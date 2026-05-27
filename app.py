# -*- coding: utf-8 -*-
"""
SCRACE Case 2 - Streamlit Web App
Upload monthly input files, run the CKD delivery optimizer, review KPI checks, and download the final Excel plan.
"""

import os
import tempfile
import traceback
from datetime import datetime
from pathlib import Path
from io import BytesIO

import pandas as pd
import streamlit as st

import scrace_case2_auto_optimizer as opt


APP_TITLE = "The Supply Chain Race 2026 - CKD Delivery Planning Optimizer"


st.set_page_config(
    page_title="The Supply Chain Race 2026 Optimizer",
    page_icon="🚗",
    layout="wide",
)


# ---------- UI helpers ----------

def money(v):
    try:
        return f"{float(v):,.0f} VND"
    except Exception:
        return "-"


def pct(v):
    try:
        return f"{float(v) * 100:.2f}%" if float(v) <= 1.5 else f"{float(v):.2f}%"
    except Exception:
        return "-"


def num(v, digits=2):
    try:
        return f"{float(v):,.{digits}f}"
    except Exception:
        return "-"


def save_upload(uploaded_file, folder: str, fallback_name: str) -> str:
    safe_name = uploaded_file.name if uploaded_file is not None else fallback_name
    path = os.path.join(folder, safe_name)
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path


def read_sheet_safe(path: str, sheet_name: str, nrows=None):
    try:
        return pd.read_excel(path, sheet_name=sheet_name, nrows=nrows)
    except Exception:
        return None


def get_uploaded_sheet_names(uploaded_file):
    if uploaded_file is None:
        return []
    try:
        data = uploaded_file.getvalue()
        names = pd.ExcelFile(BytesIO(data)).sheet_names
        return names
    except Exception:
        return []


def guess_sheet_option(sheet_names, key, default_index=0):
    if not sheet_names:
        return default_index
    aliases = opt.SHEET_ALIASES.get(key, []) if hasattr(opt, "SHEET_ALIASES") else []
    alias_norms = [opt.normalize_sheet_key(a) for a in aliases]
    for i, sh in enumerate(sheet_names):
        nsh = opt.normalize_sheet_key(sh)
        if nsh in alias_norms:
            return i
    for i, sh in enumerate(sheet_names):
        nsh = opt.normalize_sheet_key(sh)
        if any(a and (a in nsh or nsh in a) for a in alias_norms):
            return i
    return min(default_index, len(sheet_names) - 1)


def show_dataframe_block(title: str, df: pd.DataFrame, height: int = 320):
    st.subheader(title)
    if df is None or df.empty:
        st.info("Không có dữ liệu để hiển thị.")
    else:
        st.dataframe(df, use_container_width=True, height=height)


# ---------- Sidebar config ----------

st.title(APP_TITLE)
st.caption("Tự động phân bổ và lập kế hoạch giao xe CKD theo đơn đặt hàng, kế hoạch sản xuất, bảng giá vận tải và khung giờ đại lý.")

with st.sidebar:
    st.header("1) Upload dữ liệu")
    input_file = st.file_uploader(
        "File worksheet chính (.xlsx)",
        type=["xlsx"],
        help="File gồm đơn đặt hàng, kế hoạch sản xuất, bảng giá vận tải."
    )
    window_file = st.file_uploader(
        "File khung giờ đại lý (.xlsx) - optional",
        type=["xlsx"],
        help="Có thể bỏ trống nếu không muốn ràng buộc khung giờ đại lý."
    )

    sheet_map_ui = {}
    sheet_names = get_uploaded_sheet_names(input_file)
    if input_file is not None:
        st.caption("Mapping sheet trong worksheet chính")
        if sheet_names:
            st.write("Sheets detected:", ", ".join(sheet_names))
            sheet_map_ui["orders"] = st.selectbox(
                "Sheet đơn đặt hàng",
                options=sheet_names,
                index=guess_sheet_option(sheet_names, "orders", 0),
            )
            sheet_map_ui["production"] = st.selectbox(
                "Sheet kế hoạch sản xuất",
                options=sheet_names,
                index=guess_sheet_option(sheet_names, "production", 1 if len(sheet_names) > 1 else 0),
            )
            sheet_map_ui["prices"] = st.selectbox(
                "Sheet bảng giá vận tải",
                options=sheet_names,
                index=guess_sheet_option(sheet_names, "prices", 2 if len(sheet_names) > 2 else 0),
            )
        else:
            st.error("Không đọc được danh sách sheet của file chính. Hãy kiểm tra file .xlsx.")

    st.divider()
    st.header("2) Cấu hình tối ưu")
    target_inventory_days = st.number_input(
        "Mục tiêu tồn kho TB < ngày",
        min_value=0.1,
        max_value=10.0,
        value=float(opt.CONFIG.get("target_inventory_days", 1.49)),
        step=0.01,
    )
    optimize_daily_cap = st.checkbox(
        "Tự tìm daily cap để trade-off tồn kho/Heijunka",
        value=bool(opt.CONFIG.get("optimize_daily_cap", True)),
    )
    dispatch_must_match_dealer_window = st.checkbox(
        "Bắt buộc giờ xuất bãi khớp khung giờ đại lý",
        value=bool(opt.CONFIG.get("dispatch_must_match_dealer_window", True)),
    )
    no_sunday_dispatch = st.checkbox(
        "Không xuất bãi Chủ nhật",
        value=bool(opt.CONFIG.get("no_sunday_dispatch", True)),
    )
    arrival_must_be_in_same_month = st.checkbox(
        "Xe phải tới đại lý trong cùng tháng kế hoạch",
        value=bool(opt.CONFIG.get("arrival_must_be_in_same_month", True)),
    )
    min_working_buffer_minutes = st.number_input(
        "Buffer tối thiểu sau xuất xưởng (phút làm việc)",
        min_value=0,
        max_value=1440,
        value=int(opt.CONFIG.get("min_working_buffer_minutes", 240)),
        step=15,
    )

    st.divider()
    st.header("3) Chính sách vận tải")
    st.write("- Chi phí tính theo chuyến/full slot.")
    st.write("- FIFO theo model/SKU.")
    st.write("- Không ghép đa đại lý mặc định.")
    extra_drop_fee_vnd = st.number_input(
        "Phí điểm giao thêm nếu sau này bật ghép đa đại lý",
        min_value=0,
        max_value=5_000_000,
        value=int(opt.CONFIG.get("extra_drop_fee_vnd", 300_000)),
        step=50_000,
    )


# ---------- Main tabs ----------

tab_intro, tab_run, tab_results, tab_notes = st.tabs([
    "Tổng quan",
    "Chạy tối ưu",
    "Kết quả & kiểm tra",
    "Logic hệ thống",
])

with tab_intro:
    st.markdown(
        """
        ### App này làm gì?
        App nhận dữ liệu tháng mới và tự tạo kế hoạch giao xe CKD với các nguyên tắc:

        - Giao đủ 100% đơn đặt hàng theo từng đại lý và từng model.
        - Dùng toàn bộ nguồn cung có trong kế hoạch sản xuất, bao gồm tồn từ tháng trước nếu xuất hiện trong input.
        - FIFO theo từng model/SKU vì đơn hàng đại lý là theo model.
        - Xe chỉ được xuất bãi sau tối thiểu 4 giờ làm việc thực tế kể từ giờ xuất xưởng.
        - Không xuất bãi Chủ nhật nếu bật constraint.
        - Xe phải tới đại lý trong tháng kế hoạch nếu bật constraint.
        - Chọn phương tiện theo bảng giá: có giá thì có tuyến.
        - Chi phí tính theo chuyến/full capacity, bao gồm cả slot rỗng.
        - Không tự ghép đa đại lý nếu chưa có Google Maps/distance matrix.
        """
    )

    st.warning(
        "Nếu muốn ghép nhiều đại lý trên cùng xe, cần bổ sung distance matrix/Google Maps API để chứng minh cùng tuyến và tính phụ phí điểm giao thêm."
    )

with tab_run:
    st.subheader("Chạy optimizer")
    if input_file is None:
        st.info("Upload file worksheet chính ở sidebar trước.")
    else:
        st.success(f"Đã nhận file chính: {input_file.name}")
        if window_file is not None:
            st.success(f"Đã nhận file khung giờ: {window_file.name}")
        else:
            st.info("Không upload file khung giờ: app sẽ chạy không ràng buộc khung giờ đại lý.")

        run_button = st.button("🚀 Chạy tối ưu và xuất file final", type="primary", use_container_width=True)

        if run_button:
            with tempfile.TemporaryDirectory() as tmpdir:
                try:
                    main_path = save_upload(input_file, tmpdir, "input.xlsx")
                    window_path = save_upload(window_file, tmpdir, "windows.xlsx") if window_file is not None else None
                    output_path = os.path.join(tmpdir, f"SCRACE_case2_STREAMLIT_FINAL_{datetime.now():%Y%m%d_%H%M%S}.xlsx")

                    config = dict(opt.CONFIG)
                    config.update({
                        "target_inventory_days": float(target_inventory_days),
                        "optimize_daily_cap": bool(optimize_daily_cap),
                        "dispatch_must_match_dealer_window": bool(dispatch_must_match_dealer_window and window_file is not None),
                        "no_sunday_dispatch": bool(no_sunday_dispatch),
                        "arrival_must_be_in_same_month": bool(arrival_must_be_in_same_month),
                        "min_working_buffer_minutes": int(min_working_buffer_minutes),
                        "allow_multi_dealer_combine": False,
                        "extra_drop_fee_vnd": int(extra_drop_fee_vnd),
                    })

                    with st.spinner("Đang đọc dữ liệu, tối ưu chuyến, kiểm tra FIFO/4h/cost và export Excel..."):
                        out_file, metrics = opt.run_optimizer(
                            input_file=main_path,
                            window_file=window_path,
                            output_file=output_path,
                            config=config,
                            sheet_map=sheet_map_ui,
                        )

                    with open(out_file, "rb") as f:
                        file_bytes = f.read()

                    st.session_state["output_bytes"] = file_bytes
                    st.session_state["metrics"] = metrics
                    st.session_state["output_name"] = Path(out_file).name

                    # Keep a copy in a temp-independent path for preview during this session.
                    preview_path = os.path.join(tempfile.gettempdir(), Path(out_file).name)
                    with open(preview_path, "wb") as f:
                        f.write(file_bytes)
                    st.session_state["preview_path"] = preview_path

                    st.success("Đã chạy xong và tạo file final.")

                except Exception as exc:
                    st.error("Có lỗi khi chạy optimizer. Kiểm tra lại mapping sheet/file input ở sidebar.")
                    st.warning(str(exc))
                    with st.expander("Chi tiết kỹ thuật"):
                        st.code("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))

    if "output_bytes" in st.session_state:
        st.download_button(
            "⬇️ Tải Excel final",
            data=st.session_state["output_bytes"],
            file_name=st.session_state.get("output_name", "SCRACE_case2_STREAMLIT_FINAL.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

with tab_results:
    st.subheader("Dashboard KPI")
    metrics = st.session_state.get("metrics")
    if not metrics:
        st.info("Chưa có kết quả. Qua tab 'Chạy tối ưu' để chạy trước.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Xe giao", f"{int(metrics.get('total_delivered', 0)):,}")
        c2.metric("Số chuyến", f"{int(metrics.get('trip_count', 0)):,}")
        c3.metric("Tổng chi phí", money(metrics.get("total_trip_cost")))
        c4.metric("Tồn kho TB", num(metrics.get("avg_inventory_days"), 4))

        c5, c6, c7, c8 = st.columns(4)
        c5.metric("Điểm tồn kho", f"{int(metrics.get('inventory_score', 0))}/25")
        c6.metric("Điểm FIFO", f"{int(metrics.get('fifo_score_if_model_level', 0))}/5")
        c7.metric("Điểm Heijunka", f"{int(metrics.get('heijunka_score', 0))}/15")
        c8.metric("Slot rỗng tính phí", f"{int(metrics.get('empty_slots_paid', 0)):,}")

        c9, c10, c11, c12 = st.columns(4)
        c9.metric("Utilization slot", pct(metrics.get("slot_utilization")))
        c10.metric("Xe giao sau tháng", f"{int(metrics.get('arrival_after_month_count', 0)):,}")
        c11.metric("Xe xuất CN", f"{int(metrics.get('sunday_dispatch_vehicle_count', 0)):,}")
        c12.metric("Điểm vận hành", f"{int(metrics.get('inventory_score', 0)) + int(metrics.get('fifo_score_if_model_level', 0)) + int(metrics.get('heijunka_score', 0))}/45")

        preview_path = st.session_state.get("preview_path")
        if preview_path and os.path.exists(preview_path):
            st.divider()
            subtab1, subtab2, subtab3, subtab4 = st.tabs(["Trip Summary", "Dealer Coverage", "FIFO Check", "Heijunka"])
            with subtab1:
                show_dataframe_block("Trip_Summary", read_sheet_safe(preview_path, "Trip_Summary"))
            with subtab2:
                show_dataframe_block("Dealer_Coverage", read_sheet_safe(preview_path, "Dealer_Coverage"))
            with subtab3:
                show_dataframe_block("FIFO_Check", read_sheet_safe(preview_path, "FIFO_Check"))
            with subtab4:
                show_dataframe_block("Heijunka_Check", read_sheet_safe(preview_path, "Heijunka_Check"))

with tab_notes:
    st.markdown(
        """
        ### Logic tính chính

        **1. Phân bổ nguồn cung**  
        Hệ thống coi toàn bộ xe có trong kế hoạch sản xuất/input là nguồn cung khả dụng. Nếu trong input có xe tháng trước, xe đó được xem như tồn đầu kỳ và được dùng cho đơn tháng mới theo FIFO.

        **2. FIFO theo model/SKU**  
        Với từng model, hệ thống sắp xếp xe theo thời gian xuất xưởng và chọn các xe xuất xưởng sớm nhất để giao. Cách này phù hợp vì đơn đại lý đặt theo model cụ thể.

        **3. Buffer 4 giờ làm việc**  
        Xe chỉ được xuất bãi sau khi đã đủ số phút làm việc cấu hình, mặc định 240 phút. Thời gian ngoài ca làm việc không được tính vào buffer.

        **4. Chọn phương tiện**  
        Với từng đại lý, hệ thống chỉ xét các phương tiện có trong bảng giá. Tổ hợp chuyến được chọn theo chi phí thấp nhất, có xét capacity và slot rỗng.

        **5. Chi phí theo chuyến/full slot**  
        Nếu dùng xe lồng 7 chở 5 xe thì vẫn trả đủ 7 slot. Chi phí trip được phân bổ ngược về từng xe trong chuyến để tổng chi phí trong kế hoạch bằng tổng chi phí trip.

        **6. Ghép chuyến**  
        Bản app mặc định không ghép nhiều đại lý trên cùng xe. Lý do: cần distance matrix/Google Maps API để chứng minh cùng tuyến, tính thứ tự stop, leadtime, và phụ phí điểm giao thêm. Module này có thể mở rộng sau.
        """
    )
