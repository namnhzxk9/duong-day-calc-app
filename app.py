
import math
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Tính sơ bộ Đường dây 110-500kV",
    page_icon="⚡",
    layout="wide"
)

# =========================
# DATA LIBRARY
# =========================

CONDUCTOR_LIB = pd.DataFrame([
    {"Mã dây": "ACSR-185/29", "Tiết diện mm2": 185, "I cho phép A": 430, "R ohm/km": 0.157, "X ohm/km": 0.40, "Khối lượng kg/km": 732, "Ghi chú": "Tham khảo sơ bộ"},
    {"Mã dây": "ACSR-240/32", "Tiết diện mm2": 240, "I cho phép A": 520, "R ohm/km": 0.120, "X ohm/km": 0.39, "Khối lượng kg/km": 922, "Ghi chú": "Tham khảo sơ bộ"},
    {"Mã dây": "ACSR-300/39", "Tiết diện mm2": 300, "I cho phép A": 610, "R ohm/km": 0.097, "X ohm/km": 0.38, "Khối lượng kg/km": 1136, "Ghi chú": "Tham khảo sơ bộ"},
    {"Mã dây": "ACSR-400/51", "Tiết diện mm2": 400, "I cho phép A": 760, "R ohm/km": 0.073, "X ohm/km": 0.37, "Khối lượng kg/km": 1511, "Ghi chú": "Tham khảo sơ bộ"},
    {"Mã dây": "ACSR-500/64", "Tiết diện mm2": 500, "I cho phép A": 900, "R ohm/km": 0.058, "X ohm/km": 0.36, "Khối lượng kg/km": 1889, "Ghi chú": "Tham khảo sơ bộ"},
    {"Mã dây": "ACSR-630/72", "Tiết diện mm2": 630, "I cho phép A": 1050, "R ohm/km": 0.046, "X ohm/km": 0.35, "Khối lượng kg/km": 2277, "Ghi chú": "Tham khảo sơ bộ"},
])

VOLTAGE_DEFAULTS = {
    "110kV": {"span_ref": 300, "bundle": 1, "jkt": 1.1, "tower_weight": 7.5, "foundation_exc": 28, "foundation_conc": 11, "rebar": 1.2, "anchor": 0.18, "grounding": 0.10},
    "220kV": {"span_ref": 420, "bundle": 1, "jkt": 1.0, "tower_weight": 18.0, "foundation_exc": 65, "foundation_conc": 24, "rebar": 3.2, "anchor": 0.45, "grounding": 0.18},
    "500kV": {"span_ref": 520, "bundle": 4, "jkt": 0.9, "tower_weight": 55.0, "foundation_exc": 180, "foundation_conc": 70, "rebar": 9.0, "anchor": 1.2, "grounding": 0.35},
}

TOWER_TYPES = ["Đỡ thẳng", "Đỡ góc", "Néo góc", "Cột cuối", "Cột vượt", "Cột đặc biệt"]


# =========================
# FUNCTIONS
# =========================

def current_3phase(p_mw: float, u_kv: float, cos_phi: float) -> float:
    if u_kv <= 0 or cos_phi <= 0:
        return 0
    return p_mw * 1000 / (math.sqrt(3) * u_kv * cos_phi)

def select_conductor(i_a: float, s_econ: float) -> pd.DataFrame:
    df = CONDUCTOR_LIB.copy()
    df["Đạt dòng tải"] = df["I cho phép A"] >= i_a
    df["Đạt tiết diện KT"] = df["Tiết diện mm2"] >= s_econ
    df["Đạt sơ bộ"] = df["Đạt dòng tải"] & df["Đạt tiết diện KT"]
    return df

def voltage_drop_percent(i_a, u_kv, length_km, r_ohm_km, x_ohm_km, cos_phi, circuits=1, bundle=1):
    if u_kv <= 0 or circuits <= 0 or bundle <= 0:
        return 0
    sin_phi = math.sqrt(max(0, 1 - cos_phi**2))
    # Simplified: equivalent R/X reduce by number of circuits and bundle.
    r_eq = r_ohm_km / max(1, circuits * bundle)
    x_eq = x_ohm_km / max(1, circuits)
    delta_v = math.sqrt(3) * i_a * (r_eq * cos_phi + x_eq * sin_phi) * length_km
    return delta_v / (u_kv * 1000) * 100

def estimate_towers(length_km, span_m):
    if span_m <= 0:
        return 0, 0
    spans = math.ceil(length_km * 1000 / span_m)
    towers = spans + 1
    return spans, towers

def round_count(value):
    return int(max(0, round(value)))

def make_excel_bytes(sheets: dict) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe = name[:31]
            df.to_excel(writer, index=False, sheet_name=safe)
    return output.getvalue()

def status_badge(ok: bool, text_ok="PASS", text_bad="FAIL"):
    return text_ok if ok else text_bad


# =========================
# SIDEBAR
# =========================

st.sidebar.title("⚡ Input chính")
st.sidebar.caption("Bảng tính sơ bộ. Không thay thế hồ sơ tính toán chính thức.")

project_name = st.sidebar.text_input("Tên dự án", "Đường dây truyền tải sơ bộ")
voltage_level = st.sidebar.selectbox("Cấp điện áp", list(VOLTAGE_DEFAULTS.keys()), index=0)
u_kv = float(voltage_level.replace("kV", ""))
defaults = VOLTAGE_DEFAULTS[voltage_level]

p_mw = st.sidebar.number_input("Công suất truyền tải P (MW)", min_value=1.0, value=80.0, step=5.0)
length_km = st.sidebar.number_input("Chiều dài tuyến L (km)", min_value=0.1, value=10.0, step=0.5)
cos_phi = st.sidebar.slider("Hệ số công suất cosφ", 0.70, 1.00, 0.90, 0.01)
circuits = st.sidebar.selectbox("Số mạch", [1, 2, 3, 4], index=0)
bundle = st.sidebar.number_input("Số dây/bó/pha", min_value=1, max_value=6, value=int(defaults["bundle"]), step=1)
reserve_factor = st.sidebar.number_input("Hệ số dự phòng chiều dài dây", min_value=1.00, max_value=1.30, value=1.05, step=0.01)

st.sidebar.divider()
jkt = st.sidebar.number_input("Mật độ dòng điện kinh tế Jkt (A/mm2)", min_value=0.1, value=float(defaults["jkt"]), step=0.05)
span_ref = st.sidebar.number_input("Khoảng cột tham khảo (m)", min_value=50, value=int(defaults["span_ref"]), step=10)

# =========================
# CALC BASE
# =========================

i_load = current_3phase(p_mw, u_kv, cos_phi)
s_econ = i_load / jkt if jkt > 0 else 0
conductor_check = select_conductor(i_load / max(circuits * bundle, 1), s_econ / max(circuits * bundle, 1))
passed = conductor_check[conductor_check["Đạt sơ bộ"] == True]
recommended = passed.iloc[0] if not passed.empty else conductor_check.iloc[-1]

r_sel = float(recommended["R ohm/km"])
x_sel = float(recommended["X ohm/km"])
dv_percent = voltage_drop_percent(i_load, u_kv, length_km, r_sel, x_sel, cos_phi, circuits, bundle)
spans, tower_total = estimate_towers(length_km, span_ref)

conductor_length_km = length_km * circuits * 3 * bundle * reserve_factor
conductor_weight_t = conductor_length_km * float(recommended["Khối lượng kg/km"]) / 1000

opgw_length_km = length_km * reserve_factor
opgw_weight_t = opgw_length_km * 0.50  # rough placeholder t/km


# =========================
# HEADER
# =========================

st.title("⚡ Tính sơ bộ dây dẫn, cột, móng đường dây 110kV–500kV")
st.caption("Dùng cho feasibility / đấu thầu / kiểm BOQ sơ bộ. Các số thư viện là tham khảo và cần thay bằng dữ liệu thiết kế dự án.")

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Dòng tải 3 pha", f"{i_load:,.0f} A")
m2.metric("Tiết diện KT sơ bộ", f"{s_econ:,.0f} mm²")
m3.metric("Dây khuyến nghị", str(recommended["Mã dây"]))
m4.metric("Số cột sơ bộ", f"{tower_total:,}")
m5.metric("Tổng dây dẫn", f"{conductor_weight_t:,.1f} tấn")


# =========================
# TABS
# =========================

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "01. Dây dẫn",
    "02. Cột",
    "03. Móng & tiếp địa",
    "04. BOQ tổng hợp",
    "05. Check / RFI",
    "06. Hướng dẫn"
])

with tab1:
    st.subheader("01. Tính sơ bộ dây dẫn")

    c1, c2 = st.columns([1.2, 1])
    with c1:
        st.write("**Thư viện dây dẫn & kiểm sơ bộ**")
        show_df = conductor_check.copy()
        show_df["I tính trên 1 dây A"] = i_load / max(circuits * bundle, 1)
        show_df["S yêu cầu trên 1 dây mm2"] = s_econ / max(circuits * bundle, 1)
        st.dataframe(show_df, use_container_width=True, hide_index=True)

    with c2:
        st.write("**Kết quả tính nhanh**")
        st.table(pd.DataFrame([
            ["Cấp điện áp", voltage_level],
            ["Công suất truyền tải", f"{p_mw:,.1f} MW"],
            ["Chiều dài tuyến", f"{length_km:,.2f} km"],
            ["Số mạch", circuits],
            ["Số dây/bó/pha", bundle],
            ["Dòng tải tổng", f"{i_load:,.0f} A"],
            ["Dòng trên 1 dây", f"{i_load / max(circuits*bundle, 1):,.0f} A"],
            ["Dây chọn sơ bộ", recommended["Mã dây"]],
            ["Tổn thất điện áp sơ bộ", f"{dv_percent:,.2f} %"],
        ], columns=["Thông số", "Giá trị"]))

    st.info("Ghi chú: Kiểm tra điện áp rơi ở đây là công thức đơn giản. Thiết kế chính thức phải kiểm cơ lý dây, sag-tension, clearance, vầng quang, ngắn mạch và điều kiện vận hành theo hồ sơ dự án.")

with tab2:
    st.subheader("02. Nhập cơ cấu cột và tính khối lượng thép")

    st.write("Nhập tỷ lệ hoặc số lượng từng loại cột. Nếu đã có tower schedule, ưu tiên nhập số lượng thực tế.")

    default_mix = pd.DataFrame([
        {"Loại cột": "Đỡ thẳng", "Tỷ lệ %": 70, "Số lượng nhập tay": 0, "Khối lượng thép/cột trước mạ (tấn)": defaults["tower_weight"] * 0.75, "Hệ số mạ/hao hụt": 1.04},
        {"Loại cột": "Đỡ góc", "Tỷ lệ %": 8, "Số lượng nhập tay": 0, "Khối lượng thép/cột trước mạ (tấn)": defaults["tower_weight"] * 0.95, "Hệ số mạ/hao hụt": 1.04},
        {"Loại cột": "Néo góc", "Tỷ lệ %": 12, "Số lượng nhập tay": 0, "Khối lượng thép/cột trước mạ (tấn)": defaults["tower_weight"] * 1.35, "Hệ số mạ/hao hụt": 1.04},
        {"Loại cột": "Cột cuối", "Tỷ lệ %": 2, "Số lượng nhập tay": 2, "Khối lượng thép/cột trước mạ (tấn)": defaults["tower_weight"] * 1.50, "Hệ số mạ/hao hụt": 1.04},
        {"Loại cột": "Cột vượt", "Tỷ lệ %": 0, "Số lượng nhập tay": 0, "Khối lượng thép/cột trước mạ (tấn)": defaults["tower_weight"] * 2.50, "Hệ số mạ/hao hụt": 1.04},
        {"Loại cột": "Cột đặc biệt", "Tỷ lệ %": 0, "Số lượng nhập tay": 0, "Khối lượng thép/cột trước mạ (tấn)": defaults["tower_weight"] * 2.00, "Hệ số mạ/hao hụt": 1.04},
    ])

    tower_input = st.data_editor(
        default_mix,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Tỷ lệ %": st.column_config.NumberColumn(min_value=0, max_value=100, step=1),
            "Số lượng nhập tay": st.column_config.NumberColumn(min_value=0, step=1),
            "Khối lượng thép/cột trước mạ (tấn)": st.column_config.NumberColumn(min_value=0.0, step=0.1),
            "Hệ số mạ/hao hụt": st.column_config.NumberColumn(min_value=1.0, step=0.01),
        }
    )

    calc_rows = []
    remaining_manual = tower_input["Số lượng nhập tay"].sum()
    auto_base = max(0, tower_total - remaining_manual)
    for _, row in tower_input.iterrows():
        qty = int(row["Số lượng nhập tay"]) if row["Số lượng nhập tay"] > 0 else round_count(auto_base * row["Tỷ lệ %"] / 100)
        steel_before = qty * row["Khối lượng thép/cột trước mạ (tấn)"]
        steel_after = steel_before * row["Hệ số mạ/hao hụt"]
        calc_rows.append({
            "Loại cột": row["Loại cột"],
            "Số lượng": qty,
            "KL thép/cột trước mạ (tấn)": row["Khối lượng thép/cột trước mạ (tấn)"],
            "Tổng thép trước mạ (tấn)": steel_before,
            "Tổng thép sau mạ/hao hụt (tấn)": steel_after,
        })

    tower_calc = pd.DataFrame(calc_rows)
    st.write("**Kết quả cột**")
    st.dataframe(tower_calc, use_container_width=True, hide_index=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Số khoảng cột", f"{spans:,}")
    c2.metric("Số cột sơ bộ", f"{tower_calc['Số lượng'].sum():,}")
    c3.metric("Tổng thép cột sau mạ", f"{tower_calc['Tổng thép sau mạ/hao hụt (tấn)'].sum():,.1f} tấn")

    st.bar_chart(tower_calc.set_index("Loại cột")["Số lượng"])

with tab3:
    st.subheader("03. Móng, tiếp địa và khối lượng xây lắp sơ bộ")

    st.write("Nhập suất khối lượng trung bình cho 1 cột. Khi có bản vẽ móng, thay bằng số liệu thực tế theo từng vị trí.")

    foundation_df = pd.DataFrame([
        {"Hạng mục": "Đào đất đá hố móng", "Đơn vị": "m3/cột", "Suất KL": defaults["foundation_exc"]},
        {"Hạng mục": "Bê tông móng", "Đơn vị": "m3/cột", "Suất KL": defaults["foundation_conc"]},
        {"Hạng mục": "Cốt thép móng", "Đơn vị": "tấn/cột", "Suất KL": defaults["rebar"]},
        {"Hạng mục": "Bu lông neo", "Đơn vị": "tấn/cột", "Suất KL": defaults["anchor"]},
        {"Hạng mục": "Tiếp địa", "Đơn vị": "tấn/cột", "Suất KL": defaults["grounding"]},
        {"Hạng mục": "Lấp đất hố móng", "Đơn vị": "m3/cột", "Suất KL": defaults["foundation_exc"] * 0.75},
        {"Hạng mục": "San gạt/kè/rãnh", "Đơn vị": "m3/cột", "Suất KL": 0.0},
    ])

    foundation_input = st.data_editor(
        foundation_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={"Suất KL": st.column_config.NumberColumn(min_value=0.0, step=0.1)}
    )
    foundation_input["Số cột áp dụng"] = tower_calc["Số lượng"].sum()
    foundation_input["Tổng khối lượng"] = foundation_input["Suất KL"] * foundation_input["Số cột áp dụng"]

    st.dataframe(foundation_input, use_container_width=True, hide_index=True)

    c1, c2, c3, c4 = st.columns(4)
    def get_total(name):
        return foundation_input.loc[foundation_input["Hạng mục"] == name, "Tổng khối lượng"].iloc[0]
    c1.metric("Đào đất đá", f"{get_total('Đào đất đá hố móng'):,.0f} m³")
    c2.metric("Bê tông móng", f"{get_total('Bê tông móng'):,.0f} m³")
    c3.metric("Cốt thép móng", f"{get_total('Cốt thép móng'):,.1f} tấn")
    c4.metric("Tiếp địa", f"{get_total('Tiếp địa'):,.1f} tấn")

with tab4:
    st.subheader("04. BOQ tổng hợp sơ bộ")

    boq = pd.DataFrame([
        {"Nhóm": "Dây dẫn", "Hạng mục": f"Dây dẫn {recommended['Mã dây']}", "Đơn vị": "km", "Khối lượng": conductor_length_km, "Nguồn": "Tính từ chiều dài tuyến x số mạch x 3 pha x số dây/bó x dự phòng"},
        {"Nhóm": "Dây chống sét/OPGW", "Hạng mục": "OPGW / dây chống sét", "Đơn vị": "km", "Khối lượng": opgw_length_km, "Nguồn": "Tính sơ bộ theo chiều dài tuyến x dự phòng"},
        {"Nhóm": "Cột", "Hạng mục": "Tổng số cột", "Đơn vị": "cột", "Khối lượng": tower_calc["Số lượng"].sum(), "Nguồn": "Chiều dài tuyến / khoảng cột tham khảo"},
        {"Nhóm": "Cột", "Hạng mục": "Thép cột sau mạ/hao hụt", "Đơn vị": "tấn", "Khối lượng": tower_calc["Tổng thép sau mạ/hao hụt (tấn)"].sum(), "Nguồn": "Cơ cấu cột x khối lượng/cột"},
        {"Nhóm": "Móng", "Hạng mục": "Đào đất đá hố móng", "Đơn vị": "m3", "Khối lượng": get_total("Đào đất đá hố móng"), "Nguồn": "Suất KL/cột x số cột"},
        {"Nhóm": "Móng", "Hạng mục": "Bê tông móng", "Đơn vị": "m3", "Khối lượng": get_total("Bê tông móng"), "Nguồn": "Suất KL/cột x số cột"},
        {"Nhóm": "Móng", "Hạng mục": "Cốt thép móng", "Đơn vị": "tấn", "Khối lượng": get_total("Cốt thép móng"), "Nguồn": "Suất KL/cột x số cột"},
        {"Nhóm": "Móng", "Hạng mục": "Bu lông neo", "Đơn vị": "tấn", "Khối lượng": get_total("Bu lông neo"), "Nguồn": "Suất KL/cột x số cột"},
        {"Nhóm": "Tiếp địa", "Hạng mục": "Vật liệu tiếp địa", "Đơn vị": "tấn", "Khối lượng": get_total("Tiếp địa"), "Nguồn": "Suất KL/cột x số cột"},
    ])

    st.dataframe(boq, use_container_width=True, hide_index=True)

    excel_bytes = make_excel_bytes({
        "Input_summary": pd.DataFrame([
            {"Thông số": "Tên dự án", "Giá trị": project_name},
            {"Thông số": "Cấp điện áp", "Giá trị": voltage_level},
            {"Thông số": "Công suất MW", "Giá trị": p_mw},
            {"Thông số": "Chiều dài km", "Giá trị": length_km},
            {"Thông số": "cos phi", "Giá trị": cos_phi},
            {"Thông số": "Số mạch", "Giá trị": circuits},
            {"Thông số": "Số dây/bó/pha", "Giá trị": bundle},
            {"Thông số": "Khoảng cột tham khảo m", "Giá trị": span_ref},
        ]),
        "Day_dan": conductor_check,
        "Cot": tower_calc,
        "Mong_Tiep_dia": foundation_input,
        "BOQ": boq,
    })

    st.download_button(
        "⬇️ Xuất BOQ ra Excel",
        data=excel_bytes,
        file_name=f"BOQ_so_bo_{voltage_level}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with tab5:
    st.subheader("05. Kiểm tra logic và RFI")

    checks = []
    checks.append({"Nhóm": "Dây dẫn", "Điều kiện": "Dòng trên 1 dây <= I cho phép", "Trạng thái": status_badge((i_load / max(circuits*bundle,1)) <= float(recommended["I cho phép A"])), "Ghi chú": "Nếu FAIL, tăng tiết diện, tăng số mạch hoặc số dây/bó."})
    checks.append({"Nhóm": "Dây dẫn", "Điều kiện": "Tiết diện chọn >= tiết diện kinh tế", "Trạng thái": status_badge(float(recommended["Tiết diện mm2"]) >= s_econ / max(circuits*bundle,1)), "Ghi chú": "Kiểm tra lại Jkt và điều kiện vận hành."})
    checks.append({"Nhóm": "Điện áp", "Điều kiện": "Tổn thất điện áp sơ bộ <= 5%", "Trạng thái": status_badge(dv_percent <= 5.0), "Ghi chú": "Ngưỡng 5% chỉ để kiểm sơ bộ; thay bằng yêu cầu dự án nếu khác."})
    checks.append({"Nhóm": "Cột", "Điều kiện": "Có số lượng cột sơ bộ", "Trạng thái": status_badge(tower_calc["Số lượng"].sum() > 0), "Ghi chú": "Khi có tower schedule, thay số lượng sơ bộ bằng số thực tế."})
    checks.append({"Nhóm": "Móng", "Điều kiện": "Có suất bê tông/cốt thép móng", "Trạng thái": status_badge(get_total("Bê tông móng") > 0 and get_total("Cốt thép móng") > 0), "Ghi chú": "Nếu chưa có bản vẽ móng, ghi Assumption/Open-RFI."})
    checks.append({"Nhóm": "Dữ liệu", "Điều kiện": "Số liệu thư viện chỉ là tham khảo", "Trạng thái": "OPEN-RFI", "Ghi chú": "Cần thay bằng catalog/bản vẽ/tính toán chính thức trước khi chốt BOQ."})

    check_df = pd.DataFrame(checks)
    st.dataframe(check_df, use_container_width=True, hide_index=True)

    st.warning("Các dòng OPEN-RFI cần chốt bằng hồ sơ thiết kế: mặt cắt dọc, tower schedule, bảng đặc tính cột, phụ lục cơ lý dây, bản vẽ móng, đo điện trở suất đất.")

with tab6:
    st.subheader("06. Hướng dẫn sử dụng")

    guide = pd.DataFrame([
        {"Bước": 1, "Làm gì": "Nhập thông tin chính ở sidebar", "Lấy số liệu ở đâu": "Nhiệm vụ thiết kế, thuyết minh dự án, sơ đồ đấu nối"},
        {"Bước": 2, "Làm gì": "Kiểm dây dẫn ở tab 01", "Lấy số liệu ở đâu": "Phụ lục tính toán điện, catalog dây dẫn, yêu cầu vận hành"},
        {"Bước": 3, "Làm gì": "Nhập cơ cấu cột ở tab 02", "Lấy số liệu ở đâu": "Tower schedule, mặt cắt dọc, bảng thống kê cột"},
        {"Bước": 4, "Làm gì": "Nhập suất móng/tiếp địa ở tab 03", "Lấy số liệu ở đâu": "Bản vẽ móng, phụ lục tính móng, đo điện trở suất đất"},
        {"Bước": 5, "Làm gì": "Xem BOQ ở tab 04", "Lấy số liệu ở đâu": "Tự tổng hợp từ các tab trước"},
        {"Bước": 6, "Làm gì": "Xử lý cảnh báo ở tab 05", "Lấy số liệu ở đâu": "RFI với thiết kế/chủ đầu tư nếu thiếu dữ liệu"},
        {"Bước": 7, "Làm gì": "Xuất Excel", "Lấy số liệu ở đâu": "Nút download tại tab BOQ tổng hợp"},
    ])
    st.dataframe(guide, use_container_width=True, hide_index=True)

    st.markdown("""
**Nguyên tắc nhập số liệu**

- Nếu có hồ sơ thiết kế: nhập theo hồ sơ, không dùng giá trị mặc định.
- Nếu chưa có hồ sơ: dùng giá trị mặc định để estimate và ghi rõ Assumption.
- Với 220kV/500kV: không dùng bảng tính này để thay thế PLS-CADD/TOWER/PLS-POLE hoặc phụ lục tính toán chính thức.
- Cột vượt, cột đặc biệt, móng đặc biệt: luôn để riêng, không gộp với cột điển hình.
""")
