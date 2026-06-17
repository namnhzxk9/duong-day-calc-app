
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
    {"Mã dây": "ACSR-330", "Tiết diện mm2": 330, "I cho phép A": 680, "R ohm/km": 0.088, "X ohm/km": 0.38, "Khối lượng kg/km": 1260, "Ghi chú": "Tham khảo sơ bộ / cần thay bằng catalog"},
    {"Mã dây": "ACSR-400", "Tiết diện mm2": 400, "I cho phép A": 760, "R ohm/km": 0.073, "X ohm/km": 0.37, "Khối lượng kg/km": 1511, "Ghi chú": "Tham khảo sơ bộ"},
    {"Mã dây": "ACSR-500/64", "Tiết diện mm2": 500, "I cho phép A": 900, "R ohm/km": 0.058, "X ohm/km": 0.36, "Khối lượng kg/km": 1889, "Ghi chú": "Tham khảo sơ bộ"},
    {"Mã dây": "ACSR-630/72", "Tiết diện mm2": 630, "I cho phép A": 1050, "R ohm/km": 0.046, "X ohm/km": 0.35, "Khối lượng kg/km": 2277, "Ghi chú": "Tham khảo sơ bộ"},
])

VOLTAGE_DEFAULTS = {
    "110kV": {"span_ref": 300, "bundle": 1, "jkt": 1.1, "tower_weight": 7.5, "foundation_exc": 28, "foundation_conc": 11, "rebar": 1.2, "anchor": 0.18, "grounding": 0.10},
    "220kV": {"span_ref": 420, "bundle": 2, "jkt": 1.0, "tower_weight": 18.0, "foundation_exc": 65, "foundation_conc": 24, "rebar": 3.2, "anchor": 0.45, "grounding": 0.18},
    "500kV": {"span_ref": 520, "bundle": 4, "jkt": 0.9, "tower_weight": 55.0, "foundation_exc": 180, "foundation_conc": 70, "rebar": 9.0, "anchor": 1.2, "grounding": 0.35},
}

TRANSIT_HELP = """
**Đấu nối transit trên đường dây hiện trạng**

Với phương án đấu nối transit/cắt xen trên đường dây hiện hữu, dây dẫn của đoạn đấu nối nên theo cấu hình đường dây hiện trạng
để đồng bộ cơ điện, phụ kiện, vận hành và giảm rủi ro tại điểm đấu nối. App sẽ ưu tiên dùng `Dây hiện trạng` làm dây BOQ,
đồng thời vẫn kiểm dòng tải và cảnh báo nếu không đạt.
"""

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

def conductor_from_lib(code):
    df = CONDUCTOR_LIB[CONDUCTOR_LIB["Mã dây"] == code]
    return df.iloc[0] if not df.empty else None

# =========================
# SIDEBAR
# =========================

st.sidebar.title("⚡ Input chính")
st.sidebar.caption("Bảng tính sơ bộ. Không thay thế hồ sơ tính toán chính thức.")

project_name = st.sidebar.text_input("Tên dự án", "Đường dây truyền tải sơ bộ")
voltage_level = st.sidebar.selectbox("Cấp điện áp", list(VOLTAGE_DEFAULTS.keys()), index=1)
u_kv = float(voltage_level.replace("kV", ""))
defaults = VOLTAGE_DEFAULTS[voltage_level]

connection_type = st.sidebar.selectbox(
    "Kiểu đấu nối",
    ["Đấu nối mới", "Transit / cắt xen trên ĐD hiện trạng"],
    index=1
)

p_mw = st.sidebar.number_input("Công suất truyền tải P (MW)", min_value=1.0, value=90.0, step=5.0)
length_km = st.sidebar.number_input("Chiều dài đoạn đấu nối / tuyến mới L (km)", min_value=0.1, value=5.0, step=0.5)
cos_phi = st.sidebar.slider("Hệ số công suất cosφ", 0.70, 1.00, 0.90, 0.01)
circuits = st.sidebar.selectbox("Số mạch", [1, 2, 3, 4], index=1)
bundle = st.sidebar.number_input("Số dây/bó/pha của đoạn thiết kế", min_value=1, max_value=6, value=int(defaults["bundle"]), step=1)
reserve_factor = st.sidebar.number_input("Hệ số dự phòng chiều dài dây", min_value=1.00, max_value=1.30, value=1.05, step=0.01)

st.sidebar.divider()
jkt = st.sidebar.number_input("Mật độ dòng điện kinh tế Jkt (A/mm2)", min_value=0.1, value=float(defaults["jkt"]), step=0.05)
span_ref = st.sidebar.number_input("Khoảng cột tham khảo (m)", min_value=50, value=int(defaults["span_ref"]), step=10)

st.sidebar.divider()
st.sidebar.subheader("Đường dây hiện trạng")
existing_voltage = st.sidebar.selectbox("Cấp điện áp hiện trạng", ["110kV", "220kV", "500kV"], index=list(VOLTAGE_DEFAULTS.keys()).index(voltage_level))
existing_conductor_code = st.sidebar.selectbox("Dây hiện trạng", CONDUCTOR_LIB["Mã dây"].tolist(), index=4)
existing_bundle = st.sidebar.number_input("Số dây/bó/pha hiện trạng", min_value=1, max_value=6, value=2 if voltage_level == "220kV" else int(defaults["bundle"]), step=1)
existing_circuits = st.sidebar.selectbox("Số mạch hiện trạng", [1, 2, 3, 4], index=1)

existing_conductor = conductor_from_lib(existing_conductor_code)

# If transit, force BOQ conductor to existing conductor and configuration.
if connection_type.startswith("Transit"):
    boq_conductor = existing_conductor
    boq_bundle = existing_bundle
    boq_circuits = existing_circuits
else:
    boq_conductor = None
    boq_bundle = bundle
    boq_circuits = circuits

# =========================
# CALC BASE
# =========================

i_load = current_3phase(p_mw, u_kv, cos_phi)
i_per_wire_design = i_load / max(circuits * bundle, 1)
s_econ = i_load / jkt if jkt > 0 else 0
s_per_wire_design = s_econ / max(circuits * bundle, 1)

conductor_check = select_conductor(i_per_wire_design, s_per_wire_design)
passed = conductor_check[conductor_check["Đạt sơ bộ"] == True]
recommended = passed.iloc[0] if not passed.empty else conductor_check.iloc[-1]

if boq_conductor is None:
    boq_conductor = recommended

i_per_wire_boq = i_load / max(boq_circuits * boq_bundle, 1)
s_per_wire_boq = s_econ / max(boq_circuits * boq_bundle, 1)
r_sel = float(boq_conductor["R ohm/km"])
x_sel = float(boq_conductor["X ohm/km"])
dv_percent = voltage_drop_percent(i_load, u_kv, length_km, r_sel, x_sel, cos_phi, boq_circuits, boq_bundle)

spans, tower_total = estimate_towers(length_km, span_ref)

conductor_length_km = length_km * boq_circuits * 3 * boq_bundle * reserve_factor
conductor_weight_t = conductor_length_km * float(boq_conductor["Khối lượng kg/km"]) / 1000

opgw_length_km = length_km * reserve_factor
opgw_weight_t = opgw_length_km * 0.50

existing_capacity_ok = i_per_wire_boq <= float(boq_conductor["I cho phép A"])
existing_section_ok = float(boq_conductor["Tiết diện mm2"]) >= s_per_wire_boq
config_match = (boq_bundle == bundle and boq_circuits == circuits)

# =========================
# HEADER
# =========================

st.title("⚡ Tính sơ bộ dây dẫn, cột, móng đường dây 110kV–500kV")
st.caption("Có logic riêng cho đấu nối transit/cắt xen trên đường dây hiện trạng.")

if connection_type.startswith("Transit"):
    st.info(TRANSIT_HELP)

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Dòng tải 3 pha", f"{i_load:,.0f} A")
m2.metric("Dòng trên 1 dây", f"{i_per_wire_boq:,.0f} A")
m3.metric("Dây dùng cho BOQ", str(boq_conductor["Mã dây"]))
m4.metric("Số cột sơ bộ", f"{tower_total:,}")
m5.metric("Tổng dây dẫn", f"{conductor_weight_t:,.1f} tấn")

if connection_type.startswith("Transit"):
    if not existing_capacity_ok:
        st.error("Dây hiện trạng đang chọn không đạt kiểm tra dòng tải sơ bộ. Cần kiểm tra lại công suất truyền tải, số mạch, số dây/bó hoặc hồ sơ hiện trạng.")
    elif not existing_section_ok:
        st.warning("Dây hiện trạng đạt dòng tải nhưng nhỏ hơn tiết diện kinh tế sơ bộ. Cần ghi chú là cấu hình theo đường dây hiện trạng và kiểm tra lại theo tiêu chí dự án.")
    else:
        st.success("Cấu hình dây hiện trạng đạt kiểm tra sơ bộ theo dữ liệu đang nhập.")

# =========================
# TABS
# =========================

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "00. Transit",
    "01. Dây dẫn",
    "02. Cột",
    "03. Móng & tiếp địa",
    "04. BOQ tổng hợp",
    "05. Check / RFI",
    "06. Hướng dẫn"
])

with tab1:
    st.subheader("00. Logic đấu nối transit / cắt xen")

    transit_df = pd.DataFrame([
        {"Thông số": "Kiểu đấu nối", "Giá trị": connection_type, "Ghi chú": "Nếu transit, ưu tiên đồng bộ với đường dây hiện trạng"},
        {"Thông số": "Cấp điện áp hiện trạng", "Giá trị": existing_voltage, "Ghi chú": "Phải khớp với điểm đấu nối"},
        {"Thông số": "Dây hiện trạng", "Giá trị": existing_conductor_code, "Ghi chú": "Lấy từ hồ sơ hoàn công / khảo sát / EVN cung cấp"},
        {"Thông số": "Số mạch hiện trạng", "Giá trị": existing_circuits, "Ghi chú": "Áp dụng cho đoạn đấu nối nếu cắt xen"},
        {"Thông số": "Số dây/bó/pha hiện trạng", "Giá trị": existing_bundle, "Ghi chú": "Ví dụ 220kV phân pha ACSR 2x330 hoặc dây đơn ACSR 400"},
        {"Thông số": "Dây dùng trong BOQ", "Giá trị": boq_conductor["Mã dây"], "Ghi chú": "Transit: lấy theo hiện trạng; Đấu nối mới: lấy theo dây khuyến nghị"},
    ])
    st.dataframe(transit_df, use_container_width=True, hide_index=True)

    st.write("**Kiểm tra đồng bộ dây hiện trạng**")
    check_existing = pd.DataFrame([
        {"Điều kiện": "Dòng trên 1 dây <= I cho phép của dây hiện trạng", "Giá trị tính": f"{i_per_wire_boq:,.0f} A <= {boq_conductor['I cho phép A']:,.0f} A", "Trạng thái": status_badge(existing_capacity_ok), "Ghi chú": "Nếu FAIL thì không đủ cơ sở dùng dây hiện trạng theo tải nhập."},
        {"Điều kiện": "Tiết diện dây hiện trạng >= tiết diện kinh tế sơ bộ", "Giá trị tính": f"{boq_conductor['Tiết diện mm2']:,.0f} mm² >= {s_per_wire_boq:,.0f} mm²", "Trạng thái": status_badge(existing_section_ok), "Ghi chú": "Nếu FAIL chưa chắc sai, nhưng phải có cơ sở EVN/thiết kế chấp thuận."},
        {"Điều kiện": "Cấu hình đoạn thiết kế đồng bộ hiện trạng", "Giá trị tính": f"Thiết kế: {circuits} mạch, {bundle} dây/bó | Hiện trạng: {existing_circuits} mạch, {existing_bundle} dây/bó", "Trạng thái": status_badge(config_match), "Ghi chú": "Nếu khác, cần RFI làm rõ phạm vi chuyển tiếp/đấu nối."},
    ])
    st.dataframe(check_existing, use_container_width=True, hide_index=True)

    st.warning("Không tự tăng/giảm size dây tại vị trí transit nếu chưa có chấp thuận. Nếu đường dây hiện trạng là ACSR 2x330 thì BOQ đoạn đấu nối nên dùng ACSR 2x330, kèm kiểm tra khả năng tải và phụ kiện tương ứng.")

with tab1:
    st.write("**Ví dụ từ nhóm dự án solar 220kV**")
    example_df = pd.DataFrame([
        {"Dự án": "Bắc Ái 7 - Hồ Sông Cái", "Dây đấu nối theo bảng": "Dây đơn ACSR 400", "Chiều dài": "~5 km", "Ghi chú": "Transit trên ĐD hiện trạng: nhập dây hiện trạng ACSR 400, số dây/bó = 1 nếu đúng hồ sơ hiện trạng"},
        {"Dự án": "Hồ Sông Sắt", "Dây đấu nối theo bảng": "Dây đơn ACSR 400", "Chiều dài": "~2 km", "Ghi chú": "Tương tự, không lấy tự động theo công suất nếu hiện trạng đã chốt ACSR 400"},
        {"Dự án": "Ninh Phước 6.3 / Phước Ninh MR GĐ2", "Dây đấu nối theo bảng": "Phân pha ACSR 2x330", "Chiều dài": "~1 km / 8.82 km", "Ghi chú": "Nhập dây hiện trạng ACSR-330, số dây/bó = 2"},
    ])
    st.dataframe(example_df, use_container_width=True, hide_index=True)

with tab2:
    st.subheader("01. Tính sơ bộ dây dẫn")

    c1, c2 = st.columns([1.2, 1])
    with c1:
        st.write("**Thư viện dây dẫn & kiểm sơ bộ theo cấu hình thiết kế nhập ở sidebar**")
        show_df = conductor_check.copy()
        show_df["I tính trên 1 dây A"] = i_per_wire_design
        show_df["S yêu cầu trên 1 dây mm2"] = s_per_wire_design
        st.dataframe(show_df, use_container_width=True, hide_index=True)

    with c2:
        st.write("**Kết quả tính nhanh**")
        st.table(pd.DataFrame([
            ["Cấp điện áp", voltage_level],
            ["Công suất truyền tải", f"{p_mw:,.1f} MW"],
            ["Chiều dài đoạn đấu nối/tuyến mới", f"{length_km:,.2f} km"],
            ["Kiểu đấu nối", connection_type],
            ["Số mạch dùng BOQ", boq_circuits],
            ["Số dây/bó/pha dùng BOQ", boq_bundle],
            ["Dòng tải tổng", f"{i_load:,.0f} A"],
            ["Dòng trên 1 dây dùng BOQ", f"{i_per_wire_boq:,.0f} A"],
            ["Dây chọn theo tính sơ bộ", recommended["Mã dây"]],
            ["Dây dùng trong BOQ", boq_conductor["Mã dây"]],
            ["Tổn thất điện áp sơ bộ", f"{dv_percent:,.2f} %"],
        ], columns=["Thông số", "Giá trị"]))

    st.info("Nếu là transit, bảng chọn dây chỉ dùng để kiểm. BOQ sẽ ưu tiên dây hiện trạng đã nhập.")

with tab3:
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

with tab4:
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

with tab5:
    st.subheader("04. BOQ tổng hợp sơ bộ")

    boq = pd.DataFrame([
        {"Nhóm": "Dây dẫn", "Hạng mục": f"Dây dẫn {boq_conductor['Mã dây']}", "Đơn vị": "km", "Khối lượng": conductor_length_km, "Nguồn": "Transit: theo dây hiện trạng; Đấu nối mới: theo dây khuyến nghị"},
        {"Nhóm": "Dây dẫn", "Hạng mục": f"Khối lượng dây dẫn {boq_conductor['Mã dây']}", "Đơn vị": "tấn", "Khối lượng": conductor_weight_t, "Nguồn": "Chiều dài dây x kg/km"},
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
            {"Thông số": "Kiểu đấu nối", "Giá trị": connection_type},
            {"Thông số": "Dây hiện trạng", "Giá trị": existing_conductor_code},
            {"Thông số": "Số mạch hiện trạng", "Giá trị": existing_circuits},
            {"Thông số": "Số dây/bó/pha hiện trạng", "Giá trị": existing_bundle},
            {"Thông số": "Dây dùng BOQ", "Giá trị": boq_conductor["Mã dây"]},
            {"Thông số": "Công suất MW", "Giá trị": p_mw},
            {"Thông số": "Chiều dài km", "Giá trị": length_km},
            {"Thông số": "cos phi", "Giá trị": cos_phi},
            {"Thông số": "Khoảng cột tham khảo m", "Giá trị": span_ref},
        ]),
        "Transit_check": check_existing if "check_existing" in locals() else pd.DataFrame(),
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

with tab6:
    st.subheader("05. Kiểm tra logic và RFI")

    checks = []
    checks.append({"Nhóm": "Dây dẫn", "Điều kiện": "Dòng trên 1 dây <= I cho phép của dây dùng BOQ", "Trạng thái": status_badge(existing_capacity_ok), "Ghi chú": "Nếu FAIL, cần kiểm lại hiện trạng hoặc phương án đấu nối."})
    checks.append({"Nhóm": "Dây dẫn", "Điều kiện": "Tiết diện dây dùng BOQ >= tiết diện kinh tế sơ bộ", "Trạng thái": status_badge(existing_section_ok), "Ghi chú": "Nếu FAIL, cần ghi chú/được EVN chấp thuận khi vẫn dùng dây hiện trạng."})
    checks.append({"Nhóm": "Transit", "Điều kiện": "Dây dùng BOQ đồng bộ đường dây hiện trạng", "Trạng thái": "PASS" if connection_type.startswith("Transit") else "N/A", "Ghi chú": "Transit phải lấy theo hồ sơ hiện trạng, không lấy theo dây tự chọn nếu chưa có chấp thuận."})
    checks.append({"Nhóm": "Điện áp", "Điều kiện": "Tổn thất điện áp sơ bộ <= 5%", "Trạng thái": status_badge(dv_percent <= 5.0), "Ghi chú": "Ngưỡng 5% chỉ để kiểm sơ bộ; thay bằng yêu cầu dự án nếu khác."})
    checks.append({"Nhóm": "Cột", "Điều kiện": "Có số lượng cột sơ bộ", "Trạng thái": status_badge(tower_calc["Số lượng"].sum() > 0), "Ghi chú": "Khi có tower schedule, thay số lượng sơ bộ bằng số thực tế."})
    checks.append({"Nhóm": "Dữ liệu", "Điều kiện": "Có hồ sơ hiện trạng đường dây", "Trạng thái": "OPEN-RFI" if connection_type.startswith("Transit") else "N/A", "Ghi chú": "Cần xác nhận dây dẫn hiện trạng, phụ kiện, khoảng cột, loại cột, sức tải, hành lang, phương án cắt điện."})

    check_df = pd.DataFrame(checks)
    st.dataframe(check_df, use_container_width=True, hide_index=True)

    st.warning("RFI cần chốt cho transit: dây hiện trạng, số mạch, số dây/bó, phụ kiện đấu nối, loại chuỗi néo/đỡ, điểm cắt xen, phương án cắt điện, khả năng truyền tải sau đấu nối.")

with tab7:
    st.subheader("06. Hướng dẫn sử dụng")

    guide = pd.DataFrame([
        {"Bước": 1, "Làm gì": "Chọn kiểu đấu nối", "Lấy số liệu ở đâu": "Phương án đấu nối / văn bản thỏa thuận đấu nối"},
        {"Bước": 2, "Làm gì": "Nếu transit: nhập dây hiện trạng, số mạch, số dây/bó", "Lấy số liệu ở đâu": "Hồ sơ hoàn công, khảo sát tuyến, EVN cung cấp"},
        {"Bước": 3, "Làm gì": "Nhập công suất, chiều dài đoạn đấu nối", "Lấy số liệu ở đâu": "Báo cáo NCKT/TKCS, sơ đồ đấu nối, bảng dự án"},
        {"Bước": 4, "Làm gì": "Xem tab Transit để kiểm dây hiện trạng", "Lấy số liệu ở đâu": "App tự tính từ input"},
        {"Bước": 5, "Làm gì": "Nhập cơ cấu cột", "Lấy số liệu ở đâu": "Mặt cắt dọc, tower schedule, bản vẽ sơ đồ cột"},
        {"Bước": 6, "Làm gì": "Nhập suất móng/tiếp địa", "Lấy số liệu ở đâu": "Bản vẽ móng, phụ lục tính móng, đo điện trở suất đất"},
        {"Bước": 7, "Làm gì": "Xuất BOQ", "Lấy số liệu ở đâu": "Tab BOQ tổng hợp"},
    ])
    st.dataframe(guide, use_container_width=True, hide_index=True)

    st.markdown("""
**Nguyên tắc riêng cho transit**

- Dây dẫn, số dây/bó và phụ kiện đấu nối phải ưu tiên theo đường dây hiện trạng.
- App vẫn có phần chọn dây sơ bộ, nhưng với transit chỉ dùng để kiểm, không dùng để tự thay dây.
- Nếu dây hiện trạng không đạt kiểm tra tải sơ bộ, không tự nâng size trong BOQ; phải mở RFI/technical query.
- Đoạn vào TBA hoặc đoạn chuyển tiếp nếu có cấu hình khác phải tách thành hạng mục riêng.
""")
