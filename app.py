
import streamlit as st

st.set_page_config(
    page_title="Tinh thue TNCN va luong thuc lanh",
    page_icon="💰",
    layout="wide"
)

st.title("💰 Phần mềm tính tiền lương- Song Hào")
st.subheader("Tinh luong va thu nhap thuc lanh nam 2026")
st.info("Don vi nhap: Trieu dong / thang")

# NHAP THONG TIN
st.header("1. NHAP THONG TIN THU NHAP")

col1, col2 = st.columns(2)

with col1:
    LUONGCOBAN = st.number_input(
        "Luong co ban (trieu dong)",
        min_value=0.0,
        value=20.0,
        step=0.5
    )

    THUONG = st.number_input(
        "Tien thuong (trieu dong)",
        min_value=0.0,
        value=10.0,
        step=0.5
    )

    PHUCAPCOBHXH = st.number_input(
        "Phu cap co tinh BHXH (trieu dong)",
        min_value=0.0,
        value=2.0,
        step=0.5
    )

with col2:
    PHUCAPKHONGBHXH = st.number_input(
        "Phu cap khong tinh BHXH (trieu dong)",
        min_value=0.0,
        value=3.0,
        step=0.5
    )

    SONGUOIPHUTHUOC = st.number_input(
        "So nguoi phu thuoc",
        min_value=0,
        value=1,
        step=1
    )

# TINH TONG THU NHAP
TONGTHUNHAP = (
    LUONGCOBAN
    + THUONG
    + PHUCAPCOBHXH
    + PHUCAPKHONGBHXH
)

# TINH BHXH
TIENDONGBHXH = (
    LUONGCOBAN + PHUCAPCOBHXH
) * 0.105

# TINH THU NHAP TINH THUE
THUNHAPTINHTHUETNCN = (
    TONGTHUNHAP
    - 15.5
    - 6.2 * SONGUOIPHUTHUOC
    - TIENDONGBHXH
)

# TINH THUE TNCN
if THUNHAPTINHTHUETNCN <= 0:
    THUETNCN = 0
elif THUNHAPTINHTHUETNCN <= 10:
    THUETNCN = THUNHAPTINHTHUETNCN * 0.05
elif THUNHAPTINHTHUETNCN <= 30:
    THUETNCN = THUNHAPTINHTHUETNCN * 0.1 - 0.5
elif THUNHAPTINHTHUETNCN <= 60:
    THUETNCN = THUNHAPTINHTHUETNCN * 0.2 - 3.5
elif THUNHAPTINHTHUETNCN <= 100:
    THUETNCN = THUNHAPTINHTHUETNCN * 0.3 - 9.5
else:
    THUETNCN = THUNHAPTINHTHUETNCN * 0.35 - 14.5

# TINH LUONG THUC LANH
THUCLANH = (
    TONGTHUNHAP
    - TIENDONGBHXH
    - THUETNCN
)

# HIEN THI KET QUA
st.divider()
st.header("2. KET QUA TINH TOAN")

c1, c2, c3 = st.columns(3)

c1.metric(
    "Tong thu nhap",
    f"{TONGTHUNHAP:,.3f} trieu"
)

c2.metric(
    "Thue TNCN",
    f"{THUETNCN:,.3f} trieu"
)

c3.metric(
    "Tien dong BHXH",
    f"{TIENDONGBHXH:,.3f} trieu"
)

st.success(
    f"LUONG THUC LANH: {THUCLANH:,.3f} TRIEU DONG"
)

# BANG CHI TIET
st.header("3. BANG CHI TIET THU NHAP")

data = {
    "Khoan muc": [
        "Luong co ban",
        "Tien thuong",
        "Phu cap co BHXH",
        "Phu cap khong BHXH",
        "Tong thu nhap",
        "Giam tru ban than",
        "Giam tru nguoi phu thuoc",
        "Tien dong BHXH",
        "Thu nhap tinh thue",
        "Thue TNCN",
        "Luong thuc lanh"
    ],
    "So tien (trieu dong)": [
        LUONGCOBAN,
        THUONG,
        PHUCAPCOBHXH,
        PHUCAPKHONGBHXH,
        TONGTHUNHAP,
        15.5,
        6.2 * SONGUOIPHUTHUOC,
        TIENDONGBHXH,
        max(0, THUNHAPTINHTHUETNCN),
        THUETNCN,
        THUCLANH
    ]
}

st.dataframe(
    data,
    use_container_width=True,
    hide_index=True
)

# BIEU DO
st.header("4. CO CAU THU NHAP")

chart_data = {
    "Khoan muc": [
        "Luong thuc lanh",
        "Thue TNCN",
        "BHXH"
    ],
    "So tien": [
        THUCLANH,
        THUETNCN,
        TIENDONGBHXH
    ]
}

st.bar_chart(
    chart_data,
    x="Khoan muc",
    y="So tien"
)

st.caption(
    "Cong cu tham khao. Gia dinh tat ca phu cap "
    "deu chiu thue va can cu dong bao hiem "
    "duoc nhap dung theo quy dinh."
)
