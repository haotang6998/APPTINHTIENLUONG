
import streamlit as st
import pandas as pd
import os
from io import BytesIO
from datetime import datetime

# CAU HINH TRANG
st.set_page_config(
    page_title="Phan mem tinh luong - Song Hao",
    page_icon="💰",
    layout="wide"
)

# HIEN THI HINH ANH
if os.path.exists("IMG_2692.jpeg"):
    st.image("IMG_2692.jpeg", use_container_width=True)

st.title("💰 PHAN MEM TINH TIEN LUONG - SONG HAO")
st.subheader("Tinh thue TNCN va luong thuc lanh nam 2026")
st.info("Don vi tinh: Trieu dong / thang")

# CAU HINH TINH TOAN
GIAMTRUBANTHAN = 15.5
GIAMTRUPHUTHUOC = 6.2
TYLEBAOHIEM = 0.105

# CAC COT NHAP LIEU
COTNHAP = [
    "Ma NV",
    "Ho va ten",
    "Luong co ban",
    "Tien thuong",
    "Phu cap co BHXH",
    "Phu cap khong BHXH",
    "So nguoi phu thuoc"
]

COTSO = [
    "Luong co ban",
    "Tien thuong",
    "Phu cap co BHXH",
    "Phu cap khong BHXH",
    "So nguoi phu thuoc"
]

# TAO DU LIEU NHAN VIEN
def tao_nhan_vien(i):
    return {
        "Ma NV": f"NV{i:03d}",
        "Ho va ten": "",
        "Luong co ban": 0.0,
        "Tien thuong": 0.0,
        "Phu cap co BHXH": 0.0,
        "Phu cap khong BHXH": 0.0,
        "So nguoi phu thuoc": 0
    }

# KHOI TAO DU LIEU
if "bang_luong" not in st.session_state:
    st.session_state.bang_luong = pd.DataFrame(
        [tao_nhan_vien(i) for i in range(1, 6)]
    )

# DONG BO THAY DOI TU BANG NHAP
def luu_chinh_sua():
    thay_doi = st.session_state.get("bang_nhap", {})
    df = st.session_state.bang_luong.copy()

    for vi_tri, cac_cot in thay_doi.get(
        "edited_rows", {}
    ).items():
        vi_tri = int(vi_tri)

        if vi_tri < len(df):
            for ten_cot, gia_tri in cac_cot.items():
                df.at[vi_tri, ten_cot] = gia_tri

    st.session_state.bang_luong = df

# HAM TINH THUE TNCN
def tinh_thue(TNTT):
    if TNTT <= 0:
        return 0
    elif TNTT <= 10:
        return TNTT * 0.05
    elif TNTT <= 30:
        return TNTT * 0.1 - 0.5
    elif TNTT <= 60:
        return TNTT * 0.2 - 3.5
    elif TNTT <= 100:
        return TNTT * 0.3 - 9.5
    else:
        return TNTT * 0.35 - 14.5

# HAM TINH LUONG CHO NHIEU NHAN VIEN
def tinh_bang_luong(df):
    kq = df.copy()

    for cot in COTSO:
        kq[cot] = pd.to_numeric(
            kq[cot], errors="coerce"
        ).fillna(0).clip(lower=0)

    kq["So nguoi phu thuoc"] = (
        kq["So nguoi phu thuoc"].astype(int)
    )

    # TONG THU NHAP
    kq["Tong thu nhap"] = (
        kq["Luong co ban"]
        + kq["Tien thuong"]
        + kq["Phu cap co BHXH"]
        + kq["Phu cap khong BHXH"]
    )

    # TIEN DONG BAO HIEM
    kq["Tien dong BHXH"] = (
        kq["Luong co ban"]
        + kq["Phu cap co BHXH"]
    ) * TYLEBAOHIEM

    # GIAM TRU GIA CANH
    kq["Giam tru ban than"] = GIAMTRUBANTHAN

    kq["Giam tru phu thuoc"] = (
        kq["So nguoi phu thuoc"]
        * GIAMTRUPHUTHUOC
    )

    # THU NHAP TINH THUE
    kq["Thu nhap tinh thue"] = (
        kq["Tong thu nhap"]
        - kq["Tien dong BHXH"]
        - kq["Giam tru ban than"]
        - kq["Giam tru phu thuoc"]
    ).clip(lower=0)

    # THUE TNCN
    kq["Thue TNCN"] = (
        kq["Thu nhap tinh thue"]
        .apply(tinh_thue)
    )

    # LUONG THUC LANH
    kq["Luong thuc lanh"] = (
        kq["Tong thu nhap"]
        - kq["Tien dong BHXH"]
        - kq["Thue TNCN"]
    )

    return kq

# HAM XUAT FILE EXCEL
def xuat_excel(df):
    output = BytesIO()

    tong_hop = pd.DataFrame({
        "Chi tieu": [
            "So luong nhan vien",
            "Tong thu nhap",
            "Tong tien dong BHXH",
            "Tong thu nhap tinh thue",
            "Tong thue TNCN",
            "Tong luong thuc lanh"
        ],
        "Gia tri": [
            len(df),
            df["Tong thu nhap"].sum(),
            df["Tien dong BHXH"].sum(),
            df["Thu nhap tinh thue"].sum(),
            df["Thue TNCN"].sum(),
            df["Luong thuc lanh"].sum()
        ]
    })

    with pd.ExcelWriter(
        output, engine="xlsxwriter"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="Bang luong chi tiet",
            index=False,
            startrow=2
        )

        tong_hop.to_excel(
            writer,
            sheet_name="Tong hop",
            index=False,
            startrow=2
        )

        workbook = writer.book

        tieude = workbook.add_format({
            "bold": True,
            "font_size": 16,
            "font_color": "white",
            "bg_color": "#17365D",
            "align": "center",
            "valign": "vcenter"
        })

        header = workbook.add_format({
            "bold": True,
            "bg_color": "#DCEAF7",
            "border": 1,
            "align": "center",
            "text_wrap": True
        })

        tien = workbook.add_format({
            "num_format": "#,##0.000",
            "border": 1
        })

        dong = workbook.add_format({
            "num_format": "#,##0",
            "border": 1
        })

        for ten_sheet, data_sheet in [
            ("Bang luong chi tiet", df),
            ("Tong hop", tong_hop)
        ]:
            ws = writer.sheets[ten_sheet]
            so_cot = len(data_sheet.columns)

            ws.merge_range(
                0, 0, 0, so_cot - 1,
                "PHAN MEM TINH LUONG - SONG HAO",
                tieude
            )

            ws.set_row(0, 30)
            ws.set_row(2, 36)
            ws.freeze_panes(3, 2)

            ws.autofilter(
                2, 0,
                len(data_sheet) + 2,
                so_cot - 1
            )

            for j, ten_cot in enumerate(
                data_sheet.columns
            ):
                ws.write(2, j, ten_cot, header)

                if ten_cot == "Ho va ten":
                    ws.set_column(j, j, 25)
                elif ten_cot == "Ma NV":
                    ws.set_column(j, j, 12)
                else:
                    ws.set_column(j, j, 22)

                if ten_cot not in [
                    "Ma NV", "Ho va ten",
                    "Chi tieu"
                ]:
                    if ten_cot == "So nguoi phu thuoc":
                        ws.set_column(j, j, 22, dong)
                    else:
                        ws.set_column(j, j, 22, tien)

        ws = writer.sheets["Tong hop"]
        ws.set_column(0, 0, 32)
        ws.set_column(1, 1, 25)

    output.seek(0)
    return output.getvalue()

# 1. LUA CHON SO LUONG NHAN VIEN
st.header("1. SO LUONG NGUOI LAO DONG")

SOLUONG = st.number_input(
    "Chon so luong nhan vien can tinh",
    min_value=1,
    max_value=500,
    value=5,
    step=1
)

# DIEU CHINH SO DONG THEO SO LUONG
df = st.session_state.bang_luong.copy()

if SOLUONG > len(df):
    moi = pd.DataFrame([
        tao_nhan_vien(i)
        for i in range(len(df) + 1, SOLUONG + 1)
    ])
    df = pd.concat(
        [df, moi], ignore_index=True
    )

elif SOLUONG < len(df):
    df = df.iloc[:SOLUONG].copy()

st.session_state.bang_luong = (
    df.reset_index(drop=True)
)

st.caption(
    f"Dang tinh luong cho {SOLUONG} nhan vien. "
    "Co the chinh sua truc tiep trong bang."
)

# 2. BANG NHAP THONG TIN
st.header("2. NHAP THONG TIN NHAN VIEN")

st.data_editor(
    st.session_state.bang_luong,
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",
    height=430,
    key="bang_nhap",
    on_change=luu_chinh_sua,
    column_config={
        "Ma NV": st.column_config.TextColumn(
            "Ma nhan vien"
        ),
        "Ho va ten": st.column_config.TextColumn(
            "Ho va ten"
        ),
        "Luong co ban": st.column_config.NumberColumn(
            "Luong co ban",
            min_value=0.0,
            format="%.3f"
        ),
        "Tien thuong": st.column_config.NumberColumn(
            "Tien thuong",
            min_value=0.0,
            format="%.3f"
        ),
        "Phu cap co BHXH": st.column_config.NumberColumn(
            "Phu cap co BHXH",
            min_value=0.0,
            format="%.3f"
        ),
        "Phu cap khong BHXH": st.column_config.NumberColumn(
            "Phu cap khong BHXH",
            min_value=0.0,
            format="%.3f"
        ),
        "So nguoi phu thuoc": st.column_config.NumberColumn(
            "So nguoi phu thuoc",
            min_value=0,
            step=1,
            format="%d"
        )
    }
)

# TINH TOAN TOAN BO NHAN VIEN
KETQUA = tinh_bang_luong(
    st.session_state.bang_luong
)

# 3. TONG HOP
st.divider()
st.header("3. TONG HOP BANG LUONG")

TONGTHUNHAP = KETQUA["Tong thu nhap"].sum()
TONGBHXH = KETQUA["Tien dong BHXH"].sum()
TONGTHUE = KETQUA["Thue TNCN"].sum()
TONGTHUCLANH = KETQUA["Luong thuc lanh"].sum()

c1, c2, c3, c4 = st.columns(4)

c1.metric(
    "So nhan vien",
    f"{len(KETQUA)} nguoi"
)

c2.metric(
    "Tong thu nhap",
    f"{TONGTHUNHAP:,.3f} trieu"
)

c3.metric(
    "Tong thue TNCN",
    f"{TONGTHUE:,.3f} trieu"
)

c4.metric(
    "Tong BHXH",
    f"{TONGBHXH:,.3f} trieu"
)

st.success(
    f"💰 TONG LUONG THUC LANH: "
    f"{TONGTHUCLANH:,.3f} TRIEU DONG"
)

# 4. BANG KET QUA CHI TIET
st.header("4. KET QUA TUNG NHAN VIEN")

st.dataframe(
    KETQUA,
    use_container_width=True,
    hide_index=True,
    height=430
)

# 5. TIM KIEM NHAN VIEN
st.header("5. TRA CUU NHAN VIEN")

TIMKIEM = st.text_input(
    "Nhap ma nhan vien hoac ho ten"
)

if TIMKIEM:
    ket_qua_tim = KETQUA[
        KETQUA["Ma NV"].astype(str).str.contains(
            TIMKIEM, case=False, regex=False
        )
        |
        KETQUA["Ho va ten"].astype(str).str.contains(
            TIMKIEM, case=False, regex=False
        )
    ]

    st.dataframe(
        ket_qua_tim,
        use_container_width=True,
        hide_index=True
    )

# 6. BIEU DO
st.header("6. CO CAU QUY LUONG")

bieudo = pd.DataFrame({
    "Khoan muc": [
        "Luong thuc lanh",
        "Thue TNCN",
        "BHXH"
    ],
    "So tien": [
        TONGTHUCLANH,
        TONGTHUE,
        TONGBHXH
    ]
})

st.bar_chart(
    bieudo,
    x="Khoan muc",
    y="So tien"
)

# 7. XUAT EXCEL
st.header("7. XUAT BANG LUONG EXCEL")

FILEEXCEL = xuat_excel(KETQUA)

st.download_button(
    label="📥 TAI FILE EXCEL BANG LUONG",
    data=FILEEXCEL,
    file_name=(
        "Bang_luong_Song_Hao_"
        + datetime.now().strftime("%Y%m%d")
        + ".xlsx"
    ),
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    use_container_width=True
)

st.caption(
    "Phan mem tinh luong Song Hao - 2026. "
    "Ho tro toi da 500 nguoi lao dong. "
    "So lieu tinh theo cac gia dinh da nhap."
)

st.warning(
    "Luu y: Cong cu dang giu nguyen cong thuc "
    "thue va bao hiem cua bai tap. "
    "Chua xu ly phu cap mien thue, tran dong "
    "bao hiem va cac truong hop dac biet. "
    "Can kiem tra quy dinh truoc khi dung "
    "cho bang luong thuc te."
)
