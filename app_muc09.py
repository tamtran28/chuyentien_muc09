import io
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Mục 09 - Tổng hợp chuyển tiền", layout="wide")
st.title("📊 Mục 09 — Tổng hợp theo PART_NAME & Mục đích (3 năm gần nhất)")
st.caption("Bản đơn giản, chỉ hỗ trợ .xlsx (engine openpyxl).")

# --------- Cấu hình tên cột mặc định ----------
DEFAULT_COLS = {
    "date":  "TRAN_DATE",
    "id":    "TRAN_ID",
    "party": "PART_NAME",
    "purpose": "PURPOSE_OF_REMITTANCE",
    "amount": "QUY_DOI_USD",
}

with st.expander("⚙️ Tuỳ chỉnh tên cột (nếu file của bạn khác)"):
    COL_DATE   = st.text_input("Cột ngày giao dịch", DEFAULT_COLS["date"])
    COL_ID     = st.text_input("Cột mã giao dịch", DEFAULT_COLS["id"])
    COL_PART   = st.text_input("Cột PART_NAME", DEFAULT_COLS["party"])
    COL_PURP   = st.text_input("Cột PURPOSE_OF_REMITTANCE", DEFAULT_COLS["purpose"])
    COL_AMT    = st.text_input("Cột QUY_DOI_USD (số tiền)", DEFAULT_COLS["amount"])

uploaded = st.file_uploader("Tải file Excel (.xlsx)", type=["xlsx"])
run = st.button("▶️ Chạy tổng hợp")

def read_xlsx_openpyxl(uploaded_file) -> pd.DataFrame | None:
    """Chỉ đọc .xlsx bằng openpyxl. Báo lỗi nếu không đúng định dạng."""
    if not uploaded_file:
        return None
    name = uploaded_file.name.lower()
    if not name.endswith(".xlsx"):
        st.error("❌ Chỉ hỗ trợ .xlsx. Hãy lưu file .xls thành .xlsx rồi tải lên.")
        return None
    try:
        return pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Không đọc được file .xlsx: {e}")
        return None

def build_output(df: pd.DataFrame) -> tuple[pd.DataFrame, list[int]]:
    """Theo đúng logic của bạn: loại trùng 4 trường, gom theo mục đích & năm."""
    # Chuẩn hoá kiểu dữ liệu
    df = df.copy()
    df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors="coerce")
    df["YEAR"]   = df[COL_DATE].dt.year
    df[COL_AMT]  = pd.to_numeric(df[COL_AMT], errors="coerce")

    # Bỏ dòng thiếu ngày/năm
    df = df.dropna(subset=[COL_DATE, "YEAR"])
    if df.empty:
        return pd.DataFrame(), []

    # Loại trùng đúng 4 trường
    df = df.drop_duplicates(subset=[COL_PART, COL_PURP, COL_DATE, COL_ID])

    # 3 năm gần nhất (nếu thiếu vẫn chạy với số năm hiện có)
    years = sorted(df["YEAR"].dropna().astype(int).unique())
    nam_T = years[-1]
    cac_nam = [y for y in years if y >= nam_T - 2][-3:]  # tối đa 3 năm

    ket_qua = pd.DataFrame()
    ds_muc_dich = df[COL_PURP].dropna().astype(str).unique()

    for muc_dich in ds_muc_dich:
        df_md = df[df[COL_PURP] == muc_dich]
        for nam in cac_nam:
            df_y = df_md[df_md["YEAR"] == nam]
            if df_y.empty:
                continue

            pivot = (
                df_y.groupby(COL_PART, dropna=False)
                    .agg(
                        tong_lan_nhan=(COL_ID, "count"),
                        tong_tien_usd=(COL_AMT, "sum"),
                    )
                    .reset_index()
            )

            col_lan  = f"{muc_dich}_LAN_{nam}"
            col_tien = f"{muc_dich}_TIEN_{nam}"
            pivot.rename(
                columns={"tong_lan_nhan": col_lan, "tong_tien_usd": col_tien},
                inplace=True,
            )

            ket_qua = pivot if ket_qua.empty else ket_qua.merge(pivot, on=COL_PART, how="outer")

    # Điền NaN & ép kiểu
    if ket_qua.empty:
        return ket_qua, cac_nam
    for c in ket_qua.columns:
        if c == COL_PART: 
            continue
        if "_LAN_" in c:
            ket_qua[c] = pd.to_numeric(ket_qua[c], errors="coerce").fillna(0).astype(int)
        elif "_TIEN_" in c:
            ket_qua[c] = pd.to_numeric(ket_qua[c], errors="coerce").fillna(0.0).astype(float)

    # Đưa cột PART_NAME lên đầu
    ket_qua = ket_qua[[COL_PART] + [c for c in ket_qua.columns if c != COL_PART]]
    return ket_qua, cac_nam

if run:
    if not uploaded:
        st.warning("Hãy tải file .xlsx trước.")
        st.stop()

    df_raw = read_xlsx_openpyxl(uploaded)
    if df_raw is None:
        st.stop()

    # Kiểm tra cột bắt buộc
    required = [COL_DATE, COL_ID, COL_PART, COL_PURP, COL_AMT]
    missing = [c for c in required if c not in df_raw.columns]
    if missing:
        st.error(f"Thiếu các cột bắt buộc: {missing}")
        st.stop()

    ket_qua, years_used = build_output(df_raw)

    if ket_qua.empty:
        st.info("Không có dữ liệu phù hợp để tổng hợp.")
    else:
        st.success("Tổng hợp xong" + (f" cho các năm: {', '.join(map(str, years_used))}" if years_used else ""))
        st.dataframe(ket_qua, use_container_width=True)

        # Xuất Excel tải về
        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine="openpyxl") as writer:
            ket_qua.to_excel(writer, sheet_name="tong_hop", index=False)
        st.download_button(
            "⬇️ Tải Excel tổng hợp",
            data=bio.getvalue(),
            file_name="tong_hop_chuyen_tien_Muc09.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
