
import io
import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Mục 09 - Tổng hợp chuyển tiền theo mục đích", layout="wide")
st.title("📦 Mục 09 — Tổng hợp theo PART_NAME & Mục đích chuyển tiền (3 năm gần nhất)")

st.markdown('''
Tải lên file Excel **MUC 09** với các cột:
- `TRAN_DATE` (ngày giao dịch)
- `TRAN_ID` (mã giao dịch)
- `PART_NAME` (đối tượng nhận/gửi)
- `PURPOSE_OF_REMITTANCE` (mục đích chuyển tiền)
- `QUY_DOI_USD` (số tiền quy đổi USD)
''')

file = st.file_uploader("Chọn file Excel", type=["xlsx","xls"])

with st.expander("⚙️ Tuỳ chỉnh cột (nếu tên cột khác)"):
    col_date = st.text_input("Cột ngày giao dịch", "TRAN_DATE")
    col_tranid = st.text_input("Cột mã giao dịch", "TRAN_ID")
    col_party = st.text_input("Cột PART_NAME", "PART_NAME")
    col_purpose = st.text_input("Cột PURPOSE_OF_REMITTANCE", "PURPOSE_OF_REMITTANCE")
    col_amount = st.text_input("Cột số tiền (USD)", "QUY_DOI_USD")

run = st.button("▶️ Chạy tổng hợp")

def read_excel_any(file):
    \"\"\"Đọc .xls bằng calamine, .xlsx bằng openpyxl; tự fallback an toàn.\"\"\"
    # streamlit UploadedFile có thuộc tính name, readable như file-like
    name = getattr(file, "name", str(file))
    ext = os.path.splitext(name)[1].lower()
    if ext == ".xls":
        try:
            return pd.read_excel(file, engine="calamine")
        except Exception as e:
            st.warning(f"Đọc .xls bằng calamine lỗi: {e}. Thử lại với engine mặc định.")
            return pd.read_excel(file)  # có thể vẫn là calamine nếu cài đặt mặc định
    else:
        try:
            return pd.read_excel(file, engine="openpyxl")
        except Exception as e:
            st.warning(f"Đọc .xlsx bằng openpyxl lỗi: {e}. Thử engine calamine.")
            return pd.read_excel(file, engine="calamine")

def build_output(df: pd.DataFrame, col_date, col_tranid, col_party, col_purpose, col_amount):
    # Chuẩn hoá ngày và năm
    df = df.copy()
    df[col_date] = pd.to_datetime(df[col_date], errors="coerce")
    df["YEAR"] = df[col_date].dt.year

    # Xác định 3 năm gần nhất theo dữ liệu thật có
    years = sorted([int(y) for y in df["YEAR"].dropna().unique()])
    if not years:
        return pd.DataFrame(), []
    nam_T = years[-1]
    cac_nam = [y for y in years if y >= nam_T - 2][-3:]  # <= 3 năm gần nhất

    # Loại trùng
    df = df.drop_duplicates(subset=[col_party, col_purpose, col_date, col_tranid])

    ket_qua = pd.DataFrame()
    ds_muc_dich = df[col_purpose].dropna().unique()

    for muc_dich in ds_muc_dich:
        df_md = df[df[col_purpose] == muc_dich]
        for nam in cac_nam:
            df_year = df_md[df_md["YEAR"] == nam]
            if df_year.empty:
                continue
            pivot = df_year.groupby(col_party).agg(
                tong_lan_nhan=(col_tranid, "count"),
                tong_tien_usd=(col_amount, "sum")
            ).reset_index()

            col_lan = f"{muc_dich}_LAN_{nam}"
            col_tien = f"{muc_dich}_TIEN_{nam}"
            pivot.rename(columns={"tong_lan_nhan": col_lan, "tong_tien_usd": col_tien}, inplace=True)

            ket_qua = pivot if ket_qua.empty else pd.merge(ket_qua, pivot, on=col_party, how="outer")

    # Điền NaN & ép kiểu
    for c in ket_qua.columns:
        if "_LAN_" in c:
            ket_qua[c] = ket_qua[c].fillna(0).astype(int)
        elif "_TIEN_" in c:
            ket_qua[c] = ket_qua[c].fillna(0.0).astype(float)

    # Đưa PART_NAME (hoặc cột party) lên đầu
    if col_party in ket_qua.columns:
        cols = [col_party] + [c for c in ket_qua.columns if c != col_party]
        ket_qua = ket_qua[cols]

    return ket_qua, cac_nam

if run:
    if not file:
        st.warning("Hãy chọn file Excel trước khi chạy.")
        st.stop()

    try:
        df = read_excel_any(file)

        # Kiểm tra cột yêu cầu
        missing = [c for c in [col_date, col_tranid, col_party, col_purpose, col_amount] if c not in df.columns]
        if missing:
            st.error(f"Thiếu cột bắt buộc: {missing}")
            st.stop()

        ket_qua, cac_nam = build_output(df, col_date, col_tranid, col_party, col_purpose, col_amount)

        if ket_qua.empty:
            st.info("Không có dữ liệu phù hợp để tổng hợp.")
        else:
            st.success(f"Tổng hợp xong cho các năm: {', '.join(map(str, cac_nam))}")
            st.dataframe(ket_qua.head(200), use_container_width=True)

            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                ket_qua.to_excel(writer, sheet_name="tong_hop", index=False)
            st.download_button(
                "⬇️ Tải Excel tổng hợp",
                data=bio.getvalue(),
                file_name="tong_hop_chuyen_tien.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error("Đã xảy ra lỗi khi xử lý.")
        st.exception(e)
