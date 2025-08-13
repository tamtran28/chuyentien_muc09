
import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Mục 09 - Tổng hợp chuyển tiền theo mục đích", layout="wide")
st.title("📦 Mục 09 — Tổng hợp theo PART_NAME & Mục đích chuyển tiền (3 năm gần nhất)")

st.markdown('''
Tải lên file Excel **MUC 09.xlsx** (hoặc tương tự) với các cột:
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

def build_output(df: pd.DataFrame, col_date, col_tranid, col_party, col_purpose, col_amount):
    # Chuẩn hoá ngày và năm
    df = df.copy()
    df[col_date] = pd.to_datetime(df[col_date], errors="coerce")
    df["YEAR"] = df[col_date].dt.year

    # Xác định 3 năm gần nhất theo dữ liệu thật có
    years_sorted = sorted(df["YEAR"].dropna().unique())
    if not years_sorted:
        return pd.DataFrame(), []
    nam_T = years_sorted[-1]
    cac_nam = [y for y in years_sorted if y >= nam_T - 2][-3:]  # bảo vệ nếu dữ liệu không đủ 3 năm

    # Loại trùng
    df = df.drop_duplicates(subset=[col_party, col_purpose, col_date, col_tranid])

    ket_qua = pd.DataFrame()
    ds_muc_dich = df[col_purpose].dropna().unique()

    for muc_dich in ds_muc_dich:
        df_muc_dich = df[df[col_purpose] == muc_dich]
        for nam in cac_nam:
            df_nam = df_muc_dich[df_muc_dich["YEAR"] == nam]
            if df_nam.empty:
                continue
            pivot = df_nam.groupby(col_party).agg(
                tong_lan_nhan=(col_tranid, "count"),
                tong_tien_usd=(col_amount, "sum")
            ).reset_index()
            col_lan = f"{muc_dich}_LAN_{nam}"
            col_tien = f"{muc_dich}_TIEN_{nam}"
            pivot.rename(columns={
                "tong_lan_nhan": col_lan,
                "tong_tien_usd": col_tien
            }, inplace=True)
            ket_qua = pivot if ket_qua.empty else pd.merge(ket_qua, pivot, on=col_party, how="outer")

    # Fillna & kiểu dữ liệu
    for c in ket_qua.columns:
        if "_LAN_" in c:
            ket_qua[c] = ket_qua[c].fillna(0).astype(int)
        elif "_TIEN_" in c:
            ket_qua[c] = ket_qua[c].fillna(0.0).astype(float)

    # Sắp xếp cột: PART_NAME trước, sau đó theo từng mục đích/năm
    if col_party in ket_qua.columns:
        cols = [col_party] + [c for c in ket_qua.columns if c != col_party]
        ket_qua = ket_qua[cols]

    return ket_qua, cac_nam

if run:
    if file is None:
        st.warning("Hãy chọn file Excel trước khi chạy.")
        st.stop()

    try:
        # Ưu tiên engine calamine để đọc cả .xls/.xlsx nếu có
        try:
            df = pd.read_excel(file, engine="calamine")
        except Exception:
            df = pd.read_excel(file)  # fallback

        # Kiểm tra cột
        missing = [c for c in [col_date, col_tranid, col_party, col_purpose, col_amount] if c not in df.columns]
        if missing:
            st.error(f"Thiếu cột bắt buộc: {missing}")
            st.stop()

        ket_qua, cac_nam = build_output(df, col_date, col_tranid, col_party, col_purpose, col_amount)

        if ket_qua.empty:
            st.info("Không có dữ liệu phù hợp để tổng hợp (có thể thiếu 3 năm hoặc dữ liệu rỗng).")
        else:
            st.success(f"Tổng hợp xong cho các năm: {', '.join(map(str, cac_nam))}")
            st.dataframe(ket_qua.head(200))

            # Xuất Excel
            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                ket_qua.to_excel(writer, sheet_name="tong_hop", index=False)
            st.download_button(
                "⬇️ Tải Excel tổng hợp",
                data=bio.getvalue(),
                file_name="tong_hop_chuyen_tien.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.exception(e)
