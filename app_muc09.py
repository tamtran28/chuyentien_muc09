import pandas as pd
import streamlit as st
import io

st.set_page_config(page_title="Tổng hợp Mục 09", layout="wide")
st.title("📊 Tổng hợp chuyển tiền Mục 09")

# Hàm đọc Excel không dùng calamine
def read_excel_any(uploaded_file):
    filename = uploaded_file.name.lower()
    if filename.endswith(".xlsx"):
        return pd.read_excel(uploaded_file, engine="openpyxl")
    else:
        st.error("❌ File .xls không được hỗ trợ. Vui lòng lưu lại thành .xlsx rồi tải lên.")
        return None

uploaded_file = st.file_uploader("Tải file MUC 09 (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = read_excel_any(uploaded_file)
    if df is not None:
        # Xử lý dữ liệu
        df['TRAN_DATE'] = pd.to_datetime(df['TRAN_DATE'], errors='coerce')
        df['YEAR'] = df['TRAN_DATE'].dt.year

        nam_max = df['YEAR'].max()
        nam_T = nam_max
        nam_T1 = nam_T - 1
        nam_T2 = nam_T - 2

        # Loại bỏ PART_NAME trùng
        df = df.drop_duplicates(subset=['PART_NAME', 'PURPOSE_OF_REMITTANCE', 'TRAN_DATE', 'TRAN_ID'])

        ket_qua = pd.DataFrame()
        cac_nam = [nam_T2, nam_T1, nam_T]
        ds_muc_dich = df['PURPOSE_OF_REMITTANCE'].dropna().unique()

        for muc_dich in ds_muc_dich:
            df_muc_dich = df[df['PURPOSE_OF_REMITTANCE'] == muc_dich]

            for nam in cac_nam:
                df_nam = df_muc_dich[df_muc_dich['YEAR'] == nam]
                if df_nam.empty:
                    continue

                pivot = df_nam.groupby('PART_NAME').agg(
                    tong_lan_nhan=('TRAN_ID', 'count'),
                    tong_tien_usd=('QUY_DOI_USD', 'sum')
                ).reset_index()

                col_lan = f"{muc_dich}_LAN_{nam}"
                col_tien = f"{muc_dich}_TIEN_{nam}"
                pivot.rename(columns={
                    'tong_lan_nhan': col_lan,
                    'tong_tien_usd': col_tien
                }, inplace=True)

                if ket_qua.empty:
                    ket_qua = pivot
                else:
                    ket_qua = pd.merge(ket_qua, pivot, on='PART_NAME', how='outer')

        for col in ket_qua.columns:
            if "_LAN_" in col:
                ket_qua[col] = ket_qua[col].fillna(0).astype(int)
            elif "_TIEN_" in col:
                ket_qua[col] = ket_qua[col].fillna(0.0).astype(float)

        # Hiển thị kết quả
        st.dataframe(ket_qua)

        # Xuất file Excel
        output = io.BytesIO()
        ket_qua.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="📥 Tải file kết quả",
            data=output,
            file_name="tong_hop_chuyen_tien_Muc09.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
