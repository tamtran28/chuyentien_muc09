import io
import os
import pandas as pd
import streamlit as st

# ================== CONFIG ==================
st.set_page_config(page_title="Mục 09 - Tổng hợp chuyển tiền", layout="wide")
st.title("📦 Mục 09 — Tổng hợp theo PART_NAME & Mục đích chuyển tiền (3 năm gần nhất)")
st.caption("Đọc được cả .xls và .xlsx (không cần xlrd).")

st.markdown(
    """
**Yêu cầu cột dữ liệu** (có thể đổi tên ở phần ⚙️):
- `TRAN_DATE` (ngày giao dịch)
- `TRAN_ID` (mã giao dịch)
- `PART_NAME` (bên liên quan)
- `PURPOSE_OF_REMITTANCE` (mục đích chuyển tiền)
- `QUY_DOI_USD` (số tiền USD)
"""
)

# ================== INPUTS ==================
file = st.file_uploader("Chọn file Excel", type=["xls", "xlsx"])

with st.expander("⚙️ Tuỳ chỉnh cột (nếu khác tên mặc định)"):
    col_date = st.text_input("Cột ngày giao dịch", "TRAN_DATE")
    col_tranid = st.text_input("Cột mã giao dịch", "TRAN_ID")
    col_party = st.text_input("Cột PART_NAME", "PART_NAME")
    col_purpose = st.text_input("Cột PURPOSE_OF_REMITTANCE", "PURPOSE_OF_REMITTANCE")
    col_amount = st.text_input("Cột số tiền (USD)", "QUY_DOI_USD")

run = st.button("▶️ Chạy tổng hợp")

# ================== HELPERS ==================
def read_excel_any(uploaded_file):
    """
    Đọc .xls bằng calamine, .xlsx bằng openpyxl; nếu lỗi sẽ tự fallback qua engine còn lại.
    """
    name = getattr(uploaded_file, "name", str(uploaded_file))
    ext = os.path.splitext(name)[1].lower()

    if ext == ".xls":
        # Ưu tiên calamine cho .xls
        try:
            return pd.read_excel(uploaded_file, engine="calamine")
        except Exception as e:
            st.warning(f"Đọc .xls bằng calamine lỗi: {e}. Thử engine mặc định.")
            return pd.read_excel(uploaded_file)  # có thể vẫn là calamine nếu pandas cấu hình mặc định
    else:
        # Ưu tiên openpyxl cho .xlsx
        try:
            return pd.read_excel(uploaded_file, engine="openpyxl")
        except Exception as e:
            st.warning(f"Đọc .xlsx bằng openpyxl lỗi: {e}. Thử engine calamine.")
            return pd.read_excel(uploaded_file, engine="calamine")

def build_output(df: pd.DataFrame, col_date, col_tranid, col_party, col_purpose, col_amount):
    # Chuẩn hoá ngày & năm
    df = df.copy()
    df[col_date] = pd.to_datetime(df[col_date], errors="coerce")
    df["YEAR"] = df[col_date].dt.year

    # Lấy 3 năm gần nhất theo dữ liệu có thật
    years = sorted([int(y) for y in df["YEAR"].dropna().unique()])
    if not years:
        return pd.DataFrame(), []
    nam_T = years[-1]
    cac_nam = [y for y in years if y >= nam_T - 2][-3:]  # bảo vệ khi không đủ 3 năm

    # Loại trùng theo đúng logic
    df = df.drop_duplicates(subset=[col_party, col_purpose, col_date, col_tranid])

    ket_qua = pd.DataFrame()
    ds_muc_dich = df[col_purpose].dropna().unique()

    for muc_dich in ds_muc_dich:
        df_md = df[df[col_purpose] == muc_dich]
        for nam in cac_nam:
            df_y = df_md[df_md["YEAR"] == nam]
            if df_y.empty:
                continue

            pivot = (
                df_y.groupby(col_party)
                .agg(
                    tong_lan_nhan=(col_tranid, "count"),
                    tong_tien_usd=(col_amount, "sum"),
                )
                .reset_index()
            )

            # Đặt tên cột theo yêu cầu
            col_lan = f"{muc_dich}_LAN_{nam}"
            col_tien = f"{muc_dich}_TIEN_{nam}"
            pivot.rename(
                columns={"tong_lan_nhan": col_lan, "tong_tien_usd": col_tien}, inplace=True
            )

            ket_qua = pivot if ket_qua.empty else pd.merge(ket_qua, pivot, on=col_party, how="outer")

    # Fill & ép kiểu
    for c in ket_qua.columns:
        if "_LAN_" in c:
            ket_qua[c] = ket_qua[c].fillna(0).astype(int)
        elif "_TIEN_" in c:
            ket_qua[c] = ket_qua[c].fillna(0.0).astype(float)

    # Đưa cột party lên đầu
    if col_party in ket_qua.columns:
        ket_qua = ket_qua[[col_party] + [c for c in ket_qua.columns if c != col_party]]

    return ket_qua, cac_nam

# ================== RUN ==================
if run:
    if not file:
        st.error("Vui lòng chọn file Excel trước khi chạy.")
        st.stop()

    try:
        df = read_excel_any(file)

        # Kiểm tra cột
        required = [col_date, col_tranid, col_party, col_purpose, col_amount]
        missing = [c for c in required if c not in df.columns]
        if missing:
            st.error(f"Thiếu cột bắt buộc: {missing}")
            st.stop()

        ket_qua, cac_nam = build_output(df, col_date, col_tranid, col_party, col_purpose, col_amount)

        if ket_qua.empty:
            st.info("Không có dữ liệu phù hợp để tổng hợp.")
        else:
            st.success(f"Tổng hợp xong cho các năm: {', '.join(map(str, cac_nam))}")
            st.dataframe(ket_qua, use_container_width=True)

            # Xuất Excel
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
        st.error("Đã xảy ra lỗi khi xử lý:")
        st.exception(e)
