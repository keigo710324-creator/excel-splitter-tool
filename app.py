import streamlit as st
import pandas as pd
import io
import zipfile

st.set_page_config(page_title="Excel 自動切割助手", layout="wide")
st.title("✂️ Excel 自動切割助手 (完整版)")

# --- 步驟一：上傳檔案 ---
uploaded_file = st.file_uploader("請選擇 Excel 檔案 (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success(f"成功讀取檔案！")

    # --- 步驟二：設定參數 ---
    col1, col2 = st.columns(2)
    with col1:
        split_column = st.selectbox("請選擇要切割的欄位：", df.columns)
    with col2:
        mode = st.radio("請選擇分割方式：", ["分割到不同工作表 (Sheets)", "分割到不同檔案 (ZIP打包)"])

    # --- 步驟三：執行與下載 ---
    if st.button("🚀 開始執行自動切割"):
        output = io.BytesIO()
        
        if mode == "分割到不同工作表 (Sheets)":
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for category, data in df.groupby(split_column):
                    data.to_excel(writer, sheet_name=str(category)[:30], index=False)
            
            st.download_button("📥 下載 Excel 成果", data=output.getvalue(), file_name="split_sheets.xlsx")

        elif mode == "分割到不同檔案 (ZIP打包)":
            # 建立一個緩衝區來存放 ZIP 檔
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for category, data in df.groupby(split_column):
                    # 將每個分組轉成 Excel 二進位資料
                    excel_buffer = io.BytesIO()
                    data.to_excel(excel_buffer, index=False)
                    # 將該 Excel 加入 ZIP
                    zip_file.writestr(f"{category}.xlsx", excel_buffer.getvalue())
            
            st.success("檔案已成功打包成 ZIP！")
            st.download_button("📥 下載所有檔案 (ZIP)", data=zip_buffer.getvalue(), file_name="split_files.zip")
