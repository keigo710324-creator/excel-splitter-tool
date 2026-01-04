import streamlit as st
import pandas as pd
import io
import zipfile
import re

# 設定網頁外觀
st.set_page_config(page_title="Excel 自動切割助手", page_icon="✂️", layout="wide")

# 漂亮標題
st.markdown("<h1 style='text-align: center; color: #4A90E2;'>📊 Excel 自動切割助手</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>由 Gemini 協助製作：簡單快速地分割您的數據資料</p>", unsafe_allow_html=True)
st.divider()

# 定義清理 Excel 工作表名稱的函數
def clean_sheet_name(name):
    # 移除 Excel 不允許的字元: \ / ? * : [ ]
    clean_name = re.sub(r'[\\/*?:\[\]]', "_", str(name))
    # 限制長度為 31 字元
    return clean_name[:31]

# --- 步驟一：上傳檔案 ---
st.header("步驟一：上傳檔案")
uploaded_file = st.file_uploader("請選擇 Excel 檔案 (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success(f"成功讀取檔案！共有 {len(df)} 筆資料。")

    # --- 步驟二：設定參數 ---
    st.header("步驟二：設定參數")
    col1, col2 = st.columns(2)
    with col1:
        split_column = st.selectbox("請選擇要切割的欄位：", df.columns)
    with col2:
        mode = st.radio("請選擇分割方式：", ["分割到不同工作表 (Sheets)", "分割到不同檔案 (ZIP打包)"])

    # --- 步驟三：執行與下載 ---
    st.header("步驟三：執行並下載")
    if st.button("🚀 開始執行自動切割"):
        output = io.BytesIO()
        
        try:
            if mode == "分割到不同工作表 (Sheets)":
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for category, data in df.groupby(split_column):
                        # 使用清理過的名稱
                        sheet_name = clean_sheet_name(category)
                        data.to_excel(writer, sheet_name=sheet_name, index=False)
                
                st.success("切割完成！")
                st.download_button("📥 下載 Excel 成果", data=output.getvalue(), file_name=f"split_by_{split_column}.xlsx")

            elif mode == "分割到不同檔案 (ZIP打包)":
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for category, data in df.groupby(split_column):
                        excel_buffer = io.BytesIO()
                        data.to_excel(excel_buffer, index=False)
                        # 檔案名稱也進行清理
                        safe_name = clean_sheet_name(category)
                        zip_file.writestr(f"{safe_name}.xlsx", excel_buffer.getvalue())
                
                st.success("檔案已成功打包成 ZIP！")
                st.download_button("📥 下載所有檔案 (ZIP)", data=zip_buffer.getvalue(), file_name="split_files.zip")
        except Exception as e:
            st.error(f"執行時發生錯誤：{e}")
            st.info("提示：請檢查選擇的欄位內容是否包含過多特殊字元。")
