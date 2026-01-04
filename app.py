import streamlit as st
import pandas as pd
import io
import zipfile
import re

# 1. 強力清洗函數：處理所有 Excel 命名地雷
def clean_sheet_name(name):
    # 移除非法字元 \ / ? * : [ ] 並替換為底線
    clean = re.sub(r'[\\/*?:\[\]]', "_", str(name))
    # 移除開頭或結尾的單引號 (Excel 也不喜歡)
    clean = clean.strip("'")
    # 限制長度 31 字元 (Excel 極限)
    clean = clean[:31]
    # 如果清洗後變空值，給個預設名
    return clean if clean else "Sheet_Split"

st.set_page_config(page_title="Excel 自動切割助手", page_icon="✂️")
st.title("📊 Excel 自動切割助手 (最終修補版)")

uploaded_file = st.file_uploader("步驟一：上傳檔案", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    split_col = st.selectbox("步驟二：選擇切割欄位", df.columns)
    mode = st.radio("選擇輸出方式", ["分割到不同工作表 (Sheets)", "分割到不同檔案 (ZIP打包)"])

    if st.button("🚀 開始執行自動切割"):
        output = io.BytesIO()
        try:
            if mode == "分割到不同工作表 (Sheets)":
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for category, data in df.groupby(split_col):
                        # --- 這裡使用清洗後的名稱 ---
                        safe_name = clean_sheet_name(category)
                        data.to_excel(writer, sheet_name=safe_name, index=False)
                st.success("切割完成！非法字元已自動處理。")
                st.download_button("📥 下載 Excel 成果", output.getvalue(), "split_result.xlsx")

            else:
                with zipfile.ZipFile(output, "a") as zip_file:
                    for category, data in df.groupby(split_col):
                        buf = io.BytesIO()
                        data.to_excel(buf, index=False)
                        safe_file_name = f"{clean_sheet_name(category)}.xlsx"
                        zip_file.writestr(safe_file_name, buf.getvalue())
                st.success("打包完成！")
                st.download_button("📥 下載 ZIP 包", output.getvalue(), "all_files.zip")
        except Exception as e:
            st.error(f"發生未預期錯誤：{e}")
