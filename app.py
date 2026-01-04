import streamlit as st
import pandas as pd
import io
import zipfile
import re

# 設定網頁
st.set_page_config(page_title="Excel 自動切割助手", page_icon="✂️")

# 1. 定義「超強濾網」函數：清理 Excel 不接受的字元
def clean_name(name):
    # 將所有非法字元 \ / ? * : [ ] 替換成底線 _
    clean = re.sub(r'[\\/*?:\[\]]', "_", str(name))
    # 確保長度不超過 31 個字元
    return clean[:31]

st.title("📊 Excel 自動切割助手")

uploaded_file = st.file_uploader("請選擇 Excel 檔案", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    split_column = st.selectbox("請選擇切割欄位：", df.columns)
    mode = st.radio("分割方式：", ["不同工作表 (Sheets)", "不同檔案 (ZIP)"])

    if st.button("🚀 開始執行"):
        output = io.BytesIO()
        try:
            if mode == "不同工作表 (Sheets)":
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for category, data in df.groupby(split_column):
                        # --- 關鍵修正：套用清理函數 ---
                        safe_sheet_name = clean_name(category)
                        data.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                st.success("切割完成！")
                st.download_button("📥 下載 Excel", output.getvalue(), "result.xlsx")

            else:
                with zipfile.ZipFile(output, "a") as zip_file:
                    for category, data in df.groupby(split_column):
                        excel_buf = io.BytesIO()
                        data.to_excel(excel_buf, index=False)
                        # 檔名也清理一下比較保險
                        safe_file_name = f"{clean_name(category)}.xlsx"
                        zip_file.writestr(safe_file_name, excel_buf.getvalue())
                st.success("打包完成！")
                st.download_button("📥 下載 ZIP", output.getvalue(), "files.zip")
        except Exception as e:
            st.error(f"仍有錯誤發生：{e}")
