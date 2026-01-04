import streamlit as st
import pandas as pd
import io
import zipfile
import re

# 1. 強力過濾器：將所有 Excel 不允許的符號通通變成底線
def clean_sheet_name(name):
    # 找尋 \ / ? * : [ ] 並替換為 _
    clean_name = re.sub(r'[\\/*?:\[\]]', "_", str(name))
    # 限制長度在 31 字元內 (Excel 極限)
    return clean_name[:31]

st.set_page_config(page_title="Excel 自動切割助手", page_icon="✂️")
st.title("📊 Excel 自動切割助手")
st.write("這是一個自動修復名稱版本，會自動處理非法字元（如斜線、冒號）。")

uploaded_file = st.file_uploader("上傳 Excel 檔案", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    cols = df.columns.tolist()
    split_col = st.selectbox("選擇切割欄位", cols)
    mode = st.radio("輸出模式", ["不同工作表 (Sheets)", "不同檔案 (ZIP打包)"])

    if st.button("🚀 開始執行"):
        output = io.BytesIO()
        try:
            if mode == "不同工作表 (Sheets)":
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for category, data in df.groupby(split_col):
                        # --- 這裡是最關鍵的修正點 ---
                        safe_name = clean_sheet_name(category)
                        data.to_excel(writer, sheet_name=safe_name, index=False)
                st.success("成功！非法字元已自動替換。")
                st.download_button("📥 下載 Excel", output.getvalue(), "split_result.xlsx")

            else:
                with zipfile.ZipFile(output, "a") as zip_file:
                    for category, data in df.groupby(split_col):
                        buf = io.BytesIO()
                        data.to_excel(buf, index=False)
                        # 檔名也套用清理規則
                        safe_filename = f"{clean_sheet_name(category)}.xlsx"
                        zip_file.writestr(safe_filename, buf.getvalue())
                st.success("打包完成！")
                st.download_button("📥 下載 ZIP 包", output.getvalue(), "all_files.zip")
        except Exception as e:
            st.error(f"發生未預期錯誤：{e}")
