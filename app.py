import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.worksheet.page import PageMargins
import io

# 設定網頁標題與寬度
st.set_page_config(page_title="購書通知單產生器", layout="centered")

st.title("📚 班級購書費通知單自動產生器")
st.markdown("把繁瑣的書單比對交給系統，一鍵產出 A4 裁切版通知單！")
st.divider()

# --- 1. 檔案上傳區 ---
st.subheader("Step 1: 上傳班級資料與書單")
st.info("💡 提示：未來這裡可以讓老師上傳他們班的 Excel 名單與書商報價單。為了示範，目前點擊下方按鈕會直接使用內建的 801 班測試資料。")

# --- 2. 核心處理函式 (包裝我們寫好的邏輯) ---
def generate_receipts_excel():
    # 這裡放我們稍早確認過的 books_raw 和 students_updated 資料
    # (為了程式碼簡潔，我先縮減示意，您測試時可以把完整的 list 貼進來)
    students_updated = [
        [1, "王佑予", "1", "1A", "1A"], [2, "何承恩", "1", "6B", "6B"],
        [10, "陳昱學", "資", "1A", "資"] # 示範包含資優生
    ]
    books_raw = [
        ["國文科八年級小卷", "國", "1", 57], ["康軒 文法即時通", "英", "1A", 105],
        ["鼎甲國中 良師講義{康}自然(4)", "自", "1", 171], ["美術材料費", "其他", "1", 90]
    ]

    # ... 中間省略：計算數量、配置版面、2x2 網格排版邏輯 (直接複製我們上一版的迴圈內容) ...
    # 為了示範網頁運作，這邊建立一個簡單的 Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "購書通知單(4張一頁)"
    ws.cell(row=1, column=1, value="這是由 Streamlit 產生的通知單！").font = Font(size=14, bold=True)
    # -------------------------------------------------------------------------

    # 關鍵差異：在網頁版中，我們不直接 save 成實體檔案，而是存入記憶體 (BytesIO)
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 3. 執行與下載區 ---
st.subheader("Step 2: 產生與下載")

if st.button("🚀 點我開始產生通知單", type="primary"):
    with st.spinner("系統正在排版中，請稍候..."):
        # 呼叫處理函式
        excel_data = generate_receipts_excel()
        
    st.success("🎉 排版完成！請點擊下方按鈕下載。")
    
    # 顯示下載按鈕
    st.download_button(
        label="📥 下載 Excel 列印檔 (A4 2x2排版)",
        data=excel_data,
        file_name="班級購書通知單_排版完成.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
