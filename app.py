import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.worksheet.page import PageMargins
import io

st.set_page_config(page_title="全校通用購書單產生器", layout="centered")

st.title("📚 全校通用購書費通知單產生器")
st.markdown("支援一至三年級分組差異，並自動處理「語文/數理資優生」免購邏輯。")
st.divider()

# --- 1. 檔案上傳區 ---
st.subheader("Step 1: 上傳檔案")
st.info("請確保您的 Excel 檔案包含指定的欄位名稱，系統才能正確判讀。")

col1, col2 = st.columns(2)
with col1:
    student_file = st.file_uploader("上傳「學生名單」Excel", type=["xlsx"])
    st.caption("必要欄位：座號, 姓名, 資優類別(語資/數資), 英組, 數組, 自組")
with col2:
    book_file = st.file_uploader("上傳「書商報價單」Excel", type=["xlsx"])
    st.caption("必要欄位：商品名稱, 科目(國/英/數/自/社會/其他), 分組代號(1/1A/2B...), 單價")

# --- 2. 核心判斷邏輯 ---
def should_buy_book(b_subj, b_code, s_gifted, s_eng, s_math, s_sci):
    """判斷該名學生是否需要購買該本書籍的通用邏輯"""
    
    # 規則 A：資優生免買邏輯
    s_gifted = str(s_gifted).strip()
    if s_gifted == "語資" and b_subj in ["國", "英"]:
        return False
    if s_gifted == "數資" and b_subj in ["數", "自"]:
        return False
        
    # 規則 B：全班共同書目
    if b_code in ["1", "全"]:
        return True
        
    # 規則 C：分組書目比對
    # (如果是一年級，學生的 s_eng 等於 "1"，書本不是 "1" 就不會買到)
    if b_subj == "英" and str(b_code) == str(s_eng): return True
    if b_subj == "數" and str(b_code) == str(s_math): return True
    if b_subj == "自" and str(b_code) == str(s_sci): return True
    
    return False

# --- 3. 處理與排版 (略過複雜排版細節，展示流程) ---
def generate_excel(df_students, df_books):
    wb = Workbook()
    ws = wb.active
    ws.title = "購書通知單(4張一頁)"
    
    # ... (這裡放入我們之前寫好的 2x2 排版與 openpyxl 畫線邏輯) ...
    
    # 雙層迴圈範例：列出每個學生要買的書
    for index, student in df_students.iterrows():
        subtotal = 0
        for _, book in df_books.iterrows():
            if should_buy_book(
                book['科目'], book['分組代號'], 
                student['資優類別'], student['英組'], student['數組'], student['自組']
            ):
                subtotal += book['單價']
                # (實際程式碼會將書名與價格寫入對應的 Excel 儲存格)
        
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 4. 執行與下載區 ---
st.divider()
st.subheader("Step 2: 產生與下載")

if student_file and book_file:
    if st.button("🚀 開始比對並產生通知單", type="primary"):
        with st.spinner("系統正在運算與排版中..."):
            # 讀取上傳的 Excel
            df_s = pd.read_excel(student_file).fillna("") # 把空白填補起來
            df_b = pd.read_excel(book_file).fillna("")
            
            # 產出最終 Excel
            final_excel = generate_excel(df_s, df_b)
            
        st.success("🎉 排版完成！")
        st.download_button(
            label="📥 下載通知單列印檔 (A4 2x2排版)",
            data=final_excel,
            file_name="全校購書通知單_排版完成.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("請先在上方上傳名單與書單檔案。")
