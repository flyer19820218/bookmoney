import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.worksheet.page import PageMargins
import re
import io

# ==========================================
# 0. 網頁基本設定
# ==========================================
st.set_page_config(page_title="全校通用購書單系統", layout="centered", page_icon="📚")

st.title("📚 全校通用購書費通知單系統")
st.markdown("只要貼上 Google 試算表網址，系統自動完成對帳、計算資優生差額，並排版成 2x2 的 A4 通知單！")
st.divider()

# ==========================================
# 1. 核心邏輯區
# ==========================================
def get_google_sheet_xlsx_url(url):
    """將 Google Sheet 網址轉換為直接下載 XLSX 的連結，這樣才能一次讀取所有分頁"""
    try:
        pattern = r'https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9-_]+)'
        sheet_id = re.search(pattern, url).group(1)
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    except:
        return None

def should_buy_book(b_subj, b_code, s_gifted, s_eng, s_math, s_sci):
    """判斷該名學生是否需要購買該本書籍的通用邏輯"""
    s_gifted = str(s_gifted).strip()
    b_code = str(b_code).strip()
    
    # 規則 A：資優生免買邏輯
    if s_gifted == "語資" and b_subj in ["國", "英"]: return False
    if s_gifted == "數資" and b_subj in ["數", "自"]: return False
        
    # 規則 B：全班共同書目 (代號為 1 或 全)
    if b_code in ["1", "全"]: return True
        
    # 規則 C：分組書目比對
    if b_subj == "英" and b_code == str(s_eng).strip(): return True
    if b_subj == "數" and b_code == str(s_math).strip(): return True
    if b_subj == "自" and b_code == str(s_sci).strip(): return True
    
    return False

def generate_excel(df_students, df_books):
    """排版並生成 2x2 格式的 Excel 檔案"""
    wb = Workbook()
    ws = wb.active
    ws.title = "購書通知單(4張一頁)"

    # 版面與邊界設定
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5)

    # 欄寬設定
    ws.column_dimensions['A'].width = 8   ; ws.column_dimensions['B'].width = 38
    ws.column_dimensions['C'].width = 7   ; ws.column_dimensions['D'].width = 2
    ws.column_dimensions['E'].width = 8   ; ws.column_dimensions['F'].width = 38
    ws.column_dimensions['G'].width = 7

    # 字體與樣式設定
    f_title = Font(name="微軟正黑體", size=13, bold=True)
    f_info = Font(name="微軟正黑體", size=12, bold=True)
    f_norm = Font(name="微軟正黑體", size=9)
    f_bold = Font(name="微軟正黑體", size=10, bold=True)
    f_tot = Font(name="微軟正黑體", size=12, bold=True, color="FF0000")
    
    al_c = Alignment(horizontal="center", vertical="center")
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")
    
    thin = Side(style='thin', color='000000')
    b_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    RECEIPT_ROWS = 12

    for i, row in df_students.iterrows():
        seat = row.get("座號", "")
        name = row.get("姓名", "")
        gifted = row.get("資優類別", "")
        eng = row.get("英組", "1")
        math = row.get("數組", "1")
        sci = row.get("自組", "1")

        # 整理該學生的書單
        student_books = {"國": [], "英": [], "數": [], "自": [], "社會": [], "其他": []}
        for _, book in df_books.iterrows():
            b_name = book.get("商品名稱", "")
            b_subj = book.get("科目", "")
            b_code = book.get("分組代號", "1")
            b_price = book.get("單價", 0)
            
            # 將社會科系歸類到社會，以便顯示
            if b_subj in ["歷", "地", "公", "社會三科"]: b_subj = "社會"
            
            if should_buy_book(b_subj, b_code, gifted, eng, math, sci):
                if b_subj in student_books:
                    student_books[b_subj].append({"name": b_name, "price": b_price})
                else:
                    student_books["其他"].append({"name": b_name, "price": b_price})

        # 計算 Excel 寫入位置
        page = i // 4
        pos = i % 4
        start_row = page * (RECEIPT_ROWS * 2 + 2) + (0 if pos < 2 else RECEIPT_ROWS + 1) + 1
        start_col = 1 if pos % 2 == 0 else 5
        
        # 標題與基本資料
        ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col+2)
        c1 = ws.cell(row=start_row, column=start_col, value="學期購書費通知單")
        c1.font = f_title ; c1.alignment = al_c
        
        ws.merge_cells(start_row=start_row+1, start_column=start_col, end_row=start_row+1, end_column=start_col+2)
        c2 = ws.cell(row=start_row+1, column=start_col, value=f"座號：{seat}      姓名：{name}")
        c2.font = f_info ; c2.alignment = Alignment(horizontal="left", vertical="center")
        
        # 表頭
        headers = ["科目", "購買明細", "小計"]
        for col_offset, text in enumerate(headers):
            cell = ws.cell(row=start_row+2, column=start_col+col_offset, value=text)
            cell.font = f_bold ; cell.alignment = al_c ; cell.border = b_all

        # 內容寫入
        cat_order = ["國", "英", "數", "自", "社會", "其他"]
        curr_row = start_row + 3
        grand_total = 0
        
        for cat in cat_order:
            items = student_books[cat]
            if not items:
                det_str, subtotal = "免購", 0
            else:
                det_str = "、".join([f"{item['name']}(${int(item['price'])})" for item in items])
                subtotal = sum([item['price'] for item in items])
            grand_total += subtotal
            
            # 寫入分類、明細、小計
            c_cat = ws.cell(row=curr_row, column=start_col, value="社會三科" if cat == "社會" else cat)
            c_cat.font = f_bold ; c_cat.alignment = al_c ; c_cat.border = b_all
            
            c_det = ws.cell(row=curr_row, column=start_col+1, value=det_str)
            c_det.font = f_norm ; c_det.alignment = al_l ; c_det.border = b_all
            
            c_sub = ws.cell(row=curr_row, column=start_col+2, value=subtotal)
            c_sub.font = f_bold ; c_sub.alignment = al_c ; c_sub.border = b_all
            
            ws.row_dimensions[curr_row].height = 28
            curr_row += 1
            
        # 總計行
        ws.merge_cells(start_row=curr_row, start_column=start_col, end_row=curr_row, end_column=start_col+1)
        c_tot_l = ws.cell(row=curr_row, column=start_col, value="應收總計：")
        c_tot_l.font = f_tot ; c_tot_l.alignment = al_r ; c_tot_l.border = b_all
        ws.cell(row=curr_row, column=start_col+1).border = b_all
        
        c_tot_v = ws.cell(row=curr_row, column=start_col+2, value=f"{int(grand_total)}")
        c_tot_v.font = f_tot ; c_tot_v.alignment = al_c ; c_tot_v.border = b_all
        ws.row_dimensions[curr_row].height = 25
        
        # 繪製最外圍大框線 (方便裁切)
        for r in range(start_row, curr_row + 1):
            for c in range(start_col, start_col + 3):
                ws.cell(row=r, column=c).border = b_all

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# 2. 介面互動區
# ==========================================
st.subheader("Step 1: 輸入 Google 試算表網址")
st.info("⚠️ 請確保您的試算表已開啟「知道連結的人即可檢視」權限。\n\n"
        "分頁必須包含：第一頁「學生名單」(需有座號,姓名,資優類別,英組,數組,自組)\n"
        "第二頁「書局報價」(需有商品名稱,科目,分組代號,單價)")

sheet_url = st.text_input("請貼上試算表網址：", placeholder="https://docs.google.com/spreadsheets/d/...")

st.subheader("Step 2: 產生列印檔")

if sheet_url:
    xlsx_url = get_google_sheet_xlsx_url(sheet_url)
    if xlsx_url:
        try:
            with st.spinner("正在連線至 Google Sheets 讀取資料..."):
                # 直接讀取 XLSX 格式的兩個 Sheet
                df_students = pd.read_excel(xlsx_url, sheet_name=0).fillna("")
                df_books = pd.read_excel(xlsx_url, sheet_name=1).fillna("")
            
            st.success(f"✅ 讀取成功！共載入 {len(df_students)} 位學生、{len(df_books)} 筆書目報價。")
            
            if st.button("🚀 開始產生通知單 (A4 2x2排版)", type="primary"):
                with st.spinner("系統正在排版中..."):
                    excel_data = generate_excel(df_students, df_books)
                    
                st.balloons()
                st.download_button(
                    label="📥 點此下載 Excel 通知單列印檔",
                    data=excel_data,
                    file_name="全校購書通知單_列印排版.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        except Exception as e:
            st.error(f"❌ 讀取失敗！請確認網址是否正確，且權限已設為「知道連結的人即可檢視」。(錯誤細節: {e})")
    else:
        st.warning("⚠️ 無法解析網址，請確認您貼上的是完整的 Google Sheets 網址。")
