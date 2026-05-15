import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.worksheet.page import PageMargins
import re
import io
import requests

# ==========================================
# 0. 網頁基本設定
# ==========================================
st.set_page_config(page_title="全校通用購書單系統", layout="centered", page_icon="📚")

st.title("📚 全校通用購書費通知單系統")
st.markdown("自動比對名單、計算資優生差額，一鍵產出 **2x2 的 A4 裁切版通知單**！")
st.divider()

# ==========================================
# 1. 老師無腦使用說明區 
# ==========================================
st.header("💡 導師專屬「一鍵排版」全攻略")

with st.expander("👉 第一步：領取並填寫您的專屬表格 (點我展開)", expanded=True):
    st.markdown("""
    1. 點擊進入 [校內購書公版範本 (短網址：https://reurl.cc/K2LgNe)](https://reurl.cc/K2LgNe)。
    2. 進去後，點選左上角的 **「檔案」 > 「建立副本」**。（一定要建立副本才能編輯喔！）
    3. **不知道怎麼打字？請 AI 幫忙！**
       手邊只有「紙本估價單」或「分組名單」？直接拍照傳給 **Gemini 或 ChatGPT**，並對它說：
       > *"這是我班上的購書估價單，請幫我整理成「商品名稱、科目、分組代號、單價」四個欄位的表格。"*
       接著直接把 AI 做好的表格**複製貼上**到您的 Google 試算表即可！
    """)

with st.expander("👉 第二步：開啟「共用」權限 (這步沒做會失敗喔！)", expanded=True):
    st.markdown("""
    1. 點擊您 Google 試算表右上角大大的 **「共用」** 按鈕。
    2. 在「一般存取權」下方，將「限制」改為 **「知道連結的人即可檢視」**。
    3. 按下 **「複製連結」**。
    """)

with st.expander("👉 第三步：貼上網址，領取列印檔", expanded=True):
    st.markdown("""
    1. 將剛剛複製的網址，**貼到下方輸入框**。
    2. 畫面顯示「讀取成功」後，按下 **「🚀 開始產生通知單」**。
    3. 下載產出的 Excel 檔準備列印。
    """)

# 新增的 99 分真實宣告
st.warning("⚠️ **列印小叮嚀：** 系統排版能做到 99 分，但因為 **Mac 和 PC 的 Excel 預設邊界不一樣**，下載後可能不會「剛好」完美塞滿一張 A4。**請列印時自行點選「預覽列印」，並微調縮放比例或邊界喔！**")

st.info("💡 **資優生怎麼辦？** \n 在名單分頁的資優生欄位選「語資」或「數資」，系統會自動幫他扣掉不用買的講義費。一年級沒分組的話，分組代號通通填 `1` 就好！")
st.divider()

# ==========================================
# 2. 核心邏輯區
# ==========================================
def get_google_sheet_xlsx_url(url):
    """將 Google Sheet 網址(含短網址)轉換為直接下載 XLSX 的連結"""
    try:
        if "docs.google.com" not in url:
            response = requests.head(url, allow_redirects=True)
            url = response.url
            
        pattern = r'https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9-_]+)'
        match = re.search(pattern, url)
        if match:
            sheet_id = match.group(1)
            return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
        return None
    except Exception as e:
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

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5)

    ws.column_dimensions['A'].width = 8   ; ws.column_dimensions['B'].width = 38
    ws.column_dimensions['C'].width = 7   ; ws.column_dimensions['D'].width = 2
    ws.column_dimensions['E'].width = 8   ; ws.column_dimensions['F'].width = 38
    ws.column_dimensions['G'].width = 7

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

        student_books = {"國": [], "英": [], "數": [], "自": [], "社會": [], "其他": []}
        for _, book in df_books.iterrows():
            b_name = book.get("商品名稱", "")
            b_subj = book.get("科目", "")
            b_code = book.get("分組代號", "1")
            b_price = book.get("單價", 0)
            
            if b_subj in ["歷", "地", "公", "社會三科"]: b_subj = "社會"
            
            if should_buy_book(b_subj, b_code, gifted, eng, math, sci):
                if b_subj in student_books:
                    student_books[b_subj].append({"name": b_name, "price": b_price})
                else:
                    student_books["其他"].append({"name": b_name, "price": b_price})

        page = i // 4
        pos = i % 4
        start_row = page * (RECEIPT_ROWS * 2 + 2) + (0 if pos < 2 else RECEIPT_ROWS + 1) + 1
        start_col = 1 if pos % 2 == 0 else 5
        
        ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col+2)
        c1 = ws.cell(row=start_row, column=start_col, value="學期購書費通知單")
        c1.font = f_title ; c1.alignment = al_c
        
        ws.merge_cells(start_row=start_row+1, start_column=start_col, end_row=start_row+1, end_column=start_col+2)
        c2 = ws.cell(row=start_row+1, column=start_col, value=f"座號：{seat}      姓名：{name}")
        c2.font = f_info ; c2.alignment = Alignment(horizontal="left", vertical="center")
        
        headers = ["科目", "購買明細", "小計"]
        for col_offset, text in enumerate(headers):
            cell = ws.cell(row=start_row+2, column=start_col+col_offset, value=text)
            cell.font = f_bold ; cell.alignment = al_c ; cell.border = b_all

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
            
            c_cat = ws.cell(row=curr_row, column=start_col, value="社會三科" if cat == "社會" else cat)
            c_cat.font = f_bold ; c_cat.alignment = al_c ; c_cat.border = b_all
            
            c_det = ws.cell(row=curr_row, column=start_col+1, value=det_str)
            c_det.font = f_norm ; c_det.alignment = al_l ; c_det.border = b_all
            
            c_sub = ws.cell(row=curr_row, column=start_col+2, value=subtotal)
            c_sub.font = f_bold ; c_sub.alignment = al_c ; c_sub.border = b_all
            
            ws.row_dimensions[curr_row].height = 28
            curr_row += 1
            
        ws.merge_cells(start_row=curr_row, start_column=start_col, end_row=curr_row, end_column=start_col+1)
        c_tot_l = ws.cell(row=curr_row, column=start_col, value="應收總計：")
        c_tot_l.font = f_tot ; c_tot_l.alignment = al_r ; c_tot_l.border = b_all
        ws.cell(row=curr_row, column=start_col+1).border = b_all
        
        c_tot_v = ws.cell(row=curr_row, column=start_col+2, value=f"{int(grand_total)}")
        c_tot_v.font = f_tot ; c_tot_v.alignment = al_c ; c_tot_v.border = b_all
        ws.row_dimensions[curr_row].height = 25
        
        for r in range(start_row, curr_row + 1):
            for c in range(start_col, start_col + 3):
                ws.cell(row=r, column=c).border = b_all

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# 3. 介面操作區
# ==========================================
st.subheader("🛠️ 開始作業")

default_url = "https://reurl.cc/K2LgNe"
sheet_url = st.text_input("👇 請在下方輸入您的試算表網址：", value=default_url)

if sheet_url:
    xlsx_url = get_google_sheet_xlsx_url(sheet_url)
    if xlsx_url:
        try:
            with st.spinner("正在連線至 Google Sheets 讀取資料..."):
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
                    file_name="班級購書通知單_排版完成.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                # 下載按鈕下方再提醒一次
                st.caption("☝️ 下載後請記得先「預覽列印」調整邊界與縮放比例喔！")
                
        except Exception as e:
            st.error("❌ 讀取失敗！請確認網址是否正確，且【共用】權限已設為「知道連結的人即可檢視」。")
            st.warning(f"系統錯誤代碼：{e}")
    else:
        st.warning("⚠️ 無法解析網址，請確認您貼上的是正確的網址。")
