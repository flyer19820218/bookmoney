import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.worksheet.page import PageMargins
import re
import io
import requests
import os

# 🌟 新增：ReportLab PDF 核心繪圖套件
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader

# ==========================================
# 0. 網頁基本設定與自動字型下載
# ==========================================
st.set_page_config(page_title="全校通用購書單系統", layout="centered", page_icon="📚")

FONT_NAME = "NotoSansTC-Regular.ttf"
FONT_URL = "https://cdn.jsdelivr.net/gh/themoeway/noto-sans-tc-ttf@master/ttf/NotoSansTC-Regular.ttf"

@st.cache_resource
def init_fonts():
    """確保 Streamlit Cloud 上有中文字型可用，避免 PDF 亂碼"""
    if not os.path.exists(FONT_NAME):
        try:
            with st.spinner("正在下載微軟正黑體/Noto中文字型以支援 PDF 產出..."):
                response = requests.get(FONT_URL, headers={'User-Agent': 'Mozilla/5.0'}, timeout=60)
                with open(FONT_NAME, "wb") as f:
                    f.write(response.content)
        except Exception as e:
            st.error(f"字型下載失敗：{e}")
            return False
    try:
        pdfmetrics.registerFont(TTFont('CustomFont', FONT_NAME))
        return True
    except Exception as e:
        st.error(f"字型註冊失敗：{e}")
        return False

HAS_FONT = init_fonts()

st.title("📚 班級各項費用與通知單系統")
st.markdown("請選擇上方分頁切換您要使用的功能！")

# 🌟 建立兩個分頁
tab1, tab2 = st.tabs(["📝 雲端試算表自動版 (現有功能)", "🖨️ 懶人 PDF 產生器 (新功能測試)"])

# =====================================================================
# 🌟 分頁 1：原本的雲端自動系統 (保持不變)
# =====================================================================
with tab1:
    st.header("💡 導師專屬「一鍵排版」全攻略")

    with st.expander("👉 第一步：領取並填寫您的專屬表格 (點我展開)", expanded=True):
        st.markdown("""
        1. 點擊進入 [校內購書公版範本 (短網址：https://reurl.cc/K2LgNe)](https://reurl.cc/K2LgNe)。
        2. 進去後，點選左上角的 **「檔案」 > 「建立副本」**。（一定要建立副本才能編輯喔！）
        3. **不想自己打字？讓 AI 幫你精準整理！**
           手邊只有「紙本估價單」？直接拍照傳給 **Gemini 或 ChatGPT**，並**完整複製貼上以下這段指令給 AI**：
           
           > 你現在是一位專業、細心的「學校資料輸入員」。請將我上傳的書商估價單照片，精準地轉換成表格。
           > 
           > **【表格必須嚴格包含這 4 個欄位】：**
           > 1. **商品名稱**：完整保留書名或材料費名稱。
           > 2. **科目**：請根據書名自動分類為「國、英、數、自、社會、其他」這六類。（注意：歷史、地理、公民請一律歸類為「社會」；聯絡簿、桌墊、材料費請歸類為「其他」）。
           > 3. **分組代號**：若圖片中無特別標示分組，請一律填入數字「1」。如果有分組，請直接填寫代號（如 5A、6B）。
           > 4. **單價**：只填寫純數字，不要加上 $ 或 元。
           > 
           > **【嚴格限制】：** 請勿遺漏項目，不要加入任何問候語或廢話，你的回覆只能有「一個表格」。
           
           接著直接把 AI 做好的表格**全選複製、貼上**到您的 Google 試算表即可！
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
        2. 畫面顯示「讀取成功」後，按下 **「🚀 開始產生檔案」**。
        3. 您可以分別下載「給家長的 4格通知單」以及「給導師的 對帳收費總表」。
        """)

    st.warning("⚠️ **列印小叮嚀：** 下載家長通知單後，因各電腦 Excel 預設邊界不同，請在列印時自行點選「預覽列印」微調縮放比例！")
    st.info("💡 **資優生怎麼辦？** 在名單分頁的資優生欄位選「語資」或「數資」，系統會自動扣除。一年級沒分組，代號全填 `1` 就好！")
    st.divider()

    # 原本的核心邏輯函數
    def get_google_sheet_xlsx_url(url):
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
        s_gifted = str(s_gifted).strip()
        b_code = str(b_code).strip()
        
        if s_gifted == "語資" and b_subj in ["國", "英"]: return False
        if s_gifted == "數資" and b_subj in ["數", "自"]: return False
        if b_code in ["1", "全", "", "nan"]: return True
            
        if b_subj == "英" and str(s_eng).strip() in b_code: return True
        if b_subj == "數" and str(s_math).strip() in b_code: return True
        if b_subj == "自" and str(s_sci).strip() in b_code: return True
        return False

    def generate_receipts_excel(df_students, df_books):
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
            seat, name = row.get("座號", ""), row.get("姓名", "")
            gifted = row.get("資優類別", "")
            eng, math, sci = row.get("英組", "1"), row.get("數組", "1"), row.get("自組", "1")

            student_books = {"國": [], "英": [], "數": [], "自": [], "社會": [], "其他": []}
            for _, book in df_books.iterrows():
                b_name, b_subj, b_code, b_price = book.get("商品名稱", ""), book.get("科目", ""), book.get("分組代號", "1"), book.get("單價", 0)
                if b_subj in ["歷", "地", "公", "社會三科"]: b_subj = "社會"
                if should_buy_book(b_subj, b_code, gifted, eng, math, sci):
                    if b_subj in student_books: student_books[b_subj].append({"name": b_name, "price": b_price})
                    else: student_books["其他"].append({"name": b_name, "price": b_price})

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
            
            for col_offset, text in enumerate(["科目", "購買明細", "小計"]):
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
                for c in range(start_col, start_col + 3): ws.cell(row=r, column=c).border = b_all

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    def generate_master_excel(df_students, df_books):
        wb = Workbook()
        ws = wb.active
        ws.title = "班級收費總表"
        
        book_rows = []
        for _, book in df_books.iterrows():
            b_name, b_subj, b_code, b_price = book.get("商品名稱", ""), book.get("科目", ""), book.get("分組代號", "1"), book.get("單價", 0)
            b_subj_check = "社會" if b_subj in ["歷", "地", "公", "社會三科"] else b_subj
            
            qty = 0
            for _, s in df_students.iterrows():
                if should_buy_book(b_subj_check, b_code, s.get("資優類別", ""), s.get("英組", "1"), s.get("數組", "1"), s.get("自組", "1")):
                    qty += 1
            book_rows.append([b_name, b_subj, b_code, qty, b_price])

        student_rows = []
        for _, s in df_students.iterrows():
            seat, name = s.get("座號", ""), s.get("姓名", "")
            gifted = s.get("資優類別", "")
            eng, math, sci = s.get("英組", "1"), s.get("數組", "1"), s.get("自組", "1")
            
            subtotal = 0
            for _, book in df_books.iterrows():
                b_subj, b_code, b_price = book.get("科目", ""), book.get("分組代號", "1"), book.get("單價", 0)
                b_subj_check = "社會" if b_subj in ["歷", "地", "公", "社會三科"] else b_subj
                if should_buy_book(b_subj_check, b_code, gifted, eng, math, sci):
                    subtotal += b_price
            student_rows.append([seat, name, gifted, eng, math, sci, subtotal])

        headers_left = ["商品名稱", "科目", "分組代號", "購買數量", "單價"]
        for col_idx, h in enumerate(headers_left, 1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")

        headers_right = ["座號", "姓名", "資優", "英組", "數組", "自組", "應收總額"]
        for col_idx, h in enumerate(headers_right, 7):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")

        thin_border = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                             top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

        for r_idx, b_row in enumerate(book_rows, 2):
            for c_idx, val in enumerate(b_row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = thin_border
                if c_idx > 1: cell.alignment = Alignment(horizontal="center")

        for r_idx, s_row in enumerate(student_rows, 2):
            for c_idx, val in enumerate(s_row, 7):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = thin_border
                if c_idx != 8: cell.alignment = Alignment(horizontal="center")
                if c_idx == 13: cell.font = Font(bold=True)

        last_row = max(len(book_rows), len(student_rows)) + 2
        ws.cell(row=last_row, column=12, value="班級總計").font = Font(bold=True)
        total_sum = sum([s[-1] for s in student_rows])
        tot_cell = ws.cell(row=last_row, column=13, value=total_sum)
        tot_cell.font = Font(bold=True, color="FF0000"); tot_cell.border = thin_border

        ws.column_dimensions['A'].width = 35; ws.column_dimensions['B'].width = 8
        ws.column_dimensions['C'].width = 10; ws.column_dimensions['D'].width = 10; ws.column_dimensions['E'].width = 8
        ws.column_dimensions['F'].width = 3  
        ws.column_dimensions['G'].width = 6 ; ws.column_dimensions['H'].width = 12
        ws.column_dimensions['I'].width = 8 ; ws.column_dimensions['J'].width = 8
        ws.column_dimensions['K'].width = 8 ; ws.column_dimensions['L'].width = 8
        ws.column_dimensions['M'].width = 12

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

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
                
                if st.button("🚀 開始產生檔案", type="primary"):
                    with st.spinner("系統正在計算與排版中..."):
                        receipts_data = generate_receipts_excel(df_students, df_books)
                        master_data = generate_master_excel(df_students, df_books)
                        
                    st.balloons()
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(label="📥 下載【家長通知單】(A4列印版)", data=receipts_data, file_name="家長購書通知單.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    with col2:
                        st.download_button(label="📥 下載【導師對帳總表】(收費明細)", data=master_data, file_name="導師收費對帳總表.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error("❌ 讀取失敗！請確認網址與權限。")


# =====================================================================
# 🌟 分頁 2：全新升級 ─ 智慧交叉比對系統 (先測試產出 Excel 版)
# =====================================================================
with tab2:
    st.subheader("🖨️ 學校名單 A/B 分組交叉比對系統 (產出自動化 Excel)")
    st.markdown("直接上傳學校名單原始檔與書商 CSV，**自動校正人數**，產出一人 4 格的通知單與對帳總表。")

    st.markdown("#### 📁 第一步：上傳學校名單與書商報價")

    col_a, col_b = st.columns(2)
    with col_a:
        file_class = st.file_uploader("1. 班級總名單 (Excel 檔)", type=["xlsx", "xls"])
        file_books = st.file_uploader("2. 書商報價單 (CSV 或 Excel 檔)", type=["csv", "xlsx", "xls"])
    with col_b:
        file_eng   = st.file_uploader("3. 英文分組名單 (可選)", type=["csv", "xlsx", "xls"])
        file_math  = st.file_uploader("4. 數學分組名單 (可選)", type=["csv", "xlsx", "xls"])

    # 🌟 內部工具：智慧欄位比對
    def find_column(df, keywords, default_name):
        for col in df.columns:
            if any(kw in str(col) for kw in keywords):
                return col
        return None
        
    # 🌟 內部工具：自動通靈猜測科目 (解決書商沒給科目的問題)
    def guess_subject(name):
        name = str(name).replace("國中", "").replace("國小", "")
        if any(k in name for k in ["英"]): return "英"
        if any(k in name for k in ["數"]): return "數"
        if any(k in name for k in ["自", "理化", "生物", "地科"]): return "自"
        if any(k in name for k in ["歷", "地", "公", "社會"]): return "社會"
        if any(k in name for k in ["國", "文"]): return "國"
        return "其他"

    # 🌟 內部工具：智慧解析「並排式」分組名單 (破解合併儲存格)
    def parse_horizontal_group_file(uploaded_file):
        if uploaded_file.name.endswith('.csv'):
            df_grp = pd.read_csv(uploaded_file, header=None).fillna("")
        else:
            df_grp = pd.read_excel(uploaded_file, header=None).fillna("")
            
        header_idx = -1
        for idx, row in df_grp.iterrows():
            if any(isinstance(v, str) and "姓名" in v for v in row.values):
                header_idx = idx ; break
                
        mapping = {}
        if header_idx != -1:
            col_to_group = {}
            current_group = "1"
            for col_idx in range(df_grp.shape[1]):
                for r in range(0, header_idx):
                    val = str(df_grp.iloc[r, col_idx]).strip()
                    if val and val != "nan":
                        match = re.search(r'[A-Za-z]', val)
                        if match: current_group = match.group(0).upper()
                        elif len(val) <= 2: current_group = val
                col_to_group[col_idx] = current_group

            for col_idx in range(df_grp.shape[1]):
                if "姓名" in str(df_grp.iloc[header_idx, col_idx]):
                    group_for_this_col = col_to_group[col_idx]
                    names = df_grp.iloc[header_idx+1:, col_idx].astype(str).str.strip()
                    for name in names:
                        clean_name = name.replace(" ", "")
                        if clean_name and clean_name != "nan":
                            mapping[clean_name] = group_for_this_col
        return mapping

    # 🌟 Tab2 專用 Excel 產生器：家長 4 格通知單
    def generate_excel_receipts_tab2(df_students, df_books):
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
            seat = str(row.get("座號", "")).split('.')[0]
            name = str(row.get("姓名", "")).strip()
            gifted = str(row.get("資優類別", "")).strip()
            eng, math, sci = str(row.get("英組", "1")), str(row.get("數組", "1")), str(row.get("自組", "1"))

            student_books = {"國": [], "英": [], "數": [], "自": [], "社會": [], "其他": []}
            for _, book in df_books.iterrows():
                b_name, b_subj, b_code, b_price = book['name'], book['subj'], book['code'], book['price']
                
                is_match = False
                if b_code in ["1", "全", "", "nan"]: is_match = True
                elif b_subj == "英" and eng in b_code: is_match = True
                elif b_subj == "數" and math in b_code: is_match = True
                elif b_subj == "自" and sci in b_code: is_match = True
                
                if gifted == "語資" and b_subj in ["國", "英"]: is_match = False
                if gifted == "數資" and b_subj in ["數", "自"]: is_match = False
                if "語資" in gifted and "數資" in gifted and b_subj in ["國", "英", "數", "自"]: is_match = False
                
                if is_match:
                    if b_subj in student_books: student_books[b_subj].append({"name": b_name, "price": b_price})
                    else: student_books["其他"].append({"name": b_name, "price": b_price})

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
            
            for col_offset, text in enumerate(["科目", "購買明細", "小計"]):
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
                
                c_cat = ws.cell(row=curr_row, column=start_col, value=cat)
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
                for c in range(start_col, start_col + 3): ws.cell(row=r, column=c).border = b_all

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    # 🌟 Tab2 專用 Excel 產生器：導師收費總表
    def generate_excel_master_tab2(df_students, df_books):
        wb = Workbook()
        ws = wb.active
        ws.title = "班級收費總表"
        
        book_rows = []
        for _, b in df_books.iterrows():
            qty = 0
            for _, s in df_students.iterrows():
                is_match = False
                if b['code'] in ["1", "全", "", "nan"]: is_match = True
                elif b['subj'] == "英" and str(s.get("英組","1")) in b['code']: is_match = True
                elif b['subj'] == "數" and str(s.get("數組","1")) in b['code']: is_match = True
                elif b['subj'] == "自" and str(s.get("自組","1")) in b['code']: is_match = True
                
                gifted = str(s.get("資優類別", ""))
                if gifted == "語資" and b['subj'] in ["國", "英"]: is_match = False
                if gifted == "數資" and b['subj'] in ["數", "自"]: is_match = False
                if "語資" in gifted and "數資" in gifted and b['subj'] in ["國", "英", "數", "自"]: is_match = False
                
                if is_match: qty += 1
            book_rows.append([b['name'], b['subj'], b['code'], qty, b['price']])

        student_rows = []
        for _, s in df_students.iterrows():
            seat, name = str(s.get("座號", "")).split('.')[0], s.get("姓名", "")
            gifted, eng, math, sci = s.get("資優類別", ""), s.get("英組", "1"), s.get("數組", "1"), s.get("自組", "1")
            
            subtotal = 0
            for _, b in df_books.iterrows():
                is_match = False
                if b['code'] in ["1", "全", "", "nan"]: is_match = True
                elif b['subj'] == "英" and str(eng) in b['code']: is_match = True
                elif b['subj'] == "數" and str(math) in b['code']: is_match = True
                elif b['subj'] == "自" and str(sci) in b['code']: is_match = True
                
                if gifted == "語資" and b['subj'] in ["國", "英"]: is_match = False
                if gifted == "數資" and b['subj'] in ["數", "自"]: is_match = False
                if "語資" in gifted and "數資" in gifted and b['subj'] in ["國", "英", "數", "自"]: is_match = False
                
                if is_match: subtotal += b['price']
            student_rows.append([seat, name, gifted, eng, math, sci, subtotal])

        headers_left = ["商品名稱", "科目", "分組代號", "購買數量", "單價"]
        for col_idx, h in enumerate(headers_left, 1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")

        headers_right = ["座號", "姓名", "資優", "英組", "數組", "自組", "應收總額"]
        for col_idx, h in enumerate(headers_right, 7):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")

        thin_border = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                             top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

        for r_idx, b_row in enumerate(book_rows, 2):
            for c_idx, val in enumerate(b_row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = thin_border
                if c_idx > 1: cell.alignment = Alignment(horizontal="center")

        for r_idx, s_row in enumerate(student_rows, 2):
            for c_idx, val in enumerate(s_row, 7):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = thin_border
                if c_idx != 8: cell.alignment = Alignment(horizontal="center")
                if c_idx == 13: cell.font = Font(bold=True)

        last_row = max(len(book_rows), len(student_rows)) + 2
        ws.cell(row=last_row, column=12, value="班級總計").font = Font(bold=True)
        total_sum = sum([s[-1] for s in student_rows])
        tot_cell = ws.cell(row=last_row, column=13, value=total_sum)
        tot_cell.font = Font(bold=True, color="FF0000"); tot_cell.border = thin_border

        ws.column_dimensions['A'].width = 35; ws.column_dimensions['B'].width = 8
        ws.column_dimensions['C'].width = 10; ws.column_dimensions['D'].width = 10; ws.column_dimensions['E'].width = 8
        ws.column_dimensions['F'].width = 3  
        ws.column_dimensions['G'].width = 6 ; ws.column_dimensions['H'].width = 12
        ws.column_dimensions['I'].width = 8 ; ws.column_dimensions['J'].width = 8
        ws.column_dimensions['K'].width = 8 ; ws.column_dimensions['L'].width = 8
        ws.column_dimensions['M'].width = 12

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    # 🌟 執行區塊
    if file_class and file_books:
        try:
            # 1. 智慧讀取原班名單
            df_temp = pd.read_excel(file_class, header=None).fillna("")
            header_idx = 0
            for idx, row in df_temp.iterrows():
                if any("姓名" in str(v) for v in row.values):
                    header_idx = idx ; break
            df_s = pd.read_excel(file_class, skiprows=header_idx).fillna("")
            
            c_seat = find_column(df_s, ["座號", "號碼", "序號"], "座號")
            c_name = find_column(df_s, ["姓名", "名稱", "學生"], "姓名")
            if c_seat and c_seat != "座號": df_s = df_s.rename(columns={c_seat: "座號"})
            if c_name and c_name != "姓名": df_s = df_s.rename(columns={c_name: "姓名"})
            
            for col in ["英組", "數組", "自組", "資優類別"]:
                if col not in df_s.columns: df_s[col] = "1"

            # 🌟 智慧整合分組名單 ＆ 資優生判定邏輯
            eng_map = parse_horizontal_group_file(file_eng) if file_eng else {}
            math_map = parse_horizontal_group_file(file_math) if file_math else {}

            for idx, row in df_s.iterrows():
                clean_name = str(row["姓名"]).replace(" ", "").strip()
                if file_eng:
                    if clean_name in eng_map: df_s.at[idx, "英組"] = eng_map[clean_name]
                    else:
                        df_s.at[idx, "資優類別"] = "語資"
                        df_s.at[idx, "英組"] = "免"
                
                if file_math:
                    if clean_name in math_map: df_s.at[idx, "數組"] = math_map[clean_name]
                    else:
                        current_gifted = str(df_s.at[idx, "資優類別"])
                        if current_gifted == "1" or current_gifted == "": df_s.at[idx, "資優類別"] = "數資"
                        elif current_gifted == "語資": df_s.at[idx, "資優類別"] = "語資/數資"
                        df_s.at[idx, "數組"] = "免"

            with st.expander("👀 步驟 1.5：核對學生名條與資優生判定 (點我展開)", expanded=True):
                st.write("請確認「分組」是否成功，若找不到分組，系統會自動將該生標示為「語資/數資」：")
                st.dataframe(df_s[["座號", "姓名", "英組", "數組", "資優類別"]])

            # 2. 智慧讀取書商報價單
            header_skip = 0
            keywords_header = ["品名", "商品", "名稱", "書籍", "單價", "價格", "金額", "序號"]
            if file_books.name.endswith('.csv'):
                bytes_data = file_books.read()
                lines = []
                used_enc = 'utf-8'
                for enc in ['utf-8', 'big5', 'cp950', 'utf-8-sig']:
                    try:
                        lines = bytes_data.decode(enc).splitlines()
                        used_enc = enc ; break
                    except: continue
                for idx, line in enumerate(lines):
                    if any(kw in line for kw in keywords_header):
                        header_skip = idx ; break
                file_books.seek(0)
                try: df_b = pd.read_csv(file_books, skiprows=header_skip, encoding=used_enc).fillna("")
                except: 
                    file_books.seek(0)
                    df_b = pd.read_csv(file_books, skiprows=header_skip, on_bad_lines='skip').fillna("")
            else:
                df_temp2 = pd.read_excel(file_books, header=None).fillna("")
                for idx, row in df_temp2.iterrows():
                    row_str = "".join([str(val) for val in row.values])
                    if any(kw in row_str for kw in keywords_header):
                        header_skip = idx ; break
                file_books.seek(0)
                df_b = pd.read_excel(file_books, skiprows=header_skip).fillna("")

            b_col_name = find_column(df_b, ["品名", "商品", "名稱", "書籍"], "商品名稱")
            b_col_price = find_column(df_b, ["單價", "價格", "金額"], "單價")
            b_col_code = find_column(df_b, ["附記", "備註", "分組", "代號"], "分組代號")
            
            # 🌟 啟動通靈術：如果書商沒給「科目」欄位，AI 自己猜！
            b_col_subj = find_column(df_b, ["科目", "類別"], "科目")
            if b_col_subj:
                subj_series = df_b[b_col_subj].astype(str)
            else:
                # 找不到科目欄位時，針對「品名」直接呼叫猜測函數
                subj_series = df_b[b_col_name].apply(guess_subject)

            df_books_clean = pd.DataFrame({
                'name': df_b[b_col_name],
                'price': pd.to_numeric(df_b[b_col_price], errors='coerce').fillna(0),
                'code': df_b[b_col_code].astype(str),
                'subj': subj_series
            })
            df_books_clean['subj'] = df_books_clean['subj'].apply(lambda x: "社會" if x in ["歷", "地", "公", "歷史", "地理", "公民"] else x)

            with st.expander("👀 步驟 1.8：核對書商書籍清單 (點我展開)"):
                st.write("這是系統抓到的完整書目，請確認 AI 通靈猜測的「科目」是否正確：")
                st.dataframe(df_books_clean)

            st.divider()
            st.markdown("#### 🚀 第二步：執行交叉智慧扣合與產出 Excel")

            if st.button("🎯 確認無誤，開始產出 Excel 繳費單與總表", type="primary"):
                with st.spinner("正在進行交叉對帳與排版中..."):
                    receipts_data = generate_excel_receipts_tab2(df_s, df_books_clean)
                    master_data = generate_excel_master_tab2(df_s, df_books_clean)
                    
                    st.balloons()
                    st.success("🎉 對帳完成！已為您準備好熟悉的 Excel 檔案格式。")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(label="📥 下載【家長通知單】(A4 4格版)", data=receipts_data, file_name="全班通知單_核對版.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    with col2:
                        st.download_button(label="📥 下載【導師對帳總表】(收費明細)", data=master_data, file_name="導師總表_核對版.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        except Exception as e:
            st.error(f"系統讀取發生錯誤：{e}")
            st.info("提示：請檢查上傳的檔案格式是否正確。")
    else:
        st.info("💡 請先上傳名單與報價單，核對預覽畫面就會出現喔！")
        st.info("💡 請先上傳名單與報價單，核對預覽畫面就會出現喔！")
