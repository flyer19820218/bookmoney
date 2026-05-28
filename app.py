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
# 🌟 分頁 2：全新升級 ─ ReportLab PDF 智慧交叉比對產生器
# =====================================================================
with tab2:
    st.subheader("🖨️ 學校名單 A/B 分組交叉比對系統 (產出 PDF)")
    st.markdown("直接上傳學校名單原始檔與書商 CSV，**自動校正人數**，產生一人一張的 A4 PDF 通知單。")

    st.markdown("#### 📁 第一步：上傳學校名單與書商報價")

    col_a, col_b = st.columns(2)
    with col_a:
        file_class = st.file_uploader("1. 班級總名單 (Excel 檔)", type=["xlsx", "xls"])
        file_books = st.file_uploader("2. 書商報價單 (CSV 或 Excel 檔)", type=["csv", "xlsx", "xls"])
    with col_b:
        file_eng   = st.file_uploader("3. 英文分組名單 (CSV 或 Excel 檔)", type=["csv", "xlsx", "xls"])
        file_math  = st.file_uploader("4. 數學分組名單 (CSV 或 Excel 檔)", type=["csv", "xlsx", "xls"])

    # 🌟 智慧欄位比對函數工具
    def find_column(df, keywords, default_name):
        for col in df.columns:
            if any(kw in str(col) for kw in keywords):
                return col
        return default_name

    def generate_reportlab_pdf(df_students, df_books_clean):
        """利用 ReportLab 產出完美 A4 中文購書單 PDF (一人一頁)"""
        pdf_io = io.BytesIO()
        c = canvas.Canvas(pdf_io, pagesize=A4)
        width, height = A4
        pdf_font = 'CustomFont' if HAS_FONT else 'Helvetica'
        
        for _, student in df_students.iterrows():
            seat = str(student.get("座號", "")).split('.')[0]
            name = str(student.get("姓名", "")).strip()
            s_eng = str(student.get("英組", "1")).strip()
            s_math = str(student.get("數組", "1")).strip()
            s_sci = str(student.get("自組", "1")).strip()
            s_gifted = str(student.get("資優類別", "")).strip()

            # 1. 建立校正後的個人書單
            personal_list = []
            total_amount = 0
            
            for _, b in df_books_clean.iterrows():
                b_name = b['name']
                b_price = b['price']
                b_code = b['code']
                b_subj = b['subj']
                
                # 判斷學生的分組是否符合購書條件
                is_match = False
                if b_code in ["1", "全", "", "nan"]:
                    is_match = True
                elif b_subj == "英" and s_eng in b_code:
                    is_match = True
                elif b_subj == "數" and s_math in b_code:
                    is_match = True
                elif b_subj == "自" and s_sci in b_code:
                    is_match = True
                
                # 資優生排除規則
                if s_gifted == "語資" and b_subj in ["國", "英"]: is_match = False
                if s_gifted == "數資" and b_subj in ["數", "自"]: is_match = False
                
                if is_match:
                    personal_list.append((b_name, b_price))
                    total_amount += b_price

            # 2. 開始繪製該學生的 PDF 頁面
            c.setFont(pdf_font, 22)
            c.drawCentredString(width/2, height - 60, "學 期 各 項 費 用 通 知 單")
            
            # 學生基本資訊橫條
            c.setFont(pdf_font, 12)
            c.drawString(60, height - 100, f"座號：{seat}        姓名：{name}")
            c.drawRightString(width - 60, height - 100, f"分組狀態：英({s_eng}) 數({s_math})")
            c.setStrokeColorRGB(0, 0, 0)
            c.setLineWidth(1.5)
            c.line(60, height - 110, width - 60, height - 110)
            
            # 畫出明細表頭
            c.setFont(pdf_font, 11)
            c.drawString(70, height - 135, "項目 / 書籍名稱")
            c.drawRightString(width - 70, height - 135, "金額 (元)")
            c.setLineWidth(0.5)
            c.line(60, height - 145, width - 60, height - 145)
            
            # 填入扣合分組後的實用明細
            y_pos = height - 170
            c.setFont(pdf_font, 10)
            if not personal_list:
                c.drawString(70, y_pos, "（本學期無特殊選購書籍，免繳費）")
                y_pos -= 25
            else:
                for b_name, price in personal_list:
                    if y_pos < 180: # 防止書籍項目太多溢出頁面
                        c.drawString(70, y_pos, "...項目過多未完...")
                        break
                    c.drawString(70, y_pos, b_name)
                    c.drawRightString(width - 70, y_pos, f"$ {int(price)}")
                    y_pos -= 22
            
            # 計算總計線與金額
            c.line(60, y_pos + 10, width - 60, y_pos + 10)
            c.setFont(pdf_font, 13)
            c.drawString(70, y_pos - 10, "應繳總計金額：")
            c.setFillColorRGB(0.8, 0, 0) # 紅色強調總金額
            c.drawRightString(width - 70, y_pos - 10, f"$ {int(total_amount)} 元")
            c.setFillColorRGB(0, 0, 0) # 恢復黑色
            
            # 家長回條簽章區塊 (鎖死在頁面底部，方便導師收單)
            box_y = 60
            c.setStrokeColorRGB(0.4, 0.4, 0.4)
            c.roundRect(60, box_y, width - 120, 75, 6, stroke=1, fill=0)
            c.setFont(pdf_font, 11)
            c.drawString(75, box_y + 50, "【家長簽章回條】")
            c.setFont(pdf_font, 10)
            c.drawString(75, box_y + 25, f"本人已確認上述 座號 {seat} {name} 之購書明細與金額無誤。")
            c.drawString(width - 200, box_y + 25, "家長簽名：__________________")
            
            # 完成此學生，換下一頁
            c.showPage()
            
        c.save()
        pdf_io.seek(0)
        return pdf_io

    st.divider()
    st.markdown("#### 🚀 第二步：執行交叉智慧扣合與產出")

    # 檢查必要檔案
    if file_class and file_books:
        st.success("✅ 基礎名單與書商報價單已偵測到！")
        
        if st.button("🎯 執行全班 AB 分組精準對帳並產生 PDF", type="primary"):
            with st.spinner("正在以學校分組資料覆蓋書商錯誤數量，完美排版中..."):
                try:
                    # 1. 讀取班級總表
                    df_s = pd.read_excel(file_class).fillna("")
                    
                    # 2. 智慧讀取書商報價單
                    if file_books.name.endswith('.csv'):
                        df_b = pd.read_csv(file_books).fillna("")
                    else:
                        df_b = pd.read_excel(file_books).fillna("")
                        
                    # 3. 如果有上傳獨立的英文/數學組別 CSV，執行動態合併校正
                    if file_eng:
                        df_e = pd.read_csv(file_eng) if file_eng.name.endswith('.csv') else pd.read_excel(file_eng)
                        # 比對座號或姓名將組別合進主表
                        c_seat = find_column(df_e, ["座號", "號碼"], "座號")
                        c_group = find_column(df_e, ["組", "英文"], "英組")
                        df_e_clean = df_e[[c_seat, c_group]].rename(columns={c_seat: "座號", c_group: "英組"})
                        df_s = df_s.drop(columns=["英組"], errors="ignore").merge(df_e_clean, on="座號", how="left")
                        
                    if file_math:
                        df_m = pd.read_csv(file_math) if file_math.name.endswith('.csv') else pd.read_excel(file_math)
                        c_seat = find_column(df_m, ["座號", "號碼"], "座號")
                        c_group = find_column(df_m, ["組", "數學"], "數組")
                        df_m_clean = df_m[[c_seat, c_group]].rename(columns={c_seat: "座號", c_group: "數組"})
                        df_s = df_s.drop(columns=["數組"], errors="ignore").merge(df_m_clean, on="座號", how="left")

                    # 4. 統一書商名單的欄位標籤
                    col_name = find_column(df_b, ["品名", "商品", "名稱", "書籍"], "商品名稱")
                    col_price = find_column(df_b, ["單價", "價格", "金額"], "單價")
                    col_code = find_column(df_b, ["附記", "備註", "分組", "代號"], "分組代號")
                    col_subj = find_column(df_b, ["科目", "類別"], "科目")
                    
                    df_books_clean = pd.DataFrame({
                        'name': df_b[col_name],
                        'price': pd.to_numeric(df_b[col_price], errors='coerce').fillna(0),
                        'code': df_b[col_code].astype(str),
                        'subj': df_b[col_subj].astype(str)
                    })

                    # 自動判斷科目簡寫 (例如歷史/地理/公民自動歸入社會，方便匹配)
                    df_books_clean['subj'] = df_books_clean['subj'].apply(lambda x: "社會" if x in ["歷", "地", "公", "歷史", "地理", "公民"] else x)

                    # 5. 生成完美的 ReportLab PDF
                    pdf_result = generate_reportlab_pdf(df_s, df_books_clean)
                    
                    st.balloons()
                    st.success("🎉 交叉比對校正成功！已自動用學校名單人數剔除書商出貨誤差。")
                    
                    st.download_button(
                        label="📥 下載全班 A4 PDF 繳費通知單 (一人一張完美列印版)",
                        data=pdf_result,
                        file_name="全班購書單_校正完美版.pdf",
                        mime="application/pdf"
                    )
                except Exception as e:
                    st.error(f"❌ 比對程序發生異常：{e}")
                    st.warning("提示：請確認您的名單或 CSV 檔案內容是否含有『座號』與『姓名』欄位標題。")
    else:
        st.info("💡 請至少先上傳『1. 班級總名單』與『2. 書商報價單』，系統的智慧對帳按鈕就會解鎖出現喔！")
