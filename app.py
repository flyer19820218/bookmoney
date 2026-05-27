import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.worksheet.page import PageMargins
import re
import io
import requests
from fpdf import FPDF  # 新增：用來產生 PDF 的套件

# ==========================================
# 0. 網頁基本設定
# ==========================================
st.set_page_config(page_title="全校通用購書單系統", layout="centered", page_icon="📚")

st.title("📚 班級各項費用與通知單系統")
st.markdown("請選擇上方分頁切換您要使用的功能！")

# 🌟 建立兩個分頁
tab1, tab2 = st.tabs(["📝 雲端試算表自動版 (現有功能)", "🖨️ 懶人 PDF 產生器 (新功能測試)"])

# =====================================================================
# 🌟 分頁 1：原本的雲端自動系統
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
        if b_code in ["1", "全"]: return True
            
        if b_subj == "英" and b_code == str(s_eng).strip(): return True
        if b_subj == "數" and b_code == str(s_math).strip(): return True
        if b_subj == "自" and b_code == str(s_sci).strip(): return True
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
                        st.download_button(
                            label="📥 下載【家長通知單】(A4列印版)",
                            data=receipts_data,
                            file_name="家長購書通知單.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    with col2:
                        st.download_button(
                            label="📥 下載【導師對帳總表】(收費明細)",
                            data=master_data,
                            file_name="導師收費對帳總表.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
            except Exception as e:
                st.error("❌ 讀取失敗！請確認網址是否正確，且【共用】權限已設為「知道連結的人即可檢視」。")
                st.warning(f"系統錯誤代碼：{e}")
        else:
            st.warning("⚠️ 無法解析網址，請確認您貼上的是正確的網址。")


# =====================================================================
# 🌟 分頁 2：新的懶人 PDF 產生器測試區
# =====================================================================
with tab2:
    st.subheader("🖨️ 懶人版 PDF 通知單生成器 (開發測試中)")
    st.markdown("請上傳您的 Excel 檔案，系統將為您合併產出 PDF。這個功能可以直接吃學校的生肉資料！")

    st.markdown("#### 📁 第一步：上傳檔案")

    col_a, col_b = st.columns(2)
    with col_a:
        file_class = st.file_uploader("1. 班級總名單", type=["xlsx", "xls"])
        file_books = st.file_uploader("2. 書商估價單", type=["xlsx", "xls"])
        file_eng   = st.file_uploader("3. 英文分組名單", type=["xlsx", "xls"])
    with col_b:
        file_math  = st.file_uploader("4. 數學分組名單", type=["xlsx", "xls"])
        file_extra = st.file_uploader("5. 補充檔案 (可選)", type=["xlsx", "xls"])

    def create_test_pdf(file_names):
        """測試用的 PDF 產生引擎"""
        pdf = FPDF()
        pdf.add_page()
        
        # 測試版先用內建英文字型 (之後再幫您套入繁體中文字型)
        pdf.set_font("Arial", size=16)
        pdf.cell(200, 10, txt="System Test: PDF Generation OK!", ln=True, align='C')
        pdf.ln(10)
        
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt="Uploaded Files Received:", ln=True)
        
        for name in file_names:
            safe_name = name.encode('latin-1', 'replace').decode('latin-1')
            pdf.cell(200, 10, txt=f"- {safe_name}", ln=True)
            
        pdf.ln(20)
        pdf.cell(200, 10, txt="Next Step: We will build the real table here.", ln=True)
        
        return pdf.output(dest='S').encode('latin-1')

    st.divider()
    st.markdown("#### 🚀 第二步：產出 PDF")

    required_files = [file_class, file_books, file_eng, file_math]

    if all(required_files):
        st.success("✅ 必要檔案皆已上傳！可以開始測試生成 PDF。")
        
        if st.button("產生 A4 測試通知單", type="primary", key="pdf_btn"):
            with st.spinner("PDF 產生中..."):
                files_uploaded = [file_class, file_books, file_eng, file_math, file_extra]
                file_names = [f.name for f in files_uploaded if f is not None]
                
                pdf_data = create_test_pdf(file_names)
                
                st.download_button(
                    label="📥 下載 PDF",
                    data=pdf_data,
                    file_name="測試通知單.pdf",
                    mime="application/pdf"
                )
    else:
        st.info("請先上傳前 4 個必要的 Excel 檔案，生成按鈕才會出現喔！")
