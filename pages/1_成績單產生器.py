import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
import re
import io
import requests
import os

# 解決 matplotlib 在 Streamlit 上的執行緒警告
import matplotlib
matplotlib.use('Agg')

# ==========================================
# 0. 網頁基本設定 & 字體準備
# ==========================================
st.set_page_config(page_title="AI 成績單產生器", layout="centered", page_icon="📈")

# 【防護機制】確保有字體檔可以畫中文。如果真的沒有，就先用系統預設
# 老師請注意：若要在 Streamlit Cloud 完美顯示中文，建議將一個中文字體檔 (如 msjh.ttc 或 NotoSans.ttf) 
# 直接上傳到您的 GitHub 資料夾，並將下方的 'msjh.ttc' 改成您的字體檔名。
FONT_NAME = "NotoSansTC-Regular.ttf"
HAS_FONT = os.path.exists(FONT_NAME)

if not HAS_FONT:
    # 退而求其次的備用方案下載
    try:
        import urllib.request
        urllib.request.urlretrieve("https://github.com/google/fonts/raw/main/ofl/notosanstc/NotoSansTC-Regular.ttf", FONT_NAME)
        HAS_FONT = True
    except:
        pass

if HAS_FONT:
    plt.rcParams['font.sans-serif'] = ['Noto Sans TC', 'Microsoft JhengHei', 'sans-serif']
    plt.rcParams['axes.unicode_minus'] = False

# ==========================================
# 1. 說明區
# ==========================================
st.title("📈 全校通用 AI 成績單產生器")
st.markdown("貼上 Google 試算表，自動計算五科平均，並產出附帶**雷達圖**的精美 PDF 成績單！")

with st.expander("👉 點我查看【試算表標準格式】說明", expanded=True):
    st.info("""
    **請確保您的試算表第一列有以下欄位名稱 (順序不拘)：**
    * **一年級版**：`座號`、`姓名`、`國文`、`英文`、`數學`、`自然`、`社會`
    * **二三年級版**：`座號`、`姓名`、`國文`、`英文`、`數學`、`自然`、`歷史`、`地理`、`公民`
    """)

# ==========================================
# 2. 核心邏輯區
# ==========================================
def get_google_sheet_csv_url(url):
    try:
        if "docs.google.com" not in url:
            response = requests.head(url, allow_redirects=True)
            url = response.url
        pattern = r'https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9-_]+)'
        match = re.search(pattern, url)
        if match:
            sheet_id = match.group(1)
            return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
        return None
    except:
        return None

def create_radar_chart(labels, scores):
    """繪製雷達圖並回傳圖片記憶體物件"""
    num_vars = len(labels)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    
    # 完成圓形閉環
    scores = scores + [scores[0]]
    angles = angles + [angles[0]]
    
    # 使用 Agg 背景畫圖，避免 Streamlit 報錯
    fig, ax = plt.subplots(figsize=(4, 4), subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    
    # 畫座標軸與標籤
    # 如果沒有字體，可能會顯示方塊，但程式不會崩潰
    ax.set_thetagrids(np.degrees(angles[:-1]), labels, fontsize=12)
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], color="grey", size=8)
    
    # 填滿圖表
    ax.plot(angles, scores, color='#4F81BD', linewidth=2, linestyle='solid')
    ax.fill(angles, scores, color='#4F81BD', alpha=0.4)
    
    # 存成圖片並徹底關閉畫布釋放記憶體
    img_io = io.BytesIO()
    plt.savefig(img_io, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    img_io.seek(0)
    return img_io

def generate_pdf_report(df):
    """產生包含雷達圖的 PDF 檔案"""
    pdf_io = io.BytesIO()
    c = canvas.Canvas(pdf_io, pagesize=A4)
    width, height = A4
    
    # PDF 中文字體註冊
    font_registered = False
    if HAS_FONT:
        try:
            pdfmetrics.registerFont(TTFont('CustomFont', FONT_NAME))
            font_registered = True
        except:
            pass
            
    current_font = 'CustomFont' if font_registered else 'Helvetica'
    
    # 判斷是 5 科還是 7 科
    has_7_subjects = all(col in df.columns for col in ['歷史', '地理', '公民'])
    
    for _, row in df.iterrows():
        seat, name = row.get('座號', ''), row.get('姓名', '')
        
        # --- 處理分數邏輯 ---
        scores_dict = {}
        if has_7_subjects:
            subjects = ['國文', '英文', '數學', '自然', '歷史', '地理', '公民']
            radar_labels = subjects
            for sub in subjects:
                scores_dict[sub] = pd.to_numeric(row.get(sub, 0), errors='coerce')
            
            radar_scores = [scores_dict[s] for s in subjects]
            # 社會科平均 = (歷+地+公) / 3
            soc_avg = (scores_dict['歷史'] + scores_dict['地理'] + scores_dict['公民']) / 3
            total = scores_dict['國文'] + scores_dict['英文'] + scores_dict['數學'] + scores_dict['自然'] + soc_avg
            avg = total / 5
            
            print_text = [
                f"國文: {scores_dict['國文']}", f"英文: {scores_dict['英文']}", 
                f"數學: {scores_dict['數學']}", f"自然: {scores_dict['自然']}",
                f"歷史: {scores_dict['歷史']} | 地理: {scores_dict['地理']} | 公民: {scores_dict['公民']}"
            ]
        else:
            subjects = ['國文', '英文', '數學', '自然', '社會']
            radar_labels = subjects
            for sub in subjects:
                scores_dict[sub] = pd.to_numeric(row.get(sub, 0), errors='coerce')
                
            radar_scores = [scores_dict[s] for s in subjects]
            total = sum(radar_scores)
            avg = total / 5
            print_text = [f"{s}: {scores_dict[s]}" for s in subjects]

        # --- 繪製 PDF 版面 ---
        c.setFont(current_font, 24)
        c.drawCentredString(width/2, height - 80, "學 生 個 人 成 績 單" if font_registered else "Student Report Card")
        
        c.setFont(current_font, 14)
        c.drawString(100, height - 130, f"座號(No.): {int(seat) if pd.notna(seat) else ''}      姓名(Name): {name}")
        c.line(100, height - 135, width - 100, height - 135) 
        
        y_pos = height - 170
        c.setFont(current_font, 12)
        for text in print_text:
            c.drawString(120, y_pos, text)
            y_pos -= 25
            
        y_pos -= 10
        c.drawString(120, y_pos, f"五科總分(Total): {total:.1f}")
        c.drawString(120, y_pos - 25, f"五科平均(Average): {avg:.2f}")

        # --- 繪製雷達圖 ---
        chart_img = create_radar_chart(radar_labels, radar_scores)
        c.drawImage(ImageReader(chart_img), 280, height - 420, width=250, height=250, mask='auto')
        
        c.showPage()
        
    c.save()
    pdf_io.seek(0)
    return pdf_io

# ==========================================
# 3. 介面操作區
# ==========================================
sheet_url = st.text_input("🔗 請在下方輸入您的成績試算表網址：", placeholder="https://docs.google.com/spreadsheets/...")

if sheet_url:
    csv_url = get_google_sheet_csv_url(sheet_url)
    if csv_url:
        try:
            with st.spinner("正在連線讀取成績資料..."):
                df_scores = pd.read_csv(csv_url).fillna(0)
            
            st.success(f"✅ 成功讀取 {len(df_scores)} 位學生的成績！")
            
            if st.button("🚀 一鍵產生 PDF 成績單 (附雷達圖)", type="primary"):
                with st.spinner("AI 正在繪製成績地圖與排版 PDF..."):
                    pdf_data = generate_pdf_report(df_scores)
                    
                st.balloons()
                st.download_button(
                    label="📥 下載全班 PDF 成績單",
                    data=pdf_data,
                    file_name="全班個人成績單.pdf",
                    mime="application/pdf"
                )
        except Exception as e:
            st.error("❌ 產生失敗！請檢查試算表欄位名稱是否為：座號、姓名、國文、英文...")
            st.warning(f"系統錯誤訊息：{e}")
