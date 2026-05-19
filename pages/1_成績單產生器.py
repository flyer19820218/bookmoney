import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
import re
import io
import requests
import os
import urllib.request

# ==========================================
# 0. 網頁基本與字體設定 (防呆：自動下載中文字體)
# ==========================================
st.set_page_config(page_title="AI 成績單產生器", layout="centered", page_icon="📈")

# 自動下載 NotoSans 中文字體，解決 PDF 亂碼問題
font_path = "NotoSansTC-Regular.ttf"
if not os.path.exists(font_path):
    with st.spinner("首次啟動，正在為您安裝中文字體，請稍候..."):
        font_url = "https://github.com/google/fonts/raw/main/ofl/notosanstc/NotoSansTC-Regular.ttf"
        urllib.request.urlretrieve(font_url, font_path)

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
    
    *(系統會自動判斷您是否有輸入史地公，自動切換 5 科或 7 科雷達圖，並統一以「5科」計算總分與平均)*
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
    font_prop = FontProperties(fname=font_path)
    num_vars = len(labels)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    
    # 完成圓形閉環
    scores = scores + [scores[0]]
    angles = angles + [angles[0]]
    
    fig, ax = plt.subplots(figsize=(4, 4), subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    
    # 畫座標軸與標籤
    ax.set_thetagrids(np.degrees(angles[:-1]), labels, fontproperties=font_prop, fontsize=12)
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], color="grey", size=8)
    
    # 填滿圖表
    ax.plot(angles, scores, color='#4F81BD', linewidth=2, linestyle='solid')
    ax.fill(angles, scores, color='#4F81BD', alpha=0.4)
    
    # 存成圖片
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
    
    # 註冊中文字體
    pdfmetrics.registerFont(TTFont('NotoSans', font_path))
    
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
                f"國文：{scores_dict['國文']}", f"英文：{scores_dict['英文']}", 
                f"數學：{scores_dict['數學']}", f"自然：{scores_dict['自然']}",
                f"歷史：{scores_dict['歷史']} | 地理：{scores_dict['地理']} | 公民：{scores_dict['公民']}"
            ]
        else:
            subjects = ['國文', '英文', '數學', '自然', '社會']
            radar_labels = subjects
            for sub in subjects:
                scores_dict[sub] = pd.to_numeric(row.get(sub, 0), errors='coerce')
                
            radar_scores = [scores_dict[s] for s in subjects]
            total = sum(radar_scores)
            avg = total / 5
            print_text = [f"{s}：{scores_dict[s]}" for s in subjects]

        # --- 繪製 PDF 版面 ---
        # 標題
        c.setFont('NotoSans', 24)
        c.drawCentredString(width/2, height - 80, "學 生 個 人 成 績 單")
        
        # 基本資料
        c.setFont('NotoSans', 14)
        c.drawString(100, height - 130, f"座號：{int(seat) if pd.notna(seat) else ''}      姓名：{name}")
        c.line(100, height - 135, width - 100, height - 135) # 分隔線
        
        # 列印各科分數
        y_pos = height - 170
        c.setFont('NotoSans', 12)
        for text in print_text:
            c.drawString(120, y_pos, text)
            y_pos -= 25
            
        # 列印總分與平均 (粗體效果：疊字)
        y_pos -= 10
        c.drawString(120, y_pos, f"➡ 五科總分：{total:.1f}")
        c.drawString(120, y_pos - 25, f"➡ 五科平均：{avg:.2f}")

        # --- 繪製雷達圖 ---
        chart_img = create_radar_chart(radar_labels, radar_scores)
        # 將圖片放入 PDF (x, y, width, height)
        c.drawImage(ImageReader(chart_img), 280, height - 420, width=250, height=250, mask='auto')
        
        # 分頁
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
            st.dataframe(df_scores.head(3)) # 預覽前三筆
            
            if st.button("🚀 一鍵產生 PDF 成績單 (附雷達圖)", type="primary"):
                with st.spinner("AI 正在繪製成績地圖與排版 PDF，這會花幾秒鐘的時間..."):
                    pdf_data = generate_pdf_report(df_scores)
                    
                st.balloons()
                st.download_button(
                    label="📥 下載全班 PDF 成績單",
                    data=pdf_data,
                    file_name="全班個人成績單.pdf",
                    mime="application/pdf"
                )
        except Exception as e:
            st.error("❌ 讀取失敗！請確認試算表格式是否正確，以及是否已開啟「知道連結的人即可檢視」。")
            st.warning(f"錯誤細節：{e}")
