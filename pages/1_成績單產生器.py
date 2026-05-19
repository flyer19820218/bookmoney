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

# 確保 Matplotlib 不會與 Streamlit 衝突
import matplotlib
matplotlib.use('Agg')

# ==========================================
# 0. 網頁基本設定 & 字體
# ==========================================
st.set_page_config(page_title="AI 成績單產生器", layout="centered", page_icon="📈")

# 自動處理字體，避免 PDF 亂碼
FONT_NAME = "NotoSansTC-Regular.ttf"
if not os.path.exists(FONT_NAME):
    try:
        import urllib.request
        urllib.request.urlretrieve("https://github.com/google/fonts/raw/main/ofl/notosanstc/NotoSansTC-Regular.ttf", FONT_NAME)
    except:
        pass

if os.path.exists(FONT_NAME):
    plt.rcParams['font.sans-serif'] = ['Noto Sans TC', 'Microsoft JhengHei', 'sans-serif']
    plt.rcParams['axes.unicode_minus'] = False

# ==========================================
# 1. 核心邏輯區
# ==========================================
def get_google_sheet_csv_url(url):
    pattern = r'https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9-_]+)'
    match = re.search(pattern, url)
    if match:
        sheet_id = match.group(1)
        gid_match = re.search(r'gid=([0-9]+)', url)
        gid_param = f"&gid={gid_match.group(1)}" if gid_match else ""
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv{gid_param}"
    return None

def create_radar_chart(labels, student_scores, avg_scores):
    num_vars = len(labels)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    student_scores = student_scores + [student_scores[0]]
    avg_scores = avg_scores + [avg_scores[0]]
    angles = angles + [angles[0]]
    
    fig, ax = plt.subplots(figsize=(4, 4), subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.set_thetagrids(np.degrees(angles[:-1]), labels, fontsize=12)
    ax.set_ylim(0, 100)
    
    ax.plot(angles, student_scores, color='#4F81BD', linewidth=2, label='學生個人')
    ax.fill(angles, student_scores, color='#4F81BD', alpha=0.25)
    ax.plot(angles, avg_scores, color='#C0504D', linewidth=2, linestyle='--', label='班級平均')
    plt.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1))
    
    img_io = io.BytesIO()
    plt.savefig(img_io, format='png', bbox_inches='tight', dpi=100)
    plt.close(fig)
    img_io.seek(0)
    return img_io

def generate_pdf(df):
    pdf_io = io.BytesIO()
    c = canvas.Canvas(pdf_io, pagesize=A4)
    w, h = A4
    
    # 註冊字體
    if os.path.exists(FONT_NAME):
        pdfmetrics.registerFont(TTFont('CustomFont', FONT_NAME))
        font = 'CustomFont'
    else:
        font = 'Helvetica'

    # 計算全班平均與 PR 值
    has_7_sub = all(col in df.columns for col in ['歷史', '地理', '公民'])
    
    # 統一計算總分與 PR
    def calc_total(row):
        base = row['國文'] + row['英文'] + row['數學'] + row['自然']
        return base + ((row['歷史'] + row['地理'] + row['公民'])/3 if has_7_sub else row['社會'])
    
    df['總分'] = df.apply(calc_total, axis=1)
    df['PR值'] = df['總分'].rank(pct=True) * 100
    
    # 計算各科班級平均供雷達圖用
    if has_7_sub:
        subjects = ['國文', '英文', '數學', '自然', '歷史', '地理', '公民']
        avg_scores = [df[s].mean() for s in subjects]
    else:
        subjects = ['國文', '英文', '數學', '自然', '社會']
        avg_scores = [df[s].mean() for s in subjects]

    for _, row in df.iterrows():
        # 繪製內容
        c.setFont(font, 20)
        c.drawCentredString(w/2, h-60, "學生個人成績單")
        c.setFont(font, 12)
        c.drawString(60, h-100, f"座號: {int(row['座號'])}  姓名: {row['姓名']}")
        
        y = h - 140
        for s in subjects:
            c.drawString(70, y, f"{s}: {row[s]:.1f}")
            y -= 20
        
        c.drawString(70, y-20, f"總分: {row['總分']:.1f}")
        c.drawString(70, y-40, f"PR值: {row['PR值']:.1f}")
        
        # 繪圖
        student_scores = [float(row[s]) for s in subjects]
        chart = create_radar_chart(subjects, student_scores, avg_scores)
        c.drawImage(ImageReader(chart), 200, h-400, 300, 300, mask='auto')
        c.showPage()
        
    c.save()
    pdf_io.seek(0)
    return pdf_io

# ==========================================
# 3. 網頁介面
# ==========================================
st.title("📈 801 班級成績單產生器")
url = st.text_input("輸入成績試算表網址：")

if url and (csv_url := get_google_sheet_csv_url(url)):
    df = pd.read_csv(csv_url)
    # 過濾：座號必須是數字且 > 0
    df['座號'] = pd.to_numeric(df['座號'], errors='coerce')
    df = df[df['座號'] > 0].sort_values('座號')
    
    st.success(f"已偵測到 {len(df)} 位學生")
    if st.button("產生 PDF"):
        pdf = generate_pdf(df)
        st.download_button("下載成績單", pdf, "成績單.pdf", "application/pdf")
