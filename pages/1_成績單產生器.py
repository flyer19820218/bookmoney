import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
import io
import requests

# 為了穩定，固定使用 Agg 畫圖
import matplotlib
matplotlib.use('Agg')

# ==========================================
# 1. 核心繪圖函數 (強化雷達圖)
# ==========================================
def create_radar_chart(labels, scores, avg_scores):
    # 設定字體 (嘗試使用內建或指定)
    try:
        plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'sans-serif']
    except:
        pass
    
    num_vars = len(labels)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    
    scores = scores + [scores[0]]
    avg_scores = avg_scores + [avg_scores[0]]
    angles = angles + [angles[0]]
    
    fig, ax = plt.subplots(figsize=(4, 4), subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    
    ax.set_thetagrids(np.degrees(angles[:-1]), labels, fontsize=10)
    ax.set_ylim(0, 100)
    
    # 畫線
    ax.plot(angles, scores, color='#1f77b4', linewidth=2, label='學生')
    ax.fill(angles, scores, color='#1f77b4', alpha=0.2)
    ax.plot(angles, avg_scores, color='#d62728', linewidth=2, linestyle='--', label='平均')
    
    img_io = io.BytesIO()
    plt.savefig(img_io, format='png', bbox_inches='tight', dpi=100)
    plt.close(fig)
    img_io.seek(0)
    return img_io

# ==========================================
# 2. 核心 PDF 產生函數 (徹底修正排版)
# ==========================================
def generate_pdf(df):
    pdf_io = io.BytesIO()
    c = canvas.Canvas(pdf_io, pagesize=A4)
    w, h = A4
    
    # 字體註冊 (若倉庫有字體檔則載入)
    font = 'Helvetica' # 預設英數
    if os.path.exists("msjh.ttf"):
        pdfmetrics.registerFont(TTFont('msjh', 'msjh.ttf'))
        font = 'msjh'
    
    has_7_sub = all(col in df.columns for col in ['歷史', '地理', '公民'])
    subjects = ['國文', '英文', '數學', '自然', '歷史', '地理', '公民'] if has_7_sub else ['國文', '英文', '數學', '自然', '社會']
    
    # 計算全班平均供雷達圖用
    avg_vals = [df[s].mean() for s in subjects]
    
    for _, row in df.iterrows():
        c.setFont(font, 20)
        c.drawCentredString(w/2, h-50, "學生個人成績單")
        
        c.setFont(font, 14)
        c.drawString(60, h-100, f"座號: {int(row['座號'])}      姓名: {row['姓名']}")
        c.line(60, h-110, w-60, h-110)
        
        # 分數明細
        y = h - 150
        for s in subjects:
            score = row[s] if pd.notna(row[s]) else 0
            c.drawString(70, y, f"{s}: {float(score):.1f}")
            y -= 25
            
        # 總分與 PR
        c.setFont(font, 16)
        c.drawString(70, y-20, f"五科總分: {row['總分']:.1f}")
        c.drawString(70, y-50, f"PR值: {row['PR值']:.1f}")
        
        # 繪圖
        s_scores = [float(row[s]) if pd.notna(row[s]) else 0 for s in subjects]
        chart = create_radar_chart(subjects, s_scores, avg_vals)
        c.drawImage(ImageReader(chart), 250, h-400, 300, 300, mask='auto')
        c.showPage()
    
    c.save()
    pdf_io.seek(0)
    return pdf_io

# ==========================================
# 3. Streamlit 網頁主程式
# ==========================================
st.title("📊 801 班級成績單自動化系統")
url = st.text_input("輸入成績試算表網址：")

if url:
    # 抓資料並清理
    try:
        csv_url = get_google_sheet_csv_url(url)
        df = pd.read_csv(csv_url)
        # 強制過濾座號 1-31
        df['座號'] = pd.to_numeric(df['座號'], errors='coerce')
        df = df[df['座號'].between(1, 31)].sort_values('座號')
        
        # 計算總分 (自動切換)
        if all(col in df.columns for col in ['歷史', '地理', '公民']):
            df['總分'] = df['國文'] + df['英文'] + df['數學'] + df['自然'] + (df['歷史'] + df['地理'] + df['公民'])/3
        else:
            df['總分'] = df['國文'] + df['英文'] + df['數學'] + df['自然'] + df['社會']
        
        df['PR值'] = df['總分'].rank(pct=True) * 100
        
        st.write(f"✅ 讀取到 {len(df)} 位學生資料")
        if st.button("產生全班成績單"):
            pdf = generate_pdf(df)
            st.download_button("下載 PDF", pdf, "成績單.pdf", "application/pdf")
    except Exception as e:
        st.error(f"資料讀取失敗，請確認欄位名稱正確：{e}")
