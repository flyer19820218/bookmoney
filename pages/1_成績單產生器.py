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
import re

# 確保繪圖不會因為環境衝突而失效
import matplotlib
matplotlib.use('Agg')

st.set_page_config(page_title="班級成績單公版產生器", layout="wide")

# ==========================================
# 1. 核心邏輯：通用資料處理
# ==========================================
def get_csv_from_url(url):
    try:
        if "docs.google.com" in url:
            # 強制轉換為 CSV 匯出連結
            match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
            if match:
                sheet_id = match.group(1)
                return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
        return url
    except:
        return None

# ==========================================
# 2. PDF 生成核心 (公版邏輯)
# ==========================================
def generate_pdf(df, subjects):
    pdf_io = io.BytesIO()
    c = canvas.Canvas(pdf_io, pagesize=A4)
    w, h = A4
    
    # 嘗試載入中文字體 (若無法載入，改用 Helvetica 避免程式崩潰)
    font_name = "msjh"
    try:
        # 請確保 GitHub 根目錄有 msjh.ttf，或是系統已有此字體
        pdfmetrics.registerFont(TTFont(font_name, 'msjh.ttf'))
    except:
        font_name = "Helvetica"
        st.warning("警告：未偵測到中文字體檔(msjh.ttf)，中文可能會顯示為空白。")

    # 計算全班平均
    avg_vals = [df[s].mean() for s in subjects]
    
    for _, row in df.iterrows():
        c.setFont(font_name, 20)
        c.drawCentredString(w/2, h-50, "學生個人成績單")
        
        c.setFont(font_name, 14)
        c.drawString(60, h-100, f"座號: {row['座號']}      姓名: {row['姓名']}")
        c.line(60, h-110, w-60, h-110)
        
        # 顯示科目分數
        y = h - 140
        for s in subjects:
            score = row[s] if pd.notna(row[s]) else 0
            c.drawString(70, y, f"{s}: {float(score):.1f}")
            y -= 25
            
        c.drawString(70, y-20, f"總分: {float(row['總分']):.1f}")
        c.drawString(70, y-50, f"PR值: {float(row['PR值']):.1f}")
        
        c.showPage()
    
    c.save()
    pdf_io.seek(0)
    return pdf_io

# ==========================================
# 3. UI 介面
# ==========================================
st.title("📊 班級成績單通用公版")
st.markdown("請確認您的 Google Sheet 第一列包含：**座號、姓名、國文、英文、數學、自然、社會** (或歷史、地理、公民)")

url = st.text_input("輸入成績試算表網址：")

if url:
    csv_url = get_csv_from_url(url)
    try:
        df = pd.read_csv(csv_url)
        
        # 自動清理資料：移除沒有座號的列
        df['座號'] = pd.to_numeric(df['座號'], errors='coerce')
        df = df.dropna(subset=['座號']).sort_values('座號')
        
        # 動態偵測科目
        all_cols = df.columns.tolist()
        potential_subjects = ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']
        found_subjects = [s for s in potential_subjects if s in all_cols]
        
        st.write(f"✅ 系統已抓取到 {len(df)} 位學生，偵測到的科目: {', '.join(found_subjects)}")
        
        # 計算總分與 PR
        if '歷史' in found_subjects: # 七科版
            df['總分'] = df['國文'] + df['英文'] + df['數學'] + df['自然'] + (df['歷史'] + df['地理'] + df['公民'])/3
        else: # 五科版
            df['總分'] = df['國文'] + df['英文'] + df['數學'] + df['自然'] + df['社會']
            
        df['PR值'] = df['總分'].rank(pct=True) * 100
        
        st.dataframe(df.head(5)) # 預覽前5筆
        
        if st.button("產生全班 PDF 成績單"):
            pdf = generate_pdf(df, found_subjects)
            st.download_button("下載 PDF", pdf, "成績單.pdf", "application/pdf")
            
    except Exception as e:
        st.error(f"❌ 發生錯誤，請檢查您的試算表格式: {e}")
