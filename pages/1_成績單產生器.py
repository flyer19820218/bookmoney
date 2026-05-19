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

# 強制設定 Matplotlib 環境
import matplotlib
matplotlib.use('Agg')

# 【關鍵修正】請務必確保這個檔案在 GitHub 倉庫裡
FONT_PATH = "NotoSansTC-Regular.ttf" 

# 設定 Matplotlib 全局字體
if os.path.exists(FONT_PATH):
    plt.rcParams['font.sans-serif'] = ['NotoSans']
    plt.rcParams['axes.unicode_minus'] = False
else:
    # 如果找不到字體，這裡先暫時用英文頂著，但我建議您一定要上傳字體檔
    pass

# ==========================================
# 核心雷達圖繪製 (加入了字體註冊)
# ==========================================
def create_radar_chart(labels, student_scores, avg_scores):
    num_vars = len(labels)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    
    student_scores = student_scores + [student_scores[0]]
    avg_scores = avg_scores + [avg_scores[0]]
    angles = angles + [angles[0]]
    
    fig, ax = plt.subplots(figsize=(4, 4), subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    
    # 這裡確保 labels 顯示中文
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

# ==========================================
# 核心 PDF 產生
# ==========================================
def generate_pdf(df):
    pdf_io = io.BytesIO()
    c = canvas.Canvas(pdf_io, pagesize=A4)
    w, h = A4
    
    # 【關鍵修正】強制將字體註冊給 ReportLab
    if os.path.exists(FONT_PATH):
        pdfmetrics.registerFont(TTFont('NotoSans', FONT_PATH))
        font = 'NotoSans'
    else:
        font = 'Helvetica'
        st.error("找不到字體檔！請確認 NotoSansTC-Regular.ttf 是否已上傳至 GitHub。")

    # (中間計算分數邏輯同上...)
    # 畫圖時改用註冊的字體：
    # c.setFont(font, 20)
    # c.drawCentredString(...)
    
    # ... 其餘邏輯相同 ...
