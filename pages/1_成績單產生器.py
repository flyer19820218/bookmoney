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
import matplotlib
matplotlib.use('Agg')

# 強制設定字體，解決中文字爆炸
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

def create_radar_chart(labels, student_scores, avg_scores):
    num_vars = len(labels)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    
    student_scores += [student_scores[0]]
    avg_scores += [avg_scores[0]]
    angles += [angles[0]]
    
    fig, ax = plt.subplots(figsize=(5, 5), subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    
    ax.set_thetagrids(np.degrees(angles[:-1]), labels, fontsize=12)
    ax.set_ylim(0, 100)
    
    # 畫學生線
    ax.plot(angles, student_scores, color='#4F81BD', linewidth=2, label='學生個人')
    ax.fill(angles, student_scores, color='#4F81BD', alpha=0.25)
    
    # 畫班級平均線
    ax.plot(angles, avg_scores, color='#C0504D', linewidth=2, linestyle='--', label='班級平均')
    
    plt.legend(loc='upper right', bbox_to_anchor=(0.1, 0.1))
    
    img_io = io.BytesIO()
    plt.savefig(img_io, format='png', bbox_inches='tight', dpi=100)
    plt.close(fig)
    img_io.seek(0)
    return img_io

# [後續 PDF 生成邏輯...]
# 在 generate_pdf_report 中呼叫時，傳入 avg_scores 即可
