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

# 解決 matplotlib 執行緒警告
import matplotlib
matplotlib.use('Agg')

# ==========================================
# 0. 網頁基本設定 & 字體準備
# ==========================================
st.set_page_config(page_title="AI 成績單產生器", layout="centered", page_icon="📈")

# 【字體防呆】如果沒有上傳字體，就試著用系統內建的
FONT_NAME = "NotoSansTC-Regular.ttf"
HAS_FONT = os.path.exists(FONT_NAME)
if not HAS_FONT:
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
st.markdown("貼上成績連結，系統會**自動偵測班級人數**並排除底部的平均值與圖表，精準產出 **PDF 雷達圖成績單**！")

with st.expander("👉 點我查看【試算表標準格式】說明", expanded=True):
    st.info("""
    **請確保您的試算表第一列有以下欄位名稱 (順序不拘)：**
    * **一年級版 (5科)**：`座號`、`姓名`、`國文`、`英文`、`數學`、`自然`、`社會`
    * **二三年級版 (7科)**：`座號`、`姓名`、`國文`、`英文`、`數學`、`自然`、`歷史`、`地理`、`公民`
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
            # 支援特定工作表 (gid) 的抓取
            gid_match = re.search(r'gid=([0-9]+)', url)
            gid_param = f"&gid={gid_match.group(1)}" if gid_match else ""
            return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv{gid_param}"
        return None
    except:
        return None

def create_radar_chart(labels, scores):
    num_vars = len(labels)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    
    scores = scores + [scores[0]]
    angles = angles + [angles[0]]
    
    fig, ax = plt.subplots(figsize=(4, 4), subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    
    ax.set_thetagrids(np.degrees(angles[:-1]), labels, fontsize=12)
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], color="grey", size=8)
    
    ax.plot(angles, scores, color='#4F81BD', linewidth=2, linestyle='solid')
    ax.fill(angles, scores, color='#4F81BD', alpha=0.4)
    
    img_io = io.BytesIO()
    plt.savefig(img_io, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    img_io.seek(0)
    return img_io

def generate_pdf_report(df):
    pdf_io = io.BytesIO()
    c = canvas.Canvas(pdf_io, pagesize=A4)
    width, height = A4
    
    font_registered = False
    if HAS_FONT:
        try:
            pdfmetrics.registerFont(TTFont('CustomFont', FONT_NAME))
            font_registered = True
        except:
            pass
            
    current_font = 'CustomFont' if font_registered else 'Helvetica'
    
    has_7_subjects = all(col in df.columns for col in ['歷史', '地理', '公民'])
    
    for _, row in df.iterrows():
        seat, name = row.get('座號', ''), row.get('姓名', '')
        name = str(name)
        
        scores_dict = {}
        if has_7_subjects:
            subjects = ['國文', '英文', '數學', '自然', '歷史', '地理', '公民']
            radar_labels = subjects
            for sub in subjects:
                scores_dict[sub] = pd.to_numeric(row.get(sub, 0), errors='coerce')
                if pd.isna(scores_dict[sub]): scores_dict[sub] = 0
            
            radar_scores = [scores_dict[s] for s in subjects]
            soc_avg = (scores_dict['歷史'] + scores_dict['地理'] + scores_dict['公民']) / 3
            total = scores_dict['國文'] + scores_dict['英文'] + scores_dict['數學'] + scores_dict['自然'] + soc_avg
            avg = total / 5
            
            print_text = [
                f"國文: {scores_dict['國文']}", f"英文: {scores_dict['英文']}", 
                f"數學: {scores_dict['數學']}", f"自然: {scores_dict['自然']}",
                f"歷史: {scores_dict['歷史']}  地理: {scores_dict['地理']}  公民: {scores_dict['公民']}"
            ]
        else:
            subjects = ['國文', '英文', '數學', '自然', '社會']
            radar_labels = subjects
            for sub in subjects:
                scores_dict[sub] = pd.to_numeric(row.get(sub, 0), errors='coerce')
                if pd.isna(scores_dict[sub]): scores_dict[sub] = 0
                
            radar_scores = [scores_dict[s] for s in subjects]
            total = sum(radar_scores)
            avg = total / 5
            print_text = [f"{s}: {scores_dict[s]}" for s in subjects]

        # --- PDF 版面設計 ---
        c.setFont(current_font, 22)
        c.drawCentredString(width/2, height - 70, "學 生 個 人 成 績 單")
        
        c.setFont(current_font, 14)
        c.drawString(60, height - 120, f"座號: {int(seat)}      姓名: {name}")
        c.line(60, height - 130, width - 60, height - 130)
        
        y_pos = height - 170
        c.setFont(current_font, 12)
        for text in print_text:
            c.drawString(70, y_pos, text)
            y_pos -= 25
            
        y_pos -= 15
        c.setFont(current_font, 14)
        c.drawString(70, y_pos, f"⭐ 五科總分: {total:.1f}")
        c.drawString(70, y_pos - 30, f"⭐ 五科平均: {avg:.2f}")

        # 右側雷達圖
        chart_img = create_radar_chart(radar_labels, radar_scores)
        c.drawImage(ImageReader(chart_img), width - 320, height - 420, width=280, height=280, mask='auto')
        
        c.showPage()
        
    c.save()
    pdf_io.seek(0)
    return pdf_io

# ==========================================
# 3. 介面操作區
# ==========================================
# 預設直接放老師的網址方便測試
default_url = "https://docs.google.com/spreadsheets/d/1lp0F45BnLO0Hn2l47vJ7orawr0KfIaT_/edit?gid=1366647975#gid=1366647975"
sheet_url = st.text_input("🔗 成績試算表網址：", value=default_url)

if sheet_url:
    csv_url = get_google_sheet_csv_url(sheet_url)
    if csv_url:
        try:
            with st.spinner("正在連線過濾資料..."):
                df_scores = pd.read_csv(csv_url)
                
                # 【智慧核心】動態過濾學生人數，不綁死 31 這個數字
                # 1. 將座號轉換為數字，如果遇到「總平均」、「分佈圖」這種文字，會被變成空值 (NaN)
                df_scores['座號'] = pd.to_numeric(df_scores['座號'], errors='coerce')
                
                # 2. 刪除空值，並只保留座號 > 0 的列 (徹底過濾雜訊)
                df_scores = df_scores.dropna(subset=['座號'])
                df_scores = df_scores[df_scores['座號'] > 0]
                
                # 3. 重新排序並重置索引
                df_scores = df_scores.sort_values(by='座號').reset_index(drop=True)
                
                # 動態取得學生人數
                student_count = len(df_scores)
                
            st.success(f"✅ 成功鎖定！系統自動偵測到本班共 **{student_count}** 位學生，並已完美過濾下方圖表雜訊。")
            st.dataframe(df_scores[['座號', '姓名', '國文', '英文', '數學', '自然']].head(5)) # 預覽前五筆
            
            if st.button(f"🚀 一鍵產生 {student_count} 人 PDF 成績單", type="primary"):
                with st.spinner("AI 正在繪製成績地圖與排版 PDF，這會花幾秒鐘的時間..."):
                    pdf_data = generate_pdf_report(df_scores)
                    
                st.balloons()
                st.download_button(
                    label=f"📥 下載全班 ({student_count}人) PDF 成績單",
                    data=pdf_data,
                    file_name="全班個人成績單_雷達圖版.pdf",
                    mime="application/pdf"
                )
        except Exception as e:
            st.error("❌ 讀取失敗！請確認試算表欄位名稱是否包含：座號、姓名、國文、英文...")
            st.warning(f"錯誤細節：{e}")
