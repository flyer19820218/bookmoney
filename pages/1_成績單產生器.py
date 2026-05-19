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
from matplotlib import font_manager

# 解決 matplotlib 執行緒警告
import matplotlib
matplotlib.use('Agg')

# ==========================================
# 0. 網頁基本設定 & 字體準備
# ==========================================
st.set_page_config(page_title="AI 成績單產生器", layout="centered", page_icon="📈")

FONT_NAME = "NotoSansTC-Regular.ttf"
FONT_URL = "https://cdn.jsdelivr.net/gh/themoeway/noto-sans-tc-ttf@master/ttf/NotoSansTC-Regular.ttf"

@st.cache_resource
def init_fonts():
    """初始化字體：自動偵測本地字體，若無則從網路下載，並註冊到 Matplotlib 與 ReportLab"""
    if not os.path.exists(FONT_NAME):
        try:
            with st.spinner("未偵測到本地字體，正在從網路下載中文字體 (Noto Sans TC)..."):
                headers = {'User-Agent': 'Mozilla/5.0'}
                response = requests.get(FONT_URL, headers=headers, timeout=60)
                response.raise_for_status()
                with open(FONT_NAME, "wb") as f:
                    f.write(response.content)
        except Exception as e:
            st.error(f"字體下載失敗，請確保 {FONT_NAME} 已手動放置於專案根目錄。錯誤: {e}")
            return False
    
    try:
        font_manager.fontManager.addfont(FONT_NAME)
        dynamic_font_name = font_manager.FontProperties(fname=FONT_NAME).get_name()
        plt.rcParams['font.sans-serif'] = [dynamic_font_name, 'Microsoft JhengHei', 'Arial']
        plt.rcParams['axes.unicode_minus'] = False
        pdfmetrics.registerFont(TTFont('CustomFont', FONT_NAME))
        return True
    except Exception as e:
        st.error(f"字體註冊失敗: {e}")
        return False

HAS_FONT = init_fonts()

# ==========================================
# 1. 說明區
# ==========================================
st.title("📈 全校通用 AI 成績單產生器 (PR版)")
st.markdown("自動偵測人數、排除雜訊，產出包含 **各科平均、班級名次、PR值雷達圖** 的專業成績單！")

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
            gid_match = re.search(r'gid=([0-9]+)', url)
            gid_param = f"&gid={gid_match.group(1)}" if gid_match else ""
            return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv{gid_param}"
        return None
    except:
        return None

def create_pr_radar_chart(labels, pr_scores):
    """使用 PR 值繪製雷達圖 (滿分為 PR 99)"""
    num_vars = len(labels)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    
    plot_scores = list(pr_scores) + [pr_scores[0]]
    angles = angles + [angles[0]]
    
    fig, ax = plt.subplots(figsize=(4, 4), subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    
    # 標籤加上 "PR" 字樣提示學生
    display_labels = [f"{label} (PR)" for label in labels]
    ax.set_thetagrids(np.degrees(angles[:-1]), display_labels, fontsize=11)
    
    # PR 值最高到 99 (我們設定軸的最大值為 100)
    ax.set_ylim(0, 100)
    ax.set_yticks([25, 50, 75, 100])
    ax.set_yticklabels(["PR25", "PR50", "PR75", "PR99"], color="grey", size=8)
    
    # 畫 PR 50 (班級中位數) 的虛線參考基準
    ax.plot(angles, [50]*len(angles), color='#C0504D', linewidth=1.5, linestyle='--', label='PR 50 (中位數)')
    
    # 畫學生的 PR 分佈
    ax.plot(angles, plot_scores, color='#4F81BD', linewidth=2, linestyle='solid', label='個人優勢 (PR)')
    ax.fill(angles, plot_scores, color='#4F81BD', alpha=0.4)
    
    plt.legend(loc='upper right', bbox_to_anchor=(1.35, 1.15), fontsize=9)
    
    img_io = io.BytesIO()
    plt.savefig(img_io, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    img_io.seek(0)
    return img_io

def generate_pdf_report(df):
    pdf_io = io.BytesIO()
    c = canvas.Canvas(pdf_io, pagesize=A4)
    width, height = A4
    current_font = 'CustomFont' if HAS_FONT else 'Helvetica'
    
    # 決定是 5 科還是 7 科，並統一歸納成 5 大科供繪圖用
    has_7_subjects = all(col in df.columns for col in ['歷史', '地理', '公民'])
    core_subjects = ['國文', '英文', '數學', '自然', '社會']
    
    # 計算各科班級平均 (用於文字顯示)
    avg_dict = {}
    if has_7_subjects:
        for sub in ['國文', '英文', '數學', '自然', '歷史', '地理', '公民']:
            avg_dict[sub] = df[sub].mean()
        avg_dict['社會'] = (avg_dict['歷史'] + avg_dict['地理'] + avg_dict['公民']) / 3
    else:
        for sub in core_subjects:
            avg_dict[sub] = df[sub].mean()

    for _, row in df.iterrows():
        seat, name = row.get('座號', ''), str(row.get('姓名', ''))
        
        # 準備 PDF 文字內容
        if has_7_subjects:
            print_text = [
                f"國文: {row['國文']:.1f}  (班均: {avg_dict['國文']:.1f})", 
                f"英文: {row['英文']:.1f}  (班均: {avg_dict['英文']:.1f})", 
                f"數學: {row['數學']:.1f}  (班均: {avg_dict['數學']:.1f})", 
                f"自然: {row['自然']:.1f}  (班均: {avg_dict['自然']:.1f})",
                f"社會: {row['社會']:.1f}  (班均: {avg_dict['社會']:.1f})",
                f"  └ 歷:{row['歷史']:.1f} / 地:{row['地理']:.1f} / 公:{row['公民']:.1f}"
            ]
        else:
            print_text = [f"{s}: {row[s]:.1f}  (班均: {avg_dict[s]:.1f})" for s in core_subjects]

        # 取得該學生的各科 PR 值供雷達圖繪製
        pr_scores = [row[f'{s}_PR'] for s in core_subjects]

        # --- PDF 版面設計 ---
        c.setFont(current_font, 24)
        c.drawCentredString(width/2, height - 70, "學 生 個 人 成 績 單")
        
        c.setFont(current_font, 14)
        c.drawString(60, height - 120, f"座號: {int(seat)}      姓名: {name}")
        c.line(60, height - 130, width - 60, height - 130)
        
        # 列印各科與平均
        y_pos = height - 170
        c.setFont(current_font, 12)
        for text in print_text:
            c.drawString(70, y_pos, text)
            y_pos -= 25
            
        y_pos -= 10
        c.setFont(current_font, 14)
        c.drawString(70, y_pos, f"⭐ 五科總分: {row['總分']:.1f}")
        c.drawString(70, y_pos - 30, f"⭐ 班級名次: 第 {int(row['名次'])} 名")
        c.drawString(70, y_pos - 60, f"⭐ 總分 PR 值: {row['總PR']:.1f}")

        # 右側 PR 雷達圖
        chart_img = create_pr_radar_chart(core_subjects, pr_scores)
        c.drawImage(ImageReader(chart_img), width - 330, height - 440, width=300, height=300, mask='auto')
        
        c.showPage()
        
    c.save()
    pdf_io.seek(0)
    return pdf_io

# ==========================================
# 3. 介面操作區
# ==========================================
default_url = "https://docs.google.com/spreadsheets/d/1lp0F45BnLO0Hn2l47vJ7orawr0KfIaT_/edit?gid=1366647975#gid=1366647975"
sheet_url = st.text_input("🔗 成績試算表網址：", value=default_url)

if sheet_url:
    csv_url = get_google_sheet_csv_url(sheet_url)
    if csv_url:
        try:
            with st.spinner("正在連線讀取並計算 PR 值與名次..."):
                df = pd.read_csv(csv_url)
                
                # 過濾資料
                df['座號'] = pd.to_numeric(df['座號'], errors='coerce')
                df = df.dropna(subset=['座號'])
                df = df[df['座號'] > 0]
                
                # 分數轉數字，填補 NaN 為 0
                cols_to_convert = ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']
                for col in cols_to_convert:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
                # 計算社會科總結 (如果是 7 科)
                if all(c in df.columns for c in ['歷史', '地理', '公民']):
                    df['社會'] = (df['歷史'] + df['地理'] + df['公民']) / 3
                elif '社會' not in df.columns:
                    df['社會'] = 0 # 防呆

                # 核心大變動：計算各科 PR 值！
                core_subjects = ['國文', '英文', '數學', '自然', '社會']
                for sub in core_subjects:
                    df[f'{sub}_PR'] = df[sub].rank(pct=True) * 100
                
                # 計算總分、總PR與名次
                df['總分'] = df[core_subjects].sum(axis=1)
                df['總PR'] = df['總分'].rank(pct=True) * 100
                df['名次'] = df['總分'].rank(ascending=False, method='min') # 同分則同名次
                
                df = df.sort_values(by='座號').reset_index(drop=True)
                student_count = len(df)
                
            st.success(f"✅ 成功鎖定！系統自動偵測到本班共 **{student_count}** 位學生，各科 PR 與名次計算完成。")
            
            # 預覽包含名次與 PR 的資料
            preview_cols = ['座號', '姓名', '總分', '名次', '總PR']
            st.dataframe(df[preview_cols].head(5))
            
            if not HAS_FONT:
                st.error("⚠️ 系統未能成功載入中文字體，產出的 PDF 將無法正常顯示中文。")
            else:
                if st.button(f"🚀 一鍵產生 {student_count} 人 PDF 成績單 (PR版)", type="primary"):
                    with st.spinner("AI 正在繪製 PR 雷達地圖與排版 PDF，這會花幾秒鐘的時間..."):
                        pdf_data = generate_pdf_report(df)
                        
                    st.balloons()
                    st.download_button(
                        label=f"📥 下載全班 ({student_count}人) PDF 成績單",
                        data=pdf_data,
                        file_name="全班個人成績單_PR雷達版.pdf",
                        mime="application/pdf"
                    )
        except Exception as e:
            st.error("❌ 讀取失敗！請確認試算表格式或欄位名稱。")
            st.warning(f"錯誤細節：{e}")
