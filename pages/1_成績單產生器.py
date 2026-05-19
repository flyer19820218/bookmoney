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
import os
from matplotlib import font_manager
import matplotlib
matplotlib.use('Agg')

st.set_page_config(page_title="AI 成績單產生器", layout="centered", page_icon="📈")

FONT_NAME = "NotoSansTC-Regular.ttf"
FONT_URL = "https://cdn.jsdelivr.net/gh/themoeway/noto-sans-tc-ttf@master/ttf/NotoSansTC-Regular.ttf"

@st.cache_resource
def init_fonts():
    if not os.path.exists(FONT_NAME):
        try:
            with st.spinner("未偵測到本地字體，正在從網路下載..."):
                response = requests.get(FONT_URL, headers={'User-Agent': 'Mozilla/5.0'}, timeout=60)
                with open(FONT_NAME, "wb") as f:
                    f.write(response.content)
        except Exception as e:
            return False
    try:
        font_manager.fontManager.addfont(FONT_NAME)
        dynamic_font_name = font_manager.FontProperties(fname=FONT_NAME).get_name()
        plt.rcParams['font.sans-serif'] = [dynamic_font_name, 'Microsoft JhengHei', 'Arial']
        plt.rcParams['axes.unicode_minus'] = False
        pdfmetrics.registerFont(TTFont('CustomFont', FONT_NAME))
        return True
    except:
        return False

HAS_FONT = init_fonts()

def get_google_sheet_csv_url(url):
    try:
        if "export?format=csv" in url: return url
        import re
        match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
        if match:
            gid_match = re.search(r'gid=([0-9]+)', url)
            gid_param = f"&gid={gid_match.group(1)}" if gid_match else ""
            return f"https://docs.google.com/spreadsheets/d/{match.group(1)}/export?format=csv{gid_param}"
    except: pass
    return None

def create_pr_radar_chart(labels, pr_scores):
    num_vars = len(labels)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    plot_scores = list(pr_scores) + [pr_scores[0]]
    angles = angles + [angles[0]]
    
    fig, ax = plt.subplots(figsize=(4.5, 4.5), subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    
    display_labels = [f"{label}\n(PR)" for label in labels]
    ax.set_thetagrids(np.degrees(angles[:-1]), display_labels, fontsize=10)
    ax.set_ylim(0, 100)
    ax.set_yticks([25, 50, 75, 100])
    ax.set_yticklabels(["25", "50", "75", "99"], color="grey", size=8)
    
    ax.plot(angles, [50]*len(angles), color='#C0504D', linewidth=1.5, linestyle='--', label='PR 50 (中位數)')
    ax.plot(angles, plot_scores, color='#4F81BD', linewidth=2, linestyle='solid', label='個人優勢')
    ax.fill(angles, plot_scores, color='#4F81BD', alpha=0.35)
    plt.legend(loc='upper right', bbox_to_anchor=(1.35, 1.15), fontsize=9)
    
    img_io = io.BytesIO()
    plt.savefig(img_io, format='png', bbox_inches='tight', transparent=True)
    plt.close(fig)
    img_io.seek(0)
    return img_io

def generate_pdf_report(df_students, df_stats, subjects, has_7_subjects):
    pdf_io = io.BytesIO()
    c = canvas.Canvas(pdf_io, pagesize=A4)
    width, height = A4
    current_font = 'CustomFont' if HAS_FONT else 'Helvetica'
    
    def fmt_2f(val):
        try:
            return f"{float(val):.2f}"
        except:
            return str(val)

    for _, row in df_students.iterrows():
        c.setFont(current_font, 24)
        c.drawCentredString(width/2, height - 70, "學 生 個 人 成 績 單")
        
        c.setFont(current_font, 14)
        c.drawString(60, height - 120, f"座號: {int(row['座號'])}      姓名: {row['姓名']}")
        c.line(60, height - 130, width - 60, height - 130)
        
        y_pos = height - 165
        c.setFont(current_font, 11)
        
        if has_7_subjects:
            display_subs = ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']
        else:
            display_subs = ['國文', '英文', '數學', '自然', '社會']

        for s in display_subs:
            high = df_stats.get(s, {}).get('高標', '')
            avg = df_stats.get(s, {}).get('平均', '')
            
            stat_str = f"  (均: {fmt_2f(avg)} | 高: {fmt_2f(high)})" if str(avg) != 'nan' and avg != '' else ""
            
            if s in ['歷史', '地理', '公民']:
                c.drawString(65, y_pos, f"  └ {s}: {row.get(s, 0):.2f}{stat_str}")
            else:
                c.drawString(65, y_pos, f"{s}: {row.get(s, 0):.2f}{stat_str}")
            y_pos -= 25
            
        y_pos -= 10
        c.setFont(current_font, 13)
        c.drawString(65, y_pos, f"五科總分 :  {row['總分']:.2f}")
        c.drawString(65, y_pos - 25, f"班級名次 :  第 {int(row['名次'])} 名")
        c.drawString(65, y_pos - 50, f"總分 PR 值 :  {row['總PR']:.2f}")

        # 雷達圖
        radar_subs = ['國文', '英文', '數學', '自然', '歷史', '地理', '公民'] if has_7_subjects else ['國文', '英文', '數學', '自然', '社會']
        pr_scores = [row[f'{s}_PR'] for s in radar_subs]
        chart_img = create_pr_radar_chart(radar_subs, pr_scores)
        c.drawImage(ImageReader(chart_img), width - 350, height - 440, width=320, height=320, mask='auto')
        
        # 導師的話
        msg_y = height - 480
        c.setFont(current_font, 12)
        c.drawString(60, msg_y, "【導師勉勵】")
        c.setFont(current_font, 10)
        msg_lines = [
            "即將升上三年級，代表你們準備好迎接國中階段最重要的挑戰了！無論這次表現",
            "如何，它都只是當下的標記，絕不代表你的最終極限。未來的這一年，只要靜",
            "下心來，找到適合的讀書節奏，確實訂定目標，每一次的努力與修正都會讓你",
            "變得更強大。老師相信你們都有無限的潛力，讓我們一起迎接精彩的國三生活！"
        ]
        msg_y -= 20
        for line in msg_lines:
            c.drawString(65, msg_y, line)
            msg_y -= 18

        # 反省區
        box_y = msg_y - 150 
        c.setFont(current_font, 12)
        c.drawString(60, box_y + 125, "【自我反省與下階段目標】") 
        c.setStrokeColorRGB(0.5, 0.5, 0.5)
        c.roundRect(60, box_y, width - 120, 115, 8, stroke=1, fill=0)
        
        c.showPage()
    c.save()
    pdf_io.seek(0)
    return pdf_io

# ==========================================
# 網頁 UI 介面設定
# ==========================================
st.title("📈 班級成績單自動產生器")

# === 【新增】網頁使用說明與公版連結 ===
st.markdown("---")
st.markdown("### 📝 使用教學")
st.info("""
歡迎使用成績單產生系統！請依照以下 **3 個步驟** 操作：
1. **取得公版檔案**：請點擊下方連結，複製一份專屬的試算表到您的雲端硬碟。
   👉 [**點我取得公版成績單 (Google Sheets)**](https://reurl.cc/9Wqr6v)
2. **輸入學生成績**：在您複製的表格中填入班級成績，請保留表格底部的「高標」與「平均」列，系統會自動抓取這些數值。
3. **貼上連結產出**：將您填好的試算表網址貼在下方框框，點擊按鈕即可一鍵產出全班 PDF！
   *(💡 記得將試算表共用權限設為「**知道連結的人均可檢視**」喔！)*
""")
st.markdown("---")
# =====================================

sheet_url = st.text_input("🔗 請在此貼上您的成績試算表網址：", placeholder="https://docs.google.com/spreadsheets/d/...")

if sheet_url:
    csv_url = get_google_sheet_csv_url(sheet_url)
    if csv_url:
        try:
            with st.spinner("正在讀取您的原始數據..."):
                raw_df = pd.read_csv(csv_url)
                
                raw_df['_is_student'] = pd.to_numeric(raw_df['座號'], errors='coerce').notna()
                
                df_students = raw_df[raw_df['_is_student']].copy()
                df_students['座號'] = pd.to_numeric(df_students['座號'])
                df_students = df_students[df_students['座號'] > 0]
                
                df_bottom = raw_df[~raw_df['_is_student']].copy()
                stats_dict = {}
                cols_to_extract = ['國文', '英文', '數學', '自然', '社會', '歷史', '地理', '公民']
                
                for _, row in df_bottom.iterrows():
                    row_name = str(row['座號']).strip() 
                    if '高標' in row_name or '高標' in str(row['姓名']):
                        for c in cols_to_extract:
                            if c in row:
                                if c not in stats_dict: stats_dict[c] = {}
                                stats_dict[c]['高標'] = row[c]
                    if '平均' in row_name or '均標' in row_name or '平均' in str(row['姓名']):
                        for c in cols_to_extract:
                            if c in row:
                                if c not in stats_dict: stats_dict[c] = {}
                                stats_dict[c]['平均'] = row[c]

                for col in cols_to_extract:
                    if col in df_students.columns:
                        df_students[col] = pd.to_numeric(df_students[col], errors='coerce').fillna(0)
                
                has_7_subjects = all(c in df_students.columns for c in ['歷史', '地理', '公民'])
                
                if has_7_subjects:
                    radar_subjects = ['國文', '英文', '數學', '自然', '歷史', '地理', '公民']
                else:
                    radar_subjects = ['國文', '英文', '數學', '自然', '社會']
                    if '社會' not in df_students.columns: df_students['社會'] = 0

                if '總分' not in df_students.columns:
                    if has_7_subjects:
                        df_students['總分'] = df_students['國文'] + df_students['英文'] + df_students['數學'] + df_students['自然'] + (df_students['歷史'] + df_students['地理'] + df_students['公民']) / 3
                    else:
                        df_students['總分'] = df_students['國文'] + df_students['英文'] + df_students['數學'] + df_students['自然'] + df_students['社會']
                
                for sub in radar_subjects:
                    df_students[f'{sub}_PR'] = df_students[sub].rank(pct=True) * 100
                
                df_students['總PR'] = df_students['總分'].rank(pct=True) * 100
                if '名次' not in df_students.columns:
                    df_students['名次'] = df_students['總分'].rank(ascending=False, method='min') 
                
                df_students = df_students.sort_values(by='座號').reset_index(drop=True)
                student_count = len(df_students)
                
            st.success(f"✅ 成功鎖定！已讀取 {student_count} 位學生資料與各科高標/平均值。")
            
            if st.button("🚀 產生全班 PDF 成績單", type="primary"):
                with st.spinner("完美排版中，請稍候..."):
                    pdf_data = generate_pdf_report(df_students, stats_dict, radar_subjects, has_7_subjects)
                    
                st.balloons()
                st.download_button("📥 下載全班成績單 (PDF)", pdf_data, "成績單_完美版.pdf", "application/pdf")
        except Exception as e:
            st.error(f"❌ 發生錯誤，請確認您的試算表欄位名稱是否與公版一致：{e}")
