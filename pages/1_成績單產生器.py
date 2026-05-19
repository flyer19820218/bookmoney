# ==========================================
# === 區塊 1: 模組與初始化 ===
# ==========================================
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

# ==========================================
# === 區塊 2: 圖表產生器 ===
# ==========================================
def create_pr_radar_chart(labels, pr_scores):
    num_vars = len(labels)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    plot_scores = list(pr_scores) + [pr_scores[0]]
    angles = angles + [angles[0]]
    
    fig, ax = plt.subplots(figsize=(5, 5), subplot_kw=dict(polar=True))
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    
    display_labels = [f"{label}\n(PR)" for label in labels]
    ax.set_thetagrids(np.degrees(angles[:-1]), display_labels, fontsize=10)
    ax.set_ylim(0, 100)
    ax.set_yticks([25, 50, 75, 100])
    ax.set_yticklabels(["25", "50", "75", "100"], color="grey", size=8)
    
    ax.plot(angles, [50]*len(angles), color='#C0504D', linewidth=1.5, linestyle='--', label='PR 50 (中位數)')
    ax.plot(angles, plot_scores, color='#4F81BD', linewidth=2, linestyle='solid', label='個人優勢')
    ax.fill(angles, plot_scores, color='#4F81BD', alpha=0.35)
    plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.1), fontsize=9, ncol=2)
    
    img_io = io.BytesIO()
    plt.savefig(img_io, format='png', bbox_inches='tight', transparent=True, dpi=150)
    plt.close(fig)
    img_io.seek(0)
    return img_io

# ==========================================
# === 區塊 3: PDF 產生器 (排版核心) ===
# ==========================================
def generate_pdf_report(df_students, df_stats, subjects, has_7_subjects, selected_quote):
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
        
        c.drawImage(ImageReader(chart_img), width - 360, height - 450, width=320, height=320, preserveAspectRatio=True, mask='auto')
        
        # --- 導師的話 ---
        msg_y = height - 480
        c.setFont(current_font, 12)
        c.drawString(60, msg_y, "【導師勉勵】")
        
        c.setFont(current_font, 11) 
        
        msg_lines = []
        for i in range(0, len(selected_quote), 33):
            msg_lines.append(selected_quote[i:i+33])

        msg_y -= 20
        for line in msg_lines:
            c.drawString(65, msg_y, line)
            msg_y -= 20 

        # --- 反省區 (往下平移 40 像素) ---
        c.setFont(current_font, 12)
        c.drawString(60, height - 640, "【自我反省與下階段目標】") 
        c.setStrokeColorRGB(0.5, 0.5, 0.5)
        c.roundRect(60, height - 770, width - 120, 115, 8, stroke=1, fill=0)
        
        c.showPage()
    c.save()
    pdf_io.seek(0)
    return pdf_io

# ==========================================
# === 區塊 4: 網頁 UI 與主程式 ===
# ==========================================
quotes_list = [
    "✏️ 我想自己寫...", 
    "每次的考試都是檢視自己學習歷程的絕佳機會。無論結果如何，這都只是學習旅程中的一個標記。請保持求知若渴的心，繼續穩紮穩打，未來的你一定會感謝現在努力的自己！",
    "分數只是數字，更重要的是你從中學到了什麼。找出自己的弱點並勇敢面對它，就是進步的開始。老師相信你的潛力無限，下個階段我們一起設定新目標，繼續前進！",
    "「不怕慢，只怕站。」學習就像跑馬拉松，重點不是瞬間的爆發力，而是持續不懈的毅力。調整好步伐，堅持每天進步一點點，最後的勝利一定屬於你。",
    "看到你在這段期間的努力，老師感到非常欣慰。也許成果還沒有完全展現，但所有的汗水都不會白流。請繼續保持這份熱忱與堅持，閃耀的時刻就在不遠處等著你！",
    "學習的路上難免會遇到挫折，但挫折是為了讓你變得更堅強。不要因為一次的失利而氣餒，把錯誤當作墊腳石，勇敢地跨越它，相信下一次的你一定會更加出色。",
    "優秀是一種習慣，而你正在慢慢培養這種習慣。繼續保持你良好的學習態度，不要害怕發問，不要害怕挑戰困難，你的努力終將為你帶來豐碩的果實。繼續加油！",
    "每一次的努力都是在為未來打地基。現在的辛苦，是為了讓未來的自己有更多的選擇權。相信自己的能力，勇敢迎接接下來的每一個挑戰，老師會一直在背後支持你。",
    "成功沒有捷徑，只有一步一腳印的踏實。請認真檢視這次的成績，找到需要補強的地方，並且確實執行你的讀書計畫。只要你願意付出，一定能看到改變的發生。",
    "你的進步老師都看在眼裡！這份成績單是你努力的證明，請為自己感到驕傲。但別忘了，這只是一個新的起點，繼續保持渴望學習的心，去探索更廣闊的知識領域吧！",
    "不要和別人比，只要今天的你比昨天的你進步，這就是最大的成功。每個人都有自己的學習步調，找到適合自己的方法最重要。對自己有信心，你絕對做得到！"
]

st.title("📈 班級成績單自動產生器 (公版)")

st.markdown("---")
st.markdown("### 📝 使用教學 (請務必詳讀)")
st.info("""
歡迎使用成績單產生系統！請依照以下 **3 個步驟** 產出您班上的專屬成績單：

**步驟 1：建立專屬的成績單副本**
這是公版的成績單，**請勿直接在上面修改！**
👉 請點擊：[**【點我開啟公版成績單】**](https://reurl.cc/9Wqr6v)
進入後，點擊左上角的 **「檔案」 ➜ 「建立副本」**，將它存到您自己的 Google 雲端硬碟中。

**步驟 2：輸入您班上的學生成績**
打開您剛剛建立的 **「副本」**，填入您班上的學生成績。
*💡 貼心提醒：請保留表格最底部的「高標」與「平均」列，不要刪除，系統會自動幫您抓取這些數值！*

**步驟 3：貼上連結，一鍵產出 PDF！**
確認成績無誤後，請將您的試算表共用權限設為「**知道連結的人均可檢視**」。然後將網址複製，貼到下方的框框中，點擊產出按鈕即可！
""")
st.markdown("---")

sheet_url = st.text_input("🔗 請在此貼上「您自己的成績試算表」網址：", placeholder="https://docs.google.com/spreadsheets/d/...")

st.markdown("### 💬 選擇要印在成績單上的導師勉勵")
selected_quote = st.selectbox("請選擇一段符合班級現狀的鼓勵語：", quotes_list)

if selected_quote == "✏️ 我想自己寫...":
    st.warning("⚠️ **請注意：為了不與下方的學生反省區重疊，請務必將字數控制在 120 字以內！**")
    custom_quote = st.text_area("請在此輸入您的自訂勉勵語：", max_chars=120)
    final_quote = custom_quote
else:
    final_quote = selected_quote

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
                if final_quote.strip() == "":
                    st.warning("⚠️ 請記得選擇或輸入導師勉勵語喔！")
                else:
                    with st.spinner("完美排版中，請稍候..."):
                        pdf_data = generate_pdf_report(df_students, stats_dict, radar_subjects, has_7_subjects, final_quote)
                        
                    st.balloons()
                    st.download_button("📥 下載全班成績單 (PDF)", pdf_data, "成績單_完美版.pdf", "application/pdf")
        except Exception as e:
            st.error(f"❌ 發生錯誤，請確認您的試算表欄位名稱是否與公版一致：{e}")
