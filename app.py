import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.worksheet.page import PageMargins
import re
import io
import requests
import os

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

# 🌟 新增：ReportLab PDF 核心繪圖套件
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader

# ==========================================
# 0. 網頁基本設定與自動字型下載
# ==========================================
st.set_page_config(page_title="全校通用購書單系統", layout="centered", page_icon="📚")

FONT_NAME = "NotoSansTC-Regular.ttf"
FONT_URL = "https://raw.githubusercontent.com/google/fonts/main/ofl/notosanstc/NotoSansTC%5Bwght%5D.ttf"

@st.cache_resource
def init_fonts():
    """確保 Streamlit Cloud 上有中文字型可用，避免 PDF 亂碼"""
    if not os.path.exists(FONT_NAME):
        try:
            with st.spinner("正在下載微軟正黑體/Noto中文字型以支援 PDF 產出..."):
                response = requests.get(FONT_URL, headers={'User-Agent': 'Mozilla/5.0'}, timeout=60)
                with open(FONT_NAME, "wb") as f:
                    f.write(response.content)
        except Exception as e:
            st.error(f"字型下載失敗：{e}")
            return False
    try:
        pdfmetrics.registerFont(TTFont('CustomFont', FONT_NAME))
        return True
    except Exception as e:
        st.error(f"字型註冊失敗：{e}")
        return False

HAS_FONT = init_fonts()

# ==========================================
# 0.5 共用工具：個別學生免購排除與即時書目預覽
# ==========================================
def normalize_seat(val):
    text = str(val).strip()
    if text in ["", "nan", "None"]:
        return ""
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".")[0]
    return text

def safe_price(val):
    num = pd.to_numeric(val, errors="coerce")
    if pd.isna(num):
        return 0
    return int(round(float(num))) if abs(float(num) - round(float(num))) < 1e-6 else float(num)

def attach_student_keys(df_students, default_group="無"):
    df_out = df_students.copy()
    occ = {}
    keys = []
    records = []
    for row_idx, row in df_out.iterrows():
        seat = normalize_seat(row.get("座號", ""))
        name = str(row.get("姓名", "")).strip()
        base = f"{seat}|{name}"
        occ[base] = occ.get(base, 0) + 1
        s_key = f"{base}#{occ[base]}"
        keys.append(s_key)

        gifted = str(row.get("資優類別", default_group)).strip()
        eng = str(row.get("英組", default_group)).strip()
        math = str(row.get("數組", default_group)).strip()
        sci = str(row.get("自組", default_group)).strip()
        label = f"{seat}號 {name}" if seat else name
        records.append({
            "key": s_key,
            "row_idx": row_idx,
            "seat": seat,
            "name": name,
            "label": label,
            "gifted": gifted,
            "eng": eng,
            "math": math,
            "sci": sci,
        })
    df_out["_student_key"] = keys
    return df_out, records

def attach_book_keys(df_books, mode="tab2"):
    df_out = df_books.copy()
    occ = {}
    keys = []
    records = []
    for row_idx, book in df_out.iterrows():
        if mode == "tab1":
            b_name = str(book.get("商品名稱", "")).strip()
            b_subj = str(book.get("科目", "")).strip()
            b_code = str(book.get("分組代號", "1")).strip()
            b_price = safe_price(book.get("單價", 0))
            b_pub = ""
        else:
            b_name = str(book.get("name", "")).strip()
            b_subj = str(book.get("subj", "")).strip()
            b_code = str(book.get("code", "1")).strip()
            b_price = safe_price(book.get("price", 0))
            b_pub = str(book.get("publisher", "")).strip()

        b_subj_check = "社會" if b_subj in ["歷", "地", "公", "歷史", "地理", "公民", "社會三科"] else b_subj
        base = f"{b_pub}|{b_subj_check}|{b_code}|{b_name}|{b_price}"
        occ[base] = occ.get(base, 0) + 1
        b_key = f"{base}#{occ[base]}"
        keys.append(b_key)

        records.append({
            "key": b_key,
            "row_idx": row_idx,
            "name": b_name,
            "subj": b_subj,
            "subj_check": b_subj_check,
            "code": b_code,
            "price": b_price,
            "publisher": b_pub,
        })
    df_out["_book_key"] = keys
    return df_out, records

def format_student_group_summary(student):
    parts = []
    eng = student.get("eng", "")
    math = student.get("math", "")
    sci = student.get("sci", "")
    gifted = student.get("gifted", "")
    if eng not in ["", "1", "無", "免", "nan"]:
        parts.append(f"英:{eng}")
    if math not in ["", "1", "無", "免", "nan"]:
        parts.append(f"數:{math}")
    if sci not in ["", "1", "無", "免", "nan"] and sci != math:
        parts.append(f"自:{sci}")
    if gifted not in ["", "1", "無", "免", "nan"]:
        parts.append(f"資優:{gifted}")
    return " / ".join(parts) if parts else "一般"

def compute_student_book_plan(student_records, book_records, eligibility_fn, excluded_pairs):
    excluded_pairs = excluded_pairs or set()
    student_buying = {s["key"]: [] for s in student_records}
    student_excluded = {s["key"]: [] for s in student_records}
    student_eligible = {s["key"]: [] for s in student_records}
    book_buying = {b["key"]: [] for b in book_records}
    book_excluded = {b["key"]: [] for b in book_records}
    book_eligible = {b["key"]: [] for b in book_records}

    for student in student_records:
        s_key = student["key"]
        for book in book_records:
            b_key = book["key"]
            if eligibility_fn(student, book):
                student_eligible[s_key].append(book)
                book_eligible[b_key].append(student)
                if (s_key, b_key) in excluded_pairs:
                    student_excluded[s_key].append(book)
                    book_excluded[b_key].append(student)
                else:
                    student_buying[s_key].append(book)
                    book_buying[b_key].append(student)

    return {
        "student_buying": student_buying,
        "student_excluded": student_excluded,
        "student_eligible": student_eligible,
        "book_buying": book_buying,
        "book_excluded": book_excluded,
        "book_eligible": book_eligible,
    }

def wrap_pdf_text(text, font_name, font_size, max_width, max_lines=3):
    if not text:
        return []
    lines = []
    current = ""
    for ch in str(text):
        candidate = current + ch
        if current and pdfmetrics.stringWidth(candidate, font_name, font_size) > max_width:
            lines.append(current)
            current = ch
        else:
            current = candidate
    if current:
        lines.append(current)
    if len(lines) > max_lines:
        trimmed = lines[max_lines - 1]
        while trimmed and pdfmetrics.stringWidth(trimmed + "…", font_name, font_size) > max_width:
            trimmed = trimmed[:-1]
        lines = lines[:max_lines - 1] + [trimmed + "…"]
    return lines

def write_individual_booklist_sheet(wb, student_records, plan):
    ws = wb.create_sheet("個別學生書目明細")
    headers = ["座號", "姓名", "分組／資優", "購買冊數", "購買書目明細", "已排除書目(免購)", "應收總額"]
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF'),
    )
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = Font(name="微軟正黑體", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    grand_total = 0
    for r_idx, s in enumerate(student_records, 2):
        s_key = s["key"]
        buying = plan["student_buying"][s_key]
        excluded = plan["student_excluded"][s_key]
        buying_str = "、".join([f"{b['name']}(${int(b['price'])})" for b in buying]) if buying else "免購"
        excluded_str = "、".join([f"{b['name']}(${int(b['price'])})" for b in excluded]) if excluded else ""
        subtotal = sum(b["price"] for b in buying)
        grand_total += subtotal

        row_vals = [
            s["seat"],
            s["name"],
            format_student_group_summary(s),
            len(buying),
            buying_str,
            excluded_str,
            subtotal,
        ]
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = Font(name="微軟正黑體", size=10, bold=(c_idx == 7))
            cell.border = thin_border
            if c_idx in [5, 6]:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    last_row = len(student_records) + 2
    ws.cell(row=last_row, column=6, value="班級總計").font = Font(name="微軟正黑體", bold=True)
    ws.cell(row=last_row, column=6).alignment = Alignment(horizontal="right", vertical="center")
    tot_cell = ws.cell(row=last_row, column=7, value=grand_total)
    tot_cell.font = Font(name="微軟正黑體", bold=True, color="FF0000")
    tot_cell.alignment = Alignment(horizontal="center", vertical="center")
    tot_cell.border = thin_border

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 52
    ws.column_dimensions['F'].width = 30
    ws.column_dimensions['G'].width = 12

def _trigger_rerun():
    if hasattr(st, "rerun"):
        st.rerun()
    elif hasattr(st, "experimental_rerun"):
        st.experimental_rerun()

def _on_bulk_exclusions(state_key, ver_key, pairs, exclude=True):
    exclusions = set(st.session_state.get(state_key, set()))
    if exclude:
        exclusions.update(pairs)
    else:
        exclusions.difference_update(pairs)
    st.session_state[state_key] = exclusions
    st.session_state[ver_key] = st.session_state.get(ver_key, 0) + 1

def _on_clear_exclusions(state_key, ver_key):
    st.session_state[state_key] = set()
    st.session_state[ver_key] = st.session_state.get(ver_key, 0) + 1

def render_exclusion_and_preview_section(prefix, student_records, book_records, eligibility_fn, show_publisher=False):
    state_key = f"{prefix}_excluded_pairs"
    ver_key = f"{prefix}_excluded_ver"
    msg_key = f"{prefix}_excluded_msg"
    if state_key not in st.session_state:
        st.session_state[state_key] = set()
    if ver_key not in st.session_state:
        st.session_state[ver_key] = 0

    valid_s_keys = {s["key"] for s in student_records}
    valid_b_keys = {b["key"] for b in book_records}
    raw_exclusions = st.session_state[state_key]
    active_exclusions = {
        (sk, bk) for (sk, bk) in raw_exclusions
        if sk in valid_s_keys and bk in valid_b_keys
    }

    plan = compute_student_book_plan(student_records, book_records, eligibility_fn, active_exclusions)
    book_by_key = {b["key"]: b for b in book_records}
    student_by_key = {s["key"]: s for s in student_records}
    ver = st.session_state[ver_key]

    st.markdown("#### 🙋 個別學生「已自購／不購買」書籍勾選")
    st.caption("若有學生已自行購買某本書，請在下方選單中選擇書目（或選擇學生），一次勾選完所有不買的名單後，再按「✅ 勾選完畢，套用免購名單」。")

    if st.session_state.get(msg_key):
        st.success(st.session_state.pop(msg_key))

    if not student_records or not book_records:
        st.info("目前尚無完整的學生或書目資料可設定。")
        return active_exclusions, plan

    mode = st.radio(
        "選擇勾選方式：",
        options=["📘 依書目勾選（這本書誰不買）", "🧑‍🎓 依學生勾選（這位學生哪些書不買）"],
        horizontal=True,
        key=f"{prefix}_excl_mode",
    )

    if mode.startswith("📘"):
        book_keys = [b["key"] for b in book_records]
        book_labels = {}
        for b in book_records:
            bk = b["key"]
            elig_cnt = len(plan["book_eligible"][bk])
            buy_cnt = len(plan["book_buying"][bk])
            excl_cnt = len(plan["book_excluded"][bk])
            code_tag = f" [{b['code']}]" if b["code"] not in ["1", "", "全"] else ""
            excl_tag = f" ★已排除 {excl_cnt} 人" if excl_cnt > 0 else ""
            book_labels[bk] = f"【{b['subj']}】{b['name']}{code_tag}（${int(b['price'])}｜{buy_cnt}/{elig_cnt}人購{excl_tag}）"

        selected_b_key = st.selectbox(
            "請選擇書目：",
            options=book_keys,
            format_func=lambda k: book_labels[k],
            key=f"{prefix}_select_book_excl",
        )
        selected_book = book_by_key[selected_b_key]
        elig_students = plan["book_eligible"][selected_b_key]
        excl_students = plan["book_excluded"][selected_b_key]

        if not elig_students:
            st.info("此書目依目前分組／資優條件無對應購買學生。")
        else:
            buy_cnt = len(elig_students) - len(excl_students)
            deducted = len(excl_students) * selected_book["price"]
            st.write(
                f"單價 **${int(selected_book['price'])}** ｜ 原預計 **{len(elig_students)}** 人 ｜ "
                f"目前已套用不購買 **{len(excl_students)}** 人（扣除 **${int(deducted):,}**）｜ 實際結算 **{buy_cnt}** 人"
            )

            with st.form(key=f"{prefix}_form_bybook_{ver}_{selected_b_key}"):
                st.caption("👇 請勾選所有「已自購／不購買此書」的學生，全部勾好後按下方按鈕套用（勾選過程中畫面不會跳動）：")
                cols = st.columns(3)
                checked_map = {}
                for idx, s in enumerate(elig_students):
                    sk = s["key"]
                    chk_key = f"{prefix}_fb_{ver}_{selected_b_key}_{sk}"
                    is_excl = (sk, selected_b_key) in active_exclusions
                    with cols[idx % 3]:
                        checked_map[sk] = st.checkbox(
                            s["label"],
                            value=is_excl,
                            key=chk_key,
                        )

                submitted = st.form_submit_button(
                    "✅ 勾選完畢，套用免購名單",
                    type="primary",
                    use_container_width=True,
                )

            if submitted:
                updated = set(st.session_state.get(state_key, set()))
                selected_cnt = 0
                for sk, is_checked in checked_map.items():
                    if is_checked:
                        updated.add((sk, selected_b_key))
                        selected_cnt += 1
                    else:
                        updated.discard((sk, selected_b_key))
                st.session_state[state_key] = updated
                st.session_state[ver_key] = st.session_state.get(ver_key, 0) + 1
                deduct_sum = int(selected_cnt * selected_book["price"])
                st.session_state[msg_key] = (
                    f"✅ 已套用「{selected_book['name']}」免購設定：共 {selected_cnt} 位學生不購買（扣除 ${deduct_sum:,} 元）。"
                )
                _trigger_rerun()

            btn_col1, btn_col2, _ = st.columns([1.2, 1.2, 2])
            all_book_pairs = [(s["key"], selected_b_key) for s in elig_students]
            with btn_col1:
                st.button(
                    "勾選本書全部不買",
                    key=f"{prefix}_book_excl_all_{selected_b_key}",
                    on_click=_on_bulk_exclusions,
                    args=(state_key, ver_key, all_book_pairs, True),
                )
            with btn_col2:
                st.button(
                    "取消本書全部勾選",
                    key=f"{prefix}_book_clear_{selected_b_key}",
                    on_click=_on_bulk_exclusions,
                    args=(state_key, ver_key, all_book_pairs, False),
                )
    else:
        student_keys = [s["key"] for s in student_records]
        student_labels = {}
        for s in student_records:
            sk = s["key"]
            payable = sum(b["price"] for b in plan["student_buying"][sk])
            excl_cnt = len(plan["student_excluded"][sk])
            excl_tag = f"｜已排除 {excl_cnt} 本" if excl_cnt > 0 else ""
            student_labels[sk] = f"{s['label']}（實繳 ${int(payable):,}{excl_tag}）"

        selected_s_key = st.selectbox(
            "請選擇學生：",
            options=student_keys,
            format_func=lambda k: student_labels[k],
            key=f"{prefix}_select_student_excl",
        )
        selected_student = student_by_key[selected_s_key]
        elig_books = plan["student_eligible"][selected_s_key]
        excl_books = plan["student_excluded"][selected_s_key]
        if not elig_books:
            st.info("此學生目前無需購買任何書目。")
        else:
            orig_amt = sum(b["price"] for b in elig_books)
            deduct_amt = sum(b["price"] for b in excl_books)
            final_amt = orig_amt - deduct_amt
            st.write(
                f"原應繳 **${int(orig_amt):,}** ｜ 目前已套用排除 **{len(excl_books)}** 本（扣除 **${int(deduct_amt):,}**）｜ "
                f"實繳 **${int(final_amt):,}**"
            )

            with st.form(key=f"{prefix}_form_bystu_{ver}_{selected_s_key}"):
                st.caption("👇 請勾選這位學生所有「已自購／不購買」的書目，全部勾好後按下方按鈕套用（勾選過程中畫面不會跳動）：")
                checked_books = {}
                for b in elig_books:
                    bk = b["key"]
                    chk_key = f"{prefix}_fs_{ver}_{selected_s_key}_{bk}"
                    is_excl = (selected_s_key, bk) in active_exclusions
                    code_tag = f" [{b['code']}]" if b["code"] not in ["1", "", "全"] else ""
                    checked_books[bk] = st.checkbox(
                        f"【{b['subj']}】{b['name']}{code_tag}（${int(b['price'])}）",
                        value=is_excl,
                        key=chk_key,
                    )

                submitted_stu = st.form_submit_button(
                    "✅ 勾選完畢，套用免購書目",
                    type="primary",
                    use_container_width=True,
                )

            if submitted_stu:
                updated = set(st.session_state.get(state_key, set()))
                selected_cnt = 0
                deduct_sum = 0
                for bk, is_checked in checked_books.items():
                    if is_checked:
                        updated.add((selected_s_key, bk))
                        selected_cnt += 1
                        deduct_sum += int(book_by_key[bk]["price"])
                    else:
                        updated.discard((selected_s_key, bk))
                st.session_state[state_key] = updated
                st.session_state[ver_key] = st.session_state.get(ver_key, 0) + 1
                st.session_state[msg_key] = (
                    f"✅ 已套用「{selected_student['label']}」免購設定：共排除 {selected_cnt} 本（扣除 ${deduct_sum:,} 元）。"
                )
                _trigger_rerun()

            btn_col1, btn_col2, _ = st.columns([1.2, 1.2, 2])
            all_stu_pairs = [(selected_s_key, b["key"]) for b in elig_books]
            with btn_col1:
                st.button(
                    "此生全部書目不買",
                    key=f"{prefix}_stu_excl_all_{selected_s_key}",
                    on_click=_on_bulk_exclusions,
                    args=(state_key, ver_key, all_stu_pairs, True),
                )
            with btn_col2:
                st.button(
                    "取消此生全部勾選",
                    key=f"{prefix}_stu_clear_{selected_s_key}",
                    on_click=_on_bulk_exclusions,
                    args=(state_key, ver_key, all_stu_pairs, False),
                )

    # 已勾選免購總覽
    excl_summary_rows = []
    for s in student_records:
        for b in plan["student_excluded"][s["key"]]:
            excl_summary_rows.append({
                "座號": s["seat"],
                "姓名": s["name"],
                "已排除書目(不購買)": b["name"],
                "科目": b["subj"],
                "扣除金額": int(b["price"]),
            })

    if excl_summary_rows:
        total_deducted_all = sum(r["扣除金額"] for r in excl_summary_rows)
        with st.expander(
            f"📌 目前已勾選 {len(excl_summary_rows)} 筆免購項目（共扣除 ${total_deducted_all:,} 元）",
            expanded=True,
        ):
            st.dataframe(pd.DataFrame(excl_summary_rows), use_container_width=True, hide_index=True)
            st.button(
                "🗑️ 清除全部免購勾選",
                key=f"{prefix}_clear_all_excl",
                on_click=_on_clear_exclusions,
                args=(state_key, ver_key),
            )

    st.divider()
    st.markdown("#### 👀 個別學生名單與書目預覽（免列印直接查看）")

    total_payable_all = sum(
        sum(b["price"] for b in plan["student_buying"][s["key"]])
        for s in student_records
    )
    total_excluded_cnt = len(excl_summary_rows)
    total_excluded_amt = sum(r["扣除金額"] for r in excl_summary_rows)

    m1, m2, m3 = st.columns(3)
    m1.metric("班級學生人數", f"{len(student_records)} 人")
    m2.metric(
        "已排除自購免購",
        f"{total_excluded_cnt} 人次",
        delta=f"-${total_excluded_amt:,}" if total_excluded_amt > 0 else "無排除",
        delta_color="normal" if total_excluded_amt == 0 else "inverse",
    )
    m3.metric("全班應收總額", f"$ {int(total_payable_all):,}")

    preview_tab1, preview_tab2, preview_tab3 = st.tabs([
        "📋 全班個別書目總表",
        "🔍 單一學生明細檢視",
        "📚 各書目購買人數統計",
    ])

    with preview_tab1:
        f_col1, f_col2 = st.columns([2, 1.3])
        with f_col1:
            search_kw = st.text_input(
                "搜尋學生（輸入座號或姓名）",
                value="",
                placeholder="例如：03 或 王小明",
                key=f"{prefix}_preview_search",
            ).strip()
        with f_col2:
            only_excl = st.checkbox(
                "僅顯示有排除免購的學生",
                value=False,
                key=f"{prefix}_preview_only_excl",
            )

        roster_rows = []
        for s in student_records:
            sk = s["key"]
            buying = plan["student_buying"][sk]
            excluded = plan["student_excluded"][sk]
            if only_excl and not excluded:
                continue
            if search_kw and (search_kw not in s["seat"] and search_kw not in s["name"]):
                continue
            roster_rows.append({
                "座號": s["seat"],
                "姓名": s["name"],
                "分組／資優": format_student_group_summary(s),
                "購買冊數": len(buying),
                "購買書目明細": "、".join([f"{b['name']}(${int(b['price'])})" for b in buying]) if buying else "免購",
                "已排除書目(免購)": "、".join([f"{b['name']}(-${int(b['price'])})" for b in excluded]) if excluded else "無",
                "應繳總額": int(sum(b["price"] for b in buying)),
            })

        if roster_rows:
            st.dataframe(pd.DataFrame(roster_rows), use_container_width=True, hide_index=True)
        else:
            st.info("沒有符合篩選條件的學生。")

    with preview_tab2:
        inspect_keys = [s["key"] for s in student_records]
        inspect_labels = {}
        for s in student_records:
            sk = s["key"]
            amt = int(sum(b["price"] for b in plan["student_buying"][sk]))
            ex_cnt = len(plan["student_excluded"][sk])
            ex_tag = f"｜免購 {ex_cnt} 本" if ex_cnt > 0 else ""
            inspect_labels[sk] = f"{s['label']}（應繳 ${amt:,}{ex_tag}）"

        inspect_s_key = st.selectbox(
            "選擇學生查看個人完整書目：",
            options=inspect_keys,
            format_func=lambda k: inspect_labels[k],
            key=f"{prefix}_inspect_student",
        )
        stu = student_by_key[inspect_s_key]
        stu_buying = plan["student_buying"][inspect_s_key]
        stu_excluded = plan["student_excluded"][inspect_s_key]
        stu_total = int(sum(b["price"] for b in stu_buying))

        st.markdown(f"**{stu['label']}** ｜ 分組狀態：`{format_student_group_summary(stu)}`")
        if stu_buying:
            detail_rows = []
            for b in stu_buying:
                row_item = {
                    "科目": b["subj"],
                    "商品名稱": b["name"],
                    "分組代號": b["code"],
                    "單價": int(b["price"]),
                }
                if show_publisher:
                    row_item["書商／來源"] = b["publisher"]
                detail_rows.append(row_item)
            st.dataframe(pd.DataFrame(detail_rows), use_container_width=True, hide_index=True)
        else:
            st.info("此學生本學期無應購項目（免繳費）。")

        st.markdown(f"**💰 應繳總計：`$ {stu_total:,} 元`**（共 {len(stu_buying)} 項）")

        if stu_excluded:
            excl_txt = "、".join([f"{b['name']} (${int(b['price'])})" for b in stu_excluded])
            excl_sum = int(sum(b["price"] for b in stu_excluded))
            st.warning(f"🚫 已排除自購免購（不計費，共扣除 ${excl_sum:,} 元）：{excl_txt}")

    with preview_tab3:
        book_stat_rows = []
        for b in book_records:
            bk = b["key"]
            elig_cnt = len(plan["book_eligible"][bk])
            excl_list = plan["book_excluded"][bk]
            buy_cnt = len(plan["book_buying"][bk])
            row_item = {
                "商品名稱": b["name"],
                "科目": b["subj"],
                "分組代號": b["code"],
            }
            if show_publisher:
                row_item["書商／來源"] = b["publisher"]
            row_item.update({
                "原預計人數": elig_cnt,
                "免購人數": len(excl_list),
                "實際結算人數": buy_cnt,
                "單價": int(b["price"]),
                "小計金額": int(buy_cnt * b["price"]),
                "免購名單": "、".join([s["label"] for s in excl_list]) if excl_list else "無",
            })
            book_stat_rows.append(row_item)
        st.dataframe(pd.DataFrame(book_stat_rows), use_container_width=True, hide_index=True)

    return active_exclusions, plan

@st.cache_data(ttl=300, show_spinner=False)
def load_google_sheet_data(xlsx_url):
    df_students = pd.read_excel(xlsx_url, sheet_name=0).fillna("")
    df_books = pd.read_excel(xlsx_url, sheet_name=1).fillna("")
    return df_students, df_books

st.title("📚 班級各項費用與通知單系統")
st.markdown("請選擇上方分頁切換您要使用的功能！")

# 🌟 前兩頁為已驗證的書籍費系統；第三頁新增個人功課表。
tab1, tab2, tab3 = st.tabs([
    "📝 雲端試算表自動版",
    "🖨️ 懶人 PDF 產生器",
    "🗓️ 個人功課表",
])

# =====================================================================
# 🌟 分頁 1：原本的雲端自動系統 (保持不變)
# =====================================================================
with tab1:
    st.header("💡 導師專屬「一鍵排版」全攻略")

    with st.expander("👉 第一步：領取並填寫您的專屬表格 (點我展開)", expanded=True):
        st.markdown("""
        1. 點擊進入 [校內購書公版範本 (短網址：https://reurl.cc/K2LgNe)](https://reurl.cc/K2LgNe)。
        2. 進去後，點選左上角的 **「檔案」 > 「建立副本」**。（一定要建立副本才能編輯喔！）
        3. **不想自己打字？讓 AI 幫你精準整理！**
            手邊只有「紙本估價單」？直接拍照傳給 **Gemini 或 ChatGPT**，並**完整複製貼上以下這段指令給 AI**：
            
            > 你現在是一位專業、細心的「學校資料輸入員」。請將我上傳的書商估價單照片，精準地轉換成表格。
            > 
            > **【表格必須嚴格包含這 4 個欄位】：**
            > 1. **商品名稱**：完整保留書名或材料費名稱。
            > 2. **科目**：請根據書名自動分類為「國、英、數、自、社會、其他」這六類。（注意：歷史、地理、公民請一律歸類為「社會」；聯絡簿、桌墊、材料費請歸類為「其他」）。
            > 3. **分組代號**：若圖片中無特別標示分組，請一律填入數字「1」。如果有分組，請直接填寫代號（如 5A、6B）。
            > 4. **單價**：只填寫純數字，不要加上 $ 或 元。
            > 
            > **【嚴格限制】：** 請勿遺漏項目，不要加入任何問候語或廢話，你的回覆只能有「一個表格」。
            
            接著直接把 AI 做好的表格**全選複製、貼上**到您的 Google 試算表即可！
        """)

    with st.expander("👉 第二步：開啟「共用」權限 (這步沒做會失敗喔！)", expanded=True):
        st.markdown("""
        1. 點擊您 Google 試算表右上角大大的 **「共用」** 按鈕。
        2. 在「一般存取權」下方，將「限制」改為 **「知道連結的人即可檢視」**。
        3. 按下 **「複製連結」**。
        """)

    with st.expander("👉 第三步：貼上網址，領取列印檔", expanded=True):
        st.markdown("""
        1. 將剛剛複製的網址，**貼到下方輸入框**。
        2. 畫面顯示「讀取成功」後，按下 **「🚀 開始產生檔案」**。
        3. 您可以分別下載「給家長的 4格通知單」以及「給導師的 對帳收費總表」。
        """)

    st.warning("⚠️ **列印小叮嚀：** 下載家長通知單後，因各電腦 Excel 預設邊界不同，請在列印時自行點選「預覽列印」微調縮放比例！")
    st.info("💡 **資優生怎麼辦？** 在名單分頁的資優生欄位選「語資」或「數資」，系統會自動扣除。一年級沒分組，代號全填 `1` 就好！")
    st.divider()

    # 原本的核心邏輯函數
    @st.cache_data(ttl=600, show_spinner=False)
    def get_google_sheet_xlsx_url(url):
        try:
            if "docs.google.com" not in url:
                response = requests.head(url, allow_redirects=True, timeout=15)
                url = response.url
                
            pattern = r'https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9-_]+)'
            match = re.search(pattern, url)
            if match:
                sheet_id = match.group(1)
                return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
            return None
        except Exception as e:
            return None

    def should_buy_book(b_subj, b_code, s_gifted, s_eng, s_math, s_sci):
        s_gifted = str(s_gifted).strip()
        b_code = str(b_code).strip()
        
        if s_gifted == "語資" and b_subj in ["國", "英"]: return False
        if s_gifted == "數資" and b_subj in ["數", "自"]: return False
        if b_code in ["1", "全", "", "nan"]: return True
            
        if b_subj == "英" and str(s_eng).strip() in b_code: return True
        if b_subj == "數" and str(s_math).strip() in b_code: return True
        if b_subj == "自" and str(s_sci).strip() in b_code: return True
        return False

    def generate_receipts_excel(df_students, df_books, excluded_pairs=None):
        excluded_pairs = excluded_pairs or set()
        df_s_use = df_students if "_student_key" in df_students.columns else attach_student_keys(df_students, default_group="1")[0]
        df_b_use = df_books if "_book_key" in df_books.columns else attach_book_keys(df_books, mode="tab1")[0]

        wb = Workbook()
        ws = wb.active
        ws.title = "購書通知單(4張一頁)"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5)

        ws.column_dimensions['A'].width = 8   ; ws.column_dimensions['B'].width = 38
        ws.column_dimensions['C'].width = 7   ; ws.column_dimensions['D'].width = 2
        ws.column_dimensions['E'].width = 8   ; ws.column_dimensions['F'].width = 38
        ws.column_dimensions['G'].width = 7

        f_title = Font(name="微軟正黑體", size=13, bold=True)
        f_info = Font(name="微軟正黑體", size=12, bold=True)
        f_norm = Font(name="微軟正黑體", size=9)
        f_bold = Font(name="微軟正黑體", size=10, bold=True)
        f_tot = Font(name="微軟正黑體", size=12, bold=True, color="FF0000")
        al_c = Alignment(horizontal="center", vertical="center")
        al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
        al_r = Alignment(horizontal="right", vertical="center")
        thin = Side(style='thin', color='000000')
        b_all = Border(left=thin, right=thin, top=thin, bottom=thin)

        RECEIPT_ROWS = 12

        for i, (_, row) in enumerate(df_s_use.iterrows()):
            seat, name = normalize_seat(row.get("座號", "")), row.get("姓名", "")
            gifted = row.get("資優類別", "")
            eng, math, sci = row.get("英組", "1"), row.get("數組", "1"), row.get("自組", "1")
            s_key = row.get("_student_key", "")

            student_books = {"國": [], "英": [], "數": [], "自": [], "社會": [], "其他": []}
            for _, book in df_b_use.iterrows():
                b_name = book.get("商品名稱", "")
                b_subj = book.get("科目", "")
                b_code = book.get("分組代號", "1")
                b_price = safe_price(book.get("單價", 0))
                b_key = book.get("_book_key", "")
                if b_subj in ["歷", "地", "公", "歷史", "地理", "公民", "社會三科"]: b_subj = "社會"
                if should_buy_book(b_subj, b_code, gifted, eng, math, sci) and (s_key, b_key) not in excluded_pairs:
                    if b_subj in student_books: student_books[b_subj].append({"name": b_name, "price": b_price})
                    else: student_books["其他"].append({"name": b_name, "price": b_price})

            page = i // 4
            pos = i % 4
            start_row = page * (RECEIPT_ROWS * 2 + 2) + (0 if pos < 2 else RECEIPT_ROWS + 1) + 1
            start_col = 1 if pos % 2 == 0 else 5
            
            ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col+2)
            c1 = ws.cell(row=start_row, column=start_col, value="學期購書費通知單")
            c1.font = f_title ; c1.alignment = al_c
            
            ws.merge_cells(start_row=start_row+1, start_column=start_col, end_row=start_row+1, end_column=start_col+2)
            c2 = ws.cell(row=start_row+1, column=start_col, value=f"座號：{seat}      姓名：{name}")
            c2.font = f_info ; c2.alignment = Alignment(horizontal="left", vertical="center")
            
            for col_offset, text in enumerate(["科目", "購買明細", "小計"]):
                cell = ws.cell(row=start_row+2, column=start_col+col_offset, value=text)
                cell.font = f_bold ; cell.alignment = al_c ; cell.border = b_all

            cat_order = ["國", "英", "數", "自", "社會", "其他"]
            curr_row = start_row + 3
            grand_total = 0
            for cat in cat_order:
                items = student_books[cat]
                if not items:
                    det_str, subtotal = "免購", 0
                else:
                    det_str = "、".join([f"{item['name']}(${int(item['price'])})" for item in items])
                    subtotal = sum([item['price'] for item in items])
                grand_total += subtotal
                
                c_cat = ws.cell(row=curr_row, column=start_col, value="社會三科" if cat == "社會" else cat)
                c_cat.font = f_bold ; c_cat.alignment = al_c ; c_cat.border = b_all
                c_det = ws.cell(row=curr_row, column=start_col+1, value=det_str)
                c_det.font = f_norm ; c_det.alignment = al_l ; c_det.border = b_all
                c_sub = ws.cell(row=curr_row, column=start_col+2, value=subtotal)
                c_sub.font = f_bold ; c_sub.alignment = al_c ; c_sub.border = b_all
                ws.row_dimensions[curr_row].height = 28
                curr_row += 1
                
            ws.merge_cells(start_row=curr_row, start_column=start_col, end_row=curr_row, end_column=start_col+1)
            c_tot_l = ws.cell(row=curr_row, column=start_col, value="應收總計：")
            c_tot_l.font = f_tot ; c_tot_l.alignment = al_r ; c_tot_l.border = b_all
            ws.cell(row=curr_row, column=start_col+1).border = b_all
            c_tot_v = ws.cell(row=curr_row, column=start_col+2, value=f"{int(grand_total)}")
            c_tot_v.font = f_tot ; c_tot_v.alignment = al_c ; c_tot_v.border = b_all
            ws.row_dimensions[curr_row].height = 25
            
            for r in range(start_row, curr_row + 1):
                for c in range(start_col, start_col + 3): ws.cell(row=r, column=c).border = b_all

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    def generate_master_excel(df_students, df_books, excluded_pairs=None, student_records=None, plan=None):
        excluded_pairs = excluded_pairs or set()
        df_s_use = df_students if "_student_key" in df_students.columns else attach_student_keys(df_students, default_group="1")[0]
        df_b_use = df_books if "_book_key" in df_books.columns else attach_book_keys(df_books, mode="tab1")[0]

        wb = Workbook()
        ws = wb.active
        ws.title = "班級收費總表"
        
        book_rows = []
        for _, book in df_b_use.iterrows():
            b_name = book.get("商品名稱", "")
            b_subj = book.get("科目", "")
            b_code = book.get("分組代號", "1")
            b_price = safe_price(book.get("單價", 0))
            b_key = book.get("_book_key", "")
            b_subj_check = "社會" if b_subj in ["歷", "地", "公", "歷史", "地理", "公民", "社會三科"] else b_subj
            
            qty = 0
            for _, s in df_s_use.iterrows():
                s_key = s.get("_student_key", "")
                if should_buy_book(b_subj_check, b_code, s.get("資優類別", ""), s.get("英組", "1"), s.get("數組", "1"), s.get("自組", "1")) and (s_key, b_key) not in excluded_pairs:
                    qty += 1
            book_rows.append([b_name, b_subj, b_code, qty, b_price])

        student_rows = []
        has_any_excl = False
        for _, s in df_s_use.iterrows():
            seat, name = normalize_seat(s.get("座號", "")), s.get("姓名", "")
            gifted = s.get("資優類別", "")
            eng, math, sci = s.get("英組", "1"), s.get("數組", "1"), s.get("自組", "1")
            s_key = s.get("_student_key", "")
            
            subtotal = 0
            excl_names = []
            for _, book in df_b_use.iterrows():
                b_name = book.get("商品名稱", "")
                b_subj = book.get("科目", "")
                b_code = book.get("分組代號", "1")
                b_price = safe_price(book.get("單價", 0))
                b_key = book.get("_book_key", "")
                b_subj_check = "社會" if b_subj in ["歷", "地", "公", "歷史", "地理", "公民", "社會三科"] else b_subj
                if should_buy_book(b_subj_check, b_code, gifted, eng, math, sci):
                    if (s_key, b_key) in excluded_pairs:
                        excl_names.append(str(b_name))
                        has_any_excl = True
                    else:
                        subtotal += b_price
            student_rows.append([seat, name, gifted, eng, math, sci, subtotal, "、".join(excl_names)])

        headers_left = ["商品名稱", "科目", "分組代號", "購買數量", "單價"]
        for col_idx, h in enumerate(headers_left, 1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")

        headers_right = ["座號", "姓名", "資優", "英組", "數組", "自組", "應收總額"]
        if has_any_excl:
            headers_right.append("已自購免購")
        for col_idx, h in enumerate(headers_right, 7):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")

        thin_border = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                             top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

        for r_idx, b_row in enumerate(book_rows, 2):
            for c_idx, val in enumerate(b_row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = thin_border
                if c_idx > 1: cell.alignment = Alignment(horizontal="center")

        for r_idx, s_row in enumerate(student_rows, 2):
            row_out = s_row if has_any_excl else s_row[:7]
            for c_idx, val in enumerate(row_out, 7):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = thin_border
                if c_idx == 14:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                elif c_idx != 8:
                    cell.alignment = Alignment(horizontal="center")
                if c_idx == 13: cell.font = Font(bold=True)

        last_row = max(len(book_rows), len(student_rows)) + 2
        ws.cell(row=last_row, column=12, value="班級總計").font = Font(bold=True)
        total_sum = sum([s[6] for s in student_rows])
        tot_cell = ws.cell(row=last_row, column=13, value=total_sum)
        tot_cell.font = Font(bold=True, color="FF0000"); tot_cell.border = thin_border

        ws.column_dimensions['A'].width = 35; ws.column_dimensions['B'].width = 8
        ws.column_dimensions['C'].width = 10; ws.column_dimensions['D'].width = 10; ws.column_dimensions['E'].width = 8
        ws.column_dimensions['F'].width = 3  
        ws.column_dimensions['G'].width = 6 ; ws.column_dimensions['H'].width = 12
        ws.column_dimensions['I'].width = 8 ; ws.column_dimensions['J'].width = 8
        ws.column_dimensions['K'].width = 8 ; ws.column_dimensions['L'].width = 8
        ws.column_dimensions['M'].width = 12
        if has_any_excl:
            ws.column_dimensions['N'].width = 28

        if student_records is not None and plan is not None:
            write_individual_booklist_sheet(wb, student_records, plan)

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    st.subheader("🛠️ 開始作業")
    default_url = "https://reurl.cc/K2LgNe"
    url_col, refresh_col = st.columns([4, 1])
    with url_col:
        sheet_url = st.text_input("👇 請在下方輸入您的試算表網址：", value=default_url)
    with refresh_col:
        st.write("")
        st.write("")
        if st.button("🔄 重新讀取", key="tab1_refresh_sheet"):
            load_google_sheet_data.clear()

    if sheet_url:
        xlsx_url = get_google_sheet_xlsx_url(sheet_url)
        if xlsx_url:
            try:
                with st.spinner("正在連線至 Google Sheets 讀取資料..."):
                    df_students_raw, df_books_raw = load_google_sheet_data(xlsx_url)

                df_students = df_students_raw.copy()
                if "姓名" in df_students.columns:
                    df_students = df_students[df_students["姓名"].astype(str).str.strip() != ""].reset_index(drop=True)
                df_books = df_books_raw.copy()
                if "商品名稱" in df_books.columns:
                    df_books = df_books[df_books["商品名稱"].astype(str).str.strip() != ""].reset_index(drop=True)

                df_students, tab1_student_records = attach_student_keys(df_students, default_group="1")
                df_books, tab1_book_records = attach_book_keys(df_books, mode="tab1")

                st.success(f"✅ 讀取成功！共載入 {len(df_students)} 位學生、{len(df_books)} 筆書目報價。")

                def tab1_eligibility(student, book):
                    return should_buy_book(
                        book["subj_check"],
                        book["code"],
                        student["gifted"],
                        student["eng"],
                        student["math"],
                        student["sci"],
                    )

                st.divider()
                tab1_excluded_pairs, tab1_plan = render_exclusion_and_preview_section(
                    prefix="tab1",
                    student_records=tab1_student_records,
                    book_records=tab1_book_records,
                    eligibility_fn=tab1_eligibility,
                    show_publisher=False,
                )

                st.divider()
                tab1_sig = (
                    xlsx_url,
                    tuple(s["key"] for s in tab1_student_records),
                    tuple(b["key"] for b in tab1_book_records),
                    tuple(sorted(tab1_excluded_pairs)),
                )
                if st.session_state.get("tab1_report_sig") != tab1_sig:
                    st.session_state.tab1_receipts_output = None
                    st.session_state.tab1_master_output = None

                if st.button("🚀 開始產生檔案", type="primary"):
                    with st.spinner("系統正在計算與排版中..."):
                        st.session_state.tab1_receipts_output = generate_receipts_excel(
                            df_students, df_books, excluded_pairs=tab1_excluded_pairs
                        )
                        st.session_state.tab1_master_output = generate_master_excel(
                            df_students,
                            df_books,
                            excluded_pairs=tab1_excluded_pairs,
                            student_records=tab1_student_records,
                            plan=tab1_plan,
                        )
                        st.session_state.tab1_report_sig = tab1_sig
                    st.balloons()

                if st.session_state.get("tab1_receipts_output") and st.session_state.get("tab1_master_output"):
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            label="📥 下載【家長通知單】(A4列印版)",
                            data=st.session_state.tab1_receipts_output,
                            file_name="家長購書通知單.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    with col2:
                        st.download_button(
                            label="📥 下載【導師對帳總表】(收費明細)",
                            data=st.session_state.tab1_master_output,
                            file_name="導師收費對帳總表.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
            except Exception as e:
                st.error(f"❌ 讀取失敗！請確認網址與權限。（{e}）")


# =====================================================================
# 分頁 2：班級購書與費用對帳系統 (正式專業版)
# =====================================================================
with tab2:
    st.subheader("班級購書與費用對帳系統")
    
    st.info("""
    **【系統使用說明】**
    1. 請至「台中市教育雲端系統」下載 **班級名條** (若該年級有實施分組教學，請一併下載英文/數學分組名條)。
    2. 請向各家書商索取本次的 **報價單 CSV 或 Excel 檔**。
    3. 依序上傳下方檔案，若有其他自訂收費項目，可於「步驟二」表格手動新增。
    *(註：一年級免分組，僅需上傳「班級名單」與「書商報價單」即可，系統會自動產出全班明細)*
    """)

    st.markdown("#### 步驟一：上傳名單與報價單")

    col_a, col_b = st.columns(2)
    with col_a:
        file_class = st.file_uploader("1. 班級總名單 (Excel 檔)", type=["xlsx", "xls"])
        file_books_list = st.file_uploader("2. 書商報價單 (可框選多份 CSV/Excel)", type=["csv", "xlsx", "xls"], accept_multiple_files=True)
    with col_b:
        file_eng   = st.file_uploader("3. 英文分組名單 (可選)", type=["csv", "xlsx", "xls"])
        file_math  = st.file_uploader("4. 數學/自然分組名單 (可選)", type=["csv", "xlsx", "xls"])
        st.caption("提示：系統若讀取到 9 或 3 開頭之三年級名單，將自動連動數學與自然科為同組。")

    st.markdown("#### 步驟二：新增其他收費 (選填)")
    st.write("若無其他收費請留空。填寫「分組代號」系統將自動結算人數 (若為全班性收費請填 `1`)。")
    
    if 'custom_fees' not in st.session_state:
        st.session_state.custom_fees = pd.DataFrame([
            {"商品名稱": "", "科目": "其他", "分組代號": "1", "單價": 0}
        ])
        
    edited_custom_fees = st.data_editor(
        st.session_state.custom_fees,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "商品名稱": st.column_config.TextColumn("商品名稱 (必填)"),
            "科目": st.column_config.SelectboxColumn("科目分類", options=["國", "英", "數", "自", "社會", "其他"]),
            "分組代號": st.column_config.TextColumn("分組代號 (全班請填 1)"),
            "單價": st.column_config.NumberColumn("每人單價 (元)", min_value=0, step=1),
        }
    )

    if 'pdf_output' not in st.session_state: st.session_state.pdf_output = None
    if 'excel_output' not in st.session_state: st.session_state.excel_output = None

    def find_column(df, keywords, default_name):
        for col in df.columns:
            if any(kw in str(col) for kw in keywords): return col
        return None
        
    def infer_subject(name, code=""):
        # 依據書名或代號推斷科目分類
        name = str(name).replace("國中", "").replace("國小", "")
        if any(k in name for k in ["英", "文法", "單字", "聽力", "ABC", "abc"]): return "英"
        if any(k in name for k in ["數", "幾何", "代數", "算"]): return "數"
        if any(k in name for k in ["自", "理化", "生物", "地科", "科學"]): return "自"
        if any(k in name for k in ["歷", "地", "公", "社會"]): return "社會"
        if any(k in name for k in ["國", "文", "閱讀"]): return "國"
        
        code_str = str(code).replace(" ", "").upper()
        if "英" in code_str: return "英"
        if "數" in code_str: return "數"
        if "自" in code_str: return "自"
        return "其他"

    def standardize_book_code(val):
        val = str(val).strip()
        if val in ["1", "全", "", "nan", "None", "無"]: return "1"
        match = re.search(r'(\d+).*?([A-Za-z])', val)
        if match:
            cls_num = int(match.group(1))
            if cls_num >= 100: cls_num = cls_num % 100 
            return f"{cls_num}{match.group(2).upper()}"
        match_alpha = re.search(r'[A-Za-z]', val)
        if match_alpha: return match_alpha.group(0).upper()
        return val

    def parse_horizontal_group_file(uploaded_file, subj_hint=""):
        # 處理名單之合併儲存格與分組代號繼承
        uploaded_file.seek(0)
        if uploaded_file.name.endswith('.csv'): df_grp = pd.read_csv(uploaded_file, header=None).fillna("")
        else: df_grp = pd.read_excel(uploaded_file, header=None).fillna("")
        header_idx = -1
        for idx, row in df_grp.iterrows():
            if any(isinstance(v, str) and "姓名" in v for v in row.values):
                header_idx = idx ; break
        mapping = {}
        if header_idx != -1:
            col_to_group = {}
            current_group = "無"
            current_grade = 0
            for col_idx in range(df_grp.shape[1]):
                for r in range(0, header_idx):
                    val = str(df_grp.iloc[r, col_idx]).strip()
                    if val and val != "nan":
                        match_num_alpha = re.search(r'(\d+).*?([A-Za-z])', val)
                        if match_num_alpha:
                            cls_num = int(match_num_alpha.group(1))
                            current_grade = cls_num // 100 
                            if cls_num >= 100: cls_num = cls_num % 100 
                            current_group = f"{subj_hint}{cls_num}{match_num_alpha.group(2).upper()}"
                        else:
                            match_alpha = re.search(r'[A-Za-z]', val)
                            if match_alpha: current_group = f"{subj_hint}{match_alpha.group(0).upper()}"
                            elif len(val) <= 4: current_group = val
                col_to_group[col_idx] = (current_group, current_grade)

            for col_idx in range(df_grp.shape[1]):
                if "姓名" in str(df_grp.iloc[header_idx, col_idx]):
                    group_info = col_to_group.get(col_idx, ("無", 0))
                    names = df_grp.iloc[header_idx+1:, col_idx].astype(str).str.strip()
                    for name in names:
                        clean_name = name.replace(" ", "")
                        if clean_name and clean_name != "nan":
                            mapping[clean_name] = group_info
        return mapping

    def check_group_match(s_grp, b_code):
        if s_grp in ["1", "無", "免", ""]: return False
        b_clean = str(b_code).strip().upper()
        s_clean = str(s_grp).strip().upper()
        if s_clean == b_clean: return True
        
        s_match = re.search(r'(\d*)([A-Z]+)', s_clean)
        b_match = re.search(r'(\d*)([A-Z]+)', b_clean)
        if s_match and b_match:
            s_num, s_letter = s_match.groups()
            b_num, b_letter = b_match.groups()
            if s_letter != b_letter: return False
            if b_num and s_num and b_num != s_num: return False
            return True
        return False

    def auto_correct_group_code(df_b, df_s):
        # 依據名單人數自動校正書商提供的模糊分組代號
        counts = {
            '英': df_s[~df_s['英組'].isin(['無', '免', '1', ''])]['英組'].value_counts().to_dict(),
            '數': df_s[~df_s['數組'].isin(['無', '免', '1', ''])]['數組'].value_counts().to_dict(),
            '自': df_s[~df_s['自組'].isin(['無', '免', '1', ''])]['自組'].value_counts().to_dict()
        }
        
        for idx, row in df_b.iterrows():
            b_code = str(row['code']).strip().upper()
            b_subj = str(row['subj'])
            b_qty = row['qty']
            
            if b_qty > 0 and re.fullmatch(r'[A-Z]', b_code) and b_subj in counts:
                subj_counts = counts[b_subj]
                possible_fixes = []
                for grp_name, grp_count in subj_counts.items():
                    if grp_count == b_qty and b_code in grp_name:
                        possible_fixes.append(grp_name)
                        
                if len(possible_fixes) == 1:
                    df_b.at[idx, 'code'] = possible_fixes[0]
                    
        return df_b

    def is_book_for_student(b_subj, b_code, s_eng, s_math, s_sci, s_gifted):
        if s_gifted == "語資" and b_subj in ["國", "英"]: return False
        if s_gifted == "數資" and b_subj in ["數", "自"]: return False
        if "語資" in s_gifted and "數資" in s_gifted and b_subj in ["國", "英", "數", "自"]: return False
        
        if b_code == "1": return True
        
        if b_subj == "英": return check_group_match(s_eng, b_code)
        if b_subj == "數": return check_group_match(s_math, b_code)
        if b_subj == "自": return check_group_match(s_sci, b_code)
        if b_subj not in ["英", "數", "自"]: 
            return check_group_match(s_eng, b_code) or check_group_match(s_math, b_code)
        return False

    def generate_smart_pdf(df_students, df_books_clean, excluded_pairs=None, show_excluded_note=True):
        excluded_pairs = excluded_pairs or set()
        df_s_use = df_students if "_student_key" in df_students.columns else attach_student_keys(df_students, default_group="無")[0]
        df_b_use = df_books_clean if "_book_key" in df_books_clean.columns else attach_book_keys(df_books_clean, mode="tab2")[0]

        pdf_io = io.BytesIO()
        c = canvas.Canvas(pdf_io, pagesize=A4)
        width, height = A4
        pdf_font = 'CustomFont' if HAS_FONT else 'Helvetica'
        
        for _, student in df_s_use.iterrows():
            seat = normalize_seat(student.get("座號", ""))
            name = str(student.get("姓名", "")).strip()
            s_eng = str(student.get("英組", "無")).strip()
            s_math = str(student.get("數組", "無")).strip()
            s_sci = str(student.get("自組", "無")).strip()
            s_gifted = str(student.get("資優類別", "無")).strip()
            s_key = student.get("_student_key", "")

            personal_list = []
            excluded_names = []
            total_amount = 0
            for _, b in df_b_use.iterrows():
                b_key = b.get("_book_key", "")
                if is_book_for_student(b['subj'], b['code'], s_eng, s_math, s_sci, s_gifted):
                    if (s_key, b_key) in excluded_pairs:
                        excluded_names.append(str(b['name']))
                    else:
                        personal_list.append((b['name'], b['price']))
                        total_amount += b['price']

            c.setFont(pdf_font, 24)
            c.drawCentredString(width/2, height - 80, "學 期 各 項 費 用 通 知 單")
            s_eng_display = s_eng if s_eng not in ["1", "無", "免"] else "無"
            s_math_display = s_math if s_math not in ["1", "無", "免"] else "無"
            s_gifted_display = s_gifted if s_gifted not in ["1", "無", "免"] else "無"
            
            c.setFont(pdf_font, 14)
            c.drawString(60, height - 130, f"座號：{seat}        姓名：{name}")
            c.drawRightString(width - 60, height - 130, f"狀態：英({s_eng_display}) 數({s_math_display}) 資優({s_gifted_display})")
            c.setStrokeColorRGB(0, 0, 0)
            c.setLineWidth(2)
            c.line(60, height - 145, width - 60, height - 145)
            
            c.setFont(pdf_font, 12)
            c.drawString(70, height - 170, "項目名稱")
            c.drawRightString(width - 70, height - 170, "金額 (元)")
            c.setLineWidth(0.5)
            c.line(60, height - 180, width - 60, height - 180)
            
            has_note = bool(show_excluded_note and excluded_names)
            start_y = height - 205
            end_y = 195 if has_note else 160
            available_space = start_y - end_y
            item_count = len(personal_list)
            
            if item_count == 0:
                c.setFont(pdf_font, 12)
                c.drawString(70, start_y, "（本學期無收費項目，免繳費）")
                y_pos = start_y - 25
                line_height = 20
            else:
                default_line_height = 22 
                if item_count * default_line_height > available_space:
                    line_height = available_space / item_count
                    dynamic_font_size = max(9, int(line_height * 0.6)) 
                else:
                    line_height = default_line_height
                    dynamic_font_size = 12
                
                c.setFont(pdf_font, dynamic_font_size)
                y_pos = start_y
                for b_name, price in personal_list:
                    c.drawString(70, y_pos, b_name)
                    c.drawRightString(width - 70, y_pos, f"$ {int(price)}")
                    y_pos -= line_height
            
            c.setLineWidth(1)
            c.line(60, y_pos + (line_height * 0.6), width - 60, y_pos + (line_height * 0.6))
            c.setFont(pdf_font, 16)
            c.drawString(70, y_pos - 20, "應繳總計金額：")
            c.setFillColorRGB(0.8, 0, 0)
            c.drawRightString(width - 70, y_pos - 20, f"$ {int(total_amount)} 元")
            c.setFillColorRGB(0, 0, 0)

            if has_note:
                note_text = "※ 已自購免購（不計費）：" + "、".join(excluded_names)
                note_font_size = 9
                note_lines = wrap_pdf_text(note_text, pdf_font, note_font_size, width - 140, max_lines=3)
                c.setFont(pdf_font, note_font_size)
                c.setFillColorRGB(0.35, 0.35, 0.35)
                note_y = y_pos - 42
                for n_line in note_lines:
                    c.drawString(70, note_y, n_line)
                    note_y -= 12
                c.setFillColorRGB(0, 0, 0)
            
            box_y = 55
            c.setFont(pdf_font, 12)
            c.drawString(60, box_y + 55, "【家長簽章回條】")
            c.setFont(pdf_font, 11)
            c.drawString(60, box_y + 25, f"本人已確認上述 座號 {seat} {name} 之明細與金額無誤。")
            c.drawString(width - 150, box_y + 25, "家長簽名：")
            c.showPage()
        c.save()
        pdf_io.seek(0)
        return pdf_io

    def fill_sheet_data(ws, sheet_title, df_students, df_books, excluded_pairs=None):
        excluded_pairs = excluded_pairs or set()
        df_s_use = df_students if "_student_key" in df_students.columns else attach_student_keys(df_students, default_group="無")[0]
        df_b_use = df_books if "_book_key" in df_books.columns else attach_book_keys(df_books, mode="tab2")[0]

        ws.title = sheet_title
        book_rows = []
        for _, book in df_b_use.iterrows():
            b_name, b_subj, b_price = book['name'], book['subj'], book['price']
            b_code_str = str(book['code']).strip()
            b_key = book.get("_book_key", "")
            qty = 0
            for _, s in df_s_use.iterrows():
                s_eng, s_math, s_sci = str(s.get("英組","無")), str(s.get("數組","無")), str(s.get("自組","無"))
                s_gifted = str(s.get("資優類別", "無"))
                s_key = s.get("_student_key", "")
                if is_book_for_student(b_subj, b_code_str, s_eng, s_math, s_sci, s_gifted) and (s_key, b_key) not in excluded_pairs:
                    qty += 1
            book_rows.append([b_name, b_subj, b_code_str, qty, b_price])

        student_rows = []
        has_any_excl = False
        for _, s in df_s_use.iterrows():
            seat, name = normalize_seat(s.get("座號", "")), s.get("姓名", "")
            gifted = str(s.get("資優類別", "無"))
            s_eng, s_math, s_sci = str(s.get("英組","無")), str(s.get("數組","無")), str(s.get("自組","無"))
            s_key = s.get("_student_key", "")
            subtotal = 0
            excl_names = []
            for _, book in df_b_use.iterrows():
                b_key = book.get("_book_key", "")
                if is_book_for_student(book['subj'], book['code'], s_eng, s_math, s_sci, gifted):
                    if (s_key, b_key) in excluded_pairs:
                        excl_names.append(str(book['name']))
                        has_any_excl = True
                    else:
                        subtotal += book['price']
            student_rows.append([seat, name, gifted, s_eng, s_math, s_sci, subtotal, "、".join(excl_names)])

        headers_left = ["項目名稱", "科目", "分組代號", "結算數量", "單價"]
        for col_idx, h in enumerate(headers_left, 1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")

        headers_right = ["座號", "姓名", "資優", "英組", "數組", "自組", "應收總額"]
        if has_any_excl:
            headers_right.append("已自購免購")
        for col_idx, h in enumerate(headers_right, 7):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")

        thin_border = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'), top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))
        for r_idx, b_row in enumerate(book_rows, 2):
            for c_idx, val in enumerate(b_row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = thin_border
                if c_idx > 1: cell.alignment = Alignment(horizontal="center")

        for r_idx, s_row in enumerate(student_rows, 2):
            row_out = s_row if has_any_excl else s_row[:7]
            for c_idx, val in enumerate(row_out, 7):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = thin_border
                if c_idx == 14:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                elif c_idx != 8:
                    cell.alignment = Alignment(horizontal="center")
                if c_idx == 13: cell.font = Font(bold=True)

        last_row = max(len(book_rows), len(student_rows)) + 2
        ws.cell(row=last_row, column=12, value="單頁總計").font = Font(bold=True)
        total_sum = sum([s[6] for s in student_rows])
        tot_cell = ws.cell(row=last_row, column=13, value=total_sum)
        tot_cell.font = Font(bold=True, color="FF0000"); tot_cell.border = thin_border
        
        ws.column_dimensions['A'].width = 35; ws.column_dimensions['B'].width = 8
        ws.column_dimensions['C'].width = 10; ws.column_dimensions['D'].width = 10; ws.column_dimensions['E'].width = 8
        ws.column_dimensions['F'].width = 3  
        ws.column_dimensions['G'].width = 6 ; ws.column_dimensions['H'].width = 12
        ws.column_dimensions['I'].width = 8 ; ws.column_dimensions['J'].width = 8
        ws.column_dimensions['K'].width = 8 ; ws.column_dimensions['L'].width = 8
        ws.column_dimensions['M'].width = 12
        if has_any_excl:
            ws.column_dimensions['N'].width = 28

    def generate_excel_master_dynamic(df_students, df_books_clean, excluded_pairs=None, student_records=None, plan=None):
        excluded_pairs = excluded_pairs or set()
        df_s_use = df_students if "_student_key" in df_students.columns else attach_student_keys(df_students, default_group="無")[0]
        df_b_use = df_books_clean if "_book_key" in df_books_clean.columns else attach_book_keys(df_books_clean, mode="tab2")[0]

        wb = Workbook()
        ws1 = wb.active
        fill_sheet_data(ws1, "班級收費總表(全)", df_s_use, df_b_use, excluded_pairs=excluded_pairs)

        if student_records is not None and plan is not None:
            write_individual_booklist_sheet(wb, student_records, plan)
        
        unique_publishers = df_b_use['publisher'].unique()
        for pub in unique_publishers:
            df_pub = df_b_use[df_b_use['publisher'] == pub]
            ws_pub = wb.create_sheet()
            safe_pub_name = str(pub).replace("/", "").replace("\\", "")[:12]
            fill_sheet_data(ws_pub, f"{safe_pub_name}對帳表", df_s_use, df_pub, excluded_pairs=excluded_pairs)
            
        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    # 執行系統主程式
    if file_class and file_books_list and len(file_books_list) > 0:
        try:
            file_class.seek(0)
            df_temp = pd.read_excel(file_class, header=None).fillna("")
            header_idx = 0
            for idx, row in df_temp.iterrows():
                if any("姓名" in str(v) for v in row.values):
                    header_idx = idx ; break
            file_class.seek(0)
            df_s = pd.read_excel(file_class, skiprows=header_idx).fillna("")
            
            c_name = find_column(df_s, ["姓名", "名稱", "學生"], "姓名")
            if c_name: df_s = df_s[df_s[c_name].astype(str).str.strip() != ""] 
            
            c_seat = find_column(df_s, ["座號", "號碼", "序號"], "座號")
            if c_seat and c_seat != "座號": df_s = df_s.rename(columns={c_seat: "座號"})
            if c_name and c_name != "姓名": df_s = df_s.rename(columns={c_name: "姓名"})
            
            for col in ["英組", "數組", "自組", "資優類別"]:
                if col not in df_s.columns: df_s[col] = "無"

            eng_map = parse_horizontal_group_file(file_eng, "英") if file_eng else {}
            math_map = parse_horizontal_group_file(file_math, "數") if file_math else {}

            for idx, row in df_s.iterrows():
                clean_name = str(row["姓名"]).replace(" ", "").strip()
                
                if file_eng:
                    if clean_name in eng_map: df_s.at[idx, "英組"] = eng_map[clean_name][0]
                    else:
                        df_s.at[idx, "資優類別"] = "語資"
                        df_s.at[idx, "英組"] = "免"
                        
                if file_math:
                    if clean_name in math_map:
                        grp_code, grade = math_map[clean_name]
                        df_s.at[idx, "數組"] = grp_code
                        if grade in [9, 3]: 
                            df_s.at[idx, "自組"] = grp_code  
                        else:
                            df_s.at[idx, "自組"] = "免"
                    else:
                        current_gifted = str(df_s.at[idx, "資優類別"])
                        if current_gifted in ["無", "", "1"]: df_s.at[idx, "資優類別"] = "數資"
                        elif current_gifted == "語資": df_s.at[idx, "資優類別"] = "語資/數資"
                        df_s.at[idx, "數組"] = "免"
                        df_s.at[idx, "自組"] = "免"

            df_s = df_s.reset_index(drop=True)
            df_s, tab2_student_records = attach_student_keys(df_s, default_group="無")

            with st.expander("名單與分組狀態預覽"):
                st.write("請確認全班名單與分組判定結果是否正確：")
                st.dataframe(df_s[["座號", "姓名", "英組", "數組", "自組", "資優類別"]])

            all_books_clean_list = []
            
            for fb in file_books_list:
                publisher_name = fb.name.split('.')[0]
                extracted_books = []
                header_skip = 0
                keywords_header = ["品名", "商品", "名稱", "單價", "價格", "金額", "數量", "件數"]
                
                fb.seek(0)
                if fb.name.endswith('.csv'):
                    bytes_data = fb.read()
                    lines = []
                    used_enc = 'utf-8'
                    for enc in ['utf-8', 'big5', 'cp950', 'utf-8-sig']:
                        try:
                            lines = bytes_data.decode(enc).splitlines()
                            used_enc = enc ; break
                        except: continue
                    
                    for idx, line in enumerate(lines):
                        if sum(1 for kw in keywords_header if kw in line) >= 2:
                            header_skip = idx ; break
                            
                    if header_skip > 0 and len(lines) > 0:
                        clean_first_line = lines[0].replace(',', '').strip()
                        if clean_first_line:
                            publisher_name = clean_first_line.replace("估價單", "").replace("報價單", "").strip()
                            
                    fb.seek(0)
                    try: df_b = pd.read_csv(fb, skiprows=header_skip, encoding=used_enc).fillna("")
                    except: 
                        fb.seek(0)
                        df_b = pd.read_csv(fb, skiprows=header_skip, on_bad_lines='skip').fillna("")
                
                else:
                    df_temp2 = pd.read_excel(fb, header=None).fillna("")
                    for idx, row in df_temp2.iterrows():
                        row_str = "".join([str(val) for val in row.values])
                        if sum(1 for kw in keywords_header if kw in row_str) >= 2:
                            header_skip = idx ; break
                            
                    if header_skip > 0:
                        first_row_str = "".join([str(val) for val in df_temp2.iloc[0].values if str(val).strip() and str(val) != "nan"])
                        if first_row_str:
                            publisher_name = first_row_str.replace("估價單", "").replace("報價單", "").strip()
                            
                    fb.seek(0)
                    df_b = pd.read_excel(fb, skiprows=header_skip).fillna("")

                b_col_name = find_column(df_b, ["品名", "商品", "名稱", "書籍"], "商品名稱")
                b_col_price = find_column(df_b, ["單價", "價格", "金額"], "單價")
                b_col_code = find_column(df_b, ["附記", "備註", "分組", "代號"], "分組代號")
                b_col_qty = find_column(df_b, ["數量", "量", "件數"], "數量")
                b_col_subj = find_column(df_b, ["科目", "類別"], "科目")
                
                df_b['parsed_price'] = pd.to_numeric(df_b[b_col_price], errors='coerce')
                
                for i in range(len(df_b)):
                    price_val = df_b['parsed_price'].iloc[i]
                    if pd.notna(price_val) and price_val > 0: 
                        raw_name = str(df_b[b_col_name].iloc[i])
                        raw_code = standardize_book_code(df_b[b_col_code].iloc[i]) if b_col_code else "1"
                        raw_qty = df_b[b_col_qty].iloc[i] if b_col_qty else 0
                        qty_val = pd.to_numeric(raw_qty, errors='coerce')
                        qty_val = int(qty_val) if pd.notna(qty_val) else 0
                        
                        subj_val = str(df_b[b_col_subj].iloc[i]) if b_col_subj else infer_subject(raw_name, raw_code)

                        extracted_books.append({
                            'name': raw_name,
                            'price': int(price_val),
                            'code': raw_code,
                            'qty': qty_val,
                            'subj': subj_val,
                            'publisher': publisher_name
                        })

                if extracted_books:
                    df_temp_clean = pd.DataFrame(extracted_books)
                    all_books_clean_list.append(df_temp_clean)
            
            custom_books = []
            for _, row in edited_custom_fees.iterrows():
                name_val = str(row.get("商品名稱", "")).strip()
                price_val = pd.to_numeric(row.get("單價", 0), errors='coerce')
                if name_val and name_val != "nan" and pd.notna(price_val) and price_val > 0:
                    c_subj = str(row.get("科目", "其他"))
                    c_code = standardize_book_code(row.get("分組代號", "1"))
                    custom_books.append({
                        'name': name_val,
                        'price': int(price_val),
                        'code': c_code,
                        'qty': 0,
                        'subj': c_subj,
                        'publisher': "新增自訂項目"
                    })
            if custom_books:
                all_books_clean_list.append(pd.DataFrame(custom_books))
                
            df_books_clean = pd.concat(all_books_clean_list, ignore_index=True) if all_books_clean_list else pd.DataFrame()

            if not df_books_clean.empty:
                df_books_clean = auto_correct_group_code(df_books_clean, df_s)
                df_books_clean['subj'] = df_books_clean['subj'].apply(lambda x: "社會" if x in ["歷", "地", "公", "歷史", "地理", "公民"] else x)

            st.divider()
            st.markdown("#### 步驟三：移除全班皆不需購買之書籍 (若無請跳過)")
            if not df_books_clean.empty:
                all_book_names = df_books_clean['name'].unique().tolist()
                remove_books = st.multiselect(
                    "若書商提供之報價單包含全班不需購買之品項，請於此處選擇並剔除：",
                    options=all_book_names,
                    help="被選中的書籍將不會出現在最終的通知單與對帳表中。"
                )
                if remove_books:
                    df_books_clean = df_books_clean[~df_books_clean['name'].isin(remove_books)].reset_index(drop=True)
                    st.success(f"已成功剔除 {len(remove_books)} 項書籍。")

            tab2_book_records = []
            if not df_books_clean.empty:
                df_books_clean, tab2_book_records = attach_book_keys(df_books_clean, mode="tab2")

            with st.expander("書目與收費項目預覽"):
                st.write("系統已讀取以下書目與收費資料。若代號有缺漏，系統會自動嘗試依據人數與科目進行校正：")
                if not df_books_clean.empty:
                    st.dataframe(df_books_clean.drop(columns=["_book_key"], errors="ignore"))
                else:
                    st.warning("尚未讀取到有效之書目資料，請確認檔案格式。")

            tab2_excluded_pairs = set()
            tab2_plan = None
            if not df_books_clean.empty:
                def tab2_eligibility(student, book):
                    return is_book_for_student(
                        book["subj"],
                        book["code"],
                        student["eng"],
                        student["math"],
                        student["sci"],
                        student["gifted"],
                    )

                st.divider()
                st.markdown("#### 步驟四：個別學生已自購免購設定與書目預覽")
                tab2_excluded_pairs, tab2_plan = render_exclusion_and_preview_section(
                    prefix="tab2",
                    student_records=tab2_student_records,
                    book_records=tab2_book_records,
                    eligibility_fn=tab2_eligibility,
                    show_publisher=True,
                )

            st.divider()
            st.markdown("#### 步驟五：產生報表")
            show_excluded_note_pdf = st.checkbox(
                "在 PDF 家長通知單上註明「已自購免購（不計費）」書目",
                value=True,
                key="tab2_pdf_show_excl_note",
            )

            tab2_sig = (
                tuple(s["key"] for s in tab2_student_records),
                tuple(b["key"] for b in tab2_book_records),
                tuple(sorted(tab2_excluded_pairs)),
                show_excluded_note_pdf,
            )
            if st.session_state.get("tab2_report_sig") != tab2_sig:
                if st.session_state.get("tab2_report_sig") is not None and (
                    st.session_state.pdf_output or st.session_state.excel_output
                ):
                    st.info("ℹ️ 名單或免購設定已更新，請重新點擊下方按鈕產生最新報表。")
                st.session_state.pdf_output = None
                st.session_state.excel_output = None

            if not df_books_clean.empty and st.button("確認無誤，開始產生報表", type="primary"):
                with st.spinner("正在產生 PDF 與 Excel 檔案中..."):
                    st.session_state.pdf_output = generate_smart_pdf(
                        df_s,
                        df_books_clean,
                        excluded_pairs=tab2_excluded_pairs,
                        show_excluded_note=show_excluded_note_pdf,
                    )
                    st.session_state.excel_output = generate_excel_master_dynamic(
                        df_s,
                        df_books_clean,
                        excluded_pairs=tab2_excluded_pairs,
                        student_records=tab2_student_records,
                        plan=tab2_plan,
                    )
                    st.session_state.tab2_report_sig = tab2_sig
                    st.success("報表產生完成！請點擊下方按鈕下載。")

            if st.session_state.pdf_output and st.session_state.excel_output:
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="下載【家長通知單】(PDF 檔)", 
                        data=st.session_state.pdf_output, 
                        file_name="家長通知單_費用明細.pdf", 
                        mime="application/pdf"
                    )
                with col2:
                    st.download_button(
                        label="下載【導師對帳總表】(Excel 檔)", 
                        data=st.session_state.excel_output, 
                        file_name="導師對帳總表.xlsx", 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        
        except Exception as e:
            st.error(f"系統讀取發生錯誤：{e}")
            st.info("請檢查上傳的檔案格式是否正確。")
    else:
        st.info("請先依序上傳名單與報價單，系統將自動顯示預覽畫面。")


# =====================================================================
# 分頁 3：依正式課表、老師與分組名單產生個人功課表
# =====================================================================
with tab3:
    st.subheader("個人功課表產生器")
    st.info("""
    系統先以老師辨識分組節次，再依數理、英文與資優名單安排跑班。
    **輸出只保留正式課表上的科目名稱**，不會把內部判斷的真實課程寫出去。
    自動判斷後一定會先顯示「節次規則」及「學生分組」供老師校正，確認後才產生課表。
    **請老師務必自行修改學生的「資優類別」；系統不會預存任何學生姓名。**
    語文資優生於英文、國文節次前往資優班；數理分組及其他課程仍依原班規則上課。
    """)

    if pdfplumber is None:
        st.error("目前環境缺少 pdfplumber。請在 requirements.txt 加入 `pdfplumber` 後重新部署。")
        st.stop()

    DAY_NAMES = ["星期一", "星期二", "星期三", "星期四", "星期五"]
    PERIOD_LABELS = {
        0: "早修", 1: "第1節", 2: "第2節", 3: "第3節", 4: "第4節",
        5: "第5節", 6: "第6節", 7: "第7節", 8: "第8節", 9: "第9節",
    }
    PERIOD_NUMBERS = {v: k for k, v in PERIOD_LABELS.items()}
    RULE_NORMAL = "不分組"
    RULE_STEM = "數理分組"
    RULE_ENGLISH = "英文分組"
    RULE_OPTIONS = [RULE_NORMAL, RULE_STEM, RULE_ENGLISH]

    # 為保護學生隱私，程式不預存資優生姓名；請老師在校正表自行指定。
    VERIFIED_STEM_GIFTED = set()

    def clean_pdf_lines(value):
        if value is None:
            return []
        result = []
        for line in str(value).splitlines():
            line = re.sub(r"\s+", "", line).strip()
            if line:
                result.append(line)
        return result

    def parse_course_cell(value):
        lines = clean_pdf_lines(value)
        return {
            "subject": lines[0] if lines else "",
            "teacher": lines[1] if len(lines) >= 2 else "",
            "raw": "\n".join(lines),
        }

    @st.cache_data(show_spinner=False)
    def parse_timetable_pdf(pdf_bytes):
        schedules = {}
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                # 部分學校PDF會把 901 抽成「9 01」，因此先容許班號中含空白。
                class_match = re.search(r"班級\s*[：:]\s*([0-9\s]{3,6})", page_text)
                if not class_match:
                    continue
                class_digits = re.sub(r"\s+", "", class_match.group(1))
                if len(class_digits) < 3:
                    continue
                class_no = int(class_digits[:3])
                tables = page.extract_tables() or []
                if not tables:
                    continue
                table = max(tables, key=lambda t: len(t) if t else 0)
                class_schedule = {}

                # pdfplumber表格：第1列標題、第2列早修、第3～11列為第1～9節。
                for row_index, row in enumerate(table):
                    if row_index == 1:
                        period_no = 0
                    elif 2 <= row_index <= 10:
                        period_no = row_index - 1
                    else:
                        continue

                    row = list(row or [])
                    while len(row) < 8:
                        row.append("")
                    for day_index in range(5):
                        class_schedule[(period_no, day_index)] = parse_course_cell(row[3 + day_index])

                schedules[class_no] = class_schedule
        return schedules

    @st.cache_data(show_spinner=False)
    def parse_group_pdf(pdf_bytes, group_kind):
        rows = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if not pdf.pages:
                return rows
            page = pdf.pages[0]
            page_text = page.extract_text() or ""
            tables = page.extract_tables() or []

            if group_kind == "數理":
                markers = re.findall(r"(\d{3})\s*數理\s*([A-Za-z])", page_text)
            else:
                markers = re.findall(r"(\d{3})\s*英\s*([A-Za-z])", page_text)

            if len(markers) < len(tables):
                raise ValueError(f"無法從{group_kind}分組PDF辨識每一欄的目的班級。")

            for table_index, table in enumerate(tables):
                destination_class = int(markers[table_index][0])
                group_letter = markers[table_index][1].upper()
                group_label = f"{destination_class}{group_kind}{group_letter}"
                grade_base = (destination_class // 100) * 100

                for source_row in (table or [])[1:]:
                    source_row = list(source_row or [])
                    while len(source_row) < 5:
                        source_row.append("")
                    original_class = str(source_row[1] or "").strip()
                    seat = str(source_row[2] or "").strip()
                    student_id = str(source_row[3] or "").strip()
                    name = re.sub(r"\s+", "", str(source_row[4] or ""))

                    if not name or not original_class:
                        continue
                    try:
                        home_class = grade_base + int(float(original_class))
                        seat_no = int(float(seat))
                    except (TypeError, ValueError):
                        continue

                    student_key = student_id if student_id else f"{home_class}-{seat_no}-{name}"
                    rows.append({
                        "學生識別": student_key,
                        "學號": student_id,
                        "姓名": name,
                        "原班": home_class,
                        "座號": seat_no,
                        "分組種類": group_kind,
                        "分組代號": group_label,
                        "目的班級": destination_class,
                    })
        return rows

    def build_student_group_table(math_rows, english_rows):
        students = {}
        for group_rows in [math_rows, english_rows]:
            for item in group_rows:
                key = item["學生識別"]
                if key not in students:
                    students[key] = {
                        "學生識別": key,
                        "原班": item["原班"],
                        "座號": item["座號"],
                        "學號": item["學號"],
                        "姓名": item["姓名"],
                        "數理去向": "未設定",
                        "英文去向": "未設定",
                        "資優類別": "無",
                    }
                if item["分組種類"] == "數理":
                    students[key]["數理去向"] = item["分組代號"]
                else:
                    students[key]["英文去向"] = item["分組代號"]

        result = pd.DataFrame(students.values())
        if result.empty:
            return result
        for idx, student in result.iterrows():
            if student["姓名"] in VERIFIED_STEM_GIFTED:
                result.at[idx, "資優類別"] = "數理資優"
        return result.sort_values(["原班", "座號"]).reset_index(drop=True)

    def teacher_seed_pools(schedules, class_numbers):
        pools = {}
        stem_words = ["數學", "自然科學", "理化", "生物", "地科", "STEM", "數學專題"]
        for class_no in class_numbers:
            stem_teachers = set()
            english_teachers = set()
            for cell in schedules.get(class_no, {}).values():
                subject = str(cell.get("subject", ""))
                teacher = str(cell.get("teacher", ""))
                if not teacher:
                    continue
                if any(word in subject for word in stem_words):
                    stem_teachers.add(teacher)
                if "英語文" in subject:
                    english_teachers.add(teacher)
            pools[class_no] = {"stem": stem_teachers, "english": english_teachers}
        return pools

    def infer_slot_rules(schedules, cluster_classes, gifted_class):
        all_classes = list(cluster_classes)
        if gifted_class:
            all_classes.append(gifted_class)
        pools = teacher_seed_pools(schedules, all_classes)
        rules = {}

        for period_no in range(10):
            for day_index in range(5):
                english_hits = 0
                stem_hits = 0
                for class_no in cluster_classes:
                    teacher = schedules.get(class_no, {}).get((period_no, day_index), {}).get("teacher", "")
                    if teacher and teacher in pools.get(class_no, {}).get("english", set()):
                        english_hits += 1
                    if teacher and teacher in pools.get(class_no, {}).get("stem", set()):
                        stem_hits += 1

                gifted_teacher = ""
                gifted_is_stem = False
                if gifted_class:
                    gifted_teacher = schedules.get(gifted_class, {}).get((period_no, day_index), {}).get("teacher", "")
                    gifted_is_stem = bool(
                        gifted_teacher and
                        gifted_teacher in pools.get(gifted_class, {}).get("stem", set())
                    )

                if english_hits == len(cluster_classes) and english_hits > 0:
                    rules[(period_no, day_index)] = RULE_ENGLISH
                elif gifted_is_stem and stem_hits >= max(2, len(cluster_classes) - 1):
                    # 914同節有數理資優課，且至少兩個跑班教室出現數理老師。
                    # 可涵蓋「科目名稱或其中一位老師被偽裝」的節次。
                    rules[(period_no, day_index)] = RULE_STEM
                else:
                    rules[(period_no, day_index)] = RULE_NORMAL
        return rules

    def compact_cell(cell):
        if not cell or not cell.get("raw"):
            return ""
        return "／".join(clean_pdf_lines(cell.get("raw", "")))

    def build_rule_editor_table(schedules, cluster_classes, gifted_class, inferred_rules):
        rows = []
        visible_classes = list(cluster_classes)
        if gifted_class:
            visible_classes.append(gifted_class)
        for period_no in range(10):
            for day_index, day_name in enumerate(DAY_NAMES):
                row = {
                    "星期": day_name,
                    "節次": PERIOD_LABELS[period_no],
                    "課程類型": inferred_rules[(period_no, day_index)],
                }
                for class_no in visible_classes:
                    row[str(class_no)] = compact_cell(
                        schedules.get(class_no, {}).get((period_no, day_index), {})
                    )
                rows.append(row)
        return pd.DataFrame(rows)

    def route_class_from_label(route_label):
        match = re.match(r"\s*(\d{3})", str(route_label))
        return int(match.group(1)) if match else None

    def is_home_chinese_slot(schedules, home_class, period_no, day_index):
        """先由正式國文課建立老師名單，再用老師判斷該節是否為國文。"""
        home_schedule = schedules.get(home_class, {})
        chinese_teachers = {
            str(cell.get("teacher", ""))
            for cell in home_schedule.values()
            if any(word in str(cell.get("subject", "")) for word in ["國語文", "國文"])
            and str(cell.get("teacher", ""))
        }
        current_teacher = str(
            home_schedule.get((period_no, day_index), {}).get("teacher", "")
        )
        return bool(current_teacher and current_teacher in chinese_teachers)

    def make_personal_grid(student, schedules, rule_map, gifted_class):
        home_class = int(student["原班"])
        gifted_type = str(student.get("資優類別", "無"))
        result = {}
        errors = []

        for period_no in range(10):
            for day_index in range(5):
                rule = rule_map.get((period_no, day_index), RULE_NORMAL)
                destination_class = home_class
                route_label = "原班"
                cell_kind = "normal"

                if rule == RULE_STEM:
                    cell_kind = "stem"
                    if "數理資優" in gifted_type:
                        destination_class = gifted_class
                        route_label = f"{gifted_class}數理資優"
                        cell_kind = "gifted"
                    else:
                        route_label = str(student.get("數理去向", "未設定"))
                        destination_class = route_class_from_label(route_label)
                elif rule == RULE_ENGLISH:
                    cell_kind = "english"
                    if "語文資優" in gifted_type:
                        destination_class = gifted_class
                        route_label = f"{gifted_class}語文資優"
                        cell_kind = "gifted"
                    else:
                        route_label = str(student.get("英文去向", "未設定"))
                        destination_class = route_class_from_label(route_label)
                elif "語文資優" in gifted_type and is_home_chinese_slot(
                    schedules, home_class, period_no, day_index
                ):
                    destination_class = gifted_class
                    route_label = f"{gifted_class}語文資優"
                    cell_kind = "gifted"
                    # 國文不是一般跑班規則，但語文資優生在此節仍需抽離。
                    rule = "語文資優抽離"

                if destination_class is None or destination_class not in schedules:
                    errors.append(
                        f"{student['姓名']}：{DAY_NAMES[day_index]}{PERIOD_LABELS[period_no]}的去向「{route_label}」無有效課表。"
                    )
                    cell = {"subject": "⚠️去向未設定", "teacher": "", "raw": "⚠️去向未設定"}
                    cell_kind = "warning"
                else:
                    cell = schedules.get(destination_class, {}).get(
                        (period_no, day_index), {"subject": "", "teacher": "", "raw": ""}
                    )

                official_lines = clean_pdf_lines(cell.get("raw", ""))
                display_lines = list(official_lines)
                if rule != RULE_NORMAL:
                    # 即使目的班級與原班相同，也仍標示為跑班課。
                    display_lines.append(f"【跑班：{route_label}／{destination_class}教室】")
                    if not official_lines:
                        errors.append(
                            f"{student['姓名']}：{DAY_NAMES[day_index]}{PERIOD_LABELS[period_no]}的"
                            f"{destination_class}課表是空白。"
                        )
                        display_lines.insert(0, "⚠️正式課表空白")
                        cell_kind = "warning"

                result[(period_no, day_index)] = {
                    "text": "\n".join(display_lines),
                    "kind": cell_kind,
                }
        return result, errors

    def personal_grid_dataframe(grid):
        rows = []
        for period_no in range(10):
            row = {"節次": PERIOD_LABELS[period_no]}
            for day_index, day_name in enumerate(DAY_NAMES):
                row[day_name] = grid[(period_no, day_index)]["text"]
            rows.append(row)
        return pd.DataFrame(rows).set_index("節次")

    def write_personal_sheet(ws, student, grid):
        title = f"{int(student['原班'])}班 {int(student['座號'])}號 {student['姓名']} 個人功課表"
        ws.merge_cells("A1:F1")
        ws["A1"] = title
        ws["A1"].font = Font(name="微軟正黑體", size=16, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        ws.merge_cells("A2:F2")
        ws["A2"] = (
            f"數理：{student.get('數理去向', '未設定')}　"
            f"英文：{student.get('英文去向', '未設定')}　"
            f"資優：{student.get('資優類別', '無')}"
        )
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

        headers = ["節次"] + DAY_NAMES
        for column_no, header in enumerate(headers, 1):
            cell = ws.cell(3, column_no, header)
            cell.font = Font(name="微軟正黑體", bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4472C4")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        fills = {
            "normal": PatternFill("solid", fgColor="FFFFFF"),
            "stem": PatternFill("solid", fgColor="DDEBF7"),
            "english": PatternFill("solid", fgColor="FCE4D6"),
            "gifted": PatternFill("solid", fgColor="E4DFEC"),
            "warning": PatternFill("solid", fgColor="FFC7CE"),
        }
        thin = Side(style="thin", color="808080")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for period_no in range(10):
            excel_row = period_no + 4
            label_cell = ws.cell(excel_row, 1, PERIOD_LABELS[period_no])
            label_cell.font = Font(name="微軟正黑體", bold=True)
            label_cell.alignment = Alignment(horizontal="center", vertical="center")
            label_cell.border = border
            ws.row_dimensions[excel_row].height = 58

            for day_index in range(5):
                item = grid[(period_no, day_index)]
                cell = ws.cell(excel_row, day_index + 2, item["text"])
                cell.font = Font(name="微軟正黑體", size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.fill = fills[item["kind"]]
                cell.border = border

        for col in "BCDEF":
            ws.column_dimensions[col].width = 24
        ws.column_dimensions["A"].width = 10
        ws.freeze_panes = "B4"
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.4, bottom=0.4)

    def generate_schedule_workbook(student_records, schedules, rule_map, gifted_class):
        wb = Workbook()
        wb.remove(wb.active)
        all_errors = []
        used_sheet_names = set()

        for student in student_records:
            grid, errors = make_personal_grid(student, schedules, rule_map, gifted_class)
            all_errors.extend(errors)
            base_name = f"{int(student['座號']):02d}{student['姓名']}"[:31]
            sheet_name = base_name
            suffix = 2
            while sheet_name in used_sheet_names:
                sheet_name = f"{base_name[:27]}_{suffix}"
                suffix += 1
            used_sheet_names.add(sheet_name)
            ws = wb.create_sheet(sheet_name)
            write_personal_sheet(ws, student, grid)

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue(), all_errors

    def _pdf_wrapped_lines(text, font_name, font_size, max_width):
        """依實際字寬切行，保留課表原本的換行。"""
        wrapped = []
        pdf_text = str(text or "").replace("【", "[").replace("】", "]").replace("／", "/")
        for source_line in pdf_text.splitlines() or [""]:
            current = ""
            for char in source_line:
                candidate = current + char
                if current and pdfmetrics.stringWidth(candidate, font_name, font_size) > max_width:
                    wrapped.append(current)
                    current = char
                else:
                    current = candidate
            wrapped.append(current)
        return wrapped

    def draw_personal_schedule_pdf_page(pdf, student, grid):
        """在既有 PDF 畫布加入一位學生的一頁橫式 A4 功課表。"""
        page_width, page_height = A4[1], A4[0]
        font_name = "CustomFont" if HAS_FONT else "Helvetica"

        title = f"{int(student['原班'])}班 {int(student['座號'])}號 {student['姓名']} 個人功課表"
        subtitle = (
            f"數理：{student.get('數理去向', '未設定')}　"
            f"英文：{student.get('英文去向', '未設定')}　"
            f"資優：{student.get('資優類別', '無')}"
        )
        pdf.setTitle(title)
        pdf.setFont(font_name, 16)
        pdf.drawCentredString(page_width / 2, page_height - 28, title)
        pdf.setFont(font_name, 9)
        pdf.drawCentredString(page_width / 2, page_height - 45, subtitle)

        left = 20
        bottom = 20
        table_top = page_height - 58
        table_height = table_top - bottom
        header_height = 27
        row_height = (table_height - header_height) / 10
        first_width = 58
        day_width = (page_width - 40 - first_width) / 5
        widths = [first_width] + [day_width] * 5
        headers = ["節次"] + DAY_NAMES

        fill_colors = {
            "normal": (1, 1, 1),
            "stem": (0.87, 0.92, 0.97),
            "english": (0.99, 0.89, 0.84),
            "gifted": (0.89, 0.87, 0.93),
            "warning": (1, 0.78, 0.81),
        }

        x = left
        for col_no, header in enumerate(headers):
            width = widths[col_no]
            pdf.setFillColorRGB(0.27, 0.45, 0.71)
            pdf.rect(x, table_top - header_height, width, header_height, fill=1, stroke=0)
            pdf.setFillColorRGB(1, 1, 1)
            pdf.setFont(font_name, 9)
            pdf.drawCentredString(x + width / 2, table_top - 18, header)
            x += width

        for period_no in range(10):
            y_top = table_top - header_height - period_no * row_height
            x = left
            pdf.setFillColorRGB(0.95, 0.95, 0.95)
            pdf.rect(x, y_top - row_height, first_width, row_height, fill=1, stroke=0)
            pdf.setFillColorRGB(0, 0, 0)
            pdf.setFont(font_name, 8.5)
            pdf.drawCentredString(x + first_width / 2, y_top - row_height / 2 - 3, PERIOD_LABELS[period_no])
            x += first_width

            for day_index in range(5):
                item = grid[(period_no, day_index)]
                rgb = fill_colors.get(item["kind"], fill_colors["normal"])
                pdf.setFillColorRGB(*rgb)
                pdf.rect(x, y_top - row_height, day_width, row_height, fill=1, stroke=0)
                pdf.setFillColorRGB(0, 0, 0)
                font_size = 7.2
                lines = _pdf_wrapped_lines(item["text"], font_name, font_size, day_width - 8)
                if len(lines) > 5:
                    font_size = 6.4
                    lines = _pdf_wrapped_lines(item["text"], font_name, font_size, day_width - 8)
                lines = lines[:6]
                line_height = font_size + 1.2
                text_y = y_top - (row_height - len(lines) * line_height) / 2 - font_size
                pdf.setFont(font_name, font_size)
                for line in lines:
                    pdf.drawCentredString(x + day_width / 2, text_y, line)
                    text_y -= line_height
                x += day_width

        pdf.setStrokeColorRGB(0.5, 0.5, 0.5)
        x = left
        for width in widths:
            pdf.rect(x, bottom, width, table_height, fill=0, stroke=1)
            x += width
        for row_no in range(11):
            y = table_top - header_height - row_no * row_height
            pdf.line(left, y, page_width - 20, y)
        pdf.line(left, table_top, page_width - 20, table_top)

        pdf.showPage()

    def generate_personal_schedule_pdf(student, grid):
        """製作單一學生、一頁橫式 A4 的個人功課表 PDF。"""
        output = io.BytesIO()
        page_width, page_height = A4[1], A4[0]
        pdf = canvas.Canvas(output, pagesize=(page_width, page_height))
        draw_personal_schedule_pdf_page(pdf, student, grid)
        pdf.save()
        return output.getvalue()

    def generate_all_schedule_pdf(student_records, schedules, rule_map, gifted_class):
        """製作全班共用的多頁 PDF，每位學生固定一頁。"""
        output = io.BytesIO()
        all_errors = []
        page_width, page_height = A4[1], A4[0]
        pdf = canvas.Canvas(output, pagesize=(page_width, page_height))
        for student in student_records:
            grid, errors = make_personal_grid(student, schedules, rule_map, gifted_class)
            all_errors.extend(errors)
            draw_personal_schedule_pdf_page(pdf, student, grid)
        pdf.save()
        return output.getvalue(), all_errors

    st.markdown("#### 步驟一：上傳正式課表與分組名單")
    col_t1, col_t2, col_t3 = st.columns(3)
    with col_t1:
        timetable_file = st.file_uploader(
            "正式課表 PDF（同一跑班班群＋資優班）", type=["pdf"], key="personal_timetable_pdf"
        )
    with col_t2:
        timetable_math_file = st.file_uploader(
            "數理分組名單 PDF", type=["pdf"], key="personal_math_pdf"
        )
    with col_t3:
        timetable_english_file = st.file_uploader(
            "英文分組名單 PDF", type=["pdf"], key="personal_english_pdf"
        )

    if timetable_file and timetable_math_file and timetable_english_file:
        try:
            schedules = parse_timetable_pdf(timetable_file.getvalue())
            math_rows = parse_group_pdf(timetable_math_file.getvalue(), "數理")
            english_rows = parse_group_pdf(timetable_english_file.getvalue(), "英")
            students_df = build_student_group_table(math_rows, english_rows)

            if len(schedules) < 3:
                raise ValueError("正式課表PDF至少需要包含三個跑班班級。")
            if students_df.empty:
                raise ValueError("分組PDF沒有讀到學生資料。")

            schedule_classes = sorted(schedules.keys())
            default_gifted = 914 if 914 in schedule_classes else schedule_classes[-1]
            regular_defaults = [c for c in schedule_classes if c != default_gifted][:3]

            st.success(
                f"已讀取 {len(schedules)} 份正式課表、{len(students_df)} 位學生的分組資料。"
            )

            col_cfg1, col_cfg2 = st.columns(2)
            with col_cfg1:
                cluster_classes = st.multiselect(
                    "本次跑班班群（請選三班）",
                    options=schedule_classes,
                    default=regular_defaults,
                    key="personal_cluster_classes",
                )
            with col_cfg2:
                gifted_class = st.selectbox(
                    "資優班課表",
                    options=schedule_classes,
                    index=schedule_classes.index(default_gifted),
                    key="personal_gifted_class",
                )

            if len(cluster_classes) != 3:
                st.warning("請選擇剛好三個跑班班級。")
                st.stop()
            if gifted_class in cluster_classes:
                st.warning("資優班不能同時列入三個一般跑班班級。")
                st.stop()

            inferred_rules = infer_slot_rules(schedules, cluster_classes, gifted_class)
            rule_df = build_rule_editor_table(schedules, cluster_classes, gifted_class, inferred_rules)
            auto_grouped = rule_df[rule_df["課程類型"] != RULE_NORMAL]

            st.markdown("#### 步驟二：確認哪些節次需要跑班")
            st.caption(
                "自動判斷只看老師與三班／資優班同節課表；科目名稱僅供老師核對，不作為最終依據。"
            )
            count_stem = int((rule_df["課程類型"] == RULE_STEM).sum())
            count_english = int((rule_df["課程類型"] == RULE_ENGLISH).sum())
            st.write(f"目前判定：數理跑班 **{count_stem}節**、英文跑班 **{count_english}節**。")
            st.dataframe(auto_grouped, use_container_width=True, hide_index=True)

            with st.expander("老師校正：分組／不分組節次", expanded=False):
                editable_rule_columns = ["星期", "節次", "課程類型"] + [str(c) for c in cluster_classes] + [str(gifted_class)]
                edited_rule_df = st.data_editor(
                    rule_df[editable_rule_columns],
                    use_container_width=True,
                    hide_index=True,
                    height=520,
                    disabled=["星期", "節次"] + [str(c) for c in cluster_classes] + [str(gifted_class)],
                    column_config={
                        "課程類型": st.column_config.SelectboxColumn(
                            "課程類型", options=RULE_OPTIONS, required=True
                        )
                    },
                    key="personal_rule_editor",
                )

            rule_map = {}
            for _, row in edited_rule_df.iterrows():
                day_index = DAY_NAMES.index(row["星期"])
                period_no = PERIOD_NUMBERS[row["節次"]]
                rule_map[(period_no, day_index)] = row["課程類型"]

            available_home_classes = sorted(
                c for c in students_df["原班"].dropna().astype(int).unique().tolist()
                if c in cluster_classes
            )
            if not available_home_classes:
                available_home_classes = sorted(students_df["原班"].dropna().astype(int).unique().tolist())

            st.markdown("#### 步驟三：老師校正學生分組與資優身分")
            target_class = st.selectbox(
                "要產生哪一班的個人課表？",
                options=available_home_classes,
                index=available_home_classes.index(901) if 901 in available_home_classes else 0,
                key="personal_target_class",
            )

            class_students = students_df[students_df["原班"] == target_class].copy().reset_index(drop=True)
            math_options = ["未設定"] + sorted({r["分組代號"] for r in math_rows})
            english_options = ["未設定"] + sorted({r["分組代號"] for r in english_rows})

            edited_students = st.data_editor(
                class_students[["原班", "座號", "學號", "姓名", "數理去向", "英文去向", "資優類別"]],
                use_container_width=True,
                hide_index=True,
                disabled=["原班", "座號", "學號", "姓名"],
                column_config={
                    "數理去向": st.column_config.SelectboxColumn(
                        "數理去向", options=math_options, required=True
                    ),
                    "英文去向": st.column_config.SelectboxColumn(
                        "英文去向", options=english_options, required=True
                    ),
                    "資優類別": st.column_config.SelectboxColumn(
                        "資優類別",
                        options=["無", "數理資優", "語文資優", "數理／語文資優"],
                        required=True,
                    ),
                },
                key=f"personal_student_editor_{target_class}",
            )

            missing_math = edited_students[edited_students["數理去向"] == "未設定"]
            missing_english = edited_students[edited_students["英文去向"] == "未設定"]
            if not missing_math.empty or not missing_english.empty:
                st.error(
                    f"仍有 {len(missing_math)} 位學生未設定數理去向、"
                    f"{len(missing_english)} 位學生未設定英文去向；請先校正。"
                )
                st.stop()

            st.markdown("#### 步驟四：預覽與下載個人功課表")
            student_labels = {
                f"{int(row['座號']):02d}號 {row['姓名']}": idx
                for idx, row in edited_students.iterrows()
            }
            selected_label = st.selectbox("選擇學生預覽", options=list(student_labels.keys()))
            selected_student = edited_students.loc[student_labels[selected_label]].to_dict()
            selected_grid, selected_errors = make_personal_grid(
                selected_student, schedules, rule_map, gifted_class
            )
            st.dataframe(
                personal_grid_dataframe(selected_grid),
                use_container_width=True,
                height=470,
            )

            if selected_errors:
                with st.expander(f"此學生有 {len(selected_errors)} 個需要確認的地方", expanded=True):
                    for message in selected_errors:
                        st.warning(message)

            selected_excel, selected_excel_errors = generate_schedule_workbook(
                [selected_student], schedules, rule_map, gifted_class
            )
            all_excel, all_excel_errors = generate_schedule_workbook(
                edited_students.to_dict("records"), schedules, rule_map, gifted_class
            )
            all_pdf, all_pdf_errors = generate_all_schedule_pdf(
                edited_students.to_dict("records"), schedules, rule_map, gifted_class
            )

            col_download1, col_download2, col_download3 = st.columns(3)
            with col_download1:
                st.download_button(
                    "下載這位學生的個人功課表",
                    data=selected_excel,
                    file_name=f"{target_class}_{int(selected_student['座號']):02d}_{selected_student['姓名']}_個人功課表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with col_download2:
                st.download_button(
                    f"下載{target_class}全班個人功課表",
                    data=all_excel,
                    file_name=f"{target_class}_全班個人功課表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with col_download3:
                st.download_button(
                    f"下載{target_class}全班個人功課表 PDF",
                    data=all_pdf,
                    file_name=f"{target_class}_全班個人功課表.pdf",
                    mime="application/pdf",
                )

            unique_errors = list(dict.fromkeys(all_excel_errors + all_pdf_errors))
            if unique_errors:
                with st.expander(f"全班共有 {len(unique_errors)} 個資料警告", expanded=False):
                    for message in unique_errors:
                        st.warning(message)

        except Exception as e:
            st.error(f"個人功課表讀取失敗：{e}")
            st.info("請確認PDF是學校正式課表及原始分組名單；若自動判斷不完整，可在校正表修改。")
    else:
        st.info("請先上傳正式課表、數理分組名單及英文分組名單。")
