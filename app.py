import math
import io
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="成績單產生器", page_icon="📊", layout="centered")

st.title("📊 成績單產生器")
st.caption("上傳成績 Excel，自動排名並輸出格式化報表")

# ════════════════════════════════
#  樣式輔助
# ════════════════════════════════
FONT_NAME = "新細明體"
FILLS = {
    "A++": None,
    "A+":  PatternFill("solid", fgColor="E6E6E6"),
    "A":   PatternFill("solid", fgColor="BFBFBF"),
    "B++": PatternFill("solid", fgColor="808080"),
    "":    PatternFill("solid", fgColor="808080"),
}

def _med(): return Side(style="medium", color="000000")
def _thn(): return Side(style="thin",   color="000000")
def all_thin():
    s = _thn()
    return Border(left=s, right=s, top=s, bottom=s)

def outer_med(r, c, r1, c1, r2, c2):
    return Border(
        left   = _med() if c == c1 else _thn(),
        right  = _med() if c == c2 else _thn(),
        top    = _med() if r == r1 else _thn(),
        bottom = _med() if r == r2 else _thn(),
    )

def sc(ws, row, col, value, bold=False, size=10, fill=None, border=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(name=FONT_NAME, bold=bold, size=size)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if border: c.border = border
    if fill:   c.fill   = fill
    return c

def get_grade(total, th_app, th_ap, th_a, th_bpp):
    if total >= th_app: return "A++"
    if total >= th_ap:  return "A+"
    if total >= th_a:   return "A"
    if total >= th_bpp: return "B++"
    return ""

def read_students(uploaded):
    wb = openpyxl.load_workbook(uploaded, data_only=True)
    students = []
    for row in wb.active.iter_rows(values_only=True):
        if not row[0]: continue
        try:
            students.append((str(row[0]).strip(), float(row[1]), float(row[2]), float(row[3])))
        except (TypeError, ValueError): pass
    return students

# ════════════════════════════════
#  建立報表
# ════════════════════════════════
def build_report(students, exam_lines, th_app, th_ap, th_a, th_bpp):
    sorted_s = sorted(students, key=lambda x: -x[3])
    n = len(sorted_s)

    avg_sel    = round(sum(s  for _, s,  _, _ in sorted_s) / n, 2)
    avg_nonsel = round(sum(ns for _, _, ns, _ in sorted_s) / n, 2)
    avg_total  = round(sum(t  for _, _, _,  t in sorted_s) / n, 2)
    
    counts = {"A++": 0, "A+": 0, "A": 0, "B++": 0}
    for _, _, _, t in sorted_s:
        g = get_grade(t, th_app, th_ap, th_a, th_bpp)
        if g in counts: counts[g] += 1

    # 計算矩形行數 (學生數 + 平均，平分三欄)
    total_slots = n + 1 
    rows_per_block = math.ceil(total_slots / 3)
    if rows_per_block < 2: rows_per_block = 2
    
    HEADER_ROW = 1
    DATA_START = 2
    FINAL_ROW = DATA_START + rows_per_block - 1

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "成績報表"

    # 設定基本欄寬
    for b in range(3):
        base = b * 4 + 1
        ws.column_dimensions[get_column_letter(base)].width   = 9
        ws.column_dimensions[get_column_letter(base+1)].width = 7.5
        ws.column_dimensions[get_column_letter(base+2)].width = 7.5
        ws.column_dimensions[get_column_letter(base+3)].width = 7.5
    
    ws.column_dimensions["M"].width = 0.4
    ws.column_dimensions["N"].width = 7
    ws.column_dimensions["O"].width = 7
    ws.column_dimensions["P"].width = 7

    # 填入標頭
    for b in range(3):
        base = b * 4 + 1
        for i, h in enumerate(["姓名", "選擇", "非選", "總分"]):
            sc(ws, HEADER_ROW, base + i, h, border=all_thin())

    # 填入學生
    for idx, (name, sel, nonsel, total) in enumerate(sorted_s):
        b = idx // rows_per_block
        r = DATA_START + (idx % rows_per_block)
        col = b * 4 + 1
        g = get_grade(total, th_app, th_ap, th_a, th_bpp)
        f = FILLS[g]
        sc(ws, r, col,     name,   fill=f, border=all_thin())
        sc(ws, r, col + 1, sel,    fill=f, border=all_thin())
        sc(ws, r, col + 2, nonsel, fill=f, border=all_thin())
        sc(ws, r, col + 3, total,  fill=f, border=all_thin())

    # 處理平均值 (修正框線遺失問題)
    avg_pos = n
    b_avg = avg_pos // rows_per_block
    r_avg_start = DATA_START + (avg_pos % rows_per_block)
    r_avg_end = FINAL_ROW
    col_avg = b_avg * 4 + 1
    avg_vals = ["平均", avg_sel, avg_nonsel, avg_total]

    for i, val in enumerate(avg_vals):
        curr_col = col_avg + i
        # 重點：在合併前，先對該範圍內所有儲存格畫線
        for fill_r in range(r_avg_start, r_avg_end + 1):
            sc(ws, fill_r, curr_col, "", border=all_thin())
        
        # 執行合併
        if r_avg_end > r_avg_start:
            ws.merge_cells(start_row=r_avg_start, start_column=curr_col, end_row=r_avg_end, end_column=curr_col)
        
        # 填入數值與設定樣式
        sc(ws, r_avg_start, curr_col, val, bold=True, size=12, border=all_thin())

    # 補齊所有空格的邊框 (確保完美矩形)
    for b in range(3):
        base = b * 4 + 1
        for r in range(DATA_START, FINAL_ROW + 1):
            if not ws.cell(row=r, column=base).border:
                for i in range(4):
                    sc(ws, r, base + i, "", border=all_thin())

    # 右側統計資訊
    TITLE_R1, TITLE_R2, TITLE_C1, TITLE_C2 = 2, 7, 14, 16
    ws.merge_cells(start_row=TITLE_R1, start_column=TITLE_C1, end_row=TITLE_R2, end_column=TITLE_C2)
    tc = ws.cell(row=TITLE_R1, column=TITLE_C1)
    tc.value = "\n".join(exam_lines)
    tc.font = Font(name=FONT_NAME, bold=True, size=18)
    tc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(TITLE_R1, TITLE_R2 + 1):
        for c in range(TITLE_C1, TITLE_C2 + 1):
            ws.cell(row=r, column=c).border = outer_med(r, c, TITLE_R1, TITLE_C1, TITLE_R2, TITLE_C2)

    visible = [(g, counts[g]) for g in ["A++", "A+", "A", "B++"] if counts[g] > 0]
    GRADE_R1 = TITLE_R2 + 2
    for i, (g, cnt) in enumerate(visible):
        row = GRADE_R1 + i
        f = FILLS[g]
        for col, val in [(14, g), (15, cnt)]:
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = Font(name=FONT_NAME, bold=True, size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = outer_med(row, col, GRADE_R1, 14, GRADE_R1 + len(visible) - 1, 15)
            if f: cell.fill = f

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, counts, avg_sel, avg_nonsel, avg_total, n

# ════════════════════════════════
#  UI 介面 (保留所有原始功能)
# ════════════════════════════════
with st.expander("⚙️ 考試設定", expanded=True):
    exam_name = st.text_input("考試名稱（以空格分三段）", placeholder="國三 金安模擬考 第六回")
    c1, c2, c3, c4 = st.columns(4)
    th_app = c1.number_input("A++ ≥", value=93.2, step=0.1)
    th_ap  = c2.number_input("A+  ≥", value=85.7, step=0.1)
    th_a   = c3.number_input("A   ≥", value=76.2, step=0.1)
    th_bpp = c4.number_input("B++ ≥", value=67.1, step=0.1)

st.markdown("---")
uploaded = st.file_uploader("📂 上傳成績檔案（xlsx）", type=["xlsx", "xls"])

if uploaded:
    try:
        students = read_students(uploaded)
        st.success(f"✅ 成功讀取 {len(students)} 位學生")
        if st.button("🚀 產生報表", type="primary", use_container_width=True):
            if not exam_name.strip():
                st.error("請填入考試名稱")
            else:
                parts = exam_name.strip().split()
                if   len(parts) >= 3: lines = [parts[0], " ".join(parts[1:-1]), parts[-1]]
                elif len(parts) == 2: lines = [parts[0], "", parts[1]]
                else:                 lines = ["", exam_name.strip(), ""]

                buf, counts, avg_sel, avg_nonsel, avg_total, n = build_report(students, lines, th_app, th_ap, th_a, th_bpp)
                st.download_button("⬇️ 下載 Excel 報表", data=buf, file_name=f"{exam_name}.xlsx", use_container_width=True, type="primary")

                st.markdown("**報表摘要**")
                cols = st.columns(4)
                for col, (g, bg) in zip(cols, [("A++","#dceeff"),("A+","#e6e6e6"),("A","#bfbfbf"),("B++","#aaaaaa")]):
                    col.markdown(f'<div style="background:{bg};border-radius:8px;padding:10px;text-align:center"><div style="font-size:12px;font-weight:600">{g}</div><div style="font-size:24px;font-weight:700">{counts[g]}</div></div>', unsafe_allow_html=True)
                st.markdown(f"平均　選擇 **{avg_sel}**　非選 **{avg_nonsel}**　總分 **{avg_total}**　共 **{n}** 人")
    except Exception as e:
        st.error(f"錯誤：{e}")
