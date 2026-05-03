import math
import io
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ════════════════════════════════
#  網頁設定與 CSS 注入
# ════════════════════════════════
st.set_page_config(page_title="成績報表產生器", page_icon="📊", layout="centered")

# 自定義 CSS
st.markdown("""
    <style>
    /* 全域背景與字體 */
    .stApp {
        background-color: #f8f9fa;
        font-family: "Microsoft JhengHei", sans-serif;
    }
    
    /* 標題樣式 */
    .main-title {
        font-size: 3rem;
        font-weight: 800;
        color: #1E1E1E;
        text-align: center;
        margin-bottom: 0.5rem;
        background: -webkit-linear-gradient(#1e3c72, #2a5298);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    /* 卡片式容器 */
    div[data-testid="stExpander"], .stFileUploader {
        background-color: white;
        border-radius: 15px;
        border: none;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        padding: 10px;
        margin-bottom: 20px;
    }
    
    /* 按鈕優化 */
    .stButton > button {
        border-radius: 10px;
        height: 3em;
        transition: all 0.3s ease;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        font-weight: bold;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 15px rgba(0,0,0,0.15);
        color: white;
    }
    
    /* 摘要卡片樣式 */
    .summary-card {
        transition: transform 0.3s;
        border: 1px solid rgba(0,0,0,0.05);
    }
    .summary-card:hover {
        transform: scale(1.05);
    }
    </style>
    """, unsafe_allow_html=True)

# 顯示精美標題
st.markdown('<h1 class="main-title">📊 成績報表產生器</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center; color:#666;">專業、自動、精準的成績數據處理工具</p>', unsafe_allow_html=True)

# ════════════════════════════════
#  樣式輔助 (原始邏輯保留)
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
#  建立報表 (包含你要求的矩形與平均邏輯)
# ════════════════════════════════
def build_report(students, exam_lines, th_app, th_ap, th_a, th_bpp):
    sorted_s = sorted(students, key=lambda x: -x[3])
    n = len(sorted_s)
    avg_sel = round(sum(s for _, s, _, _ in sorted_s) / n, 2)
    avg_nonsel = round(sum(ns for _, _, ns, _ in sorted_s) / n, 2)
    avg_total = round(sum(t for _, _, _, t in sorted_s) / n, 2)
    
    counts = {"A++": 0, "A+": 0, "A": 0, "B++": 0}
    for _, _, _, t in sorted_s:
        g = get_grade(t, th_app, th_ap, th_a, th_bpp)
        if g in counts: counts[g] += 1

    # 矩形計算邏輯
    rows_per_block = math.ceil((n + 1) / 3)
    remaining_space = rows_per_block - (n % rows_per_block)
    if (n % rows_per_block != 0) and remaining_space < 2:
        rows_per_block += 1

    HEADER_ROW, DATA_START = 1, 2
    FINAL_ROW = DATA_START + rows_per_block - 1

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "成績報表"

    for b in range(3):
        base = b * 4 + 1
        ws.column_dimensions[get_column_letter(base)].width = 9
        for i in range(1, 4): ws.column_dimensions[get_column_letter(base+i)].width = 7.5

    for b in range(3):
        base = b * 4 + 1
        for i, h in enumerate(["姓名", "選擇", "非選", "總分"]):
            sc(ws, HEADER_ROW, base + i, h, border=all_thin())

    for idx, (name, sel, nonsel, total) in enumerate(sorted_s):
        b, r = idx // rows_per_block, DATA_START + (idx % rows_per_block)
        col, g = b * 4 + 1, get_grade(total, th_app, th_ap, th_a, th_bpp)
        sc(ws, r, col, name, fill=FILLS[g], border=all_thin())
        sc(ws, r, col + 1, sel, fill=FILLS[g], border=all_thin())
        sc(ws, r, col + 2, nonsel, fill=FILLS[g], border=all_thin())
        sc(ws, r, col + 3, total, fill=FILLS[g], border=all_thin())

    # 平均區塊 (確保多行且有框線)
    avg_pos = n
    b_avg = avg_pos // rows_per_block
    r_avg_start = DATA_START + (avg_pos % rows_per_block)
    col_avg = b_avg * 4 + 1
    avg_vals = ["平均", avg_sel, avg_nonsel, avg_total]

    for i, val in enumerate(avg_vals):
        curr_col = col_avg + i
        for fill_r in range(r_avg_start, FINAL_ROW + 1):
            sc(ws, fill_r, curr_col, "", border=all_thin())
        if FINAL_ROW > r_avg_start:
            ws.merge_cells(start_row=r_avg_start, start_column=curr_col, end_row=FINAL_ROW, end_column=curr_col)
        sc(ws, r_avg_start, curr_col, val, bold=True, size=12, border=all_thin())

    # 補齊矩形邊框
    for b in range(3):
        base = b * 4 + 1
        for r in range(DATA_START, FINAL_ROW + 1):
            if not ws.cell(row=r, column=base).border:
                for i in range(4): sc(ws, r, base + i, "", border=all_thin())

    # 右側統計標題
    TITLE_R1, TITLE_R2, TITLE_C1, TITLE_C2 = 2, 7, 14, 16
    ws.merge_cells(start_row=TITLE_R1, start_column=TITLE_C1, end_row=TITLE_R2, end_column=TITLE_C2)
    tc = ws.cell(row=TITLE_R1, column=TITLE_C1)
    tc.value, tc.font = "\n".join(exam_lines), Font(name=FONT_NAME, bold=True, size=18)
    tc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(TITLE_R1, TITLE_R2 + 1):
        for c in range(TITLE_C1, TITLE_C2 + 1):
            ws.cell(row=r, column=c).border = outer_med(r, c, TITLE_R1, TITLE_C1, TITLE_R2, TITLE_C2)

    visible = [(g, counts[g]) for g in ["A++", "A+", "A", "B++"] if counts[g] > 0]
    for i, (g, cnt) in enumerate(visible):
        row = TITLE_R2 + 2 + i
        for col, val in [(14, g), (15, cnt)]:
            cell = sc(ws, row, col, val, bold=True, size=11, border=outer_med(row, col, TITLE_R2+2, 14, TITLE_R2+1+len(visible), 15))
            if FILLS[g]: cell.fill = FILLS[g]

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, counts, avg_sel, avg_nonsel, avg_total, n

# ════════════════════════════════
#  UI 流程
# ════════════════════════════════
with st.expander("⚙️ 參數設定", expanded=True):
    exam_name = st.text_input("考試名稱", placeholder="例如：國三 金安模擬考 第六回")
    c1, c2, c3, c4 = st.columns(4)
    th_app = c1.number_input("A++", value=95.0)
    th_ap  = c2.number_input("A+", value=90.0)
    th_a   = c3.number_input("A", value=80.0)
    th_bpp = c4.number_input("B++", value=70.0)

uploaded = st.file_uploader("📂 上傳成績 Excel", type=["xlsx"])

if uploaded:
    students = read_students(uploaded)
    st.info(f"已讀取 {len(students)} 名學生資料")
    if st.button("🚀 開始產生美化報表", use_container_width=True):
        parts = exam_name.strip().split()
        if len(parts) >= 3: lines = [parts[0], " ".join(parts[1:-1]), parts[-1]]
        elif len(parts) == 2: lines = [parts[0], "", parts[1]]
        else: lines = ["", exam_name.strip(), ""]

        buf, counts, avg_sel, avg_nonsel, avg_total, n = build_report(students, lines, th_app, th_ap, th_a, th_bpp)
        
        st.success("報表產生完成！")
        st.download_button("⬇️ 立即下載 Excel 檔案", data=buf, file_name=f"{exam_name}.xlsx", use_container_width=True)

        # 美化的摘要統計
        st.markdown("### 📊 本次考試概況")
        cols = st.columns(4)
        for col, (g, bg, txt) in zip(cols, [("A++","#E3F2FD","#0D47A1"),("A+","#F5F5F5","#424242"),("A","#EEEEEE","#212121"),("B++","#E0E0E0","#000000")]):
            col.markdown(f"""
                <div class="summary-card" style="background:{bg}; color:{txt}; border-radius:12px; padding:15px; text-align:center;">
                    <small>{g}</small><br><b style="font-size:24px;">{counts[g]}</b> 人
                </div>
            """, unsafe_allow_html=True)
        
        st.divider()
        st.markdown(f"""
            <div style="background:white; padding:20px; border-radius:15px; box-shadow: 0 2px 10px rgba(0,0,0,0.05);">
                平均數據：選擇 <b>{avg_sel}</b> | 非選 <b>{avg_nonsel}</b> | 總分 <b>{avg_total}</b> | 總人數 <b>{n}</b>
            </div>
        """, unsafe_allow_html=True)
