import math
import io
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ════════════════════════════════
# 🚀 現代化介面 CSS (全功能強化版)
# ════════════════════════════════
st.set_page_config(page_title="成績報表系統", page_icon="📊", layout="centered")

st.markdown("""
    <style>
    /* 全域背景與現代感字體 */
    .stApp {
        background: #F8FAFC;
        font-family: "Microsoft JhengHei", sans-serif;
    }
    
    /* 標題動畫設計 */
    .header-text {
        text-align: center;
        background: linear-gradient(120deg, #1E3A8A 0%, #3B82F6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 900;
        font-size: 3rem !important;
        margin-bottom: 5px;
    }

    /* 玻璃擬態卡片容器 */
    .glass-card {
        background: rgba(255, 255, 255, 0.9);
        padding: 30px;
        border-radius: 20px;
        border: 1px solid rgba(255, 255, 255, 0.3);
        box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.07);
        margin-bottom: 25px;
    }

    /* 按鈕美化 */
    .stButton > button {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        color: white !important;
        border: none !important;
        padding: 12px 30px !important;
        border-radius: 12px !important;
        font-weight: bold !important;
        width: 100%;
        transition: all 0.3s ease !important;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(37, 99, 235, 0.3);
    }
    </style>
    """, unsafe_allow_html=True)

# ════════════════════════════════
# 🛠️ 原始功能邏輯 (保證不刪減)
# ════════════════════════════════

def _med(): return Side(style="medium", color="000000")
def _thn(): return Side(style="thin",   color="000000")
def all_thin(): return Border(left=_thn(), right=_thn(), top=_thn(), bottom=_thn())

def outer_med(r, c, r1, c1, r2, c2):
    return Border(
        left   = _med() if c == c1 else _thn(),
        right  = _med() if c == c2 else _thn(),
        top    = _med() if r == r1 else _thn(),
        bottom = _med() if r == r2 else _thn(),
    )

def sc(ws, row, col, value, bold=False, size=10, fill=None, border=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(name="新細明體", bold=bold, size=size)
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
    ws = wb.active
    students = []
    # 支援 Google 表單下載：跳過 A 欄(時間戳記)，從 B(姓名), C(選擇), D(非選), E(總分) 開始
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 5 or row[1] is None: continue
        try:
            students.append((str(row[1]).strip(), float(row[2]), float(row[3]), float(row[4])))
        except (TypeError, ValueError): continue
    return students

def build_report(students, exam_lines, th_app, th_ap, th_a, th_bpp):
    sorted_s = sorted(students, key=lambda x: -x[3])
    n = len(sorted_s)
    avg_sel = round(sum(s[1] for s in sorted_s) / n, 2)
    avg_nonsel = round(sum(s[2] for s in sorted_s) / n, 2)
    avg_total = round(sum(s[3] for s in sorted_s) / n, 2)
    
    counts = {"A++": 0, "A+": 0, "A": 0, "B++": 0}
    fills = {
        "A++": None, "A+": PatternFill("solid", fgColor="E6E6E6"),
        "A": PatternFill("solid", fgColor="BFBFBF"), "B++": PatternFill("solid", fgColor="808080")
    }

    # 矩形計算邏輯：確保平均區塊至少 2 行
    rows_per_block = math.ceil((n + 1) / 3)
    if (n % rows_per_block != 0) and (rows_per_block - (n % rows_per_block) < 2):
        rows_per_block += 1
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "成績報表"

    # 設定寬度與標頭
    for b in range(3):
        base = b * 4 + 1
        for i, h in enumerate(["姓名", "選擇", "非選", "總分"]):
            sc(ws, 1, base + i, h, border=all_thin())
            ws.column_dimensions[get_column_letter(base + i)].width = 8.5

    # 填入學生資料
    for idx, (name, sel, ns, tot) in enumerate(sorted_s):
        b, r = idx // rows_per_block, 2 + (idx % rows_per_block)
        col, grade = b * 4 + 1, get_grade(tot, th_app, th_ap, th_a, th_bpp)
        f = fills.get(grade)
        for i, v in enumerate([name, sel, ns, tot]):
            sc(ws, r, col + i, v, fill=f, border=all_thin())
            if grade in counts and i == 0: counts[grade] += 1

    # 平均區塊：多行合併並繪製框線
    avg_pos = n
    b_avg, r_avg_start = avg_pos // rows_per_block, 2 + (avg_pos % rows_per_block)
    r_avg_end = 2 + rows_per_block - 1
    col_avg = b_avg * 4 + 1
    for i, val in enumerate(["平均", avg_sel, avg_nonsel, avg_total]):
        curr_c = col_avg + i
        for r_fill in range(r_avg_start, r_avg_end + 1): # 先補線
            sc(ws, r_fill, curr_c, "", border=all_thin())
        if r_avg_end > r_avg_start:
            ws.merge_cells(start_row=r_avg_start, start_column=curr_c, end_row=r_avg_end, end_column=curr_c)
        sc(ws, r_avg_start, curr_c, val, bold=True, size=11, border=all_thin())

    # 補齊空位框線確保矩形
    for b in range(3):
        for r in range(2, r_avg_end + 1):
            if not ws.cell(row=r, column=b*4+1).border:
                for i in range(4): sc(ws, r, b*4+1+i, "", border=all_thin())

    # 右側統計標題 (原創 14, 15, 16 欄位)
    TITLE_R1, TITLE_R2, TITLE_C1, TITLE_C2 = 2, 7, 14, 16
    ws.merge_cells(start_row=TITLE_R1, start_column=TITLE_C1, end_row=TITLE_R2, end_column=TITLE_C2)
    tc = ws.cell(row=TITLE_R1, column=TITLE_C1)
    tc.value, tc.font = "\n".join(exam_lines), Font(name="新細明體", bold=True, size=18)
    tc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(TITLE_R1, TITLE_R2 + 1):
        for c in range(TITLE_C1, TITLE_C2 + 1):
            ws.cell(row=r, column=c).border = outer_med(r, c, TITLE_R1, TITLE_C1, TITLE_R2, TITLE_C2)

    # 級分人數統計
    visible = [(g, counts[g]) for g in ["A++", "A+", "A", "B++"] if counts[g] > 0]
    for i, (g, cnt) in enumerate(visible):
        row = TITLE_R2 + 2 + i
        for col, val in [(14, g), (15, cnt)]:
            cell = sc(ws, row, col, val, bold=True, border=outer_med(row, col, TITLE_R2+2, 14, TITLE_R2+1+len(visible), 15))
            if fills.get(g): cell.fill = fills.get(g)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), counts, avg_sel, avg_nonsel, avg_total, n

# ════════════════════════════════
# 📱 現代感 UI 配置
# ════════════════════════════════

st.markdown('<h1 class="header-text">成績報表產出精靈</h1>', unsafe_allow_html=True)

# 卡片區塊 1: 參數設定
st.markdown('<div class="glass-card">', unsafe_allow_html=True)
st.subheader("⚙️ 考試參數設定")
exam_name = st.text_input("考試全稱（會自動分三行）", value="國三 金安模擬考 第六回")
c1, c2, c3, c4 = st.columns(4)
th_app = c1.number_input("A++ ≥", value=95.0)
th_ap  = c2.number_input("A+ ≥", value=90.0)
th_a   = c3.number_input("A ≥", value=80.0)
th_bpp = c4.number_input("B++ ≥", value=70.0)
st.markdown('</div>', unsafe_allow_html=True)

# 卡片區塊 2: 上傳
st.markdown('<div class="glass-card">', unsafe_allow_html=True)
st.subheader("📂 檔案上傳")
uploaded = st.file_uploader("請上傳 Excel 成績檔", type=["xlsx"])
st.markdown('</div>', unsafe_allow_html=True)

if uploaded:
    students = read_students(uploaded)
    if st.button("🚀 生成報表"):
        # 分行邏輯：維持原始三段拆分
        parts = exam_name.strip().split()
        if len(parts) >= 3: lines = [parts[0], " ".join(parts[1:-1]), parts[-1]]
        elif len(parts) == 2: lines = [parts[0], "", parts[1]]
        else: lines = ["", exam_name.strip(), ""]
        
        content, counts, a_s, a_n, a_t, n = build_report(students, lines, th_app, th_ap, th_a, th_bpp)
        
        st.balloons()
        
        # 下載區卡片
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.success(f"✅ 報表製作完成！已成功分析 {n} 位學生資料。")
        
        # 顯示統計儀表板
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("總人數", f"{n} 人")
        col_m2.metric("總平均", f"{a_t}")
        col_m3.metric("A級以上", f"{counts['A++']+counts['A+']+counts['A']} 人")
        
        st.download_button("📥 點擊下載 Excel 報表", data=content, file_name=f"{exam_name}.xlsx", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
