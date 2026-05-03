import math
import io
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="成績報表產生器", page_icon="📊", layout="centered")

# ════════════════════════════════
#  深色 / 淺色模式切換
# ════════════════════════════════
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

# ════════════════════════════════
#  全域 CSS 注入
# ════════════════════════════════
def inject_css(dark: bool):
    if dark:
        bg          = "#0f1117"
        surface     = "#1a1d27"
        surface2    = "#22263a"
        border_col  = "#2e3250"
        text        = "#e8ecf4"
        text_sub    = "#8892b0"
        accent      = "#4f8ef7"
        accent_glow = "rgba(79,142,247,0.18)"
        accent2     = "#63d8b4"
        btn_bg      = "#1e2235"
        btn_hover   = "#2a3050"
        input_bg    = "#1a1d27"
        tag_bg      = "#252a40"
        success_bg  = "rgba(99,216,180,0.12)"
        success_col = "#63d8b4"
        err_bg      = "rgba(255,100,100,0.1)"
        err_col     = "#ff6464"
        toggle_icon = "☀️"
        toggle_tip  = "切換為淺色模式"
        shadow      = "0 8px 32px rgba(0,0,0,0.45)"
        divider     = "rgba(255,255,255,0.06)"
    else:
        bg          = "#f4f6fb"
        surface     = "#ffffff"
        surface2    = "#f0f3fa"
        border_col  = "#dce2f0"
        text        = "#1a1f36"
        text_sub    = "#5a6480"
        accent      = "#2563eb"
        accent_glow = "rgba(37,99,235,0.12)"
        accent2     = "#059669"
        btn_bg      = "#eef2ff"
        btn_hover   = "#dde5ff"
        input_bg    = "#f8faff"
        tag_bg      = "#eef2ff"
        success_bg  = "rgba(5,150,105,0.08)"
        success_col = "#059669"
        err_bg      = "rgba(220,38,38,0.07)"
        err_col     = "#dc2626"
        toggle_icon = "🌙"
        toggle_tip  = "切換為深色模式"
        shadow      = "0 4px 24px rgba(30,50,120,0.10)"
        divider     = "rgba(0,0,0,0.07)"

    css = f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

    :root {{
        --bg:         {bg};
        --surface:    {surface};
        --surface2:   {surface2};
        --border:     {border_col};
        --text:       {text};
        --text-sub:   {text_sub};
        --accent:     {accent};
        --accent-g:   {accent_glow};
        --accent2:    {accent2};
        --btn-bg:     {btn_bg};
        --btn-hover:  {btn_hover};
        --input-bg:   {input_bg};
        --tag-bg:     {tag_bg};
        --shadow:     {shadow};
        --divider:    {divider};
    }}

    /* ── Base ── */
    html, body, .stApp {{
        background: var(--bg) !important;
        color: var(--text) !important;
        font-family: 'Noto Sans TC', sans-serif !important;
    }}
    .block-container {{
        padding-top: 2rem !important;
        padding-bottom: 3rem !important;
        max-width: 760px !important;
    }}

    /* ── Header ── */
    .app-header {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 0.25rem;
    }}
    .app-title {{
        font-size: 1.7rem;
        font-weight: 700;
        color: var(--text);
        letter-spacing: -0.03em;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }}
    .app-title .icon-badge {{
        background: var(--accent);
        color: #fff;
        border-radius: 10px;
        width: 38px; height: 38px;
        display: inline-flex; align-items: center; justify-content: center;
        font-size: 1.1rem;
        box-shadow: 0 2px 12px var(--accent-g);
    }}
    .app-subtitle {{
        color: var(--text-sub);
        font-size: 0.85rem;
        margin-top: 0.15rem;
        margin-bottom: 1.5rem;
        letter-spacing: 0.01em;
    }}

    /* ── Divider ── */
    .styled-divider {{
        height: 1px;
        background: var(--divider);
        margin: 1.25rem 0;
        border: none;
    }}

    /* ── Card ── */
    .card {{
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 16px;
        padding: 1.5rem 1.75rem;
        box-shadow: var(--shadow);
        margin-bottom: 1.25rem;
        transition: box-shadow 0.2s;
    }}
    .card:hover {{
        box-shadow: 0 8px 40px var(--accent-g);
    }}
    .card-label {{
        font-size: 0.7rem;
        font-weight: 600;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: var(--accent);
        margin-bottom: 0.85rem;
        display: flex; align-items: center; gap: 0.4rem;
    }}
    .card-label::after {{
        content: '';
        flex: 1;
        height: 1px;
        background: var(--border);
        margin-left: 0.5rem;
    }}

    /* ── Expander ── */
    [data-testid="stExpander"] {{
        background: var(--surface) !important;
        border: 1px solid var(--border) !important;
        border-radius: 16px !important;
        box-shadow: var(--shadow) !important;
        overflow: hidden;
    }}
    [data-testid="stExpander"] summary {{
        color: var(--text) !important;
        font-weight: 600 !important;
        font-size: 0.92rem !important;
        padding: 1rem 1.25rem !important;
        background: var(--surface2) !important;
    }}
    [data-testid="stExpander"] summary:hover {{
        background: var(--btn-hover) !important;
    }}
    [data-testid="stExpander"] > div > div {{
        padding: 1.2rem 1.25rem 1.4rem !important;
        background: var(--surface) !important;
    }}

    /* ── Inputs & Selects ── */
    .stTextInput input,
    .stNumberInput input {{
        background: var(--input-bg) !important;
        border: 1.5px solid var(--border) !important;
        border-radius: 10px !important;
        color: var(--text) !important;
        font-family: 'Noto Sans TC', sans-serif !important;
        font-size: 0.92rem !important;
        padding: 0.55rem 0.85rem !important;
        transition: border-color 0.2s, box-shadow 0.2s !important;
    }}
    .stTextInput input:focus,
    .stNumberInput input:focus {{
        border-color: var(--accent) !important;
        box-shadow: 0 0 0 3px var(--accent-g) !important;
        outline: none !important;
    }}
    .stTextInput label,
    .stNumberInput label {{
        color: var(--text-sub) !important;
        font-size: 0.78rem !important;
        font-weight: 500 !important;
        letter-spacing: 0.03em !important;
    }}

    /* ── Buttons ── */
    .stButton > button {{
        background: var(--btn-bg) !important;
        color: var(--accent) !important;
        border: 1.5px solid var(--border) !important;
        border-radius: 10px !important;
        font-family: 'Noto Sans TC', sans-serif !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
        padding: 0.55rem 1.5rem !important;
        transition: all 0.2s !important;
        letter-spacing: 0.02em !important;
    }}
    .stButton > button:hover {{
        background: var(--btn-hover) !important;
        border-color: var(--accent) !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 16px var(--accent-g) !important;
    }}
    /* Primary button */
    .stButton > button[kind="primary"] {{
        background: var(--accent) !important;
        color: #ffffff !important;
        border-color: var(--accent) !important;
    }}
    .stButton > button[kind="primary"]:hover {{
        filter: brightness(1.1) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px var(--accent-g) !important;
    }}
    /* Download button */
    .stDownloadButton > button {{
        background: var(--accent) !important;
        color: #fff !important;
        border: none !important;
        border-radius: 10px !important;
        font-family: 'Noto Sans TC', sans-serif !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
        padding: 0.65rem 1.5rem !important;
        transition: all 0.2s !important;
        letter-spacing: 0.02em !important;
    }}
    .stDownloadButton > button:hover {{
        filter: brightness(1.12) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px var(--accent-g) !important;
    }}

    /* ── File Uploader ── */
    [data-testid="stFileUploaderDropzone"] {{
        background: var(--surface2) !important;
        border: 2px dashed var(--border) !important;
        border-radius: 14px !important;
        padding: 2rem !important;
        transition: border-color 0.2s, background 0.2s !important;
    }}
    [data-testid="stFileUploaderDropzone"]:hover {{
        border-color: var(--accent) !important;
        background: var(--accent-g) !important;
    }}
    [data-testid="stFileUploaderDropzone"] * {{
        color: var(--text-sub) !important;
    }}
    [data-testid="stFileUploaderDropzoneInstructions"] span {{
        color: var(--text-sub) !important;
        font-size: 0.88rem !important;
    }}

    /* ── Alerts ── */
    [data-testid="stAlert"] {{
        border-radius: 12px !important;
        border: none !important;
        font-size: 0.88rem !important;
    }}
    .stSuccess {{
        background: {success_bg} !important;
        color: {success_col} !important;
        border-left: 3px solid {success_col} !important;
    }}
    .stError {{
        background: {err_bg} !important;
        color: {err_col} !important;
        border-left: 3px solid {err_col} !important;
    }}

    /* ── Summary metric tiles ── */
    .metric-row {{
        display: flex;
        gap: 0.75rem;
        margin: 1rem 0 0.75rem;
    }}
    .metric-tile {{
        flex: 1;
        border-radius: 12px;
        padding: 0.85rem 0.6rem;
        text-align: center;
        border: 1px solid var(--border);
        background: var(--surface2);
        transition: transform 0.15s;
    }}
    .metric-tile:hover {{ transform: translateY(-2px); }}
    .metric-tile .grade-label {{
        font-size: 0.7rem;
        font-weight: 700;
        letter-spacing: 0.08em;
        color: var(--text-sub);
        margin-bottom: 0.3rem;
    }}
    .metric-tile .grade-count {{
        font-size: 1.8rem;
        font-weight: 800;
        color: var(--text);
        line-height: 1;
    }}

    .avg-row {{
        display: flex;
        gap: 1.5rem;
        flex-wrap: wrap;
        background: var(--surface2);
        border: 1px solid var(--border);
        border-radius: 12px;
        padding: 0.85rem 1.25rem;
        margin-top: 0.75rem;
        font-size: 0.88rem;
        color: var(--text-sub);
    }}
    .avg-row strong {{ color: var(--accent); font-weight: 700; }}
    .avg-sep {{
        color: var(--border);
        font-weight: 300;
    }}

    /* ── Threshold badges ── */
    .threshold-row {{
        display: flex; gap: 0.6rem; flex-wrap: wrap; margin-top: 0.5rem;
    }}
    .threshold-badge {{
        font-size: 0.72rem;
        font-weight: 600;
        padding: 0.2rem 0.6rem;
        border-radius: 999px;
        background: var(--tag-bg);
        color: var(--text-sub);
        border: 1px solid var(--border);
        font-family: 'DM Mono', monospace;
    }}

    /* ── Mode toggle button ── */
    .mode-toggle-wrap {{
        display: flex;
        justify-content: flex-end;
        margin-bottom: -0.5rem;
    }}

    /* ── Section label ── */
    .section-label {{
        font-size: 0.7rem;
        font-weight: 700;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: var(--accent);
        margin-bottom: 0.5rem;
        margin-top: 1.25rem;
    }}

    /* ── Misc overrides ── */
    .stMarkdown p {{ color: var(--text) !important; }}
    [data-testid="stCaptionContainer"] p {{ color: var(--text-sub) !important; }}
    hr {{ border-color: var(--divider) !important; }}

    /* hide streamlit branding */
    #MainMenu, footer, header {{ visibility: hidden !important; }}
    </style>
    """
    return css, toggle_icon, toggle_tip

dark = st.session_state.dark_mode
css_str, toggle_icon, toggle_tip = inject_css(dark)
st.markdown(css_str, unsafe_allow_html=True)

# ════════════════════════════════
#  樣式輔助 (原始功能，未修改)
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
#  建立報表 (原始功能，未修改)
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

    rows_per_block = math.ceil((n + 1) / 3)
    remaining_space = rows_per_block - (n % rows_per_block)
    
    if (n % rows_per_block != 0) and remaining_space < 2:
        rows_per_block += 1
    elif (n % rows_per_block == 0):
        rows_per_block = rows_per_block

    HEADER_ROW = 1
    DATA_START = 2
    FINAL_ROW = DATA_START + rows_per_block - 1

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "成績報表"

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

    for b in range(3):
        base = b * 4 + 1
        for i, h in enumerate(["姓名", "選擇", "非選", "總分"]):
            sc(ws, HEADER_ROW, base + i, h, border=all_thin())

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

    avg_pos = n
    b_avg = avg_pos // rows_per_block
    r_avg_start = DATA_START + (avg_pos % rows_per_block)
    r_avg_end = FINAL_ROW
    col_avg = b_avg * 4 + 1
    avg_vals = ["平均", avg_sel, avg_nonsel, avg_total]

    for i, val in enumerate(avg_vals):
        curr_col = col_avg + i
        for fill_r in range(r_avg_start, r_avg_end + 1):
            sc(ws, fill_r, curr_col, "", border=all_thin())
        if r_avg_end > r_avg_start:
            ws.merge_cells(start_row=r_avg_start, start_column=curr_col, end_row=r_avg_end, end_column=curr_col)
        sc(ws, r_avg_start, curr_col, val, bold=True, size=12, border=all_thin())

    for b in range(3):
        base = b * 4 + 1
        for r in range(DATA_START, FINAL_ROW + 1):
            if not ws.cell(row=r, column=base).border:
                for i in range(4):
                    sc(ws, r, base + i, "", border=all_thin())

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
#  美化 UI
# ════════════════════════════════

# 模式切換按鈕（右上角）
toggle_col1, toggle_col2 = st.columns([8, 1])
with toggle_col2:
    if st.button(toggle_icon, help=toggle_tip, key="mode_btn"):
        st.session_state.dark_mode = not st.session_state.dark_mode
        st.rerun()

# 標題區
st.markdown("""
<div class="app-title">
  <span class="icon-badge">📊</span>
  成績報表產生器
</div>
<div class="app-subtitle">上傳成績 Excel，自動排名並輸出格式化報表</div>
""", unsafe_allow_html=True)

# ── 考試設定 ──
with st.expander("⚙️　考試設定", expanded=True):
    exam_name = st.text_input(
        "考試名稱（以空格分三段）",
        placeholder="國三 金安模擬考 第六回",
        label_visibility="visible"
    )
    st.markdown('<div style="height:0.6rem"></div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    th_app = c1.number_input("A++ 門檻", value=93.2, step=0.1, format="%.1f")
    th_ap  = c2.number_input("A+  門檻", value=85.7, step=0.1, format="%.1f")
    th_a   = c3.number_input("A   門檻", value=76.2, step=0.1, format="%.1f")
    th_bpp = c4.number_input("B++ 門檻", value=67.1, step=0.1, format="%.1f")

    # 視覺化門檻預覽標籤
    st.markdown(
        f'<div class="threshold-row">'
        f'<span class="threshold-badge">A++ ≥ {th_app}</span>'
        f'<span class="threshold-badge">A+ ≥ {th_ap}</span>'
        f'<span class="threshold-badge">A ≥ {th_a}</span>'
        f'<span class="threshold-badge">B++ ≥ {th_bpp}</span>'
        f'</div>',
        unsafe_allow_html=True
    )

st.markdown('<hr class="styled-divider">', unsafe_allow_html=True)

# ── 上傳區 ──
st.markdown('<div class="section-label">📂 上傳成績檔案</div>', unsafe_allow_html=True)
uploaded = st.file_uploader("上傳 Excel 成績檔案（xlsx）", type=["xlsx", "xls"], label_visibility="collapsed")

if uploaded:
    try:
        students = read_students(uploaded)
        st.success(f"✅　成功讀取 **{len(students)}** 位學生資料")

        st.markdown('<div style="height:0.5rem"></div>', unsafe_allow_html=True)

        if st.button("🚀　產生報表", type="primary", use_container_width=True):
            if not exam_name.strip():
                st.error("⚠️　請先填入考試名稱")
            else:
                parts = exam_name.strip().split()
                if   len(parts) >= 3: lines = [parts[0], " ".join(parts[1:-1]), parts[-1]]
                elif len(parts) == 2: lines = [parts[0], "", parts[1]]
                else:                 lines = ["", exam_name.strip(), ""]

                buf, counts, avg_sel, avg_nonsel, avg_total, n = build_report(
                    students, lines, th_app, th_ap, th_a, th_bpp
                )

                st.markdown('<hr class="styled-divider">', unsafe_allow_html=True)

                # 下載按鈕
                st.download_button(
                    "⬇️　下載 Excel 報表",
                    data=buf,
                    file_name=f"{exam_name}.xlsx",
                    use_container_width=True,
                    type="primary"
                )

                # 等級磚塊
                st.markdown('<div class="section-label" style="margin-top:1.25rem">📋 報表摘要</div>', unsafe_allow_html=True)

                grade_styles = {
                    "A++": ("linear-gradient(135deg,#dbeafe,#bfdbfe)", "#1d4ed8"),
                    "A+":  ("#e9e9e9", "#444"),
                    "A":   ("#d4d4d4", "#333"),
                    "B++": ("#b0b0b0", "#222"),
                }
                tiles_html = '<div class="metric-row">'
                for g, (bg_tile, col_tile) in grade_styles.items():
                    tiles_html += (
                        f'<div class="metric-tile" style="background:{bg_tile};border-color:{col_tile}22">'
                        f'  <div class="grade-label" style="color:{col_tile}">{g}</div>'
                        f'  <div class="grade-count" style="color:{col_tile}">{counts[g]}</div>'
                        f'</div>'
                    )
                tiles_html += '</div>'
                st.markdown(tiles_html, unsafe_allow_html=True)

                # 平均列
                st.markdown(
                    f'<div class="avg-row">'
                    f'<span>📐 選擇題均分　<strong>{avg_sel}</strong></span>'
                    f'<span class="avg-sep">|</span>'
                    f'<span>非選均分　<strong>{avg_nonsel}</strong></span>'
                    f'<span class="avg-sep">|</span>'
                    f'<span>總分均分　<strong>{avg_total}</strong></span>'
                    f'<span class="avg-sep">|</span>'
                    f'<span>共　<strong>{n}</strong> 人</span>'
                    f'</div>',
                    unsafe_allow_html=True
                )

    except Exception as e:
        st.error(f"❌　讀取失敗：{e}")
