import streamlit as st
import pdfplumber
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re

st.set_page_config(page_title="PDF to Excel", page_icon="📄", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700&family=DM+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'Sora', sans-serif; }
.stApp { background: #0f0f13; color: #e8e6e0; }
.hero { text-align: center; padding: 48px 0 32px; }
.hero h1 { font-size: 2.8rem; font-weight: 700; color: #f0ede6; letter-spacing: -1px; margin-bottom: 8px; }
.hero h1 span { color: #00e5a0; }
.hero p { color: #888; font-size: 1rem; margin: 0; }
[data-testid="stFileUploader"] { background: #1a1a22; border: 1.5px dashed #2e2e3a; border-radius: 16px; padding: 24px; }
[data-testid="stFileUploader"]:hover { border-color: #00e5a0; }
.stButton > button { background: #00e5a0; color: #0f0f13; font-family: 'Sora', sans-serif; font-weight: 600; border: none; border-radius: 10px; padding: 12px 28px; width: 100%; }
[data-testid="stDownloadButton"] > button { background: #1a1a22; color: #00e5a0; font-family: 'Sora', sans-serif; font-weight: 600; border: 1.5px solid #00e5a0; border-radius: 10px; padding: 12px 28px; width: 100%; }
[data-testid="stDownloadButton"] > button:hover { background: #00e5a0; color: #0f0f13; }
.stat-card { background: #1a1a22; border: 1px solid #2e2e3a; border-radius: 12px; padding: 20px; text-align: center; }
.stat-num { font-size: 2rem; font-weight: 700; color: #00e5a0; font-family: 'DM Mono', monospace; }
.stat-label { font-size: 0.8rem; color: #666; margin-top: 4px; text-transform: uppercase; letter-spacing: 1px; }
.badge { display: inline-block; padding: 4px 12px; border-radius: 20px; font-size: 0.78rem; font-weight: 600; }
.badge-text { background: #0d2e1f; color: #00e5a0; }
.badge-ocr { background: #2e1a0d; color: #ff9f43; }
.section-label { font-size: 0.75rem; font-weight: 600; color: #555; text-transform: uppercase; letter-spacing: 1.5px; margin-bottom: 12px; }
hr { border-color: #2e2e3a; }
#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


def is_ocr_available():
    try:
        import pytesseract
        from pdf2image import convert_from_bytes
        return True
    except:
        return False

OCR_AVAILABLE = is_ocr_available()


def extract_text(file_bytes):
    text = ""
    method = "text"
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
    if len(text.strip()) < 50:
        if OCR_AVAILABLE:
            import pytesseract
            from pdf2image import convert_from_bytes
            method = "ocr"
            text = ""
            images = convert_from_bytes(file_bytes, dpi=300)
            for img in images:
                text += pytesseract.image_to_string(img)
        else:
            method = "ocr_unavailable"
    return text, method


def extract_smart_tables(file_bytes):
    all_tables = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            settings_list = [
                {},
                {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
                {"vertical_strategy": "text",  "horizontal_strategy": "text"},
            ]
            best_tables = []
            for settings in settings_list:
                try:
                    tables = page.extract_tables(settings) if settings else page.extract_tables()
                    if tables and len(tables) > len(best_tables):
                        best_tables = tables
                except:
                    continue

            for tbl in best_tables:
                if not tbl or len(tbl) < 2:
                    continue
                clean_rows = []
                for row in tbl:
                    if row and any(cell and str(cell).strip() for cell in row):
                        clean_row = [str(cell).strip() if cell else "" for cell in row]
                        clean_rows.append(clean_row)
                if len(clean_rows) < 2:
                    continue
                header = clean_rows[0]
                data   = clean_rows[1:]
                valid_headers = [h for h in header if h]
                if len(valid_headers) < 2:
                    continue
                df = pd.DataFrame(data, columns=header)
                df = df.replace("", pd.NA).dropna(axis=1, how='all').fillna("")
                df = df[df.apply(lambda r: any(str(v).strip() for v in r), axis=1)]
                df = df.reset_index(drop=True)
                if len(df) == 0:
                    continue
                all_tables.append({
                    'page': page_num,
                    'df':   df,
                    'rows': len(df),
                    'cols': len(df.columns)
                })

    # Merge same-column tables across pages
    if not all_tables:
        return all_tables
    merged = []
    used = set()
    for i, tbl in enumerate(all_tables):
        if i in used:
            continue
        current = tbl.copy()
        for j, other in enumerate(all_tables):
            if j <= i or j in used:
                continue
            if list(current['df'].columns) == list(other['df'].columns):
                current['df']   = pd.concat([current['df'], other['df']], ignore_index=True)
                current['rows'] = len(current['df'])
                used.add(j)
        merged.append(current)
        used.add(i)
    return merged


def build_excel(tables_data, text, filename):
    wb = openpyxl.Workbook()
    header_fill = PatternFill("solid", fgColor="00E5A0")
    header_font = Font(name="Calibri", bold=True, color="0F0F13", size=11)
    alt_fill    = PatternFill("solid", fgColor="F2FBF7")
    white_fill  = PatternFill("solid", fgColor="FFFFFF")
    data_font   = Font(name="Calibri", size=10)
    bold_font   = Font(name="Calibri", bold=True, size=10)
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left        = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin        = Side(style="thin", color="D0D0D0")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_sheet(ws, df):
        for col_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_idx, value=str(col_name) if col_name else f"Col{col_idx}")
            cell.font = header_font; cell.fill = header_fill
            cell.alignment = center; cell.border = border
        for row_idx, row in enumerate(df.itertuples(index=False), 2):
            fill = alt_fill if row_idx % 2 == 0 else white_fill
            for col_idx, val in enumerate(row, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=str(val) if val is not None else "")
                cell.font = data_font; cell.fill = fill
                cell.alignment = left; cell.border = border
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 45)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 30

    if tables_data:
        for idx, tbl in enumerate(tables_data, 1):
            sheet_name = f"Table_{idx}_P{tbl['page']}"[:31]
            ws = wb.create_sheet(sheet_name)
            style_sheet(ws, tbl['df'])
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
    else:
        ws = wb.active
        ws.title = "Extracted Data"
        for col_name, col_idx in [("Line No", 1), ("Content", 2)]:
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = header_font; cell.fill = header_fill
            cell.alignment = center; cell.border = border
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        for row_idx, line in enumerate(lines, 2):
            fill = alt_fill if row_idx % 2 == 0 else white_fill
            ws.cell(row=row_idx, column=1, value=row_idx-1).font = data_font
            ws.cell(row=row_idx, column=1).fill = fill
            ws.cell(row=row_idx, column=1).alignment = center
            ws.cell(row=row_idx, column=1).border = border
            ws.cell(row=row_idx, column=2, value=line).font = data_font
            ws.cell(row=row_idx, column=2).fill = fill
            ws.cell(row=row_idx, column=2).alignment = left
            ws.cell(row=row_idx, column=2).border = border
        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 80

    # Summary sheet
    ws_sum = wb.create_sheet("Summary", 0)
    ws_sum.column_dimensions["A"].width = 25
    ws_sum.column_dimensions["B"].width = 40
    ws_sum.cell(row=1, column=1, value="PDF to Excel — Summary").font = Font(name="Calibri", bold=True, size=13, color="0F0F13")
    ws_sum.cell(row=1, column=1).fill = header_fill
    ws_sum.merge_cells("A1:B1")
    ws_sum.cell(row=1, column=1).alignment = center
    ws_sum.row_dimensions[1].height = 30
    for row_idx, (key, val) in enumerate([("Source file", filename), ("Tables found", str(len(tables_data))), ("Sheets created", str(len(wb.sheetnames)-1)), ("Extracted by", "PDF to Excel App")], 3):
        ws_sum.cell(row=row_idx, column=1, value=key).font  = bold_font
        ws_sum.cell(row=row_idx, column=1).fill             = alt_fill
        ws_sum.cell(row=row_idx, column=1).alignment        = left
        ws_sum.cell(row=row_idx, column=1).border           = border
        ws_sum.cell(row=row_idx, column=2, value=val).font  = data_font
        ws_sum.cell(row=row_idx, column=2).fill             = white_fill
        ws_sum.cell(row=row_idx, column=2).alignment        = left
        ws_sum.cell(row=row_idx, column=2).border           = border

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()


# ═══════════════════════════════════════════════════
#  UI
# ═══════════════════════════════════════════════════

st.markdown("""
<div class="hero">
    <h1>PDF to <span>Excel</span></h1>
    <p>Upload any PDF — tables extracted instantly, download as .xlsx</p>
</div>
""", unsafe_allow_html=True)

if not OCR_AVAILABLE:
    st.warning("⚠️ OCR not available — scanned PDFs may not extract well.", icon="⚠️")

st.markdown('<div class="section-label">Upload your PDF</div>', unsafe_allow_html=True)

uploaded_files = st.file_uploader("", type=["pdf"], accept_multiple_files=True, label_visibility="collapsed")

if uploaded_files:
    st.markdown("---")
    for uploaded_file in uploaded_files:
        file_bytes = uploaded_file.read()
        filename   = uploaded_file.name
        with st.spinner(f"Processing **{filename}**..."):
            text, method = extract_text(file_bytes)
            tables       = extract_smart_tables(file_bytes)
            excel_bytes  = build_excel(tables, text, filename)

        st.markdown(f"### 📄 {filename}")
        if method == 'text':
            st.markdown('<span class="badge badge-text">✓ Text PDF — direct extraction</span>', unsafe_allow_html=True)
        elif method == 'ocr':
            st.markdown('<span class="badge badge-ocr">⚡ Scanned PDF — OCR used</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="badge badge-ocr">⚠ Scanned PDF — OCR not available</span>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        total_rows = sum(t['rows'] for t in tables) if tables else len([l for l in text.split('\n') if l.strip()])
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f'<div class="stat-card"><div class="stat-num">{len(tables)}</div><div class="stat-label">Tables found</div></div>', unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="stat-card"><div class="stat-num">{total_rows}</div><div class="stat-label">Rows extracted</div></div>', unsafe_allow_html=True)
        with col3:
            st.markdown(f'<div class="stat-card"><div class="stat-num">{round(len(excel_bytes)/1024,1)}K</div><div class="stat-label">Excel size</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        if tables:
            biggest = max(tables, key=lambda t: t['rows'])
            st.markdown(f'<div class="section-label">Preview — Table ({biggest["rows"]} rows × {biggest["cols"]} columns)</div>', unsafe_allow_html=True)
            st.dataframe(biggest['df'].head(15), use_container_width=True)
        else:
            st.markdown('<div class="section-label">Preview — extracted text</div>', unsafe_allow_html=True)
            st.text_area("", value=text[:800], height=180, label_visibility="collapsed")

        out_name = filename.replace('.pdf', '.xlsx')
        st.download_button(label=f"⬇ Download {out_name}", data=excel_bytes, file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_{filename}")
        st.markdown("---")

else:
    st.markdown("""
    <div style="text-align:center; padding: 48px 0; color: #444;">
        <div style="font-size: 3rem; margin-bottom: 16px;">📂</div>
        <div style="font-size: 1rem;">Drop your PDF above to get started</div>
        <div style="font-size: 0.85rem; margin-top: 8px; color: #333;">Works with bank statements, invoices, reports and any PDF with tables</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown('<div style="text-align:center;padding:32px 0 16px;color:#333;font-size:0.78rem;">PDF to Excel App · Built with Streamlit · Free to use</div>', unsafe_allow_html=True)
