import streamlit as st
import pdfplumber
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io, re

st.set_page_config(page_title="PDF to Excel", page_icon="📄", layout="centered")
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700&family=DM+Mono:wght@400;500&display=swap');
html,body,[class*="css"]{font-family:'Sora',sans-serif;}
.stApp{background:#0f0f13;color:#e8e6e0;}
.hero{text-align:center;padding:48px 0 32px;}
.hero h1{font-size:2.8rem;font-weight:700;color:#f0ede6;letter-spacing:-1px;margin-bottom:8px;}
.hero h1 span{color:#00e5a0;}
.hero p{color:#888;font-size:1rem;margin:0;}
[data-testid="stFileUploader"]{background:#1a1a22;border:1.5px dashed #2e2e3a;border-radius:16px;padding:24px;}
[data-testid="stFileUploader"]:hover{border-color:#00e5a0;}
.stButton>button{background:#00e5a0;color:#0f0f13;font-family:'Sora',sans-serif;font-weight:600;border:none;border-radius:10px;padding:12px 28px;width:100%;}
[data-testid="stDownloadButton"]>button{background:#1a1a22;color:#00e5a0;font-family:'Sora',sans-serif;font-weight:600;border:1.5px solid #00e5a0;border-radius:10px;padding:12px 28px;width:100%;}
[data-testid="stDownloadButton"]>button:hover{background:#00e5a0;color:#0f0f13;}
.stat-card{background:#1a1a22;border:1px solid #2e2e3a;border-radius:12px;padding:20px;text-align:center;}
.stat-num{font-size:2rem;font-weight:700;color:#00e5a0;font-family:'DM Mono',monospace;}
.stat-label{font-size:0.8rem;color:#666;margin-top:4px;text-transform:uppercase;letter-spacing:1px;}
.badge{display:inline-block;padding:4px 12px;border-radius:20px;font-size:0.78rem;font-weight:600;}
.badge-text{background:#0d2e1f;color:#00e5a0;}
.badge-ocr{background:#2e1a0d;color:#ff9f43;}
.section-label{font-size:0.75rem;font-weight:600;color:#555;text-transform:uppercase;letter-spacing:1.5px;margin-bottom:12px;}
hr{border-color:#2e2e3a;}
#MainMenu,footer,header{visibility:hidden;}
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

def is_scanned_pdf(file_bytes):
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if len(text.strip()) > 50:
                return False
    return True

def extract_tables_pdfplumber(file_bytes):
    all_tables = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            for settings in [{}, {"vertical_strategy":"lines","horizontal_strategy":"lines"}, {"vertical_strategy":"text","horizontal_strategy":"text"}]:
                try:
                    tables = page.extract_tables(settings) if settings else page.extract_tables()
                    for tbl in (tables or []):
                        if not tbl or len(tbl) < 2:
                            continue
                        clean = [[str(c).strip() if c else "" for c in row] for row in tbl if any(c and str(c).strip() for c in row)]
                        if len(clean) < 2 or len([h for h in clean[0] if h]) < 2:
                            continue
                        df = pd.DataFrame(clean[1:], columns=clean[0])
                        df = df.replace("", pd.NA).dropna(axis=1, how='all').fillna("")
                        df = df[df.apply(lambda r: any(str(v).strip() for v in r), axis=1)].reset_index(drop=True)
                        if len(df) > 0:
                            all_tables.append({'page': page_num, 'df': df, 'rows': len(df), 'cols': len(df.columns)})
                except:
                    continue
    # Merge same-column tables
    if not all_tables:
        return []
    merged, used = [], set()
    for i, t in enumerate(all_tables):
        if i in used: continue
        cur = t.copy()
        for j, o in enumerate(all_tables):
            if j <= i or j in used: continue
            if list(cur['df'].columns) == list(o['df'].columns):
                cur['df'] = pd.concat([cur['df'], o['df']], ignore_index=True)
                cur['rows'] = len(cur['df'])
                used.add(j)
        merged.append(cur)
        used.add(i)
    return merged

def extract_bank_statement_ocr(file_bytes):
    """Extract bank statement from scanned PDF using OCR + column position detection."""
    import pytesseract
    from pdf2image import convert_from_bytes
    
    images = convert_from_bytes(file_bytes, dpi=300)
    all_rows = []
    
    for page_num, img in enumerate(images, 1):
        width, height = img.size
        data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
        
        # Column x-boundaries (as fraction of page width)
        # TxnDate | ValDate | Description | ChequeNo | Debit | Credit | Balance
        col_bounds = [
            (0,       0.13,   'Transaction Date'),
            (0.13,    0.26,   'Value Date'),
            (0.26,    0.53,   'Description/Narration'),
            (0.53,    0.68,   'Cheque/Reference No.'),
            (0.68,    0.79,   'Debit (₹)'),
            (0.79,    0.90,   'Credit (₹)'),
            (0.90,    1.00,   'Balance (₹)'),
        ]
        
        # Group words by line
        lines = {}
        for i in range(len(data['text'])):
            text = data['text'][i].strip()
            if not text or data['conf'][i] < 25:
                continue
            top  = round(data['top'][i] / 18) * 18
            left = data['left'][i]
            if top not in lines:
                lines[top] = []
            lines[top].append((left, text))
        
        for top in sorted(lines.keys()):
            words = sorted(lines[top], key=lambda x: x[0])
            full_text = ' '.join(w[1] for w in words)
            
            # Only process rows that start with a date
            date_match = re.match(r'^(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})', full_text)
            if not date_match:
                continue
            
            # Assign each word to a column by x position
            row = {col[2]: '' for col in col_bounds}
            for left, text in words:
                x_frac = left / width
                for x0, x1, col_name in col_bounds:
                    if x0 <= x_frac < x1:
                        row[col_name] += ' ' + text
                        break
            
            row = {k: v.strip() for k, v in row.items()}
            
            # Clean number columns
            for col in ['Debit (₹)', 'Credit (₹)', 'Balance (₹)']:
                val = row[col].replace(',', '').replace('|', '').replace('-', '').strip()
                try:
                    row[col] = float(val) if val else None
                except:
                    row[col] = None
            
            # Remove pipe/noise from text columns
            for col in ['Transaction Date', 'Value Date', 'Description/Narration', 'Cheque/Reference No.']:
                row[col] = re.sub(r'[|><=]', '', row[col]).strip()
            
            if row['Transaction Date']:
                all_rows.append(row)
    
    if not all_rows:
        return []
    
    df = pd.DataFrame(all_rows)
    return [{'page': 1, 'df': df, 'rows': len(df), 'cols': len(df.columns)}]

def build_excel(tables_data, filename):
    wb = openpyxl.Workbook()
    hf = PatternFill("solid", fgColor="00E5A0")
    hfont = Font(name="Calibri", bold=True, color="0F0F13", size=11)
    af = PatternFill("solid", fgColor="F2FBF7")
    wf = PatternFill("solid", fgColor="FFFFFF")
    df2 = Font(name="Calibri", size=10)
    bf = Font(name="Calibri", bold=True, size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    right  = Alignment(horizontal="right",  vertical="center")
    thin   = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_ws(ws, df):
        for ci, cn in enumerate(df.columns, 1):
            c = ws.cell(row=1, column=ci, value=str(cn) if cn else f"Col{ci}")
            c.font=hfont; c.fill=hf; c.alignment=center; c.border=border
        for ri, row in enumerate(df.itertuples(index=False), 2):
            fill = af if ri%2==0 else wf
            for ci, val in enumerate(row, 1):
                v = val if val is not None else ""
                c = ws.cell(row=ri, column=ci, value=v)
                c.font=df2; c.fill=fill; c.border=border
                # Right-align numeric columns
                if isinstance(v, float):
                    c.alignment=right
                    c.number_format='#,##0.00'
                else:
                    c.alignment=left
        for col in ws.columns:
            ml = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width=min(ml+4,45)
        ws.freeze_panes="A2"
        ws.auto_filter.ref=ws.dimensions
        ws.row_dimensions[1].height=30

    if tables_data:
        for idx, tbl in enumerate(tables_data, 1):
            sn = f"Table_{idx}"[:31]
            ws = wb.create_sheet(sn)
            style_ws(ws, tbl['df'])
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
    
    # Summary
    ws2 = wb.create_sheet("Summary", 0)
    ws2.column_dimensions["A"].width=25
    ws2.column_dimensions["B"].width=40
    c = ws2.cell(row=1, column=1, value="PDF to Excel — Summary")
    c.font=Font(name="Calibri",bold=True,size=13,color="0F0F13")
    c.fill=hf; ws2.merge_cells("A1:B1"); c.alignment=center; ws2.row_dimensions[1].height=30
    for ri,(k,v) in enumerate([("Source file",filename),("Tables found",str(len(tables_data))),("Total rows",str(sum(t['rows'] for t in tables_data))),("Extracted by","PDF to Excel App")],3):
        ws2.cell(row=ri,column=1,value=k).font=bf
        ws2.cell(row=ri,column=1).fill=af; ws2.cell(row=ri,column=1).alignment=left; ws2.cell(row=ri,column=1).border=border
        ws2.cell(row=ri,column=2,value=v).font=df2
        ws2.cell(row=ri,column=2).fill=wf; ws2.cell(row=ri,column=2).alignment=left; ws2.cell(row=ri,column=2).border=border

    out = io.BytesIO()
    wb.save(out); out.seek(0)
    return out.read()

# ── UI ────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <h1>PDF to <span>Excel</span></h1>
    <p>Upload any PDF — tables extracted instantly, download as .xlsx</p>
</div>
""", unsafe_allow_html=True)

if not OCR_AVAILABLE:
    st.warning("⚠️ OCR not available. Install pytesseract & pdf2image for scanned PDFs.", icon="⚠️")

st.markdown('<div class="section-label">Upload your PDF</div>', unsafe_allow_html=True)
uploaded_files = st.file_uploader("", type=["pdf"], accept_multiple_files=True, label_visibility="collapsed")

if uploaded_files:
    st.markdown("---")
    for uf in uploaded_files:
        file_bytes = uf.read()
        filename   = uf.name

        with st.spinner(f"Processing **{filename}**..."):
            scanned = is_scanned_pdf(file_bytes)
            
            if scanned and OCR_AVAILABLE:
                method = "ocr"
                tables = extract_bank_statement_ocr(file_bytes)
                # fallback to generic table extraction if no rows
                if not tables:
                    tables = extract_tables_pdfplumber(file_bytes)
            else:
                method = "text"
                tables = extract_tables_pdfplumber(file_bytes)
            
            excel_bytes = build_excel(tables, filename)

        st.markdown(f"### 📄 {filename}")
        if method == 'text':
            st.markdown('<span class="badge badge-text">✓ Text PDF — direct extraction</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="badge badge-ocr">⚡ Scanned PDF — OCR used</span>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        total_rows = sum(t['rows'] for t in tables)
        c1,c2,c3 = st.columns(3)
        with c1: st.markdown(f'<div class="stat-card"><div class="stat-num">{len(tables)}</div><div class="stat-label">Tables found</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="stat-card"><div class="stat-num">{total_rows}</div><div class="stat-label">Rows extracted</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="stat-card"><div class="stat-num">{round(len(excel_bytes)/1024,1)}K</div><div class="stat-label">Excel size</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        if tables:
            biggest = max(tables, key=lambda t: t['rows'])
            st.markdown(f'<div class="section-label">Preview ({biggest["rows"]} rows × {biggest["cols"]} columns)</div>', unsafe_allow_html=True)
            st.dataframe(biggest['df'].head(15), use_container_width=True)
        else:
            st.warning("No tables found in this PDF.")

        out_name = filename.replace('.pdf', '.xlsx')
        st.download_button(label=f"⬇ Download {out_name}", data=excel_bytes, file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_{filename}")
        st.markdown("---")
else:
    st.markdown("""
    <div style="text-align:center;padding:48px 0;color:#444;">
        <div style="font-size:3rem;margin-bottom:16px;">📂</div>
        <div style="font-size:1rem;">Drop your PDF above to get started</div>
        <div style="font-size:0.85rem;margin-top:8px;color:#333;">Works with bank statements, invoices, reports and any PDF with tables</div>
    </div>""", unsafe_allow_html=True)

st.markdown('<div style="text-align:center;padding:32px 0 16px;color:#333;font-size:0.78rem;">PDF to Excel App · Built with Streamlit · Free to use</div>', unsafe_allow_html=True)
