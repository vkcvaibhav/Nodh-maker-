import streamlit as st
import google.generativeai as genai
import sqlite3
import datetime
from PIL import Image
import io
import PyPDF2
from docx import Document as DocxReader
import urllib.request
import requests
import pandas as pd
import re
import os

# Word Document Generation Imports
from docx import Document
from docx.shared import Mm, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENTATION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Define global logo paths
NAU_LOGO = "logos/nau_logo.png"
ICAR_LOGO = "logos/icar_logo.png"

# ==========================================
# Database Setup for Archiving & Workflow
# ==========================================
DB_FILE = "sadar_nondh_archive.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # Table for Sadar Nondh
    c.execute('''CREATE TABLE IF NOT EXISTS archive 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  date TEXT, month TEXT, year TEXT, subject TEXT, content TEXT)''')
    
    # Table for Purchase Orders (Workflow Tracking)
    c.execute('''CREATE TABLE IF NOT EXISTS purchase_orders 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  vendor_name TEXT, out_no TEXT, date TEXT, amount REAL, status TEXT)''')
    conn.commit()
    conn.close()

def save_to_db(subject, content):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    now = datetime.datetime.now()
    c.execute("INSERT INTO archive (date, month, year, subject, content) VALUES (?, ?, ?, ?, ?)", 
              (now.strftime("%d/%m/%Y"), now.strftime("%m"), now.strftime("%Y"), subject, content))
    conn.commit()
    conn.close()

def save_po_to_db(vendor_name, out_no, date, amount):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO purchase_orders (vendor_name, out_no, date, amount, status) VALUES (?, ?, ?, ?, 'Unfinished')", 
              (vendor_name, out_no, date, amount))
    conn.commit()
    conn.close()

def get_unfinished_pos():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT id, vendor_name, out_no, date, amount FROM purchase_orders WHERE status = 'Unfinished'")
    data = c.fetchall()
    conn.close()
    return data

def mark_po_as_paid(po_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("UPDATE purchase_orders SET status = 'Paid' WHERE id = ?", (po_id,))
    conn.commit()
    conn.close()

def get_archives(month, year, keyword=""):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    query = "SELECT date, subject, content FROM archive WHERE 1=1"
    params = []
    if year != "All":
        query += " AND year=?"
        params.append(year)
    if month != "All":
        query += " AND month=?"
        params.append(month)
    if keyword:
        query += " AND (subject LIKE ? OR content LIKE ?)"
        params.extend([f"%{keyword}%", f"%{keyword}%"])
    c.execute(query, tuple(params))
    data = c.fetchall()
    conn.close()
    return data

init_db()

# ==========================================
# Permanent Attachments & Parsing (GitHub)
# ==========================================
@st.cache_data(ttl=3600) 
def load_permanent_context():
    statute_text = "Statute 121 Rules:\n"
    sample_text = "Sample Nondh Format:\n"
    pdf_url = "https://raw.githubusercontent.com/vkcvaibhav/Nodh-maker-/main/121_Statutes.pdf"
    docx_url = "https://raw.githubusercontent.com/vkcvaibhav/Nodh-maker-/main/sample_nondh.docx"
    try:
        r_pdf = requests.get(pdf_url)
        if r_pdf.status_code == 200:
            f = io.BytesIO(r_pdf.content)
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages: statute_text += page.extract_text() + "\n"
    except Exception: pass
    try:
        r_docx = requests.get(docx_url)
        if r_docx.status_code == 200:
            f = io.BytesIO(r_docx.content)
            doc = DocxReader(f)
            for para in doc.paragraphs: sample_text += para.text + "\n"
    except Exception: pass
    return statute_text, sample_text

def search_sample_nondh(keyword, month, year):
    try: _, sample_text = load_permanent_context()
    except Exception: return []
    if not sample_text: return []
    guj_to_eng = str.maketrans("૦૧૨૩૪૫૬૭૮૯", "0123456789")
    blocks = re.split(r'\n(?=તા\.\s*)', '\n' + sample_text)
    results = []
    for block in blocks:
        block = block.strip()
        if not block or "સાદર નોંધ" not in block: continue
        eng_block = block.translate(guj_to_eng)
        if keyword and keyword.lower() not in block.lower() and keyword.lower() not in eng_block.lower(): continue
        date_match = re.search(r'તા\.\s*([\d/ \-]+)', eng_block)
        date_str = date_match.group(1).strip() if date_match else "Unknown"
        if year != "All" and year not in date_str: continue
        if month != "All":
            if f"/{month}/" not in date_str and f"-{month}-" not in date_str and f"{month}/" not in date_str: continue
        sub_match = re.search(r'વિષય:\s*([^\n]+)', block)
        subject_str = sub_match.group(1).strip() if sub_match else "Historical Reference"
        orig_date_match = re.search(r'તા\.\s*([^\n]+)', block)
        display_date = orig_date_match.group(1).strip() if orig_date_match else "Unknown"
        results.append((display_date + " [Old Sample Ref]", subject_str, block))
    return results

def parse_markdown_to_parts(text):
    lines = text.split('\n')
    pre_text, table_lines, post_text = [], [], []
    in_table, table_done = False, False
    for line in lines:
        if line.strip().startswith('|'):
            in_table = True
            table_lines.append(line)
        else:
            if in_table:
                table_done = True
                in_table = False
            if not table_done: pre_text.append(line)
            else: post_text.append(line)
    df = pd.DataFrame()
    if table_lines:
        parts = table_lines[0].split('|')
        if len(parts) > 2:
            header = [x.strip() for x in parts[1:-1]]
            data = []
            for line in table_lines[2:]: 
                row_parts = line.split('|')
                if len(row_parts) > 2:
                    data.append([x.strip() for x in row_parts[1:-1]])
            if data:
                df = pd.DataFrame(data, columns=header)
                if 'Details' in df.columns: df = df[~df['Details'].astype(str).str.contains('Grand Total', case=False, na=False)]
    return "\n".join(pre_text), df, "\n".join(post_text)

def df_to_markdown_with_total(df):
    if df.empty: return ""
    grand_total = 0
    if 'Total Price' in df.columns:
        grand_total = pd.to_numeric(df['Total Price'], errors='coerce').fillna(0).sum()
    markdown = "|" + "|".join(df.columns) + "|\n|" + "|".join(["---"] * len(df.columns)) + "|\n"
    for _, row in df.iterrows():
        clean_row = [str(int(x)) if isinstance(x, float) and x.is_integer() else str(x) for x in row]
        markdown += "|" + "|".join(clean_row) + "|\n"
    if len(df.columns) >= 5:
        total_row = [""] * len(df.columns)
        total_row[1] = "**Grand Total**"
        total_row[-1] = f"**{grand_total:.2f}**"
        markdown += "|" + "|".join(total_row) + "|\n"
    return markdown

def add_bottom_border(paragraph, size='24'):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), size)
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)
    pPr.append(pBdr)

# ==========================================
# Word Generators
# ==========================================
def create_docx(content):
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Mm(210), Mm(297)
    section.left_margin, section.right_margin = Mm(42), Mm(15)
    section.top_margin, section.bottom_margin = Mm(15), Mm(15)
    
    font = doc.styles['Normal'].font
    font.name = 'Times New Roman'
    font.size = Pt(11)
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Shruti')
    font._element.append(rFonts)

    lines = content.split('\n')
    table_data = []
    in_table = False
    sig_buffer = []

    def flush_signatures():
        if sig_buffer:
            doc.add_paragraph().paragraph_format.space_before = Pt(20)
            sig_table = doc.add_table(rows=1, cols=3)
            for c in sig_table.columns: 
                for cell in c.cells: cell.width = Mm(51)
            for i, sig in enumerate(sig_buffer):
                if i < 3:
                    p = sig_table.cell(0, i).paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    parts = sig.split(',')
                    for j, part in enumerate(parts):
                        run = p.add_run(part.strip())
                        if j < len(parts) - 1: run.add_break()
            sig_buffer.clear()

    def build_and_format_table(data):
        num_cols = len(data[0])
        table = doc.add_table(rows=len(data), cols=num_cols)
        table.style = 'Table Grid'
        widths = [Mm(12), Mm(65), Mm(19), Mm(19), Mm(19), Mm(19)]
        for row_idx, row_data in enumerate(data):
            for col_idx, cell_text in enumerate(row_data):
                cell = table.cell(row_idx, col_idx)
                if num_cols == 6: table.columns[col_idx].width = widths[col_idx]
                is_bold = (row_idx == 0) or ('**' in cell_text)
                cell.text = cell_text.replace('**', '')
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if (col_idx == 1 and row_idx > 0) else WD_ALIGN_PARAGRAPH.CENTER
                if is_bold:
                    for run in p.runs: run.bold = True
        doc.add_paragraph()

    for line in lines:
        line_stripped = line.strip()
        if not line_stripped: continue
        if line_stripped.startswith('|'):
            in_table = True
            if not line_stripped.replace('|', '').replace('-', '').replace(' ', ''): continue
            parts = line_stripped.split('|')
            if len(parts) > 2: table_data.append([cell.strip() for cell in parts[1:-1]])
        else:
            if in_table:
                if table_data: build_and_format_table(table_data)
                table_data, in_table = [], False

            if line_stripped.startswith("તા.") or line_stripped.startswith("સ્થળ:"):
                flush_signatures()
                p = doc.add_paragraph(line_stripped)
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.space_after = Pt(0) 
            elif "સાદર નોંધ" in line_stripped:
                flush_signatures()
                p = doc.add_paragraph()
                p.add_run(line_stripped).bold = True
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                pPr = p._p.get_or_add_pPr()
                pBdr = OxmlElement('w:pBdr')
                bottom = OxmlElement('w:bottom')
                bottom.set(qn('w:val'), 'dotted')
                bottom.set(qn('w:sz'), '4')
                bottom.set(qn('w:space'), '1')
                bottom.set(qn('w:color'), '000000')
                pBdr.append(bottom)
                pPr.append(pBdr)
            elif line_stripped.startswith("વિષય:"):
                flush_signatures()
                p = doc.add_paragraph()
                p.add_run(line_stripped).bold = True
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            elif any(role in line_stripped for role in ["અધિકારી", "ઈન્ચાર્જ", "પ્રાધ્યાપક", "વડા"]) and not any(r in line_stripped for r in ["આચાર્ય", "ડીનશ્રી"]):
                sig_buffer.append(line_stripped)
            elif any(role in line_stripped for role in ["આચાર્ય", "ડીનશ્રી", "મહાવિધાયલય", "ન.કૃ.યુ"]):
                flush_signatures()
                doc.add_paragraph().paragraph_format.space_before = Pt(30)
                p_table = doc.add_table(rows=1, cols=2)
                p_table.columns[0].width, p_table.columns[1].width = Mm(79), Mm(74) 
                formatted_line = "\n".join([p.strip() for p in line_stripped.split(",")])
                p = p_table.cell(0, 1).paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run(formatted_line)
            else:
                flush_signatures()
                doc.add_paragraph(line_stripped).alignment = WD_ALIGN_PARAGRAPH.JUSTIFY 

    flush_signatures()
    if in_table and table_data: build_and_format_table(table_data)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_purchase_order_docx(vendor_name, vendor_address, out_no, po_date, df_items):
    doc = Document()
    for section in doc.sections:
        section.top_margin, section.bottom_margin = Inches(0.4), Inches(0.5)
        section.left_margin, section.right_margin = Inches(0.8), Inches(0.8)
        
    style = doc.styles['Normal']
    style.font.size = Pt(12)
    style.paragraph_format.space_after, style.paragraph_format.space_before = Pt(0), Pt(0)
        
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    table.columns[0].width, table.columns[1].width, table.columns[2].width = Inches(1.8), Inches(3.6), Inches(1.4)
    
    if 'NAU_LOGO' in globals() and NAU_LOGO and os.path.exists(NAU_LOGO):
        cell_left = table.cell(0, 0)
        cell_left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_left = cell_left.paragraphs[0]
        p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_left.add_run().add_picture(NAU_LOGO, width=Inches(1.8))
        
    p_center = table.cell(0, 1).paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_center.paragraph_format.line_spacing = 0.85
    r1 = p_center.add_run("કીટકશાસ્ત્ર વિભાગ\n")
    r1.bold = True
    r1.font.size = Pt(22)
    r2 = p_center.add_run("ન. મ. કૃષિ મહાવિદ્યાલય\nનવસારી કૃષિ યુનિવર્સિટી\nનવસારી- ૩૯૬ ૪૫૦ (ગુજરાત)")
    r2.bold, r2.font.size = True, Pt(14)
    
    if 'ICAR_LOGO' in globals() and ICAR_LOGO and os.path.exists(ICAR_LOGO):
        p_right = table.cell(0, 2).paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_right.add_run().add_picture(ICAR_LOGO, width=Inches(1.5))
        
    p_thick1 = doc.add_paragraph()
    p_thick1.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_thick1.paragraph_format.line_spacing = Pt(1)
    p_thick1.add_run().font.size = Pt(1) 
    add_bottom_border(p_thick1, size='24')
    
    table2 = doc.add_table(rows=1, cols=2)
    table2.columns[0].width, table2.columns[1].width = Inches(3.4), Inches(3.4)

    table2.cell(0,0).paragraphs[0].add_run("ડૉ. સચિન ડી. પટેલ\nપ્રાધ્યાપક અને વડા")
    p2 = table2.cell(0,1).paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run("મોબાઇલ: +૯૧ ૯૪૨૭૮ ૬૭૯૨૫\nઇમેલ: headentonau@gmail.com")
    
    for cell in table2.rows[0].cells:
        for p in cell.paragraphs: p.paragraph_format.line_spacing = 0.8
            
    p_thick2 = doc.add_paragraph()
    p_thick2.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_thick2.paragraph_format.line_spacing = Pt(1)
    p_thick2.add_run().font.size = Pt(1) 
    add_bottom_border(p_thick2, size='24')
    
    letter_year = po_date.split('/')[-1] if '/' in po_date else (po_date.split('.')[-1] if '.' in po_date else "૨૦૨૬")
    
    table3 = doc.add_table(rows=1, cols=2)
    table3.cell(0,0).paragraphs[0].add_run(f"જા.નં. એસીએન/એન્ટો/એઆઈએનપી-એએ/{out_no}/{letter_year}, નવસારી")
    p_date = table3.cell(0,1).paragraphs[0]
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_date.add_run(f"તારીખ: {po_date}")
            
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    p_to = doc.add_paragraph()
    p_to.add_run("પ્રતિ,\n").bold = True
    doc.add_paragraph(vendor_name).runs[0].bold = True
    doc.add_paragraph(vendor_address)
    
    p_subj = doc.add_paragraph()
    p_subj.add_run("વિષય: ખરીદી હુકમ").bold = True
    
    doc.add_paragraph("જય ભારત સહ ઉપરોક્ત વિષય અન્વયે જણાવવાનું કે, અત્રેના કીટકશાસ્ત્ર વિભાગ ખાતે નિચેની વસ્તુઓ બિલ સહિત રજુ કરવા વિનંતી.").alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.add_paragraph() 

    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    
    headers = ["નં.", "વસ્તુઓના નામ", "જથ્થો", "ભાવ પ્રતિ નંગ", "કુલ રકમ"]
    for i, ht in enumerate(headers):
        table.cell(0,i).text = ht
        p = table.cell(0,i).paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        table.cell(0,i).vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    widths = [Inches(0.5), Inches(3.0), Inches(1.0), Inches(1.0), Inches(1.0)]
    for i in range(5): table.columns[i].width = widths[i]

    total_amount = 0.0
    for index, row in df_items.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(index + 1)
        row_cells[1].text = str(row.get('Details', ''))
        row_cells[2].text = str(row.get('Required Quantity', '')) + " " + str(row.get('Available Pkt/Unit', ''))
        
        unit_price = pd.to_numeric(row.get('Unit/Pkt Price', 0), errors='coerce')
        total_price = pd.to_numeric(row.get('Total Price', 0), errors='coerce')
        row_cells[3].text, row_cells[4].text = f"{unit_price:.2f}", f"{total_price:.2f}"
        total_amount += total_price
        
        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    total_row = table.add_row().cells
    total_row[3].text = "Total"
    total_row[3].paragraphs[0].runs[0].bold = True
    total_row[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    total_row[4].text = f"{total_amount:.2f}"
    total_row[4].paragraphs[0].runs[0].bold = True
    total_row[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()
    doc.add_paragraph()

    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig.add_run("પ્રાધ્યાપક અને વડા\nકીટકશાસ્ત્ર વિભાગ").bold = True

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- NEW: Bill Payment Form & Pasting Form ---
def create_bill_payment_form(budget_head, bill_no, bill_date, party_name, amount, amount_words):
    doc = Document()
    for section in doc.sections:
        section.top_margin, section.bottom_margin = Inches(0.8), Inches(0.8)
        section.left_margin, section.right_margin = Inches(1.0), Inches(1.0)
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    style.paragraph_format.space_after = Pt(0)
    
    # Header logic
    p_header = doc.add_paragraph()
    # Left part
    p_header.add_run("No. ACN/ENTO/BILL/       /202\n")
    # Right part
    r_right = p_header.add_run(f"NAVSARI-396450, Date: {datetime.date.today().strftime('%d/%m/%Y')}")
    p_header.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # Custom alignment using spaces to mimic the PDF format strictly
    doc.add_paragraph().paragraph_format.space_after = Pt(10)
    
    doc.add_paragraph("To,")
    doc.add_paragraph("The Principal and Dean,")
    doc.add_paragraph("N.M. College of Agriculture,")
    doc.add_paragraph("Navsari").paragraph_format.space_after = Pt(12)
    
    p_sub = doc.add_paragraph()
    p_sub.add_run("Sub: Submission of bill(s) for payment............").bold = True
    p_sub.paragraph_format.space_after = Pt(12)
    
    p_body = doc.add_paragraph("With reference to the above subject, I am submitting herewith the following bill(s) for making payment to the respective party and debit the same in Budget Head No- ")
    p_body.add_run(budget_head).bold = True
    p_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.add_paragraph()
    
    # Table exactly like the PDF
    table = doc.add_table(rows=2, cols=4)
    table.style = 'Table Grid'
    
    hdr0 = table.rows[0].cells
    hdr1 = table.rows[1].cells
    
    hdr0[0].text = "Sr."
    hdr1[0].text = "No."
    
    hdr0[1].text = "No. of Bill/Date"
    
    hdr0[2].text = "Name of the Party"
    
    hdr0[3].text = "Amount"
    hdr1[3].text = "Rs.              Ps."
    
    for c in [0, 1, 2, 3]:
        if c in [1, 2]: table.cell(0, c).merge(table.cell(1, c))
        for r in [0, 1]:
            p = table.cell(r, c).paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs: run.bold = True
            table.cell(r, c).vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    table.columns[0].width = Inches(0.5)
    table.columns[1].width = Inches(1.5)
    table.columns[2].width = Inches(2.5)
    table.columns[3].width = Inches(1.5)

    # Data row
    row = table.add_row().cells
    row[0].text = "1"
    row[1].text = f"No: {bill_no}\nDt: {bill_date}"
    row[2].text = party_name
    row[3].text = f"{float(amount):.2f}"
    
    for i in range(4): 
        row[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Total Row
    total_row = table.add_row().cells
    table.cell(3,0).merge(table.cell(3,2))
    p_tot = total_row[0].paragraphs[0]
    p_tot.add_run("Total:   ").bold = True
    p_tot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    total_row[3].text = f"{float(amount):.2f}"
    total_row[3].paragraphs[0].runs[0].bold = True
    total_row[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # In Words Row
    word_row = table.add_row().cells
    table.cell(4,0).merge(table.cell(4,3))
    p_word = word_row[0].paragraphs[0]
    p_word.add_run("In words: ").bold = True
    p_word.add_run(f"Rupees {amount_words} Only.")
    
    # Name of Party Row
    party_row = table.add_row().cells
    table.cell(5,0).merge(table.cell(5,3))
    p_party = party_row[0].paragraphs[0]
    p_party.add_run("Name of Party for Payment: ").bold = True
    p_party.add_run(party_name)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(20)
    
    doc.add_paragraph("Encl: Cash/Credit Bill in original")
    doc.add_paragraph(f"No. {bill_no} with entry").paragraph_format.space_after = Pt(20)
    
    doc.add_paragraph(f"Copy F.W.C.S. to M/S: {party_name}").paragraph_format.space_after = Pt(30)
    
    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig.add_run("Professor and Head\nDepartment of Entomology\nN.M.C.A., N.A.U., Navsari").bold = True

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_bill_pasting_form(budget_head, party_name, amount, amount_words):
    doc = Document()
    for section in doc.sections:
        section.top_margin, section.bottom_margin = Inches(0.5), Inches(0.5)
        section.left_margin, section.right_margin = Inches(0.8), Inches(0.8)
    
    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_head.add_run("NAVSARI AGRICULTURAL UNIVERSITY\n").bold = True
    p_head.add_run("DEPARTMENT OF ENTOMOLOGY, NMCA, NAVSARI\n").bold = True
    p_head.add_run("--- BILL PASTING & CERTIFICATE ---").bold = True
    doc.add_paragraph()
    
    p_cert = doc.add_paragraph()
    p_cert.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    cert_text = (
        "Certified that the materials/equipment listed in the attached bill have been received in good condition, "
        "as per the specifications ordered, and have been entered into the Dead Stock/Consumable register. "
        f"Passed for payment of Rs. {float(amount):.2f}/- (Rupees {amount_words} Only) to M/s {party_name} "
        f"under the Budget Head No: {budget_head}."
    )
    p_cert.add_run(cert_text)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(30)
    
    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig.add_run("Professor and Head\nDepartment of Entomology").bold = True
    
    doc.add_paragraph().paragraph_format.space_after = Pt(20)
    
    # Blank box for pasting
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    cell = table.cell(0,0)
    cell.height = Inches(6.0)
    p_box = cell.paragraphs[0]
    p_box.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_box.add_run("\n\n\n\n\n[ PASTE ORIGINAL BILL HERE ]").font.color.rgb = None

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# ==========================================
# Streamlit App UI
# ==========================================
st.set_page_config(page_title="સાદર નોંધ જનરેટર", layout="wide")
st.title("સાદર નોંધ જનરેટર (Intelligent Sadar Nondh App)")

api_key = st.sidebar.text_input("Enter Gemini API Key", type="password")

tab1, tab2, tab3, tab4 = st.tabs([
    "નવી સાદર નોંધ (Create)", 
    "જુની નોંધ (Archives)", 
    "ખરીદી હુકમ (Purchase Order)",
    "બિલ પેમેન્ટ (Bill Payment)"
])

with tab1:
    st.markdown("### જરૂરિયાતની વિગત આપો (Provide Requirements)")
    col1, col2 = st.columns(2)
    with col1:
        text_prompt = st.text_area("તમારી જરૂરિયાત લખો:", placeholder="e.g., need 10 entomological pins...")
    with col2:
        uploaded_image = st.file_uploader("અથવા હાથથી લખેલી ચબરખીનો ફોટો:", type=["jpg", "jpeg", "png"])
    
    if st.button("જનરેટ કરો (Generate)"):
        if not api_key: st.error("Please enter your Gemini API Key in the sidebar.")
        elif not text_prompt and not uploaded_image: st.warning("Please provide either a text requirement or an image.")
        else:
            with st.spinner("સ્ટેચ્યુટ ૧૨૧ ની ચકાસણી અને નોંધ તૈયાર કરવામાં આવી રહી છે..."):
                try:
                    statute_context, sample_context = load_permanent_context()
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel('gemini-3.1-flash-lite-preview')
                    sys_prompt = f"""
                    You are an expert administrative AI for the Department of Entomology, N. M. College of Agriculture, NAU, Navsari.
                    Your task is to generate a formal 'સાદર નોંધ' in Gujarati.
                    [CONTEXT START]
                    {statute_context[:15000]}
                    {sample_context}
                    [CONTEXT END]
                    Format REQUIRED:
                    તા. {datetime.date.today().strftime('%d/%m/%Y')}
                    સ્થળ: નવસારી
                    સાદર નોંધ:
                    વિષય: [Appropriate Subject...]
                    સવિનય ઉપરોક્ત વિષય અન્વયે જણાવવાનું કે, અત્રેનાં કીટકશાસ્ત્ર વિભાગની આઈ.સી.એ.આર. યોજના AINP on Agril Acarology બ.સ. ૩૦૩/૨૦૯૨ અંતર્ગત [Detailed logical reason]. સદર વસ્તુનો કુલ અંદાજિત ખર્ચ [Total Amount] થનાર છે.
                    જે આપ સાહેબશ્રીને સ્ટેચ્યુટ ૧૨૧ની આઈટમ નંબર [DETERMINED_ITEM_NUMBER] મુજબ એનાયત થયેલ સત્તા અનુસાર સૈદ્ધાંતિક મંજુરી આપવા વિનંતી. સદર ખર્ચ અત્રેના વિભાગમાં ચાલતી આઈ.સી.એ.આર યોજના (બ.સ. ૩૦૩/૨૦૯૨) માં કરવામાં આવશે.
                    
                    ખેતીવાડી અધિકારી,કીટકશાસ્ત્ર વિભાગ
                    પ્રોજેકટ ઈન્ચાર્જ,કીટકશાસ્ત્ર વિભાગ
                    પ્રાધ્યાપક અને વડા,કીટકશાસ્ત્ર વિભાગ
                    આચાર્ય અને ડીનશ્રી, ન. મ. કૃષિ મહાવિધાયલય, ન.કૃ.યુ. નવસારી
                    
                    ==== AI STATUTE ANALYSIS ====
                    1. **Original Statute 121 Details:** ...
                    """
                    inputs = [sys_prompt, text_prompt]
                    if uploaded_image: inputs.append(Image.open(uploaded_image))
                    response = model.generate_content(inputs)
                    res_text = response.text
                    if "==== AI STATUTE ANALYSIS ====" in res_text:
                        parts = res_text.split("==== AI STATUTE ANALYSIS ====")
                        st.session_state['generated_nondh'] = parts[0].strip()
                        st.session_state['statute_analysis'] = parts[1].strip()
                    else:
                        st.session_state['generated_nondh'] = res_text.strip()
                        st.session_state['statute_analysis'] = ""
                    st.success("સાદર નોંધ સફળતાપૂર્વક તૈયાર થઈ ગઈ છે!")
                except Exception as e: st.error(f"Error generating document: {e}")

    if 'generated_nondh' in st.session_state:
        if 'statute_analysis' in st.session_state and st.session_state['statute_analysis']:
            with st.expander("🔍 Statute 121 Analysis & Justification (AI Reasoning)", expanded=True):
                st.info("આ વિભાગ ફક્ત તમારી જાણકારી માટે છે અને વર્ડ ડોક્યુમેન્ટ (DOCX) માં પ્રિન્ટ થશે નહીં.")
                st.markdown(st.session_state['statute_analysis'])
        st.markdown("---")
        pre_text, df, post_text = parse_markdown_to_parts(st.session_state['generated_nondh'])
        edit_pre = st.text_area("ઉપરનું લખાણ:", pre_text, height=150)
        
        if not df.empty:
            edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
            if 'Required Quantity' in edited_df.columns and 'Unit/Pkt Price' in edited_df.columns and 'Total Price' in edited_df.columns:
                req_qty = edited_df['Required Quantity'].astype(str).str.extract(r'(\d+\.?\d*)')[0].astype(float).fillna(0)
                unit_price = edited_df['Unit/Pkt Price'].astype(str).str.extract(r'(\d+\.?\d*)')[0].astype(float).fillna(0)
                edited_df['Total Price'] = (req_qty * unit_price).round(2)
                grand_total_calc = edited_df['Total Price'].sum()
                st.success(f"**Grand Total (કુલ રકમ): ₹ {grand_total_calc:,.2f}**")
                edit_pre = re.sub(r'(અંદાજિત ખર્ચ\s*).*?(\s*થનાર)', f'\g<1>{grand_total_calc:,.2f}\g<2>', edit_pre)
        else:
            edited_df = pd.DataFrame()
            
        edit_post = st.text_area("નીચેનું લખાણ:", post_text, height=150)
        final_document = f"{edit_pre}\n\n{df_to_markdown_with_total(edited_df)}\n{edit_post}" if not edited_df.empty else f"{edit_pre}\n\n{edit_post}"
        
        st.markdown("---")
        col_save, col_down = st.columns(2)
        with col_save:
            if st.button("આર્કાઇવમાં સેવ કરો (Save)"):
                subj = "No Subject"
                for line in final_document.split('\n'):
                    if "વિષય:" in line:
                        subj = line.replace("વિષય:", "").strip()
                        break
                save_to_db(subj, final_document)
                st.success("નોંધ સાચવી લેવામાં આવી છે! (હવે તમે Tab 3 માંથી ખરીદી હુકમ બનાવી શકશો)")
        with col_down:
            docx_data = create_docx(final_document)
            st.download_button("Download as Word (DOCX)", data=docx_data, file_name=f"Sadar_Nondh_{datetime.date.today().strftime('%d_%m_%Y')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

with tab2:
    st.markdown("### 🗄️ જૂના રેકોર્ડ શોધો (Archive Search)")
    search_query = st.text_input("🔍 Smart Search (Type in English or Gujarati)")
    db_records = get_archives("All", "All")
    sample_records = search_sample_nondh("", "All", "All")
    all_records = db_records + sample_records

    def smart_search_gemini(query, records):
        if not query.strip(): return records
        model = genai.GenerativeModel('gemini-3.1-pro-preview')
        records_context = ""
        for i, record in enumerate(records):
            if len(record) == 3: date, subject, content = record
            elif len(record) == 2: subject, content = record
            else: continue 
            records_context += f"ID: {i} | Subject: {subject} | Content: {content[:150]}...\n"
        prompt = f"Find matching IDs for: '{query}'. Records: {records_context}. Return ONLY comma-separated IDs (e.g. 0,3). If none, return NONE."
        try:
            result = model.generate_content(prompt).text.strip()
            if result == "NONE" or not result: return []
            matched_ids = [int(x.strip()) for x in result.split(",") if x.strip().isdigit()]
            return [records[i] for i in matched_ids if i < len(records)]
        except: return records

    display_records = smart_search_gemini(search_query, all_records) if search_query else all_records
    if display_records:
        st.success(f"કુલ {len(display_records)} રેકોર્ડ મળ્યા.")
        for idx, record in enumerate(display_records):
            if len(record) == 3: date, subject, content = record
            elif len(record) == 2: subject, content, date = record[0], record[1], "જૂનો રેકોર્ડ"
            with st.expander(f"{date} - {subject}"):
                st.markdown(content)
                st.download_button("Download (Word)", data=create_docx(content), file_name=f"Archive_{idx}.docx", key=f"dl_{idx}") 

with tab3:
    st.markdown("### 📝 ખરીદી હુકમ બનાવો (Generate Purchase Order)")
    st.info("નોંધ મંજૂર થયા પછી સપ્લાયરને ખરીદીનો ઓર્ડર મોકલવા માટે અહીં વિગતો ભરો.")
    
    st.markdown("#### ૧. મંજૂર થયેલ નોંધ પસંદ કરો (Select Approved Nondh)")
    db_records = get_archives("All", "All")
    options = ["-- જાતે માહિતી ભરો (Manual Entry) --"]
    record_dict = {}
    for idx, row in enumerate(db_records):
        if len(row) == 3:
            label = f"[{idx+1}] {row[0]} - {row[1]}"
            record_dict[label] = row[2]
            options.append(label)
    selected_nondh = st.selectbox("અગાઉ સેવ કરેલ નોંધ પસંદ કરો:", options)
    
    # Dedicated section to save the Signed PO
    st.markdown("#### ૨. સહી કરેલ ખરીદી હુકમ અપલોડ કરો (Upload Signed PO - Optional)")
    uploaded_po = st.file_uploader("મંજૂર થયેલ/સહીવાળો ઓર્ડર અપલોડ કરો:", type=["pdf", "jpg", "jpeg", "png"], key="po_up")
    if uploaded_po:
        os.makedirs("signed_pos", exist_ok=True)
        with open(os.path.join("signed_pos", uploaded_po.name), "wb") as f: f.write(uploaded_po.getbuffer())
        st.success("સહી કરેલ ફાઈલ સેવ થઈ ગઈ છે!")

    st.markdown("---")
    st.markdown("#### ૩. સપ્લાયર અને ઓર્ડરની વિગત (Supplier & Order Details)")
    col_po1, col_po2 = st.columns(2)
    with col_po1:
        vendor_name = st.text_input("સપ્લાયરનું નામ (Vendor Name)", value="DUTT ENTERPRISE")
        outward_no = st.text_input("જાવક નંબર (Outward No.)", value="139")
    with col_po2:
        vendor_address = st.text_area("સપ્લાયરનું સરનામું", value="A/5, Krishna complex, borsad chokadi,\nAnand sojitra road, Anand 388 001", height=110)
        po_date = st.date_input("તારીખ (Date)", value=datetime.date.today())
        
    default_df = pd.DataFrame(columns=["Details", "Required Quantity", "Available Pkt/Unit", "Unit/Pkt Price", "Total Price"])
    if selected_nondh != "-- જાતે માહિતી ભરો (Manual Entry) --":
        _, session_df, _ = parse_markdown_to_parts(record_dict[selected_nondh])
        if not session_df.empty: default_df = session_df
    
    po_df = st.data_editor(default_df, num_rows="dynamic", use_container_width=True, key="po_editor")
    
    if st.button("📄 ખરીદી હુકમ ડાઉનલોડ કરો (Download PO & Send to Bill Payment Queue)"):
        if vendor_name and not po_df.empty:
            formatted_date = po_date.strftime("%d.%m.%Y")
            grand_total = pd.to_numeric(po_df['Total Price'], errors='coerce').fillna(0).sum()
            
            po_docx = create_purchase_order_docx(vendor_name, vendor_address, outward_no, formatted_date, po_df)
            
            # Save into Database for Tab 4 workflow
            save_po_to_db(vendor_name, outward_no, formatted_date, grand_total)
            
            st.download_button("Download Purchase Order (DOCX)", data=po_docx, file_name=f"PO_{vendor_name}.docx")
            st.success("ખરીદી હુકમ તૈયાર છે અને પેમેન્ટ માટે Tab 4 માં મોકલી દેવામાં આવ્યો છે!")

# --- NEW TAB 4 ---
with tab4:
    st.markdown("### 💳 બિલ પેમેન્ટ ફોર્મ (Bill Payment & Pasting workflow)")
    st.info("જે ખરીદીના હુકમ (Purchase Orders) માટે બિલ ચૂકવવાનું બાકી છે, તે જ અહીં દેખાશે.")
    
    unfinished_pos = get_unfinished_pos()
    
    if not unfinished_pos:
        st.success("હાલમાં કોઈ બિલ પેમેન્ટ બાકી નથી! (No unfinished purchase orders).")
    else:
        # Create a dictionary to map selectbox labels to PO data
        po_dict = {}
        po_options = []
        for po in unfinished_pos:
            po_id, v_name, o_no, p_date, amt = po
            label = f"PO #{o_no} - {v_name} - ₹{amt} ({p_date})"
            po_options.append(label)
            po_dict[label] = po
            
        # THE MISSING LINE HAS BEEN RESTORED HERE:
        selected_po_label = st.selectbox("પેમેન્ટ માટે ઓર્ડર પસંદ કરો (Select Pending PO):", po_options)
        
        if selected_po_label:
            po_id, v_name, o_no, p_date, amt = po_dict[selected_po_label]
            
            # --- NEW: Reset Session State when switching between different POs ---
            if st.session_state.get("current_po_id") != po_id:
                st.session_state.current_po_id = po_id
                st.session_state.ext_bill_no = "INV-"
                st.session_state.ext_amt = float(amt)
                st.session_state.ext_words = ""
                st.session_state.last_invoice = None
            # ---------------------------------------------------------------------

            st.markdown("#### ઇન્વોઇસ અને બજેટની વિગતો (Invoice Details)")
            
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                budget_head = st.text_input("Budget Head No.", value="303/2092 (AINP on Agril Acarology)")
                
                # CHANGED: Connected to Session State
                bill_no = st.text_input("ઇન્વોઇસ/બિલ નંબર (Vendor Bill No.)", value=st.session_state.ext_bill_no)
                
                invoice_upload = st.file_uploader("પાર્ટીનું બિલ અપલોડ કરો (Upload Vendor Invoice PDF/Img)", type=["pdf", "jpg", "png"])
                
                if invoice_upload:
                    os.makedirs("vendor_invoices", exist_ok=True)
                    with open(os.path.join("vendor_invoices", invoice_upload.name), "wb") as f: 
                        f.write(invoice_upload.getbuffer())
                    st.success("ઇન્વોઇસ સેવ થઈ ગયું!")

                    # --- NEW: AI Auto-Extraction Logic ---
                    if invoice_upload.name != st.session_state.last_invoice and api_key:
                        with st.spinner("AI દ્વારા બિલની વિગતો વાંચવામાં આવી રહી છે... (Extracting...)"):
                            try:
                                import json
                                genai.configure(api_key=api_key)
                                model = genai.GenerativeModel('gemini-3.1-pro-preview') 
                                
                                prompt = """
                                Extract the following from this invoice:
                                1. Invoice/Bill Number
                                2. Grand Total Amount (as a pure number)
                                3. Grand Total Amount in English words (e.g. "Four Thousand Two Hundred Forty Eight")
                                Return ONLY a valid JSON object in this exact format:
                                {"bill_no": "INV-123", "amount": 1234.50, "amount_words": "One Thousand..."}
                                """
                                
                                # Read PDF text or Pass Image directly
                                if invoice_upload.type == "application/pdf":
                                    reader = PyPDF2.PdfReader(invoice_upload)
                                    text = "".join([page.extract_text() for page in reader.pages])
                                    response = model.generate_content([prompt, text])
                                else:
                                    img = Image.open(invoice_upload)
                                    response = model.generate_content([prompt, img])
                                
                                # Parse the JSON response securely
                                res_text = response.text.strip().replace("```json", "").replace("```", "")
                                data = json.loads(res_text)
                                
                                # Update Session State
                                st.session_state.ext_bill_no = str(data.get("bill_no", "INV-"))
                                st.session_state.ext_amt = float(data.get("amount", amt))
                                st.session_state.ext_words = str(data.get("amount_words", ""))
                                st.session_state.last_invoice = invoice_upload.name
                                
                                st.rerun() # Refresh the UI with new values
                                
                            except Exception as e:
                                st.warning(f"આપમેળે વિગત મેળવવામાં ભૂલ: {e}. કૃપા કરીને જાતે ભરો.")
                    # -------------------------------------

            with col_b2:
                bill_date = st.date_input("ઇન્વોઇસની તારીખ (Bill Date)", value=datetime.date.today())
                
                # CHANGED: Connected to Session State
                final_amt = st.number_input("ચૂકવવા પાત્ર રકમ (Amount to Pay)", value=st.session_state.ext_amt)
                amount_words = st.text_input("રકમ શબ્દોમાં (Amount in Words - English)", value=st.session_state.ext_words, placeholder="e.g., Four Thousand Two Hundred Forty Eight")
            
            st.markdown("---")
            col_btn1, col_btn2, col_btn3 = st.columns(3)
            
            with col_btn1:
                # 1. Download Bill Payment Form
                if st.button("📄 Generate Bill Payment Form"):
                    if not amount_words: st.error("Please enter the amount in words!")
                    else:
                        bp_docx = create_bill_payment_form(budget_head, bill_no, bill_date.strftime("%d/%m/%Y"), v_name, final_amt, amount_words)
                        st.download_button("Download Bill Payment Form", data=bp_docx, file_name=f"Payment_Form_{v_name}.docx")
            
            with col_btn2:
                # 2. Download Bill Pasting Form
                if st.button("📑 Generate Bill Pasting Form"):
                    if not amount_words: st.error("Please enter the amount in words!")
                    else:
                        pst_docx = create_bill_pasting_form(budget_head, v_name, final_amt, amount_words)
                        st.download_button("Download Pasting Form", data=pst_docx, file_name=f"Pasting_Form_{v_name}.docx")
            
            with col_btn3:
                # 3. Mark as Paid
                if st.button("✅ બિલ પેમેન્ટ પૂરું કરો (Mark as Paid)"):
                    mark_po_as_paid(po_id)
                    st.success("ઓર્ડર પેમેન્ટ લિસ્ટમાંથી દૂર કરવામાં આવ્યો છે! રિફ્રેશ કરો.")
                    st.rerun()
