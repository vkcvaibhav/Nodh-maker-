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

# Word Document Generation Imports
from docx import Document
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_ORIENTATION
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn

# ==========================================
# Database Setup for Archiving
# ==========================================
DB_FILE = "sadar_nondh_archive.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS archive 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  date TEXT, 
                  month TEXT, 
                  year TEXT, 
                  subject TEXT, 
                  content TEXT)''')
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

# Keyword searching for the database query
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
        # Search inside both subject and content
        query += " AND (subject LIKE ? OR content LIKE ?)"
        params.extend([f"%{keyword}%", f"%{keyword}%"])
        
    c.execute(query, tuple(params))
    data = c.fetchall()
    conn.close()
    return data

init_db()

# ==========================================
# Permanent Attachments (Direct from GitHub)
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
            for page in reader.pages:
                statute_text += page.extract_text() + "\n"
    except Exception as e:
        pass

    try:
        r_docx = requests.get(docx_url)
        if r_docx.status_code == 200:
            f = io.BytesIO(r_docx.content)
            doc = DocxReader(f)
            for para in doc.paragraphs:
                sample_text += para.text + "\n"
    except Exception as e:
        pass
            
    return statute_text, sample_text

# ==========================================
# NEW FEATURE: Parser for searching the Sample DOCX History
# ==========================================
def search_sample_nondh(keyword, month, year):
    try:
        _, sample_text = load_permanent_context()
    except Exception:
        return []
        
    if not sample_text: return []
    
    # Translation table to map Gujarati numbers to English so month/year filters work seamlessly
    guj_to_eng = str.maketrans("૦૧૨૩૪૫૬૭૮૯", "0123456789")
    
    # Split the massive document into individual blocks based on the "તા." prefix
    blocks = re.split(r'\n(?=તા\.\s*)', '\n' + sample_text)
    results = []
    
    for block in blocks:
        block = block.strip()
        if not block or "સાદર નોંધ" not in block: continue
        
        eng_block = block.translate(guj_to_eng)
        
        # 1. Keyword Filter (Check both Gujarati and English variations)
        if keyword and keyword.lower() not in block.lower() and keyword.lower() not in eng_block.lower():
            continue
            
        # 2. Extract Date for Dropdown Filtering
        date_match = re.search(r'તા\.\s*([\d/ \-]+)', eng_block)
        date_str = date_match.group(1).strip() if date_match else "Unknown"
        
        # 3. Year Filter
        if year != "All" and year not in date_str:
            continue
            
        # 4. Month Filter
        if month != "All":
            if f"/{month}/" not in date_str and f"-{month}-" not in date_str and f"{month}/" not in date_str:
                continue
        
        # Extract Subject for the Expander UI Title
        sub_match = re.search(r'વિષય:\s*([^\n]+)', block)
        subject_str = sub_match.group(1).strip() if sub_match else "Historical Reference"
        
        # Extract Original Gujarati Date for UI
        orig_date_match = re.search(r'તા\.\s*([^\n]+)', block)
        display_date = orig_date_match.group(1).strip() if orig_date_match else "Unknown"
        
        # Append with a visual tag to differentiate from Database records
        results.append((display_date + " [Old Sample Ref]", subject_str, block))
        
    return results


# ==========================================
# Table Parsing Helpers for Smart Calculations
# ==========================================
def parse_markdown_to_parts(text):
    """Splits the LLM output into Pre-Text, Dataframe (Table), and Post-Text"""
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
            
            if not table_done:
                pre_text.append(line)
            else:
                post_text.append(line)

    df = pd.DataFrame()
    if table_lines:
        parts = table_lines[0].split('|')
        if len(parts) > 2:
            header = [x.strip() for x in parts[1:-1]]
            data = []
            for line in table_lines[2:]: 
                row_parts = line.split('|')
                if len(row_parts) > 2:
                    row = [x.strip() for x in row_parts[1:-1]]
                    data.append(row)
            if data:
                df = pd.DataFrame(data, columns=header)
                # Remove any hallucinated "Grand Total" row so it doesn't break pandas math
                if 'Details' in df.columns:
                    df = df[~df['Details'].astype(str).str.contains('Grand Total', case=False, na=False)]

    return "\n".join(pre_text), df, "\n".join(post_text)

def df_to_markdown_with_total(df):
    if df.empty: return ""
    
    # Calculate Grand Total Dynamically
    grand_total = 0
    if 'Total Price' in df.columns:
        grand_total = pd.to_numeric(df['Total Price'], errors='coerce').fillna(0).sum()

    markdown = "|" + "|".join(df.columns) + "|\n"
    markdown += "|" + "|".join(["---"] * len(df.columns)) + "|\n"
    
    for _, row in df.iterrows():
        clean_row = [str(int(x)) if isinstance(x, float) and x.is_integer() else str(x) for x in row]
        markdown += "|" + "|".join(clean_row) + "|\n"
        
    # Append the Grand Total Row cleanly
    if len(df.columns) >= 5:
        total_row = [""] * len(df.columns)
        total_row[1] = "**Grand Total**"
        total_row[-1] = f"**{grand_total:.2f}**"
        markdown += "|" + "|".join(total_row) + "|\n"
        
    return markdown

# ==========================================
# Document Generation (A4 Portrait - 20:80 Split)
# ==========================================
def create_docx(content):
    doc = Document()
    section = doc.sections[0]
    
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    
    section.left_margin = Mm(42) 
    section.right_margin = Mm(15)
    section.top_margin = Mm(15)
    section.bottom_margin = Mm(15)
    
    style = doc.styles['Normal']
    font = style.font
    # Set the primary ASCII (English) font
    font.name = 'Times New Roman'
    font.size = Pt(11)
    
    # Configure the Complex Script (CS) font specifically for Gujarati
    # This ensures Word doesn't force Times New Roman on Gujarati characters
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Shruti') # Ensure Shruti handles the Gujarati script
    font._element.append(rFonts)

    lines = content.split('\n')
    table_data = []
    in_table = False
    sig_buffer = []

    def flush_signatures():
        if sig_buffer:
            doc.add_paragraph().paragraph_format.space_before = Pt(20)
            sig_table = doc.add_table(rows=1, cols=3)
            sig_table.autofit = False
            
            for cell in sig_table.columns[0].cells: cell.width = Mm(51)
            for cell in sig_table.columns[1].cells: cell.width = Mm(51)
            for cell in sig_table.columns[2].cells: cell.width = Mm(51)

            for i, sig in enumerate(sig_buffer):
                if i < 3:
                    p = sig_table.cell(0, i).paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    parts = sig.split(',')
                    for j, part in enumerate(parts):
                        run = p.add_run(part.strip())
                        if j < len(parts) - 1:
                            run.add_break()
            sig_buffer.clear()

    def build_and_format_table(data):
        num_cols = len(data[0])
        table = doc.add_table(rows=len(data), cols=num_cols)
        table.style = 'Table Grid'
        table.autofit = False
        table.allow_autofit = False
        
        # Define explicit column widths (Total usable width ~153mm)
        widths = [Mm(12), Mm(65), Mm(19), Mm(19), Mm(19), Mm(19)]
        if num_cols == 6:
            for col_idx in range(6):
                table.columns[col_idx].width = widths[col_idx]

        for row_idx, row_data in enumerate(data):
            for col_idx, cell_text in enumerate(row_data):
                cell = table.cell(row_idx, col_idx)
                if num_cols == 6:
                    cell.width = widths[col_idx]
                
                is_bold = (row_idx == 0) or ('**' in cell_text)
                cell.text = cell_text.replace('**', '')
                p = cell.paragraphs[0]
                
                # Align left for Details, Center for others
                if col_idx == 1 and row_idx > 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                if is_bold:
                    for run in p.runs: run.bold = True
        doc.add_paragraph()

    for line in lines:
        line_stripped = line.strip()
        if not line_stripped: continue

        if line_stripped.startswith('|'):
            in_table = True
            
            if not line_stripped.replace('|', '').replace('-', '').replace(' ', ''):
                continue
                
            parts = line_stripped.split('|')
            if len(parts) > 2:
                row = [cell.strip() for cell in parts[1:-1]]
                table_data.append(row)
        else:
            if in_table:
                if table_data:
                    build_and_format_table(table_data)
                table_data = []
                in_table = False

            if line_stripped.startswith("તા.") or line_stripped.startswith("સ્થળ:"):
                flush_signatures()
                p = doc.add_paragraph(line_stripped)
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                # Reduce space between date and location
                p.paragraph_format.space_after = Pt(0) 
            elif "સાદર નોંધ" in line_stripped:
                flush_signatures()
                p = doc.add_paragraph()
                p.add_run(line_stripped).bold = True
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                # Add thin, black, dotted line under/behind the text
                pPr = p._p.get_or_add_pPr()
                pBdr = OxmlElement('w:pBdr')
                bottom = OxmlElement('w:bottom')
                bottom.set(qn('w:val'), 'dotted')  # Makes the line dotted
                bottom.set(qn('w:sz'), '4')        # Size 4 = thin line (1/2 pt)
                bottom.set(qn('w:space'), '1')     # Spacing from text
                bottom.set(qn('w:color'), '000000')# Black color hex code
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
                p_table.columns[0].width = Mm(79) 
                p_table.columns[1].width = Mm(74) 
                
                parts = line_stripped.split(",")
                formatted_line = "\n".join([p.strip() for p in parts])
                p = p_table.cell(0, 1).paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run(formatted_line)
            else:
                flush_signatures()
                p = doc.add_paragraph(line_stripped)
                # Ensure all Gujarati paragraphs are fully justified
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY 

    flush_signatures()
    
    if in_table and table_data:
        build_and_format_table(table_data)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_purchase_order_docx(vendor_name, vendor_address, out_no, po_date, df_items):
       doc = Document()
        for section in doc.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    style = doc.styles['Normal']
    style.font.size = Pt(12)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)
        
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    table.columns[0].width = Inches(1.8)
    table.columns[1].width = Inches(3.6)
    table.columns[2].width = Inches(1.4)
    
    if 'NAU_LOGO' in globals() and NAU_LOGO and os.path.exists(NAU_LOGO):
        cell_left = table.cell(0, 0)
        cell_left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_left = cell_left.paragraphs[0]
        p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_left = p_left.add_run()
        r_left.add_picture(NAU_LOGO, width=Inches(1.8))
        
    p_center = table.cell(0, 1).paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_center.paragraph_format.line_spacing = 0.85
    r1 = p_center.add_run("કીટકશાસ્ત્ર વિભાગ\n")
    r1.bold = True
    r1.font.size = Pt(22)
    r2 = p_center.add_run("ન. મ. કૃષિ મહાવિદ્યાલય\nનવસારી કૃષિ યુનિવર્સિટી\nનવસારી- ૩૯૬ ૪૫૦ (ગુજરાત)")
    r2.bold = True
    r2.font.size = Pt(14)
    
    if 'ICAR_LOGO' in globals() and ICAR_LOGO and os.path.exists(ICAR_LOGO):
        p_right = table.cell(0, 2).paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_right = p_right.add_run()
        r_right.add_picture(ICAR_LOGO, width=Inches(1.5))
        
    p_thick1 = doc.add_paragraph()
    p_thick1.paragraph_format.space_before = Pt(0)
    p_thick1.paragraph_format.space_after = Pt(0)
    p_thick1.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_thick1.paragraph_format.line_spacing = Pt(1)
    p_thick1.add_run().font.size = Pt(1) 
    add_bottom_border(p_thick1, size='24')
    
    table2 = doc.add_table(rows=1, cols=2)
    table2.autofit = False
    table2.columns[0].width = Inches(3.4)
    table2.columns[1].width = Inches(3.4)

    p1 = table2.cell(0,0).paragraphs[0]
    p1.add_run("ડૉ. સચિન ડી. પટેલ\nપ્રાધ્યાપક અને વડા")
    p2 = table2.cell(0,1).paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run("મોબાઇલ: +૯૧ ૯૪૨૭૮ ૬૭૯૨૫\nઇમેલ: headentonau@gmail.com")
    
    for cell in table2.rows[0].cells:
        for p in cell.paragraphs:
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = 0.8
            
    p_thick2 = doc.add_paragraph()
    p_thick2.paragraph_format.space_before = Pt(0)
    p_thick2.paragraph_format.space_after = Pt(0)
    p_thick2.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_thick2.paragraph_format.line_spacing = Pt(1)
    p_thick2.add_run().font.size = Pt(1) 
    add_bottom_border(p_thick2, size='24')
    
    # Dynamically grab the year from the date for the ref no
    letter_year = letter_date.split('/')[-1] if '/' in letter_date else "૨૦૨૬"
    
    table3 = doc.add_table(rows=1, cols=2)
    p_ref = table3.cell(0,0).paragraphs[0]
    p_ref.add_run(f"જા.નં. એસીએન/એન્ટો/એઆઈએનપી-એએ/{ref_no}/{letter_year}, નવસારી")
    p_date = table3.cell(0,1).paragraphs[0]
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_date.add_run(f"તારીખ: {letter_date}")
            
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    p_to = doc.add_paragraph()
    p_to.add_run("પ્રતિ,\n").bold = True
    doc.add_paragraph(vendor_name).runs[0].bold = True
    doc.add_paragraph(vendor_address)
    p_to.paragraph_format.space_after = Pt(6)

    # --- SUBJECT ---
    p_subj = doc.add_paragraph()
    p_subj.add_run("વિષય: ખરીદી હુકમ").bold = True
    p_to.paragraph_format.space_after = Pt(6)
    
    # --- BODY ---
    p_body = doc.add_paragraph("જય ભારત સહ ઉપરોક્ત વિષય અન્વયે જણાવવાનું કે, અત્રેના કીટકશાસ્ત્ર વિભાગ ખાતે નિચેની વસ્તુઓ બિલ સહિત રજુ કરવા વિનંતી.")
    p_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.add_paragraph() # Spacing

    # --- ITEMS TABLE ---
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    
    # Set headers
    hdr_cells = table.rows[0].cells
    headers = ["નં.", "વસ્તુઓના નામ", "જથ્થો", "ભાવ પ્રતિ નંગ", "કુલ રકમ"]
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Adjust column widths
    widths = [Inches(0.5), Inches(3.0), Inches(1.0), Inches(1.0), Inches(1.0)]
    for i in range(5):
        table.columns[i].width = widths[i]

    # Add Data Rows
    total_amount = 0.0
    for index, row in df_items.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(index + 1)
        row_cells[1].text = str(row.get('Details', ''))
        row_cells[2].text = str(row.get('Required Quantity', '')) + " " + str(row.get('Available Pkt/Unit', ''))
        
        unit_price = pd.to_numeric(row.get('Unit/Pkt Price', 0), errors='coerce')
        total_price = pd.to_numeric(row.get('Total Price', 0), errors='coerce')
        
        row_cells[3].text = f"{unit_price:.2f}"
        row_cells[4].text = f"{total_price:.2f}"
        
        total_amount += total_price
        
        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Add Total Row
    total_row = table.add_row().cells
    total_row[3].text = "Total"
    total_row[3].paragraphs[0].runs[0].bold = True
    total_row[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    total_row[4].text = f"{total_amount:.2f}"
    total_row[4].paragraphs[0].runs[0].bold = True
    total_row[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()
    doc.add_paragraph()

    # --- SIGNATURE ---
    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig.add_run("પ્રાધ્યાપક અને વડા\nકીટકશાસ્ત્ર વિભાગ").bold = True

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()
# ==========================================
# Streamlit App UI
# ==========================================
st.set_page_config(page_title="સાદર નોંધ જનરેટર", layout="wide")
st.title("સાદર નોંધ જનરેટર (Intelligent Sadar Nondh App)")

api_key = st.sidebar.text_input("Enter Gemini API Key", type="password")

tab1, tab2, tab3 = st.tabs(["નવી સાદર નોંધ (Create)", "જુની નોંધ (Archives)", "ખરીદી હુકમ (Purchase Order)"])

with tab1:
    st.markdown("### જરૂરિયાતની વિગત આપો (Provide Requirements)")
    col1, col2 = st.columns(2)
    with col1:
        text_prompt = st.text_area("તમારી જરૂરિયાત લખો:", placeholder="e.g., need 10 entomological pins...")
    with col2:
        uploaded_image = st.file_uploader("અથવા હાથથી લખેલી ચબરખીનો ફોટો:", type=["jpg", "jpeg", "png"])
    
    if st.button("જનરેટ કરો (Generate)"):
        if not api_key:
            st.error("Please enter your Gemini API Key in the sidebar.")
        elif not text_prompt and not uploaded_image:
            st.warning("Please provide either a text requirement or an image.")
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

                    STATUTE 121 ITEM NUMBER DETERMINATION (CRITICAL INSTRUCTION):
                    1. You MUST read the 'Sample Nondh Format' provided in the context to find the historically correct 'આઈટમ નંબર' (Item Number) for this type of purchase.
                    2. Treat the Sample Nondh as the ultimate precedent. If the purchase involves laboratory chemicals, research materials, or AINP scheme-related items, strictly use the exact item number found in the sample (e.g., "૫૪ (i)" or "54 (i)").
                    3. Only rely on the dense 'Statute 121 Rules' PDF text if the item category is completely new and not covered by the sample document precedent. 
                    4. Replace [DETERMINED_ITEM_NUMBER] with the exact number in Gujarati format (like ૫૪ (i)). Do not guess or hallucinate numbers like 45 (ii) (ii).

                    TABLE LOGIC:
                    Analyze the user request to determine if a table is required. 
                    - If the request is a general administrative note WITHOUT specific items to purchase, DO NOT include a table.
                    - If the request involves purchasing, requesting, or listing items with quantities and prices, YOU MUST include a markdown table.
                    
                    If a table is included, the headers MUST strictly be in ENGLISH and use these EXACT columns:
                    Sr. No. | Details | Required Quantity | Available Pkt/Unit | Unit/Pkt Price | Total Price
                    
                    CRITICAL TABLE RULES:
                    1. The 'Details' column MUST contain ONLY the item name without the package size (e.g., "ACETIC ACID GLACIAL 99.5% Extra Pure").
                    2. The 'Available Pkt/Unit' column MUST contain the package size and unit type (e.g., "500 ML", "25 GM", "1 Unit").
                    3. The 'Required Quantity' and 'Unit/Pkt Price' columns MUST contain pure numbers only (e.g., "1" or "335.95"). Do NOT put units like "ML" or "GM" in these two columns.
                    4. Do NOT generate a "Grand Total" row. The system will calculate it automatically.

                    ખેતીવાડી અધિકારી,કીટકશાસ્ત્ર વિભાગ
                    પ્રોજેકટ ઈન્ચાર્જ,કીટકશાસ્ત્ર વિભાગ
                    પ્રાધ્યાપક અને વડા,કીટકશાસ્ત્ર વિભાગ

                    આચાર્ય અને ડીનશ્રી, ન. મ. કૃષિ મહાવિધાયલય, ન.કૃ.યુ. નવસારી
                    
                    ==== AI STATUTE ANALYSIS ====
                    1. **Original Statute 121 Details:** - **Item Number Used:** [State the specific rule number you used].
                       - **Original Statute Text:** [Provide the EXACT quote/sentence directly from the attached Statute 121 PDF for this specific rule number].
                    2. **Justification:** [Explain exactly WHY this specific statute applies to the requested purchase. Relate the items being bought to the statute's wording].
                    3. **Similar Precedent from Sample Nondh:** [Find a similar past purchase in the uploaded 'Sample Nondh' context. List its Subject, Date, and the Statute Item Number it used to prove your choice is historically accurate].
                    4. **Rejected Alternative Statute:** [Find another statute item number from the PDF that looks similar but is INCORRECT (e.g., a rule for furniture instead of chemicals). Quote it and explicitly explain why it is NOT compatible with this purchase].
                    """
                    
                    inputs = [sys_prompt, text_prompt]
                    if uploaded_image:
                        inputs.append(Image.open(uploaded_image))
                        
                    response = model.generate_content(inputs)
                    res_text = response.text
                    
                    # Intercept and Split the AI Response safely
                    if "==== AI STATUTE ANALYSIS ====" in res_text:
                        parts = res_text.split("==== AI STATUTE ANALYSIS ====")
                        st.session_state['generated_nondh'] = parts[0].strip()
                        st.session_state['statute_analysis'] = parts[1].strip()
                    else:
                        st.session_state['generated_nondh'] = res_text.strip()
                        st.session_state['statute_analysis'] = ""
                        
                    st.success("સાદર નોંધ સફળતાપૂર્વક તૈયાર થઈ ગઈ છે!")
                    
                except Exception as e:
                    st.error(f"Error generating document: {e}")

    if 'generated_nondh' in st.session_state:
        
        # FEATURE ADDITION: Display Statute Analysis clearly separated from the Document Draft
        if 'statute_analysis' in st.session_state and st.session_state['statute_analysis']:
            with st.expander("🔍 Statute 121 Analysis & Justification (AI Reasoning)", expanded=True):
                st.info("આ વિભાગ ફક્ત તમારી જાણકારી માટે છે અને વર્ડ ડોક્યુમેન્ટ (DOCX) માં પ્રિન્ટ થશે નહીં.")
                st.markdown(st.session_state['statute_analysis'])
        
        st.markdown("---")
        st.markdown("### ડ્રાફ્ટ એડિટિંગ (Smart Editor)")
        
        pre_text, df, post_text = parse_markdown_to_parts(st.session_state['generated_nondh'])
        
        edit_pre = st.text_area("ઉપરનું લખાણ:", pre_text, height=150)
        
        if not df.empty:
            st.markdown("#### સ્માર્ટ ટેબલ (Smart Table)")
            st.info("નોંધ: 'Required Quantity' અથવા 'Unit/Pkt Price' બદલશો તો 'Total Price' અને લખાણમાં રહેલ 'અંદાજિત ખર્ચ' આપોઆપ બદલાઈ જશે.")
            
            edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
            
            # Intelligent Math Calculation using Regex to strip any accidental text
            if 'Required Quantity' in edited_df.columns and 'Unit/Pkt Price' in edited_df.columns and 'Total Price' in edited_df.columns:
                
                # Extract pure numbers safely
                req_qty = edited_df['Required Quantity'].astype(str).str.extract(r'(\d+\.?\d*)')[0].astype(float).fillna(0)
                unit_price = edited_df['Unit/Pkt Price'].astype(str).str.extract(r'(\d+\.?\d*)')[0].astype(float).fillna(0)
                
                # Perform the calculation
                edited_df['Total Price'] = (req_qty * unit_price).round(2)
                
                # Dynamic Grand Total Calculation
                grand_total_calc = edited_df['Total Price'].sum()
                st.success(f"**Grand Total (કુલ રકમ): ₹ {grand_total_calc:,.2f}**")
                
                # Automatically sync the paragraph text with the accurate Grand Total
                edit_pre = re.sub(r'(અંદાજિત ખર્ચ\s*).*?(\s*થનાર)', f'\g<1>{grand_total_calc:,.2f}\g<2>', edit_pre)
        else:
            edited_df = pd.DataFrame()
            st.info("આ નોંધમાં ટેબલની જરૂરિયાત જણાઈ નથી. (No table required for this note based on the context).")
            
        edit_post = st.text_area("નીચેનું લખાણ:", post_text, height=150)
        
        # Re-stitch using the custom markdown generator that handles the Grand Total
        final_document = f"{edit_pre}\n\n{df_to_markdown_with_total(edited_df)}\n{edit_post}" if not edited_df.empty else f"{edit_pre}\n\n{edit_post}"
        
        st.markdown("---")
        st.markdown("### દસ્તાવેજ પ્રીવ્યુ (Visual Preview - 20/80 Layout)")
        
        with st.container(border=True):
            prev_blank, prev_content = st.columns([2, 8]) 
            with prev_content:
                st.markdown(final_document)
        
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
                st.success("નોંધ સાચવી લેવામાં આવી છે!")
                
        with col_down:
            docx_data = create_docx(final_document)
            st.download_button(label="Download as Word (DOCX)",
                               data=docx_data,
                               file_name=f"Sadar_Nondh_{datetime.date.today().strftime('%d_%m_%Y')}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

with tab2:
    st.markdown("### 🗄️ જૂના રેકોર્ડ શોધો (Archive Search)")
    
    # --- 1. THE SEARCH BAR ---
    search_query = st.text_input("🔍 Smart Search (Type in English or Gujarati)")

    # --- 2. FETCH THE DATA ---
    # Call your database and document functions to actually pull the records
    db_records = get_archives(month="All", year="All")
    sample_records = search_sample_nondh(keyword="", month="All", year="All")

    # Now that the variables hold data, you can safely combine them
    all_records = db_records + sample_records

    # --- 2. THE GEMINI LOGIC ---
    def smart_search_gemini(query, records):
        if not query.strip():
            return records
        
        model = genai.GenerativeModel('gemini-3.1-pro-preview')
        
        records_context = ""
        for i, record in enumerate(records):
            
            # --- SAFE UNPACKING FOR GEMINI ---
            if len(record) == 3:
                date, subject, content = record
            elif len(record) == 2:
                subject, content = record
            else:
                continue # Skip weird formatting
            
            records_context += f"ID: {i} | Subject: {subject} | Content: {content[:150]}...\n"
            
        prompt = f"""
        You are an intelligent, cross-lingual search engine assistant. 
        The user is searching for: "{query}" (This may be in English, but the records are mostly in Gujarati).
        
        Task: Find the records that semantically match the user's intent. Understand the context across both languages (e.g., if the user searches "approval for snacks", match it with "ચા-નાસ્તો મંજુરી", or "seeds" with "બિયારણ").
        
        Available Records:
        {records_context}
        
        Respond ONLY with a comma-separated list of the matching IDs (e.g., 0, 3, 4). If no records match, return NONE.
        """
        
        try:
            response = model.generate_content(prompt)
            result = response.text.strip()
            
            if result == "NONE" or not result:
                return []
                
            matched_ids = [int(x.strip()) for x in result.split(",") if x.strip().isdigit()]
            return [records[i] for i in matched_ids if i < len(records)]
            
        except Exception as e:
            st.error(f"Search error: {e}")
            return records

# --- 3. FILTERING ---
    if search_query:
        with st.spinner("માહિતી શોધાઈ રહી છે... (Searching semantically...)"):
            display_records = smart_search_gemini(search_query, all_records)
    else:
        display_records = all_records

    # --- 4. DISPLAYING THE RESULTS ---
    if display_records:
        st.success(f"કુલ {len(display_records)} રેકોર્ડ મળ્યા.")
        for idx, record in enumerate(display_records):
            
            # --- THE FIX: Handle both 3-item (Database) and 2-item (GitHub) records ---
            if len(record) == 3:
                date, subject, content = record
            elif len(record) == 2:
                subject, content = record
                date = "જૂનો રેકોર્ડ (GitHub)" # Default date for Word docs
            else:
                date, subject, content = "N/A", "Unknown", str(record)
            # --------------------------------------------------------------------------

            with st.expander(f"{date} - {subject}"):
                arc_col1, arc_col2 = st.columns([2, 8])
                with arc_col2:
                    st.markdown(content)
                
                archived_docx = create_docx(content)
                st.download_button(label="Download (Word)",
                                   data=archived_docx,
                                   file_name=f"Archive_Doc_{idx}.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                   key=f"dl_smart_{idx}") 
    else:
        st.info("કોઈ રેકોર્ડ મળેલ નથી (No matching records found).")

with tab3:
    st.markdown("### 📝 ખરીદી હુકમ બનાવો (Generate Purchase Order)")
    st.info("નોંધ મંજૂર થયા પછી સપ્લાયરને ખરીદીનો ઓર્ડર મોકલવા માટે અહીં વિગતો ભરો.")
    
    col_po1, col_po2 = st.columns(2)
    
    with col_po1:
        vendor_name = st.text_input("સપ્લાયરનું નામ (Vendor Name)", value="DUTT ENTERPRISE")
        outward_no = st.text_input("જાવક નંબર (Outward No.)", value="139")
    
    with col_po2:
        vendor_address = st.text_area("સપ્લાયરનું સરનામું (Vendor Address)", value="A/5, Krishna complex, borsad chokadi,\nAnand sojitra road, Anand 388 001", height=110)
        po_date = st.date_input("તારીખ (Date)", value=datetime.date.today())
        
    st.markdown("#### ખરીદીની વસ્તુઓ (Items to Purchase)")
    
    # Try to load the dataframe from Tab 1 if it exists so the user doesn't have to re-type
    default_df = pd.DataFrame(columns=["Details", "Required Quantity", "Available Pkt/Unit", "Unit/Pkt Price", "Total Price"])
    
    if 'generated_nondh' in st.session_state:
        _, session_df, _ = parse_markdown_to_parts(st.session_state['generated_nondh'])
        if not session_df.empty:
            default_df = session_df
            st.success("Tab 1 માંથી વસ્તુઓ આપમેળે લેવામાં આવી છે! (Items imported from Tab 1)")
            
    po_df = st.data_editor(default_df, num_rows="dynamic", use_container_width=True, key="po_editor")
    
    if st.button("📄 ખરીદી હુકમ ડાઉનલોડ કરો (Download PO)"):
        if vendor_name and not po_df.empty:
            try:
                # Format the date nicely for the letter
                formatted_date = po_date.strftime("%d.%m.%Y")
                
                # Generate Document
                po_docx = create_purchase_order_docx(
                    vendor_name=vendor_name,
                    vendor_address=vendor_address,
                    out_no=outward_no,
                    po_date=formatted_date,
                    df_items=po_df
                )
                
                st.download_button(
                    label="Download Purchase Order (DOCX)",
                    data=po_docx,
                    file_name=f"Purchase_Order_{vendor_name.replace(' ', '_')}_{formatted_date}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.success("ખરીદી હુકમ તૈયાર છે! ડાઉનલોડ બટન પર ક્લિક કરો.")
            except Exception as e:
                st.error(f"Error generating document: {e}")
        else:
            st.warning("કૃપા કરીને સપ્લાયરનું નામ અને વસ્તુઓની વિગત ઉમેરો.")
