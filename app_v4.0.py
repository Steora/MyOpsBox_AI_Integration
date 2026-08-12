# --- STYLED STREAMLIT APP ENGINE FOR MY OPSBOX ARCHIVER ---
import io
import os
import tempfile
import subprocess
import platform
import PyPDF2
from datetime import datetime, timedelta
import requests
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

# --- GEMINI MODEL CONFIGURATION ---
PRIMARY_GEMINI_MODEL = "gemini-3.5-flash"
FALLBACK_GEMINI_MODEL = "gemini-3.5-pro-latest"

# --- BRAND PALETTE DEFINITIONS ---
HEX_PURPLE_MAIN = "642795"      
HEX_WINE_SECONDARY = "7E0380"   
HEX_LIGHT_BG = "F7F8FA"         

COLOR_PURPLE_MAIN = RGBColor(100, 39, 149)
COLOR_WINE_SECONDARY = RGBColor(126, 3, 128)
COLOR_TEXT_DARK = RGBColor(40, 40, 40)
COLOR_SECONDARY = RGBColor(115, 115, 115) 

# --- CONFIGURATION & CONSTANTS ---
FATHOM_API_URL = "https://api.fathom.ai/external/v1/meetings".strip()

# LOGO & SIGNATURE ASSET PATHS
MY_OPSBOX_LOGO_URL = r"MyOpsBox.svg"  
MY_OPSBOX_LOGO_PNG = r"MyOpsBox.png"
STEORA_LOGO_URL = r"Steora.svg"       
PATTI_SIG_PNG = r"Patricia Zapparolli Signature.png" 
MAIN_COLOR_LOGO_PNG = r"Main Color Logo.png"


def call_gemini_api(prompt_text, api_key, system_instruction=None, temperature=0.25):
    """Sends a prompt to Google's Gemini REST API and returns the clean generated string."""
    if not api_key:
        raise ValueError("Gemini API key is missing.")

    models_to_try = [PRIMARY_GEMINI_MODEL, FALLBACK_GEMINI_MODEL]
    last_exception = None

    for model in models_to_try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        payload = {
            "generationConfig": {"temperature": temperature},
            "contents": [{"role": "user", "parts": [{"text": prompt_text}]}],
        }

        if system_instruction:
            payload["systemInstruction"] = {"parts": [{"text": system_instruction}]}

        try:
            response = requests.post(url, headers={"Content-Type": "application/json"}, json=payload, timeout=180)
            if response.status_code == 200:
                data = response.json()
                return data["candidates"][0]["content"]["parts"][0]["text"]
            else:
                last_exception = RuntimeError(f"Gemini API ({model}) error {response.status_code}: {response.text}")
        except Exception as err:
            last_exception = err

    raise last_exception or RuntimeError("Failed to communicate with Gemini API endpoints.")

def clean_ai_markdown(text):
    if not text: return ""
    text = text.strip()
    if text.startswith("```markdown"): text = text[11:]
    elif text.startswith("```"): text = text[3:]
    if text.endswith("```"): text = text[:-3]
    return text.strip()

def extract_pdf_text(uploaded_file):
    if uploaded_file is None: return ""
    try:
        reader = PyPDF2.PdfReader(uploaded_file)
        return "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
    except Exception as e:
        st.error(f"Error extracting PDF: {e}")
        return ""

# --- XML HELPER FUNCTIONS ---
def set_cell_background(cell, fill_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{margin}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="D3D3D3"):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tblBorders.append(border)
    tblPr.append(tblBorders)

# --- GENERIC MARKDOWN PARSER FOR DOCX ---
def parse_markdown_to_docx(doc, text_content, primary_color=COLOR_PURPLE_MAIN, primary_hex=HEX_PURPLE_MAIN):
    """Parses clean markdown and injects it into a docx Document object."""
    lines = text_content.split('\n')
    idx = 0
    while idx < len(lines):
        line = lines[idx].strip()
        if not line or line.startswith("---") or line.startswith("___") or line.startswith(":---"):
            idx += 1
            continue

        if line.startswith("|"):
            table_rows = []
            while idx < len(lines) and lines[idx].strip().startswith("|"):
                row_raw = lines[idx].strip()
                if not (":---" in row_raw or "---:" in row_raw):
                    row_cells = [c.strip() for c in row_raw.split("|")[1:-1]]
                    if row_cells: table_rows.append(row_cells)
                idx += 1

            if table_rows:
                max_cols = max(len(r) for r in table_rows)
                grid_table = doc.add_table(rows=0, cols=max_cols)
                grid_table.style = 'Normal Table'
                set_table_borders(grid_table)
                is_metadata_card = (max_cols == 4)

                for r_num, row_data in enumerate(table_rows):
                    # Skip rendering the row entirely if it's an empty markdown header (e.g., | | | )
                    if r_num == 0 and all(not c.strip() for c in row_data):
                        continue

                    tr = grid_table.add_row()
                    trPr = tr._tr.get_or_add_trPr()
                    cantSplit = OxmlElement('w:cantSplit')
                    trPr.append(cantSplit)
                    
                    row_cells = tr.cells
                    is_header_row = (r_num == 0 and not is_metadata_card and not all(not c.strip() for c in table_rows[0]))

                    if is_header_row:
                        tblHeader = OxmlElement('w:tblHeader')
                        trPr.append(tblHeader)

                    for c_num, val in enumerate(row_data):
                        if c_num < len(row_cells):
                            cell = row_cells[c_num]
                            set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
                            cell_p = cell.paragraphs[0]
                            cell_p.paragraph_format.space_after = Pt(0)
                            
                            # Safely extract bolding from markdown table cells
                            is_bold = "**" in val
                            clean_val = val.replace("**", "").strip()
                            
                            run = cell_p.add_run(clean_val)
                            run.font.name = 'Montserrat' if (run.font.name and 'Montserrat' in run.font.name) else 'Arial'

                            if is_header_row:
                                set_cell_background(cell, primary_hex)
                                run.font.bold = True; run.font.size = Pt(10); run.font.color.rgb = RGBColor(255, 255, 255)
                            else:
                                run.font.size = Pt(9.5)
                                run.font.bold = is_bold
                                
                                if is_metadata_card and c_num in [0, 2]:
                                    run.font.bold = True; run.font.color.rgb = primary_color
                                else:
                                    run.font.color.rgb = COLOR_TEXT_DARK

                                if not is_metadata_card and r_num % 2 == 1 and max_cols > 2:
                                    set_cell_background(cell, HEX_LIGHT_BG)
                doc.add_paragraph().paragraph_format.space_after = Pt(6)
            continue

        # Intercept the exact Signature Block Marker for PSA
        if "[SIGNATURE_BLOCK]" in line:
            doc.add_paragraph().paragraph_format.space_before = Pt(24)
            sig_p = doc.add_paragraph()
            sig_run = sig_p.add_run("Acceptance\nIN WITNESS WHEREOF, the parties have executed this Final Acceptance Section as of the date first written below.")
            sig_run.font.name = 'Arial'; sig_run.font.size = Pt(10)
            
            sig_table = doc.add_table(rows=5, cols=2)
            sig_table.autofit = False
            sig_table.columns[0].width = Inches(3.2)
            sig_table.columns[1].width = Inches(3.2)
            
            sig_table.cell(0,0).text = "Client: ___________________________"
            sig_table.cell(0,1).text = "Provider: My OpsBox"
            sig_table.cell(2,0).text = "X _________________________________"
            
            prov_sig_p = sig_table.cell(2,1).paragraphs[0]
            if os.path.exists(PATTI_SIG_PNG):
                prov_sig_p.add_run().add_picture(PATTI_SIG_PNG, width=Inches(1.8))
            else:
                prov_sig_p.add_run("X _________________________________")
            
            sig_table.cell(3,0).text = "Client Representative Name and Title"
            sig_table.cell(3,1).text = "Patti Zapparolli, Founder\nPROVIDER Representative Name and Title"
            sig_table.cell(4,0).text = "___________________________________\nDate Signed"
            sig_table.cell(4,1).text = f"{datetime.now().strftime('%B %d, %Y')}\nDate Signed"
            
            for row in sig_table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.name = 'Arial'
                            run.font.size = Pt(10)
            idx += 1
            continue

        if line.startswith("## "):
            clean_title = line.replace("## ", "").replace("**", "").strip()
            banner_table = doc.add_table(rows=1, cols=1)
            banner_table.autofit = False
            banner_table.columns[0].width = Inches(6.5)
            cell = banner_table.rows[0].cells[0]
            set_cell_background(cell, primary_hex)
            set_cell_margins(cell, top=140, bottom=140, left=120, right=120)

            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(2); p.paragraph_format.keep_with_next = True
            run = p.add_run(clean_title)
            run.font.name = 'Arial'; run.font.size = Pt(11.5); run.font.bold = True; run.font.color.rgb = RGBColor(255, 255, 255)

            spacer = doc.add_paragraph()
            spacer.paragraph_format.space_after = Pt(4)
            idx += 1
            continue

        if line.startswith("### "):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(4); p.paragraph_format.keep_with_next = True
            run = p.add_run(line.replace("### ", "").replace("**", "").strip())
            run.font.name = 'Arial'; run.font.size = Pt(11); run.font.bold = True; run.font.color.rgb = primary_color
            idx += 1
            continue

        if line.startswith("*") or line.startswith("-") or line.startswith("•"):
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(3)
            clean_line = line.lstrip("*•- ").strip()

            if " : " in clean_line or "**" in clean_line:
                clean_line = clean_line.replace("**", "")
                splitter = " : " if " : " in clean_line else ":"
                parts = clean_line.split(splitter, 1)

                r_bold = p.add_run(parts[0] + splitter)
                r_bold.font.name = 'Arial'; r_bold.font.size = Pt(10.5); r_bold.font.bold = True; r_bold.font.color.rgb = COLOR_TEXT_DARK
                if len(parts) > 1:
                    r_body = p.add_run(parts[1])
                    r_body.font.name = 'Arial'; r_body.font.size = Pt(10.5)
            else:
                run = p.add_run(clean_line)
                run.font.name = 'Arial'; run.font.size = Pt(10.5)
            idx += 1
            continue

        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(8); p.paragraph_format.line_spacing = 1.15
        clean_text = line.replace("**", "").strip()

        if ":" in clean_text and len(clean_text.split(":")[0]) < 30 and not clean_text.startswith("http"):
            parts = clean_text.split(":", 1)
            r_bold = p.add_run(parts[0] + ":")
            r_bold.font.name = 'Arial'; r_bold.font.size = Pt(10.5); r_bold.font.bold = True
            r_body = p.add_run(parts[1])
            r_body.font.name = 'Arial'; r_body.font.size = Pt(10.5)
        else:
            run = p.add_run(clean_text)
            run.font.name = 'Arial'; run.font.size = Pt(10.5)
        idx += 1

# --- DOCUMENT ENGINES ---
def audit_to_docx_buffer(text_content, meta_dict):
    """Dedicated DOCX engine for Growth Drag Audit matching the exact design spec."""
    text_content = clean_ai_markdown(text_content)
    doc = Document()
    
    client_name = meta_dict.get('client_business_name', 'Client')
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
        # --- HEADER WITH DYNAMIC PAGE NUMBERS ---
        header = section.header
        header_p = header.paragraphs[0]
        header_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        h_run = header_p.add_run("My OpsBox | Growth Drag Audit Analysis & Summary Report    ")
        h_run.font.name = 'Montserrat'
        h_run.font.size = Pt(9)
        h_run.font.color.rgb = RGBColor(160, 110, 210) 
        
        # Inject robust XML for dynamic Page Numbers
        pg_run = header_p.add_run()
        pg_run.font.name = 'Montserrat'
        pg_run.font.size = Pt(9)
        pg_run.font.color.rgb = RGBColor(160, 110, 210)
        
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = "PAGE"
        
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'separate')
        
        # Fallback text so the field isn't blank before Word updates it
        t_fallback = OxmlElement('w:t')
        t_fallback.text = "1"
        
        fldChar3 = OxmlElement('w:fldChar')
        fldChar3.set(qn('w:fldCharType'), 'end')
        
        pg_run._r.append(fldChar1)
        pg_run._r.append(instrText)
        pg_run._r.append(fldChar2)
        pg_run._r.append(t_fallback)
        pg_run._r.append(fldChar3)
        
        # --- CENTERED FOOTER ---
        footer = section.footer
        footer_p = footer.paragraphs[0]
        footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        f_run = footer_p.add_run(f"Prepared for {client_name} | Confidential Client Summary | Prepared by My OpsBox")
        f_run.font.name = 'Montserrat'
        f_run.font.size = Pt(9)
        f_run.font.color.rgb = RGBColor(160, 160, 160) # Light gray text from reference

    # --- CENTERED LOGO INJECTION ---
    if os.path.exists(MAIN_COLOR_LOGO_PNG):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(0)
        p_logo.paragraph_format.space_after = Pt(0)
        r_logo = p_logo.add_run()
        r_logo.add_picture(MAIN_COLOR_LOGO_PNG, width=Inches(3.5))
    
    # --- CENTERED TITLE BLOCK ---
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(2)
    p_title.paragraph_format.line_spacing = 1.15
    
    run_title1 = p_title.add_run("Growth Drag Audit\n")
    run_title1.font.name = 'Montserrat'
    run_title1.font.size = Pt(28)
    run_title1.font.color.rgb = COLOR_PURPLE_MAIN
    
    run_title2 = p_title.add_run("Analysis & Summary Report")
    run_title2.font.name = 'Montserrat'
    run_title2.font.size = Pt(14)
    run_title2.font.italic = True
    run_title2.font.color.rgb = COLOR_TEXT_DARK
    
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_meta.paragraph_format.space_after = Pt(12)
    p_meta.paragraph_format.line_spacing = 1.3
    
    r1 = p_meta.add_run(f"Prepared for {client_name}\n")
    r1.font.italic = True
    r1.font.name = 'Montserrat'
    r1.font.size = Pt(12)
    r1.font.color.rgb = COLOR_TEXT_DARK
    
    bp_date = meta_dict.get('blueprint_date_time', 'TBD') if meta_dict.get('blueprint_booked') == 'Yes' else 'TBD'
    r2 = p_meta.add_run(f"Client Contact: {meta_dict.get('client_contact_name_title', 'Client')} | Audit Date: {meta_dict.get('audit_date', 'TBD')} | Blueprint Review: {bp_date}\n")
    r2.font.name = 'Montserrat'
    r2.font.size = Pt(10)
    r2.font.color.rgb = COLOR_TEXT_DARK
    
    r3 = p_meta.add_run(f"Prepared by {meta_dict.get('prepared_by', 'Patti Zapparolli, Founder')} | My OpsBox")
    r3.font.name = 'Montserrat'
    r3.font.size = Pt(10)
    r3.font.color.rgb = COLOR_TEXT_DARK
    
    # --- PURPLE DIVIDER LINE ---
    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_div.paragraph_format.space_after = Pt(18)
    r_div = p_div.add_run("―" * 65)
    r_div.font.color.rgb = COLOR_PURPLE_MAIN
    r_div.font.bold = True

    # Parse remaining markdown body
    parse_markdown_to_docx(doc, text_content, primary_color=COLOR_PURPLE_MAIN, primary_hex=HEX_PURPLE_MAIN)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def text_to_docx_buffer(text_content, title_text):
    """Standard DOCX engine for Phase 1 Operational Assessments."""
    text_content = clean_ai_markdown(text_content)
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1); section.bottom_margin = Inches(1); section.left_margin = Inches(1); section.right_margin = Inches(1)

    header_table = doc.add_table(rows=1, cols=2)
    header_table.autofit = False
    header_table.rows[0].cells[0].width = Inches(3.8)
    header_table.rows[0].cells[1].width = Inches(2.7)

    left_cell = header_table.rows[0].cells[0]
    left_p = left_cell.paragraphs[0]
    if os.path.exists(MY_OPSBOX_LOGO_PNG):
        try: left_p.add_run().add_picture(MY_OPSBOX_LOGO_PNG, width=Inches(2.4))
        except: left_p.add_run("My OpsBox®").font.size = Pt(24)
    else:
        left_p.add_run("My OpsBox®").font.size = Pt(24)

    sub_p = left_cell.add_paragraph()
    sub_run = sub_p.add_run(title_text)
    sub_run.font.italic = True; sub_run.font.color.rgb = COLOR_SECONDARY

    right_cell = header_table.rows[0].cells[1]
    right_p = right_cell.paragraphs[0]
    right_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tcPr = right_cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="12" w:space="0" w:color="{HEX_PURPLE_MAIN}"/></w:tcBorders>')
    tcPr.append(borders)
    set_cell_margins(right_cell, top=60, bottom=60, left=180, right=60)

    for r_text, r_bold, r_color, r_size in [("Prepared by:\n", True, COLOR_SECONDARY, 9.5), ("Patti Zapparolli\n", True, COLOR_TEXT_DARK, 10), ("patti@myopsbox.com\n", False, COLOR_PURPLE_MAIN, 9.5), ("727-919-7323", False, COLOR_SECONDARY, 9.5)]:
        run = right_p.add_run(r_text); run.bold = r_bold; run.font.color.rgb = r_color; run.font.size = Pt(r_size)

    doc.add_paragraph().add_run("―" * 60).font.color.rgb = RGBColor(210, 214, 219)
    parse_markdown_to_docx(doc, text_content, primary_color=COLOR_PURPLE_MAIN, primary_hex=HEX_PURPLE_MAIN)
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def psa_to_docx_buffer(text_content, project_type_footer="PROJECT TYPE: Master Service Agreement | PSA#: TBD"):
    """Dedicated DOCX engine that exactly matches the My OpsBox PSA Template formatting."""
    text_content = clean_ai_markdown(text_content)
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1); section.bottom_margin = Inches(1); section.left_margin = Inches(1); section.right_margin = Inches(1)
        
        header = section.header
        header_p = header.paragraphs[0]
        if os.path.exists(MY_OPSBOX_LOGO_PNG):
            try: header_p.add_run().add_picture(MY_OPSBOX_LOGO_PNG, width=Inches(2.5))
            except: header_p.add_run("My OpsBox®")
        
        footer = section.footer
        footer_p = footer.paragraphs[0]
        footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        f_run = footer_p.add_run(project_type_footer)
        f_run.font.size = Pt(9); f_run.font.color.rgb = COLOR_SECONDARY

    parse_markdown_to_docx(doc, text_content, primary_color=COLOR_PURPLE_MAIN, primary_hex=HEX_PURPLE_MAIN)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def docx_to_pdf_buffer(docx_buffer):
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "temp.docx")
        pdf_path = os.path.join(tmpdir, "temp.pdf")
        with open(docx_path, "wb") as f: f.write(docx_buffer.getvalue())

        if platform.system() == "Windows":
            try:
                from docx2pdf import convert
                import pythoncom 
                pythoncom.CoInitialize()
                convert(docx_path, pdf_path)
                pythoncom.CoUninitialize()
            except Exception as win_err:
                st.error(f"Local Word Conversion Error: {win_err}")
                return None
        else:
            try:
                subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", tmpdir], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except FileNotFoundError:
                st.error("LibreOffice missing. Ensure packages.txt exists with 'libreoffice'.")
                return None

        with open(pdf_path, "rb") as f:
            return io.BytesIO(f.read())


# --- STREAMLIT UI SETUP ---
st.set_page_config(page_title="My OpsBox Pipeline Engine", page_icon=r"Favicon with Initials.png", layout="wide", initial_sidebar_state="collapsed")

logo_col1, logo_col2 = st.columns([1, 1])
with logo_col1:
    try: st.image(MY_OPSBOX_LOGO_URL, width=220)
    except: st.markdown("### **My OpsBox**")
with logo_col2:
    try:
        if os.path.exists(STEORA_LOGO_URL):
            with open(STEORA_LOGO_URL, "r", encoding="utf-8") as svg_file:
                st.markdown(f'<div style="display: flex; justify-content: flex-end;"><div style="width: 140px;">{svg_file.read()}</div></div>', unsafe_allow_html=True)
    except: pass

st.markdown("---")
st.title("🚀 My OpsBox Intelligent Pipeline Engine")
st.markdown("Sync Fathom meetings, run deep-dive strategic audits, and generate client-ready proposals instantly.")

st.sidebar.header("⚙️ Pipeline Status")
secret_gemini = st.secrets.get("GEMINI_API_KEY") or st.secrets.get("general", {}).get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")
secret_fathom = st.secrets.get("FATHOM_API_KEY") or st.secrets.get("general", {}).get("FATHOM_API_KEY") or os.getenv("FATHOM_API_KEY")

user_gemini = st.sidebar.text_input("Gemini API Key", type="password", placeholder="•••••••••••••••• (Autofetched)" if secret_gemini else "Enter Key")
user_fathom = st.sidebar.text_input("Fathom API Key", type="password", placeholder="•••••••••••••••• (Autofetched)" if secret_fathom else "Enter Key")

gemini_credential = user_gemini if user_gemini else secret_gemini
fathom_credential = user_fathom if user_fathom else secret_fathom
time_frame = st.sidebar.selectbox("Lookback Window", ["Today", "Past 7 Days", "Past 30 Days", "All Time"], index=3)

if not gemini_credential or not fathom_credential:
    st.info("💡 Please input your Fathom and Gemini API Keys in the sidebar.")
    st.stop()

headers = {"X-Api-Key": fathom_credential}
created_after_param = None
if time_frame == "Today": created_after_param = (datetime.utcnow() - timedelta(days=1)).strftime("%Y-%m-%dT%H:%M:%SZ")
elif time_frame == "Past 7 Days": created_after_param = (datetime.utcnow() - timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%SZ")
elif time_frame == "Past 30 Days": created_after_param = (datetime.utcnow() - timedelta(days=30)).strftime("%Y-%m-%dT%H:%M:%SZ")

if st.button("🔄 Sync Active Discovery Call Stream", type="primary"):
    with st.spinner("Connecting to Fathom Cloud Gateways..."):
        try:
            clean_endpoint = FATHOM_API_URL.split("://")[-1].replace(")", "").replace("]", "").strip()
            
            # 1. Drop the heavy transcript flag to prevent strict Fathom throttling
            # 2. Ask for max allowed limit per page (100)
            params = {
                "calendar_invitees_domains_type": "all",
                "limit": 100
            }
            if created_after_param: 
                params["created_after"] = created_after_param
            
            all_meetings = []
            cursor = None
            
            # 3. Standard Pagination Loop (Fetch up to 5 pages / 500 calls max to prevent long wait times)
            for _ in range(5):
                if cursor:
                    params["cursor"] = cursor
                    
                response = requests.get(f"https://{clean_endpoint}", headers=headers, params=params, timeout=30)
                
                if response.status_code == 200:
                    data = response.json()
                    all_meetings.extend(data.get("items", []))
                    
                    # Check if Fathom provided a cursor for the next page of results
                    cursor = data.get("next_cursor")
                    if not cursor:
                        break  # No more calls in the selected timeframe
                else:
                    if not all_meetings: # Only show error if the very first request failed
                        st.error(f"Fathom API Error {response.status_code}: {response.text}")
                    break
                    
            if all_meetings:
                st.session_state["meetings_list"] = all_meetings
                st.success(f"Synced {len(all_meetings)} calls from Fathom stream.")
                
        except Exception as e: 
            st.error(f"Pipeline failure: {e}")

if "meetings_list" in st.session_state and st.session_state["meetings_list"]:
    meetings = st.session_state["meetings_list"]
    options = [(f"{m.get('created_at', 'TBD')[:10]} — {m.get('title', 'Call')}", idx) for idx, m in enumerate(meetings)]

    # --- MULTI-SELECT UPGRADE ---
    selected_calls = st.multiselect("Select Call Stream(s) for Analysis (e.g. Discovery + Audit):", options=options, format_func=lambda x: x[0])

    raw_transcript_block = ""
    safe_name = "Analysis"
    if selected_calls:
        duration_minutes_total = 0
        safe_name = "".join([c if c.isalnum() else "_" for c in meetings[selected_calls[0][1]].get('title', 'Analysis')])
        
        with st.spinner("Downloading raw call data elements..."):
            for idx, call_tuple in enumerate(selected_calls):
                target = meetings[call_tuple[1]]
                rec_id = target.get("recording_id")
                call_title = target.get("title", "Unknown Call")
                call_date = target.get("created_at", "TBD")[:10]
                
                # --- 1. INITIAL DURATION CALCULATION ---
                raw_duration = target.get('duration') or target.get('duration_seconds') or target.get('length') or 0
                try:
                    duration_seconds = int(raw_duration)
                except (ValueError, TypeError):
                    duration_seconds = 0
                dur_mins = duration_seconds // 60

                # --- 2. FETCH TRANSCRIPT FIRST ---
                transcript_text = ""
                turns = []
                if rec_id:
                    clean_rec_path = f"api.fathom.ai/external/v1/recordings/{rec_id}/transcript".strip()
                    try:
                        t_resp = requests.get(f"https://{clean_rec_path}", headers=headers, timeout=30)
                        if t_resp.status_code == 200:
                            turns = t_resp.json().get("transcript", [])
                            transcript_text = "\n".join([f"[{t.get('timestamp', '00:00')}] {t.get('speaker', {}).get('display_name', 'Speaker')}: {t.get('text', '')}" for t in turns])
                        else:
                            st.warning(f"Failed to fetch transcript for {call_title}. Fathom returned code {t_resp.status_code}.")
                    except requests.exceptions.RequestException as net_err:
                        st.warning(f"Network drop while downloading {call_title}. Skipping transcript. Error: {net_err}")
                
                if not transcript_text:
                    raw_t = target.get("transcript")
                    if isinstance(raw_t, list):
                        turns = raw_t
                        transcript_text = "\n".join([f"[{t.get('timestamp', '00:00')}] {t.get('speaker', {}).get('display_name', 'Speaker')}: {t.get('text', '')}" for t in turns])
                    else:
                        transcript_text = str(raw_t or "No transcript payload present.")

                # --- 3. DURATION FALLBACK (Using the fetched turns) ---
                if dur_mins == 0:
                    if turns:
                        try:
                            last_ts = turns[-1].get("timestamp", "0:00").split(":")
                            if len(last_ts) == 2:
                                dur_mins = int(last_ts[0])
                            elif len(last_ts) == 3:
                                dur_mins = (int(last_ts[0]) * 60) + int(last_ts[1])
                        except Exception:
                            pass
                    if dur_mins == 0:
                        dur_mins = 15  # Default reasonable estimate if no metadata exists

                duration_minutes_total += dur_mins
                
                raw_transcript_block += f"\n\n--- [TRANSCRIPT {idx+1}: {call_title} ({call_date})] ---\n"
                raw_transcript_block += transcript_text

        col_m1, col_m2 = st.columns(2)
        col_m1.metric("Selected Client Calls", f"{len(selected_calls)} Meeting(s)")
        col_m2.metric("Total Duration Pool", f"{duration_minutes_total} minutes")

        st.markdown("---")
        
        # --- TABBED ARCHITECTURE ---
        tab1, tab2, tab3 = st.tabs(["📊 Phase 1: Growth Drag Audit Report", "🧠 Phase 2: Operational Assessment", "🤝 Phase 3: PSA Generator"])
        
        # ==========================================
        # TAB 1: GROWTH DRAG AUDIT REPORT AGENT
        # ==========================================
        with tab1:
            st.subheader("📊 Growth Drag Audit Summary Report Agent")
            st.markdown("Generate a diagnostic, client-facing summary report (No pricing or scope proposals).")
            
            with st.expander("📝 Required Intake Fields", expanded=True):
                col_i1, col_i2 = st.columns(2)
                with col_i1:
                    client_business_name = st.text_input("Client/Business Name")
                    prepared_for = st.text_input("Prepared For")
                    client_contact = st.text_input("Client Contact Name/Title")
                    industry = st.text_input("Industry")
                    audit_date = st.date_input("Audit Date")
                with col_i2:
                    attendees = st.text_input("Attendees")
                    pre_call_context = st.text_area("Pre-Call Context (Known goals, background)")
                    prepared_by = st.text_input("Prepared By", value="Patti Zapparolli")
                    blueprint_booked = st.radio("Blueprint Review Booked?", ["Yes", "No"], horizontal=True)
                    blueprint_date_time = st.text_input("Blueprint Date/Time (If Yes)")

            if st.button("🔥 Generate Growth Drag Audit Report", type="primary"):
                if not client_business_name or not prepared_for:
                    st.warning("Please fill out at least the Client Business Name and Prepared For fields.")
                else:
                    with st.spinner("Diagnosing operational drag and formatting client report..."):
                        try:
                            audit_prompt = f"""
You are an expert Fractional COO acting on behalf of My OpsBox. Generate a client-facing "Growth Drag Audit Summary Report" based strictly on the source hierarchy and rules below.

### INTAKE METADATA:
- Business Name: {client_business_name}
- Prepared For: {prepared_for}
- Contact Name/Title: {client_contact}
- Industry: {industry}
- Audit Date: {audit_date}
- Attendees: {attendees}
- Pre-Call Context: {pre_call_context}
- Prepared By: {prepared_by}
- Blueprint Booked: {blueprint_booked}
- Blueprint Date/Time: {blueprint_date_time}

### SOURCE TRANSCRIPT(S):
{raw_transcript_block}

### TRANSFORMATION RULES (CLIENT-LANGUAGE SANITIZATION):
Transform informal complaints into polished operational gaps. 
Example 1 -> Raw: "I'm the problem." | Polished: "Growth activity currently depends heavily on the owner's availability, comfort level, and follow-through."
Example 2 -> Raw: "I just don't want to do it." | Polished: "The task does not align with the owner's highest-value role, reducing execution consistency."

### MANDATORY REPORT STRUCTURE (LOCKED HEADINGS ONLY):
Do NOT output a title block or meta block. Start immediately with '## Executive Summary'.

## Executive Summary
[2 short paragraphs summarizing the primary source of operational drag, current situation, and strongest opportunity. Diagnostic only.]

## Current State Snapshot
| Category | Summary |
| --- | --- |
| Business Model | [1-3 sentences on core offer & client model] |
| Current Support | [1-3 sentences on current team/contractors/systems] |
| 12-Month Goal | [1-3 sentences on primary target] |
| Current Pressure | [1-3 sentences on seasonal or operational pressure] |
| Primary Operating Gap | [1-3 sentences on where bottleneck occurs] |

## Primary Growth Drag Findings
[Provide 3 to 5 findings using exact sub-structure below:]
### Drag Area 1: [Clear Finding Title]
**What showed up:** [Summary of challenge from transcript]
**What to do about it:** [High-level diagnostic direction. No packages/pricing.]

### Drag Area 2: [Clear Finding Title]
**What showed up:** [Summary of challenge from transcript]
**What to do about it:** [High-level diagnostic direction.]

## Where to Start
[Short introductory paragraph followed by 3-5 high-level strategic bullet points focusing on immediate relief and stabilization. NO task lists, NO 30-60-90 day plans.]

## Expected Business Outcomes
| Outcome | What It Means for {client_business_name} |
| --- | --- |
| [e.g., Lower Owner Dependency] | [Operational explanation] |
| [e.g., More Consistent Pipeline] | [Operational explanation] |

## Next Step: Traction Blueprint Review
{'The Traction Blueprint Review on ' + blueprint_date_time + ' will turn these findings into a practical recommended plan. The review will clarify the highest-impact starting priorities, the support structure, how progress should be tracked, and what path will create the best traction. ' + client_business_name + ' is already moving in the right direction by identifying the drag before adding more activity. The next step is to align process, ownership, and measurement so growth can become more consistent, visible, and manageable.' if blueprint_booked == 'Yes' else 'The recommended next step is to schedule the Traction Blueprint Review. This review will turn the audit findings into a practical recommended plan, including the highest-impact starting priorities, the right-fit support structure, how progress should be tracked, and where the strongest operational value is likely to come from. The Growth Drag Audit identified where the business is getting stuck. The Blueprint Review is where those findings become a clear path forward.'}

### CRITICAL RULES:
1. Target length: 2 to 4 pages.
2. NEVER use ASCII art boxes or lines (e.g., +---+ or ┌───┐). Use standard Markdown tables for Section 2 and 5.
3. STRICT EXCLUSIONS: Do NOT include pricing, packages, service tiers, staffing lists, legal terms, or guarantees.
"""
                            sys_inst = "You are a world-class Fractional COO diagnostician. Output strict Markdown."
                            audit_resp = call_gemini_api(audit_prompt, gemini_credential, system_instruction=sys_inst)
                            
                            st.session_state["compiled_audit"] = clean_ai_markdown(audit_resp)
                            st.session_state["audit_meta"] = {
                                "client_business_name": client_business_name, "prepared_for": prepared_for,
                                "client_contact_name_title": client_contact, "audit_date": str(audit_date),
                                "blueprint_booked": blueprint_booked, "blueprint_date_time": blueprint_date_time,
                                "prepared_by": prepared_by
                            }
                            st.success("Growth Drag Audit Report Generated!")
                        except Exception as e:
                            st.error(f"Audit Generation Error: {e}")

            if "compiled_audit" in st.session_state:
                st.markdown("### 📋 Generated Growth Drag Audit")
                st.markdown(st.session_state["compiled_audit"])
                st.markdown("### 📥 Audit Export Options")
                
                # Fetch business name and scrub it for safe file saving
                raw_b_name = st.session_state["audit_meta"].get("client_business_name", "Client")
                safe_b_name = "".join([c if c.isalnum() or c in " _-" else "_" for c in raw_b_name]).strip()
                audit_file_name = f"{safe_b_name} Growth Drag Audit Analysis"
                
                if "export_audit_docx" not in st.session_state or st.session_state.get("last_audit_title") != safe_name:
                    st.session_state["export_audit_docx"] = audit_to_docx_buffer(st.session_state["compiled_audit"], st.session_state["audit_meta"])
                    with st.spinner("Rendering exact-match PDF..."):
                        st.session_state["export_audit_pdf"] = docx_to_pdf_buffer(st.session_state["export_audit_docx"])
                    st.session_state["last_audit_title"] = safe_name

                col_a1, col_a2 = st.columns(2)
                with col_a1: 
                    st.download_button("💾 Save Audit (Word)", data=st.session_state["export_audit_docx"], file_name=f"{audit_file_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                with col_a2: 
                    if st.session_state["export_audit_pdf"]:
                        st.download_button("💾 Save Audit (PDF)", data=st.session_state["export_audit_pdf"], file_name=f"{audit_file_name}.pdf", mime="application/pdf")

        # ==========================================
        # TAB 2: OPERATIONAL ASSESSMENT
        # ==========================================
        with tab2:
            st.subheader("🧠 Operational Assessment Matrix")
            if st.button("🔥 Run Operational Assessment Analysis", type="secondary"):
                with st.spinner("Processing deep analysis matrix..."):
                    try:
                        master_prompt = f"""
You are an expert Fractional Chief Operating Officer (Fractional COO) acting on behalf of My OpsBox. 
Your objective is to thoroughly analyze the attached raw Fathom call transcripts and generate a comprehensive Operational Assessment document.

### CRITICAL OUTPUT FORMATTING INSTRUCTIONS FOR THE DOCX PROCESSING ENGINE:
1. CLIENT CARD META BLOCK:
| Business Name | [Insert] | Assessment Date | {datetime.now().strftime('%m/%d/%y')} |
| Website | [Insert] | Email | [Insert] |
| Business Type | [Insert] | Business Industry | [Insert] |
| Number of full-time employees | [ headcount ] | Annual revenue range | [Estimated] |

2. PRIMARY HEADINGS:
## 1. Executive Summary & Business Context
## 2. Key Operational Challenges & Gaps
## 3. Strategic 12-Month Objectives
## 4. Proposed Service Package Strategy
## 5. Structured 30-60-90 Day Transformation Timeline Plan
## 6. Scorecard Matrix Projections

3. STRICT NO ASCII ART OR FLOWCHARTS:
Never use ASCII block shapes or lines (e.g., ┌───┐). Represent all objectives strictly as standard bulleted lists.

### LIVE CALL TRANSCRIPT TEXT TO PROCESS AND ANALYZE:
{raw_transcript_block}
"""
                        sys_inst = "You are a world-class Fractional COO. Output strict Markdown without ASCII art."
                        ai_resp = call_gemini_api(master_prompt, gemini_credential, system_instruction=sys_inst)
                        st.session_state["compiled_analysis"] = clean_ai_markdown(ai_resp)
                        st.success("AI Strategic Assessment Generated Flawlessly!")
                    except Exception as ai_err: st.error(f"Exception: {ai_err}")

            if "compiled_analysis" in st.session_state:
                st.markdown("### 📋 Generated Strategic Analysis")
                st.markdown(st.session_state["compiled_analysis"])
                if "export_docx_buf" not in st.session_state or st.session_state.get("last_call_title") != safe_name:
                    st.session_state["export_docx_buf"] = text_to_docx_buffer(st.session_state["compiled_analysis"], "Operational Assessment & Strategic Analysis")
                    with st.spinner("Rendering PDF..."):
                        st.session_state["export_pdf_buf"] = docx_to_pdf_buffer(st.session_state["export_docx_buf"])
                    st.session_state["last_call_title"] = safe_name

                col_e1, col_e2 = st.columns(2)
                with col_e1: st.download_button("💾 Save Assessment (Word)", data=st.session_state["export_docx_buf"], file_name=f"Assessment_{safe_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                with col_e2: 
                    if st.session_state["export_pdf_buf"]:
                        st.download_button("💾 Save Assessment (PDF)", data=st.session_state["export_pdf_buf"], file_name=f"Assessment_{safe_name}.pdf", mime="application/pdf")

        # ==========================================
        # TAB 3: PSA GENERATOR
        # ==========================================
        with tab3:
            st.subheader("🤝 Proposal & Service Agreement (PSA) Generator")
            col_r1, col_r2 = st.columns(2)
            with col_r1: managed_services_file = st.file_uploader("Upload Managed Services Guide", type=["pdf"], key="msg_file")
            with col_r2: psa_template_file = st.file_uploader("Upload Sample PSA Template", type=["pdf"], key="psa_file")

            col_p1, col_p2 = st.columns(2)
            with col_p1: contract_length = st.text_input("Minimum Contract Length", value="12 months")
            with col_p2: pricing_structure = st.text_area("Pricing / Rate Structure Per Month")

            if st.button("📑 Generate Final PSA Document", type="primary"):
                if not managed_services_file or not psa_template_file or not pricing_structure.strip():
                    st.warning("Please upload both PDFs and provide the pricing structure.")
                elif "compiled_analysis" not in st.session_state:
                    st.warning("Please run the Operational Assessment (Phase 2 Tab) first to gather context.")
                else:
                    with st.spinner("Extracting template rules and generating highly customized PSA..."):
                        try:
                            msg_text = extract_pdf_text(managed_services_file)
                            psa_text = extract_pdf_text(psa_template_file)
                            psa_prompt = f"""
You are an expert Fractional Chief Operating Officer (Fractional COO) and elite sales closer.
### SOURCE MATERIALS:
1. **Assessment Notes:** {st.session_state["compiled_analysis"]}
2. **Managed Services Guide:** {msg_text}
3. **Sample PSA Context:** {psa_text}
4. **Pricing:** {pricing_structure} | **Length:** {contract_length}

### INSTRUCTIONS FOR DOCUMENT GENERATION (STRICT MARKDOWN):
**Prepared For:** [Insert Client Name, Company]
**My PreOpsBox:** [Insert Proposed Role/Package]
**Date Prepared:** {datetime.now().strftime('%B %d, %Y')}

This Proposal and Service Agreement for the [Insert Service] (the "Agreement") will be effective from [Insert Start Date] ("Effective Date") and continue until [Insert End Date].

**BETWEEN:**
[Insert Company Name] (The "CLIENT")...
**AND:**
My OpsBox (The "PROVIDER")...

**Service Length:** {contract_length}
**Monthly Rate:** {pricing_structure}

**Scope Transparency:**
- Rate covers only items listed in Scope of Service.
- Rate includes $50 monthly phone fee.

## Scope of Work
[Generate custom scope mapped from Guide]

## Value
[Generate custom ROI]

## Time and Payment
[Include standard payment terms]

## Standard Operating Procedures
[Include standard SOPs from template]

## Legal Notices
[Include standard legal clauses]

## Entire Agreement
[Include final clauses]

[SIGNATURE_BLOCK]
"""
                            sys_inst = "You are a master fractional executive and sales closer. Output pure Markdown documents only."
                            psa_resp = call_gemini_api(psa_prompt, gemini_credential, system_instruction=sys_inst)
                            st.session_state["compiled_psa"] = clean_ai_markdown(psa_resp)
                            st.success("PSA Generated Successfully!")
                        except Exception as e: st.error(f"PSA Generation Error: {e}")

            if "compiled_psa" in st.session_state:
                st.markdown("### 📋 Generated Proposal & Service Agreement")
                st.markdown(st.session_state["compiled_psa"])
                project_footer = f"PROJECT TYPE: Proposal & Service Agreement | PSA#: {safe_name.upper()}"
                
                if "export_psa_docx" not in st.session_state or st.session_state.get("last_psa_title") != safe_name:
                    st.session_state["export_psa_docx"] = psa_to_docx_buffer(st.session_state["compiled_psa"], project_footer)
                    with st.spinner("Rendering exact-match PDF..."):
                        st.session_state["export_psa_pdf"] = docx_to_pdf_buffer(st.session_state["export_psa_docx"])
                    st.session_state["last_psa_title"] = safe_name

                col_psa1, col_psa2 = st.columns(2)
                with col_psa1: st.download_button("💾 Save PSA (Word)", data=st.session_state["export_psa_docx"], file_name=f"PSA_{safe_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                with col_psa2: 
                    if st.session_state["export_psa_pdf"]:
                        st.download_button("💾 Save PSA (PDF)", data=st.session_state["export_psa_pdf"], file_name=f"PSA_{safe_name}.pdf", mime="application/pdf")