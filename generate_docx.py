import os
import re
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

md_path = r"g:\My Drive\Dev\Einstein\Le_Pacte_Nice_IA.md"
docx_path = r"g:\My Drive\Dev\Einstein\Le_Pacte_Nice_IA.docx"

with open(md_path, 'r', encoding='utf-8') as f:
    md_text = f.read()

doc = Document()

# Set page margins (A4)
sections = doc.sections
for section in sections:
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

# Color Palette
CHAMPAGNE = RGBColor(184, 151, 98)
NAVY_BLUE = RGBColor(15, 23, 42)
SKY_BLUE = RGBColor(2, 132, 199)
DARK_TEXT = RGBColor(30, 41, 59)

def set_cell_background(cell, fill_color):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_color}"/>')
    tcPr.append(shd)

def add_styled_heading(doc, text, level):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    
    run = p.add_run(text)
    if level == 1:
        run.font.name = 'Georgia'
        run.font.size = Pt(20)
        run.font.bold = True
        run.font.color.rgb = NAVY_BLUE
        p.paragraph_format.space_before = Pt(20)
    elif level == 2:
        run.font.name = 'Georgia'
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = SKY_BLUE
    elif level == 3:
        run.font.name = 'Georgia'
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = CHAMPAGNE
    elif level == 4:
        run.font.name = 'Arial'
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = NAVY_BLUE
    return p

def add_formatted_text(paragraph, text):
    # Process bold **text** and italic *text*
    parts = re.split(r'(\*\*.*?\*\*|\*.*?\*|\[.*?\]\(.*?\))', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            r = paragraph.add_run(part[2:-2])
            r.font.bold = True
            r.font.color.rgb = NAVY_BLUE
        elif part.startswith('*') and part.endswith('*'):
            r = paragraph.add_run(part[1:-1])
            r.font.italic = True
        elif part.startswith('[') and ']' in part and '(' in part and part.endswith(')'):
            m = re.match(r'\[(.*?)\]\((.*?)\)', part)
            if m:
                r = paragraph.add_run(m.group(1))
                r.font.color.rgb = SKY_BLUE
                r.font.underline = True
        else:
            if part:
                r = paragraph.add_run(part)
                r.font.color.rgb = DARK_TEXT

lines = md_text.split('\n')
i = 0
in_table = False
table_rows = []

while i < len(lines):
    line = lines[i].strip()
    
    if not line:
        i += 1
        continue
        
    # Check Header
    if line.startswith('#'):
        level = len(line.split()[0])
        text = line.lstrip('#').strip()
        add_styled_heading(doc, text, level)
        i += 1
        continue
        
    # Check Table
    if line.startswith('|') and line.endswith('|'):
        table_rows = []
        while i < len(lines) and lines[i].strip().startswith('|') and lines[i].strip().endswith('|'):
            row_line = lines[i].strip()
            if '---' not in row_line:
                cells = [c.strip() for c in row_line.split('|')[1:-1]]
                table_rows.append(cells)
            i += 1
            
        if table_rows:
            num_cols = len(table_rows[0])
            table = doc.add_table(rows=len(table_rows), cols=num_cols)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            for row_idx, row_data in enumerate(table_rows):
                for col_idx, cell_value in enumerate(row_data):
                    cell = table.cell(row_idx, col_idx)
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_before = Pt(4)
                    p.paragraph_format.space_after = Pt(4)
                    
                    if row_idx == 0:
                        set_cell_background(cell, "0F172A")
                        r = p.add_run(cell_value)
                        r.font.bold = True
                        r.font.color.rgb = RGBColor(216, 185, 138) # Champagne
                        r.font.size = Pt(9.5)
                    else:
                        if row_idx % 2 == 1:
                            set_cell_background(cell, "F8FAFC")
                        else:
                            set_cell_background(cell, "FFFFFF")
                        add_formatted_text(p, cell_value)
                        for run in p.runs:
                            run.font.size = Pt(9)
            doc.add_paragraph().paragraph_format.space_after = Pt(6)
        continue
        
    # Check Bullet Lists
    if line.startswith('* ') or line.startswith('- '):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        add_formatted_text(p, line[2:])
        i += 1
        continue
        
    # Check Numbered Lists
    if re.match(r'^\d+\.\s', line):
        text = re.sub(r'^\d+\.\s', '', line)
        p = doc.add_paragraph(style='List Number')
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        add_formatted_text(p, text)
        i += 1
        continue
        
    # Check Blockquotes
    if line.startswith('> '):
        quote_text = line.lstrip('> ').replace('[!IMPORTANT]', '').strip()
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.4)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        add_formatted_text(p, quote_text)
        i += 1
        continue
        
    # Normal Paragraph
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    add_formatted_text(p, line)
    i += 1

doc.save(docx_path)
print(f"Document Word (compatible Google Docs) généré avec succès : {docx_path}")
