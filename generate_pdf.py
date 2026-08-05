import os
import re
import subprocess

md_path = r"g:\My Drive\Dev\Einstein\Le_Pacte_Nice_IA.md"
html_path = r"g:\My Drive\Dev\Einstein\temp_pdf_template.html"
pdf_path = r"g:\My Drive\Dev\Einstein\Le_Pacte_Nice_IA.pdf"

with open(md_path, 'r', encoding='utf-8') as f:
    md_content = f.read()

def md_to_html(md):
    html = md
    
    # Headers
    html = re.sub(r'^##### (.*$)', r'<h5>\1</h5>', html, flags=re.M)
    html = re.sub(r'^#### (.*$)', r'<h4>\1</h4>', html, flags=re.M)
    html = re.sub(r'^### (.*$)', r'<h3>\1</h3>', html, flags=re.M)
    html = re.sub(r'^## (.*$)', r'<h2>\1</h2>', html, flags=re.M)
    html = re.sub(r'^# (.*$)', r'<h1>\1</h1>', html, flags=re.M)
    
    # Text styles
    html = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', html)
    html = re.sub(r'\*(.*?)\*', r'<em>\1</em>', html)
    html = re.sub(r'\[(.*?)\]\((.*?)\)', r'<a href="\2">\1</a>', html)
    html = re.sub(r'^---$', r'<hr>', html, flags=re.M)
    
    # Alerts / Quotes
    html = re.sub(r'^> \[\!IMPORTANT\]\n> (.*$)', r'<div class="alert alert-important"><strong>⚡ Important:</strong> \1</div>', html, flags=re.M)
    html = re.sub(r'^> (.*$)', r'<blockquote>\1</blockquote>', html, flags=re.M)
    
    # Lines & Tables
    lines = html.split('\n')
    in_table = False
    table_html = ''
    new_lines = []
    
    for line in lines:
        l = line.strip()
        if l.startswith('|') and l.endswith('|'):
            if '---' in l:
                continue
            cells = [c.strip() for c in l.split('|')[1:-1]]
            if not in_table:
                in_table = True
                table_html = '<table><thead><tr>' + ''.join(f'<th>{c}</th>' for c in cells) + '</tr></thead><tbody>'
            else:
                table_html += '<tr>' + ''.join(f'<td>{c}</td>' for c in cells) + '</tr>'
        else:
            if in_table:
                in_table = False
                table_html += '</tbody></table>'
                new_lines.append(table_html)
                table_html = ''
            if l.startswith('* ') or l.startswith('- '):
                l = f'<li>{l[2:]}</li>'
            new_lines.append(l)
            
    if in_table:
        table_html += '</tbody></table>'
        new_lines.append(table_html)
        
    html = '\n'.join(new_lines)
    html = re.sub(r'\n\n', r'</p><p>', html)
    return f'<p>{html}</p>'

parsed_body = md_to_html(md_content)

full_html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Le Pacte Nice IA — Document Stratégique</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Playfair+Display:ital,wght@0,600;0,700;1,400&display=swap');
        
        @page {{
            size: A4;
            margin: 20mm 15mm 20mm 15mm;
        }}
        
        body {{
            font-family: 'Inter', sans-serif;
            color: #1E293B;
            line-height: 1.6;
            font-size: 10.5pt;
            background: #FFFFFF;
        }}
        
        .header-banner {{
            border-bottom: 3px solid #D8B98A;
            padding-bottom: 12px;
            margin-bottom: 24px;
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
        }}
        
        .doc-title {{
            font-family: 'Playfair Display', serif;
            font-size: 22pt;
            font-weight: 700;
            color: #0F172A;
            margin: 0 0 6px 0;
            letter-spacing: -0.5px;
        }}
        
        .doc-subtitle {{
            font-size: 11pt;
            color: #475569;
            margin: 0;
            font-weight: 500;
        }}
        
        .meta-tag {{
            font-size: 9pt;
            color: #94A3B8;
            text-transform: uppercase;
            letter-spacing: 1px;
            font-weight: 600;
        }}
        
        h1 {{
            font-family: 'Playfair Display', serif;
            font-size: 18pt;
            color: #0F172A;
            border-bottom: 2px solid #E2E8F0;
            padding-bottom: 6px;
            margin-top: 24px;
            margin-bottom: 14px;
            page-break-after: avoid;
        }}
        
        h2 {{
            font-family: 'Playfair Display', serif;
            font-size: 14pt;
            color: #0284C7;
            margin-top: 20px;
            margin-bottom: 10px;
            page-break-after: avoid;
        }}
        
        h3 {{
            font-size: 12pt;
            color: #D8B98A;
            margin-top: 16px;
            margin-bottom: 8px;
            page-break-after: avoid;
        }}
        
        p {{
            margin-bottom: 10px;
            text-align: justify;
        }}
        
        strong {{
            color: #0F172A;
        }}
        
        hr {{
            border: none;
            border-top: 1px solid #E2E8F0;
            margin: 20px 0;
        }}
        
        blockquote {{
            background: #F8FAFC;
            border-left: 4px solid #D8B98A;
            margin: 14px 0;
            padding: 10px 16px;
            font-style: italic;
            color: #334155;
        }}
        
        .alert-important {{
            background: #F0F9FF;
            border-left: 4px solid #0284C7;
            padding: 12px 16px;
            margin: 14px 0;
            border-radius: 0 6px 6px 0;
            font-style: normal;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 16px 0;
            font-size: 9.5pt;
            page-break-inside: avoid;
        }}
        
        th {{
            background: #0F172A;
            color: #F8FAFC;
            padding: 8px 10px;
            text-align: left;
            font-weight: 600;
        }}
        
        td {{
            padding: 8px 10px;
            border-bottom: 1px solid #E2E8F0;
        }}
        
        tr:nth-child(even) td {{
            background: #F8FAFC;
        }}
        
        ul {{
            margin: 8px 0 14px 20px;
            padding: 0;
        }}
        
        li {{
            margin-bottom: 4px;
        }}
        
        .footer-note {{
            margin-top: 40px;
            border-top: 1px solid #E2E8F0;
            padding-top: 12px;
            font-size: 8pt;
            color: #94A3B8;
            text-align: center;
        }}
    </style>
</head>
<body>
    <div class="header-banner">
        <div>
            <div class="meta-tag">Arx Consulting — Document de Position Officiel</div>
            <div class="doc-title">🤝 LE PACTE NICE IA</div>
            <div class="doc-subtitle">Doctrine Stratégique & Justifications Chiffrées (2026-2029)</div>
        </div>
    </div>
    
    {parsed_body}
    
    <div class="footer-note">
        Le Pacte Nice IA — Document confidentiel métropolitain rédigé par Benoît Sigwald — Août 2026 — arx-consulting.com/nice
    </div>
</body>
</html>
"""

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(full_html)

chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
cmd = [
    chrome_path,
    "--headless",
    "--disable-gpu",
    f"--print-to-pdf={pdf_path}",
    "--no-pdf-header-footer",
    html_path
]

print("Generation du PDF via Chrome headless...")
subprocess.run(cmd, check=True)
print(f"PDF généré avec succès à l'emplacement : {pdf_path}")
