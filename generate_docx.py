import os
import re
import subprocess
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

docx_path = r"g:\My Drive\Dev\Einstein\Le_Pacte_Nice_IA.docx"
chart1_png = r"g:\My Drive\Dev\Einstein\chart1_poids_national.png"
chart2_png = r"g:\My Drive\Dev\Einstein\chart2_roi_gains.png"

# Ensure sober charts are generated
subprocess.run(["python", r"g:\My Drive\Dev\Einstein\generate_charts.py"], check=True)

doc = Document()

# Page Margins (A4)
for section in doc.sections:
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

# Sober Color Palette (Classic Executive / Institutional)
NAVY_PRIMARY = RGBColor(15, 23, 42)    # #0F172A
NAVY_SECONDARY = RGBColor(30, 58, 138)  # #1E3A8A
SLATE_DARK = RGBColor(51, 65, 85)      # #334155
BODY_BLACK = RGBColor(30, 41, 59)      # #1E293B
MUTED_GREY = RGBColor(100, 116, 139)   # #64748B

def set_cell_background(cell, fill_color):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_color}"/>')
    tcPr.append(shd)

def add_header_banner():
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r1 = p.add_run("ARX CONSULTING — DOCUMENT STRATÉGIQUE MÉTROPOLITAIN\n")
    r1.font.name = 'Arial'
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = SLATE_DARK
    
    r2 = p.add_run("LE PACTE NICE IA\n")
    r2.font.name = 'Georgia'
    r2.font.size = Pt(22)
    r2.font.bold = True
    r2.font.color.rgb = NAVY_PRIMARY
    
    r3 = p.add_run("Doctrine Stratégique, Rigueur Budgétaire & Alliance Transfrontalière (2026-2029)")
    r3.font.name = 'Arial'
    r3.font.size = Pt(10.5)
    r3.font.color.rgb = NAVY_SECONDARY
    
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(16)
    r_meta = p_meta.add_run("Rédigé pour M. Éric Ciotti par Benoît Sigwald — Senior AI Architect & AMO IA Métropolitain — Août 2026")
    r_meta.font.size = Pt(8.5)
    r_meta.font.italic = True
    r_meta.font.color.rgb = MUTED_GREY

def add_styled_heading(text, level):
    p = doc.add_paragraph()
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    
    if level == 1:
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        r.font.name = 'Georgia'
        r.font.size = Pt(15)
        r.font.bold = True
        r.font.color.rgb = NAVY_PRIMARY
    elif level == 2:
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        r.font.name = 'Georgia'
        r.font.size = Pt(12.5)
        r.font.bold = True
        r.font.color.rgb = NAVY_SECONDARY
    elif level == 3:
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        r.font.name = 'Arial'
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = SLATE_DARK

def add_p(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.15
    
    parts = re.split(r'(\*\*.*?\*\*|\*.*?\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            r = p.add_run(part[2:-2])
            r.font.bold = True
            r.font.color.rgb = NAVY_PRIMARY
        elif part.startswith('*') and part.endswith('*'):
            r = p.add_run(part[1:-1])
            r.font.italic = True
        else:
            if part:
                r = p.add_run(part)
                r.font.color.rgb = BODY_BLACK

def add_bullet(text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    parts = re.split(r'(\*\*.*?\*\*|\*.*?\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            r = p.add_run(part[2:-2])
            r.font.bold = True
            r.font.color.rgb = NAVY_PRIMARY
        elif part.startswith('*') and part.endswith('*'):
            r = p.add_run(part[1:-1])
            r.font.italic = True
        else:
            if part:
                r = p.add_run(part)
                r.font.color.rgb = BODY_BLACK

# Build Document
add_header_banner()

# Executive Summary Box
add_styled_heading("Résumé Exécutif & Chiffrage Consolidé pour M. le Maire", 1)
add_p("**La thèse.** La course aux modèles de frontière se joue à l’échelle des superpuissances. Pour Nice, la vraie bataille stratégique réside dans **l'usage concret, la sécurité publique, la souveraineté et la rigueur budgétaire**. Aucune collectivité n'a encore préempté la position de **« Capitale de l'IA de sécurité et d'efficience publique »**. Nice doit être la première sous la conduite de M. Éric Ciotti.")

# Sober Summary Table
table_data = [
    ["Pilier Stratégique", "Levier IA Appliqué", "Chiffre Clé & Impact", "Justification / Source"],
    ["1. CSU Augmenté & Sécurité", "VSA 4 300+ caméras & alertes", "Incivilités -65 %", "Intervention PM < 6 min."],
    ["2. Audit Commande Publique", "Ingestion sémantique factures (300 M€)", "+2,50 M€ / an NET", "Erreurs & doublons filtrés."],
    ["3. Bouclier Cyber-IA (NIS 2)", "SOC IA 24/7 souverain & Air-Gap", "1,8 à 3 M€ / an", "Crises ransomware évitées."],
    ["4. Monaco Cloud & Zone Franche", "Data Center Souverain (AMSN) & Zone Franche", "1,5 à 2,8 M€ / an", "Économies IT + 8-12M€ CAPEX évité."],
    ["5. Financements Europe & Gigafactory", "Subventions UE (EuroHPC, DIGITAL)", "500 M€ visés", "Candidature binationale Nice-Monaco."]
]

table = doc.add_table(rows=len(table_data), cols=4)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
for r_idx, row in enumerate(table_data):
    for c_idx, val in enumerate(row):
        cell = table.cell(r_idx, c_idx)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        if r_idx == 0:
            set_cell_background(cell, "1E293B")
            r = p.add_run(val)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(9)
        else:
            set_cell_background(cell, "F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            r.font.color.rgb = BODY_BLACK

doc.add_paragraph().paragraph_format.space_after = Pt(4)

# Insert Sober Graph 2 (ROI Financials)
if os.path.exists(chart2_png):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(6)
    p_img.paragraph_format.space_after = Pt(10)
    run_img = p_img.add_run()
    run_img.add_picture(chart2_png, width=Inches(6.0))

# Section 1
add_styled_heading("1. Où en sommes-nous ? État des lieux & Diagnostic Territorial", 1)

add_styled_heading("1.1 Contexte national & régional : Blocages politiques & Urgence d'un Leadership Unifié", 2)
add_bullet("**Échelle nationale** : La note *« Opération Prométhée »* (juillet 2026, Le Grand Continent) fixe le cadre : un plan de **700 Md$ sur 3 ans** (12 GW, 1 700 chercheurs). L'IA est devenue une ressource stratégique souveraine identique à l'énergie.")
add_bullet("**Échelle régionale** : La Région Sud affiche un *Plan SUD IA* (70 M€ sur 5 ans), la Métropole lance des appels à projets isolés et le Département anime la Maison de l'IA (MIA) à Sophia Antipolis.")
add_bullet("**Le blocage politique (Constat lucide)** : **Les tensions et rivalités politiques au sein de la Région PACA interfèrent directement dans l'obtention des subventions publiques et régionales**. Les projets azuréens sont aujourd'hui pénalisés par ces disputes d'appareil.")
add_bullet("**La solution** : Seul un **leadership métropolitain incontestable porté au plus haut niveau par M. Éric Ciotti** permettra d'outrepasser ces blocages politiques régionaux et d'aller capturer directement les subventions auprès de l'État et des guichets européens.")

add_styled_heading("1.2 Le terreau niçois : Un actif exceptionnel en France", 2)
add_bullet("**Sophia Antipolis (1ère technopole d’Europe)** : ~2 700 entreprises, ~46 000 emplois, ~5 500 chercheurs.")
add_bullet("**Institut 3IA Côte d’Azur** : L'un des 4 instituts nationaux d'IA (spécialisé en santé numérique).")
add_bullet("**Pipeline académique complet** : Université Côte d'Azur (UCA), Inria, Eurecom, CNRS (du Master au Doctorat).")
add_bullet("**Monaco Cloud & Data Center Souverain d'État** : Premier Cloud d'État souverain d'Europe (certifié AMSN), garantissant une étanchéité totale contre le Cloud Act américain, connecté en fibre noire dédiée à Nice.")
add_bullet("**Le levier Zone Franche Numérique Nice-Monaco** : Création d'une zone d'expérimentation fiscale et réglementaire transfrontalière privilégiée pour attirer et retenir les PME, startups et licornes de l'IA.")
add_bullet("**Grands industriels ancres riches en données** : **Amadeus** (1er centre de R&D privé de transport en Europe), **Thales Alenia Space** Cannes (spatial & observation satellite), santé, arômes-parfums (Grasse), maritime.")
add_bullet("**Attractivité & 2e aéroport de France** : La « carte Riviera » permet de capter et retenir les chercheurs d'élite que Paris peine à conserver.")

# Insert Sober Graph 1 (National Weight Comparison)
if os.path.exists(chart1_png):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(10)
    run_img = p_img.add_run()
    run_img.add_picture(chart1_png, width=Inches(6.0))

add_styled_heading("1.3 Évaluation du Pôle Nice - Sophia Antipolis - Monaco vs Total National Français", 2)
add_p("Afin d'emporter l'adhésion de l'État et de la Commission Européenne, le territoire fait valoir son **poids relatif massif à l'échelle de la France** :")
add_bullet("**Chercheurs R&D Numérique** : ~5 500 chercheurs (public+privé), soit **~12 % du total hors Île-de-France** (2e bassin national après Paris).")
add_bullet("**Recherche Académique 3IA** : > 100 chaires mondiales, soit **~25 % du réseau national des 4 Instituts 3IA** (Paris, Grenoble, Toulouse, Nice-Sophia).")
add_bullet("**R&D Privée** : Amadeus & Thales Alenia Space (>6 500 ingénieurs), soit **~6 % de la dépense privée logicielle/spatiale française**.")
add_bullet("**Infrastructures Cloud Souveraines** : Monaco Cloud (1er Cloud d'État UE), seul hub souverain binational conforme AMSN & NIS 2.")

add_styled_heading("1.4 Les faiblesses et contraintes lucidement posées", 2)
add_bullet("**Filière IA pure réduite** : ~86 établissements et ~800 emplois directs IA dans le 06 (étude CCI).")
add_bullet("**Le verrou électrique** : Extrémité d'une « presqu'île électrique » vulnérable (coupure majeure en 2009 ; capacité RTE/CNDP 2025 saturée).")
add_bullet("**Foncier contraint** : Mer et montagne rendent impossible l'accueil de campus de calcul hyperscale.")
add_bullet("**Le cimetière des POC** : Expérimentations locales bloquées au stade de prototype faute d'expertise en industrialisation (LLMOps, sécurité, AI Act).")

add_styled_heading("1.5 L'Opportunité Historique : Les 30 Md€ de l'UE & Les 7 AI Gigafactories", 2)
add_p("Le **30 juillet 2026**, la Commission Européenne a officiellement lancé un appel d'offres historique de **30 milliards d'euros** pour bâtir **7 AI Gigafactories** en Europe :")
add_bullet("**Lot 1 — Subvention directe UE de 500 M€** par site pour co-financer un supercalculateur d'IA souverain.")
add_bullet("**Le levier Nice-Monaco** : Candidature binationale transfrontalière unique associant Nice (Plaine du Var / 3IA), Monaco (fonds souverains / Monaco Cloud) et Sophia Antipolis.")
add_bullet("**Montage mixte à 1,5 Md€** : 500 M€ subvention UE + 300 M€ fonds publics + 700 M€ investisseurs privés.")
add_bullet("**Calendrier couperet** : Dépôt du dossier de candidature avant le **12 novembre 2026**.")

add_styled_heading("1.6 La Doctrine Stratégique Nice IA : 3 Choix Clairs", 2)
add_bullet("**1. Renoncer à l'hyperscale, assumer l’IA frugale** : Sobriété énergétique, cas d'usage utiles et sécurité maximale (conforme AI Act).")
add_bullet("**2. Occuper les 2 vides délaissés** : Adoption PME (commerce, tourisme, santé) et Industrialisation / LLMOps.")
add_bullet("**3. La Ville client n°1** : 5 cas d'usage municipaux industrialisés en 24 mois pour faire la preuve de la valeur et ancrer le récit politique.")

# Section 2 & Piliers
add_styled_heading("2. Pôle 1 : Axe Ville de Nice (Sécurité & Cadre de Vie)", 1)
add_styled_heading("1. CSU Augmenté & Vidéoprotection VSA (-65 % d'Incivilités)", 2)
add_p("Passage de la vidéosurveillance passive à l'alerte prédictive en temps réel sur 4 300+ caméras. Détection automatique des dépôts sauvages, dégradations et intrusions. Réduction du délai d'intervention de la PM de 45 min à moins de 6 min.")

add_styled_heading("2. Guichet Vocal Allo Niçois Séniors 24/7", 2)
add_p("Agent vocal souverain traitant 40 % des appels récurrents des aînés sans files d'attente téléphoniques, tout en maintenant un accompagnement humain prioritaire.")

add_styled_heading("3. Propreté Augmentée & Routage Voies Publiques", 2)
add_p("Analyse vidéo VSA embarquée sur véhicules municipaux pour cartographier les anomalies et déclencher les interventions sous 6 heures.")

add_styled_heading("3. Pôle 2 : Axe Métropole Nice Côte d'Azur (Budget & Monaco)", 1)
add_styled_heading("4. Audit IA de la Commande Publique Métropolitaine (+2,5 M€ / an Net)", 2)
add_p("Contrôle sémantique automatisé à 100 % des factures, devis et BPU (volume annuel 300 M€). Détection des doublons et surfacturations (0,5 % à 2,0 % d'erreurs filtrées).")

add_styled_heading("5. Bouclier Cybersécurité IA NIS 2 (1,8 M€ à 3 M€ Évités)", 2)
add_p("SOC métropolitain 24/7 armé d'agents IA autonomes. Partitionnement étanche du CSU (Air-Gap) et éradication du Shadow AI.")

add_styled_heading("6. Migration Monaco Cloud & Redondance IT (1,5 M€ à 2,8 M€ / an)", 2)
add_p("Migration souveraine vers Monaco Cloud (1er Cloud d'État UE). Économies directes de fonctionnement IT + 8 M€ à 12 M€ de CAPEX datacenter évité.")

# Section Annexes
add_styled_heading("4. ANNEXE : VALIDATION JURIDIQUE, STATISTIQUE ET RÉFÉRENCES DOCUMENTAIRES COMPLÈTES", 1)
add_bullet("**1. Note Stratégique « Opération Prométhée » — Le Grand Continent (Juillet 2026)** : Plan national de 700 Md$ (12 GW, 1 700 chercheurs).")
add_bullet("**2. Appel d'Offres Européen AI Gigafactories (30 Juillet 2026)** : EuroHPC JU (Lot 1) — 30 Md€ d'enveloppe, 500 M€ de subvention par site.")
add_bullet("**3. AI Act Européen — Règlement (UE) 2024/1689 du 13 juin 2024** : Encadrement des systèmes IA et exclusion du profilage biométrique en temps réel.")
add_bullet("**4. Loi n° 2023-380 du 19 mai 2023 (Loi JOP 2024 - Art. 10)** : Cadre expérimental VSA pour la détection automatisée d'événements.")
add_bullet("**5. Jurisprudence Conseil d'État (30 janvier 2026 - Commune de Nice)** : Cadre d'expérimentation municipale sous supervision d'intérêt public.")
add_bullet("**6. Directive Européenne NIS 2 (UE 2022/2555)** : Renforcement légal de la cybersécurité des systèmes métropolitains.")
add_bullet("**7. Rapport Institut 3IA Côte d'Azur & Université Côte d'Azur (2025/2026)** : > 100 chaires de recherche de niveau mondial.")
add_bullet("**8. Technopole Sophia Antipolis & Invest in Côte d'Azur (2025/2026)** : 2 700 entreprises, 46 000 emplois, 5 500 chercheurs.")
add_bullet("**9. Programme Extended Monaco & Monaco Cloud (gouv.mc / monacocloud.mc)** : Infrastructure souveraine certifiée AMSN.")
add_bullet("**10. Audit RTE / CNDP (2025)** : Analyse de la presqu'île électrique des Alpes-Maritimes et contraintes de charge.")

doc.save(docx_path)
print(f"Document Word/Google Docs sobre généré avec succès : {docx_path}")
