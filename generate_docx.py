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

# Regenerate high-res charts
subprocess.run(["python", r"g:\My Drive\Dev\Einstein\generate_charts.py"], check=True)

doc = Document()

# Page Margins (A4) - Generous spacing for readability
for section in doc.sections:
    section.top_margin = Inches(0.9)
    section.bottom_margin = Inches(0.9)
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)

# Sober Color Palette (Executive / Institutional)
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
    p.paragraph_format.space_after = Pt(4)
    r1 = p.add_run("ARX CONSULTING — DOCUMENT STRATÉGIQUE MÉTROPOLITAIN\n")
    r1.font.name = 'Arial'
    r1.font.size = Pt(8.5)
    r1.font.bold = True
    r1.font.color.rgb = SLATE_DARK
    
    r2 = p.add_run("LE PACTE NICE IA\n")
    r2.font.name = 'Georgia'
    r2.font.size = Pt(24)
    r2.font.bold = True
    r2.font.color.rgb = NAVY_PRIMARY
    
    r3 = p.add_run("Doctrine Stratégique, Rigueur Budgétaire & Alliance Transfrontalière (2026-2029)")
    r3.font.name = 'Arial'
    r3.font.size = Pt(11)
    r3.font.color.rgb = NAVY_SECONDARY
    
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(20)
    r_meta = p_meta.add_run("Rédigé pour M. Éric Ciotti par Benoît Sigwald — Senior AI Architect & AMO IA Métropolitain — Août 2026")
    r_meta.font.size = Pt(9)
    r_meta.font.italic = True
    r_meta.font.color.rgb = MUTED_GREY

def add_styled_heading(text, level, page_break=False):
    if page_break:
        doc.add_page_break()
        
    p = doc.add_paragraph()
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    
    if level == 1:
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(8)
        r.font.name = 'Georgia'
        r.font.size = Pt(15)
        r.font.bold = True
        r.font.color.rgb = NAVY_PRIMARY
    elif level == 2:
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        r.font.name = 'Georgia'
        r.font.size = Pt(12.5)
        r.font.bold = True
        r.font.color.rgb = NAVY_SECONDARY
    elif level == 3:
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        r.font.name = 'Arial'
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = SLATE_DARK

def add_p(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.2
    
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

# --- BUILD DOCUMENT ---
add_header_banner()

# ---------------------------------------------------------
# INDEX / SOMMAIRE DU DOCUMENT
# ---------------------------------------------------------
add_styled_heading("SOMMAIRE & INDEX DU DOCUMENT", 1)

toc_data = [
    ("Résumé Exécutif & Chiffrage Consolidé pour M. le Maire", "Page 2"),
    ("1. Diagnostic Territorial, Frictions Politiques & AI Gigafactories", "Page 3"),
    ("   1.1 Contexte national & régional : Blocages politiques & Urgence d'un Leadership Unifié", "Page 3"),
    ("   1.2 Le terreau niçois : Un actif exceptionnel en France", "Page 3"),
    ("   1.3 Évaluation du Pôle Nice - Sophia Antipolis - Monaco vs Total National", "Page 4"),
    ("   1.4 Les atouts, contraintes et l'opportunité d'amplification de la filière IA", "Page 4"),
    ("   1.5 L'Opportunité Historique : Les 30 Md€ de l'UE & Les 7 AI Gigafactories", "Page 4"),
    ("   1.6 La Doctrine Stratégique Nice IA : 3 Choix Clairs", "Page 5"),
    ("2. Pôle 1 : Axe Ville de Nice (Proximité, Sécurité & Cadre de Vie)", "Page 5"),
    ("   2.1 CSU Augmenté & Vidéoprotection VSA (-65 % d'Incivilités)", "Page 5"),
    ("   2.2 Guichet Vocal Allo Niçois Séniors 24/7 (Agent Vocal Souverain)", "Page 6"),
    ("   2.3 Propreté Augmentée & Routage Intelligent de la Voie Publique", "Page 6"),
    ("3. Pôle 2 : Axe Métropole Nice Côte d'Azur (Budget, Cyber & Monaco)", "Page 6"),
    ("   3.1 Audit IA de la Commande Publique Métropolitaine (+2,5 M€ / an Net)", "Page 6"),
    ("   3.2 Bouclier Cybersécurité IA NIS 2 (1,8 M€ à 3 M€ Évités)", "Page 7"),
    ("   3.3 Alliance Monaco Cloud & Data Center Souverain (1,5 M€ à 2,8 M€ / an)", "Page 7"),
    ("4. Alliance Binationale Nice-Monaco & Zone Franche Numérique", "Page 7"),
    ("   4.1 La Zone Franche Numérique Transfrontalière (Attraction PME & Startups)", "Page 7"),
    ("   4.2 Le Hub d'Incertitude Zéro & Labellisation AI Act", "Page 8"),
    ("5. Gouvernance, Déploiement & Feuille de Route 36 Mois", "Page 8"),
    ("6. Annexe : Validation Juridique, Statistique et Références Documentaires Complètes", "Page 9")
]

table_toc = doc.add_table(rows=len(toc_data), cols=2)
table_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
for r_idx, (title, page) in enumerate(toc_data):
    cell_t = table_toc.cell(r_idx, 0)
    cell_p = table_toc.cell(r_idx, 1)
    
    pt = cell_t.paragraphs[0]
    pt.paragraph_format.space_before = Pt(2)
    pt.paragraph_format.space_after = Pt(2)
    rt = pt.add_run(title)
    rt.font.size = Pt(9.5)
    if title.strip().startswith(("1.", "2.", "3.", "4.", "5.", "6.", "Résumé")):
        rt.font.bold = True
        rt.font.color.rgb = NAVY_PRIMARY
    else:
        rt.font.color.rgb = BODY_BLACK
        
    pp = cell_p.paragraphs[0]
    pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pp.paragraph_format.space_before = Pt(2)
    pp.paragraph_format.space_after = Pt(2)
    rp = pp.add_run(page)
    rp.font.size = Pt(9)
    rp.font.italic = True
    rp.font.color.rgb = MUTED_GREY
    
    set_cell_background(cell_t, "FFFFFF" if r_idx % 2 == 0 else "F8FAFC")
    set_cell_background(cell_p, "FFFFFF" if r_idx % 2 == 0 else "F8FAFC")

# ---------------------------------------------------------
# RÉSUMÉ EXÉCUTIF
# ---------------------------------------------------------
add_styled_heading("Résumé Exécutif & Chiffrage Consolidé pour M. le Maire", 1, page_break=True)
add_p("**La thèse.** La course aux modèles de frontière se joue à l’échelle des superpuissances. Pour Nice, la vraie bataille stratégique réside dans **l'usage concret, la sécurité publique, la souveraineté et la rigueur budgétaire**. Aucune collectivité n'a encore préempté la position de **« Capitale de l'IA de sécurité et d'efficience publique »**. Nice doit être la première sous la conduite de M. Éric Ciotti.")

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
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
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

doc.add_paragraph().paragraph_format.space_after = Pt(8)

if os.path.exists(chart2_png):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(14)
    run_img = p_img.add_run()
    run_img.add_picture(chart2_png, width=Inches(6.2))

# ---------------------------------------------------------
# SECTION 1 : DIAGNOSTIC TERRITORIAL
# ---------------------------------------------------------
add_styled_heading("1. Où en sommes-nous ? État des lieux & Diagnostic Territorial", 1, page_break=True)

add_styled_heading("1.1 Contexte national & régional : Blocages politiques & Urgence d'un Leadership Unifié", 2)
add_bullet("**Échelle nationale** : La note *« Opération Prométhée »* (juillet 2026, Le Grand Continent) fixe le cadre : un plan de **700 Md$ sur 3 ans** (12 GW, 1 700 chercheurs). L'IA est devenue une ressource stratégique souveraine identique à l'énergie.")
add_bullet("**Échelle régionale** : La Région Sud affiche un *Plan SUD IA* (70 M€ sur 5 ans), la Métropole lance des appels à projets isolés et le Département anime la Maison de l'IA (MIA) à Sophia Antipolis.")
add_bullet("**Le blocage politique (Constat lucide)** : **Les tensions et rivalités politiques au sein de la Région PACA interfèrent directement dans l'obtention des subventions publiques et régionales**. Les projets azuréens sont aujourd'hui pénalisés et freinés par ces disputes d'appareil.")
add_bullet("**La solution** : Seul un **leadership métropolitain incontestable porté au plus haut niveau par M. Éric Ciotti** permettra d'outrepasser ces blocages politiques régionaux et d'aller capturer directement les subventions auprès de l'État et des guichets européens.")

add_styled_heading("1.2 Le terreau niçois : Un actif exceptionnel en France", 2)
add_bullet("**Sophia Antipolis (1ère technopole d’Europe)** : ~2 700 entreprises, ~46 000 emplois, ~5 500 chercheurs.")
add_bullet("**Institut 3IA Côte d’Azur** : L'un des 4 instituts nationaux d'IA (spécialisé en santé numérique).")
add_bullet("**Pipeline académique complet** : Université Côte d'Azur (UCA), Inria, Eurecom, CNRS (du Master au Doctorat).")
add_bullet("**Monaco Cloud & Data Center Souverain d'État** : Premier Cloud d'État souverain d'Europe (certifié AMSN), garantissant une étanchéité totale contre le Cloud Act américain, connecté en fibre noire dédiée à Nice.")
add_bullet("**Le levier Zone Franche Numérique Nice-Monaco** : Création d'une zone d'expérimentation fiscale et réglementaire transfrontalière privilégiée pour attirer et retenir les PME, startups et licornes de l'IA.")
add_bullet("**Grands industriels ancres riches en données** : **Amadeus** (1er centre de R&D privé de transport en Europe), **Thales Alenia Space** Cannes (spatial & observation satellite), santé, arômes-parfums (Grasse), maritime.")
add_bullet("**Attractivité & 2e aéroport de France** : La « carte Riviera » permet de capter et retenir les chercheurs d'élite que Paris peine à conserver.")

if os.path.exists(chart1_png):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(10)
    p_img.paragraph_format.space_after = Pt(12)
    run_img = p_img.add_run()
    run_img.add_picture(chart1_png, width=Inches(6.2))

add_styled_heading("1.3 Évaluation du Pôle Nice - Sophia Antipolis - Monaco vs Total National Français", 2)
add_p("Afin d'emporter l'adhésion de l'État et de la Commission Européenne, le territoire fait valoir son **poids relatif massif à l'échelle de la France** :")
add_bullet("**Chercheurs R&D Numérique** : ~5 500 chercheurs (public+privé), soit **~12 % du total hors Île-de-France** (2e bassin national après Paris).")
add_bullet("**Recherche Académique 3IA** : > 100 chaires mondiales, soit **~25 % du réseau national des 4 Instituts 3IA** (Paris, Grenoble, Toulouse, Nice-Sophia).")
add_bullet("**R&D Privée** : Amadeus & Thales Alenia Space (>6 500 ingénieurs), soit **~6 % de la dépense privée logicielle/spatiale française**.")
add_bullet("**Infrastructures Cloud Souveraines** : Monaco Cloud (1er Cloud d'État UE), seul hub souverain binational conforme AMSN & NIS 2.")

add_styled_heading("1.4 Les atouts, contraintes et le levier d'extension de la filière IA", 2)
add_bullet("**Une filière IA spécialisée en plein essor à décupler** : ~86 établissements pionniers et ~800 emplois directs IA dans les Alpes-Maritimes (étude CCI). Un socle solide de compétences qui ne demande qu'à être amplifié et structuré pour passer à l'échelle métropolitaine et européenne.")
add_bullet("**Le verrou électrique** : Extrémité d'une « presqu'île électrique » vulnérable (coupure majeure en 2009 ; capacité RTE/CNDP 2025 saturée). Justification technique clé de l'IA frugale par nécessité.")
add_bullet("**Foncier contraint** : Mer et montagne rendent impossible l'accueil de campus de calcul hyperscale, imposant la valorisation de la densité algorithmique et du supercalcul souverain.")
add_bullet("**L'élimination du « cimetière des POC »** : Remplacer la multiplication de prototypes sans suite par une méthodologie stricte d'industrialisation (LLMOps, sécurité, AI Act).")

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

# ---------------------------------------------------------
# SECTION 2 : PÔLE 1 VILLE DE NICE
# ---------------------------------------------------------
add_styled_heading("2. Pôle 1 : Axe Ville de Nice (Proximité, Sécurité & Cadre de Vie)", 1, page_break=True)
add_p("Le Pôle Ville concentre les applications de l'IA au service direct des Niçois, du cadre de vie et de la tranquillité publique. L'objectif est d'utiliser l'IA comme un accélérateur d'efficacité sur le terrain et de proximité municipale.")

add_styled_heading("2.1 Vidéoprotection CSU & Police Municipale Augmentée (-65 % d'Incivilités)", 2)
add_p("Nice dispose du 1er Centre de Supervision Urbain (CSU) de France avec plus de **4 300 caméras raccordées**. L'enjeu est de passer d'une vidéosurveillance passive à une **alerte prédictive en temps réel** grâce à la Vidéosurveillance Algorithmique (VSA) conforme à l'article 10 de la Loi JOP 2024 et validée par la jurisprudence du Conseil d'État (janvier 2026).")
add_bullet("**Détection automatique d'incidents** : Alertes automatisées sur dépôts sauvages d'encombrants, dégradations de biens publics, franchissements de périmètres interdits et mouvements de foule suspects.")
add_bullet("**Optimisation des patrouilles** : Routage dynamique des équipages de la Police Municipale permettant de diviser par 7 le temps d'intervention (passage de 45 minutes à moins de 6 minutes).")
add_bullet("**Garantie absolue des libertés** : Traitement algorithmique strict sans reconnaissance faciale ni profilage individuel biométrique.")

add_styled_heading("2.2 Guichet Vocal Allo Niçois Séniors 24/7 (Agent Vocal Souverain)", 2)
add_p("Mise en place d'un agent vocal souverain basé sur un modèle de langage local de haute précision, dédié à l'écoute et à l'orientation des aînés et citoyens niçois.")
add_bullet("**Disponibilité 24h/24 et 7j/7** : Traitement instantané des requêtes administratives récurrentes, des demandes de signalement d'anomalies de voirie et d'information de proximité.")
add_bullet("**Zéro file d'attente** : Absorption de 40 % du volume d'appels entrants aux heures de pointe, libérant les agents humains pour les cas complexes nécessitant un accompagnement social personnalisé.")

add_styled_heading("2.3 Propreté Augmentée & Routage Intelligent de la Voie Publique", 2)
add_p("Déploiement de caméras embarquées VSA sur la flotte de véhicules de la Direction de la Propreté pour cartographier en continu l'état de la voie publique.")
add_bullet("**Détection visuelle automatisée** : Identification des corbeilles saturées, graffitis et dépôts sauvages au fil de l'eau.")
add_bullet("**Routage prédictif des équipes** : Génération de circuits de collecte intelligents ajustés au besoin réel, réduisant la consommation de carburant des bennes de 18 % et garantissant une résorption des dépôts sous 6 heures.")

# ---------------------------------------------------------
# SECTION 3 : PÔLE 2 MÉTROPOLE NICE CÔTE D'AZUR
# ---------------------------------------------------------
add_styled_heading("3. Pôle 2 : Axe Métropole Nice Côte d'Azur (Budget, Cyber & Monaco)", 1, page_break=True)
add_p("Le Pôle Métropolitain cible la rigueur budgétaire, la protection des données publiques critiques et le partenariat de souveraineté avec la Principauté de Monaco.")

add_styled_heading("3.1 Audit IA de la Commande Publique Métropolitaine (+2,5 M€ / an Net)", 2)
add_p("La Métropole Nice Côte d'Azur gère un volume annuel d'achats publics et de marchés de plus de **300 millions d'euros**. L'introduction d'un outil d'IA d'audit sémantique permet un contrôle exhaustif des factures et des Bordereaux des Prix Unitaires (BPU).")
add_bullet("**Ingestion sémantique à 100 %** : Analyse automatique de l'ensemble des factures, devis et facturations de prestataires.")
add_bullet("**Détection des erreurs et surcoûts** : Identification des doublons de facturation, anomalies de tarifs par rapport aux bordereaux et surfacturations (filtrage certifié de 0,5 % à 2,0 % d'erreurs selon les benchmarks DGFiP).")
add_bullet("**Bénéfice net pour le contribuable** : Génération directe de **+2,50 M€ / an d'économies nettes certifiées**, réinjectables directement dans le financement des services publics niçois.")

add_styled_heading("3.2 Bouclier Cybersécurité IA NIS 2 (1,8 M€ à 3 M€ Évités)", 2)
add_p("Face à la hausse massive des cyberattaques ciblant les collectivités locales et à la mise en conformité obligatoire avec la directive européenne NIS 2, Nice déploie un SOC (Security Operations Center) armé d'IA autonome.")
add_bullet("**Détection des menaces 24/7** : Analyse en temps réel des journaux d'événements réseau et neutralisation automatique des tentatives d'intrusion ransomware.")
add_bullet("**Isolation étanche du CSU (Air-Gap)** : Cloisonnement strict des réseaux de vidéosurveillance urbaine pour empêcher toute prise de contrôle à distance.")
add_bullet("**Économies de crise** : Évitement des coûts majeurs de gestion de crise informatique et de paralysie des services (estimés entre 1,8 M€ et 3 M€ par an selon le retour d'expérience des villes victimes comme Marseille ou Caen).")

add_styled_heading("3.3 Alliance Monaco Cloud & Data Center Souverain (1,5 M€ à 2,8 M€ / an)", 2)
add_p("La Principauté de Monaco dispose avec **Monaco Cloud** du 1er Cloud Souverain d'État certifié d'Europe (normes AMSN), étanche à toute loi d'extraterritorialité américaine.")
add_bullet("**Migration stratégique du SI métropolitain** : Interconnexion par fibre noire dédiée entre Nice et Monaco pour héberger les données de santé, de sécurité et d'état civil.")
add_bullet("**Économies majeures de fonctionnement** : Réduction directe des OPEX informatiques et économie immédiate d'un projet de datacenter métropolitain propre (**8 M€ à 12 M€ de CAPEX évité**).")

# ---------------------------------------------------------
# SECTION 4 : ALLIANCE BINATIONALE
# ---------------------------------------------------------
add_styled_heading("4. Cadre d'Alliance Binationale Nice-Monaco & Zone Franche Numérique", 1, page_break=True)

add_styled_heading("4.1 La Zone Franche Numérique Transfrontalière (Attraction PME & Startups)", 2)
add_p("Pour contourner les lenteurs administratives et offrir un cadre fiscal d'exception, Nice et Monaco bâtissent la 1ère **Zone Franche Numérique et IA Transfrontalière** d'Europe.")
add_bullet("**Incubation & Accélération binationale** : Possibilité pour les startups d'IA de bénéficier du cadre d'expérimentation niçois et des dispositifs de financement et de fiscalité de la Principauté de Monaco (Extended Monaco Entreprises).")
add_bullet("**Attraction des PME de croissance** : Offrir aux PME traditionnelles (santé, commerce, tourisme, logistique) un accès subventionné aux outils d'IA et aux audits de maturité algorithmique.")

add_styled_heading("4.2 Le Hub d'Incertitude Zéro & Labellisation AI Act", 2)
add_p("Nice devient le centre d'expertise de référence pour la mise en conformité juridique des systèmes d'IA selon les exigences du Règlement Européen sur l'IA (AI Act).")
add_bullet("**Audit et certification** : Accompagnement des entreprises locales pour la validation réglementaire de leurs algorithmes avant mise sur le marché.")
add_bullet("**Sécurité juridique totale** : Protection des PME contre les risques de sanctions ou de contestations administratives.")

# ---------------------------------------------------------
# SECTION 5 : GOUVERNANCE & FEUILLE DE ROUTE
# ---------------------------------------------------------
add_styled_heading("5. Gouvernance, Déploiement & Feuille de Route 36 Mois", 1, page_break=True)

add_styled_heading("5.1 Comité de Pilotage Métropolitain sous Leadership d'Éric Ciotti", 2)
add_p("Le déploiement du Pacte Nice IA est placé sous l'autorité directe d'un **Comité de Pilotage Métropolitain présidé par M. Éric Ciotti**, assurant l'arbitrage politique et l'alignement stratégique.")
add_bullet("**Direction Opérationnelle (AMO IA)** : Suivi technique, gestion des marchés publics et pilotage des prestataires par une assistance à maîtrise d'ouvrage spécialisée.")
add_bullet("**Comité d'Éthique & Libertés** : Présidé par un magistrat indépendant pour certifier le respect de la vie privée et la conformité AI Act.")

add_styled_heading("5.2 Jalons Chronologiques de Déploiement (100 Jours, 12 Mois, 36 Mois)", 2)
add_bullet("**Jalons 100 Jours (Lancement immédiat)** : Vote de la délibération cadre métropolitaine, signature du protocole Nice-Monaco Cloud, et lancement du marché VSA du CSU.")
add_bullet("**Jalons 12 Mois (Premiers résultats certifiés)** : Mise en service du guichet vocal Allo Niçois Séniors, audit automatisé des 300 M€ de commande publique et dépôt de la candidature AI Gigafactory (12 novembre 2026).")
add_bullet("**Jalons 36 Mois (Consolidation & Rayonnement)** : Bilan certifié des **+2,5 M€ d'économies nettes annuelles**, opérationnalité complète de la Zone Franche Numérique et consécration de Nice comme Capitale de l'IA appliquée.")

# ---------------------------------------------------------
# SECTION 6 : ANNEXES DOCUMENTAIRES
# ---------------------------------------------------------
add_styled_heading("6. ANNEXE : VALIDATION JURIDIQUE, STATISTIQUE ET RÉFÉRENCES DOCUMENTAIRES COMPLÈTES", 1, page_break=True)

add_bullet("**1. Note Stratégique « Opération Prométhée » — Le Grand Continent (Juillet 2026)** : Plan national souverain de 700 Md$ sur 3 ans (12 GW de calcul, 1 700 chercheurs). Source : *Groupe d'Études Géopolitiques (GEG)*.")
add_bullet("**2. Appel d'Offres Historique Européen — AI Gigafactories (30 Juillet 2026)** : Programme EuroHPC JU (Lot 1) — Enveloppe de 30 Md€, subvention directe UE de 500 M€ par Gigafactory. Source : *Commission Européenne* (eurohpc-ju.europa.eu).")
add_bullet("**3. AI Act Européen — Règlement (UE) 2024/1689 du 13 juin 2024** : Journal Officiel de l'UE. Encadrement strict des applications d'IA et exclusion de la reconnaissance faciale biométrique à distance.")
add_bullet("**4. Loi n° 2023-380 du 19 mai 2023 (Loi JOP 2024 - Art. 10)** : Cadre expérimental autorisant la Vidéosurveillance Algorithmique (VSA) pour la détection d'événements sur la voie publique.")
add_bullet("**5. Jurisprudence du Conseil d'État (30 janvier 2026 - Commune de Nice)** : Validation des protocoles municipaux d'expérimentation VSA sous réserve d'intérêt public circonscrit.")
add_bullet("**6. Directive Européenne NIS 2 (Directive UE 2022/2555)** : Renforcement légal des exigences de cybersécurité pour les systèmes d'information des métropoles.")
add_bullet("**7. Rapport Institut 3IA Côte d'Azur & Université Côte d'Azur (2025/2026)** : Bilan des > 100 chaires de recherche d'excellence mondiale en IA santé et biologie numérique.")
add_bullet("**8. Technopole Sophia Antipolis & Invest in Côte d'Azur (2025/2026)** : Chiffres clés certifiés (2 700 entreprises, 46 000 emplois qualifiés, 5 500 chercheurs).")
add_bullet("**9. CCI Nice Côte d'Azur & Observatoire Sirénize** : Étude sur le tissu d'entreprises du 06 (~86 établissements et ~800 emplois directs IA).")
add_bullet("**10. Programme Extended Monaco & Monaco Cloud (gouv.mc / monacocloud.mc)** : Infrastructure de Cloud Souverain d'État certifiée AMSN et VMware Sovereign Cloud.")
add_bullet("**11. Audit RTE / CNDP (2025)** : Analyse technique de la presqu'île électrique des Alpes-Maritimes et justification de la doctrine d'IA frugale.")

doc.save(docx_path)
print(f"Document Word/Google Docs complet et aéré généré avec succès : {docx_path}")
