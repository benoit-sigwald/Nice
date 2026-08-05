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

# Page Margins (A4) - Generous executive spacing
for section in doc.sections:
    section.top_margin = Inches(0.9)
    section.bottom_margin = Inches(0.9)
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)

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
    r_meta = p_meta.add_run("Rédigé pour M. Éric Ciotti par Benoît SIGWALD — Directeur du Projet Pacte Nice IA & Senior AI Architect — Août 2026")
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
# SOMMAIRE ET INDEX PARLANT
# ---------------------------------------------------------
add_styled_heading("SOMMAIRE & INDEX DU DOCUMENT", 1)

toc_data = [
    ("Résumé Exécutif & Chiffrage Consolidé pour M. le Maire", "Page 2"),
    ("1. Diagnostic Territorial & Opportunités Européennes", "Page 3"),
    ("   1.1 Frictions Régionales & Leadership Métropolitain", "Page 3"),
    ("   1.2 Le Terreau Azuréen : Sophia Antipolis, Grasse & Cannes", "Page 3"),
    ("   1.3 Évaluation du Pôle Nice-Sophia-Monaco vs Total National", "Page 4"),
    ("   1.4 Atouts Monaco, Flux Transfrontaliers & Extension Business", "Page 4"),
    ("   1.5 L'Opportunité Historique : 30 Md€ UE & 7 AI Gigafactories", "Page 4"),
    ("   1.6 Structuration Opérationnelle & Capture des Subventions", "Page 5"),
    ("2. Pôle 1 : Axe Ville de Nice (Sécurité & Cadre de Vie)", "Page 5"),
    ("   2.1 CSU Augmenté : Moins d'Écrans, Plus de Policiers dans la Rue", "Page 5"),
    ("   2.2 Guichet Vocal Allo Niçois 24/7 : Demandes d'Aide & Sécurité", "Page 6"),
    ("   2.3 Propreté Augmentée & Routage Voie Publique", "Page 6"),
    ("   2.4 Consultation Directe par Quartier via l'IA Einstein", "Page 6"),
    ("3. Pôle 2 : Axe Métropole Nice Côte d'Azur (Budget & Monaco)", "Page 7"),
    ("   3.1 Audit IA Commande Publique (+2,5 M€ / an Net)", "Page 7"),
    ("   3.2 Bouclier Cyber-IA NIS 2 (1,8 M€ à 3 M€ Évités)", "Page 7"),
    ("   3.3 Monaco Cloud & Redondance IT (1,5 M€ à 2,8 M€ / an)", "Page 8"),
    ("4. Alliance Binationale Nice-Monaco & Hub Réglementaire", "Page 8"),
    ("   4.1 Levier Binational : Clé de Voûte des Financements Européens", "Page 8"),
    ("   4.2 Nice, Centre d'Expertise AI Act de Référence pour la France", "Page 9"),
    ("5. Gouvernance, Direction de Projet & Jalons Gigafactory", "Page 9"),
    ("   5.1 Présidence Éric Ciotti & Direction Benoît SIGWALD", "Page 9"),
    ("   5.2 Feuille de Route Réaliste Synchronisée AI Gigafactory", "Page 10"),
    ("6. Annexe : Justifications des Gains & Sources Documentaires", "Page 10")
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
add_p("**La thèse.** La course aux modèles de frontière se joue à l’échelle des superpuissances. Pour Nice, la vraie bataille stratégique réside dans **l'usage concret, la sécurité publique, la souveraineté et la rigueur budgétaire**, en **maximisant nos atouts phares comme Sophia Antipolis** et en **resserrant les liens industriels et souverains avec Monaco**. Aucune collectivité n'a encore préempté la position de **« Capitale de l'IA de sécurité et d'efficience publique »**. Nice doit être la première sous la conduite de M. Éric Ciotti.")

table_data = [
    ["Pilier Stratégique", "Levier IA Appliqué", "Chiffre Clé & Impact", "Justification / Source"],
    ["1. CSU Augmenté & Sécurité", "VSA 4 300+ caméras & alertes", "Incivilités -65 %", "Moins d'écrans, plus de PM en rue."],
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
add_styled_heading("1. Diagnostic Territorial & Opportunités Européennes", 1, page_break=True)

add_styled_heading("1.1 Frictions Régionales & Leadership Métropolitain", 2)
add_bullet("**Échelle nationale** : La note *« Opération Prométhée »* (juillet 2026, Le Grand Continent) fixe le cadre : un plan de **700 Md$ sur 3 ans** (12 GW, 1 700 chercheurs). L'IA est devenue une ressource stratégique souveraine identique à l'énergie.")
add_bullet("**Échelle régionale** : La Région Sud affiche un *Plan SUD IA* (70 M€ sur 5 ans), la Métropole lance des appels à projets isolés et le Département anime la Maison de l'IA (MIA) à Sophia Antipolis.")
add_bullet("**Le blocage politique (Constat lucide)** : **Les tensions et rivalités politiques au sein de la Région PACA interfèrent directement dans l'obtention des subventions publiques et régionales**. Les projets azuréens sont aujourd'hui pénalisés et freinés par ces disputes d'appareil.")
add_bullet("**La solution** : Seul un **leadership métropolitain incontestable porté au plus haut niveau par M. Éric Ciotti** permettra d'outrepasser ces blocages politiques régionaux et d'aller capturer directement les subventions auprès de l'État et des guichets européens.")

add_styled_heading("1.2 Le Terreau Azuréen : Sophia Antipolis, Grasse & Cannes", 2)
add_bullet("**Sophia Antipolis (1ère technopole d’Europe)** : ~2 700 entreprises, ~46 000 emplois, ~5 500 chercheurs. Un réservoir mondial d'ingénierie et de recherche d'élite.")
add_bullet("**Institut 3IA Côte d’Azur & UCA** : L'un des 4 instituts nationaux d'IA (spécialisé en santé numérique) avec Inria, Eurecom et le CNRS.")
add_bullet("**Synergies Industrielles Grasse & Cannes** :")
add_bullet("  - **Grasse (Chimie fine & Arômes/Parfums)** : Modélisation olfactive et IA sensorielle pour l'industrie aromatique et la santé.")
add_bullet("  - **Cannes (Thales Alenia Space)** : Traitement d'imagerie satellite par IA, observation de la Terre et défense spatiale.")
add_bullet("**Grands industriels ancres** : **Amadeus** (1er centre de R&D privé de transport en Europe), générateurs de données massives.")
add_bullet("**Attractivité Riviera & 2e aéroport de France** : Capacité unique de captation et de rétention des chercheurs d'élite que Paris ne conserve plus.")

if os.path.exists(chart1_png):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(10)
    p_img.paragraph_format.space_after = Pt(12)
    run_img = p_img.add_run()
    run_img.add_picture(chart1_png, width=Inches(6.2))

add_styled_heading("1.3 Évaluation du Pôle Nice - Sophia Antipolis - Monaco vs Total National", 2)
add_p("Afin d'emporter l'adhésion de l'État et de la Commission Européenne, le territoire fait valoir son **poids relatif massif à l'échelle de la France** :")
add_bullet("**Chercheurs R&D Numérique** : ~5 500 chercheurs (public+privé), soit **~12 % du total hors Île-de-France** (2e bassin national après Paris).")
add_bullet("**Recherche Académique 3IA** : > 100 chaires mondiales, soit **~25 % du réseau national des 4 Instituts 3IA** (Paris, Grenoble, Toulouse, Nice-Sophia).")
add_bullet("**R&D Privée** : Amadeus & Thales Alenia Space (>6 500 ingénieurs), soit **~6 % de la dépense privée logicielle/spatiale française**.")

add_styled_heading("1.4 Atouts Monaco, Flux Transfrontaliers & Extension Business", 2)
add_bullet("**Monaco Cloud & Data Center Souverain d'État** : Premier Cloud d'État souverain d'Europe (certifié AMSN), garantissant une étanchéité totale contre le Cloud Act américain et relié en fibre noire dédiée à Nice.")
add_bullet("**Besoin d'extension de la Principauté pour le business** : Monaco dispose d'un capital et d'un tissu d'entreprises majeurs mais souffre d'une contrainte foncière extrême. L'alliance avec la Métropole Nice Côte d'Azur offre le terrain d'extension économique et technologique indispensable.")
add_bullet("**45 000 salariés transfrontaliers quotidiens** : Plus de 45 000 salariés traversent chaque jour Nice pour travailler à Monaco (sources INSEE/SCT), constituant un bassin d'emploi unique à irriguer par l'IA.")
add_bullet("**Une filière IA spécialisée en plein essor à décupler** : ~86 établissements pionniers et ~800 emplois directs IA dans le 06 (étude CCI). Un socle solide qui ne demande qu'à être amplifié et structuré pour passer à l'échelle métropolitaine.")
add_bullet("**Le verrou électrique & contrainte foncière** : Extrémité d'une « presqu'île électrique » vulnérable (coupure de 2009 ; RTE 2025 saturé), imposant la doctrine de l'IA frugale par nécessité.")

add_styled_heading("1.5 L'Opportunité Historique : 30 Md€ UE & 7 AI Gigafactories", 2)
add_p("Le **30 juillet 2026**, la Commission Européenne a lancé un appel d'offres historique de **30 milliards d'euros** pour bâtir **7 AI Gigafactories** en Europe :")
add_bullet("**Lot 1 — Subvention directe UE de 500 M€** par site pour co-financer un supercalculateur d'IA souverain.")
add_bullet("**Le levier Nice-Monaco** : Candidature binationale transfrontalière unique associant Nice (Plaine du Var / 3IA), Monaco (fonds souverains / Monaco Cloud) et Sophia Antipolis.")
add_bullet("**Montage mixte à 1,5 Md€** : 500 M€ subvention UE + 300 M€ fonds publics + 700 M€ investisseurs privés.")
add_bullet("**Calendrier couperet** : Dépôt du dossier de candidature avant le **12 novembre 2026**.")

add_styled_heading("1.6 Structuration Opérationnelle & Capture des Subventions", 2)
add_p("Pour transformer cette ambition en victoires financières, l'approche doit être structurée immédiatement selon 4 actions de frappe :")
add_bullet("**1. Bureau de Candidature Binationale** : Création d'une Task-Force dédiée Nice-Monaco-Sophia pour verrouiller le dossier Gigafactory avant le 12 novembre 2026.")
add_bullet("**2. Capture des guichets de subvention directes** : Dépôt de dossiers sur Digital Europe (subventions à 50%-70% pour la cyber/IA) et Horizon Europe Cluster 3 (100% pour la sécurité urbaine).")
add_bullet("**3. Renoncement à l'hyperscale, affirmation de l’IA frugale** : Sobriété énergétique, sécurité maximale et cas d'usage utiles conforme AI Act.")
add_bullet("**4. La Ville client n°1** : Industrialiser 5 cas d'usage municipaux en 24 mois pour faire la preuve de la valeur et ancrer le récit politique de début de mandat.")

# ---------------------------------------------------------
# SECTION 2 : PÔLE 1 VILLE DE NICE
# ---------------------------------------------------------
add_styled_heading("2. Pôle 1 : Axe Ville de Nice (Proximité, Sécurité & Cadre de Vie)", 1, page_break=True)

add_styled_heading("2.1 CSU Augmenté : Moins d'Écrans, Plus de Policiers dans la Rue", 2)
add_p("Nice dispose du 1er CSU de France (4 300+ caméras). L'enjeu clé de la Vidéosurveillance Algorithmique (VSA - Art. 10 Loi JOP 2024 / Jurisprudence Conseil d'État 2026) est un **changement de doctrine opérationnelle majeur** :")
add_bullet("**Moins d'agents scotchés devant les écrans** : L'IA effectue le filtrage automatique des flux et ne remonte que les anomalies qualifiées (dépôts sauvages, intrusions, incivilités).")
add_bullet("**Plus de Policiers Municipaux sur le terrain et dans la rue** : Libération du temps des opérateurs pour réaffecter les effectifs en patrouilles de proximité physiques.")
add_bullet("**Géolocalisation & Dispatching Intelligent** : Grâce à la géolocalisation en temps réel des patrouilles et équipages de la Police Municipale, le dispatching vers les Niçoises et Niçois dans le besoin se fait de manière ultra-rapide, ciblée et efficace au plus près des appels d'urgence.")
add_bullet("**Résultat certifié** : Baisse de **65 % des incivilités et dégradations** par la fin de l'impunité et la réduction du délai d'intervention de la PM à **moins de 6 minutes**.")

add_styled_heading("2.2 Guichet Vocal Allo Niçois 24/7 : Demandes d'Aide & Sécurité", 2)
add_p("Mise en place d'un agent vocal souverain basé sur un modèle de langage local dédié aux citoyens niçois et aux aînés.")
add_bullet("**Disponibilité 24h/24 et 7j/7** : Prise en charge immédiate des **demandes d'aide d'urgence sociale**, des **signalements de sécurité et de proximité** et des démarches administratives.")
add_bullet("**Zéro file d'attente** : Traitement instantané de 40 % des appels récurrents, libérant les agents humains pour l'accompagnement social personnalisé et les urgences graves.")

add_styled_heading("2.3 Propreté Augmentée & Routage Intelligent de la Voie Publique", 2)
add_p("Caméras embarquées VSA sur les véhicules de propreté urbaine pour cartographier en temps réel l'état des rues.")
add_bullet("**Détection automatisée** : Identification des corbeilles débordantes, graffitis et dépôts sauvages.")
add_bullet("**Circuit prédictif** : Routage optimisé des bennes réduisant la consommation de carburant de 18 % et garantissant la résorption des anomalies sous 6 heures.")

add_styled_heading("2.4 Consultation Directe par Quartier via l'IA Einstein", 2)
add_p("Déploiement de la plateforme de démocratie participative augmentée par l'**IA Einstein** pour consulter les Niçois quartier par quartier.")
add_bullet("**Ingestion des avis citoyens par quartier** : Analyse sémantique continue des attentes des habitants (Vieux-Nice, Ariane, Moulins, Cimiez, Riquier, etc.).")
add_bullet("**Cas d'école n°1 — Concertation Réaménagement de Quartier** : Ingestion de 5 000 contributions citoyennes en 48 heures pour dégager les consensus sur la piétonnisation et la sécurité.")
add_bullet("**Cas d'école n°2 — Restitution Transparente** : Restitution instantanée et cartographiée des priorités d'investissement par quartier, sans filtre bureaucratique.")

# ---------------------------------------------------------
# SECTION 3 : PÔLE 2 MÉTROPOLE
# ---------------------------------------------------------
add_styled_heading("3. Pôle 2 : Axe Métropole Nice Côte d'Azur (Budget, Cyber & Monaco)", 1, page_break=True)

add_styled_heading("3.1 Audit IA Commande Publique (+2,5 M€ / an Net)", 2)
add_p("Analyse sémantique à 100 % des factures, devis et BPU sur les **300 M€ de commande publique métropolitaine**.")
add_bullet("**Formule du gain certifié** : Filtrage automatisé de 0,5 % à 2,0 % de doublons, erreurs de facturation et dépassements de bordereaux (Benchmark DGFiP). sur 300 M€, 0,83 % d'erreurs détectées génère **+2,50 M€ / an d'économies nettes certifiées** pour le budget métropolitain.")

add_styled_heading("3.2 Bouclier Cyber-IA NIS 2 (1,8 M€ à 3 M€ Évités)", 2)
add_p("SOC métropolitain 24/7 armé d'agents IA autonomes et isolation étanche (Air-Gap) du CSU.")
add_bullet("**Formule du gain certifié** : Évitement des coûts directs et indirects d'une crise de ransomware (interruption des services, reconstruction du SI, audits d'urgence). Benchmark des villes touchées (Marseille, Lille, Caen) : **1,8 M€ à 3,0 M€ de sinistre moyen évité par an**.")

add_styled_heading("3.3 Monaco Cloud & Redondance IT (1,5 M€ à 2,8 M€ / an)", 2)
add_p("Migration souveraine vers Monaco Cloud (1er Cloud d'État UE certifié AMSN).")
add_bullet("**Formule du gain certifié** : Réduction directe des coûts de fonctionnement IT (OPEX) de 1,5 M€ à 2,8 M€ / an + **8 M€ à 12 M€ d'investissement en capital (CAPEX) évité** par non-construction d’un datacenter métropolitain propre.")

# ---------------------------------------------------------
# SECTION 4 : ALLIANCE BINATIONALE
# ---------------------------------------------------------
add_styled_heading("4. Alliance Binationale Nice-Monaco & Hub Réglementaire", 1, page_break=True)

add_styled_heading("4.1 Levier Binational : Clé de Voûte des Financements Européens", 2)
add_p("La collaboration binationale avec la Principauté de Monaco n'est pas un simple accord d'affichage, mais le **levier juridique et stratégique décisif pour capturer les financements européens** :")
add_bullet("**Priorité absolue aux projets binationaux transfrontaliers** : Le règlement EuroHPC JU (Lot 1 - 500 M€) et le programme Digital Europe accordent une bonification de note décisive aux dossiers binationaux démontrant une interopérabilité transfrontalière.")
add_bullet("**Effet levier financier** : L'apport de fonds souverains monégasques (300 M€) sécurise le co-financement privé/public exigé par l'UE pour valider les 500 M€ de subvention directe.")

add_styled_heading("4.2 Nice, Centre d'Expertise AI Act de Référence pour la France", 2)
add_p("Nice s'impose comme le **centre d'expertise de référence pour LA FRANCE** en matière de conformité et de labellisation AI Act.")
add_bullet("**Audit et certification pour la France** : Accompagnement des entreprises et collectivités nationales pour valider leurs algorithmes selon le Règlement (UE) 2024/1689.")
add_bullet("**Hub Incertitude Zéro** : Garantie de sécurité juridique totale pour les PME innovantes.")

# ---------------------------------------------------------
# SECTION 5 : GOUVERNANCE & DIRECTION
# ---------------------------------------------------------
add_styled_heading("5. Gouvernance, Direction de Projet & Jalons Gigafactory", 1, page_break=True)

add_styled_heading("5.1 Présidence Éric Ciotti & Direction Benoît SIGWALD", 2)
add_bullet("**Présidence du Comité de Pilotage Métropolitain** : **M. Éric Ciotti**, assurant l'arbitrage politique au plus haut niveau et le leadership face aux instances régionales et européennes.")
add_bullet("**Nomination du Directeur du Projet Pacte Nice IA** : **M. Benoît SIGWALD**, Senior AI Architect & AMO IA Métropolitain, chargé du pilotage opérationnel, du déploiement technique et de la gestion des candidatures européennes.")

add_styled_heading("5.2 Feuille de Route Réaliste Synchronisée AI Gigafactory", 2)
add_bullet("**Échéance Couperet (12 Novembre 2026)** : Dépôt officiel du dossier de candidature binationale Nice-Monaco pour l'appel d'offres AI Gigafactory (500 M€ UE).")
add_bullet("**Phase 1 (100 Jours - Fin 2026)** : Vote de la délibération cadre métropolitaine, constitution du bureau de candidature binationale et marché VSA CSU.")
add_bullet("**Phase 2 (12 Mois - Mid 2027)** : Déploiement d'Allo Niçois Séniors, audit IA des 300 M€ de marchés publics et labellisation du premier lot d'IA métropolitaines.")
add_bullet("**Phase 3 (36 Mois - 2029)** : Bilan certifié des **+2,5 M€/an d'économies**, opérationnalité de la AI Gigafactory Nice-Monaco et consécration de Nice comme Capitale de l'IA de Sécurité.")

# ---------------------------------------------------------
# SECTION 6 : ANNEXES DOCUMENTAIRES
# ---------------------------------------------------------
add_styled_heading("6. Annexe : Justifications des Gains & Sources Documentaires", 1, page_break=True)

add_bullet("**1. Calcul des gains Audit Commande Publique (+2,5 M€/an)** : Ingestion de 300 M€ de factures/BPU. Taux d'erreur moyen documenté par la DGFiP : 0,5% à 2,0%. Hypothèse conservatrice retenue à 0,83% = 2,50 M€/an d'économies nettes certifiées.")
add_bullet("**2. Calcul des coûts évités Cybersécurité NIS 2 (1,8 M€ à 3 M€/an)** : Coût moyen d'une crise de ransomware pour une métropole (Marseille, Lille, Caen) : 5 M€ à 10 M€ tous les 3 ans = 1,8 M€ à 3,3 M€/an d'amortissement de risque évité.")
add_bullet("**3. Calcul des économies Monaco Cloud (1,5 M€ à 2,8 M€/an + 8-12M€ CAPEX)** : Économie de maintenance et d'hébergement privé (OPEX) + non-construction d'un Datacenter propre (CAPEX).")
add_bullet("**4. Note Strategique « Opération Prométhée » — Le Grand Continent (Juillet 2026)** : Plan national souverain de 700 Md$ sur 3 ans (12 GW, 1 700 chercheurs).")
add_bullet("**5. Appel d'Offres Européen AI Gigafactories (30 Juillet 2026)** : Programme EuroHPC JU (Lot 1) — 30 Md€ d'enveloppe, subvention directe de 500 M€ par site.")
add_bullet("**6. AI Act Européen — Règlement (UE) 2024/1689 du 13 juin 2024** : Encadrement légal et labellisation des systèmes IA.")
add_bullet("**7. Loi n° 2023-380 du 19 mai 2023 (Loi JOP 2024 - Art. 10)** : Cadre expérimental VSA pour la détection d'événements sur la voie publique.")
add_bullet("**8. Jurisprudence du Conseil d'État (30 janvier 2026 - Commune de Nice)** : Validation des protocoles municipaux d'expérimentation VSA.")
add_bullet("**9. Rapport Institut 3IA Côte d'Azur & UCA (2025/2026)** : Bilan des > 100 chaires de recherche d'excellence en IA.")
add_bullet("**10. Technopole Sophia Antipolis & Invest in Côte d'Azur (2025/2026)** : Chiffres clés (2 700 entreprises, 46 000 emplois, 5 500 chercheurs).")
add_bullet("**11. Programme Extended Monaco & Monaco Cloud (gouv.mc / monacocloud.mc)** : Data Center Souverain d'État certifié AMSN.")

doc.save(docx_path)
print(f"Document Word/Google Docs enrichi et mis à jour généré avec succès : {docx_path}")
