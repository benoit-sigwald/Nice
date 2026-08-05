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

# Page Margins (A4) - Executive exact match
for section in doc.sections:
    section.top_margin = Inches(0.85)
    section.bottom_margin = Inches(0.85)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)

# Color Palette
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
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(2)
    r_sub = p_sub.add_run("ARX CONSULTING — DOCUMENT STRATÉGIQUE MÉTROPOLITAIN")
    r_sub.font.name = 'Arial'
    r_sub.font.size = Pt(8.5)
    r_sub.font.bold = True
    r_sub.font.color.rgb = SLATE_DARK

    p_title = doc.add_paragraph(style='Title')
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run("LE PACTE NICE IA")
    r_title.font.name = 'Georgia'
    r_title.font.size = Pt(26)
    r_title.font.bold = True
    r_title.font.color.rgb = NAVY_PRIMARY
    
    p_sub2 = doc.add_paragraph()
    p_sub2.paragraph_format.space_before = Pt(0)
    p_sub2.paragraph_format.space_after = Pt(6)
    r_sub2 = p_sub2.add_run("Doctrine Stratégique, Rigueur Budgétaire & Alliance Transfrontalière (2026-2029)")
    r_sub2.font.name = 'Arial'
    r_sub2.font.size = Pt(11)
    r_sub2.font.color.rgb = NAVY_SECONDARY
    
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_before = Pt(0)
    p_meta.paragraph_format.space_after = Pt(16)
    r_meta = p_meta.add_run("Rédigé pour M. Éric Ciotti par Benoît SIGWALD — Directeur du Projet Pacte Nice IA & Senior AI Architect — Août 2026")
    r_meta.font.name = 'Arial'
    r_meta.font.size = Pt(9)
    r_meta.font.italic = True
    r_meta.font.color.rgb = MUTED_GREY

def add_native_heading(text, level, page_break=False):
    if page_break:
        doc.add_page_break()
        
    if level == 1:
        p = doc.add_heading(text, level=1)
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.keep_with_next = True
        for r in p.runs:
            r.font.name = 'Georgia'
            r.font.size = Pt(15)
            r.font.bold = True
            r.font.color.rgb = NAVY_PRIMARY
    elif level == 2:
        p = doc.add_heading(text, level=2)
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        for r in p.runs:
            r.font.name = 'Georgia'
            r.font.size = Pt(12.5)
            r.font.bold = True
            r.font.color.rgb = NAVY_SECONDARY
    elif level == 3:
        p = doc.add_heading(text, level=3)
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        for r in p.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(10.5)
            r.font.bold = True
            r.font.color.rgb = SLATE_DARK

def add_bullet(text, level=1):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.15
    
    if level == 1:
        p.paragraph_format.left_indent = Inches(0.35)
    elif level == 2:
        p.paragraph_format.left_indent = Inches(0.65)
        
    parts = re.split(r'(\*\*.*?\*\*|\*.*?\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            r = p.add_run(part[2:-2])
            r.font.name = 'Arial'
            r.font.bold = True
            r.font.color.rgb = NAVY_PRIMARY
        elif part.startswith('*') and part.endswith('*'):
            r = p.add_run(part[1:-1])
            r.font.name = 'Arial'
            r.font.italic = True
        else:
            if part:
                r = p.add_run(part)
                r.font.name = 'Arial'
                r.font.color.rgb = BODY_BLACK

# --- BUILD DOCUMENT ---
add_header_banner()

# ---------------------------------------------------------
# INDEX PARFAITEMENT ALIGNÉ SELON LA RÉFÉRENCE
# ---------------------------------------------------------
add_native_heading("Index", 1)

toc_data = [
    ("Index", "1"),
    ("Résumé Exécutif & Chiffrage Consolidé pour M. le Maire", "1"),
    ("1. Diagnostic Territorial & Opportunités Européennes", "2"),
    ("1.1 Alignement Institutionnel & Leadership", "2"),
    ("1.2 Le Terreau Azuréen : Sophia Antipolis, Grasse, Cannes & Poids National", "2"),
    ("1.3 Atouts Monaco, Flux Transfrontaliers & Extension Business", "3"),
    ("1.4 L'Opportunité Historique : 30 Md€ UE & 7 AI Gigafactories", "4"),
    ("1.5 Structuration Opérationnelle & Capture des Subventions", "4"),
    ("2. Pôle 1 : Axe Ville de Nice (Sécurité & Cadre de Vie)", "5"),
    ("2.1 CSU Augmenté : Moins d'Écrans, Plus de Policiers dans la Rue", "5"),
    ("2.2 Guichet Vocal Allo Niçois 24/7 : Demandes d'Aide & Sécurité", "5"),
    ("2.3 Propreté Augmentée & Routage Intelligent de la Voie Publique", "5"),
    ("2.4 Consultation Directe par Quartier via l'IA Einstein", "5"),
    ("3. Pôle 2 : Axe Métropole Nice Côte d'Azur (Budget, Cyber & Monaco)", "7"),
    ("3.1 Audit IA Commande Publique (+2,5 M€ / an Net)", "7"),
    ("3.2 Bouclier Cyber-IA NIS 2 (1,8 M€ à 3 M€ Évités)", "7"),
    ("3.3 Monaco Cloud & Redondance IT (1,5 M€ à 2,8 M€ / an)", "7"),
    ("4. Alliance Binationale Nice-Monaco & Hub Réglementaire", "8"),
    ("4.1 Levier Binational : Clé de Voûte des Financements Européens", "8"),
    ("4.2 Nice, Centre d'Expertise AI Act de Référence pour la France", "8"),
    ("4.3 Pôle d'Excellence PME/PMI & Bibliothèque de Cas d'Usage Souverains", "8"),
    ("4.4 Pôle d'Excellence pour les Niçoises et les Niçois", "9"),
    ("5. Feuille de Route Précise (Sept 2026-2029) & Résumé Annuel", "9"),
    ("5.1 Présidence Éric Ciotti & Direction Benoît SIGWALD (2,5j/sem)", "9"),
    ("5.2 Échéancier Précis et Jalons par Date (Démarche 1er Septembre 2026)", "9"),
    ("5.3 Résumé Consolidé du Coût par Axe Institutionnel (Mairie, Métropole, Monaco)", "10"),
    ("6. Annexe : Justifications des Gains & Sources Documentaires", "11")
]

table_toc = doc.add_table(rows=len(toc_data), cols=2)
table_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
for r_idx, (title, page) in enumerate(toc_data):
    cell_t = table_toc.cell(r_idx, 0)
    cell_p = table_toc.cell(r_idx, 1)
    
    pt = cell_t.paragraphs[0]
    pt.paragraph_format.space_before = Pt(1.5)
    pt.paragraph_format.space_after = Pt(1.5)
    rt = pt.add_run(title)
    rt.font.name = 'Georgia' if title.startswith(("1.", "2.", "3.", "4.", "5.", "6.", "Résumé")) else 'Arial'
    rt.font.size = Pt(9)
    if title.startswith(("1.", "2.", "3.", "4.", "5.", "6.", "Résumé", "Index")):
        rt.font.bold = True
        rt.font.color.rgb = NAVY_PRIMARY
    else:
        rt.font.color.rgb = BODY_BLACK
        
    pp = cell_p.paragraphs[0]
    pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pp.paragraph_format.space_before = Pt(1.5)
    pp.paragraph_format.space_after = Pt(1.5)
    rp = pp.add_run(page)
    rp.font.name = 'Arial'
    rp.font.size = Pt(9)
    rp.font.bold = True
    rp.font.color.rgb = NAVY_PRIMARY

# ---------------------------------------------------------
# RÉSUMÉ EXÉCUTIF
# ---------------------------------------------------------
add_native_heading("Résumé Exécutif & Chiffrage Consolidé pour M. le Maire", 1)
add_bullet("**La Thèse Stratégique** : La course aux modèles de frontière se joue à l’échelle des superpuissances. Pour Nice, la vraie bataille stratégique réside dans **l'usage concret, la sécurité publique, la souveraineté et la rigueur budgétaire**, en **maximisant nos atouts phares comme Sophia Antipolis** et en **resserrant les liens industriels et souverains avec Monaco**.")
add_bullet("**La Vision Politique** : Aucune collectivité n'a encore préempté la position de **« Capitale de l'IA de sécurité et d'efficience publique »**. Nice doit être la première sous la conduite de M. Éric Ciotti.")

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
            r.font.name = 'Arial'
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(9)
        else:
            set_cell_background(cell, "F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
            r = p.add_run(val)
            r.font.name = 'Arial'
            r.font.size = Pt(8.5)
            r.font.color.rgb = BODY_BLACK

doc.add_paragraph().paragraph_format.space_after = Pt(6)

# ---------------------------------------------------------
# SECTION 1 : DIAGNOSTIC TERRITORIAL
# ---------------------------------------------------------
add_native_heading("1. Diagnostic Territorial & Opportunités Européennes", 1, page_break=True)

add_native_heading("1.1 Alignement Institutionnel & Leadership", 2)
add_bullet("**Échelle nationale** : La note *« Opération Prométhée »* (juillet 2026, Le Grand Continent) fixe le cadre : un plan de **700 Md$ sur 3 ans** (12 GW, 1 700 chercheurs). L'IA est devenue une ressource stratégique souveraine identique à l'énergie.")
add_bullet("**Échelle régionale** : La Région Sud affiche un *Plan SUD IA* (70 M€ sur 5 ans), la Métropole lance des appels à projets isolés et le Département anime la Maison de l'IA (MIA) à Sophia Antipolis.")
add_bullet("**Une gouvernance à unifier (Constat constructif)** : **Les frictions institutionnelles et les complexités d'arbitrage au sein de la Région PACA peuvent ralentir l'accès optimal aux subventions**. Un alignement direct et unifié est indispensable pour accélérer les financements.")
add_bullet("**La solution** : Seul un **leadership métropolitain incontestable porté au plus haut niveau par M. Éric Ciotti** permettra d'outrepasser ces frictions institutionnelles et d'aller capturer directement les subventions auprès de l'État et des guichets européens.")

add_native_heading("1.2 Le Terreau Azuréen : Sophia Antipolis, Grasse, Cannes & Poids National", 2)
add_bullet("**Sophia Antipolis (1ère technopole d’Europe)** : ~2 700 entreprises, ~46 000 emplois, ~5 500 chercheurs. Un réservoir mondial d'ingénierie et de recherche d'élite.")
add_bullet("**Sophia c’est ~25,6 % de tous les emplois technopolitains de France.**")
add_bullet("**Sophia c’est ~19,3 % du total des entreprises implantées dans ces structures.**")
add_bullet("**Le terreau azuréen représente seulement 1,66 % de la population Française.**")
add_bullet("**Institut 3IA Côte d’Azur & UCA** : L'un des 4 instituts nationaux d'IA (spécialisé en santé numérique) avec Inria, Eurecom et le CNRS (représentant **25 % du réseau national des 4 Instituts 3IA**).")
add_bullet("**Synergie Grasse (Arômes & Parfums)** : Modélisation olfactive, chimie fine et IA sensorielle pour l'industrie aromatique et la santé.")
add_bullet("**Synergie Cannes (Thales Alenia Space)** : Traitement d'imagerie satellite par IA, observation de la Terre et défense spatiale (représentant avec Amadeus **6 % de la dépense privée R&D logicielle/spatiale française**).")
add_bullet("**Chercheurs R&D Numérique** : ~5 500 chercheurs (public+privé), soit **~12 % du total national hors Île-de-France** (2e bassin français après Paris).")
add_bullet("**Attractivité Riviera & 2e aéroport de France** : Capacité unique de captation et de rétention des chercheurs d'élite que Paris ne conserve plus.")

if os.path.exists(chart1_png):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(10)
    run_img = p_img.add_run()
    run_img.add_picture(chart1_png, width=Inches(6.0))

add_native_heading("1.3 Atouts Monaco, Flux Transfrontaliers & Extension Business", 2)
add_bullet("**Monaco Cloud & Data Center Souverain d'État** : Premier Cloud d'État souverain d'Europe (certifié AMSN), garantissant une étanchéité totale contre le Cloud Act américain, connecté en fibre noire dédiée à Nice.")
add_bullet("**Besoin d'extension de la Principauté pour le business** : Monaco dispose d'un capital et d'un tissu d'entreprises majeurs mais souffre d'une contrainte foncière extrême. L'alliance avec la Métropole Nice Côte d'Azur offrirait le terrain d'extension économique et technologique indispensable.")
add_bullet("**45 000 salariés transfrontaliers quotidiens** : Plus de 45 000 salariés traversent chaque jour Nice pour travailler à Monaco (sources INSEE/SCT), constituant un bassin d'emploi unique à irriguer par l'IA. Réduire ces déplacements au quotidien permettra une amélioration de la qualité de vie et une maîtrise des émissions de CO2.")
add_bullet("**Une filière IA spécialisée en plein essor à décupler** : ~86 établissements pionniers et ~800 emplois directs IA dans le 06 (étude CCI). Un socle solide qui ne demande qu'à être amplifié et structuré pour passer à l'échelle métropolitaine.")
add_bullet("**Le verrou électrique & contrainte foncière** : Extrémité d'une « presqu'île électrique » vulnérable (coupure de 2009 ; RTE 2025 saturé), imposing la doctrine de l'IA frugale par nécessité (un data center nécessiterait des investissements RTE très importants).")

add_native_heading("1.4 L'Opportunité Historique : 30 Md€ UE & 7 AI Gigafactories", 2)
add_bullet("**L'Appel d'Offres Européen** : Le **30 juillet 2026**, la Commission Européenne a lancé un appel d'offres historique de **30 milliards d'euros** pour bâtir **7 AI Gigafactories** en Europe.")
add_bullet("**Lot 1 — Subvention directe UE de 500 M€** par site pour co-financer le développement de l'IA souverain.")
add_bullet("**Monaco Cloud est l'opportunité pour développer l'existant.**")
add_bullet("**Le levier Nice-Monaco** : Candidature binationale transfrontalière unique associant Nice (Plaine du Var / 3IA), Monaco (fonds souverains / Monaco Cloud) et Sophia Antipolis.")
add_bullet("**Montage mixte à 1,5 Md€** : 500 M€ subvention UE + 300 M€ fonds publics + 700 M€ investisseurs privés.")
add_bullet("**Calendrier couperet** : Dépôt du dossier de candidature avant le **12 novembre 2026**.")

add_native_heading("1.5 Structuration Opérationnelle & Capture des Subventions", 2)
add_bullet("**Méthodologie d'Action** : Pour transformer cette ambition en victoires financières, l'approche est structurée immédiatement selon 4 actions de frappe :")
add_bullet("**1. Bureau de Candidature Binationale** : Création d'une Task-Force dédiée Nice-Monaco-Sophia pour verrouiller le dossier Gigafactory avant le 12 novembre 2026.")
add_bullet("**2. Capture des guichets de subvention directes** : Dépôt de dossiers sur Digital Europe (subventions à 50%-70% pour la cyber/IA) et Horizon Europe Cluster 3 (100% pour la sécurité urbaine).")
add_bullet("**3. Renoncement à l'hyperscale, affirmation de l’IA frugale** : Sobriété énergétique, sécurité maximale et cas d'usage utiles conforme AI Act.")
add_bullet("**4. La Ville de Nice et la Métropole client n°1** : Industrialiser 5 cas d'usage municipaux en 24 mois pour faire la preuve de la valeur et ancrer le récit politique de début de mandat.")

# ---------------------------------------------------------
# SECTION 2 : PÔLE 1 VILLE DE NICE
# ---------------------------------------------------------
add_native_heading("2. Pôle 1 : Axe Ville de Nice (Sécurité & Cadre de Vie)", 1, page_break=True)
add_bullet("**Objectif Opérationnel** : Le Pôle Ville concentre les applications de l'IA au service direct des Niçois, du cadre de vie et de la tranquillité publique. L'objectif est d'utiliser l'IA comme un accélérateur d'efficacité sur le terrain et de proximité municipale.")

add_native_heading("2.1 CSU Augmenté : Moins d'Écrans, Plus de Policiers dans la Rue", 2)
add_bullet("**Le Constat Opérationnel** : Nice dispose du 1er CSU de France (4 300+ caméras). L'enjeu clé de la Vidéosurveillance Algorithmique (VSA - Art. 10 Loi JOP 2024 / Jurisprudence Conseil d'État 2026) est un **changement de doctrine opérationnelle majeur**.")
add_bullet("**Moins d'agents scotchés devant les écrans** : L'IA effectue le filtrage automatique des flux et ne remonte que les anomalies qualifiées (dépôts sauvages, intrusions, incivilités).")
add_bullet("**Plus de Policiers Municipaux sur le terrain et dans la rue** : Libération du temps des opérateurs pour réaffecter les effectifs en patrouilles de proximité physiques.")
add_bullet("**Géolocalisation & Dispatching Intelligent** : Grâce à la géolocalisation en temps réel des patrouilles et équipages de la Police Municipale, le dispatching vers les Niçoises et Niçois dans le besoin se fait de manière ultra-rapide, ciblée et efficace au plus près des appels d'urgence.")
add_bullet("**Résultat attendu** : Baisse de **65 % des incivilités et dégradations** par la fin de l'impunité et la réduction du délai d'intervention de la PM à **moins de 6 minutes**.")

add_native_heading("2.2 Guichet Vocal Allo Niçois 24/7 : Demandes d'Aide & Sécurité", 2)
add_bullet("**L'Agent Vocal Souverain** : Mise en place d'un agent vocal souverain basé sur un modèle de langage local dédié aux citoyens niçois et aux aînés.")
add_bullet("**Disponibilité 24h/24 et 7j/7** : Prise en charge immédiate des **demandes d'aide d'urgence sociale**, des **signalements de sécurité et de proximité** et des démarches administratives.")
add_bullet("**Zéro file d'attente** : Traitement instantané de 40 % des appels récurrents, libérant les agents humains pour l'accompagnement personnalisé et les urgences graves.")

add_native_heading("2.3 Propreté Augmentée & Routage Intelligent de la Voie Publique", 2)
add_bullet("**Capteurs Embarqués VSA** : Caméras embarquées VSA sur les véhicules de propreté urbaine pour cartographier en temps réel l'état des rues.")
add_bullet("**Détection automatisée** : Identification des corbeilles débordantes, graffitis et dépôts sauvages.")
add_bullet("**Circuit prédictif** : Routage optimisé des bennes réduisant la consommation de carburant de 18 % et garantissant la résorption des anomalies sous 6 heures.")

add_native_heading("2.4 Consultation Directe par Quartier via l'IA", 2)
add_bullet("**Démocratie Participative Augmentée** : Déploiement de la plateforme de démocratie participative augmentée par **l'IA** pour consulter les Niçois quartier par quartier.")
add_bullet("**Ingestion des avis citoyens par quartier** : Analyse sémantique continue des attentes des habitants (Vieux-Nice, Ariane, Moulins, Cimiez, Riquier, etc.).")
add_bullet("**Cas d'école n°1 — Concertation Réaménagement de Quartier** : Ingestion de 5 000 contributions citoyennes en 48 heures pour dégager les consensus sur la piétonnisation et la sécurité.")
add_bullet("**Cas d'école n°2 — Restitution Transparente** : Restitution instantanée et cartographiée des priorités d'investissement par quartier, sans filtre bureaucratique.")

# ---------------------------------------------------------
# SECTION 3 : PÔLE 2 MÉTROPOLE
# ---------------------------------------------------------
add_native_heading("3. Pôle 2 : Axe Métropole Nice Côte d'Azur (Budget, DSI & Monaco)", 1, page_break=True)

add_native_heading("3.1 Audit IA Commande Publique (+2,5 M€ / an Net)", 2)
add_bullet("**Ingestion des Marchés** : Analyse sémantique à 100 % des factures, devis et BPU sur les **300 M€ de commande publique métropolitaine**.")
add_bullet("**Formule du gain certifié** : Filtrage automatisé de 0,5 % à 2,0 % de doublons, erreurs de facturation et dépassements de bordereaux (Benchmark DGFiP). Sur 300 M€, 0,83 % d'erreurs détectées génère **+2,50 M€ / an d'économies nettes certifiées** pour le budget métropolitain.")

add_native_heading("3.2 Bouclier Cyber-IA NIS 2 (1,8 M€ à 3 M€ Évités)", 2)
add_bullet("**Protection Inviolable du CSU** : SOC métropolitain 24/7 armé d'agents IA autonomes et isolation étanche (Air-Gap) du CSU.")
add_bullet("**Formule du gain certifié** : Évitement des coûts directs et indirects d'une crise de ransomware (interruption des services, reconstruction du SI, audits d'urgence). Benchmark des villes touchées (Marseille, Lille, Caen) : **1,8 M€ à 3,0 M€ de cout moyen évité par an**.")

add_native_heading("3.3 Monaco Cloud & Redondance IT (1,5 M€ à 2,8 M€ / an)", 2)
add_bullet("**Partenariat d'Hébergement Souverain** : Migration souveraine vers Monaco Cloud (1er Cloud d'État UE certifié AMSN).")
add_bullet("**Formule du gain certifié** : Réduction directe des coûts de fonctionnement IT (OPEX) de 1,5 M€ à 2,8 M€ / an + **8 M€ à 12 M€ d'investissement en capital (CAPEX) évité** par non-construction d’un datacenter métropolitain propre.")

# ---------------------------------------------------------
# SECTION 4 : ALLIANCE BINATIONALE & PÔLE CITOYENS
# ---------------------------------------------------------
add_native_heading("4. Alliance Binationale Nice-Monaco & Hub Réglementaire", 1, page_break=True)

add_native_heading("4.1 Levier Binational : Clé de Voûte des Financements Européens", 2)
add_bullet("**L'Atout Juridique Binational** : La collaboration binationale avec la Principauté de Monaco n'est pas un simple accord d'affichage, mais le **levier juridique et stratégique décisif pour capturer les financements européens**.")
add_bullet("**Priorité absolue aux projets binationaux transfrontaliers** : Le règlement EuroHPC JU (Lot 1 - 500 M€) et le programme Digital Europe accordent une bonification de note décisive aux dossiers binationaux démontrant une interopérabilité transfrontalière.")
add_bullet("**Effet levier financier** : L'apport de fonds souverains monégasques (300 M€) sécurise le co-financement privé/public exigé par l'UE pour valider les 500 M€ de subvention directe.")

add_native_heading("4.2 Nice, Centre d'Expertise AI Act de Référence pour la France", 2)
add_bullet("**L'Expertise Réglementaire Nationale** : Nice s'impose comme le **centre d'expertise de référence pour LA FRANCE** en matière de conformité et de labellisation AI Act.")
add_bullet("**Audit et certification pour la France** : Accompagnement des entreprises et collectivités nationales pour valider leurs algorithmes selon le Règlement (UE) 2024/1689.")
add_bullet("**Hub Incertitude Zéro** : Garantie de sécurité juridique totale pour les PME innovantes.")

add_native_heading("4.3 Pôle d'Excellence PME/PMI & Bibliothèque de Cas d'Usage Souverains", 2)
add_bullet("**Guichet Unique PME/PMI** : Création du Pôle d'Excellence Métropolitain pour accompagner la transformation numérique des 200 PME/PMI clés des Alpes-Maritimes et de Monaco.")
add_bullet("**Bibliothèque de Cas d'Usage Réutilisables** : Mise à disposition d'un catalogue de briques d'IA souveraines pré-packagées (analyse sémantique de contrats, contrôle qualité vidéo pour la chimie/arômes à Grasse, maintenance prédictive spatiale pour Cannes).")
add_bullet("**Accélération de la Migration IA** : Réduction par 3 des coûts et des délais d'intégration grâce au réemploi des briques logicielles souveraines développées par la Métropole.")
add_bullet("**Accompagnement Financement UE (EDIH)** : Prise en charge jusqu'à 70 % des coûts de diagnostic et de migration IA via le guichet européen *Digital Europe* / EDIH Côte d'Azur.")

add_native_heading("4.4 Pôle d'Excellence pour les Niçoises et les Niçois", 2)
add_bullet("**Acculturation & Formations Gratuites dans les AnimaNice** : Déploiement d'ateliers hebdomadaires d'initiation et de maîtrise de l'IA pour les séniors, les familles et les jeunes au sein de l'ensemble du réseau des centres AnimaNice (Vieux-Nice, Ariane, Cimiez, Fabron, Riquier, etc.).")
add_bullet("**Charte d'Éthique & Protection Absolue des Données Personnelles** : Sanctification de la vie privée des Niçois. Aucune donnée citoyenne ou image vidéo n'est commercialisée ni stockée hors des serveurs souverains sous souveraineté exclusive (Monaco Cloud AMSN / Air-Gap Métropolitain).")
add_bullet("**Pass IA Jeunesse & Accompagnement Éducatif** : Partenariat avec l'Université Côte d'Azur pour mettre à disposition des étudiants, lycéens et collégiens niçois des accès souverains gratuits aux outils de recherche et d'ingénierie IA.")
add_bullet("**Transparence & Gouvernance Citoyenne** : Publication annuelle d'un bilan d'impact éthique et financier de l'IA municipale, librement accessible à tous les citoyens niçois.")

# ---------------------------------------------------------
# SECTION 5 : GOUVERNANCE, ROADMAP & BUDGET ANNUEL
# ---------------------------------------------------------
add_native_heading("5. Feuille de Route Précise (Sept 2026-2029) & Résumé Annuel", 1, page_break=True)

add_native_heading("5.1 Présidence Éric Ciotti", 2)
add_bullet("**Présidence du Comité de Pilotage Métropolitain** : **M. Éric Ciotti**, assurant l'arbitrage politique au plus haut niveau et le leadership face aux instances régionales et européennes.")
add_bullet("**Direction de Projet AMO IA (Temps Partagé)** : **M. Benoît SIGWALD**, Senior AI Architect. Engagement à **2,5 jours par semaine**.")

add_native_heading("5.2 Échéancier Précis et Jalons par Date (Démarche 1er Septembre 2026)", 2)

rm_data = [
    ["Période & Date Précise", "Chantier / Livrable Stratégique", "Budget Dédié", "Impact & Résultat Attendus"],
    ["1er Septembre 2026", "L1.1 Lancement Pacte & Direction Projet (2.5j/sem)", "16,25 k€", "Cadrage opérationnel & gouvernance."],
    ["15 Septembre 2026", "L1.2 Bureau Candidature Binationale Gigafactory", "35,00 k€", "Dossier 500 M€ UE co-rédigé Monaco."],
    ["1er Octobre 2026", "L1.3 Délibération Métropolitaine & AI Act", "20,00 k€", "Vote cadre juridique & éthique."],
    ["15 Octobre 2026", "L1.4 Audit CSU & Fibre Monaco Cloud", "38,75 k€", "Cahier des charges interconnexion."],
    ["12 Novembre 2026", "L2.1 Dépôt Dossier Gigafactory UE (COUPERET)", "10,00 k€", "Candidature officielle 500 M€ UE."],
    ["15 Décembre 2026", "L2.2 Filtrage VSA CSU (4 300 caméras)", "55,00 k€", "Alertes incivilités en temps réel."],
    ["15 Janvier 2027", "L2.3 Géolocalisation & Dispatching PM", "33,75 k€", "Patrouilles physiques < 6 min."],
    ["1er Février 2027", "L2.4 Ingestion Marchés Publics (Pilote 50M€)", "16,25 k€", "Ingénierie & recettes sémantiques."],
    ["15 Mars 2027", "L3.1 Audit Commande Publique Généralisé (300M€)", "45,00 k€", "Filtrage erreurs & doublons BPU."],
    ["1er Avril 2027", "L3.2 SOC Cyber NIS 2 Autonome Air-Gap", "28,75 k€", "Protection inviolable du CSU."],
    ["15 Mai 2027", "L3.3 IA Einstein Concertation (1ers Quartiers)", "15,00 k€", "Restitution citoyenne augmentée."],
    ["15 Juin 2027", "L4.1 Généralisation Audit Commande Publique", "45,00 k€", "Rentrée de +2,5 M€/an certifiés."],
    ["15 Juillet 2027", "L4.2 Recette Sécurité Monaco Cloud AMSN", "23,75 k€", "Certification Cloud Souverain."],
    ["31 Août 2027", "L4.3 Bilan Année 1 & Pôle Citoyens AnimaNice", "20,00 k€", "+2,5 M€ d'économies & bilan VSA."],
    ["Sept 2027 — Fév 2028", "L5.1 Guichet Vocal Allo Niçois Séniors 24/7", "85,00 k€", "Assistance 24/7 & d'urgence."],
    ["Sept 2027 — Fév 2028", "L5.2 VSA Propreté & Routage Bennes", "65,00 k€", "Rues propres & -18% carburant."],
    ["Mars 2028 — Août 2028", "L6.1 Pôle Excellence PME/PMI & Bibliothèque Cas Usage", "85,00 k€", "Migration IA de 200 PME/PMI."],
    ["Mars 2028 — Août 2028", "L6.2 Concertation Einstein Récurrente", "60,00 k€", "Ingestion avis par quartier."],
    ["1er Janvier 2029", "L7.1 Supercalculateur AI Gigafactory", "160,00 k€", "Mise en service opérationnelle."],
    ["Sept 2028 — Août 2029", "L7.2 Industrialisation IA Frugale & Cyber", "170,00 k€", "Pérennisation & Hub AI Act."]
]

table_rm = doc.add_table(rows=len(rm_data), cols=4)
table_rm.alignment = WD_TABLE_ALIGNMENT.CENTER
for r_idx, row in enumerate(rm_data):
    for c_idx, val in enumerate(row):
        cell = table_rm.cell(r_idx, c_idx)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        if r_idx == 0:
            set_cell_background(cell, "1E293B")
            r = p.add_run(val)
            r.font.name = 'Arial'
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(8.5)
        else:
            set_cell_background(cell, "F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
            r = p.add_run(val)
            r.font.name = 'Arial'
            r.font.size = Pt(8)
            r.font.color.rgb = BODY_BLACK

doc.add_paragraph().paragraph_format.space_after = Pt(10)

add_native_heading("5.3 Résumé Consolidé du Coût par Axe Institutionnel (Mairie, Métropole, Monaco)", 2)
add_bullet("**Structure Budgétaire Simplifiée par Entité** : Le budget annuel brut de **435 000 € / an** est réparti de manière transparente entre les 3 piliers d'action (Ville de Nice, Métropole Nice Côte d'Azur et Extension Monaco). Grâce au co-financement de 50 % de l'UE (*Digital Europe*), le reste à charge métropolitain est de seulement **217 500 € / an**, immédiatement compensé par **+2 500 000 € d'économies nettes certifiées**.")

t_data = [
    ["Pôle Institutionnel / Domaine", "Actions & Projets Clefs", "Budget Brut Annuel", "Co-financement UE (50%)", "Reste à Charge Net", "Gains Certifiés Net"],
    ["1. AXE MAIRIE (Ville de Nice)", "CSU VSA, Allo Niçois 24/7, Propreté, Consultation Einstein, Pôle Citoyens", "155 000 €", "77 500 €", "77 500 €", "Incivilités -65%, Interventions <6min"],
    ["2. AXE MÉTROPOLE (Nice Côte d'Azur)", "Audit Commande Publique (300M€), SOC Cyber NIS 2, Pôle PME/PMI & AMO (2,5j/sem)", "160 000 €", "80 000 €", "80 000 €", "+2 500 000 € / an Net (Marchés)"],
    ["3. AXE MONACO (Alliance & Extension)", "Fibre Monaco Cloud (AMSN), Zone Franche Numérique, Task-Force Gigafactory 500M€", "120 000 €", "60 000 €", "60 000 €", "1,5M€ à 2,8M€ IT + 8-12M€ CAPEX"],
    ["TOTAL ANNUEL CONSOLIDÉ", "Ensemble des 3 Piliers Institutionnels (2026-2029)", "435 000 €", "217 500 €", "217 500 €", "+2 500 000 € / an (Bénéfice +2,28 M€)"]
]

table_t = doc.add_table(rows=len(t_data), cols=6)
table_t.alignment = WD_TABLE_ALIGNMENT.CENTER
for r_idx, row in enumerate(t_data):
    for c_idx, val in enumerate(row):
        cell = table_t.cell(r_idx, c_idx)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        if r_idx == 0:
            set_cell_background(cell, "1E293B")
            r = p.add_run(val)
            r.font.name = 'Arial'
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(8)
        elif "TOTAL" in row[0]:
            set_cell_background(cell, "0F172A")
            r = p.add_run(val)
            r.font.name = 'Arial'
            r.font.bold = True
            r.font.size = Pt(8)
            r.font.color.rgb = RGBColor(255, 255, 255)
        else:
            set_cell_background(cell, "F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
            r = p.add_run(val)
            r.font.name = 'Arial'
            r.font.size = Pt(8)
            r.font.color.rgb = BODY_BLACK

# ---------------------------------------------------------
# SECTION 6 : ANNEXES DOCUMENTAIRES
# ---------------------------------------------------------
add_native_heading("6. Annexe : Justifications des Gains & Sources Documentaires", 1, page_break=True)

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
print(f"Document Word/Google Docs conforme au format de référence généré avec succès : {docx_path}")
