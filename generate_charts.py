import os
import subprocess

chart1_html = r"g:\My Drive\Dev\Einstein\temp_chart1.html"
chart2_html = r"g:\My Drive\Dev\Einstein\temp_chart2.html"
chart1_png = r"g:\My Drive\Dev\Einstein\chart1_poids_national.png"
chart2_png = r"g:\My Drive\Dev\Einstein\chart2_roi_gains.png"

html1_content = """<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    body {
        margin: 0;
        padding: 30px;
        background: #FFFFFF;
        color: #0F172A;
        font-family: 'Inter', sans-serif;
        width: 1400px;
        box-sizing: border-box;
        border: 2px solid #E2E8F0;
        border-radius: 8px;
    }
    .title {
        font-size: 26px;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 4px;
    }
    .subtitle {
        font-size: 18px;
        color: #64748B;
        margin-bottom: 28px;
    }
    .bar-container {
        margin-bottom: 22px;
    }
    .bar-header {
        display: flex;
        justify-content: space-between;
        font-size: 20px;
        font-weight: 600;
        margin-bottom: 8px;
    }
    .bar-label { color: #334155; }
    .bar-val { color: #1E3A8A; font-weight: 700; }
    .bar-bg {
        background: #F1F5F9;
        border-radius: 6px;
        height: 28px;
        overflow: hidden;
    }
    .bar-fill {
        height: 100%;
        border-radius: 6px;
        background: #1E3A8A;
    }
</style>
</head>
<body>
    <div class="title">Poids du Pôle Nice - Sophia Antipolis - Monaco vs Total National Français</div>
    <div class="subtitle">Part relative dans les capacités nationales de Recherche & IA en France (Sources INSEE / 3IA 2026)</div>
    
    <div class="bar-container">
        <div class="bar-header">
            <span class="bar-label">Recherche Académique IA d'Élite (Réseau National des 4 Instituts 3IA)</span>
            <span class="bar-val">25 % de la France</span>
        </div>
        <div class="bar-bg"><div class="bar-fill" style="width: 25%;"></div></div>
    </div>
    
    <div class="bar-container">
        <div class="bar-header">
            <span class="bar-label">Chercheurs & Ingénieurs R&D Numérique (hors Île-de-France)</span>
            <span class="bar-val">12 % du hors-IDF</span>
        </div>
        <div class="bar-bg"><div class="bar-fill" style="width: 12%;"></div></div>
    </div>
    
    <div class="bar-container">
        <div class="bar-header">
            <span class="bar-label">Dépense de R&D Privée Logicielle / Spatiale (Amadeus & Thales Alenia Space)</span>
            <span class="bar-val">6 % du total France</span>
        </div>
        <div class="bar-bg"><div class="bar-fill" style="width: 6%;"></div></div>
    </div>
</body>
</html>
"""

html2_content = """<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    body {
        margin: 0;
        padding: 30px;
        background: #FFFFFF;
        color: #0F172A;
        font-family: 'Inter', sans-serif;
        width: 1400px;
        box-sizing: border-box;
        border: 2px solid #E2E8F0;
        border-radius: 8px;
    }
    .title {
        font-size: 26px;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 4px;
    }
    .subtitle {
        font-size: 18px;
        color: #64748B;
        margin-bottom: 24px;
    }
    .grid {
        display: flex;
        gap: 20px;
        margin-bottom: 24px;
    }
    .card {
        flex: 1;
        background: #F8FAFC;
        border: 1.5px solid #CBD5E1;
        border-radius: 8px;
        padding: 18px;
        text-align: center;
    }
    .card-title {
        font-size: 15px;
        color: #475569;
        font-weight: 600;
        text-transform: uppercase;
    }
    .card-val {
        font-size: 32px;
        font-weight: 700;
        color: #0F172A;
        margin-top: 8px;
    }
    .card-label {
        font-size: 14px;
        color: #64748B;
        margin-top: 4px;
    }
    
    .alloc-title {
        font-size: 18px;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 12px;
        border-bottom: 2px solid #E2E8F0;
        padding-bottom: 6px;
    }
    
    .alloc-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 15px;
    }
    .alloc-table th {
        background: #1E293B;
        color: #FFFFFF;
        text-align: left;
        padding: 8px 12px;
    }
    .alloc-table td {
        padding: 8px 12px;
        border-bottom: 1px solid #E2E8F0;
    }
    .alloc-table tr:nth-child(even) td {
        background: #F8FAFC;
    }
</style>
</head>
<body>
    <div class="title">Bilan Financier & Ventilation Budgétaire Métropolitaine</div>
    <div class="subtitle">Découpage de l'Allocation des Dépenses (0,50 M€ / an) & Rentrée Nette Certifiée</div>
    
    <div class="grid">
        <div class="card">
            <div class="card-title">Audit Commande Publique</div>
            <div class="card-val" style="color: #059669;">+2,50 M€ / an</div>
            <div class="card-label">Économies nettes certifiées</div>
        </div>
        <div class="card">
            <div class="card-title">Coût Budget IA Total</div>
            <div class="card-val" style="color: #DC2626;">0,50 M€ / an</div>
            <div class="card-label">Reste à charge Métropole: 250 k€ (50% UE)</div>
        </div>
        <div class="card">
            <div class="card-title">Bénéfice Net Public</div>
            <div class="card-val" style="color: #1E3A8A;">+2,00 M€ / an</div>
            <div class="card-label">Réinjectés dans le service public</div>
        </div>
    </div>
    
    <div class="alloc-title">Détail de l'Allocation du Budget IA (500 k€ / an)</div>
    <table class="alloc-table">
        <thead>
            <tr>
                <th>Poste d'Allocation Budgétaire</th>
                <th>Montant Annuel</th>
                <th>Part (%)</th>
                <th>Description / Destination Opérationnelle</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td><strong>1. Licences Algorithmiques & Modèles IA Souverains</strong></td>
                <td><strong>150 k€</strong></td>
                <td>30 %</td>
                <td>Inférence VSA CSU, modèles locaux et API sécurisées sans fuite de données.</td>
            </tr>
            <tr>
                <td><strong>2. Cloud Souverain & Interconnexion Monaco Cloud</strong></td>
                <td><strong>120 k€</strong></td>
                <td>24 %</td>
                <td>Instances GPU d'inférence certifiées AMSN et liaison fibre noire dédiée.</td>
            </tr>
            <tr>
                <td><strong>3. Direction de Projet (AMO IA) & Ingénierie</strong></td>
                <td><strong>130 k€</strong></td>
                <td>26 %</td>
                <td>Direction de projet (Benoît SIGWALD), suivi des marchés et recettes.</td>
            </tr>
            <tr>
                <td><strong>4. Audit AI Act, Cybersécurité NIS 2 & Éthique</strong></td>
                <td><strong>60 k€</strong></td>
                <td>12 %</td>
                <td>Tests d'intrusion, audits de conformité AI Act et secrétariat du Comité d'Éthique.</td>
            </tr>
            <tr>
                <td><strong>5. Démocratie Participative & IA Einstein par Quartier</strong></td>
                <td><strong>40 k€</strong></td>
                <td>8 %</td>
                <td>Plateforme de concertation citoyenne augmentée et restitution par quartier.</td>
            </tr>
        </tbody>
    </table>
</body>
</html>
"""

with open(chart1_html, 'w', encoding='utf-8') as f:
    f.write(html1_content)

with open(chart2_html, 'w', encoding='utf-8') as f:
    f.write(html2_content)

chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

subprocess.run([
    chrome_path, "--headless", "--disable-gpu", 
    "--force-device-scale-factor=2",
    "--window-size=1440,460", 
    f"--screenshot={chart1_png}", chart1_html
], check=True)

subprocess.run([
    chrome_path, "--headless", "--disable-gpu", 
    "--force-device-scale-factor=2",
    "--window-size=1440,650", 
    f"--screenshot={chart2_png}", chart2_html
], check=True)

print("Graphique de bilan financier avec découpage budgétaire généré avec succès !")
