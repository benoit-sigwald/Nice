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
        margin-bottom: 28px;
    }
    .grid {
        display: flex;
        gap: 24px;
    }
    .card {
        flex: 1;
        background: #F8FAFC;
        border: 1.5px solid #CBD5E1;
        border-radius: 8px;
        padding: 22px;
        text-align: center;
    }
    .card-title {
        font-size: 16px;
        color: #475569;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .card-val {
        font-size: 36px;
        font-weight: 700;
        color: #0F172A;
        margin-top: 10px;
    }
    .card-label {
        font-size: 15px;
        color: #64748B;
        margin-top: 6px;
    }
</style>
</head>
<body>
    <div class="title">Bilan Financier & Impact Métropolitain</div>
    <div class="subtitle">Modèle d'IA Frugale & Rigueur Budgétaire Certifiée (Données Annuelles Nettes)</div>
    
    <div class="grid">
        <div class="card">
            <div class="card-title">Audit Commande Publique</div>
            <div class="card-val" style="color: #059669;">+2,50 M€</div>
            <div class="card-label">Économies nettes / an certifiées</div>
        </div>
        <div class="card">
            <div class="card-title">Coût Budget IA Total</div>
            <div class="card-val" style="color: #DC2626;">0,50 M€</div>
            <div class="card-label">Co-financé par l'UE</div>
        </div>
        <div class="card">
            <div class="card-title">Bénéfice Net Public</div>
            <div class="card-val" style="color: #1E3A8A;">+2,00 M€</div>
            <div class="card-label">Réinjectés dans le service public</div>
        </div>
    </div>
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
    "--window-size=1440,360", 
    f"--screenshot={chart2_png}", chart2_html
], check=True)

print("Graphiques Ultra-Nets 4K (High-DPI) générés avec succès !")
