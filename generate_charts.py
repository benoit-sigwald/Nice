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
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&family=Playfair+Display:wght@700&display=swap');
    body {
        margin: 0;
        padding: 24px;
        background: #0F172A;
        color: #F8FAFC;
        font-family: 'Inter', sans-serif;
        width: 750px;
        box-sizing: border-box;
    }
    .title {
        font-family: 'Playfair Display', serif;
        font-size: 20px;
        color: #D8B98A;
        margin-bottom: 4px;
    }
    .subtitle {
        font-size: 13px;
        color: #94A3B8;
        margin-bottom: 24px;
    }
    .bar-container {
        margin-bottom: 18px;
    }
    .bar-header {
        display: flex;
        justify-content: space-between;
        font-size: 14px;
        font-weight: 600;
        margin-bottom: 6px;
    }
    .bar-label { color: #E2E8F0; }
    .bar-val { color: #38BDF8; font-weight: 700; }
    .bar-bg {
        background: rgba(255, 255, 255, 0.08);
        border-radius: 8px;
        height: 22px;
        overflow: hidden;
        position: relative;
    }
    .bar-fill {
        height: 100%;
        border-radius: 8px;
        background: linear-gradient(90deg, #0284C7 0%, #D8B98A 100%);
    }
</style>
</head>
<body>
    <div class="title">Poids du Pôle Nice - Sophia Antipolis - Monaco</div>
    <div class="subtitle">Part relative dans les capacités nationales de Recherche & IA en France</div>
    
    <div class="bar-container">
        <div class="bar-header">
            <span class="bar-label">Recherche Académique IA d'Élite (Réseau National 3IA)</span>
            <span class="bar-val">25 % de la France</span>
        </div>
        <div class="bar-bg"><div class="bar-fill" style="width: 25%;"></div></div>
    </div>
    
    <div class="bar-container">
        <div class="bar-header">
            <span class="bar-label">Chercheurs & Ingénieurs R&D Numérique (hors Paris)</span>
            <span class="bar-val">12 % du hors-IDF</span>
        </div>
        <div class="bar-bg"><div class="bar-fill" style="width: 12%;"></div></div>
    </div>
    
    <div class="bar-container">
        <div class="bar-header">
            <span class="bar-label">Dépense de R&D Privée Logicielle/Spatiale (Amadeus/Thales)</span>
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
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&family=Playfair+Display:wght@700&display=swap');
    body {
        margin: 0;
        padding: 24px;
        background: #0F172A;
        color: #F8FAFC;
        font-family: 'Inter', sans-serif;
        width: 750px;
        box-sizing: border-box;
    }
    .title {
        font-family: 'Playfair Display', serif;
        font-size: 20px;
        color: #D8B98A;
        margin-bottom: 4px;
    }
    .subtitle {
        font-size: 13px;
        color: #94A3B8;
        margin-bottom: 24px;
    }
    .grid {
        display: flex;
        gap: 16px;
    }
    .card {
        flex: 1;
        background: rgba(30, 41, 59, 0.8);
        border: 1px solid rgba(216, 185, 138, 0.3);
        border-radius: 12px;
        padding: 16px;
        text-align: center;
    }
    .card-val {
        font-size: 24px;
        font-weight: 700;
        color: #34D399;
        margin-top: 8px;
    }
    .card-label {
        font-size: 12px;
        color: #CBD5E1;
        margin-top: 4px;
    }
</style>
</head>
<body>
    <div class="title">Bilan Financier & Impact Métropolitain</div>
    <div class="subtitle">Modèle d'IA Frugale & Rigueur Budgétaire Certifiée (Chiffres Annuels)</div>
    
    <div class="grid">
        <div class="card">
            <div style="font-size: 13px; color: #D8B98A; font-weight: 600;">Audit Commande Publique</div>
            <div class="card-val">+2,50 M€</div>
            <div class="card-label">Économies nettes / an certifiées</div>
        </div>
        <div class="card">
            <div style="font-size: 13px; color: #38BDF8; font-weight: 600;">Coût Budget IA Total</div>
            <div class="card-val" style="color: #F43F5E;">0,50 M€</div>
            <div class="card-label">Co-financé par l'UE</div>
        </div>
        <div class="card">
            <div style="font-size: 13px; color: #34D399; font-weight: 600;">Bénéfice Net Public</div>
            <div class="card-val">+2,00 M€</div>
            <div class="card-label">Réinjectés dans les services</div>
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

subprocess.run([chrome_path, "--headless", "--disable-gpu", "--window-size=800,260", f"--screenshot={chart1_png}", chart1_html], check=True)
subprocess.run([chrome_path, "--headless", "--disable-gpu", "--window-size=800,200", f"--screenshot={chart2_png}", chart2_html], check=True)

print("Graphiques haute définition générés avec succès : chart1_poids_national.png et chart2_roi_gains.png")
