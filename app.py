"""
Registre des Timbres Fiscaux — Application Flask
Étude de commissaire de justice

SUPPRESSION / ANNULATION : Ces opérations ne sont pas disponibles dans
l'interface pour des raisons d'intégrité du registre. Procédure manuelle :
  - Suppression : éditer data/timbres_{année}.json + supprimer le PDF dans data/pdfs/
  - Annulation d'attribution : remettre statut="disponible", dossier=null,
    date_utilisation=null dans le fichier de l'année correspondante.
"""

import json
import re
import threading
import uuid
from datetime import date, datetime
from pathlib import Path

import openpyxl
from flask import (Flask, flash, redirect, render_template_string,
                   request, send_file, url_for)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pypdf import PdfReader, PdfWriter

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
MONTANT_TIMBRE = 50.0
SEUIL_ALERTE = 5

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
PDF_DIR = DATA_DIR / "pdfs"

_lock = threading.Lock()

# ---------------------------------------------------------------------------
# Initialisation Flask
# ---------------------------------------------------------------------------
app = Flask(__name__)
app.secret_key = "timbres-fiscaux-secret-key-2025"

# ---------------------------------------------------------------------------
# Couche données
# ---------------------------------------------------------------------------

def data_file(year: int) -> Path:
    return DATA_DIR / f"timbres_{year}.json"


def load_year(year: int) -> list:
    f = data_file(year)
    if not f.exists():
        return []
    with open(f, "r", encoding="utf-8") as fh:
        return json.load(fh)


def save_year(year: int, data: list):
    f = data_file(year)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with _lock:
        with open(f, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)


def annees_disponibles() -> list:
    return sorted(
        [int(p.stem.split("_")[1]) for p in DATA_DIR.glob("timbres_*.json")],
        reverse=True,
    )


def load_all() -> list:
    result = []
    for year in sorted(annees_disponibles()):
        result.extend(load_year(year))
    return result


def save_timbre(timbre: dict):
    year = int(timbre["date_achat"][:4])
    timbres = load_year(year)
    for i, t in enumerate(timbres):
        if t["id"] == timbre["id"]:
            timbres[i] = timbre
            break
    save_year(year, timbres)


# ---------------------------------------------------------------------------
# Extraction du numéro de timbre
# ---------------------------------------------------------------------------

PATTERNS = [
    re.compile(r"\b3[A-Z0-9]{2}[\s\-]?[0-9]{4}[\s\-]?[0-9]{4}[\s\-]?[0-9]{4}\b"),
    re.compile(r"\b[0-9]{4}[\s\-][0-9]{4}[\s\-][0-9]{4}[\s\-][0-9]{4}\b"),
    re.compile(r"\b[0-9]{13,20}\b"),
    re.compile(r"[Nn]°\s*([A-Z0-9\-]{8,})"),
    re.compile(r"[Rr][ée]f[ée]rence\s*:?\s*([A-Z0-9\-]{6,})"),
]


def extraire_numero(text: str) -> str | None:
    for pattern in PATTERNS:
        m = pattern.search(text)
        if m:
            return m.group(1) if m.lastindex else m.group(0)
    return None


# ---------------------------------------------------------------------------
# Lancement — banner terminal
# ---------------------------------------------------------------------------

def print_banner():
    annees = annees_disponibles()
    annees_str = ", ".join(str(a) for a in sorted(annees)) if annees else "aucune"
    print("=" * 60)
    print("  ⚖  Registre des Timbres Fiscaux")
    print("=" * 60)
    print(f"  Accès local  : http://localhost:5000")
    print(f"  Données      : {DATA_DIR / 'timbres_{année}.json'}")
    print(f"  PDFs         : {PDF_DIR}/")
    print(f"  Années       : {annees_str}")
    print("=" * 60)


# ---------------------------------------------------------------------------
# Template HTML base
# ---------------------------------------------------------------------------

BASE_HTML = """<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{ title }} — Timbres Fiscaux</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=EB+Garamond:ital,wght@0,400;0,600;0,700;1,400&display=swap" rel="stylesheet">
<style>
  :root {
    --marine:   #1a2744;
    --marine2:  #243156;
    --or:       #c9a84c;
    --creme:    #f5e6c8;
    --fond:     #f7f5f0;
    --fond-row: #faf9f6;
    --border:   #e8e3d8;
    --border2:  #d0c8b8;
    --vert-bg:  #d4edda;
    --vert-fg:  #155724;
    --rouge-bg: #f8d7da;
    --rouge-fg: #721c24;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'EB Garamond', Georgia, serif;
    background: var(--fond);
    color: #2c2c2c;
    min-height: 100vh;
  }
  /* NAV */
  nav {
    background: linear-gradient(135deg, var(--marine) 0%, var(--marine2) 100%);
    border-bottom: 3px solid var(--or);
    padding: 0 2rem;
    display: flex;
    align-items: center;
    gap: 2rem;
    height: 58px;
  }
  .nav-brand {
    font-size: 1.25rem;
    font-weight: 700;
    color: var(--creme);
    text-decoration: none;
    white-space: nowrap;
    letter-spacing: .03em;
  }
  .nav-brand span { color: var(--or); }
  .nav-links { display: flex; gap: 0; flex: 1; }
  .nav-links a {
    color: #c8d4e8;
    text-decoration: none;
    padding: .4rem 1.1rem;
    font-size: 1rem;
    border-bottom: 3px solid transparent;
    margin-bottom: -3px;
    transition: color .2s, border-color .2s;
  }
  .nav-links a:hover { color: var(--creme); }
  .nav-links a.active { color: var(--or); border-color: var(--or); }
  .nav-right { margin-left: auto; }
  .btn-excel {
    border: 1.5px solid var(--or);
    color: var(--or);
    background: transparent;
    padding: .35rem .9rem;
    border-radius: 6px;
    text-decoration: none;
    font-size: .92rem;
    font-family: inherit;
    cursor: pointer;
    transition: background .2s, color .2s;
  }
  .btn-excel:hover { background: var(--or); color: var(--marine); }
  /* MAIN */
  main { max-width: 1100px; margin: 0 auto; padding: 2rem 1.5rem; }
  h1 { font-size: 1.7rem; color: var(--marine); margin-bottom: 1.5rem; font-weight: 600; }
  h2 { font-size: 1.25rem; color: var(--marine); margin-bottom: 1rem; font-weight: 600; }
  /* ALERTS */
  .alert {
    padding: .85rem 1.2rem;
    border-radius: 8px;
    margin-bottom: 1.2rem;
    font-size: 1rem;
    border-left: 5px solid;
  }
  .alert-danger  { background: var(--rouge-bg); color: var(--rouge-fg); border-color: #e74c3c; }
  .alert-success { background: var(--vert-bg);  color: var(--vert-fg);  border-color: #27ae60; }
  .alert-error   { background: var(--rouge-bg); color: var(--rouge-fg); border-color: #e74c3c; }
  /* CARDS */
  .card {
    background: #fff;
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1.4rem 1.6rem;
    box-shadow: 0 2px 8px rgba(26,39,68,.07);
    margin-bottom: 1.4rem;
  }
  /* STATS */
  .stats-grid { display: grid; grid-template-columns: repeat(3,1fr); gap: 1.2rem; margin-bottom: 1.4rem; }
  .stat-card {
    background: #fff;
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1.3rem 1.5rem;
    box-shadow: 0 2px 8px rgba(26,39,68,.07);
    text-align: center;
  }
  .stat-card .label { font-size: .88rem; color: #666; text-transform: uppercase; letter-spacing: .06em; margin-bottom: .5rem; }
  .stat-card .value { font-size: 2.4rem; font-weight: 700; line-height: 1; }
  .stat-card.total .value  { color: var(--marine); }
  .stat-card.dispo .value  { color: var(--vert-fg); }
  .stat-card.utilise .value{ color: var(--rouge-fg); }
  /* WIDGETS */
  .widgets-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 1.2rem; margin-bottom: 1.4rem; }
  .widget {
    background: #fff;
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1.2rem 1.5rem;
    box-shadow: 0 2px 8px rgba(26,39,68,.07);
  }
  .widget .wlabel { font-size: .82rem; text-transform: uppercase; letter-spacing: .06em; color: #888; margin-bottom: .4rem; }
  .widget .wmain  { font-size: 1.1rem; font-weight: 600; color: var(--marine); }
  .widget .wsub   { font-size: .92rem; color: #555; margin-top: .2rem; }
  /* FORM */
  label { display: block; font-size: .95rem; color: #444; margin-bottom: .3rem; }
  input[type=text], input[type=date] {
    width: 100%; padding: .55rem .8rem;
    border: 1.5px solid var(--border2);
    border-radius: 7px;
    font-family: inherit;
    font-size: 1rem;
    background: var(--fond-row);
    transition: border-color .2s;
  }
  input[type=text]:focus, input[type=date]:focus {
    outline: none; border-color: var(--or);
  }
  .form-row { margin-bottom: 1rem; }
  .btn {
    background: var(--or);
    color: var(--marine);
    border: none;
    padding: .6rem 1.4rem;
    border-radius: 7px;
    font-family: inherit;
    font-size: 1rem;
    font-weight: 700;
    cursor: pointer;
    transition: filter .2s;
  }
  .btn:hover { filter: brightness(1.08); }
  /* DROPZONE */
  .dropzone {
    border: 2.5px dashed var(--or);
    border-radius: 10px;
    background: #fffdf7;
    padding: 2rem;
    text-align: center;
    cursor: pointer;
    transition: background .2s;
    margin-bottom: .5rem;
  }
  .dropzone:hover, .dropzone.dragover { background: #fff8e8; }
  .dropzone .dz-icon { font-size: 2.5rem; margin-bottom: .5rem; }
  .dropzone .dz-text { color: #666; font-size: .95rem; }
  .dropzone .dz-price { color: var(--or); font-weight: 600; margin-top: .4rem; font-size: .9rem; }
  #pdf-input { display: none; }
  #file-name  { font-size: .9rem; color: var(--marine); margin-top: .4rem; font-style: italic; }
  /* TIMBRE FICHE */
  .timbre-fiche {
    display: flex;
    align-items: center;
    background: #fff;
    border: 1px solid var(--border);
    border-radius: 10px;
    box-shadow: 0 2px 8px rgba(26,39,68,.07);
    margin-bottom: 1.2rem;
    overflow: hidden;
  }
  .timbre-fiche .tf-cell {
    padding: 1.1rem 1.5rem;
    border-right: 1px solid var(--border);
    flex: 1;
  }
  .timbre-fiche .tf-cell:last-child { border-right: none; }
  .timbre-fiche .tf-label { font-size: .78rem; text-transform: uppercase; letter-spacing: .06em; color: #888; margin-bottom: .25rem; }
  .timbre-fiche .tf-value { font-size: 1.08rem; font-weight: 600; color: var(--marine); }
  .timbre-fiche .tf-mono  { font-family: 'Courier New', monospace; font-size: 1.05rem; font-weight: 700; color: var(--marine); letter-spacing: .04em; }
  /* TABLE */
  table { width: 100%; border-collapse: collapse; font-size: .97rem; }
  thead tr { background: var(--marine); color: var(--creme); }
  thead th {
    padding: .75rem 1rem;
    text-align: left;
    font-weight: 600;
    border-bottom: 2px solid var(--or);
    letter-spacing: .03em;
  }
  tbody tr { border-bottom: 1px solid var(--border); }
  tbody tr:nth-child(odd)  { background: #fff; }
  tbody tr:nth-child(even) { background: var(--fond-row); }
  tbody td { padding: .65rem 1rem; }
  .badge {
    display: inline-block;
    padding: .2rem .7rem;
    border-radius: 50px;
    font-size: .82rem;
    font-weight: 600;
  }
  .badge-dispo  { background: var(--vert-bg);  color: var(--vert-fg); }
  .badge-utilise{ background: var(--rouge-bg); color: var(--rouge-fg); }
  .mono { font-family: 'Courier New', monospace; font-size: .95rem; }
  /* LOT HEADER */
  .lot-header {
    background: var(--marine);
    color: var(--creme);
    padding: .75rem 1.1rem;
    border-radius: 8px 8px 0 0;
    border-bottom: 2px solid var(--or);
    font-size: .97rem;
    display: flex;
    align-items: center;
    gap: 1.2rem;
  }
  .lot-block { margin-bottom: 1.6rem; border-radius: 8px; overflow: hidden; border: 1px solid var(--border); }
  .lot-block table { border-radius: 0; }
  /* TOOLBAR */
  .toolbar { display: flex; gap: 1rem; align-items: center; margin-bottom: 1.2rem; flex-wrap: wrap; }
  .toolbar input[type=text] { max-width: 300px; }
  select {
    padding: .5rem .8rem;
    border: 1.5px solid var(--border2);
    border-radius: 7px;
    font-family: inherit;
    font-size: 1rem;
    background: var(--fond-row);
    cursor: pointer;
  }
  select:focus { outline: none; border-color: var(--or); }
  /* SPINNER PAGE */
  .spinner-wrap { display: flex; flex-direction: column; align-items: center; justify-content: center; min-height: 60vh; gap: 1.5rem; }
  .spinner {
    width: 60px; height: 60px;
    border: 5px solid #e8e3d8;
    border-top-color: var(--or);
    border-radius: 50%;
    animation: spin .8s linear infinite;
  }
  @keyframes spin { to { transform: rotate(360deg); } }
  .spinner-text { font-size: 1.2rem; color: var(--marine); }
  /* MISC */
  .subtitle { color: #555; font-size: 1rem; margin-bottom: 1.2rem; }
  .note { font-size: .9rem; color: #777; margin-top: .8rem; font-style: italic; }
  .empty { text-align: center; padding: 2.5rem; color: #888; font-size: 1.1rem; }
  .hidden { display: none !important; }
  @media (max-width: 700px) {
    .stats-grid, .widgets-grid { grid-template-columns: 1fr; }
    .timbre-fiche { flex-direction: column; }
    .timbre-fiche .tf-cell { border-right: none; border-bottom: 1px solid var(--border); width: 100%; }
  }
</style>
</head>
<body>
<nav>
  <a href="/" class="nav-brand">⚖ Timbres <span>Fiscaux</span></a>
  <div class="nav-links">
    <a href="/" class="{{ 'active' if active=='dashboard' else '' }}">Tableau de bord</a>
    <a href="/disponibles" class="{{ 'active' if active=='disponibles' else '' }}">Disponibles</a>
    <a href="/historique" class="{{ 'active' if active=='historique' else '' }}">Historique</a>
  </div>
  <div class="nav-right">
    <a href="/export-excel" class="btn-excel">⬇ Export Excel</a>
  </div>
</nav>
<main>
{% with messages = get_flashed_messages(with_categories=true) %}
  {% for cat, msg in messages %}
    <div class="alert alert-{{ cat }}">{{ msg }}</div>
  {% endfor %}
{% endwith %}
{{ content }}
</main>
</body>
</html>"""


def render_page(title: str, active: str, content: str) -> str:
    return render_template_string(
        BASE_HTML, title=title, active=active, content=content
    )


# ---------------------------------------------------------------------------
# PAGE 1 — TABLEAU DE BORD
# ---------------------------------------------------------------------------

@app.route("/")
def dashboard():
    timbres = load_all()
    disponibles = [t for t in timbres if t["statut"] == "disponible"]
    utilises    = [t for t in timbres if t["statut"] == "utilisé"]

    nb_dispo  = len(disponibles)
    nb_utilise = len(utilises)
    nb_total   = len(timbres)

    # Dernier import : lot avec date_achat la plus récente
    dernier_import = None
    if timbres:
        date_max = max(t["date_achat"] for t in timbres)
        lot_recents = [t for t in timbres if t["date_achat"] == date_max]
        dernier_import = {
            "date": date_max,
            "nb": len(lot_recents),
            "montant": MONTANT_TIMBRE,
        }

    # Dernière attribution
    derniere_attrib = None
    if utilises:
        u = max(utilises, key=lambda t: t.get("date_utilisation") or "")
        derniere_attrib = u

    today = date.today().isoformat()

    alert_html = ""
    if nb_dispo <= SEUIL_ALERTE:
        alert_html = f"""<div class="alert alert-danger">
          ⚠ Stock bas — {nb_dispo} timbre(s) disponible(s) · Pensez à commander un nouveau lot
        </div>"""

    import_widget = "Aucun import"
    if dernier_import:
        import_widget = f"""
          <div class="wmain">{dernier_import['date']} · {dernier_import['nb']} timbre(s)</div>
          <div class="wsub">{dernier_import['montant']:.2f} € / timbre</div>"""

    attrib_widget = "Aucune attribution"
    if derniere_attrib:
        attrib_widget = f"""
          <div class="wmain mono">{derniere_attrib['numero']}</div>
          <div class="wsub">{derniere_attrib.get('dossier','') or ''}</div>
          <div class="wsub">{derniere_attrib.get('date_utilisation','') or ''}</div>"""

    content = f"""
{alert_html}
<h1>Tableau de bord</h1>

<div class="stats-grid">
  <div class="stat-card total">
    <div class="label">Total</div>
    <div class="value">{nb_total}</div>
  </div>
  <div class="stat-card dispo">
    <div class="label">Disponibles</div>
    <div class="value">{nb_dispo}</div>
  </div>
  <div class="stat-card utilise">
    <div class="label">Utilisés</div>
    <div class="value">{nb_utilise}</div>
  </div>
</div>

<div class="widgets-grid">
  <div class="widget">
    <div class="wlabel">Dernier import</div>
    {import_widget if dernier_import else '<div class="wmain" style="color:#999">Aucun import</div>'}
  </div>
  <div class="widget">
    <div class="wlabel">Dernière attribution</div>
    {attrib_widget if derniere_attrib else '<div class="wmain" style="color:#999">Aucune attribution</div>'}
  </div>
</div>

<div class="card">
  <h2>Importer un lot de timbres</h2>
  <form method="post" action="/import" enctype="multipart/form-data">
    <div class="form-row">
      <label for="date_achat">Date d'achat</label>
      <input type="date" id="date_achat" name="date_achat" value="{today}" required style="max-width:220px">
    </div>
    <div class="form-row">
      <label>Fichier PDF du lot</label>
      <div class="dropzone" id="dropzone" onclick="document.getElementById('pdf-input').click()">
        <div class="dz-icon">📄</div>
        <div class="dz-text">Glissez le PDF ici ou <strong>cliquez pour parcourir</strong></div>
        <div class="dz-price">50,00 € / timbre</div>
      </div>
      <input type="file" id="pdf-input" name="pdf" accept="application/pdf" required>
      <div id="file-name"></div>
    </div>
    <button type="submit" class="btn">Importer le lot</button>
  </form>
</div>

<script>
const dz = document.getElementById('dropzone');
const inp = document.getElementById('pdf-input');
const fn  = document.getElementById('file-name');
inp.addEventListener('change', () => {{
  fn.textContent = inp.files[0] ? '📎 ' + inp.files[0].name : '';
}});
dz.addEventListener('dragover', e => {{ e.preventDefault(); dz.classList.add('dragover'); }});
dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
dz.addEventListener('drop', e => {{
  e.preventDefault(); dz.classList.remove('dragover');
  inp.files = e.dataTransfer.files;
  fn.textContent = inp.files[0] ? '📎 ' + inp.files[0].name : '';
}});
</script>
"""
    return render_page("Tableau de bord", "dashboard", content)


# ---------------------------------------------------------------------------
# IMPORT LOT
# ---------------------------------------------------------------------------

@app.route("/import", methods=["POST"])
def import_lot():
    pdf_file = request.files.get("pdf")
    date_achat = request.form.get("date_achat", "").strip()

    if not pdf_file or not date_achat:
        flash("Fichier PDF et date d'achat requis.", "error")
        return redirect(url_for("dashboard"))

    try:
        year = int(date_achat[:4])
        reader = PdfReader(pdf_file.stream)
        nb_pages = len(reader.pages)

        PDF_DIR.mkdir(parents=True, exist_ok=True)
        DATA_DIR.mkdir(parents=True, exist_ok=True)

        # Construire le numéro de séquence basé sur les timbres existants
        existing = load_year(year)
        start_idx = len(existing) + 1

        nouveaux = []
        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            numero = extraire_numero(text)
            if not numero:
                numero = f"TIMBRE-{date_achat}-{start_idx + i:03d}"

            pdf_uuid = uuid.uuid4().hex + ".pdf"
            writer = PdfWriter()
            writer.add_page(page)
            pdf_path = PDF_DIR / pdf_uuid
            with open(pdf_path, "wb") as fout:
                writer.write(fout)

            timbre = {
                "id": str(uuid.uuid4()),
                "numero": numero,
                "date_achat": date_achat,
                "montant": MONTANT_TIMBRE,
                "statut": "disponible",
                "pdf": pdf_uuid,
                "dossier": None,
                "date_utilisation": None,
            }
            nouveaux.append(timbre)

        existing.extend(nouveaux)
        save_year(year, existing)
        flash(f"{nb_pages} timbre(s) importé(s) avec succès ({year}).", "success")

    except Exception as exc:
        flash(f"Erreur lors de l'import : {exc}", "error")

    return redirect(url_for("dashboard"))


# ---------------------------------------------------------------------------
# PAGE 2 — DISPONIBLES
# ---------------------------------------------------------------------------

@app.route("/disponibles")
def disponibles():
    timbres = load_all()
    dispos = sorted(
        [t for t in timbres if t["statut"] == "disponible"],
        key=lambda t: (t["date_achat"], t["id"]),
    )
    nb_dispo = len(dispos)
    prochain = dispos[0] if dispos else None
    en_attente = nb_dispo - 1 if nb_dispo > 1 else 0

    if prochain:
        position = 1
        total_lot = sum(
            1 for t in timbres
            if t["date_achat"] == prochain["date_achat"]
        )
        fiche = f"""
<div class="timbre-fiche">
  <div class="tf-cell">
    <div class="tf-label">N° Timbre</div>
    <div class="tf-mono">{prochain['numero']}</div>
  </div>
  <div class="tf-cell">
    <div class="tf-label">Montant</div>
    <div class="tf-value">{prochain['montant']:.2f} €</div>
  </div>
  <div class="tf-cell">
    <div class="tf-label">Date d'achat</div>
    <div class="tf-value">{prochain['date_achat']}</div>
  </div>
  <div class="tf-cell">
    <div class="tf-label">File</div>
    <div class="tf-value">{position} / {nb_dispo}</div>
  </div>
</div>

<div class="card">
  <h2>Attribuer ce timbre</h2>
  <form method="post" action="/utiliser">
    <input type="hidden" name="timbre_id" value="{prochain['id']}">
    <div class="form-row">
      <label for="dossier">Référence dossier / affaire</label>
      <input type="text" id="dossier" name="dossier" placeholder="ex. 2025-042 · Dupont / SCI Les Pins"
             autofocus required>
    </div>
    <button type="submit" class="btn">Attribuer ce timbre</button>
  </form>
</div>
{"<p class='note'>📋 " + str(en_attente) + " timbre(s) en attente derrière celui-ci.</p>" if en_attente > 0 else ""}
"""
        subtitle = f"{nb_dispo} timbre(s) en stock · le suivant sera servi après attribution"
    else:
        fiche = "<div class='empty'>📭 Aucun timbre disponible. Importez un lot depuis le tableau de bord.</div>"
        subtitle = "Aucun timbre en stock"

    content = f"""
<h1>Timbres disponibles</h1>
<p class="subtitle">{subtitle}</p>
{fiche}
"""
    return render_page("Disponibles", "disponibles", content)


# ---------------------------------------------------------------------------
# ATTRIBUTION
# ---------------------------------------------------------------------------

@app.route("/utiliser", methods=["POST"])
def utiliser():
    timbre_id = request.form.get("timbre_id", "").strip()
    dossier   = request.form.get("dossier", "").strip()

    if not timbre_id or not dossier:
        flash("Données manquantes.", "error")
        return redirect(url_for("disponibles"))

    timbres = load_all()
    timbre  = next((t for t in timbres if t["id"] == timbre_id), None)

    if not timbre:
        flash("Timbre introuvable.", "error")
        return redirect(url_for("disponibles"))

    if timbre["statut"] != "disponible":
        flash("Ce timbre n'est plus disponible.", "error")
        return redirect(url_for("disponibles"))

    timbre["statut"] = "utilisé"
    timbre["dossier"] = dossier
    timbre["date_utilisation"] = date.today().isoformat()
    save_timbre(timbre)

    return redirect(url_for("telecharger", timbre_id=timbre_id))


@app.route("/attribution/telecharger/<timbre_id>")
def telecharger(timbre_id: str):
    timbres = load_all()
    timbre  = next((t for t in timbres if t["id"] == timbre_id), None)

    if not timbre:
        flash("Timbre introuvable.", "error")
        return redirect(url_for("disponibles"))

    numero  = timbre["numero"]
    dossier = timbre.get("dossier") or ""
    pdf_url = url_for("serve_pdf", filename=timbre["pdf"])

    content = f"""
<div class="spinner-wrap">
  <div class="spinner"></div>
  <div class="spinner-text">Téléchargement du timbre {numero} en cours…</div>
  <div style="color:#888;font-size:.9rem">{dossier}</div>
</div>
<a id="dl-link" href="{pdf_url}" download="{numero}.pdf" style="display:none"></a>
<script>
  window.onload = function() {{
    document.getElementById('dl-link').click();
    setTimeout(function() {{
      window.location.href = "{url_for('disponibles')}";
    }}, 1500);
  }};
</script>
"""
    # Préparer le flash pour la page de destination
    flash(f"Timbre {numero} attribué au dossier « {dossier} » et téléchargé.", "success")
    return render_page("Téléchargement", "disponibles", content)


# ---------------------------------------------------------------------------
# SERVE PDF (protégé)
# ---------------------------------------------------------------------------

@app.route("/pdfs/<filename>")
def serve_pdf(filename: str):
    timbres = load_all()
    timbre  = next((t for t in timbres if t.get("pdf") == filename), None)
    if not timbre:
        return "Fichier introuvable.", 404
    if timbre["statut"] == "disponible":
        return "Accès refusé : ce timbre n'a pas encore été attribué.", 403
    pdf_path = PDF_DIR / filename
    if not pdf_path.exists():
        return "Fichier PDF manquant.", 404
    return send_file(pdf_path, mimetype="application/pdf")


# ---------------------------------------------------------------------------
# PAGE 3 — HISTORIQUE
# ---------------------------------------------------------------------------

@app.route("/historique")
def historique():
    annee_param = request.args.get("annee", "toutes")
    annees = annees_disponibles()

    if annee_param == "toutes":
        timbres = load_all()
    else:
        try:
            timbres = load_year(int(annee_param))
        except ValueError:
            timbres = load_all()
            annee_param = "toutes"

    utilises = [t for t in timbres if t["statut"] == "utilisé"]

    # Regrouper par lot (date_achat, montant)
    lots: dict[tuple, list] = {}
    for t in utilises:
        key = (t["date_achat"], t["montant"])
        lots.setdefault(key, []).append(t)

    # Trier chaque lot par date_utilisation décroissante
    for key in lots:
        lots[key].sort(key=lambda t: t.get("date_utilisation") or "", reverse=True)

    # Trier les lots par date_achat décroissante
    lots_sorted = sorted(lots.items(), key=lambda kv: kv[0][0], reverse=True)

    # Totaux pour l'en-tête de lot : nb dans le fichier année toutes statuts
    def total_lot_all(date_a, montant):
        yr = int(date_a[:4])
        all_yr = load_year(yr)
        return sum(1 for t in all_yr if t["date_achat"] == date_a and t["montant"] == montant)

    nb_utilises_total = len(utilises)
    montant_total = nb_utilises_total * MONTANT_TIMBRE
    nb_lots = len(lots)

    # Options années
    opts = '<option value="toutes"' + (' selected' if annee_param == "toutes" else "") + '>Toutes les années</option>\n'
    for a in annees:
        sel = ' selected' if str(a) == annee_param else ""
        opts += f'<option value="{a}"{sel}>{a}</option>\n'

    # Construire les blocs de lots
    lots_html = ""
    for (date_a, montant), items in lots_sorted:
        nb_utilises_lot = len(items)
        total_lot = total_lot_all(date_a, montant)
        montant_lot = nb_utilises_lot * montant
        rows = ""
        for t in items:
            pdf_btn = f'<a href="/pdfs/{t["pdf"]}" target="_blank" class="btn" style="padding:.25rem .7rem;font-size:.82rem">Voir</a>' if t.get("pdf") else "—"
            rows += f"""<tr class="hist-row">
              <td class="mono">{t['numero']}</td>
              <td>{t.get('dossier') or '—'}</td>
              <td>{t.get('date_utilisation') or '—'}</td>
              <td>{pdf_btn}</td>
            </tr>"""
        lots_html += f"""
<div class="lot-block" data-lot>
  <div class="lot-header">
    📦 Lot du {date_a} &nbsp;·&nbsp; {montant:.2f} € / timbre
    &nbsp;&nbsp;|&nbsp;&nbsp; {nb_utilises_lot} utilisé(s) / {total_lot} dans ce lot
    &nbsp;&nbsp;|&nbsp;&nbsp; {montant_lot:,.2f} €
  </div>
  <table>
    <thead><tr>
      <th>N° Timbre</th><th>Dossier / Affaire</th><th>Date utilisation</th><th>PDF</th>
    </tr></thead>
    <tbody>{rows}</tbody>
  </table>
</div>"""

    if not lots_html:
        annee_txt = f" pour {annee_param}" if annee_param != "toutes" else ""
        lots_html = f"<div class='empty'>Aucun timbre utilisé{annee_txt}.</div>"

    subtitle = f"{nb_utilises_total} timbre(s) utilisé(s) sur {nb_lots} lot(s) — total : {montant_total:,.2f} €"

    content = f"""
<h1>Historique des attributions</h1>
<p class="subtitle">{subtitle}</p>

<div class="toolbar">
  <input type="text" id="search" placeholder="Rechercher…" oninput="filtrer()">
  <select id="annee-sel" onchange="window.location.href='/historique?annee='+this.value">
    {opts}
  </select>
</div>

{lots_html}

<script>
function filtrer() {{
  const q = document.getElementById('search').value.toLowerCase();
  document.querySelectorAll('[data-lot]').forEach(function(bloc) {{
    let visible = 0;
    bloc.querySelectorAll('tr.hist-row').forEach(function(row) {{
      const match = row.textContent.toLowerCase().includes(q);
      row.style.display = match ? '' : 'none';
      if (match) visible++;
    }});
    bloc.style.display = visible > 0 ? '' : 'none';
  }});
}}
</script>
"""
    return render_page("Historique", "historique", content)


# ---------------------------------------------------------------------------
# EXPORT EXCEL
# ---------------------------------------------------------------------------

@app.route("/export-excel")
def export_excel():
    annees = annees_disponibles()
    if not annees:
        flash("Aucune donnée à exporter.", "error")
        return redirect(url_for("dashboard"))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    marine   = "1a2744"
    creme    = "f5e6c8"
    or_hex   = "c9a84c"
    vert_bg  = "d4edda"
    vert_fg  = "155724"
    rouge_bg = "f8d7da"
    rouge_fg = "721c24"

    side_or = Side(style="thin", color=or_hex)
    border_or = Border(left=side_or, right=side_or, top=side_or, bottom=side_or)

    for year in sorted(annees, reverse=True):
        timbres = load_year(year)
        timbres.sort(key=lambda t: (t["statut"], t["date_achat"]))

        ws = wb.create_sheet(title=str(year))

        headers = ["N° Timbre", "Date achat", "Montant (€)", "Statut", "Dossier", "Date utilisation"]
        ws.append(headers)
        for col, _ in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col)
            cell.font = Font(bold=True, color=creme, name="Calibri")
            cell.fill = PatternFill("solid", fgColor=marine)
            cell.border = border_or
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for t in timbres:
            statut_label = "Disponible" if t["statut"] == "disponible" else "Utilisé"
            row = [
                t["numero"],
                t["date_achat"],
                t["montant"],
                statut_label,
                t.get("dossier") or "",
                t.get("date_utilisation") or "",
            ]
            ws.append(row)
            r = ws.max_row
            # Statut colonne D
            stat_cell = ws.cell(row=r, column=4)
            if t["statut"] == "disponible":
                stat_cell.fill = PatternFill("solid", fgColor=vert_bg)
                stat_cell.font = Font(color=vert_fg, name="Calibri")
            else:
                stat_cell.fill = PatternFill("solid", fgColor=rouge_bg)
                stat_cell.font = Font(color=rouge_fg, name="Calibri")
            stat_cell.alignment = Alignment(horizontal="center")

        col_widths = [22, 14, 14, 14, 45, 18]
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

        ws.row_dimensions[1].height = 22

    today_str = date.today().isoformat()
    filename  = f"timbres-fiscaux-{today_str}.xlsx"
    tmp_path  = DATA_DIR / filename
    wb.save(tmp_path)

    return send_file(
        tmp_path,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    PDF_DIR.mkdir(parents=True, exist_ok=True)
    print_banner()
    app.run(host="0.0.0.0", port=5000, debug=False)
