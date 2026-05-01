#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FEC Audit Pro v5.0 — Interface Streamlit
Thème universel : compatible mode clair ET sombre.
"""

import streamlit as st
import tempfile, os, sys, warnings, json
warnings.filterwarnings('ignore')
from collections import Counter

# ── Config page ───────────────────────────────────────────────
st.set_page_config(
    page_title="FEC Audit Pro v5",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Import moteur métier ───────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from fec_audit_pro_v5 import (
    charger_fec, detecter_secteur,
    analyser_equilibre, analyser_benford, analyser_montants_ronds,
    analyser_weekend, analyser_soldes, analyser_doublons,
    analyser_inversees, analyser_fin_exercice, analyser_concentration,
    analyser_saisonnalite, analyser_marges, analyser_aberrants,
    analyser_benchmark, analyser_charges_salariales, analyser_scoring_fiscal,
    calculer_bouclier_fiscal, generer_commentaires, generer_rapport,
    generer_previsionnel, generer_rapport_mission_ia,
    BENCHMARKS, SECTEURS_LISTE, ANTHROPIC_AVAILABLE, DOCX_AVAILABLE,
)

# ══════════════════════════════════════════════════════════════
# CSS — COMPATIBLE THÈME CLAIR & SOMBRE
# Principe : couleurs TOUJOURS explicites dans les composants HTML.
# On évite tout héritage du thème Streamlit pour les blocs custom.
# ══════════════════════════════════════════════════════════════
st.markdown("""
<style>
/* ── KPI cards ── */
.fec-kpi-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 12px; margin: 12px 0;
}
.fec-kpi {
    background: #0f2640;
    border-radius: 10px; padding: 14px 16px;
    border-left: 4px solid #00c8b4;
}
.fec-kpi .kpi-label {
    color: #7fc4e8; font-size: .72rem; font-weight: 600;
    text-transform: uppercase; letter-spacing: .06em; margin-bottom: 4px;
}
.fec-kpi .kpi-value { color: #ffffff; font-size: 1.5rem; font-weight: 800; line-height: 1.1; }
.fec-kpi .kpi-sub   { color: #7fc4e8; font-size: .7rem; margin-top: 2px; }

/* ── Score cards ── */
.fec-score-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 14px; margin: 16px 0; }
.fec-score      { border-radius: 14px; padding: 22px 16px; text-align: center; }
.fec-score .sc-num { font-size: 3rem; font-weight: 900; line-height: 1; }
.fec-score .sc-lbl { font-size: .78rem; font-weight: 700; text-transform: uppercase;
                     letter-spacing: .08em; margin-top: 4px; }
.fec-score .sc-sub { font-size: .72rem; margin-top: 6px; opacity: .85; }

/* ── Anomalie cards ── */
.fec-anomalie { border-radius: 8px; padding: 12px 16px; margin-bottom: 8px; border-left: 5px solid; }
.fec-anomalie .anom-header { display: flex; align-items: center; gap: 10px; margin-bottom: 4px; }
.fec-anomalie .anom-badge  {
    font-size: .68rem; font-weight: 800; padding: 2px 10px;
    border-radius: 20px; white-space: nowrap; letter-spacing: .05em;
}
.fec-anomalie .anom-title  { font-weight: 700; font-size: .93rem; }
.fec-anomalie .anom-detail { font-size: .84rem; margin-top: 2px; line-height: 1.45; }
.fec-anomalie .anom-meta   { font-size: .74rem; margin-top: 4px; }

/* CRITIQUE */
.anom-crit { background: #3d0f0f; border-color: #ef4444; }
.anom-crit .anom-badge  { background: #ef4444; color: #fff; }
.anom-crit .anom-title  { color: #fca5a5; }
.anom-crit .anom-detail { color: #fcd0d0; }
.anom-crit .anom-meta   { color: #f87171; }

/* ALERTE */
.anom-alert { background: #3d1f0a; border-color: #f97316; }
.anom-alert .anom-badge  { background: #f97316; color: #fff; }
.anom-alert .anom-title  { color: #fdba74; }
.anom-alert .anom-detail { color: #fed7aa; }
.anom-alert .anom-meta   { color: #fb923c; }

/* ATTENTION */
.anom-att { background: #2d2506; border-color: #eab308; }
.anom-att .anom-badge  { background: #eab308; color: #000; }
.anom-att .anom-title  { color: #fde047; }
.anom-att .anom-detail { color: #fef08a; }
.anom-att .anom-meta   { color: #facc15; }

/* INFO */
.anom-info { background: #052e1a; border-color: #22c55e; }
.anom-info .anom-badge  { background: #22c55e; color: #000; }
.anom-info .anom-title  { color: #86efac; }
.anom-info .anom-detail { color: #bbf7d0; }
.anom-info .anom-meta   { color: #4ade80; }

/* ── Lignes tableau ── */
.fec-row {
    display: flex; justify-content: space-between; align-items: center;
    padding: 8px 12px; border-bottom: 1px solid #1e3050; font-size: .87rem;
}
.fec-row .row-label { color: #94a3b8; }
.fec-row .row-val   { font-weight: 700; color: #e2e8f0; }
.fec-row .row-pos   { font-weight: 700; color: #4ade80; }
.fec-row .row-neg   { font-weight: 700; color: #f87171; }
.fec-row .row-warn  { font-weight: 700; color: #fbbf24; }

/* ── Axe bouclier ── */
.fec-axe { background: #0f2036; border-radius: 8px; padding: 10px 14px; margin-bottom: 8px; }
.fec-axe .axe-top { display: flex; justify-content: space-between; margin-bottom: 4px; }
.fec-axe .axe-lbl { color: #cbd5e1; font-weight: 600; font-size: .87rem; }
.fec-axe .axe-pts { font-weight: 800; font-size: .9rem; }
.fec-axe .axe-bar { background: #1e3a50; border-radius: 4px; height: 8px; margin: 5px 0; }
.fec-axe .axe-fill{ height: 8px; border-radius: 4px; }
.fec-axe .axe-det { color: #64748b; font-size: .77rem; }

/* ── Commentaire ── */
.fec-comment { border-radius: 8px; padding: 10px 14px; margin-bottom: 8px; border-left: 4px solid; }
.com-green  { background: #052e1a; border-color: #22c55e; }
.com-green  .com-title { color: #4ade80; font-weight: 700; }
.com-green  .com-text  { color: #bbf7d0; }
.com-orange { background: #3d1f0a; border-color: #f97316; }
.com-orange .com-title { color: #fdba74; font-weight: 700; }
.com-orange .com-text  { color: #fed7aa; }
.com-red    { background: #3d0f0f; border-color: #ef4444; }
.com-red    .com-title { color: #fca5a5; font-weight: 700; }
.com-red    .com-text  { color: #fcd0d0; }
.com-gray   { background: #0f1e2e; border-color: #475569; }
.com-gray   .com-title { color: #94a3b8; font-weight: 700; }
.com-gray   .com-text  { color: #cbd5e1; }

/* ── Suggestion ── */
.fec-suggest {
    background: #0f2640; border: 1px solid #1e4a7a;
    border-radius: 8px; padding: 10px 16px; margin-bottom: 8px;
    color: #93c5fd; font-size: .87rem;
}
.fec-suggest b { color: #60a5fa; }

/* ── Prévisionnel table ── */
.prev-table { width: 100%; border-collapse: collapse; font-size: .86rem; }
.prev-table th {
    background: #1e3a5c; color: #93c5fd;
    padding: 8px 12px; text-align: left; font-weight: 700;
}
.prev-table td { padding: 7px 12px; border-bottom: 1px solid #1e3050; }
.prev-table tr:nth-child(even) td { background: #0a1928; }
.prev-table .td-label { color: #94a3b8; }
.prev-table .td-val   { color: #e2e8f0; font-weight: 600; text-align: right; }
.prev-table .td-pos   { color: #4ade80; font-weight: 700; text-align: right; }
.prev-table .td-neg   { color: #f87171; font-weight: 700; text-align: right; }
.prev-table .td-bold  { color: #60a5fa; font-weight: 800; }

/* ── Section header ── */
.sec-header {
    background: linear-gradient(135deg, #0f2640 0%, #1a3a5c 100%);
    color: #e2e8f0; padding: 10px 18px; border-radius: 8px;
    font-weight: 700; font-size: 1rem; margin-bottom: 12px;
    border-left: 4px solid #00c8b4;
}
/* ── Empty state ── */
.fec-empty {
    background: #052e1a; color: #4ade80; padding: 12px 16px;
    border-radius: 8px; text-align: center; font-weight: 600;
    border: 1px solid #166534;
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# HELPERS HTML
# ══════════════════════════════════════════════════════════════

def anom_card(a: dict) -> str:
    g = a.get('gravite', 'INFO')
    cls = {'CRITIQUE': 'anom-crit', 'ALERTE': 'anom-alert',
           'ATTENTION': 'anom-att', 'INFO': 'anom-info'}.get(g, 'anom-info')
    ref = f'<div class="anom-meta">📎 {a["reference"]}</div>' if a.get('reference') else ''
    mnt = (f'<div class="anom-meta">💶 {a["montant"]:,.0f} €</div>'
           if a.get('montant', 0) else '')
    return (
        f'<div class="fec-anomalie {cls}">'
        f'<div class="anom-header">'
        f'<span class="anom-badge">{g}</span>'
        f'<span class="anom-title">{a.get("type","")}</span>'
        f'</div>'
        f'<div class="anom-detail">{a.get("detail","")}</div>'
        f'{mnt}{ref}'
        f'</div>'
    )

def row_html(label, val, kind='normal'):
    cls = {'pos': 'row-pos', 'neg': 'row-neg', 'warn': 'row-warn'}.get(kind, 'row-val')
    return (f'<div class="fec-row">'
            f'<span class="row-label">{label}</span>'
            f'<span class="{cls}">{val}</span>'
            f'</div>')

def axe_card(info: dict) -> str:
    pts = info.get('points', 0); mx = info.get('max', 1)
    lbl = info.get('label', '');  det = info.get('detail', '')
    pct = round(pts / mx * 100) if mx else 0
    col = '#ef4444' if pct >= 80 else '#f97316' if pct >= 50 else '#eab308' if pct >= 25 else '#22c55e'
    return (
        f'<div class="fec-axe">'
        f'<div class="axe-top">'
        f'<span class="axe-lbl">{lbl}</span>'
        f'<span class="axe-pts" style="color:{col}">{pts}/{mx}</span>'
        f'</div>'
        f'<div class="axe-bar"><div class="axe-fill" style="width:{pct}%;background:{col}"></div></div>'
        f'<div class="axe-det">{det}</div>'
        f'</div>'
    )

def comment_card(com: dict) -> str:
    cmap = {'red': 'com-red', 'orange': 'com-orange', 'green': 'com-green', 'gray': 'com-gray'}
    imap = {'red': '🔴', 'orange': '🟠', 'green': '🟢', 'gray': '⚪'}
    cls  = cmap.get(com.get('couleur', 'gray'), 'com-gray')
    icon = imap.get(com.get('couleur', 'gray'), '⚪')
    return (
        f'<div class="fec-comment {cls}">'
        f'<div class="com-title">{icon} {com.get("titre","")}</div>'
        f'<div class="com-text" style="margin-top:4px;font-size:.84rem">{com.get("texte","")}</div>'
        f'</div>'
    )

def score_card_html(num, lbl, sub, color):
    return (
        f'<div class="fec-score" style="background:{color}22;border:2px solid {color}">'
        f'<div class="sc-num" style="color:{color}">{num}</div>'
        f'<div class="sc-lbl" style="color:{color}">{lbl}</div>'
        f'<div class="sc-sub" style="color:{color}">{sub}</div>'
        f'</div>'
    )

def couleur_score(s):    return '#ef4444' if s >= 50 else '#f97316' if s >= 25 else '#22c55e'
def couleur_bouclier(s): return '#ef4444' if s >= 80 else '#f97316' if s >= 60 else '#eab308' if s >= 40 else '#22c55e'


# ══════════════════════════════════════════════════════════════
# CONTRÔLES
# ══════════════════════════════════════════════════════════════
def run_controls(df, secteur_code):
    all_a = []; stats = {}
    controles = [
        ("Équilibre D/C",       analyser_equilibre,         'equilibre'),
        ("Loi de Benford",      analyser_benford,           'benford'),
        ("Montants ronds",      analyser_montants_ronds,    'montants_ronds'),
        ("Week-end",            analyser_weekend,           'weekend'),
        ("Soldes anormaux",     analyser_soldes,            'soldes'),
        ("Doublons",            analyser_doublons,          'doublons'),
        ("Écritures inversées", analyser_inversees,         'inversees'),
        ("Fin d'exercice",      analyser_fin_exercice,      'fin_exercice'),
        ("Concentration",       analyser_concentration,     'concentration'),
        ("Saisonnalité",        analyser_saisonnalite,      'saisonnalite'),
        ("Marges et TVA",       analyser_marges,            'marges'),
        ("Montants aberrants",  analyser_aberrants,         'aberrants'),
        ("Benchmark sectoriel", lambda d: analyser_benchmark(d, secteur_code), 'benchmark'),
        ("Charges salariales",  analyser_charges_salariales,                   'charges_salariales'),
        ("Scoring fiscal",      analyser_scoring_fiscal,                        'scoring_fiscal'),
    ]
    prog = st.progress(0, text="Démarrage…")
    for i, (nom, fn, key) in enumerate(controles):
        try:
            a, s = fn(df); all_a.extend(a); stats[key] = s
        except Exception as e:
            st.warning(f"Contrôle **{nom}** ignoré : {e}")
        prog.progress((i + 1) / len(controles), text=f"⚙️ {nom}…")
    try:
        stats['_anomalies_raw'] = all_a
        _, bouclier = calculer_bouclier_fiscal(df, stats, all_a, secteur_code)
        stats['bouclier_fiscal'] = bouclier
    except Exception as e:
        st.warning(f"Bouclier fiscal : {e}")
    prog.empty()
    ordre = {'CRITIQUE': 0, 'ALERTE': 1, 'ATTENTION': 2, 'INFO': 3}
    all_a.sort(key=lambda x: ordre.get(x['gravite'], 4))
    return all_a, stats


# ══════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🔍 FEC Audit Pro v5")
    st.caption("Analyse & Contrôle des Fichiers des Écritures Comptables")
    st.divider()

    st.subheader("📂 Fichier FEC")
    uploaded = st.file_uploader(
        "Déposez votre FEC",
        type=["txt", "csv", "tsv"],
        help="Format DGFiP : séparateur tabulation, pipe ou point-virgule.",
    )
    st.divider()

    st.subheader("🏢 Secteur d'activité")
    choix_secteur = st.selectbox(
        "Sélectionnez ou laissez Auto",
        options=["AUTO"] + [c for _, c, _ in SECTEURS_LISTE if c != 'INCONNU'] + ["INCONNU"],
        format_func=lambda c: (
            "🔍 Auto-détection" if c == "AUTO" else
            next((label for _, code, label in SECTEURS_LISTE if code == c), c)
        ),
    )
    st.divider()

    st.subheader("⚙️ Modules v5")
    with st.expander("📈 Prévisionnel N+1"):
        activer_prev = st.toggle("Activer", value=True)
        croissance_ca = st.slider("Hypothèse croissance CA (%)", -20, 50, 5, 1)
    with st.expander("🤖 Rapport IA (Claude)"):
        activer_ia = st.toggle("Activer", value=False)
        api_key_input = st.text_input("Clé API Anthropic", type="password", placeholder="sk-ant-…")
        if not ANTHROPIC_AVAILABLE:
            st.warning("Package `anthropic` absent.")
        if not DOCX_AVAILABLE:
            st.warning("Package `python-docx` absent.")
    st.divider()
    lancer = st.button("🚀 Lancer l'analyse", type="primary", use_container_width=True)


# ══════════════════════════════════════════════════════════════
# ÉCRAN ACCUEIL
# ══════════════════════════════════════════════════════════════
if not uploaded:
    st.markdown("""
    ## 🔍 FEC Audit Pro v5.0
    **Outil d'analyse et de contrôle des Fichiers des Écritures Comptables (FEC) — DGFiP**

    | Module | Détail |
    |--------|--------|
    | **15 contrôles automatiques** | Benford, doublons, week-end, soldes, TVA, marges… |
    | **Benchmark sectoriel** | Médianes Banque de France / INSEE par secteur |
    | **🛡️ Bouclier Fiscal** | Score probabilité contrôle DGFiP (0-100) |
    | **📈 Prévisionnel N+1** | Projection CA / charges / résultat |
    | **🤖 Rapport IA** | Rapport de mission rédigé par Claude |

    ---
    **Pour démarrer :** déposez votre FEC dans la barre latérale puis cliquez **🚀 Lancer l'analyse**.

    > Formats acceptés : `.txt`, `.csv`, `.tsv` — séparateur `\\t`, `|` ou `;`
    """)
    st.stop()


# ══════════════════════════════════════════════════════════════
# CHARGEMENT FEC
# ══════════════════════════════════════════════════════════════
@st.cache_data(show_spinner="Chargement du FEC…")
def load_fec(fb, fn):
    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(fn)[1]) as t:
        t.write(fb); tmp = t.name
    df = charger_fec(tmp); os.unlink(tmp); return df

try:
    df = load_fec(uploaded.read(), uploaded.name)
except Exception as e:
    st.error(f"❌ Impossible de charger le FEC : {e}"); st.stop()

with st.expander(f"📋 Aperçu — {len(df):,} lignes × {len(df.columns)} colonnes", expanded=False):
    st.dataframe(df.head(50), use_container_width=True)

with st.spinner("Détection du secteur…"):
    secteur_auto, conf_auto, _ = detecter_secteur(df)
secteur_code = secteur_auto if choix_secteur == "AUTO" else choix_secteur
sect_label = BENCHMARKS[secteur_code]['label']

if choix_secteur == "AUTO":
    st.info(f"🔍 Secteur auto-détecté : **{sect_label}** (confiance {conf_auto}%)")
else:
    st.success(f"✅ Secteur : **{sect_label}**")

if not lancer:
    st.info("👈 Cliquez **🚀 Lancer l'analyse** pour démarrer.", icon="💡")
    st.stop()


# ══════════════════════════════════════════════════════════════
# ANALYSE
# ══════════════════════════════════════════════════════════════
with st.spinner("Analyse en cours…"):
    all_anomalies, stats = run_controls(df, secteur_code)
    commentaires, suggestions = generer_commentaires(df, all_anomalies, stats, secteur_code)
    prev_data = None
    if activer_prev:
        try:
            _, prev_data = generer_previsionnel(
                df, stats, secteur_code,
                croissance_ca=croissance_ca / 100,
                filepath=uploaded.name)
            stats['previsionnel'] = prev_data
        except Exception as e:
            st.warning(f"Prévisionnel : {e}")
    html_rapport = generer_rapport(
        uploaded.name, df, all_anomalies, stats, commentaires, suggestions, secteur_code)


# ══════════════════════════════════════════════════════════════
# TABLEAU DE BORD
# ══════════════════════════════════════════════════════════════
c_cnt = Counter(a['gravite'] for a in all_anomalies)
nc, na, nat, ni = c_cnt['CRITIQUE'], c_cnt['ALERTE'], c_cnt['ATTENTION'], c_cnt['INFO']
score_risque = min(100, nc * 25 + na * 10 + nat * 3 + ni)
sc_color  = couleur_score(score_risque)
sc_label  = 'ÉLEVÉ' if score_risque >= 50 else 'MODÉRÉ' if score_risque >= 25 else 'FAIBLE'

bouclier  = stats.get('bouclier_fiscal', {})
b_score   = bouclier.get('score_global', 0)
b_niveau  = bouclier.get('niveau_risque', 'N/A')
b_color   = couleur_bouclier(b_score)

mg        = stats.get('marges', {})
bench     = stats.get('benchmark', {})
tm_e      = bench.get('tm_entreprise', 0) or 0
tm_med    = bench.get('tm_mediane', bench.get('tm_med', 0)) or 0
tm_low    = bench.get('tm_low', 0) or 0
tm_high   = bench.get('tm_high', 0) or 0
m_color   = '#22c55e' if tm_low <= tm_e <= tm_high else '#f97316'

st.divider()
st.markdown('<div class="sec-header">📊 Tableau de bord</div>', unsafe_allow_html=True)

# KPI row
kpi_html = (
    '<div class="fec-kpi-grid">'
    + f'<div class="fec-kpi"><div class="kpi-label">Lignes FEC</div>'
    + f'<div class="kpi-value">{len(df):,}</div></div>'
    + f'<div class="fec-kpi" style="border-color:{"#ef4444" if nc>0 else "#22c55e"}">'
    + f'<div class="kpi-label">Anomalies</div><div class="kpi-value">{len(all_anomalies)}</div>'
    + f'<div class="kpi-sub">🔴 {nc} critique(s)</div></div>'
    + f'<div class="fec-kpi"><div class="kpi-label">CA (70x)</div>'
    + f'<div class="kpi-value">{mg.get("ca",0):,.0f} €</div></div>'
    + f'<div class="fec-kpi" style="border-color:{m_color}">'
    + f'<div class="kpi-label">Marge brute</div>'
    + f'<div class="kpi-value" style="color:{m_color}">{mg.get("tm",0):.1f}%</div>'
    + f'<div class="kpi-sub">Médiane : {tm_med:.0f}% [{tm_low:.0f}–{tm_high:.0f}%]</div></div>'
    + f'<div class="fec-kpi"><div class="kpi-label">Résultat exploit.</div>'
    + f'<div class="kpi-value" style="color:{"#4ade80" if mg.get("res",0)>=0 else "#f87171"}">'
    + f'{mg.get("res",0):,.0f} €</div></div>'
    + '</div>'
)
st.markdown(kpi_html, unsafe_allow_html=True)

# Scores
sc_html = (
    '<div class="fec-score-grid">'
    + score_card_html(score_risque, "Score Anomalies", f"Risque {sc_label}", sc_color)
    + score_card_html(b_score, "🛡️ Bouclier Fiscal", b_niveau, b_color)
    + score_card_html(f"{tm_e:.1f}%", "Marge brute",
                      f"Médiane : {tm_med:.0f}% [{tm_low:.0f}–{tm_high:.0f}%]", m_color)
    + '</div>'
)
st.markdown(sc_html, unsafe_allow_html=True)
st.divider()


# ══════════════════════════════════════════════════════════════
# ONGLETS
# ══════════════════════════════════════════════════════════════
tabs = st.tabs([
    "🔴 Anomalies",
    "📉 Marges & TVA",
    "🛡️ Bouclier Fiscal",
    "🏢 Benchmark",
    "📈 Prévisionnel N+1",
    "💡 Commentaires",
    "🤖 Rapport IA",
    "⬇️ Téléchargements",
])


# ── TAB 1 : Anomalies ─────────────────────────────────────────
with tabs[0]:
    st.markdown(
        f'<div class="sec-header">Anomalies : {len(all_anomalies)}'
        f' &nbsp;|&nbsp; 🔴 {nc} critique(s) · 🟠 {na} alerte(s)'
        f' · 🟡 {nat} attention(s) · 🟢 {ni} info(s)</div>',
        unsafe_allow_html=True)

    filtres = st.multiselect(
        "Filtrer par gravité",
        ["CRITIQUE", "ALERTE", "ATTENTION", "INFO"],
        default=["CRITIQUE", "ALERTE", "ATTENTION", "INFO"])

    filtrees = [a for a in all_anomalies if a['gravite'] in filtres]
    if not filtrees:
        st.markdown('<div class="fec-empty">✅ Aucune anomalie pour ce filtre</div>',
                    unsafe_allow_html=True)
    else:
        st.markdown("".join(anom_card(a) for a in filtrees), unsafe_allow_html=True)


# ── TAB 2 : Marges & TVA ──────────────────────────────────────
with tabs[1]:
    st.markdown('<div class="sec-header">📉 Marges & TVA</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Compte de résultat simplifié**")
        ca = mg.get('ca', 0); achats = mg.get('achats', 0); mb = mg.get('mb', 0)
        tm = mg.get('tm', 0); cp = mg.get('cp', 0); ce = mg.get('ce', 0)
        res = mg.get('res', 0); tr = mg.get('tr', 0)
        tp = mg.get('tp', 0); tc_val = mg.get('tc', 0)
        data_rows = [
            ("Chiffre d'affaires (70x)",   f"{ca:,.0f} €",     'pos' if ca > 0 else 'normal'),
            ("— Achats (60x)",             f"{achats:,.0f} €", 'normal'),
            ("= Marge brute",              f"{mb:,.0f} €",     'pos' if mb >= 0 else 'neg'),
            ("  Taux marge brute",         f"{tm:.2f}%",        'pos' if tm > 0 else 'neg'),
            ("— Charges externes (61-62)", f"{ce:,.0f} €",     'normal'),
            ("— Charges personnel (64x)",  f"{cp:,.0f} €",     'normal'),
            ("Total produits (7xx)",       f"{tp:,.0f} €",     'normal'),
            ("Total charges (6xx)",        f"{tc_val:,.0f} €", 'normal'),
            ("= Résultat exploitation",    f"{res:,.0f} €",    'pos' if res >= 0 else 'neg'),
            ("  Taux de résultat",         f"{tr:.2f}%",        'pos' if tr >= 0 else 'neg'),
        ]
        st.markdown("".join(row_html(l, v, k) for l, v, k in data_rows), unsafe_allow_html=True)

    with col2:
        st.markdown("**TVA**")
        tva_c = mg.get('tva_c', 0); tva_d = mg.get('tva_d', 0); tx_tva = mg.get('tx_tva', 0)
        tva_rows = [
            ("TVA collectée (4457x)",  f"{tva_c:,.0f} €", 'normal'),
            ("TVA déductible (4456x)", f"{tva_d:,.0f} €", 'normal'),
            ("Taux TVA apparent",      f"{tx_tva:.2f}%",   'neg' if tx_tva > 20 else 'pos'),
        ]
        st.markdown("".join(row_html(l, v, k) for l, v, k in tva_rows), unsafe_allow_html=True)

        st.markdown("<br>**Anomalies marges & TVA**", unsafe_allow_html=True)
        marge_anom = [a for a in all_anomalies
                      if any(x in a['type'] for x in ['Marge', 'Personnel', 'TVA'])]
        if marge_anom:
            st.markdown("".join(anom_card(a) for a in marge_anom), unsafe_allow_html=True)
        else:
            st.markdown('<div class="fec-empty">✅ Aucune anomalie marge/TVA</div>',
                        unsafe_allow_html=True)


# ── TAB 3 : Bouclier Fiscal ───────────────────────────────────
with tabs[2]:
    st.markdown('<div class="sec-header">🛡️ Bouclier Fiscal — Score de Probabilité de Contrôle DGFiP</div>',
                unsafe_allow_html=True)
    if not bouclier:
        st.warning("Données non disponibles.")
    else:
        conseil = bouclier.get('conseil', '')
        details = bouclier.get('details', {})
        col_g, col_d = st.columns([1, 2])
        with col_g:
            st.markdown(
                f'<div class="fec-score" style="background:{b_color}22;border:2px solid {b_color}">'
                f'<div class="sc-num" style="color:{b_color}">{b_score}/100</div>'
                f'<div class="sc-lbl" style="color:{b_color}">{b_niveau}</div>'
                f'<div class="sc-sub" style="color:{b_color};font-size:.78rem;margin-top:10px">{conseil}</div>'
                f'</div>', unsafe_allow_html=True)
        with col_d:
            st.markdown("**Détail par axe**")
            st.markdown("".join(axe_card(v) for v in details.values()),
                        unsafe_allow_html=True)


# ── TAB 4 : Benchmark ─────────────────────────────────────────
with tabs[3]:
    st.markdown(f'<div class="sec-header">🏢 Benchmark — {sect_label}</div>',
                unsafe_allow_html=True)
    note = BENCHMARKS[secteur_code].get('note', '')
    st.markdown(
        f'<div class="fec-comment com-gray">'
        f'<div class="com-text">💡 {note}</div>'
        f'</div>', unsafe_allow_html=True)

    sal  = stats.get('charges_salariales', {})
    fisc = stats.get('scoring_fiscal', {})
    col_b1, col_b2 = st.columns(2)

    with col_b1:
        st.markdown("**Comparaison marges**")
        ecart = tm_e - tm_med
        in_range = tm_low <= tm_e <= tm_high
        c_e = 'pos' if in_range else 'neg'
        st.markdown("".join(row_html(l, v, k) for l, v, k in [
            ("Marge brute entreprise", f"{tm_e:.1f}%", c_e),
            ("Médiane sectorielle",    f"{tm_med:.0f}%", 'normal'),
            ("Fourchette secteur",     f"[{tm_low:.0f}% — {tm_high:.0f}%]", 'normal'),
            ("Écart vs médiane",       f"{'+'if ecart>=0 else ''}{ecart:.1f} pts", c_e),
        ]), unsafe_allow_html=True)

    with col_b2:
        st.markdown("**Charges & Fiscal**")
        cp_e   = bench.get('cp_ca_entreprise', 0) or 0
        cp_med = bench.get('cp_ca_mediane', 0) or 0
        ecart_cp = cp_e - cp_med
        c_cp = 'neg' if abs(ecart_cp) > 15 else 'pos'
        rc = sal.get('ratio_cotis', 0)
        st.markdown("".join(row_html(l, v, k) for l, v, k in [
            ("Personnel/CA entreprise", f"{cp_e:.1f}%", c_cp),
            ("Personnel/CA secteur",    f"{cp_med:.0f}%", 'normal'),
            ("Écart charges pers.",     f"{'+'if ecart_cp>=0 else ''}{ecart_cp:.1f} pts", c_cp),
            ("Salaires bruts",          f"{sal.get('sal_bruts',0):,.0f} €", 'normal'),
            ("Cotisations patronales",  f"{sal.get('cotis_pat',0):,.0f} €", 'normal'),
            ("Ratio cotis./salaires",   f"{rc:.1f}%", 'pos' if 20 <= rc <= 60 else 'neg'),
            ("Effectif estimé",         f"~{sal.get('nb_sal_estime',0)} salarié(s)", 'normal'),
            ("IS comptabilisé",         f"{fisc.get('is_d',0):,.0f} €", 'normal'),
            ("Taux IS apparent",        f"{fisc.get('tx_is',0):.1f}%",
             'neg' if fisc.get('tx_is', 0) > 35 else 'normal'),
        ]), unsafe_allow_html=True)

    st.markdown("<br>**Anomalies benchmark**", unsafe_allow_html=True)
    bench_anom = [a for a in all_anomalies
                  if any(x in a['type'].lower() for x in
                         ['mediane', 'sectoriel', 'cotisations', 'inattendue'])]
    if bench_anom:
        st.markdown("".join(anom_card(a) for a in bench_anom), unsafe_allow_html=True)
    else:
        st.markdown('<div class="fec-empty">✅ Aucune anomalie sectorielle</div>',
                    unsafe_allow_html=True)


# ── TAB 5 : Prévisionnel N+1 ─────────────────────────────────
with tabs[4]:
    st.markdown('<div class="sec-header">📈 Prévisionnel N+1</div>', unsafe_allow_html=True)
    if not activer_prev:
        st.info("Module désactivé. Activez-le dans la barre latérale.")
    elif not prev_data or not prev_data.get('ca_n1'):
        st.warning("Prévisionnel non généré (CA non détecté ou données insuffisantes).")
    else:
        an = str(prev_data.get('annee_n', 'N'))
        try: an1 = str(int(an) + 1)
        except: an1 = 'N+1'

        st.markdown(
            f'<div class="fec-comment com-gray"><div class="com-text">'
            f'Hypothèse croissance CA : <b>+{croissance_ca}%</b> &nbsp;|&nbsp; '
            f'Taux IS : <b>{prev_data.get("taux_is",0.25)*100:.0f}%</b>'
            f'</div></div>', unsafe_allow_html=True)

        def fmt(v): return "—" if v is None else f"{v:,.0f} €"
        def var_str(n, n1):
            if n is None or n == 0: return "—"
            v = n1 - n; p = v / abs(n) * 100
            return f"{'+'if v>=0 else ''}{v:,.0f} ({p:+.1f}%)"
        def var_cls(n, n1):
            if n is None or n == 0: return "td-val"
            return "td-pos" if n1 >= n else "td-neg"

        rows_p = [
            ("Chiffre d'affaires", prev_data.get('ca_n',0),  prev_data.get('ca_n1',0),  True),
            ("Achats",             prev_data.get('achats_n',0), prev_data.get('achats_n1',0), False),
            ("Marge brute",        prev_data.get('ca_n',0)-prev_data.get('achats_n',0),
                                   prev_data.get('mb_n1',0), True),
            ("Charges externes",   prev_data.get('ce_n',0),  prev_data.get('ce_n1',0),  False),
            ("Charges personnel",  prev_data.get('cp_n',0),  prev_data.get('cp_n1',0),  False),
            ("EBE",                None,                      prev_data.get('ebe_n1',0), True),
            ("Amortissements",     prev_data.get('dot_amo_n',0), prev_data.get('dot_amo_n1',0), False),
            ("Résultat exploitation", prev_data.get('res_n',0), prev_data.get('rex_n1',0), True),
            ("IS",                 None,                      prev_data.get('is_n1',0),  False),
            ("Résultat net",       prev_data.get('res_n',0), prev_data.get('res_net_n1',0), True),
        ]

        thead = (f'<thead><tr>'
                 f'<th class="td-label" style="width:38%">Poste</th>'
                 f'<th style="text-align:right">{an} (réel)</th>'
                 f'<th style="text-align:right">{an1} (prévis.)</th>'
                 f'<th style="text-align:right">Variation</th>'
                 f'</tr></thead>')
        tbody = '<tbody>'
        for lbl, n, n1, bold in rows_p:
            bold_cls = ' td-bold' if bold else ''
            vc = var_cls(n, n1)
            tbody += (f'<tr>'
                      f'<td class="td-label{bold_cls}">{lbl}</td>'
                      f'<td class="td-val" style="text-align:right">{fmt(n)}</td>'
                      f'<td class="td-val{bold_cls}" style="text-align:right">{fmt(n1)}</td>'
                      f'<td class="{vc}" style="text-align:right">{var_str(n, n1)}</td>'
                      f'</tr>')
        tbody += '</tbody>'
        st.markdown(f'<table class="prev-table">{thead}{tbody}</table>',
                    unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        c1.metric(f"Taux marge brute {an1}",  f"{prev_data.get('tm_n1',0):.1f}%")
        c2.metric(f"Taux résultat net {an1}", f"{prev_data.get('tr_n1',0):.1f}%")
        c3.metric(f"Charges pers./CA {an1}",  f"{prev_data.get('cp_ca_n1',0):.1f}%")


# ── TAB 6 : Commentaires ──────────────────────────────────────
with tabs[5]:
    st.markdown('<div class="sec-header">💡 Commentaires & Suggestions</div>',
                unsafe_allow_html=True)
    if commentaires:
        st.markdown("**Commentaires automatiques**")
        st.markdown("".join(comment_card(c) for c in commentaires), unsafe_allow_html=True)
    if suggestions:
        st.markdown("<br>**Recommandations**", unsafe_allow_html=True)
        st.markdown(
            "".join(f'<div class="fec-suggest"><b>{i}.</b> {s}</div>'
                    for i, s in enumerate(suggestions, 1)),
            unsafe_allow_html=True)


# ── TAB 7 : Rapport IA ────────────────────────────────────────
with tabs[6]:
    st.markdown('<div class="sec-header">🤖 Rapport de Mission IA — Module 1</div>',
                unsafe_allow_html=True)
    if not activer_ia:
        st.info("Activez le module dans la barre latérale et saisissez votre clé API.")
    elif not api_key_input:
        st.warning("Clé API Anthropic manquante (barre latérale).")
    elif not ANTHROPIC_AVAILABLE:
        st.error("Package `anthropic` non installé — ajoutez-le dans requirements.txt.")
    else:
        if st.button("✍️ Générer le rapport de mission", type="primary"):
            with st.spinner("Rédaction par Claude…"):
                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as t:
                    tmp_docx = t.name
                path, texte = generer_rapport_mission_ia(
                    uploaded.name, stats, all_anomalies, secteur_code,
                    api_key=api_key_input, output_docx=tmp_docx)
            if texte and not texte.startswith("MODULE"):
                st.success("✅ Rapport généré !")
                with st.expander("📄 Aperçu du rapport", expanded=True):
                    st.markdown(
                        f'<div class="fec-comment com-gray">'
                        f'<div class="com-text" style="white-space:pre-wrap;font-size:.84rem">'
                        f'{texte}</div></div>',
                        unsafe_allow_html=True)
                if path and os.path.exists(path):
                    with open(path, 'rb') as f: data = f.read()
                    st.download_button(
                        "⬇️ Télécharger rapport Word (.docx)", data,
                        file_name=f"rapport_{os.path.splitext(uploaded.name)[0]}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    os.unlink(path)
                else:
                    st.download_button(
                        "⬇️ Télécharger rapport (texte)", texte.encode('utf-8'),
                        file_name=f"rapport_{os.path.splitext(uploaded.name)[0]}.txt",
                        mime="text/plain")
            else:
                st.error(f"Erreur : {texte}")


# ── TAB 8 : Téléchargements ───────────────────────────────────
with tabs[7]:
    st.markdown('<div class="sec-header">⬇️ Téléchargements</div>', unsafe_allow_html=True)

    st.markdown("**Rapport HTML complet** (graphiques interactifs Chart.js)")
    st.download_button(
        "⬇️ Rapport HTML", html_rapport.encode('utf-8'),
        file_name=f"{os.path.splitext(uploaded.name)[0]}_audit_v5.html",
        mime="text/html", type="primary")

    st.divider()
    st.markdown("**Anomalies — CSV**")
    if all_anomalies:
        import pandas as pd
        st.download_button(
            "⬇️ Anomalies CSV",
            pd.DataFrame(all_anomalies).to_csv(index=False, sep=';', encoding='utf-8-sig'),
            file_name=f"{os.path.splitext(uploaded.name)[0]}_anomalies.csv",
            mime="text/csv")
    else:
        st.info("Aucune anomalie à exporter.")

    st.divider()
    st.markdown("**Statistiques — JSON brut**")
    stats_export = {k: v for k, v in stats.items() if k != '_anomalies_raw'}
    st.download_button(
        "⬇️ Stats JSON",
        json.dumps(stats_export, ensure_ascii=False, indent=2, default=str).encode('utf-8'),
        file_name=f"{os.path.splitext(uploaded.name)[0]}_stats.json",
        mime="application/json")


# ── Footer ────────────────────────────────────────────────────
st.divider()
st.markdown(
    '<div style="color:#475569;font-size:.75rem;text-align:center">'
    "FEC Audit Pro v5.0 — Référentiels DGFiP &amp; Banque de France — "
    "Document d'aide à la révision, à valider par un expert-comptable."
    '</div>', unsafe_allow_html=True)
