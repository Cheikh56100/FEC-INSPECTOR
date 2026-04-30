#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FEC Audit Pro v5.0 — Interface Streamlit
Héberge toutes les fonctionnalités de fec_audit_pro_v5.py dans une UI web.
"""

import streamlit as st
import tempfile, os, io, sys, warnings
warnings.filterwarnings('ignore')

# ── Configuration de la page ──────────────────────────────────
st.set_page_config(
    page_title="FEC Audit Pro v5",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Import du moteur métier ────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from fec_audit_pro_v5 import (
    charger_fec,
    detecter_secteur,
    analyser_equilibre,
    analyser_benford,
    analyser_montants_ronds,
    analyser_weekend,
    analyser_soldes,
    analyser_doublons,
    analyser_inversees,
    analyser_fin_exercice,
    analyser_concentration,
    analyser_saisonnalite,
    analyser_marges,
    analyser_aberrants,
    analyser_benchmark,
    analyser_charges_salariales,
    analyser_scoring_fiscal,
    calculer_bouclier_fiscal,
    generer_commentaires,
    generer_rapport,
    generer_previsionnel,
    generer_rapport_mission_ia,
    BENCHMARKS,
    SECTEURS_LISTE,
    ANTHROPIC_AVAILABLE,
    DOCX_AVAILABLE,
)

# ══════════════════════════════════════════════════════════════
# CSS PERSONNALISÉ
# ══════════════════════════════════════════════════════════════
st.markdown("""
<style>
/* Palette */
:root {
    --brand: #1a3a5c;
    --cyan:  #00e5c8;
    --red:   #ef4444;
    --orange:#f97316;
    --green: #22c55e;
    --yellow:#eab308;
}

/* Sidebar header */
section[data-testid="stSidebar"] { background: #f0f4f8; }

/* Badges gravité */
.badge-critique { background:#fef2f2; color:#dc2626; border:1px solid #dc2626;
                  border-radius:4px; padding:2px 8px; font-size:.75rem; font-weight:700; }
.badge-alerte   { background:#fff7ed; color:#ea580c; border:1px solid #ea580c;
                  border-radius:4px; padding:2px 8px; font-size:.75rem; font-weight:700; }
.badge-attention{ background:#fffbeb; color:#ca8a04; border:1px solid #ca8a04;
                  border-radius:4px; padding:2px 8px; font-size:.75rem; font-weight:700; }
.badge-info     { background:#f0fdf4; color:#16a34a; border:1px solid #16a34a;
                  border-radius:4px; padding:2px 8px; font-size:.75rem; font-weight:700; }

/* Score jauge */
.score-card { text-align:center; padding:1rem; border-radius:12px; }
.score-big  { font-size:3.5rem; font-weight:800; line-height:1; }
.score-lbl  { font-size:1rem; font-weight:600; letter-spacing:.05em; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# FONCTIONS UTILITAIRES
# ══════════════════════════════════════════════════════════════

def badge_html(gravite: str) -> str:
    cls = {'CRITIQUE': 'badge-critique', 'ALERTE': 'badge-alerte',
           'ATTENTION': 'badge-attention', 'INFO': 'badge-info'}.get(gravite, 'badge-info')
    return f'<span class="{cls}">{gravite}</span>'


def couleur_score(score: int) -> str:
    if score >= 50: return "#ef4444"
    if score >= 25: return "#f97316"
    return "#22c55e"


def afficher_anomalies(anomalies, label_vide="Aucune anomalie détectée ✅"):
    if not anomalies:
        st.success(label_vide)
        return
    for a in anomalies:
        bg = {'CRITIQUE': '#fef2f2', 'ALERTE': '#fff7ed',
              'ATTENTION': '#fffbeb', 'INFO': '#f0fdf4'}.get(a['gravite'], '#f9fafb')
        bc = {'CRITIQUE': '#dc2626', 'ALERTE': '#ea580c',
              'ATTENTION': '#ca8a04', 'INFO': '#16a34a'}.get(a['gravite'], '#6b7280')
        st.markdown(
            f"""<div style="background:{bg};border-left:4px solid {bc};
                padding:.6rem 1rem;border-radius:6px;margin-bottom:.4rem">
                <b style="color:{bc}">[{a['gravite']}]</b>
                <b> {a['type']}</b><br>
                <small>{a['detail']}</small>
                {"<br><small><i>Réf : " + a['reference'] + "</i></small>" if a.get('reference') else ""}
                {"<br><small><b>Montant : " + f"{a['montant']:,.0f} EUR" + "</b></small>" if a.get('montant',0) else ""}
            </div>""",
            unsafe_allow_html=True,
        )


def run_controls(df, secteur_code):
    """Lance tous les contrôles et retourne (all_anomalies, stats_dict)."""
    all_a = []; stats = {}
    controles = [
        ("Équilibre D/C",        analyser_equilibre,         'equilibre'),
        ("Loi de Benford",       analyser_benford,           'benford'),
        ("Montants ronds",       analyser_montants_ronds,    'montants_ronds'),
        ("Week-end",             analyser_weekend,           'weekend'),
        ("Soldes anormaux",      analyser_soldes,            'soldes'),
        ("Doublons",             analyser_doublons,          'doublons'),
        ("Écritures inversées",  analyser_inversees,         'inversees'),
        ("Fin d'exercice",       analyser_fin_exercice,      'fin_exercice'),
        ("Concentration",        analyser_concentration,     'concentration'),
        ("Saisonnalité",         analyser_saisonnalite,      'saisonnalite'),
        ("Marges et TVA",        analyser_marges,            'marges'),
        ("Montants aberrants",   analyser_aberrants,         'aberrants'),
        ("Benchmark sectoriel",  lambda d: analyser_benchmark(d, secteur_code), 'benchmark'),
        ("Charges salariales",   analyser_charges_salariales,                   'charges_salariales'),
        ("Scoring fiscal",       analyser_scoring_fiscal,                        'scoring_fiscal'),
    ]
    progress = st.progress(0, text="Initialisation des contrôles…")
    for i, (nom, fn, key) in enumerate(controles):
        try:
            a, s = fn(df)
            all_a.extend(a); stats[key] = s
        except Exception as e:
            st.warning(f"Contrôle **{nom}** : erreur — {e}")
        progress.progress((i + 1) / len(controles), text=f"Contrôle : {nom}…")

    # Bouclier fiscal
    try:
        stats['_anomalies_raw'] = all_a
        _, bouclier = calculer_bouclier_fiscal(df, stats, all_a, secteur_code)
        stats['bouclier_fiscal'] = bouclier
    except Exception as e:
        st.warning(f"Bouclier fiscal : erreur — {e}")

    progress.empty()

    ordre = {'CRITIQUE': 0, 'ALERTE': 1, 'ATTENTION': 2, 'INFO': 3}
    all_a.sort(key=lambda x: ordre.get(x['gravite'], 4))
    return all_a, stats


# ══════════════════════════════════════════════════════════════
# SIDEBAR — PARAMÈTRES
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.image("https://img.icons8.com/color/96/search-property.png", width=60)
    st.title("FEC Audit Pro v5")
    st.caption("Analyse & Contrôle des Fichiers des Écritures Comptables")
    st.divider()

    # ── Fichier FEC ──
    st.subheader("📂 Fichier FEC")
    uploaded = st.file_uploader(
        "Déposez votre FEC ici",
        type=["txt", "csv", "tsv"],
        help="Format DGFiP : séparateur tabulation, pipe ou point-virgule.",
    )

    st.divider()

    # ── Secteur d'activité ──
    st.subheader("🏢 Secteur d'activité")
    choix_secteur = st.selectbox(
        "Sélectionnez (ou laissez sur Auto-détection)",
        options=["AUTO"] + [code for _, code, _ in SECTEURS_LISTE if code != 'INCONNU'] + ["INCONNU"],
        format_func=lambda c: (
            "🔍 Auto-détection" if c == "AUTO" else
            next((label for _, code, label in SECTEURS_LISTE if code == c), c)
        ),
    )

    st.divider()

    # ── Paramètres modules v5 ──
    st.subheader("⚙️ Modules v5")

    with st.expander("📈 Prévisionnel N+1", expanded=False):
        activer_prev = st.toggle("Générer le prévisionnel N+1", value=True)
        croissance_ca = st.slider(
            "Hypothèse croissance CA (%)", min_value=-20, max_value=50, value=5, step=1
        )

    with st.expander("🤖 Rapport de Mission IA (Claude)", expanded=False):
        activer_ia = st.toggle("Générer le rapport IA", value=False)
        api_key_input = st.text_input(
            "Clé API Anthropic",
            type="password",
            placeholder="sk-ant-...",
            help="Nécessite un accès à l'API Anthropic (claude-sonnet-4-20250514).",
        )
        if not ANTHROPIC_AVAILABLE:
            st.warning("Package `anthropic` non installé.")
        if not DOCX_AVAILABLE:
            st.warning("Package `python-docx` non installé (export Word désactivé).")

    st.divider()

    lancer = st.button("🚀 Lancer l'analyse", type="primary", use_container_width=True)

# ══════════════════════════════════════════════════════════════
# ÉCRAN D'ACCUEIL (avant upload)
# ══════════════════════════════════════════════════════════════
if not uploaded:
    st.markdown("""
    ## 🔍 FEC Audit Pro v5.0
    **Outil d'analyse et de contrôle des Fichiers des Écritures Comptables (FEC)**

    ### Fonctionnalités
    | Module | Description |
    |--------|-------------|
    | **15 contrôles automatiques** | Benford, doublons, week-end, soldes anormaux, TVA, marges… |
    | **Benchmark sectoriel** | Comparaison vs médianes Banque de France / INSEE par secteur |
    | **🛡️ Bouclier Fiscal** | Score de probabilité de contrôle DGFiP (0-100) |
    | **📈 Prévisionnel N+1** | Projection automatique CA / charges / résultat |
    | **🤖 Rapport IA** | Rédaction automatique d'un rapport de mission par Claude |

    ### Pour commencer
    1. Déposez votre fichier FEC dans la barre latérale (`.txt`, `.csv`, `.tsv`)
    2. Sélectionnez ou confirmez le secteur d'activité
    3. Configurez les modules souhaités
    4. Cliquez sur **🚀 Lancer l'analyse**

    ---
    > Format attendu : FEC DGFiP — séparateur tabulation, pipe (`|`) ou point-virgule.
    > Colonnes obligatoires : `CompteNum`, `Debit`, `Credit`, `EcritureDate`.
    """)
    st.stop()

# ══════════════════════════════════════════════════════════════
# CHARGEMENT DU FICHIER
# ══════════════════════════════════════════════════════════════
@st.cache_data(show_spinner="Chargement du FEC…")
def load_fec(file_bytes, filename):
    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(filename)[1]) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    df = charger_fec(tmp_path)
    os.unlink(tmp_path)
    return df

try:
    df = load_fec(uploaded.read(), uploaded.name)
except Exception as e:
    st.error(f"❌ Impossible de charger le FEC : {e}")
    st.stop()

# Affichage aperçu
with st.expander(f"📋 Aperçu du FEC — {len(df):,} lignes × {len(df.columns)} colonnes", expanded=False):
    st.dataframe(df.head(50), use_container_width=True)

# ── Détection du secteur ──
with st.spinner("Détection du secteur d'activité…"):
    secteur_auto, conf_auto, _ = detecter_secteur(df)

if choix_secteur == "AUTO":
    secteur_code = secteur_auto
    st.info(
        f"🔍 Secteur auto-détecté : **{BENCHMARKS[secteur_code]['label']}** "
        f"(confiance {conf_auto}%). Vous pouvez forcer le secteur dans la barre latérale.",
        icon="ℹ️",
    )
else:
    secteur_code = choix_secteur
    st.success(f"✅ Secteur sélectionné : **{BENCHMARKS[secteur_code]['label']}**")

# ══════════════════════════════════════════════════════════════
# LANCEMENT DE L'ANALYSE
# ══════════════════════════════════════════════════════════════
if not lancer:
    st.info("👈 Cliquez sur **🚀 Lancer l'analyse** pour démarrer.", icon="💡")
    st.stop()

# ── Analyse principale ──
with st.spinner("Analyse en cours…"):
    all_anomalies, stats = run_controls(df, secteur_code)
    commentaires, suggestions = generer_commentaires(df, all_anomalies, stats, secteur_code)

    # Prévisionnel
    prev_data = None
    if activer_prev:
        try:
            _, prev_data = generer_previsionnel(
                df, stats, secteur_code,
                croissance_ca=croissance_ca / 100,
                filepath=uploaded.name,
            )
            stats['previsionnel'] = prev_data
        except Exception as e:
            st.warning(f"Prévisionnel N+1 : {e}")

    # Rapport HTML
    html_rapport = generer_rapport(
        uploaded.name, df, all_anomalies, stats, commentaires, suggestions, secteur_code
    )

# ══════════════════════════════════════════════════════════════
# DASHBOARD RÉSULTATS
# ══════════════════════════════════════════════════════════════
from collections import Counter

c = Counter(a['gravite'] for a in all_anomalies)
nc, na, nat, ni = c.get('CRITIQUE', 0), c.get('ALERTE', 0), c.get('ATTENTION', 0), c.get('INFO', 0)
score_risque = min(100, nc * 25 + na * 10 + nat * 3 + ni)
sc_color = couleur_score(score_risque)
sc_label = 'ÉLEVÉ' if score_risque >= 50 else 'MODÉRÉ' if score_risque >= 25 else 'FAIBLE'

bouclier = stats.get('bouclier_fiscal', {})
b_score  = bouclier.get('score_global', 0)
b_niveau = bouclier.get('niveau_risque', 'N/A')
b_color  = '#ef4444' if b_score >= 80 else '#f97316' if b_score >= 60 else '#eab308' if b_score >= 40 else '#22c55e'

mg = stats.get('marges', {})

st.divider()
st.subheader("📊 Tableau de bord")

# ── Métriques principales ──
col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("Lignes FEC", f"{len(df):,}")
with col2:
    st.metric("Anomalies totales", len(all_anomalies),
              delta=f"{nc} critique(s)", delta_color="inverse" if nc > 0 else "off")
with col3:
    st.metric("CA (70x)", f"{mg.get('ca', 0):,.0f} €")
with col4:
    st.metric("Marge brute", f"{mg.get('tm', 0):.1f}%")
with col5:
    st.metric("Résultat exploitation", f"{mg.get('res', 0):,.0f} €")

st.divider()

# ── Scores ──
col_s1, col_s2, col_s3 = st.columns(3)
with col_s1:
    st.markdown(
        f"""<div class="score-card" style="background:{sc_color}22;border:2px solid {sc_color}">
            <div class="score-big" style="color:{sc_color}">{score_risque}</div>
            <div class="score-lbl" style="color:{sc_color}">Score Anomalies</div>
            <div style="font-size:.85rem;margin-top:.3rem">Risque {sc_label}</div>
        </div>""", unsafe_allow_html=True
    )
with col_s2:
    st.markdown(
        f"""<div class="score-card" style="background:{b_color}22;border:2px solid {b_color}">
            <div class="score-big" style="color:{b_color}">{b_score}</div>
            <div class="score-lbl" style="color:{b_color}">🛡️ Bouclier Fiscal</div>
            <div style="font-size:.85rem;margin-top:.3rem">{b_niveau}</div>
        </div>""", unsafe_allow_html=True
    )
with col_s3:
    bench = stats.get('benchmark', {})
    tm_e  = bench.get('tm_entreprise', 0) or 0
    tm_med = bench.get('tm_mediane', bench.get('tm_med', 0)) or 0
    tm_low = bench.get('tm_low', 0) or 0
    tm_high= bench.get('tm_high', 0) or 0
    pos_color = "#22c55e" if tm_low <= tm_e <= tm_high else "#f97316"
    st.markdown(
        f"""<div class="score-card" style="background:{pos_color}22;border:2px solid {pos_color}">
            <div class="score-big" style="color:{pos_color}">{tm_e:.0f}%</div>
            <div class="score-lbl" style="color:{pos_color}">Marge brute</div>
            <div style="font-size:.85rem;margin-top:.3rem">Médiane secteur : {tm_med:.0f}% [{tm_low:.0f}–{tm_high:.0f}%]</div>
        </div>""", unsafe_allow_html=True
    )

st.divider()

# ══════════════════════════════════════════════════════════════
# ONGLETS DÉTAILLÉS
# ══════════════════════════════════════════════════════════════
tabs = st.tabs([
    "🔴 Anomalies",
    "📉 Marges & TVA",
    "🛡️ Bouclier Fiscal",
    "🏢 Benchmark Sectoriel",
    "📈 Prévisionnel N+1",
    "💡 Suggestions",
    "🤖 Rapport IA",
    "⬇️ Téléchargements",
])

# ── Onglet 1 : Anomalies ──────────────────────────────────────
with tabs[0]:
    st.subheader(f"Anomalies détectées : {len(all_anomalies)} ({nc} critique(s))")

    filtres_gravite = st.multiselect(
        "Filtrer par gravité",
        options=["CRITIQUE", "ALERTE", "ATTENTION", "INFO"],
        default=["CRITIQUE", "ALERTE", "ATTENTION", "INFO"],
    )

    filtrees = [a for a in all_anomalies if a['gravite'] in filtres_gravite]

    c2, c3, c4, c5 = st.columns(4)
    c2.metric("🔴 Critiques",  nc)
    c3.metric("🟠 Alertes",    na)
    c4.metric("🟡 Attention",  nat)
    c5.metric("🟢 Info",       ni)

    st.divider()
    afficher_anomalies(filtrees)

# ── Onglet 2 : Marges & TVA ──────────────────────────────────
with tabs[1]:
    st.subheader("Marges & TVA")
    col1, col2 = st.columns(2)
    with col1:
        ca = mg.get('ca', 0)
        achats = mg.get('achats', 0)
        mb = mg.get('mb', 0)
        tm = mg.get('tm', 0)
        cp = mg.get('cp', 0)
        ce = mg.get('ce', 0)
        tp = mg.get('tp', 0)
        tc_val = mg.get('tc', 0)
        res = mg.get('res', 0)
        tr = mg.get('tr', 0)

        rows = [
            ("Chiffre d'affaires (70x)", f"{ca:,.2f} €", "normal"),
            ("Achats (60x)",             f"{achats:,.2f} €", "normal"),
            ("Marge brute",              f"{mb:,.2f} €", "positive" if mb >= 0 else "negative"),
            ("Taux marge brute",         f"{tm:.2f}%", "positive" if tm > 0 else "negative"),
            ("Charges externes (61+62)", f"{ce:,.2f} €", "normal"),
            ("Charges de personnel (64x)",f"{cp:,.2f} €", "normal"),
            ("Total produits (7xx)",     f"{tp:,.2f} €", "normal"),
            ("Total charges (6xx)",      f"{tc_val:,.2f} €", "normal"),
            ("Résultat d'exploitation",  f"{res:,.2f} €", "positive" if res >= 0 else "negative"),
            ("Taux de résultat",         f"{tr:.2f}%", "normal"),
        ]
        for label, val, kind in rows:
            color = "#22c55e" if kind == "positive" else "#ef4444" if kind == "negative" else ""
            style = f"color:{color};font-weight:bold" if color else ""
            st.markdown(
                f'<div style="display:flex;justify-content:space-between;padding:.35rem 0;'
                f'border-bottom:1px solid #eee"><span>{label}</span>'
                f'<span style="{style}">{val}</span></div>',
                unsafe_allow_html=True,
            )

    with col2:
        tva_c = mg.get('tva_c', 0)
        tva_d = mg.get('tva_d', 0)
        tx_tva = mg.get('tx_tva', 0)
        st.markdown("#### TVA")
        for label, val, kind in [
            ("TVA collectée (4457x)", f"{tva_c:,.2f} €", "normal"),
            ("TVA déductible (4456x)", f"{tva_d:,.2f} €", "normal"),
            ("Taux TVA apparent", f"{tx_tva:.2f}%",
             "negative" if tx_tva > 20 else "positive" if tva_c > 0 else "normal"),
        ]:
            color = "#22c55e" if kind == "positive" else "#ef4444" if kind == "negative" else ""
            style = f"color:{color};font-weight:bold" if color else ""
            st.markdown(
                f'<div style="display:flex;justify-content:space-between;padding:.35rem 0;'
                f'border-bottom:1px solid #eee"><span>{label}</span>'
                f'<span style="{style}">{val}</span></div>',
                unsafe_allow_html=True,
            )

        st.markdown("#### Anomalies Marges & TVA")
        marge_anom = [a for a in all_anomalies
                      if any(x in a['type'] for x in ['Marge', 'Personnel', 'TVA'])]
        afficher_anomalies(marge_anom, "Aucune anomalie sur les marges/TVA ✅")

# ── Onglet 3 : Bouclier Fiscal ────────────────────────────────
with tabs[2]:
    st.subheader("🛡️ Bouclier Fiscal — Score de Probabilité de Contrôle DGFiP")
    if not bouclier:
        st.warning("Données du bouclier fiscal non disponibles.")
    else:
        conseil = bouclier.get('conseil', '')
        details = bouclier.get('details', {})

        col_g, col_d = st.columns([1, 2])
        with col_g:
            st.markdown(
                f"""<div class="score-card" style="background:{b_color}22;border:2px solid {b_color}">
                    <div class="score-big" style="color:{b_color}">{b_score}/100</div>
                    <div class="score-lbl" style="color:{b_color}">{b_niveau}</div>
                    <div style="font-size:.85rem;margin-top:.5rem;color:#555">{conseil}</div>
                </div>""", unsafe_allow_html=True
            )

        with col_d:
            st.markdown("**Détail par axe**")
            for key, info in details.items():
                pts = info.get('points', 0); mx = info.get('max', 1)
                lbl = info.get('label', key); det = info.get('detail', '')
                pct = pts / mx * 100
                color = '#ef4444' if pct >= 80 else '#f97316' if pct >= 50 else '#eab308' if pct >= 25 else '#22c55e'
                st.markdown(
                    f"""<div style="margin-bottom:.6rem">
                        <div style="display:flex;justify-content:space-between">
                            <b>{lbl}</b>
                            <span style="color:{color};font-weight:700">{pts}/{mx}</span>
                        </div>
                        <div style="background:#e5e7eb;border-radius:4px;height:8px;margin:.3rem 0">
                            <div style="width:{pct:.0f}%;background:{color};height:8px;border-radius:4px"></div>
                        </div>
                        <small style="color:#6b7280">{det}</small>
                    </div>""", unsafe_allow_html=True
                )

# ── Onglet 4 : Benchmark Sectoriel ───────────────────────────
with tabs[3]:
    st.subheader(f"🏢 Benchmark Sectoriel — {BENCHMARKS[secteur_code]['label']}")
    bench = stats.get('benchmark', {})
    sal   = stats.get('charges_salariales', {})
    fisc  = stats.get('scoring_fiscal', {})

    if not bench:
        st.warning("Données benchmark non disponibles.")
    else:
        note = BENCHMARKS[secteur_code].get('note', '')
        st.info(f"💡 {note}")

        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.markdown("**Comparaison marges**")
            tm_e   = bench.get('tm_entreprise', 0) or 0
            tm_med = bench.get('tm_mediane', bench.get('tm_med', 0)) or 0
            tm_low = bench.get('tm_low', 0) or 0
            tm_high= bench.get('tm_high', 0) or 0
            ecart  = tm_e - tm_med
            in_range = tm_low <= tm_e <= tm_high
            c_e = "#22c55e" if in_range else "#ef4444"
            for label, val, color in [
                ("Marge entreprise", f"{tm_e:.1f}%", c_e),
                ("Médiane secteur",  f"{tm_med:.0f}%", "#6b7280"),
                ("Fourchette secteur", f"[{tm_low:.0f}% — {tm_high:.0f}%]", "#6b7280"),
                ("Écart vs médiane", f"{'+'if ecart>=0 else ''}{ecart:.1f} pts", c_e),
            ]:
                st.markdown(
                    f'<div style="display:flex;justify-content:space-between;padding:.3rem 0;'
                    f'border-bottom:1px solid #eee"><span>{label}</span>'
                    f'<span style="color:{color};font-weight:bold">{val}</span></div>',
                    unsafe_allow_html=True
                )

        with col_b2:
            st.markdown("**Charges de personnel & Fiscal**")
            cp_e = bench.get('cp_ca_entreprise', 0) or 0
            cp_med = bench.get('cp_ca_mediane', 0) or 0
            ecart_cp = cp_e - cp_med
            c_cp = "#ef4444" if abs(ecart_cp) > 15 else "#22c55e"
            for label, val, color in [
                ("Personnel/CA entreprise", f"{cp_e:.1f}%", c_cp),
                ("Personnel/CA secteur",   f"{cp_med:.0f}%", "#6b7280"),
                ("Écart charges pers.",    f"{'+'if ecart_cp>=0 else ''}{ecart_cp:.1f} pts", c_cp),
                ("Salaires bruts",         f"{sal.get('sal_bruts',0):,.0f} €", ""),
                ("Cotisations patronales", f"{sal.get('cotis_pat',0):,.0f} €", ""),
                ("Ratio cotis./salaires",  f"{sal.get('ratio_cotis',0):.1f}%",
                 "#ef4444" if not (20 <= sal.get('ratio_cotis', 0) <= 60) else "#22c55e"),
                ("Effectif estimé",        f"~{sal.get('nb_sal_estime',0)} salarié(s)", ""),
                ("IS comptabilisé",        f"{fisc.get('is_d',0):,.0f} €", ""),
                ("Taux IS apparent",       f"{fisc.get('tx_is',0):.1f}%",
                 "#ef4444" if fisc.get('tx_is', 0) > 35 else ""),
            ]:
                st.markdown(
                    f'<div style="display:flex;justify-content:space-between;padding:.3rem 0;'
                    f'border-bottom:1px solid #eee"><span>{label}</span>'
                    f'<span style="color:{color};font-weight:bold">{val}</span></div>',
                    unsafe_allow_html=True
                )

        st.markdown("**Anomalies benchmark**")
        bench_anom = [a for a in all_anomalies
                      if any(x in a['type'].lower() for x in ['mediane', 'sectoriel', 'cotisations', 'inattendue'])]
        afficher_anomalies(bench_anom, "Aucune anomalie sectorielle ✅")

# ── Onglet 5 : Prévisionnel N+1 ──────────────────────────────
with tabs[4]:
    st.subheader("📈 Prévisionnel N+1")
    if not activer_prev:
        st.info("Le module Prévisionnel est désactivé. Activez-le dans la barre latérale.")
    elif not prev_data:
        st.warning("Le prévisionnel n'a pas pu être généré (CA non détecté ?).")
    else:
        an = str(prev_data.get('annee_n', 'N'))
        try:
            an1 = str(int(an) + 1)
        except:
            an1 = 'N+1'

        st.markdown(f"**Hypothèse de croissance CA : +{croissance_ca}%** | Taux IS : {prev_data.get('taux_is',0.25)*100:.0f}%")

        import pandas as pd
        prev_rows = [
            ("Chiffre d'affaires", prev_data.get('ca_n', 0), prev_data.get('ca_n1', 0)),
            ("Achats",             prev_data.get('achats_n', 0), prev_data.get('achats_n1', 0)),
            ("Marge brute",        prev_data.get('ca_n', 0) - prev_data.get('achats_n', 0), prev_data.get('mb_n1', 0)),
            ("Charges externes",   prev_data.get('ce_n', 0), prev_data.get('ce_n1', 0)),
            ("Charges personnel",  prev_data.get('cp_n', 0), prev_data.get('cp_n1', 0)),
            ("EBE",                None, prev_data.get('ebe_n1', 0)),
            ("Amortissements",     prev_data.get('dot_amo_n', 0), prev_data.get('dot_amo_n1', 0)),
            ("Résultat exploitation", prev_data.get('res_n', 0), prev_data.get('rex_n1', 0)),
            ("IS",                 None, prev_data.get('is_n1', 0)),
            ("Résultat net",       prev_data.get('res_n', 0), prev_data.get('res_net_n1', 0)),
        ]

        def fmt(v):
            if v is None: return "—"
            return f"{v:,.0f} €"

        def variation(n, n1):
            if n is None or n == 0: return "—"
            v = n1 - n; p = v / abs(n) * 100
            return f"{'+'if v>=0 else ''}{v:,.0f} ({p:+.1f}%)"

        table_data = {
            "Poste": [r[0] for r in prev_rows],
            f"{an} (réel)": [fmt(r[1]) for r in prev_rows],
            f"{an1} (prévis.)": [fmt(r[2]) for r in prev_rows],
            "Variation": [variation(r[1], r[2]) for r in prev_rows],
        }
        st.dataframe(pd.DataFrame(table_data), use_container_width=True, hide_index=True)

        col_r1, col_r2, col_r3 = st.columns(3)
        col_r1.metric(f"Taux marge brute {an1}", f"{prev_data.get('tm_n1', 0):.1f}%")
        col_r2.metric(f"Taux résultat net {an1}", f"{prev_data.get('tr_n1', 0):.1f}%")
        col_r3.metric(f"Charges pers./CA {an1}", f"{prev_data.get('cp_ca_n1', 0):.1f}%")

# ── Onglet 6 : Suggestions ───────────────────────────────────
with tabs[5]:
    st.subheader("💡 Commentaires & Suggestions")

    st.markdown("#### Commentaires automatiques")
    for com in commentaires:
        color_map = {'red': '🔴', 'orange': '🟠', 'green': '🟢', 'gray': '⚪'}
        icon = color_map.get(com.get('couleur', 'gray'), '⚪')
        bg = {'red': '#fef2f2', 'orange': '#fff7ed', 'green': '#f0fdf4', 'gray': '#f9fafb'}.get(com.get('couleur', 'gray'), '#f9fafb')
        border = {'red': '#ef4444', 'orange': '#f97316', 'green': '#22c55e', 'gray': '#9ca3af'}.get(com.get('couleur', 'gray'), '#9ca3af')
        st.markdown(
            f"""<div style="background:{bg};border-left:4px solid {border};
                padding:.6rem 1rem;border-radius:6px;margin-bottom:.5rem">
                {icon} <b>{com.get('titre', '')}</b><br>
                <span style="font-size:.9rem">{com.get('texte', '')}</span>
            </div>""",
            unsafe_allow_html=True
        )

    if suggestions:
        st.markdown("#### Recommandations")
        for i, s in enumerate(suggestions, 1):
            st.markdown(f"**{i}.** {s}")

# ── Onglet 7 : Rapport IA ────────────────────────────────────
with tabs[6]:
    st.subheader("🤖 Rapport de Mission IA (Module 1)")

    if not activer_ia:
        st.info("Activez le module Rapport IA dans la barre latérale et saisissez votre clé API Anthropic.")
    elif not api_key_input:
        st.warning("Veuillez saisir votre clé API Anthropic dans la barre latérale.")
    elif not ANTHROPIC_AVAILABLE:
        st.error("Package `anthropic` non installé. Ajoutez-le dans requirements.txt.")
    else:
        if st.button("✍️ Générer le rapport de mission", type="primary"):
            with st.spinner("Rédaction du rapport en cours (Claude)…"):
                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_docx:
                    tmp_docx_path = tmp_docx.name

                rapport_path, rapport_texte = generer_rapport_mission_ia(
                    uploaded.name, stats, all_anomalies, secteur_code,
                    api_key=api_key_input,
                    output_docx=tmp_docx_path,
                )

            if rapport_texte and not rapport_texte.startswith("MODULE"):
                st.success("✅ Rapport généré avec succès !")
                with st.expander("📄 Aperçu du rapport", expanded=True):
                    st.markdown(rapport_texte)

                # Téléchargement Word si disponible
                if rapport_path and os.path.exists(rapport_path):
                    with open(rapport_path, 'rb') as f:
                        st.download_button(
                            "⬇️ Télécharger le rapport Word (.docx)",
                            data=f.read(),
                            file_name=f"rapport_mission_{os.path.splitext(uploaded.name)[0]}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                    os.unlink(rapport_path)
                else:
                    st.download_button(
                        "⬇️ Télécharger le rapport (texte)",
                        data=rapport_texte.encode('utf-8'),
                        file_name=f"rapport_mission_{os.path.splitext(uploaded.name)[0]}.txt",
                        mime="text/plain",
                    )
            else:
                st.error(f"Erreur : {rapport_texte}")

# ── Onglet 8 : Téléchargements ────────────────────────────────
with tabs[7]:
    st.subheader("⬇️ Téléchargements")

    st.markdown("#### Rapport HTML complet")
    st.download_button(
        label="⬇️ Télécharger le rapport HTML",
        data=html_rapport.encode('utf-8'),
        file_name=f"{os.path.splitext(uploaded.name)[0]}_audit_v5.html",
        mime="text/html",
        type="primary",
    )
    st.caption("Rapport interactif complet avec graphiques Chart.js — à ouvrir dans un navigateur.")

    st.divider()

    st.markdown("#### Données JSON brutes")
    import json
    stats_export = {k: v for k, v in stats.items() if k != '_anomalies_raw'}
    st.download_button(
        label="⬇️ Télécharger les statistiques (JSON)",
        data=json.dumps(stats_export, ensure_ascii=False, indent=2, default=str).encode('utf-8'),
        file_name=f"{os.path.splitext(uploaded.name)[0]}_stats.json",
        mime="application/json",
    )

    st.divider()

    st.markdown("#### Liste des anomalies (CSV)")
    import pandas as pd
    if all_anomalies:
        df_anom = pd.DataFrame(all_anomalies)
        st.download_button(
            label="⬇️ Télécharger les anomalies (CSV)",
            data=df_anom.to_csv(index=False, sep=';', encoding='utf-8-sig'),
            file_name=f"{os.path.splitext(uploaded.name)[0]}_anomalies.csv",
            mime="text/csv",
        )
    else:
        st.info("Aucune anomalie à exporter.")

# ══════════════════════════════════════════════════════════════
# FOOTER
# ══════════════════════════════════════════════════════════════
st.divider()
st.caption(
    "FEC Audit Pro v5.0 — Moteur d'analyse basé sur les référentiels DGFiP et Banque de France. "
    "Document généré à titre d'aide à la révision — à valider par un expert-comptable."
)
