#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fec_patch_v6.py — Correctifs et améliorations pour fec_audit_pro_v5.py

Améliorations apportées :
  1. Détection de secteur refaite (e-commerce, marketplace, signaux fournisseurs/clients)
  2. Bouclier fiscal élargi : 12 axes au lieu de 8, seuils recalibrés
  3. Nouvelle fonction : analyser_ecommerce() — contrôles spécifiques marketplace
  4. Nouveaux benchmarks : ECOMMERCE_MARKETPLACE, RESTAURATION enrichi
  5. SECTEURS_LISTE mise à jour avec le nouveau code

Import dans app.py :
    import fec_patch_v6 as _patch
    _patch.appliquer(fec_audit_pro_v5_module)
"""

import pandas as pd
import numpy as np
import math
from datetime import datetime
from collections import Counter


# ══════════════════════════════════════════════════════════════
# 1. NOUVEAUX BENCHMARKS
# ══════════════════════════════════════════════════════════════

BENCHMARKS_SUPPLEMENTAIRES = {
    'ECOMMERCE_MARKETPLACE': {
        'label': 'E-commerce / Marketplace (Amazon, Cdiscount…)',
        'tm_med': 28.0, 'tm_low': 18.0, 'tm_high': 42.0,
        'tr_med': 3.5,  'cp_ca': 6.0,   'tx_tva': 20.0,
        'ratio_frais_envoi_ca_max': 8.0,
        'ratio_commission_ca_max': 12.0,
        'concentration_client_seuil': 60.0,
        'note': (
            "Vente en ligne via marketplace (Amazon FBA/FBM, Cdiscount, Fnac). "
            "Marge brute 20-42% selon catégorie produit. "
            "Commissions marketplace (622) typiquement 8-15% du CA. "
            "Frais d'envoi (626/La Poste) < 8% du CA — au-delà c'est anormal. "
            "Masse salariale très faible si activité solo/autoentrepreneur-like. "
            "Concentration client unique (Amazon) : CRITIQUE si > 90% du CA — "
            "dépendance totale à la plateforme, risque de suspension de compte."
        ),
    },
    'ECOMMERCE_PROPRE': {
        'label': 'E-commerce site propre / BtoC en ligne',
        'tm_med': 45.0, 'tm_low': 32.0, 'tm_high': 62.0,
        'tr_med': 5.0,  'cp_ca': 12.0,  'tx_tva': 20.0,
        'note': (
            "Vente en ligne via site propre (Shopify, Prestashop…). "
            "Marges supérieures à la marketplace car pas de commissions plateforme. "
            "Frais marketing (623) peuvent être élevés (acquisition client). "
            "Coût logistique en 624/626 à surveiller."
        ),
    },
}

SECTEURS_LISTE_COMPLET = [
    ('1',  'RESTAURATION',           'Restauration / Hotellerie / Traiteur'),
    ('2',  'COMMERCE_DETAIL',        'Commerce de detail (magasin physique)'),
    ('3',  'COMMERCE_GROS',          'Commerce de gros / Negoce / Distribution'),
    ('4',  'BTP',                    'BTP / Construction / Renovation'),
    ('5',  'SERVICES_B2B',           'Conseil / Services aux entreprises / Cabinet'),
    ('6',  'PROFESSION_LIBERALE',    'Profession liberale / Sante / Juridique'),
    ('7',  'INDUSTRIE',              'Industrie / Fabrication / Transformation'),
    ('8',  'TRANSPORT',              'Transport / Logistique / Livraison'),
    ('9',  'IMMOBILIER',             'Immobilier / Agence immobiliere'),
    ('10', 'INFORMATIQUE',           'Informatique / ESN / Developpement / SaaS'),
    ('11', 'ECOMMERCE_MARKETPLACE',  'E-commerce Marketplace (Amazon, Cdiscount…)'),
    ('12', 'ECOMMERCE_PROPRE',       'E-commerce site propre / BtoC en ligne'),
    ('0',  'INCONNU',                'Autre / Ne pas comparer au marche'),
]


# ══════════════════════════════════════════════════════════════
# 2. DÉTECTION DE SECTEUR AMÉLIORÉE
# ══════════════════════════════════════════════════════════════

def detecter_secteur_v2(df):
    """
    Détection de secteur v2 — signaux forts fournisseurs/clients + ratios.
    Priorité aux signaux non ambigus (Amazon client dominant, La Poste massif…).
    """
    scores = {}
    from fec_audit_pro_v5 import BENCHMARKS
    all_benchmarks = {**BENCHMARKS, **BENCHMARKS_SUPPLEMENTAIRES}
    for k in all_benchmarks:
        if k != 'INCONNU':
            scores[k] = 0

    def s(pfx, col='Debit'):
        mask = df['CompteNum'].str.startswith(pfx)
        return float(df.loc[mask, col].sum())

    ca_total  = s('70', 'Credit') - s('70', 'Debit')
    ach_total = s('60', 'Debit')  - s('60', 'Credit')
    cp_total  = s('64', 'Debit')
    ce_total  = s('61', 'Debit') + s('62', 'Debit')
    tm_brute  = (ca_total - ach_total) / ca_total * 100 if ca_total > 0 else 50

    libs_raw = (
        ' '.join(df.get('EcritureLib', pd.Series(dtype=str)).fillna('').tolist()) + ' ' +
        ' '.join(df.get('CompteLib',   pd.Series(dtype=str)).fillna('').tolist()) + ' ' +
        ' '.join(df.get('CompAuxLib',  pd.Series(dtype=str)).fillna('').tolist())
    ).upper()

    # ── Signaux FORTS fournisseurs/clients ────────────────────

    # 1. Amazon comme CLIENT dominant
    clients_df = df[df['CompteNum'].str.startswith('411')]
    ca_amazon_client = 0.0
    if not clients_df.empty:
        mask_amazon = clients_df['CompteLib'].str.upper().str.contains('AMAZON', na=False)
        ca_amazon_client = float(clients_df.loc[mask_amazon, 'Credit'].sum())
    ratio_amazon_ca = ca_amazon_client / ca_total if ca_total > 0 else 0

    if ratio_amazon_ca > 0.5:
        scores['ECOMMERCE_MARKETPLACE'] += 30
        if ratio_amazon_ca > 0.85:
            scores['ECOMMERCE_MARKETPLACE'] += 20  # signal ultra-fort

    # 2. Commissions Amazon en 622
    comm_622 = df[df['CompteNum'].str.startswith('622')]
    comm_amazon = 0.0
    if not comm_622.empty:
        mask_comm = comm_622['EcritureLib'].str.upper().str.contains('AMAZON', na=False)
        comm_amazon = float(comm_622.loc[mask_comm, 'Debit'].sum())
    if comm_amazon > 0:
        scores['ECOMMERCE_MARKETPLACE'] += 15

    # 3. La Poste / Colissimo / Chronopost en 626
    telecom_df = df[df['CompteNum'].str.startswith('626')]
    laposte_montant = 0.0
    if not telecom_df.empty:
        kw_envoi = ['POSTE', 'COLISSIMO', 'CHRONOPOST', 'DHL', 'UPS', 'FEDEX', 'MONDIAL RELAY']
        mask_envoi = telecom_df['EcritureLib'].str.upper().str.contains(
            '|'.join(kw_envoi), na=False)
        laposte_montant = float(telecom_df.loc[mask_envoi, 'Debit'].sum())
    ratio_envoi_ca = laposte_montant / ca_total if ca_total > 0 else 0
    if ratio_envoi_ca > 0.05:
        scores['ECOMMERCE_MARKETPLACE'] += 12
        scores['ECOMMERCE_PROPRE']      += 6
    if ratio_envoi_ca > 0.12:
        scores['ECOMMERCE_MARKETPLACE'] += 8

    # 4. Stocks (37x) + achats marchandises (607) = revendeur
    stock_37 = s('37', 'Debit')
    ach_marc  = s('607', 'Debit')
    if ach_marc > 0 and ca_total > 0:
        scores['ECOMMERCE_MARKETPLACE'] += 5
        scores['ECOMMERCE_PROPRE']      += 5
        scores['COMMERCE_DETAIL']       += 4
        scores['COMMERCE_GROS']         += 3

    # 5. Hébergement web / Shopify / Prestashop → site propre
    kw_site = ['SHOPIFY', 'PRESTASHOP', 'WOOCOMMERCE', 'HOSTINGER',
               'HEBERGEMENT WEB', 'NOM DE DOMAINE', 'STRIPE']
    for kw in kw_site:
        if kw in libs_raw:
            scores['ECOMMERCE_PROPRE'] += 8

    # 6. FBM/FBA Amazon ou "AMAZON" en fournisseur (achats pour revente)
    four_df = df[df['CompteNum'].str.startswith('401')]
    if not four_df.empty:
        mask_amz_four = four_df['CompteLib'].str.upper().str.contains('AMAZON', na=False)
        amz_achat = float(four_df.loc[mask_amz_four, 'Debit'].sum())
        if amz_achat > 0 and comm_amazon > 0:
            scores['ECOMMERCE_MARKETPLACE'] += 10

    # ── Signaux sectoriels classiques ────────────────────────

    kw_map = {
        'RESTAURATION':        ['RESTAURANT','REPAS','TRAITEUR','HOTEL','BRASSERIE',
                                 'CUISINE','UBER EATS','DELIVEROO','MENU'],
        'BTP':                 ['CHANTIER','TRAVAUX','BATIMENT','CONSTRUCTION','RENOVATION',
                                 'MACONNERIE','ELECTRICITE','PLOMBERIE','COUVERTURE'],
        'SERVICES_B2B':        ['HONORAIRES','CONSEIL','CONSULTING','MISSION','FORMATION',
                                 'AUDIT','EXPERTISE','ASSISTANCE','MANAGEMENT'],
        'INFORMATIQUE':        ['LICENCE','SAAS','CLOUD','SERVEUR','INFRA',
                                 'DEVELOPPEMENT','MAINTENANCE INFORMATIQUE'],
        'PROFESSION_LIBERALE': ['CABINET MEDICAL','CONSULTATION','ORDONNANCE','CLINIQUE',
                                 'KINESITHERAPIE','DENTISTE','MEDECIN'],
        'IMMOBILIER':          ['COMMISSION AGENCE','MANDAT','TRANSACTION IMMO','BAIL'],
        'TRANSPORT':           ['GASOIL','PEAGE','FLOTTE','FRET','EXPEDITION','CHAUFFEUR'],
    }
    for sect, kws in kw_map.items():
        hits = sum(1 for k in kws if k in libs_raw)
        scores[sect] += hits * 3

    # ── Signaux financiers (ratios) ───────────────────────────
    if tm_brute > 80:
        scores['SERVICES_B2B']        += 4
        scores['INFORMATIQUE']        += 3
        scores['PROFESSION_LIBERALE'] += 3
    elif 18 <= tm_brute <= 55:
        scores['COMMERCE_DETAIL']     += 3
        scores['COMMERCE_GROS']       += 2
        scores['ECOMMERCE_MARKETPLACE'] += 2

    if ca_total > 0:
        ratio_cp = cp_total / ca_total * 100
        if ratio_cp > 50:
            scores['INFORMATIQUE']  += 3
            scores['SERVICES_B2B']  += 2
        if ratio_cp < 5:
            scores['ECOMMERCE_MARKETPLACE'] += 3
            scores['ECOMMERCE_PROPRE']      += 2

    cpts_2 = df['CompteNum'].str[:2].value_counts()
    if '86' in cpts_2.index or '69' in cpts_2.index:
        scores['PROFESSION_LIBERALE'] += 5

    # ── Sélection ────────────────────────────────────────────
    # Filtrer les scores nuls
    valid = {k: v for k, v in scores.items() if v > 0}
    if not valid:
        return 'INCONNU', 0, scores

    best  = max(valid, key=lambda k: valid[k])
    total = sum(valid.values())
    conf  = round(valid[best] / total * 100) if total > 0 else 0

    return best, conf, scores


# ══════════════════════════════════════════════════════════════
# 3. ANALYSE SPÉCIFIQUE E-COMMERCE
# ══════════════════════════════════════════════════════════════

def analyser_ecommerce(df):
    """
    Contrôles spécifiques aux activités e-commerce / marketplace.
    Retourne (anomalies, stats) comme tous les autres analyseurs.
    """
    anomalies = []
    stats = {}

    def s(pfx, col='Debit'):
        return float(df[df['CompteNum'].str.startswith(pfx)][col].sum())

    ca_total = s('70', 'Credit') - s('70', 'Debit')
    if ca_total <= 0:
        return anomalies, stats

    # ── 1. Concentration client ────────────────────────────────
    clients = df[df['CompteNum'].str.startswith('411')]
    if not clients.empty:
        grp = clients.groupby('CompteLib')['Credit'].sum().sort_values(ascending=False)
        top1_lib = grp.index[0] if len(grp) else '—'
        top1_val = float(grp.iloc[0]) if len(grp) else 0
        ratio_top1 = top1_val / ca_total * 100

        stats['concentration_client_top1'] = ratio_top1
        stats['client_top1'] = top1_lib

        if ratio_top1 > 95:
            anomalies.append({
                'type': 'Concentration client EXTRÊME',
                'gravite': 'CRITIQUE',
                'detail': (
                    f"{top1_lib} représente {ratio_top1:.1f}% du CA ({top1_val:,.0f} €). "
                    "Dépendance totale à une seule plateforme : risque de suspension = "
                    "faillite immédiate. La DGFiP considère ce profil comme à très haut "
                    "risque (critère de fragilité structurelle)."
                ),
                'montant': top1_val,
                'reference': '411xx — comptes clients',
            })
        elif ratio_top1 > 75:
            anomalies.append({
                'type': 'Concentration client critique',
                'gravite': 'ALERTE',
                'detail': (
                    f"{top1_lib} représente {ratio_top1:.1f}% du CA. "
                    "Diversification insuffisante — risque opérationnel et financier majeur."
                ),
                'montant': top1_val,
                'reference': '411xx',
            })

    # ── 2. Frais d'envoi disproportionnés ─────────────────────
    tel_df = df[df['CompteNum'].str.startswith('626')]
    kw_envoi = ['POSTE', 'COLISSIMO', 'CHRONOPOST', 'DHL', 'UPS', 'FEDEX', 'MONDIAL']
    if not tel_df.empty:
        mask = tel_df['EcritureLib'].str.upper().str.contains('|'.join(kw_envoi), na=False)
        frais_envoi = float(tel_df.loc[mask, 'Debit'].sum())
        ratio_envoi = frais_envoi / ca_total * 100
        stats['frais_envoi'] = frais_envoi
        stats['ratio_frais_envoi_ca'] = ratio_envoi

        if ratio_envoi > 20:
            anomalies.append({
                'type': 'Frais d\'envoi excessifs',
                'gravite': 'CRITIQUE',
                'detail': (
                    f"Frais d'envoi (La Poste/transporteurs) = {frais_envoi:,.0f} € "
                    f"soit {ratio_envoi:.1f}% du CA. "
                    "Seuil critique : > 15%. Possible manque de négociation tarifaire, "
                    "frais privés refacturés à l'entreprise, ou volume fictif d'expéditions."
                ),
                'montant': frais_envoi,
                'reference': '626xx — frais postaux',
            })
        elif ratio_envoi > 12:
            anomalies.append({
                'type': 'Frais d\'envoi élevés',
                'gravite': 'ALERTE',
                'detail': (
                    f"Frais d'envoi = {frais_envoi:,.0f} € ({ratio_envoi:.1f}% du CA). "
                    "Normale marketplace : 4-10%. Justifier les conditions tarifaires La Poste."
                ),
                'montant': frais_envoi,
                'reference': '626xx',
            })
        elif ratio_envoi > 8:
            anomalies.append({
                'type': 'Frais d\'envoi à surveiller',
                'gravite': 'ATTENTION',
                'detail': (
                    f"Frais d'envoi = {frais_envoi:,.0f} € ({ratio_envoi:.1f}% du CA). "
                    "Légèrement au-dessus de la norme marketplace (4-8%)."
                ),
                'montant': frais_envoi,
                'reference': '626xx',
            })

    # ── 3. Commissions marketplace vs CA ──────────────────────
    comm_df = df[df['CompteNum'].str.startswith('622')]
    if not comm_df.empty:
        commissions = float(comm_df['Debit'].sum())
        ratio_comm = commissions / ca_total * 100
        stats['commissions_marketplace'] = commissions
        stats['ratio_comm_ca'] = ratio_comm

        if ratio_comm > 18:
            anomalies.append({
                'type': 'Commissions marketplace anormalement élevées',
                'gravite': 'CRITIQUE',
                'detail': (
                    f"Commissions (622) = {commissions:,.0f} € soit {ratio_comm:.1f}% du CA. "
                    "Les commissions Amazon varient de 8-15% selon catégorie. "
                    "Au-delà de 18% : vérifier les remboursements comptabilisés en charges, "
                    "les frais FBA, ou des charges personnelles refacturées."
                ),
                'montant': commissions,
                'reference': '622xx — commissions et courtage',
            })
        elif ratio_comm > 14:
            anomalies.append({
                'type': 'Commissions marketplace élevées',
                'gravite': 'ALERTE',
                'detail': (
                    f"Commissions = {commissions:,.0f} € ({ratio_comm:.1f}% du CA). "
                    "Norme Amazon : 8-15%. Détailler la composition des frais."
                ),
                'montant': commissions,
                'reference': '622xx',
            })

    # ── 4. Stocks non apurés (37x au bilan) ───────────────────
    stocks_d = s('37', 'Debit')
    stocks_c = s('37', 'Credit')
    solde_stock = stocks_d - stocks_c
    stats['solde_stocks'] = solde_stock
    if solde_stock > 0 and ca_total > 0:
        ratio_stock = solde_stock / ca_total * 100
        if ratio_stock > 30:
            anomalies.append({
                'type': 'Stock marchandises excessif au bilan',
                'gravite': 'ALERTE',
                'detail': (
                    f"Stocks (37x) = {solde_stock:,.0f} € soit {ratio_stock:.1f}% du CA. "
                    "Risque de dépréciation ou de stocks fictifs. "
                    "Inventaire physique à justifier (article 13 du CGI)."
                ),
                'montant': solde_stock,
                'reference': '37xx — stocks marchandises',
            })

    # ── 5. Achats Amazon (fournisseur) sans séparation FBA ────
    four_amz_df = df[df['CompteNum'].str.startswith('401')]
    if not four_amz_df.empty:
        mask_amz = four_amz_df['CompteLib'].str.upper().str.contains('AMAZON', na=False)
        amz_achat = float(four_amz_df.loc[mask_amz, 'Debit'].sum())
        if amz_achat > 0 and 'commissions_marketplace' in stats:
            ratio_achat_comm = amz_achat / stats['commissions_marketplace'] if stats['commissions_marketplace'] > 0 else 0
            stats['achats_amazon_fournisseur'] = amz_achat
            if amz_achat > stats['commissions_marketplace'] * 0.5:
                anomalies.append({
                    'type': 'Flux Amazon fournisseur/client mélangés',
                    'gravite': 'ATTENTION',
                    'detail': (
                        f"Achats chez Amazon (401) = {amz_achat:,.0f} € ET "
                        f"commissions Amazon (622) = {stats['commissions_marketplace']:,.0f} €. "
                        "Vérifier la séparation achats de marchandises / frais de plateforme. "
                        "Les commissions FBA ne doivent pas être en 607."
                    ),
                    'montant': amz_achat,
                    'reference': '401AMAZON / 622xx',
                })

    # ── 6. TVA collectée vs CA marketplace ────────────────────
    tva_c = float(df[df['CompteNum'].str.startswith('44571')]['Credit'].sum())
    if tva_c > 0 and ca_total > 0:
        tx_tva_apparent = tva_c / ca_total * 100
        stats['tx_tva_apparent_marketplace'] = tx_tva_apparent
        # Amazon collecte la TVA depuis 2021 pour les vendeurs (représentant fiscal)
        if tx_tva_apparent < 5 and ca_total > 50000:
            anomalies.append({
                'type': 'TVA collectée anormalement faible (marketplace)',
                'gravite': 'ALERTE',
                'detail': (
                    f"TVA collectée = {tva_c:,.0f} € soit {tx_tva_apparent:.1f}% du CA. "
                    "Depuis 2021, Amazon collecte la TVA pour les vendeurs tiers. "
                    "Vérifier que le chiffre d'affaires Amazon est bien déclaré TTC "
                    "et que la TVA est correctement comptabilisée en 4457x."
                ),
                'montant': tva_c,
                'reference': '4457x — TVA collectée',
            })

    # ── 7. Dépenses mixtes (privé/pro) ────────────────────────
    kw_perso = {
        'RESTAURANT': ('62512', 'repas/restaurants', 3.0),
        'APPLE':      ('401',   'achats Apple',      2.0),
        'FNAC':       ('401',   'achats Fnac/Darty', 2.0),
        'DARTY':      ('401',   'achats Fnac/Darty', 2.0),
        'BRICO':      ('401',   'bricolage',         1.5),
        'IKEA':       ('401',   'IKEA',              1.5),
        'LECLERC':    ('401',   'grande surface',    1.5),
        'CARREFOR':   ('401',   'grande surface',    1.5),
        'PHARMACIE':  ('401',   'pharmacie',         1.0),
    }
    perso_total = 0.0
    perso_detail = []
    for kw, (pfx, lib, _seuil) in kw_perso.items():
        subset = df[df['CompteNum'].str.startswith(pfx)]
        mask   = subset['CompteLib'].str.upper().str.contains(kw, na=False)
        val    = float(subset.loc[mask, 'Debit'].sum())
        if val > 0:
            perso_total += val
            perso_detail.append(f"{lib} {val:,.0f} €")

    stats['charges_potentiellement_perso'] = perso_total
    if perso_total > 2000:
        anomalies.append({
            'type': 'Charges potentiellement personnelles',
            'gravite': 'ATTENTION',
            'detail': (
                f"Dépenses à caractère potentiellement personnel = {perso_total:,.0f} € : "
                f"{', '.join(perso_detail[:5])}. "
                "Ces achats (grande surface, pharmacie, IKEA, Apple…) doivent être "
                "justifiés par leur usage professionnel exclusif (article 39 CGI)."
            ),
            'montant': perso_total,
            'reference': 'Charges mixtes pro/perso',
        })
    elif perso_total > 500:
        anomalies.append({
            'type': 'Dépenses mixtes à justifier',
            'gravite': 'INFO',
            'detail': (
                f"Petites dépenses à valider ({perso_total:,.0f} €) : {', '.join(perso_detail)}."
            ),
            'montant': perso_total,
            'reference': 'Charges mixtes',
        })

    return anomalies, stats


# ══════════════════════════════════════════════════════════════
# 4. BOUCLIER FISCAL v2 — 12 axes, seuils recalibrés
# ══════════════════════════════════════════════════════════════

GRILLE_RISQUE_FISCAL_V2 = [
    (85, 'TRÈS ÉLEVÉ',   'Contrôle fiscal quasi-certain. Pré-audit recommandé en urgence. '
                          'Rassembler TOUS les justificatifs avant tout contact DGFiP.'),
    (65, 'ÉLEVÉ',        'Profil à haut risque DGFiP. Revue contradictoire urgente '
                          'des points identifiés. Envisager une procédure de régularisation.'),
    (45, 'MODÉRÉ',       'Plusieurs signaux détectés. Renforcer la documentation '
                          'et sécuriser les positions fiscales avant clôture.'),
    (25, 'FAIBLE',       'Risque limité. Maintenir la rigueur de saisie '
                          'et de justification des charges.'),
    ( 0, 'TRÈS FAIBLE',  'Aucun signal majeur. Dossier conforme aux normes DGFiP.'),
]


def calculer_bouclier_fiscal_v2(df, stats, anomalies, secteur_code):
    """
    Bouclier Fiscal v2 — 12 axes, seuils recalibrés, analyse des anomalies brutes.
    Score 0-100 : probabilité estimée de contrôle DGFiP.
    """
    from fec_audit_pro_v5 import BENCHMARKS
    all_bm = {**BENCHMARKS, **BENCHMARKS_SUPPLEMENTAIRES}

    bench = stats.get('benchmark', {})
    ben   = stats.get('benford', {})
    fe    = stats.get('fin_exercice', {})
    mg    = stats.get('marges', {})
    fisc  = stats.get('scoring_fiscal', {})
    sal   = stats.get('charges_salariales', {})
    ab    = stats.get('aberrants', {})
    dup   = stats.get('doublons', {})
    ec    = stats.get('ecommerce', {})  # si analyser_ecommerce a tourné
    details = {}

    def s(pfx, col='Debit'):
        return float(df[df['CompteNum'].str.startswith(pfx)][col].sum())

    ca_total = s('70', 'Credit') - s('70', 'Debit')

    # ── AXE 1 : Loi de Benford (20 pts) ──────────────────────
    chi2 = ben.get('chi2', 0)
    tot  = ben.get('total', 0)
    if tot >= 100:
        if   chi2 > 30: pts, lbl = 20, f"Distribution très anormale (χ²={chi2:.1f}) — manipulation probable"
        elif chi2 > 20: pts, lbl = 14, f"Distribution suspecte (χ²={chi2:.1f})"
        elif chi2 > 12: pts, lbl =  7, f"Légère déviation (χ²={chi2:.1f})"
        else:           pts, lbl =  0, f"Distribution conforme (χ²={chi2:.1f})"
    else:
        pts, lbl = 0, "Non applicable (< 100 montants)"
    details['benford'] = {'points': pts, 'max': 20, 'label': '📊 Loi de Benford', 'detail': lbl}

    # ── AXE 2 : Cohérence TVA (18 pts) ───────────────────────
    tva_c  = mg.get('tva_c', 0)
    tva_d  = mg.get('tva_d', 0)
    tx_tva = mg.get('tx_tva', 0)
    btva   = all_bm.get(secteur_code, {}).get('tx_tva', 20)
    pts = 0; lbl = "TVA cohérente"
    if tva_c > 0:
        r = tva_d / tva_c
        if   r > 2.5: pts, lbl = 18, f"Crédit TVA structurel très suspect (déductible={r:.1f}× collectée)"
        elif r > 1.8: pts, lbl = 12, f"TVA déductible élevée ({r:.1f}× la collectée)"
        elif r > 1.3: pts, lbl =  6, f"TVA déductible supérieure à la collectée ({r:.1f}×)"
        elif btva == 0 and tx_tva > 1: pts, lbl = 14, "TVA inattendue (secteur normalement exonéré)"
        elif btva > 0 and abs(tx_tva - btva) > 10:
            pts, lbl = 9, f"Taux TVA apparent {tx_tva:.1f}% éloigné du taux sectoriel {btva:.0f}%"
    elif tva_c == 0 and ca_total > 50000 and btva > 0:
        pts, lbl = 10, "Aucune TVA collectée pour un CA significatif — vérifier assujettissement"
    details['tva'] = {'points': pts, 'max': 18, 'label': '💰 Cohérence TVA', 'detail': lbl}

    # ── AXE 3 : Concentration fin d'exercice (12 pts) ────────
    p7  = fe.get('p7', 0)
    p30 = fe.get('p30', 0)
    if   p7 > 40:  pts, lbl = 12, f"{p7:.1f}% du volume sur les 7 derniers jours — écrêtage massif"
    elif p7 > 28:  pts, lbl =  9, f"{p7:.1f}% du volume sur les 7 derniers jours — suspect"
    elif p30 > 55: pts, lbl =  6, f"{p30:.1f}% du volume en fin de mois"
    elif p7 > 18:  pts, lbl =  3, f"{p7:.1f}% sur 7 derniers jours — à surveiller"
    else:          pts, lbl =  0, f"Répartition normale ({p7:.1f}% sur 7j)"
    details['fin_exercice'] = {'points': pts, 'max': 12, 'label': '📅 Concentration fin exercice', 'detail': lbl}

    # ── AXE 4 : Montants aberrants (10 pts) ──────────────────
    nb_ab = ab.get('nb', 0)
    if   nb_ab >= 10: pts, lbl = 10, f"{nb_ab} montants > 3σ — fort risque de manipulation"
    elif nb_ab >=  5: pts, lbl =  7, f"{nb_ab} montants aberrants à justifier"
    elif nb_ab >=  2: pts, lbl =  4, f"{nb_ab} montants atypiques"
    elif nb_ab ==  1: pts, lbl =  2, "1 montant aberrant à justifier"
    else:             pts, lbl =  0, "Aucun montant aberrant détecté"
    details['aberrants'] = {'points': pts, 'max': 10, 'label': '⚠️ Montants aberrants', 'detail': lbl}

    # ── AXE 5 : Marge vs secteur (12 pts) ────────────────────
    bm_s  = all_bm.get(secteur_code, {})
    tm_e  = bench.get('tm_entreprise', mg.get('tm', 0)) or 0
    tm_med = bm_s.get('tm_med', bm_s.get('tm_mediane', 0)) or 0
    tm_low = bm_s.get('tm_low', 0) or 0
    tm_high= bm_s.get('tm_high', 100) or 100
    pts = 0; lbl = "Marge dans la norme sectorielle"
    if tm_med > 0:
        ec_rel = abs(tm_e - tm_med) / tm_med * 100
        if tm_e < tm_low and ec_rel > 40:
            pts, lbl = 12, f"Marge {tm_e:.1f}% très inférieure à la médiane {tm_med:.0f}% — sous-valorisation probable"
        elif tm_e > tm_high and ec_rel > 40:
            pts, lbl = 10, f"Marge {tm_e:.1f}% très supérieure à la médiane {tm_med:.0f}% — sur-valorisation ou erreur secteur"
        elif tm_e < tm_low:
            pts, lbl =  7, f"Marge {tm_e:.1f}% sous le bas de fourchette ({tm_low:.0f}%)"
        elif tm_e > tm_high:
            pts, lbl =  5, f"Marge {tm_e:.1f}% au-dessus du haut de fourchette ({tm_high:.0f}%)"
        elif ec_rel > 30:
            pts, lbl =  3, f"Marge atypique (écart {ec_rel:.0f}% vs médiane secteur)"
    details['marge'] = {'points': pts, 'max': 12, 'label': '📈 Marge vs secteur', 'detail': lbl}

    # ── AXE 6 : IS / Résultat (10 pts) ───────────────────────
    tx_is  = fisc.get('tx_is', 0)
    is_d   = fisc.get('is_d', 0)
    res_f  = fisc.get('res', mg.get('res', 0)) or 0
    pts = 0; lbl = "Position IS normale"
    if res_f > 30000 and is_d == 0:
        pts, lbl = 10, f"Résultat {res_f:,.0f} € sans IS comptabilisé — vérifier TVA et régime fiscal"
    elif tx_is > 40:
        pts, lbl =  8, f"Taux IS apparent {tx_is:.1f}% — supérieur au taux légal 25%"
    elif 0 < tx_is < 5 and res_f > 20000:
        pts, lbl =  7, f"Taux IS très faible {tx_is:.1f}% pour résultat {res_f:,.0f} €"
    elif res_f < -50000:
        pts, lbl =  5, f"Perte importante {res_f:,.0f} € — vérifier charges exceptionnelles"
    details['is'] = {'points': pts, 'max': 10, 'label': '🏛️ Impôt sur les Sociétés', 'detail': lbl}

    # ── AXE 7 : Charges salariales (8 pts) ───────────────────
    rc  = sal.get('ratio_cotis', 0)
    sal_bruts = sal.get('sal_bruts', 0)
    pts = 0; lbl = "Charges salariales normales"
    if rc > 0:
        if   rc < 15: pts, lbl = 8, f"Cotisations/salaires {rc:.1f}% — anormal (attendu 42-55%)"
        elif rc < 25: pts, lbl = 6, f"Cotisations/salaires {rc:.1f}% — trop bas"
        elif rc > 70: pts, lbl = 5, f"Cotisations/salaires {rc:.1f}% — trop élevé"
        else:         lbl = f"Ratio cotisations/salaires {rc:.1f}% — OK"
    elif sal_bruts > 5000:
        pts, lbl = 4, f"Salaires {sal_bruts:,.0f} € sans cotisations sociales détectées"
    details['salaires'] = {'points': pts, 'max': 8, 'label': '👥 Charges salariales', 'detail': lbl}

    # ── AXE 8 : Intégrité FEC (6 pts) ────────────────────────
    nb_dup = dup.get('nb', 0)
    if   nb_dup > 30: pts, lbl = 6, f"{nb_dup} doublons — importation multiple probable"
    elif nb_dup > 10: pts, lbl = 4, f"{nb_dup} doublons à vérifier"
    elif nb_dup >  3: pts, lbl = 2, f"{nb_dup} doublons mineurs"
    else:             pts, lbl = 0, "Intégrité FEC vérifiée"
    details['integrite'] = {'points': pts, 'max': 6, 'label': '🔒 Intégrité du FEC', 'detail': lbl}

    # ── AXE 9 : Concentration client (10 pts) — NOUVEAU ──────
    # Compter la gravité des anomalies de concentration issues des contrôles
    conc_anom = [a for a in anomalies
                 if 'concentration' in a.get('type','').lower()
                 or 'client' in a.get('type','').lower()]
    # Aussi via stats ecommerce
    ratio_top1 = ec.get('concentration_client_top1', 0)
    if ratio_top1 == 0:
        # Calculer directement
        clients_df = df[df['CompteNum'].str.startswith('411')]
        if not clients_df.empty and ca_total > 0:
            grp = clients_df.groupby('CompteLib')['Credit'].sum()
            if len(grp):
                ratio_top1 = float(grp.max()) / ca_total * 100

    if   ratio_top1 > 90: pts, lbl = 10, f"Client unique {ratio_top1:.1f}% du CA — dépendance totale"
    elif ratio_top1 > 70: pts, lbl =  7, f"Client dominant {ratio_top1:.1f}% du CA"
    elif ratio_top1 > 50: pts, lbl =  4, f"Concentration client {ratio_top1:.1f}%"
    elif len(conc_anom) >= 2: pts, lbl = 3, "Plusieurs anomalies de concentration détectées"
    else:                  pts, lbl =  0, "Concentration client acceptable"
    details['concentration'] = {'points': pts, 'max': 10, 'label': '🎯 Concentration client', 'detail': lbl}

    # ── AXE 10 : Charges mixtes pro/perso (8 pts) — NOUVEAU ──
    perso = ec.get('charges_potentiellement_perso', 0)
    if perso == 0:
        # détecter directement
        kw_perso = ['PHARMACIE','RESTAURANT','APPLE','FNAC','DARTY','IKEA',
                    'LECLERC','CARREFOUR','BRICO','AMAZON ACHATS']
        four_df = df[df['CompteNum'].str.startswith('401')]
        perso = 0.0
        if not four_df.empty:
            for kw in kw_perso:
                mask = four_df['CompteLib'].str.upper().str.contains(kw, na=False)
                perso += float(four_df.loc[mask, 'Debit'].sum())

    ratio_perso = perso / ca_total * 100 if ca_total > 0 else 0
    if   ratio_perso > 8: pts, lbl = 8, f"Charges potentiellement personnelles {perso:,.0f} € ({ratio_perso:.1f}% CA)"
    elif ratio_perso > 4: pts, lbl = 5, f"Charges mixtes à justifier {perso:,.0f} €"
    elif ratio_perso > 1: pts, lbl = 2, f"Petites dépenses mixtes {perso:,.0f} €"
    else:                 pts, lbl = 0, "Aucune charge personnelle détectée"
    details['perso'] = {'points': pts, 'max': 8, 'label': '🛒 Charges mixtes pro/perso', 'detail': lbl}

    # ── AXE 11 : Montants ronds suspects (6 pts) — NOUVEAU ───
    mr = stats.get('montants_ronds', {})
    pct_ronds = mr.get('pct', 0)
    if   pct_ronds > 40: pts, lbl = 6, f"{pct_ronds:.1f}% de montants ronds — inventions probables"
    elif pct_ronds > 25: pts, lbl = 4, f"{pct_ronds:.1f}% de montants ronds — à justifier"
    elif pct_ronds > 15: pts, lbl = 2, f"{pct_ronds:.1f}% de montants ronds — surveillance"
    else:                pts, lbl = 0, f"Montants ronds {pct_ronds:.1f}% — normal"
    details['montants_ronds'] = {'points': pts, 'max': 6, 'label': '🔢 Montants ronds suspects', 'detail': lbl}

    # ── AXE 12 : Frais d'envoi disproportionnés (8 pts) — NOUVEAU ──
    ratio_envoi = ec.get('ratio_frais_envoi_ca', 0)
    if ratio_envoi == 0 and ca_total > 0:
        tel = df[df['CompteNum'].str.startswith('626')]
        kw  = ['POSTE', 'COLISSIMO', 'CHRONOPOST', 'DHL', 'UPS', 'FEDEX']
        if not tel.empty:
            mask = tel['EcritureLib'].str.upper().str.contains('|'.join(kw), na=False)
            ratio_envoi = float(tel.loc[mask, 'Debit'].sum()) / ca_total * 100

    if   ratio_envoi > 20: pts, lbl = 8, f"Frais d'envoi {ratio_envoi:.1f}% du CA — CRITIQUE (norme < 8%)"
    elif ratio_envoi > 12: pts, lbl = 5, f"Frais d'envoi {ratio_envoi:.1f}% du CA — élevés"
    elif ratio_envoi > 8:  pts, lbl = 2, f"Frais d'envoi {ratio_envoi:.1f}% du CA — légèrement élevés"
    else:                  pts, lbl = 0, f"Frais d'envoi {ratio_envoi:.1f}% — normaux"
    details['frais_envoi'] = {'points': pts, 'max': 8, 'label': '📦 Frais d\'envoi / logistique', 'detail': lbl}

    # ── Score global ──────────────────────────────────────────
    score_g = min(100, sum(v['points'] for v in details.values()))
    max_theo = sum(v['max'] for v in details.values())

    niveau = 'TRÈS FAIBLE'; conseil = GRILLE_RISQUE_FISCAL_V2[-1][2]
    for seuil, lbl_n, cons in GRILLE_RISQUE_FISCAL_V2:
        if score_g >= seuil:
            niveau = lbl_n; conseil = cons; break

    resultats = {
        'score_global':   score_g,
        'niveau_risque':  niveau,
        'conseil':        conseil,
        'details':        details,
        'nb_axes_risque': sum(1 for v in details.values() if v['points'] > 0),
        'max_theorique':  max_theo,
    }
    return [], resultats


# ══════════════════════════════════════════════════════════════
# 5. FONCTION D'APPLICATION DU PATCH
# ══════════════════════════════════════════════════════════════

def appliquer(module):
    """
    Remplace les fonctions clés du module fec_audit_pro_v5 par les versions v2.
    Appelé depuis app.py après l'import du module moteur.
    """
    # Enrichir BENCHMARKS avec les nouveaux secteurs
    module.BENCHMARKS.update(BENCHMARKS_SUPPLEMENTAIRES)

    # Remplacer SECTEURS_LISTE
    module.SECTEURS_LISTE = SECTEURS_LISTE_COMPLET

    # Remplacer la détection de secteur
    module.detecter_secteur = detecter_secteur_v2

    # Remplacer le bouclier fiscal
    module.calculer_bouclier_fiscal = calculer_bouclier_fiscal_v2

    # Ajouter l'analyseur e-commerce
    module.analyser_ecommerce = analyser_ecommerce

    print("[PATCH v6] Appliqué : détection secteur v2, bouclier fiscal 12 axes, "
          "analyseur e-commerce injecté.")
