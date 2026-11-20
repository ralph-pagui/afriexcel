"""
Test de Validation Phase 3 : Non-Régression et Équilibre Absolu des Modèles Financiers
Vérifie les 3 cas d'usage majeurs :
Test A: BENIN FRUIT INDUSTRIES SARL (Bénin, Cotonou, Jus de mangue)
Test B: TECH DISTRIB CAMEROUN (Cameroun, Douala, Matériel informatique, 25M FCFA)
Test C: SEN TRANS LOGISTICS SARL (Sénégal, Dakar, Transport de marchandises, 5 camions 10t)
"""

import sys
import io
import os
import openpyxl

# Ajouter backend au path
sys.path.insert(0, os.path.abspath("backend"))

from engine.models import PlanRequest
from engine.prompt_parser import PromptParser
from engine.financial_engine import FinancialEngine
from engine.excel_generator import ExcelReportGenerator
from engine.templates import detect_or_get_template
from engine.quality_control import QualityControl
from engine.sectors import SECTOR_CATALOG, detect_sector

def run_test_a():
    print("=" * 80)
    print("TEST A : BENIN FRUIT INDUSTRIES SARL (Bénin, Cotonou - Transformation Mangues)")
    print("=" * 80)

    prompt = """Génère un business plan professionnel, complet et bancable directement dans ce classeur Excel pour la création d'une entreprise agroalimentaire spécialisée dans la transformation de mangues fraîches en jus naturels conditionnés en bouteilles au Bénin.

# 1. IDENTITÉ DU PROJET
- Nom de l'entreprise : BENIN FRUIT INDUSTRIES SARL
- Pays : Bénin
- Ville d'implantation : Cotonou
- Secteur : Industrie agroalimentaire
- Activité principale : Transformation de mangues fraîches en jus naturels.
- Produits commercialisés : Jus de mangue naturel en bouteilles de 33 cl et 1 litre.
- Capacité de production initiale : 3 000 litres de jus par jour.
- Jours de production : 22 jours par mois.
- Horizon prévisionnel : 5 ans.
- Marché cible : Supermarchés, grossistes, restaurants, hôtels, distributeurs, commerces de proximité et particuliers.
- Positionnement : Jus naturels fabriqués localement, avec une qualité constante et un prix compétitif.

# 2. ÉTUDE DE MARCHÉ
Réalise une étude de marché approfondie sur le secteur des jus naturels au Bénin. Ne confonds pas la réglementation béninoise avec celle du Cameroun.

# 3. INVESTISSEMENTS & FINANCEMENT
Établis un plan d'investissement complet.
Financement : 30% d'apport personnel et 70% d'emprunt bancaire au taux de 11% sur 5 ans.
"""

    parsed = PromptParser.parse(prompt, default_country="Côte d'Ivoire")
    print(f"1. Parsing: Pays={parsed['country']}, Ville={parsed['city']}, Secteur={parsed['sector_id']}, Societe={parsed['company_name']}")
    assert parsed["country"] == "Bénin", f"Erreur pays: attendu Bénin, obtenu {parsed['country']}"
    assert parsed["sector_id"] == "jus", f"Erreur secteur: attendu jus, obtenu {parsed['sector_id']}"
    assert "BENIN FRUIT" in parsed["company_name"], f"Erreur nom: {parsed['company_name']}"

    # Vérification des produits extraits
    prods = parsed.get("explicit_products", [])
    print(f"2. Produits extraits: {[p['name'] for p in prods]}")
    assert len(prods) == 2, f"Attendu 2 produits (33cl et 1L), obtenu {len(prods)}"
    for p in prods:
        assert "ananas" not in p["name"].lower() and "bissap" not in p["name"].lower(), f"Produit contaminé: {p['name']}"

    # Construction du profil
    profile = SECTOR_CATALOG["jus"]
    from engine.models import ProductItem
    profile.products = [
        ProductItem(
            id=p.get("id", idx + 1),
            name=p["name"],
            description=p.get("description", p["name"]),
            share_pct=float(p.get("share_pct", 0.5)),
            unit_price=float(p.get("unit_price", 1000.0))
        )
        for idx, p in enumerate(prods)
    ]

    # Modélisation financière
    fin_data = FinancialEngine.compute(
        profile,
        currency=parsed["currency"],
        horizon_years=parsed["horizon_years"],
        country=parsed["country"],
        equity_ratio=parsed["equity_ratio"],
        loan_ratio=parsed["loan_ratio"],
        loan_interest_rate=parsed["loan_interest_rate"],
        loan_duration_years=parsed["loan_duration_years"]
    )

    kpis = fin_data["kpis"]
    print(f"3. KPIs: Investissement={kpis.total_investment:,.0f} FCFA, Apport={kpis.equity_promoter:,.0f} ({kpis.equity_pct}%), Emprunt={kpis.bank_loan:,.0f} ({kpis.loan_pct}%)")
    assert kpis.equity_pct == 30.0, f"Attendu 30% FP, obtenu {kpis.equity_pct}"
    assert kpis.loan_pct == 70.0, f"Attendu 70% Emprunt, obtenu {kpis.loan_pct}"
    assert kpis.loan_interest_rate == 11.0, f"Attendu 11% taux, obtenu {kpis.loan_interest_rate}"

    # Vérification équilibre Emplois == Ressources
    inv = fin_data["investments"]
    fin = fin_data["financing"]
    tot_emplois = inv["frais_etab"] + inv["total"] + fin["bfr_initial"]
    tot_ressources = fin["apport"] + fin["emprunt"]
    diff_fin = abs(tot_emplois - tot_ressources)
    print(f"4. Plan de Financement: Emplois={tot_emplois:,.0f} FCFA | Ressources={tot_ressources:,.0f} FCFA | Diff={diff_fin:,.0f} FCFA")
    assert diff_fin == 0, f"Déséquilibre Emplois/Ressources : {diff_fin} FCFA"

    # Vérification équilibre Bilan 5 ans
    for b in fin_data["bilan_5years"]:
        assert abs(b["diff"]) == 0, f"Bilan déséquilibré {b['year']}: Actif={b['total_actif']} != Passif={b['total_passif']}"
    print("5. Bilan Prévisionnel sur 5 ans : 100% équilibré (Actif == Passif pour chaque année)")

    # Détection template
    theme = detect_or_get_template("auto", prompt, profile.activity, sector_id="jus")
    print(f"6. Template: {theme.name} | Subtitle='{theme.subtitle}' | Tagline='{theme.tagline}'")
    assert "transport" not in theme.subtitle.lower() and "fret" not in theme.subtitle.lower(), f"Template contaminé par transport: {theme.subtitle}"
    assert "agroalimentaire" in theme.subtitle.lower() or "fruits" in theme.subtitle.lower(), f"Template inadéquat: {theme.subtitle}"

    # Génération du classeur Excel
    excel_bytes = ExcelReportGenerator.build(fin_data, project_name=parsed["company_name"], country=parsed["country"], template_id=theme.id)
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    assert len(wb.sheetnames) == 20, f"Attendu 20 feuilles, obtenu {len(wb.sheetnames)}"

    # Audit Workbook
    qc = QualityControl.audit_workbook(wb, sector_id="jus")
    print(f"7. Audit Qualité Workbook: Passed={qc.passed}, Errors={qc.errors}")
    assert qc.passed, f"Audit Qualité échoué : {qc.errors}"

    # Vérification cellule Sheet 11 Emplois vs Ressources
    ws11 = wb["11_Plan_Financement"]
    val_e = ws11.cell(row=8, column=3).value
    val_r = ws11.cell(row=8, column=6).value
    # Dans data_only, les formules SUM ne sont pas calculées par openpyxl sans Excel, vérifions les cellules sources
    c_frais = ws11.cell(row=5, column=3).value
    c_inv = ws11.cell(row=6, column=3).value
    c_bfr = ws11.cell(row=7, column=3).value
    c_app = ws11.cell(row=5, column=6).value
    c_emp = ws11.cell(row=6, column=6).value
    sum_emp = c_frais + c_inv + c_bfr
    sum_res = c_app + c_emp
    print(f"8. Sheet 11 Cellules: Emplois={sum_emp:,.0f} FCFA, Ressources={sum_res:,.0f} FCFA, Ecart={abs(sum_emp - sum_res)} FCFA")
    assert sum_emp == sum_res, f"Incohérence Sheet 11: Emplois ({sum_emp}) != Ressources ({sum_res})"

    # Sauvegarde du fichier pour inspection
    out_path = "Test_A_BENIN_FRUIT_INDUSTRIES.xlsx"
    with open(out_path, "wb") as f:
        f.write(excel_bytes)
    print(f"Fichier généré : {out_path} ({len(excel_bytes):,} octets)")
    print(">>> TEST A : VALIDÉ AVEC SUCCÈS (100% SUCCÈS) <<<\n")


def run_test_b():
    print("=" * 80)
    print("TEST B : TECH DISTRIB CAMEROUN (Cameroun, Douala - Matériel Informatique 25M FCFA)")
    print("=" * 80)

    prompt = """Génère un business plan pour une entreprise de distribution de matériel informatique et solutions réseaux à Douala au Cameroun.
Nom : TECH DISTRIB CAMEROUN SARL
Budget total : 25 000 000 FCFA
Durée : 5 ans
Financement : 40% apport personnel, 60% emprunt bancaire à 12% l'an sur 5 ans.
Activité : Vente d'ordinateurs portables professionnels, serveurs et périphériques réseaux pour PME.
"""

    parsed = PromptParser.parse(prompt, default_country="Cameroun")
    print(f"1. Parsing: Pays={parsed['country']}, Ville={parsed['city']}, Secteur={parsed['sector_id']}, Budget={parsed['total_budget']}")
    assert parsed["country"] == "Cameroun", f"Erreur pays: attendu Cameroun, obtenu {parsed['country']}"
    assert parsed["sector_id"] == "commerce_informatique", f"Erreur secteur: attendu commerce_informatique, obtenu {parsed['sector_id']}"
    assert parsed["total_budget"] == 25000000.0, f"Erreur budget: attendu 25M, obtenu {parsed['total_budget']}"

    profile = SECTOR_CATALOG["commerce_informatique"]
    fin_data = FinancialEngine.compute(
        profile,
        currency="FCFA",
        horizon_years=5,
        target_budget=25000000.0,
        country="Cameroun",
        equity_ratio=0.40,
        loan_ratio=0.60,
        loan_interest_rate=0.12,
        loan_duration_years=5
    )

    kpis = fin_data["kpis"]
    print(f"2. KPIs: Investissement={kpis.total_investment:,.0f} FCFA, Apport={kpis.equity_promoter:,.0f} (40%), Emprunt={kpis.bank_loan:,.0f} (60%)")
    assert kpis.total_investment == 25000000.0, f"Investissement non égal au budget : {kpis.total_investment}"
    assert kpis.equity_promoter == 10000000.0, f"Apport non égal à 10M : {kpis.equity_promoter}"
    assert kpis.bank_loan == 15000000.0, f"Emprunt non égal à 15M : {kpis.bank_loan}"

    # Vérification équilibre Emplois == Ressources
    inv = fin_data["investments"]
    fin = fin_data["financing"]
    tot_emplois = inv["frais_etab"] + inv["total"] + fin["bfr_initial"]
    tot_ressources = fin["apport"] + fin["emprunt"]
    diff_fin = abs(tot_emplois - tot_ressources)
    print(f"3. Plan de Financement: Emplois={tot_emplois:,.0f} FCFA | Ressources={tot_ressources:,.0f} FCFA | Diff={diff_fin:,.0f} FCFA")
    assert diff_fin == 0, f"Déséquilibre Emplois/Ressources : {diff_fin} FCFA"

    # Vérification fiscalité Cameroun (IS = 28%, CNPS = 14%)
    assert fin_data["is_rate"] == 0.28, f"Taux IS Cameroun erroné: {fin_data['is_rate']}"
    assert fin_data["cnps_rate"] == 0.14, f"Taux CNPS Cameroun erroné: {fin_data['cnps_rate']}"
    print("4. Règles fiscales Cameroun validées (IS 28%, CNPS 14%)")

    # Template
    theme = detect_or_get_template("auto", prompt, profile.activity, sector_id="commerce_informatique")
    print(f"5. Template: {theme.name} | Subtitle='{theme.subtitle}'")
    assert "informatique" in theme.subtitle.lower() or "technolog" in theme.subtitle.lower(), f"Template inadapté: {theme.subtitle}"

    excel_bytes = ExcelReportGenerator.build(fin_data, project_name=parsed["company_name"], country=parsed["country"], template_id=theme.id)
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    qc = QualityControl.audit_workbook(wb, sector_id="commerce_informatique")
    print(f"6. Audit Qualité: Passed={qc.passed}, Errors={qc.errors}")
    assert qc.passed, f"Audit Qualité échoué : {qc.errors}"

    out_path = "Test_B_TECH_DISTRIB_CAMEROUN.xlsx"
    with open(out_path, "wb") as f:
        f.write(excel_bytes)
    print(f"Fichier généré : {out_path} ({len(excel_bytes):,} octets)")
    print(">>> TEST B : VALIDÉ AVEC SUCCÈS (100% SUCCÈS) <<<\n")


def run_test_c():
    print("=" * 80)
    print("TEST C : SEN TRANS LOGISTICS SARL (Sénégal, Dakar - Transport Marchandises 5 Camions)")
    print("=" * 80)

    prompt = """Génère un business plan bancable pour SEN TRANS LOGISTICS SARL au Sénégal à Dakar.
Activité : Transport routier de marchandises, interurbain et logistique avec une flotte de 5 camions 10 tonnes.
Zone : Dakar, Thiès, Kaolack, Saint-Louis.
Financement : 35% fonds propres et 65% dette bancaire à 11% sur 5 ans.
"""

    parsed = PromptParser.parse(prompt, default_country="Sénégal")
    trucks = parsed["fleet_ops"].get("truck_count") or parsed["fleet_ops"].get("num_trucks")
    print(f"1. Parsing: Pays={parsed['country']}, Ville={parsed['city']}, Secteur={parsed['sector_id']}, Camions={trucks}")
    assert parsed["country"] == "Sénégal", f"Erreur pays: attendu Sénégal, obtenu {parsed['country']}"
    assert parsed["sector_id"] == "transport", f"Erreur secteur: attendu transport, obtenu {parsed['sector_id']}"
    assert trucks == 5, f"Erreur camions: {parsed['fleet_ops']}"

    profile = SECTOR_CATALOG["transport"]
    fin_data = FinancialEngine.compute(
        profile,
        currency="FCFA",
        horizon_years=5,
        country="Sénégal",
        equity_ratio=parsed["equity_ratio"],
        loan_ratio=parsed["loan_ratio"],
        loan_interest_rate=parsed["loan_interest_rate"],
        loan_duration_years=parsed["loan_duration_years"]
    )

    kpis = fin_data["kpis"]
    print(f"2. KPIs: Investissement={kpis.total_investment:,.0f} FCFA, Apport={kpis.equity_promoter:,.0f} (35%), Emprunt={kpis.bank_loan:,.0f} (65%)")
    assert kpis.equity_pct == 35.0, f"Attendu 35% FP, obtenu {kpis.equity_pct}"
    assert kpis.loan_pct == 65.0, f"Attendu 65% Emprunt, obtenu {kpis.loan_pct}"
    assert kpis.loan_interest_rate == 11.0, f"Attendu 11% taux, obtenu {kpis.loan_interest_rate}"

    inv = fin_data["investments"]
    fin = fin_data["financing"]
    tot_emplois = inv["frais_etab"] + inv["total"] + fin["bfr_initial"]
    tot_ressources = fin["apport"] + fin["emprunt"]
    diff_fin = abs(tot_emplois - tot_ressources)
    print(f"3. Plan de Financement: Emplois={tot_emplois:,.0f} FCFA | Ressources={tot_ressources:,.0f} FCFA | Diff={diff_fin:,.0f} FCFA")
    assert diff_fin == 0, f"Déséquilibre Emplois/Ressources : {diff_fin} FCFA"

    theme = detect_or_get_template("auto", prompt, profile.activity, sector_id="transport")
    print(f"4. Template: {theme.name} | Subtitle='{theme.subtitle}'")
    assert "transport" in theme.subtitle.lower() or "fret" in theme.subtitle.lower(), f"Template inadapté: {theme.subtitle}"

    excel_bytes = ExcelReportGenerator.build(fin_data, project_name=parsed["company_name"], country=parsed["country"], template_id=theme.id)
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    qc = QualityControl.audit_workbook(wb, sector_id="transport")
    print(f"5. Audit Qualité: Passed={qc.passed}, Errors={qc.errors}")
    assert qc.passed, f"Audit Qualité échoué : {qc.errors}"

    out_path = "Test_C_SEN_TRANS_LOGISTICS.xlsx"
    with open(out_path, "wb") as f:
        f.write(excel_bytes)
    print(f"Fichier généré : {out_path} ({len(excel_bytes):,} octets)")
    print(">>> TEST C : VALIDÉ AVEC SUCCÈS (100% SUCCÈS) <<<\n")


if __name__ == "__main__":
    try:
        run_test_a()
        run_test_b()
        run_test_c()
        print("=" * 80)
        print("TOUS LES 3 TESTS DE VALIDATION PHASE 3 ONT RÉUSSI AVEC SUCCÈS (100%) !")
        print("=" * 80)
    except Exception as e:
        import traceback
        print(f"\n[ÉCHEC TEST] : {e}")
        traceback.print_exc()
        sys.exit(1)
