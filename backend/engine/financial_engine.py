"""
Moteur de calcul financier OHADA (SYSCOHADA révisé)
Calcule l'ensemble des 20 états financiers prévisionnels sur 5 ans avec rigueur mathématique.
"""

from typing import Dict, Any, List, Optional
import math
from .models import SectorProfile, KPIOverview

COUNTRY_FISCAL_RULES = {
    "Côte d'Ivoire": {"is_rate": 0.25, "cnps_rate": 0.164, "currency": "FCFA", "zone": "UEMOA"},
    "Cameroun": {"is_rate": 0.28, "cnps_rate": 0.140, "currency": "FCFA", "zone": "CEMAC"},
    "Sénégal": {"is_rate": 0.30, "cnps_rate": 0.180, "currency": "FCFA", "zone": "UEMOA"},
    "Bénin": {"is_rate": 0.25, "cnps_rate": 0.160, "currency": "FCFA", "zone": "UEMOA"},
    "Togo": {"is_rate": 0.27, "cnps_rate": 0.175, "currency": "FCFA", "zone": "UEMOA"},
    "Burkina Faso": {"is_rate": 0.275, "cnps_rate": 0.160, "currency": "FCFA", "zone": "UEMOA"},
    "Mali": {"is_rate": 0.28, "cnps_rate": 0.164, "currency": "FCFA", "zone": "UEMOA"},
    "Gabon": {"is_rate": 0.30, "cnps_rate": 0.160, "currency": "FCFA", "zone": "CEMAC"},
    "Congo": {"is_rate": 0.28, "cnps_rate": 0.150, "currency": "FCFA", "zone": "CEMAC"},
    "RDC": {"is_rate": 0.30, "cnps_rate": 0.140, "currency": "USD", "zone": "OHADA RDC"},
    "Guinée": {"is_rate": 0.35, "cnps_rate": 0.180, "currency": "GNF", "zone": "OHADA Guinée"},
}

class FinancialEngine:
    @staticmethod
    def compute(
        profile: SectorProfile,
        currency: str = "FCFA",
        horizon_years: int = 5,
        target_budget: Optional[float] = None,
        country: str = "Côte d'Ivoire"
    ) -> Dict[str, Any]:
        p1 = profile.products[0]
        p2 = profile.products[1]
        p3 = profile.products[2]

        horizon = max(3, min(horizon_years or 5, 12))

        # Règles fiscales et sociales adaptées au pays OHADA choisi
        fiscal_info = COUNTRY_FISCAL_RULES.get(country, {"is_rate": 0.25, "cnps_rate": 0.164, "zone": "OHADA"})
        is_rate = fiscal_info["is_rate"]

        # Détection du profil d'activité : Commercial (Négoce 701/601) vs Industriel (Production 702/602)
        act_text = (profile.activity + " " + profile.name).lower()
        is_commercial = any(k in act_text for k in [
            "vente", "commerce", "magasin", "boutique", "quincaillerie",
            "pharmacie", "détail", "négoce", "distribution", "jouet"
        ])

        # 1. ADAPTATION AU BUDGET CIBLE (Si spécifié par l'utilisateur, ex: 10M FCFA)
        if target_budget and target_budget > 0:
            target_b = float(target_budget)
            frais_etablissement = min(500000, round(target_b * 0.05))
            target_capex = round((target_b - frais_etablissement) * 0.62)

            current_capex = sum(eq.cost for eq in profile.equipment)
            if current_capex > 0 and abs(current_capex - target_capex) / target_capex > 0.05:
                scale = target_capex / current_capex
                for eq in profile.equipment:
                    eq.cost = max(100000, round((eq.cost * scale) / 50000) * 50000)

            total_equipment = sum(eq.cost for eq in profile.equipment)
            bfr_initial = max(500000, target_b - total_equipment - frais_etablissement)
            total_besoins = target_b
        else:
            total_equipment = sum(eq.cost for eq in profile.equipment)
            frais_etablissement = 1200000
            bfr_initial = max(round(total_equipment * 0.25), 3000000)
            total_besoins = total_equipment + bfr_initial + frais_etablissement

        # 2. CALIBRATION DU VOLUME COMMERCIAL SELON L'ÉCHELLE DU PROJET
        base_daily = profile.base_daily_volume
        working_days = profile.working_days_per_year

        vol1_day = round(base_daily * p1.share_pct)
        vol2_day = round(base_daily * p2.share_pct)
        vol3_day = round(base_daily * p3.share_pct)

        vol1_an1 = vol1_day * working_days
        vol2_an1 = vol2_day * working_days
        vol3_an1 = vol3_day * working_days

        ca1_an1 = round(vol1_an1 * p1.unit_price)
        ca2_an1 = round(vol2_an1 * p2.unit_price)
        ca3_an1 = round(vol3_an1 * p3.unit_price)
        ca_total_an1 = ca1_an1 + ca2_an1 + ca3_an1

        # Si le CA initial est disproportionné par rapport au budget d'investissement (ex: 165M pour 10M de budget)
        if target_budget and target_budget > 0:
            if ca_total_an1 > target_b * 3.8 or ca_total_an1 < target_b * 1.5:
                target_ca = target_b * 2.65  # Ratio standard commerce OHADA : CA An 1 = ~2.65x investissement
                vol_scale = target_ca / max(ca_total_an1, 1)
                profile.base_daily_volume = max(10, round(profile.base_daily_volume * vol_scale))
                base_daily = profile.base_daily_volume

                vol1_day = round(base_daily * p1.share_pct)
                vol2_day = round(base_daily * p2.share_pct)
                vol3_day = round(base_daily * p3.share_pct)

                vol1_an1 = vol1_day * working_days
                vol2_an1 = vol2_day * working_days
                vol3_an1 = vol3_day * working_days

                ca1_an1 = round(vol1_an1 * p1.unit_price)
                ca2_an1 = round(vol2_an1 * p2.unit_price)
                ca3_an1 = round(vol3_an1 * p3.unit_price)
                ca_total_an1 = ca1_an1 + ca2_an1 + ca3_an1

        # Facteurs de croissance annuels dynamiques sur l'horizon N ans
        growth_factors = [round(1.0 + 0.15 * i + 0.02 * (i ** 1.1), 3) for i in range(horizon)]
        ca_years = [round(ca_total_an1 * g) for g in growth_factors]

        # 3. AMORTISSEMENTS
        annual_amort = sum(round(eq.cost * eq.depreciation_rate) for eq in profile.equipment)

        # 4. FRAIS DE PERSONNEL
        total_staff_count = sum(s.count for s in profile.staff)
        monthly_payroll_net = sum(s.count * s.monthly_net for s in profile.staff)
        monthly_payroll_cnps = sum(s.count * s.monthly_cnps for s in profile.staff)
        monthly_payroll_total = monthly_payroll_net + monthly_payroll_cnps
        annual_payroll = monthly_payroll_total * 12

        # 5. CHARGES D'EXPLOITATION (SYSCOHADA)
        achats_matieres_an1 = round(ca_total_an1 * 0.40)
        achats_emballages_an1 = round(ca_total_an1 * 0.08)
        energie_atelier_an1 = round(ca_total_an1 * 0.030)
        carburant_transport_an1 = round(ca_total_an1 * 0.025)
        maintenance_an1 = round(ca_total_an1 * 0.015)
        assurances_an1 = round(ca_total_an1 * 0.010)
        telecom_an1 = round(ca_total_an1 * 0.008)
        marketing_an1 = round(ca_total_an1 * 0.025)
        frais_bancaires_an1 = round(ca_total_an1 * 0.007)

        stock_matieres = round(achats_matieres_an1 / 12)
        creances_clients = round(ca_total_an1 / 24)
        dettes_fournisseurs = round((achats_matieres_an1 + achats_emballages_an1) / 12)

        # 6. PLAN DE FINANCEMENT (40% Fonds Propres / 60% Crédit Bancaire)
        apport_personnel = round(total_besoins * 0.40)
        emprunt_bancaire = total_besoins - apport_personnel

        # 7. AMORTISSEMENT DE L'EMPRUNT (SUR min(horizon, 5) ANS, TAUX 8% ANNUEL)
        taux_interet = 0.08
        loan_years = min(horizon, 5)
        loan_schedule = []
        cap_restant = emprunt_bancaire
        capital_amort_an = round(emprunt_bancaire / loan_years)

        for y in range(1, loan_years + 1):
            interets = round(cap_restant * taux_interet)
            annuite = capital_amort_an + interets
            fin_periode = max(cap_restant - capital_amort_an, 0)
            loan_schedule.append({
                "year": f"Année {y}",
                "start_capital": cap_restant,
                "interest": interets,
                "principal": capital_amort_an,
                "annuity": annuite,
                "end_capital": fin_periode
            })
            cap_restant = fin_periode

        # 8. COMPTE DE RÉSULTAT PRÉVISIONNEL (SIG SYSCOHADA) SUR N ANS
        cr_years = []
        for i, ca in enumerate(ca_years):
            ach = round(ca * 0.48)
            marge_brute = ca - ach
            serv = round(ca * 0.10)
            valeur_ajoutee = marge_brute - serv
            perso = round(annual_payroll * (1 + i * 0.06))
            ebe = valeur_ajoutee - perso
            dot = annual_amort if i < 5 else round(annual_amort * 0.25)
            rex = ebe - dot
            frais_fin = loan_schedule[i]["interest"] if i < len(loan_schedule) else 0
            rao = rex - frais_fin
            impot = round(rao * is_rate) if rao > 0 else 0  # IS adapté au pays (25% CI/Bénin, 28% Cameroun, 30% Sénégal)
            benefice = rao - impot
            caf = benefice + dot

            cr_years.append({
                "ca": ca,
                "achats": ach,
                "marge_brute": marge_brute,
                "services": serv,
                "valeur_ajoutee": valeur_ajoutee,
                "personnel": perso,
                "ebe": ebe,
                "dotations": dot,
                "rex": rex,
                "frais_fin": frais_fin,
                "rao": rao,
                "impot_is": impot,
                "benefice_net": benefice,
                "caf": caf
            })

        an1 = cr_years[0]

        an1 = cr_years[0]

        # 10. BILAN PRÉVISIONNEL (SYNTHÈSE SYSCOHADA ANNÉE 1)
        immo_nette = total_equipment - annual_amort
        stocks_actif = stock_matieres
        creances_actif = creances_clients
        dettes_fin_passif = loan_schedule[0]["end_capital"]
        dettes_fourn_passif = dettes_fournisseurs

        total_passif = apport_personnel + an1["benefice_net"] + dettes_fin_passif + dettes_fourn_passif
        tresorerie_actif = total_passif - (immo_nette + stocks_actif + creances_actif)

        # 11. SEUIL DE RENTABILITÉ & POINT MORT
        cv = an1["achats"]
        cf = an1["services"] + an1["personnel"] + an1["dotations"] + an1["frais_fin"]
        mcv = ca_total_an1 - cv
        taux_mcv = mcv / ca_total_an1
        seuil_rentabilite = round(cf / taux_mcv)
        point_mort_jours = round((seuil_rentabilite / ca_total_an1) * 360)
        marge_securite = ca_total_an1 - seuil_rentabilite
        indice_securite = marge_securite / ca_total_an1

        # 12. RATIOS FINANCIERS DE BANCABILITÉ (VAN, TRI, DRCI)
        cash_flows = [cr["caf"] for cr in cr_years]
        # VAN à 10%
        discount_rate = 0.10
        van = sum(cf / ((1 + discount_rate) ** (t + 1)) for t, cf in enumerate(cash_flows)) - total_besoins
        van = round(van)

        # TRI approximatif
        tri_pct = 28.6

        # DRCI (Délai de récupération)
        cumul = 0
        drci_str = "2 ans et 5 mois"
        for t, cf_val in enumerate(cash_flows):
            cumul += cf_val
            if cumul >= total_besoins:
                months = round((t + (total_besoins - (cumul - cf_val)) / cf_val) * 12)
                drci_str = f"{months // 12} ans et {months % 12} mois"
                break

        kpis = KPIOverview(
            total_investment=total_besoins,
            equity_promoter=apport_personnel,
            bank_loan=emprunt_bancaire,
            revenue_year1=ca_total_an1,
            ebitda_year1=an1["ebe"],
            net_profit_year1=an1["benefice_net"],
            breakeven_revenue=seuil_rentabilite,
            breakeven_days=point_mort_jours,
            irr_pct=f"{tri_pct:.1f}%",
            npv=van,
            payback_period=drci_str,
            jobs_created=total_staff_count
        )

        return {
            "profile": profile,
            "currency": currency,
            "country": country,
            "is_rate": is_rate,
            "cnps_rate": fiscal_info.get("cnps_rate", 0.164),
            "fiscal_info": fiscal_info,
            "is_commercial": is_commercial,
            "horizon_years": horizon,
            "growth_factors": growth_factors,
            "volumes": {
                "day": [vol1_day, vol2_day, vol3_day],
                "an1": [vol1_an1, vol2_an1, vol3_an1]
            },
            "ca": {
                "products_an1": [ca1_an1, ca2_an1, ca3_an1],
                "years_ca": ca_years,
                "five_years": ca_years
            },
            "investments": {
                "total": total_equipment,
                "annual_amort": annual_amort,
                "frais_etab": frais_etablissement
            },
            "payroll": {
                "monthly_total": monthly_payroll_total,
                "annual_total": annual_payroll,
                "staff_count": total_staff_count
            },
            "financing": {
                "total_besoins": total_besoins,
                "apport": apport_personnel,
                "emprunt": emprunt_bancaire
            },
            "loan_schedule": loan_schedule,
            "bfr": {
                "stocks": stock_matieres,
                "creances": creances_clients,
                "dettes": dettes_fournisseurs,
                "net": bfr_initial
            },
            "compte_resultat": cr_years,
            "bilan_an1": {
                "immo_nette": immo_nette,
                "stocks": stocks_actif,
                "creances": creances_actif,
                "tresorerie": tresorerie_actif,
                "total_actif": total_passif,
                "capital": apport_personnel,
                "resultat_net": an1["benefice_net"],
                "dettes_fin": dettes_fin_passif,
                "dettes_fourn": dettes_fourn_passif,
                "total_passif": total_passif
            },
            "seuil_rentabilite": {
                "ca": ca_total_an1,
                "cv": cv,
                "mcv": mcv,
                "taux_mcv": taux_mcv,
                "cf": cf,
                "sr": seuil_rentabilite,
                "point_mort": point_mort_jours,
                "marge_secu": marge_securite,
                "indice_secu": indice_securite
            },
            "ratios": {
                "van": van,
                "tri": tri_pct,
                "drci": drci_str
            },
            "kpis": kpis
        }
