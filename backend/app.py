"""
Serveur FastAPI AfriExcel - Moteur d'IA & Modélisation Financière OHADA en Python
Génère des Business Plans bancables certifiés OHADA en format natif Excel (.xlsx)
"""

import base64
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, Response

from engine.models import PlanRequest, PlanResponse, KPIOverview
from engine.sectors import SECTOR_CATALOG, detect_sector
from engine.financial_engine import FinancialEngine
from engine.excel_generator import ExcelReportGenerator
from engine.gemini_engine import GeminiSectorEngine
from engine.templates import TEMPLATE_CATALOG, detect_or_get_template
from engine.prompt_parser import PromptParser
from engine.quality_control import QualityControl

app = FastAPI(
    title="AfriExcel Financial Copilot API",
    description="Moteur de modélisation financière OHADA et génération de Business Plans Excel bancables",
    version="2.0.0"
)

# Configuration CORS intégrale pour Office.js, Netlify et Localhost
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def root():
    return {
        "service": "AfriExcel Python Copilot API",
        "status": "online",
        "version": "2.0.0",
        "docs_url": "/docs"
    }

@app.get("/api/health")
def health_check():
    return {
        "status": "ok",
        "service": "AfriExcel Python Financial Engine",
        "version": "2.0.0",
        "sectors_available": list(SECTOR_CATALOG.keys())
    }

@app.get("/api/sectors")
def list_sectors():
    return {
        "count": len(SECTOR_CATALOG),
        "sectors": [
            {
                "id": s.sector_id,
                "name": s.name,
                "activity": s.activity,
                "unit": s.unit,
                "products_count": len(s.products),
                "equipment_count": len(s.equipment),
                "staff_count": len(s.staff)
            }
            for s in SECTOR_CATALOG.values()
        ]
    }

@app.get("/api/templates")
def list_templates():
    return {
        "count": len(TEMPLATE_CATALOG),
        "templates": [
            {
                "id": t.id,
                "name": t.name,
                "subtitle": t.subtitle,
                "description": t.description,
                "icon": t.icon,
                "tagline": t.tagline,
                "vision_quote": t.vision_quote,
                "banner_bg": t.banner_bg,
                "header_bg": t.header_bg,
                "accent_color": t.accent_color,
                "highlight_bg": t.highlight_bg,
                "total_bg": t.total_bg
            }
            for t in TEMPLATE_CATALOG.values()
        ]
    }

@app.post("/api/generate-business-plan", response_model=PlanResponse)
def generate_business_plan(req: PlanRequest):
    try:
        # 1. Analyse sémantique approfondie du prompt utilisateur (extraction entités & contraintes financières)
        parsed = PromptParser.parse(
            req.prompt,
            default_country=req.country or "Côte d'Ivoire",
            default_currency=req.currency or "FCFA"
        )

        # Priorité absolue aux entités explicitement détectées dans le prompt
        if parsed.get("has_explicit_country"):
            country = parsed["country"]
        elif req.country and req.country != "Côte d'Ivoire":
            country = req.country
        else:
            country = parsed["country"]

        currency = parsed["currency"] if parsed.get("currency") and parsed["currency"] != "FCFA" else (req.currency or "FCFA")
        project_name = parsed["company_name"] if parsed.get("company_name") else (req.company_name or "Projet Entreprise")

        extracted_budget = req.total_budget or parsed["total_budget"]
        extracted_years = req.horizon_years or parsed["horizon_years"] or 5

        # Paramètres de financement : priorité absolue aux mentions du prompt
        if parsed.get("has_explicit_financing"):
            equity_ratio = parsed["equity_ratio"]
            loan_ratio = parsed["loan_ratio"]
        else:
            equity_ratio = req.equity_ratio if req.equity_ratio is not None else parsed["equity_ratio"]
            loan_ratio = req.loan_ratio if req.loan_ratio is not None else parsed["loan_ratio"]

        if parsed.get("has_explicit_interest_rate"):
            loan_interest_rate = parsed["loan_interest_rate"]
        else:
            loan_interest_rate = req.loan_interest_rate if req.loan_interest_rate is not None else parsed["loan_interest_rate"]

        loan_duration_years = req.loan_duration_years or parsed["loan_duration_years"] or 5
        sector_id = parsed.get("sector_id") or req.sector_id

        # 2. Conception ou sélection du profil sectoriel sans AUCUNE contamination étrangère
        dynamic_profile = GeminiSectorEngine.generate_dynamic_sector(
            user_prompt=req.prompt,
            country=country,
            currency=currency,
            target_budget=extracted_budget,
            horizon_years=extracted_years,
            clean_topic=project_name
        )

        if dynamic_profile:
            profile = dynamic_profile
        elif sector_id in SECTOR_CATALOG:
            profile = SECTOR_CATALOG[sector_id]
        else:
            profile = detect_sector(req.prompt)

        # Application stricte des produits explicites demandés par l'utilisateur
        if parsed.get("explicit_products"):
            from engine.models import ProductItem
            profile.products = [
                ProductItem(
                    id=p.get("id", idx + 1),
                    name=p["name"],
                    description=p.get("description", p["name"]),
                    share_pct=float(p.get("share_pct", 1.0 / len(parsed["explicit_products"]))),
                    unit_price=float(p.get("unit_price", 1000.0))
                )
                for idx, p in enumerate(parsed["explicit_products"])
            ]

        # Application des contraintes opérationnelles (flotte, jours de travail)
        if parsed.get("fleet_ops"):
            f_ops = parsed["fleet_ops"]
            if f_ops.get("working_days_per_year"):
                profile.working_days_per_year = f_ops["working_days_per_year"]
            if f_ops.get("daily_capacity") and profile.sector_id == "jus":
                profile.base_daily_volume = round(f_ops["daily_capacity"])

        # 3. Calcul rigoureux du modèle financier OHADA
        fin_data = FinancialEngine.compute(
            profile,
            currency=currency,
            horizon_years=extracted_years,
            target_budget=extracted_budget,
            country=country,
            equity_ratio=equity_ratio,
            loan_ratio=loan_ratio,
            loan_interest_rate=loan_interest_rate,
            loan_duration_years=loan_duration_years,
            discount_rate=req.discount_rate or 0.10,
            capacity_info=parsed.get("capacity_info")
        )
        kpis: KPIOverview = fin_data["kpis"]

        # 4. Contrôle Qualité Automatique (Quality Gate 14 points)
        qc_report = QualityControl.audit(fin_data, requested_params={
            "equity_ratio": equity_ratio,
            "loan_ratio": loan_ratio,
            "loan_interest_rate": loan_interest_rate
        })
        if not qc_report.passed:
            print(f"[Quality Gate Warning] {qc_report.errors}")

        # 5. Choix et adaptation chirurgicale du Template Visuel
        selected_theme = detect_or_get_template(req.template_id, req.prompt, profile.activity, sector_id=profile.sector_id)
        fin_data["template_id"] = selected_theme.id
        fin_data["theme"] = selected_theme

        # 6. Génération du classeur Excel natif avec openpyxl (20 feuilles 100% dynamiques et équilibrées)
        excel_bytes = ExcelReportGenerator.build(fin_data, project_name=project_name, country=country, template_id=selected_theme.id)
        b64_file = base64.b64encode(excel_bytes).decode("utf-8")

        # 7. Synthèse exécutive professionnelle
        banner = f"🔄 **Business Plan Modifié & Recalculé avec Succès : {project_name}**" if req.is_modification else f"✅ **Business Plan Bancable OHADA généré avec succès pour : {project_name}**"
        summary_text = (
            f"{banner}\n\n"
            f"🎨 **Design & Template :** {selected_theme.icon} **{selected_theme.name}** ({selected_theme.subtitle})\n"
            f"📍 **Implantation :** {country} ({parsed.get('location', country)}) | **Devise :** {currency} | **Horizon Temporel :** {extracted_years} ans\n"
            f"💼 **Investissement Total (Besoins) :** {kpis.total_investment:,.0f} {currency} "
            f"(Apport promoteur {kpis.equity_pct:.0f}% : {kpis.equity_promoter:,.0f} {currency} | Emprunt bancaire {kpis.loan_pct:.0f}% : {kpis.bank_loan:,.0f} {currency} @ {kpis.loan_interest_rate:.1f}% sur {loan_duration_years} ans)\n\n"
            f"📊 **Indicateurs Clés de Performance (Calculés selon normes bancaires UEMOA/CEMAC) :**\n"
            f"- **Chiffre d'Affaires Année 1 :** {kpis.revenue_year1:,.0f} {currency}\n"
            f"- **Excédent Brut d'Exploitation (EBE) :** {kpis.ebitda_year1:,.0f} {currency}\n"
            f"- **Bénéfice Net Année 1 :** {kpis.net_profit_year1:,.0f} {currency}\n"
            f"- **Seuil de Rentabilité :** {kpis.breakeven_revenue:,.0f} {currency} (Atteint au jour {kpis.breakeven_days})\n"
            f"- **Rentabilité Financière Réelle :** TRI de **{kpis.irr_pct}** | VAN (à 10%) de **{kpis.npv:,.0f} {currency}** | Récupération en **{kpis.payback_period}**\n"
            f"- **Indice de Profitabilité (IP) :** {kpis.profitability_index:.2f} | **DSCR Moyen :** {kpis.dscr_avg:.2f}x\n"
            f"- **Emplois créés :** {kpis.jobs_created} salariés permanents\n\n"
            f"📁 **20 feuilles professionnelles prêtes à l'emploi (Modélisation complète sur {extracted_years} ans) :**\n"
            f"Page de garde, Présentation & Vision à {extracted_years} ans, Étude de marché, Mix 4P, Itinéraire technique, Prévisions de production ({extracted_years} ans), "
            f"Chiffre d'affaires ({extracted_years} ans), Investissements & Amortissements, Charges SYSCOHADA, Salaires & CNPS, Plan de financement ({kpis.equity_pct:.0f}% / {kpis.loan_pct:.0f}%), "
            f"Remboursement emprunt {kpis.loan_interest_rate:.1f}%, BFR normatif, Compte de résultat SIG ({extracted_years} ans), Bilan prévisionnel équilibré (5 ans), "
            f"Plan de trésorerie (12M & 5 ans), Seuil de rentabilité, Stress-test sensibilité, Ratios bancaires et Tableau de bord exécutif."
        ).replace(",", " ")

        filename = f"BusinessPlan_{profile.sector_id.capitalize()}_{extracted_years}Ans_{country.replace(' ', '_')}.xlsx"

        return PlanResponse(
            success=True,
            project_name=project_name,
            sector=profile.name,
            country=country,
            currency=currency,
            summary=summary_text,
            kpis=kpis,
            sheets_count=20,
            filename=filename,
            file_base64=b64_file,
            template_id=selected_theme.id,
            template_name=selected_theme.name
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Erreur de modélisation financière : {str(e)}")

@app.post("/api/download-excel")
def download_excel(req: PlanRequest):
    """Endpoint direct pour télécharger le binaire XLSX."""
    profile = detect_sector(req.prompt)
    project_name = req.company_name or f"Projet {profile.name}"
    currency = req.currency or "FCFA"
    fin_data = FinancialEngine.compute(profile, currency=currency)
    country = req.country or "Côte d'Ivoire"
    excel_bytes = ExcelReportGenerator.build(fin_data, project_name=project_name, country=country)

    filename = f"BusinessPlan_{profile.sector_id.capitalize()}.xlsx"
    return Response(
        content=excel_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )
