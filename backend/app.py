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
        import re
        country = req.country or "Côte d'Ivoire"
        currency = req.currency or "FCFA"

        # 1. Extraction NLP chirurgicale du budget et de l'horizon dans la requête utilisateur
        extracted_budget = req.total_budget
        if not extracted_budget:
            m_budget = re.search(r'(\d+(?:[.,]\d+)?)\s*(?:millions?|million|m|mfcfa)\b', req.prompt, re.I)
            if m_budget:
                extracted_budget = float(m_budget.group(1).replace(',', '.')) * 1_000_000
            else:
                m_num = re.search(r'(\d[\d\s]{4,})\s*(?:fcfa|cfa|f|€|\$)?\b', req.prompt, re.I)
                if m_num:
                    clean_digits = re.sub(r'\s+', '', m_num.group(1))
                    if clean_digits.isdigit() and int(clean_digits) >= 100_000:
                        extracted_budget = float(clean_digits)

        extracted_years = req.horizon_years or 5
        m_years = re.search(r'(\d+)\s*ans?\b', req.prompt, re.I)
        if m_years:
            parsed_y = int(m_years.group(1))
            if 2 <= parsed_y <= 15:
                extracted_years = parsed_y

        # 2. Nettoyage chirurgical du sujet pour un titre professionnel sans bugs de mots écourtés
        clean_p = req.prompt.strip()
        clean_topic = None

        # Si le prompt combine "Sujet : Instruction de modification" (ex: "Plantation de cacao : modifie l'investissement à 15 millions")
        if ":" in clean_p:
            parts = clean_p.split(":", 1)
            candidate = parts[0].strip()
            if candidate and not re.match(r'^(?:modifie|change|ajuste|adapte|recalcule|passe|mets?)\b', candidate, re.I):
                clean_topic = candidate
                clean_p = parts[1].strip()

        if not clean_topic:
            clean_p = re.sub(r'^(?:je\s+veux|je\s+souhaite|fais[- ]moi|fait[- ]moi|peux[- ]tu\s+faire|cr[ée]e[- ]moi|g[ée]n[èe]re[- ]moi|donne[- ]moi|il\s+me\s+faut|modifie|change|ajuste|adapte|recalcule|passe|mets?)\s+(?:le\s+)?(?:montant\s+d[\'’]investissement|budget|plan|projet|l[\'’]emprunt)?\s*(?:à|au|de|pour|en)?\s*', '', clean_p, flags=re.I)
            clean_p = re.sub(r'^(?:un|le|mon|du)?\s*(?:business\s*plan|plan\s*d[\'’]affaires?|projet)\s+(?:complet|bancable|d[ée]taill[ée])?\s*', '', clean_p, flags=re.I)
            clean_p = re.sub(r'^(?:pour|sur|concernant)\s+', '', clean_p, flags=re.I)
            clean_p = re.sub(r'^(?:mon|ma|mes|notre|nos|un|une|des|le|la|les)\s+', '', clean_p, flags=re.I)
            clean_p = re.sub(r'\s+(?:sur\s+\d+\s*ans?|avec\s+(?:un\s+)?budget|\s*avec\s+\d+|j[\'’]ai\s+|\bà\s+[A-Za-zÀ-ÿ]+|\ben\s+[A-Za-zÀ-ÿ]+|\bau\s+[A-Za-zÀ-ÿ]+).*$', '', clean_p, flags=re.I)
            clean_topic = clean_p.strip(' ,;.')

        # Prise en compte du sujet précédent si modification
        if getattr(req, "previous_topic", None) and (not clean_topic or clean_topic.lower() in ["projet", "mon projet", "entreprise", "plan", "business plan"]):
            clean_topic = req.previous_topic

        if clean_topic:
            clean_topic = clean_topic[0].upper() + clean_topic[1:]
        else:
            clean_topic = "Plantation de Cacao"

        # 3. Moteur IA Gemini 3.6 Flash : Conception dynamique sur-mesure pour chaque projet spécifique
        dynamic_profile = GeminiSectorEngine.generate_dynamic_sector(
            user_prompt=req.prompt,
            country=country,
            currency=currency,
            target_budget=extracted_budget,
            horizon_years=extracted_years,
            clean_topic=clean_topic
        )

        # Si Gemini a conçu le profil sur mesure, on l'utilise directement ; sinon fallback catalogue expert
        if dynamic_profile:
            profile = dynamic_profile
        else:
            profile = detect_sector(req.prompt)

        # 4. Nom valorisant du projet
        project_name = req.company_name or f"Projet - {clean_topic}"

        # 5. Calcul rigoureux du modèle financier OHADA avec respect strict du budget et de l'horizon
        fin_data = FinancialEngine.compute(
            profile,
            currency=currency,
            horizon_years=extracted_years,
            target_budget=extracted_budget,
            country=country
        )
        kpis: KPIOverview = fin_data["kpis"]

        # 6. Choix et application du Template de Design
        selected_theme = detect_or_get_template(req.template_id, req.prompt, profile.activity)
        fin_data["template_id"] = selected_theme.id
        fin_data["theme"] = selected_theme

        # 7. Génération du classeur Excel natif avec openpyxl (20 feuilles dynamiques avec graphiques)
        excel_bytes = ExcelReportGenerator.build(fin_data, project_name=project_name, country=country, template_id=selected_theme.id)
        b64_file = base64.b64encode(excel_bytes).decode("utf-8")

        # 8. Synthèse exécutive professionnelle
        banner = f"🔄 **Business Plan Modifié & Recalculé avec Succès : {project_name}**" if req.is_modification else f"✅ **Business Plan Bancable OHADA généré avec succès pour : {project_name}**"
        summary_text = (
            f"{banner}\n\n"
            f"🎨 **Design & Template :** {selected_theme.icon} **{selected_theme.name}** ({selected_theme.subtitle})\n"
            f"📍 **Implantation :** {country} | **Devise :** {currency} | **Horizon Temporel :** {extracted_years} ans\n"
            f"💼 **Investissement Total (Besoins) :** {kpis.total_investment:,.0f} {currency} "
            f"(Apport promoteur 40% : {kpis.equity_promoter:,.0f} {currency} | Emprunt bancaire 60% : {kpis.bank_loan:,.0f} {currency})\n\n"
            f"📊 **Indicateurs Clés de Performance :**\n"
            f"- **Chiffre d'Affaires Année 1 :** {kpis.revenue_year1:,.0f} {currency}\n"
            f"- **Excédent Brut d'Exploitation (EBE) :** {kpis.ebitda_year1:,.0f} {currency}\n"
            f"- **Bénéfice Net Année 1 :** {kpis.net_profit_year1:,.0f} {currency}\n"
            f"- **Seuil de Rentabilité :** {kpis.breakeven_revenue:,.0f} {currency} (Atteint au jour {kpis.breakeven_days})\n"
            f"- **Rentabilité :** TRI de **{kpis.irr_pct}** | VAN de **{kpis.npv:,.0f} {currency}** | Récupération en **{kpis.payback_period}**\n"
            f"- **Emplois créés :** {kpis.jobs_created} salariés permanents\n\n"
            f"📁 **20 feuilles professionnelles prêtes à l'emploi (Modélisation complète sur {extracted_years} ans) :**\n"
            f"Page de garde, Présentation & Vision à {extracted_years} ans, Étude de marché, Mix 4P, Itinéraire technique, Prévisions de production ({extracted_years} ans), "
            f"Chiffre d'affaires ({extracted_years} ans), Investissements & Amortissements, Charges SYSCOHADA, Salaires & CNPS, Plan de financement, Remboursement emprunt 8%, "
            f"BFR normatif, Compte de résultat SIG ({extracted_years} ans), Bilan Actif/Passif, Trésorerie 12M, Seuil de rentabilité, "
            f"Stress-test sensibilité, Ratios bancaires et Tableau de bord exécutif."
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
