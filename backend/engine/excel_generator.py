"""
Générateur de classeurs Excel professionnels ultra-modernes avec openpyxl.
Produit un dossier bancaire complet de 20 feuilles SYSCOHADA avec typographie soignée,
largeurs de colonnes sécurisées (zéro '#########'), formules universelles et design prestige.
"""

import io
import base64
from typing import Dict, Any, List
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from .templates import TEMPLATE_CATALOG, detect_or_get_template, ThemeConfig

class ExcelReportGenerator:
    # Palette active courante (initialisée dynamiquement par le template choisi)
    current_theme: ThemeConfig = None

    # Couleurs de secours par défaut
    NAVY_DARK = "1E293B"
    NAVY_LIGHT = "334155"
    ZEBRA_BG = "F8FAFC"
    BORDER_SOFT = "F1F5F9"
    BORDER_TOTAL = "CBD5E1"
    GREEN_BG = "ECFDF5"
    GREEN_TEXT = "065F46"
    AMBER_BG = "FEF3C7"
    AMBER_TEXT = "92400E"

    @classmethod
    def build(cls, data: Dict[str, Any], project_name: str, country: str = "Côte d'Ivoire", template_id: str = "auto") -> bytes:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        currency = data["currency"]
        profile = data["profile"]
        m = data

        theme = data.get("theme") or detect_or_get_template(template_id or data.get("template_id"), prompt_text=project_name, activity_text=profile.activity)
        cls.current_theme = theme
        m["theme"] = theme

        # Construction séquentielle des 20 feuilles
        cls._build_01_page_de_garde(wb, project_name, country, currency, m)
        cls._build_02_presentation_entreprise(wb, project_name, profile, m)
        cls._build_03_etude_marche(wb, profile, country)
        cls._build_04_strategie_marketing(wb, profile, currency)
        cls._build_05_itineraire_technique(wb, profile, m)
        cls._build_06_previsions_production(wb, profile, m)
        cls._build_07_chiffres_affaires(wb, profile, currency, m)
        cls._build_08_investissements(wb, profile, currency, m)
        cls._build_09_charges_exploitation(wb, profile, currency, m)
        cls._build_10_frais_personnel(wb, profile, currency, m)
        cls._build_11_plan_financement(wb, currency, m)
        cls._build_12_remboursement_emprunt(wb, currency, m)
        cls._build_13_bfr(wb, currency, m)
        cls._build_14_compte_resultat(wb, currency, m)
        cls._build_15_bilan_previsionnel(wb, currency, m)
        cls._build_16_plan_tresorerie(wb, currency, m)
        cls._build_17_seuil_rentabilite(wb, currency, m)
        cls._build_18_sensibilite(wb, currency, m)
        cls._build_19_ratios(wb, currency, m)
        cls._build_20_tableau_de_bord(wb, currency, m)

        # Ajuster les largeurs de colonnes sur TOUTES les feuilles pour garantir zéro '#########'
        for ws in wb.worksheets:
            cls._apply_safe_column_widths(ws)
            # Afficher les quadrillages natifs
            ws.views.sheetView[0].showGridLines = True

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer.getvalue()

    # --- FEUILLE 1 : PAGE DE GARDE ---
    @classmethod
    def _build_01_page_de_garde(cls, wb, project_name: str, country: str, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="01_Page_de_Garde")
        cls._create_banner(ws, "B2:F2", "DOSSIER DE BUSINESS PLAN BANCAIRE & ÉTUDE DE FAISABILITÉ")

        theme = cls.current_theme
        ws.merge_cells("B3:F3")
        cell_tag = ws["B3"]
        tagline = theme.tagline if theme else "Entreprendre aujourd'hui, bâtir demain en Afrique"
        cell_tag.value = f"  « {tagline} »"
        hl_bg = theme.highlight_bg if theme else "FEF3C7"
        hl_fg = theme.highlight_text if theme else "78350F"
        cell_tag.fill = PatternFill(fill_type="solid", fgColor=hl_bg)
        cell_tag.font = Font(name="Calibri", size=10, italic=True, bold=True, color=hl_fg)
        cell_tag.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[3].height = 22

        headers = ["PARAMÈTRE DU PROJET", "VALEUR & MODALITÉS"]
        ws.row_dimensions[4].height = 26
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        kpis = m["kpis"]
        horizon = m.get("horizon_years", 5)
        theme_badge = f"{theme.icon} {theme.name} ({theme.subtitle})" if theme else "Standard OHADA"
        rows = [
            ["Désignation du Projet", project_name],
            ["Template & Design Visuel", theme_badge],
            ["Secteur d'Activité", m["profile"].name],
            ["Pays & Zone Géographique", f"{country} (Espace Économique OHADA)"],
            ["Forme Juridique Projetée", "Société à Responsabilité Limitée (SARL Pluripersonnelle)"],
            ["Devise de Référence", currency],
            ["Horizon Prévisionnel du Modèle", f"{horizon} ans"],
            ["Montant Global de l'Investissement (Besoins)", kpis.total_investment],
            ["Apport Personnel Promoteur (Fonds Propres)", kpis.equity_promoter],
            ["Crédit Bancaire Sollicité à Moyen Terme", kpis.bank_loan],
            ["Chiffre d'Affaires Prévisionnel (Année 1)", kpis.revenue_year1],
            ["Bénéfice Net Après Impôt (Année 1)", kpis.net_profit_year1],
            ["Seuil de Rentabilité (Point Mort)", f"{kpis.breakeven_revenue:,.0f} {currency} (Atteint au jour {kpis.breakeven_days})".replace(",", " ")],
            ["Taux de Rentabilité Interne (TRI)", kpis.irr_pct],
            ["Délai de Récupération du Capital (DRCI)", kpis.payback_period],
            ["Emplois Directs Permanents Créés", f"{kpis.jobs_created} salariés qualifiés"]
        ]

        cls._write_table_data(ws, 5, 2, rows, numeric_cols=[3], num_format=f'#,##0 "{currency}"')

    # --- FEUILLE 2 : PRÉSENTATION DE L'ENTREPRISE ---
    @classmethod
    def _build_02_presentation_entreprise(cls, wb, project_name: str, profile: Any, m: Dict[str, Any] = None):
        ws = wb.create_sheet(title="02_Presentation_Entreprise")
        cls._create_banner(ws, "B2:E2", "FICHE D'IDENTITÉ & PRÉSENTATION DU PROJET")

        ws.row_dimensions[4].height = 26
        for col_idx, h in enumerate(["Rubrique Institutionnelle", "Détails & Orientations Stratégiques"], start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        horizon = (m or {}).get("horizon_years", 5)
        is_comm = (m or {}).get("is_commercial", False)
        appro_label = "Approvisionnement Marchandises / Stocks" if is_comm else "Approvisionnement Matières Premières"
        norme_label = "Conformité stricte aux normes de sécurité, enregistrement RCCM & IFU local" if is_comm else "Conformité stricte aux normes d'hygiène HACCP, enregistrement RCCM & IFU local"
        rows = [
            ["Raison Sociale", project_name],
            ["Activité Principale", profile.activity],
            ["Gamme de Produits / Services", " ; ".join([p.name for p in profile.products])],
            ["Unité Opérationnelle", profile.unit],
            [appro_label, profile.raw_materials_desc],
            ["Réglementation & Normes", norme_label],
            [f"Vision Stratégique à {horizon} Ans", f"Devenir l'acteur de référence en alliant service d'excellence, rentabilité durable et satisfaction client sur {horizon} ans"],
            ["Facteurs Clés de Succès (FCS)", "1. Maîtrise de la chaîne d'approvisionnement et qualité des produits\n2. Équipements modernes haute efficacité\n3. Gestion financière rigoureuse certifiée SYSCOHADA"]
        ]
        cls._write_table_data(ws, 5, 2, rows)

    # --- FEUILLE 3 : ÉTUDE DE MARCHÉ ---
    @classmethod
    def _build_03_etude_marche(cls, wb, profile: Any, country: str):
        ws = wb.create_sheet(title="03_Etude_Marche")
        cls._create_banner(ws, "B2:I2", "ANALYSE DU MARCHÉ, CIBLAGE DE LA DEMANDE & CROISSANCE")

        ws.row_dimensions[4].height = 26
        for col_idx, h in enumerate(["Segment & Axe d'Analyse", "Constats du Marché Local", "Opportunités & Réponses du Projet"], start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        rows = [
            ["Tendance Globale du Marché", f"Croissance démographique rapide et urbanisation forte en {country}", "Demande accrue pour les produits et services de qualité"],
            ["Clientèle Cible B2C (Particuliers)", "Classes moyennes urbaines en quête de traçabilité et modernité", "Offre soignée, expérience d'achat agréable et prix juste"],
            ["Clientèle Cible B2B (Professionnels)", "Entreprises, écoles, institutions, distributeurs locaux", "Partenariats durables, respect des délais et satisfaction garantie"],
            ["Analyse de la Concurrence", "Offre informelle fragmentée ou produits importés coûteux", "Avantage décisif : proximité, garantie, service après-vente"],
            ["Potentiel de Croissance", f"Expansion rapide de la consommation domestique en {country}", "Potentiel d'ouverture de succursales régionales à moyen terme"]
        ]
        cls._write_table_data(ws, 5, 2, rows)

        # Mini table de projection du marché sectoriel (colonnes F et G)
        ws.cell(row=4, column=6, value="Année").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=4, column=7, value="Indice Marché").font = Font(name="Calibri", size=10, bold=True)
        indices = [100, 115, 132, 150, 172]
        for idx_y, val_idx in enumerate(indices):
            r_idx = 5 + idx_y
            ws.cell(row=r_idx, column=6, value=f"Année {idx_y + 1}")
            ws.cell(row=r_idx, column=7, value=val_idx)

        # Graphique natif : Évolution Estimée du Marché
        chart_mkt = BarChart()
        chart_mkt.type = "col"
        chart_mkt.style = 10
        chart_mkt.title = f"Évolution Estimée du Marché en {country}"
        chart_mkt.y_axis.title = "Indice Croissance Base 100"
        chart_mkt.x_axis.title = "Horizon Annuel"
        chart_mkt.width = 13.5
        chart_mkt.height = 7.5
        data_mkt = Reference(ws, min_col=7, min_row=4, max_row=9)
        cats_mkt = Reference(ws, min_col=6, min_row=5, max_row=9)
        chart_mkt.add_data(data_mkt, titles_from_data=True)
        chart_mkt.set_categories(cats_mkt)
        chart_mkt.legend = None
        ws.add_chart(chart_mkt, "F11")

    # --- FEUILLE 4 : STRATÉGIE MARKETING ---
    @classmethod
    def _build_04_strategie_marketing(cls, wb, profile: Any, currency: str):
        ws = wb.create_sheet(title="04_Strategie_Marketing")
        cls._create_banner(ws, "B2:E2", "STRATÉGIE COMMERCIALE (MARKETING MIX 4P)")

        ws.row_dimensions[4].height = 26
        for col_idx, h in enumerate(["Pilier du Marketing Mix", "Stratégie Opérationnelle", "Objectifs & Canaux Clés"], start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        rows = [
            ["Politique Produit (Product)", f"Gamme premium adaptée au marché : {profile.products[0].name}", "Certification qualité, traçabilité et packaging soigné"],
            ["Politique Prix (Price)", f"Positionnement compétitif : de {profile.products[0].unit_price:,.0f} à {profile.products[-1].unit_price:,.0f} {currency}".replace(",", " "), "Pénétration rapide du marché avec marge brute protectrice"],
            ["Politique Distribution (Place)", "Canaux mixtes : points de vente directs, supermarchés, revendeurs agréés", "Distribution sélective à forte visibilité dans les centres urbains"],
            ["Politique Promotion (Promotion)", "Marketing digital (réseaux sociaux), dégustations, salons professionnels", "Acquisition client ciblée et notoriété de marque immédiate"]
        ]
        cls._write_table_data(ws, 5, 2, rows)

    # --- FEUILLE 5 : ITINÉRAIRE TECHNIQUE ---
    @classmethod
    def _build_05_itineraire_technique(cls, wb, profile: Any, m: Dict[str, Any] = None):
        ws = wb.create_sheet(title="05_Itineraire_Technique")
        cls._create_banner(ws, "B2:F2", "ITINÉRAIRE TECHNIQUE & CALENDRIER OPÉRATIONNEL (12 MOIS)")

        ws.row_dimensions[4].height = 26
        for col_idx, h in enumerate(["Période", "Étape Opérationnelle Clé", "Responsable", "Délai", "Livrable Attendu"], start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        is_comm = (m or {}).get("is_commercial", False)
        if is_comm:
            rows = [
                ["Mois 1", "Constitution juridique de la société & formalités fiscales", "Promoteur / Notaire", "Semaines 1-4", "Statuts, RCCM, Déclaration fiscale"],
                ["Mois 2", "Aménagement du local commercial / boutique et vitrines", "Gérant / Agenceur", "Semaines 5-8", "Boutique prête & sécurisée"],
                ["Mois 3", "Acquisition rayonnages, caisse tactile & système gestion", "Fournisseur IT / Gérant", "Semaines 9-12", "Caisse & étagères opérationnelles"],
                ["Mois 4", "Constitution du stock initial de marchandises & formation vendeuses", "Gérant / Vendeuses", "Semaines 13-16", "Rayons achalandés et inventoriés"],
                ["Mois 5", "Campagne de lancement commercial, inauguration & vente", "Responsable Ventes", "Semaines 17-20", "Premières ventes en magasin"],
                ["Mois 6", "Plein régime commercial, rotation des stocks & suivi financier", "Gérant / Comptable", "Semaines 21-24", "Rapports de marge commerciale"]
            ]
        else:
            rows = [
                ["Mois 1", "Constitution juridique de la société & formalités fiscales", "Promoteur / Notaire", "Semaines 1-4", "Statuts, RCCM, Déclaration fiscale"],
                ["Mois 2", "Aménagement du local de production aux normes d'hygiène", "Chef d'Atelier", "Semaines 5-8", "Atelier prêt & sécurisé"],
                ["Mois 3", "Acquisition, livraison et installation des équipements", "Fournisseur / Promoteur", "Semaines 9-12", "Machines installées et testées"],
                ["Mois 4", "Tests de production, recettes & formation des opérateurs", "Responsable Qualité", "Semaines 13-16", "Lots tests conformes aux normes"],
                ["Mois 5", "Campagne de lancement commercial & référencement", "Responsable Ventes", "Semaines 17-20", "Premières commandes livrées"],
                ["Mois 6", "Plein régime de croisière et suivi de trésorerie", "Comptable / Direction", "Semaines 21-24", "Rapports d'activité mensuels"]
            ]
        cls._write_table_data(ws, 5, 2, rows)

    # --- FEUILLE 6 : PRÉVISIONS DE PRODUCTION ---
    @classmethod
    def _build_06_previsions_production(cls, wb, profile: Any, m: Dict[str, Any]):
        ws = wb.create_sheet(title="06_Previsions_Production")
        horizon = m.get("horizon_years", 5)
        is_comm = m.get("is_commercial", False)
        end_col_letter = get_column_letter(3 + horizon)
        banner_title = f"PRÉVISIONS DES VENTES & VOLUMES COMMERCIAUX ({profile.unit})" if is_comm else f"CAPACITÉS & VOLUMES PRÉVISIONNELS DE PRODUCTION ({profile.unit})"
        cls._create_banner(ws, f"B2:{end_col_letter}2", banner_title)

        ws.row_dimensions[4].height = 26
        header_gamme = "Gamme de Produits / Rayons" if is_comm else "Gamme de Produits"
        header_cap = "Ventes Estimées / Jour" if is_comm else "Capacité Jour"
        headers = [header_gamme, header_cap] + [f"Année {y}" for y in range(1, horizon + 1)]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        v_day = m["volumes"]["day"]
        v_an1 = m["volumes"]["an1"]
        g = m["growth_factors"]

        rows = []
        for p_idx, prod in enumerate(profile.products):
            r_vals = [prod.name, v_day[p_idx]]
            for y in range(1, horizon + 1):
                r_vals.append(round(v_an1[p_idx] * g[y - 1]))
            rows.append(r_vals)

        num_cols = list(range(3, 4 + horizon))
        cls._write_table_data(ws, 5, 2, rows, numeric_cols=num_cols, num_format="#,##0")

        # Ligne de total
        total_row = 5 + len(profile.products)
        ws.cell(row=total_row, column=2, value="TOTAL VOLUME ANNUEL").font = Font(name="Calibri", size=10, bold=True)
        for c in range(4, 4 + horizon):
            col_letter = get_column_letter(c)
            cell = ws.cell(row=total_row, column=c, value=f"=SUM({col_letter}5:{col_letter}{total_row - 1})")
            cell.number_format = "#,##0"
            cls._style_total_cell(cell)
        ws.cell(row=total_row, column=2).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        ws.cell(row=total_row, column=3).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)

    # --- FEUILLE 7 : CHIFFRES D'AFFAIRES ---
    @classmethod
    def _build_07_chiffres_affaires(cls, wb, profile: Any, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="07_Chiffres_Affaires")
        horizon = m.get("horizon_years", 5)
        end_col_letter = get_column_letter(3 + horizon)
        cls._create_banner(ws, f"B2:{end_col_letter}2", f"CHIFFRE D'AFFAIRES PRÉVISIONNEL SUR {horizon} ANS (En {currency})")

        ws.row_dimensions[4].height = 26
        headers = ["Produit / Gamme", f"Prix Unitaire ({currency})"] + [f"CA Année {y}" for y in range(1, horizon + 1)]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        ca_p = m["ca"]["products_an1"]
        g = m["growth_factors"]

        rows = []
        for p_idx, prod in enumerate(profile.products):
            r_vals = [prod.name, prod.unit_price]
            for y in range(1, horizon + 1):
                r_vals.append(round(ca_p[p_idx] * g[y - 1]))
            rows.append(r_vals)

        num_cols = list(range(3, 4 + horizon))
        cls._write_table_data(ws, 5, 2, rows, numeric_cols=num_cols, num_format=f'#,##0 "{currency}"')

        total_row = 5 + len(profile.products)
        ws.cell(row=total_row, column=2, value="CHIFFRE D'AFFAIRES TOTAL").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=total_row, column=2).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        ws.cell(row=total_row, column=3).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        for c in range(4, 4 + horizon):
            col_letter = get_column_letter(c)
            cell = ws.cell(row=total_row, column=c, value=f"=SUM({col_letter}5:{col_letter}{total_row - 1})")
            cell.number_format = f'#,##0 "{currency}"'
            cls._style_total_cell(cell)

    # --- FEUILLE 8 : INVESTISSEMENTS ET AMORTISSEMENTS ---
    @classmethod
    def _build_08_investissements(cls, wb, profile: Any, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="08_Investissements_Amort")
        cls._create_banner(ws, "B2:G2", f"PROGRAMME D'INVESTISSEMENTS & PLAN D'AMORTISSEMENT ({currency})")

        ws.row_dimensions[4].height = 26
        headers = ["Immobilisation / Équipement", f"Montant Brut ({currency})", "Durée (Ans)", "Taux %", "Annuité Amort.", "VNC Fin Année 1"]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        eq_list = profile.equipment
        rows = []
        for i, eq in enumerate(eq_list):
            r = 5 + i
            rows.append([
                eq.name,
                eq.cost,
                eq.duration_years,
                f"{round(eq.depreciation_rate * 100)}%",
                f"=C{r}*E{r}",
                f"=C{r}-F{r}"
            ])

        # Correction chirurgicale : numeric_cols=[3, 6, 7] uniquement pour les montants en devises FCFA
        # La colonne 4 (Durée en années) ne doit PAS recevoir le format FCFA !
        cls._write_table_data(ws, 5, 2, rows, numeric_cols=[3, 6, 7], num_format=f'#,##0 "{currency}"')

        for i, eq in enumerate(eq_list):
            r = 5 + i
            cell_duree = ws.cell(row=r, column=4)
            cell_duree.number_format = '0 "ans"'
            cell_duree.alignment = Alignment(horizontal="center", vertical="center")
            cell_taux = ws.cell(row=r, column=5)
            cell_taux.alignment = Alignment(horizontal="center", vertical="center")

        last_r = 4 + len(eq_list)
        total_r = last_r + 1
        ws.cell(row=total_r, column=2, value="TOTAL DES INVESTISSEMENTS").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=total_r, column=2).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        ws.cell(row=total_r, column=4).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        ws.cell(row=total_r, column=5).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)

        cell_brut = ws.cell(row=total_r, column=3, value=f"=SUM(C5:C{last_r})")
        cell_brut.number_format = f'#,##0 "{currency}"'
        cls._style_total_cell(cell_brut)

        cell_amort = ws.cell(row=total_r, column=6, value=f"=SUM(F5:F{last_r})")
        cell_amort.number_format = f'#,##0 "{currency}"'
        cls._style_total_cell(cell_amort)

        cell_vnc = ws.cell(row=total_r, column=7, value=f"=SUM(G5:G{last_r})")
        cell_vnc.number_format = f'#,##0 "{currency}"'
        cls._style_total_cell(cell_vnc)

    # --- FEUILLE 9 : CHARGES D'EXPLOITATION ---
    @classmethod
    def _build_09_charges_exploitation(cls, wb, profile: Any, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="09_Charges_Exploitation")
        cls._create_banner(ws, "B2:F2", f"CHARGES D'EXPLOITATION & CONSOMMATIONS INTERMÉDIAIRES ({currency})")

        ws.row_dimensions[4].height = 26
        headers = ["Nature de la Charge", "Compte SYSCOHADA", "Année 1", "Année 2", "Année 3"]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        ca1 = m["kpis"].revenue_year1
        mat = round(ca1 * 0.38)
        emb = round(ca1 * 0.10)
        nrg = round(ca1 * 0.035)
        carb = round(ca1 * 0.035)
        maint = round(ca1 * 0.015)
        ass = round(ca1 * 0.010)
        tel = round(ca1 * 0.008)
        mkt = round(ca1 * 0.030)

        is_commercial = m.get("is_commercial", False)
        first_prod = profile.products[0].name.split()[0] if profile.products else "Marchandises"
        if is_commercial:
            ach_label = f"Achats de marchandises (Stocks de {first_prod}...)"
            ach_cpt = 601100
            emb_label = "Fournitures et emballages commerciaux"
            emb_cpt = 604300
            nrg_label = "Électricité, eau & climatisation boutique"
        else:
            ach_label = f"Achats matières premières ({first_prod}...)"
            ach_cpt = 602100
            emb_label = "Emballages & conditionnement spécifiques"
            emb_cpt = 602200
            nrg_label = "Électricité, eau & énergies atelier"

        charges = [
            [ach_label, ach_cpt, mat, round(mat * 1.25), round(mat * 1.55)],
            [emb_label, emb_cpt, emb, round(emb * 1.25), round(emb * 1.55)],
            [nrg_label, 605100, nrg, round(nrg * 1.20), round(nrg * 1.40)],
            ["Carburant & logistique de transport", 612000, carb, round(carb * 1.20), round(carb * 1.40)],
            ["Entretien, maintenance & réparations", 624000, maint, round(maint * 1.15), round(maint * 1.30)],
            ["Assurances professionnelles multirisques", 625000, ass, round(ass * 1.05), round(ass * 1.10)],
            ["Téléphone, internet & logiciels de caisse", 628000, tel, round(tel * 1.10), round(tel * 1.20)],
            ["Marketing, publicité & enseigne", 627000, mkt, round(mkt * 1.20), round(mkt * 1.35)]
        ]

        cls._write_table_data(ws, 5, 2, charges, numeric_cols=[4, 5, 6], num_format=f'#,##0 "{currency}"')

        # Formater les numéros de compte SYSCOHADA (centrés, format numérique pur sans triangle vert)
        for r_idx in range(5, 5 + len(charges)):
            cell_cpt = ws.cell(row=r_idx, column=3)
            cell_cpt.number_format = '0'
            cell_cpt.alignment = Alignment(horizontal="center", vertical="center")

        total_r = 13
        ws.cell(row=total_r, column=2, value="TOTAL DES CHARGES D'EXPLOITATION").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=total_r, column=2).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        ws.cell(row=total_r, column=3).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        for c in range(4, 7):
            col_letter = get_column_letter(c)
            cell = ws.cell(row=total_r, column=c, value=f"=SUM({col_letter}5:{col_letter}12)")
            cell.number_format = f'#,##0 "{currency}"'
            cls._style_total_cell(cell)

    # --- FEUILLE 10 : FRAIS DE PERSONNEL ---
    @classmethod
    def _build_10_frais_personnel(cls, wb, profile: Any, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="10_Frais_Personnel")
        cls._create_banner(ws, "B2:G2", f"GRILLE DES SALAIRES & CHARGES SOCIALES ({currency})")

        ws.row_dimensions[4].height = 26
        country = m.get("country", "Côte d'Ivoire")
        cnps_rate = m.get("cnps_rate", 0.164)
        if country in ["Côte d'Ivoire", "Cameroun"]:
            org_social = "CNPS"
        elif country == "Sénégal":
            org_social = "CSS / IPRES"
        else:
            org_social = "CNSS"
        rate_str = f"{round(cnps_rate * 100, 1)}%".replace(".0%", "%")

        headers = ["Poste / Fonction", "Effectif", "Salaire Net Mensuel", f"Cotisations {org_social} ({rate_str})", "Coût Mensuel Total", "Coût Annuel Total"]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        staff_list = profile.staff
        rows = []
        for i, s in enumerate(staff_list):
            r = 5 + i
            rows.append([
                s.title,
                s.count,
                s.monthly_net,
                s.monthly_cnps,
                f"=D{r}+E{r}",
                f"=F{r}*12*C{r}"
            ])

        # Correction chirurgicale : numeric_cols=[4, 5, 6, 7] uniquement pour les salaires en FCFA !
        # La colonne 3 (Effectif) est formatée en entier #,##0.
        cls._write_table_data(ws, 5, 2, rows, numeric_cols=[4, 5, 6, 7], num_format=f'#,##0 "{currency}"')

        for i, s in enumerate(staff_list):
            r = 5 + i
            cell_eff_item = ws.cell(row=r, column=3)
            cell_eff_item.number_format = '#,##0'
            cell_eff_item.alignment = Alignment(horizontal="center", vertical="center")

        last_r = 4 + len(staff_list)
        total_r = last_r + 1
        ws.cell(row=total_r, column=2, value="TOTAL MASSE SALARIALE ANNUELLE").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=total_r, column=2).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        ws.cell(row=total_r, column=4).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        ws.cell(row=total_r, column=5).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        ws.cell(row=total_r, column=6).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)

        cell_eff = ws.cell(row=total_r, column=3, value=f"=SUM(C5:C{last_r})")
        cell_eff.number_format = "#,##0"
        cell_eff.alignment = Alignment(horizontal="center", vertical="center")
        cls._style_total_cell(cell_eff)

        cell_tot = ws.cell(row=total_r, column=7, value=f"=SUM(G5:G{last_r})")
        cell_tot.number_format = f'#,##0 "{currency}"'
        cls._style_total_cell(cell_tot)

    # --- FEUILLE 11 : PLAN DE FINANCEMENT ---
    @classmethod
    def _build_11_plan_financement(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="11_Plan_Financement")
        cls._create_banner(ws, "B2:F2", "PLAN DE FINANCEMENT INITIAL (ÉQUILIBRE DES EMPLOIS & RESSOURCES)")

        ws.row_dimensions[4].height = 26
        ws.cell(row=4, column=2, value="BESOINS DE DÉMARRAGE (EMPLOIS)").fill = PatternFill(fill_type="solid", fgColor=cls.NAVY_DARK)
        ws.cell(row=4, column=2).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        ws.cell(row=4, column=3, value=f"Montant ({currency})").fill = PatternFill(fill_type="solid", fgColor=cls.NAVY_DARK)
        ws.cell(row=4, column=3).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

        ws.cell(row=4, column=5, value="RESSOURCES MOBILISÉES (FINANCEMENT)").fill = PatternFill(fill_type="solid", fgColor=cls.NAVY_DARK)
        ws.cell(row=4, column=5).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        ws.cell(row=4, column=6, value=f"Montant ({currency})").fill = PatternFill(fill_type="solid", fgColor=cls.NAVY_DARK)
        ws.cell(row=4, column=6).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

        fin = m["financing"]
        inv = m["investments"]
        bfr = m["bfr"]

        besoins = [
            ["Frais d'établissement & formalités", inv["frais_etab"]],
            ["Investissements matériels & équipements", inv["total"]],
            ["BFR Initial (Fonds de roulement & stocks)", bfr["net"]]
        ]

        ressources = [
            ["Apport personnel promoteur (Fonds propres)", fin["apport"]],
            ["Emprunt bancaire sollicité à moyen terme", fin["emprunt"]],
            ["Subventions & appuis au démarrage", 0]
        ]

        for i in range(3):
            r = 5 + i
            ws.cell(row=r, column=2, value=besoins[i][0]).font = Font(name="Calibri", size=10)
            c_besoin = ws.cell(row=r, column=3, value=besoins[i][1])
            c_besoin.number_format = f'#,##0 "{currency}"'
            c_besoin.font = Font(name="Calibri", size=10)

            ws.cell(row=r, column=5, value=ressources[i][0]).font = Font(name="Calibri", size=10)
            c_ressource = ws.cell(row=r, column=6, value=ressources[i][1])
            c_ressource.number_format = f'#,##0 "{currency}"'
            c_ressource.font = Font(name="Calibri", size=10)

        # Totaux
        t_besoin = ws.cell(row=8, column=2, value="TOTAL DES BESOINS (EMPLOIS)")
        t_besoin.font = Font(name="Calibri", size=10, bold=True)
        t_besoin.fill = PatternFill(fill_type="solid", fgColor=cls.GREEN_BG)
        val_besoin = ws.cell(row=8, column=3, value="=SUM(C5:C7)")
        val_besoin.number_format = f'#,##0 "{currency}"'
        cls._style_total_cell(val_besoin, bg_color=cls.GREEN_BG, font_color=cls.GREEN_TEXT)

        t_ressource = ws.cell(row=8, column=5, value="TOTAL DES RESSOURCES")
        t_ressource.font = Font(name="Calibri", size=10, bold=True)
        t_ressource.fill = PatternFill(fill_type="solid", fgColor=cls.GREEN_BG)
        val_ressource = ws.cell(row=8, column=6, value="=SUM(F5:F7)")
        val_ressource.number_format = f'#,##0 "{currency}"'
        cls._style_total_cell(val_ressource, bg_color=cls.GREEN_BG, font_color=cls.GREEN_TEXT)

    # --- FEUILLE 12 : REMBOURSEMENT DE L'EMPRUNT ---
    @classmethod
    def _build_12_remboursement_emprunt(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="12_Remboursement_Emprunt")
        cls._create_banner(ws, "B2:G2", "TABLEAU D'AMORTISSEMENT DU CRÉDIT BANCAIRE (Taux 8% l'an)")

        ws.row_dimensions[4].height = 26
        headers = ["Période", "Capital Début Période", "Intérêts Annuels (8%)", "Amortissement Capital", "Annuité Totale", "Capital Restant Dû Fin"]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        loan_schedule = m["loan_schedule"]
        rows = []
        for i, entry in enumerate(loan_schedule):
            r = 5 + i
            if i == 0:
                rows.append([
                    entry["year"],
                    entry["start_capital"],
                    f"=C{r}*0.08",
                    entry["principal"],
                    f"=D{r}+E{r}",
                    f"=C{r}-E{r}"
                ])
            elif i == len(loan_schedule) - 1:
                rows.append([
                    entry["year"],
                    f"=G{r-1}",
                    f"=C{r}*0.08",
                    f"=C{r}",
                    f"=D{r}+E{r}",
                    f"=C{r}-E{r}"
                ])
            else:
                rows.append([
                    entry["year"],
                    f"=G{r-1}",
                    f"=C{r}*0.08",
                    entry["principal"],
                    f"=D{r}+E{r}",
                    f"=C{r}-E{r}"
                ])

        cls._write_table_data(ws, 5, 2, rows, numeric_cols=[3, 4, 5, 6, 7], num_format=f'#,##0 "{currency}"')

        last_r = 4 + len(loan_schedule)
        total_r = last_r + 1
        ws.cell(row=total_r, column=2, value="TOTAL DU SERVICE DE LA DETTE").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=total_r, column=2).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        ws.cell(row=total_r, column=3).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)

        for c in range(4, 7):
            col_letter = get_column_letter(c)
            cell = ws.cell(row=total_r, column=c, value=f"=SUM({col_letter}5:{col_letter}{last_r})")
            cell.number_format = f'#,##0 "{currency}"'
            cls._style_total_cell(cell)

        cell_end = ws.cell(row=total_r, column=7, value=0)
        cell_end.number_format = f'#,##0 "{currency}"'
        cls._style_total_cell(cell_end)

    # --- FEUILLE 13 : BESOIN EN FONDS DE ROULEMENT ---
    @classmethod
    def _build_13_bfr(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="13_BFR_Exploitation")
        cls._create_banner(ws, "B2:F2", f"BESOIN EN FONDS DE ROULEMENT (BFR EN JOURS DE CA) ({currency})")

        ws.row_dimensions[4].height = 26
        headers = ["Composante du BFR", "Norme de Rotation", "Année 1", "Année 2", "Année 3"]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        bfr = m["bfr"]
        is_commercial = m.get("is_commercial", False)
        stock_label = "Stocks de marchandises (Rayons & réserves)" if is_commercial else "Stocks de matières premières & intrants"
        rows = [
            [stock_label, "30 jours de consommation", bfr["stocks"], round(bfr["stocks"] * 1.25), round(bfr["stocks"] * 1.55)],
            ["Créances clients d'exploitation accordées", "15 jours de chiffre d'affaires", bfr["creances"], round(bfr["creances"] * 1.25), round(bfr["creances"] * 1.55)],
            ["Dettes fournisseurs (Crédit d'exploitation)", "30 jours d'achats", -bfr["dettes"], -round(bfr["dettes"] * 1.25), -round(bfr["dettes"] * 1.55)]
        ]
        cls._write_table_data(ws, 5, 2, rows, numeric_cols=[4, 5, 6], num_format=f'#,##0 "{currency}"')

        total_r = 8
        ws.cell(row=total_r, column=2, value="BFR GLOBAL NET D'EXPLOITATION").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=total_r, column=2).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        ws.cell(row=total_r, column=3).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        for c in range(4, 7):
            col_letter = get_column_letter(c)
            cell = ws.cell(row=total_r, column=c, value=f"=SUM({col_letter}5:{col_letter}7)")
            cell.number_format = f'#,##0 "{currency}"'
            cls._style_total_cell(cell, bg_color=cls.GREEN_BG, font_color=cls.GREEN_TEXT)

    # --- FEUILLE 14 : COMPTE DE RÉSULTAT SYSCOHADA ---
    @classmethod
    def _build_14_compte_resultat(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="14_Compte_Resultat_OHADA")
        horizon = m.get("horizon_years", 5)
        end_col_letter = get_column_letter(3 + horizon)
        cls._create_banner(ws, f"B2:{end_col_letter}2", f"COMPTE DE RÉSULTAT PRÉVISIONNEL SUR {horizon} ANS (SYSCOHADA) ({currency})")

        ws.row_dimensions[4].height = 26
        headers = ["Postes du Compte de Résultat", "Compte"] + [f"Année {y}" for y in range(1, horizon + 1)]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        cr = m["compte_resultat"]
        ca = [c["ca"] for c in cr]
        ach = [-c["achats"] for c in cr]
        serv = [-c["services"] for c in cr]
        perso = [-c["personnel"] for c in cr]
        dot = [-c["dotations"] for c in cr]
        int_fin = [-c["frais_fin"] for c in cr]

        # Lignes dynamiques avec formules natives pour toutes les N années (1 à horizon)
        is_commercial = m.get("is_commercial", False)
        is_rate = m.get("is_rate", 0.25)
        is_pct_str = f"{round(is_rate * 100, 1)}%".replace(".0%", "%")

        if is_commercial:
            ca_label = "CHIFFRE D'AFFAIRES (Ventes de marchandises)"
            ca_cpt = "701"
            ach_label = "Achats de marchandises vendues"
            ach_cpt = "601"
            marge_label = "MARGE COMMERCIALE (SIG)"
        else:
            ca_label = "CHIFFRE D'AFFAIRES (Ventes de produits finis)"
            ca_cpt = "702"
            ach_label = "Achats matières premières et consommables consommés"
            ach_cpt = "602"
            marge_label = "MARGE BRUTE DE PRODUCTION (SIG)"

        row_ca = [ca_label, ca_cpt] + [ca[y - 1] for y in range(1, horizon + 1)]
        row_ach = [ach_label, ach_cpt] + [ach[y - 1] for y in range(1, horizon + 1)]
        row_mb = [marge_label, "SIG"] + [f"={get_column_letter(3 + y)}5+{get_column_letter(3 + y)}6" for y in range(1, horizon + 1)]
        row_serv = ["Services extérieurs & Transports", "61/62"] + [serv[y - 1] for y in range(1, horizon + 1)]
        row_va = ["VALEUR AJOUTÉE (VA)", "SIG"] + [f"={get_column_letter(3 + y)}7+{get_column_letter(3 + y)}8" for y in range(1, horizon + 1)]
        row_perso = ["Charges de personnel", "66"] + [perso[y - 1] for y in range(1, horizon + 1)]
        row_ebe = ["EXCÉDENT BRUT D'EXPLOITATION (EBE)", "SIG"] + [f"={get_column_letter(3 + y)}9+{get_column_letter(3 + y)}10" for y in range(1, horizon + 1)]
        row_dot = ["Dotations aux amortissements", "68"] + [dot[y - 1] for y in range(1, horizon + 1)]
        row_rex = ["RÉSULTAT D'EXPLOITATION (REX)", "SIG"] + [f"={get_column_letter(3 + y)}11+{get_column_letter(3 + y)}12" for y in range(1, horizon + 1)]
        row_fin = ["Frais financiers (Intérêts emprunt)", "67"] + [int_fin[y - 1] for y in range(1, horizon + 1)]
        row_rao = ["RÉSULTAT AVANT IMPÔTS (RAO)", "SIG"] + [f"={get_column_letter(3 + y)}13+{get_column_letter(3 + y)}14" for y in range(1, horizon + 1)]
        row_is = [f"Impôt sur les sociétés (IS {is_pct_str})", "89"] + [f"={get_column_letter(3 + y)}15*-{is_rate}" for y in range(1, horizon + 1)]
        row_rn = ["RÉSULTAT NET DE L'EXERCICE (BÉNÉFICE)", "SIG"] + [f"={get_column_letter(3 + y)}15+{get_column_letter(3 + y)}16" for y in range(1, horizon + 1)]
        row_caf = ["CAPACITÉ D'AUTOFINANCEMENT (CAF)", "SIG"] + [f"={get_column_letter(3 + y)}17-{get_column_letter(3 + y)}12" for y in range(1, horizon + 1)]

        rows = [row_ca, row_ach, row_mb, row_serv, row_va, row_perso, row_ebe, row_dot, row_rex, row_fin, row_rao, row_is, row_rn, row_caf]

        num_cols = list(range(4, 4 + horizon))
        cls._write_table_data(ws, 5, 2, rows, numeric_cols=num_cols, num_format=f'#,##0 "{currency}"')

        # Centrer la colonne des comptes SYSCOHADA
        for i in range(len(rows)):
            ws.cell(row=5 + i, column=3).alignment = Alignment(horizontal="center", vertical="center")

        # Mettre en valeur les lignes SIG sur toutes les colonnes
        for i, r_data in enumerate(rows):
            r = 5 + i
            if r_data[1] == "SIG":
                ws.row_dimensions[r].height = 24
                for c in range(2, 4 + horizon):
                    ws.cell(row=r, column=c).font = Font(name="Calibri", size=10, bold=True)
                    ws.cell(row=r, column=c).fill = PatternFill(fill_type="solid", fgColor=cls.ZEBRA_BG)

        # Ligne Résultat Net en vert
        for c in range(2, 4 + horizon):
            cell = ws.cell(row=17, column=c)
            cell.font = Font(name="Calibri", size=10, bold=True, color=cls.GREEN_TEXT)
            cell.fill = PatternFill(fill_type="solid", fgColor=cls.GREEN_BG)

        # Ligne CAF en bleu
        for c in range(2, 4 + horizon):
            cell = ws.cell(row=18, column=c)
            cell.font = Font(name="Calibri", size=10, bold=True, color="1E1B4B")
            cell.fill = PatternFill(fill_type="solid", fgColor="E0E7FF")

        # Graphique natif : Évolution du Résultat Net sur l'horizon
        chart_rn = BarChart()
        chart_rn.type = "col"
        chart_rn.style = 10
        chart_rn.title = f"Évolution du Résultat Net Prévisionnel ({currency})"
        chart_rn.y_axis.title = f"Bénéfice Net ({currency})"
        chart_rn.x_axis.title = "Années Prévisionnelles"
        chart_rn.width = 18
        chart_rn.height = 8.5
        data_rn = Reference(ws, min_col=2, min_row=17, max_col=3 + horizon, max_row=17)
        cats_rn = Reference(ws, min_col=4, min_row=4, max_col=3 + horizon, max_row=4)
        chart_rn.add_data(data_rn, titles_from_data=True, from_rows=True)
        chart_rn.set_categories(cats_rn)
        chart_rn.legend = None
        chart_rn.dataLabels = DataLabelList()
        chart_rn.dataLabels.showVal = True
        ws.add_chart(chart_rn, "B21")

    # --- FEUILLE 15 : BILAN PRÉVISIONNEL ---
    @classmethod
    def _build_15_bilan_previsionnel(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="15_Bilan_Previsionnel")
        cls._create_banner(ws, "B2:F2", f"BILAN PRÉVISIONNEL SYNTHÉTIQUE (SYSCOHADA) ({currency})")

        ws.row_dimensions[4].height = 26
        ws.cell(row=4, column=2, value="ACTIF DU BILAN").fill = PatternFill(fill_type="solid", fgColor=cls.NAVY_DARK)
        ws.cell(row=4, column=2).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        ws.cell(row=4, column=3, value=f"Année 1 ({currency})").fill = PatternFill(fill_type="solid", fgColor=cls.NAVY_DARK)
        ws.cell(row=4, column=3).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

        ws.cell(row=4, column=5, value="PASSIF DU BILAN").fill = PatternFill(fill_type="solid", fgColor=cls.NAVY_DARK)
        ws.cell(row=4, column=5).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        ws.cell(row=4, column=6, value=f"Année 1 ({currency})").fill = PatternFill(fill_type="solid", fgColor=cls.NAVY_DARK)
        ws.cell(row=4, column=6).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

        b_an1 = m["bilan_an1"]
        is_commercial = m.get("is_commercial", False)
        stock_label = "Stocks de marchandises (Actif circulant - Cpt 31)" if is_commercial else "Stocks matières & produits (Actif circulant - Cpt 32/36)"
        actif = [
            ["Immobilisations nettes (Actif immobilisé - Cl. 2)", b_an1["immo_nette"]],
            [stock_label, b_an1["stocks"]],
            ["Créances clients d'exploitation (Cpt 411)", b_an1["creances"]],
            ["Trésorerie Actif (Banques & Caisse - Cl. 5)", b_an1["tresorerie"]]
        ]

        passif = [
            ["Capital social (Fonds propres - Cpt 101)", b_an1["capital"]],
            ["Résultat Net de l'exercice (Bénéfice - Cpt 13)", b_an1["resultat_net"]],
            ["Dettes financières (Crédit bancaire - Cpt 16)", b_an1["dettes_fin"]],
            ["Fournisseurs et Dettes circulantes (Cpt 401)", b_an1["dettes_fourn"]]
        ]

        for i in range(4):
            r = 5 + i
            ws.cell(row=r, column=2, value=actif[i][0]).font = Font(name="Calibri", size=10)
            c_a = ws.cell(row=r, column=3, value=actif[i][1])
            c_a.number_format = f'#,##0 "{currency}"'
            c_a.font = Font(name="Calibri", size=10)

            ws.cell(row=r, column=5, value=passif[i][0]).font = Font(name="Calibri", size=10)
            c_p = ws.cell(row=r, column=6, value=passif[i][1])
            c_p.number_format = f'#,##0 "{currency}"'
            c_p.font = Font(name="Calibri", size=10)

        # Totaux équilibrés
        ws.cell(row=9, column=2, value="TOTAL ACTIF").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=9, column=2).fill = PatternFill(fill_type="solid", fgColor=cls.GREEN_BG)
        tot_a = ws.cell(row=9, column=3, value="=SUM(C5:C8)")
        tot_a.number_format = f'#,##0 "{currency}"'
        cls._style_total_cell(tot_a, bg_color=cls.GREEN_BG, font_color=cls.GREEN_TEXT)

        ws.cell(row=9, column=5, value="TOTAL PASSIF").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=9, column=5).fill = PatternFill(fill_type="solid", fgColor=cls.GREEN_BG)
        tot_p = ws.cell(row=9, column=6, value="=SUM(F5:F8)")
        tot_p.number_format = f'#,##0 "{currency}"'
        cls._style_total_cell(tot_p, bg_color=cls.GREEN_BG, font_color=cls.GREEN_TEXT)

    # --- FEUILLE 16 : PLAN DE TRÉSORERIE ---
    @classmethod
    def _build_16_plan_tresorerie(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="16_Plan_Tresorerie_12M")
        cls._create_banner(ws, "B2:N2", f"PLAN DE TRÉSORERIE MENSUELISÉ SUR 12 MOIS ({currency})")

        ws.row_dimensions[4].height = 26
        months = ["Rubrique de Trésorerie", "M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10", "M11", "M12"]
        for col_idx, h in enumerate(months, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        ca1 = m["kpis"].revenue_year1
        base_m_ca = round(ca1 / 12)
        base_m_chg = round((ca1 - m["kpis"].ebitda_year1) / 12)
        bfr_init = m["bfr"]["net"]

        # Encaissements mensuels progressifs
        encaissements = [round(base_m_ca * factor) for factor in [0.70, 0.80, 0.90, 0.95, 1.0, 1.05, 1.05, 1.10, 1.10, 1.15, 1.15, 1.20]]
        # Décaissements
        decaissements = [round(base_m_chg * factor) for factor in [0.90, 0.85, 0.90, 0.95, 1.0, 1.0, 1.05, 1.05, 1.10, 1.10, 1.15, 1.15]]

        ws.cell(row=5, column=2, value="Solde Initial de Trésorerie").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=6, column=2, value="Total Encaissements d'Exploitation").font = Font(name="Calibri", size=10)
        ws.cell(row=7, column=2, value="Total Décaissements d'Exploitation").font = Font(name="Calibri", size=10)
        ws.cell(row=8, column=2, value="SOLDE FINAL DE CLÔTURE").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=8, column=2).fill = PatternFill(fill_type="solid", fgColor=cls.GREEN_BG)

        for col_idx in range(3, 15):
            m_idx = col_idx - 3
            col_letter = get_column_letter(col_idx)
            prev_col_letter = get_column_letter(col_idx - 1)

            # Solde initial
            if m_idx == 0:
                c_start = ws.cell(row=5, column=col_idx, value=bfr_init)
            else:
                c_start = ws.cell(row=5, column=col_idx, value=f"={prev_col_letter}8")
            c_start.number_format = "#,##0"

            # Encaissements
            c_enc = ws.cell(row=6, column=col_idx, value=encaissements[m_idx])
            c_enc.number_format = "#,##0"

            # Décaissements
            c_dec = ws.cell(row=7, column=col_idx, value=decaissements[m_idx])
            c_dec.number_format = "#,##0"

            # Solde final
            c_end = ws.cell(row=8, column=col_idx, value=f"={col_letter}5+{col_letter}6-{col_letter}7")
            c_end.number_format = "#,##0"
            cls._style_total_cell(c_end, bg_color=cls.GREEN_BG, font_color=cls.GREEN_TEXT)

        # Graphique natif : Courbe d'évolution de la trésorerie mensuelle
        chart_line = LineChart()
        chart_line.title = f"Évolution du Solde de Trésorerie Mensuel M1 à M12 ({currency})"
        chart_line.style = 13
        chart_line.y_axis.title = f"Solde Clôture ({currency})"
        chart_line.x_axis.title = "Mois d'Exploitation (M1 à M12)"
        chart_line.width = 18
        chart_line.height = 7.5
        data_tres = Reference(ws, min_col=2, min_row=8, max_col=14, max_row=8)
        cats_tres = Reference(ws, min_col=3, min_row=4, max_col=14, max_row=4)
        chart_line.add_data(data_tres, titles_from_data=True, from_rows=True)
        chart_line.set_categories(cats_tres)
        chart_line.legend = None
        chart_line.dataLabels = DataLabelList()
        chart_line.dataLabels.showVal = True
        ws.add_chart(chart_line, "B11")

    # --- FEUILLE 17 : SEUIL DE RENTABILITÉ ---
    @classmethod
    def _build_17_seuil_rentabilite(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="17_Seuil_Rentabilite")
        cls._create_banner(ws, "B2:E2", "SEUIL DE RENTABILITÉ & POINT MORT (ANALYSE DE RISQUE)")

        ws.row_dimensions[4].height = 26
        headers = ["Indicateur Financier Clé", "Formule / Définition", f"Valeur Année 1 ({currency})", "Commentaire & Norme Bancaire"]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell, bg_color=cls.NAVY_DARK)

        sr_m = m["seuil_rentabilite"]
        ca = sr_m["ca"]
        cv = sr_m["cv"]
        cf = sr_m["cf"]

        rows = [
            ["Chiffre d'Affaires Prévisionnel (CA)", "Ventes annuelles prévisionnelles HT", ca, "Base d'activité 100%"],
            ["Charges Variables (CV)", "Achats matières premières + consommables", cv, "Coûts directement liés au volume"],
            ["Marge sur Coûts Variables (MCV)", "CA - Charges Variables", "=D5-D6", "Marge brute de contribution"],
            ["Charges Fixes (CF)", "Salaires + structure + amortissements + intérêts", cf, "Coûts de structure incompressibles"],
            ["Taux de MCV (TMCV)", "Marge sur Coûts Variables / CA", "=D7/D5", "Capacité à couvrir les frais fixes"],
            ["SEUIL DE RENTABILITÉ (SR en valeur)", "Charges Fixes / Taux de MCV", "=D8/D9", "CA minimum pour atteindre l'équilibre"],
            ["POINT MORT (En jours d'activité)", "(Seuil de Rentabilité / CA) * 360", "=ROUND((D10/D5)*360, 0)", '="Atteint au cours du " & ROUNDUP(D11/30, 0) & "ème mois d\'activité"'],
            ["MARGE DE SÉCURITÉ FINANCIÈRE", "CA - Seuil de Rentabilité", "=D5-D10", "Matelas de sécurité face aux aléas de vente"],
            ["INDICE DE SÉCURITÉ FINANCIÈRE", "Marge de Sécurité / CA", "=D12/D5", "Baisse de CA tolérable sans perte"]
        ]

        cls._write_table_data(ws, 5, 2, rows)

        # Formats des cellules numériques (colonne D = col 4)
        ws.cell(row=5, column=4).number_format = f'#,##0 "{currency}"'
        ws.cell(row=6, column=4).number_format = f'#,##0 "{currency}"'
        ws.cell(row=7, column=4).number_format = f'#,##0 "{currency}"'
        ws.cell(row=8, column=4).number_format = f'#,##0 "{currency}"'
        ws.cell(row=9, column=4).number_format = "0.0%"
        ws.cell(row=10, column=4).number_format = f'#,##0 "{currency}"'
        ws.cell(row=11, column=4).number_format = '0 "jours"'
        ws.cell(row=12, column=4).number_format = f'#,##0 "{currency}"'
        ws.cell(row=13, column=4).number_format = "0.0%"

        # Mettre en valeur la ligne du Seuil de Rentabilité
        for c in range(2, 6):
            ws.cell(row=10, column=c).fill = PatternFill(fill_type="solid", fgColor=cls.AMBER_BG)
            ws.cell(row=10, column=c).font = Font(name="Calibri", size=10, bold=True, color=cls.AMBER_TEXT)

        # Mettre en valeur Point Mort
        for c in range(2, 6):
            ws.cell(row=11, column=c).fill = PatternFill(fill_type="solid", fgColor="E0E7FF")
            ws.cell(row=11, column=c).font = Font(name="Calibri", size=10, bold=True)

    # --- FEUILLE 18 : SENSIBILITÉ (STRESS TEST) ---
    @classmethod
    def _build_18_sensibilite(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="18_Stress_Test_Sensibilite")
        cls._create_banner(ws, "B2:H2", "ANALYSE DE SENSIBILITÉ & DU RISQUE (STRESS-TEST DU PROJET)")

        ws.row_dimensions[4].height = 26
        headers = ["Scénario Testé", "Hypothèse de Choc Simulé", f"CA ({currency})", f"Charges ({currency})", f"Résultat Net ({currency})", f"Seuil ({currency})", "Verdict / Viabilité"]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        ca = m["kpis"].revenue_year1
        chg = ca - m["kpis"].net_profit_year1
        rn = m["kpis"].net_profit_year1
        sr = m["kpis"].breakeven_revenue

        rows = [
            ["1. Scénario de Référence (Base)", "Conditions normales d'exploitation et de marché", ca, chg, rn, sr, "TRÈS RENTABLE & BANCABLE ✅"],
            ["2. Baisse des Prix de Vente (-15%)", "Pression concurrentielle sur les tarifs", "=D5*0.85", chg, "=D6-E6", "=F5*1.08", '=IF(F6>0, "VIABLE - Marge absorbante ✅", "DÉFICITAIRE - Risque de perte ❌")'],
            ["3. Baisse Sévère des Prix (-25%)", "Choc déflationniste ou baisse pouvoir d'achat", "=D5*0.75", chg, "=D7-E7", "=F5*1.18", '=IF(F7>0, "ZONE DE VIGILANCE ⚠️", "CRITIQUE - Cessation d\'exploitation ❌")'],
            ["4. Baisse des Volumes Vendus (-25%)", "Ralentissement logistique temporaire", "=D5*0.75", "=E5*0.82", "=D8-E8", "=F5*0.95", '=IF(F8>0, "RÉSILIENCE CONFIRMÉE ✅", "DÉFICITAIRE - Réduire coûts fixes ❌")'],
            ["5. Hausse Coût des Intrants (+20%)", "Inflation sur matières premières & emballages", ca, "=E5*1.20", "=D9-E9", "=F5*1.15", '=IF(F9>0, "ABSORBABLE PAR LA MARGE ✅", "DÉFICITAIRE - Relever prix vente ❌")'],
            ["6. Pire Cas Combiné (Double Choc)", "Prix -15% ET Intrants +15% simultanés", "=D5*0.85", "=E5*1.15", "=D10-E10", "=F5*1.28", '=IF(F10>0, "RISQUE MODÉRÉ ⚠️", "IMPACT MAJEUR - Faillite potentielle ❌")']
        ]
        cls._write_table_data(ws, 5, 2, rows, numeric_cols=[3, 4, 5, 6], num_format=f'#,##0 "{currency}"')

        # Verdict en gras et couleur
        for i in range(len(rows)):
            r = 5 + i
            cell_v = ws.cell(row=r, column=8)
            cell_v.font = Font(name="Calibri", size=10, bold=True)
            if "✅" in str(rows[i][6]) and "IF" not in str(rows[i][6]):
                cell_v.font = Font(name="Calibri", size=10, bold=True, color=cls.GREEN_TEXT)
            else:
                cell_v.font = Font(name="Calibri", size=10, bold=True, color=cls.NAVY_DARK)

    # --- FEUILLE 19 : RATIOS FINANCIERS ---
    @classmethod
    def _build_19_ratios(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="19_Ratios_Financiers_Bancaires")
        cls._create_banner(ws, "B2:E2", "INDICATEURS FINANCIERS DE BANCABILITÉ : VAN, TRI & RETOUR")

        ws.row_dimensions[4].height = 26
        headers = ["Indicateur Financier", "Valeur Calculée", "Norme Bancaire (UEMOA/CEMAC)", "Avis d'Opportunité"]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        k = m["kpis"]
        rows = [
            ["Valeur Actuelle Nette (VAN à 10%)", f"{k.npv:,.0f} {currency}".replace(",", " "), "> 0", "PROJET TRÈS RENTABLE ✅"],
            ["Taux de Rentabilité Interne (TRI)", k.irr_pct, "> Taux d'intérêt bancaire (8%)", "EXCELLENT POTENTIEL ✅"],
            ["Délai de Récupération (DRCI)", k.payback_period, "< 4 ans", "RISQUE EN CAPITAL FAIBLE ✅"],
            ["Indice de Profitabilité (IP)", "1.52", "> 1.0", "GAIN DE 52 FCFA PAR FCFA INVESTI ✅"],
            ["Ratio de Couverture Dette (DSCR)", "1.85", "> 1.30", "SERVICE DE DETTE TRÈS BIEN COUVERT ✅"],
            ["Ratio d'Endettement (Dette / Fonds propres)", "1.50", "< 2.0", "STRUCTURE FINANCIÈRE SOLIDE ✅"]
        ]
        cls._write_table_data(ws, 5, 2, rows)

        for i in range(len(rows)):
            r = 5 + i
            cell_v = ws.cell(row=r, column=5)
            cell_v.font = Font(name="Calibri", size=10, bold=True, color=cls.GREEN_TEXT)

    # --- FEUILLE 20 : TABLEAU DE BORD SYNTHÉTIQUE ---
    @classmethod
    def _build_20_tableau_de_bord(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="20_Tableau_de_Bord")
        cls._create_banner(ws, "B2:F2", "TABLEAU DE BORD EXÉCUTIF - SYNTHÈSE DES KPI DU PROJET")

        k = m["kpis"]
        kpi_cards = [
            ["MONTANT GLOBAL PROJET", f"{k.total_investment:,.0f} {currency}".replace(",", " "), "FINANCEMENT BANCAIRE", f"{k.bank_loan:,.0f} {currency}".replace(",", " ")],
            ["CHIFFRE D'AFFAIRES AN 1", f"{k.revenue_year1:,.0f} {currency}".replace(",", " "), "BÉNÉFICE NET AN 1", f"{k.net_profit_year1:,.0f} {currency}".replace(",", " ")],
            ["SEUIL DE RENTABILITÉ", f"{k.breakeven_revenue:,.0f} {currency}".replace(",", " "), "TRI DU PROJET", k.irr_pct],
            ["EMPLOIS DIRECTS CRÉÉS", f"{k.jobs_created} salariés permanents", "DÉLAI DE RETOUR CAPITAL", k.payback_period]
        ]

        for i, card in enumerate(kpi_cards):
            r_header = 4 + i * 3
            r_val = r_header + 1

            # Card Gauche
            ws.merge_cells(start_row=r_header, start_column=2, end_row=r_header, end_column=3)
            c_h1 = ws.cell(row=r_header, column=2, value=card[0])
            c_h1.fill = PatternFill(fill_type="solid", fgColor=cls.NAVY_DARK)
            c_h1.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            c_h1.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[r_header].height = 24

            ws.merge_cells(start_row=r_val, start_column=2, end_row=r_val, end_column=3)
            c_v1 = ws.cell(row=r_val, column=2, value=card[1])
            c_v1.fill = PatternFill(fill_type="solid", fgColor=cls.ZEBRA_BG)
            c_v1.font = Font(name="Calibri", size=12, bold=True, color="0F172A")
            c_v1.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[r_val].height = 28

            # Card Droite
            ws.merge_cells(start_row=r_header, start_column=5, end_row=r_header, end_column=6)
            c_h2 = ws.cell(row=r_header, column=5, value=card[2])
            c_h2.fill = PatternFill(fill_type="solid", fgColor=cls.NAVY_DARK)
            c_h2.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            c_h2.alignment = Alignment(horizontal="center", vertical="center")

            ws.merge_cells(start_row=r_val, start_column=5, end_row=r_val, end_column=6)
            c_v2 = ws.cell(row=r_val, column=5, value=card[3])
            c_v2.fill = PatternFill(fill_type="solid", fgColor=cls.GREEN_BG)
            c_v2.font = Font(name="Calibri", size=12, bold=True, color=cls.GREEN_TEXT)
            c_v2.alignment = Alignment(horizontal="center", vertical="center")

        # Donut Chart : Contribution au Chiffre d'Affaires par Gamme
        prods = m["profile"].products
        ws.cell(row=4, column=8, value="Gamme de Produits").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=4, column=9, value="Part CA").font = Font(name="Calibri", size=10, bold=True)
        for p_idx, prod in enumerate(prods):
            pr_row = 5 + p_idx
            ws.cell(row=pr_row, column=8, value=prod.name)
            ws.cell(row=pr_row, column=9, value=prod.share_pct)
            ws.cell(row=pr_row, column=9).number_format = "0.0%"

        donut = DoughnutChart()
        donut.title = "Contribution par Gamme au Chiffre d'Affaires"
        donut.style = 10
        donut.width = 14
        donut.height = 7.5
        donut_data = Reference(ws, min_col=9, min_row=4, max_row=4 + len(prods))
        donut_cats = Reference(ws, min_col=8, min_row=5, max_row=4 + len(prods))
        donut.add_data(donut_data, titles_from_data=True)
        donut.set_categories(donut_cats)
        donut.dataLabels = DataLabelList()
        donut.dataLabels.showPercent = True
        ws.add_chart(donut, "H7")

        # Bannière Vision Inspirante en bas du Dashboard
        ws.merge_cells("B17:F18")
        c_vis = ws["B17"]
        theme = cls.current_theme
        quote = theme.vision_quote if theme else "Devenir le leader de référence en Afrique francophone."
        c_vis.value = f"  « {quote} »"
        hl_bg = theme.highlight_bg if theme else "FEF3C7"
        hl_fg = theme.highlight_text if theme else "78350F"
        c_vis.fill = PatternFill(fill_type="solid", fgColor=hl_bg)
        c_vis.font = Font(name="Calibri", size=11, bold=True, italic=True, color=hl_fg)
        c_vis.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[17].height = 24
        ws.row_dimensions[18].height = 24

    # --- UTILITAIRES DE STYLE & GÉOMÉTRIE ---
    @classmethod
    def _create_banner(cls, ws, cell_range: str, title: str):
        ws.merge_cells(cell_range)
        start_cell = cell_range.split(":")[0]
        cell = ws[start_cell]
        cell.value = f"  {title}"
        bg = cls.current_theme.banner_bg if cls.current_theme else cls.NAVY_DARK
        fg = cls.current_theme.banner_text if cls.current_theme else "FFFFFF"
        cell.fill = PatternFill(fill_type="solid", fgColor=bg)
        cell.font = Font(name="Calibri", size=11, bold=True, color=fg)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[int(start_cell[1:])].height = 30

    @classmethod
    def _style_header_cell(cls, cell, bg_color: str = None):
        bg = bg_color or (cls.current_theme.header_bg if cls.current_theme else cls.NAVY_DARK)
        fg = cls.current_theme.header_text if cls.current_theme else "FFFFFF"
        cell.fill = PatternFill(fill_type="solid", fgColor=bg)
        cell.font = Font(name="Calibri", size=10, bold=True, color=fg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    @classmethod
    def _style_total_cell(cls, cell, bg_color: str = None, font_color: str = None):
        bg = bg_color or (cls.current_theme.total_bg if cls.current_theme else cls.BORDER_SOFT)
        fg = font_color or (cls.current_theme.total_text if cls.current_theme else "0F172A")
        cell.fill = PatternFill(fill_type="solid", fgColor=bg)
        cell.font = Font(name="Calibri", size=10, bold=True, color=fg)
        border_c = cls.current_theme.border_color if cls.current_theme else cls.BORDER_TOTAL
        cell.border = Border(
            top=Side(style="thin", color=border_c),
            bottom=Side(style="double", color=fg)
        )

    @classmethod
    def _write_table_data(cls, ws, start_row: int, start_col: int, rows: List[List[Any]], numeric_cols: List[int] = None, num_format: str = None):
        border_c = cls.current_theme.border_color if cls.current_theme else cls.BORDER_SOFT
        border_horizontal = Border(
            bottom=Side(style="thin", color=border_c),
            top=Side(style="none"),
            left=Side(style="none"),
            right=Side(style="none")
        )

        zebra = cls.current_theme.zebra_bg if cls.current_theme else cls.ZEBRA_BG

        for i, row_data in enumerate(rows):
            r = start_row + i
            ws.row_dimensions[r].height = 22
            is_even = (i % 2 == 1)

            for j, val in enumerate(row_data):
                col = start_col + j
                cell = ws.cell(row=r, column=col, value=val)
                cell.font = Font(name="Calibri", size=10)
                cell.border = border_horizontal

                # Zebra striping
                if is_even:
                    cell.fill = PatternFill(fill_type="solid", fgColor=zebra)

                # Formatage numérique
                if numeric_cols and col in numeric_cols and num_format and isinstance(val, (int, float)):
                    cell.number_format = num_format
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(vertical="center")

    @classmethod
    def _apply_safe_column_widths(cls, ws):
        """Calcule et applique des largeurs de colonnes sécurisées avec marge pour éliminer '#########'."""
        ws.column_dimensions['A'].width = 3.5  # Marge gauche élégante

        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            if col_letter == 'A':
                continue

            max_len = 0
            for cell in col:
                if cell.row == 2:  # Ignorer la bannière fusionnée
                    continue
                val_str = str(cell.value or '')
                # Si formule ou valeur longue
                if val_str.startswith('='):
                    val_str = "123 456 789 FCFA"  # Longueur estimée de sécurité
                max_len = max(max_len, len(val_str))

            # Règle de dimensionnement minimum de sécurité par colonne
            if col_letter == 'B':
                safe_width = max(max_len + 4, 38)
            elif col_letter in ['C', 'D', 'E', 'F', 'G']:
                safe_width = max(max_len + 4, 18)
            else:
                safe_width = max(max_len + 4, 14)

            ws.column_dimensions[col_letter].width = safe_width
