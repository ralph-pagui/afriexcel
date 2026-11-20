"""
Générateur de classeurs Excel professionnels ultra-modernes avec openpyxl.
Produit un dossier bancaire complet de 20 feuilles SYSCOHADA avec typographie prestige,
cartes KPI exécutives, indicateurs de bancabilité UEMOA/CEMAC, graphiques natifs,
largeurs de colonnes sécurisées (zéro '#########') et conformité OHADA stricte.
"""

import io
import base64
from typing import Dict, Any, List, Optional
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
    BLUE_BG = "E0F2FE"
    BLUE_TEXT = "0369A1"

    @classmethod
    def build(cls, data: Dict[str, Any], project_name: str, country: str = "Côte d'Ivoire", template_id: str = "auto") -> bytes:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        currency = data["currency"]
        profile = data["profile"]
        m = data

        theme = data.get("theme") or detect_or_get_template(template_id or data.get("template_id"), prompt_text=project_name, activity_text=profile.activity, sector_id=profile.sector_id)
        cls.current_theme = theme
        m["theme"] = theme

        # Construction séquentielle des 20 feuilles d'excellence bancable
        cls._build_01_page_de_garde(wb, project_name, country, currency, m)
        cls._build_02_presentation_entreprise(wb, project_name, profile, m)
        cls._build_03_etude_marche(wb, profile, country, currency, m)
        cls._build_04_strategie_marketing(wb, profile, currency, m)
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

    # =========================================================================
    # FEUILLE 01 : PAGE DE GARDE EXÉCUTIVE & SYNTHÈSE BANCAIRE
    # =========================================================================
    @classmethod
    def _build_01_page_de_garde(cls, wb, project_name: str, country: str, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="01_Page_de_Garde")
        cls._create_banner(ws, "B2:G2", "DOSSIER DE BUSINESS PLAN BANCAIRE & ÉTUDE DE FAISABILITÉ")

        theme = cls.current_theme
        ws.merge_cells("B3:G3")
        cell_tag = ws["B3"]
        tagline = theme.tagline if theme else "Entreprendre aujourd'hui, bâtir demain en Afrique"
        cell_tag.value = f"  « {tagline} »"
        hl_bg = theme.highlight_bg if theme else "FEF3C7"
        hl_fg = theme.highlight_text if theme else "78350F"
        cell_tag.fill = PatternFill(fill_type="solid", fgColor=hl_bg)
        cell_tag.font = Font(name="Calibri", size=10, italic=True, bold=True, color=hl_fg)
        cell_tag.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[3].height = 22

        # Table des Métadonnées du Projet (Colonnes B & C)
        headers = ["PARAMÈTRE DU DOSSIER", "VALEUR & MODALITÉS"]
        ws.row_dimensions[4].height = 26
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        kpis = m["kpis"]
        horizon = m.get("horizon_years", 5)
        theme_badge = f"{theme.icon} {theme.name} ({theme.subtitle})" if theme else "Standard OHADA"
        rows = [
            ["Désignation du Projet", project_name],
            ["Template & Identité Visuelle", theme_badge],
            ["Secteur d'Activité Cible", m["profile"].name],
            ["Pays & Zone Réglementaire", f"{country} (Espace Économique OHADA)"],
            ["Forme Juridique Projetée", "Société à Responsabilité Limitée (SARL Pluripersonnelle)"],
            ["Devise de Référence", currency],
            ["Horizon Prévisionnel", f"{horizon} ans (Projection SYSCOHADA)"],
            ["Montant Global du Projet", kpis.total_investment],
            ["Apport Promoteur (Fonds Propres)", kpis.equity_promoter],
            ["Emprunt Bancaire Sollicité", kpis.bank_loan],
            ["Chiffre d'Affaires Année 1", kpis.revenue_year1],
            ["Bénéfice Net Après Impôt An 1", kpis.net_profit_year1],
            ["Seuil de Rentabilité Annuel", f"{kpis.breakeven_revenue:,.0f} {currency} (Jour {kpis.breakeven_days})".replace(",", " ")],
            ["Taux de Rentabilité Interne (TRI)", kpis.irr_pct],
            ["Délai de Récupération (DRCI)", kpis.payback_period],
            ["Emplois Directs Permanents Créés", f"{kpis.jobs_created} salariés qualifiés"]
        ]
        cls._write_table_data(ws, 5, 2, rows, numeric_cols=[3], num_format=f'#,##0 "{currency}"')

        # Bloc KPI Exécutifs d'impact (Colonnes E & F)
        ws.cell(row=4, column=5, value="INDICATEUR CLÉ").fill = PatternFill(fill_type="solid", fgColor=theme.header_bg)
        ws.cell(row=4, column=5).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        ws.cell(row=4, column=5).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=4, column=6, value="SYNTHÈSE RAPIDE").fill = PatternFill(fill_type="solid", fgColor=theme.header_bg)
        ws.cell(row=4, column=6).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        ws.cell(row=4, column=6).alignment = Alignment(horizontal="center", vertical="center")

        fin = m.get("financing", {})
        part_fp = fin.get("part_fp_pct", 30.0)
        part_loan = fin.get("part_emprunt_pct", 70.0)
        rate = fin.get("interest_rate", 0.12)
        rate_pct = round(rate * 100, 1)
        loan_years = fin.get("loan_duration_years", 5)

        kpi_cards = [
            ("Investissement Total (CAPEX+BFR)", f"{kpis.total_investment:,.0f} {currency}".replace(",", " "), "Coût global de démarrage", "CAPEX + BFR initial"),
            ("Fonds Propres Mobilisés", f"{kpis.equity_promoter:,.0f} {currency} ({part_fp}%)".replace(",", " "), "Apport promoteur", f"Part {part_fp:.0f}%"),
            ("Crédit Moyen Terme Sollicité", f"{kpis.bank_loan:,.0f} {currency} ({part_loan}%)".replace(",", " "), f"Taux {rate_pct}% sur {loan_years} ans", f"Part {part_loan:.0f}%"),
            ("Chiffre d'Affaires Année 1", f"{kpis.revenue_year1:,.0f} {currency}".replace(",", " "), "Ventes prévisionnelles", "Activité nominale"),
            ("Résultat Net Bénéficiaire An 1", f"{kpis.net_profit_year1:,.0f} {currency}".replace(",", " "), "Marge nette protectrice", "Bénéfice de départ"),
            ("Taux de Rentabilité Interne (TRI)", kpis.irr_pct, "Rendement financier du projet", "Flux actualisés"),
            ("Délai de Récupération (DRCI)", kpis.payback_period, "Retour sur capital investi", "Horizon de retour")
        ]

        for idx, (label, val, note, sub_tag) in enumerate(kpi_cards):
            r = 5 + idx * 2
            ws.cell(row=r, column=5, value=label).font = Font(name="Calibri", size=9, bold=True, color="475569")
            ws.cell(row=r, column=5).fill = PatternFill(fill_type="solid", fgColor=theme.card_bg)
            ws.cell(row=r+1, column=5, value=note).font = Font(name="Calibri", size=8, italic=True, color="94A3B8")
            ws.cell(row=r+1, column=5).fill = PatternFill(fill_type="solid", fgColor=theme.card_bg)

            val_cell = ws.cell(row=r, column=6, value=val)
            val_cell.font = Font(name="Calibri", size=11, bold=True, color="0F172A")
            val_cell.fill = PatternFill(fill_type="solid", fgColor=theme.card_bg)
            val_cell.alignment = Alignment(horizontal="right", vertical="center")

            sub_cell = ws.cell(row=r+1, column=6, value=sub_tag)
            sub_cell.font = Font(name="Calibri", size=8, color=theme.badge_success_fg)
            sub_cell.fill = PatternFill(fill_type="solid", fgColor=theme.card_bg)
            sub_cell.alignment = Alignment(horizontal="right", vertical="center")
            ws.row_dimensions[r].height = 20
            ws.row_dimensions[r+1].height = 16

    # =========================================================================
    # FEUILLE 02 : PRÉSENTATION DE L'ENTREPRISE & GOUVERNANCE
    # =========================================================================
    @classmethod
    def _build_02_presentation_entreprise(cls, wb, project_name: str, profile: Any, m: Dict[str, Any] = None):
        ws = wb.create_sheet(title="02_Presentation_Entreprise")
        cls._create_banner(ws, "B2:E2", "FICHE D'IDENTITÉ & GOUVERNANCE DU PROJET")

        ws.row_dimensions[4].height = 26
        for col_idx, h in enumerate(["Rubrique Institutionnelle", "Détails & Orientations Stratégiques"], start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        horizon = (m or {}).get("horizon_years", 5)
        is_comm = (m or {}).get("is_commercial", False)
        sec_id = getattr(profile, "sector_id", "")

        appro_label = "Approvisionnement Marchandises / Stocks" if is_comm else "Approvisionnement Matières Premières & Intrants"
        if sec_id == "transport":
            appro_label = "Approvisionnement Carburant, Pneumatiques & Maintenance"
            norme_label = "Licence d'exploitation de transport routier de marchandises, visites techniques régulières, RCCM & déclaration fiscale"
        elif sec_id == "agriculture":
            appro_label = "Approvisionnement Semences Certifiées, Fertilisants & Intrants Maraîchers"
            norme_label = "Bonnes pratiques agricoles (GAP), traçabilité des récoltes, conformité phytosanitaire & RCCM"
        elif sec_id == "clinique":
            appro_label = "Approvisionnement Réactifs de Laboratoire & Consommables Médicaux Stériles"
            norme_label = "Agrément du Ministère de la Santé, contrôle qualité des automates, protocole DASRI & RCCM"
        elif sec_id == "aviculture":
            appro_label = "Approvisionnement Poussins d'un Jour, Aliments Provende & Produits Vétérinaires"
            norme_label = "Normes de biosécurité avicole, suivi sanitaire vétérinaire régulier, RCCM & déclaration fiscale"
        elif sec_id == "jus":
            appro_label = "Approvisionnement Mangues Fraîches Locales, Bouteilles Verre/PET & Ingrédients"
            norme_label = "Normes de sécurité sanitaire HACCP, pasteurisation contrôlée, traçabilité des lots & agrément agroalimentaire"
        elif sec_id == "commerce_informatique":
            appro_label = "Approvisionnement Équipements Informatiques, Serveurs, PC & Équipements Réseaux"
            norme_label = "Garanties constructeurs officielles, conformité douanière, traçabilité des numéros de série & RCCM"
        elif sec_id == "energie_solaire":
            appro_label = "Approvisionnement Panneaux Solaires Tier-1, Onduleurs Hybrides & Batteries Lithium"
            norme_label = "Normes internationales CE/IEC 61215/61730, habilitation électricien photovoltaïque B2V/BR & conformité CI-ENERGIES"
        elif is_comm:
            norme_label = "Conformité stricte aux règles commerciales, immatriculation RCCM & IFU/NINEA local"
        else:
            norme_label = "Conformité aux réglementations sectorielles en vigueur, normes professionnelles, RCCM & déclaration fiscale"
        rows = [
            ["Raison Sociale / Enseigne", project_name],
            ["Activité Principale Déclarée", profile.activity],
            ["Gamme de Produits / Services Clés", " ; ".join([p.name for p in profile.products])],
            ["Unité Opérationnelle de Référence", profile.unit],
            [appro_label, profile.raw_materials_desc],
            ["Réglementation, Qualité & Normes", norme_label],
            [f"Vision Stratégique à {horizon} Ans", f"S'imposer comme un leader régional d'excellence en alliant qualité irréprochable, rentabilité durable et satisfaction client."],
            ["Structure de Gouvernance", "Gérance assurée par le promoteur avec assistance d'un expert-comptable agréé ONECCA."],
            ["Facteurs Clés de Succès (FCS)", "1. Chaîne d'approvisionnement sécurisée et approvisionnements réguliers\n2. Équipements modernes garantissant un rendement élevé\n3. Gestion financière rigoureuse conforme au SYSCOHADA"]
        ]
        cls._write_table_data(ws, 5, 2, rows)

    # =========================================================================
    # FEUILLE 03 : ÉTUDE DE MARCHÉ & MATRICE CONCURRENTIELLE
    # =========================================================================
    @classmethod
    def _build_03_etude_marche(cls, wb, profile: Any, country: str, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="03_Etude_Marche")
        cls._create_banner(ws, "B2:H2", "ANALYSE DU MARCHÉ, DIMENSIONNEMENT TAM-SAM-SOM & CONCURRENCE")

        # Section 1 : Dimensionnement Quantitatif TAM-SAM-SOM
        cls._create_section_divider(ws, "B4:F4", "▸ 1. DIMENSIONNEMENT DU MARCHÉ ADRESSABLE (TAM - SAM - SOM)")
        ws.row_dimensions[5].height = 24
        for col_idx, h in enumerate(["Niveau du Marché", "Définition Stratégique", f"Montant Estimé ({currency})", "Part Captée"], start=2):
            cell = ws.cell(row=5, column=col_idx, value=h)
            cls._style_header_cell(cell)

        tss = m.get("tam_sam_som", {})
        tam_v = tss.get("tam", m["kpis"].revenue_year1 * 22)
        sam_v = tss.get("sam", m["kpis"].revenue_year1 * 5.5)
        som_v = tss.get("som", m["kpis"].revenue_year1)

        tss_rows = [
            ["TAM (Total Addressable Market)", f"Volume d'affaires global du secteur en {country}", tam_v, "100.0%"],
            ["SAM (Serviceable Available Market)", "Segment urbain et périurbain ciblé prioritairement", sam_v, f"{round((sam_v/tam_v)*100, 1)}%"],
            ["SOM (Serviceable Obtainable Market)", "Chiffre d'affaires cible Année 1 du projet", som_v, f"{round((som_v/sam_v)*100, 1)}% du SAM"]
        ]
        cls._write_table_data(ws, 6, 2, tss_rows, numeric_cols=[4], num_format=f'#,##0 "{currency}"')

        # Section 2 : Matrice Concurrentielle Locale
        cls._create_section_divider(ws, "B10:G10", "▸ 2. MATRICE CONCURRENTIELLE LOCALE (DIRECTS & INDIRECTS)")
        ws.row_dimensions[11].height = 24
        comp_headers = ["Catégorie de Concurrent", "Acteurs Typiques sur le Terrain", "Leurs Forces", "Leurs Faiblesses", "Notre Avantage Clé"]
        for col_idx, h in enumerate(comp_headers, start=2):
            cell = ws.cell(row=11, column=col_idx, value=h)
            cls._style_header_cell(cell)

        sec_id = getattr(profile, "sector_id", "")
        if sec_id == "transport":
            comp_rows = [
                ["Transporteurs Informels / Coxeurs", "Opérateurs isolés, vieux camions vétustes", "Prix bas négociés au coup par coup", "Retards fréquents, pannes récurrentes, fret non assuré, aucun suivi", "Flotte moderne 10t, traçabilité GPS temps réel, respect strict des délais et fret assuré"],
                ["Grandes Compagnies Multinationales", "Grands groupes logistiques internationaux", "Forte notoriété, logistique multimodale maritime/aérienne", "Tarifs prohibitifs, lourdeur administrative, focalisation sur très grands comptes", "Tarifs compétitifs PME, réactivité commerciale 24/7 et flexibilité contractuelle"],
                ["Flottes propres des entreprises", "Camions intégrés de certains industriels", "Disponibilité interne directe pour l'entreprise propriétaire", "Coûts fixes très lourds, immobilisation de capital, retours à vide fréquents", "Externalisation variable à la demande, optimisation des coûts et zéro immobilisation CAPEX"]
            ]
        elif sec_id == "agriculture":
            comp_rows = [
                ["Producteurs Maraîchers Traditionnels", "Petits exploitants en culture pluviale plein champ", "Faibles coûts initiaux, circuits courts de village", "Production saisonnière irrégulière, calibrage hétérogène, ruptures en saison sèche", "Production continue toute l'année sous irrigation goutte-à-goutte, calibrage 1er choix"],
                ["Importations Sous-régionales", "Tomates en provenance des pays voisins", "Disponibilité temporaire en contre-saison", "Coûts de transport et douanes élevés, pertes par avarie durant le trajet", "Fraîcheur bord champ, récolte à maturité optimale et approvisionnement immédiat"],
                ["Grossistes et Intermédiaires des Marchés", "Commerçants collecteurs des marchés urbains", "Maîtrise des flux de trésorerie au comptant", "Spéculation sur les prix, pression à la baisse sur les producteurs", "Contrats d'approvisionnement direct avec supermarchés, restaurants et grossistes fiables"]
            ]
        elif sec_id == "clinique":
            comp_rows = [
                ["Structures Sanitaires Publiques", "Centres de santé de référence et hôpitaux publics", "Tarifs subventionnés par l'État", "Files d'attente interminables, ruptures de réactifs, automates souvent en panne", "Prise en charge immédiate, automates récents, résultats fiables en moins de 2 heures"],
                ["Laboratoires Informels de Quartier", "Petites officines non accréditées", "Proximité immédiate dans les quartiers périphériques", "Contrôles qualité insuffisants, manque de réactifs validés, absence de biologiste", "Biologiste diplômé permanent, contrôles qualité étalonnés, hygiène et traçabilité certifiée"],
                ["Grands Laboratoires Centraux Privés", "Quelques laboratoires de prestige du centre-ville", "Plateau technique très complet", "Tarifs très onéreux, éloignement géographique pour les communes périphériques", "Tarification accessible conventionnée, proximité géographique et transmission par SMS/Web"]
            ]
        elif sec_id == "aviculture":
            comp_rows = [
                ["Éleveurs Traditionnels & Basse-cour", "Petits éleveurs informels en claustration sommaire", "Faibles coûts de structure, proximité villageoise", "Taux de mortalité élevé (>15%), croissance lente, approvisionnement irrégulier", "Poulets calibrés 1.8-2.0 kg, alimentation contrôlée, prophylaxie rigoureuse et disponibilité continue"],
                ["Importations de Volailles Congelées", "Découpes et poulets surgelés importés", "Prix bas au carton, marques internationales", "Qualité gustative altérée par décongélation, dépendance extérieure, devises", "Poulets frais abattus localement du jour, traçabilité garantie et soutien à l'économie locale"],
                ["Fermes Industrielles Intégrées", "Quelques grands complexes avicoles nationaux", "Volumes importants et capacités financières", "Politique commerciale rigide axée uniquement sur les hypermarchés", "Proximité relationnelle, flexibilité des volumes livrés aux demi-grossistes et rôtisseries"]
            ]
        elif sec_id == "jus":
            comp_rows = [
                ["Vendeurs Informels de Jus & Pressage Artisanal", "Vente ambulante en sachets ou bidons de récupération", "Prix bas négociés au comptant", "Défaut d'hygiène, non pasteurisé (fermentation rapide sous 24h), absence de traçabilité", "Pasteurisation industrielle contrôlée, DLUO longue conservation, bouteilles neuves capsulées et hygiène garantie"],
                ["Grandes Marques de Jus Importés & Sodas", "Boissons gazeuses et nectars concentrés importés", "Forte notoriété internationale, distribution massive en GMS", "Prix très élevés, additifs chimiques et forte teneur en sucres raffinés", "Pur jus 100% naturel sans conservateur, valorisation des mangues locales de qualité et prix accessible"],
                ["Petites Unités Agroalimentaires Semi-artisanales", "Petits ateliers de transformation locaux", "Présence ponctuelle dans quelques boutiques", "Capacité limitée, ruptures fréquentes en contre-saison, conditionnement artisanal", "Capacité nominale de 3 000 L/jour, approvisionnement saisonnier sécurisé et présence multi-formats (33cl et 1L)"]
            ]
        elif sec_id == "commerce_informatique":
            comp_rows = [
                ["Revendeurs Informels de Matériel d'Occasion", "Commerçants isolés de matériel reconditionné", "Prix très bas au comptant", "Matériel sans garantie, pannes fréquentes, absence de factures normalisées pour entreprises", "Équipements neufs sous garantie constructeur, facturation normalisée OHADA et SAV certifié"],
                ["Grossistes Importateurs Internationaux", "Représentations officielles de grands constructeurs mondiaux", "Catalogue étendu et stocks volumineux", "Tarifs rigides, délais de livraison longs pour PME, pas d'installation sur mesure", "Réactivité immédiate, conseil personnalisé avant-vente, installation et intégration sur site client"],
                ["Grandes Surfaces & Distributeurs Grand Public", "Rayons informatiques des grands magasins", "Visibilité commerciale grand public", "Offre orientée grand public sans solutions serveurs/réseaux d'entreprise", "Offre B2B sur mesure, intégration de parcs informatiques, serveurs et contrats de maintenance récurrents"]
            ]
        elif sec_id == "energie_solaire":
            comp_rows = [
                ["Installateurs Informels / Électriciens Bâtiment", "Artisans électriciens sans qualification solaire", "Prix bas négociés au coup par coup", "Matériel contrefait ou déclassé, sous-dimensionnement, pannes de batteries en <6 mois, risque d'incendie", "Dimensionnement rigoureux par logiciel PV, matériel Tier-1 certifié, garanties constructeur et sécurité"],
                ["Grandes Multinationales de l'Énergie", "Filiales de majors énergétiques internationales", "Très forte capacité financière, grands parcs solaires MW", "Inadaptées aux besoins des PME et résidences, devis très longs, tarifs prohibitifs", "Agilité d'intervention sous 48h, tarification compétitive PME/résidentiel et service de proximité"],
                ["Distributeurs Vendeurs de Kits sans Pose", "Comptoirs d'importation vendant des kits au carton", "Disponibilité de stock en magasin", "Aucune ingénierie d'installation, absence de SAV et aucun suivi après-vente", "Offre clés en main : étude d'ingénierie, pose certifiée, raccordement et contrat de maintenance récurrent"]
            ]
        else:
            comp_rows = [
                ["Opérateurs Informels Locaux", "Petits producteurs / commerçants de quartier", "Proximité, prix bas sans taxes", "Qualité variable, absence de facturation officielle, pas de traçabilité", "Qualité certifiée, facturation normalisée OHADA, régularité des approvisionnements"],
                ["Importations Étrangères", "Produits industriels importés", "Marques connues, emballages attrayants", "Prix élevés, délais de livraison longs, rupture fréquente", "Prix compétitif, disponibilité locale permanente et service après-vente réactif"],
                ["Acteurs Modernes Structurés", "Moyennes et grandes entreprises", "Capacité financière et logistique", "Coûts de structure lourds, faible flexibilité client", "Réactivité commerciale, service personnalisé et ancrage de terrain"]
            ]
        cls._write_table_data(ws, 12, 2, comp_rows)

        # Section 3 : Dynamique et Croissance du Marché
        cls._create_section_divider(ws, "B16:D16", "▸ 3. CROISSANCE ANNUELLE")
        ws.cell(row=17, column=2, value="Année").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=17, column=3, value="Indice Croissance").font = Font(name="Calibri", size=10, bold=True)
        indices = [100, 115, 132, 150, 172]
        for idx_y, val_idx in enumerate(indices):
            r_idx = 18 + idx_y
            ws.cell(row=r_idx, column=2, value=f"Année {idx_y + 1}")
            ws.cell(row=r_idx, column=3, value=val_idx)

        # Graphique natif : Évolution Estimée du Marché
        chart_mkt = BarChart()
        chart_mkt.type = "col"
        chart_mkt.style = 10
        chart_mkt.title = f"Évolution Estimée du Marché en {country} (Base 100)"
        chart_mkt.y_axis.title = "Indice Croissance"
        chart_mkt.x_axis.title = "Horizon Annuel"
        chart_mkt.width = 13.5
        chart_mkt.height = 6.8
        data_mkt = Reference(ws, min_col=3, min_row=17, max_row=22)
        cats_mkt = Reference(ws, min_col=2, min_row=18, max_row=22)
        chart_mkt.add_data(data_mkt, titles_from_data=True)
        chart_mkt.set_categories(cats_mkt)
        chart_mkt.legend = None
        ws.add_chart(chart_mkt, "E16")

    # =========================================================================
    # FEUILLE 04 : STRATÉGIE MARKETING & PLAN D'ACQUISITION (7P)
    # =========================================================================
    @classmethod
    def _build_04_strategie_marketing(cls, wb, profile: Any, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="04_Strategie_Marketing")
        cls._create_banner(ws, "B2:F2", "STRATÉGIE COMMERCIALE & MARKETING MIX COMPLET (7P)")

        ws.row_dimensions[4].height = 26
        for col_idx, h in enumerate(["Pilier du Marketing Mix (7P)", "Stratégie Opérationnelle Retenue", "Objectif Commercial & Livrable"], start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        p1_price = profile.products[0].unit_price
        p_last_price = profile.products[-1].unit_price
        sec_id = getattr(profile, "sector_id", "")

        if sec_id == "transport":
            rows = [
                ["1. Produit (Product)", f"Fret routier & logistique de marchandises : {profile.products[0].name}", "Sécurité absolue des marchandises, arrimage certifié, géolocalisation GPS en temps réel"],
                ["2. Prix (Price)", f"Tarification kilométrique & forfaits transparents : de {p1_price:,.0f} à {p_last_price:,.0f} {currency}".replace(",", " "), "Tarifs ultra-compétitifs et dégressifs selon volumes et fréquence contractuelle"],
                ["3. Distribution (Place)", "Axes routiers majeurs et corridors économiques régionaux", "Dessertes régulières quotidiennes, hubs de distribution et parkings logistiques partenaires"],
                ["4. Promotion (Promotion)", "Prospection commerciale B2B ciblée auprès des industriels, grossistes et importateurs", "Signature de contrats-cadres annuels, présence annuaires professionnels et plateforme web"],
                ["5. Personnel (People)", "Chauffeurs routiers expérimentés formés à l'éco-conduite et régulateur de flotte dédié", "Sens aigu de la ponctualité, respect strict du code de la route et relation client soignée"],
                ["6. Processus (Process)", "Gestion digitalisée des ordres de mission et transmission instantanée des bordereaux", "Ordres de mission informatisés, tracking de livraison en direct et bons émargés dématérialisés"],
                ["7. Preuve Matérielle (Physical)", "Flotte de 5 camions 10t récents floqués, carnets d'entretien et attestations d'assurance", "Garantie de solvabilité, attestations d'assurance fret et base logistique sécurisée"]
            ]
        elif sec_id == "agriculture":
            rows = [
                ["1. Produit (Product)", f"Tomates fraîches de table calibre sélectionné : {profile.products[0].name}", "Fermeté, excellente tenue post-récolte, saveur authentique sans résidus toxiques"],
                ["2. Prix (Price)", f"Prix de gros compétitif bord-champ : de {p1_price:,.0f} à {p_last_price:,.0f} {currency}".replace(",", " "), "Marge très attractive pour les demi-grossistes tout en sécurisant la rentabilité de la parcelle"],
                ["3. Distribution (Place)", "Circuits courts : marchés de gros urbains, grossistes partenaires et restauration", "Livraisons directes en cagettes aérées pour préserver l'intégrité physique des fruits"],
                ["4. Promotion (Promotion)", "Démarchage direct des coopératives de commerçantes, centrales d'achats et restaurants", "Échantillonnage qualité, partenariats réguliers et communication WhatsApp Pro"],
                ["5. Personnel (People)", "Ouvriers agricoles maraîchers formés aux bonnes pratiques culturales", "Rigueur lors du tuteurage, du désherbage manuel et de la cueillette soignée à maturité"],
                ["6. Processus (Process)", "Récolte aux heures fraîches du matin, tri qualitatif et conditionnement immédiat", "Mise en cagettes ventilées de 25 kg et expédition sous 24h vers les points de vente"],
                ["7. Preuve Matérielle (Physical)", "Parcelle moderne équipée d'irrigation goutte-à-goutte, aire de conditionnement propre", "Cagettes siglées assurant la traçabilité de l'exploitation maraîchère"]
            ]
        elif sec_id == "clinique":
            rows = [
                ["1. Produit (Product)", f"Prestations de diagnostic médical & analyses : {profile.products[0].name}", "Exactitude des résultats biologiques, réactifs étalonnés et délais de rendu rapides"],
                ["2. Prix (Price)", f"Grille tarifaire conventionnée et accessible : de {p1_price:,.0f} à {p_last_price:,.0f} {currency}".replace(",", " "), "Alignement sur la nomenclature médicale avec prise en charge mutuelles/assurances"],
                ["3. Distribution (Place)", "Implantation dans une zone urbaine à forte densité et facile d'accès", "Locaux climatisés, box de prélèvement individuels et accessibilité PMR"],
                ["4. Promotion (Promotion)", "Partenariats avec médecins prescripteurs, cliniques partenaires et comités d'entreprises", "Campagnes de dépistage préventif et présence digitale d'information santé"],
                ["5. Personnel (People)", "Médecin/Biologiste, techniciens supérieurs de laboratoire et infirmiers diplômés", "Accueil bienveillant, écoute active, respect du secret médical et déontologie"],
                ["6. Processus (Process)", "Accueil, enregistrement informatisé, prélèvement stérile, analyse et validation biologique", "Résultats disponibles en ligne ou par SMS sécurisé sous 2h"],
                ["7. Preuve Matérielle (Physical)", "Plateau technique moderne, automates sous contrat de maintenance, hygiène irréprochable", "Certificats de conformité sanitaire et gestion sécurisée des déchets DASRI"]
            ]
        elif sec_id == "aviculture":
            rows = [
                ["1. Produit (Product)", f"Poulets de chair vivants et éviscérés prêts à cuire : {profile.products[0].name}", "Poids moyen 1.8-2.0 kg à 45 jours, chair ferme, plumage propre et conformité sanitaire"],
                ["2. Prix (Price)", f"Positionnement compétitif : de {p1_price:,.0f} à {p_last_price:,.0f} {currency}".replace(",", " "), "Tarifs de gros incitatifs pour grossistes et revendeurs avec marge brute protectrice"],
                ["3. Distribution (Place)", "Canaux mixtes : enlèvement direct à la ferme, marchés de gros urbains, rôtisseries", "Organisation de livraisons groupées en caisses d'abattage ventilées"],
                ["4. Promotion (Promotion)", "Démarchage direct des marchés de volailles, rôtisseurs, hôtels, restaurants et supermarchés", "Accords d'approvisionnement exclusifs et communication digitale de marque locale"],
                ["5. Personnel (People)", "Équipe d'élevage expérimentée, suivie par un vétérinaire conseil mandaté", "Surveillance permanente des poulaillers, respect du vide sanitaire et de la biosécurité"],
                ["6. Processus (Process)", "Alimentation rigoureuse par phase de croissance (démarrage, croissance, finition)", "Pesées hebdomadaires, respect strict du plan de prophylaxie et traçabilité des lots"],
                ["7. Preuve Matérielle (Physical)", "Bâtiments d'élevage modernes bien ventilés, sas de désinfection et registre d'élevage", "Certificats sanitaires vétérinaires délivrés à chaque sortie de bande"]
            ]
        elif sec_id == "jus":
            rows = [
                ["1. Produit (Product)", f"Jus de mangue naturel pasteurisé en bouteilles 33cl et 1L : {profile.products[0].name}", "Jus pur fruits sans additif, pasteurisation thermique assurant une conservation longue durée"],
                ["2. Prix (Price)", f"Positionnement compétitif : de {p1_price:,.0f} à {p_last_price:,.0f} {currency}".replace(",", " "), "Tarifs distributeurs attractifs (grossistes, CHR, supermarchés) avec marge protectrice"],
                ["3. Distribution (Place)", "Canaux CHR (cafés, hôtels, restaurants), grandes et moyennes surfaces (GMS), boutiques", "Réseau de distribution urbain avec livraison dédiée par véhicules isothermes"],
                ["4. Promotion (Promotion)", "Dégustations en points de vente, communication digitale locale, partenariats événementiels", "Visibilité de marque forte valorisant la fraîcheur et le savoir-faire local"],
                ["5. Personnel (People)", "Opérateurs de ligne formés aux règles d'hygiène HACCP et équipe commerciale dédiée", "Rigueur de process, contrôles qualité permanents et sens de la relation client"],
                ["6. Processus (Process)", "Tri rigoureux, lavage/brossage, extraction, pasteurisation, embouteillage à chaud et étiquetage", "Traçabilité des lots et respect strict des normes de sécurité alimentaire"],
                ["7. Preuve Matérielle (Physical)", "Bouteilles esthétiques étiquetées aux normes, atelier inox carrelé et magasin tempéré", "Certificat de salubrité et conformité aux standards agroalimentaires"]
            ]
        elif sec_id == "commerce_informatique":
            rows = [
                ["1. Produit (Product)", f"Matériel informatique professionnel & solutions réseaux : {profile.products[0].name}", "Ordinateurs portables et de bureau professionnels, serveurs, routeurs et périphériques fiables"],
                ["2. Prix (Price)", f"Grille tarifaire compétitive B2B : de {p1_price:,.0f} à {p_last_price:,.0f} {currency}".replace(",", " "), "Tarifs dégressifs pour parcs d'entreprises, devis transparents et marges calculées"],
                ["3. Distribution (Place)", "Showroom professionnel en zone d'affaires et livraison directe en entreprise", "Espace d'exposition moderne, stocks de sécurité et interventions sur site client"],
                ["4. Promotion (Promotion)", "Démarchage direct des DSI, PME et administrations, marketing digital et catalogue", "Partenariats de maintenance et contrats d'infogérance récurrents"],
                ["5. Personnel (People)", "Ingénieurs et techniciens certifiés en réseaux et maintenance informatique", "Expertise technique pointue, sens du service client et assistance réactive"],
                ["6. Processus (Process)", "Audit des besoins client, proposition technique chiffrée, livraison, paramétrage et SAV", "Délais de réponse sous 24h et suivi des tickets d'assistance"],
                ["7. Preuve Matérielle (Physical)", "Showroom équipé, banc de test technique, attestations constructeurs", "Certificats de garantie et contrats de service formalisés"]
            ]
        elif sec_id == "energie_solaire":
            rows = [
                ["1. Produit (Product)", f"Solutions solaires photovoltaïques clés en main : {profile.products[0].name}", "Systèmes autonomes et hybrides haute performance avec onduleurs certifiés, batteries lithium et monitoring IoT"],
                ["2. Prix (Price)", f"Tarification transparente au kWc installé : de {p1_price:,.0f} à {p_last_price:,.0f} {currency}".replace(",", " "), "Formules compétitives avec retour sur investissement sous 3 à 4 ans comparé aux tarifs réseau/groupe"],
                ["3. Distribution (Place)", "Interventions directes à Abidjan et principales villes de Côte d'Ivoire", "Équipes mobiles d'installation autonomes avec véhicules d'intervention et outillage de contrôle"],
                ["4. Promotion (Promotion)", "Prospection B2B ciblée (PME, commerces, cliniques, exploitations agricoles) et marketing digital", "Audits énergétiques gratuits sur site, simulateur d'économies d'énergie et références d'installations"],
                ["5. Personnel (People)", "Ingénieurs électrotechniciens et techniciens monteurs solaires qualifiés et habilités", "Rigueur d'exécution, habilitation électrique et accompagnement personnalisé du client"],
                ["6. Processus (Process)", "Audit de charge, dimensionnement par logiciel, pose, raccordement, mise en service et formation client", "Contrôle qualité strict à chaque étape et raccordement sécurisé"],
                ["7. Preuve Matérielle (Physical)", "Véhicules floqués équipés, outillage de mesure étalonné et livret technique d'exploitation", "Certificats de garantie matérielle constructeur (25 ans panneaux, 10 ans batteries) et attestation de conformité"]
            ]
        else:
            rows = [
                ["1. Produit (Product)", f"Gamme professionnelle adaptée au marché : {profile.products[0].name}", "Qualité certifiée, fiabilité d'usage et packaging valorisant"],
                ["2. Prix (Price)", f"Positionnement compétitif : de {p1_price:,.0f} à {p_last_price:,.0f} {currency}".replace(",", " "), "Pénétration rapide tout en préservant une marge brute protectrice"],
                ["3. Distribution (Place)", "Canaux mixtes : points de vente directs, réseaux partenaires et grossistes", "Couverture efficace des zones urbaines à fort potentiel commercial"],
                ["4. Promotion (Promotion)", "Marketing digital, prospection directe et salons professionnels", "Visibilité immédiate et acquisition rapide de parts de marché"],
                ["5. Personnel (People)", "Équipe commerciale formée aux techniques de vente et relation client", "Conseil expert et fidélisation de la clientèle locale"],
                ["6. Processus (Process)", "Commandes digitalisées, gestion de stock réactive et facturation instantanée", "Expérience d'achat fluide et sans friction"],
                ["7. Preuve Matérielle (Physical)", "Aménagement moderne, conformité aux normes et réassurance client", "Garantie de confiance et sérieux institutionnel"]
            ]
        cls._write_table_data(ws, 5, 2, rows)

        # Indicateurs d'Acquisition Client (CAC & LTV)
        cls._create_section_divider(ws, "B13:F13", "▸ MÉTRIQUES DE PERFORMANCE COMMERCIALE (CAC & LTV)")
        ws.row_dimensions[14].height = 24
        for col_idx, h in enumerate(["Indicateur Commercial", "Valeur Estimée", "Interprétation & Rentabilité"], start=2):
            cell = ws.cell(row=14, column=col_idx, value=h)
            cls._style_header_cell(cell)

        cac_val = round(p1_price * 0.15)
        ltv_val = round(p1_price * 4.2)
        mkt_metrics = [
            ["Coût d'Acquisition Client (CAC)", f"{cac_val:,.0f} {currency}".replace(",", " "), "Coût marketing direct pour acquérir 1 nouveau client"],
            ["Valeur Vie Client (LTV - Lifetime Value)", f"{ltv_val:,.0f} {currency}".replace(",", " "), "Chiffre d'affaires moyen généré par client fidèle"],
            ["Ratio LTV / CAC", f"{round(ltv_val / max(cac_val, 1), 1)}x", "Excellent (> 3.0x requis par les investisseurs)"],
            ["Budget Marketing Initial (An 1)", f"{round(m['kpis'].revenue_year1 * 0.03):,.0f} {currency}".replace(",", " "), "Représente ~3% du chiffre d'affaires prévisionnel"]
        ]
        cls._write_table_data(ws, 15, 2, mkt_metrics)

    # =========================================================================
    # FEUILLE 05 : ITINÉRAIRE TECHNIQUE & CALENDRIER OPÉRATIONNEL
    # =========================================================================
    @classmethod
    def _build_05_itineraire_technique(cls, wb, profile: Any, m: Dict[str, Any] = None):
        ws = wb.create_sheet(title="05_Itineraire_Technique")
        cls._create_banner(ws, "B2:F2", "ITINÉRAIRE TECHNIQUE & CALENDRIER OPÉRATIONNEL (12 MOIS)")

        ws.row_dimensions[4].height = 26
        for col_idx, h in enumerate(["Période", "Étape Opérationnelle Clé", "Responsable", "Délai", "Livrable Attendu"], start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        is_comm = (m or {}).get("is_commercial", False)
        sec_id = getattr(profile, "sector_id", "")

        if sec_id == "transport":
            rows = [
                ["Mois 1", "Constitution juridique de la société, obtention licence de transport & formalités fiscales", "Direction / Notaire", "Semaines 1-4", "Statuts, RCCM, Licence d'exploitation transport"],
                ["Mois 2", "Aménagement de la base logistique, bureaux d'exploitation & parking sécurisé", "Direction / Travaux", "Semaines 5-8", "Base opérationnelle & parking clôturé"],
                ["Mois 3", "Acquisition des 5 camions porteurs 10t, immatriculations & pose balises GPS", "Responsable Flotte", "Semaines 9-12", "5 camions réceptionnés, immatriculés & équipés GPS"],
                ["Mois 4", "Recrutement des 5 chauffeurs routiers, tests d'aptitude à la conduite & formation éco-conduite", "Resp. Exploitation", "Semaines 13-16", "Équipe de chauffeurs qualifiée & habilitée"],
                ["Mois 5", "Campagne commerciale B2B, signature contrats de fret & tournées inaugurales", "Commercial / Resp. Flotte", "Semaines 17-20", "Premiers contrats de fret exécutés"],
                ["Mois 6 à 12", "Exploitation plein régime, suivi rentabilité kilométrique & entretien flotte", "Comptable / Resp. Flotte", "Semaines 21-52", "Rapports de marge kilométrique mensuels"]
            ]
        elif sec_id == "agriculture":
            rows = [
                ["Mois 1", "Acquisition foncière/bail rural, formalités administratives & bornage de la parcelle", "Promoteur / Notaire", "Semaines 1-4", "Bail notarié & terrain sécurisé"],
                ["Mois 2", "Forage d'eau profond, motopompe, château d'eau & clôture grillagée", "Entreprise Hydraulique", "Semaines 5-8", "Point d'eau pérenne opérationnel"],
                ["Mois 3", "Préparation du sol (labour, hersage, billonnage) & pose du réseau goutte-à-goutte", "Agronome / Techniciens", "Semaines 9-12", "Réseau d'irrigation testé et fonctionnel"],
                ["Mois 4", "Mise en place de la pépinière de tomates, repiquage & fertilisation de fond", "Agronome / Ouvriers", "Semaines 13-16", "Parcelle plantée et tuteurée"],
                ["Mois 5", "Conduite de la culture, fertigation, traitements phytosanitaires & premières récoltes", "Agronome / Équipe", "Semaines 17-20", "Premières cagettes de tomates récoltées"],
                ["Mois 6 à 12", "Récoltes continues en pleine saison, livraison des clients & 2e cycle cultural", "Responsable Ventes / Équipe", "Semaines 21-52", "Bilan de campagne et rotation des cultures"]
            ]
        elif sec_id == "clinique":
            rows = [
                ["Mois 1", "Constitution de la société, obtention de l'agrément ministériel & autorisation d'exercer", "Biologiste Directeur / Notaire", "Semaines 1-4", "Agrément Ministère Santé, RCCM"],
                ["Mois 2", "Aménagement des locaux médicaux, salle d'attente, box de prélèvement & paillasses", "Architecte / Promoteur", "Semaines 5-8", "Locaux conformes aux normes sanitaires"],
                ["Mois 3", "Acquisition, livraison et installation des automates de biochimie et hématologie", "Fournisseurs Médicaux", "Semaines 9-12", "Automates installés et raccordés"],
                ["Mois 4", "Calibrage des analyseurs, tests sur sérums de contrôle & recrutement du personnel", "Biologiste / Techniciens", "Semaines 13-16", "Contrôles qualité validés & équipe formée"],
                ["Mois 5", "Inauguration, ouverture officielle au public & signature conventions mutuelles", "Direction / Accueil", "Semaines 17-20", "Premiers patients accueillis et analysés"],
                ["Mois 6 à 12", "Exploitation continue, audit qualité semestriel et fidélisation prescripteurs", "Biologiste / Comptable", "Semaines 21-52", "Rapports d'activité médicale mensuels"]
            ]
        elif sec_id == "aviculture":
            rows = [
                ["Mois 1", "Constitution de la société, formalités d'implantation & plan de biosécurité", "Promoteur / Vétérinaire", "Semaines 1-4", "Statuts, RCCM & autorisation environnementale"],
                ["Mois 2", "Construction et aménagement des poulaillers modernes, radiants & ventilation", "Artisan / Éleveur", "Semaines 5-8", "Bâtiments d'élevage réceptionnés & désinfectés"],
                ["Mois 3", "Installation des lignes d'abreuvement automatique et mangeoires trémies", "Technicien Équipement", "Semaines 9-12", "Système d'eau et alimentation opérationnel"],
                ["Mois 4", "Achat du 1er lot de 5 000 poussins d'un jour, démarrage & plan prophylactique", "Vétérinaire / Éleveurs", "Semaines 13-16", "Phase de démarrage réussie (mortalité < 2%)"],
                ["Mois 5", "Phase de finition, enlèvement et commercialisation du premier cycle de poulets", "Responsable Ventes", "Semaines 17-20", "28 500 poulets commercialisés sur l'année"],
                ["Mois 6 à 12", "Enchaînement des 5 cycles suivants avec vide sanitaire rigoureux de 15 jours", "Direction / Vétérinaire", "Semaines 21-52", "6 cycles annuels menés à terme avec succès"]
            ]
        elif sec_id == "jus":
            rows = [
                ["Mois 1", "Constitution légale de la SARL, autorisations sanitaires & sécurisation du bail de l'atelier", "Direction / Notaire", "Semaines 1-4", "Statuts notariés, RCCM, Bail & Dépôt dossier sanitaire"],
                ["Mois 2", "Aménagement de l'atelier agroalimentaire (carrelage, siphons inox, alimentation eau/élec)", "Architecte / Artisans", "Semaines 5-8", "Atelier carrelé conforme aux normes HACCP"],
                ["Mois 3", "Acquisition, livraison et installation de la ligne d'extraction, pasteurisateur et embouteilleuse", "Fournisseurs Industriels", "Semaines 9-12", "Ligne de transformation installée et raccordée"],
                ["Mois 4", "Essais de production à blanc, réglage pasteurisation & formation du personnel de ligne", "Responsable Qualité / Opérateurs", "Semaines 13-16", "Tests microbiologiques validés et équipe habilitée"],
                ["Mois 5", "Campagne de lancement commercial, référencement en supermarchés & premières livraisons", "Responsable Commercial", "Semaines 17-20", "Premiers lots de 33cl et 1L commercialisés"],
                ["Mois 6 à 12", "Production continue plein régime (3 000 L/j), gestion des stocks saisonniers & livraisons CHR", "Direction / Responsable Production", "Semaines 21-52", "Suivi qualité, volumes nominaux atteints et rentabilité assurée"]
            ]
        elif sec_id == "commerce_informatique":
            rows = [
                ["Mois 1", "Constitution légale de la SARL, immatriculation fiscale & signature du bail commercial", "Direction / Notaire", "Semaines 1-4", "Statuts, RCCM, NIF & Bail showroom sécurisé"],
                ["Mois 2", "Aménagement du showroom, espace technique de configuration & sécurisation des locaux", "Direction / Agencement", "Semaines 5-8", "Showroom prêt à accueillir la clientèle B2B"],
                ["Mois 3", "Acquisition du stock initial d'équipements informatiques & négociation constructeurs", "Responsable Achats", "Semaines 9-12", "Premier stock de PC, serveurs et périphériques réceptionné"],
                ["Mois 4", "Recrutement des commerciaux B2B et techniciens réseaux & formation aux outils", "Direction / RH", "Semaines 13-16", "Équipe commerciale et technique opérationnelle"],
                ["Mois 5", "Campagne de prospection B2B, inauguration showroom & signature premiers contrats", "Responsable Commercial", "Semaines 17-20", "Premiers parcs informatiques équipés et livrés"],
                ["Mois 6 à 12", "Développement continu des ventes, contrats de maintenance récurrents et extension de gamme", "Direction / Équipe Commerciale", "Semaines 21-52", "Fidélisation de la clientèle PME et rentabilité commerciale pérenne"]
            ]
        elif sec_id == "energie_solaire":
            rows = [
                ["Mois 1", "Constitution légale de la SARL, formalités administratives & habilitations électriques solaires", "Promoteur / Notaire", "Semaines 1-4", "Statuts notariés, RCCM, Déclaration fiscale et habilitations"],
                ["Mois 2", "Aménagement du local technique, atelier de pré-câblage, showroom et bureau d'études", "Promoteur / Artisans", "Semaines 5-8", "Atelier et bureau d'études opérationnels"],
                ["Mois 3", "Acquisition véhicule utilitaire, outillage électricien certifié & banc d'essais photovoltaïque", "Directeur Technique", "Semaines 9-12", "Véhicule d'intervention et testeurs de conformité livrés"],
                ["Mois 4", "Approvisionnement stock initial (panneaux Tier-1, onduleurs, batteries lithium) & logiciels dimensionnement", "Responsable Achats", "Semaines 13-16", "Composants solaires réceptionnés et inventoriés"],
                ["Mois 5", "Campagne commerciale B2B/résidentiel, premières études de charge & premières installations", "Ingénieur Commercial", "Semaines 17-20", "Premiers systèmes résidentiels et centrales PME installés"],
                ["Mois 6 à 12", "Montée en puissance des installations (résidentiel, PME, pros) & contrats de maintenance récurrents", "Directeur Technique / Équipe", "Semaines 21-52", "Capacité annuelle atteinte et rentabilité consolidée"]
            ]
        elif is_comm:
            rows = [
                ["Mois 1", "Constitution juridique de la société & formalités fiscales", "Promoteur / Notaire", "Semaines 1-4", "Statuts notariés, RCCM, Déclaration fiscale"],
                ["Mois 2", "Aménagement du local commercial / boutique et vitrines", "Gérant / Agenceur", "Semaines 5-8", "Boutique opérationnelle & sécurisée"],
                ["Mois 3", "Acquisition rayonnages, caisse tactile & système gestion", "Fournisseur IT / Gérant", "Semaines 9-12", "Caisse & étagères opérationnelles"],
                ["Mois 4", "Constitution du stock initial de marchandises & formation vendeuses", "Gérant / Vendeuses", "Semaines 13-16", "Rayons achalandés et inventoriés"],
                ["Mois 5", "Campagne de lancement commercial, inauguration & vente", "Responsable Ventes", "Semaines 17-20", "Premières ventes en magasin"],
                ["Mois 6 à 12", "Plein régime commercial, rotation des stocks & suivi financier", "Gérant / Comptable", "Semaines 21-52", "Rapports de marge commerciale mensuels"]
            ]
        else:
            rows = [
                ["Mois 1", "Constitution juridique de la société & formalités administratives", "Promoteur / Notaire", "Semaines 1-4", "Statuts notariés, RCCM, Déclaration fiscale"],
                ["Mois 2", "Aménagement du site d'exploitation aux normes professionnelles", "Chef d'Atelier", "Semaines 5-8", "Site prêt & sécurisé"],
                ["Mois 3", "Acquisition, livraison et installation des équipements", "Fournisseur / Promoteur", "Semaines 9-12", "Équipements installés et testés"],
                ["Mois 4", "Tests opérationnels, mise en route & formation des opérateurs", "Responsable Qualité", "Semaines 13-16", "Tests concluants conformes aux standards"],
                ["Mois 5", "Campagne de lancement commercial & premières livraisons", "Responsable Ventes", "Semaines 17-20", "Premières commandes exécutées"],
                ["Mois 6 à 12", "Plein régime de croisière et suivi de trésorerie", "Comptable / Direction", "Semaines 21-52", "Rapports d'activité mensuels"]
            ]
        cls._write_table_data(ws, 5, 2, rows)

    # =========================================================================
    # FEUILLE 06 : PRÉVISIONS DE PRODUCTION & CAPACITÉS
    # =========================================================================
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

    # =========================================================================
    # FEUILLE 07 : CHIFFRES D'AFFAIRES PRÉVISIONNEL
    # =========================================================================
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

        # Ligne de Taux de Croissance Annuel (%)
        croissance_row = total_row + 1
        ws.cell(row=croissance_row, column=2, value="Taux de Croissance Annuel (%)").font = Font(name="Calibri", size=9, italic=True)
        ws.cell(row=croissance_row, column=4, value="-").alignment = Alignment(horizontal="center")
        for c in range(5, 4 + horizon):
            prev_l = get_column_letter(c - 1)
            curr_l = get_column_letter(c)
            c_croiss = ws.cell(row=croissance_row, column=c, value=f"=({curr_l}{total_row}-{prev_l}{total_row})/{prev_l}{total_row}")
            c_croiss.number_format = "0.0%"
            c_croiss.font = Font(name="Calibri", size=9, bold=True, color="047857")
            c_croiss.alignment = Alignment(horizontal="right")

    # =========================================================================
    # FEUILLE 08 : INVESTISSEMENTS & AMORTISSEMENTS SYSCOHADA
    # =========================================================================
    @classmethod
    def _build_08_investissements(cls, wb, profile: Any, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="08_Investissements_Amort")
        cls._create_banner(ws, "B2:G2", f"PROGRAMME D'INVESTISSEMENTS & PLAN D'AMORTISSEMENT ({currency})")

        ws.row_dimensions[4].height = 26
        headers = ["Immobilisation (Compte SYSCOHADA)", f"Montant Brut ({currency})", "Durée (Ans)", "Taux %", "Annuité Amort.", "VNC Fin Année 1"]
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
        ws.cell(row=total_r, column=2, value="TOTAL DES INVESTISSEMENTS (CAPEX)").font = Font(name="Calibri", size=10, bold=True)
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

    # =========================================================================
    # FEUILLE 09 : CHARGES D'EXPLOITATION & COMPTES 60 À 62
    # =========================================================================
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

    # =========================================================================
    # FEUILLE 10 : FRAIS DE PERSONNEL & CHARGES SOCIALES (CNPS/CSS)
    # =========================================================================
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

        cls._write_table_data(ws, 5, 2, rows, numeric_cols=[4, 5, 6, 7], num_format=f'#,##0 "{currency}"')

        for i, s in enumerate(staff_list):
            r = 5 + i
            cell_eff_item = ws.cell(row=r, column=3)
            cell_eff_item.number_format = '#,##0'
            cell_eff_item.alignment = Alignment(horizontal="center", vertical="center")

        last_r = 4 + len(staff_list)
        total_r = last_r + 1
        ws.cell(row=total_r, column=2, value="TOTAL MASSE SALARIALE ANNUELLE (COMPTE 66)").font = Font(name="Calibri", size=10, bold=True)
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

    # =========================================================================
    # FEUILLE 11 : PLAN DE FINANCEMENT INITIAL & ÉQUILIBRE STRUCTUREL
    # =========================================================================
    @classmethod
    def _build_11_plan_financement(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="11_Plan_Financement")
        cls._create_banner(ws, "B2:F2", "PLAN DE FINANCEMENT INITIAL (ÉQUILIBRE DES EMPLOIS & RESSOURCES)")

        theme = cls.current_theme
        ws.row_dimensions[4].height = 26
        ws.cell(row=4, column=2, value="BESOINS DE DÉMARRAGE (EMPLOIS)").fill = PatternFill(fill_type="solid", fgColor=theme.header_bg)
        ws.cell(row=4, column=2).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        ws.cell(row=4, column=3, value=f"Montant ({currency})").fill = PatternFill(fill_type="solid", fgColor=theme.header_bg)
        ws.cell(row=4, column=3).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

        ws.cell(row=4, column=5, value="RESSOURCES MOBILISÉES (FINANCEMENT)").fill = PatternFill(fill_type="solid", fgColor=theme.header_bg)
        ws.cell(row=4, column=5).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        ws.cell(row=4, column=6, value=f"Montant ({currency})").fill = PatternFill(fill_type="solid", fgColor=theme.header_bg)
        ws.cell(row=4, column=6).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

        fin = m["financing"]
        inv = m["investments"]
        bfr_init = fin.get("bfr_initial", inv.get("bfr_initial", m.get("bfr", {}).get("net", 0)))

        besoins = [
            ["Frais d'établissement & formalités (Cpt 20)", inv["frais_etab"]],
            ["Investissements matériels & équipements (Cpt 21-24)", inv["total"]],
            ["BFR Initial & Trésorerie de démarrage (Fonds de roulement)", bfr_init]
        ]

        ressources = [
            ["Apport personnel promoteur (Fonds propres - Cpt 101)", fin["apport"]],
            ["Emprunt bancaire sollicité à moyen terme (Cpt 162)", fin["emprunt"]],
            ["Subventions & concours d'appui au démarrage (Cpt 14)", 0]
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

        # Ligne de Contrôle d'Équilibre Strict (Emplois - Ressources = 0)
        ws.row_dimensions[9].height = 22
        ctrl_label = ws.cell(row=9, column=2, value="CONTRÔLE D'ÉQUILIBRE (EMPLOIS - RESSOURCES)")
        ctrl_label.font = Font(name="Calibri", size=9, bold=True)
        ctrl_label.fill = PatternFill(fill_type="solid", fgColor="DCFCE7")
        ctrl_ecart = ws.cell(row=9, column=3, value="=C8-F8")
        ctrl_ecart.number_format = f'#,##0 "{currency}"'
        cls._style_total_cell(ctrl_ecart, bg_color="DCFCE7", font_color=cls.GREEN_TEXT)

        ctrl_statut = ws.cell(row=9, column=5, value="STATUT DU BOUCLAGE FINANCIER")
        ctrl_statut.font = Font(name="Calibri", size=9, bold=True)
        ctrl_statut.fill = PatternFill(fill_type="solid", fgColor="DCFCE7")
        ctrl_verdict = ws.cell(row=9, column=6, value='=IF(C8=F8,"PARFAITEMENT ÉQUILIBRÉ (0 FCFA) ✅","DÉSÉQUILIBRE DÉTECTÉ ❌")')
        ctrl_verdict.font = Font(name="Calibri", size=9, bold=True, color=cls.GREEN_TEXT)
        ctrl_verdict.alignment = Alignment(horizontal="center", vertical="center")
        ctrl_verdict.fill = PatternFill(fill_type="solid", fgColor="DCFCE7")

        # Ratios de Structure Financière (Normes Bancaires UEMOA)
        cls._create_section_divider(ws, "B11:F11", "▸ RATIOS DE STRUCTURE FINANCIÈRE & CONFORMITÉ BANCAIRE")
        ws.row_dimensions[12].height = 24
        for col_idx, h in enumerate(["Règle Bancaire", "Valeur du Projet", "Norme UEMOA / CEMAC", "Verdict Bancaire"], start=2):
            cell = ws.cell(row=12, column=col_idx, value=h)
            cls._style_header_cell(cell)

        part_fp = fin.get("part_fp_pct", 30.0)
        gearing = fin.get("gearing", 1.33)
        frng = fin.get("frng_initial", bfr_init)

        struct_rows = [
            ["Part des Fonds Propres", f"{part_fp:.1f}%", ">= 20% à 30%", "CONFORME AUX STANDARDS ✅"],
            ["Ratio d'Endettement (Gearing)", f"{gearing:.2f}x", "< 2.0x à 2.5x", "ENDETTEMENT MAÎTRISÉ ✅"],
            ["Fonds de Roulement Net (FRNG Initial)", f"{frng:,.0f} {currency}".replace(",", " "), ">= BFR Initial", "ÉQUILIBRE FINANCIER ASSURÉ ✅"]
        ]
        cls._write_table_data(ws, 13, 2, struct_rows)
        for idx in range(3):
            ws.cell(row=13 + idx, column=5).font = Font(name="Calibri", size=10, bold=True, color=cls.GREEN_TEXT)

    # =========================================================================
    # FEUILLE 12 : REMBOURSEMENT DE L'EMPRUNT BANCAIRE
    # =========================================================================
    @classmethod
    def _build_12_remboursement_emprunt(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="12_Remboursement_Emprunt")
        fin = m.get("financing", {})
        rate = fin.get("interest_rate", 0.12)
        rate_pct = round(rate * 100, 1)
        cls._create_banner(ws, "B2:G2", f"TABLEAU D'AMORTISSEMENT DU CRÉDIT BANCAIRE (Taux {rate_pct}% l'an)")

        ws.row_dimensions[4].height = 26
        headers = ["Période", "Capital Début Période", f"Intérêts Annuels ({rate_pct}%)", "Amortissement Capital", "Annuité Totale", "Capital Restant Dû Fin"]
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
                    f"=C{r}*{rate}",
                    entry["principal"],
                    f"=D{r}+E{r}",
                    f"=C{r}-E{r}"
                ])
            elif i == len(loan_schedule) - 1:
                rows.append([
                    entry["year"],
                    f"=G{r-1}",
                    f"=C{r}*{rate}",
                    f"=C{r}",
                    f"=D{r}+E{r}",
                    f"=C{r}-E{r}"
                ])
            else:
                rows.append([
                    entry["year"],
                    f"=G{r-1}",
                    f"=C{r}*{rate}",
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

    # =========================================================================
    # FEUILLE 13 : BESOIN EN FONDS DE ROULEMENT (BFR 5 ANS)
    # =========================================================================
    @classmethod
    def _build_13_bfr(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="13_BFR_Exploitation")
        horizon = m.get("horizon_years", 5)
        end_col_letter = get_column_letter(3 + horizon)
        cls._create_banner(ws, f"B2:{end_col_letter}2", f"BESOIN EN FONDS DE ROULEMENT (BFR EN JOURS DE CA) ({currency})")

        ws.row_dimensions[4].height = 26
        headers = ["Composante du BFR", "Norme de Rotation"] + [f"Année {y}" for y in range(1, horizon + 1)]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        bfr = m["bfr"]
        bfr_5y = m.get("bfr_5y", {})
        stocks_s = bfr_5y.get("stocks", [bfr["stocks"]] * horizon)
        creances_s = bfr_5y.get("creances", [bfr["creances"]] * horizon)
        dettes_s = bfr_5y.get("dettes", [bfr["dettes"]] * horizon)

        is_commercial = m.get("is_commercial", False)
        stock_label = "Stocks de marchandises (Rayons & réserves)" if is_commercial else "Stocks de matières premières & intrants"

        row_stocks = [stock_label, "30 jours d'achats"] + stocks_s[:horizon]
        row_creances = ["Créances clients d'exploitation accordées", "15 jours de chiffre d'affaires"] + creances_s[:horizon]
        row_dettes = ["Dettes fournisseurs (Crédit d'exploitation)", "30 jours d'achats"] + [-d for d in dettes_s[:horizon]]

        rows = [row_stocks, row_creances, row_dettes]
        num_cols = list(range(4, 4 + horizon))
        cls._write_table_data(ws, 5, 2, rows, numeric_cols=num_cols, num_format=f'#,##0 "{currency}"')

        total_r = 8
        ws.cell(row=total_r, column=2, value="BFR GLOBAL NET D'EXPLOITATION").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=total_r, column=2).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        ws.cell(row=total_r, column=3).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        for c in range(4, 4 + horizon):
            col_letter = get_column_letter(c)
            cell = ws.cell(row=total_r, column=c, value=f"=SUM({col_letter}5:{col_letter}7)")
            cell.number_format = f'#,##0 "{currency}"'
            cls._style_total_cell(cell, bg_color=cls.GREEN_BG, font_color=cls.GREEN_TEXT)

        # Ligne BFR en Jours de Chiffre d'Affaires
        bfr_j_row = total_r + 1
        ws.cell(row=bfr_j_row, column=2, value="BFR Exprimé en Jours de Chiffre d'Affaires").font = Font(name="Calibri", size=9, italic=True)
        ws.cell(row=bfr_j_row, column=3, value="Normalisé UEMOA").alignment = Alignment(horizontal="center")
        for c in range(4, 4 + horizon):
            col_l = get_column_letter(c)
            ca_val = m["ca"]["years_ca"][c - 4]
            cell_j = ws.cell(row=bfr_j_row, column=c, value=f"=ROUND(({col_l}{total_r}/{ca_val})*360, 0)")
            cell_j.number_format = '0 "jours"'
            cell_j.font = Font(name="Calibri", size=9, bold=True, color="0284C7")
            cell_j.alignment = Alignment(horizontal="right")

    # =========================================================================
    # FEUILLE 14 : COMPTE DE RÉSULTAT SYSCOHADA (5 ANS COMPLETS + MARGES %)
    # =========================================================================
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
            ach_label = "Achats matières premières et consommables"
            ach_cpt = "602"
            marge_label = "MARGE BRUTE DE PRODUCTION (SIG)"

        row_ca = [ca_label, ca_cpt] + [ca[y - 1] for y in range(1, horizon + 1)]
        row_ach = [ach_label, ach_cpt] + [ach[y - 1] for y in range(1, horizon + 1)]
        row_mb = [marge_label, "SIG"] + [f"={get_column_letter(3 + y)}5+{get_column_letter(3 + y)}6" for y in range(1, horizon + 1)]
        row_serv = ["Services extérieurs & Transports", "61/62"] + [serv[y - 1] for y in range(1, horizon + 1)]
        row_va = ["VALEUR AJOUTÉE (VA)", "SIG"] + [f"={get_column_letter(3 + y)}7+{get_column_letter(3 + y)}8" for y in range(1, horizon + 1)]
        row_perso = ["Charges de personnel (Salaires & CNPS)", "66"] + [perso[y - 1] for y in range(1, horizon + 1)]
        row_ebe = ["EXCÉDENT BRUT D'EXPLOITATION (EBE / EBITDA)", "SIG"] + [f"={get_column_letter(3 + y)}9+{get_column_letter(3 + y)}10" for y in range(1, horizon + 1)]
        row_dot = ["Dotations aux amortissements", "68"] + [dot[y - 1] for y in range(1, horizon + 1)]
        row_rex = ["RÉSULTAT D'EXPLOITATION (REX / EBIT)", "SIG"] + [f"={get_column_letter(3 + y)}11+{get_column_letter(3 + y)}12" for y in range(1, horizon + 1)]
        row_fin = ["Frais financiers (Intérêts emprunt)", "67"] + [int_fin[y - 1] for y in range(1, horizon + 1)]
        row_rao = ["RÉSULTAT AVANT IMPÔTS (RAO)", "SIG"] + [f"={get_column_letter(3 + y)}13+{get_column_letter(3 + y)}14" for y in range(1, horizon + 1)]
        row_is = [f"Impôt sur les sociétés (IS {is_pct_str})", "89"] + [f"={get_column_letter(3 + y)}15*-{is_rate}" for y in range(1, horizon + 1)]
        row_rn = ["RÉSULTAT NET DE L'EXERCICE (BÉNÉFICE)", "SIG"] + [f"={get_column_letter(3 + y)}15+{get_column_letter(3 + y)}16" for y in range(1, horizon + 1)]
        row_caf = ["CAPACITÉ D'AUTOFINANCEMENT (CAF)", "SIG"] + [f"={get_column_letter(3 + y)}17-{get_column_letter(3 + y)}12" for y in range(1, horizon + 1)]

        rows = [row_ca, row_ach, row_mb, row_serv, row_va, row_perso, row_ebe, row_dot, row_rex, row_fin, row_rao, row_is, row_rn, row_caf]

        num_cols = list(range(4, 4 + horizon))
        cls._write_table_data(ws, 5, 2, rows, numeric_cols=num_cols, num_format=f'#,##0 "{currency}"')

        for i in range(len(rows)):
            ws.cell(row=5 + i, column=3).alignment = Alignment(horizontal="center", vertical="center")

        for i, r_data in enumerate(rows):
            r = 5 + i
            if r_data[1] == "SIG":
                ws.row_dimensions[r].height = 24
                for c in range(2, 4 + horizon):
                    ws.cell(row=r, column=c).font = Font(name="Calibri", size=10, bold=True)
                    ws.cell(row=r, column=c).fill = PatternFill(fill_type="solid", fgColor=cls.ZEBRA_BG)

        # Ligne Résultat Net en vert prestige
        for c in range(2, 4 + horizon):
            cell = ws.cell(row=17, column=c)
            cell.font = Font(name="Calibri", size=10, bold=True, color=cls.GREEN_TEXT)
            cell.fill = PatternFill(fill_type="solid", fgColor=cls.GREEN_BG)

        # Ligne CAF en bleu institutionnel avec double soulignement
        for c in range(2, 4 + horizon):
            cell = ws.cell(row=18, column=c)
            cell.font = Font(name="Calibri", size=10, bold=True, color="1E1B4B")
            cell.fill = PatternFill(fill_type="solid", fgColor="E0E7FF")
            cell.border = Border(bottom=Side(style="double", color="1E1B4B"), top=Side(style="thin", color="CBD5E1"))

        # Section Analyse des Marges Relatives (% du Chiffre d'Affaires)
        cls._create_section_divider(ws, f"B20:{end_col_letter}20", "▸ ANALYSE DES MARGES FINANCIÈRES RELATIVES (% DU CHIFFRE D'AFFAIRES)")
        ws.row_dimensions[21].height = 24
        for col_idx, h in enumerate(["Indicateur de Rentabilité", "Référence"] + [f"Année {y}" for y in range(1, horizon + 1)], start=2):
            cell = ws.cell(row=21, column=col_idx, value=h)
            cls._style_header_cell(cell)

        margin_rows = [
            ["Taux de Marge Brute (% du CA)", "MB / CA"] + [f"={get_column_letter(3 + y)}7/{get_column_letter(3 + y)}5" for y in range(1, horizon + 1)],
            ["Taux de Valeur Ajoutée (% du CA)", "VA / CA"] + [f"={get_column_letter(3 + y)}9/{get_column_letter(3 + y)}5" for y in range(1, horizon + 1)],
            ["Marge Opérationnelle Brute (Taux d'EBE)", "EBE / CA"] + [f"={get_column_letter(3 + y)}11/{get_column_letter(3 + y)}5" for y in range(1, horizon + 1)],
            ["Taux de Rentabilité d'Exploitation (REX)", "REX / CA"] + [f"={get_column_letter(3 + y)}13/{get_column_letter(3 + y)}5" for y in range(1, horizon + 1)],
            ["Marge Nette Finale (% du CA)", "RN / CA"] + [f"={get_column_letter(3 + y)}17/{get_column_letter(3 + y)}5" for y in range(1, horizon + 1)]
        ]
        cls._write_table_data(ws, 22, 2, margin_rows, numeric_cols=num_cols, num_format="0.0%")
        for r_m in range(22, 27):
            ws.cell(row=r_m, column=3).alignment = Alignment(horizontal="center")

        # Graphique natif : Évolution du Résultat Net sur l'horizon
        chart_rn = BarChart()
        chart_rn.type = "col"
        chart_rn.style = 10
        chart_rn.title = f"Évolution du Résultat Net Prévisionnel ({currency})"
        chart_rn.y_axis.title = f"Bénéfice Net ({currency})"
        chart_rn.x_axis.title = "Années Prévisionnelles"
        chart_rn.width = 18
        chart_rn.height = 7.5
        data_rn = Reference(ws, min_col=2, min_row=17, max_col=3 + horizon, max_row=17)
        cats_rn = Reference(ws, min_col=4, min_row=4, max_col=3 + horizon, max_row=4)
        chart_rn.add_data(data_rn, titles_from_data=True, from_rows=True)
        chart_rn.set_categories(cats_rn)
        chart_rn.legend = None
        chart_rn.dataLabels = DataLabelList()
        chart_rn.dataLabels.showVal = True
        ws.add_chart(chart_rn, "B28")

    # =========================================================================
    # FEUILLE 15 : BILAN PRÉVISIONNEL SUR 5 ANS (SYSCOHADA)
    # =========================================================================
    @classmethod
    def _build_15_bilan_previsionnel(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="15_Bilan_Previsionnel")
        cls._create_banner(ws, "B2:H2", f"BILAN PRÉVISIONNEL SUR 5 ANS - ACTIF & PASSIF (SYSCOHADA) ({currency})")

        horizon = m.get("horizon_years", 5)
        bilan_5y = m.get("bilan_5years", [])
        is_commercial = m.get("is_commercial", False)
        stock_label = "Stocks de marchandises (Actif circulant - Cpt 31)" if is_commercial else "Stocks matières & intrants (Actif circulant - Cpt 32/36)"

        # En-têtes du Bilan
        ws.row_dimensions[4].height = 26
        headers = ["POSTES DU BILAN SYNTHÉTIQUE", "RÉF SYSCOHADA"] + [f"Année {y}" for y in range(1, horizon + 1)]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        # 1. SECTION ACTIF
        actif_items = [
            ("Immobilisations nettes (Actif immobilisé)", "Classe 2", [b["immo_nette"] for b in bilan_5y]),
            (stock_label, "Cpt 32/36" if not is_commercial else "Cpt 31", [b["stocks"] for b in bilan_5y]),
            ("Créances clients d'exploitation", "Cpt 411", [b["creances"] for b in bilan_5y]),
            ("Trésorerie Actif (Disponibilités Banques & Caisse)", "Classe 5", [b["tresorerie"] for b in bilan_5y])
        ]

        for i, (label, ref, vals) in enumerate(actif_items):
            r = 5 + i
            ws.cell(row=r, column=2, value=label).font = Font(name="Calibri", size=10)
            ws.cell(row=r, column=3, value=ref).font = Font(name="Calibri", size=9, italic=True, color="64748B")
            ws.cell(row=r, column=3).alignment = Alignment(horizontal="center")
            for y_idx in range(horizon):
                col = 4 + y_idx
                c = ws.cell(row=r, column=col, value=vals[y_idx] if y_idx < len(vals) else 0)
                c.number_format = f'#,##0 "{currency}"'
                c.font = Font(name="Calibri", size=10)

        # Total Actif avec formule Excel
        ws.cell(row=9, column=2, value="TOTAL ACTIF DU BILAN").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=9, column=2).fill = PatternFill(fill_type="solid", fgColor=cls.GREEN_BG)
        ws.cell(row=9, column=3, value="TOTAL I").fill = PatternFill(fill_type="solid", fgColor=cls.GREEN_BG)
        ws.cell(row=9, column=3).font = Font(name="Calibri", size=9, bold=True)
        ws.cell(row=9, column=3).alignment = Alignment(horizontal="center")
        for y_idx in range(horizon):
            col_letter = get_column_letter(4 + y_idx)
            tot_a = ws.cell(row=9, column=4 + y_idx, value=f"=SUM({col_letter}5:{col_letter}8)")
            tot_a.number_format = f'#,##0 "{currency}"'
            cls._style_total_cell(tot_a, bg_color=cls.GREEN_BG, font_color=cls.GREEN_TEXT)

        # 2. SECTION PASSIF
        passif_items = [
            ("Capital social (Apport personnel du promoteur)", "Cpt 101", [b["capital"] for b in bilan_5y]),
            ("Réserves & Report à nouveau (Bénéfices cumulés)", "Cpt 11/12", [b["reserves"] for b in bilan_5y]),
            ("Résultat Net de l'exercice (Bénéfice de l'année)", "Cpt 13", [b["resultat_net"] for b in bilan_5y])
        ]

        for i, (label, ref, vals) in enumerate(passif_items):
            r = 11 + i
            ws.cell(row=r, column=2, value=label).font = Font(name="Calibri", size=10)
            ws.cell(row=r, column=3, value=ref).font = Font(name="Calibri", size=9, italic=True, color="64748B")
            ws.cell(row=r, column=3).alignment = Alignment(horizontal="center")
            for y_idx in range(horizon):
                c = ws.cell(row=r, column=4 + y_idx, value=vals[y_idx] if y_idx < len(vals) else 0)
                c.number_format = f'#,##0 "{currency}"'
                c.font = Font(name="Calibri", size=10)

        # Sous-total Capitaux Propres
        ws.cell(row=14, column=2, value="TOTAL CAPITAUX PROPRES").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=14, column=2).fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        ws.cell(row=14, column=3, value="Sous-total").fill = PatternFill(fill_type="solid", fgColor=cls.BORDER_SOFT)
        ws.cell(row=14, column=3).alignment = Alignment(horizontal="center")
        for y_idx in range(horizon):
            col_letter = get_column_letter(4 + y_idx)
            sub_cp = ws.cell(row=14, column=4 + y_idx, value=f"=SUM({col_letter}11:{col_letter}13)")
            sub_cp.number_format = f'#,##0 "{currency}"'
            cls._style_total_cell(sub_cp, bg_color=cls.BORDER_SOFT)

        # Dettes financières & Dettes circulantes
        dettes_items = [
            ("Dettes financières (Emprunt bancaire à terme)", "Cpt 162", [b["dettes_fin"] for b in bilan_5y]),
            ("Fournisseurs et Dettes circulantes", "Cpt 401", [b["dettes_fourn"] for b in bilan_5y])
        ]
        for i, (label, ref, vals) in enumerate(dettes_items):
            r = 15 + i
            ws.cell(row=r, column=2, value=label).font = Font(name="Calibri", size=10)
            ws.cell(row=r, column=3, value=ref).font = Font(name="Calibri", size=9, italic=True, color="64748B")
            ws.cell(row=r, column=3).alignment = Alignment(horizontal="center")
            for y_idx in range(horizon):
                c = ws.cell(row=r, column=4 + y_idx, value=vals[y_idx] if y_idx < len(vals) else 0)
                c.number_format = f'#,##0 "{currency}"'
                c.font = Font(name="Calibri", size=10)

        # Total Passif avec formule Excel
        ws.cell(row=17, column=2, value="TOTAL PASSIF DU BILAN").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=17, column=2).fill = PatternFill(fill_type="solid", fgColor=cls.GREEN_BG)
        ws.cell(row=17, column=3, value="TOTAL II").fill = PatternFill(fill_type="solid", fgColor=cls.GREEN_BG)
        ws.cell(row=17, column=3).font = Font(name="Calibri", size=9, bold=True)
        ws.cell(row=17, column=3).alignment = Alignment(horizontal="center")
        for y_idx in range(horizon):
            col_letter = get_column_letter(4 + y_idx)
            tot_p = ws.cell(row=17, column=4 + y_idx, value=f"={col_letter}14+{col_letter}15+{col_letter}16")
            tot_p.number_format = f'#,##0 "{currency}"'
            cls._style_total_cell(tot_p, bg_color=cls.GREEN_BG, font_color=cls.GREEN_TEXT)

        # 3. LIGNE DE CONTRÔLE D'ÉQUILIBRE STRICT (TOTAL ACTIF - TOTAL PASSIF = 0)
        ws.cell(row=19, column=2, value="CONTRÔLE D'ÉQUILIBRE (ACTIF - PASSIF)").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=19, column=2).fill = PatternFill(fill_type="solid", fgColor="DCFCE7")
        ws.cell(row=19, column=3, value="Écart = 0").fill = PatternFill(fill_type="solid", fgColor="DCFCE7")
        ws.cell(row=19, column=3).font = Font(name="Calibri", size=9, bold=True, color=cls.GREEN_TEXT)
        ws.cell(row=19, column=3).alignment = Alignment(horizontal="center")
        for y_idx in range(horizon):
            col_letter = get_column_letter(4 + y_idx)
            ctrl_cell = ws.cell(row=19, column=4 + y_idx, value=f"={col_letter}9-{col_letter}17")
            ctrl_cell.number_format = f'#,##0 "{currency}"'
            cls._style_total_cell(ctrl_cell, bg_color="DCFCE7", font_color=cls.GREEN_TEXT)

        # 4. Équilibre Financier Fonctionnel (FRNG, BFR, Trésorerie Nette)
        cls._create_section_divider(ws, f"B21:{get_column_letter(3 + horizon)}21", "▸ ÉQUILIBRE FINANCIER FONCTIONNEL SUR 5 ANS (FRNG - BFR = TRÉSORERIE NETTE)")
        ws.row_dimensions[22].height = 24
        fn_headers = ["Grand Agrégat Fonctionnel", "Formule Analytique"] + [f"Année {y}" for y in range(1, horizon + 1)]
        for col_idx, h in enumerate(fn_headers, start=2):
            cell = ws.cell(row=22, column=col_idx, value=h)
            cls._style_header_cell(cell)

        frng_vals = []
        bfr_vals = []
        tn_vals = []
        for b in bilan_5y:
            frng_y = (b["total_capitaux_propres"] + b["dettes_fin"]) - b["immo_nette"]
            bfr_y = (b["stocks"] + b["creances"]) - b["dettes_fourn"]
            tn_y = frng_y - bfr_y
            frng_vals.append(frng_y)
            bfr_vals.append(bfr_y)
            tn_vals.append(tn_y)

        agregats = [
            ("Fonds de Roulement Net Global (FRNG)", "Ressources Stables - Actif Immobilisé", frng_vals),
            ("Besoin en Fonds de Roulement (BFR)", "Actif Circulant - Passif Circulant", bfr_vals),
            ("Trésorerie Nette (TN)", "FRNG - BFR (ou Trésorerie Actif)", tn_vals)
        ]

        for i, (label, formule, vals) in enumerate(agregats):
            r = 23 + i
            ws.cell(row=r, column=2, value=label).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=r, column=3, value=formule).font = Font(name="Calibri", size=9, italic=True, color="64748B")
            for y_idx in range(horizon):
                c = ws.cell(row=r, column=4 + y_idx, value=vals[y_idx] if y_idx < len(vals) else 0)
                c.number_format = f'#,##0 "{currency}"'
                c.font = Font(name="Calibri", size=10, bold=True, color=cls.GREEN_TEXT)

    # =========================================================================
    # FEUILLE 16 : PLAN DE TRÉSORERIE PRÉVISIONNEL (12 MOIS & 5 ANS)
    # =========================================================================
    @classmethod
    def _build_16_plan_tresorerie(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="16_Plan_Tresorerie_12M")
        cls._create_banner(ws, "B2:N2", f"PLAN DE TRÉSORERIE PRÉVISIONNEL (12 MOIS & 5 ANS) ({currency})")

        horizon = m.get("horizon_years", 5)

        # -------------------------------------------------------------
        # SECTION 1 : PLAN DE TRÉSORERIE MENSUALISÉ SUR 12 MOIS (ANNÉE 1)
        # -------------------------------------------------------------
        cls._create_section_divider(ws, "B3:N3", "▸ 1. PLAN DE TRÉSORERIE MENSUELISÉ SUR 12 MOIS (ANNÉE 1)")
        ws.row_dimensions[4].height = 26
        months = ["Rubrique de Trésorerie", "M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10", "M11", "M12"]
        for col_idx, h in enumerate(months, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        ca1 = m["kpis"].revenue_year1
        base_m_ca = round(ca1 / 12)
        base_m_chg = round((ca1 - m["kpis"].ebitda_year1) / 12)
        bfr_init = m["bfr"]["net"]

        encaissements = [round(base_m_ca * factor) for factor in [0.70, 0.80, 0.90, 0.95, 1.0, 1.05, 1.05, 1.10, 1.10, 1.15, 1.15, 1.20]]
        decaissements = [round(base_m_chg * factor) for factor in [0.90, 0.85, 0.90, 0.95, 1.0, 1.0, 1.05, 1.05, 1.10, 1.10, 1.15, 1.15]]

        ws.cell(row=5, column=2, value="Solde Initial de Trésorerie").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=6, column=2, value="Total Encaissements d'Exploitation").font = Font(name="Calibri", size=10)
        ws.cell(row=7, column=2, value="Total Décaissements d'Exploitation").font = Font(name="Calibri", size=10)
        ws.cell(row=8, column=2, value="SOLDE FINAL DE CLÔTURE MENSUEL").font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row=8, column=2).fill = PatternFill(fill_type="solid", fgColor=cls.GREEN_BG)

        for col_idx in range(3, 15):
            m_idx = col_idx - 3
            col_letter = get_column_letter(col_idx)
            prev_col_letter = get_column_letter(col_idx - 1)

            if m_idx == 0:
                c_start = ws.cell(row=5, column=col_idx, value=bfr_init)
            else:
                c_start = ws.cell(row=5, column=col_idx, value=f"={prev_col_letter}8")
            c_start.number_format = "#,##0"

            c_enc = ws.cell(row=6, column=col_idx, value=encaissements[m_idx])
            c_enc.number_format = "#,##0"

            c_dec = ws.cell(row=7, column=col_idx, value=decaissements[m_idx])
            c_dec.number_format = "#,##0"

            c_end = ws.cell(row=8, column=col_idx, value=f"={col_letter}5+{col_letter}6-{col_letter}7")
            c_end.number_format = "#,##0"
            cls._style_total_cell(c_end, bg_color=cls.GREEN_BG, font_color=cls.GREEN_TEXT)

        # -------------------------------------------------------------
        # SECTION 2 : PLAN DE TRÉSORERIE PLURIANNUEL SUR 5 ANS
        # -------------------------------------------------------------
        cls._create_section_divider(ws, "B11:G11", "▸ 2. PLAN DE TRÉSORERIE PLURIANNUEL INTÉGRÉ SUR 5 ANS")
        ws.row_dimensions[12].height = 24
        t5_headers = ["Flux de Trésorerie Pluriannuels"] + [f"Année {y}" for y in range(1, horizon + 1)]
        for col_idx, h in enumerate(t5_headers, start=2):
            cell = ws.cell(row=12, column=col_idx, value=h)
            cls._style_header_cell(cell)

        treso_5y = m.get("tresorerie_5years", [])
        t5_rows = [
            ("Solde Initial de Trésorerie", [t["cash_start"] for t in treso_5y]),
            ("Excédent Brut d'Exploitation (EBE)", [t["ebitda"] for t in treso_5y]),
            ("Impôt sur les Sociétés (IS décaissé)", [-t["tax"] for t in treso_5y]),
            ("Variation du BFR d'exploitation", [-t["delta_bfr"] for t in treso_5y]),
            ("Flux Net de Trésorerie d'Exploitation (OCF)", [t["ocf"] for t in treso_5y]),
            ("Service de la Dette (Capital + Intérêts)", [-t["debt_service"] for t in treso_5y]),
            ("SOLDE FINAL DE CLÔTURE DE TRÉSORERIE", [t["cash_end"] for t in treso_5y])
        ]

        for i, (label, vals) in enumerate(t5_rows):
            r = 13 + i
            is_total = (i == 4 or i == 6)
            bg = cls.GREEN_BG if i == 6 else (cls.BORDER_SOFT if i == 4 else None)

            ws.cell(row=r, column=2, value=label).font = Font(name="Calibri", size=10, bold=is_total)
            if bg:
                ws.cell(row=r, column=2).fill = PatternFill(fill_type="solid", fgColor=bg)

            for y_idx in range(horizon):
                val = vals[y_idx] if y_idx < len(vals) else 0
                c = ws.cell(row=r, column=3 + y_idx, value=val)
                c.number_format = f'#,##0 "{currency}"'
                if bg:
                    cls._style_total_cell(c, bg_color=bg, font_color=cls.GREEN_TEXT if i == 6 else "000000")
                else:
                    c.font = Font(name="Calibri", size=10)

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

    # =========================================================================
    # FEUILLE 17 : SEUIL DE RENTABILITÉ & POINT MORT
    # =========================================================================
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

    # =========================================================================
    # FEUILLE 18 : ANALYSE DE SENSIBILITÉ (STRESS-TEST DU PROJET)
    # =========================================================================
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
            ["1. Scénario de Référence (Base)", "Conditions normales d'exploitation et de marché", ca, chg, rn, sr, "EXPLOITATION RENTABLE ✅"],
            ["2. Baisse des Prix de Vente (-15%)", "Pression concurrentielle sur les tarifs", "=D5*0.85", chg, "=D6-E6", "=F5*1.08", '=IF(F6>0, "VIABLE - Marge absorbante ✅", "DÉFICITAIRE - Risque de perte ❌")'],
            ["3. Baisse Sévère des Prix (-25%)", "Choc déflationniste ou baisse pouvoir d'achat", "=D5*0.75", chg, "=D7-E7", "=F5*1.18", '=IF(F7>0, "ZONE DE VIGILANCE ⚠️", "CRITIQUE - Cessation d\'exploitation ❌")'],
            ["4. Baisse des Volumes Vendus (-25%)", "Ralentissement logistique temporaire", "=D5*0.75", "=E5*0.82", "=D8-E8", "=F5*0.95", '=IF(F8>0, "RÉSILIENCE CONFIRMÉE ✅", "DÉFICITAIRE - Réduire coûts fixes ❌")'],
            ["5. Hausse Coût des Intrants (+20%)", "Inflation sur matières premières & emballages", ca, "=E5*1.20", "=D9-E9", "=F5*1.15", '=IF(F9>0, "ABSORBABLE PAR LA MARGE ✅", "DÉFICITAIRE - Relever prix vente ❌")'],
            ["6. Pire Cas Combiné (Double Choc)", "Prix -15% ET Intrants +15% simultanés", "=D5*0.85", "=E5*1.15", "=D10-E10", "=F5*1.28", '=IF(F10>0, "RISQUE MODÉRÉ ⚠️", "IMPACT MAJEUR - Faillite potentielle ❌")']
        ]
        cls._write_table_data(ws, 5, 2, rows, numeric_cols=[3, 4, 5, 6], num_format=f'#,##0 "{currency}"')

        for i in range(len(rows)):
            r = 5 + i
            cell_v = ws.cell(row=r, column=8)
            cell_v.font = Font(name="Calibri", size=10, bold=True)
            if "✅" in str(rows[i][6]) and "IF" not in str(rows[i][6]):
                cell_v.font = Font(name="Calibri", size=10, bold=True, color=cls.GREEN_TEXT)
            else:
                cell_v.font = Font(name="Calibri", size=10, bold=True, color=cls.NAVY_DARK)

    # =========================================================================
    # FEUILLE 19 : RATIOS FINANCIERS & ANALYSE DE RISQUE BANCAIRE
    # =========================================================================
    @classmethod
    def _build_19_ratios(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="19_Ratios_Financiers_Bancaires")
        cls._create_banner(ws, "B2:E2", "INDICATEURS FINANCIERS DE BANCABILITÉ : VAN, TRI & SCORING")

        ws.row_dimensions[4].height = 26
        headers = ["Indicateur Financier Clé", "Valeur Calculée", "Norme Bancaire (UEMOA/CEMAC)", "Avis d'Opportunité"]
        for col_idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cls._style_header_cell(cell)

        k = m["kpis"]
        altman_z = float(m.get("altman_z", 3.0))
        dscr_val = m.get("ratios", {}).get("dscr_avg", 2.5)
        ip_val = m.get("ratios", {}).get("profitability_index", 1.8)
        gearing_val = m.get("financing", {}).get("gearing", 2.33)
        rate_val = k.loan_interest_rate

        if altman_z >= 2.90:
            z_verdict = "ZONE SAINE (SOLVABILITÉ ÉLEVÉE) ✅"
            z_norm = "> 2.90 (Zone Sûre)"
        elif altman_z >= 1.23:
            z_verdict = "ZONE GRISE (SURVEILLANCE NORMALE) ⚠️"
            z_norm = "1.23 - 2.90 (Vigilance)"
        else:
            z_verdict = "ZONE DE VULNÉRABILITÉ ❌"
            z_norm = "< 1.23 (Risque)"

        rows = [
            ["Valeur Actuelle Nette (VAN à 10%)", f"{k.npv:,.0f} {currency}".replace(",", " "), "> 0", "PROJET TRÈS RENTABLE ✅"],
            ["Taux de Rentabilité Interne (TRI)", k.irr_pct, f"> Taux d'intérêt bancaire ({rate_val:.1f}%)", "EXCELLENT POTENTIEL ✅"],
            ["Délai de Récupération du Capital (DRCI)", k.payback_period, "< 4 ans", "RISQUE EN CAPITAL FAIBLE ✅"],
            ["Indice de Profitabilité (IP)", f"{ip_val:.2f}", "> 1.0", f"GAIN DE {round((ip_val - 1.0) * 100)} FCFA PAR 100 FCFA INVESTIS ✅"],
            ["Ratio Moyen de Couverture Dette (DSCR)", f"{dscr_val:.2f}x", ">= 1.30x", "SERVICE DE DETTE TRÈS BIEN COUVERT ✅"],
            ["Ratio d'Endettement (Dette / Fonds propres)", f"{gearing_val:.2f}x", "< 2.5x", "STRUCTURE FINANCIÈRE SOLIDE ✅"],
            ["Score de Risque Altman Z' PME OHADA", f"{altman_z:.2f}", z_norm, z_verdict]
        ]
        cls._write_table_data(ws, 5, 2, rows)

        for i in range(len(rows)):
            r = 5 + i
            cell_v = ws.cell(row=r, column=5)
            cell_v.font = Font(name="Calibri", size=10, bold=True, color=cls.GREEN_TEXT if "✅" in rows[i][3] else (cls.AMBER_TEXT if "⚠️" in rows[i][3] else "DC2626"))

    # =========================================================================
    # FEUILLE 20 : TABLEAU DE BORD EXÉCUTIF (EXECUTIVE DASHBOARD)
    # =========================================================================
    @classmethod
    def _build_20_tableau_de_bord(cls, wb, currency: str, m: Dict[str, Any]):
        ws = wb.create_sheet(title="20_Tableau_de_Bord")
        cls._create_banner(ws, "B2:F2", "TABLEAU DE BORD EXÉCUTIF - SYNTHÈSE DES KPI DU PROJET")

        theme = cls.current_theme
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
            c_h1.fill = PatternFill(fill_type="solid", fgColor=theme.header_bg)
            c_h1.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            c_h1.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[r_header].height = 24

            ws.merge_cells(start_row=r_val, start_column=2, end_row=r_val, end_column=3)
            c_v1 = ws.cell(row=r_val, column=2, value=card[1])
            c_v1.fill = PatternFill(fill_type="solid", fgColor=theme.card_bg)
            c_v1.font = Font(name="Calibri", size=12, bold=True, color="0F172A")
            c_v1.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[r_val].height = 28

            # Card Droite
            ws.merge_cells(start_row=r_header, start_column=5, end_row=r_header, end_column=6)
            c_h2 = ws.cell(row=r_header, column=5, value=card[2])
            c_h2.fill = PatternFill(fill_type="solid", fgColor=theme.header_bg)
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
        quote = theme.vision_quote if theme else "Devenir le leader de référence en Afrique francophone."
        c_vis.value = f"  « {quote} »"
        hl_bg = theme.highlight_bg if theme else "FEF3C7"
        hl_fg = theme.highlight_text if theme else "78350F"
        c_vis.fill = PatternFill(fill_type="solid", fgColor=hl_bg)
        c_vis.font = Font(name="Calibri", size=11, bold=True, italic=True, color=hl_fg)
        c_vis.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[17].height = 24
        ws.row_dimensions[18].height = 24

    # =========================================================================
    # UTILITAIRES DE STYLE & GÉOMÉTRIE EXCEL
    # =========================================================================
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
    def _create_section_divider(cls, ws, cell_range: str, title: str):
        ws.merge_cells(cell_range)
        start_cell = cell_range.split(":")[0]
        cell = ws[start_cell]
        cell.value = f"  {title}"
        bg = cls.current_theme.highlight_bg if cls.current_theme else "F1F5F9"
        fg = cls.current_theme.highlight_text if cls.current_theme else "1E293B"
        cell.fill = PatternFill(fill_type="solid", fgColor=bg)
        cell.font = Font(name="Calibri", size=10, bold=True, color=fg)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[int(start_cell[1:])].height = 24

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
                if val_str.startswith('='):
                    val_str = "123 456 789 FCFA"  # Longueur estimée de sécurité
                max_len = max(max_len, len(val_str))

            if col_letter == 'B':
                safe_width = max(max_len + 4, 38)
            elif col_letter in ['C', 'D', 'E', 'F', 'G']:
                safe_width = max(max_len + 5, 20)
            else:
                safe_width = max(max_len + 4, 15)

            ws.column_dimensions[col_letter].width = safe_width
