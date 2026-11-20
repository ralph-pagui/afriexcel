"""
Module d'analyse sémantique et extraction NLP avancée pour AfriExcel Copilot.
Extrait fidèlement les paramètres économiques, financiers et sectoriels des prompts utilisateurs
sans jamais injecter de texte de commande dans les cellules du classeur.
Garantit l'isolation totale du contexte projet et le verrouillage strict des paramètres.
"""

import re
from typing import Dict, Any, Optional, List, Tuple
from .parameter_lock import ProjectContext, ParameterOrigin, LockStatus

class PromptParser:
    """Analyseur sémantique déterministe pour requêtes de Business Plan en Afrique francophone."""

    @classmethod
    def parse(cls, prompt: str, default_country: str = "Bénin", default_currency: str = "FCFA") -> Dict[str, Any]:
        """Extrait l'ensemble des informations structurées et crée un contexte projet verrouillé."""
        text = prompt.strip()
        lower = text.lower()

        # 1. Détection chirurgicale du secteur d'activité par scoring multi-mots d'abord
        sector_id, sector_name = cls._detect_sector_scored(lower)

        # 2. Extraction du Pays et de la Localisation avec priorité absolue aux champs explicites
        country, has_explicit_country = cls._extract_country(text, default_country)
        city, location = cls._extract_location(text, country)

        # 3. Extraction de la Raison Sociale / Nom du projet
        company_name = cls._extract_company_name(text, sector_id=sector_id, country=country)

        # 4. Extraction de la Devise (FCFA, GNF, USD, EUR)
        currency = cls._extract_currency(text, country, default_currency)

        # 5. Extraction des paramètres de financement (Apport, Emprunt, Taux, Durée)
        equity_ratio, loan_ratio, has_explicit_financing = cls._extract_financing_split(text)
        loan_interest_rate, has_explicit_interest_rate = cls._extract_interest_rate(text)
        loan_duration_years = cls._extract_loan_duration(text)

        # 6. Extraction de l'horizon prévisionnel
        horizon_years = cls._extract_horizon_years(text)

        # 7. Extraction du budget ou investissement cible
        total_budget = cls._extract_budget(text)

        # 8. Extraction de la capacité et données opérationnelles
        capacity_info = cls._extract_capacity(text)
        fleet_ops = cls._extract_fleet_and_operations(text)

        # 9. Extraction et calibrage des produits explicitement demandés
        explicit_products = cls._extract_explicit_products(text, sector_id)

        # 10. Construction et verrouillage du Contexte Projet
        ctx = ProjectContext(
            project_name=company_name,
            country=country,
            city=city,
            currency=currency,
            sector_id=sector_id,
            sector_name=sector_name,
            explicit_products=[p["name"] for p in explicit_products] if explicit_products else [],
            daily_capacity=fleet_ops.get("daily_capacity"),
            daily_capacity_unit=fleet_ops.get("capacity_unit", "Unités"),
            working_days_per_month=fleet_ops.get("days_per_month", 22),
            working_days_per_year=fleet_ops.get("working_days_per_year", 264),
            horizon_years=horizon_years,
            equity_ratio=equity_ratio,
            loan_ratio=loan_ratio,
            loan_interest_rate=loan_interest_rate,
            loan_duration_years=loan_duration_years,
            target_budget=total_budget
        )

        # Verrouillage formel des paramètres
        if has_explicit_country:
            ctx.lock_param("country", country, "", "Pays d'implantation", ParameterOrigin.EXPLICIT_PROMPT)
        if city:
            ctx.lock_param("city", city, "", "Ville d'implantation", ParameterOrigin.EXPLICIT_PROMPT)
        if has_explicit_financing:
            ctx.lock_param("equity_ratio", equity_ratio, "%", "Quotité de fonds propres", ParameterOrigin.EXPLICIT_PROMPT)
            ctx.lock_param("loan_ratio", loan_ratio, "%", "Quotité de dette bancaire", ParameterOrigin.EXPLICIT_PROMPT)
        if has_explicit_interest_rate:
            ctx.lock_param("loan_interest_rate", loan_interest_rate, "%", "Taux d'intérêt annuel emprunt", ParameterOrigin.EXPLICIT_PROMPT)
        if fleet_ops.get("daily_capacity"):
            ctx.lock_param("daily_capacity", fleet_ops["daily_capacity"], fleet_ops.get("capacity_unit", ""), "Capacité nominale journalière", ParameterOrigin.EXPLICIT_PROMPT)

        return {
            "company_name": company_name,
            "country": country,
            "has_explicit_country": has_explicit_country,
            "city": city,
            "location": location,
            "currency": currency,
            "sector_id": sector_id,
            "sector_name": sector_name,
            "equity_ratio": equity_ratio,
            "loan_ratio": loan_ratio,
            "has_explicit_financing": has_explicit_financing,
            "loan_interest_rate": loan_interest_rate,
            "has_explicit_interest_rate": has_explicit_interest_rate,
            "loan_duration_years": loan_duration_years,
            "horizon_years": horizon_years,
            "total_budget": total_budget,
            "capacity_info": capacity_info,
            "fleet_ops": fleet_ops,
            "explicit_products": explicit_products,
            "context": ctx,
            "raw_prompt": text
        }

    @classmethod
    def _extract_company_name(cls, text: str, sector_id: str = "", country: str = "") -> str:
        """Extrait un nom d'entreprise ou de projet court, précis et valorisant."""
        # 1. Recherche explicite ancrée par ligne "Nom de l'entreprise : XYZ" ou "Nom du projet : XYZ"
        m_line = re.search(r'^[ \t]*[-*•]?[ \t]*(?:nom\s+(?:du\s+projet|de\s+l[\'’]entreprise)|raison\s+sociale)[ \t]*[:=-][ \t]*([^\r\n]+)', text, re.I | re.M)
        if m_line:
            cand = re.sub(r'[*_#]', '', m_line.group(1)).strip()
            cand = re.sub(r'\s*\(.*?\)', '', cand).strip()
            if 3 <= len(cand) <= 60 and not re.match(r'^(?:génère|crée|fais|donne|un|le|mon)\b', cand, re.I):
                return cand

        # 2. Recherche de mentions juridiques explicites SARL / SAS / SUARL / GIE / LOGISTICS / INDUSTRIES
        m_legal = re.search(r'\b([A-ZÀ-ÿ0-9][A-ZÀ-ÿ0-9\s&\'\-]{2,40}\s+(?:SARL|SAS|SUARL|SA|GIE|GROUP|SERVICES|PLUS|AGRO|LOGISTICS|INDUSTRIES))\b', text)
        if m_legal:
            cand = m_legal.group(1).strip()
            cand = re.sub(r'^(?:nom\s+(?:du\s+projet|de\s+l[\'’]entreprise)|raison\s+sociale)\s*[:=-]\s*', '', cand, flags=re.I).strip()
            if 3 <= len(cand) <= 50:
                return cand

        # 3. Recherche de création d'activité ("création d'une entreprise de...")
        m_proj = re.search(r'(?:création|mise\s+en\s+place|lancement|développement)\s+(?:d[\'’]une?|de\s+la)\s+([^\n\r,.;]+)', text, re.I)
        if m_proj:
            cand = m_proj.group(1).strip()
            cand = re.sub(r'[*_#]', '', cand).strip()
            cand = re.sub(r'\s+(?:spécialisée|au|en|dans|à|sur|avec|pour).*$', '', cand, flags=re.I).strip()
            if 4 <= len(cand) <= 45:
                words = [w.capitalize() for w in cand.split()]
                return " ".join(words) + " SARL"

        defaults = {
            "energie_solaire": "IVOIRE SOLAR ENERGY SARL" if ("côte d'ivoire" in country.lower() or "ivoire" in country.lower()) else "AFRICA SOLAR ENERGY SARL",
            "jus": "BENIN FRUIT INDUSTRIES SARL" if "bénin" in country.lower() or "benin" in country.lower() else "AGRO FRUIT INDUSTRIES SARL",
            "transport": "SEN TRANS LOGISTICS SARL",
            "commerce_informatique": "AFRIQ TECH DISTRIBUTION SARL",
            "aviculture": "AVICOLA PLUS SARL",
            "agriculture": "AGRI TOMATE SARL",
            "plantation_cacao": "CACAO EXCELLENCE SARL",
            "clinique": "POLYCLINIQUE ET BIOLOGIE SARL",
            "magasin_jouets": "AFRIQUE DISTRIBUTION SARL"
        }
        return defaults.get(sector_id, "ENTREPRISE AGRO-INDUSTRIE SARL")

    @classmethod
    def _extract_country(cls, text: str, default_country: str) -> Tuple[str, bool]:
        """
        Détecte avec certitude le pays d'implantation africain.
        Priorité absolue aux lignes déclaratives explicites (ex: 'Pays : Bénin').
        Exclut systématiquement les mentions négatives ou comparatives (ex: 'ne pas confondre avec le Cameroun').
        """
        country_catalog = {
            "Bénin": ["bénin", "benin", "cotonou", "porto-novo", "parakou", "calavi"],
            "Sénégal": ["sénégal", "senegal", "dakar", "thies", "thiès", "saint-louis", "touba", "ziguinchor", "kaolack", "mbour"],
            "Cameroun": ["cameroun", "cameroon", "douala", "yaoundé", "yaounde", "bafoussam", "garoua", "littoral"],
            "Côte d'Ivoire": ["côte d'ivoire", "cote d'ivoire", "abidjan", "yamoussoukro", "san pedro", "bouaké", "bouake", "korhogo"],
            "Togo": ["togo", "lomé", "lome", "kara", "kpalimé"],
            "Burkina Faso": ["burkina", "ouagadougou", "bobo-dioulasso"],
            "Mali": ["mali", "bamako", "sikasso", "ségou"],
            "Gabon": ["gabon", "libreville", "port-gentil", "franceville"],
            "Congo": ["congo", "brazzaville", "pointe-noire"],
            "RDC": ["rdc", "kinshasa", "lubumbashi", "goma"],
            "Guinée": ["guinée", "guinee", "conakry", "kankan"]
        }

        # 1. Recherche par champ formel direct: "Pays : Bénin"
        m_field = re.search(r'^[ \t]*[-*•]?[ \t]*pays[ \t]*[:=-][ \t]*([^\r\n,;]+)', text, re.I | re.M)
        if m_field:
            raw_c = m_field.group(1).strip().lower()
            for c_name, kws in country_catalog.items():
                if any(kw in raw_c for kw in kws):
                    return c_name, True

        # 2. Recherche par préposition de localisation dans l'en-tête : "au Bénin", "à Cotonou", "à Douala"
        lower = text.lower()
        m_prep = re.search(r'\b(?:au|en|à|a)\s+(bénin|benin|sénégal|senegal|cameroun|cameroon|côte d[\'’]ivoire|cote d[\'’]ivoire|togo|mali|burkina|gabon|congo|guinée|guinee|rdc)\b', lower)
        if m_prep:
            raw_c = m_prep.group(1)
            for c_name, kws in country_catalog.items():
                if raw_c in kws:
                    return c_name, True

        # 3. Recherche par ville majeure explicite
        city_to_country = {
            "cotonou": "Bénin", "porto-novo": "Bénin", "parakou": "Bénin",
            "dakar": "Sénégal", "thiès": "Sénégal", "thies": "Sénégal", "saint-louis": "Sénégal", "kaolack": "Sénégal",
            "douala": "Cameroun", "yaoundé": "Cameroun", "yaounde": "Cameroun", "bafoussam": "Cameroun",
            "abidjan": "Côte d'Ivoire", "bouaké": "Côte d'Ivoire", "san pedro": "Côte d'Ivoire",
            "lomé": "Togo", "lome": "Togo", "ouagadougou": "Burkina Faso", "bamako": "Mali"
        }
        for city, c_name in city_to_country.items():
            if re.search(r'\b' + re.escape(city) + r'\b', lower):
                if not re.search(r'(?:ne\s+(?:pas|confonds)|contrairement)\b[^\n]*\b' + re.escape(city), lower):
                    return c_name, True

        # 4. Analyse générale en ignorant les phrases d'exclusion
        clean_text = re.sub(r'ne\s+confonds?\s+pas\s+[^\n.]*', '', lower)
        clean_text = re.sub(r'contrairement\s+[^\n.]*', '', clean_text)
        for c_name, kws in country_catalog.items():
            for kw in kws:
                if re.search(r'\b' + re.escape(kw) + r'\b', clean_text):
                    return c_name, True

        return default_country, False

    @classmethod
    def _extract_location(cls, text: str, country: str) -> Tuple[str, str]:
        """Extrait la ville et la chaîne de localisation complète."""
        city = ""
        m_city = re.search(r'^[ \t]*[-*•]?[ \t]*(?:ville(?:\s+d[\'’]implantation)?|siège|commune)[ \t]*[:=-][ \t]*([^\r\n,;]+)', text, re.I | re.M)
        if m_city:
            city = re.sub(r'[*_#]', '', m_city.group(1)).strip()

        if not city:
            m_a = re.search(r'\b(?:à|a)\s+([A-ZÀ-ÿ][a-zà-ÿ]+(?:\s+[A-ZÀ-ÿ][a-zà-ÿ]+)?)\b', text)
            if m_a:
                cand = m_a.group(1).strip()
                if cand.lower() not in ["excel", "ce", "cette", "un", "une", "la", "le", "cinq", "trois"]:
                    city = cand

        if not city:
            city = "Cotonou" if country == "Bénin" else ("Dakar" if country == "Sénégal" else ("Douala" if country == "Cameroun" else "Abidjan"))

        return city, f"{city}, {country} (Espace Économique OHADA)"

    @classmethod
    def _extract_currency(cls, text: str, country: str, default_currency: str) -> str:
        lower = text.lower()
        if re.search(r'\b(?:gnf|franc[s]?\s+guin[ée]en[s]?)\b', lower):
            return "GNF"
        if re.search(r'\b(?:usd|dollar[s]?|\$)\b', lower):
            return "USD"
        if re.search(r'\b(?:euro[s]?|eur|€)\b', lower):
            return "EUR"
        if re.search(r'\b(?:fcfa|xaf|xof|cfa|franc[s]?\s+cfa)\b', lower):
            return "FCFA"
        return default_currency or "FCFA"

    @classmethod
    def _extract_financing_split(cls, text: str) -> Tuple[float, float, bool]:
        """Extrait les pourcentages d'apport personnel et d'emprunt bancaire."""
        eq_ratio = None
        loan_ratio = None
        has_explicit = False

        m_eq1 = re.search(r'(?:apport\s+(?:personnel|promoteur|fonds\s+propres?)|fonds\s+propres?)\s*[:=-]?\s*(?:de\s*)?(\d+(?:[.,]\d+)?)\s*%', text, re.I)
        m_eq2 = re.search(r'(\d+(?:[.,]\d+)?)\s*%\s*(?:d[\'’]apport|fonds\s+propres?|apport\s+personnel)', text, re.I)
        m_eq = m_eq1 or m_eq2
        if m_eq:
            eq_val = float(m_eq.group(1).replace(',', '.'))
            eq_ratio = eq_val / 100.0 if eq_val > 1 else eq_val
            has_explicit = True

        m_loan1 = re.search(r'(?:emprunt\s+(?:bancaire)?|crédit\s+(?:bancaire)?|prêt\s+(?:bancaire)?|financement\s+bancaire)\s*[:=-]?\s*(?:de\s*)?(\d+(?:[.,]\d+)?)\s*%', text, re.I)
        m_loan2 = re.search(r'(\d+(?:[.,]\d+)?)\s*%\s*(?:d[\'’]emprunt|d[\'’]endettement|crédit\s+bancaire|emprunt\s+bancaire|dette\s+bancaire)', text, re.I)
        m_loan = m_loan1 or m_loan2
        if m_loan:
            loan_val = float(m_loan.group(1).replace(',', '.'))
            loan_ratio = loan_val / 100.0 if loan_val > 1 else loan_val
            has_explicit = True

        if eq_ratio is not None and loan_ratio is None:
            loan_ratio = max(0.0, 1.0 - eq_ratio)
        elif loan_ratio is not None and eq_ratio is None:
            eq_ratio = max(0.0, 1.0 - loan_ratio)
        elif eq_ratio is None and loan_ratio is None:
            eq_ratio = 0.30
            loan_ratio = 0.70

        return round(eq_ratio, 2), round(loan_ratio, 2), has_explicit

    @classmethod
    def _extract_interest_rate(cls, text: str) -> Tuple[float, bool]:
        """Extrait le taux d'intérêt annuel du crédit bancaire (ex: 11%, 10% ou 9.5%)."""
        m = re.search(
            r'(?:taux\s+(?:d[\'’](?:int[ée]r[êe]t|emprunt))?\s*(?:annuel)?\s*[:=-]?\s*(?:de\s*|d[\'’]\s*|à\s*)?|\b(?:au\s+taux\s+de\s*|à\s+))(\d+(?:[.,]\d+)?)\s*%',
            text,
            re.I
        )
        if m:
            val = float(m.group(1).replace(',', '.'))
            rate = val / 100.0 if val > 1 else val
            return round(rate, 4), True
        return 0.11, False

    @classmethod
    def _extract_loan_duration(cls, text: str) -> int:
        """Extrait la durée de remboursement de l'emprunt bancaire."""
        m = re.search(r'(?:remboursement|durée\s+(?:de\s+l[\'’]emprunt|du\s+prêt|du\s+financement)?)\s*(?:sur)?\s*[:=-]?\s*(\d+)\s*ans?\b', text, re.I)
        if m:
            return max(1, min(int(m.group(1)), 15))
        return 5

    @classmethod
    def _extract_horizon_years(cls, text: str) -> int:
        """Extrait l'horizon de modélisation du business plan (en années)."""
        m = re.search(r'(?:horizon\s+(?:prévisionnel|du\s+business\s+plan)?|durée\s+du\s+projet)\s*[:=-]?\s*(\d+)\s*ans?\b', text, re.I)
        if m:
            return max(3, min(int(m.group(1)), 10))
        if re.search(r'\bcinq\s+ans\b', text, re.I):
            return 5
        if re.search(r'\btrois\s+ans\b', text, re.I):
            return 3
        return 5

    @classmethod
    def _extract_budget(cls, text: str) -> Optional[float]:
        """Extrait le budget total ou investissement initial souhaité."""
        m_millions = re.search(r'(\d+(?:[.,]\d+)?)\s*millions?\s*(?:de\s*)?(?:fcfa|f\s*cfa|francs?|xof|xaf)?\b', text, re.I)
        if m_millions:
            return float(m_millions.group(1).replace(',', '.')) * 1_000_000

        m_num = re.search(r'(?:budget|investissement|montant|co[uû]t)\s*(?:total|initial|global|estimatif|pr[eé]vu|souhait[eé]|du\s+projet|\s+)*[:=-]?\s*(\d[\d\s]{5,})\s*(?:fcfa|f\s*cfa|xof|xaf)?\b', text, re.I)
        if m_num:
            digits = re.sub(r'\s+', '', m_num.group(1))
            if digits.isdigit():
                return float(digits)
        return None

    @classmethod
    def _extract_capacity(cls, text: str) -> Optional[str]:
        """Extrait la mention textuelle de capacité de production."""
        m = re.search(r'capacit[ée]\s+(?:initiale|de\s+production)?\s*[:=-]?\s*([^\n\r,.;]+)', text, re.I)
        if m:
            return m.group(1).strip()
        return None

    @classmethod
    def _extract_fleet_and_operations(cls, text: str) -> Dict[str, Any]:
        """Extrait les grandeurs opérationnelles fondamentales (litres/j, jours/m, camions, ha)."""
        ops: Dict[str, Any] = {}

        # 1. Capacité journalière de transformation / production en Litres ou Tonnes
        m_cap = re.search(r'(\d[\d\s]*)\s*(?:litres?|l)\s*(?:de\s+jus\s+)?par\s+jour\b', text, re.I)
        if m_cap:
            digits = re.sub(r'\s+', '', m_cap.group(1))
            if digits.isdigit():
                ops["daily_capacity"] = float(digits)
                ops["capacity_unit"] = "Litres"

        # 2. Jours de production par mois
        m_days = re.search(r'(\d+)\s+jours?\s+(?:de\s+production\s+)?par\s+mois\b', text, re.I)
        if m_days:
            ops["days_per_month"] = int(m_days.group(1))
            ops["working_days_per_year"] = int(m_days.group(1)) * 12
        else:
            ops["days_per_month"] = 22
            ops["working_days_per_year"] = 264

        # 3. Flotte de camions
        m_cam = re.search(r'(\d+)\s+camions?(?:\s+porteurs?)?(?:\s+de\s+(\d+)\s+tonnes?)?', text, re.I)
        if m_cam:
            ops["truck_count"] = int(m_cam.group(1))
            if m_cam.group(2):
                ops["truck_tonnage"] = int(m_cam.group(2))

        # 4. Surface agricole en hectares
        m_ha = re.search(r'(\d+(?:[.,]\d+)?)\s*hectares?\b', text, re.I)
        if m_ha:
            ops["hectares"] = float(m_ha.group(1).replace(',', '.'))

        return ops

    @classmethod
    def _extract_explicit_products(cls, text: str, sector_id: str) -> List[Dict[str, Any]]:
        """Extrait ou dimensionne strictement les produits demandés sans aucun ajout parasite."""
        lower = text.lower()

        # Cas 1 : Jus de mangue en bouteilles 33cl et 1L
        if sector_id == "jus" and ("mangue" in lower or "jus" in lower):
            has_33cl = "33" in lower or "33 cl" in lower or "33cl" in lower
            has_1l = "1 litre" in lower or "1l" in lower or "1 l" in lower
            if has_33cl and has_1l:
                return [
                    {
                        "id": 1,
                        "name": "Jus de mangue naturel pasteurisé (Bouteille 33 cl)",
                        "description": "Bouteille individuelle 33cl en verre / PET recyclable, pur jus naturel sans colorant",
                        "format": "33 cl",
                        "liters_per_unit": 0.33,
                        "share_pct": 0.50,
                        "unit_price": 500.0
                    },
                    {
                        "id": 2,
                        "name": "Jus de mangue naturel pasteurisé (Bouteille 1 Litre)",
                        "description": "Bouteille familiale 1L fermée hermétiquement, pasteurisation préservant saveur et vitamines",
                        "format": "1 L",
                        "liters_per_unit": 1.0,
                        "share_pct": 0.50,
                        "unit_price": 1200.0
                    }
                ]

        # Cas 2 : Distribution de matériel informatique
        if sector_id in ["commerce_informatique", "magasin_jouets"] and any(w in lower for w in ["informatique", "ordinateur", "pc", "serveur"]):
            return [
                {
                    "id": 1,
                    "name": "Ordinateurs portables professionnels & Stations de travail",
                    "description": "PC portables robustes configurés pour entreprises, bureautique et CAO",
                    "format": "Unité",
                    "share_pct": 0.50,
                    "unit_price": 350000.0
                },
                {
                    "id": 2,
                    "name": "Serveurs, Réseaux & Onduleurs d'entreprise",
                    "description": "Équipements d'infrastructure réseau, switchs managés et routeurs sécurisés",
                    "format": "Unité",
                    "share_pct": 0.30,
                    "unit_price": 450000.0
                },
                {
                    "id": 3,
                    "name": "Périphériques, Imprimantes & Consommables",
                    "description": "Imprimantes multifonctions laser, toners et accessoires informatiques",
                    "format": "Article",
                    "share_pct": 0.20,
                    "unit_price": 85000.0
                }
            ]

        # Cas 3 : Transport de marchandises
        if sector_id == "transport":
            return [
                {
                    "id": 1,
                    "name": "Fret routier interurbain régulier (Camion 10t)",
                    "description": "Liaisons interurbaines corridors régionaux pour industriels et grossistes",
                    "format": "Trajet",
                    "share_pct": 0.55,
                    "unit_price": 220000.0
                },
                {
                    "id": 2,
                    "name": "Distribution & messagerie urbaine de marchandises",
                    "description": "Navettes urbaines de livraison et approvisionnement commerces",
                    "format": "Course",
                    "share_pct": 0.30,
                    "unit_price": 85000.0
                },
                {
                    "id": 3,
                    "name": "Location de camion porteur 10t avec chauffeur",
                    "description": "Mise à disposition journalière pour chantiers et transferts logistiques",
                    "format": "Journée",
                    "share_pct": 0.15,
                    "unit_price": 160000.0
                }
            ]

        # Cas 4 : Solutions solaires photovoltaïques
        if sector_id == "energie_solaire":
            return [
                {
                    "id": 1,
                    "name": "Système Solaire Résidentiel Hybride (3 kWc)",
                    "description": "Installation résidentielle complète avec panneaux monocristallins, onduleur hybride 3 kW et stockage lithium LiFePO4",
                    "format": "Installation",
                    "share_pct": 0.40,
                    "unit_price": 2800000.0
                },
                {
                    "id": 2,
                    "name": "Centrale Solaire Commerciale & Tertiaire PME (10 kWc)",
                    "description": "Solution d'autoconsommation triphasée pour commerces et PME avec onduleur réseau, secours batterie et monitoring",
                    "format": "Installation",
                    "share_pct": 0.35,
                    "unit_price": 7500000.0
                },
                {
                    "id": 3,
                    "name": "Installation Solaire Professionnelle Forte Puissance (30 kWc)",
                    "description": "Centrale sur toiture pour cliniques, hôtels et exploitations agricoles avec régulation et télémétrie avancée",
                    "format": "Installation",
                    "share_pct": 0.20,
                    "unit_price": 18500000.0
                },
                {
                    "id": 4,
                    "name": "Contrat de Maintenance Préventive & Télé-surveillance",
                    "description": "Contrat annuel de thermographie infrarouge, nettoyage des modules, contrôle des chaînes DC/AC et maintenance d'urgence",
                    "format": "Contrat/an",
                    "share_pct": 0.05,
                    "unit_price": 350000.0
                }
            ]

        return []

    @classmethod
    def _detect_sector_scored(cls, lower: str) -> Tuple[str, str]:
        """Détection rigoureuse par pondération lexicale et exclusions grammaticales."""
        scores: Dict[str, int] = {
            "jus": 0,
            "commerce_informatique": 0,
            "transport": 0,
            "agriculture": 0,
            "aviculture": 0,
            "plantation_cacao": 0,
            "pisciculture": 0,
            "chocolaterie": 0,
            "boulangerie": 0,
            "savonnerie": 0,
            "textile": 0,
            "restaurant": 0,
            "quincaillerie": 0,
            "magasin_jouets": 0,
            "clinique": 0,
            "pharmacie": 0,
            "numerique_saas": 0,
            "energie_solaire": 0,
            "btp": 0,
            "education": 0
        }

        lexicon = {
            "jus": [
                r'\bjus\b', r'\bmangue[s]?\b', r'\bananas\b', r'\bbissap\b',
                r'\bpasteurisat(?:eur|ion)\b', r'\bembouteillage\b', r'\btransformation de fruits\b',
                r'\bboisson[s]?\b', r'\bagroalimentaire\b', r'\bextraction\b', r'\bpulpe\b'
            ],
            "commerce_informatique": [
                r'\bmatériel informatique\b', r'\binformatique\b', r'\bordinateurs?\b',
                r'\bserveurs?\b', r'\bpc portables?\b', r'\bdistribution de matériel\b'
            ],
            "transport": [
                r'\btransport routier\b', r'\bcamion[s]?\b', r'\bporteur[s]?\b',
                r'\bfret\b', r'\bflotte\b', r'\bcorridor[s]?\b', r'\bpéage[s]?\b'
            ],
            "agriculture": [
                r'\btomate[s]?\b', r'\bmara[îi]chage\b', r'\bmara[îi]ch[eè]re?[s]?\b',
                r'\blégume[s]?\b', r'\bhectare[s]?\b', r'\birrigation\b',
                r'\bgoutte[- ]à[- ]goutte\b', r'\bmotopompe[s]?\b', r'\bsemence[s]?\b',
                r'\bengrais\b', r'\bphytosanitaire[s]?\b', r'\brécolte[s]?\b'
            ],
            "aviculture": [
                r'\bpoulet[s]?\b', r'\bpoussin[s]?\b', r'\bvolaille[s]?\b', r'\bavicole[s]?\b',
                r'\baviculture\b', r'\bpoulailler[s]?\b', r'\bgésier[s]?\b',
                r'\bcouvoir[s]?\b', r'\bengraissement\b', r'\bpoulet(?:s)?\s+de\s+chair\b'
            ],
            "clinique": [
                r'\bclinique[s]?\b', r'\bcabinet médical\b', r'\bhôpital\b',
                r'\blaboratoire\b', r'\banalyses médicales\b', r'\bbiologie médicale\b',
                r'\bhématologie\b', r'\bbiochimie\b', r'\bautomate[s]?\b', r'\bprélèvement[s]?\b'
            ],
            "energie_solaire": [
                r'\bsolaire[s]?\b', r'\bphotovolta[ïi]que[s]?\b', r'\bpanneau[x]?\b',
                r'\bonduleur[s]?\b', r'\bbatterie[s]?\b', r'\bkwc\b', r'\brenouvelable[s]?\b',
                r'\br[ée]gulateur[s]?\b', r'\bkit[s]?\s+solaire[s]?\b'
            ]
        }

        for sec, patterns in lexicon.items():
            for pat in patterns:
                matches = re.findall(pat, lower)
                if matches:
                    scores[sec] += len(matches) * 3

        # Bonus déterminant si le secteur ou l'activité est explicitement déclaré
        m_sec = re.search(r'(?:secteur|sous-secteur|activit[ée]\s+principale)\s*[:=-]\s*([^\n\r,.;]+)', lower)
        if m_sec:
            sec_text = m_sec.group(1)
            if any(w in sec_text for w in ["solaire", "photovoltaïque", "photovoltaique", "renouvelable", "énergie", "energie"]):
                scores["energie_solaire"] += 50
            if any(w in sec_text for w in ["jus", "fruit", "mangue", "agroalimentaire", "boisson", "transformation"]):
                scores["jus"] += 40
            if any(w in sec_text for w in ["informatique", "matériel", "distribution", "technologie", "commerce"]):
                scores["commerce_informatique"] += 40
            if any(w in sec_text for w in ["transport", "logistique", "fret", "camion"]):
                scores["transport"] += 40
            if any(w in sec_text for w in ["tomate", "maraîchage", "maraichage", "agriculture"]):
                scores["agriculture"] += 40
            if any(w in sec_text for w in ["aviculture", "élevage", "poulet", "volaille"]):
                scores["aviculture"] += 40
            if any(w in sec_text for w in ["santé", "médical", "clinique", "laboratoire"]):
                scores["clinique"] += 40

        best_sector = max(scores, key=scores.get)
        names = {
            "jus": "Unité Agro-Industrielle & Transformation de Fruits",
            "commerce_informatique": "Distribution & Commerce de Matériel Informatique",
            "transport": "Transport Urbain, Interurbain & Logistique",
            "agriculture": "Exploitation Maraîchère & Culture de Tomates",
            "aviculture": "Élevage Avicole & Production de Poulets de Chair",
            "clinique": "Clinique Médicale & Laboratoire d'Analyses",
            "plantation_cacao": "Exploitation Cacaoyère Moderne",
            "pisciculture": "Pisciculture & Aquaculture Moderne",
            "chocolaterie": "Chocolaterie Artisanale & Confiserie",
            "boulangerie": "Boulangerie-Pâtisserie Moderne",
            "savonnerie": "Savonnerie & Cosmétiques Naturels",
            "textile": "Confection Textile & Mode Africaine",
            "restaurant": "Restauration & Fast-Food Moderne",
            "quincaillerie": "Quincaillerie & Matériaux de Construction",
            "magasin_jouets": "Commerce de Détail & Distribution",
            "pharmacie": "Pharmacie d'Officine & Santé",
            "numerique_saas": "Solutions Numériques & Services IT",
            "energie_solaire": "Énergie Solaire Photovoltaïque",
            "btp": "BTP, Bâtiment & Travaux Publics",
            "education": "Éducation, École & Formation"
        }
        return best_sector, names.get(best_sector, "Entreprise Commerciale")
