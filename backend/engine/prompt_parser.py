"""
Module d'analyse sémantique et extraction NLP avancée pour AfriExcel Copilot.
Extrait fidèlement les paramètres économiques, financiers et sectoriels des prompts utilisateurs
sans jamais injecter de texte de commande dans les cellules du classeur.
"""

import re
from typing import Dict, Any, Optional, List

class PromptParser:
    """Analyseur sémantique déterministe pour requêtes de Business Plan en Afrique francophone."""

    @classmethod
    def parse(cls, prompt: str, default_country: str = "Côte d'Ivoire", default_currency: str = "FCFA") -> Dict[str, Any]:
        """Extrait l'ensemble des informations structurées du prompt utilisateur."""
        text = prompt.strip()
        lower = text.lower()

        # 1. Extraction de la Raison Sociale / Nom du projet
        company_name = cls._extract_company_name(text)

        # 2. Extraction du Pays et de la Localisation
        country = cls._extract_country(text, default_country)
        location = cls._extract_location(text, country)

        # 3. Extraction de la Devise
        currency = cls._extract_currency(text, default_currency)

        # 4. Extraction des paramètres de financement (Apport, Emprunt, Taux, Durée)
        equity_ratio, loan_ratio = cls._extract_financing_split(text)
        loan_interest_rate = cls._extract_interest_rate(text)
        loan_duration_years = cls._extract_loan_duration(text)

        # 5. Extraction de l'horizon prévisionnel
        horizon_years = cls._extract_horizon_years(text)

        # 6. Extraction du budget ou investissement
        total_budget = cls._extract_budget(text)

        # 7. Détection chirurgicale du secteur d'activité par scoring multi-mots
        sector_id, sector_name = cls._detect_sector_scored(lower)

        # 8. Extraction de la capacité et données opérationnelles
        capacity_info = cls._extract_capacity(text)

        return {
            "company_name": company_name,
            "country": country,
            "location": location,
            "currency": currency,
            "sector_id": sector_id,
            "sector_name": sector_name,
            "equity_ratio": equity_ratio,
            "loan_ratio": loan_ratio,
            "loan_interest_rate": loan_interest_rate,
            "loan_duration_years": loan_duration_years,
            "horizon_years": horizon_years,
            "total_budget": total_budget,
            "capacity_info": capacity_info,
            "raw_prompt": text
        }

    @classmethod
    def _extract_company_name(cls, text: str) -> str:
        """Extrait un nom d'entreprise ou de projet court et valorisant."""
        # Recherche explicite "Nom du projet : XYZ" ou "Raison sociale : XYZ"
        m = re.search(r'(?:nom\s+du\s+projet|raison\s+sociale|nom\s+de\s+l[\'’]entreprise|entreprise|projet)\s*[:=-]\s*([^\n\r,.;]+)', text, re.I)
        if m:
            cand = m.group(1).strip()
            # Nettoyer les balises Markdown éventuelles
            cand = re.sub(r'[*_#]', '', cand).strip()
            if 3 <= len(cand) <= 60 and not re.match(r'^(?:génère|crée|fais|donne|un|le|mon)\b', cand, re.I):
                return cand

        # Recherche de mentions types SARL / SAS / SUARL / GIE / Ferme / Clinique
        m_legal = re.search(r'\b([A-ZÀ-ÿ0-9\s&\'\-]+(?:\s+(?:SARL|SAS|SUARL|SA|GIE|GROUP|SERVICES|PLUS|AGRO)))\b', text)
        if m_legal:
            cand = m_legal.group(1).strip()
            if 3 <= len(cand) <= 45:
                return cand

        # Recherche d'expression "création d'une ferme avicole moderne"
        m_proj = re.search(r'(?:création|mise\s+en\s+place|lancement|développement)\s+(?:d[\'’]une?|de\s+la)\s+([^\n\r,.;]+)', text, re.I)
        if m_proj:
            cand = m_proj.group(1).strip()
            cand = re.sub(r'[*_#]', '', cand).strip()
            cand = re.sub(r'\s+(?:spécialisée|au|en|dans|à|sur|avec|pour).*$', '', cand, flags=re.I).strip()
            if 4 <= len(cand) <= 50:
                words = [w.capitalize() for w in cand.split()]
                return " ".join(words)

        return "AVICOLA PLUS SARL"

    @classmethod
    def _extract_country(cls, text: str, default_country: str) -> str:
        """Détecte avec précision le pays d'implantation africain."""
        countries = [
            ("Cameroun", ["cameroun", "cameroon", "douala", "yaoundé", "yaounde", "bafoussam", "garoua", "littoral"]),
            ("Côte d'Ivoire", ["côte d'ivoire", "cote d'ivoire", "abidjan", "yamoussoukro", "san pedro", "bouaké", "bouake", "korhogo"]),
            ("Sénégal", ["sénégal", "senegal", "dakar", "thies", "thiès", "saint-louis", "touba", "ziguinchor"]),
            ("Bénin", ["bénin", "benin", "cotonou", "porto-novo", "parakou", "calavi"]),
            ("Togo", ["togo", "lomé", "lome", "kara", "kpalimé"]),
            ("Burkina Faso", ["burkina", "ouagadougou", "bobo-dioulasso"]),
            ("Mali", ["mali", "bamako", "sikasso", "ségou"]),
            ("Gabon", ["gabon", "libreville", "port-gentil", "franceville"]),
            ("Congo", ["congo", "brazzaville", "pointe-noire"]),
            ("RDC", ["rdc", "kinshasa", "lubumbashi", "goma"]),
            ("Guinée", ["guinée", "guinee", "conakry", "kankan"])
        ]
        lower = text.lower()
        for c_name, keywords in countries:
            for kw in keywords:
                if re.search(r'\b' + re.escape(kw) + r'\b', lower):
                    return c_name
        return default_country

    @classmethod
    def _extract_location(cls, text: str, country: str) -> str:
        """Extrait la localisation précise (ville, région)."""
        m = re.search(r'(?:localisation|ville|siège|région)\s*[:=-]\s*([^\n\r,.;]+)', text, re.I)
        if m:
            loc = m.group(1).strip()
            loc = re.sub(r'[*_#]', '', loc).strip()
            if 3 <= len(loc) <= 50:
                return loc

        # Recherche de motifs "à Douala" ou "dans la région du Littoral"
        locations = []
        m_ville = re.search(r'\b(?:à|a)\s+([A-ZÀ-ÿ][a-zà-ÿ]+(?:\s+[A-ZÀ-ÿ][a-zà-ÿ]+)?)\b', text)
        if m_ville:
            v = m_ville.group(1).strip()
            if v.lower() not in ["excel", "ce", "cette", "un", "une"]:
                locations.append(v)

        m_reg = re.search(r'\b(?:région|zone)\s+(?:du|de|de\s+la)\s+([A-ZÀ-ÿ][a-zà-ÿ]+)\b', text, re.I)
        if m_reg:
            locations.append(m_reg.group(1).capitalize())

        if locations:
            return f"{', '.join(locations)}, {country}"
        return f"{country} (Zone OHADA)"

    @classmethod
    def _extract_currency(cls, text: str, default_currency: str) -> str:
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
    def _extract_financing_split(cls, text: str) -> (float, float):
        """Extrait les pourcentages d'apport personnel et d'emprunt bancaire."""
        eq_ratio = None
        loan_ratio = None

        # Ordre 1: "apport personnel : 30%" ou "fonds propres 25%"
        m_eq1 = re.search(r'(?:apport\s+(?:personnel|promoteur|fonds\s+propres?)|fonds\s+propres?)\s*[:=-]?\s*(?:de\s*)?(\d+(?:[.,]\d+)?)\s*%', text, re.I)
        # Ordre 2: "30% d'apport" ou "25 % fonds propres"
        m_eq2 = re.search(r'(\d+(?:[.,]\d+)?)\s*%\s*(?:d[\'’]apport|fonds\s+propres?|apport\s+personnel)', text, re.I)
        m_eq = m_eq1 or m_eq2
        if m_eq:
            eq_val = float(m_eq.group(1).replace(',', '.'))
            eq_ratio = eq_val / 100.0 if eq_val > 1 else eq_val

        # Ordre 1: "emprunt bancaire : 70%" ou "crédit 75%"
        m_loan1 = re.search(r'(?:emprunt\s+(?:bancaire)?|crédit\s+(?:bancaire)?|prêt\s+(?:bancaire)?|financement\s+bancaire)\s*[:=-]?\s*(?:de\s*)?(\d+(?:[.,]\d+)?)\s*%', text, re.I)
        # Ordre 2: "70% d'emprunt" ou "75 % crédit bancaire"
        m_loan2 = re.search(r'(\d+(?:[.,]\d+)?)\s*%\s*(?:d[\'’]emprunt|d[\'’]endettement|crédit\s+bancaire|emprunt\s+bancaire|dette\s+bancaire)', text, re.I)
        m_loan = m_loan1 or m_loan2
        if m_loan:
            loan_val = float(m_loan.group(1).replace(',', '.'))
            loan_ratio = loan_val / 100.0 if loan_val > 1 else loan_val

        # Cohérence et complétion
        if eq_ratio is not None and loan_ratio is None:
            loan_ratio = max(0.0, 1.0 - eq_ratio)
        elif loan_ratio is not None and eq_ratio is None:
            eq_ratio = max(0.0, 1.0 - loan_ratio)
        elif eq_ratio is None and loan_ratio is None:
            eq_ratio = 0.30
            loan_ratio = 0.70

        return round(eq_ratio, 2), round(loan_ratio, 2)

    @classmethod
    def _extract_interest_rate(cls, text: str) -> float:
        """Extrait le taux d'intérêt annuel du crédit bancaire (ex: 12% ou 10%)."""
        # Forme 1: "taux d'intérêt annuel de 12%" ou "taux de 10%"
        m1 = re.search(r'taux\s+(?:d[\'’]int[ée]r[êe]t)?\s*(?:annuel)?\s*(?:de\s*|d[\'’]\s*|à\s*)?(\d+(?:[.,]\d+)?)\s*%', text, re.I)
        # Forme 2: "crédit bancaire à 10 %" ou "emprunt à 8%"
        m2 = re.search(r'(?:crédit|emprunt|prêt)\s+(?:bancaire\s+)?à\s*(\d+(?:[.,]\d+)?)\s*%', text, re.I)
        m = m1 or m2
        if m:
            val = float(m.group(1).replace(',', '.'))
            return val / 100.0 if val > 1 else val
        return 0.12  # Défaut bancaire standard PME zone CFA : 12%

    @classmethod
    def _extract_loan_duration(cls, text: str) -> int:
        """Extrait la durée du crédit bancaire en années."""
        m1 = re.search(r'(?:remboursement|échéancier|prêt|crédit|emprunt)\s*(?:sur|de)?\s*(\d+)\s*ans?\b', text, re.I)
        m2 = re.search(r'(?:à\s+\d+\s*%\s+)?sur\s+(\d+)\s*ans?\b', text, re.I)
        m = m1 or m2
        if m:
            d = int(m.group(1))
            if 1 <= d <= 15:
                return d
        return 5

    @classmethod
    def _extract_horizon_years(cls, text: str) -> int:
        """Extrait la durée du business plan (horizon temporel)."""
        m = re.search(r'(?:horizon|durée|prévisions|modèle|sur)\s*(?:prévisionnel(?:le)?\s+)?(?:de\s*)?(\d+)\s*ans?\b', text, re.I)
        if m:
            y = int(m.group(1))
            if 2 <= y <= 15:
                return y
        if "cinq ans" in text.lower():
            return 5
        if "trois ans" in text.lower():
            return 3
        return 5

    @classmethod
    def _extract_budget(cls, text: str) -> Optional[float]:
        """Extrait le montant global de l'investissement si précisé."""
        m_millions = re.search(r'(?:budget|investissement|montant|besoin)\s*(?:global|total)?\s*(?:de\s*)?(\d+(?:[.,]\d+)?)\s*(?:millions?|m)\b', text, re.I)
        if m_millions:
            return float(m_millions.group(1).replace(',', '.')) * 1_000_000

        m_num = re.search(r'(?:budget|investissement|montant|besoin)\s*(?:global|total)?\s*(?:de\s*)?(\d[\d\s]{4,})\s*(?:fcfa|cfa|f)?\b', text, re.I)
        if m_num:
            clean_digits = re.sub(r'\s+', '', m_num.group(1))
            if clean_digits.isdigit() and int(clean_digits) >= 500_000:
                return float(clean_digits)
        return None

    @classmethod
    def _extract_capacity(cls, text: str) -> Optional[str]:
        """Extrait les informations opérationnelles clés (ex: 5000 poussins, 6 cycles)."""
        m = re.search(r'(\d[\d\s]*\s*(?:poussins?|sujets?|poulets?|hectares?|tonnes?|têtes?|litres?)\s*(?:par\s+cycle|par\s+an|par\s+jour)?)', text, re.I)
        if m:
            return m.group(1).strip()
        return None

    @classmethod
    def _detect_sector_scored(cls, lower: str) -> (str, str):
        """
        Détection rigoureuse par pondération lexicale et exclusions grammaticales.
        Évite TOUTE collision de sous-chaîne (comme 'soins' dans 'besoins').
        """
        scores: Dict[str, int] = {
            "aviculture": 0,
            "plantation_cacao": 0,
            "pisciculture": 0,
            "chocolaterie": 0,
            "jus": 0,
            "boulangerie": 0,
            "savonnerie": 0,
            "textile": 0,
            "restaurant": 0,
            "quincaillerie": 0,
            "magasin_jouets": 0,
            "transport": 0,
            "clinique": 0,
            "pharmacie": 0,
            "numerique_saas": 0,
            "energie_solaire": 0,
            "btp": 0,
            "education": 0
        }

        # Dictionnaires de mots-clés stricts avec détection mot entier
        lexicon = {
            "aviculture": [
                r'\bpoulet[s]?\b', r'\bpoussin[s]?\b', r'\bvolaille[s]?\b', r'\bavicole[s]?\b',
                r'\baviculture\b', r'\bpoulailler[s]?\b', r'\bgésier[s]?\b', r'\bélevage\b',
                r'\bcouvoir[s]?\b', r'\bengraissement\b', r'\bchair\b', r'\bpondeuse[s]?\b'
            ],
            "plantation_cacao": [
                r'\bcacao\b', r'\bcacaoy(?:er|ère|ere|ers)\b', r'\bfève[s]?\b', r'\bfeve[s]?\b',
                r'\bcabosse[s]?\b', r'\bcabossage\b', r'\bplantation\b', r'\birad\b', r'\bminader\b'
            ],
            "pisciculture": [
                r'\bpoisson[s]?\b', r'\bpisciculture\b', r'\baquaculture\b', r'\btilapia[s]?\b',
                r'\bsilure[s]?\b', r'\balevin[s]?\b', r'\bétang[s]?\b', r'\betang[s]?\b'
            ],
            "chocolaterie": [
                r'\bchocolat[s]?\b', r'\bchocolaterie\b', r'\bconfiserie\b', r'\btorréfacteur\b',
                r'\bconcheuse\b', r'\bbean[- ]to[- ]bar\b'
            ],
            "jus": [
                r'\bjus\b', r'\bananas\b', r'\bmangue[s]?\b', r'\bbissap\b', r'\bpasteurisat(?:eur|ion)\b',
                r'\bembouteillage\b', r'\bboisson[s]?\b'
            ],
            "boulangerie": [
                r'\bpain[s]?\b', r'\bboulangerie\b', r'\bpâtisserie\b', r'\bpatisserie\b',
                r'\bfournil\b', r'\bpétrin\b', r'\bbaguette[s]?\b', r'\bviennoiserie[s]?\b'
            ],
            "savonnerie": [
                r'\bsavon[s]?\b', r'\bsavonnerie\b', r'\bkarité\b', r'\bkarite\b',
                r'\bcosmétique[s]?\b', r'\bcosmetique[s]?\b', r'\bsaponifi(?:cation|é)\b'
            ],
            "textile": [
                r'\btextile[s]?\b', r'\bcouture\b', r'\bvêtement[s]?\b', r'\bvetement[s]?\b',
                r'\bwax\b', r'\bdan fani\b', r'\bconfection\b', r'\bmode\b'
            ],
            "restaurant": [
                r'\brestaurant[s]?\b', r'\brestauration\b', r'\bfast[- ]food\b', r'\btraiteur\b',
                r'\bmaquis\b', r'\bchoukouya\b', r'\bgrillade[s]?\b'
            ],
            "quincaillerie": [
                r'\bquincaillerie\b', r'\bciment\b', r'\bfer à béton\b', r'\bmatériaux de construction\b'
            ],
            "magasin_jouets": [
                r'\bjouet[s]?\b', r'\bludique[s]?\b', r'\bpeluche[s]?\b', r'\bpoupée[s]?\b', r'\bjeux éducatifs\b'
            ],
            "transport": [
                r'\btransport\b', r'\blogistique\b', r'\bflotte\b', r'\bcamion[s]?\b', r'\blivraison frigorifique\b'
            ],
            "clinique": [
                r'\bclinique[s]?\b', r'\bcabinet médical\b', r'\bhôpital\b', r'\bhopital\b',
                r'\bdocteur[s]?\b', r'\bmédecin[s]?\b', r'\bdispensaire\b', r'\bhospitalisation[s]?\b',
                r'\bconsultation(?:s)? médicale(?:s)?\b',
                # MOT 'soins' SEULEMENT S'IL N'EST PAS PRÉCÉDÉ PAR 'be' (évite 'besoins')
                r'(?<!be)soin[s]?\s+(?:médicaux|de santé|hospitaliers|infirmiers)\b'
            ],
            "pharmacie": [
                r'\bpharmacie\b', r'\bofficine\b', r'\bmédicament[s]?\b', r'\bmedicament[s]?\b', r'\bparapharmacie\b'
            ],
            "numerique_saas": [
                r'\bsaas\b', r'\bapplication web\b', r'\blogiciel\b', r'\bstartup tech\b', r'\bplateforme digitale\b'
            ],
            "energie_solaire": [
                r'\bsolaire\b', r'\bphotovoltaïque\b', r'\bpanneaux solaires\b', r'\bénergie renouvelable\b'
            ],
            "btp": [
                r'\bbtp\b', r'\bgénie civil\b', r'\bconstruction de bâtiments\b', r'\bterrassement\b'
            ],
            "education": [
                r'\bécole privée\b', r'\becole\b', r'\bcentre de formation\b', r'\bcollège\b', r'\blycée\b'
            ]
        }

        for sec, patterns in lexicon.items():
            for pat in patterns:
                matches = re.findall(pat, lower)
                if matches:
                    scores[sec] += len(matches) * 3

        # Bonus si le secteur est explicitement cité dans une section "Secteur : ..."
        m_sec = re.search(r'secteur\s*[:=-]\s*([^\n\r,.;]+)', lower)
        if m_sec:
            sec_text = m_sec.group(1)
            if any(w in sec_text for w in ["aviculture", "élevage", "elevage", "agriculture", "poulet", "chair"]):
                scores["aviculture"] += 15
            if any(w in sec_text for w in ["santé", "médical", "clinique", "soins"]):
                scores["clinique"] += 15
            if any(w in sec_text for w in ["commerce", "distribution", "jouet"]):
                scores["magasin_jouets"] += 10

        best_sector = max(scores, key=scores.get)
        if scores[best_sector] > 0:
            names = {
                "aviculture": "Élevage Avicole & Production de Poulets de Chair",
                "plantation_cacao": "Exploitation Cacaoyère Moderne",
                "pisciculture": "Pisciculture & Aquaculture Moderne",
                "chocolaterie": "Chocolaterie Artisanale & Confiserie",
                "jus": "Production & Mise en Bouteille de Jus de Fruits",
                "boulangerie": "Boulangerie-Pâtisserie Moderne",
                "savonnerie": "Savonnerie & Cosmétiques Naturels",
                "textile": "Confection Textile & Mode Africaine",
                "restaurant": "Restauration & Fast-Food Moderne",
                "quincaillerie": "Quincaillerie & Matériaux de Construction",
                "magasin_jouets": "Commerce de Détail & Distribution",
                "transport": "Transport Urbain & Logistique",
                "clinique": "Clinique Médicale & Centre de Soins",
                "pharmacie": "Pharmacie d'Officine & Santé",
                "numerique_saas": "Solutions Numériques & Services IT",
                "energie_solaire": "Énergie Solaire Photovoltaïque",
                "btp": "BTP, Bâtiment & Travaux Publics",
                "education": "Éducation, École & Formation"
            }
            return best_sector, names.get(best_sector, "Entreprise Commerciale")

        return "aviculture", "Ferme Avicole Moderne"
