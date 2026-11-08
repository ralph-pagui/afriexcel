"""
Moteur d'intelligence artificielle Gemini Flash pour AfriExcel Backend.
Utilise les 7 clés API avec rotation automatique et cascade de modèles haute disponibilité.
"""

import os
import json
import requests
from requests.adapters import HTTPAdapter
from urllib3.util import Retry
from typing import Optional, Dict, Any
from .models import SectorProfile, ProductItem, EquipmentItem, StaffItem

# Rotation de clés Gemini via variables d'environnement (sécurisé pour production)
_raw_keys = os.environ.get("GEMINI_API_KEYS", "") or os.environ.get("GEMINI_API_KEY", "")
GEMINI_KEYS = [k.strip() for k in _raw_keys.split(",") if k.strip()]
if not GEMINI_KEYS:
    GEMINI_KEYS = [os.environ.get("GEMINI_API_KEY", "DEMO_GEMINI_API_KEY")]

MODELS_CASCADE = ["gemini-3.6-flash", "gemini-2.5-flash", "gemini-flash-latest"]

# Session HTTP partagée avec pool de connexions persistant
_session = requests.Session()
_adapter = HTTPAdapter(pool_connections=10, pool_maxsize=10, max_retries=Retry(total=2, backoff_factor=0.5))
_session.mount("https://", _adapter)

class GeminiSectorEngine:
    current_key_idx = 0

    @classmethod
    def generate_dynamic_sector(
        cls,
        user_prompt: str,
        country: str = "Côte d'Ivoire",
        currency: str = "FCFA",
        target_budget: Optional[float] = None,
        horizon_years: Optional[int] = 5,
        clean_topic: Optional[str] = None
    ) -> Optional[SectorProfile]:
        """
        Interroge Gemini 3.6 Flash (avec cascade automatique vers 2.5 Flash) pour concevoir un profil sectoriel
        100% sur-mesure et rigoureusement calibré sur le budget et l'activité demandée par l'utilisateur.
        """
        budget_clause = ""
        if target_budget and target_budget > 0:
            target_capex = target_budget * 0.60
            budget_clause = (
                f"\nCONTRAINTE DE BUDGET CRUCIALE :\n"
                f"- L'utilisateur a un budget global de {target_budget:,.0f} {currency}.\n"
                f"- La somme totale des investissements dans la liste 'equipment' DOIT ABSOLUMENT avoisiner {target_capex:,.0f} {currency} (environ 55% à 65% du budget total), "
                f"le solde étant réservé au BFR initial (stock marchandise de départ) et aux frais de constitution.\n"
                f"- Les charges de personnel et le volume d'activité journalier (base_daily_volume) doivent être proportionnés à ce budget, "
                f"pour générer un Chiffre d'Affaires annuel en Année 1 réaliste d'environ 2.2x à 3.2x l'investissement initial.\n"
            )

        system_instruction = f"""Tu es l'Analyste Financier Senior et Banquier d'Affaires OHADA de classe mondiale, expert en modélisation d'entreprise en Afrique Francophone ({country}, zone UEMOA/CEMAC en {currency}).
Analyse rigoureusement la demande de l'utilisateur pour concevoir un modèle métier complet, ultra-professionnel, chirurgicalement adapté au secteur demandé et respectant toutes les contraintes de budget et de durée.{budget_clause}

GÉNÈRE STRICTEMENT UN OBJET JSON CONFORME À CETTE STRUCTURE :
{{
  "sector_id": "identifiant_court_du_metier",
  "name": "Nom valorisant du projet",
  "activity": "Description précise de l'activité commerciale, médicale ou industrielle",
  "unit": "Unité de vente principale",
  "base_daily_volume": 30,
  "working_days_per_year": 312,
  "raw_materials_desc": "Fournisseurs clés, approvisionnements réels et grossistes agréés",
  "products": [
    {{"id": 1, "name": "Gamme Principale", "description": "Description détaillée", "share_pct": 0.45, "unit_price": 5000}},
    {{"id": 2, "name": "Gamme Secondaire", "description": "Description détaillée", "share_pct": 0.35, "unit_price": 7500}},
    {{"id": 3, "name": "Gamme Complémentaire", "description": "Description détaillée", "share_pct": 0.20, "unit_price": 12000}}
  ],
  "equipment": [
    {{"name": "Équipement réel indispensable n°1", "cost": 3000000, "duration_years": 8, "depreciation_rate": 0.125}},
    {{"name": "Équipement réel indispensable n°2", "cost": 1500000, "duration_years": 5, "depreciation_rate": 0.20}},
    {{"name": "Équipement réel indispensable n°3", "cost": 1000000, "duration_years": 4, "depreciation_rate": 0.25}},
    {{"name": "Équipement réel indispensable n°4", "cost": 800000, "duration_years": 5, "depreciation_rate": 0.20}}
  ],
  "staff": [
    {{"title": "Gérant / Responsable qualifié", "count": 1, "monthly_net": 250000, "monthly_cnps": 40000}},
    {{"title": "Personnel commercial / caisse", "count": 1, "monthly_net": 120000, "monthly_cnps": 20000}}
  ]
}}
RÈGLES STRICTES :
1. Les équipements doivent correspondre EXACTEMENT et UNIQUEMENT au métier demandé (ex: pour un magasin de jouets : rayonnages colorés, bacs et gondoles, caisse tactile et logiciel stock, climatiseur boutique, enseigne lumineuse enfantine, kit vidéosurveillance).
2. Les montants doivent être réalistes pour l'Afrique de l'Ouest/Centrale en {currency}.
3. Le total des share_pct des produits doit être rigoureusement égal à 1.0 (ex: 0.45 + 0.35 + 0.20 = 1.0)."""

        num_keys = len(GEMINI_KEYS)

        for attempt in range(num_keys):
            key_idx = (cls.current_key_idx + attempt) % num_keys
            api_key = GEMINI_KEYS[key_idx]

            for model in MODELS_CASCADE:
                url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
                body = {
                    "contents": [{"parts": [{"text": f"Génère le modèle sectoriel bancable OHADA sur mesure pour : '{user_prompt}' à implanter en {country}."}]}],
                    "systemInstruction": {"parts": [{"text": system_instruction}]},
                    "generationConfig": {
                        "responseMimeType": "application/json",
                        "temperature": 0.15
                    }
                }

                try:
                    r = _session.post(url, json=body, timeout=16)
                    if r.status_code == 200:
                        data = r.json()
                        text = data.get("candidates", [{}])[0].get("content", {}).get("parts", [{}])[0].get("text", "")
                        if text:
                            parsed = json.loads(text)
                            profile = cls._dict_to_profile(parsed)
                            if profile:
                                cls.current_key_idx = (key_idx + 1) % num_keys
                                print(f"[Gemini Flash] Modélisation sectorielle générée avec succès via {model} (Clé #{key_idx + 1})")
                                return profile
                    else:
                        print(f"[Gemini Flash] Modèle {model} (Clé #{key_idx + 1}) a répondu avec le statut {r.status_code}, basculement vers le modèle suivant...")
                        continue
                except Exception as e:
                    print(f"[Gemini Flash] Modèle {model} (Clé #{key_idx + 1}) timeout/erreur ({type(e).__name__}), basculement...")
                    continue

        print("[Gemini Flash] Toutes les tentatives d'IA dynamique ont échoué, basculement vers le catalogue de secours.")
        return None

    @classmethod
    def _dict_to_profile(cls, data: Dict[str, Any]) -> Optional[SectorProfile]:
        try:
            # Gestion des cas où Gemini encapsule le JSON dans une clé racine
            for sub_key in ["project_info", "data", "business_plan", "model", "sector_profile"]:
                if sub_key in data and isinstance(data[sub_key], dict):
                    for k, v in data[sub_key].items():
                        if k not in data:
                            data[k] = v
            products = [
                ProductItem(
                    id=p.get("id", i + 1),
                    name=p.get("name", f"Gamme {i + 1}"),
                    description=p.get("description", ""),
                    share_pct=float(p.get("share_pct", 0.33)),
                    unit_price=float(p.get("unit_price", 2500))
                )
                for i, p in enumerate(data.get("products", []))
            ]

            equipment = [
                EquipmentItem(
                    name=eq.get("name", "Équipement d'exploitation"),
                    cost=float(eq.get("cost", 3000000)),
                    duration_years=int(eq.get("duration_years", 5)),
                    depreciation_rate=float(eq.get("depreciation_rate", 0.20))
                )
                for eq in data.get("equipment", [])
            ]

            staff = [
                StaffItem(
                    title=st.get("title", "Salarié qualifié"),
                    count=int(st.get("count", 1)),
                    monthly_net=float(st.get("monthly_net", 150000)),
                    monthly_cnps=float(st.get("monthly_cnps", 25000))
                )
                for st in data.get("staff", [])
            ]

            return SectorProfile(
                sector_id=data.get("sector_id", "commerce"),
                name=data.get("name", "Entreprise Commerciale / Industrielle"),
                activity=data.get("activity", "Activité d'exploitation commerciale"),
                unit=data.get("unit", "Unité"),
                base_daily_volume=int(data.get("base_daily_volume", 150)),
                working_days_per_year=int(data.get("working_days_per_year", 360)),
                raw_materials_desc=data.get("raw_materials_desc", "Approvisionnements généraux"),
                products=products,
                equipment=equipment,
                staff=staff
            )
        except Exception as e:
            print("Erreur de conversion JSON vers SectorProfile:", e)
            return None
