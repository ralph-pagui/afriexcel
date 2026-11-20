"""
Système de Verrouillage des Paramètres Utilisateur & Contexte de Projet Isolé
AfriExcel Copilot - Moteur d'Intégrité Bancaire OHADA

Garantit que chaque business plan possède un identifiant unique, un contexte
entièrement cloisonné et des paramètres explicitement verrouillés qui ne peuvent
être ni écrasés par d'anciens dossiers, ni déformés par des heuristiques ou des modèles IA.
"""

from typing import Dict, Any, List, Optional
from enum import Enum
import uuid
from dataclasses import dataclass, field

class ParameterOrigin(str, Enum):
    EXPLICIT_PROMPT = "EXPLICIT_PROMPT"      # Fourni formellement par l'utilisateur
    DERIVED_BUSINESS = "DERIVED_BUSINESS"    # Dérivé des équations métiers du secteur
    DEFAULT_OHADA = "DEFAULT_OHADA"          # Hypothèse normative par défaut UEMOA/CEMAC

class LockStatus(str, Enum):
    LOCKED = "LOCKED"        # Intangible, aucune modification autorisée
    UNLOCKED = "UNLOCKED"    # Hypothèse indicative modifiable

@dataclass
class LockedParameter:
    key: str
    value: Any
    unit: str
    origin: ParameterOrigin
    status: LockStatus
    description: str

@dataclass
class ProjectContext:
    """Source de vérité unique et inviolable pour une génération de Business Plan."""
    session_id: str = field(default_factory=lambda: str(uuid.uuid4()))
    project_name: str = "PROJET D'ENTREPRISE SARL"
    country: str = "Bénin"
    city: str = "Cotonou"
    currency: str = "FCFA"
    sector_id: str = "jus"
    sector_name: str = "Industrie Agroalimentaire"
    sub_sector: str = "Transformation de fruits"
    activity_desc: str = "Transformation de mangues fraîches en jus naturels pasteurisés"
    
    # Produits explicitement demandés (liste stricte, zéro ajout automatique)
    explicit_products: List[str] = field(default_factory=list)
    
    # Capacités opérationnelles verrouillées
    daily_capacity: Optional[float] = None
    daily_capacity_unit: str = "Litres"
    working_days_per_month: int = 22
    working_days_per_year: int = 264
    
    # Paramètres financiers verrouillés
    horizon_years: int = 5
    equity_ratio: float = 0.30
    loan_ratio: float = 0.70
    loan_interest_rate: float = 0.11
    loan_duration_years: int = 5
    target_budget: Optional[float] = None
    
    # Registre centralisé des verrous
    locks: Dict[str, LockedParameter] = field(default_factory=dict)

    def lock_param(self, key: str, value: Any, unit: str, description: str, origin: ParameterOrigin = ParameterOrigin.EXPLICIT_PROMPT):
        self.locks[key] = LockedParameter(
            key=key,
            value=value,
            unit=unit,
            origin=origin,
            status=LockStatus.LOCKED if origin == ParameterOrigin.EXPLICIT_PROMPT else LockStatus.UNLOCKED,
            description=description
        )

    def is_locked(self, key: str) -> bool:
        return key in self.locks and self.locks[key].status == LockStatus.LOCKED

    def get_locked_value(self, key: str, default: Any = None) -> Any:
        if key in self.locks:
            return self.locks[key].value
        return default

    def validate_downstream(self, computed_data: Dict[str, Any]) -> List[str]:
        """Vérifie que les calculs et structures aval n'ont violé aucun paramètre verrouillé."""
        violations = []
        kpis = computed_data.get("kpis")
        if not kpis:
            return ["KPIs manquants dans les résultats calculés."]

        # 1. Vérification Quotité Fonds Propres
        if self.is_locked("equity_ratio"):
            expected_eq = round(self.locks["equity_ratio"].value * 100, 1)
            actual_eq = round(kpis.equity_pct, 1)
            if abs(expected_eq - actual_eq) > 0.5:
                violations.append(f"Violation apport personnel : demandé {expected_eq}%, calculé {actual_eq}%")

        # 2. Vérification Quotité Emprunt Bancaire
        if self.is_locked("loan_ratio"):
            expected_loan = round(self.locks["loan_ratio"].value * 100, 1)
            actual_loan = round(kpis.loan_pct, 1)
            if abs(expected_loan - actual_loan) > 0.5:
                violations.append(f"Violation emprunt bancaire : demandé {expected_loan}%, calculé {actual_loan}%")

        # 3. Vérification Taux d'intérêt de l'emprunt
        if self.is_locked("loan_interest_rate"):
            expected_rate = round(self.locks["loan_interest_rate"].value * 100, 1)
            actual_rate = round(kpis.loan_interest_rate, 1)
            if abs(expected_rate - actual_rate) > 0.1:
                violations.append(f"Violation taux d'intérêt : demandé {expected_rate}%, calculé {actual_rate}%")

        # 4. Vérification Durée de modélisation
        actual_horizon = computed_data.get("horizon_years", 5)
        if actual_horizon != self.horizon_years:
            violations.append(f"Violation horizon : demandé {self.horizon_years} ans, calculé {actual_horizon} ans")

        # 5. Vérification Équilibre strict Emplois = Ressources
        total_inv = kpis.total_investment
        total_ressources = kpis.equity_promoter + kpis.bank_loan
        if abs(total_inv - total_ressources) > 1.0:
            violations.append(f"Déséquilibre Plan de Financement : Emplois ({total_inv:,.0f}) != Ressources ({total_ressources:,.0f})")

        return violations
