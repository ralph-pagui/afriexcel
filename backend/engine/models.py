from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any

class PlanRequest(BaseModel):
    prompt: str = Field(..., description="Description du projet par l'utilisateur (ex: 'business plan pour magasin de jouet sur 8 ans, budget 10 millions')")
    country: Optional[str] = Field("Côte d'Ivoire", description="Pays d'implantation (ex: Côte d'Ivoire, Sénégal, Cameroun)")
    currency: Optional[str] = Field("FCFA", description="Devise monétaire (FCFA, GNF, USD, EUR)")
    total_budget: Optional[float] = Field(None, description="Budget cible souhaité en devise locale")
    horizon_years: Optional[int] = Field(None, description="Durée de modélisation en années (ex: 5, 7, 8, 10 ans)")
    promoter_name: Optional[str] = Field(None, description="Nom du promoteur / fondateur")
    company_name: Optional[str] = Field(None, description="Raison sociale de l'entreprise")
    template_id: Optional[str] = Field("auto", description="Identifiant du template de design (ex: 'emerald_africa', 'sahara_commerce', 'royal_navy', 'cyber_indigo')")
    is_modification: Optional[bool] = Field(False, description="Indique s'il s'agit d'une mise à jour chirurgicale d'un plan existant")
    previous_topic: Optional[str] = Field(None, description="Sujet ou secteur du projet précédent lors d'une modification")

class ProductItem(BaseModel):
    id: int
    name: str
    description: str
    share_pct: float  # ex: 0.45
    unit_price: float

class EquipmentItem(BaseModel):
    name: str
    cost: float
    duration_years: int
    depreciation_rate: float

class StaffItem(BaseModel):
    title: str
    count: int
    monthly_net: float
    monthly_cnps: float

class SectorProfile(BaseModel):
    sector_id: str
    name: str
    activity: str
    unit: str
    products: List[ProductItem]
    equipment: List[EquipmentItem]
    staff: List[StaffItem]
    raw_materials_desc: str
    base_daily_volume: int = 300
    working_days_per_year: int = 280

class KPIOverview(BaseModel):
    total_investment: float
    equity_promoter: float
    bank_loan: float
    revenue_year1: float
    ebitda_year1: float  # EBE
    net_profit_year1: float
    breakeven_revenue: float
    breakeven_days: int
    irr_pct: str  # TRI
    npv: float  # VAN à 10%
    payback_period: str
    jobs_created: int

class PlanResponse(BaseModel):
    success: bool
    project_name: str
    sector: str
    country: str
    currency: str
    summary: str
    kpis: KPIOverview
    sheets_count: int
    filename: str
    file_base64: str
    template_id: Optional[str] = None
    template_name: Optional[str] = None
