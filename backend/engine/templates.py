"""
Catalogue des Templates de Design Professionnels pour Business Plans OHADA AfriExcel
Chaque template définit une identité visuelle complète (palettes, typographies, bannières,
cartes graphiques et charts natifs Excel) inspirée des standards des comités de crédit bancaire.
"""

from typing import Dict, Any, List
from dataclasses import dataclass

@dataclass
class ThemeConfig:
    id: str
    name: str
    subtitle: str
    description: str
    icon: str
    tagline: str
    vision_quote: str
    banner_bg: str       # Hex sans #
    banner_text: str
    header_bg: str
    header_text: str
    accent_color: str
    highlight_bg: str
    highlight_text: str
    zebra_bg: str
    border_color: str
    total_bg: str
    total_text: str
    chart_primary: str
    card_bg: str = "F8FAFC"
    card_border: str = "CBD5E1"
    badge_success_bg: str = "DCFCE7"
    badge_success_fg: str = "14532D"
    badge_warning_bg: str = "FEF3C7"
    badge_warning_fg: str = "78350F"
    chart_secondary: str = "0284C7"
    chart_accent: str = "10B981"

TEMPLATE_CATALOG: Dict[str, ThemeConfig] = {
    "emerald_africa": ThemeConfig(
        id="emerald_africa",
        name="Émeraude Panafricaine",
        subtitle="Agrobusiness, Économie Verte & Terroir",
        description="Design inspiré de la nature et de l'essor économique vert en Afrique.",
        icon="🌿",
        tagline="Entreprendre aujourd'hui, bâtir demain en Afrique",
        vision_quote="Bâtir un écosystème d'excellence et valoriser les richesses du continent africain.",
        banner_bg="0F5132",
        banner_text="FFFFFF",
        header_bg="14532D",
        header_text="FFFFFF",
        accent_color="F59E0B",
        highlight_bg="FEF3C7",
        highlight_text="78350F",
        zebra_bg="F0FDF4",
        border_color="BBF7D0",
        total_bg="DCFCE7",
        total_text="14532D",
        chart_primary="15803D",
        card_bg="F0FDF4",
        card_border="86EFAC",
        badge_success_bg="DCFCE7",
        badge_success_fg="14532D",
        badge_warning_bg="FEF3C7",
        badge_warning_fg="78350F",
        chart_secondary="F59E0B",
        chart_accent="059669"
    ),
    "sahara_commerce": ThemeConfig(
        id="sahara_commerce",
        name="Sahara & Terre Cuite",
        subtitle="Commerce de Détail, Jouets & Distribution",
        description="Chaleureux, élégant et axé sur le merchandising et l'accueil client.",
        icon="🏺",
        tagline="Innover, distribuer et rayonner au cœur des marchés modernes",
        vision_quote="Devenir la référence incontournable de la distribution de qualité et de proximité.",
        banner_bg="9A3412",
        banner_text="FFFFFF",
        header_bg="C2410C",
        header_text="FFFFFF",
        accent_color="D97706",
        highlight_bg="FFEDD5",
        highlight_text="7C2D12",
        zebra_bg="FFF7ED",
        border_color="FED7AA",
        total_bg="FFEDD5",
        total_text="9A3412",
        chart_primary="C2410C",
        card_bg="FFF7ED",
        card_border="FDBA74",
        badge_success_bg="DCFCE7",
        badge_success_fg="14532D",
        badge_warning_bg="FEF3C7",
        badge_warning_fg="78350F",
        chart_secondary="D97706",
        chart_accent="B45309"
    ),
    "royal_navy": ThemeConfig(
        id="royal_navy",
        name="Bleu Royal Institutionnel",
        subtitle="Banque, Industrie & Finance Bancable",
        description="Sérieux formel, inspirant la confiance absolue des comités de crédit bancaire.",
        icon="🏛️",
        tagline="Excellence financière, rigueur et pérennité d'entreprise",
        vision_quote="Garantir une solidité financière exemplaire aux meilleurs standards internationaux.",
        banner_bg="0F172A",
        banner_text="FFFFFF",
        header_bg="1E293B",
        header_text="FFFFFF",
        accent_color="0284C7",
        highlight_bg="E0F2FE",
        highlight_text="075985",
        zebra_bg="F8FAFC",
        border_color="BAE6FD",
        total_bg="E0F2FE",
        total_text="0369A1",
        chart_primary="1E3A8A",
        card_bg="F8FAFC",
        card_border="CBD5E1",
        badge_success_bg="DCFCE7",
        badge_success_fg="14532D",
        badge_warning_bg="FEF3C7",
        badge_warning_fg="78350F",
        chart_secondary="0284C7",
        chart_accent="3B82F6"
    ),
    "cyber_indigo": ThemeConfig(
        id="cyber_indigo",
        name="Cyber Indigo & Tech",
        subtitle="Startups, Digital & Solutions Innovantes",
        description="Moderne, dynamique, conçu pour les levées de fonds et l'innovation numérique.",
        icon="⚡",
        tagline="Accélérer la transformation numérique et l'impact panafricain",
        vision_quote="Créer des solutions technologiques scalables et créatrices de haute valeur ajoutée.",
        banner_bg="312E81",
        banner_text="FFFFFF",
        header_bg="4338CA",
        header_text="FFFFFF",
        accent_color="7C3AED",
        highlight_bg="EDE9FE",
        highlight_text="5B21B6",
        zebra_bg="F5F3FF",
        border_color="DDD6FE",
        total_bg="EDE9FE",
        total_text="4338CA",
        chart_primary="4F46E5",
        card_bg="F5F3FF",
        card_border="C7D2FE",
        badge_success_bg="DCFCE7",
        badge_success_fg="14532D",
        badge_warning_bg="FEF3C7",
        badge_warning_fg="78350F",
        chart_secondary="06B6D4",
        chart_accent="8B5CF6"
    )
}

def detect_or_get_template(template_id: str = None, prompt_text: str = "", activity_text: str = "") -> ThemeConfig:
    """Retourne le template sélectionné ou le détecte intelligemment selon le secteur."""
    if template_id and template_id in TEMPLATE_CATALOG:
        return TEMPLATE_CATALOG[template_id]

    txt = (prompt_text + " " + activity_text).lower()

    if any(k in txt for k in ["jouet", "boutique", "magasin", "commerce", "vente", "mode", "restaurant", "cosmétique", "quincaillerie", "détail"]):
        return TEMPLATE_CATALOG["sahara_commerce"]
    elif any(k in txt for k in ["agri", "chocolat", "savon", "poulet", "volaille", "ferme", "pisciculture", "miel", "anacarde", "manioc", "vert", "bio", "agrotech", "plantation"]):
        return TEMPLATE_CATALOG["emerald_africa"]
    elif any(k in txt for k in ["tech", "logiciel", "app", "digital", "saas", "fintech", "startup", "plateforme", "ia", "intelligence", "cloud", "web"]):
        return TEMPLATE_CATALOG["cyber_indigo"]
    else:
        return TEMPLATE_CATALOG["emerald_africa"]
