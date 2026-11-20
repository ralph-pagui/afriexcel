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

def detect_or_get_template(template_id: str = None, prompt_text: str = "", activity_text: str = "", sector_id: str = "") -> ThemeConfig:
    """
    Retourne le template sélectionné ou le détecte intelligemment selon le secteur
    sans jamais imposer de textes ou de sous-titres inappropriés.
    """
    from .prompt_parser import PromptParser
    parsed = PromptParser.parse(prompt_text + " " + activity_text)
    effective_sec_id = sector_id or parsed.get("sector_id", "general")

    if template_id and template_id in TEMPLATE_CATALOG and template_id != "auto":
        base_tpl = TEMPLATE_CATALOG[template_id]
        return _adapt_template_text(base_tpl, prompt_text, activity_text, effective_sec_id)

    # Sélection du thème de design selon la nature du projet
    agro_sectors = ["aviculture", "agriculture", "plantation_cacao", "pisciculture", "jus", "chocolaterie", "savonnerie"]
    tech_sectors = ["numerique_saas", "energie_solaire", "commerce_informatique"]
    commerce_sectors = ["magasin_jouets", "quincaillerie", "restaurant", "textile", "boulangerie", "commerce_general"]
    corp_sectors = ["transport", "clinique", "pharmacie", "btp", "education"]

    if effective_sec_id in agro_sectors:
        chosen_id = "emerald_africa"
    elif effective_sec_id in tech_sectors:
        chosen_id = "cyber_indigo"
    elif effective_sec_id in commerce_sectors:
        chosen_id = "sahara_commerce"
    elif effective_sec_id in corp_sectors:
        chosen_id = "royal_navy"
    else:
        chosen_id = "royal_navy"

    base_tpl = TEMPLATE_CATALOG[chosen_id]
    return _adapt_template_text(base_tpl, prompt_text, activity_text, effective_sec_id)

def _adapt_template_text(base: ThemeConfig, prompt_text: str, activity_text: str, sector_id: str = "") -> ThemeConfig:
    """Adapte dynamiquement le sous-titre, slogan et vision au secteur spécifique sans aucune contamination."""
    import copy
    tpl = copy.deepcopy(base)

    txt = (prompt_text + " " + activity_text).lower()

    # Priorité 1 : Matching strict par sector_id
    if sector_id == "jus":
        tpl.subtitle = "Industrie Agroalimentaire & Transformation de Fruits"
        tpl.tagline = "Valorisation des fruits locaux en jus naturels de qualité supérieure"
        tpl.vision_quote = "Bâtir une référence industrielle de la transformation agroalimentaire locale, saine et compétitive."
        return tpl

    if sector_id == "commerce_informatique":
        tpl.subtitle = "Distribution & Services Informatiques, Réseaux & Équipements"
        tpl.tagline = "Solutions technologiques performantes et équipements informatiques de pointe"
        tpl.vision_quote = "Accompagner la transformation numérique des entreprises par des équipements et services fiables."
        return tpl

    if sector_id == "transport":
        tpl.subtitle = "Transport Routier, Interurbain & Fret Logistique"
        tpl.tagline = "Connecter les corridors économiques avec rapidité, sécurité et fiabilité"
        tpl.vision_quote = "Devenir le partenaire logistique de référence pour le transport de marchandises en Afrique de l'Ouest."
        return tpl

    if sector_id == "aviculture":
        tpl.subtitle = "Élevage Avicole & Production de Volailles de Chair"
        tpl.tagline = "Excellence agro-pastorale, traçabilité et disponibilité locale continue"
        tpl.vision_quote = "Bâtir une filière avicole moderne et souveraine au service de l'alimentation de qualité."
        return tpl

    if sector_id == "agriculture":
        tpl.subtitle = "Agro-Business, Maraîchage & Cultures Maraîchères Durables"
        tpl.tagline = "Valoriser la terre africaine par des cultures saines, productives et locales"
        tpl.vision_quote = "Garantir un approvisionnement régulier en légumes frais et de qualité supérieure aux marchés locaux."
        return tpl

    if sector_id == "plantation_cacao":
        tpl.subtitle = "Exploitation Cacaoyère & Filière d'Excellence"
        tpl.tagline = "Valoriser le terroir africain avec des fèves de qualité supérieure"
        tpl.vision_quote = "Produire durablement en alliant rentabilité agro-écologique et qualité export."
        return tpl

    if sector_id == "pisciculture":
        tpl.subtitle = "Aquaculture Moderne & Élevage Piscicole"
        tpl.tagline = "Production durable de poissons de qualité supérieure en bassins aménagés"
        tpl.vision_quote = "Contribuer à la sécurité alimentaire par une aquaculture moderne et rentable."
        return tpl

    if sector_id == "clinique":
        tpl.subtitle = "Santé, Biologie Médicale & Analyses de Proximité"
        tpl.tagline = "Précision diagnostique, excellence clinique et accessibilité des soins"
        tpl.vision_quote = "Offrir un plateau technique médical de pointe garantissant des diagnostics fiables et bienveillants."
        return tpl

    if sector_id == "pharmacie":
        tpl.subtitle = "Officine Pharmaceutique & Produits de Santé"
        tpl.tagline = "Disponibilité continue des médicaments essentiels et conseil pharmaceutique de pointe"
        tpl.vision_quote = "Sécuriser l'accès aux soins et aux produits de santé avec une rigueur absolue."
        return tpl

    if sector_id == "magasin_jouets":
        tpl.subtitle = "Commerce de Détail, Jouets & Éveil Enfants"
        tpl.tagline = "Éveiller l'imaginaire et accompagner l'épanouissement des enfants"
        tpl.vision_quote = "Devenir la référence incontournable du jeu éducatif de qualité."
        return tpl

    if sector_id == "quincaillerie":
        tpl.subtitle = "Commerce de Matériaux de Construction & Quincaillerie Générale"
        tpl.tagline = "Fourniture fiable de matériaux et d'outillages pour le bâtiment et l'industrie"
        tpl.vision_quote = "Être le partenaire privilégié des artisans, chantiers et constructeurs locaux."
        return tpl

    if sector_id == "boulangerie":
        tpl.subtitle = "Boulangerie-Pâtisserie Artisanale & Moderne"
        tpl.tagline = "Pains frais, viennoiseries dorées et saveurs authentiques au quotidien"
        tpl.vision_quote = "Offrir chaque matin des produits de boulangerie frais, sains et savoureux."
        return tpl

    if sector_id == "restaurant":
        tpl.subtitle = "Restauration & Gastronomie"
        tpl.tagline = "Saveurs authentiques, convivialité et excellence du service"
        tpl.vision_quote = "Offrir une expérience culinaire raffinée et accessible valorisant les saveurs locales."
        return tpl

    if sector_id == "energie_solaire":
        tpl.subtitle = "Énergies Renouvelables & Solutions Solaires"
        tpl.tagline = "L'énergie propre et durable au cœur du développement africain"
        tpl.vision_quote = "Fournir une énergie solaire continue et compétitive aux entreprises et ménages."
        return tpl

    if sector_id == "numerique_saas":
        tpl.subtitle = "Technologies Numériques, Solutions Logicielles & SaaS"
        tpl.tagline = "Solutions logicielles scalables et transformation digitale"
        tpl.vision_quote = "Digitaliser les processus métiers avec agilité et robustesse technologique."
        return tpl

    # Priorité 2 : Détection par mots-clés spécifiques uniquement si sector_id inconnu
    if any(w in txt for w in ["jus de", "bouteille de jus", "mangue", "ananas", "fruits"]):
        tpl.subtitle = "Industrie Agroalimentaire & Transformation de Fruits"
        tpl.tagline = "Valorisation des fruits locaux en jus naturels de qualité supérieure"
        tpl.vision_quote = "Bâtir une référence industrielle de la transformation agroalimentaire locale, saine et compétitive."
    elif any(w in txt for w in ["informatique", "ordinateur", "matériel informatique", "serveurs", "réseaux"]):
        tpl.subtitle = "Distribution & Services Informatiques, Réseaux & Équipements"
        tpl.tagline = "Solutions technologiques performantes et équipements informatiques de pointe"
        tpl.vision_quote = "Accompagner la transformation numérique des entreprises par des équipements et services fiables."
    elif any(w in txt for w in ["camion de transport", "transport routier", "fret routier"]):
        tpl.subtitle = "Transport Routier, Interurbain & Fret Logistique"
        tpl.tagline = "Connecter les corridors économiques avec rapidité, sécurité et fiabilité"
        tpl.vision_quote = "Devenir le partenaire logistique de référence pour le transport de marchandises en Afrique de l'Ouest."
    elif any(w in txt for w in ["poulet", "poussin", "volaille", "avicole", "couvoir"]):
        tpl.subtitle = "Élevage Avicole & Production de Volailles de Chair"
        tpl.tagline = "Excellence agro-pastorale, traçabilité et disponibilité locale continue"
        tpl.vision_quote = "Bâtir une filière avicole moderne et souveraine au service de l'alimentation de qualité."
    elif any(w in txt for w in ["tomate", "maraîchage", "maraichage", "légume"]):
        tpl.subtitle = "Agro-Business, Maraîchage & Cultures Maraîchères Durables"
        tpl.tagline = "Valoriser la terre africaine par des cultures saines, productives et locales"
        tpl.vision_quote = "Garantir un approvisionnement régulier en légumes frais et de qualité supérieure aux marchés locaux."
    elif any(w in txt for w in ["clinique", "médical", "medical", "laboratoire d'analyses"]):
        tpl.subtitle = "Santé, Biologie Médicale & Analyses de Proximité"
        tpl.tagline = "Précision diagnostique, excellence clinique et accessibilité des soins"
        tpl.vision_quote = "Offrir un plateau technique médical de pointe garantissant des diagnostics fiables et bienveillants."
    else:
        tpl.subtitle = "Entreprise Commerciale & Prestations de Services"
        tpl.tagline = "Rigueur opérationnelle, performance financière et pérennité"
        tpl.vision_quote = "Développer un modèle d'affaires solide, rentable et générateur d'emplois durables."

    return tpl
