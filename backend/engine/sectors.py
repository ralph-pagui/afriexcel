"""
Base de données experte des secteurs d'activité porteurs en Afrique Francophone (OHADA)
Équipements réels, coûts réalistes en FCFA, gammes de produits et structures de charges.
"""

from typing import Dict, Any, List
from .models import SectorProfile, ProductItem, EquipmentItem, StaffItem

SECTOR_CATALOG: Dict[str, SectorProfile] = {
    # 0. EXPLOITATION CACAOYÈRE MODERNE & PRODUCTION DE FÈVES DE CACAO
    "plantation_cacao": SectorProfile(
        sector_id="plantation_cacao",
        name="Exploitation Cacaoyère Moderne du Centre-Cameroun",
        activity="Culture de cacaoyers haut rendement, cabossage, fermentation contrôlée en bacs d'Iroko, séchage solaire et commercialisation de fèves de cacao marchandes de qualité supérieure.",
        unit="Kg de fèves séchées",
        base_daily_volume=49,
        working_days_per_year=280,
        raw_materials_desc="Plants de cacaoyers hybrides certifiés IRAD, engrais organiques et NPK spécifiques cacao, fongicides et insecticides homologués MINADER, sacs en jute pour le stockage export.",
        products=[
            ProductItem(id=1, name="Fèves de cacao Grade 1 (Fermentées & Séchées Export)", description="Fèves de cacao qualité export sous sacs en jute", share_pct=0.65, unit_price=2200),
            ProductItem(id=2, name="Fèves de cacao Grade 2 (Marché National / Transformateurs)", description="Fèves pour transformateurs et broyeurs locaux", share_pct=0.25, unit_price=1700),
            ProductItem(id=3, name="Plants de cacaoyers hybrides en sachet", description="Plants pépinière hybrides IRAD en sachets", share_pct=0.10, unit_price=400)
        ],
        equipment=[
            EquipmentItem(name="Bacs de fermentation en bois dur (Iroko) et claies de séchage solaires", cost=3500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Atomiseurs à moteur thermique et pulvérisateurs dorsaux", cost=1800000, duration_years=4, depreciation_rate=0.25),
            EquipmentItem(name="Motopompe d'irrigation pour pépinière avec kit de tuyauterie", cost=1200000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Bascule mécanique homologuée 500 kg et matériel d'entreposage", cost=1000000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Kit d'outillage de récolte (ébougnoirs, machettes, sécateurs, bacs plastic)", cost=800000, duration_years=3, depreciation_rate=0.33),
            EquipmentItem(name="Véhicule tricycle utilitaire motorisé pour ramassage et logistique", cost=1500000, duration_years=4, depreciation_rate=0.25)
        ],
        staff=[
            StaffItem(title="Chef d'exploitation agricole / Agronome", count=1, monthly_net=200000, monthly_cnps=32000),
            StaffItem(title="Ouvriers agricoles permanents (Entretien & Fermentation)", count=2, monthly_net=80000, monthly_cnps=12800),
            StaffItem(title="Magasinier et gardien de sécurité", count=1, monthly_net=70000, monthly_cnps=11200)
        ]
    ),

    # 1. CHOCOLATERIE ARTISANALE & TRANSFORMATION DU CACAO (BEAN-TO-BAR)
    "chocolaterie": SectorProfile(
        sector_id="chocolaterie",
        name="Chocolaterie Artisanale, Cacao & Confiserie Fine (Bean-to-Bar)",
        activity="Transformation locale de fèves de cacao de terroir en tablettes grand cru, chocolats fins et pâtes à tartiner",
        unit="Tablette 100g / Pot 250g",
        base_daily_volume=300,
        working_days_per_year=280,
        raw_materials_desc="Fèves de cacao sélectionnées grade 1 (Côte d'Ivoire/Cameroun), sucre de canne cristal, beurre de cacao pur, poudre de lait entier, noix de cajou locales, gousses de vanille, emballages alu doré et étuis kraft recyclable",
        products=[
            ProductItem(id=1, name="Tablette Chocolat Noir Terroir 70% Cacao pur", description="Tablette 100g sous étui kraft prestige avec fenêtre & opercule alu or", share_pct=0.45, unit_price=1800),
            ProductItem(id=2, name="Tablette Chocolat au Lait & Éclats de Cajou torréfiés", description="Tablette 100g gourmande aux éclats de noix de cajou caramélisées", share_pct=0.35, unit_price=2000),
            ProductItem(id=3, name="Pâte à tartiner artisanale Cacao-Cajou sans huile de palme", description="Pot en verre hermétique 250g scellé à froid", share_pct=0.20, unit_price=2800)
        ],
        equipment=[
            EquipmentItem(name="Torréfacteur à tambour rotatif pour fèves de cacao (15 kg/lot)", cost=5000000, duration_years=6, depreciation_rate=0.166),
            EquipmentItem(name="Concasseur-tarare séparateur de coques et grué de cacao", cost=3500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Broyeuse-raffineuse concheuse universelle meules granit (50 kg)", cost=6500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Tempéreuse automatique continue avec table vibrante à moules", cost=5500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Armoire frigorifique de cristallisation & démoulage hygrométrie contrôlée", cost=4000000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Lot de moules polycarbonate professionnels gravés & thermosoudeuse", cost=1800000, duration_years=3, depreciation_rate=0.33),
            EquipmentItem(name="Véhicule utilitaire isotherme climatisé pour livraisons boutiques", cost=6500000, duration_years=4, depreciation_rate=0.25),
            EquipmentItem(name="Aménagement de l'atelier aux normes HACCP & balances de précision", cost=3000000, duration_years=8, depreciation_rate=0.125)
        ],
        staff=[
            StaffItem(title="Directeur Général & Maître Chocolatier", count=1, monthly_net=450000, monthly_cnps=75000),
            StaffItem(title="Responsable Production & Torréfaction", count=1, monthly_net=260000, monthly_cnps=43000),
            StaffItem(title="Opérateurs d'Atelier (Moulage & Emballage)", count=3, monthly_net=130000, monthly_cnps=22000),
            StaffItem(title="Commercial & Livreur B2B", count=1, monthly_net=160000, monthly_cnps=27000)
        ]
    ),

    # 2. COSMÉTIQUE & SAVONNERIE NATURELLE (KARITÉ & COCO)
    "savonnerie": SectorProfile(
        sector_id="savonnerie",
        name="Savonnerie & Cosmétiques Naturels au Beurre de Karité",
        activity="Fabrication artisanale et semi-industrielle de savons de toilette saponifiés à froid et baumes exfoliants",
        unit="Pain 150g / Pot 100ml",
        base_daily_volume=350,
        working_days_per_year=280,
        raw_materials_desc="Beurre de karité brut bio, huile de coco vierge, huile de palme raffinée, soude caustique pure, huiles essentielles, argile blanche, miel",
        products=[
            ProductItem(id=1, name="Savon exfoliant au Karité & Miel sauvage", description="Pain 150g sous étui carton biodégradable", share_pct=0.50, unit_price=1000),
            ProductItem(id=2, name="Savon surgras neutre pour peaux sensibles", description="Pain 120g sans parfum sous film végétal", share_pct=0.30, unit_price=800),
            ProductItem(id=3, name="Baume corporel pur Karité fouetté", description="Pot en aluminium 100ml operculé", share_pct=0.20, unit_price=1500)
        ],
        equipment=[
            EquipmentItem(name="Cuve de saponification et chauffe en inox 500L", cost=4500000, duration_years=6, depreciation_rate=0.166),
            EquipmentItem(name="Mélangeur mécanique industriel à vitesse variable", cost=3500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Table de découpe pneumatique & moules silicone grand format", cost=2500000, duration_years=4, depreciation_rate=0.25),
            EquipmentItem(name="Armoire de cure et séchage ventilée thermorégulée", cost=2000000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Machine d'étiquetage automatique & dateuse", cost=2500000, duration_years=4, depreciation_rate=0.25),
            EquipmentItem(name="Véhicule utilitaire de livraison urbaine", cost=6000000, duration_years=4, depreciation_rate=0.25),
            EquipmentItem(name="Aménagement atelier hygiène & matériel laboratoire", cost=2500000, duration_years=6, depreciation_rate=0.166)
        ],
        staff=[
            StaffItem(title="Gérant & Formulateur Cosmétique", count=1, monthly_net=400000, monthly_cnps=68000),
            StaffItem(title="Chef d'Atelier Saponification", count=1, monthly_net=220000, monthly_cnps=37000),
            StaffItem(title="Opératrices Découpe & Conditionnement", count=3, monthly_net=120000, monthly_cnps=20000),
            StaffItem(title="Commercial Boutiques & Pharmacies", count=1, monthly_net=150000, monthly_cnps=25000)
        ]
    ),

    # 3. ÉLEVAGE AVICOLE (POULETS DE CHAIR MODERNES)
    "aviculture": SectorProfile(
        sector_id="aviculture",
        name="Ferme Avicole Moderne & Production de Poulets de Chair",
        activity="Élevage rationnel, engraissement et abattage de volailles de chair destinées aux marchés urbains, restaurants et supermarchés",
        unit="Sujet prêt à cuire (1.8 - 2.2 kg)",
        base_daily_volume=200,
        working_days_per_year=360,
        raw_materials_desc="Poussins d'un jour souche certifiée vaccinés, aliments industriels équilibrés (démarrage, croissance, finition), vitamines, copeaux de litière",
        products=[
            ProductItem(id=1, name="Poulet de chair entier éviscéré calibré", description="Sujet de 1.8 à 2.0 kg sous film étirable étiqueté", share_pct=0.60, unit_price=3000),
            ProductItem(id=2, name="Découpes de poulet (Cuisses & Filets)", description="Barquette fraîcheur scellée de 1 kg", share_pct=0.25, unit_price=3500),
            ProductItem(id=3, name="Fientes de volaille séchées (Engrais bio)", description="Sac tissé de 50 kg pour producteurs maraîchers", share_pct=0.15, unit_price=2500)
        ],
        equipment=[
            EquipmentItem(name="Construction de 2 poulaillers modernes isolés (5 000 sujets)", cost=8500000, duration_years=10, depreciation_rate=0.10),
            EquipmentItem(name="Lignes automatiques d'abreuvement à pipettes & trémies", cost=4000000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Ligne d'abattage semi-automatique avec plumeuse rotative", cost=4500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Chambre froide positive 20m³ & congélateur bahut de réserve", cost=5500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Groupe électrogène de secours 15 kVA automatique", cost=3000000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Tricycle / Pick-up de livraison frigorifié", cost=4500000, duration_years=4, depreciation_rate=0.25)
        ],
        staff=[
            StaffItem(title="Directeur d'Exploitation & Vétérinaire conseil", count=1, monthly_net=380000, monthly_cnps=65000),
            StaffItem(title="Chef d'Élevage & Nutritionniste", count=1, monthly_net=220000, monthly_cnps=37000),
            StaffItem(title="Ouvriers Avicoles & Nettoyage", count=3, monthly_net=110000, monthly_cnps=18000),
            StaffItem(title="Chauffeur Livreur & Vendeur", count=1, monthly_net=140000, monthly_cnps=24000)
        ]
    ),

    # 4. PISCICULTURE MODERNE (TILAPIAS & SILURES HORS-SOL)
    "pisciculture": SectorProfile(
        sector_id="pisciculture",
        name="Pisciculture Moderne & Élevage Intensif de Tilapias et Silures",
        activity="Grossissement intensif en bacs hors-sol recirculés et fumage de poisson d'eau douce",
        unit="Kg de poisson frais",
        base_daily_volume=180,
        working_days_per_year=360,
        raw_materials_desc="Alevins sélectionnés vigoureux (30g), aliments flottants extrudés haute teneur en protéines, sel de traitement, copeaux de bois pour fumoir",
        products=[
            ProductItem(id=1, name="Tilapias frais entiers calibrés", description="Calibre moyen 400g-600g au kilo sur lit de glace", share_pct=0.55, unit_price=2500),
            ProductItem(id=2, name="Silures (Clarias) frais vivants ou fumés", description="Calibre 800g-1.2kg au kilo emballage sous vide", share_pct=0.30, unit_price=3000),
            ProductItem(id=3, name="Alevins certifiés pour pisciculteurs", description="Lot de 100 alevins acclimatés vigoureux", share_pct=0.15, unit_price=5000)
        ],
        equipment=[
            EquipmentItem(name="Installation de 8 bacs circulaires géotextile & tuyauterie PVC", cost=7000000, duration_years=8, depreciation_rate=0.125),
            EquipmentItem(name="Système de pompage solaire immergé & filtration biologique", cost=4500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Aérateurs à palettes solaires & surpresseurs d'oxygène", cost=2500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Four de fumage moderne inox multi-étages", cost=2500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Machine à glace en écailles 300 kg/jour & bacs isothermes", cost=3500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Véhicule tricycle de livraison avec bacs oxygénés", cost=4000000, duration_years=4, depreciation_rate=0.25)
        ],
        staff=[
            StaffItem(title="Responsable Aquacole & Gérant", count=1, monthly_net=380000, monthly_cnps=65000),
            StaffItem(title="Technicien d'Élevage & Contrôle Eau", count=1, monthly_net=220000, monthly_cnps=37000),
            StaffItem(title="Pisciculteurs de Bassin & Fumage", count=2, monthly_net=115000, monthly_cnps=19000),
            StaffItem(title="Commercial Marchés & Restaurants", count=1, monthly_net=140000, monthly_cnps=24000)
        ]
    ),

    # 5. BOULANGERIE & PÂTISSERIE ARTISANALE MODERNE
    "boulangerie": SectorProfile(
        sector_id="boulangerie",
        name="Boulangerie-Pâtisserie Artisanale & Viennoiserie",
        activity="Production et vente quotidienne de pains traditionnels, pains spéciaux enrichis et viennoiseries de qualité",
        unit="Baguette / Pièce viennoiserie",
        base_daily_volume=1800,
        working_days_per_year=360,
        raw_materials_desc="Farine de blé T55 enrichie, levure boulangère fraîche, sel iodé, beurre pur, sucre de canne, chocolat de couverture, emballages kraft",
        products=[
            ProductItem(id=1, name="Baguette tradition croustillante 250g", description="Farine de blé blanche enrichie, fermentation longue", share_pct=0.55, unit_price=200),
            ProductItem(id=2, name="Pains spéciaux (Complet, Maïs, Son)", description="Pain boule de 350g haute valeur nutritionnelle", share_pct=0.25, unit_price=400),
            ProductItem(id=3, name="Viennoiseries fines (Croissants, Pains au chocolat)", description="Pâte feuilletée pur beurre croustillante", share_pct=0.20, unit_price=300)
        ],
        equipment=[
            EquipmentItem(name="Four à soles électrique 4 étages avec buée programmable", cost=12500000, duration_years=8, depreciation_rate=0.125),
            EquipmentItem(name="Pétrin à spirale 60 kg cuve inox & vitesse variable", cost=4000000, duration_years=6, depreciation_rate=0.166),
            EquipmentItem(name="Chambre de pousse contrôlée 16 grilles thermorégulée", cost=3500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Diviseuse volumétrique & façonneuse à baguettes", cost=3000000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Vitrine d'exposition réfrigérée & comptoir magasin", cost=3000000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Tricycle utilitaire aménagé pour dépôts de pain", cost=3000000, duration_years=4, depreciation_rate=0.25)
        ],
        staff=[
            StaffItem(title="Maître Boulanger & Chef de Fournil", count=1, monthly_net=350000, monthly_cnps=60000),
            StaffItem(title="Pâtissier & Tourrier", count=1, monthly_net=220000, monthly_cnps=37000),
            StaffItem(title="Aides-Boulangers & Cuiseurs", count=3, monthly_net=120000, monthly_cnps=20000),
            StaffItem(title="Vendeuses en Boutique & Caisse", count=2, monthly_net=100000, monthly_cnps=17000)
        ]
    ),

    # 6. AGRO-TRANSFORMATION & JUS DE FRUITS NATURELS
    "jus": SectorProfile(
        sector_id="jus",
        name="Unité de Transformation & Jus de Fruits Naturels Pasteurisés",
        activity="Pressage, pasteurisation et conditionnement de jus d'ananas, mangue, gingembre et bissap en bouteilles verre recyclables",
        unit="Bouteille 33cl / 1L",
        base_daily_volume=500,
        working_days_per_year=280,
        raw_materials_desc="Fruits frais locaux (Ananas pain de sucre, mangues kente, racines de gingembre, calices d'hibiscus/bissap), sucre de canne, bouteilles verre 33cl/1L et capsules twist-off",
        products=[
            ProductItem(id=1, name="Pur Jus d'Ananas Pain de Sucre 100% Naturel", description="Bouteille en verre 33cl sans conservateur ni eau ajoutée", share_pct=0.45, unit_price=600),
            ProductItem(id=2, name="Cocktail Énergisant Gingembre-Ananas-Bissap", description="Bouteille en verre 33cl aux épices naturelles", share_pct=0.35, unit_price=700),
            ProductItem(id=3, name="Jus de Mangue Onctueux format familial 1L", description="Bouteille 1L verre hermétique pour réceptions et hôtels", share_pct=0.20, unit_price=1500)
        ],
        equipment=[
            EquipmentItem(name="Presse hydraulique industrielle & extracteur de jus inox", cost=5500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Pasteurisateur tubulaire continu en acier inoxydable 300L/h", cost=6500000, duration_years=6, depreciation_rate=0.166),
            EquipmentItem(name="Ligne de soutirage remplisseuse-capsuleuse semi-automatique", cost=4500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Bacs de lavage fruits avec bullage d'air & table de parage", cost=2000000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Chambre froide positive 15m³ pour stockage fruits & jus", cost=4500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Véhicule utilitaire de distribution supermarchés & maquis", cost=5500000, duration_years=4, depreciation_rate=0.25)
        ],
        staff=[
            StaffItem(title="Responsable Agroalimentaire & Gérant", count=1, monthly_net=380000, monthly_cnps=65000),
            StaffItem(title="Opérateur Pasteurisateur & Contrôle Qualité", count=1, monthly_net=220000, monthly_cnps=37000),
            StaffItem(title="Opérateurs Épluchage, Pressage & Remplissage", count=3, monthly_net=110000, monthly_cnps=18000),
            StaffItem(title="Commercial Grande Distribution & Horeca", count=1, monthly_net=150000, monthly_cnps=25000)
        ]
    ),

    # 7. CONFECTION TEXTILE & STYLISME WAX / FASO DAN FANI
    "textile": SectorProfile(
        sector_id="textile",
        name="Atelier de Confection Textile & Stylisme Moderne Africain",
        activity="Création et confection en série de tenues professionnelles, uniformes scolaires et vêtements prêt-à-porter en pagne Wax et cotonnades",
        unit="Pièce confectionnée",
        base_daily_volume=35,
        working_days_per_year=280,
        raw_materials_desc="Pagnes Wax certifiés, tissage Faso Dan Fani / Kente, tissu tergal infroissable, doublures soyeuses, fils industriels, fermetures YKK, boutons",
        products=[
            ProductItem(id=1, name="Tenues de travail professionnelles & blouses d'atelier", description="Tissu tergal ultra-résistant brodé au logo client", share_pct=0.45, unit_price=9500),
            ProductItem(id=2, name="Chemises et Robes modernes en Wax & Dan Fani", description="Prêt-à-porter cintré chic finitions haut de gamme", share_pct=0.35, unit_price=18000),
            ProductItem(id=3, name="Accessoires assortis (Sacs d'ordinateur, pochettes)", description="Doublure molletonnée avec finitions cuir synthétique", share_pct=0.20, unit_price=6500)
        ],
        equipment=[
            EquipmentItem(name="Lot de 6 machines à coudre industrielles coupe-fil Juki", cost=4800000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="2 Surjeteuses professionnelles 4 fils industrielles", cost=2200000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Machine à broder numérique 4 têtes programmable", cost=5500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Grande table de découpe 4m x 2m avec ciseau vertical électrique", cost=2000000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Centrale vapeur industrielle avec table à repasser aspirante", cost=1800000, duration_years=4, depreciation_rate=0.25),
            EquipmentItem(name="Agencement showroom de vente & mannequins d'exposition", cost=2500000, duration_years=5, depreciation_rate=0.20)
        ],
        staff=[
            StaffItem(title="Styliste Modéliste & Directrice d'Atelier", count=1, monthly_net=380000, monthly_cnps=65000),
            StaffItem(title="Chef d'Atelier Coupe & Assemblage", count=1, monthly_net=220000, monthly_cnps=37000),
            StaffItem(title="Couturiers et Piqueuses Qualifiés", count=4, monthly_net=120000, monthly_cnps=20000),
            StaffItem(title="Commercial Entreprises & Écoles", count=1, monthly_net=150000, monthly_cnps=25000)
        ]
    ),

    # 8. PHARMACIE D'OFFICINE & PARAPHARMACIE MODERNE
    "pharmacie": SectorProfile(
        sector_id="pharmacie",
        name="Pharmacie d'Officine & Espace Parapharmacie Moderne",
        activity="Dispensation de médicaments sous ordonnance, médicaments génériques DCI, conseils santé et vente de produits de parapharmacie et matériel médical",
        unit="Ordonnance / Boîte / Article",
        base_daily_volume=160,
        working_days_per_year=360,
        raw_materials_desc="Approvisionnement auprès des grossistes répartiteurs pharmaceutiques agréés (Laborex, Copharmed, Tedis, UBIPHARM) en spécialités princeps, génériques DCI, solutés, vaccins et dermo-cosmétique",
        products=[
            ProductItem(id=1, name="Spécialités éthiques & Médicaments d'ordonnance", description="Traitements chroniques, antibiotiques et spécialités sous prescription", share_pct=0.55, unit_price=4500),
            ProductItem(id=2, name="Médicaments génériques DCI essentiels", description="Molécules génériques de qualité certifiée à forte rotation", share_pct=0.30, unit_price=2000),
            ProductItem(id=3, name="Parapharmacie, Puériculture & Matériel Médical", description="Soins dermo-cosmétiques, laits infantiles, tensiomètres et pansements", share_pct=0.15, unit_price=5500)
        ],
        equipment=[
            EquipmentItem(name="Agencement complet officine : rayonnages métalliques, tiroirs sécurisés & comptoirs", cost=12000000, duration_years=8, depreciation_rate=0.125),
            EquipmentItem(name="Logiciel de gestion officine (LGPI / Pharmagest) & 3 postes informatiques complets", cost=4500000, duration_years=4, depreciation_rate=0.25),
            EquipmentItem(name="Armoire frigorifique médicale qualifiée +2°C/+8°C avec enregistreur de température", cost=3500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Croix de pharmacie lumineuse LED programmable extérieure & enseigne", cost=2500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Onduleur haute capacité & groupe électrogène de secours 15 kVA", cost=4000000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Système de vidéosurveillance HD & alarme connectée", cost=2000000, duration_years=4, depreciation_rate=0.25),
            EquipmentItem(name="Climatisation intégrale de l'officine (préservation stricte des médicaments)", cost=3000000, duration_years=5, depreciation_rate=0.20)
        ],
        staff=[
            StaffItem(title="Docteur en Pharmacie (Pharmacien Titulaire / Gérant)", count=1, monthly_net=650000, monthly_cnps=110000),
            StaffItem(title="Préparateurs en Pharmacie diplômés", count=2, monthly_net=220000, monthly_cnps=37000),
            StaffItem(title="Caissières vendeuses en pharmacie", count=2, monthly_net=120000, monthly_cnps=20000),
            StaffItem(title="Magasinier gestionnaire des réceptions & stocks", count=1, monthly_net=140000, monthly_cnps=24000)
        ]
    ),

    # 9. CLINIQUE MÉDICALE & CABINET DE SOINS
    "clinique": SectorProfile(
        sector_id="clinique",
        name="Clinique Médicale & Centre de Soins Polyvalent",
        activity="Consultations de médecine générale, pédiatrie, gynécologie, examens de laboratoire et hospitalisations de jour",
        unit="Acte médical / Consultation / Journée",
        base_daily_volume=45,
        working_days_per_year=360,
        raw_materials_desc="Consommables médicaux stériles, réactifs de laboratoire d'analyses, médicaments d'urgence, matériel à usage unique",
        products=[
            ProductItem(id=1, name="Consultations médicales spécialisées & Générales", description="Médecine générale, pédiatrie, gynécologie obstétrique", share_pct=0.45, unit_price=10000),
            ProductItem(id=2, name="Analyses de laboratoire & Imagerie de base", description="NFS, glycémie, goutte épaisse, échographie générale", share_pct=0.35, unit_price=15000),
            ProductItem(id=3, name="Hospitalisation de jour & Petits actes de chirurgie", description="Mise en observation, perfusions, pansements et petite chirurgie", share_pct=0.20, unit_price=25000)
        ],
        equipment=[
            EquipmentItem(name="Échographe doppler couleur polyvalent avec sondes", cost=9500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Automate d'analyses médicales hématologie & biochimie", cost=7000000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Lits d'hospitalisation médicalisés & monitoring multiparamétrique", cost=5000000, duration_years=6, depreciation_rate=0.166),
            EquipmentItem(name="Table de consultation gynécologique & matériel chirurgical stérile", cost=3500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Autoclave de stérilisation à vapeur d'eau 50L", cost=2500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Groupe électrogène de secours 20 kVA à démarrage automatique", cost=4500000, duration_years=5, depreciation_rate=0.20)
        ],
        staff=[
            StaffItem(title="Médecin Directeur / Médecin Généraliste", count=1, monthly_net=700000, monthly_cnps=120000),
            StaffItem(title="Infirmiers d'État diplômés", count=2, monthly_net=230000, monthly_cnps=39000),
            StaffItem(title="Technicien Supérieur de Laboratoire", count=1, monthly_net=250000, monthly_cnps=42000),
            StaffItem(title="Aides-soignantes & Accueil", count=2, monthly_net=110000, monthly_cnps=18000)
        ]
    ),

    # 10. QUINCAILLERIE GÉNÉRALE & MATÉRIAUX DE CONSTRUCTION
    "quincaillerie": SectorProfile(
        sector_id="quincaillerie",
        name="Quincaillerie Moderne & Matériaux de Construction",
        activity="Vente de ciment, fer à béton, outillage professionnel, plomberie, électricité et peintures",
        unit="Article / Sac / Barre",
        base_daily_volume=90,
        working_days_per_year=312,
        raw_materials_desc="Cimenterie locale (CIMAF, Lafarge), tréfileries de fer à béton, fabricants de peinture, importateurs d'outillage",
        products=[
            ProductItem(id=1, name="Gros œuvre : Ciment, Fer à béton & Agrégats", description="Sacs de ciment 50kg CPJ 42.5 et barres de fer FE 500", share_pct=0.55, unit_price=4500),
            ProductItem(id=2, name="Peintures, Enduits & Produits d'étanchéité", description="Seaux de peinture acrylique pro et monocouche", share_pct=0.25, unit_price=18000),
            ProductItem(id=3, name="Plomberie, Électricité & Outillage de chantier", description="Tubes PVC, câblage certifié, disjoncteurs et électroportatif", share_pct=0.20, unit_price=7500)
        ],
        equipment=[
            EquipmentItem(name="Camionnette benne utilitaire 3.5 tonnes pour livraisons chantiers", cost=12000000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Chariot élévateur manuel & transpalettes hydrauliques", cost=2500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Machine à teinter automatique de peinture & mélangeur", cost=4500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Rayonnages métalliques charges lourdes pour magasin et cour", cost=4000000, duration_years=8, depreciation_rate=0.125),
            EquipmentItem(name="Système informatique de caisse, lecteur code-barres et balance", cost=2000000, duration_years=3, depreciation_rate=0.33)
        ],
        staff=[
            StaffItem(title="Gérant Quincaillerie & Acheteur", count=1, monthly_net=400000, monthly_cnps=68000),
            StaffItem(title="Vendeurs conseils en outillage & bâtiment", count=2, monthly_net=150000, monthly_cnps=25000),
            StaffItem(title="Chauffeur livreur poids léger", count=1, monthly_net=140000, monthly_cnps=24000),
            StaffItem(title="Manutentionnaires déchargement", count=2, monthly_net=100000, monthly_cnps=17000)
        ]
    ),

    # 11. RESTAURATION MODERNE & FAST-FOOD AFRICAIN
    "restaurant": SectorProfile(
        sector_id="restaurant",
        name="Restaurant Moderne & Traiteur aux Saveurs d'Afrique",
        activity="Restauration sur place, plats à emporter et livraison de spécialités africaines revisitées et grillades",
        unit="Plat / Menu / Formule",
        base_daily_volume=120,
        working_days_per_year=360,
        raw_materials_desc="Viandes et poissons frais de première qualité, féculents locaux (attiéké, alloco, igname, riz parfumé), légumes maraîchers, épices du terroir, emballages biodégradables",
        products=[
            ProductItem(id=1, name="Formule Déjeuner Express (Plats signatures & Grillades)", description="Menu complet avec boisson locale naturelle", share_pct=0.55, unit_price=3500),
            ProductItem(id=2, name="Spécialités Braisées & Plats de prestige du soir", description="Poisson braisé, choukouya, brochettes et accompagnements", share_pct=0.30, unit_price=6000),
            ProductItem(id=3, name="Service Traiteur & Cocktails d'entreprises", description="Plateaux repas et buffets événementiels par personne", share_pct=0.15, unit_price=10000)
        ],
        equipment=[
            EquipmentItem(name="Piano de cuisson 6 feux inox avec four gaz professionnel", cost=4500000, duration_years=6, depreciation_rate=0.166),
            EquipmentItem(name="Grilloir à charbon & braiseur inox avec hotte aspirante motorisée", cost=3500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Chambre froide positive 10m³ & congélateur bahut 500L", cost=5000000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Mobilier de salle moderne : 15 tables & 60 chaises confort", cost=4000000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Matériel de cuisine inox (batterie, casseroles, bacs gastro, vaisselle)", cost=3000000, duration_years=4, depreciation_rate=0.25),
            EquipmentItem(name="Caisse enregistreuse tactile, imprimante commande cuisine & TPE", cost=1800000, duration_years=3, depreciation_rate=0.33)
        ],
        staff=[
            StaffItem(title="Chef de Cuisine & Gérant", count=1, monthly_net=400000, monthly_cnps=68000),
            StaffItem(title="Seconds de cuisine & Cuisiniers grillades", count=2, monthly_net=170000, monthly_cnps=29000),
            StaffItem(title="Serveurs en salle & Caisse", count=3, monthly_net=110000, monthly_cnps=18000),
            StaffItem(title="Plongeur & Hygiène salle", count=1, monthly_net=95000, monthly_cnps=16000)
        ]
    ),

    # 12. MAGASIN DE JOUETS & ARTICLES LUDO-ÉDUCATIFS
    "magasin_jouets": SectorProfile(
        sector_id="magasin_jouets",
        name="Magasin de Jouets & Jeux Éducatifs",
        activity="Vente au détail de jouets d'éveil, jeux de société, figurines, puzzles et articles récréatifs pour enfants",
        unit="Article / Jouet",
        base_daily_volume=25,
        working_days_per_year=312,
        raw_materials_desc="Fournisseurs grossistes agréés de jouets éducatifs, distributeurs officiels et artisans créateurs locaux",
        products=[
            ProductItem(id=1, name="Jeux d'éveil & Premier âge (bois, cubes, hochets)", description="Jouets ludo-éducatifs sensoriels pour les tout-petits", share_pct=0.40, unit_price=6500),
            ProductItem(id=2, name="Poupées, Figurines de héros & Miniatures", description="Univers ludiques, figurines d'action et véhicules pour enfants", share_pct=0.35, unit_price=8500),
            ProductItem(id=3, name="Jeux de société familiaux & Plein air (trottinettes)", description="Jeux de stratégie familiale et activités extérieures", share_pct=0.25, unit_price=14000)
        ],
        equipment=[
            EquipmentItem(name="Agencement boutique : rayonnages colorés, présentoirs & bacs", cost=2800000, duration_years=8, depreciation_rate=0.125),
            EquipmentItem(name="Système de caisse tactile, lecteur code-barres & logiciel stock", cost=1200000, duration_years=4, depreciation_rate=0.25),
            EquipmentItem(name="Climatiseur Split inverter 18 000 BTU pour confort boutique", cost=650000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Enseigne lumineuse extérieure colorée & vitrophanie enfantine", cost=850000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Kit sécurité 4 caméras surveillance HD & onduleur de secours", cost=500000, duration_years=4, depreciation_rate=0.25)
        ],
        staff=[
            StaffItem(title="Gérant(e) Responsable de Boutique & Achats", count=1, monthly_net=220000, monthly_cnps=36000),
            StaffItem(title="Vendeuse-Conseil & Accueil Enfants", count=1, monthly_net=120000, monthly_cnps=20000)
        ]
    )
}

def detect_sector(prompt: str) -> SectorProfile:
    """Détecte intelligemment le profil sectoriel approprié à partir de la demande utilisateur."""
    p = prompt.lower()
    
    # 1. Pharmacie & Médical
    if any(k in p for k in ["pharmacie", "officine", "médicament", "medicament", "santé", "sante", "parapharmacie"]):
        return SECTOR_CATALOG["pharmacie"]
    # 2. Clinique & Cabinet de soins
    elif any(k in p for k in ["clinique", "cabinet médical", "cabinet medical", "soins", "hôpital", "hopital", "docteur", "dispensaire"]):
        return SECTOR_CATALOG["clinique"]
    # 3. Quincaillerie & Matériaux
    elif any(k in p for k in ["quincaillerie", "ciment", "fer à béton", "matériaux", "bâtiment", "quincailler"]):
        return SECTOR_CATALOG["quincaillerie"]
    # 4. Restauration & Traiteur
    elif any(k in p for k in ["restaurant", "restauration", "fast food", "fast-food", "traiteur", "maquis", "grillade", "choukouya", "bar"]):
        return SECTOR_CATALOG["restaurant"]
    # 4b. Plantation & Culture de Cacao
    elif ("cacao" in p or "cacaoy" in p) and any(k in p for k in ["plantation", "culture", "champ", "agricole", "exploitation", "fève", "feve", "cabosse"]):
        return SECTOR_CATALOG["plantation_cacao"]
    # 5. Chocolat & Transformation
    elif any(k in p for k in ["chocolat", "cacao", "confiserie", "praline", "pâte à tartiner", "pate a tartiner"]):
        return SECTOR_CATALOG["chocolaterie"]
    # 6. Savonnerie & Karité
    elif any(k in p for k in ["savon", "karité", "karite", "cosmétique", "cosmetique", "baume", "pommade"]):
        return SECTOR_CATALOG["savonnerie"]
    # 7. Élevage avicole
    elif any(k in p for k in ["poulet", "volaille", "avicole", "poussin", "œuf", "oeuf", "ferme avicole"]):
        return SECTOR_CATALOG["aviculture"]
    # 8. Pisciculture
    elif any(k in p for k in ["poisson", "pisciculture", "aquaculture", "tilapia", "silure", "alevin"]):
        return SECTOR_CATALOG["pisciculture"]
    # 9. Boulangerie
    elif any(k in p for k in ["pain", "boulangerie", "pâtisserie", "patisserie", "viennoiserie", "croissant"]):
        return SECTOR_CATALOG["boulangerie"]
    # 10. Jus de fruits
    elif any(k in p for k in ["jus", "fruit", "ananas", "mangue", "gingembre", "bissap", "boisson"]):
        return SECTOR_CATALOG["jus"]
    # 11. Textile & Mode
    elif any(k in p for k in ["textile", "couture", "vêtement", "vetement", "wax", "dan fani", "mode", "confection"]):
        return SECTOR_CATALOG["textile"]
    # 12. Magasin de Jouets & Articles Enfants
    elif any(k in p for k in ["jouet", "jouets", "ludique", "ludo", "peluche", "poupée", "poupee"]):
        return SECTOR_CATALOG["magasin_jouets"]
    
    # 13. Profil dynamique par défaut (adapté au mot clé avec nettoyage chirurgical)
    import re
    clean_p = prompt.strip()
    clean_p = re.sub(r'^(?:je\s+veux|je\s+souhaite|fais[- ]moi|peux[- ]tu\s+faire|cr[ée]e[- ]moi|g[ée]n[èe]re[- ]moi|donne[- ]moi|il\s+me\s+faut)\s+', '', clean_p, flags=re.I)
    clean_p = re.sub(r'^(?:un|le|mon|du)?\s*(?:business\s*plan|plan\s*d[\'’]affaires?|projet)\s+(?:complet|bancable|d[ée]taill[ée])?\s*', '', clean_p, flags=re.I)
    clean_p = re.sub(r'^(?:pour|sur|concernant)\s+', '', clean_p, flags=re.I)
    clean_p = re.sub(r'^(?:mon|ma|mes|notre|nos|un|une|des|le|la|les)\s+', '', clean_p, flags=re.I)
    clean_p = re.sub(r'\s+(?:sur\s+\d+\s*ans?|avec\s+(?:un\s+)?budget|\s*avec\s+\d+|j[\'’]ai\s+|\bà\s+[A-Za-zÀ-ÿ]+|\ben\s+[A-Za-zÀ-ÿ]+|\bau\s+[A-Za-zÀ-ÿ]+).*$', '', clean_p, flags=re.I)
    clean_title = clean_p.strip(' ,;.')
    if clean_title:
        clean_title = clean_title[0].upper() + clean_title[1:]
    else:
        clean_title = "Entreprise Commerciale & Prestation"

    return SectorProfile(
        sector_id="entreprise_generale",
        name=f"Projet - {clean_title}",
        activity=f"Exploitation et développement d'activité commerciale et de services : {clean_title}",
        unit="Prestation / Article",
        base_daily_volume=80,
        working_days_per_year=300,
        raw_materials_desc="Approvisionnements spécifiques, consommables d'exploitation et fournitures professionnelles",
        products=[
            ProductItem(id=1, name="Offre Principale Standard", description="Prestation ou produit phare de référence", share_pct=0.55, unit_price=5000),
            ProductItem(id=2, name="Offre Supérieure Premium", description="Option à haute valeur ajoutée avec service client dédié", share_pct=0.30, unit_price=12000),
            ProductItem(id=3, name="Services Complémentaires & Accessoires", description="Produits dérivés et prestations annexes", share_pct=0.15, unit_price=3500)
        ],
        equipment=[
            EquipmentItem(name="Aménagement du local d'exploitation & enseigne lumineuse", cost=6000000, duration_years=8, depreciation_rate=0.125),
            EquipmentItem(name="Équipements métiers spécifiques & outillage professionnel", cost=8500000, duration_years=5, depreciation_rate=0.20),
            EquipmentItem(name="Matériel informatique, caisse tactile et logiciel de gestion", cost=2500000, duration_years=4, depreciation_rate=0.25),
            EquipmentItem(name="Véhicule utilitaire de livraison ou d'intervention terrain", cost=6500000, duration_years=4, depreciation_rate=0.25),
            EquipmentItem(name="Groupe électrogène de secours et onduleur", cost=3000000, duration_years=5, depreciation_rate=0.20)
        ],
        staff=[
            StaffItem(title="Directeur / Gérant d'Entreprise", count=1, monthly_net=450000, monthly_cnps=75000),
            StaffItem(title="Technicien spécialisé / Responsable des opérations", count=2, monthly_net=200000, monthly_cnps=34000),
            StaffItem(title="Commercial / Vendeur qualifié", count=1, monthly_net=150000, monthly_cnps=25000),
            StaffItem(title="Assistant(e) administratif(ve) et caisse", count=1, monthly_net=120000, monthly_cnps=20000)
        ]
    )
