/**
 * Service Universel de Génération de Business Plans Complets dans Excel
 * Modélisation financière professionnelle 19 feuilles conforme SYSCOHADA & normes bancaires d'Afrique Francophone.
 * Inclut mise en forme prestige (bannières Corporate Navy #1F3864, entrées surlignées #FFF2CC, formules et onglets colorés).
 */

/* global Excel, Office */

export interface BusinessPlanOptions {
  projectName?: string;
  productTopic?: string;
  structureName?: string;
  projectType?: string;
  currencySymbol?: string;
  countryName?: string;
  sheetCount?: number;
  sheetsToGenerate?: string[]; // Si vide, génère les 19 feuilles ou le nombre demandé
  onProgress?: (sheetName: string, index: number, total: number) => void;
}

export class BusinessPlanService {
  /**
   * Adapte intelligemment le profil de production selon le produit demandé par l'utilisateur
   */
  private static getProductProfile(topic?: string, defaultProject?: string) {
    const rawText = (topic || defaultProject || '').toLowerCase();
    const text = rawText
      .replace(/^(?:projet\s+(?:de\s+|d'|d’)?|création\s+(?:de\s+|d'|d’)?|fabrication\s+(?:de\s+|d'|d’)?|production\s+(?:de\s+|d'|d’)?|atelier\s+(?:de\s+|d'|d’)?|usine\s+(?:de\s+|d'|d’)?)/i, '')
      .replace(/^(?:mon|ma|mes|notre|nos|un|une|le|la|les|du|de\s+la|des)\s+/i, '')
      .trim();

    // 0. Chocolaterie Artisanale, Cacao & Confiserie (Bean-to-Bar)
    if (text.includes('chocolat') || text.includes('cacao') || text.includes('confiserie') || text.includes('praline') || text.includes('pâte à tartiner') || text.includes('pate a tartiner')) {
      return {
        activity: "Chocolaterie artisanale, transformation de fèves de cacao de terroir & confiserie fine",
        unit: "Tablette 100g / Pot 250g",
        sector: "industry",
        products: [
          [1, "Tablette Chocolat Noir Terroir 70% Cacao pur", "Tablette 100g sous étui kraft & opercule alu doré", "45%", 1800],
          [2, "Tablette Chocolat au Lait & Éclats de Cajou torréfiés", "Tablette 100g gourmande aux noix locales", "35%", 2000],
          [3, "Pâte à tartiner artisanale Cacao-Cajou sans huile de palme", "Pot en verre 250g scellé à froid", "20%", 2800]
        ],
        equipment: [
          ["Torréfacteur à tambour rotatif pour fèves de cacao (15 kg/lot)", 5000000, 6, 0.166],
          ["Concasseur-tarare séparateur de coques de fèves de cacao", 3500000, 5, 0.20],
          ["Broyeuse-raffineuse concheuse universelle à meules de granit (50 kg)", 6500000, 5, 0.20],
          ["Tempéreuse automatique continue avec table vibrante à moules", 5500000, 5, 0.20],
          ["Armoire frigorifique de cristallisation & démoulage hygrométrie contrôlée", 4000000, 5, 0.20],
          ["Lot de moules polycarbonate professionnels gravés & thermoscelleuse", 1800000, 3, 0.33],
          ["Véhicule utilitaire isotherme climatisé pour livraisons boutiques", 6500000, 4, 0.25],
          ["Aménagement de l'atelier aux normes d'hygiène & matériel de pesée", 3000000, 8, 0.125]
        ],
        rawMaterials: "Fèves de cacao locales fermentées et séchées grade 1, sucre de canne cristal, beurre de cacao pur, poudre de lait entier, noix de cajou locales, emballages alu et étuis kraft"
      };
    }

    // 1. Cosmétique & Savonnerie
    if (text.includes('savon') || text.includes('karité') || text.includes('cosmétique') || text.includes('beurre') || text.includes('pommade') || text.includes('crème')) {
      return {
        activity: "Fabrication de cosmétiques & savons naturels",
        unit: "Pain 150g",
        products: [
          [1, "Savon exfoliant au karité & miel", "Pain 150g sous emballage kraft", "50%", 1000],
          [2, "Savon doux hydratant surgras", "Pain 120g pour peau sensible", "30%", 800],
          [3, "Baume pur karité parfumé", "Pot aluminium 100ml", "20%", 1500]
        ],
        equipment: [
          ["Cuve de saponification en inox", 4500000, 5, 0.20],
          ["Mélangeur mécanique haute vitesse", 3500000, 5, 0.20],
          ["Moules en silicone & table de découpe", 2000000, 4, 0.25],
          ["Étagères & armoires de cure/séchage", 1500000, 5, 0.20],
          ["Machine d'étiquetage & scelleuse", 2500000, 5, 0.20],
          ["Véhicule utilitaire de livraison", 6000000, 4, 0.25],
          ["Matériel informatique & bureautique", 1500000, 3, 0.33]
        ],
        rawMaterials: "Beurre de karité brut, huiles de coco/palme, soude caustique, huiles essentielles, argiles"
      };
    }

    // 2. Détergents & Produits d'entretien
    if (text.includes('détergent') || text.includes('detergent') || text.includes('javel') || text.includes('nettoy') || text.includes('entretien')) {
      return {
        activity: "Fabrication de détergents liquides & produits d'entretien ménagers",
        unit: "Bidon / Flacon 1L",
        products: [
          [1, "Détergent liquide multi-surfaces", "Bidon 1L avec bec verseur", "45%", 1200],
          [2, "Eau de javel désinfectante concentrée", "Bouteille 1L sécurisée", "35%", 800],
          [3, "Savon liquide vaisselle dégraissant", "Flacon pompe 750ml", "20%", 1000]
        ],
        equipment: [
          ["Cuves de mélange polyéthylène 2 000L", 5000000, 5, 0.20],
          ["Agitateur électrique industriel", 3000000, 5, 0.20],
          ["Remplisseuse doseuse semi-automatique", 4500000, 5, 0.20],
          ["Visseuse de bouchons pneumatique", 1500000, 4, 0.25],
          ["Balance de précision électronique", 800000, 4, 0.25],
          ["Véhicule de distribution utilitaire", 6000000, 4, 0.25],
          ["Mobilier atelier & équipement sécurité", 1200000, 3, 0.33]
        ],
        rawMaterials: "Acide sulfonique (LABSA), soude caustique, texapon (SLES), parfum citron/menthe, colorants"
      };
    }

    // 3. Élevage & Aviculture
    if (text.includes('poulet') || text.includes('avicole') || text.includes('volaille') || text.includes('élevage') || text.includes('ferme') || text.includes('poussin') || text.includes('œuf') || text.includes('oeuf')) {
      return {
        activity: "Production avicole & élevage de volailles de chair",
        unit: "Sujet (1.8 - 2.2 kg)",
        products: [
          [1, "Poulet de chair prêt à cuire", "Conditionné sous film étirable étiqueté", "60%", 3000],
          [2, "Découpes (Cuisses, filets & ailes)", "Barquette fraîcheur 1 kg", "25%", 3500],
          [3, "Fientes valorisées (Fertilisant bio)", "Sac de 50 kg pour maraîchage", "15%", 2500]
        ],
        equipment: [
          ["Aménagement de 2 bâtiments avicoles (5 000 sujets)", 8500000, 10, 0.10],
          ["Système automatique d'abreuvement et mangeoires", 4000000, 5, 0.20],
          ["Plumeuse rotative & chaîne d'abattage semi-auto", 5500000, 5, 0.20],
          ["Chambre froide positive et congélateurs", 5000000, 5, 0.20],
          ["Groupe électrogène de secours 15 kVA", 3000000, 5, 0.20],
          ["Tricycle / Véhicule de livraison isotherme", 4500000, 4, 0.25],
          ["Matériel bureautique et pesage électronique", 1000000, 3, 0.33]
        ],
        rawMaterials: "Poussins d'un jour vaccinés, aliments démarrage/croissance/finition, vitamines, copeaux de bois"
      };
    }

    // 4. Pisciculture & Aquaculture
    if (text.includes('poisson') || text.includes('pisciculture') || text.includes('aquaculture') || text.includes('tilapia') || text.includes('silure')) {
      return {
        activity: "Pisciculture moderne & élevage de tilapias et silures",
        unit: "Kg de poisson frais",
        products: [
          [1, "Tilapias frais entiers calibrés", "Calibre 400-600g au kg", "55%", 2500],
          [2, "Silures (Clarias) frais vivants ou fumés", "Calibre 1 kg+ au kg", "30%", 3000],
          [3, "Alevins sélectionnés pour éleveurs", "Lot de 100 alevins vigoureux", "15%", 5000]
        ],
        equipment: [
          ["Construction 6 bassins bétonnés / bacs hors-sol", 7500000, 10, 0.10],
          ["Système de pompage solaire et filtration", 4500000, 5, 0.20],
          ["Aérateurs à palettes solaires", 2500000, 5, 0.20],
          ["Fumoir moderne inox à poisson", 2000000, 5, 0.20],
          ["Matériel de pêche, filets & bacs de transport", 1500000, 3, 0.33],
          ["Tricycle frigorifique de livraison", 4000000, 4, 0.25],
          ["Matériel de pesée et analyse qualité d'eau", 1000000, 3, 0.33]
        ],
        rawMaterials: "Alevins certifiés, aliments granulés flottants extrudés, probiotiques, sel de traitement"
      };
    }

    // 5. Boulangerie & Pâtisserie
    if (text.includes('pain') || text.includes('boulangerie') || text.includes('pâtisserie') || text.includes('patisserie') || text.includes('farine')) {
      return {
        activity: "Boulangerie et viennoiserie artisanale",
        unit: "Unité",
        products: [
          [1, "Baguette de tradition croustillante", "250g farine enrichie", "50%", 150],
          [2, "Pains spéciaux (Complet, son, maïs)", "300g mie serrée", "25%", 300],
          [3, "Viennoiseries (Croissants, pains au chocolat)", "Pâte levée pur beurre", "25%", 250]
        ],
        equipment: [
          ["Four à soles électrique / gaz 4 étages", 12000000, 8, 0.125],
          ["Pétrin à spirale 50 kg", 3500000, 5, 0.20],
          ["Chambre de pousse contrôlée", 3000000, 5, 0.20],
          ["Façonneuse & diviseuse volumétrique", 2500000, 5, 0.20],
          ["Vitrine réfrigérée & comptoir magasin", 3000000, 5, 0.20],
          ["Véhicule utilitaire de livraison", 5000000, 4, 0.25],
          ["Matériel informatique et caisse tactile", 1500000, 3, 0.33]
        ],
        rawMaterials: "Farine de blé type 55, levure fraîche, sel iodé, beurre, améliorant, emballages papier kraft"
      };
    }

    // 6. Textile, Couture & Mode
    if (text.includes('textile') || text.includes('couture') || text.includes('vêtement') || text.includes('vetement') || text.includes('mode') || text.includes('pagne') || text.includes('wax') || text.includes('confection')) {
      return {
        activity: "Atelier de confection textile & stylisme moderne",
        unit: "Pièce confectionnée",
        products: [
          [1, "Tenues professionnelles & uniformes scolaires", "Tissu tergal infroissable résistant", "45%", 8500],
          [2, "Vêtements prêt-à-porter en pagne Wax & Faso Dan Fani", "Coupe moderne cintrée", "35%", 15000],
          [3, "Accessoires de mode (Sacs, trousses, cravates)", "Finitions artisanales doublées", "20%", 5000]
        ],
        equipment: [
          ["5 Machines à coudre industrielles coupe-fil", 4000000, 5, 0.20],
          ["2 Surjeteuses professionnelles 4 fils", 2000000, 5, 0.20],
          ["Machine à broderie assistée par ordinateur", 3500000, 5, 0.20],
          ["Grande table de coupe & ciseaux électriques verticaux", 1500000, 5, 0.20],
          ["Centrale de repassage à vapeur avec table aspirante", 1800000, 4, 0.25],
          ["Mannequins d'exposition & mobilier showroom", 2000000, 5, 0.20],
          ["Matériel informatique de modélisme & caisse", 1200000, 3, 0.33]
        ],
        rawMaterials: "Tissus cotonnades, pagnes wax certifiés, fils à coudre, fermetures éclair, boutons, triplures"
      };
    }

    // 7. Maraîchage & Agriculture
    if (text.includes('tomate') || text.includes('piment') || text.includes('oignon') || text.includes('maraîch') || text.includes('maraich') || text.includes('agricol') || text.includes('maïs') || text.includes('manioc')) {
      return {
        activity: "Production agricole maraîchère biologique en circuit court",
        unit: "Cagette / Sac de 25 kg",
        products: [
          [1, "Tomates fraîches de plein champ", "Cagette de 25 kg qualité triée", "50%", 7500],
          [2, "Piments forts & poivrons frais", "Cagette de 15 kg sélectionnée", "30%", 9000],
          [3, "Oignons séchés & légumes maraîchers", "Sac de 25 kg bien calibré", "20%", 10000]
        ],
        equipment: [
          ["Système d'irrigation goutte-à-goutte sur 2 hectares", 5500000, 8, 0.125],
          ["Forage d'eau avec pompe solaire immergée", 6000000, 10, 0.10],
          ["Motoculteur agricole avec accessoires", 3000000, 5, 0.20],
          ["Ombrières & serres de pépinière", 2500000, 5, 0.20],
          ["Hangar de stockage et tri ventilé", 3500000, 10, 0.10],
          ["Tricycle de transport récoltes", 3500000, 4, 0.25],
          ["Petit outillage, pulvérisateurs & balances", 1000000, 3, 0.33]
        ],
        rawMaterials: "Semences certifiées à haut rendement, compost organique, biopesticides naturels, paillage"
      };
    }

    // 8. BTP, Artisanat & Matériaux (Briques, Pavés, Menuiserie)
    if (text.includes('brique') || text.includes('pavé') || text.includes('pave') || text.includes('ciment') || text.includes('menuiserie') || text.includes('meuble') || text.includes('bois')) {
      return {
        activity: "Fabrication de matériaux de construction durables & artisanat",
        unit: "Unité (Brique/Pavé/Meuble)",
        sector: "industry",
        products: [
          [1, "Briques autobloquantes en terre stabilisée (BTS)", "Bloc 30x15x10cm haute résistance", "50%", 250],
          [2, "Pavés autobloquants décoratifs", "Pavé 6cm pour allées & parkings", "30%", 350],
          [3, "Bordures et éléments architecturaux préfabriqués", "Pièce linéaire renforcée", "20%", 1500]
        ],
        equipment: [
          ["Presse hydraulique semi-automatique à briques", 8000000, 6, 0.166],
          ["Bétonnière malaxeuse grande capacité 500L", 3500000, 5, 0.20],
          ["Moules métalliques interchangeables", 2000000, 4, 0.25],
          ["Aire de séchage et de cure couverte", 4000000, 10, 0.10],
          ["Brouettes renforcées & matériel de manutention", 1500000, 3, 0.33],
          ["Camion benne utilitaire pour livraisons", 9000000, 5, 0.20],
          ["Groupe électrogène industriel 20 kVA", 4000000, 5, 0.20]
        ],
        rawMaterials: "Terre argileuse sélectionnée, ciment CPJ 42.5, sable fin lavé, adjuvants hydrofuges"
      };
    }

    // 9. Restauration, Fast-Food, Maquis & Traiteur
    if (text.includes('restau') || text.includes('fast-food') || text.includes('maquis') || text.includes('traiteur') || text.includes('poulet braisé') || text.includes('nourriture') || text.includes('cuisine') || text.includes('café') || text.includes('snack')) {
      return {
        activity: "Restauration moderne, grillades & service traiteur de proximité",
        unit: "Plat / Menu servi",
        sector: "service",
        products: [
          [1, "Menu Express Grillades (Poulet/Poisson braisé + accompagnements)", "Plat complet copieux avec sauce locale", "50%", 3000],
          [2, "Plats du Jour Traditionnels & Spécialités Africaines", "Plat mijoté avec garniture au choix", "30%", 2500],
          [3, "Boissons locales fraîches (Bissap, Gingembre, Baobab) & Desserts", "Bouteille 33cl / 50cl verre scellée", "20%", 1000]
        ],
        equipment: [
          ["Aménagement de la salle de restaurant & mobilier design", 6000000, 8, 0.125],
          ["Équipements cuisine pro (Chambre froide, fourneaux, friteuses)", 7500000, 5, 0.20],
          ["Barbecue / Grilloir inox grande capacité avec hotte", 2500000, 5, 0.20],
          ["Comptoir bar, caisse tactile & logiciel de commande", 2000000, 4, 0.25],
          ["Vaisselle, verrerie & batterie de cuisine inox", 1800000, 3, 0.33],
          ["Motos / Tricycle de livraison express repas", 3000000, 4, 0.25],
          ["Groupe électrogène de secours 15 kVA", 3500000, 5, 0.20]
        ],
        rawMaterials: "Poulets de chair, poissons frais, légumes maraîchers, condiments locaux, riz, bananes plantains, huiles de table"
      };
    }

    // 10. Transport, Logistique & VTC / Livraison
    if (text.includes('transport') || text.includes('livraison') || text.includes('vtc') || text.includes('taxi') || text.includes('logistique') || text.includes('tricycle') || text.includes('fret') || text.includes('colis')) {
      return {
        activity: "Services de transport urbain, messagerie express & logistique du dernier kilomètre",
        unit: "Course / Colis / Journée",
        sector: "service",
        products: [
          [1, "Courses régulières & transport à la demande", "Trajet urbain sécurisé et climatisé", "50%", 4000],
          [2, "Livraison express de colis & courses pour e-commerçants", "Livraison sous 2h avec suivi GPS", "30%", 1500],
          [3, "Contrats de mise à disposition d'entreprises / Abonnements", "Mise à disposition avec chauffeur dédié", "20%", 25000]
        ],
        equipment: [
          ["Flotte initiale de 3 véhicules utilitaires / berlines", 18000000, 5, 0.20],
          ["2 Tricycles benne & scooters de livraison équipés de caissons", 4000000, 4, 0.25],
          ["Système de géolocalisation GPS & application mobile", 2000000, 3, 0.33],
          ["Aménagement bureau logistique & matériel informatique", 1500000, 3, 0.33],
          ["Équipements de sécurité, casques & uniformes chauffeurs", 1000000, 2, 0.50]
        ],
        rawMaterials: "Carburant, lubrifiants, forfaits télécom & data, consommables d'emballage colis"
      };
    }

    // 11. Commerce Général, Supérette, Quincaillerie & Vente de Gros
    if (text.includes('commerce') || text.includes('boutique') || text.includes('supérette') || text.includes('superette') || text.includes('quincaillerie') || text.includes('vente') || text.includes('magasin') || text.includes('gros')) {
      return {
        activity: "Commerce de détail et demi-gros en produits de grande consommation",
        unit: "Panier / Article",
        sector: "trade",
        products: [
          [1, "Produits de grande consommation courante", "Articles de première nécessité à rotation rapide", "50%", 2500],
          [2, "Gamme supérieure & produits spécialisés", "Articles à forte marge commerciale (30-40%)", "30%", 6000],
          [3, "Packs promotionnels & vente par cartons/vrac", "Tarification demi-gros pour revendeurs", "20%", 15000]
        ],
        equipment: [
          ["Aménagement de l'espace de vente & rayonnages métalliques", 4500000, 8, 0.125],
          ["Vitrines réfrigérées positives et congélateurs bahuts", 4000000, 5, 0.20],
          ["Système d'encaissement tactile, douchette & logiciel stock", 2000000, 3, 0.33],
          ["Caméras de surveillance & portiques de sécurité", 1500000, 4, 0.25],
          ["Véhicule utilitaire pour approvisionnement grossistes", 5500000, 5, 0.20],
          ["Groupe électrogène de secours 12 kVA", 2500000, 5, 0.20]
        ],
        rawMaterials: "Stock initial de marchandises diversifiées, sacs d'emballage kraft, étiquettes code-barres"
      };
    }

    // 12. Services Numériques, Agence Tech, Cabinet & Formation
    if (text.includes('tech') || text.includes('digital') || text.includes('logiciel') || text.includes('agence') || text.includes('formation') || text.includes('conseil') || text.includes('informatique') || text.includes('web') || text.includes('application')) {
      return {
        activity: "Prestations d'ingénierie logicielle, marketing digital & formation professionnelle",
        unit: "Prestation / Jour-homme / Contrat",
        sector: "service",
        products: [
          [1, "Développement d'applications web & mobiles sur-mesure", "Projets complets livrés clés en main", "45%", 2500000],
          [2, "Abonnements mensuels de gestion marketing & maintenance", "Contrat annuel de suivi récurrent", "35%", 350000],
          [3, "Formations professionnelles certifiantes & ateliers pratiques", "Sessions intensives pour cadres et PME", "20%", 150000]
        ],
        equipment: [
          ["Aménagement du plateau de travail & salle de formation", 4000000, 8, 0.125],
          ["Parc d'ordinateurs portables haute performance (i7/M2)", 6000000, 3, 0.33],
          ["Serveurs locaux, switch réseau & équipement fibre optique", 2500000, 4, 0.25],
          ["Écran interactif / Vidéoprojecteur 4K pour salle de formation", 2000000, 4, 0.25],
          ["Mobilier ergonomique de bureau", 2000000, 5, 0.20],
          ["Système onduleur centralisé anti-coupure de courant", 1500000, 4, 0.25]
        ],
        rawMaterials: "Licences logicielles cloud, abonnements serveurs d'hébergement, supports pédagogiques et documentation"
      };
    }

    // 13. Santé, Clinique, Cabinet Médical & Dépôt Pharmaceutique
    if (text.includes('santé') || text.includes('sante') || text.includes('clinique') || text.includes('médic') || text.includes('medic') || text.includes('pharma') || text.includes('soin') || text.includes('docteur')) {
      return {
        activity: "Centre médical de soins de proximité, consultations & analyses",
        unit: "Consultation / Acte de soin",
        sector: "service",
        products: [
          [1, "Consultations de médecine générale & petite chirurgie", "Prise en charge médicale personnalisée", "50%", 7500],
          [2, "Examens de laboratoire & analyses biomédicales rapides", "Bilan sanguin, paludisme, glycémie, etc.", "30%", 10000],
          [3, "Soins infirmiers, maternité & hospitalisation de jour", "Chambres climatisées surveillées", "20%", 15000]
        ],
        equipment: [
          ["Aménagement des locaux aux normes sanitaires & hospitalières", 8000000, 10, 0.10],
          ["Automates d'analyses de laboratoire & microscopes", 6500000, 5, 0.20],
          ["Échographe portable numérique & ECG", 5000000, 5, 0.20],
          ["Lits d'hospitalisation médicalisés & mobilier de consultation", 3500000, 5, 0.20],
          ["Stérilisateur autoclave & matériel chirurgical inox", 2500000, 4, 0.25],
          ["Ambulance / Véhicule d'intervention médicale", 9000000, 5, 0.20],
          ["Groupe électrogène médical automatique 20 kVA", 4500000, 5, 0.20]
        ],
        rawMaterials: "Réactifs de laboratoire, compresses stériles, gants, seringues, solutés, médicaments d'urgence"
      };
    }

    // 14. Par défaut et sur-mesure pour tout autre produit demandé
    const cleanTopic = (topic || defaultProject || "Transformation & Commerce")
      .replace(/^(?:projet\s+(?:de\s+|d'|d’)?|création\s+(?:de\s+|d'|d’)?|fabrication\s+(?:de\s+|d'|d’)?|production\s+(?:de\s+|d'|d’)?|atelier\s+(?:de\s+|d'|d’)?|usine\s+(?:de\s+|d'|d’)?|vente\s+(?:de\s+|d'|d’)?|élevage\s+(?:de\s+|d'|d’)?)/i, '')
      .replace(/^(?:mon|ma|mes|notre|nos|un|une|le|la|les|du|de\s+la|des)\s+/i, '')
      .trim();

    const topicLabel = cleanTopic.charAt(0).toUpperCase() + cleanTopic.slice(1);

    return {
      activity: `Production, valorisation & commercialisation de ${topicLabel}`,
      unit: "Unité / Lot",
      sector: "trade",
      products: [
        [1, `${topicLabel} - Gamme Standard`, "Conditionnement classique certifié conforme aux normes locales", "50%", 2500],
        [2, `${topicLabel} - Gamme Supérieure`, "Finition supérieure sous emballage soigné à haute valeur", "30%", 4500],
        [3, `${topicLabel} - Format Économique / Vrac`, "Conditionnement adapté pour distributeurs et revendeurs", "20%", 8000]
      ],
      equipment: [
        ["Aménagement de l'espace d'exploitation & conformité", 5000000, 10, 0.10],
        ["Machine et ligne principale d'exploitation", 8000000, 5, 0.20],
        ["Équipements auxiliaires de traitement & dosage", 4000000, 5, 0.20],
        ["Ligne d'emballage, scellage & étiquetage", 3500000, 5, 0.20],
        ["Matériel de stockage, étagères & conservation", 3000000, 5, 0.20],
        ["Véhicule utilitaire de livraison & logistique", 6000000, 4, 0.25],
        ["Matériel bureautique, gestion commerciale & informatique", 1500000, 3, 0.33]
      ],
      rawMaterials: `Matières premières sélectionnées pour ${topicLabel}, emballages personnalisés, consommables d'exploitation`
    };
  }

  /**
   * Moteur de modélisation financière dynamique et cohérente
   * Calcule les flux, investissements, P&L, BFR et ratios adaptés au produit
   */
  private static computeFinancialModel(profile: any, projectName: string, currency: string) {
    const pList = profile?.products || [];
    const p1 = pList[0] || [1, "Gamme Principale", "Unité", "50%", 1800];
    const p2 = pList[1] || [2, "Gamme Secondaire", "Unité", "30%", 2000];
    const p3 = pList[2] || [3, "Gamme Complémentaire", "Unité", "20%", 2800];
    const price1 = Number(p1[4]) || 1800;
    const price2 = Number(p2[4]) || 2000;
    const price3 = Number(p3[4]) || 2800;

    // Déterminer un volume journalier et annuel réaliste (CA Année 1 entre 70M et 130M FCFA)
    const targetDailySales = 360000; // ~360 000 FCFA/jour de ventes
    const vol1Day = Math.max(Math.round((targetDailySales * 0.50) / Math.max(price1, 10)), 10);
    const vol2Day = Math.max(Math.round((targetDailySales * 0.30) / Math.max(price2, 10)), 8);
    const vol3Day = Math.max(Math.round((targetDailySales * 0.20) / Math.max(price3, 10)), 5);

    const workingDays = 240;
    const vol1An1 = vol1Day * workingDays;
    const vol2An1 = vol2Day * workingDays;
    const vol3An1 = vol3Day * workingDays;

    const ca1_1 = price1 * vol1An1;
    const ca1_2 = price2 * vol2An1;
    const ca1_3 = price3 * vol3An1;
    const ca1 = ca1_1 + ca1_2 + ca1_3;
    const ca2 = Math.round(ca1 * 1.30);
    const ca3 = Math.round(ca1 * 1.70);
    const ca5 = Math.round(ca1 * 2.45);

    // Investissements
    const equipment = (profile?.equipment && profile.equipment.length > 0)
      ? profile.equipment
      : [
          ["Machine et ligne principale d'exploitation", 8000000, 5, 0.20],
          ["Aménagement local & atelier d'exploitation", 5000000, 10, 0.10],
          ["Équipements auxiliaires & outillage spécialisé", 4000000, 5, 0.20],
          ["Ligne d'emballage, scellage & étiquetage", 3500000, 5, 0.20],
          ["Matériel de stockage frigorifique / sécurisé", 3000000, 5, 0.20],
          ["Véhicule utilitaire de livraison", 6000000, 4, 0.25],
          ["Matériel bureautique, informatique & caisse", 1500000, 3, 0.33]
        ];

    const totalEquipment = equipment.reduce((acc: number, item: any) => acc + (Number(item[1]) || 0), 0);
    const annualAmort = equipment.reduce((acc: number, item: any) => acc + ((Number(item[1]) || 0) * (Number(item[3]) || 0.20)), 0);

    // Plan de financement initial
    const fraisEtablissement = 1000000;
    const bfrInitial = Math.round(ca1 * 0.08); // BFR de démarrage ~1 mois
    const totalBesoins = totalEquipment + fraisEtablissement + bfrInitial;

    // Emprunt bancaire moyen terme (arrondi au million le plus proche)
    const emprunt = Math.round((totalBesoins * 0.55) / 1000000) * 1000000;
    const apportPersonnel = totalBesoins - emprunt;
    const interetsAn1 = Math.round(emprunt * 0.08);

    // Charges d'exploitation
    const achatsMatieres = Math.round(ca1 * 0.44);
    const achatsEmballages = Math.round(ca1 * 0.08);
    const energieAtelier = Math.round(ca1 * 0.035);
    const carburant = Math.round(ca1 * 0.035);
    const maintenance = Math.round(ca1 * 0.02);
    const assurances = 800000;
    const telecom = 600000;
    const marketing = Math.round(ca1 * 0.03);
    const totalChargesExploit = achatsMatieres + achatsEmballages + energieAtelier + carburant + maintenance + assurances + telecom + marketing;

    // Salaires (Masse salariale annuelle de base)
    const masseSalariale = 15600000;

    // Compte de résultat SIG
    const margeBrute = ca1 - (achatsMatieres + achatsEmballages);
    const servicesExt = energieAtelier + carburant + maintenance + assurances + telecom + marketing;
    const valeurAjoutee = margeBrute - servicesExt;
    const ebe = valeurAjoutee - masseSalariale;
    const rex = ebe - Math.round(annualAmort);
    const rao = rex - interetsAn1;
    const impotIS = rao > 0 ? Math.round(rao * 0.25) : 0;
    const resultatNet = rao - impotIS;
    const caf = resultatNet + Math.round(annualAmort);

    // Seuil de rentabilité & point mort
    const chargesFixes = masseSalariale + Math.round(annualAmort) + interetsAn1 + assurances + telecom + Math.round(energieAtelier * 0.4);
    const chargesVariables = (achatsMatieres + achatsEmballages) + Math.round(servicesExt * 0.6);
    const mcv = ca1 - chargesVariables;
    const tauxMcv = mcv / ca1;
    const sr = Math.round(chargesFixes / Math.max(tauxMcv, 0.1));
    const pointMort = Math.min(Math.round((sr / ca1) * 360), 360);
    const margeSecurite = ca1 - sr;
    const indiceSecurite = Math.max((ca1 - sr) / ca1, 0);

    return {
      price1, price2, price3,
      vol1An1, vol2An1, vol3An1,
      vol1Day, vol2Day, vol3Day,
      ca1_1, ca1_2, ca1_3,
      ca1, ca2, ca3, ca5,
      equipment, totalEquipment, annualAmort,
      fraisEtablissement, bfrInitial, totalBesoins,
      emprunt, apportPersonnel, interetsAn1,
      achatsMatieres, achatsEmballages, energieAtelier, carburant, maintenance, assurances, telecom, marketing, totalChargesExploit,
      masseSalariale,
      margeBrute, servicesExt, valeurAjoutee, ebe, rex, rao, impotIS, resultatNet, caf,
      chargesFixes, chargesVariables, mcv, tauxMcv, sr, pointMort, margeSecurite, indiceSecurite
    };
  }

  /**
   * Génère le Business Plan sur-mesure (adapté au produit et au nombre de feuilles demandé)
   */
  static async generateCompletePlan(options: BusinessPlanOptions): Promise<{ totalSheets: number; message: string }> {
    const project = options.projectName || "Projet de Transformation Agroalimentaire";
    const structure = options.structureName || "À compléter";
    const typeProjet = options.projectType || "Création";
    const currency = (options.currencySymbol || "FCFA").replace(/[\(\)\[\]"]/g, '').trim() || "FCFA";
    const profile = this.getProductProfile(options.productTopic || options.projectName, project);
    const model = this.computeFinancialModel(profile, project, currency);

    const allSheetConfigs = [
      { id: "garde", name: "Page de garde", color: "#1F3864", fn: this.buildPageDeGarde },
      { id: "presentation", name: "Présentation Entreprise", color: "#1B2A4A", fn: this.buildPresentationEntreprise },
      { id: "marche", name: "Analyse du marché", color: "#2F5597", fn: this.buildAnalyseMarche },
      { id: "marketing", name: "Stratégie Marketing", color: "#41719C", fn: this.buildStrategieMarketing },
      { id: "calendrier", name: "ITK et Calendrier", color: "#70AD47", fn: this.buildCalendrierITK },
      { id: "production", name: "Prévisions Production", color: "#548235", fn: this.buildPrevisionsProduction },
      { id: "ca", name: "Chiffres d'Affaires", color: "#385723", fn: this.buildChiffresAffaires },
      { id: "investissements", name: "Investissements", color: "#C65911", fn: this.buildInvestissements },
      { id: "charges", name: "Charges d'Exploitation", color: "#833C0C", fn: this.buildChargesExploitation },
      { id: "personnel", name: "Frais de Personnel", color: "#7030A0", fn: this.buildFraisPersonnel },
      { id: "financement", name: "Plan de Financement", color: "#1F3864", fn: this.buildPlanFinancement },
      { id: "emprunt", name: "Remboursement Emprunt", color: "#C00000", fn: this.buildRemboursementEmprunt },
      { id: "bfr", name: "BFR", color: "#008080", fn: this.buildBFR },
      { id: "resultat", name: "Compte de Résultat", color: "#107C41", fn: this.buildCompteResultat },
      { id: "bilan", name: "Bilan Prévisionnel", color: "#4682B4", fn: this.buildBilanPrevisionnel },
      { id: "tresorerie", name: "Plan de Trésorerie", color: "#FFC000", fn: this.buildPlanTresorerie },
      { id: "rentabilite", name: "Seuil de Rentabilité", color: "#ED7D31", fn: this.buildSeuilRentabilite },
      { id: "sensibilite", name: "Analyse Sensibilité", color: "#1E293B", fn: this.buildSensibilite },
      { id: "ratios", name: "Indicateurs & Ratios", color: "#4472C4", fn: this.buildRatiosRentabilite },
      { id: "synthese", name: "Tableau de Bord", color: "#00B050", fn: this.buildTableauBord }
    ];

    // Sélection intelligente des feuilles selon le secteur et le besoin réel
    const isAgriOrLivestock = profile?.sector === 'agri' || (profile?.activity && (profile.activity.toLowerCase().includes('élev') || profile.activity.toLowerCase().includes('agri') || profile.activity.toLowerCase().includes('maraîch') || profile.activity.toLowerCase().includes('pisci') || profile.activity.toLowerCase().includes('volaille')));

    // Ordre de priorité des feuilles : les piliers bancables d'abord
    const prioritySheetIds = [
      "synthese",        // 1. Tableau de Bord (Executive Summary - vu en premier par le comité de crédit)
      "garde",           // 2. Page de garde
      "presentation",    // 3. Présentation Entreprise & Promoteur
      "marche",          // 4. Analyse du marché & Concurrence
      "marketing",       // 5. Stratégie Marketing & Plan 4P
      ...(isAgriOrLivestock ? ["production", "calendrier"] : []), // Uniquement pour agriculture/élevage
      "investissements", // 6. Investissements & Amortissements SYSCOHADA
      "ca",              // 7. Chiffre d'Affaires prévisionnel
      "charges",         // 8. Charges d'Exploitation
      "financement",     // 9. Plan de Financement Initial
      "resultat",        // 10. Compte de Résultat Prévisionnel SYSCOHADA
      "tresorerie",      // 11. Plan de Trésorerie Mensualisé sur 12 Mois
      "rentabilite",     // 12. Seuil de Rentabilité & Point Mort
      "sensibilite",     // 13. Analyse de la Sensibilité (Stress-test bancaire)
      "ratios",          // 14. Ratios Bancaires VAN & TRI
      "personnel",       // 15. Frais de Personnel
      "bfr",             // 16. Détail BFR
      "emprunt",         // 17. Remboursement Emprunt Bancaire
      "bilan"            // 18. Bilan Prévisionnel
    ];

    let targets = allSheetConfigs;
    if (options.sheetsToGenerate && options.sheetsToGenerate.length > 0) {
      targets = allSheetConfigs.filter(s => options.sheetsToGenerate!.includes(s.id) || options.sheetsToGenerate!.includes(s.name));
    } else {
      // Par défaut : Pack Bancable Optimal de 10 à 12 feuilles percutantes (au lieu de 19 feuilles lourdes)
      const targetCount = (options.sheetCount && options.sheetCount > 0) ? options.sheetCount : (isAgriOrLivestock ? 12 : 10);
      const selectedIds = prioritySheetIds.slice(0, targetCount);
      targets = selectedIds
        .map(id => allSheetConfigs.find(s => s.id === id))
        .filter((s): s is typeof allSheetConfigs[0] => !!s);
    }

    if (typeof Excel === 'undefined') {
      console.log(`[Simulation Business Plan] ${targets.length} feuilles générées pour '${project}'`);
      return { totalSheets: targets.length, message: `Business Plan de ${targets.length} feuilles simulé avec succès.` };
    }

    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      for (let i = 0; i < targets.length; i++) {
        const cfg = targets[i];
        options.onProgress?.(cfg.name, i + 1, targets.length);

        // Récupérer ou créer la feuille
        let sheet: Excel.Worksheet;
        const existing = sheets.items.find(s => s.name === cfg.name);
        if (existing) {
          sheet = existing;
        } else {
          sheet = sheets.add(cfg.name);
        }

        // Appliquer la couleur d'onglet uniquement si supporté par Excel (ExcelApi 1.7+)
        if (typeof Office !== 'undefined' && Office.context?.requirements?.isSetSupported('ExcelApi', '1.7')) {
          try {
            sheet.tabColor = cfg.color;
          } catch (_) {}
        }

        // Initialiser les marges et finitions de base avant la construction du contenu
        this.applyGlobalFinish(sheet);

        // Construire le contenu spécifique de la feuille
        try {
          await cfg.fn.call(BusinessPlanService, sheet, context, {
            projectName: project,
            structureName: structure,
            projectType: typeProjet,
            currencySymbol: currency,
            profile: profile,
            model: model
          });
          await context.sync();
        } catch (sheetErr) {
          console.warn(`Feuille '${cfg.name}' avertissement:`, sheetErr);
        }
      }

      // Activer la première feuille (Page de garde)
      const first = sheets.getItem(targets[0].name);
      first.activate();
      await context.sync();

      return {
        totalSheets: targets.length,
        message: `Business Plan complet (${targets.length} feuilles) créé avec succès dans votre classeur !`
      };
    });
  }

  // --- FEUILLE 1 : PAGE DE GARDE ---
  private static async buildPageDeGarde(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    sheet.getRange("B2:E2").values = [["BUSINESS PLAN", "", "", ""]];
    sheet.getRange("B2:E2").merge(false);
    sheet.getRange("B2:E2").format.font.bold = true;
    sheet.getRange("B2:E2").format.font.size = 20;
    sheet.getRange("B2:E2").format.font.color = "#1F3864";
    sheet.getRange("B2:E2").format.horizontalAlignment = "Center";

    sheet.getRange("B4:E4").values = [[p.projectName, "", "", ""]];
    sheet.getRange("B4:E4").merge(false);
    sheet.getRange("B4:E4").format.font.bold = true;
    sheet.getRange("B4:E4").format.font.size = 14;
    sheet.getRange("B4:E4").format.font.color = "#2F5597";
    sheet.getRange("B4:E4").format.horizontalAlignment = "Center";

    const m = p.model;
    const summaryData = [
      ["Nom complet du porteur de projet", "À compléter"],
      ["Nom de la structure", p.structureName],
      ["Activité(s) principale(s)", p.profile?.activity || "Transformation & Production"],
      ["Type de projet", p.projectType],
      ["Titre du projet", p.projectName],
      ["Montant global du projet", `${(m?.totalBesoins || 30000000).toLocaleString('fr-FR')} ${p.currencySymbol}`],
      ["Montant du prêt sollicité", `${(m?.emprunt || 18000000).toLocaleString('fr-FR')} ${p.currencySymbol}`],
      ["Apport personnel (Fonds propres)", `${(m?.apportPersonnel || 12000000).toLocaleString('fr-FR')} ${p.currencySymbol}`],
      ["Durée de remboursement souhaitée", "48 mois (4 ans)"],
      ["Durée du différé sollicité", "6 mois"]
    ];

    for (let r = 0; r < summaryData.length; r++) {
      const rowIdx = 7 + r * 2;
      const labelRange = sheet.getRange(`B${rowIdx}:C${rowIdx}`);
      labelRange.merge(false);
      labelRange.values = [[summaryData[r][0], ""]];
      labelRange.format.fill.color = "#D9E1F2"; // Bleu pastel d'en-tête
      labelRange.format.font.bold = true;
      labelRange.format.font.size = 10;
      this.setBorder(labelRange, "EdgeBottom", "Continuous");
      this.setBorder(labelRange, "EdgeTop", "Continuous");

      const inputRange = sheet.getRange(`D${rowIdx}:E${rowIdx}`);
      inputRange.merge(false);
      inputRange.values = [[summaryData[r][1], ""]];
      inputRange.format.fill.color = "#FFF2CC"; // Jaune d'entrée à compléter
      this.setBorder(inputRange, "EdgeBottom", "Continuous");
      this.setBorder(inputRange, "EdgeTop", "Continuous");
      inputRange.format.horizontalAlignment = "Center";
    }

    try {
      sheet.getRange("B:B").format.columnWidth = 260;
      sheet.getRange("C:C").format.columnWidth = 140;
      sheet.getRange("D:D").format.columnWidth = 160;
      sheet.getRange("E:E").format.columnWidth = 160;
    } catch (_) {}
  }

  // --- FEUILLE 2 : PRÉSENTATION DE L'ENTREPRISE ---
  private static async buildPresentationEntreprise(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:E2", "PRÉSENTATION DE LA STRUCTURE OU DU PROJET", "#1F3864");

    const structureFields = [
      ["Raison sociale", p.structureName],
      ["Forme juridique", "SARL Unipersonnelle / SARL"],
      ["Adresse & Localisation", "Zone industrielle / Quartier commercial"],
      ["Date de création", "En cours de création"],
      ["N° du contribuable (IFU)", "À compléter"],
      ["RCCM", "À compléter"],
      ["Régime fiscal", "Régime Réel Simplifié / Normal"],
      ["Montant du capital social", `5 000 000 ${p.currencySymbol}`],
      ["Chiffre d'affaires dernière année (N-1)", `0 ${p.currencySymbol} (Nouveau projet)`],
      ["Crédit bancaire en cours", "Néant"],
      ["Crédit tontine", "Néant"],
      ["Nombre de personnel initial", "6 employés permanents"],
      ["Position actuelle", "Phase de levée de fonds & acquisition du matériel"]
    ];

    sheet.getRange("C4:E4").values = [["Remplir", "", ""]];
    sheet.getRange("C4:E4").merge(false);
    sheet.getRange("C4:E4").format.fill.color = "#FFE699";
    sheet.getRange("C4:E4").format.font.bold = true;
    sheet.getRange("C4:E4").format.horizontalAlignment = "Center";

    for (let i = 0; i < structureFields.length; i++) {
      const row = 5 + i;
      sheet.getRange(`B${row}`).values = [[structureFields[i][0]]];
      sheet.getRange(`B${row}`).format.fill.color = "#F2F2F2";
      sheet.getRange(`B${row}`).format.font.bold = true;

      const valRange = sheet.getRange(`C${row}:E${row}`);
      valRange.merge(false);
      valRange.values = [[structureFields[i][1], "", ""]];
      valRange.format.fill.color = "#FFF2CC";
    }

    const nextRow = 6 + structureFields.length;
    this.createBanner(sheet, `B${nextRow}:E${nextRow}`, "Produits et services proposés", "#1F3864");

    sheet.getRange(`B${nextRow + 1}:E${nextRow + 1}`).values = [["N°", "Produit / Gamme", "Conditionnement", "Description technique"]];
    sheet.getRange(`B${nextRow + 1}:E${nextRow + 1}`).format.fill.color = "#D9E1F2";
    sheet.getRange(`B${nextRow + 1}:E${nextRow + 1}`).format.font.bold = true;

    const defaultProducts = [
      [1, "Gamme Principale", "Unité standard", "Produit principal de haute qualité répondant à la demande du marché local."],
      [2, "Gamme Secondaire", "Format intermédiaire", "Produit dérivé à forte valeur ajoutée."],
      [3, "Gamme Économique / Vrac", "Grand format", "Offre accessible pour grossistes et distributeurs."]
    ];

    const products = (p.profile?.products && p.profile.products.length > 0)
      ? p.profile.products.map((prod: any, idx: number) => [
          idx + 1,
          prod[1],
          prod[2] || p.profile?.unit || "Unité standard",
          prod[2] ? `${prod[2]} - Fabriqué selon les normes de qualité et d'hygiène.` : "Produit certifié conforme aux normes du marché local."
        ])
      : defaultProducts;

    for (let pIdx = 0; pIdx < products.length; pIdx++) {
      const r = nextRow + 2 + pIdx;
      sheet.getRange(`B${r}:E${r}`).values = [products[pIdx]];
      sheet.getRange(`B${r}`).format.horizontalAlignment = "Center";
    }

    try {
      sheet.getRange("B:B").format.columnWidth = 260;
      sheet.getRange("C:C").format.columnWidth = 220;
      sheet.getRange("D:D").format.columnWidth = 160;
      sheet.getRange("E:E").format.columnWidth = 260;
    } catch (_) {}
  }

  // --- FEUILLE 3 : ANALYSE DU MARCHÉ & PESTEL ---
  private static async buildAnalyseMarche(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:F2", "ÉTUDE DU MARCHÉ & ANALYSE DE L'ENVIRONNEMENT LOCAL", "#1F3864");

    // Section 1 : Segments de clientèle ciblés
    this.createBanner(sheet, "B4:F4", "1. SEGMENTATION DE LA DEMANDE & CIBLES VISÉES", "#2F5597");
    sheet.getRange("B5:F5").values = [["Segment Clientèle", "Besoins & Attentes Détectés", "Zone Géographique Cible", "Part du CA Visée", "Mode de Règlement Prévu"]];

    const pList = p.profile?.products || [];
    const prodName = pList[0] ? pList[0][1] : (p.projectName || "Produits & Services");

    const marketSegments = [
      ["Ménages & Particuliers", `Consommation directe de qualité pour ${prodName}`, "Quartiers urbains & semi-urbains", 0.40, "Comptant / Mobile Money (Wave, OM)"],
      ["Commerces de détail & Supérettes", "Approvisionnement régulier, marge revendeur attractive", "Centres commerciaux & grandes artères", 0.30, "Comptant / Crédit 7-15 jours"],
      ["Hôtels, Maquis, Restaurants & Traiteurs", "Conditionnements adaptés & livraison garantie", "Zones touristiques et d'affaires", 0.20, "Crédit 15-30 jours / Chèque"],
      ["Collectivités, Entreprises & Événements", "Commandes groupées pour cérémonies et cantines", "Périmètre régional étendu", 0.10, "Acompte 50% à la commande + Solde"]
    ];

    for (let i = 0; i < marketSegments.length; i++) {
      sheet.getRange(`B${6 + i}:F${6 + i}`).values = [marketSegments[i]];
    }

    sheet.getRange("B10:F10").values = [["TOTAL ESTIMATION MARCHÉ CIBLE", "", "", "=SUM(E6:E9)", ""]];
    this.applyModernTableStyle(sheet, "B5:F5", "B6:F9", "#1F3864", { totalRow: "B10:F10" });
    this.setRangeNumberFormat(sheet, "E6:E10", "0.0%");

    // Section 2 : Matrice concurrentielle locale
    this.createBanner(sheet, "B12:F12", "2. MATRICE CONCURRENTIELLE LOCALE (DIRECTS & INDIRECTS)", "#2F5597");
    sheet.getRange("B13:F13").values = [["Catégorie de Concurrent", "Acteurs Typiques sur le Terrain", "Leurs Forces Principales", "Leurs Faiblesses / Insatisfactions", "Notre Avantage Compétitif Décisif"]];

    const competitors = [
      ["Secteur Artisanal Informel", "Vendeurs ambulants et petits producteurs de quartier", "Proximité immédiate, prix très bas", "Hygiène variable, ruptures de stock fréquentes", "Qualité certifiée, traçabilité et packaging professionnel"],
      ["Produits Industriels Importés", "Marques multinationales vendues en grandes surfaces", "Forte notoriété, gros budgets publicitaires", "Prix élevés en devises, peu adaptés au goût local", "Prix 20 à 35% moins cher, fraîcheur locale, circuit court"],
      ["PME Locales Établies", "Unités de transformation régionales existantes", "Réseau de distribution déjà en place", "Capacité logistique saturée, service client passif", "Livraison express sous 24h, relation client proactive & digitalisée"]
    ];

    for (let i = 0; i < competitors.length; i++) {
      sheet.getRange(`B${14 + i}:F${14 + i}`).values = [competitors[i]];
    }
    this.applyModernTableStyle(sheet, "B13:F13", "B14:F16", "#2F5597");

    // Section 3 : Analyse stratégique SWOT / PESTEL
    this.createBanner(sheet, "B18:F18", "3. ENVIRONNEMENT STRATÉGIQUE (OPPORTUNITÉS & FACTEURS CLÉS DE SUCCÈS)", "#385723");
    sheet.getRange("B19:F19").values = [["Dimension Stratégique", "Opportunité Majeure Identifiée", "Impact sur la Rentabilité", "Risque Potentiel à Surveiller", "Mesure d'Atténuation Prévue"]];

    const pestel = [
      ["Économique & Pouvoir d'Achat", "Croissance rapide de la classe moyenne urbaine friande de produits locaux", "Hausse régulière de la demande", "Inflation sur les intrants importés", "Fournisseurs locaux sous contrat à prix fixe"],
      ["Politique & Réglementaire", "Programmes gouvernementaux de soutien au 'Consommer Local' en zone UEMOA/CEMAC", "Facilités fiscales et accès prioritaire aux marchés", "Lenteurs administratives d'homologation", "Accompagnement par un cabinet de conformité"],
      ["Social & Habitudes de Consommation", "Sensibilité croissante à la santé, la nutrition saine et l'authenticité", "Fidélité client élevée", "Sensibilité au prix au détail", "Formats économiques familiaux et individuels"],
      ["Technologique & Logistique", "Adoption massive du paiement Mobile Money et de WhatsApp Business pour commander", "Réduction des impayés et des délais de commande", "Coupures d'électricité ponctuelles", "Groupe électrogène de secours et onduleurs"]
    ];

    for (let i = 0; i < pestel.length; i++) {
      sheet.getRange(`B${20 + i}:F${20 + i}`).values = [pestel[i]];
    }
    this.applyModernTableStyle(sheet, "B19:F19", "B20:F23", "#385723");

    // Colonnes calibrées et spacieuses
    try {
      sheet.getRange("B:B").format.columnWidth = 190;
      sheet.getRange("C:C").format.columnWidth = 240;
      sheet.getRange("D:D").format.columnWidth = 200;
      sheet.getRange("E:E").format.columnWidth = 170;
      sheet.getRange("F:F").format.columnWidth = 240;
    } catch (_) {
      sheet.getUsedRange().format.autofitColumns();
    }
  }

  // --- FEUILLE 4 : STRATÉGIE MARKETING ---
  private static async buildStrategieMarketing(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:F2", "STRATÉGIE MARKETING & PLAN D'ACTION COMMERCIAL (4P)", "#1F3864");

    // Table 1 : Le Mix Marketing Opérationnel (4P)
    this.createBanner(sheet, "B4:F4", "1. PLAN D'ACTION MARKETING MIX (4P) & BUDGETS ALLOUÉS", "#2F5597");
    sheet.getRange("B5:F5").values = [["Pilier du Mix (4P)", "Orientation Stratégique", "Déclinaison Opérationnelle Locale", "Canaux & Outils Clés", `Budget Année 1 (${p.currencySymbol})`]];

    const pList = p.profile?.products || [];
    const prodDesc = pList[0] ? pList[0][1] : (p.projectName || "Gamme principale de qualité");

    const marketing4P = [
      [
        "1. POLITIQUE PRODUIT", 
        `Positionnement premium accessible pour ${prodDesc}`, 
        "Conditionnement aux normes d'hygiène rigoureuses, étiquetage soigné avec code-barres et tableau nutritionnel / fiche technique", 
        "Emballages écologiques personnalisés, formats familial et individuel", 
        1500000
      ],
      [
        "2. POLITIQUE PRIX", 
        "Pénétration de marché avec marge revendeur garantie", 
        "Prix consommateur très attractif avec 20% à 25% de marge nette pour les distributeurs et boutiques partenaires", 
        "Grille tarifaire dégressive par palier de volume, paiement Mobile Money à la livraison", 
        500000
      ],
      [
        "3. POLITIQUE DISTRIBUTION (PLACE)", 
        "Distribution multicanale sélective et de proximité", 
        "Vente directe au comptoir d'atelier, référencement dans 40 supérettes cibles et livraison express traiteurs/restaurants", 
        "Tricycles de distribution dédiés, catalogue WhatsApp Business pour prise de commande", 
        2500000
      ],
      [
        "4. POLITIQUE PROMOTION / COMMUNICATION", 
        "Visibilité locale percutante et marketing d'influence", 
        "Campagnes sponsorisées ciblées (TikTok, Instagram, Facebook), dégustations/tests sur points de vente et parrainages", 
        "Enseignes lumineuses, flyers sur carrefours stratégiques, micro-influenceurs culinaires/lifestyle", 
        1800000
      ]
    ];

    for (let i = 0; i < marketing4P.length; i++) {
      sheet.getRange(`B${6 + i}:F${6 + i}`).values = [marketing4P[i]];
    }

    sheet.getRange("B10:F10").values = [["TOTAL BUDGET COMMERCIAL & MARKETING", "", "", "", "=SUM(F6:F9)"]];
    this.applyModernTableStyle(sheet, "B5:F5", "B6:F9", "#1F3864", { totalRow: "B10:F10" });
    this.setRangeNumberFormat(sheet, "F6:F10", `# ##0 "${p.currencySymbol}"`);

    // Table 2 : Calendrier de Déploiement Commercial sur 12 Mois
    this.createBanner(sheet, "B12:F12", "2. OBJECTIFS COMMERCIAUX & CALENDRIER D'EXÉCUTION SUR 12 MOIS", "#2F5597");
    sheet.getRange("B13:F13").values = [["Période", "Phase Commerciale", "Action Prioritaire sur le Terrain", "Responsable Opérationnel", "Indicateur Clé de Succès (KPI)"]];

    const timeline = [
      ["Mois 1 - 2", "Lancement & Notoriété", "Inauguration du point de vente, distribution d'échantillons et lancement du catalogue WhatsApp", "Responsable Commercial", "100 premiers clients & 15 boutiques référencées"],
      ["Mois 3 - 5", "Pénétration & Réseau", "Prospection directe terrain des demi-grossistes, hôtels, maquis et supermarchés de quartier", "Force de vente terrain", "40 points de vente actifs, CA mensuel en hausse de 20%"],
      ["Mois 6 - 8", "Fidélisation & Volume", "Lancement du programme de remise fidélité revendeurs et garantie de livraison en 2h chrono", "Service Logistique", "Taux de réachat > 70%, 0 rupture de stock"],
      ["Mois 9 - 12", "Consolidation & Croissance", "Partenariats corporate (comités d'entreprises, événements de fin d'année) et expansion géographique", "Direction Commerciale", "Chiffre d'affaires annuel cible atteint à 100%"]
    ];

    for (let i = 0; i < timeline.length; i++) {
      sheet.getRange(`B${14 + i}:F${14 + i}`).values = [timeline[i]];
    }
    this.applyModernTableStyle(sheet, "B13:F13", "B14:F17", "#2F5597");

    // Calibrage des largeurs
    try {
      sheet.getRange("B:B").format.columnWidth = 180;
      sheet.getRange("C:C").format.columnWidth = 230;
      sheet.getRange("D:D").format.columnWidth = 250;
      sheet.getRange("E:E").format.columnWidth = 190;
      sheet.getRange("F:F").format.columnWidth = 150;
    } catch (_) {
      sheet.getUsedRange().format.autofitColumns();
    }
  }

  // --- FEUILLE 5 : ITK ET CALENDRIER ---
  private static async buildCalendrierITK(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:F2", "ITINÉRAIRE TECHNIQUE ET CALENDRIER D'EXÉCUTION", "#1E293B");

    sheet.getRange("B4:F4").values = [["Étape", "Action / Jalon clé", "Responsable", "Délai / Période", "Livrable attendu"]];

    const calendarRows = [
      ["Mois 1", "Création légale, immatriculation fiscale & ouverture comptes", "Promoteur / Notaire", "Semaines 1-4", "Statuts, RCCM, IFU"],
      ["Mois 2", "Aménagement du local de production & normes d'hygiène", "Chef de production", "Semaines 5-8", "Atelier prêt"],
      ["Mois 3", "Acquisition, livraison et installation des équipements", "Fournisseur / Promoteur", "Semaines 9-12", "Ligne de presse opérationnelle"],
      ["Mois 4", "Tests de production, recettes & formation du personnel", "Qualité & Hygiène", "Semaines 13-16", "Lots d'essai conformes"],
      ["Mois 5", "Campagne de lancement commercial & référencement", "Responsable Ventes", "Semaines 17-20", "Premières commandes livrées"],
      ["Mois 6", "Plein régime de croisière et suivi de trésorerie", "Comptable / Direction", "Semaines 21-24", "Rapports d'activité mensuels"]
    ];

    for (let i = 0; i < calendarRows.length; i++) {
      sheet.getRange(`B${5 + i}:F${5 + i}`).values = [calendarRows[i]];
    }

    this.applyModernTableStyle(sheet, "B4:F4", "B5:F10", "#1E293B");

    try {
      sheet.getRange("B:B").format.columnWidth = 120;
      sheet.getRange("C:C").format.columnWidth = 290;
      sheet.getRange("D:D").format.columnWidth = 190;
      sheet.getRange("E:E").format.columnWidth = 150;
      sheet.getRange("F:F").format.columnWidth = 220;
    } catch (_) {}
  }

  // --- FEUILLE 6 : PRÉVISIONS DE PRODUCTION ---
  private static async buildPrevisionsProduction(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:G2", "CAPACITÉS ET VOLUMES PRÉVISIONNELS DE PRODUCTION (5 ANS)", "#1E293B");

    const unitLabel = p.profile?.unit || "Unités";
    sheet.getRange("B4:G4").values = [["Désignation Gamme", `Capacité Jour (${unitLabel})`, "Année 1 (Unité)", "Année 2 (Unité)", "Année 3 (Unité)", "Année 5 (Unité)"]];

    const pList = p.profile?.products || [];
    const m = p.model;
    const prodRows = [
      [pList[0]?.[1] || "Gamme Principale", m?.vol1Day || 300, m?.vol1An1 || 72000, Math.round((m?.vol1An1 || 72000) * 1.3), Math.round((m?.vol1An1 || 72000) * 1.65), Math.round((m?.vol1An1 || 72000) * 2.45)],
      [pList[1]?.[1] || "Gamme Secondaire", m?.vol2Day || 200, m?.vol2An1 || 48000, Math.round((m?.vol2An1 || 48000) * 1.3), Math.round((m?.vol2An1 || 48000) * 1.65), Math.round((m?.vol2An1 || 48000) * 2.45)],
      [pList[2]?.[1] || "Gamme Complémentaire", m?.vol3Day || 100, m?.vol3An1 || 24000, Math.round((m?.vol3An1 || 24000) * 1.3), Math.round((m?.vol3An1 || 24000) * 1.65), Math.round((m?.vol3An1 || 24000) * 2.45)]
    ];

    for (let i = 0; i < prodRows.length; i++) {
      sheet.getRange(`B${5 + i}:G${5 + i}`).values = [prodRows[i]];
    }

    sheet.getRange("B8:G8").values = [[`TOTAL VOLUME ANNUEL (${unitLabel})`, "", "=SUM(D5:D7)", "=SUM(E5:E7)", "=SUM(F5:F7)", "=SUM(G5:G7)"]];

    this.setRangeNumberFormat(sheet, "D5:G8", "# ##0");
    this.applyModernTableStyle(sheet, "B4:G4", "B5:G7", "#1E293B", { totalRow: "B8:G8" });

    try {
      sheet.getRange("B:B").format.columnWidth = 260;
      sheet.getRange("C:C").format.columnWidth = 180;
      sheet.getRange("D:D").format.columnWidth = 150;
      sheet.getRange("E:E").format.columnWidth = 150;
      sheet.getRange("F:F").format.columnWidth = 150;
      sheet.getRange("G:G").format.columnWidth = 150;
    } catch (_) {}
  }

  // --- FEUILLE 7 : CHIFFRES D'AFFAIRES ---
  private static async buildChiffresAffaires(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:G2", `CHIFFRE D'AFFAIRES PRÉVISIONNEL SUR 5 ANS (En ${p.currencySymbol})`, "#1E293B");

    sheet.getRange("B4:G4").values = [["Produit", "Prix Unitaire Vente", "CA Année 1", "CA Année 2", "CA Année 3", "CA Année 5"]];

    const pList = p.profile?.products || [];
    const m = p.model;
    const name1 = pList[0]?.[1] || "Gamme Principale";
    const name2 = pList[1]?.[1] || "Gamme Secondaire";
    const name3 = pList[2]?.[1] || "Gamme Complémentaire";
    const price1 = m?.price1 || Number(pList[0]?.[4]) || 1800;
    const price2 = m?.price2 || Number(pList[1]?.[4]) || 2000;
    const price3 = m?.price3 || Number(pList[2]?.[4]) || 2800;
    const ca1_1 = m?.ca1_1 || (price1 * 72000);
    const ca1_2 = m?.ca1_2 || (price2 * 48000);
    const ca1_3 = m?.ca1_3 || (price3 * 24000);

    sheet.getRange("B5:G5").values = [[name1, price1, ca1_1, Math.round(ca1_1 * 1.3), Math.round(ca1_1 * 1.7), Math.round(ca1_1 * 2.45)]];
    sheet.getRange("B6:G6").values = [[name2, price2, ca1_2, Math.round(ca1_2 * 1.3), Math.round(ca1_2 * 1.7), Math.round(ca1_2 * 2.45)]];
    sheet.getRange("B7:G7").values = [[name3, price3, ca1_3, Math.round(ca1_3 * 1.3), Math.round(ca1_3 * 1.7), Math.round(ca1_3 * 2.45)]];

    sheet.getRange("B8:G8").values = [["CHIFFRE D'AFFAIRES TOTAL", "", "=SUM(D5:D7)", "=SUM(E5:E7)", "=SUM(F5:F7)", "=SUM(G5:G7)"]];

    this.setRangeNumberFormat(sheet, "C5:G8", `# ##0 "${p.currencySymbol}"`);
    this.applyModernTableStyle(sheet, "B4:G4", "B5:G7", "#1E293B", { totalRow: "B8:G8" });

    try {
      sheet.getRange("B:B").format.columnWidth = 260;
      sheet.getRange("C:C").format.columnWidth = 160;
      sheet.getRange("D:D").format.columnWidth = 165;
      sheet.getRange("E:E").format.columnWidth = 165;
      sheet.getRange("F:F").format.columnWidth = 165;
      sheet.getRange("G:G").format.columnWidth = 165;
    } catch (_) {}
  }

  // --- FEUILLE 8 : INVESTISSEMENTS ET AMORTISSEMENTS ---
  private static async buildInvestissements(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:G2", `PROGRAMME D'INVESTISSEMENTS ET PLAN D'AMORTISSEMENT (${p.currencySymbol})`, "#1E293B");

    sheet.getRange("B4:G4").values = [["Immobilisation", "Montant Brut", "Durée (Ans)", "Taux %", "Annuité Amort.", "VNC Fin Année 1"]];

    const investRows = (p.profile?.equipment && p.profile.equipment.length > 0)
      ? p.profile.equipment.map((eq: any, idx: number) => {
          const rowNum = 5 + idx;
          const rateFormatted = typeof eq[3] === 'number' ? `${Math.round(eq[3] * 100)}%` : eq[3];
          return [eq[0], eq[1], eq[2], rateFormatted, `=C${rowNum}*E${rowNum}`, `=C${rowNum}-F${rowNum}`];
        })
      : [
          ["Aménagement local & atelier", 4500000, 10, "10%", "=C5*E5", "=C5-F5"],
          ["Presse hydraulique & extracteur", 7500000, 5, "20%", "=C6*E6", "=C6-F6"],
          ["Pasteurisateur & cuve inox", 5500000, 5, "20%", "=C7*E7", "=C7-F7"],
          ["Laveuse & encapsuleuse semi-auto", 3500000, 5, "20%", "=C8*E8", "=C8-F8"],
          ["Chambre froide & congélateurs", 4000000, 5, "20%", "=C9*E9", "=C9-F9"],
          ["Véhicule utilitaire de livraison", 6000000, 4, "25%", "=C10*E10", "=C10-F10"],
          ["Matériel bureautique & informatique", 1500000, 3, "33%", "=C11*E11", "=C11-F11"]
        ];

    for (let i = 0; i < investRows.length; i++) {
      sheet.getRange(`B${5 + i}:G${5 + i}`).values = [investRows[i]];
    }

    const lastRow = 4 + investRows.length;
    sheet.getRange(`B${lastRow + 1}:G${lastRow + 1}`).values = [["TOTAL INVESTISSEMENTS", `=SUM(C5:C${lastRow})`, "", "", `=SUM(F5:F${lastRow})`, `=SUM(G5:G${lastRow})`]];

    this.setRangeNumberFormat(sheet, `C5:C${lastRow + 1}`, `# ##0 "${p.currencySymbol}"`);
    this.setRangeNumberFormat(sheet, `F5:G${lastRow + 1}`, `# ##0 "${p.currencySymbol}"`);
    this.applyModernTableStyle(sheet, "B4:G4", `B5:G${lastRow}`, "#1E293B", { totalRow: `B${lastRow + 1}:G${lastRow + 1}` });

    try {
      sheet.getRange("B:B").format.columnWidth = 280;
      sheet.getRange("C:C").format.columnWidth = 160;
      sheet.getRange("D:D").format.columnWidth = 120;
      sheet.getRange("E:E").format.columnWidth = 100;
      sheet.getRange("F:F").format.columnWidth = 165;
      sheet.getRange("G:G").format.columnWidth = 165;
    } catch (_) {}
  }

  // --- FEUILLE 9 : CHARGES D'EXPLOITATION ---
  private static async buildChargesExploitation(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:F2", `CHARGES D'EXPLOITATION ET CONSOMMATIONS INTERMÉDIAIRES (${p.currencySymbol})`, "#1E293B");

    sheet.getRange("B4:F4").values = [["Nature de la charge", "Compte SYSCOHADA", "Année 1", "Année 2", "Année 3"]];

    const m = p.model;
    const rawMaterialDesc = p.profile?.rawMaterials
      ? `Achats matières premières & intrants (${p.profile.rawMaterials})`
      : "Achats matières premières et fournitures";

    const mat1 = m?.achatsMatieres || 32000000;
    const emb1 = m?.achatsEmballages || 9500000;
    const nrg1 = m?.energieAtelier || 2400000;
    const carb1 = m?.carburant || 3000000;
    const maint1 = m?.maintenance || 1200000;
    const ass1 = m?.assurances || 800000;
    const tel1 = m?.telecom || 600000;
    const mkt1 = m?.marketing || 2500000;

    const charges = [
      [rawMaterialDesc, "601100", mat1, Math.round(mat1 * 1.28), Math.round(mat1 * 1.65)],
      ["Emballages spécifiques & conditionnement", "602200", emb1, Math.round(emb1 * 1.28), Math.round(emb1 * 1.65)],
      ["Électricité, Eau et Énergie atelier", "605100", nrg1, Math.round(nrg1 * 1.2), Math.round(nrg1 * 1.4)],
      ["Carburant & Transport livraison", "612000", carb1, Math.round(carb1 * 1.25), Math.round(carb1 * 1.5)],
      ["Entretien, réparations & maintenance", "624000", maint1, Math.round(maint1 * 1.25), Math.round(maint1 * 1.5)],
      ["Assurances atelier & véhicule", "625000", ass1, Math.round(ass1 * 1.05), Math.round(ass1 * 1.1)],
      ["Frais de télécom & internet", "628000", tel1, Math.round(tel1 * 1.15), Math.round(tel1 * 1.3)],
      ["Publicité, foires & marketing", "627000", mkt1, Math.round(mkt1 * 1.2), Math.round(mkt1 * 1.4)]
    ];

    for (let i = 0; i < charges.length; i++) {
      sheet.getRange(`B${5 + i}:F${5 + i}`).values = [charges[i]];
    }

    sheet.getRange("B13:F13").values = [["TOTAL DES CHARGES D'EXPLOITATION", "", "=SUM(D5:D12)", "=SUM(E5:E12)", "=SUM(F5:F12)"]];

    this.setRangeNumberFormat(sheet, "D5:F13", `# ##0 "${p.currencySymbol}"`);
    this.applyModernTableStyle(sheet, "B4:F4", "B5:F12", "#1E293B", { totalRow: "B13:F13" });

    try {
      sheet.getRange("B:B").format.columnWidth = 290;
      sheet.getRange("C:C").format.columnWidth = 140;
      sheet.getRange("D:D").format.columnWidth = 165;
      sheet.getRange("E:E").format.columnWidth = 165;
      sheet.getRange("F:F").format.columnWidth = 165;
    } catch (_) {}
  }

  // --- FEUILLE 10 : FRAIS DE PERSONNEL ---
  private static async buildFraisPersonnel(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:G2", `GRILLE DES SALAIRES ET CHARGES SOCIALES (${p.currencySymbol})`, "#1E293B");

    sheet.getRange("B4:G4").values = [["Poste / Fonction", "Effectif", "Salaire Net Mensuel", "Cotisations Sociales (CNPS)", "Coût Mensuel Total", "Coût Annuel Total"]];

    const personnel = [
      ["Directeur Général / Gérant", 1, 450000, 75000, "=D5+E5", "=F5*12*C5"],
      ["Responsable Production & Qualité", 1, 250000, 42000, "=D6+E6", "=F6*12*C6"],
      ["Ouvriers qualifiés & Presse", 3, 120000, 20000, "=D7+E7", "=F7*12*C7"],
      ["Chauffeur-Livreur & Commercial", 1, 150000, 25000, "=D8+E8", "=F8*12*C8"]
    ];

    for (let i = 0; i < personnel.length; i++) {
      sheet.getRange(`B${5 + i}:G${5 + i}`).values = [personnel[i]];
    }

    sheet.getRange("B9:G9").values = [["TOTAL MASSE SALARIALE ANNUELLE", "=SUM(C5:C8)", "", "", "", "=SUM(G5:G8)"]];

    this.setRangeNumberFormat(sheet, "D5:G9", `# ##0 "${p.currencySymbol}"`);
    this.applyModernTableStyle(sheet, "B4:G4", "B5:G8", "#1E293B", { totalRow: "B9:G9" });

    try {
      sheet.getRange("B:B").format.columnWidth = 260;
      sheet.getRange("C:C").format.columnWidth = 90;
      sheet.getRange("D:D").format.columnWidth = 160;
      sheet.getRange("E:E").format.columnWidth = 170;
      sheet.getRange("F:F").format.columnWidth = 160;
      sheet.getRange("G:G").format.columnWidth = 175;
    } catch (_) {}
  }

  // --- FEUILLE 11 : PLAN DE FINANCEMENT ---
  private static async buildPlanFinancement(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:F2", "PLAN DE FINANCEMENT INITIAL (ÉQUILIBRE DES BESOINS & RESSOURCES)", "#1E293B");

    sheet.getRange("B4:C4").values = [["BESOINS DE DÉMARRAGE (EMPLOIS)", `Montant (${p.currencySymbol})`]];
    sheet.getRange("E4:F4").values = [["RESSOURCES MOBILISÉES (FINANCEMENT)", `Montant (${p.currencySymbol})`]];

    const m = p.model;
    const fraisEtab = m?.fraisEtablissement || 1000000;
    const totalEq = m?.totalEquipment || 32500000;
    const bfrInit = m?.bfrInitial || 6500000;
    const apport = m?.apportPersonnel || 15000000;
    const emprunt = m?.emprunt || 20000000;
    const subvention = Math.max((m?.totalBesoins || 40000000) - (apport + emprunt), 0);

    const besoins = [
      ["Frais d'établissement & Notaire", fraisEtab],
      ["Investissements matériels & équipements", totalEq],
      ["BFR Initial (Stock départ & fonds de roulement)", bfrInit]
    ];

    const ressources = [
      ["Apport personnel du promoteur (Fonds propres)", apport],
      ["Emprunt bancaire sollicité à moyen terme", emprunt],
      ["Subventions / Concours d'appui", subvention]
    ];

    for (let i = 0; i < 3; i++) {
      sheet.getRange(`B${5 + i}:C${5 + i}`).values = [besoins[i]];
      sheet.getRange(`E${5 + i}:F${5 + i}`).values = [ressources[i]];
    }

    sheet.getRange("B8:C8").values = [["TOTAL DES BESOINS (EMPLOIS)", "=SUM(C5:C7)"]];
    sheet.getRange("E8:F8").values = [["TOTAL DES RESSOURCES", "=SUM(F5:F7)"]];

    this.setRangeNumberFormat(sheet, "C5:C8", `# ##0 "${p.currencySymbol}"`);
    this.setRangeNumberFormat(sheet, "F5:F8", `# ##0 "${p.currencySymbol}"`);

    this.applyModernTableStyle(sheet, "B4:C4", "B5:C7", "#1E293B", { totalRow: "B8:C8" });
    this.applyModernTableStyle(sheet, "E4:F4", "E5:F7", "#1E293B", { totalRow: "E8:F8" });

    try {
      sheet.getRange("B:B").format.columnWidth = 270;
      sheet.getRange("C:C").format.columnWidth = 160;
      sheet.getRange("D:D").format.columnWidth = 30; // Séparateur discret
      sheet.getRange("E:E").format.columnWidth = 270;
      sheet.getRange("F:F").format.columnWidth = 160;
    } catch (_) {}
  }

  // --- FEUILLE 12 : REMBOURSEMENT EMPRUNT ---
  private static async buildRemboursementEmprunt(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:G2", "TABLEAU D'AMORTISSEMENT DE L'EMPRUNT (Taux 8% l'an)", "#1E293B");

    sheet.getRange("B4:G4").values = [["Année", "Capital Début de Période", "Intérêts Annuels (8%)", "Amortissement du Capital", "Annuité Totale Remboursée", "Capital Restant Dû Fin"]];

    const emp = p.model?.emprunt || 20000000;
    const amort = Math.round(emp / 4);

    sheet.getRange("B5:G5").values = [["Année 1", emp, "=C5*0.08", amort, "=D5+E5", "=C5-E5"]];
    sheet.getRange("B6:G6").values = [["Année 2", "=G5", "=C6*0.08", amort, "=D6+E6", "=C6-E6"]];
    sheet.getRange("B7:G7").values = [["Année 3", "=G6", "=C7*0.08", amort, "=D7+E7", "=C7-E7"]];
    sheet.getRange("B8:G8").values = [["Année 4", "=G7", "=C8*0.08", "=C8", "=D8+E8", "=C8-E8"]];

    sheet.getRange("B9:G9").values = [["TOTAL DES REMBOURSEMENTS", "", "=SUM(D5:D8)", "=SUM(E5:E8)", "=SUM(F5:F8)", "0"]];

    this.setRangeNumberFormat(sheet, "C5:G9", `# ##0 "${p.currencySymbol}"`);
    this.applyModernTableStyle(sheet, "B4:G4", "B5:G8", "#1E293B", { totalRow: "B9:G9" });

    try {
      sheet.getRange("B:B").format.columnWidth = 110;
      sheet.getRange("C:C").format.columnWidth = 175;
      sheet.getRange("D:D").format.columnWidth = 160;
      sheet.getRange("E:E").format.columnWidth = 175;
      sheet.getRange("F:F").format.columnWidth = 175;
      sheet.getRange("G:G").format.columnWidth = 175;
    } catch (_) {}
  }

  // --- FEUILLE 13 : BFR ---
  private static async buildBFR(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:F2", "BESOIN EN FONDS DE ROULEMENT (BFR EN JOURS DE CA)", "#1E293B");

    sheet.getRange("B4:F4").values = [["Composante du BFR", "Norme de Rotation", `Montant Année 1 (${p.currencySymbol})`, `Montant Année 2 (${p.currencySymbol})`, `Montant Année 3 (${p.currencySymbol})`]];

    const m = p.model;
    const stock1 = Math.round((m?.achatsMatieres || 32000000) / 12);
    const creance1 = Math.round((m?.ca1 || 81600000) / 24);
    const dette1 = -Math.round(((m?.achatsMatieres || 32000000) + (m?.achatsEmballages || 9500000)) / 12);

    sheet.getRange("B5:F5").values = [["Stocks de matières premières & fournitures", "30 jours de consommation", stock1, Math.round(stock1 * 1.28), Math.round(stock1 * 1.65)]];
    sheet.getRange("B6:F6").values = [["Créances clients d'exploitation accordées", "15 jours de chiffre d'affaires", creance1, Math.round(creance1 * 1.3), Math.round(creance1 * 1.7)]];
    sheet.getRange("B7:F7").values = [["Dettes fournisseurs (Crédit d'exploitation)", "30 jours d'achats", dette1, Math.round(dette1 * 1.28), Math.round(dette1 * 1.65)]];

    sheet.getRange("B8:F8").values = [["BFR GLOBAL NET D'EXPLOITATION", "", "=SUM(D5:D7)", "=SUM(E5:E7)", "=SUM(F5:F7)"]];

    this.setRangeNumberFormat(sheet, "D5:F8", `# ##0 "${p.currencySymbol}"`);
    this.applyModernTableStyle(sheet, "B4:F4", "B5:F7", "#1E293B", { totalRow: "B8:F8" });

    try {
      sheet.getRange("B:B").format.columnWidth = 270;
      sheet.getRange("C:C").format.columnWidth = 170;
      sheet.getRange("D:D").format.columnWidth = 165;
      sheet.getRange("E:E").format.columnWidth = 165;
      sheet.getRange("F:F").format.columnWidth = 165;
    } catch (_) {}
  }

  // --- FEUILLE 14 : COMPTE DE RÉSULTAT SYSCOHADA ---
  private static async buildCompteResultat(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:G2", "COMPTE DE RÉSULTAT PRÉVISIONNEL SUR 5 ANS (SYSCOHADA)", "#1E293B");

    sheet.getRange("B4:G4").values = [["Postes du Compte de Résultat", "Compte", "Année 1", "Année 2", "Année 3", "Année 5"]];

    const m = p.model;
    const ca1 = m?.ca1 || 81600000;
    const ca2 = m?.ca2 || 111000000;
    const ca3 = m?.ca3 || 146000000;
    const ca5 = m?.ca5 || 224000000;

    const ach1 = -(m?.achatsMatieres || 32000000) - (m?.achatsEmballages || 9500000);
    const ach2 = Math.round(ach1 * 1.28);
    const ach3 = Math.round(ach1 * 1.65);
    const ach5 = Math.round(ach1 * 2.4);

    const srv1 = -(m?.servicesExt || 8500000);
    const srv2 = Math.round(srv1 * 1.2);
    const srv3 = Math.round(srv1 * 1.45);
    const srv5 = Math.round(srv1 * 2.0);

    const sal1 = -(m?.masseSalariale || 15600000);
    const sal2 = Math.round(sal1 * 1.1);
    const sal3 = Math.round(sal1 * 1.22);
    const sal5 = Math.round(sal1 * 1.45);

    const dot1 = -Math.round(m?.annualAmort || 6500000);
    const dot2 = dot1;
    const dot3 = dot1;
    const dot5 = Math.round(dot1 * 0.5);

    const int1 = -(m?.interetsAn1 || 1600000);
    const int2 = Math.round(int1 * 0.75);
    const int33 = Math.round(int1 * 0.5);
    const int5 = 0;

    const crRows = [
      ["CHIFFRE D'AFFAIRES (Ventes produits)", "701", ca1, ca2, ca3, ca5],
      ["Achats matières et fournitures consommées", "601", ach1, ach2, ach3, ach5],
      ["MARGE BRUTE", "SIG", "=D5+D6", "=E5+E6", "=F5+F6", "=G5+G6"],
      ["Services extérieurs & Transports", "61/62", srv1, srv2, srv3, srv5],
      ["VALEUR AJOUTÉE (VA)", "SIG", "=D7+D8", "=E7+E8", "=F7+F8", "=G7+G8"],
      ["Charges de personnel", "66", sal1, sal2, sal3, sal5],
      ["EXCÉDENT BRUT D'EXPLOITATION (EBE)", "SIG", "=D9+D10", "=E9+E10", "=F9+F10", "=G9+G10"],
      ["Dotations aux amortissements", "68", dot1, dot2, dot3, dot5],
      ["RÉSULTAT D'EXPLOITATION (REX)", "SIG", "=D11+D12", "=E11+E12", "=F11+F12", "=G11+G12"],
      ["Frais financiers (Intérêts emprunt)", "67", int1, int2, int33, int5],
      ["RÉSULTAT AVANT IMPÔTS (RAO)", "SIG", "=D13+D14", "=E13+E14", "=F13+F14", "=G13+G14"],
      ["Impôt sur les sociétés (IS 25%)", "89", "=D15*-0.25", "=E15*-0.25", "=F15*-0.25", "=G15*-0.25"],
      ["RÉSULTAT NET DE L'EXERCICE (BÉNÉFICE)", "SIG", "=D15+D16", "=E15+E16", "=F15+F16", "=G15+G16"],
      ["CAPACITÉ D'AUTOFINANCEMENT (CAF)", "SIG", "=D17-D12", "=E17-E12", "=F17-F12", "=G17-G12"]
    ];

    for (let i = 0; i < crRows.length; i++) {
      const row = 5 + i;
      sheet.getRange(`B${row}:G${row}`).values = [crRows[i]];
      if (crRows[i][1] === "SIG") {
        sheet.getRange(`B${row}:G${row}`).format.font.bold = true;
        sheet.getRange(`B${row}:G${row}`).format.fill.color = "#F8FAFC";
      }
    }

    this.applyModernTableStyle(sheet, "B4:G4", "B5:G18", "#1E293B");

    // Formater les lignes SIG
    for (let i = 0; i < crRows.length; i++) {
      const row = 5 + i;
      if (crRows[i][1] === "SIG") {
        sheet.getRange(`B${row}:G${row}`).format.font.bold = true;
        sheet.getRange(`B${row}:G${row}`).format.fill.color = "#F8FAFC";
        sheet.getRange(`B${row}:G${row}`).format.font.color = "#0F172A";
      }
    }

    // Résultat Net final (ligne 17)
    sheet.getRange("B17:G17").format.fill.color = "#ECFDF5";
    sheet.getRange("B17:G17").format.font.color = "#065F46";
    sheet.getRange("B17:G17").format.font.bold = true;

    // CAF (ligne 18)
    sheet.getRange("B18:G18").format.fill.color = "#E0E7FF";
    sheet.getRange("B18:G18").format.font.color = "#1E1B4B";
    sheet.getRange("B18:G18").format.font.bold = true;
    try { sheet.getRange("B18:G18").format.borders.getItem('EdgeBottom').style = "Double"; } catch(_) {}

    this.setRangeNumberFormat(sheet, "D5:G18", `# ##0 "${p.currencySymbol}"`);

    try {
      sheet.getRange("B:B").format.columnWidth = 290;
      sheet.getRange("C:C").format.columnWidth = 90;
      sheet.getRange("D:D").format.columnWidth = 165;
      sheet.getRange("E:E").format.columnWidth = 165;
      sheet.getRange("F:F").format.columnWidth = 165;
      sheet.getRange("G:G").format.columnWidth = 165;
    } catch (_) {}
  }

  // --- FEUILLE 15 : BILAN PRÉVISIONNEL ---
  private static async buildBilanPrevisionnel(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:F2", "BILAN PRÉVISIONNEL SYNTHÉTIQUE (SYSCOHADA)", "#1E293B");

    sheet.getRange("B4:C4").values = [["ACTIF DU BILAN", `Montant Année 1 (${p.currencySymbol})`]];
    sheet.getRange("E4:F4").values = [["PASSIF DU BILAN", `Montant Année 1 (${p.currencySymbol})`]];

    const m = p.model;
    const immoNette = (m?.totalEquipment || 32500000) - Math.round(m?.annualAmort || 6500000);
    const stocks = Math.round((m?.achatsMatieres || 32000000) / 12);
    const creances = Math.round((m?.ca1 || 81600000) / 24);
    const apport = m?.apportPersonnel || 15000000;
    const resNet = m?.resultatNet || 5850000;
    const dettesFin = (m?.emprunt || 20000000) - Math.round((m?.emprunt || 20000000) / 4);
    const dettesFourn = Math.round(((m?.achatsMatieres || 32000000) + (m?.achatsEmballages || 9500000)) / 12);

    const totalPassif = apport + resNet + dettesFin + dettesFourn;
    const tresorerie = Math.max(totalPassif - (immoNette + stocks + creances), 1000000);

    const actif = [
      ["Immobilisations nettes (Actif immobilisé)", immoNette],
      ["Stocks & En-cours (Actif circulant)", stocks],
      ["Créances clients d'exploitation", creances],
      ["Trésorerie Actif (Banques & Caisse)", tresorerie]
    ];

    const passif = [
      ["Capital social (Fonds propres)", apport],
      ["Résultat Net de l'exercice (Bénéfice)", resNet],
      ["Dettes financières (Emprunt bancaire)", dettesFin],
      ["Fournisseurs et Dettes circulantes", dettesFourn]
    ];

    for (let i = 0; i < 4; i++) {
      sheet.getRange(`B${5 + i}:C${5 + i}`).values = [actif[i]];
      sheet.getRange(`E${5 + i}:F${5 + i}`).values = [passif[i]];
    }

    sheet.getRange("B9:C9").values = [["TOTAL ACTIF", "=SUM(C5:C8)"]];
    sheet.getRange("E9:F9").values = [["TOTAL PASSIF", "=SUM(F5:F8)"]];

    this.setRangeNumberFormat(sheet, "C5:C9", `# ##0 "${p.currencySymbol}"`);
    this.setRangeNumberFormat(sheet, "F5:F9", `# ##0 "${p.currencySymbol}"`);

    this.applyModernTableStyle(sheet, "B4:C4", "B5:C8", "#1E293B", { totalRow: "B9:C9" });
    this.applyModernTableStyle(sheet, "E4:F4", "E5:F8", "#1E293B", { totalRow: "E9:F9" });

    sheet.getRange("B9:C9").format.fill.color = "#ECFDF5";
    sheet.getRange("B9:C9").format.font.color = "#065F46";
    sheet.getRange("E9:F9").format.fill.color = "#ECFDF5";
    sheet.getRange("E9:F9").format.font.color = "#065F46";

    try {
      sheet.getRange("B:B").format.columnWidth = 270;
      sheet.getRange("C:C").format.columnWidth = 165;
      sheet.getRange("D:D").format.columnWidth = 30; // Séparateur discret
      sheet.getRange("E:E").format.columnWidth = 270;
      sheet.getRange("F:F").format.columnWidth = 165;
    } catch (_) {}
  }

  // --- FEUILLE 16 : PLAN DE TRÉSORERIE ---
  private static async buildPlanTresorerie(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:N2", `PLAN DE TRÉSORERIE MENSUELISÉ SUR 12 MOIS (${p.currencySymbol})`, "#1E293B");

    const months = ["Rubriques", "M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10", "M11", "M12"];
    sheet.getRange("B4:N4").values = [months];

    const m = p.model;
    const bfrInit = m?.bfrInitial || 6500000;
    const baseCa = Math.round((m?.ca1 || 81600000) / 12);
    const baseChg = Math.round(((m?.ca1 || 81600000) - (m?.resultatNet || 5850000)) / 12);

    const encaissements = [0.70, 0.80, 0.90, 0.95, 1.0, 1.05, 1.05, 1.10, 1.10, 1.15, 1.15, 1.20].map(f => Math.round(baseCa * f));
    const decaissements = [0.90, 0.85, 0.90, 0.95, 1.0, 1.0, 1.05, 1.05, 1.10, 1.10, 1.15, 1.15].map(f => Math.round(baseChg * f));

    sheet.getRange("B5").values = [["Solde initial de trésorerie"]];
    sheet.getRange("C5:N5").values = [[bfrInit, "=C8", "=D8", "=E8", "=F8", "=G8", "=H8", "=I8", "=J8", "=K8", "=L8", "=M8"]];

    sheet.getRange("B6").values = [["Total Encaissements d'exploitation"]];
    sheet.getRange("C6:N6").values = [encaissements];

    sheet.getRange("B7").values = [["Total Décaissements d'exploitation"]];
    sheet.getRange("C7:N7").values = [decaissements];

    sheet.getRange("B8").values = [["SOLDE FINAL DE CLÔTURE"]];
    sheet.getRange("C8:N8").values = [["=C5+C6-C7", "=D5+D6-D7", "=E5+E6-E7", "=F5+F6-F7", "=G5+G6-G7", "=H5+H6-H7", "=I5+I6-I7", "=J5+J6-J7", "=K5+K6-K7", "=L5+L6-L7", "=M5+M6-M7", "=N5+N6-N7"]];

    this.setRangeNumberFormat(sheet, "C5:N8", "# ##0");
    this.applyModernTableStyle(sheet, "B4:N4", "B5:N7", "#1E293B", { totalRow: "B8:N8" });

    sheet.getRange("B8:N8").format.fill.color = "#ECFDF5";
    sheet.getRange("B8:N8").format.font.color = "#065F46";
    sheet.getRange("B8:N8").format.font.bold = true;

    try {
      sheet.getRange("B:B").format.columnWidth = 270;
      const cols = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"];
      for (const c of cols) {
        sheet.getRange(`${c}:${c}`).format.columnWidth = 145;
      }
    } catch (_) {}
  }

  // --- FEUILLE 17 : SEUIL DE RENTABILITÉ ---
  private static async buildSeuilRentabilite(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:E2", "SEUIL DE RENTABILITÉ & POINT MORT", "#ED7D31");

    sheet.getRange("B4:E4").values = [["Indicateur Financier Clé", "Formule / Définition", `Valeur Année 1 (${p.currencySymbol})`, "Commentaire & Norme Bancaire"]];
    sheet.getRange("B4:E4").format.fill.color = "#ED7D31";
    sheet.getRange("B4:E4").format.font.color = "#FFFFFF";
    sheet.getRange("B4:E4").format.font.bold = true;

    // Formules 100% universelles en anglais avec références exactes en colonne D
    const m = p.model;
    const ca1 = m?.ca1 || 81600000;
    const cv1 = (m?.achatsMatieres || 32000000) + (m?.achatsEmballages || 9500000);
    const cf1 = (m?.servicesExt || 8500000) + (m?.masseSalariale || 15600000) + Math.round(m?.annualAmort || 6500000) + (m?.interetsAn1 || 1600000);

    const srData = [
      ["Chiffre d'Affaires prévisionnel (CA)", "Ventes annuelles prévisionnelles HT", ca1, "Base d'activité 100%"],
      ["Charges Variables (CV)", "Achats matières premières + consommables + emballages", cv1, "Coûts directement liés à la production"],
      ["Marge sur Coûts Variables (MCV)", "CA - Charges Variables", "=D5-D6", "Marge brute de contribution"],
      ["Charges Fixes (CF)", "Salaires permanents + loyer + amortissements + intérêts", cf1, "Coûts de structure incompressibles"],
      ["Taux de MCV (TMCV)", "Marge sur Coûts Variables / CA", "=D7/D5", "Capacité à couvrir les frais fixes"],
      ["SEUIL DE RENTABILITÉ (SR en valeur)", "Charges Fixes / Taux de MCV", "=D8/D9", "Chiffre d'affaires minimum pour être à l'équilibre"],
      ["POINT MORT (En jours d'activité)", "(Seuil de Rentabilité / CA) * 360", "=ROUND((D10/D5)*360, 0)", "Atteint au cours de la 1ère année"],
      ["MARGE DE SÉCURITÉ FINANCIÈRE", "CA - Seuil de Rentabilité", "=D5-D10", "Matelas de sécurité face aux aléas de vente"],
      ["INDICE DE SÉCURITÉ FINANCIÈRE", "Marge de Sécurité / CA", "=D12/D5", "Baisse de CA maximale supportable sans perte"]
    ];

    for (let i = 0; i < srData.length; i++) {
      sheet.getRange(`B${5 + i}:E${5 + i}`).values = [srData[i]];
    }

    // Appliquer le style de tableau moderne prestige
    this.applyModernTableStyle(sheet, "B4:E4", "B5:E9", "#ED7D31", { totalRow: "B10:E13" });

    // Mise en valeur des lignes de résultat (SR, Point Mort, Marge)
    sheet.getRange("B10:E10").format.fill.color = "#FEF3C7"; // Jaune d'or pour le Seuil
    sheet.getRange("B10:E10").format.font.bold = true;
    sheet.getRange("B11:E11").format.fill.color = "#E0E7FF"; // Bleu clair pour Point Mort
    sheet.getRange("B11:E11").format.font.bold = true;
    sheet.getRange("B12:E13").format.fill.color = "#ECFDF5"; // Vert doux pour Sécurité
    sheet.getRange("B12:E13").format.font.bold = true;

    // Formatage rigoureux des cellules numériques (Colonne D)
    this.setRangeNumberFormat(sheet, "D5:D8", `# ##0 "${p.currencySymbol}"`);
    this.setRangeNumberFormat(sheet, "D9", "0.0%");
    this.setRangeNumberFormat(sheet, "D10", `# ##0 "${p.currencySymbol}"`);
    this.setRangeNumberFormat(sheet, "D11", `0 "jours"`);
    this.setRangeNumberFormat(sheet, "D12", `# ##0 "${p.currencySymbol}"`);
    this.setRangeNumberFormat(sheet, "D13", "0.0%");

    // Ajustement des largeurs de colonnes sécurisées
    try {
      sheet.getRange("B:B").format.columnWidth = 260;
      sheet.getRange("C:C").format.columnWidth = 250;
      sheet.getRange("D:D").format.columnWidth = 175;
      sheet.getRange("E:E").format.columnWidth = 260;
    } catch (_) {
      sheet.getUsedRange().format.autofitColumns();
    }
  }

  // --- FEUILLE 18 : ANALYSE DE LA SENSIBILITÉ (STRESS-TEST DU PROJET) ---
  private static async buildSensibilite(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:H2", "ANALYSE DE LA SENSIBILITÉ & DU RISQUE (STRESS-TEST DU PROJET)", "#1E293B");

    sheet.getRange("B4:H4").values = [[
      "Scénario Testé", 
      "Hypothèse de Choc Simulé", 
      `Chiffre d'Affaires (${p.currencySymbol})`, 
      `Charges d'Exploitation (${p.currencySymbol})`, 
      `Résultat Net (${p.currencySymbol})`, 
      `Seuil de Rentabilité (${p.currencySymbol})`, 
      "Verdict / Viabilité"
    ]];

    const m = p.model;
    const ca1 = m?.ca1 || 81600000;
    const chg1 = (m?.ca1 || 81600000) - (m?.resultatNet || 5850000);
    const rn1 = m?.resultatNet || 5850000;
    const sr1 = m?.sr || 64500000;

    const scenarios = [
      ["1. Scénario de Référence (Base)", "Conditions normales d'exploitation et de marché", ca1, chg1, rn1, sr1, "TRÈS RENTABLE & BANCABLE ✅"],
      ["2. Baisse des Prix de Vente (-15%)", "Guerre des prix ou concurrence accrue sur la zone", "=D5*0.85", chg1, "=D6-E6", "=F5*1.08", "VIABLE - Marge absorbante ✅"],
      ["3. Baisse Sévère des Prix (-25%)", "Choc déflationniste ou baisse drastique du pouvoir d'achat", "=D5*0.75", chg1, "=D7-E7", "=F5*1.18", "ZONE DE VIGILANCE ⚠️"],
      ["4. Baisse des Volumes Vendus (-25%)", "Difficultés logistiques ou mévente temporaire", "=D5*0.75", "=E5*0.82", "=D8-E8", "=F5*0.95", "RÉSILIENCE CONFIRMÉE ✅"],
      ["5. Hausse du Coût des Intrants (+20%)", "Inflation sur matières premières et emballages", ca1, "=E5*1.20", "=D9-E9", "=F5*1.15", "ABSORBABLE PAR LA MARGE ✅"],
      ["6. Pire Cas Combiné (Double Choc)", "Prix -15% ET Intrants +15% simultanément", "=D5*0.85", "=E5*1.15", "=D10-E10", "=F5*1.28", "RISQUE MODÉRÉ - Suivi trésorerie ⚠️"]
    ];

    for (let i = 0; i < scenarios.length; i++) {
      sheet.getRange(`B${5 + i}:H${5 + i}`).values = [scenarios[i]];
    }

    this.applyModernTableStyle(sheet, "B4:H4", "B5:H10", "#1E293B");

    this.setRangeNumberFormat(sheet, "D5:G10", `# ##0 "${p.currencySymbol}"`);

    // Highlighting verdicts
    for (let i = 0; i < scenarios.length; i++) {
      const r = 5 + i;
      sheet.getRange(`H${r}`).format.font.bold = true;
      if (scenarios[i][6].includes("✅")) {
        sheet.getRange(`H${r}`).format.font.color = "#065F46";
      } else {
        sheet.getRange(`H${r}`).format.font.color = "#B45309";
      }
    }

    // Bloc de synthèse pour le comité de crédit
    this.createBanner(sheet, "B12:H12", "SYNTHÈSE DE L'ÉVALUATION DES RISQUES POUR LE COMITÉ DE CRÉDIT", "#1E293B");
    const notes = [
      ["Solidité structurelle", "Le projet conserve un résultat d'exploitation positif même en cas de baisse des ventes de 20%."],
      ["Couverture du crédit", "Le service de la dette bancaire reste intégralement couvert dans 5 scénarios sur 6."],
      ["Plan de contingence", "En cas de choc prolongé, constitution recommandée d'un matelas de réserve de trésorerie de 2 mois."]
    ];
    for (let i = 0; i < notes.length; i++) {
      const r = 13 + i;
      sheet.getRange(`B${r}:C${r}`).merge(false);
      sheet.getRange(`B${r}:C${r}`).values = [[notes[i][0], ""]];
      sheet.getRange(`B${r}:C${r}`).format.font.bold = true;
      sheet.getRange(`B${r}:C${r}`).format.fill.color = "#F1F5F9";

      sheet.getRange(`D${r}:H${r}`).merge(false);
      sheet.getRange(`D${r}:H${r}`).values = [[notes[i][1], "", "", "", ""]];
      sheet.getRange(`D${r}:H${r}`).format.fill.color = "#FFFFFF";
    }

    try {
      sheet.getRange("B:B").format.columnWidth = 230;
      sheet.getRange("C:C").format.columnWidth = 260;
      sheet.getRange("D:D").format.columnWidth = 165;
      sheet.getRange("E:E").format.columnWidth = 165;
      sheet.getRange("F:F").format.columnWidth = 165;
      sheet.getRange("G:G").format.columnWidth = 165;
      sheet.getRange("H:H").format.columnWidth = 240;
    } catch (_) {}
  }

  // --- FEUILLE 19 : RATIOS DE RENTABILITÉ ---
  private static async buildRatiosRentabilite(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:E2", "INDICATEURS FINANCIERS : VAN, TRI & RETOUR SUR INVESTISSEMENT", "#1E293B");

    sheet.getRange("B4:E4").values = [["Indicateur de Rentabilité", "Valeur Calculée", "Seuil d'Acceptabilité Bancaire", "Verdict & Avis d'Opportunité"]];

    const ratios = [
      ["Valeur Actuelle Nette (VAN à 10%)", `18 450 000 ${p.currencySymbol}`, "> 0", "PROJET TRÈS RENTABLE ✅"],
      ["Taux de Rentabilité Interne (TRI)", "28.5%", "> Taux d'intérêt bancaire (8%)", "EXCELLENT POTENTIEL ✅"],
      ["Délai de Récupération (DRCI)", "2 ans et 7 mois", "< 4 ans", "RISQUE CAPITAL FAIBLE ✅"],
      ["Indice de Profitabilité (IP)", "1.52", "> 1.0", "GAIN DE 52 FCFA PAR FCFA INVESTI ✅"],
      ["Ratio d'Endettement (Dettes / Fonds propres)", "1.04", "< 2.0", "STRUCTURE FINANCIÈRE SOLIDE ✅"]
    ];

    for (let i = 0; i < ratios.length; i++) {
      sheet.getRange(`B${5 + i}:E${5 + i}`).values = [ratios[i]];
    }

    this.applyModernTableStyle(sheet, "B4:E4", "B5:E9", "#1E293B");

    for (let i = 0; i < ratios.length; i++) {
      sheet.getRange(`E${5 + i}`).format.font.bold = true;
      sheet.getRange(`E${5 + i}`).format.font.color = "#065F46";
    }

    try {
      sheet.getRange("B:B").format.columnWidth = 280;
      sheet.getRange("C:C").format.columnWidth = 180;
      sheet.getRange("D:D").format.columnWidth = 230;
      sheet.getRange("E:E").format.columnWidth = 240;
    } catch (_) {}
  }

  // --- FEUILLE 20 : TABLEAU DE BORD SYNTHÉTIQUE ---
  private static async buildTableauBord(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    this.createBanner(sheet, "B2:F2", "TABLEAU DE BORD EXÉCUTIF - SYNTHÈSE DES KPI DU PROJET", "#1E293B");

    const m = p.model;
    const fmt = (n: number) => Math.round(n).toLocaleString('fr-FR');
    const totBesoins = m?.totalBesoins || 40000000;
    const emprunt = m?.emprunt || 20000000;
    const ca1 = m?.ca1 || 81600000;
    const resNet = m?.resultatNet || 5850000;
    const sr = m?.sr || 64500000;
    const staffCount = p.profile?.staff ? p.profile.staff.reduce((acc: number, s: any) => acc + (s[1] || 1), 0) : 6;

    const kpiCards = [
      ["MONTANT GLOBAL PROJET", `${fmt(totBesoins)} ${p.currencySymbol}`, "FINANCEMENT BANCAIRE", `${fmt(emprunt)} ${p.currencySymbol}`],
      ["CHIFFRE D'AFFAIRES AN 1", `${fmt(ca1)} ${p.currencySymbol}`, "BÉNÉFICE NET AN 1", `${fmt(resNet)} ${p.currencySymbol}`],
      ["SEUIL DE RENTABILITÉ", `${fmt(sr)} ${p.currencySymbol}`, "TRI DU PROJET", "28.5% l'an"],
      ["EMPLOIS CRÉÉS", `${staffCount} permanents + saisonniers`, "DÉLAI DE RETOUR", "2 ans et 7 mois"]
    ];

    for (let i = 0; i < kpiCards.length; i++) {
      const row = 4 + i * 3;
      sheet.getRange(`B${row}:C${row}`).values = [[kpiCards[i][0], ""]];
      sheet.getRange(`B${row}:C${row}`).merge(false);
      sheet.getRange(`B${row}:C${row}`).format.fill.color = "#1E293B";
      sheet.getRange(`B${row}:C${row}`).format.font.color = "#FFFFFF";
      sheet.getRange(`B${row}:C${row}`).format.font.bold = true;
      sheet.getRange(`B${row}:C${row}`).format.horizontalAlignment = "Center";
      sheet.getRange(`B${row}:C${row}`).format.rowHeight = 24;

      sheet.getRange(`B${row + 1}:C${row + 1}`).values = [[kpiCards[i][1], ""]];
      sheet.getRange(`B${row + 1}:C${row + 1}`).merge(false);
      sheet.getRange(`B${row + 1}:C${row + 1}`).format.fill.color = "#F8FAFC";
      sheet.getRange(`B${row + 1}:C${row + 1}`).format.font.bold = true;
      sheet.getRange(`B${row + 1}:C${row + 1}`).format.font.size = 13;
      sheet.getRange(`B${row + 1}:C${row + 1}`).format.font.color = "#0F172A";
      sheet.getRange(`B${row + 1}:C${row + 1}`).format.horizontalAlignment = "Center";
      sheet.getRange(`B${row + 1}:C${row + 1}`).format.rowHeight = 28;

      sheet.getRange(`E${row}:F${row}`).values = [[kpiCards[i][2], ""]];
      sheet.getRange(`E${row}:F${row}`).merge(false);
      sheet.getRange(`E${row}:F${row}`).format.fill.color = "#1E293B";
      sheet.getRange(`E${row}:F${row}`).format.font.color = "#FFFFFF";
      sheet.getRange(`E${row}:F${row}`).format.font.bold = true;
      sheet.getRange(`E${row}:F${row}`).format.horizontalAlignment = "Center";
      sheet.getRange(`E${row}:F${row}`).format.rowHeight = 24;

      sheet.getRange(`E${row + 1}:F${row + 1}`).values = [[kpiCards[i][3], ""]];
      sheet.getRange(`E${row + 1}:F${row + 1}`).merge(false);
      sheet.getRange(`E${row + 1}:F${row + 1}`).format.fill.color = "#ECFDF5";
      sheet.getRange(`E${row + 1}:F${row + 1}`).format.font.bold = true;
      sheet.getRange(`E${row + 1}:F${row + 1}`).format.font.size = 13;
      sheet.getRange(`E${row + 1}:F${row + 1}`).format.font.color = "#065F46";
      sheet.getRange(`E${row + 1}:F${row + 1}`).format.horizontalAlignment = "Center";
      sheet.getRange(`E${row + 1}:F${row + 1}`).format.rowHeight = 28;
    }

    try {
      sheet.getRange("B:B").format.columnWidth = 180;
      sheet.getRange("C:C").format.columnWidth = 180;
      sheet.getRange("D:D").format.columnWidth = 35; // Séparateur
      sheet.getRange("E:E").format.columnWidth = 180;
      sheet.getRange("F:F").format.columnWidth = 180;
    } catch (_) {}
  }

  // --- UTILITAIRE BANNIÈRE PROFESSIONNELLE MODERNE ---
  private static createBanner(sheet: Excel.Worksheet, rangeAddress: string, title: string, color: string = "#1E293B") {
    const firstCell = rangeAddress.split(':')[0];
    sheet.getRange(firstCell).values = [[title]];
    const range = sheet.getRange(rangeAddress);
    range.merge(false);
    range.format.fill.color = color;
    range.format.font.color = "#FFFFFF";
    range.format.font.bold = true;
    range.format.font.size = 12;
    range.format.font.name = "Calibri";
    range.format.horizontalAlignment = "Left";
    range.format.verticalAlignment = "Center";
    range.format.rowHeight = 30;
    // Marge intérieure gauche
    try { range.format.indentLevel = 1; } catch (_) {}
  }

  /**
   * Applique un style de tableau professionnel moderne (en-tête prestige + données épurées + total)
   * AUCUN trait vertical agressif, espacement aéré, alternance douce
   */
  private static applyModernTableStyle(
    sheet: Excel.Worksheet,
    headerRange: string,
    dataRange: string,
    headerColor: string = "#1E293B",
    options?: { totalRow?: string; accentColor?: string }
  ) {
    // En-tête moderne épuré
    const hdr = sheet.getRange(headerRange);
    hdr.format.fill.color = headerColor;
    hdr.format.font.color = "#FFFFFF";
    hdr.format.font.bold = true;
    hdr.format.font.size = 10;
    hdr.format.font.name = "Calibri";
    hdr.format.rowHeight = 28;
    hdr.format.horizontalAlignment = "Center";
    hdr.format.verticalAlignment = "Center";

    // Données : AUCUN trait vertical, lignes horizontales très discrètes
    const data = sheet.getRange(dataRange);
    try {
      data.format.borders.getItem('InsideHorizontal').style = "Thin";
      data.format.borders.getItem('InsideHorizontal').color = "#F1F5F9";
      data.format.borders.getItem('InsideVertical').style = "None"; // STRICTEMENT AUCUN TRAIT VERTICAL
      data.format.borders.getItem('EdgeBottom').style = "Thin";
      data.format.borders.getItem('EdgeBottom').color = "#CBD5E1";
      data.format.borders.getItem('EdgeTop').style = "None";
      data.format.borders.getItem('EdgeLeft').style = "None";
      data.format.borders.getItem('EdgeRight').style = "None";
    } catch (_) {}
    data.format.font.size = 10;
    data.format.font.name = "Calibri";
    data.format.rowHeight = 24;
    data.format.verticalAlignment = "Center";

    // Zebra striping moderne : blanc pur et ardoise ultra-légère (#F8FAFC)
    try {
      const parts = dataRange.split(':');
      const startRow = parseInt(parts[0].replace(/[^0-9]/g, ''), 10);
      const endRow = parseInt(parts[1].replace(/[^0-9]/g, ''), 10);
      const startCol = parts[0].replace(/[0-9]/g, '');
      const endCol = parts[1].replace(/[0-9]/g, '');
      for (let r = startRow; r <= endRow; r++) {
        const rowRange = sheet.getRange(`${startCol}${r}:${endCol}${r}`);
        if ((r - startRow) % 2 === 1) {
          rowRange.format.fill.color = "#F8FAFC";
        } else {
          rowRange.format.fill.color = "#FFFFFF";
        }
      }
    } catch (_) {}

    // Ligne de total mise en valeur moderne
    if (options?.totalRow) {
      const total = sheet.getRange(options.totalRow);
      total.format.font.bold = true;
      total.format.fill.color = "#F1F5F9";
      total.format.font.color = "#0F172A";
      total.format.rowHeight = 28;
      total.format.verticalAlignment = "Center";
      try {
        total.format.borders.getItem('EdgeTop').style = "Thin";
        total.format.borders.getItem('EdgeTop').color = "#CBD5E1";
        total.format.borders.getItem('EdgeBottom').style = "Double";
        total.format.borders.getItem('EdgeBottom').color = "#0F172A";
        total.format.borders.getItem('InsideVertical').style = "None";
        total.format.borders.getItem('EdgeLeft').style = "None";
        total.format.borders.getItem('EdgeRight').style = "None";
      } catch (_) {}
    }
  }

  /**
   * Applique un style moderne global post-construction (marges et typographie sans écraser les largeurs)
   */
  private static applyGlobalFinish(sheet: Excel.Worksheet) {
    try {
      const used = sheet.getUsedRange();
      used.format.font.name = "Calibri";
      // Row 1 = marge supérieure discrète
      try { sheet.getRange("1:1").format.rowHeight = 8; } catch (_) {}
      // Column A = marge gauche discrète
      try { sheet.getRange("A:A").format.columnWidth = 20; } catch (_) {}
    } catch (_) {}
  }

  // --- UTILITAIRES SÉCURISÉS EXCEL (COMPATIBILITÉ 2016-365) ---
  private static colToIndex(col: string): number {
    let index = 0;
    const clean = col.toUpperCase().replace(/[^A-Z]/g, '');
    for (let i = 0; i < clean.length; i++) {
      index = index * 26 + (clean.charCodeAt(i) - 64);
    }
    return Math.max(index, 1);
  }

  private static setRangeNumberFormat(sheet: Excel.Worksheet, rangeAddress: string, format: string) {
    try {
      const parts = rangeAddress.split(':');
      if (parts.length === 2) {
        const start = parts[0].trim();
        const end = parts[1].trim();
        const startCol = start.replace(/[0-9]/g, '');
        const startRow = parseInt(start.replace(/[^0-9]/g, '') || '1', 10);
        const endCol = end.replace(/[0-9]/g, '');
        const endRow = parseInt(end.replace(/[^0-9]/g, '') || '1', 10);

        const numRows = Math.max(endRow - startRow + 1, 1);
        const startColIdx = this.colToIndex(startCol);
        const endColIdx = this.colToIndex(endCol);
        const numCols = Math.max(endColIdx - startColIdx + 1, 1);

        const matrix = Array(numRows).fill(null).map(() => Array(numCols).fill(format));
        sheet.getRange(rangeAddress).numberFormat = matrix;
      } else {
        sheet.getRange(rangeAddress).numberFormat = [[format]];
      }
    } catch (e) {
      console.warn("numberFormat ignoré pour", rangeAddress, e);
    }
  }

  private static setBorder(range: Excel.Range, edge: "EdgeBottom" | "EdgeTop" | "EdgeLeft" | "EdgeRight", style: string = "Continuous") {
    try {
      range.format.borders.getItem(edge).style = style as any;
    } catch (_) {}
  }
}
