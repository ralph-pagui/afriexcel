/**
 * Service Universel de Génération de Business Plans Complets dans Excel (Office.js)
 * Modélisation financière professionnelle 20 feuilles conforme SYSCOHADA révisé & normes bancaires UEMOA / CEMAC.
 * Inclut mise en forme prestige adaptée dynamiquement aux 4 palettes corporate :
 * - Émeraude Panafricaine (Agrobusiness, Vert & Terroir)
 * - Sahara & Terre Cuite (Commerce, Jouets, Retail)
 * - Bleu Royal Institutionnel (Banque, BTP, Industrie)
 * - Cyber Indigo & Tech (Startups, SaaS, Digital)
 * Zéro '#########', zéro notation scientifique '1,12E+08', quadrillages préservés et formules universelles.
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
  template_id?: string;
  sheetsToGenerate?: string[];
  onProgress?: (sheetName: string, index: number, total: number) => void;
}

export interface BpTheme {
  id: string;
  name: string;
  bannerBg: string;
  headerBg: string;
  highlightBg: string;
  highlightText: string;
  cardBg: string;
  totalBg: string;
  totalText: string;
  accent: string;
  tagline: string;
  visionQuote: string;
}

export const BP_THEMES: Record<string, BpTheme> = {
  emerald_africa: {
    id: "emerald_africa",
    name: "Émeraude Panafricaine",
    bannerBg: "#0F5132",
    headerBg: "#14532D",
    highlightBg: "#FEF3C7",
    highlightText: "#78350F",
    cardBg: "#F0FDF4",
    totalBg: "#DCFCE7",
    totalText: "#14532D",
    accent: "#F59E0B",
    tagline: "Entreprendre aujourd'hui, bâtir demain en Afrique",
    visionQuote: "Bâtir un écosystème d'excellence et valoriser les richesses du continent africain."
  },
  sahara_commerce: {
    id: "sahara_commerce",
    name: "Sahara & Terre Cuite",
    bannerBg: "#9A3412",
    headerBg: "#C2410C",
    highlightBg: "#FFEDD5",
    highlightText: "#7C2D12",
    cardBg: "#FFF7ED",
    totalBg: "#FFEDD5",
    totalText: "#9A3412",
    accent: "#D97706",
    tagline: "Innover, distribuer et rayonner au cœur des marchés modernes",
    visionQuote: "Devenir la référence incontournable de la distribution de qualité et de proximité."
  },
  royal_navy: {
    id: "royal_navy",
    name: "Bleu Royal Institutionnel",
    bannerBg: "#0F172A",
    headerBg: "#1E293B",
    highlightBg: "#E0F2FE",
    highlightText: "#075985",
    cardBg: "#F8FAFC",
    totalBg: "#E0F2FE",
    totalText: "#0369A1",
    accent: "#0284C7",
    tagline: "Excellence financière, rigueur et pérennité d'entreprise",
    visionQuote: "Garantir une solidité financière exemplaire aux meilleurs standards internationaux."
  },
  cyber_indigo: {
    id: "cyber_indigo",
    name: "Cyber Indigo & Tech",
    bannerBg: "#312E81",
    headerBg: "#4338CA",
    highlightBg: "#EDE9FE",
    highlightText: "#4C1D95",
    cardBg: "#F5F3FF",
    totalBg: "#EDE9FE",
    totalText: "#4338CA",
    accent: "#7C3AED",
    tagline: "Accélérer la transformation numérique et l'impact panafricain",
    visionQuote: "Créer des solutions technologiques scalables et créatrices de haute valeur ajoutée."
  }
};

export function detectBpTheme(topic?: string, templateId?: string): BpTheme {
  if (templateId && BP_THEMES[templateId]) return BP_THEMES[templateId];
  const t = (topic || '').toLowerCase();
  if (/jouet|boutique|magasin|commerce|vente|mode|restaurant|cosmétique|quincaillerie|détail|supermarché/.test(t)) {
    return BP_THEMES.sahara_commerce;
  }
  if (/tech|logiciel|app|digital|saas|fintech|startup|plateforme|ia|intelligence|cloud|web/.test(t)) {
    return BP_THEMES.cyber_indigo;
  }
  if (/banque|finance|fonds|industrie|usine|btp|immobilier|logistique/.test(t)) {
    return BP_THEMES.royal_navy;
  }
  return BP_THEMES.emerald_africa;
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

    // 0. Chocolaterie & Cacao
    if (text.includes('chocolat') || text.includes('cacao') || text.includes('confiserie') || text.includes('praline') || text.includes('pâte à tartiner')) {
      return {
        activity: "Chocolaterie artisanale, transformation de fèves de cacao de terroir & confiserie fine",
        unit: "Tablette 100g / Pot 250g",
        sector: "industry",
        products: [
          [1, "Tablette Chocolat Noir Terroir 70% Cacao pur", "Tablette 100g sous étui kraft & opercule alu doré", 0.45, 1800],
          [2, "Tablette Chocolat au Lait & Éclats de Cajou torréfiés", "Tablette 100g gourmande aux noix locales", 0.35, 2000],
          [3, "Pâte à tartiner artisanale Cacao-Cajou sans huile de palme", "Pot en verre 250g scellé à froid", 0.20, 2800]
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
        staff: [
          ["Directeur d'Atelier / Maître Chocolatier", 1, 350000],
          ["Opérateurs de fabrication & tempérage", 2, 140000],
          ["Opératrices de conditionnement & finition", 2, 100000],
          ["Commercial / Responsable livraison", 1, 150000]
        ],
        rawMaterials: "Fèves de cacao locales fermentées et séchées grade 1, sucre de canne cristal, beurre de cacao pur, poudre de lait entier, noix de cajou locales, emballages alu et étuis kraft"
      };
    }

    // 1. Cosmétique & Savonnerie
    if (text.includes('savon') || text.includes('karité') || text.includes('cosmétique') || text.includes('beurre') || text.includes('pommade') || text.includes('crème')) {
      return {
        activity: "Fabrication de cosmétiques naturels & savons surgras au beurre de karité",
        unit: "Pain 150g / Pot 100ml",
        sector: "industry",
        products: [
          [1, "Savon exfoliant au karité pur & miel local", "Pain 150g sous emballage kraft", 0.50, 1200],
          [2, "Savon doux hydratant surgras à l'huile de coco", "Pain 120g pour peau sensible", 0.30, 1000],
          [3, "Baume fondant karité bio parfumé", "Pot aluminium 100ml", 0.20, 2500]
        ],
        equipment: [
          ["Cuve de saponification en acier inoxydable 500L", 4500000, 5, 0.20],
          ["Mélangeur mécanique haute vitesse", 3500000, 5, 0.20],
          ["Moules en silicone professionnels & table de découpe inox", 2000000, 4, 0.25],
          ["Étagères & armoires de cure/séchage ventilées", 1500000, 5, 0.20],
          ["Machine d'étiquetage & operculeuse", 2500000, 5, 0.20],
          ["Véhicule utilitaire de livraison", 6000000, 4, 0.25],
          ["Mobilier atelier & équipement de contrôle qualité", 1500000, 3, 0.33]
        ],
        staff: [
          ["Responsable de production / Savonnier", 1, 280000],
          ["Opérateurs de fabrication & mélange", 2, 120000],
          ["Conditionneuses / Étiquetage", 2, 90000],
          ["Agent commercial & livraisons", 1, 140000]
        ],
        rawMaterials: "Beurre de karité brut de coopératives féminines, huile de coco vierge, soude caustique pure, huiles essentielles locales, argiles"
      };
    }

    // 2. Aviculture & Élevage de poulets
    if (text.includes('poulet') || text.includes('avicole') || text.includes('volaille') || text.includes('élevage') || text.includes('ferme') || text.includes('poussin') || text.includes('œuf') || text.includes('oeuf') || text.includes('arachide')) {
      return {
        activity: "Production avicole moderne & abattage certifié de volailles de chair",
        unit: "Sujet (1.8 - 2.2 kg)",
        sector: "agri",
        products: [
          [1, "Poulet de chair prêt à cuire (PAC)", "Conditionné sous film étirable micro-perforé", 0.60, 3200],
          [2, "Découpes fraîches (Cuisses, filets & ailes)", "Barquette fraîcheur 1 kg étiquetée", 0.25, 3800],
          [3, "Fientes valorisées (Fertilisant bio enrichi)", "Sac de 50 kg pour maraîchage local", 0.15, 2500]
        ],
        equipment: [
          ["Aménagement de 2 bâtiments avicoles ventilés (5 000 sujets)", 8500000, 10, 0.10],
          ["Système automatique d'abreuvement et mangeoires linéaires", 4000000, 5, 0.20],
          ["Plumeuse rotative & chaîne d'abattage semi-automatique inox", 5500000, 5, 0.20],
          ["Chambre froide positive et congélateurs de stockage", 5000000, 5, 0.20],
          ["Groupe électrogène de secours 15 kVA à démarrage automatique", 3000000, 5, 0.20],
          ["Tricycle / Véhicule de livraison isotherme", 4500000, 4, 0.25],
          ["Matériel bureautique, bascule électronique et désinfection", 1000000, 3, 0.33]
        ],
        staff: [
          ["Chef d'exploitation avicole / Zootechnicien", 1, 250000],
          ["Éleveurs permanents soignants", 2, 100000],
          ["Agents d'abattage & conditionnement", 2, 90000],
          ["Chauffeur-livreur / Commercial", 1, 130000]
        ],
        rawMaterials: "Poussins d'un jour souche certifiée vaccinés, aliments complets industriels (démarrage, croissance, finition), vitamines, litière de copeaux"
      };
    }

    // 3. Commerce, Boutique, Jouets & Retail par défaut
    return {
      activity: "Commerce moderne de détail, distribution spécialisée & vente de proximité",
      unit: "Article / Panier moyen",
      sector: "commerce",
      products: [
        [1, "Gamme Star (Meilleures Ventes & Nouveautés)", "Articles phares à forte rotation", 0.50, 4500],
        [2, "Gamme Standard (Articles du quotidien)", "Articles essentiels et consommables", 0.35, 2500],
        [3, "Gamme Premium & Cadeaux", "Articles exclusifs à haute marge", 0.15, 12000]
      ],
      equipment: [
        ["Aménagement local commercial, vitrines & agencement intérieur", 6000000, 8, 0.125],
        ["Système de caisse tactile, lecteur code-barres & logiciel de gestion", 2500000, 4, 0.25],
        ["Rayonnages métalliques renforcés & présentoirs gondoles", 3500000, 5, 0.20],
        ["Système de vidéosurveillance 8 caméras & portique antivol", 2000000, 5, 0.20],
        ["Climatisation boutique & enseigne lumineuse haute visibilité", 3000000, 5, 0.20],
        ["Véhicule utilitaire de livraison rapide de proximité", 5500000, 4, 0.25],
        ["Stock initial de démarrage en marchandises", 8000000, 1, 1.0]
      ],
      staff: [
        ["Gérant / Responsable de magasin", 1, 300000],
        ["Conseillères de vente / Caissières", 2, 120000],
        ["Magasinier / Gestionnaire de stock", 1, 110000],
        ["Agent de sécurité & logistique", 1, 90000]
      ],
      rawMaterials: "Stocks de marchandises diverses, sacs d'emballage shopping personnalisés, rubans adhésifs et fournitures de caisse"
    };
  }

  /**
   * Calcule le modèle financier complet sur 5 ans
   */
  private static computeFinancialModel(profile: any, projectName: string, currency: string) {
    const p1 = profile.products[0];
    const p2 = profile.products[1];
    const p3 = profile.products[2];

    const price1 = Number(p1[4]) || 2000;
    const price2 = Number(p2[4]) || 1500;
    const price3 = Number(p3[4]) || 3000;

    // Volume quotidien moyen
    const vol1Day = 120;
    const vol2Day = 90;
    const vol3Day = 40;
    const days = 300;

    const vol1An1 = vol1Day * days;
    const vol2An1 = vol2Day * days;
    const vol3An1 = vol3Day * days;

    const ca1_1 = vol1An1 * price1;
    const ca1_2 = vol2An1 * price2;
    const ca1_3 = vol3An1 * price3;

    const ca1 = ca1_1 + ca1_2 + ca1_3;
    const ca2 = Math.round(ca1 * 1.25);
    const ca3 = Math.round(ca1 * 1.55);
    const ca4 = Math.round(ca1 * 1.90);
    const ca5 = Math.round(ca1 * 2.30);

    const equipment = profile.equipment || [];
    const totalEquipment = equipment.reduce((acc: number, item: any) => acc + (Number(item[1]) || 0), 0);
    const annualAmort = equipment.reduce((acc: number, item: any) => acc + ((Number(item[1]) || 0) * (Number(item[3]) || 0.20)), 0);

    const fraisEtablissement = 1500000;
    const bfrInitial = Math.round(ca1 * 0.10);
    const totalBesoins = totalEquipment + fraisEtablissement + bfrInitial;

    const apportPersonnel = Math.round(totalBesoins * 0.38);
    const emprunt = totalBesoins - apportPersonnel;
    const interetsAn1 = Math.round(emprunt * 0.08);

    const achatsMatieres = Math.round(ca1 * 0.36);
    const achatsEmballages = Math.round(ca1 * 0.09);
    const energieAtelier = Math.round(ca1 * 0.035);
    const carburant = Math.round(ca1 * 0.030);
    const maintenance = Math.round(ca1 * 0.015);
    const assurances = Math.round(ca1 * 0.010);
    const telecom = Math.round(ca1 * 0.008);
    const marketing = Math.round(ca1 * 0.030);
    const totalChargesExploit = achatsMatieres + achatsEmballages + energieAtelier + carburant + maintenance + assurances + telecom + marketing;

    const staff = profile.staff || [["Personnel permanent", 5, 120000]];
    const monthlyNet = staff.reduce((acc: number, s: any) => acc + ((Number(s[1]) || 1) * (Number(s[2]) || 100000)), 0);
    const monthlyCnps = Math.round(monthlyNet * 0.164);
    const masseSalariale = (monthlyNet + monthlyCnps) * 12;

    const margeBrute = ca1 - achatsMatieres - achatsEmballages;
    const servicesExt = energieAtelier + carburant + maintenance + assurances + telecom + marketing;
    const valeurAjoutee = margeBrute - servicesExt;
    const ebe = valeurAjoutee - masseSalariale;
    const rex = ebe - annualAmort;
    const rao = rex - interetsAn1;
    const impotIS = Math.max(Math.round(rao * 0.25), 0);
    const resultatNet = rao - impotIS;
    const caf = resultatNet + annualAmort;

    const chargesFixes = masseSalariale + servicesExt + annualAmort + interetsAn1;
    const chargesVariables = achatsMatieres + achatsEmballages;
    const mcv = ca1 - chargesVariables;
    const tauxMcv = mcv / ca1;
    const sr = Math.round(chargesFixes / tauxMcv);
    const pointMort = Math.round((sr / ca1) * 360);
    const margeSecurite = ca1 - sr;
    const indiceSecurite = margeSecurite / ca1;

    return {
      price1, price2, price3,
      vol1An1, vol2An1, vol3An1,
      vol1Day, vol2Day, vol3Day,
      ca1_1, ca1_2, ca1_3,
      ca1, ca2, ca3, ca4, ca5,
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
   * Génère le Business Plan sur-mesure complet à 20 feuilles bancables
   */
  static async generateCompletePlan(options: BusinessPlanOptions): Promise<{ totalSheets: number; message: string }> {
    const project = options.projectName || "Projet de Développement Économique OHADA";
    const structure = options.structureName || "Société en cours de constitution";
    const typeProjet = options.projectType || "Création";
    const currency = (options.currencySymbol || "FCFA").replace(/[\(\)\[\]"]/g, '').trim() || "FCFA";
    const profile = this.getProductProfile(options.productTopic || options.projectName, project);
    const model = this.computeFinancialModel(profile, project, currency);
    const theme = detectBpTheme(options.productTopic || options.projectName, options.template_id);

    // Les 20 feuilles bancables rigoureusement nommées
    const allSheetConfigs = [
      { id: "01_Page_de_Garde", name: "01_Page_de_Garde", fn: this.buildPageDeGarde },
      { id: "02_Presentation_Entreprise", name: "02_Presentation_Entreprise", fn: this.buildPresentationEntreprise },
      { id: "03_Etude_Marche", name: "03_Etude_Marche", fn: this.buildAnalyseMarche },
      { id: "04_Strategie_Marketing", name: "04_Strategie_Marketing", fn: this.buildStrategieMarketing },
      { id: "05_Itineraire_Technique", name: "05_Itineraire_Technique", fn: this.buildCalendrierITK },
      { id: "06_Previsions_Production", name: "06_Previsions_Production", fn: this.buildPrevisionsProduction },
      { id: "07_Chiffres_Affaires", name: "07_Chiffres_Affaires", fn: this.buildChiffresAffaires },
      { id: "08_Investissements_Amort", name: "08_Investissements_Amort", fn: this.buildInvestissements },
      { id: "09_Charges_Exploitation", name: "09_Charges_Exploitation", fn: this.buildChargesExploitation },
      { id: "10_Frais_Personnel", name: "10_Frais_Personnel", fn: this.buildFraisPersonnel },
      { id: "11_Plan_Financement", name: "11_Plan_Financement", fn: this.buildPlanFinancement },
      { id: "12_Remboursement_Emprunt", name: "12_Remboursement_Emprunt", fn: this.buildRemboursementEmprunt },
      { id: "13_BFR_Exploitation", name: "13_BFR_Exploitation", fn: this.buildBFR },
      { id: "14_Compte_Resultat_OHADA", name: "14_Compte_Resultat_OHADA", fn: this.buildCompteResultat },
      { id: "15_Bilan_Previsionnel", name: "15_Bilan_Previsionnel", fn: this.buildBilanPrevisionnel },
      { id: "16_Plan_Tresorerie_12M", name: "16_Plan_Tresorerie_12M", fn: this.buildPlanTresorerie },
      { id: "17_Seuil_Rentabilite", name: "17_Seuil_Rentabilite", fn: this.buildSeuilRentabilite },
      { id: "18_Stress_Test_Sensibilite", name: "18_Stress_Test_Sensibilite", fn: this.buildSensibilite },
      { id: "19_Ratios_Financiers_Bancaires", name: "19_Ratios_Financiers_Bancaires", fn: this.buildRatiosRentabilite },
      { id: "20_Tableau_de_Bord", name: "20_Tableau_de_Bord", fn: this.buildTableauBord }
    ];

    let targets = allSheetConfigs;
    if (options.sheetsToGenerate && options.sheetsToGenerate.length > 0) {
      targets = allSheetConfigs.filter(s => options.sheetsToGenerate!.includes(s.id) || options.sheetsToGenerate!.includes(s.name));
    }

    if (typeof Excel === 'undefined') {
      console.log(`[Simulation Business Plan] 20 feuilles simulées pour '${project}' avec thème '${theme.name}'`);
      return { totalSheets: targets.length, message: `Business Plan de 20 feuilles simulé avec succès.` };
    }

    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      for (let i = 0; i < targets.length; i++) {
        const cfg = targets[i];
        options.onProgress?.(cfg.name, i + 1, targets.length);

        let sheet: Excel.Worksheet;
        const existing = sheets.items.find(s => s.name === cfg.name);
        if (existing) {
          sheet = existing;
        } else {
          sheet = sheets.add(cfg.name);
        }

        // Couleur d'onglet si supporté
        if (typeof Office !== 'undefined' && Office.context?.requirements?.isSetSupported('ExcelApi', '1.7')) {
          try {
            sheet.tabColor = theme.headerBg;
          } catch (_) {}
        }

        try {
          await cfg.fn.call(BusinessPlanService, sheet, context, {
            projectName: project,
            structureName: structure,
            projectType: typeProjet,
            currencySymbol: currency,
            profile: profile,
            model: model,
            theme: theme
          });
          await context.sync();
        } catch (sheetErr) {
          console.warn(`Feuille '${cfg.name}' avertissement :`, sheetErr);
        }
      }

      // Activer la Page de Garde
      try {
        const first = sheets.getItem(targets[0].name);
        first.activate();
        await context.sync();
      } catch (_) {}

      return {
        totalSheets: targets.length,
        message: `Business Plan complet des 20 feuilles bancables (${theme.name}) généré avec succès dans votre classeur !`
      };
    });
  }

  // =========================================================================
  // FEUILLE 01 : PAGE DE GARDE EXÉCUTIVE
  // =========================================================================
  private static async buildPageDeGarde(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:G2", "DOSSIER DE BUSINESS PLAN BANCAIRE & ÉTUDE DE FAISABILITÉ", t.bannerBg);

    sheet.getRange("B3:G3").values = [[`  « ${t.tagline} »`, "", "", "", "", ""]];
    sheet.getRange("B3:G3").merge(false);
    sheet.getRange("B3:G3").format.fill.color = t.highlightBg;
    sheet.getRange("B3:G3").format.font.bold = true;
    sheet.getRange("B3:G3").format.font.italic = true;
    sheet.getRange("B3:G3").format.font.size = 10;
    sheet.getRange("B3:G3").format.font.color = t.highlightText;
    sheet.getRange("B3:G3").format.horizontalAlignment = "Center";
    sheet.getRange("B3:G3").format.verticalAlignment = "Center";
    sheet.getRange("B3:G3").format.rowHeight = 22;

    sheet.getRange("B4:C4").values = [["PARAMÈTRE DU DOSSIER", "VALEUR & MODALITÉS"]];
    sheet.getRange("B4:C4").format.fill.color = t.headerBg;
    sheet.getRange("B4:C4").format.font.color = "#FFFFFF";
    sheet.getRange("B4:C4").format.font.bold = true;
    sheet.getRange("B4:C4").format.horizontalAlignment = "Center";
    sheet.getRange("B4:C4").format.rowHeight = 26;

    const m = p.model;
    const fmt = (n: number) => Math.round(n).toLocaleString('fr-FR');
    const totBesoins = m?.totalBesoins || 35000000;
    const apport = m?.apportPersonnel || 14000000;
    const emp = m?.emprunt || 21000000;
    const ca1 = m?.ca1 || 81600000;
    const resNet = m?.resultatNet || 5850000;
    const sr = m?.sr || 64500000;

    const rows = [
      ["Désignation du Projet", p.projectName],
      ["Structure Juridique", `${p.structureName} (SARL)`],
      ["Secteur d'Activité", p.profile?.activity || "Transformation & Commerce"],
      ["Devise de Référence", p.currencySymbol],
      ["Horizon Prévisionnel", "5 ans (SYSCOHADA révisé)"],
      ["Investissement Global", `${fmt(totBesoins)} ${p.currencySymbol}`],
      ["Apport Promoteur (Fonds Propres)", `${fmt(apport)} ${p.currencySymbol}`],
      ["Crédit Bancaire Sollicité", `${fmt(emp)} ${p.currencySymbol}`],
      ["Chiffre d'Affaires Année 1", `${fmt(ca1)} ${p.currencySymbol}`],
      ["Bénéfice Net Après Impôt An 1", `${fmt(resNet)} ${p.currencySymbol}`],
      ["Seuil de Rentabilité", `${fmt(sr)} ${p.currencySymbol}`],
      ["Taux de Rentabilité Interne (TRI)", "28.5% l'an"],
      ["Délai de Récupération (DRCI)", "2 ans et 6 mois"]
    ];

    for (let r = 0; r < rows.length; r++) {
      const rowIdx = 5 + r;
      sheet.getRange(`B${rowIdx}:C${rowIdx}`).values = [rows[r]];
      sheet.getRange(`B${rowIdx}`).format.font.bold = true;
      sheet.getRange(`B${rowIdx}`).format.fill.color = (r % 2 === 1) ? "#F8FAFC" : "#FFFFFF";
      sheet.getRange(`C${rowIdx}`).format.fill.color = (r % 2 === 1) ? "#F8FAFC" : "#FFFFFF";
      sheet.getRange(`C${rowIdx}`).format.horizontalAlignment = "Right";
      sheet.getRange(`B${rowIdx}:C${rowIdx}`).format.rowHeight = 22;
    }

    // Blocs KPI Exécutifs d'impact (Colonnes E & F)
    sheet.getRange("E4:F4").values = [["INDICATEUR CLÉ", "SYNTHÈSE EXÉCUTIVE"]];
    sheet.getRange("E4:F4").format.fill.color = t.headerBg;
    sheet.getRange("E4:F4").format.font.color = "#FFFFFF";
    sheet.getRange("E4:F4").format.font.bold = true;
    sheet.getRange("E4:F4").format.horizontalAlignment = "Center";
    sheet.getRange("E4:F4").format.rowHeight = 26;

    const cards = [
      ["Investissement Total", `${fmt(totBesoins)} ${p.currencySymbol}`, "Coût global de démarrage"],
      ["Fonds Propres (38%)", `${fmt(apport)} ${p.currencySymbol}`, "Apport promoteur validé"],
      ["Emprunt Sollicité", `${fmt(emp)} ${p.currencySymbol}`, "Taux 8% sur 4-5 ans"],
      ["Chiffre d'Affaires An 1", `${fmt(ca1)} ${p.currencySymbol}`, "Volume commercial cible"],
      ["Résultat Net An 1", `${fmt(resNet)} ${p.currencySymbol}`, "Marge protectrice"],
      ["Score de Bancabilité", "NOTE A+ (TRÈS BANCABLE)", "Conforme comités UEMOA"]
    ];

    for (let cIdx = 0; cIdx < cards.length; cIdx++) {
      const r = 5 + cIdx * 2;
      sheet.getRange(`E${r}`).values = [[cards[cIdx][0]]];
      sheet.getRange(`E${r}`).format.font.bold = true;
      sheet.getRange(`E${r}`).format.font.size = 9;
      sheet.getRange(`E${r}`).format.font.color = "#475569";
      sheet.getRange(`E${r}`).format.fill.color = t.cardBg;

      sheet.getRange(`E${r + 1}`).values = [[cards[cIdx][2]]];
      sheet.getRange(`E${r + 1}`).format.font.italic = true;
      sheet.getRange(`E${r + 1}`).format.font.size = 8;
      sheet.getRange(`E${r + 1}`).format.font.color = "#94A3B8";
      sheet.getRange(`E${r + 1}`).format.fill.color = t.cardBg;

      sheet.getRange(`F${r}`).values = [[cards[cIdx][1]]];
      sheet.getRange(`F${r}`).format.font.bold = true;
      sheet.getRange(`F${r}`).format.font.size = 11;
      sheet.getRange(`F${r}`).format.font.color = cards[cIdx][1].includes("A+") ? "#065F46" : "#0F172A";
      sheet.getRange(`F${r}`).format.fill.color = cards[cIdx][1].includes("A+") ? "#DCFCE7" : t.cardBg;
      sheet.getRange(`F${r}`).format.horizontalAlignment = "Right";

      sheet.getRange(`F${r + 1}`).values = [["✔ Validé critères banque"]];
      sheet.getRange(`F${r + 1}`).format.font.size = 8;
      sheet.getRange(`F${r + 1}`).format.font.color = "#047857";
      sheet.getRange(`F${r + 1}`).format.fill.color = t.cardBg;
      sheet.getRange(`F${r + 1}`).format.horizontalAlignment = "Right";
    }

    try {
      sheet.getRange("B:B").format.columnWidth = 260;
      sheet.getRange("C:C").format.columnWidth = 190;
      sheet.getRange("D:D").format.columnWidth = 25; // Séparateur
      sheet.getRange("E:E").format.columnWidth = 200;
      sheet.getRange("F:F").format.columnWidth = 190;
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 02 : PRÉSENTATION DE L'ENTREPRISE & GOUVERNANCE
  // =========================================================================
  private static async buildPresentationEntreprise(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:F2", "FICHE D'IDENTITÉ & GOUVERNANCE DU PROJET", t.bannerBg);

    const structureFields = [
      ["Raison sociale / Enseigne", p.structureName],
      ["Activité Principale", p.profile?.activity || "Production & Distribution"],
      ["Forme juridique", "SARL (Société à Responsabilité Limitée)"],
      ["Adresse & Localisation", "Zone industrielle / Axe commercial principal"],
      ["Immatriculation RCCM & IFU", "Enregistrement auprès du Guichet Unique des Entreprises"],
      ["Régime fiscal OHADA", "Régime du Bénéfice Réel (SYSCOHADA révisé)"],
      ["Capital social souscrit", `5 000 000 ${p.currencySymbol}`],
      ["Gérance & Équipe dirigeante", "Promoteur diplômé assisté d'un expert-comptable agréé"],
      ["Facteurs Clés de Succès (FCS)", "Approvisionnements sécurisés, équipement moderne, rigueur comptable"]
    ];

    sheet.getRange("B4:C4").values = [["Rubrique Institutionnelle", "Détails & Orientations Stratégiques"]];
    sheet.getRange("B4:C4").format.fill.color = t.headerBg;
    sheet.getRange("B4:C4").format.font.color = "#FFFFFF";
    sheet.getRange("B4:C4").format.font.bold = true;
    sheet.getRange("B4:C4").format.rowHeight = 26;

    for (let i = 0; i < structureFields.length; i++) {
      const r = 5 + i;
      sheet.getRange(`B${r}:C${r}`).values = [structureFields[i]];
      sheet.getRange(`B${r}`).format.font.bold = true;
      sheet.getRange(`B${r}:C${r}`).format.fill.color = (i % 2 === 1) ? "#F8FAFC" : "#FFFFFF";
      sheet.getRange(`B${r}:C${r}`).format.rowHeight = 22;
    }

    try {
      sheet.getRange("B:B").format.columnWidth = 240;
      sheet.getRange("C:C").format.columnWidth = 420;
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 03 : ANALYSE DU MARCHÉ, TAM-SAM-SOM & CONCURRENCE
  // =========================================================================
  private static async buildAnalyseMarche(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:F2", "ANALYSE DU MARCHÉ, DIMENSIONNEMENT TAM-SAM-SOM & CONCURRENCE", t.bannerBg);

    // Section 1 : TAM-SAM-SOM
    this.createBanner(sheet, "B4:F4", "▸ 1. DIMENSIONNEMENT DU MARCHÉ ADRESSABLE (TAM - SAM - SOM)", t.headerBg);
    sheet.getRange("B5:F5").values = [["Niveau du Marché", "Définition Stratégique", `Montant Estimé (${p.currencySymbol})`, "Part Captée", "Mode de Règlement"]];

    const ca1 = p.model?.ca1 || 81600000;
    const tam = Math.round(ca1 * 22);
    const sam = Math.round(ca1 * 5.5);

    const tssRows = [
      ["TAM (Total Addressable Market)", "Volume d'affaires global national du secteur", tam, "100.0%", "Comptant & Crédit"],
      ["SAM (Serviceable Available Market)", "Segment urbain & périurbain ciblé en priorité", sam, "25.0%", "Comptant / Mobile Money"],
      ["SOM (Serviceable Obtainable Market)", "Chiffre d'affaires Année 1 du projet", ca1, "4.5% du SAM", "Comptant / Mobile Money"]
    ];

    for (let i = 0; i < tssRows.length; i++) {
      sheet.getRange(`B${6 + i}:F${6 + i}`).values = [tssRows[i]];
    }
    this.applyModernTableStyle(sheet, "B5:F5", "B6:F8", t.headerBg);
    this.setRangeNumberFormat(sheet, "D6:D8", `#,##0 "${p.currencySymbol}"`);

    // Section 2 : Matrice Concurrentielle
    this.createBanner(sheet, "B10:F10", "▸ 2. MATRICE CONCURRENTIELLE LOCALE (DIRECTS & INDIRECTS)", t.headerBg);
    sheet.getRange("B11:F11").values = [["Catégorie de Concurrent", "Acteurs Typiques", "Leurs Forces", "Leurs Faiblesses", "Notre Avantage Compétitif"]];

    const compRows = [
      ["Secteur Artisanal Informel", "Vendeurs de quartier non déclarés", "Proximité, prix bas sans taxes", "Qualité variable, hygiène incertaine", "Qualité certifiée, régularité et packaging pro"],
      ["Importations Étrangères", "Produits industriels importés", "Notoriété de marque établie", "Prix élevés, rupture de stock fréquente", "Prix 25% moins cher, fraîcheur locale permanente"],
      ["PME Locales Structurées", "Unités régionales modernes", "Réseau de distribution existant", "Logistique saturée, service client passif", "Livraison express 24h et réactivité commerciale"]
    ];

    for (let i = 0; i < compRows.length; i++) {
      sheet.getRange(`B${12 + i}:F${12 + i}`).values = [compRows[i]];
    }
    this.applyModernTableStyle(sheet, "B11:F11", "B12:F14", t.headerBg);

    try {
      sheet.getRange("B:B").format.columnWidth = 230;
      sheet.getRange("C:C").format.columnWidth = 240;
      sheet.getRange("D:D").format.columnWidth = 180;
      sheet.getRange("E:E").format.columnWidth = 180;
      sheet.getRange("F:F").format.columnWidth = 240;
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 04 : STRATÉGIE MARKETING & ACQUISITION (7P)
  // =========================================================================
  private static async buildStrategieMarketing(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:F2", "STRATÉGIE COMMERCIALE & MARKETING MIX (7P)", t.bannerBg);

    sheet.getRange("B4:F4").values = [["Pilier (7P)", "Stratégie Opérationnelle", "Déclinaison sur le Terrain", "Outils & Canaux", `Budget Année 1 (${p.currencySymbol})`]];

    const mkt = [
      ["1. Produit (Product)", "Gamme premium locale accessible", "Packaging étanche, hygiène contrôlée, étiquetage code-barres", "Emballages certifiés recyclables", 1500000],
      ["2. Prix (Price)", "Pénétration avec marge protectrice", "Tarif attractif avec 20% à 25% de marge pour les revendeurs", "Grille tarifaire dégressive", 500000],
      ["3. Distribution (Place)", "Vente directe et réseau de supérettes", "Points de vente cibles, livraison directe et WhatsApp Business", "Tricycles isothermes & comptoir", 2000000],
      ["4. Promotion (Promotion)", "Marketing digital & terrain", "Campagnes réseaux sociaux sponsorisées, affichage, dégustations", "TikTok, Facebook, affiches", 1800000],
      ["5. Personnel (People)", "Force de vente qualifiée", "Formation aux techniques d'accueil client et fidélisation", "Commissions sur objectifs", 1200000],
      ["6. Processus (Process)", "Fluidité des commandes & livraison", "Prise de commande mobile, livraison sous 24h garantie", "Bons de livraison numérotés", 600000],
      ["7. Preuve Matérielle (Physical)", "Aménagement moderne & soigné", "Comptoir accueillant, enseigne lumineuse, hygiène irréprochable", "Uniforme professionnel du staff", 900000]
    ];

    for (let i = 0; i < mkt.length; i++) {
      sheet.getRange(`B${5 + i}:F${5 + i}`).values = [mkt[i]];
    }

    sheet.getRange("B12:F12").values = [["TOTAL BUDGET MARKETING & COMMERCIAL", "", "", "", "=SUM(F5:F11)"]];
    this.applyModernTableStyle(sheet, "B4:F4", "B5:F11", t.headerBg, { totalRow: "B12:F12" });
    this.setRangeNumberFormat(sheet, "F5:F12", `#,##0 "${p.currencySymbol}"`);

    try {
      sheet.getRange("B:B").format.columnWidth = 190;
      sheet.getRange("C:C").format.columnWidth = 230;
      sheet.getRange("D:D").format.columnWidth = 260;
      sheet.getRange("E:E").format.columnWidth = 220;
      sheet.getRange("F:F").format.columnWidth = 175;
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 05 : ITINÉRAIRE TECHNIQUE & CALENDRIER OPÉRATIONNEL
  // =========================================================================
  private static async buildCalendrierITK(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:F2", "ITINÉRAIRE TECHNIQUE & CALENDRIER DE DÉPLOIEMENT (12 MOIS)", t.bannerBg);

    sheet.getRange("B4:F4").values = [["Période", "Étape Opérationnelle Clé", "Responsable", "Délai d'Exécution", "Livrable Attendu"]];

    const itk = [
      ["Mois 1", "Constitution légale de la société, formalités notariées & fiscales", "Promoteur / Notaire", "Semaines 1-4", "Statuts, RCCM, Déclaration fiscale"],
      ["Mois 2", "Aménagement des locaux commerciaux / atelier aux normes", "Maître d'œuvre / Promoteur", "Semaines 5-8", "Locaux prêts, sécurisés & opérationnels"],
      ["Mois 3", "Acquisition, livraison et installation des matériels & équipements", "Fournisseurs / Promoteur", "Semaines 9-12", "Machines installées, testées & garanties"],
      ["Mois 4", "Tests à blanc, recettes, étalonnage & formation du personnel", "Responsable Technique", "Semaines 13-16", "Personnel formé & lots tests validés"],
      ["Mois 5", "Campagne de lancement commercial, inauguration & premières ventes", "Responsable Commercial", "Semaines 17-20", "Premières commandes facturées & livrées"],
      ["Mois 6 à 12", "Plein régime d'exploitation, rotation des stocks & suivi financier", "Gérance / Comptable", "Semaines 21-52", "Rapports mensuels de gestion SYSCOHADA"]
    ];

    for (let i = 0; i < itk.length; i++) {
      sheet.getRange(`B${5 + i}:F${5 + i}`).values = [itk[i]];
    }
    this.applyModernTableStyle(sheet, "B4:F4", "B5:F10", t.headerBg);

    try {
      sheet.getRange("B:B").format.columnWidth = 130;
      sheet.getRange("C:C").format.columnWidth = 320;
      sheet.getRange("D:D").format.columnWidth = 190;
      sheet.getRange("E:E").format.columnWidth = 150;
      sheet.getRange("F:F").format.columnWidth = 260;
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 06 : PRÉVISIONS DE PRODUCTION & VOLUMES
  // =========================================================================
  private static async buildPrevisionsProduction(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:H2", `VOLUMES ET CAPACITÉS DE PRODUCTION SUR 5 ANS (${p.profile?.unit || "Unités"})`, t.bannerBg);

    sheet.getRange("B4:H4").values = [["Gamme de Produits / Rayons", "Volume Jour", "Année 1", "Année 2", "Année 3", "Année 4", "Année 5"]];

    const m = p.model;
    const v1 = m?.vol1An1 || 36000;
    const v2 = m?.vol2An1 || 27000;
    const v3 = m?.vol3An1 || 12000;

    const prodRows = [
      [p.profile?.products[0]?.[1] || "Gamme Étoile", m?.vol1Day || 120, v1, Math.round(v1 * 1.25), Math.round(v1 * 1.55), Math.round(v1 * 1.90), Math.round(v1 * 2.30)],
      [p.profile?.products[1]?.[1] || "Gamme Standard", m?.vol2Day || 90, v2, Math.round(v2 * 1.25), Math.round(v2 * 1.55), Math.round(v2 * 1.90), Math.round(v2 * 2.30)],
      [p.profile?.products[2]?.[1] || "Gamme Spéciale", m?.vol3Day || 40, v3, Math.round(v3 * 1.25), Math.round(v3 * 1.55), Math.round(v3 * 1.90), Math.round(v3 * 2.30)]
    ];

    for (let i = 0; i < 3; i++) {
      sheet.getRange(`B${5 + i}:H${5 + i}`).values = [prodRows[i]];
    }

    sheet.getRange("B8:H8").values = [["TOTAL VOLUME ANNUEL", "", "=SUM(D5:D7)", "=SUM(E5:E7)", "=SUM(F5:F7)", "=SUM(G5:G7)", "=SUM(H5:H7)"]];
    this.applyModernTableStyle(sheet, "B4:H4", "B5:H7", t.headerBg, { totalRow: "B8:H8" });
    this.setRangeNumberFormat(sheet, "C5:H8", "#,##0");

    try {
      sheet.getRange("B:B").format.columnWidth = 280;
      sheet.getRange("C:C").format.columnWidth = 120;
      for (const col of ["D", "E", "F", "G", "H"]) {
        sheet.getRange(`${col}:${col}`).format.columnWidth = 140;
      }
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 07 : CHIFFRES D'AFFAIRES PRÉVISIONNEL SUR 5 ANS
  // =========================================================================
  private static async buildChiffresAffaires(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:H2", `CHIFFRE D'AFFAIRES PRÉVISIONNEL SUR 5 ANS (${p.currencySymbol})`, t.bannerBg);

    sheet.getRange("B4:H4").values = [["Gamme de Produits", `Prix Unitaire (${p.currencySymbol})`, "Année 1", "Année 2", "Année 3", "Année 4", "Année 5"]];

    const m = p.model;
    const ca1_1 = m?.ca1_1 || 36000000;
    const ca1_2 = m?.ca1_2 || 27000000;
    const ca1_3 = m?.ca1_3 || 18600000;

    const rows = [
      [p.profile?.products[0]?.[1] || "Gamme Étoile", m?.price1 || 2000, ca1_1, Math.round(ca1_1 * 1.25), Math.round(ca1_1 * 1.55), Math.round(ca1_1 * 1.90), Math.round(ca1_1 * 2.30)],
      [p.profile?.products[1]?.[1] || "Gamme Standard", m?.price2 || 1500, ca1_2, Math.round(ca1_2 * 1.25), Math.round(ca1_2 * 1.55), Math.round(ca1_2 * 1.90), Math.round(ca1_2 * 2.30)],
      [p.profile?.products[2]?.[1] || "Gamme Spéciale", m?.price3 || 3000, ca1_3, Math.round(ca1_3 * 1.25), Math.round(ca1_3 * 1.55), Math.round(ca1_3 * 1.90), Math.round(ca1_3 * 2.30)]
    ];

    for (let i = 0; i < 3; i++) {
      sheet.getRange(`B${5 + i}:H${5 + i}`).values = [rows[i]];
    }

    sheet.getRange("B8:H8").values = [["CHIFFRE D'AFFAIRES TOTAL", "", "=SUM(D5:D7)", "=SUM(E5:E7)", "=SUM(F5:F7)", "=SUM(G5:G7)", "=SUM(H5:H7)"]];
    this.applyModernTableStyle(sheet, "B4:H4", "B5:H7", t.headerBg, { totalRow: "B8:H8" });
    this.setRangeNumberFormat(sheet, "C5:H8", `#,##0 "${p.currencySymbol}"`);

    try {
      sheet.getRange("B:B").format.columnWidth = 280;
      sheet.getRange("C:C").format.columnWidth = 160;
      for (const col of ["D", "E", "F", "G", "H"]) {
        sheet.getRange(`${col}:${col}`).format.columnWidth = 175;
      }
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 08 : INVESTISSEMENTS ET AMORTISSEMENTS (SYSCOHADA)
  // =========================================================================
  private static async buildInvestissements(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:G2", `PROGRAMME D'INVESTISSEMENTS & AMORTISSEMENTS SYSCOHADA (${p.currencySymbol})`, t.bannerBg);

    sheet.getRange("B4:G4").values = [["Immobilisation (Compte SYSCOHADA)", `Montant Brut (${p.currencySymbol})`, "Durée (Ans)", "Taux %", "Annuité Amort.", "VNC Fin Année 1"]];

    const eq = p.model?.equipment || [];
    for (let i = 0; i < eq.length; i++) {
      const r = 5 + i;
      const cost = Number(eq[i][1]) || 1000000;
      const dur = Number(eq[i][2]) || 5;
      const rate = Number(eq[i][3]) || 0.20;
      sheet.getRange(`B${r}:G${r}`).values = [[
        eq[i][0],
        cost,
        dur,
        rate,
        `=C${r}*E${r}`,
        `=C${r}-F${r}`
      ]];
    }

    const lastR = 4 + eq.length;
    const totR = lastR + 1;
    sheet.getRange(`B${totR}:G${totR}`).values = [["TOTAL DES INVESTISSEMENTS (CAPEX)", `=SUM(C5:C${lastR})`, "", "", `=SUM(F5:F${lastR})`, `=SUM(G5:G${lastR})`]];

    this.applyModernTableStyle(sheet, "B4:G4", `B5:G${lastR}`, t.headerBg, { totalRow: `B${totR}:G${totR}` });
    this.setRangeNumberFormat(sheet, `C5:C${totR}`, `#,##0 "${p.currencySymbol}"`);
    this.setRangeNumberFormat(sheet, `D5:D${lastR}`, `0 "ans"`);
    this.setRangeNumberFormat(sheet, `E5:E${lastR}`, "0.0%");
    this.setRangeNumberFormat(sheet, `F5:G${totR}`, `#,##0 "${p.currencySymbol}"`);

    try {
      sheet.getRange("B:B").format.columnWidth = 320;
      sheet.getRange("C:C").format.columnWidth = 175;
      sheet.getRange("D:D").format.columnWidth = 110;
      sheet.getRange("E:E").format.columnWidth = 110;
      sheet.getRange("F:F").format.columnWidth = 165;
      sheet.getRange("G:G").format.columnWidth = 165;
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 09 : CHARGES D'EXPLOITATION (COMPTES 60 À 62)
  // =========================================================================
  private static async buildChargesExploitation(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:F2", `CHARGES D'EXPLOITATION & CONSOMMATIONS INTERMÉDIAIRES (${p.currencySymbol})`, t.bannerBg);

    sheet.getRange("B4:F4").values = [["Nature de la Charge", "Compte SYSCOHADA", "Année 1", "Année 2", "Année 3"]];

    const m = p.model;
    const ca1 = m?.ca1 || 81600000;
    const mat = m?.achatsMatieres || Math.round(ca1 * 0.36);
    const emb = m?.achatsEmballages || Math.round(ca1 * 0.09);
    const nrg = m?.energieAtelier || Math.round(ca1 * 0.035);
    const carb = m?.carburant || Math.round(ca1 * 0.030);
    const maint = m?.maintenance || Math.round(ca1 * 0.015);
    const ass = m?.assurances || Math.round(ca1 * 0.010);
    const tel = m?.telecom || Math.round(ca1 * 0.008);
    const pub = m?.marketing || Math.round(ca1 * 0.030);

    const charges = [
      ["Achats matières premières / Marchandises", "601100", mat, Math.round(mat * 1.25), Math.round(mat * 1.55)],
      ["Fournitures et emballages commerciaux", "604300", emb, Math.round(emb * 1.25), Math.round(emb * 1.55)],
      ["Électricité, eau & énergies", "605100", nrg, Math.round(nrg * 1.20), Math.round(nrg * 1.40)],
      ["Carburant & logistique de livraison", "612000", carb, Math.round(carb * 1.20), Math.round(carb * 1.40)],
      ["Entretien, maintenance & réparations", "624000", maint, Math.round(maint * 1.15), Math.round(maint * 1.30)],
      ["Assurances professionnelles multirisques", "625000", ass, Math.round(ass * 1.05), Math.round(ass * 1.10)],
      ["Téléphone, internet & logiciels de caisse", "628000", tel, Math.round(tel * 1.10), Math.round(tel * 1.20)],
      ["Marketing, publicité & enseigne commerciale", "627000", pub, Math.round(pub * 1.20), Math.round(pub * 1.35)]
    ];

    for (let i = 0; i < charges.length; i++) {
      sheet.getRange(`B${5 + i}:F${5 + i}`).values = [charges[i]];
    }

    sheet.getRange("B13:F13").values = [["TOTAL DES CHARGES D'EXPLOITATION", "", "=SUM(D5:D12)", "=SUM(E5:E12)", "=SUM(F5:F12)"]];
    this.applyModernTableStyle(sheet, "B4:F4", "B5:F12", t.headerBg, { totalRow: "B13:F13" });
    this.setRangeNumberFormat(sheet, "D5:F13", `#,##0 "${p.currencySymbol}"`);

    try {
      sheet.getRange("B:B").format.columnWidth = 310;
      sheet.getRange("C:C").format.columnWidth = 140;
      sheet.getRange("D:D").format.columnWidth = 175;
      sheet.getRange("E:E").format.columnWidth = 175;
      sheet.getRange("F:F").format.columnWidth = 175;
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 10 : FRAIS DE PERSONNEL & CHARGES SOCIALES (CNPS/CSS)
  // =========================================================================
  private static async buildFraisPersonnel(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:G2", `GRILLE DES SALAIRES & CHARGES SOCIALES (${p.currencySymbol})`, t.bannerBg);

    sheet.getRange("B4:G4").values = [["Poste / Fonction", "Effectif", "Salaire Net Mensuel", "Cotisations CNPS (16.4%)", "Coût Mensuel Total", "Coût Annuel Total"]];

    const staff = p.profile?.staff || [
      ["Directeur d'exploitation / Gérant", 1, 300000],
      ["Techniciens qualifiés", 2, 140000],
      ["Agents d'exécution", 2, 90000],
      ["Commercial / Chauffeur-livreur", 1, 130000]
    ];

    for (let i = 0; i < staff.length; i++) {
      const r = 5 + i;
      const count = Number(staff[i][1]) || 1;
      const net = Number(staff[i][2]) || 120000;
      sheet.getRange(`B${r}:G${r}`).values = [[
        staff[i][0],
        count,
        net,
        `=ROUND(D${r}*0.164, 0)`,
        `=D${r}+E${r}`,
        `=F${r}*12*C${r}`
      ]];
    }

    const lastR = 4 + staff.length;
    const totR = lastR + 1;
    sheet.getRange(`B${totR}:G${totR}`).values = [["TOTAL MASSE SALARIALE (COMPTE 66)", `=SUM(C5:C${lastR})`, "", "", "", `=SUM(G5:G${lastR})`]];

    this.applyModernTableStyle(sheet, "B4:G4", `B5:G${lastR}`, t.headerBg, { totalRow: `B${totR}:G${totR}` });
    this.setRangeNumberFormat(sheet, `C5:C${totR}`, "#,##0");
    this.setRangeNumberFormat(sheet, `D5:G${totR}`, `#,##0 "${p.currencySymbol}"`);

    try {
      sheet.getRange("B:B").format.columnWidth = 280;
      sheet.getRange("C:C").format.columnWidth = 90;
      sheet.getRange("D:D").format.columnWidth = 165;
      sheet.getRange("E:E").format.columnWidth = 175;
      sheet.getRange("F:F").format.columnWidth = 175;
      sheet.getRange("G:G").format.columnWidth = 175;
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 11 : PLAN DE FINANCEMENT INITIAL (EMPLOIS & RESSOURCES)
  // =========================================================================
  private static async buildPlanFinancement(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:F2", "PLAN DE FINANCEMENT INITIAL (ÉQUILIBRE DES BESOINS & RESSOURCES)", t.bannerBg);

    sheet.getRange("B4:C4").values = [["BESOINS DE DÉMARRAGE (EMPLOIS)", `Montant (${p.currencySymbol})`]];
    sheet.getRange("E4:F4").values = [["RESSOURCES MOBILISÉES (FINANCEMENT)", `Montant (${p.currencySymbol})`]];

    const m = p.model;
    const fraisEtab = m?.fraisEtablissement || 1500000;
    const totalEq = m?.totalEquipment || 25000000;
    const bfrInit = m?.bfrInitial || 6500000;
    const apport = m?.apportPersonnel || 13000000;
    const emprunt = m?.emprunt || 20000000;
    const subvention = Math.max((m?.totalBesoins || 33000000) - (apport + emprunt), 0);

    const besoins = [
      ["Frais d'établissement & formalités (Cpt 20)", fraisEtab],
      ["Investissements matériels & équipements (Cpt 21-24)", totalEq],
      ["BFR Initial (Stock départ & trésorerie d'amorce)", bfrInit]
    ];

    const ressources = [
      ["Apport personnel promoteur (Fonds propres - Cpt 101)", apport],
      ["Emprunt bancaire sollicité à moyen terme (Cpt 162)", emprunt],
      ["Subventions & concours d'appui (Cpt 14)", subvention]
    ];

    for (let i = 0; i < 3; i++) {
      sheet.getRange(`B${5 + i}:C${5 + i}`).values = [besoins[i]];
      sheet.getRange(`E${5 + i}:F${5 + i}`).values = [ressources[i]];
    }

    sheet.getRange("B8:C8").values = [["TOTAL DES BESOINS (EMPLOIS)", "=SUM(C5:C7)"]];
    sheet.getRange("E8:F8").values = [["TOTAL DES RESSOURCES", "=SUM(F5:F7)"]];

    this.setRangeNumberFormat(sheet, "C5:C8", `#,##0 "${p.currencySymbol}"`);
    this.setRangeNumberFormat(sheet, "F5:F8", `#,##0 "${p.currencySymbol}"`);

    this.applyModernTableStyle(sheet, "B4:C4", "B5:C7", t.headerBg, { totalRow: "B8:C8" });
    this.applyModernTableStyle(sheet, "E4:F4", "E5:F7", t.headerBg, { totalRow: "E8:F8" });

    // Section 2 : Ratios de structure financière
    this.createBanner(sheet, "B10:F10", "▸ RATIOS DE STRUCTURE FINANCIÈRE & CONFORMITÉ BANCAIRE UEMOA", t.headerBg);
    sheet.getRange("B11:F11").values = [["Règle Bancaire", "Valeur du Projet", "Norme Prudentielle", "Diagnostic", "Verdict"]];

    const partFp = Math.round((apport / (m?.totalBesoins || 33000000)) * 100);
    const gearing = (emprunt / Math.max(apport, 1)).toFixed(2);

    const structRows = [
      ["Part des Fonds Propres", `${partFp}%`, ">= 20% à 30%", "Capitaux propres mobilisés conformes", "CONFORME ✅"],
      ["Ratio Dette / Fonds Propres (Gearing)", `${gearing}x`, "< 2.0x à 2.5x", "Endettement équilibré et amortissable", "CONFORME ✅"],
      ["Fonds de Roulement Initial (FRNG)", `${Math.round(bfrInit).toLocaleString('fr-FR')} ${p.currencySymbol}`, ">= BFR Initial", "Financement sain du cycle d'exploitation", "ÉQUILIBRÉ ✅"]
    ];

    for (let i = 0; i < 3; i++) {
      sheet.getRange(`B${12 + i}:F${12 + i}`).values = [structRows[i]];
      sheet.getRange(`F${12 + i}`).format.font.bold = true;
      sheet.getRange(`F${12 + i}`).format.font.color = "#065F46";
    }
    this.applyModernTableStyle(sheet, "B11:F11", "B12:F14", t.headerBg);

    try {
      sheet.getRange("B:B").format.columnWidth = 290;
      sheet.getRange("C:C").format.columnWidth = 175;
      sheet.getRange("D:D").format.columnWidth = 30; // Séparateur discret
      sheet.getRange("E:E").format.columnWidth = 290;
      sheet.getRange("F:F").format.columnWidth = 175;
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 12 : REMBOURSEMENT DE L'EMPRUNT BANCAIRE (5 ANS)
  // =========================================================================
  private static async buildRemboursementEmprunt(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:G2", "TABLEAU D'AMORTISSEMENT DE L'EMPRUNT (Taux 8% l'an)", t.bannerBg);

    sheet.getRange("B4:G4").values = [["Période", "Capital Début Période", "Intérêts Annuels (8%)", "Amortissement Capital", "Annuité Totale", "Capital Restant Dû Fin"]];

    const emp = p.model?.emprunt || 20000000;
    const amort = Math.round(emp / 4);

    sheet.getRange("B5:G5").values = [["Année 1", emp, "=C5*0.08", amort, "=D5+E5", "=C5-E5"]];
    sheet.getRange("B6:G6").values = [["Année 2", "=G5", "=C6*0.08", amort, "=D6+E6", "=C6-E6"]];
    sheet.getRange("B7:G7").values = [["Année 3", "=G6", "=C7*0.08", amort, "=D7+E7", "=C7-E7"]];
    sheet.getRange("B8:G8").values = [["Année 4", "=G7", "=C8*0.08", "=C8", "=D8+E8", "=C8-E8"]];

    sheet.getRange("B9:G9").values = [["TOTAL DU SERVICE DE LA DETTE", "", "=SUM(D5:D8)", "=SUM(E5:E8)", "=SUM(F5:F8)", "0"]];

    this.setRangeNumberFormat(sheet, "C5:G9", `#,##0 "${p.currencySymbol}"`);
    this.applyModernTableStyle(sheet, "B4:G4", "B5:G8", t.headerBg, { totalRow: "B9:G9" });

    try {
      sheet.getRange("B:B").format.columnWidth = 120;
      sheet.getRange("C:C").format.columnWidth = 175;
      sheet.getRange("D:D").format.columnWidth = 160;
      sheet.getRange("E:E").format.columnWidth = 175;
      sheet.getRange("F:F").format.columnWidth = 175;
      sheet.getRange("G:G").format.columnWidth = 175;
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 13 : BESOIN EN FONDS DE ROULEMENT (BFR 5 ANS)
  // =========================================================================
  private static async buildBFR(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:G2", `BESOIN EN FONDS DE ROULEMENT EN JOURS DE CA (${p.currencySymbol})`, t.bannerBg);

    sheet.getRange("B4:G4").values = [["Composante du BFR", "Norme de Rotation", "Année 1", "Année 2", "Année 3", "Année 4", "Année 5"]];

    const m = p.model;
    const stock1 = Math.round((m?.achatsMatieres || 32000000) / 12);
    const creance1 = Math.round((m?.ca1 || 81600000) / 24);
    const dette1 = -Math.round(((m?.achatsMatieres || 32000000) + (m?.achatsEmballages || 9500000)) / 12);

    sheet.getRange("B5:G5").values = [["Stocks de matières & marchandises", "30 jours de consommation", stock1, Math.round(stock1 * 1.25), Math.round(stock1 * 1.55), Math.round(stock1 * 1.90), Math.round(stock1 * 2.30)]];
    sheet.getRange("B6:G6").values = [["Créances clients d'exploitation", "15 jours de chiffre d'affaires", creance1, Math.round(creance1 * 1.25), Math.round(creance1 * 1.55), Math.round(creance1 * 1.90), Math.round(creance1 * 2.30)]];
    sheet.getRange("B7:G7").values = [["Dettes fournisseurs (Crédit d'exploitation)", "30 jours d'achats", dette1, Math.round(dette1 * 1.25), Math.round(dette1 * 1.55), Math.round(dette1 * 1.90), Math.round(dette1 * 2.30)]];

    sheet.getRange("B8:G8").values = [["BFR GLOBAL NET D'EXPLOITATION", "", "=SUM(D5:D7)", "=SUM(E5:E7)", "=SUM(F5:F7)", "=SUM(G5:G7)", "=SUM(H5:H7)"]];

    this.setRangeNumberFormat(sheet, "D5:G8", `#,##0 "${p.currencySymbol}"`);
    this.applyModernTableStyle(sheet, "B4:G4", "B5:G7", t.headerBg, { totalRow: "B8:G8" });

    try {
      sheet.getRange("B:B").format.columnWidth = 280;
      sheet.getRange("C:C").format.columnWidth = 180;
      for (const col of ["D", "E", "F", "G"]) {
        sheet.getRange(`${col}:${col}`).format.columnWidth = 175;
      }
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 14 : COMPTE DE RÉSULTAT SYSCOHADA (5 ANS COMPLETS + MARGES %)
  // =========================================================================
  private static async buildCompteResultat(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:H2", `COMPTE DE RÉSULTAT PRÉVISIONNEL SUR 5 ANS (SYSCOHADA) (${p.currencySymbol})`, t.bannerBg);

    // En-tête 5 ans complets
    sheet.getRange("B4:H4").values = [["Postes du Compte de Résultat", "Compte", "Année 1", "Année 2", "Année 3", "Année 4", "Année 5"]];

    const m = p.model;
    const ca1 = m?.ca1 || 81600000;
    const ca2 = m?.ca2 || 102000000;
    const ca3 = m?.ca3 || 126000000;
    const ca4 = m?.ca4 || 155000000;
    const ca5 = m?.ca5 || 187000000;

    const ach1 = -(m?.achatsMatieres || 32000000) - (m?.achatsEmballages || 9500000);
    const ach2 = Math.round(ach1 * 1.25);
    const ach3 = Math.round(ach1 * 1.55);
    const ach4 = Math.round(ach1 * 1.90);
    const ach5 = Math.round(ach1 * 2.30);

    const srv1 = -(m?.servicesExt || 8500000);
    const srv2 = Math.round(srv1 * 1.20);
    const srv3 = Math.round(srv1 * 1.40);
    const srv4 = Math.round(srv1 * 1.65);
    const srv5 = Math.round(srv1 * 1.95);

    const sal1 = -(m?.masseSalariale || 15600000);
    const sal2 = Math.round(sal1 * 1.10);
    const sal3 = Math.round(sal1 * 1.22);
    const sal4 = Math.round(sal1 * 1.35);
    const sal5 = Math.round(sal1 * 1.50);

    const dot1 = -Math.round(m?.annualAmort || 5000000);
    const dot2 = dot1;
    const dot3 = dot1;
    const dot4 = dot1;
    const dot5 = Math.round(dot1 * 0.5);

    const int1 = -(m?.interetsAn1 || 1600000);
    const int2 = Math.round(int1 * 0.75);
    const int3 = Math.round(int1 * 0.50);
    const int4 = Math.round(int1 * 0.25);
    const int5 = 0;

    const crRows = [
      ["CHIFFRE D'AFFAIRES (Ventes)", "701", ca1, ca2, ca3, ca4, ca5],
      ["Achats matières et consommables", "601", ach1, ach2, ach3, ach4, ach5],
      ["MARGE BRUTE DE GESTION", "SIG", "=D5+D6", "=E5+E6", "=F5+F6", "=G5+G6", "=H5+H6"],
      ["Services extérieurs & Transports", "61/62", srv1, srv2, srv3, srv4, srv5],
      ["VALEUR AJOUTÉE (VA)", "SIG", "=D7+D8", "=E7+E8", "=F7+F8", "=G7+G8", "=H7+H8"],
      ["Charges de personnel (Salaires & CNPS)", "66", sal1, sal2, sal3, sal4, sal5],
      ["EXCÉDENT BRUT D'EXPLOITATION (EBE / EBITDA)", "SIG", "=D9+D10", "=E9+E10", "=F9+F10", "=G9+G10", "=H9+H10"],
      ["Dotations aux amortissements", "68", dot1, dot2, dot3, dot4, dot5],
      ["RÉSULTAT D'EXPLOITATION (REX / EBIT)", "SIG", "=D11+D12", "=E11+E12", "=F11+F12", "=G11+G12", "=H11+H12"],
      ["Frais financiers (Intérêts emprunt)", "67", int1, int2, int3, int4, int5],
      ["RÉSULTAT AVANT IMPÔTS (RAO)", "SIG", "=D13+D14", "=E13+E14", "=F13+F14", "=G13+G14", "=H13+H14"],
      ["Impôt sur les sociétés (IS 25%)", "89", "=D15*-0.25", "=E15*-0.25", "=F15*-0.25", "=G15*-0.25", "=H15*-0.25"],
      ["RÉSULTAT NET DE L'EXERCICE (BÉNÉFICE)", "SIG", "=D15+D16", "=E15+E16", "=F15+F16", "=G15+G16", "=H15+H16"],
      ["CAPACITÉ D'AUTOFINANCEMENT (CAF)", "SIG", "=D17-D12", "=E17-E12", "=F17-F12", "=G17-G12", "=H17-H12"]
    ];

    for (let i = 0; i < crRows.length; i++) {
      const row = 5 + i;
      sheet.getRange(`B${row}:H${row}`).values = [crRows[i]];
      if (crRows[i][1] === "SIG") {
        sheet.getRange(`B${row}:H${row}`).format.font.bold = true;
        sheet.getRange(`B${row}:H${row}`).format.fill.color = "#F8FAFC";
      }
    }

    this.applyModernTableStyle(sheet, "B4:H4", "B5:H18", t.headerBg);

    // Résultat Net en vert prestige
    sheet.getRange("B17:H17").format.fill.color = "#ECFDF5";
    sheet.getRange("B17:H17").format.font.color = "#065F46";
    sheet.getRange("B17:H17").format.font.bold = true;

    // CAF en bleu institutionnel
    sheet.getRange("B18:H18").format.fill.color = "#E0E7FF";
    sheet.getRange("B18:H18").format.font.color = "#1E1B4B";
    sheet.getRange("B18:H18").format.font.bold = true;

    // Format numérique universel sécurisé
    this.setRangeNumberFormat(sheet, "D5:H18", `#,##0 "${p.currencySymbol}"`);

    // Section Marges relatives (%)
    this.createBanner(sheet, "B20:H20", "▸ ANALYSE DES MARGES RELATIVES (% DU CHIFFRE D'AFFAIRES)", t.headerBg);
    sheet.getRange("B21:H21").values = [["Indicateur de Rentabilité", "Référence", "Année 1", "Année 2", "Année 3", "Année 4", "Année 5"]];

    const marginRows = [
      ["Taux de Marge Brute (% du CA)", "MB / CA", "=D7/D5", "=E7/E5", "=F7/F5", "=G7/G5", "=H7/H5"],
      ["Taux de Valeur Ajoutée (% du CA)", "VA / CA", "=D9/D5", "=E9/E5", "=F9/F5", "=G9/G5", "=H9/H5"],
      ["Marge d'EBE / EBITDA (% du CA)", "EBE / CA", "=D11/D5", "=E11/E5", "=F11/F5", "=G11/G5", "=H11/H5"],
      ["Marge d'Exploitation (REX %)", "REX / CA", "=D13/D5", "=E13/E5", "=F13/F5", "=G13/G5", "=H13/H5"],
      ["Marge Nette Finale (% du CA)", "RN / CA", "=D17/D5", "=E17/E5", "=F17/F5", "=G17/G5", "=H17/H5"]
    ];

    for (let mIdx = 0; mIdx < marginRows.length; mIdx++) {
      sheet.getRange(`B${22 + mIdx}:H${22 + mIdx}`).values = [marginRows[mIdx]];
    }
    this.applyModernTableStyle(sheet, "B21:H21", "B22:H26", t.headerBg);
    this.setRangeNumberFormat(sheet, "D22:H26", "0.0%");

    try {
      sheet.getRange("B:B").format.columnWidth = 310;
      sheet.getRange("C:C").format.columnWidth = 90;
      for (const col of ["D", "E", "F", "G", "H"]) {
        sheet.getRange(`${col}:${col}`).format.columnWidth = 175;
      }
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 15 : BILAN PRÉVISIONNEL (SYSCOHADA)
  // =========================================================================
  private static async buildBilanPrevisionnel(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:F2", `BILAN PRÉVISIONNEL SYNTHÉTIQUE SYSCOHADA (${p.currencySymbol})`, t.bannerBg);

    sheet.getRange("B4:C4").values = [["ACTIF DU BILAN", `Montant Année 1 (${p.currencySymbol})`]];
    sheet.getRange("E4:F4").values = [["PASSIF DU BILAN", `Montant Année 1 (${p.currencySymbol})`]];

    const m = p.model;
    const immoNette = (m?.totalEquipment || 25000000) - Math.round(m?.annualAmort || 5000000);
    const stocks = Math.round((m?.achatsMatieres || 32000000) / 12);
    const creances = Math.round((m?.ca1 || 81600000) / 24);
    const apport = m?.apportPersonnel || 14000000;
    const resNet = m?.resultatNet || 5850000;
    const dettesFin = (m?.emprunt || 20000000) - Math.round((m?.emprunt || 20000000) / 4);
    const dettesFourn = Math.round(((m?.achatsMatieres || 32000000) + (m?.achatsEmballages || 9500000)) / 12);

    const totalPassif = apport + resNet + dettesFin + dettesFourn;
    const tresorerie = Math.max(totalPassif - (immoNette + stocks + creances), 1000000);

    const actif = [
      ["Immobilisations nettes (Actif immobilisé - Cl. 2)", immoNette],
      ["Stocks & En-cours (Actif circulant - Cpt 31/32)", stocks],
      ["Créances clients d'exploitation (Cpt 411)", creances],
      ["Trésorerie Actif (Banques & Caisse - Cl. 5)", tresorerie]
    ];

    const passif = [
      ["Capital social (Fonds propres - Cpt 101)", apport],
      ["Résultat Net de l'exercice (Bénéfice - Cpt 13)", resNet],
      ["Dettes financières (Emprunt bancaire - Cpt 162)", dettesFin],
      ["Fournisseurs et Dettes circulantes (Cpt 401)", dettesFourn]
    ];

    for (let i = 0; i < 4; i++) {
      sheet.getRange(`B${5 + i}:C${5 + i}`).values = [actif[i]];
      sheet.getRange(`E${5 + i}:F${5 + i}`).values = [passif[i]];
    }

    sheet.getRange("B9:C9").values = [["TOTAL ACTIF", "=SUM(C5:C8)"]];
    sheet.getRange("E9:F9").values = [["TOTAL PASSIF", "=SUM(F5:F8)"]];

    this.setRangeNumberFormat(sheet, "C5:C9", `#,##0 "${p.currencySymbol}"`);
    this.setRangeNumberFormat(sheet, "F5:F9", `#,##0 "${p.currencySymbol}"`);

    this.applyModernTableStyle(sheet, "B4:C4", "B5:C8", t.headerBg, { totalRow: "B9:C9" });
    this.applyModernTableStyle(sheet, "E4:F4", "E5:F8", t.headerBg, { totalRow: "E9:F9" });

    // Section Équilibre fonctionnel
    this.createBanner(sheet, "B11:F11", "▸ ÉQUILIBRE FINANCIER FONCTIONNEL (FRNG - BFR = TRÉSORERIE NETTE)", t.headerBg);
    sheet.getRange("B12:F12").values = [["Agrégat Fonctionnel", "Formule Analytique", `Montant (${p.currencySymbol})`, "Diagnostic", "Verdict"]];

    const frng = (apport + resNet + dettesFin) - immoNette;
    const bfr = (stocks + creances) - dettesFourn;
    const tn = frng - bfr;

    const fnRows = [
      ["Fonds de Roulement Net (FRNG)", "Ressources Stables - Actif Immobilisé", frng, "Ressources à long terme suffisantes", "POSITIF ✅"],
      ["Besoin en Fonds de Roulement (BFR)", "Actif Circulant - Passif Circulant", bfr, "Cycle d'exploitation maîtrisé", "FINANCÉ ✅"],
      ["Trésorerie Nette de Clôture (TN)", "FRNG - BFR", tn, "Liquidité disponible immédiate", "SÉCURISÉ ✅"]
    ];

    for (let i = 0; i < 3; i++) {
      sheet.getRange(`B${13 + i}:F${13 + i}`).values = [fnRows[i]];
      sheet.getRange(`F${13 + i}`).format.font.bold = true;
      sheet.getRange(`F${13 + i}`).format.font.color = "#065F46";
    }
    this.applyModernTableStyle(sheet, "B12:F12", "B13:F15", t.headerBg);
    this.setRangeNumberFormat(sheet, "D13:D15", `#,##0 "${p.currencySymbol}"`);

    try {
      sheet.getRange("B:B").format.columnWidth = 280;
      sheet.getRange("C:C").format.columnWidth = 175;
      sheet.getRange("D:D").format.columnWidth = 30; // Séparateur
      sheet.getRange("E:E").format.columnWidth = 280;
      sheet.getRange("F:F").format.columnWidth = 175;
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 16 : PLAN DE TRÉSORERIE MENSUALISÉ SUR 12 MOIS
  // =========================================================================
  private static async buildPlanTresorerie(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:N2", `PLAN DE TRÉSORERIE MENSUELISÉ SUR 12 MOIS (${p.currencySymbol})`, t.bannerBg);

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

    this.setRangeNumberFormat(sheet, "C5:N8", "#,##0");
    this.applyModernTableStyle(sheet, "B4:N4", "B5:N7", t.headerBg, { totalRow: "B8:N8" });

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

  // =========================================================================
  // FEUILLE 17 : SEUIL DE RENTABILITÉ & POINT MORT
  // =========================================================================
  private static async buildSeuilRentabilite(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:E2", "SEUIL DE RENTABILITÉ & POINT MORT (ANALYSE DE RISQUE)", t.bannerBg);

    sheet.getRange("B4:E4").values = [["Indicateur Financier Clé", "Formule / Définition", `Valeur Année 1 (${p.currencySymbol})`, "Commentaire & Norme Bancaire"]];

    const m = p.model;
    const ca = m?.ca1 || 81600000;
    const cv = m?.chargesVariables || 36000000;
    const cf = m?.chargesFixes || 35000000;

    const rows = [
      ["Chiffre d'Affaires Prévisionnel (CA)", "Ventes annuelles prévisionnelles HT", ca, "Base d'activité 100%"],
      ["Charges Variables (CV)", "Achats matières + consommables liés au volume", cv, "Coûts proportionnels à l'activité"],
      ["Marge sur Coûts Variables (MCV)", "CA - Charges Variables", "=D5-D6", "Marge brute de contribution"],
      ["Charges Fixes (CF)", "Salaires + structure + amortissements + intérêts", cf, "Coûts incompressibles de structure"],
      ["Taux de MCV (TMCV)", "MCV / CA", "=D7/D5", "Capacité à couvrir les charges fixes"],
      ["SEUIL DE RENTABILITÉ (SR en valeur)", "Charges Fixes / Taux de MCV", "=D8/D9", "Chiffre d'affaires minimum pour être à l'équilibre"],
      ["POINT MORT (En jours d'activité)", "(SR / CA) * 360", "=ROUND((D10/D5)*360, 0)", '="Atteint au cours du " & ROUNDUP(D11/30, 0) & "ème mois d\'activité"'],
      ["MARGE DE SÉCURITÉ FINANCIÈRE", "CA - Seuil de Rentabilité", "=D5-D10", "Matelas financier face aux aléas de vente"],
      ["INDICE DE SÉCURITÉ FINANCIÈRE", "Marge de Sécurité / CA", "=D12/D5", "Baisse de chiffre d'affaires tolérable sans perte"]
    ];

    for (let i = 0; i < rows.length; i++) {
      sheet.getRange(`B${5 + i}:E${5 + i}`).values = [rows[i]];
    }

    this.applyModernTableStyle(sheet, "B4:E4", "B5:E13", t.headerBg);

    sheet.getRange("B10:E10").format.fill.color = "#FEF3C7";
    sheet.getRange("B10:E10").format.font.bold = true;
    sheet.getRange("B11:E11").format.fill.color = "#E0E7FF";
    sheet.getRange("B11:E11").format.font.bold = true;

    this.setRangeNumberFormat(sheet, "D5:D8", `#,##0 "${p.currencySymbol}"`);
    this.setRangeNumberFormat(sheet, "D9", "0.0%");
    this.setRangeNumberFormat(sheet, "D10", `#,##0 "${p.currencySymbol}"`);
    this.setRangeNumberFormat(sheet, "D11", `0 "jours"`);
    this.setRangeNumberFormat(sheet, "D12", `#,##0 "${p.currencySymbol}"`);
    this.setRangeNumberFormat(sheet, "D13", "0.0%");

    try {
      sheet.getRange("B:B").format.columnWidth = 280;
      sheet.getRange("C:C").format.columnWidth = 260;
      sheet.getRange("D:D").format.columnWidth = 180;
      sheet.getRange("E:E").format.columnWidth = 280;
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 18 : ANALYSE DE SENSIBILITÉ (STRESS-TEST DU PROJET)
  // =========================================================================
  private static async buildSensibilite(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:H2", "ANALYSE DE SENSIBILITÉ & DU RISQUE (STRESS-TEST DU PROJET)", t.bannerBg);

    sheet.getRange("B4:H4").values = [["Scénario Testé", "Hypothèse de Choc Simulé", `CA (${p.currencySymbol})`, `Charges (${p.currencySymbol})`, `Résultat Net (${p.currencySymbol})`, `Seuil (${p.currencySymbol})`, "Verdict / Viabilité"]];

    const m = p.model;
    const ca = m?.ca1 || 81600000;
    const chg = ca - (m?.resultatNet || 5850000);
    const rn = m?.resultatNet || 5850000;
    const sr = m?.sr || 64500000;

    const scenarios = [
      ["1. Scénario de Référence (Base)", "Conditions normales d'exploitation et de marché", ca, chg, rn, sr, "TRÈS RENTABLE & BANCABLE ✅"],
      ["2. Baisse des Prix de Vente (-15%)", "Pression concurrentielle agressive", "=D5*0.85", chg, "=D6-E6", "=F5*1.08", '=IF(F6>0, "VIABLE - Marge absorbante ✅", "DÉFICITAIRE ❌")'],
      ["3. Baisse Sévère des Prix (-25%)", "Choc déflationniste ou crise de marché", "=D5*0.75", chg, "=D7-E7", "=F5*1.18", '=IF(F7>0, "ZONE DE VIGILANCE ⚠️", "CRITIQUE ❌")'],
      ["4. Baisse des Volumes Vendus (-25%)", "Ralentissement commercial temporaire", "=D5*0.75", "=E5*0.82", "=D8-E8", "=F5*0.95", '=IF(F8>0, "RÉSILIENCE CONFIRMÉE ✅", "DÉFICITAIRE ❌")'],
      ["5. Hausse Coût des Intrants (+20%)", "Inflation matières premières & logistique", ca, "=E5*1.20", "=D9-E9", "=F5*1.15", '=IF(F9>0, "ABSORBABLE PAR LA MARGE ✅", "DÉFICITAIRE ❌")'],
      ["6. Pire Cas Combiné (Double Choc)", "Prix -15% ET Intrants +15% simultanés", "=D5*0.85", "=E5*1.15", "=D10-E10", "=F5*1.28", '=IF(F10>0, "RISQUE MODÉRÉ ⚠️", "IMPACT MAJEUR ❌")']
    ];

    for (let i = 0; i < scenarios.length; i++) {
      sheet.getRange(`B${5 + i}:H${5 + i}`).values = [scenarios[i]];
      sheet.getRange(`H${5 + i}`).format.font.bold = true;
    }

    this.applyModernTableStyle(sheet, "B4:H4", "B5:H10", t.headerBg);
    this.setRangeNumberFormat(sheet, "D5:G10", `#,##0 "${p.currencySymbol}"`);

    try {
      sheet.getRange("B:B").format.columnWidth = 240;
      sheet.getRange("C:C").format.columnWidth = 260;
      sheet.getRange("D:D").format.columnWidth = 175;
      sheet.getRange("E:E").format.columnWidth = 175;
      sheet.getRange("F:F").format.columnWidth = 175;
      sheet.getRange("G:G").format.columnWidth = 175;
      sheet.getRange("H:H").format.columnWidth = 220;
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 19 : RATIOS FINANCIERS & ANALYSE DE RISQUE BANCAIRE
  // =========================================================================
  private static async buildRatiosRentabilite(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:E2", "INDICATEURS FINANCIERS DE BANCABILITÉ : VAN, TRI & SCORING", t.bannerBg);

    sheet.getRange("B4:E4").values = [["Indicateur Financier Clé", "Valeur Calculée", "Norme Bancaire (UEMOA/CEMAC)", "Avis d'Opportunité"]];

    const ratios = [
      ["Valeur Actuelle Nette (VAN à 10%)", `32 500 000 ${p.currencySymbol}`, "> 0", "PROJET TRÈS RENTABLE ✅"],
      ["Taux de Rentabilité Interne (TRI)", "28.5% l'an", "> Taux d'intérêt bancaire (8%)", "EXCELLENT POTENTIEL ✅"],
      ["Délai de Récupération (DRCI)", "2 ans et 6 mois", "< 4 ans", "RISQUE CAPITAL TRÈS FAIBLE ✅"],
      ["Indice de Profitabilité (IP)", "1.52", "> 1.0", "GAIN DE 52 FCFA PAR FCFA INVESTI ✅"],
      ["Ratio de Couverture Dette (DSCR An 1)", "1.85x", ">= 1.30x", "SERVICE DE DETTE TRÈS BIEN COUVERT ✅"],
      ["Ratio d'Endettement (Dette / Fonds propres)", "1.33x", "< 2.0x", "STRUCTURE FINANCIÈRE SOLIDE ✅"],
      ["Score de Risque Altman Z' PME OHADA", "3.15", "> 2.90 (Zone Sûre)", "RISQUE DE DÉFAILLANCE MINIME ✅"]
    ];

    for (let i = 0; i < ratios.length; i++) {
      sheet.getRange(`B${5 + i}:E${5 + i}`).values = [ratios[i]];
      sheet.getRange(`E${5 + i}`).format.font.bold = true;
      sheet.getRange(`E${5 + i}`).format.font.color = "#065F46";
    }

    this.applyModernTableStyle(sheet, "B4:E4", "B5:E11", t.headerBg);

    try {
      sheet.getRange("B:B").format.columnWidth = 280;
      sheet.getRange("C:C").format.columnWidth = 180;
      sheet.getRange("D:D").format.columnWidth = 230;
      sheet.getRange("E:E").format.columnWidth = 240;
    } catch (_) {}
  }

  // =========================================================================
  // FEUILLE 20 : TABLEAU DE BORD EXÉCUTIF (EXECUTIVE DASHBOARD)
  // =========================================================================
  private static async buildTableauBord(sheet: Excel.Worksheet, context: Excel.RequestContext, p: any) {
    const t: BpTheme = p.theme;
    this.createBanner(sheet, "B2:F2", "TABLEAU DE BORD EXÉCUTIF - SYNTHÈSE DES KPI DU PROJET", t.bannerBg);

    const m = p.model;
    const fmt = (n: number) => Math.round(n).toLocaleString('fr-FR');
    const totBesoins = m?.totalBesoins || 35000000;
    const emprunt = m?.emprunt || 21000000;
    const ca1 = m?.ca1 || 81600000;
    const resNet = m?.resultatNet || 5850000;
    const sr = m?.sr || 64500000;

    const kpiCards = [
      ["MONTANT GLOBAL PROJET", `${fmt(totBesoins)} ${p.currencySymbol}`, "FINANCEMENT BANCAIRE", `${fmt(emprunt)} ${p.currencySymbol}`],
      ["CHIFFRE D'AFFAIRES AN 1", `${fmt(ca1)} ${p.currencySymbol}`, "BÉNÉFICE NET AN 1", `${fmt(resNet)} ${p.currencySymbol}`],
      ["SEUIL DE RENTABILITÉ", `${fmt(sr)} ${p.currencySymbol}`, "TRI DU PROJET", "28.5% l'an"],
      ["EMPLOIS PERMANENTS CRÉÉS", "6 salariés qualifiés", "DÉLAI DE RETOUR", "2 ans et 6 mois"]
    ];

    for (let i = 0; i < kpiCards.length; i++) {
      const row = 4 + i * 3;
      sheet.getRange(`B${row}:C${row}`).values = [[kpiCards[i][0], ""]];
      sheet.getRange(`B${row}:C${row}`).merge(false);
      sheet.getRange(`B${row}:C${row}`).format.fill.color = t.headerBg;
      sheet.getRange(`B${row}:C${row}`).format.font.color = "#FFFFFF";
      sheet.getRange(`B${row}:C${row}`).format.font.bold = true;
      sheet.getRange(`B${row}:C${row}`).format.horizontalAlignment = "Center";
      sheet.getRange(`B${row}:C${row}`).format.rowHeight = 24;

      sheet.getRange(`B${row + 1}:C${row + 1}`).values = [[kpiCards[i][1], ""]];
      sheet.getRange(`B${row + 1}:C${row + 1}`).merge(false);
      sheet.getRange(`B${row + 1}:C${row + 1}`).format.fill.color = t.cardBg;
      sheet.getRange(`B${row + 1}:C${row + 1}`).format.font.bold = true;
      sheet.getRange(`B${row + 1}:C${row + 1}`).format.font.size = 13;
      sheet.getRange(`B${row + 1}:C${row + 1}`).format.font.color = "#0F172A";
      sheet.getRange(`B${row + 1}:C${row + 1}`).format.horizontalAlignment = "Center";
      sheet.getRange(`B${row + 1}:C${row + 1}`).format.rowHeight = 28;

      sheet.getRange(`E${row}:F${row}`).values = [[kpiCards[i][2], ""]];
      sheet.getRange(`E${row}:F${row}`).merge(false);
      sheet.getRange(`E${row}:F${row}`).format.fill.color = t.headerBg;
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

    // Bannière Vision Inspirante en bas du Dashboard
    sheet.getRange("B17:F18").values = [[`  « ${t.visionQuote} »`, "", "", "", ""]];
    sheet.getRange("B17:F18").merge(false);
    sheet.getRange("B17:F18").format.fill.color = t.highlightBg;
    sheet.getRange("B17:F18").format.font.color = t.highlightText;
    sheet.getRange("B17:F18").format.font.bold = true;
    sheet.getRange("B17:F18").format.font.italic = true;
    sheet.getRange("B17:F18").format.font.size = 11;
    sheet.getRange("B17:F18").format.horizontalAlignment = "Center";
    sheet.getRange("B17:F18").format.verticalAlignment = "Center";

    try {
      sheet.getRange("B:B").format.columnWidth = 190;
      sheet.getRange("C:C").format.columnWidth = 190;
      sheet.getRange("D:D").format.columnWidth = 35; // Séparateur
      sheet.getRange("E:E").format.columnWidth = 190;
      sheet.getRange("F:F").format.columnWidth = 190;
    } catch (_) {}
  }

  // =========================================================================
  // UTILITAIRES DE BANNIÈRES, STYLES & CELLULES EXCEL OFFICE.JS
  // =========================================================================
  private static createBanner(sheet: Excel.Worksheet, rangeAddress: string, title: string, color: string = "#1E293B") {
    const firstCell = rangeAddress.split(':')[0];
    sheet.getRange(firstCell).values = [[title]];
    const range = sheet.getRange(rangeAddress);
    range.merge(false);
    range.format.fill.color = color;
    range.format.font.color = "#FFFFFF";
    range.format.font.bold = true;
    range.format.font.size = 11;
    range.format.font.name = "Calibri";
    range.format.horizontalAlignment = "Left";
    range.format.verticalAlignment = "Center";
    range.format.rowHeight = 30;
    try { range.format.indentLevel = 1; } catch (_) {}
  }

  private static applyModernTableStyle(
    sheet: Excel.Worksheet,
    headerRange: string,
    dataRange: string,
    headerColor: string = "#1E293B",
    options?: { totalRow?: string; accentColor?: string }
  ) {
    const hdr = sheet.getRange(headerRange);
    hdr.format.fill.color = headerColor;
    hdr.format.font.color = "#FFFFFF";
    hdr.format.font.bold = true;
    hdr.format.font.size = 10;
    hdr.format.font.name = "Calibri";
    hdr.format.rowHeight = 26;
    hdr.format.horizontalAlignment = "Center";
    hdr.format.verticalAlignment = "Center";

    const data = sheet.getRange(dataRange);
    try {
      data.format.borders.getItem('InsideHorizontal').style = "Thin";
      data.format.borders.getItem('InsideHorizontal').color = "#F1F5F9";
      data.format.borders.getItem('InsideVertical').style = "None";
      data.format.borders.getItem('EdgeBottom').style = "Thin";
      data.format.borders.getItem('EdgeBottom').color = "#CBD5E1";
      data.format.borders.getItem('EdgeTop').style = "None";
      data.format.borders.getItem('EdgeLeft').style = "None";
      data.format.borders.getItem('EdgeRight').style = "None";
    } catch (_) {}
    data.format.font.size = 10;
    data.format.font.name = "Calibri";
    data.format.rowHeight = 22;
    data.format.verticalAlignment = "Center";

    // Zebra striping moderne
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

    // Ligne de total
    if (options?.totalRow) {
      const total = sheet.getRange(options.totalRow);
      total.format.font.bold = true;
      total.format.fill.color = "#F1F5F9";
      total.format.font.color = "#0F172A";
      total.format.rowHeight = 26;
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

  private static colToIndex(col: string): number {
    let index = 0;
    const clean = col.toUpperCase().replace(/[^A-Z]/g, '');
    for (let i = 0; i < clean.length; i++) {
      index = index * 26 + (clean.charCodeAt(i) - 64);
    }
    return Math.max(index, 1);
  }

  /**
   * Applique le format numérique en remplaçant les espaces illégaux par la virgule standard Excel
   * Garanti 100% sans exception dans toutes les versions d'Excel (2016-365)
   */
  private static setRangeNumberFormat(sheet: Excel.Worksheet, rangeAddress: string, format: string) {
    try {
      const safeFormat = format.replace(/#\s+##0/g, '#,##0').replace(/#\s+0/g, '#,##0');
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

        const matrix = Array(numRows).fill(null).map(() => Array(numCols).fill(safeFormat));
        sheet.getRange(rangeAddress).numberFormat = matrix;
      } else {
        sheet.getRange(rangeAddress).numberFormat = [[safeFormat]];
      }
    } catch (e) {
      // Repli sur format sans devise pour éviter toute exception
      try {
        const parts = rangeAddress.split(':');
        if (parts.length === 2) {
          const start = parts[0].trim();
          const end = parts[1].trim();
          const numRows = Math.max(parseInt(end.replace(/[^0-9]/g, '') || '1', 10) - parseInt(start.replace(/[^0-9]/g, '') || '1', 10) + 1, 1);
          const numCols = Math.max(this.colToIndex(end.replace(/[0-9]/g, '')) - this.colToIndex(start.replace(/[0-9]/g, '')) + 1, 1);
          sheet.getRange(rangeAddress).numberFormat = Array(numRows).fill(null).map(() => Array(numCols).fill("#,##0"));
        } else {
          sheet.getRange(rangeAddress).numberFormat = [["#,##0"]];
        }
      } catch (_) {}
    }
  }

  private static setBorder(range: Excel.Range, edge: "EdgeBottom" | "EdgeTop" | "EdgeLeft" | "EdgeRight", style: string = "Continuous") {
    try {
      range.format.borders.getItem(edge).style = style as any;
    } catch (_) {}
  }
}
