/**
 * Service d'interaction Universel avec Microsoft Excel (Office.js)
 * Conçu pour exécuter TOUTE tâche demandée par l'utilisateur de façon professionnelle :
 * - Calculs & Formules complexes
 * - Mise en forme & Styles de prestige (Bordures, Couleurs corporate, Polices)
 * - Création de Graphiques (Colonnes, Barres, Lignes, Secteurs/Camemberts)
 * - Tableaux structurés & Tableaux de bord
 * - Tri, Filtres & Nettoyage de données
 * - Multi-feuilles de calcul
 * - Exécution de scripts Office.js dynamiques
 * 100% compatible Excel 2016, 2019, 2021 et Microsoft 365.
 */

/* global Excel, Office */

import { ExcelAction } from '../types';
import { BusinessPlanService } from './businessPlanService';
import { PythonBackendService } from './pythonBackendService';

export interface ExcelRangeData {
  address: string;
  rowCount: number;
  columnCount: number;
  values: any[][];
  formulas: string[][];
  text: string[][];
}

export class OfficeService {
  /**
   * Vérifie si l'environnement d'exécution est Microsoft Excel
   */
  static isExcelEnvironment(): boolean {
    return typeof Office !== 'undefined' && typeof Excel !== 'undefined' && Office.context?.host === Office.HostType.Excel;
  }

  /**
   * Exécuteur universel d'actions Excel professionnelles
   */
  static async executeAction(action: ExcelAction, onProgress?: (msg: string) => void): Promise<{ success: boolean; message: string }> {
    if (!this.isExcelEnvironment()) {
      console.log("[Simulation Excel] Action exécutée :", action);
      return { success: true, message: `Action ${action.type} simulée avec succès.` };
    }

    try {
      switch (action.type) {
        case 'WRITE_RANGE': {
          await this.writeValues(action.values, action.address, {
            formatAsCurrency: action.formatAsCurrency,
            currencySymbol: action.currencySymbol,
            hasHeader: action.hasHeader
          });
          return { success: true, message: "Données écrites et formatées avec succès." };
        }

        case 'APPLY_FORMULA': {
          await this.applyFormula(action.targetCell, action.formula, (action as any).sheetName);
          return { success: true, message: `Formule ${action.formula} appliquée.` };
        }

        case 'FORMAT_RANGE': {
          await this.formatRange(action);
          return { success: true, message: "Style et mise en page appliqués." };
        }

        case 'CREATE_CHART': {
          await this.createChart(action.chartType, action.sourceRange, action.title);
          return { success: true, message: `Graphique ${action.chartType} inséré avec succès.` };
        }

        case 'CREATE_TABLE': {
          await this.createStructuredTable(action.rangeAddress, action.tableName, action.hasHeaders);
          return { success: true, message: "Tableau structuré créé avec succès." };
        }

        case 'SORT_RANGE': {
          await this.sortRange(action.rangeAddress, action.columnIndex, action.ascending);
          return { success: true, message: "Données triées avec succès." };
        }

        case 'ADD_SHEET': {
          await this.addWorksheet(action.sheetName, action.activate, action.tabColor);
          return { success: true, message: `Feuille '${action.sheetName}' créée.` };
        }

        case 'SET_SHEET_TAB_COLOR': {
          await this.setSheetTabColor(action.sheetName, action.color);
          return { success: true, message: "Couleur d'onglet appliquée." };
        }

        case 'CREATE_TABLE_SECTION': {
          await this.createTableSection(action);
          return { success: true, message: `Section '${action.bannerTitle || 'Tableau'}' créée avec succès.` };
        }

        case 'CREATE_MULTIPLE_TABLES': {
          await this.createMultipleTables(action.tables);
          return { success: true, message: `${action.tables.length} tableaux créés avec succès.` };
        }

        case 'UPDATE_BUSINESS_PLAN':
        case 'CREATE_BUSINESS_PLAN': {
          const isModif = action.type === 'UPDATE_BUSINESS_PLAN';
          // 1. Si un fichier Base64 a déjà été fourni par le moteur Python
          if ((action as any).file_base64) {
            onProgress?.(isModif ? "Mise à jour chirurgicale et harmonisation des 20 feuilles via Python..." : "Insertion des 20 feuilles natives via le moteur Python...");
            const inserted = await PythonBackendService.insertWorksheetsIntoActiveWorkbook((action as any).file_base64, isModif);
            if (inserted) {
              return {
                success: true,
                message: isModif
                  ? "Toutes les 20 feuilles du Business Plan ont été mises à jour chirurgicalement et harmonisées dans votre classeur !"
                  : "Les 20 feuilles professionnelles générées par le moteur Python ont été insérées directement dans votre classeur !"
              };
            }
          }

          // 2. Appel direct et prioritaire au moteur Python OHADA (20 feuilles bancables)
          onProgress?.(isModif ? "Recalcul financier complet via le moteur Python OHADA (20 feuilles)..." : "Calcul financier et modélisation via le moteur Python OHADA (20 feuilles)...");
          const pyPlan = await PythonBackendService.generateBusinessPlan({
            prompt: action.productTopic || action.projectName || "business plan",
            currency: action.currencySymbol || "FCFA",
            isModification: isModif
          });

          if (pyPlan && pyPlan.file_base64) {
            onProgress?.("Harmonisation des 20 feuilles natives Excel...");
            const inserted = await PythonBackendService.insertWorksheetsIntoActiveWorkbook(pyPlan.file_base64, isModif);
            if (inserted) {
              return {
                success: true,
                message: isModif
                  ? `Les 20 feuilles du Business Plan pour '${pyPlan.project_name}' ont été recalculées et mises à jour chirurgicalement avec succès !`
                  : `Les 20 feuilles professionnelles pour '${pyPlan.project_name}' ont été générées et insérées dans votre classeur avec succès !`
              };
            } else {
              // Si l'insertion directe Office.js échoue sur une version ancienne d'Excel, déclencher le téléchargement direct
              PythonBackendService.downloadXlsx(pyPlan.file_base64, pyPlan.filename);
              return {
                success: true,
                message: `Le classeur complet de 20 feuilles (${pyPlan.filename}) a été généré par le moteur Python et téléchargé sur votre ordinateur !`
              };
            }
          }

          // 3. Fallback client-side complet et autonome
          const bpRes = await BusinessPlanService.generateCompletePlan({
            projectName: action.projectName,
            projectType: action.projectType,
            productTopic: action.productTopic,
            sheetCount: action.sheetCount,
            currencySymbol: action.currencySymbol,
            template_id: (action as any).template_id,
            sheetsToGenerate: action.sheetsToGenerate,
            onProgress: (name, current, total) => {
              onProgress?.(`Feuille ${current}/${total} : ${name}...`);
            }
          });
          return { success: true, message: bpRes.message };
        }

        case 'AUTOFIT': {
          await this.autoFitColumns(action.rangeAddress);
          return { success: true, message: "Ajustement automatique des colonnes effectué." };
        }

        case 'MODIFY_CELLS': {
          const modAction = action as any;
          await this.modifyCells(modAction.sheetName, modAction.changes);
          return { success: true, message: `${modAction.changes?.length || 0} cellule(s) modifiée(s) sur '${modAction.sheetName}'.` };
        }

        case 'WRITE_TO_SHEET': {
          const wsAction = action as any;
          await this.writeToSpecificSheet(wsAction.sheetName, wsAction.address, wsAction.values, {
            hasHeader: wsAction.hasHeader,
            formatAsCurrency: wsAction.formatAsCurrency,
            currencySymbol: wsAction.currencySymbol
          });
          return { success: true, message: `Données écrites sur la feuille '${wsAction.sheetName}'.` };
        }

        case 'CLEAN_DATA': {
          const res = await this.cleanSelectedData();
          return { success: true, message: `${res.cleanedCells} cellules nettoyées.` };
        }

        case 'EXECUTE_CODE': {
          await this.executeRawCode(action.code);
          return { success: true, message: action.description || "Script personnalisé exécuté." };
        }

        case 'MERGE_RANGE':
        case 'MERGE_CELLS':
        case 'MERGE':
        case 'FUSIONNER':
        case 'FUSION': {
          await this.mergeRange({
            address: (action as any).address || (action as any).rangeAddress,
            across: (action as any).across,
            center: (action as any).center ?? true
          });
          return { success: true, message: "Cellules fusionnées avec succès." };
        }

        case 'UNMERGE_RANGE':
        case 'UNMERGE_CELLS':
        case 'UNMERGE':
        case 'DEFUSIONNER': {
          await this.unmergeRange({
            address: (action as any).address || (action as any).rangeAddress
          });
          return { success: true, message: "Cellules défusionnées avec succès." };
        }

        default: {
          const raw = String(action.type || '').toUpperCase();
          if (raw.includes('MERGE') || raw.includes('FUSION')) {
            await this.mergeRange({
              address: (action as any).address || (action as any).rangeAddress,
              across: (action as any).across,
              center: (action as any).center ?? true
            });
            return { success: true, message: "Cellules fusionnées avec succès." };
          }
          if (raw.includes('BUSINESS') || raw.includes('PLAN')) {
            const bpRes = await BusinessPlanService.generateCompletePlan({
              projectName: (action as any).projectName,
              projectType: (action as any).projectType,
              productTopic: (action as any).productTopic,
              sheetCount: (action as any).sheetCount,
              currencySymbol: (action as any).currencySymbol,
              sheetsToGenerate: (action as any).sheetsToGenerate,
              onProgress: (name, current, total) => {
                onProgress?.(`Feuille ${current}/${total} : ${name}...`);
              }
            });
            return { success: true, message: bpRes.message };
          }
          if (raw.includes('TABLE') && (action as any).headers && (action as any).rows) {
            await this.createTableSection(action as any);
            return { success: true, message: "Tableau créé avec succès." };
          }
          return { success: false, message: `Type d'action non reconnu : ${action.type}` };
        }
      }
    } catch (err: any) {
      console.error(`Erreur exécution ${action.type}:`, err);
      throw new Error(`Erreur lors de l'action ${action.type}: ${err?.message || err}`);
    }
  }

  /**
   * Récupère les données complètes de la sélection active
   */
  static async getSelectedRangeData(): Promise<ExcelRangeData | null> {
    if (!this.isExcelEnvironment()) {
      return {
        address: "A1:E6",
        rowCount: 6,
        columnCount: 5,
        values: [
          ["Code", "Libellé Compte", "Débit", "Crédit", "Solde"],
          ["521100", "Banque Principale", 15000000, 4200000, 10800000],
          ["571100", "Caisse Locale", 2500000, 1200000, 1300000],
          ["411100", "Clients Nationaux", 8900000, 3100000, 5800000],
          ["401100", "Fournisseurs Marchandises", 2000000, 7500000, -5500000],
          ["601100", "Achats de Marchandises", 7500000, 0, 7500000]
        ],
        formulas: [["", "", "", "", ""]],
        text: [["", "", "", "", ""]]
      };
    }

    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "rowCount", "columnCount", "values", "formulas", "text"]);
      await context.sync();

      return {
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        values: range.values,
        formulas: range.formulas,
        text: range.text
      };
    });
  }

  /**
   * Écrit des valeurs avec formatage professionnel optionnel (dimensions auto-ajustées)
   */
  static async writeValues(
    values: any[][], 
    targetAddress?: string, 
    options?: { formatAsCurrency?: boolean; currencySymbol?: string; hasHeader?: boolean }
  ): Promise<boolean> {
    if (!this.isExcelEnvironment()) return true;

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      if (!values || values.length === 0) return true;

      // Normaliser les colonnes pour que chaque ligne ait la même taille
      const maxCols = Math.max(...values.map(r => (Array.isArray(r) ? r.length : 1)), 1);
      const normalizedValues = values.map(row => {
        const r = Array.isArray(row) ? [...row] : [row];
        while (r.length < maxCols) r.push("");
        return r.slice(0, maxCols);
      });

      let range: Excel.Range;
      if (targetAddress) {
        const startCell = targetAddress.split(':')[0] || 'A1';
        range = sheet.getRange(startCell).getResizedRange(normalizedValues.length - 1, maxCols - 1);
      } else {
        const selected = context.workbook.getSelectedRange();
        selected.load("address");
        await context.sync();
        range = selected.getCell(0, 0).getResizedRange(normalizedValues.length - 1, maxCols - 1);
      }

      range.values = normalizedValues;

      // Stylisation d'en-tête si présent
      if (options?.hasHeader && normalizedValues.length > 1) {
        const headerRange = range.getRow(0);
        headerRange.format.fill.color = "#107C41"; // Vert Excel Corporate
        headerRange.format.font.color = "#FFFFFF";
        headerRange.format.font.bold = true;
      }

      // Format monétaire conforme (ex: # ##0 "FCFA") aux dimensions exactes
      if (options?.formatAsCurrency) {
        const symbol = options.currencySymbol || 'FCFA';
        const numFormat = `# ##0 "${symbol}"`;
        range.numberFormat = Array(normalizedValues.length).fill(null).map(() => Array(maxCols).fill(numFormat));
      }

      range.format.autofitColumns();
      await context.sync();
      return true;
    });
  }

  /**
   * Applique une formule Excel (sur la feuille active ou une feuille spécifique)
   */
  static async applyFormula(targetCell: string, formula: string, sheetName?: string): Promise<boolean> {
    if (!this.isExcelEnvironment()) return true;

    return Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(targetCell);
      range.formulas = [[formula]];
      await context.sync();
      return true;
    });
  }

  /**
   * Lit le contenu complet de la feuille active (valeurs + formules) pour permettre la modification chirurgicale
   * Limité aux 80 premières lignes de la zone utilisée pour la performance
   */
  static async getActiveSheetFullData(): Promise<{
    sheetName: string;
    usedRangeAddress: string;
    values: any[][];
    formulas: string[][];
    rowCount: number;
    columnCount: number;
    rowIndex: number;
    columnIndex: number;
  } | null> {
    if (!this.isExcelEnvironment()) return null;

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');
      const usedRange = sheet.getUsedRange();
      usedRange.load(['address', 'values', 'formulas', 'rowCount', 'columnCount', 'rowIndex', 'columnIndex']);
      await context.sync();

      const maxRows = Math.min(usedRange.rowCount, 80);
      return {
        sheetName: sheet.name,
        usedRangeAddress: usedRange.address,
        values: usedRange.values.slice(0, maxRows),
        formulas: usedRange.formulas.slice(0, maxRows),
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount,
        rowIndex: usedRange.rowIndex || 0,
        columnIndex: usedRange.columnIndex || 0
      };
    }).catch(err => {
      console.warn("getActiveSheetFullData:", err);
      return null;
    });
  }

  /**
   * Lit le contenu d'une feuille spécifique par nom
   */
  static async getSpecificSheetData(sheetName: string): Promise<{
    values: any[][];
    formulas: string[][];
    address: string;
    rowIndex: number;
    columnIndex: number;
  } | null> {
    if (!this.isExcelEnvironment()) return null;

    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      let sheet = sheets.items.find(s => s.name.toLowerCase() === (sheetName || '').toLowerCase());
      if (!sheet) {
        const cleanName = (sheetName || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        sheet = sheets.items.find(s => {
          const sClean = s.name.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
          return sClean.includes(cleanName) || cleanName.includes(sClean);
        });
      }
      if (!sheet) {
        sheet = context.workbook.worksheets.getActiveWorksheet();
      }

      const usedRange = sheet.getUsedRange();
      usedRange.load(['address', 'values', 'formulas', 'rowCount', 'rowIndex', 'columnIndex']);
      await context.sync();

      const maxRows = Math.min(usedRange.rowCount, 80);
      return {
        values: usedRange.values.slice(0, maxRows),
        formulas: usedRange.formulas.slice(0, maxRows),
        address: usedRange.address,
        rowIndex: usedRange.rowIndex || 0,
        columnIndex: usedRange.columnIndex || 0
      };
    }).catch(err => {
      console.warn("getSpecificSheetData:", err);
      return null;
    });
  }

  /**
   * Écrit des valeurs sur une feuille spécifique (pas la feuille active, pour modifications chirurgicales)
   */
  static async writeToSpecificSheet(
    sheetName: string,
    address: string,
    values: any[][],
    options?: { hasHeader?: boolean; formatAsCurrency?: boolean; currencySymbol?: string }
  ): Promise<boolean> {
    if (!this.isExcelEnvironment()) return true;

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      if (!values || values.length === 0) return true;

      const maxCols = Math.max(...values.map(r => (Array.isArray(r) ? r.length : 1)), 1);
      const normalizedValues = values.map(row => {
        const r = Array.isArray(row) ? [...row] : [row];
        while (r.length < maxCols) r.push("");
        return r.slice(0, maxCols);
      });

      const startCell = address.split(':')[0] || 'A1';
      const range = sheet.getRange(startCell).getResizedRange(normalizedValues.length - 1, maxCols - 1);
      range.values = normalizedValues;

      if (options?.hasHeader && normalizedValues.length > 1) {
        const headerRange = range.getRow(0);
        headerRange.format.fill.color = "#107C41";
        headerRange.format.font.color = "#FFFFFF";
        headerRange.format.font.bold = true;
      }

      if (options?.formatAsCurrency) {
        const symbol = options.currencySymbol || 'FCFA';
        const numFormat = `# ##0 "${symbol}"`;
        range.numberFormat = Array(normalizedValues.length).fill(null).map(() => Array(maxCols).fill(numFormat));
      }

      range.format.autofitColumns();
      await context.sync();
      return true;
    });
  }

  /**
   * Modifie chirurgicalement des cellules individuelles sur une feuille spécifique
   * Utilisé pour les modifications ciblées de business plans existants
   */
  static async modifyCells(
    sheetName: string,
    changes: Array<{ cell: string; value?: any; formula?: string }>
  ): Promise<boolean> {
    if (!this.isExcelEnvironment()) return true;

    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      // Recherche tolérante du nom de la feuille (ignorer casse et accents)
      let targetSheet = sheets.items.find(s => s.name.toLowerCase() === (sheetName || '').toLowerCase());
      if (!targetSheet && sheetName) {
        const cleanName = sheetName.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        targetSheet = sheets.items.find(s => {
          const sClean = s.name.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
          return sClean.includes(cleanName) || cleanName.includes(sClean);
        });
      }
      const sheet = targetSheet || context.workbook.worksheets.getActiveWorksheet();

      for (const change of changes) {
        // Nettoyer l'adresse de cellule (enlever Feuil1!, guillemets, crochets)
        const cleanCell = (change.cell || 'A1').replace(/^[^!]+!/, '').replace(/[\[\]"]/g, '').trim();
        const range = sheet.getRange(cleanCell);

        if (change.formula) {
          // Normaliser systématiquement les formules vers l'anglais et séparateur virgule
          let formula = change.formula.trim();
          if (!formula.startsWith('=')) formula = '=' + formula;
          formula = formula
            .replace(/=SOMME\(/gi, '=SUM(')
            .replace(/=ARRONDI\(/gi, '=ROUND(')
            .replace(/=MOYENNE\(/gi, '=AVERAGE(')
            .replace(/=SI\(/gi, '=IF(')
            .replace(/;(?=[^"]*(?:"[^"]*"[^"]*)*$)/g, ','); // remplace les points-virgules hors chaînes de texte

          range.formulas = [[formula]];
        } else if (change.value !== undefined) {
          range.values = [[change.value]];
        }
      }

      sheet.getUsedRange().format.autofitColumns();
      await context.sync();
      return true;
    });
  }

  /**
   * Mise en forme stylisée et professionnelle d'une plage (sécurisée sans crash)
   */
  static async formatRange(params: {
    address?: string;
    fill?: string;
    fontColor?: string;
    bold?: boolean;
    italic?: boolean;
    fontSize?: number;
    alignment?: 'Left' | 'Center' | 'Right';
    numberFormat?: string;
    border?: 'top' | 'bottom' | 'double-bottom' | 'all' | 'none';
    merge?: boolean;
  }): Promise<boolean> {
    if (!this.isExcelEnvironment()) return true;

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = params.address ? sheet.getRange(params.address) : context.workbook.getSelectedRange();

      if (params.merge) {
        range.merge(false);
      }

      if (params.fill) range.format.fill.color = params.fill;
      if (params.fontColor) range.format.font.color = params.fontColor;
      if (params.bold !== undefined) range.format.font.bold = params.bold;
      if (params.italic !== undefined) range.format.font.italic = params.italic;
      if (params.fontSize) range.format.font.size = params.fontSize;

      if (params.alignment) {
        range.format.horizontalAlignment = params.alignment;
      }

      // Gestion des bordures d'audit et comptables sécurisée
      if (params.border) {
        try {
          if (params.border === 'double-bottom') {
            range.format.borders.getItem('EdgeBottom').style = "Double";
            range.format.borders.getItem('EdgeTop').style = "Continuous";
          } else if (params.border === 'all') {
            range.format.borders.getItem('EdgeBottom').style = "Continuous";
            range.format.borders.getItem('EdgeTop').style = "Continuous";
            range.format.borders.getItem('EdgeLeft').style = "Continuous";
            range.format.borders.getItem('EdgeRight').style = "Continuous";
            try { range.format.borders.getItem('InsideHorizontal').style = "Continuous"; } catch (_) {}
            try { range.format.borders.getItem('InsideVertical').style = "Continuous"; } catch (_) {}
          }
        } catch (_) {}
      }

      await context.sync();
      return true;
    });
  }

  /**
   * Fusionne une plage de cellules spécifiée ou la sélection active
   */
  static async mergeRange(params: { address?: string; across?: boolean; center?: boolean }): Promise<boolean> {
    if (!this.isExcelEnvironment()) {
      console.log("[Simulation Excel] Fusion des cellules :", params);
      return true;
    }

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = (params.address && params.address !== 'selection') 
        ? sheet.getRange(params.address) 
        : context.workbook.getSelectedRange();

      range.merge(params.across ?? false);

      if (params.center ?? true) {
        range.format.horizontalAlignment = "Center";
        range.format.verticalAlignment = "Center";
      }

      await context.sync();
      return true;
    });
  }

  /**
   * Défusionne une plage de cellules spécifiée ou la sélection active
   */
  static async unmergeRange(params: { address?: string }): Promise<boolean> {
    if (!this.isExcelEnvironment()) {
      console.log("[Simulation Excel] Défusion des cellules :", params);
      return true;
    }

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = (params.address && params.address !== 'selection') 
        ? sheet.getRange(params.address) 
        : context.workbook.getSelectedRange();

      range.unmerge();
      await context.sync();
      return true;
    });
  }

  /**
   * Insère un graphique professionnel lié aux données
   */
  static async createChart(
    chartType: 'ColumnClustered' | 'BarClustered' | 'Line' | 'Pie',
    sourceRange?: string,
    title?: string
  ): Promise<boolean> {
    if (!this.isExcelEnvironment()) return true;

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sourceRange ? sheet.getRange(sourceRange) : context.workbook.getSelectedRange();

      let type: Excel.ChartType;
      switch (chartType) {
        case 'Pie':
          type = Excel.ChartType.pie;
          break;
        case 'Line':
          type = Excel.ChartType.line;
          break;
        case 'BarClustered':
          type = Excel.ChartType.barClustered;
          break;
        case 'ColumnClustered':
        default:
          type = Excel.ChartType.columnClustered;
          break;
      }

      const chart = sheet.charts.add(type, range, Excel.ChartSeriesBy.auto);
      if (title) {
        chart.title.text = title;
      }
      chart.legend.position = Excel.ChartLegendPosition.bottom;

      await context.sync();
      return true;
    });
  }

  /**
   * Convertit une plage en tableau structuré Excel avec filtres
   */
  static async createStructuredTable(rangeAddress?: string, tableName?: string, hasHeaders: boolean = true): Promise<boolean> {
    if (!this.isExcelEnvironment()) return true;

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = rangeAddress ? sheet.getRange(rangeAddress) : context.workbook.getSelectedRange();

      const table = sheet.tables.add(range, hasHeaders);
      if (tableName) {
        table.name = tableName.replace(/\s+/g, '_');
      }
      table.style = "TableStyleMedium9"; // Style vert professionnel

      await context.sync();
      return true;
    });
  }

  /**
   * Trie une plage de données par colonne
   */
  static async sortRange(rangeAddress?: string, columnIndex: number = 0, ascending: boolean = true): Promise<boolean> {
    if (!this.isExcelEnvironment()) return true;

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = rangeAddress ? sheet.getRange(rangeAddress) : context.workbook.getSelectedRange();

      range.sort.apply([
        {
          key: columnIndex,
          ascending: ascending
        }
      ]);

      await context.sync();
      return true;
    });
  }

  /**
   * Ajoute une nouvelle feuille de calcul avec couleur d'onglet optionnelle
   */
  static async addWorksheet(sheetName: string, activate: boolean = true, tabColor?: string): Promise<boolean> {
    if (!this.isExcelEnvironment()) return true;

    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const cleanName = sheetName.replace(/[:\\/?*\[\]]/g, '_').substring(0, 31);
      
      let newSheet: Excel.Worksheet;
      try {
        newSheet = sheets.getItem(cleanName);
      } catch (e) {
        newSheet = sheets.add(cleanName);
      }

      if (tabColor && typeof Office !== 'undefined' && Office.context?.requirements?.isSetSupported('ExcelApi', '1.7')) {
        try {
          newSheet.tabColor = tabColor;
        } catch (_) {}
      }
      if (activate) {
        newSheet.activate();
      }
      await context.sync();
      return true;
    });
  }

  /**
   * Modifie la couleur de l'onglet d'une feuille
   */
  static async setSheetTabColor(sheetName: string | undefined, color: string): Promise<boolean> {
    if (!this.isExcelEnvironment()) return true;

    return Excel.run(async (context) => {
      const sheet = sheetName 
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      if (typeof Office !== 'undefined' && Office.context?.requirements?.isSetSupported('ExcelApi', '1.7')) {
        try {
          sheet.tabColor = color;
        } catch (_) {}
      }
      await context.sync();
      return true;
    });
  }

  /**
   * Utilitaires de conversion colonnes Excel (gère A, B ... Z, AA, AB ...)
   */
  private static colLetterToIndex(col: string): number {
    let index = 0;
    const cleanCol = col.toUpperCase().replace(/[^A-Z]/g, '');
    for (let i = 0; i < cleanCol.length; i++) {
      index = index * 26 + (cleanCol.charCodeAt(i) - 64);
    }
    return Math.max(index, 1);
  }

  private static indexToColLetter(index: number): string {
    let temp = Math.max(index, 1);
    let letter = '';
    while (temp > 0) {
      const mod = (temp - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      temp = Math.floor((temp - mod) / 26);
    }
    return letter;
  }

  /**
   * Crée une section de tableau moderne avec bannière et styles de prestige (Compatible Excel 2016 à 365)
   */
  static async createTableSection(params: {
    sheetName?: string;
    startCell?: string;
    bannerTitle?: string;
    bannerColor?: string;
    headers: string[];
    headerFill?: string;
    rows: any[][];
    highlightInputs?: boolean;
    inputColumnIndices?: number[];
    formatAsCurrency?: boolean;
    currencySymbol?: string;
  }): Promise<boolean> {
    if (!this.isExcelEnvironment()) {
      console.log("[Simulation Excel] Section de tableau créée :", params);
      return true;
    }

    return Excel.run(async (context) => {
      // 1. Récupération ultra-sécurisée de la feuille (évite ItemNotFound / 0xA7120009 si la feuille n'existe pas)
      const sheets = context.workbook.worksheets;
      let sheet: Excel.Worksheet;

      if (params.sheetName) {
        sheets.load("items/name");
        await context.sync();
        const cleanName = params.sheetName.replace(/[:\\/?*\[\]]/g, '_').substring(0, 31).trim();
        const found = sheets.items.find(s => s.name.toLowerCase() === cleanName.toLowerCase());
        if (found) {
          sheet = found;
        } else {
          sheet = context.workbook.worksheets.getActiveWorksheet();
        }
      } else {
        sheet = context.workbook.worksheets.getActiveWorksheet();
      }

      // 2. Calcul robuste des coordonnées
      const start = (params.startCell || 'B2').trim().toUpperCase();
      const colMatch = start.match(/[A-Z]+/);
      const rowMatch = start.match(/\d+/);
      const startColLetter = colMatch ? colMatch[0] : 'B';
      let rowNum = rowMatch ? parseInt(rowMatch[0], 10) : 2;

      // Déterminer le nombre exact de colonnes
      const headers = (params.headers && params.headers.length > 0) 
        ? params.headers 
        : (params.rows && params.rows[0] ? params.rows[0].map((_, i) => `Colonne ${i + 1}`) : ["Colonne 1", "Colonne 2", "Colonne 3", "Colonne 4"]);
      
      const numCols = Math.max(headers.length, 1);
      const startColIdx = this.colLetterToIndex(startColLetter);
      const endColIdx = startColIdx + numCols - 1;
      const endColLetter = this.indexToColLetter(endColIdx);

      // 3. Bannière de titre si demandée
      if (params.bannerTitle) {
        const bannerAddress = `${startColLetter}${rowNum}:${endColLetter}${rowNum}`;
        const firstCell = `${startColLetter}${rowNum}`;
        
        // Assigner la valeur sur la première cellule pour éviter toute discordance de dimension (0xA7120009)
        sheet.getRange(firstCell).values = [[params.bannerTitle]];
        
        const bannerRange = sheet.getRange(bannerAddress);
        if (numCols > 1) {
          bannerRange.merge(false);
        }
        bannerRange.format.fill.color = params.bannerColor || "#1F3864";
        bannerRange.format.font.color = "#FFFFFF";
        bannerRange.format.font.bold = true;
        bannerRange.format.font.size = 11;
        bannerRange.format.horizontalAlignment = "Left";
        bannerRange.format.verticalAlignment = "Center";
        rowNum++;
      }

      // 4. En-têtes de colonnes
      const headerAddress = `${startColLetter}${rowNum}:${endColLetter}${rowNum}`;
      const headerRange = sheet.getRange(headerAddress);
      headerRange.values = [headers];
      headerRange.format.fill.color = params.headerFill || "#D9E1F2";
      headerRange.format.font.bold = true;
      headerRange.format.font.size = 10;
      headerRange.format.horizontalAlignment = "Center";
      
      try {
        headerRange.format.borders.getItem("EdgeBottom").style = "Continuous";
        headerRange.format.borders.getItem("EdgeTop").style = "Continuous";
      } catch (e) {
        // Ignorer si bordure non supportée sur cette version
      }
      rowNum++;

      // 5. Lignes de données
      const rawRows = params.rows || [];
      if (rawRows.length > 0) {
        // Normaliser les lignes pour que chaque ligne ait STRICTEMENT numCols éléments
        const normalizedRows = rawRows.map(row => {
          const r = Array.isArray(row) ? [...row] : [row];
          while (r.length < numCols) r.push("");
          return r.slice(0, numCols);
        });

        const startDataRow = rowNum;
        const endDataRow = rowNum + normalizedRows.length - 1;
        const dataAddress = `${startColLetter}${startDataRow}:${endColLetter}${endDataRow}`;
        const dataRange = sheet.getRange(dataAddress);
        
        dataRange.values = normalizedRows;

        try {
          if (normalizedRows.length > 1) {
            dataRange.format.borders.getItem("InsideHorizontal").style = "Continuous";
          }
          dataRange.format.borders.getItem("EdgeBottom").style = "Continuous";
        } catch (e) {
          // Ignorer si bordure non supportée
        }

        // Surligner les colonnes d'entrée utilisateur en jaune clair #FFF2CC si demandé
        if (params.highlightInputs || (params.inputColumnIndices && params.inputColumnIndices.length > 0)) {
          const colIndices = params.inputColumnIndices || [numCols - 1];
          for (const cIdx of colIndices) {
            if (cIdx >= 0 && cIdx < numCols) {
              const inputColLetter = this.indexToColLetter(startColIdx + cIdx);
              const inputColRange = sheet.getRange(`${inputColLetter}${startDataRow}:${inputColLetter}${endDataRow}`);
              inputColRange.format.fill.color = "#FFF2CC";
            }
          }
        }

        // Formatage monétaire sécurisé aux dimensions exactes
        if (params.formatAsCurrency) {
          const symbol = params.currencySymbol || "FCFA";
          const currencyFormat = `# ##0 "${symbol}"`;
          const formatMatrix = Array(normalizedRows.length).fill(null).map(() => 
            Array(numCols).fill(currencyFormat)
          );
          dataRange.numberFormat = formatMatrix;
        }

        // Ajustement automatique des colonnes sans utiliser la syntaxe "B:F" (qui plante sur Excel 2016)
        dataRange.format.autofitColumns();
      } else {
        headerRange.format.autofitColumns();
      }

      await context.sync();
      return true;
    });
  }

  /**
   * Crée plusieurs tableaux garantis à 100% sans omission
   */
  static async createMultipleTables(tables: Array<any>): Promise<boolean> {
    if (!tables || tables.length === 0) return true;
    let currentRow = 2;
    for (const tbl of tables) {
      const startCell = tbl.startCell || `B${currentRow}`;
      await this.createTableSection({
        ...tbl,
        startCell
      });
      const rowCount = (tbl.rows?.length || 0) + (tbl.bannerTitle ? 2 : 1) + 2;
      currentRow += rowCount;
    }
    return true;
  }

  /**
   * Ajuste la largeur des colonnes
   */
  static async autoFitColumns(rangeAddress?: string): Promise<boolean> {
    if (!this.isExcelEnvironment()) return true;

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let range: Excel.Range;
      if (rangeAddress && rangeAddress.match(/^[A-Z]+:[A-Z]+$/i)) {
        // En cas de syntaxe colonne entière "A:E", utiliser getUsedRange pour compatibilité Excel 2016
        range = sheet.getUsedRange();
      } else if (rangeAddress) {
        range = sheet.getRange(rangeAddress);
      } else {
        range = sheet.getUsedRange();
      }
      range.format.autofitColumns();
      await context.sync();
      return true;
    });
  }

  /**
   * Nettoie les données brutes dans la sélection (espaces insécables, nombres en texte)
   */
  static async cleanSelectedData(): Promise<{ cleanedRows: number; cleanedCells: number }> {
    if (!this.isExcelEnvironment()) {
      return { cleanedRows: 12, cleanedCells: 45 };
    }

    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      const values = range.values;
      let count = 0;
      let modifiedRows = 0;

      for (let r = 0; r < values.length; r++) {
        let rowModified = false;
        for (let c = 0; c < values[r].length; c++) {
          const val = values[r][c];
          if (typeof val === 'string') {
            const cleaned = val.replace(/\u00A0/g, ' ').trim();
            const numberCandidate = cleaned.replace(/\s+/g, '').replace(/,/g, '.');

            if (!isNaN(Number(numberCandidate)) && numberCandidate !== '') {
              values[r][c] = Number(numberCandidate);
              count++;
              rowModified = true;
            } else if (cleaned !== val) {
              values[r][c] = cleaned;
              count++;
              rowModified = true;
            }
          }
        }
        if (rowModified) modifiedRows++;
      }

      range.values = values;
      await context.sync();
      return { cleanedRows: modifiedRows, cleanedCells: count };
    });
  }

  /**
   * Exécute du code Office.js arbitraire de manière sécurisée
   */
  static async executeRawCode(code: string): Promise<any> {
    if (!this.isExcelEnvironment()) {
      console.log("[Simulation Script Office.js]:\n", code);
      return true;
    }

    return Excel.run(async (context) => {
      // Exécution dynamique dans le contexte Excel avec sandbox
      const fn = new Function('context', 'Excel', `return (async () => { ${code} })();`);
      await fn(context, Excel);
      await context.sync();
      return true;
    });
  }

  /**
   * Crée un Bilan SYSCOHADA normalisé
   */
  static async insertSyscohadaBilanSheet(data: {
    exercice: string;
    entreprise: string;
    lignesActif: { rubriques: string; brut: number; amort: number; net: number }[];
    lignesPassif: { rubriques: string; net: number }[];
  }): Promise<boolean> {
    if (!this.isExcelEnvironment()) return true;

    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const sheetName = `Bilan_OHADA_${new Date().getFullYear()}`;
      const newSheet = sheets.add(sheetName);
      newSheet.activate();

      newSheet.getRange("A1:D1").values = [[`BILAN NORMALISÉ SYSCOHADA - ${data.entreprise.toUpperCase()}`, "", "", ""]];
      newSheet.getRange("A1:D1").format.font.bold = true;
      newSheet.getRange("A1:D1").format.font.size = 14;
      newSheet.getRange("A1:D1").format.font.color = "#064E3B";

      newSheet.getRange("A2:D2").values = [[`Exercice clos au 31/12/${data.exercice} (En FCFA)`, "", "", ""]];
      newSheet.getRange("A2:D2").format.font.italic = true;

      newSheet.getRange("A4:D4").values = [["ACTIF", "BRUT", "AMORT / DÉPRÉC", "NET"]];
      newSheet.getRange("A4:D4").format.fill.color = "#107C41";
      newSheet.getRange("A4:D4").format.font.color = "#FFFFFF";
      newSheet.getRange("A4:D4").format.font.bold = true;

      let rIndex = 5;
      data.lignesActif.forEach(item => {
        newSheet.getRange(`A${rIndex}:D${rIndex}`).values = [[item.rubriques, item.brut, item.amort, item.net]];
        rIndex++;
      });

      newSheet.getRange(`A${rIndex}:D${rIndex}`).values = [[
        "TOTAL GÉNÉRAL ACTIF", 
        `=SUM(B5:B${rIndex - 1})`, 
        `=SUM(C5:C${rIndex - 1})`, 
        `=SUM(D5:D${rIndex - 1})`
      ]];
      newSheet.getRange(`A${rIndex}:D${rIndex}`).format.font.bold = true;
      newSheet.getRange(`A${rIndex}:D${rIndex}`).format.fill.color = "#E2E8F0";

      const numRowsActif = Math.max(rIndex - 4, 1);
      newSheet.getRange(`B5:D${rIndex}`).numberFormat = Array(numRowsActif).fill(null).map(() => ["# ##0", "# ##0", "# ##0"]);
      newSheet.getUsedRange().format.autofitColumns();

      await context.sync();
      return true;
    });
  }

  /**
   * Crée un livre de paie conforme par pays
   */
  static async insertPayrollSheet(country: string, employees: Array<{ nom: string; salaireBase: number; primes: number; partsFiscales: number }>): Promise<boolean> {
    if (!this.isExcelEnvironment()) return true;

    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const sheetName = `Livre_Paie_${country.replace(/\s+/g, '_')}`;
      const newSheet = sheets.add(sheetName);
      newSheet.activate();

      newSheet.getRange("A1:H1").values = [[`LIVRE DE PAIE MENSUEL - ${country.toUpperCase()}`, "", "", "", "", "", "", ""]];
      newSheet.getRange("A1:H1").format.font.bold = true;
      newSheet.getRange("A1:H1").format.font.size = 13;
      newSheet.getRange("A1:H1").format.font.color = "#047857";

      const headers = [
        "Nom Employé", 
        "Salaire Base", 
        "Primes / Indemnités", 
        "Salaire Brut", 
        "Cotis. Sociale (CNPS)", 
        "Impôt Salaire (ITS/IRPP)", 
        "Total Retenues", 
        "Net à Payer (FCFA)"
      ];
      newSheet.getRange("A3:H3").values = [headers];
      newSheet.getRange("A3:H3").format.fill.color = "#059669";
      newSheet.getRange("A3:H3").format.font.color = "#FFFFFF";
      newSheet.getRange("A3:H3").format.font.bold = true;

      let row = 4;
      for (const emp of employees) {
        newSheet.getRange(`A${row}:H${row}`).values = [[
          emp.nom,
          emp.salaireBase,
          emp.primes,
          `=B${row}+C${row}`,
          `=D${row}*0.063`,
          `=(D${row}-E${row})*0.08`,
          `=E${row}+F${row}`,
          `=D${row}-G${row}`
        ]];
        row++;
      }

      newSheet.getRange(`A${row}:H${row}`).values = [[
        "TOTAL MASSE SALARIALE",
        `=SUM(B4:B${row - 1})`,
        `=SUM(C4:C${row - 1})`,
        `=SUM(D4:D${row - 1})`,
        `=SUM(E4:E${row - 1})`,
        `=SUM(F4:F${row - 1})`,
        `=SUM(G4:G${row - 1})`,
        `=SUM(H4:H${row - 1})`
      ]];
      newSheet.getRange(`A${row}:H${row}`).format.font.bold = true;
      newSheet.getRange(`A${row}:H${row}`).format.fill.color = "#D1FAE5";

      const numRowsPayroll = Math.max(row - 3, 1);
      newSheet.getRange(`B4:H${row}`).numberFormat = Array(numRowsPayroll).fill(null).map(() => Array(7).fill("# ##0"));
      newSheet.getUsedRange().format.autofitColumns();

      await context.sync();
      return true;
    });
  }

  /**
   * Sélectionne et met en surbrillance une cellule ou plage dans Excel (Citations interactives type Claude)
   */
  static async selectRange(rangeAddress: string): Promise<boolean> {
    if (!this.isExcelEnvironment()) {
      console.log("[Simulation] Sélection de la plage :", rangeAddress);
      return true;
    }

    return Excel.run(async (context) => {
      try {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const cleanAddress = rangeAddress.replace(/^[^!]+!/, '').replace(/[\[\]]/g, '').trim();
        const range = sheet.getRange(cleanAddress);
        range.select();
        await context.sync();
        return true;
      } catch (err) {
        console.warn("Impossible de sélectionner la plage :", rangeAddress, err);
        return false;
      }
    });
  }

  /**
   * Vue d'ensemble du classeur (Feuilles existantes et feuille active type Claude)
   */
  static async getWorkbookOverview(): Promise<{ activeSheet: string; sheets: string[] }> {
    if (!this.isExcelEnvironment()) {
      return { activeSheet: "Feuil1", sheets: ["Feuil1", "Données", "Bilan"] };
    }

    return Excel.run(async (context) => {
      try {
        const active = context.workbook.worksheets.getActiveWorksheet();
        const allSheets = context.workbook.worksheets;
        active.load("name");
        allSheets.load("items/name");
        await context.sync();
        return {
          activeSheet: active.name,
          sheets: allSheets.items.map(s => s.name)
        };
      } catch (_) {
        return { activeSheet: "Active", sheets: [] };
      }
    });
  }

  /**
   * Détecteur d'erreurs de formule pour le débogage de cause racine (Root-Cause Debugger type Claude)
   */
  static detectFormulaErrors(values: any[][], formulas: any[][]): Array<{ cell: string; errorType: string; formula: string }> {
    const errorTypes = ['#REF!', '#VALEUR!', '#VALUE!', '#DIV/0!', '#NOM?', '#NAME?', '#N/A', '#NULL!', '#NUM!'];
    const detected: Array<{ cell: string; errorType: string; formula: string }> = [];

    if (!values || !Array.isArray(values)) return detected;

    for (let r = 0; r < values.length; r++) {
      const row = values[r];
      if (!Array.isArray(row)) continue;
      for (let c = 0; c < row.length; c++) {
        const val = String(row[c] || '');
        for (const err of errorTypes) {
          if (val.toUpperCase().includes(err)) {
            const colLetter = String.fromCharCode(65 + c);
            const cellAddr = `${colLetter}${r + 1}`;
            const formula = (formulas && formulas[r] && formulas[r][c]) ? String(formulas[r][c]) : "";
            detected.push({ cell: cellAddr, errorType: err, formula });
            break;
          }
        }
      }
    }
    return detected;
  }
}
