export type AfricanCountry = 
  | 'CI' // Côte d'Ivoire (XOF - UEMOA)
  | 'SN' // Sénégal (XOF - UEMOA)
  | 'CM' // Cameroun (XAF - CEMAC)
  | 'CD' // RD Congo (CDF)
  | 'BF' // Burkina Faso (XOF - UEMOA)
  | 'ML' // Mali (XOF - UEMOA)
  | 'BJ' // Bénin (XOF - UEMOA)
  | 'TG' // Togo (XOF - UEMOA)
  | 'GN' // Guinée (GNF)
  | 'GA'; // Gabon (XAF - CEMAC)

export interface CountryInfo {
  code: AfricanCountry;
  name: string;
  currency: string;
  currencySymbol: string;
  flag: string;
  tvaRate: number;
  cnssLabel: string;
  taxLabel: string;
}

export const SUPPORTED_COUNTRIES: Record<AfricanCountry, CountryInfo> = {
  CI: { code: 'CI', name: "Côte d'Ivoire", currency: 'XOF', currencySymbol: 'FCFA', flag: '🇨🇮', tvaRate: 18, cnssLabel: 'CNPS (6.3%)', taxLabel: 'ITS / IGR' },
  SN: { code: 'SN', name: "Sénégal", currency: 'XOF', currencySymbol: 'FCFA', flag: '🇸🇳', tvaRate: 18, cnssLabel: 'IPRES / CSS (8.4%)', taxLabel: 'IR (Barème progressif)' },
  CM: { code: 'CM', name: "Cameroun", currency: 'XAF', currencySymbol: 'FCFA', flag: '🇨🇲', tvaRate: 19.25, cnssLabel: 'CNPS (4.2%)', taxLabel: 'IRPP + CAC (10%)' },
  CD: { code: 'CD', name: "RD Congo", currency: 'CDF', currencySymbol: 'CDF', flag: '🇨🇩', tvaRate: 16, cnssLabel: 'CNSS (5%)', taxLabel: 'IPR (3% à 40%)' },
  BF: { code: 'BF', name: "Burkina Faso", currency: 'XOF', currencySymbol: 'FCFA', flag: '🇧🇫', tvaRate: 18, cnssLabel: 'CNSS (5.5%)', taxLabel: 'IUTS' },
  ML: { code: 'ML', name: "Mali", currency: 'XOF', currencySymbol: 'FCFA', flag: '🇲🇱', tvaRate: 18, cnssLabel: 'INPS (3.6%)', taxLabel: 'ITS' },
  BJ: { code: 'BJ', name: "Bénin", currency: 'XOF', currencySymbol: 'FCFA', flag: '🇧🇯', tvaRate: 18, cnssLabel: 'CNSS (3.6%)', taxLabel: 'AIB / IPTS' },
  TG: { code: 'TG', name: "Togo", currency: 'XOF', currencySymbol: 'FCFA', flag: '🇹🇬', tvaRate: 18, cnssLabel: 'CNSS (4%)', taxLabel: 'IRPP' },
  GN: { code: 'GN', name: "Guinée", currency: 'GNF', currencySymbol: 'GNF', flag: '🇬🇳', tvaRate: 18, cnssLabel: 'CNSS (5%)', taxLabel: 'RTS' },
  GA: { code: 'GA', name: "Gabon", currency: 'XAF', currencySymbol: 'FCFA', flag: '🇬🇦', tvaRate: 18, cnssLabel: 'CNSS (2.5%)', taxLabel: 'IRPP' },
};

export interface ChatMessage {
  id: string;
  sender: 'user' | 'assistant';
  text: string;
  timestamp: Date;
  actions?: ExcelAction[];
  suggestedPrompt?: string;
  status?: 'pending' | 'applied' | 'error';
  thought?: string; // Raisonnement & analyse financière préalable de l'agent
}

export interface BusinessPlanProject {
  id: string;
  title: string;
  productTopic: string;
  country: AfricanCountry;
  createdAt: string;
  updatedAt: string;
  messages: ChatMessage[];
  keyMetrics?: {
    investmentTotal?: number;
    equity?: number;
    loan?: number;
    annualTurnover?: number;
    netMargin?: number;
    irr?: string;
    breakEvenPoint?: string;
    dscr?: string;
  };
}

// Catalogue complet des actions professionnelles exécutables dans Excel
export type ExcelAction = 
  | { 
      type: 'WRITE_RANGE'; 
      address?: string; 
      values: any[][]; 
      formatAsCurrency?: boolean; 
      currencySymbol?: string; 
      hasHeader?: boolean 
    }
  | { 
      type: 'APPLY_FORMULA'; 
      targetCell: string; 
      formula: string; 
      explanation?: string 
    }
  | { 
      type: 'FORMAT_RANGE'; 
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
    }
  | { 
      type: 'MERGE_RANGE' | 'MERGE_CELLS' | 'MERGE' | 'FUSIONNER'; 
      address?: string; 
      across?: boolean; 
      center?: boolean; 
    }
  | { 
      type: 'UNMERGE_RANGE' | 'UNMERGE_CELLS' | 'UNMERGE' | 'DEFUSIONNER'; 
      address?: string; 
    }
  | { 
      type: 'CREATE_CHART'; 
      chartType: 'ColumnClustered' | 'BarClustered' | 'Line' | 'Pie'; 
      sourceRange?: string; 
      title?: string 
    }
  | { 
      type: 'CREATE_TABLE'; 
      rangeAddress?: string; 
      tableName?: string; 
      hasHeaders?: boolean 
    }
  | { 
      type: 'SORT_RANGE'; 
      rangeAddress?: string; 
      columnIndex: number; 
      ascending?: boolean 
    }
  | { 
      type: 'ADD_SHEET'; 
      sheetName: string; 
      activate?: boolean;
      tabColor?: string;
    }
  | {
      type: 'SET_SHEET_TAB_COLOR';
      sheetName?: string;
      color: string;
    }
  | {
      type: 'CREATE_TABLE_SECTION';
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
    }
  | {
      type: 'CREATE_MULTIPLE_TABLES';
      tables: Array<{
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
      }>;
    }
  | {
      type: 'CREATE_BUSINESS_PLAN';
      projectName?: string;
      projectType?: string;
      productTopic?: string;
      sheetCount?: number;
      currencySymbol?: string;
      sheetsToGenerate?: string[];
    }
  | { 
      type: 'AUTOFIT'; 
      rangeAddress?: string 
    }
  | { 
      type: 'CLEAN_DATA'; 
      rangeAddress?: string 
    }
  | { 
      type: 'EXECUTE_CODE'; 
      code: string; 
      description?: string 
    }
  | {
      type: 'MODIFY_CELLS';
      sheetName: string;
      changes: Array<{ cell: string; value?: any; formula?: string }>;
    }
  | {
      type: 'WRITE_TO_SHEET';
      sheetName: string;
      address: string;
      values: any[][];
      hasHeader?: boolean;
      formatAsCurrency?: boolean;
      currencySymbol?: string;
    };

export interface UserSubscription {
  plan: 'free' | 'pro_mensuel' | 'pack_credits';
  creditsRemaining: number;
  isActive: boolean;
  expiryDate?: string;
}

export interface BankReconciliationMatch {
  date: string;
  libelle: string;
  montant: number;
  matchedWith?: string;
  status: 'lettre' | 'ecart' | 'a_regulariser';
  note: string;
}
