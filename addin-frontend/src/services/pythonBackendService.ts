/**
 * Service de communication avec le moteur financier Python (FastAPI + openpyxl)
 * Permet l'injection instantanée de classeurs Excel natifs et le téléchargement direct.
 */

const PRIMARY_URL = (import.meta as any).env?.VITE_PYTHON_BACKEND_URL || "https://afriexcel.onrender.com";
const LOCAL_URL = "http://127.0.0.1:8000";

export interface PythonPlanResponse {
  success: boolean;
  project_name: string;
  sector: string;
  country: string;
  currency: string;
  summary: string;
  kpis: {
    total_investment: number;
    equity_promoter: number;
    bank_loan: number;
    revenue_year1: number;
    ebitda_year1: number;
    net_profit_year1: number;
    breakeven_revenue: number;
    breakeven_days: number;
    irr_pct: string;
    npv: number;
    payback_period: string;
    jobs_created: number;
  };
  sheets_count: number;
  filename: string;
  file_base64: string;
  template_id?: string;
  template_name?: string;
}

export class PythonBackendService {
  /**
   * Vérifie si le serveur Python local ou hébergé est en ligne (timeout tolérant 8s)
   */
  static async isAvailable(): Promise<boolean> {
    const urls = [PRIMARY_URL, LOCAL_URL];
    for (const url of urls) {
      try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 8000);
        const res = await fetch(`${url}/api/health`, {
          method: "GET",
          signal: controller.signal
        });
        clearTimeout(timeoutId);
        if (res.ok) return true;
      } catch (_) {}
    }
    return false;
  }

  /**
   * Génère le Business Plan complet à 20 feuilles via le moteur Python OHADA
   * Tente en priorité le serveur de production Render, avec repli automatique sur local
   */
  static async generateBusinessPlan(params: {
    prompt: string;
    country?: string;
    currency?: string;
    totalBudget?: number;
    horizonYears?: number;
    template_id?: string;
    isModification?: boolean;
    previousTopic?: string;
  }): Promise<PythonPlanResponse | null> {
    const candidateUrls = [PRIMARY_URL, LOCAL_URL];

    for (const url of candidateUrls) {
      try {
        const controller = new AbortController();
        // 60s pour permettre le démarrage à froid de Render et la modélisation complète
        const timeoutId = setTimeout(() => controller.abort(), 60000);

        const res = await fetch(`${url}/api/generate-business-plan`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            prompt: params.prompt,
            country: params.country || "Côte d'Ivoire",
            currency: params.currency || "FCFA",
            total_budget: params.totalBudget,
            horizon_years: params.horizonYears,
            template_id: params.template_id || "auto",
            is_modification: params.isModification || false,
            previous_topic: params.previousTopic
          }),
          signal: controller.signal
        });

        clearTimeout(timeoutId);

        if (res.ok) {
          const data: PythonPlanResponse = await res.json();
          if (data && data.file_base64) {
            console.log(`[Python Engine] Plan généré avec succès depuis ${url} (20 feuilles)`);
            return data;
          }
        }
      } catch (e) {
        console.warn(`[Python Engine] Échec sur ${url}, essai suivant :`, e);
      }
    }

    return null;
  }

  /**
   * Télécharge directement le fichier .xlsx généré sur l'ordinateur de l'utilisateur
   */
  static downloadXlsx(base64Data: string, filename: string = "BusinessPlan_OHADA.xlsx") {
    try {
      const byteCharacters = atob(base64Data);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      return true;
    } catch (e) {
      console.error("Erreur de téléchargement XLSX:", e);
      return false;
    }
  }

  /**
   * Insère ou remplace chirurgicalement les 20 feuilles du Business Plan dans le classeur Excel actif via Office.js
   */
  static async insertWorksheetsIntoActiveWorkbook(base64Data: string, isModification: boolean = false): Promise<boolean> {
    try {
      return await Excel.run(async (context) => {
        const workbook = context.workbook as any;
        if (typeof workbook.insertWorksheetsFromBase64 === "function") {
          const worksheets = context.workbook.worksheets;
          worksheets.load("items/name");
          await context.sync();

          // 1. Détection des anciennes feuilles de business plan (commençant par 01_, 02_... ou noms clés)
          const oldBpSheets = worksheets.items.filter(s =>
            /^\d{2}_/.test(s.name) ||
            s.name.includes("Page_de_Garde") ||
            s.name.includes("Tableau_de_Bord") ||
            s.name.includes("Plan_Financement")
          );

          let tempSheet: any = null;
          // Si des feuilles de BP existent déjà, on les supprime proprement pour éviter les doublons ("01_Page_de_Garde 1")
          if (oldBpSheets.length > 0) {
            // Excel impose qu'au moins une feuille reste visible en permanence dans le classeur
            if (oldBpSheets.length >= worksheets.items.length) {
              tempSheet = worksheets.add("AfriExcel_Sync");
              await context.sync();
            }

            for (const s of oldBpSheets) {
              s.delete();
            }
            await context.sync();
          }

          // 2. Insertion chirurgicale des 20 nouvelles feuilles natives calculées et cohérentes
          workbook.insertWorksheetsFromBase64(base64Data, {
            positionType: Excel.WorksheetPositionType.end
          });
          await context.sync();

          // 3. Suppression de la feuille tampon si créée
          if (tempSheet) {
            try {
              tempSheet.delete();
              await context.sync();
            } catch (_) {}
          }

          // 4. Nettoyage de la feuille vide par défaut 'Sheet1' ou 'Feuil1' si elle n'est plus utile
          try {
            worksheets.load("items/name");
            await context.sync();
            const emptyDefault = worksheets.items.find(s =>
              (s.name === "Sheet1" || s.name === "Feuil1") && worksheets.items.length > 1
            );
            if (emptyDefault) {
              emptyDefault.delete();
              await context.sync();
            }
          } catch (_) {}

          // 5. Activation immédiate de la Page de Garde pour l'utilisateur
          try {
            const firstSheet = context.workbook.worksheets.getItem("01_Page_de_Garde");
            firstSheet.activate();
            await context.sync();
          } catch (_) {}

          return true;
        } else {
          console.warn("insertWorksheetsFromBase64 n'est pas supporté par cette version d'Excel.");
          return false;
        }
      });
    } catch (err) {
      console.error("Erreur lors de l'insertion insertWorksheetsFromBase64:", err);
      return false;
    }
  }
}
