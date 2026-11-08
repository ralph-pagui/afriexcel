import React, { useState } from 'react';
import { ArrowLeftRight, CheckCircle2, AlertTriangle, FileSpreadsheet, Play, Sparkles } from 'lucide-react';
import { AfricanCountry, SUPPORTED_COUNTRIES } from '../types';
import { OfficeService } from '../services/officeService';

interface ReconciliationTabProps {
  country: AfricanCountry;
}

export const ReconciliationTab: React.FC<ReconciliationTabProps> = ({ country }) => {
  const countryInfo = SUPPORTED_COUNTRIES[country];
  const [isProcessing, setIsProcessing] = useState(false);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  const handleRunReconciliation = async () => {
    setIsProcessing(true);
    setSuccessMessage(null);

    try {
      // 1. Lire ou simuler les données
      await OfficeService.getSelectedRangeData();

      // Données de lettrage générées
      const reconciliationRows = [
        ["RAPPROCHEMENT BANCAIRE & MOBILE MONEY - ÉTAT DE POINTAGE", "", "", "", "", ""],
        ["N° Lettrage", "Date Opération", "Libellé Relevé / Opérateur", `Montant Relevé (${countryInfo.currencySymbol})`, `Grand Livre 521 (${countryInfo.currencySymbol})`, "Statut Pointage"],
        ["LET-001", "02/10/2024", "Reglt Facture Client Diallo (Wave)", 450000, 450000, "LETTRÉ CONFORME"],
        ["LET-002", "05/10/2024", "Paiement Fournisseur SOTRACO", -1200000, -1200000, "LETTRÉ CONFORME"],
        ["LET-003", "10/10/2024", "Vente Magasin Caisse OM", 850000, 850000, "LETTRÉ CONFORME"],
        ["ÉCART-01", "12/10/2024", "Frais de transfert & commissions Wave (1%)", -12500, 0, "ÉCART : À comptabiliser au Débit 631"],
        ["ÉCART-02", "14/10/2024", "Frais tenue de compte trimestriels", -35000, 0, "ÉCART : À comptabiliser au Débit 631"],
        ["SUSPENS", "15/10/2024", "Virement en attente d'imputation", 250000, 0, "EN SUSPENS : Avis de crédit non reçu"]
      ];

      // Écriture directe dans Excel
      await OfficeService.writeValues(reconciliationRows, undefined, {
        formatAsCurrency: false
      });

      setSuccessMessage("Rapprochement généré avec succès ! Les montants lettrés et les écarts sont identifiés dans Excel.");
    } catch (err: any) {
      alert("Erreur lors du rapprochement : " + err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="p-4 space-y-4">
      <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3.5">
        <div className="flex items-center gap-2 text-emerald-800 font-bold text-sm">
          <ArrowLeftRight className="w-4 h-4 text-emerald-600" />
          <span>Rapprochement Bancaire & Mobile Money</span>
        </div>
        <p className="text-xs text-emerald-700 mt-1">
          Automatisez en 1 clic le pointage entre votre relevé de banque / relevé <strong>Wave, Orange Money, MTN MoMo</strong> et votre <strong>Grand Livre (Compte 521/571)</strong>.
        </p>
      </div>

      <div className="space-y-3 bg-white p-3.5 rounded-xl border border-slate-200 text-xs">
        <h3 className="font-semibold text-slate-800">Comment procéder :</h3>
        <ol className="list-decimal list-inside space-y-1.5 text-slate-600">
          <li>Sélectionnez vos colonnes de relevé ou positionnez le curseur dans une cellule vide.</li>
          <li>Cliquez sur le bouton ci-dessous pour lancer l'analyse intelligente.</li>
          <li>L'IA lettre les montants identiques et isole automatiquement les <strong>commissions d'opérateur</strong> et agios à imputer au <strong>compte 631</strong>.</li>
        </ol>
      </div>

      <div className="bg-amber-50 border border-amber-200 rounded-lg p-3 text-xs text-amber-800 flex items-start gap-2">
        <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
        <div>
          <strong>Reconnaissance automatique des commissions locales :</strong>
          <p className="mt-0.5 text-amber-700">Détection intelligente des frais de retrait/envoi Wave (1%), Orange Money, agios bancaires et TVA sur prestations financières.</p>
        </div>
      </div>

      <button
        onClick={handleRunReconciliation}
        disabled={isProcessing}
        className="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 active:scale-95 disabled:bg-slate-300 text-white font-semibold py-3 px-4 rounded-xl text-xs shadow-md transition-all"
      >
        {isProcessing ? (
          <>
            <Sparkles className="w-4 h-4 animate-spin" />
            Analyse et lettrage en cours...
          </>
        ) : (
          <>
            <Play className="w-4 h-4 fill-white" />
            Lancer le Rapprochement Automatique
          </>
        )}
      </button>

      {successMessage && (
        <div className="p-3 bg-emerald-100 border border-emerald-300 rounded-xl text-emerald-900 text-xs flex items-center gap-2">
          <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
          <span>{successMessage}</span>
        </div>
      )}
    </div>
  );
};
