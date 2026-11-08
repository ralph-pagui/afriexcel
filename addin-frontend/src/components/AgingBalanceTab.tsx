import React, { useState } from 'react';
import { Clock, CheckCircle2, Play, AlertCircle, MessageSquare } from 'lucide-react';
import { AfricanCountry, SUPPORTED_COUNTRIES } from '../types';
import { OfficeService } from '../services/officeService';

interface AgingBalanceTabProps {
  country: AfricanCountry;
}

export const AgingBalanceTab: React.FC<AgingBalanceTabProps> = ({ country }) => {
  const countryInfo = SUPPORTED_COUNTRIES[country];
  const [isGenerating, setIsGenerating] = useState(false);
  const [statusMessage, setStatusMessage] = useState<string | null>(null);

  const handleGenerateAging = async () => {
    setIsGenerating(true);
    setStatusMessage(null);

    try {
      const rows = [
        ["BALANCE ÂGÉE CLIENTS & SUIVI DU RECOUVREMENT", "", "", "", "", "", ""],
        ["Code Client", "Raison Sociale Client", `Total Créance (${countryInfo.currencySymbol})`, "Non échu (0-30j)", "Retard 31-60j", "Retard 61-90j", "Contentieux (>90j)"],
        ["CL-001", "Société Générale de Bâtiment", 12500000, 12500000, 0, 0, 0],
        ["CL-002", "Quincaillerie de l'Avenir", 4800000, 0, 4800000, 0, 0],
        ["CL-003", "Pharmacie du Centre", 7200000, 0, 0, 7200000, 0],
        ["CL-004", "ETS Traoré & Frères", 9500000, 0, 0, 0, 9500000],
        ["TOTAL CRÉANCES", "=SUM(C3:C6)", "=SUM(D3:D6)", "=SUM(E3:E6)", "=SUM(F3:F6)", "=SUM(G3:G6)"]
      ];

      await OfficeService.writeValues(rows, undefined, { formatAsCurrency: false });
      setStatusMessage("Balance âgée dynamique générée avec formules et tranches d'antériorité !");
    } catch (e: any) {
      alert("Erreur Balance Âgée : " + e.message);
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="p-4 space-y-4">
      <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3.5">
        <div className="flex items-center gap-2 text-emerald-800 font-bold text-sm">
          <Clock className="w-4 h-4 text-emerald-600" />
          <span>Balance Âgée & Recouvrement Clients</span>
        </div>
        <p className="text-xs text-emerald-700 mt-1">
          Identifiez en quelques secondes vos créances en retard par tranches d'échéances et réduisez vos délais de paiement.
        </p>
      </div>

      <div className="bg-white p-3.5 rounded-xl border border-slate-200 space-y-2.5 text-xs">
        <h3 className="font-semibold text-slate-800">Tranches de classification automatique :</h3>
        <div className="space-y-1.5 text-slate-600">
          <div className="flex justify-between items-center p-1.5 bg-slate-50 rounded">
            <span>🟢 Créances courantes</span>
            <span className="font-medium text-slate-700">0 à 30 jours</span>
          </div>
          <div className="flex justify-between items-center p-1.5 bg-slate-50 rounded">
            <span>🟡 Retards modérés</span>
            <span className="font-medium text-slate-700">31 à 60 jours</span>
          </div>
          <div className="flex justify-between items-center p-1.5 bg-slate-50 rounded">
            <span>🟠 Retards critiques</span>
            <span className="font-medium text-slate-700">61 à 90 jours</span>
          </div>
          <div className="flex justify-between items-center p-1.5 bg-slate-50 rounded">
            <span>🔴 Créances douteuses (Prov. OHADA)</span>
            <span className="font-medium text-red-600">+90 jours</span>
          </div>
        </div>
      </div>

      <button
        onClick={handleGenerateAging}
        disabled={isGenerating}
        className="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 active:scale-95 disabled:bg-slate-300 text-white font-semibold py-3 px-4 rounded-xl text-xs shadow-md transition-all"
      >
        <Play className="w-4 h-4 fill-white" />
        <span>Générer la Balance Âgée dans Excel</span>
      </button>

      {statusMessage && (
        <div className="p-3 bg-emerald-100 border border-emerald-300 rounded-xl text-emerald-900 text-xs flex items-center gap-2">
          <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
          <span>{statusMessage}</span>
        </div>
      )}
    </div>
  );
};
