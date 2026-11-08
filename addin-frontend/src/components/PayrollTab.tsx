import React, { useState } from 'react';
import { Users, CheckCircle2, Play, DollarSign, Calculator } from 'lucide-react';
import { AfricanCountry, SUPPORTED_COUNTRIES } from '../types';
import { OfficeService } from '../services/officeService';

interface PayrollTabProps {
  country: AfricanCountry;
}

export const PayrollTab: React.FC<PayrollTabProps> = ({ country }) => {
  const countryInfo = SUPPORTED_COUNTRIES[country];
  const [isGenerating, setIsGenerating] = useState(false);
  const [statusMessage, setStatusMessage] = useState<string | null>(null);

  const handleGeneratePayroll = async () => {
    setIsGenerating(true);
    setStatusMessage(null);

    try {
      const sampleEmployees = [
        { nom: "Mamadou DIOP", salaireBase: 450000, primes: 50000, partsFiscales: 2 },
        { nom: "Awa KONE", salaireBase: 650000, primes: 75000, partsFiscales: 3 },
        { nom: "Christian TCHOUA", salaireBase: 350000, primes: 25000, partsFiscales: 1 },
        { nom: "Béatrice N'GUESSAN", salaireBase: 800000, primes: 120000, partsFiscales: 4 }
      ];

      await OfficeService.insertPayrollSheet(countryInfo.name, sampleEmployees);
      setStatusMessage(`Livre de paie pour ${countryInfo.name} créé avec succès avec formules ${countryInfo.cnssLabel} et ${countryInfo.taxLabel} !`);
    } catch (e: any) {
      alert("Erreur Paie : " + e.message);
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="p-4 space-y-4">
      <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3.5">
        <div className="flex items-center gap-2 text-emerald-800 font-bold text-sm">
          <Users className="w-4 h-4 text-emerald-600" />
          <span>Livre de Paie & Fiscalité Sociale ({countryInfo.name})</span>
        </div>
        <p className="text-xs text-emerald-700 mt-1">
          Générez instantanément votre grille de salaires conforme au Code du Travail et à la fiscalité de <strong>{countryInfo.name}</strong>.
        </p>
      </div>

      <div className="bg-white p-3.5 rounded-xl border border-slate-200 space-y-3 text-xs">
        <h3 className="font-semibold text-slate-800 flex items-center gap-1.5">
          <Calculator className="w-3.5 h-3.5 text-emerald-600" />
          Barèmes appliqués automatiquement :
        </h3>
        <div className="grid grid-cols-2 gap-2">
          <div className="p-2.5 bg-slate-50 rounded-lg border border-slate-100">
            <span className="text-[11px] text-slate-400 block">Cotisation Sociale</span>
            <span className="font-bold text-slate-700 text-xs">{countryInfo.cnssLabel}</span>
          </div>
          <div className="p-2.5 bg-slate-50 rounded-lg border border-slate-100">
            <span className="text-[11px] text-slate-400 block">Impôt sur Salaires</span>
            <span className="font-bold text-slate-700 text-xs">{countryInfo.taxLabel}</span>
          </div>
        </div>

        <div className="p-2 bg-emerald-50 rounded-lg border border-emerald-100 text-emerald-800 text-[11px]">
          💡 Toutes les retenues (Brut, Charges, Net à Payer) sont écrites sous forme de <strong>formules Excel dynamiques</strong> réutilisables chaque mois.
        </div>
      </div>

      <button
        onClick={handleGeneratePayroll}
        disabled={isGenerating}
        className="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 active:scale-95 disabled:bg-slate-300 text-white font-semibold py-3 px-4 rounded-xl text-xs shadow-md transition-all"
      >
        <Play className="w-4 h-4 fill-white" />
        <span>Générer le Livre de Paie Complet</span>
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
