import React, { useState } from 'react';
import { Eraser, CheckCircle2, Sparkles, Filter, Wand2 } from 'lucide-react';
import { OfficeService } from '../services/officeService';

export const DataCleanTab: React.FC = () => {
  const [isCleaning, setIsCleaning] = useState(false);
  const [result, setResult] = useState<{ cleanedRows: number; cleanedCells: number } | null>(null);

  const handleClean = async () => {
    setIsCleaning(true);
    setResult(null);

    try {
      const res = await OfficeService.cleanSelectedData();
      setResult(res);
    } catch (e: any) {
      alert("Erreur Nettoyage : " + e.message);
    } finally {
      setIsCleaning(false);
    }
  };

  return (
    <div className="p-4 space-y-4">
      <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3.5">
        <div className="flex items-center gap-2 text-emerald-800 font-bold text-sm">
          <Wand2 className="w-4 h-4 text-emerald-600" />
          <span>Nettoyage Express d'Exports Comptables</span>
        </div>
        <p className="text-xs text-emerald-700 mt-1">
          Réparez instantanément les exports <strong>Sage Saari, Ciel, Gescom</strong> ou extraits bancaires corrompus.
        </p>
      </div>

      <div className="space-y-2 bg-white p-3.5 rounded-xl border border-slate-200 text-xs">
        <h3 className="font-semibold text-slate-800">Corrections automatiques appliquées :</h3>
        <ul className="space-y-2 text-slate-600">
          <li className="flex items-start gap-2">
            <span className="text-emerald-600 font-bold">✓</span>
            <div>
              <strong>Suppression des espaces insécables :</strong>
              <p className="text-[11px] text-slate-500">Transforme les montants comme <code className="bg-slate-100 px-1">"1 500 000 "</code> (vus comme texte) en vrais nombres calculables.</p>
            </div>
          </li>
          <li className="flex items-start gap-2">
            <span className="text-emerald-600 font-bold">✓</span>
            <div>
              <strong>Normalisation des dates :</strong>
              <p className="text-[11px] text-slate-500">Conversion des formats inversés vers le standard francophone <code className="bg-slate-100 px-1">JJ/MM/AAAA</code>.</p>
            </div>
          </li>
          <li className="flex items-start gap-2">
            <span className="text-emerald-600 font-bold">✓</span>
            <div>
              <strong>Élimination des fusions parasites :</strong>
              <p className="text-[11px] text-slate-500">Défusionne les en-têtes qui bloquent les filtres et les formules de recherche.</p>
            </div>
          </li>
        </ul>
      </div>

      <button
        onClick={handleClean}
        disabled={isCleaning}
        className="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 active:scale-95 disabled:bg-slate-300 text-white font-semibold py-3 px-4 rounded-xl text-xs shadow-md transition-all"
      >
        {isCleaning ? (
          <>
            <Sparkles className="w-4 h-4 animate-spin" />
            Nettoyage en cours...
          </>
        ) : (
          <>
            <Eraser className="w-4 h-4" />
            Nettoyer la Sélection Active
          </>
        )}
      </button>

      {result && (
        <div className="p-3 bg-emerald-100 border border-emerald-300 rounded-xl text-emerald-900 text-xs flex items-center gap-2">
          <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
          <span>Nettoyage terminé : {result.cleanedCells} cellules corrigées sur {result.cleanedRows} lignes !</span>
        </div>
      )}
    </div>
  );
};
