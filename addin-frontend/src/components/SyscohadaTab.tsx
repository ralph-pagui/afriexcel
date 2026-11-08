import React, { useState } from 'react';
import { BookOpen, CheckCircle2, FileText, ArrowRight, Table, Layers } from 'lucide-react';
import { AfricanCountry, SUPPORTED_COUNTRIES } from '../types';
import { OfficeService } from '../services/officeService';

interface SyscohadaTabProps {
  country: AfricanCountry;
}

export const SyscohadaTab: React.FC<SyscohadaTabProps> = ({ country }) => {
  const countryInfo = SUPPORTED_COUNTRIES[country];
  const [enterpriseName, setEnterpriseName] = useState('ENTREPRISE AFRICA SARL');
  const [exerciseYear, setExerciseYear] = useState('2024');
  const [isGenerating, setIsGenerating] = useState(false);
  const [statusText, setStatusText] = useState<string | null>(null);

  const handleGenerateBilan = async () => {
    setIsGenerating(true);
    setStatusText(null);

    try {
      await OfficeService.insertSyscohadaBilanSheet({
        exercice: exerciseYear,
        entreprise: enterpriseName,
        lignesActif: [
          { rubriques: "Frais d'établissement et de développement (211)", brut: 3000000, amort: 1000000, net: 2000000 },
          { rubriques: "Brevets, licences, logiciels (212)", brut: 4500000, amort: 1500000, net: 3000000 },
          { rubriques: "Terrains & Constructions (22)", brut: 45000000, amort: 9000000, net: 36000000 },
          { rubriques: "Matériel, mobilier et actifs biologiques (24)", brut: 18500000, amort: 6500000, net: 12000000 },
          { rubriques: "Matériel de transport (245)", brut: 12000000, amort: 4000000, net: 8000000 },
          { rubriques: "Titres de participation & Dépôts (26-27)", brut: 5000000, amort: 0, net: 5000000 },
          { rubriques: "Stocks de marchandises (31)", brut: 15400000, amort: 400000, net: 15000000 },
          { rubriques: "Créances Clients et comptes rattachés (41)", brut: 24600000, amort: 1600000, net: 23000000 },
          { rubriques: "Autres créances & État débiteur (44-46)", brut: 6200000, amort: 0, net: 6200000 },
          { rubriques: "Banques, chèques postaux et caisse (52-57)", brut: 14800000, amort: 0, net: 14800000 }
        ],
        lignesPassif: [
          { rubriques: "Capital social ou personnel (10)", net: 50000000 },
          { rubriques: "Réserves légales et statutaires (11)", net: 12000000 },
          { rubriques: "Report à nouveau (12)", net: 4500000 },
          { rubriques: "Résultat net de l'exercice (13)", net: 8500000 },
          { rubriques: "Emprunts et dettes financières (16)", net: 20000000 },
          { rubriques: "Dettes Fournisseurs et comptes rattachés (40)", net: 18000000 },
          { rubriques: "Dettes fiscales et sociales (42-43-44)", net: 9000000 },
          { rubriques: "Trésorerie Passif / Découverts bancaires (56)", net: 3000000 }
        ]
      });

      setStatusText("Feuille 'Bilan_OHADA' créée avec succès avec formules et mise en forme officielle !");
    } catch (err: any) {
      alert("Erreur Bilan : " + err.message);
    } finally {
      setIsGenerating(false);
    }
  };

  const handleGenerateResultat = async () => {
    setIsGenerating(true);
    setStatusText(null);

    try {
      const resultatRows = [
        [`COMPTE DE RÉSULTAT SYSCOHADA RÉVISÉ - ${enterpriseName.toUpperCase()}`, "", ""],
        [`Exercice clos au 31/12/${exerciseYear} (Montants en ${countryInfo.currencySymbol})`, "", ""],
        ["RUBRIQUES DU SYSTÈME NORMAL", "RÉFÉRENCE", `MONTANT NET (${countryInfo.currencySymbol})`],
        ["Ventes de marchandises (701)", "TA", 125000000],
        ["Achats de marchandises (601)", "RA", -78000000],
        ["Variation de stocks de marchandises (6031)", "RB", -2000000],
        ["MARGE COMMERCIALE (Somme TA à RB)", "XA", "=C4+C5+C6"],
        ["Ventes de produits fabriqués & Travaux (702-704)", "TB", 42000000],
        ["Consommations intermédiaires (602-608)", "RC", -15000000],
        ["Services extérieurs (61-63)", "RD", -18000000],
        ["VALEUR AJOUTÉE (XA + TB + RC + RD)", "XB", "=C7+C8+C9+C10"],
        ["Charges de personnel - Salaires et CNPS (66)", "RE", -22000000],
        ["EXCÉDENT BRUT D'EXPLOITATION (EBE)", "XC", "=C11+C12"],
        ["Dotations aux amortissements & provisions (68)", "RF", -6500000],
        ["Reprises de provisions (79)", "TD", 1200000],
        ["RÉSULTAT D'EXPLOITATION", "XD", "=C13+C14+C15"],
        ["Revenus financiers & Intérêts (77)", "TE", 850000],
        ["Frais financiers & Pertes de change (67)", "RG", -2100000],
        ["RÉSULTAT FINANCIER", "XE", "=C17+C18"],
        ["Résultat des Activités Ordinaires (RAO)", "XF", "=C16+C19"],
        ["Résultat Hors Activités Ordinaires - HAO (81-88)", "XG", 500000],
        ["Participation des travailleurs (87)", "RH", 0],
        ["Impôts sur le résultat / IS (89)", "RI", -3450000],
        ["RÉSULTAT NET DE L'EXERCICE (Bénéfice/Perte)", "XH", "=C20+C21+C22+C23"]
      ];

      await OfficeService.writeValues(resultatRows, undefined, { formatAsCurrency: false });
      setStatusText("Compte de Résultat SYSCOHADA généré avec calcul de la Marge, VA, EBE et Résultat Net !");
    } catch (err: any) {
      alert("Erreur Compte de Résultat : " + err.message);
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="p-4 space-y-4">
      <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3.5">
        <div className="flex items-center gap-2 text-emerald-800 font-bold text-sm">
          <BookOpen className="w-4 h-4 text-emerald-600" />
          <span>Normes Comptables SYSCOHADA Révisé</span>
        </div>
        <p className="text-xs text-emerald-700 mt-1">
          Génération automatisée des états financiers normalisés OHADA pour votre <strong>Liasse Fiscale (DSF)</strong>.
        </p>
      </div>

      {/* Paramètres de l'entreprise */}
      <div className="bg-white p-3 rounded-xl border border-slate-200 space-y-2.5 text-xs">
        <h3 className="font-semibold text-slate-800 flex items-center gap-1.5">
          <Layers className="w-3.5 h-3.5 text-emerald-600" />
          En-tête de la déclaration fiscale
        </h3>
        <div className="grid grid-cols-2 gap-2">
          <div>
            <label className="text-slate-500 block mb-1">Raison Sociale</label>
            <input
              type="text"
              value={enterpriseName}
              onChange={e => setEnterpriseName(e.target.value)}
              className="w-full border border-slate-300 rounded px-2 py-1 text-xs focus:ring-1 focus:ring-emerald-500 outline-none"
            />
          </div>
          <div>
            <label className="text-slate-500 block mb-1">Exercice</label>
            <input
              type="text"
              value={exerciseYear}
              onChange={e => setExerciseYear(e.target.value)}
              className="w-full border border-slate-300 rounded px-2 py-1 text-xs focus:ring-1 focus:ring-emerald-500 outline-none"
            />
          </div>
        </div>
      </div>

      {/* Boutons d'action */}
      <div className="space-y-2">
        <button
          onClick={handleGenerateBilan}
          disabled={isGenerating}
          className="w-full flex items-center justify-between bg-emerald-600 hover:bg-emerald-700 active:scale-95 disabled:bg-slate-300 text-white font-semibold py-2.5 px-3.5 rounded-xl text-xs shadow transition-all"
        >
          <div className="flex items-center gap-2">
            <Table className="w-4 h-4" />
            <span>Créer le Bilan SYSCOHADA (Actif & Passif)</span>
          </div>
          <ArrowRight className="w-3.5 h-3.5" />
        </button>

        <button
          onClick={handleGenerateResultat}
          disabled={isGenerating}
          className="w-full flex items-center justify-between bg-slate-800 hover:bg-slate-900 active:scale-95 disabled:bg-slate-300 text-white font-semibold py-2.5 px-3.5 rounded-xl text-xs shadow transition-all"
        >
          <div className="flex items-center gap-2">
            <FileText className="w-4 h-4" />
            <span>Créer le Compte de Résultat (Marge, VA, EBE)</span>
          </div>
          <ArrowRight className="w-3.5 h-3.5" />
        </button>
      </div>

      {statusText && (
        <div className="p-3 bg-emerald-100 border border-emerald-300 rounded-xl text-emerald-900 text-xs flex items-center gap-2">
          <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
          <span>{statusText}</span>
        </div>
      )}

      <div className="bg-slate-50 border border-slate-200 rounded-lg p-3 text-[11px] text-slate-500">
        <p><strong>Note de conformité OHADA :</strong> Les classes de comptes 1 à 8 sont mappées selon le Système Normal SYSCOHADA avec formules de bouclage automatique sans erreur circulaire.</p>
      </div>
    </div>
  );
};
