import React, { useState, useEffect } from 'react';
import { Sparkles, ChevronDown, Globe, Laptop, ShieldCheck } from 'lucide-react';
import { AfricanCountry, SUPPORTED_COUNTRIES } from './types';
import { ChatTab } from './components/ChatTab';
import { LandingPage } from './components/LandingPage';
import { KPayService } from './services/kpayService';

export const App: React.FC = () => {
  const [selectedCountry, setSelectedCountry] = useState<AfricanCountry>('CI');
  const [activePlan, setActivePlan] = useState<any>(null);

  // Détection du mode initial :
  // Si le paramètre URL contient ?mode=addin ou si Office.js est détecté, on affiche le copilote
  // Sinon on affiche le site vitrine dédié
  const [viewMode, setViewMode] = useState<'landing' | 'addin'>(() => {
    if (typeof window !== 'undefined') {
      const params = new URLSearchParams(window.location.search);
      if (params.get('mode') === 'addin' || params.get('view') === 'addin') {
        return 'addin';
      }
      if ((window as any).Office?.context?.document) {
        return 'addin';
      }
    }
    return 'landing';
  });

  useEffect(() => {
    // Vérifier si un plan KPay a été activé
    const plan = KPayService.getActivePlan();
    if (plan) {
      setActivePlan(plan);
    }
  }, [viewMode]);

  const currentCountry = SUPPORTED_COUNTRIES[selectedCountry];

  // 1. Vue Site Web Dédié (Vitrine d'acquisition, Tarifs KPay, Installation 2 clics)
  if (viewMode === 'landing') {
    return (
      <LandingPage
        onOpenAddin={() => setViewMode('addin')}
        selectedCountry={selectedCountry}
        onSelectCountry={setSelectedCountry}
      />
    );
  }

  // 2. Vue Volet Copilote Excel (Taskpane)
  return (
    <div className="flex flex-col h-screen bg-slate-100 text-slate-800 font-sans select-none overflow-hidden">
      {/* En-tête exécutif : Copilote Business Plan Bancable OHADA */}
      <header className="bg-slate-900 text-white px-3 py-2 border-b border-slate-800 shrink-0 shadow-md">
        <div className="flex items-center justify-between">
          {/* Logo & Titre de l'assistant spécialisé */}
          <div className="flex items-center gap-2">
            <div className="w-7 h-7 bg-gradient-to-tr from-emerald-600 to-teal-500 rounded-xl flex items-center justify-center font-bold text-white shadow-sm ring-1 ring-white/10">
              <Sparkles className="w-3.5 h-3.5 text-emerald-100" />
            </div>
            <div>
              <div className="flex items-center gap-1.5">
                <h1 className="text-xs font-black tracking-tight text-white">
                  AfriExcel Copilot
                </h1>
                <span className="text-[9px] bg-emerald-950 text-emerald-400 border border-emerald-800/80 px-1 py-0.2 rounded font-semibold uppercase tracking-wider">
                  {activePlan ? 'PRO ACTIF' : 'OHADA'}
                </span>
              </div>
              <div className="text-[9px] text-slate-400 flex items-center gap-1">
                <span className="inline-block w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse"></span>
                <span>20 Feuilles SYSCOHADA</span>
              </div>
            </div>
          </div>

          {/* Boutons d'action rapides : Retour au site web & Sélecteur de pays */}
          <div className="flex items-center gap-1.5">
            <button
              onClick={() => setViewMode('landing')}
              title="Consulter le site web officiel et les offres KPay"
              className="flex items-center gap-1 text-[10px] font-bold bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white px-2 py-1 rounded-lg border border-slate-700 transition"
            >
              <Globe className="w-3 h-3 text-emerald-400" />
              <span className="hidden sm:inline">Site & Offres</span>
            </button>

            {/* Sélecteur de Pays & Zone Monétaire */}
            <div className="relative">
              <select
                value={selectedCountry}
                onChange={e => setSelectedCountry(e.target.value as AfricanCountry)}
                className="appearance-none bg-slate-800 hover:bg-slate-700 border border-slate-700 text-white text-[10px] font-semibold rounded-lg pl-1.5 pr-5 py-1 cursor-pointer focus:outline-none focus:ring-1 focus:ring-emerald-500"
                title="Changer de pays d'application financière"
              >
                {Object.values(SUPPORTED_COUNTRIES).map(c => (
                  <option key={c.code} value={c.code}>
                    {c.flag} {c.currencySymbol}
                  </option>
                ))}
              </select>
              <ChevronDown className="w-2.5 h-2.5 text-slate-400 absolute right-1 top-2 pointer-events-none" />
            </div>
          </div>
        </div>
      </header>

      {/* Interface du Copilote IA Bancable */}
      <main className="flex-1 overflow-hidden">
        <ChatTab country={selectedCountry} />
      </main>
    </div>
  );
};

export default App;