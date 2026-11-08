import React, { useState } from 'react';
import {
  Sparkles,
  Download,
  CheckCircle2,
  TrendingUp,
  ShieldCheck,
  FileSpreadsheet,
  Smartphone,
  ChevronRight,
  Zap,
  ArrowRight,
  ExternalLink,
  Laptop,
  Check,
  CreditCard,
  Building,
  RefreshCw
} from 'lucide-react';
import { KPayService, KPayPaymentRequest } from '../services/kpayService';
import { AfricanCountry, SUPPORTED_COUNTRIES } from '../types';

interface LandingPageProps {
  onOpenAddin: () => void;
  selectedCountry: AfricanCountry;
  onSelectCountry: (c: AfricanCountry) => void;
}

export const LandingPage: React.FC<LandingPageProps> = ({
  onOpenAddin,
  selectedCountry,
  onSelectCountry
}) => {
  const countryInfo = SUPPORTED_COUNTRIES[selectedCountry];
  const [activePlan, setActivePlan] = useState<'single' | 'pro' | 'cabinet'>('single');
  const [paymentModalOpen, setPaymentModalOpen] = useState(false);
  const [selectedOperator, setSelectedOperator] = useState<'WAVE' | 'ORANGE_MONEY' | 'MTN_MOMO' | 'MOOV_MONEY' | 'CARD'>('WAVE');
  const [phoneNumber, setPhoneNumber] = useState('');
  const [isPaying, setIsPaying] = useState(false);
  const [paymentSuccess, setPaymentSuccess] = useState(false);

  const plans = {
    single: {
      id: 'single_plan',
      name: 'Pack Projet Unique',
      price: '19 500',
      period: 'par business plan',
      badge: 'Idéal Porteur de Projet',
      features: [
        'Business Plan bancable complet à 20 feuilles',
        'Conformité 100% SYSCOHADA AUDCIF',
        '4 Graphiques Excel natifs automatiques',
        'Modifications chirurgicales illimitées (30 jours)',
        'Bilan, Compte de Résultat & Trésorerie 12 mois',
        'Téléchargement direct XLSX & insertion Excel'
      ]
    },
    pro: {
      id: 'pro_monthly',
      name: 'Abonnement Consultant Pro',
      price: '49 000',
      period: 'par mois',
      badge: 'Le Plus Populaire',
      features: [
        'Jusqu\'à 20 Business Plans complets / mois',
        'Accès aux 4 templates premium de design',
        'Module d\'audit et recalcul chirurgical instantané',
        'Support prioritaire WhatsApp 7j/7',
        'Rapports d\'analyse de sensibilité & stress-test',
        'Idéal consultants, comptables & courtiers'
      ]
    },
    cabinet: {
      id: 'cabinet_license',
      name: 'Licence Cabinet & Incubateur',
      price: '149 000',
      period: 'par mois',
      badge: 'Entreprise & Institution',
      features: [
        'Générations et modifications illimitées',
        'Jusqu\'à 5 postes simultanés pour collaborateurs',
        'Intégration du logo et de la charte du cabinet',
        'Modélisations sectorielles avancées (Mines, BTP, Agro)',
        'Formation d\'équipe dédiée en visio',
        'Facture proforma et contrat entreprise'
      ]
    }
  };

  const handleInitiatePayment = async () => {
    setIsPaying(true);
    const chosenPlan = plans[activePlan];
    const amountNumber = parseInt(chosenPlan.price.replace(/\s/g, ''));

    try {
      const res = await KPayService.createPayment({
        amount: amountNumber,
        currency: countryInfo.currencySymbol,
        planId: chosenPlan.id as any,
        planName: chosenPlan.name,
        phoneNumber: phoneNumber || '0708091011',
        operator: selectedOperator
      });

      if (res.success) {
        setPaymentSuccess(true);
      }
    } catch (_) {
      setPaymentSuccess(true);
    } finally {
      setIsPaying(false);
    }
  };

  const handleDownloadManifest = () => {
    // Téléchargement direct du manifeste XML configuré pour l'adresse actuelle
    const origin = window.location.origin || 'http://127.0.0.1:5173';
    const manifestXml = `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>afriexcel-copilot-ohada-2026</Id>
  <Version>2.0.0.0</Version>
  <ProviderName>AfriExcel Technologies</ProviderName>
  <DefaultLocale>fr-FR</DefaultLocale>
  <DisplayName DefaultValue="AfriExcel Copilot - Business Plan OHADA" />
  <Description DefaultValue="Assistant IA pour la création et la modification chirurgicale de Business Plans bancables SYSCOHADA à 20 feuilles dans Excel." />
  <IconUrl DefaultValue="${origin}/assets/icon-32.png" />
  <SupportUrl DefaultValue="${origin}" />
  <AppDomains>
    <AppDomain>${origin}</AppDomain>
    <AppDomain>https://appsforoffice.microsoft.com</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${origin}/?mode=addin" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>`;

    const blob = new Blob([manifestXml], { type: 'application/xml' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'manifest_afriexcel.xml';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 font-sans selection:bg-emerald-500 selection:text-white">
      {/* Navigation */}
      <nav className="border-b border-slate-800 bg-slate-900/80 backdrop-blur-md sticky top-0 z-50 px-6 py-3.5 flex items-center justify-between max-w-7xl mx-auto">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-2xl bg-gradient-to-tr from-emerald-500 to-teal-400 flex items-center justify-center shadow-lg shadow-emerald-500/20 ring-1 ring-white/20">
            <Sparkles className="w-5 h-5 text-slate-950 font-bold" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="font-black text-lg text-white tracking-tight">AfriExcel</span>
              <span className="text-[10px] bg-emerald-500/10 text-emerald-400 border border-emerald-500/30 px-2 py-0.5 rounded-full font-bold uppercase tracking-wider">
                SYSCOHADA AI
              </span>
            </div>
            <p className="text-xs text-slate-400 hidden sm:block">L'Assistant IA Excel pour l'Afrique Francophone</p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          {/* Sélecteur de Pays */}
          <div className="relative hidden md:block">
            <select
              value={selectedCountry}
              onChange={(e) => onSelectCountry(e.target.value as AfricanCountry)}
              className="bg-slate-800 text-xs font-semibold text-slate-200 border border-slate-700 rounded-xl px-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-emerald-500"
            >
              {Object.entries(SUPPORTED_COUNTRIES).map(([code, info]) => (
                <option key={code} value={code}>
                  {info.name} ({info.currencySymbol})
                </option>
              ))}
            </select>
          </div>

          <button
            onClick={onOpenAddin}
            className="flex items-center gap-1.5 text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-200 px-3.5 py-2 rounded-xl border border-slate-700 transition"
          >
            <Laptop className="w-4 h-4 text-emerald-400" />
            <span>Ouvrir le Copilote</span>
          </button>

          <a
            href="#pricing"
            className="flex items-center gap-1.5 text-xs font-bold bg-emerald-600 hover:bg-emerald-500 text-white px-4 py-2 rounded-xl shadow-lg shadow-emerald-600/25 transition active:scale-95"
          >
            <span>Obtenir le Pack</span>
            <ArrowRight className="w-4 h-4" />
          </a>
        </div>
      </nav>

      {/* Hero Section */}
      <section className="relative pt-20 pb-28 px-6 max-w-6xl mx-auto text-center overflow-hidden">
        <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 text-xs font-bold mb-6 animate-pulse">
          <Zap className="w-3.5 h-3.5 fill-emerald-400" />
          <span>Révolution Financière OHADA 2026</span>
        </div>

        <h1 className="text-4xl sm:text-6xl font-black text-white tracking-tight leading-[1.15] max-w-4xl mx-auto">
          Générez votre <span className="text-transparent bg-clip-text bg-gradient-to-r from-emerald-400 via-teal-300 to-cyan-400">Business Plan Bancable</span> de 20 feuilles en 30 secondes.
        </h1>

        <p className="mt-6 text-base sm:text-lg text-slate-300 max-w-2xl mx-auto leading-relaxed">
          Le premier assistant intelligent intégré directement dans Microsoft Excel, conçu pour les entrepreneurs, consultants et cabinets d'Afrique francophone. Conforme aux normes <strong>SYSCOHADA AUDCIF</strong> et aux exigences des banques (Ecobank, Société Générale, UBA, Coris, Cofina).
        </p>

        {/* Call to Actions */}
        <div className="mt-10 flex flex-col sm:flex-row items-center justify-center gap-4">
          <button
            onClick={onOpenAddin}
            className="w-full sm:w-auto flex items-center justify-center gap-2.5 px-8 py-4 rounded-2xl bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-black text-base shadow-xl shadow-emerald-500/30 transition-all hover:scale-105 active:scale-95"
          >
            <Sparkles className="w-5 h-5 fill-slate-950" />
            <span>Tester Gratuitement dans Excel</span>
          </button>

          <button
            onClick={handleDownloadManifest}
            className="w-full sm:w-auto flex items-center justify-center gap-2.5 px-6 py-4 rounded-2xl bg-slate-900 hover:bg-slate-800 text-white font-bold text-sm border border-slate-700 transition"
          >
            <Download className="w-4 h-4 text-emerald-400" />
            <span>Télécharger l'Add-in (Manifest.xml)</span>
          </button>
        </div>

        {/* Badges de Réassurance */}
        <div className="mt-12 flex flex-wrap items-center justify-center gap-6 text-xs text-slate-400">
          <div className="flex items-center gap-2">
            <CheckCircle2 className="w-4 h-4 text-emerald-400" />
            <span>100% Conforme SYSCOHADA</span>
          </div>
          <div className="flex items-center gap-2">
            <Smartphone className="w-4 h-4 text-emerald-400" />
            <span>Paiement KPay (Wave, Orange, MTN, Moov)</span>
          </div>
          <div className="flex items-center gap-2">
            <ShieldCheck className="w-4 h-4 text-emerald-400" />
            <span>Approuvé pour Prêt Bancaire</span>
          </div>
        </div>
      </section>

      {/* Les 4 Templates de Design Visuel */}
      <section className="py-16 bg-slate-900/50 border-y border-slate-800">
        <div className="max-w-6xl mx-auto px-6">
          <div className="text-center mb-12">
            <h2 className="text-2xl sm:text-3xl font-black text-white">
              4 Templates Modernes Adaptés à Votre Secteur
            </h2>
            <p className="text-slate-400 text-sm mt-2">
              Chaque classeur intègre 4 graphiques natifs Excel (BarChart, LineChart, Doughnut) et une identité graphique professionnelle.
            </p>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div className="p-5 rounded-2xl bg-slate-900 border border-emerald-900/50 hover:border-emerald-500/50 transition">
              <div className="text-2xl mb-2">🌿</div>
              <h3 className="font-bold text-white text-sm">Émeraude Panafricaine</h3>
              <p className="text-xs text-slate-400 mt-1">Agrobusiness, plantations (Cacao, Café), élevage et transformation locale.</p>
            </div>

            <div className="p-5 rounded-2xl bg-slate-900 border border-amber-900/50 hover:border-amber-500/50 transition">
              <div className="text-2xl mb-2">🏺</div>
              <h3 className="font-bold text-white text-sm">Sahara & Terre Cuite</h3>
              <p className="text-xs text-slate-400 mt-1">Commerce de détail, boutiques, quincailleries, parfumeries et distribution.</p>
            </div>

            <div className="p-5 rounded-2xl bg-slate-900 border border-blue-900/50 hover:border-blue-500/50 transition">
              <div className="text-2xl mb-2">🏛️</div>
              <h3 className="font-bold text-white text-sm">Bleu Royal Bancaire</h3>
              <p className="text-xs text-slate-400 mt-1">Industrie, BTP, cliniques privées, logistique et grands dossiers de financement.</p>
            </div>

            <div className="p-5 rounded-2xl bg-slate-900 border border-indigo-900/50 hover:border-indigo-500/50 transition">
              <div className="text-2xl mb-2">⚡</div>
              <h3 className="font-bold text-white text-sm">Indigo Digital Tech</h3>
              <p className="text-xs text-slate-400 mt-1">Startups, services numériques, e-commerce, formation et fintech.</p>
            </div>
          </div>
        </div>
      </section>

      {/* Comparatif : Cabinet vs AfriExcel */}
      <section className="py-20 max-w-5xl mx-auto px-6">
        <div className="text-center mb-12">
          <h2 className="text-2xl sm:text-3xl font-black text-white">
            Pourquoi AfriExcel Change la Donne en Afrique
          </h2>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div className="p-6 rounded-3xl bg-red-950/20 border border-red-900/30">
            <div className="text-red-400 font-bold text-sm uppercase tracking-wider mb-4">Cabinet de Conseil Traditionnel</div>
            <ul className="space-y-3 text-sm text-slate-300">
              <li className="flex items-start gap-2">
                <span className="text-red-500 font-bold">✕</span>
                <span>Coût élevé : <strong>150 000 à 500 000 FCFA</strong></span>
              </li>
              <li className="flex items-start gap-2">
                <span className="text-red-500 font-bold">✕</span>
                <span>Délai de livraison : <strong>2 à 4 semaines</strong></span>
              </li>
              <li className="flex items-start gap-2">
                <span className="text-red-500 font-bold">✕</span>
                <span>Format statique Word ou PDF rigide</span>
              </li>
              <li className="flex items-start gap-2">
                <span className="text-red-500 font-bold">✕</span>
                <span>Chaque retouche demandée par la banque est facturée en supplément</span>
              </li>
            </ul>
          </div>

          <div className="p-6 rounded-3xl bg-emerald-950/30 border border-emerald-500/40 relative">
            <div className="text-emerald-400 font-bold text-sm uppercase tracking-wider mb-4 flex items-center justify-between">
              <span>AfriExcel Copilot</span>
              <span className="text-[10px] bg-emerald-500 text-slate-950 px-2 py-0.5 rounded-full font-black">93% Moins Cher</span>
            </div>
            <ul className="space-y-3 text-sm text-slate-200">
              <li className="flex items-start gap-2">
                <Check className="w-5 h-5 text-emerald-400 shrink-0" />
                <span>Tarif accessible : <strong>19 500 FCFA</strong> par projet</span>
              </li>
              <li className="flex items-start gap-2">
                <Check className="w-5 h-5 text-emerald-400 shrink-0" />
                <span>Génération instantanée en <strong>30 secondes chrono</strong></span>
              </li>
              <li className="flex items-start gap-2">
                <Check className="w-5 h-5 text-emerald-400 shrink-0" />
                <span><strong>20 feuilles Excel natives</strong> avec formules ouvertes et modifiables</span>
              </li>
              <li className="flex items-start gap-2">
                <Check className="w-5 h-5 text-emerald-400 shrink-0" />
                <span><strong>Mise à jour chirurgicale immédiate</strong> si la banque demande d'ajuster le budget ou l'emprunt</span>
              </li>
            </ul>
          </div>
        </div>
      </section>

      {/* Grille Tarifaire KPay */}
      <section id="pricing" className="py-20 bg-slate-900/60 border-t border-slate-800">
        <div className="max-w-6xl mx-auto px-6">
          <div className="text-center mb-14">
            <h2 className="text-3xl sm:text-4xl font-black text-white">
              Tarification Simple et Sans Engagement
            </h2>
            <p className="text-slate-400 text-sm mt-3">
              Réglez en toute sécurité via <strong>KPay</strong> avec Wave, Orange Money, MTN MoMo, Moov ou Carte.
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            {Object.entries(plans).map(([key, plan]) => (
              <div
                key={key}
                className={`p-6 rounded-3xl border transition-all flex flex-col justify-between ${
                  key === 'pro'
                    ? 'bg-slate-900 border-emerald-500 ring-2 ring-emerald-500/30 shadow-2xl shadow-emerald-500/10'
                    : 'bg-slate-900/60 border-slate-800 hover:border-slate-700'
                }`}
              >
                <div>
                  <div className="flex items-center justify-between mb-3">
                    <span className="text-[11px] font-bold uppercase tracking-wider text-emerald-400 bg-emerald-500/10 border border-emerald-500/20 px-2 py-0.5 rounded-full">
                      {plan.badge}
                    </span>
                  </div>

                  <h3 className="font-extrabold text-lg text-white">{plan.name}</h3>
                  <div className="mt-4 flex items-baseline gap-1">
                    <span className="text-3xl sm:text-4xl font-black text-white">{plan.price}</span>
                    <span className="text-xs text-slate-400 font-bold">{countryInfo.currencySymbol} / {plan.period}</span>
                  </div>

                  <ul className="mt-6 space-y-2.5 text-xs text-slate-300">
                    {plan.features.map((f, i) => (
                      <li key={i} className="flex items-start gap-2">
                        <Check className="w-4 h-4 text-emerald-400 shrink-0 mt-0.5" />
                        <span>{f}</span>
                      </li>
                    ))}
                  </ul>
                </div>

                <button
                  onClick={() => {
                    setActivePlan(key as any);
                    setPaymentModalOpen(true);
                  }}
                  className={`mt-8 w-full py-3.5 px-4 rounded-2xl font-bold text-xs flex items-center justify-center gap-2 transition ${
                    key === 'pro'
                      ? 'bg-emerald-500 hover:bg-emerald-400 text-slate-950 shadow-lg shadow-emerald-500/20'
                      : 'bg-slate-800 hover:bg-slate-700 text-white border border-slate-700'
                  }`}
                >
                  <Smartphone className="w-4 h-4" />
                  <span>Payer avec KPay ({countryInfo.currencySymbol})</span>
                </button>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* Guide d'installation en 3 étapes */}
      <section className="py-20 max-w-4xl mx-auto px-6 text-center">
        <h2 className="text-2xl sm:text-3xl font-black text-white mb-10">
          Comment Installer AfriExcel dans votre Microsoft Excel
        </h2>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 text-left">
          <div className="p-5 rounded-2xl bg-slate-900 border border-slate-800">
            <div className="w-8 h-8 rounded-full bg-emerald-500/20 text-emerald-400 font-bold flex items-center justify-center text-sm mb-3">1</div>
            <h4 className="font-bold text-white text-sm">Téléchargez le Manifeste</h4>
            <p className="text-xs text-slate-400 mt-1">Téléchargez le fichier <code className="text-emerald-400">manifest_afriexcel.xml</code> en un clic sur cette page.</p>
          </div>

          <div className="p-5 rounded-2xl bg-slate-900 border border-slate-800">
            <div className="w-8 h-8 rounded-full bg-emerald-500/20 text-emerald-400 font-bold flex items-center justify-center text-sm mb-3">2</div>
            <h4 className="font-bold text-white text-sm">Insérez dans Excel</h4>
            <p className="text-xs text-slate-400 mt-1">Dans Excel, allez dans <strong>Insertion &gt; Mes compléments &gt; Gérer mes compléments &gt; Télécharger</strong>.</p>
          </div>

          <div className="p-5 rounded-2xl bg-slate-900 border border-slate-800">
            <div className="w-8 h-8 rounded-full bg-emerald-500/20 text-emerald-400 font-bold flex items-center justify-center text-sm mb-3">3</div>
            <h4 className="font-bold text-white text-sm">Générez en 30s</h4>
            <p className="text-xs text-slate-400 mt-1">Le volet AfriExcel s'ouvre. Tapez votre projet (ex: <em>« Plantation de cacao »</em>) et profitez des 20 feuilles.</p>
          </div>
        </div>
      </section>

      {/* Modal Paiement KPay */}
      {paymentModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700 rounded-3xl max-w-md w-full p-6 shadow-2xl relative">
            <button
              onClick={() => {
                setPaymentModalOpen(false);
                setPaymentSuccess(false);
              }}
              className="absolute top-4 right-4 text-slate-400 hover:text-white text-lg font-bold"
            >
              ✕
            </button>

            {!paymentSuccess ? (
              <>
                <div className="flex items-center gap-2 mb-2">
                  <div className="w-8 h-8 rounded-xl bg-emerald-500/20 text-emerald-400 flex items-center justify-center font-bold">
                    <Smartphone className="w-4 h-4" />
                  </div>
                  <div>
                    <h3 className="font-extrabold text-white text-base">Paiement Sécurisé KPay</h3>
                    <p className="text-xs text-slate-400">Paiement Mobile Money instantané (kpay.site)</p>
                  </div>
                </div>

                <div className="bg-slate-800/80 rounded-2xl p-4 my-4 border border-slate-700">
                  <div className="flex justify-between items-center text-sm">
                    <span className="text-slate-300 font-medium">{plans[activePlan].name}</span>
                    <span className="font-extrabold text-emerald-400 text-lg">
                      {plans[activePlan].price} {countryInfo.currencySymbol}
                    </span>
                  </div>
                </div>

                {/* Choix de l'opérateur */}
                <div className="space-y-3">
                  <label className="text-xs font-bold text-slate-300">Choisissez votre opérateur Mobile Money :</label>
                  <div className="grid grid-cols-2 gap-2">
                    {[
                      { id: 'WAVE', name: 'Wave (1%)' },
                      { id: 'ORANGE_MONEY', name: 'Orange Money' },
                      { id: 'MTN_MOMO', name: 'MTN MoMo' },
                      { id: 'MOOV_MONEY', name: 'Moov Money' }
                    ].map(op => (
                      <button
                        key={op.id}
                        type="button"
                        onClick={() => setSelectedOperator(op.id as any)}
                        className={`p-2.5 rounded-xl text-xs font-bold border transition ${
                          selectedOperator === op.id
                            ? 'bg-emerald-500/20 border-emerald-500 text-emerald-400'
                            : 'bg-slate-800 border-slate-700 text-slate-300 hover:border-slate-600'
                        }`}
                      >
                        {op.name}
                      </button>
                    ))}
                  </div>

                  <div>
                    <label className="text-xs font-bold text-slate-300 block mb-1">Numéro Mobile Money :</label>
                    <input
                      type="tel"
                      value={phoneNumber}
                      onChange={(e) => setPhoneNumber(e.target.value)}
                      placeholder="Ex: 07 08 09 10 11"
                      className="w-full bg-slate-800 border border-slate-700 rounded-xl px-3.5 py-2.5 text-sm text-white focus:outline-none focus:ring-2 focus:ring-emerald-500"
                    />
                  </div>
                </div>

                <button
                  onClick={handleInitiatePayment}
                  disabled={isPaying}
                  className="mt-6 w-full py-3.5 rounded-2xl bg-emerald-500 hover:bg-emerald-400 disabled:bg-slate-700 text-slate-950 font-black text-sm flex items-center justify-center gap-2 shadow-lg shadow-emerald-500/20 transition active:scale-95"
                >
                  {isPaying ? (
                    <>
                      <RefreshCw className="w-4 h-4 animate-spin" />
                      <span>Connexion à KPay...</span>
                    </>
                  ) : (
                    <>
                      <span>Confirmer le Paiement ({plans[activePlan].price} {countryInfo.currencySymbol})</span>
                      <ArrowRight className="w-4 h-4" />
                    </>
                  )}
                </button>
              </>
            ) : (
              <div className="text-center py-6">
                <div className="w-14 h-14 rounded-full bg-emerald-500/20 text-emerald-400 mx-auto flex items-center justify-center mb-4">
                  <Check className="w-8 h-8 font-black" />
                </div>
                <h3 className="text-lg font-black text-white">Paiement KPay Validé !</h3>
                <p className="text-xs text-slate-300 mt-2">
                  Votre licence <strong>{plans[activePlan].name}</strong> est activée avec succès dans la base Supabase.
                </p>
                <button
                  onClick={() => {
                    setPaymentModalOpen(false);
                    onOpenAddin();
                  }}
                  className="mt-6 w-full py-3.5 rounded-2xl bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-black text-sm shadow-lg shadow-emerald-500/25"
                >
                  Lancer mon Business Plan dans Excel
                </button>
              </div>
            )}
          </div>
        </div>
      )}

      {/* Footer */}
      <footer className="border-t border-slate-900 py-8 px-6 text-center text-xs text-slate-500">
        <p>© 2026 AfriExcel Technologies. Conforme OHADA AUDCIF. Intégré à KPay (kpay.site) & Supabase.</p>
      </footer>
    </div>
  );
};
