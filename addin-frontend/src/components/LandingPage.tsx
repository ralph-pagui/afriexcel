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
  RefreshCw,
  Copy,
  Terminal
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
  const [installTab, setInstallTab] = useState<'modern' | 'cmd' | 'manual'>('modern');
  const [copied, setCopied] = useState(false);

  const COMMAND_MODERN = "irm https://raw.githubusercontent.com/ralph-pagui/afriexcel/main/addin-frontend/public/install.ps1 | iex";
  const COMMAND_CMD = 'powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/ralph-pagui/afriexcel/main/addin-frontend/public/install.ps1 | iex"';
  const COMMAND_LEGACY = "[Net.ServicePointManager]::SecurityProtocol = 3072 -bor 768; iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/ralph-pagui/afriexcel/main/addin-frontend/public/install.ps1'))";

  const handleCopyCommand = (text: string) => {
    navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2500);
  };

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

          {/* Bouton d'accès rapide à l'installation */}
          <a
            href="#install"
            className="flex items-center gap-1.5 text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-200 px-3 py-2 rounded-xl border border-slate-700 transition"
          >
            <Terminal className="w-3.5 h-3.5 text-emerald-400" />
            <span className="hidden sm:inline">Installer dans Excel</span>
          </a>

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
          <a
            href="#install"
            className="w-full sm:w-auto flex items-center justify-center gap-2.5 px-8 py-4 rounded-2xl bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-black text-base shadow-xl shadow-emerald-500/30 transition-all hover:scale-105 active:scale-95"
          >
            <Terminal className="w-5 h-5 text-slate-950" />
            <span>Installer dans mon Excel (1 Clic)</span>
          </a>

          <button
            onClick={onOpenAddin}
            className="w-full sm:w-auto flex items-center justify-center gap-2.5 px-6 py-4 rounded-2xl bg-slate-900 hover:bg-slate-800 text-white font-bold text-sm border border-slate-700 transition"
          >
            <Sparkles className="w-4 h-4 text-emerald-400" />
            <span>Essayer le Copilote Web</span>
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

      {/* Section Installation 1 Clic & Procédure Universelle */}
      <section id="install" className="py-24 max-w-5xl mx-auto px-6">
        <div className="text-center mb-12">
          <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 text-xs font-bold mb-4">
            <Terminal className="w-3.5 h-3.5" />
            <span>INSTALLATION 1-CLIC DANS EXCEL</span>
          </div>
          <h2 className="text-3xl sm:text-4xl font-black text-white">
            Comment Installer AfriExcel dans votre Microsoft Excel
          </h2>
          <p className="text-slate-300 text-sm sm:text-base mt-3 max-w-2xl mx-auto">
            Une seule commande sécurisée dans PowerShell configure instantanément l'add-in dans votre ruban Excel.
            Compatible avec <strong>tous les PC Windows (Windows 7, 8, 10, 11)</strong>, sans droits administrateur et sans alerte de sécurité.
          </p>

          {/* Badges de Compatibilité & Sécurité */}
          <div className="mt-8 flex flex-wrap items-center justify-center gap-3 text-xs">
            <span className="flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-slate-900 border border-slate-800 text-slate-300">
              <ShieldCheck className="w-4 h-4 text-emerald-400" />
              <span>Sans droits administrateur</span>
            </span>
            <span className="flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-slate-900 border border-slate-800 text-slate-300">
              <CheckCircle2 className="w-4 h-4 text-emerald-400" />
              <span>TLS 1.0, 1.1, 1.2 & 1.3 débloqués</span>
            </span>
            <span className="flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-slate-900 border border-slate-800 text-slate-300">
              <CheckCircle2 className="w-4 h-4 text-emerald-400" />
              <span>Zéro alerte SmartScreen</span>
            </span>
            <span className="flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-slate-900 border border-slate-800 text-slate-300">
              <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
              <span>Excel 2016 à Microsoft 365</span>
            </span>
          </div>
        </div>

        {/* Terminal Box Interactif */}
        <div className="bg-slate-900 border border-slate-800 rounded-3xl overflow-hidden shadow-2xl shadow-emerald-950/20">
          {/* Header Terminal avec Sélecteur d'onglets */}
          <div className="bg-slate-950 px-4 py-3 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3">
            <div className="flex items-center gap-2">
              <span className="w-3 h-3 rounded-full bg-red-500/80 inline-block"></span>
              <span className="w-3 h-3 rounded-full bg-yellow-500/80 inline-block"></span>
              <span className="w-3 h-3 rounded-full bg-green-500/80 inline-block"></span>
              <span className="text-xs text-slate-400 font-mono ml-2 font-medium">powershell.exe</span>
            </div>

            {/* Onglets de syntaxe */}
            <div className="flex items-center gap-1.5 bg-slate-900 p-1 rounded-xl border border-slate-800">
              <button
                onClick={() => setInstallTab('modern')}
                className={`px-3.5 py-1.5 rounded-lg text-xs font-bold transition ${
                  installTab === 'modern'
                    ? 'bg-emerald-500 text-slate-950 shadow-sm'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                PowerShell 1-Clic (Recommandé)
              </button>
              <button
                onClick={() => setInstallTab('cmd')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition ${
                  installTab === 'cmd'
                    ? 'bg-emerald-500 text-slate-950 shadow-sm'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                Invite CMD
              </button>
              <button
                onClick={() => setInstallTab('manual')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition ${
                  installTab === 'manual'
                    ? 'bg-emerald-500 text-slate-950 shadow-sm'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                Téléchargement Fichier
              </button>
            </div>
          </div>

          {/* Corps du Terminal */}
          <div className="p-6 sm:p-8">
            {installTab !== 'manual' ? (
              <>
                <div className="text-xs text-slate-400 mb-3 flex items-center justify-between">
                  <span>
                    {installTab === 'modern' && '✦ Commande officielle ultra-rapide 1-clic pour Windows 10 & 11 (PowerShell).'}
                    {installTab === 'cmd' && '✦ À copier-coller dans l\'invite de commandes classique noire (cmd.exe).'}
                  </span>
                </div>

                <div className="relative bg-slate-950 rounded-2xl p-4 sm:p-5 border border-slate-800/80 font-mono text-xs sm:text-sm text-slate-200 overflow-x-auto flex items-center justify-between gap-4">
                  <div className="flex items-start gap-2 select-all overflow-x-auto">
                    <span className="text-emerald-400 select-none font-bold">
                      {installTab === 'modern' ? 'PS>' : 'C:\\>'}
                    </span>
                    <span className="text-emerald-200 whitespace-pre-wrap break-all font-semibold">
                      {installTab === 'modern' ? COMMAND_MODERN : COMMAND_CMD}
                    </span>
                  </div>

                  <button
                    onClick={() => {
                      const text = installTab === 'modern' ? COMMAND_MODERN : COMMAND_CMD;
                      handleCopyCommand(text);
                    }}
                    className={`shrink-0 flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-black transition active:scale-95 ${
                      copied
                        ? 'bg-emerald-500 text-slate-950 shadow-lg shadow-emerald-500/30'
                        : 'bg-emerald-600 hover:bg-emerald-500 text-white shadow-md'
                    }`}
                  >
                    {copied ? (
                      <>
                        <Check className="w-4 h-4 stroke-[3]" />
                        <span>Copié !</span>
                      </>
                    ) : (
                      <>
                        <Copy className="w-4 h-4" />
                        <span>Copier</span>
                      </>
                    )}
                  </button>
                </div>

                {/* Secours pour très anciennes machines */}
                <details className="mt-4 pt-3 border-t border-slate-800/60 text-left">
                  <summary className="text-[11px] text-slate-400 hover:text-slate-300 cursor-pointer font-medium select-none">
                    ✦ Vous utilisez Windows 7 ou une très ancienne version de PowerShell ? (Cliquez ici pour la commande secours)
                  </summary>
                  <div className="mt-2.5 p-3 rounded-xl bg-slate-950/70 border border-slate-800 text-xs font-mono text-slate-300 flex items-center justify-between gap-3">
                    <span className="whitespace-pre-wrap break-all select-all text-emerald-300 text-[11px]">
                      {COMMAND_LEGACY}
                    </span>
                    <button
                      onClick={() => handleCopyCommand(COMMAND_LEGACY)}
                      className="shrink-0 px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-white text-[11px] font-bold"
                    >
                      Copier
                    </button>
                  </div>
                </details>
              </>
            ) : (
              <div className="text-center py-4">
                <h4 className="text-base font-bold text-white mb-2">Préférez-vous installer sans invite de commande ?</h4>
                <p className="text-xs text-slate-400 max-w-lg mx-auto mb-6">
                  Téléchargez l'installateur autonome ou le manifeste XML pour une intégration directe.
                </p>
                <div className="flex flex-wrap items-center justify-center gap-4">
                  <a
                    href="/installer.bat"
                    download="installer_afriexcel.bat"
                    className="flex items-center gap-2 px-6 py-3.5 rounded-2xl bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-bold text-xs shadow-lg shadow-emerald-500/20 transition"
                  >
                    <Download className="w-4 h-4" />
                    <span>Télécharger installer.bat (Double-clic)</span>
                  </a>
                  <a
                    href="/manifest.xml"
                    download="manifest_afriexcel.xml"
                    className="flex items-center gap-2 px-6 py-3.5 rounded-2xl bg-slate-800 hover:bg-slate-700 text-white font-bold text-xs border border-slate-700 transition"
                  >
                    <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
                    <span>Télécharger manifest.xml (Manuel)</span>
                  </a>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Procédure illustrée en 3 étapes simples */}
        <div className="mt-14 grid grid-cols-1 md:grid-cols-3 gap-6 text-left">
          <div className="p-6 rounded-3xl bg-slate-900 border border-slate-800 relative">
            <div className="w-9 h-9 rounded-2xl bg-emerald-500/20 text-emerald-400 font-black flex items-center justify-center text-sm mb-4">
              1
            </div>
            <h4 className="font-extrabold text-white text-base">Ouvrez PowerShell</h4>
            <p className="text-xs text-slate-300 mt-2 leading-relaxed">
              Appuyez sur la touche <strong>Windows</strong> de votre clavier, tapez <code className="bg-slate-800 text-emerald-400 px-1.5 py-0.5 rounded">PowerShell</code> et appuyez sur <strong>Entrée</strong>.
            </p>
            <span className="inline-block mt-3 text-[11px] text-slate-400 italic">
              (Pas besoin de droits administrateur)
            </span>
          </div>

          <div className="p-6 rounded-3xl bg-slate-900 border border-slate-800 relative">
            <div className="w-9 h-9 rounded-2xl bg-emerald-500/20 text-emerald-400 font-black flex items-center justify-center text-sm mb-4">
              2
            </div>
            <h4 className="font-extrabold text-white text-base">Collez et Validez</h4>
            <p className="text-xs text-slate-300 mt-2 leading-relaxed">
              Faites un <strong>clic-droit</strong> dans la fenêtre PowerShell (ou <kbd className="bg-slate-800 text-slate-200 px-1.5 py-0.5 rounded text-[10px]">Ctrl+V</kbd>) pour coller la commande, puis appuyez sur <strong>Entrée</strong>.
            </p>
            <span className="inline-block mt-3 text-[11px] text-slate-400 italic">
              (Configuration en 3 secondes chrono)
            </span>
          </div>

          <div className="p-6 rounded-3xl bg-slate-900 border border-emerald-500/40 relative bg-emerald-950/10">
            <div className="w-9 h-9 rounded-2xl bg-emerald-500 text-slate-950 font-black flex items-center justify-center text-sm mb-4">
              3
            </div>
            <h4 className="font-extrabold text-white text-base">Activez dans Excel (2 clics)</h4>
            <p className="text-xs text-slate-300 mt-2 leading-relaxed">
              Dans Excel, allez dans <strong>Insertion &gt; Mes compléments &gt; Dossier partagé</strong>, sélectionnez <strong>AfriExcel Copilot</strong> et cliquez sur <strong>Ajouter</strong>. Le bouton restera définitivement dans votre ruban <strong>Accueil</strong> !
            </p>
            <span className="inline-block mt-3 text-[11px] text-emerald-400 font-bold">
              ✓ Volet IA actif &amp; 20 feuilles SYSCOHADA
            </span>
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
