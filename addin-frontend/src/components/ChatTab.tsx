import React, { useState, useRef, useEffect } from 'react';
import { 
  Send, 
  Sparkles, 
  CheckCircle, 
  ArrowRight, 
  RefreshCw, 
  FileSpreadsheet, 
  Brain, 
  ChevronDown, 
  ChevronUp, 
  FolderPlus, 
  Folder, 
  Trash2, 
  Plus, 
  TrendingUp, 
  ShieldCheck, 
  Smartphone,
  Zap,
  HelpCircle,
  Download
} from 'lucide-react';
import { ChatMessage, AfricanCountry, SUPPORTED_COUNTRIES, BusinessPlanProject } from '../types';
import { AIService } from '../services/aiService';
import { OfficeService } from '../services/officeService';
import { PaymentService } from '../services/paymentService';
import { PythonBackendService } from '../services/pythonBackendService';

interface ChatTabProps {
  country: AfricanCountry;
  onActionExecuted?: () => void;
  onOpenPayment?: () => void;
}

const STORAGE_KEY = 'afriexcel_bp_projects_v2';

// Modèles de démarrage à fort taux de conversion pour l'Afrique francophone
const SPECIALIZED_STARTERS = [
  {
    icon: '🐔',
    title: 'Élevage avicole moderne',
    subtitle: '5 000 poulets de chair / pondeuses',
    prompt: "Génère un business plan bancable SYSCOHADA complet de 19 feuilles pour une ferme avicole moderne de 5 000 poulets de chair. Investissement initial estimé à 15 000 000 FCFA avec apport personnel de 30% et emprunt bancaire sur 4 ans à 9.5%. Calcule la trésorerie mensuelle, le BFR et le ratio DSCR pour le comité de crédit."
  },
  {
    icon: '🧼',
    title: 'Savonnerie & Cosmétiques karité',
    subtitle: 'Unité semi-industrielle & gels',
    prompt: "Conçois le modèle financier bancable pour une unité semi-industrielle de fabrication de savons au beurre de karité et cosmétiques naturels. Investissement de 12 000 000 FCFA (cuves inox, moules, camionnette). Démontre la rentabilité avec un Compte de Résultat SYSCOHADA (SIG), seuil de rentabilité et VAN/TRI sur 5 ans."
  },
  {
    icon: '🍍',
    title: 'Transformation agroalimentaire',
    subtitle: 'Jus de fruits locaux & farines',
    prompt: "Génère le business plan bancable de 19 feuilles pour une usine de pasteurisation et mise en bouteille de jus de fruits locaux (mangue, ananas, bissap). Investissement de 22 000 000 FCFA avec apport de 35%. Prévois un BFR de 45 jours de stocks et le plan de trésorerie mensuel sans découvert."
  },
  {
    icon: '🥖',
    title: 'Boulangerie & Pâtisserie moderne',
    subtitle: 'Fournil, boutique & salon',
    prompt: "Crée le business plan d'une boulangerie-pâtisserie moderne avec four rotatif et salon de dégustation. Investissement 28 000 000 FCFA. Établis la masse salariale (boulangers, vendeurs), le tableau d'amortissement de la dette bancaire et la couverture de dette DSCR."
  },
  {
    icon: '🚚',
    title: 'Transport & Logistique urbaine',
    subtitle: 'Flotte de livraison frigorifique',
    prompt: "Modélise un projet d'entreprise de transport et livraison urbaine sous température dirigée (flotte de 3 véhicules utilitaires). Investissement de 18 000 000 FCFA. Intègre le carburant, maintenance, amortissements linéaires sur 4 ans et flux de trésorerie prévisionnels."
  },
  {
    icon: '🏥',
    title: 'Clinique & Soins de proximité',
    subtitle: 'Cabinet médical & laboratoire',
    prompt: "Établis le dossier financier pour la création d'un centre de santé et de soins de proximité (consultations, laboratoire d'analyses, pharmacie intégrée). Investissement 35 000 000 FCFA avec prêt bancaire de 25 000 000 FCFA sur 5 ans. Présente le compte de résultat et bilan SYSCOHADA équilibré."
  }
];

export const TEMPLATE_OPTIONS = [
  { id: 'auto', name: 'Auto (IA)', icon: '✨', desc: 'Choix intelligent selon le secteur' },
  { id: 'emerald_africa', name: 'Émeraude Panafricaine', icon: '🌿', desc: 'Agro, Terroir & Éco-verte' },
  { id: 'sahara_commerce', name: 'Sahara & Terre Cuite', icon: '🏺', desc: 'Commerce, Jouets & Détail' },
  { id: 'royal_navy', name: 'Bleu Royal', icon: '🏛️', desc: 'Banque, Industrie & Finance' },
  { id: 'cyber_indigo', name: 'Cyber Indigo', icon: '⚡', desc: 'Tech, Startups & Innovation' },
];

export const ChatTab: React.FC<ChatTabProps> = ({ country, onActionExecuted, onOpenPayment }) => {
  const countryInfo = SUPPORTED_COUNTRIES[country];
  const [selectedTemplate, setSelectedTemplate] = useState<string>('auto');

  // Gestion des Projets Multi-Sessions persistants dans localStorage
  const [projects, setProjects] = useState<BusinessPlanProject[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          return parsed.map(p => ({
            ...p,
            messages: p.messages.map((m: any) => ({
              ...m,
              timestamp: new Date(m.timestamp)
            }))
          }));
        }
      }
    } catch (e) {
      console.warn("Erreur chargement projets depuis localStorage :", e);
    }

    // Projet initial par défaut
    const defaultProj: BusinessPlanProject = {
      id: 'proj_' + Date.now(),
      title: 'Mon Premier Business Plan',
      productTopic: 'Général',
      country: country,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      messages: [
        {
          id: 'welcome',
          sender: 'assistant',
          text: `Bonjour ! Je suis votre **Copilote IA Business Plan Bancable** certifié pour les normes **SYSCOHADA** et les critères des comités de crédit bancaire (${countryInfo.name} - ${countryInfo.currencySymbol}).\n\n🎯 Choisissez un secteur clé ci-dessous ou décrivez votre projet pour générer votre dossier complet de 19 feuilles directement dans Excel.`,
          timestamp: new Date()
        }
      ]
    };
    return [defaultProj];
  });

  const [activeProjectId, setActiveProjectId] = useState<string>(() => {
    return projects[0]?.id || 'proj_default';
  });

  const [isProjectDrawerOpen, setIsProjectDrawerOpen] = useState(false);
  const [isMobileMoneyModalOpen, setIsMobileMoneyModalOpen] = useState(false);
  const [expandedThoughts, setExpandedThoughts] = useState<Record<string, boolean>>({});

  // Sauvegarde automatique permanente des projets
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(projects));
    } catch (e) {
      console.error("Erreur sauvegarde localStorage projets :", e);
    }
  }, [projects]);

  const activeProject = projects.find(p => p.id === activeProjectId) || projects[0];

  const [input, setInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [activeProject?.messages, isLoading]);

  // Création d'un nouveau projet
  const handleCreateNewProject = (customTitle?: string, customTopic?: string) => {
    const newId = 'proj_' + Date.now();
    const newProj: BusinessPlanProject = {
      id: newId,
      title: customTitle || `Dossier #${projects.length + 1} (${countryInfo.currencySymbol})`,
      productTopic: customTopic || 'Projet bancable',
      country: country,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      messages: [
        {
          id: 'welcome_' + newId,
          sender: 'assistant',
          text: `Nouveau dossier de financement bancable ouvert pour **${countryInfo.name}**.\n\nQuel est votre projet d'investissement ? Je suis prêt à concevoir l'ensemble des 19 feuilles prévisionnelles SYSCOHADA.`,
          timestamp: new Date()
        }
      ]
    };

    setProjects(prev => [newProj, ...prev]);
    setActiveProjectId(newId);
    setIsProjectDrawerOpen(false);
  };

  // Suppression d'un projet
  const handleDeleteProject = (projId: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (projects.length <= 1) {
      alert("Vous devez conserver au moins un dossier de travail.");
      return;
    }
    if (confirm("Voulez-vous vraiment supprimer ce dossier de business plan ?")) {
      const remaining = projects.filter(p => p.id !== projId);
      setProjects(remaining);
      if (activeProjectId === projId) {
        setActiveProjectId(remaining[0].id);
      }
    }
  };

  const toggleThought = (msgId: string) => {
    setExpandedThoughts(prev => ({
      ...prev,
      [msgId]: !prev[msgId]
    }));
  };

  // Envoi de message avec mémoire contextuelle complète
  const handleSend = async (textToSend?: string) => {
    const query = textToSend || input;
    if (!query.trim() || isLoading) return;

    const userMsg: ChatMessage = {
      id: Date.now().toString(),
      sender: 'user',
      text: query,
      timestamp: new Date()
    };

    // Mettre à jour l'historique du projet actif immédiatement
    const updatedMessages = [...activeProject.messages, userMsg];
    setProjects(prev =>
      prev.map(p =>
        p.id === activeProjectId
          ? { ...p, messages: updatedMessages, updatedAt: new Date().toISOString() }
          : p
      )
    );

    if (!textToSend) setInput('');
    setIsLoading(true);

    try {
      // 1. Lire le contexte Excel actuel (sélection + feuille complète pour modifications)
      const rangeData = await OfficeService.getSelectedRangeData();
      const workbookInfo = await OfficeService.getWorkbookOverview();
      const detectedErrors = (rangeData?.values && rangeData?.formulas)
        ? OfficeService.detectFormulaErrors(rangeData.values, rangeData.formulas)
        : [];
      
      // Lire le contenu complet de la feuille active pour permettre les modifications chirurgicales
      let activeSheetData: any = null;
      try {
        activeSheetData = await OfficeService.getActiveSheetFullData();
      } catch (e) {
        console.warn("Lecture feuille active (non-bloquant) :", e);
      }

      // 2. Interroger Gemini 3.6 Flash avec l'HISTORIQUE CONVERSATIONNEL COMPLET du projet
      const response = await AIService.queryAssistant(
        query,
        country,
        {
          address: rangeData?.address,
          values: rangeData?.values,
          formulas: rangeData?.formulas,
          activeSheet: workbookInfo?.activeSheet,
          workbookSheets: workbookInfo?.sheets,
          detectedErrors: detectedErrors,
          activeSheetData: activeSheetData
        },
        updatedMessages, // Envoi de tout le contexte pour mémoire à long terme
        selectedTemplate
      );

      const assistantMsg: ChatMessage = {
        id: (Date.now() + 1).toString(),
        sender: 'assistant',
        text: response.message,
        thought: response.thought,
        timestamp: new Date(),
        actions: response.actions,
        suggestedPrompt: response.suggestedFollowUp,
        status: response.actions && response.actions.length > 0 ? 'pending' : undefined
      };

      // Si le projet a un titre générique, l'adapter selon le sujet détecté
      let newTitle = activeProject.title;
      if (activeProject.title === 'Mon Premier Business Plan' || activeProject.title.startsWith('Dossier #')) {
        const firstAction = response.actions?.find((a: any) => a.type === 'CREATE_BUSINESS_PLAN' || a.type === 'UPDATE_BUSINESS_PLAN');
        if (firstAction && (firstAction as any).projectName) {
          newTitle = (firstAction as any).projectName;
        } else if (firstAction && (firstAction as any).productTopic) {
          newTitle = `Projet ${(firstAction as any).productTopic}`;
        }
      }

      setProjects(prev =>
        prev.map(p =>
          p.id === activeProjectId
            ? {
                ...p,
                title: newTitle,
                messages: [...updatedMessages, assistantMsg],
                updatedAt: new Date().toISOString()
              }
            : p
        )
      );

      // Ouvrir automatiquement le panneau d'analyse s'il y a un raisonnement financier
      if (response.thought) {
        setExpandedThoughts(prev => ({
          ...prev,
          [assistantMsg.id]: true
        }));
      }

    } catch (err: any) {
      const errorMsg: ChatMessage = {
        id: (Date.now() + 1).toString(),
        sender: 'assistant',
        text: `Désolé, une difficulté est survenue : ${err?.message || 'Erreur inconnue'}.\n\nVous pouvez reformuler ou cliquer sur un bouton d'action directe.`,
        timestamp: new Date()
      };
      setProjects(prev =>
        prev.map(p =>
          p.id === activeProjectId
            ? { ...p, messages: [...updatedMessages, errorMsg], updatedAt: new Date().toISOString() }
            : p
        )
      );
    } finally {
      setIsLoading(false);
    }
  };

  const [isApplying, setIsApplying] = useState<string | null>(null);
  const [actionProgress, setActionProgress] = useState<string | null>(null);
  const [actionError, setActionError] = useState<{ id: string; message: string } | null>(null);

  // Application chirurgicale dans Excel
  const handleApplyAction = async (msgId: string, actions?: any[]) => {
    if (!actions || actions.length === 0) return;

    setIsApplying(msgId);
    setActionError(null);
    setActionProgress(`Initialisation de l'architecture Excel...`);

    try {
      for (let i = 0; i < actions.length; i++) {
        const act = actions[i];
        if (act.type === 'CREATE_BUSINESS_PLAN') {
          setActionProgress(`Génération des ${(act as any).sheetCount || 20} feuilles bancables...`);
        } else if (act.type === 'UPDATE_BUSINESS_PLAN') {
          setActionProgress(`Mise à jour chirurgicale des 20 feuilles bancables...`);
        } else {
          setActionProgress(`Application dans Excel (${i + 1}/${actions.length})...`);
        }

        const res = await OfficeService.executeAction(act, (prog) => {
          setActionProgress(prog);
        });
        if (res && !res.success) {
          throw new Error(res.message);
        }
      }

      // Marquer le message comme appliqué
      setProjects(prev =>
        prev.map(p =>
          p.id === activeProjectId
            ? {
                ...p,
                messages: p.messages.map(m => m.id === msgId ? { ...m, status: 'applied' } : m)
              }
            : p
        )
      );

      onActionExecuted?.();
    } catch (e: any) {
      console.error("Erreur application dans Excel :", e);
      setActionError({ id: msgId, message: e?.message || "Erreur lors de la mise à jour dans Excel." });
    } finally {
      setIsApplying(null);
      setActionProgress(null);
    }
  };

  // Citations de cellules cliquables
  const renderMessageTextWithCitations = (text: string) => {
    const citationRegex = /(\[[A-Za-z0-9_!:]+\])/g;
    const parts = text.split(citationRegex);

    return parts.map((part, i) => {
      if (part.startsWith('[') && part.endsWith(']')) {
        const citation = part.slice(1, -1);
        const isCellRange = /^[A-Za-z0-9_!:]*?[A-Z]{1,3}[0-9]{1,7}(?::[A-Z]{1,3}[0-9]{1,7})?$/i.test(citation);
        if (isCellRange) {
          return (
            <button
              key={i}
              type="button"
              onClick={() => OfficeService.selectRange(citation)}
              title={`Cliquer pour sélectionner [${citation}] dans Excel`}
              className="inline-flex items-center gap-0.5 mx-0.5 px-1.5 py-0.5 bg-emerald-100 hover:bg-emerald-200 text-emerald-900 font-mono text-xs rounded border border-emerald-300 font-semibold cursor-pointer transition-transform active:scale-95 shadow-2xs"
            >
              <span className="text-[10px]">📍</span>
              {citation}
            </button>
          );
        }
      }
      return <span key={i}>{part}</span>;
    });
  };

  // Suggestions d'affinage financier en cours d'échange
  const ongoingPills = [
    "📊 Générer le Pack Bancable (19 feuilles)",
    "⚖️ Ajuster l'apport personnel à 30%",
    "📈 Simuler +20% sur les ventes de l'An 2",
    "🏦 Calculer le DSCR et la capacité d'emprunt",
    "💧 Vérifier la trésorerie mensuelle (12 mois)"
  ];

  // Gestion Mobile Money Modal
  const [momoOperator, setMomoOperator] = useState<'WAVE' | 'ORANGE_MONEY' | 'MTN_MOMO' | 'MOOV_MONEY'>('WAVE');
  const [momoPhone, setMomoPhone] = useState('');
  const [momoIsLoading, setMomoIsLoading] = useState(false);
  const [momoSuccess, setMomoSuccess] = useState<string | null>(null);

  const handleMobileMoneySubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setMomoIsLoading(true);
    setMomoSuccess(null);
    try {
      const res = await PaymentService.initiateMobileMoneyPayment({
        operator: momoOperator,
        phoneNumber: momoPhone,
        amount: 15000,
        currency: countryInfo.currencySymbol,
        planName: 'Pack Business Plan Bancable Illimité'
      });
      if (res.success) {
        setMomoSuccess(res.message);
      } else {
        alert(res.message);
      }
    } catch (err: any) {
      alert("Erreur de paiement : " + err.message);
    } finally {
      setMomoIsLoading(false);
    }
  };

  return (
    <div className="flex flex-col h-full bg-slate-50 relative">
      {/* Barre supérieure du Dossier Actif & Sélecteur de Projets */}
      <div className="bg-white border-b border-slate-200 px-3 py-2 flex items-center justify-between shadow-2xs shrink-0">
        <button
          onClick={() => setIsProjectDrawerOpen(prev => !prev)}
          className="flex items-center gap-2 hover:bg-slate-100 p-1.5 rounded-lg transition-colors text-left max-w-[70%]"
          title="Gérer les dossiers de Business Plan"
        >
          <div className="w-7 h-7 rounded-lg bg-emerald-50 border border-emerald-200 flex items-center justify-center text-emerald-700 shrink-0">
            <Folder className="w-4 h-4" />
          </div>
          <div className="min-w-0">
            <div className="text-[10px] text-slate-400 uppercase font-bold tracking-wider flex items-center gap-1">
              <span>Dossier en cours</span>
              <span className="bg-emerald-100 text-emerald-800 text-[9px] px-1 rounded font-mono font-normal">
                {projects.length} dossier{projects.length > 1 ? 's' : ''}
              </span>
            </div>
            <div className="text-xs font-bold text-slate-900 truncate">
              {activeProject.title}
            </div>
          </div>
          <ChevronDown className="w-3.5 h-3.5 text-slate-400 shrink-0 ml-1" />
        </button>

        <div className="flex items-center gap-1.5">
          <button
            onClick={() => handleCreateNewProject()}
            className="flex items-center gap-1 text-[11px] bg-slate-900 hover:bg-slate-800 text-white font-semibold px-2.5 py-1.5 rounded-lg shadow-sm transition-transform active:scale-95"
            title="Créer un nouveau dossier de Business Plan"
          >
            <Plus className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">Nouveau</span>
          </button>

          <button
            onClick={() => setIsMobileMoneyModalOpen(true)}
            className="flex items-center gap-1 text-[11px] bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white font-bold px-2.5 py-1.5 rounded-lg shadow-sm transition-transform active:scale-95"
            title="Activer la licence illimitée via Wave, Orange Money ou MTN MoMo"
          >
            <Zap className="w-3.5 h-3.5 text-amber-300 fill-amber-300" />
            <span>Passer Pro</span>
          </button>
        </div>
      </div>

      {/* Tiroir de gestion des dossiers (Project Drawer) */}
      {isProjectDrawerOpen && (
        <div className="absolute top-[49px] left-0 right-0 z-30 bg-white border-b border-slate-300 shadow-xl p-3 space-y-2 max-h-[380px] overflow-y-auto animate-in slide-in-from-top-2 duration-150">
          <div className="flex items-center justify-between pb-2 border-b border-slate-100">
            <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
              <FolderPlus className="w-4 h-4 text-emerald-600" />
              Vos Dossiers de Business Plan ({projects.length})
            </span>
            <button
              onClick={() => handleCreateNewProject()}
              className="text-xs text-emerald-700 hover:text-emerald-800 font-semibold flex items-center gap-1"
            >
              <Plus className="w-3.5 h-3.5" />
              Créer un dossier
            </button>
          </div>

          <div className="space-y-1.5">
            {projects.map(proj => {
              const isActive = proj.id === activeProjectId;
              return (
                <div
                  key={proj.id}
                  onClick={() => {
                    setActiveProjectId(proj.id);
                    setIsProjectDrawerOpen(false);
                  }}
                  className={`flex items-center justify-between p-2.5 rounded-lg border text-left cursor-pointer transition-all ${
                    isActive
                      ? 'border-emerald-500 bg-emerald-50/60 ring-1 ring-emerald-500'
                      : 'border-slate-200 hover:border-slate-300 hover:bg-slate-50'
                  }`}
                >
                  <div className="min-w-0 flex-1">
                    <div className="text-xs font-bold text-slate-900 truncate flex items-center gap-1.5">
                      {proj.title}
                      {isActive && (
                        <span className="text-[9px] bg-emerald-600 text-white font-medium px-1.5 py-0.2 rounded-full">
                          Actif
                        </span>
                      )}
                    </div>
                    <div className="text-[10px] text-slate-400 mt-0.5">
                      {proj.messages.length} échange{proj.messages.length > 1 ? 's' : ''} • Mis à jour le {new Date(proj.updatedAt).toLocaleDateString()}
                    </div>
                  </div>

                  {projects.length > 1 && (
                    <button
                      onClick={(e) => handleDeleteProject(proj.id, e)}
                      className="text-slate-400 hover:text-red-600 p-1 rounded transition-colors ml-2"
                      title="Supprimer ce dossier"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Zone principale des messages */}
      <div className="flex-1 overflow-y-auto p-3 space-y-3">
        {/* Affichage des Starters Avancés quand le projet est nouveau */}
        {activeProject.messages.length <= 1 && (
          <div className="space-y-2 mb-3">
            <div className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
              <Sparkles className="w-3.5 h-3.5 text-emerald-600" />
              <span>Secteurs bancables prêts à générer en 1 clic :</span>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
              {SPECIALIZED_STARTERS.map((starter, idx) => (
                <button
                  key={idx}
                  onClick={() => handleSend(starter.prompt)}
                  className="p-2.5 bg-white hover:bg-emerald-50/50 border border-slate-200 hover:border-emerald-300 rounded-xl text-left transition-all shadow-2xs group flex items-start gap-2.5 active:scale-[0.98]"
                >
                  <span className="text-xl shrink-0 group-hover:scale-110 transition-transform">
                    {starter.icon}
                  </span>
                  <div className="min-w-0">
                    <div className="font-bold text-xs text-slate-900 group-hover:text-emerald-900">
                      {starter.title}
                    </div>
                    <div className="text-[10px] text-slate-500 mt-0.5 line-clamp-1">
                      {starter.subtitle}
                    </div>
                  </div>
                </button>
              ))}
            </div>
          </div>
        )}

        {/* Liste des messages de la conversation */}
        {activeProject.messages.map(msg => (
          <div
            key={msg.id}
            className={`flex flex-col ${msg.sender === 'user' ? 'items-end' : 'items-start'}`}
          >
            <div
              className={`max-w-[92%] rounded-2xl p-3 text-xs leading-relaxed shadow-xs ${
                msg.sender === 'user'
                  ? 'bg-emerald-600 text-white rounded-br-xs'
                  : 'bg-white text-slate-800 border border-slate-200 rounded-bl-xs'
              }`}
            >
              {/* Carte de Réflexion / Analyse du Comité de Crédit (Deep Reasoning) */}
              {msg.thought && (
                <div className="mb-2.5 border border-emerald-200/90 bg-emerald-50/60 rounded-xl overflow-hidden shadow-2xs">
                  <button
                    type="button"
                    onClick={() => toggleThought(msg.id)}
                    className="w-full flex items-center justify-between p-2 text-[11px] font-bold text-emerald-950 bg-emerald-100/50 hover:bg-emerald-100 transition-colors"
                  >
                    <span className="flex items-center gap-1.5">
                      <Brain className="w-3.5 h-3.5 text-emerald-700 shrink-0" />
                      <span>Analyse du Comité de Crédit & Faisabilité</span>
                    </span>
                    {expandedThoughts[msg.id] ? (
                      <ChevronUp className="w-3.5 h-3.5 text-emerald-700" />
                    ) : (
                      <ChevronDown className="w-3.5 h-3.5 text-emerald-700" />
                    )}
                  </button>
                  {expandedThoughts[msg.id] && (
                    <div className="p-2.5 text-[11px] text-slate-700 font-sans leading-relaxed whitespace-pre-line border-t border-emerald-200/70 bg-white/70">
                      {msg.thought}
                    </div>
                  )}
                </div>
              )}

              {/* Contenu textuel avec citations interactives */}
              <div className="whitespace-pre-line">
                {msg.sender === 'assistant' ? renderMessageTextWithCitations(msg.text) : msg.text}
              </div>

              {/* Carte d'Aperçu & Bouton d'Action Excel */}
              {msg.actions && msg.actions.length > 0 && (
                <div className="mt-3 pt-2.5 border-t border-slate-100">
                  {msg.status === 'applied' ? (
                    <div className="flex items-center gap-1.5 text-xs text-emerald-800 font-semibold bg-emerald-50 py-1.5 px-2.5 rounded-lg border border-emerald-200">
                      <CheckCircle className="w-4 h-4 text-emerald-600 shrink-0" />
                      <span>Business Plan injecté avec succès dans Excel !</span>
                    </div>
                  ) : (
                    <div className="space-y-2">
                      <div className="bg-slate-50 border border-slate-200 rounded-xl p-2.5 text-xs text-slate-700 space-y-1.5">
                        <div className="flex items-center justify-between font-bold text-slate-800 border-b border-slate-200 pb-1.5">
                          <span className="flex items-center gap-1">
                            <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
                            Structure financière prête ({msg.actions.length} action{msg.actions.length > 1 ? 's' : ''})
                          </span>
                          <span className="text-[10px] bg-emerald-100 text-emerald-800 font-bold px-1.5 py-0.5 rounded">
                            Conforme SYSCOHADA
                          </span>
                        </div>

                        {msg.actions.map((act: any, idx: number) => (
                          <div key={idx} className="flex items-start gap-1.5 text-[11px] text-slate-700">
                            <span className="text-slate-400 font-mono text-[10px] mt-0.5">#{idx + 1}</span>
                            <div className="flex-1">
                              {(act.type === 'CREATE_BUSINESS_PLAN' || act.type === 'UPDATE_BUSINESS_PLAN') && (
                                <span>
                                  <strong className="text-emerald-900 font-bold">{act.type === 'UPDATE_BUSINESS_PLAN' ? '🔄 Mise à jour Globale' : 'Pack Complet'}</strong> : {act.sheetCount || 20} feuilles bancables ({act.projectName || act.productTopic || "Projet"})
                                </span>
                              )}
                              {act.type === 'CREATE_TABLE_SECTION' && (
                                <span><strong className="text-slate-900 font-semibold">Tableau</strong> : {act.bannerTitle || "Section financière"}</span>
                              )}
                              {act.type === 'APPLY_FORMULA' && (
                                <span><strong className="text-blue-900 font-semibold">Formule</strong> : <code className="bg-white px-1 py-0.5 rounded border border-slate-200 font-mono text-blue-800">{act.formula}</code> sur {act.targetCell}</span>
                              )}
                              {act.type === 'WRITE_RANGE' && (
                                <span><strong className="text-slate-900 font-semibold">Données</strong> : Plage {act.address} ({act.values?.length || 0} lignes)</span>
                              )}
                              {act.type === 'MODIFY_CELLS' && (
                                <span><strong className="text-amber-900 font-semibold">✏️ Modification</strong> : {act.changes?.length || 0} cellule(s) sur « {act.sheetName} »</span>
                              )}
                              {act.type === 'WRITE_TO_SHEET' && (
                                <span><strong className="text-blue-900 font-semibold">📝 Écriture ciblée</strong> : Plage {act.address} sur « {act.sheetName} »</span>
                              )}
                              {!['CREATE_BUSINESS_PLAN', 'UPDATE_BUSINESS_PLAN', 'CREATE_TABLE_SECTION', 'APPLY_FORMULA', 'WRITE_RANGE', 'MODIFY_CELLS', 'WRITE_TO_SHEET'].includes(act.type) && (
                                <span><strong className="text-slate-900 font-semibold">{act.type}</strong> : {act.address || act.sheetName || "Action directe"}</span>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>

                      <button
                        onClick={() => handleApplyAction(msg.id, msg.actions)}
                        disabled={isApplying === msg.id}
                        className="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-emerald-400 text-white font-bold py-2 px-3 rounded-xl text-xs shadow transition-all active:scale-95"
                      >
                        {isApplying === msg.id ? (
                          <>
                            <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                            <span>{actionProgress || "Application dans Excel..."}</span>
                          </>
                        ) : (
                          <>
                            <FileSpreadsheet className="w-4 h-4" />
                            <span>{msg.actions.some((a: any) => a.type === 'UPDATE_BUSINESS_PLAN') ? "🔄 Mettre à jour et harmoniser dans Excel" : "Générer et appliquer directement dans Excel"}</span>
                          </>
                        )}
                      </button>

                      {msg.actions.some((act: any) => act.type === 'CREATE_BUSINESS_PLAN' || act.type === 'UPDATE_BUSINESS_PLAN') && (
                        <button
                          onClick={() => {
                            const bpAct = msg.actions?.find((a: any) => a.type === 'CREATE_BUSINESS_PLAN' || a.type === 'UPDATE_BUSINESS_PLAN') as any;
                            if (bpAct?.file_base64) {
                              PythonBackendService.downloadXlsx(bpAct.file_base64, bpAct.filename || "BusinessPlan_OHADA.xlsx");
                            } else {
                              PythonBackendService.generateBusinessPlan({
                                prompt: bpAct?.productTopic || bpAct?.projectName || "Business Plan",
                                currency: bpAct?.currencySymbol || "FCFA"
                              }).then(pyPlan => {
                                if (pyPlan && pyPlan.file_base64) {
                                  PythonBackendService.downloadXlsx(pyPlan.file_base64, pyPlan.filename);
                                }
                              });
                            }
                          }}
                          className="w-full flex items-center justify-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold py-1.5 px-3 rounded-xl text-xs border border-slate-300 transition-all active:scale-95 shadow-sm"
                        >
                          <Download className="w-3.5 h-3.5 text-slate-600" />
                          <span>📥 Télécharger le classeur complet (.xlsx)</span>
                        </button>
                      )}

                      {actionError && actionError.id === msg.id && (
                        <div className="text-[11px] text-red-700 bg-red-50 p-2 rounded-lg border border-red-200">
                          ⚠️ {actionError.message}
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Prompt de suivi suggéré */}
            {msg.suggestedPrompt && (
              <button
                onClick={() => handleSend(msg.suggestedPrompt)}
                className="mt-1 text-[11px] text-emerald-800 hover:text-emerald-950 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 px-2.5 py-1 rounded-full flex items-center gap-1 transition-all"
              >
                <span>💡 {msg.suggestedPrompt}</span>
                <ArrowRight className="w-3 h-3 text-emerald-600" />
              </button>
            )}

            <span className="text-[10px] text-slate-400 mt-1 px-1">
              {msg.timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
            </span>
          </div>
        ))}

        {isLoading && (
          <div className="flex items-center gap-2.5 text-xs text-slate-600 bg-white border border-slate-200 p-2.5 rounded-xl w-fit shadow-2xs">
            <RefreshCw className="w-3.5 h-3.5 animate-spin text-emerald-600" />
            <span>Le Copilote IA analyse la faisabilité financière et prépare le classeur...</span>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* Suggestions rapides en cours de discussion */}
      {activeProject.messages.length > 1 && (
        <div className="p-1.5 bg-slate-100 border-t border-slate-200 flex items-center gap-1.5 overflow-x-auto text-xs whitespace-nowrap scrollbar-none">
          <span className="text-[10px] text-slate-500 font-bold uppercase shrink-0 px-1">Ajustements :</span>
          {ongoingPills.map((qp, idx) => (
            <button
              key={idx}
              onClick={() => handleSend(qp)}
              className="px-2 py-0.5 bg-white hover:bg-emerald-50 text-slate-700 hover:text-emerald-900 rounded-md border border-slate-200 hover:border-emerald-300 text-[11px] transition-colors shadow-2xs shrink-0"
            >
              {qp}
            </button>
          ))}
        </div>
      )}

      {/* Sélecteur de Template Professionnel */}
      <div className="px-2.5 py-1.5 bg-slate-50 border-t border-slate-200 flex items-center gap-1.5 overflow-x-auto text-[11px] whitespace-nowrap scrollbar-none">
        <span className="text-[10px] text-slate-500 font-bold uppercase shrink-0 flex items-center gap-1">
          <Sparkles className="w-3 h-3 text-emerald-600" />
          Template :
        </span>
        {TEMPLATE_OPTIONS.map(tpl => (
          <button
            key={tpl.id}
            type="button"
            onClick={() => setSelectedTemplate(tpl.id)}
            className={`px-2 py-0.5 rounded-full border text-[11px] font-medium transition-all shrink-0 flex items-center gap-1 ${
              selectedTemplate === tpl.id
                ? 'bg-emerald-600 text-white border-emerald-700 shadow-xs font-semibold'
                : 'bg-white text-slate-600 border-slate-200 hover:border-emerald-300 hover:bg-emerald-50'
            }`}
            title={tpl.desc}
          >
            <span>{tpl.icon}</span>
            <span>{tpl.name}</span>
          </button>
        ))}
      </div>

      {/* Zone de saisie conversationnelle */}
      <div className="p-2.5 bg-white border-t border-slate-200">
        <form
          onSubmit={e => {
            e.preventDefault();
            handleSend();
          }}
          className="flex items-center gap-2"
        >
          <input
            type="text"
            value={input}
            onChange={e => setInput(e.target.value)}
            placeholder={`Posez une question ou demandez un ajustement (${countryInfo.currencySymbol})...`}
            disabled={isLoading}
            className="flex-1 text-xs border border-slate-300 rounded-xl px-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent transition-all shadow-2xs"
          />
          <button
            type="submit"
            disabled={!input.trim() || isLoading}
            className="bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-300 text-white p-2.5 rounded-xl transition-all shadow-sm shrink-0 active:scale-95"
            title="Envoyer"
          >
            <Send className="w-4 h-4" />
          </button>
        </form>
      </div>

      {/* Modal Mobile Money sans friction pour premier million */}
      {isMobileMoneyModalOpen && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-3 animate-in fade-in duration-150">
          <div className="bg-white rounded-2xl max-w-sm w-full p-4 shadow-2xl border border-slate-200 space-y-3.5">
            <div className="flex items-center justify-between border-b border-slate-100 pb-2.5">
              <div className="flex items-center gap-2">
                <div className="w-7 h-7 rounded-lg bg-emerald-100 flex items-center justify-center text-emerald-800">
                  <Smartphone className="w-4 h-4 text-emerald-700" />
                </div>
                <div>
                  <h3 className="font-bold text-xs text-slate-900">Pack Business Plan Illimité</h3>
                  <p className="text-[10px] text-slate-500">Paiement Mobile Money en 10 secondes</p>
                </div>
              </div>
              <button
                onClick={() => setIsMobileMoneyModalOpen(false)}
                className="text-slate-400 hover:text-slate-600 p-1 text-sm font-bold"
              >
                ✕
              </button>
            </div>

            <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-2.5 text-xs text-emerald-900">
              <div className="flex justify-between items-center font-bold">
                <span>Accès Complet Pro</span>
                <span className="text-sm font-extrabold text-emerald-700">15 000 {countryInfo.currencySymbol}</span>
              </div>
              <p className="text-[10px] text-emerald-700 mt-1">
                ✅ Économisez 500 000 FCFA d'honoraires de cabinet d'expertise. Modèles 100% bancables certifiés SYSCOHADA.
              </p>
            </div>

            {/* Opérateurs */}
            <div className="space-y-1.5">
              <label className="text-[11px] font-bold text-slate-700">Sélectionnez votre opérateur :</label>
              <div className="grid grid-cols-3 gap-1.5 text-xs">
                {[
                  { id: 'WAVE', name: 'Wave' },
                  { id: 'ORANGE_MONEY', name: 'Orange' },
                  { id: 'MTN_MOMO', name: 'MTN' }
                ].map(op => (
                  <button
                    key={op.id}
                    type="button"
                    onClick={() => setMomoOperator(op.id as any)}
                    className={`py-2 px-1.5 rounded-lg border font-bold text-center transition-all ${
                      momoOperator === op.id
                        ? 'border-emerald-600 bg-emerald-50 text-emerald-800 ring-1 ring-emerald-500'
                        : 'border-slate-200 text-slate-600 hover:bg-slate-50'
                    }`}
                  >
                    {op.name}
                  </button>
                ))}
              </div>
            </div>

            {/* Numéro de téléphone */}
            <form onSubmit={handleMobileMoneySubmit} className="space-y-2.5">
              <div>
                <label className="text-[11px] font-bold text-slate-700">Votre numéro {momoOperator} :</label>
                <input
                  type="tel"
                  required
                  placeholder="Ex: 07 08 09 10 11"
                  value={momoPhone}
                  onChange={e => setMomoPhone(e.target.value)}
                  className="w-full mt-1 text-xs border border-slate-300 rounded-lg px-2.5 py-2 focus:ring-2 focus:ring-emerald-500 focus:outline-none"
                />
              </div>

              {momoSuccess && (
                <div className="p-2 bg-emerald-50 border border-emerald-200 rounded-lg text-xs text-emerald-800">
                  🎉 {momoSuccess}
                </div>
              )}

              <button
                type="submit"
                disabled={momoIsLoading}
                className="w-full bg-emerald-600 hover:bg-emerald-700 disabled:bg-emerald-300 text-white font-bold py-2.5 rounded-xl text-xs shadow transition-all flex items-center justify-center gap-1.5 active:scale-95"
              >
                {momoIsLoading ? (
                  <>
                    <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                    <span>Connexion Mobile Money...</span>
                  </>
                ) : (
                  <>
                    <ShieldCheck className="w-4 h-4" />
                    <span>Valider (15 000 {countryInfo.currencySymbol})</span>
                  </>
                )}
              </button>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
