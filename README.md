# 📊 Assistant IA Excel - Spécialisé Afrique Francophone & SYSCOHADA

Un complément Microsoft Excel professionnel (Office.js) propulsé par **Google Gemini 3.6 Flash** et **Supabase**, conçu pour résoudre les tâches les plus complexes et chronophages des entreprises et cabinets en Afrique francophone.

---

## 🌟 Fonctionnalités Clés Intégrées

1. **Rapprochement Bancaire & Mobile Money (Wave, Orange Money, MTN MoMo)** :
   * Lettrage automatique des relevés et du Grand Livre (Compte 521).
   * Détection et régularisation automatique des commissions opérateurs au compte 631.

2. **Génération d'États Financiers SYSCOHADA Révisé** :
   * Conversion dynamique de balance à 6 colonnes vers le **Bilan (Actif & Passif)** et le **Compte de Résultat**.
   * Calcul automatique des soldes intermédiaires de gestion : Marge brute, Valeur Ajoutée (VA), Excédent Brut d'Exploitation (EBE), Résultat Net.

3. **Nettoyage Express d'Exports Bruts (Sage Saari 100, ERP, Banques)** :
   * Élimination des espaces insécables (`\u00A0`) rendant les calculs nuls.
   * Conversion instantanée des nombres textuels en valeurs réelles.
   * Normalisation des dates au format `JJ/MM/AAAA`.

4. **Livre de Paie & Fiscalité Sociale par Pays** :
   * Barèmes et cotisations sociales intégrés : **Côte d'Ivoire (CNPS, ITS)**, **Sénégal (IPRES, IR)**, **Cameroun (CNPS, IRPP)**, **RDC (CNSS, IPR)**, Bénin, Mali, etc.
   * Calcul dynamique du Net à payer et des charges patronales.

5. **Balance Âgée & Relance Créances Clients** :
   * Ventilation des factures échues (0-30j, 31-60j, 61-90j, +90j) avec alertes de provisionnement OHADA.

6. **Paiement Fluide par Mobile Money** :
   * Module intégré Wave, Orange Money, MTN MoMo, Moov Money avec activation instantanée par webhook.

---

## 🖥️ Compatibilité Garantie

* **Versions d'Excel Desktop** : Excel 2016, Excel 2019, Excel 2021 et Microsoft 365.
* **Excel Online** : Compatible Office sur le Web.
* **Système d'exploitation** : Windows 10, Windows 11.

---

## 📂 Structure du Répertoire

```text
f:\cop\
├── addin-frontend/             # Interface React (Office.js + Tailwind CSS)
│   ├── src/
│   │   ├── components/         # Onglets métiers (Chat, Rapprochement, SYSCOHADA, Paie...)
│   │   ├── services/           # Office.js API, Gemini 3.6 Flash, Supabase & Paiement
│   │   ├── App.tsx             # Application principale
│   │   └── main.tsx            # Initialisation Office.js avec fallback Web
│   ├── public/assets/          # Icônes officielles Excel Ribbon (16, 32, 64, 80, 128px)
│   ├── vite.config.ts          # Compilation compatible ES2015 & Netlify Drop
│   └── package.json
├── manifest.xml                # Manifeste officiel Office Add-in
├── supabase/                   # Configuration Backend & IA
│   ├── functions/
│   │   ├── ai-assistant/       # Edge Function Deno avec Google Gemini 3.6 Flash
│   │   └── mobile-money-webhook/ # Webhook pour Wave, CinetPay, FedaPay
│   └── schema.sql              # Schéma PostgreSQL (utilisateurs, abonnements, crédits)
├── installer/                  # Installeur Windows 1-Clic SANS AppSource
│   ├── Installer-Assistant-IA.bat
│   └── install.ps1
└── DEPLOY_NETLIFY.md           # Guide de mise en ligne gratuite par Drag & Drop
```

---

## 🚀 Démarrage Rapide

### 1. Développement local
```bash
cd addin-frontend
npm install
npm run dev
```

### 2. Déploiement gratuit sur Netlify (Drag & Drop)
```bash
cd addin-frontend
npm run build
```
Glissez-déposez le dossier `addin-frontend/dist` sur [https://app.netlify.com/drop](https://app.netlify.com/drop).

### 3. Installation dans Excel
Double-cliquez sur `installer/Installer-Assistant-IA.bat` et ouvrez Excel !
