# 🚀 Guide de Déploiement Gratuit sur Netlify (Drag & Drop en 30 secondes)

Ce guide explique comment mettre en ligne votre Assistant IA Excel **gratuitement et sans aucune ligne de commande complexe**, grâce à **Netlify Drop**.

---

## Étape 1 : Générer le dossier de production (`dist`)

Dans le terminal, exécutez la commande de compilation :
```bash
cd addin-frontend
npm run build
```
Cela crée un dossier optimisé et prêt à l'emploi nommé :
📁 `f:\cop\addin-frontend\dist`

---

## Étape 2 : Glisser-Déposer sur Netlify Drop

1. Rendez-vous sur la page officielle : **[https://app.netlify.com/drop](https://app.netlify.com/drop)** (ou connectez-vous à votre compte Netlify gratuit).
2. Ouvrez l'explorateur de fichiers Windows à l'emplacement `f:\cop\addin-frontend`.
3. **Glissez et déposez simplement le dossier `dist`** directement dans le rectangle pointillé de Netlify Drop dans votre navigateur.
4. En 10 secondes, Netlify déploie le site et vous fournit une URL HTTPS sécurisée gratuite, par exemple :
   `https://mon-assistant-excel-afrique.netlify.app`

> [!TIP]
> Netlify attribue automatiquement un certificat SSL (HTTPS) valide, ce qui est obligatoire pour les compléments Microsoft Office.

---

## Étape 3 : Configurer votre URL dans l'Add-in Excel

Une fois que vous avez votre URL Netlify (ex: `https://mon-assistant-excel-afrique.netlify.app`) :

### Option A : Installation automatique en 1 clic
Exécutez dans PowerShell :
```powershell
.\installer\install.ps1 -NetlifyUrl "https://mon-assistant-excel-afrique.netlify.app"
```
Le script configure immédiatement le registre Windows d'Excel avec votre URL Netlify !

### Option B : Modification manuelle du manifest.xml
Ouvrez le fichier [manifest.xml](file:///f:/cop/manifest.xml) et remplacez les occurrences de `https://mon-assistant-excel.netlify.app` par votre propre URL Netlify.

---

## Étape 4 : Lancer dans Microsoft Excel (2016 à 365)

1. Ouvrez Excel.
2. Allez dans l'onglet **Insertion** > **Mes compléments** (ou **Compléments Office**).
3. Cliquez sur l'onglet **Dossier partagé** (Shared Folder).
4. Sélectionnez **Assistant IA Excel - Afrique** et cliquez sur **Ajouter**.
5. Le bouton **"Assistant IA"** apparaît dans votre ruban et le volet s'ouvre !
