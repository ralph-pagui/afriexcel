# 💻 Guide d'Installation 1-Clic pour l'Utilisateur Final

L'Assistant IA Excel fonctionne sur **toutes les versions d'Excel Desktop sous Windows de 2016 à Microsoft 365**, sans passer par AppSource.

---

### Pour installer en 1 seul clic :

1. Faites un **double-clic** sur le fichier :
   `Installer-Assistant-IA.bat`
2. Laissez le script s'exécuter (2 secondes). Il va :
   - Déposer le fichier de configuration dans `%LOCALAPPDATA%\AssistantIAExcel`.
   - Enregistrer le catalogue de confiance dans le Registre Windows (`HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs`).
   - Nettoyer le cache Office pour une prise en compte immédiate.
3. Lancez **Microsoft Excel**.
4. Allez dans le menu : **Insertion** > **Compléments** (ou **Mes compléments**) > onglet **Dossier partagé**.
5. Cliquez sur **Assistant IA Excel - Afrique** puis sur **Ajouter**.
6. Un bouton **"Assistant IA"** est désormais présent dans votre ruban Excel !

---

### En cas de pare-feu d'entreprise ou blocage de scripts :
Si Windows affiche un avertissement "SmartScreen", cliquez sur **Informations complémentaires** puis sur **Exécuter quand même**.
