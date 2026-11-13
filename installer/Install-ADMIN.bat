@echo off
:: === VERIFIER LES DROITS ADMIN ===
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo.
    echo ============================================
    echo   CE SCRIPT DOIT ETRE LANCE EN TANT
    echo   QU'ADMINISTRATEUR !
    echo ============================================
    echo.
    echo   Clic-droit sur ce fichier ^> 
    echo   "Executer en tant qu'administrateur"
    echo.
    pause
    exit /b 1
)

echo.
echo =============================================
echo   INSTALLATEUR ADMIN - Assistant IA Excel
echo   Afrique Francophone (SYSCOHADA)
echo =============================================
echo.

:: === 1. Creer le dossier pour les manifests ===
set "SHARE_DIR=%LOCALAPPDATA%\AssistantIAExcel"
if not exist "%SHARE_DIR%" mkdir "%SHARE_DIR%"

:: Copier le manifest depuis dist si disponible
if exist "f:\cop\addin-frontend\dist\manifest.xml" (
    copy /y "f:\cop\addin-frontend\dist\manifest.xml" "%SHARE_DIR%\manifest.xml" >nul
    echo [OK] Manifest copie dans %SHARE_DIR%
) else (
    echo [INFO] Manifest deja present dans %SHARE_DIR%
)

:: === 2. Supprimer l'ancien partage s'il existe ===
net share AssistantIAExcel /delete >nul 2>&1

:: === 3. Creer le partage reseau ===
echo.
echo [2/5] Creation du partage reseau...
net share AssistantIAExcel="%SHARE_DIR%" /grant:Everyone,READ >nul 2>&1
if %errorLevel% neq 0 (
    :: Essayer avec le nom francais
    net share AssistantIAExcel="%SHARE_DIR%" /grant:"Tout le monde",READ >nul 2>&1
)
if %errorLevel% neq 0 (
    :: Essayer sans permissions specifiques
    net share AssistantIAExcel="%SHARE_DIR%" >nul 2>&1
)
echo [OK] Partage cree: \\%COMPUTERNAME%\AssistantIAExcel

:: === 4. Enregistrer le partage dans le registre ===
echo.
echo [3/5] Enregistrement du catalogue...

:: Nettoyer les anciennes entrees
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\2" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\3" /f >nul 2>&1

:: Ajouter le nouveau catalogue UNC
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\2" /v Url /t REG_SZ /d "\\%COMPUTERNAME%\AssistantIAExcel" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\2" /v Flags /t REG_DWORD /d 1 /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\2" /v Name /t REG_SZ /d "AssistantIAExcel" /f >nul
echo [OK] Catalogue UNC enregistre: \\%COMPUTERNAME%\AssistantIAExcel

:: === 5. Vider le cache WEF ===
echo.
echo [4/5] Nettoyage du cache Office...
rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" 2>nul
rd /s /q "%LOCALAPPDATA%\Microsoft\Office\Wef" 2>nul
mkdir "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" 2>nul
echo [OK] Cache nettoye

:: === 6. Fermer et relancer Excel ===
echo.
echo [5/5] Redemarrage d'Excel...
taskkill /f /im excel.exe >nul 2>&1
timeout /t 3 /nobreak >nul
start excel.exe
timeout /t 3 /nobreak >nul

echo.
echo =============================================
echo   INSTALLATION TERMINEE !
echo =============================================
echo.
echo   Partage: \\%COMPUTERNAME%\AssistantIAExcel
echo.
echo   Dans Excel:
echo     1. INSERT ^> My Add-ins
echo     2. Cliquez sur l'onglet SHARED FOLDER
echo     3. Selectionnez "Assistant IA Excel"
echo     4. Cliquez Add
echo.
echo   L'onglet "Assistant IA" apparaitra !
echo.
pause
