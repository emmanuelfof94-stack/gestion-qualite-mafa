@echo off
chcp 65001 >nul
title Installation Démarrage Automatique — Gestion Qualité
cd /d "%~dp0"

echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║   INSTALLATION DU DÉMARRAGE AUTOMATIQUE AU BOOT     ║
echo  ╚══════════════════════════════════════════════════════╝
echo.
echo  Ce script configure le serveur pour démarrer automatiquement
echo  à chaque allumage du PC — même sans ouvrir de fenêtre.
echo.

:: Vérifier droits admin
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo  [!] Droits administrateur requis.
    echo  Clic droit sur ce fichier ^> "Exécuter en tant qu'administrateur"
    pause
    exit /b 1
)

:: S'assurer que PM2 est démarré et sauvegardé
where pm2 >nul 2>&1
if %errorlevel% neq 0 (
    echo  [!] PM2 non trouvé — installation...
    npm install -g pm2
)

:: Démarrer le service si pas encore fait
pm2 delete gestion-qualite >nul 2>&1
pm2 start ecosystem.config.js
pm2 save

:: Créer le service Windows via pm2-startup ou tâche planifiée
echo  [*] Configuration du démarrage automatique Windows...

:: Méthode : Tâche Planifiée Windows
set TASK_NAME=GestionQualite-Serveur
set SCRIPT_PATH=%~dp0demarrer-pm2-silencieux.bat

:: Créer un script silencieux pour le démarrage
(
echo @echo off
echo cd /d "%~dp0"
echo pm2 resurrect
) > "%~dp0demarrer-pm2-silencieux.bat"

:: Enregistrer la tâche planifiée au démarrage de session
schtasks /create /tn "%TASK_NAME%" /tr "\"%~dp0demarrer-pm2-silencieux.bat\"" /sc ONLOGON /rl HIGHEST /f >nul 2>&1

if %errorlevel% equ 0 (
    echo.
    echo  ✓ Tâche planifiée créée : %TASK_NAME%
    echo  ✓ Le serveur démarrera automatiquement à chaque connexion Windows
    echo.
) else (
    echo.
    echo  [!] Échec tâche planifiée. Essai méthode alternative (Registre)...
    reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "GestionQualite" /t REG_SZ /d "\"%~dp0demarrer-pm2-silencieux.bat\"" /f >nul 2>&1
    if %errorlevel% equ 0 (
        echo  ✓ Entrée registre ajoutée — démarrage automatique configuré
    ) else (
        echo  [!] Impossible de configurer automatiquement.
        echo      Ajoutez manuellement le raccourci de demarrer-pm2-silencieux.bat
        echo      dans : shell:startup
    )
)

echo.
echo  Pour vérifier : Redémarrez le PC puis ouvrez https://localhost:3443/admin
echo.
pause
