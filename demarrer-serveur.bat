@echo off
chcp 65001 >nul
title Gestion Qualité — Démarrage PM2
cd /d "%~dp0"

echo.
echo  ╔══════════════════════════════════════════╗
echo  ║   SYSTÈME GESTION QUALITÉ — DÉMARRAGE   ║
echo  ╚══════════════════════════════════════════╝
echo.

:: Vérifier si PM2 est installé
where pm2 >nul 2>&1
if %errorlevel% neq 0 (
    echo  [!] PM2 non trouve — installation en cours...
    npm install -g pm2
)

:: Arrêter l'instance précédente proprement
pm2 delete gestion-qualite >nul 2>&1

:: Tuer les anciens processus sur les ports
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr ":3443" ^| findstr "LISTENING"') do (
    taskkill /PID %%a /F >nul 2>&1
)
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr ":3080" ^| findstr "LISTENING"') do (
    taskkill /PID %%a /F >nul 2>&1
)

:: Créer le dossier logs si nécessaire
if not exist "logs" mkdir logs

:: Démarrer avec PM2
echo  [*] Démarrage du serveur avec PM2...
pm2 start ecosystem.config.js

:: Sauvegarder
pm2 save

echo.
echo  ✓ Serveur actif en arrière-plan (redémarrage automatique activé)
echo.
echo  Commandes utiles:
echo    pm2 status                     — voir l'état
echo    pm2 logs gestion-qualite       — voir les logs
echo    pm2 restart gestion-qualite    — redémarrer
echo    pm2 stop gestion-qualite       — arrêter
echo.
pause
