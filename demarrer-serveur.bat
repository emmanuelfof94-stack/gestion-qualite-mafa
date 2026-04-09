@echo off
chcp 65001 >nul
title Gestion Qualité — Démarrage
cd /d "%~dp0"

echo.
echo  ╔══════════════════════════════════════════════╗
echo  ║    SYSTÈME GESTION QUALITÉ — CRYSTAL SOL.   ║
echo  ╚══════════════════════════════════════════════╝
echo.

:: ── Vérifier Node.js ──
where node >nul 2>&1
if %errorlevel% neq 0 (
    echo  [ERREUR] Node.js n'est pas installe !
    echo.
    echo  Telechargez et installez Node.js depuis :
    echo  https://nodejs.org  (version LTS recommandee)
    echo.
    pause
    exit /b 1
)

:: ── Vérifier npm ──
where npm >nul 2>&1
if %errorlevel% neq 0 (
    echo  [ERREUR] npm introuvable. Reinstallez Node.js.
    pause
    exit /b 1
)

:: ── Installer les dépendances si nécessaire ──
if not exist "node_modules" (
    echo  [*] Installation des dependances (premiere fois)...
    npm install
    if %errorlevel% neq 0 (
        echo  [ERREUR] Echec npm install.
        pause
        exit /b 1
    )
    echo  [OK] Dependances installees.
    echo.
)

:: ── Créer les dossiers nécessaires ──
if not exist "data"               mkdir data
if not exist "logs"               mkdir logs
if not exist "public\photos"      mkdir public\photos
if not exist "public\pieces-jointes" mkdir public\pieces-jointes

:: ── Libérer les ports si occupés ──
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr ":3443 " ^| findstr "LISTENING"') do (
    taskkill /PID %%a /F >nul 2>&1
)
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr ":3080 " ^| findstr "LISTENING"') do (
    taskkill /PID %%a /F >nul 2>&1
)
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr ":3000 " ^| findstr "LISTENING"') do (
    taskkill /PID %%a /F >nul 2>&1
)
timeout /t 1 /nobreak >nul

:: ── Essayer PM2 (si installé) ──
where pm2 >nul 2>&1
if %errorlevel% equ 0 (
    echo  [*] Demarrage avec PM2 (redemarrage automatique)...
    pm2 delete gestion-qualite >nul 2>&1
    pm2 start ecosystem.config.js
    pm2 save >nul 2>&1
    echo.
    echo  ╔══════════════════════════════════════════════╗
    echo  ║  Serveur actif en arriere-plan via PM2      ║
    echo  ║                                              ║
    echo  ║  Admin    : https://localhost:3443/admin     ║
    echo  ║  Badge    : voir IP dans les logs PM2        ║
    echo  ╚══════════════════════════════════════════════╝
    echo.
    echo  pm2 logs gestion-qualite   = voir les logs
    echo  pm2 stop gestion-qualite   = arreter
    echo.
    pause
    exit /b 0
)

:: ── Fallback : démarrage direct Node.js ──
echo  [*] PM2 non installe — demarrage direct Node.js...
echo  (La fenetre doit rester ouverte pour que le serveur fonctionne)
echo.
echo  ╔══════════════════════════════════════════════╗
echo  ║  Serveur en cours de demarrage...           ║
echo  ║                                              ║
echo  ║  Admin    : https://localhost:3443/admin     ║
echo  ║  Ne fermez pas cette fenetre !              ║
echo  ╚══════════════════════════════════════════════╝
echo.
node server.js
pause
