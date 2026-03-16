@echo off
chcp 65001 >nul
title Compilation — Timbres Fiscaux

echo.
echo ============================================================
echo   Compilation du Registre des Timbres Fiscaux
echo   Produit : dist\timbres-fiscaux.exe
echo ============================================================
echo.

:: ── 1. Vérification Python ──────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo   ERREUR — Python n'est pas trouvé dans le PATH.
    echo.
    echo   Installez Python depuis : https://www.python.org/downloads/
    echo   IMPORTANT : cochez "Add Python to PATH" pendant l'installation.
    echo.
    pause
    exit /b 1
)
echo   Python détecté :
python --version
echo.

:: ── 2. Installation des dépendances ─────────────────────────
echo   Installation des dépendances (flask, pypdf, openpyxl, pyinstaller)...
pip install flask pypdf openpyxl pyinstaller --quiet
if errorlevel 1 (
    echo.
    echo   ERREUR — L'installation des dépendances a échoué.
    echo   Vérifiez votre connexion Internet et les droits d'exécution.
    echo.
    pause
    exit /b 1
)
echo   Dépendances installées.
echo.

:: ── 3. Compilation PyInstaller ───────────────────────────────
echo   Compilation en cours (peut prendre 1-2 minutes)...
pyinstaller --onefile --noconsole --name timbres-fiscaux app.py
if errorlevel 1 (
    echo.
    echo   ERREUR — La compilation PyInstaller a échoué.
    echo   Consultez les messages ci-dessus pour le détail.
    echo.
    pause
    exit /b 1
)

:: ── 4. Succès ────────────────────────────────────────────────
echo.
echo ============================================================
echo   Compilation réussie !
echo ============================================================
echo.
echo   Fichier produit : dist\timbres-fiscaux.exe
echo.
echo   Étapes suivantes :
echo   ──────────────────────────────────────────────────────
echo   1. Copier dist\timbres-fiscaux.exe dans :
echo         \\SERVEUR\COMMUN\GESTION-TF\
echo.
echo   2. Sur chaque poste utilisateur, créer un raccourci
echo      bureau vers :
echo         \\SERVEUR\COMMUN\GESTION-TF\timbres-fiscaux.exe
echo.
echo   3. Utilisation : double-clic sur le raccourci.
echo      Le navigateur s'ouvre automatiquement.
echo.
echo   4. Les données sont stockées dans :
echo         \\SERVEUR\COMMUN\GESTION-TF\data\
echo   ──────────────────────────────────────────────────────
echo.
pause
