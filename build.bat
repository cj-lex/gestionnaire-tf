@echo off
title Compilation Timbres Fiscaux

echo.
echo ============================================================
echo   Compilation du Registre des Timbres Fiscaux
echo   Produit : dist\timbres-fiscaux.exe
echo ============================================================
echo.

:: 1. Verification Python
python --version >nul 2>&1
if errorlevel 1 (
    echo   ERREUR : Python n'est pas trouve dans le PATH.
    echo.
    echo   Installez Python depuis :
    echo   https://www.python.org/downloads/
    echo.
    echo   IMPORTANT : cochez "Add Python to PATH" lors de l'installation.
    echo.
    pause
    exit /b 1
)
echo   Python detecte :
python --version
echo.

:: 2. Installation des dependances
echo   Installation des dependances...
echo   (flask, pypdf, openpyxl, pyinstaller)
pip install flask pypdf openpyxl pyinstaller --quiet
if errorlevel 1 (
    echo.
    echo   ERREUR : L'installation des dependances a echoue.
    echo   Verifiez votre connexion Internet.
    echo.
    pause
    exit /b 1
)
echo   Dependances installees.
echo.

:: 3. Compilation PyInstaller
echo   Compilation en cours (1-2 minutes)...
pyinstaller --onefile --noconsole --name timbres-fiscaux app.py
if errorlevel 1 (
    echo.
    echo   ERREUR : La compilation PyInstaller a echoue.
    echo   Consultez les messages ci-dessus pour le detail.
    echo.
    pause
    exit /b 1
)

:: 4. Succes
echo.
echo ============================================================
echo   Compilation reussie !
echo ============================================================
echo.
echo   Fichier produit : dist\timbres-fiscaux.exe
echo.
echo   Etapes suivantes :
echo   ------------------------------------------------------------
echo   1. Copier dist\timbres-fiscaux.exe vers :
echo         \\SERVEUR\COMMUN\GESTION-TF\
echo.
echo   2. Sur chaque poste, creer un raccourci bureau vers :
echo         \\SERVEUR\COMMUN\GESTION-TF\timbres-fiscaux.exe
echo.
echo   3. Utilisation : double-clic sur le raccourci.
echo      Le navigateur s'ouvre automatiquement.
echo.
echo   4. Les donnees sont dans :
echo         \\SERVEUR\COMMUN\GESTION-TF\data\
echo   ------------------------------------------------------------
echo.
pause
