# PowerShell Script to manage PDF processing tasks

# Get the directory of the script
$ScriptDir = $PSScriptRoot

# --- Configuration ---
$BaseProduitsFile = Join-Path $ScriptDir "FICHIER GENERAL.xlsx"
$FicheTechniqueDir = Join-Path $ScriptDir "Fiche technique a traite"
$ExtractEanScript = Join-Path $ScriptDir "programme/extract_ean.py"
$RenameFilesScript = Join-Path $ScriptDir "programme/rename_files.py"
$ExtractImagesScript = Join-Path $ScriptDir "programme/extract_images.py"
$RequirementsFile = Join-Path $ScriptDir "programme/requirements.txt"

# --- Vérification de Python et des dépendances ---
function CheckPythonDependencies {
    # Vérifier si Python est installé
    $pythonInstalled = $false
    try {
        $pythonVersion = python --version 2>&1
        if ($pythonVersion -like "*Python 3*") {
            Write-Host "Python est installé: $pythonVersion" -ForegroundColor Green
            $pythonInstalled = $true
        } else {
            Write-Host "Python 3 n'est pas détecté" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Python n'est pas installé ou n'est pas dans le PATH" -ForegroundColor Red
    }

    if (-not $pythonInstalled) {
        Write-Host "Veuillez installer Python 3 depuis https://www.python.org/downloads/" -ForegroundColor Red
        Read-Host "Appuyez sur Entree pour quitter"
        exit
    }

    # Vérifier si pip est installé
    $pipInstalled = $false
    try {
        $pipVersion = python -m pip --version 2>&1
        Write-Host "Pip est installé: $pipVersion" -ForegroundColor Green
        $pipInstalled = $true
    } catch {
        Write-Host "Pip n'est pas installé correctement" -ForegroundColor Red
    }

    if (-not $pipInstalled) {
        Write-Host "Installation de pip..." -ForegroundColor Yellow
        try {
            Invoke-Expression "python -m ensurepip --upgrade"
            Write-Host "Pip a été installé" -ForegroundColor Green
        } catch {
            Write-Host "Échec de l'installation de pip. Veuillez l'installer manuellement." -ForegroundColor Red
            Read-Host "Appuyez sur Entree pour quitter"
            exit
        }
    }

    # Installer les dépendances requises
    if (Test-Path $RequirementsFile) {
        Write-Host "Installation des dépendances Python nécessaires..." -ForegroundColor Yellow
        try {
            Invoke-Expression "python -m pip install -r `"$RequirementsFile`""
            Write-Host "Toutes les dépendances ont été installées avec succès" -ForegroundColor Green
        } catch {
            Write-Host "Erreur lors de l'installation des dépendances" -ForegroundColor Red
            Write-Host $_.Exception.Message
            $installManually = Read-Host "Voulez-vous continuer sans les dépendances? (O/N)"
            if ($installManually -ne "O") {
                exit
            }
        }
    } else {
        Write-Host "Fichier requirements.txt non trouvé dans $RequirementsFile" -ForegroundColor Red
    }
}

# --- Check for required files and directories ---
if (-not (Test-Path $BaseProduitsFile)) {
    Write-Host "Erreur : Le fichier '$BaseProduitsFile' est manquant." -ForegroundColor Red
    Read-Host "Appuyez sur Entree pour quitter."
    exit
}

if (-not (Test-Path $FicheTechniqueDir -PathType Container)) {
    Write-Host "Erreur : Le dossier '$FicheTechniqueDir' est manquant." -ForegroundColor Red
    Read-Host "Appuyez sur Entree pour quitter."
    exit
}

# --- Vérifier l'installation de Python et des dépendances ---
Write-Host "Vérification de Python et des dépendances requises..." -ForegroundColor Cyan
CheckPythonDependencies

# --- Main Menu Loop ---
while ($true) {
    Clear-Host
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "   Menu de Traitement des Fichiers PDF   " -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host
    Write-Host "1. Rechercher les codes EAN ou code produit fournisseur dans les fichiers PDF dans dossier Fiche technique a traite" -ForegroundColor Yellow
    Write-Host "2. Lancer le match du ean_codes.csv avec le FICHIER GENERAL.xlsx pour renommer les fichier PDF" -ForegroundColor Yellow
    Write-Host "(ne pas laisser ouvert le fichier ean_codes.csv)"
    Write-Host "3. Extraire les images des fichiers PDF" -ForegroundColor Yellow
    Write-Host "4. Quitter" -ForegroundColor Yellow
    Write-Host

    $choice = Read-Host "Faites votre choix"

    switch ($choice) {
        "1" {
            Write-Host "Lancement du script d'extraction EAN..." -ForegroundColor Green
            python $ExtractEanScript
            Write-Host "`nOperation terminee." -ForegroundColor Green
            Write-Host "Veuillez corriger et completer le fichier 'ean_codes.csv' avant de lancer l'etape 2." -ForegroundColor Magenta
        }
        "2" {
            Write-Host "Lancement du script de renommage de fichiers..." -ForegroundColor Green
            python $RenameFilesScript
            Write-Host "`nOperation terminee." -ForegroundColor Green
        }
        "3" {
            Write-Host "Lancement du script d'extraction d'images..." -ForegroundColor Green
            python $ExtractImagesScript
            Write-Host "`nOperation terminee." -ForegroundColor Green
        }
        "4" {
            Write-Host "Au revoir !"
            exit
        }
        default {
            Write-Host "Choix invalide. Veuillez reessayer." -ForegroundColor Red
        }
    }

    Read-Host "`nAppuyez sur Entree pour retourner au menu..."
}
