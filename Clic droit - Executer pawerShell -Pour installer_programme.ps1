# PowerShell Script pour installer le programme et ses dépendances

# Get the directory of the script
$ScriptDir = $PSScriptRoot

# --- Configuration ---
# Chemin du fichier requirements.txt
$RequirementsFile = Join-Path $ScriptDir "programme/requirements.txt"
$MainScript = Join-Path $ScriptDir "1-Clic droit Executer PowerShell.ps1"

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

# --- Main Menu Loop ---
while ($true) {
    Clear-Host
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "   Menu Principal   " -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host
    Write-Host "1. Installer le programme" -ForegroundColor Yellow
    Write-Host "2. Quitter" -ForegroundColor Yellow
    Write-Host
    
    $mainChoice = Read-Host "Faites votre choix"
    
    switch ($mainChoice) {
        "1" {
            # Installation du programme
            Write-Host "Installation du programme et des dépendances..." -ForegroundColor Cyan
            CheckPythonDependencies
            
            # Vérifier que les dossiers nécessaires existent
            $FicheTechniqueDir = Join-Path $ScriptDir "A Fiches techniques a traiter"
            $FichesTraiteesDir = Join-Path $ScriptDir "B Fiches techniques traitees"
            
            if (-not (Test-Path $FicheTechniqueDir -PathType Container)) {
                Write-Host "Création du dossier '$FicheTechniqueDir'..." -ForegroundColor Yellow
                New-Item -Path $FicheTechniqueDir -ItemType Directory | Out-Null
            }
            
            if (-not (Test-Path $FichesTraiteesDir -PathType Container)) {
                Write-Host "Création du dossier '$FichesTraiteesDir'..." -ForegroundColor Yellow
                New-Item -Path $FichesTraiteesDir -ItemType Directory | Out-Null
            }
            
            Write-Host "`nInstallation terminée." -ForegroundColor Green
            Read-Host "Appuyez sur Entree pour retourner au menu principal..."
        }
        "2" {
            Write-Host "Au revoir !" -ForegroundColor Cyan
            exit
        }
        default {
            Write-Host "Choix invalide. Veuillez reessayer." -ForegroundColor Red
            Read-Host "Appuyez sur Entree pour continuer..."
        }
    }
}
