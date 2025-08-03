# PowerShell Script to manage PDF processing tasks

# Get the directory of the script
$ScriptDir = $PSScriptRoot

# --- Configuration ---
$BaseProduitsFile = Join-Path $ScriptDir "FICHIER GENERAL.xlsx"
$FicheTechniqueDir = Join-Path $ScriptDir "Fiche technique"
$ExtractEanScript = Join-Path $ScriptDir "extract_ean.py"
$RenameFilesScript = Join-Path $ScriptDir "rename_files.py"
$ExtractImagesScript = Join-Path $ScriptDir "extract_images.py"

# --- Check for required files and directories ---
if (-not (Test-Path $BaseProduitsFile)) {
    Write-Host "Erreur : Le fichier '$BaseProduitsFile' est manquant." -ForegroundColor Red
    Read-Host "Appuyez sur Entrée pour quitter."
    exit
}

if (-not (Test-Path $FicheTechniqueDir -PathType Container)) {
    Write-Host "Erreur : Le dossier '$FicheTechniqueDir' est manquant." -ForegroundColor Red
    Read-Host "Appuyez sur Entrée pour quitter."
    exit
}

# --- Main Menu Loop ---
while ($true) {
    Clear-Host
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "   Menu de Traitement des Fichiers PDF   " -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host
    Write-Host "1. Rechercher les codes EAN ou code produit fournisseur dans les fichiers PDF dans dossier Fiche technique" -ForegroundColor Yellow
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
