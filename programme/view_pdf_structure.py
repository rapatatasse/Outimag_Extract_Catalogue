import os
import sys
import PyPDF2
from pathlib import Path

def view_pdf_structure(pdf_path):
    """
    Affiche la structure d'un fichier PDF et ses métadonnées
    """
    if not os.path.exists(pdf_path):
        print(f"Erreur: Le fichier '{pdf_path}' n'existe pas.")
        return
    
    try:
        # Ouvrir le fichier PDF
        with open(pdf_path, 'rb') as file:
            # Créer un lecteur PDF
            pdf_reader = PyPDF2.PdfReader(file)
            
            # Afficher les informations de base
            print(f"\n{'='*50}")
            print(f"STRUCTURE DU FICHIER PDF: {os.path.basename(pdf_path)}")
            print(f"{'='*50}")
            
            # Nombre de pages
            num_pages = len(pdf_reader.pages)
            print(f"\nNombre de pages: {num_pages}")
            
            # Métadonnées
            print("\nMétadonnées:")
            metadata = pdf_reader.metadata
            if metadata:
                for key, value in metadata.items():
                    if value:
                        print(f"  {key}: {value}")
            else:
                print("  Aucune métadonnée trouvée")
            
            # Afficher la structure des pages
            print("\nStructure des pages:")
            for i in range(num_pages):
                page = pdf_reader.pages[i]
                print(f"\nPage {i+1}:")
                print(f"  Dimensions: {page.mediabox.width} x {page.mediabox.height} points")
                
                # Extraire le texte (les 100 premiers caractères)
                text = page.extract_text()
                if text:
                    preview = text[:100000] + "..." if len(text) > 100000 else text
                    print(f"  Aperçu du texte: {preview}")
                else:
                    print("  Aucun texte extractible")
                
                # Vérifier s'il y a des images
                if "/XObject" in page:
                    xobjects = page["/XObject"].get_object()
                    if xobjects:
                        image_count = sum(1 for obj in xobjects.values() if "/Subtype" in obj and obj["/Subtype"] == "/Image")
                        print(f"  Nombre d'images: {image_count}")
                    else:
                        print("  Aucune image détectée")
                else:
                    print("  Aucune image détectée")
            
            print(f"\n{'='*50}")
            
    except Exception as e:
        print(f"Erreur lors de la lecture du PDF: {e}")
        import traceback
        traceback.print_exc()

def main():
    # Vérifier si un chemin de fichier a été fourni
    if len(sys.argv) < 2:
        print("Usage: python view_pdf_structure.py <chemin_du_pdf>")
        print("\nOu entrez le chemin du fichier PDF manuellement:")
        pdf_path = input("Chemin du fichier PDF: ").strip()
        if not pdf_path:
            print("Aucun fichier spécifié. Fin du programme.")
            return
    else:
        pdf_path = sys.argv[1]
    
    # Convertir en chemin absolu si nécessaire
    pdf_path = os.path.abspath(pdf_path)
    
    # Afficher la structure du PDF
    view_pdf_structure(pdf_path)

if __name__ == "__main__":
    main()
