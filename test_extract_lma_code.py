import os
import fitz  # PyMuPDF
import re

def extract_lma_info(pdf_path):
    """Extrait le contenu d'un PDF LMA et tente de trouver le numéro à 4-6 chiffres."""
    try:
        # Ouvrir le document PDF
        doc = fitz.open(pdf_path)
        print(f"\nAnalyse du fichier: {os.path.basename(pdf_path)}")
        print("-" * 50)
        
        # Variables pour stocker les informations
        found_code = None
        contains_lma = False
        full_text = ""
        
        # Examiner chaque page
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text()
            full_text += text
            
            # Vérifier s'il s'agit d'un fichier LMA
            if "www.lma-lebeurre.com" in text:
                contains_lma = True
        
        if contains_lma:
            print("✓ C'est bien un fichier LMA")
            
            # Afficher les 1000 premiers caractères pour analyse
            print("\nAperçu du contenu (1000 premiers caractères):")
            print(full_text[:1000])
            
            # Stratégie 1: Chercher un nombre à 4-6 chiffres dans une ligne contenant "TERREAU" ou avant "Descriptif"
            lines = full_text.split('\n')
            for line in lines:
                # Chercher un nombre à 4-6 chiffres dans la même ligne que TERREAU ou similaire
                title_match = re.search(r'(\d{4,6})\s+([A-Z]+)', line)
                if title_match:
                    found_code = title_match.group(1)
                    product_name = title_match.group(2)
                    print(f"\nStratégie 1 - Trouvé dans le titre: Code={found_code}, Nom={product_name}")
                    break
            
            # Stratégie 2: Chercher avant "Descriptif"
            if not found_code:
                descriptif_idx = full_text.find("Descriptif")
                if descriptif_idx > 0:
                    before_descriptif = full_text[:descriptif_idx]
                    code_match = re.search(r'\b(\d{4,6})\b', before_descriptif[-50:])
                    if code_match:
                        found_code = code_match.group(1)
                        print(f"\nStratégie 2 - Trouvé avant Descriptif: Code={found_code}")
            
            # Stratégie 3: Simplement chercher n'importe quel nombre à 4-6 chiffres
            if not found_code:
                all_numbers = re.findall(r'\b(\d{4,6})\b', full_text)
                if all_numbers:
                    # Prendre le premier nombre trouvé
                    found_code = all_numbers[0]
                    print(f"\nStratégie 3 - Premier nombre à 4-6 chiffres trouvé: {found_code}")
                    print(f"Tous les nombres à 4-6 chiffres trouvés: {all_numbers[:10]}")
        else:
            print("Ce n'est pas un fichier LMA")
        
        doc.close()
        return found_code, contains_lma
        
    except Exception as e:
        print(f"Erreur lors de l'analyse du fichier: {e}")
        return None, False

def main():
    """Fonction principale pour tester l'extraction sur un PDF."""
    # Répertoire des PDFs
    pdf_directory = os.path.join(os.path.dirname(__file__), 'Fiche technique')
    
    if not os.path.isdir(pdf_directory):
        print(f"Erreur: Répertoire introuvable '{pdf_directory}'")
        return
    
    # Chercher les fichiers PDF qui pourraient être des fichiers LMA
    lma_found = False
    for filename in os.listdir(pdf_directory):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(pdf_directory, filename)
            
            # Vérifier rapidement si c'est un fichier LMA
            try:
                doc = fitz.open(pdf_path)
                for page_num in range(min(len(doc), 2)):  # Vérifier seulement les 2 premières pages
                    page = doc.load_page(page_num)
                    text = page.get_text()
                    if "www.lma-lebeurre.com" in text:
                        # Si c'est un fichier LMA, faire une analyse complète
                        code, is_lma = extract_lma_info(pdf_path)
                        if code:
                            print(f"\n✓ CODE LMA EXTRAIT: {code}")
                            lma_found = True
                            break
                doc.close()
                if lma_found:
                    break
            except Exception as e:
                print(f"Erreur lors de la vérification de {filename}: {e}")
    
    if not lma_found:
        print("\nAucun fichier LMA trouvé dans le répertoire. Assurez-vous que les fichiers contiennent 'www.lma-lebeurre.com'.")

if __name__ == "__main__":
    main()
