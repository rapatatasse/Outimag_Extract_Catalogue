import os
import re
import fitz  # PyMuPDF
import csv

def extract_ean_from_pdf(pdf_path):
    """Extracts all 13-digit numbers (EAN codes) from a PDF file.
    Also checks for specific text "www.lma-lebeurre.com" if no EAN codes found.
    For LMA files, extracts the product code after "WORKWEAR 1880".
    For AUTOBEST files, extracts the reference at the beginning of the text before "AUTOBEST"."""
    ean_codes = []
    contains_lma = False
    contains_autobest = False
    lma_product_code = None
    autobest_code = None
    full_text = ""
    fournisseur = ""
    
    try:
        doc = fitz.open(pdf_path)
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text()
            full_text += text
            
            # Regex to find all 13-digit numbers
            found_codes = re.findall(r'\b\d{13}\b', text)
            if found_codes:
                ean_codes.extend(found_codes)
            
            # Vérifier si le texte contient www.lma-lebeurre.com
            if "www.lma-lebeurre.com" in text:
                contains_lma = True
                
            # Vérifier si le texte contient AUTOBEST
            if "AUTOBEST" in text:
                contains_autobest = True
                       
        # Si c'est un fichier LMA, chercher le code produit après "WORKWEAR 1880"
        if contains_lma:
            # Rechercher le motif "WORKWEAR 1880" suivi d'un nombre à 4+ chiffres
            match = re.search(r'WORKWEAR\s+1880\s+(\d{4,})', full_text)
            if match:
                lma_product_code = match.group(1)
                print(f"Code LMA trouvé: {lma_product_code}")
        
        # Si c'est un fichier AUTOBEST, extraire la référence après "Aperçu du texte:"
        if contains_autobest:
                       # Rechercher le motif "Aperçu du texte:" suivi de chiffres
            # Afficher le texte autour de "Aperçu du texte:" pour débogage
            apercu_index = full_text.find("AUTOBEST - BP 67")
            if apercu_index != -1:
                debug_text = full_text[apercu_index:apercu_index+100]
               
                # Extraire les chiffres après "AUTOBEST - BP 67"
                match = re.search(r'\s*(\d+)\s*AUTOBEST - BP 67', full_text)
                if match:
                    autobest_code = match.group(1).strip()
                    print(f"Code AUTOBEST trouvé dans le contenu: {autobest_code}")
                else:
                    print("Motif 'Code AUTOBEST' non trouvé")
        
        doc.close()
    except Exception as e:
        print(f"Error processing file {os.path.basename(pdf_path)}: {e}")
    
    # Return unique EAN codes
    unique_codes = list(set(ean_codes))
    
    # Si aucun code EAN trouvé et c'est un fichier LMA, utiliser le code produit LMA s'il existe
    if not unique_codes and contains_lma:
        if lma_product_code:
            # Ajouter le code produit LMA comme "code EAN"
            unique_codes = [lma_product_code]
            fournisseur = "LMA"
        else:
            fournisseur = "LMA"
    
    # Si c'est un fichier AUTOBEST, utiliser le code AUTOBEST extrait du contenu
    if contains_autobest and autobest_code:
        unique_codes = [autobest_code]
        fournisseur = "AUTOBEST"
    
    return unique_codes, fournisseur

def main():
    """Main function to process all PDFs in a directory and write to a CSV."""
    pdf_directory = os.path.join(os.path.dirname(__file__), '../A Fiches techniques a traiter')
    output_csv_path = os.path.join(os.path.dirname(__file__), '../ean_codes.csv')

    if not os.path.isdir(pdf_directory):
        print(f"Error: Directory not found at '{pdf_directory}'")
        return

    print(f"Searching for PDF files in: {pdf_directory}\n")

    with open(output_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
        csv_writer = csv.writer(csvfile, delimiter=';')
        csv_writer.writerow(['Filename', 'EAN Codes', 'Fournisseur'])

        for filename in os.listdir(pdf_directory):
            if filename.lower().endswith('.pdf'):
                ean_codes = []
                # Analyser le contenu du PDF pour tous les fichiers
                print(f"--- Analyzing PDF content of: {filename} ---")
                pdf_path = os.path.join(pdf_directory, filename)
                ean_codes, fournisseur = extract_ean_from_pdf(pdf_path)
                
                ean_codes_str = ';'.join(ean_codes)
                csv_writer.writerow([filename, ean_codes_str, fournisseur])

                if ean_codes:
                    print(f"Found codes: {ean_codes_str}")
                else:
                    print("No codes found.")
                print("-" * (len(filename) + 20) + "\n")

    print(f"\nProcessing complete. Results saved to '{output_csv_path}'")


if __name__ == "__main__":
    main()
