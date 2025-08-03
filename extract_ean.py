import os
import re
import fitz  # PyMuPDF
import csv

def extract_ean_from_pdf(pdf_path):
    """Extracts all 13-digit numbers (EAN codes) from a PDF file.
    Also checks for specific text "www.lma-lebeurre.com" if no EAN codes found.
    For LMA files, extracts the product code after "WORKWEAR 1880"."""
    ean_codes = []
    contains_lma = False
    lma_product_code = None
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
        
        # Si c'est un fichier LMA, chercher le code produit après "WORKWEAR 1880"
        if contains_lma:
            # Rechercher le motif "WORKWEAR 1880" suivi d'un nombre à 4+ chiffres
            match = re.search(r'WORKWEAR\s+1880\s+(\d{4,})', full_text)
            if match:
                lma_product_code = match.group(1)
                print(f"Code LMA trouvé: {lma_product_code}")
        
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
    
    return unique_codes, fournisseur

def main():
    """Main function to process all PDFs in a directory and write to a CSV."""
    pdf_directory = os.path.join(os.path.dirname(__file__), 'Fiche technique')
    output_csv_path = os.path.join(os.path.dirname(__file__), 'ean_codes.csv')

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
                # Check if the filename is exactly a 6-digit or 5-digit number followed by .pdf
                match = re.match(r'^(\d{6})\.pdf$', filename, re.IGNORECASE) #exactement 6 digits
                match2 = re.match(r'^(\d{5})\.pdf$', filename, re.IGNORECASE) #exactement 5 digits
                
                if match:
                    # If a 6-digit number is found, use it as the code
                    ean_codes = [match.group(1)]
                    print(f"--- Found 6-digit code in filename: {filename} ---")
                    fournisseur = "AUTOBEST"
                elif match2:
                    ean_codes = [match2.group(1)]
                    print(f"--- Found 5-digit code in filename: {filename} ---")
                    fournisseur = "AUTOBEST"
                else:
                    # Otherwise, search for 13-digit EANs in the PDF content
                    print(f"--- Searching for 13-digit EANs in: {filename} ---")
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
