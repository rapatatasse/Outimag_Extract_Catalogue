import os
import pandas as pd
import re
import shutil
import time
import sys

start_total = time.time()

def sanitize_filename(filename):
    return re.sub(r'[\\/*?"<>|]', "", filename)

def rename_pdfs():
    """Renames PDF files based on EAN codes and an Excel mapping file."""
    start_total = time.time()
    # --- Configuration ---
    ean_csv_path = './ean_codes.csv'
    product_db_path = './FICHIER GENERAL.xlsx'
    pdf_directory = './A Fiches techniques a traiter'
    pdf_directory_traiter = './B Fiches techniques traitees'

    # Column indices (0-based)
    product_name_col_index = 3  # Column 4 (D) for Product Name
    ean_col_index = 4           # Column 5 (E) for 13-digit EAN
    ref_fourn_col = 5           # Column 6 (F) for Ref Four
    brand_col_index = 6         # Column 7 (G) for Brand ('AUTOBEST')

    # --- File and Directory Checks ---
    if not os.path.exists(ean_csv_path):
        print(f"Error: '{ean_csv_path}' not found. Please run the EAN extraction script first.")
        return
    if not os.path.exists(product_db_path):
        print(f"Error: Product database '{product_db_path}' not found.")
        return
    if not os.path.isdir(pdf_directory):
        print(f"Error: PDF directory '{pdf_directory}' not found.")
        return
    
    # Vérifier si le dossier de destination existe, sinon le créer
    if not os.path.exists(pdf_directory_traiter):
        os.makedirs(pdf_directory_traiter)
        print(f"Created directory '{pdf_directory_traiter}'.")

    try:
        # --- Load and Prepare Product Database ---
        print(f"Chargement du fichier Excel '{product_db_path}'...")
        start_excel = time.time()
        db_df = pd.read_excel(product_db_path, header=6, dtype=str)
        excel_time = time.time() - start_excel
        print(f"Successfully loaded '{product_db_path}' en {excel_time:.2f} secondes.")
        # Get column names by index to avoid issues with header names
        start_mapping = time.time()
        product_name_col = db_df.columns[product_name_col_index]
        ean_col = db_df.columns[ean_col_index]
        ref_fourn_col = db_df.columns[ref_fourn_col]
        brand_col = db_df.columns[brand_col_index]

        # --- Create Mapping for 13-digit EANs ---
        ean_map_df = db_df.dropna(subset=[product_name_col, ean_col])
        ean_to_name_map = pd.Series(ean_map_df[product_name_col].values, index=ean_map_df[ean_col]).to_dict()
        print(f"Created {len(ean_to_name_map)} mappings for 13-digit EANs.")

        # --- Create Mapping for 6-digit AUTOBEST codes ---
        autobest_df = db_df[db_df[brand_col] == 'AUTOBEST'].dropna(subset=[product_name_col, ref_fourn_col])
        autobest_to_name_map = pd.Series(autobest_df[product_name_col].values, index=autobest_df[ref_fourn_col]).to_dict()
        print(f"Created {len(autobest_to_name_map)} mappings for 6-digit AUTOBEST codes.")

        # --- Pour LMA: extraire tous les produits avec le même code de base ---
        lma_df = db_df[db_df[brand_col] == 'LMA'].dropna(subset=[product_name_col, ref_fourn_col])
        
        # Mapper chaque code LMA de base à toutes ses références OUT correspondantes
        lma_code_to_refs = {}
        
        # Parcourir toutes les lignes LMA
        for _, row in lma_df.iterrows():
            ref_fourn = str(row[ref_fourn_col])
            product_name = row[product_name_col]
            
            # Extraire le code de base (avant espace ou tiret)
            base_code = re.split(r'[-\s]', ref_fourn)[0]
            
            # Vérifier que le code de base a au moins 4 chiffres
            if base_code.isdigit() and len(base_code) >= 4:
                if base_code not in lma_code_to_refs:
                    lma_code_to_refs[base_code] = []
                
                lma_code_to_refs[base_code].append({
                    'product_name': product_name,
                    'ref_fourn': ref_fourn
                })
        
        print(f"Created mappings for {len(lma_code_to_refs)} LMA base codes with {sum(len(refs) for refs in lma_code_to_refs.values())} total products.")
        
        mapping_time = time.time() - start_mapping
        print(f"Création des mappings terminée en {mapping_time:.2f} secondes.")

        # --- Read CSV and Process Files ---
        start_csv = time.time()
        pdf_ean_df = pd.read_csv(ean_csv_path, sep=';', dtype=str)
        csv_time = time.time() - start_csv
        print(f"\nCSV chargé en {csv_time:.2f} secondes. Processing {len(pdf_ean_df)} files from '{ean_csv_path}'.\n")

        for index, row in pdf_ean_df.iterrows():
            try:
                original_filename = row['Filename']
                codes_str = row.get('EAN Codes', '')
                fournisseur = row.get('Fournisseur', '')

                if not codes_str or pd.isna(codes_str):
                    print(f"Skipping '{original_filename}': No codes found.")
                    continue

                # Process all codes in the file
                codes = [code.strip() for code in codes_str.split(';') if code.strip()]
                matches_found = False
                original_path = os.path.join(pdf_directory, original_filename)
                
                # Skip if original file doesn't exist
                if not os.path.exists(original_path):
                    print(f"Skipping '{original_filename}': file not found.")
                    continue
                    
                # Process each code
                current_source_path = original_path  # Garde une trace du chemin source actuel
                first_renamed_file = None  # Pour stocker le premier fichier renommé
                successfully_renamed = False  # Flag pour vérifier si au moins un fichier a été renommé
                
                # Initialiser le compteur de codes trouvés
                nb_codes_trouves = 0
                created_files = []  # Liste pour stocker les chemins des fichiers créés
                
                # S'assurer que le fichier original existe toujours
                if not os.path.exists(original_path):
                    print(f"Warning: Original file '{original_filename}' no longer exists. Skipping.")
                    continue
                    
                for code_index, code in enumerate(codes):
                    try:
                        new_name_base = None
                        
                        # Determine which mapping to use based on fournisseur
                        if fournisseur == "AUTOBEST":
                            new_name_base = autobest_to_name_map.get(code)
                            lookup_type = "6-digit AUTOBEST code"
                            
                            if new_name_base:
                                sanitized_name = sanitize_filename(str(new_name_base))
                                _, extension = os.path.splitext(original_filename)
                                new_filename = f"{sanitized_name}{extension}"
                                
                                new_path = os.path.join(pdf_directory_traiter, new_filename)
                                
                                # Handle duplicate filenames by adding a counter if needed
                                counter = 1
                                base_path = new_path
                                while os.path.exists(new_path):
                                    name_parts = os.path.splitext(base_path)
                                    new_path = f"{name_parts[0]}_{counter}{name_parts[1]}"
                                    counter += 1
                                
                                # Copier le fichier original vers la nouvelle destination
                                shutil.copy2(original_path, new_path)
                                print(f"Created file for code '{code}' as '{os.path.basename(new_path)}'")
                                created_files.append(new_path)
                                nb_codes_trouves += 1
                            else:
                                print(f"Code '{code}' from '{original_filename}': {lookup_type} not found in database.")
                            
                        elif fournisseur == "LMA":
                            # Pour LMA, rechercher toutes les références avec ce code de base
                            lookup_type = "LMA base code"
                            matching_refs = lma_code_to_refs.get(code, [])
                            
                            if matching_refs:
                                print(f"Found {len(matching_refs)} products for LMA base code {code}")
                                
                                # Pour chaque référence trouvée, créer un fichier
                                for i, ref_data in enumerate(matching_refs):
                                    _, extension = os.path.splitext(original_filename)
                                    new_filename = f"{ref_data['product_name']}{extension}"
                                    
                                    new_path = os.path.join(pdf_directory_traiter, new_filename)
                                    
                                    # Handle duplicate filenames if needed
                                    counter = 1
                                    base_path = new_path
                                    while os.path.exists(new_path):
                                        name_parts = os.path.splitext(base_path)
                                        new_path = f"{name_parts[0]}_{counter}{name_parts[1]}"
                                        counter += 1
                                    
                                    # Copier le fichier original vers la nouvelle destination
                                    shutil.copy2(original_path, new_path)
                                    print(f"Created file for LMA product '{ref_data['product_name']}' as '{os.path.basename(new_path)}'")
                                    created_files.append(new_path)
                                    nb_codes_trouves += 1
                            else:
                                print(f"Code '{code}' from '{original_filename}': {lookup_type} not found in database.")
                        
                        else:
                            # Méthode standard pour les codes EAN
                            new_name_base = ean_to_name_map.get(code)
                            lookup_type = "13-digit EAN"
                            
                            if new_name_base:
                                sanitized_name = sanitize_filename(str(new_name_base))
                                _, extension = os.path.splitext(original_filename)
                                new_filename = f"{sanitized_name}{extension}"
                                
                                new_path = os.path.join(pdf_directory_traiter, new_filename)
                                
                                # Handle duplicate filenames by adding a counter if needed
                                counter = 1
                                base_path = new_path
                                while os.path.exists(new_path):
                                    name_parts = os.path.splitext(base_path)
                                    new_path = f"{name_parts[0]}_{counter}{name_parts[1]}"
                                    counter += 1
                                
                                # Copier le fichier original vers la nouvelle destination
                                shutil.copy2(original_path, new_path)
                                print(f"Created file for code '{code}' as '{os.path.basename(new_path)}'")
                                created_files.append(new_path)
                                nb_codes_trouves += 1
                            else:
                                print(f"Code '{code}' from '{original_filename}': {lookup_type} not found in database.")
                    except Exception as e:
                        print(f"Error processing code '{code}' from '{original_filename}': {e}")
                        continue  # Continue with next code
                
                # Après avoir traité tous les codes, supprimer l'original si au moins un fichier a été créé
                if nb_codes_trouves > 0:
                    # Supprimer le fichier original car au moins une copie a été créée
                    os.remove(original_path)
                    print(f"Fichier original '{original_filename}' supprimé après création de {nb_codes_trouves} fichier(s).")
                else:
                    print(f"No valid codes found for '{original_filename}'. File not renamed.")
                
            except Exception as e:
                print(f"Error processing file entry {index}: {e}")
                continue  # Continue with next file

    except FileNotFoundError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    rename_pdfs()
    total_time = time.time() - start_total
    print(f"\nRenaming process complete en {total_time:.2f} secondes.")
