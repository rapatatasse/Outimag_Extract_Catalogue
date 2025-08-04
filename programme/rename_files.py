import os
import pandas as pd
import re
import shutil

def sanitize_filename(filename):
    return re.sub(r'[\\/*?"<>|]', "", filename)

def rename_pdfs():
    """Renames PDF files based on EAN codes and an Excel mapping file."""
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
        db_df = pd.read_excel(product_db_path, header=6, dtype=str)
        print(f"Successfully loaded '{product_db_path}'.")

        # Get column names by index to avoid issues with header names
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

        # --- Read CSV and Process Files ---
        pdf_ean_df = pd.read_csv(ean_csv_path, sep=';', dtype=str)
        print(f"\nProcessing {len(pdf_ean_df)} files from '{ean_csv_path}'.\n")

        for index, row in pdf_ean_df.iterrows():
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
            
            for code_index, code in enumerate(codes):
                new_name_base = None
                
                # Determine which mapping to use based on fournisseur
                if fournisseur == "AUTOBEST":
                    new_name_base = autobest_to_name_map.get(code)
                    lookup_type = "6-digit AUTOBEST code"
                    
                    if new_name_base:
                        matches_found = True
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
                        
                        import shutil
                        
                        if code_index == 0:
                            # Pour le premier code, renommer le fichier original
                            os.rename(current_source_path, new_path)
                            print(f"Renamed '{original_filename}' to '{os.path.basename(new_path)}'")
                            first_renamed_file = new_path  # Sauvegarde le chemin du premier fichier renommé
                            current_source_path = new_path  # Met à jour le chemin source actuel
                        else:
                            # Pour les codes suivants, copier à partir du premier fichier renommé
                            shutil.copy2(first_renamed_file, new_path)
                            print(f"Created copy as '{os.path.basename(new_path)}'")
                    else:
                        print(f"Code '{code}' from '{original_filename}': {lookup_type} not found in database.")
                    
                elif fournisseur == "LMA":
                    # Pour LMA, rechercher toutes les références avec ce code de base
                    lookup_type = "LMA base code"
                    matching_refs = lma_code_to_refs.get(code, [])
                    
                    if matching_refs:
                        matches_found = True
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
                            
                            import shutil
                            
                            if i == 0 and code_index == 0:
                                # Pour la première référence du premier code, renommer le fichier original
                                os.rename(current_source_path, new_path)
                                print(f"Renamed '{original_filename}' to '{os.path.basename(new_path)}'")
                                first_renamed_file = new_path  # Sauvegarde le chemin du premier fichier renommé
                                current_source_path = new_path  # Met à jour le chemin source actuel
                            else:
                                # Pour les autres références, copier à partir du fichier original/premier renommé
                                shutil.copy2(first_renamed_file or original_path, new_path)
                                print(f"Created copy as '{os.path.basename(new_path)}' for ref {ref_data['product_name']}")
                    else:
                        print(f"Code '{code}' from '{original_filename}': {lookup_type} not found in database.")
                
                else:
                    # Méthode standard pour les codes EAN
                    new_name_base = ean_to_name_map.get(code)
                    lookup_type = "13-digit EAN"
                    
                    if new_name_base:
                        matches_found = True
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
                        
                        import shutil
                        
                        if code_index == 0:
                            # Pour le premier code, renommer le fichier original
                            os.rename(current_source_path, new_path)
                            print(f"Renamed '{original_filename}' to '{os.path.basename(new_path)}'")
                            first_renamed_file = new_path  # Sauvegarde le chemin du premier fichier renommé
                            current_source_path = new_path  # Met à jour le chemin source actuel
                        else:
                            # Pour les codes suivants, copier à partir du premier fichier renommé
                            shutil.copy2(first_renamed_file, new_path)
                            print(f"Created copy as '{os.path.basename(new_path)}'")
                    else:
                        print(f"Code '{code}' from '{original_filename}': {lookup_type} not found in database.")
                
                # Ce bloc est géré au-dessus dans les sections spécifiques de chaque fournisseur

    except FileNotFoundError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    rename_pdfs()
    print("\nRenaming process complete.")
