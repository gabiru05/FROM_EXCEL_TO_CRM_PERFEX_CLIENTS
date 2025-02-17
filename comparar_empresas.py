import pandas as pd
import os
import argparse
from glob import glob
import re
from fuzzywuzzy import fuzz
from unidecode import unidecode

def normalize_string(text):
    """Normaliza una cadena: minúsculas, sin tildes, solo alfanumérico."""
    if not isinstance(text, str):
        return ""
    text = unidecode(text).lower()
    text = re.sub(r'[^a-z0-9\s]', '', text)
    return text.strip()

def compare_companies(target_company, other_company):
    """Compara dos nombres de empresa usando fuzzy matching."""
    if not isinstance(target_company, str) or not isinstance(other_company, str):
        return 0
    return fuzz.ratio(normalize_string(target_company), normalize_string(other_company))

def process_and_compare(input_file, compare_dir, output_file):
    """
    Procesa el archivo principal y lo compara con archivos en un directorio.
    """
    try:
        # Intenta leer como Excel primero
        df_target = pd.read_excel(input_file, engine='openpyxl')
    except Exception:
        try:
            # Intenta leer como CSV con UTF-8
            df_target = pd.read_csv(input_file, sep='\t', encoding='utf-8')
        except Exception:
            try:
                #Si UTF8 falla intenta con latin-1
                df_target = pd.read_csv(input_file, sep='\t', encoding='latin-1')
            except:
                print(f"Error: No se pudo leer el archivo principal '{input_file}'.")
                return

    # --- Detección automática de la columna 'Company' ---
    company_col = None
    for col in df_target.columns:
        if normalize_string(col) == 'company':  # Compara normalizado
            company_col = col
            break  # Sale del bucle si encuentra la columna

    if company_col is None:
        print("Error: El archivo principal no tiene una columna que se llame 'Company' (o similar).")
        return
    #-------------------------------------------------------

    df_target['Company_Normalized'] = df_target[company_col].apply(normalize_string)

    all_results = []

    compare_files = glob(os.path.join(compare_dir, "*.xlsx")) + \
                    glob(os.path.join(compare_dir, "*.xls")) + \
                    glob(os.path.join(compare_dir, "*.csv"))

    for compare_file in compare_files:
        if os.path.basename(compare_file).startswith("~$"):
            continue

        try:
            df_other = pd.read_excel(compare_file, engine='openpyxl')
        except Exception:
            try:
                df_other = pd.read_csv(compare_file, sep='\t', encoding='utf-8')
            except Exception:
                try:
                    df_other = pd.read_csv(compare_file, sep='\t', encoding='latin-1')
                except:
                    print(f"Error: No se pudo leer '{compare_file}'. Se omite.")
                    continue

        # --- Detección automática de 'Nombre_Comercial' ---
        nombre_comercial_col = None
        for col in df_other.columns:
            if normalize_string(col) == 'nombrecomercial':
                nombre_comercial_col = col
                break

        if nombre_comercial_col is None:
            print(f"Advertencia: '{compare_file}' no tiene 'Nombre_Comercial'. Se omite.")
            continue
       # ------------------------------------------------------
        df_other['Nombre_Comercial_Normalized'] = df_other[nombre_comercial_col].apply(normalize_string)

        for index, row in df_target.iterrows():
            target_company = row['Company_Normalized']
            for other_index, other_row in df_other.iterrows():
                other_company = other_row['Nombre_Comercial_Normalized']
                similarity = compare_companies(target_company, other_company)
                if similarity >= 80:
                    telefono1 = other_row.get('Telefono', '')
                    telefono2 = other_row.get('Telefono.1', '') if 'Telefono.1' in other_row else ''
                    email = other_row.get('Email', '')

                    all_results.append({
                        'Empresa_Original': row[company_col],  # Usa la columna original
                        'Empresa_Coincidente': other_row[nombre_comercial_col],  # Usa la columna original
                        'Telefono1': telefono1,
                        'Telefono2': telefono2,
                        'Email': email,
                        'Archivo_Origen': os.path.basename(compare_file),
                        'Score_Coincidencia': similarity
                    })

    if all_results:
        df_results = pd.DataFrame(all_results)
        df_results = df_results[['Empresa_Original', 'Empresa_Coincidente', 'Telefono1', 'Telefono2', 'Email', 'Archivo_Origen', 'Score_Coincidencia']]
        df_results.to_excel(output_file, index=False)
        print(f"Resultados guardados en '{output_file}'")
    else:
        print("No se encontraron coincidencias.")



def main():
    parser = argparse.ArgumentParser(description="Compara nombres de empresas.")
    parser.add_argument("input_dir", help="Directorio del archivo principal.")
    parser.add_argument("compare_dir", help="Directorio de archivos a comparar.")
    parser.add_argument("-o", "--output_file", default="resultados.xlsx",
                        help="Archivo de salida (por defecto: resultados.xlsx).")
    args = parser.parse_args()

    input_files = glob(os.path.join(args.input_dir, "*.xlsx")) + \
                 glob(os.path.join(args.input_dir, "*.xls")) + \
                 glob(os.path.join(args.input_dir, "*.csv"))

    if not input_files:
        print(f"No hay archivos en input: {args.input_dir}")
        return

    input_file = input_files[0]
    process_and_compare(input_file, args.compare_dir, args.output_file)

if __name__ == "__main__":
    main()