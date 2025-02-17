import pandas as pd
import os
import argparse
from glob import glob
import re
from fuzzywuzzy import fuzz
from unidecode import unidecode
import zipfile  # <--- AGREGA ESTA LÍNEA

def normalize_string(text):
    """Normaliza una cadena: minúsculas, sin tildes, solo alfanumérico."""
    if not isinstance(text, str):
        return ""
    text = unidecode(text).lower()  # Elimina tildes y convierte a minúsculas
    text = re.sub(r'[^a-z0-9\s]', '', text)  # Elimina caracteres no alfanuméricos
    return text.strip()


def find_matching_columns(df_target, df_other):
    """Encuentra las columnas (o pares) con nombres/apellidos coincidentes."""
    matching_cols = []
    target_names = []

    # Prepara los nombres del archivo target (input)
    for col in ['Firstname', 'Lastname']:  # Considera Firstname y Lastname
        if col in df_target.columns:
            target_names.extend(df_target[col].dropna().astype(str).tolist())
    # Agrega una columna combinada "Nombre Completo" si existen ambas
    if 'Firstname' in df_target.columns and 'Lastname' in df_target.columns:
        df_target['Nombre Completo'] = df_target['Firstname'].fillna('') + " " + df_target['Lastname'].fillna('')
        target_names.extend(df_target['Nombre Completo'].dropna().astype(str).tolist())

    target_names = [normalize_string(name) for name in target_names]

    # Itera sobre columnas y pares de columnas del otro archivo
    for i in range(len(df_other.columns)):
        col1_name = df_other.columns[i]
        col1_values = df_other.iloc[:, i].dropna().astype(str).tolist()  # .iloc for positional
        col1_values_normalized = [normalize_string(val) for val in col1_values]

        # Compara columna individual
        if any(any(fuzz.partial_ratio(target_name, col_value) >= 80  # Ajusta el umbral según sea necesario
                   for col_value in col1_values_normalized)
               for target_name in target_names):
            matching_cols.append((col1_name,))  # Tupla de un elemento

        # Compara pares de columnas
        for j in range(i + 1, len(df_other.columns)):
            col2_name = df_other.columns[j]
            col2_values = df_other.iloc[:, j].dropna().astype(str).tolist()

            # Combina las dos columnas
            combined_values = (df_other.iloc[:, i].fillna('').astype(str) + " " +
                               df_other.iloc[:, j].fillna('').astype(str)).tolist()
            combined_values_normalized = [normalize_string(val) for val in combined_values]


            if any(any(fuzz.partial_ratio(target_name, combined_value) >= 80
                       for combined_value in combined_values_normalized)
                   for target_name in target_names):
                matching_cols.append((col1_name, col2_name))  # Tupla de dos elementos

    return matching_cols


def extract_info(df_target, df_other, matching_cols, filename):
    """Extrae información (correo, teléfono) basada en coincidencias de nombre."""
    results = []

    # Normaliza los nombres objetivo una vez, fuera del bucle
    target_names = []
    for col in ['Firstname', 'Lastname']:
        if col in df_target.columns:
            target_names.extend(df_target[col].dropna().astype(str).tolist())
    if 'Firstname' in df_target.columns and 'Lastname' in df_target.columns:
        df_target['Nombre Completo'] = df_target['Firstname'].fillna('') + " " + df_target['Lastname'].fillna('')
        target_names.extend(df_target['Nombre Completo'].dropna().astype(str).tolist())
    target_names_normalized = [normalize_string(name) for name in target_names]


    for match in matching_cols:
        if len(match) == 1:
            # Coincidencia con una sola columna
            col_name = match[0]
            other_names = df_other[col_name].dropna().astype(str).tolist()
            other_names_normalized = [normalize_string(name) for name in other_names]

            for target_name, target_name_norm in zip(target_names, target_names_normalized):
                for other_name, other_name_norm, (index, row) in zip(other_names, other_names_normalized, df_other.iterrows()):
                    if fuzz.partial_ratio(target_name_norm, other_name_norm) >= 80:

                        # Extrae info (maneja si las columnas no existen)
                        email = row.get('Email', '')  # Usa .get() para evitar KeyError
                        phone = row.get('Contact phonenumber',
                                        row.get('Phonenumber', '') if 'Phonenumber' in df_other.columns else '')

                        results.append({
                            'Nombre_Archivo_Target': input_file, #Nombre archivo target
                            'Nombre_Objetivo': target_name,
                            'Nombre_Coincidente': other_name,
                            'Email': email,
                            'Telefono': phone,
                            'Nombre_Archivo': filename
                        })
        else:
             # Coincidencia con par de columnas
            col1_name, col2_name = match
            combined_names = (df_other[col1_name].fillna('') + " " + df_other[col2_name].fillna('')).astype(str).tolist()
            combined_names_normalized = [normalize_string(name) for name in combined_names]

            for target_name, target_name_norm in zip(target_names, target_names_normalized):
                for combined_name, combined_name_norm, (index, row) in zip(combined_names, combined_names_normalized, df_other.iterrows()):
                    if fuzz.partial_ratio(target_name_norm, combined_name_norm) >= 80:
                        email = row.get('Email', '')
                        phone = row.get('Contact phonenumber',
                                        row.get('Phonenumber', '') if 'Phonenumber' in df_other.columns else '')
                        results.append({
                            'Nombre_Archivo_Target': input_file,
                            'Nombre_Objetivo': target_name,
                            'Nombre_Coincidente': combined_name,
                            'Email': email,
                            'Telefono': phone,
                            'Nombre_Archivo': filename
                        })

    return results


def main():
    parser = argparse.ArgumentParser(description="Compara nombres entre archivos Excel/CSV.")
    parser.add_argument("input_dir", help="Directorio que contiene el archivo principal.")
    parser.add_argument("compare_dir", help="Directorio que contiene los archivos a comparar.")
    parser.add_argument("-o", "--output_file", default="resultados_comparacion.xlsx",
                        help="Nombre del archivo de salida (por defecto: resultados_comparacion.xlsx).")
    args = parser.parse_args()

    input_files = glob(os.path.join(args.input_dir, "*.xlsx")) + \
                 glob(os.path.join(args.input_dir, "*.xls")) + \
                 glob(os.path.join(args.input_dir, "*.csv"))

    if not input_files:
      print(f"No se encontraron archivos en la carpeta input: {args.input_dir}")
      return
    #Como solo nos interesa un archivo, tomamos el primero
    input_file = input_files[0]


    try:
        # Intenta leer como Excel primero
        df_target = pd.read_excel(input_file, engine='openpyxl')
    except (FileNotFoundError, ValueError, KeyError, TypeError) as e1:
        try:
            # Si falla Excel, intenta leer como CSV con UTF-8
            df_target = pd.read_csv(input_file, sep='\t', encoding='utf-8')
        except (FileNotFoundError, pd.errors.ParserError, UnicodeDecodeError) as e2:
            try:
                # Si UTF-8 falla, intenta con latin-1
                df_target = pd.read_csv(input_file, sep='\t', encoding='latin-1')
            except (FileNotFoundError, pd.errors.ParserError, UnicodeDecodeError) as e3:
                print(f"Error: No se pudo leer el archivo '{input_file}' ni como Excel ni como CSV.")
                print(f"Errores:\nExcel: {e1}\nCSV (utf-8): {e2}\nCSV (latin-1): {e3}")
                return

    all_results = []

    compare_files = glob(os.path.join(args.compare_dir, "*.xlsx")) + \
                    glob(os.path.join(args.compare_dir, "*.xls")) + \
                    glob(os.path.join(args.compare_dir, "*.csv"))

    if not compare_files:
      print(f"No se encontraron archivos en la carpeta de comparación: {args.compare_dir}")
      return


    for compare_file in compare_files:
        # Ignora archivos temporales de Excel
        if os.path.basename(compare_file).startswith("~$"):
            continue

        try:
            # Intenta leer como Excel
            df_other = pd.read_excel(compare_file, engine='openpyxl')
        except (FileNotFoundError, ValueError, KeyError, TypeError, zipfile.BadZipFile) as e1:  # Añade zipfile.BadZipFile
            try:
                # Si falla Excel, intenta CSV con UTF-8
                df_other = pd.read_csv(compare_file, sep='\t', encoding='utf-8')
            except (FileNotFoundError, pd.errors.ParserError, UnicodeDecodeError) as e2:
                try:
                    # Si falla UTF-8, intenta latin-1
                    df_other = pd.read_csv(compare_file, sep='\t', encoding='latin-1')
                except (FileNotFoundError, pd.errors.ParserError, UnicodeDecodeError) as e3:
                    print(f"Error: No se pudo leer el archivo '{compare_file}' ni como Excel ni como CSV. Se omitirá.")
                    print(f"Errores:\nExcel: {e1}\nCSV (utf-8): {e2}\nCSV (latin-1): {e3}")
                    continue #Continua al siguiente ciclo

        filename = os.path.basename(compare_file)
        matching_cols = find_matching_columns(df_target, df_other)
        results = extract_info(df_target, df_other, matching_cols, filename)
        all_results.extend(results)

    if all_results:
        df_results = pd.DataFrame(all_results)
        try:
          df_results.to_excel(args.output_file, index=False)
          print(f"Resultados guardados en '{args.output_file}'")
        except Exception as e:
          print(f"No se pudo guardar, error: {e}")

    else:
        print("No se encontraron coincidencias.")


if __name__ == "__main__":
    main()