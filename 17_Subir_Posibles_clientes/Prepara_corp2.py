import pandas as pd
import re
import os
import argparse
import csv
from glob import glob
import chardet
from unicodedata import normalize
import numpy as np
# --- Constantes ---
DEFAULT_ROLES_MAPPING = {
    'Dueño': ['RAZON_SOCIAL', 'EMAIL']  # Usamos RAZON_SOCIAL como nombre
}

def split_name(full_name):
    """Divide un nombre completo en nombre y apellido (sin cambios)."""
    if not isinstance(full_name, str):
        return "", ""
    parts = full_name.split()
    if len(parts) == 0:
        return "", ""
    elif len(parts) == 1:
        return parts[0], ""
    elif len(parts) == 2:
        return parts[0], parts[1]
    else:
        firstname = []
        lastname = []
        compound_names = ["María", "Ana", "Juan", "Luis", "José", "Carlos",
                          "San", "Santa", "De", "Del", "La", "El", "Los",
                          "Da", "Do", "Das", "Dos", "D'", "L'", "O'"]

        i = 0
        while i < len(parts):
            if i < len(parts) - 1 and parts[i] in compound_names:
                firstname.append(parts[i] + " " + parts[i + 1])
                i += 2
            else:
                firstname.append(parts[i])
                i += 1

        if len(firstname) >= 3:
            mid = len(firstname) // 2
            lastname = firstname[mid:]
            firstname = firstname[:mid]

        return " ".join(firstname), " ".join(lastname)

def detect_encoding(filepath):
    """Detecta la codificación de un archivo usando chardet."""
    with open(filepath, 'rb') as file:
        rawdata = file.read()
    result = chardet.detect(rawdata)
    return result['encoding']


def detect_delimiter(filepath, num_lines=5):
    """Detecta el delimitador de un archivo CSV (mejorado con encoding)."""
    encoding = detect_encoding(filepath)
    try:
        with open(filepath, 'r', encoding=encoding) as file:
            sample_lines = [file.readline() for _ in range(num_lines)]
    except:
        return '\t'

    sniffer = csv.Sniffer()
    for line in sample_lines:
        try:
            dialect = sniffer.sniff(line)
            return dialect.delimiter
        except csv.Error:
            continue
    return '\t'


def read_csv_robust(filepath, delimiter=None):
    """Lee un CSV intentando múltiples codificaciones si es necesario."""
    if delimiter is None:
        delimiter = detect_delimiter(filepath)

    encodings_to_try = ['utf-8', 'latin-1', 'cp1252', 'utf-8-sig', 'utf-16']

    try:
        encoding = detect_encoding(filepath)
        df = pd.read_csv(filepath, sep=delimiter, encoding=encoding, errors='replace')
        return df
    except:
        pass

    for encoding in encodings_to_try:
        try:
            df = pd.read_csv(filepath, sep=delimiter, encoding=encoding, errors='replace')
            print(f"Archivo leído con éxito usando la codificación: {encoding}")
            return df
        except Exception as e:
            print(f"Fallo al leer con {encoding}: {e}")
            continue  # Prueba la siguiente codificación

    print(f"No se pudo leer el archivo CSV: {filepath}")
    return None  # Retorna None si falla

def fix_encoding_issues(text):
    """Intenta corregir problemas comunes de codificación."""
    if isinstance(text, str):
        # Diccionario de reemplazos
        replacements = {
            "Ã³": "ó",
            "Ã¡": "á",
            "Ã©": "é",
            "Ã­": "í",
            "Ãº": "ú",
            "Ã±": "ñ",
            "Ã": "Á",
            "Ã‰": "É",
            "Ã": "Í",
            "Ã“": "Ó",
            "Ãš": "Ú",
            "Ã‘": "Ñ",
            "Â¿": "¿",
            "Â¡": "¡",
            "â€œ": "“",
            "â€": "”",
            "â€”": "—",
            "â€“": "–",
            "â€¦": "…",
            "Â": "",  # Remover Â (carácter de control)
        }
        #Aplica los remplazos
        for incorrect, correct in replacements.items():
            text = text.replace(incorrect, correct)

        return text
    else:
        return text #Retorna sin modificar

def process_and_transform_excel(input_file, output_dir, output_format="excel", roles_mapping=None, chunksize=None):
    """Procesa, transforma y (opcionalmente) divide los datos."""

    try:
        df = pd.read_excel(input_file, engine='openpyxl')
    except (FileNotFoundError, ValueError, KeyError, TypeError, pd.errors.EmptyDataError) as e1:
        df = read_csv_robust(input_file)
        if df is None:
            return

    if roles_mapping is None:
        roles_mapping = DEFAULT_ROLES_MAPPING

    data = []
    for index, row in df.iterrows():
        company = row.get('NOMBRE_COMERCIAL', '')
        for position, cols in roles_mapping.items():
            name_col, email_col = cols
            try:
                full_name = str(row[name_col]) if pd.notna(row[name_col]) else ''
                email = row[email_col] if pd.notna(row[email_col]) else ''
            except KeyError as e:
                print(f"Advertencia: Columna '{e}' no encontrada en '{input_file}'. Omitiendo.")
                continue

            if email and full_name:
                firstname, lastname = split_name(full_name)
                phone1 = str(row.get('TELEFONO', '')).strip() if pd.notna(row.get('TELEFONO', '')) else ''
                phone2 = str(row.get('TELEFONO_2', '')).strip() if pd.notna(row.get('TELEFONO_2', '')) else ''
                phone1 = re.sub(r'\D', '', phone1)
                phone2 = re.sub(r'\D', '', phone2)
                phone_numbers = [p for p in [phone1, phone2] if p]
                combined_phone_number = ",".join(phone_numbers)

                actividades = row.get('ACTIVIDADES', '')
                if actividades:
                    actividades = str(actividades)
                    tags_list = re.split(r'[;,]| y ', actividades)
                    tags_list = [tag.strip().lower() for tag in tags_list if tag.strip()]
                    tags = ",".join(tags_list)
                    description = fix_encoding_issues(actividades.strip())
                else:
                    tags = ""
                    description = ""

                address_parts = [
                    str(row.get('PROVINCIA', '')).strip(),
                    str(row.get('DISTRITO', '')).strip(),
                    str(row.get('CORREGIMIENTO', '')).strip(),
                    str(row.get('URBANIZACION', '')).strip(),
                    str(row.get('DESCRIPCION_DEL_AREA', '')).strip(),
                    str(row.get('CALLE', '')).strip(),
                    str(row.get('CASA', '')).strip(),
                    str(row.get('EDIFICIO', '')).strip(),
                    str(row.get('APARTAMENTO', '')).strip(),
                ]
                address = ", ".join(part for part in address_parts if part and part.lower() != "nan")

                person_data = {
                    'Name': f"{firstname} {lastname}".strip(),
                    'Position': "Dueño",
                    'Company': company,
                    'Description': description,
                    'Country': 'Panama',
                    'Zip': '',
                    'City': row.get('DISTRITO', ''),
                    'State': row.get('PROVINCIA', ''),
                    'Address': address,
                    'Status': '',
                    'Source': '',
                    'Email': email,
                    'Website': '',
                    'Phonenumber': combined_phone_number,
                    'Lead value': '',
                    'Tags': tags
                }
                data.append(person_data)

    output_df = pd.DataFrame(data)
    if output_df.empty:
        print(f"No hay datos para procesar en '{input_file}', no se genera archivo.")
        return

    column_order = [
        'Name', 'Position', 'Company', 'Description', 'Country', 'Zip',
        'City', 'State', 'Address', 'Status', 'Source', 'Email',
        'Website', 'Phonenumber', 'Lead value', 'Tags'
    ]
    output_df = output_df[column_order]

    # --- División en chunks (si chunksize se proporciona) ---
    if chunksize is not None:
        output_base = os.path.join(output_dir, "separadas", f"TO_Dashboard_{os.path.splitext(os.path.basename(input_file))[0]}")
        os.makedirs(os.path.join(output_dir, "separadas"), exist_ok=True)  # Crea "separadas"

        for i, chunk in enumerate(np.array_split(output_df, len(output_df) // chunksize + 1)):
            output_filepath = f"{output_base}_{i+1}.{output_format}"
            if output_format == "csv":
                try:
                    with open(output_filepath, 'w', newline='', encoding='utf-8') as csvfile:
                        writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                        writer.writerow(output_df.columns)  # Encabezados en cada chunk
                        for row in chunk.values:
                            fixed_row = [fix_encoding_issues(str(x)) for x in row]
                            writer.writerow(fixed_row)
                    print(f"Datos guardados en '{output_filepath}'")
                except Exception as e:
                    print(f"Error al guardar '{output_filepath}': {e}")

            else: #Excel
                try:
                    chunk.to_excel(output_filepath, index=False, engine='openpyxl')
                    print(f"Datos guardados en '{output_filepath}'")
                except Exception as e:
                    print(f"Error al guardar '{output_filepath}': {e}")

    else:
        # --- Guardado normal (sin división) ---
        output_filepath = os.path.join(output_dir, f"TO_Dashboard_{os.path.basename(input_file)}.{output_format}")
        if output_format == "csv":
            try:
                with open(output_filepath, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                    writer.writerow(output_df.columns)
                    for row in output_df.values:
                        fixed_row = [fix_encoding_issues(str(x)) for x in row]
                        writer.writerow(fixed_row)
                print(f"Datos guardados en '{output_filepath}'")
            except Exception as e:
                print(f"Error al guardar '{output_filepath}': {e}")
        else:  # Excel
            try:
                output_df.to_excel(output_filepath, index=False, engine='openpyxl')
                print(f"Datos guardados en '{output_filepath}'")
            except Exception as e:
                print(f"Error al guardar '{output_filepath}': {e}")



def main():
    """Función principal."""
    parser = argparse.ArgumentParser(description="Procesa archivos Excel/CSV, los consolida y opcionalmente los divide.")
    parser.add_argument("input_dir", help="Directorio de entrada.")
    parser.add_argument("-o", "--output_dir",
                        help="Directorio de salida (si no, usa el directorio de entrada).")
    parser.add_argument("-f", "--format", choices=["excel", "csv"], default="excel",
                        help="Formato de salida ('excel' o 'csv', por defecto: 'excel').")
    parser.add_argument("-c", "--chunksize", type=int,
                        help="Tamaño de los chunks para dividir el archivo (opcional).")

    args = parser.parse_args()

    # --- Configuración del directorio de salida ---
    if args.output_dir:
        output_dir = args.output_dir
    else:
        output_dir = os.path.join(args.input_dir, "output")  # Crea carpeta "output"
    os.makedirs(output_dir, exist_ok=True)  # Crea el directorio si no existe

    input_files = glob(os.path.join(args.input_dir, "*"))
    filtered_input_files = [f for f in input_files if not os.path.basename(f).startswith("~$") and os.path.isfile(f)]

    if not filtered_input_files:
        print(f"No se encontraron archivos válidos en: {args.input_dir}")
        return

    for input_file in filtered_input_files:
        print(f"Procesando archivo: {input_file}")
        process_and_transform_excel(input_file, output_dir, args.format, chunksize=args.chunksize)

if __name__ == "__main__":
    main()