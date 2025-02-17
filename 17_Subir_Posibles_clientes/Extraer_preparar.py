#Extraer y preparar los posibles clientes de Pyme a subir al dashboard
#Para Directorio  Empresas Pyme

import pandas as pd
import re
import os
import argparse
import csv
from glob import glob
import chardet
from unicodedata import normalize

# --- Constantes ---
DEFAULT_ROLES_MAPPING = {
    'Dueño': ['Nombre_Propietario', 'Email']
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


def process_and_transform_excel(input_file, output_file, output_format="excel", roles_mapping=None):
    """Procesa y transforma los datos (con manejo avanzado de codificación)."""

    try:
        # Intenta leer como Excel
        df = pd.read_excel(input_file, engine='openpyxl')
    except (FileNotFoundError, ValueError, KeyError, TypeError, pd.errors.EmptyDataError) as e1:
        # Si falla Excel, intenta leer como CSV (usando la función robusta)
        df = read_csv_robust(input_file)
        if df is None:  # Si read_csv_robust retorna None, no se pudo leer
            return

    if roles_mapping is None:
        roles_mapping = DEFAULT_ROLES_MAPPING

    data = []
    for index, row in df.iterrows():
        company = row.get('Nombre_Comercial', '')
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
                phone1 = str(row.get('Telefono', '')).strip() if pd.notna(row.get('Telefono', '')) else ''
                phone2 = str(row.get('Telefono2', '')).strip() if pd.notna(row.get('Telefono2', '')) else ''
                phone1 = re.sub(r'\D', '', phone1)  # Quitar NO dígitos
                phone2 = re.sub(r'\D', '', phone2)  # Quitar NO dígitos
                phone_numbers = [p for p in [phone1, phone2] if p]
                combined_phone_number = ",".join(phone_numbers)

                actividades = row.get('Actividades', '')
                if actividades:
                    actividades = str(actividades)
                    tags_list = re.split(r'[;,]| y ', actividades)
                    tags_list = [tag.strip().lower() for tag in tags_list if tag.strip()]
                    tags = ",".join(tags_list)
                else:
                    tags = ""
                # --- Construcción de la dirección (CORREGIDA) ---
                address_parts = [
                    str(row.get('Provincia', '')).strip(),
                    str(row.get('Distrito', '')).strip(),
                    str(row.get('Corregimiento', '')).strip(),
                    str(row.get('Urbanizacion', '')).strip(),
                    str(row.get('Descripcion_Del_Area', '')).strip(),
                    str(row.get('Calle', '')).strip(),
                    str(row.get('Casa', '')).strip() if pd.notna(row.get('Casa', '')) else '',
                    str(row.get('Edificio', '')).strip(),
                    str(row.get('Apartamento', '')).strip(),
                ]
                # Filtra las partes que NO estén vacías Y que NO sean "nan"
                address = ", ".join(part for part in address_parts if part and part.lower() != "nan")

                person_data = {
                    'Name': f"{firstname} {lastname}".strip(),
                    'Position': "Dueño",
                    'Company': company,
                    'Description': '',
                    'Country': 'Panama',
                    'Zip': '',
                    'City': row.get('Distrito', ''),
                    'State': row.get('Provincia', ''),
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
    if not output_df.empty:
        column_order = [
            'Name', 'Position', 'Company', 'Description', 'Country', 'Zip',
            'City', 'State', 'Address', 'Status', 'Source', 'Email',
            'Website', 'Phonenumber', 'Lead value', 'Tags'
        ]
        output_df = output_df[column_order]

        if output_format == "csv":
            try:
                with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                    writer.writerow(output_df.columns)
                    for row in output_df.values:
                         #Aplica la función fix_encoding_issues a cada elemento de la fila
                        fixed_row = [fix_encoding_issues(str(x)) for x in row]
                        writer.writerow(fixed_row)
                print(f"Datos guardados en '{output_file}'")
            except Exception as e:
                print(f"Error al guardar '{output_file}': {e}")
        else:
            try:
                output_df.to_excel(output_file, index=False, engine='openpyxl')
                print(f"Datos guardados en '{output_file}'")
            except Exception as e:
                print(f"Error al guardar '{output_file}': {e}")
    else:
        print(f"No hay datos para procesar en '{input_file}', no se genera archivo.")

def main():
    """Función principal (sin cambios mayores)."""
    parser = argparse.ArgumentParser(description="Procesa archivos Excel/CSV y los consolida en un solo archivo.")
    parser.add_argument("input_dir", help="Directorio de entrada.")
    parser.add_argument("-o", "--output_file",
                        help="Archivo de salida (nombre completo con extensión .xlsx o .csv).")
    parser.add_argument("-f", "--format", choices=["excel", "csv"], default="excel",
                        help="Formato de salida ('excel' o 'csv', por defecto: 'excel').")

    args = parser.parse_args()

    if args.output_file:
        output_file = args.output_file
        if args.format == "excel" and not output_file.lower().endswith(".xlsx"):
            output_file += ".xlsx"
        elif args.format == "csv" and not output_file.lower().endswith(".csv"):
            output_file += ".csv"
    else:
        output_dir = args.input_dir
        if args.format == "excel":
            output_file = os.path.join(output_dir, "consolidado.xlsx")
        else:
            output_file = os.path.join(output_dir, "consolidado.csv")


    input_files = glob(os.path.join(args.input_dir, "*"))
    filtered_input_files = [f for f in input_files if not os.path.basename(f).startswith("~$") and os.path.isfile(f)]

    if not filtered_input_files:
        print(f"No se encontraron archivos válidos en: {args.input_dir}")
        return

    for input_file in filtered_input_files:
        print(f"Procesando archivo: {input_file}")
        process_and_transform_excel(input_file, output_file, args.format)

if __name__ == "__main__":
    main()