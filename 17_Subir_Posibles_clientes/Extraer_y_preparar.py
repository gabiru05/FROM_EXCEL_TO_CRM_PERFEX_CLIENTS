#Extraer y preparar los posibles clientes de Pyme a subir al dashboard
#Para Directorio 154

import pandas as pd
import re
import os
import argparse
from glob import glob
import csv  # Importar el módulo csv

# --- Constantes ---
DEFAULT_GROUP_BY = 'GRUPO / TALLER'
DEFAULT_ROLES_MAPPING = {
    'Representante Principal': ['Representante Principal', 'Email'],
    'Representante Suplente': ['Representante Suplente', 'Email.1'],
    'Asistente de Gerencia': ['Asistente de Gerencia', 'Email.2'],
    'Gerente General': ['Gerente General', 'Email.3'],
    'Recursos Humanos': ['Recursos Humanos', 'Email.4'],
    'Mercadeo': ['Mercadeo', 'Email.5'],
    'Ventas': ['Ventas', 'Email.6'],
}

def split_name(full_name):
    """Divide un nombre completo en nombre y apellido."""
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
                          "Da", "Do", "Das", "Dos", "D'", "L'", "O'"]  # Añadidos más nombres compuestos

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


def detect_delimiter(filepath, num_lines=5):
    """Detecta el delimitador más probable de un archivo CSV."""
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            sample_lines = [file.readline() for _ in range(num_lines)]
    except UnicodeDecodeError:
        try:
            with open(filepath, 'r', encoding='latin-1') as file:
                sample_lines = [file.readline() for _ in range(num_lines)]
        except UnicodeDecodeError:
            with open(filepath, 'r', encoding='cp1252') as file: #Prueba con encoding cp1252
               sample_lines = [file.readline() for _ in range(num_lines)]

    sniffer = csv.Sniffer()
    for line in sample_lines:
        try:
            dialect = sniffer.sniff(line)
            return dialect.delimiter
        except csv.Error:
            continue
    return '\t'  # Delimitador por defecto


def process_and_split_excel(input_file, output_dir, output_format="excel",
                            group_by_col=DEFAULT_GROUP_BY, roles_mapping=None):
    """Procesa archivos, transforma datos y divide en archivos por grupo."""

    try:
        # Intenta leer como Excel primero
        df = pd.read_excel(input_file, engine='openpyxl')
    except (FileNotFoundError, ValueError, KeyError, TypeError, pd.errors.EmptyDataError) as e1:
        try:
            # Si falla Excel, intenta leer como CSV
            delimiter = detect_delimiter(input_file)  # Detecta el delimitador
            df = pd.read_csv(input_file, sep=delimiter, encoding='utf-8')
        except (FileNotFoundError, pd.errors.ParserError, UnicodeDecodeError, pd.errors.EmptyDataError) as e2:
            try:
                # Si UTF-8 falla, intenta con latin-1
                df = pd.read_csv(input_file, sep=delimiter, encoding='latin-1')
            except (FileNotFoundError, pd.errors.ParserError, UnicodeDecodeError, pd.errors.EmptyDataError) as e3:
                try:
                    # Si latin-1 falla, intenta con cp1252
                    df = pd.read_csv(input_file, sep=delimiter, encoding='cp1252')
                except (FileNotFoundError, pd.errors.ParserError, UnicodeDecodeError, pd.errors.EmptyDataError) as e4:
                    print(f"Error: No se pudo leer '{input_file}' ni como Excel ni como CSV.")
                    print(f"Errores:\nExcel: {e1}\nCSV (utf-8): {e2}\nCSV (latin-1): {e3}\nCSV(cp1252): {e4}")
                    return

    if roles_mapping is None:
        roles_mapping = DEFAULT_ROLES_MAPPING

    grouped = df.groupby(group_by_col)

    for group_name, group_df in grouped:
        data = []
        for index, row in group_df.iterrows():
            company = row['Nombre_empresa']
            phone_number = str(row['Telefonos']) if pd.notna(row['Telefonos']) else ''
            phone_number = re.sub(r'\D', '', phone_number) # Limpiar numero de telefono

            # --- Obtener 'Actividad' y 'GRUPO / TALLER' ---
            actividad = row.get('Actividad', '')  # Usar .get() por si no existe
            grupo_taller = row.get(group_by_col, '')  # Usar group_by_col y .get()

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

                    # --- Crear los tags ---
                    # Priorizar 'Actividad Comercial', luego 'Actividad', y finalmente cadena vacía
                    tags_source = row.get('Actividad Comercial', row.get('Actividad', ''))
                    if tags_source:
                        # Dividir por comas, punto y coma, o "y" (con espacios opcionales)
                        tags_list = re.split(r'[;,]| y ', tags_source)
                        tags_list = [tag.strip().lower() for tag in tags_list if tag.strip()]  # Limpiar y a minúsculas
                        tags = ",".join(tags_list)
                    else:
                        tags = ""

                    person_data = {
                        'Name': f"{firstname} {lastname}".strip(),  # Combina nombre y apellido
                        'Position': position,
                        'Company': company,
                        'Description': '',  # Valores por defecto para las nuevas columnas
                        'Country': 'Panama',
                        'Zip': '',
                        'City': '',  # Ya no se pone SAN FELIPE por defecto
                        'State': '',
                        'Address': '',
                        'Status': '',
                        'Source': '',
                        'Email': email,
                        'Website': '',
                        'Phonenumber': phone_number,
                        'Lead value': '',
                        'Tags': tags  # Agregar los tags
                    }
                    data.append(person_data)

        output_df = pd.DataFrame(data)
        if not output_df.empty:
            # --- Orden de columnas CORRECTO ---
            column_order = [
                'Name', 'Position', 'Company', 'Description', 'Country', 'Zip',
                'City', 'State', 'Address', 'Status', 'Source', 'Email',
                'Website', 'Phonenumber', 'Lead value', 'Tags'
            ]
            output_df = output_df[column_order]

            try:
                output_filename = re.sub(r'[\\/*?:"<>|]', "", str(group_name))
                output_filename = output_filename.strip() #Eliminar espacios al inicio y al final
            except TypeError:
                output_filename = "grupo_invalido"

            if output_format == "csv":
                output_filepath = os.path.join(output_dir, f"{output_filename}.csv")
                try:
                    # --- Usa csv.writer para un CSV bien formado ---
                    with open(output_filepath, 'w', newline='', encoding='utf-8') as csvfile:
                        writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                        writer.writerow(output_df.columns)  # Encabezados
                        for row in output_df.values:
                            writer.writerow(row)
                    # ------------------------------------------------
                    print(f"Datos de '{group_name}' guardados en '{output_filepath}'")
                except Exception as e:
                    print(f"Error al guardar '{output_filepath}': {e}")
            else:
                output_filepath = os.path.join(output_dir, f"{output_filename}.xlsx")
                try:
                    output_df.to_excel(output_filepath, index=False, engine='openpyxl')
                    print(f"Datos de '{group_name}' guardados en '{output_filepath}'")
                except Exception as e:
                    print(f"Error al guardar '{output_filepath}': {e}")
        else:
            print(f"No hay datos para el grupo '{group_name}' en '{input_file}', no se genera archivo.")


def main():
    """Función principal."""
    parser = argparse.ArgumentParser(description="Procesa y divide archivos Excel/CSV por grupos.")
    parser.add_argument("input_dir", help="Directorio de entrada.")
    parser.add_argument("-o", "--output_dir",
                        help="Directorio de salida (si no, usa el de entrada).")
    parser.add_argument("-f", "--format", choices=["excel", "csv"], default="excel",
                        help="Formato de salida ('excel' o 'csv', por defecto: 'excel').")
    parser.add_argument("-g", "--group_by", default=DEFAULT_GROUP_BY,
                        help="Columna para agrupar (por defecto: 'GRUPO / TALLER').")
    parser.add_argument("-m", "--mapping",
                        help="Ruta a un archivo de mapeo de columnas (opcional).")
    args = parser.parse_args()

    output_dir = args.output_dir if args.output_dir else args.input_dir
    if output_dir != args.input_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    roles_mapping = None
    if args.mapping:
        try:
            mapping_df = pd.read_csv(args.mapping)
            roles_mapping = {}
            for index, row in mapping_df.iterrows():
                roles_mapping[row['Role']] = [row['NameColumn'], row['EmailColumn']]
        except Exception as e:
            print(f"Error al leer mapeo: {e}. Usando mapeo predeterminado.")

    # --- MODIFICACIÓN AQUÍ: Simplificar la búsqueda de archivos ---
    input_files = glob(os.path.join(args.input_dir, "*"))  # Busca *cualquier* archivo

    filtered_input_files = [f for f in input_files if not os.path.basename(f).startswith("~$") and os.path.isfile(f)]


    if not filtered_input_files:
        print(f"No se encontraron archivos válidos en: {args.input_dir}")
        return
    # ----------------------------------------------------------------

    for input_file in filtered_input_files:
        print(f"Procesando archivo: {input_file}")
        process_and_split_excel(input_file, output_dir, args.format, args.group_by, roles_mapping)

if __name__ == "__main__":
    main()