import pandas as pd
import re
import os
import argparse
from glob import glob
import csv  # Importante: Importar el módulo csv


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
                        "San", "Santa", "De", "Del", "La", "El", "Los"]

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


def process_and_split_excel(input_file, output_dir, output_format="excel",
                            group_by_col='GRUPO / TALLER', roles_mapping=None):
    """
    Procesa archivos, transforma datos y divide en archivos por grupo.
    """

    try:
        # Intenta leer como Excel primero
        df = pd.read_excel(input_file, engine='openpyxl')
    except (FileNotFoundError, ValueError, KeyError, TypeError) as e1:
        try:
            # Si falla Excel, intenta leer como CSV con UTF-8 y delimitador \t
            df = pd.read_csv(input_file, sep='\t', encoding='utf-8')
        except (FileNotFoundError, pd.errors.ParserError, UnicodeDecodeError) as e2:
            try:
                # Si UTF-8 falla, intenta con latin-1 y delimitador \t
                df = pd.read_csv(input_file, sep='\t', encoding='latin-1')
            except (FileNotFoundError, pd.errors.ParserError, UnicodeDecodeError) as e3:
                print(f"Error: No se pudo leer el archivo '{input_file}' ni como Excel ni como CSV.")
                print(f"Errores:\nExcel: {e1}\nCSV (utf-8): {e2}\nCSV (latin-1): {e3}")
                return

    if roles_mapping is None:
        roles_mapping = {
            'Representante Principal': ['Representante Principal', 'Email'],
            'Representante Suplente': ['Representante Suplente', 'Email.1'],
            'Asistente de Gerencia': ['Asistente de Gerencia', 'Email.2'],
            'Gerente General': ['Gerente General', 'Email.3'],
            'Recursos Humanos': ['Recursos Humanos', 'Email.4'],
            'Mercadeo': ['Mercadeo', 'Email.5'],
            'Ventas': ['Ventas', 'Email.6'],
        }

    grouped = df.groupby(group_by_col)

    for group_name, group_df in grouped:
        data = []
        for index, row in group_df.iterrows():
            company = row['Nombre_empresa']
            phone_number = str(row['Telefonos']) if pd.notna(row['Telefonos']) else ''
            phone_number = re.sub(r'\D', '', phone_number)

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
                    person_data = {
                        'Firstname': firstname, 'Lastname': lastname, 'Email': email,
                        'Contact phonenumber': phone_number, 'Position': position,
                        'Company': company, 'Vat': '', 'Phonenumber': '', 'Country': 'Panama',
                        'City': 'SAN FELIPE', 'Zip': '', 'State': '', 'Address': '',
                        'Website': '', 'Billing street': '', 'Billing city': 'Panama',
                        'Billing state': 'SAN FELIPE', 'Billing zip': '', 'Billing country': 'Panama',
                        'Shipping street': '', 'Shipping city': '', 'Shipping state': '',
                        'Shipping zip': '', 'Shipping country': '', 'Longitude': '', 'Latitude': '',
                        'Stripe id': ''
                    }
                    data.append(person_data)

        output_df = pd.DataFrame(data)
        if not output_df.empty:
            column_order = [
                'Firstname', 'Lastname', 'Email', 'Contact phonenumber', 'Position',
                'Company', 'Vat', 'Phonenumber', 'Country', 'City', 'Zip', 'State',
                'Address', 'Website', 'Billing street', 'Billing city', 'Billing state',
                'Billing zip', 'Billing country', 'Shipping street', 'Shipping city',
                'Shipping state', 'Shipping zip', 'Shipping country', 'Longitude',
                'Latitude', 'Stripe id'
            ]
            output_df = output_df[column_order]

            try:
                output_filename = re.sub(r'[\\/*?:"<>|]', "", str(group_name))
                output_filename = output_filename.strip()
            except TypeError:
                output_filename = "grupo_invalido"

            if output_format == "csv":
                output_filepath = os.path.join(output_dir, f"{output_filename}.csv")
                try:
                    # --- CAMBIO IMPORTANTE AQUÍ: Usar csv.writer ---
                    with open(output_filepath, 'w', newline='', encoding='utf-8') as csvfile:
                        writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                        writer.writerow(output_df.columns)  # Escribe los encabezados
                        for row in output_df.values:
                            writer.writerow(row)  # Escribe cada fila
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
    """Función principal que maneja los argumentos de la línea de comandos."""

    parser = argparse.ArgumentParser(description="Procesa y divide archivos Excel/CSV por grupos.")
    parser.add_argument("input_dir", help="Directorio de entrada.")
    parser.add_argument("-o", "--output_dir",
                        help="Directorio de salida (si no, usa el de entrada).")
    parser.add_argument("-f", "--format", choices=["excel", "csv"], default="excel",
                        help="Formato de salida ('excel' o 'csv', por defecto: 'excel').")
    parser.add_argument("-g", "--group_by", default="GRUPO / TALLER",
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
                role = row['Role']
                name_col = row['NameColumn']
                email_col = row['EmailColumn']
                roles_mapping[role] = [name_col, email_col]

        except Exception as e:
            print(f"Error al leer el archivo de mapeo: {e}. Se usará el mapeo predeterminado.")

    input_files = glob(os.path.join(args.input_dir, "*.xlsx")) + \
                  glob(os.path.join(args.input_dir, "*.xls")) + \
                  glob(os.path.join(args.input_dir, "*.csv"))

    if not input_files:
        print(f"No se encontraron archivos Excel/CSV en: {args.input_dir}")
        return

    for input_file in input_files:
        print(f"Procesando archivo: {input_file}")
        process_and_split_excel(input_file, output_dir, args.format, args.group_by, roles_mapping)


if __name__ == "__main__":
    main()