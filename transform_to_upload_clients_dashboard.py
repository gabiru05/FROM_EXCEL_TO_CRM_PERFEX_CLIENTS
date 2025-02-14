import pandas as pd
import re
import os
import argparse
from glob import glob  # Importamos glob


def split_name(full_name):
    """Divide un nombre completo en nombre y apellido."""
    if not full_name:
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
    Procesa un archivo Excel o CSV, transforma los datos y los divide en archivos
    según una columna de agrupación.  Maneja errores de lectura individualmente
    por archivo.

    Args:
        input_file (str): Ruta al archivo de entrada (Excel o CSV).
        output_dir (str): Directorio donde se guardarán los archivos de salida.
        output_format (str): Formato de los archivos de salida ("excel" o "csv").
        group_by_col (str): Nombre de la columna para agrupar.
        roles_mapping (dict): Diccionario de mapeo de roles a columnas.
    """

    try:
        # Intenta leer como CSV
        df = pd.read_csv(input_file, sep='\t', encoding='latin-1')
    except (FileNotFoundError, pd.errors.ParserError, UnicodeDecodeError) as e1:
        try:
            # Si falla, intenta leer como Excel
            df = pd.read_excel(input_file, engine='openpyxl')
        except (FileNotFoundError, pd.errors.ParserError, UnicodeDecodeError) as e2:
            try:
                # Si falla, intenta leer como CSV con utf-8
                df = pd.read_csv(input_file, sep='\t', encoding='utf-8')
            except (FileNotFoundError, pd.errors.ParserError, UnicodeDecodeError) as e3:
                print(f"Error: No se pudo leer el archivo '{input_file}' ni como CSV ni como Excel.")
                print(f"Errores:\nCSV (latin-1): {e1}\nExcel: {e2}\nCSV (utf-8): {e3}")
                return  # Sale de la función si no se puede leer el archivo

    # Usa el mapeo de roles proporcionado o uno predeterminado
    if roles_mapping is None:
        roles_mapping = {  # Mapeo predeterminado
            'Representante Principal': ['Representante Principal', 'Email'],
            'Representante Suplente': ['Representante Suplente', 'Email.1'],
            'Asistente de Gerencia': ['Asistente de Gerencia', 'Email.2'],
            'Gerente General': ['Gerente General', 'Email.3'],
            'Recursos Humanos': ['Recursos Humanos', 'Email.4'],
            'Mercadeo': ['Mercadeo', 'Email.5'],
            'Ventas': ['Ventas', 'Email.6'],
        }

    # Agrupa por la columna especificada
    grouped = df.groupby(group_by_col)

    # Itera sobre cada grupo
    for group_name, group_df in grouped:
        data = []
        # Procesa cada grupo
        for index, row in group_df.iterrows():
            company = row['Nombre_empresa']
            phone_number = str(row['Telefonos']) if pd.notna(row['Telefonos']) else ''
            phone_number = re.sub(r'\D', '', phone_number)

            for position, cols in roles_mapping.items():
                name_col, email_col = cols
                # Manejo para diferentes nombres de columnas
                try:
                    full_name = str(row[name_col]) if pd.notna(row[name_col]) else ''
                    email = row[email_col] if pd.notna(row[email_col]) else ''
                except KeyError as e:
                    print(f"Advertencia: La columna '{e}' no se encontró en el archivo '{input_file}'.  Se omitirá esta columna.")
                    continue  # Continua con la siguiente iteración

                if email and full_name:
                    firstname, lastname = split_name(full_name)
                    person_data = {
                        'Firstname': firstname,
                        'Lastname': lastname,
                        'Email': email,
                        'Contact phonenumber': phone_number,
                        'Position': position,
                        'Company': company,
                        'Vat': '',
                        'Phonenumber': '',
                        'Country': 'Panama',
                        'City': '',
                        'Zip': '',
                        'State': '',
                        'Address': '',
                        'Website': '',
                        'Billing street': '',
                        'Billing city': 'Panama',
                        'Billing state': '',
                        'Billing zip': '',
                        'Billing country': 'Panama',
                        'Shipping street': '',
                        'Shipping city': '',
                        'Shipping state': '',
                        'Shipping zip': '',
                        'Shipping country': '',
                        'Longitude': '',
                        'Latitude': '',
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

            # Limpia el nombre del grupo
            output_filename = re.sub(r'[\\/*?:"<>|]', "", str(group_name))
            output_filename = output_filename.strip()
            # Determina la extension y guarda
            if output_format == "csv":
                output_filepath = os.path.join(output_dir, f"{output_filename}.csv")
                try:
                    output_df.to_csv(output_filepath, index=False)
                    print(f"Datos de '{group_name}' guardados en '{output_filepath}'")
                except Exception as e:
                    print(f"Error al guardar '{output_filepath}': {e}")

            else:  # Por defecto, guarda como Excel
                output_filepath = os.path.join(output_dir, f"{output_filename}.xlsx")
                try:
                    output_df.to_excel(output_filepath, index=False)
                    print(f"Datos de '{group_name}' guardados en '{output_filepath}'")
                except Exception as e:
                    print(f"Error al guardar '{output_filepath}': {e}")
        else:
            print(f"No hay datos para el grupo '{group_name}' en '{input_file}', no se genera archivo.")


def main():
    """Función principal que maneja los argumentos de la línea de comandos."""

    parser = argparse.ArgumentParser(description="Procesa y divide archivos Excel/CSV por grupos.")
    parser.add_argument("input_dir", help="Ruta al directorio de entrada que contiene los archivos Excel/CSV.")
    parser.add_argument("-o", "--output_dir",
                        help="Directorio de salida (si no se especifica, se usa el mismo que la entrada).")
    parser.add_argument("-f", "--format", choices=["excel", "csv"], default="excel",
                        help="Formato de salida ('excel' o 'csv', por defecto: 'excel').")
    parser.add_argument("-g", "--group_by", default="GRUPO / TALLER",
                        help="Nombre de la columna para agrupar (por defecto: 'GRUPO / TALLER').")
    parser.add_argument("-m", "--mapping",
                        help="Ruta a un archivo de mapeo de columnas (opcional).")

    args = parser.parse_args()

    # Si no se especifica directorio de salida, usa el de entrada
    output_dir = args.output_dir if args.output_dir else args.input_dir

    # Crea el directorio de salida si no existe y si es diferente al de entrada
    if output_dir != args.input_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)


    # Procesa el mapeo si fue provisto
    roles_mapping = None
    if args.mapping:
        try:
            # Intenta cargar desde un archivo .csv
            mapping_df = pd.read_csv(args.mapping)
            roles_mapping = {}
            # Itera por el mapeo
            for index, row in mapping_df.iterrows():
                role = row['Role']
                name_col = row['NameColumn']
                email_col = row['EmailColumn']
                # Guardamos en el diccionario
                roles_mapping[role] = [name_col, email_col]

        except Exception as e:
            print(f"Error al leer el archivo de mapeo: {e}. Se usará el mapeo predeterminado.")
            # Si hay error se sigue con el mapeo default

    # Usa glob para encontrar todos los archivos .xlsx, .xls y .csv en el directorio de entrada
    input_files = glob(os.path.join(args.input_dir, "*.xlsx")) + \
                  glob(os.path.join(args.input_dir, "*.xls")) + \
                  glob(os.path.join(args.input_dir, "*.csv"))

    if not input_files:
        print(f"No se encontraron archivos Excel/CSV en el directorio: {args.input_dir}")
        return

    for input_file in input_files:
        print(f"Procesando archivo: {input_file}")
        process_and_split_excel(input_file, output_dir, args.format, args.group_by, roles_mapping)


if __name__ == "__main__":
    main()