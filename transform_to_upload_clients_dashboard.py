import pandas as pd
import re
import os  # Importamos el módulo os


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


def process_and_split_excel(input_file, output_dir="output"):
    """Procesa el Excel, transforma los datos y los divide en archivos."""

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
                print(f"Error: No se pudo leer el archivo ni como CSV ni como Excel.")
                print(f"Errores originales:\nCSV (latin-1): {e1}\nExcel (utf-8): {e2}\nCSV (utf-8): {e3}")
                return

    # Crea el directorio de salida si no existe
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Agrupa por 'GRUPO / TALLER'
    grouped = df.groupby('GRUPO / TALLER')

    # Itera sobre cada grupo
    for group_name, group_df in grouped:
        data = []
        #Procesa cada grupo (mismo procesamiento que teniamos antes)
        for index, row in group_df.iterrows():
            company = row['Nombre_empresa']
            phone_number = str(row['Telefonos']) if pd.notna(row['Telefonos']) else ''
            phone_number = re.sub(r'\D', '', phone_number)

            roles_mapping = {
                'Representante Principal': ['Representante Principal', 'Email'],
                'Representante Suplente': ['Representante Suplente', 'Email.1'],
                'Asistente de Gerencia': ['Asistente de Gerencia', 'Email.2'],
                'Gerente General': ['Gerente General', 'Email.3'],
                'Recursos Humanos': ['Recursos Humanos', 'Email.4'],
                'Mercadeo': ['Mercadeo', 'Email.5'],
                'Ventas': ['Ventas', 'Email.6'],
            }

            for position, cols in roles_mapping.items():
                name_col, email_col = cols
                full_name = str(row[name_col]) if pd.notna(row[name_col]) else ''
                email = row[email_col] if pd.notna(row[email_col]) else ''

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
        #Si el dataframe resultante esta vacio no guardamos nada
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

            # Limpia el nombre del grupo para usarlo como nombre de archivo
           output_filename = re.sub(r'[\\/*?:"<>|]', "", group_name)  # Elimina caracteres inválidos
           output_filename = output_filename.strip() #Remueve espacios
           output_filepath = os.path.join(output_dir, f"{output_filename}.xlsx")

            # Guarda el DataFrame en un archivo Excel
           try:
                output_df.to_excel(output_filepath, index=False)
                print(f"Datos de '{group_name}' guardados en '{output_filepath}'")
           except Exception as e:
                print(f"Error al guardar '{output_filepath}': {e}")
        else:
            print(f"No hay datos para el grupo '{group_name}', no se genera archivo.")



input_excel_file = "!-Directorio 154.xlsx"  # Cambia por el nombre de tu archivo
output_directory = "output"  # Carpeta donde se guardarán los archivos
process_and_split_excel(input_excel_file, output_directory)