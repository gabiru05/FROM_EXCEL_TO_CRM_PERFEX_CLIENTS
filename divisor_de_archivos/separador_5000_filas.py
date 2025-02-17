import os
import pandas as pd
import math

def dividir_archivo(ruta_entrada, ruta_salida, filas_por_parte=5000, fila_encabezado=0):
    """
    Divide un archivo CSV, Excel (xls, xlsx) o Parquet en múltiples archivos Excel,
    preservando el encabezado en cada parte.
    """
    try:
        nombre_base = os.path.splitext(os.path.basename(ruta_entrada))[0]
        extension = os.path.splitext(ruta_entrada)[1].lower()

        if extension == '.csv':
            try:
                df = pd.read_csv(ruta_entrada, encoding='utf-8', header=fila_encabezado)
                encabezado = df.columns  # Guarda el encabezado
            except UnicodeDecodeError:
                try:
                    df = pd.read_csv(ruta_entrada, encoding='latin-1', header=fila_encabezado)
                    encabezado = df.columns
                except UnicodeDecodeError:
                    df = pd.read_csv(ruta_entrada, encoding='cp1252', header=fila_encabezado)
                    encabezado = df.columns

            for i in range(0, len(df), filas_por_parte):
                chunk = df[i:i + filas_por_parte]
                # Agrega el encabezado al chunk
                chunk = pd.concat([pd.DataFrame(columns=encabezado), chunk], ignore_index=False)
                nombre_archivo_salida = f"{nombre_base}_parte_{i // filas_por_parte + 1}.xlsx"
                ruta_completa_salida = os.path.join(ruta_salida, nombre_archivo_salida)
                chunk.to_excel(ruta_completa_salida, index=False, engine='openpyxl')
                print(f"Guardado: {nombre_archivo_salida}")

        elif extension in ['.xls', '.xlsx']:
            xls = pd.ExcelFile(ruta_entrada)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=fila_encabezado)
                encabezado = df.columns  # Guarda el encabezado

                for i in range(0, len(df), filas_por_parte):
                    chunk = df[i:i + filas_por_parte]
                     # Agrega el encabezado al chunk
                    chunk = pd.concat([pd.DataFrame(columns=encabezado), chunk], ignore_index=False)
                    nombre_archivo_salida = f"{nombre_base}_{sheet_name}_parte_{i // filas_por_parte + 1}.xlsx"
                    ruta_completa_salida = os.path.join(ruta_salida, nombre_archivo_salida)
                    chunk.to_excel(ruta_completa_salida, index=False, sheet_name="Sheet1", engine='openpyxl')
                    print(f"Guardado: {nombre_archivo_salida}")
            xls.close()


        elif extension == '.parquet':
            df = pd.read_parquet(ruta_entrada, engine='pyarrow')
            if fila_encabezado != 0:
              header_rows = pd.read_parquet(ruta_entrada, engine='pyarrow').head(fila_encabezado)
              new_header = ['_'.join(col).strip() for col in header_rows.columns.to_flat_index()]
              df = pd.read_parquet(ruta_entrada, engine='pyarrow')
              df = df.iloc[fila_encabezado:]  # Ya no elimina filas
              df.columns = new_header  # Ya no elimina filas

            encabezado = df.columns  # Guarda el encabezado

            for i in range(0, len(df), filas_por_parte):
                chunk = df[i:i + filas_por_parte]
                # Agrega el encabezado
                chunk = pd.concat([pd.DataFrame(columns=encabezado), chunk], ignore_index=False)
                nombre_archivo_salida = f"{nombre_base}_parte_{i // filas_por_parte + 1}.xlsx"
                ruta_completa_salida = os.path.join(ruta_salida, nombre_archivo_salida)
                chunk.to_excel(ruta_completa_salida, index=False, engine='openpyxl')
                print(f"Guardado: {nombre_archivo_salida}")

        else:
            print(f"Error: Tipo de archivo no soportado ({extension}).")


    except FileNotFoundError:
        print(f"Error: Archivo no encontrado: {ruta_entrada}")
    except pd.errors.EmptyDataError:
        print(f"Error: El archivo {ruta_entrada} está vacío.")
    except Exception as e:
        print(f"Error inesperado: {e}")


def procesar_carpeta_input(carpeta_input="input", carpeta_output="output", fila_encabezado=0):

    if not os.path.exists(carpeta_output):
        os.makedirs(carpeta_output)
        print(f"Carpeta de salida '{carpeta_output}' creada.")

    if not os.path.exists(carpeta_input):
        print(f"Error: La carpeta de entrada '{carpeta_input}' no existe.")
        return

    for nombre_archivo in os.listdir(carpeta_input):
        ruta_completa_entrada = os.path.join(carpeta_input, nombre_archivo)

        if os.path.isfile(ruta_completa_entrada):
            print(f"Procesando: {nombre_archivo}")
            dividir_archivo(ruta_completa_entrada, carpeta_output, fila_encabezado=fila_encabezado)

    archivos_salida = [f for f in os.listdir(carpeta_output) if os.path.isfile(os.path.join(carpeta_output, f))]
    num_archivos = len(archivos_salida)

    if num_archivos > 0:
        num_carpetas = 2
        archivos_por_carpeta = math.ceil(num_archivos / num_carpetas)

        carpeta1 = os.path.join(carpeta_output, "parte1")
        carpeta2 = os.path.join(carpeta_output, "parte2")
        os.makedirs(carpeta1, exist_ok=True)
        os.makedirs(carpeta2, exist_ok=True)

        for i, archivo in enumerate(archivos_salida):
            if i < archivos_por_carpeta:
                origen = os.path.join(carpeta_output, archivo)
                destino = os.path.join(carpeta1, archivo)
            else:
                origen = os.path.join(carpeta_output, archivo)
                destino = os.path.join(carpeta2, archivo)
            os.rename(origen, destino)

if __name__ == "__main__":
    procesar_carpeta_input(fila_encabezado=0)  # Ajusta si es necesario
    print("Proceso completado.")