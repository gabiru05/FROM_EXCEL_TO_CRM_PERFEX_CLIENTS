# Procesador de Directorios Excel/CSV

Este script de Python procesa archivos Excel (.xlsx, .xls) y CSV en un directorio, transforma los datos, y los divide en múltiples archivos de salida (Excel o CSV) agrupados por una columna específica.  Es altamente configurable mediante argumentos de línea de comandos. Ideal Si tiene que analizar archivos Excel para subirlo en clientes PERFEXCRM. 

## Requisitos

*   Python 3.6+
*   Librerías de Python:
    *   `pandas`
    *   `openpyxl` (necesaria para escribir archivos Excel .xlsx)
    *   `argparse`
    *   `glob`
    *   `re`
    * `os`

Instala las librerías necesarias (si no las tienes) usando `pip`:

```bash
pip install pandas openpyxl
```


# Uso

El script se ejecuta desde la línea de comandos (Terminal en macOS/Linux, Símbolo del sistema o PowerShell en Windows).

Formato de uso: carpeta de entrada -o carpeta de salida -f formato de salida 

```bash
python procesar_directorio.py entrada -o salida -f csv
```

ejemplos para mi uso:

```bash
python transform_to_upload_clients_dashboard.py input -o output -f csv
```