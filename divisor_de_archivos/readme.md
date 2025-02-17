# Divisor de Archivos Grandes (CSV, Excel, Parquet)

Este script de Python te permite dividir archivos grandes (CSV, Excel o Parquet) en archivos más pequeños, facilitando su manejo.

## 🚀 Guía Rápida

1.  **Preparación:**
    *   Descarga este script (el archivo `.py`).
    *   Crea dos carpetas en el mismo directorio que el script:
        *   `input`:  Aquí colocarás los archivos grandes a dividir.
        *   `output`:  Aquí se guardarán los archivos divididos.

2.  **División de Archivos:**

    1.  **Mueve los archivos:**  Copia los archivos CSV, Excel (.xls, .xlsx) o Parquet que deseas dividir a la carpeta `input`.
    2.  **Ejecuta el script:**
        *   **Windows:**
            *   Abre el "Símbolo del sistema" (busca "cmd" en el menú Inicio).
            *   Navega hasta la carpeta que contiene el script usando el comando `cd`.  Por ejemplo:
                ```
                cd Desktop\MiCarpetaConElScript
                ```
            *   Ejecuta el script:
                ```
                python nombre_del_script.py
                ```
                (Reemplaza `nombre_del_script.py` con el nombre real del archivo).
        *   **macOS / Linux:**
            *   Abre la "Terminal".
            *   Navega hasta la carpeta del script con `cd`:
                ```
                cd /ruta/a/la/carpeta/del/script
                ```
            *   Ejecuta el script:
                ```
                python3 nombre_del_script.py
                ```
                (Usa `python3` en lugar de `python` si es necesario).

    3.  **Espera:** El script mostrará mensajes de progreso en la terminal.
    4.  **Resultado:** Los archivos divididos se encontrarán en la carpeta `output`.

## ⚙️ Configuración Avanzada (Opcional)

### Tamaño de los Archivos de Salida

Por defecto, el script divide los archivos en partes de 5000 filas. Para cambiar esto:

1.  Abre el archivo `.py` con un *editor de texto plano* (Bloc de notas, TextEdit, *no* Word).
2.  Busca la línea:
    ```python
    def dividir_archivo(ruta_entrada, ruta_salida, filas_por_parte=5000, fila_encabezado=0):
    ```
3.  Reemplaza `5000` con el número deseado de filas por archivo.
4.  Guarda el archivo.

### Formato de archivo de Salida

5.  Guarda los cambios en el script.


### Fila de Encabezado (Archivos con Encabezados Desplazados)

*   **¿Qué es la fila de encabezado?**  La fila con los nombres de las columnas (ej: "Nombre", "Edad").
*   **¿Cómo encontrarla en Excel?**
    1.  Abre tu archivo Excel.
    2.  Cuenta las filas desde arriba (empezando en 1) hasta la fila de los nombres de columna.  Ese es el número de fila *en Excel*.
    3.  **Importante:** Para el script, resta 1 a ese número (las filas en el script empiezan desde 0).
*   **Ejemplos:**

    | Fila en Excel |  `fila_encabezado` (para el script) |
    |---------------|---------------------------------------|
    | 1             | 0                                     |
    | 2             | 1                                     |
    | 3             | 2                                     |

*   **Cómo configurar el script:**

    1.  Abre el archivo `.py` en un editor de texto.
    2.  Busca la línea que contiene:  `procesar_carpeta_input(fila_encabezado=...)`.  Por ejemplo:
       `procesar_carpeta_input(fila_encabezado=2)`
    3.  Cambia el número (ej: `2`) por el valor correcto de `fila_encabezado` (el número de fila en Excel *menos 1*).
    4.  Guarda el archivo.
    5. Opcionalmente, si se desea que ese encabezado sea el predefinido, cambiar también en:
        ```python
         def dividir_archivo(ruta_entrada, ruta_salida, filas_por_parte=5000, fila_encabezado=0):
        ```
        el `fila_encabezado=0` con el valor deseado.

## 📁 Estructura de Carpetas

*   **`input`:**  Coloca aquí los archivos originales que quieres dividir.
*   **`output`:**  Los archivos divididos aparecerán aquí.  El script creará esta carpeta si no existe.

## 📝 Notas

*   Este script requiere Python 3.
*   Necesitarás las bibliotecas `pandas` y `openpyxl` (y `pyarrow` si usas archivos Parquet). Si no las tienes, instálalas con:
    ```bash
    pip install pandas openpyxl pyarrow
    ```
    (o `pip3` en lugar de `pip` si es necesario).

¡Listo!  Con este script, dividir archivos grandes debería ser mucho más fácil.