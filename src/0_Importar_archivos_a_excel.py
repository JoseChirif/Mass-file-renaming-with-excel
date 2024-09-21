#Importo librerias a usar
import os
import sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Protection
from decouple import config # Para importarvariables de .env 

# Importo funciones a usar
sys.path.append(os.path.abspath(os.path.dirname(__file__) + '/../'))  # Agrega el directorio donde se organiza el proyecto (incluyendo funciones)
from functions.functions import directorio_a_trabajar, preguntar_reemplazo, mostrar_mensaje, mostrar_error, modificar_excel_dataframe  # Me da el directorio donde trabajaré


# Importo y declaro las variables de .env
nombre_archivo_excel_base = config('NOMBRE_ARCHIVO_EXCEL')
columna_nombre_original_excel_inicial = config('COLUMNA_1_EXCEL')
columna_extension_excel_incial = config('COLUMNA_2_EXCEL')
columna_nombre_nuevo_excel_inicial = config('COLUMNA_3_EXCEL')
nombre_carpeta_destino = config("CARPETA_DESTINO") # Para poner en las notas del excel

#Declaro varaibles tipo lista
# Nombres de archivos a excluir
archivos_excluir = [
    nombre_archivo_excel_base,
    nombre_carpeta_destino, #Para evitar un bucle que se vaya copiando el contenido dentro cada que ejecuto
    os.path.basename(__file__) # Obtiene el nombre del archivo actual
]

# Declaro variables únicas de este script
filas_adicionales_a_desbloquear_excel = 100



## EJECUCIÓN
# Aseguro que el Excel se cree en la carpeta padre del proyecto
nombre_archivo_excel_ruta = os.path.join(directorio_a_trabajar(), nombre_archivo_excel_base)

## DATAFRAME
# Crea una lista para almacenar los datos
datos = []

# Recorre los archivos en el directorio actual
for archivo in os.listdir(directorio_a_trabajar()):
    ruta_archivo = os.path.join(directorio_a_trabajar(), archivo)
    
    # Si el archivo no está en la lista de exclusión
    if archivo not in archivos_excluir:
        if os.path.isfile(ruta_archivo):
            # Separa el nombre y la extensión
            nombre, extension = os.path.splitext(archivo)
            # Añade los datos a la lista (con extensión)
            datos.append([nombre, extension, ""])  # Nombre nuevo vacío
        elif os.path.isdir(ruta_archivo):
            # Si es una carpeta, se añade con extensión vacía
            datos.append([archivo, "", ""])  # Nombre nuevo vacío

# Crea un DataFrame con los datos
df = pd.DataFrame(datos, columns=[columna_nombre_original_excel_inicial, columna_extension_excel_incial, columna_nombre_nuevo_excel_inicial])

# Eliminar la carpeta del proyecto en caso se este trabajando sobre el mismo Python. Esto evitará un bucle de carpetas infinitas una dentro de otra.
nombre_carpeta_actual = os.path.basename(os.getcwd())
df = df[df[columna_nombre_original_excel_inicial] != nombre_carpeta_actual]


##EXCEL
# Parámetros de aviso creación excel
mensaje = f'Plantilla guardada en {nombre_archivo_excel_base}. \n \nNo modifique los encabezados de la celda A1:C1, No mueva las columnas A:C, no inserte columnas antes de la columna C y no modifique el nombre del excel.'
titulo = "Archivo creado"

# Verifica si el archivo ya existe
if os.path.exists(nombre_archivo_excel_ruta):
    if not preguntar_reemplazo(nombre_archivo_excel_ruta):
        mostrar_mensaje("Archivo no creado", "Usted selecciono no reemplazar el archivo existente")
        # Corto la ejecución
        sys.exit()
    else:
        try:
            df.to_excel(nombre_archivo_excel_ruta, index=False)
            modificar_excel_dataframe(nombre_archivo_excel_ruta, columna_nombre_nuevo_excel_inicial, nombre_carpeta_destino, filas_adicionales_a_desbloquear_excel)
            mostrar_mensaje("Archivo creado", f"El archivo {nombre_archivo_excel_base} se reeemplazo satisfactoriamente.")
        except PermissionError: #Si el excel esta abierto.
            mostrar_error(f"Cerrar '{nombre_archivo_excel_base}' para poder guardar uno nuevo.")
            # Corto la ejecución
            sys.exit()
else:
    try:
        df.to_excel(nombre_archivo_excel_ruta, index=False)
        modificar_excel_dataframe(nombre_archivo_excel_ruta, columna_nombre_nuevo_excel_inicial, nombre_carpeta_destino, filas_adicionales_a_desbloquear_excel)
        mostrar_mensaje(titulo, mensaje)
    except PermissionError: #Si el excel esta abierto.
        mostrar_error(f"Cerrar '{nombre_archivo_excel_base}' para poder guardar uno nuevo.")
        # Corto la ejecución
        sys.exit()
        

