# DESACTIVAR PROTECCION DE LA HOJA EXCEL
#Importo librerias
import os
from openpyxl import load_workbook
import sys
from decouple import config  # Para importar variables de .env


# Importo funciones a usar
sys.path.append(os.path.abspath(os.path.dirname(__file__) + '/../'))  # Me dirije al directorio del proyecto
from functions.functions import directorio_a_trabajar, mostrar_error, mostrar_mensaje

# Importo y declaro las variables de .env
archivo_excel = config('NOMBRE_ARCHIVO_EXCEL')



## EJECUCIÓN
# Aseguro estar en el directorio correcto
directorio = directorio_a_trabajar()

# Ruta completa del archivo Excel
archivo_excel_ruta = os.path.join(directorio, archivo_excel)


## TRABAJO Y LIMPIEZA DEL DATAFRAME
# Verifica si el archivo Excel existe
if not os.path.exists(archivo_excel_ruta):
    # Muestra un mensaje de error si el archivo no existe
    mensaje_error = f"El excel '{archivo_excel}' no se encuentra en la carpeta. \n \n Se creará el archivo en la carpeta."
    mostrar_error(mensaje_error)

    # Corre el script src/0 Importar archivos a excel.py
    os.system("Python src/0_Importar_archivos_a_excel.py")
    
    # Corto la ejecución
    sys.exit()

else:
    print(archivo_excel_ruta)
    # Si el archivo existe, lo abriré para quitarle la protección
    try:
        wb = load_workbook(archivo_excel_ruta)
    except PermissionError: #Si el excel esta abierto.
        mostrar_error(f"Cerrar el excel para poder quitarle la protección de hoja.")
        # Corto la ejecución
        sys.exit()
        
        
    ws = wb.active  # Selecciona la primera hoja de trabajo (la activa)
    ws.protection.sheet = False
    # Guardar el archivo modificado
    wb.save(archivo_excel_ruta)
    wb.close()
    
    #Mostrar mensaje
    mostrar_mensaje('Desbloqueado', 'El excel fue desbloquado satisfactoriamente. \n \nNo modifique los encabezados de la celda A1:C1, No mueva las columnas A:C, no inserte columnas antes de la columna C y no modifique el nombre del excel.')