# DESACTIVAR PROTECCION DE LA HOJA EXCEL
#Importo librerias
import os
from openpyxl import load_workbook
import sys


# Añadir el directorio del proyecto a sys.path para llamar functions.functions en mis scripts
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
# Importo funciones a usar
from functions.functions import directorio_a_trabajar, mostrar_error, mostrar_mensaje, obtener_nombre_archivo_o_directorio, cargar_traducciones,choose_language, exit_if_directly_executed
# Importo parámetros de config
from config.config import excel_name

# Importo otros scripts
from src.a_Importar_archivos_a_excel import main as script0_main


language = choose_language()


# Declaro la funcion main para al ejecutarlo en otros scripts trasmita el parámetro language
def main(language):
    # Cargar traducciones
    traducciones = cargar_traducciones(language)

    # Declaro textos de locales (cargar_traducciones) _1
    renamed = traducciones["renamed"]

    # Declaro textos de locales (cargar_traducciones) _2 _FINAL
    error_text = traducciones['error_text']
    error_excel_not_found = traducciones['error_excel_not_found'].format(excel_name=excel_name)
    error_excel_open = traducciones['error_excel_open'].format(excel_name=excel_name)
    unlocked_title = traducciones['unlocked_title']
    unlocked_message = traducciones['unlocked_message']

    ## EJECUCIÓN
    # Aseguro estar en el directorio correcto
    directorio = directorio_a_trabajar()

    # Ruta completa del archivo Excel
    excel_name_ruta = os.path.join(directorio, excel_name)


    ## TRABAJO Y LIMPIEZA DEL DATAFRAME
    # Verifica si el archivo Excel existe
    if not os.path.exists(excel_name_ruta):
        # Muestra un mensaje de error si el archivo no existe
        mostrar_error(error_text, error_excel_not_found)

        # Corre el script src/0 Importar archivos a excel.py
        script0_main(language)
        
        # Corto la ejecución
        exit_if_directly_executed()

    else:
        #print(excel_name_ruta)
        # Si el archivo existe, lo abriré para quitarle la protección
        try:
            wb = load_workbook(excel_name_ruta)
        except PermissionError: #Si el excel esta abierto.
            mostrar_error(error_text, error_excel_open)
            # Corto la ejecución
            exit_if_directly_executed()
            
            
        ws = wb.active  # Selecciona la primera hoja de trabajo (la activa)
        ws.protection.sheet = False
        # Guardar el archivo modificado
        wb.save(excel_name_ruta)
        wb.close()
        
        #Mostrar mensaje
        mostrar_mensaje(unlocked_title, unlocked_message)


if __name__ == "__main__":
    main(language)