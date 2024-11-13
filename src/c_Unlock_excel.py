# DESACTIVAR PROTECCION DE LA HOJA EXCEL
#Importo librerias
import os
from openpyxl import load_workbook



# Importo funciones a usar
from functions.functions import working_directory, show_error, show_message, load_translations,choose_language, exit_if_directly_executed
# Importo parámetros de config
from config.config import excel_name

# Importo otros scripts
from src.a_Create_excel import main as script0_main


language = choose_language()


# Declaro la funcion main para al ejecutarlo en otros scripts trasmita el parámetro language
def main(language):
    """
    Unlocks the Excel sheet with the same name as the project or executable file, enabling full editing. The function uses the provided language parameter to handle any language-specific functionality.

    This function is located in src/c_Unlock_excel.py

    Arg:
        language (str): The language code (e.g., 'en', 'es') passed to customize language-specific content during the unlocking process.

    Returns:
        None: This function does not return any value. It modifies the Excel file by unlocking the sheet to allow full editing.
    """
    # Cargar traducciones
    traducciones = load_translations(language)

    # Declaro textos de locales (load_translations) _1
    renamed = traducciones["renamed"]

    # Declaro textos de locales (load_translations) _2 _FINAL
    error_text = traducciones['error_text']
    error_excel_not_found = traducciones['error_excel_not_found'].format(excel_name=excel_name)
    error_excel_open = traducciones['error_excel_open'].format(excel_name=excel_name)
    unlocked_title = traducciones['unlocked_title']
    unlocked_message = traducciones['unlocked_message']

    ## EJECUCIÓN
    # Aseguro estar en el directorio correcto
    directorio = working_directory()

    # Ruta completa del archivo Excel
    excel_name_ruta = os.path.join(directorio, excel_name)


    ## TRABAJO Y LIMPIEZA DEL DATAFRAME
    # Verifica si el archivo Excel existe
    if not os.path.exists(excel_name_ruta):
        # Muestra un mensaje de error si el archivo no existe
        show_error(error_text, error_excel_not_found)

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
            show_error(error_text, error_excel_open)
            # Corto la ejecución
            exit_if_directly_executed()
            
            
        ws = wb.active  # Selecciona la primera hoja de trabajo (la activa)
        ws.protection.sheet = False
        # Guardar el archivo modificado
        wb.save(excel_name_ruta)
        wb.close()
        
        #Mostrar mensaje
        show_message(unlocked_title, unlocked_message)


if __name__ == "__main__":
    main(language)