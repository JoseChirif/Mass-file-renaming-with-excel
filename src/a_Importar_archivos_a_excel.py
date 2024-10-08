#Importo librerias a usar
import os
import sys
import pandas as pd

# Añadir el directorio del proyecto a sys.path para llamar functions.functions en mis scripts
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
# Importo funciones a usar
from functions.functions import choose_language, exit_if_directly_executed, directorio_a_trabajar, preguntar_reemplazo, mostrar_mensaje, mostrar_error, modificar_excel_dataframe, cargar_traducciones

# Importo parámetros de config
from config.config import project_name, excel_name, files_to_avoid

    
    
## DECLARACION VARIABLES
# Obtener el language desde los argumentos de línea de comandos
language = choose_language()

# Declaro la funcion main para al ejecutarlo en otros scripts trasmita el parámetro language
def main(language):
    # Cargar traducciones
    traducciones = cargar_traducciones(language)
    # Declaro textos de locales (cargar_traducciones) _1
    renamed = traducciones["renamed"]

    # Obtener carpeta destino
    folder_with_renamed_files = project_name + " - " + renamed
    


    # Declaro textos de locales (cargar_traducciones)_2 _final
    # Valores para Excel
    excel_column_original_name = traducciones["excel_column_original_name"]
    excel_column_extention = traducciones["excel_column_extention"]
    excel_column_new_name = traducciones["excel_column_new_name"]
    notes_title = traducciones["notes_title"]
    excel_note1 = traducciones["excel_note1"].format(excel_column_new_name=excel_column_new_name)
    excel_note2 = traducciones["excel_note2"].format(excel_column_new_name=excel_column_new_name)
    excel_note3 = traducciones["excel_note3"]
    excel_note4 = traducciones["excel_note4"].format(project_name=project_name, folder_with_renamed_files=folder_with_renamed_files)
    excel_note5 = traducciones["excel_note5"]
    # Mensajes Tkinter
    file_created_title = traducciones["file_created_title"]
    file_created_message = traducciones["file_created_message"].format(excel_name = excel_name)
    file_not_created_selected_title = traducciones["file_not_created_selected_title"]
    file_not_created_selected_message = traducciones["file_not_created_selected_message"]
    file_replaced_title = traducciones["file_replaced_title"]
    file_replaced_message = traducciones["file_replaced_message"].format(excel_name = excel_name)
    error_open_file = traducciones["error_open_file"].format(excel_name = excel_name)
    file_already_exist_title = traducciones['file_already_exist_title']
    file_already_exist_message = traducciones['file_already_exist_message'].format(excel_name = excel_name)
    error_text = traducciones['error_text']




    # Declaro variables únicas de este script
    excel_aditional_rows_to_unlock = 100



## EJECUCIÓN
    # Aseguro que el Excel se cree en la carpeta padre del proyecto
    excel_save_route = os.path.join(directorio_a_trabajar(), excel_name)


    ## DATAFRAME
    # Crea una lista para almacenar los data
    data = []

    # Recorre los files en el directorio actual
    for file in os.listdir(directorio_a_trabajar()):
        file_route = os.path.join(directorio_a_trabajar(), file)
        
        # Si el file no está en la lista de exclusión
        if file not in files_to_avoid:
            if os.path.isfile(file_route):
                # Separa el name y la extensión
                name, extention = os.path.splitext(file)
                # Añade los data a la lista (con extensión)
                data.append([name, extention, ""])  # name nuevo vacío
            elif os.path.isdir(file_route):
                # Si es una carpeta, se añade con extensión vacía
                data.append([file, "", ""])  # name nuevo vacío

    # Crea un DataFrame con los data
    df = pd.DataFrame(data, columns=[excel_column_original_name, excel_column_extention, excel_column_new_name])



    ##EXCEL
    # Verifica si el file ya existe
    if os.path.exists(excel_save_route):
        if not preguntar_reemplazo(file_already_exist_title, file_already_exist_message):
            mostrar_mensaje(file_not_created_selected_title, file_not_created_selected_message)
            # Corto la ejecución
            exit_if_directly_executed()
        else:
            try:
                df.to_excel(excel_save_route, index=False)
                modificar_excel_dataframe(excel_save_route, excel_aditional_rows_to_unlock, notes_title, excel_note1, excel_note2, excel_note3, excel_note4, excel_note5)
                mostrar_mensaje(file_replaced_title, file_replaced_message)
            except PermissionError: #Si el excel esta abierto.
                mostrar_error(error_text, error_open_file)
                # Corto la ejecución
                exit_if_directly_executed()
    else:
        try:
            df.to_excel(excel_save_route, index=False)
            modificar_excel_dataframe(excel_save_route, excel_aditional_rows_to_unlock, notes_title, excel_note1, excel_note2, excel_note3, excel_note4, excel_note5)
            mostrar_mensaje(file_created_title, file_created_message)
        except PermissionError: #Si el excel esta abierto.
            mostrar_error(error_text, error_open_file)
            # Corto la ejecución
            exit_if_directly_executed()
            
            
            

    
    
if __name__ == "__main__":
    main(language)
