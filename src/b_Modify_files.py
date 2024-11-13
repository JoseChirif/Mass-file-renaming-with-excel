#Importo librerias
import os
import pandas as pd


# Importo funciones a usar
from functions.functions import working_directory, show_error, show_message, show_options, unlock_excel_sheet, copy_files_with_new_names, rename_files_locally, load_translations, choose_language, exit_if_directly_executed
# Importo parámetros de config
from config.config import project_name, excel_name, files_to_avoid

# Importo otros scripts
from src.a_Create_excel import main as script0_main


language = choose_language()

# Declaro la funcion main para al ejecutarlo en otros scripts trasmita el parámetro language
def main(language):
    """
    Modifies the names of files in the current directory based on the Excel filescreated earlier. The Excel file must have the same name as the project or executable. The function uses the provided language parameter to handle any language-specific functionality.
    
    This function is located in src/b_Modify_files.py

    Arg:
        language (str): The language code (e.g., 'en', 'es') passed to customize language-specific content during the file renaming process.

    Returns:
        None: This function does not return any value. It modifies the names of the files in the directory based on the Excel file data.
    """
    # Cargar traducciones
    traducciones = load_translations(language)


    # Declaro textos de locales (load_translations) _1
    renamed = traducciones["renamed"]

    # Obtener carpeta destino
    folder_with_renamed_files = project_name + " - " + renamed    
    # Aseguro estar en el directorio correcto
    directorio = working_directory()


    # Declaro textos de locales (load_translations) _2 _Final
    select_an_option_text = traducciones["select_an_option_text"]
    error_open_file = traducciones["error_open_file"].format(excel_name=excel_name)
    excel_column_original_name = traducciones['excel_column_original_name']
    excel_column_extention = traducciones['excel_column_extention']
    excel_column_new_name = traducciones["excel_column_new_name"]
    excel_column_original_name_final = traducciones["excel_column_original_name_final"]
    excel_column_new_name_final = traducciones["excel_column_new_name_final"]
    excel_error_column_new_names_empty = traducciones["excel_error_column_new_names_empty"].format(excel_column_new_name=excel_column_new_name)
    excel_column_status = traducciones["excel_column_status"]
    excel_column_error = traducciones["excel_column_error"]
    option1 = traducciones["option1"]
    option2 = traducciones["option2"]
    error_excel_not_found = traducciones["error_excel_not_found"].format(excel_name=excel_name)
    error_excel_structure_modified = traducciones["error_excel_structure_modified"].format(excel_column_original_name=excel_column_original_name, excel_column_extention=excel_column_extention, excel_column_new_name=excel_column_new_name)
    error_unknown = traducciones["error_unknown"]
    excel_error_file_not_found = traducciones["excel_error_file_not_found"]
    excel_error_file_already_proceced = traducciones["excel_error_file_already_proceced"].format(excel_column_new_name_final=excel_column_new_name_final)
    excel_error_file_already_exist_in_new_directory = traducciones["excel_error_file_already_exist_in_new_directory"].format(excel_column_new_name_final=excel_column_new_name_final, folder_with_renamed_files=folder_with_renamed_files)
    excel_error_file_already_exist_in_the_same_folder = traducciones['excel_error_file_already_exist_in_the_same_folder'].format(excel_column_new_name_final=excel_column_new_name_final)
    excel_error_conflic = traducciones['excel_error_conflic'].format(excel_column_new_name_final=excel_column_new_name_final)
    excel_template_error_not_found = traducciones["excel_template_error_not_found"].format(excel_name=excel_name)
    excel_template_error_not_allowed = traducciones["excel_template_error_not_allowed"].format(excel_name=excel_name)
    success_title = traducciones['success_title']
    success_message = traducciones['success_message'].format(folder_with_renamed_files=folder_with_renamed_files)
    error_text = traducciones['error_text']
    excel_error_files_not_found = traducciones['excel_error_files_not_found']
    excel_errors_log_message = traducciones['excel_errors_log_message'].format(excel_column_error = excel_column_error)
    cancel_title = traducciones['cancel_title']
    start_function_title = traducciones['start_function_title']
    start_function_message = traducciones['start_function_message']
    operation_canceled_title = traducciones['operation_canceled_title']
    operation_canceled_message = traducciones['operation_canceled_message']






    #opciones metodo
    border1 = "black"
    border2 = "red"




    ## EJECUCIÓN


    # Ruta completa del archivo Excel
    excel_save_route = os.path.join(directorio, excel_name)


    ## TRABAJO Y LIMPIEZA DEL DATAFRAME
    # Verifica si el archivo Excel existe
    if not os.path.exists(excel_save_route):
        # Muestra un mensaje de error si el archivo no existe
        show_error(error_text, error_excel_not_found)

        # Corre el script src/0 Importar archivos a excel.py
        script0_main(language)
        
        # Corto la ejecución
        exit_if_directly_executed()

    else:
        # Si el archivo existe, intantará leerlo
        try:
            df = pd.read_excel(excel_save_route, usecols=[0, 1, 2])
        except PermissionError:
            # Si el archivo está abierto o en uso, se muestra un mensaje de error y corto la ejecución
            show_error(error_text, error_open_file)
            exit_if_directly_executed()
            
        
        # Si la columna 'A' esta vacia, desbloqueamos el excel y mostramos error:
        if df.iloc[:, 0].isnull().all():
            unlock_excel_sheet(excel_save_route)
            show_error(error_text, error_excel_structure_modified)
            # Corto la ejecución
            exit_if_directly_executed()
        
        
        # Elimino las celdas que en la columna 'A' están vacias
        df = df[pd.notna(df.iloc[:, 0])]
        
        #Crear valores vacio en vez de la asignación 'nan' en el df
        df = df.fillna('')
        
    
        # Convertir los nombres a tipo str
        df.iloc[:, [0, 2]] = df.iloc[:, [0, 2]].astype(str)
    

        # Crear la columna excel_column_original_name_final concatenando columna_nombre_original_excel_inicial y columna_extention_excel_inicial
        df[excel_column_original_name_final] = df.iloc[:, 0] + df.iloc[:, 1]

        
        #LIMPIEZA
        # Creo df_conflicto, filtrando df con files_to_avoid para mostrar error al final
        df_conflicto = df[df.loc[:, excel_column_original_name_final].isin(files_to_avoid)]

        # ELimino los archivos que pueden generar conflicto
        df = df[~df.loc[:, excel_column_original_name_final].isin(files_to_avoid)]


        # Crear la columna excel_column_new_name_final (original + extention si nuevo esta vacio. Sino, nuevo + extention)
        df[excel_column_new_name_final] = df.apply(
        lambda row: (row.iloc[0] + row.iloc[1])  
        if row.iloc[2] == ''  
        else (row.iloc[2] + row.iloc[1]), 
        axis=1
        )
        
        #repetir con df_conflicto para crear el log y crear la columna error para que quede igual al df al final del proceso
        df_conflicto[excel_column_new_name_final] = df_conflicto.apply(
        lambda row: (row.iloc[0] + row.iloc[1])  
        if row.iloc[2] == ''  
        else (row.iloc[2] + row.iloc[1]), 
        axis=1
        )
        df_conflicto[excel_column_error] = excel_error_conflic
        
        
        # Verifico si hay nombres para modificar. Si no hay, muestró mensaje de error y cierro el programa.
        if df.iloc[:, 2].isnull().all() or (df.iloc[:, 2] == "").all():
            show_error(error_text, excel_error_column_new_names_empty)
            exit_if_directly_executed()

    



    ## MÉTODOS PARA MODIFICAR NOMBRES  
    # Pregunta método
    nro_opcion = show_options(select_an_option_text, option1, option2, border1, border2, cancel_title)

    # Verificando nro_opcion
    # print(f'Se escogió la opción {nro_opcion}')
    # if nro_opcion == None:
    #     print("Cancelar")
    # else:
    #     print(f"Opción seleccionada: {globals()[f'opcion{nro_opcion}']}" if f'opcion{nro_opcion}' in globals() else "Error: Opción no válida.")
    
    
    # Si da cancelar, corto la ejecución.   
    if nro_opcion == None:
        exit_if_directly_executed()
    # Si da opción 1, usaré método de nueva carpeta. 
    elif nro_opcion == 1:
        #Crea la carpeta y copia los archivos con nuevo nombre:
        ruta_carpeta_destino = os.path.join(directorio, folder_with_renamed_files)
        # Correr metodo
        copy_files_with_new_names(df,directorio,ruta_carpeta_destino, excel_error_file_not_found, excel_error_file_already_proceced, excel_error_file_already_exist_in_new_directory, excel_template_error_not_found, excel_template_error_not_allowed, excel_column_status)
    #Si se da la opción 2, usaré el mismo metodo copiandolo en la misma carpeta
    elif nro_opcion == 2:
        # Correr metodo sobre la misma carpeta
        df = rename_files_locally(df, directorio, excel_error_file_not_found, excel_error_file_already_proceced, excel_error_file_already_exist_in_the_same_folder, excel_template_error_not_found, excel_template_error_not_allowed, excel_column_status, start_function_title, start_function_message, operation_canceled_title, operation_canceled_message)
        folder_with_renamed_files = working_directory()
        
        
    
    # Otra opcion marcará error y saldrá del programa
    else:
        show_error(error_text, error_unknown)
        exit_if_directly_executed()
        




    ## CREACIÓN DEL LOG DE excel_column_error EN EXCEL
    # Filtrar el DataFrame para que excel_column_status sea diferente de 'Ok'
    df_error = df[df[excel_column_status] != 'Ok']

    #Si el len(df_error) = 0, mostrar mensaje de éxito y salir.
    if len(df_error)==0:
        show_message(success_title, success_message)
        exit_if_directly_executed()
        
    #Si len(df_error) != 0 Entonces crear excel con log de excel_column_error
    else: 
        # Cambiar el nombre de la columna excel_column_status a excel_column_error
        df_error = df_error.rename(columns={excel_column_status: excel_column_error})
        
        # Añado df_conflicto
        df_error = pd.concat([df_conflicto, df_error], ignore_index=True)

        # Guardar el DataFrame en un archivo Excel (no reemplazar el último):
        excel_error_log = excel_name + "_" + error_text
        extention = ".xlsx"
        contador = 1
        
        # Ruta completa inicial
        ruta_completa = os.path.join(directorio, f"{excel_error_log} - {contador}{extention}")

        # Verificar si el archivo existe y agregar sufijos con números si es necesario
        while os.path.exists(ruta_completa):
            contador += 1
            ruta_completa = os.path.join(directorio, f"{excel_error_log} - {contador}{extention}")


        # Filtrar la columna 6 (excel_column_error) para asegurarse que la columna A y B formen archivos de la carpeta
        filtro = df_error[df_error.iloc[:, 5].str.startswith(excel_error_file_not_found)]
        # Si el filtro y df tiene la misma cantidad de filas => La columna 'A' tiene nombres que no corresponden a los archivos de la carpeta:
        if len(filtro) == len(df):
            # Si se da la condición marco error
            unlock_excel_sheet(excel_save_route)
            show_error(error_text, excel_error_files_not_found + error_excel_structure_modified)
            # Finalizar la ejecución
            exit_if_directly_executed()

            
        # Modifico los valores de la columna 6 (errores) que tengan el titulo de la columna 5 (nombre nuevo final)
        df_error.iloc[:, 5] = df_error.apply(lambda row: row.iloc[5].replace(f"{excel_column_new_name_final}", str(row.iloc[4])) if isinstance(row.iloc[5], str) else row.iloc[5], axis=1)

        
        # Guardar el archivo en la ruta final
        df_error.to_excel(ruta_completa, index=False)

        show_error(error_text, excel_errors_log_message + f'{excel_error_log} - {contador}{extention}')
        # Finalizar la ejecución
        exit_if_directly_executed()



if __name__ == "__main__":
    main(language)
    
    
