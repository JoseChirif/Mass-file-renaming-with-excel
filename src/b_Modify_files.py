#Import libraries
import os
import pandas as pd

# go to the parent directory if you are running this script directly (uncomment the following lines)
# import sys
# sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import functions
from functions.functions import working_directory, show_error, show_message, show_options, unlock_excel_sheet, copy_files_with_new_names, rename_files_locally, load_translations, choose_language, exit_if_directly_executed
# Import parameters from config
from config.config import project_name, excel_name, files_to_avoid

# Import another scritps
from src.a_Create_excel import main as script0_main


language = choose_language()

# Declare the main function to transmit the language parameter when executing it in other scripts.
def main(language):
    """
    Modifies the names of files in the current directory based on the Excel filescreated earlier. The Excel file must have the same name as the project or executable. The function uses the provided language parameter to handle any language-specific functionality.
    
    This function is located in src/b_Modify_files.py

    Arg:
        language (str): The language code (e.g., 'en', 'es') passed to customize language-specific content during the file renaming process.

    Returns:
        None: This function does not return any value. It modifies the names of the files in the directory based on the Excel file data.
    """
    # Load translations
    translations = load_translations(language)


    # load translations (part 1)
    renamed = translations["renamed"]

    # Get destination folder
    folder_with_renamed_files = project_name + " - " + renamed    
    # Work into working directory
    directory = working_directory()


    # load translations (part 1 - final)
    select_an_option_text = translations["select_an_option_text"]
    error_open_file = translations["error_open_file"].format(excel_name=excel_name)
    excel_column_original_name = translations['excel_column_original_name']
    excel_column_extention = translations['excel_column_extention']
    excel_column_new_name = translations["excel_column_new_name"]
    excel_column_original_name_final = translations["excel_column_original_name_final"]
    excel_column_new_name_final = translations["excel_column_new_name_final"]
    excel_error_column_new_names_empty = translations["excel_error_column_new_names_empty"].format(excel_column_new_name=excel_column_new_name)
    excel_column_status = translations["excel_column_status"]
    excel_column_error = translations["excel_column_error"]
    option1 = translations["option1"]
    option2 = translations["option2"]
    error_excel_not_found = translations["error_excel_not_found"].format(excel_name=excel_name)
    error_excel_structure_modified = translations["error_excel_structure_modified"].format(excel_column_original_name=excel_column_original_name, excel_column_extention=excel_column_extention, excel_column_new_name=excel_column_new_name)
    error_unknown = translations["error_unknown"]
    excel_error_file_not_found = translations["excel_error_file_not_found"]
    excel_error_file_already_proceced = translations["excel_error_file_already_proceced"].format(excel_column_new_name_final=excel_column_new_name_final)
    excel_error_file_already_exist_in_new_directory = translations["excel_error_file_already_exist_in_new_directory"].format(excel_column_new_name_final=excel_column_new_name_final, folder_with_renamed_files=folder_with_renamed_files)
    excel_error_file_already_exist_in_the_same_folder = translations['excel_error_file_already_exist_in_the_same_folder'].format(excel_column_new_name_final=excel_column_new_name_final)
    excel_error_conflic = translations['excel_error_conflic'].format(excel_column_new_name_final=excel_column_new_name_final)
    excel_template_error_not_found = translations["excel_template_error_not_found"].format(excel_name=excel_name)
    excel_template_error_not_allowed = translations["excel_template_error_not_allowed"].format(excel_name=excel_name)
    success_title = translations['success_title']
    success_message = translations['success_message'].format(folder_with_renamed_files=folder_with_renamed_files)
    error_text = translations['error_text']
    excel_error_files_not_found = translations['excel_error_files_not_found']
    excel_errors_log_message = translations['excel_errors_log_message'].format(excel_column_error = excel_column_error)
    cancel_title = translations['cancel_title']
    start_function_title = translations['start_function_title']
    start_function_message = translations['start_function_message']
    operation_canceled_title = translations['operation_canceled_title']
    operation_canceled_message = translations['operation_canceled_message']


    # border colors
    border1 = "black"
    border2 = "red"




    ## EXECUTION


    # Excel file path
    excel_save_route = os.path.join(directory, excel_name)


    ## DATAFRAME WORKING AND CLEANING
    # Check if the Excel file exists
    if not os.path.exists(excel_save_route):
        #If it doesn't exist
        # Displays an error message if the file does not exist
        show_error(error_text, error_excel_not_found)
        # Run src/a_Create_excel.py
        script0_main(language)
        # Abort the execution
        exit_if_directly_executed()

    else:
        # If the file exists, it will try to read it.
        try:
            df = pd.read_excel(excel_save_route, usecols=[0, 1, 2])
        except PermissionError:
            # If the file is open or in use, an error message is displayed and execution is aborted.
            show_error(error_text, error_open_file)
            exit_if_directly_executed()
            
        
        # If column 'A' is empty, we unlock excel and display error:
        if df.iloc[:, 0].isnull().all():
            unlock_excel_sheet(excel_save_route)
            show_error(error_text, error_excel_structure_modified)
            # Abort the execution
            exit_if_directly_executed()
        
        
        # Delete the cells that are empty in column 'A'.
        df = df[pd.notna(df.iloc[:, 0])]
        
        # Create empty values instead of the assignment 'nan' in df
        df = df.fillna('')
        
    
        # Convert names to type str
        df.iloc[:, [0, 2]] = df.iloc[:, [0, 2]].astype(str)
    

        # Create column excel_column_original_original_name_final by concatenating column_original_name_excel_initial and column_extention_excel_initial
        df[excel_column_original_name_final] = df.iloc[:, 0] + df.iloc[:, 1]

        
        # CLEANING
        # I create df_conflict, filtering df with files_to_avoid to show error at the end
        df_conflict = df[df.loc[:, excel_column_original_name_final].isin(files_to_avoid)]

        # Delete files that may generate conflict
        df = df[~df.loc[:, excel_column_original_name_final].isin(files_to_avoid)]


        # Create column excel_column_new_name_final (original + extention if new is empty, otherwise new + extention)
        df[excel_column_new_name_final] = df.apply(
        lambda row: (row.iloc[0] + row.iloc[1])  
        if row.iloc[2] == ''  
        else (row.iloc[2] + row.iloc[1]), 
        axis=1
        )
        
        # Repeat the process with df_conflict to create the log and create the error column to be equal to the df at the end of the process
        df_conflict[excel_column_new_name_final] = df_conflict.apply(
        lambda row: (row.iloc[0] + row.iloc[1])  
        if row.iloc[2] == ''  
        else (row.iloc[2] + row.iloc[1]), 
        axis=1
        )
        df_conflict[excel_column_error] = excel_error_conflic
        
        
        # Check if there are names to modify. If there are not, I show an error message and close the program..
        if df.iloc[:, 2].isnull().all() or (df.iloc[:, 2] == "").all():
            show_error(error_text, excel_error_column_new_names_empty)
            exit_if_directly_executed()

    



    ## METHODS TO MODIFY NAMES  
    # Ask method
    Option_selected = show_options(select_an_option_text, option1, option2, border1, border2, cancel_title)

    # Verifying Option_selected
    # print(f'The option selected is {Option_selected}')
    # if Option_selected == None:
    #     print("Cancel")
    # else:
    #     print(f"Option selected: {globals()[f'opcion{Option_selected}']}" if f'option{Option_selected}' in globals() else "Error: Option invalid.")
    
    
    # If cancel is given, the execution is aborted.   
    if Option_selected == None:
        exit_if_directly_executed()
    # If option 1 is given, the "create new folder" method will be used: 
    elif Option_selected == 1:
        # Create the folder and copy the files with new name:
        ruta_carpeta_destino = os.path.join(directory, folder_with_renamed_files)
        # Run function
        copy_files_with_new_names(df,directory,ruta_carpeta_destino, excel_error_file_not_found, excel_error_file_already_proceced, excel_error_file_already_exist_in_new_directory, excel_template_error_not_found, excel_template_error_not_allowed, excel_column_status)
    # If option 2 is given, the "overwriting files" method will be usde:
    elif Option_selected == 2:
        # Run function
        df = rename_files_locally(df, directory, excel_error_file_not_found, excel_error_file_already_proceced, excel_error_file_already_exist_in_the_same_folder, excel_template_error_not_found, excel_template_error_not_allowed, excel_column_status, start_function_title, start_function_message, operation_canceled_title, operation_canceled_message)
        folder_with_renamed_files = working_directory()
        
        
    
    # Another option will fail and abort the script.
    else:
        show_error(error_text, error_unknown)
        exit_if_directly_executed()
        




    ## CREATING THE LOG OF excel_column_error IN EXCEL
    # Filter the DataFrame so that excel_column_status is different from 'Ok'.
    df_error = df[df[excel_column_status] != 'Ok']

    # If len(df_error) = 0, show success message and exit
    if len(df_error)==0:
        show_message(success_title, success_message)
        exit_if_directly_executed()
        
    # If len(df_error) != 0 Then create excel with log of excel_column_error
    else: 
        # Rename column name excel_column_status to excel_column_error
        df_error = df_error.rename(columns={excel_column_status: excel_column_error})
        
        # Add df_conflict
        df_error = pd.concat([df_conflict, df_error], ignore_index=True)

        # Save the DataFrame in an Excel file (do not replace the last one):
        excel_error_log = excel_name + "_" + error_text
        extention = ".xlsx"
        counter = 1
        
        # Saving path
        excel_error_logs_path = os.path.join(directory, f"{excel_error_log} - {counter}{extention}")

        # Check if the file exists and add suffixes with numbers
        while os.path.exists(excel_error_logs_path):
            counter += 1
            excel_error_logs_path = os.path.join(directory, f"{excel_error_log} - {counter}{extention}")


        # Filter column 6 (excel_column_error) to make sure that column A and B are files in the folder
        filter = df_error[df_error.iloc[:, 5].str.startswith(excel_error_file_not_found)]
        # If filter_df and processed_df have the same number of rows => Column 'A' has names that do not correspond to the files in the folder:
        if len(filter) == len(df):
            # If the frame condition save the errors
            unlock_excel_sheet(excel_save_route)
            show_error(error_text, excel_error_files_not_found + error_excel_structure_modified)
            # Abort the execution
            exit_if_directly_executed()

            
        # Modify the values of column 6 (errors) to have the title of column 5 (new name at the end).
        df_error.iloc[:, 5] = df_error.apply(lambda row: row.iloc[5].replace(f"{excel_column_new_name_final}", str(row.iloc[4])) if isinstance(row.iloc[5], str) else row.iloc[5], axis=1)

        
        # Save the file in the final path
        df_error.to_excel(excel_error_logs_path, index=False)

        show_error(error_text, excel_errors_log_message + f'{excel_error_log} - {counter}{extention}')
        # Abort the execution
        exit_if_directly_executed()



if __name__ == "__main__":
    main(language)
    
    
