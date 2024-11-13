# DISABLE EXCEL SHEET PROTECTION
#Import libraries
import os
from openpyxl import load_workbook



# Import functions
from functions.functions import working_directory, show_error, show_message, load_translations,choose_language, exit_if_directly_executed
# Import parameters form config
from config.config import excel_name

# Import other scripts
from src.a_Create_excel import main as script0_main


language = choose_language()


# Declare the main function to transmit the language parameter when executing it in other scripts.
def main(language):
    """
    Unlocks the Excel sheet with the same name as the project or executable file, enabling full editing. The function uses the provided language parameter to handle any language-specific functionality.

    This function is located in src/c_Unlock_excel.py

    Arg:
        language (str): The language code (e.g., 'en', 'es') passed to customize language-specific content during the unlocking process.

    Returns:
        None: This function does not return any value. It modifies the Excel file by unlocking the sheet to allow full editing.
    """
    # Load translations
    traducciones = load_translations(language)


    error_text = traducciones['error_text']
    error_excel_not_found = traducciones['error_excel_not_found'].format(excel_name=excel_name)
    error_excel_open = traducciones['error_excel_open'].format(excel_name=excel_name)
    unlocked_title = traducciones['unlocked_title']
    unlocked_message = traducciones['unlocked_message']

    ## EXECUTION
    # Make sure you are in the working directory
    directory = working_directory()

    # Excel file path
    excel_name_ruta = os.path.join(directory, excel_name)


    ## DATAFRAME WORKING AND CLEANING
    # Check if the Excel file exists
    if not os.path.exists(excel_name_ruta):
        # If it doesn't exist
        # Displays an error message
        show_error(error_text, error_excel_not_found)
        # Run src/a_Create_excel.py
        script0_main(language)
        # Abort the execution
        exit_if_directly_executed()

    else:
        #print(excel_name_ruta)
        # If the file exists, I will open it to unprotect it.
        try:
            wb = load_workbook(excel_name_ruta)
        except PermissionError: # If the excel is open.
            show_error(error_text, error_excel_open)
            # Abort the execution
            exit_if_directly_executed()
            
            
        ws = wb.active  # Selects the first worksheet (the active one)
        ws.protection.sheet = False
        # Save the unprotected excel
        wb.save(excel_name_ruta)
        wb.close()
        
        # Show message
        show_message(unlocked_title, unlocked_message)


if __name__ == "__main__":
    main(language)