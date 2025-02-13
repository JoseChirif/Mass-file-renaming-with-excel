#Import libraries
import tkinter as tk
from tkinter import ttk
import sys
import os
import subprocess
from PIL import Image, ImageTk

# go to the parent directory if you are running this script directly (uncomment the following lines)
# sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from functions.functions import open_web_page, load_available_languages, load_translations,  get_instructions, adjust_text, execute_script0, execute_script1, execute_script2

# Import parameters from config
from config.config import icon_picture_png, icon_picture_ico, logo_github_png, __version__

current_version = f'v{__version__}'

language = "en" # default language


# Load language names
names_languages = load_available_languages()
# Order them alphabetically for the menu
names_languages = dict(sorted(names_languages.items(), key=lambda item: item[1]))

# Invert the dictionary to facilitate selection
languages_names = {v: k for k, v in names_languages.items()}




# print(f'names_languages = {names_languages}')
# print(f'languages_names = {languages_names}')

translations = load_translations(language)

# Declaro textos de locales (load_translations)
project_title = translations["project_title"]
language_text = translations['language_text']
select_an_option_text = translations["select_an_option_text"]
text_button_1 = translations["text_button_1"]
text_button_2 = translations["text_button_2"]
text_button_3 = translations["text_button_3"]
notes_title = translations["notes_title"]
notes_content = translations["notes_content"]
instructions_text = translations["instructions_text"]
    

# Styles
background = '#F0F0F0'
button_padding = 2
button_spacing = 10
button_border_width = 2
button_color_border = "black"
padding_text_button_x = 20  # Horizontal padding between text and button border
padding_text_button_y = 2  # Vertical padding between text and button border
window_size = "650x650"   # Custom window size (width x height)
margin = 30        # left margin to align texts
minimum_window_width = 300
link_color = "#0770E0"





def select_language(event):
    """
    Handles the language selection in the application. This function is triggered when the user selects a language from the available options. It updates the interface and loads the corresponding language settings.

    This function is located in src/main_menu.py

    Arg:
        event (Event): The event object that is passed when the user selects a language from the menu or interface.

    Returns:
        None: This function does not return any value. It updates the language settings and refreshes the interface according to the selected language.
    """
    global language
    language_selected = combo_languages.get()  # Get the selected language from the Combobox
    language = languages_names[language_selected]  # Get language code from languages_names

    translations = load_translations(language)

    # Update texts with loaded translations
    global project_title, language_text, select_an_option_text
    global text_button_1, text_button_2, text_button_3, notes_title, notes_content

    project_title = translations["project_title"]
    language_text = translations['language_text']
    select_an_option_text = translations["select_an_option_text"]
    text_button_1 = translations["text_button_1"]
    text_button_2 = translations["text_button_2"]
    text_button_3 = translations["text_button_3"]
    notes_title = translations["notes_title"]
    notes_content = translations["notes_content"]
    instructions_text = translations["instructions_text"]
    # Update texts in GUI widgets
    lbl_project_title.config(text=project_title)
    lbl_languages.config(text=language_text)
    lbl_menu.config(text=select_an_option_text)
    btn_option_1.config(text=text_button_1)
    btn_option_2.config(text=text_button_2)
    btn_option_3.config(text=text_button_3)
    lbl_project_title_notes_content.config(text=notes_title)
    lbl_notes_content.config(text=notes_content)
    lbl_instructions.config(text=instructions_text)



    
    
# Function to create the main menu
def main_menu():
    """
    Displays the main menu for the project, allowing the user to navigate between different functionalities. 
    The function customizes the content and interface based on the provided language parameter.

    This function is located in src/main_menu.py

    Arg:
        language (str): The language code (e.g., 'en', 'es') passed to customize the language-specific content and labels on the menu.

    Returns:
        None: This function does not return any value. It opens the main menu interface and handles user interaction for navigation.
    """
    global window, combo_languages # Make 'window' and combo_languages global
    global lbl_project_title, lbl_languages, lbl_menu, language_text, btn_option_1, btn_option_2, btn_option_3, lbl_project_title_notes_content, lbl_notes_content, lbl_instructions  # Make the widgets global to update them in select_language..
    window = tk.Tk()
    window.title(project_title)
    
    # Set the minimum width
    window.wm_minsize(minimum_window_width, 0)
    
    # Set window size and background size
    window.geometry(window_size)
    window.configure(bg=background)

    # Make the window come to the foreground
    window.attributes("-topmost", True)
    window.after(1, lambda: window.attributes("-topmost", False))  # Return to normal state after opening
    
    # Add the icon to the window
    window.iconbitmap(icon_picture_ico)
    
    # Close the whole program when I close the menu
    window.protocol("WM_DELETE_WINDOW", lambda: (window.destroy(), sys.exit()))
    
    


    
    
    
    ## TITLE
    empty_space = tk.Frame(window, height=20)  # An empty frame
    empty_space.pack()
    # Create a frame for the project_title and the icon_picture_png image
    frame_project_title = tk.Frame(window, bg=background)
    frame_project_title.pack(pady=10, padx=margin, fill='x')

    # Load the icon_picture_png
    img = Image.open(icon_picture_png)
    img = img.resize((50, 50), Image.LANCZOS)  # Sets the icon size
    icon_picture_png_tk = ImageTk.PhotoImage(img)
    


    # Label for the icon_picture_png
    lbl_icon_picture_png = tk.Label(frame_project_title, image=icon_picture_png_tk, bg=background)
    lbl_icon_picture_png.image = icon_picture_png_tk  # Keep a reference to prevent deletion
    lbl_icon_picture_png.pack(side=tk.LEFT)

    # project_title
    lbl_project_title = tk.Label(frame_project_title, text=project_title, font=("Arial", 14, "bold"), bg=background)
    lbl_project_title.pack(fill='x', expand=True, padx=10)
    
    
    
    ## LANGUAGES
    # Create a frame to contain the Combobox
    frame_language = tk.Frame(window, bg=background)  # Create a frame with the window background
    # Label for “languages” in boldface type
    lbl_languages = tk.Label(frame_language, text=language_text, font=("Arial", 10, "bold"), bg=background, anchor="e")
    lbl_languages.pack(side=tk.LEFT, padx=(0, 5))  # Packing the label on the left with a right margin
    
    # Drop-down menu to select language using ttk.Combobox
    variable_language = tk.StringVar(window)
    variable_language.set(names_languages[language])  # Value with the full name of the language

    # Create a Combobox for the language to display the full names
    combo_languages = ttk.Combobox(frame_language, textvariable=variable_language, values=list(names_languages.values()))
    combo_languages.bind("<<ComboboxSelected>>", select_language)  # Calling the function when selecting a language
    # Packs the Combobox on the right with a specific margin
    combo_languages.pack(side=tk.RIGHT, padx=margin)

    # Pack the frame in the window
    frame_language.pack(side=tk.TOP, anchor='ne', padx=(margin, 10), pady=(2, 0))
    
    
    
    
 
    
    
    ## OPTIONS
    frame_options_and_notes = tk.Frame(window, bg=background)
    frame_options_and_notes.pack(expand=True, fill='both', pady=5)
    
    # Select an option text
    lbl_menu = tk.Label(frame_options_and_notes, text=select_an_option_text, font=("Arial", 10, "bold"), bg=background, anchor="w")
    lbl_menu.pack(expand=True, fill='both', pady=5, padx=margin, anchor="w")

    # Button 1: Create Excel
    btn_option_1 = tk.Button(frame_options_and_notes, text=text_button_1, font=("Arial", 10), command=lambda: execute_script0(language),
                            borderwidth=button_border_width, highlightbackground=button_color_border,
                            padx=padding_text_button_x, pady=padding_text_button_y)
    btn_option_1.pack(expand=True, fill='both', padx=margin+10, pady=(button_padding, button_spacing))

    # Button 2: Modify names
    btn_option_2 = tk.Button(frame_options_and_notes, text=text_button_2, font=("Arial", 10),command=lambda: execute_script1(language),
                            borderwidth=button_border_width, highlightbackground=button_color_border,
                            padx=padding_text_button_x, pady=padding_text_button_y)
    btn_option_2.pack(expand=True, fill='both', padx=margin+10, pady=(button_padding, button_spacing))

    # Button 3: Unlock Excel sheet
    btn_option_3 = tk.Button(frame_options_and_notes, text=text_button_3, font=("Arial", 10),command=lambda: execute_script2(language),
                            borderwidth=button_border_width, highlightbackground=button_color_border,
                            padx=padding_text_button_x, pady=padding_text_button_y)
    btn_option_3.pack(expand=True, fill='both', padx=margin+10, pady=(button_padding, button_spacing))

    # Notes title text
    lbl_project_title_notes_content = tk.Label(frame_options_and_notes, text=notes_title, font=("Arial", 10, "bold"), bg=background, anchor="w")
    lbl_project_title_notes_content.pack(expand=True, fill='both', pady=(0, 0), padx=margin, anchor="n")

    # notes content
    lbl_notes_content = tk.Label(frame_options_and_notes, text=notes_content, font=("Arial", 10), 
                                justify="left", bg=background, anchor="n")
    lbl_notes_content.pack(expand=True, fill='both', padx=margin, pady=(0, 0))







    ## REPOSITORY
    # Create a frame to align the icon_picture_png and the license
    frame_github = tk.Frame(window, bg=background)
    frame_github.place(relx=1.0, rely=0.0, anchor="se", x=-margin + 15, y=38)

    # Github text
    lbl_github_text = tk.Label(frame_github, text="GitHub", font=("Bell MT", 14), bg=background, fg=link_color, cursor="hand2") 
    lbl_github_text.pack(side=tk.LEFT, padx=(100, 0), pady=(5, 0))
    
    # Repository link (with the text)
    lbl_github_text.bind("<Button-1>", lambda e: open_web_page('https://github.com/JoseChirif/Mass-file-renaming-with-excel','https://github.com/JoseChirif?tab=repositories', 'https://github.com/JoseChirif'))
    
    #PNG    
    # Load Github icon
    img_logo_github = Image.open(logo_github_png)
    img_logo_github = img_logo_github.resize((30, 30), Image.LANCZOS)
    icon_picture_png_Link_repositorio_tk = ImageTk.PhotoImage(img_logo_github)

    # Label for the icon_picture_png of Link_repository
    lbl_github_logo = tk.Label(frame_github, image=icon_picture_png_Link_repositorio_tk, bg=background, cursor="hand2")  # Hand cursor
    lbl_github_logo.image = icon_picture_png_Link_repositorio_tk  
    lbl_github_logo.pack(side=tk.LEFT, padx=(0, 10))  # Add padding on the right side

    # Repository link (with the icon)
    lbl_github_logo.bind("<Button-1>", lambda e: open_web_page('https://github.com/JoseChirif/Mass-file-renaming-with-excel','https://github.com/JoseChirif?tab=repositories', 'https://github.com/JoseChirif'))
    
    



    ## LICENSE
    #frame_github = tk.Frame(window, bg=background)
    #frame_github.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-5)
    
    def get_license_route():
        """
        Returns the absolute path to the LICENSE.txt file located in the project's parent directory.

        This function is located in src/main_menu.py

        Returns:
            str: The absolute path to LICENSE.txt, allowing other parts of the application to access the license file regardless of the current working directory.
        """
        return os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'LICENSE'))
        
    def open_license():
        """
        Opens the LICENSE file in Notepad, regardless of its extension.
        """
        license_path = get_license_route()
        try:
            subprocess.run(["notepad", license_path], check=True)
        except FileNotFoundError:
            print("Notepad not found. Please ensure it is installed.")
        except Exception as e:
            print(f"An error occurred while trying to open the license file: {e}")


    # Show "MIT License" 
    lbl_license = tk.Label(window, text="MIT License", font=("Arial", 9), fg=link_color, cursor="hand2")    
    lbl_license.place(relx=1.0, rely=1.0, anchor="se", x=-margin, y=-5)
    # Llamar a open_license al hacer clic
    lbl_license.bind("<Button-1>", lambda e: open_license())


    ## Version
    lbl_version = tk.Label(window, text=current_version, font=("Arial", 9), fg="black")    
    lbl_version.place(relx=1.0, rely=1.0, anchor="se", x=-margin, y=-25)


    ## Instructions
    lbl_instructions = tk.Label(window, text=instructions_text, font=("Arial", 9, "bold"), fg="black", cursor="hand2", bg=window.cget('bg'), relief="raised", bd=2, padx=5, pady=5)

    lbl_instructions.place(relx=0.0, rely=1.0, anchor="sw", x=margin-10, y=-5)
    # Open the instructions page by clicking on
    lbl_instructions.bind("<Button-1>", lambda e: get_instructions(language))
    
    
    
    
    
    ## FINAL SETTINGS
    window.bind("<Configure>", lambda event: adjust_text(event, lbl_project_title, lbl_menu, btn_option_1, btn_option_2, btn_option_3, lbl_project_title_notes_content, lbl_notes_content, margin=20))
    
    window.mainloop()





    
    
# Call main_menu
if __name__ == "__main__":
    main_menu()
    
    
