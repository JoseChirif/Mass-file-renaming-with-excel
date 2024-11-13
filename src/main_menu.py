#Importo librerias
import tkinter as tk
from tkinter import ttk
import sys
import os
import subprocess
from PIL import Image, ImageTk


from functions.functions import open_web_page, load_available_languages, load_translations,  get_instructions, adjust_text, execute_script0, execute_script1, execute_script2

# Importo parámetros de config
from config.config import icon_picture_png, icon_picture_ico, logo_github_png, __version__

current_version = f'v{__version__}'

idioma = "en" # Idioma predeterminado


# Cargar los nombres de los idiomas
nombres_idiomas = load_available_languages()
# Los ordeno alfabeticamente para el menú
nombres_idiomas = dict(sorted(nombres_idiomas.items(), key=lambda item: item[1]))

# Invertir el diccionario para facilitar la selección
idiomas_nombres = {v: k for k, v in nombres_idiomas.items()}




# print(f'nombres_idiomas = {nombres_idiomas}')
# print(f'idiomas_nombres = {idiomas_nombres}')

traducciones = load_translations(idioma)

# Declaro textos de locales (load_translations)
project_title = traducciones["project_title"]
language_text = traducciones['language_text']
select_an_option_text = traducciones["select_an_option_text"]
text_button_1 = traducciones["text_button_1"]
text_button_2 = traducciones["text_button_2"]
text_button_3 = traducciones["text_button_3"]
notes_title = traducciones["notes_title"]
notes_content = traducciones["notes_content"]
instructions_text = traducciones["instructions_text"]
    

# Estilos
fondo_ventana = '#F0F0F0'
padding_botones = 2  # Espacio alrededor de los botones
espaciado_botones = 10  # Espaciado vertical entre botones
ancho_bordes_botones = 2   # Ancho del borde de los botones
color_borde_botones = "black"  # Color del borde de los botones
padding_text_button_x = 20  # Padding horizontal entre texto y borde del botón
padding_text_button_y = 2  # Padding vertical entre texto y borde del botón
tamano_ventana = "650x650"   # Tamaño personalizado de la ventana (ancho x alto)
margin = 30        # margin izquierdo para alinear los textos
ancho_minimo_ventana = 300
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
    global idioma
    idioma_seleccionado = combo_idiomas.get()  # Obtener el idioma seleccionado del Combobox
    idioma = idiomas_nombres[idioma_seleccionado]  # Obtener el código del idioma desde idiomas_nombres

    traducciones = load_translations(idioma)

    # Actualizar los textos con las traducciones cargadas
    global project_title, language_text, select_an_option_text
    global text_button_1, text_button_2, text_button_3, notes_title, notes_content

    project_title = traducciones["project_title"]
    language_text = traducciones['language_text']
    select_an_option_text = traducciones["select_an_option_text"]
    text_button_1 = traducciones["text_button_1"]
    text_button_2 = traducciones["text_button_2"]
    text_button_3 = traducciones["text_button_3"]
    notes_title = traducciones["notes_title"]
    notes_content = traducciones["notes_content"]
    instructions_text = traducciones["instructions_text"]
    # Actualizar los textos en los widgets de la interfaz gráfica
    lbl_project_title.config(text=project_title)
    etiqueta_idiomas.config(text=language_text)
    lbl_menu.config(text=select_an_option_text)
    btn_opcion_1.config(text=text_button_1)
    btn_opcion_2.config(text=text_button_2)
    btn_opcion_3.config(text=text_button_3)
    lbl_project_title_notes_content.config(text=notes_title)
    lbl_notes_content.config(text=notes_content)
    lbl_instructions.config(text=instructions_text)



    
    
# Función para crear la interfaz gráfica
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
    global ventana, combo_idiomas # Hacer que 'ventana' y combo_idiomas sea global
    global lbl_project_title, etiqueta_idiomas, lbl_menu, language_text, btn_opcion_1, btn_opcion_2, btn_opcion_3, lbl_project_title_notes_content, lbl_notes_content, lbl_instructions  # Hacer los widgets globales para actualizarlos en select_language.
    ventana = tk.Tk()
    ventana.title(project_title)
    
    # Establecer el ancho mínimo
    ventana.wm_minsize(ancho_minimo_ventana, 0)
    
    # Configurar el tamaño de la ventana y el fondo
    ventana.geometry(tamano_ventana)
    ventana.configure(bg=fondo_ventana)

    # Hacer que la ventana salga en primer plano
    ventana.attributes("-topmost", True)
    ventana.after(1, lambda: ventana.attributes("-topmost", False))  # Volver a estado normal después de abrir
    
    # Agregar el icono a la ventana
    ventana.iconbitmap(icon_picture_ico)
    
    # Cerrar el programa completo cuando cierro el menu
    ventana.protocol("WM_DELETE_WINDOW", lambda: (ventana.destroy(), sys.exit()))
    
    


    
    
    
    ## TITULO
    espacio_vacio = tk.Frame(ventana, height=20)  # Un Frame vacío
    espacio_vacio.pack()
    # Crear un frame para el project_title y la imagen icon_picture_png
    frame_project_title = tk.Frame(ventana, bg=fondo_ventana)
    frame_project_title.pack(pady=10, padx=margin, fill='x')

    # Cargar la icon_picture_png
    img = Image.open(icon_picture_png)
    img = img.resize((50, 50), Image.LANCZOS)  # Cambia el tamaño según sea necesario
    icon_picture_png_tk = ImageTk.PhotoImage(img)
    


    # Etiqueta para la icon_picture_png
    lbl_icon_picture_png = tk.Label(frame_project_title, image=icon_picture_png_tk, bg=fondo_ventana)
    lbl_icon_picture_png.image = icon_picture_png_tk  # Mantener una referencia para evitar que se elimine
    lbl_icon_picture_png.pack(side=tk.LEFT)

    # project_title
    lbl_project_title = tk.Label(frame_project_title, text=project_title, font=("Arial", 14, "bold"), bg=fondo_ventana)
    lbl_project_title.pack(fill='x', expand=True, padx=10)
    
    
    
    ## IDIOMA
    # Crear un marco para contener el Combobox
    frame_idioma = tk.Frame(ventana, bg=fondo_ventana)  # Crea un frame con el fondo de la ventana
    # Etiqueta para "IDIOMAS" en negrita
    etiqueta_idiomas = tk.Label(frame_idioma, text=language_text, font=("Arial", 10, "bold"), bg=fondo_ventana, anchor="e")
    etiqueta_idiomas.pack(side=tk.LEFT, padx=(0, 5))  # Empaquetar la etiqueta a la izquierda con un margin a la derecha
    
    # Menú desplegable para seleccionar idioma usando ttk.Combobox
    variable_idioma = tk.StringVar(ventana)
    variable_idioma.set(nombres_idiomas[idioma])  # Valor inicial con el nombre completo del idioma

    # Crear un Combobox para el idioma que muestre los nombres completos
    combo_idiomas = ttk.Combobox(frame_idioma, textvariable=variable_idioma, values=list(nombres_idiomas.values()))
    combo_idiomas.bind("<<ComboboxSelected>>", select_language)  # Llamar a la función al seleccionar un idioma
    # Empaqueta el Combobox a la derecha con un margin específico
    combo_idiomas.pack(side=tk.RIGHT, padx=margin)

    # Empaquetar el frame en la ventana
    frame_idioma.pack(side=tk.TOP, anchor='ne', padx=(margin, 10), pady=(2, 0))
    
    
    
    
 
    
    
    ## OPCIONES
    frame_opciones_y_notas = tk.Frame(ventana, bg=fondo_ventana)
    frame_opciones_y_notas.pack(expand=True, fill='both', pady=5)
    
    # Menú de opciones
    lbl_menu = tk.Label(frame_opciones_y_notas, text=select_an_option_text, font=("Arial", 10, "bold"), bg=fondo_ventana, anchor="w")
    lbl_menu.pack(expand=True, fill='both', pady=5, padx=margin, anchor="w")

    # Botón 1: Crear Excel
    btn_opcion_1 = tk.Button(frame_opciones_y_notas, text=text_button_1, font=("Arial", 10), command=lambda: execute_script0(idioma),
                            borderwidth=ancho_bordes_botones, highlightbackground=color_borde_botones,
                            padx=padding_text_button_x, pady=padding_text_button_y)
    btn_opcion_1.pack(expand=True, fill='both', padx=margin+10, pady=(padding_botones, espaciado_botones))

    # Botón 2: Modificar nombres
    btn_opcion_2 = tk.Button(frame_opciones_y_notas, text=text_button_2, font=("Arial", 10),command=lambda: execute_script1(idioma),
                            borderwidth=ancho_bordes_botones, highlightbackground=color_borde_botones,
                            padx=padding_text_button_x, pady=padding_text_button_y)
    btn_opcion_2.pack(expand=True, fill='both', padx=margin+10, pady=(padding_botones, espaciado_botones))

    # Botón 3: Desbloquear hoja Excel
    btn_opcion_3 = tk.Button(frame_opciones_y_notas, text=text_button_3, font=("Arial", 10),command=lambda: execute_script2(idioma),
                            borderwidth=ancho_bordes_botones, highlightbackground=color_borde_botones,
                            padx=padding_text_button_x, pady=padding_text_button_y)
    btn_opcion_3.pack(expand=True, fill='both', padx=margin+10, pady=(padding_botones, espaciado_botones))

    # Etiqueta notes_title en negrita antes de las notes_content
    lbl_project_title_notes_content = tk.Label(frame_opciones_y_notas, text=notes_title, font=("Arial", 10, "bold"), bg=fondo_ventana, anchor="w")
    lbl_project_title_notes_content.pack(expand=True, fill='both', pady=(0, 0), padx=margin, anchor="n")

    # notes_content
    lbl_notes_content = tk.Label(frame_opciones_y_notas, text=notes_content, font=("Arial", 10), 
                                justify="left", bg=fondo_ventana, anchor="n")
    lbl_notes_content.pack(expand=True, fill='both', padx=margin, pady=(0, 0))







    ## LINK AL REPOSITORIO Y LICENCIA
    # Crear un frame para alinear la icon_picture_png y la licencia
    frame_github = tk.Frame(ventana, bg=fondo_ventana)
    frame_github.place(relx=1.0, rely=0.0, anchor="se", x=-margin + 15, y=38)

    #texto github
    lbl_github_text = tk.Label(frame_github, text="GitHub", font=("Bell MT", 14), bg=fondo_ventana, fg=link_color, cursor="hand2") 
    lbl_github_text.pack(side=tk.LEFT, padx=(100, 0), pady=(5, 0))
    
    # Vínculo al repositorio
    lbl_github_text.bind("<Button-1>", lambda e: open_web_page('https://github.com/JoseChirif/Mass-file-renaming-with-excel','https://github.com/JoseChirif?tab=repositories', 'https://github.com/JoseChirif'))
    
    #PNG    
    # Cargar la icon_picture_png de link a GitHub
    img_logo_github = Image.open(logo_github_png)
    img_logo_github = img_logo_github.resize((30, 30), Image.LANCZOS)
    icon_picture_png_Link_repositorio_tk = ImageTk.PhotoImage(img_logo_github)

    # Etiqueta para la icon_picture_png de Link_repositorio
    lbl_github_logo = tk.Label(frame_github, image=icon_picture_png_Link_repositorio_tk, bg=fondo_ventana, cursor="hand2")  # Cursor como mano
    lbl_github_logo.image = icon_picture_png_Link_repositorio_tk  
    lbl_github_logo.pack(side=tk.LEFT, padx=(0, 10))  # Agrega un poco de padding a la derecha

    # Vínculo al repositorio
    lbl_github_logo.bind("<Button-1>", lambda e: open_web_page('https://github.com/JoseChirif/Mass-file-renaming-with-excel','https://github.com/JoseChirif?tab=repositories', 'https://github.com/JoseChirif'))
    
    



    ## LICENCIA
    #frame_github = tk.Frame(ventana, bg=fondo_ventana)
    #frame_github.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-5)
    
    def obtener_ruta_license():
        # Obtiene la ruta completa al archivo LICENSE.txt en el directorio padre
        return os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'LICENSE.txt'))
    # Función para leer licencia
    def leer_license():
        # Abre LICENSE.txt en el Bloc de notes_content
        ruta_license = obtener_ruta_license()
        subprocess.Popen(['notepad.exe', ruta_license])

    # Mostrar "MIT License" 
    lbl_licencia = tk.Label(ventana, text="MIT License", font=("Arial", 9), fg=link_color, cursor="hand2")    
    lbl_licencia.place(relx=1.0, rely=1.0, anchor="se", x=-margin, y=-5)
    # Llamar a leer_license al hacer clic
    lbl_licencia.bind("<Button-1>", lambda e: leer_license())


    ## Version
    lbl_version = tk.Label(ventana, text=current_version, font=("Arial", 9), fg="black")    
    lbl_version.place(relx=1.0, rely=1.0, anchor="se", x=-margin, y=-25)


    ## Instructions
    lbl_instructions = tk.Label(ventana, text=instructions_text, font=("Arial", 9, "bold"), fg="black", cursor="hand2", bg=ventana.cget('bg'), relief="raised", bd=2, padx=5, pady=5)

    lbl_instructions.place(relx=0.0, rely=1.0, anchor="sw", x=margin-10, y=-5)
    # Abrir la pagina de instruciones al hacer click
    lbl_instructions.bind("<Button-1>", lambda e: get_instructions(idioma))
    
    
    
    
    
    ## CONFIGURACIONES FINALES
    ventana.bind("<Configure>", lambda event: adjust_text(event, lbl_project_title, lbl_menu, btn_opcion_1, btn_opcion_2, btn_opcion_3, lbl_project_title_notes_content, lbl_notes_content, margin=20))
    
    ventana.mainloop()





    
    
# Llamada para crear la interfaz gráfica   
if __name__ == "__main__":
    main_menu()
    
    
