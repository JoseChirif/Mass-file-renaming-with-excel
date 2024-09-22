#Importo librerias
import tkinter as tk
from tkinter import messagebox
import sys
import os
from PIL import Image, ImageTk


# Importo funciones a usar
sys.path.append(os.path.abspath(os.path.dirname(__file__) + '/../'))  # Me dirije al directorio del proyecto
from functions.functions import ejecutar_script_0_Importar_archivos_a_excel, ejecutar_script_1_Modificar_nombres, ejecutar_script_2_Desbloquear_excel, abrir_github


# Textos y funciones
titulo = "Renombrar archivos masivamente"
imagen = "icon/icon.png"
encabezado = "Renombrar archivos masivamente"
seleccionar_opcion = "Seleccione una opción:"
texto_boton_1= "1. Crear Excel con los nombres de los archivos en esta carpeta."
accion_boton_1 = ejecutar_script_0_Importar_archivos_a_excel
texto_boton_2= "2. Modificar nombres con la plantilla excel."
accion_boton_2 = ejecutar_script_1_Modificar_nombres
texto_boton_3= "Desbloquear hoja excel (opcional)"
accion_boton_3 = ejecutar_script_2_Desbloquear_excel
notas = "- Se recomienda trabajar dentro de una carpeta creada. \n- No trabajar directamente sobre el escritorio o disco duro (Ej: C, D). \n- El usuario es responsable de los archivos que modifica."

# Estilos
padding_botones = 10  # Espacio alrededor de los botones
ancho_botones = 50    # Ancho de los botones
espaciado_botones = 10  # Espaciado vertical entre botones
ancho_bordes_botones = 2   # Ancho del borde de los botones
color_borde_botones = "black"  # Color del borde de los botones
padding_texto_boton_x = 20  # Padding horizontal entre texto y borde del botón
padding_texto_boton_y = 10  # Padding vertical entre texto y borde del botón
tamano_ventana = "600x500"   # Tamaño personalizado de la ventana (ancho x alto)
margen_izquierdo = 30        # Margen izquierdo para alinear los textos


# Función para crear la interfaz gráfica
def menu_principal():
    global ventana  # Hacer que 'ventana' sea global
    ventana = tk.Tk()
    ventana.title(titulo)
    
    # Configurar el tamaño de la ventana y el fondo blanco
    ventana.geometry(tamano_ventana)
    ventana.configure(bg="white")

    # Hacer que la ventana salga en primer plano
    ventana.attributes("-topmost", True)
    ventana.after(1, lambda: ventana.attributes("-topmost", False))  # Volver a estado normal después de abrir
    
    # Añadirle el icono
    icon = "icon/icon.ico"
    # Agregar el icono a la ventana
    icon = ventana.iconbitmap(icon)

    # Crear un frame para el encabezado y la imagen
    frame_encabezado = tk.Frame(ventana, bg="white")
    frame_encabezado.pack(pady=10)

    # Cargar la imagen
    img = Image.open(imagen)
    img = img.resize((50, 50), Image.LANCZOS)  # Cambia el tamaño según sea necesario
    imagen_tk = ImageTk.PhotoImage(img)

    # Etiqueta para la imagen
    lbl_imagen = tk.Label(frame_encabezado, image=imagen_tk, bg="white")
    lbl_imagen.image = imagen_tk  # Mantener una referencia para evitar que se elimine
    lbl_imagen.pack(side=tk.LEFT)

    # Encabezado
    lbl_encabezado = tk.Label(frame_encabezado, text=encabezado, font=("Arial", 14, "bold"), bg="white")
    lbl_encabezado.pack(side=tk.LEFT, padx=10)  # Espaciado horizontal

    # Menú de opciones
    lbl_menu = tk.Label(ventana, text=seleccionar_opcion, font=("Arial", 10), bg="white", anchor="w")
    lbl_menu.pack(pady=5, padx=margen_izquierdo, anchor="w")

    # Botón 1: Crear Excel
    btn_opcion_1 = tk.Button(ventana, text=texto_boton_1, command=accion_boton_1, width=ancho_botones,
                             borderwidth=ancho_bordes_botones, highlightbackground=color_borde_botones,
                             padx=padding_texto_boton_x, pady=padding_texto_boton_y)
    btn_opcion_1.pack(pady=(padding_botones, espaciado_botones))

    # Botón 2: Modificar nombres
    btn_opcion_2 = tk.Button(ventana, text=texto_boton_2, command=accion_boton_2, width=ancho_botones,
                             borderwidth=ancho_bordes_botones, highlightbackground=color_borde_botones,
                             padx=padding_texto_boton_x, pady=padding_texto_boton_y)
    btn_opcion_2.pack(pady=espaciado_botones)

    # Botón 3: Desbloquear hoja Excel
    btn_opcion_3 = tk.Button(ventana, text=texto_boton_3, command=accion_boton_3, width=ancho_botones,
                             borderwidth=ancho_bordes_botones, highlightbackground=color_borde_botones,
                             padx=padding_texto_boton_x, pady=padding_texto_boton_y)
    btn_opcion_3.pack(pady=espaciado_botones)

    # Etiqueta "Notas:" en negrita antes de las notas
    lbl_titulo_notas = tk.Label(ventana, text="Notas:", font=("Arial", 10, "bold"), bg="white", anchor="w")
    lbl_titulo_notas.pack(pady=(20, 5), padx=margen_izquierdo, anchor="w")

    # Notas
    lbl_notas = tk.Label(ventana, text=notas, font=("Arial", 10), justify="left", bg="white", anchor="w")
    lbl_notas.pack(padx=margen_izquierdo, anchor="w")

    # Agregar espacio en blanco
    tk.Label(ventana, text="", bg="white").pack(pady=2)
    

    # Cargar la imagen de link a github
    img_Link_repositorio = Image.open("icon/Link_repositorio.jpg")
    img_Link_repositorio = img_Link_repositorio.resize((95, 50), Image.LANCZOS)
    imagen_Link_repositorio_tk = ImageTk.PhotoImage(img_Link_repositorio)

    # Etiqueta para la imagen de Link_repositorio
    lbl_imagen_Link_repositorio = tk.Label(ventana, image=imagen_Link_repositorio_tk, bg="white", cursor="hand2")  # Cursor como mano
    lbl_imagen_Link_repositorio.image = imagen_Link_repositorio_tk  # Mantener una referencia para evitar que se elimine
    lbl_imagen_Link_repositorio.pack(pady=10)  # Puedes ajustar el padding según sea necesario

    # Vínculo a Google
    lbl_imagen_Link_repositorio.bind("<Button-1>", abrir_github)
    
    ventana.mainloop()
    


# Llamada para crear la interfaz gráfica
if __name__ == "__main__":
    menu_principal()