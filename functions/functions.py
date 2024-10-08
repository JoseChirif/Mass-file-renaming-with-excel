# Importar librerias para manejar carpetas y archivos
import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook
from openpyxl.styles import Font, Protection
import sys
import shutil
import webbrowser
import json



## FUNCIONES DIRIGIRSE A CARPETAS/DIRECTORIOS
#Obtner el direrctorio (utilizar directorio padre para crear el ejecutable, sino marcara error)
def directorio_a_trabajar():
    """
    Esta función obtiene el directorio padre del proyecto (3 carpetas padre a partir de functions.py).
    """
    if hasattr(sys, 'frozen'):
        # Si está ejecutándose como un ejecutable, obtenemos la carpeta del ejecutable
        directorio = os.path.dirname(sys.executable)
    else:
        # Si está ejecutándose como un script normal
        directorio = os.path.abspath(os.path.join(__file__, *(['..'] * 3)))
    
    return directorio  # Debe devolver el directorio



def directorio_proyecto():
    """
    Esta función obtiene el directorio del proyecto (carpetas padre a partir de functions.py).
    """
    if hasattr(sys, '_MEIPASS'):
        # Si está ejecutándose como un ejecutable, se usa el directorio temporal
        directorio_proyecto = sys._MEIPASS
    else:
        # Si está ejecutándose como un script normal
        directorio_proyecto = os.path.abspath(os.path.join(__file__, *(['..'] * 2)))
    
    return directorio_proyecto  # Debe devolver el directorio



def obtener_nombre_archivo_o_directorio():
    """Get the name of executable file or project directory 

    Returns:
        str: name of executable file or project directory
    """
    if getattr(sys, 'frozen', False):
        # Estás ejecutando desde un .exe
        return os.path.basename(sys.executable)
    else:
        # Si estás ejecutando desde un script Python obtiene el nombre del directorio
        return os.path.basename(directorio_proyecto()).rstrip("/").rstrip("\\")
    
    
## FUNCIONES SCRIPTS
def choose_language():
    # Obtener el idioma desde los argumentos de línea de comandos
    language = "es"  # Idioma por defecto
    if len(sys.argv) > 1:
        language = sys.argv[1]  # Tomar el argumento si se proporciona
        
    return language


        
        
def cargar_lista_de_lenguajes():
    # Obtener el directorio base del proyecto
    directorio_del_proyecto = directorio_proyecto() 
    # Cargar el diccionario de idiomas desde el archivo JSON
    ruta_json = os.path.join(directorio_del_proyecto, 'locales', 'languages.json')
    with open(ruta_json, 'r', encoding='utf-8') as archivo:
        return json.load(archivo)
    
    
def cargar_traducciones(idioma):
    """_summary_

    Args:
        idioma (_type_): _description_

    Returns:
        _type_: _description_
    """
    # Obtener el directorio base del proyecto
    directorio_del_proyecto = directorio_proyecto() 
    # Construir la ruta hacia el archivo JSON del idioma
    ruta_json = os.path.join(directorio_del_proyecto, 'locales', f'{idioma}.json')
    with open(ruta_json, 'r', encoding='utf-8') as archivo:
        traducciones = json.load(archivo)


    return traducciones


# Función para preguntar al usuario si desea reemplazar el archivo
def preguntar_reemplazo(file_already_exist_title, file_already_exist_message):
    """
    Si existe un archivo con el mismo nombre, preguntará si deseo reemplazarlo (tkinter).
    """
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal
    return messagebox.askyesno(file_already_exist_title, file_already_exist_message)

# Función para mostrar un mensaje de error
def mostrar_error(error_title ,error_message):
    """
    Muestra un mensaje de error con tkinter
    """
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal
    messagebox.showerror(error_title, error_message)
    

def mostrar_mensaje(titulo, mensaje):
    # Crear una ventana oculta
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal
    # Convertir el mensaje a string si es un DataFrame
    if isinstance(mensaje, pd.DataFrame):
        mensaje = mensaje.to_string()  # Convertir DataFrame a string
    # Mostrar el mensaje en una ventana emergente
    messagebox.showinfo(titulo, mensaje)
    # Cerrar la ventana después de que se cierre el mensaje
    root.destroy()



#Función para modificar el excel con el DF
def modificar_excel_dataframe(nombre_archivo_excel_ruta, filas_adicionales_a_desbloquear, notas_titulo, excel_note1, excel_note2, excel_note3, excel_note4, excel_note5):
    # Abrir el archivo Excel para mostrar una nota en la celda E2
    wb = load_workbook(nombre_archivo_excel_ruta)
    ws = wb.active  # Selecciona la primera hoja de trabajo (la activa)

    # Escribir notas en la columna E
    ws['E2'] = notas_titulo
    ws['E2'].font = Font(bold=True)  # Aplicar formato en negrita
    ws['E3'] = excel_note1
    ws['E4'] = excel_note2
    ws['E5'] = excel_note3
    ws['E6'] = excel_note4
    ws['E7'] = excel_note5

    # Protección de la hoja
    #1048576 filas = el excel default. Para asegurar desbloquear la hoja completa. Valido no pasar el límite
    if ws.max_row >=1048576-filas_adicionales_a_desbloquear: 
        ws.max_row =1048576-filas_adicionales_a_desbloquear
    # Desbloquear las filas del df + filas adicionales     
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row + filas_adicionales_a_desbloquear, min_col=1, max_col=16):
        for cell in row:
            cell.protection = Protection(locked=False)
            
    # Bloquear las celdas A1, B1 y C1
    for cell in ['A1', 'B1', 'C1']:
        ws[cell].protection = Protection(locked=True)

    # Habilitar la protección de la hoja
    ws.protection.sheet = True

    # Permitir ciertas acciones en la hoja
    ws.protection.sort = False           # Permitir ordenar
    ws.protection.autoFilter = False      # Permitir filtros automáticos
    ws.protection.insertRows = False      # Permitir insertar filas
    ws.protection.deleteRows = False      # Permitir eliminar filas
    ws.protection.insertColumns = False    # Permitir insertar columnas
    ws.protection.deleteColumns = False    # Permitir eliminar columnas
    ws.protection.formatColumns = False    # Permitir cambiar el formato de columnas
    ws.protection.formatRows = False       # Permitir cambiar el formato de filas



    # Guardar el archivo modificado
    wb.save(nombre_archivo_excel_ruta)
    wb.close()

    
    
#Función para modificar el excel con el DF
def desbloquear_proteccion_excel(nombre_archivo_excel_ruta):
    # Abrir el archivo Excel para mostrar una nota en la celda E2
    wb = load_workbook(nombre_archivo_excel_ruta)
    ws = wb.active  # Selecciona la primera hoja de trabajo (la activa)

    # Deshabilitar la protección de la hoja
    ws.protection.sheet = False

    # Guardar el archivo modificado
    wb.save(nombre_archivo_excel_ruta)
    wb.close()
    
  

  
def eliminar_desde_ultimo_punto(archivo):
    # Encontrar la posición del último punto
    indice_punto = archivo.rfind('.')
    
    # Si no hay punto en la cadena, devolver la cadena completa
    if indice_punto == -1:
        return archivo
    else:
        # Devolver la cadena desde el inicio hasta el último punto
        return archivo[:indice_punto]




# Función para mostrar las opciones
def mostrar_opciones(titulo, opcion1, opcion2, borde1, borde2, cancel_title):
    """
    Muestra una ventana de opciones con dos botones.
    Al seleccionar una opción, devuelve el número de la opción seleccionada (1 o 2).
    Si se cancela, cierra solo la ventana de opciones sin afectar otras ventanas.
    """
    def cancelar_operacion_actual():
        """
        Esta subfunción cancela la operación y cierra solo la ventana actual de opciones.
        """
        global ventana_opciones
        if ventana_opciones:
            ventana_opciones.destroy()  # Cierra la ventana de opciones
        
        
    def seleccionar_opcion(nro_opcion):
        """Sub-función para manejar la selección de la opción."""
        nonlocal resultado
        resultado = nro_opcion
        cancelar_operacion_actual()
        ventana_opciones.quit()  # Cierra la ventana de opciones

    global ventana_opciones
    ventana_opciones = tk.Toplevel()  # Crear una ventana secundaria
    ventana_opciones.title(titulo)  # Establecer el título de la ventana
    
   # Agregar el icono a la ventana
    ventana_opciones.iconbitmap(relative_route_to_file("assets", "icon.ico"))

    # Variable para almacenar el resultado de la selección
    resultado = None

    # Manejar el botón de cerrar (X) de la ventana
    ventana_opciones.protocol("WM_DELETE_WINDOW", cancelar_operacion_actual)

    # Marco para la primera opción
    frame1 = tk.Frame(ventana_opciones, highlightbackground=borde1, highlightthickness=2)
    frame1.pack(pady=20, padx=20)
    
    boton1 = tk.Button(frame1, text=opcion1, font=("Arial", 10), borderwidth=0, relief="flat",
                       bg="white", activebackground="lightgrey", 
                       padx=15, pady=10,  # Padding
                       command=lambda: seleccionar_opcion(1))
    boton1.pack()

    # Marco para la segunda opción
    frame2 = tk.Frame(ventana_opciones, highlightbackground=borde2, highlightthickness=2)
    frame2.pack(pady=20, padx=20)

    boton2 = tk.Button(frame2, text=opcion2, font=("Arial", 10), borderwidth=0, relief="flat",
                       bg="white", activebackground="lightgrey", 
                       padx=15, pady=10,  # Padding
                       command=lambda: seleccionar_opcion(2))
    boton2.pack()

    # Botón de cancelar
    boton_cancelar = tk.Button(ventana_opciones, text=cancel_title, font=("Arial", 10), command=cancelar_operacion_actual, 
                                padx=10, pady=5, bg="red", fg="white")
    boton_cancelar.pack(side=tk.BOTTOM, anchor=tk.SE, padx=10, pady=10)  # Esquina inferior derecha

    # Mantener la ventana abierta hasta que se seleccione una opción o se cierre
    ventana_opciones.mainloop()

    # Retornar el resultado al finalizar
    return resultado






    
#Función para crear una nueva carpeta y renombrar
def copiar_archivos_con_nuevo_nombre(DataFrame_a_procesar, carpeta_origen, carpeta_destino, excel_error_file_doesnt_found, excel_error_file_already_proceced, excel_error_file_already_exist, excel_template_error_doesnt_found, excel_template_error_not_allowed, excel_column_status):
    """
    Esta función crea una carpeta nueva (si no existe), verifica que no haya archivos con los nombres nuevos (si los hay los marcará en estados). 
    También verifica que entre los nombres nuevos no haya duplicados (también lo marcará en estados). 
    Finalmente copia los archivos de la carpeta origen y los pega con los nuevos nombres en la carpeta destino. 
    Todo se registra en estados y se devuelve en el df.
    """
    # Crear la carpeta de destino si no existe
    if not os.path.exists(carpeta_destino):
        os.makedirs(carpeta_destino)

    # Lista para guardar los estados
    estados = []
    nombres_procesados = set()  # Conjunto para rastrear nombres únicos de 'Nombre nuevo completo'

    # Iterar sobre el DataFrame para copiar archivos y carpetas
    for index, row in DataFrame_a_procesar.iterrows():
        nombre_original_completo = row.iloc[3]  # columna 4 = Archivo o carpeta de origen
        nombre_nuevo_completo = row.iloc[4]  # columna 5 = Nombre para el archivo o carpeta de destino

        # Ruta completa del archivo o carpeta de origen
        ruta_origen = os.path.join(carpeta_origen, nombre_original_completo)
        # Ruta completa para el archivo o carpeta de destino (en "Nombre modificado")
        ruta_destino = os.path.join(carpeta_destino, nombre_nuevo_completo)

        # Verificar si el archivo o carpeta existe en la carpeta de origen
        if not os.path.exists(ruta_origen):
            estados.append(excel_error_file_doesnt_found +" "+nombre_nuevo_completo)
            continue  # No procesar si no existe en origen

        # Verificar si el nombre nuevo ya ha sido procesado antes (es un duplicado en el DataFrame)
        if nombre_nuevo_completo in nombres_procesados:
            estados.append(excel_error_file_already_proceced)
            continue
        
        # Añadir el nombre nuevo al conjunto de nombres procesados
        nombres_procesados.add(nombre_nuevo_completo)

        # Comprobar si el archivo o carpeta ya existe en la carpeta de destino
        if os.path.exists(ruta_destino):
            # Si el archivo o carpeta ya existe, registrar el estado y continuar
            estados.append(excel_error_file_already_exist)
            continue

        # Intentar copiar el archivo o carpeta
        try:
            if os.path.isdir(ruta_origen):
                # Si es una carpeta, se copia toda la estructura de la carpeta
                shutil.copytree(ruta_origen, ruta_destino)
            else:
                # Si es un archivo, solo se copia el archivo
                shutil.copy(ruta_origen, ruta_destino)

            # Añadir estado 'Ok' si la copia fue exitosa
            estados.append("Ok")
        except FileNotFoundError:
            # Manejar el caso de archivo no encontrado
            estados.append(excel_template_error_doesnt_found)
        except PermissionError:
            # Manejar errores de permisos
            estados.append(excel_template_error_not_allowed)
        except Exception as e:
            # Manejo de errores en caso de otros problemas al copiar
            estados.append(f"Error: {e}")

    # Añadir la lista de estados al DataFrame como una nueva columna
    DataFrame_a_procesar[excel_column_status] = estados

    # Devolver el DataFrame con la nueva columna
    return DataFrame_a_procesar




# Función de confirmación usando tkinter
def preguntar_proceder_funcion(start_function_title, start_function_message):
    """
    Esta función pregunta al usuario si desea continuar con el renombrado de archivos.
    """
    ventana = tk.Tk()
    ventana.withdraw()  # Ocultar la ventana principal
    respuesta = messagebox.askyesno(start_function_title, start_function_message)
    ventana.destroy()
    return respuesta



# Función principal para renombrar archivos
def renombrar_archivos_local(DataFrame_a_procesar, directorio, excel_error_file_doesnt_found, excel_error_file_already_proceced, excel_error_file_already_exist, excel_template_error_doesnt_found, excel_template_error_not_allowed, excel_column_status, start_function_title, start_function_message, operation_canceled_title, operation_canceled_message):
    """
    Esta función renombra archivos localmente, utilizando la misma carpeta de origen y destino.
    Pregunta primero si el usuario desea proceder, y si no, cancela el proceso.
    """
    # Preguntar al usuario si desea proceder
    if not preguntar_proceder_funcion(start_function_title, start_function_message):
        # Si el usuario presiona "No", muestra un mensaje y retorna sin cambios
        mostrar_error(operation_canceled_title, operation_canceled_message)
        exit_if_directly_executed()  # Cancelar el script

    # Lista para guardar los estados
    estados = []
    nombres_procesados = set()  # Conjunto para rastrear nombres únicos de 'Nombre nuevo completo'

    # Iterar sobre el DataFrame para renombrar archivos y carpetas
    for index, row in DataFrame_a_procesar.iterrows():
        nombre_original_completo = row.iloc[3]  # columna 4 = Archivo o carpeta de origen
        nombre_nuevo_completo = row.iloc[4]  # columna 5 = Nombre para el archivo o carpeta de destino

        # Ruta completa del archivo o carpeta de origen
        ruta_origen = os.path.join(directorio, nombre_original_completo)
        # Ruta completa para el archivo o carpeta de destino (en "Nombre modificado")
        ruta_destino = os.path.join(directorio, nombre_nuevo_completo)

        # Verificar si el archivo o carpeta existe en la carpeta de origen
        if not os.path.exists(ruta_origen):
            estados.append(excel_error_file_doesnt_found)
            continue  # No procesar si no existe en origen

        # Verificar si el nombre nuevo ya ha sido procesado antes (es un duplicado en el DataFrame)
        if nombre_nuevo_completo in nombres_procesados:
            estados.append(excel_error_file_already_proceced)
            continue
        
        # Añadir el nombre nuevo al conjunto de nombres procesados
        nombres_procesados.add(nombre_nuevo_completo)

        # Comprobar si el archivo o carpeta ya existe con el nombre nuevo
        if os.path.exists(ruta_destino):
            # Si el archivo o carpeta ya existe, registrar el estado y continuar
            estados.append(excel_error_file_already_exist)
            continue

        # Intentar renombrar el archivo o carpeta
        try:
            os.rename(ruta_origen, ruta_destino)
            # Añadir estado 'Ok' si el renombrado fue exitoso
            estados.append("Ok")
        except FileNotFoundError:
            # Manejar el caso de archivo no encontrado
            estados.append(excel_template_error_doesnt_found)
        except PermissionError:
            # Manejar errores de permisos
            estados.append(excel_template_error_not_allowed)
        except Exception as e:
            # Manejo de errores en caso de otros problemas al renombrar
            estados.append(f"Error: {e}")

    # Añadir la lista de estados al DataFrame como una nueva columna
    DataFrame_a_procesar[excel_column_status] = estados

    # Devolver el DataFrame con la nueva columna
    return DataFrame_a_procesar


# Ajusta los textos de tkinter
def ajustar_texto(event, *args, margen):
    for label in args:
        # Ajustar el texto al ancho disponible menos un margen
        nuevo_wraplength = label.winfo_width() - margen
        if nuevo_wraplength > 0:  # Asegúrate de que el nuevo wraplength sea positivo
            label.config(wraplength=nuevo_wraplength)


def relative_route_to_file(path_to_folder, file):
    if hasattr(sys, '_MEIPASS'):
        route = os.path.join(sys._MEIPASS, path_to_folder, file)
    else:
        route = os.path.join(path_to_folder, file)
    
    return route




## FUNCIONES PARA MENU PRINCIPAL
# Función para abrir el repositorio
def open_web_page(link):
    webbrowser.open(link)







class DirectExecutionExit(Exception):
    """Excepción personalizada para manejar la salida del script sin cerrar el menú de Tkinter."""
    pass

def exit_if_directly_executed():
    """Lanza una excepción que puede ser capturada para detener el script."""
    raise DirectExecutionExit("El script fue detenido.")


## FUNCIONES PARA CORRER SCRIPTS
def ejecutar_script_src(script, idioma):
    """Ejecuta el script correspondiente y maneja el idioma."""
    try:
        # Intenta importar y ejecutar el script correspondiente
        module = __import__(f'src.{script}', fromlist=['main'])
        module.main(idioma)  # Llama a la función main pasando el idioma

    except DirectExecutionExit as e:
        print(f"El script fue detenido: {e}")
    except ImportError as e:
        print(f"Error al importar el script '{script}': {e}")
    except AttributeError as e:
        print(f"El script '{script}' no tiene una función 'main': {e}")
    except Exception as e:
        print(f"Ocurrió un error al ejecutar el script '{script}': {e}")



# Funciones específicas para cada script
def ejecutar_script0(idioma):
    ejecutar_script_src('a_Importar_archivos_a_excel', idioma)

def ejecutar_script1(idioma):
    ejecutar_script_src('b_Modificar_nombres', idioma)

def ejecutar_script2(idioma):
    ejecutar_script_src('c_Desbloquear_excel', idioma)

