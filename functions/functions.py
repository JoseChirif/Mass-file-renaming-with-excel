# Importar librerias para manejar carpetas y archivos
import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook
from openpyxl.styles import Font, Protection
import sys
import subprocess
import shutil
import webbrowser




## FUNCIONES QUE USARE EN MIS SCRIPTS
#Obtner el direrctorio (utilizar directorio padre para crear el ejecutable, sino marcara error)
def directorio_a_trabajar():
    """
    Esta función obtiene el directorio padre.
    En el código, está comentado la linea que obtiene el directorio padre, ya que para las pruebas se hagan en la misma carpeta.
    """
    # Obtener la ruta del directorio actual
    # directorio = os.getcwd()
    # Obtener la ruta del directorio padre del proyecto
    directorio = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    #Obtener directorio final
    return directorio  # Debe devolver el directorio

# Función para preguntar al usuario si desea reemplazar el archivo
def preguntar_reemplazo(archivo):
    """
    Si existe un archivo con el mismo nombre, preguntará si deseo reemplazarlo (tkinter).
    """
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal
    return messagebox.askyesno("Archivo existente", f"El archivo '{archivo}' ya existe. ¿Deseas reemplazarlo?")

# Función para mostrar un mensaje de error
def mostrar_error(mensaje):
    """
    Muestra un mensaje de error con tkinter
    """
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal
    messagebox.showerror("Error", mensaje)
    

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
def modificar_excel_dataframe(nombre_archivo_excel_ruta, columna_nombre_nuevo_excel_inicial, nombre_carpeta_destino, filas_adicionales_a_desbloquear):
        # Abrir el archivo Excel para mostrar una nota en la celda E2
    wb = load_workbook(nombre_archivo_excel_ruta)
    ws = wb.active  # Selecciona la primera hoja de trabajo (la activa)

    # Escribir notas en la columna E
    ws['E2'] = "Notas:"
    ws['E2'].font = Font(bold=True)  # Aplicar formato en negrita
    ws['E3'] = f" - La columna C ('{columna_nombre_nuevo_excel_inicial}') NO requiere extensión (Ej: .jpg .png ...), se respetarán las extensiones de la columna B."
    ws['E4'] = f" - Si deseas mantener el nombre original, deja la celda de la columna C ('{columna_nombre_nuevo_excel_inicial}') en blanco."
    ws['E5'] = f" - Mantener la estructura del Excel (No modificar las celdas A1, B1 y C1; ni insertar columnas antes de la columna C."
    ws['E6'] = f" - No incluir el ejecutable que modifica nombres ni la carpeta '{nombre_carpeta_destino}' (donde opcionalmente se almacenan las modificaciones). Y en caso estes ejecutando desde la carpeta con Python, no incluir la misma carpeta del proyecto."

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
def desbloquear_protección_excel(nombre_archivo_excel_ruta):
    # Abrir el archivo Excel para mostrar una nota en la celda E2
    wb = load_workbook(nombre_archivo_excel_ruta)
    ws = wb.active  # Selecciona la primera hoja de trabajo (la activa)

    # Deshabilitar la protección de la hoja
    ws.protection.sheet = False

    # Guardar el archivo modificado
    wb.save(nombre_archivo_excel_ruta)
    wb.close()
    
  
# Función para obtener el icono
def obtener_icono():
    """
    Esta función añade el icono a las ventanas de tkinter.
    """
    # Voy al directorio raíz del proyecto
    sys.path.append(os.path.abspath(os.path.dirname(__file__) + '/../'))
    # Obtengo el icono
    icon = "icon/icon.ico"
    # Agregar el icono a la ventana
    icon = ventana.iconbitmap(icon)
    
    return icon
      

    
    
# Función para cancelar la operación
def cancelar_operacion():
    """
    Esta función cancela la operación y cierra la ventana principal.
    """
    global ventana
    # Variable global para la ventana
    ventana = None
    
    if ventana:
        ventana.destroy()  # Cierra la ventana principal
    sys.exit()  # Detener la ejecución del programa
    


# Función para seleccionar una de 2 opciones
def mostrar_opciones(titulo, opcion1, opcion2, borde1, borde2):
    """
    Esta función brinda 2 opciones, cada una en un botón con el borde asignado (borde1 y borde2 respectivamente).
    Al final, con la sub-función seleccionar_opcion() devuelve nro 1 ó 2 según la selección o 'cancelar' si se cancela.
    """
    def seleccionar_opcion(nro_opcion):
        """Sub-función para brindar el nro de opción"""
        nonlocal resultado
        resultado = nro_opcion
        ventana.quit()  # Cerrar la ventana

    global ventana
    ventana = tk.Tk()
    ventana.title(titulo)  # Establecer el título de la ventana
    
    # Agregar el icono a la ventana
    obtener_icono()

    # Variable para almacenar el resultado
    resultado = None

    # Botón de cerrar (X)
    ventana.protocol("WM_DELETE_WINDOW", cancelar_operacion)

    # Marco para la primera opción
    frame1 = tk.Frame(ventana, highlightbackground=borde1, highlightthickness=2)
    frame1.pack(pady=20, padx=20)
    
    boton1 = tk.Button(frame1, text=opcion1, borderwidth=0, relief="flat",
                       bg="white", activebackground="lightgrey", 
                       padx=15, pady=10,  # Padding
                       command=lambda: seleccionar_opcion(1))
    boton1.pack()

    # Marco para la segunda opción
    frame2 = tk.Frame(ventana, highlightbackground=borde2, highlightthickness=2)
    frame2.pack(pady=20, padx=20)

    boton2 = tk.Button(frame2, text=opcion2, borderwidth=0, relief="flat",
                       bg="white", activebackground="lightgrey", 
                       padx=15, pady=10,  # Padding
                       command=lambda: seleccionar_opcion(2))
    boton2.pack()

    # Botón de cancelar
    boton_cancelar = tk.Button(ventana, text="Cancelar", command=cancelar_operacion, 
                                padx=10, pady=5, bg="red", fg="white")
    boton_cancelar.pack(side=tk.BOTTOM, anchor=tk.SE, padx=10, pady=10)  # Esquina inferior derecha

    ventana.mainloop()

    # Retornamos el resultado al finalizar
    return resultado


    
#Función para crear una nueva carpeta y renombrar
def copiar_archivos_con_nuevo_nombre(DataFrame_a_procesar, carpeta_origen, carpeta_destino):
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
        nombre_original_completo = row['Nombre original completo']  # Archivo o carpeta de origen
        nombre_nuevo_completo = row['Nombre nuevo completo']  # Nombre para el archivo o carpeta de destino

        # Ruta completa del archivo o carpeta de origen
        ruta_origen = os.path.join(carpeta_origen, nombre_original_completo)
        # Ruta completa para el archivo o carpeta de destino (en "Nombre modificado")
        ruta_destino = os.path.join(carpeta_destino, nombre_nuevo_completo)

        # Verificar si el archivo o carpeta existe en la carpeta de origen
        if not os.path.exists(ruta_origen):
            estados.append(f"Archivo o carpeta no encontrado: ({nombre_original_completo})")
            continue  # No procesar si no existe en origen

        # Verificar si el nombre nuevo ya ha sido procesado antes (es un duplicado en el DataFrame)
        if nombre_nuevo_completo in nombres_procesados:
            estados.append(f"El archivo '{nombre_nuevo_completo}' ya fue procesado previamente en la tabla")
            continue
        
        # Añadir el nombre nuevo al conjunto de nombres procesados
        nombres_procesados.add(nombre_nuevo_completo)

        # Comprobar si el archivo o carpeta ya existe en la carpeta de destino
        if os.path.exists(ruta_destino):
            # Si el archivo o carpeta ya existe, registrar el estado y continuar
            estados.append(f"El archivo '{nombre_nuevo_completo}' ya existe en '{carpeta_destino}'.")
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
            estados.append(f"Error: Archivo o carpeta no encontrado ({nombre_original_completo})")
        except PermissionError:
            # Manejar errores de permisos
            estados.append(f"Error: Permiso denegado al acceder a ({nombre_original_completo}). Cerrar si el archivo está abierto.")
        except Exception as e:
            # Manejo de errores en caso de otros problemas al copiar
            estados.append(f"Error: {e}")

    # Añadir la lista de estados al DataFrame como una nueva columna
    DataFrame_a_procesar['Estado'] = estados

    # Devolver el DataFrame con la nueva columna
    return DataFrame_a_procesar




# Función de confirmación usando tkinter
def preguntar_proceder_funcion():
    """
    Esta función pregunta al usuario si desea continuar con el renombrado de archivos.
    """
    ventana = tk.Tk()
    ventana.withdraw()  # Ocultar la ventana principal
    respuesta = messagebox.askyesno('Reemplazar archivos', '¿Deseas reemplazar los archivos originales? \nEste proceso no se puede deshacer.')
    ventana.destroy()
    return respuesta



# Función principal para renombrar archivos
def renombrar_archivos_local(DataFrame_a_procesar, directorio, columna_original, columna_nuevo):
    if columna_original not in DataFrame_a_procesar.columns or columna_nuevo not in DataFrame_a_procesar.columns:
        print("Error: Las columnas especificadas no existen en el DataFrame.")
        sys.exit()

    # Preguntar al usuario si desea proceder
    if not preguntar_proceder_funcion():
        # Si el usuario presiona "No", muestra un mensaje y retorna sin cambios
        mostrar_error("El proceso no se realizo.")
        sys.exit() #Cancelar el script

    DataFrame_a_procesar = DataFrame_a_procesar[DataFrame_a_procesar[columna_original] != DataFrame_a_procesar[columna_nuevo]].copy()

    estado = []
    nombres_procesados = set()  # Crear un set para nombres nuevos ya procesados

    for index, row in DataFrame_a_procesar.iterrows():
        nombre_original = row[columna_original]
        nombre_nuevo = row[columna_nuevo]
        ruta_original = os.path.join(directorio, nombre_original)
        ruta_nueva = os.path.join(directorio, nombre_nuevo)

        if not os.path.exists(ruta_original):
            estado.append('Archivo no encontrado')
        elif nombre_nuevo in nombres_procesados:
            estado.append(f'{nombre_nuevo} ya fue procesado en la tabla')
            continue  # No procesar este archivo
        elif os.path.exists(ruta_nueva):
            estado.append(f'{nombre_nuevo} ya existe')
        else:
            try:
                os.rename(ruta_original, ruta_nueva)
                estado.append('Ok')
            except Exception as e:
                estado.append(f'Error: {str(e)}')

        # Añadir el nombre nuevo al set
        nombres_procesados.add(nombre_nuevo)

    DataFrame_a_procesar['Estado'] = estado

    return DataFrame_a_procesar


## FUNCIONES PARA MENU PRINCIPAL
# Función para abrir el repositorio
def abrir_github(event):
    webbrowser.open("https://github.com/JoseChirif/Renombrar-archivos-masivamente")





## FUNCIONES PARA CORRER SCRIPTS
# Obtener la ruta absoluta del directorio actual



def ejecutar_script_0_Importar_archivos_a_excel():
    """
    Ejecuta el script 0_Importar_archivos_a_excel.py ubicado en la carpeta src.
    """
    # Voy al directorio raíz del proyecto
    sys.path.append(os.path.abspath(os.path.dirname(__file__) + '/../'))
    # Pruebo ejecutar el script
    try:
        os.system("Python src/0_Importar_archivos_a_excel.py")
    except subprocess.CalledProcessError as e:
        print(f"Error al ejecutar script 0_Importar_archivos_a_excel: {e}")

def ejecutar_script_1_Modificar_nombres():
    """
    Ejecuta el script 1_Modificar_nombres.py ubicado en la carpeta src.
    """
    # Voy al directorio raíz del proyecto
    sys.path.append(os.path.abspath(os.path.dirname(__file__) + '/../'))
    # Pruebo ejecutar el script
    try:
        os.system("Python src/1_Modificar_nombres.py")
    except subprocess.CalledProcessError as e:
        print(f"Error al ejecutar script 1_Modificar_nombres: {e}")

def ejecutar_script_2_Desbloquear_excel():
    """
    Ejecuta el script 2_Desbloquear_excel.py ubicado en la carpeta src.
    """
    # Voy al directorio raíz del proyecto
    sys.path.append(os.path.abspath(os.path.dirname(__file__) + '/../'))
    # Pruebo ejecutar el script
    try:
        os.system("Python src/2_Desbloquear_excel.py")
    except subprocess.CalledProcessError as e:
        print(f"Error al ejecutar script 2_Desbloquear_excel: {e}")

        
#test    
# print(directorio_a_trabajar())
# print(preguntar_reemplazo("Archivo"))
# print(mostrar_error("mensaje de x"))

