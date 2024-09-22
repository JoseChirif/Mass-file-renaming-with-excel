#Importo librerias
import os
import sys
import pandas as pd
import numpy as np
from decouple import config  # Para importar variables de .env

# Importo funciones a usar
sys.path.append(os.path.abspath(os.path.dirname(__file__) + '/../'))  # Me dirije al directorio del proyecto
from functions.functions import directorio_a_trabajar, mostrar_error, mostrar_mensaje, mostrar_opciones, desbloquear_protección_excel, copiar_archivos_con_nuevo_nombre, renombrar_archivos_local,ejecutar_script_0_Importar_archivos_a_excel 

# Importo y declaro las variables de .env
archivo_excel = config('NOMBRE_ARCHIVO_EXCEL')
columna_nombre_original_excel_inicial = config('COLUMNA_1_EXCEL')
columna_extension_excel_inicial = config('COLUMNA_2_EXCEL')
columna_nombre_nuevo_excel_inicial = config('COLUMNA_3_EXCEL')
nombre_carpeta_destino = config("CARPETA_DESTINO")

#Declaro varaibles tipo lista
# Nombres de archivos a excluir
archivos_excluir = [
    "system32",
    archivo_excel, #Para que no se renombre el excel que estamos utilizando
    nombre_carpeta_destino, #Para evitar un bucle que se vaya copiando el contenido dentro cada que ejecuto
    os.path.basename(os.getcwd()), #Obtiene la carpeta del proyecto (si estas en Python)
    os.path.basename(__file__) # Obtiene el nombre del archivo actual (ejecutable)
]

#Creo variables para nombres de este script
# Columnas df
nombre_original_completo = 'Nombre original completo'
nombre_nuevo_completo = 'Nombre nuevo completo'
Estado = "Estado"
Errores = "Errores"

#opciones metodo
opcion1 = "Crear una nueva carpeta y copiar los archivos con sus nuevos nombres."
borde1 = "black"
opcion2 = "¡Sobreescribir los archivos con los nuevos nombres!"
borde2 = "red"




## EJECUCIÓN
# Aseguro estar en el directorio correcto
directorio = directorio_a_trabajar()

# Ruta completa del archivo Excel
archivo_excel_ruta = os.path.join(directorio, archivo_excel)


## TRABAJO Y LIMPIEZA DEL DATAFRAME
# Verifica si el archivo Excel existe
if not os.path.exists(archivo_excel_ruta):
    # Muestra un mensaje de error si el archivo no existe
    mensaje_error = f"El excel '{archivo_excel}' no se encuentra en la carpeta. \n \n Se creará el archivo en la carpeta."
    mostrar_error(mensaje_error)

    # Corre el script src/0 Importar archivos a excel.py
    ejecutar_script_0_Importar_archivos_a_excel()
    
    
    # Corto la ejecución
    sys.exit()

else:
    # Si el archivo existe, intantará leerlo
    try:
        df = pd.read_excel(archivo_excel_ruta, usecols=[0, 1, 2])
    except PermissionError:
        # Si el archivo está abierto o en uso, se muestra un mensaje de error y corto la ejecución
        mostrar_error(f"Cerrar '{archivo_excel}' para poder modificar los nombres de los archivos.")
        sys.exit()
        
     
    # Verificamos que los encabezados sean correctos
    encabezados_esperados = [columna_nombre_original_excel_inicial, columna_extension_excel_inicial, columna_nombre_nuevo_excel_inicial]
    # Si no son correctos:
    if list(df.columns) != encabezados_esperados:
        desbloquear_protección_excel(archivo_excel_ruta)
        mostrar_error(f"La estructura debe iniciar como se brindó la plantilla desde la celda 'A1' con los encabezados: {encabezados_esperados} respectivamente. \n Se deshabilito la protección de hoja en el excel.")
        # Corto la ejecución
        sys.exit()
    
       
    # Elimino las celdas que en la columna A están vacias
    df = df[pd.notna(df[columna_nombre_original_excel_inicial])]
       
    #Crear valores vacio en vez de la asignación 'nan' en el df
    df = df.fillna('')
    
   
    # Convertir los nombres a tipo str
    df.loc[:, [columna_nombre_original_excel_inicial, columna_nombre_nuevo_excel_inicial]] = df.loc[:, [columna_nombre_original_excel_inicial, columna_nombre_nuevo_excel_inicial]].astype(str)
 

    # Crear la columna nombre_original_completo concatenando columna_nombre_original_excel_inicial y columna_extension_excel_inicial
    df[nombre_original_completo] = df[columna_nombre_original_excel_inicial] + df[columna_extension_excel_inicial]
    
    #LIMPIEZA
    # ELimino los archivos que pueden generar conflicto
    df = df[~df.loc[:, nombre_original_completo].isin(archivos_excluir)]

    # Crear la columna nombre_nuevo_completo (original + extension si nuevo esta vacio. Sino, nuevo + extension)
    df[nombre_nuevo_completo] = df.apply(
    lambda row: (row[columna_nombre_original_excel_inicial] + row[columna_extension_excel_inicial]) 
    if row[columna_nombre_nuevo_excel_inicial] == '' 
    else (row[columna_nombre_nuevo_excel_inicial] + row[columna_extension_excel_inicial]), 
    axis=1
)
    
    # Verifico si hay nombres para modificar. Si no hay, muestró mensaje de error y cierro el programa.
    if df.loc[:, columna_nombre_nuevo_excel_inicial].isnull().all() or (df.loc[:, columna_nombre_nuevo_excel_inicial] == "").all():
        mostrar_error(f"La columna {columna_nombre_nuevo_excel_inicial} en el excel esta vacía. \nLlenarla para modificar los nombres.")
        sys.exit()

   



## MÉTODOS PARA MODIFICAR NOMBRES  
# Pregunta método
nro_opcion = mostrar_opciones("Escoge una opción", opcion1, opcion2, borde1, borde2)

# Verificando nro_opcion
# print(f'Se escogió la opción {nro_opcion}')
# if nro_opcion == None:
#     print("Cancelar")
# else:
#     print(f"Opción seleccionada: {globals()[f'opcion{nro_opcion}']}" if f'opcion{nro_opcion}' in globals() else "Error: Opción no válida.")
 
 
# Si da cancelar, corto la ejecución.   
if nro_opcion == None:
    sys.exit()
# Si da opción 1, usaré método de nueva carpeta. 
elif nro_opcion == 1:
    #Crea la carpeta y copia los archivos con nuevo nombre:
    ruta_carpeta_destino = os.path.join(directorio, nombre_carpeta_destino)
    # Correr metodo
    copiar_archivos_con_nuevo_nombre(df,directorio,ruta_carpeta_destino)
#Si se da la opción 2, usaré el mismo metodo copiandolo en la misma carpeta
elif nro_opcion == 2:
    # Correr metodo sobre la misma carpeta
    df = renombrar_archivos_local(df, directorio, nombre_original_completo, nombre_nuevo_completo)
    
 
# Otra opcion marcará error y saldrá del programa
else:
    mensaje_error("Error desconocido")
    sys.exit()
    




## CREACIÓN DEL LOG DE ERRORES
# Filtrar el DataFrame para que "Estado" sea diferente de 'Ok'
df = df[df["Estado"] != 'Ok']

#Si el len(df) = 0, mostrar mensaje de éxito y salir.
if len(df)==0:
    mostrar_mensaje("Éxito", f"Los archivos fueron guardados correctamente en la carpeta {nombre_carpeta_destino}.")
    sys.exit()
    
#Si len(df) != 0 Entonces crear log de errores
else: 
    # Cambiar el nombre de la columna "Estado" a "Errores"
    df = df.rename(columns={"Estado": "Errores"})


    # Guardar el DataFrame en un archivo Excel (no reemplazar el último):
    nombre_archivofinal_errores = "0 Nombres archivos _log errores.xlsx"
    # Nombre base del archivo
    nombre_archivo_errores = "0 Nombres archivos _log errores"
    extension = ".xlsx"
    contador = 1
    
    # Ruta completa inicial
    ruta_completa = os.path.join(directorio, f"{nombre_archivo_errores}{extension}")

    # Verificar si el archivo existe y agregar sufijos con números si es necesario
    while os.path.exists(ruta_completa):
        contador += 1
        ruta_completa = os.path.join(directorio, f"{nombre_archivo_errores} - {contador}{extension}")

    # Guardar el archivo en la ruta final
    df.to_excel(ruta_completa, index=False)

    mostrar_error(f'Algunos archivos no fueron procesados. \n La lista de errores se guardo en {nombre_archivo_errores} - {contador}{extension}')
    # Mostrar mensaje de éxito
    sys.exit()

