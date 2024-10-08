#librerias
import sys
import os

# Añadir el directorio del proyecto a sys.path para llamar functions.functions en mis scripts
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from functions.functions import relative_route_to_file, obtener_nombre_archivo_o_directorio,  cargar_lista_de_lenguajes, cargar_traducciones, eliminar_desde_ultimo_punto


# Rutas y variables comunes
# icon_picture_ico
icon_picture_ico = relative_route_to_file("assets", "icon.ico")

# icon_picture_png
icon_picture_png = relative_route_to_file("assets", "icon.png")
  
# logo_github
logo_github_png = relative_route_to_file("assets", "icon_github.png")


# Nombres comunes scripts
project_name  = obtener_nombre_archivo_o_directorio()
excel_name = eliminar_desde_ultimo_punto(project_name) + '.xlsx' # eliminar_desde_ultimo_punto borrara la extensión que tenga (ejemplo .exe)


# Archivos a ignorar en df
files_to_avoid = [
    excel_name,
    project_name # Obtiene el name del file actual
]    
# Añado folder renombrado en cada idioma
# Obtener el diccionario de idiomas del archivo languages.json
idiomas = cargar_lista_de_lenguajes()
# Iterar sobre cada idioma en el diccionario de idiomas
for idioma in idiomas:
    # Cargar las traducciones del idioma actual
    traducciones = cargar_traducciones(idioma)
    # Agregar folder_with_renamed_files a la lista
    if 'renamed' in traducciones:
        renamed = traducciones['renamed']
        folder_with_renamed_files = project_name + " - " + renamed
        files_to_avoid.append(folder_with_renamed_files)
        

# print(files_to_avoid)

