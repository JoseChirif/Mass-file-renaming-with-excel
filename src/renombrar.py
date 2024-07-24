# importing python-env & python-decouple
from decouple import config

# Importinr libraries to manage folders and files
import os
import re

    
def rename_files(folder_path):
    # Obtener la lista de archivos en el directorio
    files = os.listdir(folder_path)
    
    # Expresión regular para encontrar el patrón con tres guiones bajos
    pattern = re.compile(r'^(.*?)_(.*?)_(.*?)_(.*?)\.(.*)$')
    
    for file in files:
        # Solo procesar archivos con el patrón especificado
        match = re.match(pattern, file)
        if match:
            # Extraer las partes del nombre del archivo
            prefix = match.group(1)
            middle1 = match.group(2)
            middle2 = match.group(3)
            suffix = match.group(4)
            extension = match.group(5)
            
            # Construir el nuevo nombre del archivo
            new_name = f"{prefix}_{suffix}_{middle1}_{middle2}.{extension}"
            
            # Ruta completa original y nueva
            old_path = os.path.join(folder_path, file)
            new_path = os.path.join(folder_path, new_name)
            
            # Renombrar el archivo
            os.rename(old_path, new_path)
            print(f"Renaming: {file} -> {new_name}")


# Config file:
if __name__ == "__main__":
    default_folder = os.path.join(os.path.dirname(__file__), '0files_rename')

    while True:
        print("Seleccione una opción para elegir la carpeta:")
        print("1. Carpeta por defecto (0files_rename)")
        print("2. Nombre de carpeta dentro del proyecto")
        print("3. Ruta completa en tu PC")

        choice = input("Ingrese el número de opción (1, 2, o 3): ").strip()

        if choice == '1':
            folder_path = default_folder
            break
        elif choice == '2':
            folder_name = input(f'Seleccionar ruta: {os.getcwd()}\ ').strip()
            folder_path = folder_name
            break
        elif choice == '3':
            folder_path = input("Ingrese la ruta completa de la carpeta: ").strip()
            break
        else:
            print("Opción no válida. Por favor, ingrese 1, 2 o 3.")

    print(f"Usando la carpeta: {folder_path}")
    rename_files(folder_path)