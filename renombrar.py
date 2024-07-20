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
            print(f"Renombrado: {file} -> {new_name}")

# Ejemplo de uso:
if __name__ == "__main__":
    folder_path = 'imagenes'  # Reemplazar con la ruta de tu directorio
    rename_files(folder_path)
