# Renombrar archivos en masa
¡Renombrar múltiples archivos es tan facil como editar un excel!
![gif_resume](https://github.com/user-attachments/assets/1f7a9477-ddc0-4e3f-a747-4d48567e4b85)



# Iniciar el programa:
- [Ejecutar desde python (click aquí)](#ejecutar-desde-python)

- [Ejecutar desde un archivo .exe (click aquí)](#ejecutar-desde-un-archivo-exe)

## Ejecutar desde python:
1. Instala un ambiente virtual con el comando `python -m venv venv`

2. Activa el ambiente virtual con `venv scripts activate`

3. Instala las dependencias desde requiremente.txt: `pip install -r requirements.txt` 
(asegurate de que el ambiente virtual este activo).

4. Corre run.py para que salga el menú principal: `python run.py` 
    ![father's directory](https://github.com/user-attachments/assets/efa73448-b450-4e51-a26c-c2f8ebb882a7)

    **Los scripts modifican los nombres de los archivos en el directorio padre del proyecto.**


[Ahora reviza el flujo de trabajo (click aquí)](#flujo-de-trabajo)





## Ejecutar desde un archivo exe :
1. Descarga el archivo .exe de lanzamientos:
    ![releases](https://github.com/user-attachments/assets/479ec700-56e1-4bf9-9d80-77743b8a1fbd)
    ![.exe](https://github.com/user-attachments/assets/57da40c2-41f7-4e71-b459-d785acd136d1)

2. Mueve el archivo.exe a la carpeta donde estan los archivos que deseas modificar.
    ![.exe in folder](https://github.com/user-attachments/assets/6e2859e6-dd8c-454d-ba5f-fef4369deb43)

    **Los scripts modificarán los archivos que estan ubicados en la misma carpeta que el archivo.exe.**

3. Ejecuta el archivo para abrir el menú principal.



[Ahora reviza el flujo de trabajo (click aquí)](#flujo-de-trabajo)




## Flujo de trabajo:
1. Con el menú principal abierto puedes correr todos los scripts.
    ![main menu](https://github.com/user-attachments/assets/9e9d432d-d3e2-4f8a-a4c4-dea361f054b3)

2. El primer botón creará automáticamente un excel que incluya el nombre de todos los archivos que estan en la misma carpeta de tu programa.
    ![excel template](https://github.com/user-attachments/assets/c43eb533-498d-46a5-87d3-1ab98e0f8348)

3. Llena la columna C del excel con los nombres nuevos que deseas asignar (deja la celda en blanco si deseas conservar el nombre actual).
    ![Flash fill excel](https://github.com/user-attachments/assets/ec5e8c1a-dc87-49f7-bff6-abe98b32a57c)
    **Recomendaciones:** 
    - Si modificas siguiendo un patrón, haz 2 ó 3 ejemplos y usa "Relleno rápido" de excel.
    - Si quieres desbloquear la hoja para trabajar más columnas o usar más formulas, el tercer botón del menú principal lo hace.
    

4. El segundo botón del menú modificará los nombres por los asignados en el excel.
    ![renaming](https://github.com/user-attachments/assets/e8aa9663-363b-4297-aa6f-55cae6d83c77)


    Tienes 2 opciones:
    ![options to modify names](https://github.com/user-attachments/assets/8d4136fe-5dc2-43c5-875a-fc729e16124d)
    - Opción 1: Crear un nuevo folder (con el nombre del .exe o directorio si estas en python) y copiar todos los archivos con el nuevo nombre asignado ahí (si no asignaste nuevo nombre se copiará el archivo conservado el nombre inicial).
    - Opción 2: Modificar los nombres de los archivos en la misma carpeta.


**¡LISTO, finalizaste!**













