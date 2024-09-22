#ESTE SCRIPT ES SOLO PARA MOSTRAR LOS NOMBRES DE TODAS LAS FUNCIONES DE functions/functions.py Y TENERLOS A LA MANO SI SE REQUIERE.

import inspect
from functions import *

## LISTA DE FUNCIONES
# Función para listar todas las funciones en el módulo actual
def listar_funciones():
    # Obtener todas las funciones definidas en el módulo actual
    funciones = [obj for name, obj in inspect.getmembers(__import__(__name__)) if inspect.isfunction(obj)]
    for f in funciones:
        print(f.__name__)

# Llamar a la función para imprimir las funciones
listar_funciones()
