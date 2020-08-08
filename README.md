# OT-Python
Es un programa escrito en Python, que permite crear ordenes de trabajo, apartir de un documento Gantt en excel.


Para utilizarlo tendras que descargar las siguientes librerias:
1)
    pip install datetime
Documentación aquí:https://docs.python.org/3/library/datetime.html

2)
    pip install openpyxl
Documentación aquí: https://openpyxl.readthedocs.io/en/stable/  


Despues editas el archivo excel de nombre: gantt.xlsx  allí pondras toda la información que se requiere (es importante que el nombre del archivo no cambie, el programa busca este nombre).

Dentro de "gantt.xlsx" encontraras un breve ejemplo de como llenarlo.

Luego desde terminal podras ejecutar

python OT_test1.py

Así obtendras todas las ordenes de trabajo de cada una de las actividades planteadas en el gantt.

Si quieres modificar como se verán los archivos finales, sientete libre de modificar los colores del archivo: OT-0-Plantilla.xlsx (si modificas las casillas donde deberian ir los nombres, recuerda cambiarlo en el programa de python también)


Espero que te sirva de ayuda o de inspiración.
