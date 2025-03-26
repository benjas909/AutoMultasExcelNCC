# AutoMultasExcel

Desarrollado en Python 3.13, Windows 10 Enterprise 22H2

## Prerrequisitos

- Python 3.xx

- Openpyxl

Las instrucciones para su instalación se muestran a continuación.

## Antes de ejecutar por primera vez

1. Instalar la versión más reciente de Python desde la Microsoft Store, en caso de ya estar instalado, puedes saltar este paso.
    - Puedes revisar si python está instalado saltando al paso 2 y ejecutando "python --version" en PowerShell, si obtienes una versión, Python está instalado.

2. Abrir una ventana de Powershell en la carpeta donde se encuentre este programa:

    - En la carpeta, en un espacio vacío, hacer click derecho y elegir, "Abrir en Terminal"
    - En caso de que no aparezca esta opción, hacer Shift + Click derecho y seleccionar "Abrir la ventana de PowerShell aquí".

3. Ejecutar el comando "pip --version", si se muestra alguna versión, puedes saltarte al paso 5.

4. Ejecutar el comando "python ./get-pip.py" en powershell, esperar a que finalice
    - En caso de que algún comando que comience con "python" no funcione, probar reemplazando con "python3", en caso de que no funcione, probar a reinstalar python.

5. Al finalizar, ejecutar "python -m pip install openpyxl".

6. En caso de no obtener errores, se puede ejecutar el programa

## Ejecución del programa

1. En una pestaña de PowerShell ubicada en la carpeta que contiene el programa, ejecutar el comando "python ./test.py"

2. Se abrirá una ventana con una casilla que se debe seleccionar en caso de ser necesaria la importación de la hoja de tickets. Al seleccionar esta casilla, se activará el botón para elegir el archivo a importar.

3. Además, se debe seleccionar el archivo de multas y la ubicación y nombre del archivo de salida del programa.

4. Cuando aparezca el mensaje de "Archivo Guardado", el archivo estará listo para abrir en la ubicación elegida.
