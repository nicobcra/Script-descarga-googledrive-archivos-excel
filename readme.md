# Google Drive Mass Synchronized Downloader 🚀

Este script de automatización en Python está diseñado para resolver de manera definitiva los problemas de descarga masiva desde Google Drive en entornos corporativos restrictivos. 

A diferencia de los scripts tradicionales que fallan por tokens expirados o cambios en el código de la interfaz de Google, esta herramienta toma el control de una sesión activa de Chrome (`Remote Debugging`) y utiliza un **mecanismo de bloqueo por hardware de precisión** para asegurar una alineación matemática perfecta (1:1) entre los registros de un archivo Excel y los archivos físicos descargados en disco.

## ✨ Características Principales
* **Sincronización Perfecta:** El script congela su ejecución por completo por cada fila procesada y monitorea en tiempo real el disco duro de la computadora. No avanza a la siguiente fila hasta que el archivo actual termine de descargarse (rastreo de extensiones `.crdownload`).
* **Protección del Archivo Base:** El archivo Excel original se lee en modo de solo lectura. El script genera un informe totalmente nuevo desde cero.
* **Manejo Visual de Errores:** Utilizando `openpyxl`, el script analiza el resultado final y pinta automáticamente de **color rojo** las filas cuyos enlaces presenten fallos o tiempos de espera agotados.
* **Bypass de Firewalls Corporativos:** Al conectarse a una ventana remota del navegador real de la compañía, hereda de forma automática el inicio de sesión, credenciales y cookies activas de la empresa.

## 🛠️ Requisitos e Instalación

1. **Python 3.10+** instalado en el equipo.
2. Instalar las dependencias necesarias mediante la consola de comandos:
```bash
pip install pandas selenium openpyxl webdriver-manager


## 🚀 Guía de Uso Step-by-Step
Paso 1: Inicializar Chrome en Modo Remoto
Para que Python pueda tomar el control de tu sesión activa sin requerir logueos extras y evadiendo políticas rígidas de seguridad, cierra todas tus ventanas corrientes de Google Chrome, abre el Símbolo del sistema (cmd) de Windows y ejecuta:

"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\SeleniumChromeProfile"

Nota: Deja abierta esta ventana del navegador e inicia sesión en tu Google Drive corporativo.

Paso 2: Configurar las Rutas en el Script

Abre el script downloader.py y ajusta las siguientes variables a tus rutas deseadas:
ruta_excel_original: Ruta del listado con links de Drive.
ruta_excel_nuevo_reporte: Destino del nuevo reporte limpio a generar.
carpeta_destino_descargas: Carpeta local donde se acumularán las descargas físicas.
nombre_hoja: Nombre exacto de la pestaña del Excel a procesar.


Paso 3: Ejecutar la Automatización
Corre el script desde tu entorno de desarrollo o consola:

Bash
python downloader.py


📊 Formato de Salida
El script genera un archivo Excel independiente con una estructura unificada de 5 columnas:

Cliente (Extraído de Columna B original)
Asignado (Extraído de Columna J original)
Link Drive (Extraído de Columna N original)
ID Validación / Ticket (Extraído de Columna O original)
Nombre Archivo / Status (Nueva Columna P de control)

Cualquier fila que presente un enlace vacío o una descarga fallida se resaltará automáticamente en color rojo.