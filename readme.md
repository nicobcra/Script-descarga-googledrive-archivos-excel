# Descarga masiva de Google Drive sincronizada con Excel

Script en Python que descarga de forma masiva los archivos enlazados en una hoja de Excel (links de Google Drive) y deja un reporte donde **cada fila coincide con su archivo descargado**. Está pensado para entornos corporativos restrictivos, donde los tokens y las APIs de Drive no están disponibles.

En lugar de usar la API, toma el control de una sesión de Chrome ya abierta (*Remote Debugging*) y hereda el inicio de sesión y las cookies de la empresa.

## Qué resuelve

- **Alineación 1:1 entre Excel y disco.** Procesa una fila a la vez y no avanza hasta que el archivo termina de descargarse (detecta la extensión temporal `.crdownload` de Chrome).
- **No modifica el original.** Lee el Excel en solo lectura y genera un reporte nuevo.
- **Errores visibles.** Con `openpyxl` pinta de rojo las filas con enlace vacío, fallo o tiempo de espera agotado.
- **Funciona detrás del firewall corporativo**, porque usa el navegador real ya autenticado.

## Tecnologías

`Python 3.10+` · `pandas` · `Selenium` · `openpyxl`

## Instalación

```bash
pip install pandas selenium openpyxl
```

## Uso

**1. Abre Chrome en modo remoto.** Cierra todas las ventanas de Chrome y, desde el Símbolo del sistema de Windows, ejecuta:

```bat
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\SeleniumChromeProfile"
```

Deja esa ventana abierta e inicia sesión en tu Google Drive.

**2. Configura las rutas** al inicio de `archivos.py`:

| Variable | Qué es |
| :-- | :-- |
| `ruta_excel_original` | Excel con el listado de links de Drive |
| `ruta_excel_nuevo_reporte` | Dónde se guardará el reporte generado |
| `carpeta_destino_descargas` | Carpeta local donde se acumulan las descargas |
| `nombre_hoja` | Nombre exacto de la pestaña a procesar |

**3. Ejecuta:**

```bash
python archivos.py
```

## Formato del reporte

El script genera un Excel independiente con cinco columnas:

| Columna | Origen |
| :-- | :-- |
| Cliente | Columna B del original |
| Asignado | Columna J del original |
| Link Drive | Columna N del original |
| ID Validación / Ticket | Columna O del original |
| Nombre Archivo / Estado | Columna nueva de control |

Las filas con enlace vacío o descarga fallida se resaltan en rojo.
