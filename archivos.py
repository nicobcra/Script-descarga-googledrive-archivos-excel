import os
import re
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# --- CONFIGURACIÓN DE RUTAS ---
ruta_excel_original = r"C:\Users\nicolas.becerra\Downloads\Atenea_Fz\cHARGUE LEVANTAMIENTOS.xlsx"
ruta_excel_nuevo_reporte = r"C:\Users\nicolas.becerra\Downloads\Atenea_Fz\REPORTE_DESCARGAS_FINAL.xlsx"
carpeta_destino_descargas = r"C:\Users\nicolas.becerra\Downloads\Atenea_Fz\Archivos Cargue"
nombre_hoja = "CargueFinal"

if not os.path.exists(carpeta_destino_descargas):
    os.makedirs(carpeta_destino_descargas)

# --- CONEXIÓN AL CHROME ACTIVO (PUERTO 9222) ---
options = Options()
options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

print("Conectándose a tu Chrome con sesión iniciada...")
driver = webdriver.Chrome(options=options)

# --- LEER EXCEL ORIGINAL ---
print("Leyendo archivo original desde la primera fila (Sin encabezados)...")
# header=None garantiza que la Fila 1 sea procesada como datos reales y no títulos
df_original = pd.read_excel(ruta_excel_original, sheet_name=nombre_hoja, header=None)

filas_nuevo_reporte = []

def obtener_id_drive(url):
    if pd.isna(url):
        return None
    match = re.search(r'id=([a-zA-Z0-9_-]+)', str(url))
    if match:
        return match.group(1)
    match_share = re.search(r'drive\.google\.com/file/d/([a-zA-Z0-9_-]+)', str(url))
    if match_share:
        return match_share.group(1)
    if len(str(url).strip()) > 15 and " " not in str(url):
        return str(url).strip()
    return None

print("\nIniciando descargas con bloqueo total de hardware y sincronización perfecta...")

for index, row in df_original.iterrows():
    # Extraer los datos físicos reales (B=1, J=9, N=13, O=14)
    val_B = row.iloc[1] if len(row) > 1 else ""   # Cliente
    val_J = row.iloc[9] if len(row) > 9 else ""   # Asignado
    val_N = row.iloc[13] if len(row) > 13 else "" # Link Drive
    val_O = row.iloc[14] if len(row) > 14 else "" # ID Ticket
    
    id_drive = obtener_id_drive(val_N)
    
    # Si la fila no contiene un enlace de Drive válido, guardamos el espacio vacío de forma estricta
    if not id_drive:
        filas_nuevo_reporte.append({
            "Cliente (Columna B)": val_B,
            "Asignado (Columna J)": val_J,
            "Link Drive (Columna N)": val_N,
            "ID Validación (Columna O)": val_O,
            "Nombre Archivo (Columna P Nueva)": "Sin enlace de Drive"
        })
        continue
        
    num_fila_real = index + 1
    print(f"\n[Fila {num_fila_real}] Procesando Ticket: {val_O} | Cliente: {val_B}")
    nombre_archivo_p = "Error / No Descargado"
    
    # Captura exacta de la carpeta un milisegundo antes de disparar la descarga de ESTA fila
    archivos_antes = set(os.listdir(carpeta_destino_descargas))
    
    try:
        url_descarga_directa = f"https://drive.google.com/uc?export=download&id={id_drive}"
        driver.get(url_descarga_directa)
        
        # --- BUCLE DE BLOQUEO ABSOLUTO (INMUNE AL TIEMPO) ---
        descarga_exitosa = False
        segundos_esperados = 0
        
        while segundos_esperados < 300:  # 5 minutos máximo por si el archivo está muy pesado
            time.sleep(1)
            segundos_esperados += 1
            
            archivos_actuales = set(os.listdir(carpeta_destino_descargas))
            nuevos_archivos = archivos_actuales - archivos_antes
            
            if nuevos_archivos:
                archivo_detectado = list(nuevos_archivos)[0]
                
                # Si el archivo termina en .crdownload significa que Chrome sigue escribiéndolo en disco.
                # El código se mantiene CONGELADO aquí y no avanza a la siguiente fila.
                if archivo_detectado.endswith('.crdownload'):
                    continue
                
                # Descarga completada al 100% con su nombre final
                nombre_archivo_p = archivo_detectado
                descarga_exitosa = True
                break
                
        if descarga_exitosa:
            print(f"-> [OK PERFECTO] Fila {num_fila_real} perfectamente alineada con: {nombre_archivo_p}")
        else:
            nombre_archivo_p = "Error: Tiempo de espera agotado"
            print(f"-> Advertencia: El archivo de la fila {num_fila_real} no bajó.")
            
    except Exception as e:
        print(f"-> Error crítico en fila {num_fila_real}: {e}")
        nombre_archivo_p = "Error en enlace"

    # Insertamos los datos en la lista respetando estrictamente el orden secuencial
    filas_nuevo_reporte.append({
        "Cliente (Columna B)": val_B,
        "Asignado (Columna J)": val_J,
        "Link Drive (Columna N)": val_N,
        "ID Validación (Columna O)": val_O,
        "Nombre Archivo (Columna P Nueva)": nombre_archivo_p
    })

# --- 1. CREAR EL NUEVO EXCEL DESDE CERO ---
print("\nEscribiendo el reporte base...")
df_nuevo = pd.DataFrame(filas_nuevo_reporte)
with pd.ExcelWriter(ruta_excel_nuevo_reporte, engine='openpyxl') as writer:
    df_nuevo.to_excel(writer, sheet_name="Reporte Descargas", index=False)

# --- 2. APLICAR FORMATO COLOR ROJO A FILAS CON ERROR (OPENPYXL) ---
print("Aplicando estilos visuales y marcas rojas a errores...")
wb = load_workbook(ruta_excel_nuevo_reporte)
ws = wb["Reporte Descargas"]

# Definimos el color rojo suave/claro para que sea fácil de leer el texto
relleno_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

# Recorremos el archivo Excel generado fila por fila buscando errores en la columna E (Nombre Archivo)
# Empezamos en la fila 2 porque la fila 1 contiene los encabezados del reporte nuevo
for row_idx in range(2, ws.max_row + 1):
    celda_resultado = ws.cell(row=row_idx, column=5) # Columna 5 es la 'P' (Nombre Archivo) en este reporte nuevo
    valor = str(celda_resultado.value)
    
    # Si contiene la palabra Error o Sin enlace, pintamos toda la fila de color rojo
    if "Error" in valor or "Sin enlace" in valor or valor == "No Descargado":
        for col_idx in range(1, 6): # Columnas de la 1 a la 5 (De la A a la E de este reporte)
            ws.cell(row=row_idx, column=col_idx).fill = relleno_rojo

# Guardamos el archivo con los colores aplicados
wb.save(ruta_excel_nuevo_reporte)

print(f"\n¡PROCESO FINALIZADO CON ÉXITO!")
print(f"Tu archivo original quedó intacto. Las filas con errores están pintadas de ROJO.")
print(f"Reporte perfecto con formato en: {ruta_excel_nuevo_reporte}")