import pyautogui
import time
import pandas as pd
import os
import sys
from utils import limpiar_sku, forzar_caps_off, f_monto

def cargar_listado_hijos():
    """Busca el archivo hijos.txt en la misma carpeta que el ejecutable o el script."""
    hijos = set()
    
    # Esta línea detecta automáticamente dónde está parado el .exe
    if getattr(sys, 'frozen', False):
        # Si es el .exe
        ruta_base = os.path.dirname(sys.executable)
    else:
        # Si es el código .py
        ruta_base = os.path.dirname(os.path.abspath(__file__))
        # Si el script está dentro de la carpeta 'robots', subimos un nivel
        if "robots" in ruta_base:
            ruta_base = os.path.dirname(ruta_base)

    ruta_txt = os.path.join(ruta_base, 'hijos.txt')
    
    try:
        if os.path.exists(ruta_txt):
            with open(ruta_txt, 'r', encoding='utf-8') as f:
                hijos = {line.strip() for line in f if line.strip()}
    except Exception:
        pass
    return hijos

# --- CAMBIO DE NOMBRE AQUÍ PARA QUE main_gui.py NO DE ERROR ---
def ejecutar_precios_v2(df, total, log_func, progress_func, velocidad):
    """
    Tu código ORIGINAL con el nombre corregido para el import.
    """
    pyautogui.PAUSE = velocidad
    forzar_caps_off()
    
    # 1. Cargamos el listado de hijos al inicio
    listado_hijos = cargar_listado_hijos()
    log_func(f"💰 Iniciando Módulo de Precios (Base de {len(listado_hijos)} hijos)")
    
    # --- ENTRADA AL MÓDULO (3 -> 4 -> 2) ---
    for k in ['3', '4', '2']:
        pyautogui.write(k)
        time.sleep(0.5)

    # --- BUCLE DE ARTÍCULOS ---
    for i, row in df.iterrows():
        # 1. Limpieza de datos (quitando el .0 por si las dudas)
        sku_raw = str(row[0]).split('.')[0].strip()
        sku = limpiar_sku(sku_raw)
        
        if not sku or sku.lower() in ['codigo', 'sku', 'articulo', 'código']:
            continue
            
        try:
            costo = f_monto(row[1])
            p_salon = f_monto(row[2])
            p_mayo = f_monto(row[3])
            
            # Verificamos si hay información en Columna E (índice 4)
            info_e = str(row[4]).strip() if len(row) > 4 and not pd.isna(row[4]) and str(row[4]).lower() != 'nan' else None

            # --- DETECCIÓN DE HIJO ---
            es_hijo = sku in listado_hijos
            cant_enters = 6 if es_hijo else 5

            log_func(f"🔄 SKU {sku} ({'HIJO' if es_hijo else 'NORMAL'})...")

            # --- SECUENCIA EN PUTTY (TU LÓGICA ORIGINAL) ---
            pyautogui.write(sku)
            
            # Aplicamos la cantidad de enters dinámica según el listado
            pyautogui.press('enter', presses=cant_enters, interval=0.04)
            
            pyautogui.write(costo)
            pyautogui.press('enter', presses=3, interval=0.03)
            
            pyautogui.write(p_salon)
            
            # --- CONDICIONAL POR COLUMNA E ---
            if info_e:
                pyautogui.press('enter', presses=2, interval=0.03)
                pyautogui.write(info_e)
                pyautogui.press('enter')
                pyautogui.write(p_mayo)
            else:
                pyautogui.press('enter', presses=3, interval=0.03)
                pyautogui.write(p_mayo)
            
            pyautogui.press('enter') 
            pyautogui.press('f5')    
            
            # Pausa de seguridad
            time.sleep(1.2) 

        except Exception as e:
            log_func(f"⚠️ Error en fila {i+1} (SKU {sku}): {e}")

        # Actualizar barra de progreso
        progress_func((i + 1) / total)

    # --- SALIDA AL MENÚ PRINCIPAL (TU SALIDA ORIGINAL) ---
    log_func("🧹 Limpiando y regresando al menú principal...")
    pyautogui.press('end'); time.sleep(0.5)
    pyautogui.press('end'); time.sleep(0.5)
    pyautogui.press('end'); time.sleep(0.8)

    log_func("✅ Cambio de precios finalizado con éxito.")