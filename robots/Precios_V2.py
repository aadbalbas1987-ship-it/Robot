import pyautogui
import time
import pandas as pd
import os
import sys
from utils import limpiar_sku, forzar_caps_off, f_monto

def cargar_listado_hijos():
    """Carga el archivo hijos.txt detectando si es ejecución directa o desde el .exe"""
    hijos = set()
    if getattr(sys, 'frozen', False):
        ruta_base = os.path.dirname(sys.executable)
    else:
        ruta_base = os.path.dirname(os.path.abspath(__file__))
        if "robots" in ruta_base: ruta_base = os.path.dirname(ruta_base)

    ruta_txt = os.path.join(ruta_base, 'hijos.txt')
    try:
        if os.path.exists(ruta_txt):
            with open(ruta_txt, 'r', encoding='utf-8') as f:
                hijos = {line.strip() for line in f if line.strip()}
    except: pass
    return hijos

def ejecutar_precios_v2(df, total, log_func, progress_func, velocidad):
    """
    Robot de Precios con lógica dinámica de Enters y discriminación de Hijos.
    Mapeo: A(0): SKU, B(1): Costo, C(2): P.Salón, D(3): P.Mayo, E(4): Info Extra.
    """
    pyautogui.PAUSE = velocidad
    forzar_caps_off()
    
    listado_hijos = cargar_listado_hijos()
    log_func(f"💰 Módulo Precios: {len(listado_hijos)} hijos cargados.")
    
    # 1. ENTRADA AL MÓDULO (3 -> 4 -> 2)
    for k in ['3', '4', '2']:
        pyautogui.write(k); time.sleep(0.5)

    # 2. BUCLE DE ARTÍCULOS
    for i, row in df.iterrows():
        sku_raw = str(row.iloc[0]).split('.')[0].strip()
        sku = limpiar_sku(sku_raw)
        
        if not sku or sku.lower() in ['codigo', 'sku', 'articulo']: continue
            
        try:
            costo = f_monto(row.iloc[1])
            p_salon = f_monto(row.iloc[2])
            p_mayo = f_monto(row.iloc[3])
            
            # Columna E (Índice 4): Info extra o condicional
            info_e = str(row.iloc[4]).strip() if len(row) > 4 and not pd.isna(row.iloc[4]) and str(row.iloc[4]).lower() != 'nan' else None

            # Determinar cantidad de Enters: Hijos (6), Normales (5)
            es_hijo = sku in listado_hijos
            cant_enters = 6 if es_hijo else 5

            log_func(f"🔄 Fila {i+1}: SKU {sku} ({'HIJO' if es_hijo else 'NORMAL'})")

            # --- SECUENCIA PUTTY ---
            pyautogui.write(sku)
            pyautogui.press('enter', presses=cant_enters, interval=0.05)
            
            pyautogui.write(costo)
            pyautogui.press('enter', presses=3, interval=0.05)
            
            pyautogui.write(p_salon)
            
            # Lógica condicional según Columna E
            if info_e:
                pyautogui.press('enter', presses=2, interval=0.05)
                pyautogui.write(info_e); pyautogui.press('enter')
                pyautogui.write(p_mayo)
            else:
                pyautogui.press('enter', presses=3, interval=0.05)
                pyautogui.write(p_mayo)
            
            pyautogui.press('enter') 
            pyautogui.press('f5') # Confirmar cambio de precio   
            
            time.sleep(1.0) # Espera a que el sistema procese el cambio

        except Exception as e:
            log_func(f"⚠️ Error en SKU {sku}: {e}")

        progress_func((i + 1) / total)

    # 3. SALIDA AL MENÚ (Secuencia de reseteo)
    log_func("🧹 Finalizando y regresando al menú principal...")
    for _ in range(3):
        pyautogui.press('end'); time.sleep(0.5)

    log_func("✅ Cambio de precios finalizado con éxito.")
    return True