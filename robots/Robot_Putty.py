import pyautogui
import time
import pandas as pd
from utils import forzar_caps_off

def ejecutar_stock(df, total, log_func, progress_func, velocidad):
    """
    Robot de Carga de Stock (3-6-1).
    Optimizado para ráfaga de carga y blindado contra errores de índice.
    """
    pyautogui.PAUSE = velocidad
    forzar_caps_off()
    
    if df.empty:
        log_func("❌ ERROR: El archivo Excel no tiene datos.")
        return False

    # 1. NAVEGACIÓN INICIAL
    log_func("Entrando al módulo de stock (3-6-1)...")
    for k in ['3', '6', '1']:
        pyautogui.write(k); pyautogui.press('enter')
        time.sleep(0.5)
    
    # 2. CARGA DE CABECERA (Columna C / Índice 2)
    try:
        # Extraemos Pedido (C1), Obs (C2), Impresora (C3)
        pedido = str(df.iloc[0, 2]).split('.')[0].strip() if not pd.isna(df.iloc[0, 2]) else ""
        obs    = str(df.iloc[1, 2]).strip() if not pd.isna(df.iloc[1, 2]) else ""
        imp    = str(df.iloc[2, 2]).upper().strip() if not pd.isna(df.iloc[2, 2]) else "LP1"
        
        log_func(f"📌 Cabecera: Pedido {pedido} | Obs: {obs}")

        pyautogui.write(pedido); pyautogui.press('enter'); time.sleep(0.8)
        pyautogui.press('enter') 
        pyautogui.write(obs); pyautogui.press('enter'); time.sleep(0.5)
        pyautogui.press('enter')
        pyautogui.write(imp); pyautogui.press('enter'); time.sleep(1.2)
        
        # Entrar a la grilla usando el número de pedido
        pyautogui.write(pedido); pyautogui.press('enter'); time.sleep(1.5)
        
    except Exception as e:
        log_func(f"❌ Error en cabecera: {e}")
        return False

    # 3. BUCLE DE ARTÍCULOS (RÁFAGA)
    log_func("⏳ Iniciando ráfaga de carga...")
    
    for i, row in df.iterrows():
        # Tomamos SKU (Col A / 0) y Cantidad (Col B / 1)
        val_sku = row.iloc[0]
        if pd.isna(val_sku): continue
        
        try:
            sku = str(val_sku).split('.')[0].strip()
            # Quitamos el .0 de las cantidades por si vienen de Excel
            cantidad = str(int(float(str(row.iloc[1]).strip()))) if not pd.isna(row.iloc[1]) else "0"
            
            # Lógica Modo G (Columna D / Índice 3)
            info_d = str(row.iloc[3]).strip() if len(row) > 3 and not pd.isna(row.iloc[3]) and str(row.iloc[3]).lower() != 'nan' else None

            # Escribir SKU y navegar a cantidad (4 enters)
            pyautogui.write(sku)
            pyautogui.press('enter', presses=4, interval=0.1)

            if info_d:
                # MODO G (Pesable / Galpón)
                pyautogui.write('g')
                pyautogui.write(cantidad); pyautogui.press('enter')
                time.sleep(0.4)
                pyautogui.write(info_d); pyautogui.press('enter')
            else:
                # MODO U (Unidad normal)
                pyautogui.write(f"u{cantidad}"); pyautogui.press('enter')
            
            time.sleep(0.1) 
            log_func(f"✅ Cargado: {sku} ({cantidad})")

        except Exception as row_err:
            log_func(f"⚠️ Error en fila {i+1}: {row_err}")
            
        progress_func((i + 1) / total)

    # 4. CIERRE Y GUARDADO (TU SECUENCIA)
    log_func("💾 Guardando cambios en PuTTY...")
    pyautogui.press('f5'); time.sleep(3.0)
    
    # Secuencia para volver al menú principal
    for k in ['end', 'enter', 'end', 'end']:
        pyautogui.press(k); time.sleep(0.5)

    log_func("🏁 Proceso de Stock finalizado con éxito.")
    return True