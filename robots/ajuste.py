import pyautogui
import time
import pandas as pd
from utils import forzar_caps_off

def ejecutar_ajuste(df, total, log_func, progress_func, velocidad):
    """
    Robot de Ajuste de Stock (3-6-2).
    Soporta cantidades negativas y lógica u/g según Columna D.
    """
    pyautogui.PAUSE = velocidad
    forzar_caps_off()
    
    if df.empty:
        log_func("❌ ERROR: El archivo de Ajuste no tiene datos.")
        return False

    # 1. NAVEGACIÓN (3-6-2)
    log_func("Entrando al módulo de Ajuste (3-6-2)...")
    for k in ['3', '6', '2']: 
        pyautogui.write(k); pyautogui.press('enter')
        time.sleep(0.5)
    
    # 2. CABECERA (Doc en C1, Motivo en C2, Tipo en C3)
    try:
        pedido = str(df.iloc[0, 2]).split('.')[0].strip() if not pd.isna(df.iloc[0, 2]) else ""
        obs    = str(df.iloc[1, 2]).strip() if not pd.isna(df.iloc[1, 2]) else ""
        imp    = str(df.iloc[2, 2]).upper().strip() if not pd.isna(df.iloc[2, 2]) else "AS" 
        
        log_func(f"📌 Ajuste Cabecera: Doc {pedido} | Motivo: {obs}")

        pyautogui.write(pedido); pyautogui.press('enter'); time.sleep(0.8)
        pyautogui.press('enter') 
        pyautogui.write(obs); pyautogui.press('enter'); time.sleep(0.5)
        pyautogui.press('enter')
        pyautogui.write(imp); pyautogui.press('enter'); time.sleep(1.2)
        
        # Entrada a la grilla
        pyautogui.write(pedido); pyautogui.press('enter'); time.sleep(1.5)
        
    except Exception as e:
        log_func(f"❌ Error en cabecera de ajuste: {e}")
        return False

    # 3. BUCLE DE CARGA
    log_func("⏳ Iniciando carga de ajustes...")
    for i, row in df.iterrows():
        val_sku = row.iloc[0]
        if pd.isna(val_sku): continue
        
        try:
            sku = str(val_sku).split('.')[0].strip()
            # Cantidad (Acepta negativos de Excel: -10, -50, etc.)
            cantidad = str(int(float(str(row.iloc[1]).strip()))) if not pd.isna(row.iloc[1]) else "0"
            
            # Lógica u/g (Columna D)
            info_d = str(row.iloc[3]).strip() if len(row) > 3 and not pd.isna(row.iloc[3]) and str(row.iloc[3]).lower() != 'nan' else None

            pyautogui.write(sku)
            pyautogui.press('enter', presses=4, interval=0.1)

            if info_d:
                # Caso 'g' (Galpón/Pesables)
                pyautogui.write('g')
                pyautogui.write(cantidad); pyautogui.press('enter')
                time.sleep(0.4)
                pyautogui.write(info_d); pyautogui.press('enter')
            else:
                # Caso 'u' (Unidad)
                pyautogui.write(f"u{cantidad}"); pyautogui.press('enter')

            time.sleep(0.1) 
            log_func(f"✅ Fila {i+1}: {sku} -> Cant: {cantidad}")

        except Exception as row_err:
            log_func(f"⚠️ Error fila {i+1}: {row_err}")
            
        progress_func((i + 1) / total)

    # 4. GUARDADO Y RESET (Tu secuencia: End, Enter, End, End)
    log_func("💾 Guardando Ajuste...")
    pyautogui.press('f5'); time.sleep(2.5)
    
    for k in ['end', 'enter', 'end', 'end']:
        pyautogui.press(k); time.sleep(0.5)

    log_func("🏁 Ajuste finalizado con éxito.")
    return True