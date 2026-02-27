import pyautogui
import time
import pandas as pd
from utils import forzar_caps_off, f_monto

def ejecutar_cheques(df, total, log_func, progress_func, velocidad):
    """
    Robot de Cheques V6 - Carga, Cálculo de APF/COF y Cierre de Operación.
    A2(0,0): Entidad | A3(1,0): Comisión | Col J(9): Montos para Suma.
    """
    pyautogui.PAUSE = velocidad
    forzar_caps_off()
    
    def limpiar_fecha(v):
        if pd.isna(v) or str(v).strip() == "": return ""
        try: return pd.to_datetime(v).strftime('%d%m%Y')
        except: return str(v).replace('-', '').replace('/', '').replace(' ', '').strip()

    log_func("💰 Iniciando Carga y Liquidación Final de Cheques...")
    
    try:
        # 1. EXTRACCIÓN DE CABECERA Y CÁLCULOS
        entidad = str(df.iloc[0, 0]).split('.')[0].strip()
        comision = float(df.iloc[1, 0]) if not pd.isna(df.iloc[1, 0]) else 0.0
        
        # Sumamos todos los montos de la columna J (índice 9)
        suma_cheques = df.iloc[0:, 9].sum()
        
        # APF = Comisión - Suma de Cheques (Suele ser valor negativo)
        valor_negativo_apf = comision - suma_cheques
        
        log_func(f"📌 Entidad: {entidad} | Comisión: {comision}")
        log_func(f"📊 Suma Cheques: {suma_cheques} | APF: {valor_negativo_apf}")

        # 2. NAVEGACIÓN INICIAL PuTTY
        pyautogui.write('2')
        pyautogui.press('enter', presses=6, interval=0.1)
        time.sleep(0.5)
        
        pyautogui.write(entidad); pyautogui.press('enter')
        pyautogui.write('afd'); pyautogui.press('enter')
        pyautogui.write('0'); pyautogui.press('enter')
        time.sleep(1.2)

        # 3. BUCLE DE CARGA DE GRILLA
        for i, row in df.iterrows():
            # Saltamos si la Referencia (Col B / 1) está vacía
            if pd.isna(row.iloc[1]) or str(row.iloc[1]).strip() == "": continue 
            
            ref      = str(row.iloc[1]).split('.')[0].strip()
            serie    = str(row.iloc[2]).strip()
            nro_ch   = str(row.iloc[3]).split('.')[0].strip()
            f_orig   = limpiar_fecha(row.iloc[4])
            f_depo   = limpiar_fecha(row.iloc[5])
            banco    = str(row.iloc[6]).strip()
            nombre_h = str(row.iloc[7]).strip() if not pd.isna(row.iloc[7]) else ""
            cuit_i   = str(row.iloc[8]).split('.')[0].strip() if not pd.isna(row.iloc[8]) else ""
            monto_j  = f_monto(row.iloc[9])

            log_func(f"▶️ Cheque {nro_ch} | $ {monto_j}")

            pyautogui.write(ref); pyautogui.press('enter'); time.sleep(0.5) 
            pyautogui.write(serie); pyautogui.press('enter')
            pyautogui.write(nro_ch); pyautogui.press('enter')
            pyautogui.write(f_orig); pyautogui.press('enter')
            pyautogui.write(f_depo); pyautogui.press('enter')
            pyautogui.write(banco); pyautogui.press('enter')
            
            with pyautogui.hold('shift'): pyautogui.press('t')
            pyautogui.press('enter')
            
            pyautogui.write(nombre_h); pyautogui.press('enter')
            pyautogui.write(cuit_i)
            pyautogui.press('enter', presses=3, interval=0.1)
            pyautogui.write(monto_j); pyautogui.press('enter')
            
            progress_func((i + 1) / total)

        # 4. SECUENCIA DE CIERRE (LIQUIDACIÓN)
        log_func("⚙️ Ejecutando cierre y balanceo de importes...")
        
        pyautogui.press('f5'); time.sleep(0.5)
        pyautogui.press('enter', presses=2, interval=0.2)
        
        # A. Ingresar Suma Total de Cheques
        pyautogui.write(f_monto(suma_cheques)); pyautogui.press('enter')
        
        # B. Ingresar Comisión (cof)
        pyautogui.write('cof'); pyautogui.press('enter')
        pyautogui.write(f_monto(comision)); pyautogui.press('enter')
        
        # C. Ingresar Ajuste (apf)
        pyautogui.write('apf'); pyautogui.press('enter')
        pyautogui.write(f_monto(valor_negativo_apf)); pyautogui.press('enter')
        
        # D. Salida y validación final
        pyautogui.press('enter', presses=6, interval=0.1)
        pyautogui.press('f5', presses=2, interval=0.3)
        pyautogui.press('enter', presses=2, interval=0.2)
        
        # E. Cierre con la SUMA TOTAL de cheques (Tu pedido especial)
        pyautogui.write(f_monto(suma_cheques))
        
        log_func("✅ Liquidación completada con éxito.")

        # Confirmación visual de seguridad
        pyautogui.confirm(
            f"RESUMEN DE CARGA:\n\n"
            f"Suma de Cheques: {suma_cheques}\n"
            f"Comisión: {comision}\n"
            f"Ajuste APF: {valor_negativo_apf}\n\n"
            "El robot pegó la Suma Total al final. ¿Grabar?",
            "Validación Cheques"
        )
        return True

    except Exception as e:
        log_func(f"❌ Error en Cheques: {e}")
        return False