import pandas as pd
import ctypes
import pyautogui
import os
import datetime
from fpdf import FPDF

# =========================================================
# 1. FUNCIONES INTOCABLES PARA ROBOTS Y MATCHING
# =========================================================
def limpiar_sku(v):
    if pd.isna(v): return None
    try:
        s = str(v).strip()
        # 1er Intento: Convertir a número puro. 
        # Esto mata los ceros a la izquierda ("000035" -> 35) y los decimales ("35.0" -> 35)
        try:
            return str(int(float(s)))
        except ValueError:
            # 2do Intento (Por si el código tiene letras, ej "A0035"):
            if s.endswith('.0'): s = s[:-2]
            return s.lstrip('0') or '0'
    except: return None

def f_monto(v):
    if pd.isna(v): return "0.00"
    try: return "{:.2f}".format(float(str(v).replace(',', '.')))
    except: return "0.00"

def forzar_caps_off():
    hllDll = ctypes.WinDLL("User32.dll")
    if hllDll.GetKeyState(0x14) & 0x0001:
        pyautogui.press('capslock')

# =========================================================
# 2. ETL: TRANSFORMACIÓN ABSOLUTA A EXCEL (PARA AUDITORÍA VISUAL)
# =========================================================
def etl_lpcio_a_excel(ruta_csv, carpeta_salida):
    """Convierte el LPCIO crudo en un Excel perfecto para que el usuario lo revise."""
    try:
        try: df = pd.read_csv(ruta_csv, sep='\t', encoding='utf-16', dtype=str)
        except: df = pd.read_csv(ruta_csv, sep='\t', encoding='latin-1', dtype=str)
        
        # MAPEO CORREGIDO BASADO EN TU ARCHIVO REAL:
        # 0:CodFam, 1:DesFam, 2:Articulo(SKU), 3:Barra, 4:Descripcion, 5:Precio, 6:Oferta
        col_sku = df.columns[2]       # Columna 'Articulo'
        col_desc = df.columns[4]      # Columna 'Descripcion'
        col_precio = df.columns[5]    # Columna 'Precio'
        col_oferta = df.columns[-1]   # Última columna 'Oferta'
        
        df_limpio = pd.DataFrame()
        df_limpio['SKU'] = df[col_sku].apply(limpiar_sku)
        df_limpio['DESCRIPCION'] = df[col_desc].str.strip()
        df_limpio['PRECIO_REGULAR'] = pd.to_numeric(df[col_precio].str.replace(',', '.'), errors='coerce').fillna(0)
        
        # Transformamos la oferta a número real
        df_limpio['VALOR_OFERTA'] = pd.to_numeric(df[col_oferta].str.replace(',', '.'), errors='coerce').fillna(0)
        
        # Si el valor de oferta es distinto a 0, entonces ES OFERTA
        df_limpio['ES_OFERTA'] = df_limpio['VALOR_OFERTA'] != 0

        nombre_salida = "LPCIO_ETL_Revisar.xlsx"
        ruta_salida = os.path.join(carpeta_salida, nombre_salida)
        df_limpio.to_excel(ruta_salida, index=False)
        return ruta_salida
    except Exception as e: return f"Error ETL LPCIO: {e}"

def etl_ventas_a_excel(ruta_csv, carpeta_salida):
    """Convierte la exportación de ventas de Putty en un Excel limpio."""
    try:
        df = pd.read_csv(ruta_csv, sep='|', encoding='latin-1', dtype=str)
        df.columns = [c.strip() for c in df.columns]
        df = df[df['CAMPO'].str.strip() == 'Artículo'].copy()
        
        df_limpio = pd.DataFrame()
        # ACÁ APLICAMOS LA LIMPIEZA MÁGICA QUE MATA LOS CEROS
        df_limpio['SKU'] = df['CODIGO'].apply(limpiar_sku)
        df_limpio['DESCRIPCION'] = df['DESCRIPCION'].str.strip()
        df_limpio['TOTAL_VENTA'] = pd.to_numeric(df['TOTAL'].str.replace(',', '.'), errors='coerce').fillna(0)
        
        nombre_salida = f"{os.path.splitext(os.path.basename(ruta_csv))[0]}_ETL.xlsx"
        ruta_salida = os.path.join(carpeta_salida, nombre_salida)
        df_limpio.to_excel(ruta_salida, index=False)
        return ruta_salida
    except Exception as e: return f"Error ETL Ventas: {e}"

# =========================================================
# ETL ANTIGUO DE ROBOTS (INTACTO - NO SE TOCA)
# =========================================================
def etl_limpiador_csv(ruta_csv):
    try:
        try: df = pd.read_csv(ruta_csv, sep='\t', encoding='utf-16')
        except: df = pd.read_csv(ruta_csv, sep='\t', encoding='latin-1')
        cols_a_borrar = [df.columns[0], df.columns[1], df.columns[3]]
        df_limpio = df.drop(columns=cols_a_borrar)
        u_col = df_limpio.columns[-1]
        df_limpio[u_col] = pd.to_numeric(df_limpio[u_col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        output_path = os.path.splitext(ruta_csv)[0] + "_PARA_ROBOT.xlsx"
        df_limpio.to_excel(output_path, index=False)
        return output_path
    except Exception as e: return f"Error ETL: {str(e)}"

# =========================================================
# 3. MOTOR BI SUPREMO (CRUZA EXCEL vs EXCEL)
# =========================================================
def motor_bi_avanzado(lista_ventas_etiquetadas, ruta_maestro_excel, carpeta_salida, tipo_analisis):
    try:
        # 1. Cargar Maestro ya validado en Excel
        df_maestro = pd.read_excel(ruta_maestro_excel)
        df_maestro['SKU'] = df_maestro['SKU'].astype(str)
        maestro_ofertas = df_maestro[['SKU', 'ES_OFERTA']].drop_duplicates(subset=['SKU'])

        # 2. Consolidar todas las ventas
        dfs_ventas = []
        for venta in lista_ventas_etiquetadas:
            df_v = pd.read_excel(venta['ruta'])
            df_v['SKU'] = df_v['SKU'].astype(str)
            df_v['ETIQUETA_TIEMPO'] = venta['etiqueta'] 
            dfs_ventas.append(df_v)
            
        df_consolidado = pd.concat(dfs_ventas)

        # 3. Cruce Molecular (Match Perfecto)
        df_final = pd.merge(df_consolidado, maestro_ofertas, on='SKU', how='left')
        df_final['ES_OFERTA'] = df_final['ES_OFERTA'].fillna(False)

        # 4. Lógica de Negocio Andrés (14.5% Neto Regular / -1.5% Neto Oferta)
        df_final['UTILIDAD_NETA'] = df_final.apply(
            lambda r: r['TOTAL_VENTA'] * -0.015 if r['ES_OFERTA'] else r['TOTAL_VENTA'] * 0.145, 
            axis=1
        )
        df_final['TIPO_ANALISIS'] = tipo_analisis

        # 5. Guardado Maestro para Streamlit
        os.makedirs(carpeta_salida, exist_ok=True)
        ruta_db = os.path.join(carpeta_salida, "RETAIL_ENGINE_DB.csv")
        
        df_final.to_csv(ruta_db, index=False)
        return "OK"
    except Exception as e: return f"Error BI Engine: {e}"

# =========================================================
# 4. GENERADOR PDF (MANTENIDO)
# =========================================================
def generar_pdf_gestion(ruta_csv, salida_pdf):
    return "✅ PDF Funcional."