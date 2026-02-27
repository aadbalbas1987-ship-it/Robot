import customtkinter as ctk
import pandas as pd
import pyautogui
import pygetwindow as gw
import time, os, shutil, sys, threading
from tkinter import filedialog, messagebox

from robots.Robot_Putty import ejecutar_stock
from robots.ajuste import ejecutar_ajuste
from robots.Cheques import ejecutar_cheques
from robots.Precios_V2 import ejecutar_precios_v2
from utils import etl_lpcio_a_excel, etl_ventas_a_excel, motor_bi_avanzado, generar_pdf_gestion

if getattr(sys, 'frozen', False): BASE_DIR = os.path.dirname(sys.executable)
else: BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PATH_PROCESADOS = os.path.join(BASE_DIR, "procesados")
PATH_DASHBOARD = os.path.join(BASE_DIR, "Dashboard_Data")
PATH_LISTAS = os.path.join(BASE_DIR, "listas_de_precios")
for f in ["stock", "precios", "ajuste", "cheques"]: os.makedirs(os.path.join(PATH_PROCESADOS, f), exist_ok=True)
os.makedirs(PATH_DASHBOARD, exist_ok=True)
os.makedirs(PATH_LISTAS, exist_ok=True)

class SuiteRPA(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Retail Engine Suite v4.0 - Andrés Díaz")
        self.geometry("950x700")
        self.archivo_ruta = None
        self.maestro_lpcio_excel = None # Almacena el LPCIO cargado en memoria
        self.modo = "STOCK"
        self.velocidad_tipeo = 0.05
        
        # APAGAMOS EL FRENO DE EMERGENCIA DEL MOUSE
        pyautogui.FAILSAFE = False 
        
        self.setup_ui()

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        
        # --- SIDEBAR: ROBOTS ---
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar.grid(row=0, column=0, rowspan=6, sticky="nsew")
        ctk.CTkLabel(self.sidebar, text="🤖 MÓDULOS RPA", font=("Arial", 16, "bold")).pack(pady=(20,10))
        for m in ["STOCK", "PRECIOS", "AJUSTE", "CHEQUES"]:
            ctk.CTkButton(self.sidebar, text=m, command=lambda x=m: self.set_modo(x)).pack(pady=5, padx=10)

        # --- ÁREA CENTRAL: ETL Y BI ---
        self.frame_bi = ctk.CTkFrame(self)
        self.frame_bi.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        
        ctk.CTkLabel(self.frame_bi, text="ETAPA 1: ETL (Transformar CSV a Excel Limpio)", font=("Arial", 14, "bold"), text_color="#f39c12").pack(pady=5)
        ctk.CTkButton(self.frame_bi, text="🧼 Convertir Maestro LPCIO", fg_color="#8e44ad", command=self.ejecutar_etl_lpcio).pack(pady=5)
        ctk.CTkButton(self.frame_bi, text="🧼 Convertir Archivo(s) de Ventas", fg_color="#8e44ad", command=self.ejecutar_etl_ventas).pack(pady=5)

        ctk.CTkLabel(self.frame_bi, text="ETAPA 2: Cargar Maestro en Memoria", font=("Arial", 14, "bold"), text_color="#3498db").pack(pady=(15,5))
        self.btn_load_lpcio = ctk.CTkButton(self.frame_bi, text="📥 Cargar LPCIO (Excel)", fg_color="#2980b9", command=self.cargar_lpcio_memoria)
        self.btn_load_lpcio.pack(pady=5)

        ctk.CTkLabel(self.frame_bi, text="ETAPA 3: Análisis Forense por Períodos", font=("Arial", 14, "bold"), text_color="#2ecc71").pack(pady=(15,5))
        frame_botones = ctk.CTkFrame(self.frame_bi, fg_color="transparent")
        frame_botones.pack(pady=5)
        ctk.CTkButton(frame_botones, text="📅 Diario (Máx 7)", width=120, fg_color="#16a085", command=lambda: self.lanzar_analisis("Diario", 7)).grid(row=0, column=0, padx=5)
        ctk.CTkButton(frame_botones, text="📆 Semanal (Máx 4)", width=120, fg_color="#27ae60", command=lambda: self.lanzar_analisis("Semanal", 4)).grid(row=0, column=1, padx=5)
        ctk.CTkButton(frame_botones, text="📊 Trimestral (Máx 3)", width=120, fg_color="#2ecc71", command=lambda: self.lanzar_analisis("Trimestral", 3)).grid(row=0, column=2, padx=5)
        ctk.CTkButton(frame_botones, text="🌎 Anual (1 Arc.)", width=120, fg_color="#1abc9c", command=lambda: self.lanzar_analisis("Anual", 1)).grid(row=0, column=3, padx=5)

        ctk.CTkLabel(self.frame_bi, text="ETAPA 4: Visualización", font=("Arial", 14, "bold"), text_color="#e74c3c").pack(pady=(15,5))
        ctk.CTkButton(self.frame_bi, text="🚀 LANZAR DASHBOARDS SIMULTÁNEOS", fg_color="#c0392b", hover_color="#922b21", command=self.abrir_dashboard).pack(pady=5)

        # --- LOG Y CONSOLA DE ROBOTS ---
        self.txt_log = ctk.CTkTextbox(self, height=130)
        self.txt_log.grid(row=1, column=1, padx=20, pady=5, sticky="nsew")
        
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.grid(row=2, column=1, padx=20, pady=5, sticky="ew")
        self.progress_bar.set(0)
        
        self.btn_file = ctk.CTkButton(self, text="📁 SELECCIONAR EXCEL PARA ROBOT", command=self.seleccionar_archivo, fg_color="#2c3e50")
        self.btn_file.grid(row=3, column=1, padx=20, pady=5, sticky="ew")
        
        self.btn_run = ctk.CTkButton(self, text="▶ INICIAR ROBOT RPA", command=self.run_thread, fg_color="#d35400", state="disabled")
        self.btn_run.grid(row=4, column=1, padx=20, pady=10, sticky="ew")

    def log(self, msg): 
        self.txt_log.insert("end", f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.txt_log.see("end")

    def set_modo(self, m): 
        self.modo = m
        self.log(f"Modo RPA: {m}")
    
    # --- FLUJO ETL ---
    def ejecutar_etl_lpcio(self):
        arch = filedialog.askopenfilename(title="LPCIO Crudo", filetypes=[("CSV/TXT", "*.csv *.txt")])
        if arch:
            res = etl_lpcio_a_excel(arch, PATH_LISTAS)
            if "Error" not in res: self.log(f"✅ LPCIO convertido a Excel. Verificalo en /listas_de_precios")
            else: self.log(f"❌ {res}")

    def ejecutar_etl_ventas(self):
        archivos = filedialog.askopenfilenames(title="Ventas Crudas", filetypes=[("CSV/TXT", "*.csv *.txt")])
        for arch in archivos:
            res = etl_ventas_a_excel(arch, PATH_LISTAS)
            if "Error" not in res: self.log(f"✅ Venta ETL OK: {os.path.basename(res)}")
            else: self.log(f"❌ {res}")

    def cargar_lpcio_memoria(self):
        arch = filedialog.askopenfilename(title="Seleccionar LPCIO_ETL_Revisar.xlsx", filetypes=[("Excel", "*.xlsx")])
        if arch:
            self.maestro_lpcio_excel = arch
            self.btn_load_lpcio.configure(text="✅ LPCIO CARGADO", fg_color="#27ae60")
            self.log(f"🧠 Maestro de Ofertas en memoria: {os.path.basename(arch)}")

    # --- FLUJO BI ---
    def lanzar_analisis(self, tipo_analisis, max_archivos):
        if not self.maestro_lpcio_excel:
            messagebox.showerror("Error", "Primero debés Cargar el Maestro (Etapa 2).")
            return

        archivos_procesar = []
        for i in range(1, max_archivos + 1):
            msg = f"Seleccioná el archivo {i} de {max_archivos} (Dejar vacío para terminar)"
            arch = filedialog.askopenfilename(title=msg, filetypes=[("Excel ETL", "*.xlsx")])
            if not arch: break 
            
            etiqueta = ctk.CTkInputDialog(text=f"¿Qué etiqueta le ponemos al archivo {i}?\nEj: 'Día Lunes', 'Semana 1', 'Mes Enero'", title="Etiqueta").get_input()
            if not etiqueta: etiqueta = f"Carga {i}"
            
            archivos_procesar.append({'ruta': arch, 'etiqueta': etiqueta})

        if not archivos_procesar: return

        self.log(f"⚙️ Procesando Análisis {tipo_analisis} con {len(archivos_procesar)} archivos...")
        def tarea():
            res = motor_bi_avanzado(archivos_procesar, self.maestro_lpcio_excel, PATH_DASHBOARD, tipo_analisis)
            if res == "OK": 
                self.log(f"✅ Base de datos {tipo_analisis} construida.")
                messagebox.showinfo("Listo", "Datos preparados. Podés lanzar los Dashboards.")
            else: self.log(f"❌ {res}")
        threading.Thread(target=tarea, daemon=True).start()

    def abrir_dashboard(self): 
        threading.Thread(target=lambda: os.system(f'streamlit run "{os.path.join(BASE_DIR, "app_dashboard.py")}"'), daemon=True).start()

    # =======================================================
    # ENFOQUE DE PUTTY BLINDADO (LA SOLUCIÓN AL ERROR DE HOY)
    # =======================================================
    def enfocar_putty(self):
        try:
            ventanas = [w for w in gw.getWindowsWithTitle('PuTTY') if 'PuTTY' in w.title]
            if ventanas:
                win = ventanas[0]
                if win.isMinimized: 
                    win.restore()
                try:
                    win.activate()
                except:
                    pass # Evita que Windows bloquee el comando activate()
                
                time.sleep(1) # Le damos 1 segundo a la PC para traer la ventana al frente
                return True
            return False
        except Exception as e:
            self.log(f"⚠️ Error al buscar la ventana: {e}")
            return False

    # --- FLUJO ROBOTS ---
    def seleccionar_archivo(self):
        self.archivo_ruta = filedialog.askopenfilename(initialdir=PATH_LISTAS, filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.archivo_ruta: 
            self.btn_run.configure(state="normal")
            self.log(f"Cargado RPA: {os.path.basename(self.archivo_ruta)}")

    def run_thread(self): 
        threading.Thread(target=self.ejecutar_robot, daemon=True).start()
    
    def ejecutar_robot(self):
        self.btn_run.configure(state="disabled")
        
        # Validamos primero que PuTTY esté abierto y en foco
        if not self.enfocar_putty():
            self.log("❌ ERROR: No se encontró la ventana de PuTTY. Asegurate de tenerla abierta.")
            self.btn_run.configure(state="normal")
            return
            
        try:
            df = pd.read_excel(self.archivo_ruta, header=None)
            total_filas = len(df)
            args = (df, total_filas, self.log, self.progress_bar.set, self.velocidad_tipeo)
            
            if self.modo == "STOCK": ejecutar_stock(*args)
            elif self.modo == "PRECIOS": ejecutar_precios_v2(*args)
            elif self.modo == "AJUSTE": ejecutar_ajuste(*args)
            elif self.modo == "CHEQUES": ejecutar_cheques(*args)
            
            # --- FUNCIÓN AGREGADA: MOVER ARCHIVO A PROCESADOS ---
            dest_folder = os.path.join(PATH_PROCESADOS, self.modo.lower())
            dest_file = os.path.join(dest_folder, os.path.basename(self.archivo_ruta))
            
            if os.path.exists(dest_file):
                os.remove(dest_file) # Si ya existe uno viejo con ese nombre, lo pisa
                
            shutil.move(self.archivo_ruta, dest_file)
            self.archivo_ruta = None # Limpiamos la variable para evitar reprocesos accidentales
            # ----------------------------------------------------
            
            self.log(f"✅ Módulo {self.modo} Terminado. Archivo movido.")
        except Exception as e: 
            self.log(f"❌ ERROR RPA: {e}")
        finally: 
            self.btn_run.configure(state="normal")
            self.progress_bar.set(0)

if __name__ == "__main__": 
    app = SuiteRPA()
    app.mainloop()