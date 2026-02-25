import customtkinter as ctk
import pandas as pd
import pyautogui
import pygetwindow as gw
import time
import os
import shutil
import logging
import sys
import threading
from tkinter import filedialog, messagebox
from dotenv import load_dotenv, set_key

# --- IMPORTACIÓN DE MÓDULOS DE ROBOTS ---
from robots.Robot_Putty import ejecutar_stock
from robots.ajuste import ejecutar_ajuste
from robots.Cheques import ejecutar_cheques
from robots.Precios_V2 import ejecutar_precios_v2

# --- CONFIGURACIÓN DE RUTAS ---
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ENV_PATH = os.path.join(BASE_DIR, ".env")
load_dotenv(ENV_PATH)

# Rutas de carpetas automáticas
PATH_PROCESADOS_RAIZ = os.path.join(BASE_DIR, "procesados")
PATH_ENTRADA = os.path.join(BASE_DIR, "Excel_Entrada")

# Configuración de Logging
logging.basicConfig(
    filename=os.path.join(BASE_DIR, "operaciones_rpa.log"),
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

pyautogui.FAILSAFE = True 

def asegurar_carpetas():
    """Crea la estructura de carpetas necesaria para la portabilidad."""
    # Crear carpeta de entrada principal
    if not os.path.exists(PATH_ENTRADA):
        os.makedirs(PATH_ENTRADA, exist_ok=True)
        
    # Crear carpetas de salida (procesados)
    for folder in ["stock", "precios", "ajuste", "cheques"]:
        os.makedirs(os.path.join(PATH_PROCESADOS_RAIZ, folder), exist_ok=True)

# Ejecutar la creación de entorno antes de iniciar la interfaz
asegurar_carpetas()

# --- CLASE DE LOGIN ---
class LoginWindow(ctk.CTkFrame):
    def __init__(self, master, login_callback):
        super().__init__(master)
        self.login_callback = login_callback
        self.pack(pady=100, padx=100, fill="both", expand=True)

        ctk.CTkLabel(self, text="🔐 Acceso PuTTY", font=("Arial", 24, "bold")).pack(pady=20)
        
        self.user_entry = ctk.CTkEntry(self, placeholder_text="Usuario", width=250)
        self.user_entry.insert(0, os.getenv("USER_PUTTY", ""))
        self.user_entry.pack(pady=10)
        self.user_entry.bind("<Return>", lambda event: self.validar())

        self.pass_entry = ctk.CTkEntry(self, placeholder_text="Contraseña", show="*", width=250)
        self.pass_entry.pack(pady=10)
        self.pass_entry.bind("<Return>", lambda event: self.validar())

        ctk.CTkButton(self, text="INGRESAR", command=self.validar, fg_color="#3498db").pack(pady=20)

    def validar(self):
        u, p = self.user_entry.get(), self.pass_entry.get()
        if u and p:
            set_key(ENV_PATH, "USER_PUTTY", u)
            self.login_callback(u, p)
        else:
            messagebox.showwarning("Error", "Ingrese sus credenciales")

# --- CLASE PRINCIPAL ---
class AndresRPASuite(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Andrés Díaz | RPA Suite v5.0 PRO")
        self.geometry("1100x800")
        
        ctk.set_appearance_mode("dark")
        
        self.user_putty = None
        self.pass_putty = None
        self.velocidad_tipeo = 0.3
        
        self.login_screen = LoginWindow(self, self.iniciar_suite)

    def iniciar_suite(self, user, password):
        self.user_putty, self.pass_putty = user, password
        self.login_screen.destroy()
        self.setup_ui()

    def setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Sidebar
        self.sidebar = ctk.CTkFrame(self, width=240, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        ctk.CTkLabel(self.sidebar, text="🤖 CONTROL PANEL", font=("Arial", 20, "bold")).pack(pady=20)
        
        self.btn_st = ctk.CTkButton(self.sidebar, text="📦 Carga Stock", command=lambda: self.set_modo("STOCK"))
        self.btn_st.pack(pady=8, padx=20, fill="x")
        
        self.btn_pr = ctk.CTkButton(self.sidebar, text="💰 Precios (V2 Sensor)", command=lambda: self.set_modo("PRECIOS"))
        self.btn_pr.pack(pady=8, padx=20, fill="x")
        
        self.btn_aj = ctk.CTkButton(self.sidebar, text="🔧 Ajuste Stock", command=lambda: self.set_modo("AJUSTE"))
        self.btn_aj.pack(pady=8, padx=20, fill="x")

        self.btn_ch = ctk.CTkButton(self.sidebar, text="🎫 Cheques", command=lambda: self.set_modo("CHEQUES"))
        self.btn_ch.pack(pady=8, padx=20, fill="x")

        # Velocidad
        ctk.CTkLabel(self.sidebar, text="⚡ Velocidad (ms)").pack(pady=(30, 0))
        self.slider = ctk.CTkSlider(self.sidebar, from_=0.05, to=0.8, command=self.act_vel)
        self.slider.set(0.3); self.slider.pack(pady=10, padx=20)
        self.lbl_speed = ctk.CTkLabel(self.sidebar, text="0.30s"); self.lbl_speed.pack()

        # Emergencia
        ctk.CTkLabel(self.sidebar, text="🚨 EMERGENCIA:\nMouse a esquina SUP-IZQ", 
                      text_color="#e74c3c", font=("Arial", 11, "bold")).pack(side="bottom", pady=20)

        # Panel Central
        self.main = ctk.CTkFrame(self, fg_color="transparent")
        self.main.grid(row=0, column=1, padx=30, pady=30, sticky="nsew")
        
        self.lbl_archivo = ctk.CTkLabel(self.main, text="Seleccione un archivo Excel", font=("Arial", 14))
        self.lbl_archivo.pack(pady=10)
        
        ctk.CTkButton(self.main, text="Buscar Excel", command=self.seleccionar_archivo).pack(pady=10)

        self.progress_bar = ctk.CTkProgressBar(self.main, width=600)
        self.progress_bar.set(0); self.progress_bar.pack(pady=20)

        self.btn_run = ctk.CTkButton(self.main, text="🚀 INICIAR", state="disabled", fg_color="#2ecc71", command=self.confirmar_inicio)
        self.btn_run.pack(pady=10)

        self.console = ctk.CTkTextbox(self.main, height=300, font=("Consolas", 11))
        self.console.insert("0.0", "--- SISTEMA RPA INICIADO ---\n")
        self.console.pack(padx=20, pady=20, fill="both", expand=True)

        self.modo, self.archivo_ruta = None, None

    def act_vel(self, v):
        self.velocidad_tipeo = float(v)
        self.lbl_speed.configure(text=f"{v:.2f}s")

    def log(self, txt):
        msg = f"[{time.strftime('%H:%M:%S')}] {txt}"
        self.console.insert("end", f"{msg}\n")
        self.console.see("end")
        logging.info(txt)

    def seleccionar_archivo(self):
        f = filedialog.askopenfilename(
            initialdir=PATH_ENTRADA,
            filetypes=[("Excel files", "*.xlsx")]
        )
        if f:
            self.archivo_ruta = f
            self.lbl_archivo.configure(text=os.path.basename(f), text_color="cyan")
            if self.modo: self.btn_run.configure(state="normal")

    def set_modo(self, m):
        self.modo = m
        self.log(f"Módulo {m} seleccionado.")
        if self.archivo_ruta: self.btn_run.configure(state="normal")

    def confirmar_inicio(self):
        respuesta = messagebox.askyesno("Confirmación de Seguridad", 
                                        f"¿Estás seguro de que PuTTY está en el MENU INICIAL para iniciar {self.modo}?")
        if respuesta:
            self.lanzar_hilo()
        else:
            self.log("Proceso cancelado por el usuario.")

    def lanzar_hilo(self):
        threading.Thread(target=self.ejecutar_robot, daemon=True).start()

    def enfocar_putty(self):
        ventanas = [w for w in gw.getAllWindows() if "PuTTY" in w.title]
        if ventanas:
            win = ventanas[0]
            if win.isMinimized: win.restore()
            win.activate()
            return True
        return False

    def ejecutar_robot(self):
        self.btn_run.configure(state="disabled")
        if not self.enfocar_putty():
            messagebox.showerror("Error", "No se detectó PuTTY."); self.btn_run.configure(state="normal"); return

        try:
            df = pd.read_excel(self.archivo_ruta, header=None)
            total = len(df)
            self.log(f"Ejecutando {self.modo}...")

            if self.modo == "STOCK":
                ejecutar_stock(df, total, self.log, self.progress_bar.set, self.velocidad_tipeo)
            elif self.modo == "PRECIOS":
                # Aquí llamamos a la versión V2 con sensor para la prueba
                ejecutar_precios_v2(df, total, self.log, self.progress_bar.set, self.velocidad_tipeo)
            elif self.modo == "AJUSTE":
                ejecutar_ajuste(df, total, self.log, self.progress_bar.set, self.velocidad_tipeo)
            elif self.modo == "CHEQUES":
                ejecutar_cheques(df, total, self.log, self.progress_bar.set, self.velocidad_tipeo)

            # Mover archivo a procesados
            dest_dir = os.path.join(PATH_PROCESADOS_RAIZ, self.modo.lower())
            os.makedirs(dest_dir, exist_ok=True)
            dest_path = os.path.join(dest_dir, os.path.basename(self.archivo_ruta))
            
            if os.path.exists(dest_path): os.remove(dest_path)
            shutil.move(self.archivo_ruta, dest_path)
            
            messagebox.showinfo("RPA Suite", f"Proceso {self.modo} finalizado con éxito.")
            self.log(f"✅ Archivo movido a: {dest_path}")

        except Exception as e:
            self.log(f"❌ Error crítico: {e}")
        
        self.btn_run.configure(state="normal")
        self.progress_bar.set(0)

if __name__ == "__main__":
    app = AndresRPASuite()
    app.mainloop()