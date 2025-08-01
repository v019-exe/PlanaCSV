import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import csv
from collections import defaultdict
from datetime import date, datetime
import os
import ctypes
import calendar
import requests
import threading
import tempfile
import shutil
import sys
import subprocess




RANGOS_HORARIOS = {
    '00-08': [str(h).zfill(2) for h in range(0, 8)],
    '08-09': [str(h).zfill(2) for h in range(8, 9)],
    '09-13': [str(h).zfill(2) for h in range(9, 13)],
    '13-18': [str(h).zfill(2) for h in range(13, 18)],
    '18-20': [str(h).zfill(2) for h in range(18, 20)],
    '20-22': [str(h).zfill(2) for h in range(20, 22)],
    '22-23': [str(h).zfill(2) for h in range(22, 24)]
}




SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_FILE_PATH = os.path.join(SCRIPT_DIR, 'data.csv')


VERSION = "1.0.1"  
GITHUB_REPO = "v019-exe/PlanaCSV"  
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
GITHUB_RELEASE_URL = f"https://github.com/{GITHUB_REPO}/releases/latest"

def verificar_actualizacion():
    try:
        response = requests.get(GITHUB_API_URL, timeout=5)
        if response.status_code == 200:
            data = response.json()
            ultima_version = data['tag_name'].lstrip('v')
            return ultima_version != VERSION, ultima_version, data['assets'][0]['browser_download_url']
        return False, VERSION, None
    except Exception:
        return False, VERSION, None

def descargar_actualizacion(url, callback):
    try:
        
        temp_dir = tempfile.mkdtemp()
        temp_file = os.path.join(temp_dir, "nueva_version.exe")
        
        
        response = requests.get(url, stream=True)
        with open(temp_file, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        
        
        current_exe = sys.executable
        batch_file = os.path.join(temp_dir, "actualizar.bat")
        with open(batch_file, 'w') as f:
            f.write('@echo off\n')
            f.write('timeout /t 1 /nobreak > nul\n')  
            f.write(f'copy /Y "{temp_file}" "{current_exe}"\n')
            f.write(f'start "" "{current_exe}"\n')
            f.write(f'del "%~f0"\n')  
        
        
        subprocess.Popen(['cmd', '/c', batch_file])
        sys.exit(0)
        
    except Exception as e:
        callback(False, str(e))
        return



def obtener_fechas_semana(año, mes, semana_objetivo):
    cal = calendar.Calendar(firstweekday=0)  
    try:
        semanas_del_mes = cal.monthdayscalendar(año, mes)
    except calendar.IllegalMonthError:
        return []

    if not (1 <= semana_objetivo <= len(semanas_del_mes)):
        return []

    semana_seleccionada = semanas_del_mes[semana_objetivo - 1]
    
    fechas_formateadas = []
    for dia in semana_seleccionada:
        if dia != 0:
            fechas_formateadas.append(date(año, mes, dia).strftime('%d/%m/%Y'))
    return fechas_formateadas

def contar_llamadas_perdidas_por_rango(fechas_objetivo, archivo_csv):
    conteos = defaultdict(lambda: defaultdict(int))
    
    for fecha in fechas_objetivo:
        for rango in RANGOS_HORARIOS:
            conteos[fecha][rango] = 0

    try:
        with open(archivo_csv, 'r', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile, delimiter=';')
            for row in reader:
                agente = row.get('Agente', '').strip().lower()
                fecha_raw = row.get('Fecha de inicio de la llamada', '').strip()
                hora_raw = row.get('Hora de inicio de la llamada', '').strip()

                if not fecha_raw or not hora_raw:
                    continue

                try:
                    fecha_iso = fecha_raw.split(' ')[0]
                    fecha_formateada = datetime.strptime(fecha_iso, '%Y-%m-%d').strftime('%d/%m/%Y')
                    hora = hora_raw[:2]

                    
                    if agente == 'undefined' and fecha_formateada in fechas_objetivo:
                        for rango, horas_rango in RANGOS_HORARIOS.items():
                            if hora in horas_rango:
                                conteos[fecha_formateada][rango] += 1
                                break
                except (ValueError, IndexError):
                    continue
    except FileNotFoundError:
        messagebox.showerror("Error", f"No se encontró el archivo: {archivo_csv}")
        return None
    return conteos

def contar_llamadas_por_agente_y_dia(fechas_objetivo, archivo_csv):
    contadores = defaultdict(lambda: defaultdict(int))
    try:
        with open(archivo_csv, 'r', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile, delimiter=';')
            for row in reader:
                agente = row.get('Agente', '').strip()
                fecha_raw = row.get('Fecha de inicio de la llamada', '').strip()

                if not fecha_raw or not agente or agente.lower() == 'undefined':
                    continue
                
                try:
                    fecha_iso = fecha_raw.split(' ')[0]
                    fecha_formateada = datetime.strptime(fecha_iso, '%Y-%m-%d').strftime('%d/%m/%Y')
                    
                    if fecha_formateada in fechas_objetivo:
                        contadores[agente][fecha_formateada] += 1
                except ValueError:
                    continue
    except FileNotFoundError:
        
        return None
    return contadores




class AnalizadorCsv(tk.Tk):
    def __init__(self, default_csv_path=""):
        super().__init__()

        myappid = 'plana.analizador.llamadas.1'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

        self.title(f"Analizador de llamadas Plana v{VERSION}")
        self.geometry("500x500")
        self.iconbitmap(os.path.join(SCRIPT_DIR, 'Plana.ico'))
        
        
        self.csv_path_var = tk.StringVar(value=default_csv_path)
        self.año_var = tk.StringVar(value=date.today().year)
        self.mes_var = tk.StringVar(value=date.today().month)
        self.semana_var = tk.IntVar(value=1)
        
        
        self.style = ttk.Style()
        self.style.theme_use('clam')  
        
        
        self.configure(bg="#f0f0f0")
        self.style.configure('TFrame', background="#f0f0f0")
        self.style.configure('TLabel', background="#f0f0f0", font=('Segoe UI', 10))
        self.style.configure('TButton', font=('Segoe UI', 10))
        self.style.configure('Accent.TButton', font=('Segoe UI', 10, 'bold'))
        
        self._crear_widgets()
        
        
        self.after(1000, self.verificar_actualizaciones_silencioso)

    def _crear_widgets(self):
        
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        
        main_frame = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(main_frame, text="Análisis")
        
        
        main_frame.grid_columnconfigure(1, weight=1)
        
        
        titulo_label = ttk.Label(main_frame, text="Analizador de llamadas Plana", 
                               font=('Segoe UI', 14, 'bold'))
        titulo_label.grid(row=0, column=0, columnspan=3, pady=(0, 15), sticky="w")
        
        
        param_frame = ttk.LabelFrame(main_frame, text="Parámetros", padding=10)
        param_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        param_frame.grid_columnconfigure(1, weight=1)
        
        
        ttk.Label(param_frame, text="Año:").grid(row=0, column=0, sticky="w", pady=5, padx=5)
        ttk.Entry(param_frame, textvariable=self.año_var, width=10).grid(row=0, column=1, sticky="w", pady=5)
        
        
        ttk.Label(param_frame, text="Mes:").grid(row=1, column=0, sticky="w", pady=5, padx=5)
        ttk.Entry(param_frame, textvariable=self.mes_var, width=10).grid(row=1, column=1, sticky="w", pady=5)
        
        
        ttk.Label(param_frame, text="Semana:").grid(row=2, column=0, sticky="w", pady=5, padx=5)
        semanas_opciones = [1, 2, 3, 4, 5, 6]
        semana_combo = ttk.Combobox(param_frame, textvariable=self.semana_var, values=semanas_opciones, width=8, state="readonly")
        semana_combo.grid(row=2, column=1, sticky="w", pady=5)
        semana_combo.current(0)
        
        
        file_frame = ttk.LabelFrame(main_frame, text="Archivo de datos", padding=10)
        file_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        file_frame.grid_columnconfigure(0, weight=1)
        
        
        entry_path = ttk.Entry(file_frame, textvariable=self.csv_path_var)
        entry_path.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        ttk.Button(file_frame, text="Buscar...", command=self._seleccionar_archivo).grid(row=0, column=1)
        
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=3, column=0, columnspan=3, sticky="ew")
        btn_frame.grid_columnconfigure(0, weight=1)
        
        ttk.Button(btn_frame, text="Analizar", command=self.realizar_analisis, 
                  style='Accent.TButton', width=20).grid(row=0, column=0, pady=10)
        
        
        update_frame = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(update_frame, text="Actualizaciones")
        
        update_frame.grid_columnconfigure(0, weight=1)
        
        
        ttk.Label(update_frame, text=f"Versión actual: {VERSION}", 
                font=('Segoe UI', 11)).grid(row=0, column=0, sticky="w", pady=(0, 10))
        
        
        ttk.Button(update_frame, text="Verificar actualizaciones", 
                  command=self.verificar_actualizaciones_manual).grid(row=1, column=0, sticky="w", pady=5)
        
        
        self.update_status = ttk.Label(update_frame, text="")
        self.update_status.grid(row=2, column=0, sticky="w", pady=5)
        
        
        status_frame = ttk.Frame(self)
        status_frame.pack(fill="x", side="bottom", padx=10, pady=5)
        
        self.status_label = ttk.Label(status_frame, text=f"Analizador de llamadas Plana v{VERSION}", anchor="w")
        self.status_label.pack(side="left")
        
        
        self.update_button = ttk.Button(status_frame, text="¡Nueva versión disponible!", 
                                      command=self.mostrar_dialogo_actualizacion)
        

    def _seleccionar_archivo(self):
        filepath = filedialog.askopenfilename(
            title="Seleccionar archivo CSV",
            filetypes=(("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*"))
        )
        if filepath:  
            self.csv_path_var.set(filepath)

    def _mostrar_resultados(self, texto_resultados):
        ventana_resultados = tk.Toplevel(self)
        ventana_resultados.title("Resultados del Análisis")
        ventana_resultados.geometry("750x550")
        
        ventana_resultados.transient(self)
        ventana_resultados.grab_set()
        
        try:
            ventana_resultados.iconbitmap(os.path.join(SCRIPT_DIR, 'Plana.ico'))
        except tk.TclError:
            pass
        
        
        main_frame = ttk.Frame(ventana_resultados, padding=15)
        main_frame.pack(fill="both", expand=True)
        
        
        ttk.Label(main_frame, text="Resultados del Análisis", 
                font=('Segoe UI', 14, 'bold')).pack(anchor="w", pady=(0, 15))
        
        
        text_frame = ttk.LabelFrame(main_frame, text="Informe detallado")
        text_frame.pack(fill="both", expand=True)
        
        
        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Consolas", 10), 
                            bg="#f9f9f9", padx=10, pady=10)
        scrollbar = ttk.Scrollbar(text_frame, command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        
        text_widget.insert(tk.END, texto_resultados)
        
        
        texto = texto_resultados.split('\n')
        pos = "1.0"
        for i, linea in enumerate(texto):
            if "===" in linea or linea.startswith("---"):
                text_widget.tag_add(f"titulo{i}", pos, f"{pos.split('.')[0]}.end")
                text_widget.tag_config(f"titulo{i}", font=("Consolas", 10, "bold"), foreground="#0066cc")
            elif linea.startswith("Resultados para"):
                text_widget.tag_add(f"titulo{i}", pos, f"{pos.split('.')[0]}.end")
                text_widget.tag_config(f"titulo{i}", font=("Consolas", 12, "bold"))
            elif linea.startswith("Agente:") or linea.startswith("Día:"):
                text_widget.tag_add(f"subtitulo{i}", pos, f"{pos.split('.')[0]}.end")
                text_widget.tag_config(f"subtitulo{i}", font=("Consolas", 10, "bold"))
            elif "TOTAL" in linea:
                text_widget.tag_add(f"total{i}", pos, f"{pos.split('.')[0]}.end")
                text_widget.tag_config(f"total{i}", font=("Consolas", 10, "bold"), foreground="#006600")
            
            pos = f"{int(pos.split('.')[0]) + 1}.0"
        
        text_widget.config(state='disabled')
        
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x", pady=(15, 0))
        
        ttk.Button(btn_frame, text="Cerrar", 
                 command=ventana_resultados.destroy).pack(side="right")

    def verificar_actualizaciones_silencioso(self):
        def check_bg():
            hay_actualizacion, nueva_version, download_url = verificar_actualizacion()
            if hay_actualizacion:
                self.after(0, lambda: self.mostrar_notificacion_actualizacion(nueva_version, download_url))
        
        
        threading.Thread(target=check_bg, daemon=True).start()
    
    def verificar_actualizaciones_manual(self):
        self.update_status.config(text="Verificando actualizaciones...")
        
        def check_bg():
            hay_actualizacion, nueva_version, download_url = verificar_actualizacion()
            if hay_actualizacion:
                self.after(0, lambda: self.mostrar_dialogo_actualizacion(nueva_version, download_url))
            else:
                self.after(0, lambda: self.update_status.config(
                    text=f"Estás utilizando la última versión ({VERSION})"))
        
        threading.Thread(target=check_bg, daemon=True).start()
    
    def mostrar_notificacion_actualizacion(self, nueva_version, download_url):
        self.nueva_version = nueva_version
        self.download_url = download_url
        self.update_button.pack(side="right", padx=5)
    
    def mostrar_dialogo_actualizacion(self, nueva_version=None, download_url=None):
        if nueva_version is None:
            nueva_version = self.nueva_version
            download_url = self.download_url
            
        respuesta = messagebox.askyesno(
            "Nueva versión disponible",
            f"Hay una nueva versión disponible: v{nueva_version}\n\n"
            f"Tu versión actual es: v{VERSION}\n\n"
            "¿Deseas descargar e instalar la actualización ahora?\n"
            "(La aplicación se reiniciará automáticamente)"
        )
        
        if respuesta:
            self.update_status.config(text="Descargando actualización...")
            
            def on_complete(exito, mensaje):
                if not exito:
                    self.after(0, lambda: messagebox.showerror(
                        "Error de actualización", 
                        f"No se pudo completar la actualización: {mensaje}"))
                    self.after(0, lambda: self.update_status.config(
                        text="Error al actualizar. Intente más tarde."))
            
            threading.Thread(
                target=descargar_actualizacion,
                args=(download_url, on_complete),
                daemon=True
            ).start()
    
    def realizar_analisis(self):
        csv_path = self.csv_path_var.get()
        if not csv_path or not os.path.exists(csv_path):
            messagebox.showerror("Error de Archivo", "La ruta del archivo CSV no es válida o está vacía. Por favor, selecciona un archivo.")
            return

        try:
            
            año = int(self.año_var.get())
            mes = int(self.mes_var.get())
            semana = self.semana_var.get()
        except ValueError:
            messagebox.showerror("Error de Entrada", "El año y el mes deben ser números.")
            return

        
        dias_objetivo = obtener_fechas_semana(año, mes, semana)
        if not dias_objetivo:
            messagebox.showwarning("Aviso", f"No se encontraron días para la semana {semana} del mes {mes}/{año}.")
            return

        conteos_perdidas = contar_llamadas_perdidas_por_rango(dias_objetivo, csv_path)
        conteos_agentes = contar_llamadas_por_agente_y_dia(dias_objetivo, csv_path)

        if conteos_perdidas is None or conteos_agentes is None:
            return 

        
        texto_final = f"Resultados para la Semana {semana} del Mes {mes}, Año {año}\n"
        texto_final += "=" * 50 + "\n\n"

        texto_final += "--- LLAMADAS PERDIDAS POR FRANJA HORARIA ---\n"
        for dia in dias_objetivo:
            texto_final += f'\nDía: {dia}\n'
            total_dia = sum(conteos_perdidas[dia].values())
            for rango in RANGOS_HORARIOS:
                cantidad = conteos_perdidas[dia][rango]
                texto_final += f'  - Rango {rango}: {cantidad} llamadas\n'
            texto_final += f'  TOTAL DÍA: {total_dia} llamadas perdidas\n'

        texto_final += "\n" + "=" * 50 + "\n\n"
        texto_final += "--- LLAMADAS ATENDIDAS POR AGENTE ---\n"
        for agente, fechas_dict in sorted(conteos_agentes.items()):
            texto_final += f'\nAgente: {agente}\n'
            for fecha in dias_objetivo:
                llamadas_dia = fechas_dict.get(fecha, 0)
                if llamadas_dia > 0:
                    texto_final += f'  - {fecha}: {llamadas_dia} llamadas\n'
            total_agente = sum(fechas_dict.values())
            texto_final += f'  TOTAL AGENTE: {total_agente} llamadas\n'

        
        self._mostrar_resultados(texto_final)


if __name__ == "__main__":
    app = AnalizadorCsv(default_csv_path=CSV_FILE_PATH)
    app.mainloop()