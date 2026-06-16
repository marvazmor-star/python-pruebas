import os
import sys
import subprocess
import multiprocessing
import platform
import getpass
import webbrowser
from datetime import datetime
from concurrent.futures import ProcessPoolExecutor
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from urllib.request import pathname2url

# --- DETECCIÓN AUTOMÁTICA DE LIBREOFFICE EN WINDOWS ---

def buscar_binario_libreoffice():
    """Busca el ejecutable de LibreOffice en las rutas por defecto de Windows."""
    if sys.platform != 'win32':
        return "libreoffice" # En Linux/Mac suele estar en el PATH por defecto
        
    rutas_probables = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"C:\Program Files\LibreOffice\program\libreoffice.exe"
    ]
    for ruta in rutas_probables:
        if os.path.exists(ruta):
            return ruta
    return None

# --- TRABAJADOR DE EXCEL EN PARALELO ---

def trabajador_conversion_excel(args):
    archivo, ruta_origen, ruta_destino = args
    ruta_excel = os.path.abspath(os.path.join(ruta_origen, archivo))
    nombre_base = os.path.splitext(archivo)[0]
    ruta_pdf = os.path.abspath(os.path.join(ruta_destino, f"{nombre_base}.pdf"))
    
    import comtypes.client
    import ctypes
    
    excel_app = None
    try:
        try:
            ctypes.windll.ole32.CoInitializeEx(None, 0x2) # COINIT_APARTMENTTHREADED
        except:
            pass
            
        excel_app = comtypes.client.CreateObject("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.ScreenUpdating = False
        excel_app.Interactive = False
        
        wb = excel_app.Workbooks.Open(ruta_excel, 0, True)
        wb.ExportAsFixedFormat(0, ruta_pdf)
        wb.Close(False)
        
        return True, archivo, None
    except Exception as e:
        return False, archivo, f"Error API Excel: {str(e)}"
    finally:
        if excel_app is not None:
            try: excel_app.Quit()
            except: pass
        try: ctypes.windll.ole32.CoUninitialize()
        except: pass

# --- TRABAJADOR DE LIBREOFFICE EN PARALELO ---

def trabajador_conversion_libreoffice(args):
    archivo, ruta_origen, ruta_destino, indice_proceso, ruta_binario_lo = args
    ruta_completa = os.path.abspath(os.path.join(ruta_origen, archivo))
    nombre_base = os.path.splitext(archivo)[0]
    ruta_destino_abs = os.path.abspath(ruta_destino)
    
    # Validar si encontramos LibreOffice antes de lanzar el comando
    if not ruta_binario_lo:
        return False, archivo, "LibreOffice no está instalado o no se encontró en las rutas por defecto de Windows."

    url_archivo = "file://" + pathname2url(ruta_completa)
    # Perfil aislado único por cada proceso para evitar colisiones de archivos temporales
    env_perfil = f"-env:UserInstallation=file:///{pathname2url(os.path.join(ruta_destino_abs, f'_lo_p_temp_{indice_proceso}'))}"
    
    comando = [
        ruta_binario_lo, env_perfil, "--headless", 
        "--convert-to", "pdf", "--outdir", ruta_destino_abs, url_archivo
    ]
    
    try:
        startupinfo = None
        if sys.platform == 'win32':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0 # SW_HIDE

        resultado = subprocess.run(
            comando, stdout=subprocess.PIPE, stderr=subprocess.PIPE, 
            startupinfo=startupinfo, timeout=50
        )
        if resultado.returncode == 0:
            return True, archivo, None
        else:
            err_output = resultado.stderr.decode('utf-8', errors='ignore').strip()
            return False, archivo, f"LibreOffice mandó código de error: {resultado.returncode}. {err_output}"
    except Exception as e:
        return False, archivo, f"Excepción de comando: {str(e)}"

# --- GENERADOR DE INFORME ---

def generar_y_abrir_informe(d, errores):
    ruta_informe = os.path.join(d['ruta_destino'], f"Informe_Procesamiento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
    filas_errores = ""
    if errores:
        for err in errores:
            filas_errores += f"<tr><td style='font-weight:500;color:#334155;'>{err['archivo']}</td><td><span class='status-pill status-danger'>Fallo Técnico</span></td><td style='color:#64748b;font-family:monospace;font-size:13px;'>{err['motivo']}</td></tr>"
    else:
        filas_errores = "<tr><td colspan='3' style='text-align:center;padding:30px;color:#10b981;font-weight:500;'>✓ Procesamiento limpio. Ningún archivo reportó anomalías.</td></tr>"

    css_styles = """
        :root { --bg-main: #f8fafc; --panel-bg: #ffffff; --text-main: #0f172a; --text-muted: #64748b; --primary: #1e293b; --primary-accent: #3b82f6; --success: #10b981; --danger: #ef4444; --border-color: #e2e8f0; }
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background-color: var(--bg-main); color: var(--text-main); margin: 0; padding: 40px 20px; }
        .wrapper { max-width: 1100px; margin: 0 auto; }
        .header { margin-bottom: 32px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border-color); padding-bottom: 20px; }
        .header h1 { font-size: 26px; font-weight: 700; margin: 0; color: var(--primary); }
        .dashboard-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 32px; }
        .panel { background: var(--panel-bg); border: 1px solid var(--border-color); border-radius: 12px; padding: 24px; }
        .panel h3 { margin-top: 0; margin-bottom: 16px; font-size: 16px; color: var(--text-muted); text-transform: uppercase; }
        .data-list { list-style: none; padding: 0; margin: 0; }
        .data-list li { display: flex; justify-content: space-between; padding: 10px 0; font-size: 14px; border-bottom: 1px solid #f1f5f9; }
        .kpi-container { display: grid; grid-template-columns: repeat(3, 1fr); gap: 24px; margin-bottom: 32px; }
        .kpi-card { background: var(--panel-bg); border: 1px solid var(--border-color); border-radius: 12px; padding: 20px; text-align: center; }
        .kpi-card .kpi-val { font-size: 32px; font-weight: 700; }
        table { width: 100%; border-collapse: separate; border-spacing: 0; border: 1px solid var(--border-color); border-radius: 8px; overflow: hidden; }
        th, td { padding: 14px 18px; font-size: 14px; text-align: left; }
        th { background-color: #f8fafc; color: var(--text-muted); }
        .status-pill { display: inline-flex; padding: 2px 10px; font-size: 12px; font-weight: 600; border-radius: 9999px; }
        .status-danger { background-color: #fef2f2; color: var(--danger); }
    """

    html_content = f"""<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Reporte Corporativo</title><style>{css_styles}</style></head>
<body><div class="wrapper"><div class="header"><h1>Auditoría Masiva de Procesamiento</h1></div>
<div class="dashboard-grid"><div class="panel"><h3>Entorno de Ejecución</h3><ul class="data-list"><li>Servidor/PC: {d['pc']}</li><li>Operador: {d['usuario']}</li><li>Motor Asignado: {d['motor']}</li></ul></div>
<div class="panel"><h3>Métricas Temporales</h3><ul class="data-list"><li>Inicio: {d['hora_inicio']}</li><li>Finalización: {d['hora_fin']}</li></ul></div></div>
<div class="kpi-container"><div class="kpi-card"><div>{d['total_inicial']}</div><div style="color:var(--text-muted);">Archivos en Cola</div></div>
<div class="kpi-card" style="color:var(--success);"><div>{d['total_ok']}</div><div>Procesados Correctamente</div></div>
<div class="kpi-card" style="color:var(--danger);"><div>{d['total_error']}</div><div>Fallo/Descartados</div></div></div>
<div class="panel" style="padding:0;"><table><thead><tr><th>Archivo de Origen</th><th>Estado Técnico</th><th>Razón de la Incidencia</th></tr></thead><tbody>{filas_errores}</tbody></table></div></div></body></html>"""
    
    with open(ruta_informe, "w", encoding="utf-8") as f:
        f.write(html_content)
    webbrowser.open(ruta_informe)

# --- INTERFAZ GRÁFICA ---

class AplicacionConversor:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor Supersónico Siniestro - Excel a PDF")
        self.root.geometry("620x460")
        self.root.resizable(False, False)
        
        self.var_proyecto = tk.StringVar(value="Lote_Masivo_Rendimiento_Extremo")
        self.var_origen = tk.StringVar()
        self.var_destino = tk.StringVar()
        self.var_motor = tk.StringVar(value="excel")
        
        style = ttk.Style()
        style.theme_use('vista')

        padding_opciones = {'padx': 15, 'pady': 6}
        
        ttk.Label(root, text="Identificador del Proyecto corporativo:").pack(anchor="w", **padding_opciones)
        ttk.Entry(root, textvariable=self.var_proyecto, width=85).pack(**padding_opciones)

        ttk.Label(root, text="Directorio Origen (Admite .xls, .xlsx, .xlsm de forma mixta):").pack(anchor="w", **padding_opciones)
        frame_origen = ttk.Frame(root)
        frame_origen.pack(fill="x", **padding_opciones)
        ttk.Entry(frame_origen, textvariable=self.var_origen, width=70).pack(side="left", padx=(0,5))
        ttk.Button(frame_origen, text="Buscar...", command=self.seleccionar_origen).pack(side="left")

        ttk.Label(root, text="Directorio Destino (Depósito de PDFs e Informe):").pack(anchor="w", **padding_opciones)
        frame_destino = ttk.Frame(root)
        frame_destino.pack(fill="x", **padding_opciones)
        ttk.Entry(frame_destino, textvariable=self.var_destino, width=70).pack(side="left", padx=(0,5))
        ttk.Button(frame_destino, text="Buscar...", command=self.seleccionar_destino).pack(side="left")

        ttk.Label(root, text="Selección del Motor de Procesamiento Paralelo:").pack(anchor="w", **padding_opciones)
        frame_radio = ttk.Frame(root)
        frame_radio.pack(anchor="w", padx=15, pady=3)
        ttk.Radiobutton(frame_radio, text="MS Excel (Instancia Múltiple Oculta)", variable=self.var_motor, value="excel").pack(side="left", padx=(0,20))
        ttk.Radiobutton(frame_radio, text="LibreOffice Headless (Entorno Auto-Detectado)", variable=self.var_motor, value="libreoffice").pack(side="left")

        self.btn_procesar = ttk.Button(root, text="INICIAR TRITURACIÓN DE ARCHIVOS MASIVA", command=self.ejecutar_procesamiento)
        self.btn_procesar.pack(pady=15, ipadx=20, ipady=4)

        self.progress_bar = ttk.Progressbar(root, orient="horizontal", length=560, mode="determinate")
        self.progress_bar.pack(pady=(5, 2))
        
        self.lbl_status = ttk.Label(root, text="Arquitectura lista. Soporta mayúsculas/minúsculas y formatos mixtos.", font=("Segoe UI", 9, "italic"))
        self.lbl_status.pack()

    def seleccionar_origen(self):
        ruta = filedialog.askdirectory()
        if ruta: self.var_origen.set(ruta)

    def seleccionar_destino(self):
        ruta = filedialog.askdirectory()
        if ruta: self.var_destino.set(ruta)

    def ejecutar_procesamiento(self):
        if not self.var_origen.get() or not self.var_destino.get():
            messagebox.showerror("Error de entrada", "Las carpetas corporativas de origen y destino son obligatorias.")
            return

        if not os.path.exists(self.var_origen.get()):
            messagebox.showerror("Error", "La ruta origen especificada es inaccesible.")
            return

        # MEJORA: Filtro tolerante a mayúsculas/minúsculas para extensiones xls, xlsx y xlsm de forma nativa
        formatos_validos = ('.xls', '.xlsx', '.xlsm')
        archivos = [f for f in os.listdir(self.var_origen.get()) if os.path.splitext(f)[1].lower() in formatos_validos]
        total_inicial = len(archivos)

        if total_inicial == 0:
            messagebox.showinfo("Cola Vacía", "No se detectaron libros de cálculo (.xls, .xlsx, .xlsm) en la carpeta.")
            return

        self.btn_procesar.config(state="disabled")
        self.root.update()
        
        ruta_lo_binario = buscar_binario_libreoffice() if self.var_motor.get() == "libreoffice" else None

        # Si el usuario eligió LibreOffice y no se encuentra instalado, abortamos antes de romper el programa
        if self.var_motor.get() == "libreoffice" and not ruta_lo_binario:
            messagebox.showerror("Dependencia Faltante", "No se ha detectado ninguna instalación activa de LibreOffice en este PC.\n\nPor favor, usa el motor de MS Excel o instala LibreOffice.")
            self.btn_procesar.config(state="normal")
            return

        datos_informe = {
            'nombre_proyecto': self.var_proyecto.get(),
            'ruta_origen': self.var_origen.get(),
            'ruta_destino': self.var_destino.get(),
            'pc': platform.node(),
            'usuario': getpass.getuser(),
            'motor': "Microsoft Excel Multi-Núcleo" if self.var_motor.get() == "excel" else "LibreOffice Headless Auto-Localizado",
            'hora_inicio': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'total_inicial': total_inicial
        }

        cores = multiprocessing.cpu_count()
        hilos_trabajo = max(1, int(cores * 0.8))
        
        total_ok = 0
        total_error = 0
        errores = []
        contador = 0

        with ProcessPoolExecutor(max_workers=hilos_trabajo) as executor:
            if self.var_motor.get() == "excel":
                tareas = [(f, datos_informe['ruta_origen'], datos_informe['ruta_destino']) for f in archivos]
                resultados = executor.map(trabajador_conversion_excel, tareas)
            else:
                tareas = [(f, datos_informe['ruta_origen'], datos_informe['ruta_destino'], i, ruta_lo_binario) for i, f in enumerate(archivos)]
                resultados = executor.map(trabajador_conversion_libreoffice, tareas)

            for exito, archivo, motivo in resultados:
                if exito:
                    total_ok += 1
                else:
                    total_error += 1
                    errores.append({'archivo': archivo, 'motivo': motivo})
                
                contador += 1
                porcentaje = (contador / total_inicial) * 100
                self.progress_bar['value'] = porcentaje
                self.lbl_status.config(text=f"Progreso: {contador}/{total_inicial} archivos completados ({porcentaje:.1f}%)")
                self.root.update()

        # Limpieza de las carpetas temporales de LibreOffice de forma segura
        if self.var_motor.get() != "excel":
            for i in range(total_inicial):
                p = os.path.join(os.path.abspath(datos_informe['ruta_destino']), f"_lo_p_temp_{i}")
                if os.path.exists(p):
                    import shutil
                    try: shutil.rmtree(p, ignore_errors=True)
                    except: pass

        datos_informe['hora_fin'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        datos_informe['total_ok'] = total_ok
        datos_informe['total_error'] = total_error

        self.lbl_status.config(text="Procesamiento completado con éxito.")
        self.btn_procesar.config(state="normal")

        generar_y_abrir_informe(datos_informe, errores)
        messagebox.showinfo("Lote Finalizado", f"Auditoría terminada.\n\nConvertidos con éxito: {total_ok}\nErrores guardados en reporte: {total_error}")

if __name__ == "__main__":
    multiprocessing.freeze_support()
    root = tk.Tk()
    app = AplicacionConversor(root)
    root.mainloop()
