import os
import sys
import subprocess
import multiprocessing
import platform
import getpass
import webbrowser
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# --- FUNCIONES DE CONVERSIÓN ---

def convertir_con_libreoffice_nativo(archivo_excel, ruta_origen, ruta_destino):
    ruta_completa_excel = os.path.join(ruta_origen, archivo_excel)
    comando = ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", ruta_destino, ruta_completa_excel]
    try:
        subprocess.run(comando, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True, None
    except Exception as e:
        return False, str(e)

def convertir_con_excel_nativo(archivo_excel, ruta_origen, ruta_destino):
    origen_abs = os.path.abspath(ruta_origen)
    destino_abs = os.path.abspath(ruta_destino)
    nombre_base = os.path.splitext(archivo_excel)[0]
    ruta_excel = os.path.join(origen_abs, archivo_excel)
    ruta_pdf = os.path.join(destino_abs, f"{nombre_base}.pdf")
    
    vbs_path = os.path.join(destino_abs, f"_temp_{nombre_base}.vbs")
    
    vbs_code = f"""
    On Error Resume Next
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    excelApp.ScreenUpdating = False
    Set wb = excelApp.Workbooks.Open("{ruta_excel.replace('\\', '\\\\')}", 0, True)
    wb.ExportAsFixedFormat 0, "{ruta_pdf.replace('\\', '\\\\')}"
    wb.Close False
    excelApp.Quit
    """
    try:
        with open(vbs_path, "w", encoding="utf-8") as f:
            f.write(vbs_code)
        resultado = subprocess.run(["cscript", "//NoLogo", vbs_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        
        if os.path.exists(vbs_path):
            os.remove(vbs_path)
            
        if resultado.returncode == 0:
            return True, None
        return False, "Error en la ejecución de la API nativa de MS Excel"
    except Exception as e:
        if os.path.exists(vbs_path):
            os.remove(vbs_path)
        return False, str(e)

# --- GENERADOR DE INFORME PREMIUM HTML5/CSS3 ---

def generar_y_abrir_informe(datos_informe, errores):
    ruta_informe = os.path.join(datos_informe['ruta_destino'], f"Informe_Procesamiento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
    
    filas_errores = ""
    if errores:
        for err in errores:
            filas_errores += f"""
            <tr>
                <td style="font-weight: 500; color: #334155;">{err['archivo']}</td>
                <td><span class="status-pill status-danger">Fallo Técnico</span></td>
                <td style="color: #64748b; font-family: monospace; font-size: 13px;">{err['motivo']}</td>
            </tr>"""
    else:
        filas_errores = """
        <tr>
            <td colspan="3" style="text-align: center; padding: 30px; color: #10b981; font-weight: 500;">
                ✓ No se detectaron anomalías. Todos los archivos se procesaron de forma impecable.
            </td>
        </tr>"""

    html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reporte Técnico de Conversión</title>
    <style>
        :root {{
            --bg-main: #f8fafc;
            --panel-bg: #ffffff;
            --text-main: #0f172a;
            --text-muted: #64748b;
            --primary: #1e293b;
            --primary-accent: #3b82f6;
            --success: #10b981;
            --danger: #ef4444;
            --border-color: #e2e8f0;
        }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg-main);
            color: var(--text-main);
            margin: 0;
            padding: 40px 20px;
        }}
        .wrapper {{ max-width: 1100px; margin: 0 auto; }}
        .header {{ margin-bottom: 32px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border-color); padding-bottom: 20px; }}
        .header h1 {{ font-size: 26px; font-weight: 700; margin: 0; color: var(--primary); letter-spacing: -0.5px; }}
        .header .project-badge {{ background-color: #eff6ff; color: var(--primary-accent); padding: 6px 16px; border-radius: 9999px; font-size: 14px; font-weight: 600; border: 1px solid #bfdbfe; }}
        .dashboard-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 32px; }}
        .panel {{ background: var(--panel-bg); border: 1px solid var(--border-color); border-radius: 12px; padding: 24px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05); }}
        .panel h3 {{ margin-top: 0; margin-bottom: 16px; font-size: 16px; font-weight: 600; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; }}
        .data-list {{ list-style: none; padding: 0; margin: 0; }}
        .data-list li {{ display: flex; justify-content: space-between; padding: 10px 0; font-size: 14px; border-bottom: 1px solid #f1f5f9; }}
        .data-list li:last-child {{ border-bottom: none; }}
        .data-list span.label {{ color: var(--text-muted); font-weight: 500; }}
        .data-list span.value {{ color: var(--text-main); font-weight: 600; word-break: break-all; text-align: right; max-width: 60%; }}
        .kpi-container {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 24px; margin-bottom: 32px; }}
        .kpi-card {{ background: var(--panel-bg); border: 1px solid var(--border-color); border-radius: 12px; padding: 20px; text-align: center; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05); }}
        .kpi-card .kpi-val {{ font-size: 32px; font-weight: 700; margin-bottom: 4px; color: var(--primary); }}
        .kpi-card .kpi-lbl {{ font-size: 13px; font-weight: 500; color: var(--text-muted); }}
        .kpi-card.kpi-success .kpi-val {{ color: var(--success); }}
        .kpi-card.kpi-danger .kpi-val {{ color: var(--danger); }}
        table {{ width: 100%; border-collapse: separate; border-spacing: 0; margin-top: 12px; border-radius: 8px; overflow: hidden; border: 1px solid var(--border-color); }}
        th, td {{ padding: 14px 18px; text-align: left; font-size: 14px; }}
        th {{ background-color: #f8fafc; color: var(--text-muted); font-weight: 600; border-bottom: 1px solid var(--border-color); text-transform: uppercase; font-size: 12px; letter-spacing: 0.5px; }}
        td {{ background-color: #ffffff; border-bottom: 1px solid #f1f5f9; }}
        tr:last-child td {{ border-bottom: none; }}
        .status-pill {{ display: inline-flex; align-items: center; padding: 2px 10px; font-size: 12px; font-weight: 600; border-radius: 9999px; }}
        .status-danger {{ background-color: #fef2f2; color: var(--danger); border: 1px solid #fca5a5; }}
    </style>
</head>
<body>
    <div class="wrapper">
        <div class="header">
            <h1>Auditoría y Reporte de Procesamiento</h1>
            <span class="project-badge">{datos_informe['nombre_proyecto']}</span>
        </div>
        
        <div class="dashboard-grid">
            <div class="panel">
                <h3>Metadatos de Infraestructura</h3>
                <ul class="data-list">
                    <li><span class="label">Estación de Trabajo (PC):</span> <span class="value">{datos_informe['pc']}</span></li>
                    <li><span class="label">Dominio / Grupo:</span> <span class="value">{datos_informe['dominio']}</span></li>
                    <li><span class="label">Operador del Sistema:</span> <span class="value">{datos_informe['usuario']}</span></li>
                    <li><span class="label">Tecnología de Renderizado:</span> <span class="value">{datos_informe['motor']}</span></li>
                </ul>
            </div>
            <div class="panel">
                <h3>Tiempos y Logs de Registro</h3>
                <ul class="data-list">
                    <li><span class="label">Timestamp Inicial:</span> <span class="value">{datos_informe['hora_inicio']}</span></li>
                    <li><span class="label">Timestamp Final:</span> <span class="value">{datos_informe['hora_fin']}</span></li>
                    <li><span class="label">Directorio Origen:</span> <span class="value">{datos_informe['ruta_origen']}</span></li>
                    <li><span class="label">Directorio Destino:</span> <span class="value">{datos_informe['ruta_destino']}</span></li>
                </ul>
            </div>
        </div>

        <div class="kpi-container">
            <div class="kpi-card">
                <div class="kpi-val">{datos_informe['total_inicial']}</div>
                <div class="kpi-lbl">Carga Inicial de Archivos</div>
            </div>
            <div class="kpi-card kpi-success">
                <div class="kpi-val">{datos_informe['total_ok']}</div>
                <div class="kpi-lbl">Compilados Correctamente (PDF)</div>
            </div>
            <div class="kpi-card kpi-danger">
                <div class="kpi-val">{datos_informe['total_error']}</div>
                <div class="kpi-lbl">Excepciones / Errores</div>
            </div>
        </div>

        <div class="panel" style="padding: 0;">
            <div style="padding: 24px 24px 12px 24px;"><h3 style="margin: 0;">Registro de Errores e Incidencias</h3></div>
            <table>
                <thead>
                    <tr>
                        <th style="width: 35%;">Fichero de Origen</th>
                        <th style="width: 15%;">Estado Técnico</th>
                        <th style="width: 50%;">Excepción Detectada</th>
                    </tr>
                </thead>
                <tbody>{filas_errores}</tbody>
            </table>
        </div>
    </div>
</body>
</html>
"""
    with open(ruta_informe, "w", encoding="utf-8") as f:
        f.write(html_content)
    webbrowser.open(ruta_informe)


# --- INTERFAZ GRÁFICA (TKINTER) ---

class AplicacionConversor:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor Masivo Corporativo - Excel a PDF")
        self.root.geometry("620x460") # Expandido un poco para alojar la barra de progreso de forma elegante
        self.root.resizable(False, False)
        
        self.var_proyecto = tk.StringVar(value="Lote_Procesamiento_Excel")
        self.var_origen = tk.StringVar()
        self.var_destino = tk.StringVar()
        self.var_motor = tk.StringVar(value="excel")
        
        style = ttk.Style()
        style.theme_use('vista')

        padding_opciones = {'padx': 15, 'pady': 6}
        
        ttk.Label(root, text="Identificador / Nombre del Proyecto:").pack(anchor="w", **padding_opciones)
        ttk.Entry(root, textvariable=self.var_proyecto, width=85).pack(**padding_opciones)

        ttk.Label(root, text="Directorio Raíz de Origen (Documentos Excel):").pack(anchor="w", **padding_opciones)
        frame_origen = ttk.Frame(root)
        frame_origen.pack(fill="x", **padding_opciones)
        ttk.Entry(frame_origen, textvariable=self.var_origen, width=70).pack(side="left", padx=(0,5))
        ttk.Button(frame_origen, text="Examinar...", command=self.seleccionar_origen).pack(side="left")

        ttk.Label(root, text="Directorio Raíz de Destino (PDFs de Salida y Reporte):").pack(anchor="w", **padding_opciones)
        frame_destino = ttk.Frame(root)
        frame_destino.pack(fill="x", **padding_opciones)
        ttk.Entry(frame_destino, textvariable=self.var_destino, width=70).pack(side="left", padx=(0,5))
        ttk.Button(frame_destino, text="Examinar...", command=self.seleccionar_destino).pack(side="left")

        ttk.Label(root, text="Motor de Renderizado y Conversión:").pack(anchor="w", **padding_opciones)
        frame_radio = ttk.Frame(root)
        frame_radio.pack(anchor="w", padx=15, pady=3)
        ttk.Radiobutton(frame_radio, text="Microsoft Excel (Fidelidad 100%)", variable=self.var_motor, value="excel").pack(side="left", padx=(0,20))
        ttk.Radiobutton(frame_radio, text="LibreOffice Headless (80% Cores)", variable=self.var_motor, value="libreoffice").pack(side="left")

        # --- COMPONENTES DE PROGRESO INTEGRADOS ---
        self.btn_procesar = ttk.Button(root, text="INICIAR PROCESAMIENTO MASIVO", command=self.ejecutar_procesamiento)
        self.btn_procesar.pack(pady=15, ipadx=20, ipady=4)

        self.progress_bar = ttk.Progressbar(root, orient="horizontal", length=560, mode="determinate")
        self.progress_bar.pack(pady=(5, 2))
        
        self.lbl_status = ttk.Label(root, text="Sistema listo. Esperando asignación de carpetas.", font=("Segoe UI", 9, "italic"))
        self.lbl_status.pack()

    def seleccionar_origen(self):
        ruta = filedialog.askdirectory()
        if ruta: self.var_origen.set(ruta)

    def seleccionar_destino(self):
        ruta = filedialog.askdirectory()
        if ruta: self.var_destino.set(ruta)

    def actualizar_interfaz_progreso(self, actual, total):
        """Calcula el porcentaje matemático y refresca visualmente los controles de la GUI"""
        porcentaje = (actual / total) * 100
        self.progress_bar['value'] = porcentaje
        self.lbl_status.config(text=f"Procesando: {actual} de {total} archivos... ({porcentaje:.1f}%)")
        self.root.update() # Fuerza al motor de Windows a redibujar la ventana inmediatamente

    def ejecutar_procesamiento(self):
        if not self.var_origen.get() or not self.var_destino.get():
            messagebox.showerror("Campos Incompletos", "Es mandatorio indicar las rutas de origen y destino corporativo.")
            return
        
        if not os.path.exists(self.var_origen.get()):
            messagebox.showerror("Ruta Inválida", "La ruta de origen especificada no existe en el sistema.")
            return

        archivos = [f for f in os.listdir(self.var_origen.get()) if f.endswith(('.xlsx', '.xls', '.xlsm'))]
        total_inicial = len(archivos)

        if total_inicial == 0:
            messagebox.showinfo("Carga Vacía", "No se han detectado ficheros Excel compatibles en la ruta origen.")
            return

        # Bloqueamos el botón y preparamos la barra
        self.btn_procesar.config(state="disabled")
        self.progress_bar['value'] = 0
        
        datos_informe = {
            'nombre_proyecto': self.var_proyecto.get(),
            'ruta_origen': self.var_origen.get(),
            'ruta_destino': self.var_destino.get(),
            'pc': platform.node(),
            'dominio': os.environ.get('USERDOMAIN', 'Local Network'),
            'usuario': getpass.getuser(),
            'motor': "Microsoft Excel (Motor Nativo)" if self.var_motor.get() == "excel" else "LibreOffice Multi-Core",
            'hora_inicio': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'total_inicial': total_inicial
        }

        if sys.platform == 'win32':
            import ctypes
            ctypes.windll.kernel32.SetPriorityClass(ctypes.windll.kernel32.GetCurrentProcess(), 0x00004000)

        total_ok = 0
        total_error = 0
        errores = []
        contador = 0

        if self.var_motor.get() == "excel":
            for archivo in archivos:
                exito, motivo = convertir_con_excel_nativo(archivo, datos_informe['ruta_origen'], datos_informe['ruta_destino'])
                if exito:
                    total_ok += 1
                else:
                    total_error += 1
                    errores.append({'archivo': archivo, 'motivo': motivo})
                
                contador += 1
                self.actualizar_interfaz_progreso(contador, total_inicial)
        else:
            cores = multiprocessing.cpu_count()
            hilos_80 = max(1, int(cores * 0.8))
            
            with ThreadPoolExecutor(max_workers=hilos_80) as executor:
                futuros = {executor.submit(convertir_con_libreoffice_nativo, f, datos_informe['ruta_origen'], datos_informe['ruta_destino']): f for f in archivos}
                for fut in futuros:
                    archivo = futuros[fut]
                    try:
                        exito, motivo = fut.result()
                        if exito:
                            total_ok += 1
                        else:
                            total_error += 1
                            errores.append({'archivo': archivo, 'motivo': "El binario headless de LibreOffice rechazó el documento."})
                    except Exception as e:
                        total_error += 1
                        errores.append({'archivo': archivo, 'motivo': str(e)})
                    
                    contador += 1
                    self.actualizar_interfaz_progreso(contador, total_inicial)

        datos_informe['hora_fin'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        datos_informe['total_ok'] = total_ok
        datos_informe['total_error'] = total_error

        # Restablecer controles visuales
        self.lbl_status.config(text="Proceso finalizado con éxito.")
        self.btn_procesar.config(state="normal")

        # Generar y abrir informe ejecutivo
        generar_y_abrir_informe(datos_informe, errores)
        messagebox.showinfo("Auditoría Finalizada", f"Proceso concluido.\n\nProcesados OK: {total_ok}\nErrores: {total_error}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AplicacionConversor(root)
    root.mainloop()
