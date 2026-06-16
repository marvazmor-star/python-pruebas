import os
import sys
import subprocess
import multiprocessing
from concurrent.futures import ThreadPoolExecutor
import win32com.client
from tqdm import tqdm

def limitar_prioridad_proceso():
    if sys.platform == 'win32':
        import win32process
        import win32api
        pid = win32api.GetCurrentProcessId()
        handle = win32api.OpenProcess(win32process.PROCESS_SET_INFORMATION, True, pid)
        win32process.SetPriorityClass(handle, win32process.BELOW_NORMAL_PRIORITY_CLASS)

def convertir_con_libreoffice(archivo_excel, ruta_origen, ruta_destino, posicion_hilo):
    ruta_completa_excel = os.path.join(ruta_origen, archivo_excel)
    comando = ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", ruta_destino, ruta_completa_excel]
    
    nombre_corto = archivo_excel[:20] + "..." if len(archivo_excel) > 20 else archivo_excel
    with tqdm(total=1, desc=f"Hilo #{posicion_hilo} ({nombre_corto})", position=posicion_hilo, leave=False) as pbar:
        try:
            subprocess.run(comando, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            pbar.update(1)
        except Exception:
            pbar.set_description(f"Hilo #{posicion_hilo} [FALLÓ]")

def convertir_bloque_microsoft(archivos, ruta_origen, ruta_destino):
    origen_abs = os.path.abspath(ruta_origen)
    destino_abs = os.path.abspath(ruta_destino)
    convertidos = 0
    
    try:
        excel_app = win32com.client.DispatchEx("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.ScreenUpdating = False
        excel_app.EnableEvents = False
        
        with tqdm(total=len(archivos), desc="Progreso Total MS Excel", unit="archivo") as pbar:
            for archivo in archivos:
                ruta_excel = os.path.join(origen_abs, archivo)
                nombre_base = os.path.splitext(archivo)[0]
                ruta_pdf = os.path.join(destino_abs, f"{nombre_base}.pdf")
                
                wb = None
                try:
                    wb = excel_app.Workbooks.Open(ruta_excel, ReadOnly=True, UpdateLinks=False)
                    wb.ExportAsFixedFormat(0, ruta_pdf)
                    convertidos += 1
                except Exception:
                    pass
                finally:
                    if wb:
                        wb.Close(SaveChanges=False)
                pbar.update(1)
                    
    except Exception as e:
        print(f"\nError crítico en el motor de Excel: {e}")
    finally:
        if excel_app:
            try:
                excel_app.ScreenUpdating = True
                excel_app.DisplayAlerts = True
                excel_app.EnableEvents = True
                excel_app.Quit()
            except:
                pass
    return convertidos

def main():
    print("==================================================================")
    print("       CONVERSOR MASIVO DE EXCEL A PDF (MOTOR DUAL v1.0)          ")
    print("==================================================================\n")
    
    ruta_origen = input("1. Ruta de la carpeta de ORIGEN (Excel): ").strip()
    ruta_destino = input("2. Ruta de la carpeta de DESTINO (PDF): ").strip()
    
    if not os.path.exists(ruta_origen):
        print("La ruta de origen no existe.")
        input("\nPresiona Intro para salir...")
        return
    if not os.path.exists(ruta_destino):
        os.makedirs(ruta_destino)

    archivos = [f for f in os.listdir(ruta_origen) if f.endswith(('.xlsx', '.xls', '.xlsm'))]
    total = len(archivos)
    
    if total == 0:
        print("No se encontraron archivos Excel.")
        input("\nPresiona Intro para salir...")
        return

    nucleos_totales = multiprocessing.cpu_count()
    hilos_80_porcento = max(1, int(nucleos_totales * 0.8))
    
    print(f"\nTu procesador tiene {nucleos_totales} hilos lógicos.")
    print("¿Qué motor de conversión deseas utilizar?")
    print(f" [1] Microsoft Excel Oficial (Fidelidad 100% | Barra única global)")
    print(f" [2] LibreOffice Headless    (Fidelidad Variable | {hilos_80_porcento} Barras simultáneas)")
    opcion = input("Selecciona una opción (1 o 2): ").strip()

    limitar_prioridad_proceso()

    os.system('cls' if os.name == 'nt' else 'clear')
    print(f"Procesando {total} archivos...\n" + "="*50 + "\n")

    if opcion == "1":
        exitos = convertir_bloque_microsoft(archivos, ruta_origen, ruta_destino)
        print(f"\n\n" + "="*50)
        print(f"¡Finalizado! Se convirtieron {exitos}/{total} archivos con MS Excel.")
        
    elif opcion == "2":
        print("Iniciando multihilo. Verás cambiar las tareas dinámicamente:\n")
        with ThreadPoolExecutor(max_workers=hilos_80_porcento) as executor:
            for i, archivo in enumerate(archivos):
                posicion_barra = (i % hilos_80_porcento) + 1
                executor.submit(convertir_con_libreoffice, archivo, ruta_origen, ruta_destino, posicion_barra)
                
        print("\n" * (hilos_80_porcento + 2))
        print("="*50 + "\n¡Finalizado en paralelo con LibreOffice!")
    else:
        print("Opción inválida.")
        
    input("\nPresiona Intro para salir del programa...")

if __name__ == "__main__":
    main()