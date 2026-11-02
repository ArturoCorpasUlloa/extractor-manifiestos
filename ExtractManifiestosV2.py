#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime
import threading

try:
    import pdfplumber
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
except ImportError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pdfplumber", "pandas", "openpyxl", "-q"])
    import pdfplumber
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

class ExtractorManifiestos:
    def __init__(self, root):
        self.root = root
        self.root.title("Extractor de Manifiestos Rras")
        self.root.geometry("1000x750")
        self.root.minsize(800, 600)

        self.carpeta_var = tk.StringVar()
        self.archivos_pdf = []
        self.resultados = []

        self._crear_ui()

    def _crear_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)

        ttk.Label(main_frame, text="📄 Extractor de Manifiestos Ambientales", font=("Helvetica", 16, "bold")).grid(row=0, column=0, columnspan=3, pady=(0, 15), sticky=tk.W)
        ttk.Label(main_frame, text="Solo se procesarán archivos que comiencen con 'Manifiesto'").grid(row=1, column=0, columnspan=3, sticky=tk.W)

        ttk.Label(main_frame, text="Carpeta:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.carpeta_var, width=60).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)
        ttk.Button(main_frame, text="📁 Buscar...", command=self._seleccionar_carpeta).grid(row=2, column=2, pady=5)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E))
        self.btn_procesar = ttk.Button(btn_frame, text="▶️ Procesar PDFs", command=self._iniciar_procesamiento, state=tk.DISABLED)
        self.btn_procesar.pack(side=tk.LEFT, padx=5)
        self.btn_exportar = ttk.Button(btn_frame, text="💾 Exportar a Excel", command=self._exportar_excel, state=tk.DISABLED)
        self.btn_exportar.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🧹 Limpiar", command=self._limpiar).pack(side=tk.LEFT, padx=5)

        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        self.log_text = scrolledtext.ScrolledText(main_frame, width=100, height=20, wrap=tk.WORD, font=("Consolas", 9))
        self.log_text.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        self.status_var = tk.StringVar(value="Listo.")
        ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E))

    def _log(self, mensaje):
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {mensaje}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def _seleccionar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de origen")
        if carpeta:
            self.carpeta_var.set(carpeta)
            archivos_validos = []
            total_pdfs = 0
            
            # Búsqueda recursiva
            for root_dir, dirs, files in os.walk(carpeta):
                for file in files:
                    if file.lower().endswith(".pdf"):
                        total_pdfs += 1
                        # CONDICIÓN: Debe comenzar con 'Manifiesto'
                        if file.lower().startswith("manifiesto"):
                            archivos_validos.append(os.path.join(root_dir, file))
            
            self.archivos_pdf = archivos_validos
            self._log(f"Carpeta: {carpeta}")
            self._log(f"PDFs totales encontrados: {total_pdfs}")
            self._log(f"PDFs que cumplen filtro ('Manifiesto #'): {len(self.archivos_pdf)}")
            
            if self.archivos_pdf:
                self.btn_procesar.config(state=tk.NORMAL)
                self.status_var.set(f"{len(self.archivos_pdf)} archivos listos.")
            else:
                self.btn_procesar.config(state=tk.DISABLED)
                messagebox.showwarning("Sin coincidencias", "No se encontraron PDFs que comiencen con 'Manifiesto #'.")

    def _limpiar(self):
        self.carpeta_var.set("")
        self.archivos_pdf = []
        self.resultados = []
        self.log_text.delete(1.0, tk.END)
        self.progress['value'] = 0
        self.btn_procesar.config(state=tk.DISABLED)
        self.btn_exportar.config(state=tk.DISABLED)

    def _extraer_datos_pdf(self, ruta_pdf):
        filas_archivo = []
        base = {
            'N_MANIFIESTO': '', 'CLIENTE': '', 'NIT': '', 'SEDE': '', 'DIRECCION': '',
            'FECHA': '', 'DEPARTAMENTO': '', 'MUNICIPIO': '',
            'TOTAL_KG_MANIFIESTO': '', 'EMPAQUE': '', 'ADECUADO': '', 'TOTAL_EMPAQUE': '',
            'DOTACION_GENERADOR': '', 'DOTACION_OPERADOR': '', 'PLACA_VEHICULO': '', 
            'DIAS_ALMACENADOS': '', 'NOMBRE_ARCHIVO': os.path.basename(ruta_pdf)
        }

        try:
            with pdfplumber.open(ruta_pdf) as pdf:
                page = pdf.pages[0]
                text = page.extract_text() or ""
                tables = page.extract_tables()

                # --- Regex Extracción ---
                m = re.search(r'MANIFIESTO No\.?\s*(\d+)', text)
                if m: base['N_MANIFIESTO'] = m.group(1)

                m = re.search(r'CLIENTE\s+(.*?)(?:\n|NIT)', text, re.S)
                if m: base['CLIENTE'] = m.group(1).strip().replace('\n', ' ')

                m = re.search(r'NIT\s+(\d+)', text)
                if m: base['NIT'] = m.group(1)

                m = re.search(r'SEDE\s+(.*?)(?:\n|DIRECCION)', text, re.S)
                if m: base['SEDE'] = m.group(1).strip().replace('\n', ' ')

                m = re.search(r'DIRECCION\s+(.*?)(?:\n|FECHA)', text, re.S)
                if m: base['DIRECCION'] = m.group(1).strip().replace('\n', ' ')

                m = re.search(r'FECHA\s+(\d{2}/\d{2}/\d{4}.*?)(?:\n|DEPARTAMENTO)', text, re.S)
                if m: base['FECHA'] = m.group(1).strip().replace('\n', ' ')

                m = re.search(r'DEPARTAMENTO\s+(.*?)(?:\n|MUNICIPIO)', text, re.S)
                if m: base['DEPARTAMENTO'] = m.group(1).strip()

                m = re.search(r'MUNICIPIO\s+(.*?)(?:\n|TIPO DE RESIDUO)', text, re.S)
                if m: base['MUNICIPIO'] = m.group(1).strip()

                m = re.search(r'PLACA DEL VEHÍCULO\s+([A-Z0-9]+)', text)
                if m: base['PLACA_VEHICULO'] = m.group(1)

                m = re.search(r'DÍAS ALMACENADOS\s+(\d+)', text)
                if m: base['DIAS_ALMACENADOS'] = m.group(1)

                # --- Tablas (Lógica multilínea corregida) ---
                residuos_extraidos = []
                for table in tables:
                    if not table: continue
                    for i, row in enumerate(table):
                        row_str = [str(c) if c else "" for c in row]
                        
                        # Detección de Tabla de Residuos
                        if any("TIPO DE RESIDUO" in s for s in row_str):
                            for r_idx in range(i + 1, len(table)):
                                r_data = [str(c).strip() if c else "" for c in table[r_idx]]
                                if not r_data or not r_data[0] or "TOTAL" in r_data[0].upper():
                                    if r_data and "TOTAL" in r_data[0].upper():
                                        # Intentar capturar el peso total de la fila de totales
                                        for val in r_data:
                                            if "(KG)" in val: base['TOTAL_KG_MANIFIESTO'] = val
                                    break
                                residuos_extraidos.append({
                                    'TIPO_RESIDUO': r_data[0],
                                    'CODIGO': r_data[1] if len(r_data)>1 else '',
                                    'ESTADO': r_data[2] if len(r_data)>2 else '',
                                    'CLASE': r_data[3] if len(r_data)>3 else '',
                                    'CANTIDAD': r_data[4] if len(r_data)>4 else ''
                                })

                        # Detección de Empaques
                        if any("BOLSA" in s for s in row_str):
                            base['EMPAQUE'] = row_str[0]
                            base['ADECUADO'] = row_str[1] if len(row_str)>1 else ''
                            base['TOTAL_EMPAQUE'] = row_str[2] if len(row_str)>2 else ''

                if not residuos_extraidos:
                    filas_archivo.append(base)
                else:
                    for res in residuos_extraidos:
                        nf = base.copy()
                        nf.update(res)
                        filas_archivo.append(nf)

        except Exception as e:
            self._log(f"Error en {os.path.basename(ruta_pdf)}: {e}")
        return filas_archivo

    def _iniciar_procesamiento(self):
        self.btn_procesar.config(state=tk.DISABLED)
        self.resultados = []
        threading.Thread(target=self._worker_procesar, daemon=True).start()

    def _worker_procesar(self):
        total = len(self.archivos_pdf)
        for i, ruta in enumerate(self.archivos_pdf, 1):
            self._log(f"Procesando ({i}/{total}): {os.path.basename(ruta)}")
            self.resultados.extend(self._extraer_datos_pdf(ruta))
            self.progress['value'] = (i / total) * 100
        
        self.btn_procesar.config(state=tk.NORMAL)
        if self.resultados: self.btn_exportar.config(state=tk.NORMAL)
        self._log(f"Proceso terminado. {len(self.resultados)} filas listas para exportar.")

    def _exportar_excel(self):
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not ruta: return
        try:
            df = pd.DataFrame(self.resultados)
            cols = ['N_MANIFIESTO', 'FECHA', 'CLIENTE', 'NIT', 'SEDE', 'DIRECCION', 'DEPARTAMENTO', 'MUNICIPIO', 'TIPO_RESIDUO', 'CODIGO', 'ESTADO', 'CLASE', 'CANTIDAD', 'TOTAL_KG_MANIFIESTO', 'EMPAQUE', 'ADECUADO', 'TOTAL_EMPAQUE', 'PLACA_VEHICULO', 'DIAS_ALMACENADOS', 'NOMBRE_ARCHIVO']
            df = df[[c for c in cols if c in df.columns]]
            df.to_excel(ruta, index=False)
            
            # Autoajuste de columnas básico
            wb = load_workbook(ruta)
            ws = wb.active
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                    except: pass
                ws.column_dimensions[column].width = min(max_length + 2, 60)
            wb.save(ruta)
            
            messagebox.showinfo("Éxito", f"Excel guardado con {len(df)} registros.")
            os.startfile(os.path.dirname(ruta))
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExtractorManifiestos(root)
    root.mainloop()