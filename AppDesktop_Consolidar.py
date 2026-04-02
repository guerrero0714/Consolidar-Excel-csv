#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Consolidador de Archivos Excel v4
==================================
- Consolida archivos Excel (.xlsx/.xls/.xlsm) y CSV con la MISMA estructura
- Valida TODAS las hojas de cada archivo y las incluye si coinciden
- Rechaza y reporta archivos con estructura diferente
- Exporta en CSV + Excel
- Panel de diagnóstico, log, y reporte de errores
- Gestión de duplicados y filas vacías
"""

import os, re, json, time, threading, traceback, logging, warnings
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from datetime import datetime
from collections import Counter
import pandas as pd

try:
    import openpyxl
    _EXCEL_OK = True
except ImportError:
    _EXCEL_OK = False

try:
    import chardet
    _CHARDET_OK = True
except ImportError:
    _CHARDET_OK = False

# ─────────────────────────────────────────────
#  LOGGING
# ─────────────────────────────────────────────
log = logging.getLogger("consolidador")
logging.basicConfig(level=logging.DEBUG,
                    format="%(asctime)s [%(levelname)s] %(message)s",
                    handlers=[logging.StreamHandler()])

# ─────────────────────────────────────────────
#  LECTURA SEGURA DE ARCHIVOS
# ─────────────────────────────────────────────
def normalizar_columnas(cols: list) -> list:
    """Normaliza nombres de columna: strip, lower, espacios → _"""
    resultado = []
    for c in cols:
        c = str(c).strip()
        resultado.append(c)
    return resultado


def detectar_encoding(ruta: str) -> str:
    if not _CHARDET_OK:
        return 'utf-8'
    with open(ruta, 'rb') as f:
        r = chardet.detect(f.read(20000))
    enc = r.get('encoding') or 'utf-8'
    if r.get('confidence', 1) < 0.6:
        enc = 'utf-8-sig'
    return enc


def detectar_separador(ruta: str, enc: str) -> str:
    try:
        with open(ruta, 'r', encoding=enc, errors='replace') as f:
            muestra = f.read(5000)
        counts = {sep: muestra.count(sep) for sep in [',', ';', '\t', '|']}
        return max(counts, key=counts.get)
    except:
        return ','


def leer_csv(ruta: str) -> list:
    """Lee un CSV y devuelve lista de tuplas (nombre_hoja, DataFrame)."""
    enc = detectar_encoding(ruta)
    sep = detectar_separador(ruta, enc)
    intentos = [
        dict(encoding=enc, sep=sep, engine='python'),
        dict(encoding=enc, sep=None, engine='python'),
        dict(encoding='utf-8-sig', sep=None, engine='python'),
        dict(encoding='latin-1', sep=None, engine='python'),
    ]
    errores = []
    for kw in intentos:
        try:
            df = pd.read_csv(ruta, dtype=str, **kw)
            if df.empty or len(df.columns) < 2:
                continue
            df.columns = normalizar_columnas(df.columns.tolist())
            return [("CSV", df)]
        except Exception as e:
            errores.append(str(e))
    raise ValueError(f"No se pudo leer el CSV. Intentos: {errores}")


def leer_excel_todas_hojas(ruta: str) -> list:
    """
    Lee TODAS las hojas de un Excel.
    Devuelve lista de tuplas (nombre_hoja, DataFrame).
    Ignora hojas vacías.
    """
    xls = pd.ExcelFile(ruta)
    hojas = []
    for nombre_hoja in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=nombre_hoja, dtype=str)
            # Limpiar valores 'nan' string
            df.replace({'nan': None, 'NaN': None}, inplace=True)
            df.columns = normalizar_columnas(df.columns.tolist())

            # Ignorar hojas vacías o con solo headers
            if df.dropna(how='all').empty:
                continue

            # Ignorar hojas con menos de 2 columnas (probablemente metadata)
            if len(df.columns) < 2:
                continue

            hojas.append((nombre_hoja, df))
        except Exception as e:
            hojas.append((nombre_hoja, e))  # Guardar error para reportar
    return hojas


def leer_archivo_completo(ruta: str) -> list:
    """
    Lee un archivo y devuelve lista de (nombre_hoja, DataFrame).
    Para CSV devuelve una sola "hoja".
    """
    ext = os.path.splitext(ruta)[1].lower()
    if ext in ('.xlsx', '.xls', '.xlsm'):
        return leer_excel_todas_hojas(ruta)
    elif ext == '.csv':
        return leer_csv(ruta)
    raise ValueError(f"Extensión no soportada: {ext}")


# ─────────────────────────────────────────────
#  VALIDACIÓN DE ESTRUCTURA
# ─────────────────────────────────────────────
def columnas_coinciden(cols_ref: list, cols_test: list) -> dict:
    """
    Compara dos listas de columnas.
    Retorna dict con:
        coincide: bool
        faltantes: columnas en ref que no están en test
        sobrantes: columnas en test que no están en ref
    """
    set_ref = set(cols_ref)
    set_test = set(cols_test)
    return {
        'coincide': set_ref == set_test,
        'faltantes': sorted(set_ref - set_test),
        'sobrantes': sorted(set_test - set_ref),
    }


def validar_hojas_internas(hojas: list) -> dict:
    """
    Valida que todas las hojas de un mismo archivo tengan la misma estructura.
    Retorna dict con:
        ok: bool
        estructura_ref: columnas de referencia (primera hoja válida)
        hojas_ok: lista de nombres de hojas válidas
        hojas_error: lista de {hoja, motivo}
    """
    resultado = {
        'ok': True,
        'estructura_ref': [],
        'hojas_ok': [],
        'hojas_error': [],
    }

    # Filtrar hojas válidas (DataFrames, no excepciones)
    hojas_validas = [(n, df) for n, df in hojas if isinstance(df, pd.DataFrame)]
    hojas_con_error = [(n, e) for n, e in hojas if isinstance(e, Exception)]

    for nombre, exc in hojas_con_error:
        resultado['hojas_error'].append({
            'hoja': nombre,
            'motivo': f"Error al leer: {exc}"
        })

    if not hojas_validas:
        resultado['ok'] = False
        return resultado

    # Primera hoja como referencia
    nombre_ref, df_ref = hojas_validas[0]
    resultado['estructura_ref'] = df_ref.columns.tolist()
    resultado['hojas_ok'].append(nombre_ref)

    # Comparar el resto contra la referencia
    for nombre, df in hojas_validas[1:]:
        comp = columnas_coinciden(resultado['estructura_ref'], df.columns.tolist())
        if comp['coincide']:
            resultado['hojas_ok'].append(nombre)
        else:
            detalle = []
            if comp['faltantes']:
                detalle.append(f"Faltan: {comp['faltantes']}")
            if comp['sobrantes']:
                detalle.append(f"Sobran: {comp['sobrantes']}")
            resultado['hojas_error'].append({
                'hoja': nombre,
                'motivo': f"Estructura diferente. {'; '.join(detalle)}"
            })

    return resultado


# ─────────────────────────────────────────────
#  ANÁLISIS DE ARCHIVO
# ─────────────────────────────────────────────
def analizar_archivo(ruta: str) -> dict:
    """
    Analiza un archivo completo (todas las hojas).
    Devuelve metadata del archivo.
    """
    resultado = {
        'ruta': ruta,
        'nombre': os.path.basename(ruta),
        'ok': False,
        'error': None,
        'num_hojas': 0,
        'hojas_ok': [],
        'hojas_error': [],
        'estructura': [],
        'filas_total': 0,
        'advertencias': [],
    }
    try:
        hojas = leer_archivo_completo(ruta)
        resultado['num_hojas'] = len(hojas)

        if not hojas:
            resultado['error'] = "Archivo vacío o sin hojas con datos"
            return resultado

        val = validar_hojas_internas(hojas)
        resultado['estructura'] = val['estructura_ref']
        resultado['hojas_ok'] = val['hojas_ok']
        resultado['hojas_error'] = val['hojas_error']

        if val['hojas_error']:
            for h in val['hojas_error']:
                resultado['advertencias'].append(
                    f"Hoja '{h['hoja']}': {h['motivo']}")

        # Contar filas de hojas válidas
        for nombre, df in hojas:
            if isinstance(df, pd.DataFrame) and nombre in val['hojas_ok']:
                resultado['filas_total'] += len(df)

        # Advertencias por columna (de la primera hoja válida)
        hojas_df = [(n, df) for n, df in hojas
                    if isinstance(df, pd.DataFrame) and n in val['hojas_ok']]
        if hojas_df:
            _, df_ref = hojas_df[0]
            for col in df_ref.columns:
                pct_nulo = df_ref[col].isna().mean()
                if pct_nulo > 0.9:
                    resultado['advertencias'].append(
                        f"Col '{col}': {pct_nulo*100:.0f}% nulos (posible columna vacía)")
                elif pct_nulo > 0.5:
                    resultado['advertencias'].append(
                        f"Col '{col}': {pct_nulo*100:.0f}% nulos")

        resultado['ok'] = len(val['hojas_ok']) > 0

    except Exception as e:
        resultado['error'] = str(e)

    return resultado


# ─────────────────────────────────────────────
#  CONTROL DE PROCESO
# ─────────────────────────────────────────────
class ControlProceso:
    def __init__(self):
        self.paused = self.stopped = False
    def pausar(self):   self.paused = True
    def reanudar(self): self.paused = False
    def detener(self):  self.stopped = True


# ─────────────────────────────────────────────
#  CONSOLIDACIÓN
# ─────────────────────────────────────────────
def consolidar(archivos, estructura_ref, control,
               cb_log, cb_progreso, cb_registros, cb_arch,
               eliminar_dup) -> tuple:
    """
    Consolida todos los archivos.
    Solo incluye archivos/hojas cuya estructura coincida con estructura_ref.
    Devuelve (df_final, lista_errores_detallada)
    """
    datos = []
    errores = []
    total = len(archivos)
    acum = 0

    for i, ruta in enumerate(archivos):
        if control.stopped:
            cb_log("⏹️ Detenido por el usuario.", "error")
            break
        while control.paused:
            time.sleep(0.1)
            if control.stopped:
                break

        nombre = os.path.basename(ruta)
        cb_log(f"▶️ [{i+1}/{total}] {nombre}", "info")
        t0 = time.time()

        try:
            hojas = leer_archivo_completo(ruta)
            hojas_incluidas = 0

            for nombre_hoja, df in hojas:
                if isinstance(df, Exception):
                    msg = f"Hoja '{nombre_hoja}': error al leer — {df}"
                    cb_log(f"   ⚠️  {msg}", "warning")
                    errores.append({
                        'archivo': nombre, 'hoja': nombre_hoja,
                        'tipo': 'Error lectura hoja', 'detalle': msg
                    })
                    continue

                # Validar estructura contra referencia
                comp = columnas_coinciden(estructura_ref, df.columns.tolist())

                if not comp['coincide']:
                    detalle = []
                    if comp['faltantes']:
                        detalle.append(f"Faltan: {comp['faltantes']}")
                    if comp['sobrantes']:
                        detalle.append(f"Sobran: {comp['sobrantes']}")
                    msg = f"Hoja '{nombre_hoja}': estructura diferente. {'; '.join(detalle)}"
                    cb_log(f"   ❌ {msg}", "error")
                    errores.append({
                        'archivo': nombre, 'hoja': nombre_hoja,
                        'tipo': 'Estructura diferente', 'detalle': msg
                    })
                    continue

                # Reordenar columnas según referencia
                df = df[estructura_ref]

                # Limpiar texto
                for col in df.columns:
                    if df[col].dtype == object:
                        df[col] = (df[col].str.strip()
                                   .replace({'': None, 'nan': None, 'NaN': None,
                                             'NULL': None, 'null': None,
                                             'N/A': None, 'n/a': None,
                                             '#N/A': None, '-': None}))

                # Eliminar filas completamente vacías
                mask_vacia = df.isna().all(axis=1)
                n_vacias = mask_vacia.sum()
                if n_vacias > 0:
                    cb_log(f"   🗑️  Hoja '{nombre_hoja}': {n_vacias} filas vacías eliminadas", "warning")
                    df = df[~mask_vacia]

                # Agregar columnas de trazabilidad
                df = df.copy()
                df.insert(0, '__archivo__', nombre)
                df.insert(1, '__hoja__', nombre_hoja)

                datos.append(df)
                hojas_incluidas += 1
                acum += len(df)

            elapsed = time.time() - t0
            if hojas_incluidas > 0:
                cb_log(f"   ✅ {hojas_incluidas} hoja(s) incluida(s) — acum: {acum:,} filas — {elapsed:.2f}s", "success")
            else:
                cb_log(f"   ❌ Ninguna hoja válida en este archivo", "error")
                errores.append({
                    'archivo': nombre, 'hoja': '*',
                    'tipo': 'Archivo rechazado',
                    'detalle': 'Ninguna hoja coincide con la estructura de referencia'
                })

            cb_registros(acum)

        except Exception as e:
            msg = str(e)
            cb_log(f"   ❌ ERROR: {msg}", "error")
            errores.append({
                'archivo': nombre, 'hoja': '*',
                'tipo': 'Error lectura', 'detalle': msg
            })
            errores.append({
                'archivo': nombre, 'hoja': '*',
                'tipo': 'Traceback', 'detalle': traceback.format_exc()
            })

        cb_progreso(i + 1, total)
        cb_arch(i + 1, total)

    if not datos:
        return None, errores

    df_final = pd.concat(datos, ignore_index=True)

    # Eliminar duplicados
    if eliminar_dup:
        antes = len(df_final)
        subset = [c for c in estructura_ref if c in df_final.columns]
        if subset:
            df_final.drop_duplicates(subset=subset, keep='first', inplace=True)
        n_dup = antes - len(df_final)
        if n_dup > 0:
            cb_log(f"🔁 {n_dup:,} duplicados eliminados", "warning")
            errores.append({
                'archivo': 'GLOBAL', 'hoja': '*',
                'tipo': 'Duplicados eliminados', 'detalle': str(n_dup)
            })

    return df_final, errores


# ─────────────────────────────────────────────
#  COLORES
# ─────────────────────────────────────────────
COLOR_DARK  = '#1e2a3a'
COLOR_MID   = '#2c3e50'
COLOR_ACENT = '#3498db'
COLOR_OK    = '#27ae60'
COLOR_WARN  = '#f39c12'
COLOR_ERR   = '#e74c3c'
COLOR_BG    = '#f4f6f9'
COLOR_PANEL = '#ffffff'


# ─────────────────────────────────────────────
#  APLICACIÓN GUI
# ─────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Consolidador de Archivos Excel v4")
        self.geometry("1100x860")
        self.minsize(900, 700)
        self.configure(bg=COLOR_BG)

        self.control = None
        self.hilo = None
        self.df_resultado = None
        self.errores = []
        self.analisis_cache = []
        self.estructura_ref = None  # Columnas de referencia (del primer archivo válido)
        self._preview_ok = False

        self._build_styles()
        self._build_ui()
        self._update_buttons(procesando=False)

    # ── ESTILOS ──────────────────────────────
    def _build_styles(self):
        s = ttk.Style(self)
        s.theme_use('clam')
        s.configure('.', background=COLOR_BG, font=('Segoe UI', 10))
        s.configure('TFrame', background=COLOR_BG)
        s.configure('TLabelframe', background=COLOR_BG)
        s.configure('TLabelframe.Label', font=('Segoe UI', 10, 'bold'), foreground=COLOR_MID)
        s.configure('TButton', font=('Segoe UI', 10, 'bold'), padding=6)
        s.configure('Accent.TButton', background=COLOR_ACENT, foreground='white')
        s.configure('TEntry', fieldbackground='white')
        s.configure('Horizontal.TProgressbar', troughcolor='#dde', background=COLOR_ACENT, thickness=18)

    # ── UI PRINCIPAL ─────────────────────────
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=COLOR_DARK)
        hdr.pack(fill='x')
        tk.Label(hdr, text="📊  CONSOLIDADOR DE ARCHIVOS EXCEL  v4",
                 bg=COLOR_DARK, fg='white', font=('Segoe UI', 14, 'bold')).pack(pady=12)

        # Notebook
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill='both', expand=True, padx=10, pady=8)

        self.tab_config  = ttk.Frame(self.nb)
        self.tab_diag    = ttk.Frame(self.nb)
        self.tab_log     = ttk.Frame(self.nb)
        self.tab_errores = ttk.Frame(self.nb)

        self.nb.add(self.tab_config,  text='⚙️  Configuración')
        self.nb.add(self.tab_diag,    text='🔍  Diagnóstico')
        self.nb.add(self.tab_log,     text='📋  Log')
        self.nb.add(self.tab_errores, text='⚠️  Errores')

        self._build_tab_config()
        self._build_tab_diag()
        self._build_tab_log()
        self._build_tab_errores()
        self._build_footer()

    # ── TAB: CONFIGURACIÓN ───────────────────
    def _build_tab_config(self):
        pad = dict(padx=10, pady=5)

        # Rutas
        fr = ttk.LabelFrame(self.tab_config, text="📁  Rutas", padding=10)
        fr.pack(fill='x', **pad)
        fr.columnconfigure(1, weight=1)

        self.origen_var = tk.StringVar()
        self.salida_var = tk.StringVar()

        filas_ruta = [
            ("Carpeta de origen *", self.origen_var, self._sel_origen),
            ("Carpeta de salida",   self.salida_var, self._sel_salida),
        ]
        for i, (lbl, var, cmd) in enumerate(filas_ruta):
            ttk.Label(fr, text=lbl + ":").grid(row=i, column=0, sticky='w', padx=5, pady=3)
            ttk.Entry(fr, textvariable=var).grid(row=i, column=1, sticky='ew', padx=5)
            ttk.Button(fr, text="📂 Buscar", command=cmd, width=10).grid(row=i, column=2, padx=5)

        # Opciones
        fr2 = ttk.LabelFrame(self.tab_config, text="🛠️  Opciones", padding=10)
        fr2.pack(fill='x', **pad)

        self.elim_dup_var = tk.BooleanVar(value=True)
        self.incl_origen_var = tk.BooleanVar(value=True)
        self.nombre_salida_var = tk.StringVar(value='consolidado')

        ttk.Checkbutton(fr2, text="Eliminar filas duplicadas (basado en columnas de datos)",
                        variable=self.elim_dup_var).grid(row=0, column=0, columnspan=2, sticky='w', padx=5)
        ttk.Checkbutton(fr2, text="Incluir columnas de trazabilidad (__archivo__, __hoja__)",
                        variable=self.incl_origen_var).grid(row=1, column=0, columnspan=2, sticky='w', padx=5)

        ttk.Label(fr2, text="Nombre del archivo de salida:").grid(row=2, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(fr2, textvariable=self.nombre_salida_var, width=30).grid(row=2, column=1, sticky='w', padx=5)

        # Referencia
        fr3 = ttk.LabelFrame(self.tab_config, text="📐  Estructura de referencia", padding=10)
        fr3.pack(fill='x', **pad)

        ttk.Label(fr3, text="La estructura se toma del PRIMER archivo válido. "
                            "Los demás archivos/hojas se validan contra esa referencia.",
                  font=('Segoe UI', 9, 'italic'), wraplength=700).pack(anchor='w')

        self.lbl_ref = tk.Label(fr3, text="(ejecuta el análisis para ver la estructura)",
                                bg=COLOR_BG, fg='#888', font=('Consolas', 9), anchor='w',
                                justify='left', wraplength=900)
        self.lbl_ref.pack(fill='x', pady=5)

        # Botones
        fr4 = ttk.Frame(self.tab_config)
        fr4.pack(fill='x', **pad)

        self.btn_analizar = ttk.Button(fr4, text="🔍 Analizar archivos", command=self._analizar_preview)
        self.btn_analizar.pack(side='left', padx=5)

        self.btn_iniciar = ttk.Button(fr4, text="🚀 Iniciar consolidación",
                                      command=self._iniciar, style='Accent.TButton')
        self.btn_iniciar.pack(side='left', padx=5)

        self.btn_pausar   = ttk.Button(fr4, text="⏸️ Pausar",  command=self._pausar,   state='disabled')
        self.btn_reanudar = ttk.Button(fr4, text="▶️ Reanudar", command=self._reanudar, state='disabled')
        self.btn_detener  = ttk.Button(fr4, text="⏹️ Detener",  command=self._detener,  state='disabled')
        self.btn_pausar.pack(side='left', padx=3)
        self.btn_reanudar.pack(side='left', padx=3)
        self.btn_detener.pack(side='left', padx=3)

        ttk.Button(fr4, text="❌ Salir", command=self.destroy).pack(side='right', padx=5)

    # ── TAB: DIAGNÓSTICO ─────────────────────
    def _build_tab_diag(self):
        ttk.Label(self.tab_diag,
                  text="Haz clic en '🔍 Analizar archivos' para ver la estructura detectada.",
                  font=('Segoe UI', 10, 'italic')).pack(pady=8)

        self.diag_text = ScrolledText(self.tab_diag, font=('Consolas', 9), wrap='word',
                                      bg=COLOR_PANEL, relief='flat')
        self.diag_text.pack(fill='both', expand=True, padx=10, pady=5)
        self.diag_text.tag_config('titulo', foreground=COLOR_MID, font=('Consolas', 9, 'bold'))
        self.diag_text.tag_config('ok',   foreground=COLOR_OK)
        self.diag_text.tag_config('warn', foreground=COLOR_WARN)
        self.diag_text.tag_config('err',  foreground=COLOR_ERR)

    # ── TAB: LOG ─────────────────────────────
    def _build_tab_log(self):
        bar = ttk.Frame(self.tab_log)
        bar.pack(fill='x', padx=10, pady=5)
        ttk.Button(bar, text="💾 Exportar log", command=self._exportar_log).pack(side='left', padx=5)
        ttk.Button(bar, text="🗑️ Limpiar",
                   command=lambda: self.log_text.delete(1.0, tk.END)).pack(side='left', padx=5)

        self.log_text = ScrolledText(self.tab_log, font=('Consolas', 9), wrap='word',
                                     bg='#0d1117', fg='#c9d1d9', relief='flat')
        self.log_text.pack(fill='both', expand=True, padx=10, pady=5)
        for tag, color in [('info', '#58a6ff'), ('success', '#3fb950'),
                           ('warning', '#d29922'), ('error', '#f85149')]:
            self.log_text.tag_config(tag, foreground=color)

    # ── TAB: ERRORES ─────────────────────────
    def _build_tab_errores(self):
        bar = ttk.Frame(self.tab_errores)
        bar.pack(fill='x', padx=10, pady=5)
        ttk.Button(bar, text="📤 Exportar errores CSV", command=self._exportar_errores).pack(side='left', padx=5)
        ttk.Label(bar, text="Filtrar:").pack(side='left', padx=10)
        self.filtro_err_var = tk.StringVar(value='Todos')
        self.combo_filtro = ttk.Combobox(bar, textvariable=self.filtro_err_var, width=22, state='readonly')
        self.combo_filtro['values'] = ['Todos']
        self.combo_filtro.pack(side='left')
        self.combo_filtro.bind('<<ComboboxSelected>>', lambda e: self._refrescar_errores())

        tree_frame = ttk.Frame(self.tab_errores)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=5)
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        cols = ('Archivo', 'Hoja', 'Tipo', 'Detalle')
        self.tree_err = ttk.Treeview(tree_frame, columns=cols, show='headings')
        for c, w in zip(cols, [180, 120, 150, 500]):
            self.tree_err.heading(c, text=c)
            self.tree_err.column(c, width=w, minwidth=60)
        vsb = ttk.Scrollbar(tree_frame, orient='vertical',   command=self.tree_err.yview)
        hsb = ttk.Scrollbar(tree_frame, orient='horizontal', command=self.tree_err.xview)
        self.tree_err.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree_err.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

    # ── FOOTER ───────────────────────────────
    def _build_footer(self):
        ft = tk.Frame(self, bg=COLOR_DARK)
        ft.pack(fill='x', side='bottom')

        self.progress = ttk.Progressbar(ft, length=400, mode='determinate')
        self.progress.pack(side='left', padx=10, pady=6)

        self.lbl_estado    = tk.Label(ft, text="Listo", bg=COLOR_DARK, fg='white', font=('Segoe UI', 9))
        self.lbl_total     = tk.Label(ft, text="Archivos: 0", bg=COLOR_DARK, fg='#aaa', font=('Segoe UI', 9))
        self.lbl_proc      = tk.Label(ft, text="Procesados: 0/0", bg=COLOR_DARK, fg='#aaa', font=('Segoe UI', 9))
        self.lbl_registros = tk.Label(ft, text="Registros: 0", bg=COLOR_DARK, fg='#aaa', font=('Segoe UI', 9))
        self.lbl_errcount  = tk.Label(ft, text="Errores: 0", bg=COLOR_DARK, fg=COLOR_WARN, font=('Segoe UI', 9))

        for lbl in [self.lbl_estado, self.lbl_total, self.lbl_proc,
                    self.lbl_registros, self.lbl_errcount]:
            lbl.pack(side='left', padx=10)

    # ── SELECCIÓN RUTAS ──────────────────────
    def _sel_origen(self):
        r = filedialog.askdirectory()
        if r:
            self.origen_var.set(r)
            self._contar_archivos(r)

    def _sel_salida(self):
        r = filedialog.askdirectory()
        if r:
            self.salida_var.set(r)

    def _contar_archivos(self, ruta):
        try:
            archs = [f for f in os.listdir(ruta)
                     if f.lower().endswith(('.xlsx', '.xls', '.csv', '.xlsm'))
                     and not f.startswith('~$')]
            self.lbl_total.config(text=f"Archivos: {len(archs)}")
            self.lbl_proc.config(text=f"Procesados: 0/{len(archs)}")
        except:
            pass

    # ── LOG HELPERS ──────────────────────────
    def _log(self, msg, nivel='info'):
        ts = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert(tk.END, f"[{ts}] {msg}\n", nivel)
        self.log_text.see(tk.END)
        self.update_idletasks()

    def _log_t(self, msg, nivel='info'):
        self.after(0, lambda: self._log(msg, nivel))

    def _diag(self, msg, tag=''):
        self.diag_text.insert(tk.END, msg + '\n', tag)
        self.diag_text.see(tk.END)

    # ── ANÁLISIS PREVIO ──────────────────────
    def _analizar_preview(self):
        origen = self.origen_var.get().strip()
        if not origen or not os.path.isdir(origen):
            messagebox.showerror("Error", "Selecciona una carpeta de origen válida.")
            return

        archivos = self._listar_archivos(origen)
        if not archivos:
            messagebox.showwarning("Vacío", "No se encontraron archivos .xlsx/.xls/.csv")
            return

        self.diag_text.delete(1.0, tk.END)
        self._diag(f"{'─' * 70}", 'titulo')
        self._diag(f"🔍  ANÁLISIS DE ESTRUCTURA — {len(archivos)} archivos", 'titulo')
        self._diag(f"{'─' * 70}\n", 'titulo')
        self.nb.select(self.tab_diag)

        def _run():
            resultados = []
            estructura_ref = None

            for ruta in archivos:
                self._log_t(f"Analizando {os.path.basename(ruta)}…", 'info')
                info = analizar_archivo(ruta)
                resultados.append(info)

                # Tomar estructura del primer archivo válido como referencia
                if info['ok'] and estructura_ref is None:
                    estructura_ref = info['estructura']

                # Validar contra referencia global
                if info['ok'] and estructura_ref:
                    comp = columnas_coinciden(estructura_ref, info['estructura'])
                    if not comp['coincide']:
                        info['error_vs_ref'] = comp
                    else:
                        info['error_vs_ref'] = None
                else:
                    info['error_vs_ref'] = None

                self.after(0, lambda i=info, ref=estructura_ref:
                           self._mostrar_diag_archivo(i, ref))

            self.analisis_cache = resultados
            self.estructura_ref = estructura_ref

            # Mostrar estructura de referencia en tab config
            if estructura_ref:
                txt = f"✅ {len(estructura_ref)} columnas: {', '.join(estructura_ref)}"
                self.after(0, lambda: self.lbl_ref.config(text=txt, fg=COLOR_OK))
            else:
                self.after(0, lambda: self.lbl_ref.config(
                    text="❌ No se encontró ningún archivo válido", fg=COLOR_ERR))

            # Resumen
            ok_count = sum(1 for r in resultados if r['ok'] and not r.get('error_vs_ref'))
            err_count = len(resultados) - ok_count
            self._log_t(f"✅ Análisis completado: {ok_count} compatibles, {err_count} con problemas", 'success')

        threading.Thread(target=_run, daemon=True).start()

    def _mostrar_diag_archivo(self, info, estructura_ref):
        nombre = info['nombre']
        if info['ok']:
            self._diag(f"📄  {nombre}", 'titulo')
            self._diag(f"    Hojas detectadas: {info['num_hojas']}  |  "
                       f"Hojas válidas: {len(info['hojas_ok'])}  |  "
                       f"Filas totales: {info['filas_total']:,}", 'ok')

            if info['hojas_ok']:
                self._diag(f"    Hojas OK: {', '.join(info['hojas_ok'])}", 'ok')

            # Validar vs referencia global
            err_ref = info.get('error_vs_ref')
            if err_ref:
                self._diag(f"    ❌ ESTRUCTURA DIFERENTE vs referencia:", 'err')
                if err_ref['faltantes']:
                    self._diag(f"       Faltan: {err_ref['faltantes']}", 'err')
                if err_ref['sobrantes']:
                    self._diag(f"       Sobran: {err_ref['sobrantes']}", 'err')
            else:
                self._diag(f"    ✅ Estructura compatible con la referencia", 'ok')

            # Columnas
            self._diag(f"    Columnas ({len(info['estructura'])}):", '')
            for col in info['estructura']:
                self._diag(f"        {col}", '')

            for adv in info['advertencias']:
                self._diag(f"    ⚠️  {adv}", 'warn')
        else:
            self._diag(f"❌  {nombre}: {info['error']}", 'err')

        self._diag('')

    # ── INICIO CONSOLIDACIÓN ─────────────────
    def _iniciar(self):
        origen = self.origen_var.get().strip()
        if not origen or not os.path.isdir(origen):
            messagebox.showerror("Error", "Selecciona una carpeta de origen válida.")
            return

        salida = self.salida_var.get().strip() or origen
        archivos = self._listar_archivos(origen)
        if not archivos:
            messagebox.showwarning("Sin archivos", "No hay archivos para procesar.")
            return

        self.log_text.delete(1.0, tk.END)
        self._log(f"🚀 Iniciando — {len(archivos)} archivos", 'info')
        self._log(f"📁 Origen: {origen}", 'info')
        self._log(f"📁 Salida: {salida}", 'info')

        self.errores = []
        self.df_resultado = None
        self.lbl_registros.config(text="Registros: 0")
        self.lbl_errcount.config(text="Errores: 0")
        self.progress['value'] = 0

        self.control = ControlProceso()
        self._update_buttons(procesando=True)
        self.lbl_estado.config(text="⏳ Procesando…")
        self.nb.select(self.tab_log)

        self.hilo = threading.Thread(target=self._ejecutar,
                                     args=(archivos, salida), daemon=True)
        self.hilo.start()

    def _ejecutar(self, archivos, salida):
        try:
            # 1. Determinar estructura de referencia (primer archivo válido)
            self._log_t("🔍 Determinando estructura de referencia…", 'info')
            estructura_ref = None

            for ruta in archivos:
                info = analizar_archivo(ruta)
                if info['ok'] and info['estructura']:
                    estructura_ref = info['estructura']
                    self._log_t(f"📐 Referencia: {info['nombre']} → {len(estructura_ref)} columnas", 'info')
                    self._log_t(f"   Columnas: {', '.join(estructura_ref)}", 'info')
                    break

            if not estructura_ref:
                self._log_t("❌ No se encontró ningún archivo válido para usar como referencia.", 'error')
                return

            self.estructura_ref = estructura_ref

            # 2. Consolidar
            df, errores = consolidar(
                archivos, estructura_ref, self.control,
                self._log_t, self._progreso_t, self._registros_t, self._arch_t,
                self.elim_dup_var.get()
            )

            # Quitar trazabilidad si no se quiere
            if df is not None and not self.incl_origen_var.get():
                for col in ['__archivo__', '__hoja__']:
                    if col in df.columns:
                        df.drop(columns=[col], inplace=True)

            self.errores = errores
            self.after(0, self._refrescar_errores)
            self.after(0, lambda: self.lbl_errcount.config(
                text=f"Errores: {len(errores)}",
                fg=COLOR_ERR if any(e['tipo'] == 'Error lectura' for e in errores) else COLOR_WARN))

            if df is not None and not df.empty:
                self.df_resultado = df
                self._log_t(f"✅ DataFrame listo: {len(df):,} filas × {len(df.columns)} cols", 'success')
                self.after(0, lambda: self._preview_y_guardar(df, salida))
            else:
                self._log_t("❌ No se generaron datos.", 'error')

        except Exception as e:
            self._log_t(f"🔥 Error grave: {e}", 'error')
            self._log_t(traceback.format_exc(), 'error')
        finally:
            self.after(0, self._finalizar)

    def _preview_y_guardar(self, df, salida):
        ok = self._mostrar_preview(df)
        if not ok:
            self._log("⏹️ Guardado cancelado.", 'warning')
            return

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        nombre_base = self.nombre_salida_var.get().strip() or 'consolidado'
        nombre_csv  = f"{nombre_base}_{ts}.csv"
        nombre_xlsx = f"{nombre_base}_{ts}.xlsx"

        # Guardar CSV
        ruta_csv = os.path.join(salida, nombre_csv)
        try:
            df.to_csv(ruta_csv, index=False, encoding='utf-8-sig')
            self._log(f"💾 CSV guardado: {ruta_csv}", 'success')
        except Exception as e:
            self._log(f"❌ Error guardando CSV: {e}", 'error')

        # Guardar XLSX
        if _EXCEL_OK:
            ruta_xlsx = os.path.join(salida, nombre_xlsx)
            try:
                df.to_excel(ruta_xlsx, index=False, engine='openpyxl')
                self._log(f"💾 XLSX guardado: {ruta_xlsx}", 'success')
            except Exception as e:
                self._log(f"❌ Error guardando XLSX: {e}", 'error')
        else:
            self._log("⚠️  openpyxl no instalado, no se puede generar XLSX", 'warning')

        # Reporte de errores
        if self.errores:
            ruta_rep = os.path.join(salida, f"reporte_errores_{ts}.csv")
            try:
                pd.DataFrame(self.errores).to_csv(ruta_rep, index=False, encoding='utf-8-sig')
                self._log(f"📋 Reporte de errores: {ruta_rep}", 'warning')
            except:
                pass

        self._log(f"🏁 Total filas consolidadas: {len(df):,}", 'success')
        messagebox.showinfo("Completado",
                            f"✅ Consolidación exitosa\n\n"
                            f"Filas: {len(df):,}\n"
                            f"Columnas: {len(df.columns)}\n"
                            f"CSV: {nombre_csv}\n"
                            f"XLSX: {nombre_xlsx if _EXCEL_OK else 'N/A'}")

    # ── PREVIEW ──────────────────────────────
    def _mostrar_preview(self, df) -> bool:
        win = tk.Toplevel(self)
        win.title("Vista previa del consolidado")
        win.geometry("980x620")
        win.configure(bg=COLOR_BG)
        win.grab_set()

        self._preview_ok = False

        tk.Label(win, text=f"📊  Vista previa — {len(df):,} filas × {len(df.columns)} columnas",
                 bg=COLOR_MID, fg='white', font=('Segoe UI', 11, 'bold')).pack(fill='x')

        nb2 = ttk.Notebook(win)
        nb2.pack(fill='both', expand=True, padx=5, pady=5)

        # Tab datos
        tab_datos = ttk.Frame(nb2)
        nb2.add(tab_datos, text='📄 Datos (primeras 100 filas)')

        frame_tree = ttk.Frame(tab_datos)
        frame_tree.pack(fill='both', expand=True, padx=5, pady=5)

        tree = ttk.Treeview(frame_tree, columns=list(df.columns), show='headings')
        vsb = ttk.Scrollbar(frame_tree, orient='vertical',   command=tree.yview)
        hsb = ttk.Scrollbar(frame_tree, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        frame_tree.rowconfigure(0, weight=1)
        frame_tree.columnconfigure(0, weight=1)

        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, minwidth=60, stretch=True)
        for _, row in df.head(100).iterrows():
            vals = [str(v) if pd.notna(v) else '' for v in row]
            tree.insert('', 'end', values=vals)

        # Tab resumen
        tab_res = ttk.Frame(nb2)
        nb2.add(tab_res, text='📈 Resumen')

        res_text = ScrolledText(tab_res, font=('Consolas', 9), wrap='none', bg=COLOR_PANEL)
        res_text.pack(fill='both', expand=True, padx=5, pady=5)

        lineas = [
            f"{'Columna':<40} {'Tipo':<12} {'Nulos':>8} {'%Nulo':>7} {'Únicos':>8}",
            "─" * 80
        ]
        for col in df.columns:
            nulos = int(df[col].isna().sum())
            pct   = nulos / max(len(df), 1) * 100
            uniq  = df[col].nunique()
            tipo  = str(df[col].dtype)
            alerta = " ⚠️" if pct > 50 else ""
            lineas.append(f"{col:<40} {tipo:<12} {nulos:>8} {pct:>6.1f}% {uniq:>8}{alerta}")

        # Estadísticas por archivo
        if '__archivo__' in df.columns:
            lineas.append(f"\n{'─' * 80}")
            lineas.append("📁 Filas por archivo:")
            for arch, cnt in df['__archivo__'].value_counts().items():
                lineas.append(f"    {arch}: {cnt:,} filas")

        if '__hoja__' in df.columns:
            lineas.append(f"\n📄 Filas por hoja:")
            for hoja, cnt in df['__hoja__'].value_counts().items():
                lineas.append(f"    {hoja}: {cnt:,} filas")

        res_text.insert(tk.END, '\n'.join(lineas))
        res_text.config(state='disabled')

        # Botones
        btn_fr = ttk.Frame(win)
        btn_fr.pack(pady=8)

        def _aceptar():
            self._preview_ok = True
            win.destroy()

        def _cancelar():
            self._preview_ok = False
            win.destroy()

        ttk.Button(btn_fr, text="✅ Guardar (CSV + XLSX)", command=_aceptar,
                   style='Accent.TButton').pack(side='left', padx=10)
        ttk.Button(btn_fr, text="❌ Cancelar", command=_cancelar).pack(side='left', padx=10)

        win.protocol("WM_DELETE_WINDOW", _cancelar)
        self.wait_window(win)
        return self._preview_ok

    # ── ERRORES TAB ──────────────────────────
    def _refrescar_errores(self):
        tipos = ['Todos'] + sorted({e['tipo'] for e in self.errores})
        self.combo_filtro['values'] = tipos
        filtro = self.filtro_err_var.get()
        if filtro not in tipos:
            filtro = 'Todos'

        for item in self.tree_err.get_children():
            self.tree_err.delete(item)

        for e in self.errores:
            if filtro != 'Todos' and e['tipo'] != filtro:
                continue
            detalle = str(e.get('detalle', ''))[:300]
            self.tree_err.insert('', 'end', values=(
                e.get('archivo', ''), e.get('hoja', ''),
                e.get('tipo', ''), detalle))

        self.nb.tab(self.tab_errores, text=f"⚠️  Errores ({len(self.errores)})")

    # ── PROGRESO ─────────────────────────────
    def _progreso_t(self, actual, total):
        self.after(0, lambda: self._set_progreso(actual, total))

    def _set_progreso(self, actual, total):
        self.progress['value'] = actual / max(total, 1) * 100
        self.lbl_estado.config(text=f"⏳ {actual}/{total}")

    def _registros_t(self, n):
        self.after(0, lambda: self.lbl_registros.config(text=f"Registros: {n:,}"))

    def _arch_t(self, actual, total):
        self.after(0, lambda: self.lbl_proc.config(text=f"Procesados: {actual}/{total}"))

    # ── CONTROL ──────────────────────────────
    def _pausar(self):
        if self.control:
            self.control.pausar()
            self._log("⏸️ Pausado.", 'warning')
            self.btn_pausar.config(state='disabled')
            self.btn_reanudar.config(state='normal')
            self.lbl_estado.config(text="⏸️ Pausado")

    def _reanudar(self):
        if self.control:
            self.control.reanudar()
            self._log("▶️ Reanudado.", 'info')
            self.btn_pausar.config(state='normal')
            self.btn_reanudar.config(state='disabled')
            self.lbl_estado.config(text="⏳ Procesando…")

    def _detener(self):
        if self.control:
            self.control.detener()
            self._log("⏹️ Deteniendo…", 'error')
            self.btn_pausar.config(state='disabled')
            self.btn_reanudar.config(state='disabled')
            self.btn_detener.config(state='disabled')

    def _update_buttons(self, procesando):
        if procesando:
            self.btn_iniciar.config(state='disabled')
            self.btn_analizar.config(state='disabled')
            self.btn_pausar.config(state='normal')
            self.btn_detener.config(state='normal')
            self.btn_reanudar.config(state='disabled')
        else:
            self.btn_iniciar.config(state='normal')
            self.btn_analizar.config(state='normal')
            self.btn_pausar.config(state='disabled')
            self.btn_reanudar.config(state='disabled')
            self.btn_detener.config(state='disabled')

    def _finalizar(self):
        self._update_buttons(procesando=False)
        self.lbl_estado.config(text="✅ Finalizado")
        self.progress['value'] = 0
        self.control = None
        self.hilo = None

    # ── UTILIDADES ───────────────────────────
    def _listar_archivos(self, ruta):
        result = []
        for f in sorted(os.listdir(ruta)):
            if (f.lower().endswith(('.xlsx', '.xls', '.csv', '.xlsm'))
                    and not f.startswith('~$')):
                result.append(os.path.join(ruta, f))
        return result

    def _exportar_log(self):
        ruta = filedialog.asksaveasfilename(defaultextension='.txt',
                                            filetypes=[("Texto", "*.txt")])
        if ruta:
            with open(ruta, 'w', encoding='utf-8') as f:
                f.write(self.log_text.get(1.0, tk.END))
            messagebox.showinfo("Log exportado", ruta)

    def _exportar_errores(self):
        if not self.errores:
            messagebox.showinfo("Sin datos", "No hay errores registrados.")
            return
        ruta = filedialog.asksaveasfilename(defaultextension='.csv',
                                            filetypes=[("CSV", "*.csv")])
        if ruta:
            pd.DataFrame(self.errores).to_csv(ruta, index=False, encoding='utf-8-sig')
            messagebox.showinfo("Exportado", ruta)


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────
if __name__ == '__main__':
    app = App()
    app.mainloop()