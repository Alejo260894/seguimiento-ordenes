"""
Seguimiento de órdenes — aplicación de escritorio.

- "Cargar CSV de SAP": limpia el reporte, genera el control Excel del día y actualiza el seguimiento.
- Pestaña "OC sin factura": OCs en proceso sin factura; estado, responsable, fecha compromiso y comentarios.
- Pestaña "Facturas": facturas del último reporte (Pendiente de pago / Pagado) para revisar hasta su pago.
- Los comentarios son por OC, se guardan en seguimiento.db (junto a la aplicación) y se arrastran a cada
  control nuevo; las facturas nuevas y los pagos detectados quedan registrados en el historial de la OC.
"""
import json
import os
import queue
import shutil
import sys
import threading
import traceback
from datetime import datetime, date

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import procesar_sap as ps
import seguimiento as sgm
import requerimientos as rq
import app_requerimientos as ui_req

BASE = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE, "seguimiento.db")
CFG = os.path.join(BASE, "config_seguimiento.json")
CARGAS = os.path.join(BASE, "Cargas")

SITUACIONES = ["Activas", sgm.SIT_SIN_FACT, sgm.SIT_FACTURADA, sgm.SIT_CERRADA_SAP, sgm.SIT_CANCELADA_SAP,
               sgm.SIT_FUERA, "Todas"]
COLS = [("oc", "OC", 85, "e"), ("usuario", "Comprador", 120, "w"), ("proveedor", "Proveedor", 290, "w"),
        ("mon", "Mon", 45, "center"), ("total", "Total sin IGV", 100, "e"), ("fecha_oc", "Fecha OC", 80, "center"),
        ("dias", "Días", 50, "e"), ("estado_oc", "Estado OC", 80, "w"), ("situacion", "Situación", 110, "w"),
        ("estado_seg", "Seguimiento", 150, "w"), ("responsable", "Responsable", 110, "w"),
        ("compromiso", "Compromiso", 85, "center"), ("ult_com", "Último comentario", 330, "w")]
COLS_FAC = [("oc", "OC", 85, "e"), ("usuario", "Comprador", 115, "w"), ("proveedor", "Proveedor", 250, "w"),
            ("n_factura", "N° Factura", 110, "w"), ("tipo", "Tipo", 40, "center"), ("estado_factura", "Estado factura", 115, "w"),
            ("mon", "Mon", 45, "center"), ("ft_sin_imp", "FT sin imp.", 95, "e"), ("imp_pendiente", "Imp. pendiente", 100, "e"),
            ("emision", "Emisión", 80, "center"), ("vencimiento", "Vencimiento", 85, "center"),
            ("dias_venc", "Días p/ vencer", 85, "e"), ("fecha_pago", "Fecha pago", 80, "center"),
            ("estado_oc", "Estado OC", 80, "w"), ("ult_com", "Último comentario (OC)", 300, "w")]
ESTADOS_FACT_FILTRO = ["Pendiente de pago", "Pagado", "Todos", "(sin estado)"]
MONEDAS = ["Todas", "SOL", "USD", "CAD", "AUD", "GBP", "EUR", "ZAR"]


def fmt_fecha(iso):
    return f"{iso[8:10]}/{iso[5:7]}/{iso[:4]}" if iso else ""


def leer_cfg():
    try:
        with open(CFG, encoding="utf-8-sig") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {}


def guardar_cfg(cfg):
    with open(CFG, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)


def registrar_error(texto):
    """Guarda el detalle técnico de un error en errores.log (junto a la aplicación)."""
    try:
        with open(os.path.join(BASE, "errores.log"), "a", encoding="utf-8") as fh:
            fh.write(f"\n===== {datetime.now():%d/%m/%Y %H:%M:%S}\n{texto}\n")
    except OSError:
        pass


def fecha_de(meta, ruta):
    """Fecha y hora del reporte; si el archivo no la trae, la de modificación del archivo."""
    return ps.fecha_reporte(meta) or datetime.fromtimestamp(os.path.getmtime(ruta))


def generar_control(df, meta, log, con, fr):
    """Escribe 'Control de ordenes DD_MM.xlsx' con la hoja de seguimiento. Devuelve la ruta."""
    salida = os.path.join(BASE, f"Control de ordenes {fr:%d_%m}.xlsx")
    tmp = os.path.join(BASE, f"~tmp_control_{os.getpid()}.xlsx")
    ps.escribir_libro(df, meta, log, tmp, seguimiento=sgm.tabla_para_excel(con),
                      comentarios_oc=sgm.ultimos_comentarios(con))
    try:
        os.replace(tmp, salida)
    except PermissionError:          # el control está abierto en Excel
        salida = os.path.join(BASE, f"Control de ordenes {fr:%d_%m} ({datetime.now():%H%M%S}).xlsx")
        os.replace(tmp, salida)
    return salida


class ReporteIncompleto(Exception):
    pass


def preparar_carga(ruta):
    """Paso 1 (no toca la base): limpia el reporte y revisa si parece completo.
    Devuelve (df, log, meta, fecha_reporte, aviso); aviso != '' si faltan muchas órdenes respecto del anterior."""
    df, log, meta, _ = ps.procesar(ruta)
    fr = fecha_de(meta, ruta)
    con = sgm.conectar(DB)
    try:
        rq.preparar(con)
        previas = {r[0] for r in con.execute("SELECT numero_oc FROM sap_oc")}
        ult = sgm.ultima_carga(con)
    finally:
        con.close()
    aviso = ""
    if previas and (ult is None or f"{fr:%Y-%m-%d %H:%M:%S}" >= ult["fecha_reporte"]):
        faltan = previas - set(df["Numero OC"])
        if len(faltan) > max(10, 0.05 * len(previas)):
            aviso = (f"El reporte trae {df['Numero OC'].nunique()} órdenes y el anterior tenía {len(previas)}: faltan "
                     f"{len(faltan)} órdenes que sí estaban antes.\n\nSi es un reporte filtrado o incompleto, esas "
                     f"órdenes se cerrarían como «Fuera del reporte» y sus facturas se marcarían como no vigentes.")
    return df, log, meta, fr, aviso


def archivar(ruta, fr):
    """Copia el reporte a Cargas\\ (historial). Si no se puede, devuelve la ruta original."""
    archivo = os.path.join(CARGAS, f"Control de ordenes {fr:%Y-%m-%d %H%M}{os.path.splitext(ruta)[1]}")
    try:
        os.makedirs(CARGAS, exist_ok=True)
        if not (os.path.exists(archivo) and os.path.samefile(ruta, archivo)):
            shutil.copy2(ruta, archivo)
        return archivo
    except OSError:
        return ruta


def aplicar_carga(ruta, df, log, meta, fr):
    """Paso 2: archiva el reporte, sincroniza el seguimiento y genera el control.
    Devuelve (resultado, ruta_control o None si el reporte era anterior a la última carga, log)."""
    archivo = archivar(ruta, fr)
    con = sgm.conectar(DB)
    salida = None
    try:
        res = sgm.sincronizar(con, df, fr, archivo)
        if not res["omitida"]:
            rq.preparar(con)
            rq.guardar_snapshot_sap(con, df)
            res["req_completados"] = rq.completar_desde_sap(con)
            salida = generar_control(df, meta, log, con, fr)
            con.execute("UPDATE cargas SET salida=? WHERE id=(SELECT MAX(id) FROM cargas)", (salida,))
            con.commit()
    finally:
        con.close()
    if salida:
        cfg = leer_cfg()
        cfg["ultimo_control"] = salida
        guardar_cfg(cfg)
    return res, salida, log


def cargar_reporte(ruta, forzar=False):
    """Carga completa sin preguntas (modo --cargar). Un reporte que parece incompleto no se aplica."""
    df, log, meta, fr, aviso = preparar_carga(ruta)
    if aviso and not forzar:
        raise ReporteIncompleto(aviso)
    return aplicar_carga(ruta, df, log, meta, fr)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Seguimiento de requerimientos, órdenes y facturas")
        self.geometry("1500x880")
        self.minsize(1100, 650)
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Treeview", rowheight=22, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("Titulo.TLabel", font=("Segoe UI", 12, "bold"))
        style.configure("Accent.TButton", font=("Segoe UI", 9, "bold"))

        self.cfg = leer_cfg()
        self.base = BASE
        self.guardar_cfg = lambda: guardar_cfg(self.cfg)
        self.con = sgm.conectar(DB)
        rq.preparar(self.con)
        self.filas = []
        self.facturas = []
        self.oc_sel = None
        self.fac_sel = None
        self.modo = "oc"                     # "oc" (pestaña OC sin factura) o "fac" (pestaña Facturas)
        self.orden = ("dias", True)
        self.orden_fac = ("dias_venc", False)
        self.borradores = {}                  # comentario sin enviar de cada OC
        self._com_oc = None                   # OC a la que pertenece el texto del cuadro «Nuevo comentario»
        self.sash = None                      # altura del panel de detalle al ocultarlo
        self.ultima_carga_id = None
        self._ui()
        self.bind_all("<F5>", lambda e: self.actualizar_vista())
        self.bind("<Activate>", lambda e: self.after_idle(self._revisar_cambios_externos))
        self.after(100, self._pedir_autor)
        self.refrescar()

    def report_callback_exception(self, exc, val, tb):
        """Errores en botones: en el .exe no hay consola, así que se muestran y se guardan en errores.log."""
        try:
            self.con.rollback()
        except Exception:  # noqa: BLE001
            pass
        registrar_error("".join(traceback.format_exception(exc, val, tb)))
        messagebox.showerror("Error", f"No se pudo completar la acción:\n{val}\n\nEl detalle técnico quedó en "
                                      f"errores.log, en la carpeta de la aplicación.", parent=self)

    def actualizar_vista(self):
        self._guardar_pendientes()
        self.refrescar()

    def _revisar_cambios_externos(self):
        """Si una carga programada (--cargar) actualizó la base mientras la app estaba abierta, se refresca."""
        u = sgm.ultima_carga(self.con)
        if u is not None and u["id"] != self.ultima_carga_id:
            self.actualizar_vista()

    # ------------------------------------------------------------------ UI
    def _ui(self):
        barra = ttk.Frame(self, padding=(8, 8, 8, 4))
        barra.pack(fill="x")
        ttk.Button(barra, text="📥  Cargar CSV de SAP…", style="Accent.TButton", command=self.cargar_csv).pack(side="left")
        ttk.Button(barra, text="🔄  Actualizar Excel con comentarios", command=self.regenerar_excel).pack(side="left", padx=6)
        ttk.Button(barra, text="📊  Abrir último control", command=self.abrir_control).pack(side="left")
        ttk.Button(barra, text="📁  Abrir carpeta", command=lambda: os.startfile(BASE)).pack(side="left", padx=6)
        self.lbl_autor = ttk.Label(barra, text="")
        ttk.Button(barra, text="Cambiar", command=lambda: self._pedir_autor(True)).pack(side="right")
        self.lbl_autor.pack(side="right", padx=6)

        self.lbl_estado = ttk.Label(self, padding=(10, 0), foreground="#1F3864")
        self.lbl_estado.pack(fill="x")

        paned = self.paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=8, pady=(4, 8))
        self.nb = ttk.Notebook(paned)
        paned.add(self.nb, weight=3)
        tab_req = ttk.Frame(self.nb)
        tab_oc = ttk.Frame(self.nb)
        tab_fac = ttk.Frame(self.nb)
        tab_ind = ttk.Frame(self.nb)
        self.nb.add(tab_req, text="  Requerimientos  ")
        self.nb.add(tab_oc, text="  OC sin factura (en proceso)  ")
        self.nb.add(tab_fac, text="  Facturas (pendientes de pago / pagadas)  ")
        self.nb.add(tab_ind, text="  Indicadores  ")
        self.tabs = {str(tab_req): "req", str(tab_oc): "oc", str(tab_fac): "fac", str(tab_ind): "ind"}
        ui_req.construir(self, tab_req, tab_ind)
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self.cambiar_pestana())

        filt = ttk.Frame(tab_oc, padding=(4, 6))
        filt.pack(fill="x")
        ttk.Label(filt, text="Buscar:").pack(side="left")
        self.v_buscar = tk.StringVar()
        e = ttk.Entry(filt, textvariable=self.v_buscar, width=28)
        e.pack(side="left", padx=(4, 12))
        self.v_buscar.trace_add("write", lambda *_: self.pintar())
        self.v_comprador = self._combo(filt, "Comprador:", ["Todos"], 18)
        self.v_estado_oc = self._combo(filt, "Estado OC:", ["Todos", "En proceso", "Cerrado", "Cancelado"], 13)
        self.v_situacion = self._combo(filt, "Situación:", SITUACIONES, 18)
        self.v_seg = self._combo(filt, "Seguimiento:", ["Todos"] + sgm.ESTADOS_SEG, 22)

        filt2 = ttk.Frame(tab_oc, padding=(4, 0, 4, 4))
        filt2.pack(fill="x")
        self.v_sincom = tk.BooleanVar()
        ttk.Checkbutton(filt2, text="Solo sin comentarios", variable=self.v_sincom, command=self.pintar).pack(side="left")
        self.v_venc = tk.BooleanVar()
        ttk.Checkbutton(filt2, text="Solo compromiso vencido", variable=self.v_venc, command=self.pintar).pack(side="left", padx=10)
        ttk.Button(filt2, text="Limpiar filtros", command=self.limpiar_filtros).pack(side="left", padx=(0, 16))
        self.lbl_resumen = ttk.Label(filt2, foreground="#404040", justify="left")
        self.lbl_resumen.pack(side="left")

        self.tree = self._tabla(tab_oc, COLS, self.ordenar, ("proveedor", "ult_com"))
        self.tree.tag_configure("vencido", foreground="#C00000")
        self.tree.tag_configure("sincom", background="#FFF8DC")
        self.tree.tag_configure("facturada", background="#E2EFDA")
        self.tree.tag_configure("cerrada", foreground="#808080")
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.seleccionar())

        # ---- pestaña Facturas
        ff = ttk.Frame(tab_fac, padding=(4, 6))
        ff.pack(fill="x")
        ttk.Label(ff, text="Buscar:").pack(side="left")
        self.vf_buscar = tk.StringVar()
        ttk.Entry(ff, textvariable=self.vf_buscar, width=24).pack(side="left", padx=(4, 12))
        self.vf_buscar.trace_add("write", lambda *_: self.pintar_fac())
        self.vf_estado = self._combo(ff, "Estado factura:", ESTADOS_FACT_FILTRO, 17, self.pintar_fac)
        self.vf_comprador = self._combo(ff, "Comprador:", ["Todos"], 18, self.pintar_fac)
        self.vf_estado_oc = self._combo(ff, "Estado OC:", ["Todos", "En proceso", "Cerrado", "Cancelado"], 11, self.pintar_fac)
        self.vf_mon = self._combo(ff, "Moneda:", MONEDAS, 6, self.pintar_fac)
        ff2 = ttk.Frame(tab_fac, padding=(4, 0, 4, 4))
        ff2.pack(fill="x")
        self.vf_vencidas = tk.BooleanVar()
        ttk.Checkbutton(ff2, text="Solo vencidas sin pagar", variable=self.vf_vencidas, command=self.pintar_fac).pack(side="left")
        self.vf_concom = tk.BooleanVar()
        ttk.Checkbutton(ff2, text="Solo con comentarios", variable=self.vf_concom, command=self.pintar_fac).pack(side="left", padx=10)
        self.vf_fuera = tk.BooleanVar()
        ttk.Checkbutton(ff2, text="Incluir facturas fuera del reporte", variable=self.vf_fuera, command=self.pintar_fac).pack(side="left")
        ttk.Button(ff2, text="Limpiar filtros", command=self.limpiar_filtros_fac).pack(side="left", padx=(10, 16))
        self.lbl_resumen_fac = ttk.Label(ff2, foreground="#404040", justify="left")
        self.lbl_resumen_fac.pack(side="left")
        self.tree_fac = self._tabla(tab_fac, COLS_FAC, self.ordenar_fac, ("proveedor", "ult_com"))
        self.tree_fac.tag_configure("vencida", foreground="#C00000")
        self.tree_fac.tag_configure("pagada", background="#E2EFDA")
        self.tree_fac.tag_configure("fuera", foreground="#808080")
        self.tree_fac.bind("<<TreeviewSelect>>", lambda e: self.seleccionar_fac())

        det = self.det = ttk.Frame(paned, padding=(0, 6, 0, 0))
        paned.add(det, weight=2)
        det.columnconfigure(1, weight=3)
        det.columnconfigure(2, weight=2)
        det.rowconfigure(0, weight=1)

        izq = ttk.LabelFrame(det, text="Orden seleccionada", padding=8)
        izq.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.lbl_oc = ttk.Label(izq, text="Selecciona una orden de la lista", style="Titulo.TLabel", wraplength=330)
        self.lbl_oc.grid(row=0, column=0, columnspan=2, sticky="w")
        self.lbl_info = ttk.Label(izq, text="", wraplength=330, justify="left")
        self.lbl_info.grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 8))
        ttk.Label(izq, text="Estado seguimiento:").grid(row=2, column=0, sticky="w")
        self.v_det_seg = tk.StringVar()
        self.cb_det_seg = ttk.Combobox(izq, textvariable=self.v_det_seg, values=sgm.ESTADOS_SEG, state="readonly", width=26)
        self.cb_det_seg.grid(row=2, column=1, sticky="w", pady=2)
        ttk.Label(izq, text="Responsable:").grid(row=3, column=0, sticky="w")
        self.v_det_resp = tk.StringVar()
        self.ent_resp = ttk.Entry(izq, textvariable=self.v_det_resp, width=28)
        self.ent_resp.grid(row=3, column=1, sticky="w", pady=2)
        ttk.Label(izq, text="Fecha compromiso:").grid(row=4, column=0, sticky="w")
        self.v_det_comp = tk.StringVar()
        self.ent_comp = ttk.Entry(izq, textvariable=self.v_det_comp, width=12)
        self.ent_comp.grid(row=4, column=1, sticky="w", pady=2)
        ttk.Label(izq, text="dd/mm/aaaa", foreground="#808080").grid(row=5, column=1, sticky="w")
        self.btn_guardar = ttk.Button(izq, text="💾  Guardar cambios", command=self.guardar_orden, state="disabled")
        self.btn_guardar.grid(row=6, column=0, columnspan=2, sticky="w", pady=(8, 0))

        mid = ttk.LabelFrame(det, text="Historial (trazabilidad)", padding=6)
        mid.grid(row=0, column=1, sticky="nsew", padx=(0, 6))
        mid.rowconfigure(0, weight=1)
        mid.columnconfigure(0, weight=1)
        self.txt_hist = tk.Text(mid, wrap="word", font=("Segoe UI", 9), state="disabled", relief="flat", background="#FAFAFA")
        hs = ttk.Scrollbar(mid, command=self.txt_hist.yview)
        self.txt_hist.configure(yscrollcommand=hs.set)
        self.txt_hist.grid(row=0, column=0, sticky="nsew")
        hs.grid(row=0, column=1, sticky="ns")
        self.txt_hist.tag_configure("cab", font=("Segoe UI", 8, "bold"), foreground="#1F3864")
        self.txt_hist.tag_configure("cab_sis", font=("Segoe UI", 8, "bold"), foreground="#808080")
        self.txt_hist.tag_configure("sis", foreground="#808080", font=("Segoe UI", 9, "italic"))
        self.txt_hist.tag_configure("cambio", foreground="#7F6000", font=("Segoe UI", 9, "italic"))

        der = self.lf_com = ttk.LabelFrame(det, text="Nuevo comentario", padding=6)
        der.grid(row=0, column=2, sticky="nsew")
        der.rowconfigure(0, weight=1)
        der.columnconfigure(0, weight=1)
        self.txt_com = tk.Text(der, wrap="word", font=("Segoe UI", 10), height=6)
        self.txt_com.grid(row=0, column=0, sticky="nsew")
        self.txt_com.bind("<Control-Return>", lambda e: (self.agregar_comentario(), "break")[1])
        self.btn_com = ttk.Button(der, text="➕  Agregar comentario  (Ctrl+Enter)", command=self.agregar_comentario, state="disabled")
        self.btn_com.grid(row=1, column=0, sticky="e", pady=(6, 0))

    def _tabla(self, parent, cols, ordenar, estirar):
        marco = ttk.Frame(parent)
        marco.pack(fill="both", expand=True)
        tree = ttk.Treeview(marco, columns=[c[0] for c in cols], show="headings", selectmode="browse")
        for key, txt, w, anc in cols:
            tree.heading(key, text=txt, command=lambda k=key: ordenar(k))
            tree.column(key, width=w, anchor=anc, stretch=key in estirar)
        sy = ttk.Scrollbar(marco, orient="vertical", command=tree.yview)
        sx = ttk.Scrollbar(marco, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        marco.rowconfigure(0, weight=1)
        marco.columnconfigure(0, weight=1)
        return tree

    def _combo(self, parent, etiqueta, valores, ancho, al_cambiar=None):
        ttk.Label(parent, text=etiqueta).pack(side="left")
        v = tk.StringVar(value=valores[0])
        cb = ttk.Combobox(parent, textvariable=v, values=valores, state="readonly", width=ancho)
        cb.pack(side="left", padx=(4, 12))
        cb.bind("<<ComboboxSelected>>", lambda e: (al_cambiar or self.pintar)())
        v.widget = cb
        return v

    def _pedir_autor(self, forzar=False):
        if forzar or not self.cfg.get("autor"):
            try:
                defecto = self.cfg.get("autor") or os.getlogin()
            except OSError:
                defecto = ""
            nombre = simpledialog.askstring("Tu nombre", "Nombre con el que se firmarán tus comentarios:",
                                            initialvalue=defecto, parent=self)
            if nombre and nombre.strip():
                self.cfg["autor"] = nombre.strip()
                guardar_cfg(self.cfg)
            elif not self.cfg.get("autor") and defecto:
                self.cfg["autor"] = defecto           # si se cancela, se firma con el usuario de Windows
                guardar_cfg(self.cfg)
        self.lbl_autor.config(text=f"Comentarios como: {self.cfg.get('autor', '(sin nombre)')}")

    # ------------------------------------------------------------- datos
    def refrescar(self):
        self.filas = sgm.listar(self.con)
        compradores = sorted({f["usuario"] for f in self.filas if f["usuario"]})
        self.v_comprador.widget.config(values=["Todos"] + compradores)
        u = sgm.ultima_carga(self.con)
        self.ultima_carga_id = u["id"] if u else None
        if u:
            self.lbl_estado.config(text=f"Última carga: reporte SAP del {fmt_fecha(u['fecha_reporte'])} {u['fecha_reporte'][11:16]}"
                                        f" · cargado el {fmt_fecha(u['fecha_carga'])} {u['fecha_carga'][11:16]}"
                                        f" · {u['nuevas']} nuevas, {u['facturadas']} facturadas, {u['fuera']} fuera del reporte en esa carga")
        else:
            self.lbl_estado.config(text="Aún no se ha cargado ningún reporte. Usa «Cargar CSV de SAP…».")
        self.facturas = sgm.listar_facturas(self.con)
        self.vf_comprador.widget.config(values=["Todos"] + sorted({f["usuario"] for f in self.facturas if f["usuario"]}))
        self.pintar()
        self.pintar_fac()
        self.refrescar_req()
        if self.oc_sel and self.modo in ("oc", "fac"):
            self.mostrar_detalle()        # el panel refleja lo que acaba de cambiar (p. ej. tras una carga)

    def refrescar_req(self):
        ui_req.refrescar(self)

    # -------------------------------------------------- pestaña Facturas
    def filtradas_fac(self):
        q = self.vf_buscar.get().strip().lower()
        est, comp, eoc, mon = self.vf_estado.get(), self.vf_comprador.get(), self.vf_estado_oc.get(), self.vf_mon.get()
        hoy = date.today().isoformat()
        out = []
        for f in self.facturas:
            if f["situacion"] != sgm.FACT_VIGENTE and not self.vf_fuera.get():
                continue
            if est == "(sin estado)" and f["estado_factura"]:
                continue
            if est not in ("Todos", "(sin estado)") and f["estado_factura"] != est:
                continue
            if comp != "Todos" and f["usuario"] != comp:
                continue
            if eoc != "Todos" and f["estado_oc"] != eoc:
                continue
            if mon != "Todas" and f["mon"] != mon:
                continue
            if self.vf_vencidas.get() and not (f["estado_factura"] != "Pagado" and f["vencimiento"] and f["vencimiento"] < hoy):
                continue
            if self.vf_concom.get() and not f["n_com"]:
                continue
            if q and q not in " ".join(str(f[k] or "") for k in ("numero_oc", "proveedor", "ruc", "n_factura", "usuario", "ult_com")).lower():
                continue
            out.append(f)
        return out

    @staticmethod
    def _dias_venc(f):
        if f["estado_factura"] == "Pagado" or not f["vencimiento"]:
            return None
        return -(sgm.dias_desde(f["vencimiento"]))

    def pintar_fac(self):
        filas = self.filtradas_fac()
        clave, desc = self.orden_fac
        num = {"oc": lambda f: f["numero_oc"], "ft_sin_imp": lambda f: f["ft_sin_imp"] or 0,
               "imp_pendiente": lambda f: f["imp_pendiente"] or 0,
               "dias_venc": lambda f: self._dias_venc(f) if self._dias_venc(f) is not None else 99999}
        keyf = num.get(clave, lambda f: str(f.get(clave) or "").lower())
        filas.sort(key=keyf, reverse=desc)
        t = self.tree_fac
        t.delete(*t.get_children())
        hoy = date.today().isoformat()
        for f in filas:
            tags = []
            if f["situacion"] != sgm.FACT_VIGENTE:
                tags.append("fuera")
            elif f["estado_factura"] == "Pagado":
                tags.append("pagada")
            elif f["vencimiento"] and f["vencimiento"] < hoy:
                tags.append("vencida")
            dv = self._dias_venc(f)
            t.insert("", "end", iid=f"{f['ruc']}|{f['n_factura']}", tags=tags, values=(
                f["numero_oc"], f["usuario"], f["proveedor"], f["n_factura"], f["tipo"], f["estado_factura"], f["mon"],
                f"{f['ft_sin_imp'] or 0:,.2f}", f"{f['imp_pendiente'] or 0:,.2f}", fmt_fecha(f["emision"]),
                fmt_fecha(f["vencimiento"]), "" if dv is None else dv, fmt_fecha(f["fecha_pago"]), f["estado_oc"],
                (f["ult_com"] or "").replace("\n", " ")))
        vig = [f for f in self.facturas if f["situacion"] == sgm.FACT_VIGENTE]
        pend = [f for f in vig if f["estado_factura"] == "Pendiente de pago"]
        venc = sum(1 for f in pend if f["vencimiento"] and f["vencimiento"] < hoy)
        tot = {}
        for f in filas:
            if f["estado_factura"] != "Pagado":
                tot[f["mon"]] = tot.get(f["mon"], 0) + (f["imp_pendiente"] or 0)
        txt_tot = " · ".join(f"{m} {v:,.2f}" for m, v in sorted(tot.items()) if v)
        self.lbl_resumen_fac.config(
            text=f"Mostrando {len(filas)}   |   Pendientes de pago: {len(pend)} (vencidas {venc}) · "
                 f"Pagadas: {sum(1 for f in vig if f['estado_factura'] == 'Pagado')}"
                 + (f"\nImporte pendiente de lo mostrado: {txt_tot}" if txt_tot else ""))
        if self.modo == "fac" and self.oc_sel:
            if self.fac_sel and t.exists(self.fac_sel):
                t.selection_set(self.fac_sel)
                t.see(self.fac_sel)
            else:
                self.limpiar_detalle()

    def ordenar_fac(self, clave):
        desc = not self.orden_fac[1] if self.orden_fac[0] == clave else False
        self.orden_fac = (clave, desc)
        self.pintar_fac()

    def limpiar_filtros_fac(self):
        self.vf_buscar.set("")
        self.vf_estado.set("Pendiente de pago")
        for v in (self.vf_comprador, self.vf_estado_oc):
            v.set("Todos")
        self.vf_mon.set("Todas")
        for v in (self.vf_vencidas, self.vf_concom, self.vf_fuera):
            v.set(False)
        self.pintar_fac()

    def cambiar_pestana(self):
        guardado = self._guardar_pendientes()
        self.modo = self.tabs[self.nb.select()]
        # El panel de detalle/comentarios es de las pestañas de OC y facturas
        visible = str(self.det) in self.paned.panes()
        if self.modo in ("oc", "fac") and not visible:
            self.paned.add(self.det, weight=2)
            if self.sash:
                self.after_idle(lambda: self.paned.sashpos(0, self.sash))
        elif self.modo in ("req", "ind") and visible:
            self.sash = self.paned.sashpos(0)
            self.paned.forget(self.det)
        self.limpiar_detalle()
        if guardado:
            self.refrescar()
        if self.modo == "oc" and self.tree.selection():
            self.seleccionar()
        elif self.modo == "fac" and self.tree_fac.selection():
            self.seleccionar_fac()

    def seleccionar_fac(self):
        sel = self.tree_fac.selection()
        if not sel or self.modo != "fac" or sel[0] == self.fac_sel:
            return
        self._guardar_borrador()
        self.fac_sel = sel[0]
        self.oc_sel = int(self.tree_fac.item(sel[0])["values"][0])
        self.mostrar_detalle()

    def filtradas(self):
        q = self.v_buscar.get().strip().lower()
        comp, eoc, sit, seg = self.v_comprador.get(), self.v_estado_oc.get(), self.v_situacion.get(), self.v_seg.get()
        hoy = date.today().isoformat()
        out = []
        for f in self.filas:
            if sit == "Activas" and f["situacion"] != sgm.SIT_ACTIVA:
                continue
            if sit not in ("Activas", "Todas") and f["situacion"] != sit:
                continue
            if comp != "Todos" and f["usuario"] != comp:
                continue
            if eoc != "Todos" and f["estado_oc"] != eoc:
                continue
            if seg != "Todos" and f["estado_seg"] != seg:
                continue
            if self.v_sincom.get() and f["n_com"]:
                continue
            if self.v_venc.get() and not (f["fecha_compromiso"] and f["fecha_compromiso"] < hoy):
                continue
            if q and q not in " ".join(str(f[k] or "") for k in ("numero_oc", "proveedor", "ruc", "usuario", "responsable", "ult_com")).lower():
                continue
            out.append(f)
        return out

    def pintar(self):
        filas = self.filtradas()
        clave, desc = self.orden
        keyf = {"oc": lambda f: f["numero_oc"], "dias": lambda f: sgm.dias_desde(f["fecha_oc"]) or 0,
                "total": lambda f: f["total"] or 0, "fecha_oc": lambda f: f["fecha_oc"] or "",
                "compromiso": lambda f: f["fecha_compromiso"] or "9999"}.get(clave, lambda f: str(f.get(clave) or "").lower())
        filas.sort(key=keyf, reverse=desc)
        self.tree.delete(*self.tree.get_children())
        hoy = date.today().isoformat()
        for f in filas:
            tags = []
            if f["situacion"] == sgm.SIT_FACTURADA:
                tags.append("facturada")
            elif f["situacion"] != sgm.SIT_ACTIVA:
                tags.append("cerrada")
            else:
                if not f["n_com"]:
                    tags.append("sincom")
                if f["fecha_compromiso"] and f["fecha_compromiso"] < hoy:
                    tags.append("vencido")
            self.tree.insert("", "end", iid=str(f["numero_oc"]), tags=tags, values=(
                f["numero_oc"], f["usuario"], f["proveedor"], f["mon"], f"{f['total'] or 0:,.2f}", fmt_fecha(f["fecha_oc"]),
                sgm.dias_desde(f["fecha_oc"]) or "", f["estado_oc"], f["situacion"], f["estado_seg"], f["responsable"] or "",
                fmt_fecha(f["fecha_compromiso"]), (f["ult_com"] or "").replace("\n", " ")))
        activas = [f for f in self.filas if f["situacion"] == sgm.SIT_ACTIVA]
        sin_c = sum(1 for f in activas if not f["n_com"])
        venc = sum(1 for f in activas if f["fecha_compromiso"] and f["fecha_compromiso"] < hoy)
        por_sit = {s: sum(1 for f in self.filas if f["situacion"] == s) for s in SITUACIONES[1:-1]}
        cerradas = " · ".join(f"{s}: {n}" for s, n in por_sit.items())
        texto = (f"Mostrando {len(filas)}   |   Activas: {len(activas)} (sin comentarios {sin_c}, vencidos {venc})"
                 f"   |   {cerradas}")
        sit = self.v_situacion.get()
        if not filas and sit in por_sit and por_sit[sit] == 0:
            texto += (f"\nAún no hay OCs en «{sit}»: una OC pasa a esa situación cuando un reporte posterior "
                      f"lo detecta (o, para «Cerrada sin factura», cuando tú la marcas así).")
        self.lbl_resumen.config(text=texto)
        if self.modo != "oc":
            return
        if self.oc_sel and self.tree.exists(str(self.oc_sel)):
            self.tree.selection_set(str(self.oc_sel))
            self.tree.see(str(self.oc_sel))
        elif self.oc_sel:
            self._guardar_pendientes()        # un filtro ocultó la OC: lo editado no se pierde
            self.limpiar_detalle()

    def ordenar(self, clave):
        desc = not self.orden[1] if self.orden[0] == clave else False
        self.orden = (clave, desc)
        self.pintar()

    def limpiar_filtros(self):
        self.v_buscar.set("")
        for v in (self.v_comprador, self.v_seg):
            v.set("Todos")
        self.v_estado_oc.set("Todos")
        self.v_situacion.set("Activas")
        self.v_sincom.set(False)
        self.v_venc.set(False)
        self.pintar()

    # ----------------------------------------------------------- detalle
    def seleccionar(self):
        sel = self.tree.selection()
        if not sel or self.modo != "oc":
            return
        oc = int(sel[0])
        if oc == self.oc_sel and self.txt_hist.get("1.0", "end").strip():
            return
        guardado = self._guardar_pendientes()      # lo editado en la OC anterior no se pierde
        self._guardar_borrador()
        self.oc_sel = oc
        if guardado:
            self.refrescar()
        else:
            self.mostrar_detalle()

    def _guardar_borrador(self):
        """Guarda el comentario sin enviar de la OC que está en el cuadro, para recuperarlo al volver a ella."""
        if self._com_oc:
            texto = self.txt_com.get("1.0", "end").strip()
            if texto:
                self.borradores[self._com_oc] = texto
            else:
                self.borradores.pop(self._com_oc, None)

    def limpiar_detalle(self):
        self._guardar_borrador()
        self._com_oc = None
        self.oc_sel = None
        self.fac_sel = None
        self.lbl_oc.config(text="Selecciona una orden de la lista")
        self.lbl_info.config(text="")
        for v in (self.v_det_seg, self.v_det_resp, self.v_det_comp):
            v.set("")
        self.btn_guardar.config(state="disabled")
        self.btn_com.config(state="disabled")
        self.lf_com.config(text="Nuevo comentario")
        self.txt_com.delete("1.0", "end")
        self.txt_hist.config(state="normal")
        self.txt_hist.delete("1.0", "end")
        self.txt_hist.config(state="disabled")

    def mostrar_detalle(self):
        f = next((x for x in self.filas if x["numero_oc"] == self.oc_sel), None)
        if self.modo == "fac":
            fx = next((x for x in self.facturas if f"{x['ruc']}|{x['n_factura']}" == self.fac_sel), None)
            if fx is None:
                return
            otras = [x for x in self.facturas if x["numero_oc"] == self.oc_sel]
            self.lbl_oc.config(text=f"OC {fx['numero_oc']} · {fx['proveedor']}")
            lineas = [f"Serie {fx['serie']} · Comprador: {fx['usuario']} · Estado OC: {fx['estado_oc']}",
                      f"Fecha OC: {fmt_fecha(fx['fecha_oc'])} · RUC {fx['ruc']}", "",
                      f"Factura {fx['n_factura']} ({fx['tipo'] or 's/tipo'}): {fx['estado_factura'] or 'sin estado'}",
                      f"{fx['mon']} {fx['ft_sin_imp'] or 0:,.2f} sin imp. · pendiente {fx['imp_pendiente'] or 0:,.2f}",
                      f"Emisión {fmt_fecha(fx['emision'])} · Vence {fmt_fecha(fx['vencimiento'])}"
                      + (f" · Pagada {fmt_fecha(fx['fecha_pago'])}" if fx["fecha_pago"] else "")]
            if len(otras) > 1:
                lineas += ["", f"Facturas de la OC ({len(otras)}):"] + \
                          [f"  • {x['n_factura']}: {x['estado_factura'] or 'sin estado'}" for x in otras]
            self.lbl_info.config(text="\n".join(lineas))
            editable = False
        else:
            if f is None:
                return
            self.lbl_oc.config(text=f"OC {f['numero_oc']} · {f['proveedor']}")
            info = (f"Serie {f['serie']} · Comprador: {f['usuario']}\n"
                    f"Fecha OC: {fmt_fecha(f['fecha_oc'])} ({sgm.dias_desde(f['fecha_oc'])} días) · Estado OC: {f['estado_oc']}\n"
                    f"Total sin IGV: {f['mon']} {f['total'] or 0:,.2f} · RUC {f['ruc']}\n"
                    f"Situación: {f['situacion']}" + (f" ({f['facturas']})" if f["facturas"] else "") +
                    f"\nEn seguimiento desde: {fmt_fecha(f['primer_visto'])}")
            self.lbl_info.config(text=info)
            editable = True
        if editable:
            self.v_det_seg.set(f["estado_seg"])
            self.v_det_resp.set(f["responsable"] or "")
            self.v_det_comp.set(fmt_fecha(f["fecha_compromiso"]))
        else:
            for v in (self.v_det_seg, self.v_det_resp, self.v_det_comp):
                v.set("")
        for w in (self.cb_det_seg, self.ent_resp, self.ent_comp):
            w.config(state=("readonly" if w is self.cb_det_seg else "normal") if editable else "disabled")
        self.btn_guardar.config(state="normal" if editable else "disabled")
        self.btn_com.config(state="normal")
        if self._com_oc != self.oc_sel:         # cada OC conserva su propio borrador de comentario
            self._guardar_borrador()
            self.txt_com.delete("1.0", "end")
            self.txt_com.insert("1.0", self.borradores.get(self.oc_sel, ""))
            self._com_oc = self.oc_sel
        self.lf_com.config(text=f"Nuevo comentario para la OC {self.oc_sel}")
        t = self.txt_hist
        t.config(state="normal")
        t.delete("1.0", "end")
        for c in sgm.historial(self.con, self.oc_sel):
            fecha = datetime.strptime(c["fecha"], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M")
            sis = c["tipo"] == "sistema"
            t.insert("end", f"{fecha} · {c['autor']}\n", "cab_sis" if sis else "cab")
            t.insert("end", c["texto"] + "\n\n", {"sistema": "sis", "cambio": "cambio"}.get(c["tipo"], ""))
        t.config(state="disabled")
        t.see("end")

    def _fila_sel(self):
        return next((x for x in self.filas if x["numero_oc"] == self.oc_sel), None)

    def _hay_pendientes(self):
        """¿Hay cambios escritos en el panel (estado, responsable, compromiso) que aún no se guardaron?"""
        if self.modo != "oc" or not self.oc_sel or str(self.btn_guardar.cget("state")) == "disabled":
            return False
        f = self._fila_sel()
        if f is None:
            return False
        return (self.v_det_seg.get() != f["estado_seg"] or self.v_det_resp.get().strip() != (f["responsable"] or "")
                or self.v_det_comp.get().strip() != fmt_fecha(f["fecha_compromiso"]))

    def _guardar_pendientes(self):
        """Guarda sin preguntar lo editado en el panel (al cambiar de OC, comentar, cargar…). True si guardó."""
        return self._hay_pendientes() and self.guardar_orden(refrescar=False)

    def guardar_orden(self, refrescar=True):
        f = self._fila_sel()
        if f is None:
            return False
        seg, resp, comp = self.v_det_seg.get(), self.v_det_resp.get().strip(), self.v_det_comp.get().strip()
        try:
            comp_iso = datetime.strptime(comp, "%d/%m/%Y").strftime("%Y-%m-%d") if comp else ""
        except ValueError:
            messagebox.showwarning("Fecha inválida", f"La fecha compromiso de la OC {self.oc_sel} debe tener el formato "
                                                     f"dd/mm/aaaa. Esa fecha no se guardó.", parent=self)
            if refrescar:
                return False
            comp_iso = f["fecha_compromiso"] or ""
        if seg == "Cerrar sin factura" and f["estado_seg"] != seg and not messagebox.askyesno(
                "Cerrar sin factura", f"¿Cerrar el seguimiento de la OC {self.oc_sel}?\n\nÚsalo cuando la orden no se va "
                f"a facturar. Si más adelante llega una factura, la app la marcará como Facturada.", parent=self):
            seg = f["estado_seg"]
            self.v_det_seg.set(seg)
        cambio = sgm.actualizar_orden(self.con, self.oc_sel, seg, resp, comp_iso, self.cfg.get("autor", ""))
        if cambio and refrescar:
            self.refrescar()
        return cambio

    def agregar_comentario(self):
        texto = self.txt_com.get("1.0", "end").strip()
        if not self.oc_sel or not texto:
            return
        self._guardar_pendientes()            # lo editado en estado/responsable/compromiso también se guarda
        sgm.comentar(self.con, self.oc_sel, texto, self.cfg.get("autor", ""))
        self.con.commit()
        self.txt_com.delete("1.0", "end")
        self.borradores.pop(self.oc_sel, None)
        self.refrescar()

    # ------------------------------------------------- carga / Excel
    def _en_segundo_plano(self, titulo, fn, al_terminar):
        win = tk.Toplevel(self)
        win.title(titulo)
        win.transient(self)
        win.resizable(False, False)
        win.protocol("WM_DELETE_WINDOW", lambda: None)      # no se puede cerrar mientras trabaja
        ttk.Label(win, text=titulo + "\nEsto puede tardar unos segundos…", padding=16, justify="center").pack()
        pb = ttk.Progressbar(win, mode="indeterminate", length=320)
        pb.pack(padx=16, pady=(0, 16))
        pb.start(12)
        win.update_idletasks()
        win.geometry(f"+{self.winfo_rootx() + self.winfo_width() // 2 - 180}+{self.winfo_rooty() + 200}")
        win.grab_set()
        cola = queue.Queue()                                # una cola por tarea

        def run():
            try:
                cola.put(("ok", fn(), None))
            except BaseException as e:  # noqa: BLE001 — cualquier error se muestra al usuario
                cola.put(("error", e, traceback.format_exc()))

        def esperar():
            try:
                estado, res, detalle = cola.get_nowait()
            except queue.Empty:
                self.after(150, esperar)
                return
            win.grab_release()
            win.destroy()
            if estado == "error":
                registrar_error(detalle)
                extra = "" if isinstance(res, ValueError) else "\n\nEl detalle técnico quedó en errores.log."
                messagebox.showerror("No se pudo completar", f"{res}{extra}", parent=self)
            else:
                al_terminar(res)

        threading.Thread(target=run, daemon=True).start()
        self.after(150, esperar)

    def cargar_csv(self):
        self._guardar_pendientes()
        inicial = self.cfg.get("ultimo_csv") or os.path.join(BASE, "Control de ordenes.csv")
        ruta = filedialog.askopenfilename(parent=self, title="Reporte 'Control de órdenes' descargado de SAP",
                                          initialdir=os.path.dirname(inicial), initialfile=os.path.basename(inicial),
                                          filetypes=[("Reporte SAP", "*.csv *.xlsx"), ("Todos", "*.*")])
        if not ruta:
            return
        self.cfg["ultimo_csv"] = ruta
        guardar_cfg(self.cfg)

        def paso2(datos):
            df, log, meta, fr, aviso = datos
            if aviso and not messagebox.askyesno("¿Reporte completo?", aviso + "\n\n¿Cargarlo de todas formas?",
                                                 icon="warning", default="no", parent=self):
                return
            self._en_segundo_plano("Actualizando el seguimiento", lambda: aplicar_carga(ruta, df, log, meta, fr), fin)

        def fin(r):
            res, salida, log = r
            self.cfg = leer_cfg()
            conteos = dict(log["conteos"])
            msg = (f"Filas limpias: {conteos.get('Filas finales en Sheet1')} · OCs: {conteos.get('OCs distintas')}\n"
                   f"Avisos de limpieza: {len(log['avisos'])} (ver hoja 'Log limpieza' del control)\n\n")
            if res["omitida"]:
                msg += ("⚠ " + res["motivo"] + "\nTampoco se generó un control nuevo: el último control sigue siendo "
                        "el más reciente.")
            else:
                msg = f"Control generado:\n{os.path.basename(salida)}\n\n" + msg
                msg += (f"Seguimiento actualizado:\n• {res['nuevas']} OCs nuevas sin factura\n"
                        f"• {res['facturadas']} OCs ya facturadas (cerradas)\n"
                        f"• {res.get('cerradas_sap', 0)} OCs cerradas y {res.get('canceladas_sap', 0)} canceladas en SAP "
                        f"sin factura (salen del seguimiento)\n"
                        f"• {res['fuera']} fuera del reporte\n"
                        f"• {res['reabiertas']} reabiertas\n• {res['activas']} en seguimiento activo\n\n"
                        f"Facturas: {res.get('facturas_nuevas', 0)} nuevas · {res.get('facturas_pagadas', 0)} pagadas desde la carga anterior\n"
                        f"Requerimientos completados con datos de SAP: {res.get('req_completados', 0)}")
            self.refrescar()
            if salida is None:
                messagebox.showwarning("Reporte no aplicado", msg, parent=self)
            elif messagebox.askyesno("Carga completada", msg + "\n\n¿Abrir el control en Excel?", parent=self):
                os.startfile(salida)

        self._en_segundo_plano("Leyendo el reporte SAP", lambda: preparar_carga(ruta), paso2)

    def regenerar_excel(self):
        self._guardar_pendientes()
        u = sgm.ultima_carga(self.con)
        if not u or not u["archivo"] or not os.path.exists(u["archivo"]):
            messagebox.showinfo("Sin reporte", "Primero carga un CSV de SAP.", parent=self)
            return

        def trabajo():
            df, log, meta, _ = ps.procesar(u["archivo"])
            con = sgm.conectar(DB)
            try:
                return generar_control(df, meta, log, con, fecha_de(meta, u["archivo"]))
            finally:
                con.close()

        def fin(salida):
            self.cfg["ultimo_control"] = salida
            guardar_cfg(self.cfg)
            if messagebox.askyesno("Excel actualizado", f"Se regeneró {os.path.basename(salida)} con los comentarios "
                                                        f"actuales.\n\n¿Abrirlo en Excel?", parent=self):
                os.startfile(salida)

        self._en_segundo_plano("Actualizando el control con los comentarios", trabajo, fin)

    def abrir_control(self):
        ruta = self.cfg.get("ultimo_control")
        if ruta and os.path.exists(ruta):
            os.startfile(ruta)
        else:
            messagebox.showinfo("Sin control", "Aún no se ha generado un control desde la aplicación.", parent=self)


def main():
    # Modo sin ventana (p. ej. desde el Programador de tareas):
    #   "Seguimiento Ordenes.exe" --cargar archivo.csv [--forzar]
    # --forzar aplica también un reporte que parece incompleto (sin --forzar se deja para revisarlo en la app).
    if len(sys.argv) >= 3 and sys.argv[1] == "--cargar":
        ahora = f"{datetime.now():%d/%m/%Y %H:%M}"
        with open(os.path.join(BASE, "ultima_carga.log"), "w", encoding="utf-8") as fh:
            try:
                res, salida, log = cargar_reporte(sys.argv[2], forzar="--forzar" in sys.argv[3:])
                estado = f"OMITIDO: {res['motivo']}" if res["omitida"] else "OK"
                fh.write(f"{ahora} {estado}\nControl: {salida or '(no se generó)'}\nSeguimiento: {res}\n")
                for c, v in log["conteos"]:
                    fh.write(f"  {c}: {v}\n")
            except ReporteIncompleto as e:
                fh.write(f"{ahora} REVISAR: no se aplicó la carga.\n{e}\n\nCárgalo desde la aplicación para "
                         f"confirmarlo, o vuelve a ejecutar con --forzar.\n")
                sys.exit(2)
            except BaseException:  # noqa: BLE001
                fh.write(f"{ahora} ERROR\n{traceback.format_exc()}")
                sys.exit(1)
        return
    try:
        App().mainloop()
    except Exception:  # noqa: BLE001
        messagebox.showerror("Error", traceback.format_exc())


if __name__ == "__main__":
    main()
