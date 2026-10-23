"""
Seguimiento de órdenes — aplicación de escritorio.

- "Cargar CSV de SAP": limpia el reporte, genera el control Excel del día y actualiza el seguimiento.
- Pestaña "OC sin factura": OCs en proceso sin factura; estado, responsable, fecha compromiso y comentarios.
- Pestaña "Facturas": facturas del último reporte (Pendiente de pago / Pagado) para revisar hasta su pago.
- Pestaña "Consultar OC": el status de cualquier orden (SAP, facturas, requerimientos y seguimiento) en un solo lugar.
- "Datos Power BI.xlsx": tablas para Power BI, regeneradas en cada carga, al cerrar y desde Indicadores.
- Los comentarios son por OC, se guardan en seguimiento.db (junto a la aplicación) y se arrastran a cada
  control nuevo; las facturas nuevas y los pagos detectados quedan registrados en el historial de la OC.
- Varios compradores pueden usar la misma base (carpeta de red): cada uno tiene su configuración en
  %APPDATA%\\Seguimiento Ordenes y las listas muestran por defecto lo suyo.
"""
import json
import os
import queue
import shutil
import sqlite3
import sys
import unicodedata
import threading
import traceback
from datetime import datetime, date

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import detalle_sap
import powerbi as pbi
import procesar_sap as ps
import seguimiento as sgm
import requerimientos as rq
import semaforo as sem
import app_requerimientos as ui_req
import app_consulta as ui_con

BASE = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE, "seguimiento.db")
CFG_ANTIGUA = os.path.join(BASE, "config_seguimiento.json")        # antes: una sola configuración junto a la app
CFG = os.path.join(os.environ.get("APPDATA") or BASE, "Seguimiento Ordenes", "config.json")    # una por usuario
CARGAS = os.path.join(BASE, "Cargas")
DIAS_RESPALDO = 15

SITUACIONES = ["Activas", sgm.SIT_SIN_FACT, sgm.SIT_FACTURADA, sgm.SIT_CERRADA_SAP, sgm.SIT_CANCELADA_SAP,
               sgm.SIT_FUERA, "Todas"]
COLS = [("oc", "OC", 85, "e"), ("usuario", "Comprador", 120, "w"), ("proveedor", "Proveedor", 290, "w"),
        ("mon", "Mon", 45, "center"), ("total", "Total sin IGV", 100, "e"), ("fecha_oc", "Fecha OC", 80, "center"),
        ("dias", "Días", 50, "e"), ("plazo", "Plazo", 75, "center"), ("estado_oc", "Estado OC", 80, "w"),
        ("logistica", "Logística", 150, "w"),
        ("situacion", "Situación", 110, "w"),
        ("estado_seg", "Seguimiento", 150, "w"), ("responsable", "Responsable", 110, "w"),
        ("compromiso", "Compromiso", 85, "center"), ("ult_com", "Último comentario", 330, "w")]
COLS_FAC = [("oc", "OC", 85, "e"), ("usuario", "Comprador", 115, "w"), ("proveedor", "Proveedor", 250, "w"),
            ("n_factura", "N° Factura", 110, "w"), ("tipo", "Tipo", 40, "center"), ("estado_factura", "Estado factura", 115, "w"),
            ("mon", "Mon", 45, "center"), ("ft_sin_imp", "FT sin imp.", 95, "e"), ("imp_pendiente", "Imp. pendiente", 100, "e"),
            ("emision", "Emisión", 80, "center"), ("vencimiento", "Vencimiento", 85, "center"),
            ("dias_venc", "Días p/ vencer", 85, "e"), ("plazo", "Plazo", 75, "center"), ("fecha_pago", "Fecha pago", 80, "center"),
            ("estado_oc", "Estado OC", 80, "w"), ("ult_com", "Último comentario (OC)", 300, "w")]
ESTADOS_FACT_FILTRO = ["Pendiente de pago", "Pagado", "Todos", "(sin estado)"]
MONEDAS = ["Todas", "SOL", "USD", "CAD", "AUD", "GBP", "EUR", "ZAR"]


def fmt_fecha(iso):
    return f"{iso[8:10]}/{iso[5:7]}/{iso[:4]}" if iso else ""


def texto_logistica(f):
    """Estado logístico de la OC (reporte de detalle), con la fecha de la última nota de entrada."""
    if not f.get("logistica"):
        return ""
    txt = f["logistica"]
    if f["logistica"] in (detalle_sap.PARCIAL, "Cerrada con atención parcial") and f.get("pct_ingresado") is not None:
        txt += f" ({f['pct_ingresado']:.0%})"
    if f.get("fecha_entrada") and f["logistica"] in (detalle_sap.ATENDIDA, detalle_sap.PARCIAL):
        txt += f" · {fmt_fecha(f['fecha_entrada'])}"
    return txt


def usuario_windows():
    try:
        return os.getlogin()
    except OSError:
        return os.environ.get("USERNAME", "")


def leer_cfg():
    try:
        with open(CFG, encoding="utf-8-sig") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        pass
    # Primera vez de este usuario: se toma la configuración antigua si era suya (los plazos, en cualquier caso)
    try:
        with open(CFG_ANTIGUA, encoding="utf-8-sig") as fh:
            vieja = json.load(fh)
    except (OSError, ValueError):
        return {}
    if str(vieja.get("autor", "")).lower() == usuario_windows().lower():
        return vieja
    return {k: v for k, v in vieja.items() if k == "plazos"}


def guardar_cfg(cfg):
    os.makedirs(os.path.dirname(CFG), exist_ok=True)
    with open(CFG, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)


def _sin_tildes(t):
    return "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn").lower()


def adivinar_comprador(login, nombres):
    """'asalinas' -> 'Alejandro Salinas' (inicial del nombre + apellido), si está entre los compradores."""
    login = _sin_tildes(login)
    for n in nombres:
        partes = _sin_tildes(n).split()
        if len(partes) >= 2 and login in {partes[0][0] + p for p in partes[1:]} | {partes[0]}:
            return n
    return ""


def respaldo_diario(con):
    """Copia de la base una vez al día en Respaldos\\Automáticos (se guardan los últimos DIAS_RESPALDO)."""
    carpeta = os.path.join(BASE, "Respaldos", "Automáticos")
    destino = os.path.join(carpeta, f"seguimiento_{date.today():%Y-%m-%d}.db")
    if os.path.exists(destino):
        return
    os.makedirs(carpeta, exist_ok=True)
    tmp = os.path.join(carpeta, f"~{os.getpid()}.tmp")      # nombre corto: rutas largas (límite de 260 caracteres)
    copia = sqlite3.connect(tmp)
    try:
        con.backup(copia)
    finally:
        copia.close()
    try:
        os.replace(tmp, destino)
    except OSError:              # el otro usuario lo hizo al mismo tiempo
        os.remove(tmp)
    viejos = sorted(f for f in os.listdir(carpeta) if f.startswith("seguimiento_") and f.endswith(".db"))
    for f in viejos[:-DIAS_RESPALDO]:
        try:
            os.remove(os.path.join(carpeta, f))
        except OSError:
            pass


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


def archivar(ruta, fr, nombre="Control de ordenes"):
    """Copia el reporte a Cargas\\ (historial). Si no se puede, devuelve la ruta original."""
    archivo = os.path.join(CARGAS, f"{nombre} {fr:%Y-%m-%d %H%M}{os.path.splitext(ruta)[1]}")
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
        try:
            respaldo_diario(con)
        except Exception:  # noqa: BLE001 — sin respaldo igual se carga; queda registrado
            registrar_error("Respaldo diario:\n" + traceback.format_exc())
        doc = sgm.normalizar_documentos(df)        # DI sin número cuentan como factura (el control usa df tal cual)
        res = sgm.sincronizar(con, doc, fr, archivo)
        if not res["omitida"]:
            rq.preparar(con)
            rq.guardar_snapshot_sap(con, doc)
            res["req_completados"] = rq.completar_desde_sap(con)
            salida = generar_control(df, meta, log, con, fr)
            con.execute("UPDATE cargas SET salida=? WHERE id=(SELECT MAX(id) FROM cargas)", (salida,))
            con.commit()
            try:
                pbi.exportar(con, BASE, sem.umbrales(leer_cfg()))
            except Exception:  # noqa: BLE001 — p. ej. abierto en Excel; se reintenta al cerrar la app
                registrar_error("Datos Power BI:\n" + traceback.format_exc())
    finally:
        con.close()
    if salida:
        cfg = leer_cfg()
        cfg["ultimo_control"] = salida
        guardar_cfg(cfg)
    return res, salida, log


def cargar_detalle(ruta):
    """Reporte de detalle (OC → entrada → factura → pago): lo archiva y reemplaza el detalle guardado."""
    df = detalle_sap.leer(ruta)
    archivo = archivar(ruta, datetime.fromtimestamp(os.path.getmtime(ruta)), "Seguimiento detalle")
    con = sgm.conectar(DB)
    try:
        rq.preparar(con)
        return detalle_sap.guardar(con, df, archivo)
    finally:
        con.close()


def reaplicar_control():
    """Vuelve a aplicar el último «Control de órdenes» (p. ej. tras cargar solo el detalle), para que el seguimiento
    use el estado logístico nuevo. Devuelve (resultado, ruta_control, log) o None si no hay un Control cargado."""
    con = sgm.conectar(DB)
    try:
        u = sgm.ultima_carga(con)
    finally:
        con.close()
    if not u or not u["archivo"] or not os.path.exists(u["archivo"]):
        return None
    df, log, meta, _ = ps.procesar(u["archivo"])
    return aplicar_carga(u["archivo"], df, log, meta, fecha_de(meta, u["archivo"]))


def separar_reportes(rutas):
    """(rutas de detalle, rutas de Control) según el contenido de cada archivo."""
    detalle = [r for r in rutas if detalle_sap.es_detalle(r)]
    return detalle, [r for r in rutas if r not in detalle]


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
        self.data_version = None             # cambia cuando otro usuario guarda en la base
        self._ui()
        self.bind_all("<F5>", lambda e: self.actualizar_vista())
        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        self.bind("<Activate>", lambda e: self.after_idle(self._revisar_cambios_externos))
        self.after(100, self._inicio)
        self.refrescar()

    def _inicio(self):
        self._pedir_usuario()
        self._filtros_mios()
        try:
            respaldo_diario(self.con)
        except Exception:  # noqa: BLE001
            registrar_error("Respaldo diario:\n" + traceback.format_exc())
        if self.cfg.get("pendientes_al_abrir", True):
            self.pendientes_de_hoy(auto=True)

    def _cerrar(self):
        """Al cerrar: guarda lo editado en el panel y deja al día los datos de Power BI."""
        try:
            self._guardar_pendientes()
            pbi.exportar(self.con, BASE, sem.umbrales(self.cfg))
        except Exception:  # noqa: BLE001 — cerrar nunca debe fallar
            registrar_error("Al cerrar:\n" + traceback.format_exc())
        self.destroy()

    def exportar_powerbi(self):
        self._guardar_pendientes()
        try:
            ruta = pbi.exportar(self.con, BASE, sem.umbrales(self.cfg))
        except PermissionError:
            messagebox.showwarning("Datos Power BI", f"«{pbi.ARCHIVO}» está abierto en Excel: ciérralo y vuelve a "
                                                     f"intentarlo.", parent=self)
            return
        if messagebox.askyesno("Datos Power BI", f"Se actualizó «{pbi.ARCHIVO}» en la carpeta de la aplicación.\n\n"
                                                 f"En Power BI pulsa «Actualizar» para ver los datos nuevos.\n\n"
                                                 f"¿Abrir la carpeta?", parent=self):
            os.startfile(os.path.dirname(ruta))

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
        """Al volver a la ventana: si otro usuario o una carga programada (--cargar) cambió la base, se refresca."""
        if self.con.execute("PRAGMA data_version").fetchone()[0] != self.data_version:
            self.actualizar_vista()

    # ------------------------------------------------------------------ UI
    def _ui(self):
        barra = ttk.Frame(self, padding=(8, 8, 8, 4))
        barra.pack(fill="x")
        ttk.Button(barra, text="📥  Cargar CSV de SAP…", style="Accent.TButton", command=self.cargar_csv).pack(side="left")
        ttk.Button(barra, text="🔄  Actualizar Excel con comentarios", command=self.regenerar_excel).pack(side="left", padx=6)
        ttk.Button(barra, text="📊  Abrir último control", command=self.abrir_control).pack(side="left")
        ttk.Button(barra, text="📁  Abrir carpeta", command=lambda: os.startfile(BASE)).pack(side="left", padx=6)
        ttk.Button(barra, text="🚦  Pendientes de hoy", style="Accent.TButton",
                   command=self.pendientes_de_hoy).pack(side="left", padx=(12, 0))
        ttk.Button(barra, text="⚙  Plazos…", command=self.editar_plazos).pack(side="left", padx=6)
        self.lbl_autor = ttk.Label(barra, text="")
        ttk.Button(barra, text="Cambiar", command=lambda: (self._pedir_usuario(True), self._filtros_mios())).pack(side="right")
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
        tab_con = ttk.Frame(self.nb)
        self.nb.add(tab_req, text="  Requerimientos  ")
        self.nb.add(tab_oc, text="  OC sin factura (en proceso)  ")
        self.nb.add(tab_fac, text="  Facturas (pendientes de pago / pagadas)  ")
        self.nb.add(tab_ind, text="  Indicadores  ")
        self.nb.add(tab_con, text="  🔎 Consultar OC  ")
        self.tab_req, self.tab_oc, self.tab_fac, self.tab_con = tab_req, tab_oc, tab_fac, tab_con
        self.tabs = {str(tab_req): "req", str(tab_oc): "oc", str(tab_fac): "fac", str(tab_ind): "ind", str(tab_con): "con"}
        ui_req.construir(self, tab_req, tab_ind)
        ui_con.construir(self, tab_con)
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
        self.v_semaforo = self._combo(filt, "Plazo:", sem.FILTROS, 18)

        filt2 = ttk.Frame(tab_oc, padding=(4, 0, 4, 4))
        filt2.pack(fill="x")
        self.v_sincom = tk.BooleanVar()
        ttk.Checkbutton(filt2, text="Solo sin comentarios", variable=self.v_sincom, command=self.pintar).pack(side="left")
        self.v_venc = tk.BooleanVar()
        ttk.Checkbutton(filt2, text="Solo compromiso vencido", variable=self.v_venc, command=self.pintar).pack(side="left", padx=10)
        self.v_logistica = self._combo(filt2, "Logística:", ["Todos"] + detalle_sap.LOGISTICA[:3] + [
            detalle_sap.SERVICIO, detalle_sap.SIN_DETALLE], 18)
        ttk.Button(filt2, text="Limpiar filtros", command=self.limpiar_filtros).pack(side="left", padx=(0, 16))
        self.lbl_resumen = ttk.Label(filt2, foreground="#404040", justify="left")
        self.lbl_resumen.pack(side="left")

        self.tree = self._tabla(tab_oc, COLS, self.ordenar, ("proveedor", "ult_com"))
        for nivel, color in sem.FONDO.items():
            self.tree.tag_configure(nivel, background=color)
        self.tree.tag_configure("vencido", foreground="#9C0006")
        self.tree.tag_configure("sincom", font=("Segoe UI", 9, "bold"))
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
        self.vf_semaforo = self._combo(ff2, "Plazo:", sem.FILTROS_FACT, 20, self.pintar_fac)
        self.vf_concom = tk.BooleanVar()
        ttk.Checkbutton(ff2, text="Solo con comentarios", variable=self.vf_concom, command=self.pintar_fac).pack(side="left", padx=10)
        self.vf_fuera = tk.BooleanVar()
        ttk.Checkbutton(ff2, text="Incluir facturas fuera del reporte", variable=self.vf_fuera, command=self.pintar_fac).pack(side="left")
        ttk.Button(ff2, text="Limpiar filtros", command=self.limpiar_filtros_fac).pack(side="left", padx=(10, 16))
        self.lbl_resumen_fac = ttk.Label(ff2, foreground="#404040", justify="left")
        self.lbl_resumen_fac.pack(side="left")
        self.tree_fac = self._tabla(tab_fac, COLS_FAC, self.ordenar_fac, ("proveedor", "ult_com"))
        for nivel, color in sem.FONDO.items():
            self.tree_fac.tag_configure(nivel, background=color)
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

    def _pedir_usuario(self, forzar=False):
        """Quién usa la app en este Windows: su comprador (las listas muestran lo suyo) y la firma de comentarios."""
        if forzar or not self.cfg.get("autor") or "comprador" not in self.cfg:
            login = usuario_windows()
            nombres = rq.compradores(self.con)
            win = tk.Toplevel(self)
            win.title("¿Quién usa la aplicación?")
            win.transient(self)
            win.resizable(False, False)
            c = ttk.Frame(win, padding=16)
            c.pack(fill="both", expand=True)
            ttk.Label(c, text="Comprador:").grid(row=0, column=0, sticky="w")
            v_comp = tk.StringVar(value=self.cfg.get("comprador") or adivinar_comprador(login, nombres))
            cb = ttk.Combobox(c, textvariable=v_comp, values=nombres, width=32)
            cb.grid(row=0, column=1, sticky="w", pady=3)
            ttk.Label(c, text="Tu nombre tal como aparece en la columna «Usuario» del reporte SAP. Las listas muestran "
                              "por defecto tus requerimientos, OC y facturas (con «Todos» ves también los del otro "
                              "comprador).", foreground="#595959", wraplength=420, justify="left").grid(
                row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))
            ttk.Label(c, text="Firma de comentarios:").grid(row=2, column=0, sticky="w")
            v_autor = tk.StringVar(value=self.cfg.get("autor") or login)
            ttk.Entry(c, textvariable=v_autor, width=34).grid(row=2, column=1, sticky="w", pady=3)

            def aceptar(e=None):
                self.cfg["comprador"] = " ".join(v_comp.get().split())
                self.cfg["autor"] = v_autor.get().strip() or self.cfg.get("autor") or login
                guardar_cfg(self.cfg)
                win.destroy()

            def cancelar():
                if not self.cfg.get("autor") or "comprador" not in self.cfg:
                    aceptar()                   # primera vez: se queda con lo propuesto (se puede cambiar luego)
                else:
                    win.destroy()
            pie = ttk.Frame(win, padding=(16, 0, 16, 12))
            pie.pack(fill="x")
            ttk.Button(pie, text="Aceptar", style="Accent.TButton", command=aceptar).pack(side="right")
            ttk.Button(pie, text="Cancelar", command=cancelar).pack(side="right", padx=6)
            win.protocol("WM_DELETE_WINDOW", cancelar)
            win.bind("<Return>", aceptar)
            win.bind("<Escape>", lambda e: cancelar())
            win.update_idletasks()
            win.geometry(f"+{self.winfo_rootx() + 200}+{self.winfo_rooty() + 120}")
            cb.focus_set()
            win.grab_set()
            self.wait_window(win)
        comp = self.cfg.get("comprador") or "(todos)"
        self.lbl_autor.config(text=f"Usuario: {comp} · firma: {self.cfg.get('autor', '(sin nombre)')}")

    def mio(self):
        """Valor del filtro «Comprador» por defecto: el comprador de este usuario (o Todos)."""
        return self.cfg.get("comprador") or "Todos"

    def _filtros_mios(self):
        for v in (self.v_comprador, self.vf_comprador, self.vr_comprador, self.vi_comprador):
            v.set(self.mio())
        self.pintar()
        self.pintar_fac()
        ui_req.pintar_req(self)
        ui_req.pintar_ind(self)

    # ------------------------------------------------------------- datos
    def refrescar(self):
        self.filas = sgm.listar(self.con)
        compradores = sorted({f["usuario"] for f in self.filas if f["usuario"]})
        self.v_comprador.widget.config(values=["Todos"] + compradores)
        self.data_version = self.con.execute("PRAGMA data_version").fetchone()[0]
        u = sgm.ultima_carga(self.con)
        self.ultima_carga_id = u["id"] if u else None
        if u:
            self.lbl_estado.config(text=f"Última carga: reporte SAP del {fmt_fecha(u['fecha_reporte'])} {u['fecha_reporte'][11:16]}"
                                        f" · cargado el {fmt_fecha(u['fecha_carga'])} {u['fecha_carga'][11:16]}"
                                        f" · {u['nuevas']} nuevas, {u['facturadas']} facturadas, {u['fuera']} fuera del reporte en esa carga"
                                        + self._texto_detalle())
        else:
            self.lbl_estado.config(text="Aún no se ha cargado ningún reporte. Usa «Cargar CSV de SAP…».")
        self.facturas = sgm.listar_facturas(self.con)
        self.vf_comprador.widget.config(values=["Todos"] + sorted({f["usuario"] for f in self.facturas if f["usuario"]}))
        self.pintar()
        self.pintar_fac()
        self.refrescar_req()
        ui_con.refrescar(self)
        if self.modo == "con" and self.con_oc:
            ui_con.mostrar_status(self, self.con_oc)
        if self.oc_sel and self.modo in ("oc", "fac", "con"):
            self.mostrar_detalle()        # el panel refleja lo que acaba de cambiar (p. ej. tras una carga)

    def _texto_detalle(self):
        d = detalle_sap.ultima_carga(self.con)
        if d is None:
            return "   |   Reporte de detalle: aún no cargado"
        return f"   |   Detalle (ingresos y líneas): archivo del {fmt_fecha(d['fecha_archivo'])} {d['fecha_archivo'][11:16]}"

    def refrescar_req(self):
        ui_req.refrescar(self)

    # -------------------------------------------------- pestaña Facturas
    def filtradas_fac(self):
        q = self.vf_buscar.get().strip().lower()
        est, comp, eoc, mon = self.vf_estado.get(), self.vf_comprador.get(), self.vf_estado_oc.get(), self.vf_mon.get()
        u, plazo = sem.umbrales(self.cfg), self.vf_semaforo.get()
        out = []
        for f in self.facturas:
            f["nivel"] = sem.nivel_factura(f, u)
            if not sem.pasa_filtro(f["nivel"], plazo):
                continue
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
        orden_nivel = {sem.ROJO: 0, sem.AMARILLO: 1, sem.VERDE: 2}
        num = {"oc": lambda f: f["numero_oc"], "ft_sin_imp": lambda f: f["ft_sin_imp"] or 0,
               "imp_pendiente": lambda f: f["imp_pendiente"] or 0,
               "dias_venc": lambda f: self._dias_venc(f) if self._dias_venc(f) is not None else 99999,
               "plazo": lambda f: (orden_nivel.get(f["nivel"], 3), self._dias_venc(f) or 0)}
        keyf = num.get(clave, lambda f: str(f.get(clave) or "").lower())
        filas.sort(key=keyf, reverse=desc)
        t = self.tree_fac
        t.delete(*t.get_children())
        for f in filas:
            tags = []
            if f["situacion"] != sgm.FACT_VIGENTE:
                tags.append("fuera")
            elif f["estado_factura"] == "Pagado":
                tags.append("pagada")
            elif f["nivel"] in sem.FONDO:
                tags.append(f["nivel"])
            dv = self._dias_venc(f)
            t.insert("", "end", iid=f"{f['ruc']}|{f['n_factura']}", tags=tags, values=(
                f["numero_oc"], f["usuario"], f["proveedor"], f["n_factura"], f["tipo"],
                "Pago parcial" if f["estado_detalle"] == "Pago parcial" else f["estado_factura"], f["mon"],
                f"{f['ft_sin_imp'] or 0:,.2f}", f"{f['imp_pendiente'] or 0:,.2f}", fmt_fecha(f["emision"]),
                fmt_fecha(f["vencimiento"]), "" if dv is None else dv, sem.TEXTO_FACT.get(f["nivel"], ""),
                fmt_fecha(f["fecha_pago"]), sgm.estado_visible(f["serie"], f["estado_oc"], f["fecha_oc"], f["ingresado"]),
                (f["ult_com"] or "").replace("\n", " ")))
        u = sem.umbrales(self.cfg)
        vig = [f for f in self.facturas if f["situacion"] == sgm.FACT_VIGENTE]
        pend = [f for f in vig if f["estado_factura"] == "Pendiente de pago"]
        venc = sum(1 for f in pend if sem.nivel_factura(f, u) == sem.ROJO)
        por_vencer = sum(1 for f in pend if sem.nivel_factura(f, u) == sem.AMARILLO)
        tot = {}
        for f in filas:
            if f["estado_factura"] != "Pagado":
                tot[f["mon"]] = tot.get(f["mon"], 0) + (f["imp_pendiente"] or 0)
        txt_tot = " · ".join(f"{m} {v:,.2f}" for m, v in sorted(tot.items()) if v)
        self.lbl_resumen_fac.config(
            text=f"Mostrando {len(filas)}   |   Pendientes de pago: {len(pend)} (vencidas {venc}, vencen en "
                 f"{u['fact']} días o menos {por_vencer}) · "
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
        self.vf_comprador.set(self.mio())
        self.vf_estado_oc.set("Todos")
        self.vf_mon.set("Todas")
        self.vf_semaforo.set("Todas")
        for v in (self.vf_concom, self.vf_fuera):
            v.set(False)
        self.pintar_fac()

    def cambiar_pestana(self):
        guardado = self._guardar_pendientes()
        self.modo = self.tabs[self.nb.select()]
        # El panel de detalle/comentarios es de las pestañas de OC y facturas
        visible = str(self.det) in self.paned.panes()
        if self.modo in ("oc", "fac", "con") and not visible:
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
        elif self.modo == "con" and self.tree_con.selection():
            ui_con.seleccionar(self)

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
        u, plazo = sem.umbrales(self.cfg), self.v_semaforo.get()
        out = []
        for f in self.filas:
            f["nivel"] = sem.nivel_oc(f, u)
            if not sem.pasa_filtro(f["nivel"], plazo):
                continue
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
            logi = self.v_logistica.get()
            if logi != "Todos" and (f["logistica"] or detalle_sap.SIN_DETALLE) != logi:
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
        orden_nivel = {sem.ROJO: 0, sem.AMARILLO: 1, sem.VERDE: 2}
        keyf = {"oc": lambda f: f["numero_oc"], "dias": lambda f: sgm.dias_desde(f["fecha_oc"]) or 0,
                "total": lambda f: f["total"] or 0, "fecha_oc": lambda f: f["fecha_oc"] or "",
                "compromiso": lambda f: f["fecha_compromiso"] or "9999",
                "plazo": lambda f: (orden_nivel.get(f["nivel"], 3), -(sgm.dias_desde(f["fecha_oc"]) or 0))
                }.get(clave, lambda f: str(f.get(clave) or "").lower())
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
                if f["nivel"] in sem.FONDO:
                    tags.append(f["nivel"])
                if not f["n_com"]:
                    tags.append("sincom")
                if f["fecha_compromiso"] and f["fecha_compromiso"] < hoy:
                    tags.append("vencido")
            self.tree.insert("", "end", iid=str(f["numero_oc"]), tags=tags, values=(
                f["numero_oc"], f["usuario"], f["proveedor"], f["mon"], f"{f['total'] or 0:,.2f}", fmt_fecha(f["fecha_oc"]),
                sgm.dias_desde(f["fecha_oc"]) or "", sem.texto(f["nivel"]),
                sgm.estado_visible(f["serie"], f["estado_oc"], f["fecha_oc"], f["ingresado"]), texto_logistica(f),
                f["situacion"], f["estado_seg"],
                f["responsable"] or "", fmt_fecha(f["fecha_compromiso"]), (f["ult_com"] or "").replace("\n", " ")))
        activas = [f for f in self.filas if f["situacion"] == sgm.SIT_ACTIVA]
        sin_c = sum(1 for f in activas if not f["n_com"])
        venc = sum(1 for f in activas if f["fecha_compromiso"] and f["fecha_compromiso"] < hoy)
        por_sit = {s: sum(1 for f in self.filas if f["situacion"] == s) for s in SITUACIONES[1:-1]}
        cerradas = " · ".join(f"{s}: {n}" for s, n in por_sit.items())
        u = sem.umbrales(self.cfg)
        rojas = sum(1 for f in activas if sem.nivel_oc(f, u) == sem.ROJO)
        amar = sum(1 for f in activas if sem.nivel_oc(f, u) == sem.AMARILLO)
        texto = (f"Mostrando {len(filas)}   |   Activas: {len(activas)} (atrasadas {rojas}, en atención {amar}, "
                 f"sin comentarios {sin_c} en negrita, compromisos vencidos {venc})   |   {cerradas}")
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
        for v in (self.v_seg, self.v_semaforo):
            v.set("Todos")
        self.v_comprador.set(self.mio())
        self.v_estado_oc.set("Todos")
        self.v_situacion.set("Activas")
        self.v_logistica.set("Todos")
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
            lineas = [f"Serie {fx['serie']} · Comprador: {fx['usuario']} · Estado OC: {sgm.estado_visible(fx['serie'], fx['estado_oc'], fx['fecha_oc'], fx['ingresado'])}",
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
        elif self.modo == "con":
            x = self.consulta.get(self.oc_sel)
            if x is None:
                return
            self.lbl_oc.config(text=f"OC {x['oc']} · {x['proveedor']}")
            lineas = [f"Comprador: {x['comprador'] or '-'} · Estado OC: {x['estado_sap']}", f"Status: {x['status']}", ""]
            lineas.append(f"Seguimiento de OC sin factura: {f['situacion']}" if f is not None else
                          "No pasó por «OC sin factura»: igual puedes comentarla y el comentario queda en su historial.")
            self.lbl_info.config(text="\n".join(lineas))
            editable = f is not None
        else:
            if f is None:
                return
            self.lbl_oc.config(text=f"OC {f['numero_oc']} · {f['proveedor']}")
            info = (f"Serie {f['serie']} · Comprador: {f['usuario']}\n"
                    f"Fecha OC: {fmt_fecha(f['fecha_oc'])} ({sgm.dias_desde(f['fecha_oc'])} días) · Estado OC: {sgm.estado_visible(f['serie'], f['estado_oc'], f['fecha_oc'], f['ingresado'])}\n"
                    f"Logística: {texto_logistica(f) or 'sin datos del reporte de detalle'}\n"
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
        if self.modo not in ("oc", "con") or not self.oc_sel or str(self.btn_guardar.cget("state")) == "disabled":
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

    # ------------------------------------------------- semáforo de plazos
    def pendientes_de_hoy(self, auto=False):
        """Resumen de lo atrasado y lo que requiere atención, con acceso directo a cada lista."""
        u = sem.umbrales(self.cfg)
        hoy = date.today()

        def cuenta(filas, fn):
            niveles = [fn(f) for f in filas]
            return niveles.count(sem.ROJO), niveles.count(sem.AMARILLO)

        yo = self.cfg.get("comprador")              # solo lo de este comprador (si no se indicó, todo)
        abiertos = [f for f in self.filas_req if f["etapa"] not in rq.CERRADAS and (not yo or f["comprador"] == yo)]
        hasta = [f for f in abiertos if f["etapa"] in rq.ETAPAS_HASTA_OC]
        despues = [f for f in abiertos if f["etapa"] not in rq.ETAPAS_HASTA_OC]
        activas = [f for f in self.filas if f["situacion"] == sgm.SIT_ACTIVA and (not yo or f["usuario"] == yo)]
        r1, a1 = cuenta(hasta, lambda f: sem.nivel_req(f, u))
        r2, a2 = cuenta(despues, lambda f: sem.nivel_req(f, u))
        r3, a3 = cuenta(activas, lambda f: sem.nivel_oc(f, u))
        r4, a4 = cuenta([f for f in self.facturas if not yo or f["usuario"] == yo], lambda f: sem.nivel_factura(f, u))
        comp_venc = sum(1 for f in activas if f["fecha_compromiso"] and f["fecha_compromiso"] < hoy.isoformat())
        comp_hoy = sum(1 for f in activas if f["fecha_compromiso"] == hoy.isoformat())
        if auto and not any((r1, a1, r2, a2, r3, a3, r4, a4, comp_hoy)):
            return

        win = tk.Toplevel(self)
        win.title("Pendientes de hoy")
        win.transient(self)
        win.resizable(False, False)
        cuerpo = ttk.Frame(win, padding=16)
        cuerpo.pack(fill="both", expand=True)
        ttk.Label(cuerpo, text=f"🚦 Pendientes de hoy · {hoy:%d/%m/%Y}" + (f" · {yo}" if yo else ""),
                  style="Titulo.TLabel").grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))

        def chip(parent, texto_chip, nivel, fila, col):
            tk.Label(parent, text=texto_chip, bg=sem.FONDO[nivel] if texto_chip[0] != "0" else "#F2F2F2",
                     font=("Segoe UI", 9, "bold"), padx=8, pady=2).grid(row=fila, column=col, sticky="w", padx=4)

        lineas = [
            ("Requerimientos hasta el envío de la OC", f"{r1} atrasados", f"{a1} en atención", "req_hasta"),
            ("Requerimientos después de la OC (factura y pago)", f"{r2} atrasados", f"{a2} en atención", "req_abiertos"),
            ("OC sin factura (en proceso)", f"{r3} atrasadas", f"{a3} en atención", "oc"),
            ("Facturas pendientes de pago", f"{r4} vencidas", f"{a4} vencen en {u['fact']} días o menos", "fac"),
        ]
        for i, (titulo, roja, amarilla, destino) in enumerate(lineas, start=1):
            ttk.Label(cuerpo, text=titulo, font=("Segoe UI", 10)).grid(row=i, column=0, sticky="w", pady=3)
            chip(cuerpo, roja, sem.ROJO, i, 1)
            chip(cuerpo, amarilla, sem.AMARILLO, i, 2)
            ttk.Button(cuerpo, text="Ver →", command=lambda d=destino: (self._ir(d), win.destroy())).grid(
                row=i, column=3, sticky="e", padx=(8, 0))
        ttk.Label(cuerpo, text=f"Compromisos de OC vencidos: {comp_venc} · vencen hoy: {comp_hoy}",
                  foreground="#9C0006" if comp_venc else "#404040").grid(row=5, column=0, columnspan=4, sticky="w", pady=(6, 0))

        # Lo más atrasado hasta el envío de la OC
        top = sorted((f for f in hasta if sem.nivel_req(f, u) == sem.ROJO), key=lambda f: -(f["dias_etapa"] or 0))[:5]
        if top:
            ttk.Label(cuerpo, text="Lo más atrasado hasta el envío de la OC:", font=("Segoe UI", 9, "bold")).grid(
                row=6, column=0, columnspan=4, sticky="w", pady=(12, 2))
            for j, f in enumerate(top, start=7):
                desc = f["proveedor"] or f["proyecto_nombre"] or f["proyecto"]
                ttk.Label(cuerpo, text=f"  • COD {f['id']:03d} · {f['etapa']} · {f['dias_etapa']} días · {desc[:55]}").grid(
                    row=j, column=0, columnspan=4, sticky="w")

        pie = ttk.Frame(win, padding=(16, 0, 16, 12))
        pie.pack(fill="x")
        v_abrir = tk.BooleanVar(value=self.cfg.get("pendientes_al_abrir", True))

        def cambiar_abrir():
            self.cfg["pendientes_al_abrir"] = v_abrir.get()
            guardar_cfg(self.cfg)
        ttk.Checkbutton(pie, text="Mostrar al abrir la aplicación", variable=v_abrir, command=cambiar_abrir).pack(side="left")
        ttk.Button(pie, text="Cerrar", command=win.destroy).pack(side="right")
        ttk.Button(pie, text="⚙  Plazos…", command=lambda: (win.destroy(), self.editar_plazos())).pack(side="right", padx=6)
        win.bind("<Escape>", lambda e: win.destroy())
        win.update_idletasks()
        win.geometry(f"+{self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)}+{self.winfo_rooty() + 120}")
        win.focus_set()
        return win

    def _ir(self, destino):
        """Abre la pestaña y deja aplicados los filtros de «solo lo atrasado y en atención»."""
        if destino in ("req_hasta", "req_abiertos"):
            self.nb.select(self.tab_req)
            for v in (self.vr_tipo, self.vr_controller, self.vr_mes):
                v.set("Todos")
            self.vr_comprador.set(self.mio())
            self.vr_buscar.set("")
            self.vr_etapa.set("Pendientes hasta envío de OC" if destino == "req_hasta" else "Abiertos (todos)")
            self.vr_semaforo.set("Atrasado + Atención")
            self.orden_req = ("plazo", False)
            ui_req.pintar_req(self)
        elif destino == "oc":
            self.nb.select(self.tab_oc)
            self.limpiar_filtros()
            self.v_semaforo.set("Atrasado + Atención")
            self.orden = ("plazo", False)
            self.pintar()
        else:
            self.nb.select(self.tab_fac)
            self.limpiar_filtros_fac()
            self.vf_semaforo.set("Vencidas + por vencer")
            self.orden_fac = ("dias_venc", False)
            self.pintar_fac()

    def editar_plazos(self):
        """Umbrales del semáforo: desde cuántos días un pendiente pasa a «Atención» y a «Atrasado»."""
        u = sem.umbrales(self.cfg)
        win = tk.Toplevel(self)
        win.title("Plazos del semáforo")
        win.transient(self)
        win.resizable(False, False)
        cuerpo = ttk.Frame(win, padding=16)
        cuerpo.pack(fill="both", expand=True)
        ttk.Label(cuerpo, text="Días a partir de los cuales un pendiente pasa a «Atención» (amarillo) y a «Atrasado» (rojo).",
                  wraplength=560).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        for c, t in enumerate(["Requerimientos: días en la etapa", "Atención desde", "Atrasado desde"]):
            ttk.Label(cuerpo, text=t, font=("Segoe UI", 9, "bold")).grid(row=1, column=c, sticky="w", padx=4)
        vars_req = {}
        fila = 2
        for etapa in rq.NOMBRES_ETAPA[:-1]:
            a, r = u["req"][etapa]
            va, vr = tk.StringVar(value=str(a)), tk.StringVar(value=str(r))
            vars_req[etapa] = (va, vr)
            ttk.Label(cuerpo, text=etapa + ("  (hasta la OC)" if etapa in rq.ETAPAS_HASTA_OC else "")).grid(
                row=fila, column=0, sticky="w", padx=4)
            ttk.Spinbox(cuerpo, from_=0, to=365, width=6, textvariable=va).grid(row=fila, column=1, sticky="w", padx=4, pady=1)
            ttk.Spinbox(cuerpo, from_=0, to=365, width=6, textvariable=vr).grid(row=fila, column=2, sticky="w", padx=4, pady=1)
            fila += 1
        ttk.Separator(cuerpo).grid(row=fila, column=0, columnspan=3, sticky="ew", pady=8)
        fila += 1
        oa, orr = tk.StringVar(value=str(u["oc"][0])), tk.StringVar(value=str(u["oc"][1]))
        ttk.Label(cuerpo, text="OC sin factura: días desde la fecha de la OC").grid(row=fila, column=0, sticky="w", padx=4)
        ttk.Spinbox(cuerpo, from_=0, to=999, width=6, textvariable=oa).grid(row=fila, column=1, sticky="w", padx=4)
        ttk.Spinbox(cuerpo, from_=0, to=999, width=6, textvariable=orr).grid(row=fila, column=2, sticky="w", padx=4)
        fila += 1
        ttk.Label(cuerpo, text="(una OC con el compromiso vencido siempre queda «Atrasada»)", foreground="#808080").grid(
            row=fila, column=0, columnspan=3, sticky="w", padx=4)
        fila += 1
        fv = tk.StringVar(value=str(u["fact"]))
        ttk.Label(cuerpo, text="Facturas: «Por vencer» cuando faltan (días)").grid(row=fila, column=0, sticky="w", padx=4, pady=(6, 0))
        ttk.Spinbox(cuerpo, from_=0, to=90, width=6, textvariable=fv).grid(row=fila, column=1, sticky="w", padx=4, pady=(6, 0))
        ttk.Label(cuerpo, text="(vencida = «Vencida»)", foreground="#808080").grid(row=fila, column=2, sticky="w", padx=4, pady=(6, 0))

        def restaurar():
            for etapa, (va, vr) in vars_req.items():
                va.set(str(sem.PLAZOS_REQ[etapa][0]))
                vr.set(str(sem.PLAZOS_REQ[etapa][1]))
            oa.set(str(sem.PLAZO_OC[0]))
            orr.set(str(sem.PLAZO_OC[1]))
            fv.set(str(sem.AVISO_FACTURA))

        def guardar():
            try:
                nuevo = {"req": {e: (int(va.get()), int(vr.get())) for e, (va, vr) in vars_req.items()},
                         "oc": (int(oa.get()), int(orr.get())), "fact": int(fv.get())}
            except ValueError:
                messagebox.showwarning("Plazos", "Escribe los días como números enteros.", parent=win)
                return
            malos = [e for e, (a, r) in list(nuevo["req"].items()) + [("OC sin factura", nuevo["oc"])] if a > r]
            if malos:
                messagebox.showwarning("Plazos", "«Atención desde» no puede ser mayor que «Atrasado desde» en:\n• "
                                       + "\n• ".join(malos), parent=win)
                return
            sem.guardar_umbrales(self.cfg, nuevo)
            guardar_cfg(self.cfg)
            win.destroy()
            self.refrescar()

        pie = ttk.Frame(win, padding=(16, 0, 16, 12))
        pie.pack(fill="x")
        ttk.Button(pie, text="Restaurar valores iniciales", command=restaurar).pack(side="left")
        ttk.Button(pie, text="Cancelar", command=win.destroy).pack(side="right")
        ttk.Button(pie, text="💾  Guardar", style="Accent.TButton", command=guardar).pack(side="right", padx=6)
        win.bind("<Escape>", lambda e: win.destroy())
        win.grab_set()
        return win

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
        rutas = filedialog.askopenfilenames(
            parent=self, title="Reportes de SAP: «Control de órdenes» y/o el de detalle (puedes elegir los dos con Ctrl)",
            initialdir=os.path.dirname(inicial), filetypes=[("Reportes SAP", "*.csv *.xlsx"), ("Todos", "*.*")])
        if not rutas:
            return
        detalle, control = separar_reportes(list(rutas))
        if len(control) > 1 or len(detalle) > 1:
            messagebox.showwarning("Cargar", "Elige como máximo un «Control de órdenes» y un reporte de detalle.", parent=self)
            return
        if control:
            self.cfg["ultimo_csv"] = control[0]
            guardar_cfg(self.cfg)
        if detalle:
            def tras_detalle(r):
                texto = (f"Reporte de detalle cargado: {r['ocs']} OC aprobadas, {r['lineas']} líneas, "
                         f"{r['entradas']} notas de entrada y {r['facturas']} facturas.")
                if control:
                    self.refrescar()
                    self._cargar_control(control[0], previo=texto)
                    return

                def fin_solo(r2):
                    self.refrescar()
                    extra = ("\n\nSe volvió a aplicar el último «Control de órdenes» con el estado logístico nuevo "
                             f"({r2[0].get('nuevas', 0)} OC entraron al seguimiento)." if r2 else
                             "\n\nAún no hay un «Control de órdenes» cargado: cárgalo para actualizar el seguimiento.")
                    messagebox.showinfo("Carga completada", texto + extra, parent=self)
                self._en_segundo_plano("Actualizando el seguimiento con el detalle", reaplicar_control, fin_solo)
            self._en_segundo_plano("Leyendo el reporte de detalle", lambda: cargar_detalle(detalle[0]), tras_detalle)
        else:
            self._cargar_control(control[0])

    def _cargar_control(self, ruta, previo=""):
        """Carga del «Control de órdenes»: revisión de reporte completo, sincronización y control del día."""

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
            if previo:
                msg = previo + "\n\n" + msg
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
        u = sgm.ultima_carga(self.con)
        rutas = [r for r in (self.cfg.get("ultimo_control"), u["salida"] if u else None) if r and os.path.exists(r)]
        if rutas:
            os.startfile(max(rutas, key=os.path.getmtime))   # la carga pudo hacerla el otro comprador
        else:
            messagebox.showinfo("Sin control", "Aún no se ha generado un control desde la aplicación.", parent=self)


def main():
    # Modo sin ventana (p. ej. desde el Programador de tareas):
    #   "Seguimiento Ordenes.exe" --cargar "Control de ordenes.csv" ["Seguimiento.xlsx"] [--forzar]
    # Acepta el «Control de órdenes», el reporte de detalle o los dos (el detalle se carga primero).
    # --forzar aplica también un reporte que parece incompleto (sin --forzar se deja para revisarlo en la app).
    if len(sys.argv) >= 3 and sys.argv[1] == "--cargar":
        ahora = f"{datetime.now():%d/%m/%Y %H:%M}"
        detalle, control = separar_reportes([a for a in sys.argv[2:] if not a.startswith("--")])
        with open(os.path.join(BASE, "ultima_carga.log"), "w", encoding="utf-8") as fh:
            try:
                det = [cargar_detalle(d) for d in detalle]
                if not control:
                    r = reaplicar_control()
                    fh.write(f"{ahora} OK (solo detalle; se reaplicó el último Control: {r[0] if r else 'no hay'})\n"
                             f"Detalle: {det}\n")
                    return
                res, salida, log = cargar_reporte(control[0], forzar="--forzar" in sys.argv)
                estado = f"OMITIDO: {res['motivo']}" if res["omitida"] else "OK"
                fh.write(f"{ahora} {estado}\nControl: {salida or '(no se generó)'}\nSeguimiento: {res}\n")
                if det:
                    fh.write(f"Detalle: {det}\n")
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
