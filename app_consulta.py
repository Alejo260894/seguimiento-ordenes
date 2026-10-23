"""
Pestaña «Consultar OC» de la aplicación de seguimiento: el status de cualquier orden en un solo lugar.

Se construye sobre la App de app_seguimiento.py (usa su conexión, sus pestañas y el panel de historial y
comentarios de abajo, que en esta pestaña sirve para cualquier orden).
"""
import tkinter as tk
from tkinter import ttk

import consulta as cq
import detalle_sap
import seguimiento as sgm
import app_requerimientos as ui_req

COLS_CON = [("oc", "OC", 85, "e"), ("comprador", "Comprador", 115, "w"), ("proveedor", "Proveedor", 230, "w"),
            ("fecha", "Fecha OC", 78, "center"), ("estado_sap", "Estado OC", 80, "w"),
            ("logistica", "Logística", 130, "w"), ("status", "Status", 260, "w")]
VER = ["Búsqueda", "Pendientes de ingreso", detalle_sap.PENDIENTE, detalle_sap.PARCIAL]
FONDO = {"ok": "#E2EFDA", "atencion": "#FCE4D6"}
MAXIMO = 300


def construir(app, tab):
    app.consulta = {}
    app._con_busqueda = None
    top = ttk.Frame(tab, padding=(4, 6))
    top.pack(fill="x")
    ttk.Label(top, text="Buscar orden:", font=("Segoe UI", 10, "bold")).pack(side="left")
    app.vc_buscar = tk.StringVar()
    app.ent_con = ttk.Entry(top, textvariable=app.vc_buscar, width=36, font=("Segoe UI", 10))
    app.ent_con.pack(side="left", padx=(6, 10))
    ttk.Label(top, text="N° de OC o de factura, RUC, proveedor, o COD / referencia del requerimiento   ·   Ctrl+F "
                        "desde cualquier pestaña", foreground="#808080").pack(side="left")
    app.vc_buscar.trace_add("write", lambda *_: _programar(app))
    app.ent_con.bind("<Return>", lambda e: pintar(app, elegir=True))
    app.ent_con.bind("<Down>", lambda e: _a_la_lista(app))
    app.vc_ver = app._combo(top, "   Ver:", VER, 22, lambda: pintar(app))
    app.lbl_res_con = ttk.Label(tab, padding=(6, 0, 6, 4), foreground="#404040")
    app.lbl_res_con.pack(fill="x")

    pan = ttk.PanedWindow(tab, orient="horizontal")
    pan.pack(fill="both", expand=True)
    izq = ttk.Frame(pan)
    der = ttk.LabelFrame(pan, text="Status de la orden", padding=6)
    pan.add(izq, weight=3)
    pan.add(der, weight=2)
    app.tree_con = app._tabla(izq, COLS_CON, lambda k: None, ("proveedor", "status"))
    for nivel, color in FONDO.items():
        app.tree_con.tag_configure(nivel, background=color)
    app.tree_con.tag_configure("cerrada", foreground="#808080")
    app.tree_con.bind("<<TreeviewSelect>>", lambda e: seleccionar(app))

    der.rowconfigure(0, weight=1)
    der.columnconfigure(0, weight=1)
    app.txt_con = tk.Text(der, wrap="word", font=("Segoe UI", 9), height=12, relief="flat", background="#FAFAFA")
    sy = ttk.Scrollbar(der, command=app.txt_con.yview)
    app.txt_con.configure(yscrollcommand=sy.set)
    app.txt_con.grid(row=0, column=0, sticky="nsew")
    sy.grid(row=0, column=1, sticky="ns")
    app.txt_con.tag_configure("titulo", font=("Segoe UI", 10, "bold"), foreground="#1F3864")
    app.txt_con.tag_configure("status", font=("Segoe UI", 10, "bold"), foreground="#9C5700")
    app.txt_con.bind("<Key>", lambda e: None if (e.state & 4 and e.keysym.lower() in ("c", "a")) else "break")
    bot = ttk.Frame(der)
    bot.grid(row=1, column=0, columnspan=2, sticky="we", pady=(6, 0))
    app.btn_con_copiar = ttk.Button(bot, text="📋  Copiar status", style="Accent.TButton", command=lambda: copiar(app))
    app.btn_con_copiar.pack(side="left")
    app.btn_con_oc = ttk.Button(bot, text="Ver en OC sin factura", command=lambda: ver_oc(app))
    app.btn_con_oc.pack(side="left", padx=(8, 0))
    app.btn_con_fac = ttk.Button(bot, text="Ver en Facturas", command=lambda: ver_fac(app))
    app.btn_con_fac.pack(side="left", padx=(6, 0))
    app.btn_con_req = ttk.Button(bot, text="Abrir requerimiento", command=lambda: abrir_req(app))
    app.btn_con_req.pack(side="left", padx=(6, 0))
    app.lbl_con_copiado = ttk.Label(bot, text="", foreground="#2E7D32")
    app.lbl_con_copiado.pack(side="left", padx=8)
    _limpiar_status(app)
    app.bind_all("<Control-f>", lambda e: enfocar(app) if e.widget.winfo_toplevel() is app else None)
    app.bind_all("<Control-F>", lambda e: enfocar(app) if e.widget.winfo_toplevel() is app else None)


def refrescar(app):
    """Vuelve a armar el universo de órdenes (tras una carga, un comentario o un cambio en requerimientos)."""
    app.consulta = cq.universo(app.con, app.filas_req)
    pintar(app)


def enfocar(app):
    app.nb.select(app.tab_con)
    app.ent_con.focus_set()
    app.ent_con.select_range(0, "end")
    return "break"


def _programar(app):
    if app._con_busqueda:
        app.after_cancel(app._con_busqueda)
    app._con_busqueda = app.after(250, lambda: pintar(app, elegir=True))


def _a_la_lista(app):
    hijos = app.tree_con.get_children()
    if hijos:
        app.tree_con.focus_set()
        app.tree_con.focus(hijos[0])
        app.tree_con.selection_set(hijos[0])
    return "break"


def pintar(app, elegir=False):
    """Lista las órdenes que coinciden con lo escrito. `elegir`: si hay una sola (o el N° exacto), la muestra."""
    app._con_busqueda = None
    q = app.vc_buscar.get().strip().lower()
    t = app.tree_con
    sel = t.selection()
    t.delete(*t.get_children())
    ver = app.vc_ver.get()
    if ver != "Búsqueda":
        return _pintar_pendientes(app, ver, q, sel)
    if len(q) < 2:
        app.lbl_res_con.config(text=f"Escribe para buscar entre las {len(app.consulta):,} órdenes que conoce la "
                                    f"aplicación (todas las del último reporte SAP, más las de seguimiento, facturas y "
                                    f"requerimientos).")
        return
    palabras = q.split()
    hits = [x for x in app.consulta.values() if all(p in x["texto_busqueda"] for p in palabras)]
    hits.sort(key=lambda x: x["fecha"] or "", reverse=True)
    hits.sort(key=lambda x: str(x["oc"]) != q)            # el N° exacto primero; luego las más recientes
    for x in hits[:MAXIMO]:
        t.insert("", "end", iid=str(x["oc"]), tags=(x["nivel"],), values=(
            x["oc"], x["comprador"], x["proveedor"], cq._fmt(x["fecha"]), x["estado_sap"], _logistica(x), x["status"]))
    extra = f" (se muestran las {MAXIMO} más recientes: escribe algo más específico)" if len(hits) > MAXIMO else ""
    app.lbl_res_con.config(text=f"{len(hits)} órdenes encontradas{extra}   ·   colores: verde = pagada, naranja = "
                                f"requiere atención (sin factura en proceso o factura vencida), gris = cerrada / cancelada")
    if sel and t.exists(sel[0]):
        t.selection_set(sel[0])
    elif elegir and hits and (len(hits) == 1 or str(hits[0]["oc"]) == q):
        t.selection_set(str(hits[0]["oc"]))
        t.see(str(hits[0]["oc"]))


def _logistica(x):
    sap = x["sap"] or {}
    txt = x["logistica"]
    if txt == detalle_sap.PARCIAL and sap.get("pct_atendido") is not None:
        txt += f" ({sap['pct_atendido']:.0%})"
    return txt


def _pintar_pendientes(app, ver, q, sel):
    """OC de bienes abiertas con material por ingresar (más antiguas primero)."""
    estados = detalle_sap.PENDIENTES if ver == "Pendientes de ingreso" else (ver,)
    filas = [x for x in app.consulta.values() if x["logistica"] in estados
             and (not q or all(p in x["texto_busqueda"] for p in q.split()))]
    filas.sort(key=lambda x: x["fecha"] or "")
    t = app.tree_con
    for x in filas:
        t.insert("", "end", iid=str(x["oc"]), tags=("atencion",), values=(
            x["oc"], x["comprador"], x["proveedor"], cq._fmt(x["fecha"]), x["estado_sap"], _logistica(x), x["status"]))
    app.lbl_res_con.config(text=f"{len(filas)} OC de bienes {ver.lower()}   ·   las más antiguas primero   ·   una línea "
                                f"facturada al 100 % cuenta como atendida aunque no tenga nota de entrada   ·   "
                                f"la búsqueda filtra esta lista")
    if sel and t.exists(sel[0]):
        t.selection_set(sel[0])


def seleccionar(app):
    sel = app.tree_con.selection()
    if not sel or app.modo != "con":
        return
    oc = int(sel[0])
    mostrar_status(app, oc)
    if oc == app.oc_sel and app.txt_hist.get("1.0", "end").strip():
        return
    guardado = app._guardar_pendientes()       # lo editado en la orden anterior no se pierde
    app._guardar_borrador()
    app.oc_sel = oc
    if guardado:
        app.refrescar()
    else:
        app.mostrar_detalle()


def _limpiar_status(app):
    app.txt_con.delete("1.0", "end")
    app.txt_con.insert("1.0", "Elige una orden de la lista para ver su status completo: datos de SAP, facturas, "
                              "requerimientos, seguimiento y último comentario.\n\nAbajo puedes ver su historial y "
                              "agregar un comentario, aunque la orden no esté en «OC sin factura».")
    for b in (app.btn_con_copiar, app.btn_con_oc, app.btn_con_fac, app.btn_con_req):
        b.config(state="disabled")
    app.btn_con_req.config(text="Abrir requerimiento")
    app.con_oc = None


def mostrar_status(app, oc):
    x = app.consulta.get(oc)
    if x is None:
        return _limpiar_status(app)
    app.con_oc = oc
    u = sgm.ultima_carga(app.con)
    texto = cq.texto_status(x, u["fecha_reporte"] if u else None, sgm.historial(app.con, oc),
                            lineas=detalle_sap.lineas_oc(app.con, oc))
    t = app.txt_con
    t.delete("1.0", "end")
    t.insert("1.0", texto)
    t.tag_add("titulo", "1.0", "1.end")
    pos = t.search("STATUS:", "1.0", "end")
    if pos:
        t.tag_add("status", pos, f"{pos} lineend")
    d = cq.donde(x)
    app.btn_con_copiar.config(state="normal")
    app.btn_con_oc.config(state="normal" if d["oc"] else "disabled")
    app.btn_con_fac.config(state="normal" if d["fac"] else "disabled")
    reqs = d["req"]
    app.btn_con_req.config(state="normal" if reqs else "disabled",
                           text=f"Abrir requerimiento {reqs[0]:03d}" if len(reqs) == 1 else
                           f"Abrir requerimientos ({len(reqs)})" if reqs else "Abrir requerimiento")
    app.lbl_con_copiado.config(text="")


def copiar(app):
    app.clipboard_clear()
    app.clipboard_append(app.txt_con.get("1.0", "end").strip())
    app.lbl_con_copiado.config(text="✔ Copiado: pégalo en el correo o chat")


def ver_oc(app):
    """Abre «OC sin factura» con la orden a la vista (sin filtros que la oculten)."""
    oc = app.con_oc
    app.nb.select(app.tab_oc)
    app.limpiar_filtros()
    app.v_comprador.set("Todos")
    app.v_situacion.set("Todas")
    app.v_buscar.set(str(oc))
    app.pintar()
    app.after(50, lambda: app.tree.exists(str(oc)) and (app.tree.selection_set(str(oc)), app.tree.see(str(oc))))


def ver_fac(app):
    """Abre «Facturas» con las facturas de la orden (todos los estados, también las que salieron del reporte)."""
    oc = app.con_oc
    app.nb.select(app.tab_fac)
    app.limpiar_filtros_fac()
    app.vf_estado.set("Todos")
    app.vf_comprador.set("Todos")
    app.vf_fuera.set(True)
    app.vf_buscar.set(str(oc))
    app.pintar_fac()
    hijos = app.tree_fac.get_children()
    if hijos:
        app.after(50, lambda: (app.tree_fac.selection_set(hijos[0]), app.tree_fac.see(hijos[0])))


def abrir_req(app):
    ids = cq.donde(app.consulta[app.con_oc])["req"]
    filas = {f["id"]: f for f in app.filas_req}
    if len(ids) == 1:
        ui_req.EditorReq(app, filas[ids[0]])
        return
    menu = tk.Menu(app, tearoff=0)
    for i in ids:
        f = filas[i]
        menu.add_command(label=f"COD {i:03d} · {f['referencia'] or f['proyecto']} · {f['etapa']}",
                         command=lambda f=f: ui_req.EditorReq(app, f))
    b = app.btn_con_req
    menu.tk_popup(b.winfo_rootx(), b.winfo_rooty() + b.winfo_height())
