"""
Pestañas «Requerimientos» e «Indicadores» de la aplicación de seguimiento.

Se construyen sobre la App de app_seguimiento.py (usa su conexión, configuración y carpeta base).
"""
import os
from datetime import date, datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import requerimientos as rq
import semaforo as sem

COLS_REQ = [("id", "COD", 50, "e"), ("f_solic", "F. Solic.", 80, "center"), ("proyecto", "Proyecto", 80, "w"),
            ("referencia", "Referencia", 95, "w"), ("tipo", "Tipo", 50, "center"), ("proveedor", "Proveedor", 220, "w"),
            ("etapa", "Etapa", 165, "w"), ("dias_etapa", "Días etapa", 70, "e"), ("plazo", "Plazo", 75, "center"),
            ("monto", "Monto", 90, "e"),
            ("moneda", "Mon", 45, "center"), ("orden", "Orden", 85, "e"), ("sap_estado", "Estado OC SAP", 95, "w"),
            ("controller", "Controller", 180, "w"), ("estatus", "Estatus", 120, "w"),
            ("com_oc", "Último comentario OC", 260, "w")]
FILTRO_ETAPA = ["Pendientes hasta envío de OC", "Abiertos (todos)"] + rq.NOMBRES_ETAPA + rq.CERRADAS[1:] + ["Cerrados", "Todos"]


def iso_desde_ddmm(s):
    s = s.strip()
    return datetime.strptime(s, "%d/%m/%Y").strftime("%Y-%m-%d") if s else ""


def leer_monto(s):
    """Monto escrito como 1234.50, 1,234.50, 1234,50 o 1.234,50 -> float (ValueError si no es un número)."""
    s = s.strip().replace(" ", "")
    if not s:
        return None
    if "," in s and "." in s:                       # el último separador es el decimal
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:                                  # solo comas: decimal si hay 1 o 2 dígitos al final
        ent, _, dec = s.rpartition(",")
        s = f"{ent.replace(',', '')}.{dec}" if len(dec) in (1, 2) and s.count(",") == 1 else s.replace(",", "")
    return float(s)


class Autocompletar:
    """Lista desplegable de sugerencias bajo un Entry (busca por todas las palabras escritas).
    ↓ para entrar en la lista, Enter o clic para elegir, Esc para cerrar."""
    TECLAS_IGNORADAS = {"Down", "Up", "Return", "Escape", "Tab", "Left", "Right", "Home", "End", "Shift_L", "Shift_R",
                        "Control_L", "Control_R", "Alt_L", "Alt_R", "Caps_Lock", "Win_L", "Win_R"}

    def __init__(self, entry, opciones, al_elegir=None, maximo=60):
        self.entry, self.opciones, self.al_elegir, self.maximo = entry, opciones, al_elegir, maximo
        self.popup = self.lb = None
        entry.bind("<KeyRelease>", self._tecla, add="+")
        entry.bind("<FocusOut>", lambda e: entry.after(150, self._cerrar_si_sin_foco), add="+")
        entry.bind("<Down>", self._bajar, add="+")
        entry.bind("<Escape>", lambda e: self._cerrar(), add="+")

    def _tecla(self, e):
        if e.keysym in self.TECLAS_IGNORADAS:
            return
        palabras = self.entry.get().lower().split()
        if not palabras:
            return self._cerrar()
        hits = [o for o in self.opciones() if all(p in o.lower() for p in palabras)][: self.maximo]
        if not hits:
            return self._cerrar()
        if self.popup is None:
            self.popup = tk.Toplevel(self.entry)
            self.popup.overrideredirect(True)
            self.popup.attributes("-topmost", True)
            self.lb = tk.Listbox(self.popup, height=8, font=("Segoe UI", 9), activestyle="dotbox")
            self.lb.pack(fill="both", expand=True)
            self.lb.bind("<ButtonRelease-1>", self._elegir)
            self.lb.bind("<Return>", self._elegir)
            self.lb.bind("<Escape>", lambda e: (self._cerrar(), self.entry.focus_set()))
            self.lb.bind("<FocusOut>", lambda e: self.entry.after(150, self._cerrar_si_sin_foco))
        x, y = self.entry.winfo_rootx(), self.entry.winfo_rooty() + self.entry.winfo_height()
        self.popup.geometry(f"{max(self.entry.winfo_width(), 520)}x{min(len(hits), 8) * 18 + 6}+{x}+{y}")
        self.lb.delete(0, "end")
        for h in hits:
            self.lb.insert("end", h)

    def _bajar(self, e):
        if self.popup is not None:
            self.lb.focus_set()
            self.lb.selection_clear(0, "end")
            self.lb.selection_set(0)
            self.lb.activate(0)
            return "break"

    def _elegir(self, e=None):
        sel = self.lb.curselection() if self.lb is not None else ()
        if not sel:
            return
        valor = self.lb.get(sel[0])
        self._cerrar()
        if self.al_elegir:
            self.al_elegir(valor)
        else:
            self.entry.delete(0, "end")
            self.entry.insert(0, valor)
        self.entry.focus_set()

    def _cerrar_si_sin_foco(self):
        try:
            foco = self.entry.focus_get()
        except KeyError:            # el foco está en un widget sin nombre (p. ej. otra aplicación)
            foco = None
        if foco not in (self.entry, self.lb):
            self._cerrar()

    def _cerrar(self):
        if self.popup is not None:
            self.popup.destroy()
            self.popup = self.lb = None


class EditorReq(tk.Toplevel):
    """Formulario para crear/editar un requerimiento."""

    def __init__(self, app, req=None):
        super().__init__(app)
        self.app, self.con = app, app.con
        self.req = req or {}
        nuevo = not req
        self.title("Nuevo requerimiento" if nuevo else f"Requerimiento {req['id']:03d}")
        # Tamaño ajustado a la pantalla (en portátiles o con zoom de Windows los botones deben verse)
        ancho, alto = min(1180, self.winfo_screenwidth() - 40), min(800, self.winfo_screenheight() - 90)
        self.geometry(f"{ancho}x{alto}+{max(0, (self.winfo_screenwidth() - ancho) // 2)}+10")
        self.transient(app)
        self.v = {}
        self.sap_lbl = {}
        self._sap_form = {}            # campos que el formulario tomó de SAP: {campo: valor mostrado}
        # Botonera primero y abajo: siempre visible aunque la ventana sea pequeña
        pie = ttk.Frame(self, padding=(10, 0, 10, 10))
        pie.pack(side="bottom", fill="x")
        self.lbl_etapa = ttk.Label(pie, text="", font=("Segoe UI", 10, "bold"), foreground="#1F3864")
        self.lbl_etapa.pack(side="left")
        ttk.Button(pie, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(pie, text="💾  Guardar  (Ctrl+G)", style="Accent.TButton", command=self.guardar).pack(side="right", padx=6)
        self.bind("<Control-g>", lambda e: self.guardar())
        self.bind("<Escape>", lambda e: self.destroy())
        cuerpo = ttk.Frame(self, padding=10)
        cuerpo.pack(fill="both", expand=True)
        cuerpo.columnconfigure(0, weight=1)
        cuerpo.columnconfigure(1, weight=1)

        # ---- Detalle
        det = ttk.LabelFrame(cuerpo, text="Detalle del requerimiento", padding=8)
        det.grid(row=0, column=0, columnspan=2, sticky="nsew")
        for c in (1, 3, 5):
            det.columnconfigure(c, weight=1)
        ttk.Label(det, text="Proyecto:").grid(row=0, column=0, sticky="w")
        self.v["proyecto"] = tk.StringVar()
        e_proy = ttk.Entry(det, textvariable=self.v["proyecto"], width=14)
        e_proy.grid(row=0, column=1, sticky="w")
        self.lbl_proy = ttk.Label(det, text="Escribe código o nombre del proyecto", foreground="#595959", wraplength=780)
        self.lbl_proy.grid(row=0, column=2, columnspan=4, sticky="w", padx=6)
        self.proyectos = {r["codigo"]: dict(r) for r in self.con.execute("SELECT * FROM cat_proyectos")}
        ops_proy = [f"{c} — {p['nombre_corto'] or p['nombre']} [{p['estado']}]" for c, p in self.proyectos.items()]
        Autocompletar(e_proy, lambda: ops_proy, self._elegir_proyecto)
        e_proy.bind("<FocusOut>", lambda e: self._mostrar_proyecto(), add="+")

        self._campo(det, 1, 0, "referencia", "Referencia:", ancho=18)
        ttk.Label(det, text="(SOLPED N°, CORREO…)", foreground="#808080").grid(row=1, column=2, sticky="w")
        self._combo(det, 1, 3, "tipo", "Tipo:", self._valores("tipo", rq.TIPOS))
        self._fecha(det, 1, 5, "f_solic", "F. Solic.:")
        ttk.Label(det, text="Proveedor:").grid(row=2, column=0, sticky="w")
        self.v["proveedor"] = tk.StringVar()
        e_prov = ttk.Entry(det, textvariable=self.v["proveedor"], width=60)
        e_prov.grid(row=2, column=1, columnspan=3, sticky="we", pady=2)
        provs = [r[0] for r in self.con.execute("SELECT nombre FROM cat_proveedores ORDER BY nombre")]
        Autocompletar(e_prov, lambda: provs)
        self._combo(det, 2, 4, "estatus", "Estatus:", self._valores("estatus", list(rq.ESTATUS_CIERRE)), editable=True)
        self._fecha(det, 3, 0, "fecha_inicio", "Fecha inicio:")
        self._fecha(det, 3, 2, "fecha_fin", "Fecha fin:")
        self.w_vinc = {"monto": self._campo(det, 4, 0, "monto", "Monto:", ancho=14),
                       "moneda": self._combo(det, 4, 2, "moneda", "Moneda:", rq.MONEDAS)}
        self.lbl_vinc = ttk.Label(det, text="", foreground="#2F5597", wraplength=520, justify="left")
        self.lbl_vinc.grid(row=4, column=4, columnspan=2, sticky="w", padx=10)

        # ---- Flujos
        fv = ttk.LabelFrame(cuerpo, text="Flujo valorización", padding=8)
        fv.grid(row=1, column=0, sticky="nsew", pady=(8, 0), padx=(0, 4))
        self._fecha(fv, 0, 0, "envio_valo", "Envío de valorización:")
        self._fecha(fv, 1, 0, "envio_aprob_valo", "Envío para aprobación:")
        self._fecha(fv, 2, 0, "aprob_valo", "Aprobación de valo:")

        fo = ttk.LabelFrame(cuerpo, text="Flujo orden", padding=8)
        fo.grid(row=1, column=1, sticky="nsew", pady=(8, 0), padx=(4, 0))
        ttk.Label(fo, text="N° de orden (SAP):").grid(row=0, column=0, sticky="w")
        self.v["orden"] = tk.StringVar()
        e_orden = ttk.Entry(fo, textvariable=self.v["orden"], width=14)
        e_orden.grid(row=0, column=1, sticky="w")
        e_orden.bind("<FocusOut>", lambda e: self._mostrar_sap())      # al cambiar la orden se vuelve a vincular
        e_orden.bind("<Return>", lambda e: self._mostrar_sap())
        ttk.Button(fo, text="⬇ Traer de SAP", command=self.traer_sap).grid(row=0, column=2, sticky="w", padx=6)
        self.lbl_sap = ttk.Label(fo, text="", foreground="#2F5597", wraplength=430, justify="left")
        self.lbl_sap.grid(row=4, column=0, columnspan=4, sticky="w", pady=(4, 0))
        self._fecha(fo, 1, 0, "emision_oc", "Emisión de orden:")
        self._fecha(fo, 2, 0, "aprob_oc", "Aprobación de orden:")
        self._fecha(fo, 3, 0, "envio_oc", "Envío de orden:")

        fp = ttk.LabelFrame(cuerpo, text="Flujo pago", padding=8)
        fp.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        self._fecha(fp, 0, 0, "recepcion_fact", "Recepción de factura:")
        self._fecha(fp, 0, 4, "envio_aprob_fact", "Envío de aprobación de factura:")
        self._fecha(fp, 1, 0, "validacion_fact", "Validación de factura:")
        self._campo(fp, 1, 4, "num_factura", "N° de factura:", ancho=28)
        self._fecha(fp, 2, 0, "envio_pago", "Envío para pago:")
        self._fecha(fp, 2, 4, "fecha_pago", "Fecha de pago:")

        # ---- Notas e historial
        nh = ttk.Frame(cuerpo)
        nh.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        cuerpo.rowconfigure(3, weight=1)
        nh.columnconfigure(0, weight=1)
        nh.columnconfigure(1, weight=2)
        nh.rowconfigure(0, weight=1)
        fn = ttk.LabelFrame(nh, text="Notas", padding=6)
        fn.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        fn.rowconfigure(0, weight=1)
        fn.columnconfigure(0, weight=1)
        self.txt_notas = tk.Text(fn, height=5, wrap="word", font=("Segoe UI", 9))
        self.txt_notas.grid(row=0, column=0, sticky="nsew")
        fh = ttk.LabelFrame(nh, text="Historial", padding=6)
        fh.grid(row=0, column=1, sticky="nsew")
        fh.rowconfigure(0, weight=1)
        fh.columnconfigure(0, weight=1)
        self.txt_hist = tk.Text(fh, height=5, wrap="word", font=("Segoe UI", 9), state="disabled", background="#FAFAFA")
        self.txt_hist.grid(row=0, column=0, sticky="nsew")
        self.txt_hist.tag_configure("cab", font=("Segoe UI", 8, "bold"), foreground="#1F3864")
        self.txt_hist.tag_configure("sis", foreground="#808080", font=("Segoe UI", 9, "italic"))

        self._cargar()
        self.grab_set()
        e_proy.focus_set()

    # ---------- construcción de campos
    def _valores(self, campo, base):
        usados = [r[0] for r in self.con.execute(f"SELECT DISTINCT {campo} FROM req WHERE {campo}<>''")]
        return list(dict.fromkeys(base + sorted(usados)))

    def _campo(self, parent, fila, col, campo, etiqueta, ancho=20):
        ttk.Label(parent, text=etiqueta).grid(row=fila, column=col, sticky="w")
        self.v[campo] = tk.StringVar()
        w = ttk.Entry(parent, textvariable=self.v[campo], width=ancho)
        w.grid(row=fila, column=col + 1, sticky="w", pady=2)
        return w

    def _combo(self, parent, fila, col, campo, etiqueta, valores, editable=False):
        ttk.Label(parent, text=etiqueta).grid(row=fila, column=col, sticky="w", padx=(10 if col else 0, 0))
        self.v[campo] = tk.StringVar()
        w = ttk.Combobox(parent, textvariable=self.v[campo], values=valores, width=22,
                         state="normal" if editable else "readonly")
        w.grid(row=fila, column=col + 1, sticky="w", pady=2)
        return w

    def _fecha(self, parent, fila, col, campo, etiqueta):
        ttk.Label(parent, text=etiqueta).grid(row=fila, column=col, sticky="w", padx=(10 if col else 0, 0))
        self.v[campo] = tk.StringVar()
        fr = ttk.Frame(parent)
        fr.grid(row=fila, column=col + 1, sticky="w", pady=2)
        ttk.Entry(fr, textvariable=self.v[campo], width=11).pack(side="left")
        ttk.Button(fr, text="Hoy", width=4,
                   command=lambda: self.v[campo].set(date.today().strftime("%d/%m/%Y"))).pack(side="left", padx=2)
        self.sap_lbl[campo] = ttk.Label(fr, text="", foreground="#2F5597")
        self.sap_lbl[campo].pack(side="left", padx=2)

    # ---------- datos
    def _cargar(self):
        r = self.req
        for campo, _, tipo in rq.CAMPOS:
            if campo == "id" or campo not in self.v:
                continue
            val = r.get(campo)
            if tipo == "fecha":
                self.v[campo].set(rq.fmt(val or ""))
            elif campo == "monto":
                self.v[campo].set("" if val is None else f"{val:,.2f}")
            else:
                self.v[campo].set("" if val is None else str(val))
        if not r:
            self.v["f_solic"].set(date.today().strftime("%d/%m/%Y"))
        self.txt_notas.insert("1.0", r.get("notas") or "")
        self._mostrar_proyecto()
        self._mostrar_sap()
        self.txt_hist.config(state="normal")
        for e in rq.eventos(self.con, r["id"]) if r else []:
            f = datetime.strptime(e["fecha"], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M")
            self.txt_hist.insert("end", f"{f} · {e['autor']}\n", "cab")
            self.txt_hist.insert("end", e["texto"] + "\n\n", "sis" if e["tipo"] == "sistema" else "")
        self.txt_hist.config(state="disabled")
        self.txt_hist.see("end")
        if r:
            marcados = ", ".join(rq.ETIQUETAS[c] for c in r.get("desde_sap", []))
            self.lbl_etapa.config(text=f"Etapa: {r['etapa']}" + (f"   ·   Datos desde SAP: {marcados}" if marcados else ""))

    def _elegir_proyecto(self, texto):
        self.v["proyecto"].set(texto.split(" — ")[0])
        self._mostrar_proyecto()

    def _mostrar_proyecto(self):
        p = self.proyectos.get(self.v["proyecto"].get().strip().upper())
        if p:
            self.v["proyecto"].set(p["codigo"])
            self.lbl_proy.config(foreground="#1F3864", text=(
                f"{p['nombre_corto'] or p['nombre']} · {p['estado']}\nController: {p['controller1'] or '-'} · "
                f"Controller 2: {p['controller2'] or '-'} · Jefe de proyecto: {p['responsable'] or '-'}"))
        elif self.v["proyecto"].get().strip():
            self.lbl_proy.config(foreground="#C00000", text="Proyecto no encontrado en el catálogo (Consulta1).")

    def _vincular(self, d):
        """Monto y moneda bloqueados con lo aprobado en SAP cuando la orden está en el reporte.
        Si la OC la comparten varios requerimientos, el monto queda libre (cada uno registra su parte)."""
        otros = []
        if d:
            otros = [rid for rid in rq.usos_oc(self.con).get(int(self.v["orden"].get()), []) if rid != self.req.get("id")]
        for campo, w in self.w_vinc.items():
            if d and d.get(campo) not in (None, "") and not (campo == "monto" and otros):
                valor = f"{d[campo]:,.2f}" if campo == "monto" else d[campo]
                self.v[campo].set(valor)
                self._sap_form[campo] = valor
                w.config(state="disabled")
            else:
                self._sap_form.pop(campo, None)
                w.config(state="normal" if campo == "monto" else "readonly")
        if not d:
            texto = ""
        elif otros:
            suma = sum(r[0] or 0 for r in self.con.execute(
                f"SELECT monto FROM req WHERE id IN ({','.join('?' * len(otros))})", otros))
            try:
                suma += leer_monto(self.v["monto"].get()) or 0
            except ValueError:
                pass
            ok = abs(suma - (d["monto"] or 0)) < 0.01
            texto = (f"⚠ OC compartida con el requerimiento {', '.join(f'{o:03d}' for o in otros)}: el total SAP "
                     f"({d['moneda']} {d['monto'] or 0:,.2f}) es de la orden completa. Registra aquí solo la parte de "
                     f"este requerimiento. Suma actual de los requerimientos: {suma:,.2f} "
                     + ("✔ cuadra con SAP" if ok else "✘ no cuadra con SAP"))
        else:
            texto = "🔗 Monto y moneda vinculados a SAP (total aprobado de la OC)"
        self.lbl_vinc.config(text=texto, foreground="#2F5597" if not otros else "#9C5700")

    def _mostrar_sap(self):
        for l in self.sap_lbl.values():
            l.config(text="")
        try:
            d = rq.datos_sap(self.con, int(self.v["orden"].get())) if self.v["orden"].get().strip() else None
        except ValueError:
            d = None
        self._vincular(d)
        if not d:
            self.lbl_sap.config(text="" if not self.v["orden"].get().strip() else "La orden no está en el último reporte SAP cargado.")
            return None
        for campo in rq.CAMPOS_FECHA:
            if d.get(campo) and campo in self.sap_lbl:
                self.sap_lbl[campo].config(text=f"SAP: {rq.fmt(d[campo])}")
        facts = ", ".join(f"{f['n_factura']} ({f['estado'] or 'sin estado'})" for f in d["facturas"]) or "sin factura"
        self.lbl_sap.config(text=f"SAP: {d['proveedor']} · {d['moneda']} {d['monto'] or 0:,.2f} · OC {d['estado_oc']}\n"
                                 f"Facturas: {facts}")
        return d

    def traer_sap(self):
        d = self._mostrar_sap()
        if not d:
            messagebox.showinfo("SAP", "Esa orden no está en el último reporte SAP cargado.", parent=self)
            return
        llenados = [rq.ETIQUETAS[c] + " (vinculado)" for c in rq.CAMPOS_VINCULADOS if c in self._sap_form]
        for campo in rq.CAMPOS_SAP:
            v = d.get(campo)
            if v in (None, "") or campo in rq.CAMPOS_VINCULADOS or self.v[campo].get().strip():
                continue
            valor = rq.fmt(v) if campo in rq.CAMPOS_FECHA else str(v)
            self.v[campo].set(valor)
            self._sap_form[campo] = valor          # queda marcado como dato de SAP al guardar
            llenados.append(rq.ETIQUETAS[campo])
        messagebox.showinfo("SAP", ("Completado (solo campos vacíos): " + ", ".join(llenados)) if llenados
                            else "No había campos vacíos que completar con SAP.", parent=self)

    def guardar(self):
        vals = {"id": self.req.get("id")}
        try:
            for campo, _, tipo in rq.CAMPOS:
                if campo == "id":
                    continue
                s = self.v[campo].get().strip()
                if tipo == "fecha":
                    vals[campo] = iso_desde_ddmm(s)
                elif campo == "monto":
                    vals[campo] = leer_monto(s)
                elif tipo == "int":
                    vals[campo] = int(s) if s else None
                else:
                    vals[campo] = s.upper() if campo in ("estatus", "tipo", "proyecto") else s
        except ValueError:
            messagebox.showwarning("Dato inválido", f"Revisa «{rq.ETIQUETAS[campo]}»: las fechas van como dd/mm/aaaa, "
                                                    f"el monto como número (1234.50 o 1.234,50) y la orden solo con "
                                                    f"dígitos.", parent=self)
            return
        vals["notas"] = self.txt_notas.get("1.0", "end").strip()
        if not vals["proyecto"] and not vals["referencia"]:
            messagebox.showwarning("Faltan datos", "Indica al menos el proyecto o la referencia.", parent=self)
            return
        if vals["proyecto"] and vals["proyecto"] not in self.proyectos and not messagebox.askyesno(
                "Proyecto no encontrado", f"«{vals['proyecto']}» no está en el catálogo de proyectos, así que no tendrá "
                f"controller ni jefe de proyecto.\n\n¿Guardar de todas formas?", parent=self):
            return
        # Lo que el formulario tomó de SAP (y no se modificó a mano) queda marcado como dato de SAP
        vals["_sap"] = [c for c, v in self._sap_form.items() if self.v[c].get().strip() == str(v).strip()]
        rid = rq.guardar(self.con, vals, self.app.cfg.get("autor", ""))
        self.app.refrescar_req()
        self.app.tree_req.selection_set(str(rid)) if self.app.tree_req.exists(str(rid)) else None
        self.destroy()


# ---------------------------------------------------------------------------
def construir(app, tab_req, tab_ind):
    """Agrega a la App las pestañas de requerimientos e indicadores."""
    app.filas_req = []
    app.orden_req = ("id", True)

    barra = ttk.Frame(tab_req, padding=(4, 6))
    barra.pack(fill="x")
    ttk.Button(barra, text="➕  Nuevo requerimiento", style="Accent.TButton", command=lambda: EditorReq(app)).pack(side="left")
    ttk.Button(barra, text="✏  Editar", command=lambda: editar(app)).pack(side="left", padx=6)
    ttk.Button(barra, text="📤  Exportar Excel", command=lambda: exportar(app)).pack(side="left")
    ttk.Button(barra, text="📚  Actualizar catálogos", command=lambda: catalogos(app)).pack(side="left", padx=6)
    ttk.Button(barra, text="📥  Importar desde Excel…", command=lambda: importar(app)).pack(side="left")
    app.lbl_cat = ttk.Label(barra, text="", foreground="#808080")
    app.lbl_cat.pack(side="right")

    f1 = ttk.Frame(tab_req, padding=(4, 0, 4, 4))
    f1.pack(fill="x")
    ttk.Label(f1, text="Buscar:").pack(side="left")
    app.vr_buscar = tk.StringVar()
    ttk.Entry(f1, textvariable=app.vr_buscar, width=24).pack(side="left", padx=(4, 12))
    app.vr_buscar.trace_add("write", lambda *_: pintar_req(app))
    app.vr_etapa = app._combo(f1, "Etapa:", FILTRO_ETAPA, 28, lambda: pintar_req(app))
    app.vr_tipo = app._combo(f1, "Tipo:", ["Todos"] + rq.TIPOS, 8, lambda: pintar_req(app))
    app.vr_controller = app._combo(f1, "Controller:", ["Todos"], 26, lambda: pintar_req(app))
    app.vr_mes = app._combo(f1, "Mes solicitud:", ["Todos"], 9, lambda: pintar_req(app))
    app.vr_semaforo = app._combo(f1, "Plazo:", sem.FILTROS, 18, lambda: pintar_req(app))
    app.lbl_resumen_req = ttk.Label(tab_req, padding=(6, 0, 6, 4), foreground="#404040")
    app.lbl_resumen_req.pack(fill="x")
    app.tree_req = app._tabla(tab_req, COLS_REQ, lambda k: ordenar_req(app, k), ("proveedor", "com_oc"))
    for nivel, color in sem.FONDO.items():
        app.tree_req.tag_configure(nivel, background=color)
    app.tree_req.tag_configure("pagado", background="#E2EFDA")
    app.tree_req.tag_configure("cerrado", foreground="#808080")
    app.tree_req.bind("<Double-1>", lambda e: editar(app))
    app.tree_req.bind("<Return>", lambda e: editar(app))

    # ---- Indicadores
    fi = ttk.Frame(tab_ind, padding=(4, 6))
    fi.pack(fill="x")
    app.vi_tipo = app._combo(fi, "Tipo:", ["Todos"] + rq.TIPOS, 8, lambda: pintar_ind(app))
    app.vi_desde = app._combo(fi, "Solicitudes desde:", ["Todos"], 9, lambda: pintar_ind(app))
    app.vi_hasta = app._combo(fi, "hasta:", ["Todos"], 9, lambda: pintar_ind(app))
    app.vi_controller = app._combo(fi, "Controller:", ["Todos"], 26, lambda: pintar_ind(app))
    ttk.Button(fi, text="📤  Exportar Excel", command=lambda: exportar(app)).pack(side="right")
    ttk.Label(tab_ind, text="Pendientes por etapa (las etapas resaltadas son el seguimiento hasta el envío de la OC)",
              font=("Segoe UI", 10, "bold"), padding=(6, 4)).pack(fill="x")
    app.tree_pend = ttk.Treeview(tab_ind, columns=("etapa", "n", "rojo", "amarillo", "prom", "max", "pen", "usd"),
                                 show="headings", height=11)
    for k, t, w, a in [("etapa", "Etapa", 260, "w"), ("n", "N°", 60, "e"), ("rojo", "Atrasados", 80, "e"),
                       ("amarillo", "En atención", 85, "e"), ("prom", "Días promedio en etapa", 150, "e"),
                       ("max", "Máx. días", 80, "e"), ("pen", "Monto PEN", 130, "e"), ("usd", "Monto USD", 130, "e")]:
        app.tree_pend.heading(k, text=t)
        app.tree_pend.column(k, width=w, anchor=a, stretch=False)
    app.tree_pend.tag_configure("hasta_oc", background="#FFF2CC")
    app.tree_pend.tag_configure("total", font=("Segoe UI", 9, "bold"))
    app.tree_pend.pack(fill="x", padx=6)
    ttk.Label(tab_ind, text="Tiempos del flujo (días; solo requerimientos con ambas fechas)",
              font=("Segoe UI", 10, "bold"), padding=(6, 10, 6, 4)).pack(fill="x")
    app.tree_tiempos = ttk.Treeview(tab_ind, columns=("tramo", "n", "prom", "med", "max", "neg"), show="headings", height=9)
    for k, t, w, a in [("tramo", "Tramo", 320, "w"), ("n", "N°", 70, "e"), ("prom", "Promedio", 100, "e"),
                       ("med", "Mediana", 100, "e"), ("max", "Máximo", 100, "e"), ("neg", "Anticipos / regularizaciones", 200, "e")]:
        app.tree_tiempos.heading(k, text=t)
        app.tree_tiempos.column(k, width=w, anchor=a, stretch=False)
    app.tree_tiempos.tag_configure("total", font=("Segoe UI", 9, "bold"))
    app.tree_tiempos.pack(fill="x", padx=6)


def refrescar(app):
    app.filas_req = rq.listar(app.con)
    ctrl = ["Todos"] + sorted({f["controller"] for f in app.filas_req if f["controller"]})
    meses = sorted({f["f_solic"][:7] for f in app.filas_req if f["f_solic"]})
    for v in (app.vr_controller, app.vi_controller):
        v.widget.config(values=ctrl)
    app.vr_mes.widget.config(values=["Todos"] + meses[::-1])
    for v in (app.vi_desde, app.vi_hasta):
        v.widget.config(values=["Todos"] + meses)
    n_p, n_v = app.con.execute("SELECT (SELECT COUNT(*) FROM cat_proyectos), (SELECT COUNT(*) FROM cat_proveedores)").fetchone()
    app.lbl_cat.config(text=f"Catálogos: {n_p} proyectos · {n_v} proveedores")
    pintar_req(app)
    pintar_ind(app)


def filtradas_req(app):
    q = app.vr_buscar.get().strip().lower()
    et, tp, ct, mes = app.vr_etapa.get(), app.vr_tipo.get(), app.vr_controller.get(), app.vr_mes.get()
    u, plazo = sem.umbrales(app.cfg), app.vr_semaforo.get()
    out = []
    for f in app.filas_req:
        f["nivel"] = sem.nivel_req(f, u)
        if not sem.pasa_filtro(f["nivel"], plazo):
            continue
        if et == "Pendientes hasta envío de OC" and f["etapa"] not in rq.ETAPAS_HASTA_OC:
            continue
        if et == "Abiertos (todos)" and f["etapa"] in rq.CERRADAS:
            continue
        if et == "Cerrados" and f["etapa"] not in rq.CERRADAS:
            continue
        if et not in ("Pendientes hasta envío de OC", "Abiertos (todos)", "Cerrados", "Todos") and f["etapa"] != et:
            continue
        if tp != "Todos" and f["tipo"] != tp:
            continue
        if ct != "Todos" and f["controller"] != ct:
            continue
        if mes != "Todos" and not (f["f_solic"] or "").startswith(mes):
            continue
        if q and q not in " ".join(str(f.get(k) or "") for k in (
                "id", "proyecto", "proyecto_nombre", "referencia", "proveedor", "orden", "num_factura", "notas", "controller")).lower():
            continue
        out.append(f)
    return out


def pintar_req(app):
    filas = filtradas_req(app)
    clave, desc = app.orden_req
    num = {"id", "dias_etapa", "monto", "orden"}
    if clave == "plazo":
        orden_nivel = {sem.ROJO: 0, sem.AMARILLO: 1, sem.VERDE: 2}
        llave = lambda f: (orden_nivel.get(f["nivel"], 3), -(f["dias_etapa"] or 0))
    elif clave in num:
        llave = lambda f: (f[clave] is None, f[clave] or 0)
    else:
        llave = lambda f: str(f.get(clave) or "").lower()
    filas.sort(key=llave, reverse=desc)
    t = app.tree_req
    sel = t.selection()
    t.delete(*t.get_children())
    for f in filas:
        tag = ("pagado",) if f["etapa"] == "Pagado" else ("cerrado",) if f["etapa"] in rq.CERRADAS else \
              (f["nivel"],) if f["nivel"] in sem.FONDO else ()
        t.insert("", "end", iid=str(f["id"]), tags=tag, values=(
            f"{f['id']:03d}", rq.fmt(f["f_solic"]), f["proyecto"], f["referencia"], f["tipo"], f["proveedor"], f["etapa"],
            "" if f["dias_etapa"] is None else f["dias_etapa"], sem.texto(f["nivel"]),
            "" if f["monto"] is None else f"{f['monto']:,.2f}",
            f["moneda"], f["orden"] or "", f["sap_estado"], f["controller"], f["estatus"], f["com_oc"].replace("\n", " ")))
    if sel and t.exists(sel[0]):
        t.selection_set(sel[0])
    u = sem.umbrales(app.cfg)
    abiertos = [f for f in app.filas_req if f["etapa"] not in rq.CERRADAS]
    hasta = [f for f in abiertos if f["etapa"] in rq.ETAPAS_HASTA_OC]
    rojos = sum(1 for f in hasta if sem.nivel_req(f, u) == sem.ROJO)
    amar = sum(1 for f in hasta if sem.nivel_req(f, u) == sem.AMARILLO)
    app.lbl_resumen_req.config(text=f"Mostrando {len(filas)}   |   Abiertos: {len(abiertos)}   |   Hasta envío de OC: "
                                    f"{len(hasta)} (atrasados {rojos}, en atención {amar})   ·   colores: rojo = atrasado, "
                                    f"amarillo = atención (umbrales en «⚙ Plazos…»)")


def ordenar_req(app, clave):
    app.orden_req = (clave, not app.orden_req[1] if app.orden_req[0] == clave else False)
    pintar_req(app)


def pintar_ind(app):
    tp, d, h, ct = app.vi_tipo.get(), app.vi_desde.get(), app.vi_hasta.get(), app.vi_controller.get()
    filas = [f for f in app.filas_req
             if (tp == "Todos" or f["tipo"] == tp) and (ct == "Todos" or f["controller"] == ct)
             and (d == "Todos" or (f["f_solic"] or "")[:7] >= d) and (h == "Todos" or (f["f_solic"] or "9999")[:7] <= h)]
    pend, tiempos = rq.indicadores(filas)
    u = sem.umbrales(app.cfg)
    niveles = {}
    for f in filas:
        n = sem.nivel_req(f, u)
        if n in (sem.ROJO, sem.AMARILLO):
            niveles[(f["etapa"], n)] = niveles.get((f["etapa"], n), 0) + 1
    t = app.tree_pend
    t.delete(*t.get_children())
    n_oc = r_oc = a_oc = 0
    for p in pend:
        rojos, amar = niveles.get((p["etapa"], sem.ROJO), 0), niveles.get((p["etapa"], sem.AMARILLO), 0)
        if p["hasta_oc"]:
            n_oc, r_oc, a_oc = n_oc + p["n"], r_oc + rojos, a_oc + amar
        t.insert("", "end", tags=("hasta_oc",) if p["hasta_oc"] else (), values=(
            p["etapa"], p["n"], rojos or "", amar or "", "-" if p["prom"] is None else f"{p['prom']:.1f}",
            "-" if p["max"] is None else p["max"],
            f"{p['montos'].get('PEN', 0):,.2f}" if p["montos"].get("PEN") else "-",
            f"{p['montos'].get('USD', 0):,.2f}" if p["montos"].get("USD") else "-"))
    t.insert("", "end", tags=("total",), values=("Total pendientes hasta envío de OC", n_oc, r_oc or "", a_oc or "", "", "", "", ""))
    t = app.tree_tiempos
    t.delete(*t.get_children())
    for x in tiempos:
        t.insert("", "end", tags=("total",) if "total" in x["tramo"] else (), values=(
            x["tramo"], x["n"], "-" if x["prom"] is None else f"{x['prom']:.1f}",
            "-" if x["mediana"] is None else f"{x['mediana']:.1f}", "-" if x["max"] is None else x["max"],
            x["negativos"] or ""))


def editar(app):
    sel = app.tree_req.selection()
    if not sel:
        messagebox.showinfo("Editar", "Selecciona un requerimiento de la lista.", parent=app)
        return
    f = next(x for x in app.filas_req if x["id"] == int(sel[0]))
    EditorReq(app, f)


def _ruta_excel_req(app, preguntar=False):
    """Ruta del Excel de seguimiento (catálogos). Se pide si no hay una guardada o si `preguntar`."""
    ruta = app.cfg.get("excel_req")
    if preguntar or not ruta or not os.path.exists(ruta):
        ruta = filedialog.askopenfilename(parent=app, title="Excel «SEGUIMIENTO DE REQUERIMIENTOS Y ORDENES»",
                                          initialdir=os.path.dirname(ruta) if ruta else app_base(app),
                                          filetypes=[("Excel", "*.xlsx")])
    return ruta or None


def _recordar_ruta(app, ruta):
    """Se guarda la ruta solo cuando el Excel se pudo leer bien."""
    if app.cfg.get("excel_req") != ruta:
        app.cfg["excel_req"] = ruta
        app.guardar_cfg()


def app_base(app):
    return app.base


def catalogos(app):
    ruta = _ruta_excel_req(app)
    if not ruta:
        return
    try:
        n_p, n_v = rq.actualizar_catalogos(app.con, ruta)
    except Exception as e:  # noqa: BLE001
        app.con.rollback()
        messagebox.showerror("Catálogos", f"No se pudo leer el Excel (los catálogos anteriores se mantienen):\n{e}",
                             parent=app)
        return
    _recordar_ruta(app, ruta)
    refrescar(app)
    messagebox.showinfo("Catálogos", f"Catálogos actualizados desde\n{os.path.basename(ruta)}:\n"
                                     f"• {n_p} proyectos (Consulta1)\n• {n_v} proveedores (PROVEEDORES + SAP)", parent=app)


def importar(app):
    ruta = _ruta_excel_req(app, preguntar=True)
    if not ruta:
        return
    try:
        rq.actualizar_catalogos(app.con, ruta)
        n, avisos = rq.importar_control(app.con, ruta, app.cfg.get("autor", "Importación"))
        c = rq.completar_desde_sap(app.con)
    except Exception as e:  # noqa: BLE001
        app.con.rollback()
        messagebox.showerror("Importar", f"No se pudo importar (no se guardó nada):\n{e}", parent=app)
        return
    _recordar_ruta(app, ruta)
    refrescar(app)
    texto = (f"Se importaron {n} requerimientos nuevos de la tabla CONTROL (los COD que ya existían no se tocan).\n"
             f"{c} se completaron con datos de SAP.")
    if avisos:
        texto += f"\n\n{len(avisos)} fechas no se pudieron leer y quedaron vacías:\n• " + "\n• ".join(avisos[:10])
        texto += "\n…" if len(avisos) > 10 else ""
    messagebox.showinfo("Importar", texto, parent=app)


def exportar(app):
    u = sem.umbrales(app.cfg)
    plazos = {f["id"]: sem.texto(sem.nivel_req(f, u)) for f in app.filas_req}
    salida = rq.exportar_excel(app.con, os.path.join(app_base(app), f"Seguimiento requerimientos {date.today():%d_%m}.xlsx"),
                               plazos)
    if messagebox.askyesno("Exportar", f"Se generó {os.path.basename(salida)}.\n\n¿Abrirlo en Excel?", parent=app):
        os.startfile(salida)
