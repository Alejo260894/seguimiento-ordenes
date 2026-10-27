"""
«Resumen seguimiento DD_MM.xlsx»: resumen ejecutivo de requerimientos y OC según SAP, en un archivo aparte.

Hojas:
- Panel: filtros (celdas amarillas), indicadores, 4 gráficos, ficha para explicar una OC o requerimiento y el
  listado filtrado. Todo se recalcula al cambiar un filtro (fórmulas de Excel 365: FILTER, SORTBY, UNIQUE, XMATCH).
- Datos: una fila por requerimiento (sin vuelos) con lo que dice SAP, su STATUS y el detalle. La columna VISIBLE
  indica si la fila pasa los filtros del Panel.
- Calculos: tablas que alimentan los gráficos.
- Listas: opciones de los filtros.
"""
import os
from datetime import date, datetime

import requerimientos as rq

ESTADOS = rq.STATUS_RESUMEN
PENDIENTES = ["Pendiente de pago", "Facturación parcial", "Pendiente de factura", "Sin OC (en trámite)", "OC no está en SAP"]
TODOS_PEND = "Pendientes (todo lo no completado)"
ANTIGUEDAD = [("0 a 30 días", 0, 30), ("31 a 60 días", 31, 60), ("61 a 90 días", 61, 90), ("Más de 90 días", 91, 100000)]
COLORES = {"Completado": "#16A34A", "Pendiente de pago": "#E11D48", "Facturación parcial": "#F59E0B",
           "Pendiente de factura": "#EA580C", "Sin OC (en trámite)": "#60A5FA", "OC no está en SAP": "#94A3B8",
           "Anulada": "#64748B"}
FONDOS = {"Completado": "#DCFCE7", "Pendiente de pago": "#FFE4E6", "Facturación parcial": "#FEF3C7",
          "Pendiente de factura": "#FFEDD5", "Sin OC (en trámite)": "#DBEAFE", "OC no está en SAP": "#F1F5F9",
          "Anulada": "#F1F5F9"}
AZUL = "#1E3A8A"

# Columnas de la hoja Datos: (encabezado, clave en la fila del resumen, tipo)
COLUMNAS = [("COD", "id", "int"), ("TIPO", "tipo", "txt"), ("PROYECTO", "proyecto", "txt"),
            ("REFERENCIA", "referencia", "txt"), ("PROVEEDOR", "proveedor", "txt"), ("COMPRADOR", "comprador", "txt"),
            ("ORDEN", "orden", "int"), ("FECHA OC", "fecha_oc", "fecha"), ("DÍAS", "dias", "int"),
            ("ESTADO OC", "estado_oc", "txt"), ("MONEDA", "moneda", "txt"), ("MONTO OC", "total", "num"),
            ("FACTURADO", "facturado", "num"), ("% FACTURADO", "pct", "pct"), ("SALDO POR PAGAR", "saldo", "num"),
            ("MONTO PENDIENTE", "pendiente", "num"), ("PRÓX. VENCIMIENTO", "vence", "fecha"),
            ("ÚLTIMO PAGO", "ultimo_pago", "fecha"), ("STATUS", "status", "txt"), ("DETALLE", "detalle", "txt"),
            ("NOTAS SOLPED", "notas", "txt"), ("SEGUIMIENTO OC", "seguimiento", "txt"),
            ("ÚLTIMO COMENTARIO", "ult_com", "txt"), ("COMENTARIOS", "comentarios", "txt"),
            ("PENDIENTE", "es_pendiente", "int"), ("PRIORIDAD", "prioridad", "int")]
POS = {c[0]: i for i, c in enumerate(COLUMNAS)}
LISTADO = ["COD", "ORDEN", "PROVEEDOR", "STATUS", "MONEDA", "MONTO PENDIENTE", "DÍAS", "PRÓX. VENCIMIENTO",
           "ÚLTIMO COMENTARIO", "DETALLE"]
FICHA = ["COD", "TIPO", "PROYECTO", "REFERENCIA", "PROVEEDOR", "COMPRADOR", "ORDEN", "FECHA OC", "DÍAS", "ESTADO OC",
         "MONEDA", "MONTO OC", "FACTURADO", "% FACTURADO", "SALDO POR PAGAR", "PRÓX. VENCIMIENTO", "ÚLTIMO PAGO",
         "STATUS", "DETALLE", "SEGUIMIENTO OC", "NOTAS SOLPED", "ÚLTIMO COMENTARIO", "COMENTARIOS"]


def _fecha(iso):
    return datetime.strptime(iso, "%Y-%m-%d") if iso else None


def filas_resumen(con, comprador=None):
    """Filas del resumen (sin vuelos) con días de antigüedad, prioridad y marca de pendiente."""
    reqs = [f for f in rq.listar(con) if comprador is None or f["comprador"] == comprador]
    orden = {s: i for i, s in enumerate(PENDIENTES + ["Completado", "Anulada"])}     # lo pendiente primero
    notas = {f["id"]: f["notas"] for f in reqs}
    # Comentarios de la OC (se hacen en «OC sin factura», «Facturas» o «Consultar OC»): más reciente primero
    comentarios = {}
    for c in con.execute("SELECT numero_oc, fecha, autor, texto FROM comentarios WHERE tipo='usuario' "
                         "ORDER BY fecha DESC, id DESC"):
        comentarios.setdefault(c["numero_oc"], []).append(
            f"{c['fecha'][8:10]}/{c['fecha'][5:7]}/{c['fecha'][:4]} {c['autor']}: {' '.join(c['texto'].split())}")
    seguimiento = {}
    for o in con.execute("SELECT * FROM ordenes WHERE situacion='Activa'"):
        partes = [o["estado_seg"]]
        if o["responsable"]:
            partes.append(f"responsable {o['responsable']}")
        if o["fecha_compromiso"]:
            partes.append(f"compromiso {o['fecha_compromiso'][8:10]}/{o['fecha_compromiso'][5:7]}/{o['fecha_compromiso'][:4]}")
        seguimiento[o["numero_oc"]] = " · ".join(partes)
    salida = []
    for r in rq.resumen_sap(con, reqs):
        com = comentarios.get(r["orden"], []) if r["orden"] else []
        r["notas"] = " ".join((notas.get(r["id"]) or "").split())
        r["ult_com"] = com[0] if com else ""
        r["comentarios"] = " | ".join(com)
        r["seguimiento"] = seguimiento.get(r["orden"], "") if r["orden"] else ""
        r["es_pendiente"] = int(r["status"] in PENDIENTES)
        r["prioridad"] = orden.get(r["status"], 9) + 1
        r["dias"] = ((date.today() - _fecha(r["fecha_oc"]).date()).days
                     if r.get("fecha_oc") and r["es_pendiente"] else None)
        salida.append(r)
    return salida


def exportar(con, salida, comprador=None):
    """Escribe el archivo de resumen. Si `salida` está abierto en Excel, guarda una copia con la hora."""
    import xlsxwriter
    filas = filas_resumen(con, comprador)
    destino, tmp = salida, os.path.join(os.path.dirname(salida), f"~tmp_resumen_{os.getpid()}.xlsx")
    wb = xlsxwriter.Workbook(tmp, {"strings_to_formulas": False,
                                   "default_format_properties": {"font_name": "Segoe UI", "font_size": 10}})
    F = lambda **k: wb.add_format({"font_name": "Segoe UI", "font_size": 10, **k})
    hdr = F(bold=True, font_color="white", bg_color=AZUL, border=1, border_color="#CBD5E1", valign="vcenter", text_wrap=True)
    fd, fm, fp, fe = F(num_format="dd/mm/yyyy"), F(num_format="#,##0.00"), F(num_format="0%"), F(num_format="0")

    panel = wb.add_worksheet("Panel")
    datos = wb.add_worksheet("Datos")
    calc = wb.add_worksheet("Calculos")
    listas = wb.add_worksheet("Listas")

    # ---------------------------------------------------------------- Listas (opciones de los filtros)
    compradores = sorted({f["comprador"] for f in filas if f["comprador"]})
    proyectos = sorted({f["proyecto"] for f in filas if f["proyecto"]})
    monedas = sorted({f["moneda"] for f in filas if f["moneda"]})
    opciones = {"Status": ["Todos", TODOS_PEND] + ESTADOS, "Moneda": ["Todas"] + monedas,
                "Comprador": ["Todos"] + compradores, "Proyecto": ["Todos"] + proyectos,
                "Antigüedad": ["Todos"] + [a[0] for a in ANTIGUEDAD]}
    for j, (nombre, valores) in enumerate(opciones.items()):
        listas.write(0, j, nombre, hdr)
        listas.write_column(1, j, valores)
    listas.hide()

    # ---------------------------------------------------------------- Datos (tabla con VISIBLE según los filtros)
    def valor(f, clave, tipo):
        v = f.get(clave)
        if tipo == "fecha":
            return _fecha(v)
        if v in ("", None):
            return None
        return v
    cuerpo = [[valor(f, c, t) for _, c, t in COLUMNAS] + [None] for f in filas]
    visible = ("=IF(AND(OR(Panel!$B$5=\"Todos\",AND(Panel!$B$5=\"" + TODOS_PEND + "\",[@PENDIENTE]=1),[@STATUS]=Panel!$B$5),"
               "OR(Panel!$C$5=\"Todas\",[@MONEDA]=Panel!$C$5),OR(Panel!$D$5=\"Todos\",[@COMPRADOR]=Panel!$D$5),"
               "OR(Panel!$E$5=\"Todos\",[@PROYECTO]=Panel!$E$5),"
               "OR(Panel!$F$5=\"Todos\",AND(ISNUMBER([@DÍAS]),OR("
               + ",".join(f"AND(Panel!$F$5=\"{n}\",[@DÍAS]>={lo},[@DÍAS]<={hi})" for n, lo, hi in ANTIGUEDAD) + "))),"
               "OR(Panel!$G$5=\"\",ISNUMBER(SEARCH(Panel!$G$5,[@COD]&\" \"&[@ORDEN]&\" \"&[@PROVEEDOR]&\" \"&[@REFERENCIA]"
               "&\" \"&[@PROYECTO]&\" \"&[@DETALLE]&\" \"&[@[NOTAS SOLPED]]&\" \"&[@COMENTARIOS])))),1,0)")
    formatos = {"fecha": fd, "num": fm, "pct": fp}
    cols = [{"header": h, "header_format": hdr, **({"format": formatos[t]} if t in formatos else {})} for h, _, t in COLUMNAS]
    cols.append({"header": "VISIBLE", "header_format": hdr, "formula": visible})
    n = max(len(cuerpo), 1)
    datos.add_table(0, 0, n, len(cols) - 1, {"name": "Datos", "data": cuerpo, "columns": cols,
                                             "style": "Table Style Light 9"})
    anchos = {"PROVEEDOR": 36, "DETALLE": 70, "REFERENCIA": 18, "COMPRADOR": 18, "STATUS": 22, "NOTAS SOLPED": 30,
              "SEGUIMIENTO OC": 30, "ÚLTIMO COMENTARIO": 50, "COMENTARIOS": 80}
    for i, (h, _, t) in enumerate(COLUMNAS):
        datos.set_column(i, i, anchos.get(h, 12 if t in ("fecha", "num") else 10))
    datos.freeze_panes(1, 1)
    for st, fondo in FONDOS.items():
        datos.conditional_format(1, POS["STATUS"], n, POS["STATUS"],
                                 {"type": "cell", "criteria": "==", "value": f'"{st}"', "format": F(bg_color=fondo, bold=True)})

    # ---------------------------------------------------------------- Calculos (datos de los gráficos)
    calc.write_row(0, 0, ["Status", "N°", "Monto PEN", "Monto USD", "Otras monedas"], hdr)
    for i, st in enumerate(ESTADOS, start=1):
        r = i + 1
        calc.write(i, 0, st)
        calc.write_formula(i, 1, f'=COUNTIFS(Datos[STATUS],$A{r},Datos[VISIBLE],1)')
        calc.write_formula(i, 2, f'=SUMIFS(Datos[MONTO PENDIENTE],Datos[STATUS],$A{r},Datos[MONEDA],"PEN",Datos[VISIBLE],1)', fm)
        calc.write_formula(i, 3, f'=SUMIFS(Datos[MONTO PENDIENTE],Datos[STATUS],$A{r},Datos[MONEDA],"USD",Datos[VISIBLE],1)', fm)
        calc.write_formula(i, 4, f'=SUMIFS(Datos[MONTO PENDIENTE],Datos[STATUS],$A{r},Datos[VISIBLE],1)-C{r}-D{r}', fm)
    fila_ant = len(ESTADOS) + 3
    calc.write_row(fila_ant - 1, 0, ["Antigüedad (pendientes)", "N°"], hdr)
    for k, (nombre, lo, hi) in enumerate(ANTIGUEDAD):
        calc.write(fila_ant + k, 0, nombre)
        calc.write_formula(fila_ant + k, 1, f'=COUNTIFS(Datos[PENDIENTE],1,Datos[VISIBLE],1,Datos[DÍAS],">={lo}",Datos[DÍAS],"<={hi}")')
    fila_prov = fila_ant + len(ANTIGUEDAD) + 2
    calc.write_row(fila_prov - 1, 0, ["Proveedores con pendientes", "N°", "", "Top 10", "N°"], hdr)
    calc.write_dynamic_array_formula(fila_prov, 0, fila_prov, 0,
                                     '=_xlfn.UNIQUE(_xlfn._xlws.FILTER(Datos[PROVEEDOR],(Datos[VISIBLE]=1)*(Datos[PENDIENTE]=1),"(sin pendientes)"))')
    a = f"_xlfn.ANCHORARRAY(Calculos!$A${fila_prov + 1})"
    calc.write_dynamic_array_formula(fila_prov, 1, fila_prov, 1,
                                     f'=COUNTIFS(Datos[PROVEEDOR],{a},Datos[VISIBLE],1,Datos[PENDIENTE],1)')
    b = f"_xlfn.ANCHORARRAY(Calculos!$B${fila_prov + 1})"
    calc.write_dynamic_array_formula(fila_prov, 3, fila_prov + 9, 3,
                                     f'=IFERROR(INDEX(_xlfn.SORTBY({a},{b},-1),_xlfn.SEQUENCE(10)),"")')
    calc.write_dynamic_array_formula(fila_prov, 4, fila_prov + 9, 4,
                                     f'=IFERROR(INDEX(_xlfn.SORTBY({b},{b},-1),_xlfn.SEQUENCE(10)),"")')
    calc.set_column(0, 0, 40)
    calc.set_column(1, 4, 14)
    calc.set_column(3, 3, 40)

    # ---------------------------------------------------------------- Panel
    panel.hide_gridlines(2)
    panel.set_column(0, 0, 2)
    panel.set_column(1, 8, 17)
    panel.set_column(9, 9, 2)
    panel.set_column(10, 10, 40)
    panel.set_column(11, 11, 62)
    banda = F(bold=True, font_size=18, font_color="white", bg_color=AZUL, valign="vcenter")
    panel.set_row(0, 34)
    panel.merge_range(0, 0, 0, 11, "  Resumen de seguimiento de órdenes y requerimientos (SAP)", banda)
    total_fil = len(filas)
    panel.merge_range(1, 0, 1, 11, f"  {('Comprador: ' + comprador + ' · ') if comprador else ''}{total_fil} requerimientos "
                                   f"(sin vuelos) · generado el {datetime.now():%d/%m/%Y %H:%M} · «Completado» = sin "
                                   f"pendientes de pago (saldos de hasta {rq.SALDO_MENOR:,.0f} = diferencia de montos)",
                      F(font_color="#475569", italic=True))
    etiqueta = F(bold=True, font_color=AZUL)
    filtro = F(bg_color="#FEF9C3", border=1, border_color="#CA8A04", bold=True)
    for j, nombre in enumerate(["Status", "Moneda", "Comprador", "Proyecto", "Antigüedad"]):
        panel.write(3, 1 + j, nombre, etiqueta)
        panel.write(4, 1 + j, opciones[nombre][0], filtro)
        col = chr(ord("A") + j)
        panel.data_validation(4, 1 + j, 4, 1 + j, {
            "validate": "list", "source": f"=Listas!${col}$2:${col}${len(opciones[nombre]) + 1}"})
    panel.write(3, 6, "Buscar (OC, COD, proveedor, texto)", etiqueta)
    panel.write_blank(4, 6, None, filtro)
    panel.write_comment(4, 1, "Elige un status, «Pendientes (todo lo no completado)» o «Todos».")

    # Indicadores
    kpis = [("Requerimientos", "=SUM(Datos[VISIBLE])", fe, "#2563EB"),
            ("Completados", '=COUNTIFS(Datos[VISIBLE],1,Datos[STATUS],"Completado")', fe, "#16A34A"),
            ("% completado", '=IFERROR(C8/SUM(Datos[VISIBLE]),0)', fp, "#16A34A"),
            ("Pendientes", '=SUMIFS(Datos[PENDIENTE],Datos[VISIBLE],1)', fe, "#EA580C"),
            ("Por pagar PEN", '=SUMIFS(Datos[SALDO POR PAGAR],Datos[VISIBLE],1,Datos[STATUS],"Pendiente de pago",Datos[MONEDA],"PEN")', fm, "#E11D48"),
            ("Por pagar USD", '=SUMIFS(Datos[SALDO POR PAGAR],Datos[VISIBLE],1,Datos[STATUS],"Pendiente de pago",Datos[MONEDA],"USD")', fm, "#E11D48"),
            ("Facturas vencidas", '=COUNTIFS(Datos[VISIBLE],1,Datos[STATUS],"Pendiente de pago",Datos[PRÓX. VENCIMIENTO],"<"&TODAY())', fe, "#B91C1C"),
            ("Sin factura +30 días", '=COUNTIFS(Datos[VISIBLE],1,Datos[STATUS],"Pendiente de factura",Datos[DÍAS],">30")', fe, "#EA580C")]
    for j, (nombre, formula, fmt, color) in enumerate(kpis):
        panel.write(6, 1 + j, nombre, F(bold=True, font_color="white", bg_color=color, align="center"))
        fmt_v = wb.add_format({"font_name": "Segoe UI", "font_size": 18, "bold": True, "font_color": color,
                               "align": "center", "bg_color": "#F8FAFC", "border": 1, "border_color": "#E2E8F0",
                               "num_format": {fe: "0", fm: "#,##0", fp: "0%"}[fmt]})
        panel.write_formula(7, 1 + j, formula, fmt_v)
    panel.set_row(7, 32)

    # Gráficos
    r1 = len(ESTADOS) + 1
    g1 = wb.add_chart({"type": "bar"})
    g1.add_series({"name": "Requerimientos", "categories": ["Calculos", 1, 0, r1 - 1, 0],
                   "values": ["Calculos", 1, 1, r1 - 1, 1], "data_labels": {"value": True},
                   "points": [{"fill": {"color": COLORES[s]}} for s in ESTADOS]})
    g1.set_title({"name": "Requerimientos por status", "name_font": {"size": 12, "color": AZUL}})
    g1.set_legend({"none": True})
    g1.set_y_axis({"reverse": True})
    g2 = wb.add_chart({"type": "column"})
    g2.add_series({"name": "Pendientes", "categories": ["Calculos", fila_ant, 0, fila_ant + 3, 0],
                   "values": ["Calculos", fila_ant, 1, fila_ant + 3, 1], "data_labels": {"value": True},
                   "points": [{"fill": {"color": c}} for c in ("#16A34A", "#F59E0B", "#EA580C", "#E11D48")]})
    g2.set_title({"name": "Pendientes por antigüedad (días desde la OC)", "name_font": {"size": 12, "color": AZUL}})
    g2.set_legend({"none": True})
    g3 = wb.add_chart({"type": "bar"})
    for col, nombre, color in ((2, "PEN", "#2563EB"), (3, "USD", "#0EA5A4")):
        g3.add_series({"name": nombre, "categories": ["Calculos", 2, 0, 4, 0], "values": ["Calculos", 2, col, 4, col],
                       "fill": {"color": color}, "data_labels": {"value": True, "num_format": "#,##0"}})
    g3.set_title({"name": "Monto pendiente por status (por pagar / por facturar)", "name_font": {"size": 12, "color": AZUL}})
    g3.set_y_axis({"reverse": True})
    g3.set_x_axis({"num_format": "#,##0"})
    g3.set_legend({"position": "bottom"})
    g4 = wb.add_chart({"type": "bar"})
    g4.add_series({"name": "Pendientes", "categories": ["Calculos", fila_prov, 3, fila_prov + 9, 3],
                   "values": ["Calculos", fila_prov, 4, fila_prov + 9, 4], "fill": {"color": "#7C3AED"},
                   "data_labels": {"value": True}})
    g4.set_title({"name": "Proveedores con más pendientes (top 10)", "name_font": {"size": 12, "color": AZUL}})
    g4.set_legend({"none": True})
    g4.set_y_axis({"reverse": True})
    for g, celda in ((g1, "B10"), (g2, "F10"), (g3, "B28"), (g4, "F28")):
        g.set_size({"width": 520, "height": 340})
        g.set_chartarea({"border": {"color": "#E2E8F0"}})
        panel.insert_chart(celda, g, {"x_offset": 4, "y_offset": 4})

    # Ficha: explicar una OC o un requerimiento
    panel.merge_range(3, 10, 3, 11, "EXPLICAR UNA ORDEN O REQUERIMIENTO", F(bold=True, font_color="white", bg_color=AZUL))
    panel.write(4, 10, "N° de OC o COD", etiqueta)
    panel.write_blank(4, 11, None, filtro)
    panel.write_comment(4, 11, "Escribe el N° de OC (ej. 262000968) o el COD del requerimiento (ej. 93).")
    fila_match = ('IFERROR(_xlfn.XMATCH($L$5,Datos[ORDEN]),IFERROR(_xlfn.XMATCH($L$5,Datos[COD]),'
                  'IFERROR(_xlfn.XMATCH(VALUE($L$5),Datos[ORDEN]),_xlfn.XMATCH(VALUE($L$5),Datos[COD]))))')
    panel.write(5, 10, "Explicación", etiqueta)
    explic = ('=IF($L$5="","Escribe un N° de OC o un COD en la celda amarilla.",IFERROR('
              '"COD "&INDEX(Datos[COD],' + fila_match + ')&IF(INDEX(Datos[ORDEN],' + fila_match + ')="",""," · OC "&'
              'INDEX(Datos[ORDEN],' + fila_match + '))&" · "&INDEX(Datos[PROVEEDOR],' + fila_match + ')&": "&'
              'UPPER(INDEX(Datos[STATUS],' + fila_match + '))&". "&INDEX(Datos[DETALLE],' + fila_match + ')&'
              'IF(INDEX(Datos[ÚLTIMO COMENTARIO],' + fila_match + ')="",""," · Último comentario: "&'
              'INDEX(Datos[ÚLTIMO COMENTARIO],' + fila_match + '))&'
              'IF(INDEX(Datos[NOTAS SOLPED],' + fila_match + ')="",""," · Nota SOLPED: "&INDEX(Datos[NOTAS SOLPED],' + fila_match + ')),'
              '"No se encontró: revisa el número (las OC sin aprobación y los vuelos no están en el resumen)."))')
    panel.write_formula(5, 11, explic, F(text_wrap=True, valign="top", bold=True, font_color="#0F172A"))
    panel.set_row(5, 64)
    formatos_ficha = {"FECHA OC": fd, "PRÓX. VENCIMIENTO": fd, "ÚLTIMO PAGO": fd, "MONTO OC": fm, "FACTURADO": fm,
                      "SALDO POR PAGAR": fm, "% FACTURADO": fp}
    for k, campo in enumerate(FICHA):
        r = 6 + k
        panel.write(r, 10, campo, F(font_color="#475569"))
        panel.write_formula(r, 11, f'=IF($L$5="","",IFERROR(INDEX(Datos[{campo}],{fila_match})&"","—"))'
                            if campo not in formatos_ficha else
                            f'=IF($L$5="","",IFERROR(INDEX(Datos[{campo}],{fila_match}),"—"))',
                            F(**({"num_format": formatos_ficha[campo].num_format} if campo in formatos_ficha else {}),
                              text_wrap=campo == "DETALLE", align="left"))
    for st, fondo in FONDOS.items():
        panel.conditional_format(6 + FICHA.index("STATUS"), 11, 6 + FICHA.index("STATUS"), 11,
                                 {"type": "cell", "criteria": "==", "value": f'"{st}"', "format": F(bg_color=fondo, bold=True)})

    # Listado filtrado (prioridad y luego monto pendiente)
    fila_lis = 46
    panel.merge_range(fila_lis - 2, 1, fila_lis - 2, 11, "LISTADO SEGÚN LOS FILTROS (prioridad: pendiente de pago, "
                      "facturación parcial, pendiente de factura…; luego mayor monto pendiente)",
                      F(bold=True, font_color="white", bg_color=AZUL))
    columnas_lis = [POS[c] + 1 for c in LISTADO]
    for j, c in enumerate(LISTADO):
        panel.write(fila_lis - 1, 1 + j if j < 8 else 2 + j, c, hdr)
    vis = "Datos[VISIBLE]=1"
    # En el archivo la tabla completa se escribe Datos[#Data] (Excel la muestra como «Datos»)
    sel = f'INDEX(Datos[#Data],_xlfn.SEQUENCE(ROWS(Datos[#Data])),{{{",".join(str(c) for c in columnas_lis[:8])}}})'
    formula = (f'=IFERROR(_xlfn.SORTBY(_xlfn._xlws.FILTER({sel},{vis}),_xlfn._xlws.FILTER(Datos[PRIORIDAD],{vis}),1,'
               f'_xlfn._xlws.FILTER(Datos[MONTO PENDIENTE],{vis}),-1),"Sin resultados con estos filtros")')
    panel.write_dynamic_array_formula(fila_lis, 1, fila_lis, 1, formula)
    # Último comentario (columna K) y detalle (columna L), en el mismo orden que el listado
    for col, campo in ((10, "ÚLTIMO COMENTARIO"), (11, "DETALLE")):
        f_txt = (f'=IFERROR(_xlfn.SORTBY(_xlfn._xlws.FILTER(Datos[{campo}],{vis}),_xlfn._xlws.FILTER(Datos[PRIORIDAD],{vis}),1,'
                 f'_xlfn._xlws.FILTER(Datos[MONTO PENDIENTE],{vis}),-1),"")')
        panel.write_dynamic_array_formula(fila_lis, col, fila_lis, col, f_txt)
    largo = fila_lis + len(filas) + 2
    for j, c in enumerate(LISTADO[:8]):
        fmt = {"MONTO PENDIENTE": fm, "PRÓX. VENCIMIENTO": fd}.get(c)
        if fmt:
            for r in range(fila_lis, largo):
                panel.write_blank(r, 1 + j, None, fmt)
    for st, fondo in FONDOS.items():
        panel.conditional_format(fila_lis, 4, largo, 4, {"type": "cell", "criteria": "==", "value": f'"{st}"',
                                                        "format": F(bg_color=fondo, bold=True)})
    panel.freeze_panes(5, 0)
    panel.activate()
    wb.close()
    try:
        os.replace(tmp, destino)
    except PermissionError:          # el archivo del día está abierto en Excel
        destino = destino.replace(".xlsx", f" ({datetime.now():%H%M%S}).xlsx")
        os.replace(tmp, destino)
    return destino
