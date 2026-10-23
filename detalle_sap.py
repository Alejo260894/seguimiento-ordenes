"""
Reporte de detalle «Seguimiento OC → Entrada → Factura → Pago» (consulta SAP B1 / HANA, exportada a .xlsx).

Complementa al «Control de órdenes» (que sigue siendo la fuente general):
- Ingreso de material (nota de entrada) por línea de OC: estado logístico real de cada OC.
- Facturas con su N° interno (clave única) y el estado «Pago parcial».
- Líneas de la OC con proyecto, unidad de negocio y área.
Solo se guardan las OC aprobadas (con fecha de aprobación), igual que en el Control.

Granularidad del archivo: una fila = línea de OC × línea de entrada × línea de factura; por eso cada nivel se
deduplica por su propia clave antes de guardarlo.
"""
import os
import re
import shutil
import tempfile
import zipfile
from datetime import datetime

import pandas as pd

COLUMNAS = ["N° OC", "Estado OC", "Serie", "Fecha OC", "Aprobación", "Tipo OC", "Línea", "Código", "Descripción",
            "Cant. OC", "UM", "Total Línea Sin IGV", "Proyecto", "Unidad Negocio", "Área", "N° Entrada",
            "Fecha Entrada", "Cant. Ingresada", "Almacén", "Total Ingresado Línea", "Pendiente Ingreso",
            "Estado Logístico", "N° Factura", "Tipo", "N° Interno FT", "Emisión FT", "Contab. FT", "Venc. FT",
            "Estado Factura", "Fecha Pago", "Imp. Pendiente", "RUC", "Moneda"]

# Estado logístico de la OC (artículos), del más pendiente al más avanzado
SIN_DETALLE = "(sin detalle)"
SERVICIO = "Servicio"
LOGISTICA = ["Pendiente de ingreso", "Ingreso parcial", "Material ingresado", "Facturado sin ingreso",
             "Cerrada sin ingreso", "OC cancelada", SERVICIO]


def es_detalle(ruta):
    """¿El archivo es el reporte de detalle? (lee solo el encabezado)."""
    if not ruta.lower().endswith((".xlsx", ".xlsm")):
        return False
    try:
        cab = pd.read_excel(ruta, engine="calamine", nrows=0)
    except Exception:  # noqa: BLE001 — no es un Excel legible: que lo trate el lector del Control
        return False
    return "Estado Logístico" in cab.columns and "N° Interno FT" in cab.columns


def _limpiar_xlsx(ruta):
    """Copia del .xlsx sin caracteres de control dentro de los XML (SAP a veces los exporta y rompen la lectura)."""
    tmp = os.path.join(tempfile.gettempdir(), f"detalle_{os.getpid()}.xlsx")
    malos = re.compile(rb"[\x00-\x08\x0b\x0c\x0e-\x1f]")
    with zipfile.ZipFile(ruta) as z, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for item in z.infolist():
            datos = z.read(item.filename)
            if item.filename.endswith(".xml"):
                datos = malos.sub(b"", datos)
            out.writestr(item, datos)
    return tmp


def leer(ruta):
    """DataFrame del reporte (todo como texto). ValueError con un mensaje claro si no es el archivo esperado."""
    if ruta.lower().endswith(".csv"):
        raise ValueError("Del reporte de detalle usa el Excel (.xlsx): el CSV de SAP sale en UTF-16 y sin comillas, "
                         "y se desordenan las columnas.")
    copia = os.path.join(tempfile.gettempdir(), f"detalle_copia_{os.getpid()}.xlsx")   # puede estar abierto en Excel
    shutil.copyfile(ruta, copia)
    try:
        tmp = _limpiar_xlsx(copia)
    finally:
        os.remove(copia)
    try:
        df = pd.read_excel(tmp, engine="calamine", dtype=str).fillna("")
    finally:
        os.remove(tmp)
    faltan = [c for c in COLUMNAS if c not in df.columns]
    if faltan:
        raise ValueError("El archivo no parece el reporte de detalle «Seguimiento OC → Entrada → Factura → Pago»: "
                         f"faltan las columnas {', '.join(faltan[:6])}.")
    for c in df.columns:
        df[c] = df[c].str.strip()
    return df


def _iso(s):
    s = (s or "").strip()
    if not s:
        return ""
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, f).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return ""


def _num(s):
    try:
        return float(str(s).replace(",", ""))
    except ValueError:
        return 0.0


def _oc(s):
    try:
        return int(float(s))
    except ValueError:
        return None


def estado_oc(tipo_oc, estado, lineas):
    """Estado logístico de una OC a partir de sus líneas [(cantidad, ingresado, estado_linea)]."""
    if estado == "Cancelado":
        return "OC cancelada"
    if tipo_oc == SERVICIO:
        return SERVICIO
    cant = sum(c for c, _, _ in lineas)
    ing = sum(min(i, c) if c else i for c, i, _ in lineas)
    if ing <= 0:
        if any(e.startswith("Facturado sin ingreso") for _, _, e in lineas):
            return "Facturado sin ingreso"
        return "Cerrada sin ingreso" if estado == "Cerrado" else "Pendiente de ingreso"
    return "Material ingresado" if cant and ing >= cant - 1e-9 else "Ingreso parcial"


def guardar(con, df, archivo):
    """Reemplaza el detalle guardado con el del reporte (solo OC aprobadas). Devuelve un resumen."""
    df = df[df["Aprobación"] != ""].copy()
    df["oc"] = df["N° OC"].map(_oc)
    df = df[df["oc"].notna()]
    lineas = df.drop_duplicates(["oc", "Línea"])
    entradas = df[df["N° Entrada"] != ""].drop_duplicates(["oc", "Línea", "N° Entrada"])
    facturas = df[df["N° Interno FT"] != ""].drop_duplicates(["N° Interno FT", "oc"])
    fecha_ent = entradas.groupby(["oc", "Línea"])["Fecha Entrada"].agg(lambda s: max(_iso(x) for x in s))

    with con:
        for t in ("detalle_lineas", "detalle_entradas", "detalle_facturas", "detalle_oc"):
            con.execute(f"DELETE FROM {t}")
        con.executemany("INSERT INTO detalle_lineas VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", [
            (int(r["oc"]), int(_num(r["Línea"])), r["Tipo OC"], r["Código"], r["Descripción"], _num(r["Cant. OC"]),
             r["UM"], _num(r["Total Línea Sin IGV"]), r["Moneda"], r["Proyecto"], r["Unidad Negocio"], r["Área"],
             _num(r["Total Ingresado Línea"]), _num(r["Pendiente Ingreso"]), r["Estado Logístico"],
             fecha_ent.get((r["oc"], r["Línea"]), "")) for _, r in lineas.iterrows()])
        con.executemany("INSERT INTO detalle_entradas VALUES (?,?,?,?,?,?)", [
            (int(r["oc"]), int(_num(r["Línea"])), r["N° Entrada"], _iso(r["Fecha Entrada"]), _num(r["Cant. Ingresada"]),
             r["Almacén"]) for _, r in entradas.iterrows()])
        con.executemany("INSERT INTO detalle_facturas VALUES (?,?,?,?,?,?,?,?,?,?,?)", [
            (r["N° Interno FT"], int(r["oc"]), r["N° Factura"], r["Tipo"], r["Estado Factura"], _iso(r["Emisión FT"]),
             _iso(r["Contab. FT"]), _iso(r["Venc. FT"]), _iso(r["Fecha Pago"]), _num(r["Imp. Pendiente"]), r["RUC"])
            for _, r in facturas.iterrows()])
        filas = []
        for oc, g in lineas.groupby("oc"):
            c = g.iloc[0]
            datos = [(_num(x["Cant. OC"]), _num(x["Total Ingresado Línea"]), x["Estado Logístico"]) for _, x in g.iterrows()]
            est = estado_oc(c["Tipo OC"], c["Estado OC"], datos)
            cant = sum(d[0] for d in datos)
            ing = sum(min(d[1], d[0]) if d[0] else d[1] for d in datos)
            fechas = [fecha_ent.get((oc, x), "") for x in g["Línea"]]
            # proyecto / unidad / área de la línea de mayor monto
            mayor = g.assign(t=g["Total Línea Sin IGV"].map(_num)).sort_values("t", ascending=False).iloc[0]
            filas.append((int(oc), c["Tipo OC"], est,
                          None if c["Tipo OC"] == SERVICIO else int(ing > 0),
                          max(fechas) if any(fechas) else "", round(ing / cant, 4) if cant else None,
                          mayor["Proyecto"], mayor["Unidad Negocio"], mayor["Área"], len(g)))
        con.executemany("INSERT INTO detalle_oc VALUES (?,?,?,?,?,?,?,?,?,?)", filas)
        con.execute("INSERT INTO detalle_cargas(archivo, fecha_archivo, fecha_carga, filas, ocs) VALUES (?,?,?,?,?)",
                    (archivo, datetime.fromtimestamp(os.path.getmtime(archivo)).strftime("%Y-%m-%d %H:%M:%S")
                     if os.path.exists(archivo) else "", datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                     len(df), len(filas)))
        enlazar(con)
    return {"ocs": len(filas), "lineas": len(lineas), "entradas": len(entradas), "facturas": len(facturas)}


def enlazar(con):
    """Copia a sap_oc (último Control) el tipo de OC, el estado logístico y si tiene material ingresado."""
    con.execute("""UPDATE sap_oc SET tipo_oc = d.tipo_oc, estado_logistico = d.estado_logistico,
                   ingresado = d.ingresado, fecha_entrada = d.fecha_entrada
                   FROM detalle_oc d WHERE d.numero_oc = sap_oc.numero_oc""")


def ultima_carga(con):
    return con.execute("SELECT * FROM detalle_cargas ORDER BY id DESC LIMIT 1").fetchone()


def ingresos(con):
    """{N° OC: 1/0/None} — material ingresado (None = servicio). Para la regla de nota de entrada."""
    return {r[0]: r[1] for r in con.execute("SELECT numero_oc, ingresado FROM detalle_oc")}
