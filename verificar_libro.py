"""
Verifica en Excel (vía COM) el libro generado por procesar_sap.py y lo guarda ya calculado.

Uso:
    python verificar_libro.py salida.xlsx archivo_sap.(csv|xlsx)

1. Abre el libro en Excel, recalcula y busca errores (#REF!, #NAME?, #VALUE!, #N/A, ...).
2. Revisa que las celdas de control den 0.
3. Para varias combinaciones de filtros compara, contra un cálculo independiente en pandas,
   el monto por año y moneda, el Nº de órdenes, el Nº de facturas, las facturas por estado
   y el Top 15 de proveedores.
4. Deja los filtros en su valor por defecto y guarda el libro calculado por Excel.
"""
import os
import sys
from datetime import datetime

import pandas as pd
import win32com.client as win32

import procesar_sap as ps

ERRORES = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
           -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!", -2146826238: "#SPILL!",
           -2146826220: "#CALC!"}


def serial(d):
    return (d - datetime(1899, 12, 30)).days


# ---------------------------------------------------------------------------
# Modelo independiente en pandas
# ---------------------------------------------------------------------------
def esperado(df, f, moneda_graf):
    est_fact = df["Estado Factura"].where(df["Estado Factura"] != "", "Sin factura")
    m = est_fact.isin(f["ef"]) & df["Estado"].isin(f["eoc"]) & df["Usuario"].isin(f["usr"])
    if f["prov"]:
        m &= df["Proveedor"].str.upper().isin([p.upper() for p in f["prov"]])
    for col, key in [("Fecha", "f_oc"), ("Emisión FT", "f_em"), ("Fecha Pago", "f_pago")]:
        if f.get(key):
            d, h = f[key]
            d = d or datetime(1900, 1, 1)
            h = h or datetime(9999, 12, 31)
            m &= df[col].notna() & (df[col] >= d) & (df[col] <= h)
    p = df[m].copy()
    p["Año"] = p["Fecha"].dt.year
    oc = p.drop_duplicates("Numero OC", keep="first")          # primera fila de la OC que pasa
    monto = oc.groupby(["Año", "Mon"])["Total Sin IGV"].sum()
    ordenes = oc.groupby(["Año", "Mon"]).size()
    fx = p[p["N° Factura"] != ""]
    facturas = fx.groupby(["Año", "Mon"]).size()
    por_estado = fx.groupby(["Estado Factura", "Mon"]).size()
    g = oc[oc["Mon"] == moneda_graf].groupby(oc["Proveedor"].str.upper())["Total Sin IGV"].sum()
    top = g.sort_values(ascending=False, kind="stable").head(ps.TOP_N)
    return dict(monto=monto, ordenes=ordenes, facturas=facturas, por_estado=por_estado, top=top,
                total_graf=g.sum())


# ---------------------------------------------------------------------------
def leer_rango(ws, addr):
    v = ws.Range(addr).Value2
    if not isinstance(v, tuple):
        return [[v]]
    return [list(r) for r in v]


def buscar_errores(wb):
    encontrados = {}
    total = 0
    for ws in wb.Worksheets:
        datos = ws.UsedRange.Value2
        if not isinstance(datos, tuple):
            continue
        r0, c0 = ws.UsedRange.Row, ws.UsedRange.Column
        for i, fila in enumerate(datos):
            for j, v in enumerate(fila):
                if isinstance(v, int) and v in ERRORES:
                    total += 1
                    lst = encontrados.setdefault(f"{ws.Name}: {ERRORES[v]}", [])
                    if len(lst) < 10:
                        lst.append(ws.Cells(r0 + i, c0 + j).Address(False, False))
    return total, encontrados


def aplicar_filtros(rs, L, f):
    def si_no(rango_ini, opciones, elegidas):
        for k, op in enumerate(opciones):
            rs.Range(f"K{rango_ini + k}").Value2 = "Sí" if op in elegidas else "No"
    si_no(L.ef[0], ps.ESTADOS_FACT + ["Sin factura"], f["ef"])
    si_no(L.eoc[0], ps.ESTADOS_OC, f["eoc"])
    si_no(L.usr[0], L.usuarios, f["usr"])
    for k, key in enumerate(["f_oc", "f_em", "f_pago"]):
        r = L.fec[0] + k
        d, h = f.get(key) or (None, None)
        rs.Range(f"K{r}").Value2 = serial(d) if d else None
        rs.Range(f"L{r}").Value2 = serial(h) if h else None
        if not d:
            rs.Range(f"K{r}").ClearContents()
        if not h:
            rs.Range(f"L{r}").ClearContents()
    rs.Range(f"J{L.prov[0]}:J{L.prov[1]}").ClearContents()
    for k, p in enumerate(f["prov"]):
        rs.Range(f"J{L.prov[0] + k}").Value2 = p


def comparar(xl, wb, L, df, nombre, f, moneda_graf):
    rs = wb.Worksheets("Resumen Gastos")
    gp = wb.Worksheets("Gráficos Proveedores")
    aplicar_filtros(rs, L, f)
    gp.Range("B4").Value2 = moneda_graf
    xl.CalculateFull()
    e = esperado(df, f, moneda_graf)
    fallos = []

    def tabla(key, serie, decimales):
        t = L.tablas[key]
        vals = leer_rango(rs, f"B{t['ini']}:H{t['fin']}")
        filas = L.anios if key != "estado" else ps.ESTADOS_FACT
        for i, fila_key in enumerate(filas):
            for j, mon in enumerate(ps.MONEDAS):
                x = vals[i][j] or 0
                y = float(serie.get((fila_key, mon), 0))
                if round(x - y, decimales) != 0:
                    fallos.append(f"{key} [{fila_key},{mon}] Excel={x} pandas={y}")
    tabla("monto", e["monto"], 2)
    tabla("ordenes", e["ordenes"], 6)
    tabla("facturas", e["facturas"], 6)
    tabla("estado", e["por_estado"], 6)

    controles = {f"Resumen {k}": rs.Range(f"B{t['ctl']}").Value2 for k, t in L.tablas.items()}
    controles["Gráficos vs Resumen"] = gp.Range("C76").Value2
    for k, v in controles.items():
        if v is None or abs(v) > 0.005:
            fallos.append(f"Control {k} = {v}")

    top_x = [(str(r[0]), r[1] or 0) for r in leer_rango(gp, "B58:C72")]
    for i, (prov, monto) in enumerate(e["top"].items()):
        nx, mx = top_x[i]
        if round(mx - monto, 2) != 0 or nx.upper() != prov[:ps.NOMBRE_MAX]:
            fallos.append(f"Top {i + 1}: Excel=({nx}, {mx}) pandas=({prov[:ps.NOMBRE_MAX]}, {monto})")
    tot_graf = gp.Range("C74").Value2
    if round(tot_graf - e["total_graf"], 2) != 0:
        fallos.append(f"Total gráfico Excel={tot_graf} pandas={e['total_graf']}")

    t = L.tablas
    resumen = (f"monto total por moneda={[round(v or 0, 2) for v in leer_rango(rs, f'B{t['monto']['tot']}:H{t['monto']['tot']}')[0]]}, "
               f"órdenes={sum(v or 0 for v in leer_rango(rs, f'B{t['ordenes']['tot']}:H{t['ordenes']['tot']}')[0]):.0f}, "
               f"facturas={sum(v or 0 for v in leer_rango(rs, f'B{t['facturas']['tot']}:H{t['facturas']['tot']}')[0]):.0f}, "
               f"top1 {moneda_graf}={top_x[0]}")
    print(f"\n[{nombre}] {'OK' if not fallos else 'FALLOS: ' + str(len(fallos))}")
    print("   " + resumen)
    for x in fallos[:20]:
        print("   - " + x)
    return not fallos


def main():
    salida, entrada = os.path.abspath(sys.argv[1]), sys.argv[2]
    df, log, meta, _ = ps.procesar(entrada)
    usuarios = sorted(df["Usuario"].unique())
    anios = sorted(int(y) for y in df["Fecha"].dropna().map(lambda d: d.year).unique())
    L = ps.Layout(usuarios, anios)

    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    ok = True
    try:
        wb = xl.Workbooks.Open(salida)
        xl.CalculateFull()
        total, errs = buscar_errores(wb)
        print(f"Errores de fórmula en el libro: {total}")
        for k, v in errs.items():
            print(f"   {k}: {v}")
        ok &= total == 0

        rs = wb.Worksheets("Resumen Gastos")
        dv = rs.Range(f"J{L.prov[0]}").Validation.Formula1
        print(f"Validación proveedor J{L.prov[0]}: {dv}")
        n_lista = wb.Worksheets("Listas").Range("A2").SpillingToRange.Rows.Count
        print(f"Listas!A2# derrama {n_lista} proveedores")

        todos = dict(ef=ps.ESTADOS_FACT + ["Sin factura"], eoc=ps.ESTADOS_OC, usr=usuarios, prov=[])
        prov_lista = [r[0] for r in leer_rango(wb.Worksheets("Listas"), f"A2:A{1 + n_lista}")]
        escenarios = [
            ("Todo en Sí, sin fechas, USD", todos, "USD"),
            ("Todo en Sí, sin fechas, SOL", todos, "SOL"),
            ("Solo Pendiente de pago · OC Cerrado+En proceso · 1 comprador", dict(ef=["Pendiente de pago"], eoc=["Cerrado", "En proceso"], usr=usuarios[:1], prov=[]), "SOL"),
            ("Pagado+Sin factura · Fecha OC 01/03/2026–30/06/2026 · pago desde 01/04/2026", dict(ef=["Pagado", "Sin factura"], eoc=ps.ESTADOS_OC, usr=usuarios, prov=[], f_oc=(datetime(2026, 3, 1), datetime(2026, 6, 30)), f_pago=(datetime(2026, 4, 1), None)), "USD"),
            ("3 proveedores · emisión FT hasta 31/12/2025", dict(ef=ps.ESTADOS_FACT + ["Sin factura"], eoc=ps.ESTADOS_OC, usr=usuarios, prov=prov_lista[10:13], f_em=(None, datetime(2025, 12, 31))), "SOL"),
            ("Solo Sin factura · Cancelado", dict(ef=["Sin factura"], eoc=["Cancelado"], usr=usuarios, prov=[]), "USD"),
        ]
        for nombre, f, mon in escenarios:
            ok &= comparar(xl, wb, L, df, nombre, f, mon)

        aplicar_filtros(rs, L, todos)
        wb.Worksheets("Gráficos Proveedores").Range("B4").Value2 = "USD"
        xl.CalculateFull()
        total, _ = buscar_errores(wb)
        print(f"\nErrores tras restaurar filtros: {total}")
        rs.Activate()
        rs.Range("A1").Select()
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()
    print("\nRESULTADO:", "TODO OK" if ok else "HAY DIFERENCIAS")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
