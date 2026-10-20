"""
Procesa el reporte "Control de órdenes" descargado de SAP Business One (.csv o .xlsx)
y genera un libro Excel con: Sheet1 (datos limpios), Resumen Gastos, Gráficos Proveedores,
Calc_Graficos, Listas y Log limpieza.

Uso:
    python procesar_sap.py archivo_sap.(csv|xlsx) salida.xlsx
"""
import csv
import re
import sys
from datetime import datetime, time

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
HEADERS = ["Usuario", "Estado", "Serie", "Preliminar", "Numero OC", "Fecha", "Aprobación",
           "RUC", "Proveedor", "Mon", "Total Sin IGV", "N° Factura", "Tipo", "Estado Factura",
           "FT Sin Imp.", "Emisión FT", "Contab. FT", "Venc. FT", "Envío Log.", "Fecha Pago",
           "Imp. Pendiente"]
DATE_COLS = ["Fecha", "Aprobación", "Emisión FT", "Contab. FT", "Venc. FT", "Envío Log.", "Fecha Pago"]
AMOUNT_COLS = ["Total Sin IGV", "FT Sin Imp.", "Imp. Pendiente"]
INT_COLS = ["Preliminar", "Numero OC"]
TEXT_COLS = [h for h in HEADERS if h not in DATE_COLS + AMOUNT_COLS + INT_COLS]
INVOICE_COLS = HEADERS[11:]  # N° Factura .. Imp. Pendiente

MONEDAS = ["SOL", "USD", "CAD", "AUD", "GBP", "EUR", "ZAR"]
ESTADOS_OC = ["Cerrado", "En proceso", "Cancelado"]
ESTADOS_FACT = ["Pagado", "Pendiente de pago"]
TIPOS = ["FT", "RH", "CN", "IN", "DI", "RS", "BV", ""]
SERIE_RE = re.compile(r"^(\d{2})?OC[BS]$")

LAST = 10000                 # última fila que cubren las fórmulas
N_PROV_FILTRO = 20           # casillas de proveedor en el panel de filtros
TOP_N = 15
NOMBRE_MAX = 45

OBS_SOBRANTE = "Fila sin factura en OC que sí tiene facturas"
OBS_SOLO_AN = "OC cuya única factura fue anulada (AN): queda sin factura"
OBS_EST_SIN_N = "Estado de factura sin N° Factura"
OBS_SIN_TIPO = "Factura sin Tipo"
OBS_RUC_NOMBRES = "RUC con más de un nombre de proveedor"


# ---------------------------------------------------------------------------
# Lectura del archivo crudo
# ---------------------------------------------------------------------------
def _norm(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def leer_csv(path):
    """El CSV de SAP repite en cada fila: título, fecha, hora, 'Fecha de reporte',
    los 21 encabezados, los 21 datos y el pie. También acepta un CSV tabular normal."""
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with open(path, newline="", encoding=enc) as fh:
                rows = list(csv.reader(fh))
            break
        except UnicodeDecodeError:
            continue
    meta = {"fecha": None, "hora": None}
    data, descartes, header_pos = [], [], None
    for n, r in enumerate(rows, start=1):
        rn = [_norm(x) for x in r]
        h = next((i for i in range(len(rn) - 20) if rn[i:i + 21] == HEADERS), None)
        if h is not None and len(rn) >= h + 42:          # formato "plano" de SAP
            if meta["fecha"] is None and rn[0] == "Control de ordenes":
                meta["fecha"], meta["hora"] = rn[1], rn[2]
            data.append((n, r[h + 21:h + 42]))
        elif h is not None:                               # fila de encabezados tabular
            header_pos = h
        elif header_pos is not None and len(r) >= header_pos + 21:
            data.append((n, r[header_pos:header_pos + 21]))
        else:
            descartes.append((n, "Fila que no es una orden (título/pie/blanco)", " | ".join(x for x in rn if x)[:120]))
    return data, meta, len(rows), descartes


def leer_xlsx(path):
    """El xlsx de SAP es un reporte con formato: título y fecha arriba, encabezados en una
    fila intermedia, columnas separadoras vacías y pie al final. Se ubica todo por nombre."""
    df = pd.read_excel(path, header=None, engine="calamine", dtype=object)
    meta = {"fecha": None, "hora": None}
    hdr_row = None
    for i in range(len(df)):
        vals = [_norm(v) if pd.notna(v) else "" for v in df.iloc[i]]
        if "Usuario" in vals and "Numero OC" in vals:
            hdr_row = i
            colmap = {h: vals.index(h) for h in HEADERS if h in vals}
            break
        for v in df.iloc[i]:
            if isinstance(v, datetime) and meta["fecha"] is None:
                meta["fecha"] = v
            elif isinstance(v, time) and meta["hora"] is None:
                meta["hora"] = v
    if hdr_row is None:
        sys.exit("No se encontró la fila de encabezados (Usuario, Numero OC, ...) en el xlsx.")
    faltan = [h for h in HEADERS if h not in colmap]
    if faltan:
        sys.exit(f"Faltan columnas en el xlsx: {faltan}")
    data, descartes = [], []
    for i in range(hdr_row + 1, len(df)):
        fila = df.iloc[i]
        vals = [fila[colmap[h]] for h in HEADERS]
        if all(pd.isna(v) or _norm(v) == "" for v in vals):
            otros = " | ".join(_norm(v) for v in fila if pd.notna(v))
            if otros:
                descartes.append((i + 1, "Fila que no es una orden (título/pie/blanco)", otros[:120]))
            continue
        data.append((i + 1, [None if pd.isna(v) else v for v in vals]))
    return data, meta, len(df), descartes


# ---------------------------------------------------------------------------
# Conversión de tipos
# ---------------------------------------------------------------------------
def a_fecha(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, datetime):
        return v.replace(hour=0, minute=0, second=0, microsecond=0)
    s = _norm(v)
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    raise ValueError(f"fecha no reconocida: {s!r}")


def a_hora(v):
    if v is None or isinstance(v, time):
        return v
    return datetime.strptime(_norm(v), "%H:%M:%S").time()


def a_numero(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = _norm(v).replace(",", "")
    return float(s) if s else None


def a_entero(v):
    n = a_numero(v)
    return None if n is None else int(round(n))


def a_texto(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return _norm(v)


def convertir(data):
    registros, descartes = [], []
    for n, vals in data:
        rec = dict(zip(HEADERS, vals))
        try:
            out = {"_fila": n}
            for h in HEADERS:
                v = rec[h]
                if h in DATE_COLS:
                    out[h] = a_fecha(v)
                elif h in AMOUNT_COLS:
                    out[h] = a_numero(v)
                elif h in INT_COLS:
                    out[h] = a_entero(v)
                else:
                    out[h] = a_texto(v)
        except ValueError as e:
            descartes.append((n, f"Valor no convertible ({e})", " | ".join(a_texto(x) for x in vals)[:120]))
            continue
        if out["Numero OC"] is None:
            descartes.append((n, "Fila sin Numero OC (no es una orden)", " | ".join(a_texto(x) for x in vals if x)[:120]))
            continue
        registros.append(out)
    return pd.DataFrame(registros), descartes


# ---------------------------------------------------------------------------
# Limpieza
# ---------------------------------------------------------------------------
def limpiar(df):
    log = {"descartes": [], "avisos": [], "conteos": []}
    n0 = len(df)

    # 1) Anulados: Tipo = AN, Estado Factura = Anulado o Estado (OC) = Anulado
    anul = (df["Tipo"] == "AN") | (df["Estado Factura"] == "Anulado") | (df["Estado"] == "Anulado")
    ocs_validas = set(df.loc[~anul, "Numero OC"])
    # OCs que quedarían sin ninguna fila: se conservan como una fila sin factura
    solo_an = df[anul & ~df["Numero OC"].isin(ocs_validas)]
    placeholders = []
    for oc, g in solo_an.groupby("Numero OC", sort=False):
        base = g.sort_values("Aprobación", na_position="first").iloc[-1].copy()
        for c in INVOICE_COLS:
            base[c] = None if c in DATE_COLS + AMOUNT_COLS else ""
        base["_obs_an"] = True
        base["_fila"] = g["_fila"].min()
        placeholders.append(base)
    for _, r in df[anul].iterrows():
        motivo = "Tipo = AN" if r["Tipo"] == "AN" else f"Estado/Estado Factura = Anulado"
        log["descartes"].append((r["_fila"], f"Anulada ({motivo})", f"OC {r['Numero OC']} · {r['N° Factura'] or '(sin N°)'} · {r['Estado Factura']}"))
    df = df[~anul].copy()
    df["_obs_an"] = False
    if placeholders:
        df = pd.concat([df, pd.DataFrame(placeholders)], ignore_index=True)
    df["_obs_an"] = df["_obs_an"].astype(bool)
    log["conteos"].append(("Filas anuladas eliminadas (Tipo AN / Anulado)", int(anul.sum())))
    log["conteos"].append(("OCs con solo facturas anuladas, conservadas como fila sin factura", len(placeholders)))

    # 2) Duplicados por modificación de OC: todas las columnas iguales salvo Aprobación.
    #    Se conserva la fila con la Aprobación más reciente.
    clave = [h for h in HEADERS if h != "Aprobación"]
    k = df[clave].map(lambda v: "" if pd.isna(v) else str(v))
    df["_grupo"] = k.agg("\x1f".join, axis=1)
    df["_orden"] = df.groupby("_grupo")["_fila"].transform("min")
    df = df.sort_values(["_grupo", "Aprobación"], na_position="first")
    dup = df.duplicated("_grupo", keep="last")
    for _, r in df[dup].iterrows():
        log["descartes"].append((r["_fila"], "Duplicado por modificación (se conserva la Aprobación más reciente)",
                                 f"OC {r['Numero OC']} · {r['N° Factura'] or '(sin factura)'} · Aprob. {r['Aprobación']:%d/%m/%Y}" if pd.notna(r["Aprobación"]) else f"OC {r['Numero OC']}"))
    df = df[~dup].sort_values(["_orden", "_fila"]).reset_index(drop=True)
    log["conteos"].append(("Filas duplicadas por modificación eliminadas", int(dup.sum())))

    # 3) Validaciones (se avisa, no se corrige)
    def inesperados(col, permitidos):
        malos = df.loc[~df[col].isin(permitidos), col].value_counts()
        for v, c in malos.items():
            log["avisos"].append((f"Valor inesperado en {col}", f"{v!r}: {c} fila(s)"))
    inesperados("Estado", ESTADOS_OC)
    inesperados("Mon", MONEDAS)
    inesperados("Estado Factura", ESTADOS_FACT + [""])
    inesperados("Tipo", TIPOS)
    for v, c in df.loc[~df["Serie"].str.match(SERIE_RE), "Serie"].value_counts().items():
        log["avisos"].append(("Valor inesperado en Serie", f"{v!r}: {c} fila(s)"))
    for col in ["Fecha"]:
        n = int(df[col].isna().sum())
        if n:
            log["avisos"].append((f"{col} vacía", f"{n} fila(s): no entran en las tablas por año"))

    con_fact = df[df["N° Factura"] != ""]
    dups = con_fact[con_fact.duplicated(["RUC", "N° Factura"], keep=False)]
    if len(dups):
        for (ruc, f), g in dups.groupby(["RUC", "N° Factura"]):
            log["avisos"].append(("RUC + N° Factura duplicado", f"{ruc} · {f} · OCs {sorted(g['Numero OC'].unique())}"))
    else:
        log["conteos"].append(("RUC + N° Factura único por fila", "Sí (0 duplicados)"))

    oc_multi = df.groupby("Numero OC")[["RUC", "Mon", "Total Sin IGV", "Estado", "Serie"]].nunique(dropna=False)
    for oc in oc_multi[(oc_multi > 1).any(axis=1)].index:
        log["avisos"].append(("OC con datos de cabecera distintos entre filas", f"OC {oc}"))

    # 4) Observaciones (marcas para revisar y filtrar)
    obs = [[] for _ in range(len(df))]
    tiene_fact = df.groupby("Numero OC")["N° Factura"].transform(lambda s: (s != "").any())
    for i in df.index[(df["N° Factura"] == "") & tiene_fact & ~df["_obs_an"]]:
        obs[i].append(OBS_SOBRANTE)
    for i in df.index[df["_obs_an"]]:
        obs[i].append(OBS_SOLO_AN)
    for i in df.index[(df["N° Factura"] == "") & (df["Estado Factura"] != "")]:
        obs[i].append(OBS_EST_SIN_N)
    for i in df.index[(df["N° Factura"] != "") & (df["Tipo"] == "")]:
        obs[i].append(OBS_SIN_TIPO)
    nombres = df.groupby("RUC")["Proveedor"].unique()
    ruc_multi = nombres[nombres.map(len) > 1]
    for i in df.index[df["RUC"].isin(ruc_multi.index)]:
        obs[i].append(OBS_RUC_NOMBRES)
    df["Observación"] = ["; ".join(o) for o in obs]
    for etiqueta in [OBS_SOBRANTE, OBS_SOLO_AN, OBS_EST_SIN_N, OBS_SIN_TIPO, OBS_RUC_NOMBRES]:
        n = int(df["Observación"].str.contains(etiqueta, regex=False).sum())
        log["conteos"].append((f"Marcadas: {etiqueta}", n))
    for ruc, ns in ruc_multi.items():
        propuesto = max(ns, key=len)
        log["avisos"].append(("RUC con varios nombres (NO normalizado; propuesta: usar el más largo)",
                              f"{ruc}: " + " || ".join(ns) + f"  →  {propuesto}"))

    # Nombres que romperían los criterios de SUMIFS
    for p in df["Proveedor"].unique():
        if len(p) > 255 or re.search(r"[*?~]", p) or p[:1] in "=<>":
            log["avisos"].append(("Proveedor con caracteres especiales para SUMIFS (revisar gráfico)", p))

    log["conteos"].insert(0, ("Filas de órdenes leídas", n0))
    log["conteos"].append(("Filas finales en Sheet1", len(df)))
    log["conteos"].append(("  · con N° Factura", int((df["N° Factura"] != "").sum())))
    log["conteos"].append(("  · sin N° Factura", int((df["N° Factura"] == "").sum())))
    log["conteos"].append(("OCs distintas", int(df["Numero OC"].nunique())))
    return df[HEADERS + ["Observación"]], log


# ---------------------------------------------------------------------------
# Escritura del libro
# ---------------------------------------------------------------------------
def rng(col):
    return f"Sheet1!${col}$3:${col}${LAST}"


class Layout:
    """Posiciones (1-based) del panel de filtros y de las tablas de Resumen Gastos."""
    def __init__(self, usuarios, anios):
        nu, ny = len(usuarios), len(anios)
        self.usuarios, self.anios = usuarios, anios
        self.ef = (6, 8)                       # Estado de factura (cabecera fila 5)
        self.eoc = (11, 13)                    # Estado de OC (cabecera 10)
        self.usr = (16, 15 + nu)               # Comprador (cabecera 15)
        self.fec_hdr = 17 + nu
        self.fec = (18 + nu, 20 + nu)          # OC, Emisión, Pago
        self.prov_hdr = 22 + nu
        self.prov = (24 + nu, 23 + nu + N_PROV_FILTRO)
        self.prov_act = self.prov[1] + 1
        # Tablas: título, cabecera, años, total, control
        self.tablas = {}
        r = 9
        for key in ["monto", "ordenes", "facturas"]:
            self.tablas[key] = dict(titulo=r, hdr=r + 1, ini=r + 2, fin=r + 1 + ny, tot=r + 2 + ny, ctl=r + 3 + ny)
            r += ny + 5
        self.tablas["estado"] = dict(titulo=r, hdr=r + 1, ini=r + 2, fin=r + 3, tot=r + 4, ctl=r + 5)

    def filtros_sumproduct(self):
        """Condiciones (lista) para SUMPRODUCT sobre rangos de Sheet1."""
        ef, eoc, usr, pr = self.ef, self.eoc, self.usr, self.prov
        f0 = self.fec[0]
        c = [
            f'--ISNUMBER(MATCH({rng("N")}&"|",$L${ef[0]}:$L${ef[1]}&"|",0))',
            f'--ISNUMBER(MATCH({rng("B")}&"|",$L${eoc[0]}:$L${eoc[1]}&"|",0))',
            f'--ISNUMBER(MATCH({rng("A")}&"|",$L${usr[0]}:$L${usr[1]}&"|",0))',
            f'--((($M${self.prov_act}=0)+($M${self.prov_act}=1)*ISNUMBER(MATCH({rng("I")}&"|",$M${pr[0]}:$M${pr[1]}&"|",0)))>0)',
        ]
        for k, col in enumerate("FPT"):
            r = f0 + k
            c.append(f'--((($M${r}=0)+($M${r}=1)*({rng(col)}>=$N${r})*({rng(col)}<=$O${r}))>0)')
        return c

    def filtro_fila(self, r):
        """Condición 1/0 para una fila r de Sheet1 (Calc_Graficos)."""
        RG = "'Resumen Gastos'!"
        ef, eoc, usr, pr = self.ef, self.eoc, self.usr, self.prov
        f0 = self.fec[0]
        partes = [
            f'ISNUMBER(MATCH(Sheet1!N{r}&"|",{RG}$L${ef[0]}:$L${ef[1]}&"|",0))',
            f'ISNUMBER(MATCH(Sheet1!B{r}&"|",{RG}$L${eoc[0]}:$L${eoc[1]}&"|",0))',
            f'ISNUMBER(MATCH(Sheet1!A{r}&"|",{RG}$L${usr[0]}:$L${usr[1]}&"|",0))',
            f'OR({RG}$M${self.prov_act}=0,ISNUMBER(MATCH(Sheet1!I{r}&"|",{RG}$M${pr[0]}:$M${pr[1]}&"|",0)))',
        ]
        for k, col in enumerate("FPT"):
            q = f0 + k
            partes.append(f'OR({RG}$M${q}=0,AND(Sheet1!{col}{r}>={RG}$N${q},Sheet1!{col}{r}<={RG}$O${q}))')
        return "=--AND(" + ",".join(partes) + ")"


def fecha_reporte(meta):
    """Fecha y hora del reporte SAP como datetime (o None)."""
    if meta["fecha"] is None:
        return None
    f = a_fecha(meta["fecha"])
    h = a_hora(meta["hora"]) if meta["hora"] is not None else time(0, 0)
    return datetime.combine(f.date(), h)


def escribir_libro(df, meta, log, salida, seguimiento=None, comentarios_oc=None):
    """`seguimiento`: filas de seguimiento.tabla_para_excel(); agrega la hoja 'Seguimiento'.
    `comentarios_oc`: {OC: (fecha, texto)} con el último comentario de cada OC.
    Con cualquiera de los dos se agrega la columna Y de Sheet1 (estado de seguimiento · último comentario)."""
    usuarios = sorted(df["Usuario"].unique())
    anios = sorted(int(y) for y in df["Fecha"].dropna().map(lambda d: d.year).unique())
    L = Layout(usuarios, anios)
    n_prov = df.loc[df["Proveedor"] != "", "Proveedor"].str.upper().nunique()

    wb = xlsxwriter.Workbook(salida, {"default_format_properties": {"font_name": "Arial", "font_size": 10}})
    F = lambda **k: wb.add_format({"font_name": "Arial", "font_size": 10, **k})
    AZUL = "#1F3864"
    fmt = dict(
        date=F(num_format="dd/mm/yyyy"), time=F(num_format="hh:mm:ss"), amt=F(num_format="#,##0.00"),
        int=F(num_format="0"), txt=F(num_format="@"), base=F(),
        hdr=F(bold=True, font_color="white", bg_color=AZUL), title=F(bold=True, font_size=14),
        sub=F(font_color="#595959"), h2=F(bold=True, font_size=11), bold=F(bold=True),
        panel=F(bold=True, font_color=AZUL),
        inp=F(bold=True, font_color="#0000FF", bg_color="#FFFF00"),
        inp_date=F(bold=True, font_color="#0000FF", bg_color="#FFFF00", num_format="dd/mm/yyyy"),
        aux=F(font_color="#7F7F7F"), aux_date=F(font_color="#7F7F7F", num_format="dd/mm/yyyy"),
        aux_amt=F(font_color="#7F7F7F", num_format='#,##0.00;(#,##0.00);"-"'),
        aux_int=F(font_color="#7F7F7F", num_format='#,##0;(#,##0);"-"'),
        m=F(num_format='#,##0.00;(#,##0.00);"-"'), n=F(num_format='#,##0;(#,##0);"-"'),
        year=F(num_format="0", font_color="#0000FF"), year_f=F(num_format="0"),
        tot_l=F(bold=True, bg_color="#F2F2F2"),
        tot_m=F(bold=True, bg_color="#F2F2F2", num_format='#,##0.00;(#,##0.00);"-"'),
        tot_n=F(bold=True, bg_color="#F2F2F2", num_format='#,##0;(#,##0);"-"'),
        tot_p=F(bold=True, bg_color="#F2F2F2", num_format='0.0%;-0.0%;"-"'),
        pct=F(num_format='0.0%;-0.0%;"-"'),
        no=F(bold=True, font_color="#9C0006", bg_color="#FFC7CE"),
    )

    # ---------------- Sheet1 ----------------
    ws = wb.add_worksheet("Sheet1")
    ws.write("A1", "Control de ordenes")
    fecha = a_fecha(meta["fecha"]) if meta["fecha"] is not None else None
    hora = a_hora(meta["hora"]) if meta["hora"] is not None else None
    if fecha:
        ws.write_datetime("B1", fecha, fmt["date"])
    if hora:
        ws.write_datetime("C1", hora, fmt["time"])
    ws.write("D1", "Fecha de reporte")
    extra = ["Observación", "Cuenta OC (filtros)", "Monto OC (filtros)"]
    ws.write_row(1, 0, HEADERS + extra)
    for i, rec in enumerate(df.itertuples(index=False), start=2):
        for j, h in enumerate(HEADERS + ["Observación"]):
            v = rec[j]
            if h in DATE_COLS:
                ws.write_datetime(i, j, v, fmt["date"]) if v is not None and not pd.isna(v) else ws.write_blank(i, j, None, fmt["date"])
            elif h in AMOUNT_COLS:
                ws.write_number(i, j, v, fmt["amt"]) if v is not None and not pd.isna(v) else ws.write_blank(i, j, None, fmt["amt"])
            elif h in INT_COLS:
                ws.write_number(i, j, v, fmt["int"])
            elif h == "RUC":
                ws.write_string(i, j, v, fmt["txt"])
            elif v:
                ws.write_string(i, j, v)
    for r in range(3, LAST + 1):
        ws.write_formula(f"W{r}", f"=Calc_Graficos!F{r}", fmt["aux_int"])
        ws.write_formula(f"X{r}", f"=Calc_Graficos!D{r}", fmt["aux_amt"])
    widths = dict(A=18, B=10.3, C=8.7, D=10, E=10.6, F=10.7, G=10.7, H=12.5, I=54, J=5, K=13, L=14,
                  M=5, N=16.3, O=12, P=10.7, Q=10.7, R=10.7, S=10.7, T=10.9, U=13.2, V=40, W=10, X=13, Y=60)
    for c, w in widths.items():
        ws.set_column(f"{c}:{c}", w)
    ultima_col = 21
    if seguimiento is not None or comentarios_oc is not None:
        ws.write(1, 24, "Seguimiento (estado · último comentario)")
        estado = {s["numero_oc"]: s["estado_seg"] for s in (seguimiento or [])
                  if s["situacion"] in ("Activa", "Cerrada sin factura")}
        com = comentarios_oc or {}
        seg_oc = {}
        for oc in set(estado) | set(com):
            partes = [estado[oc]] if oc in estado else []
            if oc in com:
                f, t = com[oc]
                partes.append(f"{f[8:10]}/{f[5:7]}: {t}")
            seg_oc[oc] = " · ".join(partes)
        for i, oc in enumerate(df["Numero OC"], start=2):
            if oc in seg_oc:
                ws.write_string(i, 24, seg_oc[oc])
        ultima_col = 24
    ws.autofilter(1, 0, 1 + len(df), ultima_col)
    ws.freeze_panes(2, 0)
    ws.write_comment("V2", "Marcas de revisión generadas por procesar_sap.py (filtrable). No afecta los cálculos.")
    ws.write_comment("W2", "1 si esta fila es la primera de su OC que pasa los filtros de 'Resumen Gastos'. Sirve para contar cada OC una sola vez.")
    ws.write_comment("X2", "Total Sin IGV contado una sola vez por OC (en la primera fila de la OC que pasa los filtros).")

    # ---------------- Resumen Gastos ----------------
    rs = wb.add_worksheet("Resumen Gastos")
    rs.hide_gridlines(2)
    rs.set_column("A:A", 27.8)
    rs.set_column("B:H", 15.8)
    rs.set_column("I:I", 4.7)
    rs.set_column("J:J", 40)
    rs.set_column("K:K", 14.8)
    rs.set_column("L:L", 19.4)
    rs.set_column("M:O", 12)
    rs.write("A1", "Resumen de gastos por año y moneda", fmt["title"])
    rs.write("A2", "Fuente: Sheet1 (filas 3 a 10000) · Monto = Total Sin IGV (col. K) contado una vez por OC (col. X) · "
                   "Año según Fecha de la OC (col. F) · Las monedas no se suman entre sí", fmt["sub"])
    rs.write("J3", "FILTROS — elige Sí / No en cada opción para incluirla o excluirla", fmt["panel"])

    def bloque_si_no(hdr_row, titulo, opciones, aux_vacio=None):
        r0 = hdr_row
        rs.write_row(r0 - 1, 9, [titulo, "Incluir", "Valor auxiliar"], fmt["hdr"])
        for k, op in enumerate(opciones):
            r = r0 + 1 + k
            rs.write(r - 1, 9, op)
            rs.write(r - 1, 10, "Sí", fmt["inp"])
            if aux_vacio is not None and op == aux_vacio:
                rs.write_formula(r - 1, 11, f'=IF(K{r}="Sí","","§")', fmt["aux"])
            else:
                rs.write_formula(r - 1, 11, f'=IF(K{r}="Sí",J{r},"§")', fmt["aux"])
        rango = f"K{r0 + 1}:K{r0 + len(opciones)}"
        rs.data_validation(rango, {"validate": "list", "source": ["Sí", "No"]})
        rs.conditional_format(rango, {"type": "cell", "criteria": "==", "value": '"No"', "format": fmt["no"]})
        return rango

    k_ef = bloque_si_no(5, "Estado de factura", ESTADOS_FACT + ["Sin factura"], aux_vacio="Sin factura")
    k_eoc = bloque_si_no(10, "Estado de OC", ESTADOS_OC)
    k_usr = bloque_si_no(15, "Comprador (Usuario)", usuarios)

    fh = L.fec_hdr
    rs.write_row(fh - 1, 9, ["Fechas (vacío = sin filtro)", "Desde", "Hasta", "Activo", "Desde ef.", "Hasta ef."], fmt["hdr"])
    for k, nombre in enumerate(["Fecha de la OC", "Emisión de factura", "Fecha de pago"]):
        r = L.fec[0] + k
        rs.write(r - 1, 9, nombre)
        rs.write_blank(r - 1, 10, None, fmt["inp_date"])
        rs.write_blank(r - 1, 11, None, fmt["inp_date"])
        rs.write_formula(r - 1, 12, f"=IF(COUNT(K{r}:L{r})>0,1,0)", fmt["aux"])
        rs.write_formula(r - 1, 13, f'=IF(K{r}="",1,K{r})', fmt["aux_date"])
        rs.write_formula(r - 1, 14, f'=IF(L{r}="",2958465,L{r})', fmt["aux_date"])
    rs.data_validation(f"K{L.fec[0]}:L{L.fec[1]}", {"validate": "date", "criteria": ">=", "value": datetime(2000, 1, 1),
                                                     "error_message": "Ingresa una fecha (dd/mm/yyyy) o deja vacío."})

    ph = L.prov_hdr
    rs.write_row(ph - 1, 9, ["Proveedor (vacío = todos)", "", "", "Valor auxiliar"], fmt["hdr"])
    rs.write(ph, 9, f"Elige hasta {N_PROV_FILTRO} proveedores (desplegable)", fmt["sub"])
    for r in range(L.prov[0], L.prov[1] + 1):
        rs.write_blank(r - 1, 9, None, fmt["inp"])
        rs.write_formula(r - 1, 12, f'=IF(J{r}="","§",J{r})', fmt["aux"])
    # En el XML el operador # se guarda como ANCHORARRAY; Excel lo muestra como =Listas!$A$2#
    rs.data_validation(f"J{L.prov[0]}:J{L.prov[1]}", {"validate": "list", "source": "=_xlfn.ANCHORARRAY(Listas!$A$2)",
                                                       "error_type": "warning"})
    rs.write(L.prov_act - 1, 11, "Activo", fmt["aux"])
    rs.write_formula(L.prov_act - 1, 12, f"=IF(COUNTA(J{L.prov[0]}:J{L.prov[1]})>0,1,0)", fmt["aux"])

    # Resumen de filtros activos A4:B8
    def resumen_lista(rango_k, rango_j):
        return (f'=IF(COUNTIF({rango_k},"Sí")=0,"(ninguno)",IF(COUNTIF({rango_k},"Sí")=ROWS({rango_k}),"Todos",'
                f'_xlfn.TEXTJOIN(", ",TRUE,IF({rango_k}="Sí",{rango_j},""))))')
    lineas = [("Estado de factura:", k_ef), ("Estado de OC:", k_eoc), ("Comprador:", k_usr)]
    for k, (et, rk) in enumerate(lineas):
        rs.write(3 + k, 0, et, fmt["bold"])
        rs.write_dynamic_array_formula(3 + k, 1, 3 + k, 1, resumen_lista(rk, rk.replace("K", "J")))
    rs.write(6, 0, "Fechas:", fmt["bold"])
    fr = f"M{L.fec[0]}:M{L.fec[1]}"
    rs.write_formula(6, 1, f'=IF(SUM({fr})=0,"Sin filtros de fecha",SUM({fr})&" filtro(s) de fecha activo(s)")')
    rs.write(7, 0, "Proveedor:", fmt["bold"])
    rs.write_formula(7, 1, f'=IF(M{L.prov_act}=0,"Todos",COUNTA(J{L.prov[0]}:J{L.prov[1]})&" proveedor(es) seleccionado(s)")')

    filtros = ",".join(L.filtros_sumproduct())
    y0, y1 = anios[0], anios[-1]
    rango_anios = f'--({rng("F")}>=DATE({y0},1,1)),--({rng("F")}<=DATE({y1},12,31))'

    def tabla(key, titulo, cuerpo, total_fmt, cell_fmt, control):
        t = L.tablas[key]
        rs.write(t["titulo"] - 1, 0, titulo, fmt["h2"])
        rs.write(t["hdr"] - 1, 0, "Año" if key != "estado" else "Estado", fmt["hdr"])
        for j, mon in enumerate(MONEDAS):
            if key == "monto":
                rs.write(t["hdr"] - 1, 1 + j, mon, fmt["hdr"])
            else:
                rs.write_formula(t["hdr"] - 1, 1 + j, f"={xl_col_to_name(1 + j)}{L.tablas['monto']['hdr']}", fmt["hdr"], mon)
        filas = anios if key != "estado" else ESTADOS_FACT
        for i, val in enumerate(filas):
            r = t["ini"] + i
            if key == "monto":
                rs.write_number(r - 1, 0, val, fmt["year"])
            elif key == "estado":
                rs.write_formula(r - 1, 0, f"=J{L.ef[0] + i}", None, val)
            else:
                rs.write_formula(r - 1, 0, f"=A{L.tablas['monto']['ini'] + i}", fmt["year_f"], val)
            for j in range(len(MONEDAS)):
                col = xl_col_to_name(1 + j)
                rs.write_dynamic_array_formula(r - 1, 1 + j, r - 1, 1 + j, cuerpo(col, t["hdr"], r), cell_fmt)
        rs.write(t["tot"] - 1, 0, "Total", fmt["tot_l"])
        for j in range(len(MONEDAS)):
            col = xl_col_to_name(1 + j)
            rs.write_formula(t["tot"] - 1, 1 + j, f"=SUM({col}{t['ini']}:{col}{t['fin']})", total_fmt)
        rs.write(t["ctl"] - 1, 0, "Control (debe ser 0)", fmt["aux"])
        rs.write_dynamic_array_formula(t["ctl"] - 1, 1, t["ctl"] - 1, 1, control(t), fmt["aux"])

    anio_r = lambda r: f'--({rng("F")}>=DATE($A{r},1,1)),--({rng("F")}<=DATE($A{r},12,31))'
    tabla("monto", "Monto total sin IGV (cada OC una sola vez)",
          lambda c, h, r: f'=SUMPRODUCT(--({rng("J")}={c}${h}),{anio_r(r)},{rng("X")})',
          fmt["tot_m"], fmt["m"],
          lambda t: f"=ROUND(SUMPRODUCT(--ISNUMBER(MATCH(Calc_Graficos!$B$3:$B${LAST},$A${t['ini']}:$A${t['fin']},0)),"
                    f"--ISNUMBER(MATCH(Calc_Graficos!$C$3:$C${LAST},$B${t['hdr']}:$H${t['hdr']},0)),"
                    f"Calc_Graficos!$D$3:$D${LAST},Calc_Graficos!$E$3:$E${LAST})-SUM(B{t['tot']}:H{t['tot']}),2)")
    tabla("ordenes", "Cantidad de órdenes",
          lambda c, h, r: f'=SUMPRODUCT(--({rng("J")}={c}${h}),{anio_r(r)},{rng("W")})',
          fmt["tot_n"], fmt["n"],
          lambda t: f"=SUMPRODUCT(--ISNUMBER(MATCH(Calc_Graficos!$B$3:$B${LAST},$A${t['ini']}:$A${t['fin']},0)),"
                    f"--ISNUMBER(MATCH(Calc_Graficos!$C$3:$C${LAST},$B${t['hdr']}:$H${t['hdr']},0)),"
                    f"Calc_Graficos!$F$3:$F${LAST})-SUM(B{t['tot']}:H{t['tot']})")
    tabla("facturas", "Nº de facturas (con N° de factura registrado)",
          lambda c, h, r: f'=SUMPRODUCT(--({rng("J")}={c}${h}),{anio_r(r)},--({rng("L")}<>""),{filtros})',
          fmt["tot_n"], fmt["n"],
          lambda t: f"=SUMPRODUCT(--ISNUMBER(MATCH(Calc_Graficos!$B$3:$B${LAST},$A${t['ini']}:$A${t['fin']},0)),"
                    f"--ISNUMBER(MATCH(Calc_Graficos!$C$3:$C${LAST},$B${t['hdr']}:$H${t['hdr']},0)),"
                    f'--({rng("L")}<>""),Calc_Graficos!$E$3:$E${LAST})-SUM(B{t["tot"]}:H{t["tot"]})')
    tf = L.tablas["facturas"]
    tabla("estado", "Nº de facturas por estado de factura (todos los años)",
          lambda c, h, r: f'=SUMPRODUCT(--({rng("J")}={c}${h}),--({rng("N")}=$A{r}),{rango_anios},--({rng("L")}<>""),{filtros})',
          fmt["tot_n"], fmt["n"],
          lambda t: f"=SUM(B{tf['tot']}:H{tf['tot']})-SUM(B{t['tot']}:H{t['tot']})")
    rs.write_comment(f"A{L.tablas['monto']['ctl']}", "Compara la tabla con Calc_Graficos (monto × pasa filtros).")
    rs.write_comment(f"A{L.tablas['ordenes']['ctl']}", "Compara la tabla con Calc_Graficos (OC únicas que pasan filtros).")
    rs.write_comment(f"A{L.tablas['facturas']['ctl']}", "Compara la tabla (filtros escritos en SUMPRODUCT) con la columna 'Pasa filtros' de Calc_Graficos.")
    rs.write_comment(f"A{L.tablas['estado']['ctl']}", "Nº de facturas total menos facturas por estado. Distinto de 0 = facturas con estado vacío.")

    # ---------------- Gráficos Proveedores ----------------
    gp = wb.add_worksheet("Gráficos Proveedores")
    GPN = "'Gráficos Proveedores'"
    gp.hide_gridlines(2)
    gp.set_column("A:A", 20.3)
    gp.set_column("B:B", 55.6)
    gp.set_column("C:C", 16.7)
    gp.set_column(3, 3 + len(anios) + 1, 14)
    gp.write("A1", "Proveedores con mayor gasto", fmt["title"])
    gp.write("A2", "Aplica los filtros de la hoja 'Resumen Gastos' (estado de factura, estado de OC, comprador, proveedor y fechas). "
                   "Monto = Total Sin IGV contado una vez por OC; las monedas no se suman entre sí.", fmt["sub"])
    gp.write("A4", "Moneda", fmt["bold"])
    gp.write("B4", "USD", fmt["inp"])
    gp.data_validation("B4", {"validate": "list", "source": MONEDAS})
    col_t = 4 + len(anios) + 2          # columna de los títulos auxiliares (H con 2 años)
    ct = xl_col_to_name(col_t - 1)
    gp.write(3, col_t - 2, "Título gráfico 1", fmt["aux"])
    gp.write_formula(3, col_t - 1, '="Top 15 proveedores por gasto sin IGV ("&B4&")"', fmt["aux"])
    gp.write(4, col_t - 2, "Título gráfico 2", fmt["aux"])
    gp.write_formula(4, col_t - 1, '="Top 15 proveedores por gasto sin IGV y por año ("&B4&")"', fmt["aux"])
    for k, et in enumerate(["Estado de factura:", "Estado de OC:", "Comprador:", "Fechas:", "Proveedor:"]):
        gp.write(4 + k, 0, et, fmt["bold"])
        gp.write_formula(4 + k, 1, f"='Resumen Gastos'!B{4 + k}")

    tr = 56                                  # fila del título de la tabla
    r1, r15 = tr + 2, tr + 1 + TOP_N         # 58..72
    r_otros, r_total, r_ctl = r15 + 1, r15 + 2, r15 + 4
    gp.write(tr - 1, 0, f"Detalle Top {TOP_N} proveedores", fmt["h2"])
    hdrs = ["#", "Proveedor", "Total"]
    gp.write_row(tr, 0, hdrs, fmt["hdr"])
    for k, y in enumerate(anios):
        gp.write_formula(tr, 3 + k, f"=\"Año \"&'Resumen Gastos'!A{L.tablas['monto']['ini'] + k}", fmt["hdr"], f"Año {y}")
    c_pct = 3 + len(anios)
    gp.write(tr, c_pct, "% del total", fmt["hdr"])
    gp.write(tr, c_pct + 1, "% acumulado", fmt["hdr"])
    sortby = lambda col: (f"INDEX(_xlfn.SORTBY(_xlfn.ANCHORARRAY(Calc_Graficos!${col}$3),"
                          f"_xlfn.ANCHORARRAY(Calc_Graficos!$H$3),-1),_xlfn.SEQUENCE({TOP_N}))")
    gp.write_dynamic_array_formula(f"A{r1}:A{r15}", f"=_xlfn.SEQUENCE({TOP_N})")
    gp.write_dynamic_array_formula(f"C{r1}:C{r15}", "=" + sortby("H"), fmt["m"])
    gp.write_dynamic_array_formula(f"B{r1}:B{r15}",
                                   f'=IF(_xlfn.ANCHORARRAY(C{r1})>0,LEFT({sortby("G")},{NOMBRE_MAX}),"")')
    calc_year_cols = [xl_col_to_name(8 + k) for k in range(len(anios))]       # I, J, ...
    for k, cy in enumerate(calc_year_cols):
        cc = xl_col_to_name(3 + k)
        gp.write_dynamic_array_formula(f"{cc}{r1}:{cc}{r15}", "=" + sortby(cy), fmt["m"])
    cp, ca = xl_col_to_name(c_pct), xl_col_to_name(c_pct + 1)
    gp.write_dynamic_array_formula(f"{cp}{r1}:{cp}{r15}", f"=IFERROR(_xlfn.ANCHORARRAY(C{r1})/$C${r_total},0)", fmt["pct"])
    gp.write_formula(f"{ca}{r1}", f"={cp}{r1}", fmt["pct"])
    for r in range(r1 + 1, r15 + 1):
        gp.write_formula(f"{ca}{r}", f"={ca}{r - 1}+{cp}{r}", fmt["pct"])
    gp.write(f"B{r_otros}", "Otros proveedores")
    for k, cc in enumerate(["C"] + [xl_col_to_name(3 + i) for i in range(len(anios))]):
        gp.write_formula(f"{cc}{r_otros}", f"={cc}{r_total}-SUM({cc}{r1}:{cc}{r15})", fmt["m"])
    gp.write_formula(f"{cp}{r_otros}", f"=IFERROR(C{r_otros}/$C${r_total},0)", fmt["pct"])
    gp.write(f"B{r_total}", "Total (todos los proveedores filtrados)", fmt["tot_l"])
    gp.write_formula(f"C{r_total}", "=SUM(_xlfn.ANCHORARRAY(Calc_Graficos!H3))", fmt["tot_m"])
    for k, cy in enumerate(calc_year_cols):
        gp.write_formula(f"{xl_col_to_name(3 + k)}{r_total}", f"=SUM(_xlfn.ANCHORARRAY(Calc_Graficos!{cy}3))", fmt["tot_m"])
    gp.write_formula(f"{cp}{r_total}", f"=SUM({cp}{r1}:{cp}{r_otros})", fmt["tot_p"])
    tm = L.tablas["monto"]
    gp.write(f"A{r_ctl}", "Control vs 'Resumen Gastos' (debe ser 0)", fmt["aux"])
    gp.write_formula(f"C{r_ctl}", f"=ROUND(C{r_total}-INDEX('Resumen Gastos'!$B${tm['tot']}:$H${tm['tot']},"
                                  f"MATCH($B$4,'Resumen Gastos'!$B${tm['hdr']}:$H${tm['hdr']},0)),2)", fmt["aux_amt"])

    sheet = "Gráficos Proveedores"
    ch1 = wb.add_chart({"type": "bar"})
    ch1.add_series({"name": "Total", "categories": [sheet, r1 - 1, 1, r15 - 1, 1],
                    "values": [sheet, r1 - 1, 2, r15 - 1, 2], "fill": {"color": AZUL},
                    "data_labels": {"value": True, "num_format": "#,##0", "font": {"size": 8}}})
    ch1.set_title({"name": [sheet, 3, col_t - 1], "name_font": {"size": 12}})
    ch1.set_y_axis({"reverse": True, "num_font": {"size": 8}})
    ch1.set_x_axis({"num_format": "#,##0", "major_gridlines": {"visible": True, "line": {"color": "#D9D9D9"}}})
    ch1.set_legend({"none": True})
    ch1.set_size({"width": 1000, "height": 420})
    gp.insert_chart("A11", ch1)

    ch2 = wb.add_chart({"type": "bar", "subtype": "stacked"})
    colores = ["#8EA9DB", AZUL, "#C9A227", "#70AD47", "#A5A5A5"]
    for k in range(len(anios)):
        ch2.add_series({"name": [sheet, tr, 3 + k], "categories": [sheet, r1 - 1, 1, r15 - 1, 1],
                        "values": [sheet, r1 - 1, 3 + k, r15 - 1, 3 + k],
                        "fill": {"color": colores[k % len(colores)]}, "gap": 60})
    ch2.set_title({"name": [sheet, 4, col_t - 1], "name_font": {"size": 12}})
    ch2.set_y_axis({"reverse": True, "num_font": {"size": 8}})
    ch2.set_x_axis({"num_format": "#,##0", "major_gridlines": {"visible": True, "line": {"color": "#D9D9D9"}}})
    ch2.set_legend({"position": "top"})
    ch2.set_size({"width": 1000, "height": 420})
    gp.insert_chart("A33", ch2)

    # ---------------- Seguimiento ----------------
    if seguimiento is not None:
        sg = wb.add_worksheet("Seguimiento")
        sg.hide_gridlines(2)
        sg.write("A1", "Seguimiento de órdenes sin factura", fmt["title"])
        sg.write("A2", "Generado desde la aplicación 'Seguimiento Ordenes' (los comentarios se registran allí). "
                       "Activa = sigue sin factura · Facturada / Fuera del reporte = cerrada automáticamente.", fmt["sub"])
        cols = [("OC", 11), ("Serie", 7), ("Comprador", 16), ("Estado OC", 11), ("Fecha OC", 11), ("Días desde OC", 9),
                ("Proveedor", 40), ("Mon", 6), ("Total Sin IGV", 13), ("Situación", 15), ("Estado seguimiento", 20),
                ("Responsable", 16), ("Fecha compromiso", 12), ("Último comentario", 50), ("Fecha últ. comentario", 15),
                ("Historial completo", 90), ("En seguimiento desde", 12), ("Facturas / cierre", 24)]
        wrap = F(text_wrap=True, valign="top")
        top = F(valign="top")
        topd = F(valign="top", num_format="dd/mm/yyyy")
        topm = F(valign="top", num_format="#,##0.00")
        venc = F(valign="top", num_format="dd/mm/yyyy", font_color="#9C0006", bold=True)
        hdr_wrap = F(bold=True, font_color="white", bg_color=AZUL, text_wrap=True, valign="vcenter")
        sg.set_row(3, 30)
        for j, (h, w) in enumerate(cols):
            sg.write(3, j, h, hdr_wrap)
            sg.set_column(j, j, w)
        d = lambda s: datetime.strptime(s[:10], "%Y-%m-%d") if s else None
        hoy = datetime.now()
        for i, s in enumerate(seguimiento, start=4):
            fo = d(s["fecha_oc"])
            fc = d(s["fecha_compromiso"])
            vals = [(s["numero_oc"], top), (s["serie"], top), (s["usuario"], top), (s["estado_oc"], top)]
            for j, (v, f) in enumerate(vals):
                sg.write(i, j, v, f)
            sg.write_datetime(i, 4, fo, topd) if fo else sg.write_blank(i, 4, None, topd)
            sg.write_formula(i, 5, f'=IF(E{i + 1}="","",TODAY()-E{i + 1})', top)
            sg.write(i, 6, s["proveedor"], top)
            sg.write(i, 7, s["mon"], top)
            sg.write_number(i, 8, s["total"] or 0, topm)
            sg.write(i, 9, s["situacion"], top)
            sg.write(i, 10, s["estado_seg"], top)
            sg.write(i, 11, s["responsable"] or "", top)
            if fc:
                sg.write_datetime(i, 12, fc, venc if (fc < hoy and s["situacion"] == "Activa") else topd)
            sg.write(i, 13, s["ult_com"] or "", wrap)
            uf = datetime.strptime(s["ult_fecha"], "%Y-%m-%d %H:%M:%S") if s["ult_fecha"] else None
            sg.write_datetime(i, 14, uf, topd) if uf else sg.write_blank(i, 14, None, topd)
            sg.write(i, 15, s["historial"], wrap)
            pv = d(s["primer_visto"])
            sg.write_datetime(i, 16, pv, topd) if pv else sg.write_blank(i, 16, None, topd)
            cierre = s["facturas"] or ""
            if s["fecha_cierre"]:
                cierre = (cierre + " · " if cierre else "") + f"cerrada {s['fecha_cierre'][8:10]}/{s['fecha_cierre'][5:7]}/{s['fecha_cierre'][:4]}"
            sg.write(i, 17, cierre, top)
        n = len(seguimiento)
        sg.autofilter(3, 0, 3 + max(n, 1), len(cols) - 1)
        sg.freeze_panes(4, 1)
        sg.conditional_format(4, 9, 3 + max(n, 1), 9, {"type": "cell", "criteria": "==", "value": '"Activa"',
                                                      "format": wb.add_format({"bg_color": "#FFF2CC"})})
        sg.conditional_format(4, 9, 3 + max(n, 1), 9, {"type": "cell", "criteria": "==", "value": '"Facturada"',
                                                      "format": wb.add_format({"bg_color": "#E2EFDA"})})

    # ---------------- Calc_Graficos ----------------
    cg = wb.add_worksheet("Calc_Graficos")
    cg.write("A1", "Cálculo auxiliar para 'Resumen Gastos' y 'Gráficos Proveedores' — una fila por cada fila de Sheet1 (misma numeración). No modificar.")
    cg.write_row(1, 0, ["Proveedor", "Año OC", "Moneda", "Monto sin IGV (1 vez por OC)", "Pasa filtros (1/0)",
                        "Cuenta OC (1/0)", "Proveedor (lista)", "Total (moneda elegida)"], fmt["bold"])
    for k in range(len(anios)):
        cg.write_formula(1, 8 + k, f"='Resumen Gastos'!A{tm['ini'] + k}", fmt["bold"], anios[k])
    for r in range(3, LAST + 1):
        cg.write_formula(f"A{r}", f'=IF(Sheet1!I{r}="","",Sheet1!I{r})')
        cg.write_formula(f"B{r}", f'=IF(ISNUMBER(Sheet1!F{r}),YEAR(Sheet1!F{r}),"")')
        cg.write_formula(f"C{r}", f'=Sheet1!J{r}&""')
        cg.write_formula(f"D{r}", f"=IF(F{r}=1,IF(ISNUMBER(Sheet1!K{r}),Sheet1!K{r},0),0)")
        cg.write_dynamic_array_formula(f"E{r}:E{r}", L.filtro_fila(r))
        cg.write_formula(f"F{r}", f"=IF(E{r}=1,--(COUNTIFS(Sheet1!$E$3:$E{r},Sheet1!$E{r},$E$3:$E{r},1)=1),0)")
    gfin = 2 + n_prov
    cg.write_dynamic_array_formula(f"G3:G{gfin}", "=_xlfn.ANCHORARRAY(Listas!A2)")
    suma = "+".join(f"_xlfn.ANCHORARRAY({c}3)" for c in calc_year_cols)
    cg.write_dynamic_array_formula(f"H3:H{gfin}", "=" + suma)
    for k, cy in enumerate(calc_year_cols):
        cg.write_dynamic_array_formula(
            f"{cy}3:{cy}{gfin}",
            f"=SUMIFS($D$3:$D${LAST},$A$3:$A${LAST},_xlfn.ANCHORARRAY($G$3),$C$3:$C${LAST},{GPN}!$B$4,"
            f"$E$3:$E${LAST},1,$B$3:$B${LAST},'Resumen Gastos'!$A${tm['ini'] + k})")
    cg.set_column("A:A", 40)
    cg.set_column("G:G", 40)
    cg.set_column("D:D", 14)
    cg.set_column(7, 8 + len(anios), 14)

    # ---------------- Listas ----------------
    ls = wb.add_worksheet("Listas")
    ls.set_column("A:A", 77.8)
    ls.write("A1", "Proveedores (lista automática desde Sheet1, col. I)", fmt["bold"])
    ls.write_dynamic_array_formula(f"A2:A{1 + n_prov}",
                                   f'=_xlfn._xlws.SORT(_xlfn.UNIQUE(_xlfn._xlws.FILTER(Sheet1!$I$3:$I${LAST},Sheet1!$I$3:$I${LAST}<>"")))')

    # ---------------- Log limpieza ----------------
    lg = wb.add_worksheet("Log limpieza")
    lg.set_column("A:A", 70)
    lg.set_column("B:B", 60)
    lg.set_column("C:C", 90)
    lg.write("A1", "Informe de limpieza", fmt["title"])
    lg.write("A2", f"Generado {datetime.now():%d/%m/%Y %H:%M} por procesar_sap.py", fmt["sub"])
    r = 3
    lg.write_row(r, 0, ["Concepto", "Valor"], fmt["hdr"]); r += 1
    for c, v in log["conteos"]:
        lg.write(r, 0, c); lg.write(r, 1, v); r += 1
    r += 1
    lg.write_row(r, 0, ["Aviso (no corregido)", "Detalle"], fmt["hdr"]); r += 1
    if not log["avisos"]:
        lg.write(r, 0, "Sin avisos"); r += 1
    for c, v in log["avisos"]:
        lg.write(r, 0, c); lg.write(r, 1, v); r += 1
    r += 1
    lg.write_row(r, 0, ["Fila del archivo SAP", "Motivo de descarte", "Detalle"], fmt["hdr"]); r += 1
    for n, motivo, det in sorted(log["descartes"], key=lambda x: x[0]):
        lg.write(r, 0, n); lg.write(r, 1, motivo); lg.write(r, 2, det); r += 1

    rs.activate()
    wb.close()
    return L


# ---------------------------------------------------------------------------
def procesar(entrada, salida=None):
    if entrada.lower().endswith(".csv"):
        data, meta, n_total, desc_lect = leer_csv(entrada)
    else:
        data, meta, n_total, desc_lect = leer_xlsx(entrada)
    df, desc_conv = convertir(data)
    df_limpio, log = limpiar(df)
    log["descartes"] = desc_lect + desc_conv + log["descartes"]
    log["conteos"].insert(0, ("Filas físicas leídas del archivo", n_total))
    log["conteos"].insert(1, ("Filas descartadas por no ser una orden (título/pie/blanco)", len(desc_lect) + len(desc_conv)))
    L = escribir_libro(df_limpio, meta, log, salida) if salida else None
    return df_limpio, log, meta, L


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    entrada, salida = sys.argv[1], sys.argv[2]
    df, log, meta, _ = procesar(entrada, salida)
    print(f"Archivo: {entrada}  (reporte {meta['fecha']} {meta['hora']})")
    for c, v in log["conteos"]:
        print(f"  {c}: {v}")
    print(f"Avisos ({len(log['avisos'])}):")
    for c, v in log["avisos"]:
        print(f"  - {c}: {v}")
    print(f"Libro generado: {salida}")


if __name__ == "__main__":
    main()
