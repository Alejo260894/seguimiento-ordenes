"""
Semáforo de plazos: cada pendiente queda «Atrasado» (rojo), «Atención» (amarillo) o «En plazo».

- Requerimientos: días en la etapa actual, con umbrales por etapa.
- OC sin factura (activas): días desde la fecha de la OC; un compromiso vencido la pone en rojo.
- Facturas pendientes de pago: vencida = rojo; vence dentro de N días = amarillo.

Los umbrales se guardan en la configuración (config_seguimiento.json, clave «plazos»).
"""
from datetime import date, timedelta

import requerimientos as rq
import reglas
import seguimiento as sgm

ROJO, AMARILLO, VERDE = "rojo", "amarillo", "verde"
TEXTO = {ROJO: "Atrasado", AMARILLO: "Atención", VERDE: "En plazo"}
FONDO = {ROJO: "#F8CBAD", AMARILLO: "#FFE699"}

# Valores iniciales en días: (amarillo desde, rojo desde). Tomados de los tiempos actuales del flujo
# (p. ej. de la solicitud al envío de la OC la mediana es 4 días); se ajustan en «⚙ Plazos…».
PLAZOS_REQ = {
    "Por valorizar": (2, 5), "Valo por enviar a aprobación": (1, 3), "Valo en aprobación": (2, 5),
    "OC por emitir": (1, 3), "OC por aprobar": (2, 4), "OC por enviar": (1, 2),
    "Pendiente de factura": (15, 30), "Factura recibida": (2, 5), "Factura en validación": (2, 5),
    "Por enviar a pago": (1, 3), "Pendiente de pago": (10, 20),
}
PLAZO_OC = (30, 60)          # OC sin factura: días desde la fecha de la OC
AVISO_FACTURA = 7            # facturas: días antes del vencimiento en que pasan a amarillo

FILTROS = ["Todos", "Atrasado + Atención", "Atrasado", "Atención", "En plazo"]
# En facturas se habla de vencimiento
TEXTO_FACT = {ROJO: "Vencida", AMARILLO: "Por vencer", VERDE: "Al día"}
FILTROS_FACT = ["Todas", "Vencidas + por vencer", "Vencidas", "Por vencer", "Al día"]
_EQUIV_FACT = dict(zip(FILTROS_FACT, FILTROS))


def umbrales(cfg):
    """Umbrales vigentes: los guardados en la configuración, completados con los valores iniciales."""
    p = (cfg or {}).get("plazos") or {}
    req = {e: tuple(p.get("req", {}).get(e, v)) for e, v in PLAZOS_REQ.items()}
    return {"req": req, "oc": tuple(p.get("oc", PLAZO_OC)), "fact": int(p.get("fact", AVISO_FACTURA))}


def guardar_umbrales(cfg, u):
    cfg["plazos"] = {"req": {e: list(v) for e, v in u["req"].items()}, "oc": list(u["oc"]), "fact": u["fact"]}


def _nivel(dias, amarillo, rojo):
    if dias is None:
        return None
    return ROJO if dias >= rojo else AMARILLO if dias >= amarillo else VERDE


def nivel_req(f, u):
    """Nivel de un requerimiento (fila de requerimientos.listar) o None si está cerrado."""
    if f["etapa"] in rq.CERRADAS or f["etapa"] not in u["req"]:
        return None
    return _nivel(f["dias_etapa"], *u["req"][f["etapa"]])


def nivel_oc(f, u, hoy=None):
    """Nivel de una OC en seguimiento (fila de seguimiento.listar); solo para las activas."""
    if f["situacion"] != sgm.SIT_ACTIVA:
        return None
    hoy = (hoy or date.today()).isoformat()
    if f["fecha_compromiso"] and f["fecha_compromiso"] < hoy:
        return ROJO
    return _nivel(sgm.dias_desde(f["fecha_oc"]), *u["oc"])


def nivel_factura(f, u, hoy=None):
    """Nivel de una factura (fila de seguimiento.listar_facturas); solo vigentes sin pagar con vencimiento."""
    if (f["situacion"] != sgm.FACT_VIGENTE or not f["vencimiento"]
            or reglas.saldada(f["estado_factura"], f.get("imp_pendiente"), f.get("mon"))):     # pagada o saldo residual
        return None
    hoy = hoy or date.today()
    if f["vencimiento"] < hoy.isoformat():
        return ROJO
    return AMARILLO if f["vencimiento"] <= (hoy + timedelta(days=u["fact"])).isoformat() else VERDE


def pasa_filtro(nivel, filtro):
    filtro = _EQUIV_FACT.get(filtro, filtro)
    if filtro == "Todos":
        return True
    if filtro == "Atrasado + Atención":
        return nivel in (ROJO, AMARILLO)
    return nivel == {"Atrasado": ROJO, "Atención": AMARILLO, "En plazo": VERDE}[filtro]


def texto(nivel):
    return TEXTO.get(nivel, "")
