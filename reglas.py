"""
Reglas de negocio compartidas: un solo lugar para decidir si una factura o una OC tiene pagos pendientes.

Las usan la app (pestañas, semáforo, «Pendientes de hoy»), Consultar OC, los Excel exportados y Power BI, para que
todas las vistas digan lo mismo.
- Moneda: SAP escribe SOL en las facturas y PEN en la OC; se muestran igual (PEN).
- Saldo residual: una factura «Pendiente de pago» con un saldo de hasta SALDO_MENOR (por moneda) es una diferencia
  de montos (redondeos, detracción ya depositada…): se considera saldada y no cuenta como vencida.
- Una factura «Pagado» está saldada aunque SAP no traiga la fecha de pago.
- OC completa: cerrada en SAP (salvo nota de entrada, que espera factura) o facturada al 100 % (±TOLERANCIA o hasta
  el saldo menor de su moneda).
"""
MON = {"SOL": "PEN"}
TOLERANCIA = 0.005                       # facturado vs OC: ±0.5 % se considera exacto
SALDO_MENOR = {"PEN": 10.0, "USD": 3.0}  # saldo que se toma como diferencia de montos, por moneda
SALDO_MENOR_OTRAS = 3.0                  # resto de monedas (EUR, CAD, GBP…)

PAGADO, PENDIENTE, PARCIAL, RESIDUAL, ANULADO = "Pagado", "Pendiente de pago", "Pago parcial", "Saldo residual", "Anulado"


def moneda(m):
    return MON.get(m, m) or ""


def umbral(m):
    """Saldo máximo que se considera diferencia de montos en la moneda `m`."""
    return SALDO_MENOR.get(moneda(m), SALDO_MENOR_OTRAS)


def texto_umbral():
    return " / ".join(f"{v:,.0f} {m}" for m, v in SALDO_MENOR.items()) + f" / {SALDO_MENOR_OTRAS:,.0f} en otras monedas"


def estado_factura(estado, pendiente, mon, estado_detalle=None):
    """Estado de una factura para mostrar y decidir: Pagado, Saldo residual, Pago parcial, Pendiente de pago…
    `estado_detalle`: estado del reporte de detalle (distingue el «Pago parcial»)."""
    if estado == PAGADO:
        return PAGADO
    if estado == PENDIENTE and pendiente is not None and pendiente <= umbral(mon):
        return RESIDUAL
    if estado == PENDIENTE and estado_detalle == PARCIAL:
        return PARCIAL
    return estado or ""


def saldada(estado, pendiente, mon):
    """¿Nada por pagar en esta factura? (pagada, o con un saldo residual)."""
    return estado_factura(estado, pendiente, mon) in (PAGADO, RESIDUAL)


def oc_completa(estado_oc, nota_entrada, total, facturado, mon):
    """¿La OC ya no espera más facturas? Cerrada en SAP (sin nota de entrada pendiente) o facturada al 100 %."""
    if estado_oc == "Cerrado" and not nota_entrada:
        return True
    return bool(total) and (facturado >= (1 - TOLERANCIA) * total or total - facturado <= umbral(mon))
