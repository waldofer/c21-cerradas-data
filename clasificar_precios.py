"""
Escalera de decisión para los registros con el precio de cierre en guaraníes.

Función pura: no toca archivos ni el dataset. Recibe un registro, la tabla del
BCP y el pool de comparables, y devuelve qué hacer con él.

El orden de las reglas importa: la primera que matchea define el destino.
La tasa 6100 se usa SOLO como firma forense para reconocer el patrón de
conversión del sistema origen — nunca para convertir.
"""

import statistics

from cotizaciones_bcp import cotizacion_de

UMBRAL_GS = 5_000_000          # pr por encima de esto está en guaraníes
TASA_ORIGEN = 6100             # firma forense, NO se usa para convertir
BANDA_MIN, BANDA_MAX = 0.20, 5.00
MIN_PARES = 3
UMBRAL_NO_ES_VENTA = 3_000     # USD: por debajo no es un precio de venta
UMBRAL_PARCELA_GRANDE = 9_000  # m²: el $/m² deja de ser comparable
UMBRAL_SUPERFICIE_ABSURDA = 20 # m²: superficie imposible para un inmueble

UNIT_SUBTYPES = {
    "departamento", "departamento-en-pozo", "penthouse",
    "oficinas", "estacionamiento",
}


def base_m2(reg):
    """Superficie de referencia: mc para unidades, mt para el resto."""
    b = reg.get("mc") if reg.get("su") in UNIT_SUBTYPES else reg.get("mt")
    return b if b and b > 0 else None


def construir_pool(dataset):
    """Comparables usables: los que ya están en USD y tienen superficie válida."""
    return [d for d in dataset
            if d.get("pr") and 3_000 <= d["pr"] <= UMBRAL_GS and base_m2(d)]


def _pares(reg, pool, nivel):
    """Candidatos a comparable según el nivel de relajación."""
    mismo_su = [p for p in pool if p.get("su") == reg.get("su") and p["id"] != reg["id"]]
    if nivel == "L5":
        return mismo_su
    mismo_mu = [p for p in mismo_su if p.get("mu") == reg.get("mu")]
    if nivel == "L4":
        return mismo_mu
    b = base_m2(reg)
    anho = int(reg["fc"][:4])
    if nivel == "L3":   # suelta la fecha, mantiene tamaño
        return [p for p in mismo_mu if 0.5 * b <= base_m2(p) <= 1.5 * b]
    if nivel == "L2":   # suelta el tamaño, mantiene fecha
        return [p for p in mismo_mu if abs(int(p["fc"][:4]) - anho) <= 2]
    return [p for p in mismo_mu                      # L1: todo
            if 0.5 * b <= base_m2(p) <= 1.5 * b
            and abs(int(p["fc"][:4]) - anho) <= 2]


def mediana_pares(reg, pool):
    """Mediana de $/m² de los pares, relajando el criterio hasta juntar MIN_PARES."""
    if base_m2(reg) is None:
        return None, 0, "SIN_PARES"
    for nivel in ("L1", "L2", "L3", "L4", "L5"):
        cand = _pares(reg, pool, nivel)
        if len(cand) >= MIN_PARES:
            return statistics.median(p["pr"] / base_m2(p) for p in cand), len(cand), nivel
    return None, 0, "SIN_PARES"


def clasificar(reg, tabla_bcp, pool):
    """Decide el destino de un registro. Ver docstring del módulo."""
    pr = reg.get("pr")
    if not pr or pr <= UMBRAL_GS:
        return {"accion": "sin_cambio", "causa": "no_esta_en_gs",
                "pr_nuevo": None, "detalle": ""}

    cotiz, retroceso = cotizacion_de(tabla_bcp, reg["fc"])
    if cotiz is None:
        return {"accion": "marcar", "causa": "sin_cotizacion",
                "pr_nuevo": None,
                "detalle": f"No hay cotización BCP para {reg['fc']} ni 10 días antes"}

    prL = reg.get("prL")
    firma = prL * TASA_ORIGEN if prL else None

    # REGLA 1 — dígito faltante: prL*6100 da exactamente 10x el pr guardado.
    if firma and abs(firma / pr - 10) < 0.05:
        pr_restaurado = pr * 10
        return {"accion": "corregir", "causa": "digito_faltante",
                "pr_nuevo": round(pr_restaurado / cotiz, 2),
                "detalle": f"pr {pr:,.0f} -> {pr_restaurado:,.0f} Gs (dígito restituido), "
                           f"tasa {cotiz:,.2f}"}

    pr_usd = pr / cotiz

    # REGLA 2 — el monto no es un precio de venta (alquiler, seña, pago parcial).
    if pr_usd < UMBRAL_NO_ES_VENTA:
        return {"accion": "eliminar", "causa": "no_es_precio_de_venta",
                "pr_nuevo": None,
                "detalle": f"{pr:,.0f} Gs -> USD {pr_usd:,.0f}: no es precio de venta"}

    b = base_m2(reg)

    # REGLA 3 — superficie corrupta: el precio es correcto, mt/mc es imposible.
    if b is not None and b < UMBRAL_SUPERFICIE_ABSURDA and firma and abs(firma / pr - 1) < 0.05:
        return {"accion": "corregir_y_marcar", "causa": "superficie_corrupta",
                "pr_nuevo": round(pr_usd, 2),
                "detalle": f"Precio OK (USD {pr_usd:,.0f}) pero superficie {b} m² es imposible"}

    # REGLA 4 — parcela grande: el $/m² no aplica, se acepta sin validar.
    if b is not None and b > UMBRAL_PARCELA_GRANDE:
        return {"accion": "corregir", "causa": "parcela_grande",
                "pr_nuevo": round(pr_usd, 2),
                "detalle": f"{b:,.0f} m²: $/m² no comparable, aceptado sin validar"}

    med, n_pares, nivel = mediana_pares(reg, pool)

    # REGLA 5 — sin pares: imposible de validar.
    if med is None or med <= 0:
        return {"accion": "marcar", "causa": "sin_pares",
                "pr_nuevo": None,
                "detalle": f"No hay comparables suficientes (subtipo {reg.get('su')})"}

    factor = (pr_usd / b) / med if b else None

    # REGLA 6 — dentro de banda: corrección aceptada.
    if factor is not None and BANDA_MIN <= factor <= BANDA_MAX:
        return {"accion": "corregir", "causa": "validado",
                "pr_nuevo": round(pr_usd, 2),
                "detalle": f"factor {factor:.3f} ({n_pares} pares, {nivel})"}

    # REGLA 7 — fuera de banda con referencia débil: el comparable es el problema.
    if nivel in ("L3", "L4", "L5"):
        return {"accion": "corregir", "causa": "zona_sin_comparables",
                "pr_nuevo": round(pr_usd, 2),
                "detalle": f"factor {factor:.3f} pero referencia débil ({nivel}, "
                           f"{n_pares} pares): se acepta el dato"}

    # REGLA 8 — fuera de banda con pares sólidos: ambiguo, revisión manual.
    pista = ""
    if firma and 3 < firma / pr < 6:
        pista = f" PISTA: prL*6100 da {firma/pr:.1f}x el pr, el monto puede estar mal cargado."
    return {"accion": "marcar", "causa": "ambiguo_revision_manual",
            "pr_nuevo": None,
            "detalle": f"factor {factor:.3f} con {n_pares} pares sólidos ({nivel}).{pista}"}
