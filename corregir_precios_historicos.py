"""
Corrección única de los precios de cierre almacenados en guaraníes.

Hace backup, clasifica cada candidato, aplica el destino y escribe un reporte.
Es idempotente: correrlo dos veces no cambia nada la segunda vez, porque
después de la primera pasada ya no quedan registros con pr > UMBRAL_GS salvo
los marcados, que se saltean.

Uso:
    python corregir_precios_historicos.py --dry-run   # no escribe nada
    python corregir_precios_historicos.py
"""

import argparse
import json
import shutil
from pathlib import Path

from clasificar_precios import UMBRAL_GS, base_m2, clasificar, construir_pool

REPO_DIR = Path(__file__).parent
DATASET = REPO_DIR / "data" / "cerradas.json"
BACKUP = REPO_DIR / "data" / "cerradas_backup_pre_correccion.json"
TABLA = REPO_DIR / "data" / "cotizaciones_bcp.json"
REPORTE = REPO_DIR / "reporte_correccion_precios.md"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    dataset = json.loads(DATASET.read_text(encoding="utf-8"))
    tabla = json.loads(TABLA.read_text(encoding="utf-8"))
    total_previo = len(dataset)

    if not args.dry_run:
        if not BACKUP.exists():
            shutil.copy2(DATASET, BACKUP)
            print(f"Backup creado: {BACKUP}")
        else:
            print(f"Backup ya existía, se conserva: {BACKUP}")

    pool = construir_pool(dataset)
    # Los que ya tienen px de una corrida anterior no se vuelven a tocar.
    candidatos = [d for d in dataset
                  if d.get("pr") and d["pr"] > UMBRAL_GS and not d.get("px")]
    print(f"Dataset: {total_previo} registros. Candidatos: {len(candidatos)}")

    resultados = []
    for reg in candidatos:
        r = clasificar(reg, tabla, pool)
        r["reg"] = reg
        resultados.append(r)

    a_eliminar = {r["reg"]["id"] for r in resultados if r["accion"] == "eliminar"}

    if not args.dry_run:
        for r in resultados:
            reg, acc = r["reg"], r["accion"]
            if acc in ("corregir", "corregir_y_marcar"):
                reg["pr"] = r["pr_nuevo"]
            if acc in ("marcar", "corregir_y_marcar"):
                reg["px"] = 1
        dataset = [d for d in dataset if d["id"] not in a_eliminar]
        DATASET.write_text(
            json.dumps(dataset, ensure_ascii=False, separators=(",", ":")),
            encoding="utf-8")

    escribir_reporte(resultados, total_previo, len(dataset), args.dry_run)

    print(f"\nCorregidos: {sum(1 for r in resultados if r['accion'] in ('corregir','corregir_y_marcar'))}")
    print(f"Eliminados: {len(a_eliminar)}")
    print(f"Marcados:   {sum(1 for r in resultados if r['accion'] in ('marcar','corregir_y_marcar'))}")
    print(f"Total dataset: {total_previo} -> {len(dataset)}")
    print(f"Reporte: {REPORTE}")


def escribir_reporte(resultados, total_previo, total_final, dry_run):
    por_causa = {}
    for r in resultados:
        por_causa.setdefault(r["causa"], []).append(r)

    L = ["# Reporte de corrección de precios históricos", ""]
    if dry_run:
        L += ["> **DRY RUN** — no se escribió ningún cambio.", ""]
    L += [f"Dataset: {total_previo} -> {total_final} registros.",
          f"Candidatos procesados: {len(resultados)}.", "",
          "## Resumen por causa", "",
          "| Causa | Acción | Casos |", "|---|---|---|"]
    for causa, rs in sorted(por_causa.items(), key=lambda kv: -len(kv[1])):
        L.append(f"| {causa} | {rs[0]['accion']} | {len(rs)} |")

    pendientes = por_causa.get("ambiguo_revision_manual", [])
    if pendientes:
        L += ["", "## Pendientes de revisión manual", "",
              "Quedan marcados con `px` y **sin corregir**.", "",
              "| ID | Fecha | Subtipo | Barrio | pr (Gs) | m² | Detalle |",
              "|---|---|---|---|---|---|---|"]
        for r in sorted(pendientes, key=lambda r: r["reg"]["mu"]):
            g = r["reg"]
            L.append(f"| {g['id']} | {g['fc']} | {g.get('su','')} | {g.get('mu','')} "
                     f"| {g['pr']:,.0f} | {base_m2(g) or '—'} | {r['detalle']} |")

    elim = por_causa.get("no_es_precio_de_venta", [])
    if elim:
        L += ["", "## Eliminados del dataset de ventas", "",
              "| ID | Fecha | Subtipo | Barrio | Detalle |", "|---|---|---|---|---|"]
        for r in elim:
            g = r["reg"]
            L.append(f"| {g['id']} | {g['fc']} | {g.get('su','')} | {g.get('mu','')} | {r['detalle']} |")

    for causa in ("digito_faltante", "superficie_corrupta", "sin_pares", "sin_cotizacion"):
        rs = por_causa.get(causa, [])
        if not rs:
            continue
        L += ["", f"## {causa} ({len(rs)})", "",
              "| ID | Fecha | Barrio | Detalle |", "|---|---|---|---|"]
        for r in rs:
            g = r["reg"]
            L.append(f"| {g['id']} | {g['fc']} | {g.get('mu','')} | {r['detalle']} |")

    REPORTE.write_text("\n".join(L) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
