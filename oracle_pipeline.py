"""
oracle_pipeline.py — Yve.01 Módulo Oracle
Orquesta los 4 módulos Oracle en secuencia con gestión completa de errores.
Regla crítica: nunca procesa facturas sin estado APROBADA (excepto en simulación).
Log completo de cada paso. Resumen final: X contabilizadas, importe total, X errores.
Uso: python oracle_pipeline.py [--force-sim] [--dry-run]
"""

import sys
import os
import time
from pathlib import Path
from datetime import datetime

# ─── Imports del módulo Oracle ────────────────────────────────────────────────
from oracle_auth import (
    is_simulation, test_connection, print_status, ORACLE_LEDGER_NAME
)
from oracle_lector_facturas import preparar_facturas_para_oracle
from oracle_crear_asientos import procesar_batches, guardar_simulacion_excel
from oracle_actualizar_estado import actualizar_estados, mostrar_resumen_oracle

BASE_DIR     = Path(__file__).parent
REPORTES_DIR = BASE_DIR / "reportes"
REPORTES_DIR.mkdir(exist_ok=True)
HOY = datetime.now().strftime("%Y%m%d")


def log(msg: str, level: str = "INFO"):
    """Log con timestamp y nivel."""
    ts = datetime.now().strftime("%H:%M:%S")
    icons = {"INFO": "  ", "OK": "✅", "WARN": "⚠ ", "ERROR": "❌", "STEP": "▶ "}
    icon = icons.get(level.upper(), "  ")
    print(f"[{ts}] {icon} {msg}", flush=True)


def separador(titulo: str = ""):
    """Separador visual en consola."""
    if titulo:
        pad = max(0, 62 - len(titulo) - 4)
        print(f"\n{'─'*2} {titulo} {'─'*pad}")
    else:
        print("─" * 65)


def run_pipeline(dry_run: bool = False) -> dict:
    """
    Ejecuta el pipeline completo Oracle.
    dry_run: si True, solo muestra qué haría sin ejecutar nada.
    Returns: dict con estadísticas del pipeline.
    """
    inicio = time.time()
    stats = {
        "modo":             "SIMULACION" if is_simulation() else "PRODUCCION",
        "dry_run":          dry_run,
        "inicio":           datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "facturas_leidas":  0,
        "facturas_bloqueadas": 0,
        "contabilizadas":   0,
        "errores":          0,
        "importe_total":    0.0,
        "oracle_ids":       [],
        "ruta_excel_sim":   "",
        "fin":              None,
        "duracion_s":       0,
    }

    # ════════════════════════════════════════════════════════════
    print("=" * 65)
    print("  Yve.01 — Oracle Pipeline  |  Contabilización Automática")
    print("=" * 65)
    print_status()

    if dry_run:
        log("MODO DRY-RUN: solo muestra qué haría, no ejecuta nada", "WARN")

    # ════════════════════════════════════════════════════════════
    separador("PASO 1/4  Autenticación Oracle")

    conn = test_connection()
    if not conn["ok"]:
        log(f"Conexión Oracle fallida: {conn['message']}", "ERROR")
        stats["errores"] += 1
        return stats

    log(f"{conn['message']}", "OK" if conn["ok"] else "ERROR")
    if conn.get("ledgers"):
        log(f"Ledger: {conn['ledgers'][0].get('name','?')}", "INFO")

    # ════════════════════════════════════════════════════════════
    separador("PASO 2/4  Cargar y validar facturas")

    try:
        batches, bloqueadas, df_facturas = preparar_facturas_para_oracle()
    except FileNotFoundError as e:
        log(str(e), "ERROR")
        stats["errores"] += 1
        return stats
    except Exception as e:
        log(f"Error cargando facturas: {e}", "ERROR")
        stats["errores"] += 1
        return stats

    stats["facturas_leidas"]    = len(batches)
    stats["facturas_bloqueadas"] = len(bloqueadas)

    log(f"Facturas listas para Oracle: {len(batches)}", "OK")

    if bloqueadas:
        log(f"Facturas bloqueadas (no procesan): {len(bloqueadas)}", "WARN")
        for b in bloqueadas:
            log(f"  ✗ {b['numero_factura']}: {b['motivo']}", "WARN")

    if not batches:
        log("No hay facturas que contabilizar.", "WARN")
        if is_simulation():
            log("En producción: aprueba facturas en app_aprobacion_ap.py (puerto 5002)", "INFO")
        stats["fin"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        stats["duracion_s"] = round(time.time() - inicio, 2)
        _imprimir_resumen_final(stats)
        return stats

    log(f"Importe total a contabilizar: {sum(b['total_factura'] for b in batches):,.2f} EUR", "INFO")

    if dry_run:
        log("DRY-RUN: las siguientes facturas SE CONTABILIZARÍAN:", "WARN")
        for b in batches:
            log(f"  → {b['numero_factura']} | {b['nombre_proveedor']} | {b['total_factura']:,.2f} EUR", "INFO")
        stats["fin"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        stats["duracion_s"] = round(time.time() - inicio, 2)
        return stats

    # ════════════════════════════════════════════════════════════
    separador("PASO 3/4  Enviar a Oracle GL")

    try:
        resultados = procesar_batches(batches)
    except Exception as e:
        log(f"Error procesando batches: {e}", "ERROR")
        stats["errores"] += 1
        return stats

    # Lo que el pipeline ha producido DE VERDAD, en crudo: es la unica fuente
    # del exportador GL (oracle_export_dryrun). Nunca asientos inventados.
    try:
        from oracle_actualizar_estado import guardar_asientos_producidos
        _n_as = guardar_asientos_producidos(resultados)
        if _n_as:
            log(f"Asientos guardados para el exportador GL: {_n_as}", "OK")
    except Exception as _ea:
        log(f"No se pudieron guardar los asientos para el exportador: {str(_ea)[:70]}", "WARN")

    ok_results  = [r for r in resultados if "CONTABILIZADA" in r.get("estado","")]
    err_results = [r for r in resultados if "CONTABILIZADA" not in r.get("estado","")]

    stats["contabilizadas"]  = len(ok_results)
    stats["errores"]        += len(err_results)
    stats["importe_total"]   = sum(r["total_factura"] for r in ok_results)
    stats["oracle_ids"]      = [r["oracle_id"] for r in ok_results if r.get("oracle_id")]

    if err_results:
        for r in err_results:
            log(f"Error en {r['numero_factura']}: {r.get('error','')[:60]}", "ERROR")

    # ════════════════════════════════════════════════════════════
    separador("PASO 4/4  Actualizar estados")

    try:
        df_actualizado, ruta_excel, upd_stats = actualizar_estados(resultados)
        log(f"Estados actualizados: {upd_stats['actualizadas']} facturas", "OK")
        mostrar_resumen_oracle(df_actualizado)
    except Exception as e:
        log(f"Error actualizando estados: {e}", "ERROR")
        stats["errores"] += 1

    # Guardar Excel de simulación
    if is_simulation() and resultados:
        try:
            ruta_sim = guardar_simulacion_excel(resultados)
            stats["ruta_excel_sim"] = ruta_sim
            log(f"Excel simulación: {Path(ruta_sim).name}", "OK")
        except Exception as e:
            log(f"Error guardando Excel simulación: {e}", "WARN")

    # ════════════════════════════════════════════════════════════
    stats["fin"]        = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    stats["duracion_s"] = round(time.time() - inicio, 2)

    _imprimir_resumen_final(stats)
    return stats


def _imprimir_resumen_final(stats: dict):
    """Imprime el resumen final del pipeline."""
    separador("RESUMEN FINAL")
    modo  = stats["modo"]
    icono = "🟡 SIMULACIÓN" if modo == "SIMULACION" else "🟢 PRODUCCIÓN"
    print(f"\n  Modo:                   {icono}")
    print(f"  Inicio:                 {stats['inicio']}")
    print(f"  Fin:                    {stats.get('fin','—')}")
    print(f"  Duración:               {stats['duracion_s']}s")
    print()
    ok  = stats["contabilizadas"]
    err = stats["errores"]
    imp = stats["importe_total"]
    print(f"  Facturas leídas:        {stats['facturas_leidas']}")
    print(f"  Facturas bloqueadas:    {stats['facturas_bloqueadas']}")
    if ok:
        print(f"  ✅ Contabilizadas:      {ok}")
        print(f"  💶 Importe total:       {imp:,.2f} EUR")
        if stats.get("oracle_ids"):
            print(f"  Oracle IDs:             {', '.join(stats['oracle_ids'])}")
    if err:
        print(f"  ❌ Errores:             {err}")
    if stats.get("ruta_excel_sim"):
        print(f"  📄 Excel simulación:    {Path(stats['ruta_excel_sim']).name}")
    print()
    if ok and not err:
        print("  ✅ Pipeline completado con éxito")
    elif ok and err:
        print("  ⚠  Pipeline completado con errores parciales")
    elif not ok and not err:
        print("  ℹ  Pipeline completado — sin facturas que procesar")
    else:
        print("  ❌ Pipeline completado con errores")
    print("=" * 65)


if __name__ == "__main__":
    dry_run = "--dry-run" in sys.argv

    try:
        stats = run_pipeline(dry_run=dry_run)
        # Exit code: 0 si OK, 1 si hay errores
        sys.exit(0 if stats["errores"] == 0 else 1)
    except KeyboardInterrupt:
        print("\n\n  Pipeline interrumpido por el usuario.")
        sys.exit(1)
    except Exception as e:
        print(f"\n  ERROR FATAL: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
