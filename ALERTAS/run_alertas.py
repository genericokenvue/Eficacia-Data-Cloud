"""
run_alertas.py
──────────────
Punto de entrada único del sistema de alertas de Eficacia.
Ejecuta las tres fases en secuencia:

  Fase A → calcular_cumplimientos.py
  Fase B → alertas_telegram.py
  Fase C → alertas_email.py

Uso
───
    # Ejecución completa
    python run_alertas.py

    # Solo calcular (sin enviar nada)
    python run_alertas.py --solo calcular

    # Calcular + Telegram sin correos
    python run_alertas.py --solo calcular telegram

    # Modo prueba: calcula pero no envía nada real
    python run_alertas.py --prueba

    # Solo correos (re-usa los archivos ya calculados)
    python run_alertas.py --solo email
"""

import sys
import time
import argparse
import pandas as pd
from pathlib import Path
from datetime import datetime

from alertas_logger import inicializar_logging, get_logger, ruta_log_actual
import paths   # alertas_logger ya añadió SCRIPTS/ a sys.path

BASE         = paths.BASE
DIR_ALERTAS  = paths.ALERTAS_DIR
RUTA_MAESTRA = paths.ALERTAS_MAESTRO


def main():
    parser = argparse.ArgumentParser(
        description="Sistema de alertas Eficacia — orquestador",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--solo", nargs="+",
        choices=["calcular", "telegram", "email"],
        help="Ejecutar solo las fases indicadas (default: todas)",
    )
    parser.add_argument(
        "--prueba", action="store_true",
        help="Modo prueba: calcula sin enviar mensajes ni correos reales",
    )
    parser.add_argument(
        "--setup-telegram", action="store_true",
        help="Lista los chat_id que el bot ha recibido vía getUpdates y sale.",
    )
    args = parser.parse_args()

    if args.setup_telegram:
        from setup_telegram import obtener_chat_ids
        obtener_chat_ids()
        return

    fases = args.solo or ["calcular", "telegram", "email"]
    modo_prueba = args.prueba

    ruta_log = inicializar_logging()
    log = get_logger("run_alertas")
    log.info(f"Inicio sistema de alertas | fases={fases} | prueba={modo_prueba}")

    t_inicio = time.perf_counter()
    print("\n" + "═" * 60)
    print("  SISTEMA DE ALERTAS — Eficacia")
    print(f"  Inicio: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Fases: {fases}{'  [MODO PRUEBA]' if modo_prueba else ''}")
    print(f"  Log:   {ruta_log}")
    print("═" * 60)

    resultado_calculo = None

    # ── Fase A: Calcular cumplimientos ────────────────────────────────────
    if "calcular" in fases:
        from calcular_cumplimientos import main as calcular
        resultado_calculo = calcular()
    else:
        # Re-leer desde disco si se saltan el cálculo
        candidatos = sorted(DIR_ALERTAS.glob("DETALLE_CUMPLIMIENTO_*.xlsx"), reverse=True)
        if candidatos:
            ruta = candidatos[0]
            partes = ruta.stem.split("_")
            mes  = int(partes[-2]) if len(partes) >= 2 else datetime.now().month
            anio = int(partes[-1]) if len(partes) >= 1 else datetime.now().year
            df_det = pd.read_excel(ruta)
            # Resumen también
            ruta_res = DIR_ALERTAS / f"RESUMEN_CUMPLIMIENTO_{mes:02d}_{anio}.xlsx"
            df_res = pd.read_excel(ruta_res) if ruta_res.exists() else pd.DataFrame()
            # Adjuntos
            rutas_adj = {
                f.stem.replace(f"Detalle_", "").rsplit("_", 2)[0].replace("_", " ").upper(): str(f)
                for f in (DIR_ALERTAS / "ADJUNTOS").glob("Detalle_*.xlsx")
            }
            # Sprint 13.1: detectar el rango del periodo aunque saltemos calcular
            try:
                from calcular_cumplimientos import detectar_rango_periodo
                rango_periodo = detectar_rango_periodo(mes, anio)
            except Exception:
                rango_periodo = None

            resultado_calculo = {
                "df_detalle"    : df_det,
                "df_resumen"    : df_res,
                "rutas_adjuntos": rutas_adj,
                "mes"           : mes,
                "anio"          : anio,
                "rango_periodo" : rango_periodo,
            }
        else:
            print("❌ No se encontraron archivos de cumplimiento calculados.")
            print("   Ejecuta primero con: python run_alertas.py --solo calcular")
            sys.exit(1)

    # Cargar maestra
    if not RUTA_MAESTRA.exists():
        print("❌ MAESTRO_SUPERVISORES.xlsx no encontrado.")
        print("   Ejecuta primero: python run_alertas.py --solo calcular")
        sys.exit(1)
    df_maestro = pd.read_excel(RUTA_MAESTRA)

    # ── Fase B: Telegram ──────────────────────────────────────────────────
    if "telegram" in fases and resultado_calculo:
        from alertas_telegram import enviar_resumen_telegram
        enviar_resumen_telegram(
            df_resumen   = resultado_calculo["df_resumen"],
            df_maestro   = df_maestro,
            mes          = resultado_calculo["mes"],
            anio         = resultado_calculo["anio"],
            modo_prueba  = modo_prueba,
            # DF crudos para desglosar CIF/NP/Precios/SOS por perfil PDV
            df_cif_pdv   = resultado_calculo.get("df_cif_pdv"),
            df_np        = resultado_calculo.get("df_np"),
            df_pr        = resultado_calculo.get("df_pr"),
            df_sos       = resultado_calculo.get("df_sos"),
            rango_periodo= resultado_calculo.get("rango_periodo"),
        )

    # ── Fase C: Email ─────────────────────────────────────────────────────
    if "email" in fases and resultado_calculo:
        from alertas_email import enviar_correos
        enviar_correos(
            df_detalle     = resultado_calculo["df_detalle"],
            df_maestro     = df_maestro,
            rutas_adjuntos = resultado_calculo["rutas_adjuntos"],
            mes            = resultado_calculo["mes"],
            anio           = resultado_calculo["anio"],
            modo_prueba    = modo_prueba,
            rango_periodo  = resultado_calculo.get("rango_periodo"),
        )

    t_total = time.perf_counter() - t_inicio
    log.info(f"Sistema de alertas completado en {t_total:.1f}s")
    print(f"\n  🏁 Sistema de alertas completado en {t_total:.1f}s")
    print(f"  Fin: {datetime.now().strftime('%H:%M:%S')}\n")


if __name__ == "__main__":
    main()
