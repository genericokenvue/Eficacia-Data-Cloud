"""
reintentar_telegram.py
──────────────────────
Reenvía SOLO a los supervisores especificados (útil cuando una corrida
previa falló por errores transitorios de red).

Uso
───
    python reintentar_telegram.py SUPERVISOR1 SUPERVISOR2 ...
    python reintentar_telegram.py --rapido SUP   # NO recalcula cumplimientos
                                                  # (mensaje sin desglose por
                                                   # perfil PDV)

Por defecto recalcula los cumplimientos para tener los DataFrames crudos
en memoria (df_cif_pdv, df_np, df_pr, df_sos), de modo que el mensaje
incluya el desglose por perfil PDV (Directo/Droguerías vs Proximity).
Esto toma ~30-60 s. Con --rapido se salta el cálculo y se usa el resumen
en disco — el mensaje saldrá SIN las secciones de perfil.
"""

import sys
import argparse
import pandas as pd
from pathlib import Path

from alertas_telegram import enviar_resumen_telegram
import paths


def main():
    parser = argparse.ArgumentParser(
        description="Reintentar envío de Telegram para supervisores específicos"
    )
    parser.add_argument(
        "supervisores", nargs="+",
        help="Nombres exactos (o substrings) de los supervisores a reenviar",
    )
    parser.add_argument(
        "--rapido", action="store_true",
        help="No recalcula cumplimientos (mensaje sin desglose por perfil)",
    )
    args = parser.parse_args()

    # ── Cargar / calcular insumos ────────────────────────────────────────
    if args.rapido:
        # Modo rápido: lee del disco, mensaje sin desglose por perfil
        candidatos = sorted(
            paths.ALERTAS_DIR.glob("RESUMEN_CUMPLIMIENTO_*.xlsx"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if not candidatos:
            print("❌ No hay RESUMEN_CUMPLIMIENTO_*.xlsx — ejecuta primero el cálculo.")
            sys.exit(1)
        ruta = candidatos[0]
        df_resumen = pd.read_excel(ruta)
        partes = ruta.stem.split("_")
        try:
            mes  = int(partes[-2])
            anio = int(partes[-1])
        except (ValueError, IndexError):
            from datetime import datetime
            mes = datetime.now().month
            anio = datetime.now().year
        df_cif_pdv = df_np = df_pr = df_sos = None
    else:
        # Modo completo: recalcula para obtener los DF crudos en memoria
        print("Recalculando cumplimientos (necesario para desglose por perfil)...")
        from calcular_cumplimientos import main as calcular
        resultado = calcular()
        df_resumen = resultado["df_resumen"].copy()
        mes        = resultado["mes"]
        anio       = resultado["anio"]
        df_cif_pdv = resultado.get("df_cif_pdv")
        df_np      = resultado.get("df_np")
        df_pr      = resultado.get("df_pr")
        df_sos     = resultado.get("df_sos")

    # ── Filtrar por supervisores solicitados ─────────────────────────────
    nombres_norm = [s.upper() for s in args.supervisores]
    df_resumen["SUPERVISOR_LIDER"] = df_resumen["SUPERVISOR_LIDER"].astype(str).str.upper()
    mask = df_resumen["SUPERVISOR_LIDER"].apply(
        lambda x: any(n in x for n in nombres_norm)
    )
    df_filtro = df_resumen[mask].copy()

    if df_filtro.empty:
        print(f"❌ Ningún supervisor del resumen matchea: {args.supervisores}")
        sys.exit(1)

    print(f"Reintentando envío para {len(df_filtro)} supervisor(es):")
    for s in df_filtro["SUPERVISOR_LIDER"]:
        print(f"  · {s}")
    print()

    df_maestro = pd.read_excel(paths.ALERTAS_MAESTRO)

    enviar_resumen_telegram(
        df_resumen=df_filtro,
        df_maestro=df_maestro,
        mes=mes,
        anio=anio,
        modo_prueba=False,
        df_cif_pdv=df_cif_pdv,
        df_np=df_np,
        df_pr=df_pr,
        df_sos=df_sos,
    )


if __name__ == "__main__":
    main()
