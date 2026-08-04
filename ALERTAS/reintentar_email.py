"""
reintentar_email.py
────────────────────
Envía (o prueba) el correo SOLO a los supervisores especificados — útil
para probar un cambio en el cuerpo del correo, o reenviar a alguien puntual
sin disparar los 37 correos completos.

Uso
───
    python reintentar_email.py "YENUBY MILENA LEON LUNA" --prueba
    python reintentar_email.py "YENUBY MILENA LEON LUNA"          # envío real
    python reintentar_email.py SUP1 SUP2 SUP3 --prueba            # varios a la vez

Siempre recalcula los cumplimientos primero (necesario para tener el
adjunto Excel de cada supervisor en memoria/temporal, igual que hace
run_alertas.py) — toma ~1-2 minutos.
"""

import sys
import argparse

from alertas_email import (
    construir_cola_envios,
    _enviar_correo_graph,
    _obtener_bytes_adjunto,
    _leer_excel_cloud,
    CORREO_REMITENTE,
    _validar_remitente,
    RUTA_MAESTRA_CLOUD,
)


def main():
    parser = argparse.ArgumentParser(
        description="Reintentar/probar envío de correo para supervisores específicos"
    )
    parser.add_argument(
        "supervisores", nargs="+",
        help="Nombres exactos (o substrings) de los supervisores a enviar",
    )
    parser.add_argument(
        "--prueba", action="store_true",
        help="Modo prueba: arma el correo (asunto, cuerpo, adjunto) pero NO lo envía de verdad",
    )
    args = parser.parse_args()

    if not args.prueba:
        if not CORREO_REMITENTE:
            print("❌ Falta CORREO_REMITENTE en las variables de entorno.")
            sys.exit(1)
        if not _validar_remitente():
            sys.exit(1)

    print("Recalculando cumplimientos (necesario para armar el adjunto de cada supervisor)...")
    from calcular_cumplimientos import main as calcular
    resultado = calcular()
    df_detalle    = resultado["df_detalle"]
    mes           = resultado["mes"]
    anio          = resultado["anio"]
    rutas_adj     = resultado["rutas_adjuntos"]
    rango_periodo = resultado.get("rango_periodo")

    try:
        df_maestro = _leer_excel_cloud(RUTA_MAESTRA_CLOUD, "Maestra de supervisores")
    except Exception as e:
        print(f"❌ No se pudo leer la maestra de supervisores desde SharePoint: {e}")
        sys.exit(1)

    print("\nArmando cola de envíos completa...")
    filas = construir_cola_envios(df_detalle, df_maestro, rutas_adj, mes, anio, rango_periodo=rango_periodo)

    # ── Filtrar solo a los supervisores pedidos ──────────────────────────
    nombres_norm = [s.upper() for s in args.supervisores]
    filas_filtradas = [
        f for f in filas
        if any(n in f["NOMBRE_SUP"] for n in nombres_norm)
    ]

    if not filas_filtradas:
        print(f"❌ Ningún supervisor en la cola matchea: {args.supervisores}")
        print(f"   Supervisores disponibles en la cola: {[f['NOMBRE_SUP'] for f in filas]}")
        sys.exit(1)

    print(f"\nReintentando correo para {len(filas_filtradas)} supervisor(es):")
    for f in filas_filtradas:
        print(f"  · {f['NOMBRE_SUP']} -> {f['CORREO']}")
    print()

    enviados = errores = 0
    for fila in filas_filtradas:
        if args.prueba:
            print(f"  🧪 [PRUEBA] Se enviaría a {fila['CORREO']} — adjunto: {fila['NOMBRE_ADJUNTO']}")
            enviados += 1
            continue
        try:
            bytes_adj = _obtener_bytes_adjunto(fila["NOMBRE_SUP"], fila["RUTA_ADJUNTO"], fila["NOMBRE_ADJUNTO"])
            _enviar_correo_graph(
                fila["CORREO"], fila["ASUNTO"], fila["CUERPO_HTML"],
                fila["NOMBRE_ADJUNTO"], bytes_adj,
            )
            print(f"  ✓ Enviado a {fila['CORREO']} ({fila['NOMBRE_SUP']})")
            enviados += 1
        except Exception as e:
            print(f"  ❌ Error enviando a {fila['CORREO']} ({fila['NOMBRE_SUP']}): {e}")
            errores += 1

    print(f"\nEnviados: {enviados}  |  Errores: {errores}")


if __name__ == "__main__":
    main()