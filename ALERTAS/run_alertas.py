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

import io
import sys
import time
import argparse
import urllib.parse
import pandas as pd
import requests
from pathlib import Path
from datetime import datetime

from alertas_logger import inicializar_logging, get_logger, ruta_log_actual
import paths   # alertas_logger ya añadió SCRIPTS/ a sys.path

BASE = paths.BASE

# La maestra de supervisores y los archivos de Detalle/Resumen ya no viven
# en disco local (paths.ALERTAS_DIR) — calcular_cumplimientos.py los guarda
# en SharePoint. Se leen aquí con los mismos helpers Graph API que usa el
# resto del pipeline.
RUTA_MAESTRA_CLOUD = f"{paths._BASES_ROOT}/ALERTAS/MAESTRO_SUPERVISORES.xlsx"

_DRIVE_ID_CACHE: dict = {}


def _obtener_default_drive_id(token: str) -> str:
    if "drive_id" in _DRIVE_ID_CACHE:
        return _DRIVE_ID_CACHE["drive_id"]
    headers = {"Authorization": f"Bearer {token}"}
    nombre_sitio = getattr(paths, "SHAREPOINT_SITE_NAME", "eficacia")
    res_site = requests.get(
        f"https://graph.microsoft.com/v1.0/sites?search={urllib.parse.quote(nombre_sitio)}",
        headers=headers,
    )
    if res_site.status_code == 200:
        sites = res_site.json().get("value", [])
        if sites:
            site = next(
                (s for s in sites
                 if nombre_sitio.lower() in (s.get("name", "").lower(), s.get("displayName", "").lower())),
                sites[0],
            )
            res_drive = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site.get('id')}/drive", headers=headers)
            if res_drive.status_code == 200:
                _DRIVE_ID_CACHE["drive_id"] = res_drive.json().get("id")
                return _DRIVE_ID_CACHE["drive_id"]
    res_root = requests.get("https://graph.microsoft.com/v1.0/sites/root", headers=headers)
    if res_root.status_code == 200:
        site_id = res_root.json().get("id")
        res_drive = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", headers=headers)
        if res_drive.status_code == 200:
            _DRIVE_ID_CACHE["drive_id"] = res_drive.json().get("id")
            return _DRIVE_ID_CACHE["drive_id"]
    raise Exception(f"No se pudo obtener el Drive ID del sitio SharePoint '{nombre_sitio}': {res_site.text}")


def _limpiar_ruta_graph(ruta_sharepoint: str) -> str:
    ruta_sharepoint = ruta_sharepoint.replace("\\", "/")
    for prefijo in ["Documentos compartidos/", "Shared Documents/"]:
        if ruta_sharepoint.startswith(prefijo):
            ruta_sharepoint = ruta_sharepoint.replace(prefijo, "", 1)
    return ruta_sharepoint.lstrip("/")


def _listar_hijos_cloud(ruta_carpeta: str) -> list:
    token = paths.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    drive_id = _obtener_default_drive_id(token)
    ruta_codificada = urllib.parse.quote(_limpiar_ruta_graph(ruta_carpeta))
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{ruta_codificada}:/children"
    items = []
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            raise Exception(f"Error al listar carpeta '{ruta_carpeta}': {response.status_code} - {response.text}")
        data = response.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def _leer_excel_cloud(ruta_sharepoint: str, descripcion: str) -> pd.DataFrame:
    token = paths.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    drive_id = _obtener_default_drive_id(token)
    ruta_limpia = _limpiar_ruta_graph(ruta_sharepoint)
    for r in (ruta_limpia, f"Documentos compartidos/{ruta_limpia}"):
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{urllib.parse.quote(r)}:/content"
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return pd.read_excel(io.BytesIO(response.content))
    raise FileNotFoundError(f"No se pudo leer {descripcion} desde SharePoint ({ruta_sharepoint}).")


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
    parser.add_argument(
        "--supervisor", nargs="+", metavar="NOMBRE",
        help=(
            "Corre TODO el pipeline (calcular + telegram + email) pero solo "
            "envía a el/los supervisor(es) indicados — el resto se calcula "
            "y se guarda igual en SharePoint, solo se filtra el envío. "
            "Coincidencia parcial: 'YENUBY' matchea 'YENUBY MILENA LEON LUNA'."
        ),
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
        # Re-leer desde SharePoint si se saltan el cálculo (los archivos ya
        # no viven en disco local — calcular_cumplimientos.py los guarda en
        # Equipo Información/BI/INVOLVES/SALIDAS/).
        try:
            candidatos = sorted(
                (h["name"] for h in _listar_hijos_cloud(paths._SALIDAS_ROOT)
                 if h.get("file") and h["name"].startswith("Detalle_Cumplimientos_")),
                reverse=True,
            )
        except Exception as e:
            print(f"❌ No se pudo listar SALIDAS en SharePoint: {e}")
            sys.exit(1)

        if candidatos:
            nombre = candidatos[0]  # p.ej. Detalle_Cumplimientos_2026_07.xlsx
            partes = Path(nombre).stem.split("_")
            anio = int(partes[-2]) if len(partes) >= 2 else datetime.now().year
            mes  = int(partes[-1]) if len(partes) >= 1 else datetime.now().month
            df_det = _leer_excel_cloud(f"{paths._SALIDAS_ROOT}/{nombre}", "Detalle Consolidado")

            nombre_res = f"Resumen_Supervisores_{anio}_{mes:02d}.xlsx"
            try:
                df_res = _leer_excel_cloud(f"{paths._SALIDAS_ROOT}/{nombre_res}", "Resumen por Supervisor")
            except Exception:
                df_res = pd.DataFrame()

            # Los adjuntos ya no se listan localmente: alertas_email.py los
            # descarga de la nube automáticamente por supervisor si no
            # encuentra copia local (esto pasa cuando se saltó 'calcular').
            rutas_adj = {}

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
            print("❌ No se encontraron archivos de cumplimiento calculados en SharePoint.")
            print("   Ejecuta primero con: python run_alertas.py --solo calcular")
            sys.exit(1)

    # ── Filtro opcional por supervisor ──────────────────────────────────
    # Todo el cálculo ya corrió completo y ya se guardó en SharePoint (arriba)
    # — esto solo recorta lo que se manda por Telegram/Email, para probar
    # con una persona sin disparar el envío completo a todo el equipo.
    if args.supervisor:
        nombres_norm = [s.upper() for s in args.supervisor]

        df_res = resultado_calculo["df_resumen"]
        mask_res = df_res["SUPERVISOR_LIDER"].astype(str).str.upper().apply(
            lambda x: any(n in x for n in nombres_norm)
        )
        df_res_filtrado = df_res[mask_res].copy()

        df_det = resultado_calculo["df_detalle"]
        nombre_match = df_det["NOMBRE"].astype(str).str.upper().apply(
            lambda x: any(n in x for n in nombres_norm)
        )
        acronimos_sup = df_det.loc[nombre_match, "ACRONIMO"].tolist()
        mask_det = nombre_match
        if "ACRONIMO_SUP" in df_det.columns:
            mask_det = mask_det | df_det["ACRONIMO_SUP"].isin(acronimos_sup)
        df_det_filtrado = df_det[mask_det].copy()

        if df_res_filtrado.empty:
            print(f"❌ Ningún supervisor en el resumen matchea: {args.supervisor}")
            sys.exit(1)

        print(f"\n🔎 Filtrado por --supervisor {args.supervisor}: "
              f"{len(df_res_filtrado)} supervisor(es), {len(df_det_filtrado)} personas en detalle")

        resultado_calculo["df_resumen"] = df_res_filtrado
        resultado_calculo["df_detalle"] = df_det_filtrado

    # Cargar maestra (SharePoint)
    try:
        df_maestro = _leer_excel_cloud(RUTA_MAESTRA_CLOUD, "Maestra de supervisores")
    except Exception as e:
        print(f"❌ MAESTRO_SUPERVISORES.xlsx no encontrado en SharePoint: {e}")
        print("   Ejecuta primero: python run_alertas.py --solo calcular")
        sys.exit(1)

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