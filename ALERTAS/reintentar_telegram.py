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

import io
import sys
import argparse
import urllib.parse
import requests
import pandas as pd
from pathlib import Path

from alertas_telegram import enviar_resumen_telegram
import paths


# ─────────────────────────────────────────────────────────────────────────────
# CONEXIÓN A LA NUBE (SharePoint vía Microsoft Graph API)
# ─────────────────────────────────────────────────────────────────────────────
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


RUTA_MAESTRA_CLOUD = f"{paths._BASES_ROOT}/ALERTAS/MAESTRO_SUPERVISORES.xlsx"


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
    parser.add_argument(
        "--prueba", action="store_true",
        help="Modo prueba: arma el mensaje pero NO lo envía por Telegram (solo lo imprime)",
    )
    args = parser.parse_args()

    # ── Cargar / calcular insumos ────────────────────────────────────────
    if args.rapido:
        # Modo rápido: lee de SharePoint, mensaje sin desglose por perfil
        try:
            candidatos = sorted(
                (h["name"] for h in _listar_hijos_cloud(paths._SALIDAS_ROOT)
                 if h.get("file") and h["name"].startswith("Resumen_Supervisores_")),
                reverse=True,
            )
        except Exception as e:
            print(f"❌ No se pudo listar SALIDAS en SharePoint: {e}")
            sys.exit(1)
        if not candidatos:
            print("❌ No hay Resumen_Supervisores_*.xlsx en SharePoint — ejecuta primero el cálculo.")
            sys.exit(1)
        nombre = candidatos[0]  # p.ej. Resumen_Supervisores_2026_07.xlsx
        df_resumen = _leer_excel_cloud(f"{paths._SALIDAS_ROOT}/{nombre}", "Resumen por Supervisor")
        partes = Path(nombre).stem.split("_")
        try:
            anio = int(partes[-2])
            mes  = int(partes[-1])
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

    try:
        df_maestro = _leer_excel_cloud(RUTA_MAESTRA_CLOUD, "Maestra de supervisores")
    except Exception as e:
        print(f"❌ No se pudo leer la maestra de supervisores desde SharePoint: {e}")
        sys.exit(1)

    enviar_resumen_telegram(
        df_resumen=df_filtro,
        df_maestro=df_maestro,
        mes=mes,
        anio=anio,
        modo_prueba=args.prueba,
        df_cif_pdv=df_cif_pdv,
        df_np=df_np,
        df_pr=df_pr,
        df_sos=df_sos,
    )


if __name__ == "__main__":
    main()