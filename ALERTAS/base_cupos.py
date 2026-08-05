"""
base_cupos.py
─────────────
Helpers para trabajar con `Base_cupos.xlsx` — la **tabla maestra de personas**
del proyecto. Resuelve identidades entre las distintas fuentes:

  • PT (CIF/NP/Precios/SOS):  llave nativa ACRONIMO + CEDULA
  • D&P Impactos:             Asesor "1001-NOMBRE", llave COD_ASESOR_ECOM
  • D&P Ventas:               Cod. Vendedor (numérico)

Diseño
──────
ACRONIMO es la **llave canónica** del proyecto. Todas las funciones aquí
devuelven o cruzan por ACRONIMO. Cuando una fuente externa trae otra
llave (CEDULA, COD_ASESOR_ECOM o NOMBRE), se traduce a ACRONIMO usando
Base cupos.

Roles excluidos (sin meta de cumplimiento estándar)
───────────────────────────────────────────────────
    NO APLICA, COORDINADOR GENERICO, TELEVENDEDORA, PROMOTOR TAT
"""

from __future__ import annotations

import io
import sys
import urllib.parse
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd
import requests

# Asegurar SCRIPTS/ en sys.path (paths.py vive ahí)
# ALERTAS/ vive DENTRO de SCRIPTS/ (SCRIPTS/ALERTAS/base_cupos.py), no es su
# hermana. `Path(__file__).resolve().parent.parent` YA apunta a SCRIPTS/ —
# no hay que agregarle "/SCRIPTS" de nuevo (eso generaba .../SCRIPTS/SCRIPTS,
# que no existe en ningún entorno).
_SCRIPTS_DIR = Path(__file__).resolve().parent.parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

import paths


# ─────────────────────────────────────────────────────────────────────────────
# CONEXIÓN A LA NUBE (SharePoint vía Microsoft Graph API)
# ─────────────────────────────────────────────────────────────────────────────
# Reutiliza paths.obtener_token_azure() (misma fuente de credenciales que el
# resto del pipeline) en vez de duplicar lógica de autenticación aquí.
_DRIVE_ID_CACHE: dict = {}


def _obtener_default_drive_id(token: str) -> str:
    """Resuelve el Drive ID del sitio SharePoint configurado en paths.SHAREPOINT_SITE_NAME."""
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
            site_id = site.get("id")
            res_drive = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", headers=headers)
            if res_drive.status_code == 200:
                drive = res_drive.json()
                _DRIVE_ID_CACHE["drive_id"] = drive.get("id")
                return drive.get("id")

    res_root = requests.get("https://graph.microsoft.com/v1.0/sites/root", headers=headers)
    if res_root.status_code == 200:
        site_id = res_root.json().get("id")
        res_drive = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", headers=headers)
        if res_drive.status_code == 200:
            _DRIVE_ID_CACHE["drive_id"] = res_drive.json().get("id")
            return res_drive.json().get("id")

    raise Exception(f"No se pudo obtener el Drive ID del sitio SharePoint '{nombre_sitio}': {res_site.text}")


def _limpiar_ruta_graph(ruta_sharepoint: str) -> str:
    """Normaliza separadores. NO remueve 'Equipo Información/' — es una carpeta real del sitio."""
    ruta_sharepoint = ruta_sharepoint.replace("\\", "/")
    for prefijo in ["Documentos compartidos/", "Shared Documents/"]:
        if ruta_sharepoint.startswith(prefijo):
            ruta_sharepoint = ruta_sharepoint.replace(prefijo, "", 1)
    return ruta_sharepoint.lstrip("/")


def _leer_excel_cloud(ruta_sharepoint: str, descripcion: str, sheet_name=0) -> pd.DataFrame:
    """Lee un archivo Excel (opcionalmente una hoja específica) directamente desde SharePoint vía Graph API."""
    print(f"  ⏳ Leyendo {descripcion} desde SharePoint ({ruta_sharepoint})...")
    token = paths.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    drive_id = _obtener_default_drive_id(token)
    ruta_limpia = _limpiar_ruta_graph(ruta_sharepoint)

    rutas_candidatas = [ruta_limpia, f"Documentos compartidos/{ruta_limpia}"]

    response = None
    url_usada = ""
    for r in rutas_candidatas:
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{urllib.parse.quote(r)}:/content"
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            url_usada = url
            break

    if response is not None and response.status_code == 200:
        print(f"    ✓ {descripcion} descargado exitosamente desde la nube.")
        return pd.read_excel(io.BytesIO(response.content), sheet_name=sheet_name)

    status = response.status_code if response is not None else "N/A"
    text = response.text if response is not None else "Sin respuesta"
    raise FileNotFoundError(
        f"No se pudo descargar {descripcion} desde SharePoint. "
        f"Última URL probada: {url_usada or rutas_candidatas} | Status: {status} - {text}"
    )


# Ruta en la nube de la tabla maestra de personas.
RUTA_BASE_CUPOS_CLOUD = f"{paths.RUTA_CARPETA_BASES_DYP}/Base_cupos.xlsx"


ROLES_EXCLUIDOS = {
    "NO APLICA",
    "COORDINADOR GENERICO",
    "TELEVENDEDORA",
    "PROMOTOR TAT",
}

# Roles que cuentan como "supervisor" en el universo de Base cupos.
# Cualquier ROL EN ACTIVO que contenga la palabra "SUPERVISOR".
def _es_supervisor(rol: str) -> bool:
    return "SUPERVISOR" in str(rol).upper()


# Sprint 17.15 — roles que generan correo/Telegram propio además de Supervisor.
def _es_gdd(rol: str) -> bool:
    rol_u = str(rol).upper()
    return "GENERADOR DE DEMANDA" in rol_u or rol_u == "GDD"


def _es_lider(rol: str) -> bool:
    return "LIDER" in str(rol).upper()


# ─────────────────────────────────────────────────────────────────────────────
# CARGA Y NORMALIZACIÓN
# ─────────────────────────────────────────────────────────────────────────────

def cargar() -> pd.DataFrame:
    """
    Lee `Base_cupos.xlsx` y devuelve un DataFrame normalizado con:
        ACRONIMO          str   (mayúsculas, strip)
        CEDULA            str   (sin .0 trailing)
        COD_ASESOR_ECOM   str   (entero como string, "" si nulo)
        COD_PT            str
        TIPO_SERVICIO     str   ("EXCLUSIVO" / "MIXTO")
        CANAL             str   ("DIRECTO", "DROGUERIAS", etc.)
        CIUDAD_CUPO       str
        ROL               str   (ROL EN ACTIVO original — UPPER)
        NOMBRE            str   (nombre canónico del proyecto)
        ES_SUPERVISOR     bool
        ES_ACTIVO         bool  (True si ROL no está en ROLES_EXCLUIDOS)

    Filtra registros sin ACRONIMO o sin NOMBRE.
    Genera automáticamente un archivo de auditoría Excel en caliente para depuración.
    """
    df = _leer_excel_cloud(RUTA_BASE_CUPOS_CLOUD, "Base_cupos.xlsx", sheet_name="Tabla total roles")
    df_original_raw = df.copy()  # Respaldamos el estado crudo para el reporte de auditoría

    df = df.rename(columns={
        "COD PT":         "COD_PT",
        "TIPO SERVICIO":  "TIPO_SERVICIO",
        "CIUDAD CUPO":    "CIUDAD_CUPO",
        "ROL EN ACTIVO":  "ROL",
    })

    # Normalización de strings
    # IMPORTANTE: fillna("") ANTES de astype(str) — en pandas reciente
    # (dtype "str" nativo), astype(str) sobre una columna con NaN NO la
    # convierte a la palabra "nan"; queda como NaN real (float), y el
    # .replace({"nan": ""}) de abajo nunca la atrapaba. Eso dejaba filas
    # con ACRONIMO/NOMBRE realmente vacíos sin filtrar más adelante
    # (sin_acronimo_o_nombre compara contra "", no contra NaN).
    for c in ["ACRONIMO", "COD_PT", "TIPO_SERVICIO", "CANAL",
              "CIUDAD_CUPO", "ROL", "NOMBRE"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str).str.strip().replace({"nan": ""})
    df["ACRONIMO"] = df["ACRONIMO"].str.upper()
    df["NOMBRE"]   = df["NOMBRE"].str.upper()
    df["ROL"]      = df["ROL"].str.upper()

    # CEDULA: int → str sin .0
    df["CEDULA"] = (
        pd.to_numeric(df["CEDULA"], errors="coerce")
          .astype("Int64").astype(str).replace("<NA>", "")
    )
    # COD_ASESOR_ECOM: idem
    df["COD_ASESOR_ECOM"] = (
        pd.to_numeric(df["COD_ASESOR_ECOM"], errors="coerce")
          .astype("Int64").astype(str).replace("<NA>", "")
    )

    # Identificar registros antes del descarte defensivo
    sin_acronimo_o_nombre = (df["ACRONIMO"] == "") | (df["NOMBRE"] == "")
    total_descartados = sin_acronimo_o_nombre.sum()

    # Defensivo: descartar filas sin ACRONIMO o sin NOMBRE
    df_filtrado = df[~sin_acronimo_o_nombre].copy()

    df_filtrado["ES_SUPERVISOR"] = df_filtrado["ROL"].apply(_es_supervisor)
    df_filtrado["ES_GDD"]        = df_filtrado["ROL"].apply(_es_gdd)       # Sprint 17.15
    df_filtrado["ES_LIDER"]      = df_filtrado["ROL"].apply(_es_lider)     # Sprint 17.15
    df_filtrado["ES_ACTIVO"]     = ~df_filtrado["ROL"].isin(ROLES_EXCLUIDOS)

    # ─────────────────────────────────────────────────────────────────────────
    # ⚙️ EXPORTACIÓN AUTOMÁTICA DE AUDITORÍA (SISTEMA DE DEPURACIÓN EN CALIENTE)
    # ─────────────────────────────────────────────────────────────────────────
    try:
        ruta_debug = Path(tempfile.gettempdir()) / "DEBUG_Base_Cupos_Normalizada.xlsx"
        
        # Construimos un reporte estructurado con múltiples hojas para análisis
        with pd.ExcelWriter(ruta_debug, engine="openpyxl") as writer:
            # Hoja 1: El universo procesado y limpio definitivo
            df_filtrado.to_excel(writer, sheet_name="1. Base_Limpia_Final", index=False)
            
            # Hoja 2: Muestra qué registros eliminó el filtro defensivo (sin acrónimo o nombre)
            df_descartados = df[sin_acronimo_o_nombre].copy()
            df_descartados.to_excel(writer, sheet_name="2. Registros_Descartados", index=False)
            
            # Hoja 3: Análisis rápido de duplicados por llave primaria ACRONIMO
            df_duplicados_acr = df_filtrado[df_filtrado.duplicated(subset=["ACRONIMO"], keep=False)].sort_values(by="ACRONIMO")
            df_duplicados_acr.to_excel(writer, sheet_name="3. Duplicados_Acronimo", index=False)

            # Hoja 4: MISMA CEDULA con DISTINTO ACRONIMO — la señal real de que
            # una persona quedó registrada dos veces en Base cupos (ej. "LIBARDO
            # JOYA" y "LIBARDO ANTONIO JOYA TEJADA" como dos filas separadas).
            # Esto NO lo detecta la Hoja 3 porque cada fila sí tiene un ACRONIMO
            # único — el duplicado está en la CEDULA, no en el ACRONIMO.
            df_dup_cedula = pd.DataFrame()
            if "CEDULA" in df_filtrado.columns:
                cedulas_validas = df_filtrado["CEDULA"].astype(str).str.strip()
                cedulas_validas = cedulas_validas[
                    cedulas_validas.replace({"": None, "nan": None}).notna()
                    & (cedulas_validas.str.upper() != "VACANTE")
                ]
                mask_cedula_dup = (
                    df_filtrado["CEDULA"].astype(str).str.strip().isin(
                        cedulas_validas[cedulas_validas.duplicated(keep=False)].unique()
                    )
                    & (df_filtrado["CEDULA"].astype(str).str.strip() != "")
                )
                df_dup_cedula = df_filtrado[mask_cedula_dup].sort_values(by="CEDULA")
                df_dup_cedula.to_excel(writer, sheet_name="4. Duplicados_Cedula", index=False)

        print(f"\n⚠️ [AUDITORÍA] Diagnóstico generado en: {ruta_debug}")
        if total_descartados > 0:
            print(f"⚠️ [AUDITORÍA] Cuidado: Se descartaron {total_descartados} filas por no tener ACRONIMO o NOMBRE. Revisa la pestaña 2.")
        if len(df_duplicados_acr) > 0:
            print(f"⚠️ [AUDITORÍA] Advertencia: Se encontraron acrónimos duplicados en la maestra. Revisa la pestaña 3.")
        if len(df_dup_cedula) > 0:
            personas_dup = df_dup_cedula["CEDULA"].nunique()
            print(f"⚠️ [AUDITORÍA] Advertencia: {personas_dup} cédula(s) aparecen en MÁS DE UNA fila con "
                  f"ACRONIMO distinto — probablemente la misma persona registrada dos veces en Base cupos "
                  f"(ej. con y sin segundo nombre). Revisa la pestaña 4 y elimina/fusiona la fila sobrante.")

    except Exception as e:
        print(f"\n⚠ No se pudo generar el archivo de depuración de cupos: {e}", file=sys.stderr)
    # ─────────────────────────────────────────────────────────────────────────

    return df_filtrado.reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# MAPEOS DE LLAVES → ACRONIMO
# ─────────────────────────────────────────────────────────────────────────────

def construir_indices(df_bc: pd.DataFrame) -> dict:
    """
    Devuelve un dict con índices listos para hacer .map():
        cedula_a_acronimo : {CEDULA → ACRONIMO}
        codigo_a_acronimo : {COD_ASESOR_ECOM → ACRONIMO}
        nombre_a_acronimo : {NOMBRE_UPPER → ACRONIMO}   (último recurso)
        acronimo_a_nombre : {ACRONIMO → NOMBRE_UPPER}   (nombre canónico
                             para mostrar en salidas — así una misma persona
                             siempre se ve con el mismo nombre, sin importar
                             cómo la haya escrito cada fuente/módulo)
    """
    cedula_a_acr = (
        df_bc[df_bc["CEDULA"] != ""]
            .drop_duplicates("CEDULA", keep="first")
            .set_index("CEDULA")["ACRONIMO"]
            .to_dict()
    )
    codigo_a_acr = (
        df_bc[df_bc["COD_ASESOR_ECOM"] != ""]
            .drop_duplicates("COD_ASESOR_ECOM", keep="first")
            .set_index("COD_ASESOR_ECOM")["ACRONIMO"]
            .to_dict()
    )
    nombre_a_acr = (
        df_bc.drop_duplicates("NOMBRE", keep="first")
             .set_index("NOMBRE")["ACRONIMO"]
             .to_dict()
    )
    acr_a_nombre = (
        df_bc[df_bc["ACRONIMO"] != ""]
            .drop_duplicates("ACRONIMO", keep="first")
            .set_index("ACRONIMO")["NOMBRE"]
            .to_dict()
    )
    return {
        "cedula_a_acronimo": cedula_a_acr,
        "codigo_a_acronimo": codigo_a_acr,
        "nombre_a_acronimo": nombre_a_acr,
        "acronimo_a_nombre": acr_a_nombre,
    }


def resolver_acronimo(
    serie_in: pd.Series,
    idx: dict,
    estrategia: str,
    df_pt_cedula: pd.Series | None = None,
) -> pd.Series:
    """
    Resuelve un ACRONIMO canónico para cada valor de `serie_in`.

    Parámetros
    ──────────
    serie_in: la columna de origen (puede contener ACRONIMOs, cédulas,
              códigos asesor o nombres, según `estrategia`).
    idx:      output de `construir_indices`.
    estrategia: una de:
        "acronimo+cedula" — `serie_in` es ACRONIMO; si no matchea con la
                            maestra, intentar fallback por CEDULA usando
                            `df_pt_cedula` (Series alineada con serie_in).
        "codigo"          — `serie_in` es COD_ASESOR_ECOM (numérico string).
        "nombre"          — `serie_in` es NOMBRE (UPPER).

    Devuelve la misma serie con ACRONIMOs canónicos. Filas sin match
    quedan como string vacío "".
    """
    serie_in = serie_in.astype(str).str.strip().str.upper()

    if estrategia == "codigo":
        return serie_in.map(idx["codigo_a_acronimo"]).fillna("")

    if estrategia == "nombre":
        return serie_in.map(idx["nombre_a_acronimo"]).fillna("")

    if estrategia == "acronimo+cedula":
        acr_set = set(idx["cedula_a_acronimo"].values()) \
                  | set(idx["codigo_a_acronimo"].values()) \
                  | set(idx["nombre_a_acronimo"].values())
        # 1) Si el ACRONIMO ya existe en la maestra, usarlo tal cual.
        salida = serie_in.where(serie_in.isin(acr_set), other="")
        # 2) Para los que no matchean, fallback a CEDULA.
        if df_pt_cedula is not None:
            faltan = salida == ""
            cedulas = df_pt_cedula.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            salida.loc[faltan] = cedulas[faltan].map(idx["cedula_a_acronimo"]).fillna("")
        return salida

    raise ValueError(f"Estrategia desconocida: {estrategia}")


# ─────────────────────────────────────────────────────────────────────────────
# UNIVERSO DE PERSONAS PARA EL REPORTE
# ─────────────────────────────────────────────────────────────────────────────

def universo_personas(df_bc: pd.DataFrame) -> pd.DataFrame:
    """
    Devuelve las personas activas (ES_ACTIVO=True), una fila por ACRONIMO.

    Columnas: ACRONIMO, CEDULA, COD_ASESOR_ECOM, NOMBRE, ROL, CANAL,
              ES_SUPERVISOR.
    """
    cols = ["ACRONIMO", "CEDULA", "COD_ASESOR_ECOM", "NOMBRE",
            "ROL", "CANAL", "TIPO_SERVICIO",
            "ES_SUPERVISOR", "ES_GDD", "ES_LIDER",      # Sprint 17.15
            "CIUDAD_CUPO"]
    return (
        df_bc[df_bc["ES_ACTIVO"]][cols]
            .drop_duplicates(subset=["ACRONIMO"])
            .reset_index(drop=True)
    )


def supervisores(df_bc: pd.DataFrame) -> pd.DataFrame:
    """Sólo los supervisores activos (una fila por ACRONIMO)."""
    return universo_personas(df_bc).query("ES_SUPERVISOR == True").reset_index(drop=True)



# ─────────────────────────────────────────────────────────────────────────────
# DISPARADOR DE PRUEBA (Para ejecutar directamente) 
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("🚀 Iniciando prueba forzada de base_cupos.py...")
    try:
        # Llamamos a cargar(), lo que disparará la exportación del Excel de debug
        df_prueba = cargar()
        print(f"✅ Proceso terminado con éxito. Se cargaron {len(df_prueba)} registros.")
    except Exception as e:
        print(f"❌ Error al ejecutar el archivo: {e}")