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

import sys
from pathlib import Path

import numpy as np
import pandas as pd

# Asegurar SCRIPTS/ en sys.path (paths.py vive ahí)
_SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "SCRIPTS"
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

import paths


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
    """
    if not paths.DYP_BASE_CUPOS.is_file():
        raise FileNotFoundError(
            f"No existe la tabla maestra de personas: {paths.DYP_BASE_CUPOS}\n"
            "Asegúrate de tenerla en BASES/D&P/Base_cupos.xlsx (hoja "
            "'Tabla total roles')."
        )

    df = pd.read_excel(paths.DYP_BASE_CUPOS, sheet_name="Tabla total roles")
    df = df.rename(columns={
        "COD PT":         "COD_PT",
        "TIPO SERVICIO":  "TIPO_SERVICIO",
        "CIUDAD CUPO":    "CIUDAD_CUPO",
        "ROL EN ACTIVO":  "ROL",
    })

    # Normalización de strings
    for c in ["ACRONIMO", "COD_PT", "TIPO_SERVICIO", "CANAL",
              "CIUDAD_CUPO", "ROL", "NOMBRE"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip().replace({"nan": ""})
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

    # Defensivo: descartar filas sin ACRONIMO o sin NOMBRE
    df = df[(df["ACRONIMO"] != "") & (df["NOMBRE"] != "")].copy()

    df["ES_SUPERVISOR"] = df["ROL"].apply(_es_supervisor)
    df["ES_GDD"]        = df["ROL"].apply(_es_gdd)       # Sprint 17.15
    df["ES_LIDER"]      = df["ROL"].apply(_es_lider)     # Sprint 17.15
    df["ES_ACTIVO"]     = ~df["ROL"].isin(ROLES_EXCLUIDOS)

    return df.reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# MAPEOS DE LLAVES → ACRONIMO
# ─────────────────────────────────────────────────────────────────────────────

def construir_indices(df_bc: pd.DataFrame) -> dict:
    """
    Devuelve un dict con tres índices listos para hacer .map():
        cedula_a_acronimo : {CEDULA → ACRONIMO}
        codigo_a_acronimo : {COD_ASESOR_ECOM → ACRONIMO}
        nombre_a_acronimo : {NOMBRE_UPPER → ACRONIMO}   (último recurso)
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
    return {
        "cedula_a_acronimo": cedula_a_acr,
        "codigo_a_acronimo": codigo_a_acr,
        "nombre_a_acronimo": nombre_a_acr,
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
