"""
paths.py
────────
Resolución central de rutas para Fase 1 y Fase 2.

Soporta DOS layouts:

  • PROD (operación real en OneDrive de Eficacia):
      <BASE>/CIF/BASES/{PLAN DE TRABAJO,INVOLVES,EMPLEADOS,...}/
      <BASE>/CIF/SALIDA/
      <BASE>/NO PRESENCIA/{BASES,SALIDA}/
      <BASE>/PRECIOS/{BASES,SALIDA}/
      <BASE>/SOS/{BASES,SALIDA}/
      <BASE>/EXHIBICIONES/<NN. Mes>/    (subcarpeta del mes)
      <BASE>/EXHIBICIONES/SALIDA/
      <BASE>/ALERTAS/

  • DEV (checkout del proyecto sin estructura de OneDrive):
      <BASE>/BASES/CIF/{PLAN DE TRABAJO,INVOLVES,EMPLEADOS,...}/
      <BASE>/BASES/NO PRESENCIA/
      <BASE>/BASES/PRECIOS/
      <BASE>/BASES/SOS/
      <BASE>/BASES/EXHIBICIONES/        (plano, sin subcarpeta de mes)
      <BASE>/SALIDA/CIF/...              (auto-creadas si no existen)
      <BASE>/SALIDA/NO PRESENCIA/
      <BASE>/SALIDA/PRECIOS/
      <BASE>/SALIDA/SOS/
      <BASE>/SALIDA/EXHIBICIONES/
      <BASE>/ALERTAS/

Resolución de BASE
──────────────────
Orden de prioridad:
  1. Variable de entorno EFICACIA_BASE (override explícito).
  2. Archivo `EFICACIA_BASE.txt` en la raíz del proyecto, con la ruta como
     única línea (forma recomendada para el analista — no requiere tocar
     código ni variables de entorno del sistema).
  3. Ruta de PROD (`C:\\1\\OneDrive - Eficacia\\Escritorio\\ETLS`) si existe.
  4. Ruta de DEV (raíz del checkout = parent de SCRIPTS/).

Detección de layout
───────────────────
Se considera DEV si dentro de BASE existe una carpeta `BASES/`. En ese caso
los inputs se leen de `BASE/BASES/...` y los outputs van a `BASE/SALIDA/...`.
En PROD inputs van a `BASE/MODULE/BASES/...` y outputs a `BASE/MODULE/SALIDA/...`.
"""

import os
import re
import glob
from pathlib import Path


# ─────────────────────────────────────────────────────────────────────────────
# RESOLUCIÓN DE BASE Y LAYOUT
# ─────────────────────────────────────────────────────────────────────────────
_PROD_BASE = Path(r"C:\1\OneDrive - Eficacia\Escritorio\ETLS")
_DEV_BASE  = Path(__file__).resolve().parent.parent  # parent de SCRIPTS/
_OVERRIDE  = os.environ.get("EFICACIA_BASE", "").strip()
_BASE_FILE = _DEV_BASE / "EFICACIA_BASE.txt"


def _leer_base_file() -> str:
    """Lee la primera línea no vacía/no comentario de EFICACIA_BASE.txt."""
    if not _BASE_FILE.is_file():
        return ""
    try:
        for linea in _BASE_FILE.read_text(encoding="utf-8").splitlines():
            linea = linea.strip()
            if linea and not linea.startswith("#"):
                return linea
    except Exception:
        return ""
    return ""


def _resolver_base() -> tuple[Path, str]:
    # 1) Variable de entorno
    if _OVERRIDE:
        p = Path(_OVERRIDE)
        if p.is_dir():
            return p, "dev" if (p / "BASES").is_dir() else "prod"
    # 2) Archivo EFICACIA_BASE.txt en la raíz del proyecto
    desde_archivo = _leer_base_file()
    if desde_archivo:
        p = Path(desde_archivo)
        if p.is_dir():
            return p, "dev" if (p / "BASES").is_dir() else "prod"
    # 3) DEV (parent de SCRIPTS/) si tiene BASES/
    if _DEV_BASE.is_dir() and (_DEV_BASE / "BASES").is_dir():
        return _DEV_BASE, "dev"
    # 4) PROD por defecto
    if _PROD_BASE.is_dir():
        return _PROD_BASE, "prod"
    # Fallback final: producción aunque no exista (los chequeos de archivos
    # individuales reportarán claro el problema).
    return _PROD_BASE, "prod"


BASE, LAYOUT = _resolver_base()


# ─────────────────────────────────────────────────────────────────────────────
# DIRECTORIOS POR MÓDULO  (input + output)
# ─────────────────────────────────────────────────────────────────────────────

def _input_root(modulo: str) -> Path:
    """Raíz de inputs para un módulo (CIF, NO PRESENCIA, PRECIOS, SOS, EXHIBICIONES)."""
    if LAYOUT == "dev":
        return BASE / "BASES" / modulo
    return BASE / modulo / "BASES"


def _output_root(modulo: str) -> Path:
    """Raíz de outputs para un módulo (auto-creada si no existe)."""
    if LAYOUT == "dev":
        d = BASE / "SALIDA" / modulo
    else:
        d = BASE / modulo / "SALIDA"
    d.mkdir(parents=True, exist_ok=True)
    return d


# ── CIF ──────────────────────────────────────────────────────────────────────
CIF_BASES        = _input_root("CIF")
CIF_PT_DIR       = CIF_BASES / "PLAN DE TRABAJO"
CIF_INVOLVES     = CIF_BASES / "INVOLVES"   / "informe-gerencial-visitas.xlsx"
CIF_COLABORADORES= CIF_BASES / "EMPLEADOS"  / "Informe_Colaboradores.xlsx"
CIF_BASE_VENTAS  = CIF_BASES / "PUNTOS DE VENTA" / "BASE PUNTOS DE VENTA.xlsx"
CIF_CAUSALES     = CIF_BASES / "CAUSALES"   / "CAUSALES.xlsx"
CIF_SALIDA       = _output_root("CIF")
# Detalle PDV×persona — el cliente lo conoce como CIF.xlsx (V3); mantenemos
# el alias `Plan de trabajo.xlsx` por compatibilidad hacia atrás (D8=B).
CIF_OUT_CIF      = CIF_SALIDA / "CIF.xlsx"
CIF_OUT_FINAL    = CIF_SALIDA / "Plan de trabajo.xlsx"   # alias legacy
# Resumen agregado por gestor (paso 8 V3): mes activo + histórico acumulado (D6=C)
CIF_OUT_KPIS              = CIF_SALIDA / "KPIS_CIF.xlsx"
CIF_OUT_KPIS_HISTORICO    = CIF_SALIDA / "KPIS_CIF_HISTORICO.xlsx"
CIF_OUT_ALERTAS  = CIF_SALIDA / "ALERTAS_TIEMPOS_INCONSISTENTES.csv"
# CSVs intermedios (en la misma carpeta del PT, como en producción)
CIF_CONSOLIDADO_CSV  = CIF_PT_DIR / "Plan de trabajo consolidado.csv"
CIF_AGRUPADO_CSV     = CIF_PT_DIR / "Plan de trabajo consolidado_grupado.csv"
# Procesado de Involves (escrito por paso 3 y 4, leído por paso 5/6/7)
CIF_INVOLVES_PROCESADO = CIF_BASES / "INVOLVES" / "informe_visitas_procesado.csv"

# ── NO PRESENCIA ─────────────────────────────────────────────────────────────
NP_BASES                = _input_root("NO PRESENCIA")
NP_SALIDA               = _output_root("NO PRESENCIA")
NP_OUT_KPIS             = NP_SALIDA / "NO_PRESENCIA_KPIS.xlsx"
NP_OUT_KPIS_HISTORICO   = NP_SALIDA / "NO_PRESENCIA_KPIS_HISTORICO.xlsx"

# ── PRECIOS ──────────────────────────────────────────────────────────────────
PR_BASES                = _input_root("PRECIOS")
PR_SALIDA               = _output_root("PRECIOS")
PR_OUT_KPIS             = PR_SALIDA / "PRECIOS_KPIS.xlsx"
PR_OUT_KPIS_HISTORICO   = PR_SALIDA / "PRECIOS_KPIS_HISTORICO.xlsx"

# ── SOS ──────────────────────────────────────────────────────────────────────
SOS_BASES               = _input_root("SOS")
SOS_SALIDA              = _output_root("SOS")
SOS_OUT_KPIS            = SOS_SALIDA / "SOS_KPIS.xlsx"
SOS_OUT_KPIS_HISTORICO  = SOS_SALIDA / "SOS_KPIS_HISTORICO.xlsx"

# ── EXHIBICIONES ─────────────────────────────────────────────────────────────
def _resolver_exhib_data_dir() -> Path:
    """
    En dev devuelve BASE/BASES/EXHIBICIONES (plano).
    En prod busca la subcarpeta `NN. Mes` más reciente; si no hay, devuelve
    el directorio raíz de EXHIBICIONES.
    """
    raiz = _input_root("EXHIBICIONES")
    if LAYOUT == "dev":
        return raiz
    if not raiz.is_dir():
        return raiz
    candidatos = [
        raiz / d.name
        for d in raiz.iterdir()
        if d.is_dir() and re.match(r"^\d{2}\.\s", d.name)
    ]
    if candidatos:
        return max(candidatos, key=lambda p: p.stat().st_mtime)
    return raiz


EXHIB_DATA_DIR = _resolver_exhib_data_dir()
EXHIB_SALIDA   = _output_root("EXHIBICIONES")
EXHIB_NIVEL_IMPACTO = EXHIB_DATA_DIR / "Nivel impacto x Exhibición.xlsx"
# KPIs V3 — mes activo + histórico (D6=C)
EXHIB_PAG_OUT_KPIS              = EXHIB_SALIDA / "EXHIBICIONES_PAGADAS_KPIS.xlsx"
EXHIB_PAG_OUT_KPIS_HISTORICO    = EXHIB_SALIDA / "EXHIBICIONES_PAGADAS_KPIS_HISTORICO.xlsx"
EXHIB_GRA_OUT_KPIS              = EXHIB_SALIDA / "EXHIBICIONES_GRATIS_KPIS.xlsx"
EXHIB_GRA_OUT_KPIS_HISTORICO    = EXHIB_SALIDA / "EXHIBICIONES_GRATIS_KPIS_HISTORICO.xlsx"

# ── D&P  (Ventas + Impactos + Segmentos) ─────────────────────────────────────
# Inputs:
#   <BASE>/D&P/01. Ventas/<año>/<mes>/<archivo>.xlsx     (uno por corte mensual)
#   <BASE>/D&P/02. Impactos/<año>/<mes>/<archivo>.xlsx
#   <BASE>/D&P/Listas/listas_referencia.xlsx             (hojas MSL, ListSant,
#                                                         DoyPackBaby, CremasBaby)
#   <BASE>/D&P/Rutero/Rutero_Droguerias.xlsx             (hoja "PLAN DE TRABAJO")
DYP_BASES        = _input_root("D&P")
DYP_VENTAS_DIR   = DYP_BASES / "01. Ventas"
DYP_IMPACTOS_DIR = DYP_BASES / "02. Impactos"
# Sprint 14.2: nuevos nombres V2.
#   • MSL & Listas Target Catman.xlsx (4 segmentos nuevos + MSL)
#   • RUTERO <MES> <AÑO> D&P.xlsx (nombre cambia mes a mes; glob detecta el reciente)
DYP_LISTAS_FILE  = DYP_BASES / "Listas" / "MSL & Listas Target Catman.xlsx"

def _resolver_rutero_dyp() -> Path:
    """Detecta el rutero D&P más reciente (RUTERO*D&P.xlsx por mtime). Si no
    se encuentra ninguno, cae al nombre legacy 'Rutero_Droguerias.xlsx'."""
    rutero_dir = DYP_BASES / "Rutero"
    if rutero_dir.is_dir():
        # Patrones aceptables (mes a mes el nombre cambia)
        candidatos = list(rutero_dir.glob("RUTERO*.xlsx")) \
                   + list(rutero_dir.glob("Rutero*.xlsx"))
        # Excluir archivos temporales de Excel
        candidatos = [p for p in candidatos if not p.name.startswith("~$")]
        if candidatos:
            return max(candidatos, key=lambda p: p.stat().st_mtime)
    return rutero_dir / "Rutero_Droguerias.xlsx"

DYP_RUTERO_FILE  = _resolver_rutero_dyp()
# Tabla maestra de personas (ACRONIMO, CEDULA, COD_ASESOR_ECOM, NOMBRE, ROL, CANAL).
# Es la fuente de verdad para resolver identidades entre PT y D&P.
DYP_BASE_CUPOS   = DYP_BASES / "Base_cupos.xlsx"

DYP_SALIDA           = _output_root("D&P")
DYP_OUT_VENTAS       = DYP_SALIDA / "Consolidado_Ventas.csv"
DYP_OUT_IMPACTOS     = DYP_SALIDA / "Consolidado_Impactos.csv"
DYP_OUT_SEGMENTOS    = DYP_SALIDA / "impacto_segmentos.xlsx"

# ── ALERTAS ──────────────────────────────────────────────────────────────────
ALERTAS_DIR    = BASE / "ALERTAS"
ALERTAS_DIR.mkdir(parents=True, exist_ok=True)
ALERTAS_LOGS   = ALERTAS_DIR / "logs"
ALERTAS_ADJUNTOS = ALERTAS_DIR / "ADJUNTOS"
ALERTAS_MAESTRO  = ALERTAS_DIR / "MAESTRO_SUPERVISORES.xlsx"
ALERTAS_CONFIG_ENV = ALERTAS_DIR / "config.env"
ALERTAS_XLSM   = ALERTAS_DIR / "EnviarCorreos.xlsm"


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def descubrir_archivos_pt() -> tuple[str, str]:
    """
    Localiza los archivos de PT Directo e ISM más recientes en CIF_PT_DIR.
    Devuelve (ruta_directo, ruta_ism). Vacío si no encuentra.
    """
    if not CIF_PT_DIR.is_dir():
        return "", ""
    todos = [
        Path(p) for p in glob.glob(str(CIF_PT_DIR / "*.xlsx"))
        if not Path(p).name.startswith("~$")
        and "consolidado" not in Path(p).name.lower()
    ]
    candidatos_dir = [p for p in todos if "directo" in p.name.lower()]
    candidatos_ism = [p for p in todos if "ism"     in p.name.lower()]
    ruta_dir = str(max(candidatos_dir, key=lambda p: p.stat().st_mtime)) if candidatos_dir else ""
    ruta_ism = str(max(candidatos_ism, key=lambda p: p.stat().st_mtime)) if candidatos_ism else ""
    return ruta_dir, ruta_ism


def info() -> str:
    """Resumen de la resolución de paths para diagnóstico."""
    lineas = [
        f"BASE   : {BASE}",
        f"LAYOUT : {LAYOUT}",
        f"  CIF      → in: {CIF_BASES}",
        f"             out: {CIF_SALIDA}",
        f"  NP       → in: {NP_BASES}",
        f"             out: {NP_SALIDA}",
        f"  PRECIOS  → in: {PR_BASES}",
        f"             out: {PR_SALIDA}",
        f"  SOS      → in: {SOS_BASES}",
        f"             out: {SOS_SALIDA}",
        f"  EXHIB    → in: {EXHIB_DATA_DIR}",
        f"             out: {EXHIB_SALIDA}",
        f"  D&P      → in: {DYP_BASES}",
        f"             out: {DYP_SALIDA}",
        f"  ALERTAS  → {ALERTAS_DIR}",
    ]
    return "\n".join(lineas)


if __name__ == "__main__":
    print(info())
