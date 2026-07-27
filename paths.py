import os
import datetime
import re
import glob
from pathlib import Path

# Librerías para autenticación y servicios cloud
from dotenv import load_dotenv  # pip install python-dotenv
import msal                     # pip install msal requests
# from supabase import create_client, Client  # pip install supabase (descomentar si la usas)

# ─────────────────────────────────────────────────────────────────────────────
# 1. CARGA DE VARIABLES DE ENTORNO (.ENV)
# ─────────────────────────────────────────────────────────────────────────────
# Carga automáticamente las variables del archivo .env localizado en la raíz
load_dotenv()

# Credenciales de Azure AD
AZURE_TENANT_ID     = os.getenv("AZURE_TENANT_ID")
AZURE_CLIENT_ID     = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

# Credenciales de Supabase
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]


def obtener_token_azure() -> str:
    """
    Se conecta a Azure Active Directory usando las credenciales del .env
    y obtiene el Token de Acceso para Microsoft Graph / SharePoint.
    """
    if not all([AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET]):
        raise ValueError("Faltan credenciales de Azure en el archivo .env")

    app = msal.ConfidentialClientApplication(
        AZURE_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{AZURE_TENANT_ID}",
        client_credential=AZURE_CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Error al autenticar en Azure AD: {result.get('error_description')}")


# OPCIONAL: Cliente de Supabase inicializado (descomentar si usas la SDK oficial)
# def obtener_cliente_supabase() -> Client:
#     if not SUPABASE_URL or not SUPABASE_KEY:
#         raise ValueError("Faltan credenciales de Supabase en el archivo .env")
#     return create_client(SUPABASE_URL, SUPABASE_KEY)


# ─────────────────────────────────────────────────────────────────────────────
# 2. CONFIGURACIÓN CLOUD - SHAREPOINT (RUTAS EN LA NUBE)
# ─────────────────────────────────────────────────────────────────────────────
SHAREPOINT_SITE_NAME = "JJ451"
AÑO_ACTUAL = datetime.datetime.now().year

# Carpeta base en SharePoint extraída de tu enlace
SHAREPOINT_BASE_DIR = "Equipo Información/BI"

# Rutas dinámicas
RUTA_CARPETA_PT = f"{SHAREPOINT_BASE_DIR}/INVOLVES/PLAN DE TRABAJO" #OOKK
_BASES_ROOT     = f"{SHAREPOINT_BASE_DIR}/INVOLVES/BASES DE RESPUESTAS" #OOKKK
_SALIDAS_ROOT   = f"{SHAREPOINT_BASE_DIR}/INVOLVES/SALIDAS" #OOOKKK

# Rutas SharePoint por Módulo
RUTA_CARPETA_PT_CIF       = RUTA_CARPETA_PT  #OK
RUTA_CARPETA_BASES_CIF   = f"{_BASES_ROOT}/CIF/{AÑO_ACTUAL}" #OK
RUTA_CARPETA_SALIDAS_CIF = f"{_SALIDAS_ROOT}/CIF" #OK

RUTA_CARPETA_PT_PRECIOS      = RUTA_CARPETA_PT
RUTA_CARPETA_BASES_PRECIOS   = f"{_BASES_ROOT}/PRECIOS/{AÑO_ACTUAL}" #OOKK
RUTA_CARPETA_SALIDAS_PRECIOS = f"{_SALIDAS_ROOT}/PRECIOS"            #OOKKK

RUTA_CARPETA_PT_NP       = RUTA_CARPETA_PT
RUTA_CARPETA_BASES_NP    = f"{_BASES_ROOT}/NO PRESENCIA/{AÑO_ACTUAL}" #OK
RUTA_CARPETA_SALIDAS_NP  = f"{_SALIDAS_ROOT}/NO PRESENCIA"   #OK

RUTA_CARPETA_PT_SOS      = RUTA_CARPETA_PT
RUTA_CARPETA_BASES_SOS    = f"{_BASES_ROOT}/PARTICIPACIONES/{AÑO_ACTUAL}"  #OK
RUTA_CARPETA_SALIDAS_SOS  = f"{_SALIDAS_ROOT}/SOS"                         #OK             

RUTA_CARPETA_PT_EXHIB    = RUTA_CARPETA_PT
RUTA_CARPETA_BASES_EXHIB  = f"{_BASES_ROOT}/EXHIBICIONES/{AÑO_ACTUAL}"
RUTA_CARPETA_SALIDAS_EXHIB= f"{_SALIDAS_ROOT}/EXHIBICIONES"

RUTA_CARPETA_PT_DYP      = RUTA_CARPETA_PT
RUTA_CARPETA_BASES_DYP    = f"{_BASES_ROOT}/DYP/{AÑO_ACTUAL}"
RUTA_CARPETA_SALIDAS_DYP  = f"{_SALIDAS_ROOT}/DYP"

RUTA_CARPETA_BASES   = RUTA_CARPETA_BASES_PRECIOS
RUTA_CARPETA_SALIDAS = RUTA_CARPETA_SALIDAS_PRECIOS


# ─────────────────────────────────────────────────────────────────────────────
# 3. RESOLUCIÓN DE ENTORNO Y CARPETA LOCAL (FALLBACK DESARROLLO)
# ─────────────────────────────────────────────────────────────────────────────
_DEV_BASE  = Path(__file__).resolve().parent.parent
_OVERRIDE  = os.environ.get("EFICACIA_BASE", "").strip()
_BASE_FILE = _DEV_BASE / "EFICACIA_BASE.txt"


def _leer_base_file() -> str:
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
    if _OVERRIDE:
        p = Path(_OVERRIDE)
        if p.is_dir():
            return p, "dev" if (p / "BASES").is_dir() else "prod"

    desde_archivo = _leer_base_file()
    if desde_archivo:
        p = Path(desde_archivo)
        if p.is_dir():
            return p, "dev" if (p / "BASES").is_dir() else "prod"

    if _DEV_BASE.is_dir() and (_DEV_BASE / "BASES").is_dir():
        return _DEV_BASE, "dev"
    
    return _DEV_BASE, "dev"


BASE, LAYOUT = _resolver_base()


# ─────────────────────────────────────────────────────────────────────────────
# 4. DIRECTORIOS LOCALES POR MÓDULO
# ─────────────────────────────────────────────────────────────────────────────
def _input_root(modulo: str) -> Path:
    if LAYOUT == "dev":
        return BASE / "BASES" / modulo
    return BASE / modulo / "BASES"


def _output_root(modulo: str) -> Path:
    if LAYOUT == "dev":
        d = BASE / "SALIDA" / modulo
    else:
        d = BASE / modulo / "SALIDA"
    d.mkdir(parents=True, exist_ok=True)
    return d


# CIF
CIF_BASES          = _input_root("CIF")
CIF_PT_DIR         = CIF_BASES / "PLAN DE TRABAJO"
CIF_INVOLVES       = CIF_BASES  / "informe-gerencial-visitas.xlsx"
CIF_COLABORADORES  = CIF_BASES  / "Informe_Colaboradores.xlsx"
CIF_BASE_VENTAS    = CIF_BASES  / "BASE PUNTOS DE VENTA.xlsx"
CIF_CAUSALES       = CIF_BASES  / "CAUSALES.xlsx"
CIF_SALIDA         = _output_root("CIF")

CIF_OUT_CIF            = CIF_SALIDA / "CIF.xlsx"
CIF_OUT_FINAL          = CIF_SALIDA / "Plan de trabajo.xlsx"
CIF_OUT_KPIS           = CIF_SALIDA / "KPIS_CIF.xlsx"
CIF_OUT_KPIS_HISTORICO = CIF_SALIDA / "KPIS_CIF_HISTORICO.xlsx"
CIF_OUT_ALERTAS        = CIF_SALIDA / "ALERTAS_TIEMPOS_INCONSISTENTES.csv"

# NO PRESENCIA
NP_BASES              = _input_root("NO PRESENCIA")
NP_SALIDA             = _output_root("NO PRESENCIA")
NP_OUT_KPIS           = NP_SALIDA / "NO_PRESENCIA_KPIS.xlsx"
NP_OUT_KPIS_HISTORICO = NP_SALIDA / "NO_PRESENCIA_KPIS_HISTORICO.xlsx"

# PRECIOS
PR_BASES              = _input_root("PRECIOS")
PR_SALIDA             = _output_root("PRECIOS")
PR_OUT_KPIS           = PR_SALIDA / "PRECIOS_KPIS.xlsx"
PR_OUT_KPIS_HISTORICO = PR_SALIDA / "PRECIOS_KPIS_HISTORICO.xlsx"

# SOS
SOS_BASES             = _input_root("SOS")
SOS_SALIDA            = _output_root("SOS")
SOS_OUT_KPIS          = SOS_SALIDA / "SOS_KPIS.xlsx"
SOS_OUT_KPIS_HISTORICO= SOS_SALIDA / "SOS_KPIS_HISTORICO.xlsx"

# EXHIBICIONES
def _resolver_exhib_data_dir() -> Path:
    raiz = _input_root("EXHIBICIONES")
    if LAYOUT == "dev" or not raiz.is_dir():
        return raiz
    candidatos = [
        raiz / d.name
        for d in raiz.iterdir()
        if d.is_dir() and re.match(r"^\d{2}\.\s", d.name)
    ]
    if candidatos:
        return max(candidatos, key=lambda p: p.stat().st_mtime)
    return raiz

EXHIB_DATA_DIR      = _resolver_exhib_data_dir()
EXHIB_SALIDA        = _output_root("EXHIBICIONES")
EXHIB_NIVEL_IMPACTO = EXHIB_DATA_DIR / "Nivel impacto x Exhibición.xlsx"

EXHIB_PAG_OUT_KPIS           = EXHIB_SALIDA / "EXHIBICIONES_PAGADAS_KPIS.xlsx"
EXHIB_PAG_OUT_KPIS_HISTORICO = EXHIB_SALIDA / "EXHIBICIONES_PAGADAS_KPIS_HISTORICO.xlsx"
EXHIB_GRA_OUT_KPIS           = EXHIB_SALIDA / "EXHIBICIONES_GRATIS_KPIS.xlsx"
EXHIB_GRA_OUT_KPIS_HISTORICO = EXHIB_SALIDA / "EXHIBICIONES_GRATIS_KPIS_HISTORICO.xlsx"

# D&P
DYP_BASES        = _input_root("D&P")
DYP_VENTAS_DIR   = DYP_BASES / "01. Ventas"
DYP_IMPACTOS_DIR = DYP_BASES / "02. Impactos"
DYP_LISTAS_FILE  = DYP_BASES / "Listas" / "MSL & Listas Target Catman.xlsx"

def _resolver_rutero_dyp() -> Path:
    rutero_dir = DYP_BASES / "Rutero"
    if rutero_dir.is_dir():
        candidatos = list(rutero_dir.glob("RUTERO*.xlsx")) + list(rutero_dir.glob("Rutero*.xlsx"))
        candidatos = [p for p in candidatos if not p.name.startswith("~$")]
        if candidatos:
            return max(candidatos, key=lambda p: p.stat().st_mtime)
    return rutero_dir / "Rutero_Droguerias.xlsx"

DYP_RUTERO_FILE   = _resolver_rutero_dyp()
DYP_BASE_CUPOS    = DYP_BASES / "Base_cupos.xlsx"
DYP_SALIDA        = _output_root("D&P")
DYP_OUT_VENTAS    = DYP_SALIDA / "Consolidado_Ventas.csv"
DYP_OUT_IMPACTOS  = DYP_SALIDA / "Consolidado_Impactos.csv"
DYP_OUT_SEGMENTOS = DYP_SALIDA / "impacto_segmentos.xlsx"

# ALERTAS
ALERTAS_DIR        = BASE / "ALERTAS"
ALERTAS_DIR.mkdir(parents=True, exist_ok=True)
ALERTAS_LOGS       = ALERTAS_DIR / "logs"
ALERTAS_ADJUNTOS   = ALERTAS_DIR / "ADJUNTOS"
ALERTAS_MAESTRO    = ALERTAS_DIR / "MAESTRO_SUPERVISORES.xlsx"
ALERTAS_CONFIG_ENV = ALERTAS_DIR / "config.env"
ALERTAS_XLSM       = ALERTAS_DIR / "EnviarCorreos.xlsm"


# ─────────────────────────────────────────────────────────────────────────────
# DIAGNÓSTICO
# ─────────────────────────────────────────────────────────────────────────────
def info() -> str:
    lineas = [
        "=== VERIFICACIÓN DE CREDENCIALES ===",
        f"  AZURE TENANT ID : {'✓ Detectado' if AZURE_TENANT_ID else '✗ Falta'}",
        f"  AZURE CLIENT ID : {'✓ Detectado' if AZURE_CLIENT_ID else '✗ Falta'}",
        f"  SUPABASE URL    : {'✓ Detectada' if SUPABASE_URL else '✗ Falta'}",
        "",
        f"=== CLOUD SHAREPOINT (SITIO: {SHAREPOINT_SITE_NAME}) ===",
        f"  BASE GENERAL : {SHAREPOINT_BASE_DIR}",
        f"  PLAN TRABAJO : {RUTA_CARPETA_PT}",
        "",
        "--- RUTAS MÓDULOS SHAREPOINT ---",
        f"  PRECIOS → BASES: {RUTA_CARPETA_BASES_PRECIOS} | OUTS: {RUTA_CARPETA_SALIDAS_PRECIOS}",
        f"  CIF     → BASES: {RUTA_CARPETA_BASES_CIF} | OUTS: {RUTA_CARPETA_SALIDAS_CIF}",
        f"  NP      → BASES: {RUTA_CARPETA_BASES_NP} | OUTS: {RUTA_CARPETA_SALIDAS_NP}",
        f"  SOS     → BASES: {RUTA_CARPETA_BASES_SOS} | OUTS: {RUTA_CARPETA_SALIDAS_SOS}",
        f"  EXHIB   → BASES: {RUTA_CARPETA_BASES_EXHIB} | OUTS: {RUTA_CARPETA_SALIDAS_EXHIB}",
        f"  D&P     → BASES: {RUTA_CARPETA_BASES_DYP} | OUTS: {RUTA_CARPETA_SALIDAS_DYP}",
    ]
    return "\n".join(lineas)


if __name__ == "__main__":
    print(info())
    
    # Prueba opcional de conexión a Azure al ejecutar directamente
    try:
        token = obtener_token_azure()
        print("\n[SUCCESS] Conexión exitosa a Azure AD. Token obtenido correctamente.")
    except Exception as e:
        print(f"\n[ERROR] Falló la prueba de conexión a Azure: {e}")