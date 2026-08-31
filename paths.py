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

# ─────────────────────────────────────────────────────────────────────────────
# CARPETAS DE BASES POR AÑO
# ─────────────────────────────────────────────────────────────────────────────
# Las carpetas de BASES están organizadas por año. El año debe ser el del
# PERIODO que se está procesando, NO el del reloj:
#   • Cerrar diciembre durante enero → AÑO_ACTUAL da 2027, el dato está en 2026
#   • Recalcular un periodo anterior (backfill) → misma historia
# En ambos casos la lectura devuelve 404, el DataFrame sale vacío y las hojas
# quedan en blanco sin ningún error visible.
#
# Por eso las carpetas de BASES se exponen como FUNCIONES que reciben el año.
# Las constantes se conservan (apuntando al año actual) para no romper el
# código que ya las usa.

def _bases(modulo: str, anio: int | None = None) -> str:
    """Carpeta de BASES de un módulo para el año indicado."""
    return f"{_BASES_ROOT}/{modulo}/{anio or AÑO_ACTUAL}"


def bases_cif(anio: int | None = None) -> str:     return _bases("CIF", anio)
def bases_precios(anio: int | None = None) -> str: return _bases("PRECIOS", anio)
def bases_np(anio: int | None = None) -> str:      return _bases("NO PRESENCIA", anio)
def bases_sos(anio: int | None = None) -> str:     return _bases("PARTICIPACIONES", anio)
def bases_exhib(anio: int | None = None) -> str:   return _bases("EXHIBICIONES", anio)
def bases_dyp(anio: int | None = None) -> str:     return _bases("DYP", anio)


# Rutas SharePoint por Módulo
RUTA_CARPETA_PT_CIF       = RUTA_CARPETA_PT  #OK
RUTA_CARPETA_BASES_CIF   = bases_cif() #OK
RUTA_CARPETA_SALIDAS_CIF = f"{_SALIDAS_ROOT}/CIF" #OK

RUTA_CARPETA_PT_PRECIOS      = RUTA_CARPETA_PT
RUTA_CARPETA_BASES_PRECIOS   = bases_precios() #OOKK
RUTA_CARPETA_SALIDAS_PRECIOS = f"{_SALIDAS_ROOT}/PRECIOS"            #OOKKK

RUTA_CARPETA_PT_NP       = RUTA_CARPETA_PT
RUTA_CARPETA_BASES_NP    = bases_np() #OK
RUTA_CARPETA_SALIDAS_NP  = f"{_SALIDAS_ROOT}/NO PRESENCIA"   #OK

RUTA_CARPETA_PT_SOS      = RUTA_CARPETA_PT
RUTA_CARPETA_BASES_SOS    = bases_sos()  #OK
RUTA_CARPETA_SALIDAS_SOS  = f"{_SALIDAS_ROOT}/SOS"                         #OK             

RUTA_CARPETA_PT_EXHIB    = RUTA_CARPETA_PT
RUTA_CARPETA_BASES_EXHIB  = bases_exhib()
RUTA_CARPETA_SALIDAS_EXHIB= f"{_SALIDAS_ROOT}/EXHIBICIONES"

RUTA_CARPETA_PT_DYP      = RUTA_CARPETA_PT
RUTA_CARPETA_BASES_DYP    = bases_dyp()
RUTA_CARPETA_SALIDAS_DYP  = f"{_SALIDAS_ROOT}/DYP"


# ─────────────────────────────────────────────────────────────────────────────
# AÑO DEL PERIODO vs AÑO DEL RELOJ
# ─────────────────────────────────────────────────────────────────────────────
# Las constantes de arriba se construyen con AÑO_ACTUAL, o sea el año en que
# se ejecuta el proceso. Eso funciona mientras se procese el mes en curso, pero
# falla en dos casos reales y SIN dar error:
#
#   • Cerrar diciembre durante enero → busca en .../2027 cuando el dato está
#     en .../2026. La lectura devuelve 404, el DataFrame sale vacío y las
#     hojas del reporte quedan en blanco.
#   • Reprocesar un periodo de un año anterior → lo mismo.
#
# `usar_anio()` reapunta las constantes al año del periodo que se va a
# procesar. Cada ETL la llama una vez, al empezar con un periodo.
#
# Es seguro con la ejecución en paralelo de run_all.py: los ETLs corren en
# paralelo DENTRO de un periodo (todos quieren el mismo año) y los periodos
# se procesan de forma secuencial.

def carpeta_periodo(raiz: str, anio: int | None) -> str:
    """
    Devuelve `raiz` apuntando al año indicado.

    Si la ruta ya termina en un año lo REEMPLAZA (no lo anida, que produciría
    .../2026/2026); si no lo trae, lo agrega. Con anio=None no toca nada.
    """
    raiz = str(raiz).rstrip("/")
    if not anio:
        return raiz
    anio_str = str(int(anio))
    if raiz.endswith(f"/{anio_str}"):
        return raiz
    if re.search(r"/(19|20)\d{2}$", raiz):
        return re.sub(r"/(19|20)\d{2}$", f"/{anio_str}", raiz)
    return f"{raiz}/{anio_str}"


def usar_anio(anio: int) -> None:
    """
    Reapunta las carpetas de BASES al año del periodo que se está procesando.

    Llamar al inicio de cada periodo, antes de leer nada:

        paths.usar_anio(spec.anio)
    """
    global RUTA_CARPETA_BASES_CIF, RUTA_CARPETA_BASES_PRECIOS
    global RUTA_CARPETA_BASES_NP, RUTA_CARPETA_BASES_SOS
    global RUTA_CARPETA_BASES_EXHIB, RUTA_CARPETA_BASES_DYP
    global RUTA_CARPETA_BASES, RUTA_CARPETA_SALIDAS

    RUTA_CARPETA_BASES_CIF     = bases_cif(anio)
    RUTA_CARPETA_BASES_PRECIOS = bases_precios(anio)
    RUTA_CARPETA_BASES_NP      = bases_np(anio)
    RUTA_CARPETA_BASES_SOS     = bases_sos(anio)
    RUTA_CARPETA_BASES_EXHIB   = bases_exhib(anio)
    RUTA_CARPETA_BASES_DYP     = bases_dyp(anio)

    # Alias que usa el ETL de Precios (RUTA_CARPETA_BASES/SALIDAS sin sufijo).
    RUTA_CARPETA_BASES   = RUTA_CARPETA_BASES_PRECIOS
    RUTA_CARPETA_SALIDAS = RUTA_CARPETA_SALIDAS_PRECIOS

    if int(anio) != AÑO_ACTUAL:
        print(f"  ℹ️  Carpetas de BASES apuntando al año del periodo ({anio}), "
              f"no al del reloj ({AÑO_ACTUAL}).")


# ─────────────────────────────────────────────────────────────────────────────
# INSUMOS EN LA NUBE CON MES Y AÑO EN EL NOMBRE
# ─────────────────────────────────────────────────────────────────────────────
# Punto único de verdad para los archivos cuyo nombre depende del periodo.
# Antes cada ETL los armaba por su cuenta, con el año a veces fijo.
_MESES_ES = {1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO",
             6: "JUNIO", 7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE",
             10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"}


def mes_es(mes: int, mayus: bool = True) -> str:
    """7 → 'JULIO' (o 'Julio' con mayus=False)."""
    m = _MESES_ES[int(mes)]
    return m if mayus else m.capitalize()


# --- CIF ---------------------------------------------------------------------
def cloud_cif_visitas(mes: int, anio: int) -> str:
    """BASES DE RESPUESTAS/CIF/<año>/Involves_Procesado_<MM>_<AAAA>.csv"""
    return f"{bases_cif(anio)}/Involves_Procesado_{int(mes):02d}_{anio}.csv"


# --- EXHIBICIONES ------------------------------------------------------------
def cloud_exhib_planning(mes: int, anio: int) -> str:
    """BASES DE RESPUESTAS/EXHIBICIONES/<año>/PLANNING DE <MES> <AÑO>.xlsx"""
    return f"{bases_exhib(anio)}/PLANNING DE {mes_es(mes)} {anio}.xlsx"


def cloud_exhib_base_planning(mes: int, anio: int) -> str:
    """BASES DE RESPUESTAS/EXHIBICIONES/<año>/Base Exhibiciones Planning <Mes> <Año>.xlsx"""
    return f"{bases_exhib(anio)}/Base Exhibiciones Planning {mes_es(mes, False)} {anio}.xlsx"


# --- D&P ---------------------------------------------------------------------
# Nota: el "&" de "MSL & Listas Target Catman.xlsx" NO necesita escaparse
# aparte; urllib.parse.quote() lo codifica como %26 al armar la URL de Graph.
# Los dos consolidados de D&P son archivos ÚNICOS que acumulan todos los meses:
# cada corrida quita las filas de su periodo y vuelve a escribirlas (upsert).
#
# Antes el nombre llevaba el periodo (`Consolidado_Ventas_AGOSTO_2026.csv`) pero
# el CONTENIDO era igual de acumulado, así que cada mes dejaba una copia entera
# del histórico bajo su propio nombre: seis archivos de ~130 MB con los mismos
# datos y, peor, contradiciéndose entre sí (el de MARZO decía 76.522 filas de
# marzo y el de AGOSTO decía 64.735, porque cada uno era una foto de un momento
# distinto). Con un solo archivo hay una sola verdad.
#
# `mes` y `anio` se siguen recibiendo para no romper a quien ya llama estas
# funciones con el periodo; simplemente no se usan en el nombre.
def cloud_dyp_impactos(mes: int | None = None, anio: int | None = None) -> str:
    """SALIDAS/DYP/Consolidado_Impactos.csv (único, acumula todos los meses)"""
    return f"{RUTA_CARPETA_SALIDAS_DYP}/Consolidado_Impactos.csv"


def cloud_dyp_ventas(mes: int | None = None, anio: int | None = None) -> str:
    """SALIDAS/DYP/Consolidado_Ventas.csv (único, acumula todos los meses)"""
    return f"{RUTA_CARPETA_SALIDAS_DYP}/Consolidado_Ventas.csv"


def cloud_dyp_rutero(mes: int, anio: int) -> str:
    """
    BASES DE RESPUESTAS/DYP/<año>/Rutero/RUTERO <MES> <AÑO> D&P.xlsx

    Nombre confirmado: lleva ESPACIOS y "&", no guiones bajos, y cambia cada
    mes. El "&" no necesita escaparse: urllib.parse.quote() lo codifica %26.
    """
    return f"{bases_dyp(anio)}/Rutero/RUTERO {mes_es(mes)} {anio} D&P.xlsx"


def cloud_dyp_listas(anio: int | None = None) -> str:
    """BASES DE RESPUESTAS/DYP/<año>/Listas/MSL & Listas Target Catman.xlsx"""
    return f"{bases_dyp(anio)}/Listas/MSL & Listas Target Catman.xlsx"


def cloud_dyp_base_cupos(anio: int | None = None) -> str:
    """BASES DE RESPUESTAS/DYP/<año>/Base_cupos.xlsx"""
    return f"{bases_dyp(anio)}/Base_cupos.xlsx"


# ─────────────────────────────────────────────────────────────────────────────
# INSUMOS TRANSVERSALES (carpetas propias, fuera de BASES DE RESPUESTAS)
# ─────────────────────────────────────────────────────────────────────────────
# Estos 3 insumos son ÚNICOS para todo el proyecto: los publica Involves una
# sola vez al mes y los consumen varios módulos. Viven en carpetas propias al
# mismo nivel que BASES DE RESPUESTAS / SALIDAS / PLAN DE TRABAJO.
#
# Antes cada módulo los buscaba dentro de BASES DE RESPUESTAS/CIF/<año> (y el
# de visitas también en BASES DE RESPUESTAS/VISITAS/<año>), lo que obligaba a
# duplicar el mismo archivo en 2 o 3 carpetas cada mes. Estas rutas apuntan al
# original, para que se cargue una sola vez.
#
# OJO: estas carpetas NO tienen subcarpeta por año — todos los meses conviven
# sueltos. Por eso el periodo va SIEMPRE en el nombre del archivo, y quien lo
# busque debe exigirlo (no vale agarrar "el primero que aparezca").
RUTA_CARPETA_INVOLVES   = f"{SHAREPOINT_BASE_DIR}/INVOLVES/INVOLVES"
RUTA_CARPETA_PUNTOS_VTA = f"{SHAREPOINT_BASE_DIR}/INVOLVES/PUNTOS DE VENTA"
RUTA_CARPETA_EMPLEADOS  = f"{SHAREPOINT_BASE_DIR}/INVOLVES/EMPLEADOS"


def cloud_involves_visitas(mes: int, anio: int) -> str:
    """INVOLVES/INVOLVES/informe-gerencial-visitas_<AAAAMM>.xlsx"""
    return f"{RUTA_CARPETA_INVOLVES}/informe-gerencial-visitas_{int(anio)}{int(mes):02d}.xlsx"


def cloud_puntos_venta(mes: int, anio: int) -> str:
    """PUNTOS DE VENTA/BASE PUNTOS DE VENTA <MES> <AÑO>.xlsx"""
    return f"{RUTA_CARPETA_PUNTOS_VTA}/BASE PUNTOS DE VENTA {mes_es(mes)} {int(anio)}.xlsx"


def cloud_empleados(mes: int, anio: int) -> str:
    """EMPLEADOS/Informe_Colaboradores_<MES>_<AÑO>.xlsx"""
    return f"{RUTA_CARPETA_EMPLEADOS}/Informe_Colaboradores_{mes_es(mes)}_{int(anio)}.xlsx"

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
# Nombre confirmado del archivo que publica el ETL de exhibiciones gratis en
# SALIDAS/EXHIBICIONES. Antes decía EXHIBICIONES_GRATIS_KPIS.xlsx, que no
# existe: la lectura daba 404 y ese KPI llegaba vacío al cruce.
EXHIB_GRA_OUT_KPIS           = EXHIB_SALIDA / "KPI_Exhibiciones_Gratis.xlsx"
EXHIB_GRA_OUT_KPIS_HISTORICO = EXHIB_SALIDA / "KPI_Exhibiciones_Gratis_HISTORICO.xlsx"

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