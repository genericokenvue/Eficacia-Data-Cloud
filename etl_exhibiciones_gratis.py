import os
import glob
from pathlib import Path
from io import BytesIO
import pandas as pd
import numpy as np
import urllib.parse
import requests
import msal
from dotenv import load_dotenv

import paths
import periodo_resolver as pr
import supabase_io

# ==============================================================================
# 0. CONFIGURACIÓN ENTORNO NUBE — SHAREPOINT (MICROSOFT GRAPH API)
# ==============================================================================
BASE_DIR = Path(__file__).resolve().parent
env_path = BASE_DIR / '.env'

if env_path.exists():
    load_dotenv(dotenv_path=env_path, override=True)
else:
    env_path_alt = BASE_DIR.parent / '.env'
    if env_path_alt.exists():
        load_dotenv(dotenv_path=env_path_alt, override=True)

AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

if SUPABASE_URL and SUPABASE_KEY:
    try:
        from supabase import create_client
        supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
    except ImportError:
        supabase = None
else:
    supabase = None

SHAREPOINT_BASE_DIR = os.getenv("SHAREPOINT_BASE_DIR", "Equipo Información/BI/INVOLVES").replace("\\", "/")
_BASES_ROOT = os.getenv("BASES_ROOT", f"{SHAREPOINT_BASE_DIR}/BASES DE RESPUESTAS").replace("\\", "/")
_SALIDAS_ROOT = os.getenv("SALIDAS_ROOT", f"{SHAREPOINT_BASE_DIR}/SALIDAS").replace("\\", "/")
AÑO_ACTUAL = pd.Timestamp.now().year

RUTA_CARPETA_PT = f"{SHAREPOINT_BASE_DIR}/PLAN DE TRABAJO".replace("\\", "/")
RUTA_CARPETA_BASES_EXHIB = f"{_BASES_ROOT}/EXHIBICIONES/{AÑO_ACTUAL}".replace("\\", "/")
RUTA_CARPETA_SALIDAS_EXHIB = f"{_SALIDAS_ROOT}/EXHIBICIONES".replace("\\", "/")

paths.EXHIB_SALIDA = Path(f"{RUTA_CARPETA_BASES_EXHIB}".replace("\\", "/"))


def _apuntar_rutas_al_periodo(spec: pr.PeriodoSpec) -> None:
    """
    Reapunta al año del periodo las rutas que dependen de una carpeta con año.

    Estas dos se fijaban al importar el módulo, usando AÑO_ACTUAL (el año del
    reloj). Al reprocesar un periodo de otro año se buscaban en la carpeta
    equivocada y la lectura fallaba sin explicar por qué.
    """
    global RUTA_CARPETA_BASES_EXHIB
    RUTA_CARPETA_BASES_EXHIB = f"{_BASES_ROOT}/EXHIBICIONES/{spec.anio}".replace("\\", "/")
    paths.EXHIB_SALIDA        = Path(RUTA_CARPETA_BASES_EXHIB)
    paths.EXHIB_NIVEL_IMPACTO = Path(
        f"{_BASES_ROOT}/EXHIBICIONES/{spec.anio}/Nivel impacto x Exhibición.xlsx".replace("\\", "/"))
    paths.DYP_BASE_CUPOS      = Path(
        f"{_BASES_ROOT}/DYP/{spec.anio}/Base_cupos.xlsx".replace("\\", "/"))


# Valores iniciales (año del reloj). `run()` los reapunta al periodo real.
paths.EXHIB_NIVEL_IMPACTO = Path(f"{_BASES_ROOT}/EXHIBICIONES/{AÑO_ACTUAL}/Nivel impacto x Exhibición.xlsx".replace("\\", "/"))
paths.DYP_BASE_CUPOS = Path(f"{_BASES_ROOT}/DYP/{AÑO_ACTUAL}/Base_cupos.xlsx".replace("\\", "/"))

# ==============================================================================
# 1. CONFIGURACIÓN DE COLUMNAS Y REGLAS
# ==============================================================================
OUTPUT_PATH = str(Path(RUTA_CARPETA_SALIDAS_EXHIB) / "Resultado exhibiciones gratis.xlsx").replace("\\", "/")

COL_ID_PDV          = "ID del PDV"
COL_EMPLEADO        = "Empleado"
COL_PERFIL_EMP      = "Perfil de acceso"
COL_FECHA           = "Fecha de la encuesta"
COL_MES             = "Mes del año"
COL_ANIO            = "Año"
COL_EXHIBICIONES    = "EXHIBICIONES:"

COL_TIPO_EXHIB_GC   = "Seleccionar el Tipo de la exhibicion:"
COL_MARCA_GC        = "MARCA"
COL_CANTIDAD_GC     = "*Digite el numero de exhibiciones adicionales para este tipo"

COL_TIPO_EXHIB_PAG  = "Seleccionar el Tipo de la exhibicion (Pagadas)"
COL_MARCA_PAG       = "MARCA.1"
COL_CANTIDAD_PAG    = "*Digite el numero de exhibiciones adicionales para este tipo."
COL_CAUSAL          = "Indique las causales:"

COL_TIPO_CONTRA     = "Seleccionar el Tipo de la exhibicion - CONTRAPRESTACIÓN"
COL_MARCA_CONTRA    = "MARCA - CONTRAPRESTACIÓN"
COL_CANTIDAD_CONTRA = "*Digite el numero de exhibiciones adicionales para este tipo. - CONTRAPRESTACIÓN"

COL_PLAN_ID_PDV     = "ID_PDV_INVOLVES"
COL_PLAN_ROL        = "ROL"
COL_PLAN_FREC       = "CANTIDAD_VISITAS" 
COL_PLAN_MES        = "MES"   
COL_PLAN_ANIO       = "AÑO"   

COL_VIS_ID_PDV      = "ID del PDV"
COL_VIS_EMPLEADO    = "Empleado"
COL_VIS_FECHA       = "Fecha de la visita"

EXHIB_PAGADA        = "Exhibiciones Pagadas"
EXHIB_GRATIS        = "Exhibiciones Gratis"
EXHIB_CONCURSO      = "Exhibiciones Concurso"
CAUSAL_CONTRA       = "contraprestación a la negociación"

PERFIL_PLAN         = "GESTOR"
PERFIL_A_ROL = {
    "GESTOR": "GESTOR",
    "SUPERVISOR": "SUPERVISOR",
    "GENERADOR DE DEMANDA": "GENERADOR DE DEMANDA",
}

OUTPUT_COLS = ["Mes", "Año", "Mes-Año", "ID PDV", "Categoría", "Tipo Exhibición", "Nivel Impacto", "Marca", "Cantidad", "Empleado", "Rol Empleado"]

# ==============================================================================
# 2. CONEXIÓN CLOUD MICROSOFT GRAPH API (SHAREPOINT)
# ==============================================================================

import sys
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

def _obtener_token_azure():
    authority = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
    scopes = ["https://graph.microsoft.com/.default"]
    app = msal.ConfidentialClientApplication(
        AZURE_CLIENT_ID, 
        authority=authority, 
        client_credential=AZURE_CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=scopes)
    if "access_token" in result:
        return result["access_token"]
    raise Exception(f"❌ Error de autenticación en Azure AD: {result.get('error_description')}")

def _obtener_site_id(headers):
    site_name = getattr(paths, "SHAREPOINT_SITE_NAME", "eficacia")
    url = f"https://graph.microsoft.com/v1.0/sites/root:/sites/{site_name}"
    res = requests.get(url, headers=headers).json()
    if "id" not in res:
        url_alt = f"https://graph.microsoft.com/v1.0/sites/{site_name}"
        res_alt = requests.get(url_alt, headers=headers).json()
        if "id" in res_alt:
            return res_alt["id"]
        raise Exception(f"❌ No se pudo resolver el Site ID de SharePoint. Respuesta: {res}")
    return res["id"]

def _get_sharepoint_file_stream(sharepoint_path: str) -> BytesIO:
    token = _obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    site_id = _obtener_site_id(headers)
    
    ruta_limpia = str(sharepoint_path).replace("\\", "/").strip("/")
    ruta_formateada = urllib.parse.quote(ruta_limpia)
    
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_formateada}:/content"
    response = requests.get(url, headers=headers)
    
    if response.status_code != 200:
        raise FileNotFoundError(f"❌ No se pudo descargar el archivo desde SharePoint ({sharepoint_path}). Estado HTTP {response.status_code}: {response.text}")
        
    return BytesIO(response.content)

def _upload_sharepoint_file(sharepoint_path: str, buffer_data: BytesIO):
    token = _obtener_token_azure()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
    site_id = _obtener_site_id(headers)
    
    ruta_limpia = str(sharepoint_path).replace("\\", "/").strip("/")
    ruta_formateada = urllib.parse.quote(ruta_limpia)
    
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_formateada}:/content"
    response = requests.put(url, headers=headers, data=buffer_data.getvalue())
    
    if response.status_code not in [200, 201]:
        raise Exception(f"❌ Error al subir archivo a SharePoint ({sharepoint_path}). Estado HTTP {response.status_code}: {response.text}")

def _semana_del_mes(fecha) -> int:
    if pd.isnull(fecha): return -1
    primer_dia = fecha.replace(day=1)
    offset = primer_dia.weekday()
    return (fecha.day + offset - 1) // 7 + 1

def resolve_files(spec: pr.PeriodoSpec) -> dict:
    groups: dict[str, list[str]] = {}

    carpeta_bases_cloud = f"{_BASES_ROOT}/EXHIBICIONES/{spec.anio}".replace("\\", "/")
    nombre_informar = f"Base Exhibiciones Informar {spec.etiqueta.replace('_', ' ')}.xlsx"
    nombre_planning = f"Base Exhibiciones Planning {spec.etiqueta.replace('_', ' ')}.xlsx"

    p_informar = f"{carpeta_bases_cloud}/{nombre_informar}".replace("\\", "/")
    p_planning = f"{carpeta_bases_cloud}/{nombre_planning}".replace("\\", "/")

    try:
        # Carpeta transversal INVOLVES/INVOLVES: el informe de visitas se carga
        # UNA sola vez y lo consumen CIF y este módulo. Antes se leía de
        # BASES DE RESPUESTAS/VISITAS/<año>, lo que obligaba a duplicar el
        # mismo archivo en dos carpetas cada mes.
        p_visitas = paths.cloud_involves_visitas(spec.mes, spec.anio)
    except Exception:
        p_visitas = ""

    nombre_mes_anio = spec.etiqueta.replace('_', ' ')
    p_plan_directo = f"{SHAREPOINT_BASE_DIR}/PLAN DE TRABAJO/Plan de trabajo Directo {nombre_mes_anio}.xlsx".replace("\\", "/")
    p_plan_ism = f"{SHAREPOINT_BASE_DIR}/PLAN DE TRABAJO/Plan de trabajo ISM {nombre_mes_anio}.xlsx".replace("\\", "/")

    groups["informar"]     = [p_informar] if p_informar else []
    groups["planning"]     = [p_planning] if p_planning else []
    groups["visitas"]      = [p_visitas] if p_visitas and p_visitas != "None" else []
    groups["plan_trabajo"] = [p_plan_directo, p_plan_ism]

    print(f"\n☁️ [SHAREPOINT GRAPH API] Rutas aplicadas para el periodo {spec.etiqueta}:")
    for tipo, archivos in groups.items():
        print(f"  [{tipo}] {len(archivos)} archivo(s):")
        for a in archivos:
            print(f"    - Ruta: {a}")
            
    return groups

def load_encuestas(files: dict) -> pd.DataFrame:
    dfs = []
    rutas_a_probar = files.get("informar", []) + files.get("planning", [])
    
    for path in rutas_a_probar:
        try:
            stream_data = _get_sharepoint_file_stream(path)
            df = pd.read_excel(stream_data, sheet_name="report")
            print(f"    ✓ [SharePoint] {os.path.basename(str(path))} — {len(df):,} filas")
            dfs.append(df)
        except Exception as e:
            print(f"    ⚠️ No se pudo cargar el archivo {os.path.basename(str(path))}: {e}")
            
    if not dfs: raise ValueError("No se pudieron cargar archivos de encuesta desde SharePoint.")
    cols_comunes = set(dfs[0].columns)
    for df in dfs[1:]: cols_comunes &= set(df.columns)
    cols_comunes = [c for c in dfs[0].columns if c in cols_comunes]
    df = pd.concat([d[cols_comunes] for d in dfs], ignore_index=True)

    # FIX: "Informar" y "Planning" son, por diseño, dos encuestas SIN solape
    # (Informar = Gratis/Concurso, Planning = Pagadas) — así fue en marzo,
    # abril, mayo y julio 2026 (0% de IDs en común). En agosto 2026 el export
    # de "Planning" cambió de forma y empezó a incluir TAMBIÉN las mismas
    # capturas de "Informar" (100% de sus 2.061 filas repetidas, byte a byte:
    # mismo ID de encuesta, categoría y cantidad), además de sus propias 2.353
    # filas nuevas de Pagadas. Sin deduplicar, esa captura repetida se contaba
    # dos veces antes incluso de llegar al merge con el Plan de Trabajo — el
    # PDV 9236 pasó de 24 (valor real) a 48, y de ahí el otro bug del merge
    # (filas duplicadas del plan) lo llevaba a 192.
    #
    # Deduplicar por "ID de la encuesta" es seguro en los dos escenarios: si
    # no hay solape (marzo-julio) no cambia nada; si lo hay (agosto) deja una
    # sola copia de cada captura real.
    if "ID de la encuesta" in df.columns:
        n_antes = len(df)
        df = df.drop_duplicates(subset="ID de la encuesta")
        if len(df) != n_antes:
            print(f"    ℹ️  Encuestas: {n_antes:,} filas → {len(df):,} tras deduplicar por "
                  f"'ID de la encuesta' (Informar/Planning se solapaban este periodo).")

    if "Sub canal" in df.columns:
        n_antes = len(df)
        df = df[df["Sub canal"].astype(str).str.strip().str.upper() != "DISCOUNTER"].copy()
        print(f"    🛡️ Filtro DISCOUNTER aplicado en encuestas: omitidas {n_antes - len(df):,} filas.")
    elif "SUB CANAL" in df.columns:
        n_antes = len(df)
        df = df[df["SUB CANAL"].astype(str).str.strip().str.upper() != "DISCOUNTER"].copy()
        print(f"    🛡️ Filtro SUB CANAL aplicado en encuestas: omitidas {n_antes - len(df):,} filas.")

    df[COL_FECHA] = pd.to_datetime(df[COL_FECHA], dayfirst=True, errors="coerce")
    df[COL_MES] = df[COL_FECHA].dt.month
    df[COL_ANIO] = df[COL_FECHA].dt.year
    df["_semana_mes"] = df[COL_FECHA].apply(_semana_del_mes)
    return df

def load_plan(files: dict, spec: pr.PeriodoSpec) -> pd.DataFrame:
    dfs = []
    for path in files.get("plan_trabajo", []):
        try:
            stream_data = _get_sharepoint_file_stream(path)
            df = pd.read_excel(stream_data)   
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            if "SUB CANAL" in df.columns:
                df = df[df["SUB CANAL"].astype(str).str.strip().str.upper() != "DISCOUNTER"].copy()

            if COL_PLAN_ROL in df.columns:
                mask = df[COL_PLAN_ROL].astype(str).str.upper().str.strip() == 'GESTOR'
                df = df[mask].copy()
                
            dfs.append(df)
            print(f"    ✓ [SharePoint] {os.path.basename(str(path))} — {len(df):,} gestores cargados")
        except Exception as e:
            print(f"    ⚠️ No se pudo cargar el archivo de plan de trabajo ({os.path.basename(str(path))}): {e}")
            
    if not dfs: raise ValueError("No se encontraron registros de planes de trabajo en SharePoint para este periodo.")
    df = pd.concat(dfs, ignore_index=True)

    mapa_renombres = {}
    for col in df.columns:
        c_up = col.strip().upper()
        if "ID" in c_up and ("PDV" in c_up or "INVOLVES" in c_up):
            mapa_renombres[col] = COL_PLAN_ID_PDV
        elif c_up in ["ROL", "PERFIL"]:
            mapa_renombres[col] = COL_PLAN_ROL
        elif "VISITA" in c_up or "FREC" in c_up:
            if "CANTIDAD_VISITAS" not in mapa_renombres.values():
                mapa_renombres[col] = "CANTIDAD_VISITAS"
        elif c_up in ["MES", "MONTH"]:
            mapa_renombres[col] = COL_PLAN_MES
        elif c_up in ["AÑO", "ANIO", "YEAR"]:
            mapa_renombres[col] = COL_PLAN_ANIO

    df = df.rename(columns=mapa_renombres)

    if COL_PLAN_MES not in df.columns or COL_PLAN_ANIO not in df.columns:
        if "FECHA" in df.columns:
            df["_FECHA_TMP"] = pd.to_datetime(df["FECHA"], errors="coerce")
            if COL_PLAN_MES not in df.columns:
                df[COL_PLAN_MES] = df["_FECHA_TMP"].dt.month
            if COL_PLAN_ANIO not in df.columns:
                df[COL_PLAN_ANIO] = df["_FECHA_TMP"].dt.year
            df.drop(columns=["_FECHA_TMP"], errors="ignore", inplace=True)

    # Si el plan de trabajo no trae mes/año, se usan los del periodo que se
    # está procesando. Antes había un 7/2026 fijo: cualquier periodo sin esas
    # columnas quedaba marcado como julio de 2026, sin importar qué se corriera.
    if COL_PLAN_MES not in df.columns:
        df[COL_PLAN_MES] = spec.mes
    if COL_PLAN_ANIO not in df.columns:
        df[COL_PLAN_ANIO] = spec.anio

    if "CANTIDAD_VISITAS" not in df.columns:
        v_cols = [c for c in df.columns if "VISITAS" in c]
        if v_cols:
            df["CANTIDAD_VISITAS"] = df[v_cols[0]]

    cols_requeridas = [COL_PLAN_ID_PDV, COL_PLAN_ROL, 'CANTIDAD_VISITAS', COL_PLAN_MES, COL_PLAN_ANIO]
    faltantes = [c for c in cols_requeridas if c not in df.columns]
    if faltantes:
        raise KeyError(f"❌ Faltan columnas requeridas en los planes de trabajo: {faltantes}. Columnas disponibles: {list(df.columns)}")

    df = df[cols_requeridas].copy()
    df[COL_PLAN_ID_PDV] = pd.to_numeric(df[COL_PLAN_ID_PDV], errors="coerce")
    df[COL_PLAN_MES] = pd.to_numeric(df[COL_PLAN_MES], errors="coerce")
    df[COL_PLAN_ANIO] = pd.to_numeric(df[COL_PLAN_ANIO], errors="coerce")
    df['CANTIDAD_VISITAS'] = pd.to_numeric(df['CANTIDAD_VISITAS'], errors="coerce").fillna(0)
    df.dropna(subset=[COL_PLAN_ID_PDV, COL_PLAN_MES], inplace=True)
    df[COL_PLAN_ID_PDV] = df[COL_PLAN_ID_PDV].astype(int)
    df[COL_PLAN_MES] = df[COL_PLAN_MES].astype("Int64")
    df[COL_PLAN_ANIO] = df[COL_PLAN_ANIO].astype("Int64")
    return df

def load_visitas(files: dict) -> pd.DataFrame:
    dfs = []
    for path in files.get("visitas", []):
        try:
            stream_data = _get_sharepoint_file_stream(path)
            df = pd.read_excel(stream_data)
            dfs.append(df)
            print(f"    ✓ [SharePoint] {os.path.basename(str(path))} — {len(df):,} filas")
        except Exception as e:
            print(f"    ⚠️ No se pudo cargar el archivo de visitas {os.path.basename(str(path))}: {e}")
            
    if not dfs:
        print("    ⚠️ No se encontraron archivos de Visitas válidos en SharePoint. Se continúa sin visitas.")
        return pd.DataFrame(columns=[COL_VIS_ID_PDV, COL_VIS_EMPLEADO, COL_VIS_FECHA, "_semana_mes"])
        
    df = pd.concat(dfs, ignore_index=True)

    if "Tipo de check-in" in df.columns:
        n_antes = len(df)
        df = df[df["Tipo de check-in"].astype(str).str.strip() != "Sin check-in"].copy()
        n_dropped = n_antes - len(df)
        print(f"    🔎 Filtro Tipo de check-in: descartadas {n_dropped:,} filas 'Sin check-in' "
              f"({n_dropped/n_antes*100:.1f}%); quedan {len(df):,} visitas efectivas.")
    else:
        print("    ⚠ Columna 'Tipo de check-in' no encontrada.")

    df[COL_VIS_FECHA] = pd.to_datetime(df[COL_VIS_FECHA], dayfirst=True, errors="coerce")
    df["_semana_mes"] = df[COL_VIS_FECHA].apply(_semana_del_mes)
    df[COL_VIS_ID_PDV] = pd.to_numeric(df[COL_VIS_ID_PDV], errors="coerce")
    df.dropna(subset=[COL_VIS_ID_PDV], inplace=True)
    df[COL_VIS_ID_PDV] = df[COL_VIS_ID_PDV].astype(int)
    return df[[COL_VIS_ID_PDV, COL_VIS_EMPLEADO, COL_VIS_FECHA, "_semana_mes"]]

def load_nivel_impacto() -> pd.DataFrame:
    file_nivel = str(paths.EXHIB_NIVEL_IMPACTO).replace("\\", "/")
    try:
        stream_data = _get_sharepoint_file_stream(file_nivel)
        df = pd.read_excel(stream_data)
        df.columns = df.columns.str.strip()
        return df[["Tipo Exhibición", "Nivel Impacto"]]
    except Exception as e:
        print(f"    ⚠️ No se pudo cargar el archivo de nivel de impacto: {e}. Se retorna estructura vacía.")
        return pd.DataFrame(columns=["Tipo Exhibición", "Nivel Impacto"])

# ==============================================================================
# 3. MÓDULOS DE CÁLCULO
# ==============================================================================

def calcular_gratis_concurso(df_enc, df_plan, df_visitas) -> pd.DataFrame:
    mask = df_enc[COL_EXHIBICIONES].isin([EXHIB_GRATIS, EXHIB_CONCURSO])
    df = df_enc[mask].copy()
    if df.empty: return pd.DataFrame(columns=OUTPUT_COLS)
    
    df["Cantidad"] = pd.to_numeric(df[COL_CANTIDAD_GC], errors="coerce").fillna(0)
    df["_rol"] = df[COL_PERFIL_EMP].map(PERFIL_A_ROL)
    
    plan_lookup = df_plan.rename(columns={
        COL_PLAN_ID_PDV: COL_ID_PDV, COL_PLAN_ROL: "_rol",
        'CANTIDAD_VISITAS': "_frec_plan", COL_PLAN_MES: COL_MES, COL_PLAN_ANIO: COL_ANIO
    })[[COL_ID_PDV, "_rol", "_frec_plan", COL_MES, COL_ANIO]]

    # FIX: el Plan de Trabajo trae varias filas para el mismo (PDV, rol, mes,
    # año) — probablemente una por SKU/categoría del plan, sin deduplicar por
    # PDV. El merge de abajo es por esa llave, SIN Empleado, así que sin este
    # drop_duplicates cada captura real de la encuesta se multiplicaba por la
    # cantidad de filas repetidas: un PDV con 8 filas duplicadas en el plan
    # convertía 2 capturas de 12 unidades en 16 filas de 12, sumando 192 en
    # vez de 24. Se verificó que _frec_plan es idéntica entre los duplicados
    # (incluidos los 3 casos con gestores distintos que comparten PDV+rol),
    # así que quedarse con la primera fila no pierde ningún dato real.
    n_antes = len(plan_lookup)
    plan_lookup = plan_lookup.drop_duplicates(subset=[COL_ID_PDV, "_rol", COL_MES, COL_ANIO])
    if len(plan_lookup) != n_antes:
        print(f"  ℹ️  Plan de trabajo: {n_antes:,} filas → {len(plan_lookup):,} tras "
              f"deduplicar por (PDV, rol, mes, año) — evita inflar Cantidad en el merge.")

    df = df.merge(plan_lookup, on=[COL_ID_PDV, "_rol", COL_MES, COL_ANIO], how="left")
    
    if not df_visitas.empty:
        df_visitas["_mes_vis"] = df_visitas[COL_VIS_FECHA].dt.month
        df_visitas["_anio_vis"] = df_visitas[COL_VIS_FECHA].dt.year
        frec_real = df_visitas.groupby([COL_VIS_ID_PDV, COL_VIS_EMPLEADO, "_mes_vis", "_anio_vis"]).size().reset_index(name="_frec_real")
        frec_real.columns = [COL_ID_PDV, COL_EMPLEADO, COL_MES, COL_ANIO, "_frec_real"]
        df = df.merge(frec_real, on=[COL_ID_PDV, COL_EMPLEADO, COL_MES, COL_ANIO], how="left")
    else:
        df["_frec_real"] = 0

    def frec_referencia(row):
        perfil = str(row[COL_PERFIL_EMP]).upper()
        frec_real_val = row.get("_frec_real", 0)
        if perfil == PERFIL_PLAN:
            frec_plan_val = row.get("_frec_plan", 0)
            return frec_plan_val if pd.notna(frec_plan_val) and frec_plan_val > 0 else frec_real_val
        return frec_real_val

    df["_frec_ref"] = df.apply(frec_referencia, axis=1)
    KEY = [COL_MES, COL_ANIO, COL_ID_PDV, COL_EXHIBICIONES, COL_MARCA_GC, COL_TIPO_EXHIB_GC, COL_PERFIL_EMP, COL_EMPLEADO]

    def evaluar_cumplimiento(grupo):
        semanas_distintas = grupo["_semana_mes"].nunique()
        frec_ref = grupo["_frec_ref"].iloc[0] if not grupo.empty else 0
        cumple = semanas_distintas >= 2 if frec_ref > 1 else semanas_distintas >= 1
        if cumple:
            cant_sem = grupo.groupby("_semana_mes")["Cantidad"].sum()
            return pd.Series({"Cantidad": cant_sem.mean(),
                            "_cumple": True,
                            "_semanas_distintas": int(semanas_distintas),
                            "_frec_ref": float(frec_ref) if pd.notna(frec_ref) else 0.0})
        return pd.Series({"Cantidad": float(grupo["Cantidad"].sum()),
                        "_cumple": False,
                        "_semanas_distintas": int(semanas_distintas),
                        "_frec_ref": float(frec_ref) if pd.notna(frec_ref) else 0.0})

    res = df.groupby(KEY, dropna=False).apply(evaluar_cumplimiento).reset_index()
    fuera = res[res["_cumple"] == False].copy()
    if not fuera.empty:
        fuera = fuera.rename(columns={
            COL_EXHIBICIONES: "Categoría",
            COL_MARCA_GC: "Marca",
            COL_TIPO_EXHIB_GC: "Tipo Exhibición",
        })
        fuera["Categoría"] = fuera["Categoría"].str.replace("Exhibiciones ", "").str.strip()
        ruta_fuera = f"{RUTA_CARPETA_SALIDAS_EXHIB}/Exh_Gratis_Fuera_de_Regla.xlsx".replace("\\", "/")
        try:
            try:
                stream_prev_fuera = _get_sharepoint_file_stream(ruta_fuera)
                prev = pd.read_excel(stream_prev_fuera)
            except Exception:
                prev = pd.DataFrame()

            if not prev.empty and COL_MES in prev.columns and COL_ANIO in prev.columns:
                periodos_nuevos = set(map(tuple, fuera[[COL_MES, COL_ANIO]].drop_duplicates().values.tolist()))
                prev = prev[~prev.apply(lambda r: (int(r[COL_MES]), int(r[COL_ANIO])) in periodos_nuevos, axis=1)]
                fuera_out = pd.concat([prev, fuera], ignore_index=True)
            else:
                fuera_out = fuera
            
            output_buffer = BytesIO()
            fuera_out.to_excel(output_buffer, index=False, engine="openpyxl")
            _upload_sharepoint_file(ruta_fuera, output_buffer)
        except Exception as ee:
            print(f"  ⚠️ No pude persistir Exh Gratis fuera de regla en SharePoint: {ee}")

    res = res[res["_cumple"] == True].drop(
        columns=["_cumple", "_semanas_distintas", "_frec_ref"]
    )
    res = res.rename(columns={COL_EXHIBICIONES: "Categoría", COL_MARCA_GC: "Marca", COL_TIPO_EXHIB_GC: "Tipo Exhibición"})
    res["Categoría"] = res["Categoría"].str.replace("Exhibiciones ", "").str.strip()
    return res

def calcular_pagadas(df_enc) -> pd.DataFrame:
    df = df_enc[df_enc[COL_EXHIBICIONES] == EXHIB_PAGADA].copy()
    if df.empty:
        return pd.DataFrame(columns=[COL_MES, COL_ANIO, COL_ID_PDV, "Tipo Exhibición", "Marca", "Cantidad", COL_EMPLEADO, COL_PERFIL_EMP, "Categoría"])
    
    col_impl = None
    for c in df.columns:
        c_low = c.lower()
        if "implementada" in c_low and "planning" in c_low:
            col_impl = c
            break
            
    if not col_impl:
        print("    ⚠️ No se encontró columna de implementación planning para pagadas. Se omite.")
        return pd.DataFrame(columns=[COL_MES, COL_ANIO, COL_ID_PDV, "Tipo Exhibición", "Marca", "Cantidad", COL_EMPLEADO, COL_PERFIL_EMP, "Categoría"])

    df["_impl_norm"] = df[col_impl].astype(str).str.strip().str.lower()
    df["_causal_norm"] = df[COL_CAUSAL].astype(str).str.strip().str.lower() if COL_CAUSAL in df.columns else ""
    
    m_si = df["_impl_norm"] == "si"
    m_contra = (df["_impl_norm"] == "no") & (df["_causal_norm"] == CAUSAL_CONTRA.lower())

    def extraer(sub, t, m, c):
        if sub.empty:
            return pd.DataFrame(columns=[COL_MES, COL_ANIO, COL_ID_PDV, "Tipo Exhibición", "Marca", "Cantidad", COL_EMPLEADO, COL_PERFIL_EMP])
        return pd.DataFrame({
            COL_MES: sub[COL_MES], COL_ANIO: sub[COL_ANIO], COL_ID_PDV: sub[COL_ID_PDV],
            "Tipo Exhibición": sub[t], "Marca": sub[m],
            "Cantidad": pd.to_numeric(sub[c], errors="coerce").fillna(0),
            COL_EMPLEADO: sub[COL_EMPLEADO], COL_PERFIL_EMP: sub[COL_PERFIL_EMP]
        })

    base = pd.concat([extraer(df[m_si], COL_TIPO_EXHIB_PAG, COL_MARCA_PAG, COL_CANTIDAD_PAG),
                      extraer(df[m_contra], COL_TIPO_CONTRA, COL_MARCA_CONTRA, COL_CANTIDAD_CONTRA)], ignore_index=True)
    if base.empty: return base
    agg = base.groupby([COL_MES, COL_ANIO, COL_ID_PDV, "Tipo Exhibición", "Marca", COL_EMPLEADO, COL_PERFIL_EMP], dropna=False)["Cantidad"].sum().reset_index()
    agg["Categoría"] = "Pagada"
    return agg

# ==============================================================================
# 4. ORQUESTADOR
# ==============================================================================

def run(spec: pr.PeriodoSpec):
    _apuntar_rutas_al_periodo(spec)
    print(f"── Resolviendo fuentes en SharePoint ({spec.etiqueta}) ─")
    files = resolve_files(spec)
    print("\n── Cargando fuentes desde SharePoint ────────────")
    df_enc = load_encuestas(files)
    df_plan = load_plan(files, spec)
    df_vis = load_visitas(files)
    df_nivel = load_nivel_impacto()

    print(f"\n   Encuestas consolidadas: {len(df_enc):,} filas")
    print(f"   Plan de trabajo:        {len(df_plan):,} filas")
    print(f"   Visitas realizadas:     {len(df_vis):,} filas")

    print("\n── Módulo 1: Exhibiciones Pagadas ───────────────")
    df_pag = calcular_pagadas(df_enc)
    print(f"   Implementadas pagadas:          {len(df_pag):,} filas")

    print("\n── Módulo 2: Exhibiciones Gratis y Concurso ─────")
    df_gc = calcular_gratis_concurso(df_enc, df_plan, df_vis)
    print(f"   Implementadas gratis/concurso: {len(df_gc):,} filas")

    print("\n── Consolidando y escribiendo en SharePoint ─────")
    df_out = pd.concat([df_pag, df_gc], ignore_index=True)
    if not df_nivel.empty:
        df_out = df_out.merge(df_nivel, on="Tipo Exhibición", how="left")
    else:
        df_out["Nivel Impacto"] = "SIN NIVEL"
        
    df_out[COL_MES] = df_out[COL_MES].astype("Int64")
    df_out[COL_ANIO] = df_out[COL_ANIO].astype("Int64")
    df_out["Mes-Año"] = df_out[COL_MES].astype(str).str.zfill(2) + "-" + df_out[COL_ANIO].astype(str)
    
    df_out = df_out.rename(columns={COL_MES: "Mes", COL_ANIO: "Año", COL_ID_PDV: "ID PDV", COL_EMPLEADO: "Empleado", COL_PERFIL_EMP: "Rol Empleado"})
    df_out = df_out[OUTPUT_COLS].sort_values(["Año", "Mes", "ID PDV"])
    df_out["Cantidad"] = pd.to_numeric(df_out["Cantidad"], errors="coerce").round(2)

    try:
        stream_prev = _get_sharepoint_file_stream(OUTPUT_PATH)
        df_prev = pd.read_excel(stream_prev, sheet_name="Exhibiciones_implementadas")
        if "Mes" in df_prev.columns and "Año" in df_prev.columns:
            periodos_actuales = set(
                (int(m), int(a)) for m, a in zip(df_out["Mes"], df_out["Año"])
                if pd.notna(m) and pd.notna(a)
            )
            df_prev["_mes_int"]  = pd.to_numeric(df_prev["Mes"],  errors="coerce")
            df_prev["_anio_int"] = pd.to_numeric(df_prev["Año"], errors="coerce")
            mask_keep = ~df_prev.apply(
                lambda r: (r["_mes_int"], r["_anio_int"]) in periodos_actuales,
                axis=1,
            )
            df_prev = df_prev[mask_keep].drop(columns=["_mes_int", "_anio_int"])
            df_out = pd.concat([df_prev, df_out], ignore_index=True, sort=False)
            df_out = df_out[OUTPUT_COLS].sort_values(["Año", "Mes", "ID PDV"])
        else:
            print("  ⚠️ Resultado previo sin Mes/Año — descartado en upsert.")
    except Exception as e:
        print(f"  ⚠️ No se pudo leer resultado previo de SharePoint ({e}); se sobrescribe.")

    output_buffer = BytesIO()
    df_out.to_excel(output_buffer, index=False, sheet_name="Exhibiciones_implementadas", engine="openpyxl")
    _upload_sharepoint_file(OUTPUT_PATH, output_buffer)
    print(f"\n✓ Output escrito con éxito en SharePoint: {OUTPUT_PATH} ({len(df_out):,} filas)")

    # El archivo acumula todos los meses; solo_periodo sube solo el procesado.
    supabase_io.cargar_detalle_seguro(
        "exhibiciones_gratis_detalle", df_out, spec.mes, spec.anio)
    
    m_gratis, m_pag = df_out["Categoría"].isin(["Gratis", "Concurso"]), df_out["Categoría"] == "Pagada"
    m_alto, m_medio = df_out["Nivel Impacto"] == "ALTO IMPACTO", df_out["Nivel Impacto"] == "MEDIO IMPACTO"

    print("\n── Resumen de Exhibiciones ──────────────────────")
    print(f"   Total Exhibiciones:                                     {df_out['Cantidad'].sum():>8,.0f}")
    print(f"   Total Exhibiciones Gratis:                                {df_out.loc[m_gratis, 'Cantidad'].sum():>8,.0f}")
    print(f"   Total Exhibiciones Pagadas:                              {df_out.loc[m_pag, 'Cantidad'].sum():>8,.0f}")
    print(f"\n   Total Exhibiciones Alto Impacto:                         {df_out.loc[m_alto, 'Cantidad'].sum():>8,.0f}")
    print(f"   Total Exhibiciones Medio Impacto:                        {df_out.loc[m_medio, 'Cantidad'].sum():>8,.0f}")
    print("─────────────────────────────────────────────────")

def generar_resumen_kpi_exhibiciones_gratis(spec: pr.PeriodoSpec):
    print(f"\n--- KPI (V3): RESUMEN EXHIBICIONES GRATIS por empleado ({spec.etiqueta}) ---")
    try:
        stream_data = _get_sharepoint_file_stream(OUTPUT_PATH)
        df = pd.read_excel(stream_data, engine='openpyxl')
    except Exception as e:
        print(f"❌ No se pudo leer el archivo de salida desde SharePoint: {e}")
        return
        
    df.columns = [str(c).strip() for c in df.columns]
    print(f"  Leyendo archivo desde SharePoint ({len(df)} filas)")

    df = df[df['Categoría'].astype(str).str.strip().str.lower() == 'gratis'].copy()

    df['Mes'] = pd.to_numeric(df.get('Mes'), errors='coerce').fillna(0).astype(int)
    df['Año'] = pd.to_numeric(df.get('Año'), errors='coerce').fillna(0).astype(int)
    df['Cantidad'] = pd.to_numeric(df['Cantidad'], errors='coerce').fillna(0)
    df['Nivel Impacto'] = df['Nivel Impacto'].astype(str).str.strip().str.upper()

    df_periodo = df[(df['Mes'] == spec.mes) & (df['Año'] == spec.anio)].copy()
    if df_periodo.empty:
        raise ValueError(
            f"Exh Gratis KPI: no hay filas del periodo {spec.etiqueta}."
        )
    df = df_periodo

    resumen = (
        df.pivot_table(
            index=['Mes', 'Año', 'Empleado'],
            columns='Nivel Impacto',
            values='Cantidad',
            aggfunc='sum',
            fill_value=0,
        )
        .reset_index()
    )
    for col in ['ALTO IMPACTO', 'MEDIO IMPACTO']:
        if col not in resumen.columns:
            resumen[col] = 0

    try:
        path_dy_cupos = str(paths.DYP_BASE_CUPOS).replace("\\", "/")
        stream_cupos = _get_sharepoint_file_stream(path_dy_cupos)
        bc = pd.read_excel(stream_cupos, sheet_name="Tabla total roles")
        bc["NOMBRE_N"] = (
            bc["NOMBRE"].astype(str).str.strip().str.upper()
            .str.replace(r"\s+", " ", regex=True)
        )
        canal_por_nombre = dict(zip(
            bc["NOMBRE_N"],
            bc["CANAL"].astype(str).str.strip().str.upper(),
        ))
    except Exception as e:
        print(f"  ⚠️ No pude leer Base cupos para targets por canal desde SharePoint ({e}); uso 3/5 por defecto.")
        canal_por_nombre = {}

    def _targets_por_canal(empleado: str) -> tuple[int, int]:
        canal = canal_por_nombre.get(
            str(empleado).strip().upper().replace("  ", " "),
            "",
        )
        if canal in ("PROXIMITY", "PROXIMITY TAT"):
            return 3, 5
        return 2, 8

    targets = resumen["Empleado"].apply(_targets_por_canal)
    resumen["TARGET_ALTO"]  = [t[0] for t in targets]
    resumen["TARGET_MEDIO"] = [t[1] for t in targets]
    resumen["CUMP_ALTO"]    = resumen["ALTO IMPACTO"]  / resumen["TARGET_ALTO"]
    resumen["CUMP_MEDIO"]   = resumen["MEDIO IMPACTO"] / resumen["TARGET_MEDIO"]

    cols_final = ['Mes', 'Año', 'Empleado',
                  'ALTO IMPACTO', 'MEDIO IMPACTO',
                  'TARGET_ALTO', 'TARGET_MEDIO',
                  'CUMP_ALTO', 'CUMP_MEDIO']
    resumen = resumen[cols_final].copy()
    resumen['TOTAL'] = resumen['ALTO IMPACTO'] + resumen['MEDIO IMPACTO']
    resumen = resumen.sort_values(['Año', 'Mes', 'Empleado']).reset_index(drop=True)
    cols_final.append('TOTAL')

    if not resumen.empty:
        anio_max = int(resumen['Año'].max())
        mes_max  = int(resumen[resumen['Año'] == anio_max]['Mes'].max())
        mask = (resumen['Año'] == anio_max) & (resumen['Mes'] == mes_max)
        resumen_activo = resumen[mask].copy()
    else:
        resumen_activo = resumen.copy()

    path_kpis = f"{RUTA_CARPETA_SALIDAS_EXHIB}/KPI_Exhibiciones_Gratis.xlsx".replace("\\", "/")
    buf_kpi = BytesIO()
    resumen_activo.to_excel(buf_kpi, index=False, engine='openpyxl')
    _upload_sharepoint_file(path_kpis, buf_kpi)
    print(f"  ✅ Mes activo guardado en SharePoint: {len(resumen_activo)} empleados → {os.path.basename(str(path_kpis))}")

    path_hist = f"{RUTA_CARPETA_SALIDAS_EXHIB}/KPI_Exhibiciones_Gratis_Historico.xlsx".replace("\\", "/")
    try:
        stream_hist = _get_sharepoint_file_stream(path_hist)
        hist_prev = pd.read_excel(stream_hist, engine='openpyxl')
        hist_prev['Mes'] = pd.to_numeric(hist_prev.get('Mes'), errors='coerce').fillna(0).astype(int)
        hist_prev['Año'] = pd.to_numeric(hist_prev.get('Año'), errors='coerce').fillna(0).astype(int)
    except Exception:
        hist_prev = pd.DataFrame(columns=cols_final)

    if not hist_prev.empty and not resumen.empty:
        periodos_nuevos = set(map(tuple,
            resumen[['Mes', 'Año']].drop_duplicates().values.tolist()))
        mask_keep = ~hist_prev.apply(
            lambda r: (int(r['Mes']), int(r['Año'])) in periodos_nuevos,
            axis=1,
        )
        hist_prev = hist_prev[mask_keep]

    hist_final = pd.concat([hist_prev, resumen], ignore_index=True)
    hist_final = hist_final.sort_values(['Año', 'Mes', 'Empleado'])
    
    buf_hist = BytesIO()
    hist_final.to_excel(buf_hist, index=False, engine='openpyxl')
    _upload_sharepoint_file(path_hist, buf_hist)
    print(f"  ✅ Histórico actualizado en SharePoint → {os.path.basename(str(path_hist))}")

    # El KPI agregado ya no va a Supabase: lo que se carga es el DETALLE
    # (tabla exhibiciones_gratis_detalle). El KPI sigue publicándose como
    # Excel en SharePoint.


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="ETL Exhibiciones Gratis — SharePoint Cloud")
    parser.add_argument("--solo", nargs="+", choices=["full", "kpi"],
                        help="Default: corre todo + kpi.")
    pr.cli_add_periodos_arg(parser)
    args = parser.parse_args()
    pasos = args.solo or ["full", "kpi"]
    specs = pr.periodos_de_args(args)

    if len(specs) > 1:
        print(f"🎯 ETL Exh Gratis — multi-periodo: {len(specs)} meses → "
              f"{', '.join(s.etiqueta for s in specs)}")

    for i, spec in enumerate(specs, 1):
        if len(specs) > 1:
            print(f"\n▶ Periodo {i}/{len(specs)}: {spec.etiqueta}")
        print(f"\n🎯 ETL Exh Gratis — procesandoperiodo {spec.etiqueta} ({spec})")

        try:
            if "full" in pasos: run(spec)
            if "kpi"  in pasos: generar_resumen_kpi_exhibiciones_gratis(spec)
        except Exception as e:
            print(f"\n❌ ERROR para {spec.etiqueta}: {e}")
            if len(specs) > 1:
                print(f"   Continuando con los siguientes periodos...")
            else:
                raise