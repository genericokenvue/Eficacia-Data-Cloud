import os
import glob
from pathlib import Path
import pandas as pd
import numpy as np
from dotenv import load_dotenv

import paths
import periodo_resolver as pr

# ==============================================================================
# 0. CONFIGURACIÓN ENTORNO NUBE — SHAREPOINT / AZURE (100% NUBE)
# ==============================================================================
load_dotenv()

AZURE_STORAGE_CONNECTION_STRING = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

SHAREPOINT_BASE_DIR = os.getenv("SHAREPOINT_BASE_DIR", "Equipo Información/BI/INVOLVES")
_BASES_ROOT = os.getenv("BASES_ROOT", f"{SHAREPOINT_BASE_DIR}/BASES DE RESPUESTAS")
_SALIDAS_ROOT = os.getenv("SALIDAS_ROOT", f"{SHAREPOINT_BASE_DIR}/SALIDAS")
AÑO_ACTUAL = pd.Timestamp.now().year

RUTA_CARPETA_PT = f"{SHAREPOINT_BASE_DIR}/PLAN DE TRABAJO"
RUTA_CARPETA_BASES_EXHIB = f"{_BASES_ROOT}/EXHIBICIONES/{AÑO_ACTUAL}"
RUTA_CARPETA_SALIDAS_EXHIB = f"{_SALIDAS_ROOT}/EXHIBICIONES"

# Creamos la ruta física localmente de forma temporal para satisfacer el .is_dir() de periodo_resolver.py
dir_exhib_temporal = Path(f"{_BASES_ROOT}/EXHIBICIONES/{AÑO_ACTUAL}")
dir_exhib_temporal.mkdir(parents=True, exist_ok=True)

# Forzamos rutas exclusivamente a la nube usando objetos Path
paths.EXHIB_DATA_DIR = dir_exhib_temporal
paths.CIF_OUT_FINAL = Path(f"{RUTA_CARPETA_PT}/Plan de trabajo.xlsx")
paths.EXHIB_SALIDA = Path(f"{RUTA_CARPETA_SALIDAS_EXHIB}")
paths.EXHIB_NIVEL_IMPACTO = Path(f"{_BASES_ROOT}/Nivel Impacto.xlsx")
paths.DYP_BASE_CUPOS = Path(f"{_BASES_ROOT}/Base cupos.xlsx")

# ==============================================================================
# 1. CONFIGURACIÓN
# ==============================================================================
DATA_DIR = str(paths.EXHIB_DATA_DIR)

FILE_NIVEL_IMPACTO = str(paths.EXHIB_NIVEL_IMPACTO)
OUTPUT_PATH = str(paths.EXHIB_SALIDA / "Resultado exhibiciones gratis.xlsx")
PLAN_SHEET = "Plan de trabajo"

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
COL_IMPLEMENTADA    = "La Exhibicion esta implementada de acuerdo con le planning?"
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

PERFILES_VISITAS_REALES = {"SUPERVISOR", "GENERADOR DE DEMANDA"}
PERFIL_PLAN         = "GESTOR"

PERFIL_A_ROL = {
    "GESTOR": "GESTOR",
    "SUPERVISOR": "SUPERVISOR",
    "GENERADOR DE DEMANDA": "GENERADOR DE DEMANDA",
}

OUTPUT_COLS = ["Mes", "Año", "Mes-Año", "ID PDV", "Categoría", "Tipo Exhibición", "Nivel Impacto", "Marca", "Cantidad", "Empleado", "Rol Empleado"]

# ==============================================================================
# 2. CARGADORES (loaders.py)
# ==============================================================================

import sys
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass


def _semana_del_mes(fecha) -> int:
    if pd.isnull(fecha): return -1
    primer_dia = fecha.replace(day=1)
    offset = primer_dia.weekday()
    return (fecha.day + offset - 1) // 7 + 1

def resolve_files(spec: pr.PeriodoSpec) -> dict:
    groups: dict[str, list[str]] = {}

    # Obtenemos la ruta base dinámicamente desde el módulo paths (Cloud / Local agnostic)
    base_dir = Path(paths.EXHIB_DATA_DIR)

    p_informar = str(base_dir / f"Base Exhibiciones Informar {spec.mes_str} {spec.anio}.xlsx")
    p_planning = str(base_dir / f"Base Exhibiciones Planning {spec.mes_str} {spec.anio}.xlsx")
    
    try:
        p_visitas = str(pr.cif_involves(spec))
    except Exception:
        p_visitas = ""

    # Plan de trabajo dinámico según paths
    try:
        if hasattr(paths, "CIF_OUT_FINAL") and paths.CIF_OUT_FINAL:
            p_plan_trabajo = str(paths.CIF_OUT_FINAL)
        else:
            p_plan_trabajo = str(base_dir.parent.parent / "PLAN DE TRABAJO" / "Plan de trabajo.xlsx")
    except Exception:
        p_plan_trabajo = ""

    groups["informar"]     = [p_informar] if p_informar else []
    groups["planning"]     = [p_planning] if p_planning else []
    groups["visitas"]      = [p_visitas] if p_visitas else []
    groups["plan_trabajo"] = [p_plan_trabajo] if p_plan_trabajo else []

    print(f"\n☁️ [NUBE / GITHUB ACTIONS] Rutas y archivos aplicados para el periodo {spec.etiqueta}:")
    for tipo, archivos in groups.items():
        print(f"  [{tipo}] {len(archivos)} archivo(s):")
        for a in archivos:
            print(f"    - Ruta: {a}")
            print(f"    - Archivo: {os.path.basename(a) if a else 'NINGUNO'}")
            
    return groups
            
    return groups

    def limpiar_path(p):
        nombre = os.path.basename(p)
        if "informe-gerencial-visitas" in nombre.lower():
            return f"{_BASES_ROOT}/CIF/INVOLVES/{nombre}"
        return f"{_BASES_ROOT}/EXHIBICIONES/{spec.anio}/{nombre}"

    groups["informar"]   = [limpiar_path(p_informar)]
    groups["planning"]   = [limpiar_path(p_planning)]
    groups["visitas"]    = [limpiar_path(p_visitas)]
    groups["plan_trabajo"] = [str(paths.CIF_OUT_FINAL)]

    print(f"\n☁️ [NUBE / SHAREPOINT] Rutas y archivos aplicados para el periodo {spec.etiqueta}:")
    print(f"  📂 Carpeta Base Exhibiciones: {RUTA_CARPETA_BASES_EXHIB}")
    print(f"  📂 Carpeta Plan de Trabajo:  {RUTA_CARPETA_PT}")
    print(f"  📂 Carpeta Salidas Cloud:    {RUTA_CARPETA_SALIDAS_EXHIB}")
    
    for tipo, archivos in groups.items():
        print(f"  [{tipo}] {len(archivos)} archivo(s) en SharePoint:")
        for a in archivos:
            print(f"    - Ruta cloud: {a}")
            print(f"    - Archivo: {os.path.basename(a)}")
            
    return groups

def load_encuestas(files: dict) -> pd.DataFrame:
    dfs = []
    for path in files.get("informar", []) + files.get("planning", []):
        try:
            df = pd.read_excel(path, sheet_name="report")
            dfs.append(df)
            print(f"    ✓ [Cloud] {os.path.basename(path)} — {len(df):,} filas")
        except Exception as e: print(f"    ⚠ Error leyendo desde SharePoint {os.path.basename(path)}: {e}")
    if not dfs: raise ValueError("No se encontraron archivos de encuesta en la nube.")
    cols_comunes = set(dfs[0].columns)
    for df in dfs[1:]: cols_comunes &= set(df.columns)
    cols_comunes = [c for c in dfs[0].columns if c in cols_comunes]
    df = pd.concat([d[cols_comunes] for d in dfs], ignore_index=True)
    df[COL_FECHA] = pd.to_datetime(df[COL_FECHA], dayfirst=True, errors="coerce")
    df[COL_MES] = df[COL_FECHA].dt.month
    df[COL_ANIO] = df[COL_FECHA].dt.year
    df["_semana_mes"] = df[COL_FECHA].apply(_semana_del_mes)
    return df

def load_plan(files: dict) -> pd.DataFrame:
    dfs = []
    for path in files.get("plan_trabajo", []):
        try:
            df = pd.read_excel(path)   
            df.columns = [str(c).strip().upper() for c in df.columns]
            if COL_PLAN_ROL in df.columns:
                mask = df[COL_PLAN_ROL].astype(str).str.upper().str.strip() == 'GESTOR'
                df = df[mask].copy()
            dfs.append(df)
            print(f"    ✓ [Cloud] {os.path.basename(path)} — {len(df):,} gestores cargados")
        except Exception as e:
            print(f"    ⚠ Error leyendo plan en la nube {os.path.basename(path)}: {e}")
    if not dfs: raise ValueError("No se encontraron registros de GESTOR en la nube.")
    df = pd.concat(dfs, ignore_index=True)
    df = df[[COL_PLAN_ID_PDV, COL_PLAN_ROL, 'CANTIDAD_VISITAS', COL_PLAN_MES, COL_PLAN_ANIO]].copy()
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
            df = pd.read_excel(path)
            dfs.append(df)
            print(f"    ✓ [Cloud] {os.path.basename(path)} — {len(df):,} filas")
        except Exception as e:
            print(f"    ⚠ Error leyendo visitas en la nube {os.path.basename(path)}: {e}")
    if not dfs:
        raise ValueError("No se encontraron archivos de Visitas en la nube.")
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
    df = pd.read_excel(FILE_NIVEL_IMPACTO)
    df.columns = df.columns.str.strip()
    return df[["Tipo Exhibición", "Nivel Impacto"]]

# ==============================================================================
# 3. MÓDULOS DE LÓGICA (Gratis y Pagadas)
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
    
    df = df.merge(plan_lookup, on=[COL_ID_PDV, "_rol", COL_MES, COL_ANIO], how="left")
    df_visitas["_mes_vis"] = df_visitas[COL_VIS_FECHA].dt.month
    df_visitas["_anio_vis"] = df_visitas[COL_VIS_FECHA].dt.year
    
    frec_real = df_visitas.groupby([COL_VIS_ID_PDV, COL_VIS_EMPLEADO, "_mes_vis", "_anio_vis"]).size().reset_index(name="_frec_real")
    frec_real.columns = [COL_ID_PDV, COL_EMPLEADO, COL_MES, COL_ANIO, "_frec_real"]
    df = df.merge(frec_real, on=[COL_ID_PDV, COL_EMPLEADO, COL_MES, COL_ANIO], how="left")

    def frec_referencia(row):
        perfil = str(row[COL_PERFIL_EMP]).upper()
        if perfil == PERFIL_PLAN:
            return row["_frec_plan"] if pd.notna(row["_frec_plan"]) and row["_frec_plan"] > 0 else row["_frec_real"]
        return row["_frec_real"]

    df["_frec_ref"] = df.apply(frec_referencia, axis=1)
    KEY = [COL_MES, COL_ANIO, COL_ID_PDV, COL_EXHIBICIONES, COL_MARCA_GC, COL_TIPO_EXHIB_GC, COL_PERFIL_EMP, COL_EMPLEADO]

    def evaluar_cumplimiento(grupo):
        semanas_distintas = grupo["_semana_mes"].nunique()
        frec_ref = grupo["_frec_ref"].iloc[0]
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
        ruta_fuera = f"{RUTA_CARPETA_SALIDAS_EXHIB}/Exh_Gratis_Fuera_de_Regla.xlsx"
        try:
            if os.path.exists(ruta_fuera) or str(ruta_fuera).startswith("http") or "sharepoint" in str(ruta_fuera).lower():
                prev = pd.read_excel(ruta_fuera)
                if COL_MES in prev.columns and COL_ANIO in prev.columns:
                    periodos_nuevos = set(map(tuple,
                        fuera[[COL_MES, COL_ANIO]].drop_duplicates().values.tolist()))
                    prev = prev[~prev.apply(
                        lambda r: (int(r[COL_MES]), int(r[COL_ANIO])) in periodos_nuevos,
                        axis=1,
                    )]
                    fuera_out = pd.concat([prev, fuera], ignore_index=True)
                else:
                    fuera_out = fuera
            else:
                fuera_out = fuera
            os.makedirs(os.path.dirname(str(ruta_fuera)), exist_ok=True)
            fuera_out.to_excel(ruta_fuera, index=False, engine="openpyxl")
        except Exception as ee:
            print(f"  ⚠️ No pude persistir Exh Gratis fuera de regla en nube: {ee}")

    res = res[res["_cumple"] == True].drop(
        columns=["_cumple", "_semanas_distintas", "_frec_ref"]
    )
    res = res.rename(columns={COL_EXHIBICIONES: "Categoría", COL_MARCA_GC: "Marca", COL_TIPO_EXHIB_GC: "Tipo Exhibición"})
    res["Categoría"] = res["Categoría"].str.replace("Exhibiciones ", "").str.strip()
    return res

def calcular_pagadas(df_enc) -> pd.DataFrame:
    df = df_enc[df_enc[COL_EXHIBICIONES] == EXHIB_PAGADA].copy()
    df["_impl_norm"] = df[COL_IMPLEMENTADA].astype(str).str.strip().str.lower()
    df["_causal_norm"] = df[COL_CAUSAL].astype(str).str.strip().str.lower()
    
    m_si = df["_impl_norm"] == "si"
    m_contra = (df["_impl_norm"] == "no") & (df["_causal_norm"] == CAUSAL_CONTRA.lower())

    def extraer(sub, t, m, c):
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
# 4. ORQUESTADOR (pipeline.py)
# ==============================================================================

def run(spec: pr.PeriodoSpec):
    if not DATA_DIR:
        print(f"❌ No se encontró directorio cloud de exhibiciones: {DATA_DIR}")
        return
    print(f"── Resolviendo fuentes en SharePoint ({spec.etiqueta}) ─")
    files = resolve_files(spec)
    print("\n── Cargando fuentes desde SharePoint ────────────")
    df_enc = load_encuestas(files)
    df_plan = load_plan(files)
    df_vis = load_visitas(files)
    df_nivel = load_nivel_impacto()

    print(f"\n   Encuestas consolidadas: {len(df_enc):,} filas")
    print(f"   Plan de trabajo:        {len(df_plan):,} filas")
    print(f"   Visitas realizadas:     {len(df_vis):,} filas")

    print("\n── Módulo 1: Exhibiciones Pagadas ───────────────")
    df_pag = calcular_pagadas(df_enc)
    print(f"   Implementadas pagadas:           {len(df_pag):,} filas")

    print("\n── Módulo 2: Exhibiciones Gratis y Concurso ─────")
    df_gc = calcular_gratis_concurso(df_enc, df_plan, df_vis)
    print(f"   Implementadas gratis/concurso: {len(df_gc):,} filas")

    print("\n── Consolidando y escribiendo en SharePoint ─────")
    df_out = pd.concat([df_pag, df_gc], ignore_index=True)
    df_out = df_out.merge(df_nivel, on="Tipo Exhibición", how="left")
    df_out[COL_MES] = df_out[COL_MES].astype("Int64")
    df_out[COL_ANIO] = df_out[COL_ANIO].astype("Int64")
    df_out["Mes-Año"] = df_out[COL_MES].astype(str).str.zfill(2) + "-" + df_out[COL_ANIO].astype(str)
    
    df_out = df_out.rename(columns={COL_MES: "Mes", COL_ANIO: "Año", COL_ID_PDV: "ID PDV", COL_EMPLEADO: "Empleado", COL_PERFIL_EMP: "Rol Empleado"})
    df_out = df_out[OUTPUT_COLS].sort_values(["Año", "Mes", "ID PDV"])
    df_out["Cantidad"] = pd.to_numeric(df_out["Cantidad"], errors="coerce").round(2)

    if os.path.exists(OUTPUT_PATH) or str(OUTPUT_PATH).startswith("http") or "sharepoint" in str(OUTPUT_PATH).lower():
        try:
            df_prev = pd.read_excel(OUTPUT_PATH, sheet_name="Exhibiciones_implementadas")
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
                print("  ⚠️  Resultado previo sin Mes/Año — descartado en upsert.")
        except Exception as e:
            print(f"  ⚠️  No se pudo leer resultado previo en nube ({e}); se sobrescribe.")

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    df_out.to_excel(OUTPUT_PATH, index=False, sheet_name="Exhibiciones_implementadas")
    print(f"\n✓ Output escrito con éxito en SharePoint: {OUTPUT_PATH} ({len(df_out):,} filas)")
    
    m_gratis, m_pag = df_out["Categoría"].isin(["Gratis", "Concurso"]), df_out["Categoría"] == "Pagada"
    m_alto, m_medio = df_out["Nivel Impacto"] == "ALTO IMPACTO", df_out["Nivel Impacto"] == "MEDIO IMPACTO"

    print("\n── Resumen de Exhibiciones ──────────────────────")
    print(f"   Total Exhibiciones:                                    {df_out['Cantidad'].sum():>8,.0f}")
    print(f"   Total Exhibiciones Gratis:                             {df_out.loc[m_gratis, 'Cantidad'].sum():>8,.0f}")
    print(f"   Total Exhibiciones Pagadas:                            {df_out.loc[m_pag, 'Cantidad'].sum():>8,.0f}")
    print(f"\n   Total Exhibiciones Alto Impacto:                       {df_out.loc[m_alto, 'Cantidad'].sum():>8,.0f}")
    print(f"   Total Exhibiciones Medio Impacto:                      {df_out.loc[m_medio, 'Cantidad'].sum():>8,.0f}")
    print("─────────────────────────────────────────────────")

def generar_resumen_kpi_exhibiciones_gratis(spec: pr.PeriodoSpec):
    print(f"\n--- KPI (V3): RESUMEN EXHIBICIONES GRATIS por empleado (Nube) ({spec.etiqueta}) ---")
    ruta = OUTPUT_PATH
    if not os.path.exists(ruta) and not str(ruta).startswith("http") and "sharepoint" not in str(ruta).lower():
        print(f"❌ No existe en nube: {ruta}")
        return
    df = pd.read_excel(ruta, engine='openpyxl')
    df.columns = [str(c).strip() for c in df.columns]
    print(f"  Leyendo desde SharePoint: {os.path.basename(ruta)} ({len(df)} filas)")

    df = df[df['Categoría'].astype(str).str.strip().str.lower() == 'gratis'].copy()

    df['Mes'] = pd.to_numeric(df.get('Mes'), errors='coerce').fillna(0).astype(int)
    df['Año'] = pd.to_numeric(df.get('Año'), errors='coerce').fillna(0).astype(int)
    df['Cantidad'] = pd.to_numeric(df['Cantidad'], errors='coerce').fillna(0)
    df['Nivel Impacto'] = df['Nivel Impacto'].astype(str).str.strip().str.upper()

    df_periodo = df[(df['Mes'] == spec.mes) & (df['Año'] == spec.anio)].copy()
    if df_periodo.empty:
        raise ValueError(
            f"Exh Gratis KPI: no hay filas del periodo {spec.etiqueta} en SharePoint."
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
        path_dy_cupos = paths.DYP_BASE_CUPOS
        bc = pd.read_excel(path_dy_cupos, sheet_name="Tabla total roles")
        bc["NOMBRE_N"] = (
            bc["NOMBRE"].astype(str).str.strip().str.upper()
            .str.replace(r"\s+", " ", regex=True)
        )
        canal_por_nombre = dict(zip(
            bc["NOMBRE_N"],
            bc["CANAL"].astype(str).str.strip().str.upper(),
        ))
    except Exception as e:
        print(f"  ⚠️ No pude leer Base cupos para targets por canal ({e}); uso 3/5 por defecto.")
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

    path_kpis = f"{RUTA_CARPETA_SALIDAS_EXHIB}/KPI_Exhibiciones_Gratis.xlsx"
    os.makedirs(os.path.dirname(str(path_kpis)), exist_ok=True)
    resumen_activo.to_excel(str(path_kpis), index=False, engine='openpyxl')
    print(f"  ✅ Mes activo en SharePoint: {len(resumen_activo)} empleados → {os.path.basename(str(path_kpis))}")

    path_hist = f"{RUTA_CARPETA_SALIDAS_EXHIB}/KPI_Exhibiciones_Gratis_Historico.xlsx"
    if os.path.exists(str(path_hist)) or str(path_hist).startswith("http") or "sharepoint" in str(path_hist).lower():
        try:
            hist_prev = pd.read_excel(str(path_hist), engine='openpyxl')
            hist_prev['Mes'] = pd.to_numeric(hist_prev.get('Mes'), errors='coerce').fillna(0).astype(int)
            hist_prev['Año'] = pd.to_numeric(hist_prev.get('Año'), errors='coerce').fillna(0).astype(int)
        except Exception:
            hist_prev = pd.DataFrame(columns=cols_final)
    else:
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
    os.makedirs(os.path.dirname(str(path_hist)), exist_ok=True)
    hist_final.to_excel(str(path_hist), index=False, engine='openpyxl')
    print(f"  ✅ Histórico actualizado en SharePoint → {os.path.basename(str(path_hist))}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="ETL Exhibiciones Gratis — SharePoint / Azure")
    parser.add_argument("--solo", nargs="+", choices=["full", "kpi"],
                        help="Default: corre todo + kpi.")
    pr.cli_add_periodos_arg(parser)
    args = parser.parse_args()
    pasos = args.solo or ["full", "kpi"]
    specs = pr.periodos_de_args(args)

    if len(specs) > 1:
        print(f"🎯 ETL Exh Gratis (Cloud) — multi-periodo: {len(specs)} meses → "
              f"{', '.join(s.etiqueta for s in specs)}")

    for i, spec in enumerate(specs, 1):
        if len(specs) > 1:
            print(f"\n▶ Periodo cloud {i}/{len(specs)}: {spec.etiqueta}")
        print(f"\n🎯 ETL Exh Gratis (Cloud) — procesando periodo {spec.etiqueta} ({spec})")

        try:
            if "full" in pasos: run(spec)
            if "kpi"  in pasos: generar_resumen_kpi_exhibiciones_gratis(spec)
        except Exception as e:
            print(f"\n❌ ERROR en SharePoint para {spec.etiqueta}: {e}")
            if len(specs) > 1:
                print(f"   Continuando con los siguientes periodos...")
            else:
                raise