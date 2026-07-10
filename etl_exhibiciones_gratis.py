import os
import glob
import pandas as pd
import numpy as np

import paths
import periodo_resolver as pr

# ==============================================================================
# 1. CONFIGURACIÓN
# ==============================================================================
DATA_DIR = str(paths.EXHIB_DATA_DIR)

# Sprint 15.5.8 — Fix F6: ya NO usamos FILE_PATTERNS + glob. Los archivos se
# resuelven explícitamente por periodo vía periodo_resolver:
#   - exh_base_informar(spec)  → Base Exhibiciones Informar <Mes> <Año>.xlsx
#   - exh_base_planning(spec)  → Base Exhibiciones Planning <Mes> <Año>.xlsx
#   - exh_visitas(spec)        → informe-gerencial-visitas_YYYYMM.xlsx
#                                (mismo Involves que CIF; reemplaza el viejo
#                                "Visitas realizadas*.xlsx" que no existía)
# El "plan de trabajo" sigue siendo paths.CIF_OUT_FINAL (output del ETL CIF).

FILE_NIVEL_IMPACTO = str(paths.EXHIB_NIVEL_IMPACTO)
OUTPUT_PATH = str(paths.EXHIB_SALIDA / "Resultado exhibiciones gratis.xlsx")
PLAN_SHEET = "Plan de trabajo"

# Columnas clave encuesta
COL_ID_PDV          = "ID del PDV"
COL_EMPLEADO        = "Empleado"
COL_PERFIL_EMP      = "Perfil de acceso"
COL_FECHA           = "Fecha de la encuesta"
COL_MES             = "Mes del año"
COL_ANIO            = "Año"
COL_EXHIBICIONES    = "EXHIBICIONES:"

# Gratis / Concurso
COL_TIPO_EXHIB_GC   = "Seleccionar el Tipo de la exhibicion:"
COL_MARCA_GC        = "MARCA"
COL_CANTIDAD_GC     = "*Digite el numero de exhibiciones adicionales para este tipo"

# Pagadas
COL_TIPO_EXHIB_PAG  = "Seleccionar el Tipo de la exhibicion (Pagadas)"
COL_MARCA_PAG       = "MARCA.1"
COL_CANTIDAD_PAG    = "*Digite el numero de exhibiciones adicionales para este tipo."
COL_IMPLEMENTADA    = "La Exhibicion esta implementada de acuerdo con el planning?"
COL_CAUSAL          = "Indique las causales:"

# Contraprestación
COL_TIPO_CONTRA     = "Seleccionar el Tipo de la exhibicion - CONTRAPRESTACIÓN"
COL_MARCA_CONTRA    = "MARCA - CONTRAPRESTACIÓN"
COL_CANTIDAD_CONTRA = "*Digite el numero de exhibiciones adicionales para este tipo. - CONTRAPRESTACIÓN"

# Columnas Plan de Trabajo
COL_PLAN_ID_PDV     = "ID_PDV_INVOLVES"
COL_PLAN_ROL        = "ROL"
COL_PLAN_FREC       = "CANTIDAD_VISITAS" 
COL_PLAN_MES        = "MES"   
COL_PLAN_ANIO       = "AÑO"   

# Columnas Visitas Realizadas
COL_VIS_ID_PDV      = "ID del PDV"
COL_VIS_EMPLEADO    = "Empleado"
COL_VIS_FECHA       = "Fecha de la visita"

# Valores de negocio
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


"""
    Retorna el número de semana del mes para una fecha dada.
    Definición: semana 1 = lunes al domingo que contiene el día 1.
 """
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
    """
    Sprint 15.5.8 — Fix F6: resolución determinista por periodo.
    Sustituye al viejo discover_files() que matcheaba archivos por patrón.
    """
    groups: dict[str, list[str]] = {}

    groups["informar"]   = [str(pr.exh_base_informar(spec))]
    groups["planning"]   = [str(pr.exh_base_planning(spec))]
    groups["visitas"]    = [str(pr.exh_visitas(spec))]

    cif_out = paths.CIF_OUT_FINAL
    if cif_out.is_file():
        groups["plan_trabajo"] = [str(cif_out)]
    else:
        groups["plan_trabajo"] = []
        print(
            f"    ⚠ No se encontró {cif_out.name} — ejecuta primero el ETL CIF "
            f"({spec.etiqueta}); exhibiciones gratis no podrá computar frecuencias planeadas."
        )

    for tipo, archivos in groups.items():
        print(f"  [{tipo}] {len(archivos)} archivo(s):")
        for a in archivos:
            print(f"    - {os.path.basename(a)}")
    return groups

def load_encuestas(files: dict) -> pd.DataFrame:
    dfs = []
    for path in files.get("informar", []) + files.get("planning", []):
        try:
            df = pd.read_excel(path, sheet_name="report")
            dfs.append(df)
            print(f"    ✓ {os.path.basename(path)} — {len(df):,} filas")
        except Exception as e: print(f"    ⚠ Error leyendo {os.path.basename(path)}: {e}")
    if not dfs: raise ValueError("No se encontraron archivos de encuesta.")
    cols_comunes = set(dfs[0].columns)
    for df in dfs[1:]: cols_comunes &= set(df.columns)
    cols_comunes = [c for c in dfs[0].columns if c in cols_comunes]
    df = pd.concat([d[cols_comunes] for d in dfs], ignore_index=True)
    df[COL_FECHA] = pd.to_datetime(df[COL_FECHA], dayfirst=True, errors="coerce")
    df[COL_MES] = df[COL_FECHA].dt.month
    df[COL_ANIO] = df[COL_FECHA].dt.year
    df["_semana_mes"] = df[COL_FECHA].apply(_semana_del_mes)
    return df



    """
    Carga y consolida archivos de Plan de Trabajo.
    FILTRO: Solo ROL 'GESTOR'.
    COLUMNAS: ID_PDV, ROL, CANTIDAD_VISITAS, MES y AÑO.
    """


def load_plan(files: dict) -> pd.DataFrame:
    """
    Carga el PT (output de CIF) y filtra ROL=GESTOR.
    El archivo CIF/SALIDA/Plan de trabajo.xlsx ya viene con columnas
    UPPER_SNAKE_CASE y MES/AÑO/CANTIDAD_VISITAS calculados.
    """
    dfs = []
    for path in files.get("plan_trabajo", []):
        try:
            df = pd.read_excel(path)   # única hoja en el output de CIF
            df.columns = [str(c).strip().upper() for c in df.columns]
            if COL_PLAN_ROL in df.columns:
                mask = df[COL_PLAN_ROL].astype(str).str.upper().str.strip() == 'GESTOR'
                df = df[mask].copy()
            dfs.append(df)
            print(f"    ✓ {os.path.basename(path)} — {len(df):,} gestores cargados")
        except Exception as e:
            print(f"    ⚠ Error leyendo {os.path.basename(path)}: {e}")
    if not dfs: raise ValueError("No se encontraron registros de GESTOR.")
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

    """
    Carga y consolida todos los archivos de Visitas Realizadas encontrados. SIGUE IGUAL, SE USA ARCHIVO DE VISITAS DE INVOLVES COMO SE DESCARGA DE INVOLVES, SE CALCULAN LAS SEMANAS
    LA COMPARACION ES POR SEMANAS
    """
def load_visitas(files: dict) -> pd.DataFrame:
    """
    Sprint 15.5.8 — Fix F6.b: filtra 'Tipo de check-in' != 'Sin check-in'.

    Antes contaba TODAS las filas del Involves como visitas reales, incluidas
    las 'Sin check-in' (que son visitas planeadas no realizadas). En Mayo 2026
    eso era el 92.7% del archivo — el _frec_real salía inflado ~14x.

    Valores válidos en 'Tipo de check-in': Manual, Manual y GPS, GPS.
    """
    dfs = []
    for path in files.get("visitas", []):
        try:
            df = pd.read_excel(path)
            dfs.append(df)
            print(f"    ✓ {os.path.basename(path)} — {len(df):,} filas")
        except Exception as e:
            print(f"    ⚠ Error leyendo {os.path.basename(path)}: {e}")
    if not dfs:
        raise ValueError("No se encontraron archivos de Visitas (Involves).")
    df = pd.concat(dfs, ignore_index=True)

    # Fix F6.b: filtrar visitas no efectivas.
    if "Tipo de check-in" in df.columns:
        n_antes = len(df)
        df = df[df["Tipo de check-in"].astype(str).str.strip() != "Sin check-in"].copy()
        n_dropped = n_antes - len(df)
        print(f"    🔎 Filtro Tipo de check-in: descartadas {n_dropped:,} filas 'Sin check-in' "
              f"({n_dropped/n_antes*100:.1f}%); quedan {len(df):,} visitas efectivas.")
    else:
        print("    ⚠ Columna 'Tipo de check-in' no encontrada — no se aplicó filtro F6.b.")

    df[COL_VIS_FECHA] = pd.to_datetime(df[COL_VIS_FECHA], dayfirst=True, errors="coerce")
    df["_semana_mes"] = df[COL_VIS_FECHA].apply(_semana_del_mes)
    df[COL_VIS_ID_PDV] = pd.to_numeric(df[COL_VIS_ID_PDV], errors="coerce")
    df.dropna(subset=[COL_VIS_ID_PDV], inplace=True)
    df[COL_VIS_ID_PDV] = df[COL_VIS_ID_PDV].astype(int)
    return df[[COL_VIS_ID_PDV, COL_VIS_EMPLEADO, COL_VIS_FECHA, "_semana_mes"]]

    """
    Carga la tabla auxiliar Tipo Exhibición → Nivel Impacto (ruta fija).
    """
def load_nivel_impacto() -> pd.DataFrame:
    df = pd.read_excel(FILE_NIVEL_IMPACTO)
    df.columns = df.columns.str.strip()
    return df[["Tipo Exhibición", "Nivel Impacto"]]

# ==============================================================================
# 3. MÓDULOS DE LÓGICA (Gratis y Pagadas)
# ==============================================================================


"""
    Regla de cumplimiento por clave
    (ID PDV + EXHIBICIONES + Marca + Tipo Exhibición + Perfil Empleado + Empleado):
    - Frecuencia de referencia > 1 → necesita medición en >= 2 semanas distintas
    - Frecuencia de referencia == 1 → basta con 1 fecha registrada
    - GESTOR           → Frec Mensual del Plan de Trabajo
    - SUPERVISOR / GDD → conteo de visitas realizadas (ID PDV + Empleado)

    Cantidad = promedio de sumas semanales (no promedio directo de filas).
"""



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
        # No cumple — devolvemos el conteo real (sumando todas las semanas)
        # para poder mostrarlo en el detalle de "fuera de regla".
        return pd.Series({"Cantidad": float(grupo["Cantidad"].sum()),
                          "_cumple": False,
                          "_semanas_distintas": int(semanas_distintas),
                          "_frec_ref": float(frec_ref) if pd.notna(frec_ref) else 0.0})

    res = df.groupby(KEY, dropna=False).apply(evaluar_cumplimiento).reset_index()
    # Persistir SOLO las "fuera de regla" (Sprint 17.10) para el adjunto.
    fuera = res[res["_cumple"] == False].copy()
    if not fuera.empty:
        fuera = fuera.rename(columns={
            COL_EXHIBICIONES: "Categoría",
            COL_MARCA_GC: "Marca",
            COL_TIPO_EXHIB_GC: "Tipo Exhibición",
        })
        fuera["Categoría"] = fuera["Categoría"].str.replace("Exhibiciones ", "").str.strip()
        # Atajo: dejamos el archivo accesible a calcular_cumplimientos vía paths.
        import paths as _paths_persist
        ruta_fuera = _paths_persist.EXHIB_SALIDA / "Exh_Gratis_Fuera_de_Regla.xlsx"
        # Upsert por (Mes, Año) — si ya existe, mantenemos otros periodos.
        try:
            if ruta_fuera.exists():
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
            fuera_out.to_excel(ruta_fuera, index=False, engine="openpyxl")
        except Exception as ee:
            print(f"  ⚠️ No pude persistir Exh Gratis fuera de regla: {ee}")

    res = res[res["_cumple"] == True].drop(
        columns=["_cumple", "_semanas_distintas", "_frec_ref"]
    )
    res = res.rename(columns={COL_EXHIBICIONES: "Categoría", COL_MARCA_GC: "Marca", COL_TIPO_EXHIB_GC: "Tipo Exhibición"})
    res["Categoría"] = res["Categoría"].str.replace("Exhibiciones ", "").str.strip()
    return res



"""
    Reglas:
    - Implementada si ¿Implementada? == 'Si'
    - O si ¿Implementada? == 'No' Y Causal == 'contraprestación a la negociación'
    - Contraprestación → columnas _CONTRA para tipo, marca, cantidad
    - Sí → columnas originales pagadas
    - Output: una fila por ID PDV + Tipo Exhibición + Marca + Empleado (suma cantidades)
"""


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

"""
    Ejecuta la ETL completa sobre todos los archivos encontrados en DATA_DIR
    y escribe el consolidado en output_path.
"""


def run(spec: pr.PeriodoSpec):
    if not DATA_DIR:
        print(f"❌ No se encontró carpeta de exhibiciones: {DATA_DIR}")
        return
    print(f"── Resolviendo archivos fuente ({spec.etiqueta}) ─")
    files = resolve_files(spec)
    print("\n── Cargando archivos fuente ─────────────────────")
    df_enc = load_encuestas(files)
    df_plan = load_plan(files)
    df_vis = load_visitas(files)
    df_nivel = load_nivel_impacto()

    print(f"\n  Encuestas consolidadas: {len(df_enc):,} filas")
    print(f"  Plan de trabajo:         {len(df_plan):,} filas")
    print(f"  Visitas realizadas:      {len(df_vis):,} filas")

    print("\n── Módulo 1: Exhibiciones Pagadas ───────────────")
    df_pag = calcular_pagadas(df_enc)
    print(f"  Implementadas pagadas:          {len(df_pag):,} filas")

    print("\n── Módulo 2: Exhibiciones Gratis y Concurso ─────")
    df_gc = calcular_gratis_concurso(df_enc, df_plan, df_vis)
    print(f"  Implementadas gratis/concurso: {len(df_gc):,} filas")

    print("\n── Consolidando output ──────────────────────────")
    df_out = pd.concat([df_pag, df_gc], ignore_index=True)
    df_out = df_out.merge(df_nivel, on="Tipo Exhibición", how="left")
    df_out[COL_MES] = df_out[COL_MES].astype("Int64")
    df_out[COL_ANIO] = df_out[COL_ANIO].astype("Int64")
    df_out["Mes-Año"] = df_out[COL_MES].astype(str).str.zfill(2) + "-" + df_out[COL_ANIO].astype(str)
    
    df_out = df_out.rename(columns={COL_MES: "Mes", COL_ANIO: "Año", COL_ID_PDV: "ID PDV", COL_EMPLEADO: "Empleado", COL_PERFIL_EMP: "Rol Empleado"})
    df_out = df_out[OUTPUT_COLS].sort_values(["Año", "Mes", "ID PDV"])
    df_out["Cantidad"] = pd.to_numeric(df_out["Cantidad"], errors="coerce").round(2)

    # Sprint 17.20 — upsert por (Mes, Año): preservar periodos previos para
    # que GOLD/BI mantengan el historico cuando se corre solo un mes.
    if os.path.exists(OUTPUT_PATH):
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
                print("  ⚠️  Resultado exhibiciones gratis previo sin Mes/Año — descartado en upsert.")
        except Exception as e:
            print(f"  ⚠️  No se pudo leer Resultado exhibiciones gratis previo ({e}); se sobrescribe.")

    df_out.to_excel(OUTPUT_PATH, index=False, sheet_name="Exhibiciones_implementadas")
    print(f"\n✓ Output escrito en: {OUTPUT_PATH} ({len(df_out):,} filas, periodos preservados)")
    
    # Resumen final
    m_gratis, m_pag = df_out["Categoría"].isin(["Gratis", "Concurso"]), df_out["Categoría"] == "Pagada"
    m_alto, m_medio = df_out["Nivel Impacto"] == "ALTO IMPACTO", df_out["Nivel Impacto"] == "MEDIO IMPACTO"

    print("\n── Resumen de Exhibiciones ──────────────────────")
    print(f"  Total Exhibiciones:                        {df_out['Cantidad'].sum():>8,.0f}")
    print(f"  Total Exhibiciones Gratis:                 {df_out.loc[m_gratis, 'Cantidad'].sum():>8,.0f}")
    print(f"  Total Exhibiciones Pagadas:                {df_out.loc[m_pag, 'Cantidad'].sum():>8,.0f}")
    print(f"\n  Total Exhibiciones Alto Impacto:           {df_out.loc[m_alto, 'Cantidad'].sum():>8,.0f}")
    print(f"  Total Exhibiciones Medio Impacto:          {df_out.loc[m_medio, 'Cantidad'].sum():>8,.0f}")
    print("─────────────────────────────────────────────────")

def generar_resumen_kpi_exhibiciones_gratis(spec: pr.PeriodoSpec):
    """Resumen V3: solo categoría 'Gratis', pivot por (Empleado × Nivel Impacto).

    Sprint 15.5.8: filtra el archivo al periodo solicitado y aborta si no
    encuentra filas de ese mes (señal de que el `run()` no se corrió primero).
    """
    print(f"\n--- KPI (V3): RESUMEN EXHIBICIONES GRATIS por empleado  ({spec.etiqueta}) ---")
    import paths as _paths
    ruta = _paths.EXHIB_SALIDA / "Resultado exhibiciones gratis.xlsx"
    if not ruta.exists():
        print(f"❌ No existe: {ruta}")
        return
    df = pd.read_excel(ruta, engine='openpyxl')
    df.columns = [str(c).strip() for c in df.columns]
    print(f"  Leyendo: {ruta.name} ({len(df)} filas)")

    # Filtrar solo categoría Gratis
    df = df[df['Categoría'].astype(str).str.strip().str.lower() == 'gratis'].copy()

    # Normalizar Mes/Año (en el archivo vienen ya, asegurar tipo)
    df['Mes'] = pd.to_numeric(df.get('Mes'), errors='coerce').fillna(0).astype(int)
    df['Año'] = pd.to_numeric(df.get('Año'), errors='coerce').fillna(0).astype(int)
    df['Cantidad'] = pd.to_numeric(df['Cantidad'], errors='coerce').fillna(0)
    df['Nivel Impacto'] = df['Nivel Impacto'].astype(str).str.strip().str.upper()

    # Sprint 15.5.8: filtrar al periodo solicitado y validar que haya datos.
    df_periodo = df[(df['Mes'] == spec.mes) & (df['Año'] == spec.anio)].copy()
    if df_periodo.empty:
        raise ValueError(
            f"Exh Gratis KPI: no hay filas del periodo {spec.etiqueta} en "
            f"{ruta.name}. Corre `run(spec)` para este mes antes del KPI."
        )
    df = df_periodo

    # Pivot por (Mes, Año, Empleado) × Nivel Impacto
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
    # Asegurar columnas esperadas
    for col in ['ALTO IMPACTO', 'MEDIO IMPACTO']:
        if col not in resumen.columns:
            resumen[col] = 0

    # Sprint 17.9 — targets diferenciados por canal del gestor:
    #   PROXIMITY / PROXIMITY TAT             → 3 ALTO + 5 MEDIO
    #   DIRECTO, DROGUERIAS y combinados      → 2 ALTO + 8 MEDIO
    # Canal se obtiene de Base cupos (columna CANAL) cruzando por NOMBRE.
    try:
        bc = pd.read_excel(_paths.DYP_BASE_CUPOS, sheet_name="Tabla total roles")
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
        # Proximity puro o TAT → 3/5; cualquier otro (incluye combinados con DIRECTO) → 2/8.
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

    # Mes activo
    if not resumen.empty:
        anio_max = int(resumen['Año'].max())
        mes_max  = int(resumen[resumen['Año'] == anio_max]['Mes'].max())
        mask = (resumen['Año'] == anio_max) & (resumen['Mes'] == mes_max)
        resumen_activo = resumen[mask].copy()
    else:
        resumen_activo = resumen.copy()

    os.makedirs(os.path.dirname(str(_paths.EXHIB_GRA_OUT_KPIS)), exist_ok=True)
    resumen_activo.to_excel(str(_paths.EXHIB_GRA_OUT_KPIS), index=False, engine='openpyxl')
    print(f"  ✅ Mes activo: {len(resumen_activo)} empleados → {_paths.EXHIB_GRA_OUT_KPIS.name}")

    # Histórico upsert por (Mes, Año, Empleado)
    ruta_hist = str(_paths.EXHIB_GRA_OUT_KPIS_HISTORICO)
    if os.path.exists(ruta_hist):
        try:
            hist_prev = pd.read_excel(ruta_hist, engine='openpyxl')
            hist_prev['Mes'] = pd.to_numeric(hist_prev.get('Mes'), errors='coerce').fillna(0).astype(int)
            hist_prev['Año'] = pd.to_numeric(hist_prev.get('Año'), errors='coerce').fillna(0).astype(int)
        except Exception:
            hist_prev = pd.DataFrame(columns=cols_final)
    else:
        hist_prev = pd.DataFrame(columns=cols_final)

    # Sprint 15.9 (final fix) — upsert por periodo completo.
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
    hist_final.to_excel(ruta_hist, index=False, engine='openpyxl')
    print(f"  ✅ Histórico → {_paths.EXHIB_GRA_OUT_KPIS_HISTORICO.name}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="ETL Exhibiciones Gratis — Sprint 16.1: multi-periodo")
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
        print(f"\n🎯 ETL Exh Gratis — procesando periodo {spec.etiqueta} ({spec})")

        try:
            if "full" in pasos: run(spec)
            if "kpi"  in pasos: generar_resumen_kpi_exhibiciones_gratis(spec)
        except Exception as e:
            print(f"\n❌ ERROR en {spec.etiqueta}: {e}")
            if len(specs) > 1:
                print(f"   Continuando con siguientes periodos...")
            else:
                raise