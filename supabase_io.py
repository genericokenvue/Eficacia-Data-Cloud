"""
supabase_io.py
──────────────
Punto único de carga de los archivos de DETALLE a Supabase.

Por qué existe
──────────────
Antes cada ETL subía su tabla de KPIs con su propio bloque de código copiado
(normalización de columnas, limpieza de NaN, `upsert`). Eso tenía dos problemas:

  1. Se subía el KPI agregado, no el detalle — así no se puede analizar en BI
     más allá de lo que el ETL ya calculó.
  2. `upsert()` necesita una llave única. Las tablas de detalle no la tienen
     (un PDV puede aparecer varias veces), así que reejecutar el mismo mes
     duplicaba filas en lugar de reemplazarlas.

Este módulo resuelve las dos cosas con una estrategia de **reemplazo por
periodo**:

    DELETE FROM <tabla> WHERE mes = M AND anio = A
    INSERT  INTO <tabla> (...)   ← las filas nuevas

Con eso el comportamiento es el que se espera de una ETL que corre a diario:

  • Corres agosto hoy       → carga agosto.
  • Corres agosto mañana    → reemplaza agosto (no duplica).
  • Corres septiembre       → agrega septiembre, agosto queda intacto.
  • Vuelves a correr agosto → reemplaza solo agosto.

Uso
───
    import supabase_io

    supabase_io.cargar_detalle(
        tabla="cif_detalle",
        df=df_resultado,
        mes=spec.mes,
        anio=spec.anio,
    )
"""

from __future__ import annotations

import os
import re
import unicodedata
from datetime import datetime, timedelta, timezone

import numpy as np
import pandas as pd

try:
    from supabase import create_client
except ImportError:  # pragma: no cover
    create_client = None


# ─────────────────────────────────────────────────────────────────────────────
# CONEXIÓN
# ─────────────────────────────────────────────────────────────────────────────

_SUPABASE_URL = os.environ.get("SUPABASE_URL")
_SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

_cliente = None


def cliente():
    """Devuelve el cliente de Supabase (lo crea la primera vez)."""
    global _cliente
    if _cliente is None:
        if not create_client:
            raise RuntimeError("Falta la librería 'supabase' (pip install supabase)")
        if not _SUPABASE_URL or not _SUPABASE_KEY:
            raise RuntimeError("Faltan SUPABASE_URL / SUPABASE_KEY en el entorno")
        _cliente = create_client(_SUPABASE_URL, _SUPABASE_KEY)
    return _cliente


# ─────────────────────────────────────────────────────────────────────────────
# NORMALIZACIÓN DE NOMBRES DE COLUMNA
# ─────────────────────────────────────────────────────────────────────────────
# Postgres no distingue mayúsculas salvo que se citen las columnas, y las tildes
# o los espacios obligan a citarlas siempre. Se normaliza todo a snake_case ASCII.

# Postgres trunca en silencio cualquier identificador de más de 63 caracteres,
# lo que puede hacer que dos columnas distintas colapsen en la misma.
LIMITE_IDENTIFICADOR_PG = 63

# Traducciones explícitas. Tres motivos para estar en esta lista:
#   1. El resultado genérico sería confuso ('AÑO' → 'ao').
#   2. El nombre normalizado pasa de 63 caracteres (Exhibiciones Pagadas).
#   3. El archivo trae varias columnas que normalizan al mismo nombre
#      (Impactos tiene 7 columnas '% Efect' con sufijos '.1' y '_1', que
#      normalizan igual y se pisarían entre sí).
_RENOMBRES = {
    "AÑO": "anio",
    "Año": "anio",
    "AÑO_PLAN": "anio",
    "MES": "mes",
    "Mes": "mes",
    "Mes-Año": "mes_anio",
    "%_AGOTADOS": "pct_agotados",
    "ID del PDV": "id_pdv",
    "ID PDV": "id_pdv",
    "ID_PDV_INVOLVES": "id_pdv_involves",
    "ID INVOLVES": "id_pdv_involves",

    # Exhibiciones Pagadas — nombres que exceden el límite de Postgres
    "*Pagadas - *Digite el numero de exhibiciones adicionales para este tipo.":
        "cantidad_pagadas",
    "*Digite el numero de exhibiciones adicionales para este tipo. - CONTRAPRESTACIÓN":
        "cantidad_contraprestacion",
    "Seleccionar el Tipo de la exhibicion - CONTRAPRESTACIÓN":
        "tipo_contraprestacion",
    "La Exhibicion esta implementada de acuerdo con el planning?":
        "implementada_segun_planning",
    "*La exhibicion es PAC o ExtraPAC?": "pac_o_extrapac",
    "Exhibiciones adicionales - Planning E": "exhib_adicionales_planning",

    # Impactos D&P — 7 columnas '% Efect' que normalizarían al mismo nombre.
    # Se nombran por el indicador que las precede en el archivo.
    "% Efect":    "pct_efect_venta",
    "% Efect.1":  "pct_efect_visitas",
    "% Efect.2":  "pct_efect_impactos",
    "% Efect.3":  "pct_efect_georef",
    # Estas 3 vienen repetidas al final del archivo; su significado exacto
    # está por confirmar con negocio (ver nota en DOCS/supabase_tablas.sql).
    "% Efect_1":  "pct_efect_dup_1",
    "% Efect_2":  "pct_efect_dup_2",
    "% Efect_3":  "pct_efect_dup_3",
}

_MESES_A_NUM = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
    "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "SETIEMBRE": 9,
    "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12,
}


def normalizar_nombre_columna(nombre: str) -> str:
    """
    'ID del PDV' → 'id_pdv'   ·   '%_AGOTADOS' → 'pct_agotados'
    'CANTIDAD_VISITAS' → 'cantidad_visitas'   ·   'Año' → 'anio'
    """
    nombre = str(nombre).strip()
    if nombre in _RENOMBRES:
        return _RENOMBRES[nombre]

    # La ñ se translitera a 'n' antes de quitar tildes, para que 'AÑO' no
    # quede como 'ao' si el orden de las operaciones cambiara.
    txt = nombre.replace("Ñ", "N").replace("ñ", "n")
    txt = "".join(
        c for c in unicodedata.normalize("NFD", txt)
        if unicodedata.category(c) != "Mn"
    )
    txt = txt.replace("%", "pct").replace("&", "y")
    txt = re.sub(r"[^0-9a-zA-Z]+", "_", txt)
    txt = re.sub(r"_+", "_", txt).strip("_").lower()

    if not txt:
        txt = "col"
    if txt[0].isdigit():          # Postgres no acepta identificadores que arranquen en número
        txt = f"c_{txt}"

    # Recorte explícito: si se deja pasar, Postgres trunca por su cuenta y dos
    # columnas largas con el mismo prefijo terminan siendo la misma.
    if len(txt) > LIMITE_IDENTIFICADOR_PG:
        txt = txt[:LIMITE_IDENTIFICADOR_PG].rstrip("_")
    return txt


def normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """Renombra todas las columnas a snake_case y resuelve duplicados."""
    df = df.copy()
    nuevos, vistos = [], {}
    for col in df.columns:
        base = normalizar_nombre_columna(col)
        if base in vistos:
            vistos[base] += 1
            base = f"{base}_{vistos[base]}"
        else:
            vistos[base] = 0
        nuevos.append(base)
    df.columns = nuevos
    return df


# ─────────────────────────────────────────────────────────────────────────────
# LIMPIEZA PARA JSON
# ─────────────────────────────────────────────────────────────────────────────

# Valores que llegan como TEXTO y no son datos: aparecen cuando un cálculo
# dividió por cero y el resultado se escribió al CSV como la palabra 'inf'.
# Si la columna se declara numérica en Postgres, la inserción falla con ellos.
_TEXTOS_NULOS = {"inf", "-inf", "nan", "nat", "none", "null", "#¡div/0!", "#n/a", ""}


def _intentar_numerico(serie: pd.Series) -> pd.Series | None:
    """
    Convierte una columna de texto a número si TODOS sus valores lo permiten.

    Los CSV de D&P se escriben con coma decimal, pero si la columna quedó
    mezclada con textos ('inf'), pandas la lee como texto y no aplica la
    conversión: '196,100' se queda como cadena y Postgres la rechaza si la
    columna es numérica.

    Solo se convierte la coma en punto cuando hay UNA sola coma y ningún punto.
    Con eso '196,100' (coma decimal) se convierte, pero '30,000,000' (separador
    de miles) no: ahí la coma es ambigua y es preferible dejarlo como texto
    antes que inventar un valor.

    Devuelve la serie convertida, o None si no todos los valores son números.
    """
    txt = serie.astype(str).str.strip()
    vacios = txt.str.lower().isin(_TEXTOS_NULOS)
    utiles = txt[~vacios]
    if utiles.empty:
        return None

    una_coma = (utiles.str.count(",") == 1) & (~utiles.str.contains(r"\.", regex=True))
    normalizado = utiles.where(~una_coma, utiles.str.replace(",", ".", regex=False))

    convertido = pd.to_numeric(normalizado, errors="coerce")
    if convertido.isna().any():
        return None

    salida = pd.Series(np.nan, index=serie.index, dtype="float64")
    salida.loc[convertido.index] = convertido
    return salida


def preparar_registros(df: pd.DataFrame) -> list[dict]:
    """
    Convierte el DataFrame en la lista de diccionarios que espera la API REST.

    Pandas no admite `None` en columnas numéricas: si se limpia con
    `.where(pd.notnull(df), None)` vuelve a convertir esos None en NaN, y el
    serializador de JSON falla con "Out of range float values are not JSON
    compliant". Por eso hay que castear a `object` ANTES de sustituir los nulos.
    """
    limpio = df.replace([np.inf, -np.inf], np.nan)

    # ── Pasada 1: normalizar tipos ───────────────────────────────────────
    for col in limpio.columns:
        # Las fechas no son serializables: se pasan a texto ISO.
        if pd.api.types.is_datetime64_any_dtype(limpio[col]):
            limpio[col] = limpio[col].dt.strftime("%Y-%m-%d %H:%M:%S")
        elif limpio[col].dtype == object:
            # Columna de texto que en realidad es numérica (coma decimal, o
            # enteros que quedaron como '31568062.0' tras pasar por Excel).
            numerica = _intentar_numerico(limpio[col])
            if numerica is not None:
                limpio[col] = numerica
                continue
            # Si no, al menos anular los 'inf' que llegaron como palabra.
            mask = limpio[col].astype(str).str.strip().str.lower().isin(_TEXTOS_NULOS)
            if mask.any():
                limpio.loc[mask, col] = np.nan

    # ── Pasada 2: decimales que en realidad son enteros ──────────────────
    #
    # Va DESPUÉS de la pasada 1 a propósito. Una columna entera puede llegar
    # como decimal por dos caminos distintos:
    #   • el concat del histórico con datos nuevos, si hay algún nulo
    #     (pandas asciende int64 → float64 y queda 31568062.0)
    #   • texto '31568062.0' que la pasada 1 acaba de convertir a float
    # Si esto se hiciera en la misma pasada, el segundo caso se escaparía.
    #
    # Postgres rechaza el decimal en una columna bigint:
    #     invalid input syntax for type bigint: "31568062.0"
    # Mandar el entero es seguro tanto para bigint como para numeric.
    for col in limpio.columns:
        if pd.api.types.is_float_dtype(limpio[col]):
            no_nulos = limpio[col].dropna()
            if not no_nulos.empty and (no_nulos % 1 == 0).all():
                limpio[col] = limpio[col].astype("Int64")

    limpio = limpio.astype(object).where(pd.notnull(limpio), None)
    return limpio.to_dict(orient="records")


# ─────────────────────────────────────────────────────────────────────────────
# MARCA DE TIEMPO DE LA CARGA
# ─────────────────────────────────────────────────────────────────────────────
# Colombia es UTC-5 todo el año (no aplica horario de verano desde 1993), así
# que un desfase fijo es exacto y evita depender del paquete `tzdata`, que no
# viene instalado por defecto en Windows ni en algunas imágenes de Linux.
#
# Se envía como texto sin zona horaria y la columna se declara `timestamp`
# (no `timestamptz`): así lo que se guarda es literalmente la hora de Bogotá.
# Con `timestamptz` Postgres normaliza a UTC y la consola de Supabase la
# muestra 5 horas adelantada, que es justo lo que se quería corregir.

ZONA_COLOMBIA = timezone(timedelta(hours=-5))
COLUMNA_CARGA = "cargado_en"


def ahora_colombia() -> str:
    """Hora local de Colombia como 'YYYY-MM-DD HH:MM:SS'."""
    return datetime.now(ZONA_COLOMBIA).strftime("%Y-%m-%d %H:%M:%S")


def _a_numero_mes(valor) -> int | None:
    """Acepta 7, '7', '07', 'Julio', 'JULIO' → 7."""
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return None
    txt = str(valor).strip()
    if not txt:
        return None
    if txt.isdigit():
        return int(txt)
    return _MESES_A_NUM.get(
        "".join(
            c for c in unicodedata.normalize("NFD", txt.upper())
            if unicodedata.category(c) != "Mn"
        )
    )


# ─────────────────────────────────────────────────────────────────────────────
# CARGA CON REEMPLAZO POR PERIODO
# ─────────────────────────────────────────────────────────────────────────────

# Tamaño de lote para el INSERT. Consolidado_Ventas ronda las 170.000 filas por
# mes, así que con lotes de 500 serían ~350 peticiones. Con 1.000 se reduce a la
# mitad sin acercarse al límite de tamaño de petición de PostgREST.
TAMANO_LOTE = 1000

# A partir de este número de filas se imprime el avance: una carga de 170.000
# filas tarda varios minutos y sin señal parece que el proceso se colgó.
UMBRAL_AVISO_PROGRESO = 20_000


# Errores de red o de carga del servidor, no de los datos. Se reintentan.
_ERRORES_TRANSITORIOS = (
    "server disconnected", "connection reset", "connection aborted",
    "timeout", "timed out", "temporarily unavailable", "bad gateway",
    "service unavailable", "502", "503", "504",
)

REINTENTOS_MAX = 4


def _borrar_periodo_con_reintentos(sb, tabla: str, mes: int, anio: int) -> int:
    """
    Borra las filas del periodo, reintentando si Postgres corta la sentencia.

    Con tablas grandes el DELETE puede pasarse del límite de tiempo del
    servidor ('canceling statement due to statement timeout'). Se reintenta con
    espera creciente: parte del trabajo suele quedar hecha en cada intento, así
    que los siguientes tienen menos filas que borrar y terminan más rápido.
    """
    import time

    total = 0
    espera = 3.0
    for intento in range(1, REINTENTOS_MAX + 1):
        try:
            resp = sb.table(tabla).delete().eq("mes", mes).eq("anio", anio).execute()
            return total + len(getattr(resp, "data", None) or [])
        except Exception as e:
            msg = str(e).lower()
            es_timeout = "timeout" in msg or "57014" in msg or "canceling statement" in msg
            if not es_timeout or intento == REINTENTOS_MAX:
                raise
            print(f"     ⏳ {tabla}: el borrado de {mes:02d}/{anio} superó el tiempo límite "
                  f"({intento}/{REINTENTOS_MAX}) — reintento en {espera:.0f}s")
            time.sleep(espera)
            espera *= 2
    return total


def _insertar_bloque_con_reintentos(sb, tabla: str, bloque: list, offset: int,
                                    datos: pd.DataFrame) -> None:
    """
    Inserta un lote reintentando ante fallos de red.

    Con tablas grandes (Ventas ronda las 50.000 filas por mes) la API corta la
    conexión de vez en cuando: 'Server disconnected'. No es un problema de los
    datos, así que abortar toda la carga por eso desperdicia el trabajo hecho.
    Se reintenta con espera creciente y, si sigue fallando, se parte el lote a
    la mitad — a veces el bloque es sencillamente muy grande para la petición.
    """
    import time

    espera = 2.0
    for intento in range(1, REINTENTOS_MAX + 1):
        try:
            sb.table(tabla).insert(bloque).execute()
            return
        except Exception as e:
            msg = str(e).lower()
            transitorio = any(t in msg for t in _ERRORES_TRANSITORIOS)

            # Un error de tipo o de columna no se arregla reintentando.
            if not transitorio or intento == REINTENTOS_MAX:
                if transitorio and len(bloque) > 50:
                    # Último recurso: partir en dos y reintentar cada mitad.
                    medio = len(bloque) // 2
                    print(f"     ↪ lote de {len(bloque)} partido en dos tras {intento} intentos")
                    _insertar_bloque_con_reintentos(sb, tabla, bloque[:medio], offset, datos)
                    _insertar_bloque_con_reintentos(sb, tabla, bloque[medio:], offset + medio, datos)
                    return
                raise RuntimeError(
                    f"Error insertando en '{tabla}' (filas {offset}-{offset+len(bloque)}): {e}\n"
                    f"  Columnas enviadas: {list(datos.columns)}"
                ) from e

            print(f"     ⏳ {tabla}: fallo de red en filas {offset}-{offset+len(bloque)} "
                  f"({intento}/{REINTENTOS_MAX}) — reintento en {espera:.0f}s")
            time.sleep(espera)
            espera *= 2


def cargar_detalle(
    tabla: str,
    df: pd.DataFrame,
    mes: int,
    anio: int,
    *,
    solo_periodo: bool = True,
    lote: int = TAMANO_LOTE,
) -> int:
    """
    Sube `df` a `tabla` reemplazando por completo el periodo (mes, anio).

    Args:
        tabla:  nombre de la tabla en Supabase.
        df:     DataFrame de detalle. Debe traer columnas de mes y año
                (MES/AÑO, Mes/Año, o ya normalizadas mes/anio).
        mes:    mes del periodo procesado (1-12).
        anio:   año del periodo procesado.
        solo_periodo: si el DataFrame es multi-periodo (p. ej. CIF.xlsx, que
                acumula todos los meses), deja solo las filas del periodo
                indicado antes de subir. Así una corrida de agosto no reescribe
                julio.

    Returns:
        Número de filas insertadas.
    """
    if df is None or df.empty:
        print(f"  ℹ️  {tabla}: DataFrame vacío, no se sube nada.")
        return 0

    datos = normalizar_columnas(df)

    # ── Asegurar columnas de periodo numéricas ───────────────────────────
    if "mes" not in datos.columns:
        datos["mes"] = mes
    else:
        datos["mes"] = datos["mes"].map(_a_numero_mes)
        datos["mes"] = datos["mes"].fillna(mes)

    if "anio" not in datos.columns:
        datos["anio"] = anio
    else:
        datos["anio"] = pd.to_numeric(datos["anio"], errors="coerce").fillna(anio)

    datos["mes"] = pd.to_numeric(datos["mes"], errors="coerce").fillna(mes).astype(int)
    datos["anio"] = pd.to_numeric(datos["anio"], errors="coerce").fillna(anio).astype(int)

    # ── Quedarse solo con el periodo procesado ───────────────────────────
    if solo_periodo:
        antes = len(datos)
        datos = datos[(datos["mes"] == int(mes)) & (datos["anio"] == int(anio))]
        if len(datos) != antes:
            print(f"  ℹ️  {tabla}: {antes} filas en el archivo → {len(datos)} del periodo {mes:02d}/{anio}")
        if datos.empty:
            print(f"  ⚠️  {tabla}: el archivo no tiene filas de {mes:02d}/{anio}. No se sube nada.")
            return 0

    # ── Marca de tiempo de la carga (hora de Colombia) ───────────────────
    datos[COLUMNA_CARGA] = ahora_colombia()

    sb = cliente()

    # ── 1) Borrar el periodo ─────────────────────────────────────────────
    # Se hace antes de insertar para que reejecutar el mismo mes reemplace en
    # vez de duplicar. Las tablas de detalle no tienen llave única, así que
    # `upsert()` no sirve aquí.
    try:
        n_borradas = _borrar_periodo_con_reintentos(sb, tabla, int(mes), int(anio))
        # Se informa el número porque un DELETE bloqueado por Row Level Security
        # NO da error: borra 0 filas y se ve idéntico a "no había nada que borrar".
        print(f"  🗑️  {tabla}: {n_borradas:,} fila(s) del periodo {mes:02d}/{anio} eliminadas")
    except Exception as e:
        raise RuntimeError(
            f"No se pudo limpiar el periodo {mes:02d}/{anio} en '{tabla}': {e}\n"
            f"  Si la tabla no existe todavía, créala con el script SQL "
            f"(ver DOCS/supabase_tablas.sql)."
        ) from e

    # ── 2) Insertar por lotes ────────────────────────────────────────────
    registros = preparar_registros(datos)
    total = len(registros)
    mostrar_avance = total >= UMBRAL_AVISO_PROGRESO
    if mostrar_avance:
        print(f"  ⏳ {tabla}: subiendo {total:,} filas en lotes de {lote:,}…")

    insertadas = 0
    for i in range(0, total, lote):
        bloque = registros[i:i + lote]
        _insertar_bloque_con_reintentos(sb, tabla, bloque, i, datos)
        insertadas += len(bloque)

        if mostrar_avance and (insertadas % (lote * 25) == 0 or insertadas == total):
            print(f"     {insertadas:,}/{total:,} ({insertadas*100//total}%)")

    print(f"  ✅ {tabla}: {insertadas:,} filas cargadas ({mes:02d}/{anio})")
    return insertadas


def cargar_detalle_seguro(tabla: str, df: pd.DataFrame, mes: int, anio: int, **kw) -> int:
    """
    Igual que `cargar_detalle` pero no interrumpe el ETL si Supabase falla.

    El archivo de SharePoint ya quedó escrito en ese punto; que la base de
    datos esté caída no debería tumbar toda la corrida.
    """
    try:
        return cargar_detalle(tabla, df, mes, anio, **kw)
    except Exception as e:
        print(f"  ⚠️  {tabla}: no se pudo cargar a Supabase — {e}")
        return 0
