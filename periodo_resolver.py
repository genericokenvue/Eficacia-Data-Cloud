"""
periodo_resolver.py
─────────────────────
Único punto de verdad para mapear (mes, año) → archivos concretos en BASES/.

Diseño Sprint 15.5:
  - Cada ETL recibe --mes M --anio AAAA como argumentos REQUERIDOS.
  - El ETL llama `resolver(mes, anio)` para obtener un PeriodoSpec.
  - Para localizar archivos, usa las funciones específicas del módulo
    (cif_pt_directo, np_respuestas, etc.). Cada una valida que el archivo
    exista y falla con error explícito si no.
  - Esto sustituye la "detección automática del último periodo" que generaba
    salidas mal-fechadas en uniperiodo y silenciosamente perdía datos en
    multiperiodo.

Principios:
  • Fail-fast: si no encuentra el archivo del mes solicitado, error claro.
  • No fallback al "más reciente": queremos errores explícitos, no datos viejos.
  • Patrones literales: si el nombre del archivo no encaja, el operador
    debe renombrarlo o el código debe actualizarse. No adivinar.

Convención de nombres de archivo (basada en BASES/ producción):
  CIF
    PLAN DE TRABAJO/Plan de trabajo Directo {Mes} {Año}.xlsx
    PLAN DE TRABAJO/Plan de trabajo ISM     {Mes} {Año}.xlsx
    INVOLVES/informe-gerencial-visitas_{YYYYMM}.xlsx
  NO PRESENCIA
    Respuestas no presencia {Mes} {Año}.xlsx
  PRECIOS
    Respuestas Precios {Mes} {Año}.xlsx
    Plan_Trabajo_Precios_{MES_UPPER}_{Año}.xlsx
  SOS
    Respuestas SOS {Mes} {Año}.xlsx
    Respuestas SOS Baby {Mes} {Año}.xlsx
    Colombia-sos-target-{YYYYMM}.xlsx
  EXHIBICIONES
    Base Exhibiciones Informar {Mes} {Año}.xlsx
    Base Exhibiciones Planning {Mes} {Año}.xlsx
  D&P
    01. Ventas/{Año}/{NN}. {Mes} base ventas.xlsx
    02. Impactos/{Año}/{NN}. {Mes} base impactos - cuota.xlsx
    Rutero/RUTERO {MES_UPPER} {Año} D&P.xlsx
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

import paths

# ─────────────────────────────────────────────────────────────────────────────
# Tabla de meses (español como aparece en los nombres de archivo)
# ─────────────────────────────────────────────────────────────────────────────

MESES_ES = {
    1:  "Enero",
    2:  "Febrero",
    3:  "Marzo",
    4:  "Abril",
    5:  "Mayo",
    6:  "Junio",
    7:  "Julio",
    8:  "Agosto",
    9:  "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}


# ─────────────────────────────────────────────────────────────────────────────
# Spec de periodo
# ─────────────────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class PeriodoSpec:
    """
    Identifica un periodo (mes, año) y expone las variantes de formato más
    comunes (string en español, mayúsculas, YYYYMM, etiqueta).
    """
    mes: int
    anio: int

    @property
    def mes_str(self) -> str:
        """'Mayo'"""
        return MESES_ES[self.mes]

    @property
    def mes_str_upper(self) -> str:
        """'MAYO'"""
        return self.mes_str.upper()

    @property
    def yyyymm(self) -> str:
        """'202605'"""
        return f"{self.anio}{self.mes:02d}"

    @property
    def etiqueta(self) -> str:
        """'Mayo_2026'"""
        return f"{self.mes_str}_{self.anio}"

    def __str__(self) -> str:
        return f"{self.mes:02d}/{self.anio}"


def resolver(mes: int, anio: int) -> PeriodoSpec:
    """
    Construye un PeriodoSpec validado.

    Raises:
        ValueError: si mes ∉ [1,12] o año fuera de [2020, 2099].
    """
    if not isinstance(mes, int) or not (1 <= mes <= 12):
        raise ValueError(f"Mes inválido: {mes!r} (esperado entero 1-12)")
    if not isinstance(anio, int) or not (2020 <= anio <= 2099):
        raise ValueError(f"Año inválido: {anio!r} (esperado entero 2020-2099)")
    return PeriodoSpec(mes=mes, anio=anio)


def parsear_periodo_str(s: str) -> PeriodoSpec:
    """
    Parsea 'MM/YYYY' o 'MM-YYYY' a PeriodoSpec.

    Ejemplo:
        parsear_periodo_str('05/2026') → PeriodoSpec(mes=5, anio=2026)
    """
    s = s.strip().replace("-", "/")
    if "/" not in s:
        raise ValueError(f"Formato de periodo esperado MM/YYYY, recibido: {s!r}")
    m_str, a_str = s.split("/", 1)
    return resolver(int(m_str), int(a_str))


# ─────────────────────────────────────────────────────────────────────────────
# Helper de localización fail-fast
# ─────────────────────────────────────────────────────────────────────────────

def _find_unico(
    carpeta: Path,
    patron_glob: str,
    *,
    contexto: str,
) -> Path:
    """
    Localiza UN único archivo en `carpeta` que matchee `patron_glob`.
    Falla con FileNotFoundError si no hay ninguno.
    Falla con RuntimeError si hay más de uno (situación ambigua que requiere
    intervención del operador).
    Ignora archivos temporales de Excel (`~$*`).
    """
    if not carpeta.is_dir():
        raise FileNotFoundError(
            f"[{contexto}] Carpeta no existe: {carpeta}"
        )
    matches = sorted(
        p for p in carpeta.glob(patron_glob)
        if not p.name.startswith("~$")
    )
    if not matches:
        raise FileNotFoundError(
            f"[{contexto}] No se encontró archivo con patrón '{patron_glob}' en {carpeta}"
        )
    if len(matches) > 1:
        nombres = [m.name for m in matches]
        raise RuntimeError(
            f"[{contexto}] Múltiples coincidencias para patrón '{patron_glob}' "
            f"en {carpeta}: {nombres}. Esperaba exactamente uno."
        )
    return matches[0]


# ─────────────────────────────────────────────────────────────────────────────
# CIF
# ─────────────────────────────────────────────────────────────────────────────

def cif_pt_directo(spec: PeriodoSpec) -> Path:
    return _find_unico(
        paths.CIF_PT_DIR,
        f"Plan de trabajo Directo {spec.mes_str} {spec.anio}.xlsx",
        contexto=f"CIF/PT Directo {spec.etiqueta}",
    )


def cif_pt_ism(spec: PeriodoSpec) -> Path:
    return _find_unico(
        paths.CIF_PT_DIR,
        f"Plan de trabajo ISM {spec.mes_str} {spec.anio}.xlsx",
        contexto=f"CIF/PT ISM {spec.etiqueta}",
    )


def cif_involves(spec: PeriodoSpec) -> Path:
    """informe-gerencial-visitas_YYYYMM.xlsx — visitas Involves del periodo."""
    return _find_unico(
        paths.CIF_BASES / "INVOLVES",
        f"informe-gerencial-visitas_{spec.yyyymm}.xlsx",
        contexto=f"CIF/INVOLVES {spec.etiqueta}",
    )


# ─────────────────────────────────────────────────────────────────────────────
# NO PRESENCIA
# ─────────────────────────────────────────────────────────────────────────────

def np_respuestas(spec: PeriodoSpec) -> Path:
    return _find_unico(
        paths.NP_BASES,
        f"Respuestas no presencia {spec.mes_str} {spec.anio}.xlsx",
        contexto=f"NP/Respuestas {spec.etiqueta}",
    )


# ─────────────────────────────────────────────────────────────────────────────
# PRECIOS
# ─────────────────────────────────────────────────────────────────────────────

def precios_respuestas(spec: PeriodoSpec) -> Path:
    return _find_unico(
        paths.PR_BASES,
        f"Respuestas Precios {spec.mes_str} {spec.anio}.xlsx",
        contexto=f"Precios/Respuestas {spec.etiqueta}",
    )


# Nota Sprint 15.5: 'Plan_Trabajo_Precios_<MES>_<AÑO>.xlsx' NO es input —
# es salida intermedia generada por el paso 1 del ETL Precios a partir
# del PT consolidado de CIF. Por eso no lo expone este resolver.


# ─────────────────────────────────────────────────────────────────────────────
# SOS
# ─────────────────────────────────────────────────────────────────────────────

def sos_respuestas(spec: PeriodoSpec) -> Path:
    return _find_unico(
        paths.SOS_BASES,
        f"Respuestas SOS {spec.mes_str} {spec.anio}.xlsx",
        contexto=f"SOS/Respuestas {spec.etiqueta}",
    )


def sos_respuestas_baby(spec: PeriodoSpec) -> Path:
    return _find_unico(
        paths.SOS_BASES,
        f"Respuestas SOS Baby {spec.mes_str} {spec.anio}.xlsx",
        contexto=f"SOS/Respuestas Baby {spec.etiqueta}",
    )


def sos_target(spec: PeriodoSpec) -> Path:
    return _find_unico(
        paths.SOS_BASES,
        f"Colombia-sos-target-{spec.yyyymm}.xlsx",
        contexto=f"SOS/Target {spec.etiqueta}",
    )


# ─────────────────────────────────────────────────────────────────────────────
# EXHIBICIONES
# ─────────────────────────────────────────────────────────────────────────────
#
# La carpeta EXHIBICIONES tiene DOS familias de archivos por mes:
#
#   PLANNING DE <MES> <AÑO>.xlsx
#       Planning "maestro" (origen). Columnas con prefijo asterisco:
#       *Punto de venta, *Tipo - Pagadas, *MARCA - Pagadas, etc.
#       Es lo que Exh Pagadas usa como ruta_p (planning real).
#
#   Base Exhibiciones Planning <Mes> <Año>.xlsx
#       Descarga de Involves con las respuestas de las encuestas.
#       Columnas sin prefijo: ID del PDV, PDV, etc.
#       Exh Pagadas la usa como ruta_i (involves planning).
#
#   Base Exhibiciones Informar <Mes> <Año>.xlsx
#       Descarga de Involves de la encuesta "Informar" (exhibiciones gratis).

def exh_planning_master(spec: PeriodoSpec) -> Path:
    """Planning maestro de exhibiciones — PLANNING DE <MES> <AÑO>.xlsx."""
    return _find_unico(
        paths.EXHIB_DATA_DIR,
        f"PLANNING DE {spec.mes_str_upper} {spec.anio}.xlsx",
        contexto=f"Exh/Planning maestro {spec.etiqueta}",
    )


def exh_base_planning(spec: PeriodoSpec) -> Path:
    """Involves de planning — Base Exhibiciones Planning <Mes> <Año>.xlsx."""
    return _find_unico(
        paths.EXHIB_DATA_DIR,
        f"Base Exhibiciones Planning {spec.mes_str} {spec.anio}.xlsx",
        contexto=f"Exh/Base Planning {spec.etiqueta}",
    )


def exh_base_informar(spec: PeriodoSpec) -> Path:
    """Involves de informar — Base Exhibiciones Informar <Mes> <Año>.xlsx."""
    return _find_unico(
        paths.EXHIB_DATA_DIR,
        f"Base Exhibiciones Informar {spec.mes_str} {spec.anio}.xlsx",
        contexto=f"Exh/Base Informar {spec.etiqueta}",
    )


def exh_visitas(spec: PeriodoSpec) -> Path:
    """
    Las "visitas realizadas" para Exh Gratis se obtienen del mismo informe
    Involves que CIF (en BASES/CIF/INVOLVES/). Sprint 15.5.8 unifica esto.
    """
    return cif_involves(spec)


# ─────────────────────────────────────────────────────────────────────────────
# D&P  (Ventas + Impactos + Rutero)
# ─────────────────────────────────────────────────────────────────────────────

def dyp_ventas(spec: PeriodoSpec) -> Path:
    return _find_unico(
        paths.DYP_VENTAS_DIR / str(spec.anio),
        f"{spec.mes:02d}. {spec.mes_str} base ventas.xlsx",
        contexto=f"D&P/Ventas {spec.etiqueta}",
    )


def dyp_impactos(spec: PeriodoSpec) -> Path:
    return _find_unico(
        paths.DYP_IMPACTOS_DIR / str(spec.anio),
        f"{spec.mes:02d}. {spec.mes_str} base impactos - cuota.xlsx",
        contexto=f"D&P/Impactos {spec.etiqueta}",
    )


def dyp_rutero(spec: PeriodoSpec, *, fallback_a_legacy: bool = True) -> Path:
    """
    Rutero D&P del periodo.
    Si no existe el del mes y `fallback_a_legacy=True`, intenta
    'Rutero_Droguerias.xlsx' (legacy estable). Esta concesión existe porque
    en la práctica el rutero solo se actualiza ocasionalmente.
    """
    try:
        return _find_unico(
            paths.DYP_BASES / "Rutero",
            f"RUTERO {spec.mes_str_upper} {spec.anio} D&P.xlsx",
            contexto=f"D&P/Rutero {spec.etiqueta}",
        )
    except FileNotFoundError:
        if not fallback_a_legacy:
            raise
        legacy = paths.DYP_BASES / "Rutero" / "Rutero_Droguerias.xlsx"
        if legacy.is_file():
            return legacy
        raise


# ─────────────────────────────────────────────────────────────────────────────
# CLI helper — para que cada ETL exponga --mes/--anio de manera uniforme
# ─────────────────────────────────────────────────────────────────────────────

def cli_add_periodo_args(parser: argparse.ArgumentParser) -> None:
    """
    Registra --mes (1-12) y --anio (AAAA) como argumentos REQUERIDOS en el
    parser dado. Sprint 15.5 obligatoriedad: ya no hay default de "último
    mes detectado" — quien corre el ETL debe decir qué mes quiere.

    Mono-periodo. Para multi-periodo ver cli_add_periodos_arg.
    """
    parser.add_argument(
        "--mes", type=int, required=True, choices=range(1, 13), metavar="M",
        help="Mes a procesar (1-12). Obligatorio.",
    )
    parser.add_argument(
        "--anio", type=int, required=True, metavar="AAAA",
        help="Año a procesar (ej. 2026). Obligatorio.",
    )


def periodo_de_args(args: argparse.Namespace) -> PeriodoSpec:
    """Construye PeriodoSpec desde argparse.Namespace con .mes y .anio."""
    return resolver(int(args.mes), int(args.anio))


# ── Sprint 16.1 — Multi-periodo ──────────────────────────────────────────────

def cli_add_periodos_arg(parser: argparse.ArgumentParser) -> None:
    """
    Registra --mes/--anio y --periodos en un esquema NO requerido por argparse
    pero con validación en periodos_de_args que exige exactamente uno de los
    dos modos:

      Mono-periodo:   --mes M --anio AAAA
      Multi-periodo:  --periodos M/AAAA,N/BBBB,O/CCCC

    Usar este helper en orquestadores (run_all.py) y ETLs que admiten varios
    meses en una corrida. Para CLI mono-periodo simple, usar cli_add_periodo_args.
    """
    parser.add_argument(
        "--mes", type=int, required=False, choices=range(1, 13), metavar="M",
        help="Mes a procesar (1-12). Usar junto a --anio para mono-periodo.",
    )
    parser.add_argument(
        "--anio", type=int, required=False, metavar="AAAA",
        help="Año a procesar (ej. 2026). Usar junto a --mes.",
    )
    parser.add_argument(
        "--periodos", type=str, required=False, metavar="LISTA",
        help=(
            "Lista de periodos separados por coma. Formato MM/AAAA. "
            "Ejemplo: --periodos 03/2026,04/2026,05/2026"
        ),
    )


def periodos_de_args(args: argparse.Namespace) -> "list[PeriodoSpec]":
    """
    Construye una lista de PeriodoSpec desde args registrados con
    cli_add_periodos_arg.

    Reglas:
      - Si se pasa --periodos: parsea cada token (MM/AAAA). Lista de 1+.
      - Si se pasan --mes Y --anio: lista de 1 elemento.
      - --periodos es excluyente con --mes/--anio: pasar ambos → error.
      - No pasar ninguno → error.
      - Duplicados en --periodos se preservan (caller decide deduplicar).
    """
    tiene_periodos = getattr(args, "periodos", None) is not None
    tiene_mes      = getattr(args, "mes", None) is not None
    tiene_anio     = getattr(args, "anio", None) is not None

    if tiene_periodos and (tiene_mes or tiene_anio):
        raise SystemExit(
            "Error: --periodos es excluyente con --mes/--anio. Pasá uno u otro."
        )

    if tiene_periodos:
        tokens = [t.strip() for t in str(args.periodos).split(",") if t.strip()]
        if not tokens:
            raise SystemExit("Error: --periodos vacío. Formato: MM/AAAA[,MM/AAAA,...]")
        return [parsear_periodo_str(t) for t in tokens]

    if tiene_mes and tiene_anio:
        return [resolver(int(args.mes), int(args.anio))]

    raise SystemExit(
        "Error: especificá el periodo con --mes M --anio AAAA "
        "o con --periodos MM/AAAA[,MM/AAAA,...]"
    )


# ─────────────────────────────────────────────────────────────────────────────
# Diagnóstico
# ─────────────────────────────────────────────────────────────────────────────

def diagnostico(spec: PeriodoSpec) -> str:
    """
    Devuelve un string que muestra qué archivos resuelve el periodo dado y
    cuáles existen. Útil para que el operador valide el setup antes de correr.
    """
    resolvers = [
        ("CIF.PT_Directo",     lambda: cif_pt_directo(spec)),
        ("CIF.PT_ISM",         lambda: cif_pt_ism(spec)),
        ("CIF.Involves",       lambda: cif_involves(spec)),
        ("NP.Respuestas",      lambda: np_respuestas(spec)),
        ("Precios.Respuestas", lambda: precios_respuestas(spec)),
        ("SOS.Respuestas",     lambda: sos_respuestas(spec)),
        ("SOS.Respuestas_Baby",lambda: sos_respuestas_baby(spec)),
        ("SOS.Target",         lambda: sos_target(spec)),
        ("Exh.Planning_master",lambda: exh_planning_master(spec)),
        ("Exh.Base_Planning",  lambda: exh_base_planning(spec)),
        ("Exh.Base_Informar",  lambda: exh_base_informar(spec)),
        ("DyP.Ventas",         lambda: dyp_ventas(spec)),
        ("DyP.Impactos",       lambda: dyp_impactos(spec)),
        ("DyP.Rutero",         lambda: dyp_rutero(spec)),
    ]
    lineas = [f"Periodo: {spec.etiqueta}  ({spec})", "─" * 70]
    for nombre, fn in resolvers:
        try:
            ruta = fn()
            lineas.append(f"  ✓ {nombre:<22} → {ruta.name}")
        except Exception as e:
            lineas.append(f"  ✗ {nombre:<22} → {type(e).__name__}: {e}")
    return "\n".join(lineas)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Diagnóstico de resolución de archivos por periodo",
    )
    cli_add_periodo_args(parser)
    args = parser.parse_args()
    spec = periodo_de_args(args)
    print(diagnostico(spec))
