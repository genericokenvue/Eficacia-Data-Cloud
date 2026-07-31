"""
alertas_logger.py
─────────────────
Reexporta el logger thread-safe definido en SCRIPTS/etl_logger.py para que
los scripts de Fase 2 (calcular_cumplimientos, alertas_telegram,
alertas_email, run_alertas) compartan formato, archivo de log y semáforos
con el ETL pipeline.

Uso
───
    from alertas_logger import get_logger, inicializar_logging
    log = get_logger("calcular_cumplimientos")
    log.info(...)

Dónde se guarda el log
──────────────────────
Por defecto, ALERTAS/logs/alertas_<timestamp>.log. Se puede sobreescribir
pasando ruta_log a inicializar_logging() desde run_alertas.py.
"""

import sys
from pathlib import Path
from datetime import datetime

# SCRIPTS/ es hermana de ALERTAS/. Preferimos el sibling del archivo actual
# (dev). Si no existe, caemos a la ruta de producción. El orden importa
# cuando ambas existen (máquinas con OneDrive sincronizado y checkout local).
_DEV_SCRIPTS  = Path(__file__).resolve().parent.parent / "SCRIPTS"
_PROD_SCRIPTS = Path(r"C:\1\OneDrive - Eficacia\Escritorio\ETLS\SCRIPTS")
_CANDIDATAS_SCRIPTS = [_DEV_SCRIPTS, _PROD_SCRIPTS]
_SCRIPTS_DIR = next((p for p in _CANDIDATAS_SCRIPTS if p.is_dir()), None)
if _SCRIPTS_DIR is None:
    raise RuntimeError(
        "No se encontró el directorio SCRIPTS/ (con etl_logger.py). "
        f"Probadas: {[str(p) for p in _CANDIDATAS_SCRIPTS]}"
    )
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

from etl_logger import (  # noqa: E402  (import tras path manipulation)
    inicializar_logging as _inicializar_base,
    get_logger,
    obtener_eventos_acumulados,
    ruta_log_actual,
    separador,
)
import paths  # noqa: E402  — usa la misma resolución dev/prod que el resto

_DIR_ALERTAS = paths.ALERTAS_DIR
_DIR_LOGS    = paths.ALERTAS_LOGS


def inicializar_logging(ruta_log: str | None = None) -> str:
    """
    Inicializa logging para Fase 2. Si no se da ruta, crea
    ALERTAS/logs/alertas_<timestamp>.log. Devuelve la ruta usada.
    Es seguro llamar varias veces (reconfigura handlers).
    """
    if ruta_log is None:
        _DIR_LOGS.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_log = str(_DIR_LOGS / f"alertas_{ts}.log")
    _inicializar_base(ruta_log=ruta_log)
    return ruta_log


__all__ = [
    "inicializar_logging",
    "get_logger",
    "obtener_eventos_acumulados",
    "ruta_log_actual",
    "separador",
]
