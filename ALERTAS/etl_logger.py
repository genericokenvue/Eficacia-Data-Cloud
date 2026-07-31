"""
etl_logger.py
─────────────
Sistema de logging centralizado para el conjunto de ETLs de Eficacia.

Responsabilidades
─────────────────
1. Emitir mensajes con nivel (INFO / WARNING / ERROR / CRITICAL)
   formateados con timestamp, ETL de origen y contexto.
2. Escribir simultáneamente a terminal (con color ANSI) y a un
   archivo de log en disco (sin color, legible como texto plano).
3. Ser completamente thread-safe: múltiples ETLs en paralelo
   escriben al mismo log sin entrelazar líneas ni perder eventos.
4. Acumular todos los eventos no-INFO en una lista interna para
   que el orquestador pueda incluirlos en el reporte final.

Uso desde cualquier módulo
──────────────────────────
    from etl_logger import get_logger

    log = get_logger("cif")          # nombre del ETL o módulo
    log.info("Paso 3 iniciado")
    log.warning("Sin motivos de inasistencia — paso 4 vacío")
    log.error("Archivo Involves no encontrado: {ruta}")
    log.critical("No se pudo cargar ningún PT")

Inicialización (solo desde run_all.py, una vez)
───────────────────────────────────────────────
    from etl_logger import inicializar_logging
    inicializar_logging(ruta_log="ETLS/logs/ejecucion_20260324.log")
"""

import logging
import threading
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

# ─────────────────────────────────────────────────────────────────────────────
# COLORES ANSI PARA TERMINAL
# ─────────────────────────────────────────────────────────────────────────────
_RESET  = "\033[0m"
_BOLD   = "\033[1m"
_CYAN   = "\033[36m"
_GREEN  = "\033[32m"
_YELLOW = "\033[33m"
_RED    = "\033[31m"
_BRED   = "\033[1;31m"   # bold red para CRITICAL
_GREY   = "\033[90m"

# Detectar si la terminal soporta colores (Windows cmd no los soporta por defecto)
_SUPPORTS_COLOR = (
    hasattr(sys.stdout, "isatty") and sys.stdout.isatty()
    and os.environ.get("NO_COLOR") is None
    and sys.platform != "win32"
    # En Windows, activar colores ANSI requiere `os.system('color')` o Windows 10+
)

_NIVEL_COLORES = {
    "DEBUG"   : _GREY,
    "INFO"    : _GREEN,
    "WARNING" : _YELLOW,
    "ERROR"   : _RED,
    "CRITICAL": _BRED,
}

_NIVEL_ICONOS = {
    "DEBUG"   : "·",
    "INFO"    : "✓",
    "WARNING" : "⚠",
    "ERROR"   : "✗",
    "CRITICAL": "☠",
}


# ─────────────────────────────────────────────────────────────────────────────
# ESTADO GLOBAL (singleton por proceso)
# ─────────────────────────────────────────────────────────────────────────────
_lock_global      = threading.Lock()
_eventos_no_info  = []          # lista de dicts con todos los WARNING/ERROR/CRITICAL
_inicializado     = False
_ruta_log_actual  = None


# ─────────────────────────────────────────────────────────────────────────────
# FORMATTER CON COLOR PARA TERMINAL
# ─────────────────────────────────────────────────────────────────────────────
class _ColorFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        nivel    = record.levelname
        icono    = _NIVEL_ICONOS.get(nivel, " ")
        ts       = datetime.fromtimestamp(record.created).strftime("%H:%M:%S")
        etl_tag  = f"[{record.name}]"

        if _SUPPORTS_COLOR:
            color   = _NIVEL_COLORES.get(nivel, _RESET)
            ts_str  = f"{_GREY}{ts}{_RESET}"
            tag_str = f"{_CYAN}{_BOLD}{etl_tag:<18}{_RESET}"
            lvl_str = f"{color}{_BOLD}{icono} {nivel:<8}{_RESET}"
            msg_str = f"{color}{record.getMessage()}{_RESET}"
        else:
            ts_str  = ts
            tag_str = f"{etl_tag:<18}"
            lvl_str = f"{icono} {nivel:<8}"
            msg_str = record.getMessage()

        return f"{ts_str}  {tag_str}  {lvl_str}  {msg_str}"


# ─────────────────────────────────────────────────────────────────────────────
# FORMATTER PLANO PARA ARCHIVO (sin secuencias ANSI)
# ─────────────────────────────────────────────────────────────────────────────
class _FileFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        nivel    = record.levelname
        icono    = _NIVEL_ICONOS.get(nivel, " ")
        ts       = datetime.fromtimestamp(record.created).strftime("%Y-%m-%d %H:%M:%S")
        etl_tag  = f"[{record.name}]"
        return f"{ts}  {etl_tag:<18}  {icono} {nivel:<8}  {record.getMessage()}"


# ─────────────────────────────────────────────────────────────────────────────
# HANDLER QUE ACUMULA EVENTOS NO-INFO
# ─────────────────────────────────────────────────────────────────────────────
class _AcumuladorHandler(logging.Handler):
    """
    Captura todos los registros con nivel >= WARNING y los almacena
    en la lista global _eventos_no_info para el reporte final.
    Thread-safe gracias al lock global.
    """
    def emit(self, record: logging.LogRecord) -> None:
        if record.levelno >= logging.WARNING:
            with _lock_global:
                _eventos_no_info.append({
                    "ts"      : datetime.fromtimestamp(record.created).strftime("%H:%M:%S"),
                    "etl"     : record.name,
                    "nivel"   : record.levelname,
                    "mensaje" : record.getMessage(),
                    "exc"     : self.formatException(record.exc_info) if record.exc_info else None,
                })


# ─────────────────────────────────────────────────────────────────────────────
# API PÚBLICA
# ─────────────────────────────────────────────────────────────────────────────

def inicializar_logging(
    ruta_log: Optional[str] = None,
    nivel_consola: str = "INFO",
    nivel_archivo: str = "DEBUG",
) -> None:
    """
    Configura el sistema de logging. Debe llamarse UNA SOLA VEZ
    desde run_all.py antes de lanzar cualquier ETL.

    Parámetros
    ──────────
    ruta_log       : Ruta del archivo .log. Si es None, se crea
                     automáticamente en ETLS/logs/ con timestamp.
    nivel_consola  : Nivel mínimo para mostrar en terminal (default INFO).
    nivel_archivo  : Nivel mínimo para escribir en archivo (default DEBUG).
    """
    global _inicializado, _ruta_log_actual, _eventos_no_info

    # Asegurar que stdout pueda escribir los íconos unicode (✓ ⚠ ✗)
    # en consolas Windows (cp1252) sin reventar el StreamHandler.
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass  # Stream no es un TextIOWrapper estándar — seguir adelante.

    with _lock_global:
        # Si ya estaba inicializado, reconfiguramos: limpiamos handlers
        # y volvemos a montar todo. Esto permite que un caller "tardío"
        # (típicamente run_all.py con --log-dir) imponga su ruta sobre
        # una auto-init previa hecha desde get_logger.
        if _inicializado:
            root_existing = logging.getLogger("etl")
            for h in list(root_existing.handlers):
                try:
                    h.close()
                except Exception:
                    pass
                root_existing.removeHandler(h)

        _eventos_no_info = []   # reset para ejecuciones sucesivas en tests

        # ── Determinar ruta del archivo de log ────────────────────────────
        if ruta_log is None:
            ts_nombre = datetime.now().strftime("%Y%m%d_%H%M%S")
            # Intenta crear la carpeta logs/ junto al script
            dir_log = Path(__file__).parent / "logs"
            dir_log.mkdir(exist_ok=True)
            ruta_log = str(dir_log / f"ejecucion_{ts_nombre}.log")

        _ruta_log_actual = ruta_log
        os.makedirs(os.path.dirname(os.path.abspath(ruta_log)), exist_ok=True)

        # ── Logger raíz de la familia "etl.*" ────────────────────────────
        root = logging.getLogger("etl")
        root.setLevel(logging.DEBUG)  # el nivel más bajo; los handlers filtran
        root.handlers.clear()

        # Handler 1: terminal con color
        h_consola = logging.StreamHandler(sys.stdout)
        h_consola.setLevel(getattr(logging, nivel_consola.upper()))
        h_consola.setFormatter(_ColorFormatter())

        # Handler 2: archivo plano
        h_archivo = logging.FileHandler(ruta_log, encoding="utf-8")
        h_archivo.setLevel(getattr(logging, nivel_archivo.upper()))
        h_archivo.setFormatter(_FileFormatter())

        # Handler 3: acumulador en memoria (para reporte final)
        h_acum = _AcumuladorHandler()
        h_acum.setLevel(logging.WARNING)

        root.addHandler(h_consola)
        root.addHandler(h_archivo)
        root.addHandler(h_acum)

        _inicializado = True

    logging.getLogger("etl.sistema").info(
        f"Logging inicializado — archivo: {ruta_log}"
    )


def get_logger(nombre_etl: str) -> logging.Logger:
    """
    Retorna un logger con jerarquía 'etl.<nombre_etl>'.
    Si el sistema no fue inicializado aún, hace una inicialización
    mínima automática (útil para ejecución standalone de un ETL).
    """
    if not _inicializado:
        inicializar_logging()   # inicialización mínima con defaults
    return logging.getLogger(f"etl.{nombre_etl}")


def obtener_eventos_acumulados() -> list[dict]:
    """
    Retorna una copia de la lista de eventos WARNING/ERROR/CRITICAL
    acumulados durante la ejecución. Usado por el reporte final.
    """
    with _lock_global:
        return list(_eventos_no_info)


def ruta_log_actual() -> Optional[str]:
    """Retorna la ruta del archivo de log activo."""
    return _ruta_log_actual


def separador(logger: logging.Logger, char: str = "─", ancho: int = 56) -> None:
    """Imprime una línea separadora como INFO (útil para secciones)."""
    logger.info(char * ancho)
