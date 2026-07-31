"""
config_loader.py
────────────────
Lectura del archivo ALERTAS/config.env (formato KEY=VALUE).

Importar `cargar_config()` desde cualquier script de la Fase 2 al inicio
para que las claves del archivo queden disponibles en os.environ. Las
variables de entorno reales tienen precedencia (permiten override en
CI/tests). Maneja comillas, espacios y los marcadores < > que se suelen
dejar accidentalmente al pegar tokens.
"""

import os
from pathlib import Path

# Preferimos el directorio donde vive este archivo (dev = sibling de SCRIPTS).
# Si no existe (ej. ejecución desde otra ruta), caemos al directorio de PROD.
_DEV_DIR  = Path(__file__).resolve().parent
_PROD_DIR = Path(r"C:\1\OneDrive - Eficacia\Escritorio\ETLS\ALERTAS")
DIR_ALERTAS = _DEV_DIR if _DEV_DIR.is_dir() else _PROD_DIR
RUTA_ENV    = DIR_ALERTAS / "config.env"


def cargar_config() -> dict:
    """
    Vuelca el contenido de config.env a os.environ y devuelve el dict
    cargado. Idempotente: usa os.environ.setdefault para no pisar
    variables ya exportadas externamente.
    """
    cargado: dict[str, str] = {}
    if not RUTA_ENV.exists():
        return cargado
    for linea in RUTA_ENV.read_text(encoding="utf-8").splitlines():
        linea = linea.strip()
        if not linea or linea.startswith("#") or "=" not in linea:
            continue
        clave, _, valor = linea.partition("=")
        clave = clave.strip()
        valor = valor.strip().strip('"').strip("'")
        if valor.startswith("<") and valor.endswith(">"):
            valor = valor[1:-1].strip()
        cargado[clave] = valor
        os.environ.setdefault(clave, valor)
    return cargado
