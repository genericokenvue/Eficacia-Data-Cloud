"""
setup_telegram.py
─────────────────
Utilidad para recopilar y mantener un histórico local de chat_id de Telegram.

El script:
1. Consulta getUpdates de la Bot API.
2. Extrae chat_id desde distintos tipos de actualización:
   - message
   - edited_message
   - channel_post
   - edited_channel_post
   - callback_query
   - my_chat_member
   - chat_member
3. Deduplica los registros por chat_id.
4. Genera/actualiza un CSV con el histórico acumulado.

Uso:
    cd ALERTAS
    python setup_telegram.py

Salida:
    ALERTAS/telegram_chat_ids.csv
"""

import io
import os
import csv
import json
import sys
import urllib.parse
import requests
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

from config_loader import cargar_config

# ALERTAS/ vive DENTRO de SCRIPTS/ (SCRIPTS/ALERTAS/setup_telegram.py), no es
# su hermana — mismo patrón de resolución que el resto del pipeline.
_SCRIPTS_DIR = Path(__file__).resolve().parent.parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))
import paths

DIR_ALERTAS = paths.ALERTAS_DIR  # local: solo se usa para logs/temporales

# El CSV histórico de chat_ids ya NO vive en disco local — antes se perdía
# entre corridas de GitHub Actions (cada job arranca con disco vacío). Ahora
# vive en SharePoint, junto a la maestra de supervisores.
RUTA_CARPETA_ALERTAS_CLOUD = f"{paths._SALIDAS_ROOT}/ALERTAS"
CSV_SALIDA    = f"{RUTA_CARPETA_ALERTAS_CLOUD}/telegram_chat_ids.csv"      # ruta SharePoint
RUTA_OFFSET_CLOUD = f"{RUTA_CARPETA_ALERTAS_CLOUD}/telegram_bot_offset.txt"

ZONA_HORARIA = ZoneInfo("America/Bogota")


# ─────────────────────────────────────────────────────────────────────────────
# CONEXIÓN A LA NUBE (SharePoint vía Microsoft Graph API)
# ─────────────────────────────────────────────────────────────────────────────
_DRIVE_ID_CACHE: dict = {}


def _obtener_default_drive_id(token: str) -> str:
    if "drive_id" in _DRIVE_ID_CACHE:
        return _DRIVE_ID_CACHE["drive_id"]
    headers = {"Authorization": f"Bearer {token}"}
    nombre_sitio = getattr(paths, "SHAREPOINT_SITE_NAME", "eficacia")
    res_site = requests.get(
        f"https://graph.microsoft.com/v1.0/sites?search={urllib.parse.quote(nombre_sitio)}",
        headers=headers,
    )
    if res_site.status_code == 200:
        sites = res_site.json().get("value", [])
        if sites:
            site = next(
                (s for s in sites
                 if nombre_sitio.lower() in (s.get("name", "").lower(), s.get("displayName", "").lower())),
                sites[0],
            )
            res_drive = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site.get('id')}/drive", headers=headers)
            if res_drive.status_code == 200:
                _DRIVE_ID_CACHE["drive_id"] = res_drive.json().get("id")
                return _DRIVE_ID_CACHE["drive_id"]
    res_root = requests.get("https://graph.microsoft.com/v1.0/sites/root", headers=headers)
    if res_root.status_code == 200:
        site_id = res_root.json().get("id")
        res_drive = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", headers=headers)
        if res_drive.status_code == 200:
            _DRIVE_ID_CACHE["drive_id"] = res_drive.json().get("id")
            return _DRIVE_ID_CACHE["drive_id"]
    raise Exception(f"No se pudo obtener el Drive ID del sitio SharePoint '{nombre_sitio}': {res_site.text}")


def _limpiar_ruta_graph(ruta_sharepoint: str) -> str:
    ruta_sharepoint = ruta_sharepoint.replace("\\", "/")
    for prefijo in ["Documentos compartidos/", "Shared Documents/"]:
        if ruta_sharepoint.startswith(prefijo):
            ruta_sharepoint = ruta_sharepoint.replace(prefijo, "", 1)
    return ruta_sharepoint.lstrip("/")


def _descargar_bytes_cloud(ruta_sharepoint: str, descripcion: str) -> bytes:
    token = paths.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}"}
    drive_id = _obtener_default_drive_id(token)
    ruta_limpia = _limpiar_ruta_graph(ruta_sharepoint)
    response = None
    for r in (ruta_limpia, f"Documentos compartidos/{ruta_limpia}"):
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{urllib.parse.quote(r)}:/content"
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.content
    status = response.status_code if response is not None else "N/A"
    raise FileNotFoundError(f"No se pudo descargar {descripcion} desde SharePoint ({ruta_sharepoint}): {status}")


def _subir_bytes_cloud(contenido: bytes, ruta_sharepoint: str, content_type: str) -> None:
    token = paths.obtener_token_azure()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": content_type}
    drive_id = _obtener_default_drive_id(token)
    ruta_limpia = _limpiar_ruta_graph(ruta_sharepoint)
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{urllib.parse.quote(ruta_limpia)}:/content"
    response = requests.put(url, headers=headers, data=contenido)
    if response.status_code not in (200, 201):
        raise Exception(f"Error al subir '{ruta_sharepoint}' a SharePoint: {response.status_code} - {response.text}")


def _cargar_token() -> str:
    cargar_config()
    token = os.environ.get("TELEGRAM_TOKEN", "").strip()

    if not token or token.startswith("<") or "REEMPLAZAR" in token.upper():
        raise ValueError(
            "TELEGRAM_TOKEN no configurado.\n"
            "Edita ALERTAS/config.env y pon: TELEGRAM_TOKEN=NN:AA... sin < >"
        )

    return token


def _ahora_colombia() -> str:
    return datetime.now(ZONA_HORARIA).strftime("%Y-%m-%d %H:%M:%S")


def _leer_csv_existente(ruta_cloud: str) -> dict:
    """
    Lee el CSV histórico desde SharePoint (ruta_cloud) y devuelve un
    diccionario indexado por chat_id. Antes leía de disco local; el CSV
    ya no vive ahí (se perdía entre corridas de GitHub Actions).
    """
    registros = {}
    try:
        contenido = _descargar_bytes_cloud(ruta_cloud, "telegram_chat_ids.csv")
    except Exception:
        return registros  # no existe aún en SharePoint — primera corrida

    texto = contenido.decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(texto))
    for row in reader:
        chat_id = str(row.get("chat_id", "")).strip()
        if chat_id:
            registros[chat_id] = row

    return registros


def _guardar_csv(ruta_cloud: str, registros: dict) -> None:
    """
    Guarda el histórico completo como CSV en SharePoint (ruta_cloud).
    """
    campos = [
        "chat_id",
        "chat_type",
        "chat_title",
        "telegram_user_id",
        "nombre_telegram",
        "username",
        "is_bot",
        "language_code",
        "ultimo_update_id",
        "tipo_update",
        "fecha_ultimo_registro",
    ]

    registros_ordenados = sorted(
        registros.values(),
        key=lambda x: str(x.get("nombre_telegram", "")).lower()
    )

    buffer = io.StringIO()
    writer = csv.DictWriter(buffer, fieldnames=campos)
    writer.writeheader()
    writer.writerows(registros_ordenados)
    contenido = buffer.getvalue().encode("utf-8-sig")

    _subir_bytes_cloud(contenido, ruta_cloud, "text/csv; charset=utf-8-sig")


def _extraer_desde_message(update: dict, tipo_update: str) -> dict | None:
    """
    Extrae información desde updates que contienen un objeto Message.
    """
    msg = update.get(tipo_update)
    if not msg:
        return None

    chat = msg.get("chat", {}) or {}
    sender = msg.get("from", {}) or {}

    chat_id = chat.get("id")
    if chat_id is None:
        return None

    nombre = (
        f"{sender.get('first_name', '')} {sender.get('last_name', '')}"
    ).strip()

    return {
        "chat_id": str(chat_id),
        "chat_type": chat.get("type", ""),
        "chat_title": chat.get("title", ""),
        "telegram_user_id": str(sender.get("id", "")),
        "nombre_telegram": nombre,
        "username": sender.get("username", ""),
        "is_bot": str(sender.get("is_bot", "")),
        "language_code": sender.get("language_code", ""),
        "tipo_update": tipo_update,
    }


def _extraer_desde_callback_query(update: dict) -> dict | None:
    """
    Extrae información desde callback_query, cuando el usuario interactúa
    con botones inline del bot.
    """
    cq = update.get("callback_query")
    if not cq:
        return None

    msg = cq.get("message", {}) or {}
    chat = msg.get("chat", {}) or {}
    sender = cq.get("from", {}) or {}

    chat_id = chat.get("id")
    if chat_id is None:
        return None

    nombre = (
        f"{sender.get('first_name', '')} {sender.get('last_name', '')}"
    ).strip()

    return {
        "chat_id": str(chat_id),
        "chat_type": chat.get("type", ""),
        "chat_title": chat.get("title", ""),
        "telegram_user_id": str(sender.get("id", "")),
        "nombre_telegram": nombre,
        "username": sender.get("username", ""),
        "is_bot": str(sender.get("is_bot", "")),
        "language_code": sender.get("language_code", ""),
        "tipo_update": "callback_query",
    }


def _extraer_desde_chat_member(update: dict, tipo_update: str) -> dict | None:
    """
    Extrae información desde my_chat_member o chat_member.
    Útil para detectar cuando el bot fue agregado, bloqueado o actualizado
    dentro de un chat.
    """
    cm = update.get(tipo_update)
    if not cm:
        return None

    chat = cm.get("chat", {}) or {}
    sender = cm.get("from", {}) or {}

    chat_id = chat.get("id")
    if chat_id is None:
        return None

    nombre = (
        f"{sender.get('first_name', '')} {sender.get('last_name', '')}"
    ).strip()

    return {
        "chat_id": str(chat_id),
        "chat_type": chat.get("type", ""),
        "chat_title": chat.get("title", ""),
        "telegram_user_id": str(sender.get("id", "")),
        "nombre_telegram": nombre,
        "username": sender.get("username", ""),
        "is_bot": str(sender.get("is_bot", "")),
        "language_code": sender.get("language_code", ""),
        "tipo_update": tipo_update,
    }


def _extraer_registro(update: dict) -> dict | None:
    """
    Intenta extraer un chat_id desde diferentes tipos de actualización.
    """
    update_id = update.get("update_id", "")

    tipos_message = [
        "message",
        "edited_message",
        "channel_post",
        "edited_channel_post",
        "business_message",
        "edited_business_message",
    ]

    for tipo in tipos_message:
        registro = _extraer_desde_message(update, tipo)
        if registro:
            registro["ultimo_update_id"] = str(update_id)
            registro["fecha_ultimo_registro"] = _ahora_colombia()
            return registro

    registro = _extraer_desde_callback_query(update)
    if registro:
        registro["ultimo_update_id"] = str(update_id)
        registro["fecha_ultimo_registro"] = _ahora_colombia()
        return registro

    for tipo in ["my_chat_member", "chat_member"]:
        registro = _extraer_desde_chat_member(update, tipo)
        if registro:
            registro["ultimo_update_id"] = str(update_id)
            registro["fecha_ultimo_registro"] = _ahora_colombia()
            return registro

    return None


def obtener_updates(token: str) -> list:
    """
    Consulta getUpdates.

    Nota:
    - No se usa offset aquí para no marcar como confirmadas las actualizaciones
      de manera agresiva durante la fase de configuración.
    - allowed_updates=[] indica que se solicitan los tipos permitidos por defecto.
    """
    url = f"https://api.telegram.org/bot{token}/getUpdates"

    params = {
        "timeout": 10,
        "limit": 100,
        "allowed_updates": json.dumps([]),
    }

    resp = requests.get(url, params=params, timeout=20)
    resp.raise_for_status()

    data = resp.json()

    if not data.get("ok"):
        raise RuntimeError(f"Error de la API de Telegram: {data}")

    return data.get("result", [])


def recopilar_chat_ids() -> None:
    token = _cargar_token()

    registros_historicos = _leer_csv_existente(CSV_SALIDA)
    updates = obtener_updates(token)

    nuevos = 0
    actualizados = 0

    for update in updates:
        registro = _extraer_registro(update)

        if not registro:
            continue

        chat_id = registro["chat_id"]

        if chat_id in registros_historicos:
            registros_historicos[chat_id].update(registro)
            actualizados += 1
        else:
            registros_historicos[chat_id] = registro
            nuevos += 1

    _guardar_csv(CSV_SALIDA, registros_historicos)

    print("\n=== Recopilación de chat_id finalizada ===\n")
    print(f"Archivo generado/actualizado: {CSV_SALIDA}")
    print(f"Registros nuevos          : {nuevos}")
    print(f"Registros actualizados    : {actualizados}")
    print(f"Total histórico acumulado : {len(registros_historicos)}")

    if not updates:
        print("\nNo se encontraron actualizaciones disponibles en getUpdates.")
        print("Pide a cada supervisor que escriba /start al bot y vuelve a ejecutar el script.")

    if registros_historicos:
        print("\n=== Chat IDs acumulados ===\n")

        for row in registros_historicos.values():
            username = row.get("username", "")
            username_txt = f"@{username}" if username else ""

            print(f"Nombre Telegram : {row.get('nombre_telegram', '')}")
            print(f"Username        : {username_txt}")
            print(f"Tipo chat       : {row.get('chat_type', '')}")
            print(f"chat_id         : {row.get('chat_id', '')}")
            print("-" * 45)


if __name__ == "__main__":
    recopilar_chat_ids()