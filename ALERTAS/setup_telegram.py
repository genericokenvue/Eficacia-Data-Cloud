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

import os
import csv
import json
import requests
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

from config_loader import cargar_config


DIR_ALERTAS = Path(r"C:\1\OneDrive - Eficacia\Escritorio\ETLS\ALERTAS")
CSV_SALIDA = DIR_ALERTAS / "telegram_chat_ids.csv"

ZONA_HORARIA = ZoneInfo("America/Bogota")


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


def _leer_csv_existente(path_csv: Path) -> dict:
    """
    Lee el CSV existente y devuelve un diccionario indexado por chat_id.
    Esto permite mantener un histórico acumulado.
    """
    registros = {}

    if not path_csv.exists():
        return registros

    with path_csv.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)

        for row in reader:
            chat_id = str(row.get("chat_id", "")).strip()
            if chat_id:
                registros[chat_id] = row

    return registros


def _guardar_csv(path_csv: Path, registros: dict) -> None:
    """
    Guarda el histórico completo en CSV.
    """
    path_csv.parent.mkdir(parents=True, exist_ok=True)

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

    with path_csv.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=campos)
        writer.writeheader()
        writer.writerows(registros_ordenados)


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
