"""
envio_errores.py
────────────────
Le manda a cada supervisor el link al archivo con SUS errores para que los
corrija en Excel Online.

VA SEPARADO DEL CÁLCULO A PROPÓSITO
───────────────────────────────────
`etl_precios_errores.py` detecta y escribe el archivo; este script solo lee
ese archivo y envía. Están separados porque tienen ritmos distintos: el
cálculo puede correr seguido sin molestar a nadie, mientras que el aviso
conviene mandarlo poco y en un horario razonable. Además, así se puede
reenviar sin recalcular ni arriesgar las correcciones ya escritas.

PREPARADO PARA MÁS VALIDACIONES
───────────────────────────────
Hoy solo existe la de precios. Para agregar otra basta con sumarle una entrada
a VALIDACIONES: el resto (agrupar por supervisor, buscar correos, armar el
correo, enviar) ya es común.

A QUIÉN LE LLEGA
────────────────
Solo a quien TIENE errores. Mandarle un correo a alguien para decirle que no
tiene nada que corregir es la forma más rápida de que la gente deje de abrir
estos avisos.

La maestra admite varias filas con el mismo NOMBRE_SUPERVISOR: cada una es un
destinatario, así se le puede copiar a un acompañante sin tocar código.

USO
───
    python envio_errores.py --mes 8 --anio 2026 --prueba
    python envio_errores.py --mes 8 --anio 2026
    python envio_errores.py --mes 8 --anio 2026 --supervisor "YULY ALARCON"
    python envio_errores.py --mes 8 --anio 2026 --tipo precios
"""
from __future__ import annotations

import argparse
import base64
import io
import os
import sys
import urllib.parse

import msal
import pandas as pd
import requests
from dotenv import load_dotenv

import paths
import periodo_resolver as pr

load_dotenv()

TENANT_ID = os.environ.get("AZURE_TENANT_ID")
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET")
# .strip() no es opcional: el valor va dentro de la URL de Graph, y un espacio
# o salto de línea al final del secret se codifica como %0A y deja la URL
# inválida. El servidor responde con un HTML de IIS ("Bad Request - Invalid
# URL") en vez de un error de Graph, así que el mensaje no dice qué pasó.
CORREO_REMITENTE = os.environ.get("CORREO_REMITENTE", "").strip()

RUTA_MAESTRA = f"{paths._BASES_ROOT}/ALERTAS/MAESTRO_SUPERVISORES.xlsx"


# ─────────────────────────────────────────────────────────────────────────────
# QUÉ VALIDACIONES SE ENVÍAN
# ─────────────────────────────────────────────────────────────────────────────
# Para agregar una nueva: sumá una entrada acá. `archivo` recibe el periodo y
# devuelve el nombre del Excel; `columnas` dice qué mostrar en la tabla del
# correo (encabezado → columna del archivo).
VALIDACIONES = {
    "precios": {
        "titulo": "Precios para revisar",
        "carpeta": paths.RUTA_CARPETA_SALIDAS_PRECIOS,
        "archivo": lambda s: f"ERRORES_PRECIOS_{s.mes_str_upper}_{s.anio}.xlsx",
        "que_paso": ("Se detectaron precios que se alejan mucho del precio habitual "
                     "de ese mismo producto. Casi siempre es un dígito de más o de "
                     "menos al digitar."),
        "que_hacer_adjunto": ("escribí el precio real en <b>PRECIO_CORREGIDO</b>. "
                              "Si el precio estaba bien, dejalo vacío y anotá por "
                              "qué en <b>OBSERVACION_SUPERVISOR</b>."),
        "columnas": [
            ("Gestor", "Empleado"),
            ("Producto", "NOMBRE_PRODUCTO"),
            ("Capturado", "PRECIO_CAPTURADO"),
            ("Diagnóstico", "DIAGNOSTICO"),
        ],
        "moneda": {"PRECIO_CAPTURADO"},
        "prefijo_adjunto": "Errores_Precios",
    },
    "exhibiciones": {
        "titulo": "Exhibiciones para revisar",
        "carpeta": paths.RUTA_CARPETA_SALIDAS_EXHIB,
        "archivo": lambda s: f"ERRORES_EXHIBICIONES_{s.mes_str_upper}_{s.anio}.xlsx",
        "que_paso": ("Se registraron exhibiciones con una cantidad muy alta para un "
                     "solo punto de venta. Suele ser que se digitó el total del mes "
                     "en vez de lo de esa visita."),
        "que_hacer_adjunto": ("escribí la cantidad real en <b>CANTIDAD_CORREGIDA</b>. "
                              "Si la cantidad estaba bien, dejala vacía y anotá por "
                              "qué en <b>OBSERVACION_SUPERVISOR</b>."),
        "columnas": [
            ("Gestor", "Empleado"),
            ("PDV", "ID PDV"),
            ("Tipo", "Tipo Exhibición"),
            ("Marca", "Marca"),
            ("Cantidad", "CANTIDAD"),
            ("Diagnóstico", "DIAGNOSTICO"),
        ],
        "moneda": set(),
        "prefijo_adjunto": "Errores_Exhibiciones",
    },
}


# ─────────────────────────────────────────────────────────────────────────────
# SHAREPOINT
# ─────────────────────────────────────────────────────────────────────────────

def obtener_token_azure() -> str:
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res:
        raise RuntimeError(f"Azure rechazó la autenticación: {res.get('error_description')}")
    return res["access_token"]


def obtener_site_id(headers: dict) -> str:
    url = f"https://graph.microsoft.com/v1.0/sites/root:/sites/{paths.SHAREPOINT_SITE_NAME}"
    res = requests.get(url, headers=headers).json()
    if "id" not in res:
        raise RuntimeError(f"No se encontró el sitio '{paths.SHAREPOINT_SITE_NAME}'")
    return res["id"]


def leer_excel(headers: dict, site_id: str, ruta: str, desc: str,
               obligatorio: bool = True) -> pd.DataFrame | None:
    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/"
           f"{urllib.parse.quote(ruta)}:/content")
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        if obligatorio:
            raise FileNotFoundError(f"No se pudo leer {desc}: {ruta} ({r.status_code})")
        print(f"  ℹ️  No existe {desc} ({ruta}) — se salta.")
        return None
    print(f"  ✓ {desc} leído")
    return pd.read_excel(io.BytesIO(r.content))


def link_archivo(headers: dict, site_id: str, carpeta: str, nombre: str) -> str:
    """
    Link web del archivo. Abre en Excel Online, que es donde conviene que lo
    editen: permite que varios trabajen a la vez y no bloquea el archivo como
    sí lo hace el Excel de escritorio.
    """
    url = (f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/"
           f"{urllib.parse.quote(carpeta)}/{urllib.parse.quote(nombre)}")
    r = requests.get(url, headers=headers)
    return r.json().get("webUrl", "") if r.status_code == 200 else ""


def excel_del_supervisor(filas: pd.DataFrame) -> bytes:
    """
    Excel con SOLO las filas de ese supervisor, en una sola hoja.

    Se manda adjunto en vez de un link al archivo compartido para que cada
    supervisor reciba únicamente lo suyo y no vea los casos de sus pares.
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        filas.to_excel(w, sheet_name="Errores", index=False)
    buf.seek(0)
    return buf.getvalue()


def validar_remitente() -> None:
    """
    Revisa CORREO_REMITENTE ANTES de intentar el primer envío.

    Sin esto, un valor mal formado hace fallar cada destinatario por separado
    con el mismo error repetido, y el mensaje que devuelve Graph en ese caso
    ("Bad Request - Invalid URL", en HTML) no dice cuál es el problema real.
    """
    if not CORREO_REMITENTE:
        raise RuntimeError(
            "Falta la variable de entorno CORREO_REMITENTE (la casilla desde la "
            "que salen los correos vía Graph — requiere el permiso Mail.Send)."
        )
    if "@" not in CORREO_REMITENTE or any(c.isspace() for c in CORREO_REMITENTE):
        raise RuntimeError(
            f"CORREO_REMITENTE no parece una dirección válida: {CORREO_REMITENTE!r}. "
            f"Revisá el secret en GitHub — un espacio o salto de línea al final "
            f"alcanza para invalidar la URL de Graph."
        )


def enviar_correo(destinatario: str, asunto: str, cuerpo: str,
                  nombre_adjunto: str | None = None,
                  bytes_adjunto: bytes | None = None) -> None:
    token = paths.obtener_token_azure()
    url = (f"https://graph.microsoft.com/v1.0/users/"
           f"{urllib.parse.quote(CORREO_REMITENTE)}/sendMail")
    mensaje = {
        "subject": asunto,
        "body": {"contentType": "HTML", "content": cuerpo},
        "toRecipients": [{"emailAddress": {"address": destinatario}}],
    }
    if bytes_adjunto and nombre_adjunto:
        mensaje["attachments"] = [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": nombre_adjunto,
            "contentBytes": base64.b64encode(bytes_adjunto).decode(),
        }]
    r = requests.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"message": mensaje, "saveToSentItems": "true"},
    )
    if r.status_code not in (200, 202):
        raise RuntimeError(f"Graph sendMail devolvió {r.status_code}: {r.text}")


# ─────────────────────────────────────────────────────────────────────────────
# CORREO
# ─────────────────────────────────────────────────────────────────────────────

BORDE = "padding:6px 10px;border-bottom:1px solid #eee"
ALTA_HTML = '<b style="color:#c00">ALTA</b>'
MAX_FILAS_CORREO = 15


def cuerpo_html(nombre_sup: str, filas: pd.DataFrame, link: str,
                spec: pr.PeriodoSpec, cfg: dict) -> str:
    n = len(filas)
    altas = int((filas["SEVERIDAD"] == "ALTA").sum()) if "SEVERIDAD" in filas else 0
    moneda = cfg.get("moneda", set())

    encabezados = "".join(
        f"<th style='padding:7px 10px;text-align:left'>{tit}</th>"
        for tit, _ in cfg["columnas"])
    if "SEVERIDAD" in filas.columns:
        encabezados = ("<th style='padding:7px 10px;text-align:left'>Severidad</th>"
                       + encabezados)

    def fila_html(r) -> str:
        celdas = ""
        if "SEVERIDAD" in filas.columns:
            sev = ALTA_HTML if r["SEVERIDAD"] == "ALTA" else "Media"
            celdas += f"<td style='{BORDE}'>{sev}</td>"
        for _, col in cfg["columnas"]:
            v = r.get(col, "")
            if col in moneda:
                try:
                    celdas += f"<td style='{BORDE};text-align:right'>${float(v):,.0f}</td>"
                    continue
                except (TypeError, ValueError):
                    pass
            celdas += f"<td style='{BORDE}'>{v}</td>"
        return f"<tr>{celdas}</tr>"

    tabla = "".join(fila_html(r) for _, r in filas.head(MAX_FILAS_CORREO).iterrows())
    mas = (f"<p style='color:#666'>… y {n - MAX_FILAS_CORREO} más en el archivo.</p>"
           if n > MAX_FILAS_CORREO else "")
    urgentes = (f', <b style="color:#c00">{altas} muy marcados</b>' if altas else "")
    # Sin botón: el archivo va adjunto, no hay a dónde enlazar.
    boton = ""

    return f"""
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222">
  <h2 style="color:#1F4E79;margin-bottom:4px">{cfg['titulo']}
      — {spec.mes_str} {spec.anio}</h2>
  <p style="color:#666;margin-top:0">Equipo de <b>{nombre_sup}</b></p>

  <p>Hay <b>{n} caso(s)</b> para revisar{urgentes}. {cfg['que_paso']}</p>
  {boton}
  <p><b>Qué hacer:</b> abrí el archivo adjunto, {cfg['que_hacer_adjunto']}
     Cuando termines, respondé este correo con el archivo.</p>

  <h3 style="color:#1F4E79;margin-bottom:6px">Los de tu equipo</h3>
  <table style="border-collapse:collapse;font-size:13px">
    <tr style="background:#DCE6F1">{encabezados}</tr>
    {tabla}
  </table>
  {mas}

  <p style="color:#888;font-size:12px;margin-top:24px">
     El adjunto trae solo los casos de tu equipo.</p>
</div>"""


# ─────────────────────────────────────────────────────────────────────────────
# PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

def enviar_validacion(headers: dict, site_id: str, maestra: pd.DataFrame,
                      tipo: str, spec: pr.PeriodoSpec, prueba: bool,
                      solo: list[str] | None) -> tuple[int, int, int]:
    cfg = VALIDACIONES[tipo]
    nombre = cfg["archivo"](spec)
    carpeta = cfg["carpeta"]

    print(f"\n── {tipo.upper()} ─────────────────────────────────")
    err = leer_excel(headers, site_id, f"{carpeta}/{nombre}", nombre, obligatorio=False)
    if err is None or err.empty:
        print(f"  Sin errores que enviar.")
        return 0, 0, 0

    link = link_archivo(headers, site_id, carpeta, nombre)
    if not link:
        print("  ⚠️  No se pudo obtener el link; el correo va sin botón.")

    err["SUPERVISOR_LIDER"] = err["SUPERVISOR_LIDER"].astype(str).str.strip().str.upper()
    filtro = [s.strip().upper() for s in solo] if solo else None
    enviados = fallidos = sin_correo = 0

    for nombre_sup, filas in err.groupby("SUPERVISOR_LIDER"):
        if filtro and nombre_sup not in filtro:
            continue
        if nombre_sup == "(SIN CRUZAR)":
            print(f"  ⏭️  {len(filas)} fila(s) sin supervisor asignado — no se envían.")
            continue

        correos = [c for c in maestra.loc[
            maestra["NOMBRE_SUPERVISOR"] == nombre_sup, "CORREO"].tolist()
            if c not in ("", "nan", "NAN")]
        correos = list(dict.fromkeys(correos))
        if not correos:
            print(f"  ⚠️  {nombre_sup}: sin correo en la maestra — no se le envía.")
            sin_correo += 1
            continue

        altas = int((filas["SEVERIDAD"] == "ALTA").sum()) if "SEVERIDAD" in filas else 0
        asunto = (f"{cfg['titulo']} — {spec.mes_str} {spec.anio} — "
                  f"{len(filas)} caso(s)" + (f", {altas} urgentes" if altas else ""))
        cuerpo = cuerpo_html(nombre_sup, filas, link, spec, cfg)

        # Un archivo por supervisor, con solo sus filas. El nombre lleva el
        # suyo para que al responder el correo se sepa de quién viene.
        adjunto = excel_del_supervisor(filas)
        nombre_adjunto = (f"{cfg['prefijo_adjunto']}_{nombre_sup.replace(' ', '_')}"
                          f"_{spec.mes:02d}_{spec.anio}.xlsx")

        for correo in correos:
            if prueba:
                print(f"  [PRUEBA] {nombre_sup} → {correo} "
                      f"({len(filas)} casos, {altas} altas) — adjunto: {nombre_adjunto}")
                enviados += 1
                continue
            try:
                enviar_correo(correo, asunto, cuerpo, nombre_adjunto, adjunto)
                print(f"  ✓ {nombre_sup} → {correo} ({len(filas)} casos)")
                enviados += 1
            except Exception as e:
                print(f"  ❌ {nombre_sup} → {correo}: {e}")
                fallidos += 1

    return enviados, fallidos, sin_correo


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawTextHelpFormatter)
    ap.add_argument("--mes", type=int, required=True, choices=range(1, 13))
    ap.add_argument("--anio", type=int, required=True)
    ap.add_argument("--tipo", default="todas",
                    help=f"Qué enviar: todas | {' | '.join(VALIDACIONES)}")
    ap.add_argument("--prueba", action="store_true",
                    help="Muestra a quién le llegaría, sin enviar nada.")
    ap.add_argument("--supervisor", nargs="+", default=None,
                    help="Enviar solo a estos supervisores (nombre completo).")
    args = ap.parse_args()

    spec = pr.resolver(int(args.mes), int(args.anio))
    tipos = list(VALIDACIONES) if args.tipo == "todas" else [args.tipo]
    desconocidos = [t for t in tipos if t not in VALIDACIONES]
    if desconocidos:
        print(f"❌ Validación desconocida: {desconocidos}. "
              f"Disponibles: {list(VALIDACIONES)}")
        return 1

    print("\n" + "=" * 60)
    print(f"  ENVÍO DE ERRORES — {spec.etiqueta}")
    print(f"  Validaciones: {', '.join(tipos)}")
    if args.prueba:
        print("  MODO PRUEBA — no se envía ningún correo")
    print("=" * 60)

    # Se valida el remitente ANTES de leer nada: si está mal, mejor enterarse
    # ahora que después de armar 18 correos y ver el mismo error 18 veces.
    if not args.prueba:
        validar_remitente()
        print(f"  Remitente: {CORREO_REMITENTE}")

    headers = {"Authorization": f"Bearer {obtener_token_azure()}"}
    site_id = obtener_site_id(headers)

    print("\nLeyendo la maestra de supervisores:")
    maestra = leer_excel(headers, site_id, RUTA_MAESTRA, "maestra de supervisores")
    maestra.columns = maestra.columns.str.strip().str.upper()
    maestra["NOMBRE_SUPERVISOR"] = (
        maestra["NOMBRE_SUPERVISOR"].fillna("").astype(str).str.strip().str.upper())
    maestra["CORREO"] = maestra["CORREO"].fillna("").astype(str).str.strip()

    tot_env = tot_fall = tot_sin = 0
    for tipo in tipos:
        e, f, s = enviar_validacion(headers, site_id, maestra, tipo, spec,
                                    args.prueba, args.supervisor)
        tot_env += e
        tot_fall += f
        tot_sin += s

    print(f"\n  Enviados: {tot_env} | Fallidos: {tot_fall} | Sin correo: {tot_sin}")
    return 1 if tot_fall else 0


if __name__ == "__main__":
    sys.exit(main())
