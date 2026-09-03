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
        "hoja": "Precios",
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
        "hoja": "Exhibiciones",
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


def _encabezados(cfg: dict) -> str:
    return "".join(f"<th style='padding:7px 10px;text-align:left'>{tit}</th>"
                   for tit, _ in cfg["columnas"])


def _fila_html(r, cfg: dict) -> str:
    celdas = ""
    for _, col in cfg["columnas"]:
        v = r.get(col, "")
        if col in cfg.get("moneda", set()):
            try:
                celdas += f"<td style='{BORDE};text-align:right'>${float(v):,.0f}</td>"
                continue
            except (TypeError, ValueError):
                pass
        celdas += f"<td style='{BORDE}'>{v}</td>"
    return f"<tr>{celdas}</tr>"


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

ASUNTO_BASE = "CAPTURA DE ERRORES"


def cargar_errores(headers: dict, site_id: str, tipos: list[str],
                   spec: pr.PeriodoSpec) -> dict[str, pd.DataFrame]:
    """Lee el archivo de cada validación. Devuelve {tipo: DataFrame}."""
    datos: dict[str, pd.DataFrame] = {}
    for tipo in tipos:
        cfg = VALIDACIONES[tipo]
        nombre = cfg["archivo"](spec)
        err = leer_excel(headers, site_id, f"{cfg['carpeta']}/{nombre}", nombre,
                         obligatorio=False)
        if err is None or err.empty:
            print(f"  ℹ️  {tipo}: sin errores.")
            continue
        err["SUPERVISOR_LIDER"] = (
            err["SUPERVISOR_LIDER"].astype(str).str.strip().str.upper())
        datos[tipo] = err
        print(f"  ✓ {tipo}: {len(err)} caso(s)")
    return datos


def libro_del_supervisor(por_tipo: dict[str, pd.DataFrame]) -> bytes:
    """
    UN archivo con una hoja por validación, solo con las filas de ese
    supervisor. Antes iba un archivo por validación y llegaban dos correos;
    con una hoja por tema alcanza uno solo.
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for tipo, filas in por_tipo.items():
            filas.to_excel(w, sheet_name=VALIDACIONES[tipo]["hoja"][:31], index=False)
    buf.seek(0)
    return buf.getvalue()


def cuerpo_unico(nombre_sup: str, por_tipo: dict[str, pd.DataFrame],
                 spec: pr.PeriodoSpec) -> str:
    """Un cuerpo con una sección por validación."""
    total = sum(len(f) for f in por_tipo.values())
    secciones = []
    for tipo, filas in por_tipo.items():
        cfg = VALIDACIONES[tipo]
        secciones.append(f"""
  <h3 style="color:#1F4E79;margin:26px 0 4px">{cfg['titulo']}
      — {len(filas)} caso(s)</h3>
  <p style="margin-top:0">{cfg['que_paso']}</p>
  <p style="color:#444">En la hoja <b>{cfg['hoja']}</b>, {cfg['que_hacer_adjunto']}</p>
  <table style="border-collapse:collapse;font-size:13px">
    <tr style="background:#DCE6F1">{_encabezados(cfg)}</tr>
    {''.join(_fila_html(r, cfg) for _, r in filas.head(MAX_FILAS_CORREO).iterrows())}
  </table>
  {f"<p style='color:#666'>… y {len(filas) - MAX_FILAS_CORREO} más en el archivo.</p>"
   if len(filas) > MAX_FILAS_CORREO else ""}""")

    return f"""
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222">
  <h2 style="color:#1F4E79;margin-bottom:4px">{ASUNTO_BASE}
      — {spec.mes_str} {spec.anio}</h2>
  <p style="color:#666;margin-top:0">Equipo de <b>{nombre_sup}</b></p>

  <p>Hay <b>{total} caso(s)</b> para revisar, en el archivo adjunto.
     Cada tema está en su propia hoja.</p>

  <p><b>Cuando termines, respondé este correo con el archivo.</b></p>
  {''.join(secciones)}

  <p style="color:#888;font-size:12px;margin-top:26px">
     El adjunto trae solo los casos de tu equipo. La columna <b>ID_ERROR</b>
     identifica cada caso: si preguntás por uno, mencioná ese código.</p>
</div>"""


def enviar_todo(headers: dict, site_id: str, maestra: pd.DataFrame,
                por_validacion: dict[str, pd.DataFrame], spec: pr.PeriodoSpec,
                prueba: bool, solo: list[str] | None) -> tuple[int, int, int]:
    """Un correo por supervisor, con todas sus validaciones juntas."""
    # Universo de supervisores con al menos un caso en alguna validación.
    sups = sorted({s for err in por_validacion.values()
                   for s in err["SUPERVISOR_LIDER"].unique()})
    filtro = [s.strip().upper() for s in solo] if solo else None
    enviados = fallidos = sin_correo = 0

    for nombre_sup in sups:
        if filtro and nombre_sup not in filtro:
            continue
        if nombre_sup == "(SIN CRUZAR)":
            n = sum(int((e["SUPERVISOR_LIDER"] == "(SIN CRUZAR)").sum())
                    for e in por_validacion.values())
            print(f"  ⏭️  {n} fila(s) sin supervisor asignado — no se envían.")
            continue

        por_tipo = {}
        for tipo, err in por_validacion.items():
            filas = err[err["SUPERVISOR_LIDER"] == nombre_sup]
            if not filas.empty:
                por_tipo[tipo] = filas
        if not por_tipo:
            continue

        correos = [c for c in maestra.loc[
            maestra["NOMBRE_SUPERVISOR"] == nombre_sup, "CORREO"].tolist()
            if c not in ("", "nan", "NAN")]
        correos = list(dict.fromkeys(correos))
        if not correos:
            print(f"  ⚠️  {nombre_sup}: sin correo en la maestra — no se le envía.")
            sin_correo += 1
            continue

        total = sum(len(f) for f in por_tipo.values())
        detalle = ", ".join(f"{len(f)} de {t}" for t, f in por_tipo.items())
        asunto = f"{ASUNTO_BASE} — {spec.mes_str} {spec.anio} — {total} caso(s)"
        cuerpo = cuerpo_unico(nombre_sup, por_tipo, spec)
        adjunto = libro_del_supervisor(por_tipo)
        nombre_adjunto = (f"Errores_{nombre_sup.replace(' ', '_')}"
                          f"_{spec.mes:02d}_{spec.anio}.xlsx")

        for correo in correos:
            if prueba:
                print(f"  [PRUEBA] {nombre_sup} → {correo} ({detalle}) "
                      f"— adjunto: {nombre_adjunto}")
                enviados += 1
                continue
            try:
                enviar_correo(correo, asunto, cuerpo, nombre_adjunto, adjunto)
                print(f"  ✓ {nombre_sup} → {correo} ({detalle})")
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

    print("\nLeyendo los archivos de errores:")
    por_validacion = cargar_errores(headers, site_id, tipos, spec)
    if not por_validacion:
        print("\n  No hay errores que enviar.")
        return 0

    print("\nEnviando:")
    tot_env, tot_fall, tot_sin = enviar_todo(
        headers, site_id, maestra, por_validacion, spec, args.prueba, args.supervisor)

    print(f"\n  Enviados: {tot_env} | Fallidos: {tot_fall} | Sin correo: {tot_sin}")
    return 1 if tot_fall else 0


if __name__ == "__main__":
    sys.exit(main())
