# SCRIPTS — Eficacia (canal droguerías, Kenvue)

Todo corre en la nube: los datos viven en SharePoint (Microsoft Graph API) y las
corridas se lanzan desde **GitHub Actions** (ver `.github/workflows/README.md` para
el detalle de cada workflow y en qué orden correrlos).

Hay dos sistemas separados en esta carpeta:

- **ETL principal** (abajo) — trae los datos de campo (Involves) y calcula los
  reportes de negocio.
- **ALERTAS** (`ALERTAS/`, tiene su propio README) — sistema de cumplimiento
  por Telegram/correo, no depende del ETL principal.

## Parte 1 — ETL principal

| Script | Qué hace |
|---|---|
| `run_all.py` | Orquestador: corre todos los ETLs del periodo actual y regenera GOLD. |
| `etl_cif.py` | ETL del módulo CIF. |
| `etl_nopresencia.py` | ETL de No Presencia. |
| `etl_precios.py` | Consolida las capturas de precios (lee `Respuestas_Encuesta`, genera `ANALISIS_PRECIOS`). |
| `etl_sos.py` | Consolida las encuestas SOS/espacios (genera `Encuesta_Sos_Consolidada`). |
| `etl_exhibiciones_gratis.py` | Exhibiciones gratis — junta "Base Exhibiciones Informar" + "Planning", genera `Resultado exhibiciones gratis.xlsx`. |
| `etl_exhibiciones_pagadas.py` | Exhibiciones pagadas. |
| `etl_impactos.py` / `etl_impactos_segmentos.py` | Impactos D&P (Cuota, Venta, Clientes) y su dashboard de segmentos. |
| `etl_ventas.py` | Consolida Ventas D&P. |
| `generar_gold.py` | Genera la capa GOLD (CSV) para Power BI. |
| `generar_sql_tablas.py` | Emite el SQL de las tablas de detalle en Supabase. |

**Infraestructura compartida** (no se corren solos):

| Script | Para qué |
|---|---|
| `paths.py` | Todas las rutas de SharePoint (BASES y SALIDAS) por módulo y año. |
| `periodo_resolver.py` | Resuelve (mes, año) → nombres de archivo concretos. |
| `shared_loader.py` | Lee el Plan de Trabajo una sola vez para todos los ETLs. |
| `supabase_io.py` | Carga los archivos de detalle a Supabase. |
| `etl_logger.py` / `run_log.py` | Logging centralizado y registro de corridas (JSONL). |

## Parte 2 — Ciclo de corrección de errores

Detecta errores de captura (precios, exhibiciones, espacios), se los manda a
cada supervisor por correo, recibe la respuesta corregida y la aplica al
archivo real. Es una secuencia de 4 pasos — ver el detalle completo en
`.github/workflows/README.md`.

| Script | Parte del ciclo |
|---|---|
| `etl_precios_errores.py` | Detecta precios fuera de rango. |
| `etl_exhibiciones_errores.py` | Detecta cantidades de exhibiciones sospechosas. |
| `etl_espacios_errores.py` | Detecta errores de espacios (SOS): universo superado y cambios de participación. |
| `envio_errores.py` | Manda el correo con los errores a cada supervisor. |
| `capturar_respuestas_correo.py` | Guarda el adjunto que el supervisor responde por correo. |
| `aplicar_correcciones.py` | Escribe cada corrección en el archivo real. |

## Utilidades sueltas

| Script | Para qué |
|---|---|
| `local_server.py` | Servidor Flask+SQLite para pruebas presenciales sin SharePoint. |
| `run_revision_local.py` | Orquestador de una sesión de pruebas presencial. |
| `paths_seguridad.py` | Rutas equivalentes a `paths.py` para el módulo de seguridad. |

## Dónde se ajustan las cosas que cambian seguido

- **Supervisores y sus correos:** `BASES DE RESPUESTAS/ALERTAS/MAESTRO_SUPERVISORES.xlsx` en SharePoint (columnas `NOMBRE_SUPERVISOR`, `CORREO`) — la usan tanto `envio_errores.py` como el sistema de Alertas.
- **Umbrales de detección de errores** (qué tan grande debe ser un error para marcarse): están fijos dentro de cada `etl_<módulo>_errores.py`, a propósito — así dos periodos son comparables. Cambiarlos es editar código, no un archivo de configuración.
