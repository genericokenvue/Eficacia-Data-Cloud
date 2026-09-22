# Manual — Proceso de Datos y Corrección de Errores (Eficacia)

Sistema que trae la captura de campo de las droguerías (Kenvue, Colombia), arma
los reportes de negocio, y detecta + corrige errores de digitación en precios,
exhibiciones y espacios. Todo vive en la nube (SharePoint) y se corre desde
GitHub Actions — no hace falta tener nada instalado en un computador.

---

## 1. Qué es esto, en una frase

Cada mes el equipo de campo captura precios, exhibiciones y espacios en las
droguerías. Este sistema junta esa información, arma reportes, y además
revisa si algo se digitó mal (un precio, una cantidad) — si encuentra algo
sospechoso, le avisa al supervisor responsable por correo, recibe su
corrección, y la aplica sola.

Son **dos sistemas separados**:

| | Qué es | Cadencia |
|---|---|---|
| **ETL principal** (`SCRIPTS/`) | Trae los datos y calcula reportes + corrige errores | Diario / fin de mes |
| **Alertas** (`SCRIPTS/ALERTAS/`) | Avisa a cada supervisor su cumplimiento, por Telegram y correo | Diario |

---

## 2. Guía de usuario

### Cómo se corre

Todo se lanza desde GitHub, pestaña **Actions**: eliges el workflow en la
lista de la izquierda → botón **Run workflow** → completas los campos que
pida → **Run workflow** otra vez para confirmar.

Hay una guía completa, paso a paso, con cada campo explicado:
👉 **[Cómo correr todo](https://claude.ai/artifact/En4gNJWPtN58rCv4QxA6ue)**

Resumen rápido:

| Cadencia | Workflow | Qué hace |
|---|---|---|
| Diario | `1B. ETL Diario` | Automático, no hay que tocarlo |
| Diario | `1C. Alertas` | Se lanza a mano cada día |
| Fin de mes | `2A → 2B → 2C → 2D` | Calcular errores → enviar → recibir respuesta → aplicar corrección (en ese orden) |
| Cuando haga falta | `1A. ETL Manual` | Un módulo suelto, o un mes atrasado |

### Qué hay que mantener

**Lista de supervisores y sus correos** — es lo único que cambia seguido:

```
SharePoint → BASES DE RESPUESTAS/ALERTAS/MAESTRO_SUPERVISORES.xlsx
```

Columnas `NOMBRE_SUPERVISOR` y `CORREO`. Si alguien entra, sale o cambia de
correo, se edita ahí — no hay que tocar ningún script.

Todo lo demás (qué tan grande debe ser un error para marcarse, por ejemplo)
está fijo en el código a propósito, para que un mes sea comparable con el
siguiente.

---

## 3. Guía técnica

### Arquitectura

- **Datos:** SharePoint (Microsoft Graph API) — no hay base de datos local.
- **Autenticación:** una app registrada en Azure AD, credenciales tipo
  `client_credentials` (usuario y contraseña no aplican; son 3 secrets:
  `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, guardados en
  GitHub como *repository secrets*).
- **Ejecución:** GitHub Actions. Cada workflow instala Python 3.11,
  `pip install -r requirements.txt`, y corre el script con los `env:` de
  arriba.
- **Estructura en SharePoint:**
  - `BASES DE RESPUESTAS/<módulo>/<año>/` — insumos crudos (lo que trae
    Involves/las encuestas).
  - `SALIDAS/<módulo>/` — lo que el sistema calcula (reportes, archivos de
    error).
  - `SALIDAS/GOLD/` — capa final para Power BI.

### El ciclo de corrección de errores, técnicamente

```
etl_<módulo>_errores.py      → detecta, escribe ID_ERROR en el archivo crudo
        ↓
envio_errores.py             → un correo por supervisor, con sus 3 módulos juntos
        ↓
capturar_respuestas_correo.py → guarda el adjunto que responde el supervisor
        ↓
aplicar_correcciones.py      → escribe el valor corregido en el archivo crudo
```

Puntos clave:

- **`ID_ERROR`** identifica cada caso (`PRE-202608-008` = precios, agosto
  2026, caso 8). El mes va incrustado en el propio ID — la corrección
  siempre se aplica al archivo de ESE mes, sin importar cuándo llegue la
  respuesta.
- **La corrección se escribe en el archivo CRUDO**, no en el reporte
  calculado — porque el reporte se regenera desde cero en cada corrida y
  perdería el cambio. El crudo es el insumo, así que el cambio sobrevive.
- **Columna `CORREGIDO`** (Sí/No/blanco) queda marcada en ese mismo archivo
  crudo: blanco si la fila nunca tuvo error, "No" si tiene error sin
  corregir, "Sí" si ya se corrigió.
- **`capturar_respuestas_correo.py`** reemplaza lo que iba a hacer Power
  Automate: revisa la bandeja de `generico_kenvue@eficacia.com.co`, busca
  correos "CAPTURA DE ERRORES" con adjunto, y mueve el correo a una carpeta
  "Procesados" — pero solo DESPUÉS de guardar bien el adjunto, para que un
  fallo a mitad de camino no pierda el correo.
- **`aplicar_correcciones.py`** reconoce el módulo por las **columnas** de
  cada hoja del Excel (`PRECIO_CORREGIDO`, `CANTIDAD_CORREGIDA`,
  `CM_MARCA_CORREGIDO`/`UNIVERSO_CORREGIDO`), no por el nombre de la hoja —
  así funciona igual con el archivo real o con uno suelto de prueba.
- Nunca se escribe un valor vacío encima de un dato real (protección
  agregada tras un incidente real — ver `LOG_CORRECCIONES.xlsx` para el
  historial de auditoría de cada corrección aplicada).

### Dónde está cada cosa

| Carpeta | Qué contiene |
|---|---|
| `SCRIPTS/` | ETL principal — ver [README.md](README.md) |
| `SCRIPTS/ALERTAS/` | Sistema de alertas — ver [ALERTAS/README.md](ALERTAS/README.md) |
| `SCRIPTS/.github/workflows/` | Los 7 workflows de GitHub Actions — ver su [README.md](.github/workflows/README.md) |
| `SCRIPTS/paths.py` | Todas las rutas de SharePoint |
| `SCRIPTS/periodo_resolver.py` | (mes, año) → nombres de archivo concretos |

---

## 4. Glosario rápido

| Término | Qué significa |
|---|---|
| **ETL** | El proceso que trae, limpia y consolida los datos (Extract-Transform-Load). |
| **Encuesta cruda** | El archivo tal cual sale de la captura de campo, sin procesar. |
| **GOLD** | La capa final de datos, lista para Power BI. |
| **Workflow** | Una tarea automatizada en GitHub Actions (un botón que se puede apretar). |
| **`workflow_dispatch`** | La opción que permite lanzar un workflow a mano, con el botón "Run workflow". |
