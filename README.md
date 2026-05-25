# JPQ Llaves - Streamlit

App de Streamlit para publicar en un solo link el contenido de la planilla Google Sheets de JPQ. Por defecto, si la planilla remota no responde, usa como respaldo los Excel de la carpeta `LLAVES 1er JPQ`, tomando solamente la hoja que **no** empieza con `Base`.

## Qué hace

- Lee todas las hojas de la planilla Google Sheets configurada por variable de entorno (o valor por defecto en código).
- Ignora hojas cuyo nombre comience con `Base`.
- Usa la categoría según el nombre de la hoja (por ejemplo, `c3` -> `c3`).
- Si falla la carga remota, intenta leer los `.xlsx` locales de `LLAVES 1er JPQ`.
- Consolida todo en una sola tabla.
- Permite filtrar por categoría desde la interfaz.

## Configuración (Google Sheets)

Podés configurar la fuente remota sin editar código:

- `JPQ_SHEET_URL`: URL de exportación `.xlsx` de Google Sheets.
- `JPQ_SHEET_ID`: ID del spreadsheet (opcional, recomendado para acceso privado).
- `JPQ_USE_LOCAL_FALLBACK`: define si usa respaldo local cuando falla Google Sheets.
	- `1` / `true` / `yes` (por defecto): habilitado.
	- `0` / `false` / `no`: deshabilitado (modo solo Google Sheets).

Para Google Sheets privado (sin URL publica), agregá una de estas opciones de credenciales:

- `JPQ_GOOGLE_SERVICE_ACCOUNT_FILE`: ruta al JSON de Service Account.
- `JPQ_GOOGLE_SERVICE_ACCOUNT_JSON`: contenido JSON completo (ideal para secrets en deploy).

La app intenta primero descarga pública por `JPQ_SHEET_URL`; si falla, intenta export autenticado con Service Account usando Drive API.

### Pasos para acceso privado (Service Account)

1. En Google Cloud, crear una Service Account y descargar su JSON.
2. Compartir la planilla de Google Sheets con el email de esa Service Account (permiso de lectura).
3. Definir variables de entorno y ejecutar la app.

Ejemplo:

```bash
export JPQ_SHEET_ID="tu_sheet_id"
export JPQ_GOOGLE_SERVICE_ACCOUNT_FILE="/ruta/credenciales.json"
export JPQ_USE_LOCAL_FALLBACK=0
streamlit run main.py
```

Ejemplo de ejecución en modo solo Google Sheets:

```bash
JPQ_USE_LOCAL_FALLBACK=0 streamlit run main.py
```

## Ejecutar localmente

1. Instalar dependencias:

```bash
pip install -e .
```

2. Levantar Streamlit:

```bash
streamlit run main.py
```

3. Abrir el link local que muestra Streamlit (normalmente `http://localhost:8501`).

## Publicar en un solo link

Podés desplegar esta app en Streamlit Community Cloud apuntando al repo y archivo `main.py`.

## Deploy en Streamlit Community Cloud

La app ya soporta configuracion por variables de entorno y tambien por Streamlit Secrets.

1. Subi este repo a GitHub.
2. En Streamlit Community Cloud, creá una app apuntando a `main.py`.
3. En Settings -> Secrets, cargá estas claves:

```toml
JPQ_SHEET_ID = "tu_sheet_id"
JPQ_USE_LOCAL_FALLBACK = "0"

# Opcion A: ruta local (normalmente no aplica en Cloud)
# JPQ_GOOGLE_SERVICE_ACCOUNT_FILE = "/ruta/credenciales.json"

# Opcion B (recomendada en Cloud): JSON completo en una sola linea
JPQ_GOOGLE_SERVICE_ACCOUNT_JSON = "{\"type\":\"service_account\",\"project_id\":\"...\",\"private_key\":\"-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n\",\"client_email\":\"...\",\"token_uri\":\"https://oauth2.googleapis.com/token\"}"
```

Alternativa recomendada para evitar errores de TOML en Cloud:

```toml
JPQ_SHEET_ID = "tu_sheet_id"
JPQ_USE_LOCAL_FALLBACK = "0"

[gcp_service_account]
type = "service_account"
project_id = "padelstats-492015"
private_key_id = "..."
private_key = """-----BEGIN PRIVATE KEY-----
...
-----END PRIVATE KEY-----
"""
client_email = "..."
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
universe_domain = "googleapis.com"
```

Notas:

- Si usás planilla privada, compartila con el email de la Service Account con permiso de lectura.
- En Cloud conviene usar `JPQ_GOOGLE_SERVICE_ACCOUNT_JSON` en Secrets y no commitear ningun archivo de credenciales.
