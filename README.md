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
- `JPQ_USE_LOCAL_FALLBACK`: define si usa respaldo local cuando falla Google Sheets.
	- `1` / `true` / `yes` (por defecto): habilitado.
	- `0` / `false` / `no`: deshabilitado (modo solo Google Sheets).

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
