# JPQ Llaves - Streamlit

App de Streamlit para publicar en un solo link el contenido de la planilla Google Sheets de JPQ. Si la planilla remota no responde, usa como respaldo los Excel de la carpeta `LLAVES 1er JPQ`, tomando solamente la hoja que **no** empieza con `Base`.

## Qué hace

- Lee todas las hojas de la planilla Google Sheets configurada en `main.py`.
- Ignora hojas cuyo nombre comience con `Base`.
- Usa la categoría según el nombre de la hoja (por ejemplo, `c3` -> `c3`).
- Si falla la carga remota, intenta leer los `.xlsx` locales de `LLAVES 1er JPQ`.
- Consolida todo en una sola tabla.
- Permite filtrar por categoría desde la interfaz.

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
