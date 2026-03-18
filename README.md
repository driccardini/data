# Executive AI Presenter

App en Streamlit para convertir `Resumen-Ejecutivo-2025.pptx` en un resumen ejecutivo y un guion de presentación con enfoque de IA.

## Presentación moderna (reunión)

Para presentar directamente en formato corporativo moderno:

```bash
pip install -e .
streamlit run presentacion_streamlit.py
```

Archivo por defecto usado por la app:

- `Presentacion-Ejecutiva-Corporativa-2025-2026.pptx`

La app abre directamente este archivo (sin uploader/drag-and-drop) para una experiencia de presentación limpia.

## Qué hace

- Lee el `.pptx` y extrae texto por diapositiva.
- Detecta temas dominantes (frecuencia de términos).
- Marca señales de riesgo y oportunidad (palabras clave).
- Genera un guion ejecutivo en tono `Directo`, `Narrativo` o `Data-first`.
- Permite descargar el guion en `.md`.

## Ejecutar

```bash
pip install -e .
streamlit run main.py
```

Luego abre la URL local que muestra Streamlit en tu terminal.
