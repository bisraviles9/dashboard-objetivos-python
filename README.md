# Dashboard de objetivos

Aplicación en Python + Streamlit para revisar cumplimiento por área, avance mensual, objetivos trimestrales, proyección de cierre e integrantes que requieren apoyo. Incluye un archivo Excel editable y descarga del Excel actualizado desde la aplicación.

> Los datos iniciales son demostrativos: reemplázalos por los de tu equipo.

## Ejecutar localmente

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Despliegue en Streamlit Community Cloud

1. Sube este repositorio a GitHub.
2. En Streamlit Community Cloud, crea una app desde ese repositorio.
3. Selecciona `app.py` como archivo principal.
4. La plataforma instalará las dependencias de `requirements.txt`.

## Excel incluido

`dashboard_objetivos.xlsx` tiene tres pestañas:
- **Resumen**: cumplimiento promedio calculado.
- **Objetivos**: área, meta, avance actual, responsable, unidad, fecha límite y trimestre.
- **Seguimiento mensual**: observaciones de progreso mensual.

La aplicación permite cargar un Excel que tenga una hoja `Objetivos` y opcionalmente `Seguimiento mensual`, editar objetivos en pantalla y descargar los cambios como Excel.
