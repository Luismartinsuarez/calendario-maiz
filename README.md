# Calendario fenológico del maíz (GDD) — Streamlit

Dashboard que estima las fechas de cada etapa del maíz a partir de grados-día de crecimiento
(método del promedio simple con Tb y Tu). Observaciones hasta la fecha de referencia y
climatología diaria del propio archivo para proyectar el resto del ciclo.

## Funciones
- **Estaciones incluidas** en `data/estaciones/` (planillas SENAMHI `.XLS`:
  TM102 = Tmax, TM103 = Tmin; -999 = faltante) más el ejemplo de Busan.
  Para agregar una estación basta copiar su archivo en esa carpeta. El distrito, la
  provincia y el departamento se toman de `data/estaciones/Maestra.xlsx` según el código
  de la planilla (CODIGO o NCODIGO).
- **Subir archivo**: CSV como `ejemplo_busan.csv` (`date,tmax,tmin`), planilla SENAMHI, Excel,
  o CSV horario / de 10 min (se agrega a diario). Las columnas se detectan y se pueden reasignar (en las estaciones incluidas se detectan solas).
- **Variedad** desde `data/gdd_variedades.xlsx` (fijo), hoja `requerimientos`: carga fases (desde panoja
  hasta cosecha; emergencia y aparición de hojas no se muestran), GDD, Tb y método
  (media / acotada). Avisa si la Tmedia de la estación está fuera del rango calibrado.
  Opción «Genérica» con valores de literatura.
- **Fases, GDD, Tb, Tu y método editables** (agregar/quitar filas).
- Estado actual, línea de tiempo, tabla de fechas previstas por fase y curva de GDD acumulados por año.

Formato mínimo del archivo:

```csv
fecha,tmax,tmin
2025-01-01,28.4,17.2
```

## Ejecutar localmente
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicar en Streamlit Community Cloud (gratis)
1. Subir esta carpeta a un repositorio de GitHub (`app.py`, `requirements.txt`,
   `.streamlit/config.toml` y la carpeta `data/`).
2. Entrar a <https://share.streamlit.io>, **Create app** → elegir el repositorio, rama y
   `app.py` como archivo principal → **Deploy**.
3. Los archivos que suben los usuarios solo viven en su sesión; no se guardan en el servidor.
