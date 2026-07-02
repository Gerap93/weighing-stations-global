# Weighing Stations · Global

Dashboard de estaciones de pesaje en **HTML + data.json** (sin Power BI), desplegado
en Cloudflare Pages. Este despliegue **global** incluye los cuatro centros:

- **Séneca** (Barcelona · Dory) — modelo de fórmulas (líneas, lab/robot, ingredientes)
- **Granollers** (Roxane) — modelo de muestras (órdenes, pours, sample requests, estaciones MWS/ROXY)
- **Dubai** (Dory) — modelo de fórmulas (100 % pesaje manual, sin robot)
- **Singapore** (Dory) — modelo de fórmulas (100 % pesaje manual, sin robot)

El dashboard **auto-detecta** qué centros están presentes: al cargar hace `fetch` de cada
`data.json` y muestra solo los que existen. El mismo `index.html` sirve para todos los
despliegues (global, spain, dubai, singapore); lo único que cambia son los `data.json`
incluidos.

## Estructura

```
index.html                     · dashboard (auto-detecta centros)
data/data.json                 · Séneca (regenera build_data.py)
location/granollers/data.json  · Granollers (regenera build_roxane.py)
location/dubai/data.json       · Dubai (regenera build_dubai.py)
location/singapore/data.json   · Singapore (regenera build_singapore.py)
_headers                       · Cloudflare: data.json sin caché
```

## Regenerar datos

Los datos crudos (CSV/XLSX) **no se versionan** (`.gitignore`). Para actualizar, colocar
los archivos fuente en `location/<centro>/` y ejecutar el parser correspondiente:

```
python build_data.py       # Séneca     → data/data.json
python build_roxane.py     # Granollers  → location/granollers/data.json
python build_dubai.py      # Dubai      → location/dubai/data.json
python build_singapore.py  # Singapore  → location/singapore/data.json
```

Idiomas: ES / EN / FR. Comparativa de centros disponible cuando hay ≥2 centros.
