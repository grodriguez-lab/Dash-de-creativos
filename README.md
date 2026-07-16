# Monitoreo de Creativos — Dashboard

Dashboard estático (sin backend, sin costo) que muestra el semáforo de creativos de Meta
para el equipo de Growth de Retorna. Se actualiza manualmente cada jueves después de correr
la skill `monitoreo-creativos-meta`.

## Estructura del repo

```
index.html            <- el dashboard (no lo edites salvo que quieras cambiar diseño)
data.json              <- los datos de la semana (esto es lo único que se reemplaza cada jueves)
export_dashboard.py    <- script que genera data.json a partir del Excel
```

## Setup único (la primera vez)

1. Crea un repo nuevo en GitHub (público), por ejemplo `monitoreo-creativos-dashboard`.
2. Sube estos 3 archivos a la raíz del repo (`index.html`, `data.json`, `export_dashboard.py`).
3. Ve a **Settings → Pages** del repo → en "Branch" elige `main` y carpeta `/ (root)` → Save.
4. GitHub te da una URL fija, algo como:
   `https://tu-usuario.github.io/monitoreo-creativos-dashboard/`
   Ese es el link que compartes con tu equipo. No cambia nunca.

## Flujo semanal (cada jueves)

1. Corre la skill `monitoreo-creativos-meta` como siempre: BigQuery → llenar Excel →
   `recalc.py` → confirmar `total_errors: 0`.
2. Genera el nuevo `data.json` a partir del Excel ya recalculado:
   ```bash
   python export_dashboard.py "<ruta>/Monitoreo_Creativos_Meta.xlsx" data.json --fecha "17 jul 2026"
   ```
3. Sube el `data.json` nuevo al repo. Dos formas, elige la que prefieras:
   - **Sin terminal:** entra al repo en github.com → abre `data.json` → ícono de lápiz
     (editar) → borra todo → pega el contenido nuevo → "Commit changes".
   - **Con terminal:**
     ```bash
     git add data.json
     git commit -m "Monitoreo semana del 17 jul"
     git push
     ```
4. Espera ~30 segundos y refresca el link. Ya está actualizado para todo el equipo.

## Notas

- El Excel (`Monitoreo_Creativos_Meta.xlsx` / el Google Sheet canónico) sigue siendo la
  única fuente de verdad. Este dashboard es solo una "foto" de lectura para el equipo.
- `export_dashboard.py` necesita que el Excel esté **recalculado** (paso de `recalc.py` de
  la skill xlsx) antes de correrlo, porque lee los valores en caché de las fórmulas
  (CPI, ESTADO, Acción, etc). Si corres el export sobre un archivo sin recalcular, esas
  columnas saldrán vacías.
- Si cambias nombres de columnas en el Excel, ajusta el diccionario `COLS` al inicio de
  `export_dashboard.py`.
- El repo es público por defecto en el plan gratis de GitHub Pages (no aparece en buscadores,
  pero cualquiera con el link puede verlo). Si necesitas que solo tu equipo lo vea con login,
  la alternativa gratis es Cloudflare Pages + Cloudflare Access (hasta 50 usuarios gratis).
