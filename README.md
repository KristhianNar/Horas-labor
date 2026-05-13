# Tablero Horas Labor · Super de Alimentos

Dashboard interactivo estilo Tableau (self-contained HTML) que consume los
datos desde un Excel ubicado en `data/` y se regenera automáticamente en cada
push mediante **GitHub Actions** + **GitHub Pages**.

## Estructura del repositorio

```
.
├── .github/workflows/dashboard.yml   # Workflow: build + deploy a Pages
├── data/
│   └── HorasLab.xlsx                 # Excel fuente (reemplazar para actualizar)
├── scripts/
│   ├── generate_dashboard.py         # Script Python: lee xlsx, agrega, renderiza
│   └── template.html                 # Template Jinja2 del dashboard
├── docs/
│   └── dashboard.html                # Salida del workflow (publicada en Pages)
└── README.md
```

## Cómo funciona

1. Subes o reemplazas `data/HorasLab.xlsx` (cualquier .xlsx sirve, se toma el
   primero que encuentre).
2. El workflow `dashboard.yml` corre en GitHub Actions:
   - Instala Python 3.11 + `pandas`, `openpyxl`, `jinja2`.
   - Ejecuta `scripts/generate_dashboard.py`, que:
     - Detecta la primera hoja y lee todas sus columnas.
     - Normaliza tipos (numéricos, fechas, categóricos) usando `errors='coerce'`.
     - Agrega los datos a nivel *(Cédula, Año, Mes, Área, Subárea)*.
     - Construye un payload columnar compacto y lo embebe en el HTML como
       `const DASHBOARD_DATA = {...}`.
     - Renderiza `scripts/template.html` con Jinja2 y escribe
       `docs/dashboard.html`.
   - Despliega la carpeta `docs/` en GitHub Pages con
     `actions/deploy-pages@v4`.

El HTML final es **un solo archivo**, funciona **offline** una vez generado y
usa solo CDNs para CSS/JS (Bootstrap 5.3, Bootstrap Icons, Chart.js,
SweetAlert2).

## Variables y filtros del dashboard

**Filtros globales:** Año, Área, Subárea (multi-selección).

**Granularidad temporal: semana.** Los buckets semanales siguen la regla del
negocio:
- **2025:** semanas 1 a 27 individuales. A partir de la 28 los datos ya vienen
  consolidados en las semanas-bloque {28, 31, 34, 37, 40, 43, 46, 49, 52}.
  Las filas marcadas con semanas 29, 30, 32, 33, 35, 36, … son residuales
  (ajustes de muy pocos registros) y se descartan: sumarlas al bloque infla
  el total (por ejemplo, la semana 28/2025 real = 413,656 h; incluir los
  residuales la llevaba a 441,999 h).
- **2026+:** los datos ya vienen bloqueados (3, 6, 9, …) y se usan tal cual.

**Gráfica principal:** `TLB_TotalHorasLaboradas` por Semana × Año + línea de
colaboradores únicos (cédulas distintas).

**Gráfica secundaria:** `TLB_TotalHorasLaboradas` agrupado por `TLB_Area` +
colaboradores únicos.

**Pestañas detalladas** (una por variable), cada una con filtros propios
adicionales (Año, Semana, Área, Subárea), gráfica semana-vs-semana y tarjetas
con % variación y proporción sobre `TLB_TotalHorasLaboradas`:

- TLB_HorasDiurnasOrdinarias
- TLB_HorasNocturnasOrdinarias
- TLB_HorasExtraDiurnas
- TLB_HorasExtraNocturnas
- TLB_HorasExtraFestivasDiurnas
- TLB_HorasExtraFestivasNocturnas
- TLB_HorasFestivasDiurnasDiaDescanso
- TLB_HorasFestivasNocturnasDiaDescanso
- TLB_HorasDiurnasDiaDescanso
- TLB_HorasNocturnasDiaDescanso
- TLB_TotalHorasOrdinarias
- TLB_HorasExtra
- Exceso horas extras

Cada gráfica incluye el **número total de colaboradores** (cédulas únicas
agrupadas por Semana y Año).

## Ejecución local

Requiere Python 3.11+.

```bash
pip install pandas==2.2.2 openpyxl==3.1.5 jinja2==3.1.4

# Dejar el .xlsx en data/ y correr:
python scripts/generate_dashboard.py

# Abrir el resultado:
start docs/dashboard.html      # Windows
open  docs/dashboard.html      # macOS
xdg-open docs/dashboard.html   # Linux
```

## Configurar GitHub Pages

1. Crear el repo y hacer push:

   ```bash
   gh repo create horas-labor --public --source=. --remote=origin --push
   ```

2. Habilitar GitHub Pages (una sola vez):

   - UI: *Settings → Pages → Source: GitHub Actions*.
   - O por CLI:

     ```bash
     gh api -X POST repos/:owner/:repo/pages -f build_type=workflow
     ```

3. Disparar el workflow manualmente:

   ```bash
   gh workflow run dashboard.yml
   gh run watch
   ```

4. Abrir el dashboard publicado:

   ```bash
   gh browse --branch gh-pages     # o ver la URL en Actions → deploy
   ```

## Actualizar los datos

1. Reemplaza `data/HorasLab.xlsx` por la versión nueva.
2. `git commit` + `git push` a `main`.
3. El workflow se dispara automáticamente (push con cambios en `data/**.xlsx`)
   y regenera `docs/dashboard.html`.

## Exportar CSV

El botón **Exportar CSV** (arriba a la derecha en el dashboard) descarga los
datos agregados por Mes × Año con los filtros actualmente aplicados y muestra
una alerta de confirmación con SweetAlert2.

## Notas técnicas

- Bucket semanal calculado en `scripts/generate_dashboard.py::week_bucket`:
  para 2025, individual hasta la semana 27 y sólo se aceptan las semanas-
  bloque {28, 31, 34, 37, 40, 43, 46, 49, 52} (el resto son residuales y se
  descartan); para 2026+, la semana raw es ya el bucket.
- Payload columnar (arrays paralelos + diccionarios de strings) para
  minimizar el tamaño del HTML y mantener los filtros en tiempo real.
- Conteo de colaboradores usa `Set` sobre índices de cédula, garantizando
  unicidad al combinar filtros.
