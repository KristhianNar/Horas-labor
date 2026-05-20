"""
Genera el dashboard HTML (self-contained) a partir del primer .xlsx encontrado
en /data/.

- Detecta automáticamente hojas, columnas, tipos de datos (numérico, categórico, fecha).
- Agrega los datos a nivel (Cédula, Año, Mes, Área, Subárea) con suma de todas las métricas.
- Embebe los datos como `const DASHBOARD_DATA = {...}` en el HTML final (docs/dashboard.html).
- Renderiza usando Jinja2 el template en scripts/template.html.
- Tolerante a celdas vacías, tipos mixtos y fechas (errors='coerce').
"""

from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape

# ----------------------------------------------------------------------------
# Rutas del proyecto
# ----------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
DOCS_DIR = ROOT / "docs"
SCRIPTS_DIR = ROOT / "scripts"
TEMPLATE_FILE = SCRIPTS_DIR / "template.html"
OUTPUT_FILE = DOCS_DIR / "dashboard.html"

# ----------------------------------------------------------------------------
# Columnas esperadas (mapeo lógico)
# ----------------------------------------------------------------------------
METRIC_COLUMNS = [
    "TLB_HorasDiurnasOrdinarias",
    "TLB_HorasNocturnasOrdinarias",
    "TLB_HorasExtraDiurnas",
    "TLB_HorasExtraNocturnas",
    "TLB_HorasExtraFestivasDiurnas",
    "TLB_HorasExtraFestivasNocturnas",
    "TLB_HorasFestivasDiurnasDiaDescanso",
    "TLB_HorasFestivasNocturnasDiaDescanso",
    "TLB_HorasDiurnasDiaDescanso",
    "TLB_HorasNocturnasDiaDescanso",
    "TLB_TotalHorasOrdinarias",
    "TLB_TotalHorasLaboradas",
    "TLB_HorasExtra",
    "Exceso horas extras",
]

DIMENSION_COLUMNS = {
    "cedula":  "TLB_Cedula",
    "area":    "TLB_Area",
    "subarea": "TLB_Subarea",
    "semana":  "TLB_Semana",
    "anio":    "TLB_Año",
}


def find_first_xlsx(folder: Path) -> Path:
    if not folder.exists():
        raise FileNotFoundError(f"No existe la carpeta {folder}")
    xlsx_files = sorted(folder.glob("*.xlsx"))
    if not xlsx_files:
        raise FileNotFoundError(f"No se encontró ningún .xlsx en {folder}")
    return xlsx_files[0]


BLOCK_WEEKS_2025 = {28, 31, 34, 37, 40, 43, 46, 49, 52}


def week_bucket(year: int, week: int):
    """Mapea (año, semana raw) -> bucket semanal usado en los gráficos.

    Regla acordada con el negocio:
    - 2025: semanas 1-27 individuales. Desde la 28 los datos ya vienen
      consolidados en las semanas-bloque {28, 31, 34, 37, 40, 43, 46, 49, 52}.
      Los valores que aparecen en 29, 30, 32, 33, … son residuales que NO
      deben agregarse al bloque (la base ya trae el total consolidado en la
      semana-bloque). Por eso esas filas se descartan devolviendo ``None``.
    - 2026 en adelante: los datos ya llegan bloqueados (3, 6, 9, …), cada
      número es ya un bucket semanal; se deja tal cual.

    Devuelve ``None`` cuando la fila debe excluirse.
    """
    if pd.isna(week) or pd.isna(year):
        return None
    y = int(year)
    w = int(week)
    if w < 1:
        return None
    if y == 2025:
        if w <= 27:
            return w
        if w in BLOCK_WEEKS_2025:
            return w
        return None
    return w


def load_and_clean(xlsx_path: Path) -> pd.DataFrame:
    print(f"[INFO] Leyendo {xlsx_path.name}")
    xls = pd.ExcelFile(xlsx_path)
    sheet = xls.sheet_names[0]
    print(f"[INFO] Hojas detectadas: {xls.sheet_names} -> usando '{sheet}'")

    df = pd.read_excel(xlsx_path, sheet_name=sheet)
    print(f"[INFO] Filas: {len(df):,}, Columnas: {len(df.columns)}")

    # Normalizar nombres por si vienen con espacios
    df.columns = [str(c).strip() for c in df.columns]

    # Validación suave: avisar si faltan columnas clave
    missing = [c for c in list(DIMENSION_COLUMNS.values()) + METRIC_COLUMNS if c not in df.columns]
    if missing:
        print(f"[WARN] Columnas faltantes: {missing}")

    # Asegurar tipos numéricos
    for col in METRIC_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    for col in (DIMENSION_COLUMNS["semana"], DIMENSION_COLUMNS["anio"]):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    for col in (DIMENSION_COLUMNS["cedula"],):
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    for col in (DIMENSION_COLUMNS["area"], DIMENSION_COLUMNS["subarea"]):
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace({"nan": "SIN DATO", "": "SIN DATO"})

    # Fechas (si existen) - usamos errors='coerce'
    for col in ("TLB_FechaCreacion", "TLB_FechaModificacion"):
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

    # Filtrar año válido
    anio_col = DIMENSION_COLUMNS["anio"]
    if anio_col in df.columns:
        df = df[df[anio_col] > 0].copy()

    # Derivar bucket semanal y descartar residuales (2025 > 27 que no son bloque)
    semana_col = DIMENSION_COLUMNS["semana"]
    if semana_col in df.columns and anio_col in df.columns:
        df["__Semana"] = df.apply(
            lambda r: week_bucket(r[anio_col], r[semana_col]), axis=1
        )
        antes = len(df)
        df = df[df["__Semana"].notna()].copy()
        df["__Semana"] = df["__Semana"].astype(int)
        print(f"[INFO] Filas descartadas (residuales 2025 o semanas inválidas): {antes - len(df):,}")
    else:
        df["__Semana"] = 0

    return df


def aggregate(df: pd.DataFrame) -> pd.DataFrame:
    dims = [
        DIMENSION_COLUMNS["cedula"],
        DIMENSION_COLUMNS["anio"],
        "__Semana",
        DIMENSION_COLUMNS["area"],
        DIMENSION_COLUMNS["subarea"],
    ]
    dims = [d for d in dims if d in df.columns]
    metrics = [m for m in METRIC_COLUMNS if m in df.columns]

    print(f"[INFO] Agregando por {dims} — métricas: {len(metrics)}")
    agg = df.groupby(dims, as_index=False, dropna=False)[metrics].sum()
    print(f"[INFO] Filas agregadas: {len(agg):,}")
    return agg


def build_dashboard_data(agg: pd.DataFrame) -> dict:
    """Convierte el dataframe agregado a un dict columnar compacto."""
    anio_col = DIMENSION_COLUMNS["anio"]
    cedula_col = DIMENSION_COLUMNS["cedula"]
    area_col = DIMENSION_COLUMNS["area"]
    sub_col = DIMENSION_COLUMNS["subarea"]

    years = sorted(int(y) for y in agg[anio_col].unique() if not pd.isna(y)) if anio_col in agg.columns else [0]
    areas = sorted(str(a) if not pd.isna(a) else "SIN DATO" for a in agg[area_col].unique()) if area_col in agg.columns else ["SIN DATO"]
    subareas = sorted(str(s) if not pd.isna(s) else "SIN DATO" for s in agg[sub_col].unique()) if sub_col in agg.columns else ["SIN DATO"]
    cedulas = sorted(str(c) for c in agg[cedula_col].unique() if not pd.isna(c)) if cedula_col in agg.columns else ["0"]

    year_idx = {y: i for i, y in enumerate(years)}
    area_idx = {a: i for i, a in enumerate(areas)}
    sub_idx = {s: i for i, s in enumerate(subareas)}
    ced_idx = {c: i for i, c in enumerate(cedulas)}

    # Normalise area/subarea/cedula columns in agg to str so mapping never produces NaN
    if area_col in agg.columns:
        agg = agg.copy()
        agg[area_col] = agg[area_col].apply(lambda v: "SIN DATO" if pd.isna(v) else str(v))
    if sub_col in agg.columns:
        agg[sub_col] = agg[sub_col].apply(lambda v: "SIN DATO" if pd.isna(v) else str(v))
    if cedula_col in agg.columns:
        agg[cedula_col] = agg[cedula_col].apply(lambda v: "SIN_CEDULA" if pd.isna(v) else str(v))

    data = {
        "y": agg[anio_col].map(year_idx).fillna(0).astype(int).tolist() if anio_col in agg.columns else [0] * len(agg),
        "w": agg["__Semana"].astype(int).tolist(),
        "a": agg[area_col].map(area_idx).fillna(0).astype(int).tolist() if area_col in agg.columns else [0] * len(agg),
        "s": agg[sub_col].map(sub_idx).fillna(0).astype(int).tolist() if sub_col in agg.columns else [0] * len(agg),
        "c": agg[cedula_col].map(ced_idx).fillna(0).astype(int).tolist() if cedula_col in agg.columns else [0] * len(agg),
    }

    for i, col in enumerate(METRIC_COLUMNS):
        key = f"v{i}"
        if col in agg.columns:
            data[key] = [round(float(x), 2) for x in agg[col].tolist()]
        else:
            data[key] = [0.0] * len(agg)

    return {
        "years": years,
        "areas": areas,
        "subareas": subareas,
        "cedulas": cedulas,
        "data": data,
    }


def render(dashboard_data: dict, source_file: str, rows_count: int) -> str:
    env = Environment(
        loader=FileSystemLoader(str(SCRIPTS_DIR)),
        autoescape=select_autoescape(disabled_extensions=("html",), default=False),
        variable_start_string="{{ ",
        variable_end_string=" }}",
    )
    template = env.get_template("template.html")
    return template.render(
        DASHBOARD_DATA=json.dumps(dashboard_data, ensure_ascii=False, separators=(",", ":")),
        GENERATED_AT=datetime.now().strftime("%Y-%m-%d %H:%M"),
        SOURCE_FILE=source_file,
        ROWS_COUNT=f"{rows_count:,}",
    )


def main() -> int:
    try:
        DOCS_DIR.mkdir(parents=True, exist_ok=True)
        xlsx_path = find_first_xlsx(DATA_DIR)
        df = load_and_clean(xlsx_path)
        agg = aggregate(df)
        payload = build_dashboard_data(agg)
        html = render(payload, xlsx_path.name, len(df))
        OUTPUT_FILE.write_text(html, encoding="utf-8")
        size_mb = OUTPUT_FILE.stat().st_size / (1024 * 1024)
        print(f"[OK] Dashboard generado: {OUTPUT_FILE} ({size_mb:.2f} MB)")
        return 0
    except Exception as exc:  # pragma: no cover
        print(f"[ERROR] {exc}", file=sys.stderr)
        raise


if __name__ == "__main__":
    sys.exit(main())
