"""
core/excel_exporter.py
======================
Genera el reporte Excel consolidado con formato profesional.

Fixes aplicados:
  - Hoja 1: fórmula Plataformas Activas sin #DIV/0! (usa SUMPRODUCT con IF)
  - Hoja 1: columnas A/B con ancho correcto, KPIs bien ajustados
  - Hoja 2: anchos de columna calculados según contenido real (no fijos)
  - Hoja 2: IDs de todas las plataformas visibles (ya no NaN por BOM)
  - Hoja 1/2/3: KPIs calculados directamente desde Python (valores reales,
    sin depender de referencias cross-sheet que pueden fallar)
"""

from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# ─────────────────────────────────────────────────────────────────────────────
# PALETA (hex ARGB para openpyxl)
# ─────────────────────────────────────────────────────────────────────────────
C = {
    "header_bg":  "FF003366",
    "header_fg":  "FFFFFFFF",
    "summary_bg": "FF0D2137",
    "kpi_row_bg": "FF161B22",
    "kpi_fg":     "FF58A6FF",
    "border":     "FFB8C8D8",
    "booking_bg": "FFD6E4F7",
    "airbnb_bg":  "FFFDE8E9",
    "expedia_bg": "FFFFF3CC",
    "white":      "FFFFFFFF",
    "note_fg":    "FF8B949E",
    "data_fg":    "FF1A1A2E",
}

PLATFORM_BG = {
    "Booking": C["booking_bg"],
    "Airbnb":  C["airbnb_bg"],
    "Expedia": C["expedia_bg"],
}

def _border(color=C["border"]):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

THIN   = _border()
H_FONT = Font(name="Calibri", bold=True, color=C["header_fg"], size=10)
H_FILL = PatternFill("solid", fgColor=C["header_bg"])
H_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
D_FONT  = Font(name="Calibri", size=9, color=C["data_fg"])

COLUMN_LABELS = {
    "id_reserva":   "ID Reserva",
    "plataforma":   "Plataforma",
    "huesped":      "Huésped",
    "checkin":      "Check-In",
    "checkout":     "Check-Out",
    "noches":       "Noches",
    "habitacion":   "Habitación",
    "adultos":      "Adultos",
    "estado":       "Estado",
    "tarifa_bruta": "Tarifa Bruta ($)",
    "comision":     "Comisión ($)",
    "tarifa_neta":  "Tarifa Neta ($)",
}


# ─────────────────────────────────────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────────────────────────────────────
def _auto_col_width(ws, col_idx: int, df_col: pd.Series, header: str, min_w=8, max_w=40) -> float:
    """
    Calcula el ancho de columna óptimo basado en el contenido real:
    toma el máximo entre el header y los datos (hasta 200 filas de muestra),
    aplica un factor de escala de caracteres a unidades Excel.
    """
    # Muestra representativa para no iterar millones de filas
    sample = df_col.dropna().astype(str).head(200)
    max_data = sample.map(len).max() if not sample.empty else 0
    max_len  = max(len(str(header)), int(max_data))
    # Factor empírico: Excel char width ≈ 1.2 × character count + 2 padding
    width = max_len * 1.15 + 2
    return max(min_w, min(width, max_w))


def _write_header_row(ws, headers: list, row: int = 1, height: int = 22):
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=ci, value=h)
        c.font      = H_FONT
        c.fill      = H_FILL
        c.alignment = H_ALIGN
        c.border    = THIN
    ws.row_dimensions[row].height = height


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIÓN PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────
def export_to_excel(unified_df: pd.DataFrame, output_path: str) -> tuple[bool, str]:
    if unified_df.empty:
        return False, "No hay datos para exportar."
    try:
        wb = Workbook()
        _build_summary_sheet(wb, unified_df)
        _build_detail_sheet(wb, unified_df)
        _build_platform_sheet(wb, unified_df)
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
        wb.save(output_path)
        return True, f"✓ Reporte guardado en: {output_path}"
    except Exception as e:
        return False, f"Error al generar Excel: {str(e)}"


# ─────────────────────────────────────────────────────────────────────────────
# HOJA 1: RESUMEN EJECUTIVO
# Usa valores calculados directamente desde Pandas (sin fórmulas cross-sheet
# que generan #DIV/0! cuando las celdas referenciadas tienen vacíos).
# ─────────────────────────────────────────────────────────────────────────────
def _build_summary_sheet(wb: Workbook, df: pd.DataFrame):
    ws = wb.create_sheet("📊 Resumen Ejecutivo", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 26

    # ── Título ────────────────────────────────────────────────────────────
    ws.merge_cells("A1:B1")
    tc = ws["A1"]
    tc.value     = "🏨 Hotel Channel Manager — Reporte Consolidado"
    tc.font      = Font(name="Calibri", bold=True, size=14, color=C["header_fg"])
    tc.fill      = PatternFill("solid", fgColor=C["summary_bg"])
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 38

    # ── Fecha ─────────────────────────────────────────────────────────────
    ws.merge_cells("A2:B2")
    dc = ws["A2"]
    dc.value     = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    dc.font      = Font(name="Calibri", italic=True, size=9, color=C["note_fg"])
    dc.fill      = PatternFill("solid", fgColor=C["summary_bg"])
    dc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    ws.row_dimensions[3].height = 10   # separador visual

    # ── KPIs — valores reales calculados desde Pandas ─────────────────────
    # Esto evita completamente los #DIV/0! y errores de referencia cruzada.
    plataformas_activas = int(df["plataforma"].nunique())
    total_reservas      = len(df)
    ingresos_brutos     = float(df["tarifa_bruta"].sum())
    total_comisiones    = float(df["comision"].sum())
    ingresos_netos      = float(df["tarifa_neta"].sum())
    promedio_noches     = float(df["noches"].mean())

    kpis = [
        ("📋 Total Reservas",       total_reservas,      None),
        ("🌐 Plataformas Activas",  plataformas_activas, None),
        ("💰 Ingresos Brutos",      ingresos_brutos,     "currency"),
        ("💸 Total Comisiones",     total_comisiones,    "currency"),
        ("🟢 Ingresos Netos",       ingresos_netos,      "currency"),
        ("🌙 Promedio Noches/Res.", promedio_noches,     "decimal"),
    ]

    for i, (label, value, fmt) in enumerate(kpis):
        row = 4 + i

        lc = ws.cell(row=row, column=1, value=label)
        lc.font      = Font(name="Calibri", size=10, color="FFE6EDF3")
        lc.fill      = PatternFill("solid", fgColor=C["kpi_row_bg"])
        lc.alignment = Alignment(vertical="center", indent=1)
        lc.border    = THIN

        vc = ws.cell(row=row, column=2, value=value)
        vc.font      = Font(name="Calibri", bold=True, size=13, color=C["kpi_fg"])
        vc.fill      = PatternFill("solid", fgColor=C["kpi_row_bg"])
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.border    = THIN

        if fmt == "currency":
            vc.number_format = '#,##0.00 "$"'
        elif fmt == "decimal":
            vc.number_format = "0.0"

        ws.row_dimensions[row].height = 30

    # ── Nota al pie ───────────────────────────────────────────────────────
    note_row = 4 + len(kpis) + 1
    ws.merge_cells(f"A{note_row}:B{note_row}")
    nc = ws[f"A{note_row}"]
    nc.value     = "⚠ Las comisiones son estimadas cuando no constan en el reporte original."
    nc.font      = Font(name="Calibri", italic=True, size=8, color=C["note_fg"])
    nc.alignment = Alignment(horizontal="center")
    ws.row_dimensions[note_row].height = 16


# ─────────────────────────────────────────────────────────────────────────────
# HOJA 2: DETALLE DE RESERVAS
# Anchos calculados automáticamente según contenido real de cada columna.
# ─────────────────────────────────────────────────────────────────────────────
def _build_detail_sheet(wb: Workbook, df: pd.DataFrame):
    ws = wb.create_sheet("📋 Reservas", 1)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"

    columns = list(COLUMN_LABELS.keys())
    headers = [COLUMN_LABELS[c] for c in columns]

    _write_header_row(ws, headers, row=1, height=24)

    money_cols = {columns.index("tarifa_bruta")+1,
                  columns.index("comision")+1,
                  columns.index("tarifa_neta")+1}
    date_cols  = {columns.index("checkin")+1,
                  columns.index("checkout")+1}
    int_cols   = {columns.index("noches")+1,
                  columns.index("adultos")+1}

    for row_idx, row_data in enumerate(df[columns].itertuples(index=False), start=2):
        platform       = str(getattr(row_data, "plataforma", ""))
        row_fill_color = PLATFORM_BG.get(platform, C["white"])

        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)

            # Limpiar NaN / NaT → celda vacía
            try:
                is_na = pd.isna(value)
            except (TypeError, ValueError):
                is_na = False

            if is_na or str(value) == "nan":
                cell.value = ""
            else:
                cell.value = value

            cell.font      = D_FONT
            cell.fill      = PatternFill("solid", fgColor=row_fill_color)
            cell.alignment = Alignment(vertical="center")
            cell.border    = THIN

            if col_idx in money_cols:
                cell.number_format = '#,##0.00'
                cell.alignment     = Alignment(horizontal="right", vertical="center")
            elif col_idx in date_cols and hasattr(value, "strftime"):
                cell.number_format = "DD/MM/YYYY"
                cell.alignment     = Alignment(horizontal="center", vertical="center")
            elif col_idx in int_cols:
                cell.alignment     = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[row_idx].height = 17

    # ── Tabla Excel con filtros ───────────────────────────────────────────
    last_col = get_column_letter(len(columns))
    last_row = len(df) + 1
    table = Table(displayName="Reservas", ref=f"A1:{last_col}{last_row}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False, showLastColumn=False,
        showRowStripes=True,  showColumnStripes=False,
    )
    ws.add_table(table)

    # ── Anchos automáticos basados en contenido real ──────────────────────
    # Mínimos forzados por columna para que no queden demasiado angostas
    min_widths = {
        "id_reserva": 12, "plataforma": 11, "huesped":    20,
        "checkin":    11, "checkout":   11, "noches":      8,
        "habitacion": 22, "adultos":     8, "estado":     13,
        "tarifa_bruta": 15, "comision": 13, "tarifa_neta": 14,
    }
    for col_idx, col_name in enumerate(columns, start=1):
        col_letter = get_column_letter(col_idx)
        header     = COLUMN_LABELS[col_name]
        series     = df[col_name]
        min_w      = min_widths.get(col_name, 10)
        width      = _auto_col_width(ws, col_idx, series, header, min_w=min_w)
        ws.column_dimensions[col_letter].width = width


# ─────────────────────────────────────────────────────────────────────────────
# HOJA 3: RESUMEN POR PLATAFORMA
# ─────────────────────────────────────────────────────────────────────────────
def _build_platform_sheet(wb: Workbook, df: pd.DataFrame):
    ws = wb.create_sheet("📈 Por Plataforma", 2)
    ws.sheet_view.showGridLines = False

    headers = ["Plataforma", "Reservas", "Noches Totales",
               "Ingreso Bruto ($)", "Comisiones ($)", "Ingreso Neto ($)", "% del Total"]
    _write_header_row(ws, headers, row=1, height=22)

    summary = (
        df.groupby("plataforma")
        .agg(
            reservas=("id_reserva", "count"),
            noches=("noches", "sum"),
            bruto=("tarifa_bruta", "sum"),
            comision=("comision", "sum"),
            neto=("tarifa_neta", "sum"),
        )
        .reset_index()
        .sort_values("neto", ascending=False)
        .reset_index(drop=True)
    )
    total_neto = summary["neto"].sum()

    for ri, row in summary.iterrows():
        er    = ri + 2
        color = PLATFORM_BG.get(row["plataforma"], C["white"])
        pct   = row["neto"] / total_neto if total_neto else 0
        vals  = [row["plataforma"], int(row["reservas"]), int(row["noches"]),
                 float(row["bruto"]), float(row["comision"]), float(row["neto"]), pct]
        fmts  = [None, "int", "int", "cur", "cur", "cur", "pct"]

        for ci, (val, vfmt) in enumerate(zip(vals, fmts), 1):
            c = ws.cell(row=er, column=ci, value=val)
            c.font      = D_FONT
            c.fill      = PatternFill("solid", fgColor=color)
            c.alignment = Alignment(horizontal="right" if ci > 1 else "left",
                                    vertical="center")
            c.border    = THIN
            if vfmt == "cur": c.number_format = '#,##0.00'
            if vfmt == "pct": c.number_format = "0.0%"
        ws.row_dimensions[er].height = 20

    # Fila de totales
    tr = len(summary) + 2
    totals = ["TOTAL",
              int(summary["reservas"].sum()),
              int(summary["noches"].sum()),
              float(summary["bruto"].sum()),
              float(summary["comision"].sum()),
              float(summary["neto"].sum()),
              1.0]
    fmts   = [None, "int", "int", "cur", "cur", "cur", "pct"]
    for ci, (val, vfmt) in enumerate(zip(totals, fmts), 1):
        c = ws.cell(row=tr, column=ci, value=val)
        c.font      = Font(name="Calibri", bold=True, size=10, color=C["header_fg"])
        c.fill      = H_FILL
        c.alignment = Alignment(horizontal="right" if ci > 1 else "left",
                                vertical="center")
        c.border    = THIN
        if vfmt == "cur": c.number_format = '#,##0.00'
        if vfmt == "pct": c.number_format = "0.0%"
    ws.row_dimensions[tr].height = 22

    # Anchos
    for ci, (h, col_name) in enumerate(zip(headers, 
        ["plataforma","reservas","noches","bruto","comision","neto","pct"]), 1):
        series = summary.get(col_name, pd.Series(dtype=str)) if col_name in summary.columns else pd.Series(dtype=str)
        ws.column_dimensions[get_column_letter(ci)].width = _auto_col_width(
            ws, ci, series, h, min_w=10
        )
