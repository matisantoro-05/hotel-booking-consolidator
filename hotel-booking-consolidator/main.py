"""
main.py — Hotel Channel Manager · Data Link
============================================
Compatible con Flet 0.85+

Cambios clave respecto a versiones anteriores:
  - FilePicker es completamente async en 0.85:
      archivos = await picker.pick_files(...)
      ruta     = await picker.save_file(...)
  - No existe FilePickerResultEvent ni on_result callback
  - page.run_task(coro) lanza coroutines desde handlers síncronos
  - ft.Border.all()  en vez de ft.border.all()
  - ft.Padding(...)  en vez de ft.padding.all/symmetric/only()
  - ft.Margin(...)   en vez de ft.margin.symmetric()
"""

import flet as ft
import os
import sys
import threading
from pathlib import Path
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.data_processor import HotelDataProcessor
from core.charts import build_monthly_chart, build_channel_pie, build_revenue_trend
from core.excel_exporter import export_to_excel


# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTES DE DISEÑO
# ─────────────────────────────────────────────────────────────────────────────
BG_DARK        = "#0D1117"
BG_PANEL       = "#161B22"
BG_CARD        = "#1C2128"
BORDER_COLOR   = "#30363D"
TEXT_PRIMARY   = "#E6EDF3"
TEXT_SECONDARY = "#8B949E"
ACCENT_BLUE    = "#58A6FF"
ACCENT_GREEN   = "#3FB950"
ACCENT_ORANGE  = "#D29922"
ACCENT_RED     = "#F85149"

PLATFORM_COLORS = {
    "booking": "#003580",
    "airbnb":  "#FF5A5F",
    "expedia": "#FFC72C",
}
PLATFORM_ICONS = {
    "booking": ft.Icons.HOTEL,
    "airbnb":  ft.Icons.HOME,
    "expedia": ft.Icons.FLIGHT_TAKEOFF,
}


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS DE UI
# ─────────────────────────────────────────────────────────────────────────────
def card(content, padding=16, border_radius=10, bgcolor=BG_CARD, width=None):
    p = ft.Padding(left=padding, top=padding, right=padding, bottom=padding)
    return ft.Container(
        content=content, bgcolor=bgcolor, border_radius=border_radius,
        padding=p, border=ft.Border.all(1, BORDER_COLOR), width=width,
    )

def lbl(text, size=11, color=TEXT_SECONDARY, bold=False, italic=False):
    return ft.Text(
        text, size=size, color=color, italic=italic,
        weight=ft.FontWeight.BOLD if bold else ft.FontWeight.NORMAL,
    )

def hdivider():
    return ft.Container(
        height=1, bgcolor=BORDER_COLOR,
        margin=ft.Margin(left=0, top=8, right=0, bottom=8),
    )

def badge_pill(text, bg_color):
    return ft.Container(
        ft.Text(text, size=9, color="white", weight=ft.FontWeight.BOLD),
        bgcolor=bg_color, border_radius=4,
        padding=ft.Padding(left=6, top=2, right=6, bottom=2),
    )

def b64_to_src(b64: str) -> str:
    """Convierte base64 puro a data URI para ft.Image(src=...)."""
    return f"data:image/png;base64,{b64}"


# ─────────────────────────────────────────────────────────────────────────────
# APLICACIÓN PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────
def main(page: ft.Page):
    page.title      = "Hotel Channel Manager — Data Link"
    page.bgcolor    = BG_DARK
    page.theme_mode = ft.ThemeMode.DARK
    page.padding    = 0
    page.window.width      = 1280
    page.window.height     = 820
    page.window.min_width  = 1024
    page.window.min_height = 700

    processor    = HotelDataProcessor()
    loaded_files: dict[str, str] = {}

    # ── Widgets de estado global ──────────────────────────────────────────
    status_text = ft.Text("", size=11, color=TEXT_SECONDARY, italic=True)
    status_dot  = ft.Container(width=8, height=8, border_radius=4, bgcolor=TEXT_SECONDARY)

    def set_status(msg: str, level: str = "info"):
        palette = {
            "info":    TEXT_SECONDARY,
            "success": ACCENT_GREEN,
            "error":   ACCENT_RED,
            "working": ACCENT_BLUE,
        }
        c = palette.get(level, TEXT_SECONDARY)
        status_text.value  = msg
        status_text.color  = c
        status_dot.bgcolor = c

    # ── Barra de progreso global ──────────────────────────────────────────
    progress_bar = ft.ProgressBar(
        value=0,
        bar_height=3,
        color=ACCENT_BLUE,
        bgcolor=ft.Colors.with_opacity(0.12, ACCENT_BLUE),
        visible=False,
    )
    progress_label = ft.Text(
        "", size=10, color=ACCENT_BLUE, italic=True, visible=False,
    )
    progress_label_row = ft.Container(
        content=ft.Row(
            [progress_label],
            alignment=ft.MainAxisAlignment.CENTER,
        ),
        bgcolor=ft.Colors.with_opacity(0.05, ACCENT_BLUE),
        padding=ft.Padding(left=12, top=4, right=12, bottom=4),
        visible=False,
    )

    def set_progress(value, label: str = ""):
        if value == -1:
            progress_bar.visible       = False
            progress_label.visible     = False
            progress_label_row.visible = False
        else:
            progress_bar.visible       = True
            progress_bar.value         = value
            progress_label.visible     = bool(label)
            progress_label_row.visible = bool(label)
            progress_label.value       = label

    # ─────────────────────────────────────────────────────────────────────
    # PANEL 1 — CARGA DE ARCHIVOS
    # ─────────────────────────────────────────────────────────────────────
    file_status = {
        p: ft.Text("Sin archivo cargado", size=10, color=TEXT_SECONDARY, italic=True)
        for p in ["booking", "airbnb", "expedia"]
    }
    file_badges = {p: ft.Container(visible=False) for p in ["booking", "airbnb", "expedia"]}

    preview_table = ft.DataTable(
        columns=[ft.DataColumn(ft.Text("—", size=10, color=TEXT_SECONDARY))],
        rows=[],
        border=ft.Border.all(1, BORDER_COLOR),
        border_radius=8,
        column_spacing=16,
        heading_row_color=ft.Colors.with_opacity(0.06, "white"),
        heading_row_height=36,
        data_row_min_height=30,
        data_row_max_height=30,
        bgcolor=BG_PANEL,
    )
    preview_label = ft.Text(
        "Vista previa — cargue archivos y procese para ver datos",
        size=11, color=TEXT_SECONDARY, italic=True,
    )
    process_btn = ft.ElevatedButton(
        "⚡ Procesar y Unificar Datos",
        bgcolor=ACCENT_BLUE, color="white",
        style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8)),
        height=40, disabled=True,
    )
    process_spinner = ft.ProgressRing(
        width=18, height=18, stroke_width=2, color=ACCENT_BLUE, visible=False,
    )

    def build_preview_table(df):
        """Rellena el DataTable con los primeros 100 registros consolidados."""
        cols_to_show = ["plataforma", "huesped", "checkin", "checkout",
                        "noches", "tarifa_bruta", "comision", "tarifa_neta"]
        col_labels   = ["Canal", "Huésped", "Check-In", "Check-Out",
                        "Noches", "Bruto ($)", "Comisión ($)", "Neto ($)"]

        preview_table.columns = [
            ft.DataColumn(ft.Text(h, size=10, weight=ft.FontWeight.BOLD, color=ACCENT_BLUE))
            for h in col_labels
        ]
        preview_table.rows = []

        for _, row in df.head(100).iterrows():
            cells = []
            for col in cols_to_show:
                val = row.get(col, "")
                if col in ("checkin", "checkout"):
                    try:
                        val = val.strftime("%d/%m/%Y")
                    except Exception:
                        val = "—"
                elif col in ("tarifa_bruta", "comision", "tarifa_neta"):
                    try:
                        val = f"${float(val):,.2f}"
                    except Exception:
                        val = "—"
                if str(val) in ("nan", "NaT", "None", ""):
                    val = "—"
                plat = str(row.get("plataforma", "")).lower()
                txt_color = (PLATFORM_COLORS.get(plat, TEXT_PRIMARY)
                             if col == "plataforma" else TEXT_PRIMARY)
                cells.append(ft.DataCell(ft.Text(str(val), size=9, color=txt_color)))
            preview_table.rows.append(ft.DataRow(cells=cells))

        preview_label.value = (
            f"Vista previa — {min(len(df), 100)} de {len(df)} reservas (orden por Check-In)"
        )

    # ── FilePicker handlers — async en Flet 0.85 ─────────────────────────
    # pick_files() retorna List[FilePickerFile] | None directamente (no callback)

    pickers: dict[str, ft.FilePicker] = {}
    for platform in ["booking", "airbnb", "expedia"]:
        pickers[platform] = ft.FilePicker()
        page.services.append(pickers[platform])

    def make_pick_handler(platform: str):
        """Genera un handler async para cada botón de plataforma."""
        async def handler(_):
            files = await pickers[platform].pick_files(
                allowed_extensions=["csv", "xlsx", "xls"],
                dialog_title=f"Seleccionar reporte de {platform.capitalize()}",
            )
            if not files:
                return
            filepath = files[0].path
            loaded_files[platform] = filepath
            fname = Path(filepath).name

            file_status[platform].value = f"✓ {fname}"
            file_status[platform].color = ACCENT_GREEN
            file_badges[platform].visible = True
            file_badges[platform].content = badge_pill(
                platform.capitalize(), PLATFORM_COLORS[platform]
            )
            process_btn.disabled = False
            set_status(f"Archivo {platform} cargado: {fname}", "info")
            page.update()
        return handler

    def on_process(_):
        """Lanza el pipeline ETL en un thread separado para no bloquear la UI."""
        process_btn.disabled    = True
        process_spinner.visible = True
        set_status("Procesando archivos...", "working")
        set_progress(0, "Leyendo archivos...")
        page.update()

        def progress_cb(value, label):
            """Recibe notificaciones reales desde el interior del pipeline."""
            set_progress(value, label)
            page.update()

        def run():
            import time

            # Lectura de archivos: tramo inicial 0% → primeros pasos del pipeline
            n = len(loaded_files)
            for i, (platform, filepath) in enumerate(loaded_files.items()):
                set_progress(i / n * 0.05, f"Leyendo {platform.capitalize()}...")
                page.update()
                ok, msg = processor.load_file(filepath, platform)
                if not ok:
                    set_status(msg, "error")

            # El pipeline reporta desde adentro vía callback (0.05 → 0.95)
            ok, msg = processor.process_all(progress_cb=progress_cb)

            # Vista previa (95% → 100%)
            set_progress(0.95, "Construyendo vista previa...")
            page.update()
            if ok:
                build_preview_table(processor.unified_df)
                set_status(msg, "success")
                process_btn.text = "✓ Procesado — volver a ejecutar"
            else:
                set_status(msg, "error")

            set_progress(1.0, "")
            page.update()
            time.sleep(0.5)
            set_progress(-1)
            process_btn.disabled    = False
            process_spinner.visible = False
            page.update()

        page.run_thread(run)

    process_btn.on_click = on_process

    def platform_card(platform: str):
        color = PLATFORM_COLORS[platform]
        icon  = PLATFORM_ICONS[platform]
        desc  = {
            "booking": "CSV desde Extranet · Booking.com",
            "airbnb":  "CSV desde Panel Host · Airbnb",
            "expedia": "CSV desde Partner Central · Expedia",
        }[platform]
        return card(
            ft.Column([
                ft.Row([
                    ft.Icon(icon, color=color, size=20),
                    ft.Text(platform.capitalize(), size=13, color=TEXT_PRIMARY,
                            weight=ft.FontWeight.BOLD),
                    file_badges[platform],
                ], spacing=8),
                ft.Text(desc, size=9, color=TEXT_SECONDARY),
                ft.Container(height=6),
                file_status[platform],
                ft.Container(height=8),
                ft.ElevatedButton(
                    f"Seleccionar {platform.capitalize()}",
                    icon=ft.Icons.UPLOAD_FILE,
                    bgcolor=ft.Colors.with_opacity(0.15, color),
                    color=color,
                    style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=6)),
                    height=34,
                    # on_click recibe una coroutine — Flet 0.85 la ejecuta automáticamente
                    on_click=make_pick_handler(platform),
                ),
            ], spacing=4),
            padding=14,
        )

    load_panel = ft.Column(
        [
            ft.Text("Carga de Reportes", size=16, color=TEXT_PRIMARY,
                    weight=ft.FontWeight.BOLD),
            lbl("Seleccione los archivos exportados de cada plataforma"),
            ft.Container(height=8),
            ft.Row([platform_card(p) for p in ["booking", "airbnb", "expedia"]], spacing=12),
            ft.Container(height=12),
            ft.Row([process_btn, process_spinner], spacing=12,
                   vertical_alignment=ft.CrossAxisAlignment.CENTER),
            ft.Container(height=6),
            lbl("Nota: si el reporte no incluye comisión → Booking 15% | Airbnb 3% | Expedia 18%",
                size=9, italic=True),
            hdivider(),
            preview_label,
            ft.Container(height=6),
            ft.Container(
                content=ft.Column([preview_table], scroll=ft.ScrollMode.AUTO, expand=True),
                expand=True,
                border=ft.Border.all(1, BORDER_COLOR),
                border_radius=8,
                bgcolor=BG_PANEL,
            ),
        ],
        expand=True,
        scroll=ft.ScrollMode.AUTO,
        spacing=2,
    )

    # ─────────────────────────────────────────────────────────────────────
    # PANEL 2 — DASHBOARD
    # ─────────────────────────────────────────────────────────────────────
    kpi_row = ft.Row([], spacing=10, wrap=True)

    img_monthly = ft.Image(src="", width=620, height=240,
                           fit=ft.BoxFit.CONTAIN, visible=False)
    img_pie     = ft.Image(src="", width=320, height=240,
                           fit=ft.BoxFit.CONTAIN, visible=False)
    img_trend   = ft.Image(src="", width=950, height=200,
                           fit=ft.BoxFit.CONTAIN, visible=False)

    dash_spinner = ft.ProgressRing(width=32, height=32, stroke_width=3,
                                   color=ACCENT_BLUE, visible=False)
    dash_msg     = ft.Text("Procese los archivos para ver el dashboard",
                           size=12, color=TEXT_SECONDARY, italic=True)

    def kpi_card(title, value, subtitle, color=ACCENT_BLUE, icon=ft.Icons.ANALYTICS):
        return card(
            ft.Column([
                ft.Row([ft.Icon(icon, color=color, size=16),
                        lbl(title, size=9, color=TEXT_SECONDARY)], spacing=6),
                ft.Text(str(value), size=20, color=color, weight=ft.FontWeight.BOLD),
                lbl(subtitle, size=9),
            ], spacing=3),
            padding=12,
            width=185,
        )

    def load_dashboard():
        if processor.unified_df.empty:
            dash_msg.value = "No hay datos procesados aún."
            page.update()
            return

        dash_spinner.visible = True
        img_monthly.visible  = False
        img_pie.visible      = False
        img_trend.visible    = False
        dash_msg.value       = ""
        set_progress(0, "Calculando KPIs...")
        page.update()

        def run():
            import time

            # KPIs — rápido, cálculos en memoria
            set_progress(0.05, "Calculando KPIs...")
            page.update()
            stats   = processor.get_summary_stats()
            monthly = processor.get_monthly_occupation()
            channel = processor.get_channel_distribution()

            kpi_row.controls = [
                kpi_card("RESERVAS TOTALES",
                         int(stats.get("total_reservas", 0)),
                         "Todas las plataformas", ACCENT_BLUE, ft.Icons.BOOK_ONLINE),
                kpi_card("INGRESO BRUTO",
                         f"${stats.get('total_bruto', 0):,.0f}",
                         "Antes de comisiones", ACCENT_ORANGE, ft.Icons.PAYMENTS),
                kpi_card("INGRESO NETO",
                         f"${stats.get('total_neto', 0):,.0f}",
                         "Después de comisiones", ACCENT_GREEN, ft.Icons.TRENDING_UP),
                kpi_card("COMISIONES",
                         f"${stats.get('total_comisiones', 0):,.0f}",
                         "Pagadas a OTAs", ACCENT_RED, ft.Icons.MONEY_OFF),
                kpi_card("PROM. NOCHES",
                         f"{stats.get('promedio_noches', 0):.1f}",
                         "Por reserva", ACCENT_BLUE, ft.Icons.NIGHTS_STAY),
            ]
            page.update()

            # Gráfico 1 — el más pesado, Matplotlib renderiza la imagen
            set_progress(0.20, "Renderizando gráfico de ocupación mensual...")
            page.update()
            img_monthly.src = b64_to_src(build_monthly_chart(monthly))
            set_progress(0.55, "✓ Ocupación mensual lista")
            page.update()

            # Gráfico 2 — torta de canales
            set_progress(0.58, "Renderizando gráfico de canales...")
            page.update()
            img_pie.src = b64_to_src(build_channel_pie(channel))
            set_progress(0.78, "✓ Distribución por canal lista")
            page.update()

            # Gráfico 3 — tendencia de ingresos
            set_progress(0.80, "Renderizando tendencia de ingresos...")
            page.update()
            img_trend.src = b64_to_src(build_revenue_trend(monthly))
            set_progress(1.0, "")
            page.update()

            time.sleep(0.5)
            set_progress(-1)
            dash_spinner.visible = False
            img_monthly.visible  = True
            img_pie.visible      = True
            img_trend.visible    = True
            page.update()

        page.run_thread(run)

    dashboard_panel = ft.Column(
        [
            ft.Row([
                ft.Text("Dashboard de Control", size=16, color=TEXT_PRIMARY,
                        weight=ft.FontWeight.BOLD),
                ft.IconButton(ft.Icons.REFRESH, icon_color=ACCENT_BLUE,
                              tooltip="Actualizar gráficos",
                              on_click=lambda _: load_dashboard()),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            lbl("KPIs y visualizaciones del inventario consolidado"),
            ft.Container(height=8),
            kpi_row,
            ft.Container(height=8),
            dash_spinner,
            dash_msg,
            ft.Row([
                card(ft.Column([lbl("Ocupación Mensual por Canal", bold=True),
                                ft.Container(height=4), img_monthly]), width=650),
                card(ft.Column([lbl("Origen de Reservas", bold=True),
                                ft.Container(height=4), img_pie]),     width=330),
            ], spacing=12),
            ft.Container(height=8),
            card(ft.Column([lbl("Tendencia de Ingresos Netos", bold=True),
                            ft.Container(height=4), img_trend])),
        ],
        expand=True,
        scroll=ft.ScrollMode.AUTO,
        spacing=4,
    )

    # ─────────────────────────────────────────────────────────────────────
    # PANEL 3 — EXPORTAR
    # ─────────────────────────────────────────────────────────────────────
    export_path = ft.TextField(
        value=str(Path.home() / "Reporte_Consolidado_Hotel.xlsx"),
        bgcolor=BG_PANEL,
        border_color=BORDER_COLOR,
        focused_border_color=ACCENT_BLUE,
        color=TEXT_PRIMARY,
        text_size=11,
        height=40,
        expand=True,
    )
    export_log     = ft.Column([], scroll=ft.ScrollMode.AUTO, spacing=3)
    export_spinner = ft.ProgressRing(width=18, height=18, stroke_width=2,
                                     color=ACCENT_GREEN, visible=False)

    save_picker = ft.FilePicker()
    page.services.append(save_picker)

    def add_log(msg: str, color: str = TEXT_SECONDARY):
        ts = datetime.now().strftime("%H:%M:%S")
        export_log.controls.append(
            ft.Text(f"[{ts}] {msg}", size=10, color=color, selectable=True)
        )
        if len(export_log.controls) > 60:
            export_log.controls.pop(0)
        page.update()

    def on_export(_):
        if processor.unified_df.empty:
            add_log("⚠ No hay datos procesados. Procese los archivos primero.", ACCENT_ORANGE)
            return
        out = export_path.value.strip()
        if not out.endswith(".xlsx"):
            out += ".xlsx"
        export_spinner.visible = True
        set_progress(0, "Preparando datos para exportar...")
        page.update()

        def run():
            import time
            add_log(f"Generando Excel: {out}", ACCENT_BLUE)

            # Hoja 1 — Resumen Ejecutivo
            set_progress(0.10, "Construyendo Hoja 1: Resumen Ejecutivo...")
            page.update()
            time.sleep(0.05)

            # Hoja 2 — Detalle de Reservas
            set_progress(0.35, "Construyendo Hoja 2: Detalle de Reservas...")
            page.update()
            time.sleep(0.05)

            # Hoja 3 — Resumen por Plataforma
            set_progress(0.65, "Construyendo Hoja 3: Resumen por Plataforma...")
            page.update()
            time.sleep(0.05)

            # Escritura en disco
            set_progress(0.85, "Guardando archivo en disco...")
            page.update()
            ok, msg = export_to_excel(processor.unified_df, out)

            set_progress(1.0, "")
            page.update()
            add_log(msg, ACCENT_GREEN if ok else ACCENT_RED)
            if ok:
                add_log(f"📊 {len(processor.unified_df)} reservas exportadas en 3 hojas.",
                        TEXT_PRIMARY)
                for err in processor.error_log:
                    add_log(err, ACCENT_ORANGE)

            time.sleep(0.4)
            set_progress(-1)
            export_spinner.visible = False
            page.update()

        page.run_thread(run)

    # save_file también es async en Flet 0.85
    async def on_pick_save_path(_):
        path = await save_picker.save_file(
            file_name="Reporte_Consolidado_Hotel.xlsx",
            allowed_extensions=["xlsx"],
            dialog_title="Guardar reporte como...",
        )
        if path:
            export_path.value = path if path.endswith(".xlsx") else path + ".xlsx"
            page.update()

    export_panel = ft.Column(
        [
            ft.Text("Exportar Reporte Consolidado", size=16, color=TEXT_PRIMARY,
                    weight=ft.FontWeight.BOLD),
            lbl("Genera un Excel profesional con todas las reservas unificadas"),
            ft.Container(height=12),
            card(ft.Column([
                lbl("Contenido del Excel generado:", bold=True),
                ft.Container(height=6),
                ft.Column([
                    ft.Row([ft.Icon(ft.Icons.CHECK_CIRCLE, color=ACCENT_GREEN, size=14),
                            lbl("Hoja 1 · Resumen Ejecutivo — KPIs con fórmulas Excel dinámicas")]),
                    ft.Row([ft.Icon(ft.Icons.CHECK_CIRCLE, color=ACCENT_GREEN, size=14),
                            lbl("Hoja 2 · Detalle completo — filtros activos, color por plataforma")]),
                    ft.Row([ft.Icon(ft.Icons.CHECK_CIRCLE, color=ACCENT_GREEN, size=14),
                            lbl("Hoja 3 · Resumen por canal — totales automáticos")]),
                ], spacing=7),
            ])),
            ft.Container(height=12),
            lbl("Ruta de destino:", bold=True),
            ft.Container(height=4),
            ft.Row([
                export_path,
                ft.IconButton(
                    ft.Icons.FOLDER_OPEN, icon_color=ACCENT_BLUE,
                    tooltip="Seleccionar ubicación",
                    on_click=on_pick_save_path,   # async handler directo
                ),
            ], spacing=8),
            ft.Container(height=12),
            ft.Row([
                ft.ElevatedButton(
                    "📥 Exportar a Excel",
                    bgcolor=ACCENT_GREEN, color="white",
                    style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8)),
                    height=42, width=200,
                    on_click=on_export,
                ),
                export_spinner,
            ], spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            ft.Container(height=16),
            hdivider(),
            lbl("Log de operaciones:", bold=True),
            ft.Container(height=4),
            ft.Container(
                content=export_log,
                bgcolor=BG_PANEL,
                border=ft.Border.all(1, BORDER_COLOR),
                border_radius=8,
                padding=ft.Padding(left=10, top=10, right=10, bottom=10),
                height=200,
            ),
        ],
        expand=True,
        scroll=ft.ScrollMode.AUTO,
        spacing=4,
    )

    # ─────────────────────────────────────────────────────────────────────
    # NAVEGACIÓN LATERAL
    # ─────────────────────────────────────────────────────────────────────
    panels     = {"load": load_panel, "dashboard": dashboard_panel, "export": export_panel}
    panel_keys = ["load", "dashboard", "export"]

    content_area = ft.Container(
        content=load_panel,
        expand=True,
        padding=ft.Padding(left=20, top=20, right=20, bottom=20),
        bgcolor=BG_DARK,
    )

    def nav_change(e):
        key = panel_keys[e.control.selected_index]
        content_area.content = panels[key]
        if key == "dashboard" and not processor.unified_df.empty:
            load_dashboard()
        page.update()

    rail = ft.NavigationRail(
        selected_index=0,
        label_type=ft.NavigationRailLabelType.ALL,
        min_width=72,
        bgcolor=BG_PANEL,
        indicator_color=ft.Colors.with_opacity(0.15, ACCENT_BLUE),
        destinations=[
            ft.NavigationRailDestination(
                icon=ft.Icons.UPLOAD_FILE_OUTLINED,
                selected_icon=ft.Icons.UPLOAD_FILE,
                label="Carga",
            ),
            ft.NavigationRailDestination(
                icon=ft.Icons.DASHBOARD_OUTLINED,
                selected_icon=ft.Icons.DASHBOARD,
                label="Dashboard",
            ),
            ft.NavigationRailDestination(
                icon=ft.Icons.DOWNLOAD_OUTLINED,
                selected_icon=ft.Icons.DOWNLOAD,
                label="Exportar",
            ),
        ],
        on_change=nav_change,
    )

    # ─────────────────────────────────────────────────────────────────────
    # APP BAR
    # ─────────────────────────────────────────────────────────────────────
    page.appbar = ft.AppBar(
        leading=ft.Icon(ft.Icons.HOTEL, color=ACCENT_BLUE, size=22),
        leading_width=48,
        title=ft.Row([
            ft.Text("Hotel Channel Manager", size=14,
                    weight=ft.FontWeight.BOLD, color=TEXT_PRIMARY),
            ft.Text("— Data Link", size=12, color=TEXT_SECONDARY),
        ], spacing=4),
        center_title=False,
        bgcolor=BG_PANEL,
        actions=[
            ft.Container(
                ft.Row([status_dot, status_text], spacing=6),
                padding=ft.Padding(left=0, top=0, right=16, bottom=0),
            ),
        ],
    )

    # ─────────────────────────────────────────────────────────────────────
    # LAYOUT RAÍZ
    # ─────────────────────────────────────────────────────────────────────
    page.add(
        ft.Column([
            # Barra de progreso global — visible durante operaciones largas
            progress_label_row,
            progress_bar,
            ft.Row([
                rail,
                ft.VerticalDivider(width=1, color=BORDER_COLOR),
                content_area,
            ], expand=True, spacing=0),
        ], spacing=0, expand=True)
    )

    set_status("Listo — cargue los archivos de cada plataforma", "info")
    page.update()


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    ft.run(main)
