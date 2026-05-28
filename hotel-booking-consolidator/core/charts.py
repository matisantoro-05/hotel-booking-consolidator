"""
core/charts.py
==============
Módulo de visualización. Genera gráficos Matplotlib embebibles en Flet
como imágenes base64 (sin depender de un servidor web externo).
"""

import io
import base64
import pandas as pd
import matplotlib
matplotlib.use("Agg")          # Backend sin GUI (necesario para Flet/headless)
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.ticker import FuncFormatter
import numpy as np


# ─────────────────────────────────────────────────────────────────────────────
# PALETA Y ESTILO
# ─────────────────────────────────────────────────────────────────────────────
DARK_BG      = "#0D1117"
PANEL_BG     = "#161B22"
GRID_COLOR   = "#21262D"
TEXT_COLOR   = "#E6EDF3"
ACCENT_COLOR = "#58A6FF"

PLATFORM_COLORS = {
    "Booking":  "#003580",
    "Airbnb":   "#FF5A5F",
    "Expedia":  "#FFC72C",
    "Otros":    "#58A6FF",
}

def _apply_dark_style(fig, ax):
    """Aplica el tema oscuro consistente a cualquier figura."""
    fig.patch.set_facecolor(DARK_BG)
    ax.set_facecolor(PANEL_BG)
    ax.tick_params(colors=TEXT_COLOR, labelsize=9)
    ax.xaxis.label.set_color(TEXT_COLOR)
    ax.yaxis.label.set_color(TEXT_COLOR)
    ax.title.set_color(TEXT_COLOR)
    for spine in ax.spines.values():
        spine.set_edgecolor(GRID_COLOR)
    ax.grid(color=GRID_COLOR, linewidth=0.5, alpha=0.7)

def _fig_to_base64(fig) -> str:
    """Convierte una figura Matplotlib a string base64 para Flet Image."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=130, bbox_inches="tight",
                facecolor=DARK_BG, edgecolor="none")
    buf.seek(0)
    img_b64 = base64.b64encode(buf.read()).decode("utf-8")
    plt.close(fig)
    return img_b64


# ─────────────────────────────────────────────────────────────────────────────
# GRÁFICO 1: OCUPACIÓN MENSUAL (barras apiladas por plataforma)
# ─────────────────────────────────────────────────────────────────────────────
def build_monthly_chart(monthly_df: pd.DataFrame) -> str:
    """
    Recibe el DataFrame de get_monthly_occupation() y genera un gráfico
    de barras apiladas: eje X = mes, eje Y = cantidad de reservas,
    colores por plataforma.
    Retorna la imagen codificada en base64.
    """
    if monthly_df.empty:
        return _empty_chart("Sin datos de ocupación mensual")

    # Pivotar: filas = meses, columnas = plataformas
    pivot = monthly_df.pivot_table(
        index="mes", columns="plataforma", values="reservas", aggfunc="sum", fill_value=0
    )
    pivot = pivot.sort_index()

    fig, ax = plt.subplots(figsize=(9, 4))
    _apply_dark_style(fig, ax)

    x = np.arange(len(pivot.index))
    width = 0.6
    bottom = np.zeros(len(pivot))

    for platform in pivot.columns:
        color = PLATFORM_COLORS.get(platform, ACCENT_COLOR)
        values = pivot[platform].values
        bars = ax.bar(x, values, width, bottom=bottom, label=platform,
                      color=color, alpha=0.88, zorder=2)
        # Etiqueta dentro de la barra si hay espacio
        for bar, val in zip(bars, values):
            if val > 0:
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_y() + bar.get_height() / 2,
                    str(int(val)), ha="center", va="center",
                    color="white", fontsize=7.5, fontweight="bold"
                )
        bottom += values

    ax.set_xticks(x)
    ax.set_xticklabels(pivot.index, rotation=35, ha="right", fontsize=8)
    ax.set_ylabel("Reservas", color=TEXT_COLOR, fontsize=9)
    ax.set_title("Ocupación Mensual por Canal", color=TEXT_COLOR, fontsize=11, pad=10)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: str(int(v))))

    legend = ax.legend(
        framealpha=0, labelcolor=TEXT_COLOR, fontsize=8,
        loc="upper left", borderpad=0.5
    )

    fig.tight_layout(pad=1.5)
    return _fig_to_base64(fig)


# ─────────────────────────────────────────────────────────────────────────────
# GRÁFICO 2: ORIGEN DE RESERVAS (torta con donut interior)
# ─────────────────────────────────────────────────────────────────────────────
def build_channel_pie(channel_df: pd.DataFrame) -> str:
    """
    Recibe el DataFrame de get_channel_distribution() y genera un gráfico
    de donut mostrando el % de reservas por canal.
    Retorna la imagen codificada en base64.
    """
    if channel_df.empty:
        return _empty_chart("Sin datos de canales")

    labels  = channel_df["plataforma"].tolist()
    values  = channel_df["reservas"].tolist()
    colors  = [PLATFORM_COLORS.get(p, ACCENT_COLOR) for p in labels]

    fig, ax = plt.subplots(figsize=(5, 4))
    fig.patch.set_facecolor(DARK_BG)
    ax.set_facecolor(DARK_BG)

    wedges, texts, autotexts = ax.pie(
        values,
        labels=None,
        colors=colors,
        autopct="%1.1f%%",
        startangle=140,
        pctdistance=0.78,
        wedgeprops={"width": 0.55, "edgecolor": DARK_BG, "linewidth": 1.5},
    )

    for at in autotexts:
        at.set_color("white")
        at.set_fontsize(8.5)
        at.set_fontweight("bold")

    # Leyenda lateral
    patches = [
        mpatches.Patch(color=colors[i], label=f"{labels[i]} ({values[i]})")
        for i in range(len(labels))
    ]
    ax.legend(
        handles=patches, loc="lower center", bbox_to_anchor=(0.5, -0.12),
        ncol=len(labels), framealpha=0, labelcolor=TEXT_COLOR, fontsize=8
    )

    ax.set_title("Origen de Reservas", color=TEXT_COLOR, fontsize=11, pad=8)

    # Texto central con el total
    total = sum(values)
    ax.text(0, 0, f"{total}\nRes.", ha="center", va="center",
            color=TEXT_COLOR, fontsize=10, fontweight="bold")

    fig.tight_layout(pad=1.2)
    return _fig_to_base64(fig)


# ─────────────────────────────────────────────────────────────────────────────
# GRÁFICO 3: INGRESOS NETOS MENSUALES (línea de tendencia)
# ─────────────────────────────────────────────────────────────────────────────
def build_revenue_trend(monthly_df: pd.DataFrame) -> str:
    """
    Gráfico de línea mostrando la evolución de ingresos netos mes a mes.
    """
    if monthly_df.empty:
        return _empty_chart("Sin datos de ingresos")

    revenue_by_month = (
        monthly_df.groupby("mes")["ingresos"].sum().reset_index().sort_values("mes")
    )

    fig, ax = plt.subplots(figsize=(9, 3.5))
    _apply_dark_style(fig, ax)

    x = range(len(revenue_by_month))
    y = revenue_by_month["ingresos"].values

    ax.fill_between(x, y, alpha=0.18, color=ACCENT_COLOR)
    ax.plot(x, y, color=ACCENT_COLOR, linewidth=2, marker="o",
            markersize=5, markerfacecolor=ACCENT_COLOR, zorder=3)

    # Etiquetas de valor en cada punto
    for xi, yi in zip(x, y):
        ax.annotate(
            f"${yi:,.0f}", (xi, yi),
            textcoords="offset points", xytext=(0, 8),
            ha="center", color=TEXT_COLOR, fontsize=7.5
        )

    ax.set_xticks(list(x))
    ax.set_xticklabels(revenue_by_month["mes"].tolist(), rotation=35, ha="right", fontsize=8)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"${v:,.0f}"))
    ax.set_title("Tendencia de Ingresos Netos", color=TEXT_COLOR, fontsize=11, pad=10)
    ax.set_ylabel("Ingresos Netos ($)", color=TEXT_COLOR, fontsize=9)

    fig.tight_layout(pad=1.5)
    return _fig_to_base64(fig)


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: gráfico vacío cuando no hay datos
# ─────────────────────────────────────────────────────────────────────────────
def _empty_chart(message: str) -> str:
    fig, ax = plt.subplots(figsize=(6, 3))
    fig.patch.set_facecolor(DARK_BG)
    ax.set_facecolor(PANEL_BG)
    ax.text(0.5, 0.5, message, ha="center", va="center",
            color="#8B949E", fontsize=11, transform=ax.transAxes)
    ax.axis("off")
    return _fig_to_base64(fig)
