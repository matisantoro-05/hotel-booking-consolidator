"""
core/data_processor.py
======================
Motor de procesamiento de datos para el Hotel Channel Manager.
Responsable de: ingestión, normalización, mapeo de columnas,
cálculos financieros y deduplicación de reservas.
"""

import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path


# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE PLATAFORMAS
# Define los esquemas de cada OTA: columnas esperadas, comisión y formato de fecha.
# Al agregar una nueva plataforma, solo hay que extender este diccionario.
# ─────────────────────────────────────────────────────────────────────────────
PLATFORM_CONFIG = {
    "booking": {
        "commission_rate": 0.15,          # 15% comisión estándar Booking.com
        "date_format": "%d/%m/%Y",        # Formato europeo DD/MM/YYYY
        "column_map": {
            # clave = columna del CSV original → valor = nombre estándar interno
            "Nombre del huésped":         "huesped",
            "Nombre":                     "huesped",
            "Número de reserva":          "id_reserva",
            "Booking Number":             "id_reserva",
            "Fecha de llegada":           "checkin",
            "Check-in Date":              "checkin",
            "Fecha de salida":            "checkout",
            "Check-out Date":             "checkout",
            "Habitación":                 "habitacion",
            "Room":                       "habitacion",
            "Precio total":               "tarifa_bruta",
            "Total Price":                "tarifa_bruta",
            "Comisión":                   "comision_explicita",
            "Commission":                 "comision_explicita",
            "Noches":                     "noches",
            "Nights":                     "noches",
            "Adultos":                    "adultos",
            "Estado":                     "estado",
            "Status":                     "estado",
        },
    },
    "airbnb": {
        "commission_rate": 0.03,          # 3% comisión host Airbnb
        "date_format": "%Y-%m-%d",        # Formato ISO YYYY-MM-DD
        "column_map": {
            "Guest Name":                 "huesped",
            "Guest":                      "huesped",
            "Confirmation Code":          "id_reserva",
            "Confirmation":               "id_reserva",
            "Start Date":                 "checkin",
            "End Date":                   "checkout",
            "Listing":                    "habitacion",
            "Amount":                     "tarifa_bruta",
            "Payout":                     "tarifa_bruta",
            "Host Fee":                   "comision_explicita",
            "Nights":                     "noches",
            "# Guests":                   "adultos",
            "Status":                     "estado",
            "Type":                       "estado",
        },
    },
    "expedia": {
        "commission_rate": 0.18,          # 18% comisión estándar Expedia
        "date_format": "%m/%d/%Y",        # Formato americano MM/DD/YYYY
        "column_map": {
            "Cliente":                    "huesped",
            "Guest Name":                 "huesped",
            "Traveler Name":              "huesped",
            "ID de reserva":              "id_reserva",
            "Reservation ID":             "id_reserva",
            "Booking ID":                 "id_reserva",
            "Fecha entrada":              "checkin",
            "Arrival Date":               "checkin",
            "Check In":                   "checkin",
            "Fecha salida":               "checkout",
            "Departure Date":             "checkout",
            "Check Out":                  "checkout",
            "Tipo de habitación":         "habitacion",
            "Room Type":                  "habitacion",
            "Tarifa":                     "tarifa_bruta",
            "Total Amount":               "tarifa_bruta",
            "Rate":                       "tarifa_bruta",
            "Comisión Expedia":           "comision_explicita",
            "Commission":                 "comision_explicita",
            "Duración":                   "noches",
            "Length of Stay":             "noches",
            "Estado":                     "estado",
            "Booking Status":             "estado",
        },
    },
}

# Columnas del schema unificado (siempre presentes en el DataFrame final)
UNIFIED_COLUMNS = [
    "id_reserva", "plataforma", "huesped", "checkin", "checkout",
    "noches", "habitacion", "adultos", "estado",
    "tarifa_bruta", "comision", "tarifa_neta",
]


# ─────────────────────────────────────────────────────────────────────────────
# CLASE PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────
class HotelDataProcessor:
    """
    Centraliza toda la lógica de ETL (Extract-Transform-Load) para
    los reportes de Booking.com, Airbnb y Expedia.
    """

    def __init__(self):
        self.raw_dataframes: dict[str, pd.DataFrame] = {}   # DataFrames originales por plataforma
        self.unified_df: pd.DataFrame = pd.DataFrame()      # DataFrame consolidado final
        self.error_log: list[str] = []                      # Registro de errores no críticos

    # ── INGESTIÓN ─────────────────────────────────────────────────────────────
    def load_file(self, filepath: str, platform: str) -> tuple[bool, str]:
        """
        Carga un archivo CSV/XLSX de una plataforma específica.
        Retorna (éxito: bool, mensaje: str).
        """
        path = Path(filepath)
        if not path.exists():
            return False, f"Archivo no encontrado: {filepath}"

        try:
            if path.suffix.lower() in (".xlsx", ".xls"):
                df = pd.read_excel(filepath, dtype=str)
            elif path.suffix.lower() == ".csv":
                # Intentar detectar el separador automáticamente
                df = pd.read_csv(filepath, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
            else:
                return False, f"Formato no soportado: {path.suffix}"

            self.raw_dataframes[platform] = df
            return True, f"✓ {path.name} cargado ({len(df)} filas)"

        except Exception as e:
            return False, f"Error al leer {path.name}: {str(e)}"

    # ── TRANSFORMACIÓN ────────────────────────────────────────────────────────
    def _map_columns(self, df: pd.DataFrame, platform: str) -> pd.DataFrame:
        """
        Renombra las columnas del DataFrame original al schema unificado.
        Usa el 'column_map' de PLATFORM_CONFIG para encontrar equivalencias.
        Las columnas sin mapeo se descartan; las faltantes se rellenan con NaN.
        """
        col_map = PLATFORM_CONFIG[platform]["column_map"]

        # Construir diccionario de renombrado solo con columnas que existen en el DF
        rename_dict = {
            original: unified
            for original, unified in col_map.items()
            if original in df.columns
        }

        df = df.rename(columns=rename_dict)

        # Asegurar que todas las columnas del schema existen (rellenar con NaN si faltan)
        for col in UNIFIED_COLUMNS:
            if col not in df.columns:
                df[col] = np.nan

        return df[UNIFIED_COLUMNS]

    def _normalize_dates(self, df: pd.DataFrame, platform: str) -> pd.DataFrame:
        """
        Convierte las fechas de checkin/checkout al tipo datetime de Pandas.
        - Si la columna ya es datetime (re-ejecución del pipeline), no hace nada.
        - Primero intenta el formato declarado en PLATFORM_CONFIG.
        - Si falla, usa pd.to_datetime con inferencia automática como fallback.
        """
        fmt = PLATFORM_CONFIG[platform]["date_format"]
        dayfirst = fmt.startswith("%d")   # True para formatos europeos DD/MM/...

        for col in ["checkin", "checkout"]:
            if col not in df.columns or df[col].isna().all():
                continue

            # Si ya es datetime (pipeline re-ejecutado), no tocar
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                continue

            # Convertir a string limpio antes de parsear (elimina espacios, etc.)
            col_str = df[col].astype(str).str.strip()

            # Intento 1: formato exacto declarado en PLATFORM_CONFIG
            parsed = pd.to_datetime(col_str, format=fmt, errors="coerce")

            # Intento 2: inferencia automática para las filas que fallaron
            if parsed.isna().any():
                fallback = pd.to_datetime(col_str, dayfirst=dayfirst, errors="coerce")
                # Usar fallback solo donde el intento exacto produjo NaT
                parsed = parsed.fillna(fallback)

            df[col] = parsed

        return df

    def _calculate_financials(self, df: pd.DataFrame, platform: str) -> pd.DataFrame:
        """
        Lógica de negocio financiera:
        1. Convierte tarifa_bruta a numérico (limpia símbolo $, comas, etc.)
        2. Si hay comisión explícita en el reporte → la usa directamente.
        3. Si NO hay comisión explícita → la calcula aplicando el % del PLATFORM_CONFIG.
        4. Calcula tarifa_neta = tarifa_bruta - comision.
        """
        commission_rate = PLATFORM_CONFIG[platform]["commission_rate"]

        # --- Limpiar y convertir tarifa_bruta ---
        df["tarifa_bruta"] = (
            df["tarifa_bruta"]
            .astype(str)
            .str.replace(r"[€$£,\s]", "", regex=True)   # Eliminar símbolos de moneda
            .str.replace(",", ".", regex=False)           # Normalizar decimales
            .replace("", np.nan)
            .pipe(pd.to_numeric, errors="coerce")
        )

        # --- Comisión: explícita vs calculada ---
        if "comision_explicita" in df.columns and df["comision_explicita"].notna().any():
            # Limpiar y convertir la comisión declarada en el reporte
            df["comision_explicita"] = (
                df["comision_explicita"]
                .astype(str)
                .str.replace(r"[€$£,\s%]", "", regex=True)
                .replace("", np.nan)
                .pipe(pd.to_numeric, errors="coerce")
            )
            # Si el valor es < 1, asumimos que es porcentaje (ej: 0.15 en lugar de 15%)
            mask_pct = df["comision_explicita"] < 1
            df.loc[mask_pct, "comision_explicita"] = (
                df.loc[mask_pct, "comision_explicita"] * df.loc[mask_pct, "tarifa_bruta"]
            )
            df["comision"] = df["comision_explicita"].fillna(
                df["tarifa_bruta"] * commission_rate
            )
        else:
            # No hay comisión en el reporte → calcular según tarifa de la plataforma
            df["comision"] = df["tarifa_bruta"] * commission_rate

        # --- Tarifa neta real ---
        df["tarifa_neta"] = df["tarifa_bruta"] - df["comision"]

        # Redondear a 2 decimales
        df[["tarifa_bruta", "comision", "tarifa_neta"]] = df[
            ["tarifa_bruta", "comision", "tarifa_neta"]
        ].round(2)

        return df

    def _clean_metadata(self, df: pd.DataFrame, platform: str) -> pd.DataFrame:
        """
        Limpieza general: agrega columna 'plataforma', normaliza texto,
        calcula noches si están ausentes, y estandariza el campo estado.
        """
        df["plataforma"] = platform.capitalize()

        # Normalizar nombre de huésped (capitalize words, strip espacios)
        if "huesped" in df.columns:
            df["huesped"] = df["huesped"].astype(str).str.strip().str.title().replace("Nan", "")

        # Calcular noches a partir de fechas si la columna está vacía
        if df["noches"].isna().all() or (df["noches"].astype(str).str.strip() == "").all():
            mask = df["checkin"].notna() & df["checkout"].notna()
            df.loc[mask, "noches"] = (
                df.loc[mask, "checkout"] - df.loc[mask, "checkin"]
            ).dt.days

        # Convertir noches a entero
        df["noches"] = pd.to_numeric(df["noches"], errors="coerce").fillna(0).astype(int)

        # Adultos: rellenar con 1 si está ausente
        df["adultos"] = pd.to_numeric(df["adultos"], errors="coerce").fillna(1).astype(int)

        # Estandarizar estado (texto limpio)
        if "estado" in df.columns:
            df["estado"] = df["estado"].astype(str).str.strip().str.capitalize()

        return df

    def _deduplicate(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Deduplicación inteligente para evitar doble conteo cuando se cargan
        archivos con solapamiento de fechas.

        Estrategia:
        1. Prioridad por id_reserva (si existe en ambos registros, conservar el primero cargado).
        2. Fallback: detectar reservas "gemelas" por huésped + checkin + checkout + tarifa_bruta.
           En ese caso conservar el registro con mayor información (menos NaN).
        """
        original_count = len(df)

        # --- Paso 1: deduplicar por ID de reserva (excluyendo IDs vacíos) ---
        mask_has_id = df["id_reserva"].astype(str).str.strip().ne("") & df["id_reserva"].notna()
        df_with_id = df[mask_has_id].drop_duplicates(subset=["id_reserva"], keep="first")
        df_no_id = df[~mask_has_id]

        # --- Paso 2: deduplicar los sin ID por huésped + fechas + tarifa ---
        dedup_keys = ["huesped", "checkin", "checkout", "tarifa_bruta"]
        df_no_id = df_no_id.drop_duplicates(subset=dedup_keys, keep="first")

        result = pd.concat([df_with_id, df_no_id], ignore_index=True)
        removed = original_count - len(result)

        if removed > 0:
            self.error_log.append(f"⚠ Se eliminaron {removed} reserva(s) duplicada(s).")

        return result

    # ── PIPELINE PRINCIPAL ───────────────────────────────────────────────────
    def process_all(self, progress_cb=None) -> tuple[bool, str]:
        """
        Ejecuta el pipeline ETL completo sobre todos los archivos cargados.
        progress_cb(value: float, label: str) se llama en cada paso real
        para que la UI pueda actualizar la barra de progreso desde adentro.
        Retorna (éxito: bool, mensaje: str).
        """
        def report(value, label):
            if progress_cb:
                progress_cb(value, label)

        if not self.raw_dataframes:
            return False, "No hay archivos cargados. Cargue al menos un archivo."

        # Resetear log en cada ejecución para no acumular errores de corridas anteriores
        self.error_log = []

        platforms = list(self.raw_dataframes.items())
        n = len(platforms)
        processed_frames = []

        # Cada plataforma ocupa un tramo proporcional de 0% → 70%
        # Dentro de cada plataforma: 4 sub-pasos de igual peso
        for idx, (platform, df) in enumerate(platforms):
            base = (idx / n) * 0.70        # inicio del tramo de esta plataforma
            step = (1 / n) * 0.70 / 4     # cada sub-paso vale 1/4 del tramo

            try:
                report(base,            f"[{platform.capitalize()}] Mapeando columnas...")
                df = self._map_columns(df.copy(), platform)

                report(base + step,     f"[{platform.capitalize()}] Normalizando fechas...")
                df = self._normalize_dates(df, platform)

                report(base + step*2,   f"[{platform.capitalize()}] Calculando comisiones...")
                df = self._calculate_financials(df, platform)

                report(base + step*3,   f"[{platform.capitalize()}] Limpiando metadatos...")
                df = self._clean_metadata(df, platform)

                processed_frames.append(df)
            except Exception as e:
                self.error_log.append(f"Error procesando {platform}: {str(e)}")

        if not processed_frames:
            return False, "Ningún archivo pudo procesarse correctamente."

        report(0.70, "Concatenando plataformas...")
        combined = pd.concat(processed_frames, ignore_index=True)

        report(0.82, "Deduplicando reservas...")
        combined = self._deduplicate(combined)

        report(0.92, "Ordenando por fecha de check-in...")
        combined = combined.sort_values("checkin", ascending=True, na_position="last")
        combined = combined.reset_index(drop=True)

        self.unified_df = combined
        return True, f"✓ Procesado exitoso: {len(combined)} reservas consolidadas."

    # ── MÉTRICAS PARA DASHBOARD ──────────────────────────────────────────────
    def get_summary_stats(self) -> dict:
        """Calcula KPIs generales para los cards del dashboard."""
        if self.unified_df.empty:
            return {}

        df = self.unified_df
        return {
            "total_reservas":   len(df),
            "total_bruto":      df["tarifa_bruta"].sum(),
            "total_neto":       df["tarifa_neta"].sum(),
            "total_comisiones": df["comision"].sum(),
            "promedio_noches":  df["noches"].mean(),
            "plataformas":      df["plataforma"].nunique(),
        }

    def get_monthly_occupation(self) -> pd.DataFrame:
        """Agrupa reservas por mes y plataforma para el gráfico de ocupación."""
        if self.unified_df.empty:
            return pd.DataFrame()

        df = self.unified_df.copy()
        df = df.dropna(subset=["checkin"])
        df["mes"] = df["checkin"].dt.to_period("M").astype(str)

        return (
            df.groupby(["mes", "plataforma"])
            .agg(reservas=("id_reserva", "count"), ingresos=("tarifa_neta", "sum"))
            .reset_index()
        )

    def get_channel_distribution(self) -> pd.DataFrame:
        """Distribución de reservas por canal (para gráfico de torta)."""
        if self.unified_df.empty:
            return pd.DataFrame()

        return (
            self.unified_df.groupby("plataforma")
            .agg(
                reservas=("id_reserva", "count"),
                ingresos_netos=("tarifa_neta", "sum"),
            )
            .reset_index()
        )
