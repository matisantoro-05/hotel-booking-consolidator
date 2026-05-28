"""
sample_data/generate_samples.py
================================
Genera archivos CSV de muestra para cada plataforma con datos realistas
de un hotel ficticio: "Hotel Mirador del Sur".
Ejecutar una vez para tener archivos de prueba.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
import os

random.seed(42)
np.random.seed(42)

GUESTS_ES = [
    "García López, Carlos",    "Martínez Ruiz, Ana",     "González Fernández, Pedro",
    "Rodríguez Sánchez, María","López Jiménez, José",    "Pérez Torres, Laura",
    "Sánchez Moreno, Antonio", "Ramírez Castro, Isabel", "Flores Ortega, Diego",
    "Vargas Medina, Sofía",
]
GUESTS_EN = [
    "James Wilson",     "Emma Thompson",   "Oliver Brown",    "Sophie Johnson",
    "William Davis",    "Charlotte White",  "Benjamin Moore",  "Amelia Taylor",
    "Henry Anderson",   "Grace Jackson",
]
ROOMS_BOOKING = ["Habitación Doble Superior", "Suite Junior", "Habitación Individual", "Suite Familiar"]
ROOMS_AIRBNB  = ["Entire apartment · Mirador", "Private room · Garden View", "Entire villa · Rooftop"]
ROOMS_EXPEDIA = ["Standard Double", "Junior Suite", "Deluxe King", "Family Room"]
STATUSES_ES = ["Confirmada", "Cancelada", "No-Show", "Confirmada", "Confirmada"]  # Más confirmadas


def random_dates(n: int, year: int = 2024):
    """Genera pares checkin/checkout aleatorios dentro del año dado."""
    dates = []
    for _ in range(n):
        month = random.randint(1, 12)
        day   = random.randint(1, 25)
        try:
            checkin  = datetime(year, month, day)
        except ValueError:
            checkin  = datetime(year, month, 1)
        nights   = random.randint(1, 7)
        checkout = checkin + timedelta(days=nights)
        dates.append((checkin, checkout, nights))
    return dates


def generate_booking_csv(filepath: str, n: int = 30):
    """Formato Booking.com: fechas DD/MM/YYYY, columnas en español."""
    dates = random_dates(n)
    records = []
    for i, (ci, co, nights) in enumerate(dates):
        bruto = round(random.uniform(80, 450) * nights, 2)
        commission = round(bruto * 0.15, 2)
        records.append({
            "Número de reserva":   f"BK{2024000 + i}",
            "Nombre del huésped":  random.choice(GUESTS_ES),
            "Fecha de llegada":    ci.strftime("%d/%m/%Y"),    # DD/MM/YYYY
            "Fecha de salida":     co.strftime("%d/%m/%Y"),
            "Habitación":          random.choice(ROOMS_BOOKING),
            "Noches":              nights,
            "Adultos":             random.randint(1, 3),
            "Precio total":        f"${bruto}",
            "Comisión":            f"${commission}",
            "Estado":              random.choice(STATUSES_ES),
        })
    pd.DataFrame(records).to_csv(filepath, index=False, encoding="utf-8-sig")
    print(f"✓ {filepath} generado ({n} reservas)")


def generate_airbnb_csv(filepath: str, n: int = 20):
    """Formato Airbnb: fechas YYYY-MM-DD, columnas en inglés."""
    dates = random_dates(n)
    records = []
    for i, (ci, co, nights) in enumerate(dates):
        payout = round(random.uniform(60, 380) * nights, 2)
        records.append({
            "Confirmation Code":  f"AIR{8000 + i}",
            "Guest Name":         random.choice(GUESTS_EN),
            "Start Date":         ci.strftime("%Y-%m-%d"),     # ISO YYYY-MM-DD
            "End Date":           co.strftime("%Y-%m-%d"),
            "Listing":            random.choice(ROOMS_AIRBNB),
            "Nights":             nights,
            "# Guests":           random.randint(1, 4),
            "Payout":             payout,
            # Airbnb no reporta comisión explícita → se calcula al procesar
            "Status":             random.choice(["confirmed", "cancelled", "confirmed", "confirmed"]),
        })
    pd.DataFrame(records).to_csv(filepath, index=False, encoding="utf-8")
    print(f"✓ {filepath} generado ({n} reservas)")


def generate_expedia_csv(filepath: str, n: int = 25):
    """Formato Expedia: fechas MM/DD/YYYY, columnas mixtas."""
    dates = random_dates(n)
    records = []
    for i, (ci, co, nights) in enumerate(dates):
        rate = round(random.uniform(75, 420) * nights, 2)
        commission = round(rate * 0.18, 2)
        records.append({
            "ID de reserva":       f"EXP{5000 + i}",
            "Cliente":             random.choice(GUESTS_ES + GUESTS_EN),
            "Fecha entrada":       ci.strftime("%m/%d/%Y"),    # MM/DD/YYYY americano
            "Fecha salida":        co.strftime("%m/%d/%Y"),
            "Tipo de habitación":  random.choice(ROOMS_EXPEDIA),
            "Duración":            nights,
            "Adultos":             random.randint(1, 3),
            "Tarifa":              rate,
            "Comisión Expedia":    commission,
            "Estado":              random.choice(["Confirmada", "Cancelada", "Confirmada", "Confirmada"]),
        })
    pd.DataFrame(records).to_csv(filepath, index=False, encoding="utf-8-sig")
    print(f"✓ {filepath} generado ({n} reservas)")


if __name__ == "__main__":
    out_dir = os.path.dirname(os.path.abspath(__file__))
    generate_booking_csv(os.path.join(out_dir, "sample_booking.csv"))
    generate_airbnb_csv(os.path.join(out_dir, "sample_airbnb.csv"))
    generate_expedia_csv(os.path.join(out_dir, "sample_expedia.csv"))
    print("\n✅ Todos los archivos de muestra generados en sample_data/")
