# 🏨 Hotel Booking Consolidator

Aplicación de escritorio profesional para unificar y analizar reportes de reservas
de Booking.com, Airbnb y Expedia en un dashboard único con exportación a Excel.

---

## 📁 Estructura del Proyecto

```
hotel_channel_manager/
│
├── main.py                        ← Punto de entrada / UI Flet
│
├── core/
│   ├── __init__.py
│   ├── data_processor.py          ← ETL: ingestión, normalización, finanzas
│   ├── charts.py                  ← Gráficos Matplotlib embebibles
│   └── excel_exporter.py          ← Exportación Excel profesional
│
├── sample_data/
│   ├── generate_samples.py        ← Genera CSVs de prueba
│   ├── sample_booking.csv
│   ├── sample_airbnb.csv
│   └── sample_expedia.csv
│
├── output/                        ← Carpeta de salida por defecto
├── requirements.txt
└── README.md
```

---

## ⚡ Instalación y Ejecución

### 1. Crear entorno virtual (recomendado)
```bash
python -m venv venv
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate
```

### 2. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 3. Ejecutar la aplicación
```bash
python main.py
```

### 4. (Opcional) Generar archivos de prueba
```bash
python sample_data/generate_samples.py
```

---

## 📊 Flujo de Trabajo

```
1. CARGA     → Seleccionar CSV/XLSX de cada plataforma
2. PROCESAR  → Click en "Procesar y Unificar Datos"
3. REVISAR   → Ver tabla de previsualización (hasta 100 filas)
4. DASHBOARD → Ir a Dashboard para ver KPIs y gráficos
5. EXPORTAR  → Generar Reporte_Consolidado_Hotel.xlsx
```

---

## 🧠 Lógica de Negocio Clave

### Formatos de fecha soportados

| Plataforma | Formato CSV   | Ejemplo       |
|------------|---------------|---------------|
| Booking    | DD/MM/YYYY    | 25/12/2024    |
| Airbnb     | YYYY-MM-DD    | 2024-12-25    |
| Expedia    | MM/DD/YYYY    | 12/25/2024    |

> Si el formato falla, se usa inferencia automática como fallback.

### Mapeo de columnas (ejemplos)

| Booking              | Airbnb        | Expedia             | → Unificado   |
|----------------------|---------------|---------------------|---------------|
| Nombre del huésped   | Guest Name    | Cliente             | `huesped`     |
| Número de reserva    | Confirmation  | ID de reserva       | `id_reserva`  |
| Precio total         | Payout        | Tarifa              | `tarifa_bruta`|

### Cálculo de comisiones

```
Si el reporte incluye comisión → se usa el valor declarado
Si NO incluye comisión         → se calcula:
  Booking:  tarifa_bruta × 15%
  Airbnb:   tarifa_bruta ×  3%
  Expedia:  tarifa_bruta × 18%

tarifa_neta = tarifa_bruta − comision
```

### Deduplicación

```
Paso 1: eliminar duplicados por id_reserva (mismo ID = misma reserva)
Paso 2: eliminar duplicados por (huésped + checkin + checkout + tarifa)
        → detecta solapamiento de archivos con fechas compartidas
```

---

## 📦 Generar Ejecutable (.exe / app)

### Windows — PyInstaller
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "HotelChannelManager" \
  --add-data "sample_data;sample_data" main.py
# El ejecutable queda en: dist/HotelChannelManager.exe
```

### macOS — PyInstaller
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "HotelChannelManager" \
  --add-data "sample_data:sample_data" main.py
# El ejecutable queda en: dist/HotelChannelManager
```

### Alternativa multiplataforma — Flet Pack
```bash
# Flet incluye su propio empaquetador basado en Flutter
flet pack main.py --name "HotelChannelManager" --product-name "Hotel Channel Manager"
```

> **Nota**: Para distribución empresarial se recomienda `flet pack` ya que
> empaqueta correctamente los assets de Flutter que usa Flet internamente.

---

## 🔧 Agregar Nueva Plataforma (extensión)

1. Abrir `core/data_processor.py`
2. Agregar entrada en `PLATFORM_CONFIG`:

```python
"nueva_ota": {
    "commission_rate": 0.12,          # 12%
    "date_format": "%Y/%m/%d",
    "column_map": {
        "Nombre Huésped": "huesped",
        "Número Booking": "id_reserva",
        # ... mapear según el CSV de esa OTA
    },
},
```

3. En `main.py`, agregar a `PLATFORM_COLORS` e `PLATFORM_ICONS`.
4. ¡Listo! El pipeline ETL se aplica automáticamente.

---

## 🐛 Troubleshooting

| Problema                          | Solución                                              |
|-----------------------------------|-------------------------------------------------------|
| `ModuleNotFoundError: flet`       | `pip install flet`                                    |
| Fechas como `NaT` en la tabla     | Verificar formato de fecha del CSV fuente             |
| Comisiones en $0                  | El CSV no tiene columna de comisión; se calcula auto  |
| Archivo no se abre                | Verificar que sea CSV UTF-8 o XLSX válido             |
| Gráficos no aparecen en dashboard | Presionar el botón ↻ Refresh del dashboard            |

---

## 📄 Licencia

Proyecto de uso interno hotelero. Libre para modificar y adaptar.
