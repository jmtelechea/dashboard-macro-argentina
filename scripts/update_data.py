from __future__ import annotations

import csv
import json
import os
import shutil
import subprocess
import tempfile
import time
import unicodedata
from datetime import datetime, timezone
from collections import defaultdict
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import xlrd
import x13binary


ROOT = Path(__file__).resolve().parents[1]
DATA_FILE = ROOT / "data" / "series.json"
TASK_WORK_DIR = ROOT / "actividad_sectorial_emae"
IPC_HISTORY_FILE = ROOT / "extender_ipc_1997" / "ipc_alternativo_1997_2016.csv"
API_URL = "https://apis.datos.gob.ar/series/api/series"
BCRA_API_URL = "https://api.bcra.gob.ar/estadisticas/v4.0/Monetarias"
INDEC_COMEX_XLS_URL = "https://www.indec.gob.ar/ftp/cuadros/economia/serie_mensual_indices_comex.xls"
INDEC_EMAE_SECTORS_XLS_URL = "https://www.indec.gob.ar/ftp/cuadros/economia/sh_emae_actividad_base2004.xls"
INDEC_SALARY_HISTORY_XLS_URL = "https://www.indec.gob.ar/ftp/cuadros/sociedad/serie_is_2012.xls"
INDEC_SALARY_CURRENT_URL_TEMPLATE = "https://www.indec.gob.ar/ftp/cuadros/sociedad/variaciones_salarios_{month:02d}_{year:02d}.xls"
INDEC_COMEX_EXPORT_QUANTITIES_ID = "indec_comex_export_quantities_2004"
INDEC_COMEX_IMPORT_QUANTITIES_ID = "indec_comex_import_quantities_2004"
INDEC_EMAE_SECTORS_ID = "indec_emae_sectorial_sa_nov2023"
EMAE_SECTORS_CHART_ID = "emae_general_sectorial_sa_nov2023"
SALARY_SOURCE_CODE = "IS_PRIVATE_REGISTERED"
SALARY_SOURCE_ID = "indec_is_private_registered_spliced_2001"
SALARY_REAL_CHART_ID = "indec_is_private_registered_real_2001"
SALARY_REAL_COMPARISON_CHART_ID = "indec_is_sector_real_comparison_2022"
SALARY_REAL_BASE_START = "2023-01-01"
SALARY_REAL_BASE_END = "2023-11-01"
SALARY_COMPARISON_START = "2022-01-01"

SERIES = [
    {"code": "P1", "id": "145.3_INGNACUAL_DICI_M_38", "title": "IPC nacional (serie empalmada)", "subtitle": "Variacion mensual; alternativa hasta dic-16, INDEC desde ene-17", "group": "Precios", "format": "percent"},
    {"code": "P2", "id": "148.3_INUCLEONAL_DICI_M_19", "title": "IPC nucleo", "subtitle": "Variacion mensual", "group": "Precios", "format": "percent", "transform": "percent_change"},
    {"code": "P3", "id": "148.3_IREGULANAL_DICI_M_22", "title": "IPC regulados", "subtitle": "Variacion mensual", "group": "Precios", "format": "percent", "transform": "percent_change"},
    {"code": "P6", "id": "147.3_IBIENESNAL_DICI_T_19", "title": "IPC bienes", "subtitle": "Variacion mensual", "group": "Precios", "format": "percent", "transform": "percent_change"},
    {"code": "P7", "id": "147.3_ISERVICNAL_DICI_T_22", "title": "IPC servicios", "subtitle": "Variacion mensual", "group": "Precios", "format": "percent", "transform": "percent_change"},
    {"code": "P8", "id": "148.1_IPC_ESTACINAL_DICI_T_25", "title": "IPC estacionales", "subtitle": "Variacion trimestral", "group": "Precios", "format": "percent", "transform": "percent_change"},
    {"code": "A2", "id": "143.3_NO_PR_2004_A_31", "title": "EMAE", "subtitle": "Serie desestacionalizada", "group": "Actividad", "format": "number"},
    {"code": "S1", "id": "158.1_REPTE_0_0_5", "title": "RIPTE real", "subtitle": "Pesos constantes de noviembre de 2023", "group": "Ingresos", "format": "currency", "deflate": True},
    {"code": "S5", "id": "58.1_MP_0_M_13", "title": "Haber jubilatorio minimo real", "subtitle": "Pesos constantes de noviembre de 2023", "group": "Ingresos", "format": "currency", "deflate": True},
    {"code": "E1", "id": "42.3_EPH_PUNTUATAL_0_M_30", "title": "Desocupacion EPH", "subtitle": "Serie trimestral", "group": "Empleo", "format": "percent"},
    {"code": "EXT_EXPORT", "id": "74.3_IET_0_M_16", "title": "Exportaciones", "subtitle": "Millones de dolares", "group": "Sector externo", "format": "usd_millions", "hidden": True},
    {"code": "EXT_IMPORT", "id": "74.3_IIT_0_M_25", "title": "Importaciones", "subtitle": "Millones de dolares", "group": "Sector externo", "format": "usd_millions", "hidden": True},
    {"code": "EXT_BALANCE", "id": "74.3_ISC_0_M_19", "title": "Saldo comercial", "subtitle": "Millones de dolares", "group": "Sector externo", "format": "usd_millions"},
    {"code": "FISC_PRIMARY", "id": "379.9_RESULTADO_017__31_73", "title": "Resultado primario sin rentas", "subtitle": "Sector Publico Nacional, millones de pesos", "group": "Fiscal", "format": "ars_millions", "hidden": True},
    {"code": "FISC_FINANCIAL", "id": "379.9_RESULTADO_017__18_38", "title": "Resultado financiero", "subtitle": "Sector Publico Nacional, millones de pesos", "group": "Fiscal", "format": "ars_millions", "hidden": True},
    {"code": "GDP", "id": "4.4_OGP_2004_T_17", "title": "PIB nominal trimestral", "subtitle": "Millones de pesos corrientes", "group": "Auxiliar", "format": "ars_millions", "hidden": True},
    {"code": "B1248", "id": "1248", "title": "Base monetaria", "subtitle": "Saldo diario", "group": "BCRA", "format": "ars_millions", "provider": "bcra", "derive_real": True, "derive_gdp": True, "hidden": True},
    {"code": "B1266", "id": "1266", "title": "Depositos del Gobierno en el BCRA en moneda extranjera", "subtitle": "Saldo diario expresado en pesos", "group": "BCRA", "format": "ars_millions", "provider": "bcra"},
    {"code": "B1265", "id": "1265", "title": "Depositos del Gobierno en el BCRA en pesos", "subtitle": "Saldo diario", "group": "BCRA", "format": "ars_millions", "provider": "bcra"},
    {"code": "B107", "id": "107", "title": "Depositos totales en dolares", "subtitle": "Sectores publico y privado no financieros, saldo diario en millones de dolares", "group": "BCRA", "format": "usd_millions", "provider": "bcra"},
    {"code": "B1244", "id": "1244", "title": "Reservas internacionales BCRA", "subtitle": "Saldo diario", "group": "BCRA", "format": "usd_millions", "provider": "bcra"},
    {"code": "B1187", "id": "1187", "title": "Banda cambiaria: limite inferior", "subtitle": "Pesos por dolar", "group": "BCRA", "format": "exchange_rate", "provider": "bcra", "hidden": True},
    {"code": "B1188", "id": "1188", "title": "Banda cambiaria: limite superior", "subtitle": "Pesos por dolar", "group": "BCRA", "format": "exchange_rate", "provider": "bcra", "hidden": True},
    {"code": "B7", "id": "7", "title": "Tasa BADLAR de bancos privados", "subtitle": "Tasa nominal anual", "group": "BCRA", "format": "percent", "provider": "bcra"},
    {"code": "B150", "id": "150", "title": "Tasa de pases a 1 dia", "subtitle": "Tasa nominal anual", "group": "BCRA", "format": "percent", "provider": "bcra"},
    {"code": "B74", "id": "74", "title": "Reservas internacionales sin asignaciones DEG 2009", "subtitle": "Saldo diario en millones de dolares", "group": "BCRA", "format": "usd_millions", "provider": "bcra"},
    {"code": "B78", "id": "78", "title": "Compra de divisas del BCRA", "subtitle": "Variacion diaria de reservas", "group": "BCRA", "format": "usd_millions", "provider": "bcra", "hidden": True},
    {"code": "B84", "id": "84", "title": "Tipo de cambio de valuacion contable", "subtitle": "Pesos por dolar estadounidense", "group": "BCRA", "format": "exchange_rate", "provider": "bcra"},
    {"code": "B144", "id": "144", "title": "Prestamos personales en pesos", "subtitle": "Tasa nominal anual", "group": "BCRA", "format": "percent", "provider": "bcra", "hidden": True},
    {"code": "B1189", "id": "1189", "title": "Depositos a plazo fijo en pesos", "subtitle": "Tasa nominal anual", "group": "BCRA", "format": "percent", "provider": "bcra", "hidden": True},
    {"code": "B1193", "id": "1193", "title": "Depositos a plazo fijo en dolares", "subtitle": "Tasa nominal anual", "group": "BCRA", "format": "percent", "provider": "bcra"},
    {"code": "B5", "id": "5", "title": "Tipo de cambio mayorista de referencia", "subtitle": "Pesos por dolar", "group": "BCRA", "format": "exchange_rate", "provider": "bcra", "hidden": True},
    {"code": "B1341", "id": "1341", "title": "Prestamos al sector privado", "subtitle": "Saldo diario", "group": "BCRA", "format": "ars_millions", "provider": "bcra", "derive_real": True, "derive_gdp": True, "hidden": True},
    {"code": "B197", "id": "197", "title": "M2 transaccional del sector privado", "subtitle": "Saldo diario", "group": "BCRA", "format": "ars_millions", "provider": "bcra", "derive_real": True, "derive_gdp": True, "hidden": True},
    {"code": "B198", "id": "198", "title": "Stock de pasivos remunerados", "subtitle": "Saldo diario en millones de pesos", "group": "BCRA", "format": "ars_millions", "provider": "bcra", "chart_type": "bar"},
]

IPC_CODE = "P1"
BASE_MONTH = "2023-11-01"
REAL_CODES = ("B197", "B1341", "B1248")
EMAE_SECTOR_COLUMNS = (
    ("Agro (8%)", 2, "#C0504D", 8.1),
    ("Mineria y petroleo (5%)", 4, "#9BBB59", 5.0),
    ("Industria (19%)", 5, "#4BACC6", 18.9),
    ("Finanzas (3,1%)", 11, "#8064A2", 3.1),
    ("Comercio (12%)", 8, "#F79646", 12.4),
)


def month_key(date: str) -> str:
    return f"{date[:7]}-01"


def quarter_key(date: str) -> str:
    year, month = map(int, date[:7].split("-"))
    quarter_month = ((month - 1) // 3) * 3 + 1
    return f"{year:04d}-{quarter_month:02d}-01"


def extend_ipc_history(item: dict) -> dict:
    if not IPC_HISTORY_FILE.exists():
        raise RuntimeError(f"No existe la serie alternativa de IPC: {IPC_HISTORY_FILE}")
    levels = []
    with IPC_HISTORY_FILE.open(encoding="utf-8", newline="") as source:
        for row in csv.DictReader(source):
            levels.append({"date": row["date"], "value": float(row["level"])})
    if len(levels) < 2 or levels[0]["date"] != "1997-01-01" or levels[-1]["date"] != "2016-12-01":
        raise RuntimeError("La serie alternativa de IPC no cubre enero de 1997-diciembre de 2016")
    alternative = [
        {"date": current["date"], "value": current["value"] / previous["value"] - 1}
        for previous, current in zip(levels, levels[1:])
    ]
    official = [point for point in item["data"] if point["date"] >= "2017-01-01"]
    if not official or official[0]["date"] != "2017-01-01":
        raise RuntimeError("El IPC oficial no comienza en enero de 2017")
    result = dict(item)
    result.update({
        "description": "Variacion mensual del IPC: serie alternativa basada en indices provinciales hasta diciembre de 2016 e IPC nacional del INDEC desde enero de 2017",
        "units": "Variacion porcentual respecto del mes anterior",
        "source": "Elaboracion propia basada en indices provinciales hasta diciembre de 2016; INDEC desde enero de 2017",
        "frequency": "month",
        "data": alternative + official,
        "calculation": {
            "alternative_source": str(IPC_HISTORY_FILE.relative_to(ROOT)).replace("\\", "/"),
            "alternative_period": "1997-02 a 2016-12",
            "official_series": item["id"],
            "official_period": "desde 2017-01",
            "method": "Variaciones mensuales de la serie alternativa empalmadas con el IPC nacional oficial",
        },
    })
    return result


def make_real_series(item: dict, price_index: dict) -> dict:
    comparable = [point for point in item["data"] if month_key(point["date"]) in price_index]
    if not comparable:
        raise RuntimeError(f"No hay meses comparables entre IPC y {item['code']}")
    base_month = month_key(comparable[0]["date"])
    base_level = price_index[base_month]
    result = dict(item)
    result.update({
        "code": f"{item['code']}_REAL",
        "id": f"{item['id']}_real",
        "title": f"{item['title']} real",
        "subtitle": f"Millones de pesos constantes de {base_month[:7]}",
        "format": "ars_millions",
        "frequency": "day",
        "hidden": False,
        "data": [
            {"date": point["date"], "value": point["value"] * base_level / price_index[month_key(point["date"])]}
            for point in comparable
        ],
        "deflator": {"series": IPC_CODE, "base_month": base_month, "method": "IPC empalmado encadenado"},
    })
    return result


def make_gdp_series(item: dict, gdp: dict) -> dict:
    buckets = defaultdict(list)
    for point in item["data"]:
        buckets[quarter_key(point["date"])].append(point["value"])
    gdp_by_quarter = {point["date"]: point["value"] for point in gdp["data"]}
    data = []
    for quarter in sorted(set(buckets) & set(gdp_by_quarter)):
        average = sum(buckets[quarter]) / len(buckets[quarter])
        data.append({"date": quarter, "value": average / gdp_by_quarter[quarter] * 100})
    result = dict(item)
    result.update({
        "code": f"{item['code']}_GDP",
        "id": f"{item['id']}_gdp",
        "title": f"{item['title']} / PIB",
        "subtitle": "Promedio trimestral como porcentaje del PIB nominal",
        "format": "percent",
        "frequency": "quarter",
        "hidden": False,
        "data": data,
        "calculation": {"denominator": gdp["id"], "method": "Promedio diario trimestral / PIB nominal trimestral"},
    })
    return result


def make_exchange_chart(by_code: dict) -> dict:
    components = [by_code[code] for code in ("B1187", "B1188", "B5")]
    return {
        "code": "BANDAS_FX",
        "id": "1187_1188_5",
        "title": "Bandas cambiarias y tipo de cambio mayorista",
        "subtitle": "Pesos por dolar",
        "group": "BCRA",
        "format": "exchange_rate",
        "frequency": "day",
        "source": "Banco Central de la Republica Argentina",
        "lines": [
            {"label": component["title"], "data": component["data"], "color": color}
            for component, color in zip(components, ("rgb(150, 175, 209)", "#667788", "#0a2540"))
        ],
        "data": components[0]["data"],
    }


def make_rolling_average(item: dict, window: int = 5) -> dict:
    data = []
    values = [point["value"] for point in item["data"]]
    for index in range(window - 1, len(item["data"])):
        data.append({
            "date": item["data"][index]["date"],
            "value": sum(values[index - window + 1:index + 1]) / window,
        })
    result = dict(item)
    result.update({
        "code": f"{item['code']}_MA{window}",
        "id": f"{item['id']}_ma{window}",
        "title": "Compra de divisas del BCRA",
        "subtitle": f"Promedio movil de {window} observaciones diarias, millones de USD",
        "hidden": False,
        "data": data,
        "calculation": {"source": item["id"], "method": f"Promedio movil de {window} observaciones"},
    })
    return result


def make_peso_rates_chart(by_code: dict) -> dict:
    personal = by_code["B144"]
    term_deposits = by_code["B1189"]
    return {
        "code": "TASAS_PESOS",
        "id": "144_1189",
        "title": "Tasas de prestamos personales y plazos fijos",
        "subtitle": "Tasa nominal anual en pesos",
        "group": "BCRA",
        "format": "percent",
        "frequency": "day",
        "source": "Banco Central de la Republica Argentina",
        "lines": [
            {"label": personal["title"], "data": personal["data"], "color": "#0a2540"},
            {"label": term_deposits["title"], "data": term_deposits["data"], "color": "rgb(150, 175, 209)"},
        ],
        "data": term_deposits["data"],
    }


def make_two_line_chart(by_code: dict, first_code: str, second_code: str, code: str, title: str, subtitle: str, group: str, format_name: str) -> dict:
    first = by_code[first_code]
    second = by_code[second_code]
    return {
        "code": code,
        "id": f"{first['id']}_{second['id']}",
        "title": title,
        "subtitle": subtitle,
        "group": group,
        "format": format_name,
        "frequency": "month",
        "source": first["source"],
        "lines": [
            {"label": first["title"], "data": first["data"], "color": "#0a2540"},
            {"label": second["title"], "data": second["data"], "color": "rgb(150, 175, 209)"},
        ],
        "data": first["data"],
    }


def add_months(date: str, offset: int) -> str:
    year, month = map(int, date[:7].split("-"))
    month_index = year * 12 + month - 1 + offset
    return f"{month_index // 12:04d}-{month_index % 12 + 1:02d}-01"


def make_fiscal_gdp_chart(by_code: dict, gdp: dict) -> dict:
    primary = by_code["FISC_PRIMARY"]
    financial = by_code["FISC_FINANCIAL"]
    gdp_monthly = {
        add_months(point["date"], offset): point["value"]
        for point in gdp["data"]
        for offset in range(3)
    }

    def rolling_ratio(item: dict) -> list[dict]:
        points = item["data"]
        result = []
        for index in range(11, len(points)):
            window = points[index - 11:index + 1]
            expected_dates = [add_months(window[0]["date"], offset) for offset in range(12)]
            actual_dates = [point["date"] for point in window]
            if actual_dates != expected_dates or any(date not in gdp_monthly for date in actual_dates):
                continue
            rolling_result = sum(point["value"] for point in window)
            rolling_gdp = sum(gdp_monthly[date] for date in actual_dates) / 12
            if rolling_gdp:
                result.append({"date": window[-1]["date"], "value": rolling_result / rolling_gdp})
        return result

    primary_data = rolling_ratio(primary)
    financial_data = rolling_ratio(financial)
    if not primary_data or not financial_data:
        raise RuntimeError("No hay datos suficientes para calcular los resultados fiscales sobre PIB")
    return {
        "code": "FISC_RESULTS",
        "id": f"{primary['id']}_{financial['id']}_rolling12_gdp",
        "title": "Resultados primario y financiero",
        "subtitle": "Acumulado de 12 meses como porcentaje del PIB nominal",
        "group": "Fiscal",
        "format": "percent",
        "frequency": "month",
        "source": primary["source"],
        "lines": [
            {"label": primary["title"], "data": primary_data, "color": "#0a2540"},
            {"label": financial["title"], "data": financial_data, "color": "rgb(150, 175, 209)"},
        ],
        "data": primary_data,
        "calculation": {
            "primary_series": primary["id"],
            "financial_series": financial["id"],
            "denominator": gdp["id"],
            "method": "Suma movil de 12 resultados mensuales / promedio del PIB nominal trimestral anualizado correspondiente a esos 12 meses",
        },
    }


def normalize_text(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value).strip().lower())
    return "".join(character for character in text if not unicodedata.combining(character))


def run_x13(values: list[float], start_year: int, start_month: int) -> list[float]:
    formatted = [f"{value:.12g}" for value in values]
    data_lines = "\n  ".join(" ".join(formatted[index:index + 8]) for index in range(0, len(formatted), 8))
    spec = f"""series{{
 title=\"EMAE sectorial\"
 start={start_year}.{start_month}
 period=12
 data=({data_lines})
}}
transform{{function=auto}}
automdl{{}}
x11{{save=(d11)}}
"""
    TASK_WORK_DIR.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="emae_x13_", dir=TASK_WORK_DIR) as temp_name:
        temp = Path(temp_name)
        stem = temp / "sector"
        stem.with_suffix(".spc").write_text(spec, encoding="ascii")
        source_binary = Path(x13binary.find_x13_bin())
        binary = temp / ("x13as.exe" if os.name == "nt" else "x13as")
        shutil.copy2(source_binary, binary)
        if os.name != "nt":
            binary.chmod(binary.stat().st_mode | 0o111)
        result = subprocess.run(
            [str(binary), stem.name], cwd=temp, capture_output=True, text=True, timeout=120,
        )
        output = stem.with_suffix(".d11")
        if result.returncode != 0 or not output.exists():
            detail = (result.stderr or result.stdout).strip()[-800:]
            raise RuntimeError(f"X-13 no genero la serie desestacionalizada: {detail}")
        adjusted = []
        for line in output.read_text(encoding="ascii", errors="ignore").splitlines():
            parts = line.split()
            if len(parts) >= 2 and len(parts[0]) == 6 and parts[0].isdigit():
                adjusted.append(float(parts[1]))
        if len(adjusted) != len(values):
            raise RuntimeError(f"X-13 devolvio {len(adjusted)} valores para {len(values)} observaciones")
        return adjusted


def fetch_indec_emae_sectors() -> dict:
    months = {
        "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
        "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
        "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
    }
    request = Request(INDEC_EMAE_SECTORS_XLS_URL, headers={"User-Agent": "dashboard-macro-argentina/1.0"})
    last_error = None
    for attempt in range(3):
        try:
            with urlopen(request, timeout=45) as response:
                workbook = xlrd.open_workbook(file_contents=response.read())
            if "Tabla Letras" not in workbook.sheet_names():
                raise RuntimeError("El Excel sectorial no contiene la hoja Tabla Letras")
            sheet = workbook.sheet_by_name("Tabla Letras")
            dates = []
            values_by_column = {column: [] for _, column, _, _ in EMAE_SECTOR_COLUMNS}
            current_year = None
            for row in range(4, sheet.nrows):
                year_cell = sheet.cell_value(row, 0)
                if isinstance(year_cell, (int, float)) and year_cell:
                    current_year = int(year_cell)
                month = months.get(normalize_text(sheet.cell_value(row, 1)))
                if current_year is None or month is None:
                    continue
                row_values = [sheet.cell_value(row, column) for column in values_by_column]
                if not all(isinstance(value, (int, float)) for value in row_values):
                    continue
                dates.append(f"{current_year:04d}-{month:02d}-01")
                for column, value in zip(values_by_column, row_values):
                    values_by_column[column].append(float(value))
            if not dates or dates[0] != "2004-01-01":
                raise RuntimeError("La serie sectorial no comienza en enero de 2004")
            if BASE_MONTH not in dates:
                raise RuntimeError(f"El Excel sectorial no contiene el mes base {BASE_MONTH}")
            base_index = dates.index(BASE_MONTH)
            lines = []
            for label, column, color, weight in EMAE_SECTOR_COLUMNS:
                adjusted = run_x13(values_by_column[column], 2004, 1)
                base_value = adjusted[base_index]
                if base_value == 0:
                    raise RuntimeError(f"La serie {label} tiene base cero en noviembre de 2023")
                lines.append({
                    "label": label,
                    "color": color,
                    "weight_2004": weight,
                    "data": [
                        {"date": date, "value": value / base_value * 100}
                        for date, value in zip(dates, adjusted)
                    ],
                })
            return {
                "code": "EMAE_SECTORS_SOURCE",
                "id": INDEC_EMAE_SECTORS_ID,
                "title": "Actividad economica por sectores",
                "subtitle": "Series desestacionalizadas, noviembre de 2023=100",
                "group": "Actividad",
                "format": "number",
                "frequency": "month",
                "source": "INDEC, Estimador Mensual de Actividad Economica por sector",
                "seasonal_adjustment": "X-13ARIMA-SEATS, X-11 final seasonal adjustment (d11)",
                "base_month": BASE_MONTH,
                "weights_base_2004": {label: weight for label, _, _, weight in EMAE_SECTOR_COLUMNS},
                "lines": lines,
                "data": lines[0]["data"],
                "hidden": True,
            }
        except Exception as exc:
            last_error = exc
            time.sleep(2 ** attempt)
    raise RuntimeError(f"No se pudo procesar el Excel sectorial del EMAE: {last_error}")


def make_emae_sectors_chart(emae: dict, sectors: dict) -> dict:
    emae_by_date = {point["date"]: point["value"] for point in emae["data"]}
    base_value = emae_by_date.get(BASE_MONTH)
    if base_value is None or base_value == 0:
        raise RuntimeError(f"El EMAE general no contiene el mes base {BASE_MONTH}")
    emae_line = {
        "label": "EMAE",
        "color": "#111111",
        "borderDash": [8, 6],
        "data": [
            {"date": point["date"], "value": point["value"] / base_value * 100}
            for point in emae["data"]
        ],
    }
    result = dict(sectors)
    result.update({
        "code": "EMAE_SECTORS",
        "id": EMAE_SECTORS_CHART_ID,
        "hidden": False,
        "lines": [emae_line, *sectors["lines"]],
        "data": sectors["data"],
        "calculation": {
            "base": "Noviembre de 2023=100 para cada serie",
            "general": emae["id"],
            "sectors": sectors["id"],
        },
    })
    return result


def fetch_indec_comex_quantities() -> list[dict]:
    months = {
        "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
        "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
        "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
    }
    request = Request(INDEC_COMEX_XLS_URL, headers={"User-Agent": "dashboard-macro-argentina/1.0"})
    last_error = None
    for attempt in range(3):
        try:
            with urlopen(request, timeout=45) as response:
                workbook = xlrd.open_workbook(file_contents=response.read())
            sheet = workbook.sheet_by_index(0)
            export_data = []
            import_data = []
            current_year = None
            for row in range(5, sheet.nrows):
                year_cell = str(sheet.cell_value(row, 0)).strip()
                if year_cell:
                    digits = "".join(character for character in year_cell if character.isdigit())
                    if len(digits) >= 4:
                        current_year = int(digits[:4])
                month = months.get(normalize_text(sheet.cell_value(row, 1)))
                if current_year is None or month is None:
                    continue
                export_value = sheet.cell_value(row, 4)
                import_value = sheet.cell_value(row, 8)
                if not isinstance(export_value, (int, float)) or not isinstance(import_value, (int, float)):
                    continue
                date = f"{current_year:04d}-{month:02d}-01"
                export_data.append({"date": date, "value": float(export_value)})
                import_data.append({"date": date, "value": float(import_value)})
            if not export_data or len(export_data) != len(import_data):
                raise RuntimeError("El Excel no contiene las dos series de cantidades esperadas")
            common = {
                "subtitle": "Indice de cantidad, base 2004=100",
                "group": "Sector externo",
                "format": "number",
                "frequency": "month",
                "source": "INDEC, indices de comercio exterior",
            }
            return [
                {
                    **common,
                    "code": "EXPORT_QUANTITIES",
                    "id": INDEC_COMEX_EXPORT_QUANTITIES_ID,
                    "title": "Cantidades exportadas",
                    "data": export_data,
                },
                {
                    **common,
                    "code": "IMPORT_QUANTITIES",
                    "id": INDEC_COMEX_IMPORT_QUANTITIES_ID,
                    "title": "Cantidades importadas",
                    "data": import_data,
                },
            ]
        except Exception as exc:
            last_error = exc
            time.sleep(2 ** attempt)
    raise RuntimeError(f"No se pudo descargar el Excel mensual de comercio exterior: {last_error}")


def download_salary_workbooks() -> tuple[xlrd.book.Book, xlrd.book.Book, str]:
    request = Request(INDEC_SALARY_HISTORY_XLS_URL, headers={"User-Agent": "dashboard-macro-argentina/1.0"})
    last_error = None
    for attempt in range(3):
        try:
            with urlopen(request, timeout=45) as response:
                history = xlrd.open_workbook(file_contents=response.read())
            break
        except Exception as exc:
            last_error = exc
            time.sleep(2 ** attempt)
    else:
        raise RuntimeError(f"No se pudo descargar el Excel historico del Indice de Salarios: {last_error}")

    now = datetime.now(timezone.utc)
    current = None
    current_url = None
    candidate_errors = []
    for offset in range(13):
        absolute_month = now.year * 12 + now.month - 1 - offset
        year, zero_based_month = divmod(absolute_month, 12)
        month = zero_based_month + 1
        url = INDEC_SALARY_CURRENT_URL_TEMPLATE.format(month=month, year=year % 100)
        try:
            request = Request(url, headers={"User-Agent": "dashboard-macro-argentina/1.0"})
            with urlopen(request, timeout=45) as response:
                current = xlrd.open_workbook(file_contents=response.read())
            current_url = url
            break
        except Exception as exc:
            candidate_errors.append(f"{url}: {exc}")
    if current is None or current_url is None:
        raise RuntimeError(f"No se encontro un Excel vigente del Indice de Salarios: {candidate_errors[-1]}")
    return history, current, current_url


def fetch_indec_salary_index() -> dict:
    months = {
        "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
        "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
        "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
    }
    history_workbook, current_workbook, current_url = download_salary_workbooks()
    history_sheet = history_workbook.sheet_by_name("Serie historica IS") if "Serie historica IS" in history_workbook.sheet_names() else history_workbook.sheet_by_index(1)
    historical = []
    for row in range(7, history_sheet.nrows):
        date_cell = history_sheet.cell_value(row, 0)
        value = history_sheet.cell_value(row, 1)
        if not isinstance(date_cell, (int, float)) or not isinstance(value, (int, float)):
            continue
        date = xlrd.xldate_as_datetime(date_cell, history_workbook.datemode).strftime("%Y-%m-01")
        historical.append({"date": date, "value": float(value)})

    current_sheet = current_workbook.sheet_by_name("Cuadro 1")
    current = []
    current_year = None
    for row in range(10, current_sheet.nrows):
        year_cell = current_sheet.cell_value(row, 0)
        if isinstance(year_cell, (int, float)) and year_cell:
            current_year = int(year_cell)
        month = months.get(normalize_text(current_sheet.cell_value(row, 1)))
        value = current_sheet.cell_value(row, 3)
        if current_year is None or month is None or not isinstance(value, (int, float)):
            continue
        current.append({"date": f"{current_year:04d}-{month:02d}-01", "value": float(value)})

    public_sheet = current_workbook.sheet_by_name("Cuadro 3")
    public_national = []
    public_provincial = []
    public_year = None
    national_level = 100.0
    provincial_level = 100.0
    for row in range(7, public_sheet.nrows):
        year_cell = public_sheet.cell_value(row, 0)
        if isinstance(year_cell, (int, float)) and year_cell:
            public_year = int(year_cell)
        month = months.get(normalize_text(public_sheet.cell_value(row, 1)))
        national_change = public_sheet.cell_value(row, 2)
        provincial_change = public_sheet.cell_value(row, 5)
        if (
            public_year is None or month is None
            or not isinstance(national_change, (int, float))
            or not isinstance(provincial_change, (int, float))
        ):
            continue
        date = f"{public_year:04d}-{month:02d}-01"
        national_level *= 1 + float(national_change) / 100
        provincial_level *= 1 + float(provincial_change) / 100
        public_national.append({"date": date, "value": national_level})
        public_provincial.append({"date": date, "value": provincial_level})

    if not historical or historical[0]["date"] != "2001-10-01" or historical[-1]["date"] != "2015-10-01":
        raise RuntimeError("La serie historica del Indice de Salarios no cubre octubre de 2001-octubre de 2015")
    if not current or current[0]["date"] != "2015-10-01":
        raise RuntimeError("La serie vigente del Indice de Salarios no comienza en octubre de 2015")
    if (
        not public_national or not public_provincial
        or public_national[0]["date"] != SALARY_COMPARISON_START
        or public_provincial[0]["date"] != SALARY_COMPARISON_START
        or public_national[-1]["date"] != current[-1]["date"]
        or public_provincial[-1]["date"] != current[-1]["date"]
    ):
        raise RuntimeError("Las series salariales publicas no cubren enero de 2022 hasta el ultimo mes vigente")
    historical_anchor = historical[-1]["value"]
    current_anchor = current[0]["value"]
    if historical_anchor == 0:
        raise RuntimeError("El Indice de Salarios historico tiene un ancla igual a cero")
    splice_factor = current_anchor / historical_anchor
    data = [
        {"date": point["date"], "value": point["value"] * splice_factor}
        for point in historical[:-1]
    ] + current
    return {
        "code": SALARY_SOURCE_CODE,
        "id": SALARY_SOURCE_ID,
        "title": "Indice de salarios del sector privado registrado",
        "subtitle": "Serie nominal empalmada",
        "group": "Ingresos",
        "format": "number",
        "frequency": "month",
        "source": "INDEC, Indice de salarios",
        "hidden": True,
        "data": data,
        "public_national": public_national,
        "public_provincial": public_provincial,
        "calculation": {
            "current_source": current_url,
            "historical_source": INDEC_SALARY_HISTORY_XLS_URL,
            "splice_month": "2015-10",
            "method": "Serie vigente desde octubre de 2015; hacia atras, variaciones mensuales de la serie historica",
            "splice_factor": splice_factor,
        },
    }


def make_real_salary_index(item: dict, price_index: dict) -> dict:
    raw = [
        {"date": point["date"], "value": point["value"] / price_index[point["date"]]}
        for point in item["data"]
        if point["date"] in price_index
    ]
    base_values = [
        point["value"] for point in raw
        if SALARY_REAL_BASE_START <= point["date"] <= SALARY_REAL_BASE_END
    ]
    if len(base_values) != 11:
        raise RuntimeError("El Indice de Salarios real no contiene los once meses de enero-noviembre de 2023")
    base_value = sum(base_values) / len(base_values)
    if base_value == 0:
        raise RuntimeError("La base del Indice de Salarios real es cero")
    return {
        "code": "IS_PRIVATE_REGISTERED_REAL",
        "id": SALARY_REAL_CHART_ID,
        "title": "Salario privado registrado real",
        "subtitle": "Indice, promedio enero-noviembre de 2023=100",
        "group": "Ingresos",
        "format": "number",
        "frequency": "month",
        "source": "INDEC, Indice de salarios; elaboracion propia con IPC empalmado",
        "data": [
            {"date": point["date"], "value": point["value"] / base_value * 100}
            for point in raw
        ],
        "deflator": {"series": IPC_CODE, "method": "IPC empalmado encadenado"},
        "calculation": {
            "nominal_series": item["id"],
            "real_method": "Indice de salarios nominal / nivel del IPC empalmado",
            "base": "Promedio enero-noviembre de 2023=100",
            "base_months": 11,
            "sources": item["calculation"],
        },
    }


def make_real_salary_comparison(item: dict, price_index: dict) -> dict:
    nominal_lines = [
        ("Privado registrado", [point for point in item["data"] if point["date"] >= SALARY_COMPARISON_START], "#0a2540"),
        ("Publico nacional", item["public_national"], "rgb(150, 175, 209)"),
        ("Publico provincial", item["public_provincial"], "#667788"),
    ]
    lines = []
    expected_dates = None
    for label, nominal_data, color in nominal_lines:
        raw = [
            {"date": point["date"], "value": point["value"] / price_index[point["date"]]}
            for point in nominal_data
            if point["date"] in price_index
        ]
        dates = [point["date"] for point in raw]
        if expected_dates is None:
            expected_dates = dates
        elif dates != expected_dates:
            raise RuntimeError("Las tres series salariales no tienen la misma cobertura mensual")
        base_values = [
            point["value"] for point in raw
            if SALARY_REAL_BASE_START <= point["date"] <= SALARY_REAL_BASE_END
        ]
        if len(base_values) != 11:
            raise RuntimeError(f"La serie {label} no contiene los once meses de la base")
        base_value = sum(base_values) / len(base_values)
        if base_value == 0:
            raise RuntimeError(f"La base de la serie {label} es cero")
        lines.append({
            "label": label,
            "color": color,
            "data": [
                {"date": point["date"], "value": point["value"] / base_value * 100}
                for point in raw
            ],
        })
    return {
        "code": "IS_SECTOR_REAL_COMPARISON",
        "id": SALARY_REAL_COMPARISON_CHART_ID,
        "title": "Salarios reales por sector",
        "subtitle": "Indice, promedio enero-noviembre de 2023=100",
        "group": "Ingresos",
        "format": "number",
        "frequency": "month",
        "source": "INDEC, Indice de salarios; elaboracion propia con IPC empalmado",
        "data": lines[0]["data"],
        "lines": lines,
        "deflator": {"series": IPC_CODE, "method": "IPC empalmado encadenado"},
        "calculation": {
            "nominal_private_series": item["id"],
            "public_source": item["calculation"]["current_source"],
            "public_columns": "Cuadro 3: variacion mensual del sector publico nacional y provincial",
            "public_method": "Indices encadenados a partir de las variaciones mensuales desde enero de 2022",
            "real_method": "Indice salarial nominal / nivel del IPC empalmado",
            "base": "Promedio enero-noviembre de 2023=100 para cada serie",
            "base_months": 11,
        },
    }


def fetch_bcra(item: dict) -> dict:
    points = []
    offset = 0
    total = None
    while total is None or offset < total:
        params = urlencode({"offset": offset, "limit": 1000})
        request = Request(
            f"{BCRA_API_URL}/{item['id']}?{params}",
            headers={"User-Agent": "dashboard-macro-argentina/1.0"},
        )
        with urlopen(request, timeout=45) as response:
            payload = json.load(response)
        if payload.get("status") != 200 or not payload.get("results"):
            raise RuntimeError("La API del BCRA no devolvio resultados")
        total = payload.get("metadata", {}).get("resultset", {}).get("count", 0)
        detail = payload["results"][0].get("detalle", [])
        points.extend({"date": point["fecha"], "value": point["valor"]} for point in detail if point.get("valor") is not None)
        if not detail:
            break
        offset += len(detail)
    points.sort(key=lambda point: point["date"])
    if not points:
        raise RuntimeError("La API del BCRA devolvio una serie sin observaciones")
    result = dict(item)
    result.update({
        "description": item["title"],
        "units": item["subtitle"],
        "source": "Banco Central de la Republica Argentina",
        "frequency": "day",
        "data": points,
    })
    return result


def fetch_datos_argentina(item: dict) -> dict:
    query = {"ids": item["id"], "last": 1000, "metadata": "simple"}
    params = urlencode(query)
    request = Request(f"{API_URL}?{params}", headers={"User-Agent": "dashboard-macro-argentina/1.0"})
    last_error = None
    for attempt in range(3):
        try:
            with urlopen(request, timeout=45) as response:
                payload = json.load(response)
            if not payload.get("data"):
                raise RuntimeError("La API devolvio una serie sin observaciones")
            meta = payload.get("meta", [])
            field = meta[1].get("field", {}) if len(meta) > 1 else {}
            dataset = meta[1].get("dataset", {}) if len(meta) > 1 else {}
            points = [{"date": row[0], "value": row[1]} for row in payload["data"] if len(row) > 1 and row[1] is not None]
            if item.get("transform") == "percent_change":
                points = [
                    {"date": current["date"], "value": current["value"] / previous["value"] - 1}
                    for previous, current in zip(points, points[1:])
                    if previous["value"] != 0
                ]
            result = dict(item)
            result.update({
                "description": field.get("description", item["title"]),
                "units": "Variacion porcentual respecto del periodo anterior" if item.get("transform") == "percent_change" else (field.get("representation_mode_units") or field.get("units") or ""),
                "source": dataset.get("source") or "Datos Argentina",
                "frequency": meta[0].get("frequency") if meta else None,
                "data": points,
            })
            if not result["data"]:
                raise RuntimeError("La serie solo contiene valores nulos")
            return result
        except Exception as exc:
            last_error = exc
            time.sleep(2 ** attempt)
    raise RuntimeError(f"No se pudo descargar {item['code']} ({item['id']}): {last_error}")


def fetch_one(item: dict) -> dict:
    if item.get("provider") == "bcra":
        last_error = None
        for attempt in range(3):
            try:
                return fetch_bcra(item)
            except Exception as exc:
                last_error = exc
                time.sleep(2 ** attempt)
        raise RuntimeError(f"No se pudo descargar {item['code']} ({item['id']}): {last_error}")
    result = fetch_datos_argentina(item)
    return extend_ipc_history(result) if item["code"] == IPC_CODE else result


def main() -> None:
    previous = {}
    if DATA_FILE.exists():
        try:
            previous_payload = json.loads(DATA_FILE.read_text(encoding="utf-8"))
            previous = {item["id"]: item for item in previous_payload.get("series", [])}
        except (OSError, json.JSONDecodeError):
            previous = {}

    series = []
    errors = []
    for item in SERIES:
        try:
            series.append(fetch_one(item))
        except Exception as exc:
            if item["id"] in previous:
                series.append(previous[item["id"]])
                errors.append(f"{item['code']}: se mantuvo la copia anterior")
            else:
                errors.append(str(exc))

    try:
        series.extend(fetch_indec_comex_quantities())
    except Exception as exc:
        quantity_ids = (INDEC_COMEX_EXPORT_QUANTITIES_ID, INDEC_COMEX_IMPORT_QUANTITIES_ID)
        if all(series_id in previous for series_id in quantity_ids):
            series.extend(previous[series_id] for series_id in quantity_ids)
            errors.append("COMEX_QUANTITIES: se mantuvo la copia anterior")
        else:
            errors.append(str(exc))

    try:
        series.append(fetch_indec_emae_sectors())
    except Exception as exc:
        if EMAE_SECTORS_CHART_ID in previous:
            series.append(previous[EMAE_SECTORS_CHART_ID])
            errors.append("EMAE_SECTORS: se mantuvo la copia anterior")
        else:
            errors.append(str(exc))

    try:
        series.append(fetch_indec_salary_index())
    except Exception as exc:
        if SALARY_REAL_CHART_ID in previous:
            series.append(previous[SALARY_REAL_CHART_ID])
            errors.append("IS_PRIVATE_REGISTERED_REAL: se mantuvo la copia anterior")
            if SALARY_REAL_COMPARISON_CHART_ID in previous:
                series.append(previous[SALARY_REAL_COMPARISON_CHART_ID])
                errors.append("IS_SECTOR_REAL_COMPARISON: se mantuvo la copia anterior")
        else:
            errors.append(str(exc))

    by_code = {item["code"]: item for item in series}
    ipc = by_code.get(IPC_CODE)
    if ipc:
        price_level = 1.0
        price_index = {}
        for point in ipc["data"]:
            price_level *= 1 + point["value"]
            price_index[point["date"]] = price_level
        base_level = price_index.get(BASE_MONTH)
        if base_level is None:
            raise RuntimeError(f"No existe IPC para el mes base {BASE_MONTH}")
        for item in series:
            if not item.get("deflate"):
                continue
            real_data = []
            for point in item["data"]:
                level = price_index.get(point["date"])
                if level is not None:
                    real_data.append({"date": point["date"], "value": point["value"] * base_level / level})
            if not real_data:
                raise RuntimeError(f"No hay meses comparables entre IPC y {item['code']}")
            item["data"] = real_data
            item["units"] = "Pesos constantes de noviembre de 2023"
            item["deflator"] = {"series": IPC_CODE, "base_month": BASE_MONTH, "method": "IPC empalmado encadenado"}

        price_level = 1.0
        price_index = {}
        for point in ipc["data"]:
            price_level *= 1 + point["value"]
            price_index[point["date"]] = price_level
        gdp = by_code.get("GDP")
        if not gdp:
            raise RuntimeError("No se pudo obtener el PIB nominal trimestral")
        derived = []
        for code in REAL_CODES:
            source = by_code.get(code)
            if not source:
                raise RuntimeError(f"No se pudo obtener la serie fuente {code}")
            derived.extend((make_real_series(source, price_index), make_gdp_series(source, gdp)))
        series.extend(derived)
        series.append(make_exchange_chart(by_code))
        series.append(make_rolling_average(by_code["B78"]))
        series.append(make_peso_rates_chart(by_code))
        series.append(make_two_line_chart(
            by_code, "EXT_EXPORT", "EXT_IMPORT", "TRADE_FLOWS",
            "Exportaciones e importaciones", "Millones de dolares por mes",
            "Sector externo", "usd_millions",
        ))
        series.append(make_fiscal_gdp_chart(by_code, gdp))
        sectors = by_code.get("EMAE_SECTORS_SOURCE")
        if sectors:
            series.append(make_emae_sectors_chart(by_code["A2"], sectors))
        salary_source = by_code.get(SALARY_SOURCE_CODE)
        if salary_source:
            try:
                series.append(make_real_salary_index(salary_source, price_index))
            except Exception as exc:
                if SALARY_REAL_CHART_ID in previous:
                    series.append(previous[SALARY_REAL_CHART_ID])
                    errors.append("IS_PRIVATE_REGISTERED_REAL: se mantuvo la copia anterior")
                else:
                    raise RuntimeError(f"No se pudo calcular el Indice de Salarios real: {exc}") from exc
            try:
                series.append(make_real_salary_comparison(salary_source, price_index))
            except Exception as exc:
                if SALARY_REAL_COMPARISON_CHART_ID in previous:
                    series.append(previous[SALARY_REAL_COMPARISON_CHART_ID])
                    errors.append("IS_SECTOR_REAL_COMPARISON: se mantuvo la copia anterior")
                else:
                    raise RuntimeError(f"No se pudo calcular la comparacion de salarios reales: {exc}") from exc

    series = [item for item in series if not item.get("hidden")]

    if not series:
        raise RuntimeError("No se pudo obtener ninguna serie y no existe una copia anterior")

    payload = {
        "updated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "api": API_URL,
        "errors": errors,
        "series": series,
    }
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    DATA_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Actualizadas {len(series)} series visibles")
    if errors:
        print("Advertencias:")
        for error in errors:
            print(f"- {error}")


if __name__ == "__main__":
    main()
