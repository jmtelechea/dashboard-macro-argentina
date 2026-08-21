import json
from urllib.request import urlopen

import xlrd


DATA_URL = "https://www.indec.gob.ar/ftp/cuadros/sociedad/variaciones_salarios_08_26.xls"
MONTHS = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
}


with open("data/series.json", encoding="utf-8") as source:
    payload = json.load(source)

comparison = next(
    item for item in payload["series"]
    if item["id"] == "indec_is_sector_real_comparison_2022"
)
ipc = {
    point["date"]: point["value"]
    for point in next(item for item in payload["series"] if item["code"] == "P1")["data"]
}

workbook = xlrd.open_workbook(file_contents=urlopen(DATA_URL, timeout=45).read())
sheet = workbook.sheet_by_name("Cuadro 3")
rates = {}
year = None
for row in range(7, sheet.nrows):
    year_cell = sheet.cell_value(row, 0)
    if isinstance(year_cell, (int, float)) and year_cell:
        year = int(year_cell)
    month = MONTHS.get(str(sheet.cell_value(row, 1)).lower())
    national = sheet.cell_value(row, 2)
    provincial = sheet.cell_value(row, 5)
    if (
        year is not None and month is not None
        and isinstance(national, (int, float))
        and isinstance(provincial, (int, float))
    ):
        rates[f"{year:04d}-{month:02d}-01"] = (float(national) / 100, float(provincial) / 100)

errors = []
for line, column in zip(comparison["lines"][1:], (0, 1)):
    for previous, current in zip(line["data"], line["data"][1:]):
        implied_nominal_change = (
            current["value"] / previous["value"]
            * (1 + ipc[current["date"]])
            - 1
        )
        errors.append(abs(implied_nominal_change - rates[current["date"]][column]))

print(f"comparisons={len(errors)}")
print(f"max_error={max(errors):.16g}")
