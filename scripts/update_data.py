from __future__ import annotations

import json
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen


ROOT = Path(__file__).resolve().parents[1]
DATA_FILE = ROOT / "data" / "series.json"
API_URL = "https://apis.datos.gob.ar/series/api/series"

SERIES = [
    {"code": "P1", "id": "145.3_INGNACUAL_DICI_M_38", "title": "IPC nacional", "subtitle": "Variacion mensual", "group": "Precios", "format": "percent"},
    {"code": "P2", "id": "148.3_INUCLEONAL_DICI_M_19", "title": "IPC nucleo", "subtitle": "Variacion mensual", "group": "Precios", "format": "percent", "transform": "percent_change"},
    {"code": "P3", "id": "148.3_IREGULANAL_DICI_M_22", "title": "IPC regulados", "subtitle": "Variacion mensual", "group": "Precios", "format": "percent", "transform": "percent_change"},
    {"code": "P6", "id": "147.3_IBIENESNAL_DICI_T_19", "title": "IPC bienes", "subtitle": "Variacion mensual", "group": "Precios", "format": "percent", "transform": "percent_change"},
    {"code": "P7", "id": "147.3_ISERVICNAL_DICI_T_22", "title": "IPC servicios", "subtitle": "Variacion mensual", "group": "Precios", "format": "percent", "transform": "percent_change"},
    {"code": "P8", "id": "148.1_IPC_ESTACINAL_DICI_T_25", "title": "IPC estacionales", "subtitle": "Variacion trimestral", "group": "Precios", "format": "percent", "transform": "percent_change"},
    {"code": "A2", "id": "143.3_NO_PR_2004_A_31", "title": "EMAE", "subtitle": "Serie desestacionalizada", "group": "Actividad", "format": "number"},
    {"code": "S1", "id": "158.1_REPTE_0_0_5", "title": "RIPTE real", "subtitle": "Pesos constantes de noviembre de 2023", "group": "Ingresos", "format": "currency", "deflate": True},
    {"code": "S5", "id": "58.1_MP_0_M_13", "title": "Haber jubilatorio minimo real", "subtitle": "Pesos constantes de noviembre de 2023", "group": "Ingresos", "format": "currency", "deflate": True},
    {"code": "E1", "id": "42.3_EPH_PUNTUATAL_0_M_30", "title": "Desocupacion EPH", "subtitle": "Serie trimestral", "group": "Empleo", "format": "percent"},
    {"code": "X1", "id": "168.1_T_CAMBIOR_D_0_0_26", "title": "Tipo de cambio BNA vendedor", "subtitle": "Pesos por dolar", "group": "Sector externo", "format": "currency"},
    {"code": "X2", "id": "174.1_RRVAS_IDOS_0_0_36", "title": "Reservas internacionales", "subtitle": "Saldo mensual", "group": "Sector externo", "format": "usd_millions"},
]

IPC_CODE = "P1"
BASE_MONTH = "2023-11-01"


def fetch_one(item: dict) -> dict:
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
            item["deflator"] = {"series": ipc["id"], "base_month": BASE_MONTH, "method": "IPC nacional encadenado"}

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
    print(f"Actualizadas {len(series)} de {len(SERIES)} series")
    if errors:
        print("Advertencias:")
        for error in errors:
            print(f"- {error}")


if __name__ == "__main__":
    main()
