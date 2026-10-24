# Reparacion de la actualizacion diaria

Fecha: 2026-10-24.

El workflow diario fallaba al construir el grafico fiscal con `KeyError: FISC_PRIMARY`.
Las series de origen ocultas se eliminaban antes de guardar el JSON publico, por lo
que no existia respaldo para recuperarlas cuando la descarga fallaba.

La correccion guarda niveles de origen sin transformar en `source_series.json`.
Los cinco graficos fiscales recuperan de forma independiente su ultima copia valida
si falta una fuente o falla su calculo. Las advertencias conservan la causa y los
datos respaldados conservan sus fechas originales. Los niveles ya deflactados del
JSON publico no se vuelven a deflactar cuando se utilizan como respaldo.

GitHub Actions conserva tambien este archivo de origen, evita ejecuciones
simultaneas y ejecuta una actualizacion al cambiar el script o su configuracion.

Prueba de regresion:

```text
python reparar_actualizacion_dashboard/test_update_fallback.py
```
