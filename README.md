# Dashboard macro Argentina

Pagina unica con indicadores macroeconomicos y sociales de Argentina obtenidos desde la API oficial de Series de Tiempo y la API de Estadisticas del BCRA.

## Actualizacion

GitHub Actions ejecuta `scripts/update_data.py` todos los dias a las 10:17 UTC. El script reintenta fallas transitorias y conserva la ultima copia valida cuando una serie no responde.

## Uso local

1. Ejecutar `python scripts/update_data.py`.
2. Servir la carpeta con un servidor web local.
3. Abrir `index.html` desde ese servidor.

## Fuente

- API: `https://apis.datos.gob.ar/series/api/series`
- API BCRA: `https://api.bcra.gob.ar/estadisticas/v4.0/Monetarias`
- Datos publicados por INDEC, BCRA y otros organismos oficiales.
