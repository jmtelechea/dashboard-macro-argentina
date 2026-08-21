# Dashboard macro Argentina

Pagina unica con indicadores macroeconomicos y sociales de Argentina obtenidos desde la API oficial de Series de Tiempo y la API de Estadisticas del BCRA.

## Actualizacion

GitHub Actions ejecuta `scripts/update_data.py` todos los dias a las 10:17 UTC. El script reintenta fallas transitorias y conserva la ultima copia valida cuando una serie no responde.

## Metodologia

- Base monetaria, prestamos al sector privado y M2 transaccional se deflactan con el IPC nacional encadenado. La base es el primer mes comparable de cada serie: enero de 2017 para base monetaria y prestamos, y enero de 2021 para M2 transaccional.
- Los cocientes sobre PIB usan el promedio de los saldos diarios de cada trimestre y el PIB nominal a precios corrientes del mismo trimestre.
- El grafico cambiario combina los limites inferior y superior del regimen de bandas con el tipo de cambio mayorista de referencia.
- La variacion de reservas por compra de divisas se presenta como promedio movil de cinco observaciones diarias.
- Las tasas de prestamos personales y de depositos a plazo fijo en pesos se comparan en un mismo grafico; la tasa de plazos fijos en dolares se muestra por separado.
- Exportaciones e importaciones mensuales se comparan en un mismo grafico; el saldo comercial se presenta por separado, en millones de dolares.
- Los indices mensuales de cantidades exportadas e importadas, base 2004=100, se presentan en graficos separados y se leen directamente del Excel que actualiza el INDEC.
- Los resultados primario y financiero mensuales corresponden al Sector Publico Nacional, base caja y metodologia 2017.

## Uso local

1. Ejecutar `python scripts/update_data.py`.
2. Servir la carpeta con un servidor web local.
3. Abrir `index.html` desde ese servidor.

## Fuente

- API: `https://apis.datos.gob.ar/series/api/series`
- API BCRA: `https://api.bcra.gob.ar/estadisticas/v4.0/Monetarias`
- Datos publicados por INDEC, BCRA y otros organismos oficiales.
