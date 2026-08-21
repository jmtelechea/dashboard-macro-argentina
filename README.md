# Dashboard macro Argentina

Pagina unica con indicadores macroeconomicos y sociales de Argentina obtenidos principalmente de fuentes oficiales. El IPC incorpora una extension historica alternativa explicitamente identificada.

## Actualizacion

GitHub Actions ejecuta `scripts/update_data.py` todos los dias a las 10:17 UTC. El script reintenta fallas transitorias y conserva la ultima copia valida cuando una serie no responde.

## Metodologia

- El grafico de IPC usa una serie empalmada: variaciones calculadas desde un indice alternativo basado en indices provinciales entre febrero de 1997 y diciembre de 2016, e IPC nacional oficial del INDEC desde enero de 2017.
- Todas las variables reales se deflactan con ese IPC empalmado. RIPTE y haber jubilatorio se expresan en pesos constantes de noviembre de 2023; las series monetarias derivadas toman como base el primer mes comparable de cada serie.
- Los cocientes sobre PIB usan el promedio de los saldos diarios de cada trimestre y el PIB nominal a precios corrientes del mismo trimestre.
- El grafico cambiario combina los limites inferior y superior del regimen de bandas con el tipo de cambio mayorista de referencia.
- La variacion de reservas por compra de divisas se presenta como promedio movil de cinco observaciones diarias.
- Las tasas de prestamos personales y de depositos a plazo fijo en pesos se comparan en un mismo grafico; la tasa de plazos fijos en dolares se muestra por separado.
- Exportaciones e importaciones mensuales se comparan en un mismo grafico; el saldo comercial se presenta por separado, en millones de dolares.
- Los indices mensuales de cantidades exportadas e importadas, base 2004=100, se presentan en graficos separados y se leen directamente del Excel que actualiza el INDEC.
- El grafico de actividad economica por sectores combina el EMAE general desestacionalizado con cinco sectores del Excel mensual del INDEC. Los sectores se ajustan cada dia con X-13ARIMA-SEATS (ajuste final X-11, tabla d11) y todas las lineas se expresan con noviembre de 2023=100. La leyenda muestra las ponderaciones del ano base 2004 redondeadas como en la referencia visual.
- Los resultados primario y financiero mensuales corresponden al Sector Publico Nacional, base caja y metodologia 2017.

## Uso local

1. Ejecutar `python scripts/update_data.py`.
2. Servir la carpeta con un servidor web local.
3. Abrir `index.html` desde ese servidor.

## Fuente

- API: `https://apis.datos.gob.ar/series/api/series`
- API BCRA: `https://api.bcra.gob.ar/estadisticas/v4.0/Monetarias`
- Excel sectorial del EMAE: `https://www.indec.gob.ar/ftp/cuadros/economia/sh_emae_actividad_base2004.xls`
- IPC alternativo 1997-2016: `extender_ipc_1997/ipc_alternativo_1997_2016.csv`
- Datos publicados por INDEC, BCRA y otros organismos oficiales.
