# Dashboard macro Argentina

Pagina unica con indicadores macroeconomicos y sociales de Argentina obtenidos principalmente de fuentes oficiales. El IPC incorpora una extension historica alternativa explicitamente identificada.

## Actualizacion

GitHub Actions ejecuta `scripts/update_data.py` todos los dias a las 10:17 UTC. El script reintenta fallas transitorias y conserva la ultima copia valida cuando una serie no responde.

## Metodologia

- El grafico de IPC usa una serie empalmada: variaciones calculadas desde un indice alternativo basado en indices provinciales entre febrero de 1997 y diciembre de 2016, e IPC nacional oficial del INDEC desde enero de 2017.
- Todas las variables reales se deflactan con ese IPC empalmado. RIPTE y haber jubilatorio se expresan en pesos constantes de noviembre de 2023; las series monetarias derivadas toman como base el primer mes comparable de cada serie.
- El salario del sector privado registrado real se construye con el Indice de Salarios vigente del INDEC desde octubre de 2015 y, hacia atras, con las variaciones de la serie historica hasta octubre de 2001. Se deflacta con el IPC empalmado y se expresa con promedio enero-noviembre de 2023=100.
- El grafico comparativo de salarios reales presenta desde enero de 2022 los indices del sector privado registrado, publico nacional y publico provincial. Las dos series publicas se encadenan con las variaciones mensuales del Cuadro 3 del Excel vigente del INDEC; las tres se deflactan con el IPC empalmado y se normalizan por separado con promedio enero-noviembre de 2023=100.
- Los cocientes sobre PIB usan el promedio de los saldos diarios de cada trimestre y el PIB nominal a precios corrientes del mismo trimestre.
- El grafico cambiario combina los limites inferior y superior del regimen de bandas con el tipo de cambio mayorista de referencia.
- La variacion de reservas por compra de divisas se presenta como promedio movil de cinco observaciones diarias.
- Las tasas de prestamos personales y de depositos a plazo fijo en pesos se comparan en un mismo grafico; la tasa de plazos fijos en dolares se muestra por separado.
- Los depositos totales en dolares corresponden a los sectores publico y privado no financieros y se muestran como saldo diario en millones de dolares, tal como los publica la serie 107 del BCRA.
- Exportaciones e importaciones mensuales se comparan en un mismo grafico; el saldo comercial se presenta por separado, en millones de dolares.
- Los indices mensuales de cantidades exportadas e importadas, base 2004=100, se presentan en graficos separados y se leen directamente del Excel que actualiza el INDEC.
- El grafico de actividad economica por sectores combina el EMAE general desestacionalizado con cinco sectores del Excel mensual del INDEC. Los sectores se ajustan cada dia con X-13ARIMA-SEATS (ajuste final X-11, tabla d11) y todas las lineas se expresan con noviembre de 2023=100. La leyenda muestra las ponderaciones del ano base 2004 redondeadas como en la referencia visual.
- El grafico fiscal usa el resultado primario sin rentas y el resultado financiero del Sector Publico Nacional, base caja y metodologia 2017. Para hacer comparables periodos con inflacion muy distinta, presenta la suma movil de 12 meses como porcentaje del PIB nominal: el denominador es el promedio de los niveles trimestrales anualizados correspondientes a esos 12 meses.

## Uso local

1. Ejecutar `python scripts/update_data.py`.
2. Servir la carpeta con un servidor web local.
3. Abrir `index.html` desde ese servidor.

## Fuente

- API: `https://apis.datos.gob.ar/series/api/series`
- API BCRA: `https://api.bcra.gob.ar/estadisticas/v4.0/Monetarias`
- Excel sectorial del EMAE: `https://www.indec.gob.ar/ftp/cuadros/economia/sh_emae_actividad_base2004.xls`
- IPC alternativo 1997-2016: `extender_ipc_1997/ipc_alternativo_1997_2016.csv`
- Indice de Salarios historico: `https://www.indec.gob.ar/ftp/cuadros/sociedad/serie_is_2012.xls`
- Indice de Salarios vigente: ultimo archivo mensual `variaciones_salarios_MM_AA.xls` publicado por el INDEC.
- Datos publicados por INDEC, BCRA y otros organismos oficiales.
