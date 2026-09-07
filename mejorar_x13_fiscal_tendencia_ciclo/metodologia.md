# Mejora del ajuste fiscal X-13

Los graficos de recaudacion total, ingresos fiscales y gasto primario
incorporan dos lineas:

- Serie desestacionalizada: salida final X-11 `d11`.
- Tendencia-ciclo: salida final X-11 `d12`.

El modelo X-13ARIMA-SEATS se estima sobre toda la historia real disponible.
La especificacion prueba efectos de dias habiles y Pascua mediante AIC e
incorpora deteccion automatica de valores atipicos. Esto mejora el ajuste
estacional sin confundir la volatilidad irregular de `d11` con estacionalidad
residual. La tendencia-ciclo `d12` ofrece la lectura suavizada solicitada.

La serie de subsidios a la energia no se modifica y permanece solamente
deflactada, sin ajuste estacional ni tendencia-ciclo.
