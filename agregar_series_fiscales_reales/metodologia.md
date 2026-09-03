# Series fiscales reales

Esta tarea incorpora cuatro series mensuales al grupo Fiscal:

- Recaudacion total real y desestacionalizada.
- Ingresos fiscales reales, concatenados y desestacionalizados.
- Gasto primario real, concatenado y desestacionalizado.
- Subsidios a la energia en terminos reales, sin desestacionalizar.

Los niveles nominales se deflactan con el IPC empalmado del dashboard,
normalizado de modo que el promedio de 2014 sea igual a 100. Los tramos
fiscales de 1993-2006, 2007-2014 y desde 2015 se concatenan sin modificar
los saltos de enero de 2007 y enero de 2015. X-13ARIMA-SEATS se aplica
sobre la historia real completa de las tres series desestacionalizadas y
se utiliza la salida X-11 `d11`.

Aunque los tramos nominales de ingresos y gasto comienzan en 1993, las
series reales comienzan en enero de 1997 porque ese es el primer mes con
cobertura del IPC empalmado disponible en el dashboard.
