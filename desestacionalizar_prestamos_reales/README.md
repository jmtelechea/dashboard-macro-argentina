# Prestamos reales desestacionalizados

Carpeta de trabajo para reemplazar la serie diaria del grafico de prestamos reales por su version mensual desestacionalizada.

- Serie fuente: prestamos al sector privado, BCRA 1341.
- Deflactor: IPC empalmado del dashboard.
- Agregacion: promedio mensual de saldos diarios reales.
- Ajuste estacional: X-13ARIMA-SEATS, ajuste final X-11, tabla d11.
- Titulo: `Prestamos al sector privado real`.
- Subtitulo: `En millones de pesos constantes`.
- La serie original no se muestra en el grafico.
