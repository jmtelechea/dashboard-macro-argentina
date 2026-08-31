# Etiquetas adaptativas del eje temporal

Correccion global de las etiquetas de fecha para evitar meses repetidos al reducir la ventana con el slider.

- Series diarias con una ventana de hasta 180 dias: dia y mes (`15 jun`).
- Series diarias con una ventana mas extensa: mes y ano (`jun 26`).
- Series mensuales, trimestrales, semestrales y anuales: mes y ano.
- Se conserva el limite de siete marcas visibles para evitar superposiciones.
