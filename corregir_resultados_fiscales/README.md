# Correccion de resultados fiscales

Carpeta de trabajo para diagnosticar y corregir las series de resultado primario y financiero del Sector Publico Nacional que alimentan el dashboard.

## Diagnostico

- El grafico anterior usaba la serie de superavit primario, que incluye rentas, en lugar del resultado primario sin rentas.
- La serie financiera estaba identificada como resultado antes de figurativos, aunque el catalogo ofrece el resultado financiero explicito.
- Los flujos mensuales nominales no eran comparables en el tiempo y comprimian visualmente los deficit antiguos por la inflacion.

## Correccion

- Resultado primario sin rentas: `379.9_RESULTADO_017__31_73`.
- Resultado financiero: `379.9_RESULTADO_017__18_38`.
- Presentacion: suma movil de 12 meses como porcentaje del PIB nominal.
- Validacion diciembre de 2015: resultado primario `-4,1%` del PIB y resultado financiero `-5,1%`.
