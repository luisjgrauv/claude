# Distribuidora Andina S.A. (ficticia) — Índice de entregables

Estado: **completo**, integrado en la rama `claude/armado-de-reina-cqmydu` el 22/11/2026.

## Lectura recomendada
1. `entregables/eecc-2026/eecc_andina_2026.xlsx`: ESP, ER, EEPN y EFE 2026, comparativos con 2025, en moneda de dic-2026.
2. `entregables/notas-anexos-2026/notas_andina_2026.md` y `anexos_andina_2026.xlsx`: 11 notas y los anexos I a V.
3. `entregables/ajuste/resumen.md`: método, coeficientes, RECPAM y Dudas del ajuste.

## Cadena completa
| Etapa | Archivos |
|---|---|
| Histórico 2025 | `insumos/andina/bss_2025.*`, `auxiliares_2025.xlsx`, `auxiliares_complementarios_2025.xlsx` |
| Histórico 2026 | `insumos/andina/bss_2026.*`, `auxiliares_2026.xlsx`, `entregables/historicos/eecc_historicos_2025_2026.xlsx` |
| Índices | `insumos/indices/indices_facpce.csv` y `README.md` |
| Ajuste por inflación | `entregables/ajuste/ajuste_2025.xlsx`, `ajuste_2026.xlsx`, `comparativo_2025_en_moneda_dic2026.xlsx` |
| EECC 2026 | `entregables/eecc-2026/` y `entregables/notas-anexos-2026/` |

Cada etapa tiene un script generador y otro verificador (`*.py`). Para regenerar con los índices definitivos de nov y dic-2026: actualizar el CSV de índices y correr los generadores en orden (ajuste → eecc → notas).

## Cifras clave
| Moneda dic-2026 | 2026 | 2025 |
|---|---|---|
| Activo | 10.368.152.850 | 10.663.951.279 |
| Pasivo | 2.341.770.750 | 2.783.069.779 |
| Patrimonio neto | 8.026.382.100 | 7.880.881.500 |
| Resultado del ejercicio | 496.161.444 | 611.508.366 |
| RECPAM | (174.121.078) | — |
| Resultado histórico (nominal) | 1.250.174.250 | 926.354.325 |

## Dudas abiertas (requieren validación profesional)
1. **Reserva legal (art. 70 LGS):** si la base es capital + ajuste, el faltante a dic-26 es de $731,0 M. No se contabilizó.
2. **Impuesto a las ganancias:**
   - El cargo usa una tasa plana del 35 % sobre el resultado histórico. La ley prevé una escala del 25 %, 30 % y 35 %.
   - No se evaluó el ajuste por inflación impositivo.
   - No se reconoció impuesto diferido.
3. **Índices de nov y dic-2026:** son estimados. Hay que regenerar el ajuste cuando se publiquen (nov-26: 16/12/2026).
4. **RT 54 (t.o. RT 59):** falta confirmar los números de párrafo y el alcance de la derogación de la RT 9 contra el PDF oficial de la FACPCE.
5. **Datos faltantes:** valor recuperable / VNR, cronograma del préstamo, saldos mensuales de efectivo para el RECPAM del EFE y sumas aseguradas.
