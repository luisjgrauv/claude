# Distribuidora Andina S.A. (FICTICIA) — Notas a los estados contables al 31/12/2026

Ejercicio económico iniciado el 01/01/2026 y finalizado el 31/12/2026, presentado en forma comparativa con el ejercicio 2025.
Cifras en pesos de **diciembre de 2026**, sin decimales. Datos ficticios, de práctica.

> **Advertencia sobre las cifras:** los índices de noviembre y diciembre de 2026 son **estimados**, no oficiales. Todas las cifras en moneda de dic-2026 deben recalcularse cuando se publiquen los índices oficiales. Ver la Nota 1.3.

Los anexos I a V están en `anexos_andina_2026.xlsx`. La tabla de control final concilia estas notas con los BSS ajustados.

## Nota 1 — Bases de preparación

### 1.1 Normas contables aplicadas
- Los estados contables se prepararon según las normas contables profesionales argentinas de la FACPCE.
- **Norma aplicada:** RT 54, Norma Unificada Argentina de Contabilidad (NUA), texto ordenado por la RT 59. Es de aplicación obligatoria para ejercicios iniciados desde el 01/01/2025. La fecha de vigencia la fija cada consejo profesional; la jurisdicción del ente es CABA (CPCECABA).
- La RT 54 derogó, entre otras, las RT 6, 8, 9, 17 y 41. Incorpora el procedimiento de reexpresión que antes estaba en la RT 6.
- **Ley 19.550 General de Sociedades (LGS):** arts. 62 a 65 (documentación contable), art. 64 (información del estado de resultados: Anexo IV) y art. 70 (reserva legal: Nota 9).
- La denominación "Bienes de uso" se mantiene por coherencia con el plan de cuentas. En la RT 54 el rubro se denomina "Propiedades, planta y equipo".
- **Duda:** confirmar (a) el número de la resolución del CPCECABA que adoptó la RT 54 y la RT 59 y (b) la numeración exacta de los párrafos de la RT 54 sobre unidad de medida, antes de citarlos en la versión final.

### 1.2 Unidad de medida (moneda homogénea)
- Los estados contables se expresan en moneda homogénea de diciembre de 2026. Se reconocen los efectos de la inflación según el procedimiento de reexpresión de la RT 54, que tomó el de la RT 6, y la Res. JG FACPCE 539/18 (modif. Res. JG 553/19 y 584/21).
- **Índice:** serie FACPCE de la Res. JG 539/18. Es el IPC Nacional del INDEC (base dic-2016 = 100) desde ene-2017, empalmado con el IPIM para los meses anteriores. Fuente: `insumos/indices/indices_facpce.csv`, consultado el 22/11/2026.
- **Contexto de inflación:** la variación acumulada de los últimos tres años (dic-2023 a dic-2026) es de 275,9 %. Supera la pauta cuantitativa del 100 % acumulado en tres años (RT 6, modif. RT 39, retomada por la RT 54; contexto declarado por la Res. JG 539/18).

| Concepto | dic-2024 | dic-2025 | dic-2026 |
|---|---:|---:|---:|
| Índice FACPCE (Res. JG 539/18) | 7.694,0075 | 10.121,3715 | 13.281,9696 (estimado) |
| Coeficiente del ejercicio (dic-previo → dic) | — | 1,315488 | 1,312270 |
| Inflación anual | — | 31,55 % (oficial; INDEC publicó 31,5 %) | 31,23 % (estimada) |

- **Coeficiente de reexpresión del comparativo:** cifras de 2025 en moneda de dic-2025 × 1,312270.
- **Método:**
  - Las partidas monetarias se exponen a su valor nominal.
  - Las partidas no monetarias se reexpresan desde su fecha de origen: bienes de uso bien por bien, mercaderías por capas FIFO, seguros por fecha de pago y patrimonio neto por componente.
  - Los resultados se reexpresan con el coeficiente de su mes de devengamiento. Los que consumen activos no monetarios (CMV, amortizaciones, seguros y venta de bienes de uso) se reexpresan desde su auxiliar.
  - El RECPAM se determina por diferencia y se controla por el método directo. La diferencia entre ambos es menor a $1.
- **RECPAM:** pérdida de $174.121.078 en 2026 y de $81.993.622 en 2025 (moneda dic-2026). Es pérdida porque la posición monetaria neta es activa: $702.443.075 al 31/12/2026 y $303.519.889 al 31/12/2025.

### 1.3 Estimaciones y cifras sujetas a revisión
- **Estimación — índices de nov-2026 y dic-2026.** Son una estimación propia: la media geométrica de las variaciones mensuales de ago a oct-2026 (1,9213 %). No son datos oficiales.
  - El índice de nov-2026 se publica el 16/12/2026. El de dic-2026, en enero de 2027.
  - Sensibilidad: si nov y dic repitieran el 2,17 % de octubre, el índice de dic-2026 sería alrededor de 0,5 % mayor. Todas las cifras en moneda de dic-2026 cambiarían aproximadamente en esa proporción.
  - Recomendación: regenerar el ajuste y estas notas con los índices oficiales antes de emitir los estados contables.
- Otras estimaciones: vidas útiles, previsión para deudores incobrables e impuesto a las ganancias.

### 1.4 Información comparativa
- Se presentan comparativos del ejercicio 2025 para el ESP, el estado de resultados, el EEPN, las notas y los anexos. Están reexpresados a moneda de dic-2026.
- **Limitación:** 2025 es el primer ejercicio en que se aplicó el ajuste por inflación en esta preparación.
  - Los resultados no asignados al 31/12/2024 se determinaron por diferencia entre activos y pasivos reexpresados y los demás componentes del PN.
  - El comparativo depende de los mismos supuestos de origen que los saldos de 2026: fechas de alta, capas de mercaderías y fechas de pago de las primas.
- El ESP al 31/12/2024 está en `comparativo_2025_en_moneda_dic2026.xlsx`. No se presenta en estas notas.

## Nota 2 — Principales criterios de medición

| Rubro | Criterio |
|---|---|
| Caja y bancos | Valor nominal. |
| Inversiones (plazo fijo) | Valor nominal más intereses devengados al cierre. Vencimiento original menor a 90 días: se considera equivalente de efectivo para el EFE. |
| Créditos por ventas | Valor nominal, neto de la previsión para deudores incobrables. No se segregan componentes financieros implícitos: Dato faltante sobre plazos de cobro; se asume que no son significativos. |
| Otros créditos | Anticipos e IVA a valor nominal. Seguros pagados por adelantado: parte no devengada de las primas, reexpresada desde el mes de pago. |
| Bienes de cambio | Costo de adquisición reexpresado, método FIFO por capas mensuales de compra. Dato faltante: prueba de valor neto de realización o costo de reposición. |
| Bienes de uso | Costo de adquisición reexpresado desde el mes de alta, menos amortización acumulada. Amortización lineal: año de alta completo y año de baja sin amortización. Vidas útiles: edificios 50 años, camiones 10 años, vehículos ejecutivos y camionetas 5 años. Los terrenos no se amortizan. Dato faltante: prueba de valor recuperable. |
| Deudas (comerciales, remuneraciones, cargas fiscales y préstamos) | Valor nominal más intereses devengados. El préstamo se separa en porción corriente y no corriente. |
| Impuesto a las ganancias | Método del impuesto corriente (impuesto determinado del ejercicio), según el supuesto simplificador del insumo. Ver la Duda de la Nota 8. |
| Patrimonio neto | Capital a su valor nominal. La reexpresión se expone en "Ajuste de capital". La reserva legal y los resultados no asignados se reexpresan desde su fecha de origen. |
| Resultados | Reexpresados desde el mes de devengamiento. Los resultados por consumo de activos no monetarios se reexpresan desde el auxiliar de esos activos. |

## Nota 3 — Caja y bancos

| Concepto | 31/12/2026 | 31/12/2025 |
|---|---:|---:|
| Caja | 10.000.000 | 10.498.158 |
| Banco Ficticio — cuenta corriente | 539.713.825 | 536.193.418 |
| **Total** | **549.713.825** | **546.691.576** |

## Nota 4 — Inversiones

| Concepto | 31/12/2026 | 31/12/2025 |
|---|---:|---:|
| Plazo fijo (vencimiento original < 90 días) | 270.000.000 | 288.699.344 |
| **Total corriente** | **270.000.000** | **288.699.344** |

Los intereses ganados en el ejercicio fueron $80.312.136 en 2026 y $81.523.008 en 2025, ambos en moneda de dic-2026 (cuenta 4.2.01).

## Nota 5 — Créditos por ventas

| Concepto | 31/12/2026 | 31/12/2025 |
|---|---:|---:|
| Deudores por ventas | 2.275.000.000 | 2.296.472.054 |
| Menos: previsión para deudores incobrables (Anexo II) | (70.500.000) | (68.894.162) |
| **Total corriente** | **2.204.500.000** | **2.227.577.893** |

Dato faltante: antigüedad de los créditos y plazos de cobro. Por esa razón no se presenta la apertura por vencimiento.

## Nota 6 — Otros créditos

| Concepto | 31/12/2026 | 31/12/2025 |
|---|---:|---:|
| Anticipos de impuesto a las ganancias | 350.000.000 | 196.840.462 |
| IVA saldo técnico a favor | 20.000.000 | 23.620.855 |
| Seguros pagados por adelantado (Anexo V) | 59.617.953 | 58.778.198 |
| Subtotal (saldos brutos) | 429.617.953 | 279.239.515 |
| Menos: anticipos de IG compensados con el impuesto a pagar (Nota 7.5) | (350.000.000) | (196.840.462) |
| **Total corriente (neto)** | **79.617.953** | **82.399.053** |

**Compensación:** los anticipos de impuesto a las ganancias se compensan con el impuesto determinado a pagar. Se trata del mismo impuesto y del mismo fisco, y el ente tiene el derecho legal de cancelar por el neto. Es el mismo criterio de presentación de los estados básicos (`eecc-2026`), definido por la reina el 22/11/2026. En el BSS los saldos siguen brutos: la compensación solo afecta la exposición.

## Nota 7 — Bienes de cambio, bienes de uso y deudas

### 7.1 Bienes de cambio

| Concepto | 31/12/2026 | 31/12/2025 |
|---|---:|---:|
| Mercaderías de reventa (insumos), FIFO | 1.632.937.362 | 1.652.008.929 |
| **Total corriente** | **1.632.937.362** | **1.652.008.929** |

La existencia al 31/12/2026 corresponde a compras de noviembre ($413.120.000 nominales) y diciembre de 2026 ($1.211.880.000 nominales). El costo de ventas está en el Anexo III.

### 7.2 Bienes de uso (detalle en el Anexo I y en `insumos/andina/auxiliares_2026.xlsx`, hoja Bienes de uso)

| Rubro | VO 31/12/2026 | AA 31/12/2026 | Neto 31/12/2026 | Neto 31/12/2025 |
|---|---:|---:|---:|---:|
| Terrenos | 1.357.370.288 | 0 | 1.357.370.288 | 1.357.370.288 |
| Edificios | 3.589.772.809 | (622.410.132) | 2.967.362.678 | 3.039.158.134 |
| Rodados — camiones | 2.298.609.677 | (1.112.107.148) | 1.186.502.529 | 1.265.021.431 |
| Rodados — vehículos ejecutivos | 406.437.365 | (307.694.385) | 98.742.980 | 162.111.347 |
| Rodados — camionetas de reparto | 107.540.247 | (86.135.011) | 21.405.236 | 42.913.285 |
| **Total no corriente** | **7.759.730.387** | **(2.128.346.676)** | **5.631.383.710** | **5.866.574.485** |

Composición al 31/12/2026, según el auxiliar:
- **Inmuebles (2):** INM-01, sede administrativa (alta 03/2016), e INM-02, planta / centro de distribución (alta 07/2019). Cada uno se registra con terreno y edificio por separado.
- **Camiones (10):** C02 a C11.
  - Baja: C01, vendido el 15/06/2026 por $28.000.000 + IVA. Su resultado reexpresado es una **pérdida de $(13.750.256)**: el valor residual reexpresado desde may-2018 supera el precio de venta reexpresado. En valores históricos había sido una ganancia de $27.525.000.
  - Alta: C11, adquirido el 15/06/2026 por $175.000.000 nominales.
- **Vehículos ejecutivos (5):** E01 a E05. E01 y E02 están totalmente amortizados al cierre.
- **Camionetas de reparto (2):** U01, totalmente amortizada, y U02.
- **Amortizaciones del ejercicio:** $386.532.840 en 2026 y $407.396.305 en 2025. Su imputación por función está en el Anexo IV.
- **Dato faltante — valor recuperable:** el valor residual reexpresado ($5.631.383.710) multiplica por 10,6 el valor residual histórico ($530.189.000).
  - Los inmuebles se reexpresan con coeficientes de ×153 (mar-2016) y ×58 (jul-2019).
  - La RT 54 exige comparar con el valor recuperable cuando hay indicios de deterioro. No se recibieron tasaciones ni proyecciones, por lo que no se pudo hacer la prueba.
  - Si el valor recuperable fuera menor, correspondería una desvalorización con cargo a resultados.

### 7.3 Deudas comerciales

| Concepto | 31/12/2026 | 31/12/2025 |
|---|---:|---:|
| Proveedores | 1.495.000.000 | 1.509.110.207 |
| **Total corriente** | **1.495.000.000** | **1.509.110.207** |

### 7.4 Remuneraciones y cargas sociales

| Concepto | 31/12/2026 | 31/12/2025 |
|---|---:|---:|
| Sueldos y cargas sociales a pagar | 169.000.000 | 170.595.067 |
| **Total corriente** | **169.000.000** | **170.595.067** |

### 7.5 Cargas fiscales

| Concepto | 31/12/2026 | 31/12/2025 |
|---|---:|---:|
| Impuesto a las ganancias determinado (Nota 8) | 673.170.750 | 654.568.252 |
| Menos: anticipos de impuesto a las ganancias compensados (Nota 6) | (350.000.000) | (196.840.462) |
| Impuesto a las ganancias a pagar, neto | 323.170.750 | 457.727.790 |
| Impuesto sobre los ingresos brutos a pagar | 54.600.000 | 55.115.329 |
| **Total corriente** | **377.770.750** | **512.843.120** |

### 7.6 Préstamos

| Concepto | 31/12/2026 | 31/12/2025 |
|---|---:|---:|
| Préstamo bancario — porción corriente | 150.000.000 | 196.840.462 |
| Préstamo bancario — porción no corriente | 150.000.000 | 393.680.924 |
| **Total** | **300.000.000** | **590.521.385** |

- Préstamo bancario a tasa fija. En 2026 se amortizaron $150.000.000 de capital, en cuotas trimestrales de $37.500.000. Saldo nominal al 31/12/2026: $300.000.000.
- Intereses devengados: $118.601.067 en 2026 y $179.756.770 en 2025, en moneda de dic-2026.
- **Supuesto (S):** la porción corriente es de $150.000.000 nominales por año. Al 31/12/2026: $150.000.000 vencen en 2027 y $150.000.000 en 2028.
- **Dato faltante:** cronograma contractual, tasa y garantías.

## Nota 8 — Impuesto a las ganancias

**Conclusión:** el cargo del ejercicio ($673.170.750) no equivale al 35 % del resultado contable ajustado. El impuesto se determinó sobre el resultado histórico, con el supuesto simplificador del insumo. La diferencia es el efecto del ajuste por inflación contable sobre la base del cálculo.

| Concepto | 2026 | 2025 |
|---|---:|---:|
| Resultado antes del impuesto (moneda dic-2026) | 1.169.332.194 | 1.266.076.619 |
| Tasa | 35 % | 35 % |
| Impuesto a la tasa sobre el resultado contable ajustado | 409.266.268 | 443.126.817 |
| Efecto del ajuste por inflación contable no reconocido en la base del impuesto (resultado histórico − ajustado, × 35 %) | 263.904.482 | 211.441.436 |
| **Cargo por impuesto a las ganancias (estado de resultados)** | **673.170.750** | **654.568.252** |
| Tasa efectiva sobre el resultado ajustado | 57,6 % | 51,7 % |

Control:
- Resultado histórico antes del impuesto 2026: $1.923.345.000 × 35 % = $673.170.750, igual al cargo.
- Resultado histórico antes del impuesto 2025: $1.425.160.500 × 35 % = $498.806.175 en moneda dic-2025, que reexpresado da $654.568.252.

Pasivo al 31/12/2026: impuesto determinado $673.170.750, menos anticipos pagados en 2026 $350.000.000, igual a $323.170.750 neto, expuesto en Cargas fiscales (Notas 6 y 7.5).

**Dudas y limitaciones (requieren validación profesional):**
1. **Impuesto diferido.** La RT 54 requiere, como regla general, el método del impuesto diferido. La reexpresión de bienes de uso, bienes de cambio y seguros genera diferencias temporarias con sus valores fiscales.
   - Esas diferencias darían lugar a un pasivo por impuesto diferido que no se reconoce, por el supuesto simplificador del insumo.
   - Hay que confirmar si el ente puede usar alguna opción de la RT 54 para entes pequeños o medianos. Si no puede, hay que calcular el diferido.
2. **Tasa.** El 35 % es un supuesto del insumo. La ley del impuesto (Ley 27.630) establece una escala de alícuotas del 25 %, 30 % y 35 % por tramos de ganancia neta imponible, con montos que se actualizan cada año. Hay que verificar los tramos vigentes para 2026.
3. **Ajuste por inflación impositivo** (Título VI de la ley del impuesto). No se modeló. El impuesto real puede diferir del cargo expuesto.

## Nota 9 — Patrimonio neto, reserva legal y restricciones a la distribución de resultados

### 9.1 Capital
- Capital suscripto e integrado: $20.000.000 nominales, integrado el 14/03/2014.
- Ajuste de capital: $4.735.227.863. Total reexpresado: $4.755.227.863.
- **Dato faltante:** clases de acciones, valor nominal y votos (datos adicionales 2026).

### 9.2 Distribución aprobada en el ejercicio
- La asamblea ordinaria del 28/04/2026 aprobó dividendos en efectivo por $300.000.000 nominales sobre el resultado 2025. Se pagaron el 30/04/2026.
- Reexpresados a moneda de dic-2026 suman $350.660.845. La asamblea no asignó importes a la reserva legal.

### 9.3 Reserva legal (art. 70 LGS)
- El art. 70 de la LGS obliga a destinar a la reserva legal al menos el 5 % de las ganancias realizadas y líquidas de cada ejercicio, hasta alcanzar el 20 % del capital social.
- En valores nominales, la reserva ($4.000.000) alcanzó en 04/2022 el 20 % del capital suscripto nominal ($20.000.000).
- **Duda (societaria):** no está definido si la base del 20 % debe incluir el ajuste de capital.
  - Si la base es capital + ajuste de capital, en moneda homogénea la reserva queda por debajo del tope.
  - Esa es la interpretación propuesta en la fase 2, que siguen, entre otros, el CPCE de Córdoba y la IPJ de Córdoba (RG 95/T).
  - Para CABA hay que validar el criterio de la IGJ con un asesor societario. La RG IGJ 7/2015 fue reemplazada por la RG IGJ 15/2024: verificar qué dice la norma vigente.

| Concepto (moneda dic-2026) | Importe |
|---|---:|
| Base: capital suscripto + ajuste de capital | 4.755.227.863 |
| Tope: 20 % de la base | 951.045.573 |
| Reserva legal reexpresada al 31/12/2026 | 220.012.066 |
| Faltante para alcanzar el tope | 731.033.507 |
| 5 % del resultado 2025 ajustado, no asignado por la asamblea del 28/04/2026 | 30.575.418 |
| 5 % del resultado 2026 ajustado | 24.808.072 |
| **Asignación mínima propuesta a la asamblea que trate los EECC 2026** | **55.383.491** |

- **No se contabilizó ninguna asignación.** Asignar la reserva es una decisión de la asamblea. Es una reclasificación dentro del PN y no cambia su total.
- **Restricción a la distribución:** con el criterio propuesto, de los resultados no asignados al 31/12/2026 ($3.051.142.171) deben asignarse a reserva legal al menos $55.383.491 antes de distribuir dividendos. Esa cifra es el mínimo legal; el art. 70 no impide asignar más.
- **Riesgo a validar:** la distribución de $300.000.000 de dividendos sin asignar el 5 % del resultado 2025 ajustado. Si se adopta la base con ajuste, la asamblea podría tener que regularizarlo.
- Los dividendos solo pueden distribuirse de ganancias realizadas y líquidas, según un balance confeccionado conforme a la ley y aprobado por la asamblea (art. 68 LGS).

## Nota 10 — Seguros contratados

- Al 31/12/2026 hay 10 pólizas vigentes con Aseguradora Ficticia S.A. Cubren automotores, incendio, robo, responsabilidad civil y riesgos de la naturaleza. El detalle de bienes, coberturas y vigencias está en el Anexo V.
- Las primas renovadas en 2026 suman $184.680.000 nominales ($215.125.357 en moneda dic-2026).
- Gasto de seguros 2026: $214.285.602. Comercialización $183.654.163; administración $30.631.439.
- Parte no devengada al cierre: $59.617.953 (Nota 6).
- **Dato faltante:** sumas aseguradas y franquicias. Sin esos datos no se puede evaluar si la cobertura es suficiente frente al valor de los bienes.

## Nota 11 — Hechos posteriores al cierre

- Entre el 31/12/2026 y la fecha de emisión no hay hechos posteriores informados que requieran ajustar o revelar estos estados contables.
- **Dato faltante:** fecha de emisión y confirmación de la gerencia a esa fecha.
- **Hecho conocido después del cierre que afecta las estimaciones:** la publicación de los índices oficiales de noviembre (16/12/2026) y diciembre de 2026 (enero de 2027). Si ocurre antes de la emisión, corresponde recalcular las cifras en moneda de dic-2026 (Nota 1.3).

## Tabla de control — notas y anexos contra los BSS ajustados

Fuentes:
- Columna BSS: `entregables/ajuste/ajuste_2026.xlsx`, hoja BSS ajustado; el 31/12/2025 es el saldo inicial en moneda dic-2026. En Otros créditos, Cargas fiscales y los totales, el BSS se muestra neto de la compensación de anticipos de IG (Nota 6).
- Estado de resultados 2025: `comparativo_2025_en_moneda_dic2026.xlsx`.
- Columna eecc: `entregables/eecc-2026/eecc_andina_2026.xlsx`, hojas ESP y ER (estados básicos de la sesión eecc-2026).
- Diferencia aceptada: menor a $1.

### Estado de situación patrimonial

| Rubro | Nota 31/12/2026 | BSS 31/12/2026 | eecc 31/12/2026 | Nota 31/12/2025 | BSS 31/12/2025 | eecc 31/12/2025 | Resultado |
|---|---:|---:|---:|---:|---:|---:|---|
| Caja y bancos | 549.713.825 | 549.713.825 | 549.713.825 | 546.691.576 | 546.691.576 | 546.691.576 | OK |
| Inversiones | 270.000.000 | 270.000.000 | 270.000.000 | 288.699.344 | 288.699.344 | 288.699.344 | OK |
| Créditos por ventas | 2.204.500.000 | 2.204.500.000 | 2.204.500.000 | 2.227.577.893 | 2.227.577.893 | 2.227.577.893 | OK |
| Otros créditos (neto de compensación) | 79.617.953 | 79.617.953 | 79.617.953 | 82.399.053 | 82.399.053 | 82.399.053 | OK |
| Bienes de cambio | 1.632.937.362 | 1.632.937.362 | 1.632.937.362 | 1.652.008.929 | 1.652.008.929 | 1.652.008.929 | OK |
| Bienes de uso (Anexo I) | 5.631.383.710 | 5.631.383.710 | 5.631.383.710 | 5.866.574.485 | 5.866.574.485 | 5.866.574.485 | OK |
| Deudas comerciales | 1.495.000.000 | 1.495.000.000 | 1.495.000.000 | 1.509.110.207 | 1.509.110.207 | 1.509.110.207 | OK |
| Remuneraciones y cargas sociales | 169.000.000 | 169.000.000 | 169.000.000 | 170.595.067 | 170.595.067 | 170.595.067 | OK |
| Cargas fiscales (neto de compensación) | 377.770.750 | 377.770.750 | 377.770.750 | 512.843.120 | 512.843.120 | 512.843.120 | OK |
| Préstamos (corriente + no corriente) | 300.000.000 | 300.000.000 | 300.000.000 | 590.521.385 | 590.521.385 | 590.521.385 | OK |
| **Total activo** | 10.368.152.850 | 10.368.152.850 | 10.368.152.850 | 10.663.951.279 | 10.663.951.279 | 10.663.951.279 | OK |
| **Total pasivo** | 2.341.770.750 | 2.341.770.750 | 2.341.770.750 | 2.783.069.779 | 2.783.069.779 | 2.783.069.779 | OK |
| **Patrimonio neto (activo − pasivo)** | 8.026.382.100 | 8.026.382.100 | 8.026.382.100 | 7.880.881.500 | 7.880.881.500 | 7.880.881.500 | OK |

### Estado de resultados (egresos con signo positivo; resultado del ejercicio: ganancia positiva)

| Concepto | Nota / anexo 2026 | BSS 2026 | eecc 2026 | Nota / anexo 2025 | Comparativo 2025 | eecc 2025 | Resultado |
|---|---:|---:|---:|---:|---:|---:|---|
| Costo de mercaderías vendidas (Anexo III) | 14.648.408.412 | 14.648.408.412 | 14.648.408.412 | 14.955.044.768 | 14.955.044.768 | 14.955.044.768 | OK |
| Gastos de comercialización (Anexo IV) | 3.173.091.074 | 3.173.091.074 | 3.173.091.074 | 3.228.372.658 | 3.228.372.658 | 3.228.372.658 | OK |
| Gastos de administración (Anexo IV) | 1.081.445.216 | 1.081.445.216 | 1.081.445.216 | 1.119.537.023 | 1.119.537.023 | 1.119.537.023 | OK |
| Intereses del préstamo (Anexo IV) | 118.601.067 | 118.601.067 | 118.601.067 | 179.756.770 | 179.756.770 | 179.756.770 | OK |
| Cargo por incobrables = aumentos de la previsión (Anexo II) | 43.490.830 | 43.490.830 | — | 44.454.766 | 44.454.766 | — | OK |
| Amortizaciones (Anexo I) = cuentas de amortización | 386.532.840 | 386.532.840 | — | 407.396.305 | 407.396.305 | — | OK |
| Impuesto a las ganancias (Nota 8) | 673.170.750 | 673.170.750 | 673.170.750 | 654.568.252 | 654.568.252 | 654.568.252 | OK |
| **Resultado del ejercicio** | 496.161.444 | 496.161.444 | 496.161.444 | 611.508.366 | 611.508.366 | 611.508.366 | OK |

### Otros controles

| Control | Nota / anexo | Fuente | Dif. | Resultado |
|---|---:|---:|---:|---|
| Compras históricas 2026 (cifra de control del encargo eecc: $13.115.000.000) | 13.115.000.000 | 13.115.000.000 | 0 | OK |
| Anexo V: gasto de seguros 2026 = 5.2.06 + 5.3.05 | 214.285.602 | 214.285.602 | 0 | OK |
| Anexo V: no devengado 31/12/2026 = 1.4.03 | 59.617.953 | 59.617.953 | 0 | OK |
| Previsión: saldo final 2026 (Anexo II) = 1.3.02 | 70.500.000 | 70.500.000 | 0 | OK |
| Nota 8: impuesto a la tasa + efecto ajuste = cargo 2026 | 673.170.750 | 673.170.750 | 0 | OK |
| Nota 8: resultado histórico antes de IG 2026 × 35 % = cargo | 673.170.750 | 673.170.750 | 0 | OK |
| Nota 9: RNA 31/12/2026 = RNA inicio − dividendos + resultado | 3.051.142.171 | 3.051.142.171 | 0 | OK |

**Resultado global de los controles: OK.** La hoja Controles de `anexos_andina_2026.xlsx` repite los cruces con fórmulas.

**Cruce con `eecc-2026`: OK.** Los rubros y totales del ESP y del ER de las notas coinciden con `eecc_andina_2026.xlsx` en 2026 y en 2025 (columnas eecc).

## Supuestos, Dudas y datos faltantes (resumen)

| Tipo | Tema | Nota |
|---|---|---|
| Estimación | Índices de nov y dic-2026 no oficiales: afectan todas las cifras en moneda de dic-2026 | 1.3 |
| Duda | Número de la resolución de adopción de la RT 54 / RT 59 en el CPCECABA y párrafos de la RT 54 sobre unidad de medida | 1.1 |
| Duda | Base de la reserva legal (art. 70 LGS) con o sin ajuste de capital: faltante de $731.033.507 y asignación mínima propuesta de $55.383.491 | 9.3 |
| Duda | Impuesto diferido no reconocido (supuesto del insumo) y escala de alícuotas de la Ley 27.630 | 8 |
| Dato faltante | Prueba de valor recuperable de bienes de uso y de VNR de bienes de cambio | 2, 7.1, 7.2 |
| Dato faltante | Fechas de aplicación de la previsión para incobrables (supuesto lineal mensual) | Anexo II |
| Dato faltante | Cronograma, tasa y garantías del préstamo | 7.6 |
| Dato faltante | Sumas aseguradas y franquicias | 10 |
| Dato faltante | Antigüedad de los créditos; acciones (clases y votos); fecha de emisión y hechos posteriores | 5, 9.1, 11 |

---
Generado con `entregables/notas-anexos-2026/generador_notas_anexos.py` y verificado con `verificar_notas_anexos.py` (Python + LibreOffice).
