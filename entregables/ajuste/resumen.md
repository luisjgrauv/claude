# Ajuste por inflación — Distribuidora Andina S.A. (FICTICIA)

**Conclusión:** el ajuste reduce el resultado del ejercicio a la mitad en ambos años. 2025: de $926,4 M (histórico) a **$466,0 M** (moneda de dic-2025). 2026: de $1.250,2 M a **$496,2 M** (moneda de dic-2026). El RECPAM es pérdida en los dos ejercicios, porque la posición monetaria neta es activa y crece. Los tres libros cuadran. Los 27 controles independientes (Python + LibreOffice) dan OK.

## 1. Resultados
| Concepto | 2025 (moneda dic-2025) | 2026 (moneda dic-2026) |
|---|---:|---:|
| Resultado histórico | 926.354.325 | 1.250.174.250 |
| Resultado antes del RECPAM (ajustado) | 528.475.179 | 670.282.523 |
| RECPAM por diferencia | **(62.482.293)** | **(174.121.078)** |
| RECPAM método directo (control) | (62.482.293) | (174.121.078) |
| **Resultado del ejercicio ajustado** | **465.992.887** | **496.161.444** |
| PN ajustado al inicio | 5.681.007.004 | 7.880.881.500 |
| PN ajustado al cierre | 6.005.534.707 | 8.026.382.100 |
| Coeficiente dic-previo → cierre | 1,315488 | 1,312270 (estimado) |

Comparativo 2025 en moneda de dic-2026: resultado **$611.508.366**; PN al 31/12/2025 **$7.880.881.500**, igual al PN inicial de 2026.

**Por qué cae el resultado (inferencia sobre los datos):**
- El CMV sale de compras de meses anteriores. Se reexpresa con coeficientes más altos que las ventas. El margen bruto baja del 30,0 % al 27,9 % en 2025.
- Las amortizaciones se calculan sobre bienes de 2016–2025 reexpresados. Por ejemplo, los coeficientes a dic-2025 son ×116,2 (mar-2016) y ×43,9 (jul-2019).
- El impuesto a las ganancias está imputado en diciembre (coeficiente 1). No baja en la misma proporción que el resultado.
- RECPAM: la posición monetaria neta pasa de −$78 M a +$231 M en 2025 y de +$231 M a +$702 M en 2026.

## 2. Norma y método
- **Norma:** RT 6 FACPCE (modif. RT 19, 27 y 39) y Res. JG 539/18 (modif. Res. JG 553/19 y 584/21). La RT 54 (NUA), vigente para ejercicios iniciados desde el 01/01/2025, incorpora el procedimiento de la RT 6. **Duda:** confirmar la numeración exacta de los párrafos de la RT 54 (t.o. RT 59) sobre unidad de medida antes de citarla en las notas.
- **Índice:** serie FACPCE Res. JG 539/18 (`insumos/indices/indices_facpce.csv`). Coeficiente = índice del mes de cierre / índice del mes de origen.
- **Partidas monetarias:** valor nominal al cierre. Saldo inicial × coeficiente dic-previo → cierre.
- **Bienes de uso:** bien por bien. VO, amortización acumulada y amortización del ejercicio × coeficiente del mes de alta. La baja del C01 se calcula como precio × coeficiente de jun-2026 − VRC reexpresado. Resultado: **pérdida de $13,8 M** ajustada contra una ganancia de $27,5 M histórica. El VRC reexpresado desde may-2018 supera el precio de venta.
- **Seguros:** no devengado × coeficiente del mes de pago. Gasto = devengado de cada prima × su coeficiente.
- **Mercaderías (FIFO):** capas × coeficiente del mes de compra. CMV con la matriz FIFO por mes de compra.
- **PN:** capital nominal ($20 M) y ajuste de capital por separado (coeficiente de mar-2014). Reserva legal por fecha de cada asamblea. Dividendos × coeficiente del mes de la asamblea.
- **RNA al inicio de 2025:** por diferencia entre activo y pasivo reexpresados y los demás componentes del PN (primer ejercicio de aplicación). Como dato informativo, el libro reexpresa también las capas de RNA por año y muestra la diferencia, que no se contabiliza.
- **RNA al inicio de 2026:** cierre 2025 ajustado (con el resultado 2025) × 1,312270.
- **Resultados:** importe del mes × coeficiente del mes. CMV, amortizaciones, seguros y venta de bienes de uso salen de su auxiliar.
- **RECPAM:** por diferencia. Es la cuenta de cuadre del BSS ajustado: resultado total (ΔPN ajustado + dividendos) − resultados reexpresados.

## 3. RECPAM: diferencia entre los dos métodos (criterio de aceptación)
- **Método directo:** PMN al inicio de cada mes × inflación del mes × coeficiente mes → cierre. Los flujos monetarios mensuales salen de los mismos auxiliares: resultados monetarios, compras, altas y venta de bienes de uso, primas y dividendos.
- **Diferencia:** < $0,01 en ambos años. **Explicación:** con flujos fechados en el mismo mes que la partida no monetaria de contrapartida, los dos métodos son algebraicamente idénticos: Σ PMNₘ₋₁ × inflaciónₘ × coefₘ = PMN cierre − PMN inicio × coef − Σ flujoₘ × coefₘ. Por eso la igualdad no prueba que las fechas de origen sean correctas. Sí prueba que no falta ninguna partida y que la clasificación monetaria/no monetaria es coherente. Además, la PMN de cierre que se arma con los flujos es igual a la del BSS.
- **Estimación simplificada (orden de magnitud):** −PMN promedio × inflación anual. La hoja RECPAM, bloque C, la muestra. Su diferencia con el método por diferencia viene de la distribución de la inflación y de la PMN dentro del año.

**Confirmación de independencia (pedida por la reina).** El método directo es un cálculo independiente. No usa ninguna cifra del cálculo por diferencia.
- **Datos de entrada (hoja RECPAM, bloque B):**
  - PMN de enero = suma de los saldos iniciales monetarios del BSS histórico.
  - Cada mes, PMN fin = PMN inicio + flujo monetario neto del mes. El flujo sale de las filas mensuales históricas de la hoja Resultados (ingresos y gastos monetarios), de las compras (Mercaderías), de las altas y la venta de bienes de uso (Bienes de uso), de las primas pagadas (Seguros) y de los dividendos (Patrimonio neto).
  - RECPAM del mes = −PMN al inicio del mes × (índiceₘ / índiceₘ₋₁ − 1) × (índice de cierre / índiceₘ).
- **Qué no usa:** el saldo ajustado del BSS, el PN ajustado ni la celda del RECPAM por diferencia.
- **Control cruzado:** la PMN de diciembre armada con los flujos coincide con la del BSS de cierre.
- **Recálculo propio:** `verificar_ajuste.py` repite el directo en Python, leyendo directamente los insumos.

| Año | RECPAM por diferencia | RECPAM método directo | Diferencia |
|---|---:|---:|---:|
| 2025 | −62.482.292,5062761 | −62.482.292,5062747 | −0,0000014 |
| 2026 | −174.121.078,291697 | −174.121.078,2917 | +0,0000027 |

**Por qué la diferencia es solo error de coma flotante (≈10⁻⁶ $):**
- Cada partida devengada mensualmente se reexpresa con el coeficiente del mes en que genera su flujo monetario. Por eso no se espera una diferencia económica.
- Las partidas con desfase (CMV, amortizaciones y seguros) no generan flujo monetario en el mes en que se devengan. Su flujo monetario es la compra, el alta o el pago de la prima, y entra en el directo con ese mes. Es el mismo coeficiente con que el método por diferencia reexpresa la partida no monetaria. Por eso tampoco generan diferencia.
- Una diferencia aparecería solo si se cambiara la convención de fechas, por ejemplo con un flujo a mitad de mes o un índice de cierre de mes en lugar del índice promedio. Como referencia, la estimación simplificada (PMN promedio × inflación anual) da −64.648.090 en 2025 (dif. 2.165.797) y −191.970.780 en 2026 (dif. 17.849.702).

## 4. Entregables
| Archivo | Contenido |
|---|---|
| `ajuste_2025.xlsx` | Notas, Controles, Coeficientes, BSS histórico, Bienes de uso, Seguros, Mercaderías, Patrimonio neto, Resultados, RECPAM, BSS ajustado, Conciliación |
| `ajuste_2026.xlsx` | La misma estructura, más la hoja **Inicio** (cierre 2025 ajustado × 1,312270) y la conciliación de los saldos iniciales contra los auxiliares desde el origen |
| `comparativo_2025_en_moneda_dic2026.xlsx` | ESP (31/12/2025 y 31/12/2024), ER y EEPN 2025 en moneda de dic-2026, con controles |
| `generador_ajuste.py` / `verificar_ajuste.py` | Generación reproducible y verificación (Python + LibreOffice 24.2) |

Las fórmulas están vinculadas. El azul indica un insumo y el negro una fórmula. Los libros usan INDEX/MATCH y SUMPRODUCT en lugar de XLOOKUP o LET, porque deben recalcularse con LibreOffice 24.2. Los saldos de apertura de 2026 y los datos del comparativo son **valores importados** de `ajuste_2025.xlsx` recalculado. El verificador comprueba que coinciden.

## 5. Controles (todos OK)
- BSS ajustados: inicio y cierre cuadran en 2025 y 2026.
- PN 31/12/2025 ajustado × 1,312270 = PN inicial 2026 ($7.880.881.500). Coincide también con el comparativo.
- Auxiliares ajustados = BSS ajustado: Σ|dif.| = 0. Saldos iniciales 2026 desde el origen = Inicio.
- RECPAM por diferencia = método directo. PMN de cierre por flujos = BSS.
- Recálculo independiente en Python de resultados, RECPAM y PN: coincide con LibreOffice.
- Sin importes fijos en las columnas ajustadas.

## 6. Dudas y datos faltantes
1. **Duda — Reserva legal (art. 70 LGS).** *Propuesta:* la base del 20 % es capital suscripto + ajuste de capital. Es la interpretación habitual y debe validarla el asesor societario con el criterio de la IGJ.
   - En moneda homogénea la reserva reexpresada ($167,7 M a dic-2025; $220,0 M a dic-2026) queda por debajo del tope: $724,7 M y $951,0 M respectivamente. El faltante es de $557,1 M y $731,0 M.
   - La asamblea del 28/04/2026 no asignó el 5 % del resultado 2025 ajustado ($30,6 M en moneda dic-2026) y distribuyó $300 M de dividendos.
   - **No se contabiliza** ninguna asignación: requiere decisión de asamblea y es una reclasificación dentro del PN, sin efecto en el total.
   - Se propone informarlo en nota y que la asamblea de 2027 asigne hasta $55,4 M: el 5 % omitido de 2025 más el 5 % de 2026.
   - Riesgo societario a validar: la distribución de dividendos sin reconstituir la reserva (art. 70, 2.º párrafo).
2. **Estimación — Índices de nov-2026 y dic-2026** (no oficiales). Afectan todo lo expresado en moneda de dic-2026: ajuste 2026 y comparativo. Hay que regenerar con `generador_ajuste.py` cuando se publiquen (16/12/2026 y ene-2027).
3. **Dato faltante — Cronograma del préstamo.** Se mantiene el supuesto trimestral lineal. Solo afecta la distribución mensual de intereses y, en forma marginal, el RECPAM.
4. **Supuestos (S):**
   - La póliza anterior se pagó el mismo día del año previo.
   - Se usa el índice del mes de origen, sin prorrateo diario.
   - No se hizo prueba de valor recuperable ni de VNR. Bienes de uso con valor reexpresado de $5.866,6 M en moneda dic-2026 (31/12/2025) requieren esa prueba en la fase 3.
   - Préstamo corriente: $150 M nominales.
