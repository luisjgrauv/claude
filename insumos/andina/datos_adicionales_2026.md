# Distribuidora Andina S.A. (FICTICIA) — Datos adicionales del ejercicio 2026

Datos inventados para práctica. No corresponden a ninguna empresa real.

## Generales
- Ejercicio económico: del 01/01/2026 al 31/12/2026 (ejercicio ficticio completo). Saldos iniciales: saldos finales 2025 después del cierre de resultados.
- **Valores históricos nominales, sin reexpresar.** Cada partida no monetaria conserva su fecha de origen en los auxiliares.
- Inflación nominal de referencia usada para fijar importes: ~2 % mensual (solo para verosimilitud; no es un índice).
- Generado con `entregables/historicos/generador_historicos.py`; verificado con `entregables/historicos/verificar_historicos.py` (Python + LibreOffice).

## Archivos
| Archivo | Contenido |
|---|---|
| `bss_2026.xlsx` / `.csv` | BSS al 31/12/2026: saldo inicial, débitos, créditos y saldo final; cuentas de resultado abiertas. Cuenta nueva: 4.3.01 Resultado por venta de bienes de uso. |
| `auxiliares_2026.xlsx` | Bienes de uso (con altas y bajas), Seguros, Mercaderías (FIFO por mes de compra), Patrimonio neto (por año de origen), Resultados mensuales 2026, Conciliación BSS y Notas. |
| `auxiliares_complementarios_2025.xlsx` | Mercaderías, Patrimonio neto y Resultados mensuales 2025, con conciliación contra `bss_2025.xlsx`. |
| `../../entregables/historicos/eecc_historicos_2025_2026.xlsx` | ESP (31/12/2024, 2025 y 2026) y ER (2025 y 2026) históricos. Estado de trabajo. |
| `../../entregables/historicos/control_generador.json` | Cifras de control y asientos resumidos 2026 (insumo para el EFE). |

## Hechos del ejercicio 2026
1. **Asamblea del 28/04/2026:** dividendos en efectivo $300.000.000 sobre el resultado 2025, pagados el 30/04/2026. La reserva legal queda en $4.000.000 (ya en el 20 % nominal del capital). **Duda:** base del art. 70 LGS después del ajuste (fase 2).
2. **Baja por venta del camión C01** (alta 14/05/2018, VO $2.375.000) el 15/06/2026. Precio $28.000.000 + IVA 21 %. Amortización acumulada $1.900.000; VRC $475.000; **resultado por venta $27.525.000**. Año de baja sin amortización.
3. **Alta del camión C11** (pesado 26 t, caja térmica, dominio AJ012WX) el 15/06/2026 por $175.000.000 + IVA 10,5 %. Vida útil 10 años, año de alta completo. La flota sigue en 10 camiones.
4. **Renovación de las 10 pólizas** con la misma vigencia y prima +35 % nominal (total primas $184.680.000), pagadas al inicio de cada vigencia.
5. **Préstamo bancario:** amortización de capital $150.000.000 (37,5 M por trimestre). Saldo al 31/12/2026: $300.000.000. Intereses $105.000.000.
6. **Ventas** $18.200.000.000 (+30 %), margen bruto 30 %. **Gastos operativos** +30 % nominal. IIBB 3 % de las ventas.
7. **Mercaderías:** compras $13.115.000.000; existencia final $1.625.000.000 (nov-2026 $413.120.000 y dic-2026 $1.211.880.000).
8. **Impuesto a las ganancias:** 35 % del resultado contable, sin diferido. Se canceló el impuesto 2025 ($498.806.175) neto de anticipos ($150.000.000). Anticipos 2026: $350.000.000.

Al 31/12/2026 hay: 2 inmuebles (terreno y edificio por separado), 10 camiones (C02–C11), 5 vehículos ejecutivos (E01–E05; E01 totalmente amortizado), 2 camionetas (U01–U02) y 10 pólizas.

## Cifras de control 2026
| Concepto | Importe ($) |
|---|---|
| Ventas | 18.200.000.000 |
| Costo de mercaderías vendidas | 12.740.000.000 |
| Resultado por venta de bienes de uso | 27.525.000 |
| Resultado antes de impuesto | 1.923.345.000 |
| Impuesto a las ganancias | 673.170.750 |
| **Resultado del ejercicio** | **1.250.174.250** |
| Amortizaciones 2026 | 84.150.000 |
| Gasto de seguros 2026 | 171.030.000 |
| Seguros no devengados 31/12/2026 | 52.650.000 |
| Totales débitos = créditos | 97.923.246.925 |
| Saldos deudores = acreedores al 31/12/2026 | 22.948.044.575 |
| RNA inicial 2026 (incluye resultado 2025) | 1.936.107.825 |

## ESP histórico (totales)
| Concepto | 31/12/2024 | 31/12/2025 | 31/12/2026 |
|---|---|---|---|
| Total activo | 2.933.753.500 | 4.230.914.000 | 5.602.052.825 |
| Total pasivo | 1.780.000.000 | 2.270.806.175 | 2.691.770.750 |
| Total patrimonio neto | 1.153.753.500 | 1.960.107.825 | 2.910.282.075 |

## Patrimonio neto: historia (supuestos)
- Capital $20.000.000 integrado el 14/03/2014 (el mes es dato; el día es supuesto).
- Reserva legal: 5 % de los resultados 2014–2021, constituida en las asambleas de abril de 2015 a 2022 (tope $4.000.000 alcanzado en 04/2022).
- Resultados 2014–2023 y dividendos 2019–2024: supuestos verosímiles. Resultado 2024 = $702.703.500, residuo que cuadra los RNA al 31/12/2024 ($1.129.753.500).
- Cada dividendo se imputa al resultado del ejercicio inmediato anterior a la asamblea.

## Supuestos simplificadores relevantes para las fases 2 y 3
- Seguros sin IVA discriminado (como en 2025).
- Impuesto a las ganancias imputado en diciembre en los resultados mensuales.
- Préstamo: porción corriente = $150.000.000 por año (supuesto para el ESP).
- Existencia inicial 2025 abierta en nov-2024 ($160 M) y dic-2024 ($740 M): supuesto.

## Dudas y datos faltantes
- **Duda:** base de la reserva legal (art. 70 LGS) después del ajuste por inflación. La define la fase 2.
- **Dato faltante:** cronograma contractual del préstamo (se asumió amortización trimestral lineal).
