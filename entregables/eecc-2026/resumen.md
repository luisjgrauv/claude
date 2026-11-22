# EECC 2026 — Distribuidora Andina S.A. (FICTICIA)

**Conclusión:** los estados básicos al 31/12/2026 (comparativos con 2025, en moneda de dic-2026) cierran. El resultado da **$496.161.444**, el PN **$8.026.382.100** y el activo **$10.368.152.850**. Los 19 controles del libro dan cero tras recalcular con LibreOffice. El verificador independiente en Python también da OK. Todas las cifras coinciden con la fase 2.

## 1. Estado de situación patrimonial (en $)
| Rubro | 31/12/2026 | 31/12/2025 |
|---|---:|---:|
| Caja y bancos | 549.713.825 | 546.691.576 |
| Inversiones (plazo fijo) | 270.000.000 | 288.699.344 |
| Créditos por ventas | 2.204.500.000 | 2.227.577.893 |
| Otros créditos (sin anticipos de IG) | 79.617.953 | 82.399.053 |
| Bienes de cambio | 1.632.937.362 | 1.652.008.929 |
| **Activo corriente** | **4.736.769.139** | **4.797.376.794** |
| Bienes de uso | 5.631.383.710 | 5.866.574.485 |
| **Total activo** | **10.368.152.850** | **10.663.951.279** |
| Deudas comerciales | 1.495.000.000 | 1.509.110.207 |
| Remuneraciones y cargas sociales | 169.000.000 | 170.595.067 |
| Cargas fiscales (IG neto de anticipos) | 377.770.750 | 512.843.120 |
| Préstamos corrientes | 150.000.000 | 196.840.462 |
| **Pasivo corriente** | **2.191.770.750** | **2.389.388.855** |
| Préstamos no corrientes | 150.000.000 | 393.680.924 |
| **Total pasivo** | **2.341.770.750** | **2.783.069.779** |
| **Patrimonio neto** | **8.026.382.100** | **7.880.881.500** |

El ESP al 31/12/2025 difiere del comparativo de la fase 2 en una sola reclasificación de $196.840.462: los anticipos de IG se compensan con el pasivo. El encargo pide presentar el impuesto neto de anticipos. El PN no cambia.

## 2. Estado de resultados (en $)
| Concepto | 2026 | 2025 |
|---|---:|---:|
| Ventas netas | 20.298.437.162 | 20.749.258.452 |
| Costo de mercaderías vendidas | (14.648.408.412) | (14.955.044.768) |
| **Resultado bruto** | **5.650.028.749** | **5.794.213.684** |
| Gastos de comercialización | (3.173.091.074) | (3.228.372.658) |
| Gastos de administración | (1.081.445.216) | (1.119.537.023) |
| Resultado por venta de bienes de uso | (13.750.256) | – |
| Intereses ganados | 80.312.136 | 81.523.008 |
| Intereses del préstamo | (118.601.067) | (179.756.770) |
| RECPAM | (174.121.078) | (81.993.622) |
| **Resultado antes de IG** | **1.169.332.194** | **1.266.076.619** |
| Impuesto a las ganancias | (673.170.750) | (654.568.252) |
| **Resultado del ejercicio** | **496.161.444** | **611.508.366** |

## 3. EEPN y EFE
- **EEPN:** dividendos reexpresados por $350.660.845 en 2026 y $185.640.480 en 2025. Capital $20.000.000 nominal. Ajuste de capital $4.735.227.863. Reserva legal $220.012.066 (sin cambios). RNA al 31/12/2026: $3.051.142.171.
- **EFE (método indirecto; efectivo = caja + bancos + plazo fijo):**

| Causa | 2026 | 2025 |
|---|---:|---:|
| Operativas | 885.396.271 | 1.149.816.797 |
| Inversión | (165.092.321) | (298.223.804) |
| Financiación | (515.041.527) | (403.797.030) |
| RECPAM generado por el efectivo (estimado) | (220.939.518) | (192.433.303) |
| **Variación del efectivo** | **(15.677.095)** | **255.362.659** |

## 4. Controles (todos en cero, tolerancia $0,50)
- A = P + PN en 2026 y en 2025.
- Resultado del ER = variación de RNA del EEPN (+ dividendos) en 2026 y en 2025.
- Variación de efectivo del EFE = variación del ESP en 2026 y en 2025. El efectivo al cierre de 2025 es igual al de inicio de 2026.
- Resultado, PN y ESP comparativo = fase 2 (rubro por rubro, neto de la reclasificación de anticipos).
- EEPN por componente = BSS. Los dos BSS cuadran. El préstamo corriente es de $150 M.
- `verificar_eecc.py` recalcula con LibreOffice y replica en Python el ESP, el resultado y el EFE. Resultado: OK.

## 5. Norma
- **RT 54 FACPCE (NUA), t.o. RT 59.** Obligatoria para ejercicios iniciados desde el 01/01/2025.
- El texto ordenado (copia del CPCECABA) deja sin efecto la RT 8 y la RT 6. Fuentes de CPCE Salta indican que también deroga totalmente la RT 9.
- **Duda:** no pude leer el PDF oficial porque la red de la sesión lo bloqueó. Faltan verificar la derogación de la RT 9 y los números de párrafo.
- La Res. JG 660/26 es una dispensa transitoria de presentación, solo para el primer ejercicio de aplicación de la RT 54.
- **RECPAM:** en el ER, renglón propio dentro de resultados financieros y por tenencia. En el EFE, el del efectivo va como causa separada (Interpretación 2).

## 6. Dudas y datos faltantes
1. **Duda — Reserva legal (art. 70 LGS):** el faltante es de $731,0 M al 31/12/2026, sin contabilizar. Debe validarlo el asesor societario con el criterio de la IGJ.
2. **Estimación — RECPAM del efectivo:** no hay saldos mensuales de efectivo; se supuso una evolución lineal. Solo cambia la distribución entre operativas y RECPAM, no la variación total.
3. **Estimación — Índices de nov y dic-2026:** no son oficiales. Hay que regenerar los libros de la fase 2 y este (`generador_eecc.py`).
4. **Dato faltante — Valor recuperable y VNR** (bienes de uso $5.631,4 M; bienes de cambio $1.632,9 M).
5. **Dato faltante — Cronograma del préstamo:** se supusieron cuotas trimestrales de $37,5 M en mar, jun, sep y dic.
6. **Duda — Res. JG 660/26:** aplica solo si 2026 fuera el primer ejercicio de aplicación de la RT 54.

## Archivos
- `eecc_andina_2026.xlsx`: hojas Supuestos, Controles, ESP, ER, EEPN, EFE, BSS 2026, BSS 2025, Datos EFE y Fase 2.
- `generador_eecc.py` y `verificar_eecc.py`: generación y verificación reproducibles.
