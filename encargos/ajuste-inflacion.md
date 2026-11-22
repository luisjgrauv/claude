# Encargo: ajuste-inflacion

**Estado:** EN_CURSO (hito 1/3 cerrado)
**Rama:** obrera/ajuste-inflacion
**Creado por la reina:** 2026-11-22 · Fase 2 de `PLAN_ANDINA.md`

## Objetivo
1. Ajustar por inflación el ejercicio 2025 a moneda de dic-2025.
2. Ajustar por inflación el ejercicio 2026 a moneda de dic-2026.
3. Reexpresar las cifras 2025 ajustadas a moneda de dic-2026, para usarlas como comparativo.

## Contexto
- Insumos: `insumos/andina/` (BSS y auxiliares 2025 y 2026) e `insumos/indices/indices_facpce.csv`.
- Normativa: la vigente de la FACPCE en materia de unidad de medida (RT 6 / RT 54 y resoluciones de la JG). Verificá vigencia y método, y citalo.
- Método esperado: reexpresión de partidas no monetarias por coeficiente desde su fecha de origen (bienes de uso bien por bien, seguros por fecha de pago, mercaderías por mes de compra, PN por componente y fecha); resultados mensuales; RECPAM por diferencia, con una estimación por el método directo como control de razonabilidad.
- Patrimonio neto al inicio: capital y ajuste de capital por separado. Definí y documentá cómo tratás la reserva legal (art. 70 LGS) y si hace falta reconocer una asignación faltante. Esto es **Duda**: dejá tu propuesta.

## Datos de la fase 1 (integrados y verificados por la reina)
- Índices: `insumos/indices/indices_facpce.csv` (155 oficiales; nov-26 y dic-26 **estimados**). Coeficientes de control: dic24→dic25 1,315488; dic25→dic26 1,312270; dic24→dic26 1,726275.
- Resultados históricos: 2025 $926.354.325; 2026 $1.250.174.250. La apertura 2026 coincide con el cierre 2025 (verificado).
- `entregables/historicos/control_generador.json` tiene los asientos 2026 (útil para RECPAM directo y para el EFE).
- PN: las capas de RNA por año de origen están en `auxiliares_complementarios_2025.xlsx`.
- Mercaderías: las capas FIFO por mes de compra están en los auxiliares complementarios 2025 y en los auxiliares 2026.
- Préstamo: cronograma trimestral **supuesto** (Dato faltante).

## Entregables
- [x] `entregables/ajuste/ajuste_2025.xlsx`: hojas de coeficientes, bienes de uso, seguros, mercaderías, PN, resultados, RECPAM y BSS ajustado al 31/12/2025.
- [ ] `entregables/ajuste/ajuste_2026.xlsx`: la misma estructura, con saldos iniciales = cierre 2025 ajustado × coeficiente dic-2025→dic-2026.
- [ ] `entregables/ajuste/comparativo_2025_en_moneda_dic2026.xlsx`: ESP, ER y EEPN 2025 en moneda de dic-2026.
- [ ] `entregables/ajuste/resumen.md`: método, norma citada, coeficientes usados, resultado histórico vs. ajustado por año, RECPAM y Dudas.

## Criterios de aceptación
- Los BSS ajustados cuadran.
- El patrimonio neto ajustado al cierre 2025, multiplicado por el coeficiente dic-25→dic-26, es igual al PN inicial 2026.
- Los auxiliares ajustados concilian con el BSS ajustado.
- La diferencia entre el RECPAM por diferencia y el método directo está explicada.
- Las fórmulas son vinculadas, sin importes fijos donde corresponde una fórmula. Recalculadas con LibreOffice.

## Fuera de alcance
- No hagas la presentación formal de los EECC ni las notas (eso es la fase 3).

---
## Bitácora
- 2026-11-22 · **Hito 1 — Ajuste 2025 (moneda dic-2025): cerrado.** `ajuste_2025.xlsx` generado por `generador_ajuste.py` (fórmulas vinculadas, recalculado con LibreOffice). Resultado histórico $926.354.325 → ajustado **$465.992.887**. RECPAM por diferencia **−$62.482.293** = método directo (dif. < $0,01; identidad algebraica explicada en hoja RECPAM). PN ajustado 31/12/2025 $6.005.534.707. RNA inicial por diferencia (RT 6, primer ejercicio). Controles: todos OK. **Duda (no bloquea):** reserva legal art. 70 LGS. Propuesta: base = capital + ajuste de capital; faltante $557,1 M a dic-2025; no se contabiliza asignación (requiere asamblea). Falta: hitos 2 y 3.
