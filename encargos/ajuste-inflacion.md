# Encargo: ajuste-inflacion

**Estado:** PENDIENTE (arranca cuando termine la fase 1)
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

## Entregables
- [ ] `entregables/ajuste/ajuste_2025.xlsx`: hojas de coeficientes, bienes de uso, seguros, mercaderías, PN, resultados, RECPAM y BSS ajustado al 31/12/2025.
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
