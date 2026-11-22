# Encargo: eecc-basicos

**Estado:** DETENIDA — obsoleto: el usuario redefinió el alcance (ejercicio 2025, valores históricos). Pendiente de nuevo encargo.
**Rama:** obrera/eecc-basicos
**Creado por la reina:** 2026-11-22

## Objetivo
Elaborar los estados contables básicos de Distribuidora Andina S.A. (ficticia) al 31/12/2026 en un libro Excel.

## Contexto
- Insumos: `insumos/andina/bss_2026.xlsx` (también en .csv) y `insumos/andina/datos_adicionales.md`.
- Los importes ya están en moneda homogénea. **No reexpreses.**
- Resultado del ejercicio esperado: **$73.125.000** (ganancia). Si tu cálculo da otra cifra, detente y anótalo en la bitácora.
- Normativa: normas contables profesionales argentinas vigentes (FACPCE) para ejercicios iniciados el 01/01/2026. Verificá cuál es la RT aplicable a presentación (RT 8/9 o la que las haya reemplazado) y cítala en la hoja "Supuestos".

## Entregables
- [ ] `entregables/eecc-basicos/eecc_andina_2026.xlsx` con estas hojas:
  - **ESP**: estado de situación patrimonial, corriente / no corriente, comparativo con 31/12/2025.
  - **ER**: estado de resultados 2026. Columna 2025 con la leyenda "Dato faltante".
  - **EEPN**: estado de evolución del patrimonio neto, con saldos iniciales, asamblea y resultado.
  - **EFE**: estado de flujo de efectivo, método indirecto. Efectivo = caja + bancos + plazo fijo.
  - **Controles**: cuadres (Activo = Pasivo + PN; resultado del ER = variación de RNA del EEPN; variación de efectivo del EFE = variación del ESP).
  - **Supuestos**: norma citada, supuestos y Dudas.
- [ ] `entregables/eecc-basicos/resumen.md`: totales por rubro, resultado y controles, en texto.

## Criterios de aceptación
- Todos los controles en cero.
- Fórmulas vinculadas a una hoja "BSS" copiada dentro del libro. No escribas importes fijos donde pueda ir una fórmula.
- Préstamo dividido en corriente $20M y no corriente $20M.
- Importes en pesos sin decimales y con separador de miles.

## Fuera de alcance
- No hagas notas ni anexos (los hace la obrera `notas-anexos`).
- No modifiques `insumos/`.

## Dependencias
- `notas-anexos` usa los mismos insumos. La reina cruza los totales de ambas al final: los rubros del ESP y del ER tienen que coincidir con las notas.

---
## Bitácora
