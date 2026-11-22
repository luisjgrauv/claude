# Encargo: notas-anexos

**Estado:** DETENIDA — obsoleto: el usuario redefinió el alcance (ejercicio 2025, valores históricos). Pendiente de nuevo encargo.
**Rama:** obrera/notas-anexos
**Creado por la reina:** 2026-11-22

## Objetivo
Redactar las notas y los anexos de los estados contables de Distribuidora Andina S.A. (ficticia) al 31/12/2026.

## Contexto
- Insumos: `insumos/andina/bss_2026.xlsx` (también en .csv) y `insumos/andina/datos_adicionales.md`.
- Los importes ya están en moneda homogénea. **No reexpreses.**
- Resultado del ejercicio: **$73.125.000** (ganancia).
- Normativa: normas contables profesionales argentinas vigentes (FACPCE) y art. 64 de la Ley 19.550 (anexo de gastos). Verificá la vigencia de cada norma y citala.

## Entregables
- [ ] `entregables/notas-anexos/notas_andina_2026.md` con:
  - Nota 1: bases de preparación (norma aplicable, moneda homogénea, comparativos y su limitación).
  - Nota 2: principales criterios de medición por rubro.
  - Nota 3 en adelante: composición de rubros (caja y bancos, inversiones, créditos por ventas, otros créditos, deudas comerciales, remuneraciones, cargas fiscales, préstamos con su apertura corriente / no corriente).
  - Nota de impuesto a las ganancias (conciliación tasa × resultado vs. cargo).
  - Nota de restricción por reserva legal y nota de hechos posteriores.
- [ ] `entregables/notas-anexos/anexos_andina_2026.xlsx` con:
  - Anexo Bienes de uso: valores de origen, altas, bajas, amortizaciones y neto.
  - Anexo Previsiones: saldo inicial, aumentos, aplicaciones y saldo final.
  - Anexo Costo de mercaderías vendidas: EI + compras − EF.
  - Anexo Gastos (art. 64 LGS): apertura por función (comercialización / administración).

## Criterios de aceptación
- Los totales de las notas cuadran con el BSS. Agregá una tabla de control al final de las notas.
- Compras del ejercicio = $615.000.000 (control del anexo de CMV).
- Importes en pesos sin decimales y con separador de miles.

## Fuera de alcance
- No armes los estados básicos (los hace la obrera `eecc-basicos`).
- No modifiques `insumos/`.

## Dependencias
- La reina cruza tus totales con los de `eecc-basicos`.

---
## Bitácora
