# Encargo: indices

**Estado:** PENDIENTE
**Rama:** obrera/indices
**Creado por la reina:** 2026-11-22 · Fase 1 de `PLAN_ANDINA.md`

## Objetivo
Compilar la serie mensual de índices que exige la normativa contable argentina para el ajuste por inflación, desde **dic-2013** hasta **dic-2026**.

## Contexto
- La serie a usar es la publicada por la FACPCE para la reexpresión en moneda homogénea: el empalme IPIM / IPC nacional del INDEC según la Resolución JG FACPCE 539/18 y sus actualizaciones. Verificá la resolución vigente y citala.
- Hoy es 22/11/2026. Los meses todavía no publicados (como mínimo, nov-2026 y dic-2026) se **estiman**.
- El método de estimación tiene que estar documentado (por ejemplo, el promedio de la variación de los últimos 3 meses o el REM del BCRA) y cada valor estimado se marca como tal.
- Si no podés obtener valores oficiales de una fuente verificable, **no inventes**. Marcá esos meses como "Dato faltante" y avisá a la reina con ESPERA_DECISION.

## Entregables
- [ ] `insumos/indices/indices_facpce.csv` con las columnas `mes (AAAA-MM); indice; variacion_mensual; fuente; estado (oficial|estimado|faltante)`.
- [ ] `insumos/indices/README.md`: fuente, resolución citada, fecha de consulta, método de estimación y cuadro de coeficientes de control:
  - dic-2024 → dic-2025
  - dic-2025 → dic-2026
  - dic-2024 → dic-2026
  - Inflación anual 2025 y 2026.

## Criterios de aceptación
- La serie mensual está completa, sin huecos, salvo los meses marcados como faltantes.
- La variación anual 2025 calculada con la serie coincide con el IPC anual publicado por el INDEC (con tolerancia de redondeo).
- Cada valor tiene su fuente trazable.

## Fuera de alcance
- No hagas el ajuste. No toques `insumos/andina/`.

---
## Bitácora
