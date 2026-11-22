# Encargo: notas-anexos-2026

**Estado:** TERMINADO
**Rama:** obrera/notas-anexos (sesión reutilizada)
**Creado por la reina:** 2026-11-22 · Fase 3 de `PLAN_ANDINA.md`

## Objetivo
Redactar las notas y anexos de los EECC al 31/12/2026, comparativos con 2025, en moneda de dic-2026.

## Contexto
- Insumos: `entregables/ajuste/`, `insumos/andina/` (auxiliares 2025 y 2026) e `insumos/indices/README.md`.
- Normativa vigente FACPCE y art. 64 de la Ley 19.550. Verificá la vigencia y citala.

## Datos de la fase 2 (integrados y verificados por la reina)
- Ajuste: `entregables/ajuste/` (27 controles OK; verificado por la reina con `verificar_ajuste.py`).
- Resultado ajustado 2026: **$496.161.444** (moneda de dic-26). Comparativo 2025 en moneda de dic-26: **$611.508.366**.
- RECPAM 2026: −$174.121.078. Posición monetaria neta activa.
- PN ajustado: al 31/12/2025 (moneda de dic-26) $7.880.881.500; al 31/12/2026 $8.026.382.100.
- Índices de nov y dic-26 **estimados**: dejarlo dicho en la nota de unidad de medida.
- **Duda** reserva legal (art. 70 LGS): faltante de $731,0 M a dic-26 si la base es capital + ajuste. Exponer como Duda, sin contabilizar asignaciones.
- Pendiente: prueba de valor recuperable / VNR. Si no hay datos, Dato faltante.

## Entregables
- [x] `entregables/notas-anexos-2026/notas_andina_2026.md`. Debe incluir:
  - Bases de preparación y unidad de medida: índice usado, coeficientes e inflación anual.
  - Criterios de medición.
  - Composición de rubros.
  - Bienes de uso y seguros con referencia a los auxiliares.
  - Impuesto a las ganancias.
  - Reserva legal (art. 70 LGS) y restricciones a la distribución.
  - Hechos posteriores.
- [x] `entregables/notas-anexos-2026/anexos_andina_2026.xlsx`, con los anexos:
  - Bienes de uso: origen, altas, bajas, amortizaciones y neto, comparativo.
  - Previsiones.
  - Costo de mercaderías vendidas.
  - Gastos por función (art. 64 LGS).
  - Seguros contratados: póliza, cobertura, bienes y vigencia.

## Criterios de aceptación
- Las notas y anexos concilian con los BSS ajustados y con `eecc-2026`. Agregá una tabla de control.
- Valores en moneda de dic-2026, comparativo incluido.

---
## Bitácora
2026-11-22 · Hecho: merge de la rama integrada; notas (11 notas + tabla de control) y anexos I–V (Bienes de uso comparativo, Previsiones, CMV, Gastos art. 64 LGS, Seguros) generados desde `entregables/ajuste/` con `generador_notas_anexos.py`. Verificación con `verificar_notas_anexos.py` (Python + LibreOffice): 27 controles del Excel OK; tabla de control de las notas OK (ESP 2026/2025, ER 2026/2025, otros). Totales: activo $10.718.152.850; pasivo $2.691.770.750; PN $8.026.382.100; resultado 2026 $496.161.444; comparativo 2025 $611.508.366; bienes de uso $5.631.383.710; CMV $14.648.408.412; compras históricas $13.115.000.000. · Falta: cruce con `eecc-2026` (lo hace la reina). · Siguiente paso: regenerar con los índices oficiales de nov/dic-2026 cuando se publiquen. · Decisión pendiente (no bloquea): base de la reserva legal (art. 70 LGS); impuesto diferido; compensación de anticipos de IG (unificar con eecc-2026). Dudas/datos faltantes listados al final de las notas.
2026-11-22 · Hecho: por decisión de la reina, compensación de anticipos de IG contra el IG a pagar en las Notas 6, 7.5 y 8 (bruto, compensación y neto); tabla de control con columna de cruce contra `entregables/eecc-2026/eecc_andina_2026.xlsx` (ESP y ER, 2026 y 2025): todo OK. Activo neto $10.368.152.850 (2026) y $10.663.951.279 (2025); pasivo $2.341.770.750 y $2.783.069.779. Verificador ampliado con el cruce contra eecc-2026: OK. · Falta: nada del encargo. · Siguiente paso: regenerar con los índices oficiales de nov/dic-2026. · Decisión pendiente (no bloquea): reserva legal e impuesto diferido.
