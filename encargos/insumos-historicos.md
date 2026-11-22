# Encargo: insumos-historicos

**Estado:** TERMINADO
**Rama:** obrera/insumos-historicos
**Creado por la reina:** 2026-11-22 · Fase 1 de `PLAN_ANDINA.md`

## Objetivo
Completar la base **histórica nominal** de Distribuidora Andina S.A. (ficticia, distribuidora de insumos) para 2025 y 2026.

## Contexto
- Ya existen y **no se modifican**: `insumos/andina/bss_2025.xlsx`, `auxiliares_2025.xlsx` y `datos_adicionales.md` (leelos enteros). El generador original no está en el repo: replicá su lógica a partir de los auxiliares.
- Todo en valores históricos nominales. El ajuste por inflación lo hace otra obrera; vos solo dejás la **fecha de origen** de cada partida no monetaria.
- Convenciones: amortización lineal, año de alta completo, año de baja sin amortización, sin valor residual. Seguros anuales con prima pagada al inicio de la vigencia y devengamiento mensual.
- Inflación nominal de referencia para fijar precios 2026: alrededor del 2 % mensual. Es solo para que los importes sean verosímiles.

## Entregables
1. `insumos/andina/auxiliares_complementarios_2025.xlsx`, consistente con el BSS 2025 sin cambiarlo:
   - **Mercaderías:** compras mensuales 2025 (que sumen $10.150.000.000), existencia inicial ($900M) y final ($1.250M) abiertas por mes de compra (FIFO, rotación de alrededor de 1 mes), y CMV ($9.800M).
   - **Patrimonio neto:** capital $20M (integrado en 03/2014); reserva legal $4M con fechas de constitución; Resultados no asignados al 31/12/2024 ($1.129.753.500) abiertos por año de origen (resultados 2014–2024 menos dividendos, con fechas); dividendos 2025 ($120M, abril).
   - **Resultados mensuales 2025:** ventas, compras y gastos por mes, que sumen los saldos del BSS. Es la base del ajuste de resultados.
2. `insumos/andina/bss_2026.xlsx` (y `.csv`): BSS histórico al 31/12/2026, con saldo inicial = saldo final 2025 después del cierre de resultados.
   - Hechos 2026 a modelar:
     - Asamblea de abril de 2026 que distribuye dividendos sobre el resultado 2025. La reserva legal sigue tal como está (ver Dudas).
     - Baja por venta del camión C01 y alta de un camión nuevo C11, para que la flota siga en 10 camiones.
     - Renovación de las 10 pólizas con la misma vigencia y primas nominales más altas.
     - Amortización del préstamo.
     - Ventas y gastos 2026 nominales, de alrededor de +30 %.
   - Al 31/12/2026 tiene que haber: 2 inmuebles, 10 camiones, 5 vehículos ejecutivos, 2 camionetas y 10 pólizas.
3. `insumos/andina/auxiliares_2026.xlsx`: el mismo formato que `auxiliares_2025.xlsx` (bienes de uso, seguros, conciliación) más las hojas de mercaderías, PN y resultados mensuales 2026.
4. `insumos/andina/datos_adicionales_2026.md`: hechos 2026 y cifras de control (resultado, totales).
5. `entregables/historicos/eecc_historicos_2025_2026.xlsx`: ESP y ER históricos de cada año. Es un estado de trabajo, no una presentación formal.

## Criterios de aceptación
- El BSS 2026 cuadra (débitos = créditos; saldos iniciales y finales en cero).
- Todas las conciliaciones de auxiliares contra el BSS dan diferencia 0, en 2025 y en 2026.
- Toda partida no monetaria tiene su fecha de origen.
- Verificado con Python y con los totales recalculados por LibreOffice (`soffice --headless --convert-to csv`).

## Fuera de alcance
- No reexpreses nada. No modifiques los archivos 2025 existentes.

## Dudas conocidas (no las resuelvas, dejalas anotadas)
- Base de la reserva legal (art. 70 LGS) después del ajuste: la define la fase 2.

---
## Bitácora
2026-11-22 · Hecho: generador `entregables/historicos/generador_historicos.py` (replica la lógica de amortización y seguros de los auxiliares 2025, con asserts contra ellos); `auxiliares_complementarios_2025.xlsx` (mercaderías FIFO por mes de compra, PN por año de origen, resultados mensuales 2025); `bss_2026.xlsx/.csv`; `auxiliares_2026.xlsx`; `datos_adicionales_2026.md`; `entregables/historicos/eecc_historicos_2025_2026.xlsx`; `control_generador.json` (asientos resumidos 2026 para el EFE). · Falta: nada. · Siguiente paso: revisión e integración por la reina. · Decisión pendiente: ninguna.
2026-11-22 · Hecho: verificación `verificar_historicos.py` con Python y LibreOffice (soffice → csv de todas las hojas): 28 controles OK, 0 errores. BSS 2026 cuadra (D = C = 97.923.246.925; saldos iniciales y finales 0); conciliaciones 2025 y 2026 con diferencia 0; fechas de origen en bienes, pólizas, capas de mercaderías y PN; originales 2025 sin cambios. Resultado 2025 926.354.325; resultado 2026 1.250.174.250. · Falta: nada. · Siguiente paso: fase 2 (ajuste-inflacion). · Decisión pendiente: ninguna. Duda abierta (no resuelta, por encargo): base de la reserva legal art. 70 LGS después del ajuste.
