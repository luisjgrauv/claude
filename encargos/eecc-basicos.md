# Encargo: eecc-basicos

**Estado:** DETENIDA
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
- 2026-11-22 · **FRENO recibido de la reina (insumos en revisión: BSS nuevo con más detalle y auxiliares).** · **Hecho:** leídos encargo, CLAUDE.md, BSS y datos adicionales actuales. Verificación preliminar con el BSS viejo (no vigente): resultado $73.125.000 OK; ESP cuadra (Activo 573.750.000 = Pasivo 191.875.000 + PN 381.875.000); EFE indirecto concilia (operativas 65.250.000; inversión −22.000.000; financiación −50.000.000; variación −6.750.000). Norma verificada: RT 54 (NUA, T.O. RT 59) deroga RT 8 y RT 9; presentación según RT 54 (ver Res. JG 660/26, dispensa transitoria de presentación). No se generó ningún Excel ni entregable. · **Falta:** todo el entregable (libro Excel con BSS, ESP, ER, EEPN, EFE, Controles, Supuestos; resumen.md), sobre los insumos nuevos. · **Siguiente paso exacto:** cuando la reina avise, releer `insumos/andina/` nuevos y el encargo (por si cambia el resultado de control), y construir el libro con openpyxl + fórmulas vinculadas a la hoja BSS; recalcular con LibreOffice (disponible en /usr/bin/soffice) y verificar controles en cero. Dudas a trasladar: RECPAM de efectivo no informado por separado en EFE (dato faltante); intereses pagados clasificados en operativas. · **Decisión pendiente:** ninguna; espero insumos nuevos.
