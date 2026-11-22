# Plan: Distribuidora Andina S.A. — del histórico a los EECC 2026 en moneda homogénea

## Objetivos del usuario
1. Estados financieros **históricos** (nominales) para 2025 y 2026.
2. Ajuste por inflación de **2025** (moneda de dic-2025) y de **2026** (moneda de dic-2026).
3. Reexpresión de 2025 a moneda de dic-2026 como comparativo.
4. EECC 2026 comparativos, con notas y anexos, según la normativa argentina vigente.

## Fases y obreras
| Fase | Obrera | Entrega | Depende de |
|---|---|---|---|
| 1 | `insumos-historicos` | Auxiliares de mercaderías y PN 2025; BSS 2026 histórico y auxiliares 2026; EECC históricos 2025 y 2026 (ESP y ER) | BSS 2025 (listo) |
| 1 | `indices` | Serie de índices FACPCE dic-2013 a dic-2026, con fuente y estado de cada valor (oficial o estimado) | — |
| 2 | `ajuste-inflacion` | Ajuste 2025, ajuste 2026 y reexpresión de 2025 a moneda de dic-2026 | Fase 1 completa |
| 3 | `eecc-2026` (reutiliza la sesión eecc-basicos) | ESP, ER, EEPN y EFE 2026 comparativos | Fase 2 |
| 3 | `notas-anexos-2026` (reutiliza la sesión notas-anexos) | Notas y anexos 2026 comparativos | Fase 2 |

## Integración (la hace la reina)
- Cada obrera trabaja en su rama `obrera/<nombre>`.
- Al terminar, la reina revisa los controles y fusiona esa rama en `claude/armado-de-reina-cqmydu`.
- Las obreras de la fase siguiente arrancan desde esa rama integrada.
- La reina hace el control cruzado final: ESP y ER contra notas y anexos; el resultado del EEPN contra el ER; el EFE contra el ESP.

## Decisiones tomadas por la reina (revisables por el usuario)
- Índices: serie oficial FACPCE. Los meses no publicados se **estiman** y se marcan como tales; la estimación nunca se presenta como dato oficial.
- Ejercicio 2026: ficticio y completo al 31/12/2026, aunque hoy sea 22/11/2026.
