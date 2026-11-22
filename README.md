# Colmena — abeja reina en Claude Code (nube)

Una sola sesión (la **reina**) recibe tus objetivos, escribe un encargo por trabajo y abre sesiones **obreras** en la nube. Las obreras trabajan en paralelo, anotan su avance en git y la reina te reporta.

| Archivo | Para qué sirve |
|---|---|
| `CLAUDE.md` | Normas comunes. Toda sesión las lee al arrancar. |
| `REINA.md` | Rol, límites y procedimientos de la reina. |
| `FRENOS.md` | Límites de gasto y protocolo FRENO. |
| `ESTADO.md` | Tablero de obreras. |
| `plantillas/ENCARGO.md` | Plantilla del encargo escrito. |
| `encargos/` | Un archivo por obrera: encargo y bitácora. |
| `entregables/` | Resultados. |

## Uso
- Arrancar: *"Eres la reina, lee REINA.md. Objetivo: …"*
- Parte: *"¿Cómo va la colmena?"*
- Cambio transversal: *"El precio del bono pasa a X; avisa a quien lo use."*
- Freno: *"FRENO"* · Reanudar: *"Retoma lo que dejamos, trabajo a trabajo."*

## Tres tipos de ayuda (no confundir)
| Tipo | Qué es | Cuándo usarlo |
|---|---|---|
| **Obrera** | Sesión completa con su propio contexto; puedes entrar a corregirla. | Trabajos largos. |
| **Subagente** | Lo lanza la obrera para algo rápido (buscar, revisar un archivo); devuelve un resumen y desaparece. | Tareas cortas; no hablas con él. |
| **Pregunta lateral** | Consulta a una obrera sin desviarla de su tarea ("¿qué haces?", "¿cuánto te falta?"). | Dudas, sin detener a nadie. |

## Cuándo hablar con la reina y cuándo con la obrera
- **Reina:** arrancar varias cosas, cambios que afectan a varios encargos, decisiones, reanudar tras un corte.
- **Obrera directa:** ajuste puntual de un solo trabajo (cambiar una frase, quitar un rótulo). Cuesta menos tokens y es más rápido.
- La reina revisa la miel; **la valida el usuario**.
