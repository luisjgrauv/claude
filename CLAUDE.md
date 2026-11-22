# Normas de la colmena

Toda sesión que trabaje en este repositorio (reina u obrera) lee este archivo al arrancar y lo cumple.

## Idioma y estilo
- Español. Estilo directo, ejecutivo y técnico.
- Conclusión primero, explicación después.
- No inventes datos. Lo que falte se marca como **Duda** o **Dato faltante**.
- Distingue hechos verificados, supuestos, estimaciones y recomendaciones.

## Archivos
- No modifiques archivos originales del usuario. Trabaja sobre una copia o un archivo nuevo.
- Los entregables van a `entregables/<nombre-del-encargo>/`.
- No subas secretos, credenciales ni datos personales de terceros.

## Contabilidad y normativa (cuando aplique)
- Argentina: prioriza la normativa vigente (FACPCE/RT, ARCA, IGJ, BCRA, CNV) y verifica que siga vigente.
- Distingue los tratamientos contable, fiscal, societario y financiero.
- Cita la norma aplicable. No presentes una práctica habitual como si fuera obligación legal.

## Excel / Power BI
- Usa la solución moderna primero: matrices dinámicas, LET, LAMBDA, Power Query, DAX.
- No inventes columnas ni tablas. Si asumes una estructura, dilo.

## Reglas para obreras
1. Tu encargo está en `encargos/<tu-nombre>.md`. Léelo entero antes de hacer nada.
2. Si la bitácora del encargo tiene entradas, continúa desde la última. No empieces de cero.
3. Al cerrar cada hito, anota en la bitácora qué hiciste, qué falta y qué decisión necesitas. Después haz commit y push.
4. Si necesitas una decisión del usuario, pon el estado `ESPERA_DECISION`, deja la pregunta en la bitácora y detente.
5. Si recibes la orden **FRENO**, deja un punto de control (qué hiciste, qué falta, siguiente paso exacto), haz push y detente.
6. Usa subagentes solo para tareas cortas: buscar, revisar o resumir.
7. Revisa tu propio trabajo antes de marcarlo `TERMINADO`: lógica, cálculos, fechas, referencias.

## Lecciones aprendidas
<!-- La reina agrega aquí, con fecha, toda corrección que deba valer para todas las obreras futuras. -->
