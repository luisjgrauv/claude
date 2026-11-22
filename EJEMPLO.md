# Ejemplo guiado: cierre de ejercicio con la colmena

Cliente ficticio: **Distribuidora Andina SA** · Cierre: **31/12/2026**.
Los mensajes en `>` son los que escribís vos. El resto es lo que hace la reina.

---

## Paso 0 · Piloto (5 minutos, costo mínimo)
> Reina, abrí una obrera de prueba llamada "piloto". Encargo: crear `entregables/piloto/hola.md` con la fecha de hoy y un resumen de 3 líneas de CLAUDE.md. Avisame cuando termine.

Qué validás:
- que la reina escriba `encargos/piloto.md` y abra la sesión `Obrera · piloto`;
- que la obrera lea el encargo, deje su bitácora y haga push;
- que la reina te dé el parte y actualice `ESTADO.md`.

Si algo falla, se corrige acá y no con un trabajo real.

---

## Paso 1 · Preparar los insumos (antes de arrancar)
- Subí los insumos a `insumos/andina/`: balance de sumas y saldos 2026 (.xlsx o .csv) y EECC 2025 aprobados.
- **Anonimizá antes de subir**: sin CUIT, nombres de personas ni datos bancarios. Usá códigos (CLIENTE-01, PROV-03).
- Con el repo **público**, no subas datos reales de clientes. Pasalo a privado antes (GitHub → Settings → Change visibility).

---

## Paso 2 · Un solo mensaje con todo el objetivo
> Reina: cierre 31/12/2026 de Distribuidora Andina SA. Insumos en `insumos/andina/`. Necesito tres cosas:
> 1. Planilla de ajuste por inflación del ejercicio, con índices FACPCE.
> 2. Borrador de notas a los EECC y políticas contables según la norma vigente.
> 3. Lista de documentación pendiente y correo al cliente pidiéndola.
> Las notas usan las cifras ajustadas de la 1. La lista 3 junta todos los faltantes que detecten la 1 y la 2. Repartí y avisame.

Respuesta esperada de la reina:
```
Reparto (3 obreras):
- ajuste-inflacion → planilla reexpresión + resumen de RECPAM. Dependencia: alimenta a notas-eecc.
- notas-eecc       → notas y políticas. Espera las cifras de ajuste-inflacion; mientras, redacta la parte cualitativa.
- pedido-cliente   → recoge los faltantes de las otras dos y redacta el correo.
Datos compartidos: fecha de cierre 31/12/2026; índice FACPCE dic-2026 = Dato faltante (pendiente de publicación).
Encargos escritos en encargos/. Abro las 3 sesiones.
```

---

## Paso 3 · Así queda un encargo (`encargos/ajuste-inflacion.md`)
```
# Encargo: ajuste-inflacion
Estado: TRABAJANDO · Rama: obrera/ajuste-inflacion

## Objetivo
Reexpresar los EECC al 31/12/2026 en moneda de cierre.
## Contexto
Insumos en insumos/andina/. Índices: serie FACPCE (verificar última publicación).
Si falta el índice de diciembre, usar el último disponible y marcarlo como Dato faltante.
## Entregables
- [ ] entregables/ajuste-inflacion/reexpresion.xlsx (una hoja por rubro, coeficientes trazables)
- [ ] entregables/ajuste-inflacion/resumen.md (RECPAM, supuestos, partidas sin fecha de origen)
## Criterios de aceptación
- El activo y el patrimonio neto ajustados cuadran. El RECPAM se concilia por diferencia.
- Toda partida sin fecha de origen queda listada como Duda.
## Fuera de alcance
- No tocar los insumos originales.
## Dependencias
- notas-eecc usa el resumen.md. pedido-cliente usa la lista de Dudas.
---
## Bitácora
2026-11-23 10:40 · Hecho: carga de BSS y anticuación de Bienes de uso · Falta: PN y resultados · Siguiente: reexpresar el PN · Decisión pendiente: —
```

---

## Paso 4 · Pedir el parte
> ¿Cómo va la colmena?

```
TERMINADAS:       —
TRABAJANDO:       ajuste-inflacion — reexpresando el PN (hito 2 de 3)
                  notas-eecc — políticas contables listas, faltan las cifras
ESPERAN DECISIÓN: pedido-cliente — ¿el correo lo firmás vos o sale a nombre del estudio?
DETENIDAS:        —
```
> A nombre del estudio.

La reina se lo pasa solo a `pedido-cliente`.

---

## Paso 5 · Cambio que afecta a varias obreras
> Se publicó el índice de diciembre: <valor>. Actualizá a quien lo use.

La reina actualiza el contexto de `ajuste-inflacion` y de `notas-eecc`, les envía un aviso y deja constancia en ESTADO.md. `pedido-cliente` no se entera porque no usa ese dato.

## Paso 6 · Ajuste puntual (sin pasar por la reina)
Para cambiar una frase del correo, entrá directo a la sesión `Obrera · pedido-cliente` y decíselo ahí. Cuesta menos tokens.

## Paso 7 · Lección para todas
> Reina, de ahora en más los importes van sin decimales y con separador de miles en punto. Que quede para todas.

La reina lo agrega a "Lecciones aprendidas" en CLAUDE.md.

## Paso 8 · Freno y reanudación
> FRENO

Cada obrera deja su punto de control en la bitácora, hace push y se detiene.

Al día siguiente:
> Retomá lo que dejamos, trabajo a trabajo.

La reina lee ESTADO.md y las bitácoras, reabre lo pendiente, y cada obrera sigue desde su último paso.

## Paso 9 · Cierre
> Mostrame los entregables y lo que quedó como Duda.

La reina revisa cada entregable contra sus criterios de aceptación y te lista las Dudas y los Datos faltantes. La validación profesional final es tuya.
