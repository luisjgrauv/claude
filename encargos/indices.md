# Encargo: indices

**Estado:** TERMINADO
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
- [x] `insumos/indices/indices_facpce.csv` con las columnas `mes (AAAA-MM); indice; variacion_mensual; fuente; estado (oficial|estimado|faltante)`.
- [x] `insumos/indices/README.md`: fuente, resolución citada, fecha de consulta, método de estimación y cuadro de coeficientes de control:
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

- 2026-11-22 · **Hecho:** leí encargo, CLAUDE.md y PLAN_ANDINA.md. Intenté obtener la serie oficial. La política de red del entorno bloquea todos los hosts necesarios: `www.facpce.org.ar` (DNS/403 proxy), `www.indec.gob.ar` (403 proxy), `apis.datos.gob.ar` (403 proxy) y también sitios secundarios (`contadoresenred.com`, DNS). Solo funciona el buscador web, que devuelve fragmentos de terceros, no la serie. · **Falta:** todo el entregable (CSV dic-2013 a dic-2026 y README). No generé el CSV para no cargar valores sin fuente verificable. · **Siguiente paso:** con acceso a red, descargar el Excel de índices RT 6 de FACPCE (Res. JG 539/18) y cotejar con la serie IPC nacional base dic-2016 del INDEC; luego estimar nov y dic-2026. · **Decisión pendiente:** ver abajo.

### ESPERA_DECISION — pregunta para la reina / usuario
No puedo acceder a ninguna fuente oficial desde este entorno. Opciones:
1. **(Recomendada)** Habilitar en el entorno (Network access → Allowed domains) los hosts `www.facpce.org.ar`, `www.indec.gob.ar` y `apis.datos.gob.ar`, y reactivar esta obrera.
2. Que el usuario suba a `insumos/indices/` el Excel oficial de FACPCE (índices RT 6) descargado por él; yo lo proceso, estimo nov/dic-2026 y armo el README.
3. Aceptar fuentes secundarias (blogs profesionales) marcadas como "no verificadas" — no recomendado: incumple el criterio "fuente trazable oficial".

Datos obtenidos solo como fragmentos de búsqueda (terceros, **no verificados**, no usados en entregables; consulta 2026-11-22):
- Índice FACPCE/IPC nacional oct-2026: 12.785,93 (var. mensual 2,17 %; interanual 33,13 %; acumulado 2026: 26,3 %). Fuentes: criteria.com.ar, ajusteya.com.ar, siap.blogdelcontador.com.ar.
- sep-2026: 12.514,15; ene-2026: 10.413,03 (contadoresenred.com, vía buscador). Control: 12.785,93 / 12.514,15 = 1,0217 → coherente con 2,17 %.
- **Dato faltante:** dic-2013 a dic-2025 completos, feb-ago 2026 y la inflación anual 2025 oficial.

- 2026-11-22 16:20 · **Hecho:** la reina informó la decisión del usuario (opción 1: hosts habilitados). Descargué el Excel oficial FACPCE Res. JG 539/18 (hasta oct-2026) y lo crucé con INDEC (`sh_ipc_01_26.xls`) y con la API datos.gob.ar (serie 148.3_INIVELNAL_DICI_M_26): coinciden exactamente. Generé `insumos/indices/indices_facpce.csv` (157 meses: 155 oficiales y 2 estimados, ninguno faltante) y el README con fuentes, método y coeficientes. El valor de oct-2026 (12.785,9321) quedó confirmado contra la fuente oficial. Inflación 2025: 31,55 % según la serie, contra 31,5 % del INDEC ✔. 2026: 31,23 % (estimada). · **Falta:** nada del alcance. · **Siguiente paso:** reemplazar los estimados de nov y dic-2026 cuando se publiquen (16/12/2026 y ene-2027). · **Decisión pendiente:** ninguna. Duda menor: la estimación usa la media geométrica de ago-oct (1,9213 % mensual); el REM del BCRA no fue accesible.
