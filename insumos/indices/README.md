# Índices para el ajuste por inflación (RT 6 · Res. JG FACPCE 539/18)

**Conclusión:** la serie dic-2013 → oct-2026 es oficial y está completa (155 meses). Nov-2026 y dic-2026 son **estimados**. La inflación 2025 que surge de la serie (31,55 %) coincide con el dato publicado por el INDEC (31,5 %).

## 1. Fuente y norma
| Concepto | Detalle |
|---|---|
| Norma | Resolución JG FACPCE 539/18. Define el índice de la RT 6 como el empalme del IPC Nacional INDEC (base dic-2016 = 100) con el IPIM (Res. JG 517/16). La página de FACPCE lo rotula "Índice RT 6 – RT 54", así que la serie sigue vigente con la RT 54. |
| Fuente principal | Excel FACPCE "Indice-FACPCE-Res.-JG-539-18-2026-10.xlsx": https://www.facpce.org.ar/wp-content/uploads/2026/11/Indice-FACPCE-Res.-JG-539-18-2026-10.xlsx (página: https://www.facpce.org.ar/indices-facpce/). Último mes publicado: oct-2026. |
| Metodología FACPCE | https://www.facpce.org.ar/wp-content/uploads/2020/05/MetodologiaIndices.pdf. Tramos: ene-1993 a oct-2015, IPIM INDEC empalmado; nov-2015 a dic-2016, IPIM FACPCE (Res. JG 517/16) empalmado; desde ene-2017, IPC Nacional INDEC. |
| Control INDEC (1) | Cuadro INDEC `sh_ipc_01_26.xls`: https://www.indec.gob.ar/ftp/cuadros/economia/sh_ipc_01_26.xls (índices dic-2016 a dic-2025 y variación interanual). |
| Control INDEC (2) | API Series de Tiempo, serie `148.3_INIVELNAL_DICI_M_26` (fuente INDEC): https://apis.datos.gob.ar/series/api/series/?ids=148.3_INIVELNAL_DICI_M_26&format=csv (dic-2016 a oct-2026). |
| Fecha de consulta | 22/11/2026, para todas las fuentes. |
| Próxima publicación | Según la nota del Excel FACPCE, el índice de nov-2026 se publica el 16/12/2026. |

**Verificación:** entre dic-2016 y oct-2026, la serie FACPCE es idéntica a la serie del INDEC de la API (diferencia relativa máxima de 2·10⁻¹⁶). El índice de dic-2025 (10.121,3715) coincide además con el cuadro INDEC `sh_ipc_01_26.xls`.

## 2. Archivo `indices_facpce.csv`
- Separador `;`. Decimal con punto.
- `indice`: base dic-2016 = 100, redondeado a 4 decimales. Antes de 2017, el Excel FACPCE trae más decimales; el redondeo no tiene efecto material.
- `variacion_mensual`: fracción decimal (0,021718 = 2,1718 %). Para dic-2013 se calculó contra nov-2013 del mismo Excel.
- `estado`: `oficial` (dic-2013 a oct-2026) o `estimado` (nov-2026 y dic-2026). No hay meses `faltante`.

## 3. Método de estimación (nov-2026 y dic-2026)
- Tasa mensual = media geométrica de las variaciones de ago, sep y oct-2026: (12.785,9321 / 12.076,3937)^(1/3) − 1 = **1,9213 %**.
- Variaciones de referencia: ago 1,6592 %; sep 1,9336 %; oct 2,1718 %.
- Nov-2026 = 12.785,9321 × 1,019213 = **13.031,5909** (estimado).
- Dic-2026 = 13.031,5909 × 1,019213 = **13.281,9696** (estimado).
- Es una **estimación propia, no un dato oficial**. Se elige porque el encargo la admite y es reproducible.
- **Duda / alternativa:** el REM del BCRA es el dato de mercado habitual, pero `www.bcra.gob.ar` no es accesible desde este entorno. La serie va en aumento (1,66 % → 2,17 %): si nov y dic repitieran el 2,17 % de octubre, el índice de dic-2026 sería ≈ 13.347 (+0,5 % sobre lo estimado).
- Cuando salga el dato oficial de nov-2026 (16/12/2026), hay que reemplazar el estimado.

## 4. Coeficientes de control
| Tramo | Índice inicial | Índice final | Coeficiente | Variación | Estado |
|---|---|---|---|---|---|
| dic-2024 → dic-2025 | 7.694,0075 | 10.121,3715 | **1,315488** | **31,55 %** | Oficial |
| dic-2025 → dic-2026 | 10.121,3715 | 13.281,9696 | **1,312270** | **31,23 %** | Incluye 2 meses estimados |
| dic-2024 → dic-2026 | 7.694,0075 | 13.281,9696 | **1,726275** | **72,63 %** | Incluye 2 meses estimados |
| dic-2025 → oct-2026 (acumulado) | 10.121,3715 | 12.785,9321 | 1,263261 | 26,33 % | Oficial |
| oct-2025 → oct-2026 (interanual) | — | 12.785,9321 | 1,331332 | 33,13 % | Oficial |

- **Inflación anual 2025:** 31,55 % según la serie. El INDEC publicó 31,5 % (variación interanual de dic-2025, `sh_ipc_01_26.xls`). ✔ Coincide con tolerancia de redondeo.
- **Inflación anual 2026:** 31,23 % (**estimada**). 26,33 % oficial a octubre, más los dos meses estimados.

## 5. Limitaciones
- Los coeficientes que llegan a dic-2026 dependen de la estimación. Hay que recalcularlos con los índices oficiales de nov y dic-2026 (publicación de dic-2026 y ene-2027).
- Las cifras oficiales del INDEC para 2026 se verificaron a través de la API datos.gob.ar (fuente INDEC) y del Excel FACPCE. No se descargó el informe técnico del INDEC de oct-2026.
