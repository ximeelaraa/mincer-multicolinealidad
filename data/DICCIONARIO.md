# Datos documentados

**Equipo 4:** Ximena Lara, Camila Ricardi y Sara Enriquez · Fundamentos de Econometría, Universidad Panamericana

## 1. Base principal: `wage2`

| Campo | Detalle |
|---|---|
| Archivo utilizado | [`wage2_muestra.csv`](wage2_muestra.csv): la muestra limpia (722 filas) con las variables transformadas |
| Fuente original | National Longitudinal Survey of Young Men (NLS), EE. UU., 1980. La base la compilaron Blackburn y Neumark (1992, *Quarterly Journal of Economics* 107(4): 1421–1436) |
| Distribución | Conjunto de datos `WAGE2` de Wooldridge, *Introductory Econometrics: A Modern Approach*. Se carga con el paquete de Python `wooldridge` (v0.5.0, PyPI): `wooldridge.data("wage2")` |
| Fecha de consulta | 18 de noviembre de 2026 |
| Licencia o uso | Datos públicos con fines académicos que acompañan al libro de texto. No contienen información que identifique a las personas |
| Unidad de análisis | Hombre trabajador (una fila por persona) |
| Población y periodo | Hombres de 28 a 38 años de la cohorte NLS, entrevistados en 1980 |
| Observaciones | 935 originales y 722 en la muestra de análisis |

### Diccionario de variables

| Variable | Definición | Unidad | Signo esperado sobre ln(salario) |
|---|---|---|---|
| `wage` | Ingreso mensual | USD de 1980 | (variable dependiente en niveles) |
| `lwage` | ln(`wage`) | log | **Variable dependiente (Y)** |
| `educ` | Escolaridad | años | + (variable de interés) |
| `exper` | Experiencia laboral | años | + |
| `tenure` | Antigüedad con el empleador actual | años | + |
| `age` | Edad | años | ? (casi colineal con educ + exper) |
| `IQ` | Coeficiente intelectual | puntos (media aprox. 100) | + (proxy de habilidad) |
| `KWW` | Puntaje de la prueba *Knowledge of the World of Work* | puntos | + (proxy de habilidad) |
| `meduc` | Escolaridad de la madre | años | + (entorno familiar) |
| `feduc` | Escolaridad del padre | años | + (entorno familiar) |
| `married` | 1 si está casado, 0 si no | dummy | + |
| `black` | 1 si es afroamericano, 0 si no | dummy | − |
| `south` | 1 si vive en el sur de EE. UU., 0 si no | dummy | − |
| `urban` | 1 si vive en un área metropolitana (SMSA), 0 si no | dummy | + |
| `sibs`, `brthord`, `hours` | Número de hermanos, orden de nacimiento, horas semanales | conteo, horas | Solo se usan en la sección de R² ajustada |

**Grupo de referencia de las dummies:** no casado, no afroamericano, fuera del sur y fuera de un área metropolitana.

### Limpieza y transformaciones (en `scripts/analisis.py`)
1. Se eliminan 213 observaciones sin escolaridad de la madre o del padre (`meduc` o `feduc` faltante): de 935 quedan **722**. Esto puede introducir selección: las personas sin ese dato podrían tener otro perfil (ver limitaciones en el informe).
2. `exper_c` = `exper` − media(`exper`): experiencia centrada, para reducir la colinealidad entre `exper` y `exper²`.
3. `I(exper_c**2)`: cuadrado de la experiencia centrada (perfil de ciclo de vida).
4. `peduc` = (`meduc` + `feduc`)/2: escolaridad promedio de los padres. Solo se usa en el análisis de sensibilidad.
5. `pexper` = `age` − `educ` − 6: experiencia potencial. Solo se construye para ilustrar la multicolinealidad **perfecta**; no entra en ningún modelo estimado.

No se recopilaron datos propios ni datos personales identificables.

## 2. Bases complementarias (solo en el anexo de temas del curso)

| Base | Fuente y carga | Fecha de consulta | Uso |
|---|---|---|---|
| `hprice1` | 88 casas del área de Boston, Wooldridge. Se carga con `wooldridge.data("hprice1")` | 18/11/2026 | Precios hedónicos (4.6) |
| `stocks` | Precios semanales normalizados de GOOG, AAPL, AMZN, FB, NFLX y MSFT, 2018–2019. Se carga con `plotly.express.data.stocks()` (plotly 7.1.0) | 18/11/2026 | Beta de mercado (4.9) |
| `macrodata` | Datos trimestrales de EE. UU., 1959Q1–2009Q3 (BEA y Reserva Federal). Se carga con `statsmodels.datasets.macrodata` (statsmodels 0.15.0) | 18/11/2026 | Autocorrelación, ARDL y rezagos (5.7–5.12) |
