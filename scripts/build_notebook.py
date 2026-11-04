import nbformat as nbf
import sys

nb = nbf.v4.new_notebook()
C = []
md = lambda s: C.append(nbf.v4.new_markdown_cell(s))
code = lambda s: C.append(nbf.v4.new_code_cell(s))

md("""# Multicolinealidad en la ecuación de salarios de Mincer

**Curso:** Fundamentos de Econometría

**Pregunta:** ¿Cómo afectan la educación, la experiencia y otras características al salario? ¿Qué pasa con nuestras estimaciones cuando los regresores están muy correlacionados?

**Datos:** `wage2` de Wooldridge (*Introductory Econometrics*): 935 hombres trabajadores, *National Longitudinal Survey*, EE. UU., 1980.

**Modelo base (Mincer, 1974):**

$$\\ln(w_i) = \\beta_0 + \\beta_1\\, educ_i + \\beta_2\\, exper_i + \\beta_3\\, exper_i^2 + u_i$$""")

code("""import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import statsmodels.formula.api as smf
import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import RidgeCV
import wooldridge

sns.set_theme(style="whitegrid")
pd.set_option("display.float_format", "{:.4f}".format)""")

md("## 1. Datos y estadística descriptiva")
code("""df = wooldridge.data("wage2")
vars_ = ["wage", "lwage", "educ", "exper", "tenure", "age", "IQ", "KWW", "meduc", "feduc"]
df = df[vars_ + ["married", "black", "south", "urban"]].dropna().reset_index(drop=True)
print(f"Observaciones después de eliminar faltantes: {len(df)}")
df[vars_].describe().T""")

md("""**Variables:**
- `wage`: salario mensual (USD 1980); `lwage`: su logaritmo
- `educ`: años de escolaridad; `exper`: años de experiencia; `tenure`: años con el empleador actual; `age`: edad
- `IQ`: coeficiente intelectual; `KWW`: puntaje en una prueba de conocimiento del mundo laboral
- `meduc`, `feduc`: escolaridad de la madre y del padre""")

md("## 2. Matriz de correlaciones")
code("""X_vars = ["educ", "exper", "tenure", "age", "IQ", "KWW", "meduc", "feduc"]
corr = df[X_vars].corr()
plt.figure(figsize=(8, 6))
sns.heatmap(corr, annot=True, fmt=".2f", cmap="RdBu_r", vmin=-1, vmax=1, square=True)
plt.title("Correlaciones entre regresores")
plt.tight_layout()
plt.show()""")

md("""Las correlaciones más altas son de esperarse: `IQ`–`KWW` (dos medidas de habilidad), `meduc`–`feduc` (emparejamiento de los padres) y `educ`–`IQ`. Ninguna por sí sola es extrema, pero **la multicolinealidad puede existir aunque ninguna correlación bivariada sea alta**. Por eso necesitamos el VIF.""")

md("""## 3. Multicolinealidad **perfecta**: la experiencia potencial

En muchas bases no se observa la experiencia real y se usa la *experiencia potencial*:

$$pexper = age - educ - 6$$

Si incluimos `age`, `educ` y `pexper` en el mismo modelo, una columna es combinación lineal exacta de las otras y la matriz $X'X$ **no es invertible**.""")
code("""df["pexper"] = df["age"] - df["educ"] - 6
X = sm.add_constant(df[["educ", "age", "pexper"]])
print("Rango de X:", np.linalg.matrix_rank(X.values), "de", X.shape[1], "columnas")
print("Número de condición de X'X: {:.2e}".format(np.linalg.cond(X.T @ X)))""")
code("""# statsmodels usa la pseudo-inversa y "estima" algo, pero los coeficientes no están identificados:
m_perf = smf.ols("lwage ~ educ + age + pexper", data=df).fit()
print(m_perf.params)
print("\\nValor propio mínimo de X'X:", np.linalg.eigvalsh(X.T @ X).min())""")
md("""El rango es menor que el número de columnas y el valor propio mínimo es prácticamente cero, así que **no hay solución única**. Solución: quitar una de las tres variables, porque la información de la tercera ya está contenida en las otras dos. Stata o R eliminarían una automáticamente (`omitted because of collinearity` / `NA`).""")

md("""## 4. Multicolinealidad **imperfecta**: el modelo "kitchen sink"

Estimamos un modelo con todos los regresores disponibles:""")
code("""f_full = "lwage ~ educ + exper + I(exper**2) + tenure + age + IQ + KWW + meduc + feduc + married + black + south + urban"
m_full = smf.ols(f_full, data=df).fit()
print(m_full.summary())""")

md("### 4.1 Factor de Inflación de Varianza (VIF)\n\n$$VIF_j = \\frac{1}{1 - R_j^2}$$\n\ndonde $R_j^2$ es el R² de regresar $x_j$ sobre los demás regresores. Regla práctica: $VIF > 10$ indica un problema serio.")
code("""def tabla_vif(model):
    X = pd.DataFrame(model.model.exog, columns=model.model.exog_names)
    out = pd.DataFrame({
        "variable": X.columns,
        "VIF": [variance_inflation_factor(X.values, i) for i in range(X.shape[1])],
    })
    return out[out.variable != "Intercept"].sort_values("VIF", ascending=False).reset_index(drop=True)

vif_full = tabla_vif(m_full)
vif_full""")
code("""plt.figure(figsize=(8, 5))
colors = ["#c0392b" if v > 10 else "#2980b9" for v in vif_full.VIF]
plt.barh(vif_full.variable, vif_full.VIF, color=colors)
plt.axvline(10, ls="--", color="k", label="umbral VIF = 10")
plt.xscale("log"); plt.xlabel("VIF (escala log)"); plt.legend()
plt.gca().invert_yaxis(); plt.title("VIF – modelo completo")
plt.tight_layout(); plt.show()""")

md("### 4.2 Número de condición")
code("""X = m_full.model.exog
Xs = X / np.linalg.norm(X, axis=0)          # columnas normalizadas (Belsley, Kuh & Welsch)
cond = np.linalg.cond(Xs)
print(f"Número de condición (X escalada): {cond:.1f}")
print("Regla: < 10 bien | 10–30 moderada | > 30 severa")""")

md("""### 4.3 Prueba de Farrar-Glauber (multicolinealidad global)

$H_0$: los regresores son ortogonales (la matriz de correlaciones $R$ es la identidad).

$$\\chi^2 = -\\left[n - 1 - \\frac{2k + 5}{6}\\right] \\ln |R| \\;\\sim\\; \\chi^2_{k(k-1)/2}$$

donde $k$ es el número de regresores (sin la constante) y $|R|$ el determinante de su matriz de correlaciones. Si $|R| \\to 0$, la multicolinealidad es severa.""")
code("""from scipy import stats

def farrar_glauber(model):
    X = pd.DataFrame(model.model.exog, columns=model.model.exog_names).drop(columns="Intercept")
    n, k = X.shape
    R = np.corrcoef(X.values, rowvar=False)
    detR = np.linalg.det(R)
    chi2 = -(n - 1 - (2 * k + 5) / 6) * np.log(detR)
    gl = k * (k - 1) / 2
    return {"|R|": detR, "chi2": chi2, "gl": gl, "valor crítico 5%": stats.chi2.ppf(0.95, gl), "p-valor": stats.chi2.sf(chi2, gl)}

pd.Series(farrar_glauber(m_full))""")
md("""Como $\\chi^2$ supera con mucho el valor crítico, **rechazamos $H_0$**: hay multicolinealidad en el conjunto de regresores. El determinante $|R|$ cercano a cero lo confirma.""")

md("""### 4.4 Regresiones auxiliares (regla de Klein)

Regresamos cada $x_j$ sobre los demás regresores. **Regla de Klein:** la multicolinealidad es problemática si $R_j^2 > R^2$ del modelo original. Note que $VIF_j = 1/(1-R_j^2)$.""")
code("""X = pd.DataFrame(m_full.model.exog, columns=m_full.model.exog_names)
aux = []
for v in X.columns.drop("Intercept"):
    r2 = sm.OLS(X[v], X.drop(columns=v)).fit().rsquared
    aux.append({"variable": v, "R² auxiliar": r2, "VIF": 1 / (1 - r2), "¿R²_j > R²?": r2 > m_full.rsquared})
aux = pd.DataFrame(aux).sort_values("R² auxiliar", ascending=False).reset_index(drop=True)
print(f"R² del modelo completo = {m_full.rsquared:.3f}")
aux""")
md("""Varios regresores tienen $R_j^2$ mayor que el $R^2$ del modelo (0.26). Esto ocurre porque los salarios son difíciles de predecir, pero los regresores se predicen muy bien entre sí. Según Klein, la multicolinealidad es relevante, y el caso más grave es el de `exper` y `exper²`, con $R_j^2 \\approx 0.95$.""")

md("""### 4.5 Síntomas clásicos
1. **VIF > 20** en `exper` y `exper²`, porque la experiencia y su cuadrado tienen una correlación de 0.975. El número de condición (77) también indica multicolinealidad severa.
2. **Errores estándar inflados**: comparemos el coeficiente de `exper` en el modelo base y en el completo.
3. **Pruebas t individuales débiles, pero la prueba F conjunta es fuerte.**""")
code("""m_base = smf.ols("lwage ~ educ + exper + I(exper**2)", data=df).fit()

def comparar(models, names, vars_):
    rows = []
    for m, n in zip(models, names):
        for v in vars_:
            if v in m.params:
                rows.append({"modelo": n, "variable": v, "coef": m.params[v], "ee": m.bse[v], "t": m.tvalues[v]})
    return pd.DataFrame(rows).pivot(index="variable", columns="modelo")

comparar([m_base, m_full], ["base", "completo"], ["educ", "exper", "I(exper ** 2)"])""")
code("""# Prueba F conjunta: ¿son exper, exper² y age conjuntamente significativas?
print(m_full.f_test("exper = 0, I(exper ** 2) = 0, age = 0"))
print("\\nt individuales:")
print(m_full.tvalues[["exper", "I(exper ** 2)", "age"]])""")
md("""Las variables de ciclo de vida (`exper`, `exper²`, `age`) son **conjuntamente** significativas, aunque individualmente sus t sean pequeñas. Esa es la firma de la multicolinealidad: los datos saben que *el ciclo de vida importa*, pero no pueden separar el efecto de cada variable.""")

md("## 5. Soluciones")
md("""### 5.1 Centrar la experiencia
El VIF de `exper` y `exper²` es en buena parte *artificial* (multicolinealidad "no esencial"). Al centrar, $(exper - \\bar{exper})$ y su cuadrado dejan de estar casi alineados. **El modelo es exactamente el mismo** (mismo R² y mismas predicciones); solo cambia la interpretación del coeficiente lineal, que pasa a ser el efecto marginal en la experiencia media.""")
code("""df["exper_c"] = df["exper"] - df["exper"].mean()
print("corr(exper, exper²)     =", round(np.corrcoef(df.exper, df.exper**2)[0, 1], 3))
print("corr(exper_c, exper_c²) =", round(np.corrcoef(df.exper_c, df.exper_c**2)[0, 1], 3))

f_c = f_full.replace("exper + I(exper**2)", "exper_c + I(exper_c**2)")
m_c = smf.ols(f_c, data=df).fit()
print(f"\\nR² original = {m_full.rsquared:.6f} | R² centrado = {m_c.rsquared:.6f}")
tabla_vif(m_c)""")

md("""### 5.2 Eliminar variables con justificación teórica
- `age` es teóricamente redundante dado `educ` + `exper` (la edad ≈ escolaridad + experiencia + 6), aunque en estos datos su VIF es moderado.
- `IQ` y `KWW` miden lo mismo (habilidad) → usar una sola.
- `meduc` y `feduc` → promediarlas en `peduc` (escolaridad de los padres).""")
code("""df["peduc"] = (df["meduc"] + df["feduc"]) / 2
f_red = "lwage ~ educ + exper_c + I(exper_c**2) + tenure + IQ + peduc + married + black + south + urban"
m_red = smf.ols(f_red, data=df).fit()
print(m_red.summary().tables[1])
tabla_vif(m_red)""")

md("""### 5.3 Componentes principales sobre las variables de habilidad y entorno familiar
Combinamos `IQ`, `KWW`, `meduc` y `feduc` en un índice (primer componente principal).""")
code("""hab = ["IQ", "KWW", "meduc", "feduc"]
Z = StandardScaler().fit_transform(df[hab])
pca = PCA().fit(Z)
print("Varianza explicada por componente:", np.round(pca.explained_variance_ratio_, 3))
print("Cargas del CP1:", dict(zip(hab, np.round(pca.components_[0], 3))))
df["cp1_habilidad"] = pca.transform(Z)[:, 0]

f_pca = "lwage ~ educ + exper_c + I(exper_c**2) + tenure + cp1_habilidad + married + black + south + urban"
m_pca = smf.ols(f_pca, data=df).fit()
print(m_pca.summary().tables[1])""")

md("""### 5.4 Regresión Ridge (extensión)
Ridge minimiza $\\sum u_i^2 + \\lambda \\sum \\beta_j^2$: acepta un poco de sesgo a cambio de reducir mucho la varianza de los estimadores.""")
code("""cols = ["educ", "exper", "tenure", "age", "IQ", "KWW", "meduc", "feduc", "married", "black", "south", "urban"]
Xr = df[cols].assign(exper2=df.exper**2)
sc = StandardScaler().fit(Xr)
ridge = RidgeCV(alphas=np.logspace(-3, 4, 200)).fit(sc.transform(Xr), df.lwage)
print(f"λ óptimo (validación cruzada): {ridge.alpha_:.2f}")

ols_std = sm.OLS(df.lwage, sm.add_constant(sc.transform(Xr))).fit()
pd.DataFrame({"MCO (estand.)": np.asarray(ols_std.params)[1:], "Ridge (estand.)": ridge.coef_}, index=Xr.columns)""")
code("""alphas = np.logspace(-2, 4, 100)
from sklearn.linear_model import Ridge
coefs = [Ridge(alpha=a).fit(sc.transform(Xr), df.lwage).coef_ for a in alphas]
plt.figure(figsize=(9, 5))
plt.plot(alphas, coefs)
plt.xscale("log"); plt.axvline(ridge.alpha_, ls="--", color="k")
plt.xlabel("λ"); plt.ylabel("coeficiente estandarizado"); plt.title("Trayectorias Ridge")
plt.legend(Xr.columns, bbox_to_anchor=(1.02, 1), loc="upper left", fontsize=8)
plt.tight_layout(); plt.show()""")

md("""## 6. Validación del modelo preferido

Elegimos el **modelo reducido** (5.2) como especificación final: VIF bajos, criterios de información iguales o mejores que el modelo completo y coeficientes fáciles de interpretar. Antes de interpretarlo, verificamos los demás supuestos de MCO.""")
md("""### 6.1 Heteroscedasticidad: Breusch-Pagan y White
$H_0$: varianza del error constante (homoscedasticidad).""")
code("""from statsmodels.stats.diagnostic import het_breuschpagan, het_white, linear_reset
from statsmodels.stats.stattools import jarque_bera

bp = het_breuschpagan(m_red.resid, m_red.model.exog)
wh = het_white(m_red.resid, m_red.model.exog)
pd.DataFrame({"estadístico LM": [bp[0], wh[0]], "p-valor": [bp[1], wh[1]]}, index=["Breusch-Pagan", "White"])""")
code("""fig, ax = plt.subplots(1, 2, figsize=(11, 4))
ax[0].scatter(m_red.fittedvalues, m_red.resid, alpha=.4, s=12)
ax[0].axhline(0, color="k", lw=1); ax[0].set_xlabel("valores ajustados"); ax[0].set_ylabel("residuos")
ax[0].set_title("Residuos vs. ajustados")
sm.qqplot(m_red.resid, line="s", ax=ax[1]); ax[1].set_title("Q-Q de residuos")
plt.tight_layout(); plt.show()""")
md("""Las pruebas no coinciden: **Breusch-Pagan rechaza $H_0$ al 5%** (p ≈ 0.02), pero **White no la rechaza** (p ≈ 0.27). White es más general, pero pierde potencia porque incluye muchos términos (cuadrados e interacciones). La advertencia de rango deficiente aparece porque el cuadrado de una variable binaria es igual a la variable misma, y statsmodels lo maneja sin problema. Ante la duda, lo prudente es usar **errores estándar robustos**.""")

md("""### 6.2 Errores estándar robustos (White / HC1)
Los errores robustos son válidos con o sin heteroscedasticidad. Si las conclusiones no cambian al usarlos, los resultados son sólidos.""")
code("""m_red_rob = smf.ols(f_red, data=df).fit(cov_type="HC1")
pd.DataFrame({"coef": m_red.params, "ee MCO": m_red.bse, "ee robusto": m_red_rob.bse,
              "p MCO": m_red.pvalues, "p robusto": m_red_rob.pvalues})""")
md("""### 6.3 Especificación: prueba RESET de Ramsey
$H_0$: el modelo está bien especificado (no faltan potencias de los valores ajustados).""")
code("""reset = linear_reset(m_red, power=3, use_f=True)
print(f"RESET: F = {reset.fvalue:.3f}, p-valor = {reset.pvalue:.4f}")""")
md("""Con errores robustos, la significancia de los coeficientes prácticamente no cambia. La única diferencia es que `black` pasa de p ≈ 0.08 a p ≈ 0.06, que sigue siendo significativo solo al 10%. En la prueba RESET, el p-valor es alto, así que **no hay evidencia de mala especificación** de la forma funcional.""")
md("""### 6.4 Normalidad de los residuos: Jarque-Bera
$H_0$: los residuos son normales. Con $n$ grande, la inferencia es válida asintóticamente aunque se rechace $H_0$.""")
code("""jb, jb_p, asim, curt = jarque_bera(m_red.resid)
print(f"Jarque-Bera = {jb:.2f}, p-valor = {jb_p:.4f}, asimetría = {asim:.3f}, curtosis = {curt:.3f}")""")
md("""### 6.5 Interpretación del modelo final
Como la variable dependiente es $\\ln(w)$, cada coeficiente $\\beta_j \\times 100$ es aproximadamente el **cambio porcentual del salario** ante un aumento de una unidad en $x_j$. Para variables binarias, el efecto exacto es $100\\,(e^{\\beta_j}-1)$.""")
code("""interp = pd.DataFrame({"coef": m_red.params, "efecto % aprox.": 100 * m_red.params,
                       "efecto % exacto": 100 * (np.exp(m_red.params) - 1), "p-valor": m_red_rob.pvalues}).drop("Intercept")
interp""")

md("## 7. Comparación de modelos")
code("""modelos = {"Base": m_base, "Completo": m_full, "Centrado": m_c, "Reducido": m_red, "PCA": m_pca}
resumen = pd.DataFrame({
    n: {"coef educ": m.params["educ"], "ee educ": m.bse["educ"],
        "R²": m.rsquared, "R² aj.": m.rsquared_adj, "AIC": m.aic, "BIC": m.bic,
        "VIF máx": tabla_vif(m).VIF.max(), "k": int(m.df_model)}
    for n, m in modelos.items()
}).T
resumen""")

md("""## 8. Conclusiones

1. **Multicolinealidad perfecta**: incluir `age`, `educ` y la experiencia potencial `age − educ − 6` hace que $X'X$ sea singular. Los coeficientes no están identificados y la única solución es quitar una variable.
2. **Multicolinealidad imperfecta**: el modelo completo muestra VIF > 20 en `exper` y `exper²`, errores estándar inflados y t individuales pequeñas pese a una F conjunta significativa. **MCO sigue siendo MELI**: la multicolinealidad no sesga los estimadores, solo los vuelve imprecisos.
3. **Centrar** elimina la multicolinealidad "no esencial" entre `exper` y `exper²` sin cambiar el ajuste del modelo.
4. **Eliminar variables redundantes** (`age`) o **combinarlas** (promedio de escolaridad de los padres, CP1 de habilidad) reduce los VIF con poca pérdida de R² y criterios de información iguales o mejores.
5. El **retorno a la educación** se mantiene positivo y significativo en todas las especificaciones, pero baja de ≈7.8% (modelo base) a ≈4–5% al controlar por habilidad y entorno familiar (sesgo por variable omitida), que es otra lección importante.
6. **Ridge** encoge los coeficientes de las variables colineales y estabiliza las estimaciones a cambio de algo de sesgo.
7. **Validación del modelo final:**
   - Hay indicios de heteroscedasticidad (Breusch-Pagan), pero los errores robustos no cambian las conclusiones.
   - La prueba RESET no detecta mala especificación.
   - Los residuos no son normales (Jarque-Bera), pero con n = 722 la inferencia es válida asintóticamente.
8. **Interpretación del modelo final:** un año más de educación eleva el salario ≈ 4.9%, estar casado ≈ 22% y vivir en zona urbana ≈ 20%, manteniendo lo demás constante.

**Referencias**
- Wooldridge, J. M. *Introducción a la econometría: un enfoque moderno*. Cap. 3 y 6.
- Gujarati, D. y Porter, D. *Econometría*. Cap. 10, "Multicolinealidad".
- Mincer, J. (1974). *Schooling, Experience, and Earnings*. NBER.
- Blackburn, M. y Neumark, D. (1992). Datos originales de `wage2`.""")

nb["cells"] = C
nb["metadata"]["kernelspec"] = {"name": "python3", "display_name": "Python 3", "language": "python"}
nbf.write(nb, sys.argv[1])
