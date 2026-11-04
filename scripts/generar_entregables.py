"""Genera las figuras, el reporte (Word) y la presentación (PowerPoint) del proyecto.

Uso (desde la carpeta del proyecto):
    .venv/bin/python scripts/generar_entregables.py
"""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
import statsmodels.api as sm
import statsmodels.formula.api as smf
import wooldridge
from scipy import stats
from sklearn.linear_model import Ridge, RidgeCV
from sklearn.preprocessing import StandardScaler
from statsmodels.stats.diagnostic import het_breuschpagan, het_white, linear_reset
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.stats.stattools import jarque_bera

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches as PI, Pt as PPt

ROOT = Path(__file__).resolve().parent.parent
FIG = ROOT / "figuras"
FIG.mkdir(exist_ok=True)
AUTORA = "Ximena Lara"
CURSO = "Fundamentos de Econometría"
FECHA = "Noviembre 2026"

# ---------------------------------------------------------------- análisis
vars_ = ["wage", "lwage", "educ", "exper", "tenure", "age", "IQ", "KWW", "meduc", "feduc"]
df = wooldridge.data("wage2")[vars_ + ["married", "black", "south", "urban"]].dropna().reset_index(drop=True)
n = len(df)
df["pexper"] = df.age - df.educ - 6
df["exper_c"] = df.exper - df.exper.mean()
df["peduc"] = (df.meduc + df.feduc) / 2


def vif(m):
    X = pd.DataFrame(m.model.exog, columns=m.model.exog_names)
    v = pd.Series([variance_inflation_factor(X.values, i) for i in range(X.shape[1])], index=X.columns)
    return v.drop("Intercept").sort_values(ascending=False)


NOMBRES = {"I(exper ** 2)": "exper²", "I(exper_c ** 2)": "exper_c²"}
nom = lambda s: NOMBRES.get(s, s)

f_base = "lwage ~ educ + exper + I(exper**2)"
f_full = "lwage ~ educ + exper + I(exper**2) + tenure + age + IQ + KWW + meduc + feduc + married + black + south + urban"
f_c = f_full.replace("exper + I(exper**2)", "exper_c + I(exper_c**2)")
f_red = "lwage ~ educ + exper_c + I(exper_c**2) + tenure + IQ + peduc + married + black + south + urban"

m_base = smf.ols(f_base, df).fit()
m_full = smf.ols(f_full, df).fit()
m_c = smf.ols(f_c, df).fit()
m_red = smf.ols(f_red, df).fit()
m_red_rob = smf.ols(f_red, df).fit(cov_type="HC1")

hab = ["IQ", "KWW", "meduc", "feduc"]
from sklearn.decomposition import PCA
Z = StandardScaler().fit_transform(df[hab])
pca = PCA().fit(Z)
df["cp1"] = pca.transform(Z)[:, 0]
m_pca = smf.ols("lwage ~ educ + exper_c + I(exper_c**2) + tenure + cp1 + married + black + south + urban", df).fit()

# multicolinealidad perfecta
Xp = sm.add_constant(df[["educ", "age", "pexper"]]).values
rango_p = np.linalg.matrix_rank(Xp)

# diagnósticos
vif_full, vif_c, vif_red = vif(m_full), vif(m_c), vif(m_red)
Xf = m_full.model.exog
cond = np.linalg.cond(Xf / np.linalg.norm(Xf, axis=0))
Xd = pd.DataFrame(Xf, columns=m_full.model.exog_names).drop(columns="Intercept")
k = Xd.shape[1]
detR = np.linalg.det(np.corrcoef(Xd.values, rowvar=False))
fg_chi2 = -(n - 1 - (2 * k + 5) / 6) * np.log(detR)
fg_gl = k * (k - 1) / 2
fg_crit = stats.chi2.ppf(0.95, fg_gl)
r2_aux = 1 - 1 / vif_full
ftest = m_full.f_test("exper = 0, I(exper ** 2) = 0, age = 0")
corr_e = np.corrcoef(df.exper, df.exper ** 2)[0, 1]
corr_ec = np.corrcoef(df.exper_c, df.exper_c ** 2)[0, 1]

# validación
bp = het_breuschpagan(m_red.resid, m_red.model.exog)
wh = het_white(m_red.resid, m_red.model.exog)
reset = linear_reset(m_red, power=3, use_f=True)
jb, jb_p, asim, curt = jarque_bera(m_red.resid)

# ridge
cols = ["educ", "exper", "tenure", "age", "IQ", "KWW", "meduc", "feduc", "married", "black", "south", "urban"]
Xr = df[cols].assign(exper2=df.exper ** 2)
sc = StandardScaler().fit(Xr)
ridge = RidgeCV(alphas=np.logspace(-3, 4, 200)).fit(sc.transform(Xr), df.lwage)
ols_std = sm.OLS(df.lwage, sm.add_constant(sc.transform(Xr))).fit()

modelos = {"Base": m_base, "Completo": m_full, "Centrado": m_c, "Reducido": m_red, "PCA": m_pca}
resumen = pd.DataFrame({nm: {"β educ": m.params["educ"], "e.e. educ": m.bse["educ"], "R²": m.rsquared,
                             "R² aj.": m.rsquared_adj, "AIC": m.aic, "BIC": m.bic,
                             "VIF máx": vif(m).max(), "k": int(m.df_model)}
                        for nm, m in modelos.items()}).T

# ---------------------------------------------------------------- figuras
sns.set_theme(style="whitegrid", font_scale=1.05)
AZUL, ROJO, GRIS = "#2563a8", "#c0392b", "#6b7280"

plt.figure(figsize=(7.5, 6))
X_vars = ["educ", "exper", "tenure", "age", "IQ", "KWW", "meduc", "feduc"]
sns.heatmap(df[X_vars].corr(), annot=True, fmt=".2f", cmap="RdBu_r", vmin=-1, vmax=1, square=True, cbar_kws={"shrink": .8})
plt.title("Correlaciones entre regresores")
plt.tight_layout(); plt.savefig(FIG / "correlaciones.png", dpi=200); plt.close()

fig, ax = plt.subplots(figsize=(9, 5))
comp = pd.DataFrame({"Completo": vif_full.rename(nom)})
comp["Centrado"] = vif_c.rename(lambda s: nom(s).replace("_c", "")).reindex(comp.index)
comp = comp.sort_values("Completo")
y = np.arange(len(comp))
ax.barh(y + .2, comp.Completo, .4, color=ROJO, label="Modelo completo")
ax.barh(y - .2, comp.Centrado, .4, color=AZUL, label="Experiencia centrada")
ax.set_yticks(y, comp.index); ax.axvline(10, ls="--", color="k", lw=1)
ax.text(10.3, 0, "umbral = 10", fontsize=9)
ax.set_xlabel("VIF"); ax.set_title("VIF antes y después de centrar la experiencia"); ax.legend(loc="lower right")
plt.tight_layout(); plt.savefig(FIG / "vif.png", dpi=200); plt.close()

alphas = np.logspace(-2, 4, 100)
coefs = np.array([Ridge(alpha=a).fit(sc.transform(Xr), df.lwage).coef_ for a in alphas])
fig, ax = plt.subplots(figsize=(9, 5))
for j, c in enumerate(Xr.columns):
    hl = c in ("exper", "exper2")
    ax.plot(alphas, coefs[:, j], lw=2.4 if hl else 1, color=(ROJO if c == "exper" else "#e67e22") if hl else GRIS,
            alpha=1 if hl else .6, label=("exper²" if c == "exper2" else c) if hl else None)
ax.set_xscale("log"); ax.axvline(ridge.alpha_, ls="--", color="k", lw=1)
ax.text(ridge.alpha_ * 1.1, ax.get_ylim()[1] * .9, f"λ óptimo ≈ {ridge.alpha_:.0f}", fontsize=9)
ax.set_xlabel("λ (penalización)"); ax.set_ylabel("coeficiente estandarizado")
ax.set_title("Trayectorias Ridge: los coeficientes colineales se encogen primero"); ax.legend()
plt.tight_layout(); plt.savefig(FIG / "ridge.png", dpi=200); plt.close()

fig, ax = plt.subplots(1, 2, figsize=(11, 4.2))
ax[0].scatter(m_red.fittedvalues, m_red.resid, alpha=.35, s=12, color=AZUL)
ax[0].axhline(0, color="k", lw=1); ax[0].set_xlabel("valores ajustados"); ax[0].set_ylabel("residuos")
ax[0].set_title("Residuos vs. ajustados")
sm.qqplot(m_red.resid, line="s", ax=ax[1], markersize=3, alpha=.5); ax[1].set_title("Q-Q de residuos")
plt.tight_layout(); plt.savefig(FIG / "residuos.png", dpi=200); plt.close()

fig, ax = plt.subplots(figsize=(8, 4.5))
cmp = resumen["β educ"]; ee = resumen["e.e. educ"]
ax.errorbar(cmp.index, cmp, yerr=1.96 * ee, fmt="o", color=AZUL, capsize=6, ms=8)
ax.set_ylabel("β educ (IC 95%)"); ax.set_title("Retorno a la educación según el modelo")
plt.tight_layout(); plt.savefig(FIG / "educ.png", dpi=200); plt.close()

# ---------------------------------------------------------------- Word
f3 = lambda x: f"{x:.3f}"
f4 = lambda x: f"{x:.4f}"


def tabla(doc, df_, fmt=f3, index_name=""):
    t = doc.add_table(rows=1, cols=len(df_.columns) + 1)
    t.style = "Light Grid Accent 1"
    hdr = t.rows[0].cells
    hdr[0].text = index_name
    for j, c in enumerate(df_.columns):
        hdr[j + 1].text = str(c)
    for idx, row in df_.iterrows():
        cells = t.add_row().cells
        cells[0].text = str(idx)
        for j, v in enumerate(row):
            cells[j + 1].text = fmt(v) if isinstance(v, (float, np.floating)) else str(v)
    for r in t.rows:
        for c in r.cells:
            for p in c.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9.5)
    return t


def pie(doc, txt):
    p = doc.add_paragraph(txt)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].italic = True
    p.runs[0].font.size = Pt(9.5)


def ecuacion(doc, txt):
    p = doc.add_paragraph(txt)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].italic = True


doc = Document()
st = doc.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(11)
for s in doc.sections:
    s.left_margin = s.right_margin = Inches(1)

for _ in range(6):
    doc.add_paragraph()
t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("Multicolinealidad en la ecuación de salarios de Mincer"); r.bold = True; r.font.size = Pt(22)
t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("Diagnóstico y corrección con datos de la National Longitudinal Survey"); r.font.size = Pt(14)
r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
for _ in range(4):
    doc.add_paragraph()
for line in (AUTORA, CURSO, FECHA):
    p = doc.add_paragraph(line); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_page_break()

doc.add_heading("Resumen", 1)
doc.add_paragraph(
    f"Este trabajo estima una ecuación de salarios tipo Mincer con datos de {n} hombres trabajadores de EE. UU. (base wage2 de Wooldridge) "
    "para estudiar la multicolinealidad. Primero se muestra un caso de multicolinealidad perfecta: al incluir la edad, la escolaridad y la "
    f"experiencia potencial (edad − escolaridad − 6), la matriz de regresores tiene rango {rango_p} de 4 columnas y el modelo no está identificado. "
    f"Después, en un modelo con 13 regresores, se detecta multicolinealidad imperfecta: VIF de hasta {vif_full.max():.1f}, número de condición de "
    f"{cond:.0f}, la prueba de Farrar-Glauber rechaza la ortogonalidad (χ² = {fg_chi2:.0f}) y la experiencia y la edad no son significativas "
    f"individualmente, pero sí de forma conjunta (F = {ftest.fvalue:.2f}, p = {float(ftest.pvalue):.3f}). Se comparan cuatro soluciones: centrar variables, "
    "eliminar o combinar regresores redundantes, componentes principales y regresión Ridge. En el modelo final, todos los VIF son menores a 2 "
    f"y un año adicional de educación se asocia con un salario {100*(np.exp(m_red.params['educ'])-1):.1f}% mayor.")

doc.add_heading("1. Introducción", 1)
doc.add_paragraph(
    "La ecuación de Mincer (1974) es uno de los modelos más utilizados en economía laboral: relaciona el logaritmo del salario con la "
    "escolaridad y la experiencia. En la práctica se suelen agregar controles como la habilidad, el entorno familiar o la región. Muchas "
    "de esas variables están relacionadas entre sí: las personas con más escolaridad tienden a tener mayor puntaje en pruebas de habilidad "
    "y padres más educados, y la edad es casi la suma de la escolaridad y la experiencia. Esta situación, en la que los regresores se mueven juntos, "
    "se conoce como multicolinealidad.")
doc.add_paragraph(
    "El objetivo de este proyecto es (i) mostrar con datos reales la diferencia entre multicolinealidad perfecta e imperfecta, "
    "(ii) aplicar las herramientas de diagnóstico vistas en el curso y (iii) comparar distintas estrategias de corrección y sus consecuencias "
    "para la interpretación del retorno a la educación.")

doc.add_heading("2. Marco teórico", 1)
doc.add_heading("2.1 La ecuación de Mincer", 2)
ecuacion(doc, "ln(wᵢ) = β₀ + β₁·educᵢ + β₂·experᵢ + β₃·experᵢ² + uᵢ")
doc.add_paragraph(
    "β₁ mide el retorno porcentual aproximado de un año más de escolaridad. El término cuadrático permite que el salario crezca con la "
    "experiencia a tasa decreciente, de modo que el perfil edad-ingreso tiene forma de U invertida.")
doc.add_heading("2.2 Multicolinealidad", 2)
doc.add_paragraph(
    "Hay multicolinealidad perfecta cuando un regresor es combinación lineal exacta de otros. En ese caso X'X no es invertible y "
    "los estimadores de MCO no existen de forma única; esto viola el supuesto de rango completo. Hay multicolinealidad imperfecta cuando "
    "la relación es fuerte pero no exacta. En ese caso, MCO sigue siendo insesgado y MELI (teorema de Gauss-Markov), pero las varianzas de los estimadores "
    "se inflan:")
ecuacion(doc, "Var(β̂ⱼ) = σ² / [ SSTⱼ · (1 − Rⱼ²) ]")
doc.add_paragraph(
    "donde Rⱼ² es el R² de regresar xⱼ sobre los demás regresores. Las consecuencias típicas son errores estándar grandes, estadísticos t "
    "pequeños con un R² y una prueba F conjunta significativos, coeficientes con signos o magnitudes poco razonables, y estimaciones muy sensibles "
    "a cambios en la muestra o la especificación.")
doc.add_heading("2.3 Herramientas de diagnóstico", 2)
for b in ["Factor de Inflación de Varianza: VIFⱼ = 1/(1 − Rⱼ²). Un valor mayor a 10 suele considerarse problemático.",
          "Número de condición de la matriz X normalizada (Belsley, Kuh y Welsch): entre 10 y 30 indica multicolinealidad moderada y más de 30, severa.",
          "Prueba de Farrar-Glauber: χ² = −[n − 1 − (2k + 5)/6]·ln|R|, con k(k − 1)/2 grados de libertad, bajo H₀ de regresores ortogonales.",
          "Regla de Klein: la multicolinealidad es preocupante si algún Rⱼ² auxiliar supera el R² del modelo."]:
    doc.add_paragraph(b, style="List Bullet")

doc.add_heading("3. Datos", 1)
doc.add_paragraph(
    "Se utiliza la base wage2 (Blackburn y Neumark, 1992), incluida en Wooldridge (2020), con información de la National "
    f"Longitudinal Survey de 1980 sobre hombres empleados. Tras eliminar observaciones con escolaridad de los padres faltante, la muestra es de n = {n}.")
desc = df[["wage", "educ", "exper", "tenure", "age", "IQ", "KWW", "meduc", "feduc"]].describe().T[["mean", "std", "min", "max"]]
desc.columns = ["Media", "Desv. est.", "Mín.", "Máx."]
tabla(doc, desc, fmt=lambda x: f"{x:.2f}", index_name="Variable")
pie(doc, "Tabla 1. Estadística descriptiva. wage: salario mensual (USD de 1980); educ, exper, tenure, age: años; "
         "IQ: coeficiente intelectual; KWW: prueba de conocimiento del mundo laboral; meduc, feduc: escolaridad de la madre y del padre.")
doc.add_picture(str(FIG / "correlaciones.png"), width=Inches(4.6))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
pie(doc, "Figura 1. Matriz de correlaciones entre regresores.")
doc.add_paragraph(
    "Las correlaciones más altas se dan entre las medidas de habilidad (IQ y KWW), entre la escolaridad de los padres y entre educ e IQ. "
    "Ninguna correlación bivariada es extrema. Sin embargo, la multicolinealidad puede surgir de combinaciones de varios regresores, por lo que "
    "las correlaciones simples no bastan como diagnóstico.")

doc.add_heading("4. Metodología", 1)
for b in ["Caso de multicolinealidad perfecta: regresión de ln(w) sobre educ, age y la experiencia potencial pexper = age − educ − 6.",
          "Modelo completo: 13 regresores (educ, exper, exper², tenure, age, IQ, KWW, meduc, feduc, married, black, south, urban).",
          "Diagnóstico: VIF, número de condición, Farrar-Glauber, regresiones auxiliares y contraste entre las pruebas t individuales y la F conjunta.",
          "Corrección: (a) centrar la experiencia; (b) eliminar age y combinar meduc y feduc en su promedio, conservando solo IQ como medida de habilidad; "
          "(c) primer componente principal de IQ, KWW, meduc y feduc; (d) regresión Ridge con λ elegido por validación cruzada.",
          "Validación del modelo final: Breusch-Pagan y White (heteroscedasticidad), errores robustos HC1, RESET de Ramsey (especificación) y "
          "Jarque-Bera (normalidad)."]:
    doc.add_paragraph(b, style="List Number")

doc.add_heading("5. Resultados", 1)
doc.add_heading("5.1 Multicolinealidad perfecta", 2)
doc.add_paragraph(
    f"La matriz X = [1, educ, age, pexper] tiene rango {rango_p}, menor que sus 4 columnas, porque pexper = age − educ − 6 de forma exacta. "
    "X'X es singular, así que el sistema de ecuaciones normales tiene infinitas soluciones. El software puede reportar números usando "
    "una pseudo-inversa, pero esos coeficientes no tienen interpretación: no es posible variar la edad manteniendo fijas la escolaridad y la "
    "experiencia potencial. La única solución es excluir una de las tres variables.")

doc.add_heading("5.2 Multicolinealidad imperfecta: diagnóstico", 2)
t_vif = pd.DataFrame({"VIF": vif_full.values, "Rⱼ² auxiliar": r2_aux.values}, index=[nom(i) for i in vif_full.index])
tabla(doc, t_vif, fmt=lambda x: f"{x:.3f}" if x < 1 else f"{x:.2f}", index_name="Variable")
pie(doc, f"Tabla 2. VIF y R² auxiliares del modelo completo (R² del modelo = {m_full.rsquared:.3f}).")
doc.add_paragraph(
    f"La experiencia y su cuadrado tienen VIF de {vif_full['exper']:.1f} y {vif_full['I(exper ** 2)']:.1f}, muy por encima del umbral de 10, "
    f"porque su correlación es de {corr_e:.3f}. El número de condición es {cond:.1f}, lo que indica multicolinealidad severa. La prueba de Farrar-Glauber "
    f"da |R| = {detR:.4f} y χ² = {fg_chi2:.1f}, frente a un valor crítico de {fg_crit:.1f} con {fg_gl:.0f} grados de libertad, por lo que se rechaza "
    f"con claridad la hipótesis de regresores ortogonales. Ocho regresores tienen un Rⱼ² auxiliar mayor que el R² del modelo ({m_full.rsquared:.3f}), "
    "por lo que la regla de Klein también señala multicolinealidad.")
sint = pd.DataFrame({
    "Coef. (base)": [m_base.params[v] for v in ["educ", "exper", "I(exper ** 2)"]],
    "t (base)": [m_base.tvalues[v] for v in ["educ", "exper", "I(exper ** 2)"]],
    "Coef. (completo)": [m_full.params[v] for v in ["educ", "exper", "I(exper ** 2)"]],
    "t (completo)": [m_full.tvalues[v] for v in ["educ", "exper", "I(exper ** 2)"]],
}, index=["educ", "exper", "exper²"])
tabla(doc, sint, fmt=f4, index_name="Variable")
pie(doc, "Tabla 3. Coeficientes y estadísticos t en los modelos base y completo.")
doc.add_paragraph(
    f"El síntoma clásico aparece con claridad: en el modelo completo, exper (t = {m_full.tvalues['exper']:.2f}), exper² "
    f"(t = {m_full.tvalues['I(exper ** 2)']:.2f}) y age (t = {m_full.tvalues['age']:.2f}) no son significativas individualmente, pero la prueba F "
    f"conjunta las rechaza (F = {ftest.fvalue:.2f}, p = {float(ftest.pvalue):.4f}). Los datos indican que el ciclo de vida laboral importa, pero no "
    "permiten separar el efecto de cada una de estas variables.")
doc.add_picture(str(FIG / "vif.png"), width=Inches(5.8))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
pie(doc, "Figura 2. VIF del modelo completo y del modelo con experiencia centrada.")

doc.add_heading("5.3 Corrección", 2)
doc.add_paragraph(
    f"(a) Centrar. Al restar la media a exper, la correlación entre la variable y su cuadrado baja de {corr_e:.3f} a {corr_ec:.3f}, y el VIF máximo "
    f"del modelo pasa de {vif_full.max():.1f} a {vif_c.max():.2f}. El modelo es algebraicamente el mismo: el R² ({m_c.rsquared:.4f}) y las "
    "predicciones no cambian. Esto muestra que esa parte de la multicolinealidad era no esencial, producto de cómo se escribió el modelo.")
doc.add_paragraph(
    f"(b) Eliminar y combinar. Se excluye age, redundante dados educ y exper, y se combinan meduc y feduc en su promedio. El modelo reducido "
    f"tiene un VIF máximo de {vif_red.max():.2f}, conserva prácticamente el mismo R² ajustado ({m_red.rsquared_adj:.4f} frente a "
    f"{m_full.rsquared_adj:.4f}) y mejora el BIC ({m_red.bic:.1f} frente a {m_full.bic:.1f}).")
doc.add_paragraph(
    f"(c) Componentes principales. El primer componente de IQ, KWW, meduc y feduc explica el {100*pca.explained_variance_ratio_[0]:.1f}% de su "
    "varianza y tiene cargas positivas y similares en las cuatro variables, por lo que puede leerse como un índice de habilidad y entorno familiar. "
    f"El modelo con este índice obtiene el menor AIC ({m_pca.aic:.1f}) y BIC ({m_pca.bic:.1f}).")
doc.add_paragraph(
    f"(d) Ridge. Con λ ≈ {ridge.alpha_:.0f} (validación cruzada), el coeficiente estandarizado de exper se reduce de "
    f"{np.asarray(ols_std.params)[1+list(Xr.columns).index('exper')]:.3f} (MCO) a {ridge.coef_[list(Xr.columns).index('exper')]:.3f}. Ridge acepta algo de sesgo "
    "a cambio de una gran reducción de la varianza, y encoge primero a los regresores colineales.")
doc.add_picture(str(FIG / "ridge.png"), width=Inches(5.8))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
pie(doc, "Figura 3. Trayectorias de los coeficientes estandarizados en la regresión Ridge.")
t_res = resumen.copy(); t_res["k"] = t_res["k"].astype(int).astype(str)
tabla(doc, t_res, fmt=lambda x: f"{x:.4f}" if abs(x) < 1 else f"{x:.1f}", index_name="Modelo")
pie(doc, "Tabla 4. Comparación de especificaciones.")
doc.add_picture(str(FIG / "educ.png"), width=Inches(5.2))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
pie(doc, "Figura 4. Retorno a la educación estimado en cada modelo (intervalo de confianza al 95%).")
doc.add_paragraph(
    f"El retorno a la educación baja de {m_base.params['educ']:.3f} en el modelo base a entre {m_full.params['educ']:.3f} y "
    f"{m_red.params['educ']:.3f} al controlar por habilidad y entorno familiar. Esto no se debe a la multicolinealidad, sino al sesgo por variable "
    "omitida del modelo base: la educación captaba en parte el efecto de la habilidad. En cambio, sí es efecto de la multicolinealidad que el "
    f"error estándar de educ aumente de {m_base.bse['educ']:.4f} a {m_full.bse['educ']:.4f} al pasar del modelo base al completo.")

doc.add_heading("5.4 Validación del modelo final", 2)
val = pd.DataFrame({"Estadístico": [bp[0], wh[0], reset.fvalue, jb], "p-valor": [bp[1], wh[1], reset.pvalue, jb_p],
                    "Conclusión (5%)": ["Rechaza H₀", "No rechaza H₀", "No rechaza H₀", "Rechaza H₀"]},
                   index=["Breusch-Pagan", "White", "RESET (Ramsey)", "Jarque-Bera"])
tabla(doc, val, index_name="Prueba")
pie(doc, "Tabla 5. Pruebas de validación del modelo reducido.")
doc.add_paragraph(
    "Breusch-Pagan sugiere heteroscedasticidad, mientras que White, más general pero con menor potencia, no la detecta. Por precaución se "
    "reportan errores robustos (HC1), que prácticamente no cambian las conclusiones. La prueba RESET no detecta problemas de forma funcional. Los "
    f"residuos no son normales (asimetría {asim:.2f}, curtosis {curt:.2f}), pero con n = {n} la inferencia es válida asintóticamente.")
doc.add_picture(str(FIG / "residuos.png"), width=Inches(6.2))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
pie(doc, "Figura 5. Diagnóstico gráfico de los residuos del modelo final.")

doc.add_heading("5.5 Interpretación", 2)
ip = m_red.params.drop("Intercept")
it = pd.DataFrame({"Coeficiente": ip, "e.e. robusto": m_red_rob.bse.drop("Intercept"),
                   "p-valor": m_red_rob.pvalues.drop("Intercept"), "Efecto % exacto": 100 * (np.exp(ip) - 1)})
it.index = [nom(i) for i in it.index]
tabla(doc, it, fmt=f4, index_name="Variable")
pie(doc, f"Tabla 6. Modelo final (variable dependiente: ln(salario); n = {n}; R² = {m_red.rsquared:.3f}; errores robustos HC1).")
e = lambda v: 100 * (np.exp(m_red.params[v]) - 1)
doc.add_paragraph(
    f"Manteniendo lo demás constante: un año más de escolaridad se asocia con un salario {e('educ'):.1f}% mayor; un año de experiencia (en la "
    f"experiencia media), con {e('exper_c'):.1f}% más; un año de antigüedad con el empleador, con {e('tenure'):.1f}% más; y 10 puntos de IQ, con "
    f"{100*(np.exp(10*m_red.params['IQ'])-1):.1f}% más. Los hombres casados ganan {e('married'):.1f}% más y quienes viven en zonas urbanas "
    f"{e('urban'):.1f}% más; vivir en el sur se asocia con {abs(e('south')):.1f}% menos. El término cuadrático de experiencia no es significativo "
    "en esta muestra, posiblemente porque la experiencia observada tiene poco rango (de 1 a 23 años).")

doc.add_heading("6. Conclusiones", 1)
for b in ["La multicolinealidad perfecta (pexper = age − educ − 6) impide identificar el modelo; la única solución es eliminar una variable.",
          "La multicolinealidad imperfecta no sesga MCO, pero infla los errores estándar. En el modelo completo produce t individuales no "
          "significativas para variables que son significativas en conjunto.",
          "Los cuatro diagnósticos (VIF, número de condición, Farrar-Glauber y Klein) coinciden en que el problema principal está en exper y exper².",
          "Centrar elimina la multicolinealidad no esencial sin cambiar el ajuste. Eliminar o combinar variables redundantes y usar componentes "
          "principales producen modelos más parsimoniosos con criterios de información iguales o mejores.",
          f"El retorno a la educación es robusto en todas las especificaciones (≈ {100*m_red.params['educ']:.0f}–{100*m_pca.params['educ']:.0f}% al "
          "controlar por habilidad). Su caída respecto al modelo base refleja un sesgo por variable omitida, no un efecto de la multicolinealidad."]:
    doc.add_paragraph(b, style="List Bullet")
doc.add_paragraph(
    "Una limitación es que la muestra solo incluye hombres de 1980, por lo que los resultados no se generalizan a otras poblaciones. "
    "Una extensión natural sería replicar el análisis con datos mexicanos (ENOE, INEGI).")

doc.add_heading("Referencias", 1)
for ref in ["Belsley, D., Kuh, E. y Welsch, R. (1980). Regression Diagnostics. Wiley.",
            "Blackburn, M. y Neumark, D. (1992). Unobserved ability, efficiency wages, and interindustry wage differentials. "
            "Quarterly Journal of Economics, 107(4), 1421–1436.",
            "Farrar, D. y Glauber, R. (1967). Multicollinearity in regression analysis: the problem revisited. "
            "Review of Economics and Statistics, 49(1), 92–107.",
            "Gujarati, D. y Porter, D. (2010). Econometría (5.ª ed.). McGraw-Hill. Cap. 10.",
            "Mincer, J. (1974). Schooling, Experience, and Earnings. NBER.",
            "Wooldridge, J. M. (2020). Introducción a la econometría: un enfoque moderno (7.ª ed.). Cengage. Caps. 3 y 6."]:
    doc.add_paragraph(ref)

doc.save(ROOT / "Reporte_Multicolinealidad_Mincer.docx")

# ---------------------------------------------------------------- PowerPoint
OSC = PRGB(0x1F, 0x2A, 0x44); AC = PRGB(0x25, 0x63, 0xA8); TXT = PRGB(0x33, 0x33, 0x33); BL = PRGB(0xFF, 0xFF, 0xFF)
RO = PRGB(0xC0, 0x39, 0x2B)
prs = Presentation(); prs.slide_width = PI(13.333); prs.slide_height = PI(7.5)
BLANK = prs.slide_layouts[6]


def caja(sl, x, y, w, h, texto, size=20, color=TXT, bold=False, align=None):
    tb = sl.shapes.add_textbox(PI(x), PI(y), PI(w), PI(h)); tf = tb.text_frame; tf.word_wrap = True
    lines = texto if isinstance(texto, list) else [texto]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = ln; p.font.size = PPt(size); p.font.color.rgb = color; p.font.bold = bold
        p.space_after = PPt(8)
        if align: p.alignment = align
    return tb


def slide(titulo, num=None):
    sl = prs.slides.add_slide(BLANK)
    bar = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, PI(1.1))
    bar.fill.solid(); bar.fill.fore_color.rgb = OSC; bar.line.fill.background()
    caja(sl, .6, .22, 12, .8, titulo, 30, BL, True)
    if num:
        caja(sl, 12.2, 6.9, 1, .4, str(num), 12, PRGB(0x99, 0x99, 0x99))
    return sl


def bullets(sl, items, x=.7, y=1.5, w=12, h=5.5, size=21):
    return caja(sl, x, y, w, h, ["•  " + i for i in items], size)


def ptabla(sl, df_, x, y, w, fmt=f3, size=14, index_name=""):
    rows, cols = len(df_) + 1, len(df_.columns) + 1
    tb = sl.shapes.add_table(rows, cols, PI(x), PI(y), PI(w), PI(.4 * rows)).table
    tb.cell(0, 0).text = index_name
    for j, c in enumerate(df_.columns): tb.cell(0, j + 1).text = str(c)
    for i, (idx, row) in enumerate(df_.iterrows(), 1):
        tb.cell(i, 0).text = str(idx)
        for j, v in enumerate(row):
            tb.cell(i, j + 1).text = fmt(v) if isinstance(v, (float, np.floating)) else str(v)
    for r in range(rows):
        for c in range(cols):
            for p in tb.cell(r, c).text_frame.paragraphs: p.font.size = PPt(size)


def dato(sl, x, y, valor, etiqueta, color=AC):
    caja(sl, x, y, 3.8, 1, valor, 44, color, True)
    caja(sl, x, y + 1, 3.8, 1, etiqueta, 15)


# 1 portada
sl = prs.slides.add_slide(BLANK)
bg = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
bg.fill.solid(); bg.fill.fore_color.rgb = OSC; bg.line.fill.background()
caja(sl, .9, 2.2, 11.5, 1.5, "Multicolinealidad en la ecuación de salarios de Mincer", 42, BL, True)
caja(sl, .9, 3.8, 11.5, .8, "Diagnóstico y corrección con datos reales (NLS, wage2)", 22, PRGB(0xBF, 0xD4, 0xEE))
caja(sl, .9, 5.4, 11.5, 1.2, [AUTORA, f"{CURSO}  ·  {FECHA}"], 18, BL)

# 2 pregunta
sl = slide("Pregunta y motivación", 2)
bullets(sl, ["¿Cuánto aumenta el salario un año más de educación?",
             "Para responderla controlamos por experiencia, edad, habilidad y entorno familiar…",
             "…pero esos controles están muy relacionados entre sí → multicolinealidad",
             "Objetivo: diagnosticarla con datos reales, corregirla y ver qué cambia en las conclusiones"])
caja(sl, .7, 5.6, 12, .8, "ln(wᵢ) = β₀ + β₁·educᵢ + β₂·experᵢ + β₃·experᵢ² + uᵢ", 26, AC, True, )

# 3 datos
sl = slide("Datos: wage2 (Wooldridge)", 3)
bullets(sl, [f"National Longitudinal Survey, EE. UU., 1980: {n} hombres empleados",
             "Salario mensual, educación, experiencia, antigüedad, edad",
             "Habilidad: IQ y KWW",
             "Entorno familiar: escolaridad de la madre y del padre",
             "Correlaciones altas: IQ–KWW, meduc–feduc, educ–IQ"], w=6.3, size=19)
sl.shapes.add_picture(str(FIG / "correlaciones.png"), PI(7.2), PI(1.3), height=PI(5.9))

# 4 perfecta
sl = slide("Multicolinealidad perfecta", 4)
caja(sl, .7, 1.5, 12, .8, "Experiencia potencial:  pexper = age − educ − 6", 26, AC, True)
bullets(sl, ["Si incluimos age, educ y pexper juntas, una es combinación lineal exacta de las otras",
             f"Rango de X = {rango_p} de 4 columnas  →  X'X no es invertible",
             "Los coeficientes no están identificados: no se puede cambiar la edad dejando fijas las otras dos",
             "Solución: eliminar una de las tres variables (Stata o R lo hacen automáticamente)"], y=2.6)

# 5 síntomas
sl = slide("Multicolinealidad imperfecta: síntomas", 5)
dato(sl, .7, 1.5, f"{vif_full.max():.0f}", "VIF máximo (exper²)\numbral = 10", RO)
dato(sl, 4.7, 1.5, f"{cond:.0f}", "Número de condición\n> 30 = severa", RO)
dato(sl, 8.7, 1.5, f"{corr_e:.3f}", "corr(exper, exper²)", RO)
caja(sl, .7, 4.0, 12, 2.5, [
    f"exper, exper² y age tienen t individuales de {m_full.tvalues['exper']:.2f}, {m_full.tvalues['I(exper ** 2)']:.2f} y "
    f"{m_full.tvalues['age']:.2f}, ninguna significativa…",
    f"…pero la prueba F conjunta es de {ftest.fvalue:.2f} (p = {float(ftest.pvalue):.3f})",
    "Los datos indican que el ciclo de vida importa, pero no permiten separar el efecto de cada variable."], 20)

# 6 diagnósticos formales
sl = slide("Diagnóstico formal", 6)
diag = pd.DataFrame({"Resultado": [f"VIF exper = {vif_full['exper']:.1f}, exper² = {vif_full['I(exper ** 2)']:.1f}", f"{cond:.1f}",
                                   f"χ² = {fg_chi2:.0f}  >  {fg_crit:.1f}  (|R| = {detR:.4f})",
                                   f"8 de 13 regresores con Rⱼ² > R² = {m_full.rsquared:.2f}"],
                     "Conclusión": ["Problema en experiencia", "Severa", "Rechaza ortogonalidad", "Multicolinealidad relevante"]},
                    index=["VIF", "Número de condición", "Farrar-Glauber", "Regla de Klein"])
ptabla(sl, diag, .7, 1.6, 11.9, size=17, index_name="Prueba")

# 7 centrar
sl = slide("Solución 1: centrar la experiencia", 7)
sl.shapes.add_picture(str(FIG / "vif.png"), PI(.4), PI(1.3), height=PI(5.6))
caja(sl, 10.3, 1.6, 2.8, 5, [f"corr: {corr_e:.3f} → {corr_ec:.3f}", f"VIF máx: {vif_full.max():.0f} → {vif_c.max():.1f}",
                              f"R²: {m_full.rsquared:.4f} = {m_c.rsquared:.4f}", "Mismo modelo, solo cambia la interpretación"], 18)

# 8 eliminar / PCA / ridge
sl = slide("Soluciones 2–4: eliminar, combinar, Ridge", 8)
bullets(sl, ["Eliminar age (redundante dados educ y exper)",
             "Promediar meduc y feduc; usar solo IQ como habilidad",
             f"PCA: el CP1 de IQ, KWW, meduc y feduc explica el {100*pca.explained_variance_ratio_[0]:.0f}% de su varianza",
             f"Ridge (λ ≈ {ridge.alpha_:.0f}): encoge primero a exper y exper²"], w=5.6, size=18)
sl.shapes.add_picture(str(FIG / "ridge.png"), PI(6.3), PI(1.4), width=PI(6.8))

# 9 comparación
sl = slide("Comparación de modelos", 9)
t_res2 = resumen[["β educ", "e.e. educ", "R² aj.", "BIC", "VIF máx"]].copy(); t_res2["k"] = resumen["k"].astype(int).astype(str)
ptabla(sl, t_res2, .5, 1.5, 7.4, fmt=lambda x: f"{x:.3f}" if abs(x) < 1 else f"{x:.1f}", size=14, index_name="Modelo")
sl.shapes.add_picture(str(FIG / "educ.png"), PI(8.1), PI(1.5), width=PI(5))
caja(sl, .5, 5.3, 12.3, 1.5, "El β de educ baja por sesgo de variable omitida (habilidad), no por la multicolinealidad. "
     "La multicolinealidad se nota en el error estándar.", 17)

# 10 validación
sl = slide("Validación del modelo final", 10)
ptabla(sl, val, .5, 1.5, 6.6, size=14, index_name="Prueba")
sl.shapes.add_picture(str(FIG / "residuos.png"), PI(7.3), PI(1.5), width=PI(5.8))
caja(sl, .5, 4.2, 6.6, 2.8, ["Usamos errores robustos (HC1): las conclusiones no cambian",
                             "RESET: no hay mala especificación",
                             f"No normalidad, pero n = {n} grande → inferencia asintótica"], 16)

# 11 interpretación
sl = slide("Interpretación (manteniendo lo demás constante)", 11)
dato(sl, .7, 1.5, f"+{e('educ'):.1f}%", "por año de educación")
dato(sl, 4.7, 1.5, f"+{e('married'):.0f}%", "estar casado")
dato(sl, 8.7, 1.5, f"+{e('urban'):.0f}%", "vivir en zona urbana")
dato(sl, .7, 4.0, f"+{e('exper_c'):.1f}%", "por año de experiencia")
dato(sl, 4.7, 4.0, f"+{100*(np.exp(10*m_red.params['IQ'])-1):.1f}%", "por 10 puntos de IQ")
dato(sl, 8.7, 4.0, f"−{abs(e('south')):.1f}%", "vivir en el sur", RO)

# 12 conclusiones
sl = slide("Conclusiones", 12)
bullets(sl, ["Multicolinealidad perfecta → el modelo no está identificado; hay que eliminar una variable",
             "Multicolinealidad imperfecta → MCO sigue siendo MELI, pero las varianzas se inflan",
             "Los cuatro diagnósticos apuntan a exper y exper²",
             "Centrar resuelve la parte no esencial; eliminar, combinar o PCA dan modelos más parsimoniosos",
             f"El retorno a la educación (≈ {100*m_red.params['educ']:.0f}%) es robusto en todas las especificaciones"], size=20)

prs.save(ROOT / "Presentacion_Multicolinealidad_Mincer.pptx")
print("Listo:", ROOT / "Reporte_Multicolinealidad_Mincer.docx", "|", ROOT / "Presentacion_Multicolinealidad_Mincer.pptx")
