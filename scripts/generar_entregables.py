"""Genera el reporte (Word), la presentación (PowerPoint) y la presentación Beamer (.tex)
a partir de resultados.json y figuras/, producidos por scripts/analisis.py.

Uso: .venv/bin/python scripts/analisis.py && .venv/bin/python scripts/generar_entregables.py
"""
import json
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches as PI, Pt as PPt

ROOT = Path(__file__).resolve().parent.parent
FIG = ROOT / "figuras"
R = json.loads((ROOT / "resultados.json").read_text())
AUTORA, CURSO, FECHA = "Ximena Lara", "Fundamentos de Econometría", "Noviembre 2026"

S, F_, A = R["simple"], R["full"], R["anova"]
C, VF, H, D = R["comparacion"], R["vif_full"], R["hetero"], R["dummies"]
PF = R["pruebas_F"]
FG, FIN, AC, ARDL, REZ = R["farrar"], R["finanzas"], R["autocorr"], R["ardl"], R["rezagos"]
JB, HED, NA, Q = R["jb"], R["hedonico"], R["no_anidados"], R["cuantil"]
RED = R["reducido"]
n = R["n"]
pct = lambda b: 100 * (np.exp(b) - 1)
p3 = lambda p: "<0.001" if p < 0.001 else f"{p:.3f}"
NOM = {"I(exper ** 2)": "exper²", "I(exper_c ** 2)": "exper_c²", "Intercept": "constante"}
nm = lambda s: NOM.get(s, s)

# ===================================================================== tablas comunes
T_simple = pd.DataFrame({
    "Valor": [f"{S['xb']:.4f}", f"{S['yb']:.4f}", f"{S['Sxx']:.2f}", f"{S['Sxy']:.3f}", f"{S['Syy']:.3f}",
              f"{S['b1']:.5f}", f"{S['b0']:.4f}", f"{S['se_b1']:.5f}", f"{S['t']:.2f}",
              f"[{S['ic'][0]:.4f}, {S['ic'][1]:.4f}]", f"{S['R2']:.4f} = {S['rho']:.4f}²", f"{S['F']:.2f} = t²"]},
    index=["x̄ (educ)", "ȳ (ln w)", "Sxx", "Sxy", "Syy = SST", "β̂1 = Sxy/Sxx", "β̂0 = ȳ − β̂1x̄",
           "se(β̂1)", "t", "IC 95% β1", "R² = ρ̂²", "F(1, 720)"])

T_pruebas = pd.DataFrame([{"H0": p["H0"], "q": p["q"], "F": f"{p['F']:.3f}", "F crítico 5%": f"{p['Fcrit']:.3f}",
                           "p": p3(p["p"]), "Decisión": p["decision"]} for p in PF]).set_index("H0")

T_vif = pd.DataFrame({k: {"VIF": f"{v['VIF']:.2f}", "Tolerancia": f"{v['Tolerancia']:.3f}",
                          "R² aux": f"{v['R2_aux']:.3f}"} for k, v in VF.items()}).T

T_comp = pd.DataFrame({k: {"β educ": f"{v['beta_educ']:.4f}", "ee educ": f"{v['ee_educ']:.4f}",
                           "R² aj.": f"{v['R2_aj']:.4f}", "BIC": f"{v['BIC']:.1f}", "VIF máx": f"{v['VIF_max']:.1f}",
                           "k": str(int(v["k"]))} for k, v in C.items()}).T

T_formas = pd.DataFrame([{"Forma": f["forma"], "β̂": f"{f['beta']:.4f}", "R²": f"{f['R2']:.4f}",
                          "Lectura": f["lectura"]} for f in R["formas"]]).set_index("Forma")

T_diag = pd.DataFrame([
    ["Breusch–Pagan (wage2)", f"LM = {H['LM']:.2f}", p3(H["bp_p"]), "Rechaza homoscedasticidad"],
    ["White (wage2)", f"LM = {H['white_LM']:.2f}", p3(H["white_p"]), "No rechaza"],
    ["RESET (wage2)", f"F = {R['reset']['F']:.3f}", p3(R["reset"]["p"]), "No rechaza"],
    ["Jarque–Bera (wage2)", f"JB = {JB['JB']:.1f}", p3(JB["p"]), "Rechaza normalidad"],
    ["Durbin–Watson (consumo estático)", f"DW = {AC['dw']:.3f}", "—", "Autocorrelación positiva"],
    ["Breusch–Godfrey(4) (consumo estático)", f"LM = {AC['bg_LM']:.1f}", p3(AC["bg_p"]), "Rechaza no autocorrelación"],
    ["Breusch–Godfrey(4) (ARDL)", f"LM = {ARDL['bg_LM']:.1f}", p3(ARDL["bg_p"]), "Aún rechaza"],
], columns=["Prueba", "Estadístico", "p", "Conclusión 5%"]).set_index("Prueba")

T_red = pd.DataFrame({nm(c): {"β̂": f"{RED['params'][c]:.4f}", "ee HC1": f"{RED['bse_hc1'][c]:.4f}",
                               "p (HC1)": p3(RED["p_hc1"][c]), "Efecto % exacto": f"{RED['exacto'][c]:.2f}"}
                      for c in RED["params"] if c != "Intercept"}).T

# ===================================================================== WORD
def tabla(doc, df_, idx="", size=9):
    t = doc.add_table(rows=1, cols=len(df_.columns) + 1)
    t.style = "Light Grid Accent 1"
    t.rows[0].cells[0].text = idx
    for j, c in enumerate(df_.columns):
        t.rows[0].cells[j + 1].text = str(c)
    for i, row in df_.iterrows():
        cells = t.add_row().cells
        cells[0].text = str(i)
        for j, v in enumerate(row):
            cells[j + 1].text = str(v)
    for r_ in t.rows:
        for c_ in r_.cells:
            for p in c_.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(size)


def pie(doc, txt):
    p = doc.add_paragraph(txt)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].italic = True
    p.runs[0].font.size = Pt(9)


def figura(doc, nombre, txt, w=5.8):
    doc.add_picture(str(FIG / f"{nombre}.png"), width=Inches(w))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    pie(doc, txt)


def ec(doc, txt):
    p = doc.add_paragraph(txt)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].italic = True


def bullets(doc, items, estilo="List Bullet"):
    for b in items:
        doc.add_paragraph(b, style=estilo)


doc = Document()
doc.styles["Normal"].font.name = "Calibri"
doc.styles["Normal"].font.size = Pt(11)
for s_ in doc.sections:
    s_.left_margin = s_.right_margin = Inches(1)
for _ in range(6):
    doc.add_paragraph()
for txt, sz, b in [("Multicolinealidad en la ecuación de salarios de Mincer", 22, True),
                   ("Regresión lineal simple y múltiple, inferencia y diagnóstico con datos reales", 13, False)]:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ = p.add_run(txt)
    r_.bold, r_.font.size = b, Pt(sz)
for _ in range(4):
    doc.add_paragraph()
for line in (AUTORA, CURSO, FECHA):
    doc.add_paragraph(line).alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_page_break()

doc.add_heading("Resumen", 1)
doc.add_paragraph(
    f"Se estima una ecuación de salarios tipo Mincer con {n} trabajadores de la National Longitudinal Survey (base wage2 "
    "de Wooldridge) y se recorren todos los temas del segundo parcial: regresión lineal simple, supuestos, interpretación, "
    "pruebas t y F, variables dummy e interacciones, R² ajustada, precios hedónicos, modelos no anidados, regresión "
    "cuantílica, aplicaciones financieras, diagnóstico de residuos, heteroscedasticidad, autocorrelación, modelos "
    "dinámicos, rezagos y Jarque–Bera. El eje del proyecto es la multicolinealidad: en el modelo con 13 regresores, la "
    f"experiencia y su cuadrado tienen VIF de {VF['exper']['VIF']:.1f} y {VF['exper²']['VIF']:.1f}, la prueba de Farrar–Glauber "
    f"rechaza la ortogonalidad (χ² = {FG['chi2']:.0f}) y la experiencia y la edad no son significativas individualmente "
    f"pero sí en conjunto (F = {PF[1]['F']:.2f}, p = {PF[1]['p']:.3f}). Centrar, combinar regresores, componentes principales "
    f"y Ridge reducen el problema. El retorno a la educación pasa de {100*S['b1']:.1f}% en la regresión simple a "
    f"≈{100*RED['params']['educ']:.1f}% al controlar por habilidad, lo que refleja sesgo por variable omitida y no "
    "multicolinealidad. Para los temas de series de tiempo se usan datos reales de consumo de EE. UU. y de acciones tecnológicas.")

doc.add_heading("1. Introducción y datos", 1)
doc.add_paragraph(
    "La ecuación de Mincer (1974) relaciona el logaritmo del salario con la escolaridad y la experiencia. Al agregar "
    "controles (edad, antigüedad, habilidad, escolaridad de los padres, región) aparecen regresores que se mueven juntos: "
    "multicolinealidad. El objetivo es diagnosticarla, corregirla y aplicar a los mismos datos las herramientas de "
    "inferencia y diagnóstico del curso.")
desc = pd.DataFrame(R["descriptivos"]).T.rename(columns={"mean": "Media", "std": "Desv. est.", "min": "Mín.", "max": "Máx."})
tabla(doc, desc.map(lambda v: f"{v:.2f}"), "Variable")
pie(doc, "Tabla 1. Estadística descriptiva de wage2 (n = 722). wage: salario mensual (USD de 1980).")
tabla(doc, pd.DataFrame({"Fuente": ["Wooldridge (NLS 1980)", "Wooldridge", "Plotly Express", "statsmodels"],
                         "Uso": ["Base principal", "4.6 Precios hedónicos", "4.9 Finanzas", "5.7–5.12 Series de tiempo"]},
                        index=["wage2", "hprice1", "stocks (2018–2019)", "macrodata (1959–2009)"]), "Base")
pie(doc, "Tabla 2. Bases de datos reales utilizadas.")

doc.add_heading("2. Regresión lineal simple", 1)
ec(doc, "ln(wᵢ) = β₀ + β₁·educᵢ + uᵢ,   β̂₁ = Sxy/Sxx,   β̂₀ = ȳ − β̂₁x̄")
tabla(doc, T_simple, "Cantidad")
pie(doc, "Tabla 3. Regresión simple de ln(salario) sobre educación con las fórmulas del curso.")
doc.add_paragraph(
    f"Un año más de educación se asocia con un salario {100*S['b1']:.2f}% mayor (exacto: {S['efecto_exacto']:.2f}%). "
    f"El estadístico t = {S['t']:.2f} supera con mucho el crítico {S['tcrit']:.3f}, y el IC 95% no contiene cero. "
    f"Como hay una sola restricción, F = t² = {S['F']:.2f}. La educación explica sólo el {100*S['R2']:.1f}% de la variación "
    "del log-salario (R² = ρ̂²): es significativa pero deja mucho sin explicar. Se cumplen las ecuaciones normales "
    "(Σû = 0, Σx·û = 0) y los residuos son ortogonales a los valores ajustados.")
figura(doc, "simple_recta", "Figura 1. Recta de MCO y residuales de la regresión simple.")
bullets(doc, [
    f"Superficie SCR(b₀, b₁): el mínimo es el punto MCO. La elipse es muy alargada porque corr(β̂₀, β̂₁) = "
    f"{S['corr_b0_b1']:.3f}: la constante y educ están casi alineadas porque x̄ = {S['xb']:.1f} está lejos de 0 "
    "(una primera forma de 'multicolinealidad' que desaparece al centrar x).",
    f"Reescalamiento: en nivel–nivel β̂₁ = {S['nivel_b1']:.2f} USD/mes; con educación en meses y salario en miles de "
    "USD, la pendiente se multiplica por c/b = (1/1000)/12 y R² no cambia.",
    f"Dummy: en ln w ~ urban, β̂₀ = {S['dummy']['b0']:.4f} = media rural y β̂₁ = {S['dummy']['b1']:.4f} = diferencia de medias "
    "urbano − rural.",
    f"Influencia: al quitar las 10 observaciones con mayor distancia de Cook, β̂₁ pasa de {S['b1']:.4f} a "
    f"{S['b1_sin_influyentes']:.4f}; con n = 722 ningún punto domina la recta.",
    f"Gauss–Markov: con la educación real como X fija, en 5,000 simulaciones MCO tiene varianza "
    f"{R['gauss_markov']['var_mco']:.1e} contra {R['gauss_markov']['var_alt']:.1e} de otro estimador lineal insesgado; ambos "
    "están centrados en β₁, pero MCO es el de menor varianza (MELI)."])
figura(doc, "simple_scr", "Figura 2. Superficie de la suma de cuadrados residual.", 4.3)
figura(doc, "gauss_markov", "Figura 3. Simulación del teorema de Gauss–Markov.", 4.8)

doc.add_heading("3. Supuestos del modelo de regresión lineal múltiple (4.1)", 1)
ec(doc, "ln wᵢ = β₀ + β₁educᵢ + β₂experᵢ + β₃experᵢ² + β₄tenureᵢ + … + εᵢ")
tabla(doc, pd.DataFrame([
    ["Linealidad en parámetros", "Sí (exper² y logs son transformaciones de X)", "RESET p = " + p3(R["reset"]["p"])],
    ["Rango completo", "Falla con pexper = age − educ − 6", f"rango {R['perfecta']['rango']} de {R['perfecta']['cols']}"],
    ["E(ε|X) = 0", "Dudoso sin controlar habilidad", f"β̂ educ cae de {C['Base']['beta_educ']:.3f} a {C['Completo']['beta_educ']:.3f}"],
    ["Homoscedasticidad", "Rechazada al 5% (BP)", "BP p = " + p3(H["bp_p"])],
    ["No autocorrelación", "No aplica en corte transversal aleatorio", f"DW cambia con el orden: {R['dw_wage2'][0]:.2f} vs {R['dw_wage2'][1]:.2f}"],
    ["Normalidad", "Rechazada, pero n grande", "JB p = " + p3(JB["p"])]],
    columns=["Supuesto", "¿Se cumple?", "Evidencia"]).set_index("Supuesto"), "Supuesto")
pie(doc, "Tabla 4. Revisión de los supuestos con los datos del proyecto.")

doc.add_heading("4. Parámetros: estimación e interpretación (4.2)", 1)
tabla(doc, T_formas, "Forma")
pie(doc, "Tabla 5. Las cuatro formas funcionales con educación como regresor.")
coef_full = pd.DataFrame({nm(c): {"β̂": f"{F_['params'][c]:.4f}", "ee": f"{F_['bse'][c]:.4f}", "t": f"{F_['t'][c]:.2f}",
                                  "p": p3(F_["p"][c])} for c in F_["params"]}).T
tabla(doc, coef_full, "Variable")
pie(doc, f"Tabla 6. Modelo completo (n = {n}, R² = {F_['R2']:.3f}, R² aj. = {F_['R2a']:.3f}).")
doc.add_paragraph(
    "Cada coeficiente es un efecto parcial ceteris paribus: compara trabajadores que difieren en una regresora y tienen "
    "iguales las demás. Por eso, cuando dos regresoras casi siempre se mueven juntas (exper y exper², exper y age), los "
    "datos tienen poca información para separar sus efectos y los errores estándar crecen.")

doc.add_heading("5. Prueba F: significancia conjunta y restricciones (4.3)", 1)
doc.add_paragraph(
    f"ANOVA del modelo completo: SST = {A['SST']:.3f} = SSE {A['SSE']:.3f} + SSR {A['SSR']:.3f}; MSR = SSE/k = {A['MSR']:.4f}, "
    f"MSE = SSR/(n−k−1) = {A['MSE']:.4f}, F = MSR/MSE = {A['F']:.2f} > F(0.05; {A['k']}, {A['dfr']}) = {A['Fcrit']:.3f}. "
    "El modelo es globalmente significativo.")
tabla(doc, T_pruebas, "H₀")
pie(doc, f"Tabla 7. Pruebas F con F = [(SSR_R − SSR_NR)/q] / [SSR_NR/(n−k−1)], gl del denominador = {A['dfr']}.")
doc.add_paragraph(
    f"Lectura clave: age no es significativa sola (F = t² = {PF[3]['F']:.2f}, p = {PF[3]['p']:.2f}), pero exper, exper² y "
    f"age juntas sí (F = {PF[1]['F']:.2f}, p = {PF[1]['p']:.3f}). Es la lección del ejercicio 1 de la guía de pruebas t y F: "
    "con regresoras correlacionadas el bloque explica, aunque los datos no permitan separar cada efecto. La restricción "
    "lineal β_meduc = β_feduc (q = 1) no se rechaza, lo que justifica promediar la escolaridad de los padres.")
figura(doc, "regiones_rechazo", "Figura 4. Regiones de rechazo de la prueba t (age) y de la prueba F (bloque de ciclo de vida).")

doc.add_heading("6. Variables cualitativas: dummies e interacciones (4.4)", 1)
ec(doc, "ln w = β₀ + β₁educ + … + δ·black + γ·(educ × black) + ε")
doc.add_paragraph(
    f"Categorías base: no casado, no negro, fuera del sur, rural. Estar casado se asocia con {D['married_exacto']:.1f}% más "
    f"salario y vivir en zona urbana con {D['urban_exacto']:.1f}% más (efecto exacto 100(e^β − 1)). Con la interacción, "
    f"la pendiente de educación es {D['educ']:.4f} para el grupo base y β₁ + γ = {D['pend_black']:.4f} para trabajadores "
    f"negros (γ con p = {D['p_int']:.2f}, no significativa). El coeficiente de black por sí solo ({D['black']:.3f}) sólo es la "
    f"diferencia cuando educ = 0; con 12 años de educación la diferencia es δ + 12γ = {D['dif_black_12']:.3f}.")
doc.add_paragraph(
    f"Trampa de las dummies: con constante, urban y rural = 1 − urban, la matriz X tiene rango {R['trampa_rango'][0]} de "
    f"{R['trampa_rango'][1]} columnas: colinealidad perfecta. Con m categorías se incluyen m − 1 dummies.")

doc.add_heading("7. R² y R² ajustada (4.5)", 1)
tabla(doc, pd.DataFrame([{"Modelo": r_["modelo"], "k": r_["k"], "R²": f"{r_['R2']:.4f}", "R² aj.": f"{r_['R2_aj']:.4f}"}
                         for r_ in R["r2"]]).set_index("Modelo"), "Modelo")
pie(doc, "Tabla 8. Al agregar sibs el R² sube pero el R² ajustado baja; hours sí mejora ambos.")

doc.add_heading("8. Precios hedónicos (4.6)", 1)
ec(doc, "ln Pᵢ = β₀ + β₁ ln(lotsize) + β₂ ln(sqrft) + β₃ bdrms + β₄ colonial + uᵢ")
hp_ = HED["params"]
doc.add_paragraph(
    f"Con 88 casas (hprice1), la elasticidad precio–superficie es {hp_['lsqrft']:.3f} y la del terreno {hp_['llotsize']:.3f} "
    f"(ambas p < 0.001). Una recámara adicional, manteniendo la superficie, se asocia con {HED['bdrms_exacto']:.1f}% más precio "
    f"pero no es significativa (p = {HED['p']['bdrms']:.2f}): en la misma superficie, más recámaras no agregan valor. "
    f"R² = {HED['R2']:.3f}. Al agregar el avalúo (lassess), construido con los mismos atributos, el VIF de lsqrft sube de "
    f"{HED['vif']['lsqrft']:.2f} a {HED['vif_assess']['lsqrft']:.2f} y su error estándar de {HED['se_lsqrft']:.3f} a "
    f"{HED['se_lsqrft_assess']:.3f}.")

doc.add_heading("9. Modelos no anidados (4.7)", 1)
doc.add_paragraph(
    "Modelo A agrega habilidad (IQ, KWW); modelo B agrega entorno familiar (meduc, feduc). Ninguno es restricción del otro, "
    "así que se comparan con criterios de información y con la prueba J de Davidson–MacKinnon.")
tabla(doc, pd.DataFrame({"A (habilidad)": [f"{NA['aicA']:.1f}", f"{NA['bicA']:.1f}", f"{NA['R2aA']:.4f}",
                                           f"ŷ_B en A: t = {NA['tJ_A']:.2f}, p = {NA['pJ_A']:.3f}"],
                         "B (familia)": [f"{NA['aicB']:.1f}", f"{NA['bicB']:.1f}", f"{NA['R2aB']:.4f}",
                                         f"ŷ_A en B: t = {NA['tJ_B']:.2f}, p = {p3(NA['pJ_B'])}"]},
                        index=["AIC", "BIC", "R² aj.", "Prueba J"]), "Criterio")
pie(doc, "Tabla 9. Comparación de modelos no anidados.")
doc.add_paragraph("AIC, BIC y R² ajustada prefieren A, pero la prueba J rechaza en ambas direcciones: ninguno abarca al otro, "
                  "lo que justifica combinarlos… y por eso aparece la multicolinealidad del modelo completo.")

doc.add_heading("10. Regresión cuantílica (4.8)", 1)
tabla(doc, pd.DataFrame([{"τ": q["tau"], "β̂ educ": f"{q['beta_educ']:.4f}", "ee": f"{q['ee']:.4f}",
                          "IC 95%": f"[{q['inf']:.3f}, {q['sup']:.3f}]"} for q in Q["tabla"]]).set_index("τ"), "τ")
pie(doc, f"Tabla 10. Retorno a la educación por cuantil condicional del log-salario (MCO = {Q['ols']:.4f}).")
figura(doc, "cuantilica", "Figura 5. Coeficiente de educación por cuantil.", 4.8)
doc.add_paragraph("El retorno va de ≈4.3% en τ = 0.10 a ≈6.0% en τ = 0.90, pero los intervalos se traslapan: la "
                  "heterogeneidad es moderada y la pendiente media resume razonablemente la relación.")

doc.add_heading("11. Aplicaciones financieras (4.9)", 1)
doc.add_paragraph(
    f"Con rendimientos logarítmicos semanales de 2018–2019 (n = {FIN['n']}), la beta de AAPL respecto al índice tecnológico "
    f"equiponderado es {FIN['beta']:.3f} (IC 95% [{FIN['ic_beta'][0]:.3f}, {FIN['ic_beta'][1]:.3f}]); el alfa "
    f"({FIN['alpha']:.4f}) no es significativo (p = {FIN['p_alpha']:.2f}). Diagnóstico: BP p = {FIN['bp_p']:.3f}, BG(4) "
    f"p = {FIN['bg_p']:.3f}, JB p = {p3(FIN['jb_p'])}: colas pesadas típicas de los rendimientos. Al regresar AAPL sobre MSFT, "
    f"AMZN y GOOG, la F global es {FIN['fac_F']:.2f} (p < 0.001) pero sólo MSFT es significativa: las acciones están "
    f"correlacionadas entre sí (VIF máximo {max(FIN['vif'].values()):.2f}).")
figura(doc, "capm", "Figura 6. Beta de mercado de AAPL.", 4.5)

doc.add_heading("12. Multicolinealidad: diagnóstico y corrección", 1)
P = R["perfecta"]
doc.add_paragraph(
    f"Perfecta: con constante, educ, age y pexper = age − educ − 6, la matriz X tiene rango {P['rango']} de {P['cols']} columnas "
    f"y el número de condición de X'X es {P['cond']:.1e}: (X'X)⁻¹ no existe y los coeficientes no están identificados.")
figura(doc, "correlaciones", "Figura 7. Correlaciones entre regresoras.", 4.4)
tabla(doc, T_vif, "Variable")
pie(doc, f"Tabla 11. VIF, tolerancia y R² auxiliar del modelo completo (R² del modelo = {F_['R2']:.3f}; regla de Klein).")
eig = R["autovalores"]
doc.add_paragraph(
    f"Autovalores de la matriz de correlaciones de las regresoras estandarizadas: el mayor es {eig[0]:.3f} y el menor "
    f"{eig[-1]:.4f}, de modo que el índice de condición máximo es √({eig[0]:.3f}/{eig[-1]:.4f}) = "
    f"{R['indices_condicion'][-1]:.2f} (número de condición de la matriz estandarizada). Con columnas normalizadas y "
    f"constante (Belsley–Kuh–Welsch) el número de condición es {R['num_cond_bkw']:.1f} > 30. Farrar–Glauber: |R| = "
    f"{FG['detR']:.4f}, χ² = {FG['chi2']:.1f} > {FG['crit']:.1f} (gl = {FG['gl']:.0f}), se rechaza la ortogonalidad.")
doc.add_paragraph(
    f"Síntomas: el error estándar de educ sube de {R['base']['bse']['educ']:.4f} (modelo base) a {F_['bse']['educ']:.4f} "
    f"(completo); exper (t = {F_['t']['exper']:.2f}), exper² (t = {F_['t']['I(exper ** 2)']:.2f}) y age "
    f"(t = {F_['t']['age']:.2f}) no son significativas individualmente pero sí en conjunto.")
doc.add_paragraph(
    f"Soluciones: (a) centrar exper reduce corr(exper, exper²) de {R['corr_exper_exper2']:.3f} a {R['corr_experc']:.3f} y el "
    f"VIF máximo de {C['Completo']['VIF_max']:.1f} a {R['vif_centrado_max']:.2f} sin cambiar R²; (b) eliminar age y promediar "
    f"meduc y feduc (validado por la prueba F de β_meduc = β_feduc); (c) el primer componente principal de IQ, KWW, meduc y feduc "
    f"explica el {100*R['pca']['var'][0]:.1f}% de su varianza; (d) Ridge con λ ≈ {R['ridge']['alpha']:.0f} encoge primero "
    "a exper y exper².")
figura(doc, "vif", "Figura 8. VIF antes y después de centrar la experiencia.")
tabla(doc, T_comp, "Modelo")
pie(doc, "Tabla 12. Comparación de especificaciones.")
figura(doc, "educ", "Figura 9. Retorno a la educación por especificación (IC 95%).", 4.8)
figura(doc, "ridge", "Figura 10. Trayectorias de Ridge.", 5)

doc.add_heading("13. MANOVA", 1)
mw, mp_ = R["manova"]["Wilks' lambda"], R["manova"]["Pillai's trace"]
doc.add_paragraph(
    f"¿Difieren conjuntamente educ, IQ y KWW entre el sur y el resto del país? Lambda de Wilks = {mw['valor']:.4f}, traza "
    f"de Pillai = {mp_['valor']:.4f}, F({mw['num']:.0f}, {mw['den']:.0f}) = {mw['F']:.2f}, p = {p3(mw['p'])}: sí difieren. "
    "MANOVA estudia varias variables dependientes a la vez; no diagnostica multicolinealidad, que es una propiedad de las "
    "columnas de X.")

doc.add_heading("14. Distribución de los residuos y diagnóstico (5.1–5.2)", 1)
figura(doc, "residuos", "Figura 11. Residuos vs ajustados, Q–Q, histograma y leverage–Cook (modelo reducido).")
doc.add_paragraph(f"No hay curvatura ni un abanico marcado. La cola izquierda es más pesada que la normal. Hay "
                  f"{R['n_cook_4n']} observaciones con distancia de Cook > 4/n, pero la mayor vale sólo "
                  f"{R['cook_max_red']:.3f}: ninguna domina la estimación.")

doc.add_heading("15. Heteroscedasticidad y Breusch–Pagan (5.3–5.6)", 1)
doc.add_paragraph(
    f"Regresión auxiliar de ê² sobre las {H['q']} regresoras: R²aux = {H['R2aux']:.4f}, LM = n·R²aux = {n}×{H['R2aux']:.4f} = "
    f"{H['LM']:.2f} > χ²({H['q']}; 0.05) = {H['crit']:.2f} (p = {H['bp_p']:.3f}): se rechaza la homoscedasticidad. White "
    f"no la rechaza (p = {H['white_p']:.2f}). Los errores HC0–HC3 casi no cambian y WLS/FGLS da β educ = {H['wls_educ']:.4f} "
    f"(ee {H['wls_se']:.4f}): la heteroscedasticidad es leve. Por prudencia, el modelo final se reporta con errores HC1.")
tabla(doc, T_red, "Variable")
pie(doc, f"Tabla 13. Modelo final (reducido) con errores robustos HC1; R² = {RED['R2']:.3f}.")

doc.add_heading("16. Autocorrelación: DW, BG y HAC (5.7–5.10)", 1)
doc.add_paragraph(
    f"En wage2 (corte transversal) el orden de las filas es arbitrario: DW vale {R['dw_wage2'][0]:.3f} en el orden original y "
    f"{R['dw_wage2'][1]:.3f} si se reordenan, así que la prueba no tiene sentido. Con consumo e ingreso trimestrales de EE. UU. "
    f"(macrodata, n = {AC['n']}), el modelo estático ln c = α + β ln y tiene R² = {AC['R2']:.4f} pero DW = {AC['dw']:.3f} ≈ "
    f"2(1 − ρ̂) con ρ̂ = {AC['rho']:.3f}, y BG(4) = {AC['bg_LM']:.1f} (p ≈ 0): autocorrelación positiva severa. El error "
    f"estándar HAC de β ({AC['se_hac']:.4f}) es más del doble que el clásico ({AC['se']:.4f}).")
figura(doc, "autocorrelacion", "Figura 12. Residuos y ACF del modelo estático de consumo.")

doc.add_heading("17. Modelos dinámicos y rezagos (5.11–5.12)", 1)
pa = ARDL["params"]
ec(doc, "ln cₜ = α + φ ln cₜ₋₁ + β₀ ln yₜ + β₁ ln yₜ₋₁ + uₜ")
doc.add_paragraph(
    f"φ̂ = {pa['lc_1']:.4f} (|φ| < 1, estable), β̂₀ = {pa['ly']:.4f}, β̂₁ = {pa['ly_1']:.4f}; el multiplicador de largo plazo "
    f"(β₀ + β₁)/(1 − φ) = {ARDL['lr']:.3f} ≈ 1: en el largo plazo el consumo crece en la misma proporción que el ingreso. "
    f"Multicolinealidad entre rezagos: corr(ln yₜ, ln yₜ₋₁) = {ARDL['corr_y_y1']:.4f} y los VIF superan 3,000. BG(4) en el "
    f"ARDL todavía rechaza (p = {ARDL['bg_p']:.1e}), por lo que hace falta más dinámica o inferencia HAC. DW no se usa porque "
    "hay dependiente rezagada.")
tabla(doc, pd.DataFrame([{"Rezagos": r_["rezagos"], "AIC": f"{r_['AIC']:.1f}", "BIC": f"{r_['BIC']:.1f}",
                          "Σβ (efecto acumulado)": f"{r_['suma_betas']:.3f}", "R² aj.": f"{r_['R2_aj']:.3f}"}
                         for r_ in REZ["tabla"]]).set_index("Rezagos"), "Rezagos")
pie(doc, "Tabla 14. Rezagos distribuidos en tasas de crecimiento: Δln cₜ sobre Δln yₜ₋ⱼ.")
doc.add_paragraph(f"BIC elige {REZ['p_bic']} rezago y AIC {REZ['p_aic']}: BIC penaliza más y favorece la parsimonia. Con "
                  f"{REZ['p_bic']} rezago, el efecto acumulado de un aumento de 1 punto en el crecimiento del ingreso es "
                  f"{REZ['acumulado']:.3f} puntos de crecimiento del consumo.")

doc.add_heading("18. Normalidad: Jarque–Bera (5.13)", 1)
doc.add_paragraph(
    f"JB = n/6·(S² + (K − 3)²/4) = {n}/6·({JB['S']:.3f}² + ({JB['K']:.3f} − 3)²/4) = {JB['JB']:.1f} > χ²(2; 0.05) = 5.99: se "
    "rechaza la normalidad (colas pesadas, ligera asimetría negativa). Con n = 722 la inferencia es válida asintóticamente; "
    "rechazar JB no implica sesgo de MCO.")
tabla(doc, T_diag, "Prueba")
pie(doc, "Tabla 15. Resumen de diagnósticos.")

doc.add_heading("19. Inferencia frente a predicción", 1)
tabla(doc, pd.DataFrame([{"Modelo": p["modelo"], "R² prueba": f"{p['R2_test']:.4f}", "RMSE prueba": f"{p['RMSE_test']:.4f}"}
                         for p in R["prediccion"]]).set_index("Modelo"), "Modelo")
pie(doc, "Tabla 16. Desempeño fuera de muestra (80% entrenamiento, 20% prueba).")
doc.add_paragraph("El modelo completo y el reducido predicen casi igual: la multicolinealidad daña la interpretación de "
                  "coeficientes individuales, no la predicción. Ridge mejora ligeramente el RMSE a cambio de sesgo.")

doc.add_heading("20. Conclusiones", 1)
bullets(doc, [
    f"El retorno a la educación es positivo y robusto: {100*S['b1']:.1f}% en la regresión simple y ≈4–5% al controlar habilidad y "
    "entorno familiar. La caída se debe a sesgo por variable omitida, no a multicolinealidad.",
    "La multicolinealidad perfecta impide estimar; la imperfecta no sesga MCO (sigue siendo MELI) pero infla las varianzas.",
    "VIF, tolerancia, Klein, autovalores, número de condición y Farrar–Glauber coinciden: el problema está en exper, exper² y age.",
    "La prueba F conjunta detecta lo que las pruebas t individuales no pueden: el ciclo de vida sí importa.",
    "Centrar, eliminar o combinar con justificación teórica, usar componentes principales o Ridge reducen el problema sin perder ajuste.",
    "Diagnóstico: heteroscedasticidad leve (se reporta con HC1), no normalidad irrelevante con n grande y sin problemas de forma "
    "funcional. En series de tiempo, la autocorrelación del modelo estático exige dinámica o HAC.",
    "La misma lección aparece en precios hedónicos, finanzas y rezagos distribuidos."])
doc.add_heading("Referencias", 1)
for ref in ["Belsley, D., Kuh, E. y Welsch, R. (1980). Regression Diagnostics. Wiley.",
            "Blackburn, M. y Neumark, D. (1992). Quarterly Journal of Economics, 107(4), 1421–1436.",
            "Davidson, R. y MacKinnon, J. (1981). Several tests for model specification. Econometrica, 49(3), 781–793.",
            "Farrar, D. y Glauber, R. (1967). Review of Economics and Statistics, 49(1), 92–107.",
            "Galindo, C. (2026). Notas de regresión lineal múltiple, pruebas t y F, inferencia y diagnóstico. Universidad Panamericana.",
            "Gujarati, D. y Porter, D. (2010). Econometría (5.ª ed.). McGraw-Hill.",
            "Koenker, R. y Bassett, G. (1978). Regression quantiles. Econometrica, 46(1), 33–50.",
            "Mincer, J. (1974). Schooling, Experience, and Earnings. NBER.",
            "Wooldridge, J. M. (2020). Introducción a la econometría (7.ª ed.). Cengage."]:
    doc.add_paragraph(ref)
doc.save(ROOT / "Reporte_Multicolinealidad_Mincer.docx")

# ===================================================================== POWERPOINT
OSC, AC_, TXT, BL, RO = PRGB(0x1D, 0x24, 0x33), PRGB(0x2F, 0x6F, 0x9F), PRGB(0x33, 0x33, 0x33), PRGB(255, 255, 255), PRGB(0xD1, 0x60, 0x3D)
prs = Presentation()
prs.slide_width, prs.slide_height = PI(13.333), PI(7.5)
BLANK = prs.slide_layouts[6]
num = [1]


def caja(sl, x, y, w, h, texto, size=20, color=TXT, bold=False):
    tf = sl.shapes.add_textbox(PI(x), PI(y), PI(w), PI(h)).text_frame
    tf.word_wrap = True
    for i, ln in enumerate(texto if isinstance(texto, list) else [texto]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text, p.font.size, p.font.color.rgb, p.font.bold = ln, PPt(size), color, bold
        p.space_after = PPt(7)


def slide(titulo, seccion=""):
    sl = prs.slides.add_slide(BLANK)
    bar = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, PI(1.05))
    bar.fill.solid()
    bar.fill.fore_color.rgb = OSC
    bar.line.fill.background()
    caja(sl, .55, .2, 10.5, .8, titulo, 28, BL, True)
    if seccion:
        caja(sl, 10.6, .33, 2.5, .5, seccion, 13, PRGB(0xBF, 0xD4, 0xEE))
    num[0] += 1
    caja(sl, 12.4, 6.95, .8, .4, str(num[0]), 11, PRGB(0x99, 0x99, 0x99))
    return sl


def viñetas(sl, items, x=.6, y=1.35, w=12.1, h=5.6, size=19):
    caja(sl, x, y, w, h, ["•  " + i for i in items], size)


def img(sl, nombre, x, y, w=None, h=None):
    sl.shapes.add_picture(str(FIG / f"{nombre}.png"), PI(x), PI(y), width=PI(w) if w else None, height=PI(h) if h else None)


def ptabla(sl, df_, x, y, w, size=13, idx=""):
    rows, cols = len(df_) + 1, len(df_.columns) + 1
    tb = sl.shapes.add_table(rows, cols, PI(x), PI(y), PI(w), PI(.37 * rows)).table
    tb.cell(0, 0).text = idx
    for j, c in enumerate(df_.columns):
        tb.cell(0, j + 1).text = str(c)
    for i, (ix, row) in enumerate(df_.iterrows(), 1):
        tb.cell(i, 0).text = str(ix)
        for j, v in enumerate(row):
            tb.cell(i, j + 1).text = str(v)
    for r_ in range(rows):
        for c_ in range(cols):
            for p in tb.cell(r_, c_).text_frame.paragraphs:
                p.font.size = PPt(size)


def dato(sl, x, y, valor, etiqueta, color=AC_):
    caja(sl, x, y, 3.9, .9, valor, 40, color, True)
    caja(sl, x, y + .95, 3.9, 1, etiqueta, 14)


sl = prs.slides.add_slide(BLANK)
bg = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
bg.fill.solid()
bg.fill.fore_color.rgb = OSC
bg.line.fill.background()
caja(sl, .9, 2.0, 11.5, 1.6, "Multicolinealidad en la ecuación de salarios de Mincer", 42, BL, True)
caja(sl, .9, 3.75, 11.5, .8, "Regresión simple y múltiple, inferencia y diagnóstico con datos reales", 21, PRGB(0xBF, 0xD4, 0xEE))
caja(sl, .9, 5.3, 11.5, 1.2, [AUTORA, f"{CURSO}  ·  {FECHA}"], 18, BL)

sl = slide("Pregunta y datos")
viñetas(sl, ["¿Cuánto aumenta el salario un año más de educación?",
             "Controles (experiencia, edad, habilidad, padres) muy correlacionados → multicolinealidad",
             f"wage2 (Wooldridge): {n} hombres, NLS 1980 — base principal",
             "hprice1 (hedónicos), acciones tecnológicas (finanzas) y macrodata (series de tiempo)",
             "Recorremos todos los temas del parcial: 4.1–4.9 y 5.1–5.13"], size=20)
caja(sl, .6, 5.6, 12, .8, "ln wᵢ = β₀ + β₁ educᵢ + β₂ experᵢ + β₃ experᵢ² + … + εᵢ", 26, AC_, True)

sl = slide("Regresión lineal simple", "MCO")
ptabla(sl, T_simple, .5, 1.35, 5.6, 12, "Cantidad")
img(sl, "simple_recta", 6.3, 1.4, w=6.8)
caja(sl, 6.3, 4.4, 6.8, 2.6, [f"+1 año de educación → +{100*S['b1']:.1f}% salario (exacto {S['efecto_exacto']:.1f}%)",
                              f"R² = ρ̂² = {S['R2']:.3f}: significativa pero explica poco",
                              f"Dummy urban: β̂₁ = diferencia de medias = {S['dummy']['b1']:.3f}",
                              f"Gauss–Markov: var MCO {R['gauss_markov']['var_mco']:.1e} < alternativo {R['gauss_markov']['var_alt']:.1e}"], 15)

sl = slide("Supuestos del RLM", "4.1")
ptabla(sl, pd.DataFrame([["Linealidad en parámetros", "Sí", f"RESET p = {R['reset']['p']:.2f}"],
                         ["Rango completo", "Falla con age − educ − 6", f"rango {P['rango']} de {P['cols']}"],
                         ["E(ε|X) = 0", "Dudoso sin habilidad", f"β educ {C['Base']['beta_educ']:.3f} → {C['Completo']['beta_educ']:.3f}"],
                         ["Homoscedasticidad", "Rechazada (leve)", f"BP p = {H['bp_p']:.3f}"],
                         ["No autocorrelación", "No aplica (corte transversal)", "DW depende del orden"],
                         ["Normalidad", "Rechazada; n grande", f"JB = {JB['JB']:.0f}"]],
                        columns=["¿Se cumple?", "Evidencia", "x"]).set_index("¿Se cumple?").rename(columns={"Evidencia": "¿Se cumple?", "x": "Evidencia"}),
       .6, 1.5, 12.1, 16, "Supuesto")

sl = slide("Parámetros e interpretación", "4.2")
ptabla(sl, T_formas, .6, 1.45, 12.1, 15, "Forma")
caja(sl, .6, 4.0, 12, 2.5, ["Cada β es un efecto parcial ceteris paribus.",
                            "Si dos regresoras casi siempre se mueven juntas, los datos no pueden separar sus efectos."], 19)

sl = slide("Prueba F: significancia conjunta", "4.3")
ptabla(sl, T_pruebas, .4, 1.35, 12.5, 12, "H₀")
img(sl, "regiones_rechazo", 2.0, 3.95, w=9.3)

sl = slide("Dummies e interacciones", "4.4")
viñetas(sl, [f"Casado: +{D['married_exacto']:.1f}%   ·   Urbano: +{D['urban_exacto']:.1f}%   ·   Sur: {pct(D['south']):.1f}%",
             f"Pendiente de educación: grupo base {D['educ']:.4f}; negros β₁ + γ = {D['pend_black']:.4f} (γ: p = {D['p_int']:.2f})",
             f"δ de black sola = diferencia sólo cuando educ = 0; con 12 años: δ + 12γ = {D['dif_black_12']:.3f}",
             f"Trampa de las dummies: constante + urban + rural → rango {R['trampa_rango'][0]} de {R['trampa_rango'][1]}"], size=19)

sl = slide("R² ajustada, hedónicos y no anidados", "4.5–4.7")
ptabla(sl, pd.DataFrame([{"Modelo": r_["modelo"], "R²": f"{r_['R2']:.4f}", "R² aj.": f"{r_['R2_aj']:.4f}"} for r_ in R["r2"]]).set_index("Modelo"),
       .5, 1.4, 5.8, 13, "Modelo")
caja(sl, 6.6, 1.35, 6.3, 5.5, [
    f"Hedónico (88 casas): elasticidad precio–superficie {hp_['lsqrft']:.2f}; recámaras no significativas (p = {HED['p']['bdrms']:.2f})",
    f"Con el avalúo: VIF de lsqrft {HED['vif']['lsqrft']:.1f} → {HED['vif_assess']['lsqrft']:.1f}",
    f"No anidados: BIC A = {NA['bicA']:.0f} < B = {NA['bicB']:.0f}",
    f"Prueba J rechaza en ambas direcciones (p = {NA['pJ_A']:.3f} y {p3(NA['pJ_B'])}) → combinar"], 16)

sl = slide("Regresión cuantílica", "4.8")
img(sl, "cuantilica", .4, 1.4, w=8.2)
caja(sl, 8.9, 1.6, 4.1, 5, [f"τ = 0.10: {100*Q['tabla'][0]['beta_educ']:.1f}%", f"τ = 0.50: {100*Q['tabla'][2]['beta_educ']:.1f}%",
                            f"τ = 0.90: {100*Q['tabla'][4]['beta_educ']:.1f}%", f"MCO: {100*Q['ols']:.1f}%",
                            "IC se traslapan → heterogeneidad moderada"], 17)

sl = slide("Aplicaciones financieras", "4.9")
img(sl, "capm", .4, 1.4, w=7.9)
caja(sl, 8.6, 1.5, 4.5, 5.5, [f"β AAPL = {FIN['beta']:.3f}", f"α = {FIN['alpha']:.4f} (p = {FIN['p_alpha']:.2f})",
                              f"BP p = {FIN['bp_p']:.3f}, BG(4) p = {FIN['bg_p']:.2f}", f"JB p = {p3(FIN['jb_p'])}: colas pesadas",
                              f"AAPL ~ MSFT + AMZN + GOOG: F = {FIN['fac_F']:.1f}, sólo MSFT significativa"], 16)

sl = slide("Multicolinealidad: síntomas", "Diagnóstico")
dato(sl, .6, 1.4, f"{VF['exper²']['VIF']:.0f}", "VIF máximo (exper²)", RO)
dato(sl, 4.6, 1.4, f"{R['num_cond_bkw']:.0f}", "número de condición (BKW) > 30", RO)
dato(sl, 8.6, 1.4, f"{FG['chi2']:.0f}", f"χ² Farrar–Glauber (crítico {FG['crit']:.0f})", RO)
caja(sl, .6, 3.9, 12, 3, [f"Perfecta: age, educ y age − educ − 6 → rango {P['rango']} de {P['cols']}",
                          f"t individuales: exper {F_['t']['exper']:.2f}, exper² {F_['t']['I(exper ** 2)']:.2f}, age {F_['t']['age']:.2f}",
                          f"…pero F conjunta = {PF[1]['F']:.2f} (p = {PF[1]['p']:.3f})",
                          f"Autovalor mínimo {R['autovalores'][-1]:.3f} → índice de condición {R['indices_condicion'][-1]:.1f}"], 19)

sl = slide("VIF, tolerancia y correlaciones", "Diagnóstico")
ptabla(sl, T_vif, .4, 1.3, 5.6, 11, "Variable")
img(sl, "correlaciones", 6.4, 1.3, h=5.8)

sl = slide("Soluciones", "Corrección")
img(sl, "vif", .3, 1.3, w=6.4)
img(sl, "ridge", 6.8, 1.3, w=6.3)
caja(sl, .5, 5.0, 12.3, 2, [f"Centrar: VIF máx {C['Completo']['VIF_max']:.0f} → {R['vif_centrado_max']:.1f}, mismo R²  ·  PCA: CP1 = {100*R['pca']['var'][0]:.0f}% de varianza",
                            f"Eliminar age + promediar padres (F de β_meduc = β_feduc: p = {PF[4]['p']:.2f})  ·  Ridge λ ≈ {R['ridge']['alpha']:.0f}"], 16)

sl = slide("Comparación de modelos", "Corrección")
ptabla(sl, T_comp, .4, 1.4, 7.6, 13, "Modelo")
img(sl, "educ", 8.2, 1.4, w=4.9)
caja(sl, .4, 4.3, 12.4, 2, "La caída del β de educ se debe a sesgo por variable omitida (habilidad); la multicolinealidad se ve en el error estándar.", 16)

sl = slide("Residuos y MANOVA", "5.1–5.2")
img(sl, "residuos", .4, 1.25, h=5.9)
caja(sl, 8.7, 1.5, 4.4, 5.5, [f"Cook máx = {R['cook_max_red']:.3f}: sin observaciones dominantes",
                              f"MANOVA (educ, IQ, KWW) ~ sur: Wilks = {mw['valor']:.3f}, p = {p3(mw['p'])}",
                              "MANOVA ≠ prueba de multicolinealidad"], 16)

sl = slide("Heteroscedasticidad y Jarque–Bera", "5.3–5.6, 5.13")
viñetas(sl, [f"BP: R²aux = {H['R2aux']:.4f} → LM = n·R² = {H['LM']:.2f} > χ²({H['q']}) = {H['crit']:.2f} (p = {H['bp_p']:.3f})",
             f"White: p = {H['white_p']:.2f} (pierde potencia)",
             f"HC0–HC3 casi iguales a los clásicos; WLS: β educ = {H['wls_educ']:.4f}",
             f"JB = {n}/6·(S² + (K−3)²/4) = {JB['JB']:.0f} > 5.99 → no normalidad (K = {JB['K']:.2f})",
             f"RESET p = {R['reset']['p']:.2f}: forma funcional adecuada"], size=19)

sl = slide("Autocorrelación", "5.7–5.10")
img(sl, "autocorrelacion", .4, 1.3, w=12.5)
caja(sl, .5, 5.6, 12.3, 1.5, [f"wage2: DW {R['dw_wage2'][0]:.2f} vs {R['dw_wage2'][1]:.2f} al reordenar → no aplica en corte transversal",
                              f"Consumo estático: DW = {AC['dw']:.3f} ≈ 2(1−ρ̂), BG(4) = {AC['bg_LM']:.0f}; ee HAC {AC['se_hac']:.4f} vs clásico {AC['se']:.4f}"], 15)

sl = slide("Modelos dinámicos y rezagos", "5.11–5.12")
viñetas(sl, [f"ARDL: ln cₜ = α + φ ln cₜ₋₁ + β₀ ln yₜ + β₁ ln yₜ₋₁:  φ̂ = {pa['lc_1']:.3f} (estable)",
             f"Multiplicador de largo plazo (β₀ + β₁)/(1 − φ) = {ARDL['lr']:.3f} ≈ 1",
             f"Rezagos casi idénticos: corr(ln yₜ, ln yₜ₋₁) = {ARDL['corr_y_y1']:.4f}, VIF > 3,000",
             f"BG(4) en el ARDL aún rechaza → más dinámica o HAC",
             f"Rezagos distribuidos (crecimientos): BIC elige {REZ['p_bic']}, AIC elige {REZ['p_aic']}; efecto acumulado {REZ['acumulado']:.2f}"], size=19)

sl = slide("Inferencia vs predicción", "scikit-learn")
ptabla(sl, pd.DataFrame([{"Modelo": p["modelo"], "R² prueba": f"{p['R2_test']:.3f}", "RMSE prueba": f"{p['RMSE_test']:.3f}"} for p in R["prediccion"]]).set_index("Modelo"),
       .6, 1.5, 7, 16, "Modelo")
caja(sl, 8, 1.5, 5, 4, ["Completo ≈ reducido fuera de muestra", "La multicolinealidad afecta la interpretación, no la predicción",
                        "Ridge: menor RMSE a cambio de sesgo"], 17)

sl = slide("Conclusiones")
viñetas(sl, [f"Retorno a la educación robusto: {100*S['b1']:.1f}% (simple) → ≈{100*RED['params']['educ']:.1f}% (controlando habilidad)",
             "Perfecta: no hay MCO único · Imperfecta: MCO sigue siendo MELI, pero las varianzas se inflan",
             "Todos los diagnósticos apuntan a exper, exper² y age; la F conjunta revela lo que las t no",
             "Centrar, combinar, PCA o Ridge corrigen sin perder ajuste",
             "Heteroscedasticidad leve (HC1), no normalidad irrelevante con n grande; en series de tiempo hace falta dinámica o HAC",
             "La misma lección aparece en hedónicos, finanzas y rezagos"], size=19)
prs.save(ROOT / "Presentacion_Multicolinealidad_Mincer.pptx")

# ===================================================================== BEAMER
def T(_txt, **kw):
    for k_, v in kw.items():
        _txt = _txt.replace(f"@{k_}@", str(v))
    return _txt


f2, f3, f4 = (lambda x: f"{x:.2f}"), (lambda x: f"{x:.3f}"), (lambda x: f"{x:.4f}")


def pl(p):
    return r"$<0.001$" if p < 0.001 else f"{p:.3f}"


def tex_tabla(cab, filas, alin=None):
    alin = alin or "l" + "r" * (len(cab) - 1)
    out = [r"\begin{tabular}{" + alin + "}", r"\toprule", " & ".join(cab) + r" \\", r"\midrule"]
    out += [" & ".join(str(c) for c in f) + r" \\" for f in filas]
    out += [r"\bottomrule", r"\end{tabular}"]
    return "\n".join(out)


VN = {"exper²": r"exper$^2$", "IQ": "IQ"}
vif_rows = [[VN.get(k_, k_), f2(v["VIF"]), f3(v["Tolerancia"]), f3(v["R2_aux"])] for k_, v in VF.items()]
pf_tex = [r"$\beta_{IQ}=\beta_{KWW}=\beta_{meduc}=\beta_{feduc}=0$", r"$\beta_{exper}=\beta_{exper^2}=\beta_{age}=0$",
          r"$\beta_{south}=\beta_{urban}=0$", r"$\beta_{age}=0$", r"$\beta_{meduc}=\beta_{feduc}$"]
pf_rows = [[h_, p["q"], f3(p["F"]), f3(p["Fcrit"]), pl(p["p"]), "Rechaza" if p["p"] < .05 else "No rechaza"] for h_, p in zip(pf_tex, PF)]
formas_tex = [["Nivel--nivel", r"$w=\beta_0+\beta_1 educ$", f2(R["formas"][0]["beta"]), rf"+1 año $\to$ +{R['formas'][0]['beta']:.1f} USD/mes"],
              ["Log--nivel", r"$\ln w=\beta_0+\beta_1 educ$", f4(R["formas"][1]["beta"]), rf"+1 año $\to$ +{100*R['formas'][1]['beta']:.1f}\%"],
              ["Nivel--log", r"$w=\beta_0+\beta_1\ln educ$", f2(R["formas"][2]["beta"]), rf"+1\% $\to$ +{R['formas'][2]['beta']/100:.2f} USD"],
              ["Log--log", r"$\ln w=\beta_0+\beta_1\ln educ$", f3(R["formas"][3]["beta"]), rf"elasticidad {R['formas'][3]['beta']:.2f}"]]
comp_rows = [[k_, f4(v["beta_educ"]), f4(v["ee_educ"]), f4(v["R2_aj"]), f"{v['BIC']:.1f}", f"{v['VIF_max']:.1f}", int(v["k"])] for k_, v in C.items()]
red_rows = [[NOM.get(c, c).replace("exper_c²", r"exper$_c^2$").replace("exper_c", r"exper$_c$"), f4(RED["params"][c]), f4(RED["bse_hc1"][c]),
             pl(RED["p_hc1"][c]), f2(RED["exacto"][c]).replace("-", "$-$")] for c in RED["params"] if c != "Intercept"]
q_rows = [[q["tau"], f4(q["beta_educ"]), f4(q["ee"]), f"[{q['inf']:.3f}, {q['sup']:.3f}]"] for q in Q["tabla"]]
diag_rows = [["Breusch--Pagan", rf"LM $=$ {H['LM']:.2f}", pl(H["bp_p"]), "Rechaza homoscedasticidad"],
             ["White", rf"LM $=$ {H['white_LM']:.2f}", pl(H["white_p"]), "No rechaza"],
             ["RESET", rf"F $=$ {R['reset']['F']:.3f}", pl(R["reset"]["p"]), "No rechaza"],
             ["Jarque--Bera", rf"JB $=$ {JB['JB']:.1f}", pl(JB["p"]), "Rechaza normalidad"],
             ["DW (consumo)", rf"DW $=$ {AC['dw']:.3f}", "--", r"$\rho>0$ fuerte"],
             ["BG(4) (consumo)", rf"LM $=$ {AC['bg_LM']:.1f}", pl(AC["bg_p"]), "Autocorrelación"],
             ["BG(4) (ARDL)", rf"LM $=$ {ARDL['bg_LM']:.1f}", pl(ARDL["bg_p"]), "Aún rechaza"]]
rez_rows = [[r_["rezagos"], f"{r_['AIC']:.1f}", f"{r_['BIC']:.1f}", f3(r_["suma_betas"]), f3(r_["R2_aj"])] for r_ in REZ["tabla"]]
pred_rows = [[p["modelo"], f4(p["R2_test"]), f4(p["RMSE_test"])] for p in R["prediccion"]]
r2_rows = [[r_["modelo"], r_["k"], f4(r_["R2"]), f4(r_["R2_aj"])] for r_ in R["r2"]]

B = T(r"""\documentclass[aspectratio=169,10pt]{beamer}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage[spanish,es-nodecimaldot,es-noshorthands]{babel}
\usepackage{amsmath,amssymb}
\usepackage{booktabs}
\usepackage{graphicx}
\graphicspath{{figuras/}}

\usetheme{default}
\definecolor{tinta}{HTML}{1D2433}
\definecolor{azul}{HTML}{2F6F9F}
\definecolor{naranja}{HTML}{D1603D}
\setbeamercolor{frametitle}{bg=tinta,fg=white}
\setbeamercolor{title}{fg=tinta}
\setbeamercolor{structure}{fg=azul}
\setbeamercolor{alerted text}{fg=naranja}
\setbeamertemplate{navigation symbols}{}
\setbeamertemplate{frametitle}{\nointerlineskip\begin{beamercolorbox}[wd=\paperwidth,ht=2.6ex,dp=1.2ex,leftskip=1em]{frametitle}\usebeamerfont{frametitle}\insertframetitle\end{beamercolorbox}}
\setbeamertemplate{footline}{\hfill\scriptsize\color{gray}\insertframenumber/\inserttotalframenumber\hspace{1em}\vspace{0.6em}}
\newcommand{\dato}[2]{\begin{minipage}{0.3\textwidth}\centering{\Huge\color{naranja}\textbf{#1}}\\[0.3em]\small #2\end{minipage}}

\title{Multicolinealidad en la ecuación de salarios de Mincer}
\subtitle{Regresión simple y múltiple, inferencia y diagnóstico con datos reales}
\author{@AUTORA@}
\institute{@CURSO@}
\date{@FECHA@}

\begin{document}

\begin{frame}
\titlepage
\end{frame}

\begin{frame}{Contenido}
\begin{columns}[T]
\column{0.5\textwidth}
\begin{enumerate}
\item Datos y pregunta
\item Regresión lineal simple
\item 4.1 Supuestos
\item 4.2 Parámetros y formas funcionales
\item 4.3 Prueba F y ANOVA
\item 4.4 Dummies e interacciones
\item 4.5 $R^2$ y $\bar R^2$
\item 4.6 Precios hedónicos
\item 4.7 Modelos no anidados
\item 4.8 Regresión cuantílica
\end{enumerate}
\column{0.5\textwidth}
\begin{enumerate}\setcounter{enumi}{10}
\item 4.9 Aplicaciones financieras
\item Multicolinealidad: diagnóstico
\item Multicolinealidad: corrección
\item MANOVA
\item 5.1--5.2 Residuos
\item 5.3--5.6 Heteroscedasticidad
\item 5.7--5.10 Autocorrelación
\item 5.11--5.12 Dinámica y rezagos
\item 5.13 Jarque--Bera
\item Predicción y conclusiones
\end{enumerate}
\end{columns}
\end{frame}

\section{Datos}
\begin{frame}{Pregunta y datos}
\textbf{¿Cuánto aumenta el salario un año más de educación?} Los controles (experiencia, edad, habilidad,
escolaridad de los padres) están muy correlacionados $\Rightarrow$ \alert{multicolinealidad}.
\[
\ln w_i=\beta_0+\beta_1 educ_i+\beta_2 exper_i+\beta_3 exper_i^2+\beta_4 tenure_i+\dots+\varepsilon_i
\]
\centering\small
@TAB_DATOS@
\end{frame}

\section{Regresión lineal simple}
\begin{frame}{Regresión lineal simple: $\ln w_i=\beta_0+\beta_1 educ_i+u_i$}
\begin{columns}[T]
\column{0.42\textwidth}\small
\[\hat\beta_1=\frac{S_{xy}}{S_{xx}}=\frac{@Sxy@}{@Sxx@}=@b1@\]
\[\hat\beta_0=\bar y-\hat\beta_1\bar x=@b0@\]
\begin{itemize}
\item $se(\hat\beta_1)=@se@$, \ $t=@t@$
\item IC 95\%: $[@ic0@,\ @ic1@]$
\item $R^2=\hat\rho_{xy}^2=@R2@$
\item $F_{1,720}=t^2=@F@$
\item $\sum\hat u_i=0$, \ $\sum x_i\hat u_i=0$
\end{itemize}
\column{0.58\textwidth}
\includegraphics[width=\linewidth]{simple_recta}\\[0.3em]
\small +1 año $\Rightarrow$ \alert{+@pb1@\%} de salario (exacto @pex@\%), pero la educación sola explica sólo el @pR2@\%.
\end{columns}
\end{frame}

\begin{frame}{Regresión simple: propiedades}
\begin{columns}[T]
\column{0.45\textwidth}
\includegraphics[width=\linewidth]{simple_scr}
\column{0.55\textwidth}\small
\begin{itemize}
\item Superficie $SCR(b_0,b_1)$: mínimo en MCO. Elipse alargada porque $\mathrm{corr}(\hat\beta_0,\hat\beta_1)=@cb@$ ($\bar x$ lejos de 0).
\item Dummy: en $\ln w\sim urban$, $\hat\beta_0=@d0@=\bar y_{rural}$ y $\hat\beta_1=@d1@=\bar y_{urb}-\bar y_{rural}$.
\item Reescalar $x'=12\,educ$, $y'=w/1000$: la pendiente se multiplica por $c/b$ y $R^2$ no cambia.
\item Sin las 10 observaciones más influyentes (Cook): $\hat\beta_1$ pasa de @b1@ a @b1s@.
\item Gauss--Markov (5{,}000 simulaciones): $\mathrm{Var}_{MCO}=@vm@<\mathrm{Var}_{alt}=@va@$.
\end{itemize}
\end{columns}
\end{frame}
""", AUTORA=AUTORA, CURSO=CURSO, FECHA=FECHA,
    TAB_DATOS=tex_tabla(["Base", "Fuente", "Uso"], [["wage2", "Wooldridge (NLS 1980), $n=@n@$", "Base principal"],
                                                     ["hprice1", "Wooldridge, 88 casas", "4.6 Hedónicos"],
                                                     ["stocks", "Plotly, semanal 2018--2019", "4.9 Finanzas"],
                                                     ["macrodata", "statsmodels, 1959--2009", "5.7--5.12 Series de tiempo"]], "lll"),
    Sxy=f"{S['Sxy']:.2f}", Sxx=f"{S['Sxx']:.1f}", b1=f"{S['b1']:.4f}", b0=f"{S['b0']:.4f}", se=f"{S['se_b1']:.5f}",
    t=f2(S["t"]), ic0=f4(S["ic"][0]), ic1=f4(S["ic"][1]), R2=f4(S["R2"]), F=f2(S["F"]), pb1=f"{100*S['b1']:.1f}",
    pex=f"{S['efecto_exacto']:.1f}", pR2=f"{100*S['R2']:.1f}", cb=f3(S["corr_b0_b1"]), d0=f4(S["dummy"]["b0"]),
    d1=f4(S["dummy"]["b1"]), b1s=f"{S['b1_sin_influyentes']:.4f}", vm=f"{R['gauss_markov']['var_mco']:.1e}".replace("e-05", r"\times10^{-5}"),
    va=f"{R['gauss_markov']['var_alt']:.1e}".replace("e-05", r"\times10^{-5}"), n=n)

B += T(r"""
\section{4.1--4.3}
\begin{frame}{4.1 Supuestos del modelo de regresión lineal múltiple}
\centering\small
@TAB@
\end{frame}

\begin{frame}{4.2 Parámetros: formas funcionales e interpretación ceteris paribus}
\centering\small
@TABF@
\vspace{0.8em}
\begin{itemize}\small
\item Cada $\beta_j$ es un efecto parcial: compara trabajadores iguales en las demás regresoras.
\item Si dos regresoras casi siempre se mueven juntas, hay poca información para separar sus efectos.
\end{itemize}
\end{frame}
""", TAB=tex_tabla(["Supuesto", "¿Se cumple?", "Evidencia"], [
    ["Linealidad en parámetros", "Sí", f"RESET $p={R['reset']['p']:.2f}$"],
    ["Rango completo", r"Falla con $age-educ-6$", f"rango {P['rango']} de {P['cols']}"],
    [r"$E(\varepsilon|X)=0$", "Dudoso sin habilidad", rf"$\hat\beta_{{educ}}$: {C['Base']['beta_educ']:.3f} $\to$ {C['Completo']['beta_educ']:.3f}"],
    ["Homoscedasticidad", "Rechazada (leve)", f"BP $p={H['bp_p']:.3f}$"],
    ["No autocorrelación", "No aplica (corte transversal)", "DW depende del orden"],
    ["Normalidad", r"Rechazada; $n$ grande", f"JB $={JB['JB']:.0f}$"]], "lll"))
B = B.replace("@TABF@", tex_tabla(["Forma", "Modelo", r"$\hat\beta$", "Lectura"], formas_tex, "llrl"))

B += T(r"""
\begin{frame}{4.3 ANOVA y prueba F global}
\[F=\frac{(SSR_R-SSR_{NR})/q}{SSR_{NR}/(n-k-1)}\overset{H_0}{\sim}F_{q,\,n-k-1}\]
\centering\small
@TAB@
\vspace{0.6em}

$F=@F@>F_{0.05;\,@k@,\,@d@}=@Fc@$ $\Rightarrow$ el modelo es globalmente significativo.
\end{frame}

\begin{frame}{4.3 Pruebas F de restricciones}
\centering\scriptsize
@TAB2@
\vspace{0.5em}

\includegraphics[height=0.42\textheight]{regiones_rechazo}

\small $age$ sola no es significativa ($F=t^2=@Fa@$), pero $\{exper, exper^2, age\}$ \alert{sí lo es en conjunto}.
\end{frame}
""", TAB=tex_tabla(["Fuente", "SC", "gl", "CM"], [["Explicada (SSE)", f3(A["SSE"]), A["k"], f4(A["MSR"])],
                                                  ["Residual (SSR)", f3(A["SSR"]), A["dfr"], f4(A["MSE"])],
                                                  ["Total (SST)", f3(A["SST"]), A["k"] + A["dfr"], ""]], "lrrr"),
       F=f2(A["F"]), k=A["k"], d=A["dfr"], Fc=f3(A["Fcrit"]),
       TAB2=tex_tabla([r"$H_0$", "$q$", "$F$", r"$F_{0.05}$", "$p$", "Decisión"], pf_rows, "lrrrrl"), Fa=f2(PF[3]["F"]))

B += T(r"""
\section{4.4--4.7}
\begin{frame}{4.4 Dummies e interacciones}
\[\ln w=\beta_0+\beta_1 educ+\dots+\delta\,black+\gamma\,(educ\times black)+\varepsilon\]
\begin{itemize}
\item Bases: no casado, no negro, fuera del sur, rural. Efecto exacto $100(e^{\beta}-1)$.
\item Casado: \alert{+@mar@\%}; urbano: \alert{+@urb@\%}; sur: @sur@\%.
\item Pendiente de educación: grupo base @eb@; negros $\beta_1+\gamma=@pb@$ ($\gamma$: $p=@pi@$).
\item $\delta$ sola es la diferencia sólo si $educ=0$; con 12 años: $\delta+12\gamma=@d12@$.
\item \textbf{Trampa de las dummies}: constante $+$ urban $+$ rural $\Rightarrow$ rango @r0@ de @r1@ (colinealidad perfecta).
\end{itemize}
\end{frame}

\begin{frame}{4.5 $R^2$ y $R^2$ ajustada \quad\textbar\quad 4.6 Precios hedónicos}
\begin{columns}[T]
\column{0.45\textwidth}\centering\small
@TABR2@\\[0.5em]
\raggedright Agregar \texttt{sibs} sube $R^2$ pero baja $\bar R^2$.
\column{0.55\textwidth}\small
\textbf{hprice1} (88 casas):
\[\ln P=\beta_0+\beta_1\ln lot+\beta_2\ln sqrft+\beta_3 bdrms+\beta_4 colonial\]
\begin{itemize}
\item Elasticidad precio--superficie: @sq@ ($p<0.001$).
\item Recámaras: @bd@ ($p=@pbd@$): con la misma superficie no agregan valor.
\item Con el avalúo: VIF de $\ln sqrft$ @v1@ $\to$ @v2@; $se$ @s1@ $\to$ @s2@.
\end{itemize}
\end{columns}
\end{frame}

\begin{frame}{4.7 Modelos no anidados}
A: controles $+$ habilidad (IQ, KWW) \qquad B: controles $+$ familia (meduc, feduc)
\vspace{0.5em}
\centering\small
@TABNA@
\vspace{0.6em}
\begin{itemize}\small
\item AIC, BIC y $\bar R^2$ prefieren A.
\item La prueba J de Davidson--MacKinnon rechaza en \alert{ambas} direcciones: ninguno abarca al otro $\Rightarrow$ combinarlos\ldots\ y aparece la multicolinealidad.
\end{itemize}
\end{frame}
""", mar=f"{D['married_exacto']:.1f}", urb=f"{D['urban_exacto']:.1f}", sur=f"{pct(D['south']):.1f}".replace("-", "$-$"),
       eb=f4(D["educ"]), pb=f4(D["pend_black"]), pi=f"{D['p_int']:.2f}", d12=f3(D["dif_black_12"]).replace("-", "$-$"),
       r0=R["trampa_rango"][0], r1=R["trampa_rango"][1],
       TABR2=tex_tabla(["Modelo", "$k$", "$R^2$", r"$\bar R^2$"], r2_rows, "lrrr"),
       sq=f3(hp_["lsqrft"]), bd=f3(hp_["bdrms"]), pbd=f"{HED['p']['bdrms']:.2f}", v1=f2(HED["vif"]["lsqrft"]),
       v2=f2(HED["vif_assess"]["lsqrft"]), s1=f3(HED["se_lsqrft"]), s2=f3(HED["se_lsqrft_assess"]),
       TABNA=tex_tabla(["", "A (habilidad)", "B (familia)"],
                       [["AIC", f"{NA['aicA']:.1f}", f"{NA['aicB']:.1f}"], ["BIC", f"{NA['bicA']:.1f}", f"{NA['bicB']:.1f}"],
                        [r"$\bar R^2$", f4(NA["R2aA"]), f4(NA["R2aB"])],
                        ["Prueba J", rf"$\hat y_B$ en A: $t={NA['tJ_A']:.2f}$, $p={NA['pJ_A']:.3f}$",
                         rf"$\hat y_A$ en B: $t={NA['tJ_B']:.2f}$, $p<0.001$"]], "lll"))

B += T(r"""
\section{4.8--4.9}
\begin{frame}{4.8 Regresión cuantílica: $Q_\tau(\ln w\mid X)=X'\beta_\tau$}
\begin{columns}[T]
\column{0.55\textwidth}
\includegraphics[width=\linewidth]{cuantilica}
\column{0.45\textwidth}\centering\small
@TABQ@\\[0.6em]
\raggedright MCO: @ols@. El retorno va de @q1@\% a @q9@\%, pero los IC se traslapan: heterogeneidad moderada.
\end{columns}
\end{frame}

\begin{frame}{4.9 Aplicaciones financieras: beta de mercado de AAPL}
\begin{columns}[T]
\column{0.55\textwidth}
\includegraphics[width=\linewidth]{capm}
\column{0.45\textwidth}\small
Rendimientos log semanales 2018--2019, $n=@n@$.
\begin{itemize}
\item $\hat\beta=@b@$, IC 95\% $[@i0@,\ @i1@]$
\item $\hat\alpha=@a@$ ($p=@pa@$): sin alfa
\item BP $p=@bp@$; BG(4) $p=@bg@$
\item JB $p<0.001$: colas pesadas
\item AAPL $\sim$ MSFT $+$ AMZN $+$ GOOG: $F=@F@$ ($p<0.001$), pero sólo MSFT es significativa: acciones correlacionadas.
\end{itemize}
\end{columns}
\end{frame}
""", TABQ=tex_tabla([r"$\tau$", r"$\hat\beta_{educ}$", "ee", "IC 95\\%"], q_rows, "rrrr"), ols=f4(Q["ols"]),
       q1=f"{100*Q['tabla'][0]['beta_educ']:.1f}", q9=f"{100*Q['tabla'][4]['beta_educ']:.1f}", n=FIN["n"], b=f3(FIN["beta"]),
       i0=f3(FIN["ic_beta"][0]), i1=f3(FIN["ic_beta"][1]), a=f4(FIN["alpha"]), pa=f"{FIN['p_alpha']:.2f}",
       bp=f3(FIN["bp_p"]), bg=f"{FIN['bg_p']:.2f}", F=f2(FIN["fac_F"]))

B += T(r"""
\section{Multicolinealidad}
\begin{frame}{Multicolinealidad: perfecta vs imperfecta}
\textbf{Perfecta:} experiencia potencial $pexper=age-educ-6$. Con constante, $educ$, $age$ y $pexper$:
rango$(X)=@ra@<@co@$, $\kappa(X'X)\approx @kc@$ $\Rightarrow$ $(X'X)^{-1}$ no existe.
\vspace{1em}

\textbf{Imperfecta} (modelo completo, 13 regresores):
\vspace{0.6em}

\centering
\dato{@vmax@}{VIF máximo ($exper^2$)}\hfill
\dato{@nc@}{número de condición (BKW) $>30$}\hfill
\dato{@fg@}{$\chi^2$ Farrar--Glauber (crítico @fgc@)}
\vspace{1em}

\raggedright\small
$t$ individuales: $exper$ @te@, $exper^2$ @te2@, $age$ @ta@ \ldots\ pero $F$ conjunta $=@Fcv@$ ($p=@pcv@$).
\end{frame}

\begin{frame}{Correlaciones, VIF, tolerancia y regla de Klein}
\begin{columns}[T]
\column{0.5\textwidth}
\includegraphics[width=\linewidth]{correlaciones}
\column{0.5\textwidth}\centering\scriptsize
$VIF_j=\dfrac{1}{1-R_j^2}$, \quad $Tol_j=1-R_j^2$
\vspace{0.3em}

@TABVIF@\\[0.3em]
Klein: $R_j^2>R^2=@R2@$ en 8 de 13 regresores.
\end{columns}
\end{frame}

\begin{frame}{Autovalores, número de condición y Farrar--Glauber}
\begin{itemize}
\item Regresoras estandarizadas: $\lambda_{\max}=@lmax@$, $\lambda_{\min}=@lmin@$ $\Rightarrow$ $\kappa=\sqrt{\lambda_{\max}/\lambda_{\min}}=@ic@$.
\item Columnas normalizadas con constante (Belsley--Kuh--Welsch): número de condición $=@bkw@$ $>30$ (severa).
\item Farrar--Glauber: $\chi^2=-\left[n-1-\frac{2k+5}{6}\right]\ln|R|$, con $|R|=@det@$: $\chi^2=@fg@>\chi^2_{0.05;\,@gl@}=@fgc@$ $\Rightarrow$ se rechaza ortogonalidad.
\item Síntoma clásico: $se(\hat\beta_{educ})$ sube de @se0@ (base) a @se1@ (completo).
\item \alert{MCO sigue siendo MELI}: la multicolinealidad no sesga, vuelve imprecisos los estimadores.
\end{itemize}
\end{frame}

\begin{frame}{Corrección: centrar, eliminar/combinar, PCA y Ridge}
\begin{columns}[T]
\column{0.5\textwidth}
\includegraphics[width=\linewidth]{vif}
\column{0.5\textwidth}
\includegraphics[width=\linewidth]{ridge}
\end{columns}
\small
\begin{itemize}
\item Centrar: $\mathrm{corr}(exper,exper^2)$ @c0@ $\to$ @c1@; VIF máx @vm0@ $\to$ @vm1@, mismo $R^2$.
\item Eliminar $age$ y promediar padres (prueba $F$ de $\beta_{meduc}=\beta_{feduc}$: $p=@pmf@$). PCA: CP1 $=@pca@\%$ de la varianza. Ridge: $\lambda\approx @lam@$.
\end{itemize}
\end{frame}

\begin{frame}{Comparación de especificaciones}
\begin{columns}[T]
\column{0.60\textwidth}\centering\footnotesize\setlength{\tabcolsep}{4pt}
@TABC@
\column{0.38\textwidth}
\includegraphics[width=\linewidth]{educ}
\end{columns}
\vspace{0.5em}
\small La caída de $\hat\beta_{educ}$ (base $\to$ completo) es \alert{sesgo por variable omitida} (habilidad); la multicolinealidad se nota en el error estándar.
\end{frame}
""", ra=P["rango"], co=P["cols"], kc=f"{P['cond']:.1e}".replace("e+16", r"\times10^{16}"), vmax=f"{VF['exper²']['VIF']:.0f}",
       nc=f"{R['num_cond_bkw']:.0f}", fg=f"{FG['chi2']:.0f}", fgc=f"{FG['crit']:.1f}", te=f2(F_["t"]["exper"]),
       te2=f2(F_["t"]["I(exper ** 2)"]).replace("-", "$-$"), ta=f2(F_["t"]["age"]), Fcv=f2(PF[1]["F"]), pcv=f3(PF[1]["p"]),
       TABVIF=tex_tabla(["Variable", "VIF", "Tol.", r"$R_j^2$"], vif_rows, "lrrr"), R2=f3(F_["R2"]),
       lmax=f3(R["autovalores"][0]), lmin=f4(R["autovalores"][-1]), ic=f2(R["indices_condicion"][-1]), bkw=f"{R['num_cond_bkw']:.1f}",
       det=f4(FG["detR"]), gl=f"{FG['gl']:.0f}", se0=f4(R["base"]["bse"]["educ"]), se1=f4(F_["bse"]["educ"]),
       c0=f3(R["corr_exper_exper2"]), c1=f3(R["corr_experc"]), vm0=f"{C['Completo']['VIF_max']:.1f}", vm1=f2(R["vif_centrado_max"]),
       pmf=f"{PF[4]['p']:.2f}", pca=f"{100*R['pca']['var'][0]:.0f}", lam=f"{R['ridge']['alpha']:.0f}",
       TABC=tex_tabla(["Modelo", r"$\hat\beta_{educ}$", "ee", r"$\bar R^2$", "BIC", "VIF máx", "$k$"], comp_rows, "lrrrrrr"))

B += T(r"""
\section{MANOVA y 5.1--5.6}
\begin{frame}{MANOVA \quad\textbar\quad 5.1--5.2 Distribución de los residuos}
\begin{columns}[T]
\column{0.4\textwidth}\small
\textbf{MANOVA}: ¿difieren $(educ, IQ, KWW)$ entre el sur y el resto?
\begin{itemize}
\item Wilks $\Lambda=@w@$
\item Pillai $=@pi@$
\item $F(@n1@,@n2@)=@F@$, $p<0.001$
\end{itemize}
MANOVA trata varias $Y$; \alert{no} diagnostica multicolinealidad (propiedad de las columnas de $X$).
\vspace{0.5em}

\textbf{Residuos}: sin curvatura ni abanico marcado; cola izquierda pesada; Cook máx $=@ck@$.
\column{0.6\textwidth}
\includegraphics[width=\linewidth]{residuos}
\end{columns}
\end{frame}

\begin{frame}{5.3--5.6 Heteroscedasticidad: Breusch--Pagan paso a paso}
\begin{enumerate}
\item Estimar el modelo reducido por MCO y obtener $\hat\varepsilon_i$.
\item Regresión auxiliar de $\hat\varepsilon_i^2$ sobre las $q=@q@$ regresoras: $R^2_{aux}=@r2@$.
\item $LM=nR^2_{aux}=@n@\times@r2@=@LM@$.
\item $LM>\chi^2_{0.05;\,@q@}=@cr@$ ($p=@p@$) $\Rightarrow$ \alert{se rechaza homoscedasticidad}.
\end{enumerate}
\begin{itemize}
\item White: $p=@wp@$ (pierde potencia por los muchos términos).
\item Errores HC0--HC3 casi iguales a los clásicos; WLS/FGLS: $\hat\beta_{educ}=@wls@$.
\item Los coeficientes MCO no cambian con errores robustos: sólo cambia su incertidumbre.
\end{itemize}
\end{frame}

\begin{frame}{Modelo final con errores robustos (HC1)}
\centering\scriptsize
@TABRED@
\vspace{0.5em}

\small $n=@nn@$, $R^2=@R2@$. Variable dependiente: $\ln w$. Bases: no casado, no negro, fuera del sur, rural.
\end{frame}
""", w=f4(mw["valor"]), pi=f4(mp_["valor"]), n1=f"{mw['num']:.0f}", n2=f"{mw['den']:.0f}", F=f2(mw["F"]), ck=f3(R["cook_max_red"]),
       q=H["q"], r2=f4(H["R2aux"]), n=n, LM=f2(H["LM"]), cr=f2(H["crit"]), p=f3(H["bp_p"]), wp=f"{H['white_p']:.2f}",
       wls=f4(H["wls_educ"]), TABRED=tex_tabla(["Variable", r"$\hat\beta$", "ee HC1", "$p$", "Efecto \\% exacto"], red_rows, "lrrrr"),
       nn=n, R2=f3(RED["R2"]))

B += T(r"""
\section{5.7--5.13}
\begin{frame}{5.7--5.10 Autocorrelación: Durbin--Watson, Breusch--Godfrey y HAC}
\includegraphics[width=\linewidth]{autocorrelacion}
\small
\begin{itemize}
\item wage2: DW $=@d0@$ y $@d1@$ al reordenar filas $\Rightarrow$ no aplica en corte transversal.
\item Consumo (macrodata, $n=@n@$), $\ln c=\alpha+\beta\ln y$: $R^2=@R2@$, DW $=@dw@\approx2(1-\hat\rho)$ con $\hat\rho=@rho@$; BG(4) $=@bg@$.
\item $se$ HAC (Newey--West) $=@sh@$ vs clásico @sc@: la inferencia clásica subestima la incertidumbre.
\end{itemize}
\end{frame}

\begin{frame}{5.11--5.12 Modelos dinámicos y rezagos}
\[\ln c_t=\alpha+\phi\ln c_{t-1}+\beta_0\ln y_t+\beta_1\ln y_{t-1}+u_t\]
\begin{columns}[T]
\column{0.5\textwidth}\small
\begin{itemize}
\item $\hat\phi=@phi@$ ($|\phi|<1$: estable)
\item $\hat\beta_0=@b0@$, \ $\hat\beta_1=@b1@$
\item Largo plazo: $\dfrac{\beta_0+\beta_1}{1-\phi}=@lr@\approx1$
\item $\mathrm{corr}(\ln y_t,\ln y_{t-1})=@cy@$: VIF $>3{,}000$ (\alert{multicolinealidad entre rezagos})
\item BG(4) aún rechaza: más dinámica o HAC
\end{itemize}
\column{0.5\textwidth}\centering\scriptsize
Rezagos distribuidos: $\Delta\ln c_t$ sobre $\Delta\ln y_{t-j}$\\[0.3em]
@TABREZ@\\[0.3em]
\small BIC elige @pb@ rezago; AIC elige @pa@.
\end{columns}
\end{frame}

\begin{frame}{5.13 Jarque--Bera y resumen de diagnósticos}
\[JB=\frac{n}{6}\left(S^2+\frac{(K-3)^2}{4}\right)=\frac{@nw@}{6}\left(@S@^2+\frac{(@K@-3)^2}{4}\right)=@JB@>\chi^2_{0.05;2}=5.99\]
\centering\small
@TABD@
\vspace{0.4em}

\raggedright Rechazar normalidad con $n=@nw@$ no implica sesgo: la inferencia es válida asintóticamente.
\end{frame}

\section{Cierre}
\begin{frame}{Inferencia vs predicción}
\begin{columns}[T]
\column{0.5\textwidth}\centering
@TABP@
\column{0.5\textwidth}
\begin{itemize}
\item Completo $\approx$ reducido fuera de muestra.
\item La multicolinealidad daña la \alert{interpretación}, no la predicción.
\item Ridge: menor RMSE a cambio de sesgo (no reporta $p$-values).
\end{itemize}
\end{columns}
\end{frame}

\begin{frame}{Conclusiones}
\begin{enumerate}
\item Retorno a la educación robusto: @s@\% (simple) $\to\approx$@r@\% con habilidad (sesgo por variable omitida).
\item Perfecta: no hay MCO único. Imperfecta: MCO sigue siendo MELI, pero las varianzas se inflan.
\item VIF, tolerancia, Klein, autovalores, número de condición y Farrar--Glauber señalan $exper$, $exper^2$ y $age$.
\item La $F$ conjunta revela lo que las $t$ individuales no pueden.
\item Centrar, combinar, PCA o Ridge corrigen sin perder ajuste.
\item Heteroscedasticidad leve (HC1); no normalidad irrelevante con $n$ grande; en series de tiempo, dinámica o HAC.
\item La misma lección aparece en hedónicos, finanzas y rezagos.
\end{enumerate}
\end{frame}

\begin{frame}{Referencias}
\small
\begin{itemize}
\item Belsley, Kuh y Welsch (1980). \textit{Regression Diagnostics}. Wiley.
\item Davidson y MacKinnon (1981). \textit{Econometrica}, 49(3), 781--793.
\item Farrar y Glauber (1967). \textit{Review of Economics and Statistics}, 49(1), 92--107.
\item Galindo, C. (2026). Notas de regresión lineal, pruebas $t$ y $F$, inferencia y diagnóstico. Universidad Panamericana.
\item Koenker y Bassett (1978). Regression quantiles. \textit{Econometrica}, 46(1), 33--50.
\item Mincer, J. (1974). \textit{Schooling, Experience, and Earnings}. NBER.
\item Wooldridge, J. M. (2020). \textit{Introducción a la econometría} (7.\textsuperscript{a} ed.). Cengage.
\end{itemize}
\end{frame}

\end{document}
""", d0=f3(R["dw_wage2"][0]), d1=f3(R["dw_wage2"][1]), n=AC["n"], R2=f4(AC["R2"]), dw=f3(AC["dw"]), rho=f3(AC["rho"]),
       bg=f"{AC['bg_LM']:.1f}", sh=f4(AC["se_hac"]), sc=f4(AC["se"]), phi=f4(pa["lc_1"]), b0=f4(pa["ly"]), b1=f4(pa["ly_1"]).replace("-", "$-$"),
       lr=f3(ARDL["lr"]), cy=f4(ARDL["corr_y_y1"]), TABREZ=tex_tabla(["$p$", "AIC", "BIC", r"$\sum\beta$", r"$\bar R^2$"], rez_rows, "rrrrr"),
       pb=REZ["p_bic"], pa=REZ["p_aic"])
B = T(B, nw=n, S=f"({JB['S']:.3f})".replace("-", "-"), K=f3(JB["K"]), JB=f"{JB['JB']:.1f}",
      TABD=tex_tabla(["Prueba", "Estadístico", "$p$", "Conclusión (5\\%)"], diag_rows, "lllll"[:4]),
      TABP=tex_tabla(["Modelo", r"$R^2$ prueba", "RMSE"], pred_rows, "lrr"),
      s=f"{100*S['b1']:.1f}", r=f"{100*RED['params']['educ']:.1f}")

(ROOT / "Presentacion_Beamer.tex").write_text(B, encoding="utf-8")
with zipfile.ZipFile(ROOT / "Beamer_Overleaf.zip", "w", zipfile.ZIP_DEFLATED) as z:
    z.write(ROOT / "Presentacion_Beamer.tex", "Presentacion_Beamer.tex")
    for f in sorted(FIG.glob("*.png")):
        z.write(f, f"figuras/{f.name}")
print("Listo: Reporte_Multicolinealidad_Mincer.docx, Presentacion_Multicolinealidad_Mincer.pptx, Presentacion_Beamer.tex, Beamer_Overleaf.zip")
