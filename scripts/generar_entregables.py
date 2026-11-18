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
EQUIPO = "Ximena Lara, Camila Ricardi y Sara Enriquez"
AUTORA, CURSO, FECHA = EQUIPO, "Fundamentos de Econometría · Prof. César Galindo · Universidad Panamericana", "Noviembre 2026"

S, F_, A = R["simple"], R["full"], R["anova"]
C, VF, H, D = R["comparacion"], R["vif_full"], R["hetero"], R["dummies"]
PF = R["pruebas_F"]
FG, FIN, AC, ARDL, REZ = R["farrar"], R["finanzas"], R["autocorr"], R["ardl"], R["rezagos"]
JB, HED, NA, Q = R["jb"], R["hedonico"], R["no_anidados"], R["cuantil"]
RED = R["preferido"]  # modelo preferido: todos los controles, experiencia centrada
TP = R["tabla_principal"]
M1, M2, M3 = TP["(1) Simple"], TP["(2) Mincer"], TP["(3) Preferido"]
WH = "Rechaza apenas" if H["white_p"] < .05 else "No rechaza"
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
    ["White (wage2)", f"LM = {H['white_LM']:.2f}", p3(H["white_p"]), WH],
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
for txt, sz, b in [("Anexo técnico: temas complementarios del curso", 22, True),
                   ("Multicolinealidad en la ecuación de salarios de Mincer · regresión simple y múltiple, inferencia y diagnóstico con datos reales", 13, False)]:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ = p.add_run(txt)
    r_.bold, r_.font.size = b, Pt(sz)
for _ in range(4):
    doc.add_paragraph()
for line in ("Equipo 4: " + EQUIPO, CURSO, FECHA):
    doc.add_paragraph(line).alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_page_break()

doc.add_heading("Resumen", 1)
doc.add_paragraph(
    "Este anexo acompaña al informe principal (Informe_Proyecto_Econometria). Aplica a los mismos datos el resto de las "
    "herramientas del curso; las conclusiones del proyecto están en el informe. "
    f"Se estima una ecuación de salarios tipo Mincer con {n} trabajadores de la National Longitudinal Survey (base wage2 "
    "de Wooldridge) y se recorren todos los temas del segundo parcial: regresión lineal simple, supuestos, interpretación, "
    "pruebas t y F, variables dummy e interacciones, R² ajustada, precios hedónicos, modelos no anidados, regresión "
    "cuantílica, aplicaciones financieras, diagnóstico de residuos, heteroscedasticidad, autocorrelación, modelos "
    "dinámicos, rezagos y Jarque–Bera. El eje del proyecto es la multicolinealidad: en el modelo con 13 regresores, la "
    f"experiencia y su cuadrado tienen VIF de {VF['exper']['VIF']:.1f} y {VF['exper²']['VIF']:.1f}, la prueba de Farrar–Glauber "
    f"rechaza la ortogonalidad (χ² = {FG['chi2']:.0f}) y la experiencia y la edad no son significativas individualmente "
    f"pero sí en conjunto (F = {PF[1]['F']:.2f}, p = {PF[1]['p']:.3f}). Centrar, combinar regresores, componentes principales "
    f"y Ridge reducen el problema. La asociación entre educación y salario pasa de {100*S['b1']:.1f}% en la regresión simple a "
    f"≈{100*RED['params']['educ']:.1f}% al controlar por habilidad (asociación, no efecto causal), lo que refleja sesgo por variable omitida y no "
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
    f"Soluciones (el modelo preferido usa sólo (a); (b) y (c) son análisis de sensibilidad, porque eliminar controles para bajar un VIF cambia la pregunta): (a) centrar exper reduce corr(exper, exper²) de {R['corr_exper_exper2']:.3f} a {R['corr_experc']:.3f} y el "
    f"VIF máximo de {C['Completo']['VIF_max']:.1f} a {R['vif_centrado_max']:.2f} sin cambiar R²; (b) eliminar age y promediar "
    f"meduc y feduc (validado por la prueba F de β_meduc = β_feduc); (c) el primer componente principal de IQ, KWW, meduc y feduc "
    f"explica el {100*R['pca']['var'][0]:.1f}% de su varianza; (d) Ridge con λ ≈ {R['ridge']['alpha']:.0f} encoge primero "
    "a exper y exper².")
figura(doc, "vif", "Figura 8. VIF antes y después de centrar la experiencia.")
tabla(doc, T_comp, "Modelo")
pie(doc, "Tabla 12. Comparación de especificaciones.")
figura(doc, "educ", "Figura 9. Coeficiente de educ por especificación (IC 95%).", 4.8)
figura(doc, "ridge", "Figura 10. Trayectorias de Ridge.", 5)

doc.add_heading("13. MANOVA", 1)
mw, mp_ = R["manova"]["Wilks' lambda"], R["manova"]["Pillai's trace"]
doc.add_paragraph(
    f"¿Difieren conjuntamente educ, IQ y KWW entre el sur y el resto del país? Lambda de Wilks = {mw['valor']:.4f}, traza "
    f"de Pillai = {mp_['valor']:.4f}, F({mw['num']:.0f}, {mw['den']:.0f}) = {mw['F']:.2f}, p = {p3(mw['p'])}: sí difieren. "
    "MANOVA estudia varias variables dependientes a la vez; no diagnostica multicolinealidad, que es una propiedad de las "
    "columnas de X.")

doc.add_heading("14. Distribución de los residuos y diagnóstico (5.1–5.2)", 1)
figura(doc, "residuos", "Figura 11. Residuos vs ajustados, Q–Q, histograma y leverage–Cook (modelo preferido).")
doc.add_paragraph(f"No hay curvatura ni un abanico marcado. La cola izquierda es más pesada que la normal. Hay "
                  f"{R['n_cook_4n']} observaciones con distancia de Cook > 4/n, pero la mayor vale sólo "
                  f"{R['cook_max_red']:.3f}: ninguna domina la estimación.")

doc.add_heading("15. Heteroscedasticidad y Breusch–Pagan (5.3–5.6)", 1)
doc.add_paragraph(
    f"Regresión auxiliar de ê² sobre las {H['q']} regresoras: R²aux = {H['R2aux']:.4f}, LM = n·R²aux = {n}×{H['R2aux']:.4f} = "
    f"{H['LM']:.2f} > χ²({H['q']}; 0.05) = {H['crit']:.2f} (p = {H['bp_p']:.3f}): se rechaza la homoscedasticidad. White "
    f"{'la rechaza apenas' if H['white_p'] < .05 else 'no la rechaza'} (p = {H['white_p']:.3f}). Los errores HC0–HC3 casi no cambian y WLS/FGLS da β educ = {H['wls_educ']:.4f} "
    f"(ee {H['wls_se']:.4f}): la heteroscedasticidad es leve. Por prudencia, el modelo final se reporta con errores HC1.")
tabla(doc, T_red, "Variable")
pie(doc, f"Tabla 13. Modelo preferido (todos los controles, experiencia centrada) con errores robustos HC1; R² = {RED['R2']:.3f}.")

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
    f"La asociación entre educación y salario es positiva y robusta: {100*S['b1']:.1f}% en la regresión simple y ≈4–5% al controlar habilidad y "
    "entorno familiar. La caída se debe a sesgo por variable omitida, no a multicolinealidad.",
    "La multicolinealidad perfecta impide estimar; la imperfecta no sesga MCO (sigue siendo MELI) pero infla las varianzas.",
    "VIF, tolerancia, Klein, autovalores, número de condición y Farrar–Glauber coinciden: el problema está en exper, exper² y age.",
    "La prueba F conjunta detecta lo que las pruebas t individuales no pueden: el ciclo de vida sí importa.",
    "Centrar resuelve la colinealidad estructural del polinomio; eliminar o combinar controles, PCA y Ridge sólo se usan como sensibilidad: β educ apenas cambia.",
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
doc.save(ROOT / "Anexo_Tecnico_Temas_del_Curso.docx")

print("Listo: Anexo_Tecnico_Temas_del_Curso.docx")
