"""Genera el informe principal (5–8 páginas) que pide la guía del proyecto, a partir de resultados.json.

Uso: .venv/bin/python scripts/generar_informe.py
"""
import json
from pathlib import Path

import numpy as np
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
FIG = ROOT / "figuras"
R = json.loads((ROOT / "resultados.json").read_text())
EQUIPO = "Ximena Lara, Camila Ricardi y Sara Enriquez"
n, n0 = R["n"], R["n_original"]
TP = R["tabla_principal"]
M1, M2, M3 = TP["(1) Simple"], TP["(2) Mincer"], TP["(3) Preferido"]
S, F_, VF, PF, C = R["simple"], R["full"], R["vif_full"], R["pruebas_F"], R["comparacion"]
H, JB, D = R["hetero"], R["jb"], R["descriptivos"]
pct = lambda b: 100 * (np.exp(b) - 1)
p3 = lambda p: "< 0.001" if p < 0.001 else f"{p:.3f}"
estrellas = lambda p: "***" if p < .01 else "**" if p < .05 else "*" if p < .1 else ""

doc = Document()
st = doc.styles["Normal"]
st.font.name, st.font.size = "Calibri", Pt(10.5)
st.paragraph_format.space_after = Pt(4)
for s_ in doc.sections:
    s_.left_margin = s_.right_margin = Inches(1)
    s_.top_margin = s_.bottom_margin = Inches(0.9)
for lvl, size in [(1, 13), (2, 11)]:
    h = doc.styles[f"Heading {lvl}"]
    h.font.size, h.font.color.rgb = Pt(size), RGBColor(0x1D, 0x24, 0x33)
    h.paragraph_format.space_before, h.paragraph_format.space_after = Pt(9), Pt(3)


def par(txt, **kw):
    p = doc.add_paragraph()
    # **negritas** simples
    for i, trozo in enumerate(txt.split("**")):
        r_ = p.add_run(trozo)
        r_.bold = i % 2 == 1
    for k, v in kw.items():
        setattr(p, k, v)
    return p


def pie(txt):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ = p.add_run(txt)
    r_.italic, r_.font.size = True, Pt(8.5)


def tabla(cab, filas, anchos=None, size=8.5):
    t = doc.add_table(rows=1, cols=len(cab))
    t.style = "Light Grid Accent 1"
    for j, c in enumerate(cab):
        t.rows[0].cells[j].text = c
    for f in filas:
        cells = t.add_row().cells
        for j, v in enumerate(f):
            cells[j].text = str(v)
    for r_ in t.rows:
        for j, c_ in enumerate(r_.cells):
            if anchos:
                c_.width = Inches(anchos[j])
            for p in c_.paragraphs:
                p.paragraph_format.space_after = Pt(0)
                for run in p.runs:
                    run.font.size = Pt(size)
    return t


def figura(nombre, txt, w=4.6):
    doc.add_picture(str(FIG / f"{nombre}.png"), width=Inches(w))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    pie(txt)


def bullets(items):
    for b in items:
        p = doc.add_paragraph(style="List Bullet")
        for i, trozo in enumerate(b.split("**")):
            r_ = p.add_run(trozo)
            r_.bold = i % 2 == 1
        p.paragraph_format.space_after = Pt(2)


# ------------------------------------------------------------------ encabezado
for txt, size, bold in [("Escolaridad, habilidad y salario: ¿la multicolinealidad impide medir la asociación?", 16, True),
                        ("Una ecuación de Mincer con datos de la NLS (EE. UU., 1980)", 11.5, False),
                        (f"Equipo 4: {EQUIPO}", 11, True),
                        ("Fundamentos de Econometría · Prof. César Galindo · Universidad Panamericana · Noviembre 2026", 9.5, False)]:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    r_ = p.add_run(txt)
    r_.bold, r_.font.size = bold, Pt(size)

b1, b3 = M1["params"]["educ"], M3["params"]["educ"]
ci3 = M3["ci"]["educ"]
par(f"**Resumen.** Con {n} hombres trabajadores de la National Longitudinal Survey, estimamos por MCO la asociación "
    f"entre escolaridad y salario. En la regresión simple, un año más de escolaridad se asocia con un salario "
    f"{pct(b1):.1f}% mayor. Al comparar personas con igual experiencia, antigüedad, habilidad medida (IQ, KWW) y "
    f"escolaridad de los padres, la asociación baja a {pct(b3):.1f}% (IC 95% con errores robustos: "
    f"{100*ci3[0]:.1f}%–{100*ci3[1]:.1f}% en log-puntos). Los controles están correlacionados con la escolaridad "
    f"(multicolinealidad moderada, VIF de educ = {VF['educ']['VIF']:.2f}): eso aumenta el error estándar, pero no impide "
    "responder la pregunta. Los VIF muy altos de la experiencia y su cuadrado son colinealidad estructural del polinomio "
    "y desaparecen al centrar. Los resultados son asociaciones condicionadas, no efectos causales.")

# ------------------------------------------------------------------ 1. Pregunta
doc.add_heading("1. Pregunta, hipótesis y motivación", 1)
par("**Pregunta.** ¿Cómo se asocia la escolaridad con el salario mensual de hombres trabajadores en EE. UU. en 1980, "
    "cuando se comparan personas con igual experiencia, antigüedad, habilidad medida y antecedentes familiares? ¿Qué "
    "tan precisamente podemos medir esa asociación si esos controles están correlacionados entre sí?")
par("**Hipótesis.** H₀: β_educ = 0 contra H₁: β_educ > 0, con un nivel de significancia α = 5% fijado antes de estimar. "
    "Esperamos una asociación positiva, pero menor que en la regresión simple: las personas con más habilidad suelen "
    "estudiar más y también ganar más, de modo que, sin controlar habilidad, β_educ captura parte de ella (sesgo por "
    "variable omitida).")
par("**Relevancia económica.** La asociación entre escolaridad y salario guía decisiones de inversión en educación y de "
    "política educativa. Además, es un caso natural de multicolinealidad: escolaridad, habilidad y escolaridad de los "
    "padres se mueven juntas, y la pregunta del proyecto es si esa redundancia impide separar sus asociaciones.")

# ------------------------------------------------------------------ 2. Teoría
doc.add_heading("2. Teoría breve", 1)
par("La ecuación de Mincer (1974) modela el logaritmo del salario como función lineal de los años de escolaridad y "
    "cuadrática de la experiencia: ln w = β₀ + β₁·educ + β₂·exper + β₃·exper² + u. Así, β₁ aproxima el cambio porcentual "
    "del salario asociado con un año más de escolaridad, y β₂ > 0 con β₃ < 0 describe un perfil de ingresos cóncavo a "
    "lo largo de la vida laboral. Si una variable omitida (habilidad) se correlaciona con educ y con el salario, MCO "
    "no estima β₁ de forma insesgada (E[u | X] ≠ 0). Por eso agregamos proxies de habilidad. El costo es que estos "
    "controles están correlacionados con educ (**multicolinealidad imperfecta**): MCO sigue siendo insesgado si se cumple "
    "la exogeneidad, pero sus varianzas crecen en el factor VIF_j = 1/(1 − R²_j). La multicolinealidad **perfecta** "
    "(un regresor es combinación exacta de otros) impide estimar.")

# ------------------------------------------------------------------ 3. Datos
doc.add_heading("3. Datos", 1)
par(f"**Fuente.** Base WAGE2 de Wooldridge (Introductory Econometrics), construida por Blackburn y Neumark (1992) a "
    f"partir de la National Longitudinal Survey of Young Men (EE. UU., 1980). La cargamos con el paquete de Python "
    f"wooldridge (v0.5.0) el 18/11/2026. **Unidad de análisis:** hombre trabajador de 28 a 38 años; corte transversal "
    f"de 1980. **Muestra:** de {n0} observaciones eliminamos {n0 - n} sin escolaridad de la madre o del padre; quedan "
    f"**n = {n}**. Esta eliminación puede introducir selección (sección 7). El archivo utilizado "
    "(data/wage2_muestra.csv), el diccionario completo y las transformaciones están en data/DICCIONARIO.md.")
filas = [["wage", "Ingreso mensual (USD de 1980)"], ["lwage", "ln(wage): variable dependiente"],
         ["educ", "Años de escolaridad"], ["exper", "Años de experiencia laboral"], ["tenure", "Años con el empleador actual"],
         ["age", "Edad (años)"], ["IQ", "Coeficiente intelectual (proxy de habilidad)"],
         ["KWW", "Prueba de conocimiento del mundo laboral (proxy de habilidad)"],
         ["meduc", "Escolaridad de la madre (años)"], ["feduc", "Escolaridad del padre (años)"]]
tabla(["Variable", "Definición y unidad", "Media", "Desv. est.", "Mín.", "Máx."],
      [[v, d, f"{D[v]['mean']:.2f}", f"{D[v]['std']:.2f}", f"{D[v]['min']:.0f}" if v != "lwage" else f"{D[v]['min']:.2f}",
        f"{D[v]['max']:.0f}" if v != "lwage" else f"{D[v]['max']:.2f}"] for v, d in filas],
      [0.7, 2.7, 0.65, 0.75, 0.5, 0.55])
pie(f"Tabla 1. Estadística descriptiva (n = {n}). Dummies: married, black, south, urban (1 = sí). Grupo de referencia: "
    "no casado, no afroamericano, fuera del sur, fuera de área metropolitana.")
cv = R["corr_vars"]
ce = dict(zip(cv, R["corr"][cv.index("educ")]))
par(f"**Exploración.** El salario es asimétrico a la derecha (media {D['wage']['mean']:.0f} USD, máximo "
    f"{D['wage']['max']:.0f}), por eso se modela en logaritmos. La escolaridad está correlacionada con IQ "
    f"(r = {ce['IQ']:.2f}), KWW ({ce['KWW']:.2f}), feduc ({ce['feduc']:.2f}) y meduc ({ce['meduc']:.2f}), y "
    f"negativamente con la experiencia ({ce['exper']:.2f}), porque estudiar más retrasa la entrada al mercado laboral. Estas "
    "correlaciones anticipan la multicolinealidad (Figura 1).")
figura("correlaciones", "Figura 1. Matriz de correlaciones entre los regresores.", 3.9)

# ------------------------------------------------------------------ 4. Especificación
doc.add_heading("4. Especificación", 1)
p = par("ln wᵢ = β₀ + β₁educᵢ + β₂experᶜᵢ + β₃(experᶜᵢ)² + β₄tenureᵢ + β₅ageᵢ + β₆IQᵢ + β₇KWWᵢ + β₈meducᵢ + β₉feducᵢ "
        "+ β₁₀marriedᵢ + β₁₁blackᵢ + β₁₂southᵢ + β₁₃urbanᵢ + uᵢ")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.runs[0].italic = True
bullets([
    "**Variable de interés:** educ. Signo esperado positivo; su coeficiente se lee como cambio porcentual aproximado del "
    "salario (100·β), y el cambio exacto es 100·(e^β − 1).",
    "**Experiencia centrada** (experᶜ = exper − media) y su cuadrado: perfil cóncavo (β₂ > 0, β₃ < 0). Centrar no cambia "
    "el ajuste ni los demás coeficientes, pero elimina la colinealidad artificial entre exper y exper².",
    "**tenure** (+): capital humano específico de la empresa. **age** (?): con educ y exper incluidas, la edad aporta poca "
    "variación propia (age ≈ educ + exper + 6); se conserva por diseño, porque edad y experiencia no son idénticas.",
    "**IQ y KWW** (+): proxies de habilidad, que es la razón principal para incluirlas. **meduc y feduc** (+): entorno "
    "familiar y calidad de la educación.",
    "**Dummies** married (+), black (−), south (−), urban (+); con m = 2 categorías cada una se incluye un solo "
    "indicador, para evitar la trampa de las variables ficticias.",
    "Errores estándar **robustos a heteroscedasticidad (HC1)**, justificados por la prueba de Breusch–Pagan (sección 6)."])

# ------------------------------------------------------------------ 5. Resultados
doc.add_heading("5. Resultados", 1)
VARS = [("educ", "educ"), ("exper_c", "exper (centrada)"), ("I(exper_c ** 2)", "exper² (centrada)"), ("tenure", "tenure"),
        ("age", "age"), ("IQ", "IQ"), ("KWW", "KWW"), ("meduc", "meduc"), ("feduc", "feduc"), ("married", "married"),
        ("black", "black"), ("south", "south"), ("urban", "urban"), ("Intercept", "constante")]


def celda(M, k):
    if k not in M["params"]:
        return ""
    return f"{M['params'][k]:.4f}{estrellas(M['p'][k])}\n({M['se'][k]:.4f})"


filas = [[lab, celda(M1, k), celda(M2, k), celda(M3, k)] for k, lab in VARS]
filas += [["R²", f"{M1['R2']:.3f}", f"{M2['R2']:.3f}", f"{M3['R2']:.3f}"],
          ["R² ajustada", f"{M1['R2a']:.3f}", f"{M2['R2a']:.3f}", f"{M3['R2a']:.3f}"],
          ["n", M1["n"], M2["n"], M3["n"]]]
tabla(["Variable dependiente: ln(wage)", "(1) Simple", "(2) Mincer", "(3) Preferido"], filas, [2.2, 1.3, 1.3, 1.3], 8)
pie("Tabla 2. Estimaciones MCO. Errores estándar robustos HC1 entre paréntesis. *** p < 0.01, ** p < 0.05, * p < 0.10.")

P3 = M3["params"]
wbar = D["wage"]["mean"]
par(f"**Escolaridad.** En el modelo preferido (3), comparando hombres con igual experiencia, antigüedad, edad, IQ, KWW, "
    f"escolaridad de los padres, estado civil, raza y región, un año más de escolaridad se asocia con un salario "
    f"{100*b3:.1f}% mayor (exacto {pct(b3):.1f}%), unos {wbar*pct(b3)/100:.0f} USD mensuales en el salario promedio de "
    f"{wbar:.0f} USD. El IC 95% robusto es [{ci3[0]:.4f}, {ci3[1]:.4f}], así que se rechaza H₀ al 5% "
    f"(t = {b3/M3['se']['educ']:.2f}, p {p3(M3['p']['educ'])}). Es económicamente relevante: cuatro años de universidad se "
    f"asocian con ≈{pct(4*b3):.0f}% más salario. La asociación cae de {100*b1:.1f}% (1) y {100*M2['params']['educ']:.1f}% (2) "
    f"a {100*b3:.1f}% (3). Esa caída es coherente con el **sesgo por variable omitida** de habilidad y entorno familiar, "
    "no con la multicolinealidad, que afecta la varianza y no el valor esperado de los estimadores.")
otros = []
for k, lab, esperado in [("tenure", "Antigüedad", "+"), ("IQ", "IQ", "+"), ("KWW", "KWW", "+"), ("married", "Casado", "+"),
                         ("black", "Afroamericano", "−"), ("south", "Sur", "−"), ("urban", "Urbano", "+")]:
    b, pv = P3[k], M3["p"][k]
    signo = "+" if b > 0 else "−"
    if k in ("married", "black", "south", "urban"):
        efecto = f"{pct(b):+.1f}% (exacto)"
    elif k in ("IQ", "KWW"):
        efecto = f"{100*b:+.2f}% por punto ({100*b*R['descriptivos'][k]['std']:+.1f}% por 1 desv. est.)"
    else:
        efecto = f"{100*b:+.2f}% por año"
    ok = "coincide con el signo esperado" if signo == esperado else "signo distinto al esperado"
    sig = "significativo al 5%" if pv < .05 else f"no significativo al 5% (p = {pv:.2f})"
    otros.append(f"**{lab}:** {efecto}; {sig}; {ok}.")
ex, ex2 = P3["exper_c"], P3["I(exper_c ** 2)"]
otros.insert(0, f"**Experiencia:** en la experiencia media ({D['exper']['mean']:.1f} años), un año más se asocia con "
                f"{100*ex:+.2f}% de salario (p = {M3['p']['exper_c']:.3f}). El término cuadrático tiene el signo esperado "
                f"(β₃ = {ex2:.5f}) pero es prácticamente cero y no significativo (p = {M3['p']['I(exper_c ** 2)']:.2f}): en esta "
                "cohorte joven (28–38 años) el perfil de ingresos es casi lineal y la concavidad de Mincer no se detecta.")
bullets(otros)
pe = PF[1]
pf0 = PF[0]
par(f"**Pruebas conjuntas.** La F global del modelo (3) es {M3['F']:.1f} (p {p3(M3['Fp'])}). Las variables de habilidad y "
    f"familia (IQ, KWW, meduc, feduc) son conjuntamente significativas: F({pf0['q']}, {pf0['gl2']}) = {pf0['F']:.2f}, "
    f"p {p3(pf0['p'])}. Experiencia, su cuadrado y edad no son significativas una por una, pero sí en conjunto: "
    f"F({pe['q']}, {pe['gl2']}) = {pe['F']:.2f}, p = {pe['p']:.3f}. Es el síntoma clásico de regresores redundantes. "
    f"El R² de {M3['R2']:.2f} indica que el modelo explica una cuarta parte de la variación del log-salario; no mide "
    "causalidad ni calidad del modelo por sí solo.")

# ------------------------------------------------------------------ 6. Diagnóstico
doc.add_heading("6. Diagnóstico", 1)
doc.add_heading("6.1 Supuestos y residuos", 2)
tabla(["Supuesto", "Evidencia", "Decisión"], [
    ["Linealidad en parámetros", f"RESET (potencias 2 y 3): p = {R['reset']['p']:.2f}", "No hay evidencia de mala forma funcional"],
    ["Rango completo", "VIF finitos; se evita la trampa de las dummies", "Cumple (sin colinealidad perfecta)"],
    ["Exogeneidad E[u|X] = 0", "No es comprobable con pruebas: habilidad medida con error, selección", "Supuesto dudoso: interpretamos asociaciones"],
    ["Homoscedasticidad", f"Breusch–Pagan LM = n·R²aux = {H['LM']:.1f}, p = {H['bp_p']:.3f}; White p = {H['white_p']:.2f}",
     "Se reportan errores HC1"],
    ["Independencia", "Corte transversal con muestreo de personas", "Plausible; Durbin–Watson no aplica"],
    ["Normalidad", f"Jarque–Bera = {JB['JB']:.0f} (p {p3(JB['p'])}); curtosis {JB['K']:.1f}",
     f"Se rechaza, pero con n = {n} la inferencia es asintóticamente válida"]],
    [1.5, 2.9, 2.1])
pie("Tabla 3. Revisión de supuestos del modelo preferido.")
par(f"Los residuos no muestran curvatura ni un abanico marcado (Figura 2). La cola izquierda es más pesada que la normal, y "
    f"ninguna observación domina la estimación (distancia de Cook máxima = {R['cook_max_red']:.3f}). El error estándar robusto "
    f"de educ ({M3['se']['educ']:.4f}) es apenas mayor que el clásico ({F_['bse']['educ']:.4f}), y WLS da β_educ = "
    f"{H['wls_educ']:.4f}: la heteroscedasticidad es leve y no cambia las conclusiones.")
figura("residuos", "Figura 2. Residuos vs ajustados, Q–Q, histograma y leverage–Cook del modelo preferido.", 4.3)

doc.add_heading("6.2 Multicolinealidad", 2)
vc_ = R["vif_centrado"]
filas = [[k.replace("I(exper_c ** 2)", "exper²").replace("exper_c", "exper"), f"{VF[k2]['VIF']:.2f}", f"{vc_[k]:.2f}",
          f"{1/vc_[k]:.3f}"] for k, k2 in
         [("exper_c", "exper"), ("I(exper_c ** 2)", "exper²"), ("educ", "educ"), ("age", "age"), ("KWW", "KWW"), ("IQ", "IQ"),
          ("feduc", "feduc"), ("meduc", "meduc"), ("black", "black"), ("tenure", "tenure")]]
tabla(["Regresor", "VIF (exper sin centrar)", "VIF modelo preferido", "Tolerancia"], filas, [1.4, 1.7, 1.7, 1.2])
pie("Tabla 4. Factores de inflación de varianza. Las dummies restantes tienen VIF < 1.1.")
se_b, se_p = R["base"]["bse"]["educ"], F_["bse"]["educ"]
CR, CP = C["Reducido"], C["PCA"]
par(f"**Lectura sin reglas mecánicas.** Sin centrar, exper y exper² tienen VIF de {VF['exper']['VIF']:.0f} y "
    f"{VF['exper²']['VIF']:.0f}, por encima de los umbrales de 5 o 10. Esa colinealidad es **estructural**: exper² se "
    f"construye a partir de exper (r = {R['corr_exper_exper2']:.3f}). Al centrar, la correlación baja a "
    f"{R['corr_experc']:.3f}, el VIF máximo a {R['vif_centrado_max']:.2f}, y el R² y el coeficiente de educ no cambian. Por "
    "eso no era un problema de los datos. La multicolinealidad **sustantiva** está en el bloque escolaridad–habilidad–"
    f"familia: VIF moderados (1.6–2.2), con VIF de educ = {VF['educ']['VIF']:.2f}. Ese VIF significa que el error estándar "
    f"de educ es √{VF['educ']['VIF']:.2f} ≈ {np.sqrt(VF['educ']['VIF']):.2f} veces el que tendría con regresores no "
    f"correlacionados. En la práctica, pasa de {se_b:.4f} (modelo 2) a {se_p:.4f} (modelo 3, errores clásicos).")
par(f"**¿Dificulta responder la pregunta?** Poco. Aun con el error estándar inflado, el IC 95% de β_educ excluye el cero "
    f"con holgura. Donde sí estorba es en separar IQ, KWW, meduc y feduc entre sí: varias no son significativas una por "
    f"una, pero la prueba F conjunta es contundente. Los diagnósticos globales coinciden: número de condición "
    f"(Belsley–Kuh–Welsch) = {R['num_cond_bkw']:.0f} sin centrar y {R['num_cond_estandarizado']:.0f} con regresores "
    f"estandarizados.")
par(f"**Qué hicimos y qué no hicimos.** No eliminamos controles para bajar un VIF, porque todos se justifican por teoría. "
    f"Centramos la experiencia, que no cambia la pregunta. Como análisis de **sensibilidad**, probamos especificaciones "
    f"alternativas. Si se elimina age y KWW y se promedia la escolaridad de los padres, β_educ = {CR['beta_educ']:.4f}. "
    f"Si se resumen IQ, KWW, meduc y feduc en su primer componente principal, β_educ = {CP['beta_educ']:.4f}. Ambos quedan "
    "dentro del IC del modelo preferido (Figura 3), pero cambian la pregunta, porque ya no controlan por todos los "
    "factores por separado. Por eso no son nuestra especificación principal.")
figura("educ", "Figura 3. Coeficiente de educ (IC 95%) según la especificación.", 4.0)

# ------------------------------------------------------------------ 7. Conclusiones
doc.add_heading("7. Conclusiones y limitaciones", 1)
par(f"**Respuesta a la pregunta.** Entre hombres trabajadores de EE. UU. en 1980 con igual experiencia, antigüedad, "
    f"habilidad medida y antecedentes familiares, un año más de escolaridad se asocia con un salario ≈{pct(b3):.0f}% mayor "
    f"(IC 95%: {100*ci3[0]:.1f}–{100*ci3[1]:.1f} log-puntos). Es menos de la mitad que en la regresión simple "
    f"({pct(b1):.1f}%). La multicolinealidad entre escolaridad, habilidad y familia es moderada: infla la varianza, pero no "
    "impide medir la asociación con una precisión útil. Los VIF extremos de la experiencia eran un artificio del "
    "polinomio y se corrigen al centrar.")
par("**Limitaciones.**")
bullets([
    "**Causalidad:** la regresión es observacional. IQ y KWW miden la habilidad con error, así que puede quedar sesgo por "
    "variable omitida (probablemente hacia arriba). La escolaridad también puede tener error de medición (sesgo hacia "
    "cero). No afirmamos que estudiar un año más cause un aumento de 4%.",
    f"**Selección:** solo se observan hombres con empleo y salario, y eliminamos {n0 - n} observaciones sin escolaridad de "
    "los padres. Si esas personas difieren sistemáticamente, la muestra deja de ser representativa.",
    "**Validez externa:** hombres de 28–38 años en EE. UU. en 1980; los resultados no se trasladan directamente a "
    "mujeres, a otras cohortes ni a México.",
    "**Forma funcional:** suponemos el mismo retorno para todos. El anexo explora interacciones y regresión cuantílica, "
    "y encuentra heterogeneidad moderada.",
])
par(f"**¿Qué decisión de modelado podría cambiar nuestra conclusión?** Sobre todo, incluir o no los controles de "
    f"habilidad: β_educ pasa de {b1:.3f} a {b3:.3f}. Con una mejor medida de habilidad o una estrategia causal "
    "(variables instrumentales o gemelos), la cifra podría ser distinta. En cambio, eliminar controles para bajar los VIF "
    "o combinar indicadores cambia β_educ menos de un punto porcentual.")

doc.add_heading("Referencias", 1)
for ref in ["Belsley, D., Kuh, E. y Welsch, R. (1980). Regression Diagnostics. Wiley.",
            "Blackburn, M. y Neumark, D. (1992). Unobserved ability, efficiency wages, and interindustry wage differentials. "
            "Quarterly Journal of Economics, 107(4), 1421–1436.",
            "Galindo, C. (2026). Notas de regresión lineal múltiple, inferencia y diagnóstico; Proyecto de Econometría: guía de "
            "trabajo. Universidad Panamericana.",
            "Mincer, J. (1974). Schooling, Experience, and Earnings. NBER.",
            "Wooldridge, J. M. (2020). Introducción a la econometría (7.ª ed.). Cengage. Datos: paquete wooldridge 0.5.0 (PyPI), "
            "consultado el 18/11/2026."]:
    p = doc.add_paragraph(ref)
    p.paragraph_format.space_after = Pt(1)
    p.runs[0].font.size = Pt(9)
par("**Anexo reproducible:** el código (scripts/analisis.py y multicolinealidad_mincer.ipynb), los pasos para "
    "reproducir el análisis (README) y el anexo técnico con los demás temas del curso están en "
    "github.com/ximeelaraa/mincer-multicolinealidad.")

doc.save(ROOT / "Informe_Proyecto_Econometria.docx")
print("Listo: Informe_Proyecto_Econometria.docx")
