"""Guía de estudio y exposición (lenguaje simple) para el equipo.

Uso: .venv/bin/python scripts/generar_guia.py   (después de analisis.py y generar_presentacion.py)
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
G = json.loads((FIG / "_guion.json").read_text())
TP = R["tabla_principal"]
M1, M2, M3 = TP["(1) Simple"], TP["(2) Mincer"], TP["(3) Preferido"]
VF, PF, C, H, JB, D = R["vif_full"], R["pruebas_F"], R["comparacion"], R["hetero"], R["jb"], R["descriptivos"]
n, n0 = R["n"], R["n_original"]
pct = lambda b: 100 * (np.exp(b) - 1)
e1, e2, e3 = (pct(M["params"]["educ"]) for M in (M1, M2, M3))
ci3 = [pct(x) for x in M3["ci"]["educ"]]
usd = D["wage"]["mean"] * e3 / 100
TINTA = RGBColor(0x1D, 0x24, 0x33)

doc = Document()
st = doc.styles["Normal"]
st.font.name, st.font.size = "Calibri", Pt(11)
st.paragraph_format.space_after = Pt(5)
for s_ in doc.sections:
    s_.left_margin = s_.right_margin = Inches(0.95)
for lvl, size in [(1, 15), (2, 12)]:
    h = doc.styles[f"Heading {lvl}"]
    h.font.size, h.font.color.rgb = Pt(size), TINTA


def par(txt, size=None, italic=False):
    p = doc.add_paragraph()
    for i, t in enumerate(txt.split("**")):
        r_ = p.add_run(t)
        r_.bold = i % 2 == 1
        r_.italic = italic
        if size:
            r_.font.size = Pt(size)
    return p


def lista(items, estilo="List Bullet"):
    for it in items:
        p = doc.add_paragraph(style=estilo)
        for i, t in enumerate(it.split("**")):
            p.add_run(t).bold = i % 2 == 1
        p.paragraph_format.space_after = Pt(3)


def tabla(filas, anchos, size=10):
    t = doc.add_table(rows=0, cols=len(filas[0]))
    t.style = "Light Grid Accent 1"
    for f in filas:
        cells = t.add_row().cells
        for j, v in enumerate(f):
            cells[j].text = str(v)
            cells[j].width = Inches(anchos[j])
            for p in cells[j].paragraphs:
                p.paragraph_format.space_after = Pt(0)
                for r_ in p.runs:
                    r_.font.size = Pt(size)
    doc.add_paragraph()


def recuadro(txt):
    t = doc.add_table(rows=1, cols=1)
    t.style = "Light Shading Accent 1"
    c = t.rows[0].cells[0]
    c.text = ""
    p = c.paragraphs[0]
    for i, s in enumerate(txt.split("**")):
        r_ = p.add_run(s)
        r_.bold = i % 2 == 1
        r_.font.size = Pt(11)
    doc.add_paragraph()


# ------------------------------------------------------------------ portada
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_ = p.add_run("Guía para entender y exponer el proyecto")
r_.bold, r_.font.size, r_.font.color.rgb = True, Pt(22), TINTA
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run("Escolaridad, habilidad y salario: ¿la multicolinealidad impide medir la asociación?").font.size = Pt(13)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run("Equipo 4: Ximena Lara, Camila Ricardi y Sara Enriquez · Fundamentos de Econometría").font.size = Pt(11)
par("Este documento es para estudiar, no para entregar. Explica todo en palabras simples, dice qué decir en cada "
    "diapositiva y qué responder si el profesor pregunta.", 10.5, True)

# ------------------------------------------------------------------ 0. qué presentar
doc.add_heading("0. ¿Qué presentamos y qué entregamos?", 1)
tabla([["Para qué", "Archivo", "Comentario"],
       ["Exponer (8 min)", "Presentacion_Beamer.pdf", "11 diapositivas. Es la que proyectan."],
       ["Lo mismo, editable", "Presentacion_Final.pptx (y .pdf)",
        "Idéntica a la Beamer. Úsenla si quieren cambiar algo; las notas del orador ya traen el guion."],
       ["Entregar: informe", "Informe_Proyecto_Econometria.pdf", "7 páginas (la guía pide 5–8)."],
       ["Entregar: anexo reproducible", "multicolinealidad_mincer.ipynb + carpeta scripts/", "El código; corre en Colab."],
       ["Entregar: datos documentados", "carpeta data/", "La base usada y su diccionario."],
       ["Extra (no se expone)", "Anexo_Tecnico_Temas_del_Curso.pdf", "Los demás temas del parcial con los mismos datos."],
       ["Para estudiar", "Esta guía", "No se entrega."]], [1.6, 2.5, 2.6])

# ------------------------------------------------------------------ 1. en un minuto
doc.add_heading("1. El proyecto en un minuto", 1)
recuadro(f"**Pregunta:** ¿cuánto más gana alguien con un año más de escuela, comparado con alguien igual en experiencia, "
         f"habilidad y familia?\n\n**Respuesta:** ≈ **{e3:.1f} % más** de salario (unos {usd:.0f} dólares al mes). Sin "
         f"comparar personas parecidas salía {e1:.1f} %. La diferencia se debe a que la gente que estudia más también suele "
         "tener más habilidad.\n\n**¿Y la multicolinealidad?** Las variables de control se parecen entre sí. Eso hace "
         "las estimaciones menos precisas, pero no tanto como para impedir responder la pregunta.")

doc.add_heading("La historia en 5 pasos", 2)
lista([
    f"**Con solo escolaridad**, un año más de escuela se asocia con {e1:.1f} % más salario.",
    f"**Al agregar experiencia sube a {e2:.1f} %.** Quien estudia más empezó a trabajar más tarde, así que tiene menos "
    "experiencia. Sin controlar la experiencia, ese efecto escondía parte del valor de la escuela.",
    f"**Al agregar habilidad (IQ, KWW) y familia (escolaridad de los padres) baja a {e3:.1f} %.** Parte de lo que parecía "
    "«escuela» en realidad era habilidad. A esto se le llama **sesgo por variable omitida**.",
    "**Multicolinealidad:** como escolaridad, IQ, KWW y los padres se mueven juntos, al modelo le cuesta separarlos. "
    "Eso agranda los errores estándar, pero el resultado de escolaridad sigue siendo claro.",
    "**Conclusión:** la relación es positiva, importante y robusta. Es una asociación, no prueba que la escuela «cause» "
    "el aumento."], "List Number")

# ------------------------------------------------------------------ 2. conceptos
doc.add_heading("2. Los conceptos, explicados fácil", 1)
par("Cada concepto con su definición sencilla y cómo aparece en nuestro proyecto.")
conceptos = [
    ("Regresión lineal", "Una recta (o plano) que resume cómo cambia Y cuando cambian las X.",
     "Y = logaritmo del salario; X = escolaridad, experiencia, IQ, etc."),
    ("MCO (mínimos cuadrados ordinarios)", "El método que elige la recta que deja los errores más pequeños (suma de errores al "
     "cuadrado mínima).", "Así estimamos todos los coeficientes."),
    ("Coeficiente (β)", "Cuánto cambia Y cuando X sube en 1, manteniendo todo lo demás igual.",
     f"β de escolaridad = {M3['params']['educ']:.3f}."),
    ("¿Por qué logaritmo del salario?", "Con log(Y), el coeficiente × 100 se lee como un cambio porcentual. Además el salario "
     "tiene unos pocos valores muy altos y el log los «aplana».", f"0.041 × 100 ≈ 4.1 % (exacto: {e3:.1f} %)."),
    ("«Manteniendo lo demás constante»", "Comparamos personas iguales en todas las demás variables del modelo. Es la frase "
     "más importante al interpretar.", "Dos hombres con la misma experiencia, IQ, padres… pero uno con un año más de escuela."),
    ("Error estándar", "Qué tan precisa es la estimación: entre más chico, más precisa.", f"{M3['se']['educ']:.4f} para escolaridad."),
    ("Intervalo de confianza 95 %", "Rango de valores compatibles con los datos. Si no incluye el 0, el efecto es "
     "significativo al 5 %.", f"De {ci3[0]:.1f} % a {ci3[1]:.1f} %: no incluye el 0."),
    ("Valor p", "Qué tan raros serían nuestros datos si en realidad β fuera 0. Si p < 0.05, rechazamos H₀. NO es la "
     "probabilidad de que H₀ sea cierta.", "p < 0.001 para escolaridad."),
    ("Hipótesis", "H₀ dice «no hay relación» (β = 0); H₁ dice «sí hay, y es positiva» (β > 0).", "Rechazamos H₀."),
    ("α = 5 %", "El nivel de significancia: el umbral que elegimos ANTES de ver los resultados.", "Lo fijamos al inicio."),
    ("R²", "Qué fracción de las diferencias en salario explica el modelo. No mide si el modelo es «bueno» ni si hay causalidad.",
     f"{M3['R2']:.2f}: explica una cuarta parte. Es normal en salarios: influyen muchas cosas no medidas."),
    ("Variable dummy", "Vale 1 o 0 (sí/no). Su coeficiente es la diferencia contra el grupo de referencia.",
     f"Casado: {pct(M3['params']['married']):+.0f} % frente a no casado."),
    ("Grupo de referencia y trampa de las dummies", "Con 2 categorías se pone UNA dummy; la otra es la referencia. Si se ponen "
     "las dos junto con la constante, hay multicolinealidad perfecta y el modelo no se puede estimar.",
     "Referencia: no casado, no afroamericano, fuera del sur, no urbano."),
    ("Sesgo por variable omitida", "Si dejas fuera algo que afecta a Y y está relacionado con X, el coeficiente de X se lleva "
     "un crédito que no le toca.", f"Sin habilidad, escolaridad «se llevaba» parte del efecto: {e2:.1f} % → {e3:.1f} %."),
    ("Multicolinealidad", "Cuando dos o más X se parecen mucho entre sí. Perfecta: una es combinación exacta de otras y no se "
     "puede estimar. Imperfecta: sí se estima, pero con más incertidumbre. NO sesga los coeficientes.",
     "Escolaridad, IQ, KWW y los padres están correlacionados."),
    ("VIF (factor de inflación de la varianza)", "Cuántas veces se agranda la varianza de un coeficiente por la correlación con "
     "las otras X. VIF = 1/(1 − R²ⱼ). Un VIF de 5 o 10 es una alerta, no una regla automática.",
     f"Escolaridad: VIF = {VF['educ']['VIF']:.2f} → su error estándar es √{VF['educ']['VIF']:.2f} ≈ "
     f"{np.sqrt(VF['educ']['VIF']):.1f} veces más grande."),
    ("Centrar una variable", "Restarle su promedio. No cambia el modelo, pero quita la correlación artificial entre una "
     "variable y su cuadrado.", f"exper y exper² tenían VIF de {VF['exper²']['VIF']:.0f}; al centrar bajó a "
     f"{R['vif_centrado_max']:.1f}."),
    ("Prueba F conjunta", "Pregunta si VARIAS variables juntas importan. Sirve cuando por separado no salen significativas por "
     "la multicolinealidad.", f"IQ, KWW y padres juntos: F = {PF[0]['F']:.1f}, p < 0.001."),
    ("Heteroscedasticidad", "Cuando la dispersión de los errores no es igual para todos. No sesga los β, pero los errores "
     "estándar normales quedan mal. Solución: errores robustos (HC1).",
     f"Breusch–Pagan p = {H['bp_p']:.3f} → la hay, leve → usamos HC1."),
    ("Normalidad (Jarque–Bera)", "Revisa si los errores tienen forma de campana. Con muestras grandes no es grave.",
     f"Se rechaza (JB = {JB['JB']:.0f}), pero con n = {n} la inferencia sigue siendo válida."),
    ("RESET", "Revisa si falta alguna forma no lineal en el modelo.", f"p = {R['reset']['p']:.2f} → la forma funcional está bien."),
    ("Exogeneidad", "El supuesto de que el error no está relacionado con las X. Es el supuesto que permite hablar de causalidad, "
     "y no se puede comprobar con una prueba.", "La habilidad se mide con error → hablamos de asociación."),
    ("Asociación vs causalidad", "Asociación: «van juntos». Causalidad: «uno provoca el otro». Con datos observacionales "
     "solo podemos afirmar asociación.", "Decimos «se asocia con», nunca «causa» o «provoca»."),
]
tabla([["Concepto", "Qué es (simple)", "En nuestro proyecto"]] + [list(c) for c in conceptos], [1.55, 2.95, 2.2], 9.5)

# ------------------------------------------------------------------ 3. números
doc.add_heading("3. Los números que hay que saberse", 1)
tabla([["Dato", "Valor", "Cómo decirlo"],
       ["Observaciones", f"{n} (de {n0})", f"Quitamos {n0 - n} sin escolaridad de los padres"],
       ["Escolaridad, sin controles", f"{e1:.1f} %", "Un año más de escuela, sin comparar personas parecidas"],
       ["Escolaridad, con experiencia", f"{e2:.1f} %", "Sube porque quien estudia más tiene menos experiencia"],
       ["Escolaridad, modelo final", f"{e3:.1f} %", f"≈ {usd:.0f} USD al mes; con 4 años de universidad ≈ {pct(4*M3['params']['educ']):.0f} %"],
       ["Intervalo de confianza 95 %", f"{ci3[0]:.1f} % a {ci3[1]:.1f} %", "No incluye 0 → significativo"],
       ["R² del modelo final", f"{M3['R2']:.2f}", "Explica una cuarta parte de la variación del salario"],
       ["VIF exper y exper² (sin centrar)", f"{VF['exper']['VIF']:.0f} y {VF['exper²']['VIF']:.0f}", "Artificial: una es el cuadrado de la otra"],
       ["VIF máximo al centrar", f"{R['vif_centrado_max']:.1f}", "Multicolinealidad real: moderada"],
       ["VIF de escolaridad", f"{VF['educ']['VIF']:.2f}", f"Error estándar ≈ {np.sqrt(VF['educ']['VIF']):.1f} veces más grande"],
       ["F de IQ, KWW y padres", f"{PF[0]['F']:.1f} (p < 0.001)", "Juntos sí importan"],
       ["Sensibilidad", f"{pct(C['Reducido']['beta_educ']):.1f} % y {pct(C['PCA']['beta_educ']):.1f} %",
        "Quitando o combinando controles: casi igual"],
       ["Breusch–Pagan", f"p = {H['bp_p']:.3f}", "Heteroscedasticidad leve → errores robustos"]], [2.2, 1.5, 3.0])

# ------------------------------------------------------------------ 4. guion
doc.add_heading("4. Guion de la exposición (8 minutos)", 1)
par("Propuesta de reparto: cada una habla unos 2 min 40 s. Pueden cambiarlo, pero practiquen con cronómetro. No lean: "
    "usen el guion como idea y díganlo con sus palabras.")
diapos = [("1. Portada", None, "Ximena", "Presentar al equipo y el título (10 s)."),
          ("2. La pregunta", "pregunta", None, None), ("3. Datos", "datos", None, None), ("4. El modelo", "modelo", None, None),
          ("5. Resultado principal", "resultado", None, None), ("6. ¿Qué significa?", "interpretacion", None, None),
          ("7. Multicolinealidad: ¿dónde está?", "multi", None, None), ("8. ¿Nos impide responder?", "impide", None, None),
          ("9. ¿Se cumplen los supuestos?", "diagnostico", None, None), ("10. Conclusiones", "conclusion", None, None),
          ("11. Gracias · ¿preguntas?", "cierre", None, None)]
tiempos = {"1": "0:10", "2": "0:50", "3": "0:45", "4": "0:55", "5": "1:00", "6": "0:50", "7": "0:50", "8": "0:55",
           "9": "0:40", "10": "0:45", "11": "0:10"}
filas = [["Diapositiva", "Quién", "Tiempo", "Qué decir"]]
for titulo, clave, quien, texto in diapos:
    num = titulo.split(".")[0]
    filas.append([titulo, quien or G["quien"][clave], tiempos[num], texto or G["guion"][clave]])
tabla(filas, [1.5, 0.7, 0.6, 3.9], 9.5)

# ------------------------------------------------------------------ 5. preguntas
doc.add_heading("5. Preguntas que les pueden hacer (y cómo responder)", 1)
qa = [
    ("¿Entonces estudiar un año más causa ganar 4 % más?",
     "No podemos afirmarlo. Es una asociación: comparamos personas parecidas, pero pueden quedar diferencias no medidas (la "
     "habilidad la medimos solo con IQ y KWW). Para hablar de causalidad necesitaríamos un experimento o variables "
     "instrumentales."),
    ("Si el VIF de experiencia era 24, ¿por qué no quitaron la variable?",
     "Porque era colinealidad artificial: exper² se construye a partir de exper. Al centrar la experiencia (restarle su "
     f"promedio) el VIF bajó a {R['vif_centrado_max']:.1f} sin cambiar el modelo. Quitar variables solo para bajar un VIF "
     "cambia la pregunta, y la guía del proyecto dice que no se haga."),
    ("¿La multicolinealidad sesga los coeficientes?",
     "No. Si se cumple la exogeneidad, MCO sigue siendo insesgado; la multicolinealidad solo agranda las varianzas. La caída "
     f"de {e2:.1f} % a {e3:.1f} % NO es por multicolinealidad, sino por sesgo por variable omitida."),
    ("¿Por qué IQ, KWW y los padres no son significativos y aun así los dejan?",
     f"Porque juntos sí lo son (F = {PF[0]['F']:.1f}, p < 0.001). Individualmente no salen porque se parecen entre sí: es el "
     "síntoma típico de multicolinealidad. Además, la teoría dice que deben estar para comparar personas con igual habilidad."),
    ("¿No es muy bajo un R² de 0.26?",
     "En datos de salarios por persona es normal: el salario depende de muchas cosas que no medimos (suerte, contactos, "
     "motivación). El R² no mide si el modelo es correcto ni si hay causalidad; lo importante es que el coeficiente de "
     "escolaridad sea preciso."),
    ("¿Qué hicieron con la heteroscedasticidad?",
     f"Breusch–Pagan la detectó (p = {H['bp_p']:.3f}), así que reportamos errores estándar robustos HC1. Cambian muy poco y "
     "no modifican la conclusión."),
    ("Los residuos no son normales, ¿eso invalida todo?",
     f"No. Con {n} observaciones, por el teorema del límite central las pruebas t y F siguen siendo válidas (son "
     "aproximadamente correctas)."),
    ("¿Por qué usan logaritmo del salario?",
     "Para interpretar los coeficientes como porcentajes y porque el salario está muy sesgado a la derecha (pocos salarios "
     "muy altos)."),
    ("¿Por qué el coeficiente de experiencia al cuadrado sale casi cero?",
     "Porque son hombres jóvenes (28–38 años): a esa edad el salario todavía sube casi en línea recta con la experiencia. La "
     "curva que predice Mincer se ve más en edades mayores."),
    ("¿De dónde son los datos? ¿Son confiables?",
     "De la National Longitudinal Survey de EE. UU. (1980), la usaron Blackburn y Neumark (1992) y viene con el libro de "
     "Wooldridge. Es pública y se usa mucho en econometría. Todo está documentado en data/DICCIONARIO.md."),
    ("¿Por qué quitaron 213 observaciones? ¿No sesga eso?",
     "Porque les faltaba la escolaridad de los padres. Puede introducir selección si esas personas son distintas; lo "
     "mencionamos como limitación."),
    ("¿Qué decisión de modelado cambiaría su conclusión?",
     f"Sobre todo incluir o no la habilidad: pasa de {e1:.1f} % a {e3:.1f} %. Quitar o combinar controles casi no lo mueve "
     f"({pct(C['Reducido']['beta_educ']):.1f} % y {pct(C['PCA']['beta_educ']):.1f} %)."),
    ("¿Sirve para México?",
     "No directamente: son hombres de EE. UU. en 1980. Habría que repetir el análisis con datos mexicanos, por ejemplo la "
     "ENOE del INEGI."),
]
for q, a in qa:
    p = doc.add_paragraph()
    r_ = p.add_run("P: " + q)
    r_.bold = True
    p.paragraph_format.space_after = Pt(1)
    par("R: " + a)

# ------------------------------------------------------------------ 6. qué no decir
doc.add_heading("6. Frases que NO hay que decir", 1)
tabla([["No digan", "Mejor digan"],
       ["«Estudiar causa que ganes 4 % más»", "«Un año más de escuela se asocia con 4 % más salario»"],
       ["«El VIF era alto, así que quitamos la variable»", "«El VIF alto era artificial; lo resolvimos centrando»"],
       ["«La multicolinealidad sesga los coeficientes»", "«La multicolinealidad agranda las varianzas, no sesga»"],
       ["«p = 0.03 significa 3 % de probabilidad de que H₀ sea cierta»", "«Con α = 5 % rechazamos H₀»"],
       ["«El R² es bajo, el modelo está mal»", "«El R² no mide si el modelo es correcto»"],
       ["«Es significativo, entonces es importante»", "«Es significativo y además grande: 18 % con 4 años de universidad»"]],
      [3.3, 3.4])

# ------------------------------------------------------------------ 7. rúbrica
doc.add_heading("7. Cómo nos califican la exposición (20 puntos)", 1)
tabla([["Criterio", "Puntos", "Dónde lo cubrimos"],
       ["Pregunta, motivación y datos claros", "4", "Diapositivas 2 y 3"],
       ["Modelo y coeficientes bien interpretados", "5", "Diapositivas 4, 5 y 6"],
       ["Diagnósticos, multicolinealidad y límites", "5", "Diapositivas 7, 8, 9 y 10"],
       ["Estructura, apoyos legibles y tiempo", "3", "11 diapositivas, 8 minutos: ¡ensayen con cronómetro!"],
       ["Respuestas razonadas", "3", "Sección 5 de esta guía"]], [3.0, 0.8, 2.9])
par("Recuerden también: la guía pide que **las tres puedan explicar todo el modelo**, no solo su parte, y llenar la tabla "
    "de contribuciones del README.")

doc.save(ROOT / "Guia_de_Estudio_y_Exposicion.docx")
print("Listo: Guia_de_Estudio_y_Exposicion.docx")
