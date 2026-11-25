"""Presentación final para exponer (8 minutos): PowerPoint y Beamer con el mismo contenido.

Uso: .venv/bin/python scripts/generar_presentacion.py   (requiere resultados.json de scripts/analisis.py)
"""
import json
import zipfile
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
FIG = ROOT / "figuras"
R = json.loads((ROOT / "resultados.json").read_text())
EQUIPO = "Ximena Lara · Camila Ricardi · Sara Enriquez"
TP = R["tabla_principal"]
M1, M2, M3 = TP["(1) Simple"], TP["(2) Mincer"], TP["(3) Preferido"]
VF, PF, C, H, JB, D = R["vif_full"], R["pruebas_F"], R["comparacion"], R["hetero"], R["jb"], R["descriptivos"]
n, n0 = R["n"], R["n_original"]
pct = lambda b: 100 * (np.exp(b) - 1)

b1, b2, b3 = (M["params"]["educ"] for M in (M1, M2, M3))
e1, e2, e3 = pct(b1), pct(b2), pct(b3)
ci3 = [pct(x) for x in M3["ci"]["educ"]]
usd = D["wage"]["mean"] * e3 / 100
cuatro = pct(4 * b3)
se_ratio = np.sqrt(VF["educ"]["VIF"])
lo, hi = sorted([pct(C["Reducido"]["beta_educ"]), pct(C["PCA"]["beta_educ"])])

# ------------------------------------------------------------------ paleta (rosa)
ROSA_H, ROSACL_H, BLUSH_H, CIRUELA_H, LILA_H, TEXTO_H, GRIS_H = (
    "#D6457A", "#FCE4EC", "#FFF7FA", "#5E2750", "#8E5CC4", "#4A3B47", "#9A8C98")
MALVA_H = "#B9A7B4"  # puntos de referencia en las gráficas
plt.rcParams["font.family"] = ["Avenir Next", "Avenir", "DejaVu Sans"]

# ------------------------------------------------------------------ figura: el número clave en tres pasos
fig, ax = plt.subplots(figsize=(7.2, 4.2), dpi=200)
etiquetas = ["Solo escolaridad", "+ experiencia", "+ habilidad, familia\ny demás controles"]
vals = [e1, e2, e3]
cis = [[pct(x) for x in M["ci"]["educ"]] for M in (M1, M2, M3)]
x = np.arange(3)
for i in range(3):
    col = ROSA_H if i == 2 else MALVA_H
    ax.plot([x[i], x[i]], cis[i], color=col, lw=2.2, solid_capstyle="round")
    ax.scatter(x[i], vals[i], s=110 if i == 2 else 85, color=col, zorder=3, edgecolor="white", linewidth=2)
    ax.text(x[i] + .1, vals[i], f"{vals[i]:.1f}%", va="center", fontsize=13, color=CIRUELA_H,
            fontweight="bold" if i == 2 else "normal")
ax.set_xticks(x, etiquetas, fontsize=10, color=CIRUELA_H)
ax.set_ylabel("% más salario por un año más de escuela", fontsize=9.5, color=TEXTO_H)
ax.set_ylim(0, 10.5)
ax.set_xlim(-.4, 2.6)
ax.grid(axis="y", color="#F3E3EA", lw=.8)
for s in ("top", "right", "left"):
    ax.spines[s].set_visible(False)
ax.spines["bottom"].set_color("#E8CFDA")
ax.tick_params(axis="y", colors=GRIS_H, length=0)
ax.tick_params(axis="x", length=0)
ax.set_axisbelow(True)
fig.text(.01, .01, "Líneas: intervalo de confianza al 95 % (errores robustos HC1)", fontsize=8, color=GRIS_H)
fig.tight_layout(rect=(0, .03, 1, 1))
fig.savefig(FIG / "pasos_educ.png", facecolor="white")
plt.close(fig)

# ------------------------------------------------------------------ figura: VIF antes y después de centrar
NOMBRES = {"educ": "escolaridad", "exper": "experiencia", "exper²": "experiencia²", "tenure": "antigüedad", "age": "edad",
           "IQ": "IQ", "KWW": "KWW", "meduc": "escolaridad madre", "feduc": "escolaridad padre", "married": "casado",
           "black": "afroamericano", "south": "sur", "urban": "urbano"}
CENT = {"exper": "exper_c", "exper²": "I(exper_c ** 2)"}
orden = sorted(NOMBRES, key=lambda k: VF[k]["VIF"])
fig, ax = plt.subplots(figsize=(7.2, 4.6), dpi=200)
for i, k in enumerate(orden):
    antes, despues = VF[k]["VIF"], R["vif_centrado"][CENT.get(k, k)]
    ax.plot([despues, antes], [i, i], color="#EBD3DE", lw=3, solid_capstyle="round", zorder=1)
    ax.scatter(antes, i, s=60, color=MALVA_H, zorder=2, edgecolor="white", linewidth=1.5)
    ax.scatter(despues, i, s=60, color=ROSA_H, zorder=3, edgecolor="white", linewidth=1.5)
    if k in CENT:
        ax.text(antes + .7, i, f"{antes:.0f}", va="center", fontsize=10, color=TEXTO_H)
        ax.text(despues + .7, i + .42, f"{despues:.1f}", va="center", fontsize=10, color=CIRUELA_H, fontweight="bold")
ax.axvline(10, color=GRIS_H, lw=1, ls=(0, (3, 3)), zorder=0)
ax.text(10.3, -.6, "regla común: 10", fontsize=8.5, color=GRIS_H, va="center")
ax.set_yticks(range(len(orden)), [NOMBRES[k] for k in orden], fontsize=9.5, color=CIRUELA_H)
ax.set_xlabel("VIF (factor de inflación de la varianza)", fontsize=9.5, color=TEXTO_H)
ax.set_xlim(0, 27)
ax.set_ylim(-1, len(orden))
ax.grid(axis="x", color="#F3E3EA", lw=.8)
for s in ("top", "right", "left"):
    ax.spines[s].set_visible(False)
ax.spines["bottom"].set_color("#E8CFDA")
ax.tick_params(length=0, colors=GRIS_H)
ax.tick_params(axis="y", colors=CIRUELA_H)
ax.set_axisbelow(True)
ax.scatter([], [], s=60, color=MALVA_H, label="experiencia sin centrar")
ax.scatter([], [], s=60, color=ROSA_H, label="experiencia centrada")
ax.legend(loc="lower right", frameon=False, fontsize=9, labelcolor=TEXTO_H)
fig.tight_layout()
fig.savefig(FIG / "vif_rosa.png", facecolor="white")
plt.close(fig)

# ------------------------------------------------------------------ contenido (común a PowerPoint y Beamer)
# Cada diapositiva: título, quién la presenta, guion (notas del orador)
GUION = {
    "pregunta": "Quien estudia más gana más, pero también suele tener más habilidad y padres más educados. Queremos saber cuánto "
                "del salario se asocia con la escuela comparando personas parecidas en todo lo demás, y si la multicolinealidad "
                "entre esos controles nos impide medirlo. Fijamos alfa = 5 % antes de estimar. Es asociación, no causalidad.",
    "datos": f"Usamos la base WAGE2 de Wooldridge, que viene de una encuesta real de EE. UU. de 1980. Eran {n0} hombres; quitamos "
             f"{n0 - n} sin escolaridad de los padres y quedaron {n}. El salario es mensual en dólares de 1980 y lo usamos en "
             "logaritmo para leer los coeficientes como porcentajes.",
    "modelo": "Es una ecuación de Mincer: logaritmo del salario contra escolaridad y experiencia, más controles de habilidad "
              "(IQ y KWW), de familia (escolaridad de los padres) y dummies. Esperamos signo positivo en escolaridad. "
              "Usamos errores robustos HC1.",
    "resultado": f"Este es el resultado central. Solo con escolaridad: {e1:.1f} %. Al agregar experiencia sube a {e2:.1f} %, "
                 "porque quien estudia más tiene menos años trabajando. Al agregar habilidad y familia baja a "
                 f"{e3:.1f} %: parte de lo que parecía escuela era habilidad. Eso es sesgo por variable omitida.",
    "interpretacion": f"Comparando a dos hombres iguales en todo menos en un año de escuela, el que estudió más gana en promedio "
                      f"{e3:.1f} % más, unos {usd:.0f} dólares al mes. El intervalo va de {ci3[0]:.1f} a {ci3[1]:.1f} %, no "
                      f"incluye el cero: rechazamos H0. Y es relevante: cuatro años de universidad son {cuatro:.0f} % más.",
    "multi": f"La experiencia y su cuadrado tenían VIF de {VF['exper']['VIF']:.0f} y {VF['exper²']['VIF']:.0f}, pero era un "
             f"artificio: una variable es el cuadrado de la otra. Al centrar la experiencia, el VIF máximo bajó a "
             f"{R['vif_centrado_max']:.1f} sin cambiar nada más. La multicolinealidad real está entre escolaridad, IQ, KWW y los "
             "padres, y es moderada.",
    "impide": f"La multicolinealidad hace el error estándar de escolaridad {se_ratio:.2f} veces más grande, pero el intervalo "
              f"sigue lejos de cero: sí podemos responder. Donde estorba es para separar IQ, KWW y padres entre sí: por separado "
              f"no salen significativos, pero juntos sí (F = {PF[0]['F']:.1f}). No quitamos variables para bajar el VIF; al "
              "probarlo, el resultado casi no cambia.",
    "diagnostico": "Revisamos los supuestos. Hay heteroscedasticidad leve, por eso usamos errores robustos. Los residuos no son "
                   f"normales, pero con {n} observaciones no importa. La forma funcional está bien según RESET.",
    "conclusion": f"Un año más de escuela se asocia con {e3:.1f} % más salario entre personas comparables. La multicolinealidad "
                  "infla la varianza pero no impide responder. Limitaciones: no es causal, puede haber selección y solo aplica a "
                  "hombres de EE. UU. en 1980. Lo que más cambia la conclusión es incluir o no la habilidad.",
    "cierre": "Todo es reproducible: el código, los datos y el diccionario están en GitHub. Gracias, ¿preguntas?",
}
QUIEN = {"pregunta": "Ximena", "datos": "Ximena", "modelo": "Ximena", "resultado": "Camila", "interpretacion": "Camila",
         "multi": "Camila", "impide": "Sara", "diagnostico": "Sara", "conclusion": "Sara", "cierre": "Sara"}
# Disponibles para la guía de estudio
(ROOT / "figuras" / "_guion.json").write_text(json.dumps({"guion": GUION, "quien": QUIEN}, ensure_ascii=False, indent=1))

# ------------------------------------------------------------------ POWERPOINT
hexrgb = lambda h: RGBColor.from_string(h.lstrip("#"))
ROSA, ROSACL, BLUSH, CIRUELA, LILA, TXT, SUAVE = map(hexrgb, (ROSA_H, ROSACL_H, BLUSH_H, CIRUELA_H, LILA_H, TEXTO_H, GRIS_H))
BL, LILACL = RGBColor(255, 255, 255), hexrgb("#EDE3F7")
FUENTE = "Avenir Next"
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]
cnt = [0]


def caja(sl, x, y, w, h, texto, size=20, color=TXT, bold=False, centro=False):
    tf = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
    tf.word_wrap = True
    for i, ln in enumerate(texto if isinstance(texto, list) else [texto]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text, p.font.size, p.font.color.rgb, p.font.bold, p.font.name = ln, Pt(size), color, bold, FUENTE
        p.space_after = Pt(10)
        if centro:
            p.alignment = 2
    return tf


def forma(sl, tipo, x, y, w, h, relleno=None, linea=None, grosor=1.25, radio=None):
    s = sl.shapes.add_shape(tipo, Inches(x), Inches(y), Inches(w), Inches(h))
    if relleno is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = relleno
    if linea is None:
        s.line.fill.background()
    else:
        s.line.color.rgb, s.line.width = linea, Pt(grosor)
    if radio is not None:
        s.adjustments[0] = radio
    s.shadow.inherit = False
    return s


def fondo(sl, portada=False):
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = BLUSH
    if portada:
        forma(sl, MSO_SHAPE.OVAL, 9.6, -2.2, 5.6, 5.6, ROSACL)
        forma(sl, MSO_SHAPE.OVAL, -1.6, 5.1, 3.6, 3.6, ROSACL)
        forma(sl, MSO_SHAPE.OVAL, 11.0, 3.4, .55, .55, LILACL)
        forma(sl, MSO_SHAPE.OVAL, 11.9, 4.0, .3, .3, ROSA)
        forma(sl, MSO_SHAPE.HEART, 11.6, 5.7, .5, .45, ROSA)
    else:
        forma(sl, MSO_SHAPE.OVAL, 12.0, -1.15, 2.3, 2.3, ROSACL)
        forma(sl, MSO_SHAPE.OVAL, 11.75, .95, .32, .32, LILACL)
        forma(sl, MSO_SHAPE.OVAL, -.85, 6.75, 1.7, 1.7, ROSACL)


def nueva(titulo, clave):
    sl = prs.slides.add_slide(BLANK)
    cnt[0] += 1
    fondo(sl)
    forma(sl, MSO_SHAPE.HEART, .6, .46, .32, .29, ROSA)
    caja(sl, 1.05, .22, 10.8, .8, titulo, 30, CIRUELA, True)
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, 1.12, 1.0, 1.6, .09, ROSA, radio=.5)
    forma(sl, MSO_SHAPE.OVAL, 12.4, 6.92, .42, .42, ROSA)
    caja(sl, 12.25, 6.93, .72, .4, f"{cnt[0]}", 12, BL, True, True)
    caja(sl, .9, 6.98, 6, .4, "Equipo 4 · Fundamentos de Econometría", 10, SUAVE)
    sl.notes_slide.notes_text_frame.text = f"[{QUIEN[clave]}] {GUION[clave]}"
    return sl


def dato(sl, x, y, valor, etiqueta, w=3.9, color=ROSA):
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, 2.45, BL, ROSACL if color == ROSA else LILACL, 2, .18)
    caja(sl, x, y + .2, w, 1.0, valor, 44, color, True, True)
    caja(sl, x + .15, y + 1.2, w - .3, 1.2, etiqueta, 17, TXT, False, True)


def mensaje(sl, texto, y=6.1):
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, .55, y, 12.2, .72, ROSA, radio=.5)
    caja(sl, .8, y + .11, 11.7, .6, "♡  " + texto, 17, BL, True, True)


def img(sl, nombre, x, y, h):
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x - .15, y - .12, h * 1.72 + .3, h + .24, BL, ROSACL, 1.5, .06)
    sl.shapes.add_picture(str(FIG / f"{nombre}.png"), Inches(x), Inches(y), height=Inches(h))


def tabla(sl, filas, x, y, w, anchos, size=16):
    tb = sl.shapes.add_table(len(filas), len(filas[0]), Inches(x), Inches(y), Inches(w), Inches(.5 * len(filas))).table
    for j, a in enumerate(anchos):
        tb.columns[j].width = Inches(a)
    for i, f in enumerate(filas):
        for j, v in enumerate(f):
            c = tb.cell(i, j)
            c.text = str(v)
            c.fill.solid()
            c.fill.fore_color.rgb = ROSA if i == 0 else (BL if i % 2 else ROSACL)
            for p in c.text_frame.paragraphs:
                p.font.size, p.font.bold, p.font.name = Pt(size), i == 0, FUENTE
                p.font.color.rgb = BL if i == 0 else TXT


def pastilla(sl, x, y, w, texto, relleno=ROSACL, color=CIRUELA, size=16):
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, .5, relleno, radio=.5)
    caja(sl, x, y + .04, w, .45, texto, size, color, True, True)


# 1 portada
sl = prs.slides.add_slide(BLANK)
cnt[0] += 1
fondo(sl, portada=True)
pastilla(sl, .9, 1.05, 3.6, "Fundamentos de Econometría", ROSA, BL, 14)
caja(sl, .9, 1.75, 11, 1.2, "Escolaridad, habilidad y salario", 48, CIRUELA, True)
caja(sl, .9, 2.85, 9.5, 1.2, "¿La multicolinealidad nos impide medir cuánto se asocia la escuela con el salario?", 23, ROSA)
caja(sl, .9, 4.25, 4, .4, "Equipo 4", 15, SUAVE, True)
for i, nom in enumerate(["Ximena Lara", "Camila Ricardi", "Sara Enriquez"]):
    pastilla(sl, .9 + i * 2.75, 4.7, 2.55, nom)
caja(sl, .9, 5.6, 10, .9, ["Prof. César Galindo · Universidad Panamericana · Noviembre 2026"], 15, TXT)
sl.notes_slide.notes_text_frame.text = "[Ximena] Presentar al equipo y el título."

# 2 pregunta
sl = nueva("La pregunta", "pregunta")
caja(sl, .8, 1.45, 11.8, 1.6, "¿Cuánto más gana alguien con un año más de escuela, comparado con alguien igual en "
     "experiencia, habilidad y familia?", 27, CIRUELA, True)
dato(sl, .8, 3.3, "H₁: β > 0", "esperamos una relación positiva", 3.8)
dato(sl, 4.75, 3.3, "α = 5 %", "nivel de significancia fijado antes de estimar", 3.8)
dato(sl, 8.7, 3.3, "≠ causa", "datos observacionales: medimos asociación", 3.8, LILA)
mensaje(sl, "Reto: escolaridad, habilidad y familia se mueven juntas → multicolinealidad")

# 3 datos
sl = nueva("Datos", "datos")
dato(sl, .8, 1.45, f"{n}", "hombres trabajadores de 28 a 38 años", 3.8)
dato(sl, 4.75, 1.45, "1980", "EE. UU., encuesta NLS (corte transversal)", 3.8)
dato(sl, 8.7, 1.45, "13", "variables explicativas", 3.8)
caja(sl, .8, 4.15, 11.8, 2.4, [f"♡  Fuente: base WAGE2 de Wooldridge (Blackburn y Neumark, 1992) · consultada el 18/11/2026",
                               f"♡  Limpieza: {n0} → {n} (quitamos {n0 - n} sin escolaridad de los padres)",
                               f"♡  En promedio: salario de {D['wage']['mean']:.0f} USD al mes y {D['educ']['mean']:.1f} años "
                               "de escuela",
                               "♡  Y = ln(salario mensual en USD de 1980) → los coeficientes se leen como %"], 18)

# 4 modelo
sl = nueva("El modelo", "modelo")
forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, 1.3, 1.35, 10.7, .9, BL, ROSACL, 2, .5)
caja(sl, .8, 1.47, 11.8, .9, "ln(salario) = β₀ + β₁·escolaridad + experiencia + controles + u", 23, ROSA, True, True)
tabla(sl, [["Grupo", "Variables", "Signo esperado"],
           ["De interés", "escolaridad (educ)", "+"],
           ["Experiencia", "exper (centrada), exper², antigüedad, edad", "+"],
           ["Habilidad", "IQ, KWW (prueba de conocimiento laboral)", "+"],
           ["Familia", "escolaridad de la madre y del padre", "+"],
           ["Dummies", "casado, afroamericano, sur, urbano", "+ − − +"]], 1.3, 2.6, 10.7, [2.3, 6.0, 2.4], 17)
caja(sl, .8, 6.25, 11.8, .6, "Estimación: MCO con errores robustos (HC1) · grupo de referencia: no casado, no afroamericano, "
     "fuera del sur, no urbano", 13.5, SUAVE, False, True)

# 5 resultado
sl = nueva("Resultado principal", "resultado")
img(sl, "pasos_educ", .65, 1.4, 4.4)
caja(sl, 8.95, 1.7, 4.2, 1.2, f"{e1:.1f} % → {e3:.1f} %", 36, ROSA, True)
caja(sl, 8.95, 2.9, 3.9, 3, "Al comparar personas con igual habilidad y familia, la asociación baja casi a la mitad", 20, CIRUELA)
mensaje(sl, "Parte de lo que parecía «escuela» era habilidad: sesgo por variable omitida")

# 6 interpretación
sl = nueva("¿Qué significa?", "interpretacion")
dato(sl, .8, 1.45, f"+{e3:.1f} %", f"de salario por un año más de escuela (≈ {usd:.0f} USD al mes)", 3.8)
dato(sl, 4.75, 1.45, f"{ci3[0]:.0f}–{ci3[1]:.0f} %", "intervalo de confianza 95 %: no incluye 0 → se rechaza H₀", 3.8)
dato(sl, 8.7, 1.45, f"+{cuatro:.0f} %", "con 4 años de universidad: relevante, no solo significativo", 3.8)
caja(sl, .8, 4.3, 11.8, 1.5, [f"Otros: casado {pct(M3['params']['married']):+.0f} % · urbano {pct(M3['params']['urban']):+.0f} % · "
                              f"sur {pct(M3['params']['south']):+.0f} %  ·  IQ: +{100*M3['params']['IQ']*D['IQ']['std']:.0f} % "
                              "por desviación estándar"], 18, TXT, False, True)
mensaje(sl, "Todo «manteniendo constantes las demás variables»")

# 7 multicolinealidad
sl = nueva("Multicolinealidad: ¿dónde está?", "multi")
img(sl, "vif_rosa", .65, 1.35, 4.5)
caja(sl, 8.6, 1.4, 4.4, 4.5, [f"VIF de {VF['exper²']['VIF']:.0f}: artificio", "exper² se construye con exper",
                              f"Al centrar: VIF máx. {R['vif_centrado_max']:.1f}", "",
                              "Real pero moderada:", "escolaridad, IQ, KWW y padres (VIF 1.6–2.2)"], 18.5, CIRUELA)
mensaje(sl, "Un VIF alto no se arregla quitando variables: primero hay que entender de dónde viene")

# 8 impide
sl = nueva("¿Nos impide responder?", "impide")
dato(sl, .8, 1.45, f"×{se_ratio:.1f}", "error estándar de escolaridad inflado por la correlación", 3.8, LILA)
dato(sl, 4.75, 1.45, f"F = {PF[0]['F']:.1f}", "IQ, KWW y padres: solos no significativos, juntos sí (p < 0.001)", 3.8)
dato(sl, 8.7, 1.45, f"{lo:.1f}–{hi:.1f} %", f"si quitamos o combinamos controles (vs {e3:.1f} %): casi igual", 3.8)
mensaje(sl, "No: el intervalo de escolaridad sigue lejos de cero. Solo cuesta separar los controles entre sí", 5.2)

# 9 diagnóstico
sl = nueva("¿Se cumplen los supuestos?", "diagnostico")
tabla(sl, [["Revisión", "Resultado", "Qué hicimos"],
           ["Varianza constante (Breusch–Pagan)", f"p = {H['bp_p']:.3f} → no", "Errores robustos HC1"],
           ["Forma funcional (RESET)", f"p = {R['reset']['p']:.2f} → bien", "Nada"],
           ["Normalidad (Jarque–Bera)", "p < 0.001 → no", f"Con n = {n} no afecta"],
           ["Exogeneidad", "No se puede probar", "Hablamos de asociación"]], .8, 1.5, 11.8, [4.4, 3.4, 4.0], 18)
mensaje(sl, "Ningún problema cambia la conclusión", 5.2)

# 10 conclusiones
sl = nueva("Conclusiones", "conclusion")
for i, t in enumerate([f"Un año más de escuela se asocia con +{e3:.1f} % de salario entre personas comparables",
                       "La multicolinealidad infla la varianza, pero no impide responder",
                       "Limitaciones: no es causal · posible selección · solo hombres de EE. UU. en 1980"]):
    y = 1.5 + i * 1.12
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, .8, y, 11.8, .9, BL, ROSACL, 1.5, .3)
    forma(sl, MSO_SHAPE.OVAL, 1.0, y + .17, .56, .56, ROSA)
    caja(sl, .95, y + .2, .66, .5, str(i + 1), 18, BL, True, True)
    caja(sl, 1.8, y + .17, 10.6, .7, t, 20, CIRUELA)
mensaje(sl, f"¿Qué cambiaría la conclusión? Incluir o no la habilidad: {e1:.1f} % → {e3:.1f} %", 5.15)

# 11 cierre
sl = nueva("Gracias · ¿preguntas?", "cierre")
forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, .8, 1.5, 11.8, 1.5, ROSACL, radio=.2)
caja(sl, 1.1, 1.62, 11.2, 1.3, ["En una frase:", f"un año más de escuela se asocia con +{e3:.1f} % de salario, y la "
                                "multicolinealidad no nos impide medirlo"], 21, CIRUELA, True, True)
caja(sl, .8, 3.45, 11.8, 2, ["♡  Código, datos y diccionario: github.com/ximeelaraa/mincer-multicolinealidad",
                             "♡  Fuentes: Blackburn y Neumark (1992) · Mincer (1974) · Wooldridge (2020) · Galindo (2026)"], 18)
for i, nom in enumerate(["Ximena Lara", "Camila Ricardi", "Sara Enriquez"]):
    pastilla(sl, 2.35 + i * 2.95, 5.3, 2.65, nom)
prs.save(ROOT / "Presentacion_Final.pptx")

# ------------------------------------------------------------------ BEAMER (mismo contenido)
def T(_t, **kw):
    for k, v in kw.items():
        _t = _t.replace(f"@{k}@", str(v))
    return _t


tex = T(r"""\documentclass[aspectratio=169,11pt,xcolor={table}]{beamer}
\usepackage{iftex}
\ifPDFTeX
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage{lmodern}
\else
\usepackage{fontspec}
\IfFontExistsTF{Avenir Next}{\setsansfont{Avenir Next}}{}
\usepackage[italic]{mathastext}
\fi
\usepackage[spanish,es-nodecimaldot,es-noshorthands]{babel}
\usepackage{amsmath,amssymb,booktabs,graphicx,tikz}
\usepackage[most]{tcolorbox}
\graphicspath{{figuras/}}
\definecolor{rosa}{HTML}{@ROSA@}
\definecolor{rosaclaro}{HTML}{@ROSACL@}
\definecolor{blush}{HTML}{@BLUSH@}
\definecolor{ciruela}{HTML}{@CIRUELA@}
\definecolor{lila}{HTML}{@LILA@}
\definecolor{lilaclaro}{HTML}{EDE3F7}
\definecolor{texto}{HTML}{@TEXTO@}
\definecolor{gris}{HTML}{@GRIS@}
\newcommand{\corazon}{\ensuremath{\heartsuit}}
\AtBeginDocument{\hyphenpenalty=10000\exhyphenpenalty=10000}
\setbeamercolor{normal text}{fg=texto}
\setbeamercolor{structure}{fg=rosa}
\setbeamercolor{alerted text}{fg=lila}
\setbeamertemplate{navigation symbols}{}
\setbeamertemplate{background}{%
  \begin{tikzpicture}
  \useasboundingbox (0,0) rectangle (\paperwidth,\paperheight);
  \fill[blush] (0,0) rectangle (\paperwidth,\paperheight);
  \fill[rosaclaro] (\paperwidth,\paperheight) circle (1.5cm);
  \fill[lilaclaro] (\paperwidth-1.75cm,\paperheight-1.45cm) circle (0.13cm);
  \fill[rosaclaro] (0,0) circle (0.9cm);
  \end{tikzpicture}}
\setbeamertemplate{frametitle}{%
  \vspace{0.55em}{\color{rosa}\small\corazon}\hspace{0.45em}{\color{ciruela}\Large\bfseries\insertframetitle}\par
  \vspace{-0.1em}\hspace{1.35em}\tikz\draw[rosa,line width=2.6pt,line cap=round] (0,0)--(1.6cm,0);\par}
\setbeamertemplate{footline}{%
  \hspace{1.6em}\raisebox{0.75em}{\tiny\color{gris}Equipo 4 $\cdot$ Fundamentos de Econometría}\hfill
  \raisebox{0.45em}{\tikz\node[circle,fill=rosa,text=white,inner sep=0pt,minimum size=1.35em,font=\tiny\bfseries]{\insertframenumber};}\hspace{1.2em}}
\setbeamertemplate{itemize item}{\color{rosa}\corazon}
\setbeamertemplate{enumerate item}{\tikz[baseline=(n.base)]\node[circle,fill=rosa,text=white,inner sep=0pt,minimum size=1.3em,font=\footnotesize\bfseries](n){\insertenumlabel};}
\newcommand{\pastilla}[2][rosaclaro]{\tikz[baseline=(p.base)]\node[fill=#1,rounded corners=7pt,inner xsep=9pt,inner ysep=4pt,text=ciruela,font=\small\bfseries](p){#2};}
\newtcolorbox{tarjeta}[1][rosaclaro]{enhanced,colback=white,colframe=#1,boxrule=1pt,arc=9pt,left=6pt,right=6pt,top=4pt,bottom=4pt}
\newcommand{\dato}[3][rosa]{%
  \begin{tcolorbox}[enhanced,width=0.31\textwidth,height=2.75cm,valign=center,nobeforeafter,box align=top,colback=white,
    colframe=#1!25,boxrule=1pt,arc=10pt,left=5pt,right=5pt,halign=center]
  \hyphenpenalty=10000{\fontsize{21}{25}\selectfont\color{#1}\textbf{#2}}\\[0.45em]{\small #3}\end{tcolorbox}}
\newcommand{\mensaje}[1]{\par\vfill
  \begin{tcolorbox}[enhanced,colback=rosa,colframe=rosa,boxrule=0pt,arc=11pt,halign=center,left=8pt,right=8pt,top=4pt,bottom=4pt]
  \hyphenpenalty=10000\color{white}\bfseries\corazon\ \ #1\end{tcolorbox}\vspace{0.1em}}
\newcommand{\figura}[1]{\begin{tarjeta}\includegraphics[width=\linewidth]{#1}\end{tarjeta}}
\setbeamertemplate{title page}{%
  \vspace{1.2em}
  \pastilla[rosa]{\color{white}Fundamentos de Econometría}\par\vspace{1em}
  {\fontsize{26}{30}\selectfont\color{ciruela}\bfseries\inserttitle\par}\vspace{0.6em}
  \parbox{0.72\textwidth}{\raggedright\large\color{rosa}\insertsubtitle}\par\vspace{1.6em}
  {\small\color{gris}\bfseries Equipo 4}\par\vspace{0.4em}
  \pastilla{Ximena Lara}\hspace{0.5em}\pastilla{Camila Ricardi}\hspace{0.5em}\pastilla{Sara Enriquez}\par\vspace{1.4em}
  {\small\color{texto}\insertinstitute\ $\cdot$ \insertdate\par}}

\title{Escolaridad, habilidad y salario}
\subtitle{¿La multicolinealidad nos impide medir cuánto se asocia la escuela con el salario?}
\author{Equipo 4: @EQ@}
\institute{Prof. César Galindo $\cdot$ Universidad Panamericana}
\date{Noviembre 2026}

\begin{document}
{\setbeamertemplate{background}{%
  \begin{tikzpicture}
  \useasboundingbox (0,0) rectangle (\paperwidth,\paperheight);
  \fill[blush] (0,0) rectangle (\paperwidth,\paperheight);
  \fill[rosaclaro] (\paperwidth+0.4cm,\paperheight+0.3cm) circle (3.6cm);
  \fill[lilaclaro] (\paperwidth-2.6cm,\paperheight-4.3cm) circle (0.3cm);
  \fill[rosa] (\paperwidth-1.2cm,\paperheight-5.4cm) circle (0.12cm);
  \fill[rosaclaro] (-0.3cm,-0.3cm) circle (1.9cm);
  \node[text=rosa,font=\Large] at (\paperwidth-2.3cm,1.4cm) {\corazon};
  \end{tikzpicture}}
\begin{frame}[plain]\titlepage\end{frame}}

\begin{frame}{La pregunta}
\vspace{0.3em}
{\Large\color{ciruela}\textbf{¿Cuánto más gana alguien con un año más de escuela, comparado con alguien igual en experiencia, habilidad y familia?}}
\vspace{1em}

\dato{$H_1:\beta>0$}{esperamos una relación positiva}\hfill
\dato{$\alpha=5\%$}{nivel de significancia fijado antes de estimar}\hfill
\dato[lila]{$\neq$ causa}{datos observacionales: medimos asociación}
\mensaje{Reto: escolaridad, habilidad y familia se mueven juntas $\to$ multicolinealidad}
\end{frame}

\begin{frame}{Datos}
\vspace{0.3em}
\dato{@n@}{hombres trabajadores de 28 a 38 años}\hfill
\dato{1980}{EE.\,UU., encuesta NLS (corte transversal)}\hfill
\dato{13}{variables explicativas}
\vspace{1em}
\begin{itemize}
\item Fuente: base WAGE2 de Wooldridge (Blackburn y Neumark, 1992), consultada el 18/11/2026.
\item Limpieza: @n0@ $\to$ @n@ (quitamos @nd@ sin escolaridad de los padres).
\item En promedio: salario de @wm@ USD al mes y @em@ años de escuela.
\item $Y=\ln(\text{salario mensual en USD de 1980})$ $\to$ los coeficientes se leen como \%.
\end{itemize}
\end{frame}

\begin{frame}{El modelo}
\begin{tcolorbox}[enhanced,colback=white,colframe=rosaclaro,boxrule=1pt,arc=12pt,halign=center,top=1pt,bottom=1pt]
\[
\color{rosa}\ln(\text{salario})=\beta_0+\beta_1\,\text{escolaridad}+\text{experiencia}+\text{controles}+u
\]
\end{tcolorbox}
\centering\small
\rowcolors{2}{white}{rosaclaro}
\begin{tabular}{lll}
\rowcolor{rosa}\color{white}\textbf{Grupo} & \color{white}\textbf{Variables} & \color{white}\textbf{Signo esperado}\\
De interés & escolaridad (educ) & $+$\\
Experiencia & exper (centrada), exper$^2$, antigüedad, edad & $+$\\
Habilidad & IQ, KWW (prueba de conocimiento laboral) & $+$\\
Familia & escolaridad de la madre y del padre & $+$\\
Dummies & casado, afroamericano, sur, urbano & $+\ -\ -\ +$\\
\end{tabular}

\vspace{0.7em}
\scriptsize\color{gris} Estimación: MCO con errores robustos (HC1). Referencia: no casado, no afroamericano, fuera del sur, no urbano.
\end{frame}

\begin{frame}{Resultado principal}
\begin{columns}[c]
\column{0.58\textwidth}
\figura{pasos_educ}
\column{0.4\textwidth}
{\LARGE\color{rosa}\textbf{@e1@\,\% $\to$ @e3@\,\%}}\\[0.6em]
{\color{ciruela}Al comparar personas con igual habilidad y familia, la asociación baja casi a la mitad.}
\end{columns}
\mensaje{Parte de lo que parecía «escuela» era habilidad: sesgo por variable omitida}
\end{frame}

\begin{frame}{¿Qué significa?}
\vspace{0.3em}
\dato{$+$@e3@\,\%}{de salario por un año más de escuela ($\approx$ @usd@ USD al mes)}\hfill
\dato{@c0@--@c1@\,\%}{intervalo de confianza 95\,\%: no incluye 0 $\to$ se rechaza $H_0$}\hfill
\dato{$+$@cu@\,\%}{con 4 años de universidad: relevante, no solo significativo}
\vspace{1.2em}

\centering Otros: casado @mar@\,\% $\cdot$ urbano @urb@\,\% $\cdot$ sur @sur@\,\% $\cdot$ IQ: $+$@iq@\,\% por desviación estándar.
\mensaje{Todo «manteniendo constantes las demás variables»}
\end{frame}

\begin{frame}{Multicolinealidad: ¿dónde está?}
\begin{columns}[c]
\column{0.54\textwidth}
\figura{vif_rosa}
\column{0.44\textwidth}
{\color{lila}\textbf{VIF de @v2@: artificio}}\\ exper$^2$ se construye con exper.\\ Al centrar: VIF máx.\ @vc@.\\[1em]
{\color{rosa}\textbf{Real pero moderada:}}\\ escolaridad $\leftrightarrow$ IQ $\leftrightarrow$ KWW $\leftrightarrow$ padres (VIF 1.6--2.2).
\end{columns}
\mensaje{Un VIF alto no se arregla quitando variables: primero hay que entender de dónde viene}
\end{frame}

\begin{frame}{¿Nos impide responder?}
\vspace{0.3em}
\dato[lila]{$\times$@sr@}{error estándar de escolaridad inflado por la correlación}\hfill
\dato{$F=@F0@$}{IQ, KWW y padres: solos no significativos, juntos sí ($p<0.001$)}\hfill
\dato{@br@--@bp@\,\%}{si quitamos o combinamos controles (vs @e3@\,\%): casi igual}
\mensaje{No: el intervalo de escolaridad sigue lejos de cero. Solo cuesta separar los controles entre sí}
\end{frame}

\begin{frame}{¿Se cumplen los supuestos?}
\vspace{0.6em}
\centering
\rowcolors{2}{white}{rosaclaro}
\small\renewcommand{\arraystretch}{1.45}
\begin{tabular}{lll}
\rowcolor{rosa}\color{white}\textbf{Revisión} & \color{white}\textbf{Resultado} & \color{white}\textbf{Qué hicimos}\\
Varianza constante (Breusch--Pagan) & $p=@bp_@$ $\to$ no & Errores robustos HC1\\
Forma funcional (RESET) & $p=@rs@$ $\to$ bien & Nada\\
Normalidad (Jarque--Bera) & $p<0.001$ $\to$ no & Con $n=@n@$ no afecta\\
Exogeneidad & No se puede probar & Hablamos de asociación\\
\end{tabular}
\mensaje{Ningún problema cambia la conclusión}
\end{frame}

\begin{frame}{Conclusiones}
\vspace{0.4em}
\large
\begin{enumerate}\setlength{\itemsep}{0.7em}
\item Un año más de escuela $\leftrightarrow$ $+$@e3@\,\% de salario entre personas comparables.
\item La multicolinealidad infla la varianza, pero no impide responder.
\item Limitaciones: no es causal $\cdot$ posible selección $\cdot$ solo hombres de EE.\,UU. en 1980.
\end{enumerate}
\normalsize
\mensaje{¿Qué cambiaría la conclusión? Incluir o no la habilidad: @e1@\,\% $\to$ @e3@\,\%}
\end{frame}

\begin{frame}{Gracias $\cdot$ ¿preguntas?}
\vspace{0.3em}
\begin{tcolorbox}[enhanced,colback=rosaclaro,colframe=rosaclaro,arc=12pt,halign=center]
{\color{ciruela}\bfseries En una frase: un año más de escuela se asocia con $+$@e3@\,\% de salario,\\ y la multicolinealidad no nos impide medirlo.}
\end{tcolorbox}
\begin{itemize}
\item Código, datos y diccionario: \texttt{github.com/ximeelaraa/mincer-multicolinealidad}
\item Fuentes: Blackburn y Neumark (1992) $\cdot$ Mincer (1974) $\cdot$ Wooldridge (2020) $\cdot$ Galindo (2026)
\end{itemize}
\vfill
\centering\pastilla{Ximena Lara}\hspace{0.5em}\pastilla{Camila Ricardi}\hspace{0.5em}\pastilla{Sara Enriquez}
\vspace{0.6em}
\end{frame}
\end{document}
""", EQ=EQUIPO.replace(" · ", ", "), n=n, n0=n0, nd=n0 - n, e1=f"{e1:.1f}", e3=f"{e3:.1f}", usd=f"{usd:.0f}",
        c0=f"{ci3[0]:.0f}", c1=f"{ci3[1]:.0f}", cu=f"{cuatro:.0f}", mar=f"{pct(M3['params']['married']):+.0f}",
        urb=f"{pct(M3['params']['urban']):+.0f}", sur=f"{pct(M3['params']['south']):+.0f}".replace("-", "$-$"),
        iq=f"{100*M3['params']['IQ']*D['IQ']['std']:.0f}", v2=f"{VF['exper²']['VIF']:.0f}", vc=f"{R['vif_centrado_max']:.1f}",
        sr=f"{se_ratio:.1f}", F0=f"{PF[0]['F']:.1f}", br=f"{lo:.1f}", wm=f"{D['wage']['mean']:.0f}",
        em=f"{D['educ']['mean']:.1f}", bp=f"{hi:.1f}", bp_=f"{H['bp_p']:.3f}", rs=f"{R['reset']['p']:.2f}",
        ROSA=ROSA_H[1:], ROSACL=ROSACL_H[1:], BLUSH=BLUSH_H[1:], CIRUELA=CIRUELA_H[1:], LILA=LILA_H[1:],
        TEXTO=TEXTO_H[1:], GRIS=GRIS_H[1:])
(ROOT / "Presentacion_Beamer.tex").write_text(tex, encoding="utf-8")
with zipfile.ZipFile(ROOT / "Beamer_Overleaf.zip", "w", zipfile.ZIP_DEFLATED) as z:
    z.write(ROOT / "Presentacion_Beamer.tex", "Presentacion_Beamer.tex")
    for f in ("pasos_educ", "vif_rosa"):
        z.write(FIG / f"{f}.png", f"figuras/{f}.png")
print("Listo: Presentacion_Final.pptx, Presentacion_Beamer.tex, Beamer_Overleaf.zip")
