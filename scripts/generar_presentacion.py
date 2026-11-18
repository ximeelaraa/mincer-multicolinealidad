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

# ------------------------------------------------------------------ figura: el número clave en tres pasos
TINTA, AZUL, GRIS, FONDO = "#1D2433", "#2F6F9F", "#8A8F98", "#FFFFFF"
fig, ax = plt.subplots(figsize=(7.2, 4.2), dpi=200)
etiquetas = ["Solo escolaridad", "+ experiencia", "+ habilidad, familia\ny demás controles"]
vals = [e1, e2, e3]
cis = [[pct(x) for x in M["ci"]["educ"]] for M in (M1, M2, M3)]
x = np.arange(3)
for i in range(3):
    ax.plot([x[i], x[i]], cis[i], color=AZUL if i == 2 else GRIS, lw=2, solid_capstyle="round")
    ax.scatter(x[i], vals[i], s=90, color=AZUL if i == 2 else GRIS, zorder=3, edgecolor=FONDO, linewidth=2)
    ax.text(x[i] + .09, vals[i], f"{vals[i]:.1f}%", va="center", fontsize=13, color=TINTA,
            fontweight="bold" if i == 2 else "normal")
ax.set_xticks(x, etiquetas, fontsize=10, color=TINTA)
ax.set_ylabel("% más salario por un año más de escuela", fontsize=9.5, color=TINTA)
ax.set_ylim(0, 10.5)
ax.set_xlim(-.4, 2.6)
ax.grid(axis="y", color="#E6E8EB", lw=.8)
for s in ("top", "right", "left"):
    ax.spines[s].set_visible(False)
ax.spines["bottom"].set_color("#C9CDD3")
ax.tick_params(axis="y", colors=GRIS, length=0)
ax.tick_params(axis="x", length=0)
ax.set_axisbelow(True)
fig.text(.01, .01, "Líneas: intervalo de confianza al 95 % (errores robustos HC1)", fontsize=8, color=GRIS)
fig.tight_layout(rect=(0, .03, 1, 1))
fig.savefig(FIG / "pasos_educ.png", facecolor=FONDO)
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
OSC, AC, TXT, BL, NAR, SUAVE = (RGBColor(0x1D, 0x24, 0x33), RGBColor(0x2F, 0x6F, 0x9F), RGBColor(0x33, 0x33, 0x33),
                               RGBColor(255, 255, 255), RGBColor(0xD1, 0x60, 0x3D), RGBColor(0x8A, 0x8F, 0x98))
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]
cnt = [0]


def caja(sl, x, y, w, h, texto, size=20, color=TXT, bold=False, centro=False):
    tf = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
    tf.word_wrap = True
    for i, ln in enumerate(texto if isinstance(texto, list) else [texto]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text, p.font.size, p.font.color.rgb, p.font.bold = ln, Pt(size), color, bold
        p.space_after = Pt(10)
        if centro:
            p.alignment = 2
    return tf


def nueva(titulo, clave):
    sl = prs.slides.add_slide(BLANK)
    cnt[0] += 1
    bar = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1.0))
    bar.fill.solid()
    bar.fill.fore_color.rgb = OSC
    bar.line.fill.background()
    caja(sl, .55, .18, 11.5, .8, titulo, 28, BL, True)
    caja(sl, 12.3, 6.95, .9, .4, f"{cnt[0]}", 12, SUAVE)
    sl.notes_slide.notes_text_frame.text = f"[{QUIEN[clave]}] {GUION[clave]}"
    return sl


def dato(sl, x, y, valor, etiqueta, w=3.9, color=AC):
    caja(sl, x, y, w, 1.1, valor, 48, color, True, True)
    caja(sl, x, y + 1.15, w, 1.2, etiqueta, 16, TXT, False, True)


def mensaje(sl, texto, y=6.15):
    box = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(.55), Inches(y), Inches(12.2), Inches(.75))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xEE, 0xF3, 0xF8)
    box.line.fill.background()
    caja(sl, .75, y + .12, 11.8, .6, texto, 19, OSC, True)


def img(sl, nombre, x, y, w=None, h=None):
    sl.shapes.add_picture(str(FIG / f"{nombre}.png"), Inches(x), Inches(y),
                          width=Inches(w) if w else None, height=Inches(h) if h else None)


def tabla(sl, filas, x, y, w, anchos, size=16):
    tb = sl.shapes.add_table(len(filas), len(filas[0]), Inches(x), Inches(y), Inches(w), Inches(.5 * len(filas))).table
    for j, a in enumerate(anchos):
        tb.columns[j].width = Inches(a)
    for i, f in enumerate(filas):
        for j, v in enumerate(f):
            c = tb.cell(i, j)
            c.text = str(v)
            for p in c.text_frame.paragraphs:
                p.font.size = Pt(size)
                p.font.bold = i == 0


# 1 portada
sl = prs.slides.add_slide(BLANK)
cnt[0] += 1
bg = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
bg.fill.solid()
bg.fill.fore_color.rgb = OSC
bg.line.fill.background()
caja(sl, .9, 1.7, 11.5, 1.2, "Escolaridad, habilidad y salario", 46, BL, True)
caja(sl, .9, 2.85, 11.5, 1, "¿La multicolinealidad nos impide medir cuánto se asocia la escuela con el salario?", 24,
     RGBColor(0xBF, 0xD4, 0xEE))
caja(sl, .9, 4.7, 11.5, 1.8, [f"Equipo 4: {EQUIPO}", "Fundamentos de Econometría · Prof. César Galindo · Universidad Panamericana",
                              "Noviembre 2026"], 17, BL)
sl.notes_slide.notes_text_frame.text = "[Ximena] Presentar al equipo y el título."

# 2 pregunta
sl = nueva("La pregunta", "pregunta")
caja(sl, .8, 1.5, 11.8, 1.6, "¿Cuánto más gana alguien con un año más de escuela, comparado con alguien igual en "
     "experiencia, habilidad y familia?", 28, OSC, True)
dato(sl, .8, 3.3, "H₁: β > 0", "esperamos una relación positiva", 3.8)
dato(sl, 4.75, 3.3, "α = 5 %", "nivel de significancia fijado antes de estimar", 3.8)
dato(sl, 8.7, 3.3, "≠ causa", "datos observacionales: medimos asociación", 3.8, NAR)
mensaje(sl, "Reto: escolaridad, habilidad y familia se mueven juntas → multicolinealidad")

# 3 datos
sl = nueva("Datos", "datos")
dato(sl, .8, 1.5, f"{n}", "hombres trabajadores de 28 a 38 años", 3.8)
dato(sl, 4.75, 1.5, "1980", "EE. UU., encuesta NLS (corte transversal)", 3.8)
dato(sl, 8.7, 1.5, "13", "variables explicativas", 3.8)
caja(sl, .8, 4.1, 11.8, 1.6, [f"Fuente: base WAGE2 de Wooldridge (Blackburn y Neumark, 1992) · consultada el 18/11/2026",
                              f"Limpieza: {n0} → {n} (quitamos {n0 - n} sin escolaridad de los padres)",
                              "Y = ln(salario mensual en USD de 1980) → los coeficientes se leen como %"], 18)

# 4 modelo
sl = nueva("El modelo", "modelo")
caja(sl, .8, 1.4, 11.8, .9, "ln(salario) = β₀ + β₁·escolaridad + experiencia + controles + u", 28, AC, True, True)
tabla(sl, [["Grupo", "Variables", "Signo esperado"],
           ["De interés", "escolaridad (educ)", "+"],
           ["Experiencia", "exper (centrada), exper², antigüedad, edad", "+"],
           ["Habilidad", "IQ, KWW (prueba de conocimiento laboral)", "+"],
           ["Familia", "escolaridad de la madre y del padre", "+"],
           ["Dummies", "casado, afroamericano, sur, urbano", "+ − − +"]], 1.3, 2.6, 10.7, [2.3, 6.0, 2.4], 17)
caja(sl, .8, 6.3, 11.8, .6, "Estimación: MCO con errores robustos (HC1) · grupo de referencia: no casado, no afroamericano, "
     "fuera del sur, no urbano", 14, SUAVE)

# 5 resultado
sl = nueva("Resultado principal", "resultado")
img(sl, "pasos_educ", .5, 1.25, h=4.75)
caja(sl, 8.9, 1.7, 4.2, 1.2, f"{e1:.1f} % → {e3:.1f} %", 36, AC, True)
caja(sl, 8.9, 2.9, 4.0, 3, "Al comparar personas con igual habilidad y familia, la asociación baja casi a la mitad", 20, OSC)
mensaje(sl, "Parte de lo que parecía «escuela» era habilidad: sesgo por variable omitida")

# 6 interpretación
sl = nueva("¿Qué significa?", "interpretacion")
dato(sl, .8, 1.5, f"+{e3:.1f} %", f"de salario por un año más de escuela (≈ {usd:.0f} USD al mes)", 3.8)
dato(sl, 4.75, 1.5, f"{ci3[0]:.0f}–{ci3[1]:.0f} %", "intervalo de confianza 95 %: no incluye 0 → se rechaza H₀", 3.8)
dato(sl, 8.7, 1.5, f"+{cuatro:.0f} %", "con 4 años de universidad: relevante, no solo significativo", 3.8)
caja(sl, .8, 4.35, 11.8, 1.5, [f"Otros: casado {pct(M3['params']['married']):+.0f} % · urbano {pct(M3['params']['urban']):+.0f} % · "
                               f"sur {pct(M3['params']['south']):+.0f} %  ·  IQ: +{100*M3['params']['IQ']*D['IQ']['std']:.0f} % por "
                               "desviación estándar"], 18)
mensaje(sl, "Todo «manteniendo constantes las demás variables»")

# 7 multicolinealidad
sl = nueva("Multicolinealidad: ¿dónde está?", "multi")
img(sl, "vif", .5, 1.2, h=4.7)
caja(sl, 8.9, 1.4, 4.1, 4.5, [f"VIF de {VF['exper²']['VIF']:.0f}: artificio", "exper² se construye con exper",
                              f"Al centrar: VIF máx. {R['vif_centrado_max']:.1f}", "",
                              "Real pero moderada:", "escolaridad ↔ IQ ↔ KWW ↔ padres (VIF 1.6–2.2)"], 19, OSC)
mensaje(sl, "Un VIF alto no se arregla quitando variables: primero hay que entender de dónde viene")

# 8 impide
sl = nueva("¿Nos impide responder?", "impide")
dato(sl, .8, 1.5, f"×{se_ratio:.1f}", "error estándar de escolaridad inflado por la correlación", 3.8, NAR)
dato(sl, 4.75, 1.5, f"F = {PF[0]['F']:.1f}", "IQ, KWW y padres: solos no significativos, juntos sí (p < 0.001)", 3.8)
dato(sl, 8.7, 1.5, f"{lo:.1f}–{hi:.1f} %", f"si quitamos o combinamos controles (vs {e3:.1f} %): casi igual", 3.8)
mensaje(sl, "No: el intervalo de escolaridad sigue lejos de cero. Solo cuesta separar los controles entre sí", 5.4)

# 9 diagnóstico
sl = nueva("¿Se cumplen los supuestos?", "diagnostico")
tabla(sl, [["Revisión", "Resultado", "Qué hicimos"],
           ["Varianza constante (Breusch–Pagan)", f"p = {H['bp_p']:.3f} → no", "Errores robustos HC1"],
           ["Forma funcional (RESET)", f"p = {R['reset']['p']:.2f} → bien", "Nada"],
           ["Normalidad (Jarque–Bera)", "p < 0.001 → no", f"Con n = {n} no afecta"],
           ["Exogeneidad", "No se puede probar", "Hablamos de asociación"]], .8, 1.5, 11.8, [4.4, 3.4, 4.0], 18)
mensaje(sl, "Ningún problema cambia la conclusión", 5.4)

# 10 conclusiones
sl = nueva("Conclusiones", "conclusion")
caja(sl, .8, 1.4, 11.8, 3.2, [f"1.  Un año más de escuela ↔ +{e3:.1f} % de salario entre personas comparables",
                              "2.  La multicolinealidad infla la varianza, pero no impide responder",
                              "3.  Limitaciones: no es causal · posible selección · solo hombres de EE. UU. en 1980"], 22, OSC)
mensaje(sl, f"¿Qué cambiaría la conclusión? Incluir o no la habilidad: {e1:.1f} % → {e3:.1f} %", 5.0)

# 11 cierre
sl = nueva("Gracias · ¿preguntas?", "cierre")
caja(sl, .8, 1.8, 11.8, 3, ["Código, datos y diccionario: github.com/ximeelaraa/mincer-multicolinealidad",
                            "Fuentes: Blackburn y Neumark (1992) · Mincer (1974) · Wooldridge (2020) · Galindo (2026)"], 20)
prs.save(ROOT / "Presentacion_Final.pptx")

# ------------------------------------------------------------------ BEAMER (mismo contenido)
def T(_t, **kw):
    for k, v in kw.items():
        _t = _t.replace(f"@{k}@", str(v))
    return _t


tex = T(r"""\documentclass[aspectratio=169,11pt]{beamer}
\usepackage{iftex}
\ifPDFTeX
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage{lmodern}
\else
\usepackage{fontspec}
\fi
\usepackage[spanish,es-nodecimaldot,es-noshorthands]{babel}
\usepackage{amsmath,amssymb,booktabs,graphicx}
\graphicspath{{figuras/}}
\definecolor{tinta}{HTML}{1D2433}
\definecolor{azul}{HTML}{2F6F9F}
\definecolor{naranja}{HTML}{D1603D}
\definecolor{suave}{HTML}{EEF3F8}
\setbeamercolor{frametitle}{bg=tinta,fg=white}
\setbeamercolor{structure}{fg=azul}
\setbeamercolor{alerted text}{fg=naranja}
\setbeamertemplate{navigation symbols}{}
\setbeamertemplate{frametitle}{\nointerlineskip\begin{beamercolorbox}[wd=\paperwidth,ht=3ex,dp=1.4ex,leftskip=1em]{frametitle}\usebeamerfont{frametitle}\insertframetitle\end{beamercolorbox}}
\setbeamertemplate{footline}{\hfill\scriptsize\color{gray}\insertframenumber\hspace{1em}\vspace{0.6em}}
\newcommand{\dato}[3][azul]{\begin{minipage}[t]{0.31\textwidth}\centering{\fontsize{26}{30}\selectfont\color{#1}\textbf{#2}}\\[0.5em]\small #3\end{minipage}}
\newcommand{\mensaje}[1]{\par\vfill\colorbox{suave}{\parbox{\dimexpr\textwidth-2\fboxsep}{\centering\color{tinta}\textbf{#1}}}\par\vspace{0.3em}}

\title{Escolaridad, habilidad y salario}
\subtitle{¿La multicolinealidad nos impide medir cuánto se asocia la escuela con el salario?}
\author{Equipo 4: @EQ@}
\institute{Fundamentos de Econometría · Prof. César Galindo · Universidad Panamericana}
\date{Noviembre 2026}

\begin{document}
\begin{frame}\titlepage\end{frame}

\begin{frame}{La pregunta}
\vspace{0.5em}
{\Large\color{tinta}\textbf{¿Cuánto más gana alguien con un año más de escuela, comparado con alguien igual en experiencia, habilidad y familia?}}
\vspace{1.2em}

\dato{$H_1:\beta>0$}{esperamos una relación positiva}\hfill
\dato{$\alpha=5\%$}{nivel de significancia fijado antes de estimar}\hfill
\dato[naranja]{$\neq$ causa}{datos observacionales: medimos asociación}
\mensaje{Reto: escolaridad, habilidad y familia se mueven juntas $\to$ multicolinealidad}
\end{frame}

\begin{frame}{Datos}
\vspace{0.5em}
\dato{@n@}{hombres trabajadores de 28 a 38 años}\hfill
\dato{1980}{EE.\,UU., encuesta NLS (corte transversal)}\hfill
\dato{13}{variables explicativas}
\vspace{1.5em}
\begin{itemize}
\item Fuente: base WAGE2 de Wooldridge (Blackburn y Neumark, 1992), consultada el 18/11/2026.
\item Limpieza: @n0@ $\to$ @n@ (quitamos @nd@ sin escolaridad de los padres).
\item $Y=\ln(\text{salario mensual en USD de 1980})$ $\to$ los coeficientes se leen como \%.
\end{itemize}
\end{frame}

\begin{frame}{El modelo}
\[
\color{azul}\ln(\text{salario})=\beta_0+\beta_1\,\text{escolaridad}+\text{experiencia}+\text{controles}+u
\]
\centering\small
\begin{tabular}{lll}
\toprule
\textbf{Grupo} & \textbf{Variables} & \textbf{Signo esperado}\\
\midrule
De interés & escolaridad (educ) & $+$\\
Experiencia & exper (centrada), exper$^2$, antigüedad, edad & $+$\\
Habilidad & IQ, KWW (prueba de conocimiento laboral) & $+$\\
Familia & escolaridad de la madre y del padre & $+$\\
Dummies & casado, afroamericano, sur, urbano & $+\ -\ -\ +$\\
\bottomrule
\end{tabular}

\vspace{0.8em}
\scriptsize\color{gray} Estimación: MCO con errores robustos (HC1). Referencia: no casado, no afroamericano, fuera del sur, no urbano.
\end{frame}

\begin{frame}{Resultado principal}
\begin{columns}[c]
\column{0.62\textwidth}
\includegraphics[width=\linewidth]{pasos_educ}
\column{0.36\textwidth}
{\LARGE\color{tinta}\textbf{@e1@\,\% $\to$ @e3@\,\%}}\\[0.6em]
Al comparar personas con igual habilidad y familia, la asociación baja casi a la mitad.
\end{columns}
\mensaje{Parte de lo que parecía «escuela» era habilidad: sesgo por variable omitida}
\end{frame}

\begin{frame}{¿Qué significa?}
\vspace{0.5em}
\dato{$+$@e3@\,\%}{de salario por un año más de escuela ($\approx$ @usd@ USD al mes)}\hfill
\dato{@c0@--@c1@\,\%}{intervalo de confianza 95\,\%: no incluye 0 $\to$ se rechaza $H_0$}\hfill
\dato{$+$@cu@\,\%}{con 4 años de universidad: relevante, no solo significativo}
\vspace{1.5em}

Otros: casado @mar@\,\% $\cdot$ urbano @urb@\,\% $\cdot$ sur @sur@\,\% $\cdot$ IQ: $+$@iq@\,\% por desviación estándar.
\mensaje{Todo «manteniendo constantes las demás variables»}
\end{frame}

\begin{frame}{Multicolinealidad: ¿dónde está?}
\begin{columns}[c]
\column{0.6\textwidth}
\includegraphics[width=\linewidth]{vif}
\column{0.38\textwidth}
{\color{naranja}\textbf{VIF de @v2@: artificio}}\\ exper$^2$ se construye con exper.\\ Al centrar: VIF máx.\ @vc@.\\[1em]
{\color{azul}\textbf{Real pero moderada:}}\\ escolaridad $\leftrightarrow$ IQ $\leftrightarrow$ KWW $\leftrightarrow$ padres (VIF 1.6--2.2).
\end{columns}
\mensaje{Un VIF alto no se arregla quitando variables: primero hay que entender de dónde viene}
\end{frame}

\begin{frame}{¿Nos impide responder?}
\vspace{0.5em}
\dato[naranja]{$\times$@sr@}{error estándar de escolaridad inflado por la correlación}\hfill
\dato{$F=@F0@$}{IQ, KWW y padres: solos no significativos, juntos sí ($p<0.001$)}\hfill
\dato{@br@--@bp@\,\%}{si quitamos o combinamos controles (vs @e3@\,\%): casi igual}
\mensaje{No: el intervalo de escolaridad sigue lejos de cero. Solo cuesta separar los controles entre sí}
\end{frame}

\begin{frame}{¿Se cumplen los supuestos?}
\centering
\begin{tabular}{lll}
\toprule
\textbf{Revisión} & \textbf{Resultado} & \textbf{Qué hicimos}\\
\midrule
Varianza constante (Breusch--Pagan) & $p=@bp_@$ $\to$ no & Errores robustos HC1\\
Forma funcional (RESET) & $p=@rs@$ $\to$ bien & Nada\\
Normalidad (Jarque--Bera) & $p<0.001$ $\to$ no & Con $n=@n@$ no afecta\\
Exogeneidad & No se puede probar & Hablamos de asociación\\
\bottomrule
\end{tabular}
\mensaje{Ningún problema cambia la conclusión}
\end{frame}

\begin{frame}{Conclusiones}
\Large
\begin{enumerate}
\item Un año más de escuela $\leftrightarrow$ $+$@e3@\,\% de salario entre personas comparables.
\item La multicolinealidad infla la varianza, pero no impide responder.
\item Limitaciones: no es causal $\cdot$ posible selección $\cdot$ solo hombres de EE.\,UU. en 1980.
\end{enumerate}
\normalsize
\mensaje{¿Qué cambiaría la conclusión? Incluir o no la habilidad: @e1@\,\% $\to$ @e3@\,\%}
\end{frame}

\begin{frame}{Gracias $\cdot$ ¿preguntas?}
\vspace{1.5em}
\begin{itemize}
\item Código, datos y diccionario: \texttt{github.com/ximeelaraa/mincer-multicolinealidad}
\item Fuentes: Blackburn y Neumark (1992) $\cdot$ Mincer (1974) $\cdot$ Wooldridge (2020) $\cdot$ Galindo (2026)
\end{itemize}
\end{frame}
\end{document}
""", EQ=EQUIPO.replace(" · ", ", "), n=n, n0=n0, nd=n0 - n, e1=f"{e1:.1f}", e3=f"{e3:.1f}", usd=f"{usd:.0f}",
        c0=f"{ci3[0]:.0f}", c1=f"{ci3[1]:.0f}", cu=f"{cuatro:.0f}", mar=f"{pct(M3['params']['married']):+.0f}",
        urb=f"{pct(M3['params']['urban']):+.0f}", sur=f"{pct(M3['params']['south']):+.0f}".replace("-", "$-$"),
        iq=f"{100*M3['params']['IQ']*D['IQ']['std']:.0f}", v2=f"{VF['exper²']['VIF']:.0f}", vc=f"{R['vif_centrado_max']:.1f}",
        sr=f"{se_ratio:.1f}", F0=f"{PF[0]['F']:.1f}", br=f"{lo:.1f}",
        bp=f"{hi:.1f}", bp_=f"{H['bp_p']:.3f}", rs=f"{R['reset']['p']:.2f}")
(ROOT / "Presentacion_Beamer.tex").write_text(tex, encoding="utf-8")
with zipfile.ZipFile(ROOT / "Beamer_Overleaf.zip", "w", zipfile.ZIP_DEFLATED) as z:
    z.write(ROOT / "Presentacion_Beamer.tex", "Presentacion_Beamer.tex")
    for f in ("pasos_educ", "vif"):
        z.write(FIG / f"{f}.png", f"figuras/{f}.png")
print("Listo: Presentacion_Final.pptx, Presentacion_Beamer.tex, Beamer_Overleaf.zip")
