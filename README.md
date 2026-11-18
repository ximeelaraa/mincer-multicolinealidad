# Escolaridad, habilidad y salario: ¿la multicolinealidad impide medir la asociación?

Proyecto de **Fundamentos de Econometría** (Universidad Panamericana, Prof. César Galindo).
**Equipo 4:** Ximena Lara, Camila Ricardi y Sara Enriquez.

**Pregunta:** ¿cómo se asocia la escolaridad con el salario mensual de hombres trabajadores en EE. UU. (1980) cuando se comparan personas con igual experiencia, antigüedad, habilidad medida y antecedentes familiares? ¿Qué tan precisamente podemos medir esa asociación si esos controles están correlacionados entre sí?

**Respuesta corta:** un año más de escolaridad se asocia con un salario ≈4.2% mayor (IC 95% robusto: 2.2–6.0 log-puntos), contra 6.0% sin controles. La multicolinealidad es moderada: aumenta la varianza, pero no impide responder la pregunta. Es una asociación, no un efecto causal.

## Entregables (según la guía del proyecto)
| Entregable de la guía | Archivo |
|---|---|
| 1. **Informe** (5–8 páginas) | [`Informe_Proyecto_Econometria.pdf`](Informe_Proyecto_Econometria.pdf) (7 páginas) · [`.docx`](Informe_Proyecto_Econometria.docx) |
| 2. **Anexo reproducible** | [`multicolinealidad_mincer.ipynb`](multicolinealidad_mincer.ipynb) (corre en Colab) · [`scripts/`](scripts/) · sección *Cómo reproducir* (abajo) |
| 3. **Datos documentados** | [`data/wage2_muestra.csv`](data/wage2_muestra.csv) (archivo utilizado) · [`data/DICCIONARIO.md`](data/DICCIONARIO.md) (fuente, fecha de consulta, diccionario y transformaciones) |
| 4. **Presentación** (8 min + 4 de preguntas) | **[`Presentacion_Beamer.pdf`](Presentacion_Beamer.pdf)** (11 diapositivas, la que se proyecta) · versión editable idéntica: [`Presentacion_Final.pptx`](Presentacion_Final.pptx) / [`.pdf`](Presentacion_Final.pdf), con el guion en las notas del orador |
| Material complementario | [`Anexo_Tecnico_Temas_del_Curso.pdf`](Anexo_Tecnico_Temas_del_Curso.pdf): el resto de los temas del parcial aplicados a los mismos datos |
| Para estudiar (no se entrega) | [`Guia_de_Estudio_y_Exposicion.pdf`](Guia_de_Estudio_y_Exposicion.pdf): explicación simple, conceptos, números clave, guion por diapositiva y preguntas probables |

La presentación tiene **11 diapositivas** para los 8 minutos (unos 45 segundos cada una). [`Beamer_Overleaf.zip`](Beamer_Overleaf.zip) trae el `.tex` y sus figuras, listo para Overleaf.

## Lista de verificación de la guía
- [x] La pregunta y la unidad de análisis son explícitas (informe, secciones 1 y 3).
- [x] Fuentes, variables, unidades y limpieza documentadas (`data/DICCIONARIO.md`, informe, sección 3).
- [x] La ecuación corresponde a los resultados presentados (informe, sección 4 y Tabla 2).
- [x] Coeficientes interpretados en unidades comprensibles (%, USD/mes, por desviación estándar).
- [x] Se discuten la incertidumbre (IC robustos), los supuestos y los límites de la inferencia (secciones 6 y 7).
- [x] Correlaciones y VIF evaluados y contextualizados, sin reglas mecánicas (sección 6.2).
- [x] Código, datos y tablas reproducibles (`scripts/`, notebook, `resultados.json`).
- [ ] Todas las integrantes pueden explicar el trabajo y sus decisiones (repasar juntas antes de exponer).

## Contribuciones del equipo
La guía pide registrar las contribuciones. Cada integrante llena su parte:

| Integrante | Tareas |
|---|---|
| Ximena Lara | |
| Camila Ricardi | |
| Sara Enriquez | |

## Datos
| Base | Fuente | Uso |
|---|---|---|
| `wage2` | NLS Young Men, EE. UU. 1980 (Blackburn y Neumark, 1992), vía Wooldridge; 722 hombres | **Base del proyecto** |
| `hprice1` | Wooldridge, 88 casas | Anexo: precios hedónicos |
| `px.data.stocks()` | Plotly, rendimientos semanales 2018–2019 | Anexo: finanzas |
| `macrodata` | statsmodels, EE. UU. trimestral 1959–2009 | Anexo: autocorrelación y rezagos |

## Cómo reproducir
```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
MPLBACKEND=Agg python scripts/analisis.py      # figuras, resultados.json y data/wage2_muestra.csv
python scripts/generar_informe.py              # informe (Word)
python scripts/generar_entregables.py          # anexo técnico (Word)
python scripts/generar_presentacion.py         # presentación (PowerPoint y Beamer)
python scripts/generar_guia.py                 # guía de estudio (Word)
jupytext --to ipynb --output multicolinealidad_mincer.ipynb scripts/analisis.py   # notebook
```
En **Google Colab**, sube el `.ipynb` y ejecútalo; la primera celda instala `wooldridge`.
Para el **Beamer**, sube `Beamer_Overleaf.zip` a Overleaf (*New Project → Upload Project*) y compila; funciona con pdfLaTeX o XeLaTeX.
