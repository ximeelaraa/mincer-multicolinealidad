# Multicolinealidad en la ecuación de salarios de Mincer

Proyecto para **Fundamentos de Econometría**. Usa datos reales para estimar el retorno a la educación y cubre todos los temas del segundo parcial. El hilo conductor es la multicolinealidad.

## Entregables
| Archivo | Qué es |
|---|---|
| [`multicolinealidad_mincer.ipynb`](multicolinealidad_mincer.ipynb) | Análisis completo con código y resultados (74 celdas, corre en Colab) |
| [`Reporte_Multicolinealidad_Mincer.docx`](Reporte_Multicolinealidad_Mincer.docx) | Reporte escrito (15 páginas) |
| [`Presentacion_Multicolinealidad_Mincer.pptx`](Presentacion_Multicolinealidad_Mincer.pptx) | Presentación en PowerPoint (20 diapositivas) |
| [`Presentacion_Beamer.tex`](Presentacion_Beamer.tex) | Presentación en LaTeX Beamer (28 frames) |
| [`Beamer_Overleaf.zip`](Beamer_Overleaf.zip) | El `.tex` y sus figuras, listo para subir a Overleaf |
| [`resultados.json`](resultados.json) | Todos los números que usan el reporte y las presentaciones |
| [`figuras/`](figuras/) | Gráficas en alta resolución |

## Datos reales
| Base | Fuente | Uso |
|---|---|---|
| `wage2` | Wooldridge (NLS, EE. UU., 1980), 722 hombres | Base principal |
| `hprice1` | Wooldridge, 88 casas | 4.6 Precios hedónicos |
| `px.data.stocks()` | Plotly, rendimientos semanales 2018–2019 | 4.9 Finanzas |
| `macrodata` | statsmodels, EE. UU. trimestral 1959–2009 | 5.7–5.12 Autocorrelación, dinámica y rezagos |

## Temas cubiertos
- **Regresión lineal simple:** Sxx, Sxy, β̂ = Sxy/Sxx, R² = ρ̂², ecuaciones normales, superficie SCR, reescalamiento, dummy como diferencia de medias, observaciones influyentes, Gauss–Markov.
- **4.1** Supuestos · **4.2** Parámetros y formas funcionales · **4.3** Pruebas t y F, ANOVA y restricciones · **4.4** Dummies, interacciones y trampa de las dummies · **4.5** R² y R² ajustada · **4.6** Precios hedónicos · **4.7** Modelos no anidados (AIC, BIC, prueba J) · **4.8** Regresión cuantílica · **4.9** Beta de mercado y modelo de factores.
- **Multicolinealidad:** perfecta e imperfecta, correlaciones, VIF, tolerancia, regla de Klein, autovalores, índices y número de condición, Farrar–Glauber. Correcciones: centrar, eliminar o combinar, PCA y Ridge.
- **MANOVA:** Wilks y Pillai.
- **5.1–5.2** Residuos, Q–Q, leverage y Cook · **5.3–5.6** Breusch–Pagan manual (nR²), White, HC0–HC3 y WLS · **5.7–5.10** Durbin–Watson, Breusch–Godfrey y HAC · **5.11–5.12** ARDL, multiplicador de largo plazo y rezagos distribuidos (AIC/BIC) · **5.13** Jarque–Bera manual y RESET.
- **Predicción:** train/test y Ridge.

## Cómo reproducir
```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python scripts/analisis.py              # figuras + resultados.json
python scripts/generar_entregables.py   # Word, PowerPoint y Beamer
jupytext --to ipynb --output multicolinealidad_mincer.ipynb scripts/analisis.py   # notebook
```
En **Google Colab**, sube el `.ipynb` y ejecútalo; la primera celda instala `wooldridge`.

Para el **Beamer**, sube `Beamer_Overleaf.zip` a Overleaf (*New Project → Upload Project*) y compila con pdfLaTeX.
