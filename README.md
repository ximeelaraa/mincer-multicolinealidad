# Multicolinealidad en la ecuación de salarios de Mincer

Proyecto para el curso de **Fundamentos de Econometría**. Usa datos reales para diagnosticar y corregir la multicolinealidad en una ecuación de salarios.

## Datos
`wage2` de Wooldridge (*Introductory Econometrics*): 935 hombres trabajadores de la *National Longitudinal Survey* (EE. UU., 1980), con salario, educación, experiencia, antigüedad, edad, IQ, KWW y escolaridad de los padres. Se carga directamente con el paquete [`wooldridge`](https://pypi.org/project/wooldridge/).

## Entregables
- [`Reporte_Multicolinealidad_Mincer.docx`](Reporte_Multicolinealidad_Mincer.docx): reporte escrito (9 páginas)
- [`Presentacion_Multicolinealidad_Mincer.pptx`](Presentacion_Multicolinealidad_Mincer.pptx): presentación (12 diapositivas)
- [`multicolinealidad_mincer.ipynb`](multicolinealidad_mincer.ipynb): análisis completo con código
- [`figuras/`](figuras/): gráficas en alta resolución

Para regenerar todo: `python scripts/build_notebook.py multicolinealidad_mincer.ipynb` y `python scripts/generar_entregables.py`.

## Contenido del análisis ([`multicolinealidad_mincer.ipynb`](multicolinealidad_mincer.ipynb))
1. Estadística descriptiva y matriz de correlaciones
2. **Multicolinealidad perfecta**: la experiencia potencial `age − educ − 6` hace que X'X sea singular
3. **Multicolinealidad imperfecta**: VIF, número de condición, Farrar-Glauber, regla de Klein, contraste entre las t individuales y la F conjunta
4. Soluciones:
   - centrar `exper` (el VIF baja de 24 a 2 sin cambiar el R²)
   - eliminar o combinar variables redundantes
   - componentes principales
   - regresión Ridge
5. Validación del modelo final: Breusch-Pagan, White, errores robustos HC1, RESET, Jarque-Bera
6. Comparación de modelos (R², R² ajustado, AIC, BIC, VIF máximo)

## Resultados principales
| Modelo   | β educ | R² aj. | VIF máx |
|----------|-------:|-------:|--------:|
| Base     | 0.078  | 0.137  | 20.4    |
| Completo | 0.041  | 0.251  | 24.0    |
| Centrado | 0.041  | 0.251  | 2.2     |
| Reducido | 0.048  | 0.249  | 1.8     |
| PCA      | 0.047  | 0.251  | 1.9     |

## Cómo correrlo
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
jupyter notebook multicolinealidad_mincer.ipynb
```

## Referencias
- Wooldridge, J. M. *Introducción a la econometría: un enfoque moderno*.
- Gujarati, D. y Porter, D. *Econometría*, cap. 10.
- Mincer, J. (1974). *Schooling, Experience, and Earnings*. NBER.
