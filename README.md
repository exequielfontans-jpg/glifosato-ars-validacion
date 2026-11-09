# Validación de un procedimiento de atribución estructural de vinculación corporativa

Código y resultados derivados del artículo que evalúa la validez del procedimiento
de atribución estructural propuesto en Fontans-Álvarez et al. (2018), aplicado a
la lista de literatura publicada del expediente europeo de renovación del
glifosato de 2012.

El procedimiento asigna vinculación corporativa a partir de la estructura de la
red de coautoría: la **Zona 1** reúne a los autores que declararon afiliación a
una de las empresas consideradas y a quienes firmaron ese mismo artículo junto a
ellos; la **Zona 2**, al resto. Una vez que un autor ingresa a la Zona 1 se
considera vinculada la totalidad de su obra. Este repositorio permite reproducir
la evaluación de ese criterio contra un conjunto de validación construido con
avisos editoriales publicados.

## Autores

- Exequiel Fontans-Álvarez — Departamento Tratamiento y Transferencia de
  Información, Facultad de Información y Comunicación, Universidad de la
  República — [0000-0002-4627-4678](https://orcid.org/0000-0002-4627-4678)
- Beatriz Sosa — Instituto de Ecología y Ciencias Ambientales, Facultad de
  Ciencias, Universidad de la República —
  [0000-0002-9259-0516](https://orcid.org/0000-0002-9259-0516)

## Estructura

```
scripts/     código del pipeline, una etapa por script
data/        fuentes de entrada (ver data/README.md: no todas se redistribuyen)
output/      resultados reconstruibles (ver output/README.md)
```

Cada script escribe en su propia subcarpeta de `output/` y genera un informe
`reporte_<nombre>.txt` con los estadísticos de la corrida.

## Pipeline

| Orden | Script | Qué hace | Salida principal |
|---|---|---|---|
| 1 | `extraer_publit_v3.py` | Extrae las referencias de los PDF de las listas del expediente | `output/publit/publit_v3.xlsx` |
| 2 | `extraer_corpus_gtf_v2.py` | Cosecha el sitio del Glyphosate Task Force desde Internet Archive y reconcilia DOI contra Crossref | `output/gtf/corpus_gtf_v2.csv` |
| 3 | `comparar_gtf_publit_v1.py` | Mide el solapamiento entre el sitio del consorcio y el expediente | `output/comparacion/solapamiento_gtf_publit_v1.xlsx` |
| 4 | `cruzar_publit_wos_v2.py` | Aparea las referencias del expediente contra la base de 2018 y estandariza por área temática | `output/cruce_publit/cruce_publit_v2.xlsx` |
| 5 | `cruzar_autores_publit_v5.py` | Cruza a nivel de autor y clasifica los artículos por zona | `output/cruce_autores/cruce_autores_publit_v5.xlsx` |
| 6 | `red_coautoria_publit_v4.py` | Construye la red de coautoría del expediente y calcula medidas de centralidad | `output/red/red_publit_v4.gexf` |
| 7 | `auditar_estandarizacion_v1.py` | Audita la estandarización indirecta: población, exclusiones, ponderación y fórmula | `output/cruce_publit/estandarizacion_auditoria_v1.xlsx` |
| 8 | `validar_conjunto_documental_v4.py` | Construye el conjunto de validación documental y mide la sensibilidad del criterio | `output/validacion/tabla4_y_recalculo_v4.xlsx` |

Los pasos 1 a 3 corren sin la base bibliográfica. Los pasos 4 a 8 la requieren.
El paso 8 toma además la hoja `Instancias` del paso 5; sin ella omite las tablas
que dependen de la red de autores y sigue adelante.

`PIPELINE.md` documenta cada etapa en detalle: decisiones de diseño, criterios de
apareo y limitaciones conocidas.

## Reproducir

```bash
pip install -r requirements.txt
cd scripts
python extraer_publit_v3.py
python cruzar_publit_wos_v2.py
python cruzar_autores_publit_v5.py
python red_coautoria_publit_v4.py
python auditar_estandarizacion_v1.py
python validar_conjunto_documental_v4.py
```

Los scripts resuelven sus rutas respecto de su propia ubicación, de modo que
pueden correrse desde cualquier directorio.

## Correspondencia con las tablas del artículo

| Tabla | Script | Hoja o archivo |
|---|---|---|
| Tabla 1, cobertura de la extracción | 1 | `reporte_extraer_publit_v3.txt` |
| Tabla 2, apareo contra la base | 4 | `cruce_publit_v2.xlsx`, hoja `Resumen` |
| Tabla 3, comparación estandarizada | 7 | `estandarizacion_auditoria_v1.xlsx`, hoja `4_Tabla3_articulo` |
| Tabla 4, conjunto de validación | 8 | `tabla4_y_recalculo_v4.xlsx`, hoja `1_Tabla4` |
| Tabla 5, posición en la red por grupo | 6 | `perfil_grupos_publit_v4.csv` |
| Tabla 6, perfil de los autores documentados | 8 | `tabla4_y_recalculo_v4.xlsx`, hoja `6_Tabla6` |
| Tabla 7, vínculo no detectado | 8 | `tabla4_y_recalculo_v4.xlsx`, hoja `4_Tabla7` |
| Tabla 8, anexo documental | 8 | `tabla4_y_recalculo_v4.xlsx`, hoja `5_Tabla8_anexo` |
| Trayectorias en la base (apartado 4.4) | 8 | `tabla4_y_recalculo_v4.xlsx`, hoja `7_Trayectorias` |
| Material suplementario | 8 | `tabla4_y_recalculo_v4.xlsx`, hoja `3_Articulos_documentados` |

## Fuentes del conjunto de validación

El conjunto se define en la sección `CONFIGURACIÓN` de
`validar_conjunto_documental_v4.py`, con la vinculación de cada autor
transcripta del aviso editorial correspondiente:

- Correction (2018). *Critical Reviews in Toxicology*. <https://doi.org/10.1080/10408444.2018.1522175>
- Expression of Concern (2018a), 26 de septiembre. <https://doi.org/10.1080/10408444.2018.1522786>
- Expression of Concern (2018b), 30 de noviembre. <https://doi.org/10.1080/10408444.2018.1539570>
- Retraction notice (2026). *Regulatory Toxicology and Pharmacology*, 165, 106006. <https://doi.org/10.1016/j.yrtph.2025.106006>

La documentación desclasificada en el litigio contra Monsanto no se emplea para
constituir el conjunto, por razones de trazabilidad explicadas en el artículo.

## Licencia

El código de `scripts/` se distribuye bajo licencia MIT (ver `LICENSE`). Los
resultados derivados de `output/` se distribuyen bajo
[CC BY 4.0](https://creativecommons.org/licenses/by/4.0/). Las fuentes de
terceros incluidas en `data/` conservan la licencia de sus titulares.

## Cómo citar

Ver `CITATION.cff`. Para citar una versión concreta, usar el DOI que Zenodo
asigna a cada publicación archivada desde este repositorio.
