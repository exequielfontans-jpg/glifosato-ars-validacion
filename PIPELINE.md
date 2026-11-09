# Pipeline de procesamiento

**Proyecto.** Evaluación de la validez de la atribución estructural de vinculación
corporativa en redes de coautoría: el caso del glifosato.
FIC-UDELAR. RLARS 2026.

---

## 1. Estructura de carpetas

```
proyecto/
  script/     los .py
  data/       fuentes originales, NUNCA se escribe aquí
  output/     todo lo que generan los scripts, una subcarpeta por script
  Borrador/   versiones del artículo
```

Regla de diseño: si se borra `output/` y se vuelven a correr los scripts en
orden, el resultado debe ser idéntico. Por eso las salidas no se guardan en
`data/`.

Todos los scripts derivan sus rutas de `Path(__file__).resolve().parent` y
funcionan desde cualquier directorio de trabajo. Al arrancar imprimen las rutas
de `data` y `output` que van a usar.

---

## 2. Fuentes de datos

| Archivo | Contenido | Origen |
|---|---|---|
| `data/Glif_-_Datos_ARS.xlsx` | Hoja `Base`: 8.174 registros WoS (1974-2016), 63 campos. Hoja `AU Gephi`: 17.523 autores normalizados con Zona 1/2 | Fontans-Álvarez et al. (2018) |
| `data/Glif_AU_Norm.xlsx` | Hoja `AUNorm`: 32.481 instancias autor-artículo, con campos `AF` (nombre completo) y `C1` (afiliación) | Ídem |
| `data/Descargas WOS/` | Exportaciones originales de WoS en bloques de 500 registros, más la estrategia de búsqueda | Descarga del 04/11/2016 |
| `data/GRG/*.pdf` | 8 listas de literatura abierta del expediente europeo (4 de 2012, 4 de 2020) | glyphosate.eu |
| `data/Rossi 2020 - Antología...pdf` | Corpus curado con lógica inversa | Naturaleza de Derechos |

Nota sobre la citación: todas las cifras de citas corresponden al 4 de
noviembre de 2016, fecha de descarga del corpus. Se mantiene esa ventana por
comparabilidad entre trabajos.

---

## 3. Secuencia de ejecución

Los scripts deben correrse en este orden. Cada uno consume salidas del
anterior.

### 3.1 `extraer_publit_v3.py` — corpus del expediente

Entrada: `data/GRG/*.pdf`
Salida: `output/publit/publit_v3.{csv,xlsx}`

Extrae las tablas de los PDF mediante `pdfplumber`, mapeando las columnas por
el nombre del encabezado y no por posición fija. Esto es necesario porque los
PDF de 2012 y 2020 tienen disposiciones distintas: los de 2012 traen seis
columnas (punto de anexo, autores, año, cita, categoría, evaluación) y los de
2020 solo tres.

Decisiones incorporadas:

- Se omiten las páginas de carátula (3 en 2012, 2 en 2020), configurables en
  `PAGINAS_OMITIR`.
- Se recomponen los DOI partidos por el salto de línea del PDF.
- **No se deduplica.** Una referencia puede figurar en varias secciones y
  recibir categorías distintas en cada una; el solapamiento se marca en los
  campos `n_apariciones` y `aparece_en`.
- Toda columna no reconocida se conserva en `columnas_extra`, para que ninguna
  variación de formato produzca pérdida silenciosa.

Resultado: 1.768 referencias, de las cuales 1.078 corresponden a la
presentación de 2012 (4,2 autores por referencia, 63,7% con DOI) y 690 a la de
2020.

La presentación de 2020 se excluye del análisis: sus listas consignan el primer
autor seguido de *et al.*, sin DOI, sin categoría y sin marca de evaluación.

### 3.2 `extraer_corpus_gtf_v2.py` — verificación de la extracción

Entrada: API CDX del Internet Archive
Salida: `output/gtf/corpus_gtf_v2.csv` + HTML crudo en `output/gtf/html/`

Recupera la base de literatura del sitio del consorcio (glifosato.es), caído,
desde las versiones archivadas. Toma el snapshot más temprano de cada página
para reconstruir el corpus en el estado en que sustentó el expediente de 2012.
Reconcilia contra Crossref exigiendo similitud de título y coherencia de año
además del puntaje de búsqueda.

Resultado: 548 referencias únicas en 22 categorías.

**Este corpus no se usa como fuente analítica.** Sirve como verificación
independiente de la extracción de los PDF (ver 3.3).

### 3.3 `comparar_gtf_publit_v1.py` — solapamiento

Entrada: las dos salidas anteriores
Salida: `output/comparacion/solapamiento_gtf_publit_v1.xlsx`

Coteja ambos corpus por DOI y, en su defecto, por título normalizado.

Resultado: 517 de las 548 referencias del sitio (94,3%) están en el expediente.
El sitio es un subconjunto del expediente, no un corpus independiente. La
coincidencia entre dos procedimientos de extracción técnicamente distintos
—cosecha de HTML con expresiones regulares y lectura de tablas de PDF— valida
la extracción.

### 3.4 `cruzar_publit_wos_v2.py` — cruce por artículo

Entrada: `data/Glif_-_Datos_ARS.xlsx` + `output/publit/publit_v3.csv`
Salida: `output/cruce_publit/cruce_publit_v2.{csv,xlsx}`

Aparea las referencias del expediente contra la base de 2018 por DOI, luego por
título normalizado y por último por título difuso acotado por año. El apareo a
nivel de artículo es el único que permite recuperar la categoría temática de
WoS y, con ella, estandarizar.

`SOLO_PRESENTACION = "2012"` acota el análisis a esa presentación.

Resultados:

- 749 de 1.078 apareadas (69,5%): 504 por DOI, 227 por título exacto, 18 difuso.
- 123 de 749 con al menos un autor de Zona 1 (16,4%), contra 31,5% de la
  literatura general 2001-2012 (χ² de Yates = 95,68).
- Estandarizado por área temática: 114 observados contra 159,6 esperados,
  razón 0,71. La hoja `Estandarizacion` expone el cálculo área por área.
- Cobertura desigual del apareo: 81% en destino ambiental, 74% en
  ecotoxicología, 49% en toxicología. Reponderando por secciones, la
  estimación baja de 16,4% a 15,6%.

### 3.5 `cruzar_autores_publit_v3.py` — cruce por autor

Entrada: `data/Glif_-_Datos_ARS.xlsx` + `output/publit/publit_v3.csv`
Salida: `output/cruce_autores/cruce_autores_publit_v3.{csv,xlsx}`

Normaliza los nombres de autor del expediente a la forma `APELLIDO, INICIALES`
y los cruza contra la clasificación de zonas. A diferencia del cruce por
artículo, cubre las 1.078 referencias: no requiere que el artículo esté en la
base, solo que sus autores estén clasificados.

Problemas de normalización resueltos:

- Iniciales partidas por coma (`Solomon K, R.` → `SOLOMON, KR`)
- Nombres de pila completos (`Brain Richard A` → `BRAIN, RA`)
- Iniciales con guion (`Sanin L-H` → `SANIN, LH`)
- Partículas (`de Villalobos C` → `DE VILLALOBOS, C`)
- Iniciales pegadas frente a nombres de pila: un token corto en mayúsculas
  (`KR`) se separa en iniciales; uno largo (`Richard`) aporta solo su inicial.

La ambigüedad se mide, no se supone: se marcan los casos con varias formas
crudas bajo una misma clave, los apellidos presentes en más de ocho autores
distintos de la base, y los que carecen de iniciales.

Asignación de artículos a cuatro grupos, por prioridad:
`CONTROL` > `ZONA_1` > `ZONA_2` > `SIN_INFORMACION`.

El grupo de control se define en la constante `CONTROL_DOCUMENTAL`, con los
autores nominados en la *Expression of Concern* de *Critical Reviews in
Toxicology* (2018) y en la retractación de Williams, Kroes y Munro (2000) por
*Regulatory Toxicology and Pharmacology* (2025).

Resultados:

- 4.549 instancias de autor, 3.118 autores únicos.
- 137 artículos con Zona 1 (12,7%), 13 solo con vinculación documental,
  796 Zona 2, 132 sin información (12,2%).
- De 150 artículos con vínculo, 13 (8,7%) son invisibles al criterio
  estructural. En la literatura general del período la proporción equivalente
  es 0,8%.
- 43 artículos (4,0%) clasificados sobre apellidos frecuentes: requieren
  revisión manual.

### 3.6 `red_coautoria_publit_v4.py` — red y centralidad

Entrada: `output/cruce_autores/...v3.xlsx` (hoja `Instancias`) y
`output/cruce_publit/cruce_publit_v2.csv`
Salida: `output/red/` — nodos, aristas, GEXF, perfil por grupo, centralidad

Construye la red de coautoría del expediente: cada autor un nodo, cada par de
coautores un vínculo no dirigido ponderado por trabajos en común. Calcula
grado, fuerza, clustering local e intermediación de Brandes (2001), esta última
implementada en el propio script para no depender de librerías externas.

Se usa la hoja `Instancias` y no la hoja `Autores` del cruce por artículo,
porque aquélla cubre las 1.078 referencias y trae el campo `grupo` con las
cuatro categorías.

Resultados:

- 3.118 nodos, 8.354 aristas, densidad 0,00172, grado medio 5,36.
- 505 componentes; el mayor reúne 432 nodos (13,9%).
- Intermediación media: Zona 1 = 1.258,7; Zona 2 = 151,4 (ocho veces).
- Intermediación no nula: Zona 1 = 34,2%; Zona 2 = 20,1%.
- Presencia en el componente principal: Zona 1 = 40,5%; Zona 2 = 10,0%.
- Clustering más bajo en Zona 1 (0,813): sus coautores colaboran menos entre sí.

La fragmentación de la red desaconseja la representación gráfica del grafo
completo; se opta por caracterizarla mediante medidas de posición.

---

## 4. Decisiones metodológicas y su fundamento

**Apareo por artículo, no por nombre, cuando es posible.** El apareo por
apellido e inicial genera falsos positivos en apellidos frecuentes, con sesgo
sistemático hacia apellidos asiáticos e hispanos. Cuando el artículo aparea, la
lista de autores viene de la base ya normalizada.

**Estandarización indirecta por área temática.** La tasa de Zona 1 varía de
53,2% en Agronomy a 9,6% en Environmental Sciences. Comparar corpus curados sin
controlar el área produce artefactos con apariencia de hallazgo: la comparación
cruda sobreestimaba la diferencia en unos diez puntos porcentuales.

**Ventana de referencia 2001-2012.** Tanto las tasas de referencia como el
corpus observado se restringen a esa ventana. Sin esa simetría se compararían
poblaciones de períodos distintos (efecto medido: 8 artículos, razón sin
cambios).

**Exclusión de la presentación de 2020.** Por insuficiencia de metadatos, no
por decisión analítica. Su incorporación requeriría reconciliación externa por
título y datos de publicación, viable pero no abordada.

**El grupo de control no admite estadísticos de grupo.** Con seis casos se
reporta la posición individual de cada autor dentro de la distribución del
corpus, no medianas de grupo.

---

## 5. Limitaciones conocidas

- El conjunto de validación se apoya solo en avisos editoriales publicados: es
  conservador por construcción y las proporciones son un piso.
- La clasificación hereda la normalización de nombres de 2018, en la que se
  detectaron unos 127 pares de nodos posiblemente desdoblados sobre 17.523
  autores (menos del 1%), concentrados en apellidos frecuentes. Casos
  verificados: `SOLOMON, KR` / `SOLOMON, K` y `ACQUAVELLA, JF` /
  `ACQUAVELLA, J`.
- El criterio de 2018 no considera la financiación declarada (campos `FU` y
  `FX` de WoS), disponible en el corpus y no utilizada.
- La brecha entre las 548 referencias recuperadas del sitio y las
  aproximadamente 1.000 declaradas no está explicada: puede ser cobertura
  parcial del archivo o tasa de retención del propio proceso de selección.
- El significado de la columna `categoria` del expediente depende de un
  documento de proceso del solicitante (Carr y Bleeke, 2012, informe
  ASB2012-11583) que no se pudo consultar. Se transcribe sin interpretar.

---

## 6. Líneas de trabajo abiertas

**Corpus de Rossi.** La *Antología toxicológica del glifosato* (5.ª ed., 2020)
es un corpus curado con lógica inversa: selecciona por evidencia de daño. El
PDF está en `data/`. Comparar ambas lógicas de curación era el diseño original
y quedó fuera por tiempo.

**Reconstrucción de la presentación de 2020.** Título, revista, año, volumen y
páginas alcanzan para reconciliar contra Crossref o PubMed y recuperar la
autoría completa. Son 690 referencias y el procedimiento ya está implementado
en `extraer_corpus_gtf_v2.py`.

**Extensión del corpus a 2017-2026.** Cubre la adquisición de Monsanto por
Bayer, el desenlace del litigio y la retractación. Permitiría observar si la
estructura de la producción se reconfiguró.

**Financiación declarada como vía de atribución.** Los campos están en el
corpus; incorporarlos podría reducir la tasa de falsos negativos y es testeable
contra el mismo conjunto documental.

**Documentos internos de la industria.** La colección de la UCSF Industry
Documents Library (118 documentos, 1.063 páginas, caso 3:16-md-02741-VC) tiene
identificadores citables y un wrapper de API en Python. Existe además
precedente metodológico de extracción de entidades y análisis de red sobre este
tipo de material (proyecto JUUL, UNC Libraries). Permitiría construir un
conjunto de validación mucho mayor y estimar la sensibilidad con precisión, en
lugar de con nueve casos.

**Composición de comités editoriales.** Cuarta modalidad de vinculación,
inobservable por autoría. Su relevamiento sistemático permitiría estimar su
extensión.

**Clasificación de resultados por orientación.** Guillemaud, Lombaert y
Bourguet (2016) resolvieron cómo hacerlo; es la extensión natural de esta
línea.

---

## 7. Convenciones de los scripts

- Docstring con proyecto, contexto, entradas, salidas, sensibilidad, uso,
  requisitos e historial de versiones.
- Rutas derivadas de la ubicación del archivo; se imprimen al arrancar.
- Informe de texto con las cifras principales, además de las salidas de datos.
- Salidas en Excel con formato para revisión manual: encabezado fijo, filtros,
  anchos de columna, hoja aparte con lo que requiere decisión humana.
- Los CSV se mantienen porque Gephi y UCINET los leen directamente.
- Nombre de versión en el archivo y en las salidas: una corrida nueva no pisa
  la anterior.
