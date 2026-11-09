# Fuentes de entrada

Esta carpeta contiene las entradas del pipeline. No todas se redistribuyen en
el repositorio: algunas están sujetas a licencia de terceros. Abajo se indica
qué se incluye, qué no, y cómo obtener lo que falta.

## Se incluye

### `GRG/`

Las listas de literatura publicada que el solicitante presentó en el expediente
de renovación, en PDF. Son documentos públicos, difundidos por el propio
Glyphosate Renewal Group:

<https://www.glyphosate.eu/es/transparency/scientific-dossier/public-literature/>

- `2012_PubLit_Referencelist_Toxicology_Section_03.pdf`
- `2012_PubLit_Referencelist_Metabolism-and-Residue_Section_04.pdf`
- `2012_PubLit_Referencelist_Environmental-Fate_Section_05.pdf`
- `2012_PubLit_Referencelist_Ecotoxicology_Section_06.pdf`

Las cuatro listas de 2020 están en la misma carpeta. No se usan en el análisis:
su presentación consigna solo el primer autor seguido de *et al.*, sin DOI,
categoría ni marca de evaluación, de modo que no admite apareo por autor.

En `GRG/No usados/` quedan tres listas adicionales del mismo sitio que no
corresponden a la presentación analizada.

## No se incluye

### `Descargas WOS/`

La descarga de la Web of Science que sostiene la base bibliográfica de 2018
(8.174 registros, 1974-2016) está sujeta a la licencia de Clarivate y no puede
redistribuirse. Para reconstruirla se necesita acceso institucional a la Web of
Science y la estrategia de búsqueda original, que está documentada en
`Descargas WOS/Estrategia de busqueda - 20161104.txt`.

La descarga se realizó el 4 de noviembre de 2016. Las cifras de citación del
análisis corresponden a ese momento y no a la fecha de consulta actual, lo que
se declara explícitamente en el artículo.

### `Glif_-_Datos_ARS.xlsx`

Base de trabajo derivada de la descarga anterior. Contiene dos hojas que el
pipeline necesita:

- `Base`: un registro por publicación, con los campos de la Web of Science.
- `AU Gephi`: 17.523 autores con su clasificación en Zona 1 o Zona 2.

Al derivarse de la descarga de la Web of Science, hereda su licencia. La hoja
`AU Gephi` sí puede compartirse de forma independiente, porque la clasificación
en zonas es resultado propio; para pedirla, escribir a los autores.

### `Rossi 2020 - Antología toxicológica del glifosato.pdf`

Compilación de terceros, presente en la carpeta de trabajo pero no utilizada en
este análisis. Queda anotada como línea de trabajo futura.

## Para quien quiera reproducir sin acceso a la Web of Science

Los scripts 1 a 3 corren solo con los PDF del expediente y no necesitan la base
bibliográfica. Reproducen la extracción de las referencias y el solapamiento con
el sitio del Glyphosate Task Force.

Los scripts 4 a 8 requieren la base. Las salidas derivadas que los sostienen
están publicadas en `output/`, de modo que los recuentos, las tasas por área y
las tablas del artículo pueden auditarse sin reconstruir el corpus de origen.
