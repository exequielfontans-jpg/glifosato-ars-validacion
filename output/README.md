# Resultados

Una subcarpeta por script. Cada una contiene el informe `reporte_<nombre>.txt`
de la última corrida, que no lleva número de versión, y las salidas versionadas
de ese script.

## Qué se publica y qué no

La base bibliográfica de la que parte el análisis proviene de una descarga de la
Web of Science sujeta a la licencia de Clarivate. Eso obliga a distinguir dos
clases de salida.

**Se publican** los informes de texto de cada corrida, las planillas auditables
que sostienen las tablas del artículo, los perfiles y centralidades por grupo, y
la red de coautoría del expediente, cuyos nodos provienen de las listas públicas
del solicitante:

```
comparacion/   solapamiento entre el sitio del consorcio y el expediente
publit/        referencias extraidas de los PDF del expediente
gtf/           corpus cosechado del sitio del consorcio
cruce_publit/  estandarizacion_auditoria_v1.xlsx  + reportes
validacion/    tabla4_y_recalculo_v4.xlsx         + reporte
red/           perfiles, centralidades, nodos, aristas y .gexf
```

**No se publican** los cruces a nivel de registro —`cruce_publit_v*`,
`cruce_autores_publit_v*`, `cruce_gtf_wos_v*` y los listados `autores_*`—,
porque reproducen campos de la base licenciada referencia por referencia. Están
excluidos en `.gitignore` y se regeneran corriendo los pasos 4 y 5 con acceso a
la base.

Esa distinción no afecta la verificabilidad de lo que el artículo reporta: cada
tabla tiene su respaldo en las planillas auditables y en los informes, y la
correspondencia entre tabla y salida está en el README principal.
