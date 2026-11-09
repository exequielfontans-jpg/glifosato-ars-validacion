"""
extraer_publit_v3.py
====================
Extrae las listas de publicaciones de los expedientes de renovacion europea
del glifosato (presentaciones 2012 y 2020, Glyphosate Renewal Group) desde
los PDF publicados, conservando la categoria de fiabilidad que el solicitante
asigno a cada referencia.

Estructura de carpetas del proyecto:
  proyecto/
    script/     los .py
    data/       fuentes originales, NUNCA se escribe aqui
    output/     todo lo que generan los scripts, una subcarpeta por script

Si se borra output/ y se vuelven a correr los scripts en orden, el resultado
debe ser identico. Por eso las salidas no van en data/.

Proyecto: Corpus curados sobre glifosato — comparacion antologia Rossi /
          Glyphosate Task Force / expediente regulatorio. FIC-UDELAR.
Contexto: Construccion del corpus regulatorio · parseo de PDF


Entradas:
  - data/GRG/*.pdf      - listas de referencias descargadas de glyphosate.eu
                          El nombre del archivo debe contener el anio de
                          presentacion (2012 / 2020) y la seccion.

Salidas:
  - output/publit/publit_v3.xlsx - libro de trabajo (Resumen, Solapamiento,
                                        Referencias)
  - output/publit/publit_v3.csv  - mismas filas, para cruzar con la base WoS
  - output/publit/reporte_extraer_publit.txt
  - output/publit/extraer_publit_errores.log

Sensibilidad de salidas:
  NO SENSIBLE : todas (metadatos bibliograficos)

Uso: python extraer_publit_v3.py

Requisitos: pip install pdfplumber openpyxl

Historial:
  v1 - version inicial: extraccion de tablas por pagina, union de filas
       partidas, deteccion de DOI y separacion titulo / fuente. NO se
       deduplica: una misma referencia puede figurar en varias secciones o
       en ambas presentaciones, y ese solapamiento es dato, no ruido; se
       marca en los campos n_apariciones / aparece_en.
  v2 - corrige el defecto principal de la v1: las columnas se tomaban por
       posicion fija, lo que funciona en los PDF de 2012 pero no en los de
       2020, cuya tabla tiene otra disposicion (el campo 'categoria' recogia
       el nombre de la revista). Ahora el encabezado de cada tabla se lee y
       las columnas se mapean por nombre. El informe lista el mapeo detectado
       en cada archivo para poder verificarlo.
  v3 - los PDF de 2020 traen menos columnas que los de 2012 (sin categoria,
       sin evaluacion, y con titulo y fuente separados). Se agrega una regla
       para la columna de fuente y, sobre todo, un campo 'columnas_extra' que
       guarda el contenido de toda columna no reconocida: asi ninguna
       variacion de formato produce perdida silenciosa de datos.

NOTA sobre el campo 'categoria': se transcribe tal como figura en el PDF.
Su significado NO esta establecido; verificar la leyenda de codigos en la
primera o ultima pagina del PDF original antes de interpretarlo.
"""

import csv
import re
import unicodedata
from pathlib import Path
from datetime import datetime

import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

CARPETA_SCRIPT = Path(__file__).resolve().parent
CARPETA_BASE   = CARPETA_SCRIPT.parent
CARPETA_DATOS  = CARPETA_BASE / "data"           # fuentes originales, solo lectura
CARPETA_OUT    = CARPETA_BASE / "output"         # todo lo generado por scripts
CARPETA_PDFS   = CARPETA_DATOS / "GRG"
CARPETA_SALIDA = CARPETA_OUT / "publit"

XLSX_FILE    = CARPETA_SALIDA / "publit_v3.xlsx"
CSV_FILE     = CARPETA_SALIDA / "publit_v3.csv"
REPORTE_FILE = CARPETA_SALIDA / "reporte_extraer_publit.txt"
LOG_FILE     = CARPETA_SALIDA / "extraer_publit_errores.log"

# Reglas para mapear el encabezado de la tabla a los campos internos. Se
# evaluan en orden y gana la primera que coincida, asi que las mas especificas
# van primero ("evaluation" antes que cualquier regla que capture "tion").
REGLAS_COLUMNA = [
    ("annex_point", ("annex", "reference number", "reference no")),
    ("evaluacion",  ("evaluation", "translation")),
    ("categoria",   ("category", "categories")),
    ("autores",     ("author",)),
    ("anio",        ("year",)),
    ("cita",        ("title", "citation")),
    ("fuente_col",  ("source", "journal", "bibliographic", "publication",
                     "reference")),
]
MARCAS_CABECERA = ("annex point", "author", "title", "citation", "year")

# Paginas de caratula y leyenda antes de que empiece la tabla, por presentacion.
# Los PDF de 2012 traen 3 y los de 2020 traen 2. Se omiten para que el texto
# de la leyenda no genere filas espurias.
PAGINAS_OMITIR = {"2012": 3, "2020": 2}

CAMPOS = ["id_ref", "presentacion", "seccion", "archivo", "pagina",
          "annex_point", "autores", "anio", "titulo", "fuente", "doi",
          "categoria", "evaluacion", "n_apariciones", "aparece_en",
          "clave_dup", "cita_cruda", "columnas_extra"]

RE_DOI  = re.compile(r"10\.\d{4,9}/[^\s,;]+", re.I)
RE_ANIO = re.compile(r"\b(19[5-9]\d|20[0-3]\d)\b")


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def log_error(mensaje):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{ts}] {mensaje}\n")


def escribir_csv(path, filas, campos):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=campos, extrasaction="ignore")
        w.writeheader()
        w.writerows(filas)


def normalizar(txt):
    """Une los saltos de linea del PDF y colapsa espacios."""
    if txt is None:
        return ""
    t = unicodedata.normalize("NFC", str(txt))
    for a, b in [("\u201c", '"'), ("\u201d", '"'), ("\u2019", "'"),
                 ("\u2013", "-"), ("\u2014", "-"), ("\u00a0", " ")]:
        t = t.replace(a, b)
    t = t.replace("-\n", "")          # palabra cortada por guion
    t = t.replace("\n", " ")
    return re.sub(r"\s+", " ", t).strip()


def limpiar_doi(txt):
    """Extrae el DOI de la cita y le saca la puntuacion final.

    Los PDF cortan la linea justo despues de la barra ("10.1016/
    j.chemosphere...") y a veces dentro del sufijo, asi que primero se
    recomponen esos cortes; sin eso se pierde una parte importante de los DOI.
    """
    t = normalizar(txt or "")
    t = re.sub(r"(10\.\d{4,9}/)\s+", r"\1", t)        # corte tras la barra
    t = re.sub(r"(10\.\d{4,9}/\S*?)\s+(?=[a-z0-9._-]{2,})", r"\1", t)
    m = RE_DOI.search(t)
    if not m:
        return ""
    return m.group(0).rstrip(".,;)").lower()


def partir_cita(cita):
    """Separa titulo y fuente. El titulo termina en el primer punto seguido
    de espacio y mayuscula, que es donde arranca el nombre de la revista."""
    c = normalizar(cita)
    c_sin_doi = RE_DOI.sub("", c)
    c_sin_doi = re.sub(r"\b(doi|DOI)\s*:?\s*$", "", c_sin_doi).strip()
    m = re.search(r"\.\s+(?=[A-Z])", c_sin_doi)
    if not m:
        return c_sin_doi.strip(" ."), ""
    return c_sin_doi[:m.start()].strip(" ."), c_sin_doi[m.end():].strip(" .")


def metadatos_archivo(nombre):
    """Deduce presentacion y seccion del nombre del archivo."""
    n = nombre.lower()
    pres = "2020" if "2020" in n or "air5" in n else "2012"
    secciones = {"ecotox": "Ecotoxicología", "efate": "Destino ambiental",
                 "environmental-fate": "Destino ambiental",
                 "residue": "Metabolismo y residuos",
                 "residues": "Metabolismo y residuos",
                 "toxicology": "Toxicología", "metabolism": "Metabolismo y residuos"}
    for clave, valor in secciones.items():
        if clave in n:
            return pres, valor
    return pres, "Sin identificar"


# ─── PROCESAMIENTO ────────────────────────────────────────────────────────────

def es_fila_encabezado(celdas):
    """Una fila es encabezado si menciona al menos dos rotulos conocidos."""
    textos = [re.sub(r"[^a-z]", "", normalizar(c).lower()) for c in celdas]
    aciertos = sum(1 for t in textos
                   if any(m.replace(" ", "") in t for m in MARCAS_CABECERA))
    return aciertos >= 2


def mapear_columnas(celdas):
    """Devuelve {campo: indice} a partir de la fila de encabezado.

    La v1 usaba posiciones fijas (0..5), lo que funciona en los PDF de 2012
    pero no en los de 2020, donde la tabla tiene otra disposicion. Leer el
    encabezado hace el parseo independiente del orden de las columnas.
    """
    mapa = {}
    for i, celda in enumerate(celdas):
        # Los PDF cortan las palabras del encabezado ("Evalu ation / Transl
        # ation"), asi que se compara sobre el texto sin espacios ni signos.
        t = re.sub(r"[^a-z]", "", normalizar(celda).lower())
        if not t:
            continue
        for campo, claves in REGLAS_COLUMNA:
            if campo in mapa:
                continue
            if any(k.replace(" ", "") in t for k in claves):
                mapa[campo] = i
                break
    return mapa


def valor(celdas, mapa, campo):
    i = mapa.get(campo)
    if i is None or i >= len(celdas):
        return ""
    return normalizar(celdas[i])


def procesar_pdf(path, contador, diagnostico):
    """Extrae las referencias de un PDF. Devuelve (filas, n_paginas)."""
    pres, seccion = metadatos_archivo(path.name)
    filas, n_pag = [], 0
    mapa = {}

    try:
        pdf = pdfplumber.open(path)
    except Exception as e:
        log_error(f"No se pudo abrir {path.name}: {repr(e)[:300]}")
        return [], 0

    omitir = PAGINAS_OMITIR.get(pres, 0)
    with pdf:
        n_pag = len(pdf.pages)
        for i, pagina in enumerate(pdf.pages, 1):
            if i <= omitir:
                continue
            try:
                tablas = pagina.extract_tables()
            except Exception as e:
                log_error(f"{path.name} pag {i}: {repr(e)[:200]}")
                continue

            for tabla in tablas:
                for celdas in tabla:
                    if not celdas:
                        continue
                    if es_fila_encabezado(celdas):
                        nuevo_mapa = mapear_columnas(celdas)
                        if len(nuevo_mapa) >= 3:
                            mapa = nuevo_mapa
                        continue
                    if not mapa:
                        continue           # aun no vimos encabezado utilizable

                    cita       = valor(celdas, mapa, "cita")
                    fuente_col = valor(celdas, mapa, "fuente_col")
                    # Toda columna que no se reconocio se conserva textual,
                    # con su indice, para no perder informacion en silencio.
                    usados = set(mapa.values())
                    extra = "; ".join(
                        f"[{i}] {normalizar(c)}"
                        for i, c in enumerate(celdas)
                        if i not in usados and normalizar(c))
                    autores  = valor(celdas, mapa, "autores")
                    anio_txt = valor(celdas, mapa, "anio")
                    annex    = valor(celdas, mapa, "annex_point")

                    # fila sin datos propios: es la cola de la anterior
                    if filas and not autores and not annex and \
                            not RE_ANIO.search(anio_txt):
                        if cita:
                            prev = filas[-1]
                            prev["cita_cruda"] = (prev["cita_cruda"] + " " + cita).strip()
                            prev["titulo"], prev["fuente"] = partir_cita(prev["cita_cruda"])
                            prev["doi"] = prev["doi"] or limpiar_doi(prev["cita_cruda"])
                        continue

                    if len(cita) < 20 or len(autores) < 3:
                        continue

                    completa = (cita + " " + fuente_col).strip()
                    m = RE_ANIO.search(anio_txt) or RE_ANIO.search(completa)
                    if fuente_col:
                        titulo, fuente = cita.strip(" ."), fuente_col
                    else:
                        titulo, fuente = partir_cita(cita)
                    contador[0] += 1

                    filas.append({
                        "id_ref": f"PL{contador[0]:05d}",
                        "presentacion": pres,
                        "seccion": seccion,
                        "archivo": path.name,
                        "pagina": i,
                        "annex_point": annex,
                        "autores": autores,
                        "anio": m.group(0) if m else "",
                        "titulo": titulo,
                        "fuente": fuente,
                        "doi": limpiar_doi(completa),
                        "categoria": valor(celdas, mapa, "categoria"),
                        "evaluacion": valor(celdas, mapa, "evaluacion"),
                        "n_apariciones": 1,
                        "aparece_en": "",
                        "clave_dup": "",
                        "cita_cruda": completa,
                        "columnas_extra": extra,
                    })

    diagnostico.append({
        "archivo": path.name, "presentacion": pres, "seccion": seccion,
        "paginas": max(n_pag - omitir, 0), "referencias": len(filas),
        "columnas_detectadas": ", ".join(f"{k}={v}" for k, v in sorted(
            mapa.items(), key=lambda x: x[1])) or "NINGUNA",
    })
    return filas, n_pag


def marcar_solapamiento(filas):
    """Marca las referencias que aparecen en mas de un archivo.

    No se elimina ninguna fila: que un articulo figure en varias secciones o
    en las dos presentaciones es justamente lo que se quiere poder medir.
    La clave es el DOI cuando existe y, si no, el titulo normalizado.
    """
    grupos = {}
    for f in filas:
        clave = f["doi"] or re.sub(r"[^a-z0-9 ]+", " ",
                                   f["titulo"].lower())[:90].strip()
        f["clave_dup"] = clave
        grupos.setdefault(clave, []).append(f)

    for clave, grupo in grupos.items():
        if not clave:
            continue
        origenes = sorted({f"{f['presentacion']}/{f['seccion']}" for f in grupo})
        for f in grupo:
            f["n_apariciones"] = len(grupo)
            f["aparece_en"] = "; ".join(origenes)
    return grupos


# ─── SALIDA EN EXCEL ──────────────────────────────────────────────────────────

GRIS = PatternFill("solid", fgColor="D9D9D9")
ANCHOS = {"id_ref": 10, "presentacion": 12, "seccion": 22, "archivo": 30,
          "pagina": 8, "annex_point": 12, "autores": 40, "anio": 7,
          "titulo": 62, "fuente": 40, "doi": 26, "categoria": 10,
          "evaluacion": 11, "n_apariciones": 12, "aparece_en": 42,
          "clave_dup": 30, "cita_cruda": 70, "columnas_extra": 50}
AMARILLO = PatternFill("solid", fgColor="FFF2CC")


def escribir_excel(filas, path, diag):
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Resumen"
    n = len(filas)
    cats = sorted({f["categoria"] for f in filas if f["categoria"]})
    pres = sorted({f["presentacion"] for f in filas})

    ws.append(["EXTRACCION DE LAS LISTAS DE PUBLICACIONES"])
    ws.append([])
    ws.append(["Referencias extraidas", n])
    ws.append(["Con DOI", sum(1 for f in filas if f["doi"])])
    ws.append(["Filas repetidas entre archivos",
               sum(1 for f in filas if f["n_apariciones"] > 1)])
    ws.append(["Referencias unicas (por clave)",
               len({f["clave_dup"] for f in filas if f["clave_dup"]})])
    ws.append([])
    ws.append(["Por presentacion y categoria de fiabilidad"])
    ws.append(["Presentacion"] + [f"Categoria {c}" for c in cats] + ["Total"])
    for p in pres:
        sub = [f for f in filas if f["presentacion"] == p]
        ws.append([p] + [sum(1 for f in sub if f["categoria"] == c) for c in cats]
                  + [len(sub)])
    ws.append([])
    ws.append(["Por seccion"])
    for s in sorted({f["seccion"] for f in filas}):
        ws.append([s, sum(1 for f in filas if f["seccion"] == s)])

    ws.column_dimensions["A"].width = 34
    for i in range(2, 10):
        ws.column_dimensions[get_column_letter(i)].width = 13
    for fila in ws.iter_rows():
        for c in fila:
            c.font = Font(name="Arial")
    for r in (1, 6, 7, 10):
        ws.cell(row=r, column=1).font = Font(name="Arial", bold=True)

    wsd = wb.create_sheet("Diagnostico")
    wsd.append(["archivo", "presentacion", "seccion", "paginas",
                "referencias", "columnas_detectadas"])
    for c in wsd[1]:
        c.font = Font(name="Arial", bold=True)
        c.fill = GRIS
    for d in diag:
        wsd.append([d["archivo"], d["presentacion"], d["seccion"],
                    d["paginas"], d["referencias"], d["columnas_detectadas"]])
    for i, w in enumerate([46, 13, 22, 10, 12, 54], 1):
        wsd.column_dimensions[get_column_letter(i)].width = w
    for fila in wsd.iter_rows(min_row=2):
        for c in fila:
            c.font = Font(name="Arial")

    rep = [f for f in filas if f["n_apariciones"] > 1]
    ws2 = wb.create_sheet("Solapamiento")
    ws2.append(["clave_dup", "n_apariciones", "aparece_en", "titulo",
                "categorias_asignadas"])
    for c in ws2[1]:
        c.font = Font(name="Arial", bold=True)
        c.fill = GRIS
    vistos = set()
    for f in sorted(rep, key=lambda x: (-x["n_apariciones"], x["clave_dup"])):
        if f["clave_dup"] in vistos:
            continue
        vistos.add(f["clave_dup"])
        cats = sorted({g["categoria"] for g in rep
                       if g["clave_dup"] == f["clave_dup"] and g["categoria"]})
        ws2.append([f["clave_dup"], f["n_apariciones"], f["aparece_en"],
                    f["titulo"][:180], "; ".join(cats)])
    for i, w in enumerate([28, 14, 44, 62, 20], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    for fila in ws2.iter_rows(min_row=2):
        for c in fila:
            c.font = Font(name="Arial")
    if len(vistos):
        ws2.auto_filter.ref = f"A1:E{len(vistos) + 1}"
    ws2.freeze_panes = "A2"

    ws = wb.create_sheet("Referencias")
    ws.append(CAMPOS)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True)
        c.fill = GRIS
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for f in filas:
        ws.append([f.get(k, "") for k in CAMPOS])
    for i, campo in enumerate(CAMPOS, 1):
        ws.column_dimensions[get_column_letter(i)].width = ANCHOS.get(campo, 14)
    col_n = CAMPOS.index("n_apariciones")
    for fila in ws.iter_rows(min_row=2):
        for c in fila:
            c.font = Font(name="Arial")
            c.alignment = Alignment(vertical="top")
        if (fila[col_n].value or 1) > 1:      # repetida entre archivos
            for c in fila:
                c.fill = AMARILLO
    if filas:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(CAMPOS))}{len(filas) + 1}"
    ws.freeze_panes = "B2"

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# ─── INFORME ──────────────────────────────────────────────────────────────────

def escribir_informe(filas, archivos, paginas, path, diag):
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    n = len(filas)
    pct = lambda x: f"{(100 * x / n):.1f}%" if n else "-"
    con_doi = sum(1 for f in filas if f["doi"])
    sin_cat = sum(1 for f in filas if not f["categoria"])

    lineas = [
        "=" * 64,
        "  [PUBLIT] - extraer_publit_v1",
        f"  Fecha: {ahora}",
        "=" * 64,
        "",
        f"  PDF procesados      : {archivos}",
        f"  Paginas de tabla    : {paginas}",
        f"  Referencias         : {n:,}",
        f"  Con DOI             : {con_doi:,} ({pct(con_doi)})",
        f"  Sin categoria       : {sin_cat:,} ({pct(sin_cat)})",
        "",
        "  Por presentacion:",
    ]
    for p in sorted({f["presentacion"] for f in filas}):
        k = sum(1 for f in filas if f["presentacion"] == p)
        lineas.append(f"    {p:<10} {k:>6}  ({pct(k)})")

    lineas += ["", "  Por seccion:"]
    for s in sorted({f["seccion"] for f in filas}):
        k = sum(1 for f in filas if f["seccion"] == s)
        lineas.append(f"    {s:<28} {k:>6}  ({pct(k)})")

    lineas += ["", "  Por categoria (transcripta del PDF, sin interpretar):"]
    cuenta = {}
    for f in filas:
        cuenta[f["categoria"]] = cuenta.get(f["categoria"], 0) + 1
    for c, k in sorted(cuenta.items(), key=lambda x: -x[1])[:15]:
        lineas.append(f"    {(c or '(vacia)')[:44]:<46} {k:>6}  ({pct(k)})")
    if len(cuenta) > 15:
        lineas.append(f"    ... y {len(cuenta) - 15} valores mas "
                      f"(ver hoja Referencias del xlsx)")

    lineas += ["", "  Marca de evaluacion:"]
    for marca in ("K", "T", "K T", ""):
        k = sum(1 for f in filas if f["evaluacion"].strip() == marca)
        if k:
            lineas.append(f"    {(marca or '(sin marca)'):<12} {k:>6}  ({pct(k)})")

    con_extra = sum(1 for f in filas if f.get("columnas_extra"))
    if con_extra:
        lineas += ["",
                   f"  Filas con columnas no reconocidas : {con_extra:,} "
                   f"({pct(con_extra)})",
                   "    Ver el campo 'columnas_extra' para decidir si aportan algo."]

    lineas += ["", "  Columnas detectadas por archivo:"]
    for d in diag:
        lineas.append(f"    {d['archivo'][:44]:<46} {d['referencias']:>5} refs")
        lineas.append(f"      -> {d['columnas_detectadas']}")

    rep = sum(1 for f in filas if f["n_apariciones"] > 1)
    uni = len({f["clave_dup"] for f in filas if f["clave_dup"]})
    lineas += ["",
               f"  Filas repetidas entre archivos : {rep:,} ({pct(rep)})",
               f"  Referencias unicas (por clave) : {uni:,}",
               "",
               "  NOTA: el campo 'categoria' se transcribe tal cual del PDF.",
               "        Verificar la leyenda de codigos antes de interpretarlo.",
               "", "  Archivos generados:",
               f"    {XLSX_FILE.name}", f"    {CSV_FILE.name}",
               f"    {REPORTE_FILE.name}", ""]

    reporte = "\n".join(lineas)
    print("\n" + reporte)
    path.write_text(reporte + "\n", encoding="utf-8")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 64)
    print("  [PUBLIT] - extraer_publit_v1")
    print("=" * 64)

    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    print(f"  data   : {CARPETA_DATOS}")
    print(f"  output : {CARPETA_SALIDA}\n")
    print(f"  PDFs   : {CARPETA_PDFS}")
    if not CARPETA_PDFS.exists():
        print(f"ERROR: no existe la carpeta {CARPETA_PDFS}")
        return

    pdfs = sorted(CARPETA_PDFS.glob("*.pdf"))
    if not pdfs:
        print(f"ERROR: no hay archivos .pdf en {CARPETA_PDFS}")
        return

    todas, contador, paginas, diag = [], [0], 0, []
    for path in pdfs:
        filas, n_pag = procesar_pdf(path, contador, diag)
        pres, _ = metadatos_archivo(path.name)
        om = PAGINAS_OMITIR.get(pres, 0)
        paginas += max(n_pag - om, 0)
        todas.extend(filas)
        print(f"  {path.name[:46]:<48} {len(filas):>5} refs / "
              f"{n_pag} pag (omitidas {om})")

    if not todas:
        print("ERROR: no se extrajo ninguna referencia.")
        return

    marcar_solapamiento(todas)
    escribir_csv(CSV_FILE, todas, CAMPOS)
    escribir_excel(todas, XLSX_FILE, diag)
    escribir_informe(todas, len(pdfs), paginas, REPORTE_FILE, diag)
    print(f"Archivos en {CARPETA_SALIDA}")


if __name__ == "__main__":
    main()
