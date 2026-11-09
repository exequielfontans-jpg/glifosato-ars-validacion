"""
cruzar_autores_publit_v5.py
===========================
Normaliza los autores de las listas de publicaciones del expediente europeo
(presentaciones 2012 y 2020) y los cruza contra la clasificacion Zona 1 /
Zona 2 del articulo de 2018.

Estructura de carpetas del proyecto:
  proyecto/
    script/     los .py
    data/       fuentes originales, NUNCA se escribe aqui
    output/     todo lo que generan los scripts, una subcarpeta por script

Si se borra output/ y se vuelven a correr los scripts en orden, el resultado
debe ser identico. Por eso las salidas no van en data/.

Proyecto: Corpus curados sobre glifosato — comparacion antologia Rossi /
          Glyphosate Task Force / expediente regulatorio. FIC-UDELAR.
Contexto: Cruce a nivel de autor · cobertura total del corpus

POR QUE A NIVEL DE AUTOR: el cruce por articulo solo alcanza a lo que esta
en la base de WoS, que cierra en 2016. A nivel de autor se cubre todo el
corpus, incluida la presentacion de 2020. La contrapartida es que el
resultado es una tasa DE AUTORES y no de articulos: no es comparable con la
linea de base por articulo. Reportarlas por separado.


Entradas:
  - data/Glif_-_Datos_ARS.xlsx       hojas 'Base' y 'AU Gephi'
  - output/publit/publit_v3.csv      - salida de extraer_publit_v3.py

Salidas:
  - cruce_autores_publit_v5.xlsx
       hoja 'Resumen'   - cifras por presentacion y seccion
       hoja 'Autores'   - un autor normalizado por fila, con zona y frecuencia
       hoja 'Ambiguos'  - claves con varias formas crudas o apellido frecuente
       hoja 'Instancias'- una fila por aparicion de autor en una referencia
  - cruce_autores_publit_v5.csv
  - reporte_cruzar_autores_publit_v5.txt
  - cruzar_autores_publit_v5_errores.log

Sensibilidad de salidas:
  NO SENSIBLE : todas (metadatos bibliograficos publicos)

Uso: python cruzar_autores_publit_v5.py

Requisitos: pip install openpyxl

Historial:
  v1 - version inicial: normalizacion de nombres tolerante a iniciales
       partidas por coma, nombres de pila completos e iniciales con guion;
       cruce contra AU Gephi y medicion de ambiguedad.
  v2 - agrega el conteo A NIVEL DE ARTICULO en cuatro grupos: Zona 1,
       Zona 2, grupo de control documental (vinculacion probada por avisos
       editoriales que el criterio estructural no detecta) y sin informacion
       (ningun autor clasificable). El techo de 2016 de la base no impide
       este conteo: lo que se necesita es que los AUTORES esten clasificados,
       no que el articulo este en el corpus.
  v3 - acota el analisis a una sola presentacion mediante la constante
       SOLO_PRESENTACION, fijada en 2012. La v2 procesaba las dos juntas, lo
       que mezclaba 1.078 referencias con autoria completa con 690 que solo
       consignan el primer autor: en esas ultimas cada referencia aportaba
       una sola instancia de autor y las tasas resultantes no eran
       comparables. Todas las cifras del informe pasan a referirse a 2012.
  v4 - separa la pertenencia al conjunto documental de la clasificacion por
       zona, que la v3 trataba como grupos excluyentes con CONTROL en cabeza
       de la jerarquia. Un autor puede estar documentado y ser Zona 1 a la
       vez, y de hecho tres lo son, de modo que la jerarquia marcaba como
       invisibles al criterio articulos que el criterio si detecta. El
       articulo se asigna ahora por zona y la condicion de invisibilidad se
       exige de forma explicita: tener al menos un autor del conjunto
       documental y ningun autor de Zona 1.
  v5 - completa CONTROL_DOCUMENTAL con las variantes del nombre que la base
       de 2018 registra como nodos separados de la misma persona. Faltaban
       'ACQUAVELLA, JF', que figura en el corpus con ocho referencias, y
       'SOLOMON, K'. La proporcion no detectada no cambia, porque esas
       referencias ya se contaban como Zona 1; lo que se corrige es el
       recuento de articulos con autor del conjunto, que pasa de 13 a 21.
"""

import csv
import re
import unicodedata
from pathlib import Path
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

CARPETA_SCRIPT = Path(__file__).resolve().parent
CARPETA_BASE   = CARPETA_SCRIPT.parent
CARPETA_DATOS  = CARPETA_BASE / "data"           # fuentes originales, solo lectura
CARPETA_OUT    = CARPETA_BASE / "output"         # todo lo generado por scripts
CARPETA_SALIDA = CARPETA_OUT / "cruce_autores"

EXCEL_FILE  = CARPETA_DATOS / "Glif_-_Datos_ARS.xlsx"
PUBLIT_FILE = CARPETA_OUT / "publit" / "publit_v3.csv"

XLSX_FILE    = CARPETA_SALIDA / "cruce_autores_publit_v5.xlsx"
CSV_FILE     = CARPETA_SALIDA / "cruce_autores_publit_v5.csv"
REPORTE_FILE = CARPETA_SALIDA / "reporte_cruzar_autores_publit_v5.txt"
LOG_FILE     = CARPETA_SALIDA / "cruzar_autores_publit_v5_errores.log"

HOJA_BASE  = "Base"
HOJA_ZONAS = "AU Gephi"
COL_Z_LABEL, COL_Z_ZONA, COL_Z_ARTS = 2, 4, 3
COL_B_AU = 2

# Particulas que forman parte del apellido y no deben tomarse como nombre
PARTICULAS = {"de", "del", "della", "da", "das", "dos", "van", "von", "der",
              "den", "la", "le", "el", "al", "bin", "ibn", "st", "san", "do"}

# Un apellido presente en mas de este numero de autores distintos de la base
# se considera frecuente y sus apareos se marcan para revision.
APELLIDO_FRECUENTE = 8

# Presentacion del expediente a analizar. La lista de 2020 consigna solo el
# primer autor seguido de "et al.", de modo que la autoria no puede
# recuperarse de la propia lista y sus cifras no son comparables con las de
# 2012. None procesa ambas presentaciones.
SOLO_PRESENTACION = "2012"

# GRUPO DE CONTROL: autores con vinculacion documentada por avisos editoriales
# de las propias revistas, que el criterio estructural de 2018 no detecta.
# Fuentes: Expression of Concern, Crit Rev Toxicol 48(10):891 (2018), sobre el
# suplemento 46(S1) de 2016; y la retractacion de Williams, Kroes y Munro
# (2000) por Regul Toxicol Pharmacol (2025).
# EDITAR ESTA LISTA es la unica forma de cambiar el grupo de control.
# Se incluyen las variantes del nombre que la base registra como nodos
# distintos de la misma persona: 'ACQUAVELLA, J' y 'ACQUAVELLA, JF' son el
# mismo autor, y lo mismo vale para 'SOLOMON, KR' y 'SOLOMON, K'. Omitir una
# variante deja sin marca de documentado a las referencias que la usan.
CONTROL_DOCUMENTAL = [
    "WILLIAMS, GM", "AARDEMA, M", "ACQUAVELLA, J", "ACQUAVELLA, JF",
    "BERRY, C", "BRUSICK, D", "BURNS, MM", "DE CAMARGO, JLV", "GARABRANT, D",
    "GREIM, H", "KIER, LD", "KIRKLAND, DJ", "MARSH, G", "SOLOMON, KR",
    "SOLOMON, K", "SORAHAN, T", "ROBERTS, A", "WEED, DL", "KROES, R",
    "MUNRO, IC",
]

# Prioridad de asignacion del articulo cuando conviven autores de distintas
# zonas. La pertenencia al conjunto documental NO entra aqui: es un atributo
# independiente de la zona, porque un autor documentado puede estar en la
# Zona 1. Un articulo se cuenta como invisible al criterio estructural cuando
# tiene al menos un autor documentado y ningun autor de Zona 1.
PRIORIDAD_GRUPOS = ["ZONA_1", "ZONA_2", "SIN_INFORMACION"]

CAMPOS_AUTORES = ["autor_norm", "apellido", "iniciales", "zona", "grupo",
                  "documentado", "n_instancias", "n_referencias",
                  "presentaciones", "secciones", "n_variantes", "variantes",
                  "arts_en_wos", "apellido_frecuente", "revisar"]
CAMPOS_INSTANCIAS = ["id_ref", "presentacion", "seccion", "anio",
                     "evaluacion", "autor_crudo", "autor_norm", "zona",
                     "grupo", "documentado"]
CAMPOS_ARTICULOS = ["id_ref", "presentacion", "seccion", "anio", "categoria",
                    "evaluacion", "titulo", "grupo", "solo_documental",
                    "n_autores", "n_zona1", "n_zona2", "n_documentados",
                    "n_sin_clasificar", "cobertura", "autores_documentados",
                    "autores_zona1", "apareo_dudoso"]


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def log_error(mensaje):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{ts}] {mensaje}\n")


def leer_csv_auto(path):
    if not path.exists():
        print(f"  No se encontro {path}")
        return [], []
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        for sep in (",", ";", "\t"):
            try:
                with open(path, encoding=enc, newline="") as f:
                    r = csv.DictReader(f, delimiter=sep)
                    filas = list(r)
                    cols = list(r.fieldnames or [])
                if cols and len(cols) > 1:
                    return filas, cols
            except Exception:
                continue
    return [], []


def escribir_csv(path, filas, campos):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=campos, extrasaction="ignore")
        w.writeheader()
        w.writerows(filas)


def sin_tildes(txt):
    t = unicodedata.normalize("NFKD", str(txt))
    return "".join(c for c in t if not unicodedata.combining(c))


# ─── NORMALIZACIÓN DE NOMBRES ─────────────────────────────────────────────────

def es_fragmento_iniciales(tramo):
    """True si el tramo son solo iniciales sueltas ('R.', 'K R', 'L-H').

    Hace falta porque el PDF corta con coma dentro del mismo autor:
    'Solomon K, R., Anadon A' son dos autores, no tres.
    """
    t = sin_tildes(tramo).replace(".", "").replace("-", " ").strip()
    if not t or len(t) > 5:
        return False
    return all(len(p) == 1 and p.isalpha() for p in t.split())


def separar_autores(cadena):
    """Divide la cadena de autores en tramos, reuniendo los que son solo
    iniciales con el autor anterior."""
    if not cadena:
        return []
    c = re.sub(r"\s+", " ", sin_tildes(cadena)).strip()
    c = re.sub(r"\s+(and|y|&)\s+", ", ", c, flags=re.I)
    c = c.replace(";", ",")
    tramos, actual = [], ""
    for parte in [p.strip() for p in c.split(",")]:
        if not parte:
            continue
        if actual and es_fragmento_iniciales(parte):
            actual += " " + parte          # cola del autor anterior
        else:
            if actual:
                tramos.append(actual)
            actual = parte
    if actual:
        tramos.append(actual)
    return tramos


def normalizar_autor(tramo):
    """Lleva un tramo a la forma 'APELLIDO, INICIALES' de la hoja AU Gephi.

    Tolera: iniciales pegadas ('Bernal MH'), nombre de pila completo
    ('Brain Richard A'), iniciales con guion ('Sanin L-H') y particulas
    ('de Villalobos C').
    """
    t = re.sub(r"[^A-Za-z\- ]", " ", sin_tildes(tramo))
    piezas = [p for p in re.sub(r"\s+", " ", t).strip().split(" ") if p]
    if not piezas:
        return "", "", ""

    # el apellido puede arrancar con particulas en minuscula
    i = 0
    apellido = []
    while i < len(piezas) and piezas[i].lower() in PARTICULAS:
        apellido.append(piezas[i])
        i += 1
    if i >= len(piezas):
        return "", "", ""
    apellido.append(piezas[i])
    i += 1

    iniciales = []
    for p in piezas[i:]:
        for sub in p.split("-"):
            if not sub:
                continue
            # 'KR' o 'MH' son dos iniciales pegadas; 'Richard' es un nombre
            # de pila del que solo interesa la primera letra. El criterio es
            # que un token corto y en mayusculas son iniciales.
            if sub.isupper() and len(sub) <= 3:
                iniciales.extend(list(sub))
            else:
                iniciales.append(sub[0].upper())

    ape = " ".join(apellido).upper()
    ini = "".join(iniciales)
    if not ape or len(ape) < 2:
        return "", "", ""
    return (f"{ape}, {ini}" if ini else ape), ape, ini


# ─── CARGA ────────────────────────────────────────────────────────────────────

def cargar_zonas(wb):
    """{autor_normalizado: (zona, articulos)} y frecuencia por apellido."""
    print("\n[1/4] Cargando clasificacion de zonas...")
    zonas, por_apellido = {}, {}
    filas = wb[HOJA_ZONAS].iter_rows(values_only=True)
    next(filas, None)
    for r in filas:
        if r[COL_Z_LABEL] is None:
            continue
        clave = re.sub(r"\s+", " ", sin_tildes(r[COL_Z_LABEL])).strip().upper()
        z = str(r[COL_Z_ZONA]).strip()
        if z not in ("1", "2"):
            continue
        zonas[clave] = (z, r[COL_Z_ARTS] or 0)
        ape = clave.split(",")[0].strip()
        por_apellido[ape] = por_apellido.get(ape, 0) + 1
    n1 = sum(1 for v in zonas.values() if v[0] == "1")
    print(f"  {len(zonas):,} autores | Zona 1: {n1:,} | Zona 2: {len(zonas)-n1:,}")
    return zonas, por_apellido


# ─── PROCESAMIENTO ────────────────────────────────────────────────────────────

def grupo_de(zona):
    """Grupo del autor segun su zona. La pertenencia al conjunto documental
    se registra aparte, en el campo 'documentado'."""
    if zona == "1":
        return "ZONA_1"
    if zona == "2":
        return "ZONA_2"
    return "SIN_INFORMACION"


def procesar(refs, zonas, por_apellido):
    print("\n[2/4] Normalizando y cruzando autores...")
    control = {normalizar_autor(x)[0] for x in CONTROL_DOCUMENTAL}
    control.discard("")
    instancias, agregado, articulos = [], {}, []

    for r in refs:
        crudos = separar_autores(r.get("autores", ""))
        conteo = {g: 0 for g in PRIORIDAD_GRUPOS}
        n_doc = 0
        nombres_doc, nombres_z1, dudoso = [], [], False

        for crudo in crudos:
            norm, ape, ini = normalizar_autor(crudo)
            if not norm:
                continue
            zona = zonas.get(norm, ("", 0))[0]
            grupo = grupo_de(zona)
            documentado = norm in control
            conteo[grupo] += 1
            if documentado:
                n_doc += 1
                nombres_doc.append(norm)
            if grupo == "ZONA_1":
                nombres_z1.append(norm)
            if (documentado or grupo == "ZONA_1") and \
                    por_apellido.get(ape, 0) > APELLIDO_FRECUENTE:
                dudoso = True

            instancias.append({
                "id_ref": r.get("id_ref", ""),
                "presentacion": r.get("presentacion", ""),
                "seccion": r.get("seccion", ""),
                "anio": r.get("anio", ""),
                "evaluacion": r.get("evaluacion", ""),
                "autor_crudo": crudo,
                "autor_norm": norm,
                "zona": zona or "sin clasificar",
                "grupo": grupo,
                "documentado": "SI" if documentado else "",
            })
            a = agregado.setdefault(norm, {
                "autor_norm": norm, "apellido": ape, "iniciales": ini,
                "zona": zona or "sin clasificar", "grupo": grupo,
                "documentado": "SI" if documentado else "",
                "arts_en_wos": zonas.get(norm, ("", 0))[1],
                "instancias": 0, "refs": set(), "pres": set(),
                "secs": set(), "variantes": set(),
            })
            a["instancias"] += 1
            a["refs"].add(r.get("id_ref", ""))
            a["pres"].add(r.get("presentacion", ""))
            a["secs"].add(r.get("seccion", ""))
            a["variantes"].add(crudo)

        n_aut = sum(conteo.values())
        if not n_aut:
            continue
        # el articulo toma la zona mas alta presente entre sus autores
        grupo_art = next((g for g in PRIORIDAD_GRUPOS if conteo[g]),
                         "SIN_INFORMACION")
        clasificables = n_aut - conteo["SIN_INFORMACION"]
        articulos.append({
            "id_ref": r.get("id_ref", ""),
            "presentacion": r.get("presentacion", ""),
            "seccion": r.get("seccion", ""),
            "anio": r.get("anio", ""),
            "categoria": r.get("categoria", ""),
            "evaluacion": r.get("evaluacion", ""),
            "titulo": (r.get("titulo", "") or "")[:180],
            "grupo": grupo_art,
            "solo_documental": "SI" if (n_doc and not conteo["ZONA_1"]) else "",
            "n_autores": n_aut,
            "n_zona1": conteo["ZONA_1"],
            "n_zona2": conteo["ZONA_2"],
            "n_documentados": n_doc,
            "n_sin_clasificar": conteo["SIN_INFORMACION"],
            "cobertura": round(clasificables / n_aut, 2),
            "autores_documentados": "; ".join(sorted(set(nombres_doc))),
            "autores_zona1": "; ".join(sorted(set(nombres_z1))[:6]),
            "apareo_dudoso": "SI" if dudoso else "",
        })

    autores = []
    for a in agregado.values():
        frec = por_apellido.get(a["apellido"], 0)
        motivos = []
        if len(a["variantes"]) > 1:
            motivos.append("VARIAS_FORMAS")
        if frec > APELLIDO_FRECUENTE and (a["documentado"] or
                                          a["grupo"] != "SIN_INFORMACION"):
            motivos.append("APELLIDO_FRECUENTE")
        if not a["iniciales"]:
            motivos.append("SIN_INICIALES")
        autores.append({
            "autor_norm": a["autor_norm"], "apellido": a["apellido"],
            "iniciales": a["iniciales"], "zona": a["zona"],
            "grupo": a["grupo"], "documentado": a["documentado"],
            "n_instancias": a["instancias"], "n_referencias": len(a["refs"]),
            "presentaciones": "; ".join(sorted(x for x in a["pres"] if x)),
            "secciones": "; ".join(sorted(x for x in a["secs"] if x)),
            "n_variantes": len(a["variantes"]),
            "variantes": " | ".join(sorted(a["variantes"])[:6]),
            "arts_en_wos": a["arts_en_wos"],
            "apellido_frecuente": frec,
            "revisar": "+".join(motivos),
        })
    autores.sort(key=lambda x: -x["n_instancias"])
    print(f"  {len(instancias):,} instancias | {len(autores):,} autores | "
          f"{len(articulos):,} articulos")
    return instancias, autores, articulos


# ─── SALIDA EN EXCEL ──────────────────────────────────────────────────────────

GRIS     = PatternFill("solid", fgColor="D9D9D9")
AMARILLO = PatternFill("solid", fgColor="FFF2CC")
ANCHOS = {"autor_norm": 28, "apellido": 20, "iniciales": 10, "zona": 14,
          "n_instancias": 12, "n_referencias": 13, "presentaciones": 16,
          "secciones": 34, "n_variantes": 12, "variantes": 56,
          "arts_en_wos": 12, "apellido_frecuente": 17, "revisar": 24,
          "grupo": 17, "n_zona1": 9, "n_zona2": 9, "n_documentados": 14,
          "documentado": 13, "solo_documental": 15,
          "n_sin_clasificar": 15, "cobertura": 11, "autores_documentados": 34,
          "autores_zona1": 34, "apareo_dudoso": 13, "n_autores": 11,
          "categoria": 11, "titulo": 60,
          "id_ref": 10, "presentacion": 12, "anio": 8, "evaluacion": 11,
          "autor_crudo": 30}


def _hoja(wb, nombre, filas, campos, resaltar=None):
    ws = wb.create_sheet(nombre)
    ws.append(campos)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True)
        c.fill = GRIS
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for f in filas:
        ws.append([f.get(k, "") for k in campos])
    for i, campo in enumerate(campos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ANCHOS.get(campo, 14)
    for fila in ws.iter_rows(min_row=2):
        for c in fila:
            c.font = Font(name="Arial")
            c.alignment = Alignment(vertical="top")
        if resaltar and resaltar(filas[fila[0].row - 2]):
            for c in fila:
                c.fill = AMARILLO
    if filas:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(campos))}{len(filas)+1}"
    ws.freeze_panes = "B2"
    return ws


def escribir_excel(instancias, autores, articulos, path):
    print("\n[3/4] Escribiendo el libro de Excel...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumen"

    n_i = len(instancias)
    def cuenta(pred, base):
        return sum(1 for x in base if pred(x))

    ws.append(["CRUCE DE AUTORES DEL EXPEDIENTE x ZONAS 2018"])
    ws.append([])
    ws.append(["ADVERTENCIA: esto es una tasa DE AUTORES, no de articulos."])
    ws.append(["No es comparable con la linea de base por articulo."])
    ws.append([])
    ws.append(["Instancias de autor", n_i])
    ws.append(["Autores unicos", len(autores)])
    ws.append([])
    ws.append(["Instancias", "n", "%"])
    for z in ("1", "2", "sin clasificar"):
        k = cuenta(lambda x: x["zona"] == z, instancias)
        ws.append([f"Zona {z}" if z != "sin clasificar" else "Sin clasificar",
                   k, f"{100*k/n_i:.1f}%" if n_i else "-"])
    ws.append([])
    ws.append(["Autores unicos", "n", "%"])
    na = len(autores)
    for z in ("1", "2", "sin clasificar"):
        k = cuenta(lambda x: x["zona"] == z, autores)
        ws.append([f"Zona {z}" if z != "sin clasificar" else "Sin clasificar",
                   k, f"{100*k/na:.1f}%" if na else "-"])
    ws.append([])
    ws.append(["Por presentacion (instancias)", "total", "Zona 1", "%"])
    for p in sorted({i["presentacion"] for i in instancias if i["presentacion"]}):
        sub = [i for i in instancias if i["presentacion"] == p]
        z1 = cuenta(lambda x: x["zona"] == "1", sub)
        ws.append([p, len(sub), z1, f"{100*z1/len(sub):.1f}%" if sub else "-"])
    ws.append([])
    ws.append(["ARTICULOS POR GRUPO", "n", "%"])
    na_art = len(articulos)
    for g in PRIORIDAD_GRUPOS:
        k = cuenta(lambda x: x["grupo"] == g, articulos)
        ws.append([g, k, f"{100*k/na_art:.1f}%" if na_art else "-"])
    z1_art = cuenta(lambda x: x["grupo"] == "ZONA_1", articulos)
    doc_art = cuenta(lambda x: x["n_documentados"], articulos)
    solo = cuenta(lambda x: x["solo_documental"], articulos)
    vinc = z1_art + solo
    ws.append([])
    ws.append(["Articulos con autor del conjunto documental", doc_art])
    ws.append(["  con al menos un autor de Zona 1", doc_art - solo])
    ws.append(["  sin ningun autor de Zona 1", solo])
    ws.append([])
    ws.append(["Articulos con vinculo acreditado", vinc])
    ws.append(["  detectados por el criterio estructural", z1_art])
    ws.append(["  invisibles al criterio estructural", solo])
    ws.append(["  proporcion no detectada",
               f"{100*solo/vinc:.1f}%" if vinc else "-"])
    ws.append([])
    ws.append(["Articulos con apareo dudoso",
               cuenta(lambda x: x["apareo_dudoso"], articulos)])
    ws.append(["Autores marcados para revision",
               cuenta(lambda x: x["revisar"], autores)])

    ws.column_dimensions["A"].width = 40
    for i in range(2, 5):
        ws.column_dimensions[get_column_letter(i)].width = 12
    for fila in ws.iter_rows():
        for c in fila:
            c.font = Font(name="Arial")
    for r in (1, 3, 9, 14, 19):
        ws.cell(row=r, column=1).font = Font(name="Arial", bold=True)

    _hoja(wb, "Articulos", articulos, CAMPOS_ARTICULOS,
          resaltar=lambda f: bool(f["solo_documental"]))
    _hoja(wb, "Autores", autores, CAMPOS_AUTORES,
          resaltar=lambda f: bool(f["revisar"]))
    _hoja(wb, "Ambiguos", [a for a in autores if a["revisar"]], CAMPOS_AUTORES)
    _hoja(wb, "Instancias", instancias, CAMPOS_INSTANCIAS)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# ─── INFORME ──────────────────────────────────────────────────────────────────

def escribir_informe(instancias, autores, articulos, path):  # noqa: C901
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    n_i, n_a = len(instancias), len(autores)
    pct = lambda k, d: f"{(100*k/d):.1f}%" if d else "-"

    lineas = [
        "=" * 68,
        "  [CRUCE AUTORES PUBLIT] - cruzar_autores_publit_v5",
        f"  Fecha: {ahora}",
        "=" * 68,
        "",
        "  ADVERTENCIA: tasa DE AUTORES, no de articulos. No comparable con",
        "  la linea de base por articulo del cruce anterior.",
        "",
        f"  Instancias de autor : {n_i:,}",
        f"  Autores unicos      : {n_a:,}",
        "",
        "  Instancias por zona:",
    ]
    for z, et in (("1", "Zona 1"), ("2", "Zona 2"),
                  ("sin clasificar", "Sin clasificar")):
        k = sum(1 for i in instancias if i["zona"] == z)
        lineas.append(f"    {et:<16} {k:>7,}  ({pct(k, n_i)})")

    lineas += ["", "  Autores unicos por zona:"]
    for z, et in (("1", "Zona 1"), ("2", "Zona 2"),
                  ("sin clasificar", "Sin clasificar")):
        k = sum(1 for a in autores if a["zona"] == z)
        lineas.append(f"    {et:<16} {k:>7,}  ({pct(k, n_a)})")

    for campo, titulo in (("presentacion", "Por presentacion"),
                          ("seccion", "Por seccion"),
                          ("evaluacion", "Por marca de evaluacion")):
        grupos = {}
        for i in instancias:
            k = i.get(campo) or "(sin dato)"
            g = grupos.setdefault(k, [0, 0])
            g[0] += 1
            g[1] += 1 if i["zona"] == "1" else 0
        if len(grupos) <= 1:
            continue
        lineas += ["", f"  {titulo} (instancias):",
                   f"    {'valor':<28} {'total':>8} {'Zona 1':>8} {'%':>7}"]
        for k, (t, z) in sorted(grupos.items(), key=lambda x: -x[1][0])[:12]:
            lineas.append(f"    {str(k)[:26]:<28} {t:>8,} {z:>8,} {pct(z, t):>7}")

    n_art = len(articulos)
    lineas += ["", "  ARTICULOS POR GRUPO (por zona del autor)",
               f"    {'grupo':<20} {'n':>8} {'%':>8}"]
    for g in PRIORIDAD_GRUPOS:
        k = sum(1 for a in articulos if a["grupo"] == g)
        lineas.append(f"    {g:<20} {k:>8,} {pct(k, n_art):>8}")

    z1_art = sum(1 for a in articulos if a["grupo"] == "ZONA_1")
    doc_art = sum(1 for a in articulos if a["n_documentados"])
    solo = sum(1 for a in articulos if a["solo_documental"])
    vinc = z1_art + solo
    lineas += ["",
               "  CONJUNTO DOCUMENTAL EN EL CORPUS",
               f"    Articulos con autor del conjunto  : {doc_art:,}",
               f"      con al menos un autor de Zona 1 : {doc_art - solo:,}"
               f"  ({pct(doc_art - solo, doc_art)})",
               f"      sin ningun autor de Zona 1      : {solo:,}"
               f"  ({pct(solo, doc_art)})",
               "",
               f"  Articulos con vinculo acreditado  : {vinc:,} ({pct(vinc, n_art)})",
               f"    detectados por el criterio      : {z1_art:,}",
               f"    invisibles al criterio          : {solo:,}",
               f"    proporcion que el criterio NO ve: "
               f"{pct(solo, vinc) if vinc else '-'}",
               "",
               "  Autores del conjunto documental presentes en el corpus:",
               f"    {'autor':<26} {'zona':>6} {'refs':>6} {'arts WoS':>9}"]
    for a in sorted((x for x in autores if x["documentado"]),
                    key=lambda x: (x["zona"], x["autor_norm"])):
        lineas.append(f"    {a['autor_norm'][:24]:<26} {a['zona']:>6} "
                      f"{a['n_referencias']:>6} {a['arts_en_wos']:>9}")
    lineas += ["", "  Articulos invisibles al criterio, uno por uno:"]
    for a in sorted((x for x in articulos if x["solo_documental"]),
                    key=lambda x: str(x["anio"])):
        lineas.append(f"    {str(a['anio']):>6}  {a['autores_documentados'][:40]:<42}"
                      f" {a['seccion'][:22]}")
        lineas.append(f"            {str(a['titulo'])[:86]}")

    for campo, titulo in (("presentacion", "Articulos por presentacion"),
                          ("evaluacion", "Articulos por marca de evaluacion"),
                          ("categoria", "Articulos por categoria del PDF")):
        grupos = {}
        for a in articulos:
            k = a.get(campo) or "(sin dato)"
            g = grupos.setdefault(k, [0, 0, 0])
            g[0] += 1
            g[1] += 1 if a["grupo"] == "ZONA_1" else 0
            g[2] += 1 if a["solo_documental"] else 0
        if len(grupos) <= 1:
            continue
        lineas += ["", f"  {titulo}:",
                   f"    {'valor':<24} {'total':>8} {'Zona 1':>8} {'invis.':>8}"]
        for k, (t, z, c) in sorted(grupos.items(), key=lambda x: -x[1][0])[:10]:
            lineas.append(f"    {str(k)[:22]:<24} {t:>8,} {z:>8,} {c:>8,}")

    dud = sum(1 for a in articulos if a["apareo_dudoso"])
    lineas += ["", f"  Articulos con apareo dudoso : {dud:,} ({pct(dud, n_art)})",
               "    (clasificados por un apellido frecuente; revisar a mano)"]

    rev = sum(1 for a in autores if a["revisar"])
    lineas += ["", f"  Autores marcados para revision : {rev:,} ({pct(rev, n_a)})"]
    motivos = {}
    for a in autores:
        if a["revisar"]:
            motivos[a["revisar"]] = motivos.get(a["revisar"], 0) + 1
    for m, k in sorted(motivos.items(), key=lambda x: -x[1]):
        lineas.append(f"    {m:<28} {k:>6}")

    lineas += ["", "  Autores mas frecuentes en el corpus:",
               f"    {'autor':<26} {'refs':>6} {'zona':>16} {'arts WoS':>9}"]
    for a in autores[:15]:
        lineas.append(f"    {a['autor_norm'][:24]:<26} {a['n_referencias']:>6} "
                      f"{a['zona']:>16} {a['arts_en_wos']:>9}")

    lineas += ["", "  Archivos generados:",
               f"    {XLSX_FILE.name}   <- abrir este para revisar",
               f"    {CSV_FILE.name}", f"    {REPORTE_FILE.name}", ""]

    reporte = "\n".join(lineas)
    print("\n" + reporte)
    path.write_text(reporte + "\n", encoding="utf-8")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 68)
    print("  [CRUCE AUTORES PUBLIT] - cruzar_autores_publit_v5")
    print("=" * 68)

    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    print(f"  data   : {CARPETA_DATOS}")
    print(f"  output : {CARPETA_SALIDA}\n")
    if not EXCEL_FILE.exists():
        print(f"ERROR: no se encontro {EXCEL_FILE}")
        return
    refs, _ = leer_csv_auto(PUBLIT_FILE)
    if not refs:
        print(f"ERROR: no se pudo leer {PUBLIT_FILE}")
        return

    if SOLO_PRESENTACION:
        refs = [r for r in refs
                if str(r.get("presentacion", "")).strip() == SOLO_PRESENTACION]
        print(f"  Filtrado a la presentacion {SOLO_PRESENTACION}: "
              f"{len(refs):,} referencias\n")
        if not refs:
            print("ERROR: ninguna referencia en esa presentacion.")
            return

    wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    zonas, por_apellido = cargar_zonas(wb)

    instancias, autores, articulos = procesar(refs, zonas, por_apellido)
    if not instancias:
        print("ERROR: no se extrajo ningun autor.")
        return

    escribir_csv(CSV_FILE, autores, CAMPOS_AUTORES)
    escribir_excel(instancias, autores, articulos, XLSX_FILE)
    print("\n[4/4] Escribiendo el informe...")
    escribir_informe(instancias, autores, articulos, REPORTE_FILE)
    print(f"Archivos en {CARPETA_SALIDA}")


if __name__ == "__main__":
    main()
