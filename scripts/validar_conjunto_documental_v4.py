"""
validar_conjunto_documental_v4.py
=================================
Construye el conjunto de validacion documental, lo coteja contra la red de
coautoria de 2018 y mide la sensibilidad del criterio estructural a nivel de
autor y a nivel de articulo.

Proyecto: Validacion de un procedimiento de atribucion estructural de
          vinculacion corporativa — expediente europeo de renovacion del
          glifosato (2012). FIC-UDELAR.
Contexto: Resultados · validacion externa del criterio estructural
          (apartados 4.3, 4.4 y 4.5; Tablas 4, 6, 7 y 8)

El conjunto de validacion se define en la seccion CONFIGURACION a partir de
avisos editoriales publicados: la correccion del suplemento 46(S1) de
Critical Reviews in Toxicology (https://doi.org/10.1080/10408444.2018.1522175)
y la nota de retractacion de Williams, Kroes y Munro (2000)
(https://doi.org/10.1016/j.yrtph.2025.106006). Cada autor lleva transcripta su
vinculacion declarada y las formas de nombre con que podria figurar en la base.

El script no reutiliza la asignacion jerarquica de grupos de
cruzar_autores_publit_v3.py. Aplica en su lugar la condicion explicita que el
analisis declara: un articulo queda invisible al criterio estructural cuando
tiene al menos un autor del conjunto documental y ningun autor de Zona 1.

Entradas:
  - data/Glif_-_Datos_ARS.xlsx               — hojas 'Base' y 'AU Gephi'
  - output/cruce_publit/cruce_publit_v2.csv  — cruce expediente x WoS
    (si no existe, usa cruce_publit_v1.csv)
  - output/cruce_autores/cruce_autores_publit_v5.xlsx — hoja 'Instancias',
    salida de cruzar_autores_publit_v5.py. Entrada opcional: sin ella se
    omite la Tabla 6. (si no existe, usa la version v4)

Salidas:
  - output/validacion/tabla4_y_recalculo_v4.xlsx            — planilla auditable
  - output/validacion/reporte_validar_conjunto_documental.txt — informe estadistico

Sensibilidad de salidas:
  NO SENSIBLE : ambas. La vinculacion de cada autor proviene de avisos
                editoriales publicos y los recuentos son agregados.

Uso: python validar_conjunto_documental_v4.py
     (correr desde scripts/ o cualquier ubicacion)

Requisitos: pip install openpyxl

Historial:
  v1 — version inicial. Sustituye el grupo CONTROL de cruzar_autores_publit_v3
       por la condicion explicita de ausencia de Zona 1, con lo que los
       articulos invisibles al criterio pasan de 13 a 6 (5 en la ventana
       2001-2012). Agrega la deteccion a nivel de articulo y el analisis de
       sensibilidad del conjunto sin Kroes.
  v2 — la Tabla 7 compara ambos corpus hasta 2012, el horizonte propio del
       expediente, en lugar de la ventana 2001-2012 de las tasas de referencia.
       Con eso la tabla incorpora el articulo de Williams, Kroes y Munro (2000).
       Se agrega un bloque de sensibilidad con los tres encuadres temporales.
  v3 - genera la Tabla 6, el perfil de los autores del conjunto presentes en
       el corpus del expediente, que hasta ahora no tenia respaldo
       reproducible. Toma como entrada adicional la hoja 'Instancias' del
       cruce a nivel de autor; si ese archivo falta, omite la hoja y sigue.
       Las formas del nombre que la base registra como nodos distintos de la
       misma persona se reportan unidas, a diferencia del analisis de red,
       que las trata separadas.
  v4 - agrega la trayectoria de cada autor del conjunto en la base de 2018:
       articulos, anios, coautores y cuantos de esos coautores integran la
       Zona 1. Sostiene la lectura negativa del criterio de 2018: quedar en
       Zona 2 significa no haber firmado ningun articulo del corpus junto a
       alguien con afiliacion corporativa declarada en ese mismo articulo.
"""

import bisect
import collections
import csv
import statistics
import unicodedata
from pathlib import Path
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

CARPETA_SCRIPT = Path(__file__).resolve().parent
CARPETA_PROY   = CARPETA_SCRIPT.parent
CARPETA_DATA   = CARPETA_PROY / "data"
CARPETA_CRUCE  = CARPETA_PROY / "output" / "cruce_publit"
CARPETA_SALIDA = CARPETA_PROY / "output" / "validacion"

# Entradas
BASE_XLSX = CARPETA_DATA / "Glif_-_Datos_ARS.xlsx"
CRUCE_CSV = CARPETA_CRUCE / "cruce_publit_v2.csv"
CRUCE_ALT = CARPETA_CRUCE / "cruce_publit_v1.csv"

# Entrada opcional: la red de autores del expediente, para la Tabla 6
CARPETA_AUTORES = CARPETA_PROY / "output" / "cruce_autores"
AUTORES_XLSX = CARPETA_AUTORES / "cruce_autores_publit_v5.xlsx"
AUTORES_ALT  = CARPETA_AUTORES / "cruce_autores_publit_v4.xlsx"

# Salidas
OUTPUT_FILE  = CARPETA_SALIDA / "tabla4_y_recalculo_v4.xlsx"
REPORTE_FILE = CARPETA_SALIDA / "reporte_validar_conjunto_documental.txt"

# Hojas e indices de columna de Glif_-_Datos_ARS.xlsx
HOJA_BASE, HOJA_ZONA = "Base", "AU Gephi"
COL_ID, COL_AU, COL_TI, COL_TC, COL_PY, COL_DI = 0, 2, 9, 32, 45, 55
COL_Z_LABEL, COL_Z_ARTS, COL_Z_ZONA = 2, 3, 4

# Ventanas temporales. La estandarizacion del apartado 4.2 trabaja en la ventana
# de las tasas de referencia por area tematica. La Tabla 7 no usa tasas por area,
# de modo que compara ambos corpus hasta 2012: un expediente presentado ese anio
# no puede contener nada posterior, y recortar alli la literatura general evita
# inflar su denominador con publicaciones que el expediente no podia citar.
VENTANA_REFERENCIA = (2001, 2012)
VENTANA_TABLA7     = (0, 2012)

# Autor excluido en el analisis de sensibilidad, por tener acreditacion
# colectiva y no individual en la nota de retractacion.
AUTOR_SENSIBILIDAD = "Kroes, R."

# ─── Conjunto de validacion documental ───
# (autor, formas de nombre en la base, tipo de vinculo, relacion previa,
#  vinculacion transcripta del aviso editorial)
CONJUNTO_DOCUMENTAL = [
    ("Acquavella, J. F.", ["ACQUAVELLA, JF", "ACQUAVELLA, J"],
     "Empleo en la empresa", "Si",
     "Empleado de Monsanto entre 1989 y 2004. Retribuido por el panel mediante un "
     "contrato de consultoria vigente con la empresa. Consultor en un litigio ajeno "
     "al glifosato relativo a una antigua planta quimica de Monsanto."),
    ("Kier, L. D.", ["KIER, LD"],
     "Empleo en la empresa", "Si",
     "Ex empleado de Monsanto. Contratado por la empresa como consultor para asistir "
     "al panel en genotoxicidad y estres oxidativo; incorporado como coautor despues "
     "de redactado el informe final."),
    ("Greim, H.", ["GREIM, H", "GREIM, HA"],
     "Consultoria u honorarios", "Si",
     "Consultor independiente previo de Monsanto. Coautor de Greim et al. (2015) junto "
     "a un empleado de la empresa que aporto los datos originales de los estudios de "
     "Monsanto."),
    ("Williams, G. M.", ["WILLIAMS, GM"],
     "Consultoria u honorarios", "Si",
     "Consultor independiente previo de Monsanto. Con Munro y la asistencia de Bryant "
     "preparo la evaluacion de seguridad y riesgo de Roundup financiada por la empresa "
     "(Williams et al., 2000), retractada en 2025."),
    ("Munro, I. C.", ["MUNRO, IC"],
     "Consultoria u honorarios", "Si",
     "Ex empleado de Cantox. Con Williams preparo la evaluacion de seguridad y riesgo "
     "de Roundup financiada por Monsanto."),
    ("Solomon, K. R.", ["SOLOMON, KR", "SOLOMON, K"],
     "Consultoria u honorarios", "Si",
     "Consultor independiente de Monsanto en la desregulacion de alfalfa RR en los "
     "Estados Unidos (2012-2014). En colaboracion con Cantox contribuyo a la evaluacion "
     "ecotoxicologica de Roundup (Giesy et al., 2000). Integro un consejo asesor "
     "cientifico de Dow AgroSciences entre 2014 y 2016."),
    ("Kirkland, D. J.", ["KIRKLAND, DJ"],
     "Consultoria u honorarios", "Si",
     "Consultor independiente previo de Monsanto, con participacion en el Glyphosate "
     "Task Force europeo."),
    ("Sorahan, T.", ["SORAHAN, T"],
     "Consultoria u honorarios", "Si",
     "Consultoria para Monsanto en litigios sobre glifosato. Honorarios de consultoria "
     "y viaticos de Monsanto Europe SA/NV como integrante del European Glyphosate "
     "Toxicology Advisory Panel. Observador por Monsanto en la reunion de la monografia "
     "112 de IARC."),
    ("Kroes, R.", ["KROES, R"],
     "Autoria no declarada", "Si",
     "Coautor de la evaluacion de Roundup financiada por Monsanto. La nota de "
     "retractacion de 2026 establece participacion no acreditada de personal de la "
     "empresa en la redaccion y posible compensacion financiera no declarada a los "
     "autores, referida al conjunto de ellos."),
    ("Berry, C.", ["BERRY, C", "BERRY, CL"],
     "Consultoria u honorarios", "Si",
     "Consultor independiente previo de Monsanto."),
    ("Brusick, D.", ["BRUSICK, D"],
     "Consultoria u honorarios", "Si",
     "Consultor independiente previo de Monsanto."),
    ("de Camargo, J. L. V.", ["DE CAMARGO, JLV", "CAMARGO, JLV"],
     "Consultoria u honorarios", "Si",
     "Consultor independiente previo de Monsanto."),
    ("Garabrant, D.", ["GARABRANT, D", "GARABRANT, DH"],
     "Consultoria u honorarios", "Si",
     "Retenido conjuntamente por Monsanto y otras empresas en litigios sobre glifosato "
     "y leucemia. Integro un consejo asesor cientifico de Dow AgroSciences entre 2014 "
     "y 2016. Consultoria para un abogado de Pharmacia en litigio ajeno al glifosato."),
    ("Weed, D. L.", ["WEED, DL"],
     "Consultoria u honorarios", "Si",
     "Consultoria para Monsanto en litigios ajenos al glifosato."),
    ("Roberts, A.", ["ROBERTS, A"],
     "Empleo en el intermediario", "No",
     "Empleado de Intertek; coordino la preparacion del suplemento financiado por "
     "Monsanto. Declara no haber participado antes en actividades relacionadas con "
     "el glifosato."),
    ("Aardema, M.", ["AARDEMA, M", "AARDEMA, MJ"],
     "Retribucion por la pieza", "No",
     "Retribuida por el panel mediante el contrato de Intertek financiado por Monsanto. "
     "Declara no haber participado antes en actividades relacionadas con el glifosato."),
    ("Burns, M. M.", ["BURNS, MM"],
     "Retribucion por la pieza", "No",
     "Retribuida por el panel mediante el contrato de Intertek financiado por Monsanto. "
     "Declara no haber participado antes en actividades relacionadas con el glifosato."),
    ("Marsh, G.", ["MARSH, G", "MARSH, GM"],
     "Retribucion por la pieza", "No",
     "Retribuido por el panel mediante el contrato de Intertek financiado por Monsanto. "
     "Declara no haber participado antes en actividades relacionadas con el glifosato."),
]

# Formato
AZUL = PatternFill("solid", fgColor="2F5496")
BLANCO_NEGRITA = Font(bold=True, color="FFFFFF")
NEGRITA = Font(bold=True)
ROJO_NEGRITA = Font(bold=True, color="C00000")


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def normalizar(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return s.strip().upper()


def leer_csv_auto(path):
    if not path.exists():
        print(f"  No se encontro {path.name} - omitido.")
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


def anio(valor):
    try:
        return int(str(valor)[:4])
    except (TypeError, ValueError):
        return None


def encabezar(ws, columnas, anchos):
    ws.append(columnas)
    for celda in ws[1]:
        celda.font = BLANCO_NEGRITA
        celda.fill = AZUL
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.freeze_panes = "A2"


# ─── CARGA ────────────────────────────────────────────────────────────────────

def cargar_base():
    if not BASE_XLSX.exists():
        raise FileNotFoundError(f"No se encontro {BASE_XLSX}")
    wb = openpyxl.load_workbook(BASE_XLSX, read_only=True, data_only=True)

    zona, arts = {}, {}
    for i, fila in enumerate(wb[HOJA_ZONA].iter_rows(values_only=True)):
        if i == 0:
            continue
        etiqueta = normalizar(fila[COL_Z_LABEL])
        if etiqueta:
            zona[etiqueta] = str(fila[COL_Z_ZONA]).strip()
            arts[etiqueta] = fila[COL_Z_ARTS] or 0

    registros = {}
    for i, fila in enumerate(wb[HOJA_BASE].iter_rows(values_only=True)):
        if i == 0:
            continue
        registros[str(fila[COL_ID]).strip()] = {
            "autores": [normalizar(a) for a in str(fila[COL_AU] or "").split(";") if a.strip()],
            "anio": anio(fila[COL_PY]),
            "titulo": str(fila[COL_TI] or ""),
            "doi": str(fila[COL_DI] or ""),
            "citas": fila[COL_TC],
        }
    wb.close()
    return zona, arts, registros


def cargar_cruce():
    path = CRUCE_CSV if CRUCE_CSV.exists() else CRUCE_ALT
    filas, _ = leer_csv_auto(path)
    apareadas = [f for f in filas if str(f.get("estado_apareo", "")).strip() == "APAREADO"]
    secciones = collections.defaultdict(set)
    publicaciones = {}
    for fila in apareadas:
        clave = str(fila.get("id_wos", "")).strip()
        sec = str(fila.get("seccion", "")).strip()
        if sec:
            secciones[clave].add(sec)
        publicaciones.setdefault(clave, fila)
    return apareadas, publicaciones, secciones, path


def cargar_instancias_expediente():
    """Red de autores del expediente, desde la hoja 'Instancias' del cruce a
    nivel de autor. Devuelve ({}, {}) si el archivo no esta disponible."""
    path = AUTORES_XLSX if AUTORES_XLSX.exists() else AUTORES_ALT
    if not path.exists():
        print(f"  No se encontro {AUTORES_XLSX.name}: se omite la Tabla 6.")
        return {}, {}, None
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if "Instancias" not in wb.sheetnames:
        wb.close()
        print(f"  {path.name} no tiene hoja 'Instancias': se omite la Tabla 6.")
        return {}, {}, None
    filas = wb["Instancias"].iter_rows(values_only=True)
    cols = list(next(filas, ()) or ())
    try:
        i_ref, i_aut = cols.index("id_ref"), cols.index("autor_norm")
    except ValueError:
        wb.close()
        print(f"  {path.name}: faltan columnas id_ref o autor_norm.")
        return {}, {}, None
    refs_de = collections.defaultdict(set)
    autores_de = collections.defaultdict(set)
    for fila in filas:
        autor = normalizar(fila[i_aut])
        ref = str(fila[i_ref] or "").strip()
        if autor and ref:
            refs_de[autor].add(ref)
            autores_de[ref].add(autor)
    wb.close()
    return refs_de, autores_de, path


# ─── PROCESAMIENTO ────────────────────────────────────────────────────────────

def resolver_formas(zona):
    """Para cada autor del conjunto, las formas de nombre presentes en la base."""
    resuelto = []
    for autor, formas, tipo, previa, vinculo in CONJUNTO_DOCUMENTAL:
        presentes = [f for f in formas if f in zona]
        clasif = "-"
        if presentes:
            clasif = "Zona 1" if any(zona[f] == "1" for f in presentes) else "Zona 2"
        resuelto.append({
            "autor": autor, "formas": formas, "presentes": presentes,
            "tipo": tipo, "previa": previa, "vinculo": vinculo,
            "en_red": bool(presentes), "clasificacion": clasif,
        })
    return resuelto


def articulos_documentados(ids, registros, zona, formas_doc, ventana=None):
    """Articulos con al menos un autor del conjunto documental."""
    salida = []
    for clave in ids:
        reg = registros.get(clave)
        if not reg:
            continue
        if ventana and not (reg["anio"] and ventana[0] <= reg["anio"] <= ventana[1]):
            continue
        documentados = [a for a in reg["autores"] if a in formas_doc]
        if not documentados:
            continue
        de_zona1 = [a for a in reg["autores"] if zona.get(a) == "1"]
        salida.append({
            "id": clave, "anio": reg["anio"], "titulo": reg["titulo"], "doi": reg["doi"],
            "citas": reg["citas"], "documentados": documentados, "zona1": de_zona1,
            "detectado": bool(de_zona1),
        })
    return salida


def perfil_documentados(resuelto, registros, zona, refs_de, autores_de,
                        wos_por_ref):
    """Tabla 6: perfil de los autores del conjunto presentes en el corpus.

    La citacion se reporta como mediana de los articulos del autor, no como
    total acumulado, y se acompana del percentil en dos distribuciones
    construidas sobre los autores del corpus del expediente: la de la obra
    completa en la base y la restringida a los articulos del expediente.
    """
    if not refs_de:
        return [], {}

    arts_de = collections.defaultdict(list)
    for rid, reg in registros.items():
        for autor in reg["autores"]:
            arts_de[autor].append(rid)

    def trayectoria(formas):
        """Produccion y entorno de coautoria del autor en la base de 2018."""
        ids = {i for f in formas for i in arts_de.get(f, [])}
        anios = sorted(registros[i]["anio"] for i in ids if registros[i]["anio"])
        coaut = set()
        for i in ids:
            coaut |= set(registros[i]["autores"])
        coaut -= set(formas)
        z1 = [c for c in coaut if zona.get(c) == "1"]
        return {
            "arts_base": len(ids),
            "anio_min": anios[0] if anios else None,
            "anio_max": anios[-1] if anios else None,
            "coaut_base": len(coaut),
            "coaut_base_z1": len(z1),
            "pct_coaut_z1": (len(z1) / len(coaut)) if coaut else None,
            "coaut_z1_nombres": "; ".join(sorted(z1)),
        }

    def mediana_citas(ids):
        v = [registros[i]["citas"] for i in ids
             if i in registros and isinstance(registros[i]["citas"], (int, float))]
        return statistics.median(v) if v else None

    med_obra, med_exp = {}, {}
    for autor, refs in refs_de.items():
        m = mediana_citas(arts_de.get(autor, []))
        if m is not None:
            med_obra[autor] = m
        m = mediana_citas([wos_por_ref[r] for r in refs if r in wos_por_ref])
        if m is not None:
            med_exp[autor] = m
    dist_obra = sorted(med_obra.values())
    dist_exp = sorted(med_exp.values())

    def percentil(valor, dist):
        if valor is None or not dist:
            return None
        return 100 * bisect.bisect_left(dist, valor) / len(dist)

    filas = []
    for r in resuelto:
        presentes = [f for f in r["formas"] if f in refs_de]
        if not presentes:
            continue
        refs = set().union(*(refs_de[f] for f in presentes))
        coautores = set()
        for ref in refs:
            coautores |= autores_de[ref]
        coautores -= set(presentes)
        obra = mediana_citas([i for f in presentes for i in arts_de.get(f, [])])
        exped = mediana_citas([wos_por_ref[x] for x in refs if x in wos_por_ref])
        filas.append({
            "autor": r["autor"], "clasificacion": r["clasificacion"],
            "formas": " / ".join(presentes), "n_refs": len(refs),
            "n_coautores": len(coautores),
            "cit_obra": obra, "pct_obra": percentil(obra, dist_obra),
            "cit_exped": exped, "pct_exped": percentil(exped, dist_exp),
            **trayectoria(presentes),
        })
    filas.sort(key=lambda x: (x["clasificacion"] != "Zona 1", x["autor"]))
    referencia = {
        "n_obra": len(dist_obra),
        "med_obra": statistics.median(dist_obra) if dist_obra else None,
        "n_exped": len(dist_exp),
        "med_exped": statistics.median(dist_exp) if dist_exp else None,
        "n_autores": len(refs_de),
    }
    return filas, referencia


def contar_corpus(ids, registros, zona, formas_doc, ventana=None):
    """Recuentos de un corpus: publicaciones, detectadas, no detectadas."""
    n = z1 = doc_sin_z1 = doc_con_z1 = 0
    for clave in ids:
        reg = registros.get(clave)
        if not reg:
            continue
        if ventana and not (reg["anio"] and ventana[0] <= reg["anio"] <= ventana[1]):
            continue
        n += 1
        marcado = any(zona.get(a) == "1" for a in reg["autores"])
        documentado = any(a in formas_doc for a in reg["autores"])
        if marcado:
            z1 += 1
        if documentado and marcado:
            doc_con_z1 += 1
        if documentado and not marcado:
            doc_sin_z1 += 1
    total = z1 + doc_sin_z1
    return {"n": n, "z1": z1, "no_detectadas": doc_sin_z1, "ya_contadas": doc_con_z1,
            "total": total, "pct_no_detectado": doc_sin_z1 / total if total else 0.0}


# ─── PLANILLA ─────────────────────────────────────────────────────────────────

def escribir_planilla(resuelto, arts_zona, arts_exp, articulos, exp, gen,
                      encuadres, perfil, ref6, path):
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "1_Tabla4"
    encabezar(ws, ["Autor", "Tipo de vinculo", "En la red", "Clasificacion estructural",
                   "Relacion previa", "Forma del nombre en la base de 2018",
                   "Articulos en la base", "Articulos del expediente"],
              [22, 26, 10, 14, 12, 26, 11, 13])
    orden = sorted(resuelto, key=lambda r: (not r["en_red"], r["clasificacion"] != "Zona 1",
                                            r["autor"]))
    for r in orden:
        if r["en_red"]:
            ws.append([r["autor"], r["tipo"], "Si", r["clasificacion"], r["previa"],
                       " / ".join(r["presentes"]),
                       sum(arts_zona.get(f, 0) for f in r["presentes"]),
                       sum(arts_exp.get(f, 0) for f in r["presentes"])])
        else:
            ws.append([r["autor"], r["tipo"], "No", "-", r["previa"], "(ausente)", "-", "-"])
        for j in (3, 4, 5, 7, 8):
            ws.cell(ws.max_row, j).alignment = Alignment(horizontal="center")
    ws.append([])
    ws.append(["Nota", "El tipo de vinculo consigna la relacion mas extensa acreditada en el "
                       "aviso editorial. La Tabla 8 del anexo transcribe la vinculacion de cada "
                       "autor. 'En la red' indica presencia en la red de coautoria de 17.523 "
                       "autores de Fontans-Alvarez et al. (2018).", "", "", "", "", "", ""])
    ws.cell(ws.max_row, 1).font = NEGRITA
    ws.cell(ws.max_row, 2).alignment = Alignment(wrap_text=True)

    ws = wb.create_sheet("2_Deteccion")
    encabezar(ws, ["Nivel de analisis", "Casos", "Detectados", "No detectados", "Sensibilidad"],
              [56, 10, 12, 14, 12])
    en_red = [r for r in resuelto if r["en_red"]]
    det = [r for r in en_red if r["clasificacion"] == "Zona 1"]
    ws.append(["Autores del conjunto presentes en la red", len(en_red), len(det),
               len(en_red) - len(det), round(len(det) / len(en_red), 4) if en_red else 0])
    sin_k = [r for r in en_red if r["autor"] != AUTOR_SENSIBILIDAD]
    det_k = [r for r in sin_k if r["clasificacion"] == "Zona 1"]
    ws.append([f"Idem, excluyendo a {AUTOR_SENSIBILIDAD} (acreditacion colectiva)",
               len(sin_k), len(det_k), len(sin_k) - len(det_k),
               round(len(det_k) / len(sin_k), 4) if sin_k else 0])
    n_art = len(articulos)
    det_art = sum(1 for a in articulos if a["detectado"])
    ws.append(["Articulos del expediente con autor documentado", n_art, det_art,
               n_art - det_art, round(det_art / n_art, 4) if n_art else 0])
    for fila in ws.iter_rows(min_row=2):
        fila[4].number_format = "0.0%"
        for celda in fila[1:]:
            celda.alignment = Alignment(horizontal="center")
    ws.append([])
    ws.append(["Nota", "Basta un coautor de Zona 1 para que el criterio marque el articulo, "
                       "de modo que la deteccion a nivel de articulo supera a la deteccion a "
                       "nivel de autor.", "", "", ""])
    ws.cell(ws.max_row, 1).font = NEGRITA
    ws.cell(ws.max_row, 2).alignment = Alignment(wrap_text=True)

    ws = wb.create_sheet("3_Articulos_documentados")
    encabezar(ws, ["Anio", "Titulo", "DOI", "Autor(es) del conjunto documental",
                   "Coautores de Zona 1", "El criterio lo detecta",
                   "Seccion del expediente", "Citas (WoS 2016)"],
              [7, 72, 26, 30, 44, 13, 22, 12])
    for a in sorted(articulos, key=lambda x: (not x["detectado"], x["anio"] or 0)):
        ws.append([a["anio"], a["titulo"], a["doi"], ", ".join(a["documentados"]),
                   ", ".join(a["zona1"]) if a["zona1"] else "-",
                   "Si" if a["detectado"] else "No", a["seccion"], a["citas"]])
        for celda in ws[ws.max_row]:
            celda.alignment = Alignment(wrap_text=True, vertical="top")
        for j in (1, 6, 8):
            ws.cell(ws.max_row, j).alignment = Alignment(horizontal="center", vertical="top")
        if not a["detectado"]:
            ws.cell(ws.max_row, 6).font = ROJO_NEGRITA

    ws = wb.create_sheet("4_Tabla7")
    encabezar(ws, ["Corpus", "Publicaciones", "Detectadas por el criterio (Zona 1)",
                   "Con autor documentado y sin Zona 1", "Total con vinculo acreditado",
                   "Proporcion del vinculo no detectada"], [40, 13, 18, 18, 16, 17])
    for etq, c in (("Expediente europeo 2012", exp), ("Literatura general (WoS)", gen)):
        ws.append([f"{etq}, hasta {VENTANA_TABLA7[1]}", c["n"], c["z1"], c["no_detectadas"],
                   c["total"], round(c["pct_no_detectado"], 4)])
        ws.cell(ws.max_row, 6).number_format = "0.0%"
        for celda in ws[ws.max_row][1:]:
            celda.alignment = Alignment(horizontal="center")
    razon = (exp["pct_no_detectado"] / gen["pct_no_detectado"]) if gen["pct_no_detectado"] else 0
    ws.append([])
    ws.append(["Razon entre proporciones", round(razon, 1), "veces", "", "", ""])
    ws.cell(ws.max_row, 1).font = NEGRITA
    ws.append(["Nota", "Ambas filas sobre publicaciones distintas con anio de edicion hasta "
                       "2012, el horizonte propio de un expediente presentado ese anio, y con la "
                       "condicion explicita de al menos un autor del conjunto documental y ningun "
                       "autor de Zona 1.", "", "", "", ""])
    ws.cell(ws.max_row, 1).font = NEGRITA
    ws.cell(ws.max_row, 2).alignment = Alignment(wrap_text=True)
    ws.append([])
    ws.append(["SENSIBILIDAD AL ENCUADRE TEMPORAL", "", "", "", "", ""])
    ws.cell(ws.max_row, 1).font = NEGRITA
    ws.append(["Encuadre", "Expediente, no det. / total", "%",
               "General, no det. / total", "%", "Razon"])
    for celda in ws[ws.max_row]:
        celda.font = NEGRITA
    for etq, e, g in encuadres:
        rz = (e["pct_no_detectado"] / g["pct_no_detectado"]) if g["pct_no_detectado"] else 0
        ws.append([etq, f"{e['no_detectadas']} / {e['total']}",
                   round(e["pct_no_detectado"], 4),
                   f"{g['no_detectadas']} / {g['total']}",
                   round(g["pct_no_detectado"], 4), round(rz, 1)])
        ws.cell(ws.max_row, 3).number_format = "0.0%"
        ws.cell(ws.max_row, 5).number_format = "0.0%"

    ws = wb.create_sheet("5_Tabla8_anexo")
    encabezar(ws, ["Autor", "Vinculacion documentada, transcripta del aviso editorial",
                   "Tipo de vinculo", "Relacion previa", "En la red", "Clasificacion"],
              [22, 104, 24, 12, 10, 14])
    for r in orden:
        ws.append([r["autor"], r["vinculo"], r["tipo"], r["previa"],
                   "Si" if r["en_red"] else "No", r["clasificacion"]])
        for celda in ws[ws.max_row]:
            celda.alignment = Alignment(wrap_text=True, vertical="top")
        for j in (4, 5, 6):
            ws.cell(ws.max_row, j).alignment = Alignment(horizontal="center", vertical="top")
    ws.append([])
    ws.append(["Fuentes", "Correction (2018), Critical Reviews in Toxicology, "
                          "https://doi.org/10.1080/10408444.2018.1522175 · Expression of Concern "
                          "(2018a), https://doi.org/10.1080/10408444.2018.1522786 · Expression of "
                          "Concern (2018b), https://doi.org/10.1080/10408444.2018.1539570 · nota "
                          "de retractacion, https://doi.org/10.1016/j.yrtph.2025.106006",
               "", "", "", ""])
    ws.cell(ws.max_row, 1).font = NEGRITA
    ws.cell(ws.max_row, 2).alignment = Alignment(wrap_text=True)

    if perfil:
        ws = wb.create_sheet("6_Tabla6")
        encabezar(ws, ["Autor", "Clasificacion", "Articulos del expediente",
                       "Coautores", "Citacion de la obra", "Percentil",
                       "Citacion en el expediente", "Percentil",
                       "Formas del nombre en la base"],
                  [22, 14, 13, 11, 13, 11, 14, 11, 28])
        for f in perfil:
            ws.append([f["autor"], f["clasificacion"], f["n_refs"],
                       f["n_coautores"],
                       f["cit_obra"] if f["cit_obra"] is not None else "-",
                       round(f["pct_obra"], 1) if f["pct_obra"] is not None else "-",
                       f["cit_exped"] if f["cit_exped"] is not None else "-",
                       round(f["pct_exped"], 1) if f["pct_exped"] is not None else "-",
                       f["formas"]])
            for j in range(2, 9):
                ws.cell(ws.max_row, j).alignment = Alignment(horizontal="center")
        ws.append([])
        ws.append(["Nota", "La citacion se reporta como mediana de los articulos del autor y "
                           "no como total acumulado. Los percentiles corresponden a los "
                           "autores del corpus del expediente con citacion disponible: "
                           f"{ref6['n_obra']:,} para la obra completa, de mediana "
                           f"{ref6['med_obra']}, y {ref6['n_exped']:,} para los articulos "
                           f"del expediente, de mediana {ref6['med_exped']}. Las formas del "
                           "nombre que la base registra como nodos distintos de la misma "
                           "persona se reportan unidas; el analisis de red las trata "
                           "separadas.", "", "", "", "", "", "", ""])
        ws.cell(ws.max_row, 1).font = NEGRITA
        ws.cell(ws.max_row, 2).alignment = Alignment(wrap_text=True)

        ws = wb.create_sheet("7_Trayectorias")
        encabezar(ws, ["Autor", "Clasificacion", "Articulos en la base",
                       "Primer anio", "Ultimo anio", "Coautores en la base",
                       "De ellos en Zona 1", "%", "Referencias del expediente",
                       "Coautores del expediente", "Coautores de Zona 1"],
                  [22, 14, 12, 10, 10, 13, 12, 9, 13, 13, 46])
        for f in perfil:
            ws.append([f["autor"], f["clasificacion"], f["arts_base"],
                       f["anio_min"] or "-", f["anio_max"] or "-",
                       f["coaut_base"], f["coaut_base_z1"],
                       round(f["pct_coaut_z1"], 4) if f["pct_coaut_z1"] is not None else "-",
                       f["n_refs"], f["n_coautores"], f["coaut_z1_nombres"] or "-"])
            ws.cell(ws.max_row, 8).number_format = "0.0%"
            for j in range(2, 11):
                ws.cell(ws.max_row, j).alignment = Alignment(horizontal="center")
            ws.cell(ws.max_row, 11).alignment = Alignment(wrap_text=True, vertical="top")
        ws.append([])
        ws.append(["Nota", "Produccion y entorno de coautoria de cada autor en la base de "
                           "2018, no solo en el corpus del expediente. La columna de "
                           "coautores de Zona 1 es la que sostiene la lectura negativa del "
                           "criterio: un autor queda en Zona 2 cuando no firmo ningun "
                           "articulo junto a alguien con afiliacion corporativa declarada en "
                           "ese mismo articulo (Fontans-Alvarez et al., 2018). Tener "
                           "coautores de Zona 1 es compatible con esa situacion, porque esos "
                           "coautores pueden integrar la Zona 1 por articulos que el autor no "
                           "firmo, lo que lo ubica a distancia dos.",
                   "", "", "", "", "", "", "", "", "", ""])
        ws.cell(ws.max_row, 1).font = NEGRITA
        ws.cell(ws.max_row, 2).alignment = Alignment(wrap_text=True)

    wb.save(path)


# ─── INFORME ──────────────────────────────────────────────────────────────────

def escribir_informe(resuelto, articulos, exp, gen, encuadres, perfil, ref6,
                     origen, path):
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    en_red = [r for r in resuelto if r["en_red"]]
    det = [r for r in en_red if r["clasificacion"] == "Zona 1"]
    sin_k = [r for r in en_red if r["autor"] != AUTOR_SENSIBILIDAD]
    det_k = [r for r in sin_k if r["clasificacion"] == "Zona 1"]
    det_art = sum(1 for a in articulos if a["detectado"])
    razon = (exp["pct_no_detectado"] / gen["pct_no_detectado"]) if gen["pct_no_detectado"] else 0

    lineas = [
        "=" * 68,
        "  [VALIDACION DOCUMENTAL] - validar_conjunto_documental_v4",
        f"  Fecha: {ahora}",
        f"  Cruce de origen: {origen.name}",
        "=" * 68,
        "",
        "  CONJUNTO DE VALIDACION",
        f"    Autores del conjunto                    : {len(resuelto)}",
        f"    Presentes en la red de 2018             : {len(en_red)}",
        f"    Ausentes (no inciden en ninguna cifra)  : {len(resuelto) - len(en_red)}",
        "",
        "  SENSIBILIDAD A NIVEL DE AUTOR",
        f"    Detectados (Zona 1)                     : {len(det)} de {len(en_red)}"
        f" = {100 * len(det) / len(en_red):.1f}%",
        f"    No detectados (Zona 2)                  : {len(en_red) - len(det)}",
        f"    Excluyendo a {AUTOR_SENSIBILIDAD:14s}           : {len(det_k)} de {len(sin_k)}"
        f" = {100 * len(det_k) / len(sin_k):.1f}%",
        "",
        "    Detalle de los autores presentes en la red:",
    ]
    for r in sorted(en_red, key=lambda x: (x["clasificacion"] != "Zona 1", x["autor"])):
        lineas.append(f"      {r['autor']:22s} {r['clasificacion']:7s} {r['tipo']}")
    lineas += [
        "",
        "  SENSIBILIDAD A NIVEL DE ARTICULO",
        f"    Articulos con autor documentado         : {len(articulos)}",
        f"    Detectados por tener un autor de Zona 1 : {det_art}"
        f" = {100 * det_art / len(articulos):.1f}%" if articulos else "",
        f"    Invisibles al criterio estructural      : {len(articulos) - det_art}",
        "",
        "    Articulos invisibles al criterio:",
    ]
    for a in sorted((x for x in articulos if not x["detectado"]), key=lambda x: x["anio"] or 0):
        lineas.append(f"      {a['anio']}  {', '.join(a['documentados'])}")
        lineas.append(f"            {a['titulo'][:86]}")
    lineas += [
        "",
        "  PROPORCION DEL VINCULO NO DETECTADA (Tabla 7)",
        f"    {'corpus':38s} {'publ.':>7} {'Zona 1':>7} {'no det.':>8} {'total':>7} {'%':>7}",
        f"    {'Expediente 2012, hasta ' + str(VENTANA_TABLA7[1]):38s}"
        f" {exp['n']:>7,} {exp['z1']:>7,} {exp['no_detectadas']:>8} {exp['total']:>7,}"
        f" {100 * exp['pct_no_detectado']:>6.1f}%",
        f"    {'Literatura general, hasta ' + str(VENTANA_TABLA7[1]):38s}"
        f" {gen['n']:>7,} {gen['z1']:>7,} {gen['no_detectadas']:>8} {gen['total']:>7,}"
        f" {100 * gen['pct_no_detectado']:>6.1f}%",
        f"    Razon entre proporciones                : {razon:.1f} veces",
        "",
        "  SENSIBILIDAD AL ENCUADRE TEMPORAL",
        f"    {'encuadre':44s} {'expediente':>17} {'general':>17} {'razon':>7}",
    ]
    for etq, e, g in encuadres:
        ce = f"{e['no_detectadas']}/{e['total']}  {100 * e['pct_no_detectado']:.1f}%"
        cg = f"{g['no_detectadas']}/{g['total']}  {100 * g['pct_no_detectado']:.1f}%"
        rz = e["pct_no_detectado"] / g["pct_no_detectado"] if g["pct_no_detectado"] else 0
        lineas.append(f"    {etq:44s} {ce:>17} {cg:>17} {rz:>7.1f}")
    if perfil:
        lineas += [
            "",
            "  PERFIL DE LOS AUTORES DOCUMENTADOS EN EL CORPUS (Tabla 6)",
            f"    {'autor':20s} {'clasif.':>8} {'arts':>5} {'coaut':>6}"
            f" {'cit.obra':>9} {'pctil':>6} {'cit.exp':>8} {'pctil':>6}",
        ]
        for f in perfil:
            co = f["cit_obra"] if f["cit_obra"] is not None else "-"
            ce = f["cit_exped"] if f["cit_exped"] is not None else "-"
            po = f"{f['pct_obra']:.1f}" if f["pct_obra"] is not None else "-"
            pe = f"{f['pct_exped']:.1f}" if f["pct_exped"] is not None else "-"
            lineas.append(f"    {f['autor']:20s} {f['clasificacion']:>8}"
                          f" {f['n_refs']:>5} {f['n_coautores']:>6}"
                          f" {str(co):>9} {po:>6} {str(ce):>8} {pe:>6}")
        lineas += [
            "",
            f"    Distribucion de referencia: {ref6['n_obra']:,} autores del corpus con"
            f" citacion disponible, mediana {ref6['med_obra']};",
            f"    {ref6['n_exped']:,} con citacion en el expediente, mediana"
            f" {ref6['med_exped']}.",
            "",
            "  TRAYECTORIA EN LA BASE DE 2018",
            f"    {'autor':20s} {'clasif.':>8} {'arts':>5} {'anios':>11}"
            f" {'coaut':>6} {'de Zona 1':>10}",
        ]
        for f in perfil:
            rango = (f"{f['anio_min']}-{f['anio_max']}"
                     if f["anio_min"] else "-")
            cuota = (f"{f['coaut_base_z1']} ({100 * f['pct_coaut_z1']:.0f}%)"
                     if f["pct_coaut_z1"] is not None else "-")
            lineas.append(f"    {f['autor']:20s} {f['clasificacion']:>8}"
                          f" {f['arts_base']:>5} {rango:>11}"
                          f" {f['coaut_base']:>6} {cuota:>10}")
        lineas += [
            "",
            "    Un autor de Zona 2 puede tener coautores de Zona 1: esos coautores",
            "    la integran por articulos que el autor no firmo, lo que lo ubica a",
            "    distancia dos de la afiliacion corporativa declarada.",
        ]
    lineas += [
        "",
        "  Archivos generados:",
        f"    {OUTPUT_FILE.name}",
        f"    {REPORTE_FILE.name}",
    ]
    reporte = "\n".join(l for l in lineas if l != "")
    print(reporte)
    path.write_text(reporte + "\n", encoding="utf-8")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 68)
    print("  [VALIDACION DOCUMENTAL] - validar_conjunto_documental_v4")
    print("=" * 68)

    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

    print("\n[1/4] Cargando la base de 2018...")
    zona, arts_zona, registros = cargar_base()
    print(f"      {len(registros):,} registros, {len(zona):,} autores con zona.")

    print("[2/4] Cargando el cruce expediente x WoS...")
    apareadas, publicaciones, secciones, origen = cargar_cruce()
    if not apareadas:
        print("ERROR: no hay entradas apareadas para procesar.")
        return
    print(f"      {len(apareadas):,} entradas, {len(publicaciones):,} publicaciones distintas.")

    print("[3/4] Resolviendo el conjunto documental y los articulos...")
    resuelto = resolver_formas(zona)
    formas_doc = {f for r in resuelto for f in r["presentes"]}
    faltantes = [r["autor"] for r in resuelto if not r["presentes"]]
    if faltantes:
        print(f"      ausentes de la base: {', '.join(faltantes)}")

    ids_exp = set(publicaciones)
    arts_exp = collections.Counter()
    for clave in ids_exp:
        for a in registros.get(clave, {}).get("autores", []):
            if a in formas_doc:
                arts_exp[a] += 1

    articulos = articulos_documentados(ids_exp, registros, zona, formas_doc)
    for a in articulos:
        a["seccion"] = "; ".join(sorted(secciones.get(a["id"], [])))

    refs_de, autores_de, origen_red = cargar_instancias_expediente()
    wos_por_ref = {str(f.get("id_cita", "")).strip(): str(f.get("id_wos", "")).strip()
                   for f in apareadas}
    perfil, ref6 = perfil_documentados(resuelto, registros, zona, refs_de,
                                       autores_de, wos_por_ref)
    if perfil:
        print(f"      Tabla 6: {len(perfil)} autores, desde {origen_red.name}")

    exp = contar_corpus(ids_exp, registros, zona, formas_doc, VENTANA_TABLA7)
    gen = contar_corpus(registros.keys(), registros, zona, formas_doc, VENTANA_TABLA7)
    encuadres = [
        (f"Hasta {VENTANA_TABLA7[1]} (reportado en la Tabla 7)", exp, gen),
        (f"Ventana {VENTANA_REFERENCIA[0]}-{VENTANA_REFERENCIA[1]} (la del apartado 4.2)",
         contar_corpus(ids_exp, registros, zona, formas_doc, VENTANA_REFERENCIA),
         contar_corpus(registros.keys(), registros, zona, formas_doc, VENTANA_REFERENCIA)),
        ("Sin restriccion temporal",
         contar_corpus(ids_exp, registros, zona, formas_doc),
         contar_corpus(registros.keys(), registros, zona, formas_doc)),
    ]

    print("[4/4] Escribiendo planilla e informe...\n")
    escribir_planilla(resuelto, arts_zona, arts_exp, articulos, exp, gen,
                      encuadres, perfil, ref6, OUTPUT_FILE)
    escribir_informe(resuelto, articulos, exp, gen, encuadres, perfil, ref6,
                     origen, REPORTE_FILE)
    print(f"\nArchivos en {CARPETA_SALIDA}")


if __name__ == "__main__":
    main()
