"""
cruzar_publit_wos_v2.py
=======================
Cruza las listas de publicaciones de los expedientes de renovacion europea
del glifosato (presentaciones 2012 y 2020) contra la base de WoS del articulo
de 2018, y determina que proporcion de la literatura que la industria presento
al regulador tiene autores de Zona 1.

Estructura de carpetas del proyecto:
  proyecto/
    script/     los .py
    data/       fuentes originales, NUNCA se escribe aqui
    output/     todo lo que generan los scripts, una subcarpeta por script

Si se borra output/ y se vuelven a correr los scripts en orden, el resultado
debe ser identico. Por eso las salidas no van en data/.

Proyecto: Corpus curados sobre glifosato — comparacion antologia Rossi /
          Glyphosate Task Force / expediente regulatorio. FIC-UDELAR.
Contexto: Cruce corpus regulatorio x clasificacion Zona 1 / Zona 2

LIMITACION IMPORTANTE: la base de WoS llega hasta 2016. Las referencias
posteriores no pueden aparear; se reportan aparte como FUERA_DE_BASE y no
deben contarse como no apareadas. El cruce es completo para la presentacion
de 2012 y parcial para la de 2020.


Entradas:
  - data/Glif_-_Datos_ARS.xlsx       hojas 'Base' y 'AU Gephi'
  - output/publit/publit_v3.csv      - salida de extraer_publit_v3.py

Salidas:
  - cruce_publit_v2.xlsx  hojas Resumen / Revisar / Cruce / Autores
  - cruce_publit_v2.csv   para Gephi y UCINET
  - reporte_cruzar_publit_wos_v2.txt
  - cruzar_publit_wos_v2_errores.log

Sensibilidad de salidas:
  NO SENSIBLE : todas (metadatos bibliograficos publicos)

Uso: python cruzar_publit_wos_v2.py

Requisitos: pip install openpyxl

Historial:
  v1 - version inicial: apareo por DOI y titulo normalizado, asignacion de
       zona, estandarizacion por area tematica de WoS y comparacion de la
       proporcion de Zona 1 segun presentacion, seccion y marca K.
  v2 - hace auditable la estandarizacion, que en la v1 quedaba reducida a un
       unico numero global:
       (a) las categorias tematicas de WoS se arrastran a la salida en el
           campo 'wc_wos'; antes se usaban internamente y se descartaban;
       (b) se agrega la hoja 'Estandarizacion' con el calculo por area
           —articulos del expediente, Zona 1 observados, tasa de referencia
           de la base, esperados y razon observado/esperado— mas la fila de
           total global. Es la tabla que va al articulo;
       (c) se agrega la constante SOLO_PRESENTACION para acotar el analisis
           a una de las dos presentaciones del expediente.
"""

import csv
import re
import unicodedata
from difflib import SequenceMatcher
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
CARPETA_SALIDA = CARPETA_OUT / "cruce_publit"

EXCEL_FILE  = CARPETA_DATOS / "Glif_-_Datos_ARS.xlsx"
GTF_FILE    = CARPETA_OUT / "publit" / "publit_v3.csv"

CRUCE_FILE   = CARPETA_SALIDA / "cruce_publit_v2.csv"
AUTORES_FILE = CARPETA_SALIDA / "autores_publit_v2.csv"
XLSX_FILE    = CARPETA_SALIDA / "cruce_publit_v2.xlsx"
REPORTE_FILE = CARPETA_SALIDA / "reporte_cruzar_publit_wos_v2.txt"
LOG_FILE     = CARPETA_SALIDA / "cruzar_publit_wos_v2_errores.log"

ANIO_TOPE_BASE = 2016      # ultimo anio cubierto por la base de WoS

# Presentacion del expediente a analizar. La lista de 2020 consigna solo el
# primer autor, de modo que la autoria no puede recuperarse de la propia
# lista; por eso el analisis se acota a 2012. None procesa ambas.
SOLO_PRESENTACION = "2012"

HOJA_BASE    = "Base"
HOJA_ZONAS   = "AU Gephi"

# Columnas de la hoja Base (formato de exportacion WoS)
COL_ID, COL_AU, COL_TI, COL_PY, COL_DI, COL_PM = 0, 2, 9, 45, 55, 62
COL_WC = 58                # 'Web of Science Categories', para estandarizar
# Columnas de la hoja AU Gephi
COL_Z_LABEL, COL_Z_ZONA = 2, 4

VENTANA = (2001, 2012)     # ventana de referencia para la linea de base
SIM_TITULO = 0.92          # similitud minima para aceptar apareo difuso
DELTA_ANIO = 1             # tolerancia de anio en el apareo difuso

CAMPOS_CRUCE = ["id_cita", "presentacion", "seccion", "categoria_pdf",
                "evaluacion", "titulo", "anio", "doi", "wc_wos", "n_wc",
                "estado_apareo", "metodo_apareo", "similitud_apareo",
                "id_wos", "titulo_wos", "anio_wos",
                "n_autores", "n_zona1", "n_zona2", "n_sin_zona", "tiene_zona1"]
CAMPOS_AUTORES = ["id_cita", "id_wos", "presentacion", "seccion",
                  "evaluacion", "autor", "zona", "wc_wos"]


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def log_error(mensaje):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{ts}] {mensaje}\n")


def leer_csv_auto(path):
    if not path.exists():
        print(f"  No se encontro {path.name}")
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
    print(f"  No se pudo leer {path.name}")
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


def norm_titulo(txt):
    """Clave de titulo: sin tildes, sin puntuacion, minusculas."""
    t = sin_tildes(txt).lower()
    t = re.sub(r"[^a-z0-9 ]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def norm_doi(txt):
    t = str(txt or "").strip().lower()
    t = re.sub(r"^https?://(dx\.)?doi\.org/", "", t)
    return t.strip(" .")


def norm_autor(txt):
    """Formato de AU Gephi: APELLIDO, IN (mayusculas, espacios colapsados)."""
    return re.sub(r"\s+", " ", sin_tildes(txt)).strip().upper()


def anio_int(v):
    s = re.sub(r"[^0-9]", "", str(v or ""))
    return int(s[:4]) if len(s) >= 4 else None


# ─── CARGA ────────────────────────────────────────────────────────────────────

def cargar_zonas(wb):
    """Devuelve {autor_normalizado: zona}."""
    print("\n[1/4] Cargando clasificacion de zonas...")
    zonas = {}
    filas = wb[HOJA_ZONAS].iter_rows(values_only=True)
    next(filas, None)
    for r in filas:
        if r[COL_Z_LABEL] is None:
            continue
        z = str(r[COL_Z_ZONA]).strip()
        if z in ("1", "2"):
            zonas[norm_autor(r[COL_Z_LABEL])] = z
    n1 = sum(1 for v in zonas.values() if v == "1")
    print(f"  {len(zonas):,} autores | Zona 1: {n1:,} | Zona 2: {len(zonas)-n1:,}")
    return zonas


def cargar_base(wb):
    """Devuelve los registros de WoS y los indices de apareo."""
    print("\n[2/4] Cargando base de WoS...")
    registros, por_doi, por_titulo = [], {}, {}
    filas = wb[HOJA_BASE].iter_rows(values_only=True)
    next(filas, None)
    for r in filas:
        if r[COL_ID] is None:
            continue
        reg = {
            "id_wos": str(r[COL_ID]),
            "titulo": str(r[COL_TI] or ""),
            "anio": anio_int(r[COL_PY]),
            "doi": norm_doi(r[COL_DI]),
            "pmid": str(r[COL_PM] or "").strip(),
            "autores": [a.strip() for a in str(r[COL_AU] or "").split(";") if a.strip()],
            "wc": [x.strip() for x in str(r[COL_WC] or "").split(";") if x.strip()],
        }
        reg["clave_titulo"] = norm_titulo(reg["titulo"])
        registros.append(reg)
        if reg["doi"]:
            por_doi.setdefault(reg["doi"], reg)
        if reg["clave_titulo"]:
            por_titulo.setdefault(reg["clave_titulo"], reg)
    print(f"  {len(registros):,} registros | con DOI: {len(por_doi):,}")
    return registros, por_doi, por_titulo


# ─── PROCESAMIENTO ────────────────────────────────────────────────────────────

def aparear(cita, por_doi, por_titulo, registros):
    """Aparea una referencia del GTF contra la base. DOI > titulo exacto > difuso."""
    doi = norm_doi(cita.get("doi"))
    if doi and doi in por_doi:
        return por_doi[doi], "DOI", 1.0

    clave = norm_titulo(cita.get("titulo", ""))
    if clave and clave in por_titulo:
        return por_titulo[clave], "TITULO_EXACTO", 1.0

    # Difuso, acotado por anio para no comparar contra los 8.174 registros.
    # El DOI cubre solo el 67,8% de la base, asi que sin esta pasada se
    # perderian apareos validos por diferencias menores de puntuacion.
    anio = anio_int(cita.get("anio"))
    if not clave or anio is None:
        return None, "", 0.0
    mejor, mejor_sim = None, 0.0
    for reg in registros:
        if reg["anio"] is None or abs(reg["anio"] - anio) > DELTA_ANIO:
            continue
        if not reg["clave_titulo"]:
            continue
        if abs(len(reg["clave_titulo"]) - len(clave)) > 25:
            continue
        sim = SequenceMatcher(None, clave, reg["clave_titulo"]).ratio()
        if sim > mejor_sim:
            mejor, mejor_sim = reg, sim
    if mejor and mejor_sim >= SIM_TITULO:
        return mejor, "TITULO_DIFUSO", round(mejor_sim, 3)
    return None, "", round(mejor_sim, 3)


def cruzar(citas, por_doi, por_titulo, registros, zonas):
    print("\n[3/4] Cruzando el corpus del GTF contra la base...")
    cruce, autores_out = [], []

    for i, c in enumerate(citas, 1):
        anio = anio_int(c.get("anio"))
        if anio and anio > ANIO_TOPE_BASE:
            # posterior al cierre de la base: no puede aparear, no es un fallo
            reg, metodo, sim = None, "FUERA_DE_BASE", 0.0
        else:
            reg, metodo, sim = aparear(c, por_doi, por_titulo, registros)
        fila = {
            "id_cita": c.get("id_ref", c.get("id_cita", "")),
            "presentacion": c.get("presentacion", ""),
            "seccion": c.get("seccion", ""),
            "categoria_pdf": c.get("categoria", ""),
            "evaluacion": c.get("evaluacion", ""),
            "titulo": c.get("titulo", ""),
            "anio": c.get("anio", ""),
            "doi": c.get("doi", ""),
            "wc_wos": "", "n_wc": "",
            "estado_apareo": ("APAREADO" if reg else
                              "FUERA_DE_BASE" if metodo == "FUERA_DE_BASE"
                              else "NO_APAREADO"),
            "metodo_apareo": metodo,
            "similitud_apareo": sim,
            "id_wos": "", "titulo_wos": "", "anio_wos": "",
            "n_autores": "", "n_zona1": "", "n_zona2": "", "n_sin_zona": "",
            "tiene_zona1": "",
        }
        if reg:
            z = [zonas.get(norm_autor(a)) for a in reg["autores"]]
            fila.update({
                "wc_wos": "; ".join(reg["wc"]),
                "n_wc": len(reg["wc"]),
                "id_wos": reg["id_wos"],
                "titulo_wos": reg["titulo"],
                "anio_wos": reg["anio"] or "",
                "n_autores": len(reg["autores"]),
                "n_zona1": sum(1 for x in z if x == "1"),
                "n_zona2": sum(1 for x in z if x == "2"),
                "n_sin_zona": sum(1 for x in z if x is None),
                "tiene_zona1": 1 if any(x == "1" for x in z) else 0,
            })
            for a in reg["autores"]:
                autores_out.append({
                    "id_cita": fila["id_cita"], "id_wos": reg["id_wos"],
                    "presentacion": fila["presentacion"],
                    "seccion": fila["seccion"],
                    "evaluacion": fila["evaluacion"],
                    "autor": norm_autor(a),
                    "zona": zonas.get(norm_autor(a), ""),
                    "wc_wos": fila["wc_wos"],
                })
        cruce.append(fila)
        if i % 100 == 0:
            print(f"  [{i}/{len(citas)}] procesadas")

    return cruce, autores_out


def linea_base(registros, zonas):
    """Proporcion de articulos con >=1 autor de Zona 1, dentro y fuera de ventana."""
    dentro = fuera = dentro_z1 = fuera_z1 = 0
    for reg in registros:
        if reg["anio"] is None:
            continue
        tiene = any(zonas.get(norm_autor(a)) == "1" for a in reg["autores"])
        if VENTANA[0] <= reg["anio"] <= VENTANA[1]:
            dentro += 1
            dentro_z1 += tiene
        else:
            fuera += 1
            fuera_z1 += tiene
    return {"n_ventana": dentro, "z1_ventana": dentro_z1,
            "n_fuera": fuera, "z1_fuera": fuera_z1}


def estandarizar(registros, zonas, id_apareados):
    """Estandarizacion indirecta por area tematica de WoS.

    Comparar el corpus del GTF contra el total de la base confunde dos cosas:
    la logica de curacion y la composicion tematica. La tasa de Zona 1 en la
    base va de ~53% en Agronomy a ~10% en Environmental Sciences, y el corpus
    del GTF es ambiental-toxicologico. Aca se calcula, para cada area, la tasa
    de Zona 1 en la ventana de referencia, y se aplican esas tasas a la mezcla
    tematica del corpus apareado: el resultado es cuantos articulos con autor
    de Zona 1 cabria esperar si el GTF no hubiera seleccionado por vinculacion.
    """
    tot, z1 = {}, {}
    for reg in registros:
        if reg["anio"] is None or not (VENTANA[0] <= reg["anio"] <= VENTANA[1]):
            continue
        tiene = any(zonas.get(norm_autor(a)) == "1" for a in reg["autores"])
        for wc in reg["wc"]:
            tot[wc] = tot.get(wc, 0) + 1
            z1[wc] = z1.get(wc, 0) + tiene

    global_z1 = sum(z1.values()) / sum(tot.values()) if tot else 0.0

    esperado, observado, n, sin_wc = 0.0, 0, 0, 0
    for reg in registros:
        if reg["id_wos"] not in id_apareados:
            continue
        if reg["anio"] is None or not (VENTANA[0] <= reg["anio"] <= VENTANA[1]):
            continue
        n += 1
        observado += any(zonas.get(norm_autor(a)) == "1" for a in reg["autores"])
        tasas = [z1[w] / tot[w] for w in reg["wc"] if tot.get(w, 0) >= 20]
        if tasas:
            esperado += sum(tasas) / len(tasas)
        else:
            esperado += global_z1
            sin_wc += 1

    # Detalle por area: como se reparte el corpus apareado entre areas y que
    # tasa de referencia le corresponde a cada una. Es el calculo que sostiene
    # la razon global, expuesto para que pueda auditarse.
    detalle = {}
    for reg in registros:
        if reg["id_wos"] not in id_apareados:
            continue
        if reg["anio"] is None or not (VENTANA[0] <= reg["anio"] <= VENTANA[1]):
            continue
        tiene = any(zonas.get(norm_autor(a)) == "1" for a in reg["autores"])
        for w in reg["wc"]:
            d = detalle.setdefault(w, {"n": 0, "obs": 0, "ref_z1": z1.get(w, 0),
                                       "ref_n": tot.get(w, 0)})
            d["n"] += 1
            d["obs"] += tiene

    return {"n": n, "observado": observado, "esperado": esperado,
            "sin_wc": sin_wc,
            "razon": (observado / esperado) if esperado else 0.0,
            "detalle": detalle,
            "tasas_area": {w: (z1[w] / tot[w], tot[w]) for w in tot if tot[w] >= 60}}


def chi2_2x2(a, b, c, d):
    """Chi-cuadrado con correccion de Yates. Devuelve (chi2, significativo_05)."""
    n = a + b + c + d
    if min(a + b, c + d, a + c, b + d) == 0:
        return 0.0, False
    chi2 = n * (abs(a * d - b * c) - n / 2) ** 2 / ((a + b) * (c + d) * (a + c) * (b + d))
    return round(chi2, 2), chi2 > 3.841


# ─── SALIDA EN EXCEL ──────────────────────────────────────────────────────────

FUENTE      = "Arial"
GRIS        = PatternFill("solid", fgColor="D9D9D9")
AMARILLO    = PatternFill("solid", fgColor="FFF2CC")   # necesita decision
NARANJA     = PatternFill("solid", fgColor="FCE4D6")   # no apareado
VERDE       = PatternFill("solid", fgColor="E2EFDA")   # apareo seguro

# Ancho de columna por nombre de campo. Los titulos son largos y sin esto
# el libro es tan ilegible como el CSV.
ANCHOS = {
    "id_cita": 10, "titulo": 62, "anio": 7, "doi": 26, "categorias": 30,
    "n_categorias": 8, "estado_apareo": 14, "metodo_apareo": 16,
    "similitud_apareo": 10, "id_wos": 9, "titulo_wos": 62, "anio_wos": 9,
    "n_autores": 10, "n_zona1": 9, "n_zona2": 9, "n_sin_zona": 11,
    "tiene_zona1": 11, "autor": 30, "zona": 7, "revisado_ok": 14,
}


def _volcar_hoja(ws, filas, campos, colorear=None, congelar="A2"):
    """Escribe una tabla con encabezado fijo, filtro y anchos utiles."""
    ws.append(campos)
    for c in ws[1]:
        c.font = Font(name=FUENTE, bold=True)
        c.fill = GRIS
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28

    for fila in filas:
        ws.append([fila.get(k, "") for k in campos])

    for i, campo in enumerate(campos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ANCHOS.get(campo, 14)

    for fila_ws in ws.iter_rows(min_row=2):
        for c in fila_ws:
            c.font = Font(name=FUENTE)
            c.alignment = Alignment(vertical="top", wrap_text=False)
        if colorear:
            relleno = colorear(filas[fila_ws[0].row - 2])
            if relleno:
                for c in fila_ws:
                    c.fill = relleno

    if len(filas):
        ws.auto_filter.ref = f"A1:{get_column_letter(len(campos))}{len(filas) + 1}"
    ws.freeze_panes = congelar


def _color_cruce(fila):
    if fila["estado_apareo"] == "FUERA_DE_BASE":
        return None
    if fila["estado_apareo"] != "APAREADO":
        return NARANJA
    if fila["metodo_apareo"] == "TITULO_DIFUSO":
        return AMARILLO
    return VERDE


def escribir_excel(cruce, autores_out, base_stats, path):
    """Libro de trabajo con las filas dudosas separadas del resto."""
    print("\n  Escribiendo el libro de Excel...")
    wb = openpyxl.Workbook()

    col = {c: get_column_letter(i) for i, c in enumerate(CAMPOS_CRUCE, 1)}
    n_filas = len(cruce) + 1

    # --- Resumen (con formulas: si corregis 'Revisar', los numeros cambian) ---
    ws = wb.active
    ws.title = "Resumen"
    bn, bz = base_stats["n_ventana"], base_stats["z1_ventana"]

    bloques = [
        ("CRUCE GTF x WoS", ""),
        ("", ""),
        ("APAREO", ""),
        ("Referencias del GTF", f"=COUNTA(Cruce!{col['id_cita']}2:"
                               f"{col['id_cita']}{n_filas})"),
        ("Apareadas", f"=COUNTIF(Cruce!{col['estado_apareo']}2:"
                      f"{col['estado_apareo']}{n_filas},\"APAREADO\")"),
        ("No apareadas", f"=COUNTIF(Cruce!{col['estado_apareo']}2:"
                         f"{col['estado_apareo']}{n_filas},\"NO_APAREADO\")"),
        ("  por DOI", f"=COUNTIF(Cruce!{col['metodo_apareo']}2:"
                      f"{col['metodo_apareo']}{n_filas},\"DOI\")"),
        ("  por titulo exacto", f"=COUNTIF(Cruce!{col['metodo_apareo']}2:"
                                f"{col['metodo_apareo']}{n_filas},\"TITULO_EXACTO\")"),
        ("  por titulo difuso (revisar)", f"=COUNTIF(Cruce!{col['metodo_apareo']}2:"
                                          f"{col['metodo_apareo']}{n_filas},\"TITULO_DIFUSO\")"),
        ("", ""),
        ("RESULTADO PRINCIPAL", ""),
        ("Articulos del GTF con >=1 autor de Zona 1",
         f"=COUNTIFS(Cruce!{col['estado_apareo']}2:{col['estado_apareo']}{n_filas},"
         f"\"APAREADO\",Cruce!{col['tiene_zona1']}2:{col['tiene_zona1']}{n_filas},1)"),
        ("Sobre el total apareado", "=IFERROR(B12/B5,\"\")"),
        ("", ""),
        (f"Linea de base WoS {VENTANA[0]}-{VENTANA[1]}: articulos", bn),
        ("  de ellos con >=1 autor de Zona 1", bz),
        ("  proporcion", "=IFERROR(B16/B15,\"\")"),
        ("", ""),
        ("Diferencia (GTF - linea de base)", "=IFERROR(B13-B17,\"\")"),
    ]
    for etiqueta, valor in bloques:
        ws.append([etiqueta, valor])

    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 16
    for fila in ws.iter_rows():
        for c in fila:
            c.font = Font(name=FUENTE)
    for r in (1, 3, 11):
        ws.cell(row=r, column=1).font = Font(name=FUENTE, bold=True, size=12)
    for r in (13, 17, 19):
        ws.cell(row=r, column=2).number_format = "0.0%"
    ws.cell(row=19, column=1).font = Font(name=FUENTE, bold=True)
    ws.cell(row=19, column=2).fill = AMARILLO
    ws.append([])
    ws.append(["Linea de base calculada sobre la hoja 'Base' del Excel de 2018,"])
    ws.append(["restringida a la ventana declarada por el GTF. Valores fijos."])

    # --- Estandarizacion: la Tabla 3 del articulo, con su calculo a la vista ---
    det = base_stats["estand"]["detalle"]
    filas_est = []
    for w, d in sorted(det.items(), key=lambda x: -x[1]["n"]):
        if d["n"] < 10 or d["ref_n"] < 20:
            continue
        tasa_ref = d["ref_z1"] / d["ref_n"] if d["ref_n"] else 0
        filas_est.append({
            "area_wos": w,
            "arts_expediente": d["n"],
            "z1_expediente": d["obs"],
            "pct_expediente": round(100 * d["obs"] / d["n"], 1),
            "arts_base": d["ref_n"],
            "z1_base": d["ref_z1"],
            "pct_base": round(100 * tasa_ref, 1),
            "esperados": round(d["n"] * tasa_ref, 1),
            "razon_obs_esp": (round(d["obs"] / (d["n"] * tasa_ref), 2)
                              if d["n"] * tasa_ref else ""),
        })
    campos_est = ["area_wos", "arts_expediente", "z1_expediente", "pct_expediente",
                  "arts_base", "z1_base", "pct_base", "esperados", "razon_obs_esp"]
    ws = wb.create_sheet("Estandarizacion")
    ws.append(["Calculo de la comparacion estandarizada por area tematica de WoS"])
    ws.append(["Se muestran las areas con al menos 10 articulos apareados del "
               "expediente y 20 en la base de referencia."])
    ws.append([])
    ws.append(campos_est)
    for c in ws[4]:
        c.font = Font(name=FUENTE, bold=True)
        c.fill = GRIS
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for f in filas_est:
        ws.append([f[k] for k in campos_est])
    fila_tot = len(filas_est) + 5
    ws.cell(row=fila_tot, column=1, value="TOTAL (global, todas las areas)").font = \
        Font(name=FUENTE, bold=True)
    ws.cell(row=fila_tot, column=3, value=base_stats["estand"]["observado"])
    ws.cell(row=fila_tot, column=8,
            value=round(base_stats["estand"]["esperado"], 1))
    ws.cell(row=fila_tot, column=9,
            value=round(base_stats["estand"]["razon"], 2)).font = \
        Font(name=FUENTE, bold=True)
    ws.cell(row=fila_tot, column=9).fill = AMARILLO
    for i, campo in enumerate(campos_est, 1):
        ws.column_dimensions[get_column_letter(i)].width = \
            44 if campo == "area_wos" else 15
    for f in ws.iter_rows(min_row=5, max_row=fila_tot):
        for c in f:
            if not c.font.bold:
                c.font = Font(name=FUENTE)
    ws.freeze_panes = "B5"

    # --- Revisar: solo lo que necesita decision humana ---
    dudosas = [dict(f, revisado_ok="") for f in cruce
               if f["estado_apareo"] == "NO_APAREADO"
               or f["metodo_apareo"] == "TITULO_DIFUSO"]
    campos_rev = ["revisado_ok", "id_cita", "titulo", "anio", "doi",
                  "estado_apareo", "metodo_apareo", "similitud_apareo",
                  "titulo_wos", "anio_wos", "id_wos"]
    ws = wb.create_sheet("Revisar")
    _volcar_hoja(ws, dudosas, campos_rev, colorear=_color_cruce, congelar="C2")
    ws.cell(row=1, column=1).comment = None

    # --- Cruce completo y autores ---
    _volcar_hoja(wb.create_sheet("Cruce"), cruce, CAMPOS_CRUCE,
                 colorear=_color_cruce, congelar="C2")
    _volcar_hoja(wb.create_sheet("Autores"), autores_out, CAMPOS_AUTORES)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    print(f"  {path.name}: {len(dudosas)} filas para revisar en la hoja 'Revisar'")
    return len(dudosas)


# ─── INFORME ──────────────────────────────────────────────────────────────────

def tabla_por(cruce, campo):
    """Proporcion de Zona 1 agrupada por un campo del corpus regulatorio."""
    grupos = {}
    for r in cruce:
        if r["estado_apareo"] != "APAREADO":
            continue
        k = r.get(campo) or "(sin dato)"
        g = grupos.setdefault(k, [0, 0])
        g[0] += 1
        g[1] += 1 if r["tiene_zona1"] == 1 else 0
    return grupos


def escribir_informe(citas, cruce, base_stats, path):
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    total = len(cruce)
    apareados = [r for r in cruce if r["estado_apareo"] == "APAREADO"]
    fuera = [r for r in cruce if r["estado_apareo"] == "FUERA_DE_BASE"]
    na = [r for r in cruce if r["estado_apareo"] == "NO_APAREADO"]
    n_ap = len(apareados)
    elegibles = total - len(fuera)
    pct = lambda n, d: f"{(100 * n / d):.1f}%" if d else "-"

    metodos = {}
    for r in apareados:
        metodos[r["metodo_apareo"]] = metodos.get(r["metodo_apareo"], 0) + 1

    con_z1 = sum(1 for r in apareados if r["tiene_zona1"] == 1)
    bn, bz = base_stats["n_ventana"], base_stats["z1_ventana"]
    est = base_stats["estand"]
    chi2, signif = chi2_2x2(con_z1, n_ap - con_z1,
                            bz - con_z1, (bn - bz) - (n_ap - con_z1))

    lineas = [
        "=" * 68,
        "  [CRUCE PUBLIT x WoS] - cruzar_publit_wos_v2",
        f"  Fecha: {ahora}",
        "=" * 68,
        "",
        "  APAREO",
        f"    Referencias del expediente        : {total:,}",
        f"    Posteriores a {ANIO_TOPE_BASE} (fuera de base) : {len(fuera):,} "
        f"({pct(len(fuera), total)})",
        f"    Elegibles para apareo             : {elegibles:,}",
        f"    Apareadas                         : {n_ap:,} "
        f"({pct(n_ap, elegibles)} de las elegibles)",
        f"    No apareadas                      : {len(na):,} "
        f"({pct(len(na), elegibles)})",
        "",
        "    Metodo de apareo:",
    ]
    for m, n in sorted(metodos.items(), key=lambda x: -x[1]):
        lineas.append(f"      {m:<20} {n:>5}  ({pct(n, n_ap)})")

    lineas += [
        "",
        "  RESULTADO PRINCIPAL",
        f"    Con >=1 autor de Zona 1 : {con_z1:,} / {n_ap:,} = {pct(con_z1, n_ap)}",
        f"    Linea de base WoS {VENTANA[0]}-{VENTANA[1]} : "
        f"{bz:,} / {bn:,} = {pct(bz, bn)}",
        f"    Chi-cuadrado (Yates) : {chi2}   "
        f"{'DIFERENCIA SIGNIFICATIVA (p<0.05)' if signif else 'sin diferencia significativa'}",
        "",
        "  COMPARACION ESTANDARIZADA POR AREA TEMATICA",
        f"    Observado : {est['observado']:,}",
        f"    Esperado  : {est['esperado']:.1f} bajo la mezcla tematica del corpus",
        f"    Razon obs/esp : {est['razon']:.2f}   "
        f"{'(mas Zona 1 de lo esperable)' if est['razon'] > 1.15 else '(menos Zona 1 de lo esperable)' if est['razon'] < 0.85 else '(sin apartamiento relevante)'}",
    ]

    for campo, titulo in [("presentacion", "Por presentacion"),
                          ("seccion", "Por seccion del expediente"),
                          ("evaluacion", "Por marca de evaluacion (K = resumen Tier II)"),
                          ("categoria_pdf", "Por categoria transcripta del PDF")]:
        grupos = tabla_por(cruce, campo)
        if len(grupos) <= 1:
            continue
        lineas += ["", f"  {titulo}:",
                   f"    {'valor':<28} {'apareadas':>10} {'con Zona 1':>11} {'%':>7}"]
        for k, (n, z) in sorted(grupos.items(), key=lambda x: -x[1][0])[:12]:
            lineas.append(f"    {str(k)[:26]:<28} {n:>10} {z:>11} "
                          f"{pct(z, n):>7}")

    lineas += [
        "",
        "  NOTA: las filas FUERA_DE_BASE no son fallos de apareo: corresponden",
        f"        a publicaciones posteriores a {ANIO_TOPE_BASE}, fuera del alcance",
        "        de la base de 2018. Declararlo asi en la metodologia.",
        "",
        "  Archivos generados:",
        f"    {XLSX_FILE.name}   <- abrir este para revisar",
        f"    {CRUCE_FILE.name}",
        f"    {AUTORES_FILE.name}",
        f"    {REPORTE_FILE.name}",
        "",
    ]

    reporte = "\n".join(lineas)
    print("\n" + reporte)
    path.write_text(reporte + "\n", encoding="utf-8")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 64)
    print("  [CRUCE PUBLIT x WoS] - cruzar_publit_wos_v2")
    print("=" * 64)

    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    print(f"  data   : {CARPETA_DATOS}")
    print(f"  output : {CARPETA_SALIDA}\n")

    if not EXCEL_FILE.exists():
        print(f"ERROR: no se encontro {EXCEL_FILE}")
        return
    citas, _ = leer_csv_auto(GTF_FILE)
    if not citas:
        print(f"ERROR: no se pudo leer {GTF_FILE}")
        return

    if SOLO_PRESENTACION:
        citas = [c for c in citas
                 if str(c.get("presentacion", "")).strip() == SOLO_PRESENTACION]
        print(f"  Filtrado a la presentacion {SOLO_PRESENTACION}: "
              f"{len(citas):,} referencias\n")
        if not citas:
            print("ERROR: ninguna referencia en esa presentacion.")
            return

    wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    zonas = cargar_zonas(wb)
    registros, por_doi, por_titulo = cargar_base(wb)

    cruce, autores_out = cruzar(citas, por_doi, por_titulo, registros, zonas)
    escribir_csv(CRUCE_FILE, cruce, CAMPOS_CRUCE)
    escribir_csv(AUTORES_FILE, autores_out, CAMPOS_AUTORES)

    print("\n[4/4] Calculando linea de base...")
    stats = linea_base(registros, zonas)
    ids = {r["id_wos"] for r in cruce if r["estado_apareo"] == "APAREADO"}
    stats["estand"] = estandarizar(registros, zonas, ids)
    escribir_excel(cruce, autores_out, stats, XLSX_FILE)
    escribir_informe(citas, cruce, stats, REPORTE_FILE)
    print(f"Archivos en {CARPETA_SALIDA}")


if __name__ == "__main__":
    main()
