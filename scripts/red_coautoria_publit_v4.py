"""
red_coautoria_publit_v4.py
==========================
Genera la red de coautoria del corpus del expediente europeo de 2012 en
formato Gephi: una lista de nodos con sus atributos y una lista de aristas
ponderada por numero de trabajos en comun.

Proyecto: Corpus curados sobre glifosato. FIC-UDELAR.
Contexto: Figura 1 del articulo · red de coautoria por grupo de vinculacion

POR QUE LA HOJA 'Instancias': cubre las 1.078 referencias y trae el campo
'grupo' con las cuatro categorias (CONTROL / ZONA_1 / ZONA_2 /
SIN_INFORMACION), que es lo que la figura debe mostrar. La hoja 'Autores' de
cruce_publit_v2 solo cubre los 749 articulos apareados y no distingue el
grupo de control, aunque aporta el area tematica de WoS: se usa como
enriquecimiento opcional.

Estructura de carpetas del proyecto:
  proyecto/
    script/     los .py
    data/       fuentes originales, NUNCA se escribe aqui
    output/     todo lo que generan los scripts, una subcarpeta por script

Entradas:
  - output/cruce_autores/cruce_autores_publit_v3.xlsx   hoja 'Instancias'
  - output/cruce_publit/cruce_publit_v2.csv             (opcional, para wc_wos)
  - data/Glif_-_Datos_ARS.xlsx   hojas 'Base' y 'AU Gephi' (para citacion)

Salidas:
  - output/red/nodos_publit_v4.csv     - para importar en Gephi como nodos
  - output/red/aristas_publit_v4.csv   - para importar en Gephi como aristas
  - output/red/red_publit_v4.gexf      - la red completa, lista para abrir
  - output/red/perfil_grupos_publit_v4.csv - Tabla 5 del articulo
  - output/red/reporte_red_publit_v4.txt  - metricas de la red y perfil

Sensibilidad de salidas:
  NO SENSIBLE : todas (metadatos bibliograficos publicos)

Uso: python red_coautoria_publit_v4.py

Requisitos: pip install openpyxl

Historial:
  v1 - version inicial: aristas por coautoria dentro de cada referencia,
       ponderadas por trabajos en comun; atributos de nodo por grupo, zona,
       productividad y seccion; metricas de componente y grado por grupo.
  v2 - agrega el perfil comparado de los cuatro grupos, que es la Tabla 5 del
       articulo: productividad en el expediente, coautores distintos,
       presencia en el componente gigante y citacion. La citacion se calcula
       en dos versiones, porque responden preguntas distintas: la de los
       articulos del expediente (que literatura se presento al regulador) y
       la de toda la obra del autor en la base de 2018 (que peso tiene en el
       campo). Solo esta disponible para los autores presentes en la base,
       de modo que las medianas se calculan sobre ese subconjunto.
  v3 - agrega la ubicacion individual de cada autor del grupo de control
       dentro de la distribucion de citacion. Con seis casos, una mediana de
       grupo no es interpretable: lo que si lo es, es la posicion de cada
       autor respecto de la distribucion completa (2.000+ autores) y respecto
       de las medianas de Zona 1 y Zona 2. Sale como
       control_citacion_publit_v4.csv.
  v4 - agrega medidas de centralidad, que es lo que un analisis de redes
       requiere y la v3 no tenia: fuerza (grado ponderado), coeficiente de
       clustering local e intermediacion. La intermediacion se calcula dos
       veces —sobre el grafo completo y restringida al componente principal—
       porque en una red con 505 componentes la mayoria de los nodos no tiene
       nada que intermediar y el promedio global seria enganoso.
"""

import csv
import re
from pathlib import Path
from datetime import datetime
from itertools import combinations
from collections import defaultdict

import openpyxl

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

CARPETA_SCRIPT = Path(__file__).resolve().parent
CARPETA_BASE   = CARPETA_SCRIPT.parent
CARPETA_DATOS  = CARPETA_BASE / "data"
CARPETA_OUT    = CARPETA_BASE / "output"
CARPETA_SALIDA = CARPETA_OUT / "red"

INSTANCIAS_FILE = CARPETA_OUT / "cruce_autores" / "cruce_autores_publit_v3.xlsx"
HOJA_INSTANCIAS = "Instancias"
WC_FILE         = CARPETA_OUT / "cruce_publit" / "cruce_publit_v2.csv"

EXCEL_FILE   = CARPETA_DATOS / "Glif_-_Datos_ARS.xlsx"

NODOS_FILE   = CARPETA_SALIDA / "nodos_publit_v4.csv"
PERFIL_FILE  = CARPETA_SALIDA / "perfil_grupos_publit_v4.csv"
CONTROL_FILE = CARPETA_SALIDA / "control_citacion_publit_v4.csv"
CENTRAL_FILE = CARPETA_SALIDA / "centralidad_grupos_publit_v4.csv"
ARISTAS_FILE = CARPETA_SALIDA / "aristas_publit_v4.csv"
GEXF_FILE    = CARPETA_SALIDA / "red_publit_v4.gexf"
REPORTE_FILE = CARPETA_SALIDA / "reporte_red_publit_v4.txt"

# Un articulo con muchos firmantes genera un clique enorme que domina la
# visualizacion sin aportar informacion. Los que superen este umbral se
# excluyen del calculo de aristas y se reportan aparte.
MAX_AUTORES_POR_ARTICULO = 25

CAMPOS_NODOS = ["Id", "Label", "grupo", "zona", "n_articulos", "n_coautores",
                "fuerza", "clustering", "intermediacion", "interm_gigante",
                "en_gigante", "secciones", "con_marca_K", "wc_principal"]
CAMPOS_ARISTAS = ["Source", "Target", "Type", "Weight"]


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def escribir_csv(path, filas, campos):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=campos, extrasaction="ignore")
        w.writeheader()
        w.writerows(filas)


def leer_csv_auto(path):
    if not path.exists():
        return []
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        for sep in (",", ";", "\t"):
            try:
                with open(path, encoding=enc, newline="") as f:
                    r = csv.DictReader(f, delimiter=sep)
                    filas = list(r)
                    cols = list(r.fieldnames or [])
                if cols and len(cols) > 1:
                    return filas
            except Exception:
                continue
    return []


def escapar(txt):
    """Escapa el texto para XML (GEXF)."""
    return (str(txt).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


# ─── CARGA ────────────────────────────────────────────────────────────────────

def cargar_instancias():
    """Lee la hoja 'Instancias': una fila por aparicion de autor."""
    print("\n[1/4] Cargando instancias de autor...")
    if not INSTANCIAS_FILE.exists():
        print(f"  ERROR: no se encontro {INSTANCIAS_FILE}")
        return []
    wb = openpyxl.load_workbook(INSTANCIAS_FILE, read_only=True, data_only=True)
    if HOJA_INSTANCIAS not in wb.sheetnames:
        print(f"  ERROR: el libro no tiene la hoja '{HOJA_INSTANCIAS}'")
        return []
    filas = wb[HOJA_INSTANCIAS].iter_rows(values_only=True)
    cab = [str(c).strip() if c is not None else "" for c in next(filas)]
    idx = {c: i for i, c in enumerate(cab)}
    faltan = [c for c in ("id_ref", "autor_norm", "grupo") if c not in idx]
    if faltan:
        print(f"  ERROR: faltan columnas en la hoja: {faltan}")
        return []

    datos = []
    for r in filas:
        if r[idx["id_ref"]] is None:
            continue
        autor = str(r[idx["autor_norm"]] or "").strip()
        if not autor:
            continue
        datos.append({
            "id_ref": str(r[idx["id_ref"]]),
            "autor": autor,
            "grupo": str(r[idx["grupo"]] or "SIN_INFORMACION"),
            "zona": str(r[idx["zona"]] or "") if "zona" in idx else "",
            "seccion": str(r[idx["seccion"]] or "") if "seccion" in idx else "",
            "evaluacion": str(r[idx["evaluacion"]] or "") if "evaluacion" in idx else "",
        })
    print(f"  {len(datos):,} instancias | "
          f"{len({d['id_ref'] for d in datos}):,} referencias | "
          f"{len({d['autor'] for d in datos}):,} autores")
    return datos


def cargar_areas():
    """Area tematica de WoS por referencia, si esta disponible."""
    filas = leer_csv_auto(WC_FILE)
    if not filas:
        print("  (sin cruce_publit_v2.csv: los nodos quedan sin area tematica)")
        return {}
    areas = {}
    for r in filas:
        wc = (r.get("wc_wos") or "").strip()
        if wc:
            areas[str(r.get("id_cita", ""))] = wc.split(";")[0].strip()
    print(f"  areas tematicas disponibles para {len(areas):,} referencias")
    return areas


def cargar_citacion():
    """Citas por autor desde la base de 2018.

    Devuelve dos diccionarios: citas de toda la obra del autor en la base
    (peso en el campo) y citas por id_wos (para los articulos del
    expediente). Solo cubre a los autores presentes en la base, lo que se
    declara al reportar las medianas.
    """
    print("\n  Cargando citacion desde la base de 2018...")
    if not EXCEL_FILE.exists():
        print(f"  (sin {EXCEL_FILE.name}: no se calcula citacion)")
        return {}, {}
    import unicodedata

    def norm(t):
        t = unicodedata.normalize("NFKD", str(t))
        t = "".join(c for c in t if not unicodedata.combining(c))
        return re.sub(r"\s+", " ", t).strip().upper()

    wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    citas_autor = defaultdict(list)
    citas_art = {}
    filas = wb["Base"].iter_rows(values_only=True)
    next(filas, None)
    for r in filas:
        if r[0] is None:
            continue
        tc = r[32] if isinstance(r[32], (int, float)) else 0
        citas_art[str(r[0])] = tc
        for a in str(r[2] or "").split(";"):
            a = norm(a)
            if a:
                citas_autor[a].append(tc)
    print(f"  citacion disponible para {len(citas_autor):,} autores")
    return citas_autor, citas_art


def mediana(vals):
    v = sorted(vals)
    n = len(v)
    if not n:
        return None
    return v[n // 2] if n % 2 else (v[n // 2 - 1] + v[n // 2]) / 2


def perfil_por_grupo(nodos, grado, comps, datos, citas_autor, citas_art, areas):
    """Tabla 5: comportamiento comparado de los cuatro grupos."""
    gigante = set(comps[0]) if comps else set()

    # citas de los articulos del expediente, por autor
    id_wos_por_ref = {}
    for r in leer_csv_auto(WC_FILE):
        iw = str(r.get("id_wos") or "").strip()
        if iw:
            id_wos_por_ref[str(r.get("id_cita", ""))] = iw
    arts_autor = defaultdict(set)
    for d in datos:
        iw = id_wos_por_ref.get(d["id_ref"])
        if iw:
            arts_autor[d["autor"]].add(iw)

    PRIORIDAD = ["CONTROL", "ZONA_1", "ZONA_2", "SIN_INFORMACION"]
    filas = []
    for g in PRIORIDAD:
        sub = [n for n in nodos if n["grupo"] == g]
        if not sub:
            continue
        con_citas = [n for n in sub if n["Id"] in citas_autor]
        med_obra = mediana([mediana(citas_autor[n["Id"]]) for n in con_citas]) \
            if con_citas else None
        cit_exp = []
        for n in sub:
            vals = [citas_art[i] for i in arts_autor.get(n["Id"], ())
                    if i in citas_art]
            if vals:
                cit_exp.append(mediana(vals))
        filas.append({
            "grupo": g,
            "autores": len(sub),
            "pct_autores": round(100 * len(sub) / len(nodos), 1),
            "arts_expediente": round(sum(n["n_articulos"] for n in sub) / len(sub), 2),
            "coautores": round(sum(n["n_coautores"] for n in sub) / len(sub), 2),
            "n_en_gigante": sum(1 for n in sub if n["Id"] in gigante),
            "pct_en_gigante": round(100 * sum(1 for n in sub if n["Id"] in gigante)
                                    / len(sub), 1),
            "en_base_2018": len(con_citas),
            "mediana_citas_obra": med_obra if med_obra is not None else "",
            "mediana_citas_expediente": mediana(cit_exp) if cit_exp else "",
        })
    return filas


def clustering_local(nodo, vecinos_de):
    """Fraccion de pares de coautores de un nodo que son coautores entre si."""
    v = vecinos_de.get(nodo, set())
    k = len(v)
    if k < 2:
        return 0.0
    enlaces = sum(1 for a, b in combinations(v, 2) if b in vecinos_de.get(a, ()))
    return 2 * enlaces / (k * (k - 1))


def intermediacion(nodos_ids, vecinos_de):
    """Intermediacion de Brandes para grafos no dirigidos y no ponderados.

    Se implementa aca en lugar de usar una libreria externa para que el
    script no dependa de networkx: la red tiene pocos miles de nodos y el
    algoritmo es lineal en nodos por aristas.
    """
    cb = {v: 0.0 for v in nodos_ids}
    for s0 in nodos_ids:
        pila, pred = [], {v: [] for v in nodos_ids}
        sigma = {v: 0 for v in nodos_ids}
        dist = {v: -1 for v in nodos_ids}
        sigma[s0], dist[s0] = 1, 0
        cola = [s0]
        i = 0
        while i < len(cola):
            v = cola[i]
            i += 1
            pila.append(v)
            for w in vecinos_de.get(v, ()):
                if dist[w] < 0:
                    dist[w] = dist[v] + 1
                    cola.append(w)
                if dist[w] == dist[v] + 1:
                    sigma[w] += sigma[v]
                    pred[w].append(v)
        delta = {v: 0.0 for v in nodos_ids}
        while pila:
            w = pila.pop()
            for v in pred[w]:
                delta[v] += (sigma[v] / sigma[w]) * (1 + delta[w])
            if w != s0:
                cb[w] += delta[w]
    # no dirigido: cada par se cuenta dos veces
    return {v: c / 2 for v, c in cb.items()}


def centralidad_por_grupo(nodos):
    """Medidas de centralidad agregadas por grupo de vinculacion.

    Se reportan medianas ademas de medias porque las distribuciones de
    intermediacion son muy asimetricas: unos pocos nodos concentran casi
    todo el valor y la media sola resulta enganosa.
    """
    PRIORIDAD = ["CONTROL", "ZONA_1", "ZONA_2", "SIN_INFORMACION"]
    filas = []
    for g in PRIORIDAD:
        sub = [n for n in nodos if n["grupo"] == g]
        if not sub:
            continue
        gig = [n for n in sub if n["en_gigante"] == 1]
        filas.append({
            "grupo": g,
            "autores": len(sub),
            "grado_medio": round(sum(n["n_coautores"] for n in sub) / len(sub), 2),
            "grado_mediana": mediana([n["n_coautores"] for n in sub]),
            "fuerza_media": round(sum(n["fuerza"] for n in sub) / len(sub), 2),
            "clustering_medio": round(sum(n["clustering"] for n in sub) / len(sub), 3),
            "interm_media": round(sum(n["intermediacion"] for n in sub) / len(sub), 1),
            "interm_mediana": mediana([n["intermediacion"] for n in sub]),
            "pct_interm_no_nula": round(
                100 * sum(1 for n in sub if n["intermediacion"] > 0) / len(sub), 1),
            "n_en_gigante": len(gig),
            "interm_media_gigante": (
                round(sum(n["interm_gigante"] for n in gig) / len(gig), 1)
                if gig else ""),
        })
    return filas


def percentil_en(valor, distribucion):
    """Porcentaje de la distribucion que queda por debajo del valor."""
    if not distribucion:
        return ""
    return round(100 * sum(1 for v in distribucion if v < valor)
                 / len(distribucion), 1)


def detalle_control(nodos, citas_autor, citas_art, datos):
    """Ubica a cada autor del grupo de control en la distribucion general.

    Con seis casos no corresponde calcular una mediana de grupo. Lo que si
    es interpretable es donde cae cada autor dentro de la distribucion de
    todos los autores del corpus, y como se compara con las medianas de
    Zona 1 y Zona 2.
    """
    id_wos_por_ref = {}
    for r in leer_csv_auto(WC_FILE):
        iw = str(r.get("id_wos") or "").strip()
        if iw:
            id_wos_por_ref[str(r.get("id_cita", ""))] = iw
    arts_autor = defaultdict(set)
    for d in datos:
        iw = id_wos_por_ref.get(d["id_ref"])
        if iw:
            arts_autor[d["autor"]].add(iw)

    def cit_obra(a):
        return mediana(citas_autor[a]) if a in citas_autor else None

    def cit_exp(a):
        v = [citas_art[i] for i in arts_autor.get(a, ()) if i in citas_art]
        return mediana(v) if v else None

    # distribuciones de referencia sobre todos los autores con citacion
    todos = [cit_obra(n["Id"]) for n in nodos if cit_obra(n["Id"]) is not None]
    por_grupo = {}
    for g in ("ZONA_1", "ZONA_2", "SIN_INFORMACION"):
        vals = [cit_obra(n["Id"]) for n in nodos
                if n["grupo"] == g and cit_obra(n["Id"]) is not None]
        por_grupo[g] = mediana(vals) if vals else None
    todos_exp = [cit_exp(n["Id"]) for n in nodos if cit_exp(n["Id"]) is not None]

    filas = []
    for n in sorted([x for x in nodos if x["grupo"] == "CONTROL"],
                    key=lambda x: -x["n_articulos"]):
        co, ce = cit_obra(n["Id"]), cit_exp(n["Id"])
        filas.append({
            "autor": n["Id"],
            "arts_expediente": n["n_articulos"],
            "coautores": n["n_coautores"],
            "citas_obra": co if co is not None else "",
            "percentil_obra": percentil_en(co, todos) if co is not None else "",
            "citas_expediente": ce if ce is not None else "",
            "percentil_expediente": percentil_en(ce, todos_exp) if ce is not None else "",
        })
    referencias = {
        "n_autores_con_citacion": len(todos),
        "mediana_general": mediana(todos),
        "mediana_zona1": por_grupo.get("ZONA_1"),
        "mediana_zona2": por_grupo.get("ZONA_2"),
        "mediana_general_expediente": mediana(todos_exp) if todos_exp else None,
    }
    return filas, referencias


# ─── CONSTRUCCIÓN DE LA RED ───────────────────────────────────────────────────

def construir(datos, areas):
    print("\n[2/4] Construyendo la red...")
    por_ref = defaultdict(list)
    for d in datos:
        por_ref[d["id_ref"]].append(d)

    aristas = defaultdict(int)
    grado = defaultdict(set)
    fuerza = defaultdict(int)
    excluidos = []

    for ref, filas in por_ref.items():
        autores = sorted({f["autor"] for f in filas})
        if len(autores) > MAX_AUTORES_POR_ARTICULO:
            excluidos.append((ref, len(autores)))
            continue
        for a, b in combinations(autores, 2):
            aristas[(a, b) if a < b else (b, a)] += 1
            grado[a].add(b)
            grado[b].add(a)
            fuerza[a] += 1
            fuerza[b] += 1

    # atributos de nodo
    info = defaultdict(lambda: {"refs": set(), "grupos": set(), "zonas": set(),
                                "secs": set(), "k": 0, "wc": set()})
    for d in datos:
        n = info[d["autor"]]
        n["refs"].add(d["id_ref"])
        n["grupos"].add(d["grupo"])
        if d["zona"]:
            n["zonas"].add(d["zona"])
        if d["seccion"]:
            n["secs"].add(d["seccion"])
        if "K" in d["evaluacion"]:
            n["k"] += 1
        w = areas.get(d["id_ref"])
        if w:
            n["wc"].add(w)

    # Un autor puede aparecer con distinto grupo segun el articulo; se le
    # asigna el mas informativo segun la misma jerarquia del cruce.
    PRIORIDAD = ["CONTROL", "ZONA_1", "ZONA_2", "SIN_INFORMACION"]
    nodos = []
    for autor, n in info.items():
        grupo = next((g for g in PRIORIDAD if g in n["grupos"]), "SIN_INFORMACION")
        wc = sorted(n["wc"])
        nodos.append({
            "Id": autor, "Label": autor,
            "grupo": grupo,
            "zona": sorted(n["zonas"])[0] if n["zonas"] else "sin clasificar",
            "n_articulos": len(n["refs"]),
            "n_coautores": len(grado.get(autor, ())),
            "fuerza": fuerza.get(autor, 0),
            "clustering": "", "intermediacion": "",
            "interm_gigante": "", "en_gigante": "",
            "secciones": "; ".join(sorted(n["secs"])),
            "con_marca_K": n["k"],
            "wc_principal": wc[0] if wc else "",
        })
    nodos.sort(key=lambda x: -x["n_articulos"])

    filas_aristas = [{"Source": a, "Target": b, "Type": "Undirected", "Weight": w}
                     for (a, b), w in sorted(aristas.items(), key=lambda x: -x[1])]

    print(f"  {len(nodos):,} nodos | {len(filas_aristas):,} aristas")
    if excluidos:
        print(f"  {len(excluidos)} referencias excluidas por superar "
              f"{MAX_AUTORES_POR_ARTICULO} autores")
    return nodos, filas_aristas, grado, excluidos


def componentes(nodos, grado):
    """Componentes conexos por recorrido en anchura."""
    vistos, comps = set(), []
    for n in nodos:
        a = n["Id"]
        if a in vistos:
            continue
        pila, comp = [a], []
        vistos.add(a)
        while pila:
            x = pila.pop()
            comp.append(x)
            for y in grado.get(x, ()):
                if y not in vistos:
                    vistos.add(y)
                    pila.append(y)
        comps.append(comp)
    comps.sort(key=len, reverse=True)
    return comps


# ─── SALIDA GEXF ──────────────────────────────────────────────────────────────

def escribir_gexf(nodos, aristas, path):
    """GEXF con atributos, para abrir directo en Gephi sin importar dos CSV."""
    attrs = [("grupo", "string"), ("zona", "string"),
             ("n_articulos", "integer"), ("n_coautores", "integer"),
             ("fuerza", "integer"), ("clustering", "float"),
             ("intermediacion", "float"), ("en_gigante", "integer"),
             ("secciones", "string"), ("con_marca_K", "integer"),
             ("wc_principal", "string")]
    ids = {n["Id"]: str(i) for i, n in enumerate(nodos)}
    L = ['<?xml version="1.0" encoding="UTF-8"?>',
         '<gexf xmlns="http://www.gexf.net/1.2draft" version="1.2">',
         f'  <meta lastmodifieddate="{datetime.now():%Y-%m-%d}">',
         '    <creator>red_coautoria_publit_v4.py</creator>',
         '    <description>Red de coautoria del expediente europeo 2012</description>',
         '  </meta>',
         '  <graph mode="static" defaultedgetype="undirected">',
         '    <attributes class="node">']
    for i, (nombre, tipo) in enumerate(attrs):
        L.append(f'      <attribute id="{i}" title="{nombre}" type="{tipo}"/>')
    L.append('    </attributes>')
    L.append('    <nodes>')
    for n in nodos:
        L.append(f'      <node id="{ids[n["Id"]]}" label="{escapar(n["Label"])}">')
        L.append('        <attvalues>')
        for i, (nombre, _) in enumerate(attrs):
            L.append(f'          <attvalue for="{i}" value="{escapar(n[nombre])}"/>')
        L.append('        </attvalues>')
        L.append('      </node>')
    L.append('    </nodes>')
    L.append('    <edges>')
    for i, e in enumerate(aristas):
        L.append(f'      <edge id="{i}" source="{ids[e["Source"]]}" '
                 f'target="{ids[e["Target"]]}" weight="{e["Weight"]}"/>')
    L.append('    </edges>')
    L.append('  </graph>')
    L.append('</gexf>')
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(L), encoding="utf-8")


# ─── INFORME ──────────────────────────────────────────────────────────────────

def escribir_informe(nodos, aristas, comps, excluidos, perfil, centr, ctrl, refs, path):
    n, m = len(nodos), len(aristas)
    pct = lambda k, d: f"{(100*k/d):.1f}%" if d else "-"
    dens = (2 * m / (n * (n - 1))) if n > 1 else 0
    grados = [x["n_coautores"] for x in nodos]
    gigante = comps[0] if comps else []
    en_gigante = set(gigante)

    por_grupo = defaultdict(lambda: {"n": 0, "grado": 0, "arts": 0, "gig": 0})
    for x in nodos:
        g = por_grupo[x["grupo"]]
        g["n"] += 1
        g["grado"] += x["n_coautores"]
        g["arts"] += x["n_articulos"]
        g["gig"] += 1 if x["Id"] in en_gigante else 0

    L = ["=" * 64,
         "  [RED DE COAUTORIA] - red_coautoria_publit_v4",
         f"  Fecha: {datetime.now():%Y-%m-%d %H:%M}",
         "=" * 64, "",
         f"  Nodos (autores)        : {n:,}",
         f"  Aristas (coautorias)   : {m:,}",
         f"  Densidad               : {dens:.5f}",
         f"  Grado medio            : {sum(grados)/n if n else 0:.2f}",
         f"  Grado maximo           : {max(grados) if grados else 0}",
         f"  Nodos aislados         : {sum(1 for g in grados if g == 0):,}",
         "",
         f"  Componentes            : {len(comps):,}",
         f"  Componente gigante     : {len(gigante):,} nodos ({pct(len(gigante), n)})",
         "",
         "  Por grupo de vinculacion:",
         f"    {'grupo':<18}{'nodos':>7}{'%':>8}{'grado medio':>13}"
         f"{'arts medios':>13}{'% en gigante':>14}"]
    for g in ["CONTROL", "ZONA_1", "ZONA_2", "SIN_INFORMACION"]:
        d = por_grupo.get(g)
        if not d or not d["n"]:
            continue
        L.append(f"    {g:<18}{d['n']:>7,}{pct(d['n'], n):>8}"
                 f"{d['grado']/d['n']:>13.2f}{d['arts']/d['n']:>13.2f}"
                 f"{pct(d['gig'], d['n']):>14}")

    if perfil:
        L += ["", "  PERFIL COMPARADO DE LOS CUATRO GRUPOS (Tabla 5)",
              f"    {'grupo':<17}{'autores':>8}{'arts':>7}{'coaut':>7}"
              f"{'gigante':>9}{'cit.obra':>10}{'cit.exped':>11}"]
        for f in perfil:
            L.append(f"    {f['grupo']:<17}{f['autores']:>8,}"
                     f"{f['arts_expediente']:>7.2f}{f['coautores']:>7.2f}"
                     f"{f['pct_en_gigante']:>8.1f}%"
                     f"{str(f['mediana_citas_obra']):>10}"
                     f"{str(f['mediana_citas_expediente']):>11}")
        L += ["",
              "    arts      : articulos del expediente por autor (media)",
              "    coaut     : coautores distintos por autor (media)",
              "    gigante   : porcentaje del grupo en el componente principal",
              "    cit.obra  : mediana de la citacion mediana de toda la obra",
              "                del autor en la base de 2018",
              "    cit.exped : idem, restringido a sus articulos del expediente",
              "",
              "    Autores con citacion disponible (presentes en la base 2018):"]
        for f in perfil:
            L.append(f"      {f['grupo']:<18}{f['en_base_2018']:>6} de "
                     f"{f['autores']:,}")

    if centr:
        L += ["", "  CENTRALIDAD POR GRUPO",
              f"    {'grupo':<17}{'grado':>7}{'fuerza':>8}{'clust':>8}"
              f"{'interm':>10}{'%int>0':>8}{'int.gig':>9}"]
        for f in centr:
            L.append(f"    {f['grupo']:<17}{f['grado_medio']:>7.2f}"
                     f"{f['fuerza_media']:>8.2f}{f['clustering_medio']:>8.3f}"
                     f"{f['interm_media']:>10.1f}{f['pct_interm_no_nula']:>7.1f}%"
                     f"{str(f['interm_media_gigante']):>9}")
        L += ["",
              "    grado   : coautores distintos (media)",
              "    fuerza  : grado ponderado por trabajos en comun (media)",
              "    clust   : coeficiente de clustering local (media)",
              "    interm  : intermediacion sobre el grafo completo (media)",
              "    %int>0  : porcentaje del grupo con intermediacion no nula",
              "    int.gig : intermediacion media dentro del componente principal"]

    if ctrl:
        L += ["", "  GRUPO DE CONTROL, CASO POR CASO",
              f"    (con seis autores no corresponde una mediana de grupo;",
              f"     se reporta la posicion de cada uno en la distribucion)",
              "",
              f"    {'autor':<20}{'arts':>6}{'coaut':>7}{'cit.obra':>10}"
              f"{'pctil':>8}{'cit.exp':>9}{'pctil':>8}"]
        for f in ctrl:
            L.append(f"    {f['autor'][:18]:<20}{f['arts_expediente']:>6}"
                     f"{f['coautores']:>7}{str(f['citas_obra']):>10}"
                     f"{str(f['percentil_obra']):>8}"
                     f"{str(f['citas_expediente']):>9}"
                     f"{str(f['percentil_expediente']):>8}")
        L += ["", "    Valores de referencia (citacion de la obra completa):",
              f"      autores con citacion disponible : "
              f"{refs['n_autores_con_citacion']:,}",
              f"      mediana general                 : {refs['mediana_general']}",
              f"      mediana Zona 1                  : {refs['mediana_zona1']}",
              f"      mediana Zona 2                  : {refs['mediana_zona2']}",
              f"      mediana general (expediente)    : "
              f"{refs['mediana_general_expediente']}"]

    if excluidos:
        L += ["", f"  Referencias excluidas por exceso de autores: {len(excluidos)}"]
        for ref, k in sorted(excluidos, key=lambda x: -x[1])[:5]:
            L.append(f"    {ref}: {k} autores")

    L += ["", "  Sugerencia para Gephi:",
          "    abrir red_publit_v4.gexf, aplicar ForceAtlas2 con 'Prevent",
          "    Overlap', colorear los nodos por el atributo 'grupo' y",
          "    dimensionarlos por 'n_articulos'. Filtrar por componente",
          "    gigante si la periferia de diadas dificulta la lectura.",
          "",
          "  Archivos generados:",
          f"    {GEXF_FILE.name}   <- abrir este en Gephi",
          f"    {PERFIL_FILE.name}   <- Tabla 5 del articulo",
          f"    {CONTROL_FILE.name}",
          f"    {CENTRAL_FILE.name}",
          f"    {NODOS_FILE.name}",
          f"    {ARISTAS_FILE.name}",
          f"    {REPORTE_FILE.name}", ""]

    rep = "\n".join(L)
    print("\n" + rep)
    path.write_text(rep + "\n", encoding="utf-8")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 64)
    print("  [RED DE COAUTORIA] - red_coautoria_publit_v4")
    print("=" * 64)
    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    print(f"  output : {CARPETA_SALIDA}")

    datos = cargar_instancias()
    if not datos:
        return
    areas = cargar_areas()

    citas_autor, citas_art = cargar_citacion()

    nodos, aristas, grado, excluidos = construir(datos, areas)
    if not nodos:
        print("ERROR: la red quedo vacia.")
        return
    comps = componentes(nodos, grado)

    print("\n  Calculando centralidades...")
    ids = [n["Id"] for n in nodos]
    cb = intermediacion(ids, grado)
    gig = set(comps[0]) if comps else set()
    sub = {v: {w for w in grado.get(v, ()) if w in gig} for v in gig}
    cb_gig = intermediacion(list(gig), sub) if gig else {}
    for n in nodos:
        a = n["Id"]
        n["clustering"] = round(clustering_local(a, grado), 4)
        n["intermediacion"] = round(cb.get(a, 0.0), 2)
        n["interm_gigante"] = round(cb_gig.get(a, 0.0), 2) if a in gig else ""
        n["en_gigante"] = 1 if a in gig else 0
    print(f"    intermediacion sobre {len(ids):,} nodos y "
          f"{len(gig):,} en el componente principal")
    perfil = perfil_por_grupo(nodos, grado, comps, datos, citas_autor,
                              citas_art, areas)
    ctrl, refs = detalle_control(nodos, citas_autor, citas_art, datos)
    centr = centralidad_por_grupo(nodos)

    print("\n[3/4] Escribiendo nodos, aristas y GEXF...")
    escribir_csv(NODOS_FILE, nodos, CAMPOS_NODOS)
    escribir_csv(ARISTAS_FILE, aristas, CAMPOS_ARISTAS)
    escribir_gexf(nodos, aristas, GEXF_FILE)
    if perfil:
        escribir_csv(PERFIL_FILE, perfil, list(perfil[0].keys()))
    if ctrl:
        escribir_csv(CONTROL_FILE, ctrl, list(ctrl[0].keys()))
    if centr:
        escribir_csv(CENTRAL_FILE, centr, list(centr[0].keys()))

    print("\n[4/4] Calculando metricas...")
    escribir_informe(nodos, aristas, comps, excluidos, perfil, centr, ctrl,
                     refs, REPORTE_FILE)
    print(f"Archivos en {CARPETA_SALIDA}")


if __name__ == "__main__":
    main()
