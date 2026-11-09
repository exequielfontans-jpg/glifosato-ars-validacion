"""
extraer_corpus_gtf_v2.py
========================
Cosecha desde el Wayback Machine la base de literatura del Glyphosate Task
Force (glifosato.es), parsea las citas y las reconcilia contra Crossref
para obtener DOIs y metadatos.

Estructura de carpetas del proyecto:
  proyecto/
    script/     los .py
    data/       fuentes originales, NUNCA se escribe aqui
    output/     todo lo que generan los scripts, una subcarpeta por script

Si se borra output/ y se vuelven a correr los scripts en orden, el resultado
debe ser identico. Por eso las salidas no van en data/.

Proyecto: Corpus curados sobre glifosato — comparacion antologia Rossi /
          Glyphosate Task Force. FIC-UDELAR.
Contexto: Construccion del corpus "industria" · cosecha Wayback + Crossref

Entradas:
  - (ninguna; el script consulta las APIs de Wayback y Crossref)

Salidas:
  - output/gtf/urls_gtf_v1.csv          - snapshots localizados en el Wayback
  - output/gtf/html/*.html              - HTML crudo de cada pagina
  - output/gtf/citas_gtf_v2.csv         - citas parseadas (autores/titulo/fuente)
  - output/gtf/corpus_gtf_v2.csv        - citas + DOI y metadatos de Crossref
  - output/gtf/categorias_gtf_v2.csv    - tabla larga cita x categoria (bipartita)
  - output/gtf/reporte_extraer_corpus_gtf.txt - informe estadistico
  - output/gtf/crossref_checkpoint.csv  - checkpoint (borrar tras verificar)
  - output/gtf/extraer_corpus_gtf_errores.log - log de errores

Sensibilidad de salidas:
  NO SENSIBLE : todas (metadatos bibliograficos publicos)

Uso: python extraer_corpus_gtf_v2.py
     IMPORTANTE: borrar crossref_checkpoint.csv de la corrida v1 antes de
     ejecutar; la logica de validacion de DOI cambio y el checkpoint viejo
     arrastraria los falsos positivos.

Requisitos: pip install requests beautifulsoup4

Historial:
  v1 - version inicial: CDX + descarga snapshots + parseo citas + Crossref
  v2 - corrige tres defectos de la v1:
       (a) parseo: separa multiples citas dentro de un mismo bloque HTML
           (listados separados por <br> en lugar de un <p> por cita);
       (b) deduplicacion: ya no borra la pertenencia multiple a categorias,
           se conserva la relacion cita x categoria en tabla aparte;
       (c) validacion de DOI: exige similitud de titulo y coherencia de anio
           ademas del score de Crossref; sin eso el DOI queda vacio.
"""

import csv
import re
import json
import time
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path
from datetime import datetime

import requests
from bs4 import BeautifulSoup

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

CARPETA_SCRIPT = Path(__file__).resolve().parent
CARPETA_BASE   = CARPETA_SCRIPT.parent
CARPETA_DATOS  = CARPETA_BASE / "data"           # fuentes originales, solo lectura
CARPETA_OUT    = CARPETA_BASE / "output"         # todo lo generado por scripts
CARPETA_SALIDA = CARPETA_OUT / "gtf"
CARPETA_HTML   = CARPETA_SALIDA / "html"

URLS_FILE       = CARPETA_SALIDA / "urls_gtf_v1.csv"
CITAS_FILE      = CARPETA_SALIDA / "citas_gtf_v2.csv"
OUTPUT_FILE     = CARPETA_SALIDA / "corpus_gtf_v2.csv"
CATEGORIAS_FILE = CARPETA_SALIDA / "categorias_gtf_v2.csv"
REPORTE_FILE    = CARPETA_SALIDA / "reporte_extraer_corpus_gtf.txt"
CHECKPOINT_FILE = CARPETA_SALIDA / "crossref_checkpoint.csv"
LOG_FILE        = CARPETA_SALIDA / "extraer_corpus_gtf_errores.log"

# Dominio a cosechar y filtro de rutas de interes
DOMINIO       = "glifosato.es"
PATRON_RUTA   = "base-de-datos-de-literatura"
SNAPSHOT_HASTA = "2018"        # preferir snapshots anteriores a esta fecha

# Crossref (el mailto habilita el "polite pool", mas rapido y estable)
CROSSREF_MAILTO = "exequiel.fontans@fic.edu.uy"
SCORE_MINIMO    = 60.0         # score minimo de Crossref
SIMILITUD_MINIMA = 0.85        # similitud de titulo exigida para aceptar el DOI
DELTA_ANIO_MAX   = 1           # diferencia maxima admitida entre anio citado y Crossref
CROSSREF_ROWS    = 3           # candidatos a evaluar por cita

PAUSA_WAYBACK  = 1.5
PAUSA_CROSSREF = 0.4
TIMEOUT        = 60
REINTENTOS     = 3

CAMPOS_CITAS  = ["id_cita", "categorias", "n_categorias", "url_origen",
                 "cita_cruda", "autores", "titulo", "fuente", "anio"]
CAMPOS_SALIDA = CAMPOS_CITAS + ["doi", "titulo_crossref", "revista", "anio_crossref",
                                "n_autores", "score", "similitud", "revision_manual"]


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def log_error(mensaje):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{ts}] {mensaje}\n")


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
                    cols  = list(r.fieldnames or [])
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


def get_con_reintentos(url, params=None, headers=None):
    """GET con reintentos y backoff. Devuelve la respuesta o None."""
    espera = 2
    for intento in range(1, REINTENTOS + 1):
        try:
            r = requests.get(url, params=params, headers=headers, timeout=TIMEOUT)
            if r.status_code == 200:
                return r
            log_error(f"HTTP {r.status_code} en {url} (intento {intento}) "
                      f"| body: {r.text[:300]}")
        except Exception as e:
            log_error(f"EXCEPCION en {url} (intento {intento}): {repr(e)[:300]}")
        if intento < REINTENTOS:
            time.sleep(espera)
            espera *= 2
    return None


def normalizar(txt):
    """Normaliza espacios, comillas tipograficas y guiones."""
    if not txt:
        return ""
    txt = unicodedata.normalize("NFC", txt)
    for a, b in [("\u201c", '"'), ("\u201d", '"'), ("\u2018", "'"),
                 ("\u2019", "'"), ("\u2013", "-"), ("\u2014", "-"),
                 ("\u00a0", " ")]:
        txt = txt.replace(a, b)
    return re.sub(r"\s+", " ", txt).strip()


def slug(url):
    s = re.sub(r"^https?://", "", url)
    s = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")
    return s[:120]


# ─── CARGA: enumerar snapshots en el Wayback ──────────────────────────────────

def enumerar_snapshots():
    """Consulta la API CDX y devuelve un snapshot por URL de interes."""
    print("\n[1/4] Enumerando snapshots en el Wayback Machine...")
    params = {
        "url": DOMINIO,
        "matchType": "domain",
        "output": "json",
        "fl": "original,timestamp,statuscode",
        "filter": "statuscode:200",
        "collapse": "urlkey",
    }
    r = get_con_reintentos("http://web.archive.org/cdx/search/cdx", params=params)
    if not r:
        print("  ERROR: no se pudo consultar la API CDX.")
        return []

    try:
        datos = r.json()
    except Exception as e:
        log_error(f"CDX devolvio contenido no-JSON: {repr(e)[:300]}")
        return []

    if not datos or len(datos) < 2:
        print("  No se encontraron snapshots.")
        return []

    filas = [dict(zip(datos[0], fila)) for fila in datos[1:]]

    # Filtrar a las paginas de la base de literatura y quedarse con el
    # snapshot mas antiguo de cada URL (mas cercano al corpus WoS de 2016).
    mejores = {}
    for f in filas:
        url = f.get("original", "")
        if PATRON_RUTA not in url.lower():
            continue
        ts = f.get("timestamp", "")
        prev = mejores.get(url)
        if prev is None:
            mejores[url] = ts
            continue
        # preferir snapshots anteriores a SNAPSHOT_HASTA; entre ellos, el mayor
        prev_ok = prev < SNAPSHOT_HASTA
        ts_ok   = ts < SNAPSHOT_HASTA
        if ts_ok and (not prev_ok or ts > prev):
            mejores[url] = ts
        elif not prev_ok and not ts_ok and ts < prev:
            mejores[url] = ts

    salida = []
    for url, ts in sorted(mejores.items()):
        categoria = url.rstrip("/").split("/")[-1]
        categoria = categoria.replace("base-de-datos-de-literatura-especializada", "")
        categoria = categoria.strip("-").replace("-", " ") or "general"
        salida.append({"url": url, "timestamp": ts, "categoria": categoria})

    print(f"  {len(salida)} paginas de la base localizadas.")
    return salida


# ─── PROCESAMIENTO: descargar y parsear ───────────────────────────────────────

def descargar_snapshots(snapshots):
    """Descarga el HTML crudo de cada snapshot (sufijo id_ = sin toolbar)."""
    print("\n[2/4] Descargando snapshots...")
    CARPETA_HTML.mkdir(parents=True, exist_ok=True)
    for i, s in enumerate(snapshots, 1):
        destino = CARPETA_HTML / f"{slug(s['url'])}.html"
        if destino.exists():
            continue
        url = f"https://web.archive.org/web/{s['timestamp']}id_/{s['url']}"
        r = get_con_reintentos(url)
        if r is None:
            print(f"  [{i}/{len(snapshots)}] FALLO {s['categoria']}")
            continue
        destino.write_text(r.text, encoding="utf-8", errors="replace")
        print(f"  [{i}/{len(snapshots)}] {s['categoria']}")
        time.sleep(PAUSA_WAYBACK)


RE_TITULO = re.compile(r'"([^"]{15,400})"')
RE_ANIO   = re.compile(r"\b(19[3-9]\d|20[0-2]\d)\b")


def clave_titulo(titulo):
    """Clave normalizada para deduplicar y comparar titulos."""
    t = unicodedata.normalize("NFKD", normalizar(titulo).lower())
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9 ]+", " ", t).strip()


def similitud(a, b):
    return SequenceMatcher(None, clave_titulo(a), clave_titulo(b)).ratio()


def armar_cita(autores, titulo, resto):
    """Valida y arma una cita a partir de sus tres tramos."""
    autores = normalizar(autores).strip(" .;,")
    titulo  = normalizar(titulo).strip().rstrip(".")
    resto   = re.sub(r"^\s*\[In [^\]]+\]\.?\s*", "", normalizar(resto)).strip()
    if len(titulo) < 15 or len(autores) < 4:
        return None
    anios = RE_ANIO.findall(resto) or RE_ANIO.findall(autores)
    return {"autores": autores, "titulo": titulo,
            "fuente": resto, "anio": anios[-1] if anios else ""}


def parsear_bloque(texto):
    """Extrae TODAS las citas contenidas en un bloque de texto.

    La v1 asumia una cita por bloque HTML y devolvia solo la primera. Las
    paginas del GTF listan varias citas dentro de un mismo bloque (separadas
    por <br>), asi que hay que segmentar por titulo entrecomillado: cada par
    de comillas abre una cita nueva y lo que queda entre dos titulos se
    reparte entre la fuente de la anterior y los autores de la siguiente.
    """
    t = normalizar(texto)
    marcas = list(RE_TITULO.finditer(t))
    if not marcas:
        return []
    if len(marcas) == 1:
        m = marcas[0]
        c = armar_cita(t[:m.start()], m.group(1), t[m.end():])
        return [c] if c else []

    citas = []
    for i, m in enumerate(marcas):
        ini_aut = marcas[i - 1].end() if i else 0
        fin_res = marcas[i + 1].start() if i + 1 < len(marcas) else len(t)
        entre   = t[m.end():fin_res]
        tramo_autores = t[ini_aut:m.start()]

        if i:
            # El tramo previo trae la cola de la cita anterior seguida de los
            # autores de esta. El corte esta donde termina la paginacion o el
            # anio de la anterior: un digito + punto/parentesis. No sirve
            # cortar por "punto + mayuscula" porque las iniciales de los
            # autores (B. O. Tasie, R. Kale) disparan ese patron y dejarian
            # solo el ultimo apellido.
            corte = 0
            for mm in re.finditer(r"\d\s*[.\)]\s+", tramo_autores):
                corte = mm.end()
            if not corte:
                for mm in re.finditer(r"[.\)]\s+(?=[A-Z\u00c0-\u00dc])", tramo_autores):
                    corte = mm.end()
            tramo_autores = tramo_autores[corte:]

        if i + 1 < len(marcas):
            corte = len(entre)
            for mm in re.finditer(r"[.\)]\s+(?=[A-Z\u00c0-\u00dc])", entre):
                corte = mm.end()
                break
            entre = entre[:corte]

        c = armar_cita(tramo_autores, m.group(1), entre)
        if c:
            citas.append(c)
    return citas


def parsear_paginas(snapshots):
    """Extrae las citas de cada pagina y conserva la relacion cita x categoria."""
    print("\n[3/4] Parseando citas...")
    registro, n = {}, 0

    for s in snapshots:
        path = CARPETA_HTML / f"{slug(s['url'])}.html"
        if not path.exists():
            print(f"  {s['categoria']}: SIN HTML (fallo la descarga)")
            continue

        soup = BeautifulSoup(path.read_text(encoding="utf-8", errors="replace"),
                             "html.parser")
        for tag in soup(["script", "style", "nav", "header", "footer"]):
            tag.decompose()
        for br in soup.find_all("br"):
            br.replace_with("\n")

        bloques = [b for b in soup.find_all(["p", "li", "div", "td"])
                   if not b.find(["p", "li", "div", "td"])]

        nuevas, repetidas = 0, 0
        vistos_pagina = set()
        for b in bloques:
            texto = b.get_text(" ", strip=True)
            for c in parsear_bloque(texto):
                k = clave_titulo(c["titulo"])[:120]
                if not k or k in vistos_pagina:
                    continue
                vistos_pagina.add(k)
                if k in registro:
                    # mismo articulo listado bajo otra categoria: se acumula
                    registro[k]["categorias"].append(s["categoria"])
                    repetidas += 1
                    continue
                n += 1
                registro[k] = {
                    "id_cita": f"GTF{n:05d}",
                    "categorias": [s["categoria"]],
                    "url_origen": s["url"],
                    "cita_cruda": texto[:1000],
                    **c,
                }
                nuevas += 1
        print(f"  {s['categoria']}: {nuevas} nuevas, {repetidas} ya vistas")

    citas = []
    for r in registro.values():
        cats = sorted(set(r["categorias"]))
        citas.append({**r, "categorias": "; ".join(cats), "n_categorias": len(cats)})

    print(f"  Total: {len(citas)} citas unicas.")
    return citas


def tabla_categorias(citas):
    """Tabla larga cita x categoria, lista para armar la red bipartita."""
    filas = []
    for c in citas:
        for cat in c["categorias"].split("; "):
            filas.append({"id_cita": c["id_cita"], "categoria": cat,
                          "titulo": c["titulo"], "anio": c["anio"]})
    return filas


# ─── PROCESAMIENTO: reconciliar con Crossref ──────────────────────────────────

def consultar_crossref(cita):
    """Busca el DOI en Crossref y lo acepta solo si el match es verificable.

    La v1 tomaba el primer resultado sin control, lo que producia 100% de
    cobertura aparente con falsos positivos (articulos posteriores a la fecha
    del snapshot, por ejemplo). Aca se piden varios candidatos y se exige
    similitud de titulo y coherencia de anio ademas del score.
    """
    params = {
        "query.bibliographic": f"{cita['titulo']} {cita['anio']}".strip(),
        "rows": CROSSREF_ROWS,
        "select": "DOI,title,container-title,issued,author,score",
        "mailto": CROSSREF_MAILTO,
    }
    r = get_con_reintentos("https://api.crossref.org/works", params=params)
    if r is None:
        return None
    try:
        items = r.json().get("message", {}).get("items", [])
    except Exception as e:
        log_error(f"Crossref no-JSON para {cita['id_cita']}: {repr(e)[:300]}")
        return None

    vacio = {"doi": "", "titulo_crossref": "", "revista": "", "anio_crossref": "",
             "n_autores": "", "score": 0.0, "similitud": 0.0,
             "revision_manual": "SIN_RESULTADO"}
    if not items:
        return vacio

    # elegir el candidato mas parecido al titulo citado, no el primero
    mejor, mejor_sim = None, -1.0
    for it in items:
        titulo_cr = (it.get("title") or [""])[0]
        sim = similitud(cita["titulo"], titulo_cr) if titulo_cr else 0.0
        if sim > mejor_sim:
            mejor, mejor_sim = it, sim

    fecha = mejor.get("issued", {}).get("date-parts", [[""]])[0]
    anio_cr = str(fecha[0]) if fecha and fecha[0] else ""
    score   = float(mejor.get("score", 0) or 0)

    motivos = []
    if mejor_sim < SIMILITUD_MINIMA:
        motivos.append("TITULO_DISTINTO")
    if score < SCORE_MINIMO:
        motivos.append("SCORE_BAJO")
    if cita["anio"] and anio_cr:
        if abs(int(cita["anio"]) - int(anio_cr)) > DELTA_ANIO_MAX:
            motivos.append("ANIO_DISCREPA")
    elif not anio_cr:
        motivos.append("SIN_ANIO")

    aceptado = not motivos
    return {
        "doi": mejor.get("DOI", "") if aceptado else "",
        "titulo_crossref": (mejor.get("title") or [""])[0],
        "revista": (mejor.get("container-title") or [""])[0],
        "anio_crossref": anio_cr,
        "n_autores": len(mejor.get("author", []) or []),
        "score": round(score, 1),
        "similitud": round(mejor_sim, 3),
        "revision_manual": "" if aceptado else "+".join(motivos),
    }


def reconciliar(citas):
    """Consulta Crossref para cada cita, con checkpoint."""
    print("\n[4/4] Reconciliando con Crossref...")
    previos, _ = leer_csv_auto(CHECKPOINT_FILE)
    hechos = {r["id_cita"]: r for r in previos if r.get("id_cita")}
    if hechos:
        print(f"  Checkpoint: {len(hechos):,} ya procesadas.")

    escribir_header = not CHECKPOINT_FILE.exists()
    resultado = []

    for i, c in enumerate(citas, 1):
        if c["id_cita"] in hechos:
            resultado.append(hechos[c["id_cita"]])
            continue

        datos = consultar_crossref(c)
        if datos is None:                 # error real: no guardar, se reintenta
            print(f"  [{i}/{len(citas)}] {c['id_cita']} FALLO - se reintentara")
            continue

        fila = {**c, **datos}
        resultado.append(fila)
        with open(CHECKPOINT_FILE, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=CAMPOS_SALIDA, extrasaction="ignore")
            if escribir_header:
                w.writeheader()
                escribir_header = False
            w.writerow(fila)

        if i % 25 == 0:
            print(f"  [{i}/{len(citas)}] procesadas")
        time.sleep(PAUSA_CROSSREF)

    return resultado


# ─── INFORME ──────────────────────────────────────────────────────────────────

def escribir_informe(snapshots, citas, corpus, path):
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    total = len(corpus)
    pct = lambda n: f"{(100 * n / total):.1f}%" if total else "-"

    bajadas = sum(1 for s in snapshots
                  if (CARPETA_HTML / f"{slug(s['url'])}.html").exists())
    con_doi = sum(1 for r in corpus if r.get("doi"))

    motivos = {}
    for r in corpus:
        m = r.get("revision_manual") or ""
        if m:
            motivos[m] = motivos.get(m, 0) + 1

    por_cat = {}
    for r in corpus:
        for cat in str(r.get("categorias", "")).split("; "):
            if cat:
                por_cat[cat] = por_cat.get(cat, 0) + 1

    multi = sum(1 for r in corpus if int(r.get("n_categorias", 1) or 1) > 1)
    anios = [int(r["anio_crossref"]) for r in corpus
             if str(r.get("anio_crossref", "")).isdigit() and r.get("doi")]

    lineas = [
        "=" * 64,
        "  [CORPUS GTF] - extraer_corpus_gtf_v2",
        f"  Fecha: {ahora}",
        "=" * 64,
        "",
        f"  Paginas localizadas en el Wayback : {len(snapshots):,}",
        f"  Paginas efectivamente descargadas : {bajadas:,}",
    ]
    if bajadas < len(snapshots):
        lineas.append(f"  ATENCION: faltan {len(snapshots) - bajadas} paginas. "
                      f"Volver a correr el script para reintentarlas.")
    lineas += [
        f"  Citas unicas parseadas            : {len(citas):,}",
        f"  Citas reconciliadas               : {total:,}",
        "",
        f"  DOI validado                      : {con_doi:,} ({pct(con_doi)})",
        f"  Sin DOI validado                  : {total - con_doi:,} "
        f"({pct(total - con_doi)})",
    ]
    if motivos:
        lineas.append("")
        lineas.append("  Motivos de no validacion:")
        for m, n in sorted(motivos.items(), key=lambda x: -x[1]):
            lineas.append(f"    {m:<40} {n:>5}  ({pct(n)})")

    lineas += ["",
               f"  Citas en mas de una categoria     : {multi:,} ({pct(multi)})",
               "",
               "  Distribucion por categoria (con solapamiento):"]
    for cat, n in sorted(por_cat.items(), key=lambda x: -x[1]):
        lineas.append(f"    {cat:<45} {n:>5}")

    if anios:
        lineas += ["",
                   f"  Rango de anios (solo DOI validado) : {min(anios)} - {max(anios)}",
                   f"  Mediana                            : "
                   f"{sorted(anios)[len(anios)//2]}"]

    lineas += ["",
               "  Archivos generados:",
               f"    {URLS_FILE.name}",
               f"    {CITAS_FILE.name}",
               f"    {CATEGORIAS_FILE.name}",
               f"    {OUTPUT_FILE.name}",
               f"    {REPORTE_FILE.name}",
               ""]

    reporte = "\n".join(lineas)
    print("\n" + reporte)
    path.write_text(reporte + "\n", encoding="utf-8")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  [CORPUS GTF] - extraer_corpus_gtf_v2")
    print("=" * 60)

    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    print(f"  data   : {CARPETA_DATOS}")
    print(f"  output : {CARPETA_SALIDA}\n")

    snapshots = enumerar_snapshots()
    if not snapshots:
        print("ERROR: no hay paginas para procesar.")
        return
    escribir_csv(URLS_FILE, snapshots, ["url", "timestamp", "categoria"])

    descargar_snapshots(snapshots)

    citas = parsear_paginas(snapshots)
    if not citas:
        print("ERROR: no se parsearon citas. Revisar el HTML en output/gtf/html/.")
        return
    escribir_csv(CITAS_FILE, citas, CAMPOS_CITAS)
    escribir_csv(CATEGORIAS_FILE, tabla_categorias(citas),
                 ["id_cita", "categoria", "titulo", "anio"])

    corpus = reconciliar(citas)
    escribir_csv(OUTPUT_FILE, corpus, CAMPOS_SALIDA)

    escribir_informe(snapshots, citas, corpus, REPORTE_FILE)
    print(f"Archivos en {CARPETA_SALIDA}")


if __name__ == "__main__":
    main()
