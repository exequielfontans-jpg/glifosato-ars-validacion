"""
comparar_gtf_publit_v1.py
=========================
Mide cuanto se solapan el corpus publicado en el sitio del Glyphosate Task
Force (glifosato.es) y la lista de publicaciones del expediente europeo de
2012. Responde a una sola pregunta: son dos corpus distintos o el mismo
cuerpo de literatura presentado de dos maneras.

Estructura de carpetas del proyecto:
  proyecto/
    script/     los .py
    data/       fuentes originales, NUNCA se escribe aqui
    output/     todo lo que generan los scripts, una subcarpeta por script

Si se borra output/ y se vuelven a correr los scripts en orden, el resultado
debe ser identico. Por eso las salidas no van en data/.

Proyecto: Corpus curados sobre glifosato. FIC-UDELAR.
Contexto: Verificacion previa al envio · independencia de los corpus


Entradas:
  - output/gtf/corpus_gtf_v2.csv     - salida de extraer_corpus_gtf_v2.py
  - output/publit/publit_v3.csv      - salida de extraer_publit_v3.py

Salidas:
  - output/comparacion/solapamiento_gtf_publit_v1.xlsx
       hoja 'Resumen'   - las cifras de solapamiento
       hoja 'Comunes'   - referencias presentes en ambos corpus
       hoja 'Solo GTF'  - presentes solo en el sitio
       hoja 'Solo PubLit' - presentes solo en el expediente
  - output/comparacion/reporte_comparar_gtf_publit.txt

Sensibilidad de salidas:
  NO SENSIBLE : todas (metadatos bibliograficos publicos)

Uso: python comparar_gtf_publit_v1.py

Requisitos: pip install openpyxl

Historial:
  v1 - version inicial: cotejo por DOI y, en su defecto, por titulo
       normalizado; el porcentaje se calcula sobre las referencias
       efectivamente comparables, no sobre el total.
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
CARPETA_SALIDA = CARPETA_OUT / "comparacion"

GTF_FILE    = CARPETA_OUT / "gtf" / "corpus_gtf_v2.csv"
PUBLIT_FILE = CARPETA_OUT / "publit" / "publit_v3.csv"

XLSX_FILE    = CARPETA_SALIDA / "solapamiento_gtf_publit_v1.xlsx"
REPORTE_FILE = CARPETA_SALIDA / "reporte_comparar_gtf_publit.txt"

# Solo se comparan las referencias del expediente de esta presentacion:
# la de 2020 corresponde a otro dossier y no es comparable con el sitio.
PRESENTACION = "2012"


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def leer_csv_auto(path):
    if not path.exists():
        print(f"  ERROR: no se encontro {path}")
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
    print(f"  ERROR: no se pudo leer {path.name}")
    return []


def norm_doi(txt):
    t = str(txt or "").strip().lower()
    t = re.sub(r"^https?://(dx\.)?doi\.org/", "", t)
    return t.strip(" .")


def norm_titulo(txt):
    t = unicodedata.normalize("NFKD", str(txt or "").lower())
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = re.sub(r"[^a-z0-9 ]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def claves(fila, campo_titulo="titulo"):
    """Devuelve (clave_doi, clave_titulo) de una referencia."""
    return norm_doi(fila.get("doi")), norm_titulo(fila.get(campo_titulo))


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 64)
    print("  [SOLAPAMIENTO GTF x EXPEDIENTE] - comparar_gtf_publit_v1")
    print("=" * 64)

    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    print(f"  data   : {CARPETA_DATOS}")
    print(f"  output : {CARPETA_SALIDA}\n")
    gtf = leer_csv_auto(GTF_FILE)
    pub = [r for r in leer_csv_auto(PUBLIT_FILE)
           if str(r.get("presentacion", "")).strip() == PRESENTACION]
    if not gtf or not pub:
        return

    print(f"\n  Sitio del GTF        : {len(gtf):,} referencias")
    print(f"  Expediente {PRESENTACION}       : {len(pub):,} referencias")

    # indices del expediente
    pub_doi = {}
    pub_tit = {}
    for r in pub:
        d, t = claves(r)
        if d:
            pub_doi.setdefault(d, r)
        if t:
            pub_tit.setdefault(t, r)

    comunes, solo_gtf = [], []
    comparables_doi = comparables_tit = 0
    por_doi = por_titulo = 0
    ids_pub_usados = set()

    for r in gtf:
        d, t = claves(r)
        hallado, via = None, ""
        if d:
            comparables_doi += 1
            if d in pub_doi:
                hallado, via = pub_doi[d], "DOI"
                por_doi += 1
        if hallado is None and t:
            comparables_tit += 1
            if t in pub_tit:
                hallado, via = pub_tit[t], "TITULO"
                por_titulo += 1

        if hallado is not None:
            ids_pub_usados.add(hallado.get("id_ref", ""))
            comunes.append({
                "via": via,
                "id_gtf": r.get("id_cita", ""),
                "id_publit": hallado.get("id_ref", ""),
                "doi": d,
                "titulo": (r.get("titulo") or "")[:180],
                "anio": r.get("anio", ""),
                "categorias_gtf": r.get("categorias", ""),
                "seccion_publit": hallado.get("seccion", ""),
                "categoria_publit": hallado.get("categoria", ""),
                "evaluacion_publit": hallado.get("evaluacion", ""),
            })
        else:
            solo_gtf.append({
                "id_gtf": r.get("id_cita", ""),
                "doi": d, "anio": r.get("anio", ""),
                "titulo": (r.get("titulo") or "")[:180],
                "categorias_gtf": r.get("categorias", ""),
                "comparable": "SI" if (d or t) else "NO",
            })

    solo_pub = [{
        "id_publit": r.get("id_ref", ""),
        "doi": norm_doi(r.get("doi")), "anio": r.get("anio", ""),
        "titulo": (r.get("titulo") or "")[:180],
        "seccion_publit": r.get("seccion", ""),
        "categoria_publit": r.get("categoria", ""),
    } for r in pub if r.get("id_ref", "") not in ids_pub_usados]

    # El porcentaje honesto se calcula sobre lo que se pudo comparar.
    comparables = sum(1 for r in gtf if claves(r)[0] or claves(r)[1])
    n_com = len(comunes)
    pct_gtf = 100 * n_com / comparables if comparables else 0
    pct_pub = 100 * n_com / len(pub) if pub else 0

    if pct_gtf >= 80:
        veredicto = ("MISMO CORPUS. El sitio y el expediente presentan el mismo "
                     "cuerpo de literatura. Tratarlos como un unico corpus con "
                     "dos presentaciones y rehacer la Tabla 1.")
    elif pct_gtf >= 35:
        veredicto = ("CORPUS EMPARENTADOS PERO DISTINTOS. La diferencia es en si "
                     "un resultado: que literatura se muestra en publico y cual "
                     "se presenta al regulador.")
    else:
        veredicto = ("CORPUS INDEPENDIENTES. Se sostiene el tratamiento actual "
                     "como dos corpus separados.")

    # ── informe ──
    pct = lambda k, d: f"{(100*k/d):.1f}%" if d else "-"
    lineas = [
        "=" * 64,
        "  [SOLAPAMIENTO GTF x EXPEDIENTE] - comparar_gtf_publit_v1",
        f"  Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "=" * 64,
        "",
        f"  Sitio del GTF                   : {len(gtf):,}",
        f"    de ellas comparables          : {comparables:,} "
        f"({pct(comparables, len(gtf))})",
        f"  Expediente {PRESENTACION}                  : {len(pub):,}",
        "",
        f"  Referencias en ambos corpus     : {n_com:,}",
        f"    por DOI                       : {por_doi:,}",
        f"    por titulo normalizado        : {por_titulo:,}",
        "",
        f"  Solapamiento sobre el GTF comparable : {pct_gtf:.1f}%",
        f"  Solapamiento sobre el expediente     : {pct_pub:.1f}%",
        "",
        f"  Solo en el sitio    : {len(solo_gtf):,}",
        f"  Solo en el expediente : {len(solo_pub):,}",
        "",
        "  VEREDICTO:",
    ]
    for i in range(0, len(veredicto), 62):
        lineas.append(f"    {veredicto[i:i+62]}")
    lineas += ["",
               "  NOTA: el porcentaje se calcula sobre las referencias con DOI o",
               "        titulo utilizable, no sobre el total. Revisar la hoja",
               "        'Comunes' para verificar algunos casos a mano.",
               "", f"  Archivo: {XLSX_FILE.name}", ""]
    reporte = "\n".join(lineas)
    print("\n" + reporte)
    REPORTE_FILE.write_text(reporte + "\n", encoding="utf-8")

    # ── excel ──
    gris = PatternFill("solid", fgColor="D9D9D9")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumen"
    for fila in [["SOLAPAMIENTO ENTRE LOS DOS CORPUS DEL CONSORCIO"], [],
                 ["Sitio del GTF (total)", len(gtf)],
                 ["  comparables (con DOI o titulo)", comparables],
                 [f"Expediente {PRESENTACION}", len(pub)], [],
                 ["Referencias en ambos", n_com],
                 ["  por DOI", por_doi],
                 ["  por titulo", por_titulo], [],
                 ["Solapamiento sobre el GTF comparable", round(pct_gtf, 1)],
                 ["Solapamiento sobre el expediente", round(pct_pub, 1)], [],
                 ["Solo en el sitio", len(solo_gtf)],
                 ["Solo en el expediente", len(solo_pub)], [],
                 ["VEREDICTO"], [veredicto]]:
        ws.append(fila)
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 14
    for f in ws.iter_rows():
        for c in f:
            c.font = Font(name="Arial")
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for r in (1, 16):
        ws.cell(row=r, column=1).font = Font(name="Arial", bold=True)

    for nombre, filas in (("Comunes", comunes), ("Solo GTF", solo_gtf),
                          ("Solo PubLit", solo_pub)):
        hoja = wb.create_sheet(nombre)
        if not filas:
            hoja.append(["(sin registros)"])
            continue
        campos = list(filas[0].keys())
        hoja.append(campos)
        for c in hoja[1]:
            c.font = Font(name="Arial", bold=True)
            c.fill = gris
        for f in filas:
            hoja.append([f.get(k, "") for k in campos])
        anchos = {"titulo": 62, "doi": 26, "categorias_gtf": 30,
                  "seccion_publit": 22}
        for i, campo in enumerate(campos, 1):
            hoja.column_dimensions[get_column_letter(i)].width = anchos.get(campo, 14)
        for f in hoja.iter_rows(min_row=2):
            for c in f:
                c.font = Font(name="Arial")
        hoja.auto_filter.ref = f"A1:{get_column_letter(len(campos))}{len(filas)+1}"
        hoja.freeze_panes = "A2"

    wb.save(XLSX_FILE)
    print(f"Archivos en {CARPETA_SALIDA}")


if __name__ == "__main__":
    main()
