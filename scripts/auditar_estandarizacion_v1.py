"""
auditar_estandarizacion_v1.py
=============================
Reconstruye y audita la estandarizacion indirecta por area tematica de la
Web of Science que sostiene la Tabla 3 del articulo, declarando poblacion,
exclusiones, ponderacion y tratamiento de categorias multiples.

Proyecto: Validacion de un procedimiento de atribucion estructural de
          vinculacion corporativa — expediente europeo de renovacion del
          glifosato (2012). FIC-UDELAR.
Contexto: Resultados · auditoria de la comparacion estandarizada (apartado 4.2)

El script no recalcula el apareo: lo toma de la salida de
cruzar_publit_wos_v2.py. Recupera las categorias tematicas desde la hoja
Base por identificador WoS, en lugar de leer la columna 'categorias' del
cruce, para que la auditoria sea independiente de esa columna.

Entradas:
  - data/Glif_-_Datos_ARS.xlsx                   — hojas 'Base' y 'AU Gephi'
  - output/cruce_publit/cruce_publit_v2.csv      — cruce expediente x WoS
    (si no existe, usa cruce_publit_v1.csv)

Salidas:
  - output/cruce_publit/estandarizacion_auditoria_v1.xlsx — planilla auditable
  - output/cruce_publit/reporte_auditar_estandarizacion.txt — informe estadistico

Sensibilidad de salidas:
  NO SENSIBLE : ambas. Contienen metadatos bibliograficos publicos y
                recuentos agregados.

Uso: python auditar_estandarizacion_v1.py
     (correr desde scripts/ o cualquier ubicacion)

Requisitos: pip install openpyxl

Historial:
  v1 — version inicial. Reconstruye las 22 filas de la Tabla 3 publicada con
       desvio cero, identifica la poblacion del calculo en 700 publicaciones
       distintas (no 749 entradas de lista), desarrolla la formula del valor
       esperado y agrega el chi-cuadrado corregido con correccion de Yates.
"""

import csv
import collections
import unicodedata
from pathlib import Path
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

CARPETA_SCRIPT  = Path(__file__).resolve().parent
CARPETA_PROY    = CARPETA_SCRIPT.parent
CARPETA_DATA    = CARPETA_PROY / "data"
CARPETA_SALIDA  = CARPETA_PROY / "output" / "cruce_publit"

# Entradas
BASE_XLSX    = CARPETA_DATA / "Glif_-_Datos_ARS.xlsx"
CRUCE_CSV    = CARPETA_SALIDA / "cruce_publit_v2.csv"
CRUCE_ALT    = CARPETA_SALIDA / "cruce_publit_v1.csv"

# Salidas
OUTPUT_FILE  = CARPETA_SALIDA / "estandarizacion_auditoria_v1.xlsx"
REPORTE_FILE = CARPETA_SALIDA / "reporte_auditar_estandarizacion.txt"

# Hojas e indices de columna de Glif_-_Datos_ARS.xlsx
HOJA_BASE, HOJA_ZONA = "Base", "AU Gephi"
COL_ID, COL_AU, COL_TI, COL_TC, COL_PY, COL_DI, COL_WC = 0, 2, 9, 32, 45, 55, 58
COL_Z_LABEL, COL_Z_ARTS, COL_Z_ZONA = 2, 3, 4

# Parametros del calculo
VENTANA        = (2001, 2012)   # ventana de las tasas de referencia
MIN_BASE_CALC  = 20             # minimo de articulos en la base para usar un area
MIN_EXP_MOSTRAR = 10            # minimo de publicaciones del expediente para mostrarla
MIN_BASE_MOSTRAR = 20
N_INESTABLE    = 30             # por debajo de este n la razon se marca inestable

# Formato
AZUL = PatternFill("solid", fgColor="2F5496")
BLANCO_NEGRITA = Font(bold=True, color="FFFFFF")
NEGRITA = Font(bold=True)
BORDE_INF = Border(bottom=Side(style="thin", color="BFBFBF"))


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def normalizar(s):
    """Mayusculas sin diacriticos ni espacios sobrantes, para cotejar nombres."""
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
    print(f"  No se pudo leer {path.name}")
    return [], []


def anio(valor):
    try:
        return int(str(valor)[:4])
    except (TypeError, ValueError):
        return None


def chi2_yates(a, b, c, d):
    """Chi-cuadrado de una tabla 2x2 con correccion de continuidad de Yates."""
    n = a + b + c + d
    num = n * (abs(a * d - b * c) - n / 2) ** 2
    den = (a + b) * (c + d) * (a + c) * (b + d)
    return num / den if den else 0.0


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
    """Devuelve (zona_por_autor, arts_por_autor, registros_base)."""
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
            arts[etiqueta] = fila[COL_Z_ARTS]

    registros = {}
    for i, fila in enumerate(wb[HOJA_BASE].iter_rows(values_only=True)):
        if i == 0:
            continue
        registros[str(fila[COL_ID]).strip()] = {
            "autores": [normalizar(a) for a in str(fila[COL_AU] or "").split(";") if a.strip()],
            "anio": anio(fila[COL_PY]),
            "areas": sorted({c.strip() for c in str(fila[COL_WC] or "").split(";") if c.strip()}),
            "titulo": str(fila[COL_TI] or ""),
            "doi": str(fila[COL_DI] or ""),
            "citas": fila[COL_TC],
        }
    wb.close()
    return zona, arts, registros


def cargar_cruce():
    path = CRUCE_CSV if CRUCE_CSV.exists() else CRUCE_ALT
    filas, _ = leer_csv_auto(path)
    return [f for f in filas if str(f.get("estado_apareo", "")).strip() == "APAREADO"], path


# ─── PROCESAMIENTO ────────────────────────────────────────────────────────────

VERDADERO = {"1", "TRUE", "VERDADERO", "SI", "SÍ"}


def tiene_zona1_csv(fila):
    return str(fila.get("tiene_zona1", "")).strip().upper() in VERDADERO


def tasas_referencia(zona, registros):
    """Tasa de publicaciones con >=1 autor de Zona 1, por area, en la ventana."""
    en_ventana = [r for r in registros.values()
                  if r["anio"] and VENTANA[0] <= r["anio"] <= VENTANA[1]]
    total, con_z1 = collections.Counter(), collections.Counter()
    for reg in en_ventana:
        marcado = any(zona.get(a) == "1" for a in reg["autores"])
        for area in reg["areas"]:
            total[area] += 1
            if marcado:
                con_z1[area] += 1
    tasa = {a: con_z1[a] / total[a] for a in total}
    n_base = len(en_ventana)
    z1_base = sum(1 for r in en_ventana if any(zona.get(a) == "1" for a in r["autores"]))
    return tasa, total, con_z1, n_base, z1_base


def poblacion(apareadas, registros):
    """Embudo: entradas apareadas -> publicaciones distintas -> en ventana."""
    vistos, repetidas = {}, 0
    for fila in apareadas:
        clave = str(fila.get("id_wos", "")).strip()
        if clave in vistos:
            repetidas += 1
        else:
            vistos[clave] = fila
    distintas = [(k, v) for k, v in vistos.items() if k in registros]
    en_ventana, fuera = [], 0
    for clave, fila in distintas:
        a = registros[clave]["anio"]
        if a and VENTANA[0] <= a <= VENTANA[1]:
            en_ventana.append((clave, fila))
        else:
            fuera += 1
    return en_ventana, len(distintas), repetidas, fuera


def estandarizar(seleccion, registros, tasa, total_base):
    """Estandarizacion indirecta. Devuelve el diccionario de resultados."""
    n_area, obs_area = collections.Counter(), collections.Counter()
    for clave, fila in seleccion:
        marcado = tiene_zona1_csv(fila)
        for area in registros[clave]["areas"]:
            n_area[area] += 1
            if marcado:
                obs_area[area] += 1

    areas_calculo = sorted([a for a in n_area
                            if a in tasa and total_base.get(a, 0) >= MIN_BASE_CALC],
                           key=lambda a: -n_area[a])
    instancias = sum(n_area[a] for a in areas_calculo)
    suma_esp = sum(n_area[a] * tasa[a] for a in areas_calculo)

    n_pub = len(seleccion)
    obs = sum(1 for _, f in seleccion if tiene_zona1_csv(f))
    tasa_esp = suma_esp / instancias if instancias else 0.0
    esperados = n_pub * tasa_esp

    # Ponderacion alternativa: cada publicacion con el promedio simple de sus areas
    alt_n = alt_obs = 0
    alt_esp = 0.0
    for clave, fila in seleccion:
        usables = [a for a in registros[clave]["areas"] if a in areas_calculo]
        if not usables:
            continue
        alt_n += 1
        if tiene_zona1_csv(fila):
            alt_obs += 1
        alt_esp += sum(tasa[a] for a in usables) / len(usables)

    areas_mostrar = sorted([a for a in n_area
                            if n_area[a] >= MIN_EXP_MOSTRAR
                            and total_base.get(a, 0) >= MIN_BASE_MOSTRAR],
                           key=lambda a: -n_area[a])

    return {
        "n_area": n_area, "obs_area": obs_area,
        "areas_calculo": areas_calculo, "areas_mostrar": areas_mostrar,
        "instancias": instancias, "suma_esp": suma_esp,
        "n_pub": n_pub, "obs": obs, "tasa_esp": tasa_esp, "esperados": esperados,
        "razon": obs / esperados if esperados else 0.0,
        "alt_n": alt_n, "alt_obs": alt_obs, "alt_esp": alt_esp,
        "alt_razon": alt_obs / alt_esp if alt_esp else 0.0,
    }


# ─── PLANILLA ─────────────────────────────────────────────────────────────────

def escribir_planilla(res, emb, tasa, total_base, z1_base, n_base, apareadas, path):
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "1_Poblacion"
    encabezar(ws, ["Paso", "Unidad de observacion", "n", "Comentario"], [8, 44, 10, 88])
    obs_entradas = sum(1 for f in apareadas if tiene_zona1_csv(f))
    pasos = [
        ("1", "Entradas apareadas de la lista del expediente", len(apareadas),
         "Una entrada por aparicion; una misma publicacion puede figurar en mas de una seccion"),
        ("2", "Publicaciones distintas en el conjunto apareado", emb["distintas"],
         f"Se descartan {emb['repetidas']} entradas que remiten a un registro WoS ya apareado"),
        ("3", f"Publicaciones distintas con anio {VENTANA[0]}-{VENTANA[1]}", res["n_pub"],
         f"Se excluyen {emb['fuera']} publicaciones fuera de la ventana de las tasas de referencia"),
        ("", "  de ellas, con al menos un autor de Zona 1", res["obs"],
         f"{100 * res['obs'] / res['n_pub']:.1f}% de la poblacion. Es el OBSERVADO de la estandarizacion"),
        ("", "Entradas apareadas con al menos un autor de Zona 1", obs_entradas,
         f"{100 * obs_entradas / len(apareadas):.1f}% de las entradas. Cifra de la comparacion cruda"),
    ]
    for p in pasos:
        ws.append(list(p))
    for fila in ws.iter_rows(min_row=2):
        fila[0].alignment = Alignment(horizontal="center")
        fila[2].alignment = Alignment(horizontal="center")
        fila[1].alignment = Alignment(wrap_text=True, vertical="center")
        fila[3].alignment = Alignment(wrap_text=True, vertical="center")
    ws.append([])
    ws.append(["Nota", "La estandarizacion se calcula sobre publicaciones distintas (paso 3) "
                       "y no sobre entradas de lista (paso 1). De ahi la diferencia entre ambas cifras.",
               "", ""])
    ws.cell(ws.max_row, 1).font = NEGRITA

    ws = wb.create_sheet("2_Tasas_base")
    encabezar(ws, ["Area tematica (WoS)", f"Articulos en la base {VENTANA[0]}-{VENTANA[1]}",
                   "Con >=1 autor de Zona 1", "Tasa de referencia",
                   "Usada en el calculo", "Mostrada en Tabla 3"], [46, 16, 16, 14, 14, 14])
    for area in sorted(tasa, key=lambda a: -total_base[a]):
        ws.append([area, total_base[area], round(total_base[area] * tasa[area]),
                   round(tasa[area], 5),
                   "si" if area in res["areas_calculo"] else "no",
                   "si" if area in res["areas_mostrar"] else "no"])
        ws.cell(ws.max_row, 4).number_format = "0.0%"
    ws.append([])
    ws.append([f"Total base {VENTANA[0]}-{VENTANA[1]}", n_base, z1_base,
               round(z1_base / n_base, 5), "", ""])
    for celda in ws[ws.max_row]:
        celda.font = NEGRITA
    ws.cell(ws.max_row, 4).number_format = "0.0%"

    ws = wb.create_sheet("3_Estandarizacion")
    encabezar(ws, ["Area tematica (WoS)", "Publicaciones del expediente en el area (a)",
                   "Con >=1 autor de Zona 1 (b)", "% con Zona 1 (b/a)",
                   "Tasa de referencia del area (c)", "Esperados en el area (a x c)",
                   "Razon del area (b / a x c)"], [46, 15, 14, 12, 14, 13, 14])
    for area in res["areas_calculo"]:
        esp = res["n_area"][area] * tasa[area]
        ws.append([area, res["n_area"][area], res["obs_area"][area],
                   round(res["obs_area"][area] / res["n_area"][area], 4),
                   round(tasa[area], 5), round(esp, 2),
                   round(res["obs_area"][area] / esp, 3) if esp else 0])
        ws.cell(ws.max_row, 4).number_format = "0.0%"
        ws.cell(ws.max_row, 5).number_format = "0.0%"
        ws.cell(ws.max_row, 6).number_format = "0.00"
        ws.cell(ws.max_row, 7).number_format = "0.000"
    inicio = ws.max_row + 2
    ws.cell(inicio, 1, "CALCULO GLOBAL").font = NEGRITA
    detalle = [
        ("Publicaciones distintas en la ventana  (N)", res["n_pub"], "0",
         "Paso 3 de la hoja 1_Poblacion"),
        ("Observados: publicaciones con >=1 autor de Zona 1  (O)", res["obs"], "0",
         f"{100 * res['obs'] / res['n_pub']:.1f}% de N"),
        (f"Areas con tasa utilizable (>={MIN_BASE_CALC} articulos en la base)",
         len(res["areas_calculo"]), "0",
         "Criterio de calculo, mas amplio que el de presentacion de la Tabla 3"),
        ("Instancias area-publicacion  (suma de la columna a)", res["instancias"], "0",
         "Una publicacion aporta tantas instancias como areas tenga asignadas"),
        ("Suma de esperados por area  (suma de a x c)", round(res["suma_esp"], 2), "0.00",
         "Numerador de la tasa esperada"),
        ("Tasa esperada ponderada por instancias  (T)", round(res["tasa_esp"], 5), "0.0%",
         "Promedio de las tasas de referencia, ponderado por el peso de cada area en el corpus"),
        ("ESPERADOS  (E = N x T)", round(res["esperados"], 1), "0.0",
         "Publicaciones que cabria esperar con >=1 autor de Zona 1 bajo la composicion tematica"),
        ("RAZON OBSERVADO / ESPERADO  (O / E)", round(res["razon"], 3), "0.000",
         "Valor reportado en el articulo"),
    ]
    for i, (etq, val, fmt, com) in enumerate(detalle):
        ws.cell(inicio + 1 + i, 1, etq)
        celda = ws.cell(inicio + 1 + i, 2, val)
        celda.number_format = fmt
        ws.cell(inicio + 1 + i, 3, com).alignment = Alignment(wrap_text=True)
        if i >= 5:
            ws.cell(inicio + 1 + i, 1).font = NEGRITA
            celda.font = NEGRITA

    ws = wb.create_sheet("4_Tabla3_articulo")
    encabezar(ws, ["Area tematica (WoS)", "Publicaciones", "% con Zona 1", "Obs./esp.",
                   f"n < {N_INESTABLE} (valor inestable)"], [46, 13, 12, 11, 16])
    for area in res["areas_mostrar"]:
        esp = res["n_area"][area] * tasa[area]
        ws.append([area, res["n_area"][area],
                   round(res["obs_area"][area] / res["n_area"][area], 4),
                   round(res["obs_area"][area] / esp, 2) if esp else 0,
                   "*" if res["n_area"][area] < N_INESTABLE else ""])
        ws.cell(ws.max_row, 3).number_format = "0.0%"
        ws.cell(ws.max_row, 4).number_format = "0.00"
    ws.append(["Total (todas las areas)", res["n_pub"],
               round(res["obs"] / res["n_pub"], 4), round(res["razon"], 2), ""])
    for celda in ws[ws.max_row]:
        celda.font = NEGRITA
        celda.border = BORDE_INF
    ws.cell(ws.max_row, 3).number_format = "0.0%"
    ws.cell(ws.max_row, 4).number_format = "0.00"

    ws = wb.create_sheet("5_Sensibilidad")
    encabezar(ws, ["Esquema de ponderacion", "Publicaciones", "Observados", "Esperados",
                   "Razon", "Descripcion"], [40, 13, 12, 12, 10, 74])
    ws.append(["Por instancia area-publicacion", res["n_pub"], res["obs"],
               round(res["esperados"], 1), round(res["razon"], 3),
               "Cada publicacion pesa en proporcion a la cantidad de areas que tiene asignadas. "
               "Es el esquema reportado en el articulo."])
    ws.append(["Por publicacion", res["alt_n"], res["alt_obs"], round(res["alt_esp"], 1),
               round(res["alt_razon"], 3),
               "Cada publicacion pesa uno y su tasa esperada es el promedio simple "
               "de las tasas de sus areas."])
    for fila in ws.iter_rows(min_row=2):
        fila[5].alignment = Alignment(wrap_text=True, vertical="center")
    ws.append([])
    ws.append(["Nota", "Las dos ponderaciones acotan el resultado. Ninguna altera el signo: "
                       "el expediente presenta menos publicaciones con autores de Zona 1 que lo "
                       "esperable segun su composicion tematica.", "", "", "", ""])
    ws.cell(ws.max_row, 1).font = NEGRITA

    wb.save(path)


# ─── INFORME ──────────────────────────────────────────────────────────────────

def escribir_informe(res, emb, tasa, n_base, z1_base, apareadas, origen, path):
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    obs_ent = sum(1 for f in apareadas if tiene_zona1_csv(f))
    x2_ent = chi2_yates(obs_ent, len(apareadas) - obs_ent, z1_base, n_base - z1_base)
    x2_pub = chi2_yates(res["obs"], res["n_pub"] - res["obs"], z1_base, n_base - z1_base)

    lineas = [
        "=" * 68,
        "  [AUDITORIA DE LA ESTANDARIZACION] - auditar_estandarizacion_v1",
        f"  Fecha: {ahora}",
        f"  Cruce de origen: {origen.name}",
        "=" * 68,
        "",
        "  POBLACION DEL CALCULO",
        f"    Entradas apareadas de la lista          : {len(apareadas):,}",
        f"    Publicaciones distintas                 : {emb['distintas']:,}"
        f"  (-{emb['repetidas']} entradas repetidas)",
        f"    En la ventana {VENTANA[0]}-{VENTANA[1]}                 : {res['n_pub']:,}"
        f"  (-{emb['fuera']} fuera de ventana)",
        "",
        "  COMPARACION CRUDA",
        f"    Entradas con >=1 autor de Zona 1        : {obs_ent:,} / {len(apareadas):,}"
        f" = {100 * obs_ent / len(apareadas):.1f}%",
        f"    Publicaciones con >=1 autor de Zona 1   : {res['obs']:,} / {res['n_pub']:,}"
        f" = {100 * res['obs'] / res['n_pub']:.1f}%",
        f"    Linea de base WoS {VENTANA[0]}-{VENTANA[1]}            : {z1_base:,} / {n_base:,}"
        f" = {100 * z1_base / n_base:.1f}%",
        f"    Chi-cuadrado de Yates, sobre entradas   : {x2_ent:.2f}",
        f"    Chi-cuadrado de Yates, sobre publicac.  : {x2_pub:.2f}   <- valor a reportar",
        "",
        "  ESTANDARIZACION INDIRECTA POR AREA TEMATICA",
        f"    Areas con tasa utilizable (>={MIN_BASE_CALC} en base) : {len(res['areas_calculo'])}",
        f"    Instancias area-publicacion             : {res['instancias']:,}",
        f"    Tasa esperada ponderada por instancias  : {100 * res['tasa_esp']:.1f}%",
        f"    Observados                              : {res['obs']:,}",
        f"    Esperados  (N x T)                      : {res['esperados']:.1f}",
        f"    Razon observado / esperado              : {res['razon']:.2f}",
        "",
        "  SENSIBILIDAD AL TRATAMIENTO DE CATEGORIAS MULTIPLES",
        f"    Ponderando por instancia area-publicacion : razon {res['razon']:.2f}",
        f"    Ponderando por publicacion                : razon {res['alt_razon']:.2f}",
        "",
        f"  TABLA 3 DEL ARTICULO ({len(res['areas_mostrar'])} areas mostradas)",
        f"    {'area':44s} {'n':>5} {'%z1':>7} {'o/e':>6}",
    ]
    for area in res["areas_mostrar"]:
        esp = res["n_area"][area] * tasa[area]
        marca = " *" if res["n_area"][area] < N_INESTABLE else ""
        lineas.append(f"    {area:44s} {res['n_area'][area]:>5}"
                      f" {100 * res['obs_area'][area] / res['n_area'][area]:>6.1f}%"
                      f" {res['obs_area'][area] / esp if esp else 0:>6.2f}{marca}")
    lineas += [
        f"    {'Total (todas las areas)':44s} {res['n_pub']:>5}"
        f" {100 * res['obs'] / res['n_pub']:>6.1f}% {res['razon']:>6.2f}",
        "",
        "  NOTA: las areas se solapan, de modo que las filas no suman al total.",
        f"        Las razones calculadas sobre menos de {N_INESTABLE} publicaciones (*)",
        "        son inestables y se incluyen con fines descriptivos.",
        "",
        "  Archivos generados:",
        f"    {OUTPUT_FILE.name}",
        f"    {REPORTE_FILE.name}",
    ]
    reporte = "\n".join(lineas)
    print(reporte)
    path.write_text(reporte + "\n", encoding="utf-8")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 68)
    print("  [AUDITORIA DE LA ESTANDARIZACION] - auditar_estandarizacion_v1")
    print("=" * 68)

    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

    print("\n[1/4] Cargando la base de 2018...")
    zona, _arts, registros = cargar_base()
    print(f"      {len(registros):,} registros, {len(zona):,} autores con zona.")

    print("[2/4] Cargando el cruce expediente x WoS...")
    apareadas, origen = cargar_cruce()
    if not apareadas:
        print("ERROR: no hay entradas apareadas para procesar.")
        return
    print(f"      {len(apareadas):,} entradas apareadas desde {origen.name}.")

    print("[3/4] Calculando tasas de referencia y estandarizando...")
    tasa, total_base, _z1_area, n_base, z1_base = tasas_referencia(zona, registros)
    seleccion, n_distintas, n_repetidas, n_fuera = poblacion(apareadas, registros)
    emb = {"distintas": n_distintas, "repetidas": n_repetidas, "fuera": n_fuera}
    res = estandarizar(seleccion, registros, tasa, total_base)

    print("[4/4] Escribiendo planilla e informe...\n")
    escribir_planilla(res, emb, tasa, total_base, z1_base, n_base, apareadas, OUTPUT_FILE)
    escribir_informe(res, emb, tasa, n_base, z1_base, apareadas, origen, REPORTE_FILE)
    print(f"\nArchivos en {CARPETA_SALIDA}")


if __name__ == "__main__":
    main()
