"""Comparador de precios y análisis de margen con catálogo base de Imptek en Excel."""
import csv
import json
import os
import re
import threading
import time
import webbrowser
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from urllib.parse import urlparse

import httpx
import pandas as pd
from ddgs import DDGS
from flask import Flask, jsonify, render_template, request, send_file
from groq import Groq

BASE = Path(__file__).resolve().parent
EMPRESAS_FILE = BASE / "empresas.json"
EXCEL_IMPTEK = BASE / "lista de productos.xlsx"
OUT_DIR = BASE / "resultados"
CSV_FILE = OUT_DIR / "precios.csv"
CAMPOS = ["fecha", "consulta", "canal", "empresa", "marca", "tipo", "producto",
          "presentacion", "precio", "moneda", "disponibilidad", "url"]
MAX_POR_TIENDA = 50
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "es-EC,es;q=0.9",
    "Accept": "application/json,text/html;q=0.9,*/*;q=0.8",
}
PRES_RE = re.compile(
    r"\b\d+(?:[.,]\d+)?\s?(?:kg|gr|g|gl|gal\w*|lts?|litros?|l|ml|m2|mts?|m|cm|mm)\b",
    re.I,
)

TIPOS = [
    "Membranas",
    "Revestimientos asfálticos",
    "Revestimientos acrílicos",
    "Tejas",
    "Complementarios",
    "Vial",
    "Impac",
]
OTROS = "Otros"

MAP_CATEGORIA_EXCEL = {
    "MEMBRANAS": "Membranas",
    "REVESTIMIENTO ASFALTICO": "Revestimientos asfálticos",
    "REVESTIMIENTO ASFÁLTICO": "Revestimientos asfálticos",
    "REVESTIMIENTO ACRILICO": "Revestimientos acrílicos",
    "REVESTIMIENTO ACRÍLICO": "Revestimientos acrílicos",
    "TEJAS": "Tejas",
    "COMPLEMENTARIOS": "Complementarios",
    "VIAL": "Vial",
    "IMPAC": "Impac",
}

UNIDADES = {
    "gr": "g", "kgs": "kg", "gal": "gl", "gls": "gl", "galon": "gl",
    "galón": "gl", "galones": "gl", "lt": "l", "lts": "l", "litro": "l",
    "litros": "l", "mt": "m", "mts": "m",
}

app = Flask(__name__)


# ----------------------------------------------------------------- utilidades
def cargar_env():
    env = BASE / ".env"
    if env.exists():
        for linea in env.read_text(encoding="utf-8").splitlines():
            linea = linea.strip()
            if linea and not linea.startswith("#") and "=" in linea:
                k, v = linea.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def cargar_empresas():
    return json.loads(EMPRESAS_FILE.read_text(encoding="utf-8"))


def _normalizar_texto(s):
    s = (s or "").lower()
    s = (s.replace("á", "a").replace("é", "e").replace("í", "i")
          .replace("ó", "o").replace("ú", "u").replace("ñ", "n"))
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return re.sub(r"\s+", " ", s)


def normalizar_categoria_excel(cat):
    clave = _normalizar_texto(cat).upper()
    for k, v in MAP_CATEGORIA_EXCEL.items():
        if _normalizar_texto(k).upper() == clave:
            return v
    return tipo_por_texto(cat) or OTROS


def cargar_catalogo_imptek():
    """Carga los productos y categorías base de Imptek desde lista de productos.xlsx."""
    if not EXCEL_IMPTEK.exists():
        return []
    try:
        xls = pd.ExcelFile(EXCEL_IMPTEK)
        hoja = "Lista de Precios" if "Lista de Precios" in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(EXCEL_IMPTEK, sheet_name=hoja)
    except Exception:
        return []
    if df.shape[1] < 2:
        return []
    df = df.iloc[:, :2].copy()
    df.columns = ["categoria", "descripcion"]
    catalogo = []
    cat_actual = ""
    for _, row in df.iterrows():
        raw_cat = "" if pd.isna(row["categoria"]) else str(row["categoria"]).strip()
        raw_desc = "" if pd.isna(row["descripcion"]) else str(row["descripcion"]).strip()
        if raw_cat and _normalizar_texto(raw_cat) != "categoria":
            cat_actual = normalizar_categoria_excel(raw_cat)
        if not raw_desc or _normalizar_texto(raw_desc) == "descripcion":
            continue
        tipo = cat_actual or tipo_por_texto(raw_desc) or OTROS
        mp = PRES_RE.search(raw_desc)
        catalogo.append({
            "producto": raw_desc,
            "tipo": tipo,
            "presentacion": mp.group(0).strip() if mp else "",
            "norm": _normalizar_texto(raw_desc),
        })
    return catalogo


CATALOGO_IMPTEK = cargar_catalogo_imptek()


def _get(url, **kw):
    r = httpx.get(url, headers=HEADERS, timeout=15, follow_redirects=True, **kw)
    r.raise_for_status()
    return r


def _precio(v):
    try:
        n = float(str(v).strip().replace(",", "."))
        return round(n, 2) if n > 0 else None
    except (TypeError, ValueError):
        return None


def _json(texto):
    limpio = re.sub(r"^```(?:json)?\s*|\s*```$", "", (texto or "").strip(), flags=re.I | re.M)
    m = re.search(r"\{.*\}", limpio, re.S)
    if not m:
        raise ValueError("La respuesta no contenía JSON.")
    return json.loads(re.sub(r",\s*([}\]])", r"\1", m.group(0)))


# ------------------------------------------------- clasificación fija de tipos
def tipo_por_texto(texto):
    """Clasifica estrictamente por palabras clave del producto/categoría."""
    t = _normalizar_texto(texto)
    if not t:
        return None
    if re.search(r"\b(alumband|membran\w*|manto\w*|chovatecho|chovapiso|imperglass|imperpol|superk|techofielt|asfalum)\b", t):
        return "Membranas"
    if re.search(r"\b(asfalt\w*|polibrea|ultraprimer|densobrea|cemento asfaltico)\b", t):
        return "Revestimientos asfálticos"
    if re.search(r"\b(acrilic\w*|elastomer\w*|superacryl|impermeabilizante\w*)\b", t):
        return "Revestimientos acrílicos"
    if re.search(r"\b(teja\w*|fleje|clavo)\b", t):
        return "Tejas"
    if re.search(r"\b(canal\w*|bajante\w*|codo\w*|sujetador\w*|tapa\w*|piton|rejilla\w*|chimenea\w*|cazoleta\w*|rodillo\w*|fazband)\b", t):
        return "Complementarios"
    if re.search(r"\b(vial\w*|bitumix)\b", t):
        return "Vial"
    if re.search(r"\b(impac|weider|cemenflex|wr-480|wr-700)\b", t):
        return "Impac"
    return None


def match_excel_imptek(producto):
    """Busca coincidencia exacta o similar en el Excel de Imptek."""
    if not CATALOGO_IMPTEK:
        return None
    np = _normalizar_texto(producto)
    if not np:
        return None
    # 1. Coincidencia exacta o contenida directa
    for item in CATALOGO_IMPTEK:
        if np == item["norm"] or (len(item["norm"]) >= 6 and item["norm"] in np):
            return item
    # 2. Similitud de secuencia / tokens clave
    mejor, mejor_score = None, 0.0
    tokens_p = set(np.split())
    for item in CATALOGO_IMPTEK:
        ratio = SequenceMatcher(None, np, item["norm"]).ratio()
        tokens_c = set(item["norm"].split())
        overlap = (len(tokens_p & tokens_c) / len(tokens_c)) if tokens_c else 0.0
        score = max(ratio, overlap * 0.92)
        if score > mejor_score:
            mejor, mejor_score = item, score
    return mejor if mejor_score >= 0.62 else None


def variantes_busqueda(consulta):
    """Genera términos de búsqueda exactos y similares basados en la consulta y el Excel de Imptek."""
    q = consulta.strip()
    variantes = [q]
    if q.lower().endswith("s") and len(q) > 4:
        variantes.append(q[:-1])

    q_norm = _normalizar_texto(q)
    similares_excel = [
        item for item in CATALOGO_IMPTEK
        if q_norm in item["norm"]
        or any(tok in item["norm"].split() for tok in q_norm.split() if len(tok) >= 4)
    ]
    for item in similares_excel[:8]:
        variantes.append(item["producto"])
        base_sin_medida = re.sub(r"\b\d+(?:[.,]\d+)?\s*(?:cm|mm|m2|m|kg|g|gr|ml|l|galon\w*|rollo\w*)\b", "", item["producto"], flags=re.I)
        base_sin_medida = re.sub(r"\s*[x×]\s*\d+", "", base_sin_medida)
        base_limpia = re.sub(r"\s+", " ", base_sin_medida).strip()
        if base_limpia and len(base_limpia) >= 4:
            variantes.append(base_limpia)

    out = []
    vistas = set()
    for v in variantes:
        k = _normalizar_texto(v)
        if k and k not in vistas:
            vistas.add(k)
            out.append(v)
    return out[:6]


def _pres_key(texto, producto=""):
    m = PRES_RE.search(texto or "") or PRES_RE.search(producto or "")
    if not m:
        return "N/D"
    n, u = re.match(r"([\d.,]+)\s*(.+)$", m.group(0).lower()).groups()
    try:
        n = "%g" % float(n.replace(",", "."))
    except ValueError:
        return "N/D"
    return f"{n} {UNIDADES.get(u, u)}"


# ------------------------------------------- lectores de catálogo (por plataforma)
def buscar_vtex(e, q):
    base = e["url"].rstrip("/")
    data = _get(f"{base}/api/catalog_system/pub/products/search",
                params={"ft": q, "_from": 0, "_to": MAX_POR_TIENDA - 1}).json()
    if not isinstance(data, list):
        return []
    out = []
    for p in data:
        ofertas = [s.get("commertialOffer") or {}
                   for it in p.get("items", []) for s in it.get("sellers", [])]
        ofertas = [o for o in ofertas if _precio(o.get("Price"))]
        oferta = next((o for o in ofertas if (o.get("AvailableQuantity") or 0) > 0),
                      ofertas[0] if ofertas else None)
        cats = p.get("categories") or []
        out.append({
            "producto": p.get("productName", ""),
            "marca": p.get("brand", ""),
            "precio": _precio(oferta.get("Price")) if oferta else None,
            "disponibilidad": ("en stock" if (oferta.get("AvailableQuantity") or 0) > 0 else "agotado")
                              if oferta else "N/D",
            "url": p.get("link") or base,
            "cat": cats[0].strip("/").split("/")[-1] if cats else "",
        })
    return out


def buscar_woo(e, q):
    base = e["url"].rstrip("/")
    data = _get(f"{base}/wp-json/wc/store/v1/products",
                params={"search": q, "per_page": MAX_POR_TIENDA}).json()
    if not isinstance(data, list):
        return []
    out = []
    for p in data:
        pr = p.get("prices") or {}
        try:
            precio = _precio(int(pr["price"]) / 10 ** int(pr.get("currency_minor_unit", 2)))
        except (KeyError, ValueError, TypeError):
            precio = None
        marcas = p.get("brands") or []
        cats = p.get("categories") or []
        out.append({
            "producto": re.sub(r"<[^>]+>", "", p.get("name", "")),
            "marca": marcas[0].get("name", "") if marcas and isinstance(marcas[0], dict) else "",
            "precio": precio,
            "disponibilidad": "en stock" if p.get("is_in_stock") else "agotado",
            "url": p.get("permalink") or base,
            "cat": cats[0].get("name", "") if cats and isinstance(cats[0], dict) else "",
        })
    return out


LD_RE = re.compile(r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>', re.I | re.S)


def _meta(html, prop):
    m = re.search(r'<meta[^>]+(?:property|name)=["\']%s["\'][^>]+content=["\']([^"\']+)' % re.escape(prop), html, re.I)
    return m.group(1) if m else None


def _walk(o):
    if isinstance(o, dict):
        yield o
        for v in o.values():
            yield from _walk(v)
    elif isinstance(o, list):
        for v in o:
            yield from _walk(v)


def productos_de_html(html, url):
    out = []
    for bloque in LD_RE.findall(html):
        try:
            data = json.loads(bloque.strip())
        except ValueError:
            continue
        for d in _walk(data):
            tipo = d.get("@type")
            if "Product" not in (tipo if isinstance(tipo, list) else [tipo]) or not d.get("name"):
                continue
            ofertas = d.get("offers")
            ofertas = ofertas if isinstance(ofertas, list) else [ofertas or {}]
            precio, disp = None, "N/D"
            for o in ofertas:
                if isinstance(o, dict) and _precio(o.get("price") or o.get("lowPrice")):
                    precio = _precio(o.get("price") or o.get("lowPrice"))
                    av = str(o.get("availability", ""))
                    disp = "en stock" if "InStock" in av else ("agotado" if "OutOfStock" in av or "SoldOut" in av else "N/D")
                    break
            marca = d.get("brand")
            marca = marca.get("name") if isinstance(marca, dict) else marca
            u = d.get("url")
            out.append({
                "producto": str(d["name"]).strip(),
                "marca": str(marca or ""),
                "precio": precio,
                "disponibilidad": disp,
                "url": u if isinstance(u, str) and u.startswith("http") else url,
                "cat": "",
            })
    if not out:
        precio = _meta(html, "product:price:amount") or _meta(html, "og:price:amount")
        titulo = _meta(html, "og:title")
        if precio and titulo:
            out.append({
                "producto": titulo.strip(),
                "marca": _meta(html, "product:brand") or "",
                "precio": _precio(precio),
                "disponibilidad": "N/D",
                "url": url,
                "cat": "",
            })
    return out


def buscar_web(e, q):
    dom = e["dominio"]
    urls = []
    with DDGS() as d:
        for r in d.text(f"site:{dom} {q}", max_results=12):
            u = r.get("href") or ""
            if dom in urlparse(u).netloc and u not in urls:
                urls.append(u)

    def leer(u):
        try:
            return productos_de_html(_get(u).text, u)
        except Exception:
            return []

    with ThreadPoolExecutor(6) as ex:
        return [p for lst in ex.map(leer, urls) for p in lst]


ADAPTADORES = [("VTEX", buscar_vtex), ("WooCommerce", buscar_woo), ("Búsqueda web", buscar_web)]


def consultar_tienda(e, q):
    ultimo = ""
    terminos = variantes_busqueda(q)
    for nombre, fn in ADAPTADORES:
        acumulados, vistos_url = [], set()
        for t in terminos:
            try:
                items = fn(e, t)
            except Exception as ex:
                ultimo = f"{nombre}: {type(ex).__name__}"
                continue
            for it in items:
                clave = (_normalizar_texto(it.get("producto")), it.get("url"))
                if clave not in vistos_url:
                    vistos_url.add(clave)
                    it.update(empresa=e["nombre"], canal=e.get("canal", ""), moneda="USD")
                    acumulados.append(it)
            if len(acumulados) >= MAX_POR_TIENDA:
                break
        if acumulados:
            return acumulados[:MAX_POR_TIENDA], nombre, ""
    return [], "sin datos", ultimo or "Sin resultados"


# -------------------------------------------------- clasificación: marca / tipo
def _normalizar_marcas(items):
    vistos = {}
    for it in items:
        m = (it.get("marca") or "").strip()
        if not m or m.lower() in ("sin marca", "n/d"):
            if "imptek" in _normalizar_texto(it.get("producto")):
                m = "Imptek"
            else:
                it["marca"] = "Sin marca"
                continue
        if m.isupper() or m.islower():
            m = m.title()
        it["marca"] = vistos.setdefault(m.lower(), m)


def clasificar(items, consulta, avisos):
    for it in items:
        mp = PRES_RE.search(it["producto"])
        ref_excel = match_excel_imptek(it["producto"])
        es_imptek = "imptek" in _normalizar_texto(f"{it.get('marca', '')} {it['producto']}")
        if ref_excel and (es_imptek or not it.get("marca")):
            tipo_asignado = ref_excel["tipo"]
            if not it.get("marca"):
                it["marca"] = "Imptek"
        else:
            tipo_asignado = tipo_por_texto(f"{it['producto']} {it.get('cat', '')}") or (ref_excel["tipo"] if ref_excel else OTROS)
        it.update(
            tipo=tipo_asignado,
            presentacion=mp.group(0) if mp else (ref_excel["presentacion"] if ref_excel else ""),
            relevante=True,
        )

    if not os.environ.get("GROQ_API_KEY"):
        avisos.append("Falta GROQ_API_KEY: el tipo se asignó por catálogo Excel y palabras clave.")
        return

    client = Groq(api_key=os.environ["GROQ_API_KEY"])
    modelo = os.environ.get("GROQ_MODEL", "openai/gpt-oss-120b")
    sistema = (
        "Clasificas productos de construcción e impermeabilización en Ecuador. "
        f"El usuario buscó: «{consulta}». Tipos permitidos EXACTOS: "
        "'Membranas', 'Revestimientos asfálticos', 'Revestimientos acrílicos', "
        "'Tejas', 'Complementarios', 'Vial', 'Impac', 'Otros'. "
        "No mezcles membranas (Alumband, mantos, rollos) con recubrimientos acrílicos o asfálticos. "
        "Para cada producto devuelve: marca, tipo, presentacion, relevante (1 o 0). "
        'Responde SOLO JSON: {"items":[{"i":0,"marca":"","tipo":"","presentacion":"","relevante":1}]}'
    )
    for ini in range(0, len(items), 50):
        lote = items[ini:ini + 50]
        entrada = "\n".join(
            json.dumps(
                {"i": ini + k, "n": it["producto"][:90], "b": it["marca"], "c": it.get("cat", "")},
                ensure_ascii=False,
            )
            for k, it in enumerate(lote)
        )
        for intento in range(3):
            try:
                resp = client.chat.completions.create(
                    model=modelo, temperature=0, max_tokens=3500,
                    response_format={"type": "json_object"},
                    messages=[{"role": "system", "content": sistema}, {"role": "user", "content": entrada}],
                )
                for r in _json(resp.choices[0].message.content).get("items", []):
                    i = r.get("i")
                    if isinstance(i, int) and 0 <= i < len(items):
                        it = items[i]
                        if not it["marca"] and r.get("marca"):
                            it["marca"] = str(r["marca"])
                        ref_excel = match_excel_imptek(it["producto"])
                        es_imptek = "imptek" in _normalizar_texto(f"{it.get('marca', '')} {it['producto']}")
                        if ref_excel and es_imptek:
                            it["tipo"] = ref_excel["tipo"]
                        else:
                            it["tipo"] = (
                                tipo_por_texto(r.get("tipo"))
                                or tipo_por_texto(it["producto"])
                                or (r.get("tipo") if r.get("tipo") in TIPOS + [OTROS] else None)
                                or OTROS
                            )
                        it["presentacion"] = str(r.get("presentacion") or it["presentacion"]).strip()
                        it["relevante"] = str(r.get("relevante", 1)).lower() not in ("0", "false")
                break
            except Exception as ex:
                if "429" in str(ex) or "rate" in str(ex).lower():
                    time.sleep(10 * (intento + 1))
                    continue
                avisos.append(f"No se pudo clasificar un lote con Groq ({type(ex).__name__}).")
                break


# ------------------------------------------------------------------ guardado
def guardar(consulta, productos):
    OUT_DIR.mkdir(exist_ok=True)
    if CSV_FILE.exists():
        with open(CSV_FILE, encoding="utf-8-sig", newline="") as f:
            cab = next(csv.reader(f), [])
        if cab != CAMPOS:
            CSV_FILE.rename(OUT_DIR / f"precios_anterior_{datetime.now():%Y%m%d_%H%M%S}.csv")
    nuevo = not CSV_FILE.exists()
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M")
    with open(CSV_FILE, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=CAMPOS)
        if nuevo:
            w.writeheader()
        for p in productos:
            fila = {k: p.get(k, "") for k in CAMPOS}
            fila.update(fecha=fecha, consulta=consulta, precio=p["precio"] if p["precio"] is not None else "N/D")
            w.writerow(fila)


def resumen(productos, tiendas):
    if not productos:
        return "No se encontraron productos."
    marcas = {p["marca"] for p in productos if p["marca"] != "Sin marca"}
    ok = sum(1 for t in tiendas if t["cantidad"])
    txt = f"{len(productos)} productos de {len(marcas)} marcas en {ok} de {len(tiendas)} tiendas."
    precios = [p["precio"] for p in productos if p["precio"] is not None]
    if precios:
        txt += f" Rango de precios: USD {min(precios):.2f} a USD {max(precios):.2f}."
    return txt


# ----------------------------------------------- análisis de margen de precios
CAMPOS_MARGEN = [
    "tipo", "marca_base", "producto_base", "presentacion_base", "precio_base", "tienda_base", "url_base",
    "marca_comparada", "producto_comparado", "presentacion_comparada", "precio_comparado",
    "tienda_comparada", "url_comparada", "margen_usd", "margen_pct"
]


def _resolver_tipo(row):
    ref_excel = match_excel_imptek(row.get("producto", ""))
    es_imptek = "imptek" in _normalizar_texto(f"{row.get('marca', '')} {row.get('producto', '')}")
    if ref_excel and es_imptek:
        return ref_excel["tipo"]
    t_raw = (row.get("tipo") or "").strip()
    if t_raw in TIPOS + [OTROS]:
        return t_raw
    return tipo_por_texto(t_raw) or tipo_por_texto(row.get("producto")) or (ref_excel["tipo"] if ref_excel else OTROS)


def leer_historial():
    if not CSV_FILE.exists():
        return []
    with open(CSV_FILE, encoding="utf-8-sig", newline="") as f:
        filas = list(csv.DictReader(f))
    for r in filas:
        r["tipo"] = _resolver_tipo(r)
    return filas


def filas_comparables(consulta="", tipo=""):
    ultimos = {}
    for r in leer_historial():
        p = _precio(r.get("precio"))
        if p is None or r.get("marca", "Sin marca") == "Sin marca":
            continue
        if (consulta and r["consulta"] != consulta) or (tipo and r["tipo"] != tipo):
            continue
        r["_p"] = p
        r["_pres"] = _pres_key(r.get("presentacion"), r.get("producto"))
        ultimos[(r["empresa"], r["producto"].lower(), r["url"])] = r
    return list(ultimos.values())


def _marcas_unicas(lista):
    vistas, out = set(), []
    for m in lista:
        m = m.strip()
        if m and m.lower() not in vistas:
            vistas.add(m.lower())
            out.append(m)
    return out


def _item(r):
    return {
        "marca": r["marca"],
        "tipo": r["tipo"],
        "producto": r["producto"],
        "precio": r["_p"],
        "presentacion": "" if r["_pres"] == "N/D" else r["_pres"],
        "tienda": r["empresa"],
        "url": r["url"],
    }


def analizar_margen(base, marcas, consulta="", tipo=""):
    """Analiza el margen de precios (USD y %) estrictamente dentro del mismo tipo."""
    items = [_item(r) for r in filas_comparables(consulta, tipo)]
    bases = sorted(
        (i for i in items if i["marca"].lower() == base.lower()),
        key=lambda i: (i["tipo"], i["producto"].lower(), i["tienda"], i["precio"]),
    )
    filas = []
    for b in bases:
        for marca in marcas:
            # Restricción estricta: mismo tipo de producto
            cands = [
                i for i in items
                if i["marca"].lower() == marca.lower() and i["tipo"] == b["tipo"]
            ]
            misma_pres = [c for c in cands if b["presentacion"] and c["presentacion"] == b["presentacion"]]
            candidatos = sorted(misma_pres or cands, key=lambda c: c["precio"])[:3]
            if not candidatos:
                filas.append({
                    "tipo": b["tipo"],
                    "marca_base": b["marca"],
                    "producto_base": b["producto"],
                    "presentacion_base": b["presentacion"],
                    "precio_base": b["precio"],
                    "tienda_base": b["tienda"],
                    "url_base": b["url"],
                    "marca": marca,
                    "producto_comp": None,
                    "presentacion_comp": None,
                    "precio_comp": None,
                    "tienda_comp": None,
                    "url_comp": None,
                    "margen_usd": None,
                    "margen_pct": None,
                })
                continue
            for c in candidatos:
                margen_usd = round(c["precio"] - b["precio"], 2)
                margen_pct = round(((c["precio"] - b["precio"]) / b["precio"]) * 100, 1) if b["precio"] > 0 else None
                filas.append({
                    "tipo": b["tipo"],
                    "marca_base": b["marca"],
                    "producto_base": b["producto"],
                    "presentacion_base": b["presentacion"],
                    "precio_base": b["precio"],
                    "tienda_base": b["tienda"],
                    "url_base": b["url"],
                    "marca": c["marca"],
                    "producto_comp": c["producto"],
                    "presentacion_comp": c["presentacion"],
                    "precio_comp": c["precio"],
                    "tienda_comp": c["tienda"],
                    "url_comp": c["url"],
                    "margen_usd": margen_usd,
                    "margen_pct": margen_pct,
                })

    usd_vals = [f["margen_usd"] for f in filas if f["margen_usd"] is not None]
    pct_vals = [f["margen_pct"] for f in filas if f["margen_pct"] is not None]
    metricas = {
        "total_pares": len(usd_vals),
        "margen_usd_promedio": round(sum(usd_vals) / len(usd_vals), 2) if usd_vals else None,
        "margen_pct_promedio": round(sum(pct_vals) / len(pct_vals), 1) if pct_vals else None,
        "margen_pct_min": min(pct_vals) if pct_vals else None,
        "margen_pct_max": max(pct_vals) if pct_vals else None,
    }
    return {"base": base, "marcas": marcas, "filas": filas, "metricas": metricas}


def csv_margen(res, ruta):
    mapa = {
        "marca_comparada": "marca",
        "producto_comparado": "producto_comp",
        "presentacion_comparada": "presentacion_comp",
        "precio_comparado": "precio_comp",
        "tienda_comparada": "tienda_comp",
        "url_comparada": "url_comp",
    }
    with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=CAMPOS_MARGEN)
        w.writeheader()
        for fila in res["filas"]:
            w.writerow({
                k: ("" if fila.get(mapa.get(k, k)) is None else fila.get(mapa.get(k, k)))
                for k in CAMPOS_MARGEN
            })


def parametros_margen():
    base = request.args.get("base", "").strip()
    marcas = [m for m in _marcas_unicas(request.args.getlist("marcas")) if m.lower() != base.lower()]
    return base, marcas, request.args.get("consulta", ""), request.args.get("tipo", "")


# --------------------------------------------------------------------- rutas
@app.get("/")
def index():
    return render_template("index.html", empresas=cargar_empresas())


@app.get("/api/catalogo-imptek")
def api_catalogo_imptek():
    return jsonify({
        "archivo": EXCEL_IMPTEK.name,
        "total": len(CATALOGO_IMPTEK),
        "items": [{"producto": x["producto"], "tipo": x["tipo"], "presentacion": x["presentacion"]} for x in CATALOGO_IMPTEK],
    })


@app.post("/api/consultar")
def api_consultar():
    body = request.get_json(force=True)
    consulta = (body.get("producto") or "").strip()
    dominios = body.get("dominios") or []
    if not consulta:
        return jsonify(error="Escribe un producto o categoría."), 400
    todas = cargar_empresas()
    empresas = [e for e in todas if e["dominio"] in dominios] or todas

    with ThreadPoolExecutor(len(empresas)) as ex:
        res = list(ex.map(lambda e: consultar_tienda(e, consulta), empresas))
    items, tiendas, vistos = [], [], set()
    for e, (its, metodo, nota) in zip(empresas, res):
        tiendas.append({"empresa": e["nombre"], "dominio": e["dominio"], "metodo": metodo,
                        "cantidad": len(its), "nota": nota})
        for it in its:
            clave = (e["nombre"], it["producto"].lower(), it["url"])
            if clave not in vistos:
                vistos.add(clave)
                items.append(it)

    avisos = []
    if items:
        clasificar(items, consulta, avisos)
        relevantes = [it for it in items if it["relevante"]]
        items = relevantes or items
        _normalizar_marcas(items)
        items.sort(key=lambda p: (p["precio"] is None, p["precio"] or 0, p["producto"].lower()))
        try:
            guardar(consulta, items)
        except OSError:
            avisos.append("No se pudo guardar en precios.csv: ciérralo si lo tienes abierto en Excel.")
    productos = [{k: p.get(k, "") for k in CAMPOS[2:] if k != "precio"} | {"precio": p["precio"]} for p in items]
    for t in tiendas:
        t["cantidad"] = sum(1 for p in productos if p["empresa"] == t["empresa"])
    return jsonify(productos=productos, tiendas=tiendas, avisos=avisos, resumen=resumen(productos, tiendas))


@app.get("/api/historial")
def api_historial():
    return jsonify(leer_historial()[-1500:][::-1])


@app.get("/api/marcas")
def api_marcas():
    filas = filas_comparables()
    cuenta = Counter(r["marca"] for r in filas)
    presentes = {r["tipo"] for r in filas}
    return jsonify(
        marcas=[{"marca": m, "n": n} for m, n in sorted(cuenta.items(), key=lambda x: (-x[1], x[0].lower()))],
        consultas=sorted({r["consulta"] for r in filas}, key=str.lower),
        tipos=[t for t in TIPOS + [OTROS] if t in presentes])


@app.get("/api/comparar")
def api_comparar():
    base, marcas, consulta, tipo = parametros_margen()
    if not base or not marcas:
        return jsonify(error="Elige la marca base y al menos otra marca."), 400
    return jsonify(analizar_margen(base, marcas, consulta, tipo))


@app.get("/descargar-comparacion")
def descargar_comparacion():
    base, marcas, consulta, tipo = parametros_margen()
    if not base or not marcas:
        return "Elige la marca base y al menos otra marca.", 400
    res = analizar_margen(base, marcas, consulta, tipo)
    OUT_DIR.mkdir(exist_ok=True)
    ruta = OUT_DIR / f"margen_{base}_{datetime.now():%Y%m%d_%H%M%S}.csv".replace(" ", "_")
    csv_margen(res, ruta)
    return send_file(ruta, as_attachment=True, download_name=ruta.name)


if __name__ == "__main__":
    cargar_env()
    if not CATALOGO_IMPTEK:
        print("AVISO: no se cargaron productos desde lista de productos.xlsx.")
    else:
        print(f"Catálogo Imptek cargado: {len(CATALOGO_IMPTEK)} productos desde {EXCEL_IMPTEK.name}.")
    if not os.environ.get("GROQ_API_KEY"):
        print("AVISO: no se encontró GROQ_API_KEY en .env.")
    url = "http://127.0.0.1:5000"
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    print(f"Abriendo {url}  (Ctrl+C para cerrar)")
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)