# -*- coding: utf-8 -*-
"""
build.py — Revisión de Estrategias · Abril 2026
Procesa los PPTs de las ~55 estrategias del mes, las cruza contra los catálogos
de fondos y ETFs (con data granular), aplica reglas vs house view y arma un
único HTML con la misma estética que el index.html piloto.

Uso:  python build.py
Output: sobreescribe ./index.html y loggea summary en stdout.
"""

from __future__ import annotations
import os
import re
import sys
import html
import base64
import logging
import unicodedata
from pathlib import Path
from collections import defaultdict

import pandas as pd
import pdfplumber

# ---------------------------------------------------------------------------
# LOGO
# ---------------------------------------------------------------------------
def _logo_b64():
    p = Path(__file__).parent / "assets" / "logo-latam-color.png"
    return base64.b64encode(p.read_bytes()).decode()

LOGO_B64 = _logo_b64()
LOGO_DATA = f"data:image/png;base64,{LOGO_B64}"

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
PDF_ROOT = Path(
    r"C:/Users/mmaut/LATAM ConsultUs/Santiago De Haedo - LATAM Consultus – Compartida"
    r"/Portafolios/Fin de mes/Presentaciones + Factsheets/2026/ABRIL"
)
OUT_HTML = ROOT / "index.html"

# Regex de detección de PPTs válidos (no factsheets, no informes de gestión)
PPT_PATTERNS = [
    re.compile(r"^PPT.*30[-.]?4[-.]?2026.*\.pdf$", re.I),
    re.compile(r"^PPT.*30\.04\.2026.*\.pdf$", re.I),
    re.compile(r"^PPT.*19\.05\.2026.*\.pdf$", re.I),   # Niveton
    re.compile(r"^Presentation.*30\.04\.2026.*\.pdf$", re.I),  # GWM
    re.compile(r".*Presentation.*30\.04\.2026.*\.pdf$", re.I),  # GWM con prefijo
    re.compile(r".*\bppt\s+19\.05\.2026.*\.pdf$", re.I),  # Niveton Dynamic (ppt como sufijo)
]
EXCLUDE_PATTERNS = [
    re.compile(r"^INFORME DE GESTION", re.I),
    re.compile(r"^IG ", re.I),
    re.compile(r"^FS ", re.I),
    re.compile(r"Factsheet", re.I),
    re.compile(r"OVERDRAWN", re.I),
]

# Para ITAU: las versiones SIN BENCHMARK / S-BM son redundantes; solo mantener CON BM
ITAU_EXCLUDE_PATH_RE = re.compile(r"SIN\s*BM|SIN\s*BENCHMARK", re.I)
ITAU_EXCLUDE_FNAME_RE = re.compile(r"\bS[-\s]?BM\b", re.I)

# Línea de holding: peso ISIN nombre... <métrica numérica o '-' >
LINE_RE = re.compile(
    r"^\s*(\d{1,3}\.\d{1,2})\s+([A-Z]{2}[A-Z0-9]{9}\d)\s+(.+?)\s+[-\d]"
)
TOTAL_RE = re.compile(r"^\s*100\.0?0?\s+TOTAL\b", re.I)

# Columnas que vamos a ponderar (existen tanto en fondos_data como etfs_data)
METRIC_COLS = [
    # asset allocation (escala 0-100)
    "s_aa_cash", "s_aa_fixed_income", "s_aa_equity",
    "s_aa_commodities", "s_aa_otros", "s_aa_derivados",
    # RF index
    "f_ind_ytw", "f_ind_moddur",
    # calidad crediticia
    "f_cacr_aaa", "f_cacr_aa", "f_cacr_a", "f_cacr_bbb",
    "f_cacr_bb", "f_cacr_b", "f_cacr_ccc", "f_cacr_nr",
    # sectores RF
    "f_sec_corp", "f_sec_gov", "f_sec_securitized", "f_sec_deriv",
    "f_sec_mmkt", "f_sec_liab",
    # geo RF
    "f_geo_na", "f_geo_la", "f_geo_uk", "f_geo_eurd", "f_geo_eure",
    "f_geo_mena", "f_geo_australasia", "f_geo_china", "f_geo_japan",
    "f_geo_asiad", "f_geo_asiae", "f_geo_other",
    # RV index
    "e_ind_pe", "e_ind_pb", "e_ind_ps",
    # geo RV
    "e_geo_na", "e_geo_la", "e_geo_uk", "e_geo_eurf", "e_geo_eure",
    "e_geo_mena", "e_geo_australasia", "e_geo_china", "e_geo_jap",
    "e_geo_asiad", "e_geo_asiae", "e_geo_other",
    # sec RV
    "e_sec_tech", "e_sec_fin", "e_sec_hc", "e_sec_consdis",
    "e_sec_conss", "e_sec_ind", "e_sec_en", "e_sec_u",
    "e_sec_comm", "e_sec_mat", "e_sec_re",
    # estilos
    "rv_value", "rv_blend", "rv_growth",
    # commodities
    "c_ti_precious", "c_ti_industrial_metals", "c_ti_agric", "c_ti_energy",
]

# Reglas/thresholds por perfil
THRESHOLDS = {
    "Conservador":  {"rf_min": 70, "rv_max": 30, "rv_min": 0,  "ig_min": 80, "dur_max": 6},
    "Moderado":     {"rf_min": 45, "rv_max": 65, "rv_min": 30, "ig_min": 65, "dur_max": 6},
    "Dinámico":     {"rf_min": 0,  "rv_max": 100,"rv_min": 60, "ig_min": 50, "dur_max": 7},
    "Distributivo": {"rf_min": 50, "rv_max": 50, "rv_min": 0,  "ig_min": 65, "dur_max": 6},
    "Otro":         {"rf_min": 0,  "rv_max": 100,"rv_min": 0,  "ig_min": 0,  "dur_max": 10},
}

# Paleta de acentos (CSS vars) rotativa por cliente
ACCENTS = ["--c-mod", "--c-cons", "--c-din", "--c-otro"]

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s %(message)s",
)
log = logging.getLogger("build")


# ---------------------------------------------------------------------------
# 1. CATÁLOGO unificado
# ---------------------------------------------------------------------------
def _safe(v):
    """NaN-safe getter para floats."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    return v


def load_catalog() -> dict:
    """Une fondos+ETFs en un dict por ISIN con campos del catálogo + métricas."""
    fondos = pd.read_excel(DATA / "fondos.xlsx")
    etfs   = pd.read_excel(DATA / "etfs.xlsx")
    fdata  = pd.read_excel(DATA / "fondos_data.xlsx")
    edata  = pd.read_excel(DATA / "etfs_data.xlsx")

    # Normalizar nombres de columnas que difieren entre fondos y etfs
    # fondos_data tiene s_aa_others, etfs tienen s_aa_otros
    if "s_aa_others" in fdata.columns and "s_aa_otros" not in fdata.columns:
        fdata = fdata.rename(columns={"s_aa_others": "s_aa_otros"})

    catalog: dict[str, dict] = {}

    # FONDOS — key = ISIN (dedup conservando la primera ocurrencia)
    fdata_dedup = fdata.drop_duplicates(subset=["ISIN"], keep="first")
    fdata_by_isin = fdata_dedup.set_index("ISIN", drop=False).to_dict("index")
    for _, row in fondos.iterrows():
        isin = row.get("ISIN")
        if not isinstance(isin, str):
            continue
        entry = {
            "isin": isin,
            "ticker": None,
            "nombre": _safe(row.get("Nombre")) or "",
            "categoria": _safe(row.get("Categoria")) or "",
            "sub_categoria": _safe(row.get("Sub categoria")) or "",
            "asset_class": _safe(row.get("asset_class")) or "",
            "ter": _safe(row.get("s_op_ter")),
            "mgmt_fee": _safe(row.get("s_op_mgmt_fee")),
            "kind": "fondo",
        }
        # Cargar métricas del data
        m = fdata_by_isin.get(isin, {})
        for col in METRIC_COLS:
            entry[col] = _safe(m.get(col))
        # s_aa_* viven en ambos archivos; preferir el del data si trae algo,
        # si no usar el del catálogo
        for col in ("s_aa_cash", "s_aa_fixed_income", "s_aa_equity",
                    "s_aa_commodities", "s_aa_otros", "s_aa_derivados"):
            if entry.get(col) is None:
                entry[col] = _safe(row.get(col))
        catalog[isin] = entry

    # ETFs — key = ISIN (también accesibles por ticker para fallback)
    edata_dedup = edata.drop_duplicates(subset=["ticker"], keep="first")
    edata_by_ticker = edata_dedup.set_index("ticker", drop=False).to_dict("index")
    by_ticker = {}
    for _, row in etfs.iterrows():
        isin = row.get("ISIN")
        ticker = row.get("ticker")
        if not isinstance(isin, str):
            continue
        entry = {
            "isin": isin,
            "ticker": ticker if isinstance(ticker, str) else None,
            "nombre": _safe(row.get("Nombre")) or "",
            "categoria": _safe(row.get("Categoria")) or "",
            "sub_categoria": _safe(row.get("Sub categoria")) or "",
            "asset_class": _safe(row.get("asset_class")) or "",
            "ter": _safe(row.get("s_op_ter")),
            "mgmt_fee": None,
            "kind": "etf",
        }
        m = edata_by_ticker.get(ticker, {})
        for col in METRIC_COLS:
            entry[col] = _safe(m.get(col))
        for col in ("s_aa_cash", "s_aa_fixed_income", "s_aa_equity",
                    "s_aa_commodities", "s_aa_otros", "s_aa_derivados"):
            if entry.get(col) is None:
                entry[col] = _safe(row.get(col))
        catalog[isin] = entry
        if isinstance(ticker, str):
            by_ticker[ticker] = entry

    log.info("Catálogo: %d ISINs (fondos+ETFs) cargados", len(catalog))
    return catalog


# ---------------------------------------------------------------------------
# 2. DESCUBRIR estrategias
# ---------------------------------------------------------------------------
def normalize(s: str) -> str:
    """Texto → ASCII bajo, espacios colapsados."""
    if not s:
        return ""
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip()


def infer_perfil(name: str) -> str:
    n = normalize(name).lower()
    # heurísticas en orden de preferencia
    if "1mm" in n or "platinum" in n or "utmost" in n or "evolution" in n and "dynamic" not in n:
        # Platinum/Evolution/Utmost suelen ser perfiles moderados-altos genéricos
        # caen en "Otro" salvo que el nombre indique perfil
        pass
    if any(w in n for w in ["conservador", "conservadora", "conservative", "cons "]):
        return "Conservador"
    if any(w in n for w in ["dinamica", "dinamico", "dynamic", "crecimiento", "growth"]):
        return "Dinámico"
    if any(w in n for w in ["distributiv", "distribucion"]):
        return "Distributivo"
    if any(w in n for w in ["moderada", "moderado", "moderate", "balance", "balanceada"]):
        return "Moderado"
    if "100rf" in n or "100 rf" in n:
        return "Conservador"
    if "100rv" in n or "100 rv" in n:
        return "Dinámico"
    return "Otro"


def clean_strategy_name(fname: str, cliente: str) -> str:
    s = fname
    # quitar extensión
    s = re.sub(r"\.pdf$", "", s, flags=re.I)
    # quitar fechas
    s = re.sub(r"\s*30[-.]?4[-.]?2026\s*", " ", s)
    s = re.sub(r"\s*30\.04\.2026\s*", " ", s)
    s = re.sub(r"\s*19\.05\.2026\s*", " ", s)
    # quitar prefijos típicos
    s = re.sub(r"^PPT\s+", "", s, flags=re.I)
    s = re.sub(r"^Presentation\s+", "", s, flags=re.I)
    s = re.sub(r"\s+Presentation$", "", s, flags=re.I)
    s = re.sub(r"\s+PPT$", "", s, flags=re.I)
    s = re.sub(r"\s+ppt$", "", s)
    s = re.sub(r"\s+ppt\s*$", "", s, flags=re.I)
    # ITAU: sacar el sufijo/prefijo de benchmark del nombre visible
    s = re.sub(r"\bC[-\s]?BM\b", "", s, flags=re.I)
    s = re.sub(r"\bCON\s+BENCHMARK\b", "", s, flags=re.I)
    s = re.sub(r"\bCON\s+BM\b", "", s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip()
    # algunos clientes prefijan con su nombre, lo dejamos
    return s or cliente


def discover_strategies() -> list[dict]:
    """Walk recursivo y filtro de PPTs válidos."""
    out: list[dict] = []
    if not PDF_ROOT.exists():
        log.error("PDF_ROOT no existe: %s", PDF_ROOT)
        return out

    for dirpath, dirnames, filenames in os.walk(PDF_ROOT):
        # ignorar Historial y similares
        if "Historial" in dirpath:
            continue
        for f in filenames:
            if any(p.search(f) for p in EXCLUDE_PATTERNS):
                continue
            if not any(p.match(f) for p in PPT_PATTERNS):
                continue
            full = Path(dirpath) / f
            # cliente: primera carpeta bajo PDF_ROOT
            rel = full.relative_to(PDF_ROOT)
            parts = rel.parts
            cliente_raw = parts[0]
            # Filtro ITAU: las versiones SIN BENCHMARK / S-BM son duplicados
            # redundantes (misma estrategia sin métricas vs benchmark).
            if "ITAU" in cliente_raw.upper() or "ITAÚ" in cliente_raw.upper():
                if ITAU_EXCLUDE_PATH_RE.search(str(rel)) or ITAU_EXCLUDE_FNAME_RE.search(f):
                    continue
            # limpiar nombre de cliente
            cliente = cliente_raw
            cliente = re.sub(r"\s*-\s*MaximUs\s*$", "", cliente, flags=re.I)
            cliente = re.sub(r"\s*\(.*\)\s*$", "", cliente).strip()
            # mapeos especiales
            cliente_map = {
                "FS": "FS Internacional",
                "FS Internacional": "FS Internacional",
                "Latam Consultus": "LATAM ConsultUs",
                "Latin Securities": "Latin Securities",
                "Max Valores": "Max Valores",
                "Andres Picovsky": "Andrés Picovsky",
                "Monica y Sofia": "Mónica y Sofía",
            }
            cliente = cliente_map.get(cliente, cliente)
            # estrategia: nombre limpio del filename
            estrategia = clean_strategy_name(f, cliente)
            # sub-segmento (ej. EFG/MAX bajo Max Valores; OnShore/UCITS bajo Mónica)
            sub = None
            if len(parts) >= 3 and "FS fondos separados" not in parts:
                sub_candidate = parts[1]
                if not sub_candidate.lower().startswith("13."):
                    sub = sub_candidate
            # decidir perfil
            perfil = infer_perfil(estrategia)
            out.append({
                "cliente": cliente,
                "sub": sub,
                "estrategia": estrategia,
                "perfil": perfil,
                "path": str(full),
                "filename": f,
            })
    # Dedup: si para misma cliente+estrategia hay varias (ej. 13.05.2026 vs 30-4-2026),
    # priorizar el de fin de mes (30-4-2026) que ya filtramos arriba
    seen = {}
    for s in out:
        key = (s["cliente"], s["sub"], s["estrategia"].lower())
        if key not in seen:
            seen[key] = s
    return list(seen.values())


# ---------------------------------------------------------------------------
# 3. PARSEAR PDFs — composición de la estrategia
# ---------------------------------------------------------------------------
def parse_pdf_holdings(pdf_path: str) -> tuple[list[tuple], dict]:
    """Devuelve (lista de (isin, peso, nombre, ret_3y, ret_5y)) + metadata de las páginas
    relevantes (RF/RV/perf declarados). ret_3y / ret_5y pueden ser None."""
    holdings: list[tuple] = []
    meta: dict = {}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 1) Página 3 (idx=2) — composición
            target_pages = [2, 3, 1, 4]
            for pi in target_pages:
                if pi >= len(pdf.pages):
                    continue
                txt = pdf.pages[pi].extract_text() or ""
                if "ISIN" not in txt and "ISIN" not in txt.upper():
                    continue
                if "COMPOSI" not in txt.upper() and "COMPOSITION" not in txt.upper():
                    continue
                # parsear líneas
                hs = _parse_holdings_text(txt)
                if hs:
                    holdings = hs
                    break
            # 2) Páginas siguientes — perf declarado y métricas BM
            # Buscar página con "RENDIMIENTO" o "PERFORMANCE"
            for pi in range(len(pdf.pages)):
                txt = pdf.pages[pi].extract_text() or ""
                up = txt.upper()
                if "RENDIMIENTO Y M" in up or "PERFORMANCE" in up and "BENCHMARK" in up:
                    meta["perf_text"] = txt
                    break
            # 3) Página de RF declarado
            for pi in range(len(pdf.pages)):
                txt = pdf.pages[pi].extract_text() or ""
                up = txt.upper()
                if "YTW" in up and "DURACI" in up or "YTW" in up and "DURATION" in up:
                    meta["rf_text"] = txt
                    break
    except Exception as e:
        log.warning("Error leyendo %s: %s", pdf_path, e)
    return holdings, meta


def _parse_holdings_text(txt: str) -> list[tuple]:
    """Aplica LINE_RE línea por línea sobre el texto de la página de composición.
    Además extrae los floats trailing (Ret_1m, Ret_YTD, Ret_1y, Ret_3y, Ret_5y, ...)
    para devolver también Ret_3y y Ret_5y."""
    out: list[tuple] = []
    nums_re = re.compile(r'-?\d+\.\d+')
    # Algunos nombres ocupan dos líneas (el filename del fondo) — pdfplumber suele
    # poner todo en una sola línea, pero a veces parte el "Nombre del fondo".
    # Estrategia: regex directo por línea, ya filtra por '-' o dígito tras nombre.
    for line in txt.split("\n"):
        m = LINE_RE.match(line)
        if not m:
            continue
        peso = float(m.group(1))
        isin = m.group(2)
        nombre = m.group(3).strip()
        # Trailing floats: tomamos todos los floats de la línea, descartamos el
        # primero (peso) y nos quedamos con los métricos. Índices 0..4 son
        # Ret_1m, Ret_YTD, Ret_1y, Ret_3y, Ret_5y.
        all_nums = nums_re.findall(line)
        ret_3y = None
        ret_5y = None
        # Removemos la primera ocurrencia (peso) — todas las demás son métricas
        # (los Ret pueden ser negativos, con punto decimal).
        metrics = all_nums[1:] if all_nums else []
        if len(metrics) >= 5:
            try:
                ret_3y = float(metrics[3])
            except ValueError:
                pass
            try:
                ret_5y = float(metrics[4])
            except ValueError:
                pass
        out.append((isin, peso, nombre, ret_3y, ret_5y))
    return out


# ---------------------------------------------------------------------------
# 3b. PARSEAR PDFs — métricas de performance (página 7 típicamente)
# ---------------------------------------------------------------------------
_FLOAT_RE = re.compile(r'-?\d+\.\d+')
# Detección flexible de línea de Benchmark: cubre nombres como
# "Benchmark", "Índice", "Index", "Bench", token "BM" aislado o
# patrones tipo "60-40", "70-30", "35-55-10".
_BM_PATTERN = re.compile(r"""
    \b(?:benchmark|índice|indice|index|bench)\b   |   # palabra clave
    \bBM\b                                         |   # BM aislado
    \b\d{2,3}-\d{2,3}(?:-\d{2,3})?\b                   # "60-40", "70-30", "35-55-10" SIN espacios
                                                        # (evita falsos positivos en líneas de solo números)
""", re.IGNORECASE | re.VERBOSE)


def _is_benchmark_line(line: str) -> bool:
    return bool(_BM_PATTERN.search(line))


def _extract_perf_from_text(txt: str) -> dict:
    """Intenta extraer perf+riesgo de un texto. Si no hay líneas con 7 y 10
    floats no-BM, devuelve None (señal para seguir buscando).

    También extrae los valores del benchmark (líneas que matchean BM) con
    el mismo formato (7 o 10 floats), si están presentes."""
    perf = {"ret_5y": None, "ret_3y": None, "ret_1y": None}
    risk = {"vol_5y": None, "sharpe_5y": None,
            "dd_5y": None, "dd_3y": None,
            "up_5y": None, "dn_5y": None}
    bm_perf = {"bm_ret_5y": None, "bm_ret_3y": None, "bm_ret_1y": None}
    bm_risk = {"bm_vol_5y": None, "bm_sharpe_5y": None,
               "bm_dd_5y": None, "bm_dd_3y": None,
               "bm_up_5y": None, "bm_dn_5y": None}

    for line in txt.split("\n"):
        is_bm = _is_benchmark_line(line)
        nums = _FLOAT_RE.findall(line)
        if is_bm:
            if len(nums) == 7 and bm_perf["bm_ret_5y"] is None:
                try:
                    fnums = [float(x) for x in nums]
                    bm_perf["bm_ret_1y"] = fnums[4]
                    bm_perf["bm_ret_3y"] = fnums[5]
                    bm_perf["bm_ret_5y"] = fnums[6]
                except ValueError:
                    pass
            elif len(nums) == 10 and bm_risk["bm_vol_5y"] is None:
                try:
                    fnums = [float(x) for x in nums]
                    bm_risk["bm_vol_5y"] = fnums[1]
                    bm_risk["bm_sharpe_5y"] = fnums[3]
                    bm_risk["bm_dd_3y"] = fnums[4]
                    bm_risk["bm_dd_5y"] = fnums[5]
                    bm_risk["bm_up_5y"] = fnums[7]
                    bm_risk["bm_dn_5y"] = fnums[9]
                except ValueError:
                    pass
            elif len(nums) == 6 and bm_risk["bm_vol_5y"] is None and re.search(r'-\s+-\s+-\s+-\s*$', line):
                try:
                    fnums = [float(x) for x in nums]
                    bm_risk["bm_vol_5y"] = fnums[1]
                    bm_risk["bm_sharpe_5y"] = fnums[3]
                    bm_risk["bm_dd_3y"] = fnums[4]
                    bm_risk["bm_dd_5y"] = fnums[5]
                except ValueError:
                    pass
            continue
        if len(nums) == 7 and perf["ret_5y"] is None:
            try:
                fnums = [float(x) for x in nums]
                perf["ret_1y"] = fnums[4]
                perf["ret_3y"] = fnums[5]
                perf["ret_5y"] = fnums[6]
            except ValueError:
                pass
        elif len(nums) == 10 and risk["vol_5y"] is None:
            try:
                fnums = [float(x) for x in nums]
                risk["vol_5y"] = fnums[1]
                risk["sharpe_5y"] = fnums[3]
                risk["dd_3y"] = fnums[4]
                risk["dd_5y"] = fnums[5]
                risk["up_5y"] = fnums[7]
                risk["dn_5y"] = fnums[9]
            except ValueError:
                pass
        elif len(nums) == 6 and risk["vol_5y"] is None and re.search(r'-\s+-\s+-\s+-\s*$', line):
            # Caso ITAU S-BM / Niveton: la tabla tiene Vol3 Vol5 Sharpe3 Sharpe5 DD3 DD5
            # y las columnas Up/Dn vienen con "- - - -" (sin tracking suficiente).
            try:
                fnums = [float(x) for x in nums]
                risk["vol_5y"] = fnums[1]
                risk["sharpe_5y"] = fnums[3]
                risk["dd_3y"] = fnums[4]
                risk["dd_5y"] = fnums[5]
                # up/dn quedan en None
            except ValueError:
                pass
        # No rompemos cuando termina la estrategia: seguimos buscando líneas BM
    return {**perf, **risk, **bm_perf, **bm_risk}


def parse_performance_metrics(pdf_path: str) -> dict:
    """Extrae métricas de performance y riesgo barriendo todas las páginas del PDF.
    Empieza por las páginas más probables (7, luego 6, 8, 5, 9, ...) y devuelve
    la primera donde el parser encuentre datos completos (perf + risk).

    Devuelve dict con keys: ret_5y, ret_3y, ret_1y, vol_5y, sharpe_5y,
    dd_5y, dd_3y, up_5y, dn_5y + las versiones bm_* del benchmark.
    """
    empty = {
        "ret_5y": None, "ret_3y": None, "ret_1y": None,
        "vol_5y": None, "sharpe_5y": None,
        "dd_5y": None, "dd_3y": None,
        "up_5y": None, "dn_5y": None,
        "bm_ret_5y": None, "bm_ret_3y": None, "bm_ret_1y": None,
        "bm_vol_5y": None, "bm_sharpe_5y": None,
        "bm_dd_5y": None, "bm_dd_3y": None,
        "bm_up_5y": None, "bm_dn_5y": None,
    }
    try:
        with pdfplumber.open(pdf_path) as pdf:
            n = len(pdf.pages)
            # Orden de páginas a probar: 7 (idx 6), 6, 8, 5, 9, 4, 10, ...
            order = [6]
            for offset in range(1, 8):
                if 6 - offset >= 0:
                    order.append(6 - offset)
                if 6 + offset < n:
                    order.append(6 + offset)

            best = None
            for idx in order:
                if idx >= n or idx < 0:
                    continue
                txt = pdf.pages[idx].extract_text() or ""
                if not txt:
                    continue
                # Heurística: la página debe contener al menos una palabra clave
                # del header de la tabla de métricas
                low = txt.lower()
                if not any(kw in low for kw in (
                    'rendimiento', 'performance', 'volatilidad', 'volatility',
                    'sharpe', 'drawdown', 'capture',
                )):
                    continue
                res = _extract_perf_from_text(txt)
                # Datos completos (perf + risk) → devolvemos directo
                if res.get("ret_5y") is not None and res.get("vol_5y") is not None:
                    return res
                # Si solo hay perf parcial, lo guardamos como mejor candidato
                if res.get("ret_5y") is not None:
                    score = sum(v is not None for v in res.values())
                    if best is None or score > sum(v is not None for v in best.values()):
                        best = res
            if best is not None:
                return best
    except Exception as e:
        log.warning("Error leyendo perf de %s: %s", pdf_path, e)
    return empty


# ---------------------------------------------------------------------------
# 4. COMPOSICIÓN PONDERADA
# ---------------------------------------------------------------------------
def aggregate(holdings: list[tuple], catalog: dict) -> dict:
    """Pondera todas las métricas por peso. Hay dos clases de campos:
      - Asset allocation (s_aa_*): % del fondo en cada bucket → se pondera por peso/100.
      - Índices RF (f_ind_*, f_cacr_*, f_sec_*, f_geo_*): % dentro del bucket RF del fondo →
        peso efectivo = peso * s_aa_fixed_income/100, luego normalizar por la suma de RF.
      - Índices RV (e_ind_*, e_geo_*, e_sec_*, rv_*): igual con s_aa_equity.
      - Commodities (c_ti_*): igual con s_aa_commodities.
    """
    res: dict = {col: 0.0 for col in METRIC_COLS}
    res["_weight_known"] = 0.0
    res["_weight_total"] = 0.0
    res["_missing_isins"] = []
    res["_holdings"] = []
    res["_ter_w"] = 0.0
    res["_ter_known"] = 0.0

    if not holdings:
        return res

    total_w = sum(h[1] for h in holdings)
    res["_weight_total"] = total_w

    # Acumuladores con su peso efectivo (para luego normalizar)
    rf_sum: dict[str, float] = defaultdict(float)
    rf_w_total = 0.0
    rv_sum: dict[str, float] = defaultdict(float)
    rv_w_total = 0.0
    co_sum: dict[str, float] = defaultdict(float)
    co_w_total = 0.0

    sub_acc: dict[str, float] = defaultdict(float)

    AA_COLS = {"s_aa_cash", "s_aa_fixed_income", "s_aa_equity",
               "s_aa_commodities", "s_aa_otros", "s_aa_derivados"}
    RF_COLS = {c for c in METRIC_COLS
               if c.startswith(("f_ind_", "f_cacr_", "f_sec_", "f_geo_"))}
    RV_COLS = {c for c in METRIC_COLS
               if c.startswith(("e_ind_", "e_geo_", "e_sec_", "rv_"))}
    CO_COLS = {c for c in METRIC_COLS if c.startswith("c_ti_")}

    for h in holdings:
        # Soportar tanto (isin, peso, nombre) como (isin, peso, nombre, ret3y, ret5y)
        if len(h) >= 5:
            isin, peso, nombre, ret_3y, ret_5y = h[0], h[1], h[2], h[3], h[4]
        else:
            isin, peso, nombre = h[0], h[1], h[2]
            ret_3y, ret_5y = None, None
        fund = catalog.get(isin)
        if not fund:
            res["_missing_isins"].append((isin, peso, nombre))
            continue
        res["_weight_known"] += peso

        rf_share = (fund.get("s_aa_fixed_income") or 0) / 100.0
        rv_share = (fund.get("s_aa_equity") or 0) / 100.0
        co_share = (fund.get("s_aa_commodities") or 0) / 100.0

        rf_eff = peso * rf_share
        rv_eff = peso * rv_share
        co_eff = peso * co_share

        # Asset allocation — ponderación lineal directa
        for col in AA_COLS:
            v = fund.get(col)
            if v is not None:
                res[col] += float(v) * peso / 100.0
        # RF metrics
        if rf_eff > 0:
            rf_w_total += rf_eff
            for col in RF_COLS:
                v = fund.get(col)
                if v is not None:
                    rf_sum[col] += float(v) * rf_eff
        # RV metrics
        if rv_eff > 0:
            rv_w_total += rv_eff
            for col in RV_COLS:
                v = fund.get(col)
                if v is not None:
                    rv_sum[col] += float(v) * rv_eff
        # Commodities
        if co_eff > 0:
            co_w_total += co_eff
            for col in CO_COLS:
                v = fund.get(col)
                if v is not None:
                    co_sum[col] += float(v) * co_eff

        ter = fund.get("ter")
        if ter is not None:
            res["_ter_w"] += float(ter) * peso / 100.0
            res["_ter_known"] += peso
        sub = fund.get("sub_categoria") or fund.get("categoria") or "?"
        sub_acc[sub] += peso
        res["_holdings"].append({
            "isin": isin,
            "peso": peso,
            "nombre": fund.get("nombre") or nombre,
            "sub": sub,
            "categoria": fund.get("categoria"),
            "asset_class": fund.get("asset_class"),
            "ret_3y": ret_3y,
            "ret_5y": ret_5y,
        })

    # Normalizar RF / RV / CO por su peso total efectivo
    if rf_w_total > 0:
        for col in RF_COLS:
            res[col] = rf_sum.get(col, 0.0) / rf_w_total
    if rv_w_total > 0:
        for col in RV_COLS:
            res[col] = rv_sum.get(col, 0.0) / rv_w_total
    if co_w_total > 0:
        for col in CO_COLS:
            res[col] = co_sum.get(col, 0.0) / co_w_total

    res["_sub_acc"] = dict(sub_acc)
    if res["_ter_known"] > 0:
        res["ter"] = res["_ter_w"] * 100.0 / max(res["_ter_known"], 1)
    else:
        res["ter"] = None
    return res


# ---------------------------------------------------------------------------
# 5. REGLAS vs HOUSE VIEW
# ---------------------------------------------------------------------------
def analyze(agg: dict, perfil: str) -> dict:
    """Aplica reglas y devuelve findings + veredicto."""
    th = THRESHOLDS.get(perfil, THRESHOLDS["Otro"])

    bad: list[str] = []
    good: list[str] = []

    rf  = agg.get("s_aa_fixed_income", 0) or 0
    rv  = agg.get("s_aa_equity", 0) or 0
    cash = agg.get("s_aa_cash", 0) or 0
    com  = agg.get("s_aa_commodities", 0) or 0
    dur  = agg.get("f_ind_moddur", 0) or 0
    ytw  = agg.get("f_ind_ytw", 0) or 0
    aaa = agg.get("f_cacr_aaa", 0) or 0
    aa  = agg.get("f_cacr_aa", 0) or 0
    a   = agg.get("f_cacr_a", 0) or 0
    bbb = agg.get("f_cacr_bbb", 0) or 0
    bb  = agg.get("f_cacr_bb", 0) or 0
    b   = agg.get("f_cacr_b", 0) or 0
    ccc = agg.get("f_cacr_ccc", 0) or 0
    nr  = agg.get("f_cacr_nr", 0) or 0
    # IG está expresado como % del bucket RF — lo dejamos en su propia escala
    ig_rf = aaa + aa + a + bbb            # % dentro de la cartera de RF
    sub_ig_rf = bb + b + ccc + nr
    # %tech+industrials del RV total (los e_sec_* son % de la cartera RV)
    tech = agg.get("e_sec_tech", 0) or 0
    ind  = agg.get("e_sec_ind", 0) or 0
    eq_em = (agg.get("e_geo_la", 0) or 0) + (agg.get("e_geo_eure", 0) or 0) + \
            (agg.get("e_geo_china", 0) or 0) + (agg.get("e_geo_asiae", 0) or 0)
    eq_europe = (agg.get("e_geo_eurf", 0) or 0) + (agg.get("e_geo_uk", 0) or 0)
    eq_na = agg.get("e_geo_na", 0) or 0
    non_us_rv = max(0.0, 100.0 - eq_na) if rv > 0 else 0
    ter = agg.get("ter")

    # R1 calidad crediticia
    if sub_ig_rf > (100 - th["ig_min"]) and rf > 5:
        bad.append(
            f"<strong>Calidad crediticia agresiva</strong> · {sub_ig_rf:.1f}% sub-IG en RF — "
            f"reducir concentración sub-investment grade para perfil {perfil}.")
    # R2 duration larga
    if dur > th["dur_max"] and rf > 5:
        bad.append(
            f"<strong>Duration {dur:.2f}a por encima del techo</strong> ({th['dur_max']}a) — "
            f"house view es UW duration larga US por riesgo de tasas.")
    # R3 RV alineada con OW
    if perfil != "Conservador" and perfil != "Otro" and rv < th["rv_min"]:
        bad.append(
            f"<strong>UW equities</strong> ({rv:.1f}% RV) vs piso {th['rv_min']}% del perfil — "
            f"house view OW global equities.")
    # R4 EM equity
    if perfil not in ("Conservador", "Otro") and eq_em < 3 and rv > 10:
        bad.append("<strong>Sin EM equities</strong> — house view OW EM.")
    # R5 commodities/oro
    if com < 1 and rv > 10:
        bad.append("<strong>Sin commodities/oro</strong> — falta hedge inflacionario (OW en house view).")
    # R6/R7 removidas — IA-infra es posición táctica, TER no es criterio de revisión
    # R8 Europa pura
    if eq_europe > 15:
        bad.append(
            f"<strong>Sobre-expuesto Europa</strong> · {eq_europe:.1f}% RV en Europa — "
            f"house view Neutral, redirigir a EM o Global.")
    # R9 cash exceso
    if cash > 10:
        bad.append(
            f"<strong>Cash elevado</strong> · {cash:.1f}% — "
            f"house view UW cash, preferir carry en RF corta.")
    # R10 solapamiento por sub-categoría
    sub_acc = agg.get("_sub_acc", {})
    overlaps = [(k, v) for k, v in sub_acc.items() if v > 25 and k not in ("", "?")]
    overlaps.sort(key=lambda x: -x[1])
    if overlaps:
        k, v = overlaps[0]
        bad.append(
            f"<strong>Concentración en una sub-categoría</strong> · {k} suma {v:.1f}% — "
            f"consolidar fondos con misma sub-categoría para reducir solapamiento.")

    # GOOD findings
    if ytw > 5 and rf > 5:
        good.append(f"<strong>YTW atractivo</strong> · {ytw:.2f}% en RF — captura el carry que el view marca.")
    if rf > 5 and 3 <= dur <= 6:
        good.append(f"<strong>Duration controlada</strong> · {dur:.2f}a — coherente con UW duration larga US.")
    if rv > 10 and non_us_rv > 25:
        good.append(f"<strong>RV con diversificación geográfica</strong> · {non_us_rv:.1f}% non-US.")
    if rf > 5 and ig_rf >= th["ig_min"]:
        good.append(f"<strong>Calidad crediticia coherente</strong> · IG {ig_rf:.1f}% ≥ {th['ig_min']}%.")
    if ter is not None and ter < 0.8:
        good.append(f"<strong>TER eficiente</strong> · {ter:.2f}%.")

    # Veredicto
    n_bad = len(bad)
    if n_bad >= 4:
        verdict = "Crítico"
        vd_class = "cr"
    elif n_bad >= 2:
        verdict = "Atención"
        vd_class = "wn"
    else:
        verdict = "OK"
        vd_class = "ok"

    return {
        "bad": bad,
        "good": good,
        "verdict": verdict,
        "vd_class": vd_class,
        "rf": rf, "rv": rv, "cash": cash, "com": com,
        "dur": dur, "ytw": ytw,
        "ig_rf": ig_rf, "sub_ig_rf": sub_ig_rf,
        "eq_em": eq_em, "eq_europe": eq_europe, "eq_na": eq_na,
        "tech": tech, "ind": ind,
        "ter": ter,
        "non_us_rv": non_us_rv,
    }


# ---------------------------------------------------------------------------
# 5b. RECOMENDACIONES — cambios sugeridos concretos
# ---------------------------------------------------------------------------
# Categorías que consideramos HY / sub-IG agresivos
HY_CATEGORIES = {"High Yield"}
EM_BOND_CATEGORIES = {"EM Bonds", "EM BONDS", "Asian Bonds"}
FLEX_FI_CATEGORIES = {"Flexible FI"}
EUROPE_EQ_CATEGORIES = {"Europe Equity"}
CORE_FI_CATEGORIES = {"Core Fixed Income", "Short Duration"}

# Universo cerrado: SOLO fondos de la Lista de Fondos Recomendados (TRL)
# de LATAM ConsultUs · Mayo 2026. Las sugerencias de "Sumar" deben provenir
# exclusivamente de esta lista. Categorías ausentes (oro, commodities, TIPS,
# real assets / infra) NO generan sugerencias en la tabla — el problema sigue
# apareciendo como observación en "Problemas detectados".
TRL_FUNDS = {
    # IG short duration / corporate corto plazo
    'ig_short': [
        ('LU1882441907', 'Amundi Funds - US Short-Term Bond'),
        ('IE00B51H7M53', 'Muzinich Funds - EnhancedYield Short-Term Fund'),
        ('LU2133069521', 'Vontobel - TwentyFour Absolute Return Credit Fund'),
    ],
    # EM Equity (incluye Asia ex Japan, blend EM)
    'em_equity': [
        ('LU0181495838', 'Schroder ISF Emerging Asia'),
    ],
    # EM Debt (si la lógica pide RF emergente)
    'em_debt': [
        ('IE00BDZRXR46', 'Neuberger Berman Short Duration Emerging Market Debt'),
        ('IE00B986J944', 'Neuberger Berman EM Debt - Hard Currency'),
        ('LU1670632337', 'M&G Lux Emerging Markets Bond Fund'),
        ('LU0611394940', 'Ninety One - Emerging Markets Corporate Debt Fund'),
    ],
    # IA-infra / sectorial Tech (única posición temática que existe en TRL)
    'tech_sector': [
        ('LU1235294995', 'Fidelity Funds - Global Technology Fund'),
    ],
    # Global Equity para subir RV
    'global_equity': [
        ('IE0034235188', 'PineBridge Global Focus Equity Fund'),
        ('LU0557290698', 'Schroder ISF Global Sustainable Growth'),
        ('LU0951559797', 'Robeco Capital Growth - BP Global Premium Equities'),
    ],
    # NO HAY oro/commodities/TIPS/infra real assets en la lista
}


def pick_trl_fund(category: str, current_isins: set[str]) -> str | None:
    """Devuelve string 'Nombre (ISIN)' del primer fondo TRL de la categoría
    que no esté ya en cartera. None si todos ya están o no hay candidatos."""
    for isin, name in TRL_FUNDS.get(category, []):
        if isin not in current_isins:
            return f"{name} ({isin})"
    return None


def _short_name(name: str, n: int = 70) -> str:
    """Trim de nombre de fondo para tablas (default 70 chars; el CSS wrap se encarga)."""
    if not name:
        return "—"
    name = re.sub(r"\s+", " ", str(name)).strip()
    if len(name) <= n:
        return name
    return name[: n - 1].rstrip() + "…"


def find_funds_by_categoria(holdings_detail: list[dict], catalog: dict,
                            categorias: set[str]) -> list[dict]:
    """Holdings (vista enriquecida en agg['_holdings']) cuya `categoria` está en el set."""
    out = []
    for h in holdings_detail:
        cat = (h.get("categoria") or "")
        if cat in categorias:
            isin = h["isin"]
            ter = catalog.get(isin, {}).get("ter")
            out.append({**h, "ter": ter})
    return out


def find_funds_subig(holdings_detail: list[dict], catalog: dict,
                     min_subig: float = 40.0) -> list[dict]:
    """Holdings con alto % sub-IG en su composición (HY + EM debt + flex FI sub-IG)."""
    out = []
    for h in holdings_detail:
        isin = h["isin"]
        cat = (h.get("categoria") or "")
        fund = catalog.get(isin, {})
        bb = float(fund.get("f_cacr_bb") or 0)
        b = float(fund.get("f_cacr_b") or 0)
        ccc = float(fund.get("f_cacr_ccc") or 0)
        nr = float(fund.get("f_cacr_nr") or 0)
        subig = bb + b + ccc + nr
        # HY puro siempre cuenta; otros solo si la métrica supera el threshold
        if cat in HY_CATEGORIES or subig >= min_subig:
            out.append({**h, "subig_pct": subig, "ter": fund.get("ter")})
    # ordenar por peso descendente
    out.sort(key=lambda x: -x["peso"])
    return out


def find_long_duration_funds(holdings_detail: list[dict], catalog: dict,
                             min_dur: float = 7.0) -> list[dict]:
    """Fondos con duration > min_dur y exposición RF significativa."""
    out = []
    for h in holdings_detail:
        isin = h["isin"]
        fund = catalog.get(isin, {})
        dur = fund.get("f_ind_moddur")
        rf_share = fund.get("s_aa_fixed_income") or 0
        if dur is None or rf_share < 30:
            continue
        if float(dur) > min_dur:
            out.append({**h, "dur": float(dur), "ter": fund.get("ter")})
    out.sort(key=lambda x: -x["peso"])
    return out


def find_high_ter_funds(holdings_detail: list[dict], catalog: dict,
                       n: int = 2, min_ter: float = 1.4) -> list[dict]:
    """Top N fondos con TER alto (peso * TER como criterio combinado)."""
    scored = []
    for h in holdings_detail:
        isin = h["isin"]
        ter = catalog.get(isin, {}).get("ter")
        if ter is None:
            continue
        ter_v = float(ter)
        if ter_v < min_ter:
            continue
        # score = peso * ter (impacto en el TER ponderado total)
        scored.append({**h, "ter": ter_v, "_score": h["peso"] * ter_v})
    scored.sort(key=lambda x: -x["_score"])
    return scored[:n]


def find_overlapping_sub(holdings_detail: list[dict], min_total: float = 25.0) -> list[dict]:
    """Detecta sub-categorías con 2+ fondos que juntos suman > min_total."""
    by_sub: dict[str, list[dict]] = defaultdict(list)
    for h in holdings_detail:
        sub = h.get("sub") or ""
        if not sub or sub == "?":
            continue
        by_sub[sub].append(h)
    out = []
    for sub, funds in by_sub.items():
        if len(funds) < 2:
            continue
        tot = sum(f["peso"] for f in funds)
        if tot >= min_total:
            funds_sorted = sorted(funds, key=lambda x: -x["peso"])
            out.append({"sub": sub, "total": tot, "funds": funds_sorted})
    out.sort(key=lambda x: -x["total"])
    return out


def _pp(v: str) -> float:
    """Parsea un valor tipo '5%', '10.0%', '' a float (puntos porcentuales)."""
    if not v:
        return 0.0
    try:
        return float(str(v).rstrip('%').strip() or 0)
    except ValueError:
        return 0.0


def _delta_libre(recs: list[dict]) -> float:
    """Suma de puntos liberados por Bajar/Sacar (acción dn)."""
    total = 0.0
    for r in recs:
        if r.get('dir') == 'dn':
            de = _pp(r.get('de', ''))
            a = _pp(r.get('a', ''))
            total += (de - a)
    return total


def _delta_usado(recs: list[dict]) -> float:
    """Suma de puntos agregados por Sumar (acción new)."""
    total = 0.0
    for r in recs:
        if r.get('dir') == 'new':
            de = _pp(r.get('de', ''))
            a = _pp(r.get('a', ''))
            total += (a - de)
    return total


def _balance(recs: list[dict], current_isins: set, catalog: dict, gap: float) -> float:
    """Agrega filas Sumar hasta cubrir el gap. Cada Sumar nuevo es de hasta 10pp.
    Devuelve el gap residual."""
    if gap < 1:
        return gap  # ya balanceado o tolerancia

    # Detectar asset class predominante de las Bajar para priorizar categoría
    bajas_rf_pp = 0.0
    bajas_rv_pp = 0.0
    for r in recs:
        if r.get('dir') != 'dn':
            continue
        de = _pp(r.get('de', ''))
        a = _pp(r.get('a', ''))
        delta = de - a
        # cruzar por nombre del fondo: keywords típicas de RF
        name = (r.get('position') or '').lower()
        if any(k in name for k in ['bond', 'debt', 'credit', 'income',
                                    'yield', 'duration', 'fixed', 'cash']):
            bajas_rf_pp += delta
        else:
            bajas_rv_pp += delta

    # Definir orden de categorías a probar
    if bajas_rf_pp >= bajas_rv_pp:
        prio = ['ig_short', 'em_debt']
    else:
        prio = ['global_equity', 'em_equity']

    razon_map = {
        'ig_short':      "Reasignar liberación a investment grade short duration",
        'em_debt':       "Reasignar liberación a deuda emergente diversificada",
        'global_equity': "Reasignar liberación a renta variable global",
        'em_equity':     "Reasignar liberación a renta variable emergente",
    }

    # Para cada categoría, sumar candidatos TRL hasta cerrar gap
    for cat in prio:
        for isin, name in TRL_FUNDS.get(cat, []):
            if gap < 1:
                return gap
            if isin in current_isins:
                continue
            # ¿ya está sugerido en recs?
            ya_sugerido = any(isin in (r.get('position') or '') for r in recs)
            if ya_sugerido:
                continue
            # Tamaño: min(gap, 10) — al menos 5; si gap < 5 ajustamos último Sumar
            if gap < 5:
                last = next((r for r in reversed(recs) if r.get('dir') == 'new'), None)
                if last:
                    last_a = _pp(last.get('a', ''))
                    last['a'] = f"{last_a + gap:.0f}%"
                    gap = 0
                    return gap
                # Si no hay Sumar previo, igual agregamos esta fila con gap < 5
                size = gap
            else:
                size = min(gap, 10)
            recs.append({
                'dir': 'new',
                'action': 'Sumar',
                'position': f"{name} ({isin})",
                'de': '0%',
                'a': f"{int(round(size))}%",
                'razon': razon_map.get(cat, "Reasignar liberación a fondo TRL"),
            })
            gap -= size
        if gap < 1:
            return gap
    return gap


def _balance_excess(recs: list[dict], holdings_detail: list, agg: dict, gap_neg: float) -> float:
    """Caso inverso: hay más Sumar que Bajar (gap_neg > 0 representa el exceso
    que falta financiar). Agrega filas Bajar para liberar puntos.
    Estrategia: si hay cash material, bajar cash. Si no, bajar el holding más
    grande de la categoría opuesta a las Sumar (si Sumar son RV → bajar RF
    grande; si Sumar son RF → bajar RV grande).
    Devuelve el gap residual."""
    if gap_neg < 1:
        return gap_neg

    # Detectar asset class de las Sumar para elegir contraparte
    sumar_rv_pp = 0.0
    sumar_rf_pp = 0.0
    for r in recs:
        if r.get('dir') != 'new':
            continue
        de = _pp(r.get('de', ''))
        a = _pp(r.get('a', ''))
        delta = a - de
        name = (r.get('position') or '').lower()
        if any(k in name for k in ['equity', 'equities', 'msci', 'stoxx',
                                    'global focus', 'growth', 'value premium',
                                    'global premium', 'technology', 'tech']):
            sumar_rv_pp += delta
        elif any(k in name for k in ['bond', 'debt', 'credit', 'income',
                                      'yield', 'duration', 'fixed', 'short-term',
                                      'absolute return']):
            sumar_rf_pp += delta
        else:
            sumar_rv_pp += delta  # default

    # Helper: ¿está ya sugerido un cambio sobre este ISIN/nombre?
    def _ya_tocado(nombre: str, isin: str) -> bool:
        for r in recs:
            pos = (r.get('position') or '')
            if isin and isin in pos:
                return True
            if nombre and nombre[:30].lower() in pos.lower():
                return True
        return False

    # 1) Bajar cash si hay > 5%
    cash = agg.get('s_aa_cash') or 0
    if cash > 5 and gap_neg >= 1:
        size = min(gap_neg, max(cash - 2, 0))
        if size >= 1:
            recs.append({
                'dir': 'dn',
                'action': 'Bajar',
                'position': 'Cash',
                'de': f"{cash:.1f}%",
                'a': f"{cash - size:.0f}%",
                'razon': "Reasignar cash estructural a las posiciones sugeridas (house view UW cash)",
            })
            gap_neg -= size
            if gap_neg < 1:
                return gap_neg

    # 2) Bajar el holding más grande de la categoría opuesta a las Sumar
    if sumar_rv_pp >= sumar_rf_pp:
        # Sumar son RV → bajar RV grande (rotación intra-RV)
        keywords_target = ['equity', 'equities', 'msci', 'stoxx', 'acwi',
                           's&p', 'sp 500', 'nasdaq']
        razon_bajar = "Liberar espacio dentro del bucket RV para canalizar las posiciones sugeridas"
    else:
        # Sumar son RF → bajar RF grande
        keywords_target = ['bond', 'debt', 'credit', 'income', 'yield',
                           'duration', 'fixed']
        razon_bajar = "Liberar espacio dentro del bucket RF para canalizar las posiciones sugeridas"

    # Ranking de holdings por peso, filtrando por categoría
    candidatos = []
    for h in holdings_detail:
        name = (h.get('nombre') or '').lower()
        if not any(k in name for k in keywords_target):
            continue
        if _ya_tocado(h.get('nombre') or '', h.get('isin') or ''):
            continue
        candidatos.append(h)
    candidatos.sort(key=lambda x: -x['peso'])

    for h in candidatos:
        if gap_neg < 1:
            return gap_neg
        peso = h['peso']
        # No reducir un holding a menos de 3% (mantener significancia)
        max_reducible = max(peso - 3, 0)
        if max_reducible < 1:
            continue
        size = min(gap_neg, max_reducible)
        nuevo = peso - size
        if size <= 5 and nuevo < 2:
            action = 'Sacar'
            nuevo_str = '0%'
            size = peso
        elif nuevo <= 1:
            action = 'Sacar'
            nuevo_str = '0%'
            size = peso
        else:
            action = 'Bajar'
            nuevo_str = f"{nuevo:.0f}%"
        recs.append({
            'dir': 'dn',
            'action': action,
            'position': _short_name(h.get('nombre') or ''),
            'de': f"{peso:.1f}%",
            'a': nuevo_str,
            'razon': razon_bajar,
        })
        gap_neg -= size

    return gap_neg


def recommend(agg: dict, perfil: str, holdings: list[tuple],
              catalog: dict) -> list[dict]:
    """Genera una lista de cambios sugeridos concretos basándose en la
    composición agregada vs house view.

    Cada item: {'dir': 'up|dn|sw|new', 'action': str, 'position': str,
                'de': str, 'a': str, 'razon': str}
    """
    th = THRESHOLDS.get(perfil, THRESHOLDS["Otro"])
    holdings_detail = agg.get("_holdings", []) or []
    if not holdings_detail:
        return []

    # ISINs ya en cartera — para que pick_trl_fund evite sugerir fondos
    # que el cliente ya tiene
    current_isins = {h[0] for h in holdings}

    recs: list[dict] = []
    # Evitar duplicar ajustes sobre la misma posición + dirección
    seen_keys: set[tuple[str, str]] = set()

    def add(dir_, action, position, de, a, razon):
        key = (dir_, _short_name(position, 60).lower())
        if key in seen_keys:
            return
        seen_keys.add(key)
        recs.append({
            "dir": dir_, "action": action, "position": position,
            "de": de, "a": a, "razon": razon,
        })

    # Reusamos las métricas que analyze() ya calcula
    rf  = agg.get("s_aa_fixed_income", 0) or 0
    rv  = agg.get("s_aa_equity", 0) or 0
    cash = agg.get("s_aa_cash", 0) or 0
    dur  = agg.get("f_ind_moddur", 0) or 0
    aaa = agg.get("f_cacr_aaa", 0) or 0
    aa  = agg.get("f_cacr_aa", 0) or 0
    a_q = agg.get("f_cacr_a", 0) or 0
    bbb = agg.get("f_cacr_bbb", 0) or 0
    bb_q = agg.get("f_cacr_bb", 0) or 0
    b_q  = agg.get("f_cacr_b", 0) or 0
    ccc_q = agg.get("f_cacr_ccc", 0) or 0
    nr_q  = agg.get("f_cacr_nr", 0) or 0
    sub_ig_rf = bb_q + b_q + ccc_q + nr_q
    tech = agg.get("e_sec_tech", 0) or 0
    ind  = agg.get("e_sec_ind", 0) or 0
    eq_em = (agg.get("e_geo_la", 0) or 0) + (agg.get("e_geo_eure", 0) or 0) + \
            (agg.get("e_geo_china", 0) or 0) + (agg.get("e_geo_asiae", 0) or 0)
    eq_europe = (agg.get("e_geo_eurf", 0) or 0) + (agg.get("e_geo_uk", 0) or 0)
    # eq_em y eq_europe vienen como % del bucket RV — convertir a % de cartera
    eq_em_port = eq_em * rv / 100.0
    eq_europe_port = eq_europe * rv / 100.0
    ter = agg.get("ter")

    # ---- R1 Calidad crediticia agresiva ---------------------------------
    if sub_ig_rf > (100 - th["ig_min"]) and rf > 5:
        # Identificar HY / sub-IG funds
        agg_funds = find_funds_subig(holdings_detail, catalog, min_subig=40.0)
        for f in agg_funds[:3]:
            peso = f["peso"]
            # Cambio mínimo de 5pp absolutos; si está al 5% o menos -> Sacar
            if peso <= 5.0:
                add(
                    "dn", "Sacar",
                    _short_name(f["nombre"]),
                    f"{peso:.1f}%", "0%",
                    f"Reducir concentración sub-investment grade para perfil {perfil}",
                )
            else:
                target = max(round(peso - 5.0), 0)
                add(
                    "dn", "Bajar",
                    _short_name(f["nombre"]),
                    f"{peso:.1f}%", f"{target}%",
                    f"Reducir concentración sub-investment grade para perfil {perfil}",
                )
        # Sumar IG short duration — solo si hay candidato TRL no presente en cartera
        ig_short_pick = pick_trl_fund('ig_short', current_isins)
        if ig_short_pick:
            add(
                "new", "Sumar",
                ig_short_pick,
                "0%", "5%",
                "Mejorar calidad crediticia de la cartera con investment grade short duration",
            )

    # ---- R2 Duration larga ----------------------------------------------
    if dur > th["dur_max"] and rf > 5:
        long_funds = find_long_duration_funds(holdings_detail, catalog,
                                              min_dur=th["dur_max"])
        for f in long_funds[:2]:
            peso = f["peso"]
            if peso < 5.0:
                # Cambio sub-5pp no vale la pena rebalancear; saltar
                continue
            add(
                "sw", "Rotar",
                _short_name(f["nombre"]),
                f"{peso:.1f}% (dur {f['dur']:.1f}a)",
                "Corporate short-duration",
                "Reducir duration: house view es UW duration larga US por riesgo de tasas",
            )
        if not long_funds:
            add(
                "dn", "Bajar",
                "Duration RF agregada",
                f"{dur:.1f}a", f"≤{th['dur_max']}a",
                "Reducir duration: house view es UW duration larga US por riesgo de tasas",
            )

    # ---- R3 UW equities vs perfil (vs view OW global equities) ----------
    # Disparar si RV está bajo el piso, o si está en la mitad baja del rango
    # del perfil (mientras el view es OW, conviene movernos al midpoint+).
    if perfil not in ("Conservador", "Otro"):
        midpoint = (th["rv_min"] + th["rv_max"]) / 2.0
        # umbral suave: 80% del midpoint (i.e. claramente debajo)
        soft_target = midpoint * 0.85
        if rv < max(th["rv_min"], soft_target):
            target = max(int(round(midpoint)), th["rv_min"])
            # Solo recomendar si el ajuste es de al menos 5 puntos
            if target - rv >= 5.0:
                add(
                    "up", "Subir",
                    "Renta Variable global total",
                    f"{rv:.1f}%", f"{target}%",
                    "Aumentar exposición a renta variable: house view OW global equities",
                )
                # Sugerencia concreta de fondo TRL para canalizar el aumento
                ge_pick = pick_trl_fund('global_equity', current_isins)
                if ge_pick:
                    add(
                        "new", "Sumar",
                        ge_pick,
                        "0%", "5%",
                        "Vehículo concreto de la TRL para canalizar el aumento de RV",
                    )

    # ---- R4 EM ausente --------------------------------------------------
    # Si la estrategia tiene RV > 30% (moderados o más agresivos), priorizar
    # EM Equity. Si no tiene RV o es Conservadora, priorizar EM Debt.
    if perfil not in ("Conservador", "Otro") and eq_em_port < 3 and rv > 10:
        if rv > 30:
            em_pick = pick_trl_fund('em_equity', current_isins)
        else:
            em_pick = pick_trl_fund('em_debt', current_isins)
        if em_pick:
            add(
                "new", "Sumar",
                em_pick,
                f"{eq_em_port:.1f}%", "5%",
                "Sumar exposición a mercados emergentes: house view OW EM",
            )

    # ---- R5 Commodities / gold / energy ---------------------------------
    # No hay candidatos en la TRL para oro / commodities / real assets.
    # El finding sigue apareciendo en "Problemas detectados" (vía analyze)
    # como observación, pero NO genera fila en la tabla de cambios sugeridos.

    # ---- R6/R7 removidas -----------------------------------------------
    # R6 (IA-infraestructura): posición táctica/temática, no ajuste estructural.
    # R7 (TER alto): el costo no es criterio de revisión de cartera.

    # ---- R8 Europa concentrada ------------------------------------------
    if eq_europe_port > 15:
        europe_funds = find_funds_by_categoria(holdings_detail, catalog,
                                              EUROPE_EQ_CATEGORIES)
        if europe_funds:
            europe_funds.sort(key=lambda x: -x["peso"])
            f = europe_funds[0]
            peso = f["peso"]
            if peso <= 5.0:
                add(
                    "dn", "Sacar",
                    _short_name(f["nombre"]),
                    f"{peso:.1f}%", "0%",
                    "Reducir concentración Europa: house view Neutral, redirigir a EM o Global",
                )
            else:
                target = max(round(peso - 5.0), 0)
                add(
                    "dn", "Bajar",
                    _short_name(f["nombre"]),
                    f"{peso:.1f}%", f"{target}%",
                    "Reducir concentración Europa: house view Neutral, redirigir a EM o Global",
                )
        else:
            add(
                "dn", "Bajar",
                "Exposición Europa equities",
                f"{eq_europe_port:.1f}%", "<15%",
                "Reducir concentración Europa: house view Neutral, redirigir a EM o Global",
            )

    # ---- R9 Cash exceso -------------------------------------------------
    if cash > 10:
        # cash > 10, target <5 -> siempre cambio de >5pp; ok
        target = 5
        add(
            "dn", "Bajar",
            "Cash",
            f"{cash:.1f}%", f"{target}%",
            "Reducir cash estructural: house view UW cash, preferir carry en RF corta",
        )

    # ---- R10 Solapamiento por sub-categoría -----------------------------
    overlaps = find_overlapping_sub(holdings_detail, min_total=25.0)
    for ov in overlaps[:2]:
        funds = ov["funds"][:2]
        if len(funds) < 2:
            continue
        pos_label = " + ".join(_short_name(f["nombre"], 22) for f in funds)
        add(
            "sw", "Rotar",
            f"Consolidar {pos_label}",
            f"{ov['total']:.1f}% en «{ov['sub']}»",
            "Un único vehículo",
            "Consolidar fondos con misma sub-categoría para reducir solapamiento",
        )

    # ---- Balanceo de masa: la suma de Bajar/Sacar debe igualar la suma --
    # de Sumar para que la cartera cierre al 100%.
    gap = _delta_libre(recs) - _delta_usado(recs)
    if gap > 0:
        _balance(recs, current_isins, catalog, gap)
    elif gap < 0:
        # Más Sumar que Bajar: necesitamos liberar puntos
        _balance_excess(recs, holdings_detail, agg, -gap)

    # ---- Pruning de "Subir RV global total" redundante ------------------
    # Si todas las Sumar terminaron siendo de categorías RV, la fila
    # "Subir Renta Variable global total" es redundante y la quitamos.
    rv_isins = set()
    for cat_key in ('global_equity', 'em_equity', 'tech_sector'):
        for isin, _name in TRL_FUNDS.get(cat_key, []):
            rv_isins.add(isin)
    sumas_total = sum(1 for r in recs if r['dir'] == 'new')
    sumas_rv = sum(
        1 for r in recs
        if r['dir'] == 'new' and any(
            isin in (r.get('position') or '') for isin in rv_isins
        )
    )
    if sumas_total > 0 and sumas_rv == sumas_total:
        recs = [
            r for r in recs
            if not (r.get('dir') == 'up'
                    and 'Renta Variable' in (r.get('position') or ''))
        ]

    return recs


# ---------------------------------------------------------------------------
# 6. PROCESAR pipeline completo
# ---------------------------------------------------------------------------
def process_all(catalog: dict) -> list[dict]:
    strategies = discover_strategies()
    log.info("Descubiertas %d estrategias", len(strategies))
    out = []
    for s in strategies:
        log.info("Procesando %s | %s", s["cliente"], s["estrategia"])
        holdings, _meta = parse_pdf_holdings(s["path"])
        perf = parse_performance_metrics(s["path"])
        if not holdings:
            log.warning("  Sin holdings detectables en %s", s["filename"])
            agg = aggregate([], catalog)
            ana = {
                "bad": [], "good": [], "verdict": "Pendiente", "vd_class": "pd",
                "rf": None, "rv": None, "cash": None, "com": None,
                "dur": None, "ytw": None, "ig_rf": None, "sub_ig_rf": None,
                "eq_em": None, "eq_europe": None, "eq_na": None,
                "tech": None, "ind": None, "ter": None, "non_us_rv": None,
            }
            recs = []
        else:
            agg = aggregate(holdings, catalog)
            ana = analyze(agg, s["perfil"])
            recs = recommend(agg, s["perfil"], holdings, catalog)
        s["holdings"] = holdings
        s["agg"] = agg
        s["ana"] = ana
        s["perf"] = perf
        s["recommendations"] = recs
        out.append(s)
    return out


# ---------------------------------------------------------------------------
# 7. HTML rendering
# ---------------------------------------------------------------------------
def safe_id(s: str) -> str:
    s = normalize(s).lower()
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s or "x"


def fmt(v, digits=1, suffix=""):
    if v is None:
        return '<span class="neu">—</span>'
    try:
        if pd.isna(v):
            return '<span class="neu">—</span>'
    except Exception:
        pass
    return f"{v:.{digits}f}{suffix}"


def fmt_raw(v, digits=1, suffix=""):
    if v is None:
        return "—"
    try:
        if pd.isna(v):
            return "—"
    except Exception:
        pass
    return f"{v:.{digits}f}{suffix}"


def render_index_cards(by_client: dict) -> str:
    """Cards del índice por cliente."""
    parts = []
    for i, (cliente, strats) in enumerate(by_client.items()):
        accent = ACCENTS[i % len(ACCENTS)]
        anchor = "#" + safe_id(cliente)
        n = len(strats)
        # listado de estrategias (corto)
        nombres = ", ".join(sorted({s["estrategia"] for s in strats}))
        if len(nombres) > 110:
            nombres = nombres[:107] + "…"
        parts.append(f"""
    <a class="idx-card" href="{anchor}" style="--gc:var({accent})">
      <span class="idx-n">{i+1:02d}</span>
      <h3>{html.escape(cliente)}</h3>
      <p>{html.escape(nombres)}</p>
      <div class="idx-foot"><span class="c">{n}<small>estrategias</small></span><span class="go">Ver →</span></div>
    </a>""")
    return "\n".join(parts)


def render_overview_row(s: dict, idx: int) -> str:
    """Una fila del overview + su detail row."""
    ana = s["ana"]
    agg = s["agg"]
    holdings = s["holdings"]
    has_data = bool(holdings)
    det_id = f"d-{safe_id(s['cliente'])}-{idx}"
    vd = ana["verdict"]
    vd_cls = ana["vd_class"]
    perfil = s["perfil"]

    rf = ana.get("rf")
    rv = ana.get("rv")
    ytw = ana.get("ytw")
    dur = ana.get("dur")

    perf = s.get("perf") or {}
    ret_5y = perf.get("ret_5y")
    dd_5y = perf.get("dd_5y")

    def _signed_cell(v, label, digits=2, suffix="%"):
        if v is None:
            return f'<td class="num neu" data-l="{label}">—</td>'
        cls = "pos" if v >= 0 else "neg"
        return f'<td class="num {cls}" data-l="{label}">{fmt(v, digits, suffix)}</td>'

    ret_cell = _signed_cell(ret_5y, "Ret5Y")
    dd_cell = _signed_cell(dd_5y, "DD5Y")

    sub_html = ""
    if s.get("sub"):
        sub_html = f'<span class="ov-sub">{html.escape(s["sub"])}</span>'

    row_cls = "ov-row" if has_data else "ov-row pending"
    attrs = f'data-target="{det_id}"' if has_data else ""

    row = f"""
          <tr class="{row_cls}" {attrs}>
            <td class="ov-name"><span class="chev"></span><span class="ov-nm">{html.escape(s["estrategia"])}</span>{sub_html}</td>
            <td class="ov-cat">{html.escape(perfil)}</td>
            <td class="num" data-l="RF">{fmt(rf)}</td>
            <td class="num" data-l="RV">{fmt(rv)}</td>
            <td class="num" data-l="YTW">{fmt(ytw, 2)}</td>
            <td class="num" data-l="Dur">{fmt(dur, 2)}</td>
            {ret_cell}
            {dd_cell}
            <td class="ctr"><span class="vd {vd_cls}">{html.escape(vd)}</span></td>
          </tr>"""
    if not has_data:
        return row

    # detail
    det = render_detail(s, det_id)
    return row + det


def render_detail(s: dict, det_id: str) -> str:
    ana = s["ana"]
    agg = s["agg"]
    perfil = s["perfil"]

    # Top holdings — con Ret 3Y y Ret 5Y en lugar de Sub-categoría
    def _ret_cell(v):
        if v is None:
            return '<td class="num neu">—</td>'
        cls = "pos" if v >= 0 else "neg"
        return f'<td class="num {cls}">{v:.2f}%</td>'

    hold_rows = ""
    for h in sorted(agg.get("_holdings", []), key=lambda x: -x["peso"])[:15]:
        hold_rows += (
            f'<tr><td class="num">{h["peso"]:.2f}%</td>'
            f'<td>{html.escape(h["nombre"][:64])}</td>'
            f'{_ret_cell(h.get("ret_3y"))}'
            f'{_ret_cell(h.get("ret_5y"))}</tr>'
        )
    missing = agg.get("_missing_isins") or []
    miss_html = ""
    if missing:
        miss_html = (
            f'<div class="block"><span class="lbl">ISINs no encontrados en catálogo</span>'
            f'<ul class="findings fix">' +
            "".join(f"<li><strong>{html.escape(i)}</strong> ({p:.2f}%) — {html.escape(n[:60])}</li>"
                    for i, p, n in missing[:8]) +
            "</ul></div>"
        )

    bad_html = ""
    if ana.get("bad"):
        bad_html = (
            '<div class="block"><span class="lbl">⚠️ Problemas detectados</span>'
            '<ul class="findings bad">' +
            "".join(f"<li>{x}</li>" for x in ana["bad"]) +
            "</ul></div>"
        )
    good_html = ""
    if ana.get("good"):
        good_html = (
            '<div class="block"><span class="lbl">✅ Qué está bien</span>'
            '<ul class="findings good">' +
            "".join(f"<li>{x}</li>" for x in ana["good"]) +
            "</ul></div>"
        )

    # Cambios sugeridos — tabla accionable
    recs = s.get("recommendations") or []
    if recs:
        rec_rows = ""
        for r in recs:
            rec_rows += (
                f'<tr>'
                f'<td class="dir {html.escape(r["dir"])}">{html.escape(r["action"])}</td>'
                f'<td class="act">{html.escape(r["position"])}</td>'
                f'<td class="num">{html.escape(r["de"])}</td>'
                f'<td class="num">{html.escape(r["a"])}</td>'
                f'<td>{html.escape(r["razon"])}</td>'
                f'</tr>'
            )
        changes_html = (
            '<div class="block"><span class="lbl">🔄 Cambios sugeridos</span>'
            '<div class="changes"><table>'
            '<thead><tr><th>Acción</th><th>Posición</th><th>De</th><th>A</th><th>Razón</th></tr></thead>'
            f'<tbody>{rec_rows}</tbody></table></div></div>'
        )
    else:
        changes_html = (
            '<div class="block"><span class="lbl">🔄 Cambios sugeridos</span>'
            '<p class="muted" style="font-size:13px;margin-top:6px">'
            'Sin cambios sugeridos — estrategia bien alineada con view.'
            '</p></div>'
        )

    # Bar de allocation
    rf_w = ana.get("rf") or 0
    rv_w = ana.get("rv") or 0
    co_w = ana.get("com") or 0
    ca_w = ana.get("cash") or 0
    ot_w = (agg.get("s_aa_otros") or 0) + (agg.get("s_aa_derivados") or 0)

    # Leyenda solo con buckets con peso significativo (>= 1%); evita ruido visual de "Commodities 0.0%"
    leg_items = [
        (rf_w, "var(--c-cons)", "Renta Fija"),
        (rv_w, "var(--c-din)",  "Renta Variable"),
        (co_w, "var(--c-otro)", "Commodities"),
        (ot_w, "var(--c-mod)",  "Otros/Alternativos"),
        (ca_w, "var(--g3)",     "Cash"),
    ]
    leg_html = "".join(
        f'<div><i style="background:{color}"></i>{label} <b>{fmt_raw(w)}%</b></div>'
        for w, color, label in leg_items if w >= 1.0
    )

    bar_html = (
        '<div class="alloc-bar">'
        f'<span style="width:{rf_w:.1f}%;background:var(--c-cons)"></span>'
        f'<span style="width:{rv_w:.1f}%;background:var(--c-din)"></span>'
        f'<span style="width:{co_w:.1f}%;background:var(--c-otro)"></span>'
        f'<span style="width:{ot_w:.1f}%;background:var(--c-mod)"></span>'
        f'<span style="width:{ca_w:.1f}%;background:var(--g3)"></span>'
        '</div>'
        f'<div class="alloc-leg">{leg_html}</div>'
    )

    # Auto-show RF block: solo si la estrategia tiene exposición material (>= 5%)
    show_rf = rf_w >= 5
    show_rv = rv_w >= 5
    # Mantener siempre data-grp simple (no full) — el grid del sheet-data
    # acomoda 1 o 2 columnas sin descalibrar el dl interno
    rf_block = ""
    if show_rf:
        rf_block = (
            '<div class="data-grp"><span class="grp-t">Renta Fija</span><dl>'
            f'<div class="dl"><dt>YTW</dt><dd class="strong">{fmt_raw(ana.get("ytw"),2,"%")}</dd></div>'
            f'<div class="dl"><dt>Duration</dt><dd>{fmt_raw(ana.get("dur"),2," a")}</dd></div>'
            f'<div class="dl"><dt>Corporate</dt><dd>{fmt_raw(agg.get("f_sec_corp"),1,"%")}</dd></div>'
            f'<div class="dl"><dt>Government</dt><dd>{fmt_raw(agg.get("f_sec_gov"),1,"%")}</dd></div>'
            f'<div class="dl"><dt>Securitized</dt><dd>{fmt_raw(agg.get("f_sec_securitized"),1,"%")}</dd></div>'
            f'<div class="dl"><dt>IG (AAA-BBB)</dt><dd>{fmt_raw(ana.get("ig_rf"),1,"%")}</dd></div>'
            f'<div class="dl"><dt>Sub-IG (BB-CCC+NR)</dt><dd class="neg">{fmt_raw(ana.get("sub_ig_rf"),1,"%")}</dd></div>'
            '</dl></div>'
        )

    rv_block = ""
    if show_rv:
        rv_block = (
            '<div class="data-grp"><span class="grp-t">Renta Variable</span><dl>'
            f'<div class="dl"><dt>P/E</dt><dd>{fmt_raw(agg.get("e_ind_pe"),1)}</dd></div>'
            f'<div class="dl"><dt>P/B</dt><dd>{fmt_raw(agg.get("e_ind_pb"),2)}</dd></div>'
            f'<div class="dl"><dt>P/S</dt><dd>{fmt_raw(agg.get("e_ind_ps"),2)}</dd></div>'
            f'<div class="dl"><dt>North America</dt><dd>{fmt_raw(ana.get("eq_na"),1,"%")}</dd></div>'
            f'<div class="dl"><dt>Europa</dt><dd>{fmt_raw(ana.get("eq_europe"),1,"%")}</dd></div>'
            f'<div class="dl"><dt>Emerging Markets</dt><dd>{fmt_raw(ana.get("eq_em"),1,"%")}</dd></div>'
            f'<div class="dl"><dt>Growth</dt><dd>{fmt_raw(agg.get("rv_growth"),1,"%")}</dd></div>'
            f'<div class="dl"><dt>Blend</dt><dd>{fmt_raw(agg.get("rv_blend"),1,"%")}</dd></div>'
            f'<div class="dl"><dt>Value</dt><dd>{fmt_raw(agg.get("rv_value"),1,"%")}</dd></div>'
            f'<div class="dl"><dt>Tech</dt><dd>{fmt_raw(ana.get("tech"),1,"%")}</dd></div>'
            f'<div class="dl"><dt>Industrials</dt><dd>{fmt_raw(ana.get("ind"),1,"%")}</dd></div>'
            '</dl></div>'
        )

    # Performance & riesgo vs BM (datos reales del PDF, página 7) — tabla 2 columnas
    perf = s.get("perf") or {}

    def _strat_cell(v, signed=True, digits=2, suffix="%"):
        """Estrategia: bold + color si signed=True, sin color si signed=False (vol/up/dn)."""
        if v is None:
            return '<td class="s">—</td>'
        if signed:
            cls = "s pos" if v >= 0 else "s neg"
            return f'<td class="{cls}">{fmt_raw(v, digits, suffix)}</td>'
        return f'<td class="s">{fmt_raw(v, digits, suffix)}</td>'

    def _bm_cell(v, signed=True, digits=2, suffix="%"):
        """Benchmark: color suave si signed, sin color si no. Sin bold."""
        if v is None:
            return '<td>—</td>'
        if signed:
            cls = "pos" if v >= 0 else "neg"
            return f'<td class="{cls}">{fmt_raw(v, digits, suffix)}</td>'
        return f'<td>{fmt_raw(v, digits, suffix)}</td>'

    perf_rows = [
        ("Ret 1Y",        "ret_1y",     "bm_ret_1y",     True),
        ("Ret 3Y",        "ret_3y",     "bm_ret_3y",     True),
        ("Ret 5Y",        "ret_5y",     "bm_ret_5y",     True),
        ("Vol 5Y",        "vol_5y",     "bm_vol_5y",     False),
        ("Sharpe 5Y",     "sharpe_5y",  "bm_sharpe_5y",  True),
        ("Max DD 3Y",     "dd_3y",      "bm_dd_3y",      True),
        ("Max DD 5Y",     "dd_5y",      "bm_dd_5y",      True),
        ("Up Capture 5Y", "up_5y",      "bm_up_5y",      False),
        ("Down Capture 5Y","dn_5y",     "bm_dn_5y",      False),
    ]
    perf_rows_html = ""
    for label, s_key, b_key, signed in perf_rows:
        sv = perf.get(s_key)
        bv = perf.get(b_key)
        perf_rows_html += (
            f'<tr><td>{label}</td>'
            f'{_strat_cell(sv, signed)}'
            f'{_bm_cell(bv, signed)}'
            f'</tr>'
        )
    perf_block = (
        '<section class="sheet-block">'
        '<span class="grp-t">Performance & riesgo vs BM</span>'
        '<table class="perftab">'
        '<thead><tr><th></th><th>Estrategia</th><th>Benchmark</th></tr></thead>'
        f'<tbody>{perf_rows_html}</tbody>'
        '</table></section>'
    )

    holdings_block = (
        '<section class="sheet-block">'
        '<span class="grp-t">Holdings (top 15)</span>'
        '<div class="changes"><table>'
        '<thead><tr><th>Peso</th><th>Posición</th><th class="num">Ret 3Y</th><th class="num">Ret 5Y</th></tr></thead>'
        f'<tbody>{hold_rows}</tbody></table></div></section>'
    )

    # Veredicto narrativa breve
    if ana["verdict"] == "Crítico":
        vd_text = "Necesita rebalanceo significativo: múltiples desalineaciones con la house view."
    elif ana["verdict"] == "Atención":
        vd_text = "Hay puntos a corregir. Identificá los problemas y planificá ajustes."
    elif ana["verdict"] == "OK":
        vd_text = "Cartera razonablemente alineada con la house view. Monitoreo de rutina."
    else:
        vd_text = "Composición no detectada automáticamente. Requiere revisión manual."

    return f"""
          <tr class="ov-detail" id="{det_id}"><td colspan="9"><div class="dwrap"><div class="dinner">

            <div class="sheet">

              <div class="sheet-h">
                <div class="sheet-h-l">
                  <span class="sheet-grp">{html.escape(s["cliente"])}</span>
                  <h3>{html.escape(s["estrategia"])}</h3>
                  <div class="chips">
                    <span class="chip"><span class="chip-ag">Perfil</span><span>{html.escape(perfil)}</span></span>
                    <span class="chip"><span class="chip-ag">Cliente</span><span>{html.escape(s["cliente"])}</span></span>
                    <span class="chip"><span class="chip-ag">Cierre</span><span>30-Abr-2026</span></span>
                  </div>
                </div>
                <div class="sheet-h-r">
                  <span class="vd {ana['vd_class']}">{html.escape(ana['verdict'])}</span>
                  <span class="sheet-grp">{len(ana.get('bad', []))} hallazgos · {len(ana.get('good', []))} fortalezas</span>
                </div>
              </div>

              <div class="sheet-body">
                <div class="sheet-thesis">
                  <div class="block verdict">
                    <span class="lbl">Veredicto</span>
                    <p>{html.escape(vd_text)}</p>
                  </div>
                  {good_html}
                  {bad_html}
                  {changes_html}
                  {miss_html}
                </div>
                <div class="sheet-data">
                  <div class="data-grp full">
                    <span class="grp-t">Composición ponderada</span>
                    <div class="alloc">{bar_html}</div>
                  </div>
                  {rf_block}
                  {rv_block}
                </div>
              </div>

              <div class="sheet-extra">
                {perf_block}
                {holdings_block}
              </div>

            </div>

          </div></div></td></tr>"""


def render_section(cliente: str, strats: list[dict], idx: int) -> str:
    """Sección completa para un cliente con su divider y tabla overview."""
    accent = ACCENTS[idx % len(ACCENTS)]
    anchor = safe_id(cliente)
    rows = ""
    for j, s in enumerate(sorted(strats, key=lambda x: (x.get("sub") or "", x["estrategia"]))):
        rows += render_overview_row(s, j)
    n = len(strats)
    meta = f"{n} estrategia{'s' if n != 1 else ''} en revisión. Click en cualquier fila para abrir la ficha."

    return f"""
<section class="cat" id="{anchor}" style="--gc:var({accent})">
  <div class="divider"><div class="container divider-in">
    <div>
      <span class="eyebrow">Cliente {idx+1:02d}</span>
      <h2>{html.escape(cliente)}</h2>
      <p class="divider-meta">{html.escape(meta)}</p>
    </div>
    <span class="divider-num">{idx+1:02d}</span>
  </div></div>

  <div class="container">
    <div class="ov-wrap">
      <table class="ov">
        <thead><tr>
          <th>Estrategia</th>
          <th>Perfil</th>
          <th class="num">RF %</th>
          <th class="num">RV %</th>
          <th class="num">YTW</th>
          <th class="num">Dur</th>
          <th class="num">Ret 5Y</th>
          <th class="num">DD 5Y</th>
          <th class="ctr">Veredicto</th>
        </tr></thead>
        <tbody>{rows}
        </tbody>
      </table>
    </div>
  </div>
</section>
"""


# ---------------------------------------------------------------------------
# 8. HTML full skeleton
# ---------------------------------------------------------------------------
HEAD_CSS = r"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Revisión de Estrategias · LATAM ConsultUs · Abril 2026</title>
<meta name="description" content="Revisión de las estrategias de inversión de Abril 2026 — alineación con house view, hallazgos y cambios sugeridos.">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@300;400;600;700;800&display=swap" rel="stylesheet">
<style>
:root{
  --navy:#0c2e4e; --navy-900:#081f37; --navy-700:#143d63; --navy-600:#1c4d7a;
  --sky:#4986c4; --sky-300:#7cabd6; --sky-100:#d9e6f2; --steel:#86a9c9;
  --g1:#dde3ea; --g2:#f1f4f8; --g3:#aab3c0; --g4:#647688; --ink:#0c2e4e;
  --paper:#f6f8fb; --card:#ffffff; --border:#e4e9f0;
  --fg:#13314f; --fg-muted:#5d6e80; --fg-subtle:#94a1b1;
  --pos:#1f8a5b; --pos-bg:#e6f3ec; --neg:#c0413a; --neg-bg:#f8e6e4; --neu:#8a96a4;
  --warn:#c98a2b; --warn-bg:#fbf1de;
  --c-cons:#3d78c0; --c-mod:#1f9aa0; --c-din:#6a63c8; --c-otro:#a25fa6;
  --grad-navy:linear-gradient(158deg,#081f37 0%,#0c2e4e 52%,#143d63 100%);
  --sans:'Open Sans',-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif;
  --maxw:1240px; --gut:clamp(18px,3.4vw,46px);
  --r:14px; --r-s:9px;
  --sh1:0 1px 2px rgba(12,46,78,.05),0 1px 1px rgba(12,46,78,.04);
  --sh2:0 10px 30px rgba(12,46,78,.07),0 2px 8px rgba(12,46,78,.05);
  --sh3:0 22px 60px rgba(12,46,78,.13),0 6px 16px rgba(12,46,78,.06);
}
*{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth;scroll-padding-top:72px;-webkit-text-size-adjust:100%}
body{font-family:var(--sans);color:var(--fg);background:var(--paper);font-size:15px;line-height:1.6;
  -webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;overflow-x:hidden}
img{max-width:100%;display:block}
a{color:inherit;text-decoration:none}
.num,table td,table th,.kv,.hm-v,.fig{font-variant-numeric:tabular-nums;font-feature-settings:"tnum" 1}
.container{max-width:var(--maxw);margin:0 auto;padding-inline:var(--gut)}
.eyebrow{display:inline-block;font-size:11px;font-weight:700;letter-spacing:.2em;text-transform:uppercase;color:var(--sky);margin-bottom:14px}
.eyebrow.light{color:var(--sky-300)}
.lbl{display:block;font-size:10.5px;font-weight:700;letter-spacing:.16em;text-transform:uppercase;color:var(--fg-subtle)}
.muted{color:var(--fg-muted);font-weight:400}
h1,h2,h3,h4{font-weight:300;letter-spacing:-.02em;line-height:1.06;color:var(--ink)}
h1 strong,h2 strong,h3 strong{font-weight:700}

.topnav{position:sticky;top:0;z-index:60;background:rgba(246,248,251,.86);
  backdrop-filter:saturate(150%) blur(12px);-webkit-backdrop-filter:saturate(150%) blur(12px);
  border-bottom:1px solid var(--border)}
.nav-in{display:flex;align-items:center;gap:26px;height:60px}
.brand{font-size:13px;font-weight:800;letter-spacing:.04em;color:var(--navy);text-transform:uppercase}
.brand span{color:var(--sky);margin-left:6px;font-weight:600}
.nav-links{display:flex;gap:18px;margin-left:auto;font-size:12.5px;font-weight:600;color:var(--fg-muted);
  overflow-x:auto;scrollbar-width:none;white-space:nowrap}
.nav-links::-webkit-scrollbar{display:none}
.nav-links a{padding:5px 0;border-bottom:2px solid transparent;transition:color .18s,border-color .18s}
.nav-links a:hover,.nav-links a.active{color:var(--navy);border-bottom-color:var(--sky)}
.nav-date{flex-shrink:0;font-size:10.5px;font-weight:700;letter-spacing:.13em;text-transform:uppercase;color:var(--fg-subtle)}

.hero{position:relative;background:var(--grad-navy);color:#fff;overflow:hidden;isolation:isolate}
.hero-bg{position:absolute;inset:0;z-index:-1;
  background:
   radial-gradient(120% 90% at 84% -10%,rgba(124,171,214,.32),transparent 55%),
   radial-gradient(80% 70% at -5% 110%,rgba(73,134,196,.18),transparent 60%);}
.hero-bg::after{content:"";position:absolute;inset:0;opacity:.5;
  background-image:radial-gradient(rgba(255,255,255,.05) 1px,transparent 1px);background-size:26px 26px;
  -webkit-mask-image:linear-gradient(180deg,transparent,#000 38%,#000 82%,transparent);
          mask-image:linear-gradient(180deg,transparent,#000 38%,#000 82%,transparent);}
.hero-in{position:relative;padding-block:clamp(34px,6vw,72px) clamp(46px,7vw,92px)}
.hero-brand{font-size:13px;font-weight:800;letter-spacing:.16em;text-transform:uppercase;color:#fff;margin-bottom:clamp(30px,6vw,62px)}
.hero-brand span{color:var(--sky-300);font-weight:600;margin-left:8px}
.hero h1{font-size:clamp(40px,7.5vw,96px);line-height:.95;letter-spacing:-.03em;color:#fff;margin:6px 0 0}
.hero h1 strong{color:#fff;font-weight:600}
.hero-sub{max-width:58ch;margin-top:26px;font-size:clamp(15px,1.7vw,18.5px);line-height:1.62;color:#cbd9e8;font-weight:300}
.hero-meta{display:flex;flex-wrap:wrap;gap:clamp(24px,5vw,60px);margin-top:clamp(32px,5vw,54px);
  padding-top:30px;border-top:1px solid rgba(255,255,255,.16)}
.hero-meta>div{display:flex;flex-direction:column;gap:4px}
.hm-v{font-size:clamp(26px,3.4vw,40px);font-weight:300;letter-spacing:-.02em;color:#fff}
.hm-v small{font-size:.5em;color:var(--sky-300);font-weight:600;letter-spacing:.02em}
.hm-l{font-size:11px;font-weight:700;letter-spacing:.13em;text-transform:uppercase;color:#86a0bb}

.idx{padding-block:clamp(40px,6vw,72px)}
.idx-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:clamp(14px,1.8vw,22px)}
.idx-card{position:relative;display:flex;flex-direction:column;background:var(--card);border:1px solid var(--border);
  border-radius:var(--r);box-shadow:var(--sh1);padding:22px 22px 20px;overflow:hidden;
  transition:box-shadow .25s,transform .25s;cursor:pointer}
.idx-card:hover{box-shadow:var(--sh2);transform:translateY(-3px)}
.idx-card::before{content:"";position:absolute;left:0;top:0;width:100%;height:3px;background:var(--gc)}
.idx-n{font-size:12px;font-weight:800;letter-spacing:.1em;color:var(--gc)}
.idx-card h3{font-size:clamp(17px,1.5vw,21px);font-weight:700;margin:14px 0 0;letter-spacing:-.01em}
.idx-card p{font-size:12.5px;color:var(--fg-muted);line-height:1.5;margin-top:8px;flex:1}
.idx-foot{display:flex;align-items:baseline;justify-content:space-between;margin-top:16px;padding-top:13px;border-top:1px solid var(--g2)}
.idx-foot .c{font-size:24px;font-weight:300;color:var(--navy)}
.idx-foot .c small{font-size:11px;font-weight:700;color:var(--fg-subtle);text-transform:uppercase;letter-spacing:.1em;margin-left:5px}
.idx-foot .go{font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--gc)}

.panorama{padding-block:clamp(8px,2vw,20px) clamp(40px,6vw,72px)}
.sec-head{max-width:760px;margin-bottom:clamp(26px,4vw,42px)}
.sec-head h2{font-size:clamp(27px,4.2vw,46px)}
.sec-head .lede{font-size:clamp(15px,1.6vw,18px);line-height:1.6;color:var(--fg-muted);margin-top:15px;font-weight:300;max-width:68ch}
.card{background:var(--card);border:1px solid var(--border);border-radius:var(--r);box-shadow:var(--sh2)}
.view-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:clamp(12px,1.6vw,18px)}
.view-card{padding:18px 20px;background:var(--card);border:1px solid var(--border);border-radius:var(--r);box-shadow:var(--sh1)}
.view-card .lbl{margin-bottom:6px}
.view-card h4{font-size:17px;font-weight:700;color:var(--navy)}
.view-card p{font-size:12.5px;color:var(--fg-muted);margin-top:6px;line-height:1.55}
.view-list{margin-top:18px;display:grid;grid-template-columns:repeat(2,1fr);gap:14px 24px;padding:22px 26px;background:var(--card);border:1px solid var(--border);border-radius:var(--r);box-shadow:var(--sh1)}
.view-list .vli{display:flex;align-items:baseline;gap:10px;font-size:13px;color:var(--fg)}
.tag{display:inline-flex;align-items:center;justify-content:center;min-width:88px;padding:3px 9px;border-radius:999px;font-size:10.5px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;flex-shrink:0}
.tag.ow{background:#e6f3ec;color:#1f8a5b}
.tag.uw{background:#f8e6e4;color:#c0413a}
.tag.nw{background:var(--g2);color:var(--fg-muted)}
.tag.mod{background:#fbf1de;color:#c98a2b}

.cat{scroll-margin-top:60px}
.divider{background:var(--grad-navy);color:#fff;margin-top:clamp(28px,5vw,56px);position:relative;overflow:hidden}
.divider::before{content:"";position:absolute;left:0;top:0;bottom:0;width:6px;background:var(--gc)}
.divider-in{display:flex;justify-content:space-between;align-items:flex-end;gap:24px;padding-block:clamp(32px,5vw,58px)}
.divider .eyebrow{color:var(--gc);filter:saturate(1.3) brightness(1.35)}
.divider h2{color:#fff;font-size:clamp(28px,5vw,54px);margin-top:6px}
.divider-meta{margin-top:13px;font-size:13.5px;color:#b9c8d8;font-weight:300;max-width:54ch}
.divider-num{font-size:clamp(56px,11vw,140px);font-weight:200;line-height:.8;color:rgba(255,255,255,.10);letter-spacing:-.04em}

.ov-wrap{margin-top:clamp(22px,3.5vw,40px);background:var(--card);border:1px solid var(--border);border-radius:var(--r);box-shadow:var(--sh2);overflow:hidden}
table.ov{width:100%;border-collapse:collapse;font-size:13px}
.ov thead th{background:var(--g2);color:var(--fg-muted);font-size:10px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;padding:11px 12px;text-align:left;border-bottom:1px solid var(--border);white-space:nowrap}
.ov th.num,.ov td.num{text-align:right}
.ov th.ctr,.ov td.ctr{text-align:center}
.ov tbody tr.ov-row{cursor:pointer;transition:background .14s}
.ov tbody tr.ov-row:hover{background:var(--sky-100)}
.ov tbody tr.ov-row.open{background:var(--g2)}
.ov tbody tr.ov-row.pending{cursor:default}
.ov tbody tr.ov-row.pending:hover{background:transparent}
.ov td{padding:11px 12px;border-bottom:1px solid var(--g2);vertical-align:middle}
.ov-name{min-width:200px}
.ov-name .chev{display:inline-block;width:7px;height:7px;border-right:2px solid var(--gc);border-bottom:2px solid var(--gc);transform:rotate(-45deg);margin-right:11px;transition:transform .3s ease;vertical-align:middle}
.ov-row.pending .chev{opacity:.3}
.ov-row.open .ov-name .chev{transform:rotate(45deg)}
.ov-nm{font-weight:700;color:var(--navy)}
.ov-row:hover .ov-nm{color:var(--gc)}
.ov-sub{display:block;font-size:10.5px;color:var(--fg-subtle);font-weight:600;text-transform:uppercase;letter-spacing:.04em;margin-top:2px;padding-left:18px}
.ov-cat{color:var(--fg-muted);font-size:12px}
.ov td.num{color:var(--fg)}
.ov .pos{color:var(--pos)}.ov .neg{color:var(--neg)}.ov .neu{color:var(--fg-subtle)}
.ov-detail>td{padding:0;border:0;background:#fbfcfe}
.dwrap{display:grid;grid-template-rows:0fr;transition:grid-template-rows .42s cubic-bezier(.16,1,.3,1)}
.ov-detail.open .dwrap{grid-template-rows:1fr}
.dinner{overflow:hidden;min-height:0}

.vd{display:inline-flex;align-items:center;gap:6px;font-size:11px;font-weight:800;letter-spacing:.04em;text-transform:uppercase;padding:4px 10px;border-radius:999px;white-space:nowrap}
.vd.ok{background:var(--pos-bg);color:var(--pos)}
.vd.wn{background:var(--warn-bg);color:var(--warn)}
.vd.cr{background:var(--neg-bg);color:var(--neg)}
.vd.pd{background:var(--g2);color:var(--fg-subtle)}
.vd::before{content:"";width:7px;height:7px;border-radius:50%;background:currentColor}

.sheet{padding:clamp(18px,2.4vw,30px) clamp(16px,2vw,26px) clamp(22px,2.6vw,32px)}
.sheet-h{display:flex;justify-content:space-between;align-items:flex-start;gap:18px;padding-bottom:16px;border-bottom:1px solid var(--border);flex-wrap:wrap}
.sheet-grp{display:inline-flex;align-items:center;gap:8px;font-size:10px;font-weight:700;letter-spacing:.13em;text-transform:uppercase;color:var(--fg-subtle)}
.sheet-grp::before{content:"";width:8px;height:8px;border-radius:50%;background:var(--gc)}
.sheet-h h3{font-size:clamp(17px,1.9vw,23px);margin:9px 0 0;font-weight:600;letter-spacing:-.01em}
.chips{display:flex;flex-wrap:wrap;gap:7px;margin-top:11px}
.chip{display:inline-flex;align-items:center;gap:6px;font-size:11.5px;font-weight:700;color:var(--navy);background:var(--g2);border:1px solid var(--border);border-radius:6px;padding:3px 9px}
.chip-ag{font-size:9px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--fg-subtle)}
.sheet-h-r{display:flex;flex-direction:column;align-items:flex-end;gap:9px;flex-shrink:0;min-width:172px}
.sheet-body{display:grid;grid-template-columns:1.4fr 1fr;gap:clamp(20px,3.2vw,42px);padding:20px 0 0}
.sheet-thesis .block{margin-bottom:18px}
.sheet-thesis .lbl{margin-bottom:8px}
.sheet-thesis p{font-size:13.5px;line-height:1.62;color:var(--fg);max-width:64ch}
.sheet-thesis p+p{margin-top:9px}
.sheet-thesis .block.verdict p{font-size:14px;font-weight:600;color:var(--navy)}
.findings{list-style:none;display:grid;gap:9px;margin-top:4px}
.findings li{position:relative;padding-left:24px;font-size:13px;line-height:1.55;color:var(--fg)}
.findings li::before{content:"";position:absolute;left:2px;top:8px;width:8px;height:8px;border-radius:50%}
.findings.good li::before{background:var(--pos)}
.findings.bad li::before{background:var(--neg)}
.findings.fix li::before{background:var(--warn)}
.findings li strong{color:var(--navy);font-weight:700}
.sheet-data{display:grid;grid-template-columns:1fr 1fr;gap:16px 24px;align-content:start}
.data-grp.full{grid-column:1/-1}
.data-grp .grp-t{display:block;font-size:10px;font-weight:700;letter-spacing:.13em;text-transform:uppercase;color:var(--gc);margin-bottom:8px;padding-bottom:6px;border-bottom:1px solid var(--border)}
.data-grp.full dl{display:grid;grid-template-columns:1fr 1fr;gap:0 24px}
.dl{display:flex;justify-content:space-between;align-items:baseline;gap:12px;padding:5px 0;border-bottom:1px solid var(--g2)}
.dl dt{font-size:12px;color:var(--fg-muted);white-space:nowrap}
.dl dd{font-size:13px;font-weight:700;color:var(--navy);text-align:right}
.dl dd.strong{color:var(--gc);font-weight:800}
.dl dd .muted{font-size:11px}
.dl dd.pos{color:var(--pos)}
.dl dd.neg{color:var(--neg)}
.mono{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:12px;letter-spacing:-.02em}

.alloc{margin-top:2px}
.alloc-bar{display:flex;height:14px;border-radius:4px;overflow:hidden;background:var(--g1)}
.alloc-bar span{display:block;height:100%}
.alloc-leg{display:grid;grid-template-columns:1fr 1fr;gap:5px 16px;margin-top:11px}
.alloc-leg div{display:flex;align-items:center;gap:7px;font-size:12px;color:var(--fg-muted)}
.alloc-leg i{width:9px;height:9px;border-radius:2px}
.alloc-leg b{margin-left:auto;color:var(--navy);font-weight:700}

.changes{margin-top:14px;background:#fbfcfe;border:1px solid var(--border);border-radius:var(--r-s);overflow:hidden}
.changes table{width:100%;border-collapse:collapse;font-size:12.5px}
.changes th{background:var(--g2);color:var(--fg-muted);font-size:9.5px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;padding:9px 12px;text-align:left}
.changes td{padding:9px 12px;border-top:1px solid var(--g2);vertical-align:top;line-height:1.45}
.changes td.act{font-weight:700;color:var(--navy);min-width:180px}
.changes td.dir{font-weight:800;font-size:11px;letter-spacing:.04em;text-transform:uppercase;white-space:nowrap;width:1%}
.changes td.num{white-space:nowrap;width:1%;text-align:right}
.changes td:last-child{min-width:140px}
.changes td.dir.up{color:var(--pos)}
.changes td.dir.dn{color:var(--neg)}
.changes td.dir.sw{color:var(--warn)}
.changes td.dir.new{color:var(--sky)}
.changes th.num,.changes td.num{text-align:right}
.changes td.num.pos{color:var(--pos);font-weight:700}
.changes td.num.neg{color:var(--neg);font-weight:700}
.changes td.num.neu{color:var(--fg-subtle)}

.perftab{width:100%;border-collapse:collapse;font-size:12.5px;margin-top:6px}
.perftab th{font-size:10px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--fg-subtle);text-align:right;padding:6px 10px;border-bottom:1px solid var(--border)}
.perftab th:first-child{text-align:left}
.perftab td{padding:7px 10px;border-bottom:1px solid var(--g2);text-align:right;font-variant-numeric:tabular-nums}
.perftab td:first-child{text-align:left;color:var(--fg-muted)}
.perftab td.pos{color:var(--pos);font-weight:700}
.perftab td.neg{color:var(--neg);font-weight:700}
.perftab td.s{font-weight:700;color:var(--navy)}
.perftab td.s.pos{color:var(--pos)}
.perftab td.s.neg{color:var(--neg)}

.sheet-extra{display:grid;grid-template-columns:1fr 1fr;gap:clamp(20px,3.2vw,42px);margin-top:clamp(20px,3vw,32px);padding-top:clamp(20px,3vw,32px);border-top:1px solid var(--border)}
.sheet-block{display:flex;flex-direction:column;min-width:0}
.sheet-block .grp-t{display:block;font-size:10px;font-weight:700;letter-spacing:.13em;text-transform:uppercase;color:var(--gc);margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid var(--border)}
@media(max-width:1000px){.sheet-extra{grid-template-columns:1fr;gap:24px}}

.diag{padding-block:clamp(40px,6vw,72px);background:var(--card)}
.diag-grid{display:grid;grid-template-columns:1.2fr 1fr;gap:clamp(20px,3vw,42px)}
.diag-card{background:var(--card);border:1px solid var(--border);border-radius:var(--r);box-shadow:var(--sh1);padding:22px 24px;margin-bottom:14px}
.diag-card .lbl{margin-bottom:14px}
.findings-bar{display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid var(--g2);font-size:13px}
.findings-bar:last-child{border-bottom:none}
.findings-bar .fb-label{flex:1;color:var(--fg)}
.findings-bar .fb-track{flex:1.2;height:6px;background:var(--g1);border-radius:3px;overflow:hidden}
.findings-bar .fb-fill{display:block;height:100%;background:var(--warn);border-radius:3px}
.findings-bar .fb-count{width:80px;text-align:right;font-weight:700;color:var(--navy);font-size:12.5px}
.top-problems table{width:100%;border-collapse:collapse;font-size:12.5px}
.top-problems th{background:var(--g2);color:var(--fg-muted);font-size:10px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;padding:10px 12px;text-align:left}
.top-problems td{padding:10px 12px;border-bottom:1px solid var(--g2)}
.top-problems tr{cursor:pointer;transition:background .14s}
.top-problems tbody tr:hover{background:var(--sky-100)}
.top-problems .badcount{font-weight:800;color:var(--neg)}
.client-dist table{width:100%;border-collapse:collapse;font-size:12.5px}
.client-dist th{background:var(--g2);color:var(--fg-muted);font-size:10px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;padding:10px 12px;text-align:left}
.client-dist td{padding:8px 12px;border-bottom:1px solid var(--g2)}
.client-dist td.num{text-align:right;font-variant-numeric:tabular-nums}
.client-dist .ok{color:var(--pos);font-weight:700}
.client-dist .wn{color:var(--warn);font-weight:700}
.client-dist .cr{color:var(--neg);font-weight:700}
@media(max-width:1000px){.diag-grid{grid-template-columns:1fr}}

.closing{background:var(--grad-navy);color:#fff;margin-top:clamp(40px,6vw,80px)}
.closing-in{padding-block:clamp(46px,7vw,92px)}
.closing-brand{font-size:13px;font-weight:800;letter-spacing:.16em;text-transform:uppercase;color:#fff;margin-bottom:clamp(26px,5vw,50px)}
.closing-brand span{color:var(--sky-300);font-weight:600;margin-left:8px}
.closing-headline{font-size:clamp(28px,5vw,56px);font-weight:300;line-height:1.05;letter-spacing:-.025em}
.closing-headline strong{font-weight:600}
.closing-rule{width:54px;height:2px;background:var(--sky);margin:clamp(24px,4vw,38px) 0}
.disclaimer{margin-top:clamp(32px,5vw,54px);padding-top:24px;border-top:1px solid rgba(255,255,255,.16)}
.disclaimer p{font-size:11px;line-height:1.7;color:#8ea3bb;max-width:780px}

.totop{position:fixed;right:20px;bottom:20px;z-index:70;width:44px;height:44px;border-radius:50%;
  background:var(--navy);color:#fff;display:grid;place-items:center;font-size:18px;box-shadow:var(--sh3);
  opacity:0;pointer-events:none;transform:translateY(10px);transition:opacity .3s,transform .3s;border:none;cursor:pointer}
.totop.show{opacity:1;pointer-events:auto;transform:none}
.totop:hover{background:var(--navy-700)}

@media(max-width:1000px){
  .idx-grid{grid-template-columns:repeat(2,1fr)}
  .view-grid{grid-template-columns:repeat(2,1fr)}
  .sheet-body{grid-template-columns:1fr;gap:24px}
  .view-list{grid-template-columns:1fr}
}
@media(max-width:820px){
  body{font-size:14.5px}
  .divider-in{flex-direction:column;align-items:flex-start;gap:6px}
  .divider-num{align-self:flex-end;margin-top:-30px}
  .ov thead{display:none}
  .ov,.ov tbody,.ov tr,.ov td{display:block;width:100%}
  .ov tbody tr.ov-row{padding:13px 14px;border-bottom:1px solid var(--border);position:relative}
  .ov td{border:none;padding:2px 0}
  .ov-name{padding-right:30px!important}
  .ov-row .chev{position:absolute;right:8px;top:18px;margin:0}
  .ov td.num,.ov td.ctr{display:inline-flex;align-items:baseline;gap:6px;width:auto;margin-right:16px;font-size:12.5px}
  .ov td.num::before{content:attr(data-l);font-size:9px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--fg-subtle)}
}
@media(max-width:560px){
  .idx-grid,.view-grid{grid-template-columns:1fr}
}
@media print{
  .topnav,.totop{display:none}
  .dwrap{grid-template-rows:1fr!important}
  body{background:#fff}
}
</style>
</head>
<body id="top">"""

PANORAMA_HTML = r"""
<section class="panorama" id="panorama"><div class="container">
  <div class="sec-head">
    <span class="eyebrow">House view · Mayo 2026</span>
    <h2>Pro-riesgo, pero con una <strong>vara de calidad más alta</strong></h2>
    <p class="lede">Consenso institucional moderadamente pro-riesgo, con tres fuerzas dominantes: IA, fragmentación geopolítica e inflación más volátil. Las estrategias se evalúan contra esta vara.</p>
  </div>

  <div class="view-grid">
    <article class="view-card">
      <span class="lbl">Riesgo agregado</span>
      <h4>OW moderado</h4>
      <p>Preferencia por equities y activos reales, con diversificación más deliberada.</p>
    </article>
    <article class="view-card">
      <span class="lbl">Motor dominante</span>
      <h4>IA — infraestructura física</h4>
      <p>Energía, chips, data centers, cooling, grid, cobre. No solo mega-cap tech.</p>
    </article>
    <article class="view-card">
      <span class="lbl">Riesgo clave</span>
      <h4>Inflación volátil</h4>
      <p>Shocks de oferta debilitan el hedge tradicional stock-bond.</p>
    </article>
    <article class="view-card">
      <span class="lbl">Escenario base</span>
      <h4>60% probabilidad</h4>
      <p>Crecimiento positivo, inflación sticky, Fed eventualmente recorta.</p>
    </article>
  </div>

  <div class="view-list">
    <div class="vli"><span class="tag ow">OW</span><span><strong>Global equities</strong> · earnings sostienen el ciclo, IA como motor</span></div>
    <div class="vli"><span class="tag ow">OW</span><span><strong>EM equities selectivo</strong> · Norte Asia (semis/IA) + LatAm (commodities)</span></div>
    <div class="vli"><span class="tag ow">OW</span><span><strong>Commodities</strong> · energía y cobre como hedge y beneficiarios</span></div>
    <div class="vli"><span class="tag ow">OW</span><span><strong>Gold</strong> · hedge geopolítico e inflacionario</span></div>
    <div class="vli"><span class="tag ow">OW</span><span><strong>Infraestructura / real assets</strong> · resiliencia inflacionaria</span></div>
    <div class="vli"><span class="tag mod">MOD</span><span><strong>Japan</strong> · reformas corporativas, mejora estructural</span></div>
    <div class="vli"><span class="tag mod">SEL</span><span><strong>High yield</strong> · selectivo, underwriting exigente</span></div>
    <div class="vli"><span class="tag nw">NEU</span><span><strong>Fixed income core</strong> · carry sí, duration larga no</span></div>
    <div class="vli"><span class="tag nw">NEU</span><span><strong>Investment grade credit</strong> · buen carry, spreads ajustados</span></div>
    <div class="vli"><span class="tag nw">NEU</span><span><strong>Europe equities</strong> · valuaciones OK, energía pesa</span></div>
    <div class="vli"><span class="tag uw">UW</span><span><strong>Duration larga US</strong> · term premium + déficits</span></div>
    <div class="vli"><span class="tag uw">UW</span><span><strong>Private credit</strong> · iliquidez + opacidad de valuaciones</span></div>
    <div class="vli"><span class="tag uw">UW</span><span><strong>Cash</strong> · útil táctico, malo estructural</span></div>
  </div>
</div></section>
"""

FOOTER_HTML = rf"""
<footer class="closing"><div class="container closing-in">
  <img class="closing-logo" src="{LOGO_DATA}" alt="LATAM ConsultUs" style="height:36px;width:auto;display:block;margin-bottom:clamp(26px,5vw,50px);filter:brightness(0) invert(1)">
  <div class="closing-headline">Una revisión<br><strong>vale lo que cuesta cambiar.</strong></div>
  <div class="closing-rule"></div>
  <p style="color:#cbd9e8;font-weight:300;max-width:60ch;font-size:15.5px;line-height:1.62">
    Este documento contrasta cada estrategia con el house view interno y propone cambios concretos. La prioridad de revisión surge del nivel de desalineación detectado, no del tamaño del cliente.
  </p>
  <div class="disclaimer">
    <p>Documento interno de uso exclusivo para el equipo de LATAM ConsultUs. Los valores citados corresponden al cierre del 30 de abril de 2026 según los Informes de Gestión y PPTs de cada estrategia. Las opiniones expresadas son análisis técnico interno y no constituyen recomendación de inversión para clientes finales. Cualquier cambio en cartera debe ser validado por el comité de inversiones.</p>
  </div>
</div></footer>

<button class="totop" id="totop" aria-label="Volver arriba">↑</button>

<script>
document.querySelectorAll('.ov-row[data-target]').forEach(row=>{{
  row.addEventListener('click',()=>{{
    const id=row.getAttribute('data-target');
    const det=document.getElementById(id);
    const open=row.classList.toggle('open');
    if(det){{det.classList.toggle('open',open)}}
  }});
}});

const tt=document.getElementById('totop');
window.addEventListener('scroll',()=>{{
  tt.classList.toggle('show',window.scrollY>600);
}});
tt.addEventListener('click',()=>window.scrollTo({{top:0,behavior:'smooth'}}));
</script>

</body>
</html>
"""


# ---------------------------------------------------------------------------
# 7b. RADIOGRAFÍA — vista agregada de hallazgos
# ---------------------------------------------------------------------------
# Mapeo: substring que aparece dentro del <strong> del finding bad → categoría
FINDING_CATEGORIES = [
    ("Calidad crediticia agresiva",   "Calidad crediticia agresiva"),
    ("Duration",                       "Duration larga"),
    ("UW equities",                    "RV underweight vs perfil"),
    ("Sin EM equities",                "EM equity ausente"),
    ("Sin commodities",                "Gold/commodities ausente"),
    ("Sobre-expuesto Europa",          "Europa concentrada"),
    ("Cash elevado",                   "Cash exceso (>10%)"),
    ("Concentración en una sub",       "Solapamiento de fondos"),
]


def categorize_finding(html_finding: str) -> str | None:
    """Mapea un finding bad (string HTML) a una categoría conocida."""
    for substr, label in FINDING_CATEGORIES:
        if substr in html_finding:
            return label
    return None


def render_diag(strategies: list[dict]) -> str:
    """Render de la sección Radiografía con 3 bloques."""
    if not strategies:
        return ""

    total_strats = len(strategies)

    # --- Bloque A: hallazgos más comunes ---
    cat_counter: dict[str, int] = defaultdict(int)
    for s in strategies:
        bads = s.get("ana", {}).get("bad", []) or []
        # contar máx 1 por estrategia por categoría
        seen = set()
        for b in bads:
            cat = categorize_finding(b)
            if cat and cat not in seen:
                cat_counter[cat] += 1
                seen.add(cat)
    cat_sorted = sorted(cat_counter.items(), key=lambda x: -x[1])
    top_cats = cat_sorted[:8]
    if top_cats:
        max_count = top_cats[0][1]
    else:
        max_count = 1
    bars_html = ""
    for label, cnt in top_cats:
        pct = (cnt / max_count) * 100 if max_count else 0
        share = (cnt / total_strats) * 100 if total_strats else 0
        bars_html += (
            '<div class="findings-bar">'
            f'<span class="fb-label">{html.escape(label)}</span>'
            f'<span class="fb-track"><span class="fb-fill" style="width:{pct:.1f}%"></span></span>'
            f'<span class="fb-count">{cnt} estr · {share:.0f}%</span>'
            '</div>'
        )
    if not bars_html:
        bars_html = '<p class="muted" style="font-size:13px">No se detectaron hallazgos clasificables.</p>'

    # --- Bloque B: top 10 estrategias más problemáticas ---
    ranked = []
    for s in strategies:
        bads = s.get("ana", {}).get("bad", []) or []
        if not bads:
            continue
        ranked.append((s, len(bads)))
    ranked.sort(key=lambda x: -x[1])
    top_problems = ranked[:10]
    problems_rows = ""
    for s, n_bad in top_problems:
        det_id_prefix = f"d-{safe_id(s['cliente'])}-"
        # No tenemos el idx exacto acá; saltamos al anchor del cliente
        anchor = f"#{safe_id(s['cliente'])}"
        vd = s["ana"]["verdict"]
        vd_cls = s["ana"]["vd_class"]
        problems_rows += (
            f'<tr onclick="location.hash=\'{anchor}\'">'
            f'<td>{html.escape(s["cliente"])}</td>'
            f'<td><strong>{html.escape(s["estrategia"])}</strong></td>'
            f'<td>{html.escape(s["perfil"])}</td>'
            f'<td class="badcount">{n_bad}</td>'
            f'<td><span class="vd {vd_cls}">{html.escape(vd)}</span></td>'
            '</tr>'
        )
    if not problems_rows:
        problems_rows = '<tr><td colspan="5" class="muted" style="text-align:center">Sin estrategias con hallazgos críticos.</td></tr>'

    # --- Bloque C: distribución por cliente ---
    client_stats: dict[str, dict[str, int]] = defaultdict(lambda: {"total": 0, "OK": 0, "Atención": 0, "Crítico": 0, "Pendiente": 0})
    for s in strategies:
        c = s["cliente"]
        v = s["ana"]["verdict"]
        client_stats[c]["total"] += 1
        if v in client_stats[c]:
            client_stats[c][v] += 1
    # ordenar por # crítico descendente, luego atención, luego total
    clients_sorted = sorted(
        client_stats.items(),
        key=lambda kv: (-kv[1]["Crítico"], -kv[1]["Atención"], -kv[1]["total"]),
    )
    dist_rows = ""
    for cliente, st in clients_sorted:
        dist_rows += (
            f'<tr>'
            f'<td><strong>{html.escape(cliente)}</strong></td>'
            f'<td class="num">{st["total"]}</td>'
            f'<td class="num ok">{st["OK"]}</td>'
            f'<td class="num wn">{st["Atención"]}</td>'
            f'<td class="num cr">{st["Crítico"]}</td>'
            '</tr>'
        )

    return f"""
<section class="diag" id="radiografia"><div class="container">
  <div class="sec-head">
    <span class="eyebrow">Radiografía</span>
    <h2>Qué está <strong>roto</strong> en las estrategias</h2>
    <p class="lede">Mirada agregada para priorizar revisiones: hallazgos más comunes, estrategias con más problemas y distribución por cliente.</p>
  </div>
  <div class="diag-grid">
    <div>
      <div class="diag-card">
        <span class="lbl">Hallazgos más comunes</span>
        {bars_html}
      </div>
      <div class="diag-card top-problems">
        <span class="lbl">Top 10 más problemáticas</span>
        <table>
          <thead><tr><th>Cliente</th><th>Estrategia</th><th>Perfil</th><th># hallazgos</th><th>Veredicto</th></tr></thead>
          <tbody>{problems_rows}</tbody>
        </table>
      </div>
    </div>
    <div>
      <div class="diag-card client-dist">
        <span class="lbl">Distribución por cliente</span>
        <table>
          <thead><tr><th>Cliente</th><th class="num">Total</th><th class="num">OK</th><th class="num">Atención</th><th class="num">Crítico</th></tr></thead>
          <tbody>{dist_rows}</tbody>
        </table>
      </div>
    </div>
  </div>
</div></section>
"""


def render_html(strategies: list[dict]) -> str:
    # Agrupar por cliente, orden fijo: LATAM primero, luego alfabético
    by_client: dict[str, list[dict]] = defaultdict(list)
    for s in strategies:
        by_client[s["cliente"]].append(s)

    # Orden: LATAM ConsultUs primero, después por cantidad descendente / alfabético
    def client_key(name):
        if "LATAM" in name.upper():
            return (0, name)
        return (1, name)

    ordered = dict(sorted(by_client.items(), key=lambda kv: client_key(kv[0])))

    n_estr = sum(len(v) for v in ordered.values())
    n_cli = len(ordered)
    n_cr = sum(1 for s in strategies if s["ana"]["verdict"] == "Crítico")
    n_wn = sum(1 for s in strategies if s["ana"]["verdict"] == "Atención")
    n_ok = sum(1 for s in strategies if s["ana"]["verdict"] == "OK")
    n_pd = sum(1 for s in strategies if s["ana"]["verdict"] == "Pendiente")

    nav_links = '<a href="#panorama">House View</a><a href="#radiografia">Radiografía</a><a href="#indice">Clientes</a>'
    for cliente in list(ordered.keys())[:8]:
        nav_links += f'<a href="#{safe_id(cliente)}">{html.escape(cliente)}</a>'

    hero = f"""
<nav class="topnav"><div class="container nav-in">
  <a class="brand" href="#top"><img src="{LOGO_DATA}" alt="LATAM ConsultUs" style="height:30px;width:auto;display:block"></a>
  <div class="nav-links">{nav_links}</div>
  <span class="nav-date">Abril 2026</span>
</div></nav>

<header class="hero"><div class="hero-bg"></div>
<div class="container hero-in">
  <img class="hero-logo" src="{LOGO_DATA}" alt="LATAM ConsultUs" style="height:42px;width:auto;display:block;margin-bottom:clamp(30px,6vw,62px);filter:brightness(0) invert(1)">
  <h1>Revisión de <strong>Estrategias</strong></h1>
  <p class="hero-sub">Auditoría automatizada de las {n_estr} estrategias de inversión de Abril 2026 frente al house view interno: alineación cuantitativa, hallazgos y prioridad de revisión. Una ficha por estrategia, agrupadas por cliente, con composición ponderada por holding.</p>
  <div class="hero-meta">
    <div><span class="hm-v">{n_estr}<small>&nbsp;estrategias</small></span><span class="hm-l">Universo analizado</span></div>
    <div><span class="hm-v">{n_cli}</span><span class="hm-l">Clientes</span></div>
    <div><span class="hm-v">{n_cr}<small>&nbsp;crít · {n_wn} atn · {n_ok} ok</small></span><span class="hm-l">Veredictos</span></div>
    <div><span class="hm-v">30-Abr-2026</span><span class="hm-l">Fecha de cierre</span></div>
  </div>
</div></header>
"""

    idx_section = f"""
<section class="idx" id="indice"><div class="container">
  <div class="sec-head">
    <span class="eyebrow">Índice por cliente</span>
    <h2>{n_cli} clientes · <strong>{n_estr} estrategias</strong></h2>
    <p class="lede">Cada cliente tiene entre 1 y 5 estrategias. Hacé click en una tarjeta para saltar a la sección detallada. Las fichas se expanden al clickear cada fila.</p>
  </div>

  <div class="idx-grid">{render_index_cards(ordered)}
  </div>
</div></section>
"""

    sections = ""
    for i, (cliente, strats) in enumerate(ordered.items()):
        sections += render_section(cliente, strats, i)

    diag_section = render_diag(strategies)

    return HEAD_CSS + hero + PANORAMA_HTML + diag_section + idx_section + sections + FOOTER_HTML


# ---------------------------------------------------------------------------
# 9. MAIN
# ---------------------------------------------------------------------------
def main():
    catalog = load_catalog()
    strategies = process_all(catalog)

    html_out = render_html(strategies)
    OUT_HTML.write_text(html_out, encoding="utf-8")

    # Summary
    n_cr = sum(1 for s in strategies if s["ana"]["verdict"] == "Crítico")
    n_wn = sum(1 for s in strategies if s["ana"]["verdict"] == "Atención")
    n_ok = sum(1 for s in strategies if s["ana"]["verdict"] == "OK")
    n_pd = sum(1 for s in strategies if s["ana"]["verdict"] == "Pendiente")

    # Top ISINs faltantes
    miss_counter: dict[str, tuple[int, str, float]] = {}
    for s in strategies:
        for isin, peso, nombre in s["agg"].get("_missing_isins", []):
            cnt, name, tot = miss_counter.get(isin, (0, nombre, 0.0))
            miss_counter[isin] = (cnt + 1, name, tot + peso)
    top_missing = sorted(miss_counter.items(), key=lambda kv: -kv[1][0])[:20]

    # BM detectado: estrategias donde el parser de performance encontró
    # al menos un valor del benchmark (cualquiera de los bm_*)
    bm_keys = (
        "bm_ret_5y", "bm_ret_3y", "bm_ret_1y",
        "bm_vol_5y", "bm_sharpe_5y",
        "bm_dd_5y", "bm_dd_3y", "bm_up_5y", "bm_dn_5y",
    )
    n_bm = sum(
        1 for s in strategies
        if any((s.get("perf") or {}).get(k) is not None for k in bm_keys)
    )
    no_bm = [
        f"{s['cliente']} · {s['estrategia']}"
        for s in strategies
        if not any((s.get("perf") or {}).get(k) is not None for k in bm_keys)
    ]

    print("\n" + "=" * 60)
    print("RESUMEN DEL RUN")
    print("=" * 60)
    print(f"Estrategias procesadas: {len(strategies)}")
    print(f"  Crítico:   {n_cr}")
    print(f"  Atención:  {n_wn}")
    print(f"  OK:        {n_ok}")
    print(f"  Pendiente: {n_pd}")
    print(f"\nBenchmark detectado: {n_bm} / {len(strategies)}")
    if no_bm:
        print(f"Sin BM detectado ({len(no_bm)}):")
        for x in no_bm:
            print(f"  - {x}")
    print(f"\nHTML generado: {OUT_HTML.resolve()}")
    print(f"\nTop 20 ISINs no encontrados en catálogo:")
    if top_missing:
        for isin, (cnt, name, tot) in top_missing:
            print(f"  {isin}  ×{cnt}  peso acum {tot:.1f}%  — {name[:50]}")
    else:
        print("  (ninguno)")
    print()


if __name__ == "__main__":
    main()
