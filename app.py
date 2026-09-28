#!/usr/bin/env python3
"""
Dashify — Dashboard Multi-Marketplace + Fuente Genérica
Mercado Libre  |  Tienda Nube  |  Cualquier Excel/CSV
Uso:  python app.py  →  http://127.0.0.1:7841
Deps: pip install flask pandas openpyxl requests odfpy
"""
import io, os, uuid, pickle, secrets, threading, webbrowser, traceback, re, math, csv
from pathlib import Path
from datetime import datetime
from functools import wraps
import pandas as pd
from flask import Flask, request, jsonify, session, redirect, url_for, Response

PORT = int(os.environ.get("PORT", 7841))
HOST = os.environ.get("HOST", "127.0.0.1")
app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = 150 * 1024 * 1024

# ── Login (usuario único) y persistencia en disco ───────────────────────────
# DASHIFY_PASSWORD: contraseña para entrar al dashboard (obligatoria en producción).
# DASHIFY_SECRET_KEY: clave para firmar la sesión (si no se define, se genera una
#   al arrancar — sirve para probar local, pero en producción convine fijarla,
#   porque si cambia, se cierran las sesiones abiertas).
# DASHIFY_DATA_DIR: carpeta donde se guardan los datos cargados, para que no se
#   pierdan al reiniciar el servidor. En Render, apuntar esto a un Disco
#   persistente (por ej. /var/data) — si no, en cada deploy se borra.
app.config["SECRET_KEY"] = os.environ.get("DASHIFY_SECRET_KEY") or secrets.token_hex(32)
DASHIFY_PASSWORD = os.environ.get("DASHIFY_PASSWORD")  # None = login deshabilitado (uso local)

# DATABASE_URL: si está definida (por ej. apuntando a un Postgres gratis de
#   Supabase), los datos se guardan ahí — sobreviven aunque Render reinicie o
#   borre el disco (en el plan free de Render el disco NO es persistente).
#   Si no está definida, se guarda en un archivo local (sirve para probar en tu
#   compu, pero en Render free se pierde en cada reinicio).
DATABASE_URL = os.environ.get("DATABASE_URL")

DATA_DIR = Path(os.environ.get("DASHIFY_DATA_DIR", "./dashify_data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)
STATE_FILE = DATA_DIR / "dashify_state.pkl"


# STORE: { sid: { name, source, df_raw, df, config } }
# source: "ml" | "tn" | "custom"
# config: { date_col, amount_col, qty_col, cat_col, cat2_col, label_col, ... }
STORE = {}

# ── Store persistente para Mercado Libre (upsert por id_venta) ──────────────
# Todas las planillas ML se fusionan en un único DataFrame acumulado.
# La clave de deduplicación es id_venta ("# de venta").
# Si la misma operación aparece en dos planillas, se conserva la fila más nueva.
ML_STORE = {
    "sid":     None,   # sid fijo reutilizado en STORE
    "df":      None,   # DataFrame acumulado con todas las planillas
    "config":  None,
    "col_info":None,
    "files":   [],     # historial de archivos cargados
}

# ── Store de Fichas Técnicas de ML ──────────────────────────────────────────
# Mapeo { id_publicacion → categoria, sku → categoria, titulo_lower → categoria }
# Construido desde la planilla de fichas técnicas (una hoja por categoría).
FICHAS_STORE = {
    "loaded": False,
    "by_id":    {},   # "MLA123456" → "Shampoos y acondicionadores"
    "by_sku":   {},   # "SKU001"    → "Shampoos y acondicionadores"
    "by_title": {},   # "shampoo x" → "Shampoos y acondicionadores"
    "categorias": [], # lista de categorías disponibles
    "filename": None,
}

MESES = {"enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
         "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12}

# ── Store de Publicidad (ML Ads) ─────────────────────────────────────────────
PUB_STORE = {
    "campanias": None,   # DataFrame campañas
    "anuncios":  None,   # DataFrame anuncios
    "ventas_ads": None,  # DataFrame ventas por publicidad
    "files":     [],     # archivos cargados
}

# ═══════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def _str(v):
    if v is None: return None
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)): return None
    s = str(v).strip()
    return None if s in ("","nan","None","NaN","<NA>","none","-","N/A","n/a") else s

def _num(v):
    if v is None: return None
    if isinstance(v, float) and math.isnan(v): return None
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        f = float(s) if s and s not in ("-",".") else None
        return None if (f is not None and math.isnan(f)) else f
    except: return None

def _safe(v):
    if v is None: return None
    if isinstance(v, float) and math.isnan(v): return None
    try:
        if pd.isna(v): return None
    except: pass
    return v

def _fecha_ml(s):
    if not isinstance(s, str): return None
    m = re.search(r"(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})", s.lower())
    if not m: return None
    d, mes, y = int(m.group(1)), m.group(2), int(m.group(3))
    mo = MESES.get(mes)
    if not mo: return None
    hm = re.search(r"(\d{1,2}):(\d{2})", s)
    h, mi = (int(hm.group(1)), int(hm.group(2))) if hm else (0,0)
    try: return datetime(y, mo, d, h, mi)
    except: return None

def _fecha_any(s):
    """Parsea fechas en múltiples formatos."""
    if not isinstance(s, str): return None
    # Español de ML
    r = _fecha_ml(s)
    if r: return r
    # Formatos estándar
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S",
                "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%Y/%m/%d",
                "%d %b %Y", "%b %d, %Y"):
        try: return datetime.strptime(s.strip(), fmt)
        except: pass
    return None

def _to_native(df):
    """Convierte DataFrame a tipos Python nativos (elimina StringDtype/NaN de pandas 3)."""
    rows = []
    for _, row in df.iterrows():
        rows.append({str(col): _safe(val) for col, val in row.items()})
    return pd.DataFrame(rows)

def _add_time_cols(df, dt_col="_fecha_dt"):
    df["_fecha_str"] = df[dt_col].dt.strftime("%Y-%m-%d")
    df["_mes"]       = df[dt_col].dt.strftime("%Y-%m")
    df["_mes_label"] = df[dt_col].dt.strftime("%b %Y")
    df["_anio"]      = df[dt_col].dt.year.astype(str)
    df["_sem"]       = df[dt_col].dt.isocalendar().week.astype(str)
    return df

# ═══════════════════════════════════════════════════════════════════════════
# PARSER GENÉRICO — cualquier Excel/CSV
# ═══════════════════════════════════════════════════════════════════════════

def smart_detect_header(df_raw):
    """Detecta la fila de encabezados en un DF sin headers."""
    best_row, best_score = 0, -999
    for i, row in df_raw.iterrows():
        if i > 20: break
        vals = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()
                and str(v).strip() not in ("nan","None")]
        if not vals: continue
        fill = len(vals) / max(len(row), 1)
        if fill < 0.3: continue
        score = 0
        str_ratio = sum(1 for v in vals if not _num(v)) / len(vals)
        score += str_ratio * 40
        avg_len = sum(len(v) for v in vals) / len(vals)
        if avg_len < 25: score += 15
        if avg_len < 12: score += 10
        # Siguiente fila tiene números → buen indicador de header
        if i+1 < len(df_raw):
            nxt = [str(v).strip() for v in df_raw.iloc[i+1] if pd.notna(v)]
            if sum(1 for v in nxt if _num(v)) > len(nxt)*0.3: score += 20
        # "# de venta" o indicadores de ML/TN
        if any(v.lower() in ("# de venta","número de pedido","fecha","date","total","amount","monto") for v in vals):
            score += 30
        if len(vals) >= 15: score += 10
        if len(vals) == 1: score -= 25  # probable título
        if score > best_score:
            best_score, best_row = score, i
    return best_row

def read_raw_file(raw: bytes, filename: str) -> pd.DataFrame:
    """
    Lee cualquier archivo Excel o CSV y devuelve un DataFrame limpio.
    Detecta automáticamente headers, encoding y delimitador.
    """
    ext = Path(filename).suffix.lower()

    if ext in (".xlsx", ".xlsm"):
        buf = io.BytesIO(raw)
        df0 = pd.read_excel(buf, header=None, engine="openpyxl", dtype=str)
        hrow = smart_detect_header(df0)
        buf = io.BytesIO(raw)
        df = pd.read_excel(buf, header=hrow, engine="openpyxl")

    elif ext in (".xls",):
        buf = io.BytesIO(raw)
        df0 = pd.read_excel(buf, header=None, dtype=str)
        hrow = smart_detect_header(df0)
        buf = io.BytesIO(raw)
        df = pd.read_excel(buf, header=hrow)

    elif ext in (".ods",):
        buf = io.BytesIO(raw)
        df0 = pd.read_excel(buf, header=None, engine="odf", dtype=str)
        hrow = smart_detect_header(df0)
        buf = io.BytesIO(raw)
        df = pd.read_excel(buf, header=hrow, engine="odf")

    elif ext == ".csv":
        # Detectar encoding
        for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
            try: text = raw.decode(enc); break
            except: pass
        else: text = raw.decode("latin-1", errors="replace")
        # Detectar delimitador
        try:
            sample = text[:4096]
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            sep = dialect.delimiter
        except: sep = ","
        from io import StringIO
        df = pd.read_csv(StringIO(text), sep=sep)

    else:
        raise ValueError(f"Formato no soportado: {ext}. Usá .xlsx, .xls, .csv u .ods")

    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

    if df.empty:
        raise ValueError("El archivo no tiene datos.")

    return _to_native(df)

def infer_column_roles(df: pd.DataFrame) -> dict:
    """
    Infiere automáticamente el rol de cada columna.
    Retorna un dict { col_name: { type, role, sample } }
    """
    info = {}
    cols = list(df.columns)
    used_roles = set()

    for col in cols:
        if col.startswith("_"): continue
        sample = df[col].dropna().head(30)
        if sample.empty:
            info[col] = {"type":"empty","role":None,"sample":[]}
            continue

        # Detectar tipo
        vals = sample.tolist()
        str_vals = [str(v).strip() for v in vals]

        # ¿Fecha?
        date_hits = sum(1 for v in str_vals if _fecha_any(v))
        if date_hits / len(str_vals) > 0.6:
            ctype = "date"
        # ¿Número?
        elif sum(1 for v in vals if isinstance(v,(int,float)) and not (isinstance(v,float) and math.isnan(v))) / len(vals) > 0.7:
            ctype = "number"
        elif sum(1 for v in str_vals if _num(v) and v.strip()) / len(str_vals) > 0.6:
            ctype = "number"
        else:
            ctype = "text"

        # Inferir rol por nombre + tipo
        col_l = col.lower().strip()
        role = None

        if ctype == "date" and "date" not in used_roles:
            if any(x in col_l for x in ["fecha","date","cuando","time","día","dia"]):
                role = "date"; used_roles.add("date")

        if ctype == "number":
            if any(x in col_l for x in ["total","ingreso","monto","importe","precio","price","amount","revenue","venta","subtotal","pvp","facturac"]) and "amount" not in used_roles:
                role = "amount"; used_roles.add("amount")
            elif any(x in col_l for x in ["cantidad","qty","units","unidad","unidades","quantity","cant"]) and "qty" not in used_roles:
                role = "qty"; used_roles.add("qty")
            elif any(x in col_l for x in ["costo","cost","envio","shipping","descuento","discount"]) and "cost" not in used_roles:
                role = "cost"; used_roles.add("cost")

        if ctype == "text":
            if any(x in col_l for x in ["estado","status","estado del pago","payment"]) and "status" not in used_roles:
                role = "status"; used_roles.add("status")
            elif any(x in col_l for x in ["producto","product","publicacion","artículo","articulo","item","nombre del producto"]) and "product" not in used_roles:
                role = "product"; used_roles.add("product")
            elif any(x in col_l for x in ["categoria","category","tipo","type","canal","channel"]) and "category" not in used_roles:
                role = "category"; used_roles.add("category")
            elif any(x in col_l for x in ["provincia","provincia","city","ciudad","region","state","pais","country"]) and "geo" not in used_roles:
                role = "geo"; used_roles.add("geo")
            elif any(x in col_l for x in ["cliente","customer","comprador","buyer","nombre"]) and "customer" not in used_roles:
                role = "customer"; used_roles.add("customer")
            elif any(x in col_l for x in ["sku","código","code","id","número","numero","order"]) and "id" not in used_roles:
                role = "id"; used_roles.add("id")

        info[col] = {
            "type":   ctype,
            "role":   role,
            "sample": [str(v) for v in vals[:5]]
        }

    # Segunda pasada: asignar roles faltantes a columnas sin rol
    # (el usuario los puede cambiar en el modal)
    if "date" not in used_roles:
        for col, v in info.items():
            if v["type"] == "date" and not v["role"]:
                v["role"] = "date"; used_roles.add("date"); break
    if "amount" not in used_roles:
        for col, v in info.items():
            if v["type"] == "number" and not v["role"]:
                v["role"] = "amount"; used_roles.add("amount"); break
    if "product" not in used_roles:
        for col, v in info.items():
            if v["type"] == "text" and not v["role"]:
                v["role"] = "product"; used_roles.add("product"); break

    return info

def infer_column_roles_to_config(df: pd.DataFrame) -> dict:
    """Atajo: infiere roles y devuelve directamente { role: col_name }."""
    return {info["role"]: col
            for col, info in infer_column_roles(df).items()
            if info.get("role")}

def apply_config_to_df(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    """
    Aplica la configuración de columnas al DataFrame,
    creando aliases estándar que el motor de análisis entiende.
    config = { role: col_name }  p.ej. { "date":"Fecha", "amount":"Total", ... }
    """
    role_to_alias = {
        "date":     "_fecha_dt",
        "amount":   "ingresos",
        "qty":      "unidades",
        "cost":     "costo",
        "status":   "estado",
        "product":  "publicacion",
        "category": "categoria",
        "geo":      "provincia",
        "customer": "comprador",
        "id":       "id_venta",
        "category2":"categoria2",
        "amount2":  "ingresos2",
        "qty2":     "unidades2",
    }

    for role, col in config.items():
        if not col or col not in df.columns:
            continue
        alias = role_to_alias.get(role)
        if not alias:
            continue
        # Si la columna fuente ya ES el alias calculado, no hay nada que hacer
        if col == alias:
            continue

        if role == "date":
            df["_fecha_dt"] = pd.to_datetime(
                df[col].apply(lambda v: _fecha_any(str(v)) if v else None),
                errors="coerce"
            )
            # Fallback: parseo directo pandas
            mask = df["_fecha_dt"].isna() & df[col].notna()
            if mask.any():
                df.loc[mask,"_fecha_dt"] = pd.to_datetime(
                    df.loc[mask, col], dayfirst=True, errors="coerce")
            df = _add_time_cols(df)
        elif alias in ("ingresos","costo","ingresos2","unidades","unidades2","qty2","amount2"):
            # Si "ingresos" ya fue calculado por read_ml (BN × unidades), no sobreescribir
            # con el precio unitario crudo. Solo sobreescribir si la columna fuente
            # es distinta de "ingresos" (alias ya calculado) o si no existe todavía.
            if alias == "ingresos" and "ingresos" in df.columns and col != "ingresos":
                # Verificar si ya existe un valor calculado (BN × unidades)
                # Esto ocurre cuando read_ml ya pobló "ingresos" correctamente.
                # Solo re-calcular si la columna fuente no es la BN/precio unitario.
                _col_l = col.lower()
                _es_precio_unit = any(x in _col_l for x in ["precio por unidad","precio unitario","precio unit","precio x unidad"])
                if _es_precio_unit:
                    pass  # conservar el ingresos = BN × unidades ya calculado
                else:
                    df[alias] = pd.to_numeric(
                        df[col].apply(lambda v: re.sub(r"[^\d.\-]","",str(v)) if v else ""),
                        errors="coerce")
            else:
                df[alias] = pd.to_numeric(
                    df[col].apply(lambda v: re.sub(r"[^\d.\-]","",str(v)) if v else ""),
                    errors="coerce")
        else:
            df[alias] = df[col].apply(_str)

    if "_fecha_dt" not in df.columns:
        for c in ["_fecha_dt","_fecha_str","_mes","_mes_label","_anio","_sem"]:
            df[c] = None

    if "ingresos" not in df.columns:
        df["ingresos"] = 0.0

    return df

# ═══════════════════════════════════════════════════════════════════════════
# PARSERS ESPECÍFICOS ML y TN (mantienen la lógica previa)
# ═══════════════════════════════════════════════════════════════════════════

def read_ml(raw, filename):
    """
    Parser para el reporte de Mercado Libre (.xlsx).
    Formato exacto: descarga desde Mis Ventas → Exportar reporte.
    Headers en fila 6 (filas 1-5 son metadatos de ML).
    """
    buf = io.BytesIO(raw)
    try:
        df0 = pd.read_excel(buf, header=None, engine="openpyxl", dtype=str)
    except Exception:
        raise ValueError("El archivo no es un Excel válido (.xlsx).\n"
                         "El reporte de ML debe descargarse desde:\n"
                         "Mercado Libre → Mis Ventas → Exportar reporte")
    # Detectar fila de header buscando '# de venta'
    buf.seek(0)
    hrow = None
    for i, row in df0.iterrows():
        vals = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
        if any(v.lower() == "# de venta" for v in vals):
            hrow = i; break
        if len(vals) >= 15:
            hrow = i; break
    if hrow is None:
        raise ValueError(
            "No se encontró el encabezado del reporte ML.\n"
            "Asegurate de descargar el archivo desde:\n"
            "Mercado Libre → Mis Ventas → Exportar reporte")
    buf.seek(0)
    df = pd.read_excel(buf, header=hrow, engine="openpyxl")
    df = df.dropna(how="all").reset_index(drop=True)
    df = _to_native(df)

    # ── Mapeo EXACTO de columnas del reporte de ML ──────────────────────
    # Fechas (formato: "27 de marzo de 2026 22:00 hs.")
    if "Fecha de venta" in df.columns:
        df["_fecha_dt"] = pd.to_datetime(
            [_fecha_ml(v) for v in df["Fecha de venta"]], errors="coerce")
        df = _add_time_cols(df)

    # ── Columna BN: Precio por unidad (columna 66 del Excel de ML) ──────
    # Intentar por nombre primero; si no, por posición (índice 65 = columna BN)
    PRECIO_UNIT_NAMES = ["Precio por unidad", "Precio por Unidad", "Precio unitario",
                         "Precio Unit.", "precio_por_unidad",
                         "Precio x unidad", "precio x unidad", "Precio X Unidad",
                         "precio x Unidad", "PRECIO X UNIDAD"]
    _col_precio_unit = None
    for pname in PRECIO_UNIT_NAMES:
        if pname in df.columns:
            _col_precio_unit = pname
            break
    # Fallback: usar posición BN (índice 65) si existe
    if _col_precio_unit is None:
        _raw_cols = list(df.columns)
        if len(_raw_cols) >= 66:
            _col_precio_unit = _raw_cols[65]  # índice 65 = columna BN

    # Numéricos
    NUM_ML = {
        "Ingresos por productos (ARS)": "ingresos_orig",  # guardamos el original pero no lo usamos como principal
        "Total (ARS)":                  "total_neto",
        "Unidades":                     "unidades",
        "Cargo por venta":              "cargo_venta",
        "Costo fijo":                   "costo_fijo",
        "Costos de envío (ARS)":        "costo_envio",
        "Costos de envio (ARS)":        "costo_envio",
        "Ingresos por envío (ARS)":     "ingreso_envio",
        "Ingresos por envio (ARS)":     "ingreso_envio",
        "Impuestos":                    "impuestos",
        "Descuentos":                   "descuentos",
        "Anulaciones y reembolsos (ARS)": "anulaciones",
    }
    for orig, alias in NUM_ML.items():
        if orig in df.columns and alias not in df.columns:
            df[alias] = pd.to_numeric(
                df[orig].apply(lambda v: re.sub(r"[^\d.\-]","",str(v)) if v else ""),
                errors="coerce")

    # ── Calcular ingresos = Precio por unidad (suma directa, sin × unidades) ─
    if _col_precio_unit and _col_precio_unit in df.columns:
        df["precio_por_unidad"] = pd.to_numeric(
            df[_col_precio_unit].apply(lambda v: re.sub(r"[^\d.\-]","",str(v)) if v else ""),
            errors="coerce")
        df["ingresos"] = df["precio_por_unidad"].fillna(0).round(2)
    elif "ingresos_orig" in df.columns:
        # Fallback: usar columna original si no se encontró BN
        df["ingresos"] = df["ingresos_orig"]

    # Categóricos
    CAT_ML = {
        "# de venta":                       "id_venta",
        "Estado":                           "estado",
        "Descripción del estado":           "desc_estado",
        "Descripcion del estado":           "desc_estado",
        "Título de la publicación":         "publicacion",
        "Titulo de la publicacion":         "publicacion",
        "SKU":                              "sku",
        "Variante":                         "variante",
        "Canal de venta":                   "canal_ml",
        "Forma de entrega":                 "forma_envio",
        "Transportista":                    "transportista",
        "Ciudad":                           "ciudad",
        "Estado.1":                         "provincia",
        "País":                             "pais",
        "Comprador":                        "comprador",
        "Tienda oficial":                   "tienda",
        "# de publicación":                 "id_publicacion",
        "Venta por publicidad":             "por_publicidad",
    }
    for orig, alias in CAT_ML.items():
        if orig in df.columns and alias not in df.columns:
            df[alias] = df[orig].apply(_str)

    # Calcular "costo" consolidado (cargo venta + costo envío)
    import pandas as _pd
    c1 = df["cargo_venta"].fillna(0) if "cargo_venta" in df.columns else _pd.Series([0]*len(df))
    c2 = df["costo_envio"].fillna(0) if "costo_envio" in df.columns else _pd.Series([0]*len(df))
    df["costo"] = (c1.abs() + c2.abs())

    df["fuente"] = "Mercado Libre"
    # Normalizar estados ML
    ESTADOS_ML = {
        "entregado": "Entregado",
        "en camino": "En camino",
        "procesando": "En proceso",
        "cancelad":  "Cancelado",
        "reclamo":   "Reclamo",
        "paquete":   "Paquete múltiple",
    }
    if "estado" in df.columns:
        def norm_ml(v):
            if not v: return None
            sl = v.lower()
            for k, label in ESTADOS_ML.items():
                if k in sl: return label
            return v
        df["estado"] = df["estado"].apply(norm_ml)

    # Reportar qué columna se usó para ingresos
    # Si se calculó BN × unidades, apuntamos a "ingresos" (alias calculado)
    # para que apply_config_to_df no sobreescriba el valor correcto.
    _ingreso_col = "ingresos" if (_col_precio_unit and "ingresos" in df.columns) \
                   else ("ingresos_orig" if "ingresos_orig" in df.columns else "Ingresos por productos (ARS)")
    config = {"date":"Fecha de venta","amount": _ingreso_col,
              "qty":"Unidades","status":"Estado","product":"Título de la publicación",
              "geo":"Estado.1","customer":"Comprador","id":"# de venta","category":"Forma de entrega"}
    return df, config


def read_tn(raw, filename):
    """
    Parser para el reporte de Tienda Nube (.csv).
    Formato exacto: Panel TN → Estadísticas → Exportar pedidos.
    Encoding: latin-1 | Separador: punto y coma (;)
    
    Columnas principales del export real:
      Número de orden, Fecha, Estado de la orden, Estado del pago,
      Estado del envío, Total, Subtotal de productos, Descuento,
      Costo de envío, Nombre del comprador, Nombre del producto,
      Precio del producto, Cantidad del producto, SKU,
      Provincia o estado, Ciudad, Medio de envío, Medio de pago, Canal
    """
    # Detectar encoding (TN usa latin-1 en Argentina)
    for enc in ("latin-1", "cp1252", "utf-8-sig", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw.decode("latin-1", errors="replace")

    from io import StringIO
    # Detectar separador automáticamente (TN Argentina usa ";" pero puede venir con ",")
    sample = text[:4096]
    try:
        import csv as _csv
        dialect = _csv.Sniffer().sniff(sample, delimiters=";,	|")
        sep = dialect.delimiter
    except Exception:
        # Fallback: contar ocurrencias
        sep = ";" if text[:2048].count(";") >= text[:2048].count(",") else ","

    df = pd.read_csv(StringIO(text), sep=sep, dtype=str, na_filter=False)
    df = df.replace("", None)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    df = _to_native(df)

    if df.empty:
        raise ValueError("El CSV de Tienda Nube no tiene datos.")

    # ── Detección de formato: verificar si tiene columnas típicas de TN ──
    # Incluir variantes con caracteres corruptos por encoding
    TN_KEY_COLS = {"Número de orden", "Fecha", "Total", "Estado del pago",
                   "Nombre del producto", "Nombre del comprador",
                   "NM-zmero de orden", "Estado del pago", "Nombre del comprador"}
    cols_set = set(df.columns)
    es_formato_tn = len(TN_KEY_COLS & cols_set) >= 3

    if not es_formato_tn:
        # CSV genérico subido como TN: usar parser genérico y etiquetar como Tienda Nube
        df2 = apply_config_to_df(df.copy(), infer_column_roles_to_config(df))
        df2["fuente"] = "Tienda Nube"
        if "_fecha_dt" not in df2.columns:
            for c in ["_fecha_dt","_fecha_str","_mes","_mes_label","_anio","_sem"]:
                df2[c] = None
        if "ingresos" not in df2.columns:
            df2["ingresos"] = 0.0
        config2 = {info["role"]: col
                   for col, info in infer_column_roles(df).items()
                   if info.get("role")}
        return df2, config2

    # ── Fechas (formato TN: "DD/MM/YYYY HH:MM:SS") ──────────────────────
    FC = "Fecha"  # columna de fecha en el CSV de TN
    if FC in df.columns:
        df["_fecha_dt"] = pd.to_datetime(df[FC], format="%d/%m/%Y %H:%M:%S", errors="coerce")
        # Fallback para filas con formato incompleto
        mask = df["_fecha_dt"].isna() & df[FC].notna()
        if mask.any():
            df.loc[mask, "_fecha_dt"] = pd.to_datetime(
                df.loc[mask, FC], dayfirst=True, errors="coerce")
        df = _add_time_cols(df)

    # ── Numéricos (los valores ya vienen como número o "0.00") ──────────
    NUM_TN = {
        "Total":                  "ingresos",
        "Subtotal de productos":  "subtotal",
        "Descuento":              "descuentos",
        "Costo de envío":         "costo_envio",
        "Precio del producto":    "precio_unit",
        "Cantidad del producto":  "unidades",
    }
    for orig, alias in NUM_TN.items():
        if orig in df.columns and alias not in df.columns:
            df[alias] = pd.to_numeric(
                df[orig].apply(lambda v: re.sub(r"[^\d.\-]","",str(v)) if v else ""),
                errors="coerce")

    # total_neto = Total - Costo de envío
    if "ingresos" in df.columns and "costo_envio" in df.columns:
        df["total_neto"] = df["ingresos"].fillna(0) - df["costo_envio"].fillna(0)
    elif "ingresos" in df.columns:
        df["total_neto"] = df["ingresos"]

    # Costo consolidado
    df["costo"] = df.get("costo_envio", pd.Series([0]*len(df))).fillna(0)

    # ── Categóricos ──────────────────────────────────────────────────────
    CAT_TN = {
        "Número de orden":   "id_venta",
        "Estado del pago":   "estado",         # Recibido, Pendiente, Reembolsado...
        "Estado de la orden":"estado_orden",
        "Estado del envío":  "estado_envio",
        "Nombre del comprador": "comprador",
        "Nombre del producto":  "publicacion",
        "SKU":               "sku",
        "Provincia o estado":"provincia",
        "Ciudad":            "ciudad",
        "País":              "pais",
        "Medio de envío":    "forma_envio",
        "Medio de pago":     "forma_pago",
        "Canal":             "canal_tn",        # Web, Móvil, Admin...
        "Email":             "email",
        "Localidad":         "localidad",
        "Cupón de descuento":"cupon",
        "Código de tracking del envío": "tracking",
    }
    for orig, alias in CAT_TN.items():
        if orig in df.columns and alias not in df.columns:
            df[alias] = df[orig].apply(_str)

    # ── Normalizar estados de pago de TN ─────────────────────────────────
    # Los estados reales del CSV: Recibido, Reembolsado, Pendiente,
    # Vencido, Rechazado, Parcialmente reembolsado
    ESTADOS_TN = {
        "recibido":                   "Pagado",
        "pagado":                     "Pagado",
        "pendiente":                  "Pendiente",
        "vencido":                    "Vencido",
        "rechazado":                  "Rechazado",
        "reembolsado":                "Reembolsado",
        "parcialmente reembolsado":   "Reembolso parcial",
        "cancelad":                   "Cancelado",
    }
    if "estado" in df.columns:
        def norm_tn(v):
            if not v or not isinstance(v, str): return None
            sl = v.lower().strip()
            for k, label in ESTADOS_TN.items():
                if k in sl: return label
            return v
        df["estado"] = df["estado"].apply(norm_tn)

    # Categoría de envío (simplificar nombre largo del medio)
    if "forma_envio" in df.columns:
        def simp_envio(v):
            if not v or not isinstance(v, str): return None
            v = v.strip()
            if not v: return None
            if "andreani" in v.lower(): return "Andreani"
            if "correo" in v.lower(): return "Correo Argentino"
            if "retiro" in v.lower(): return "Punto de retiro"
            if "occa" in v.lower() or "oca" in v.lower(): return "OCA"
            return v.split('"')[0].strip() or v
        df["forma_envio"] = df["forma_envio"].apply(simp_envio)

    df["fuente"] = "Tienda Nube"

    config = {"date":"Fecha","amount":"Total","qty":"Cantidad del producto",
              "status":"Estado del pago","product":"Nombre del producto",
              "geo":"Provincia o estado","customer":"Nombre del comprador",
              "id":"Número de orden","category":"Medio de envío"}
    return df, config

# ═══════════════════════════════════════════════════════════════════════════
# FICHAS TÉCNICAS DE MERCADO LIBRE
# ═══════════════════════════════════════════════════════════════════════════

def read_fichas(raw: bytes, filename: str) -> dict:
    """
    Lee una planilla de Fichas Técnicas de Mercado Libre.
    Cada hoja (excepto 'Ayuda' y 'hidden') corresponde a una categoría.
    Fila 5 (índice 4) = encabezados reales (FAMILY_ID, ID, SKU, TITLE, …).
    Fila 6+ = datos de publicaciones.
    Devuelve un dict con los índices by_id, by_sku, by_title y la lista de categorías.
    """
    buf = io.BytesIO(raw)
    from openpyxl import load_workbook as _lwb
    wb = _lwb(buf, read_only=True, data_only=True)

    SKIP_SHEETS = {"ayuda", "hidden"}
    by_id    = {}
    by_sku   = {}
    by_title = {}
    categorias = []

    for sheet_name in wb.sheetnames:
        if sheet_name.lower().strip() in SKIP_SHEETS:
            continue

        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 5:
            continue

        # La fila 4 (índice 3) contiene los encabezados técnicos (FAMILY_ID, ID, SKU, TITLE…)
        # La fila 3 (índice 2) puede ser el título de la hoja con (*) Campos requeridos
        # Buscamos la fila de headers: primera con 'ID' o 'TITLE' en alguna celda
        header_row_idx = None
        for i, row in enumerate(rows[:8]):
            vals = [str(v).strip().upper() if v else "" for v in row]
            if "ID" in vals and ("TITLE" in vals or "SKU" in vals):
                header_row_idx = i
                break
        if header_row_idx is None:
            continue

        headers = [str(v).strip().upper() if v else "" for v in rows[header_row_idx]]

        # Indices de columnas clave
        def col_idx(*names):
            for n in names:
                if n in headers:
                    return headers.index(n)
            return None

        idx_id    = col_idx("ID")            # ID de publicación (MLA...)
        idx_sku   = col_idx("SKU")
        idx_title = col_idx("TITLE")

        if idx_id is None and idx_sku is None and idx_title is None:
            continue

        # Nombre de categoría = nombre de la hoja (limpiado)
        cat = sheet_name.strip()
        categorias.append(cat)

        # Procesar filas de datos (a partir de header_row_idx + 1)
        # Salteamos las filas de tipo/descripción que tienen "FIXED", "VARIATION", etc.
        TYPE_SKIP = {"fixed", "variation", "attribute", "attribute_unit", ""}
        for row in rows[header_row_idx + 1:]:
            if not row:
                continue
            # Detectar fila de metadatos (FIXED/VARIATION/ATTRIBUTE)
            first_vals = [str(v).strip().lower() if v else "" for v in row[:5]]
            if all(v in TYPE_SKIP for v in first_vals if v):
                continue
            # Saltar filas de encabezado descriptivo (tienen texto largo con \n)
            if idx_id is not None and idx_id < len(row):
                raw_id = str(row[idx_id]).strip() if row[idx_id] else ""
                # Si el valor en la columna ID empieza con MLA, es una fila real
                if raw_id.upper().startswith("MLA"):
                    by_id[raw_id.upper()] = cat
            if idx_sku is not None and idx_sku < len(row):
                raw_sku = str(row[idx_sku]).strip() if row[idx_sku] else ""
                if raw_sku and raw_sku not in ("", "None", "nan"):
                    by_sku[raw_sku.upper()] = cat
            if idx_title is not None and idx_title < len(row):
                raw_title = str(row[idx_title]).strip() if row[idx_title] else ""
                if raw_title and len(raw_title) > 3 and raw_title.upper() not in ("TITLE","TÍTULO"):
                    by_title[raw_title.lower()] = cat

    wb.close()
    return {
        "by_id":     by_id,
        "by_sku":    by_sku,
        "by_title":  by_title,
        "categorias": categorias,
    }


def enrich_categoria_from_fichas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Intenta asignar la columna 'categoria' en el DataFrame usando FICHAS_STORE.
    Estrategia (en orden de prioridad):
      1. id_publicacion → by_id
      2. sku            → by_sku
      3. publicacion    → by_title (título de la publicación)
    Solo sobreescribe si la categoría actual es vacía/None o es el valor por defecto de ML ('Forma de entrega').
    """
    if not FICHAS_STORE["loaded"]:
        return df

    by_id    = FICHAS_STORE["by_id"]
    by_sku   = FICHAS_STORE["by_sku"]
    by_title = FICHAS_STORE["by_title"]

    if "categoria" not in df.columns:
        df["categoria"] = None

    def _lookup(row):
        current = str(row.get("categoria", "") or "").strip()
        # Solo enriquecer si está vacía o es el placeholder de ML
        if current and current not in ("", "None", "nan", "Forma de entrega"):
            return current

        # 1. Por ID de publicación
        pid = str(row.get("id_publicacion", "") or "").strip().upper()
        if pid and pid in by_id:
            return by_id[pid]

        # 2. Por SKU
        sku = str(row.get("sku", "") or "").strip().upper()
        if sku and sku in by_sku:
            return by_sku[sku]

        # 3. Por título (coincidencia exacta lowercase)
        title = str(row.get("publicacion", "") or "").strip().lower()
        if title and title in by_title:
            return by_title[title]

        return current or None

    df["categoria"] = df.apply(_lookup, axis=1)
    return df


# ═══════════════════════════════════════════════════════════════════════════
# MOTOR DE ANÁLISIS (compartido por todas las fuentes)
# ═══════════════════════════════════════════════════════════════════════════

def _S(df, col):
    if col not in df.columns: return 0.0
    return float(df[col].sum(skipna=True))

def _R(v, d=2):
    if v is None or (isinstance(v,float) and math.isnan(v)): return 0
    return round(v, d)

def metricas(df):
    n   = len(df)
    ing = _S(df,"ingresos")
    ent = 0
    mask_ent = pd.Series([False]*len(df), index=df.index)
    if "estado" in df.columns:
        estados_ok = ["entregado","pagado","paid","delivered","completad","complete","aprobado"]
        mask_ent = df["estado"].fillna("").str.lower().apply(
            lambda s: any(x in s for x in estados_ok))
        ent = mask_ent.sum()
    df_ent = df[mask_ent]
    ing_ent = _S(df_ent,"ingresos")
    uni_ent = _R(_S(df_ent,"unidades"), 0)
    return {
        "n_ventas":          n,
        "unidades":          _R(_S(df,"unidades"), 0),
        "ingresos":          _R(ing),
        "total_neto":        _R(_S(df,"total_neto") if "total_neto" in df.columns else ing),
        "costo":             _R(abs(_S(df,"costo"))),
        "descuentos":        _R(abs(_S(df,"descuentos"))) if "descuentos" in df.columns else 0,
        "ticket_prom":       _R(ing/n) if n>0 else 0,
        "tasa_ok":           _R(int(ent)/n*100,1) if n>0 else 0,
        "ingresos_entregado":_R(ing_ent),
        "unidades_entregado":int(uni_ent),
        "n_entregado":       int(ent),
    }

def grp(df, by, top=None):
    if by not in df.columns: return []
    num_cols = [c for c in ["ingresos","unidades"] if c in df.columns]
    tmp = df[[by]+num_cols].copy()
    tmp[by] = tmp[by].fillna("(vacío)")
    for c in num_cols:
        tmp[c] = pd.to_numeric(tmp[c], errors="coerce").fillna(0)
    agg = tmp.groupby(by, as_index=False)[num_cols].sum()
    agg["cantidad"] = tmp.groupby(by)[by].count().values
    agg = agg.sort_values("ingresos", ascending=False)
    if top: agg = agg.head(top)
    tot = agg["ingresos"].sum()
    agg["pct"] = (agg["ingresos"]/tot*100).round(1) if tot else 0
    agg["ingresos"] = agg["ingresos"].round(2)
    if "unidades" in agg.columns:
        agg["unidades"] = agg["unidades"].round(0).astype(int)
    return agg.rename(columns={by:"label"}).to_dict("records")

def grp_custom(df, by_col, val_col, agg_fn="sum", top=None):
    """Agrupa por cualquier columna con cualquier función."""
    if by_col not in df.columns or val_col not in df.columns: return []
    tmp = df[[by_col, val_col]].copy()
    tmp[by_col] = tmp[by_col].fillna("(vacío)").astype(str)
    tmp[val_col] = pd.to_numeric(tmp[val_col], errors="coerce").fillna(0)
    if agg_fn == "count":
        agg = tmp.groupby(by_col).size().reset_index(name="valor")
    elif agg_fn == "avg":
        agg = tmp.groupby(by_col, as_index=False)[val_col].mean().rename(columns={val_col:"valor"})
    elif agg_fn == "max":
        agg = tmp.groupby(by_col, as_index=False)[val_col].max().rename(columns={val_col:"valor"})
    elif agg_fn == "min":
        agg = tmp.groupby(by_col, as_index=False)[val_col].min().rename(columns={val_col:"valor"})
    else:
        agg = tmp.groupby(by_col, as_index=False)[val_col].sum().rename(columns={val_col:"valor"})
    agg = agg.sort_values("valor", ascending=False)
    if top: agg = agg.head(top)
    tot = agg["valor"].sum()
    agg["pct"] = (agg["valor"]/tot*100).round(1) if tot else 0
    agg["valor"] = agg["valor"].round(2)
    return agg.rename(columns={by_col:"label"}).to_dict("records")


def flex_por_zona(df, top=15):
    """
    Filtra envíos Flex (por forma_envio o transportista) y agrupa por provincia/zona.
    Devuelve lista con: zona, cantidad, pct_cantidad, ingresos, pct_ingresos
    """
    if df.empty:
        return {"zonas": [], "total_flex": 0, "total_envios": len(df)}

    # Detectar columna de envío (forma_envio tiene prioridad, luego transportista)
    col_envio = None
    for c in ["forma_envio", "transportista", "canal_ml"]:
        if c in df.columns:
            col_envio = c
            break
    if col_envio is None:
        return {"zonas": [], "total_flex": 0, "total_envios": len(df)}

    # Filtrar filas Flex
    mask_flex = df[col_envio].fillna("").astype(str).str.lower().str.contains("flex", na=False)
    df_flex = df[mask_flex].copy()

    total_envios = len(df)
    total_flex   = len(df_flex)

    if total_flex == 0:
        return {"zonas": [], "total_flex": 0, "total_envios": total_envios}

    # Agrupar por la columna más granular disponible:
    # ciudad > localidad > provincia
    # Construir etiqueta combinada ciudad + provincia cuando sea posible
    geo_col = None
    for c in ["ciudad", "localidad", "provincia"]:
        if c in df_flex.columns and df_flex[c].notna().any():
            geo_col = c
            break
    if geo_col is None:
        return {"zonas": [], "total_flex": total_flex, "total_envios": total_envios}

    # Crear etiqueta "Ciudad (Provincia)" para mayor contexto
    if geo_col == "ciudad" and "provincia" in df_flex.columns:
        df_flex = df_flex.copy()
        df_flex["_zona_label"] = df_flex["ciudad"].fillna("(sin ciudad)").astype(str).str.strip() +             " (" + df_flex["provincia"].fillna("").astype(str).str.strip() + ")"
        df_flex["_zona_label"] = df_flex["_zona_label"].str.replace(r" \(\)$", "", regex=True)
        grp_col = "_zona_label"
    else:
        grp_col = geo_col

    agg = df_flex.groupby(grp_col, as_index=False).agg(
        cantidad=(grp_col, "count")
    )
    if "ingresos" in df_flex.columns:
        ing_agg = df_flex.groupby(grp_col, as_index=False)["ingresos"].sum()
        agg = agg.merge(ing_agg, on=grp_col, how="left")
    else:
        agg["ingresos"] = 0.0

    agg = agg.sort_values("cantidad", ascending=False).head(top)
    tot_cant = agg["cantidad"].sum()
    tot_ing  = agg["ingresos"].sum() if "ingresos" in agg.columns else 0

    zonas = []
    for _, row in agg.iterrows():
        zonas.append({
            "zona":          str(row[grp_col]),
            "cantidad":      int(row["cantidad"]),
            "pct_cantidad":  round(row["cantidad"] / tot_cant * 100, 1) if tot_cant else 0,
            "ingresos":      round(float(row.get("ingresos", 0)), 2),
            "pct_ingresos":  round(float(row.get("ingresos", 0)) / tot_ing * 100, 1) if tot_ing else 0,
        })

    # ── Agregar coordenadas para el mapa de calor ──────────────────────
    # Diccionario de ciudades/barrios argentinos con coordenadas
    GEO_AR = {
        # CABA barrios
        "palermo":(-34.5755,-58.4268),"belgrano":(-34.5606,-58.4584),
        "caballito":(-34.6182,-58.4378),"flores":(-34.6273,-58.4621),
        "villa urquiza":(-34.5758,-58.4898),"recoleta":(-34.5875,-58.3928),
        "balvanera":(-34.6095,-58.4079),"almagro":(-34.6145,-58.4267),
        "villa del parque":(-34.5991,-58.4850),"montserrat":(-34.6155,-58.3808),
        "san telmo":(-34.6245,-58.3697),"la boca":(-34.6348,-58.3620),
        "barracas":(-34.6454,-58.3829),"boedo":(-34.6348,-58.4174),
        "villa crespo":(-34.5978,-58.4452),"chacarita":(-34.5854,-58.4549),
        "colegiales":(-34.5739,-58.4474),"devoto":(-34.5994,-58.5109),
        "floresta":(-34.6273,-58.4886),"liniers":(-34.6386,-58.5204),
        "mataderos":(-34.6636,-58.5115),"monte castro":(-34.6091,-58.5070),
        "nueva pompeya":(-34.6548,-58.4108),"parque chacabuco":(-34.6395,-58.4428),
        "parque patricios":(-34.6418,-58.4044),"paternal":(-34.6036,-58.4697),
        "saavedra":(-34.5557,-58.4932),"villa luro":(-34.6386,-58.4904),
        "villa pueyrredon":(-34.5879,-58.5048),"villa real":(-34.6283,-58.5069),
        "villa riachuelo":(-34.6734,-58.4566),"villa santa rita":(-34.6136,-58.4878),
        "villa soldati":(-34.6683,-58.4370),"agronomia":(-34.5994,-58.4975),
        "nuñez":(-34.5450,-58.4572),"versalles":(-34.6333,-58.5175),
        "parque avellaneda":(-34.6500,-58.4811),"puerto madero":(-34.6151,-58.3620),
        # GBA
        "avellaneda":(-34.6611,-58.3651),"quilmes":(-34.7218,-58.2536),
        "lanus":(-34.7052,-58.3924),"lomas de zamora":(-34.7606,-58.3981),
        "almirante brown":(-34.8200,-58.3800),"esteban echeverria":(-34.8167,-58.4500),
        "ezeiza":(-34.8526,-58.5161),"merlo":(-34.6826,-58.7272),
        "morón":(-34.6516,-58.6194),"moron":(-34.6516,-58.6194),
        "ituzaingo":(-34.6581,-58.6766),"hurlingham":(-34.5892,-58.6402),
        "tres de febrero":(-34.6100,-58.5600),"san martin":(-34.5747,-58.5367),
        "general san martin":(-34.5747,-58.5367),"la matanza":(-34.7704,-58.6222),
        "san justo":(-34.6847,-58.5606),"ramos mejia":(-34.6448,-58.5629),
        "haedo":(-34.6478,-58.5944),"villa tesei":(-34.6286,-58.6469),
        "ciudadela":(-34.6414,-58.5429),"tapiales":(-34.6792,-58.5553),
        "gonzales catan":(-34.7653,-58.6392),"isidro casanova":(-34.7250,-58.6083),
        "gregorio de laferrere":(-34.7450,-58.5861),"rafael castillo":(-34.7128,-58.6408),
        "virrey del pino":(-34.8108,-58.6372),"temperley":(-34.7703,-58.3978),
        "turdera":(-34.7917,-58.3889),"banfield":(-34.7406,-58.3941),
        "remedios de escalada":(-34.7539,-58.4025),"lomas":(-34.7606,-58.3981),
        "monte grande":(-34.8186,-58.4703),"canning":(-34.8803,-58.5025),
        "longchamps":(-34.8736,-58.3858),"glew":(-34.8972,-58.3747),
        "malvinas argentinas":(-34.4525,-58.7011),"don torcuato":(-34.4747,-58.6353),
        "tigre":(-34.4258,-58.5786),"el talar":(-34.4528,-58.6814),
        "benavidez":(-34.4022,-58.6897),"nordelta":(-34.3978,-58.6467),
        "escobar":(-34.3481,-58.7967),"pilar":(-34.4585,-58.9139),
        "del viso":(-34.4125,-58.8247),"maquinista savio":(-34.3900,-58.6531),
        "garín":(-34.4228,-58.7258),"garin":(-34.4228,-58.7258),
        "villa del lago":(-34.3633,-58.7633),"jose c paz":(-34.5228,-58.7756),
        "san miguel":(-34.5453,-58.7083),"bella vista":(-34.5681,-58.6878),
        "grand bourg":(-34.4978,-58.7281),"tortuguitas":(-34.4539,-58.7519),
        "general pacheco":(-34.4519,-58.6533),"ingeniero benavides":(-34.3694,-58.6789),
        "la lucila":(-34.4983,-58.5250),"martinez":(-34.4892,-58.5053),
        "vicente lopez":(-34.5269,-58.4778),"florida":(-34.5267,-58.5036),
        "munro":(-34.5300,-58.5203),"villa martelli":(-34.5567,-58.5369),
        "olivos":(-34.5097,-58.4917),"boulogne":(-34.4942,-58.5647),
        "la lonja":(-34.4419,-58.5722),"san isidro":(-34.4683,-58.5253),
        "beccar":(-34.4769,-58.5389),"acassuso":(-34.4808,-58.5125),
        "berazategui":(-34.7654,-58.2101),"florencio varela":(-34.8181,-58.2764),
        "quilmes oeste":(-34.7411,-58.2764),"bernal":(-34.7025,-58.2817),
        "ezpeleta":(-34.7528,-58.2314),"el jaguel":(-34.8425,-58.4539),
        "canning":(-34.8803,-58.5025),"spegazzini":(-34.9056,-58.5039),
        "tristan suarez":(-34.8608,-58.5508),"ezeiza":(-34.8526,-58.5161),
        "la plata":(-34.9206,-57.9544),"la plata (buenos aires)":(-34.9206,-57.9544),
        "berisso":(-34.8745,-57.8908),"ensenada":(-34.8589,-57.9108),
        "gonnet":(-34.8836,-57.9950),"city bell":(-34.8703,-58.0264),
        "villa elisa":(-34.8553,-58.0939),"manuel b gonnet":(-34.8836,-57.9950),
        "mar del plata":(-37.9995,-57.5575),"mar del plata (buenos aires)":(-37.9995,-57.5575),
        "cordoba":(-31.4201,-64.1888),"córdoba":(-31.4201,-64.1888),
        "rosario":(-32.9442,-60.6505),"santa fe":(-31.6333,-60.7000),
        "mendoza":(-32.8908,-68.8272),"tucuman":(-26.8083,-65.2176),
        "tucumán":(-26.8083,-65.2176),"salta":(-24.7833,-65.4167),
        "corrientes":(-27.4667,-58.8333),"resistencia":(-27.4500,-58.9833),
        "posadas":(-27.3667,-55.9000),"neuquén":(-38.9517,-68.0591),
        "neuquen":(-38.9517,-68.0591),"bariloche":(-41.1333,-71.3000),
        "bahia blanca":(-38.7183,-62.2663),"bahía blanca":(-38.7183,-62.2663),
        # Fallback provincias
        "buenos aires":(-34.6037,-58.3816),
        "capital federal":(-34.6037,-58.3816),
        "caba":(-34.6037,-58.3816),
    }

    def get_coords(zona_str):
        """Busca coordenadas para una zona por nombre."""
        z = zona_str.lower().strip()
        # Buscar exacto
        if z in GEO_AR: return GEO_AR[z]
        # Buscar sin paréntesis (ciudad sin provincia)
        z_base = z.split("(")[0].strip()
        if z_base in GEO_AR: return GEO_AR[z_base]
        # Buscar si alguna clave está contenida en z
        for key, coords in GEO_AR.items():
            if key in z or z in key:
                return coords
        return None

    for z in zonas:
        coords = get_coords(z["zona"])
        if coords:
            z["lat"] = coords[0]
            z["lng"] = coords[1]

    return {
        "zonas":        zonas,
        "total_flex":   total_flex,
        "total_envios": total_envios,
        "pct_flex_total": round(total_flex / total_envios * 100, 1) if total_envios else 0,
    }


def por_tiempo(df, periodo="dia", val_col="ingresos"):
    key = {"mes":"_mes","semana":"_sem","anio":"_anio"}.get(periodo,"_fecha_str")
    if key not in df.columns: return []
    num_cols = [c for c in [val_col,"unidades"] if c in df.columns]
    if not num_cols: return []
    tmp = df[[key]+num_cols].copy()
    tmp = tmp.dropna(subset=[key])
    tmp[key] = tmp[key].astype(str)
    for c in num_cols: tmp[c] = pd.to_numeric(tmp[c], errors="coerce").fillna(0)
    tmp["cantidad"] = 1
    agg = tmp.groupby(key, as_index=False)[num_cols+["cantidad"]].sum()
    agg = agg.sort_values(key).rename(columns={key:"periodo",val_col:"ingresos"})
    agg["ingresos"] = agg["ingresos"].round(2)
    return agg.to_dict("records")

def tabla_ventas(df, page=0, ps=50):
    prefer = ["id_venta","_fecha_str","publicacion","estado","fuente","categoria",
              "unidades","ingresos","total_neto","provincia","comprador"]
    cols = [c for c in prefer if c in df.columns]
    # Agregar cols numéricas extra que no estén ya
    for c in df.columns:
        if c not in cols and not c.startswith("_") and c in df.columns:
            if len(cols) < 14: cols.append(c)
    tmp = df[cols].copy()
    if "ingresos" in tmp.columns:
        tmp = tmp.sort_values("ingresos", ascending=False, na_position="last")
    total = len(tmp)
    rows = []
    for _, row in tmp.iloc[page*ps:(page+1)*ps].iterrows():
        r = {}
        for c in cols:
            v = row[c]
            if v is None or (isinstance(v,float) and math.isnan(v)): r[c]=None
            elif isinstance(v,float) and v==int(v): r[c]=int(v)
            elif isinstance(v,(float,int)): r[c]=round(float(v),2)
            else: r[c]=str(v)
        rows.append(r)
    return {"rows":rows,"total":total,"cols":cols}

def filtros_disp(df):
    def u(col):
        if col not in df.columns: return []
        return sorted(set(str(v).strip() for v in df[col].dropna()
                         if str(v).strip() not in ("","nan","None")))
    fd=fm=None
    meses=[]
    if "_fecha_dt" in df.columns:
        dts=df["_fecha_dt"].dropna()
        if len(dts): fd=str(dts.min().date()); fm=str(dts.max().date())
    if "_mes" in df.columns:
        meses = sorted(set(str(v).strip() for v in df["_mes"].dropna()
                           if str(v).strip() not in ("","nan","None")))
    # Columnas categóricas disponibles para filtrar
    cat_cols = [c for c in df.columns
                if not c.startswith("_") and df[c].dtype == object
                and df[c].nunique() <= 200 and df[c].nunique() >= 2]
    # Columnas numéricas disponibles para visualización
    num_cols = [c for c in df.columns
                if not c.startswith("_") and pd.api.types.is_numeric_dtype(df[c])
                and c not in ("_fecha_dt",)]
    return {
        "estado":     u("estado"),
        "fuente":     u("fuente"),
        "categoria":  u("categoria"),
        "provincia":  u("provincia"),
        "forma_envio":u("forma_envio") if "forma_envio" in df.columns else [],
        "fecha_min":  fd,
        "fecha_max":  fm,
        "meses":      meses,
        "cat_cols":   cat_cols[:30],  # columnas disponibles para filtro dinámico
        "num_cols":   num_cols[:30],  # columnas numéricas para visualización
        "all_cols":   [c for c in df.columns if not c.startswith("_")][:50],
    }

def apply_filters(df, f):
    mask = pd.Series([True]*len(df), index=df.index)
    for col, val in [("estado", f.get("estado","__all__")),
                     ("fuente", f.get("fuente","__all__")),
                     ("categoria", f.get("categoria","__all__")),
                     ("provincia", f.get("provincia","__all__"))]:
        if val and val != "__all__" and col in df.columns:
            mask &= df[col].fillna("") == val
    # Filtro por mes (YYYY-MM)
    mes = f.get("mes","__all__")
    if mes and mes != "__all__" and "_mes" in df.columns:
        mask &= df["_mes"].fillna("") == mes
    # Filtro dinámico extra (custom_col + custom_val) — backward compat
    if f.get("custom_col") and f.get("custom_val") and f["custom_col"] in df.columns:
        mask &= df[f["custom_col"]].fillna("").astype(str) == f["custom_val"]
    # Múltiples filtros dinámicos [{col, val}]
    for flt in (f.get("dynamic_filters") or []):
        col = flt.get("col"); val = flt.get("val")
        if col and val and val != "__all__" and col in df.columns:
            mask &= df[col].fillna("").astype(str) == val
    if f.get("fecha_desde") and "_fecha_dt" in df.columns:
        try: mask &= df["_fecha_dt"] >= pd.to_datetime(f["fecha_desde"])
        except: pass
    if f.get("fecha_hasta") and "_fecha_dt" in df.columns:
        try:
            # Incluir el día completo sumando 1 día (hasta las 23:59:59)
            hasta = pd.to_datetime(f["fecha_hasta"]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            mask &= df["_fecha_dt"] <= hasta
        except: pass
    if f.get("texto"):
        txt = f["texto"].lower()
        tmask = pd.Series([False]*len(df), index=df.index)
        for col in df.select_dtypes(include="object").columns:
            if col.startswith("_"): continue
            tmask |= df[col].fillna("").astype(str).str.lower().str.contains(txt, na=False)
        mask &= tmask
    return df[mask].reset_index(drop=True)

def _get_combined(sids):
    frames = [STORE[s]["df"] for s in sids if s in STORE]
    if not frames: return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)

# ═══════════════════════════════════════════════════════════════════════════
# API
# ═══════════════════════════════════════════════════════════════════════════

def ok(d): return jsonify({"ok":True,**d})
def err(m,c=400): return jsonify({"ok":False,"error":str(m)}),c

# ── Persistencia: guardar/cargar todo el estado (STORE, ML_STORE, etc.) ─────
def _get_pg_conn():
    import psycopg2
    conn = psycopg2.connect(DATABASE_URL)
    with conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS dashify_state (
                id INTEGER PRIMARY KEY,
                blob BYTEA NOT NULL,
                updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
            )
        """)
    conn.commit()
    return conn

def save_state():
    blob = pickle.dumps({
        "STORE": STORE, "ML_STORE": ML_STORE,
        "FICHAS_STORE": FICHAS_STORE, "PUB_STORE": PUB_STORE,
    })
    try:
        if DATABASE_URL:
            import psycopg2
            conn = _get_pg_conn()
            try:
                with conn.cursor() as cur:
                    cur.execute(
                        "INSERT INTO dashify_state (id, blob, updated_at) VALUES (1, %s, now()) "
                        "ON CONFLICT (id) DO UPDATE SET blob = EXCLUDED.blob, updated_at = now()",
                        (psycopg2.Binary(blob),),
                    )
                conn.commit()
            finally:
                conn.close()
        else:
            with open(STATE_FILE, "wb") as fh:
                fh.write(blob)
    except Exception:
        traceback.print_exc()

def load_state():
    try:
        blob = None
        if DATABASE_URL:
            import psycopg2
            conn = _get_pg_conn()
            try:
                with conn.cursor() as cur:
                    cur.execute("SELECT blob FROM dashify_state WHERE id = 1")
                    row = cur.fetchone()
                    if row: blob = bytes(row[0])
            finally:
                conn.close()
            source_desc = "Postgres (DATABASE_URL)"
        else:
            if STATE_FILE.exists():
                blob = STATE_FILE.read_bytes()
            source_desc = str(STATE_FILE)
        if not blob:
            print(f"[Dashify] Sin estado previo guardado ({source_desc}) — arranca vacío.")
            return
        data = pickle.loads(blob)
        STORE.update(data.get("STORE", {}))
        ML_STORE.update(data.get("ML_STORE", {}))
        FICHAS_STORE.update(data.get("FICHAS_STORE", {}))
        PUB_STORE.update(data.get("PUB_STORE", {}))
        print(f"[Dashify] Estado cargado desde {source_desc} "
              f"({len(STORE)} fuente(s), {len(ML_STORE.get('files') or [])} archivo(s) ML)")
    except Exception:
        traceback.print_exc()

load_state()  # cargar datos guardados apenas arranca el proceso (funciona con gunicorn también)

@app.after_request
def _persist_after_write(resp):
    try:
        if request.method in ("POST", "DELETE") and request.path.startswith("/api/") \
           and resp.status_code < 400:
            save_state()
    except Exception:
        traceback.print_exc()
    return resp

# ── Login (protege todo excepto /login y los assets) ────────────────────────
LOGIN_PAGE = """<!DOCTYPE html><html><head><meta charset="utf-8"><title>Dashify — Ingresar</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;
  background:#0c1220;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif}
.box{background:#141b2e;border:1px solid #232d45;border-radius:14px;padding:32px;width:280px}
h1{color:#fff;font-size:18px;margin:0 0 18px}
input{width:100%;padding:10px 12px;border-radius:8px;border:1px solid #2a3552;background:#0c1220;
  color:#fff;font-size:14px;box-sizing:border-box;margin-bottom:12px}
button{width:100%;padding:10px;border:none;border-radius:8px;background:#16a34a;color:#fff;
  font-weight:600;font-size:14px;cursor:pointer}
.err{color:#f87171;font-size:12px;margin-bottom:10px}
</style></head><body>
<form class="box" method="POST">
  <h1>📊 Dashify</h1>
  __ERROR__
  <input type="password" name="password" placeholder="Contraseña" autofocus>
  <button type="submit">Ingresar</button>
</form>
</body></html>"""

def login_required(view):
    @wraps(view)
    def wrapped(*a, **kw):
        if DASHIFY_PASSWORD and not session.get("auth"):
            if request.path.startswith("/api/"):
                return jsonify({"ok": False, "error": "No autenticado"}), 401
            return redirect(url_for("login"))
        return view(*a, **kw)
    return wrapped

@app.route("/login", methods=["GET", "POST"])
def login():
    if not DASHIFY_PASSWORD:
        session["auth"] = True
        return redirect(url_for("index"))
    error = ""
    if request.method == "POST":
        if request.form.get("password") == DASHIFY_PASSWORD:
            session["auth"] = True
            session.permanent = True
            return redirect(url_for("index"))
        error = '<div class="err">Contraseña incorrecta</div>'
    return LOGIN_PAGE.replace("__ERROR__", error)

@app.route("/logout")
def logout():
    session.pop("auth", None)
    return redirect(url_for("login"))

@app.before_request
def _require_login():
    if not DASHIFY_PASSWORD:
        return
    open_paths = ("/login", "/logout")
    if request.path in open_paths:
        return
    if not session.get("auth"):
        if request.path.startswith("/api/"):
            return jsonify({"ok": False, "error": "No autenticado"}), 401
        return redirect(url_for("login"))

@app.route("/")
@login_required
def index(): return HTML

@app.route("/api/upload/ml", methods=["POST"])
def api_upload_ml():
    if "file" not in request.files: return err("Sin archivo.")
    f = request.files["file"]
    try:
        raw = f.read()
        ext = f.filename.rsplit(".", 1)[-1].lower() if "." in f.filename else ""
        if ext == "csv":
            # CSV subido en ML → redirigir automáticamente al parser de TN
            df_new, auto_config = read_tn(raw, f.filename)
            df_new["fuente"] = "Tienda Nube"
            sid_tn = str(uuid.uuid4())
            col_info_tn = infer_column_roles(df_new)
            STORE[sid_tn] = {"name": f.filename, "source": "tn", "df": df_new,
                             "config": auto_config, "col_info": col_info_tn}
            return ok({"sid": sid_tn, "filename": f.filename, "source": "tn",
                       "total_rows": len(df_new), "cols": list(df_new.columns),
                       "col_info": col_info_tn, "config": auto_config,
                       "filtros": filtros_disp(df_new),
                       "_aviso": "CSV detectado: cargado como Tienda Nube automáticamente"})
        df_new, auto_config = read_ml(raw, f.filename)
        col_info = infer_column_roles(df_new)

        # ── UPSERT: fusionar con el acumulado de ML ──────────────────────
        n_added = n_updated = 0
        KEY = "id_venta"  # columna clave: "# de venta"

        if ML_STORE["df"] is None:
            # Primera carga: guardar todo
            df_merged = df_new.copy()
            n_added   = len(df_merged)
        else:
            df_acc = ML_STORE["df"].copy()

            if KEY in df_acc.columns and KEY in df_new.columns:
                keys_acc = set(df_acc[KEY].dropna().astype(str))
                keys_new = df_new[KEY].astype(str)

                # Filas nuevas que no existen en el acumulado
                mask_new  = ~keys_new.isin(keys_acc)
                # Filas que ya existen (actualizar)
                mask_upd  = keys_new.isin(keys_acc)

                n_added   = int(mask_new.sum())
                n_updated = int(mask_upd.sum())

                # Quitar del acumulado las filas que van a ser reemplazadas
                keys_to_update = set(keys_new[mask_upd].tolist())
                df_acc = df_acc[~df_acc[KEY].astype(str).isin(keys_to_update)]

                # Unir: acumulado limpio + filas nuevas + filas actualizadas
                df_merged = pd.concat([df_acc, df_new], ignore_index=True)
            else:
                # Sin columna clave: concatenar y deduplicar
                df_merged = pd.concat([df_acc, df_new], ignore_index=True).drop_duplicates()
                n_added   = len(df_new)

        # Garantizar que fuente siempre sea "Mercado Libre" en el acumulado
        df_merged["fuente"] = "Mercado Libre"

        # ── Enriquecer categoría con Fichas Técnicas (si ya se cargaron) ──
        if FICHAS_STORE["loaded"]:
            df_merged = enrich_categoria_from_fichas(df_merged)

        # Guardar en ML_STORE
        ML_STORE["df"]       = df_merged
        ML_STORE["config"]   = auto_config
        ML_STORE["col_info"] = col_info
        ML_STORE["files"].append(f.filename)

        # Crear o reutilizar el sid fijo de ML
        if ML_STORE["sid"] is None:
            ML_STORE["sid"] = str(uuid.uuid4())
        sid = ML_STORE["sid"]

        STORE[sid] = {
            "name":     "Mercado Libre (acumulado)",
            "source":   "ml",
            "df":       df_merged,
            "config":   auto_config,
            "col_info": col_info,
        }

        return ok({
            "sid":        sid,
            "filename":   f.filename,
            "source":     "ml",
            "total_rows": len(df_merged),
            "cols":       list(df_merged.columns),
            "col_info":   col_info,
            "config":     auto_config,
            "filtros":    filtros_disp(df_merged),
            "upsert": {
                "added":   n_added,
                "updated": n_updated,
                "total":   len(df_merged),
                "files":   ML_STORE["files"],
            }
        })
    except Exception as e:
        traceback.print_exc(); return err(str(e))

@app.route("/api/upload/tn", methods=["POST"])
def api_upload_tn():
    if "file" not in request.files: return err("Sin archivo.")
    f = request.files["file"]
    try:
        raw = f.read()
        df, auto_config = read_tn(raw, f.filename)
        # Garantizar fuente siempre correcta
        df["fuente"] = "Tienda Nube"
        sid = str(uuid.uuid4())
        col_info = infer_column_roles(df)
        STORE[sid] = {"name":f.filename,"source":"tn","df":df,
                      "config":auto_config,"col_info":col_info}
        return ok({"sid":sid,"filename":f.filename,"source":"tn",
                   "total_rows":len(df),"cols":list(df.columns),
                   "col_info":col_info,"config":auto_config,
                   "filtros":filtros_disp(df)})
    except Exception as e:
        traceback.print_exc(); return err(str(e))

@app.route("/api/upload/custom", methods=["POST"])
def api_upload_custom():
    """Carga cualquier Excel/CSV. Retorna info de columnas para configurar."""
    if "file" not in request.files: return err("Sin archivo.")
    f = request.files["file"]
    try:
        raw = f.read()
        df_raw = read_raw_file(raw, f.filename)
        col_info = infer_column_roles(df_raw)
        # Config inicial auto-inferida
        auto_config = {info["role"]: col
                       for col, info in col_info.items()
                       if info.get("role")}
        # Aplicar config inicial
        df = apply_config_to_df(df_raw.copy(), auto_config)
        if "fuente" not in df.columns:
            name = Path(f.filename).stem[:20]
            df["fuente"] = name
        sid = str(uuid.uuid4())
        STORE[sid] = {"name":f.filename,"source":"custom","df":df,
                      "df_raw":df_raw,"config":auto_config,"col_info":col_info}
        return ok({"sid":sid,"filename":f.filename,"source":"custom",
                   "total_rows":len(df_raw),
                   "cols":[c for c in df_raw.columns if not c.startswith("_")],
                   "col_info":col_info,"config":auto_config,
                   "filtros":filtros_disp(df)})
    except Exception as e:
        traceback.print_exc(); return err(str(e))

@app.route("/api/upload/fichas", methods=["POST"])
def api_upload_fichas():
    """
    Carga una planilla de Fichas Técnicas de Mercado Libre.
    Extrae el mapeo ID/SKU/Título → categoría (nombre de hoja) y lo guarda en FICHAS_STORE.
    Si ya hay datos de ML cargados, los re-enriquece inmediatamente.
    """
    if "file" not in request.files:
        return err("Sin archivo.")
    f = request.files["file"]
    try:
        raw = f.read()
        result = read_fichas(raw, f.filename)

        FICHAS_STORE["by_id"]     = result["by_id"]
        FICHAS_STORE["by_sku"]    = result["by_sku"]
        FICHAS_STORE["by_title"]  = result["by_title"]
        FICHAS_STORE["categorias"]= result["categorias"]
        FICHAS_STORE["filename"]  = f.filename
        FICHAS_STORE["loaded"]    = True

        enriched_rows = 0
        # Re-enriquecer el acumulado de ML si existe
        if ML_STORE["df"] is not None and ML_STORE["sid"]:
            df_enr = enrich_categoria_from_fichas(ML_STORE["df"].copy())
            enriched_rows = int((df_enr["categoria"] != ML_STORE["df"].get("categoria", "")).sum()
                                if "categoria" in ML_STORE["df"].columns else len(df_enr))
            ML_STORE["df"] = df_enr
            sid = ML_STORE["sid"]
            if sid in STORE:
                STORE[sid]["df"] = df_enr
                return ok({
                    "filename": f.filename,
                    "categorias": result["categorias"],
                    "total_ids":    len(result["by_id"]),
                    "total_skus":   len(result["by_sku"]),
                    "total_titles": len(result["by_title"]),
                    "enriched_rows": enriched_rows,
                    "filtros": filtros_disp(df_enr),
                    "sid": sid,
                })

        return ok({
            "filename": f.filename,
            "categorias": result["categorias"],
            "total_ids":    len(result["by_id"]),
            "total_skus":   len(result["by_sku"]),
            "total_titles": len(result["by_title"]),
            "enriched_rows": 0,
        })
    except Exception as e:
        traceback.print_exc(); return err(str(e))

@app.route("/api/config/<sid>", methods=["POST"])
def api_config(sid):
    """Aplica nueva configuración de columnas a un dataset custom."""
    if sid not in STORE: return err("Dataset no encontrado.", 404)
    ds = STORE[sid]
    new_config = request.json or {}
    try:
        df_raw = ds.get("df_raw", ds["df"])
        df = apply_config_to_df(df_raw.copy(), new_config)
        name = Path(ds["name"]).stem[:20]
        if "fuente" not in df.columns:
            df["fuente"] = name
        ds["config"] = new_config
        ds["df"] = df
        return ok({"sid":sid,"config":new_config,"filtros":filtros_disp(df),
                   "preview": df.head(5).fillna("").astype(str).to_dict("records")})
    except Exception as e:
        traceback.print_exc(); return err(str(e))

@app.route("/api/dashboard", methods=["POST"])
def api_dashboard():
    body = request.json or {}
    sids = body.get("sids", [])
    f    = body.get("filtros", {})
    periodo = body.get("periodo","dia")
    pg   = int(body.get("page", 0))
    # Visualización custom
    viz = body.get("viz", {})  # { by_col, val_col, agg_fn }

    if not sids: return err("No hay datasets.", 400)
    df0 = _get_combined(sids)
    if df0.empty: return err("Sin datos.", 400)
    df  = apply_filters(df0, f)

    # Métricas por fuente
    fuentes = df["fuente"].unique() if "fuente" in df.columns else []
    metricas_por_fuente = {}
    for fnt in fuentes:
        sub = df[df["fuente"]==fnt]
        metricas_por_fuente[str(fnt)] = metricas(sub)

    # Visualización custom adicional
    viz_custom = []
    by_col  = viz.get("by_col","")
    val_col = viz.get("val_col","")
    agg_fn  = viz.get("agg_fn","sum")
    top_n   = int(viz.get("top",10)) if viz.get("top") else 10
    # Si val_col es None/vacío → conteo
    if by_col and by_col in df.columns:
        if not val_col or agg_fn == "count":
            viz_custom = grp_custom(df, by_col, by_col, "count", top=top_n)
        else:
            viz_custom = grp_custom(df, by_col, val_col, agg_fn, top=top_n)

    return ok({
        "metricas":           metricas(df),
        "metricas_por_fuente":metricas_por_fuente,
        "por_estado":         grp(df,"estado"),
        "por_publicacion":    grp(df,"publicacion",top=20),
        "por_fuente":         grp(df,"fuente"),
        "por_categoria":      grp(df,"categoria",top=15),
        "por_envio":          grp(df,"forma_envio") if "forma_envio" in df.columns else [],
        "por_tiempo":         por_tiempo(df, periodo),
        "por_tiempo_fuentes": {
            str(fnt): por_tiempo(df[df["fuente"]==fnt], periodo)
            for fnt in fuentes
        },
        "por_provincia":      grp(df,"provincia",top=10),
        "flex_zonas":         flex_por_zona(df),
        "viz_custom":         viz_custom,
        "tabla":              tabla_ventas(df, pg),
        "n_filtradas":        len(df),
        "n_total":            len(df0),
        "filtros_disponibles":filtros_disp(df0),
        "col_names":          [c for c in df.columns if not c.startswith("_")],
    })

@app.route("/api/col-values", methods=["POST"])
def api_col_values():
    """Devuelve los valores únicos de una columna para filtros dinámicos."""
    body = request.json or {}
    sids = body.get("sids", [])
    col  = body.get("col", "")
    f    = body.get("filtros", {})
    if not sids or not col: return err("Faltan sids o col.")
    df0 = _get_combined(sids)
    if df0.empty: return err("Sin datos.")
    df = apply_filters(df0, f)
    if col not in df.columns: return err("Columna no encontrada.")
    vals = sorted(set(str(v).strip() for v in df[col].dropna()
                      if str(v).strip() not in ("","nan","None")))[:200]
    return ok({"values": vals, "col": col})

@app.route("/api/delete/<sid>", methods=["DELETE"])
def api_delete(sid):
    STORE.pop(sid, None)
    # Si se elimina el sid de ML, resetear el acumulado
    if ML_STORE["sid"] == sid:
        ML_STORE["sid"]     = None
        ML_STORE["df"]      = None
        ML_STORE["config"]  = None
        ML_STORE["col_info"]= None
        ML_STORE["files"]   = []
    return ok({"deleted": sid})


# ═══════════════════════════════════════════════════════════════════════════
# HTML COMPLETO
# ═══════════════════════════════════════════════════════════════════════════
HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Dashify — Multi Marketplace</title>
<script>
/* Chart.js v4.4.3 — embebido offline */
/*!
 * Chart.js v4.4.3
 * https://www.chartjs.org
 * (c) 2024 Chart.js Contributors
 * Released under the MIT License
 */
!function(t,e){"object"==typeof exports&&"undefined"!=typeof module?module.exports=e():"function"==typeof define&&define.amd?define(e):(t="undefined"!=typeof globalThis?globalThis:t||self).Chart=e()}(this,(function(){"use strict";var t=Object.freeze({__proto__:null,get Colors(){return Go},get Decimation(){return Qo},get Filler(){return ma},get Legend(){return ya},get SubTitle(){return ka},get Title(){return Ma},get Tooltip(){return Ba}});function e(){}const i=(()=>{let t=0;return()=>t++})();function s(t){return null==t}function n(t){if(Array.isArray&&Array.isArray(t))return!0;const e=Object.prototype.toString.call(t);return"[object"===e.slice(0,7)&&"Array]"===e.slice(-6)}function o(t){return null!==t&&"[object Object]"===Object.prototype.toString.call(t)}function a(t){return("number"==typeof t||t instanceof Number)&&isFinite(+t)}function r(t,e){return a(t)?t:e}function l(t,e){return void 0===t?e:t}const h=(t,e)=>"string"==typeof t&&t.endsWith("%")?parseFloat(t)/100:+t/e,c=(t,e)=>"string"==typeof t&&t.endsWith("%")?parseFloat(t)/100*e:+t;function d(t,e,i){if(t&&"function"==typeof t.call)return t.apply(i,e)}function u(t,e,i,s){let a,r,l;if(n(t))if(r=t.length,s)for(a=r-1;a>=0;a--)e.call(i,t[a],a);else for(a=0;a<r;a++)e.call(i,t[a],a);else if(o(t))for(l=Object.keys(t),r=l.length,a=0;a<r;a++)e.call(i,t[l[a]],l[a])}function f(t,e){let i,s,n,o;if(!t||!e||t.length!==e.length)return!1;for(i=0,s=t.length;i<s;++i)if(n=t[i],o=e[i],n.datasetIndex!==o.datasetIndex||n.index!==o.index)return!1;return!0}function g(t){if(n(t))return t.map(g);if(o(t)){const e=Object.create(null),i=Object.keys(t),s=i.length;let n=0;for(;n<s;++n)e[i[n]]=g(t[i[n]]);return e}return t}function p(t){return-1===["__proto__","prototype","constructor"].indexOf(t)}function m(t,e,i,s){if(!p(t))return;const n=e[t],a=i[t];o(n)&&o(a)?x(n,a,s):e[t]=g(a)}function x(t,e,i){const s=n(e)?e:[e],a=s.length;if(!o(t))return t;const r=(i=i||{}).merger||m;let l;for(let e=0;e<a;++e){if(l=s[e],!o(l))continue;const n=Object.keys(l);for(let e=0,s=n.length;e<s;++e)r(n[e],t,l,i)}return t}function b(t,e){return x(t,e,{merger:_})}function _(t,e,i){if(!p(t))return;const s=e[t],n=i[t];o(s)&&o(n)?b(s,n):Object.prototype.hasOwnProperty.call(e,t)||(e[t]=g(n))}const y={"":t=>t,x:t=>t.x,y:t=>t.y};function v(t){const e=t.split("."),i=[];let s="";for(const t of e)s+=t,s.endsWith("\\")?s=s.slice(0,-1)+".":(i.push(s),s="");return i}function M(t,e){const i=y[e]||(y[e]=function(t){const e=v(t);return t=>{for(const i of e){if(""===i)break;t=t&&t[i]}return t}}(e));return i(t)}function w(t){return t.charAt(0).toUpperCase()+t.slice(1)}const k=t=>void 0!==t,S=t=>"function"==typeof t,P=(t,e)=>{if(t.size!==e.size)return!1;for(const i of t)if(!e.has(i))return!1;return!0};function D(t){return"mouseup"===t.type||"click"===t.type||"contextmenu"===t.type}const C=Math.PI,O=2*C,A=O+C,T=Number.POSITIVE_INFINITY,L=C/180,E=C/2,R=C/4,I=2*C/3,z=Math.log10,F=Math.sign;function V(t,e,i){return Math.abs(t-e)<i}function B(t){const e=Math.round(t);t=V(t,e,t/1e3)?e:t;const i=Math.pow(10,Math.floor(z(t))),s=t/i;return(s<=1?1:s<=2?2:s<=5?5:10)*i}function W(t){const e=[],i=Math.sqrt(t);let s;for(s=1;s<i;s++)t%s==0&&(e.push(s),e.push(t/s));return i===(0|i)&&e.push(i),e.sort(((t,e)=>t-e)).pop(),e}function N(t){return!isNaN(parseFloat(t))&&isFinite(t)}function H(t,e){const i=Math.round(t);return i-e<=t&&i+e>=t}function j(t,e,i){let s,n,o;for(s=0,n=t.length;s<n;s++)o=t[s][i],isNaN(o)||(e.min=Math.min(e.min,o),e.max=Math.max(e.max,o))}function $(t){return t*(C/180)}function Y(t){return t*(180/C)}function U(t){if(!a(t))return;let e=1,i=0;for(;Math.round(t*e)/e!==t;)e*=10,i++;return i}function X(t,e){const i=e.x-t.x,s=e.y-t.y,n=Math.sqrt(i*i+s*s);let o=Math.atan2(s,i);return o<-.5*C&&(o+=O),{angle:o,distance:n}}function q(t,e){return Math.sqrt(Math.pow(e.x-t.x,2)+Math.pow(e.y-t.y,2))}function K(t,e){return(t-e+A)%O-C}function G(t){return(t%O+O)%O}function Z(t,e,i,s){const n=G(t),o=G(e),a=G(i),r=G(o-n),l=G(a-n),h=G(n-o),c=G(n-a);return n===o||n===a||s&&o===a||r>l&&h<c}function J(t,e,i){return Math.max(e,Math.min(i,t))}function Q(t){return J(t,-32768,32767)}function tt(t,e,i,s=1e-6){return t>=Math.min(e,i)-s&&t<=Math.max(e,i)+s}function et(t,e,i){i=i||(i=>t[i]<e);let s,n=t.length-1,o=0;for(;n-o>1;)s=o+n>>1,i(s)?o=s:n=s;return{lo:o,hi:n}}const it=(t,e,i,s)=>et(t,i,s?s=>{const n=t[s][e];return n<i||n===i&&t[s+1][e]===i}:s=>t[s][e]<i),st=(t,e,i)=>et(t,i,(s=>t[s][e]>=i));function nt(t,e,i){let s=0,n=t.length;for(;s<n&&t[s]<e;)s++;for(;n>s&&t[n-1]>i;)n--;return s>0||n<t.length?t.slice(s,n):t}const ot=["push","pop","shift","splice","unshift"];function at(t,e){t._chartjs?t._chartjs.listeners.push(e):(Object.defineProperty(t,"_chartjs",{configurable:!0,enumerable:!1,value:{listeners:[e]}}),ot.forEach((e=>{const i="_onData"+w(e),s=t[e];Object.defineProperty(t,e,{configurable:!0,enumerable:!1,value(...e){const n=s.apply(this,e);return t._chartjs.listeners.forEach((t=>{"function"==typeof t[i]&&t[i](...e)})),n}})})))}function rt(t,e){const i=t._chartjs;if(!i)return;const s=i.listeners,n=s.indexOf(e);-1!==n&&s.splice(n,1),s.length>0||(ot.forEach((e=>{delete t[e]})),delete t._chartjs)}function lt(t){const e=new Set(t);return e.size===t.length?t:Array.from(e)}const ht="undefined"==typeof window?function(t){return t()}:window.requestAnimationFrame;function ct(t,e){let i=[],s=!1;return function(...n){i=n,s||(s=!0,ht.call(window,(()=>{s=!1,t.apply(e,i)})))}}function dt(t,e){let i;return function(...s){return e?(clearTimeout(i),i=setTimeout(t,e,s)):t.apply(this,s),e}}const ut=t=>"start"===t?"left":"end"===t?"right":"center",ft=(t,e,i)=>"start"===t?e:"end"===t?i:(e+i)/2,gt=(t,e,i,s)=>t===(s?"left":"right")?i:"center"===t?(e+i)/2:e;function pt(t,e,i){const s=e.length;let n=0,o=s;if(t._sorted){const{iScale:a,_parsed:r}=t,l=a.axis,{min:h,max:c,minDefined:d,maxDefined:u}=a.getUserBounds();d&&(n=J(Math.min(it(r,l,h).lo,i?s:it(e,l,a.getPixelForValue(h)).lo),0,s-1)),o=u?J(Math.max(it(r,a.axis,c,!0).hi+1,i?0:it(e,l,a.getPixelForValue(c),!0).hi+1),n,s)-n:s-n}return{start:n,count:o}}function mt(t){const{xScale:e,yScale:i,_scaleRanges:s}=t,n={xmin:e.min,xmax:e.max,ymin:i.min,ymax:i.max};if(!s)return t._scaleRanges=n,!0;const o=s.xmin!==e.min||s.xmax!==e.max||s.ymin!==i.min||s.ymax!==i.max;return Object.assign(s,n),o}class xt{constructor(){this._request=null,this._charts=new Map,this._running=!1,this._lastDate=void 0}_notify(t,e,i,s){const n=e.listeners[s],o=e.duration;n.forEach((s=>s({chart:t,initial:e.initial,numSteps:o,currentStep:Math.min(i-e.start,o)})))}_refresh(){this._request||(this._running=!0,this._request=ht.call(window,(()=>{this._update(),this._request=null,this._running&&this._refresh()})))}_update(t=Date.now()){let e=0;this._charts.forEach(((i,s)=>{if(!i.running||!i.items.length)return;const n=i.items;let o,a=n.length-1,r=!1;for(;a>=0;--a)o=n[a],o._active?(o._total>i.duration&&(i.duration=o._total),o.tick(t),r=!0):(n[a]=n[n.length-1],n.pop());r&&(s.draw(),this._notify(s,i,t,"progress")),n.length||(i.running=!1,this._notify(s,i,t,"complete"),i.initial=!1),e+=n.length})),this._lastDate=t,0===e&&(this._running=!1)}_getAnims(t){const e=this._charts;let i=e.get(t);return i||(i={running:!1,initial:!0,items:[],listeners:{complete:[],progress:[]}},e.set(t,i)),i}listen(t,e,i){this._getAnims(t).listeners[e].push(i)}add(t,e){e&&e.length&&this._getAnims(t).items.push(...e)}has(t){return this._getAnims(t).items.length>0}start(t){const e=this._charts.get(t);e&&(e.running=!0,e.start=Date.now(),e.duration=e.items.reduce(((t,e)=>Math.max(t,e._duration)),0),this._refresh())}running(t){if(!this._running)return!1;const e=this._charts.get(t);return!!(e&&e.running&&e.items.length)}stop(t){const e=this._charts.get(t);if(!e||!e.items.length)return;const i=e.items;let s=i.length-1;for(;s>=0;--s)i[s].cancel();e.items=[],this._notify(t,e,Date.now(),"complete")}remove(t){return this._charts.delete(t)}}var bt=new xt;
/*!
 * @kurkle/color v0.3.2
 * https://github.com/kurkle/color#readme
 * (c) 2023 Jukka Kurkela
 * Released under the MIT License
 */function _t(t){return t+.5|0}const yt=(t,e,i)=>Math.max(Math.min(t,i),e);function vt(t){return yt(_t(2.55*t),0,255)}function Mt(t){return yt(_t(255*t),0,255)}function wt(t){return yt(_t(t/2.55)/100,0,1)}function kt(t){return yt(_t(100*t),0,100)}const St={0:0,1:1,2:2,3:3,4:4,5:5,6:6,7:7,8:8,9:9,A:10,B:11,C:12,D:13,E:14,F:15,a:10,b:11,c:12,d:13,e:14,f:15},Pt=[..."0123456789ABCDEF"],Dt=t=>Pt[15&t],Ct=t=>Pt[(240&t)>>4]+Pt[15&t],Ot=t=>(240&t)>>4==(15&t);function At(t){var e=(t=>Ot(t.r)&&Ot(t.g)&&Ot(t.b)&&Ot(t.a))(t)?Dt:Ct;return t?"#"+e(t.r)+e(t.g)+e(t.b)+((t,e)=>t<255?e(t):"")(t.a,e):void 0}const Tt=/^(hsla?|hwb|hsv)\(\s*([-+.e\d]+)(?:deg)?[\s,]+([-+.e\d]+)%[\s,]+([-+.e\d]+)%(?:[\s,]+([-+.e\d]+)(%)?)?\s*\)$/;function Lt(t,e,i){const s=e*Math.min(i,1-i),n=(e,n=(e+t/30)%12)=>i-s*Math.max(Math.min(n-3,9-n,1),-1);return[n(0),n(8),n(4)]}function Et(t,e,i){const s=(s,n=(s+t/60)%6)=>i-i*e*Math.max(Math.min(n,4-n,1),0);return[s(5),s(3),s(1)]}function Rt(t,e,i){const s=Lt(t,1,.5);let n;for(e+i>1&&(n=1/(e+i),e*=n,i*=n),n=0;n<3;n++)s[n]*=1-e-i,s[n]+=e;return s}function It(t){const e=t.r/255,i=t.g/255,s=t.b/255,n=Math.max(e,i,s),o=Math.min(e,i,s),a=(n+o)/2;let r,l,h;return n!==o&&(h=n-o,l=a>.5?h/(2-n-o):h/(n+o),r=function(t,e,i,s,n){return t===n?(e-i)/s+(e<i?6:0):e===n?(i-t)/s+2:(t-e)/s+4}(e,i,s,h,n),r=60*r+.5),[0|r,l||0,a]}function zt(t,e,i,s){return(Array.isArray(e)?t(e[0],e[1],e[2]):t(e,i,s)).map(Mt)}function Ft(t,e,i){return zt(Lt,t,e,i)}function Vt(t){return(t%360+360)%360}function Bt(t){const e=Tt.exec(t);let i,s=255;if(!e)return;e[5]!==i&&(s=e[6]?vt(+e[5]):Mt(+e[5]));const n=Vt(+e[2]),o=+e[3]/100,a=+e[4]/100;return i="hwb"===e[1]?function(t,e,i){return zt(Rt,t,e,i)}(n,o,a):"hsv"===e[1]?function(t,e,i){return zt(Et,t,e,i)}(n,o,a):Ft(n,o,a),{r:i[0],g:i[1],b:i[2],a:s}}const Wt={x:"dark",Z:"light",Y:"re",X:"blu",W:"gr",V:"medium",U:"slate",A:"ee",T:"ol",S:"or",B:"ra",C:"lateg",D:"ights",R:"in",Q:"turquois",E:"hi",P:"ro",O:"al",N:"le",M:"de",L:"yello",F:"en",K:"ch",G:"arks",H:"ea",I:"ightg",J:"wh"},Nt={OiceXe:"f0f8ff",antiquewEte:"faebd7",aqua:"ffff",aquamarRe:"7fffd4",azuY:"f0ffff",beige:"f5f5dc",bisque:"ffe4c4",black:"0",blanKedOmond:"ffebcd",Xe:"ff",XeviTet:"8a2be2",bPwn:"a52a2a",burlywood:"deb887",caMtXe:"5f9ea0",KartYuse:"7fff00",KocTate:"d2691e",cSO:"ff7f50",cSnflowerXe:"6495ed",cSnsilk:"fff8dc",crimson:"dc143c",cyan:"ffff",xXe:"8b",xcyan:"8b8b",xgTMnPd:"b8860b",xWay:"a9a9a9",xgYF:"6400",xgYy:"a9a9a9",xkhaki:"bdb76b",xmagFta:"8b008b",xTivegYF:"556b2f",xSange:"ff8c00",xScEd:"9932cc",xYd:"8b0000",xsOmon:"e9967a",xsHgYF:"8fbc8f",xUXe:"483d8b",xUWay:"2f4f4f",xUgYy:"2f4f4f",xQe:"ced1",xviTet:"9400d3",dAppRk:"ff1493",dApskyXe:"bfff",dimWay:"696969",dimgYy:"696969",dodgerXe:"1e90ff",fiYbrick:"b22222",flSOwEte:"fffaf0",foYstWAn:"228b22",fuKsia:"ff00ff",gaRsbSo:"dcdcdc",ghostwEte:"f8f8ff",gTd:"ffd700",gTMnPd:"daa520",Way:"808080",gYF:"8000",gYFLw:"adff2f",gYy:"808080",honeyMw:"f0fff0",hotpRk:"ff69b4",RdianYd:"cd5c5c",Rdigo:"4b0082",ivSy:"fffff0",khaki:"f0e68c",lavFMr:"e6e6fa",lavFMrXsh:"fff0f5",lawngYF:"7cfc00",NmoncEffon:"fffacd",ZXe:"add8e6",ZcSO:"f08080",Zcyan:"e0ffff",ZgTMnPdLw:"fafad2",ZWay:"d3d3d3",ZgYF:"90ee90",ZgYy:"d3d3d3",ZpRk:"ffb6c1",ZsOmon:"ffa07a",ZsHgYF:"20b2aa",ZskyXe:"87cefa",ZUWay:"778899",ZUgYy:"778899",ZstAlXe:"b0c4de",ZLw:"ffffe0",lime:"ff00",limegYF:"32cd32",lRF:"faf0e6",magFta:"ff00ff",maPon:"800000",VaquamarRe:"66cdaa",VXe:"cd",VScEd:"ba55d3",VpurpN:"9370db",VsHgYF:"3cb371",VUXe:"7b68ee",VsprRggYF:"fa9a",VQe:"48d1cc",VviTetYd:"c71585",midnightXe:"191970",mRtcYam:"f5fffa",mistyPse:"ffe4e1",moccasR:"ffe4b5",navajowEte:"ffdead",navy:"80",Tdlace:"fdf5e6",Tive:"808000",TivedBb:"6b8e23",Sange:"ffa500",SangeYd:"ff4500",ScEd:"da70d6",pOegTMnPd:"eee8aa",pOegYF:"98fb98",pOeQe:"afeeee",pOeviTetYd:"db7093",papayawEp:"ffefd5",pHKpuff:"ffdab9",peru:"cd853f",pRk:"ffc0cb",plum:"dda0dd",powMrXe:"b0e0e6",purpN:"800080",YbeccapurpN:"663399",Yd:"ff0000",Psybrown:"bc8f8f",PyOXe:"4169e1",saddNbPwn:"8b4513",sOmon:"fa8072",sandybPwn:"f4a460",sHgYF:"2e8b57",sHshell:"fff5ee",siFna:"a0522d",silver:"c0c0c0",skyXe:"87ceeb",UXe:"6a5acd",UWay:"708090",UgYy:"708090",snow:"fffafa",sprRggYF:"ff7f",stAlXe:"4682b4",tan:"d2b48c",teO:"8080",tEstN:"d8bfd8",tomato:"ff6347",Qe:"40e0d0",viTet:"ee82ee",JHt:"f5deb3",wEte:"ffffff",wEtesmoke:"f5f5f5",Lw:"ffff00",LwgYF:"9acd32"};let Ht;function jt(t){Ht||(Ht=function(){const t={},e=Object.keys(Nt),i=Object.keys(Wt);let s,n,o,a,r;for(s=0;s<e.length;s++){for(a=r=e[s],n=0;n<i.length;n++)o=i[n],r=r.replace(o,Wt[o]);o=parseInt(Nt[a],16),t[r]=[o>>16&255,o>>8&255,255&o]}return t}(),Ht.transparent=[0,0,0,0]);const e=Ht[t.toLowerCase()];return e&&{r:e[0],g:e[1],b:e[2],a:4===e.length?e[3]:255}}const $t=/^rgba?\(\s*([-+.\d]+)(%)?[\s,]+([-+.e\d]+)(%)?[\s,]+([-+.e\d]+)(%)?(?:[\s,/]+([-+.e\d]+)(%)?)?\s*\)$/;const Yt=t=>t<=.0031308?12.92*t:1.055*Math.pow(t,1/2.4)-.055,Ut=t=>t<=.04045?t/12.92:Math.pow((t+.055)/1.055,2.4);function Xt(t,e,i){if(t){let s=It(t);s[e]=Math.max(0,Math.min(s[e]+s[e]*i,0===e?360:1)),s=Ft(s),t.r=s[0],t.g=s[1],t.b=s[2]}}function qt(t,e){return t?Object.assign(e||{},t):t}function Kt(t){var e={r:0,g:0,b:0,a:255};return Array.isArray(t)?t.length>=3&&(e={r:t[0],g:t[1],b:t[2],a:255},t.length>3&&(e.a=Mt(t[3]))):(e=qt(t,{r:0,g:0,b:0,a:1})).a=Mt(e.a),e}function Gt(t){return"r"===t.charAt(0)?function(t){const e=$t.exec(t);let i,s,n,o=255;if(e){if(e[7]!==i){const t=+e[7];o=e[8]?vt(t):yt(255*t,0,255)}return i=+e[1],s=+e[3],n=+e[5],i=255&(e[2]?vt(i):yt(i,0,255)),s=255&(e[4]?vt(s):yt(s,0,255)),n=255&(e[6]?vt(n):yt(n,0,255)),{r:i,g:s,b:n,a:o}}}(t):Bt(t)}class Zt{constructor(t){if(t instanceof Zt)return t;const e=typeof t;let i;var s,n,o;"object"===e?i=Kt(t):"string"===e&&(o=(s=t).length,"#"===s[0]&&(4===o||5===o?n={r:255&17*St[s[1]],g:255&17*St[s[2]],b:255&17*St[s[3]],a:5===o?17*St[s[4]]:255}:7!==o&&9!==o||(n={r:St[s[1]]<<4|St[s[2]],g:St[s[3]]<<4|St[s[4]],b:St[s[5]]<<4|St[s[6]],a:9===o?St[s[7]]<<4|St[s[8]]:255})),i=n||jt(t)||Gt(t)),this._rgb=i,this._valid=!!i}get valid(){return this._valid}get rgb(){var t=qt(this._rgb);return t&&(t.a=wt(t.a)),t}set rgb(t){this._rgb=Kt(t)}rgbString(){return this._valid?(t=this._rgb)&&(t.a<255?`rgba(${t.r}, ${t.g}, ${t.b}, ${wt(t.a)})`:`rgb(${t.r}, ${t.g}, ${t.b})`):void 0;var t}hexString(){return this._valid?At(this._rgb):void 0}hslString(){return this._valid?function(t){if(!t)return;const e=It(t),i=e[0],s=kt(e[1]),n=kt(e[2]);return t.a<255?`hsla(${i}, ${s}%, ${n}%, ${wt(t.a)})`:`hsl(${i}, ${s}%, ${n}%)`}(this._rgb):void 0}mix(t,e){if(t){const i=this.rgb,s=t.rgb;let n;const o=e===n?.5:e,a=2*o-1,r=i.a-s.a,l=((a*r==-1?a:(a+r)/(1+a*r))+1)/2;n=1-l,i.r=255&l*i.r+n*s.r+.5,i.g=255&l*i.g+n*s.g+.5,i.b=255&l*i.b+n*s.b+.5,i.a=o*i.a+(1-o)*s.a,this.rgb=i}return this}interpolate(t,e){return t&&(this._rgb=function(t,e,i){const s=Ut(wt(t.r)),n=Ut(wt(t.g)),o=Ut(wt(t.b));return{r:Mt(Yt(s+i*(Ut(wt(e.r))-s))),g:Mt(Yt(n+i*(Ut(wt(e.g))-n))),b:Mt(Yt(o+i*(Ut(wt(e.b))-o))),a:t.a+i*(e.a-t.a)}}(this._rgb,t._rgb,e)),this}clone(){return new Zt(this.rgb)}alpha(t){return this._rgb.a=Mt(t),this}clearer(t){return this._rgb.a*=1-t,this}greyscale(){const t=this._rgb,e=_t(.3*t.r+.59*t.g+.11*t.b);return t.r=t.g=t.b=e,this}opaquer(t){return this._rgb.a*=1+t,this}negate(){const t=this._rgb;return t.r=255-t.r,t.g=255-t.g,t.b=255-t.b,this}lighten(t){return Xt(this._rgb,2,t),this}darken(t){return Xt(this._rgb,2,-t),this}saturate(t){return Xt(this._rgb,1,t),this}desaturate(t){return Xt(this._rgb,1,-t),this}rotate(t){return function(t,e){var i=It(t);i[0]=Vt(i[0]+e),i=Ft(i),t.r=i[0],t.g=i[1],t.b=i[2]}(this._rgb,t),this}}function Jt(t){if(t&&"object"==typeof t){const e=t.toString();return"[object CanvasPattern]"===e||"[object CanvasGradient]"===e}return!1}function Qt(t){return Jt(t)?t:new Zt(t)}function te(t){return Jt(t)?t:new Zt(t).saturate(.5).darken(.1).hexString()}const ee=["x","y","borderWidth","radius","tension"],ie=["color","borderColor","backgroundColor"];const se=new Map;function ne(t,e,i){return function(t,e){e=e||{};const i=t+JSON.stringify(e);let s=se.get(i);return s||(s=new Intl.NumberFormat(t,e),se.set(i,s)),s}(e,i).format(t)}const oe={values:t=>n(t)?t:""+t,numeric(t,e,i){if(0===t)return"0";const s=this.chart.options.locale;let n,o=t;if(i.length>1){const e=Math.max(Math.abs(i[0].value),Math.abs(i[i.length-1].value));(e<1e-4||e>1e15)&&(n="scientific"),o=function(t,e){let i=e.length>3?e[2].value-e[1].value:e[1].value-e[0].value;Math.abs(i)>=1&&t!==Math.floor(t)&&(i=t-Math.floor(t));return i}(t,i)}const a=z(Math.abs(o)),r=isNaN(a)?1:Math.max(Math.min(-1*Math.floor(a),20),0),l={notation:n,minimumFractionDigits:r,maximumFractionDigits:r};return Object.assign(l,this.options.ticks.format),ne(t,s,l)},logarithmic(t,e,i){if(0===t)return"0";const s=i[e].significand||t/Math.pow(10,Math.floor(z(t)));return[1,2,3,5,10,15].includes(s)||e>.8*i.length?oe.numeric.call(this,t,e,i):""}};var ae={formatters:oe};const re=Object.create(null),le=Object.create(null);function he(t,e){if(!e)return t;const i=e.split(".");for(let e=0,s=i.length;e<s;++e){const s=i[e];t=t[s]||(t[s]=Object.create(null))}return t}function ce(t,e,i){return"string"==typeof e?x(he(t,e),i):x(he(t,""),e)}class de{constructor(t,e){this.animation=void 0,this.backgroundColor="rgba(0,0,0,0.1)",this.borderColor="rgba(0,0,0,0.1)",this.color="#666",this.datasets={},this.devicePixelRatio=t=>t.chart.platform.getDevicePixelRatio(),this.elements={},this.events=["mousemove","mouseout","click","touchstart","touchmove"],this.font={family:"'Helvetica Neue', 'Helvetica', 'Arial', sans-serif",size:12,style:"normal",lineHeight:1.2,weight:null},this.hover={},this.hoverBackgroundColor=(t,e)=>te(e.backgroundColor),this.hoverBorderColor=(t,e)=>te(e.borderColor),this.hoverColor=(t,e)=>te(e.color),this.indexAxis="x",this.interaction={mode:"nearest",intersect:!0,includeInvisible:!1},this.maintainAspectRatio=!0,this.onHover=null,this.onClick=null,this.parsing=!0,this.plugins={},this.responsive=!0,this.scale=void 0,this.scales={},this.showLine=!0,this.drawActiveElementsOnTop=!0,this.describe(t),this.apply(e)}set(t,e){return ce(this,t,e)}get(t){return he(this,t)}describe(t,e){return ce(le,t,e)}override(t,e){return ce(re,t,e)}route(t,e,i,s){const n=he(this,t),a=he(this,i),r="_"+e;Object.defineProperties(n,{[r]:{value:n[e],writable:!0},[e]:{enumerable:!0,get(){const t=this[r],e=a[s];return o(t)?Object.assign({},e,t):l(t,e)},set(t){this[r]=t}}})}apply(t){t.forEach((t=>t(this)))}}var ue=new de({_scriptable:t=>!t.startsWith("on"),_indexable:t=>"events"!==t,hover:{_fallback:"interaction"},interaction:{_scriptable:!1,_indexable:!1}},[function(t){t.set("animation",{delay:void 0,duration:1e3,easing:"easeOutQuart",fn:void 0,from:void 0,loop:void 0,to:void 0,type:void 0}),t.describe("animation",{_fallback:!1,_indexable:!1,_scriptable:t=>"onProgress"!==t&&"onComplete"!==t&&"fn"!==t}),t.set("animations",{colors:{type:"color",properties:ie},numbers:{type:"number",properties:ee}}),t.describe("animations",{_fallback:"animation"}),t.set("transitions",{active:{animation:{duration:400}},resize:{animation:{duration:0}},show:{animations:{colors:{from:"transparent"},visible:{type:"boolean",duration:0}}},hide:{animations:{colors:{to:"transparent"},visible:{type:"boolean",easing:"linear",fn:t=>0|t}}}})},function(t){t.set("layout",{autoPadding:!0,padding:{top:0,right:0,bottom:0,left:0}})},function(t){t.set("scale",{display:!0,offset:!1,reverse:!1,beginAtZero:!1,bounds:"ticks",clip:!0,grace:0,grid:{display:!0,lineWidth:1,drawOnChartArea:!0,drawTicks:!0,tickLength:8,tickWidth:(t,e)=>e.lineWidth,tickColor:(t,e)=>e.color,offset:!1},border:{display:!0,dash:[],dashOffset:0,width:1},title:{display:!1,text:"",padding:{top:4,bottom:4}},ticks:{minRotation:0,maxRotation:50,mirror:!1,textStrokeWidth:0,textStrokeColor:"",padding:3,display:!0,autoSkip:!0,autoSkipPadding:3,labelOffset:0,callback:ae.formatters.values,minor:{},major:{},align:"center",crossAlign:"near",showLabelBackdrop:!1,backdropColor:"rgba(255, 255, 255, 0.75)",backdropPadding:2}}),t.route("scale.ticks","color","","color"),t.route("scale.grid","color","","borderColor"),t.route("scale.border","color","","borderColor"),t.route("scale.title","color","","color"),t.describe("scale",{_fallback:!1,_scriptable:t=>!t.startsWith("before")&&!t.startsWith("after")&&"callback"!==t&&"parser"!==t,_indexable:t=>"borderDash"!==t&&"tickBorderDash"!==t&&"dash"!==t}),t.describe("scales",{_fallback:"scale"}),t.describe("scale.ticks",{_scriptable:t=>"backdropPadding"!==t&&"callback"!==t,_indexable:t=>"backdropPadding"!==t})}]);function fe(){return"undefined"!=typeof window&&"undefined"!=typeof document}function ge(t){let e=t.parentNode;return e&&"[object ShadowRoot]"===e.toString()&&(e=e.host),e}function pe(t,e,i){let s;return"string"==typeof t?(s=parseInt(t,10),-1!==t.indexOf("%")&&(s=s/100*e.parentNode[i])):s=t,s}const me=t=>t.ownerDocument.defaultView.getComputedStyle(t,null);function xe(t,e){return me(t).getPropertyValue(e)}const be=["top","right","bottom","left"];function _e(t,e,i){const s={};i=i?"-"+i:"";for(let n=0;n<4;n++){const o=be[n];s[o]=parseFloat(t[e+"-"+o+i])||0}return s.width=s.left+s.right,s.height=s.top+s.bottom,s}const ye=(t,e,i)=>(t>0||e>0)&&(!i||!i.shadowRoot);function ve(t,e){if("native"in t)return t;const{canvas:i,currentDevicePixelRatio:s}=e,n=me(i),o="border-box"===n.boxSizing,a=_e(n,"padding"),r=_e(n,"border","width"),{x:l,y:h,box:c}=function(t,e){const i=t.touches,s=i&&i.length?i[0]:t,{offsetX:n,offsetY:o}=s;let a,r,l=!1;if(ye(n,o,t.target))a=n,r=o;else{const t=e.getBoundingClientRect();a=s.clientX-t.left,r=s.clientY-t.top,l=!0}return{x:a,y:r,box:l}}(t,i),d=a.left+(c&&r.left),u=a.top+(c&&r.top);let{width:f,height:g}=e;return o&&(f-=a.width+r.width,g-=a.height+r.height),{x:Math.round((l-d)/f*i.width/s),y:Math.round((h-u)/g*i.height/s)}}const Me=t=>Math.round(10*t)/10;function we(t,e,i,s){const n=me(t),o=_e(n,"margin"),a=pe(n.maxWidth,t,"clientWidth")||T,r=pe(n.maxHeight,t,"clientHeight")||T,l=function(t,e,i){let s,n;if(void 0===e||void 0===i){const o=t&&ge(t);if(o){const t=o.getBoundingClientRect(),a=me(o),r=_e(a,"border","width"),l=_e(a,"padding");e=t.width-l.width-r.width,i=t.height-l.height-r.height,s=pe(a.maxWidth,o,"clientWidth"),n=pe(a.maxHeight,o,"clientHeight")}else e=t.clientWidth,i=t.clientHeight}return{width:e,height:i,maxWidth:s||T,maxHeight:n||T}}(t,e,i);let{width:h,height:c}=l;if("content-box"===n.boxSizing){const t=_e(n,"border","width"),e=_e(n,"padding");h-=e.width+t.width,c-=e.height+t.height}h=Math.max(0,h-o.width),c=Math.max(0,s?h/s:c-o.height),h=Me(Math.min(h,a,l.maxWidth)),c=Me(Math.min(c,r,l.maxHeight)),h&&!c&&(c=Me(h/2));return(void 0!==e||void 0!==i)&&s&&l.height&&c>l.height&&(c=l.height,h=Me(Math.floor(c*s))),{width:h,height:c}}function ke(t,e,i){const s=e||1,n=Math.floor(t.height*s),o=Math.floor(t.width*s);t.height=Math.floor(t.height),t.width=Math.floor(t.width);const a=t.canvas;return a.style&&(i||!a.style.height&&!a.style.width)&&(a.style.height=`${t.height}px`,a.style.width=`${t.width}px`),(t.currentDevicePixelRatio!==s||a.height!==n||a.width!==o)&&(t.currentDevicePixelRatio=s,a.height=n,a.width=o,t.ctx.setTransform(s,0,0,s,0,0),!0)}const Se=function(){let t=!1;try{const e={get passive(){return t=!0,!1}};fe()&&(window.addEventListener("test",null,e),window.removeEventListener("test",null,e))}catch(t){}return t}();function Pe(t,e){const i=xe(t,e),s=i&&i.match(/^(\d+)(\.\d+)?px$/);return s?+s[1]:void 0}function De(t){return!t||s(t.size)||s(t.family)?null:(t.style?t.style+" ":"")+(t.weight?t.weight+" ":"")+t.size+"px "+t.family}function Ce(t,e,i,s,n){let o=e[n];return o||(o=e[n]=t.measureText(n).width,i.push(n)),o>s&&(s=o),s}function Oe(t,e,i,s){let o=(s=s||{}).data=s.data||{},a=s.garbageCollect=s.garbageCollect||[];s.font!==e&&(o=s.data={},a=s.garbageCollect=[],s.font=e),t.save(),t.font=e;let r=0;const l=i.length;let h,c,d,u,f;for(h=0;h<l;h++)if(u=i[h],null==u||n(u)){if(n(u))for(c=0,d=u.length;c<d;c++)f=u[c],null==f||n(f)||(r=Ce(t,o,a,r,f))}else r=Ce(t,o,a,r,u);t.restore();const g=a.length/2;if(g>i.length){for(h=0;h<g;h++)delete o[a[h]];a.splice(0,g)}return r}function Ae(t,e,i){const s=t.currentDevicePixelRatio,n=0!==i?Math.max(i/2,.5):0;return Math.round((e-n)*s)/s+n}function Te(t,e){(e||t)&&((e=e||t.getContext("2d")).save(),e.resetTransform(),e.clearRect(0,0,t.width,t.height),e.restore())}function Le(t,e,i,s){Ee(t,e,i,s,null)}function Ee(t,e,i,s,n){let o,a,r,l,h,c,d,u;const f=e.pointStyle,g=e.rotation,p=e.radius;let m=(g||0)*L;if(f&&"object"==typeof f&&(o=f.toString(),"[object HTMLImageElement]"===o||"[object HTMLCanvasElement]"===o))return t.save(),t.translate(i,s),t.rotate(m),t.drawImage(f,-f.width/2,-f.height/2,f.width,f.height),void t.restore();if(!(isNaN(p)||p<=0)){switch(t.beginPath(),f){default:n?t.ellipse(i,s,n/2,p,0,0,O):t.arc(i,s,p,0,O),t.closePath();break;case"triangle":c=n?n/2:p,t.moveTo(i+Math.sin(m)*c,s-Math.cos(m)*p),m+=I,t.lineTo(i+Math.sin(m)*c,s-Math.cos(m)*p),m+=I,t.lineTo(i+Math.sin(m)*c,s-Math.cos(m)*p),t.closePath();break;case"rectRounded":h=.516*p,l=p-h,a=Math.cos(m+R)*l,d=Math.cos(m+R)*(n?n/2-h:l),r=Math.sin(m+R)*l,u=Math.sin(m+R)*(n?n/2-h:l),t.arc(i-d,s-r,h,m-C,m-E),t.arc(i+u,s-a,h,m-E,m),t.arc(i+d,s+r,h,m,m+E),t.arc(i-u,s+a,h,m+E,m+C),t.closePath();break;case"rect":if(!g){l=Math.SQRT1_2*p,c=n?n/2:l,t.rect(i-c,s-l,2*c,2*l);break}m+=R;case"rectRot":d=Math.cos(m)*(n?n/2:p),a=Math.cos(m)*p,r=Math.sin(m)*p,u=Math.sin(m)*(n?n/2:p),t.moveTo(i-d,s-r),t.lineTo(i+u,s-a),t.lineTo(i+d,s+r),t.lineTo(i-u,s+a),t.closePath();break;case"crossRot":m+=R;case"cross":d=Math.cos(m)*(n?n/2:p),a=Math.cos(m)*p,r=Math.sin(m)*p,u=Math.sin(m)*(n?n/2:p),t.moveTo(i-d,s-r),t.lineTo(i+d,s+r),t.moveTo(i+u,s-a),t.lineTo(i-u,s+a);break;case"star":d=Math.cos(m)*(n?n/2:p),a=Math.cos(m)*p,r=Math.sin(m)*p,u=Math.sin(m)*(n?n/2:p),t.moveTo(i-d,s-r),t.lineTo(i+d,s+r),t.moveTo(i+u,s-a),t.lineTo(i-u,s+a),m+=R,d=Math.cos(m)*(n?n/2:p),a=Math.cos(m)*p,r=Math.sin(m)*p,u=Math.sin(m)*(n?n/2:p),t.moveTo(i-d,s-r),t.lineTo(i+d,s+r),t.moveTo(i+u,s-a),t.lineTo(i-u,s+a);break;case"line":a=n?n/2:Math.cos(m)*p,r=Math.sin(m)*p,t.moveTo(i-a,s-r),t.lineTo(i+a,s+r);break;case"dash":t.moveTo(i,s),t.lineTo(i+Math.cos(m)*(n?n/2:p),s+Math.sin(m)*p);break;case!1:t.closePath()}t.fill(),e.borderWidth>0&&t.stroke()}}function Re(t,e,i){return i=i||.5,!e||t&&t.x>e.left-i&&t.x<e.right+i&&t.y>e.top-i&&t.y<e.bottom+i}function Ie(t,e){t.save(),t.beginPath(),t.rect(e.left,e.top,e.right-e.left,e.bottom-e.top),t.clip()}function ze(t){t.restore()}function Fe(t,e,i,s,n){if(!e)return t.lineTo(i.x,i.y);if("middle"===n){const s=(e.x+i.x)/2;t.lineTo(s,e.y),t.lineTo(s,i.y)}else"after"===n!=!!s?t.lineTo(e.x,i.y):t.lineTo(i.x,e.y);t.lineTo(i.x,i.y)}function Ve(t,e,i,s){if(!e)return t.lineTo(i.x,i.y);t.bezierCurveTo(s?e.cp1x:e.cp2x,s?e.cp1y:e.cp2y,s?i.cp2x:i.cp1x,s?i.cp2y:i.cp1y,i.x,i.y)}function Be(t,e,i,s,n){if(n.strikethrough||n.underline){const o=t.measureText(s),a=e-o.actualBoundingBoxLeft,r=e+o.actualBoundingBoxRight,l=i-o.actualBoundingBoxAscent,h=i+o.actualBoundingBoxDescent,c=n.strikethrough?(l+h)/2:h;t.strokeStyle=t.fillStyle,t.beginPath(),t.lineWidth=n.decorationWidth||2,t.moveTo(a,c),t.lineTo(r,c),t.stroke()}}function We(t,e){const i=t.fillStyle;t.fillStyle=e.color,t.fillRect(e.left,e.top,e.width,e.height),t.fillStyle=i}function Ne(t,e,i,o,a,r={}){const l=n(e)?e:[e],h=r.strokeWidth>0&&""!==r.strokeColor;let c,d;for(t.save(),t.font=a.string,function(t,e){e.translation&&t.translate(e.translation[0],e.translation[1]),s(e.rotation)||t.rotate(e.rotation),e.color&&(t.fillStyle=e.color),e.textAlign&&(t.textAlign=e.textAlign),e.textBaseline&&(t.textBaseline=e.textBaseline)}(t,r),c=0;c<l.length;++c)d=l[c],r.backdrop&&We(t,r.backdrop),h&&(r.strokeColor&&(t.strokeStyle=r.strokeColor),s(r.strokeWidth)||(t.lineWidth=r.strokeWidth),t.strokeText(d,i,o,r.maxWidth)),t.fillText(d,i,o,r.maxWidth),Be(t,i,o,d,r),o+=Number(a.lineHeight);t.restore()}function He(t,e){const{x:i,y:s,w:n,h:o,radius:a}=e;t.arc(i+a.topLeft,s+a.topLeft,a.topLeft,1.5*C,C,!0),t.lineTo(i,s+o-a.bottomLeft),t.arc(i+a.bottomLeft,s+o-a.bottomLeft,a.bottomLeft,C,E,!0),t.lineTo(i+n-a.bottomRight,s+o),t.arc(i+n-a.bottomRight,s+o-a.bottomRight,a.bottomRight,E,0,!0),t.lineTo(i+n,s+a.topRight),t.arc(i+n-a.topRight,s+a.topRight,a.topRight,0,-E,!0),t.lineTo(i+a.topLeft,s)}function je(t,e=[""],i,s,n=(()=>t[0])){const o=i||t;void 0===s&&(s=ti("_fallback",t));const a={[Symbol.toStringTag]:"Object",_cacheable:!0,_scopes:t,_rootScopes:o,_fallback:s,_getTarget:n,override:i=>je([i,...t],e,o,s)};return new Proxy(a,{deleteProperty:(e,i)=>(delete e[i],delete e._keys,delete t[0][i],!0),get:(i,s)=>qe(i,s,(()=>function(t,e,i,s){let n;for(const o of e)if(n=ti(Ue(o,t),i),void 0!==n)return Xe(t,n)?Je(i,s,t,n):n}(s,e,t,i))),getOwnPropertyDescriptor:(t,e)=>Reflect.getOwnPropertyDescriptor(t._scopes[0],e),getPrototypeOf:()=>Reflect.getPrototypeOf(t[0]),has:(t,e)=>ei(t).includes(e),ownKeys:t=>ei(t),set(t,e,i){const s=t._storage||(t._storage=n());return t[e]=s[e]=i,delete t._keys,!0}})}function $e(t,e,i,s){const a={_cacheable:!1,_proxy:t,_context:e,_subProxy:i,_stack:new Set,_descriptors:Ye(t,s),setContext:e=>$e(t,e,i,s),override:n=>$e(t.override(n),e,i,s)};return new Proxy(a,{deleteProperty:(e,i)=>(delete e[i],delete t[i],!0),get:(t,e,i)=>qe(t,e,(()=>function(t,e,i){const{_proxy:s,_context:a,_subProxy:r,_descriptors:l}=t;let h=s[e];S(h)&&l.isScriptable(e)&&(h=function(t,e,i,s){const{_proxy:n,_context:o,_subProxy:a,_stack:r}=i;if(r.has(t))throw new Error("Recursion detected: "+Array.from(r).join("->")+"->"+t);r.add(t);let l=e(o,a||s);r.delete(t),Xe(t,l)&&(l=Je(n._scopes,n,t,l));return l}(e,h,t,i));n(h)&&h.length&&(h=function(t,e,i,s){const{_proxy:n,_context:a,_subProxy:r,_descriptors:l}=i;if(void 0!==a.index&&s(t))return e[a.index%e.length];if(o(e[0])){const i=e,s=n._scopes.filter((t=>t!==i));e=[];for(const o of i){const i=Je(s,n,t,o);e.push($e(i,a,r&&r[t],l))}}return e}(e,h,t,l.isIndexable));Xe(e,h)&&(h=$e(h,a,r&&r[e],l));return h}(t,e,i))),getOwnPropertyDescriptor:(e,i)=>e._descriptors.allKeys?Reflect.has(t,i)?{enumerable:!0,configurable:!0}:void 0:Reflect.getOwnPropertyDescriptor(t,i),getPrototypeOf:()=>Reflect.getPrototypeOf(t),has:(e,i)=>Reflect.has(t,i),ownKeys:()=>Reflect.ownKeys(t),set:(e,i,s)=>(t[i]=s,delete e[i],!0)})}function Ye(t,e={scriptable:!0,indexable:!0}){const{_scriptable:i=e.scriptable,_indexable:s=e.indexable,_allKeys:n=e.allKeys}=t;return{allKeys:n,scriptable:i,indexable:s,isScriptable:S(i)?i:()=>i,isIndexable:S(s)?s:()=>s}}const Ue=(t,e)=>t?t+w(e):e,Xe=(t,e)=>o(e)&&"adapters"!==t&&(null===Object.getPrototypeOf(e)||e.constructor===Object);function qe(t,e,i){if(Object.prototype.hasOwnProperty.call(t,e)||"constructor"===e)return t[e];const s=i();return t[e]=s,s}function Ke(t,e,i){return S(t)?t(e,i):t}const Ge=(t,e)=>!0===t?e:"string"==typeof t?M(e,t):void 0;function Ze(t,e,i,s,n){for(const o of e){const e=Ge(i,o);if(e){t.add(e);const o=Ke(e._fallback,i,n);if(void 0!==o&&o!==i&&o!==s)return o}else if(!1===e&&void 0!==s&&i!==s)return null}return!1}function Je(t,e,i,s){const a=e._rootScopes,r=Ke(e._fallback,i,s),l=[...t,...a],h=new Set;h.add(s);let c=Qe(h,l,i,r||i,s);return null!==c&&((void 0===r||r===i||(c=Qe(h,l,r,c,s),null!==c))&&je(Array.from(h),[""],a,r,(()=>function(t,e,i){const s=t._getTarget();e in s||(s[e]={});const a=s[e];if(n(a)&&o(i))return i;return a||{}}(e,i,s))))}function Qe(t,e,i,s,n){for(;i;)i=Ze(t,e,i,s,n);return i}function ti(t,e){for(const i of e){if(!i)continue;const e=i[t];if(void 0!==e)return e}}function ei(t){let e=t._keys;return e||(e=t._keys=function(t){const e=new Set;for(const i of t)for(const t of Object.keys(i).filter((t=>!t.startsWith("_"))))e.add(t);return Array.from(e)}(t._scopes)),e}function ii(t,e,i,s){const{iScale:n}=t,{key:o="r"}=this._parsing,a=new Array(s);let r,l,h,c;for(r=0,l=s;r<l;++r)h=r+i,c=e[h],a[r]={r:n.parse(M(c,o),h)};return a}const si=Number.EPSILON||1e-14,ni=(t,e)=>e<t.length&&!t[e].skip&&t[e],oi=t=>"x"===t?"y":"x";function ai(t,e,i,s){const n=t.skip?e:t,o=e,a=i.skip?e:i,r=q(o,n),l=q(a,o);let h=r/(r+l),c=l/(r+l);h=isNaN(h)?0:h,c=isNaN(c)?0:c;const d=s*h,u=s*c;return{previous:{x:o.x-d*(a.x-n.x),y:o.y-d*(a.y-n.y)},next:{x:o.x+u*(a.x-n.x),y:o.y+u*(a.y-n.y)}}}function ri(t,e="x"){const i=oi(e),s=t.length,n=Array(s).fill(0),o=Array(s);let a,r,l,h=ni(t,0);for(a=0;a<s;++a)if(r=l,l=h,h=ni(t,a+1),l){if(h){const t=h[e]-l[e];n[a]=0!==t?(h[i]-l[i])/t:0}o[a]=r?h?F(n[a-1])!==F(n[a])?0:(n[a-1]+n[a])/2:n[a-1]:n[a]}!function(t,e,i){const s=t.length;let n,o,a,r,l,h=ni(t,0);for(let c=0;c<s-1;++c)l=h,h=ni(t,c+1),l&&h&&(V(e[c],0,si)?i[c]=i[c+1]=0:(n=i[c]/e[c],o=i[c+1]/e[c],r=Math.pow(n,2)+Math.pow(o,2),r<=9||(a=3/Math.sqrt(r),i[c]=n*a*e[c],i[c+1]=o*a*e[c])))}(t,n,o),function(t,e,i="x"){const s=oi(i),n=t.length;let o,a,r,l=ni(t,0);for(let h=0;h<n;++h){if(a=r,r=l,l=ni(t,h+1),!r)continue;const n=r[i],c=r[s];a&&(o=(n-a[i])/3,r[`cp1${i}`]=n-o,r[`cp1${s}`]=c-o*e[h]),l&&(o=(l[i]-n)/3,r[`cp2${i}`]=n+o,r[`cp2${s}`]=c+o*e[h])}}(t,o,e)}function li(t,e,i){return Math.max(Math.min(t,i),e)}function hi(t,e,i,s,n){let o,a,r,l;if(e.spanGaps&&(t=t.filter((t=>!t.skip))),"monotone"===e.cubicInterpolationMode)ri(t,n);else{let i=s?t[t.length-1]:t[0];for(o=0,a=t.length;o<a;++o)r=t[o],l=ai(i,r,t[Math.min(o+1,a-(s?0:1))%a],e.tension),r.cp1x=l.previous.x,r.cp1y=l.previous.y,r.cp2x=l.next.x,r.cp2y=l.next.y,i=r}e.capBezierPoints&&function(t,e){let i,s,n,o,a,r=Re(t[0],e);for(i=0,s=t.length;i<s;++i)a=o,o=r,r=i<s-1&&Re(t[i+1],e),o&&(n=t[i],a&&(n.cp1x=li(n.cp1x,e.left,e.right),n.cp1y=li(n.cp1y,e.top,e.bottom)),r&&(n.cp2x=li(n.cp2x,e.left,e.right),n.cp2y=li(n.cp2y,e.top,e.bottom)))}(t,i)}const ci=t=>0===t||1===t,di=(t,e,i)=>-Math.pow(2,10*(t-=1))*Math.sin((t-e)*O/i),ui=(t,e,i)=>Math.pow(2,-10*t)*Math.sin((t-e)*O/i)+1,fi={linear:t=>t,easeInQuad:t=>t*t,easeOutQuad:t=>-t*(t-2),easeInOutQuad:t=>(t/=.5)<1?.5*t*t:-.5*(--t*(t-2)-1),easeInCubic:t=>t*t*t,easeOutCubic:t=>(t-=1)*t*t+1,easeInOutCubic:t=>(t/=.5)<1?.5*t*t*t:.5*((t-=2)*t*t+2),easeInQuart:t=>t*t*t*t,easeOutQuart:t=>-((t-=1)*t*t*t-1),easeInOutQuart:t=>(t/=.5)<1?.5*t*t*t*t:-.5*((t-=2)*t*t*t-2),easeInQuint:t=>t*t*t*t*t,easeOutQuint:t=>(t-=1)*t*t*t*t+1,easeInOutQuint:t=>(t/=.5)<1?.5*t*t*t*t*t:.5*((t-=2)*t*t*t*t+2),easeInSine:t=>1-Math.cos(t*E),easeOutSine:t=>Math.sin(t*E),easeInOutSine:t=>-.5*(Math.cos(C*t)-1),easeInExpo:t=>0===t?0:Math.pow(2,10*(t-1)),easeOutExpo:t=>1===t?1:1-Math.pow(2,-10*t),easeInOutExpo:t=>ci(t)?t:t<.5?.5*Math.pow(2,10*(2*t-1)):.5*(2-Math.pow(2,-10*(2*t-1))),easeInCirc:t=>t>=1?t:-(Math.sqrt(1-t*t)-1),easeOutCirc:t=>Math.sqrt(1-(t-=1)*t),easeInOutCirc:t=>(t/=.5)<1?-.5*(Math.sqrt(1-t*t)-1):.5*(Math.sqrt(1-(t-=2)*t)+1),easeInElastic:t=>ci(t)?t:di(t,.075,.3),easeOutElastic:t=>ci(t)?t:ui(t,.075,.3),easeInOutElastic(t){const e=.1125;return ci(t)?t:t<.5?.5*di(2*t,e,.45):.5+.5*ui(2*t-1,e,.45)},easeInBack(t){const e=1.70158;return t*t*((e+1)*t-e)},easeOutBack(t){const e=1.70158;return(t-=1)*t*((e+1)*t+e)+1},easeInOutBack(t){let e=1.70158;return(t/=.5)<1?t*t*((1+(e*=1.525))*t-e)*.5:.5*((t-=2)*t*((1+(e*=1.525))*t+e)+2)},easeInBounce:t=>1-fi.easeOutBounce(1-t),easeOutBounce(t){const e=7.5625,i=2.75;return t<1/i?e*t*t:t<2/i?e*(t-=1.5/i)*t+.75:t<2.5/i?e*(t-=2.25/i)*t+.9375:e*(t-=2.625/i)*t+.984375},easeInOutBounce:t=>t<.5?.5*fi.easeInBounce(2*t):.5*fi.easeOutBounce(2*t-1)+.5};function gi(t,e,i,s){return{x:t.x+i*(e.x-t.x),y:t.y+i*(e.y-t.y)}}function pi(t,e,i,s){return{x:t.x+i*(e.x-t.x),y:"middle"===s?i<.5?t.y:e.y:"after"===s?i<1?t.y:e.y:i>0?e.y:t.y}}function mi(t,e,i,s){const n={x:t.cp2x,y:t.cp2y},o={x:e.cp1x,y:e.cp1y},a=gi(t,n,i),r=gi(n,o,i),l=gi(o,e,i),h=gi(a,r,i),c=gi(r,l,i);return gi(h,c,i)}const xi=/^(normal|(\d+(?:\.\d+)?)(px|em|%)?)$/,bi=/^(normal|italic|initial|inherit|unset|(oblique( -?[0-9]?[0-9]deg)?))$/;function _i(t,e){const i=(""+t).match(xi);if(!i||"normal"===i[1])return 1.2*e;switch(t=+i[2],i[3]){case"px":return t;case"%":t/=100}return e*t}const yi=t=>+t||0;function vi(t,e){const i={},s=o(e),n=s?Object.keys(e):e,a=o(t)?s?i=>l(t[i],t[e[i]]):e=>t[e]:()=>t;for(const t of n)i[t]=yi(a(t));return i}function Mi(t){return vi(t,{top:"y",right:"x",bottom:"y",left:"x"})}function wi(t){return vi(t,["topLeft","topRight","bottomLeft","bottomRight"])}function ki(t){const e=Mi(t);return e.width=e.left+e.right,e.height=e.top+e.bottom,e}function Si(t,e){t=t||{},e=e||ue.font;let i=l(t.size,e.size);"string"==typeof i&&(i=parseInt(i,10));let s=l(t.style,e.style);s&&!(""+s).match(bi)&&(console.warn('Invalid font style specified: "'+s+'"'),s=void 0);const n={family:l(t.family,e.family),lineHeight:_i(l(t.lineHeight,e.lineHeight),i),size:i,style:s,weight:l(t.weight,e.weight),string:""};return n.string=De(n),n}function Pi(t,e,i,s){let o,a,r,l=!0;for(o=0,a=t.length;o<a;++o)if(r=t[o],void 0!==r&&(void 0!==e&&"function"==typeof r&&(r=r(e),l=!1),void 0!==i&&n(r)&&(r=r[i%r.length],l=!1),void 0!==r))return s&&!l&&(s.cacheable=!1),r}function Di(t,e,i){const{min:s,max:n}=t,o=c(e,(n-s)/2),a=(t,e)=>i&&0===t?0:t+e;return{min:a(s,-Math.abs(o)),max:a(n,o)}}function Ci(t,e){return Object.assign(Object.create(t),e)}function Oi(t,e,i){return t?function(t,e){return{x:i=>t+t+e-i,setWidth(t){e=t},textAlign:t=>"center"===t?t:"right"===t?"left":"right",xPlus:(t,e)=>t-e,leftForLtr:(t,e)=>t-e}}(e,i):{x:t=>t,setWidth(t){},textAlign:t=>t,xPlus:(t,e)=>t+e,leftForLtr:(t,e)=>t}}function Ai(t,e){let i,s;"ltr"!==e&&"rtl"!==e||(i=t.canvas.style,s=[i.getPropertyValue("direction"),i.getPropertyPriority("direction")],i.setProperty("direction",e,"important"),t.prevTextDirection=s)}function Ti(t,e){void 0!==e&&(delete t.prevTextDirection,t.canvas.style.setProperty("direction",e[0],e[1]))}function Li(t){return"angle"===t?{between:Z,compare:K,normalize:G}:{between:tt,compare:(t,e)=>t-e,normalize:t=>t}}function Ei({start:t,end:e,count:i,loop:s,style:n}){return{start:t%i,end:e%i,loop:s&&(e-t+1)%i==0,style:n}}function Ri(t,e,i){if(!i)return[t];const{property:s,start:n,end:o}=i,a=e.length,{compare:r,between:l,normalize:h}=Li(s),{start:c,end:d,loop:u,style:f}=function(t,e,i){const{property:s,start:n,end:o}=i,{between:a,normalize:r}=Li(s),l=e.length;let h,c,{start:d,end:u,loop:f}=t;if(f){for(d+=l,u+=l,h=0,c=l;h<c&&a(r(e[d%l][s]),n,o);++h)d--,u--;d%=l,u%=l}return u<d&&(u+=l),{start:d,end:u,loop:f,style:t.style}}(t,e,i),g=[];let p,m,x,b=!1,_=null;const y=()=>b||l(n,x,p)&&0!==r(n,x),v=()=>!b||0===r(o,p)||l(o,x,p);for(let t=c,i=c;t<=d;++t)m=e[t%a],m.skip||(p=h(m[s]),p!==x&&(b=l(p,n,o),null===_&&y()&&(_=0===r(p,n)?t:i),null!==_&&v()&&(g.push(Ei({start:_,end:t,loop:u,count:a,style:f})),_=null),i=t,x=p));return null!==_&&g.push(Ei({start:_,end:d,loop:u,count:a,style:f})),g}function Ii(t,e){const i=[],s=t.segments;for(let n=0;n<s.length;n++){const o=Ri(s[n],t.points,e);o.length&&i.push(...o)}return i}function zi(t,e){const i=t.points,s=t.options.spanGaps,n=i.length;if(!n)return[];const o=!!t._loop,{start:a,end:r}=function(t,e,i,s){let n=0,o=e-1;if(i&&!s)for(;n<e&&!t[n].skip;)n++;for(;n<e&&t[n].skip;)n++;for(n%=e,i&&(o+=n);o>n&&t[o%e].skip;)o--;return o%=e,{start:n,end:o}}(i,n,o,s);if(!0===s)return Fi(t,[{start:a,end:r,loop:o}],i,e);return Fi(t,function(t,e,i,s){const n=t.length,o=[];let a,r=e,l=t[e];for(a=e+1;a<=i;++a){const i=t[a%n];i.skip||i.stop?l.skip||(s=!1,o.push({start:e%n,end:(a-1)%n,loop:s}),e=r=i.stop?a:null):(r=a,l.skip&&(e=a)),l=i}return null!==r&&o.push({start:e%n,end:r%n,loop:s}),o}(i,a,r<a?r+n:r,!!t._fullLoop&&0===a&&r===n-1),i,e)}function Fi(t,e,i,s){return s&&s.setContext&&i?function(t,e,i,s){const n=t._chart.getContext(),o=Vi(t.options),{_datasetIndex:a,options:{spanGaps:r}}=t,l=i.length,h=[];let c=o,d=e[0].start,u=d;function f(t,e,s,n){const o=r?-1:1;if(t!==e){for(t+=l;i[t%l].skip;)t-=o;for(;i[e%l].skip;)e+=o;t%l!=e%l&&(h.push({start:t%l,end:e%l,loop:s,style:n}),c=n,d=e%l)}}for(const t of e){d=r?d:t.start;let e,o=i[d%l];for(u=d+1;u<=t.end;u++){const r=i[u%l];e=Vi(s.setContext(Ci(n,{type:"segment",p0:o,p1:r,p0DataIndex:(u-1)%l,p1DataIndex:u%l,datasetIndex:a}))),Bi(e,c)&&f(d,u-1,t.loop,c),o=r,c=e}d<u-1&&f(d,u-1,t.loop,c)}return h}(t,e,i,s):e}function Vi(t){return{backgroundColor:t.backgroundColor,borderCapStyle:t.borderCapStyle,borderDash:t.borderDash,borderDashOffset:t.borderDashOffset,borderJoinStyle:t.borderJoinStyle,borderWidth:t.borderWidth,borderColor:t.borderColor}}function Bi(t,e){if(!e)return!1;const i=[],s=function(t,e){return Jt(e)?(i.includes(e)||i.push(e),i.indexOf(e)):e};return JSON.stringify(t,s)!==JSON.stringify(e,s)}var Wi=Object.freeze({__proto__:null,HALF_PI:E,INFINITY:T,PI:C,PITAU:A,QUARTER_PI:R,RAD_PER_DEG:L,TAU:O,TWO_THIRDS_PI:I,_addGrace:Di,_alignPixel:Ae,_alignStartEnd:ft,_angleBetween:Z,_angleDiff:K,_arrayUnique:lt,_attachContext:$e,_bezierCurveTo:Ve,_bezierInterpolation:mi,_boundSegment:Ri,_boundSegments:Ii,_capitalize:w,_computeSegments:zi,_createResolver:je,_decimalPlaces:U,_deprecated:function(t,e,i,s){void 0!==e&&console.warn(t+': "'+i+'" is deprecated. Please use "'+s+'" instead')},_descriptors:Ye,_elementsEqual:f,_factorize:W,_filterBetween:nt,_getParentNode:ge,_getStartAndCountOfVisiblePoints:pt,_int16Range:Q,_isBetween:tt,_isClickEvent:D,_isDomSupported:fe,_isPointInArea:Re,_limitValue:J,_longestText:Oe,_lookup:et,_lookupByKey:it,_measureText:Ce,_merger:m,_mergerIf:_,_normalizeAngle:G,_parseObjectDataRadialScale:ii,_pointInLine:gi,_readValueToProps:vi,_rlookupByKey:st,_scaleRangesChanged:mt,_setMinAndMaxByKey:j,_splitKey:v,_steppedInterpolation:pi,_steppedLineTo:Fe,_textX:gt,_toLeftRightCenter:ut,_updateBezierControlPoints:hi,addRoundedRectPath:He,almostEquals:V,almostWhole:H,callback:d,clearCanvas:Te,clipArea:Ie,clone:g,color:Qt,createContext:Ci,debounce:dt,defined:k,distanceBetweenPoints:q,drawPoint:Le,drawPointLegend:Ee,each:u,easingEffects:fi,finiteOrDefault:r,fontString:function(t,e,i){return e+" "+t+"px "+i},formatNumber:ne,getAngleFromPoint:X,getHoverColor:te,getMaximumSize:we,getRelativePosition:ve,getRtlAdapter:Oi,getStyle:xe,isArray:n,isFinite:a,isFunction:S,isNullOrUndef:s,isNumber:N,isObject:o,isPatternOrGradient:Jt,listenArrayEvents:at,log10:z,merge:x,mergeIf:b,niceNum:B,noop:e,overrideTextDirection:Ai,readUsedSize:Pe,renderText:Ne,requestAnimFrame:ht,resolve:Pi,resolveObjectKey:M,restoreTextDirection:Ti,retinaScale:ke,setsEqual:P,sign:F,splineCurve:ai,splineCurveMonotone:ri,supportsEventListenerOptions:Se,throttled:ct,toDegrees:Y,toDimension:c,toFont:Si,toFontString:De,toLineHeight:_i,toPadding:ki,toPercentage:h,toRadians:$,toTRBL:Mi,toTRBLCorners:wi,uid:i,unclipArea:ze,unlistenArrayEvents:rt,valueOrDefault:l});function Ni(t,e,i,s){const{controller:n,data:o,_sorted:a}=t,r=n._cachedMeta.iScale;if(r&&e===r.axis&&"r"!==e&&a&&o.length){const t=r._reversePixels?st:it;if(!s)return t(o,e,i);if(n._sharedOptions){const s=o[0],n="function"==typeof s.getRange&&s.getRange(e);if(n){const s=t(o,e,i-n),a=t(o,e,i+n);return{lo:s.lo,hi:a.hi}}}}return{lo:0,hi:o.length-1}}function Hi(t,e,i,s,n){const o=t.getSortedVisibleDatasetMetas(),a=i[e];for(let t=0,i=o.length;t<i;++t){const{index:i,data:r}=o[t],{lo:l,hi:h}=Ni(o[t],e,a,n);for(let t=l;t<=h;++t){const e=r[t];e.skip||s(e,i,t)}}}function ji(t,e,i,s,n){const o=[];if(!n&&!t.isPointInArea(e))return o;return Hi(t,i,e,(function(i,a,r){(n||Re(i,t.chartArea,0))&&i.inRange(e.x,e.y,s)&&o.push({element:i,datasetIndex:a,index:r})}),!0),o}function $i(t,e,i,s,n,o){let a=[];const r=function(t){const e=-1!==t.indexOf("x"),i=-1!==t.indexOf("y");return function(t,s){const n=e?Math.abs(t.x-s.x):0,o=i?Math.abs(t.y-s.y):0;return Math.sqrt(Math.pow(n,2)+Math.pow(o,2))}}(i);let l=Number.POSITIVE_INFINITY;return Hi(t,i,e,(function(i,h,c){const d=i.inRange(e.x,e.y,n);if(s&&!d)return;const u=i.getCenterPoint(n);if(!(!!o||t.isPointInArea(u))&&!d)return;const f=r(e,u);f<l?(a=[{element:i,datasetIndex:h,index:c}],l=f):f===l&&a.push({element:i,datasetIndex:h,index:c})})),a}function Yi(t,e,i,s,n,o){return o||t.isPointInArea(e)?"r"!==i||s?$i(t,e,i,s,n,o):function(t,e,i,s){let n=[];return Hi(t,i,e,(function(t,i,o){const{startAngle:a,endAngle:r}=t.getProps(["startAngle","endAngle"],s),{angle:l}=X(t,{x:e.x,y:e.y});Z(l,a,r)&&n.push({element:t,datasetIndex:i,index:o})})),n}(t,e,i,n):[]}function Ui(t,e,i,s,n){const o=[],a="x"===i?"inXRange":"inYRange";let r=!1;return Hi(t,i,e,((t,s,l)=>{t[a](e[i],n)&&(o.push({element:t,datasetIndex:s,index:l}),r=r||t.inRange(e.x,e.y,n))})),s&&!r?[]:o}var Xi={evaluateInteractionItems:Hi,modes:{index(t,e,i,s){const n=ve(e,t),o=i.axis||"x",a=i.includeInvisible||!1,r=i.intersect?ji(t,n,o,s,a):Yi(t,n,o,!1,s,a),l=[];return r.length?(t.getSortedVisibleDatasetMetas().forEach((t=>{const e=r[0].index,i=t.data[e];i&&!i.skip&&l.push({element:i,datasetIndex:t.index,index:e})})),l):[]},dataset(t,e,i,s){const n=ve(e,t),o=i.axis||"xy",a=i.includeInvisible||!1;let r=i.intersect?ji(t,n,o,s,a):Yi(t,n,o,!1,s,a);if(r.length>0){const e=r[0].datasetIndex,i=t.getDatasetMeta(e).data;r=[];for(let t=0;t<i.length;++t)r.push({element:i[t],datasetIndex:e,index:t})}return r},point:(t,e,i,s)=>ji(t,ve(e,t),i.axis||"xy",s,i.includeInvisible||!1),nearest(t,e,i,s){const n=ve(e,t),o=i.axis||"xy",a=i.includeInvisible||!1;return Yi(t,n,o,i.intersect,s,a)},x:(t,e,i,s)=>Ui(t,ve(e,t),"x",i.intersect,s),y:(t,e,i,s)=>Ui(t,ve(e,t),"y",i.intersect,s)}};const qi=["left","top","right","bottom"];function Ki(t,e){return t.filter((t=>t.pos===e))}function Gi(t,e){return t.filter((t=>-1===qi.indexOf(t.pos)&&t.box.axis===e))}function Zi(t,e){return t.sort(((t,i)=>{const s=e?i:t,n=e?t:i;return s.weight===n.weight?s.index-n.index:s.weight-n.weight}))}function Ji(t,e){const i=function(t){const e={};for(const i of t){const{stack:t,pos:s,stackWeight:n}=i;if(!t||!qi.includes(s))continue;const o=e[t]||(e[t]={count:0,placed:0,weight:0,size:0});o.count++,o.weight+=n}return e}(t),{vBoxMaxWidth:s,hBoxMaxHeight:n}=e;let o,a,r;for(o=0,a=t.length;o<a;++o){r=t[o];const{fullSize:a}=r.box,l=i[r.stack],h=l&&r.stackWeight/l.weight;r.horizontal?(r.width=h?h*s:a&&e.availableWidth,r.height=n):(r.width=s,r.height=h?h*n:a&&e.availableHeight)}return i}function Qi(t,e,i,s){return Math.max(t[i],e[i])+Math.max(t[s],e[s])}function ts(t,e){t.top=Math.max(t.top,e.top),t.left=Math.max(t.left,e.left),t.bottom=Math.max(t.bottom,e.bottom),t.right=Math.max(t.right,e.right)}function es(t,e,i,s){const{pos:n,box:a}=i,r=t.maxPadding;if(!o(n)){i.size&&(t[n]-=i.size);const e=s[i.stack]||{size:0,count:1};e.size=Math.max(e.size,i.horizontal?a.height:a.width),i.size=e.size/e.count,t[n]+=i.size}a.getPadding&&ts(r,a.getPadding());const l=Math.max(0,e.outerWidth-Qi(r,t,"left","right")),h=Math.max(0,e.outerHeight-Qi(r,t,"top","bottom")),c=l!==t.w,d=h!==t.h;return t.w=l,t.h=h,i.horizontal?{same:c,other:d}:{same:d,other:c}}function is(t,e){const i=e.maxPadding;function s(t){const s={left:0,top:0,right:0,bottom:0};return t.forEach((t=>{s[t]=Math.max(e[t],i[t])})),s}return s(t?["left","right"]:["top","bottom"])}function ss(t,e,i,s){const n=[];let o,a,r,l,h,c;for(o=0,a=t.length,h=0;o<a;++o){r=t[o],l=r.box,l.update(r.width||e.w,r.height||e.h,is(r.horizontal,e));const{same:a,other:d}=es(e,i,r,s);h|=a&&n.length,c=c||d,l.fullSize||n.push(r)}return h&&ss(n,e,i,s)||c}function ns(t,e,i,s,n){t.top=i,t.left=e,t.right=e+s,t.bottom=i+n,t.width=s,t.height=n}function os(t,e,i,s){const n=i.padding;let{x:o,y:a}=e;for(const r of t){const t=r.box,l=s[r.stack]||{count:1,placed:0,weight:1},h=r.stackWeight/l.weight||1;if(r.horizontal){const s=e.w*h,o=l.size||t.height;k(l.start)&&(a=l.start),t.fullSize?ns(t,n.left,a,i.outerWidth-n.right-n.left,o):ns(t,e.left+l.placed,a,s,o),l.start=a,l.placed+=s,a=t.bottom}else{const s=e.h*h,a=l.size||t.width;k(l.start)&&(o=l.start),t.fullSize?ns(t,o,n.top,a,i.outerHeight-n.bottom-n.top):ns(t,o,e.top+l.placed,a,s),l.start=o,l.placed+=s,o=t.right}}e.x=o,e.y=a}var as={addBox(t,e){t.boxes||(t.boxes=[]),e.fullSize=e.fullSize||!1,e.position=e.position||"top",e.weight=e.weight||0,e._layers=e._layers||function(){return[{z:0,draw(t){e.draw(t)}}]},t.boxes.push(e)},removeBox(t,e){const i=t.boxes?t.boxes.indexOf(e):-1;-1!==i&&t.boxes.splice(i,1)},configure(t,e,i){e.fullSize=i.fullSize,e.position=i.position,e.weight=i.weight},update(t,e,i,s){if(!t)return;const n=ki(t.options.layout.padding),o=Math.max(e-n.width,0),a=Math.max(i-n.height,0),r=function(t){const e=function(t){const e=[];let i,s,n,o,a,r;for(i=0,s=(t||[]).length;i<s;++i)n=t[i],({position:o,options:{stack:a,stackWeight:r=1}}=n),e.push({index:i,box:n,pos:o,horizontal:n.isHorizontal(),weight:n.weight,stack:a&&o+a,stackWeight:r});return e}(t),i=Zi(e.filter((t=>t.box.fullSize)),!0),s=Zi(Ki(e,"left"),!0),n=Zi(Ki(e,"right")),o=Zi(Ki(e,"top"),!0),a=Zi(Ki(e,"bottom")),r=Gi(e,"x"),l=Gi(e,"y");return{fullSize:i,leftAndTop:s.concat(o),rightAndBottom:n.concat(l).concat(a).concat(r),chartArea:Ki(e,"chartArea"),vertical:s.concat(n).concat(l),horizontal:o.concat(a).concat(r)}}(t.boxes),l=r.vertical,h=r.horizontal;u(t.boxes,(t=>{"function"==typeof t.beforeLayout&&t.beforeLayout()}));const c=l.reduce(((t,e)=>e.box.options&&!1===e.box.options.display?t:t+1),0)||1,d=Object.freeze({outerWidth:e,outerHeight:i,padding:n,availableWidth:o,availableHeight:a,vBoxMaxWidth:o/2/c,hBoxMaxHeight:a/2}),f=Object.assign({},n);ts(f,ki(s));const g=Object.assign({maxPadding:f,w:o,h:a,x:n.left,y:n.top},n),p=Ji(l.concat(h),d);ss(r.fullSize,g,d,p),ss(l,g,d,p),ss(h,g,d,p)&&ss(l,g,d,p),function(t){const e=t.maxPadding;function i(i){const s=Math.max(e[i]-t[i],0);return t[i]+=s,s}t.y+=i("top"),t.x+=i("left"),i("right"),i("bottom")}(g),os(r.leftAndTop,g,d,p),g.x+=g.w,g.y+=g.h,os(r.rightAndBottom,g,d,p),t.chartArea={left:g.left,top:g.top,right:g.left+g.w,bottom:g.top+g.h,height:g.h,width:g.w},u(r.chartArea,(e=>{const i=e.box;Object.assign(i,t.chartArea),i.update(g.w,g.h,{left:0,top:0,right:0,bottom:0})}))}};class rs{acquireContext(t,e){}releaseContext(t){return!1}addEventListener(t,e,i){}removeEventListener(t,e,i){}getDevicePixelRatio(){return 1}getMaximumSize(t,e,i,s){return e=Math.max(0,e||t.width),i=i||t.height,{width:e,height:Math.max(0,s?Math.floor(e/s):i)}}isAttached(t){return!0}updateConfig(t){}}class ls extends rs{acquireContext(t){return t&&t.getContext&&t.getContext("2d")||null}updateConfig(t){t.options.animation=!1}}const hs="$chartjs",cs={touchstart:"mousedown",touchmove:"mousemove",touchend:"mouseup",pointerenter:"mouseenter",pointerdown:"mousedown",pointermove:"mousemove",pointerup:"mouseup",pointerleave:"mouseout",pointerout:"mouseout"},ds=t=>null===t||""===t;const us=!!Se&&{passive:!0};function fs(t,e,i){t&&t.canvas&&t.canvas.removeEventListener(e,i,us)}function gs(t,e){for(const i of t)if(i===e||i.contains(e))return!0}function ps(t,e,i){const s=t.canvas,n=new MutationObserver((t=>{let e=!1;for(const i of t)e=e||gs(i.addedNodes,s),e=e&&!gs(i.removedNodes,s);e&&i()}));return n.observe(document,{childList:!0,subtree:!0}),n}function ms(t,e,i){const s=t.canvas,n=new MutationObserver((t=>{let e=!1;for(const i of t)e=e||gs(i.removedNodes,s),e=e&&!gs(i.addedNodes,s);e&&i()}));return n.observe(document,{childList:!0,subtree:!0}),n}const xs=new Map;let bs=0;function _s(){const t=window.devicePixelRatio;t!==bs&&(bs=t,xs.forEach(((e,i)=>{i.currentDevicePixelRatio!==t&&e()})))}function ys(t,e,i){const s=t.canvas,n=s&&ge(s);if(!n)return;const o=ct(((t,e)=>{const s=n.clientWidth;i(t,e),s<n.clientWidth&&i()}),window),a=new ResizeObserver((t=>{const e=t[0],i=e.contentRect.width,s=e.contentRect.height;0===i&&0===s||o(i,s)}));return a.observe(n),function(t,e){xs.size||window.addEventListener("resize",_s),xs.set(t,e)}(t,o),a}function vs(t,e,i){i&&i.disconnect(),"resize"===e&&function(t){xs.delete(t),xs.size||window.removeEventListener("resize",_s)}(t)}function Ms(t,e,i){const s=t.canvas,n=ct((e=>{null!==t.ctx&&i(function(t,e){const i=cs[t.type]||t.type,{x:s,y:n}=ve(t,e);return{type:i,chart:e,native:t,x:void 0!==s?s:null,y:void 0!==n?n:null}}(e,t))}),t);return function(t,e,i){t&&t.addEventListener(e,i,us)}(s,e,n),n}class ws extends rs{acquireContext(t,e){const i=t&&t.getContext&&t.getContext("2d");return i&&i.canvas===t?(function(t,e){const i=t.style,s=t.getAttribute("height"),n=t.getAttribute("width");if(t[hs]={initial:{height:s,width:n,style:{display:i.display,height:i.height,width:i.width}}},i.display=i.display||"block",i.boxSizing=i.boxSizing||"border-box",ds(n)){const e=Pe(t,"width");void 0!==e&&(t.width=e)}if(ds(s))if(""===t.style.height)t.height=t.width/(e||2);else{const e=Pe(t,"height");void 0!==e&&(t.height=e)}}(t,e),i):null}releaseContext(t){const e=t.canvas;if(!e[hs])return!1;const i=e[hs].initial;["height","width"].forEach((t=>{const n=i[t];s(n)?e.removeAttribute(t):e.setAttribute(t,n)}));const n=i.style||{};return Object.keys(n).forEach((t=>{e.style[t]=n[t]})),e.width=e.width,delete e[hs],!0}addEventListener(t,e,i){this.removeEventListener(t,e);const s=t.$proxies||(t.$proxies={}),n={attach:ps,detach:ms,resize:ys}[e]||Ms;s[e]=n(t,e,i)}removeEventListener(t,e){const i=t.$proxies||(t.$proxies={}),s=i[e];if(!s)return;({attach:vs,detach:vs,resize:vs}[e]||fs)(t,e,s),i[e]=void 0}getDevicePixelRatio(){return window.devicePixelRatio}getMaximumSize(t,e,i,s){return we(t,e,i,s)}isAttached(t){const e=t&&ge(t);return!(!e||!e.isConnected)}}function ks(t){return!fe()||"undefined"!=typeof OffscreenCanvas&&t instanceof OffscreenCanvas?ls:ws}var Ss=Object.freeze({__proto__:null,BasePlatform:rs,BasicPlatform:ls,DomPlatform:ws,_detectPlatform:ks});const Ps="transparent",Ds={boolean:(t,e,i)=>i>.5?e:t,color(t,e,i){const s=Qt(t||Ps),n=s.valid&&Qt(e||Ps);return n&&n.valid?n.mix(s,i).hexString():e},number:(t,e,i)=>t+(e-t)*i};class Cs{constructor(t,e,i,s){const n=e[i];s=Pi([t.to,s,n,t.from]);const o=Pi([t.from,n,s]);this._active=!0,this._fn=t.fn||Ds[t.type||typeof o],this._easing=fi[t.easing]||fi.linear,this._start=Math.floor(Date.now()+(t.delay||0)),this._duration=this._total=Math.floor(t.duration),this._loop=!!t.loop,this._target=e,this._prop=i,this._from=o,this._to=s,this._promises=void 0}active(){return this._active}update(t,e,i){if(this._active){this._notify(!1);const s=this._target[this._prop],n=i-this._start,o=this._duration-n;this._start=i,this._duration=Math.floor(Math.max(o,t.duration)),this._total+=n,this._loop=!!t.loop,this._to=Pi([t.to,e,s,t.from]),this._from=Pi([t.from,s,e])}}cancel(){this._active&&(this.tick(Date.now()),this._active=!1,this._notify(!1))}tick(t){const e=t-this._start,i=this._duration,s=this._prop,n=this._from,o=this._loop,a=this._to;let r;if(this._active=n!==a&&(o||e<i),!this._active)return this._target[s]=a,void this._notify(!0);e<0?this._target[s]=n:(r=e/i%2,r=o&&r>1?2-r:r,r=this._easing(Math.min(1,Math.max(0,r))),this._target[s]=this._fn(n,a,r))}wait(){const t=this._promises||(this._promises=[]);return new Promise(((e,i)=>{t.push({res:e,rej:i})}))}_notify(t){const e=t?"res":"rej",i=this._promises||[];for(let t=0;t<i.length;t++)i[t][e]()}}class Os{constructor(t,e){this._chart=t,this._properties=new Map,this.configure(e)}configure(t){if(!o(t))return;const e=Object.keys(ue.animation),i=this._properties;Object.getOwnPropertyNames(t).forEach((s=>{const a=t[s];if(!o(a))return;const r={};for(const t of e)r[t]=a[t];(n(a.properties)&&a.properties||[s]).forEach((t=>{t!==s&&i.has(t)||i.set(t,r)}))}))}_animateOptions(t,e){const i=e.options,s=function(t,e){if(!e)return;let i=t.options;if(!i)return void(t.options=e);i.$shared&&(t.options=i=Object.assign({},i,{$shared:!1,$animations:{}}));return i}(t,i);if(!s)return[];const n=this._createAnimations(s,i);return i.$shared&&function(t,e){const i=[],s=Object.keys(e);for(let e=0;e<s.length;e++){const n=t[s[e]];n&&n.active()&&i.push(n.wait())}return Promise.all(i)}(t.options.$animations,i).then((()=>{t.options=i}),(()=>{})),n}_createAnimations(t,e){const i=this._properties,s=[],n=t.$animations||(t.$animations={}),o=Object.keys(e),a=Date.now();let r;for(r=o.length-1;r>=0;--r){const l=o[r];if("$"===l.charAt(0))continue;if("options"===l){s.push(...this._animateOptions(t,e));continue}const h=e[l];let c=n[l];const d=i.get(l);if(c){if(d&&c.active()){c.update(d,h,a);continue}c.cancel()}d&&d.duration?(n[l]=c=new Cs(d,t,l,h),s.push(c)):t[l]=h}return s}update(t,e){if(0===this._properties.size)return void Object.assign(t,e);const i=this._createAnimations(t,e);return i.length?(bt.add(this._chart,i),!0):void 0}}function As(t,e){const i=t&&t.options||{},s=i.reverse,n=void 0===i.min?e:0,o=void 0===i.max?e:0;return{start:s?o:n,end:s?n:o}}function Ts(t,e){const i=[],s=t._getSortedDatasetMetas(e);let n,o;for(n=0,o=s.length;n<o;++n)i.push(s[n].index);return i}function Ls(t,e,i,s={}){const n=t.keys,o="single"===s.mode;let r,l,h,c;if(null!==e){for(r=0,l=n.length;r<l;++r){if(h=+n[r],h===i){if(s.all)continue;break}c=t.values[h],a(c)&&(o||0===e||F(e)===F(c))&&(e+=c)}return e}}function Es(t,e){const i=t&&t.options.stacked;return i||void 0===i&&void 0!==e.stack}function Rs(t,e,i){const s=t[e]||(t[e]={});return s[i]||(s[i]={})}function Is(t,e,i,s){for(const n of e.getMatchingVisibleMetas(s).reverse()){const e=t[n.index];if(i&&e>0||!i&&e<0)return n.index}return null}function zs(t,e){const{chart:i,_cachedMeta:s}=t,n=i._stacks||(i._stacks={}),{iScale:o,vScale:a,index:r}=s,l=o.axis,h=a.axis,c=function(t,e,i){return`${t.id}.${e.id}.${i.stack||i.type}`}(o,a,s),d=e.length;let u;for(let t=0;t<d;++t){const i=e[t],{[l]:o,[h]:d}=i;u=(i._stacks||(i._stacks={}))[h]=Rs(n,c,o),u[r]=d,u._top=Is(u,a,!0,s.type),u._bottom=Is(u,a,!1,s.type);(u._visualValues||(u._visualValues={}))[r]=d}}function Fs(t,e){const i=t.scales;return Object.keys(i).filter((t=>i[t].axis===e)).shift()}function Vs(t,e){const i=t.controller.index,s=t.vScale&&t.vScale.axis;if(s){e=e||t._parsed;for(const t of e){const e=t._stacks;if(!e||void 0===e[s]||void 0===e[s][i])return;delete e[s][i],void 0!==e[s]._visualValues&&void 0!==e[s]._visualValues[i]&&delete e[s]._visualValues[i]}}}const Bs=t=>"reset"===t||"none"===t,Ws=(t,e)=>e?t:Object.assign({},t);class Ns{static defaults={};static datasetElementType=null;static dataElementType=null;constructor(t,e){this.chart=t,this._ctx=t.ctx,this.index=e,this._cachedDataOpts={},this._cachedMeta=this.getMeta(),this._type=this._cachedMeta.type,this.options=void 0,this._parsing=!1,this._data=void 0,this._objectData=void 0,this._sharedOptions=void 0,this._drawStart=void 0,this._drawCount=void 0,this.enableOptionSharing=!1,this.supportsDecimation=!1,this.$context=void 0,this._syncList=[],this.datasetElementType=new.target.datasetElementType,this.dataElementType=new.target.dataElementType,this.initialize()}initialize(){const t=this._cachedMeta;this.configure(),this.linkScales(),t._stacked=Es(t.vScale,t),this.addElements(),this.options.fill&&!this.chart.isPluginEnabled("filler")&&console.warn("Tried to use the 'fill' option without the 'Filler' plugin enabled. Please import and register the 'Filler' plugin and make sure it is not disabled in the options")}updateIndex(t){this.index!==t&&Vs(this._cachedMeta),this.index=t}linkScales(){const t=this.chart,e=this._cachedMeta,i=this.getDataset(),s=(t,e,i,s)=>"x"===t?e:"r"===t?s:i,n=e.xAxisID=l(i.xAxisID,Fs(t,"x")),o=e.yAxisID=l(i.yAxisID,Fs(t,"y")),a=e.rAxisID=l(i.rAxisID,Fs(t,"r")),r=e.indexAxis,h=e.iAxisID=s(r,n,o,a),c=e.vAxisID=s(r,o,n,a);e.xScale=this.getScaleForId(n),e.yScale=this.getScaleForId(o),e.rScale=this.getScaleForId(a),e.iScale=this.getScaleForId(h),e.vScale=this.getScaleForId(c)}getDataset(){return this.chart.data.datasets[this.index]}getMeta(){return this.chart.getDatasetMeta(this.index)}getScaleForId(t){return this.chart.scales[t]}_getOtherScale(t){const e=this._cachedMeta;return t===e.iScale?e.vScale:e.iScale}reset(){this._update("reset")}_destroy(){const t=this._cachedMeta;this._data&&rt(this._data,this),t._stacked&&Vs(t)}_dataCheck(){const t=this.getDataset(),e=t.data||(t.data=[]),i=this._data;if(o(e)){const t=this._cachedMeta;this._data=function(t,e){const{iScale:i,vScale:s}=e,n="x"===i.axis?"x":"y",o="x"===s.axis?"x":"y",a=Object.keys(t),r=new Array(a.length);let l,h,c;for(l=0,h=a.length;l<h;++l)c=a[l],r[l]={[n]:c,[o]:t[c]};return r}(e,t)}else if(i!==e){if(i){rt(i,this);const t=this._cachedMeta;Vs(t),t._parsed=[]}e&&Object.isExtensible(e)&&at(e,this),this._syncList=[],this._data=e}}addElements(){const t=this._cachedMeta;this._dataCheck(),this.datasetElementType&&(t.dataset=new this.datasetElementType)}buildOrUpdateElements(t){const e=this._cachedMeta,i=this.getDataset();let s=!1;this._dataCheck();const n=e._stacked;e._stacked=Es(e.vScale,e),e.stack!==i.stack&&(s=!0,Vs(e),e.stack=i.stack),this._resyncElements(t),(s||n!==e._stacked)&&zs(this,e._parsed)}configure(){const t=this.chart.config,e=t.datasetScopeKeys(this._type),i=t.getOptionScopes(this.getDataset(),e,!0);this.options=t.createResolver(i,this.getContext()),this._parsing=this.options.parsing,this._cachedDataOpts={}}parse(t,e){const{_cachedMeta:i,_data:s}=this,{iScale:a,_stacked:r}=i,l=a.axis;let h,c,d,u=0===t&&e===s.length||i._sorted,f=t>0&&i._parsed[t-1];if(!1===this._parsing)i._parsed=s,i._sorted=!0,d=s;else{d=n(s[t])?this.parseArrayData(i,s,t,e):o(s[t])?this.parseObjectData(i,s,t,e):this.parsePrimitiveData(i,s,t,e);const a=()=>null===c[l]||f&&c[l]<f[l];for(h=0;h<e;++h)i._parsed[h+t]=c=d[h],u&&(a()&&(u=!1),f=c);i._sorted=u}r&&zs(this,d)}parsePrimitiveData(t,e,i,s){const{iScale:n,vScale:o}=t,a=n.axis,r=o.axis,l=n.getLabels(),h=n===o,c=new Array(s);let d,u,f;for(d=0,u=s;d<u;++d)f=d+i,c[d]={[a]:h||n.parse(l[f],f),[r]:o.parse(e[f],f)};return c}parseArrayData(t,e,i,s){const{xScale:n,yScale:o}=t,a=new Array(s);let r,l,h,c;for(r=0,l=s;r<l;++r)h=r+i,c=e[h],a[r]={x:n.parse(c[0],h),y:o.parse(c[1],h)};return a}parseObjectData(t,e,i,s){const{xScale:n,yScale:o}=t,{xAxisKey:a="x",yAxisKey:r="y"}=this._parsing,l=new Array(s);let h,c,d,u;for(h=0,c=s;h<c;++h)d=h+i,u=e[d],l[h]={x:n.parse(M(u,a),d),y:o.parse(M(u,r),d)};return l}getParsed(t){return this._cachedMeta._parsed[t]}getDataElement(t){return this._cachedMeta.data[t]}applyStack(t,e,i){const s=this.chart,n=this._cachedMeta,o=e[t.axis];return Ls({keys:Ts(s,!0),values:e._stacks[t.axis]._visualValues},o,n.index,{mode:i})}updateRangeFromParsed(t,e,i,s){const n=i[e.axis];let o=null===n?NaN:n;const a=s&&i._stacks[e.axis];s&&a&&(s.values=a,o=Ls(s,n,this._cachedMeta.index)),t.min=Math.min(t.min,o),t.max=Math.max(t.max,o)}getMinMax(t,e){const i=this._cachedMeta,s=i._parsed,n=i._sorted&&t===i.iScale,o=s.length,r=this._getOtherScale(t),l=((t,e,i)=>t&&!e.hidden&&e._stacked&&{keys:Ts(i,!0),values:null})(e,i,this.chart),h={min:Number.POSITIVE_INFINITY,max:Number.NEGATIVE_INFINITY},{min:c,max:d}=function(t){const{min:e,max:i,minDefined:s,maxDefined:n}=t.getUserBounds();return{min:s?e:Number.NEGATIVE_INFINITY,max:n?i:Number.POSITIVE_INFINITY}}(r);let u,f;function g(){f=s[u];const e=f[r.axis];return!a(f[t.axis])||c>e||d<e}for(u=0;u<o&&(g()||(this.updateRangeFromParsed(h,t,f,l),!n));++u);if(n)for(u=o-1;u>=0;--u)if(!g()){this.updateRangeFromParsed(h,t,f,l);break}return h}getAllParsedValues(t){const e=this._cachedMeta._parsed,i=[];let s,n,o;for(s=0,n=e.length;s<n;++s)o=e[s][t.axis],a(o)&&i.push(o);return i}getMaxOverflow(){return!1}getLabelAndValue(t){const e=this._cachedMeta,i=e.iScale,s=e.vScale,n=this.getParsed(t);return{label:i?""+i.getLabelForValue(n[i.axis]):"",value:s?""+s.getLabelForValue(n[s.axis]):""}}_update(t){const e=this._cachedMeta;this.update(t||"default"),e._clip=function(t){let e,i,s,n;return o(t)?(e=t.top,i=t.right,s=t.bottom,n=t.left):e=i=s=n=t,{top:e,right:i,bottom:s,left:n,disabled:!1===t}}(l(this.options.clip,function(t,e,i){if(!1===i)return!1;const s=As(t,i),n=As(e,i);return{top:n.end,right:s.end,bottom:n.start,left:s.start}}(e.xScale,e.yScale,this.getMaxOverflow())))}update(t){}draw(){const t=this._ctx,e=this.chart,i=this._cachedMeta,s=i.data||[],n=e.chartArea,o=[],a=this._drawStart||0,r=this._drawCount||s.length-a,l=this.options.drawActiveElementsOnTop;let h;for(i.dataset&&i.dataset.draw(t,n,a,r),h=a;h<a+r;++h){const e=s[h];e.hidden||(e.active&&l?o.push(e):e.draw(t,n))}for(h=0;h<o.length;++h)o[h].draw(t,n)}getStyle(t,e){const i=e?"active":"default";return void 0===t&&this._cachedMeta.dataset?this.resolveDatasetElementOptions(i):this.resolveDataElementOptions(t||0,i)}getContext(t,e,i){const s=this.getDataset();let n;if(t>=0&&t<this._cachedMeta.data.length){const e=this._cachedMeta.data[t];n=e.$context||(e.$context=function(t,e,i){return Ci(t,{active:!1,dataIndex:e,parsed:void 0,raw:void 0,element:i,index:e,mode:"default",type:"data"})}(this.getContext(),t,e)),n.parsed=this.getParsed(t),n.raw=s.data[t],n.index=n.dataIndex=t}else n=this.$context||(this.$context=function(t,e){return Ci(t,{active:!1,dataset:void 0,datasetIndex:e,index:e,mode:"default",type:"dataset"})}(this.chart.getContext(),this.index)),n.dataset=s,n.index=n.datasetIndex=this.index;return n.active=!!e,n.mode=i,n}resolveDatasetElementOptions(t){return this._resolveElementOptions(this.datasetElementType.id,t)}resolveDataElementOptions(t,e){return this._resolveElementOptions(this.dataElementType.id,e,t)}_resolveElementOptions(t,e="default",i){const s="active"===e,n=this._cachedDataOpts,o=t+"-"+e,a=n[o],r=this.enableOptionSharing&&k(i);if(a)return Ws(a,r);const l=this.chart.config,h=l.datasetElementScopeKeys(this._type,t),c=s?[`${t}Hover`,"hover",t,""]:[t,""],d=l.getOptionScopes(this.getDataset(),h),u=Object.keys(ue.elements[t]),f=l.resolveNamedOptions(d,u,(()=>this.getContext(i,s,e)),c);return f.$shared&&(f.$shared=r,n[o]=Object.freeze(Ws(f,r))),f}_resolveAnimations(t,e,i){const s=this.chart,n=this._cachedDataOpts,o=`animation-${e}`,a=n[o];if(a)return a;let r;if(!1!==s.options.animation){const s=this.chart.config,n=s.datasetAnimationScopeKeys(this._type,e),o=s.getOptionScopes(this.getDataset(),n);r=s.createResolver(o,this.getContext(t,i,e))}const l=new Os(s,r&&r.animations);return r&&r._cacheable&&(n[o]=Object.freeze(l)),l}getSharedOptions(t){if(t.$shared)return this._sharedOptions||(this._sharedOptions=Object.assign({},t))}includeOptions(t,e){return!e||Bs(t)||this.chart._animationsDisabled}_getSharedOptions(t,e){const i=this.resolveDataElementOptions(t,e),s=this._sharedOptions,n=this.getSharedOptions(i),o=this.includeOptions(e,n)||n!==s;return this.updateSharedOptions(n,e,i),{sharedOptions:n,includeOptions:o}}updateElement(t,e,i,s){Bs(s)?Object.assign(t,i):this._resolveAnimations(e,s).update(t,i)}updateSharedOptions(t,e,i){t&&!Bs(e)&&this._resolveAnimations(void 0,e).update(t,i)}_setStyle(t,e,i,s){t.active=s;const n=this.getStyle(e,s);this._resolveAnimations(e,i,s).update(t,{options:!s&&this.getSharedOptions(n)||n})}removeHoverStyle(t,e,i){this._setStyle(t,i,"active",!1)}setHoverStyle(t,e,i){this._setStyle(t,i,"active",!0)}_removeDatasetHoverStyle(){const t=this._cachedMeta.dataset;t&&this._setStyle(t,void 0,"active",!1)}_setDatasetHoverStyle(){const t=this._cachedMeta.dataset;t&&this._setStyle(t,void 0,"active",!0)}_resyncElements(t){const e=this._data,i=this._cachedMeta.data;for(const[t,e,i]of this._syncList)this[t](e,i);this._syncList=[];const s=i.length,n=e.length,o=Math.min(n,s);o&&this.parse(0,o),n>s?this._insertElements(s,n-s,t):n<s&&this._removeElements(n,s-n)}_insertElements(t,e,i=!0){const s=this._cachedMeta,n=s.data,o=t+e;let a;const r=t=>{for(t.length+=e,a=t.length-1;a>=o;a--)t[a]=t[a-e]};for(r(n),a=t;a<o;++a)n[a]=new this.dataElementType;this._parsing&&r(s._parsed),this.parse(t,e),i&&this.updateElements(n,t,e,"reset")}updateElements(t,e,i,s){}_removeElements(t,e){const i=this._cachedMeta;if(this._parsing){const s=i._parsed.splice(t,e);i._stacked&&Vs(i,s)}i.data.splice(t,e)}_sync(t){if(this._parsing)this._syncList.push(t);else{const[e,i,s]=t;this[e](i,s)}this.chart._dataChanges.push([this.index,...t])}_onDataPush(){const t=arguments.length;this._sync(["_insertElements",this.getDataset().data.length-t,t])}_onDataPop(){this._sync(["_removeElements",this._cachedMeta.data.length-1,1])}_onDataShift(){this._sync(["_removeElements",0,1])}_onDataSplice(t,e){e&&this._sync(["_removeElements",t,e]);const i=arguments.length-2;i&&this._sync(["_insertElements",t,i])}_onDataUnshift(){this._sync(["_insertElements",0,arguments.length])}}class Hs{static defaults={};static defaultRoutes=void 0;x;y;active=!1;options;$animations;tooltipPosition(t){const{x:e,y:i}=this.getProps(["x","y"],t);return{x:e,y:i}}hasValue(){return N(this.x)&&N(this.y)}getProps(t,e){const i=this.$animations;if(!e||!i)return this;const s={};return t.forEach((t=>{s[t]=i[t]&&i[t].active()?i[t]._to:this[t]})),s}}function js(t,e){const i=t.options.ticks,n=function(t){const e=t.options.offset,i=t._tickSize(),s=t._length/i+(e?0:1),n=t._maxLength/i;return Math.floor(Math.min(s,n))}(t),o=Math.min(i.maxTicksLimit||n,n),a=i.major.enabled?function(t){const e=[];let i,s;for(i=0,s=t.length;i<s;i++)t[i].major&&e.push(i);return e}(e):[],r=a.length,l=a[0],h=a[r-1],c=[];if(r>o)return function(t,e,i,s){let n,o=0,a=i[0];for(s=Math.ceil(s),n=0;n<t.length;n++)n===a&&(e.push(t[n]),o++,a=i[o*s])}(e,c,a,r/o),c;const d=function(t,e,i){const s=function(t){const e=t.length;let i,s;if(e<2)return!1;for(s=t[0],i=1;i<e;++i)if(t[i]-t[i-1]!==s)return!1;return s}(t),n=e.length/i;if(!s)return Math.max(n,1);const o=W(s);for(let t=0,e=o.length-1;t<e;t++){const e=o[t];if(e>n)return e}return Math.max(n,1)}(a,e,o);if(r>0){let t,i;const n=r>1?Math.round((h-l)/(r-1)):null;for($s(e,c,d,s(n)?0:l-n,l),t=0,i=r-1;t<i;t++)$s(e,c,d,a[t],a[t+1]);return $s(e,c,d,h,s(n)?e.length:h+n),c}return $s(e,c,d),c}function $s(t,e,i,s,n){const o=l(s,0),a=Math.min(l(n,t.length),t.length);let r,h,c,d=0;for(i=Math.ceil(i),n&&(r=n-s,i=r/Math.floor(r/i)),c=o;c<0;)d++,c=Math.round(o+d*i);for(h=Math.max(o,0);h<a;h++)h===c&&(e.push(t[h]),d++,c=Math.round(o+d*i))}const Ys=(t,e,i)=>"top"===e||"left"===e?t[e]+i:t[e]-i,Us=(t,e)=>Math.min(e||t,t);function Xs(t,e){const i=[],s=t.length/e,n=t.length;let o=0;for(;o<n;o+=s)i.push(t[Math.floor(o)]);return i}function qs(t,e,i){const s=t.ticks.length,n=Math.min(e,s-1),o=t._startPixel,a=t._endPixel,r=1e-6;let l,h=t.getPixelForTick(n);if(!(i&&(l=1===s?Math.max(h-o,a-h):0===e?(t.getPixelForTick(1)-h)/2:(h-t.getPixelForTick(n-1))/2,h+=n<e?l:-l,h<o-r||h>a+r)))return h}function Ks(t){return t.drawTicks?t.tickLength:0}function Gs(t,e){if(!t.display)return 0;const i=Si(t.font,e),s=ki(t.padding);return(n(t.text)?t.text.length:1)*i.lineHeight+s.height}function Zs(t,e,i){let s=ut(t);return(i&&"right"!==e||!i&&"right"===e)&&(s=(t=>"left"===t?"right":"right"===t?"left":t)(s)),s}class Js extends Hs{constructor(t){super(),this.id=t.id,this.type=t.type,this.options=void 0,this.ctx=t.ctx,this.chart=t.chart,this.top=void 0,this.bottom=void 0,this.left=void 0,this.right=void 0,this.width=void 0,this.height=void 0,this._margins={left:0,right:0,top:0,bottom:0},this.maxWidth=void 0,this.maxHeight=void 0,this.paddingTop=void 0,this.paddingBottom=void 0,this.paddingLeft=void 0,this.paddingRight=void 0,this.axis=void 0,this.labelRotation=void 0,this.min=void 0,this.max=void 0,this._range=void 0,this.ticks=[],this._gridLineItems=null,this._labelItems=null,this._labelSizes=null,this._length=0,this._maxLength=0,this._longestTextCache={},this._startPixel=void 0,this._endPixel=void 0,this._reversePixels=!1,this._userMax=void 0,this._userMin=void 0,this._suggestedMax=void 0,this._suggestedMin=void 0,this._ticksLength=0,this._borderValue=0,this._cache={},this._dataLimitsCached=!1,this.$context=void 0}init(t){this.options=t.setContext(this.getContext()),this.axis=t.axis,this._userMin=this.parse(t.min),this._userMax=this.parse(t.max),this._suggestedMin=this.parse(t.suggestedMin),this._suggestedMax=this.parse(t.suggestedMax)}parse(t,e){return t}getUserBounds(){let{_userMin:t,_userMax:e,_suggestedMin:i,_suggestedMax:s}=this;return t=r(t,Number.POSITIVE_INFINITY),e=r(e,Number.NEGATIVE_INFINITY),i=r(i,Number.POSITIVE_INFINITY),s=r(s,Number.NEGATIVE_INFINITY),{min:r(t,i),max:r(e,s),minDefined:a(t),maxDefined:a(e)}}getMinMax(t){let e,{min:i,max:s,minDefined:n,maxDefined:o}=this.getUserBounds();if(n&&o)return{min:i,max:s};const a=this.getMatchingVisibleMetas();for(let r=0,l=a.length;r<l;++r)e=a[r].controller.getMinMax(this,t),n||(i=Math.min(i,e.min)),o||(s=Math.max(s,e.max));return i=o&&i>s?s:i,s=n&&i>s?i:s,{min:r(i,r(s,i)),max:r(s,r(i,s))}}getPadding(){return{left:this.paddingLeft||0,top:this.paddingTop||0,right:this.paddingRight||0,bottom:this.paddingBottom||0}}getTicks(){return this.ticks}getLabels(){const t=this.chart.data;return this.options.labels||(this.isHorizontal()?t.xLabels:t.yLabels)||t.labels||[]}getLabelItems(t=this.chart.chartArea){return this._labelItems||(this._labelItems=this._computeLabelItems(t))}beforeLayout(){this._cache={},this._dataLimitsCached=!1}beforeUpdate(){d(this.options.beforeUpdate,[this])}update(t,e,i){const{beginAtZero:s,grace:n,ticks:o}=this.options,a=o.sampleSize;this.beforeUpdate(),this.maxWidth=t,this.maxHeight=e,this._margins=i=Object.assign({left:0,right:0,top:0,bottom:0},i),this.ticks=null,this._labelSizes=null,this._gridLineItems=null,this._labelItems=null,this.beforeSetDimensions(),this.setDimensions(),this.afterSetDimensions(),this._maxLength=this.isHorizontal()?this.width+i.left+i.right:this.height+i.top+i.bottom,this._dataLimitsCached||(this.beforeDataLimits(),this.determineDataLimits(),this.afterDataLimits(),this._range=Di(this,n,s),this._dataLimitsCached=!0),this.beforeBuildTicks(),this.ticks=this.buildTicks()||[],this.afterBuildTicks();const r=a<this.ticks.length;this._convertTicksToLabels(r?Xs(this.ticks,a):this.ticks),this.configure(),this.beforeCalculateLabelRotation(),this.calculateLabelRotation(),this.afterCalculateLabelRotation(),o.display&&(o.autoSkip||"auto"===o.source)&&(this.ticks=js(this,this.ticks),this._labelSizes=null,this.afterAutoSkip()),r&&this._convertTicksToLabels(this.ticks),this.beforeFit(),this.fit(),this.afterFit(),this.afterUpdate()}configure(){let t,e,i=this.options.reverse;this.isHorizontal()?(t=this.left,e=this.right):(t=this.top,e=this.bottom,i=!i),this._startPixel=t,this._endPixel=e,this._reversePixels=i,this._length=e-t,this._alignToPixels=this.options.alignToPixels}afterUpdate(){d(this.options.afterUpdate,[this])}beforeSetDimensions(){d(this.options.beforeSetDimensions,[this])}setDimensions(){this.isHorizontal()?(this.width=this.maxWidth,this.left=0,this.right=this.width):(this.height=this.maxHeight,this.top=0,this.bottom=this.height),this.paddingLeft=0,this.paddingTop=0,this.paddingRight=0,this.paddingBottom=0}afterSetDimensions(){d(this.options.afterSetDimensions,[this])}_callHooks(t){this.chart.notifyPlugins(t,this.getContext()),d(this.options[t],[this])}beforeDataLimits(){this._callHooks("beforeDataLimits")}determineDataLimits(){}afterDataLimits(){this._callHooks("afterDataLimits")}beforeBuildTicks(){this._callHooks("beforeBuildTicks")}buildTicks(){return[]}afterBuildTicks(){this._callHooks("afterBuildTicks")}beforeTickToLabelConversion(){d(this.options.beforeTickToLabelConversion,[this])}generateTickLabels(t){const e=this.options.ticks;let i,s,n;for(i=0,s=t.length;i<s;i++)n=t[i],n.label=d(e.callback,[n.value,i,t],this)}afterTickToLabelConversion(){d(this.options.afterTickToLabelConversion,[this])}beforeCalculateLabelRotation(){d(this.options.beforeCalculateLabelRotation,[this])}calculateLabelRotation(){const t=this.options,e=t.ticks,i=Us(this.ticks.length,t.ticks.maxTicksLimit),s=e.minRotation||0,n=e.maxRotation;let o,a,r,l=s;if(!this._isVisible()||!e.display||s>=n||i<=1||!this.isHorizontal())return void(this.labelRotation=s);const h=this._getLabelSizes(),c=h.widest.width,d=h.highest.height,u=J(this.chart.width-c,0,this.maxWidth);o=t.offset?this.maxWidth/i:u/(i-1),c+6>o&&(o=u/(i-(t.offset?.5:1)),a=this.maxHeight-Ks(t.grid)-e.padding-Gs(t.title,this.chart.options.font),r=Math.sqrt(c*c+d*d),l=Y(Math.min(Math.asin(J((h.highest.height+6)/o,-1,1)),Math.asin(J(a/r,-1,1))-Math.asin(J(d/r,-1,1)))),l=Math.max(s,Math.min(n,l))),this.labelRotation=l}afterCalculateLabelRotation(){d(this.options.afterCalculateLabelRotation,[this])}afterAutoSkip(){}beforeFit(){d(this.options.beforeFit,[this])}fit(){const t={width:0,height:0},{chart:e,options:{ticks:i,title:s,grid:n}}=this,o=this._isVisible(),a=this.isHorizontal();if(o){const o=Gs(s,e.options.font);if(a?(t.width=this.maxWidth,t.height=Ks(n)+o):(t.height=this.maxHeight,t.width=Ks(n)+o),i.display&&this.ticks.length){const{first:e,last:s,widest:n,highest:o}=this._getLabelSizes(),r=2*i.padding,l=$(this.labelRotation),h=Math.cos(l),c=Math.sin(l);if(a){const e=i.mirror?0:c*n.width+h*o.height;t.height=Math.min(this.maxHeight,t.height+e+r)}else{const e=i.mirror?0:h*n.width+c*o.height;t.width=Math.min(this.maxWidth,t.width+e+r)}this._calculatePadding(e,s,c,h)}}this._handleMargins(),a?(this.width=this._length=e.width-this._margins.left-this._margins.right,this.height=t.height):(this.width=t.width,this.height=this._length=e.height-this._margins.top-this._margins.bottom)}_calculatePadding(t,e,i,s){const{ticks:{align:n,padding:o},position:a}=this.options,r=0!==this.labelRotation,l="top"!==a&&"x"===this.axis;if(this.isHorizontal()){const a=this.getPixelForTick(0)-this.left,h=this.right-this.getPixelForTick(this.ticks.length-1);let c=0,d=0;r?l?(c=s*t.width,d=i*e.height):(c=i*t.height,d=s*e.width):"start"===n?d=e.width:"end"===n?c=t.width:"inner"!==n&&(c=t.width/2,d=e.width/2),this.paddingLeft=Math.max((c-a+o)*this.width/(this.width-a),0),this.paddingRight=Math.max((d-h+o)*this.width/(this.width-h),0)}else{let i=e.height/2,s=t.height/2;"start"===n?(i=0,s=t.height):"end"===n&&(i=e.height,s=0),this.paddingTop=i+o,this.paddingBottom=s+o}}_handleMargins(){this._margins&&(this._margins.left=Math.max(this.paddingLeft,this._margins.left),this._margins.top=Math.max(this.paddingTop,this._margins.top),this._margins.right=Math.max(this.paddingRight,this._margins.right),this._margins.bottom=Math.max(this.paddingBottom,this._margins.bottom))}afterFit(){d(this.options.afterFit,[this])}isHorizontal(){const{axis:t,position:e}=this.options;return"top"===e||"bottom"===e||"x"===t}isFullSize(){return this.options.fullSize}_convertTicksToLabels(t){let e,i;for(this.beforeTickToLabelConversion(),this.generateTickLabels(t),e=0,i=t.length;e<i;e++)s(t[e].label)&&(t.splice(e,1),i--,e--);this.afterTickToLabelConversion()}_getLabelSizes(){let t=this._labelSizes;if(!t){const e=this.options.ticks.sampleSize;let i=this.ticks;e<i.length&&(i=Xs(i,e)),this._labelSizes=t=this._computeLabelSizes(i,i.length,this.options.ticks.maxTicksLimit)}return t}_computeLabelSizes(t,e,i){const{ctx:o,_longestTextCache:a}=this,r=[],l=[],h=Math.floor(e/Us(e,i));let c,d,f,g,p,m,x,b,_,y,v,M=0,w=0;for(c=0;c<e;c+=h){if(g=t[c].label,p=this._resolveTickFontOptions(c),o.font=m=p.string,x=a[m]=a[m]||{data:{},gc:[]},b=p.lineHeight,_=y=0,s(g)||n(g)){if(n(g))for(d=0,f=g.length;d<f;++d)v=g[d],s(v)||n(v)||(_=Ce(o,x.data,x.gc,_,v),y+=b)}else _=Ce(o,x.data,x.gc,_,g),y=b;r.push(_),l.push(y),M=Math.max(_,M),w=Math.max(y,w)}!function(t,e){u(t,(t=>{const i=t.gc,s=i.length/2;let n;if(s>e){for(n=0;n<s;++n)delete t.data[i[n]];i.splice(0,s)}}))}(a,e);const k=r.indexOf(M),S=l.indexOf(w),P=t=>({width:r[t]||0,height:l[t]||0});return{first:P(0),last:P(e-1),widest:P(k),highest:P(S),widths:r,heights:l}}getLabelForValue(t){return t}getPixelForValue(t,e){return NaN}getValueForPixel(t){}getPixelForTick(t){const e=this.ticks;return t<0||t>e.length-1?null:this.getPixelForValue(e[t].value)}getPixelForDecimal(t){this._reversePixels&&(t=1-t);const e=this._startPixel+t*this._length;return Q(this._alignToPixels?Ae(this.chart,e,0):e)}getDecimalForPixel(t){const e=(t-this._startPixel)/this._length;return this._reversePixels?1-e:e}getBasePixel(){return this.getPixelForValue(this.getBaseValue())}getBaseValue(){const{min:t,max:e}=this;return t<0&&e<0?e:t>0&&e>0?t:0}getContext(t){const e=this.ticks||[];if(t>=0&&t<e.length){const i=e[t];return i.$context||(i.$context=function(t,e,i){return Ci(t,{tick:i,index:e,type:"tick"})}(this.getContext(),t,i))}return this.$context||(this.$context=Ci(this.chart.getContext(),{scale:this,type:"scale"}))}_tickSize(){const t=this.options.ticks,e=$(this.labelRotation),i=Math.abs(Math.cos(e)),s=Math.abs(Math.sin(e)),n=this._getLabelSizes(),o=t.autoSkipPadding||0,a=n?n.widest.width+o:0,r=n?n.highest.height+o:0;return this.isHorizontal()?r*i>a*s?a/i:r/s:r*s<a*i?r/i:a/s}_isVisible(){const t=this.options.display;return"auto"!==t?!!t:this.getMatchingVisibleMetas().length>0}_computeGridLineItems(t){const e=this.axis,i=this.chart,s=this.options,{grid:n,position:a,border:r}=s,h=n.offset,c=this.isHorizontal(),d=this.ticks.length+(h?1:0),u=Ks(n),f=[],g=r.setContext(this.getContext()),p=g.display?g.width:0,m=p/2,x=function(t){return Ae(i,t,p)};let b,_,y,v,M,w,k,S,P,D,C,O;if("top"===a)b=x(this.bottom),w=this.bottom-u,S=b-m,D=x(t.top)+m,O=t.bottom;else if("bottom"===a)b=x(this.top),D=t.top,O=x(t.bottom)-m,w=b+m,S=this.top+u;else if("left"===a)b=x(this.right),M=this.right-u,k=b-m,P=x(t.left)+m,C=t.right;else if("right"===a)b=x(this.left),P=t.left,C=x(t.right)-m,M=b+m,k=this.left+u;else if("x"===e){if("center"===a)b=x((t.top+t.bottom)/2+.5);else if(o(a)){const t=Object.keys(a)[0],e=a[t];b=x(this.chart.scales[t].getPixelForValue(e))}D=t.top,O=t.bottom,w=b+m,S=w+u}else if("y"===e){if("center"===a)b=x((t.left+t.right)/2);else if(o(a)){const t=Object.keys(a)[0],e=a[t];b=x(this.chart.scales[t].getPixelForValue(e))}M=b-m,k=M-u,P=t.left,C=t.right}const A=l(s.ticks.maxTicksLimit,d),T=Math.max(1,Math.ceil(d/A));for(_=0;_<d;_+=T){const t=this.getContext(_),e=n.setContext(t),s=r.setContext(t),o=e.lineWidth,a=e.color,l=s.dash||[],d=s.dashOffset,u=e.tickWidth,g=e.tickColor,p=e.tickBorderDash||[],m=e.tickBorderDashOffset;y=qs(this,_,h),void 0!==y&&(v=Ae(i,y,o),c?M=k=P=C=v:w=S=D=O=v,f.push({tx1:M,ty1:w,tx2:k,ty2:S,x1:P,y1:D,x2:C,y2:O,width:o,color:a,borderDash:l,borderDashOffset:d,tickWidth:u,tickColor:g,tickBorderDash:p,tickBorderDashOffset:m}))}return this._ticksLength=d,this._borderValue=b,f}_computeLabelItems(t){const e=this.axis,i=this.options,{position:s,ticks:a}=i,r=this.isHorizontal(),l=this.ticks,{align:h,crossAlign:c,padding:d,mirror:u}=a,f=Ks(i.grid),g=f+d,p=u?-d:g,m=-$(this.labelRotation),x=[];let b,_,y,v,M,w,k,S,P,D,C,O,A="middle";if("top"===s)w=this.bottom-p,k=this._getXAxisLabelAlignment();else if("bottom"===s)w=this.top+p,k=this._getXAxisLabelAlignment();else if("left"===s){const t=this._getYAxisLabelAlignment(f);k=t.textAlign,M=t.x}else if("right"===s){const t=this._getYAxisLabelAlignment(f);k=t.textAlign,M=t.x}else if("x"===e){if("center"===s)w=(t.top+t.bottom)/2+g;else if(o(s)){const t=Object.keys(s)[0],e=s[t];w=this.chart.scales[t].getPixelForValue(e)+g}k=this._getXAxisLabelAlignment()}else if("y"===e){if("center"===s)M=(t.left+t.right)/2-g;else if(o(s)){const t=Object.keys(s)[0],e=s[t];M=this.chart.scales[t].getPixelForValue(e)}k=this._getYAxisLabelAlignment(f).textAlign}"y"===e&&("start"===h?A="top":"end"===h&&(A="bottom"));const T=this._getLabelSizes();for(b=0,_=l.length;b<_;++b){y=l[b],v=y.label;const t=a.setContext(this.getContext(b));S=this.getPixelForTick(b)+a.labelOffset,P=this._resolveTickFontOptions(b),D=P.lineHeight,C=n(v)?v.length:1;const e=C/2,i=t.color,o=t.textStrokeColor,h=t.textStrokeWidth;let d,f=k;if(r?(M=S,"inner"===k&&(f=b===_-1?this.options.reverse?"left":"right":0===b?this.options.reverse?"right":"left":"center"),O="top"===s?"near"===c||0!==m?-C*D+D/2:"center"===c?-T.highest.height/2-e*D+D:-T.highest.height+D/2:"near"===c||0!==m?D/2:"center"===c?T.highest.height/2-e*D:T.highest.height-C*D,u&&(O*=-1),0===m||t.showLabelBackdrop||(M+=D/2*Math.sin(m))):(w=S,O=(1-C)*D/2),t.showLabelBackdrop){const e=ki(t.backdropPadding),i=T.heights[b],s=T.widths[b];let n=O-e.top,o=0-e.left;switch(A){case"middle":n-=i/2;break;case"bottom":n-=i}switch(k){case"center":o-=s/2;break;case"right":o-=s;break;case"inner":b===_-1?o-=s:b>0&&(o-=s/2)}d={left:o,top:n,width:s+e.width,height:i+e.height,color:t.backdropColor}}x.push({label:v,font:P,textOffset:O,options:{rotation:m,color:i,strokeColor:o,strokeWidth:h,textAlign:f,textBaseline:A,translation:[M,w],backdrop:d}})}return x}_getXAxisLabelAlignment(){const{position:t,ticks:e}=this.options;if(-$(this.labelRotation))return"top"===t?"left":"right";let i="center";return"start"===e.align?i="left":"end"===e.align?i="right":"inner"===e.align&&(i="inner"),i}_getYAxisLabelAlignment(t){const{position:e,ticks:{crossAlign:i,mirror:s,padding:n}}=this.options,o=t+n,a=this._getLabelSizes().widest.width;let r,l;return"left"===e?s?(l=this.right+n,"near"===i?r="left":"center"===i?(r="center",l+=a/2):(r="right",l+=a)):(l=this.right-o,"near"===i?r="right":"center"===i?(r="center",l-=a/2):(r="left",l=this.left)):"right"===e?s?(l=this.left+n,"near"===i?r="right":"center"===i?(r="center",l-=a/2):(r="left",l-=a)):(l=this.left+o,"near"===i?r="left":"center"===i?(r="center",l+=a/2):(r="right",l=this.right)):r="right",{textAlign:r,x:l}}_computeLabelArea(){if(this.options.ticks.mirror)return;const t=this.chart,e=this.options.position;return"left"===e||"right"===e?{top:0,left:this.left,bottom:t.height,right:this.right}:"top"===e||"bottom"===e?{top:this.top,left:0,bottom:this.bottom,right:t.width}:void 0}drawBackground(){const{ctx:t,options:{backgroundColor:e},left:i,top:s,width:n,height:o}=this;e&&(t.save(),t.fillStyle=e,t.fillRect(i,s,n,o),t.restore())}getLineWidthForValue(t){const e=this.options.grid;if(!this._isVisible()||!e.display)return 0;const i=this.ticks.findIndex((e=>e.value===t));if(i>=0){return e.setContext(this.getContext(i)).lineWidth}return 0}drawGrid(t){const e=this.options.grid,i=this.ctx,s=this._gridLineItems||(this._gridLineItems=this._computeGridLineItems(t));let n,o;const a=(t,e,s)=>{s.width&&s.color&&(i.save(),i.lineWidth=s.width,i.strokeStyle=s.color,i.setLineDash(s.borderDash||[]),i.lineDashOffset=s.borderDashOffset,i.beginPath(),i.moveTo(t.x,t.y),i.lineTo(e.x,e.y),i.stroke(),i.restore())};if(e.display)for(n=0,o=s.length;n<o;++n){const t=s[n];e.drawOnChartArea&&a({x:t.x1,y:t.y1},{x:t.x2,y:t.y2},t),e.drawTicks&&a({x:t.tx1,y:t.ty1},{x:t.tx2,y:t.ty2},{color:t.tickColor,width:t.tickWidth,borderDash:t.tickBorderDash,borderDashOffset:t.tickBorderDashOffset})}}drawBorder(){const{chart:t,ctx:e,options:{border:i,grid:s}}=this,n=i.setContext(this.getContext()),o=i.display?n.width:0;if(!o)return;const a=s.setContext(this.getContext(0)).lineWidth,r=this._borderValue;let l,h,c,d;this.isHorizontal()?(l=Ae(t,this.left,o)-o/2,h=Ae(t,this.right,a)+a/2,c=d=r):(c=Ae(t,this.top,o)-o/2,d=Ae(t,this.bottom,a)+a/2,l=h=r),e.save(),e.lineWidth=n.width,e.strokeStyle=n.color,e.beginPath(),e.moveTo(l,c),e.lineTo(h,d),e.stroke(),e.restore()}drawLabels(t){if(!this.options.ticks.display)return;const e=this.ctx,i=this._computeLabelArea();i&&Ie(e,i);const s=this.getLabelItems(t);for(const t of s){const i=t.options,s=t.font;Ne(e,t.label,0,t.textOffset,s,i)}i&&ze(e)}drawTitle(){const{ctx:t,options:{position:e,title:i,reverse:s}}=this;if(!i.display)return;const a=Si(i.font),r=ki(i.padding),l=i.align;let h=a.lineHeight/2;"bottom"===e||"center"===e||o(e)?(h+=r.bottom,n(i.text)&&(h+=a.lineHeight*(i.text.length-1))):h+=r.top;const{titleX:c,titleY:d,maxWidth:u,rotation:f}=function(t,e,i,s){const{top:n,left:a,bottom:r,right:l,chart:h}=t,{chartArea:c,scales:d}=h;let u,f,g,p=0;const m=r-n,x=l-a;if(t.isHorizontal()){if(f=ft(s,a,l),o(i)){const t=Object.keys(i)[0],s=i[t];g=d[t].getPixelForValue(s)+m-e}else g="center"===i?(c.bottom+c.top)/2+m-e:Ys(t,i,e);u=l-a}else{if(o(i)){const t=Object.keys(i)[0],s=i[t];f=d[t].getPixelForValue(s)-x+e}else f="center"===i?(c.left+c.right)/2-x+e:Ys(t,i,e);g=ft(s,r,n),p="left"===i?-E:E}return{titleX:f,titleY:g,maxWidth:u,rotation:p}}(this,h,e,l);Ne(t,i.text,0,0,a,{color:i.color,maxWidth:u,rotation:f,textAlign:Zs(l,e,s),textBaseline:"middle",translation:[c,d]})}draw(t){this._isVisible()&&(this.drawBackground(),this.drawGrid(t),this.drawBorder(),this.drawTitle(),this.drawLabels(t))}_layers(){const t=this.options,e=t.ticks&&t.ticks.z||0,i=l(t.grid&&t.grid.z,-1),s=l(t.border&&t.border.z,0);return this._isVisible()&&this.draw===Js.prototype.draw?[{z:i,draw:t=>{this.drawBackground(),this.drawGrid(t),this.drawTitle()}},{z:s,draw:()=>{this.drawBorder()}},{z:e,draw:t=>{this.drawLabels(t)}}]:[{z:e,draw:t=>{this.draw(t)}}]}getMatchingVisibleMetas(t){const e=this.chart.getSortedVisibleDatasetMetas(),i=this.axis+"AxisID",s=[];let n,o;for(n=0,o=e.length;n<o;++n){const o=e[n];o[i]!==this.id||t&&o.type!==t||s.push(o)}return s}_resolveTickFontOptions(t){return Si(this.options.ticks.setContext(this.getContext(t)).font)}_maxDigits(){const t=this._resolveTickFontOptions(0).lineHeight;return(this.isHorizontal()?this.width:this.height)/t}}class Qs{constructor(t,e,i){this.type=t,this.scope=e,this.override=i,this.items=Object.create(null)}isForType(t){return Object.prototype.isPrototypeOf.call(this.type.prototype,t.prototype)}register(t){const e=Object.getPrototypeOf(t);let i;(function(t){return"id"in t&&"defaults"in t})(e)&&(i=this.register(e));const s=this.items,n=t.id,o=this.scope+"."+n;if(!n)throw new Error("class does not have id: "+t);return n in s||(s[n]=t,function(t,e,i){const s=x(Object.create(null),[i?ue.get(i):{},ue.get(e),t.defaults]);ue.set(e,s),t.defaultRoutes&&function(t,e){Object.keys(e).forEach((i=>{const s=i.split("."),n=s.pop(),o=[t].concat(s).join("."),a=e[i].split("."),r=a.pop(),l=a.join(".");ue.route(o,n,l,r)}))}(e,t.defaultRoutes);t.descriptors&&ue.describe(e,t.descriptors)}(t,o,i),this.override&&ue.override(t.id,t.overrides)),o}get(t){return this.items[t]}unregister(t){const e=this.items,i=t.id,s=this.scope;i in e&&delete e[i],s&&i in ue[s]&&(delete ue[s][i],this.override&&delete re[i])}}class tn{constructor(){this.controllers=new Qs(Ns,"datasets",!0),this.elements=new Qs(Hs,"elements"),this.plugins=new Qs(Object,"plugins"),this.scales=new Qs(Js,"scales"),this._typedRegistries=[this.controllers,this.scales,this.elements]}add(...t){this._each("register",t)}remove(...t){this._each("unregister",t)}addControllers(...t){this._each("register",t,this.controllers)}addElements(...t){this._each("register",t,this.elements)}addPlugins(...t){this._each("register",t,this.plugins)}addScales(...t){this._each("register",t,this.scales)}getController(t){return this._get(t,this.controllers,"controller")}getElement(t){return this._get(t,this.elements,"element")}getPlugin(t){return this._get(t,this.plugins,"plugin")}getScale(t){return this._get(t,this.scales,"scale")}removeControllers(...t){this._each("unregister",t,this.controllers)}removeElements(...t){this._each("unregister",t,this.elements)}removePlugins(...t){this._each("unregister",t,this.plugins)}removeScales(...t){this._each("unregister",t,this.scales)}_each(t,e,i){[...e].forEach((e=>{const s=i||this._getRegistryForType(e);i||s.isForType(e)||s===this.plugins&&e.id?this._exec(t,s,e):u(e,(e=>{const s=i||this._getRegistryForType(e);this._exec(t,s,e)}))}))}_exec(t,e,i){const s=w(t);d(i["before"+s],[],i),e[t](i),d(i["after"+s],[],i)}_getRegistryForType(t){for(let e=0;e<this._typedRegistries.length;e++){const i=this._typedRegistries[e];if(i.isForType(t))return i}return this.plugins}_get(t,e,i){const s=e.get(t);if(void 0===s)throw new Error('"'+t+'" is not a registered '+i+".");return s}}var en=new tn;class sn{constructor(){this._init=[]}notify(t,e,i,s){"beforeInit"===e&&(this._init=this._createDescriptors(t,!0),this._notify(this._init,t,"install"));const n=s?this._descriptors(t).filter(s):this._descriptors(t),o=this._notify(n,t,e,i);return"afterDestroy"===e&&(this._notify(n,t,"stop"),this._notify(this._init,t,"uninstall")),o}_notify(t,e,i,s){s=s||{};for(const n of t){const t=n.plugin;if(!1===d(t[i],[e,s,n.options],t)&&s.cancelable)return!1}return!0}invalidate(){s(this._cache)||(this._oldCache=this._cache,this._cache=void 0)}_descriptors(t){if(this._cache)return this._cache;const e=this._cache=this._createDescriptors(t);return this._notifyStateChanges(t),e}_createDescriptors(t,e){const i=t&&t.config,s=l(i.options&&i.options.plugins,{}),n=function(t){const e={},i=[],s=Object.keys(en.plugins.items);for(let t=0;t<s.length;t++)i.push(en.getPlugin(s[t]));const n=t.plugins||[];for(let t=0;t<n.length;t++){const s=n[t];-1===i.indexOf(s)&&(i.push(s),e[s.id]=!0)}return{plugins:i,localIds:e}}(i);return!1!==s||e?function(t,{plugins:e,localIds:i},s,n){const o=[],a=t.getContext();for(const r of e){const e=r.id,l=nn(s[e],n);null!==l&&o.push({plugin:r,options:on(t.config,{plugin:r,local:i[e]},l,a)})}return o}(t,n,s,e):[]}_notifyStateChanges(t){const e=this._oldCache||[],i=this._cache,s=(t,e)=>t.filter((t=>!e.some((e=>t.plugin.id===e.plugin.id))));this._notify(s(e,i),t,"stop"),this._notify(s(i,e),t,"start")}}function nn(t,e){return e||!1!==t?!0===t?{}:t:null}function on(t,{plugin:e,local:i},s,n){const o=t.pluginScopeKeys(e),a=t.getOptionScopes(s,o);return i&&e.defaults&&a.push(e.defaults),t.createResolver(a,n,[""],{scriptable:!1,indexable:!1,allKeys:!0})}function an(t,e){const i=ue.datasets[t]||{};return((e.datasets||{})[t]||{}).indexAxis||e.indexAxis||i.indexAxis||"x"}function rn(t){if("x"===t||"y"===t||"r"===t)return t}function ln(t,...e){if(rn(t))return t;for(const s of e){const e=s.axis||("top"===(i=s.position)||"bottom"===i?"x":"left"===i||"right"===i?"y":void 0)||t.length>1&&rn(t[0].toLowerCase());if(e)return e}var i;throw new Error(`Cannot determine type of '${t}' axis. Please provide 'axis' or 'position' option.`)}function hn(t,e,i){if(i[e+"AxisID"]===t)return{axis:e}}function cn(t,e){const i=re[t.type]||{scales:{}},s=e.scales||{},n=an(t.type,e),a=Object.create(null);return Object.keys(s).forEach((e=>{const r=s[e];if(!o(r))return console.error(`Invalid scale configuration for scale: ${e}`);if(r._proxy)return console.warn(`Ignoring resolver passed as options for scale: ${e}`);const l=ln(e,r,function(t,e){if(e.data&&e.data.datasets){const i=e.data.datasets.filter((e=>e.xAxisID===t||e.yAxisID===t));if(i.length)return hn(t,"x",i[0])||hn(t,"y",i[0])}return{}}(e,t),ue.scales[r.type]),h=function(t,e){return t===e?"_index_":"_value_"}(l,n),c=i.scales||{};a[e]=b(Object.create(null),[{axis:l},r,c[l],c[h]])})),t.data.datasets.forEach((i=>{const n=i.type||t.type,o=i.indexAxis||an(n,e),r=(re[n]||{}).scales||{};Object.keys(r).forEach((t=>{const e=function(t,e){let i=t;return"_index_"===t?i=e:"_value_"===t&&(i="x"===e?"y":"x"),i}(t,o),n=i[e+"AxisID"]||e;a[n]=a[n]||Object.create(null),b(a[n],[{axis:e},s[n],r[t]])}))})),Object.keys(a).forEach((t=>{const e=a[t];b(e,[ue.scales[e.type],ue.scale])})),a}function dn(t){const e=t.options||(t.options={});e.plugins=l(e.plugins,{}),e.scales=cn(t,e)}function un(t){return(t=t||{}).datasets=t.datasets||[],t.labels=t.labels||[],t}const fn=new Map,gn=new Set;function pn(t,e){let i=fn.get(t);return i||(i=e(),fn.set(t,i),gn.add(i)),i}const mn=(t,e,i)=>{const s=M(e,i);void 0!==s&&t.add(s)};class xn{constructor(t){this._config=function(t){return(t=t||{}).data=un(t.data),dn(t),t}(t),this._scopeCache=new Map,this._resolverCache=new Map}get platform(){return this._config.platform}get type(){return this._config.type}set type(t){this._config.type=t}get data(){return this._config.data}set data(t){this._config.data=un(t)}get options(){return this._config.options}set options(t){this._config.options=t}get plugins(){return this._config.plugins}update(){const t=this._config;this.clearCache(),dn(t)}clearCache(){this._scopeCache.clear(),this._resolverCache.clear()}datasetScopeKeys(t){return pn(t,(()=>[[`datasets.${t}`,""]]))}datasetAnimationScopeKeys(t,e){return pn(`${t}.transition.${e}`,(()=>[[`datasets.${t}.transitions.${e}`,`transitions.${e}`],[`datasets.${t}`,""]]))}datasetElementScopeKeys(t,e){return pn(`${t}-${e}`,(()=>[[`datasets.${t}.elements.${e}`,`datasets.${t}`,`elements.${e}`,""]]))}pluginScopeKeys(t){const e=t.id;return pn(`${this.type}-plugin-${e}`,(()=>[[`plugins.${e}`,...t.additionalOptionScopes||[]]]))}_cachedScopes(t,e){const i=this._scopeCache;let s=i.get(t);return s&&!e||(s=new Map,i.set(t,s)),s}getOptionScopes(t,e,i){const{options:s,type:n}=this,o=this._cachedScopes(t,i),a=o.get(e);if(a)return a;const r=new Set;e.forEach((e=>{t&&(r.add(t),e.forEach((e=>mn(r,t,e)))),e.forEach((t=>mn(r,s,t))),e.forEach((t=>mn(r,re[n]||{},t))),e.forEach((t=>mn(r,ue,t))),e.forEach((t=>mn(r,le,t)))}));const l=Array.from(r);return 0===l.length&&l.push(Object.create(null)),gn.has(e)&&o.set(e,l),l}chartOptionScopes(){const{options:t,type:e}=this;return[t,re[e]||{},ue.datasets[e]||{},{type:e},ue,le]}resolveNamedOptions(t,e,i,s=[""]){const o={$shared:!0},{resolver:a,subPrefixes:r}=bn(this._resolverCache,t,s);let l=a;if(function(t,e){const{isScriptable:i,isIndexable:s}=Ye(t);for(const o of e){const e=i(o),a=s(o),r=(a||e)&&t[o];if(e&&(S(r)||_n(r))||a&&n(r))return!0}return!1}(a,e)){o.$shared=!1;l=$e(a,i=S(i)?i():i,this.createResolver(t,i,r))}for(const t of e)o[t]=l[t];return o}createResolver(t,e,i=[""],s){const{resolver:n}=bn(this._resolverCache,t,i);return o(e)?$e(n,e,void 0,s):n}}function bn(t,e,i){let s=t.get(e);s||(s=new Map,t.set(e,s));const n=i.join();let o=s.get(n);if(!o){o={resolver:je(e,i),subPrefixes:i.filter((t=>!t.toLowerCase().includes("hover")))},s.set(n,o)}return o}const _n=t=>o(t)&&Object.getOwnPropertyNames(t).some((e=>S(t[e])));const yn=["top","bottom","left","right","chartArea"];function vn(t,e){return"top"===t||"bottom"===t||-1===yn.indexOf(t)&&"x"===e}function Mn(t,e){return function(i,s){return i[t]===s[t]?i[e]-s[e]:i[t]-s[t]}}function wn(t){const e=t.chart,i=e.options.animation;e.notifyPlugins("afterRender"),d(i&&i.onComplete,[t],e)}function kn(t){const e=t.chart,i=e.options.animation;d(i&&i.onProgress,[t],e)}function Sn(t){return fe()&&"string"==typeof t?t=document.getElementById(t):t&&t.length&&(t=t[0]),t&&t.canvas&&(t=t.canvas),t}const Pn={},Dn=t=>{const e=Sn(t);return Object.values(Pn).filter((t=>t.canvas===e)).pop()};function Cn(t,e,i){const s=Object.keys(t);for(const n of s){const s=+n;if(s>=e){const o=t[n];delete t[n],(i>0||s>e)&&(t[s+i]=o)}}}function On(t,e,i){return t.options.clip?t[i]:e[i]}class An{static defaults=ue;static instances=Pn;static overrides=re;static registry=en;static version="4.4.3";static getChart=Dn;static register(...t){en.add(...t),Tn()}static unregister(...t){en.remove(...t),Tn()}constructor(t,e){const s=this.config=new xn(e),n=Sn(t),o=Dn(n);if(o)throw new Error("Canvas is already in use. Chart with ID '"+o.id+"' must be destroyed before the canvas with ID '"+o.canvas.id+"' can be reused.");const a=s.createResolver(s.chartOptionScopes(),this.getContext());this.platform=new(s.platform||ks(n)),this.platform.updateConfig(s);const r=this.platform.acquireContext(n,a.aspectRatio),l=r&&r.canvas,h=l&&l.height,c=l&&l.width;this.id=i(),this.ctx=r,this.canvas=l,this.width=c,this.height=h,this._options=a,this._aspectRatio=this.aspectRatio,this._layers=[],this._metasets=[],this._stacks=void 0,this.boxes=[],this.currentDevicePixelRatio=void 0,this.chartArea=void 0,this._active=[],this._lastEvent=void 0,this._listeners={},this._responsiveListeners=void 0,this._sortedMetasets=[],this.scales={},this._plugins=new sn,this.$proxies={},this._hiddenIndices={},this.attached=!1,this._animationsDisabled=void 0,this.$context=void 0,this._doResize=dt((t=>this.update(t)),a.resizeDelay||0),this._dataChanges=[],Pn[this.id]=this,r&&l?(bt.listen(this,"complete",wn),bt.listen(this,"progress",kn),this._initialize(),this.attached&&this.update()):console.error("Failed to create chart: can't acquire context from the given item")}get aspectRatio(){const{options:{aspectRatio:t,maintainAspectRatio:e},width:i,height:n,_aspectRatio:o}=this;return s(t)?e&&o?o:n?i/n:null:t}get data(){return this.config.data}set data(t){this.config.data=t}get options(){return this._options}set options(t){this.config.options=t}get registry(){return en}_initialize(){return this.notifyPlugins("beforeInit"),this.options.responsive?this.resize():ke(this,this.options.devicePixelRatio),this.bindEvents(),this.notifyPlugins("afterInit"),this}clear(){return Te(this.canvas,this.ctx),this}stop(){return bt.stop(this),this}resize(t,e){bt.running(this)?this._resizeBeforeDraw={width:t,height:e}:this._resize(t,e)}_resize(t,e){const i=this.options,s=this.canvas,n=i.maintainAspectRatio&&this.aspectRatio,o=this.platform.getMaximumSize(s,t,e,n),a=i.devicePixelRatio||this.platform.getDevicePixelRatio(),r=this.width?"resize":"attach";this.width=o.width,this.height=o.height,this._aspectRatio=this.aspectRatio,ke(this,a,!0)&&(this.notifyPlugins("resize",{size:o}),d(i.onResize,[this,o],this),this.attached&&this._doResize(r)&&this.render())}ensureScalesHaveIDs(){u(this.options.scales||{},((t,e)=>{t.id=e}))}buildOrUpdateScales(){const t=this.options,e=t.scales,i=this.scales,s=Object.keys(i).reduce(((t,e)=>(t[e]=!1,t)),{});let n=[];e&&(n=n.concat(Object.keys(e).map((t=>{const i=e[t],s=ln(t,i),n="r"===s,o="x"===s;return{options:i,dposition:n?"chartArea":o?"bottom":"left",dtype:n?"radialLinear":o?"category":"linear"}})))),u(n,(e=>{const n=e.options,o=n.id,a=ln(o,n),r=l(n.type,e.dtype);void 0!==n.position&&vn(n.position,a)===vn(e.dposition)||(n.position=e.dposition),s[o]=!0;let h=null;if(o in i&&i[o].type===r)h=i[o];else{h=new(en.getScale(r))({id:o,type:r,ctx:this.ctx,chart:this}),i[h.id]=h}h.init(n,t)})),u(s,((t,e)=>{t||delete i[e]})),u(i,(t=>{as.configure(this,t,t.options),as.addBox(this,t)}))}_updateMetasets(){const t=this._metasets,e=this.data.datasets.length,i=t.length;if(t.sort(((t,e)=>t.index-e.index)),i>e){for(let t=e;t<i;++t)this._destroyDatasetMeta(t);t.splice(e,i-e)}this._sortedMetasets=t.slice(0).sort(Mn("order","index"))}_removeUnreferencedMetasets(){const{_metasets:t,data:{datasets:e}}=this;t.length>e.length&&delete this._stacks,t.forEach(((t,i)=>{0===e.filter((e=>e===t._dataset)).length&&this._destroyDatasetMeta(i)}))}buildOrUpdateControllers(){const t=[],e=this.data.datasets;let i,s;for(this._removeUnreferencedMetasets(),i=0,s=e.length;i<s;i++){const s=e[i];let n=this.getDatasetMeta(i);const o=s.type||this.config.type;if(n.type&&n.type!==o&&(this._destroyDatasetMeta(i),n=this.getDatasetMeta(i)),n.type=o,n.indexAxis=s.indexAxis||an(o,this.options),n.order=s.order||0,n.index=i,n.label=""+s.label,n.visible=this.isDatasetVisible(i),n.controller)n.controller.updateIndex(i),n.controller.linkScales();else{const e=en.getController(o),{datasetElementType:s,dataElementType:a}=ue.datasets[o];Object.assign(e,{dataElementType:en.getElement(a),datasetElementType:s&&en.getElement(s)}),n.controller=new e(this,i),t.push(n.controller)}}return this._updateMetasets(),t}_resetElements(){u(this.data.datasets,((t,e)=>{this.getDatasetMeta(e).controller.reset()}),this)}reset(){this._resetElements(),this.notifyPlugins("reset")}update(t){const e=this.config;e.update();const i=this._options=e.createResolver(e.chartOptionScopes(),this.getContext()),s=this._animationsDisabled=!i.animation;if(this._updateScales(),this._checkEventBindings(),this._updateHiddenIndices(),this._plugins.invalidate(),!1===this.notifyPlugins("beforeUpdate",{mode:t,cancelable:!0}))return;const n=this.buildOrUpdateControllers();this.notifyPlugins("beforeElementsUpdate");let o=0;for(let t=0,e=this.data.datasets.length;t<e;t++){const{controller:e}=this.getDatasetMeta(t),i=!s&&-1===n.indexOf(e);e.buildOrUpdateElements(i),o=Math.max(+e.getMaxOverflow(),o)}o=this._minPadding=i.layout.autoPadding?o:0,this._updateLayout(o),s||u(n,(t=>{t.reset()})),this._updateDatasets(t),this.notifyPlugins("afterUpdate",{mode:t}),this._layers.sort(Mn("z","_idx"));const{_active:a,_lastEvent:r}=this;r?this._eventHandler(r,!0):a.length&&this._updateHoverStyles(a,a,!0),this.render()}_updateScales(){u(this.scales,(t=>{as.removeBox(this,t)})),this.ensureScalesHaveIDs(),this.buildOrUpdateScales()}_checkEventBindings(){const t=this.options,e=new Set(Object.keys(this._listeners)),i=new Set(t.events);P(e,i)&&!!this._responsiveListeners===t.responsive||(this.unbindEvents(),this.bindEvents())}_updateHiddenIndices(){const{_hiddenIndices:t}=this,e=this._getUniformDataChanges()||[];for(const{method:i,start:s,count:n}of e){Cn(t,s,"_removeElements"===i?-n:n)}}_getUniformDataChanges(){const t=this._dataChanges;if(!t||!t.length)return;this._dataChanges=[];const e=this.data.datasets.length,i=e=>new Set(t.filter((t=>t[0]===e)).map(((t,e)=>e+","+t.splice(1).join(",")))),s=i(0);for(let t=1;t<e;t++)if(!P(s,i(t)))return;return Array.from(s).map((t=>t.split(","))).map((t=>({method:t[1],start:+t[2],count:+t[3]})))}_updateLayout(t){if(!1===this.notifyPlugins("beforeLayout",{cancelable:!0}))return;as.update(this,this.width,this.height,t);const e=this.chartArea,i=e.width<=0||e.height<=0;this._layers=[],u(this.boxes,(t=>{i&&"chartArea"===t.position||(t.configure&&t.configure(),this._layers.push(...t._layers()))}),this),this._layers.forEach(((t,e)=>{t._idx=e})),this.notifyPlugins("afterLayout")}_updateDatasets(t){if(!1!==this.notifyPlugins("beforeDatasetsUpdate",{mode:t,cancelable:!0})){for(let t=0,e=this.data.datasets.length;t<e;++t)this.getDatasetMeta(t).controller.configure();for(let e=0,i=this.data.datasets.length;e<i;++e)this._updateDataset(e,S(t)?t({datasetIndex:e}):t);this.notifyPlugins("afterDatasetsUpdate",{mode:t})}}_updateDataset(t,e){const i=this.getDatasetMeta(t),s={meta:i,index:t,mode:e,cancelable:!0};!1!==this.notifyPlugins("beforeDatasetUpdate",s)&&(i.controller._update(e),s.cancelable=!1,this.notifyPlugins("afterDatasetUpdate",s))}render(){!1!==this.notifyPlugins("beforeRender",{cancelable:!0})&&(bt.has(this)?this.attached&&!bt.running(this)&&bt.start(this):(this.draw(),wn({chart:this})))}draw(){let t;if(this._resizeBeforeDraw){const{width:t,height:e}=this._resizeBeforeDraw;this._resize(t,e),this._resizeBeforeDraw=null}if(this.clear(),this.width<=0||this.height<=0)return;if(!1===this.notifyPlugins("beforeDraw",{cancelable:!0}))return;const e=this._layers;for(t=0;t<e.length&&e[t].z<=0;++t)e[t].draw(this.chartArea);for(this._drawDatasets();t<e.length;++t)e[t].draw(this.chartArea);this.notifyPlugins("afterDraw")}_getSortedDatasetMetas(t){const e=this._sortedMetasets,i=[];let s,n;for(s=0,n=e.length;s<n;++s){const n=e[s];t&&!n.visible||i.push(n)}return i}getSortedVisibleDatasetMetas(){return this._getSortedDatasetMetas(!0)}_drawDatasets(){if(!1===this.notifyPlugins("beforeDatasetsDraw",{cancelable:!0}))return;const t=this.getSortedVisibleDatasetMetas();for(let e=t.length-1;e>=0;--e)this._drawDataset(t[e]);this.notifyPlugins("afterDatasetsDraw")}_drawDataset(t){const e=this.ctx,i=t._clip,s=!i.disabled,n=function(t,e){const{xScale:i,yScale:s}=t;return i&&s?{left:On(i,e,"left"),right:On(i,e,"right"),top:On(s,e,"top"),bottom:On(s,e,"bottom")}:e}(t,this.chartArea),o={meta:t,index:t.index,cancelable:!0};!1!==this.notifyPlugins("beforeDatasetDraw",o)&&(s&&Ie(e,{left:!1===i.left?0:n.left-i.left,right:!1===i.right?this.width:n.right+i.right,top:!1===i.top?0:n.top-i.top,bottom:!1===i.bottom?this.height:n.bottom+i.bottom}),t.controller.draw(),s&&ze(e),o.cancelable=!1,this.notifyPlugins("afterDatasetDraw",o))}isPointInArea(t){return Re(t,this.chartArea,this._minPadding)}getElementsAtEventForMode(t,e,i,s){const n=Xi.modes[e];return"function"==typeof n?n(this,t,i,s):[]}getDatasetMeta(t){const e=this.data.datasets[t],i=this._metasets;let s=i.filter((t=>t&&t._dataset===e)).pop();return s||(s={type:null,data:[],dataset:null,controller:null,hidden:null,xAxisID:null,yAxisID:null,order:e&&e.order||0,index:t,_dataset:e,_parsed:[],_sorted:!1},i.push(s)),s}getContext(){return this.$context||(this.$context=Ci(null,{chart:this,type:"chart"}))}getVisibleDatasetCount(){return this.getSortedVisibleDatasetMetas().length}isDatasetVisible(t){const e=this.data.datasets[t];if(!e)return!1;const i=this.getDatasetMeta(t);return"boolean"==typeof i.hidden?!i.hidden:!e.hidden}setDatasetVisibility(t,e){this.getDatasetMeta(t).hidden=!e}toggleDataVisibility(t){this._hiddenIndices[t]=!this._hiddenIndices[t]}getDataVisibility(t){return!this._hiddenIndices[t]}_updateVisibility(t,e,i){const s=i?"show":"hide",n=this.getDatasetMeta(t),o=n.controller._resolveAnimations(void 0,s);k(e)?(n.data[e].hidden=!i,this.update()):(this.setDatasetVisibility(t,i),o.update(n,{visible:i}),this.update((e=>e.datasetIndex===t?s:void 0)))}hide(t,e){this._updateVisibility(t,e,!1)}show(t,e){this._updateVisibility(t,e,!0)}_destroyDatasetMeta(t){const e=this._metasets[t];e&&e.controller&&e.controller._destroy(),delete this._metasets[t]}_stop(){let t,e;for(this.stop(),bt.remove(this),t=0,e=this.data.datasets.length;t<e;++t)this._destroyDatasetMeta(t)}destroy(){this.notifyPlugins("beforeDestroy");const{canvas:t,ctx:e}=this;this._stop(),this.config.clearCache(),t&&(this.unbindEvents(),Te(t,e),this.platform.releaseContext(e),this.canvas=null,this.ctx=null),delete Pn[this.id],this.notifyPlugins("afterDestroy")}toBase64Image(...t){return this.canvas.toDataURL(...t)}bindEvents(){this.bindUserEvents(),this.options.responsive?this.bindResponsiveEvents():this.attached=!0}bindUserEvents(){const t=this._listeners,e=this.platform,i=(i,s)=>{e.addEventListener(this,i,s),t[i]=s},s=(t,e,i)=>{t.offsetX=e,t.offsetY=i,this._eventHandler(t)};u(this.options.events,(t=>i(t,s)))}bindResponsiveEvents(){this._responsiveListeners||(this._responsiveListeners={});const t=this._responsiveListeners,e=this.platform,i=(i,s)=>{e.addEventListener(this,i,s),t[i]=s},s=(i,s)=>{t[i]&&(e.removeEventListener(this,i,s),delete t[i])},n=(t,e)=>{this.canvas&&this.resize(t,e)};let o;const a=()=>{s("attach",a),this.attached=!0,this.resize(),i("resize",n),i("detach",o)};o=()=>{this.attached=!1,s("resize",n),this._stop(),this._resize(0,0),i("attach",a)},e.isAttached(this.canvas)?a():o()}unbindEvents(){u(this._listeners,((t,e)=>{this.platform.removeEventListener(this,e,t)})),this._listeners={},u(this._responsiveListeners,((t,e)=>{this.platform.removeEventListener(this,e,t)})),this._responsiveListeners=void 0}updateHoverStyle(t,e,i){const s=i?"set":"remove";let n,o,a,r;for("dataset"===e&&(n=this.getDatasetMeta(t[0].datasetIndex),n.controller["_"+s+"DatasetHoverStyle"]()),a=0,r=t.length;a<r;++a){o=t[a];const e=o&&this.getDatasetMeta(o.datasetIndex).controller;e&&e[s+"HoverStyle"](o.element,o.datasetIndex,o.index)}}getActiveElements(){return this._active||[]}setActiveElements(t){const e=this._active||[],i=t.map((({datasetIndex:t,index:e})=>{const i=this.getDatasetMeta(t);if(!i)throw new Error("No dataset found at index "+t);return{datasetIndex:t,element:i.data[e],index:e}}));!f(i,e)&&(this._active=i,this._lastEvent=null,this._updateHoverStyles(i,e))}notifyPlugins(t,e,i){return this._plugins.notify(this,t,e,i)}isPluginEnabled(t){return 1===this._plugins._cache.filter((e=>e.plugin.id===t)).length}_updateHoverStyles(t,e,i){const s=this.options.hover,n=(t,e)=>t.filter((t=>!e.some((e=>t.datasetIndex===e.datasetIndex&&t.index===e.index)))),o=n(e,t),a=i?t:n(t,e);o.length&&this.updateHoverStyle(o,s.mode,!1),a.length&&s.mode&&this.updateHoverStyle(a,s.mode,!0)}_eventHandler(t,e){const i={event:t,replay:e,cancelable:!0,inChartArea:this.isPointInArea(t)},s=e=>(e.options.events||this.options.events).includes(t.native.type);if(!1===this.notifyPlugins("beforeEvent",i,s))return;const n=this._handleEvent(t,e,i.inChartArea);return i.cancelable=!1,this.notifyPlugins("afterEvent",i,s),(n||i.changed)&&this.render(),this}_handleEvent(t,e,i){const{_active:s=[],options:n}=this,o=e,a=this._getActiveElements(t,s,i,o),r=D(t),l=function(t,e,i,s){return i&&"mouseout"!==t.type?s?e:t:null}(t,this._lastEvent,i,r);i&&(this._lastEvent=null,d(n.onHover,[t,a,this],this),r&&d(n.onClick,[t,a,this],this));const h=!f(a,s);return(h||e)&&(this._active=a,this._updateHoverStyles(a,s,e)),this._lastEvent=l,h}_getActiveElements(t,e,i,s){if("mouseout"===t.type)return[];if(!i)return e;const n=this.options.hover;return this.getElementsAtEventForMode(t,n.mode,n,s)}}function Tn(){return u(An.instances,(t=>t._plugins.invalidate()))}function Ln(){throw new Error("This method is not implemented: Check that a complete date adapter is provided.")}class En{static override(t){Object.assign(En.prototype,t)}options;constructor(t){this.options=t||{}}init(){}formats(){return Ln()}parse(){return Ln()}format(){return Ln()}add(){return Ln()}diff(){return Ln()}startOf(){return Ln()}endOf(){return Ln()}}var Rn={_date:En};function In(t){const e=t.iScale,i=function(t,e){if(!t._cache.$bar){const i=t.getMatchingVisibleMetas(e);let s=[];for(let e=0,n=i.length;e<n;e++)s=s.concat(i[e].controller.getAllParsedValues(t));t._cache.$bar=lt(s.sort(((t,e)=>t-e)))}return t._cache.$bar}(e,t.type);let s,n,o,a,r=e._length;const l=()=>{32767!==o&&-32768!==o&&(k(a)&&(r=Math.min(r,Math.abs(o-a)||r)),a=o)};for(s=0,n=i.length;s<n;++s)o=e.getPixelForValue(i[s]),l();for(a=void 0,s=0,n=e.ticks.length;s<n;++s)o=e.getPixelForTick(s),l();return r}function zn(t,e,i,s){return n(t)?function(t,e,i,s){const n=i.parse(t[0],s),o=i.parse(t[1],s),a=Math.min(n,o),r=Math.max(n,o);let l=a,h=r;Math.abs(a)>Math.abs(r)&&(l=r,h=a),e[i.axis]=h,e._custom={barStart:l,barEnd:h,start:n,end:o,min:a,max:r}}(t,e,i,s):e[i.axis]=i.parse(t,s),e}function Fn(t,e,i,s){const n=t.iScale,o=t.vScale,a=n.getLabels(),r=n===o,l=[];let h,c,d,u;for(h=i,c=i+s;h<c;++h)u=e[h],d={},d[n.axis]=r||n.parse(a[h],h),l.push(zn(u,d,o,h));return l}function Vn(t){return t&&void 0!==t.barStart&&void 0!==t.barEnd}function Bn(t,e,i,s){let n=e.borderSkipped;const o={};if(!n)return void(t.borderSkipped=o);if(!0===n)return void(t.borderSkipped={top:!0,right:!0,bottom:!0,left:!0});const{start:a,end:r,reverse:l,top:h,bottom:c}=function(t){let e,i,s,n,o;return t.horizontal?(e=t.base>t.x,i="left",s="right"):(e=t.base<t.y,i="bottom",s="top"),e?(n="end",o="start"):(n="start",o="end"),{start:i,end:s,reverse:e,top:n,bottom:o}}(t);"middle"===n&&i&&(t.enableBorderRadius=!0,(i._top||0)===s?n=h:(i._bottom||0)===s?n=c:(o[Wn(c,a,r,l)]=!0,n=h)),o[Wn(n,a,r,l)]=!0,t.borderSkipped=o}function Wn(t,e,i,s){var n,o,a;return s?(a=i,t=Nn(t=(n=t)===(o=e)?a:n===a?o:n,i,e)):t=Nn(t,e,i),t}function Nn(t,e,i){return"start"===t?e:"end"===t?i:t}function Hn(t,{inflateAmount:e},i){t.inflateAmount="auto"===e?1===i?.33:0:e}class jn extends Ns{static id="doughnut";static defaults={datasetElementType:!1,dataElementType:"arc",animation:{animateRotate:!0,animateScale:!1},animations:{numbers:{type:"number",properties:["circumference","endAngle","innerRadius","outerRadius","startAngle","x","y","offset","borderWidth","spacing"]}},cutout:"50%",rotation:0,circumference:360,radius:"100%",spacing:0,indexAxis:"r"};static descriptors={_scriptable:t=>"spacing"!==t,_indexable:t=>"spacing"!==t&&!t.startsWith("borderDash")&&!t.startsWith("hoverBorderDash")};static overrides={aspectRatio:1,plugins:{legend:{labels:{generateLabels(t){const e=t.data;if(e.labels.length&&e.datasets.length){const{labels:{pointStyle:i,color:s}}=t.legend.options;return e.labels.map(((e,n)=>{const o=t.getDatasetMeta(0).controller.getStyle(n);return{text:e,fillStyle:o.backgroundColor,strokeStyle:o.borderColor,fontColor:s,lineWidth:o.borderWidth,pointStyle:i,hidden:!t.getDataVisibility(n),index:n}}))}return[]}},onClick(t,e,i){i.chart.toggleDataVisibility(e.index),i.chart.update()}}}};constructor(t,e){super(t,e),this.enableOptionSharing=!0,this.innerRadius=void 0,this.outerRadius=void 0,this.offsetX=void 0,this.offsetY=void 0}linkScales(){}parse(t,e){const i=this.getDataset().data,s=this._cachedMeta;if(!1===this._parsing)s._parsed=i;else{let n,a,r=t=>+i[t];if(o(i[t])){const{key:t="value"}=this._parsing;r=e=>+M(i[e],t)}for(n=t,a=t+e;n<a;++n)s._parsed[n]=r(n)}}_getRotation(){return $(this.options.rotation-90)}_getCircumference(){return $(this.options.circumference)}_getRotationExtents(){let t=O,e=-O;for(let i=0;i<this.chart.data.datasets.length;++i)if(this.chart.isDatasetVisible(i)&&this.chart.getDatasetMeta(i).type===this._type){const s=this.chart.getDatasetMeta(i).controller,n=s._getRotation(),o=s._getCircumference();t=Math.min(t,n),e=Math.max(e,n+o)}return{rotation:t,circumference:e-t}}update(t){const e=this.chart,{chartArea:i}=e,s=this._cachedMeta,n=s.data,o=this.getMaxBorderWidth()+this.getMaxOffset(n)+this.options.spacing,a=Math.max((Math.min(i.width,i.height)-o)/2,0),r=Math.min(h(this.options.cutout,a),1),l=this._getRingWeight(this.index),{circumference:d,rotation:u}=this._getRotationExtents(),{ratioX:f,ratioY:g,offsetX:p,offsetY:m}=function(t,e,i){let s=1,n=1,o=0,a=0;if(e<O){const r=t,l=r+e,h=Math.cos(r),c=Math.sin(r),d=Math.cos(l),u=Math.sin(l),f=(t,e,s)=>Z(t,r,l,!0)?1:Math.max(e,e*i,s,s*i),g=(t,e,s)=>Z(t,r,l,!0)?-1:Math.min(e,e*i,s,s*i),p=f(0,h,d),m=f(E,c,u),x=g(C,h,d),b=g(C+E,c,u);s=(p-x)/2,n=(m-b)/2,o=-(p+x)/2,a=-(m+b)/2}return{ratioX:s,ratioY:n,offsetX:o,offsetY:a}}(u,d,r),x=(i.width-o)/f,b=(i.height-o)/g,_=Math.max(Math.min(x,b)/2,0),y=c(this.options.radius,_),v=(y-Math.max(y*r,0))/this._getVisibleDatasetWeightTotal();this.offsetX=p*y,this.offsetY=m*y,s.total=this.calculateTotal(),this.outerRadius=y-v*this._getRingWeightOffset(this.index),this.innerRadius=Math.max(this.outerRadius-v*l,0),this.updateElements(n,0,n.length,t)}_circumference(t,e){const i=this.options,s=this._cachedMeta,n=this._getCircumference();return e&&i.animation.animateRotate||!this.chart.getDataVisibility(t)||null===s._parsed[t]||s.data[t].hidden?0:this.calculateCircumference(s._parsed[t]*n/O)}updateElements(t,e,i,s){const n="reset"===s,o=this.chart,a=o.chartArea,r=o.options.animation,l=(a.left+a.right)/2,h=(a.top+a.bottom)/2,c=n&&r.animateScale,d=c?0:this.innerRadius,u=c?0:this.outerRadius,{sharedOptions:f,includeOptions:g}=this._getSharedOptions(e,s);let p,m=this._getRotation();for(p=0;p<e;++p)m+=this._circumference(p,n);for(p=e;p<e+i;++p){const e=this._circumference(p,n),i=t[p],o={x:l+this.offsetX,y:h+this.offsetY,startAngle:m,endAngle:m+e,circumference:e,outerRadius:u,innerRadius:d};g&&(o.options=f||this.resolveDataElementOptions(p,i.active?"active":s)),m+=e,this.updateElement(i,p,o,s)}}calculateTotal(){const t=this._cachedMeta,e=t.data;let i,s=0;for(i=0;i<e.length;i++){const n=t._parsed[i];null===n||isNaN(n)||!this.chart.getDataVisibility(i)||e[i].hidden||(s+=Math.abs(n))}return s}calculateCircumference(t){const e=this._cachedMeta.total;return e>0&&!isNaN(t)?O*(Math.abs(t)/e):0}getLabelAndValue(t){const e=this._cachedMeta,i=this.chart,s=i.data.labels||[],n=ne(e._parsed[t],i.options.locale);return{label:s[t]||"",value:n}}getMaxBorderWidth(t){let e=0;const i=this.chart;let s,n,o,a,r;if(!t)for(s=0,n=i.data.datasets.length;s<n;++s)if(i.isDatasetVisible(s)){o=i.getDatasetMeta(s),t=o.data,a=o.controller;break}if(!t)return 0;for(s=0,n=t.length;s<n;++s)r=a.resolveDataElementOptions(s),"inner"!==r.borderAlign&&(e=Math.max(e,r.borderWidth||0,r.hoverBorderWidth||0));return e}getMaxOffset(t){let e=0;for(let i=0,s=t.length;i<s;++i){const t=this.resolveDataElementOptions(i);e=Math.max(e,t.offset||0,t.hoverOffset||0)}return e}_getRingWeightOffset(t){let e=0;for(let i=0;i<t;++i)this.chart.isDatasetVisible(i)&&(e+=this._getRingWeight(i));return e}_getRingWeight(t){return Math.max(l(this.chart.data.datasets[t].weight,1),0)}_getVisibleDatasetWeightTotal(){return this._getRingWeightOffset(this.chart.data.datasets.length)||1}}class $n extends Ns{static id="polarArea";static defaults={dataElementType:"arc",animation:{animateRotate:!0,animateScale:!0},animations:{numbers:{type:"number",properties:["x","y","startAngle","endAngle","innerRadius","outerRadius"]}},indexAxis:"r",startAngle:0};static overrides={aspectRatio:1,plugins:{legend:{labels:{generateLabels(t){const e=t.data;if(e.labels.length&&e.datasets.length){const{labels:{pointStyle:i,color:s}}=t.legend.options;return e.labels.map(((e,n)=>{const o=t.getDatasetMeta(0).controller.getStyle(n);return{text:e,fillStyle:o.backgroundColor,strokeStyle:o.borderColor,fontColor:s,lineWidth:o.borderWidth,pointStyle:i,hidden:!t.getDataVisibility(n),index:n}}))}return[]}},onClick(t,e,i){i.chart.toggleDataVisibility(e.index),i.chart.update()}}},scales:{r:{type:"radialLinear",angleLines:{display:!1},beginAtZero:!0,grid:{circular:!0},pointLabels:{display:!1},startAngle:0}}};constructor(t,e){super(t,e),this.innerRadius=void 0,this.outerRadius=void 0}getLabelAndValue(t){const e=this._cachedMeta,i=this.chart,s=i.data.labels||[],n=ne(e._parsed[t].r,i.options.locale);return{label:s[t]||"",value:n}}parseObjectData(t,e,i,s){return ii.bind(this)(t,e,i,s)}update(t){const e=this._cachedMeta.data;this._updateRadius(),this.updateElements(e,0,e.length,t)}getMinMax(){const t=this._cachedMeta,e={min:Number.POSITIVE_INFINITY,max:Number.NEGATIVE_INFINITY};return t.data.forEach(((t,i)=>{const s=this.getParsed(i).r;!isNaN(s)&&this.chart.getDataVisibility(i)&&(s<e.min&&(e.min=s),s>e.max&&(e.max=s))})),e}_updateRadius(){const t=this.chart,e=t.chartArea,i=t.options,s=Math.min(e.right-e.left,e.bottom-e.top),n=Math.max(s/2,0),o=(n-Math.max(i.cutoutPercentage?n/100*i.cutoutPercentage:1,0))/t.getVisibleDatasetCount();this.outerRadius=n-o*this.index,this.innerRadius=this.outerRadius-o}updateElements(t,e,i,s){const n="reset"===s,o=this.chart,a=o.options.animation,r=this._cachedMeta.rScale,l=r.xCenter,h=r.yCenter,c=r.getIndexAngle(0)-.5*C;let d,u=c;const f=360/this.countVisibleElements();for(d=0;d<e;++d)u+=this._computeAngle(d,s,f);for(d=e;d<e+i;d++){const e=t[d];let i=u,g=u+this._computeAngle(d,s,f),p=o.getDataVisibility(d)?r.getDistanceFromCenterForValue(this.getParsed(d).r):0;u=g,n&&(a.animateScale&&(p=0),a.animateRotate&&(i=g=c));const m={x:l,y:h,innerRadius:0,outerRadius:p,startAngle:i,endAngle:g,options:this.resolveDataElementOptions(d,e.active?"active":s)};this.updateElement(e,d,m,s)}}countVisibleElements(){const t=this._cachedMeta;let e=0;return t.data.forEach(((t,i)=>{!isNaN(this.getParsed(i).r)&&this.chart.getDataVisibility(i)&&e++})),e}_computeAngle(t,e,i){return this.chart.getDataVisibility(t)?$(this.resolveDataElementOptions(t,e).angle||i):0}}var Yn=Object.freeze({__proto__:null,BarController:class extends Ns{static id="bar";static defaults={datasetElementType:!1,dataElementType:"bar",categoryPercentage:.8,barPercentage:.9,grouped:!0,animations:{numbers:{type:"number",properties:["x","y","base","width","height"]}}};static overrides={scales:{_index_:{type:"category",offset:!0,grid:{offset:!0}},_value_:{type:"linear",beginAtZero:!0}}};parsePrimitiveData(t,e,i,s){return Fn(t,e,i,s)}parseArrayData(t,e,i,s){return Fn(t,e,i,s)}parseObjectData(t,e,i,s){const{iScale:n,vScale:o}=t,{xAxisKey:a="x",yAxisKey:r="y"}=this._parsing,l="x"===n.axis?a:r,h="x"===o.axis?a:r,c=[];let d,u,f,g;for(d=i,u=i+s;d<u;++d)g=e[d],f={},f[n.axis]=n.parse(M(g,l),d),c.push(zn(M(g,h),f,o,d));return c}updateRangeFromParsed(t,e,i,s){super.updateRangeFromParsed(t,e,i,s);const n=i._custom;n&&e===this._cachedMeta.vScale&&(t.min=Math.min(t.min,n.min),t.max=Math.max(t.max,n.max))}getMaxOverflow(){return 0}getLabelAndValue(t){const e=this._cachedMeta,{iScale:i,vScale:s}=e,n=this.getParsed(t),o=n._custom,a=Vn(o)?"["+o.start+", "+o.end+"]":""+s.getLabelForValue(n[s.axis]);return{label:""+i.getLabelForValue(n[i.axis]),value:a}}initialize(){this.enableOptionSharing=!0,super.initialize();this._cachedMeta.stack=this.getDataset().stack}update(t){const e=this._cachedMeta;this.updateElements(e.data,0,e.data.length,t)}updateElements(t,e,i,n){const o="reset"===n,{index:a,_cachedMeta:{vScale:r}}=this,l=r.getBasePixel(),h=r.isHorizontal(),c=this._getRuler(),{sharedOptions:d,includeOptions:u}=this._getSharedOptions(e,n);for(let f=e;f<e+i;f++){const e=this.getParsed(f),i=o||s(e[r.axis])?{base:l,head:l}:this._calculateBarValuePixels(f),g=this._calculateBarIndexPixels(f,c),p=(e._stacks||{})[r.axis],m={horizontal:h,base:i.base,enableBorderRadius:!p||Vn(e._custom)||a===p._top||a===p._bottom,x:h?i.head:g.center,y:h?g.center:i.head,height:h?g.size:Math.abs(i.size),width:h?Math.abs(i.size):g.size};u&&(m.options=d||this.resolveDataElementOptions(f,t[f].active?"active":n));const x=m.options||t[f].options;Bn(m,x,p,a),Hn(m,x,c.ratio),this.updateElement(t[f],f,m,n)}}_getStacks(t,e){const{iScale:i}=this._cachedMeta,n=i.getMatchingVisibleMetas(this._type).filter((t=>t.controller.options.grouped)),o=i.options.stacked,a=[],r=t=>{const i=t.controller.getParsed(e),n=i&&i[t.vScale.axis];if(s(n)||isNaN(n))return!0};for(const i of n)if((void 0===e||!r(i))&&((!1===o||-1===a.indexOf(i.stack)||void 0===o&&void 0===i.stack)&&a.push(i.stack),i.index===t))break;return a.length||a.push(void 0),a}_getStackCount(t){return this._getStacks(void 0,t).length}_getStackIndex(t,e,i){const s=this._getStacks(t,i),n=void 0!==e?s.indexOf(e):-1;return-1===n?s.length-1:n}_getRuler(){const t=this.options,e=this._cachedMeta,i=e.iScale,s=[];let n,o;for(n=0,o=e.data.length;n<o;++n)s.push(i.getPixelForValue(this.getParsed(n)[i.axis],n));const a=t.barThickness;return{min:a||In(e),pixels:s,start:i._startPixel,end:i._endPixel,stackCount:this._getStackCount(),scale:i,grouped:t.grouped,ratio:a?1:t.categoryPercentage*t.barPercentage}}_calculateBarValuePixels(t){const{_cachedMeta:{vScale:e,_stacked:i,index:n},options:{base:o,minBarLength:a}}=this,r=o||0,l=this.getParsed(t),h=l._custom,c=Vn(h);let d,u,f=l[e.axis],g=0,p=i?this.applyStack(e,l,i):f;p!==f&&(g=p-f,p=f),c&&(f=h.barStart,p=h.barEnd-h.barStart,0!==f&&F(f)!==F(h.barEnd)&&(g=0),g+=f);const m=s(o)||c?g:o;let x=e.getPixelForValue(m);if(d=this.chart.getDataVisibility(t)?e.getPixelForValue(g+p):x,u=d-x,Math.abs(u)<a){u=function(t,e,i){return 0!==t?F(t):(e.isHorizontal()?1:-1)*(e.min>=i?1:-1)}(u,e,r)*a,f===r&&(x-=u/2);const t=e.getPixelForDecimal(0),s=e.getPixelForDecimal(1),o=Math.min(t,s),h=Math.max(t,s);x=Math.max(Math.min(x,h),o),d=x+u,i&&!c&&(l._stacks[e.axis]._visualValues[n]=e.getValueForPixel(d)-e.getValueForPixel(x))}if(x===e.getPixelForValue(r)){const t=F(u)*e.getLineWidthForValue(r)/2;x+=t,u-=t}return{size:u,base:x,head:d,center:d+u/2}}_calculateBarIndexPixels(t,e){const i=e.scale,n=this.options,o=n.skipNull,a=l(n.maxBarThickness,1/0);let r,h;if(e.grouped){const i=o?this._getStackCount(t):e.stackCount,l="flex"===n.barThickness?function(t,e,i,s){const n=e.pixels,o=n[t];let a=t>0?n[t-1]:null,r=t<n.length-1?n[t+1]:null;const l=i.categoryPercentage;null===a&&(a=o-(null===r?e.end-e.start:r-o)),null===r&&(r=o+o-a);const h=o-(o-Math.min(a,r))/2*l;return{chunk:Math.abs(r-a)/2*l/s,ratio:i.barPercentage,start:h}}(t,e,n,i):function(t,e,i,n){const o=i.barThickness;let a,r;return s(o)?(a=e.min*i.categoryPercentage,r=i.barPercentage):(a=o*n,r=1),{chunk:a/n,ratio:r,start:e.pixels[t]-a/2}}(t,e,n,i),c=this._getStackIndex(this.index,this._cachedMeta.stack,o?t:void 0);r=l.start+l.chunk*c+l.chunk/2,h=Math.min(a,l.chunk*l.ratio)}else r=i.getPixelForValue(this.getParsed(t)[i.axis],t),h=Math.min(a,e.min*e.ratio);return{base:r-h/2,head:r+h/2,center:r,size:h}}draw(){const t=this._cachedMeta,e=t.vScale,i=t.data,s=i.length;let n=0;for(;n<s;++n)null===this.getParsed(n)[e.axis]||i[n].hidden||i[n].draw(this._ctx)}},BubbleController:class extends Ns{static id="bubble";static defaults={datasetElementType:!1,dataElementType:"point",animations:{numbers:{type:"number",properties:["x","y","borderWidth","radius"]}}};static overrides={scales:{x:{type:"linear"},y:{type:"linear"}}};initialize(){this.enableOptionSharing=!0,super.initialize()}parsePrimitiveData(t,e,i,s){const n=super.parsePrimitiveData(t,e,i,s);for(let t=0;t<n.length;t++)n[t]._custom=this.resolveDataElementOptions(t+i).radius;return n}parseArrayData(t,e,i,s){const n=super.parseArrayData(t,e,i,s);for(let t=0;t<n.length;t++){const s=e[i+t];n[t]._custom=l(s[2],this.resolveDataElementOptions(t+i).radius)}return n}parseObjectData(t,e,i,s){const n=super.parseObjectData(t,e,i,s);for(let t=0;t<n.length;t++){const s=e[i+t];n[t]._custom=l(s&&s.r&&+s.r,this.resolveDataElementOptions(t+i).radius)}return n}getMaxOverflow(){const t=this._cachedMeta.data;let e=0;for(let i=t.length-1;i>=0;--i)e=Math.max(e,t[i].size(this.resolveDataElementOptions(i))/2);return e>0&&e}getLabelAndValue(t){const e=this._cachedMeta,i=this.chart.data.labels||[],{xScale:s,yScale:n}=e,o=this.getParsed(t),a=s.getLabelForValue(o.x),r=n.getLabelForValue(o.y),l=o._custom;return{label:i[t]||"",value:"("+a+", "+r+(l?", "+l:"")+")"}}update(t){const e=this._cachedMeta.data;this.updateElements(e,0,e.length,t)}updateElements(t,e,i,s){const n="reset"===s,{iScale:o,vScale:a}=this._cachedMeta,{sharedOptions:r,includeOptions:l}=this._getSharedOptions(e,s),h=o.axis,c=a.axis;for(let d=e;d<e+i;d++){const e=t[d],i=!n&&this.getParsed(d),u={},f=u[h]=n?o.getPixelForDecimal(.5):o.getPixelForValue(i[h]),g=u[c]=n?a.getBasePixel():a.getPixelForValue(i[c]);u.skip=isNaN(f)||isNaN(g),l&&(u.options=r||this.resolveDataElementOptions(d,e.active?"active":s),n&&(u.options.radius=0)),this.updateElement(e,d,u,s)}}resolveDataElementOptions(t,e){const i=this.getParsed(t);let s=super.resolveDataElementOptions(t,e);s.$shared&&(s=Object.assign({},s,{$shared:!1}));const n=s.radius;return"active"!==e&&(s.radius=0),s.radius+=l(i&&i._custom,n),s}},DoughnutController:jn,LineController:class extends Ns{static id="line";static defaults={datasetElementType:"line",dataElementType:"point",showLine:!0,spanGaps:!1};static overrides={scales:{_index_:{type:"category"},_value_:{type:"linear"}}};initialize(){this.enableOptionSharing=!0,this.supportsDecimation=!0,super.initialize()}update(t){const e=this._cachedMeta,{dataset:i,data:s=[],_dataset:n}=e,o=this.chart._animationsDisabled;let{start:a,count:r}=pt(e,s,o);this._drawStart=a,this._drawCount=r,mt(e)&&(a=0,r=s.length),i._chart=this.chart,i._datasetIndex=this.index,i._decimated=!!n._decimated,i.points=s;const l=this.resolveDatasetElementOptions(t);this.options.showLine||(l.borderWidth=0),l.segment=this.options.segment,this.updateElement(i,void 0,{animated:!o,options:l},t),this.updateElements(s,a,r,t)}updateElements(t,e,i,n){const o="reset"===n,{iScale:a,vScale:r,_stacked:l,_dataset:h}=this._cachedMeta,{sharedOptions:c,includeOptions:d}=this._getSharedOptions(e,n),u=a.axis,f=r.axis,{spanGaps:g,segment:p}=this.options,m=N(g)?g:Number.POSITIVE_INFINITY,x=this.chart._animationsDisabled||o||"none"===n,b=e+i,_=t.length;let y=e>0&&this.getParsed(e-1);for(let i=0;i<_;++i){const g=t[i],_=x?g:{};if(i<e||i>=b){_.skip=!0;continue}const v=this.getParsed(i),M=s(v[f]),w=_[u]=a.getPixelForValue(v[u],i),k=_[f]=o||M?r.getBasePixel():r.getPixelForValue(l?this.applyStack(r,v,l):v[f],i);_.skip=isNaN(w)||isNaN(k)||M,_.stop=i>0&&Math.abs(v[u]-y[u])>m,p&&(_.parsed=v,_.raw=h.data[i]),d&&(_.options=c||this.resolveDataElementOptions(i,g.active?"active":n)),x||this.updateElement(g,i,_,n),y=v}}getMaxOverflow(){const t=this._cachedMeta,e=t.dataset,i=e.options&&e.options.borderWidth||0,s=t.data||[];if(!s.length)return i;const n=s[0].size(this.resolveDataElementOptions(0)),o=s[s.length-1].size(this.resolveDataElementOptions(s.length-1));return Math.max(i,n,o)/2}draw(){const t=this._cachedMeta;t.dataset.updateControlPoints(this.chart.chartArea,t.iScale.axis),super.draw()}},PieController:class extends jn{static id="pie";static defaults={cutout:0,rotation:0,circumference:360,radius:"100%"}},PolarAreaController:$n,RadarController:class extends Ns{static id="radar";static defaults={datasetElementType:"line",dataElementType:"point",indexAxis:"r",showLine:!0,elements:{line:{fill:"start"}}};static overrides={aspectRatio:1,scales:{r:{type:"radialLinear"}}};getLabelAndValue(t){const e=this._cachedMeta.vScale,i=this.getParsed(t);return{label:e.getLabels()[t],value:""+e.getLabelForValue(i[e.axis])}}parseObjectData(t,e,i,s){return ii.bind(this)(t,e,i,s)}update(t){const e=this._cachedMeta,i=e.dataset,s=e.data||[],n=e.iScale.getLabels();if(i.points=s,"resize"!==t){const e=this.resolveDatasetElementOptions(t);this.options.showLine||(e.borderWidth=0);const o={_loop:!0,_fullLoop:n.length===s.length,options:e};this.updateElement(i,void 0,o,t)}this.updateElements(s,0,s.length,t)}updateElements(t,e,i,s){const n=this._cachedMeta.rScale,o="reset"===s;for(let a=e;a<e+i;a++){const e=t[a],i=this.resolveDataElementOptions(a,e.active?"active":s),r=n.getPointPositionForValue(a,this.getParsed(a).r),l=o?n.xCenter:r.x,h=o?n.yCenter:r.y,c={x:l,y:h,angle:r.angle,skip:isNaN(l)||isNaN(h),options:i};this.updateElement(e,a,c,s)}}},ScatterController:class extends Ns{static id="scatter";static defaults={datasetElementType:!1,dataElementType:"point",showLine:!1,fill:!1};static overrides={interaction:{mode:"point"},scales:{x:{type:"linear"},y:{type:"linear"}}};getLabelAndValue(t){const e=this._cachedMeta,i=this.chart.data.labels||[],{xScale:s,yScale:n}=e,o=this.getParsed(t),a=s.getLabelForValue(o.x),r=n.getLabelForValue(o.y);return{label:i[t]||"",value:"("+a+", "+r+")"}}update(t){const e=this._cachedMeta,{data:i=[]}=e,s=this.chart._animationsDisabled;let{start:n,count:o}=pt(e,i,s);if(this._drawStart=n,this._drawCount=o,mt(e)&&(n=0,o=i.length),this.options.showLine){this.datasetElementType||this.addElements();const{dataset:n,_dataset:o}=e;n._chart=this.chart,n._datasetIndex=this.index,n._decimated=!!o._decimated,n.points=i;const a=this.resolveDatasetElementOptions(t);a.segment=this.options.segment,this.updateElement(n,void 0,{animated:!s,options:a},t)}else this.datasetElementType&&(delete e.dataset,this.datasetElementType=!1);this.updateElements(i,n,o,t)}addElements(){const{showLine:t}=this.options;!this.datasetElementType&&t&&(this.datasetElementType=this.chart.registry.getElement("line")),super.addElements()}updateElements(t,e,i,n){const o="reset"===n,{iScale:a,vScale:r,_stacked:l,_dataset:h}=this._cachedMeta,c=this.resolveDataElementOptions(e,n),d=this.getSharedOptions(c),u=this.includeOptions(n,d),f=a.axis,g=r.axis,{spanGaps:p,segment:m}=this.options,x=N(p)?p:Number.POSITIVE_INFINITY,b=this.chart._animationsDisabled||o||"none"===n;let _=e>0&&this.getParsed(e-1);for(let c=e;c<e+i;++c){const e=t[c],i=this.getParsed(c),p=b?e:{},y=s(i[g]),v=p[f]=a.getPixelForValue(i[f],c),M=p[g]=o||y?r.getBasePixel():r.getPixelForValue(l?this.applyStack(r,i,l):i[g],c);p.skip=isNaN(v)||isNaN(M)||y,p.stop=c>0&&Math.abs(i[f]-_[f])>x,m&&(p.parsed=i,p.raw=h.data[c]),u&&(p.options=d||this.resolveDataElementOptions(c,e.active?"active":n)),b||this.updateElement(e,c,p,n),_=i}this.updateSharedOptions(d,n,c)}getMaxOverflow(){const t=this._cachedMeta,e=t.data||[];if(!this.options.showLine){let t=0;for(let i=e.length-1;i>=0;--i)t=Math.max(t,e[i].size(this.resolveDataElementOptions(i))/2);return t>0&&t}const i=t.dataset,s=i.options&&i.options.borderWidth||0;if(!e.length)return s;const n=e[0].size(this.resolveDataElementOptions(0)),o=e[e.length-1].size(this.resolveDataElementOptions(e.length-1));return Math.max(s,n,o)/2}}});function Un(t,e,i,s){const n=vi(t.options.borderRadius,["outerStart","outerEnd","innerStart","innerEnd"]);const o=(i-e)/2,a=Math.min(o,s*e/2),r=t=>{const e=(i-Math.min(o,t))*s/2;return J(t,0,Math.min(o,e))};return{outerStart:r(n.outerStart),outerEnd:r(n.outerEnd),innerStart:J(n.innerStart,0,a),innerEnd:J(n.innerEnd,0,a)}}function Xn(t,e,i,s){return{x:i+t*Math.cos(e),y:s+t*Math.sin(e)}}function qn(t,e,i,s,n,o){const{x:a,y:r,startAngle:l,pixelMargin:h,innerRadius:c}=e,d=Math.max(e.outerRadius+s+i-h,0),u=c>0?c+s+i+h:0;let f=0;const g=n-l;if(s){const t=((c>0?c-s:0)+(d>0?d-s:0))/2;f=(g-(0!==t?g*t/(t+s):g))/2}const p=(g-Math.max(.001,g*d-i/C)/d)/2,m=l+p+f,x=n-p-f,{outerStart:b,outerEnd:_,innerStart:y,innerEnd:v}=Un(e,u,d,x-m),M=d-b,w=d-_,k=m+b/M,S=x-_/w,P=u+y,D=u+v,O=m+y/P,A=x-v/D;if(t.beginPath(),o){const e=(k+S)/2;if(t.arc(a,r,d,k,e),t.arc(a,r,d,e,S),_>0){const e=Xn(w,S,a,r);t.arc(e.x,e.y,_,S,x+E)}const i=Xn(D,x,a,r);if(t.lineTo(i.x,i.y),v>0){const e=Xn(D,A,a,r);t.arc(e.x,e.y,v,x+E,A+Math.PI)}const s=(x-v/u+(m+y/u))/2;if(t.arc(a,r,u,x-v/u,s,!0),t.arc(a,r,u,s,m+y/u,!0),y>0){const e=Xn(P,O,a,r);t.arc(e.x,e.y,y,O+Math.PI,m-E)}const n=Xn(M,m,a,r);if(t.lineTo(n.x,n.y),b>0){const e=Xn(M,k,a,r);t.arc(e.x,e.y,b,m-E,k)}}else{t.moveTo(a,r);const e=Math.cos(k)*d+a,i=Math.sin(k)*d+r;t.lineTo(e,i);const s=Math.cos(S)*d+a,n=Math.sin(S)*d+r;t.lineTo(s,n)}t.closePath()}function Kn(t,e,i,s,n){const{fullCircles:o,startAngle:a,circumference:r,options:l}=e,{borderWidth:h,borderJoinStyle:c,borderDash:d,borderDashOffset:u}=l,f="inner"===l.borderAlign;if(!h)return;t.setLineDash(d||[]),t.lineDashOffset=u,f?(t.lineWidth=2*h,t.lineJoin=c||"round"):(t.lineWidth=h,t.lineJoin=c||"bevel");let g=e.endAngle;if(o){qn(t,e,i,s,g,n);for(let e=0;e<o;++e)t.stroke();isNaN(r)||(g=a+(r%O||O))}f&&function(t,e,i){const{startAngle:s,pixelMargin:n,x:o,y:a,outerRadius:r,innerRadius:l}=e;let h=n/r;t.beginPath(),t.arc(o,a,r,s-h,i+h),l>n?(h=n/l,t.arc(o,a,l,i+h,s-h,!0)):t.arc(o,a,n,i+E,s-E),t.closePath(),t.clip()}(t,e,g),o||(qn(t,e,i,s,g,n),t.stroke())}function Gn(t,e,i=e){t.lineCap=l(i.borderCapStyle,e.borderCapStyle),t.setLineDash(l(i.borderDash,e.borderDash)),t.lineDashOffset=l(i.borderDashOffset,e.borderDashOffset),t.lineJoin=l(i.borderJoinStyle,e.borderJoinStyle),t.lineWidth=l(i.borderWidth,e.borderWidth),t.strokeStyle=l(i.borderColor,e.borderColor)}function Zn(t,e,i){t.lineTo(i.x,i.y)}function Jn(t,e,i={}){const s=t.length,{start:n=0,end:o=s-1}=i,{start:a,end:r}=e,l=Math.max(n,a),h=Math.min(o,r),c=n<a&&o<a||n>r&&o>r;return{count:s,start:l,loop:e.loop,ilen:h<l&&!c?s+h-l:h-l}}function Qn(t,e,i,s){const{points:n,options:o}=e,{count:a,start:r,loop:l,ilen:h}=Jn(n,i,s),c=function(t){return t.stepped?Fe:t.tension||"monotone"===t.cubicInterpolationMode?Ve:Zn}(o);let d,u,f,{move:g=!0,reverse:p}=s||{};for(d=0;d<=h;++d)u=n[(r+(p?h-d:d))%a],u.skip||(g?(t.moveTo(u.x,u.y),g=!1):c(t,f,u,p,o.stepped),f=u);return l&&(u=n[(r+(p?h:0))%a],c(t,f,u,p,o.stepped)),!!l}function to(t,e,i,s){const n=e.points,{count:o,start:a,ilen:r}=Jn(n,i,s),{move:l=!0,reverse:h}=s||{};let c,d,u,f,g,p,m=0,x=0;const b=t=>(a+(h?r-t:t))%o,_=()=>{f!==g&&(t.lineTo(m,g),t.lineTo(m,f),t.lineTo(m,p))};for(l&&(d=n[b(0)],t.moveTo(d.x,d.y)),c=0;c<=r;++c){if(d=n[b(c)],d.skip)continue;const e=d.x,i=d.y,s=0|e;s===u?(i<f?f=i:i>g&&(g=i),m=(x*m+e)/++x):(_(),t.lineTo(e,i),u=s,x=0,f=g=i),p=i}_()}function eo(t){const e=t.options,i=e.borderDash&&e.borderDash.length;return!(t._decimated||t._loop||e.tension||"monotone"===e.cubicInterpolationMode||e.stepped||i)?to:Qn}const io="function"==typeof Path2D;function so(t,e,i,s){io&&!e.options.segment?function(t,e,i,s){let n=e._path;n||(n=e._path=new Path2D,e.path(n,i,s)&&n.closePath()),Gn(t,e.options),t.stroke(n)}(t,e,i,s):function(t,e,i,s){const{segments:n,options:o}=e,a=eo(e);for(const r of n)Gn(t,o,r.style),t.beginPath(),a(t,e,r,{start:i,end:i+s-1})&&t.closePath(),t.stroke()}(t,e,i,s)}class no extends Hs{static id="line";static defaults={borderCapStyle:"butt",borderDash:[],borderDashOffset:0,borderJoinStyle:"miter",borderWidth:3,capBezierPoints:!0,cubicInterpolationMode:"default",fill:!1,spanGaps:!1,stepped:!1,tension:0};static defaultRoutes={backgroundColor:"backgroundColor",borderColor:"borderColor"};static descriptors={_scriptable:!0,_indexable:t=>"borderDash"!==t&&"fill"!==t};constructor(t){super(),this.animated=!0,this.options=void 0,this._chart=void 0,this._loop=void 0,this._fullLoop=void 0,this._path=void 0,this._points=void 0,this._segments=void 0,this._decimated=!1,this._pointsUpdated=!1,this._datasetIndex=void 0,t&&Object.assign(this,t)}updateControlPoints(t,e){const i=this.options;if((i.tension||"monotone"===i.cubicInterpolationMode)&&!i.stepped&&!this._pointsUpdated){const s=i.spanGaps?this._loop:this._fullLoop;hi(this._points,i,t,s,e),this._pointsUpdated=!0}}set points(t){this._points=t,delete this._segments,delete this._path,this._pointsUpdated=!1}get points(){return this._points}get segments(){return this._segments||(this._segments=zi(this,this.options.segment))}first(){const t=this.segments,e=this.points;return t.length&&e[t[0].start]}last(){const t=this.segments,e=this.points,i=t.length;return i&&e[t[i-1].end]}interpolate(t,e){const i=this.options,s=t[e],n=this.points,o=Ii(this,{property:e,start:s,end:s});if(!o.length)return;const a=[],r=function(t){return t.stepped?pi:t.tension||"monotone"===t.cubicInterpolationMode?mi:gi}(i);let l,h;for(l=0,h=o.length;l<h;++l){const{start:h,end:c}=o[l],d=n[h],u=n[c];if(d===u){a.push(d);continue}const f=r(d,u,Math.abs((s-d[e])/(u[e]-d[e])),i.stepped);f[e]=t[e],a.push(f)}return 1===a.length?a[0]:a}pathSegment(t,e,i){return eo(this)(t,this,e,i)}path(t,e,i){const s=this.segments,n=eo(this);let o=this._loop;e=e||0,i=i||this.points.length-e;for(const a of s)o&=n(t,this,a,{start:e,end:e+i-1});return!!o}draw(t,e,i,s){const n=this.options||{};(this.points||[]).length&&n.borderWidth&&(t.save(),so(t,this,i,s),t.restore()),this.animated&&(this._pointsUpdated=!1,this._path=void 0)}}function oo(t,e,i,s){const n=t.options,{[i]:o}=t.getProps([i],s);return Math.abs(e-o)<n.radius+n.hitRadius}function ao(t,e){const{x:i,y:s,base:n,width:o,height:a}=t.getProps(["x","y","base","width","height"],e);let r,l,h,c,d;return t.horizontal?(d=a/2,r=Math.min(i,n),l=Math.max(i,n),h=s-d,c=s+d):(d=o/2,r=i-d,l=i+d,h=Math.min(s,n),c=Math.max(s,n)),{left:r,top:h,right:l,bottom:c}}function ro(t,e,i,s){return t?0:J(e,i,s)}function lo(t){const e=ao(t),i=e.right-e.left,s=e.bottom-e.top,n=function(t,e,i){const s=t.options.borderWidth,n=t.borderSkipped,o=Mi(s);return{t:ro(n.top,o.top,0,i),r:ro(n.right,o.right,0,e),b:ro(n.bottom,o.bottom,0,i),l:ro(n.left,o.left,0,e)}}(t,i/2,s/2),a=function(t,e,i){const{enableBorderRadius:s}=t.getProps(["enableBorderRadius"]),n=t.options.borderRadius,a=wi(n),r=Math.min(e,i),l=t.borderSkipped,h=s||o(n);return{topLeft:ro(!h||l.top||l.left,a.topLeft,0,r),topRight:ro(!h||l.top||l.right,a.topRight,0,r),bottomLeft:ro(!h||l.bottom||l.left,a.bottomLeft,0,r),bottomRight:ro(!h||l.bottom||l.right,a.bottomRight,0,r)}}(t,i/2,s/2);return{outer:{x:e.left,y:e.top,w:i,h:s,radius:a},inner:{x:e.left+n.l,y:e.top+n.t,w:i-n.l-n.r,h:s-n.t-n.b,radius:{topLeft:Math.max(0,a.topLeft-Math.max(n.t,n.l)),topRight:Math.max(0,a.topRight-Math.max(n.t,n.r)),bottomLeft:Math.max(0,a.bottomLeft-Math.max(n.b,n.l)),bottomRight:Math.max(0,a.bottomRight-Math.max(n.b,n.r))}}}}function ho(t,e,i,s){const n=null===e,o=null===i,a=t&&!(n&&o)&&ao(t,s);return a&&(n||tt(e,a.left,a.right))&&(o||tt(i,a.top,a.bottom))}function co(t,e){t.rect(e.x,e.y,e.w,e.h)}function uo(t,e,i={}){const s=t.x!==i.x?-e:0,n=t.y!==i.y?-e:0,o=(t.x+t.w!==i.x+i.w?e:0)-s,a=(t.y+t.h!==i.y+i.h?e:0)-n;return{x:t.x+s,y:t.y+n,w:t.w+o,h:t.h+a,radius:t.radius}}var fo=Object.freeze({__proto__:null,ArcElement:class extends Hs{static id="arc";static defaults={borderAlign:"center",borderColor:"#fff",borderDash:[],borderDashOffset:0,borderJoinStyle:void 0,borderRadius:0,borderWidth:2,offset:0,spacing:0,angle:void 0,circular:!0};static defaultRoutes={backgroundColor:"backgroundColor"};static descriptors={_scriptable:!0,_indexable:t=>"borderDash"!==t};circumference;endAngle;fullCircles;innerRadius;outerRadius;pixelMargin;startAngle;constructor(t){super(),this.options=void 0,this.circumference=void 0,this.startAngle=void 0,this.endAngle=void 0,this.innerRadius=void 0,this.outerRadius=void 0,this.pixelMargin=0,this.fullCircles=0,t&&Object.assign(this,t)}inRange(t,e,i){const s=this.getProps(["x","y"],i),{angle:n,distance:o}=X(s,{x:t,y:e}),{startAngle:a,endAngle:r,innerRadius:h,outerRadius:c,circumference:d}=this.getProps(["startAngle","endAngle","innerRadius","outerRadius","circumference"],i),u=(this.options.spacing+this.options.borderWidth)/2,f=l(d,r-a)>=O||Z(n,a,r),g=tt(o,h+u,c+u);return f&&g}getCenterPoint(t){const{x:e,y:i,startAngle:s,endAngle:n,innerRadius:o,outerRadius:a}=this.getProps(["x","y","startAngle","endAngle","innerRadius","outerRadius"],t),{offset:r,spacing:l}=this.options,h=(s+n)/2,c=(o+a+l+r)/2;return{x:e+Math.cos(h)*c,y:i+Math.sin(h)*c}}tooltipPosition(t){return this.getCenterPoint(t)}draw(t){const{options:e,circumference:i}=this,s=(e.offset||0)/4,n=(e.spacing||0)/2,o=e.circular;if(this.pixelMargin="inner"===e.borderAlign?.33:0,this.fullCircles=i>O?Math.floor(i/O):0,0===i||this.innerRadius<0||this.outerRadius<0)return;t.save();const a=(this.startAngle+this.endAngle)/2;t.translate(Math.cos(a)*s,Math.sin(a)*s);const r=s*(1-Math.sin(Math.min(C,i||0)));t.fillStyle=e.backgroundColor,t.strokeStyle=e.borderColor,function(t,e,i,s,n){const{fullCircles:o,startAngle:a,circumference:r}=e;let l=e.endAngle;if(o){qn(t,e,i,s,l,n);for(let e=0;e<o;++e)t.fill();isNaN(r)||(l=a+(r%O||O))}qn(t,e,i,s,l,n),t.fill()}(t,this,r,n,o),Kn(t,this,r,n,o),t.restore()}},BarElement:class extends Hs{static id="bar";static defaults={borderSkipped:"start",borderWidth:0,borderRadius:0,inflateAmount:"auto",pointStyle:void 0};static defaultRoutes={backgroundColor:"backgroundColor",borderColor:"borderColor"};constructor(t){super(),this.options=void 0,this.horizontal=void 0,this.base=void 0,this.width=void 0,this.height=void 0,this.inflateAmount=void 0,t&&Object.assign(this,t)}draw(t){const{inflateAmount:e,options:{borderColor:i,backgroundColor:s}}=this,{inner:n,outer:o}=lo(this),a=(r=o.radius).topLeft||r.topRight||r.bottomLeft||r.bottomRight?He:co;var r;t.save(),o.w===n.w&&o.h===n.h||(t.beginPath(),a(t,uo(o,e,n)),t.clip(),a(t,uo(n,-e,o)),t.fillStyle=i,t.fill("evenodd")),t.beginPath(),a(t,uo(n,e)),t.fillStyle=s,t.fill(),t.restore()}inRange(t,e,i){return ho(this,t,e,i)}inXRange(t,e){return ho(this,t,null,e)}inYRange(t,e){return ho(this,null,t,e)}getCenterPoint(t){const{x:e,y:i,base:s,horizontal:n}=this.getProps(["x","y","base","horizontal"],t);return{x:n?(e+s)/2:e,y:n?i:(i+s)/2}}getRange(t){return"x"===t?this.width/2:this.height/2}},LineElement:no,PointElement:class extends Hs{static id="point";parsed;skip;stop;static defaults={borderWidth:1,hitRadius:1,hoverBorderWidth:1,hoverRadius:4,pointStyle:"circle",radius:3,rotation:0};static defaultRoutes={backgroundColor:"backgroundColor",borderColor:"borderColor"};constructor(t){super(),this.options=void 0,this.parsed=void 0,this.skip=void 0,this.stop=void 0,t&&Object.assign(this,t)}inRange(t,e,i){const s=this.options,{x:n,y:o}=this.getProps(["x","y"],i);return Math.pow(t-n,2)+Math.pow(e-o,2)<Math.pow(s.hitRadius+s.radius,2)}inXRange(t,e){return oo(this,t,"x",e)}inYRange(t,e){return oo(this,t,"y",e)}getCenterPoint(t){const{x:e,y:i}=this.getProps(["x","y"],t);return{x:e,y:i}}size(t){let e=(t=t||this.options||{}).radius||0;e=Math.max(e,e&&t.hoverRadius||0);return 2*(e+(e&&t.borderWidth||0))}draw(t,e){const i=this.options;this.skip||i.radius<.1||!Re(this,e,this.size(i)/2)||(t.strokeStyle=i.borderColor,t.lineWidth=i.borderWidth,t.fillStyle=i.backgroundColor,Le(t,i,this.x,this.y))}getRange(){const t=this.options||{};return t.radius+t.hitRadius}}});function go(t,e,i,s){const n=t.indexOf(e);if(-1===n)return((t,e,i,s)=>("string"==typeof e?(i=t.push(e)-1,s.unshift({index:i,label:e})):isNaN(e)&&(i=null),i))(t,e,i,s);return n!==t.lastIndexOf(e)?i:n}function po(t){const e=this.getLabels();return t>=0&&t<e.length?e[t]:t}function mo(t,e,{horizontal:i,minRotation:s}){const n=$(s),o=(i?Math.sin(n):Math.cos(n))||.001,a=.75*e*(""+t).length;return Math.min(e/o,a)}class xo extends Js{constructor(t){super(t),this.start=void 0,this.end=void 0,this._startValue=void 0,this._endValue=void 0,this._valueRange=0}parse(t,e){return s(t)||("number"==typeof t||t instanceof Number)&&!isFinite(+t)?null:+t}handleTickRangeOptions(){const{beginAtZero:t}=this.options,{minDefined:e,maxDefined:i}=this.getUserBounds();let{min:s,max:n}=this;const o=t=>s=e?s:t,a=t=>n=i?n:t;if(t){const t=F(s),e=F(n);t<0&&e<0?a(0):t>0&&e>0&&o(0)}if(s===n){let e=0===n?1:Math.abs(.05*n);a(n+e),t||o(s-e)}this.min=s,this.max=n}getTickLimit(){const t=this.options.ticks;let e,{maxTicksLimit:i,stepSize:s}=t;return s?(e=Math.ceil(this.max/s)-Math.floor(this.min/s)+1,e>1e3&&(console.warn(`scales.${this.id}.ticks.stepSize: ${s} would result generating up to ${e} ticks. Limiting to 1000.`),e=1e3)):(e=this.computeTickLimit(),i=i||11),i&&(e=Math.min(i,e)),e}computeTickLimit(){return Number.POSITIVE_INFINITY}buildTicks(){const t=this.options,e=t.ticks;let i=this.getTickLimit();i=Math.max(2,i);const n=function(t,e){const i=[],{bounds:n,step:o,min:a,max:r,precision:l,count:h,maxTicks:c,maxDigits:d,includeBounds:u}=t,f=o||1,g=c-1,{min:p,max:m}=e,x=!s(a),b=!s(r),_=!s(h),y=(m-p)/(d+1);let v,M,w,k,S=B((m-p)/g/f)*f;if(S<1e-14&&!x&&!b)return[{value:p},{value:m}];k=Math.ceil(m/S)-Math.floor(p/S),k>g&&(S=B(k*S/g/f)*f),s(l)||(v=Math.pow(10,l),S=Math.ceil(S*v)/v),"ticks"===n?(M=Math.floor(p/S)*S,w=Math.ceil(m/S)*S):(M=p,w=m),x&&b&&o&&H((r-a)/o,S/1e3)?(k=Math.round(Math.min((r-a)/S,c)),S=(r-a)/k,M=a,w=r):_?(M=x?a:M,w=b?r:w,k=h-1,S=(w-M)/k):(k=(w-M)/S,k=V(k,Math.round(k),S/1e3)?Math.round(k):Math.ceil(k));const P=Math.max(U(S),U(M));v=Math.pow(10,s(l)?P:l),M=Math.round(M*v)/v,w=Math.round(w*v)/v;let D=0;for(x&&(u&&M!==a?(i.push({value:a}),M<a&&D++,V(Math.round((M+D*S)*v)/v,a,mo(a,y,t))&&D++):M<a&&D++);D<k;++D){const t=Math.round((M+D*S)*v)/v;if(b&&t>r)break;i.push({value:t})}return b&&u&&w!==r?i.length&&V(i[i.length-1].value,r,mo(r,y,t))?i[i.length-1].value=r:i.push({value:r}):b&&w!==r||i.push({value:w}),i}({maxTicks:i,bounds:t.bounds,min:t.min,max:t.max,precision:e.precision,step:e.stepSize,count:e.count,maxDigits:this._maxDigits(),horizontal:this.isHorizontal(),minRotation:e.minRotation||0,includeBounds:!1!==e.includeBounds},this._range||this);return"ticks"===t.bounds&&j(n,this,"value"),t.reverse?(n.reverse(),this.start=this.max,this.end=this.min):(this.start=this.min,this.end=this.max),n}configure(){const t=this.ticks;let e=this.min,i=this.max;if(super.configure(),this.options.offset&&t.length){const s=(i-e)/Math.max(t.length-1,1)/2;e-=s,i+=s}this._startValue=e,this._endValue=i,this._valueRange=i-e}getLabelForValue(t){return ne(t,this.chart.options.locale,this.options.ticks.format)}}class bo extends xo{static id="linear";static defaults={ticks:{callback:ae.formatters.numeric}};determineDataLimits(){const{min:t,max:e}=this.getMinMax(!0);this.min=a(t)?t:0,this.max=a(e)?e:1,this.handleTickRangeOptions()}computeTickLimit(){const t=this.isHorizontal(),e=t?this.width:this.height,i=$(this.options.ticks.minRotation),s=(t?Math.sin(i):Math.cos(i))||.001,n=this._resolveTickFontOptions(0);return Math.ceil(e/Math.min(40,n.lineHeight/s))}getPixelForValue(t){return null===t?NaN:this.getPixelForDecimal((t-this._startValue)/this._valueRange)}getValueForPixel(t){return this._startValue+this.getDecimalForPixel(t)*this._valueRange}}const _o=t=>Math.floor(z(t)),yo=(t,e)=>Math.pow(10,_o(t)+e);function vo(t){return 1===t/Math.pow(10,_o(t))}function Mo(t,e,i){const s=Math.pow(10,i),n=Math.floor(t/s);return Math.ceil(e/s)-n}function wo(t,{min:e,max:i}){e=r(t.min,e);const s=[],n=_o(e);let o=function(t,e){let i=_o(e-t);for(;Mo(t,e,i)>10;)i++;for(;Mo(t,e,i)<10;)i--;return Math.min(i,_o(t))}(e,i),a=o<0?Math.pow(10,Math.abs(o)):1;const l=Math.pow(10,o),h=n>o?Math.pow(10,n):0,c=Math.round((e-h)*a)/a,d=Math.floor((e-h)/l/10)*l*10;let u=Math.floor((c-d)/Math.pow(10,o)),f=r(t.min,Math.round((h+d+u*Math.pow(10,o))*a)/a);for(;f<i;)s.push({value:f,major:vo(f),significand:u}),u>=10?u=u<15?15:20:u++,u>=20&&(o++,u=2,a=o>=0?1:a),f=Math.round((h+d+u*Math.pow(10,o))*a)/a;const g=r(t.max,f);return s.push({value:g,major:vo(g),significand:u}),s}class ko extends Js{static id="logarithmic";static defaults={ticks:{callback:ae.formatters.logarithmic,major:{enabled:!0}}};constructor(t){super(t),this.start=void 0,this.end=void 0,this._startValue=void 0,this._valueRange=0}parse(t,e){const i=xo.prototype.parse.apply(this,[t,e]);if(0!==i)return a(i)&&i>0?i:null;this._zero=!0}determineDataLimits(){const{min:t,max:e}=this.getMinMax(!0);this.min=a(t)?Math.max(0,t):null,this.max=a(e)?Math.max(0,e):null,this.options.beginAtZero&&(this._zero=!0),this._zero&&this.min!==this._suggestedMin&&!a(this._userMin)&&(this.min=t===yo(this.min,0)?yo(this.min,-1):yo(this.min,0)),this.handleTickRangeOptions()}handleTickRangeOptions(){const{minDefined:t,maxDefined:e}=this.getUserBounds();let i=this.min,s=this.max;const n=e=>i=t?i:e,o=t=>s=e?s:t;i===s&&(i<=0?(n(1),o(10)):(n(yo(i,-1)),o(yo(s,1)))),i<=0&&n(yo(s,-1)),s<=0&&o(yo(i,1)),this.min=i,this.max=s}buildTicks(){const t=this.options,e=wo({min:this._userMin,max:this._userMax},this);return"ticks"===t.bounds&&j(e,this,"value"),t.reverse?(e.reverse(),this.start=this.max,this.end=this.min):(this.start=this.min,this.end=this.max),e}getLabelForValue(t){return void 0===t?"0":ne(t,this.chart.options.locale,this.options.ticks.format)}configure(){const t=this.min;super.configure(),this._startValue=z(t),this._valueRange=z(this.max)-z(t)}getPixelForValue(t){return void 0!==t&&0!==t||(t=this.min),null===t||isNaN(t)?NaN:this.getPixelForDecimal(t===this.min?0:(z(t)-this._startValue)/this._valueRange)}getValueForPixel(t){const e=this.getDecimalForPixel(t);return Math.pow(10,this._startValue+e*this._valueRange)}}function So(t){const e=t.ticks;if(e.display&&t.display){const t=ki(e.backdropPadding);return l(e.font&&e.font.size,ue.font.size)+t.height}return 0}function Po(t,e,i,s,n){return t===s||t===n?{start:e-i/2,end:e+i/2}:t<s||t>n?{start:e-i,end:e}:{start:e,end:e+i}}function Do(t){const e={l:t.left+t._padding.left,r:t.right-t._padding.right,t:t.top+t._padding.top,b:t.bottom-t._padding.bottom},i=Object.assign({},e),s=[],o=[],a=t._pointLabels.length,r=t.options.pointLabels,l=r.centerPointLabels?C/a:0;for(let u=0;u<a;u++){const a=r.setContext(t.getPointLabelContext(u));o[u]=a.padding;const f=t.getPointPosition(u,t.drawingArea+o[u],l),g=Si(a.font),p=(h=t.ctx,c=g,d=n(d=t._pointLabels[u])?d:[d],{w:Oe(h,c.string,d),h:d.length*c.lineHeight});s[u]=p;const m=G(t.getIndexAngle(u)+l),x=Math.round(Y(m));Co(i,e,m,Po(x,f.x,p.w,0,180),Po(x,f.y,p.h,90,270))}var h,c,d;t.setCenterPoint(e.l-i.l,i.r-e.r,e.t-i.t,i.b-e.b),t._pointLabelItems=function(t,e,i){const s=[],n=t._pointLabels.length,o=t.options,{centerPointLabels:a,display:r}=o.pointLabels,l={extra:So(o)/2,additionalAngle:a?C/n:0};let h;for(let o=0;o<n;o++){l.padding=i[o],l.size=e[o];const n=Oo(t,o,l);s.push(n),"auto"===r&&(n.visible=Ao(n,h),n.visible&&(h=n))}return s}(t,s,o)}function Co(t,e,i,s,n){const o=Math.abs(Math.sin(i)),a=Math.abs(Math.cos(i));let r=0,l=0;s.start<e.l?(r=(e.l-s.start)/o,t.l=Math.min(t.l,e.l-r)):s.end>e.r&&(r=(s.end-e.r)/o,t.r=Math.max(t.r,e.r+r)),n.start<e.t?(l=(e.t-n.start)/a,t.t=Math.min(t.t,e.t-l)):n.end>e.b&&(l=(n.end-e.b)/a,t.b=Math.max(t.b,e.b+l))}function Oo(t,e,i){const s=t.drawingArea,{extra:n,additionalAngle:o,padding:a,size:r}=i,l=t.getPointPosition(e,s+n+a,o),h=Math.round(Y(G(l.angle+E))),c=function(t,e,i){90===i||270===i?t-=e/2:(i>270||i<90)&&(t-=e);return t}(l.y,r.h,h),d=function(t){if(0===t||180===t)return"center";if(t<180)return"left";return"right"}(h),u=function(t,e,i){"right"===i?t-=e:"center"===i&&(t-=e/2);return t}(l.x,r.w,d);return{visible:!0,x:l.x,y:c,textAlign:d,left:u,top:c,right:u+r.w,bottom:c+r.h}}function Ao(t,e){if(!e)return!0;const{left:i,top:s,right:n,bottom:o}=t;return!(Re({x:i,y:s},e)||Re({x:i,y:o},e)||Re({x:n,y:s},e)||Re({x:n,y:o},e))}function To(t,e,i){const{left:n,top:o,right:a,bottom:r}=i,{backdropColor:l}=e;if(!s(l)){const i=wi(e.borderRadius),s=ki(e.backdropPadding);t.fillStyle=l;const h=n-s.left,c=o-s.top,d=a-n+s.width,u=r-o+s.height;Object.values(i).some((t=>0!==t))?(t.beginPath(),He(t,{x:h,y:c,w:d,h:u,radius:i}),t.fill()):t.fillRect(h,c,d,u)}}function Lo(t,e,i,s){const{ctx:n}=t;if(i)n.arc(t.xCenter,t.yCenter,e,0,O);else{let i=t.getPointPosition(0,e);n.moveTo(i.x,i.y);for(let o=1;o<s;o++)i=t.getPointPosition(o,e),n.lineTo(i.x,i.y)}}class Eo extends xo{static id="radialLinear";static defaults={display:!0,animate:!0,position:"chartArea",angleLines:{display:!0,lineWidth:1,borderDash:[],borderDashOffset:0},grid:{circular:!1},startAngle:0,ticks:{showLabelBackdrop:!0,callback:ae.formatters.numeric},pointLabels:{backdropColor:void 0,backdropPadding:2,display:!0,font:{size:10},callback:t=>t,padding:5,centerPointLabels:!1}};static defaultRoutes={"angleLines.color":"borderColor","pointLabels.color":"color","ticks.color":"color"};static descriptors={angleLines:{_fallback:"grid"}};constructor(t){super(t),this.xCenter=void 0,this.yCenter=void 0,this.drawingArea=void 0,this._pointLabels=[],this._pointLabelItems=[]}setDimensions(){const t=this._padding=ki(So(this.options)/2),e=this.width=this.maxWidth-t.width,i=this.height=this.maxHeight-t.height;this.xCenter=Math.floor(this.left+e/2+t.left),this.yCenter=Math.floor(this.top+i/2+t.top),this.drawingArea=Math.floor(Math.min(e,i)/2)}determineDataLimits(){const{min:t,max:e}=this.getMinMax(!1);this.min=a(t)&&!isNaN(t)?t:0,this.max=a(e)&&!isNaN(e)?e:0,this.handleTickRangeOptions()}computeTickLimit(){return Math.ceil(this.drawingArea/So(this.options))}generateTickLabels(t){xo.prototype.generateTickLabels.call(this,t),this._pointLabels=this.getLabels().map(((t,e)=>{const i=d(this.options.pointLabels.callback,[t,e],this);return i||0===i?i:""})).filter(((t,e)=>this.chart.getDataVisibility(e)))}fit(){const t=this.options;t.display&&t.pointLabels.display?Do(this):this.setCenterPoint(0,0,0,0)}setCenterPoint(t,e,i,s){this.xCenter+=Math.floor((t-e)/2),this.yCenter+=Math.floor((i-s)/2),this.drawingArea-=Math.min(this.drawingArea/2,Math.max(t,e,i,s))}getIndexAngle(t){return G(t*(O/(this._pointLabels.length||1))+$(this.options.startAngle||0))}getDistanceFromCenterForValue(t){if(s(t))return NaN;const e=this.drawingArea/(this.max-this.min);return this.options.reverse?(this.max-t)*e:(t-this.min)*e}getValueForDistanceFromCenter(t){if(s(t))return NaN;const e=t/(this.drawingArea/(this.max-this.min));return this.options.reverse?this.max-e:this.min+e}getPointLabelContext(t){const e=this._pointLabels||[];if(t>=0&&t<e.length){const i=e[t];return function(t,e,i){return Ci(t,{label:i,index:e,type:"pointLabel"})}(this.getContext(),t,i)}}getPointPosition(t,e,i=0){const s=this.getIndexAngle(t)-E+i;return{x:Math.cos(s)*e+this.xCenter,y:Math.sin(s)*e+this.yCenter,angle:s}}getPointPositionForValue(t,e){return this.getPointPosition(t,this.getDistanceFromCenterForValue(e))}getBasePosition(t){return this.getPointPositionForValue(t||0,this.getBaseValue())}getPointLabelPosition(t){const{left:e,top:i,right:s,bottom:n}=this._pointLabelItems[t];return{left:e,top:i,right:s,bottom:n}}drawBackground(){const{backgroundColor:t,grid:{circular:e}}=this.options;if(t){const i=this.ctx;i.save(),i.beginPath(),Lo(this,this.getDistanceFromCenterForValue(this._endValue),e,this._pointLabels.length),i.closePath(),i.fillStyle=t,i.fill(),i.restore()}}drawGrid(){const t=this.ctx,e=this.options,{angleLines:i,grid:s,border:n}=e,o=this._pointLabels.length;let a,r,l;if(e.pointLabels.display&&function(t,e){const{ctx:i,options:{pointLabels:s}}=t;for(let n=e-1;n>=0;n--){const e=t._pointLabelItems[n];if(!e.visible)continue;const o=s.setContext(t.getPointLabelContext(n));To(i,o,e);const a=Si(o.font),{x:r,y:l,textAlign:h}=e;Ne(i,t._pointLabels[n],r,l+a.lineHeight/2,a,{color:o.color,textAlign:h,textBaseline:"middle"})}}(this,o),s.display&&this.ticks.forEach(((t,e)=>{if(0!==e||0===e&&this.min<0){r=this.getDistanceFromCenterForValue(t.value);const i=this.getContext(e),a=s.setContext(i),l=n.setContext(i);!function(t,e,i,s,n){const o=t.ctx,a=e.circular,{color:r,lineWidth:l}=e;!a&&!s||!r||!l||i<0||(o.save(),o.strokeStyle=r,o.lineWidth=l,o.setLineDash(n.dash),o.lineDashOffset=n.dashOffset,o.beginPath(),Lo(t,i,a,s),o.closePath(),o.stroke(),o.restore())}(this,a,r,o,l)}})),i.display){for(t.save(),a=o-1;a>=0;a--){const s=i.setContext(this.getPointLabelContext(a)),{color:n,lineWidth:o}=s;o&&n&&(t.lineWidth=o,t.strokeStyle=n,t.setLineDash(s.borderDash),t.lineDashOffset=s.borderDashOffset,r=this.getDistanceFromCenterForValue(e.ticks.reverse?this.min:this.max),l=this.getPointPosition(a,r),t.beginPath(),t.moveTo(this.xCenter,this.yCenter),t.lineTo(l.x,l.y),t.stroke())}t.restore()}}drawBorder(){}drawLabels(){const t=this.ctx,e=this.options,i=e.ticks;if(!i.display)return;const s=this.getIndexAngle(0);let n,o;t.save(),t.translate(this.xCenter,this.yCenter),t.rotate(s),t.textAlign="center",t.textBaseline="middle",this.ticks.forEach(((s,a)=>{if(0===a&&this.min>=0&&!e.reverse)return;const r=i.setContext(this.getContext(a)),l=Si(r.font);if(n=this.getDistanceFromCenterForValue(this.ticks[a].value),r.showLabelBackdrop){t.font=l.string,o=t.measureText(s.label).width,t.fillStyle=r.backdropColor;const e=ki(r.backdropPadding);t.fillRect(-o/2-e.left,-n-l.size/2-e.top,o+e.width,l.size+e.height)}Ne(t,s.label,0,-n,l,{color:r.color,strokeColor:r.textStrokeColor,strokeWidth:r.textStrokeWidth})})),t.restore()}drawTitle(){}}const Ro={millisecond:{common:!0,size:1,steps:1e3},second:{common:!0,size:1e3,steps:60},minute:{common:!0,size:6e4,steps:60},hour:{common:!0,size:36e5,steps:24},day:{common:!0,size:864e5,steps:30},week:{common:!1,size:6048e5,steps:4},month:{common:!0,size:2628e6,steps:12},quarter:{common:!1,size:7884e6,steps:4},year:{common:!0,size:3154e7}},Io=Object.keys(Ro);function zo(t,e){return t-e}function Fo(t,e){if(s(e))return null;const i=t._adapter,{parser:n,round:o,isoWeekday:r}=t._parseOpts;let l=e;return"function"==typeof n&&(l=n(l)),a(l)||(l="string"==typeof n?i.parse(l,n):i.parse(l)),null===l?null:(o&&(l="week"!==o||!N(r)&&!0!==r?i.startOf(l,o):i.startOf(l,"isoWeek",r)),+l)}function Vo(t,e,i,s){const n=Io.length;for(let o=Io.indexOf(t);o<n-1;++o){const t=Ro[Io[o]],n=t.steps?t.steps:Number.MAX_SAFE_INTEGER;if(t.common&&Math.ceil((i-e)/(n*t.size))<=s)return Io[o]}return Io[n-1]}function Bo(t,e,i){if(i){if(i.length){const{lo:s,hi:n}=et(i,e);t[i[s]>=e?i[s]:i[n]]=!0}}else t[e]=!0}function Wo(t,e,i){const s=[],n={},o=e.length;let a,r;for(a=0;a<o;++a)r=e[a],n[r]=a,s.push({value:r,major:!1});return 0!==o&&i?function(t,e,i,s){const n=t._adapter,o=+n.startOf(e[0].value,s),a=e[e.length-1].value;let r,l;for(r=o;r<=a;r=+n.add(r,1,s))l=i[r],l>=0&&(e[l].major=!0);return e}(t,s,n,i):s}class No extends Js{static id="time";static defaults={bounds:"data",adapters:{},time:{parser:!1,unit:!1,round:!1,isoWeekday:!1,minUnit:"millisecond",displayFormats:{}},ticks:{source:"auto",callback:!1,major:{enabled:!1}}};constructor(t){super(t),this._cache={data:[],labels:[],all:[]},this._unit="day",this._majorUnit=void 0,this._offsets={},this._normalized=!1,this._parseOpts=void 0}init(t,e={}){const i=t.time||(t.time={}),s=this._adapter=new Rn._date(t.adapters.date);s.init(e),b(i.displayFormats,s.formats()),this._parseOpts={parser:i.parser,round:i.round,isoWeekday:i.isoWeekday},super.init(t),this._normalized=e.normalized}parse(t,e){return void 0===t?null:Fo(this,t)}beforeLayout(){super.beforeLayout(),this._cache={data:[],labels:[],all:[]}}determineDataLimits(){const t=this.options,e=this._adapter,i=t.time.unit||"day";let{min:s,max:n,minDefined:o,maxDefined:r}=this.getUserBounds();function l(t){o||isNaN(t.min)||(s=Math.min(s,t.min)),r||isNaN(t.max)||(n=Math.max(n,t.max))}o&&r||(l(this._getLabelBounds()),"ticks"===t.bounds&&"labels"===t.ticks.source||l(this.getMinMax(!1))),s=a(s)&&!isNaN(s)?s:+e.startOf(Date.now(),i),n=a(n)&&!isNaN(n)?n:+e.endOf(Date.now(),i)+1,this.min=Math.min(s,n-1),this.max=Math.max(s+1,n)}_getLabelBounds(){const t=this.getLabelTimestamps();let e=Number.POSITIVE_INFINITY,i=Number.NEGATIVE_INFINITY;return t.length&&(e=t[0],i=t[t.length-1]),{min:e,max:i}}buildTicks(){const t=this.options,e=t.time,i=t.ticks,s="labels"===i.source?this.getLabelTimestamps():this._generate();"ticks"===t.bounds&&s.length&&(this.min=this._userMin||s[0],this.max=this._userMax||s[s.length-1]);const n=this.min,o=nt(s,n,this.max);return this._unit=e.unit||(i.autoSkip?Vo(e.minUnit,this.min,this.max,this._getLabelCapacity(n)):function(t,e,i,s,n){for(let o=Io.length-1;o>=Io.indexOf(i);o--){const i=Io[o];if(Ro[i].common&&t._adapter.diff(n,s,i)>=e-1)return i}return Io[i?Io.indexOf(i):0]}(this,o.length,e.minUnit,this.min,this.max)),this._majorUnit=i.major.enabled&&"year"!==this._unit?function(t){for(let e=Io.indexOf(t)+1,i=Io.length;e<i;++e)if(Ro[Io[e]].common)return Io[e]}(this._unit):void 0,this.initOffsets(s),t.reverse&&o.reverse(),Wo(this,o,this._majorUnit)}afterAutoSkip(){this.options.offsetAfterAutoskip&&this.initOffsets(this.ticks.map((t=>+t.value)))}initOffsets(t=[]){let e,i,s=0,n=0;this.options.offset&&t.length&&(e=this.getDecimalForValue(t[0]),s=1===t.length?1-e:(this.getDecimalForValue(t[1])-e)/2,i=this.getDecimalForValue(t[t.length-1]),n=1===t.length?i:(i-this.getDecimalForValue(t[t.length-2]))/2);const o=t.length<3?.5:.25;s=J(s,0,o),n=J(n,0,o),this._offsets={start:s,end:n,factor:1/(s+1+n)}}_generate(){const t=this._adapter,e=this.min,i=this.max,s=this.options,n=s.time,o=n.unit||Vo(n.minUnit,e,i,this._getLabelCapacity(e)),a=l(s.ticks.stepSize,1),r="week"===o&&n.isoWeekday,h=N(r)||!0===r,c={};let d,u,f=e;if(h&&(f=+t.startOf(f,"isoWeek",r)),f=+t.startOf(f,h?"day":o),t.diff(i,e,o)>1e5*a)throw new Error(e+" and "+i+" are too far apart with stepSize of "+a+" "+o);const g="data"===s.ticks.source&&this.getDataTimestamps();for(d=f,u=0;d<i;d=+t.add(d,a,o),u++)Bo(c,d,g);return d!==i&&"ticks"!==s.bounds&&1!==u||Bo(c,d,g),Object.keys(c).sort(zo).map((t=>+t))}getLabelForValue(t){const e=this._adapter,i=this.options.time;return i.tooltipFormat?e.format(t,i.tooltipFormat):e.format(t,i.displayFormats.datetime)}format(t,e){const i=this.options.time.displayFormats,s=this._unit,n=e||i[s];return this._adapter.format(t,n)}_tickFormatFunction(t,e,i,s){const n=this.options,o=n.ticks.callback;if(o)return d(o,[t,e,i],this);const a=n.time.displayFormats,r=this._unit,l=this._majorUnit,h=r&&a[r],c=l&&a[l],u=i[e],f=l&&c&&u&&u.major;return this._adapter.format(t,s||(f?c:h))}generateTickLabels(t){let e,i,s;for(e=0,i=t.length;e<i;++e)s=t[e],s.label=this._tickFormatFunction(s.value,e,t)}getDecimalForValue(t){return null===t?NaN:(t-this.min)/(this.max-this.min)}getPixelForValue(t){const e=this._offsets,i=this.getDecimalForValue(t);return this.getPixelForDecimal((e.start+i)*e.factor)}getValueForPixel(t){const e=this._offsets,i=this.getDecimalForPixel(t)/e.factor-e.end;return this.min+i*(this.max-this.min)}_getLabelSize(t){const e=this.options.ticks,i=this.ctx.measureText(t).width,s=$(this.isHorizontal()?e.maxRotation:e.minRotation),n=Math.cos(s),o=Math.sin(s),a=this._resolveTickFontOptions(0).size;return{w:i*n+a*o,h:i*o+a*n}}_getLabelCapacity(t){const e=this.options.time,i=e.displayFormats,s=i[e.unit]||i.millisecond,n=this._tickFormatFunction(t,0,Wo(this,[t],this._majorUnit),s),o=this._getLabelSize(n),a=Math.floor(this.isHorizontal()?this.width/o.w:this.height/o.h)-1;return a>0?a:1}getDataTimestamps(){let t,e,i=this._cache.data||[];if(i.length)return i;const s=this.getMatchingVisibleMetas();if(this._normalized&&s.length)return this._cache.data=s[0].controller.getAllParsedValues(this);for(t=0,e=s.length;t<e;++t)i=i.concat(s[t].controller.getAllParsedValues(this));return this._cache.data=this.normalize(i)}getLabelTimestamps(){const t=this._cache.labels||[];let e,i;if(t.length)return t;const s=this.getLabels();for(e=0,i=s.length;e<i;++e)t.push(Fo(this,s[e]));return this._cache.labels=this._normalized?t:this.normalize(t)}normalize(t){return lt(t.sort(zo))}}function Ho(t,e,i){let s,n,o,a,r=0,l=t.length-1;i?(e>=t[r].pos&&e<=t[l].pos&&({lo:r,hi:l}=it(t,"pos",e)),({pos:s,time:o}=t[r]),({pos:n,time:a}=t[l])):(e>=t[r].time&&e<=t[l].time&&({lo:r,hi:l}=it(t,"time",e)),({time:s,pos:o}=t[r]),({time:n,pos:a}=t[l]));const h=n-s;return h?o+(a-o)*(e-s)/h:o}var jo=Object.freeze({__proto__:null,CategoryScale:class extends Js{static id="category";static defaults={ticks:{callback:po}};constructor(t){super(t),this._startValue=void 0,this._valueRange=0,this._addedLabels=[]}init(t){const e=this._addedLabels;if(e.length){const t=this.getLabels();for(const{index:i,label:s}of e)t[i]===s&&t.splice(i,1);this._addedLabels=[]}super.init(t)}parse(t,e){if(s(t))return null;const i=this.getLabels();return((t,e)=>null===t?null:J(Math.round(t),0,e))(e=isFinite(e)&&i[e]===t?e:go(i,t,l(e,t),this._addedLabels),i.length-1)}determineDataLimits(){const{minDefined:t,maxDefined:e}=this.getUserBounds();let{min:i,max:s}=this.getMinMax(!0);"ticks"===this.options.bounds&&(t||(i=0),e||(s=this.getLabels().length-1)),this.min=i,this.max=s}buildTicks(){const t=this.min,e=this.max,i=this.options.offset,s=[];let n=this.getLabels();n=0===t&&e===n.length-1?n:n.slice(t,e+1),this._valueRange=Math.max(n.length-(i?0:1),1),this._startValue=this.min-(i?.5:0);for(let i=t;i<=e;i++)s.push({value:i});return s}getLabelForValue(t){return po.call(this,t)}configure(){super.configure(),this.isHorizontal()||(this._reversePixels=!this._reversePixels)}getPixelForValue(t){return"number"!=typeof t&&(t=this.parse(t)),null===t?NaN:this.getPixelForDecimal((t-this._startValue)/this._valueRange)}getPixelForTick(t){const e=this.ticks;return t<0||t>e.length-1?null:this.getPixelForValue(e[t].value)}getValueForPixel(t){return Math.round(this._startValue+this.getDecimalForPixel(t)*this._valueRange)}getBasePixel(){return this.bottom}},LinearScale:bo,LogarithmicScale:ko,RadialLinearScale:Eo,TimeScale:No,TimeSeriesScale:class extends No{static id="timeseries";static defaults=No.defaults;constructor(t){super(t),this._table=[],this._minPos=void 0,this._tableRange=void 0}initOffsets(){const t=this._getTimestampsForTable(),e=this._table=this.buildLookupTable(t);this._minPos=Ho(e,this.min),this._tableRange=Ho(e,this.max)-this._minPos,super.initOffsets(t)}buildLookupTable(t){const{min:e,max:i}=this,s=[],n=[];let o,a,r,l,h;for(o=0,a=t.length;o<a;++o)l=t[o],l>=e&&l<=i&&s.push(l);if(s.length<2)return[{time:e,pos:0},{time:i,pos:1}];for(o=0,a=s.length;o<a;++o)h=s[o+1],r=s[o-1],l=s[o],Math.round((h+r)/2)!==l&&n.push({time:l,pos:o/(a-1)});return n}_generate(){const t=this.min,e=this.max;let i=super.getDataTimestamps();return i.includes(t)&&i.length||i.splice(0,0,t),i.includes(e)&&1!==i.length||i.push(e),i.sort(((t,e)=>t-e))}_getTimestampsForTable(){let t=this._cache.all||[];if(t.length)return t;const e=this.getDataTimestamps(),i=this.getLabelTimestamps();return t=e.length&&i.length?this.normalize(e.concat(i)):e.length?e:i,t=this._cache.all=t,t}getDecimalForValue(t){return(Ho(this._table,t)-this._minPos)/this._tableRange}getValueForPixel(t){const e=this._offsets,i=this.getDecimalForPixel(t)/e.factor-e.end;return Ho(this._table,i*this._tableRange+this._minPos,!0)}}});const $o=["rgb(54, 162, 235)","rgb(255, 99, 132)","rgb(255, 159, 64)","rgb(255, 205, 86)","rgb(75, 192, 192)","rgb(153, 102, 255)","rgb(201, 203, 207)"],Yo=$o.map((t=>t.replace("rgb(","rgba(").replace(")",", 0.5)")));function Uo(t){return $o[t%$o.length]}function Xo(t){return Yo[t%Yo.length]}function qo(t){let e=0;return(i,s)=>{const n=t.getDatasetMeta(s).controller;n instanceof jn?e=function(t,e){return t.backgroundColor=t.data.map((()=>Uo(e++))),e}(i,e):n instanceof $n?e=function(t,e){return t.backgroundColor=t.data.map((()=>Xo(e++))),e}(i,e):n&&(e=function(t,e){return t.borderColor=Uo(e),t.backgroundColor=Xo(e),++e}(i,e))}}function Ko(t){let e;for(e in t)if(t[e].borderColor||t[e].backgroundColor)return!0;return!1}var Go={id:"colors",defaults:{enabled:!0,forceOverride:!1},beforeLayout(t,e,i){if(!i.enabled)return;const{data:{datasets:s},options:n}=t.config,{elements:o}=n;if(!i.forceOverride&&(Ko(s)||(a=n)&&(a.borderColor||a.backgroundColor)||o&&Ko(o)))return;var a;const r=qo(t);s.forEach(r)}};function Zo(t){if(t._decimated){const e=t._data;delete t._decimated,delete t._data,Object.defineProperty(t,"data",{configurable:!0,enumerable:!0,writable:!0,value:e})}}function Jo(t){t.data.datasets.forEach((t=>{Zo(t)}))}var Qo={id:"decimation",defaults:{algorithm:"min-max",enabled:!1},beforeElementsUpdate:(t,e,i)=>{if(!i.enabled)return void Jo(t);const n=t.width;t.data.datasets.forEach(((e,o)=>{const{_data:a,indexAxis:r}=e,l=t.getDatasetMeta(o),h=a||e.data;if("y"===Pi([r,t.options.indexAxis]))return;if(!l.controller.supportsDecimation)return;const c=t.scales[l.xAxisID];if("linear"!==c.type&&"time"!==c.type)return;if(t.options.parsing)return;let{start:d,count:u}=function(t,e){const i=e.length;let s,n=0;const{iScale:o}=t,{min:a,max:r,minDefined:l,maxDefined:h}=o.getUserBounds();return l&&(n=J(it(e,o.axis,a).lo,0,i-1)),s=h?J(it(e,o.axis,r).hi+1,n,i)-n:i-n,{start:n,count:s}}(l,h);if(u<=(i.threshold||4*n))return void Zo(e);let f;switch(s(a)&&(e._data=h,delete e.data,Object.defineProperty(e,"data",{configurable:!0,enumerable:!0,get:function(){return this._decimated},set:function(t){this._data=t}})),i.algorithm){case"lttb":f=function(t,e,i,s,n){const o=n.samples||s;if(o>=i)return t.slice(e,e+i);const a=[],r=(i-2)/(o-2);let l=0;const h=e+i-1;let c,d,u,f,g,p=e;for(a[l++]=t[p],c=0;c<o-2;c++){let s,n=0,o=0;const h=Math.floor((c+1)*r)+1+e,m=Math.min(Math.floor((c+2)*r)+1,i)+e,x=m-h;for(s=h;s<m;s++)n+=t[s].x,o+=t[s].y;n/=x,o/=x;const b=Math.floor(c*r)+1+e,_=Math.min(Math.floor((c+1)*r)+1,i)+e,{x:y,y:v}=t[p];for(u=f=-1,s=b;s<_;s++)f=.5*Math.abs((y-n)*(t[s].y-v)-(y-t[s].x)*(o-v)),f>u&&(u=f,d=t[s],g=s);a[l++]=d,p=g}return a[l++]=t[h],a}(h,d,u,n,i);break;case"min-max":f=function(t,e,i,n){let o,a,r,l,h,c,d,u,f,g,p=0,m=0;const x=[],b=e+i-1,_=t[e].x,y=t[b].x-_;for(o=e;o<e+i;++o){a=t[o],r=(a.x-_)/y*n,l=a.y;const e=0|r;if(e===h)l<f?(f=l,c=o):l>g&&(g=l,d=o),p=(m*p+a.x)/++m;else{const i=o-1;if(!s(c)&&!s(d)){const e=Math.min(c,d),s=Math.max(c,d);e!==u&&e!==i&&x.push({...t[e],x:p}),s!==u&&s!==i&&x.push({...t[s],x:p})}o>0&&i!==u&&x.push(t[i]),x.push(a),h=e,m=0,f=g=l,c=d=u=o}}return x}(h,d,u,n);break;default:throw new Error(`Unsupported decimation algorithm '${i.algorithm}'`)}e._decimated=f}))},destroy(t){Jo(t)}};function ta(t,e,i,s){if(s)return;let n=e[t],o=i[t];return"angle"===t&&(n=G(n),o=G(o)),{property:t,start:n,end:o}}function ea(t,e,i){for(;e>t;e--){const t=i[e];if(!isNaN(t.x)&&!isNaN(t.y))break}return e}function ia(t,e,i,s){return t&&e?s(t[i],e[i]):t?t[i]:e?e[i]:0}function sa(t,e){let i=[],s=!1;return n(t)?(s=!0,i=t):i=function(t,e){const{x:i=null,y:s=null}=t||{},n=e.points,o=[];return e.segments.forEach((({start:t,end:e})=>{e=ea(t,e,n);const a=n[t],r=n[e];null!==s?(o.push({x:a.x,y:s}),o.push({x:r.x,y:s})):null!==i&&(o.push({x:i,y:a.y}),o.push({x:i,y:r.y}))})),o}(t,e),i.length?new no({points:i,options:{tension:0},_loop:s,_fullLoop:s}):null}function na(t){return t&&!1!==t.fill}function oa(t,e,i){let s=t[e].fill;const n=[e];let o;if(!i)return s;for(;!1!==s&&-1===n.indexOf(s);){if(!a(s))return s;if(o=t[s],!o)return!1;if(o.visible)return s;n.push(s),s=o.fill}return!1}function aa(t,e,i){const s=function(t){const e=t.options,i=e.fill;let s=l(i&&i.target,i);void 0===s&&(s=!!e.backgroundColor);if(!1===s||null===s)return!1;if(!0===s)return"origin";return s}(t);if(o(s))return!isNaN(s.value)&&s;let n=parseFloat(s);return a(n)&&Math.floor(n)===n?function(t,e,i,s){"-"!==t&&"+"!==t||(i=e+i);if(i===e||i<0||i>=s)return!1;return i}(s[0],e,n,i):["origin","start","end","stack","shape"].indexOf(s)>=0&&s}function ra(t,e,i){const s=[];for(let n=0;n<i.length;n++){const o=i[n],{first:a,last:r,point:l}=la(o,e,"x");if(!(!l||a&&r))if(a)s.unshift(l);else if(t.push(l),!r)break}t.push(...s)}function la(t,e,i){const s=t.interpolate(e,i);if(!s)return{};const n=s[i],o=t.segments,a=t.points;let r=!1,l=!1;for(let t=0;t<o.length;t++){const e=o[t],s=a[e.start][i],h=a[e.end][i];if(tt(n,s,h)){r=n===s,l=n===h;break}}return{first:r,last:l,point:s}}class ha{constructor(t){this.x=t.x,this.y=t.y,this.radius=t.radius}pathSegment(t,e,i){const{x:s,y:n,radius:o}=this;return e=e||{start:0,end:O},t.arc(s,n,o,e.end,e.start,!0),!i.bounds}interpolate(t){const{x:e,y:i,radius:s}=this,n=t.angle;return{x:e+Math.cos(n)*s,y:i+Math.sin(n)*s,angle:n}}}function ca(t){const{chart:e,fill:i,line:s}=t;if(a(i))return function(t,e){const i=t.getDatasetMeta(e),s=i&&t.isDatasetVisible(e);return s?i.dataset:null}(e,i);if("stack"===i)return function(t){const{scale:e,index:i,line:s}=t,n=[],o=s.segments,a=s.points,r=function(t,e){const i=[],s=t.getMatchingVisibleMetas("line");for(let t=0;t<s.length;t++){const n=s[t];if(n.index===e)break;n.hidden||i.unshift(n.dataset)}return i}(e,i);r.push(sa({x:null,y:e.bottom},s));for(let t=0;t<o.length;t++){const e=o[t];for(let t=e.start;t<=e.end;t++)ra(n,a[t],r)}return new no({points:n,options:{}})}(t);if("shape"===i)return!0;const n=function(t){const e=t.scale||{};if(e.getPointPositionForValue)return function(t){const{scale:e,fill:i}=t,s=e.options,n=e.getLabels().length,a=s.reverse?e.max:e.min,r=function(t,e,i){let s;return s="start"===t?i:"end"===t?e.options.reverse?e.min:e.max:o(t)?t.value:e.getBaseValue(),s}(i,e,a),l=[];if(s.grid.circular){const t=e.getPointPositionForValue(0,a);return new ha({x:t.x,y:t.y,radius:e.getDistanceFromCenterForValue(r)})}for(let t=0;t<n;++t)l.push(e.getPointPositionForValue(t,r));return l}(t);return function(t){const{scale:e={},fill:i}=t,s=function(t,e){let i=null;return"start"===t?i=e.bottom:"end"===t?i=e.top:o(t)?i=e.getPixelForValue(t.value):e.getBasePixel&&(i=e.getBasePixel()),i}(i,e);if(a(s)){const t=e.isHorizontal();return{x:t?s:null,y:t?null:s}}return null}(t)}(t);return n instanceof ha?n:sa(n,s)}function da(t,e,i){const s=ca(e),{line:n,scale:o,axis:a}=e,r=n.options,l=r.fill,h=r.backgroundColor,{above:c=h,below:d=h}=l||{};s&&n.points.length&&(Ie(t,i),function(t,e){const{line:i,target:s,above:n,below:o,area:a,scale:r}=e,l=i._loop?"angle":e.axis;t.save(),"x"===l&&o!==n&&(ua(t,s,a.top),fa(t,{line:i,target:s,color:n,scale:r,property:l}),t.restore(),t.save(),ua(t,s,a.bottom));fa(t,{line:i,target:s,color:o,scale:r,property:l}),t.restore()}(t,{line:n,target:s,above:c,below:d,area:i,scale:o,axis:a}),ze(t))}function ua(t,e,i){const{segments:s,points:n}=e;let o=!0,a=!1;t.beginPath();for(const r of s){const{start:s,end:l}=r,h=n[s],c=n[ea(s,l,n)];o?(t.moveTo(h.x,h.y),o=!1):(t.lineTo(h.x,i),t.lineTo(h.x,h.y)),a=!!e.pathSegment(t,r,{move:a}),a?t.closePath():t.lineTo(c.x,i)}t.lineTo(e.first().x,i),t.closePath(),t.clip()}function fa(t,e){const{line:i,target:s,property:n,color:o,scale:a}=e,r=function(t,e,i){const s=t.segments,n=t.points,o=e.points,a=[];for(const t of s){let{start:s,end:r}=t;r=ea(s,r,n);const l=ta(i,n[s],n[r],t.loop);if(!e.segments){a.push({source:t,target:l,start:n[s],end:n[r]});continue}const h=Ii(e,l);for(const e of h){const s=ta(i,o[e.start],o[e.end],e.loop),r=Ri(t,n,s);for(const t of r)a.push({source:t,target:e,start:{[i]:ia(l,s,"start",Math.max)},end:{[i]:ia(l,s,"end",Math.min)}})}}return a}(i,s,n);for(const{source:e,target:l,start:h,end:c}of r){const{style:{backgroundColor:r=o}={}}=e,d=!0!==s;t.save(),t.fillStyle=r,ga(t,a,d&&ta(n,h,c)),t.beginPath();const u=!!i.pathSegment(t,e);let f;if(d){u?t.closePath():pa(t,s,c,n);const e=!!s.pathSegment(t,l,{move:u,reverse:!0});f=u&&e,f||pa(t,s,h,n)}t.closePath(),t.fill(f?"evenodd":"nonzero"),t.restore()}}function ga(t,e,i){const{top:s,bottom:n}=e.chart.chartArea,{property:o,start:a,end:r}=i||{};"x"===o&&(t.beginPath(),t.rect(a,s,r-a,n-s),t.clip())}function pa(t,e,i,s){const n=e.interpolate(i,s);n&&t.lineTo(n.x,n.y)}var ma={id:"filler",afterDatasetsUpdate(t,e,i){const s=(t.data.datasets||[]).length,n=[];let o,a,r,l;for(a=0;a<s;++a)o=t.getDatasetMeta(a),r=o.dataset,l=null,r&&r.options&&r instanceof no&&(l={visible:t.isDatasetVisible(a),index:a,fill:aa(r,a,s),chart:t,axis:o.controller.options.indexAxis,scale:o.vScale,line:r}),o.$filler=l,n.push(l);for(a=0;a<s;++a)l=n[a],l&&!1!==l.fill&&(l.fill=oa(n,a,i.propagate))},beforeDraw(t,e,i){const s="beforeDraw"===i.drawTime,n=t.getSortedVisibleDatasetMetas(),o=t.chartArea;for(let e=n.length-1;e>=0;--e){const i=n[e].$filler;i&&(i.line.updateControlPoints(o,i.axis),s&&i.fill&&da(t.ctx,i,o))}},beforeDatasetsDraw(t,e,i){if("beforeDatasetsDraw"!==i.drawTime)return;const s=t.getSortedVisibleDatasetMetas();for(let e=s.length-1;e>=0;--e){const i=s[e].$filler;na(i)&&da(t.ctx,i,t.chartArea)}},beforeDatasetDraw(t,e,i){const s=e.meta.$filler;na(s)&&"beforeDatasetDraw"===i.drawTime&&da(t.ctx,s,t.chartArea)},defaults:{propagate:!0,drawTime:"beforeDatasetDraw"}};const xa=(t,e)=>{let{boxHeight:i=e,boxWidth:s=e}=t;return t.usePointStyle&&(i=Math.min(i,e),s=t.pointStyleWidth||Math.min(s,e)),{boxWidth:s,boxHeight:i,itemHeight:Math.max(e,i)}};class ba extends Hs{constructor(t){super(),this._added=!1,this.legendHitBoxes=[],this._hoveredItem=null,this.doughnutMode=!1,this.chart=t.chart,this.options=t.options,this.ctx=t.ctx,this.legendItems=void 0,this.columnSizes=void 0,this.lineWidths=void 0,this.maxHeight=void 0,this.maxWidth=void 0,this.top=void 0,this.bottom=void 0,this.left=void 0,this.right=void 0,this.height=void 0,this.width=void 0,this._margins=void 0,this.position=void 0,this.weight=void 0,this.fullSize=void 0}update(t,e,i){this.maxWidth=t,this.maxHeight=e,this._margins=i,this.setDimensions(),this.buildLabels(),this.fit()}setDimensions(){this.isHorizontal()?(this.width=this.maxWidth,this.left=this._margins.left,this.right=this.width):(this.height=this.maxHeight,this.top=this._margins.top,this.bottom=this.height)}buildLabels(){const t=this.options.labels||{};let e=d(t.generateLabels,[this.chart],this)||[];t.filter&&(e=e.filter((e=>t.filter(e,this.chart.data)))),t.sort&&(e=e.sort(((e,i)=>t.sort(e,i,this.chart.data)))),this.options.reverse&&e.reverse(),this.legendItems=e}fit(){const{options:t,ctx:e}=this;if(!t.display)return void(this.width=this.height=0);const i=t.labels,s=Si(i.font),n=s.size,o=this._computeTitleHeight(),{boxWidth:a,itemHeight:r}=xa(i,n);let l,h;e.font=s.string,this.isHorizontal()?(l=this.maxWidth,h=this._fitRows(o,n,a,r)+10):(h=this.maxHeight,l=this._fitCols(o,s,a,r)+10),this.width=Math.min(l,t.maxWidth||this.maxWidth),this.height=Math.min(h,t.maxHeight||this.maxHeight)}_fitRows(t,e,i,s){const{ctx:n,maxWidth:o,options:{labels:{padding:a}}}=this,r=this.legendHitBoxes=[],l=this.lineWidths=[0],h=s+a;let c=t;n.textAlign="left",n.textBaseline="middle";let d=-1,u=-h;return this.legendItems.forEach(((t,f)=>{const g=i+e/2+n.measureText(t.text).width;(0===f||l[l.length-1]+g+2*a>o)&&(c+=h,l[l.length-(f>0?0:1)]=0,u+=h,d++),r[f]={left:0,top:u,row:d,width:g,height:s},l[l.length-1]+=g+a})),c}_fitCols(t,e,i,s){const{ctx:n,maxHeight:o,options:{labels:{padding:a}}}=this,r=this.legendHitBoxes=[],l=this.columnSizes=[],h=o-t;let c=a,d=0,u=0,f=0,g=0;return this.legendItems.forEach(((t,o)=>{const{itemWidth:p,itemHeight:m}=function(t,e,i,s,n){const o=function(t,e,i,s){let n=t.text;n&&"string"!=typeof n&&(n=n.reduce(((t,e)=>t.length>e.length?t:e)));return e+i.size/2+s.measureText(n).width}(s,t,e,i),a=function(t,e,i){let s=t;"string"!=typeof e.text&&(s=_a(e,i));return s}(n,s,e.lineHeight);return{itemWidth:o,itemHeight:a}}(i,e,n,t,s);o>0&&u+m+2*a>h&&(c+=d+a,l.push({width:d,height:u}),f+=d+a,g++,d=u=0),r[o]={left:f,top:u,col:g,width:p,height:m},d=Math.max(d,p),u+=m+a})),c+=d,l.push({width:d,height:u}),c}adjustHitBoxes(){if(!this.options.display)return;const t=this._computeTitleHeight(),{legendHitBoxes:e,options:{align:i,labels:{padding:s},rtl:n}}=this,o=Oi(n,this.left,this.width);if(this.isHorizontal()){let n=0,a=ft(i,this.left+s,this.right-this.lineWidths[n]);for(const r of e)n!==r.row&&(n=r.row,a=ft(i,this.left+s,this.right-this.lineWidths[n])),r.top+=this.top+t+s,r.left=o.leftForLtr(o.x(a),r.width),a+=r.width+s}else{let n=0,a=ft(i,this.top+t+s,this.bottom-this.columnSizes[n].height);for(const r of e)r.col!==n&&(n=r.col,a=ft(i,this.top+t+s,this.bottom-this.columnSizes[n].height)),r.top=a,r.left+=this.left+s,r.left=o.leftForLtr(o.x(r.left),r.width),a+=r.height+s}}isHorizontal(){return"top"===this.options.position||"bottom"===this.options.position}draw(){if(this.options.display){const t=this.ctx;Ie(t,this),this._draw(),ze(t)}}_draw(){const{options:t,columnSizes:e,lineWidths:i,ctx:s}=this,{align:n,labels:o}=t,a=ue.color,r=Oi(t.rtl,this.left,this.width),h=Si(o.font),{padding:c}=o,d=h.size,u=d/2;let f;this.drawTitle(),s.textAlign=r.textAlign("left"),s.textBaseline="middle",s.lineWidth=.5,s.font=h.string;const{boxWidth:g,boxHeight:p,itemHeight:m}=xa(o,d),x=this.isHorizontal(),b=this._computeTitleHeight();f=x?{x:ft(n,this.left+c,this.right-i[0]),y:this.top+c+b,line:0}:{x:this.left+c,y:ft(n,this.top+b+c,this.bottom-e[0].height),line:0},Ai(this.ctx,t.textDirection);const _=m+c;this.legendItems.forEach(((y,v)=>{s.strokeStyle=y.fontColor,s.fillStyle=y.fontColor;const M=s.measureText(y.text).width,w=r.textAlign(y.textAlign||(y.textAlign=o.textAlign)),k=g+u+M;let S=f.x,P=f.y;r.setWidth(this.width),x?v>0&&S+k+c>this.right&&(P=f.y+=_,f.line++,S=f.x=ft(n,this.left+c,this.right-i[f.line])):v>0&&P+_>this.bottom&&(S=f.x=S+e[f.line].width+c,f.line++,P=f.y=ft(n,this.top+b+c,this.bottom-e[f.line].height));if(function(t,e,i){if(isNaN(g)||g<=0||isNaN(p)||p<0)return;s.save();const n=l(i.lineWidth,1);if(s.fillStyle=l(i.fillStyle,a),s.lineCap=l(i.lineCap,"butt"),s.lineDashOffset=l(i.lineDashOffset,0),s.lineJoin=l(i.lineJoin,"miter"),s.lineWidth=n,s.strokeStyle=l(i.strokeStyle,a),s.setLineDash(l(i.lineDash,[])),o.usePointStyle){const a={radius:p*Math.SQRT2/2,pointStyle:i.pointStyle,rotation:i.rotation,borderWidth:n},l=r.xPlus(t,g/2);Ee(s,a,l,e+u,o.pointStyleWidth&&g)}else{const o=e+Math.max((d-p)/2,0),a=r.leftForLtr(t,g),l=wi(i.borderRadius);s.beginPath(),Object.values(l).some((t=>0!==t))?He(s,{x:a,y:o,w:g,h:p,radius:l}):s.rect(a,o,g,p),s.fill(),0!==n&&s.stroke()}s.restore()}(r.x(S),P,y),S=gt(w,S+g+u,x?S+k:this.right,t.rtl),function(t,e,i){Ne(s,i.text,t,e+m/2,h,{strikethrough:i.hidden,textAlign:r.textAlign(i.textAlign)})}(r.x(S),P,y),x)f.x+=k+c;else if("string"!=typeof y.text){const t=h.lineHeight;f.y+=_a(y,t)+c}else f.y+=_})),Ti(this.ctx,t.textDirection)}drawTitle(){const t=this.options,e=t.title,i=Si(e.font),s=ki(e.padding);if(!e.display)return;const n=Oi(t.rtl,this.left,this.width),o=this.ctx,a=e.position,r=i.size/2,l=s.top+r;let h,c=this.left,d=this.width;if(this.isHorizontal())d=Math.max(...this.lineWidths),h=this.top+l,c=ft(t.align,c,this.right-d);else{const e=this.columnSizes.reduce(((t,e)=>Math.max(t,e.height)),0);h=l+ft(t.align,this.top,this.bottom-e-t.labels.padding-this._computeTitleHeight())}const u=ft(a,c,c+d);o.textAlign=n.textAlign(ut(a)),o.textBaseline="middle",o.strokeStyle=e.color,o.fillStyle=e.color,o.font=i.string,Ne(o,e.text,u,h,i)}_computeTitleHeight(){const t=this.options.title,e=Si(t.font),i=ki(t.padding);return t.display?e.lineHeight+i.height:0}_getLegendItemAt(t,e){let i,s,n;if(tt(t,this.left,this.right)&&tt(e,this.top,this.bottom))for(n=this.legendHitBoxes,i=0;i<n.length;++i)if(s=n[i],tt(t,s.left,s.left+s.width)&&tt(e,s.top,s.top+s.height))return this.legendItems[i];return null}handleEvent(t){const e=this.options;if(!function(t,e){if(("mousemove"===t||"mouseout"===t)&&(e.onHover||e.onLeave))return!0;if(e.onClick&&("click"===t||"mouseup"===t))return!0;return!1}(t.type,e))return;const i=this._getLegendItemAt(t.x,t.y);if("mousemove"===t.type||"mouseout"===t.type){const o=this._hoveredItem,a=(n=i,null!==(s=o)&&null!==n&&s.datasetIndex===n.datasetIndex&&s.index===n.index);o&&!a&&d(e.onLeave,[t,o,this],this),this._hoveredItem=i,i&&!a&&d(e.onHover,[t,i,this],this)}else i&&d(e.onClick,[t,i,this],this);var s,n}}function _a(t,e){return e*(t.text?t.text.length:0)}var ya={id:"legend",_element:ba,start(t,e,i){const s=t.legend=new ba({ctx:t.ctx,options:i,chart:t});as.configure(t,s,i),as.addBox(t,s)},stop(t){as.removeBox(t,t.legend),delete t.legend},beforeUpdate(t,e,i){const s=t.legend;as.configure(t,s,i),s.options=i},afterUpdate(t){const e=t.legend;e.buildLabels(),e.adjustHitBoxes()},afterEvent(t,e){e.replay||t.legend.handleEvent(e.event)},defaults:{display:!0,position:"top",align:"center",fullSize:!0,reverse:!1,weight:1e3,onClick(t,e,i){const s=e.datasetIndex,n=i.chart;n.isDatasetVisible(s)?(n.hide(s),e.hidden=!0):(n.show(s),e.hidden=!1)},onHover:null,onLeave:null,labels:{color:t=>t.chart.options.color,boxWidth:40,padding:10,generateLabels(t){const e=t.data.datasets,{labels:{usePointStyle:i,pointStyle:s,textAlign:n,color:o,useBorderRadius:a,borderRadius:r}}=t.legend.options;return t._getSortedDatasetMetas().map((t=>{const l=t.controller.getStyle(i?0:void 0),h=ki(l.borderWidth);return{text:e[t.index].label,fillStyle:l.backgroundColor,fontColor:o,hidden:!t.visible,lineCap:l.borderCapStyle,lineDash:l.borderDash,lineDashOffset:l.borderDashOffset,lineJoin:l.borderJoinStyle,lineWidth:(h.width+h.height)/4,strokeStyle:l.borderColor,pointStyle:s||l.pointStyle,rotation:l.rotation,textAlign:n||l.textAlign,borderRadius:a&&(r||l.borderRadius),datasetIndex:t.index}}),this)}},title:{color:t=>t.chart.options.color,display:!1,position:"center",text:""}},descriptors:{_scriptable:t=>!t.startsWith("on"),labels:{_scriptable:t=>!["generateLabels","filter","sort"].includes(t)}}};class va extends Hs{constructor(t){super(),this.chart=t.chart,this.options=t.options,this.ctx=t.ctx,this._padding=void 0,this.top=void 0,this.bottom=void 0,this.left=void 0,this.right=void 0,this.width=void 0,this.height=void 0,this.position=void 0,this.weight=void 0,this.fullSize=void 0}update(t,e){const i=this.options;if(this.left=0,this.top=0,!i.display)return void(this.width=this.height=this.right=this.bottom=0);this.width=this.right=t,this.height=this.bottom=e;const s=n(i.text)?i.text.length:1;this._padding=ki(i.padding);const o=s*Si(i.font).lineHeight+this._padding.height;this.isHorizontal()?this.height=o:this.width=o}isHorizontal(){const t=this.options.position;return"top"===t||"bottom"===t}_drawArgs(t){const{top:e,left:i,bottom:s,right:n,options:o}=this,a=o.align;let r,l,h,c=0;return this.isHorizontal()?(l=ft(a,i,n),h=e+t,r=n-i):("left"===o.position?(l=i+t,h=ft(a,s,e),c=-.5*C):(l=n-t,h=ft(a,e,s),c=.5*C),r=s-e),{titleX:l,titleY:h,maxWidth:r,rotation:c}}draw(){const t=this.ctx,e=this.options;if(!e.display)return;const i=Si(e.font),s=i.lineHeight/2+this._padding.top,{titleX:n,titleY:o,maxWidth:a,rotation:r}=this._drawArgs(s);Ne(t,e.text,0,0,i,{color:e.color,maxWidth:a,rotation:r,textAlign:ut(e.align),textBaseline:"middle",translation:[n,o]})}}var Ma={id:"title",_element:va,start(t,e,i){!function(t,e){const i=new va({ctx:t.ctx,options:e,chart:t});as.configure(t,i,e),as.addBox(t,i),t.titleBlock=i}(t,i)},stop(t){const e=t.titleBlock;as.removeBox(t,e),delete t.titleBlock},beforeUpdate(t,e,i){const s=t.titleBlock;as.configure(t,s,i),s.options=i},defaults:{align:"center",display:!1,font:{weight:"bold"},fullSize:!0,padding:10,position:"top",text:"",weight:2e3},defaultRoutes:{color:"color"},descriptors:{_scriptable:!0,_indexable:!1}};const wa=new WeakMap;var ka={id:"subtitle",start(t,e,i){const s=new va({ctx:t.ctx,options:i,chart:t});as.configure(t,s,i),as.addBox(t,s),wa.set(t,s)},stop(t){as.removeBox(t,wa.get(t)),wa.delete(t)},beforeUpdate(t,e,i){const s=wa.get(t);as.configure(t,s,i),s.options=i},defaults:{align:"center",display:!1,font:{weight:"normal"},fullSize:!0,padding:0,position:"top",text:"",weight:1500},defaultRoutes:{color:"color"},descriptors:{_scriptable:!0,_indexable:!1}};const Sa={average(t){if(!t.length)return!1;let e,i,s=new Set,n=0,o=0;for(e=0,i=t.length;e<i;++e){const i=t[e].element;if(i&&i.hasValue()){const t=i.tooltipPosition();s.add(t.x),n+=t.y,++o}}return{x:[...s].reduce(((t,e)=>t+e))/s.size,y:n/o}},nearest(t,e){if(!t.length)return!1;let i,s,n,o=e.x,a=e.y,r=Number.POSITIVE_INFINITY;for(i=0,s=t.length;i<s;++i){const s=t[i].element;if(s&&s.hasValue()){const t=q(e,s.getCenterPoint());t<r&&(r=t,n=s)}}if(n){const t=n.tooltipPosition();o=t.x,a=t.y}return{x:o,y:a}}};function Pa(t,e){return e&&(n(e)?Array.prototype.push.apply(t,e):t.push(e)),t}function Da(t){return("string"==typeof t||t instanceof String)&&t.indexOf("\n")>-1?t.split("\n"):t}function Ca(t,e){const{element:i,datasetIndex:s,index:n}=e,o=t.getDatasetMeta(s).controller,{label:a,value:r}=o.getLabelAndValue(n);return{chart:t,label:a,parsed:o.getParsed(n),raw:t.data.datasets[s].data[n],formattedValue:r,dataset:o.getDataset(),dataIndex:n,datasetIndex:s,element:i}}function Oa(t,e){const i=t.chart.ctx,{body:s,footer:n,title:o}=t,{boxWidth:a,boxHeight:r}=e,l=Si(e.bodyFont),h=Si(e.titleFont),c=Si(e.footerFont),d=o.length,f=n.length,g=s.length,p=ki(e.padding);let m=p.height,x=0,b=s.reduce(((t,e)=>t+e.before.length+e.lines.length+e.after.length),0);if(b+=t.beforeBody.length+t.afterBody.length,d&&(m+=d*h.lineHeight+(d-1)*e.titleSpacing+e.titleMarginBottom),b){m+=g*(e.displayColors?Math.max(r,l.lineHeight):l.lineHeight)+(b-g)*l.lineHeight+(b-1)*e.bodySpacing}f&&(m+=e.footerMarginTop+f*c.lineHeight+(f-1)*e.footerSpacing);let _=0;const y=function(t){x=Math.max(x,i.measureText(t).width+_)};return i.save(),i.font=h.string,u(t.title,y),i.font=l.string,u(t.beforeBody.concat(t.afterBody),y),_=e.displayColors?a+2+e.boxPadding:0,u(s,(t=>{u(t.before,y),u(t.lines,y),u(t.after,y)})),_=0,i.font=c.string,u(t.footer,y),i.restore(),x+=p.width,{width:x,height:m}}function Aa(t,e,i,s){const{x:n,width:o}=i,{width:a,chartArea:{left:r,right:l}}=t;let h="center";return"center"===s?h=n<=(r+l)/2?"left":"right":n<=o/2?h="left":n>=a-o/2&&(h="right"),function(t,e,i,s){const{x:n,width:o}=s,a=i.caretSize+i.caretPadding;return"left"===t&&n+o+a>e.width||"right"===t&&n-o-a<0||void 0}(h,t,e,i)&&(h="center"),h}function Ta(t,e,i){const s=i.yAlign||e.yAlign||function(t,e){const{y:i,height:s}=e;return i<s/2?"top":i>t.height-s/2?"bottom":"center"}(t,i);return{xAlign:i.xAlign||e.xAlign||Aa(t,e,i,s),yAlign:s}}function La(t,e,i,s){const{caretSize:n,caretPadding:o,cornerRadius:a}=t,{xAlign:r,yAlign:l}=i,h=n+o,{topLeft:c,topRight:d,bottomLeft:u,bottomRight:f}=wi(a);let g=function(t,e){let{x:i,width:s}=t;return"right"===e?i-=s:"center"===e&&(i-=s/2),i}(e,r);const p=function(t,e,i){let{y:s,height:n}=t;return"top"===e?s+=i:s-="bottom"===e?n+i:n/2,s}(e,l,h);return"center"===l?"left"===r?g+=h:"right"===r&&(g-=h):"left"===r?g-=Math.max(c,u)+n:"right"===r&&(g+=Math.max(d,f)+n),{x:J(g,0,s.width-e.width),y:J(p,0,s.height-e.height)}}function Ea(t,e,i){const s=ki(i.padding);return"center"===e?t.x+t.width/2:"right"===e?t.x+t.width-s.right:t.x+s.left}function Ra(t){return Pa([],Da(t))}function Ia(t,e){const i=e&&e.dataset&&e.dataset.tooltip&&e.dataset.tooltip.callbacks;return i?t.override(i):t}const za={beforeTitle:e,title(t){if(t.length>0){const e=t[0],i=e.chart.data.labels,s=i?i.length:0;if(this&&this.options&&"dataset"===this.options.mode)return e.dataset.label||"";if(e.label)return e.label;if(s>0&&e.dataIndex<s)return i[e.dataIndex]}return""},afterTitle:e,beforeBody:e,beforeLabel:e,label(t){if(this&&this.options&&"dataset"===this.options.mode)return t.label+": "+t.formattedValue||t.formattedValue;let e=t.dataset.label||"";e&&(e+=": ");const i=t.formattedValue;return s(i)||(e+=i),e},labelColor(t){const e=t.chart.getDatasetMeta(t.datasetIndex).controller.getStyle(t.dataIndex);return{borderColor:e.borderColor,backgroundColor:e.backgroundColor,borderWidth:e.borderWidth,borderDash:e.borderDash,borderDashOffset:e.borderDashOffset,borderRadius:0}},labelTextColor(){return this.options.bodyColor},labelPointStyle(t){const e=t.chart.getDatasetMeta(t.datasetIndex).controller.getStyle(t.dataIndex);return{pointStyle:e.pointStyle,rotation:e.rotation}},afterLabel:e,afterBody:e,beforeFooter:e,footer:e,afterFooter:e};function Fa(t,e,i,s){const n=t[e].call(i,s);return void 0===n?za[e].call(i,s):n}class Va extends Hs{static positioners=Sa;constructor(t){super(),this.opacity=0,this._active=[],this._eventPosition=void 0,this._size=void 0,this._cachedAnimations=void 0,this._tooltipItems=[],this.$animations=void 0,this.$context=void 0,this.chart=t.chart,this.options=t.options,this.dataPoints=void 0,this.title=void 0,this.beforeBody=void 0,this.body=void 0,this.afterBody=void 0,this.footer=void 0,this.xAlign=void 0,this.yAlign=void 0,this.x=void 0,this.y=void 0,this.height=void 0,this.width=void 0,this.caretX=void 0,this.caretY=void 0,this.labelColors=void 0,this.labelPointStyles=void 0,this.labelTextColors=void 0}initialize(t){this.options=t,this._cachedAnimations=void 0,this.$context=void 0}_resolveAnimations(){const t=this._cachedAnimations;if(t)return t;const e=this.chart,i=this.options.setContext(this.getContext()),s=i.enabled&&e.options.animation&&i.animations,n=new Os(this.chart,s);return s._cacheable&&(this._cachedAnimations=Object.freeze(n)),n}getContext(){return this.$context||(this.$context=(t=this.chart.getContext(),e=this,i=this._tooltipItems,Ci(t,{tooltip:e,tooltipItems:i,type:"tooltip"})));var t,e,i}getTitle(t,e){const{callbacks:i}=e,s=Fa(i,"beforeTitle",this,t),n=Fa(i,"title",this,t),o=Fa(i,"afterTitle",this,t);let a=[];return a=Pa(a,Da(s)),a=Pa(a,Da(n)),a=Pa(a,Da(o)),a}getBeforeBody(t,e){return Ra(Fa(e.callbacks,"beforeBody",this,t))}getBody(t,e){const{callbacks:i}=e,s=[];return u(t,(t=>{const e={before:[],lines:[],after:[]},n=Ia(i,t);Pa(e.before,Da(Fa(n,"beforeLabel",this,t))),Pa(e.lines,Fa(n,"label",this,t)),Pa(e.after,Da(Fa(n,"afterLabel",this,t))),s.push(e)})),s}getAfterBody(t,e){return Ra(Fa(e.callbacks,"afterBody",this,t))}getFooter(t,e){const{callbacks:i}=e,s=Fa(i,"beforeFooter",this,t),n=Fa(i,"footer",this,t),o=Fa(i,"afterFooter",this,t);let a=[];return a=Pa(a,Da(s)),a=Pa(a,Da(n)),a=Pa(a,Da(o)),a}_createItems(t){const e=this._active,i=this.chart.data,s=[],n=[],o=[];let a,r,l=[];for(a=0,r=e.length;a<r;++a)l.push(Ca(this.chart,e[a]));return t.filter&&(l=l.filter(((e,s,n)=>t.filter(e,s,n,i)))),t.itemSort&&(l=l.sort(((e,s)=>t.itemSort(e,s,i)))),u(l,(e=>{const i=Ia(t.callbacks,e);s.push(Fa(i,"labelColor",this,e)),n.push(Fa(i,"labelPointStyle",this,e)),o.push(Fa(i,"labelTextColor",this,e))})),this.labelColors=s,this.labelPointStyles=n,this.labelTextColors=o,this.dataPoints=l,l}update(t,e){const i=this.options.setContext(this.getContext()),s=this._active;let n,o=[];if(s.length){const t=Sa[i.position].call(this,s,this._eventPosition);o=this._createItems(i),this.title=this.getTitle(o,i),this.beforeBody=this.getBeforeBody(o,i),this.body=this.getBody(o,i),this.afterBody=this.getAfterBody(o,i),this.footer=this.getFooter(o,i);const e=this._size=Oa(this,i),a=Object.assign({},t,e),r=Ta(this.chart,i,a),l=La(i,a,r,this.chart);this.xAlign=r.xAlign,this.yAlign=r.yAlign,n={opacity:1,x:l.x,y:l.y,width:e.width,height:e.height,caretX:t.x,caretY:t.y}}else 0!==this.opacity&&(n={opacity:0});this._tooltipItems=o,this.$context=void 0,n&&this._resolveAnimations().update(this,n),t&&i.external&&i.external.call(this,{chart:this.chart,tooltip:this,replay:e})}drawCaret(t,e,i,s){const n=this.getCaretPosition(t,i,s);e.lineTo(n.x1,n.y1),e.lineTo(n.x2,n.y2),e.lineTo(n.x3,n.y3)}getCaretPosition(t,e,i){const{xAlign:s,yAlign:n}=this,{caretSize:o,cornerRadius:a}=i,{topLeft:r,topRight:l,bottomLeft:h,bottomRight:c}=wi(a),{x:d,y:u}=t,{width:f,height:g}=e;let p,m,x,b,_,y;return"center"===n?(_=u+g/2,"left"===s?(p=d,m=p-o,b=_+o,y=_-o):(p=d+f,m=p+o,b=_-o,y=_+o),x=p):(m="left"===s?d+Math.max(r,h)+o:"right"===s?d+f-Math.max(l,c)-o:this.caretX,"top"===n?(b=u,_=b-o,p=m-o,x=m+o):(b=u+g,_=b+o,p=m+o,x=m-o),y=b),{x1:p,x2:m,x3:x,y1:b,y2:_,y3:y}}drawTitle(t,e,i){const s=this.title,n=s.length;let o,a,r;if(n){const l=Oi(i.rtl,this.x,this.width);for(t.x=Ea(this,i.titleAlign,i),e.textAlign=l.textAlign(i.titleAlign),e.textBaseline="middle",o=Si(i.titleFont),a=i.titleSpacing,e.fillStyle=i.titleColor,e.font=o.string,r=0;r<n;++r)e.fillText(s[r],l.x(t.x),t.y+o.lineHeight/2),t.y+=o.lineHeight+a,r+1===n&&(t.y+=i.titleMarginBottom-a)}}_drawColorBox(t,e,i,s,n){const a=this.labelColors[i],r=this.labelPointStyles[i],{boxHeight:l,boxWidth:h}=n,c=Si(n.bodyFont),d=Ea(this,"left",n),u=s.x(d),f=l<c.lineHeight?(c.lineHeight-l)/2:0,g=e.y+f;if(n.usePointStyle){const e={radius:Math.min(h,l)/2,pointStyle:r.pointStyle,rotation:r.rotation,borderWidth:1},i=s.leftForLtr(u,h)+h/2,o=g+l/2;t.strokeStyle=n.multiKeyBackground,t.fillStyle=n.multiKeyBackground,Le(t,e,i,o),t.strokeStyle=a.borderColor,t.fillStyle=a.backgroundColor,Le(t,e,i,o)}else{t.lineWidth=o(a.borderWidth)?Math.max(...Object.values(a.borderWidth)):a.borderWidth||1,t.strokeStyle=a.borderColor,t.setLineDash(a.borderDash||[]),t.lineDashOffset=a.borderDashOffset||0;const e=s.leftForLtr(u,h),i=s.leftForLtr(s.xPlus(u,1),h-2),r=wi(a.borderRadius);Object.values(r).some((t=>0!==t))?(t.beginPath(),t.fillStyle=n.multiKeyBackground,He(t,{x:e,y:g,w:h,h:l,radius:r}),t.fill(),t.stroke(),t.fillStyle=a.backgroundColor,t.beginPath(),He(t,{x:i,y:g+1,w:h-2,h:l-2,radius:r}),t.fill()):(t.fillStyle=n.multiKeyBackground,t.fillRect(e,g,h,l),t.strokeRect(e,g,h,l),t.fillStyle=a.backgroundColor,t.fillRect(i,g+1,h-2,l-2))}t.fillStyle=this.labelTextColors[i]}drawBody(t,e,i){const{body:s}=this,{bodySpacing:n,bodyAlign:o,displayColors:a,boxHeight:r,boxWidth:l,boxPadding:h}=i,c=Si(i.bodyFont);let d=c.lineHeight,f=0;const g=Oi(i.rtl,this.x,this.width),p=function(i){e.fillText(i,g.x(t.x+f),t.y+d/2),t.y+=d+n},m=g.textAlign(o);let x,b,_,y,v,M,w;for(e.textAlign=o,e.textBaseline="middle",e.font=c.string,t.x=Ea(this,m,i),e.fillStyle=i.bodyColor,u(this.beforeBody,p),f=a&&"right"!==m?"center"===o?l/2+h:l+2+h:0,y=0,M=s.length;y<M;++y){for(x=s[y],b=this.labelTextColors[y],e.fillStyle=b,u(x.before,p),_=x.lines,a&&_.length&&(this._drawColorBox(e,t,y,g,i),d=Math.max(c.lineHeight,r)),v=0,w=_.length;v<w;++v)p(_[v]),d=c.lineHeight;u(x.after,p)}f=0,d=c.lineHeight,u(this.afterBody,p),t.y-=n}drawFooter(t,e,i){const s=this.footer,n=s.length;let o,a;if(n){const r=Oi(i.rtl,this.x,this.width);for(t.x=Ea(this,i.footerAlign,i),t.y+=i.footerMarginTop,e.textAlign=r.textAlign(i.footerAlign),e.textBaseline="middle",o=Si(i.footerFont),e.fillStyle=i.footerColor,e.font=o.string,a=0;a<n;++a)e.fillText(s[a],r.x(t.x),t.y+o.lineHeight/2),t.y+=o.lineHeight+i.footerSpacing}}drawBackground(t,e,i,s){const{xAlign:n,yAlign:o}=this,{x:a,y:r}=t,{width:l,height:h}=i,{topLeft:c,topRight:d,bottomLeft:u,bottomRight:f}=wi(s.cornerRadius);e.fillStyle=s.backgroundColor,e.strokeStyle=s.borderColor,e.lineWidth=s.borderWidth,e.beginPath(),e.moveTo(a+c,r),"top"===o&&this.drawCaret(t,e,i,s),e.lineTo(a+l-d,r),e.quadraticCurveTo(a+l,r,a+l,r+d),"center"===o&&"right"===n&&this.drawCaret(t,e,i,s),e.lineTo(a+l,r+h-f),e.quadraticCurveTo(a+l,r+h,a+l-f,r+h),"bottom"===o&&this.drawCaret(t,e,i,s),e.lineTo(a+u,r+h),e.quadraticCurveTo(a,r+h,a,r+h-u),"center"===o&&"left"===n&&this.drawCaret(t,e,i,s),e.lineTo(a,r+c),e.quadraticCurveTo(a,r,a+c,r),e.closePath(),e.fill(),s.borderWidth>0&&e.stroke()}_updateAnimationTarget(t){const e=this.chart,i=this.$animations,s=i&&i.x,n=i&&i.y;if(s||n){const i=Sa[t.position].call(this,this._active,this._eventPosition);if(!i)return;const o=this._size=Oa(this,t),a=Object.assign({},i,this._size),r=Ta(e,t,a),l=La(t,a,r,e);s._to===l.x&&n._to===l.y||(this.xAlign=r.xAlign,this.yAlign=r.yAlign,this.width=o.width,this.height=o.height,this.caretX=i.x,this.caretY=i.y,this._resolveAnimations().update(this,l))}}_willRender(){return!!this.opacity}draw(t){const e=this.options.setContext(this.getContext());let i=this.opacity;if(!i)return;this._updateAnimationTarget(e);const s={width:this.width,height:this.height},n={x:this.x,y:this.y};i=Math.abs(i)<.001?0:i;const o=ki(e.padding),a=this.title.length||this.beforeBody.length||this.body.length||this.afterBody.length||this.footer.length;e.enabled&&a&&(t.save(),t.globalAlpha=i,this.drawBackground(n,t,s,e),Ai(t,e.textDirection),n.y+=o.top,this.drawTitle(n,t,e),this.drawBody(n,t,e),this.drawFooter(n,t,e),Ti(t,e.textDirection),t.restore())}getActiveElements(){return this._active||[]}setActiveElements(t,e){const i=this._active,s=t.map((({datasetIndex:t,index:e})=>{const i=this.chart.getDatasetMeta(t);if(!i)throw new Error("Cannot find a dataset at index "+t);return{datasetIndex:t,element:i.data[e],index:e}})),n=!f(i,s),o=this._positionChanged(s,e);(n||o)&&(this._active=s,this._eventPosition=e,this._ignoreReplayEvents=!0,this.update(!0))}handleEvent(t,e,i=!0){if(e&&this._ignoreReplayEvents)return!1;this._ignoreReplayEvents=!1;const s=this.options,n=this._active||[],o=this._getActiveElements(t,n,e,i),a=this._positionChanged(o,t),r=e||!f(o,n)||a;return r&&(this._active=o,(s.enabled||s.external)&&(this._eventPosition={x:t.x,y:t.y},this.update(!0,e))),r}_getActiveElements(t,e,i,s){const n=this.options;if("mouseout"===t.type)return[];if(!s)return e.filter((t=>this.chart.data.datasets[t.datasetIndex]&&void 0!==this.chart.getDatasetMeta(t.datasetIndex).controller.getParsed(t.index)));const o=this.chart.getElementsAtEventForMode(t,n.mode,n,i);return n.reverse&&o.reverse(),o}_positionChanged(t,e){const{caretX:i,caretY:s,options:n}=this,o=Sa[n.position].call(this,t,e);return!1!==o&&(i!==o.x||s!==o.y)}}var Ba={id:"tooltip",_element:Va,positioners:Sa,afterInit(t,e,i){i&&(t.tooltip=new Va({chart:t,options:i}))},beforeUpdate(t,e,i){t.tooltip&&t.tooltip.initialize(i)},reset(t,e,i){t.tooltip&&t.tooltip.initialize(i)},afterDraw(t){const e=t.tooltip;if(e&&e._willRender()){const i={tooltip:e};if(!1===t.notifyPlugins("beforeTooltipDraw",{...i,cancelable:!0}))return;e.draw(t.ctx),t.notifyPlugins("afterTooltipDraw",i)}},afterEvent(t,e){if(t.tooltip){const i=e.replay;t.tooltip.handleEvent(e.event,i,e.inChartArea)&&(e.changed=!0)}},defaults:{enabled:!0,external:null,position:"average",backgroundColor:"rgba(0,0,0,0.8)",titleColor:"#fff",titleFont:{weight:"bold"},titleSpacing:2,titleMarginBottom:6,titleAlign:"left",bodyColor:"#fff",bodySpacing:2,bodyFont:{},bodyAlign:"left",footerColor:"#fff",footerSpacing:2,footerMarginTop:6,footerFont:{weight:"bold"},footerAlign:"left",padding:6,caretPadding:2,caretSize:5,cornerRadius:6,boxHeight:(t,e)=>e.bodyFont.size,boxWidth:(t,e)=>e.bodyFont.size,multiKeyBackground:"#fff",displayColors:!0,boxPadding:0,borderColor:"rgba(0,0,0,0)",borderWidth:0,animation:{duration:400,easing:"easeOutQuart"},animations:{numbers:{type:"number",properties:["x","y","width","height","caretX","caretY"]},opacity:{easing:"linear",duration:200}},callbacks:za},defaultRoutes:{bodyFont:"font",footerFont:"font",titleFont:"font"},descriptors:{_scriptable:t=>"filter"!==t&&"itemSort"!==t&&"external"!==t,_indexable:!1,callbacks:{_scriptable:!1,_indexable:!1},animation:{_fallback:!1},animations:{_fallback:"animation"}},additionalOptionScopes:["interaction"]};return An.register(Yn,jo,fo,t),An.helpers={...Wi},An._adapters=Rn,An.Animation=Cs,An.Animations=Os,An.animator=bt,An.controllers=en.controllers.items,An.DatasetController=Ns,An.Element=Hs,An.elements=fo,An.Interaction=Xi,An.layouts=as,An.platforms=Ss,An.Scale=Js,An.Ticks=ae,Object.assign(An,Yn,jo,fo,t,Ss),An.Chart=An,"undefined"!=typeof window&&(window.Chart=An),An}));
//# sourceMappingURL=chart.umd.js.map

</script>
<style>
/* Leaflet v1.9.4 CSS — embebido offline */
/* required styles */

.leaflet-pane,
.leaflet-tile,
.leaflet-marker-icon,
.leaflet-marker-shadow,
.leaflet-tile-container,
.leaflet-pane > svg,
.leaflet-pane > canvas,
.leaflet-zoom-box,
.leaflet-image-layer,
.leaflet-layer {
	position: absolute;
	left: 0;
	top: 0;
	}
.leaflet-container {
	overflow: hidden;
	}
.leaflet-tile,
.leaflet-marker-icon,
.leaflet-marker-shadow {
	-webkit-user-select: none;
	   -moz-user-select: none;
	        user-select: none;
	  -webkit-user-drag: none;
	}
/* Prevents IE11 from highlighting tiles in blue */
.leaflet-tile::selection {
	background: transparent;
}
/* Safari renders non-retina tile on retina better with this, but Chrome is worse */
.leaflet-safari .leaflet-tile {
	image-rendering: -webkit-optimize-contrast;
	}
/* hack that prevents hw layers "stretching" when loading new tiles */
.leaflet-safari .leaflet-tile-container {
	width: 1600px;
	height: 1600px;
	-webkit-transform-origin: 0 0;
	}
.leaflet-marker-icon,
.leaflet-marker-shadow {
	display: block;
	}
/* .leaflet-container svg: reset svg max-width decleration shipped in Joomla! (joomla.org) 3.x */
/* .leaflet-container img: map is broken in FF if you have max-width: 100% on tiles */
.leaflet-container .leaflet-overlay-pane svg {
	max-width: none !important;
	max-height: none !important;
	}
.leaflet-container .leaflet-marker-pane img,
.leaflet-container .leaflet-shadow-pane img,
.leaflet-container .leaflet-tile-pane img,
.leaflet-container img.leaflet-image-layer,
.leaflet-container .leaflet-tile {
	max-width: none !important;
	max-height: none !important;
	width: auto;
	padding: 0;
	}

.leaflet-container img.leaflet-tile {
	/* See: https://bugs.chromium.org/p/chromium/issues/detail?id=600120 */
	mix-blend-mode: plus-lighter;
}

.leaflet-container.leaflet-touch-zoom {
	-ms-touch-action: pan-x pan-y;
	touch-action: pan-x pan-y;
	}
.leaflet-container.leaflet-touch-drag {
	-ms-touch-action: pinch-zoom;
	/* Fallback for FF which doesn't support pinch-zoom */
	touch-action: none;
	touch-action: pinch-zoom;
}
.leaflet-container.leaflet-touch-drag.leaflet-touch-zoom {
	-ms-touch-action: none;
	touch-action: none;
}
.leaflet-container {
	-webkit-tap-highlight-color: transparent;
}
.leaflet-container a {
	-webkit-tap-highlight-color: rgba(51, 181, 229, 0.4);
}
.leaflet-tile {
	filter: inherit;
	visibility: hidden;
	}
.leaflet-tile-loaded {
	visibility: inherit;
	}
.leaflet-zoom-box {
	width: 0;
	height: 0;
	-moz-box-sizing: border-box;
	     box-sizing: border-box;
	z-index: 800;
	}
/* workaround for https://bugzilla.mozilla.org/show_bug.cgi?id=888319 */
.leaflet-overlay-pane svg {
	-moz-user-select: none;
	}

.leaflet-pane         { z-index: 400; }

.leaflet-tile-pane    { z-index: 200; }
.leaflet-overlay-pane { z-index: 400; }
.leaflet-shadow-pane  { z-index: 500; }
.leaflet-marker-pane  { z-index: 600; }
.leaflet-tooltip-pane   { z-index: 650; }
.leaflet-popup-pane   { z-index: 700; }

.leaflet-map-pane canvas { z-index: 100; }
.leaflet-map-pane svg    { z-index: 200; }

.leaflet-vml-shape {
	width: 1px;
	height: 1px;
	}
.lvml {
	behavior: url(#default#VML);
	display: inline-block;
	position: absolute;
	}


/* control positioning */

.leaflet-control {
	position: relative;
	z-index: 800;
	pointer-events: visiblePainted; /* IE 9-10 doesn't have auto */
	pointer-events: auto;
	}
.leaflet-top,
.leaflet-bottom {
	position: absolute;
	z-index: 1000;
	pointer-events: none;
	}
.leaflet-top {
	top: 0;
	}
.leaflet-right {
	right: 0;
	}
.leaflet-bottom {
	bottom: 0;
	}
.leaflet-left {
	left: 0;
	}
.leaflet-control {
	float: left;
	clear: both;
	}
.leaflet-right .leaflet-control {
	float: right;
	}
.leaflet-top .leaflet-control {
	margin-top: 10px;
	}
.leaflet-bottom .leaflet-control {
	margin-bottom: 10px;
	}
.leaflet-left .leaflet-control {
	margin-left: 10px;
	}
.leaflet-right .leaflet-control {
	margin-right: 10px;
	}


/* zoom and fade animations */

.leaflet-fade-anim .leaflet-popup {
	opacity: 0;
	-webkit-transition: opacity 0.2s linear;
	   -moz-transition: opacity 0.2s linear;
	        transition: opacity 0.2s linear;
	}
.leaflet-fade-anim .leaflet-map-pane .leaflet-popup {
	opacity: 1;
	}
.leaflet-zoom-animated {
	-webkit-transform-origin: 0 0;
	    -ms-transform-origin: 0 0;
	        transform-origin: 0 0;
	}
svg.leaflet-zoom-animated {
	will-change: transform;
}

.leaflet-zoom-anim .leaflet-zoom-animated {
	-webkit-transition: -webkit-transform 0.25s cubic-bezier(0,0,0.25,1);
	   -moz-transition:    -moz-transform 0.25s cubic-bezier(0,0,0.25,1);
	        transition:         transform 0.25s cubic-bezier(0,0,0.25,1);
	}
.leaflet-zoom-anim .leaflet-tile,
.leaflet-pan-anim .leaflet-tile {
	-webkit-transition: none;
	   -moz-transition: none;
	        transition: none;
	}

.leaflet-zoom-anim .leaflet-zoom-hide {
	visibility: hidden;
	}


/* cursors */

.leaflet-interactive {
	cursor: pointer;
	}
.leaflet-grab {
	cursor: -webkit-grab;
	cursor:    -moz-grab;
	cursor:         grab;
	}
.leaflet-crosshair,
.leaflet-crosshair .leaflet-interactive {
	cursor: crosshair;
	}
.leaflet-popup-pane,
.leaflet-control {
	cursor: auto;
	}
.leaflet-dragging .leaflet-grab,
.leaflet-dragging .leaflet-grab .leaflet-interactive,
.leaflet-dragging .leaflet-marker-draggable {
	cursor: move;
	cursor: -webkit-grabbing;
	cursor:    -moz-grabbing;
	cursor:         grabbing;
	}

/* marker & overlays interactivity */
.leaflet-marker-icon,
.leaflet-marker-shadow,
.leaflet-image-layer,
.leaflet-pane > svg path,
.leaflet-tile-container {
	pointer-events: none;
	}

.leaflet-marker-icon.leaflet-interactive,
.leaflet-image-layer.leaflet-interactive,
.leaflet-pane > svg path.leaflet-interactive,
svg.leaflet-image-layer.leaflet-interactive path {
	pointer-events: visiblePainted; /* IE 9-10 doesn't have auto */
	pointer-events: auto;
	}

/* visual tweaks */

.leaflet-container {
	background: #ddd;
	outline-offset: 1px;
	}
.leaflet-container a {
	color: #0078A8;
	}
.leaflet-zoom-box {
	border: 2px dotted #38f;
	background: rgba(255,255,255,0.5);
	}


/* general typography */
.leaflet-container {
	font-family: "Helvetica Neue", Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-size: 0.75rem;
	line-height: 1.5;
	}


/* general toolbar styles */

.leaflet-bar {
	box-shadow: 0 1px 5px rgba(0,0,0,0.65);
	border-radius: 4px;
	}
.leaflet-bar a {
	background-color: #fff;
	border-bottom: 1px solid #ccc;
	width: 26px;
	height: 26px;
	line-height: 26px;
	display: block;
	text-align: center;
	text-decoration: none;
	color: black;
	}
.leaflet-bar a,
.leaflet-control-layers-toggle {
	background-position: 50% 50%;
	background-repeat: no-repeat;
	display: block;
	}
.leaflet-bar a:hover,
.leaflet-bar a:focus {
	background-color: #f4f4f4;
	}
.leaflet-bar a:first-child {
	border-top-left-radius: 4px;
	border-top-right-radius: 4px;
	}
.leaflet-bar a:last-child {
	border-bottom-left-radius: 4px;
	border-bottom-right-radius: 4px;
	border-bottom: none;
	}
.leaflet-bar a.leaflet-disabled {
	cursor: default;
	background-color: #f4f4f4;
	color: #bbb;
	}

.leaflet-touch .leaflet-bar a {
	width: 30px;
	height: 30px;
	line-height: 30px;
	}
.leaflet-touch .leaflet-bar a:first-child {
	border-top-left-radius: 2px;
	border-top-right-radius: 2px;
	}
.leaflet-touch .leaflet-bar a:last-child {
	border-bottom-left-radius: 2px;
	border-bottom-right-radius: 2px;
	}

/* zoom control */

.leaflet-control-zoom-in,
.leaflet-control-zoom-out {
	font: bold 18px 'Lucida Console', Monaco, monospace;
	text-indent: 1px;
	}

.leaflet-touch .leaflet-control-zoom-in, .leaflet-touch .leaflet-control-zoom-out  {
	font-size: 22px;
	}


/* layers control */

.leaflet-control-layers {
	box-shadow: 0 1px 5px rgba(0,0,0,0.4);
	background: #fff;
	border-radius: 5px;
	}
.leaflet-control-layers-toggle {
	background-image: url(images/layers.png);
	width: 36px;
	height: 36px;
	}
.leaflet-retina .leaflet-control-layers-toggle {
	background-image: url(images/layers-2x.png);
	background-size: 26px 26px;
	}
.leaflet-touch .leaflet-control-layers-toggle {
	width: 44px;
	height: 44px;
	}
.leaflet-control-layers .leaflet-control-layers-list,
.leaflet-control-layers-expanded .leaflet-control-layers-toggle {
	display: none;
	}
.leaflet-control-layers-expanded .leaflet-control-layers-list {
	display: block;
	position: relative;
	}
.leaflet-control-layers-expanded {
	padding: 6px 10px 6px 6px;
	color: #333;
	background: #fff;
	}
.leaflet-control-layers-scrollbar {
	overflow-y: scroll;
	overflow-x: hidden;
	padding-right: 5px;
	}
.leaflet-control-layers-selector {
	margin-top: 2px;
	position: relative;
	top: 1px;
	}
.leaflet-control-layers label {
	display: block;
	font-size: 13px;
	font-size: 1.08333em;
	}
.leaflet-control-layers-separator {
	height: 0;
	border-top: 1px solid #ddd;
	margin: 5px -10px 5px -6px;
	}

/* Default icon URLs */
.leaflet-default-icon-path { /* used only in path-guessing heuristic, see L.Icon.Default */
	background-image: url(images/marker-icon.png);
	}


/* attribution and scale controls */

.leaflet-container .leaflet-control-attribution {
	background: #fff;
	background: rgba(255, 255, 255, 0.8);
	margin: 0;
	}
.leaflet-control-attribution,
.leaflet-control-scale-line {
	padding: 0 5px;
	color: #333;
	line-height: 1.4;
	}
.leaflet-control-attribution a {
	text-decoration: none;
	}
.leaflet-control-attribution a:hover,
.leaflet-control-attribution a:focus {
	text-decoration: underline;
	}
.leaflet-attribution-flag {
	display: inline !important;
	vertical-align: baseline !important;
	width: 1em;
	height: 0.6669em;
	}
.leaflet-left .leaflet-control-scale {
	margin-left: 5px;
	}
.leaflet-bottom .leaflet-control-scale {
	margin-bottom: 5px;
	}
.leaflet-control-scale-line {
	border: 2px solid #777;
	border-top: none;
	line-height: 1.1;
	padding: 2px 5px 1px;
	white-space: nowrap;
	-moz-box-sizing: border-box;
	     box-sizing: border-box;
	background: rgba(255, 255, 255, 0.8);
	text-shadow: 1px 1px #fff;
	}
.leaflet-control-scale-line:not(:first-child) {
	border-top: 2px solid #777;
	border-bottom: none;
	margin-top: -2px;
	}
.leaflet-control-scale-line:not(:first-child):not(:last-child) {
	border-bottom: 2px solid #777;
	}

.leaflet-touch .leaflet-control-attribution,
.leaflet-touch .leaflet-control-layers,
.leaflet-touch .leaflet-bar {
	box-shadow: none;
	}
.leaflet-touch .leaflet-control-layers,
.leaflet-touch .leaflet-bar {
	border: 2px solid rgba(0,0,0,0.2);
	background-clip: padding-box;
	}


/* popup */

.leaflet-popup {
	position: absolute;
	text-align: center;
	margin-bottom: 20px;
	}
.leaflet-popup-content-wrapper {
	padding: 1px;
	text-align: left;
	border-radius: 12px;
	}
.leaflet-popup-content {
	margin: 13px 24px 13px 20px;
	line-height: 1.3;
	font-size: 13px;
	font-size: 1.08333em;
	min-height: 1px;
	}
.leaflet-popup-content p {
	margin: 17px 0;
	margin: 1.3em 0;
	}
.leaflet-popup-tip-container {
	width: 40px;
	height: 20px;
	position: absolute;
	left: 50%;
	margin-top: -1px;
	margin-left: -20px;
	overflow: hidden;
	pointer-events: none;
	}
.leaflet-popup-tip {
	width: 17px;
	height: 17px;
	padding: 1px;

	margin: -10px auto 0;
	pointer-events: auto;

	-webkit-transform: rotate(45deg);
	   -moz-transform: rotate(45deg);
	    -ms-transform: rotate(45deg);
	        transform: rotate(45deg);
	}
.leaflet-popup-content-wrapper,
.leaflet-popup-tip {
	background: white;
	color: #333;
	box-shadow: 0 3px 14px rgba(0,0,0,0.4);
	}
.leaflet-container a.leaflet-popup-close-button {
	position: absolute;
	top: 0;
	right: 0;
	border: none;
	text-align: center;
	width: 24px;
	height: 24px;
	font: 16px/24px Tahoma, Verdana, sans-serif;
	color: #757575;
	text-decoration: none;
	background: transparent;
	}
.leaflet-container a.leaflet-popup-close-button:hover,
.leaflet-container a.leaflet-popup-close-button:focus {
	color: #585858;
	}
.leaflet-popup-scrolled {
	overflow: auto;
	}

.leaflet-oldie .leaflet-popup-content-wrapper {
	-ms-zoom: 1;
	}
.leaflet-oldie .leaflet-popup-tip {
	width: 24px;
	margin: 0 auto;

	-ms-filter: "progid:DXImageTransform.Microsoft.Matrix(M11=0.70710678, M12=0.70710678, M21=-0.70710678, M22=0.70710678)";
	filter: progid:DXImageTransform.Microsoft.Matrix(M11=0.70710678, M12=0.70710678, M21=-0.70710678, M22=0.70710678);
	}

.leaflet-oldie .leaflet-control-zoom,
.leaflet-oldie .leaflet-control-layers,
.leaflet-oldie .leaflet-popup-content-wrapper,
.leaflet-oldie .leaflet-popup-tip {
	border: 1px solid #999;
	}


/* div icon */

.leaflet-div-icon {
	background: #fff;
	border: 1px solid #666;
	}


/* Tooltip */
/* Base styles for the element that has a tooltip */
.leaflet-tooltip {
	position: absolute;
	padding: 6px;
	background-color: #fff;
	border: 1px solid #fff;
	border-radius: 3px;
	color: #222;
	white-space: nowrap;
	-webkit-user-select: none;
	-moz-user-select: none;
	-ms-user-select: none;
	user-select: none;
	pointer-events: none;
	box-shadow: 0 1px 3px rgba(0,0,0,0.4);
	}
.leaflet-tooltip.leaflet-interactive {
	cursor: pointer;
	pointer-events: auto;
	}
.leaflet-tooltip-top:before,
.leaflet-tooltip-bottom:before,
.leaflet-tooltip-left:before,
.leaflet-tooltip-right:before {
	position: absolute;
	pointer-events: none;
	border: 6px solid transparent;
	background: transparent;
	content: "";
	}

/* Directions */

.leaflet-tooltip-bottom {
	margin-top: 6px;
}
.leaflet-tooltip-top {
	margin-top: -6px;
}
.leaflet-tooltip-bottom:before,
.leaflet-tooltip-top:before {
	left: 50%;
	margin-left: -6px;
	}
.leaflet-tooltip-top:before {
	bottom: 0;
	margin-bottom: -12px;
	border-top-color: #fff;
	}
.leaflet-tooltip-bottom:before {
	top: 0;
	margin-top: -12px;
	margin-left: -6px;
	border-bottom-color: #fff;
	}
.leaflet-tooltip-left {
	margin-left: -6px;
}
.leaflet-tooltip-right {
	margin-left: 6px;
}
.leaflet-tooltip-left:before,
.leaflet-tooltip-right:before {
	top: 50%;
	margin-top: -6px;
	}
.leaflet-tooltip-left:before {
	right: 0;
	margin-right: -12px;
	border-left-color: #fff;
	}
.leaflet-tooltip-right:before {
	left: 0;
	margin-left: -12px;
	border-right-color: #fff;
	}

/* Printing */

@media print {
	/* Prevent printers from removing background-images of controls. */
	.leaflet-control {
		-webkit-print-color-adjust: exact;
		print-color-adjust: exact;
		}
	}

</style>
<script>
/* Leaflet v1.9.4 JS — embebido offline */
/* @preserve
 * Leaflet 1.9.4, a JS library for interactive maps. https://leafletjs.com
 * (c) 2010-2023 Vladimir Agafonkin, (c) 2010-2011 CloudMade
 */
!function(t,e){"object"==typeof exports&&"undefined"!=typeof module?e(exports):"function"==typeof define&&define.amd?define(["exports"],e):e((t="undefined"!=typeof globalThis?globalThis:t||self).leaflet={})}(this,function(t){"use strict";function l(t){for(var e,i,n=1,o=arguments.length;n<o;n++)for(e in i=arguments[n])t[e]=i[e];return t}var R=Object.create||function(t){return N.prototype=t,new N};function N(){}function a(t,e){var i,n=Array.prototype.slice;return t.bind?t.bind.apply(t,n.call(arguments,1)):(i=n.call(arguments,2),function(){return t.apply(e,i.length?i.concat(n.call(arguments)):arguments)})}var D=0;function h(t){return"_leaflet_id"in t||(t._leaflet_id=++D),t._leaflet_id}function j(t,e,i){var n,o,s=function(){n=!1,o&&(r.apply(i,o),o=!1)},r=function(){n?o=arguments:(t.apply(i,arguments),setTimeout(s,e),n=!0)};return r}function H(t,e,i){var n=e[1],e=e[0],o=n-e;return t===n&&i?t:((t-e)%o+o)%o+e}function u(){return!1}function i(t,e){return!1===e?t:(e=Math.pow(10,void 0===e?6:e),Math.round(t*e)/e)}function W(t){return t.trim?t.trim():t.replace(/^\s+|\s+$/g,"")}function F(t){return W(t).split(/\s+/)}function c(t,e){for(var i in Object.prototype.hasOwnProperty.call(t,"options")||(t.options=t.options?R(t.options):{}),e)t.options[i]=e[i];return t.options}function U(t,e,i){var n,o=[];for(n in t)o.push(encodeURIComponent(i?n.toUpperCase():n)+"="+encodeURIComponent(t[n]));return(e&&-1!==e.indexOf("?")?"&":"?")+o.join("&")}var V=/\{ *([\w_ -]+) *\}/g;function q(t,i){return t.replace(V,function(t,e){e=i[e];if(void 0===e)throw new Error("No value provided for variable "+t);return e="function"==typeof e?e(i):e})}var d=Array.isArray||function(t){return"[object Array]"===Object.prototype.toString.call(t)};function G(t,e){for(var i=0;i<t.length;i++)if(t[i]===e)return i;return-1}var K="data:image/gif;base64,R0lGODlhAQABAAD/ACwAAAAAAQABAAACADs=";function Y(t){return window["webkit"+t]||window["moz"+t]||window["ms"+t]}var X=0;function J(t){var e=+new Date,i=Math.max(0,16-(e-X));return X=e+i,window.setTimeout(t,i)}var $=window.requestAnimationFrame||Y("RequestAnimationFrame")||J,Q=window.cancelAnimationFrame||Y("CancelAnimationFrame")||Y("CancelRequestAnimationFrame")||function(t){window.clearTimeout(t)};function x(t,e,i){if(!i||$!==J)return $.call(window,a(t,e));t.call(e)}function r(t){t&&Q.call(window,t)}var tt={__proto__:null,extend:l,create:R,bind:a,get lastId(){return D},stamp:h,throttle:j,wrapNum:H,falseFn:u,formatNum:i,trim:W,splitWords:F,setOptions:c,getParamString:U,template:q,isArray:d,indexOf:G,emptyImageUrl:K,requestFn:$,cancelFn:Q,requestAnimFrame:x,cancelAnimFrame:r};function et(){}et.extend=function(t){function e(){c(this),this.initialize&&this.initialize.apply(this,arguments),this.callInitHooks()}var i,n=e.__super__=this.prototype,o=R(n);for(i in(o.constructor=e).prototype=o,this)Object.prototype.hasOwnProperty.call(this,i)&&"prototype"!==i&&"__super__"!==i&&(e[i]=this[i]);if(t.statics&&l(e,t.statics),t.includes){var s=t.includes;if("undefined"!=typeof L&&L&&L.Mixin){s=d(s)?s:[s];for(var r=0;r<s.length;r++)s[r]===L.Mixin.Events&&console.warn("Deprecated include of L.Mixin.Events: this property will be removed in future releases, please inherit from L.Evented instead.",(new Error).stack)}l.apply(null,[o].concat(t.includes))}return l(o,t),delete o.statics,delete o.includes,o.options&&(o.options=n.options?R(n.options):{},l(o.options,t.options)),o._initHooks=[],o.callInitHooks=function(){if(!this._initHooksCalled){n.callInitHooks&&n.callInitHooks.call(this),this._initHooksCalled=!0;for(var t=0,e=o._initHooks.length;t<e;t++)o._initHooks[t].call(this)}},e},et.include=function(t){var e=this.prototype.options;return l(this.prototype,t),t.options&&(this.prototype.options=e,this.mergeOptions(t.options)),this},et.mergeOptions=function(t){return l(this.prototype.options,t),this},et.addInitHook=function(t){var e=Array.prototype.slice.call(arguments,1),i="function"==typeof t?t:function(){this[t].apply(this,e)};return this.prototype._initHooks=this.prototype._initHooks||[],this.prototype._initHooks.push(i),this};var e={on:function(t,e,i){if("object"==typeof t)for(var n in t)this._on(n,t[n],e);else for(var o=0,s=(t=F(t)).length;o<s;o++)this._on(t[o],e,i);return this},off:function(t,e,i){if(arguments.length)if("object"==typeof t)for(var n in t)this._off(n,t[n],e);else{t=F(t);for(var o=1===arguments.length,s=0,r=t.length;s<r;s++)o?this._off(t[s]):this._off(t[s],e,i)}else delete this._events;return this},_on:function(t,e,i,n){"function"!=typeof e?console.warn("wrong listener type: "+typeof e):!1===this._listens(t,e,i)&&(e={fn:e,ctx:i=i===this?void 0:i},n&&(e.once=!0),this._events=this._events||{},this._events[t]=this._events[t]||[],this._events[t].push(e))},_off:function(t,e,i){var n,o,s;if(this._events&&(n=this._events[t]))if(1===arguments.length){if(this._firingCount)for(o=0,s=n.length;o<s;o++)n[o].fn=u;delete this._events[t]}else"function"!=typeof e?console.warn("wrong listener type: "+typeof e):!1!==(e=this._listens(t,e,i))&&(i=n[e],this._firingCount&&(i.fn=u,this._events[t]=n=n.slice()),n.splice(e,1))},fire:function(t,e,i){if(this.listens(t,i)){var n=l({},e,{type:t,target:this,sourceTarget:e&&e.sourceTarget||this});if(this._events){var o=this._events[t];if(o){this._firingCount=this._firingCount+1||1;for(var s=0,r=o.length;s<r;s++){var a=o[s],h=a.fn;a.once&&this.off(t,h,a.ctx),h.call(a.ctx||this,n)}this._firingCount--}}i&&this._propagateEvent(n)}return this},listens:function(t,e,i,n){"string"!=typeof t&&console.warn('"string" type argument expected');var o=e,s=("function"!=typeof e&&(n=!!e,i=o=void 0),this._events&&this._events[t]);if(s&&s.length&&!1!==this._listens(t,o,i))return!0;if(n)for(var r in this._eventParents)if(this._eventParents[r].listens(t,e,i,n))return!0;return!1},_listens:function(t,e,i){if(this._events){var n=this._events[t]||[];if(!e)return!!n.length;i===this&&(i=void 0);for(var o=0,s=n.length;o<s;o++)if(n[o].fn===e&&n[o].ctx===i)return o}return!1},once:function(t,e,i){if("object"==typeof t)for(var n in t)this._on(n,t[n],e,!0);else for(var o=0,s=(t=F(t)).length;o<s;o++)this._on(t[o],e,i,!0);return this},addEventParent:function(t){return this._eventParents=this._eventParents||{},this._eventParents[h(t)]=t,this},removeEventParent:function(t){return this._eventParents&&delete this._eventParents[h(t)],this},_propagateEvent:function(t){for(var e in this._eventParents)this._eventParents[e].fire(t.type,l({layer:t.target,propagatedFrom:t.target},t),!0)}},it=(e.addEventListener=e.on,e.removeEventListener=e.clearAllEventListeners=e.off,e.addOneTimeEventListener=e.once,e.fireEvent=e.fire,e.hasEventListeners=e.listens,et.extend(e));function p(t,e,i){this.x=i?Math.round(t):t,this.y=i?Math.round(e):e}var nt=Math.trunc||function(t){return 0<t?Math.floor(t):Math.ceil(t)};function m(t,e,i){return t instanceof p?t:d(t)?new p(t[0],t[1]):null==t?t:"object"==typeof t&&"x"in t&&"y"in t?new p(t.x,t.y):new p(t,e,i)}function f(t,e){if(t)for(var i=e?[t,e]:t,n=0,o=i.length;n<o;n++)this.extend(i[n])}function _(t,e){return!t||t instanceof f?t:new f(t,e)}function s(t,e){if(t)for(var i=e?[t,e]:t,n=0,o=i.length;n<o;n++)this.extend(i[n])}function g(t,e){return t instanceof s?t:new s(t,e)}function v(t,e,i){if(isNaN(t)||isNaN(e))throw new Error("Invalid LatLng object: ("+t+", "+e+")");this.lat=+t,this.lng=+e,void 0!==i&&(this.alt=+i)}function w(t,e,i){return t instanceof v?t:d(t)&&"object"!=typeof t[0]?3===t.length?new v(t[0],t[1],t[2]):2===t.length?new v(t[0],t[1]):null:null==t?t:"object"==typeof t&&"lat"in t?new v(t.lat,"lng"in t?t.lng:t.lon,t.alt):void 0===e?null:new v(t,e,i)}p.prototype={clone:function(){return new p(this.x,this.y)},add:function(t){return this.clone()._add(m(t))},_add:function(t){return this.x+=t.x,this.y+=t.y,this},subtract:function(t){return this.clone()._subtract(m(t))},_subtract:function(t){return this.x-=t.x,this.y-=t.y,this},divideBy:function(t){return this.clone()._divideBy(t)},_divideBy:function(t){return this.x/=t,this.y/=t,this},multiplyBy:function(t){return this.clone()._multiplyBy(t)},_multiplyBy:function(t){return this.x*=t,this.y*=t,this},scaleBy:function(t){return new p(this.x*t.x,this.y*t.y)},unscaleBy:function(t){return new p(this.x/t.x,this.y/t.y)},round:function(){return this.clone()._round()},_round:function(){return this.x=Math.round(this.x),this.y=Math.round(this.y),this},floor:function(){return this.clone()._floor()},_floor:function(){return this.x=Math.floor(this.x),this.y=Math.floor(this.y),this},ceil:function(){return this.clone()._ceil()},_ceil:function(){return this.x=Math.ceil(this.x),this.y=Math.ceil(this.y),this},trunc:function(){return this.clone()._trunc()},_trunc:function(){return this.x=nt(this.x),this.y=nt(this.y),this},distanceTo:function(t){var e=(t=m(t)).x-this.x,t=t.y-this.y;return Math.sqrt(e*e+t*t)},equals:function(t){return(t=m(t)).x===this.x&&t.y===this.y},contains:function(t){return t=m(t),Math.abs(t.x)<=Math.abs(this.x)&&Math.abs(t.y)<=Math.abs(this.y)},toString:function(){return"Point("+i(this.x)+", "+i(this.y)+")"}},f.prototype={extend:function(t){var e,i;if(t){if(t instanceof p||"number"==typeof t[0]||"x"in t)e=i=m(t);else if(e=(t=_(t)).min,i=t.max,!e||!i)return this;this.min||this.max?(this.min.x=Math.min(e.x,this.min.x),this.max.x=Math.max(i.x,this.max.x),this.min.y=Math.min(e.y,this.min.y),this.max.y=Math.max(i.y,this.max.y)):(this.min=e.clone(),this.max=i.clone())}return this},getCenter:function(t){return m((this.min.x+this.max.x)/2,(this.min.y+this.max.y)/2,t)},getBottomLeft:function(){return m(this.min.x,this.max.y)},getTopRight:function(){return m(this.max.x,this.min.y)},getTopLeft:function(){return this.min},getBottomRight:function(){return this.max},getSize:function(){return this.max.subtract(this.min)},contains:function(t){var e,i;return(t=("number"==typeof t[0]||t instanceof p?m:_)(t))instanceof f?(e=t.min,i=t.max):e=i=t,e.x>=this.min.x&&i.x<=this.max.x&&e.y>=this.min.y&&i.y<=this.max.y},intersects:function(t){t=_(t);var e=this.min,i=this.max,n=t.min,t=t.max,o=t.x>=e.x&&n.x<=i.x,t=t.y>=e.y&&n.y<=i.y;return o&&t},overlaps:function(t){t=_(t);var e=this.min,i=this.max,n=t.min,t=t.max,o=t.x>e.x&&n.x<i.x,t=t.y>e.y&&n.y<i.y;return o&&t},isValid:function(){return!(!this.min||!this.max)},pad:function(t){var e=this.min,i=this.max,n=Math.abs(e.x-i.x)*t,t=Math.abs(e.y-i.y)*t;return _(m(e.x-n,e.y-t),m(i.x+n,i.y+t))},equals:function(t){return!!t&&(t=_(t),this.min.equals(t.getTopLeft())&&this.max.equals(t.getBottomRight()))}},s.prototype={extend:function(t){var e,i,n=this._southWest,o=this._northEast;if(t instanceof v)i=e=t;else{if(!(t instanceof s))return t?this.extend(w(t)||g(t)):this;if(e=t._southWest,i=t._northEast,!e||!i)return this}return n||o?(n.lat=Math.min(e.lat,n.lat),n.lng=Math.min(e.lng,n.lng),o.lat=Math.max(i.lat,o.lat),o.lng=Math.max(i.lng,o.lng)):(this._southWest=new v(e.lat,e.lng),this._northEast=new v(i.lat,i.lng)),this},pad:function(t){var e=this._southWest,i=this._northEast,n=Math.abs(e.lat-i.lat)*t,t=Math.abs(e.lng-i.lng)*t;return new s(new v(e.lat-n,e.lng-t),new v(i.lat+n,i.lng+t))},getCenter:function(){return new v((this._southWest.lat+this._northEast.lat)/2,(this._southWest.lng+this._northEast.lng)/2)},getSouthWest:function(){return this._southWest},getNorthEast:function(){return this._northEast},getNorthWest:function(){return new v(this.getNorth(),this.getWest())},getSouthEast:function(){return new v(this.getSouth(),this.getEast())},getWest:function(){return this._southWest.lng},getSouth:function(){return this._southWest.lat},getEast:function(){return this._northEast.lng},getNorth:function(){return this._northEast.lat},contains:function(t){t=("number"==typeof t[0]||t instanceof v||"lat"in t?w:g)(t);var e,i,n=this._southWest,o=this._northEast;return t instanceof s?(e=t.getSouthWest(),i=t.getNorthEast()):e=i=t,e.lat>=n.lat&&i.lat<=o.lat&&e.lng>=n.lng&&i.lng<=o.lng},intersects:function(t){t=g(t);var e=this._southWest,i=this._northEast,n=t.getSouthWest(),t=t.getNorthEast(),o=t.lat>=e.lat&&n.lat<=i.lat,t=t.lng>=e.lng&&n.lng<=i.lng;return o&&t},overlaps:function(t){t=g(t);var e=this._southWest,i=this._northEast,n=t.getSouthWest(),t=t.getNorthEast(),o=t.lat>e.lat&&n.lat<i.lat,t=t.lng>e.lng&&n.lng<i.lng;return o&&t},toBBoxString:function(){return[this.getWest(),this.getSouth(),this.getEast(),this.getNorth()].join(",")},equals:function(t,e){return!!t&&(t=g(t),this._southWest.equals(t.getSouthWest(),e)&&this._northEast.equals(t.getNorthEast(),e))},isValid:function(){return!(!this._southWest||!this._northEast)}};var ot={latLngToPoint:function(t,e){t=this.projection.project(t),e=this.scale(e);return this.transformation._transform(t,e)},pointToLatLng:function(t,e){e=this.scale(e),t=this.transformation.untransform(t,e);return this.projection.unproject(t)},project:function(t){return this.projection.project(t)},unproject:function(t){return this.projection.unproject(t)},scale:function(t){return 256*Math.pow(2,t)},zoom:function(t){return Math.log(t/256)/Math.LN2},getProjectedBounds:function(t){var e;return this.infinite?null:(e=this.projection.bounds,t=this.scale(t),new f(this.transformation.transform(e.min,t),this.transformation.transform(e.max,t)))},infinite:!(v.prototype={equals:function(t,e){return!!t&&(t=w(t),Math.max(Math.abs(this.lat-t.lat),Math.abs(this.lng-t.lng))<=(void 0===e?1e-9:e))},toString:function(t){return"LatLng("+i(this.lat,t)+", "+i(this.lng,t)+")"},distanceTo:function(t){return st.distance(this,w(t))},wrap:function(){return st.wrapLatLng(this)},toBounds:function(t){var t=180*t/40075017,e=t/Math.cos(Math.PI/180*this.lat);return g([this.lat-t,this.lng-e],[this.lat+t,this.lng+e])},clone:function(){return new v(this.lat,this.lng,this.alt)}}),wrapLatLng:function(t){var e=this.wrapLng?H(t.lng,this.wrapLng,!0):t.lng;return new v(this.wrapLat?H(t.lat,this.wrapLat,!0):t.lat,e,t.alt)},wrapLatLngBounds:function(t){var e=t.getCenter(),i=this.wrapLatLng(e),n=e.lat-i.lat,e=e.lng-i.lng;return 0==n&&0==e?t:(i=t.getSouthWest(),t=t.getNorthEast(),new s(new v(i.lat-n,i.lng-e),new v(t.lat-n,t.lng-e)))}},st=l({},ot,{wrapLng:[-180,180],R:6371e3,distance:function(t,e){var i=Math.PI/180,n=t.lat*i,o=e.lat*i,s=Math.sin((e.lat-t.lat)*i/2),e=Math.sin((e.lng-t.lng)*i/2),t=s*s+Math.cos(n)*Math.cos(o)*e*e,i=2*Math.atan2(Math.sqrt(t),Math.sqrt(1-t));return this.R*i}}),rt=6378137,rt={R:rt,MAX_LATITUDE:85.0511287798,project:function(t){var e=Math.PI/180,i=this.MAX_LATITUDE,i=Math.max(Math.min(i,t.lat),-i),i=Math.sin(i*e);return new p(this.R*t.lng*e,this.R*Math.log((1+i)/(1-i))/2)},unproject:function(t){var e=180/Math.PI;return new v((2*Math.atan(Math.exp(t.y/this.R))-Math.PI/2)*e,t.x*e/this.R)},bounds:new f([-(rt=rt*Math.PI),-rt],[rt,rt])};function at(t,e,i,n){d(t)?(this._a=t[0],this._b=t[1],this._c=t[2],this._d=t[3]):(this._a=t,this._b=e,this._c=i,this._d=n)}function ht(t,e,i,n){return new at(t,e,i,n)}at.prototype={transform:function(t,e){return this._transform(t.clone(),e)},_transform:function(t,e){return t.x=(e=e||1)*(this._a*t.x+this._b),t.y=e*(this._c*t.y+this._d),t},untransform:function(t,e){return new p((t.x/(e=e||1)-this._b)/this._a,(t.y/e-this._d)/this._c)}};var lt=l({},st,{code:"EPSG:3857",projection:rt,transformation:ht(lt=.5/(Math.PI*rt.R),.5,-lt,.5)}),ut=l({},lt,{code:"EPSG:900913"});function ct(t){return document.createElementNS("http://www.w3.org/2000/svg",t)}function dt(t,e){for(var i,n,o,s,r="",a=0,h=t.length;a<h;a++){for(i=0,n=(o=t[a]).length;i<n;i++)r+=(i?"L":"M")+(s=o[i]).x+" "+s.y;r+=e?b.svg?"z":"x":""}return r||"M0 0"}var _t=document.documentElement.style,pt="ActiveXObject"in window,mt=pt&&!document.addEventListener,n="msLaunchUri"in navigator&&!("documentMode"in document),ft=y("webkit"),gt=y("android"),vt=y("android 2")||y("android 3"),yt=parseInt(/WebKit\/([0-9]+)|$/.exec(navigator.userAgent)[1],10),yt=gt&&y("Google")&&yt<537&&!("AudioNode"in window),xt=!!window.opera,wt=!n&&y("chrome"),bt=y("gecko")&&!ft&&!xt&&!pt,Pt=!wt&&y("safari"),Lt=y("phantom"),o="OTransition"in _t,Tt=0===navigator.platform.indexOf("Win"),Mt=pt&&"transition"in _t,zt="WebKitCSSMatrix"in window&&"m11"in new window.WebKitCSSMatrix&&!vt,_t="MozPerspective"in _t,Ct=!window.L_DISABLE_3D&&(Mt||zt||_t)&&!o&&!Lt,Zt="undefined"!=typeof orientation||y("mobile"),St=Zt&&ft,Et=Zt&&zt,kt=!window.PointerEvent&&window.MSPointerEvent,Ot=!(!window.PointerEvent&&!kt),At="ontouchstart"in window||!!window.TouchEvent,Bt=!window.L_NO_TOUCH&&(At||Ot),It=Zt&&xt,Rt=Zt&&bt,Nt=1<(window.devicePixelRatio||window.screen.deviceXDPI/window.screen.logicalXDPI),Dt=function(){var t=!1;try{var e=Object.defineProperty({},"passive",{get:function(){t=!0}});window.addEventListener("testPassiveEventSupport",u,e),window.removeEventListener("testPassiveEventSupport",u,e)}catch(t){}return t}(),jt=!!document.createElement("canvas").getContext,Ht=!(!document.createElementNS||!ct("svg").createSVGRect),Wt=!!Ht&&((Wt=document.createElement("div")).innerHTML="<svg/>","http://www.w3.org/2000/svg"===(Wt.firstChild&&Wt.firstChild.namespaceURI));function y(t){return 0<=navigator.userAgent.toLowerCase().indexOf(t)}var b={ie:pt,ielt9:mt,edge:n,webkit:ft,android:gt,android23:vt,androidStock:yt,opera:xt,chrome:wt,gecko:bt,safari:Pt,phantom:Lt,opera12:o,win:Tt,ie3d:Mt,webkit3d:zt,gecko3d:_t,any3d:Ct,mobile:Zt,mobileWebkit:St,mobileWebkit3d:Et,msPointer:kt,pointer:Ot,touch:Bt,touchNative:At,mobileOpera:It,mobileGecko:Rt,retina:Nt,passiveEvents:Dt,canvas:jt,svg:Ht,vml:!Ht&&function(){try{var t=document.createElement("div"),e=(t.innerHTML='<v:shape adj="1"/>',t.firstChild);return e.style.behavior="url(#default#VML)",e&&"object"==typeof e.adj}catch(t){return!1}}(),inlineSvg:Wt,mac:0===navigator.platform.indexOf("Mac"),linux:0===navigator.platform.indexOf("Linux")},Ft=b.msPointer?"MSPointerDown":"pointerdown",Ut=b.msPointer?"MSPointerMove":"pointermove",Vt=b.msPointer?"MSPointerUp":"pointerup",qt=b.msPointer?"MSPointerCancel":"pointercancel",Gt={touchstart:Ft,touchmove:Ut,touchend:Vt,touchcancel:qt},Kt={touchstart:function(t,e){e.MSPOINTER_TYPE_TOUCH&&e.pointerType===e.MSPOINTER_TYPE_TOUCH&&O(e);ee(t,e)},touchmove:ee,touchend:ee,touchcancel:ee},Yt={},Xt=!1;function Jt(t,e,i){return"touchstart"!==e||Xt||(document.addEventListener(Ft,$t,!0),document.addEventListener(Ut,Qt,!0),document.addEventListener(Vt,te,!0),document.addEventListener(qt,te,!0),Xt=!0),Kt[e]?(i=Kt[e].bind(this,i),t.addEventListener(Gt[e],i,!1),i):(console.warn("wrong event specified:",e),u)}function $t(t){Yt[t.pointerId]=t}function Qt(t){Yt[t.pointerId]&&(Yt[t.pointerId]=t)}function te(t){delete Yt[t.pointerId]}function ee(t,e){if(e.pointerType!==(e.MSPOINTER_TYPE_MOUSE||"mouse")){for(var i in e.touches=[],Yt)e.touches.push(Yt[i]);e.changedTouches=[e],t(e)}}var ie=200;function ne(t,i){t.addEventListener("dblclick",i);var n,o=0;function e(t){var e;1!==t.detail?n=t.detail:"mouse"===t.pointerType||t.sourceCapabilities&&!t.sourceCapabilities.firesTouchEvents||((e=Ne(t)).some(function(t){return t instanceof HTMLLabelElement&&t.attributes.for})&&!e.some(function(t){return t instanceof HTMLInputElement||t instanceof HTMLSelectElement})||((e=Date.now())-o<=ie?2===++n&&i(function(t){var e,i,n={};for(i in t)e=t[i],n[i]=e&&e.bind?e.bind(t):e;return(t=n).type="dblclick",n.detail=2,n.isTrusted=!1,n._simulated=!0,n}(t)):n=1,o=e))}return t.addEventListener("click",e),{dblclick:i,simDblclick:e}}var oe,se,re,ae,he,le,ue=we(["transform","webkitTransform","OTransform","MozTransform","msTransform"]),ce=we(["webkitTransition","transition","OTransition","MozTransition","msTransition"]),de="webkitTransition"===ce||"OTransition"===ce?ce+"End":"transitionend";function _e(t){return"string"==typeof t?document.getElementById(t):t}function pe(t,e){var i=t.style[e]||t.currentStyle&&t.currentStyle[e];return"auto"===(i=i&&"auto"!==i||!document.defaultView?i:(t=document.defaultView.getComputedStyle(t,null))?t[e]:null)?null:i}function P(t,e,i){t=document.createElement(t);return t.className=e||"",i&&i.appendChild(t),t}function T(t){var e=t.parentNode;e&&e.removeChild(t)}function me(t){for(;t.firstChild;)t.removeChild(t.firstChild)}function fe(t){var e=t.parentNode;e&&e.lastChild!==t&&e.appendChild(t)}function ge(t){var e=t.parentNode;e&&e.firstChild!==t&&e.insertBefore(t,e.firstChild)}function ve(t,e){return void 0!==t.classList?t.classList.contains(e):0<(t=xe(t)).length&&new RegExp("(^|\\s)"+e+"(\\s|$)").test(t)}function M(t,e){var i;if(void 0!==t.classList)for(var n=F(e),o=0,s=n.length;o<s;o++)t.classList.add(n[o]);else ve(t,e)||ye(t,((i=xe(t))?i+" ":"")+e)}function z(t,e){void 0!==t.classList?t.classList.remove(e):ye(t,W((" "+xe(t)+" ").replace(" "+e+" "," ")))}function ye(t,e){void 0===t.className.baseVal?t.className=e:t.className.baseVal=e}function xe(t){return void 0===(t=t.correspondingElement?t.correspondingElement:t).className.baseVal?t.className:t.className.baseVal}function C(t,e){if("opacity"in t.style)t.style.opacity=e;else if("filter"in t.style){var i=!1,n="DXImageTransform.Microsoft.Alpha";try{i=t.filters.item(n)}catch(t){if(1===e)return}e=Math.round(100*e),i?(i.Enabled=100!==e,i.Opacity=e):t.style.filter+=" progid:"+n+"(opacity="+e+")"}}function we(t){for(var e=document.documentElement.style,i=0;i<t.length;i++)if(t[i]in e)return t[i];return!1}function be(t,e,i){e=e||new p(0,0);t.style[ue]=(b.ie3d?"translate("+e.x+"px,"+e.y+"px)":"translate3d("+e.x+"px,"+e.y+"px,0)")+(i?" scale("+i+")":"")}function Z(t,e){t._leaflet_pos=e,b.any3d?be(t,e):(t.style.left=e.x+"px",t.style.top=e.y+"px")}function Pe(t){return t._leaflet_pos||new p(0,0)}function Le(){S(window,"dragstart",O)}function Te(){k(window,"dragstart",O)}function Me(t){for(;-1===t.tabIndex;)t=t.parentNode;t.style&&(ze(),le=(he=t).style.outlineStyle,t.style.outlineStyle="none",S(window,"keydown",ze))}function ze(){he&&(he.style.outlineStyle=le,le=he=void 0,k(window,"keydown",ze))}function Ce(t){for(;!((t=t.parentNode).offsetWidth&&t.offsetHeight||t===document.body););return t}function Ze(t){var e=t.getBoundingClientRect();return{x:e.width/t.offsetWidth||1,y:e.height/t.offsetHeight||1,boundingClientRect:e}}ae="onselectstart"in document?(re=function(){S(window,"selectstart",O)},function(){k(window,"selectstart",O)}):(se=we(["userSelect","WebkitUserSelect","OUserSelect","MozUserSelect","msUserSelect"]),re=function(){var t;se&&(t=document.documentElement.style,oe=t[se],t[se]="none")},function(){se&&(document.documentElement.style[se]=oe,oe=void 0)});pt={__proto__:null,TRANSFORM:ue,TRANSITION:ce,TRANSITION_END:de,get:_e,getStyle:pe,create:P,remove:T,empty:me,toFront:fe,toBack:ge,hasClass:ve,addClass:M,removeClass:z,setClass:ye,getClass:xe,setOpacity:C,testProp:we,setTransform:be,setPosition:Z,getPosition:Pe,get disableTextSelection(){return re},get enableTextSelection(){return ae},disableImageDrag:Le,enableImageDrag:Te,preventOutline:Me,restoreOutline:ze,getSizedParentNode:Ce,getScale:Ze};function S(t,e,i,n){if(e&&"object"==typeof e)for(var o in e)ke(t,o,e[o],i);else for(var s=0,r=(e=F(e)).length;s<r;s++)ke(t,e[s],i,n);return this}var E="_leaflet_events";function k(t,e,i,n){if(1===arguments.length)Se(t),delete t[E];else if(e&&"object"==typeof e)for(var o in e)Oe(t,o,e[o],i);else if(e=F(e),2===arguments.length)Se(t,function(t){return-1!==G(e,t)});else for(var s=0,r=e.length;s<r;s++)Oe(t,e[s],i,n);return this}function Se(t,e){for(var i in t[E]){var n=i.split(/\d/)[0];e&&!e(n)||Oe(t,n,null,null,i)}}var Ee={mouseenter:"mouseover",mouseleave:"mouseout",wheel:!("onwheel"in window)&&"mousewheel"};function ke(e,t,i,n){var o,s,r=t+h(i)+(n?"_"+h(n):"");e[E]&&e[E][r]||(s=o=function(t){return i.call(n||e,t||window.event)},!b.touchNative&&b.pointer&&0===t.indexOf("touch")?o=Jt(e,t,o):b.touch&&"dblclick"===t?o=ne(e,o):"addEventListener"in e?"touchstart"===t||"touchmove"===t||"wheel"===t||"mousewheel"===t?e.addEventListener(Ee[t]||t,o,!!b.passiveEvents&&{passive:!1}):"mouseenter"===t||"mouseleave"===t?e.addEventListener(Ee[t],o=function(t){t=t||window.event,We(e,t)&&s(t)},!1):e.addEventListener(t,s,!1):e.attachEvent("on"+t,o),e[E]=e[E]||{},e[E][r]=o)}function Oe(t,e,i,n,o){o=o||e+h(i)+(n?"_"+h(n):"");var s,r,i=t[E]&&t[E][o];i&&(!b.touchNative&&b.pointer&&0===e.indexOf("touch")?(n=t,r=i,Gt[s=e]?n.removeEventListener(Gt[s],r,!1):console.warn("wrong event specified:",s)):b.touch&&"dblclick"===e?(n=i,(r=t).removeEventListener("dblclick",n.dblclick),r.removeEventListener("click",n.simDblclick)):"removeEventListener"in t?t.removeEventListener(Ee[e]||e,i,!1):t.detachEvent("on"+e,i),t[E][o]=null)}function Ae(t){return t.stopPropagation?t.stopPropagation():t.originalEvent?t.originalEvent._stopped=!0:t.cancelBubble=!0,this}function Be(t){return ke(t,"wheel",Ae),this}function Ie(t){return S(t,"mousedown touchstart dblclick contextmenu",Ae),t._leaflet_disable_click=!0,this}function O(t){return t.preventDefault?t.preventDefault():t.returnValue=!1,this}function Re(t){return O(t),Ae(t),this}function Ne(t){if(t.composedPath)return t.composedPath();for(var e=[],i=t.target;i;)e.push(i),i=i.parentNode;return e}function De(t,e){var i,n;return e?(n=(i=Ze(e)).boundingClientRect,new p((t.clientX-n.left)/i.x-e.clientLeft,(t.clientY-n.top)/i.y-e.clientTop)):new p(t.clientX,t.clientY)}var je=b.linux&&b.chrome?window.devicePixelRatio:b.mac?3*window.devicePixelRatio:0<window.devicePixelRatio?2*window.devicePixelRatio:1;function He(t){return b.edge?t.wheelDeltaY/2:t.deltaY&&0===t.deltaMode?-t.deltaY/je:t.deltaY&&1===t.deltaMode?20*-t.deltaY:t.deltaY&&2===t.deltaMode?60*-t.deltaY:t.deltaX||t.deltaZ?0:t.wheelDelta?(t.wheelDeltaY||t.wheelDelta)/2:t.detail&&Math.abs(t.detail)<32765?20*-t.detail:t.detail?t.detail/-32765*60:0}function We(t,e){var i=e.relatedTarget;if(!i)return!0;try{for(;i&&i!==t;)i=i.parentNode}catch(t){return!1}return i!==t}var mt={__proto__:null,on:S,off:k,stopPropagation:Ae,disableScrollPropagation:Be,disableClickPropagation:Ie,preventDefault:O,stop:Re,getPropagationPath:Ne,getMousePosition:De,getWheelDelta:He,isExternalTarget:We,addListener:S,removeListener:k},Fe=it.extend({run:function(t,e,i,n){this.stop(),this._el=t,this._inProgress=!0,this._duration=i||.25,this._easeOutPower=1/Math.max(n||.5,.2),this._startPos=Pe(t),this._offset=e.subtract(this._startPos),this._startTime=+new Date,this.fire("start"),this._animate()},stop:function(){this._inProgress&&(this._step(!0),this._complete())},_animate:function(){this._animId=x(this._animate,this),this._step()},_step:function(t){var e=+new Date-this._startTime,i=1e3*this._duration;e<i?this._runFrame(this._easeOut(e/i),t):(this._runFrame(1),this._complete())},_runFrame:function(t,e){t=this._startPos.add(this._offset.multiplyBy(t));e&&t._round(),Z(this._el,t),this.fire("step")},_complete:function(){r(this._animId),this._inProgress=!1,this.fire("end")},_easeOut:function(t){return 1-Math.pow(1-t,this._easeOutPower)}}),A=it.extend({options:{crs:lt,center:void 0,zoom:void 0,minZoom:void 0,maxZoom:void 0,layers:[],maxBounds:void 0,renderer:void 0,zoomAnimation:!0,zoomAnimationThreshold:4,fadeAnimation:!0,markerZoomAnimation:!0,transform3DLimit:8388608,zoomSnap:1,zoomDelta:1,trackResize:!0},initialize:function(t,e){e=c(this,e),this._handlers=[],this._layers={},this._zoomBoundLayers={},this._sizeChanged=!0,this._initContainer(t),this._initLayout(),this._onResize=a(this._onResize,this),this._initEvents(),e.maxBounds&&this.setMaxBounds(e.maxBounds),void 0!==e.zoom&&(this._zoom=this._limitZoom(e.zoom)),e.center&&void 0!==e.zoom&&this.setView(w(e.center),e.zoom,{reset:!0}),this.callInitHooks(),this._zoomAnimated=ce&&b.any3d&&!b.mobileOpera&&this.options.zoomAnimation,this._zoomAnimated&&(this._createAnimProxy(),S(this._proxy,de,this._catchTransitionEnd,this)),this._addLayers(this.options.layers)},setView:function(t,e,i){if((e=void 0===e?this._zoom:this._limitZoom(e),t=this._limitCenter(w(t),e,this.options.maxBounds),i=i||{},this._stop(),this._loaded&&!i.reset&&!0!==i)&&(void 0!==i.animate&&(i.zoom=l({animate:i.animate},i.zoom),i.pan=l({animate:i.animate,duration:i.duration},i.pan)),this._zoom!==e?this._tryAnimatedZoom&&this._tryAnimatedZoom(t,e,i.zoom):this._tryAnimatedPan(t,i.pan)))return clearTimeout(this._sizeTimer),this;return this._resetView(t,e,i.pan&&i.pan.noMoveStart),this},setZoom:function(t,e){return this._loaded?this.setView(this.getCenter(),t,{zoom:e}):(this._zoom=t,this)},zoomIn:function(t,e){return t=t||(b.any3d?this.options.zoomDelta:1),this.setZoom(this._zoom+t,e)},zoomOut:function(t,e){return t=t||(b.any3d?this.options.zoomDelta:1),this.setZoom(this._zoom-t,e)},setZoomAround:function(t,e,i){var n=this.getZoomScale(e),o=this.getSize().divideBy(2),t=(t instanceof p?t:this.latLngToContainerPoint(t)).subtract(o).multiplyBy(1-1/n),n=this.containerPointToLatLng(o.add(t));return this.setView(n,e,{zoom:i})},_getBoundsCenterZoom:function(t,e){e=e||{},t=t.getBounds?t.getBounds():g(t);var i=m(e.paddingTopLeft||e.padding||[0,0]),n=m(e.paddingBottomRight||e.padding||[0,0]),o=this.getBoundsZoom(t,!1,i.add(n));return(o="number"==typeof e.maxZoom?Math.min(e.maxZoom,o):o)===1/0?{center:t.getCenter(),zoom:o}:(e=n.subtract(i).divideBy(2),n=this.project(t.getSouthWest(),o),i=this.project(t.getNorthEast(),o),{center:this.unproject(n.add(i).divideBy(2).add(e),o),zoom:o})},fitBounds:function(t,e){if((t=g(t)).isValid())return t=this._getBoundsCenterZoom(t,e),this.setView(t.center,t.zoom,e);throw new Error("Bounds are not valid.")},fitWorld:function(t){return this.fitBounds([[-90,-180],[90,180]],t)},panTo:function(t,e){return this.setView(t,this._zoom,{pan:e})},panBy:function(t,e){var i;return e=e||{},(t=m(t).round()).x||t.y?(!0===e.animate||this.getSize().contains(t)?(this._panAnim||(this._panAnim=new Fe,this._panAnim.on({step:this._onPanTransitionStep,end:this._onPanTransitionEnd},this)),e.noMoveStart||this.fire("movestart"),!1!==e.animate?(M(this._mapPane,"leaflet-pan-anim"),i=this._getMapPanePos().subtract(t).round(),this._panAnim.run(this._mapPane,i,e.duration||.25,e.easeLinearity)):(this._rawPanBy(t),this.fire("move").fire("moveend"))):this._resetView(this.unproject(this.project(this.getCenter()).add(t)),this.getZoom()),this):this.fire("moveend")},flyTo:function(n,o,t){if(!1===(t=t||{}).animate||!b.any3d)return this.setView(n,o,t);this._stop();var s=this.project(this.getCenter()),r=this.project(n),e=this.getSize(),a=this._zoom,h=(n=w(n),o=void 0===o?a:o,Math.max(e.x,e.y)),i=h*this.getZoomScale(a,o),l=r.distanceTo(s)||1,u=1.42,c=u*u;function d(t){t=(i*i-h*h+(t?-1:1)*c*c*l*l)/(2*(t?i:h)*c*l),t=Math.sqrt(t*t+1)-t;return t<1e-9?-18:Math.log(t)}function _(t){return(Math.exp(t)-Math.exp(-t))/2}function p(t){return(Math.exp(t)+Math.exp(-t))/2}var m=d(0);function f(t){return h*(p(m)*(_(t=m+u*t)/p(t))-_(m))/c}var g=Date.now(),v=(d(1)-m)/u,y=t.duration?1e3*t.duration:1e3*v*.8;return this._moveStart(!0,t.noMoveStart),function t(){var e=(Date.now()-g)/y,i=(1-Math.pow(1-e,1.5))*v;e<=1?(this._flyToFrame=x(t,this),this._move(this.unproject(s.add(r.subtract(s).multiplyBy(f(i)/l)),a),this.getScaleZoom(h/(e=i,h*(p(m)/p(m+u*e))),a),{flyTo:!0})):this._move(n,o)._moveEnd(!0)}.call(this),this},flyToBounds:function(t,e){t=this._getBoundsCenterZoom(t,e);return this.flyTo(t.center,t.zoom,e)},setMaxBounds:function(t){return t=g(t),this.listens("moveend",this._panInsideMaxBounds)&&this.off("moveend",this._panInsideMaxBounds),t.isValid()?(this.options.maxBounds=t,this._loaded&&this._panInsideMaxBounds(),this.on("moveend",this._panInsideMaxBounds)):(this.options.maxBounds=null,this)},setMinZoom:function(t){var e=this.options.minZoom;return this.options.minZoom=t,this._loaded&&e!==t&&(this.fire("zoomlevelschange"),this.getZoom()<this.options.minZoom)?this.setZoom(t):this},setMaxZoom:function(t){var e=this.options.maxZoom;return this.options.maxZoom=t,this._loaded&&e!==t&&(this.fire("zoomlevelschange"),this.getZoom()>this.options.maxZoom)?this.setZoom(t):this},panInsideBounds:function(t,e){this._enforcingBounds=!0;var i=this.getCenter(),t=this._limitCenter(i,this._zoom,g(t));return i.equals(t)||this.panTo(t,e),this._enforcingBounds=!1,this},panInside:function(t,e){var i=m((e=e||{}).paddingTopLeft||e.padding||[0,0]),n=m(e.paddingBottomRight||e.padding||[0,0]),o=this.project(this.getCenter()),t=this.project(t),s=this.getPixelBounds(),i=_([s.min.add(i),s.max.subtract(n)]),s=i.getSize();return i.contains(t)||(this._enforcingBounds=!0,n=t.subtract(i.getCenter()),i=i.extend(t).getSize().subtract(s),o.x+=n.x<0?-i.x:i.x,o.y+=n.y<0?-i.y:i.y,this.panTo(this.unproject(o),e),this._enforcingBounds=!1),this},invalidateSize:function(t){if(!this._loaded)return this;t=l({animate:!1,pan:!0},!0===t?{animate:!0}:t);var e=this.getSize(),i=(this._sizeChanged=!0,this._lastCenter=null,this.getSize()),n=e.divideBy(2).round(),o=i.divideBy(2).round(),n=n.subtract(o);return n.x||n.y?(t.animate&&t.pan?this.panBy(n):(t.pan&&this._rawPanBy(n),this.fire("move"),t.debounceMoveend?(clearTimeout(this._sizeTimer),this._sizeTimer=setTimeout(a(this.fire,this,"moveend"),200)):this.fire("moveend")),this.fire("resize",{oldSize:e,newSize:i})):this},stop:function(){return this.setZoom(this._limitZoom(this._zoom)),this.options.zoomSnap||this.fire("viewreset"),this._stop()},locate:function(t){var e,i;return t=this._locateOptions=l({timeout:1e4,watch:!1},t),"geolocation"in navigator?(e=a(this._handleGeolocationResponse,this),i=a(this._handleGeolocationError,this),t.watch?this._locationWatchId=navigator.geolocation.watchPosition(e,i,t):navigator.geolocation.getCurrentPosition(e,i,t)):this._handleGeolocationError({code:0,message:"Geolocation not supported."}),this},stopLocate:function(){return navigator.geolocation&&navigator.geolocation.clearWatch&&navigator.geolocation.clearWatch(this._locationWatchId),this._locateOptions&&(this._locateOptions.setView=!1),this},_handleGeolocationError:function(t){var e;this._container._leaflet_id&&(e=t.code,t=t.message||(1===e?"permission denied":2===e?"position unavailable":"timeout"),this._locateOptions.setView&&!this._loaded&&this.fitWorld(),this.fire("locationerror",{code:e,message:"Geolocation error: "+t+"."}))},_handleGeolocationResponse:function(t){if(this._container._leaflet_id){var e,i,n=new v(t.coords.latitude,t.coords.longitude),o=n.toBounds(2*t.coords.accuracy),s=this._locateOptions,r=(s.setView&&(e=this.getBoundsZoom(o),this.setView(n,s.maxZoom?Math.min(e,s.maxZoom):e)),{latlng:n,bounds:o,timestamp:t.timestamp});for(i in t.coords)"number"==typeof t.coords[i]&&(r[i]=t.coords[i]);this.fire("locationfound",r)}},addHandler:function(t,e){return e&&(e=this[t]=new e(this),this._handlers.push(e),this.options[t]&&e.enable()),this},remove:function(){if(this._initEvents(!0),this.options.maxBounds&&this.off("moveend",this._panInsideMaxBounds),this._containerId!==this._container._leaflet_id)throw new Error("Map container is being reused by another instance");try{delete this._container._leaflet_id,delete this._containerId}catch(t){this._container._leaflet_id=void 0,this._containerId=void 0}for(var t in void 0!==this._locationWatchId&&this.stopLocate(),this._stop(),T(this._mapPane),this._clearControlPos&&this._clearControlPos(),this._resizeRequest&&(r(this._resizeRequest),this._resizeRequest=null),this._clearHandlers(),this._loaded&&this.fire("unload"),this._layers)this._layers[t].remove();for(t in this._panes)T(this._panes[t]);return this._layers=[],this._panes=[],delete this._mapPane,delete this._renderer,this},createPane:function(t,e){e=P("div","leaflet-pane"+(t?" leaflet-"+t.replace("Pane","")+"-pane":""),e||this._mapPane);return t&&(this._panes[t]=e),e},getCenter:function(){return this._checkIfLoaded(),this._lastCenter&&!this._moved()?this._lastCenter.clone():this.layerPointToLatLng(this._getCenterLayerPoint())},getZoom:function(){return this._zoom},getBounds:function(){var t=this.getPixelBounds();return new s(this.unproject(t.getBottomLeft()),this.unproject(t.getTopRight()))},getMinZoom:function(){return void 0===this.options.minZoom?this._layersMinZoom||0:this.options.minZoom},getMaxZoom:function(){return void 0===this.options.maxZoom?void 0===this._layersMaxZoom?1/0:this._layersMaxZoom:this.options.maxZoom},getBoundsZoom:function(t,e,i){t=g(t),i=m(i||[0,0]);var n=this.getZoom()||0,o=this.getMinZoom(),s=this.getMaxZoom(),r=t.getNorthWest(),t=t.getSouthEast(),i=this.getSize().subtract(i),t=_(this.project(t,n),this.project(r,n)).getSize(),r=b.any3d?this.options.zoomSnap:1,a=i.x/t.x,i=i.y/t.y,t=e?Math.max(a,i):Math.min(a,i),n=this.getScaleZoom(t,n);return r&&(n=Math.round(n/(r/100))*(r/100),n=e?Math.ceil(n/r)*r:Math.floor(n/r)*r),Math.max(o,Math.min(s,n))},getSize:function(){return this._size&&!this._sizeChanged||(this._size=new p(this._container.clientWidth||0,this._container.clientHeight||0),this._sizeChanged=!1),this._size.clone()},getPixelBounds:function(t,e){t=this._getTopLeftPoint(t,e);return new f(t,t.add(this.getSize()))},getPixelOrigin:function(){return this._checkIfLoaded(),this._pixelOrigin},getPixelWorldBounds:function(t){return this.options.crs.getProjectedBounds(void 0===t?this.getZoom():t)},getPane:function(t){return"string"==typeof t?this._panes[t]:t},getPanes:function(){return this._panes},getContainer:function(){return this._container},getZoomScale:function(t,e){var i=this.options.crs;return e=void 0===e?this._zoom:e,i.scale(t)/i.scale(e)},getScaleZoom:function(t,e){var i=this.options.crs,t=(e=void 0===e?this._zoom:e,i.zoom(t*i.scale(e)));return isNaN(t)?1/0:t},project:function(t,e){return e=void 0===e?this._zoom:e,this.options.crs.latLngToPoint(w(t),e)},unproject:function(t,e){return e=void 0===e?this._zoom:e,this.options.crs.pointToLatLng(m(t),e)},layerPointToLatLng:function(t){t=m(t).add(this.getPixelOrigin());return this.unproject(t)},latLngToLayerPoint:function(t){return this.project(w(t))._round()._subtract(this.getPixelOrigin())},wrapLatLng:function(t){return this.options.crs.wrapLatLng(w(t))},wrapLatLngBounds:function(t){return this.options.crs.wrapLatLngBounds(g(t))},distance:function(t,e){return this.options.crs.distance(w(t),w(e))},containerPointToLayerPoint:function(t){return m(t).subtract(this._getMapPanePos())},layerPointToContainerPoint:function(t){return m(t).add(this._getMapPanePos())},containerPointToLatLng:function(t){t=this.containerPointToLayerPoint(m(t));return this.layerPointToLatLng(t)},latLngToContainerPoint:function(t){return this.layerPointToContainerPoint(this.latLngToLayerPoint(w(t)))},mouseEventToContainerPoint:function(t){return De(t,this._container)},mouseEventToLayerPoint:function(t){return this.containerPointToLayerPoint(this.mouseEventToContainerPoint(t))},mouseEventToLatLng:function(t){return this.layerPointToLatLng(this.mouseEventToLayerPoint(t))},_initContainer:function(t){t=this._container=_e(t);if(!t)throw new Error("Map container not found.");if(t._leaflet_id)throw new Error("Map container is already initialized.");S(t,"scroll",this._onScroll,this),this._containerId=h(t)},_initLayout:function(){var t=this._container,e=(this._fadeAnimated=this.options.fadeAnimation&&b.any3d,M(t,"leaflet-container"+(b.touch?" leaflet-touch":"")+(b.retina?" leaflet-retina":"")+(b.ielt9?" leaflet-oldie":"")+(b.safari?" leaflet-safari":"")+(this._fadeAnimated?" leaflet-fade-anim":"")),pe(t,"position"));"absolute"!==e&&"relative"!==e&&"fixed"!==e&&"sticky"!==e&&(t.style.position="relative"),this._initPanes(),this._initControlPos&&this._initControlPos()},_initPanes:function(){var t=this._panes={};this._paneRenderers={},this._mapPane=this.createPane("mapPane",this._container),Z(this._mapPane,new p(0,0)),this.createPane("tilePane"),this.createPane("overlayPane"),this.createPane("shadowPane"),this.createPane("markerPane"),this.createPane("tooltipPane"),this.createPane("popupPane"),this.options.markerZoomAnimation||(M(t.markerPane,"leaflet-zoom-hide"),M(t.shadowPane,"leaflet-zoom-hide"))},_resetView:function(t,e,i){Z(this._mapPane,new p(0,0));var n=!this._loaded,o=(this._loaded=!0,e=this._limitZoom(e),this.fire("viewprereset"),this._zoom!==e);this._moveStart(o,i)._move(t,e)._moveEnd(o),this.fire("viewreset"),n&&this.fire("load")},_moveStart:function(t,e){return t&&this.fire("zoomstart"),e||this.fire("movestart"),this},_move:function(t,e,i,n){void 0===e&&(e=this._zoom);var o=this._zoom!==e;return this._zoom=e,this._lastCenter=t,this._pixelOrigin=this._getNewPixelOrigin(t),n?i&&i.pinch&&this.fire("zoom",i):((o||i&&i.pinch)&&this.fire("zoom",i),this.fire("move",i)),this},_moveEnd:function(t){return t&&this.fire("zoomend"),this.fire("moveend")},_stop:function(){return r(this._flyToFrame),this._panAnim&&this._panAnim.stop(),this},_rawPanBy:function(t){Z(this._mapPane,this._getMapPanePos().subtract(t))},_getZoomSpan:function(){return this.getMaxZoom()-this.getMinZoom()},_panInsideMaxBounds:function(){this._enforcingBounds||this.panInsideBounds(this.options.maxBounds)},_checkIfLoaded:function(){if(!this._loaded)throw new Error("Set map center and zoom first.")},_initEvents:function(t){this._targets={};var e=t?k:S;e((this._targets[h(this._container)]=this)._container,"click dblclick mousedown mouseup mouseover mouseout mousemove contextmenu keypress keydown keyup",this._handleDOMEvent,this),this.options.trackResize&&e(window,"resize",this._onResize,this),b.any3d&&this.options.transform3DLimit&&(t?this.off:this.on).call(this,"moveend",this._onMoveEnd)},_onResize:function(){r(this._resizeRequest),this._resizeRequest=x(function(){this.invalidateSize({debounceMoveend:!0})},this)},_onScroll:function(){this._container.scrollTop=0,this._container.scrollLeft=0},_onMoveEnd:function(){var t=this._getMapPanePos();Math.max(Math.abs(t.x),Math.abs(t.y))>=this.options.transform3DLimit&&this._resetView(this.getCenter(),this.getZoom())},_findEventTargets:function(t,e){for(var i,n=[],o="mouseout"===e||"mouseover"===e,s=t.target||t.srcElement,r=!1;s;){if((i=this._targets[h(s)])&&("click"===e||"preclick"===e)&&this._draggableMoved(i)){r=!0;break}if(i&&i.listens(e,!0)){if(o&&!We(s,t))break;if(n.push(i),o)break}if(s===this._container)break;s=s.parentNode}return n=n.length||r||o||!this.listens(e,!0)?n:[this]},_isClickDisabled:function(t){for(;t&&t!==this._container;){if(t._leaflet_disable_click)return!0;t=t.parentNode}},_handleDOMEvent:function(t){var e,i=t.target||t.srcElement;!this._loaded||i._leaflet_disable_events||"click"===t.type&&this._isClickDisabled(i)||("mousedown"===(e=t.type)&&Me(i),this._fireDOMEvent(t,e))},_mouseEvents:["click","dblclick","mouseover","mouseout","contextmenu"],_fireDOMEvent:function(t,e,i){"click"===t.type&&((a=l({},t)).type="preclick",this._fireDOMEvent(a,a.type,i));var n=this._findEventTargets(t,e);if(i){for(var o=[],s=0;s<i.length;s++)i[s].listens(e,!0)&&o.push(i[s]);n=o.concat(n)}if(n.length){"contextmenu"===e&&O(t);var r,a=n[0],h={originalEvent:t};for("keypress"!==t.type&&"keydown"!==t.type&&"keyup"!==t.type&&(r=a.getLatLng&&(!a._radius||a._radius<=10),h.containerPoint=r?this.latLngToContainerPoint(a.getLatLng()):this.mouseEventToContainerPoint(t),h.layerPoint=this.containerPointToLayerPoint(h.containerPoint),h.latlng=r?a.getLatLng():this.layerPointToLatLng(h.layerPoint)),s=0;s<n.length;s++)if(n[s].fire(e,h,!0),h.originalEvent._stopped||!1===n[s].options.bubblingMouseEvents&&-1!==G(this._mouseEvents,e))return}},_draggableMoved:function(t){return(t=t.dragging&&t.dragging.enabled()?t:this).dragging&&t.dragging.moved()||this.boxZoom&&this.boxZoom.moved()},_clearHandlers:function(){for(var t=0,e=this._handlers.length;t<e;t++)this._handlers[t].disable()},whenReady:function(t,e){return this._loaded?t.call(e||this,{target:this}):this.on("load",t,e),this},_getMapPanePos:function(){return Pe(this._mapPane)||new p(0,0)},_moved:function(){var t=this._getMapPanePos();return t&&!t.equals([0,0])},_getTopLeftPoint:function(t,e){return(t&&void 0!==e?this._getNewPixelOrigin(t,e):this.getPixelOrigin()).subtract(this._getMapPanePos())},_getNewPixelOrigin:function(t,e){var i=this.getSize()._divideBy(2);return this.project(t,e)._subtract(i)._add(this._getMapPanePos())._round()},_latLngToNewLayerPoint:function(t,e,i){i=this._getNewPixelOrigin(i,e);return this.project(t,e)._subtract(i)},_latLngBoundsToNewLayerBounds:function(t,e,i){i=this._getNewPixelOrigin(i,e);return _([this.project(t.getSouthWest(),e)._subtract(i),this.project(t.getNorthWest(),e)._subtract(i),this.project(t.getSouthEast(),e)._subtract(i),this.project(t.getNorthEast(),e)._subtract(i)])},_getCenterLayerPoint:function(){return this.containerPointToLayerPoint(this.getSize()._divideBy(2))},_getCenterOffset:function(t){return this.latLngToLayerPoint(t).subtract(this._getCenterLayerPoint())},_limitCenter:function(t,e,i){var n,o;return!i||(n=this.project(t,e),o=this.getSize().divideBy(2),o=new f(n.subtract(o),n.add(o)),o=this._getBoundsOffset(o,i,e),Math.abs(o.x)<=1&&Math.abs(o.y)<=1)?t:this.unproject(n.add(o),e)},_limitOffset:function(t,e){var i;return e?(i=new f((i=this.getPixelBounds()).min.add(t),i.max.add(t)),t.add(this._getBoundsOffset(i,e))):t},_getBoundsOffset:function(t,e,i){e=_(this.project(e.getNorthEast(),i),this.project(e.getSouthWest(),i)),i=e.min.subtract(t.min),e=e.max.subtract(t.max);return new p(this._rebound(i.x,-e.x),this._rebound(i.y,-e.y))},_rebound:function(t,e){return 0<t+e?Math.round(t-e)/2:Math.max(0,Math.ceil(t))-Math.max(0,Math.floor(e))},_limitZoom:function(t){var e=this.getMinZoom(),i=this.getMaxZoom(),n=b.any3d?this.options.zoomSnap:1;return n&&(t=Math.round(t/n)*n),Math.max(e,Math.min(i,t))},_onPanTransitionStep:function(){this.fire("move")},_onPanTransitionEnd:function(){z(this._mapPane,"leaflet-pan-anim"),this.fire("moveend")},_tryAnimatedPan:function(t,e){t=this._getCenterOffset(t)._trunc();return!(!0!==(e&&e.animate)&&!this.getSize().contains(t))&&(this.panBy(t,e),!0)},_createAnimProxy:function(){var t=this._proxy=P("div","leaflet-proxy leaflet-zoom-animated");this._panes.mapPane.appendChild(t),this.on("zoomanim",function(t){var e=ue,i=this._proxy.style[e];be(this._proxy,this.project(t.center,t.zoom),this.getZoomScale(t.zoom,1)),i===this._proxy.style[e]&&this._animatingZoom&&this._onZoomTransitionEnd()},this),this.on("load moveend",this._animMoveEnd,this),this._on("unload",this._destroyAnimProxy,this)},_destroyAnimProxy:function(){T(this._proxy),this.off("load moveend",this._animMoveEnd,this),delete this._proxy},_animMoveEnd:function(){var t=this.getCenter(),e=this.getZoom();be(this._proxy,this.project(t,e),this.getZoomScale(e,1))},_catchTransitionEnd:function(t){this._animatingZoom&&0<=t.propertyName.indexOf("transform")&&this._onZoomTransitionEnd()},_nothingToAnimate:function(){return!this._container.getElementsByClassName("leaflet-zoom-animated").length},_tryAnimatedZoom:function(t,e,i){if(!this._animatingZoom){if(i=i||{},!this._zoomAnimated||!1===i.animate||this._nothingToAnimate()||Math.abs(e-this._zoom)>this.options.zoomAnimationThreshold)return!1;var n=this.getZoomScale(e),n=this._getCenterOffset(t)._divideBy(1-1/n);if(!0!==i.animate&&!this.getSize().contains(n))return!1;x(function(){this._moveStart(!0,i.noMoveStart||!1)._animateZoom(t,e,!0)},this)}return!0},_animateZoom:function(t,e,i,n){this._mapPane&&(i&&(this._animatingZoom=!0,this._animateToCenter=t,this._animateToZoom=e,M(this._mapPane,"leaflet-zoom-anim")),this.fire("zoomanim",{center:t,zoom:e,noUpdate:n}),this._tempFireZoomEvent||(this._tempFireZoomEvent=this._zoom!==this._animateToZoom),this._move(this._animateToCenter,this._animateToZoom,void 0,!0),setTimeout(a(this._onZoomTransitionEnd,this),250))},_onZoomTransitionEnd:function(){this._animatingZoom&&(this._mapPane&&z(this._mapPane,"leaflet-zoom-anim"),this._animatingZoom=!1,this._move(this._animateToCenter,this._animateToZoom,void 0,!0),this._tempFireZoomEvent&&this.fire("zoom"),delete this._tempFireZoomEvent,this.fire("move"),this._moveEnd(!0))}});function Ue(t){return new B(t)}var B=et.extend({options:{position:"topright"},initialize:function(t){c(this,t)},getPosition:function(){return this.options.position},setPosition:function(t){var e=this._map;return e&&e.removeControl(this),this.options.position=t,e&&e.addControl(this),this},getContainer:function(){return this._container},addTo:function(t){this.remove(),this._map=t;var e=this._container=this.onAdd(t),i=this.getPosition(),t=t._controlCorners[i];return M(e,"leaflet-control"),-1!==i.indexOf("bottom")?t.insertBefore(e,t.firstChild):t.appendChild(e),this._map.on("unload",this.remove,this),this},remove:function(){return this._map&&(T(this._container),this.onRemove&&this.onRemove(this._map),this._map.off("unload",this.remove,this),this._map=null),this},_refocusOnMap:function(t){this._map&&t&&0<t.screenX&&0<t.screenY&&this._map.getContainer().focus()}}),Ve=(A.include({addControl:function(t){return t.addTo(this),this},removeControl:function(t){return t.remove(),this},_initControlPos:function(){var i=this._controlCorners={},n="leaflet-",o=this._controlContainer=P("div",n+"control-container",this._container);function t(t,e){i[t+e]=P("div",n+t+" "+n+e,o)}t("top","left"),t("top","right"),t("bottom","left"),t("bottom","right")},_clearControlPos:function(){for(var t in this._controlCorners)T(this._controlCorners[t]);T(this._controlContainer),delete this._controlCorners,delete this._controlContainer}}),B.extend({options:{collapsed:!0,position:"topright",autoZIndex:!0,hideSingleBase:!1,sortLayers:!1,sortFunction:function(t,e,i,n){return i<n?-1:n<i?1:0}},initialize:function(t,e,i){for(var n in c(this,i),this._layerControlInputs=[],this._layers=[],this._lastZIndex=0,this._handlingClick=!1,this._preventClick=!1,t)this._addLayer(t[n],n);for(n in e)this._addLayer(e[n],n,!0)},onAdd:function(t){this._initLayout(),this._update(),(this._map=t).on("zoomend",this._checkDisabledLayers,this);for(var e=0;e<this._layers.length;e++)this._layers[e].layer.on("add remove",this._onLayerChange,this);return this._container},addTo:function(t){return B.prototype.addTo.call(this,t),this._expandIfNotCollapsed()},onRemove:function(){this._map.off("zoomend",this._checkDisabledLayers,this);for(var t=0;t<this._layers.length;t++)this._layers[t].layer.off("add remove",this._onLayerChange,this)},addBaseLayer:function(t,e){return this._addLayer(t,e),this._map?this._update():this},addOverlay:function(t,e){return this._addLayer(t,e,!0),this._map?this._update():this},removeLayer:function(t){t.off("add remove",this._onLayerChange,this);t=this._getLayer(h(t));return t&&this._layers.splice(this._layers.indexOf(t),1),this._map?this._update():this},expand:function(){M(this._container,"leaflet-control-layers-expanded"),this._section.style.height=null;var t=this._map.getSize().y-(this._container.offsetTop+50);return t<this._section.clientHeight?(M(this._section,"leaflet-control-layers-scrollbar"),this._section.style.height=t+"px"):z(this._section,"leaflet-control-layers-scrollbar"),this._checkDisabledLayers(),this},collapse:function(){return z(this._container,"leaflet-control-layers-expanded"),this},_initLayout:function(){var t="leaflet-control-layers",e=this._container=P("div",t),i=this.options.collapsed,n=(e.setAttribute("aria-haspopup",!0),Ie(e),Be(e),this._section=P("section",t+"-list")),o=(i&&(this._map.on("click",this.collapse,this),S(e,{mouseenter:this._expandSafely,mouseleave:this.collapse},this)),this._layersLink=P("a",t+"-toggle",e));o.href="#",o.title="Layers",o.setAttribute("role","button"),S(o,{keydown:function(t){13===t.keyCode&&this._expandSafely()},click:function(t){O(t),this._expandSafely()}},this),i||this.expand(),this._baseLayersList=P("div",t+"-base",n),this._separator=P("div",t+"-separator",n),this._overlaysList=P("div",t+"-overlays",n),e.appendChild(n)},_getLayer:function(t){for(var e=0;e<this._layers.length;e++)if(this._layers[e]&&h(this._layers[e].layer)===t)return this._layers[e]},_addLayer:function(t,e,i){this._map&&t.on("add remove",this._onLayerChange,this),this._layers.push({layer:t,name:e,overlay:i}),this.options.sortLayers&&this._layers.sort(a(function(t,e){return this.options.sortFunction(t.layer,e.layer,t.name,e.name)},this)),this.options.autoZIndex&&t.setZIndex&&(this._lastZIndex++,t.setZIndex(this._lastZIndex)),this._expandIfNotCollapsed()},_update:function(){if(this._container){me(this._baseLayersList),me(this._overlaysList),this._layerControlInputs=[];for(var t,e,i,n=0,o=0;o<this._layers.length;o++)i=this._layers[o],this._addItem(i),e=e||i.overlay,t=t||!i.overlay,n+=i.overlay?0:1;this.options.hideSingleBase&&(this._baseLayersList.style.display=(t=t&&1<n)?"":"none"),this._separator.style.display=e&&t?"":"none"}return this},_onLayerChange:function(t){this._handlingClick||this._update();var e=this._getLayer(h(t.target)),t=e.overlay?"add"===t.type?"overlayadd":"overlayremove":"add"===t.type?"baselayerchange":null;t&&this._map.fire(t,e)},_createRadioElement:function(t,e){t='<input type="radio" class="leaflet-control-layers-selector" name="'+t+'"'+(e?' checked="checked"':"")+"/>",e=document.createElement("div");return e.innerHTML=t,e.firstChild},_addItem:function(t){var e,i=document.createElement("label"),n=this._map.hasLayer(t.layer),n=(t.overlay?((e=document.createElement("input")).type="checkbox",e.className="leaflet-control-layers-selector",e.defaultChecked=n):e=this._createRadioElement("leaflet-base-layers_"+h(this),n),this._layerControlInputs.push(e),e.layerId=h(t.layer),S(e,"click",this._onInputClick,this),document.createElement("span")),o=(n.innerHTML=" "+t.name,document.createElement("span"));return i.appendChild(o),o.appendChild(e),o.appendChild(n),(t.overlay?this._overlaysList:this._baseLayersList).appendChild(i),this._checkDisabledLayers(),i},_onInputClick:function(){if(!this._preventClick){var t,e,i=this._layerControlInputs,n=[],o=[];this._handlingClick=!0;for(var s=i.length-1;0<=s;s--)t=i[s],e=this._getLayer(t.layerId).layer,t.checked?n.push(e):t.checked||o.push(e);for(s=0;s<o.length;s++)this._map.hasLayer(o[s])&&this._map.removeLayer(o[s]);for(s=0;s<n.length;s++)this._map.hasLayer(n[s])||this._map.addLayer(n[s]);this._handlingClick=!1,this._refocusOnMap()}},_checkDisabledLayers:function(){for(var t,e,i=this._layerControlInputs,n=this._map.getZoom(),o=i.length-1;0<=o;o--)t=i[o],e=this._getLayer(t.layerId).layer,t.disabled=void 0!==e.options.minZoom&&n<e.options.minZoom||void 0!==e.options.maxZoom&&n>e.options.maxZoom},_expandIfNotCollapsed:function(){return this._map&&!this.options.collapsed&&this.expand(),this},_expandSafely:function(){var t=this._section,e=(this._preventClick=!0,S(t,"click",O),this.expand(),this);setTimeout(function(){k(t,"click",O),e._preventClick=!1})}})),qe=B.extend({options:{position:"topleft",zoomInText:'<span aria-hidden="true">+</span>',zoomInTitle:"Zoom in",zoomOutText:'<span aria-hidden="true">&#x2212;</span>',zoomOutTitle:"Zoom out"},onAdd:function(t){var e="leaflet-control-zoom",i=P("div",e+" leaflet-bar"),n=this.options;return this._zoomInButton=this._createButton(n.zoomInText,n.zoomInTitle,e+"-in",i,this._zoomIn),this._zoomOutButton=this._createButton(n.zoomOutText,n.zoomOutTitle,e+"-out",i,this._zoomOut),this._updateDisabled(),t.on("zoomend zoomlevelschange",this._updateDisabled,this),i},onRemove:function(t){t.off("zoomend zoomlevelschange",this._updateDisabled,this)},disable:function(){return this._disabled=!0,this._updateDisabled(),this},enable:function(){return this._disabled=!1,this._updateDisabled(),this},_zoomIn:function(t){!this._disabled&&this._map._zoom<this._map.getMaxZoom()&&this._map.zoomIn(this._map.options.zoomDelta*(t.shiftKey?3:1))},_zoomOut:function(t){!this._disabled&&this._map._zoom>this._map.getMinZoom()&&this._map.zoomOut(this._map.options.zoomDelta*(t.shiftKey?3:1))},_createButton:function(t,e,i,n,o){i=P("a",i,n);return i.innerHTML=t,i.href="#",i.title=e,i.setAttribute("role","button"),i.setAttribute("aria-label",e),Ie(i),S(i,"click",Re),S(i,"click",o,this),S(i,"click",this._refocusOnMap,this),i},_updateDisabled:function(){var t=this._map,e="leaflet-disabled";z(this._zoomInButton,e),z(this._zoomOutButton,e),this._zoomInButton.setAttribute("aria-disabled","false"),this._zoomOutButton.setAttribute("aria-disabled","false"),!this._disabled&&t._zoom!==t.getMinZoom()||(M(this._zoomOutButton,e),this._zoomOutButton.setAttribute("aria-disabled","true")),!this._disabled&&t._zoom!==t.getMaxZoom()||(M(this._zoomInButton,e),this._zoomInButton.setAttribute("aria-disabled","true"))}}),Ge=(A.mergeOptions({zoomControl:!0}),A.addInitHook(function(){this.options.zoomControl&&(this.zoomControl=new qe,this.addControl(this.zoomControl))}),B.extend({options:{position:"bottomleft",maxWidth:100,metric:!0,imperial:!0},onAdd:function(t){var e="leaflet-control-scale",i=P("div",e),n=this.options;return this._addScales(n,e+"-line",i),t.on(n.updateWhenIdle?"moveend":"move",this._update,this),t.whenReady(this._update,this),i},onRemove:function(t){t.off(this.options.updateWhenIdle?"moveend":"move",this._update,this)},_addScales:function(t,e,i){t.metric&&(this._mScale=P("div",e,i)),t.imperial&&(this._iScale=P("div",e,i))},_update:function(){var t=this._map,e=t.getSize().y/2,t=t.distance(t.containerPointToLatLng([0,e]),t.containerPointToLatLng([this.options.maxWidth,e]));this._updateScales(t)},_updateScales:function(t){this.options.metric&&t&&this._updateMetric(t),this.options.imperial&&t&&this._updateImperial(t)},_updateMetric:function(t){var e=this._getRoundNum(t);this._updateScale(this._mScale,e<1e3?e+" m":e/1e3+" km",e/t)},_updateImperial:function(t){var e,i,t=3.2808399*t;5280<t?(i=this._getRoundNum(e=t/5280),this._updateScale(this._iScale,i+" mi",i/e)):(i=this._getRoundNum(t),this._updateScale(this._iScale,i+" ft",i/t))},_updateScale:function(t,e,i){t.style.width=Math.round(this.options.maxWidth*i)+"px",t.innerHTML=e},_getRoundNum:function(t){var e=Math.pow(10,(Math.floor(t)+"").length-1),t=t/e;return e*(t=10<=t?10:5<=t?5:3<=t?3:2<=t?2:1)}})),Ke=B.extend({options:{position:"bottomright",prefix:'<a href="https://leafletjs.com" title="A JavaScript library for interactive maps">'+(b.inlineSvg?'<svg aria-hidden="true" xmlns="http://www.w3.org/2000/svg" width="12" height="8" viewBox="0 0 12 8" class="leaflet-attribution-flag"><path fill="#4C7BE1" d="M0 0h12v4H0z"/><path fill="#FFD500" d="M0 4h12v3H0z"/><path fill="#E0BC00" d="M0 7h12v1H0z"/></svg> ':"")+"Leaflet</a>"},initialize:function(t){c(this,t),this._attributions={}},onAdd:function(t){for(var e in(t.attributionControl=this)._container=P("div","leaflet-control-attribution"),Ie(this._container),t._layers)t._layers[e].getAttribution&&this.addAttribution(t._layers[e].getAttribution());return this._update(),t.on("layeradd",this._addAttribution,this),this._container},onRemove:function(t){t.off("layeradd",this._addAttribution,this)},_addAttribution:function(t){t.layer.getAttribution&&(this.addAttribution(t.layer.getAttribution()),t.layer.once("remove",function(){this.removeAttribution(t.layer.getAttribution())},this))},setPrefix:function(t){return this.options.prefix=t,this._update(),this},addAttribution:function(t){return t&&(this._attributions[t]||(this._attributions[t]=0),this._attributions[t]++,this._update()),this},removeAttribution:function(t){return t&&this._attributions[t]&&(this._attributions[t]--,this._update()),this},_update:function(){if(this._map){var t,e=[];for(t in this._attributions)this._attributions[t]&&e.push(t);var i=[];this.options.prefix&&i.push(this.options.prefix),e.length&&i.push(e.join(", ")),this._container.innerHTML=i.join(' <span aria-hidden="true">|</span> ')}}}),n=(A.mergeOptions({attributionControl:!0}),A.addInitHook(function(){this.options.attributionControl&&(new Ke).addTo(this)}),B.Layers=Ve,B.Zoom=qe,B.Scale=Ge,B.Attribution=Ke,Ue.layers=function(t,e,i){return new Ve(t,e,i)},Ue.zoom=function(t){return new qe(t)},Ue.scale=function(t){return new Ge(t)},Ue.attribution=function(t){return new Ke(t)},et.extend({initialize:function(t){this._map=t},enable:function(){return this._enabled||(this._enabled=!0,this.addHooks()),this},disable:function(){return this._enabled&&(this._enabled=!1,this.removeHooks()),this},enabled:function(){return!!this._enabled}})),ft=(n.addTo=function(t,e){return t.addHandler(e,this),this},{Events:e}),Ye=b.touch?"touchstart mousedown":"mousedown",Xe=it.extend({options:{clickTolerance:3},initialize:function(t,e,i,n){c(this,n),this._element=t,this._dragStartTarget=e||t,this._preventOutline=i},enable:function(){this._enabled||(S(this._dragStartTarget,Ye,this._onDown,this),this._enabled=!0)},disable:function(){this._enabled&&(Xe._dragging===this&&this.finishDrag(!0),k(this._dragStartTarget,Ye,this._onDown,this),this._enabled=!1,this._moved=!1)},_onDown:function(t){var e,i;this._enabled&&(this._moved=!1,ve(this._element,"leaflet-zoom-anim")||(t.touches&&1!==t.touches.length?Xe._dragging===this&&this.finishDrag():Xe._dragging||t.shiftKey||1!==t.which&&1!==t.button&&!t.touches||((Xe._dragging=this)._preventOutline&&Me(this._element),Le(),re(),this._moving||(this.fire("down"),i=t.touches?t.touches[0]:t,e=Ce(this._element),this._startPoint=new p(i.clientX,i.clientY),this._startPos=Pe(this._element),this._parentScale=Ze(e),i="mousedown"===t.type,S(document,i?"mousemove":"touchmove",this._onMove,this),S(document,i?"mouseup":"touchend touchcancel",this._onUp,this)))))},_onMove:function(t){var e;this._enabled&&(t.touches&&1<t.touches.length?this._moved=!0:!(e=new p((e=t.touches&&1===t.touches.length?t.touches[0]:t).clientX,e.clientY)._subtract(this._startPoint)).x&&!e.y||Math.abs(e.x)+Math.abs(e.y)<this.options.clickTolerance||(e.x/=this._parentScale.x,e.y/=this._parentScale.y,O(t),this._moved||(this.fire("dragstart"),this._moved=!0,M(document.body,"leaflet-dragging"),this._lastTarget=t.target||t.srcElement,window.SVGElementInstance&&this._lastTarget instanceof window.SVGElementInstance&&(this._lastTarget=this._lastTarget.correspondingUseElement),M(this._lastTarget,"leaflet-drag-target")),this._newPos=this._startPos.add(e),this._moving=!0,this._lastEvent=t,this._updatePosition()))},_updatePosition:function(){var t={originalEvent:this._lastEvent};this.fire("predrag",t),Z(this._element,this._newPos),this.fire("drag",t)},_onUp:function(){this._enabled&&this.finishDrag()},finishDrag:function(t){z(document.body,"leaflet-dragging"),this._lastTarget&&(z(this._lastTarget,"leaflet-drag-target"),this._lastTarget=null),k(document,"mousemove touchmove",this._onMove,this),k(document,"mouseup touchend touchcancel",this._onUp,this),Te(),ae();var e=this._moved&&this._moving;this._moving=!1,Xe._dragging=!1,e&&this.fire("dragend",{noInertia:t,distance:this._newPos.distanceTo(this._startPos)})}});function Je(t,e,i){for(var n,o,s,r,a,h,l,u=[1,4,2,8],c=0,d=t.length;c<d;c++)t[c]._code=si(t[c],e);for(s=0;s<4;s++){for(h=u[s],n=[],c=0,o=(d=t.length)-1;c<d;o=c++)r=t[c],a=t[o],r._code&h?a._code&h||((l=oi(a,r,h,e,i))._code=si(l,e),n.push(l)):(a._code&h&&((l=oi(a,r,h,e,i))._code=si(l,e),n.push(l)),n.push(r));t=n}return t}function $e(t,e){var i,n,o,s,r,a,h;if(!t||0===t.length)throw new Error("latlngs not passed");I(t)||(console.warn("latlngs are not flat! Only the first ring will be used"),t=t[0]);for(var l=w([0,0]),u=g(t),c=(u.getNorthWest().distanceTo(u.getSouthWest())*u.getNorthEast().distanceTo(u.getNorthWest())<1700&&(l=Qe(t)),t.length),d=[],_=0;_<c;_++){var p=w(t[_]);d.push(e.project(w([p.lat-l.lat,p.lng-l.lng])))}for(_=r=a=h=0,i=c-1;_<c;i=_++)n=d[_],o=d[i],s=n.y*o.x-o.y*n.x,a+=(n.x+o.x)*s,h+=(n.y+o.y)*s,r+=3*s;u=0===r?d[0]:[a/r,h/r],u=e.unproject(m(u));return w([u.lat+l.lat,u.lng+l.lng])}function Qe(t){for(var e=0,i=0,n=0,o=0;o<t.length;o++){var s=w(t[o]);e+=s.lat,i+=s.lng,n++}return w([e/n,i/n])}var ti,gt={__proto__:null,clipPolygon:Je,polygonCenter:$e,centroid:Qe};function ei(t,e){if(e&&t.length){var i=t=function(t,e){for(var i=[t[0]],n=1,o=0,s=t.length;n<s;n++)(function(t,e){var i=e.x-t.x,e=e.y-t.y;return i*i+e*e})(t[n],t[o])>e&&(i.push(t[n]),o=n);o<s-1&&i.push(t[s-1]);return i}(t,e=e*e),n=i.length,o=new(typeof Uint8Array!=void 0+""?Uint8Array:Array)(n);o[0]=o[n-1]=1,function t(e,i,n,o,s){var r,a,h,l=0;for(a=o+1;a<=s-1;a++)h=ri(e[a],e[o],e[s],!0),l<h&&(r=a,l=h);n<l&&(i[r]=1,t(e,i,n,o,r),t(e,i,n,r,s))}(i,o,e,0,n-1);var s,r=[];for(s=0;s<n;s++)o[s]&&r.push(i[s]);return r}return t.slice()}function ii(t,e,i){return Math.sqrt(ri(t,e,i,!0))}function ni(t,e,i,n,o){var s,r,a,h=n?ti:si(t,i),l=si(e,i);for(ti=l;;){if(!(h|l))return[t,e];if(h&l)return!1;a=si(r=oi(t,e,s=h||l,i,o),i),s===h?(t=r,h=a):(e=r,l=a)}}function oi(t,e,i,n,o){var s,r,a=e.x-t.x,e=e.y-t.y,h=n.min,n=n.max;return 8&i?(s=t.x+a*(n.y-t.y)/e,r=n.y):4&i?(s=t.x+a*(h.y-t.y)/e,r=h.y):2&i?(s=n.x,r=t.y+e*(n.x-t.x)/a):1&i&&(s=h.x,r=t.y+e*(h.x-t.x)/a),new p(s,r,o)}function si(t,e){var i=0;return t.x<e.min.x?i|=1:t.x>e.max.x&&(i|=2),t.y<e.min.y?i|=4:t.y>e.max.y&&(i|=8),i}function ri(t,e,i,n){var o=e.x,e=e.y,s=i.x-o,r=i.y-e,a=s*s+r*r;return 0<a&&(1<(a=((t.x-o)*s+(t.y-e)*r)/a)?(o=i.x,e=i.y):0<a&&(o+=s*a,e+=r*a)),s=t.x-o,r=t.y-e,n?s*s+r*r:new p(o,e)}function I(t){return!d(t[0])||"object"!=typeof t[0][0]&&void 0!==t[0][0]}function ai(t){return console.warn("Deprecated use of _flat, please use L.LineUtil.isFlat instead."),I(t)}function hi(t,e){var i,n,o,s,r,a;if(!t||0===t.length)throw new Error("latlngs not passed");I(t)||(console.warn("latlngs are not flat! Only the first ring will be used"),t=t[0]);for(var h=w([0,0]),l=g(t),u=(l.getNorthWest().distanceTo(l.getSouthWest())*l.getNorthEast().distanceTo(l.getNorthWest())<1700&&(h=Qe(t)),t.length),c=[],d=0;d<u;d++){var _=w(t[d]);c.push(e.project(w([_.lat-h.lat,_.lng-h.lng])))}for(i=d=0;d<u-1;d++)i+=c[d].distanceTo(c[d+1])/2;if(0===i)a=c[0];else for(n=d=0;d<u-1;d++)if(o=c[d],s=c[d+1],i<(n+=r=o.distanceTo(s))){a=[s.x-(r=(n-i)/r)*(s.x-o.x),s.y-r*(s.y-o.y)];break}l=e.unproject(m(a));return w([l.lat+h.lat,l.lng+h.lng])}var vt={__proto__:null,simplify:ei,pointToSegmentDistance:ii,closestPointOnSegment:function(t,e,i){return ri(t,e,i)},clipSegment:ni,_getEdgeIntersection:oi,_getBitCode:si,_sqClosestPointOnSegment:ri,isFlat:I,_flat:ai,polylineCenter:hi},yt={project:function(t){return new p(t.lng,t.lat)},unproject:function(t){return new v(t.y,t.x)},bounds:new f([-180,-90],[180,90])},xt={R:6378137,R_MINOR:6356752.314245179,bounds:new f([-20037508.34279,-15496570.73972],[20037508.34279,18764656.23138]),project:function(t){var e=Math.PI/180,i=this.R,n=t.lat*e,o=this.R_MINOR/i,o=Math.sqrt(1-o*o),s=o*Math.sin(n),s=Math.tan(Math.PI/4-n/2)/Math.pow((1-s)/(1+s),o/2),n=-i*Math.log(Math.max(s,1e-10));return new p(t.lng*e*i,n)},unproject:function(t){for(var e,i=180/Math.PI,n=this.R,o=this.R_MINOR/n,s=Math.sqrt(1-o*o),r=Math.exp(-t.y/n),a=Math.PI/2-2*Math.atan(r),h=0,l=.1;h<15&&1e-7<Math.abs(l);h++)e=s*Math.sin(a),e=Math.pow((1-e)/(1+e),s/2),a+=l=Math.PI/2-2*Math.atan(r*e)-a;return new v(a*i,t.x*i/n)}},wt={__proto__:null,LonLat:yt,Mercator:xt,SphericalMercator:rt},Pt=l({},st,{code:"EPSG:3395",projection:xt,transformation:ht(bt=.5/(Math.PI*xt.R),.5,-bt,.5)}),li=l({},st,{code:"EPSG:4326",projection:yt,transformation:ht(1/180,1,-1/180,.5)}),Lt=l({},ot,{projection:yt,transformation:ht(1,0,-1,0),scale:function(t){return Math.pow(2,t)},zoom:function(t){return Math.log(t)/Math.LN2},distance:function(t,e){var i=e.lng-t.lng,e=e.lat-t.lat;return Math.sqrt(i*i+e*e)},infinite:!0}),o=(ot.Earth=st,ot.EPSG3395=Pt,ot.EPSG3857=lt,ot.EPSG900913=ut,ot.EPSG4326=li,ot.Simple=Lt,it.extend({options:{pane:"overlayPane",attribution:null,bubblingMouseEvents:!0},addTo:function(t){return t.addLayer(this),this},remove:function(){return this.removeFrom(this._map||this._mapToAdd)},removeFrom:function(t){return t&&t.removeLayer(this),this},getPane:function(t){return this._map.getPane(t?this.options[t]||t:this.options.pane)},addInteractiveTarget:function(t){return this._map._targets[h(t)]=this},removeInteractiveTarget:function(t){return delete this._map._targets[h(t)],this},getAttribution:function(){return this.options.attribution},_layerAdd:function(t){var e,i=t.target;i.hasLayer(this)&&(this._map=i,this._zoomAnimated=i._zoomAnimated,this.getEvents&&(e=this.getEvents(),i.on(e,this),this.once("remove",function(){i.off(e,this)},this)),this.onAdd(i),this.fire("add"),i.fire("layeradd",{layer:this}))}})),ui=(A.include({addLayer:function(t){var e;if(t._layerAdd)return e=h(t),this._layers[e]||((this._layers[e]=t)._mapToAdd=this,t.beforeAdd&&t.beforeAdd(this),this.whenReady(t._layerAdd,t)),this;throw new Error("The provided object is not a Layer.")},removeLayer:function(t){var e=h(t);return this._layers[e]&&(this._loaded&&t.onRemove(this),delete this._layers[e],this._loaded&&(this.fire("layerremove",{layer:t}),t.fire("remove")),t._map=t._mapToAdd=null),this},hasLayer:function(t){return h(t)in this._layers},eachLayer:function(t,e){for(var i in this._layers)t.call(e,this._layers[i]);return this},_addLayers:function(t){for(var e=0,i=(t=t?d(t)?t:[t]:[]).length;e<i;e++)this.addLayer(t[e])},_addZoomLimit:function(t){isNaN(t.options.maxZoom)&&isNaN(t.options.minZoom)||(this._zoomBoundLayers[h(t)]=t,this._updateZoomLevels())},_removeZoomLimit:function(t){t=h(t);this._zoomBoundLayers[t]&&(delete this._zoomBoundLayers[t],this._updateZoomLevels())},_updateZoomLevels:function(){var t,e=1/0,i=-1/0,n=this._getZoomSpan();for(t in this._zoomBoundLayers)var o=this._zoomBoundLayers[t].options,e=void 0===o.minZoom?e:Math.min(e,o.minZoom),i=void 0===o.maxZoom?i:Math.max(i,o.maxZoom);this._layersMaxZoom=i===-1/0?void 0:i,this._layersMinZoom=e===1/0?void 0:e,n!==this._getZoomSpan()&&this.fire("zoomlevelschange"),void 0===this.options.maxZoom&&this._layersMaxZoom&&this.getZoom()>this._layersMaxZoom&&this.setZoom(this._layersMaxZoom),void 0===this.options.minZoom&&this._layersMinZoom&&this.getZoom()<this._layersMinZoom&&this.setZoom(this._layersMinZoom)}}),o.extend({initialize:function(t,e){var i,n;if(c(this,e),this._layers={},t)for(i=0,n=t.length;i<n;i++)this.addLayer(t[i])},addLayer:function(t){var e=this.getLayerId(t);return this._layers[e]=t,this._map&&this._map.addLayer(t),this},removeLayer:function(t){t=t in this._layers?t:this.getLayerId(t);return this._map&&this._layers[t]&&this._map.removeLayer(this._layers[t]),delete this._layers[t],this},hasLayer:function(t){return("number"==typeof t?t:this.getLayerId(t))in this._layers},clearLayers:function(){return this.eachLayer(this.removeLayer,this)},invoke:function(t){var e,i,n=Array.prototype.slice.call(arguments,1);for(e in this._layers)(i=this._layers[e])[t]&&i[t].apply(i,n);return this},onAdd:function(t){this.eachLayer(t.addLayer,t)},onRemove:function(t){this.eachLayer(t.removeLayer,t)},eachLayer:function(t,e){for(var i in this._layers)t.call(e,this._layers[i]);return this},getLayer:function(t){return this._layers[t]},getLayers:function(){var t=[];return this.eachLayer(t.push,t),t},setZIndex:function(t){return this.invoke("setZIndex",t)},getLayerId:h})),ci=ui.extend({addLayer:function(t){return this.hasLayer(t)?this:(t.addEventParent(this),ui.prototype.addLayer.call(this,t),this.fire("layeradd",{layer:t}))},removeLayer:function(t){return this.hasLayer(t)?((t=t in this._layers?this._layers[t]:t).removeEventParent(this),ui.prototype.removeLayer.call(this,t),this.fire("layerremove",{layer:t})):this},setStyle:function(t){return this.invoke("setStyle",t)},bringToFront:function(){return this.invoke("bringToFront")},bringToBack:function(){return this.invoke("bringToBack")},getBounds:function(){var t,e=new s;for(t in this._layers){var i=this._layers[t];e.extend(i.getBounds?i.getBounds():i.getLatLng())}return e}}),di=et.extend({options:{popupAnchor:[0,0],tooltipAnchor:[0,0],crossOrigin:!1},initialize:function(t){c(this,t)},createIcon:function(t){return this._createIcon("icon",t)},createShadow:function(t){return this._createIcon("shadow",t)},_createIcon:function(t,e){var i=this._getIconUrl(t);if(i)return i=this._createImg(i,e&&"IMG"===e.tagName?e:null),this._setIconStyles(i,t),!this.options.crossOrigin&&""!==this.options.crossOrigin||(i.crossOrigin=!0===this.options.crossOrigin?"":this.options.crossOrigin),i;if("icon"===t)throw new Error("iconUrl not set in Icon options (see the docs).");return null},_setIconStyles:function(t,e){var i=this.options,n=i[e+"Size"],n=m(n="number"==typeof n?[n,n]:n),o=m("shadow"===e&&i.shadowAnchor||i.iconAnchor||n&&n.divideBy(2,!0));t.className="leaflet-marker-"+e+" "+(i.className||""),o&&(t.style.marginLeft=-o.x+"px",t.style.marginTop=-o.y+"px"),n&&(t.style.width=n.x+"px",t.style.height=n.y+"px")},_createImg:function(t,e){return(e=e||document.createElement("img")).src=t,e},_getIconUrl:function(t){return b.retina&&this.options[t+"RetinaUrl"]||this.options[t+"Url"]}});var _i=di.extend({options:{iconUrl:"marker-icon.png",iconRetinaUrl:"marker-icon-2x.png",shadowUrl:"marker-shadow.png",iconSize:[25,41],iconAnchor:[12,41],popupAnchor:[1,-34],tooltipAnchor:[16,-28],shadowSize:[41,41]},_getIconUrl:function(t){return"string"!=typeof _i.imagePath&&(_i.imagePath=this._detectIconPath()),(this.options.imagePath||_i.imagePath)+di.prototype._getIconUrl.call(this,t)},_stripUrl:function(t){function e(t,e,i){return(e=e.exec(t))&&e[i]}return(t=e(t,/^url\((['"])?(.+)\1\)$/,2))&&e(t,/^(.*)marker-icon\.png$/,1)},_detectIconPath:function(){var t=P("div","leaflet-default-icon-path",document.body),e=pe(t,"background-image")||pe(t,"backgroundImage");return document.body.removeChild(t),(e=this._stripUrl(e))?e:(t=document.querySelector('link[href$="leaflet.css"]'))?t.href.substring(0,t.href.length-"leaflet.css".length-1):""}}),pi=n.extend({initialize:function(t){this._marker=t},addHooks:function(){var t=this._marker._icon;this._draggable||(this._draggable=new Xe(t,t,!0)),this._draggable.on({dragstart:this._onDragStart,predrag:this._onPreDrag,drag:this._onDrag,dragend:this._onDragEnd},this).enable(),M(t,"leaflet-marker-draggable")},removeHooks:function(){this._draggable.off({dragstart:this._onDragStart,predrag:this._onPreDrag,drag:this._onDrag,dragend:this._onDragEnd},this).disable(),this._marker._icon&&z(this._marker._icon,"leaflet-marker-draggable")},moved:function(){return this._draggable&&this._draggable._moved},_adjustPan:function(t){var e=this._marker,i=e._map,n=this._marker.options.autoPanSpeed,o=this._marker.options.autoPanPadding,s=Pe(e._icon),r=i.getPixelBounds(),a=i.getPixelOrigin(),a=_(r.min._subtract(a).add(o),r.max._subtract(a).subtract(o));a.contains(s)||(o=m((Math.max(a.max.x,s.x)-a.max.x)/(r.max.x-a.max.x)-(Math.min(a.min.x,s.x)-a.min.x)/(r.min.x-a.min.x),(Math.max(a.max.y,s.y)-a.max.y)/(r.max.y-a.max.y)-(Math.min(a.min.y,s.y)-a.min.y)/(r.min.y-a.min.y)).multiplyBy(n),i.panBy(o,{animate:!1}),this._draggable._newPos._add(o),this._draggable._startPos._add(o),Z(e._icon,this._draggable._newPos),this._onDrag(t),this._panRequest=x(this._adjustPan.bind(this,t)))},_onDragStart:function(){this._oldLatLng=this._marker.getLatLng(),this._marker.closePopup&&this._marker.closePopup(),this._marker.fire("movestart").fire("dragstart")},_onPreDrag:function(t){this._marker.options.autoPan&&(r(this._panRequest),this._panRequest=x(this._adjustPan.bind(this,t)))},_onDrag:function(t){var e=this._marker,i=e._shadow,n=Pe(e._icon),o=e._map.layerPointToLatLng(n);i&&Z(i,n),e._latlng=o,t.latlng=o,t.oldLatLng=this._oldLatLng,e.fire("move",t).fire("drag",t)},_onDragEnd:function(t){r(this._panRequest),delete this._oldLatLng,this._marker.fire("moveend").fire("dragend",t)}}),mi=o.extend({options:{icon:new _i,interactive:!0,keyboard:!0,title:"",alt:"Marker",zIndexOffset:0,opacity:1,riseOnHover:!1,riseOffset:250,pane:"markerPane",shadowPane:"shadowPane",bubblingMouseEvents:!1,autoPanOnFocus:!0,draggable:!1,autoPan:!1,autoPanPadding:[50,50],autoPanSpeed:10},initialize:function(t,e){c(this,e),this._latlng=w(t)},onAdd:function(t){this._zoomAnimated=this._zoomAnimated&&t.options.markerZoomAnimation,this._zoomAnimated&&t.on("zoomanim",this._animateZoom,this),this._initIcon(),this.update()},onRemove:function(t){this.dragging&&this.dragging.enabled()&&(this.options.draggable=!0,this.dragging.removeHooks()),delete this.dragging,this._zoomAnimated&&t.off("zoomanim",this._animateZoom,this),this._removeIcon(),this._removeShadow()},getEvents:function(){return{zoom:this.update,viewreset:this.update}},getLatLng:function(){return this._latlng},setLatLng:function(t){var e=this._latlng;return this._latlng=w(t),this.update(),this.fire("move",{oldLatLng:e,latlng:this._latlng})},setZIndexOffset:function(t){return this.options.zIndexOffset=t,this.update()},getIcon:function(){return this.options.icon},setIcon:function(t){return this.options.icon=t,this._map&&(this._initIcon(),this.update()),this._popup&&this.bindPopup(this._popup,this._popup.options),this},getElement:function(){return this._icon},update:function(){var t;return this._icon&&this._map&&(t=this._map.latLngToLayerPoint(this._latlng).round(),this._setPos(t)),this},_initIcon:function(){var t=this.options,e="leaflet-zoom-"+(this._zoomAnimated?"animated":"hide"),i=t.icon.createIcon(this._icon),n=!1,i=(i!==this._icon&&(this._icon&&this._removeIcon(),n=!0,t.title&&(i.title=t.title),"IMG"===i.tagName&&(i.alt=t.alt||"")),M(i,e),t.keyboard&&(i.tabIndex="0",i.setAttribute("role","button")),this._icon=i,t.riseOnHover&&this.on({mouseover:this._bringToFront,mouseout:this._resetZIndex}),this.options.autoPanOnFocus&&S(i,"focus",this._panOnFocus,this),t.icon.createShadow(this._shadow)),o=!1;i!==this._shadow&&(this._removeShadow(),o=!0),i&&(M(i,e),i.alt=""),this._shadow=i,t.opacity<1&&this._updateOpacity(),n&&this.getPane().appendChild(this._icon),this._initInteraction(),i&&o&&this.getPane(t.shadowPane).appendChild(this._shadow)},_removeIcon:function(){this.options.riseOnHover&&this.off({mouseover:this._bringToFront,mouseout:this._resetZIndex}),this.options.autoPanOnFocus&&k(this._icon,"focus",this._panOnFocus,this),T(this._icon),this.removeInteractiveTarget(this._icon),this._icon=null},_removeShadow:function(){this._shadow&&T(this._shadow),this._shadow=null},_setPos:function(t){this._icon&&Z(this._icon,t),this._shadow&&Z(this._shadow,t),this._zIndex=t.y+this.options.zIndexOffset,this._resetZIndex()},_updateZIndex:function(t){this._icon&&(this._icon.style.zIndex=this._zIndex+t)},_animateZoom:function(t){t=this._map._latLngToNewLayerPoint(this._latlng,t.zoom,t.center).round();this._setPos(t)},_initInteraction:function(){var t;this.options.interactive&&(M(this._icon,"leaflet-interactive"),this.addInteractiveTarget(this._icon),pi&&(t=this.options.draggable,this.dragging&&(t=this.dragging.enabled(),this.dragging.disable()),this.dragging=new pi(this),t&&this.dragging.enable()))},setOpacity:function(t){return this.options.opacity=t,this._map&&this._updateOpacity(),this},_updateOpacity:function(){var t=this.options.opacity;this._icon&&C(this._icon,t),this._shadow&&C(this._shadow,t)},_bringToFront:function(){this._updateZIndex(this.options.riseOffset)},_resetZIndex:function(){this._updateZIndex(0)},_panOnFocus:function(){var t,e,i=this._map;i&&(t=(e=this.options.icon.options).iconSize?m(e.iconSize):m(0,0),e=e.iconAnchor?m(e.iconAnchor):m(0,0),i.panInside(this._latlng,{paddingTopLeft:e,paddingBottomRight:t.subtract(e)}))},_getPopupAnchor:function(){return this.options.icon.options.popupAnchor},_getTooltipAnchor:function(){return this.options.icon.options.tooltipAnchor}});var fi=o.extend({options:{stroke:!0,color:"#3388ff",weight:3,opacity:1,lineCap:"round",lineJoin:"round",dashArray:null,dashOffset:null,fill:!1,fillColor:null,fillOpacity:.2,fillRule:"evenodd",interactive:!0,bubblingMouseEvents:!0},beforeAdd:function(t){this._renderer=t.getRenderer(this)},onAdd:function(){this._renderer._initPath(this),this._reset(),this._renderer._addPath(this)},onRemove:function(){this._renderer._removePath(this)},redraw:function(){return this._map&&this._renderer._updatePath(this),this},setStyle:function(t){return c(this,t),this._renderer&&(this._renderer._updateStyle(this),this.options.stroke&&t&&Object.prototype.hasOwnProperty.call(t,"weight")&&this._updateBounds()),this},bringToFront:function(){return this._renderer&&this._renderer._bringToFront(this),this},bringToBack:function(){return this._renderer&&this._renderer._bringToBack(this),this},getElement:function(){return this._path},_reset:function(){this._project(),this._update()},_clickTolerance:function(){return(this.options.stroke?this.options.weight/2:0)+(this._renderer.options.tolerance||0)}}),gi=fi.extend({options:{fill:!0,radius:10},initialize:function(t,e){c(this,e),this._latlng=w(t),this._radius=this.options.radius},setLatLng:function(t){var e=this._latlng;return this._latlng=w(t),this.redraw(),this.fire("move",{oldLatLng:e,latlng:this._latlng})},getLatLng:function(){return this._latlng},setRadius:function(t){return this.options.radius=this._radius=t,this.redraw()},getRadius:function(){return this._radius},setStyle:function(t){var e=t&&t.radius||this._radius;return fi.prototype.setStyle.call(this,t),this.setRadius(e),this},_project:function(){this._point=this._map.latLngToLayerPoint(this._latlng),this._updateBounds()},_updateBounds:function(){var t=this._radius,e=this._radiusY||t,i=this._clickTolerance(),t=[t+i,e+i];this._pxBounds=new f(this._point.subtract(t),this._point.add(t))},_update:function(){this._map&&this._updatePath()},_updatePath:function(){this._renderer._updateCircle(this)},_empty:function(){return this._radius&&!this._renderer._bounds.intersects(this._pxBounds)},_containsPoint:function(t){return t.distanceTo(this._point)<=this._radius+this._clickTolerance()}});var vi=gi.extend({initialize:function(t,e,i){if(c(this,e="number"==typeof e?l({},i,{radius:e}):e),this._latlng=w(t),isNaN(this.options.radius))throw new Error("Circle radius cannot be NaN");this._mRadius=this.options.radius},setRadius:function(t){return this._mRadius=t,this.redraw()},getRadius:function(){return this._mRadius},getBounds:function(){var t=[this._radius,this._radiusY||this._radius];return new s(this._map.layerPointToLatLng(this._point.subtract(t)),this._map.layerPointToLatLng(this._point.add(t)))},setStyle:fi.prototype.setStyle,_project:function(){var t,e,i,n,o,s=this._latlng.lng,r=this._latlng.lat,a=this._map,h=a.options.crs;h.distance===st.distance?(n=Math.PI/180,o=this._mRadius/st.R/n,t=a.project([r+o,s]),e=a.project([r-o,s]),e=t.add(e).divideBy(2),i=a.unproject(e).lat,n=Math.acos((Math.cos(o*n)-Math.sin(r*n)*Math.sin(i*n))/(Math.cos(r*n)*Math.cos(i*n)))/n,!isNaN(n)&&0!==n||(n=o/Math.cos(Math.PI/180*r)),this._point=e.subtract(a.getPixelOrigin()),this._radius=isNaN(n)?0:e.x-a.project([i,s-n]).x,this._radiusY=e.y-t.y):(o=h.unproject(h.project(this._latlng).subtract([this._mRadius,0])),this._point=a.latLngToLayerPoint(this._latlng),this._radius=this._point.x-a.latLngToLayerPoint(o).x),this._updateBounds()}});var yi=fi.extend({options:{smoothFactor:1,noClip:!1},initialize:function(t,e){c(this,e),this._setLatLngs(t)},getLatLngs:function(){return this._latlngs},setLatLngs:function(t){return this._setLatLngs(t),this.redraw()},isEmpty:function(){return!this._latlngs.length},closestLayerPoint:function(t){for(var e=1/0,i=null,n=ri,o=0,s=this._parts.length;o<s;o++)for(var r=this._parts[o],a=1,h=r.length;a<h;a++){var l,u,c=n(t,l=r[a-1],u=r[a],!0);c<e&&(e=c,i=n(t,l,u))}return i&&(i.distance=Math.sqrt(e)),i},getCenter:function(){if(this._map)return hi(this._defaultShape(),this._map.options.crs);throw new Error("Must add layer to map before using getCenter()")},getBounds:function(){return this._bounds},addLatLng:function(t,e){return e=e||this._defaultShape(),t=w(t),e.push(t),this._bounds.extend(t),this.redraw()},_setLatLngs:function(t){this._bounds=new s,this._latlngs=this._convertLatLngs(t)},_defaultShape:function(){return I(this._latlngs)?this._latlngs:this._latlngs[0]},_convertLatLngs:function(t){for(var e=[],i=I(t),n=0,o=t.length;n<o;n++)i?(e[n]=w(t[n]),this._bounds.extend(e[n])):e[n]=this._convertLatLngs(t[n]);return e},_project:function(){var t=new f;this._rings=[],this._projectLatlngs(this._latlngs,this._rings,t),this._bounds.isValid()&&t.isValid()&&(this._rawPxBounds=t,this._updateBounds())},_updateBounds:function(){var t=this._clickTolerance(),t=new p(t,t);this._rawPxBounds&&(this._pxBounds=new f([this._rawPxBounds.min.subtract(t),this._rawPxBounds.max.add(t)]))},_projectLatlngs:function(t,e,i){var n,o,s=t[0]instanceof v,r=t.length;if(s){for(o=[],n=0;n<r;n++)o[n]=this._map.latLngToLayerPoint(t[n]),i.extend(o[n]);e.push(o)}else for(n=0;n<r;n++)this._projectLatlngs(t[n],e,i)},_clipPoints:function(){var t=this._renderer._bounds;if(this._parts=[],this._pxBounds&&this._pxBounds.intersects(t))if(this.options.noClip)this._parts=this._rings;else for(var e,i,n,o,s=this._parts,r=0,a=0,h=this._rings.length;r<h;r++)for(e=0,i=(o=this._rings[r]).length;e<i-1;e++)(n=ni(o[e],o[e+1],t,e,!0))&&(s[a]=s[a]||[],s[a].push(n[0]),n[1]===o[e+1]&&e!==i-2||(s[a].push(n[1]),a++))},_simplifyPoints:function(){for(var t=this._parts,e=this.options.smoothFactor,i=0,n=t.length;i<n;i++)t[i]=ei(t[i],e)},_update:function(){this._map&&(this._clipPoints(),this._simplifyPoints(),this._updatePath())},_updatePath:function(){this._renderer._updatePoly(this)},_containsPoint:function(t,e){var i,n,o,s,r,a,h=this._clickTolerance();if(this._pxBounds&&this._pxBounds.contains(t))for(i=0,s=this._parts.length;i<s;i++)for(n=0,o=(r=(a=this._parts[i]).length)-1;n<r;o=n++)if((e||0!==n)&&ii(t,a[o],a[n])<=h)return!0;return!1}});yi._flat=ai;var xi=yi.extend({options:{fill:!0},isEmpty:function(){return!this._latlngs.length||!this._latlngs[0].length},getCenter:function(){if(this._map)return $e(this._defaultShape(),this._map.options.crs);throw new Error("Must add layer to map before using getCenter()")},_convertLatLngs:function(t){var t=yi.prototype._convertLatLngs.call(this,t),e=t.length;return 2<=e&&t[0]instanceof v&&t[0].equals(t[e-1])&&t.pop(),t},_setLatLngs:function(t){yi.prototype._setLatLngs.call(this,t),I(this._latlngs)&&(this._latlngs=[this._latlngs])},_defaultShape:function(){return(I(this._latlngs[0])?this._latlngs:this._latlngs[0])[0]},_clipPoints:function(){var t=this._renderer._bounds,e=this.options.weight,e=new p(e,e),t=new f(t.min.subtract(e),t.max.add(e));if(this._parts=[],this._pxBounds&&this._pxBounds.intersects(t))if(this.options.noClip)this._parts=this._rings;else for(var i,n=0,o=this._rings.length;n<o;n++)(i=Je(this._rings[n],t,!0)).length&&this._parts.push(i)},_updatePath:function(){this._renderer._updatePoly(this,!0)},_containsPoint:function(t){var e,i,n,o,s,r,a,h,l=!1;if(!this._pxBounds||!this._pxBounds.contains(t))return!1;for(o=0,a=this._parts.length;o<a;o++)for(s=0,r=(h=(e=this._parts[o]).length)-1;s<h;r=s++)i=e[s],n=e[r],i.y>t.y!=n.y>t.y&&t.x<(n.x-i.x)*(t.y-i.y)/(n.y-i.y)+i.x&&(l=!l);return l||yi.prototype._containsPoint.call(this,t,!0)}});var wi=ci.extend({initialize:function(t,e){c(this,e),this._layers={},t&&this.addData(t)},addData:function(t){var e,i,n,o=d(t)?t:t.features;if(o){for(e=0,i=o.length;e<i;e++)((n=o[e]).geometries||n.geometry||n.features||n.coordinates)&&this.addData(n);return this}var s,r=this.options;return(!r.filter||r.filter(t))&&(s=bi(t,r))?(s.feature=Zi(t),s.defaultOptions=s.options,this.resetStyle(s),r.onEachFeature&&r.onEachFeature(t,s),this.addLayer(s)):this},resetStyle:function(t){return void 0===t?this.eachLayer(this.resetStyle,this):(t.options=l({},t.defaultOptions),this._setLayerStyle(t,this.options.style),this)},setStyle:function(e){return this.eachLayer(function(t){this._setLayerStyle(t,e)},this)},_setLayerStyle:function(t,e){t.setStyle&&("function"==typeof e&&(e=e(t.feature)),t.setStyle(e))}});function bi(t,e){var i,n,o,s,r="Feature"===t.type?t.geometry:t,a=r?r.coordinates:null,h=[],l=e&&e.pointToLayer,u=e&&e.coordsToLatLng||Li;if(!a&&!r)return null;switch(r.type){case"Point":return Pi(l,t,i=u(a),e);case"MultiPoint":for(o=0,s=a.length;o<s;o++)i=u(a[o]),h.push(Pi(l,t,i,e));return new ci(h);case"LineString":case"MultiLineString":return n=Ti(a,"LineString"===r.type?0:1,u),new yi(n,e);case"Polygon":case"MultiPolygon":return n=Ti(a,"Polygon"===r.type?1:2,u),new xi(n,e);case"GeometryCollection":for(o=0,s=r.geometries.length;o<s;o++){var c=bi({geometry:r.geometries[o],type:"Feature",properties:t.properties},e);c&&h.push(c)}return new ci(h);case"FeatureCollection":for(o=0,s=r.features.length;o<s;o++){var d=bi(r.features[o],e);d&&h.push(d)}return new ci(h);default:throw new Error("Invalid GeoJSON object.")}}function Pi(t,e,i,n){return t?t(e,i):new mi(i,n&&n.markersInheritOptions&&n)}function Li(t){return new v(t[1],t[0],t[2])}function Ti(t,e,i){for(var n,o=[],s=0,r=t.length;s<r;s++)n=e?Ti(t[s],e-1,i):(i||Li)(t[s]),o.push(n);return o}function Mi(t,e){return void 0!==(t=w(t)).alt?[i(t.lng,e),i(t.lat,e),i(t.alt,e)]:[i(t.lng,e),i(t.lat,e)]}function zi(t,e,i,n){for(var o=[],s=0,r=t.length;s<r;s++)o.push(e?zi(t[s],I(t[s])?0:e-1,i,n):Mi(t[s],n));return!e&&i&&0<o.length&&o.push(o[0].slice()),o}function Ci(t,e){return t.feature?l({},t.feature,{geometry:e}):Zi(e)}function Zi(t){return"Feature"===t.type||"FeatureCollection"===t.type?t:{type:"Feature",properties:{},geometry:t}}Tt={toGeoJSON:function(t){return Ci(this,{type:"Point",coordinates:Mi(this.getLatLng(),t)})}};function Si(t,e){return new wi(t,e)}mi.include(Tt),vi.include(Tt),gi.include(Tt),yi.include({toGeoJSON:function(t){var e=!I(this._latlngs);return Ci(this,{type:(e?"Multi":"")+"LineString",coordinates:zi(this._latlngs,e?1:0,!1,t)})}}),xi.include({toGeoJSON:function(t){var e=!I(this._latlngs),i=e&&!I(this._latlngs[0]),t=zi(this._latlngs,i?2:e?1:0,!0,t);return Ci(this,{type:(i?"Multi":"")+"Polygon",coordinates:t=e?t:[t]})}}),ui.include({toMultiPoint:function(e){var i=[];return this.eachLayer(function(t){i.push(t.toGeoJSON(e).geometry.coordinates)}),Ci(this,{type:"MultiPoint",coordinates:i})},toGeoJSON:function(e){var i,n,t=this.feature&&this.feature.geometry&&this.feature.geometry.type;return"MultiPoint"===t?this.toMultiPoint(e):(i="GeometryCollection"===t,n=[],this.eachLayer(function(t){t.toGeoJSON&&(t=t.toGeoJSON(e),i?n.push(t.geometry):"FeatureCollection"===(t=Zi(t)).type?n.push.apply(n,t.features):n.push(t))}),i?Ci(this,{geometries:n,type:"GeometryCollection"}):{type:"FeatureCollection",features:n})}});var Mt=Si,Ei=o.extend({options:{opacity:1,alt:"",interactive:!1,crossOrigin:!1,errorOverlayUrl:"",zIndex:1,className:""},initialize:function(t,e,i){this._url=t,this._bounds=g(e),c(this,i)},onAdd:function(){this._image||(this._initImage(),this.options.opacity<1&&this._updateOpacity()),this.options.interactive&&(M(this._image,"leaflet-interactive"),this.addInteractiveTarget(this._image)),this.getPane().appendChild(this._image),this._reset()},onRemove:function(){T(this._image),this.options.interactive&&this.removeInteractiveTarget(this._image)},setOpacity:function(t){return this.options.opacity=t,this._image&&this._updateOpacity(),this},setStyle:function(t){return t.opacity&&this.setOpacity(t.opacity),this},bringToFront:function(){return this._map&&fe(this._image),this},bringToBack:function(){return this._map&&ge(this._image),this},setUrl:function(t){return this._url=t,this._image&&(this._image.src=t),this},setBounds:function(t){return this._bounds=g(t),this._map&&this._reset(),this},getEvents:function(){var t={zoom:this._reset,viewreset:this._reset};return this._zoomAnimated&&(t.zoomanim=this._animateZoom),t},setZIndex:function(t){return this.options.zIndex=t,this._updateZIndex(),this},getBounds:function(){return this._bounds},getElement:function(){return this._image},_initImage:function(){var t="IMG"===this._url.tagName,e=this._image=t?this._url:P("img");M(e,"leaflet-image-layer"),this._zoomAnimated&&M(e,"leaflet-zoom-animated"),this.options.className&&M(e,this.options.className),e.onselectstart=u,e.onmousemove=u,e.onload=a(this.fire,this,"load"),e.onerror=a(this._overlayOnError,this,"error"),!this.options.crossOrigin&&""!==this.options.crossOrigin||(e.crossOrigin=!0===this.options.crossOrigin?"":this.options.crossOrigin),this.options.zIndex&&this._updateZIndex(),t?this._url=e.src:(e.src=this._url,e.alt=this.options.alt)},_animateZoom:function(t){var e=this._map.getZoomScale(t.zoom),t=this._map._latLngBoundsToNewLayerBounds(this._bounds,t.zoom,t.center).min;be(this._image,t,e)},_reset:function(){var t=this._image,e=new f(this._map.latLngToLayerPoint(this._bounds.getNorthWest()),this._map.latLngToLayerPoint(this._bounds.getSouthEast())),i=e.getSize();Z(t,e.min),t.style.width=i.x+"px",t.style.height=i.y+"px"},_updateOpacity:function(){C(this._image,this.options.opacity)},_updateZIndex:function(){this._image&&void 0!==this.options.zIndex&&null!==this.options.zIndex&&(this._image.style.zIndex=this.options.zIndex)},_overlayOnError:function(){this.fire("error");var t=this.options.errorOverlayUrl;t&&this._url!==t&&(this._url=t,this._image.src=t)},getCenter:function(){return this._bounds.getCenter()}}),ki=Ei.extend({options:{autoplay:!0,loop:!0,keepAspectRatio:!0,muted:!1,playsInline:!0},_initImage:function(){var t="VIDEO"===this._url.tagName,e=this._image=t?this._url:P("video");if(M(e,"leaflet-image-layer"),this._zoomAnimated&&M(e,"leaflet-zoom-animated"),this.options.className&&M(e,this.options.className),e.onselectstart=u,e.onmousemove=u,e.onloadeddata=a(this.fire,this,"load"),t){for(var i=e.getElementsByTagName("source"),n=[],o=0;o<i.length;o++)n.push(i[o].src);this._url=0<i.length?n:[e.src]}else{d(this._url)||(this._url=[this._url]),!this.options.keepAspectRatio&&Object.prototype.hasOwnProperty.call(e.style,"objectFit")&&(e.style.objectFit="fill"),e.autoplay=!!this.options.autoplay,e.loop=!!this.options.loop,e.muted=!!this.options.muted,e.playsInline=!!this.options.playsInline;for(var s=0;s<this._url.length;s++){var r=P("source");r.src=this._url[s],e.appendChild(r)}}}});var Oi=Ei.extend({_initImage:function(){var t=this._image=this._url;M(t,"leaflet-image-layer"),this._zoomAnimated&&M(t,"leaflet-zoom-animated"),this.options.className&&M(t,this.options.className),t.onselectstart=u,t.onmousemove=u}});var Ai=o.extend({options:{interactive:!1,offset:[0,0],className:"",pane:void 0,content:""},initialize:function(t,e){t&&(t instanceof v||d(t))?(this._latlng=w(t),c(this,e)):(c(this,t),this._source=e),this.options.content&&(this._content=this.options.content)},openOn:function(t){return(t=arguments.length?t:this._source._map).hasLayer(this)||t.addLayer(this),this},close:function(){return this._map&&this._map.removeLayer(this),this},toggle:function(t){return this._map?this.close():(arguments.length?this._source=t:t=this._source,this._prepareOpen(),this.openOn(t._map)),this},onAdd:function(t){this._zoomAnimated=t._zoomAnimated,this._container||this._initLayout(),t._fadeAnimated&&C(this._container,0),clearTimeout(this._removeTimeout),this.getPane().appendChild(this._container),this.update(),t._fadeAnimated&&C(this._container,1),this.bringToFront(),this.options.interactive&&(M(this._container,"leaflet-interactive"),this.addInteractiveTarget(this._container))},onRemove:function(t){t._fadeAnimated?(C(this._container,0),this._removeTimeout=setTimeout(a(T,void 0,this._container),200)):T(this._container),this.options.interactive&&(z(this._container,"leaflet-interactive"),this.removeInteractiveTarget(this._container))},getLatLng:function(){return this._latlng},setLatLng:function(t){return this._latlng=w(t),this._map&&(this._updatePosition(),this._adjustPan()),this},getContent:function(){return this._content},setContent:function(t){return this._content=t,this.update(),this},getElement:function(){return this._container},update:function(){this._map&&(this._container.style.visibility="hidden",this._updateContent(),this._updateLayout(),this._updatePosition(),this._container.style.visibility="",this._adjustPan())},getEvents:function(){var t={zoom:this._updatePosition,viewreset:this._updatePosition};return this._zoomAnimated&&(t.zoomanim=this._animateZoom),t},isOpen:function(){return!!this._map&&this._map.hasLayer(this)},bringToFront:function(){return this._map&&fe(this._container),this},bringToBack:function(){return this._map&&ge(this._container),this},_prepareOpen:function(t){if(!(i=this._source)._map)return!1;if(i instanceof ci){var e,i=null,n=this._source._layers;for(e in n)if(n[e]._map){i=n[e];break}if(!i)return!1;this._source=i}if(!t)if(i.getCenter)t=i.getCenter();else if(i.getLatLng)t=i.getLatLng();else{if(!i.getBounds)throw new Error("Unable to get source layer LatLng.");t=i.getBounds().getCenter()}return this.setLatLng(t),this._map&&this.update(),!0},_updateContent:function(){if(this._content){var t=this._contentNode,e="function"==typeof this._content?this._content(this._source||this):this._content;if("string"==typeof e)t.innerHTML=e;else{for(;t.hasChildNodes();)t.removeChild(t.firstChild);t.appendChild(e)}this.fire("contentupdate")}},_updatePosition:function(){var t,e,i;this._map&&(e=this._map.latLngToLayerPoint(this._latlng),t=m(this.options.offset),i=this._getAnchor(),this._zoomAnimated?Z(this._container,e.add(i)):t=t.add(e).add(i),e=this._containerBottom=-t.y,i=this._containerLeft=-Math.round(this._containerWidth/2)+t.x,this._container.style.bottom=e+"px",this._container.style.left=i+"px")},_getAnchor:function(){return[0,0]}}),Bi=(A.include({_initOverlay:function(t,e,i,n){var o=e;return o instanceof t||(o=new t(n).setContent(e)),i&&o.setLatLng(i),o}}),o.include({_initOverlay:function(t,e,i,n){var o=i;return o instanceof t?(c(o,n),o._source=this):(o=e&&!n?e:new t(n,this)).setContent(i),o}}),Ai.extend({options:{pane:"popupPane",offset:[0,7],maxWidth:300,minWidth:50,maxHeight:null,autoPan:!0,autoPanPaddingTopLeft:null,autoPanPaddingBottomRight:null,autoPanPadding:[5,5],keepInView:!1,closeButton:!0,autoClose:!0,closeOnEscapeKey:!0,className:""},openOn:function(t){return!(t=arguments.length?t:this._source._map).hasLayer(this)&&t._popup&&t._popup.options.autoClose&&t.removeLayer(t._popup),t._popup=this,Ai.prototype.openOn.call(this,t)},onAdd:function(t){Ai.prototype.onAdd.call(this,t),t.fire("popupopen",{popup:this}),this._source&&(this._source.fire("popupopen",{popup:this},!0),this._source instanceof fi||this._source.on("preclick",Ae))},onRemove:function(t){Ai.prototype.onRemove.call(this,t),t.fire("popupclose",{popup:this}),this._source&&(this._source.fire("popupclose",{popup:this},!0),this._source instanceof fi||this._source.off("preclick",Ae))},getEvents:function(){var t=Ai.prototype.getEvents.call(this);return(void 0!==this.options.closeOnClick?this.options.closeOnClick:this._map.options.closePopupOnClick)&&(t.preclick=this.close),this.options.keepInView&&(t.moveend=this._adjustPan),t},_initLayout:function(){var t="leaflet-popup",e=this._container=P("div",t+" "+(this.options.className||"")+" leaflet-zoom-animated"),i=this._wrapper=P("div",t+"-content-wrapper",e);this._contentNode=P("div",t+"-content",i),Ie(e),Be(this._contentNode),S(e,"contextmenu",Ae),this._tipContainer=P("div",t+"-tip-container",e),this._tip=P("div",t+"-tip",this._tipContainer),this.options.closeButton&&((i=this._closeButton=P("a",t+"-close-button",e)).setAttribute("role","button"),i.setAttribute("aria-label","Close popup"),i.href="#close",i.innerHTML='<span aria-hidden="true">&#215;</span>',S(i,"click",function(t){O(t),this.close()},this))},_updateLayout:function(){var t=this._contentNode,e=t.style,i=(e.width="",e.whiteSpace="nowrap",t.offsetWidth),i=Math.min(i,this.options.maxWidth),i=(i=Math.max(i,this.options.minWidth),e.width=i+1+"px",e.whiteSpace="",e.height="",t.offsetHeight),n=this.options.maxHeight,o="leaflet-popup-scrolled";(n&&n<i?(e.height=n+"px",M):z)(t,o),this._containerWidth=this._container.offsetWidth},_animateZoom:function(t){var t=this._map._latLngToNewLayerPoint(this._latlng,t.zoom,t.center),e=this._getAnchor();Z(this._container,t.add(e))},_adjustPan:function(){var t,e,i,n,o,s,r,a;this.options.autoPan&&(this._map._panAnim&&this._map._panAnim.stop(),this._autopanning?this._autopanning=!1:(t=this._map,e=parseInt(pe(this._container,"marginBottom"),10)||0,e=this._container.offsetHeight+e,a=this._containerWidth,(i=new p(this._containerLeft,-e-this._containerBottom))._add(Pe(this._container)),i=t.layerPointToContainerPoint(i),o=m(this.options.autoPanPadding),n=m(this.options.autoPanPaddingTopLeft||o),o=m(this.options.autoPanPaddingBottomRight||o),s=t.getSize(),r=0,i.x+a+o.x>s.x&&(r=i.x+a-s.x+o.x),i.x-r-n.x<(a=0)&&(r=i.x-n.x),i.y+e+o.y>s.y&&(a=i.y+e-s.y+o.y),i.y-a-n.y<0&&(a=i.y-n.y),(r||a)&&(this.options.keepInView&&(this._autopanning=!0),t.fire("autopanstart").panBy([r,a]))))},_getAnchor:function(){return m(this._source&&this._source._getPopupAnchor?this._source._getPopupAnchor():[0,0])}})),Ii=(A.mergeOptions({closePopupOnClick:!0}),A.include({openPopup:function(t,e,i){return this._initOverlay(Bi,t,e,i).openOn(this),this},closePopup:function(t){return(t=arguments.length?t:this._popup)&&t.close(),this}}),o.include({bindPopup:function(t,e){return this._popup=this._initOverlay(Bi,this._popup,t,e),this._popupHandlersAdded||(this.on({click:this._openPopup,keypress:this._onKeyPress,remove:this.closePopup,move:this._movePopup}),this._popupHandlersAdded=!0),this},unbindPopup:function(){return this._popup&&(this.off({click:this._openPopup,keypress:this._onKeyPress,remove:this.closePopup,move:this._movePopup}),this._popupHandlersAdded=!1,this._popup=null),this},openPopup:function(t){return this._popup&&(this instanceof ci||(this._popup._source=this),this._popup._prepareOpen(t||this._latlng)&&this._popup.openOn(this._map)),this},closePopup:function(){return this._popup&&this._popup.close(),this},togglePopup:function(){return this._popup&&this._popup.toggle(this),this},isPopupOpen:function(){return!!this._popup&&this._popup.isOpen()},setPopupContent:function(t){return this._popup&&this._popup.setContent(t),this},getPopup:function(){return this._popup},_openPopup:function(t){var e;this._popup&&this._map&&(Re(t),e=t.layer||t.target,this._popup._source!==e||e instanceof fi?(this._popup._source=e,this.openPopup(t.latlng)):this._map.hasLayer(this._popup)?this.closePopup():this.openPopup(t.latlng))},_movePopup:function(t){this._popup.setLatLng(t.latlng)},_onKeyPress:function(t){13===t.originalEvent.keyCode&&this._openPopup(t)}}),Ai.extend({options:{pane:"tooltipPane",offset:[0,0],direction:"auto",permanent:!1,sticky:!1,opacity:.9},onAdd:function(t){Ai.prototype.onAdd.call(this,t),this.setOpacity(this.options.opacity),t.fire("tooltipopen",{tooltip:this}),this._source&&(this.addEventParent(this._source),this._source.fire("tooltipopen",{tooltip:this},!0))},onRemove:function(t){Ai.prototype.onRemove.call(this,t),t.fire("tooltipclose",{tooltip:this}),this._source&&(this.removeEventParent(this._source),this._source.fire("tooltipclose",{tooltip:this},!0))},getEvents:function(){var t=Ai.prototype.getEvents.call(this);return this.options.permanent||(t.preclick=this.close),t},_initLayout:function(){var t="leaflet-tooltip "+(this.options.className||"")+" leaflet-zoom-"+(this._zoomAnimated?"animated":"hide");this._contentNode=this._container=P("div",t),this._container.setAttribute("role","tooltip"),this._container.setAttribute("id","leaflet-tooltip-"+h(this))},_updateLayout:function(){},_adjustPan:function(){},_setPosition:function(t){var e,i=this._map,n=this._container,o=i.latLngToContainerPoint(i.getCenter()),i=i.layerPointToContainerPoint(t),s=this.options.direction,r=n.offsetWidth,a=n.offsetHeight,h=m(this.options.offset),l=this._getAnchor(),i="top"===s?(e=r/2,a):"bottom"===s?(e=r/2,0):(e="center"===s?r/2:"right"===s?0:"left"===s?r:i.x<o.x?(s="right",0):(s="left",r+2*(h.x+l.x)),a/2);t=t.subtract(m(e,i,!0)).add(h).add(l),z(n,"leaflet-tooltip-right"),z(n,"leaflet-tooltip-left"),z(n,"leaflet-tooltip-top"),z(n,"leaflet-tooltip-bottom"),M(n,"leaflet-tooltip-"+s),Z(n,t)},_updatePosition:function(){var t=this._map.latLngToLayerPoint(this._latlng);this._setPosition(t)},setOpacity:function(t){this.options.opacity=t,this._container&&C(this._container,t)},_animateZoom:function(t){t=this._map._latLngToNewLayerPoint(this._latlng,t.zoom,t.center);this._setPosition(t)},_getAnchor:function(){return m(this._source&&this._source._getTooltipAnchor&&!this.options.sticky?this._source._getTooltipAnchor():[0,0])}})),Ri=(A.include({openTooltip:function(t,e,i){return this._initOverlay(Ii,t,e,i).openOn(this),this},closeTooltip:function(t){return t.close(),this}}),o.include({bindTooltip:function(t,e){return this._tooltip&&this.isTooltipOpen()&&this.unbindTooltip(),this._tooltip=this._initOverlay(Ii,this._tooltip,t,e),this._initTooltipInteractions(),this._tooltip.options.permanent&&this._map&&this._map.hasLayer(this)&&this.openTooltip(),this},unbindTooltip:function(){return this._tooltip&&(this._initTooltipInteractions(!0),this.closeTooltip(),this._tooltip=null),this},_initTooltipInteractions:function(t){var e,i;!t&&this._tooltipHandlersAdded||(e=t?"off":"on",i={remove:this.closeTooltip,move:this._moveTooltip},this._tooltip.options.permanent?i.add=this._openTooltip:(i.mouseover=this._openTooltip,i.mouseout=this.closeTooltip,i.click=this._openTooltip,this._map?this._addFocusListeners():i.add=this._addFocusListeners),this._tooltip.options.sticky&&(i.mousemove=this._moveTooltip),this[e](i),this._tooltipHandlersAdded=!t)},openTooltip:function(t){return this._tooltip&&(this instanceof ci||(this._tooltip._source=this),this._tooltip._prepareOpen(t)&&(this._tooltip.openOn(this._map),this.getElement?this._setAriaDescribedByOnLayer(this):this.eachLayer&&this.eachLayer(this._setAriaDescribedByOnLayer,this))),this},closeTooltip:function(){if(this._tooltip)return this._tooltip.close()},toggleTooltip:function(){return this._tooltip&&this._tooltip.toggle(this),this},isTooltipOpen:function(){return this._tooltip.isOpen()},setTooltipContent:function(t){return this._tooltip&&this._tooltip.setContent(t),this},getTooltip:function(){return this._tooltip},_addFocusListeners:function(){this.getElement?this._addFocusListenersOnLayer(this):this.eachLayer&&this.eachLayer(this._addFocusListenersOnLayer,this)},_addFocusListenersOnLayer:function(t){var e="function"==typeof t.getElement&&t.getElement();e&&(S(e,"focus",function(){this._tooltip._source=t,this.openTooltip()},this),S(e,"blur",this.closeTooltip,this))},_setAriaDescribedByOnLayer:function(t){t="function"==typeof t.getElement&&t.getElement();t&&t.setAttribute("aria-describedby",this._tooltip._container.id)},_openTooltip:function(t){var e;this._tooltip&&this._map&&(this._map.dragging&&this._map.dragging.moving()&&!this._openOnceFlag?(this._openOnceFlag=!0,(e=this)._map.once("moveend",function(){e._openOnceFlag=!1,e._openTooltip(t)})):(this._tooltip._source=t.layer||t.target,this.openTooltip(this._tooltip.options.sticky?t.latlng:void 0)))},_moveTooltip:function(t){var e=t.latlng;this._tooltip.options.sticky&&t.originalEvent&&(t=this._map.mouseEventToContainerPoint(t.originalEvent),t=this._map.containerPointToLayerPoint(t),e=this._map.layerPointToLatLng(t)),this._tooltip.setLatLng(e)}}),di.extend({options:{iconSize:[12,12],html:!1,bgPos:null,className:"leaflet-div-icon"},createIcon:function(t){var t=t&&"DIV"===t.tagName?t:document.createElement("div"),e=this.options;return e.html instanceof Element?(me(t),t.appendChild(e.html)):t.innerHTML=!1!==e.html?e.html:"",e.bgPos&&(e=m(e.bgPos),t.style.backgroundPosition=-e.x+"px "+-e.y+"px"),this._setIconStyles(t,"icon"),t},createShadow:function(){return null}}));di.Default=_i;var Ni=o.extend({options:{tileSize:256,opacity:1,updateWhenIdle:b.mobile,updateWhenZooming:!0,updateInterval:200,zIndex:1,bounds:null,minZoom:0,maxZoom:void 0,maxNativeZoom:void 0,minNativeZoom:void 0,noWrap:!1,pane:"tilePane",className:"",keepBuffer:2},initialize:function(t){c(this,t)},onAdd:function(){this._initContainer(),this._levels={},this._tiles={},this._resetView()},beforeAdd:function(t){t._addZoomLimit(this)},onRemove:function(t){this._removeAllTiles(),T(this._container),t._removeZoomLimit(this),this._container=null,this._tileZoom=void 0},bringToFront:function(){return this._map&&(fe(this._container),this._setAutoZIndex(Math.max)),this},bringToBack:function(){return this._map&&(ge(this._container),this._setAutoZIndex(Math.min)),this},getContainer:function(){return this._container},setOpacity:function(t){return this.options.opacity=t,this._updateOpacity(),this},setZIndex:function(t){return this.options.zIndex=t,this._updateZIndex(),this},isLoading:function(){return this._loading},redraw:function(){var t;return this._map&&(this._removeAllTiles(),(t=this._clampZoom(this._map.getZoom()))!==this._tileZoom&&(this._tileZoom=t,this._updateLevels()),this._update()),this},getEvents:function(){var t={viewprereset:this._invalidateAll,viewreset:this._resetView,zoom:this._resetView,moveend:this._onMoveEnd};return this.options.updateWhenIdle||(this._onMove||(this._onMove=j(this._onMoveEnd,this.options.updateInterval,this)),t.move=this._onMove),this._zoomAnimated&&(t.zoomanim=this._animateZoom),t},createTile:function(){return document.createElement("div")},getTileSize:function(){var t=this.options.tileSize;return t instanceof p?t:new p(t,t)},_updateZIndex:function(){this._container&&void 0!==this.options.zIndex&&null!==this.options.zIndex&&(this._container.style.zIndex=this.options.zIndex)},_setAutoZIndex:function(t){for(var e,i=this.getPane().children,n=-t(-1/0,1/0),o=0,s=i.length;o<s;o++)e=i[o].style.zIndex,i[o]!==this._container&&e&&(n=t(n,+e));isFinite(n)&&(this.options.zIndex=n+t(-1,1),this._updateZIndex())},_updateOpacity:function(){if(this._map&&!b.ielt9){C(this._container,this.options.opacity);var t,e=+new Date,i=!1,n=!1;for(t in this._tiles){var o,s=this._tiles[t];s.current&&s.loaded&&(o=Math.min(1,(e-s.loaded)/200),C(s.el,o),o<1?i=!0:(s.active?n=!0:this._onOpaqueTile(s),s.active=!0))}n&&!this._noPrune&&this._pruneTiles(),i&&(r(this._fadeFrame),this._fadeFrame=x(this._updateOpacity,this))}},_onOpaqueTile:u,_initContainer:function(){this._container||(this._container=P("div","leaflet-layer "+(this.options.className||"")),this._updateZIndex(),this.options.opacity<1&&this._updateOpacity(),this.getPane().appendChild(this._container))},_updateLevels:function(){var t=this._tileZoom,e=this.options.maxZoom;if(void 0!==t){for(var i in this._levels)i=Number(i),this._levels[i].el.children.length||i===t?(this._levels[i].el.style.zIndex=e-Math.abs(t-i),this._onUpdateLevel(i)):(T(this._levels[i].el),this._removeTilesAtZoom(i),this._onRemoveLevel(i),delete this._levels[i]);var n=this._levels[t],o=this._map;return n||((n=this._levels[t]={}).el=P("div","leaflet-tile-container leaflet-zoom-animated",this._container),n.el.style.zIndex=e,n.origin=o.project(o.unproject(o.getPixelOrigin()),t).round(),n.zoom=t,this._setZoomTransform(n,o.getCenter(),o.getZoom()),u(n.el.offsetWidth),this._onCreateLevel(n)),this._level=n}},_onUpdateLevel:u,_onRemoveLevel:u,_onCreateLevel:u,_pruneTiles:function(){if(this._map){var t,e,i,n=this._map.getZoom();if(n>this.options.maxZoom||n<this.options.minZoom)this._removeAllTiles();else{for(t in this._tiles)(i=this._tiles[t]).retain=i.current;for(t in this._tiles)(i=this._tiles[t]).current&&!i.active&&(e=i.coords,this._retainParent(e.x,e.y,e.z,e.z-5)||this._retainChildren(e.x,e.y,e.z,e.z+2));for(t in this._tiles)this._tiles[t].retain||this._removeTile(t)}}},_removeTilesAtZoom:function(t){for(var e in this._tiles)this._tiles[e].coords.z===t&&this._removeTile(e)},_removeAllTiles:function(){for(var t in this._tiles)this._removeTile(t)},_invalidateAll:function(){for(var t in this._levels)T(this._levels[t].el),this._onRemoveLevel(Number(t)),delete this._levels[t];this._removeAllTiles(),this._tileZoom=void 0},_retainParent:function(t,e,i,n){var t=Math.floor(t/2),e=Math.floor(e/2),i=i-1,o=new p(+t,+e),o=(o.z=i,this._tileCoordsToKey(o)),o=this._tiles[o];return o&&o.active?o.retain=!0:(o&&o.loaded&&(o.retain=!0),n<i&&this._retainParent(t,e,i,n))},_retainChildren:function(t,e,i,n){for(var o=2*t;o<2*t+2;o++)for(var s=2*e;s<2*e+2;s++){var r=new p(o,s),r=(r.z=i+1,this._tileCoordsToKey(r)),r=this._tiles[r];r&&r.active?r.retain=!0:(r&&r.loaded&&(r.retain=!0),i+1<n&&this._retainChildren(o,s,i+1,n))}},_resetView:function(t){t=t&&(t.pinch||t.flyTo);this._setView(this._map.getCenter(),this._map.getZoom(),t,t)},_animateZoom:function(t){this._setView(t.center,t.zoom,!0,t.noUpdate)},_clampZoom:function(t){var e=this.options;return void 0!==e.minNativeZoom&&t<e.minNativeZoom?e.minNativeZoom:void 0!==e.maxNativeZoom&&e.maxNativeZoom<t?e.maxNativeZoom:t},_setView:function(t,e,i,n){var o=Math.round(e),o=void 0!==this.options.maxZoom&&o>this.options.maxZoom||void 0!==this.options.minZoom&&o<this.options.minZoom?void 0:this._clampZoom(o),s=this.options.updateWhenZooming&&o!==this._tileZoom;n&&!s||(this._tileZoom=o,this._abortLoading&&this._abortLoading(),this._updateLevels(),this._resetGrid(),void 0!==o&&this._update(t),i||this._pruneTiles(),this._noPrune=!!i),this._setZoomTransforms(t,e)},_setZoomTransforms:function(t,e){for(var i in this._levels)this._setZoomTransform(this._levels[i],t,e)},_setZoomTransform:function(t,e,i){var n=this._map.getZoomScale(i,t.zoom),e=t.origin.multiplyBy(n).subtract(this._map._getNewPixelOrigin(e,i)).round();b.any3d?be(t.el,e,n):Z(t.el,e)},_resetGrid:function(){var t=this._map,e=t.options.crs,i=this._tileSize=this.getTileSize(),n=this._tileZoom,o=this._map.getPixelWorldBounds(this._tileZoom);o&&(this._globalTileRange=this._pxBoundsToTileRange(o)),this._wrapX=e.wrapLng&&!this.options.noWrap&&[Math.floor(t.project([0,e.wrapLng[0]],n).x/i.x),Math.ceil(t.project([0,e.wrapLng[1]],n).x/i.y)],this._wrapY=e.wrapLat&&!this.options.noWrap&&[Math.floor(t.project([e.wrapLat[0],0],n).y/i.x),Math.ceil(t.project([e.wrapLat[1],0],n).y/i.y)]},_onMoveEnd:function(){this._map&&!this._map._animatingZoom&&this._update()},_getTiledPixelBounds:function(t){var e=this._map,i=e._animatingZoom?Math.max(e._animateToZoom,e.getZoom()):e.getZoom(),i=e.getZoomScale(i,this._tileZoom),t=e.project(t,this._tileZoom).floor(),e=e.getSize().divideBy(2*i);return new f(t.subtract(e),t.add(e))},_update:function(t){var e=this._map;if(e){var i=this._clampZoom(e.getZoom());if(void 0===t&&(t=e.getCenter()),void 0!==this._tileZoom){var n,e=this._getTiledPixelBounds(t),o=this._pxBoundsToTileRange(e),s=o.getCenter(),r=[],e=this.options.keepBuffer,a=new f(o.getBottomLeft().subtract([e,-e]),o.getTopRight().add([e,-e]));if(!(isFinite(o.min.x)&&isFinite(o.min.y)&&isFinite(o.max.x)&&isFinite(o.max.y)))throw new Error("Attempted to load an infinite number of tiles");for(n in this._tiles){var h=this._tiles[n].coords;h.z===this._tileZoom&&a.contains(new p(h.x,h.y))||(this._tiles[n].current=!1)}if(1<Math.abs(i-this._tileZoom))this._setView(t,i);else{for(var l=o.min.y;l<=o.max.y;l++)for(var u=o.min.x;u<=o.max.x;u++){var c,d=new p(u,l);d.z=this._tileZoom,this._isValidTile(d)&&((c=this._tiles[this._tileCoordsToKey(d)])?c.current=!0:r.push(d))}if(r.sort(function(t,e){return t.distanceTo(s)-e.distanceTo(s)}),0!==r.length){this._loading||(this._loading=!0,this.fire("loading"));for(var _=document.createDocumentFragment(),u=0;u<r.length;u++)this._addTile(r[u],_);this._level.el.appendChild(_)}}}}},_isValidTile:function(t){var e=this._map.options.crs;if(!e.infinite){var i=this._globalTileRange;if(!e.wrapLng&&(t.x<i.min.x||t.x>i.max.x)||!e.wrapLat&&(t.y<i.min.y||t.y>i.max.y))return!1}return!this.options.bounds||(e=this._tileCoordsToBounds(t),g(this.options.bounds).overlaps(e))},_keyToBounds:function(t){return this._tileCoordsToBounds(this._keyToTileCoords(t))},_tileCoordsToNwSe:function(t){var e=this._map,i=this.getTileSize(),n=t.scaleBy(i),i=n.add(i);return[e.unproject(n,t.z),e.unproject(i,t.z)]},_tileCoordsToBounds:function(t){t=this._tileCoordsToNwSe(t),t=new s(t[0],t[1]);return t=this.options.noWrap?t:this._map.wrapLatLngBounds(t)},_tileCoordsToKey:function(t){return t.x+":"+t.y+":"+t.z},_keyToTileCoords:function(t){var t=t.split(":"),e=new p(+t[0],+t[1]);return e.z=+t[2],e},_removeTile:function(t){var e=this._tiles[t];e&&(T(e.el),delete this._tiles[t],this.fire("tileunload",{tile:e.el,coords:this._keyToTileCoords(t)}))},_initTile:function(t){M(t,"leaflet-tile");var e=this.getTileSize();t.style.width=e.x+"px",t.style.height=e.y+"px",t.onselectstart=u,t.onmousemove=u,b.ielt9&&this.options.opacity<1&&C(t,this.options.opacity)},_addTile:function(t,e){var i=this._getTilePos(t),n=this._tileCoordsToKey(t),o=this.createTile(this._wrapCoords(t),a(this._tileReady,this,t));this._initTile(o),this.createTile.length<2&&x(a(this._tileReady,this,t,null,o)),Z(o,i),this._tiles[n]={el:o,coords:t,current:!0},e.appendChild(o),this.fire("tileloadstart",{tile:o,coords:t})},_tileReady:function(t,e,i){e&&this.fire("tileerror",{error:e,tile:i,coords:t});var n=this._tileCoordsToKey(t);(i=this._tiles[n])&&(i.loaded=+new Date,this._map._fadeAnimated?(C(i.el,0),r(this._fadeFrame),this._fadeFrame=x(this._updateOpacity,this)):(i.active=!0,this._pruneTiles()),e||(M(i.el,"leaflet-tile-loaded"),this.fire("tileload",{tile:i.el,coords:t})),this._noTilesToLoad()&&(this._loading=!1,this.fire("load"),b.ielt9||!this._map._fadeAnimated?x(this._pruneTiles,this):setTimeout(a(this._pruneTiles,this),250)))},_getTilePos:function(t){return t.scaleBy(this.getTileSize()).subtract(this._level.origin)},_wrapCoords:function(t){var e=new p(this._wrapX?H(t.x,this._wrapX):t.x,this._wrapY?H(t.y,this._wrapY):t.y);return e.z=t.z,e},_pxBoundsToTileRange:function(t){var e=this.getTileSize();return new f(t.min.unscaleBy(e).floor(),t.max.unscaleBy(e).ceil().subtract([1,1]))},_noTilesToLoad:function(){for(var t in this._tiles)if(!this._tiles[t].loaded)return!1;return!0}});var Di=Ni.extend({options:{minZoom:0,maxZoom:18,subdomains:"abc",errorTileUrl:"",zoomOffset:0,tms:!1,zoomReverse:!1,detectRetina:!1,crossOrigin:!1,referrerPolicy:!1},initialize:function(t,e){this._url=t,(e=c(this,e)).detectRetina&&b.retina&&0<e.maxZoom?(e.tileSize=Math.floor(e.tileSize/2),e.zoomReverse?(e.zoomOffset--,e.minZoom=Math.min(e.maxZoom,e.minZoom+1)):(e.zoomOffset++,e.maxZoom=Math.max(e.minZoom,e.maxZoom-1)),e.minZoom=Math.max(0,e.minZoom)):e.zoomReverse?e.minZoom=Math.min(e.maxZoom,e.minZoom):e.maxZoom=Math.max(e.minZoom,e.maxZoom),"string"==typeof e.subdomains&&(e.subdomains=e.subdomains.split("")),this.on("tileunload",this._onTileRemove)},setUrl:function(t,e){return this._url===t&&void 0===e&&(e=!0),this._url=t,e||this.redraw(),this},createTile:function(t,e){var i=document.createElement("img");return S(i,"load",a(this._tileOnLoad,this,e,i)),S(i,"error",a(this._tileOnError,this,e,i)),!this.options.crossOrigin&&""!==this.options.crossOrigin||(i.crossOrigin=!0===this.options.crossOrigin?"":this.options.crossOrigin),"string"==typeof this.options.referrerPolicy&&(i.referrerPolicy=this.options.referrerPolicy),i.alt="",i.src=this.getTileUrl(t),i},getTileUrl:function(t){var e={r:b.retina?"@2x":"",s:this._getSubdomain(t),x:t.x,y:t.y,z:this._getZoomForUrl()};return this._map&&!this._map.options.crs.infinite&&(t=this._globalTileRange.max.y-t.y,this.options.tms&&(e.y=t),e["-y"]=t),q(this._url,l(e,this.options))},_tileOnLoad:function(t,e){b.ielt9?setTimeout(a(t,this,null,e),0):t(null,e)},_tileOnError:function(t,e,i){var n=this.options.errorTileUrl;n&&e.getAttribute("src")!==n&&(e.src=n),t(i,e)},_onTileRemove:function(t){t.tile.onload=null},_getZoomForUrl:function(){var t=this._tileZoom,e=this.options.maxZoom;return(t=this.options.zoomReverse?e-t:t)+this.options.zoomOffset},_getSubdomain:function(t){t=Math.abs(t.x+t.y)%this.options.subdomains.length;return this.options.subdomains[t]},_abortLoading:function(){var t,e,i;for(t in this._tiles)this._tiles[t].coords.z!==this._tileZoom&&((i=this._tiles[t].el).onload=u,i.onerror=u,i.complete||(i.src=K,e=this._tiles[t].coords,T(i),delete this._tiles[t],this.fire("tileabort",{tile:i,coords:e})))},_removeTile:function(t){var e=this._tiles[t];if(e)return e.el.setAttribute("src",K),Ni.prototype._removeTile.call(this,t)},_tileReady:function(t,e,i){if(this._map&&(!i||i.getAttribute("src")!==K))return Ni.prototype._tileReady.call(this,t,e,i)}});function ji(t,e){return new Di(t,e)}var Hi=Di.extend({defaultWmsParams:{service:"WMS",request:"GetMap",layers:"",styles:"",format:"image/jpeg",transparent:!1,version:"1.1.1"},options:{crs:null,uppercase:!1},initialize:function(t,e){this._url=t;var i,n=l({},this.defaultWmsParams);for(i in e)i in this.options||(n[i]=e[i]);var t=(e=c(this,e)).detectRetina&&b.retina?2:1,o=this.getTileSize();n.width=o.x*t,n.height=o.y*t,this.wmsParams=n},onAdd:function(t){this._crs=this.options.crs||t.options.crs,this._wmsVersion=parseFloat(this.wmsParams.version);var e=1.3<=this._wmsVersion?"crs":"srs";this.wmsParams[e]=this._crs.code,Di.prototype.onAdd.call(this,t)},getTileUrl:function(t){var e=this._tileCoordsToNwSe(t),i=this._crs,i=_(i.project(e[0]),i.project(e[1])),e=i.min,i=i.max,e=(1.3<=this._wmsVersion&&this._crs===li?[e.y,e.x,i.y,i.x]:[e.x,e.y,i.x,i.y]).join(","),i=Di.prototype.getTileUrl.call(this,t);return i+U(this.wmsParams,i,this.options.uppercase)+(this.options.uppercase?"&BBOX=":"&bbox=")+e},setParams:function(t,e){return l(this.wmsParams,t),e||this.redraw(),this}});Di.WMS=Hi,ji.wms=function(t,e){return new Hi(t,e)};var Wi=o.extend({options:{padding:.1},initialize:function(t){c(this,t),h(this),this._layers=this._layers||{}},onAdd:function(){this._container||(this._initContainer(),M(this._container,"leaflet-zoom-animated")),this.getPane().appendChild(this._container),this._update(),this.on("update",this._updatePaths,this)},onRemove:function(){this.off("update",this._updatePaths,this),this._destroyContainer()},getEvents:function(){var t={viewreset:this._reset,zoom:this._onZoom,moveend:this._update,zoomend:this._onZoomEnd};return this._zoomAnimated&&(t.zoomanim=this._onAnimZoom),t},_onAnimZoom:function(t){this._updateTransform(t.center,t.zoom)},_onZoom:function(){this._updateTransform(this._map.getCenter(),this._map.getZoom())},_updateTransform:function(t,e){var i=this._map.getZoomScale(e,this._zoom),n=this._map.getSize().multiplyBy(.5+this.options.padding),o=this._map.project(this._center,e),n=n.multiplyBy(-i).add(o).subtract(this._map._getNewPixelOrigin(t,e));b.any3d?be(this._container,n,i):Z(this._container,n)},_reset:function(){for(var t in this._update(),this._updateTransform(this._center,this._zoom),this._layers)this._layers[t]._reset()},_onZoomEnd:function(){for(var t in this._layers)this._layers[t]._project()},_updatePaths:function(){for(var t in this._layers)this._layers[t]._update()},_update:function(){var t=this.options.padding,e=this._map.getSize(),i=this._map.containerPointToLayerPoint(e.multiplyBy(-t)).round();this._bounds=new f(i,i.add(e.multiplyBy(1+2*t)).round()),this._center=this._map.getCenter(),this._zoom=this._map.getZoom()}}),Fi=Wi.extend({options:{tolerance:0},getEvents:function(){var t=Wi.prototype.getEvents.call(this);return t.viewprereset=this._onViewPreReset,t},_onViewPreReset:function(){this._postponeUpdatePaths=!0},onAdd:function(){Wi.prototype.onAdd.call(this),this._draw()},_initContainer:function(){var t=this._container=document.createElement("canvas");S(t,"mousemove",this._onMouseMove,this),S(t,"click dblclick mousedown mouseup contextmenu",this._onClick,this),S(t,"mouseout",this._handleMouseOut,this),t._leaflet_disable_events=!0,this._ctx=t.getContext("2d")},_destroyContainer:function(){r(this._redrawRequest),delete this._ctx,T(this._container),k(this._container),delete this._container},_updatePaths:function(){if(!this._postponeUpdatePaths){for(var t in this._redrawBounds=null,this._layers)this._layers[t]._update();this._redraw()}},_update:function(){var t,e,i,n;this._map._animatingZoom&&this._bounds||(Wi.prototype._update.call(this),t=this._bounds,e=this._container,i=t.getSize(),n=b.retina?2:1,Z(e,t.min),e.width=n*i.x,e.height=n*i.y,e.style.width=i.x+"px",e.style.height=i.y+"px",b.retina&&this._ctx.scale(2,2),this._ctx.translate(-t.min.x,-t.min.y),this.fire("update"))},_reset:function(){Wi.prototype._reset.call(this),this._postponeUpdatePaths&&(this._postponeUpdatePaths=!1,this._updatePaths())},_initPath:function(t){this._updateDashArray(t);t=(this._layers[h(t)]=t)._order={layer:t,prev:this._drawLast,next:null};this._drawLast&&(this._drawLast.next=t),this._drawLast=t,this._drawFirst=this._drawFirst||this._drawLast},_addPath:function(t){this._requestRedraw(t)},_removePath:function(t){var e=t._order,i=e.next,e=e.prev;i?i.prev=e:this._drawLast=e,e?e.next=i:this._drawFirst=i,delete t._order,delete this._layers[h(t)],this._requestRedraw(t)},_updatePath:function(t){this._extendRedrawBounds(t),t._project(),t._update(),this._requestRedraw(t)},_updateStyle:function(t){this._updateDashArray(t),this._requestRedraw(t)},_updateDashArray:function(t){if("string"==typeof t.options.dashArray){for(var e,i=t.options.dashArray.split(/[, ]+/),n=[],o=0;o<i.length;o++){if(e=Number(i[o]),isNaN(e))return;n.push(e)}t.options._dashArray=n}else t.options._dashArray=t.options.dashArray},_requestRedraw:function(t){this._map&&(this._extendRedrawBounds(t),this._redrawRequest=this._redrawRequest||x(this._redraw,this))},_extendRedrawBounds:function(t){var e;t._pxBounds&&(e=(t.options.weight||0)+1,this._redrawBounds=this._redrawBounds||new f,this._redrawBounds.extend(t._pxBounds.min.subtract([e,e])),this._redrawBounds.extend(t._pxBounds.max.add([e,e])))},_redraw:function(){this._redrawRequest=null,this._redrawBounds&&(this._redrawBounds.min._floor(),this._redrawBounds.max._ceil()),this._clear(),this._draw(),this._redrawBounds=null},_clear:function(){var t,e=this._redrawBounds;e?(t=e.getSize(),this._ctx.clearRect(e.min.x,e.min.y,t.x,t.y)):(this._ctx.save(),this._ctx.setTransform(1,0,0,1,0,0),this._ctx.clearRect(0,0,this._container.width,this._container.height),this._ctx.restore())},_draw:function(){var t,e,i=this._redrawBounds;this._ctx.save(),i&&(e=i.getSize(),this._ctx.beginPath(),this._ctx.rect(i.min.x,i.min.y,e.x,e.y),this._ctx.clip()),this._drawing=!0;for(var n=this._drawFirst;n;n=n.next)t=n.layer,(!i||t._pxBounds&&t._pxBounds.intersects(i))&&t._updatePath();this._drawing=!1,this._ctx.restore()},_updatePoly:function(t,e){if(this._drawing){var i,n,o,s,r=t._parts,a=r.length,h=this._ctx;if(a){for(h.beginPath(),i=0;i<a;i++){for(n=0,o=r[i].length;n<o;n++)s=r[i][n],h[n?"lineTo":"moveTo"](s.x,s.y);e&&h.closePath()}this._fillStroke(h,t)}}},_updateCircle:function(t){var e,i,n,o;this._drawing&&!t._empty()&&(e=t._point,i=this._ctx,n=Math.max(Math.round(t._radius),1),1!=(o=(Math.max(Math.round(t._radiusY),1)||n)/n)&&(i.save(),i.scale(1,o)),i.beginPath(),i.arc(e.x,e.y/o,n,0,2*Math.PI,!1),1!=o&&i.restore(),this._fillStroke(i,t))},_fillStroke:function(t,e){var i=e.options;i.fill&&(t.globalAlpha=i.fillOpacity,t.fillStyle=i.fillColor||i.color,t.fill(i.fillRule||"evenodd")),i.stroke&&0!==i.weight&&(t.setLineDash&&t.setLineDash(e.options&&e.options._dashArray||[]),t.globalAlpha=i.opacity,t.lineWidth=i.weight,t.strokeStyle=i.color,t.lineCap=i.lineCap,t.lineJoin=i.lineJoin,t.stroke())},_onClick:function(t){for(var e,i,n=this._map.mouseEventToLayerPoint(t),o=this._drawFirst;o;o=o.next)(e=o.layer).options.interactive&&e._containsPoint(n)&&(("click"===t.type||"preclick"===t.type)&&this._map._draggableMoved(e)||(i=e));this._fireEvent(!!i&&[i],t)},_onMouseMove:function(t){var e;!this._map||this._map.dragging.moving()||this._map._animatingZoom||(e=this._map.mouseEventToLayerPoint(t),this._handleMouseHover(t,e))},_handleMouseOut:function(t){var e=this._hoveredLayer;e&&(z(this._container,"leaflet-interactive"),this._fireEvent([e],t,"mouseout"),this._hoveredLayer=null,this._mouseHoverThrottled=!1)},_handleMouseHover:function(t,e){if(!this._mouseHoverThrottled){for(var i,n,o=this._drawFirst;o;o=o.next)(i=o.layer).options.interactive&&i._containsPoint(e)&&(n=i);n!==this._hoveredLayer&&(this._handleMouseOut(t),n&&(M(this._container,"leaflet-interactive"),this._fireEvent([n],t,"mouseover"),this._hoveredLayer=n)),this._fireEvent(!!this._hoveredLayer&&[this._hoveredLayer],t),this._mouseHoverThrottled=!0,setTimeout(a(function(){this._mouseHoverThrottled=!1},this),32)}},_fireEvent:function(t,e,i){this._map._fireDOMEvent(e,i||e.type,t)},_bringToFront:function(t){var e,i,n=t._order;n&&(e=n.next,i=n.prev,e&&((e.prev=i)?i.next=e:e&&(this._drawFirst=e),n.prev=this._drawLast,(this._drawLast.next=n).next=null,this._drawLast=n,this._requestRedraw(t)))},_bringToBack:function(t){var e,i,n=t._order;n&&(e=n.next,(i=n.prev)&&((i.next=e)?e.prev=i:i&&(this._drawLast=i),n.prev=null,n.next=this._drawFirst,this._drawFirst.prev=n,this._drawFirst=n,this._requestRedraw(t)))}});function Ui(t){return b.canvas?new Fi(t):null}var Vi=function(){try{return document.namespaces.add("lvml","urn:schemas-microsoft-com:vml"),function(t){return document.createElement("<lvml:"+t+' class="lvml">')}}catch(t){}return function(t){return document.createElement("<"+t+' xmlns="urn:schemas-microsoft.com:vml" class="lvml">')}}(),zt={_initContainer:function(){this._container=P("div","leaflet-vml-container")},_update:function(){this._map._animatingZoom||(Wi.prototype._update.call(this),this.fire("update"))},_initPath:function(t){var e=t._container=Vi("shape");M(e,"leaflet-vml-shape "+(this.options.className||"")),e.coordsize="1 1",t._path=Vi("path"),e.appendChild(t._path),this._updateStyle(t),this._layers[h(t)]=t},_addPath:function(t){var e=t._container;this._container.appendChild(e),t.options.interactive&&t.addInteractiveTarget(e)},_removePath:function(t){var e=t._container;T(e),t.removeInteractiveTarget(e),delete this._layers[h(t)]},_updateStyle:function(t){var e=t._stroke,i=t._fill,n=t.options,o=t._container;o.stroked=!!n.stroke,o.filled=!!n.fill,n.stroke?(e=e||(t._stroke=Vi("stroke")),o.appendChild(e),e.weight=n.weight+"px",e.color=n.color,e.opacity=n.opacity,n.dashArray?e.dashStyle=d(n.dashArray)?n.dashArray.join(" "):n.dashArray.replace(/( *, *)/g," "):e.dashStyle="",e.endcap=n.lineCap.replace("butt","flat"),e.joinstyle=n.lineJoin):e&&(o.removeChild(e),t._stroke=null),n.fill?(i=i||(t._fill=Vi("fill")),o.appendChild(i),i.color=n.fillColor||n.color,i.opacity=n.fillOpacity):i&&(o.removeChild(i),t._fill=null)},_updateCircle:function(t){var e=t._point.round(),i=Math.round(t._radius),n=Math.round(t._radiusY||i);this._setPath(t,t._empty()?"M0 0":"AL "+e.x+","+e.y+" "+i+","+n+" 0,23592600")},_setPath:function(t,e){t._path.v=e},_bringToFront:function(t){fe(t._container)},_bringToBack:function(t){ge(t._container)}},qi=b.vml?Vi:ct,Gi=Wi.extend({_initContainer:function(){this._container=qi("svg"),this._container.setAttribute("pointer-events","none"),this._rootGroup=qi("g"),this._container.appendChild(this._rootGroup)},_destroyContainer:function(){T(this._container),k(this._container),delete this._container,delete this._rootGroup,delete this._svgSize},_update:function(){var t,e,i;this._map._animatingZoom&&this._bounds||(Wi.prototype._update.call(this),e=(t=this._bounds).getSize(),i=this._container,this._svgSize&&this._svgSize.equals(e)||(this._svgSize=e,i.setAttribute("width",e.x),i.setAttribute("height",e.y)),Z(i,t.min),i.setAttribute("viewBox",[t.min.x,t.min.y,e.x,e.y].join(" ")),this.fire("update"))},_initPath:function(t){var e=t._path=qi("path");t.options.className&&M(e,t.options.className),t.options.interactive&&M(e,"leaflet-interactive"),this._updateStyle(t),this._layers[h(t)]=t},_addPath:function(t){this._rootGroup||this._initContainer(),this._rootGroup.appendChild(t._path),t.addInteractiveTarget(t._path)},_removePath:function(t){T(t._path),t.removeInteractiveTarget(t._path),delete this._layers[h(t)]},_updatePath:function(t){t._project(),t._update()},_updateStyle:function(t){var e=t._path,t=t.options;e&&(t.stroke?(e.setAttribute("stroke",t.color),e.setAttribute("stroke-opacity",t.opacity),e.setAttribute("stroke-width",t.weight),e.setAttribute("stroke-linecap",t.lineCap),e.setAttribute("stroke-linejoin",t.lineJoin),t.dashArray?e.setAttribute("stroke-dasharray",t.dashArray):e.removeAttribute("stroke-dasharray"),t.dashOffset?e.setAttribute("stroke-dashoffset",t.dashOffset):e.removeAttribute("stroke-dashoffset")):e.setAttribute("stroke","none"),t.fill?(e.setAttribute("fill",t.fillColor||t.color),e.setAttribute("fill-opacity",t.fillOpacity),e.setAttribute("fill-rule",t.fillRule||"evenodd")):e.setAttribute("fill","none"))},_updatePoly:function(t,e){this._setPath(t,dt(t._parts,e))},_updateCircle:function(t){var e=t._point,i=Math.max(Math.round(t._radius),1),n="a"+i+","+(Math.max(Math.round(t._radiusY),1)||i)+" 0 1,0 ",e=t._empty()?"M0 0":"M"+(e.x-i)+","+e.y+n+2*i+",0 "+n+2*-i+",0 ";this._setPath(t,e)},_setPath:function(t,e){t._path.setAttribute("d",e)},_bringToFront:function(t){fe(t._path)},_bringToBack:function(t){ge(t._path)}});function Ki(t){return b.svg||b.vml?new Gi(t):null}b.vml&&Gi.include(zt),A.include({getRenderer:function(t){t=(t=t.options.renderer||this._getPaneRenderer(t.options.pane)||this.options.renderer||this._renderer)||(this._renderer=this._createRenderer());return this.hasLayer(t)||this.addLayer(t),t},_getPaneRenderer:function(t){var e;return"overlayPane"!==t&&void 0!==t&&(void 0===(e=this._paneRenderers[t])&&(e=this._createRenderer({pane:t}),this._paneRenderers[t]=e),e)},_createRenderer:function(t){return this.options.preferCanvas&&Ui(t)||Ki(t)}});var Yi=xi.extend({initialize:function(t,e){xi.prototype.initialize.call(this,this._boundsToLatLngs(t),e)},setBounds:function(t){return this.setLatLngs(this._boundsToLatLngs(t))},_boundsToLatLngs:function(t){return[(t=g(t)).getSouthWest(),t.getNorthWest(),t.getNorthEast(),t.getSouthEast()]}});Gi.create=qi,Gi.pointsToPath=dt,wi.geometryToLayer=bi,wi.coordsToLatLng=Li,wi.coordsToLatLngs=Ti,wi.latLngToCoords=Mi,wi.latLngsToCoords=zi,wi.getFeature=Ci,wi.asFeature=Zi,A.mergeOptions({boxZoom:!0});var _t=n.extend({initialize:function(t){this._map=t,this._container=t._container,this._pane=t._panes.overlayPane,this._resetStateTimeout=0,t.on("unload",this._destroy,this)},addHooks:function(){S(this._container,"mousedown",this._onMouseDown,this)},removeHooks:function(){k(this._container,"mousedown",this._onMouseDown,this)},moved:function(){return this._moved},_destroy:function(){T(this._pane),delete this._pane},_resetState:function(){this._resetStateTimeout=0,this._moved=!1},_clearDeferredResetState:function(){0!==this._resetStateTimeout&&(clearTimeout(this._resetStateTimeout),this._resetStateTimeout=0)},_onMouseDown:function(t){if(!t.shiftKey||1!==t.which&&1!==t.button)return!1;this._clearDeferredResetState(),this._resetState(),re(),Le(),this._startPoint=this._map.mouseEventToContainerPoint(t),S(document,{contextmenu:Re,mousemove:this._onMouseMove,mouseup:this._onMouseUp,keydown:this._onKeyDown},this)},_onMouseMove:function(t){this._moved||(this._moved=!0,this._box=P("div","leaflet-zoom-box",this._container),M(this._container,"leaflet-crosshair"),this._map.fire("boxzoomstart")),this._point=this._map.mouseEventToContainerPoint(t);var t=new f(this._point,this._startPoint),e=t.getSize();Z(this._box,t.min),this._box.style.width=e.x+"px",this._box.style.height=e.y+"px"},_finish:function(){this._moved&&(T(this._box),z(this._container,"leaflet-crosshair")),ae(),Te(),k(document,{contextmenu:Re,mousemove:this._onMouseMove,mouseup:this._onMouseUp,keydown:this._onKeyDown},this)},_onMouseUp:function(t){1!==t.which&&1!==t.button||(this._finish(),this._moved&&(this._clearDeferredResetState(),this._resetStateTimeout=setTimeout(a(this._resetState,this),0),t=new s(this._map.containerPointToLatLng(this._startPoint),this._map.containerPointToLatLng(this._point)),this._map.fitBounds(t).fire("boxzoomend",{boxZoomBounds:t})))},_onKeyDown:function(t){27===t.keyCode&&(this._finish(),this._clearDeferredResetState(),this._resetState())}}),Ct=(A.addInitHook("addHandler","boxZoom",_t),A.mergeOptions({doubleClickZoom:!0}),n.extend({addHooks:function(){this._map.on("dblclick",this._onDoubleClick,this)},removeHooks:function(){this._map.off("dblclick",this._onDoubleClick,this)},_onDoubleClick:function(t){var e=this._map,i=e.getZoom(),n=e.options.zoomDelta,i=t.originalEvent.shiftKey?i-n:i+n;"center"===e.options.doubleClickZoom?e.setZoom(i):e.setZoomAround(t.containerPoint,i)}})),Zt=(A.addInitHook("addHandler","doubleClickZoom",Ct),A.mergeOptions({dragging:!0,inertia:!0,inertiaDeceleration:3400,inertiaMaxSpeed:1/0,easeLinearity:.2,worldCopyJump:!1,maxBoundsViscosity:0}),n.extend({addHooks:function(){var t;this._draggable||(t=this._map,this._draggable=new Xe(t._mapPane,t._container),this._draggable.on({dragstart:this._onDragStart,drag:this._onDrag,dragend:this._onDragEnd},this),this._draggable.on("predrag",this._onPreDragLimit,this),t.options.worldCopyJump&&(this._draggable.on("predrag",this._onPreDragWrap,this),t.on("zoomend",this._onZoomEnd,this),t.whenReady(this._onZoomEnd,this))),M(this._map._container,"leaflet-grab leaflet-touch-drag"),this._draggable.enable(),this._positions=[],this._times=[]},removeHooks:function(){z(this._map._container,"leaflet-grab"),z(this._map._container,"leaflet-touch-drag"),this._draggable.disable()},moved:function(){return this._draggable&&this._draggable._moved},moving:function(){return this._draggable&&this._draggable._moving},_onDragStart:function(){var t,e=this._map;e._stop(),this._map.options.maxBounds&&this._map.options.maxBoundsViscosity?(t=g(this._map.options.maxBounds),this._offsetLimit=_(this._map.latLngToContainerPoint(t.getNorthWest()).multiplyBy(-1),this._map.latLngToContainerPoint(t.getSouthEast()).multiplyBy(-1).add(this._map.getSize())),this._viscosity=Math.min(1,Math.max(0,this._map.options.maxBoundsViscosity))):this._offsetLimit=null,e.fire("movestart").fire("dragstart"),e.options.inertia&&(this._positions=[],this._times=[])},_onDrag:function(t){var e,i;this._map.options.inertia&&(e=this._lastTime=+new Date,i=this._lastPos=this._draggable._absPos||this._draggable._newPos,this._positions.push(i),this._times.push(e),this._prunePositions(e)),this._map.fire("move",t).fire("drag",t)},_prunePositions:function(t){for(;1<this._positions.length&&50<t-this._times[0];)this._positions.shift(),this._times.shift()},_onZoomEnd:function(){var t=this._map.getSize().divideBy(2),e=this._map.latLngToLayerPoint([0,0]);this._initialWorldOffset=e.subtract(t).x,this._worldWidth=this._map.getPixelWorldBounds().getSize().x},_viscousLimit:function(t,e){return t-(t-e)*this._viscosity},_onPreDragLimit:function(){var t,e;this._viscosity&&this._offsetLimit&&(t=this._draggable._newPos.subtract(this._draggable._startPos),e=this._offsetLimit,t.x<e.min.x&&(t.x=this._viscousLimit(t.x,e.min.x)),t.y<e.min.y&&(t.y=this._viscousLimit(t.y,e.min.y)),t.x>e.max.x&&(t.x=this._viscousLimit(t.x,e.max.x)),t.y>e.max.y&&(t.y=this._viscousLimit(t.y,e.max.y)),this._draggable._newPos=this._draggable._startPos.add(t))},_onPreDragWrap:function(){var t=this._worldWidth,e=Math.round(t/2),i=this._initialWorldOffset,n=this._draggable._newPos.x,o=(n-e+i)%t+e-i,n=(n+e+i)%t-e-i,t=Math.abs(o+i)<Math.abs(n+i)?o:n;this._draggable._absPos=this._draggable._newPos.clone(),this._draggable._newPos.x=t},_onDragEnd:function(t){var e,i,n,o,s=this._map,r=s.options,a=!r.inertia||t.noInertia||this._times.length<2;s.fire("dragend",t),!a&&(this._prunePositions(+new Date),t=this._lastPos.subtract(this._positions[0]),a=(this._lastTime-this._times[0])/1e3,e=r.easeLinearity,a=(t=t.multiplyBy(e/a)).distanceTo([0,0]),i=Math.min(r.inertiaMaxSpeed,a),t=t.multiplyBy(i/a),n=i/(r.inertiaDeceleration*e),(o=t.multiplyBy(-n/2).round()).x||o.y)?(o=s._limitOffset(o,s.options.maxBounds),x(function(){s.panBy(o,{duration:n,easeLinearity:e,noMoveStart:!0,animate:!0})})):s.fire("moveend")}})),St=(A.addInitHook("addHandler","dragging",Zt),A.mergeOptions({keyboard:!0,keyboardPanDelta:80}),n.extend({keyCodes:{left:[37],right:[39],down:[40],up:[38],zoomIn:[187,107,61,171],zoomOut:[189,109,54,173]},initialize:function(t){this._map=t,this._setPanDelta(t.options.keyboardPanDelta),this._setZoomDelta(t.options.zoomDelta)},addHooks:function(){var t=this._map._container;t.tabIndex<=0&&(t.tabIndex="0"),S(t,{focus:this._onFocus,blur:this._onBlur,mousedown:this._onMouseDown},this),this._map.on({focus:this._addHooks,blur:this._removeHooks},this)},removeHooks:function(){this._removeHooks(),k(this._map._container,{focus:this._onFocus,blur:this._onBlur,mousedown:this._onMouseDown},this),this._map.off({focus:this._addHooks,blur:this._removeHooks},this)},_onMouseDown:function(){var t,e,i;this._focused||(i=document.body,t=document.documentElement,e=i.scrollTop||t.scrollTop,i=i.scrollLeft||t.scrollLeft,this._map._container.focus(),window.scrollTo(i,e))},_onFocus:function(){this._focused=!0,this._map.fire("focus")},_onBlur:function(){this._focused=!1,this._map.fire("blur")},_setPanDelta:function(t){for(var e=this._panKeys={},i=this.keyCodes,n=0,o=i.left.length;n<o;n++)e[i.left[n]]=[-1*t,0];for(n=0,o=i.right.length;n<o;n++)e[i.right[n]]=[t,0];for(n=0,o=i.down.length;n<o;n++)e[i.down[n]]=[0,t];for(n=0,o=i.up.length;n<o;n++)e[i.up[n]]=[0,-1*t]},_setZoomDelta:function(t){for(var e=this._zoomKeys={},i=this.keyCodes,n=0,o=i.zoomIn.length;n<o;n++)e[i.zoomIn[n]]=t;for(n=0,o=i.zoomOut.length;n<o;n++)e[i.zoomOut[n]]=-t},_addHooks:function(){S(document,"keydown",this._onKeyDown,this)},_removeHooks:function(){k(document,"keydown",this._onKeyDown,this)},_onKeyDown:function(t){if(!(t.altKey||t.ctrlKey||t.metaKey)){var e,i,n=t.keyCode,o=this._map;if(n in this._panKeys)o._panAnim&&o._panAnim._inProgress||(i=this._panKeys[n],t.shiftKey&&(i=m(i).multiplyBy(3)),o.options.maxBounds&&(i=o._limitOffset(m(i),o.options.maxBounds)),o.options.worldCopyJump?(e=o.wrapLatLng(o.unproject(o.project(o.getCenter()).add(i))),o.panTo(e)):o.panBy(i));else if(n in this._zoomKeys)o.setZoom(o.getZoom()+(t.shiftKey?3:1)*this._zoomKeys[n]);else{if(27!==n||!o._popup||!o._popup.options.closeOnEscapeKey)return;o.closePopup()}Re(t)}}})),Et=(A.addInitHook("addHandler","keyboard",St),A.mergeOptions({scrollWheelZoom:!0,wheelDebounceTime:40,wheelPxPerZoomLevel:60}),n.extend({addHooks:function(){S(this._map._container,"wheel",this._onWheelScroll,this),this._delta=0},removeHooks:function(){k(this._map._container,"wheel",this._onWheelScroll,this)},_onWheelScroll:function(t){var e=He(t),i=this._map.options.wheelDebounceTime,e=(this._delta+=e,this._lastMousePos=this._map.mouseEventToContainerPoint(t),this._startTime||(this._startTime=+new Date),Math.max(i-(+new Date-this._startTime),0));clearTimeout(this._timer),this._timer=setTimeout(a(this._performZoom,this),e),Re(t)},_performZoom:function(){var t=this._map,e=t.getZoom(),i=this._map.options.zoomSnap||0,n=(t._stop(),this._delta/(4*this._map.options.wheelPxPerZoomLevel)),n=4*Math.log(2/(1+Math.exp(-Math.abs(n))))/Math.LN2,i=i?Math.ceil(n/i)*i:n,n=t._limitZoom(e+(0<this._delta?i:-i))-e;this._delta=0,this._startTime=null,n&&("center"===t.options.scrollWheelZoom?t.setZoom(e+n):t.setZoomAround(this._lastMousePos,e+n))}})),kt=(A.addInitHook("addHandler","scrollWheelZoom",Et),A.mergeOptions({tapHold:b.touchNative&&b.safari&&b.mobile,tapTolerance:15}),n.extend({addHooks:function(){S(this._map._container,"touchstart",this._onDown,this)},removeHooks:function(){k(this._map._container,"touchstart",this._onDown,this)},_onDown:function(t){var e;clearTimeout(this._holdTimeout),1===t.touches.length&&(e=t.touches[0],this._startPos=this._newPos=new p(e.clientX,e.clientY),this._holdTimeout=setTimeout(a(function(){this._cancel(),this._isTapValid()&&(S(document,"touchend",O),S(document,"touchend touchcancel",this._cancelClickPrevent),this._simulateEvent("contextmenu",e))},this),600),S(document,"touchend touchcancel contextmenu",this._cancel,this),S(document,"touchmove",this._onMove,this))},_cancelClickPrevent:function t(){k(document,"touchend",O),k(document,"touchend touchcancel",t)},_cancel:function(){clearTimeout(this._holdTimeout),k(document,"touchend touchcancel contextmenu",this._cancel,this),k(document,"touchmove",this._onMove,this)},_onMove:function(t){t=t.touches[0];this._newPos=new p(t.clientX,t.clientY)},_isTapValid:function(){return this._newPos.distanceTo(this._startPos)<=this._map.options.tapTolerance},_simulateEvent:function(t,e){t=new MouseEvent(t,{bubbles:!0,cancelable:!0,view:window,screenX:e.screenX,screenY:e.screenY,clientX:e.clientX,clientY:e.clientY});t._simulated=!0,e.target.dispatchEvent(t)}})),Ot=(A.addInitHook("addHandler","tapHold",kt),A.mergeOptions({touchZoom:b.touch,bounceAtZoomLimits:!0}),n.extend({addHooks:function(){M(this._map._container,"leaflet-touch-zoom"),S(this._map._container,"touchstart",this._onTouchStart,this)},removeHooks:function(){z(this._map._container,"leaflet-touch-zoom"),k(this._map._container,"touchstart",this._onTouchStart,this)},_onTouchStart:function(t){var e,i,n=this._map;!t.touches||2!==t.touches.length||n._animatingZoom||this._zooming||(e=n.mouseEventToContainerPoint(t.touches[0]),i=n.mouseEventToContainerPoint(t.touches[1]),this._centerPoint=n.getSize()._divideBy(2),this._startLatLng=n.containerPointToLatLng(this._centerPoint),"center"!==n.options.touchZoom&&(this._pinchStartLatLng=n.containerPointToLatLng(e.add(i)._divideBy(2))),this._startDist=e.distanceTo(i),this._startZoom=n.getZoom(),this._moved=!1,this._zooming=!0,n._stop(),S(document,"touchmove",this._onTouchMove,this),S(document,"touchend touchcancel",this._onTouchEnd,this),O(t))},_onTouchMove:function(t){if(t.touches&&2===t.touches.length&&this._zooming){var e=this._map,i=e.mouseEventToContainerPoint(t.touches[0]),n=e.mouseEventToContainerPoint(t.touches[1]),o=i.distanceTo(n)/this._startDist;if(this._zoom=e.getScaleZoom(o,this._startZoom),!e.options.bounceAtZoomLimits&&(this._zoom<e.getMinZoom()&&o<1||this._zoom>e.getMaxZoom()&&1<o)&&(this._zoom=e._limitZoom(this._zoom)),"center"===e.options.touchZoom){if(this._center=this._startLatLng,1==o)return}else{i=i._add(n)._divideBy(2)._subtract(this._centerPoint);if(1==o&&0===i.x&&0===i.y)return;this._center=e.unproject(e.project(this._pinchStartLatLng,this._zoom).subtract(i),this._zoom)}this._moved||(e._moveStart(!0,!1),this._moved=!0),r(this._animRequest);n=a(e._move,e,this._center,this._zoom,{pinch:!0,round:!1},void 0);this._animRequest=x(n,this,!0),O(t)}},_onTouchEnd:function(){this._moved&&this._zooming?(this._zooming=!1,r(this._animRequest),k(document,"touchmove",this._onTouchMove,this),k(document,"touchend touchcancel",this._onTouchEnd,this),this._map.options.zoomAnimation?this._map._animateZoom(this._center,this._map._limitZoom(this._zoom),!0,this._map.options.zoomSnap):this._map._resetView(this._center,this._map._limitZoom(this._zoom))):this._zooming=!1}})),Xi=(A.addInitHook("addHandler","touchZoom",Ot),A.BoxZoom=_t,A.DoubleClickZoom=Ct,A.Drag=Zt,A.Keyboard=St,A.ScrollWheelZoom=Et,A.TapHold=kt,A.TouchZoom=Ot,t.Bounds=f,t.Browser=b,t.CRS=ot,t.Canvas=Fi,t.Circle=vi,t.CircleMarker=gi,t.Class=et,t.Control=B,t.DivIcon=Ri,t.DivOverlay=Ai,t.DomEvent=mt,t.DomUtil=pt,t.Draggable=Xe,t.Evented=it,t.FeatureGroup=ci,t.GeoJSON=wi,t.GridLayer=Ni,t.Handler=n,t.Icon=di,t.ImageOverlay=Ei,t.LatLng=v,t.LatLngBounds=s,t.Layer=o,t.LayerGroup=ui,t.LineUtil=vt,t.Map=A,t.Marker=mi,t.Mixin=ft,t.Path=fi,t.Point=p,t.PolyUtil=gt,t.Polygon=xi,t.Polyline=yi,t.Popup=Bi,t.PosAnimation=Fe,t.Projection=wt,t.Rectangle=Yi,t.Renderer=Wi,t.SVG=Gi,t.SVGOverlay=Oi,t.TileLayer=Di,t.Tooltip=Ii,t.Transformation=at,t.Util=tt,t.VideoOverlay=ki,t.bind=a,t.bounds=_,t.canvas=Ui,t.circle=function(t,e,i){return new vi(t,e,i)},t.circleMarker=function(t,e){return new gi(t,e)},t.control=Ue,t.divIcon=function(t){return new Ri(t)},t.extend=l,t.featureGroup=function(t,e){return new ci(t,e)},t.geoJSON=Si,t.geoJson=Mt,t.gridLayer=function(t){return new Ni(t)},t.icon=function(t){return new di(t)},t.imageOverlay=function(t,e,i){return new Ei(t,e,i)},t.latLng=w,t.latLngBounds=g,t.layerGroup=function(t,e){return new ui(t,e)},t.map=function(t,e){return new A(t,e)},t.marker=function(t,e){return new mi(t,e)},t.point=m,t.polygon=function(t,e){return new xi(t,e)},t.polyline=function(t,e){return new yi(t,e)},t.popup=function(t,e){return new Bi(t,e)},t.rectangle=function(t,e){return new Yi(t,e)},t.setOptions=c,t.stamp=h,t.svg=Ki,t.svgOverlay=function(t,e,i){return new Oi(t,e,i)},t.tileLayer=ji,t.tooltip=function(t,e){return new Ii(t,e)},t.transformation=ht,t.version="1.9.4",t.videoOverlay=function(t,e,i){return new ki(t,e,i)},window.L);t.noConflict=function(){return window.L=Xi,this},window.L=t});
//# sourceMappingURL=leaflet.js.map
</script>
<script>
/* Leaflet.heat v0.2.0 — embebido offline */
/*
 (c) 2014, Vladimir Agafonkin
 simpleheat, a tiny JavaScript library for drawing heatmaps with Canvas
 https://github.com/mourner/simpleheat
*/
!function(){"use strict";function t(i){return this instanceof t?(this._canvas=i="string"==typeof i?document.getElementById(i):i,this._ctx=i.getContext("2d"),this._width=i.width,this._height=i.height,this._max=1,void this.clear()):new t(i)}t.prototype={defaultRadius:25,defaultGradient:{.4:"blue",.6:"cyan",.7:"lime",.8:"yellow",1:"red"},data:function(t,i){return this._data=t,this},max:function(t){return this._max=t,this},add:function(t){return this._data.push(t),this},clear:function(){return this._data=[],this},radius:function(t,i){i=i||15;var a=this._circle=document.createElement("canvas"),s=a.getContext("2d"),e=this._r=t+i;return a.width=a.height=2*e,s.shadowOffsetX=s.shadowOffsetY=200,s.shadowBlur=i,s.shadowColor="black",s.beginPath(),s.arc(e-200,e-200,t,0,2*Math.PI,!0),s.closePath(),s.fill(),this},gradient:function(t){var i=document.createElement("canvas"),a=i.getContext("2d"),s=a.createLinearGradient(0,0,0,256);i.width=1,i.height=256;for(var e in t)s.addColorStop(e,t[e]);return a.fillStyle=s,a.fillRect(0,0,1,256),this._grad=a.getImageData(0,0,1,256).data,this},draw:function(t){this._circle||this.radius(this.defaultRadius),this._grad||this.gradient(this.defaultGradient);var i=this._ctx;i.clearRect(0,0,this._width,this._height);for(var a,s=0,e=this._data.length;e>s;s++)a=this._data[s],i.globalAlpha=Math.max(a[2]/this._max,t||.05),i.drawImage(this._circle,a[0]-this._r,a[1]-this._r);var n=i.getImageData(0,0,this._width,this._height);return this._colorize(n.data,this._grad),i.putImageData(n,0,0),this},_colorize:function(t,i){for(var a,s=3,e=t.length;e>s;s+=4)a=4*t[s],a&&(t[s-3]=i[a],t[s-2]=i[a+1],t[s-1]=i[a+2])}},window.simpleheat=t}(),/*
 (c) 2014, Vladimir Agafonkin
 Leaflet.heat, a tiny and fast heatmap plugin for Leaflet.
 https://github.com/Leaflet/Leaflet.heat
*/
L.HeatLayer=(L.Layer?L.Layer:L.Class).extend({initialize:function(t,i){this._latlngs=t,L.setOptions(this,i)},setLatLngs:function(t){return this._latlngs=t,this.redraw()},addLatLng:function(t){return this._latlngs.push(t),this.redraw()},setOptions:function(t){return L.setOptions(this,t),this._heat&&this._updateOptions(),this.redraw()},redraw:function(){return!this._heat||this._frame||this._map._animating||(this._frame=L.Util.requestAnimFrame(this._redraw,this)),this},onAdd:function(t){this._map=t,this._canvas||this._initCanvas(),t._panes.overlayPane.appendChild(this._canvas),t.on("moveend",this._reset,this),t.options.zoomAnimation&&L.Browser.any3d&&t.on("zoomanim",this._animateZoom,this),this._reset()},onRemove:function(t){t.getPanes().overlayPane.removeChild(this._canvas),t.off("moveend",this._reset,this),t.options.zoomAnimation&&t.off("zoomanim",this._animateZoom,this)},addTo:function(t){return t.addLayer(this),this},_initCanvas:function(){var t=this._canvas=L.DomUtil.create("canvas","leaflet-heatmap-layer leaflet-layer"),i=L.DomUtil.testProp(["transformOrigin","WebkitTransformOrigin","msTransformOrigin"]);t.style[i]="50% 50%";var a=this._map.getSize();t.width=a.x,t.height=a.y;var s=this._map.options.zoomAnimation&&L.Browser.any3d;L.DomUtil.addClass(t,"leaflet-zoom-"+(s?"animated":"hide")),this._heat=simpleheat(t),this._updateOptions()},_updateOptions:function(){this._heat.radius(this.options.radius||this._heat.defaultRadius,this.options.blur),this.options.gradient&&this._heat.gradient(this.options.gradient),this.options.max&&this._heat.max(this.options.max)},_reset:function(){var t=this._map.containerPointToLayerPoint([0,0]);L.DomUtil.setPosition(this._canvas,t);var i=this._map.getSize();this._heat._width!==i.x&&(this._canvas.width=this._heat._width=i.x),this._heat._height!==i.y&&(this._canvas.height=this._heat._height=i.y),this._redraw()},_redraw:function(){var t,i,a,s,e,n,h,o,r,d=[],_=this._heat._r,l=this._map.getSize(),m=new L.Bounds(L.point([-_,-_]),l.add([_,_])),c=void 0===this.options.max?1:this.options.max,u=void 0===this.options.maxZoom?this._map.getMaxZoom():this.options.maxZoom,f=1/Math.pow(2,Math.max(0,Math.min(u-this._map.getZoom(),12))),g=_/2,p=[],v=this._map._getMapPanePos(),w=v.x%g,y=v.y%g;for(t=0,i=this._latlngs.length;i>t;t++)if(a=this._map.latLngToContainerPoint(this._latlngs[t]),m.contains(a)){e=Math.floor((a.x-w)/g)+2,n=Math.floor((a.y-y)/g)+2;var x=void 0!==this._latlngs[t].alt?this._latlngs[t].alt:void 0!==this._latlngs[t][2]?+this._latlngs[t][2]:1;r=x*f,p[n]=p[n]||[],s=p[n][e],s?(s[0]=(s[0]*s[2]+a.x*r)/(s[2]+r),s[1]=(s[1]*s[2]+a.y*r)/(s[2]+r),s[2]+=r):p[n][e]=[a.x,a.y,r]}for(t=0,i=p.length;i>t;t++)if(p[t])for(h=0,o=p[t].length;o>h;h++)s=p[t][h],s&&d.push([Math.round(s[0]),Math.round(s[1]),Math.min(s[2],c)]);this._heat.data(d).draw(this.options.minOpacity),this._frame=null},_animateZoom:function(t){var i=this._map.getZoomScale(t.zoom),a=this._map._getCenterOffset(t.center)._multiplyBy(-i).subtract(this._map._getMapPanePos());L.DomUtil.setTransform?L.DomUtil.setTransform(this._canvas,a,i):this._canvas.style[L.DomUtil.TRANSFORM]=L.DomUtil.getTranslateString(a)+" scale("+i+")"}}),L.heatLayer=function(t,i){return new L.HeatLayer(t,i)};
</script>
<script>
/* chartjs-plugin-datalabels v2.2.0 — embebido offline */
/*!
 * chartjs-plugin-datalabels v2.2.0
 * https://chartjs-plugin-datalabels.netlify.app
 * (c) 2017-2022 chartjs-plugin-datalabels contributors
 * Released under the MIT license
 */
!function(t,e){"object"==typeof exports&&"undefined"!=typeof module?module.exports=e(require("chart.js/helpers"),require("chart.js")):"function"==typeof define&&define.amd?define(["chart.js/helpers","chart.js"],e):(t="undefined"!=typeof globalThis?globalThis:t||self).ChartDataLabels=e(t.Chart.helpers,t.Chart)}(this,(function(t,e){"use strict";var r=function(){if("undefined"!=typeof window){if(window.devicePixelRatio)return window.devicePixelRatio;var t=window.screen;if(t)return(t.deviceXDPI||1)/(t.logicalXDPI||1)}return 1}(),a=function(e){var r,a=[];for(e=[].concat(e);e.length;)"string"==typeof(r=e.pop())?a.unshift.apply(a,r.split("\n")):Array.isArray(r)?e.push.apply(e,r):t.isNullOrUndef(e)||a.unshift(""+r);return a},o=function(t,e,r){var a,o=[].concat(e),n=o.length,i=t.font,l=0;for(t.font=r.string,a=0;a<n;++a)l=Math.max(t.measureText(o[a]).width,l);return t.font=i,{height:n*r.lineHeight,width:l}},n=function(t,e,r){return Math.max(t,Math.min(e,r))},i=function(t,e){var r,a,o,n,i=t.slice(),l=[];for(r=0,o=e.length;r<o;++r)n=e[r],-1===(a=i.indexOf(n))?l.push([n,1]):i.splice(a,1);for(r=0,o=i.length;r<o;++r)l.push([i[r],-1]);return l};function l(t,e){var r=e.x,a=e.y;if(null===r)return{x:0,y:-1};if(null===a)return{x:1,y:0};var o=t.x-r,n=t.y-a,i=Math.sqrt(o*o+n*n);return{x:i?o/i:0,y:i?n/i:-1}}function s(t,e,r){var a=0;return t<r.left?a|=1:t>r.right&&(a|=2),e<r.top?a|=8:e>r.bottom&&(a|=4),a}function u(t,e){var r,a,o=e.anchor,n=t;return e.clamp&&(n=function(t,e){for(var r,a,o,n=t.x0,i=t.y0,l=t.x1,u=t.y1,d=s(n,i,e),c=s(l,u,e);d|c&&!(d&c);)8&(r=d||c)?(a=n+(l-n)*(e.top-i)/(u-i),o=e.top):4&r?(a=n+(l-n)*(e.bottom-i)/(u-i),o=e.bottom):2&r?(o=i+(u-i)*(e.right-n)/(l-n),a=e.right):1&r&&(o=i+(u-i)*(e.left-n)/(l-n),a=e.left),r===d?d=s(n=a,i=o,e):c=s(l=a,u=o,e);return{x0:n,x1:l,y0:i,y1:u}}(n,e.area)),"start"===o?(r=n.x0,a=n.y0):"end"===o?(r=n.x1,a=n.y1):(r=(n.x0+n.x1)/2,a=(n.y0+n.y1)/2),function(t,e,r,a,o){switch(o){case"center":r=a=0;break;case"bottom":r=0,a=1;break;case"right":r=1,a=0;break;case"left":r=-1,a=0;break;case"top":r=0,a=-1;break;case"start":r=-r,a=-a;break;case"end":break;default:o*=Math.PI/180,r=Math.cos(o),a=Math.sin(o)}return{x:t,y:e,vx:r,vy:a}}(r,a,t.vx,t.vy,e.align)}var d=function(t,e){var r=(t.startAngle+t.endAngle)/2,a=Math.cos(r),o=Math.sin(r),n=t.innerRadius,i=t.outerRadius;return u({x0:t.x+a*n,y0:t.y+o*n,x1:t.x+a*i,y1:t.y+o*i,vx:a,vy:o},e)},c=function(t,e){var r=l(t,e.origin),a=r.x*t.options.radius,o=r.y*t.options.radius;return u({x0:t.x-a,y0:t.y-o,x1:t.x+a,y1:t.y+o,vx:r.x,vy:r.y},e)},h=function(t,e){var r=l(t,e.origin),a=t.x,o=t.y,n=0,i=0;return t.horizontal?(a=Math.min(t.x,t.base),n=Math.abs(t.base-t.x)):(o=Math.min(t.y,t.base),i=Math.abs(t.base-t.y)),u({x0:a,y0:o+i,x1:a+n,y1:o,vx:r.x,vy:r.y},e)},f=function(t,e){var r=l(t,e.origin);return u({x0:t.x,y0:t.y,x1:t.x+(t.width||0),y1:t.y+(t.height||0),vx:r.x,vy:r.y},e)},x=function(t){return Math.round(t*r)/r};function y(t,e){var r=e.chart.getDatasetMeta(e.datasetIndex).vScale;if(!r)return null;if(void 0!==r.xCenter&&void 0!==r.yCenter)return{x:r.xCenter,y:r.yCenter};var a=r.getBasePixel();return t.horizontal?{x:a,y:null}:{x:null,y:a}}function v(t,e,r){var a=r.backgroundColor,o=r.borderColor,n=r.borderWidth;(a||o&&n)&&(t.beginPath(),function(t,e,r,a,o,n){var i=Math.PI/2;if(n){var l=Math.min(n,o/2,a/2),s=e+l,u=r+l,d=e+a-l,c=r+o-l;t.moveTo(e,u),s<d&&u<c?(t.arc(s,u,l,-Math.PI,-i),t.arc(d,u,l,-i,0),t.arc(d,c,l,0,i),t.arc(s,c,l,i,Math.PI)):s<d?(t.moveTo(s,r),t.arc(d,u,l,-i,i),t.arc(s,u,l,i,Math.PI+i)):u<c?(t.arc(s,u,l,-Math.PI,0),t.arc(s,c,l,0,Math.PI)):t.arc(s,u,l,-Math.PI,Math.PI),t.closePath(),t.moveTo(e,r)}else t.rect(e,r,a,o)}(t,x(e.x)+n/2,x(e.y)+n/2,x(e.w)-n,x(e.h)-n,r.borderRadius),t.closePath(),a&&(t.fillStyle=a,t.fill()),o&&n&&(t.strokeStyle=o,t.lineWidth=n,t.lineJoin="miter",t.stroke()))}function b(t,e,r){var a=t.shadowBlur,o=r.stroked,n=x(r.x),i=x(r.y),l=x(r.w);o&&t.strokeText(e,n,i,l),r.filled&&(a&&o&&(t.shadowBlur=0),t.fillText(e,n,i,l),a&&o&&(t.shadowBlur=a))}var _=function(t,e,r,a){var o=this;o._config=t,o._index=a,o._model=null,o._rects=null,o._ctx=e,o._el=r};t.merge(_.prototype,{_modelize:function(r,a,n,i){var l,s=this,u=s._index,x=t.toFont(t.resolve([n.font,{}],i,u)),v=t.resolve([n.color,e.defaults.color],i,u);return{align:t.resolve([n.align,"center"],i,u),anchor:t.resolve([n.anchor,"center"],i,u),area:i.chart.chartArea,backgroundColor:t.resolve([n.backgroundColor,null],i,u),borderColor:t.resolve([n.borderColor,null],i,u),borderRadius:t.resolve([n.borderRadius,0],i,u),borderWidth:t.resolve([n.borderWidth,0],i,u),clamp:t.resolve([n.clamp,!1],i,u),clip:t.resolve([n.clip,!1],i,u),color:v,display:r,font:x,lines:a,offset:t.resolve([n.offset,4],i,u),opacity:t.resolve([n.opacity,1],i,u),origin:y(s._el,i),padding:t.toPadding(t.resolve([n.padding,4],i,u)),positioner:(l=s._el,l instanceof e.ArcElement?d:l instanceof e.PointElement?c:l instanceof e.BarElement?h:f),rotation:t.resolve([n.rotation,0],i,u)*(Math.PI/180),size:o(s._ctx,a,x),textAlign:t.resolve([n.textAlign,"start"],i,u),textShadowBlur:t.resolve([n.textShadowBlur,0],i,u),textShadowColor:t.resolve([n.textShadowColor,v],i,u),textStrokeColor:t.resolve([n.textStrokeColor,v],i,u),textStrokeWidth:t.resolve([n.textStrokeWidth,0],i,u)}},update:function(e){var r,o,n,i=this,l=null,s=null,u=i._index,d=i._config,c=t.resolve([d.display,!0],e,u);c&&(r=e.dataset.data[u],o=t.valueOrDefault(t.callback(d.formatter,[r,e]),r),(n=t.isNullOrUndef(o)?[]:a(o)).length&&(s=function(t){var e=t.borderWidth||0,r=t.padding,a=t.size.height,o=t.size.width,n=-o/2,i=-a/2;return{frame:{x:n-r.left-e,y:i-r.top-e,w:o+r.width+2*e,h:a+r.height+2*e},text:{x:n,y:i,w:o,h:a}}}(l=i._modelize(c,n,d,e)))),i._model=l,i._rects=s},geometry:function(){return this._rects?this._rects.frame:{}},rotation:function(){return this._model?this._model.rotation:0},visible:function(){return this._model&&this._model.opacity},model:function(){return this._model},draw:function(t,e){var r,a=t.ctx,o=this._model,i=this._rects;this.visible()&&(a.save(),o.clip&&(r=o.area,a.beginPath(),a.rect(r.left,r.top,r.right-r.left,r.bottom-r.top),a.clip()),a.globalAlpha=n(0,o.opacity,1),a.translate(x(e.x),x(e.y)),a.rotate(o.rotation),v(a,i.frame,o),function(t,e,r,a){var o,n=a.textAlign,i=a.color,l=!!i,s=a.font,u=e.length,d=a.textStrokeColor,c=a.textStrokeWidth,h=d&&c;if(u&&(l||h))for(r=function(t,e,r){var a=r.lineHeight,o=t.w,n=t.x;return"center"===e?n+=o/2:"end"!==e&&"right"!==e||(n+=o),{h:a,w:o,x:n,y:t.y+a/2}}(r,n,s),t.font=s.string,t.textAlign=n,t.textBaseline="middle",t.shadowBlur=a.textShadowBlur,t.shadowColor=a.textShadowColor,l&&(t.fillStyle=i),h&&(t.lineJoin="round",t.lineWidth=c,t.strokeStyle=d),o=0,u=e.length;o<u;++o)b(t,e[o],{stroked:h,filled:l,w:r.w,x:r.x,y:r.y+r.h*o})}(a,o.lines,i.text,o),a.restore())}});var p=Number.MIN_SAFE_INTEGER||-9007199254740991,g=Number.MAX_SAFE_INTEGER||9007199254740991;function m(t,e,r){var a=Math.cos(r),o=Math.sin(r),n=e.x,i=e.y;return{x:n+a*(t.x-n)-o*(t.y-i),y:i+o*(t.x-n)+a*(t.y-i)}}function w(t,e){var r,a,o,n,i,l=g,s=p,u=e.origin;for(r=0;r<t.length;++r)o=(a=t[r]).x-u.x,n=a.y-u.y,i=e.vx*o+e.vy*n,l=Math.min(l,i),s=Math.max(s,i);return{min:l,max:s}}function M(t,e){var r=e.x-t.x,a=e.y-t.y,o=Math.sqrt(r*r+a*a);return{vx:(e.x-t.x)/o,vy:(e.y-t.y)/o,origin:t,ln:o}}var k=function(){this._rotation=0,this._rect={x:0,y:0,w:0,h:0}};function $(t,e,r){var a=e.positioner(t,e),o=a.vx,n=a.vy;if(!o&&!n)return{x:a.x,y:a.y};var i=r.w,l=r.h,s=e.rotation,u=Math.abs(i/2*Math.cos(s))+Math.abs(l/2*Math.sin(s)),d=Math.abs(i/2*Math.sin(s))+Math.abs(l/2*Math.cos(s)),c=1/Math.max(Math.abs(o),Math.abs(n));return u*=o*c,d*=n*c,u+=e.offset*o,d+=e.offset*n,{x:a.x+u,y:a.y+d}}t.merge(k.prototype,{center:function(){var t=this._rect;return{x:t.x+t.w/2,y:t.y+t.h/2}},update:function(t,e,r){this._rotation=r,this._rect={x:e.x+t.x,y:e.y+t.y,w:e.w,h:e.h}},contains:function(t){var e=this,r=e._rect;return!((t=m(t,e.center(),-e._rotation)).x<r.x-1||t.y<r.y-1||t.x>r.x+r.w+2||t.y>r.y+r.h+2)},intersects:function(t){var e,r,a,o=this._points(),n=t._points(),i=[M(o[0],o[1]),M(o[0],o[3])];for(this._rotation!==t._rotation&&i.push(M(n[0],n[1]),M(n[0],n[3])),e=0;e<i.length;++e)if(r=w(o,i[e]),a=w(n,i[e]),r.max<a.min||a.max<r.min)return!1;return!0},_points:function(){var t=this,e=t._rect,r=t._rotation,a=t.center();return[m({x:e.x,y:e.y},a,r),m({x:e.x+e.w,y:e.y},a,r),m({x:e.x+e.w,y:e.y+e.h},a,r),m({x:e.x,y:e.y+e.h},a,r)]}});var C={prepare:function(t){var e,r,a,o,n,i=[];for(e=0,a=t.length;e<a;++e)for(r=0,o=t[e].length;r<o;++r)n=t[e][r],i.push(n),n.$layout={_box:new k,_hidable:!1,_visible:!0,_set:e,_idx:n._index};return i.sort((function(t,e){var r=t.$layout,a=e.$layout;return r._idx===a._idx?a._set-r._set:a._idx-r._idx})),this.update(i),i},update:function(t){var e,r,a,o,n,i=!1;for(e=0,r=t.length;e<r;++e)o=(a=t[e]).model(),(n=a.$layout)._hidable=o&&"auto"===o.display,n._visible=a.visible(),i|=n._hidable;i&&function(t){var e,r,a,o,n,i,l;for(e=0,r=t.length;e<r;++e)(o=(a=t[e]).$layout)._visible&&(l=new Proxy(a._el,{get:(t,e)=>t.getProps([e],!0)[e]}),n=a.geometry(),i=$(l,a.model(),n),o._box.update(i,n,a.rotation()));(function(t,e){var r,a,o,n;for(r=t.length-1;r>=0;--r)for(o=t[r].$layout,a=r-1;a>=0&&o._visible;--a)(n=t[a].$layout)._visible&&o._box.intersects(n._box)&&e(o,n)})(t,(function(t,e){var r=t._hidable,a=e._hidable;r&&a||a?e._visible=!1:r&&(t._visible=!1)}))}(t)},lookup:function(t,e){var r,a;for(r=t.length-1;r>=0;--r)if((a=t[r].$layout)&&a._visible&&a._box.contains(e))return t[r];return null},draw:function(t,e){var r,a,o,n,i,l;for(r=0,a=e.length;r<a;++r)(n=(o=e[r]).$layout)._visible&&(i=o.geometry(),l=$(o._el,o.model(),i),n._box.update(l,i,o.rotation()),o.draw(t,l))}},P="$default";function S(e,r,a,o){if(r){var n,i=a.$context,l=a.$groups;r[l._set]&&(n=r[l._set][l._key])&&!0===t.callback(n,[i,o])&&(e.$datalabels._dirty=!0,a.update(i))}}function I(t,e){var r,a,o=t.$datalabels,n=o._listeners;if(n.enter||n.leave){if("mousemove"===e.type)a=C.lookup(o._labels,e);else if("mouseout"!==e.type)return;r=o._hovered,o._hovered=a,function(t,e,r,a,o){var n,i;(r||a)&&(r?a?r!==a&&(i=n=!0):i=!0:n=!0,i&&S(t,e.leave,r,o),n&&S(t,e.enter,a,o))}(t,n,r,a,e)}}return{id:"datalabels",defaults:{align:"center",anchor:"center",backgroundColor:null,borderColor:null,borderRadius:0,borderWidth:0,clamp:!1,clip:!1,color:void 0,display:!0,font:{family:void 0,lineHeight:1.2,size:void 0,style:void 0,weight:null},formatter:function(e){if(t.isNullOrUndef(e))return null;var r,a,o,n=e;if(t.isObject(e))if(t.isNullOrUndef(e.label))if(t.isNullOrUndef(e.r))for(n="",o=0,a=(r=Object.keys(e)).length;o<a;++o)n+=(0!==o?", ":"")+r[o]+": "+e[r[o]];else n=e.r;else n=e.label;return""+n},labels:void 0,listeners:{},offset:4,opacity:1,padding:{top:4,right:4,bottom:4,left:4},rotation:0,textAlign:"start",textStrokeColor:void 0,textStrokeWidth:0,textShadowBlur:0,textShadowColor:void 0},beforeInit:function(t){t.$datalabels={_actives:[]}},beforeUpdate:function(t){var e=t.$datalabels;e._listened=!1,e._listeners={},e._datasets=[],e._labels=[]},afterDatasetUpdate:function(e,r,a){var o,n,i,l,s,u,d,c,h=r.index,f=e.$datalabels,x=f._datasets[h]=[],y=e.isDatasetVisible(h),v=e.data.datasets[h],b=function(e,r){var a,o,n,i=e.datalabels,l=[];return!1===i?null:(!0===i&&(i={}),r=t.merge({},[r,i]),o=r.labels||{},n=Object.keys(o),delete r.labels,n.length?n.forEach((function(e){o[e]&&l.push(t.merge({},[r,o[e],{_key:e}]))})):l.push(r),a=l.reduce((function(e,r){return t.each(r.listeners||{},(function(t,a){e[a]=e[a]||{},e[a][r._key||P]=t})),delete r.listeners,e}),{}),{labels:l,listeners:a})}(v,a),p=r.meta.data||[],g=e.ctx;for(g.save(),o=0,i=p.length;o<i;++o)if((d=p[o]).$datalabels=[],y&&d&&e.getDataVisibility(o)&&!d.skip)for(n=0,l=b.labels.length;n<l;++n)u=(s=b.labels[n])._key,(c=new _(s,g,d,o)).$groups={_set:h,_key:u||P},c.$context={active:!1,chart:e,dataIndex:o,dataset:v,datasetIndex:h},c.update(c.$context),d.$datalabels.push(c),x.push(c);g.restore(),t.merge(f._listeners,b.listeners,{merger:function(t,e,a){e[t]=e[t]||{},e[t][r.index]=a[t],f._listened=!0}})},afterUpdate:function(t){t.$datalabels._labels=C.prepare(t.$datalabels._datasets)},afterDatasetsDraw:function(t){C.draw(t,t.$datalabels._labels)},beforeEvent:function(t,e){if(t.$datalabels._listened){var r=e.event;switch(r.type){case"mousemove":case"mouseout":I(t,r);break;case"click":!function(t,e){var r=t.$datalabels,a=r._listeners.click,o=a&&C.lookup(r._labels,e);o&&S(t,a,o,e)}(t,r)}}},afterEvent:function(t){var e,r,a,o,n,l,s,u=t.$datalabels,d=u._actives,c=u._actives=t.getActiveElements(),h=i(d,c);for(e=0,r=h.length;e<r;++e)if((n=h[e])[1])for(a=0,o=(s=n[0].element.$datalabels||[]).length;a<o;++a)(l=s[a]).$context.active=1===n[1],l.update(l.$context);(u._dirty||h.length)&&(C.update(u._labels),t.render()),delete u._dirty}}}));

</script>
<style>
:root{
  --bg:#f3f4f7;--card:#fff;--sb:#12192b;--sb2:#0c1120;
  --acc:#2b5fd9;--acc2:#1d4ed8;--grn:#1a8a53;--red:#d33a3a;--orn:#c8860a;
  --tx:#1c2433;--mu:#707c8c;--br:#e6e8ee;--r:12px;
  --ml:#e0951f;--tn:#6d43c9;--cx:#0f8f83;
  --shadow:0 1px 2px rgba(20,25,40,.04),0 2px 10px rgba(20,25,40,.06);
  --shadow-hover:0 2px 4px rgba(20,25,40,.06),0 6px 18px rgba(20,25,40,.09);
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Segoe UI",system-ui,-apple-system,"Helvetica Neue",Arial,sans-serif;font-size:13px;
  background:var(--bg);color:var(--tx);height:100vh;display:flex;
  flex-direction:column;overflow:hidden;-webkit-font-smoothing:antialiased;letter-spacing:.01em}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-thumb{background:#c8d3de;border-radius:3px}
/* Topbar */
#top{min-height:54px;background:var(--sb);display:flex;align-items:center;
  gap:8px;padding:0 14px;flex-shrink:0;flex-wrap:wrap;
  box-shadow:0 2px 8px rgba(0,0,0,.18);position:relative;z-index:5}
.logo{display:flex;align-items:center;gap:9px;color:#fff;font-weight:700;font-size:15.5px;letter-spacing:-.01em}
.chips{display:flex;gap:3px}
.chip{font-size:9px;font-weight:800;padding:2px 6px;border-radius:3px;letter-spacing:.05em}
.chip-ml{background:var(--ml);color:#1a1a1a}
.chip-tn{background:var(--tn);color:#fff}
.chip-cx{background:var(--cx);color:#fff}
.tbtn{border:none;color:#fff;
  padding:7px 14px;border-radius:7px;cursor:pointer;font-size:12px;font-weight:600;
  transition:all .15s;display:inline-flex;align-items:center;gap:6px;letter-spacing:.01em}
.tbtn.ml{background:var(--ml);color:#fff}.tbtn.ml:hover{background:#c07d18}
.tbtn.tn{background:var(--tn);color:#fff}.tbtn.tn:hover{background:#5c37ab}
.tbtn.cx{background:var(--cx);color:#fff}.tbtn.cx:hover{background:#0c7a70}
/* Layout */
#main{display:flex;flex:1;overflow:hidden}
#sb{width:234px;background:var(--sb2);display:flex;flex-direction:column;overflow-y:auto;flex-shrink:0}
#ct{flex:1;overflow-y:auto;padding:14px 16px}
/* Sidebar */
.sb-inner{padding:13px 11px}
.ss{font-size:9px;font-weight:700;letter-spacing:.1em;color:rgba(255,255,255,.24);
  margin-bottom:6px;text-transform:uppercase}
.src-area{border-radius:8px;overflow:hidden;border:1px solid rgba(255,255,255,.07);margin-bottom:8px}
.src-head{display:flex;align-items:center;gap:7px;padding:8px 10px}
.src-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0}
.src-dot.ml{background:var(--ml)}.src-dot.tn{background:var(--tn)}.src-dot.cx{background:var(--cx)}
.src-nm{font-size:11px;font-weight:600;color:rgba(255,255,255,.75);flex:1}
.src-bd{font-size:9px;font-weight:700;padding:1px 5px;border-radius:3px}
.src-bd.ml{background:rgba(245,158,11,.18);color:var(--ml)}
.src-bd.tn{background:rgba(124,58,237,.18);color:#a78bfa}
.src-bd.cx{background:rgba(13,148,136,.18);color:#2dd4bf}
.src-files{border-top:1px solid rgba(255,255,255,.05)}
.src-file{display:flex;align-items:center;gap:5px;padding:5px 10px 5px 18px}
.src-fn{flex:1;font-size:11px;color:rgba(255,255,255,.5);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.src-fr{font-size:10px;color:rgba(255,255,255,.26)}
.sdl{background:none;border:none;cursor:pointer;color:rgba(255,60,60,.35);font-size:14px;line-height:1;padding:0}
.sdl:hover{color:#f87171}
.cfg-btn{font-size:10px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.12);
  color:rgba(255,255,255,.5);border-radius:4px;padding:2px 6px;cursor:pointer;margin-left:2px}
.cfg-btn:hover{background:rgba(255,255,255,.12);color:#fff}
.add-btn{width:100%;padding:6px;background:rgba(255,255,255,.04);
  border:1px dashed rgba(255,255,255,.1);border-radius:5px;color:rgba(255,255,255,.36);
  font-size:11px;cursor:pointer;transition:all .15s;margin:3px 0 7px}
.add-btn.ml:hover{border-color:var(--ml);color:var(--ml);background:rgba(245,158,11,.07)}
.add-btn.tn:hover{border-color:#a78bfa;color:#a78bfa;background:rgba(124,58,237,.07)}
.add-btn.cx:hover{border-color:#2dd4bf;color:#2dd4bf;background:rgba(13,148,136,.07)}
/* Filtros */
.filt-sep{height:1px;background:rgba(255,255,255,.06);margin:10px 0}
.fg{margin-bottom:7px}
.fl{font-size:10px;color:rgba(255,255,255,.36);margin-bottom:3px;display:block;font-weight:500}
select,input[type=text],input[type=date]{width:100%;background:rgba(255,255,255,.06);
  border:1px solid rgba(255,255,255,.1);border-radius:5px;padding:6px 8px;
  color:#fff;font-size:11px;outline:none;transition:border-color .15s}
select:focus,input:focus{border-color:var(--acc)}
select option{background:#1a2535}
.dr{display:flex;gap:4px}.dr input{flex:1}
.bclr{width:100%;margin-top:5px;padding:6px;border-radius:5px;
  border:1px solid rgba(255,255,255,.1);background:transparent;
  color:rgba(255,255,255,.36);cursor:pointer;font-size:11px;transition:all .15s}
.bclr:hover{background:rgba(255,255,255,.06);color:rgba(255,255,255,.7)}
/* Panel de visualización custom */
.viz-panel{background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.1);
  border-radius:10px;padding:12px;margin-top:2px}
.viz-row{display:flex;gap:4px;margin-bottom:5px}
.viz-row select{flex:1}
.viz-apply{width:100%;padding:9px;border-radius:7px;background:var(--acc);
  border:none;color:#fff;font-size:12px;cursor:pointer;font-weight:700;
  letter-spacing:.4px;transition:all .15s;margin-top:6px}
.viz-apply:hover:not(:disabled){background:var(--acc2);transform:translateY(-1px);box-shadow:0 4px 12px rgba(37,99,235,.35)}
.viz-apply:disabled{opacity:.55;cursor:not-allowed}
.viz-fn-btn{padding:5px 6px;border-radius:5px;border:1px solid rgba(255,255,255,.15);
  background:rgba(255,255,255,.06);color:rgba(255,255,255,.6);cursor:pointer;font-size:11px;
  transition:all .15s;text-align:center}
.viz-fn-btn.act{background:var(--acc);color:#fff;border-color:var(--acc);font-weight:600}
.viz-fn-btn:hover:not(.act){background:rgba(255,255,255,.12);color:rgba(255,255,255,.9)}
/* Paso labels */
.viz-step-lbl{display:flex;align-items:center;gap:5px;font-size:10.5px;font-weight:600;
  color:rgba(255,255,255,.75);text-transform:none;letter-spacing:0;margin-bottom:4px}
.viz-step-num{background:var(--acc);color:#fff;border-radius:50%;width:15px;height:15px;
  font-size:9px;display:inline-flex;align-items:center;justify-content:center;
  font-weight:800;flex-shrink:0;line-height:1}
/* Hints */
.viz-hint{font-size:9.5px;color:rgba(255,255,255,.32);margin-top:2px;
  min-height:12px;line-height:1.3;font-style:italic}
/* Status bar */
.viz-status-bar{font-size:10px;color:rgba(255,255,255,.4);text-align:center;
  margin-top:5px;min-height:14px;transition:color .3s}
/* Filtros dinámicos */
.dyn-filter-col,.dyn-filter-val{width:100%;display:block}
.viz-fn-wrap{transition:opacity .2s}
.chat-chip{padding:4px 10px;border-radius:20px;border:1px solid var(--br);background:#f8fafc;
  color:var(--mu);cursor:pointer;font-size:11px;white-space:nowrap;transition:all .15s}
.chat-chip:hover{background:#eff6ff;border-color:var(--acc);color:var(--acc)}
.msg-user{background:#eff6ff;border-radius:12px 12px 4px 12px;padding:8px 12px;font-size:13px;
  align-self:flex-end;max-width:80%;color:var(--acc);font-weight:500}
.msg-bot{background:#f8fafc;border-radius:12px 12px 12px 4px;padding:10px 14px;font-size:13px;
  align-self:flex-start;max-width:95%;border:1px solid var(--br);line-height:1.55}
.msg-bot-chart{margin-top:10px;height:240px;position:relative}
.msg-thinking{color:var(--mu);font-size:12px;font-style:italic;align-self:flex-start;padding:4px 0}
/* ── Dashi ─────────────────────────────────────────── */
@keyframes dashiBounce{0%,100%{transform:translateY(0)}50%{transform:translateY(-8px)}}
@keyframes dashiPulse{0%,100%{transform:scale(1)}50%{transform:scale(1.2)}}
@keyframes dashiWiggle{0%,100%{transform:rotate(0deg)}25%{transform:rotate(-6deg)}75%{transform:rotate(6deg)}}
@keyframes dashiSlideUp{from{opacity:0;transform:translateY(20px) scale(.96)}to{opacity:1;transform:translateY(0) scale(1)}}
@keyframes dashiTyping{0%,60%,100%{transform:translateY(0)}30%{transform:translateY(-6px)}}
#dashi-btn:hover{transform:scale(1.1) translateY(-4px)!important;filter:drop-shadow(0 8px 20px rgba(0,0,0,.4))!important}
.dashi-chip{padding:5px 11px;border-radius:20px;border:1.5px solid #f3d5d5;background:#fff5f5;
  color:#c0392b;cursor:pointer;font-size:11px;font-weight:500;transition:all .15s;white-space:nowrap}
.dashi-chip:hover{background:#c0392b;color:#fff;border-color:#c0392b}
.dmsg-user{background:linear-gradient(135deg,#c0392b,#e74c3c);color:#fff;border-radius:18px 18px 4px 18px;
  padding:9px 13px;font-size:13px;align-self:flex-end;max-width:82%;line-height:1.45;box-shadow:0 2px 8px rgba(192,57,43,.3)}
.dmsg-bot{background:#fff;border-radius:4px 18px 18px 18px;padding:11px 14px;font-size:13px;
  align-self:flex-start;max-width:90%;border:1.5px solid #f0ece8;line-height:1.55;
  box-shadow:0 2px 8px rgba(0,0,0,.06)}
.dmsg-chart{margin-top:10px;height:220px;position:relative}
.dmsg-thinking{display:flex;gap:4px;align-items:center;padding:10px 14px;background:#fff;
  border-radius:4px 18px 18px 18px;border:1.5px solid #f0ece8;align-self:flex-start}
.dmsg-thinking span{width:7px;height:7px;border-radius:50%;background:#c0392b;display:inline-block}
.dmsg-thinking span:nth-child(1){animation:dashiTyping 1.2s .0s infinite}
.dmsg-thinking span:nth-child(2){animation:dashiTyping 1.2s .2s infinite}
.dmsg-thinking span:nth-child(3){animation:dashiTyping 1.2s .4s infinite}
/* Welcome */
#wlc{display:flex;align-items:center;justify-content:center;flex:1;
  gap:18px;padding:28px;flex-direction:column}
.wlc-inner{display:flex;gap:16px;max-width:920px;width:100%;flex-wrap:wrap;justify-content:center}
.drop-card{flex:1;min-width:240px;max-width:290px;border:2px dashed var(--br);border-radius:14px;
  padding:30px 24px;display:flex;flex-direction:column;align-items:center;gap:11px;
  cursor:pointer;transition:all .2s;background:#fff;text-align:center}
.drop-card.ml:hover,.drop-card.ml.dg{border-color:var(--ml);background:#fffbeb}
.drop-card.tn:hover,.drop-card.tn.dg{border-color:var(--tn);background:#f5f3ff}
.drop-card.cx:hover,.drop-card.cx.dg{border-color:var(--cx);background:#f0fdfa}
.dc-ico{width:54px;height:54px;border-radius:13px;display:flex;align-items:center;justify-content:center}
.dc-ico.ml{background:#fef3c7}.dc-ico.tn{background:#ede9fe}.dc-ico.cx{background:#ccfbf1}
.dc-t{font-size:15px;font-weight:700}
.dc-s{color:var(--mu);font-size:12px;line-height:1.5}
.dc-btn{padding:8px 20px;border-radius:7px;font-size:12px;font-weight:600;cursor:pointer;border:none;margin-top:2px}
.dc-btn.ml{background:var(--ml);color:#1a1a1a}.dc-btn.ml:hover{background:#d97706}
.dc-btn.tn{background:var(--tn);color:#fff}.dc-btn.tn:hover{background:#6d28d9}
.dc-btn.cx{background:var(--cx);color:#fff}.dc-btn.cx:hover{background:#0f766e}
.fmt{display:flex;gap:5px;flex-wrap:wrap;justify-content:center}
.bg{padding:2px 8px;border-radius:20px;font-size:10px;font-weight:500}
.bg.ml{background:#fef3c7;color:#92400e}.bg.tn{background:#ede9fe;color:#5b21b6}
.bg.cx{background:#ccfbf1;color:#0f766e}
/* KPIs */
#kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(172px,1fr));gap:12px;margin-bottom:16px}
.kpi{background:#fff;border-radius:14px;padding:38px 16px 15px 18px;
  border:1px solid var(--br);position:relative;overflow:hidden;
  box-shadow:var(--shadow);transition:box-shadow .15s,transform .15s}
.kpi:hover{box-shadow:var(--shadow-hover);transform:translateY(-2px)}
.kico{position:absolute;top:14px;right:14px;width:28px;height:28px;border-radius:8px;
  display:flex;align-items:center;justify-content:center;flex-shrink:0}
.kico svg{width:15px;height:15px}
.kl{position:absolute;top:16px;left:18px;right:52px}
.kpi.bl{background:linear-gradient(160deg,#eef2fc 0%,#fff 55%)}
.kpi.bl .kico{background:rgba(43,95,217,.13);color:var(--acc)}
.kpi.gr{background:linear-gradient(160deg,#eaf6ef 0%,#fff 55%)}
.kpi.gr .kico{background:rgba(26,138,83,.13);color:var(--grn)}
.kpi.or{background:linear-gradient(160deg,#faf1e0 0%,#fff 55%)}
.kpi.or .kico{background:rgba(200,134,10,.14);color:var(--orn)}
.kpi.rd{background:linear-gradient(160deg,#faeaea 0%,#fff 55%)}
.kpi.rd .kico{background:rgba(211,58,58,.13);color:var(--red)}
.kpi.pu{background:linear-gradient(160deg,#f0ecfa 0%,#fff 55%)}
.kpi.pu .kico{background:rgba(109,67,201,.14);color:#6d43c9}
.kpi.tl{background:linear-gradient(160deg,#e8f5f3 0%,#fff 55%)}
.kpi.tl .kico{background:rgba(15,143,131,.14);color:#0f8f83}
.kl{font-size:9.5px;font-weight:700;color:var(--mu);letter-spacing:.08em;margin-bottom:7px;text-transform:uppercase}
.kv{font-size:23px;font-weight:700;line-height:1.15;letter-spacing:-.01em;font-variant-numeric:tabular-nums;color:var(--tx);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ks{font-size:10.5px;color:var(--mu);margin-top:5px}
/* Comparativa */
.cmp-bar{display:flex;gap:10px;margin-bottom:12px;overflow-x:auto}
.cmp-card{border-radius:var(--r);padding:13px 16px;border:2px solid;flex:1;min-width:180px}
.cmp-card.ml{background:#fffbeb;border-color:#fde68a}
.cmp-card.tn{background:#f5f3ff;border-color:#ddd6fe}
.cmp-card.cx{background:#f0fdfa;border-color:#99f6e4}
.cmp-card.gen{background:#f8fafc;border-color:#e2e8f0}
.cmp-lbl{font-size:10px;font-weight:700;letter-spacing:.06em;margin-bottom:7px}
.cmp-lbl.ml{color:#92400e}.cmp-lbl.tn{color:#5b21b6}
.cmp-lbl.cx{color:#0f766e}.cmp-lbl.gen{color:var(--mu)}
.cmp-row{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:3px}
.cmp-k{font-size:11px;color:var(--mu)}.cmp-v{font-size:13px;font-weight:600}
/* Cards */
.card{background:#fff;border-radius:var(--r);border:1px solid var(--br);overflow:hidden;box-shadow:var(--shadow);transition:box-shadow .15s}
.card:hover{box-shadow:var(--shadow-hover)}
.ch{display:flex;align-items:center;justify-content:space-between;padding:13px 17px;border-bottom:1px solid var(--br)}
.ct{font-size:13px;font-weight:650;letter-spacing:-.005em}.cs{font-size:11px;color:var(--mu);margin-top:1px}
.cb{padding:12px}.cw{position:relative;height:218px}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:12px}
.c2{grid-column:span 2}
@media(max-width:1100px){.g3{grid-template-columns:1fr 1fr}}
@media(max-width:760px){.g2,.g3{grid-template-columns:1fr}}
/* Tabla */
.tw{overflow:auto;max-height:260px}
table{width:100%;border-collapse:collapse;font-size:12px}
thead th{position:sticky;top:0;background:#f8fafc;padding:7px 11px;text-align:left;
  font-weight:600;color:var(--mu);border-bottom:2px solid var(--br);white-space:nowrap}
tbody td{padding:5px 11px;border-bottom:1px solid #f1f5f9;white-space:nowrap;
  max-width:190px;overflow:hidden;text-overflow:ellipsis}
tbody tr:hover td{background:#f8fafc}
.nr{text-align:right;font-variant-numeric:tabular-nums}
.pb{display:flex;align-items:center;gap:6px}
.pt{flex:1;height:4px;background:#e2e8f0;border-radius:3px;overflow:hidden;min-width:36px}
.pf{height:100%;border-radius:3px;transition:width .3s}
.tag{display:inline-block;padding:2px 7px;border-radius:10px;font-size:10px;font-weight:600;white-space:nowrap}
.tOK{background:#dcfce7;color:#15803d}.tC{background:#fef3c7;color:#92400e}
.tX{background:#fee2e2;color:#b91c1c}.tP{background:#e0e7ff;color:#3730a3}
.tD{background:#f1f5f9;color:#475569}
/* Paginación */
.pg{display:flex;align-items:center;justify-content:space-between;
  padding:8px 15px;border-top:1px solid var(--br);font-size:12px;color:var(--mu)}
.pgb{display:flex;gap:4px}
.pb2{padding:4px 10px;border:1px solid var(--br);border-radius:5px;background:#fff;
  cursor:pointer;font-size:12px;transition:all .15s}
.pb2:hover:not(:disabled){border-color:var(--acc);color:var(--acc)}
.pb2:disabled{opacity:.38;cursor:not-allowed}
/* Periodo */
.ptabs{display:flex;gap:2px;background:#f1f5f9;border-radius:5px;padding:2px}
.ptab{padding:3px 9px;border-radius:4px;cursor:pointer;font-size:11px;font-weight:500;
  color:var(--mu);border:none;background:none;transition:all .15s}
.ptab.act{background:#fff;color:var(--acc);box-shadow:0 1px 3px rgba(0,0,0,.09)}
/* Loading / Error */
#ld{display:none;position:fixed;inset:0;background:rgba(10,17,30,.72);
  align-items:center;justify-content:center;z-index:500;flex-direction:column;gap:12px}
#ld.show{display:flex}
.sp{width:36px;height:36px;border:3px solid rgba(255,255,255,.18);border-top-color:#fff;
  border-radius:50%;animation:spin .8s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.lm{color:#fff;font-size:13px}
#eb{display:none;padding:9px 16px;background:#fef2f2;border-bottom:1px solid #fecaca;
  align-items:center;gap:10px;font-size:13px;color:#b91c1c;flex-shrink:0}
#eb.show{display:flex}
.ex{background:none;border:none;cursor:pointer;color:#b91c1c;font-size:16px;margin-left:auto}
/* Modal configurador de columnas */
.modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:400;
  align-items:center;justify-content:center;padding:20px}
.modal-bg.open{display:flex}
.modal{background:#fff;border-radius:14px;max-width:640px;width:100%;
  max-height:90vh;display:flex;flex-direction:column;overflow:hidden}
.modal-head{display:flex;align-items:center;justify-content:space-between;
  padding:16px 20px;border-bottom:1px solid var(--br)}
.modal-title{font-size:15px;font-weight:700}
.modal-body{overflow-y:auto;padding:20px;flex:1}
.modal-foot{padding:14px 20px;border-top:1px solid var(--br);display:flex;gap:8px;justify-content:flex-end}
.role-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.role-item{border:1px solid var(--br);border-radius:8px;padding:10px 12px}
.role-lbl{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.05em;margin-bottom:5px;display:flex;align-items:center;gap:5px}
.role-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.role-desc{font-size:10px;color:var(--mu);margin-bottom:6px}
.role-sel{width:100%;background:#f8fafc;border:1px solid var(--br);border-radius:5px;
  padding:5px 8px;font-size:11px;color:var(--tx);outline:none}
.role-sel:focus{border-color:var(--acc)}
.preview-table{width:100%;font-size:11px;border-collapse:collapse;margin-top:10px}
.preview-table th{background:#f8fafc;padding:5px 8px;text-align:left;font-weight:600;
  color:var(--mu);border-bottom:1px solid var(--br)}
.preview-table td{padding:4px 8px;border-bottom:1px solid #f1f5f9}
.btn{padding:8px 18px;border-radius:7px;font-size:13px;font-weight:600;cursor:pointer;border:none;transition:all .15s}
.btn-primary{background:var(--acc);color:#fff}.btn-primary:hover{background:var(--acc2)}
.btn-ghost{background:#f1f5f9;color:var(--mu)}.btn-ghost:hover{background:#e2e8f0}
.info-tip{background:#eff6ff;border:1px solid #bfdbfe;border-radius:6px;padding:8px 12px;font-size:11px;color:#1d4ed8;margin-bottom:14px}
/* Platform tabs */
.ptf{display:flex;gap:6px;margin-bottom:14px;flex-wrap:wrap}
.ptf-btn{display:flex;align-items:center;gap:7px;padding:9px 18px;border-radius:9px;
  border:2px solid transparent;cursor:pointer;font-size:13px;font-weight:600;
  transition:all .18s;background:#fff;color:var(--mu)}
.ptf-btn:hover{transform:translateY(-1px);box-shadow:0 3px 10px rgba(0,0,0,.1)}
.ptf-btn.all{border-color:var(--br)}.ptf-btn.all.act{border-color:var(--acc);background:#eff6ff;color:var(--acc)}
.ptf-btn.ml-t{border-color:#fde68a}.ptf-btn.ml-t.act{border-color:var(--ml);background:#fffbeb;color:#92400e}
.ptf-btn.tn-t{border-color:#ddd6fe}.ptf-btn.tn-t.act{border-color:var(--tn);background:#f5f3ff;color:#5b21b6}
.ptf-btn.cx-t{border-color:#99f6e4}.ptf-btn.cx-t.act{border-color:var(--cx);background:#f0fdfa;color:#0f766e}
.ptf-dot{width:9px;height:9px;border-radius:50%;flex-shrink:0}
.ptf-cnt{font-size:10px;font-weight:700;padding:1px 7px;border-radius:10px;margin-left:2px}
.ptf-cnt.ml{background:#fef3c7;color:#92400e}
.ptf-cnt.tn{background:#ede9fe;color:#5b21b6}
.ptf-cnt.cx{background:#ccfbf1;color:#0f766e}
.ptf-cnt.all{background:#eff6ff;color:var(--acc)}
/* Dash header */
.dh{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:8px}
.dt{font-size:17px;font-weight:700}
.nb{background:#eff6ff;color:var(--acc);padding:3px 11px;border-radius:16px;font-size:12px;font-weight:600}
.dash-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.pdf-btn{display:inline-flex;align-items:center;gap:7px;padding:7px 12px;border-radius:8px;
  border:1px solid #bfdbfe;background:#fff;color:var(--acc);font-size:12px;font-weight:700;
  cursor:pointer;transition:all .15s;box-shadow:0 1px 2px rgba(15,23,42,.04)}
.pdf-btn:hover{border-color:var(--acc);background:#eff6ff;transform:translateY(-1px)}
.pdf-btn:disabled{opacity:.45;cursor:not-allowed;transform:none}
input[type=file]{display:none}
</style>
</head>
<body>
<div id="ld"><div class="sp"></div><div class="lm" id="lm">Procesando...</div></div>
<div id="eb"><span id="et"></span><button class="ex" onclick="hE()">✕</button></div>

<div id="top">
  <div class="logo">
    <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>
    Dashify
    <div class="chips">
      <span class="chip chip-ml">ML</span>
      <span class="chip chip-tn">TN</span>
      <span class="chip chip-cx">+ Excel</span>
    </div>
  </div>
  <div style="flex:1"></div>
  <button class="tbtn ml" onclick="trig('ml')">
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
    Cargar reporte ML
  </button>
  <button class="tbtn tn" onclick="trig('tn')">
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
    Cargar reporte TN
  </button>
  <button class="tbtn cx" onclick="trig('cx')">
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
    Cualquier Excel/CSV
  </button>
  <button class="tbtn" id="tbtn-ventas" onclick="showVentasModule()" style="background:#1a8a53;color:#fff">
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M3 3v18h18"/><rect x="7" y="11" width="3" height="6" rx="1"/><rect x="12" y="7" width="3" height="10" rx="1"/><rect x="17" y="4" width="3" height="13" rx="1"/></svg>
    Ventas
  </button>
  <button class="tbtn" id="tbtn-pubs" onclick="showPubsModule()" style="background:#6d43c9;color:#fff">
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="3" y="3" width="18" height="4" rx="1"/><rect x="3" y="10" width="18" height="4" rx="1"/><rect x="3" y="17" width="18" height="4" rx="1"/></svg>
    Publicaciones
  </button>
  <button class="tbtn" id="tbtn-pub" onclick="showPubModule()" style="background:#7a3f9e;color:#fff">
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="2" y="3" width="20" height="14" rx="2"/><path d="M8 21h8M12 17v4"/><path d="M7 8h10M7 12h6"/></svg>
    Publicidad ML
  </button>
  <button class="tbtn" id="tbtn-mkt" onclick="showMktModule()" style="background:#2f7fb8;color:#fff">
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="22 7 13.5 15.5 8.5 10.5 2 17"/><polyline points="16 7 22 7 22 13"/></svg>
    Tendencias
  </button>
  <button class="tbtn" id="tbtn-cotizador" onclick="showCotizadorModule()" style="background:#7a2e2e;color:#fff">
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="10"/><path d="M12 6v12M9 9a3 3 0 0 1 3-1.5c1.7 0 3 1 3 2.5s-1.3 2.5-3 2.5-3 1-3 2.5 1.3 2.5 3 2.5a3 3 0 0 0 3-1.5"/></svg>
    Cotizador
  </button>

</div>

<input type="file" id="fi-ml" accept=".xlsx,.xls" onchange="onFI(event,'ml')">
<input type="file" id="fi-tn" accept=".xlsx,.xls,.csv,.CSV" onchange="onFI(event,'tn')">
<input type="file" id="fi-cx" accept=".xlsx,.xls,.csv,.ods" onchange="onFI(event,'cx')">
<input type="file" id="fi-fichas" accept=".xlsx,.xls" onchange="onFI(event,'fichas')">
<input type="file" id="fi-pub-camp" accept=".xlsx,.xls" onchange="onPubFI(event,'campanias')">
<input type="file" id="fi-pub-an"   accept=".xlsx,.xls" onchange="onPubFI(event,'anuncios')">
<input type="file" id="fi-mkt" accept=".xlsx,.xls" multiple onchange="onMktFI(event)">

<div id="main">
<div id="sb">
<div class="sb-inner">

  <div class="src-area">
    <div class="src-head"><div class="src-dot ml"></div><span class="src-nm">Mercado Libre</span><span class="src-bd ml">ML</span></div>
    <div class="src-files" id="files-ml"><div style="padding:6px 10px 6px 18px;font-size:11px;color:rgba(255,255,255,.22)">Sin archivos</div></div>
    <button class="add-btn ml" onclick="trig('ml')">+ Cargar reporte ML</button>
    <button class="add-btn" id="btn-fichas" onclick="trig('fichas')" style="margin-top:4px;border-color:rgba(99,102,241,.5);color:rgba(200,200,255,.7);font-size:10px" title="Asigna categorías a tus publicaciones ML">📋 Cargar fichas técnicas ML</button>
    <div id="fichas-status" style="display:none;padding:4px 10px 2px 10px;font-size:10px;color:rgba(150,255,150,.75)"></div>
  </div>

  <div class="src-area">
    <div class="src-head"><div class="src-dot tn"></div><span class="src-nm">Tienda Nube</span><span class="src-bd tn">TN</span></div>
    <div class="src-files" id="files-tn"><div style="padding:6px 10px 6px 18px;font-size:11px;color:rgba(255,255,255,.22)">Sin archivos</div></div>
    <button class="add-btn tn" onclick="trig('tn')">+ Cargar reporte TN</button>
  </div>

  <div class="src-area">
    <div class="src-head"><div class="src-dot cx"></div><span class="src-nm">Fuente personalizada</span><span class="src-bd cx">XLS</span></div>
    <div class="src-files" id="files-cx"><div style="padding:6px 10px 6px 18px;font-size:11px;color:rgba(255,255,255,.22)">Sin archivos</div></div>
    <button class="add-btn cx" onclick="trig('cx')">+ Cualquier Excel/CSV</button>
  </div>

  <div class="filt-sep"></div>
  <div class="ss">FILTROS</div>

  <div class="fg">
    <label class="fl">Canal / Fuente</label>
    <select id="ff" onchange="rf()">
      <option value="__all__">Todos los canales</option>
    </select>
  </div>
  <div class="fg">
    <label class="fl">Estado</label>
    <select id="fe" onchange="rf()"><option value="__all__">Todos</option></select>
  </div>
  <div class="fg">
    <label class="fl">Categoría / Tipo</label>
    <select id="fcat" onchange="rf()"><option value="__all__">Todas</option></select>
  </div>
  <div class="fg">
    <label class="fl">Provincia / Región</label>
    <select id="fprov" onchange="rf()"><option value="__all__">Todas</option></select>
  </div>

  <!-- Filtros dinámicos múltiples -->
  <div id="dynamic-filters-container"></div>
  <div id="fg-add-filter" style="display:none;margin:4px 0 6px 0">
    <button onclick="addDynamicFilter()" style="
      width:100%;background:rgba(255,255,255,.06);border:1px dashed rgba(255,255,255,.18);
      color:rgba(255,255,255,.55);border-radius:6px;padding:5px 0;font-size:11px;
      cursor:pointer;transition:all .15s
    " onmouseover="this.style.background='rgba(255,255,255,.1)'" onmouseout="this.style.background='rgba(255,255,255,.06)'">
      + Agregar filtro por columna
    </button>
  </div>

  <div class="fg">
    <label class="fl">Mes</label>
    <select id="fmes" onchange="onMesChange()"><option value="__all__">Todos los meses</option></select>
  </div>
  <div class="fg">
    <label class="fl">Período</label>
    <div class="dr"><input type="date" id="fd" onchange="onDateChange()"><input type="date" id="fh" onchange="onDateChange()"></div>
  </div>
  <div class="fg">
    <label class="fl">Buscar en todo</label>
    <input type="text" id="fq" placeholder="Texto libre..." onkeydown="if(event.key==='Enter')rf()">
  </div>
  <button class="bclr" onclick="clrF()">↺ Limpiar filtros</button>

  <div class="filt-sep"></div>
  <div class="ss">VISUALIZACIÓN EXTRA</div>
  <div class="viz-panel">

    <!-- Paso 1: Agrupar por — solo columnas categóricas -->
    <div class="fg">
      <label class="fl viz-step-lbl">
        <span class="viz-step-num">1</span> ¿Qué querés ver en el eje X?
      </label>
      <select id="v-by" onchange="onVizByChange()">
        <option value="">— elegí cómo agrupar —</option>
      </select>
      <div id="v-by-hint" class="viz-hint"></div>
    </div>

    <!-- Paso 2: Qué medir — solo columnas numéricas + conteo -->
    <div class="fg" id="viz-val-group">
      <label class="fl viz-step-lbl">
        <span class="viz-step-num">2</span> ¿Qué querés medir?
      </label>
      <select id="v-val" onchange="onVizMetricChange()">
        <option value="__count__">📊 Cantidad de registros</option>
      </select>
      <div id="v-val-hint" class="viz-hint"></div>
    </div>

    <!-- Paso 3: Función — solo visible cuando la métrica es numérica -->
    <div class="fg viz-fn-wrap" id="viz-fn-group">
      <label class="fl viz-step-lbl">
        <span class="viz-step-num">3</span> ¿Cómo calcularlo?
      </label>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px">
        <button class="viz-fn-btn act" data-fn="sum" onclick="setVizFn('sum',this)" title="Suma todos los valores">∑ Suma total</button>
        <button class="viz-fn-btn" data-fn="avg" onclick="setVizFn('avg',this)" title="Valor promedio">⌀ Promedio</button>
        <button class="viz-fn-btn" data-fn="max" onclick="setVizFn('max',this)" title="Valor máximo">↑ Máximo</button>
        <button class="viz-fn-btn" data-fn="min" onclick="setVizFn('min',this)" title="Valor mínimo">↓ Mínimo</button>
      </div>
      <input type="hidden" id="v-agg" value="sum">
    </div>

    <!-- Paso 4: Tipo de gráfico -->
    <div class="fg">
      <label class="fl viz-step-lbl">
        <span class="viz-step-num" id="viz-step-chart-num">4</span> Tipo de gráfico
      </label>
      <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:3px">
        <button class="viz-fn-btn act" id="vt-bar" onclick="setVizType('bar',this)">▊ Barras</button>
        <button class="viz-fn-btn" id="vt-line" onclick="setVizType('line',this)">↗ Línea</button>
        <button class="viz-fn-btn" id="vt-doughnut" onclick="setVizType('doughnut',this)">◉ Torta</button>
      </div>
      <input type="hidden" id="v-type" value="bar">
    </div>

    <!-- Top N -->
    <div class="fg">
      <label class="fl viz-step-lbl" style="font-size:10px;opacity:.7">Mostrar top resultados</label>
      <div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:3px">
        <button class="viz-fn-btn" data-top="5" onclick="setVizTop(5,this)">5</button>
        <button class="viz-fn-btn act" data-top="10" onclick="setVizTop(10,this)">10</button>
        <button class="viz-fn-btn" data-top="20" onclick="setVizTop(20,this)">20</button>
        <button class="viz-fn-btn" data-top="50" onclick="setVizTop(50,this)">50</button>
      </div>
      <input type="hidden" id="v-top" value="10">
    </div>

    <button class="viz-apply" onclick="applyViz()" id="viz-apply-btn">▶ Graficar</button>
    <div id="viz-status" class="viz-status-bar"></div>
  </div>
</div>
</div>


<div id="ct">
<!-- ═══════════════════════════════════════════
     DASHI — Panda Rojo Analista IA (Flotante)
══════════════════════════════════════════════ -->

<!-- Botón flotante Dashi -->
<div id="dashi-btn" onclick="toggleDashi()" title="¡Preguntale a Dashi!" style="
  position:fixed;bottom:28px;right:28px;width:68px;height:68px;
  border-radius:50%;cursor:pointer;z-index:300;
  filter:drop-shadow(0 4px 16px rgba(0,0,0,.35));
  transition:transform .2s,filter .2s;
  animation:dashiBounce 3s ease-in-out infinite;
  user-select:none;
">
  <svg width="68" height="68" viewBox="0 0 68 68" xmlns="http://www.w3.org/2000/svg">
    <!-- Cuerpo -->
    <circle cx="34" cy="38" r="22" fill="#e8d5b7"/>
    <!-- Parche panza -->
    <ellipse cx="34" cy="42" rx="13" ry="10" fill="#f5ede0"/>
    <!-- Orejas izquierda -->
    <circle cx="14" cy="18" r="9" fill="#1a1a1a"/>
    <circle cx="14" cy="18" r="5.5" fill="#c0392b"/>
    <!-- Orejas derecha -->
    <circle cx="54" cy="18" r="9" fill="#1a1a1a"/>
    <circle cx="54" cy="18" r="5.5" fill="#c0392b"/>
    <!-- Cabeza -->
    <circle cx="34" cy="26" r="20" fill="#c0392b"/>
    <!-- Cara blanca -->
    <ellipse cx="34" cy="28" rx="14" ry="12" fill="#f5ede0"/>
    <!-- Antifaz izquierdo -->
    <ellipse cx="26" cy="24" rx="7" ry="5.5" fill="#1a1a1a" opacity=".85"/>
    <!-- Antifaz derecho -->
    <ellipse cx="42" cy="24" rx="7" ry="5.5" fill="#1a1a1a" opacity=".85"/>
    <!-- Ojos -->
    <circle cx="26" cy="24" r="3" fill="#fff"/>
    <circle cx="42" cy="24" r="3" fill="#fff"/>
    <circle cx="27" cy="24" r="1.5" fill="#1a1a1a"/>
    <circle cx="43" cy="24" r="1.5" fill="#1a1a1a"/>
    <!-- Brillo ojos -->
    <circle cx="27.8" cy="23" r=".7" fill="#fff"/>
    <circle cx="43.8" cy="23" r=".7" fill="#fff"/>
    <!-- Nariz -->
    <ellipse cx="34" cy="30" rx="2.5" ry="1.8" fill="#c0392b"/>
    <!-- Bigotes -->
    <line x1="20" y1="31" x2="30" y2="30.5" stroke="#888" stroke-width=".8"/>
    <line x1="20" y1="33" x2="30" y2="32" stroke="#888" stroke-width=".8"/>
    <line x1="38" y1="30.5" x2="48" y2="31" stroke="#888" stroke-width=".8"/>
    <line x1="38" y1="32" x2="48" y2="33" stroke="#888" stroke-width=".8"/>
    <!-- Cola roja rayada (parte visible) -->
    <path d="M52 54 Q62 46 58 36 Q56 30 60 24" stroke="#c0392b" stroke-width="5" fill="none" stroke-linecap="round"/>
    <path d="M52 54 Q62 46 58 36 Q56 30 60 24" stroke="#e8d5b7" stroke-width="5" fill="none" stroke-linecap="round" stroke-dasharray="4 5"/>
  </svg>
  <!-- Burbuja de notificación -->
  <div id="dashi-notif" style="
    position:absolute;top:-2px;right:-2px;
    background:#e53e3e;color:#fff;border-radius:50%;
    width:20px;height:20px;font-size:11px;font-weight:700;
    display:flex;align-items:center;justify-content:center;
    border:2px solid #fff;animation:dashiPulse 2s ease-in-out infinite;
  ">!</div>
</div>

<!-- Panel Dashi -->
<div id="dashi-panel" style="
  position:fixed;bottom:108px;right:28px;width:380px;max-width:94vw;
  background:#fff;border-radius:20px;
  box-shadow:0 8px 40px rgba(0,0,0,.22);
  z-index:299;display:none;flex-direction:column;overflow:hidden;
  border:1.5px solid #f3e8e8;
  animation:dashiSlideUp .25s cubic-bezier(.34,1.56,.64,1);
  max-height:82vh;
">
  <!-- Header con Dashi -->
  <div style="background:linear-gradient(135deg,#c0392b 0%,#e74c3c 100%);padding:14px 16px;display:flex;align-items:center;gap:12px;flex-shrink:0">
    <!-- Mini Dashi -->
    <div style="width:44px;height:44px;flex-shrink:0;animation:dashiWiggle 2s ease-in-out infinite">
      <svg width="44" height="44" viewBox="0 0 68 68" xmlns="http://www.w3.org/2000/svg">
        <circle cx="34" cy="38" r="22" fill="#e8d5b7"/>
        <ellipse cx="34" cy="42" rx="13" ry="10" fill="#f5ede0"/>
        <circle cx="14" cy="18" r="9" fill="#1a1a1a"/><circle cx="14" cy="18" r="5.5" fill="#c0392b"/>
        <circle cx="54" cy="18" r="9" fill="#1a1a1a"/><circle cx="54" cy="18" r="5.5" fill="#c0392b"/>
        <circle cx="34" cy="26" r="20" fill="#c0392b"/>
        <ellipse cx="34" cy="28" rx="14" ry="12" fill="#f5ede0"/>
        <ellipse cx="26" cy="24" rx="7" ry="5.5" fill="#1a1a1a" opacity=".85"/>
        <ellipse cx="42" cy="24" rx="7" ry="5.5" fill="#1a1a1a" opacity=".85"/>
        <circle cx="26" cy="24" r="3" fill="#fff"/><circle cx="42" cy="24" r="3" fill="#fff"/>
        <circle cx="27" cy="24" r="1.5" fill="#1a1a1a"/><circle cx="43" cy="24" r="1.5" fill="#1a1a1a"/>
        <circle cx="27.8" cy="23" r=".7" fill="#fff"/><circle cx="43.8" cy="23" r=".7" fill="#fff"/>
        <ellipse cx="34" cy="30" rx="2.5" ry="1.8" fill="#c0392b"/>
      </svg>
    </div>
    <div style="flex:1;color:#fff">
      <div style="font-weight:800;font-size:15px;letter-spacing:.3px">Dashi 🐾</div>
      <div style="font-size:11px;opacity:.85" id="dashi-status-txt">Tu analista de ventas</div>
    </div>
    <div style="display:flex;gap:6px">
      <button onclick="clearChat()" title="Limpiar chat" style="background:rgba(255,255,255,.18);border:none;border-radius:8px;color:#fff;cursor:pointer;padding:5px 8px;font-size:12px">🗑</button>
      <button onclick="toggleDashi()" style="background:rgba(255,255,255,.18);border:none;border-radius:8px;color:#fff;cursor:pointer;padding:5px 8px;font-size:14px">✕</button>
    </div>
  </div>

  <!-- Mensajes -->
  <div id="chat-msgs" style="flex:1;overflow-y:auto;padding:14px 14px 8px;display:flex;flex-direction:column;gap:10px;min-height:160px;max-height:420px;background:#fafaf9"></div>

  <!-- Chips -->
  <div id="chat-chips" style="padding:8px 12px 4px;display:flex;gap:5px;flex-wrap:wrap;background:#fafaf9;border-top:1px solid #f0ece8;flex-shrink:0">
    <button class="dashi-chip" onclick="askChip(this)">📈 ¿Vendí más que el mes pasado?</button>
    <button class="dashi-chip" onclick="askChip(this)">🏆 Producto más rentable</button>
    <button class="dashi-chip" onclick="askChip(this)">📍 ¿Qué provincia compra más?</button>
    <button class="dashi-chip" onclick="askChip(this)">🚚 Analizá mis envíos</button>
    <button class="dashi-chip" onclick="askChip(this)">📊 Evolución de ingresos</button>
  </div>

  <!-- Input -->
  <div style="padding:10px 12px 12px;background:#fff;border-top:1px solid #f0ece8;display:flex;gap:8px;flex-shrink:0;align-items:flex-end">
    <input id="chat-input" type="text" placeholder="Preguntale algo a Dashi..." style="
      flex:1;padding:9px 13px;border-radius:20px;border:1.5px solid #e8d5d5;
      font-size:13px;outline:none;transition:border-color .2s;background:#fafaf9;
    " onkeydown="if(event.key==='Enter')sendChat()" onfocus="this.style.borderColor='#c0392b'" onblur="this.style.borderColor='#e8d5d5'">
    <button onclick="sendChat()" id="chat-send" style="
      width:38px;height:38px;border-radius:50%;background:#c0392b;color:#fff;
      border:none;cursor:pointer;font-size:16px;display:flex;align-items:center;
      justify-content:center;flex-shrink:0;transition:transform .15s,background .15s;
    " onmouseover="this.style.background='#a93226';this.style.transform='scale(1.08)'" onmouseout="this.style.background='#c0392b';this.style.transform='scale(1)'">➤</button>
  </div>
</div>


  <!-- Welcome -->
  <div id="wlc" style="display:flex">
    <div style="text-align:center;margin-bottom:12px">
      <div style="font-size:20px;font-weight:700">Dashboard Multi-Marketplace</div>
      <div style="color:var(--mu);font-size:13px;margin-top:4px">Cargá uno o varios archivos de cualquier fuente</div>
    </div>
    <div class="wlc-inner">
      <div class="drop-card ml" id="dz-ml" ondragover="dov(event,'ml')" ondragleave="dlv('ml')" ondrop="drp(event,'ml')">
        <div class="dc-ico ml"><svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="1.8"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg></div>
        <div class="dc-t">Mercado Libre</div>
        <div class="dc-s">Reporte desde Mis Ventas → Exportar</div>
        <div class="fmt"><span class="bg ml">.xlsx</span><span class="bg ml">.xls</span></div>
        <button class="dc-btn ml" onclick="trig('ml')">Seleccionar</button>
      </div>
      <div class="drop-card tn" id="dz-tn" ondragover="dov(event,'tn')" ondragleave="dlv('tn')" ondrop="drp(event,'tn')">
        <div class="dc-ico tn"><svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#7c3aed" stroke-width="1.8"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg></div>
        <div class="dc-t">Tienda Nube</div>
        <div class="dc-s">Reporte Panel TN → Informes, o cualquier CSV/Excel etiquetado como Tienda Nube</div>
        <div class="fmt"><span class="bg tn">.xlsx</span><span class="bg tn">.csv</span></div>
        <button class="dc-btn tn" onclick="trig('tn')">Seleccionar</button>
      </div>
      <div class="drop-card cx" id="dz-cx" ondragover="dov(event,'cx')" ondragleave="dlv('cx')" ondrop="drp(event,'cx')">
        <div class="dc-ico cx"><svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#0d9488" stroke-width="1.8"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg></div>
        <div class="dc-t">Cualquier Excel / CSV</div>
        <div class="dc-s">Detecta columnas automáticamente.<br>Podés mapearlas a medida.</div>
        <div class="fmt"><span class="bg cx">.xlsx</span><span class="bg cx">.csv</span><span class="bg cx">.ods</span></div>
        <button class="dc-btn cx" onclick="trig('cx')">Seleccionar</button>
      </div>
      <div class="drop-card" id="dz-fichas" ondragover="dov(event,'fichas')" ondragleave="dlv('fichas')" ondrop="drp(event,'fichas')" style="border-color:rgba(99,102,241,.4);min-width:220px;max-width:260px">
        <div class="dc-ico" style="background:rgba(99,102,241,.12)"><svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#6366f1" stroke-width="1.8"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/><line x1="13" y1="13" x2="17" y2="13"/><line x1="13" y1="17" x2="17" y2="17"/></svg></div>
        <div class="dc-t" style="color:#a5b4fc">Fichas Técnicas ML</div>
        <div class="dc-s">Asigna categorías reales a tus publicaciones ML por ID, SKU o título.</div>
        <div class="fmt"><span class="bg" style="background:rgba(99,102,241,.15);color:#a5b4fc">.xlsx</span></div>
        <button class="dc-btn" onclick="trig('fichas')" style="background:#6366f1;color:#fff">Cargar</button>
      </div>
    </div>
    <div style="font-size:11px;color:var(--mu);text-align:center">También podés arrastrar archivos directamente sobre las tarjetas</div>
  </div>

  <!-- ═══════════════════════════════════════════
       MÓDULO PUBLICIDAD ML ADS
  ══════════════════════════════════════════════ -->
  <style>
  /* ── Publicidad Module Styles ─────────────────────────────── */
  #pub-module{--pub-bg:#0f172a;--pub-card:#1e293b;--pub-card2:#263348;
    --pub-acc:#facc15;--pub-grn:#10b981;--pub-red:#f43f5e;--pub-orn:#f59e0b;
    --pub-pur:#a78bfa;--pub-tx:#f1f5f9;--pub-mu:#94a3b8;--pub-br:#334155}
  .pub-wrap{background:#f8fafc;border-radius:14px;overflow:hidden;border:1px solid #e2e8f0}
  .pub-banner{background:var(--pub-bg);padding:20px 24px 18px;position:relative;overflow:hidden}
  .pub-banner::before{content:'';position:absolute;top:-40px;right:-60px;width:220px;height:220px;
    border-radius:50%;background:rgba(167,139,250,.08);pointer-events:none}
  .pub-banner::after{content:'';position:absolute;bottom:-30px;right:80px;width:120px;height:120px;
    border-radius:50%;background:rgba(250,204,21,.05);pointer-events:none}
  .pub-banner-top{display:flex;align-items:flex-start;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:16px}
  .pub-banner-tag{display:flex;align-items:center;gap:8px}
  .pub-banner-badge{background:var(--pub-acc);color:#0f172a;font-size:9px;font-weight:900;
    padding:3px 8px;border-radius:4px;letter-spacing:.08em;text-transform:uppercase}
  .pub-banner-title{color:#fff;font-size:17px;font-weight:800;letter-spacing:-.3px}
  .pub-banner-sub{color:var(--pub-mu);font-size:11px;margin-top:2px}
  .pub-banner-camps{display:flex;gap:6px;flex-wrap:wrap}
  .pub-camp-pill{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.12);
    color:rgba(255,255,255,.7);font-size:10px;font-weight:600;padding:3px 10px;border-radius:20px}
  .pub-hero-kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-top:4px}
  .pub-hkpi{background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.09);
    border-radius:10px;padding:12px 16px;position:relative;overflow:hidden}
  .pub-hkpi::before{content:'';position:absolute;top:0;left:0;right:0;height:2px}
  .pub-hkpi.grn::before{background:var(--pub-grn)}.pub-hkpi.red::before{background:var(--pub-red)}
  .pub-hkpi.yel::before{background:var(--pub-acc)}.pub-hkpi.pur::before{background:var(--pub-pur)}
  .pub-hkpi.orn::before{background:var(--pub-orn)}
  .pub-hkpi-lbl{font-size:9px;font-weight:700;color:var(--pub-mu);letter-spacing:.07em;text-transform:uppercase;margin-bottom:5px}
  .pub-hkpi-val{font-size:22px;font-weight:800;color:#fff;line-height:1;font-variant-numeric:tabular-nums}
  .pub-hkpi-sub{font-size:10px;color:var(--pub-mu);margin-top:4px}
  .pub-hkpi-badge{display:inline-block;font-size:9px;font-weight:700;padding:1px 6px;border-radius:10px;margin-top:4px}
  .pub-hkpi-badge.ok{background:rgba(16,185,129,.2);color:#6ee7b7}
  .pub-hkpi-badge.warn{background:rgba(245,158,11,.2);color:#fcd34d}
  .pub-hkpi-badge.bad{background:rgba(244,63,94,.2);color:#fda4af}
  /* Mes tabs */
  .pub-mes-tabs-wrap{background:#fff;padding:12px 20px;display:flex;align-items:center;gap:8px;
    border-bottom:1px solid #f1f5f9;flex-wrap:wrap}
  .pub-mes-lbl-txt{font-size:10px;font-weight:700;color:#94a3b8;letter-spacing:.07em;text-transform:uppercase;margin-right:4px;white-space:nowrap}
  .pub-mes-tab{padding:5px 14px;border-radius:20px;border:1.5px solid #e5e7eb;background:#f9fafb;
    color:#6b7280;font-size:11px;font-weight:600;cursor:pointer;transition:all .15s;white-space:nowrap}
  .pub-mes-tab:hover{border-color:#a78bfa;color:#7c3aed;background:#f5f3ff}
  .pub-mes-tab.activo{background:#7e22ce;color:#fff;border-color:#7e22ce}
  .pub-period-lbl{font-size:10px;color:#94a3b8;font-style:italic;margin-left:auto}
  /* Monthly cards strip */
  .pub-meses-strip{display:flex;gap:10px;overflow-x:auto;padding:16px 20px;background:#fff;
    border-bottom:1px solid #f1f5f9;scrollbar-width:thin}
  .pub-mes-card{flex:0 0 180px;border-radius:12px;overflow:hidden;border:2px solid transparent;
    transition:all .2s;cursor:pointer}
  .pub-mes-card:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(0,0,0,.1)}
  .pub-mes-card.sel{border-color:#7e22ce;box-shadow:0 4px 16px rgba(126,34,206,.2)}
  .pub-mes-card-head{padding:10px 12px 8px;background:var(--pub-bg)}
  .pub-mes-card-name{color:#fff;font-size:13px;font-weight:800;text-transform:uppercase;letter-spacing:.05em}
  .pub-mes-card-parcial{font-size:9px;background:var(--pub-orn);color:#fff;padding:1px 5px;border-radius:3px;font-weight:700;margin-left:5px;vertical-align:middle}
  .pub-mes-card-body{background:#f8fafc;padding:10px 12px;display:grid;grid-template-columns:1fr 1fr;gap:5px}
  .pub-mc-row{display:flex;flex-direction:column}
  .pub-mc-lbl{font-size:8px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:.06em}
  .pub-mc-val{font-size:12px;font-weight:700;color:#1e293b;font-variant-numeric:tabular-nums}
  .pub-mc-val.ok{color:#10b981}.pub-mc-val.warn{color:#f59e0b}.pub-mc-val.bad{color:#f43f5e}
  /* Main content */
  .pub-main{padding:16px 20px;background:#f8fafc;display:flex;flex-direction:column;gap:14px}
  /* Campaign cards */
  .pub-camps-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:10px}
  .pub-camp-card{background:#fff;border-radius:12px;border:1px solid #e2e8f0;overflow:hidden;transition:box-shadow .2s}
  .pub-camp-card:hover{box-shadow:0 4px 16px rgba(0,0,0,.08)}
  .pub-camp-head{padding:12px 16px 10px;border-bottom:1px solid #f1f5f9;display:flex;align-items:center;justify-content:space-between}
  .pub-camp-name{font-size:13px;font-weight:800;color:#0f172a;letter-spacing:-.2px}
  .pub-camp-status{font-size:9px;font-weight:700;padding:2px 8px;border-radius:10px}
  .pub-camp-status.activa{background:#dcfce7;color:#15803d}.pub-camp-status.pausada{background:#f1f5f9;color:#64748b}
  .pub-camp-metrics{padding:12px 16px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px}
  .pub-camp-m{display:flex;flex-direction:column;gap:1px}
  .pub-camp-ml{font-size:8.5px;font-weight:600;color:#94a3b8;text-transform:uppercase;letter-spacing:.05em}
  .pub-camp-mv{font-size:14px;font-weight:800;color:#0f172a;font-variant-numeric:tabular-nums}
  .pub-camp-mv.ok{color:#10b981}.pub-camp-mv.warn{color:#f59e0b}.pub-camp-mv.bad{color:#f43f5e}
  .pub-camp-bar{height:3px;border-radius:2px;margin:8px 16px 12px;background:#e2e8f0;overflow:hidden}
  .pub-camp-bar-fill{height:100%;border-radius:2px;transition:width .5s ease}
  /* Charts */
  .pub-charts-row{display:grid;grid-template-columns:1fr 1fr;gap:12px}
  @media(max-width:760px){.pub-charts-row{grid-template-columns:1fr}}
  .pub-chart-card{background:#fff;border-radius:12px;border:1px solid #e2e8f0;overflow:hidden}
  .pub-chart-head{padding:12px 16px 0;border-bottom:none;display:flex;align-items:flex-start;justify-content:space-between}
  .pub-chart-title{font-size:12px;font-weight:700;color:#0f172a}
  .pub-chart-sub{font-size:10px;color:#94a3b8;margin-top:1px}
  .pub-chart-body{padding:4px 12px 12px;height:200px;position:relative}
  /* Top productos */
  .pub-top-table{width:100%;border-collapse:collapse;font-size:11px}
  .pub-top-table thead th{padding:8px 12px;text-align:left;font-size:9px;font-weight:700;
    color:#94a3b8;text-transform:uppercase;letter-spacing:.06em;border-bottom:2px solid #f1f5f9;background:#fafafa}
  .pub-top-table thead th.nr{text-align:right}
  .pub-top-table tbody td{padding:9px 12px;border-bottom:1px solid #f8fafc;color:#1e293b;white-space:nowrap}
  .pub-top-table tbody td.nr{text-align:right;font-variant-numeric:tabular-nums;font-weight:600}
  .pub-top-table tbody tr:hover td{background:#fafafa}
  .pub-rank{width:22px;height:22px;border-radius:50%;display:inline-flex;align-items:center;
    justify-content:center;font-size:9px;font-weight:800;flex-shrink:0}
  .pub-rank.r1{background:#fef9c3;color:#854d0e}.pub-rank.r2{background:#f1f5f9;color:#475569}
  .pub-rank.r3{background:#fef3c7;color:#92400e}.pub-rank.rn{background:#f8fafc;color:#94a3b8}
  .pub-tag{display:inline-block;padding:2px 7px;border-radius:8px;font-size:9px;font-weight:700}
  .pub-tag.ok{background:#dcfce7;color:#15803d}.pub-tag.warn{background:#fef9c3;color:#854d0e}
  .pub-tag.bad{background:#fee2e2;color:#b91c1c}.pub-tag.na{background:#f1f5f9;color:#64748b}
  .pub-acos-bar{display:flex;align-items:center;gap:6px}
  .pub-acos-track{flex:1;height:4px;background:#e2e8f0;border-radius:2px;overflow:hidden;min-width:40px}
  .pub-acos-fill{height:100%;border-radius:2px}
  /* Upload area */
  .pub-up-area{display:flex;gap:10px;padding:16px 20px;background:#fff;border-bottom:1px solid #f1f5f9;flex-wrap:wrap}
  .pub-drop-new{flex:1;min-width:200px;border:2px dashed #e5e7eb;border-radius:10px;padding:16px;
    text-align:center;cursor:pointer;transition:all .2s;background:#fafafa}
  .pub-drop-new:hover{border-color:#7e22ce;background:#faf5ff}
  .pub-drop-new.cargado{border-color:#10b981;background:#f0fdf4;border-style:solid}
  .pub-drop-ico{font-size:24px;margin-bottom:6px}
  .pub-drop-t{font-size:12px;font-weight:700;color:#374151;margin-bottom:3px}
  .pub-drop-s{font-size:10px;color:#9ca3af}
  .pub-drop-st{font-size:10px;color:#10b981;margin-top:6px;font-weight:600;min-height:14px}
  </style>

  <div id="pub-module" style="display:none">

    <!-- ── Upload strip ──────────────────────────────────────── -->
    <div class="pub-wrap">
      <div class="pub-up-area">
        <div class="pub-drop-new" id="pub-dz-camp"
          onclick="trigPub('camp')"
          ondragover="pubDov(event,'camp')" ondrop="pubDrop(event,'campanias')">
          <div class="pub-drop-ico">📊</div>
          <div class="pub-drop-t">Reporte de Campañas</div>
          <div class="pub-drop-s">Publicidad → Campañas → Exportar xlsx</div>
          <div class="pub-drop-st" id="pub-status-camp"></div>
        </div>
        <div class="pub-drop-new" id="pub-dz-an"
          onclick="trigPub('an')"
          ondragover="pubDov(event,'an')" ondrop="pubDrop(event,'anuncios')">
          <div class="pub-drop-ico">📣</div>
          <div class="pub-drop-t">Reporte de Anuncios</div>
          <div class="pub-drop-s">Publicidad → Anuncios → Exportar xlsx</div>
          <div class="pub-drop-st" id="pub-status-an"></div>
        </div>
        <div style="flex:1;min-width:180px;max-width:220px;padding:12px 16px;background:#faf5ff;border-radius:10px;border:1px solid #e9d5ff;align-self:center">
          <div style="font-size:11px;font-weight:700;color:#6d28d9;margin-bottom:6px">ℹ️ Cómo exportar</div>
          <div style="font-size:10px;color:#6b7280;line-height:1.7">Publicidad → Campañas / Anuncios<br>Elegí el rango → <b>Exportar</b> → xlsx</div>
        </div>
      </div>
    </div>

    <!-- ── Dashboard (oculto hasta cargar) ───────────────────── -->
    <div id="pub-dash" style="display:none;margin-top:14px">

      <!-- Banner header -->
      <div class="pub-wrap" style="margin-bottom:14px">
        <div class="pub-banner">
          <div class="pub-banner-top">
            <div>
              <div class="pub-banner-tag">
                <span class="pub-banner-badge">ML Ads</span>
                <span class="pub-banner-title">Reporte de Publicidad</span>
              </div>
              <div class="pub-banner-sub" id="pub-banner-period">MercadoLibre · Últimos 90 días</div>
            </div>
            <div class="pub-banner-camps" id="pub-banner-camps"></div>
          </div>
          <div class="pub-hero-kpis" id="pub-hero-kpis"></div>
        </div>

        <!-- Mes tabs -->
        <div class="pub-mes-tabs-wrap">
          <span class="pub-mes-lbl-txt">Período</span>
          <div id="pub-mes-tabs" style="display:flex;gap:5px;flex-wrap:wrap"></div>
          <span class="pub-period-lbl" id="pub-mes-label"></span>
        </div>

        <!-- Monthly cards strip -->
        <div class="pub-meses-strip" id="pub-meses-strip" style="display:none"></div>
      </div>

      <!-- Charts row -->
      <div class="pub-charts-row" style="margin-bottom:14px">
        <div class="pub-chart-card">
          <div class="pub-chart-head">
            <div>
              <div class="pub-chart-title" id="pub-ct-ing">Ingresos vs Inversión</div>
              <div class="pub-chart-sub" id="pub-cs-ing">Por mes · ARS</div>
            </div>
          </div>
          <div class="pub-chart-body"><canvas id="pub-chart-ing"></canvas></div>
        </div>
        <div class="pub-chart-card">
          <div class="pub-chart-head">
            <div>
              <div class="pub-chart-title" id="pub-ct-acos">ACOS Mensual</div>
              <div class="pub-chart-sub" id="pub-cs-acos">% Inversión / Ingresos · objetivo &lt; 40%</div>
            </div>
          </div>
          <div class="pub-chart-body"><canvas id="pub-chart-acos"></canvas></div>
        </div>
      </div>

      <!-- Campaign cards -->
      <div class="pub-wrap" style="margin-bottom:14px">
        <div style="padding:14px 20px 10px;border-bottom:1px solid #f1f5f9;display:flex;align-items:center;justify-content:space-between">
          <div>
            <div style="font-size:13px;font-weight:700;color:#0f172a">Rendimiento por Campaña</div>
            <div class="pub-chart-sub" id="pub-cs-camp">Acumulado del período</div>
          </div>
        </div>
        <div style="padding:14px">
          <div class="pub-camps-grid" id="pub-camp-cards"></div>
        </div>
      </div>

      <!-- Top productos por rentabilidad -->
      <div class="pub-wrap" style="margin-bottom:14px">
        <div style="padding:14px 20px 10px;border-bottom:1px solid #f1f5f9;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px">
          <div>
            <div style="font-size:13px;font-weight:700;color:#0f172a">Top Productos por Rentabilidad</div>
            <div class="pub-chart-sub" id="pub-cs-tabla-camp">ACOS más bajo · solo productos con ventas</div>
          </div>
          <span style="font-size:10px;color:#94a3b8">Objetivo ACOS &lt; 40%</span>
        </div>
        <div style="overflow-x:auto">
          <table class="pub-top-table" id="pub-tabla-camp">
            <thead><tr>
              <th style="width:32px">#</th>
              <th>Producto</th>
              <th class="nr">Ingresos</th>
              <th class="nr">Inversión</th>
              <th class="nr">ACOS</th>
              <th class="nr">ROAS</th>
              <th class="nr">Clics</th>
              <th class="nr">Ventas</th>
            </tr></thead>
            <tbody></tbody>
          </table>
        </div>
      </div>

      <!-- Top productos por visibilidad -->
      <div class="pub-wrap" id="pub-anuncios-card" style="display:none;margin-bottom:14px">
        <div style="padding:14px 20px 10px;border-bottom:1px solid #f1f5f9">
          <div style="font-size:13px;font-weight:700;color:#0f172a">Top Productos por Visibilidad</div>
          <div class="pub-chart-sub" id="pub-cs-anuncios">Impresiones · Clics · Ingresos</div>
        </div>
        <div style="overflow-x:auto">
          <table class="pub-top-table" id="pub-tabla-an">
            <thead><tr>
              <th style="width:32px">#</th>
              <th>Producto</th>
              <th class="nr">Impresiones</th>
              <th class="nr">Clics</th>
              <th class="nr">Ingresos</th>
              <th class="nr">ACOS</th>
              <th class="nr">ROAS</th>
              <th class="nr">Ventas</th>
            </tr></thead>
            <tbody></tbody>
          </table>
        </div>
      </div>

    </div><!-- /pub-dash -->
  </div><!-- /pub-module -->

  <!-- ═══════════════════════════════════════════
       MÓDULO PUBLICACIONES ML
  ══════════════════════════════════════════════ -->
  <style>
  #pubs-module{padding:16px;overflow-y:auto;flex:1}
  .pubs-wrap{max-width:1100px;margin:0 auto}
  /* Upload zone */
  .pubs-upload-zone{border:2px dashed #cbd5e1;border-radius:12px;padding:32px;text-align:center;cursor:pointer;transition:all .2s;background:#f8fafc;margin-bottom:20px}
  .pubs-upload-zone:hover,.pubs-upload-zone.drag{border-color:#0d9488;background:#f0fdfa}
  .pubs-upload-zone.loaded{border-color:#0d9488;border-style:solid;background:#f0fdfa}
  .puz-ico{font-size:28px;margin-bottom:8px}
  .puz-t{font-size:14px;font-weight:600;color:#1e293b;margin-bottom:4px}
  .puz-s{font-size:12px;color:#64748b}
  .puz-btn{margin-top:12px;padding:8px 20px;background:#0d9488;color:#fff;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer}
  .puz-btn:hover{background:#0f766e}
  /* KPIs */
  .pubs-kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:18px}
  .pk{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:13px 15px}
  .pk-l{font-size:10px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px}
  .pk-v{font-size:22px;font-weight:700;color:#1e293b}
  .pk-s{font-size:11px;color:#94a3b8;margin-top:2px}
  /* Tabs */
  .pubs-tabs{display:flex;gap:2px;border-bottom:1px solid #e2e8f0;margin-bottom:16px}
  .pubs-tab{padding:8px 15px;font-size:12px;font-weight:600;color:#64748b;cursor:pointer;border:none;background:none;border-bottom:2px solid transparent;margin-bottom:-1px;transition:all .15s}
  .pubs-tab:hover{color:#1e293b}
  .pubs-tab.on{color:#7c3aed;border-bottom-color:#7c3aed}
  /* ── Módulo Tendencias ──────────────────────────── */
  #mkt-module{padding:16px;overflow-y:auto;flex:1}
  .mkt-wrap{max-width:1100px;margin:0 auto}
  .mkt-top{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin-bottom:18px;flex-wrap:wrap}
  .mkt-title{font-size:17px;font-weight:700;color:var(--tx1)}
  .mkt-sub{font-size:11px;color:var(--tx2);margin-top:3px}
  .mkt-sel{padding:5px 10px;border:0.5px solid var(--bd2);border-radius:8px;background:var(--bg2);color:var(--tx1);font-size:12px;cursor:pointer}
  .muz-btn{padding:7px 16px;background:#0ea5e9;color:#fff;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;transition:.15s}
  .muz-btn:hover{background:#0284c7}
  .mkt-dz{border:1.5px dashed #0ea5e9;border-radius:12px;padding:40px;text-align:center;cursor:pointer;transition:.2s;margin-bottom:20px;background:rgba(14,165,233,.03)}
  .mkt-dz:hover,.mkt-dz.drag{background:rgba(14,165,233,.08);border-style:solid}
  .mkt-dz.loaded{border-style:solid;border-color:#0284c7;background:rgba(14,165,233,.06)}
  .mkt-kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:18px}
  .mkt-kpi{background:var(--bg2);border-radius:var(--r2);padding:12px 14px}
  .mkt-kpi-l{font-size:10px;color:var(--tx2);margin-bottom:3px}
  .mkt-kpi-v{font-size:18px;font-weight:700;color:var(--tx1)}
  .mkt-kpi-s{font-size:10px;color:var(--tx2);margin-top:2px}
  .mkt-tabs{display:flex;gap:4px;border-bottom:0.5px solid var(--bd2);margin-bottom:16px}
  .mkt-tab{padding:8px 14px;font-size:12px;font-weight:500;color:var(--tx2);cursor:pointer;border:none;border-bottom:2px solid transparent;background:none;margin-bottom:-1px}
  .mkt-tab:hover{color:var(--tx1)}
  .mkt-tab.on{color:#0ea5e9;border-bottom-color:#0ea5e9}
  .mkt-panel{display:none}
  .mkt-panel.on{display:block}
  .mkt-panel-head{display:flex;align-items:center;gap:12px;margin-bottom:12px}
  .mkt-panel-title{font-size:13px;font-weight:600;color:var(--tx1)}
  .mkt-panel-sub{font-size:11px;color:var(--tx2)}
  .mkt-tbl{width:100%;border-collapse:collapse;font-size:11px}
  .mkt-tbl th{text-align:left;padding:7px 8px;font-size:10px;font-weight:600;color:var(--tx2);border-bottom:0.5px solid var(--bd2);white-space:nowrap}
  .mkt-tbl td{padding:8px;border-bottom:0.5px solid var(--bd2);color:var(--tx1);vertical-align:middle}
  .mkt-tbl tr:last-child td{border-bottom:none}
  .mkt-tbl tr:hover td{background:var(--bg2)}
  .mkt-nm{display:block;max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-weight:500}
  .mkt-chip{font-size:10px;padding:2px 7px;border-radius:10px;font-weight:500}
  .mkt-chip.si{background:#dcfce7;color:#166534}
  .mkt-chip.no{background:#f3f4f6;color:#4b5563}
  .mkt-chip.full{background:#dbeafe;color:#1e40af}
  .mkt-chart-card{background:var(--bg2);border-radius:var(--r2);padding:14px 16px;border:0.5px solid var(--bd2)}
  .mkt-chart-card.full2{grid-column:span 2}
  .mkt-ct{font-size:12px;font-weight:600;color:var(--tx1);margin-bottom:2px}
  .mkt-cs{font-size:10px;color:var(--tx2);margin-bottom:10px}
  .mkt-chart-wrap{position:relative;height:220px}
  .mkt-chart-wrap.tall{height:300px}
  .mkt-bench-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
  .mkt-bench-card{background:var(--bg2);border-radius:var(--r2);padding:14px 16px;border:0.5px solid var(--bd2)}
  .mkt-bench-card h4{font-size:11px;font-weight:600;color:var(--tx2);margin-bottom:8px;text-transform:uppercase;letter-spacing:.4px}
  .mkt-bench-row{display:flex;justify-content:space-between;align-items:center;padding:4px 0;font-size:12px;border-bottom:0.5px solid var(--bd2)}
  .mkt-bench-row:last-child{border-bottom:none}
  .mkt-bench-val{font-weight:600;color:var(--tx1)}
  .mkt-mes-tag{font-size:10px;padding:2px 8px;background:#e0f2fe;color:#0369a1;border-radius:10px;font-weight:500}
  .pubs-low-section{margin-top:24px;border:0.5px solid var(--bd2);border-radius:var(--r2);overflow:hidden}
  .pubs-low-header{display:flex;align-items:flex-start;justify-content:space-between;padding:14px 16px;border-bottom:0.5px solid var(--bd2);gap:12px;flex-wrap:wrap}
  .pubs-low-fbtn{padding:4px 12px;font-size:11px;font-weight:500;border:0.5px solid var(--bd2);border-radius:20px;background:none;cursor:pointer;color:var(--tx2);transition:.15s}
  .pubs-low-fbtn:hover{background:var(--bg2)}
  .pubs-low-fbtn.on{background:#7c3aed;color:#fff;border-color:#7c3aed}
  .pubs-low-kpi{background:var(--bg2);border-radius:var(--r1);padding:10px 12px}
  .pubs-low-kpi .lk-l{font-size:10px;color:var(--tx2);margin-bottom:3px}
  .pubs-low-kpi .lk-v{font-size:18px;font-weight:600;color:var(--tx1)}
  .pubs-low-kpi .lk-s{font-size:10px;color:var(--tx2);margin-top:2px}
  .diag-chip{font-size:10px;padding:2px 8px;border-radius:10px;font-weight:500;white-space:nowrap}
  .diag-novent{background:#fce7f3;color:#9d174d}
  .diag-bconv{background:#fef3c7;color:#92400e}
  .diag-sinvis{background:#f3f4f6;color:#4b5563}
  .diag-ok{background:#dcfce7;color:#166534}
  .pubs-panel{display:none}
  .pubs-panel.on{display:block}
  /* Section header */
  .pubs-sh{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}
  .pubs-sh-t{font-size:13px;font-weight:600;color:#1e293b}
  .pubs-sh-s{font-size:11px;color:#64748b}
  /* Table */
  .pubs-tbl{width:100%;border-collapse:collapse;font-size:12px;background:#fff;border-radius:10px;overflow:hidden;border:1px solid #e2e8f0}
  .pubs-tbl thead th{background:#f8fafc;padding:9px 11px;font-size:11px;font-weight:600;color:#64748b;text-align:left;border-bottom:1px solid #e2e8f0}
  .pubs-tbl td{padding:9px 11px;border-bottom:1px solid #f1f5f9;color:#1e293b;vertical-align:middle}
  .pubs-tbl tr:last-child td{border-bottom:none}
  .pubs-tbl tr:hover td{background:#f8fafc}
  .pubs-tbl .nr{text-align:right;font-variant-numeric:tabular-nums}
  .pubs-badge{font-size:10px;padding:2px 8px;border-radius:10px;font-weight:600}
  .pubs-badge.act{background:#dcfce7;color:#15803d}
  .pubs-badge.ina{background:#f1f5f9;color:#475569}
  .pubs-nm{max-width:210px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-weight:500}
  .pubs-bar{display:flex;align-items:center;gap:5px}
  .pubs-bar-track{width:60px;height:4px;background:#e2e8f0;border-radius:2px;overflow:hidden}
  .pubs-bar-fill{height:100%;border-radius:2px;transition:width .3s}
  .pubs-chip{font-size:11px;padding:2px 8px;border-radius:10px;font-weight:600}
  .pubs-chip.ok{background:#dcfce7;color:#15803d}
  .pubs-chip.warn{background:#fef9c3;color:#854d0e}
  .pubs-chip.bad{background:#fee2e2;color:#b91c1c}
  /* Charts */
  .pubs-charts{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:16px}
  .pubs-chart-card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px 16px}
  .pubs-chart-card.full{grid-column:1/-1}
  .pubs-ct{font-size:12px;font-weight:600;color:#1e293b;margin-bottom:2px}
  .pubs-cs{font-size:11px;color:#94a3b8;margin-bottom:10px}
  .pubs-chart-wrap{position:relative;height:200px}
  .pubs-chart-wrap.tall{height:240px}
  </style>
  <input type="file" id="fi-pubs" accept=".xlsx,.xls,.csv" onchange="onPubsFI(event)">

  <div id="pubs-module" style="display:none;flex:1;overflow-y:auto">
    <div style="padding:16px">
    <div class="pubs-wrap">

      <!-- Upload -->
      <div class="pubs-upload-zone" id="pubs-dz"
           onclick="document.getElementById('fi-pubs').click()"
           ondragover="pubsDov(event)" ondragleave="pubsDlv()" ondrop="pubsDrop(event)">
        <div class="puz-ico">📋</div>
        <div class="puz-t" id="pubs-dz-t">Cargar reporte de publicaciones</div>
        <div class="puz-s" id="pubs-dz-s">MercadoLibre → Publicaciones → Exportar reporte · .xlsx / .xls / .csv</div>
        <button class="puz-btn" onclick="event.stopPropagation();document.getElementById('fi-pubs').click()">Seleccionar archivo</button>
      </div>

      <!-- Dashboard (hidden until loaded) -->
      <div id="pubs-dash" style="display:none">

        <!-- KPIs -->
        <div class="pubs-kpis" id="pubs-kpis"></div>

        <!-- Tabs -->
        <div class="pubs-tabs">
          <button class="pubs-tab on" onclick="pubsTab('ventas',this)">Por ventas brutas</button>
          <button class="pubs-tab" onclick="pubsTab('visitas',this)">Por visitas</button>
          <button class="pubs-tab" onclick="pubsTab('cvr',this)">Por conversión</button>
          <button class="pubs-tab" onclick="pubsTab('graficos',this)">Gráficos</button>
        </div>

        <div id="pubs-panel-ventas" class="pubs-panel on">
          <div class="pubs-sh"><span class="pubs-sh-t">Top publicaciones por ventas brutas</span><span class="pubs-sh-s" id="pubs-periodo-v"></span></div>
          <table class="pubs-tbl"><thead><tr>
            <th>#</th><th>Publicación</th><th>Estado</th>
            <th class="nr">Visitas</th><th class="nr">Ventas</th><th class="nr">Uds.</th><th class="nr">Conv.</th><th class="nr">Ventas brutas</th>
          </tr></thead><tbody id="pubs-tb-ventas"></tbody></table>
        </div>

        <div id="pubs-panel-visitas" class="pubs-panel">
          <div class="pubs-sh"><span class="pubs-sh-t">Top publicaciones por tráfico</span><span class="pubs-sh-s" id="pubs-periodo-vis"></span></div>
          <table class="pubs-tbl"><thead><tr>
            <th>#</th><th>Publicación</th><th>Estado</th>
            <th class="nr">Visitas</th><th class="nr">CVR</th><th colspan="2">Barra</th>
          </tr></thead><tbody id="pubs-tb-visitas"></tbody></table>
        </div>

        <div id="pubs-panel-cvr" class="pubs-panel">
          <div class="pubs-sh"><span class="pubs-sh-t">Top publicaciones por conversión</span><span class="pubs-sh-s">Mín. 10 visitas</span></div>
          <table class="pubs-tbl"><thead><tr>
            <th>#</th><th>Publicación</th><th>Estado</th>
            <th class="nr">Visitas</th><th class="nr">Ventas</th><th class="nr">CVR</th><th colspan="2">Barra</th>
          </tr></thead><tbody id="pubs-tb-cvr"></tbody></table>
        </div>

        <div id="pubs-panel-graficos" class="pubs-panel">
          <div class="pubs-charts">
            <div class="pubs-chart-card full">
              <div class="pubs-ct">Ventas brutas · Top 10 publicaciones</div>
              <div class="pubs-cs">Ordenado de mayor a menor ingreso</div>
              <div class="pubs-chart-wrap tall"><canvas id="pubs-ch-ventas" role="img" aria-label="Top 10 publicaciones por ventas brutas">Top publicaciones por ingresos.</canvas></div>
            </div>
            <div class="pubs-chart-card">
              <div class="pubs-ct">Visitas vs. Conversiones</div>
              <div class="pubs-cs">Top 8 por tráfico</div>
              <div class="pubs-chart-wrap"><canvas id="pubs-ch-visitas" role="img" aria-label="Visitas y conversiones por publicación">Visitas y ventas por publicación.</canvas></div>
            </div>
            <div class="pubs-chart-card">
              <div class="pubs-ct">Tasa de conversión (CVR %)</div>
              <div class="pubs-cs">Top 8 con min. 10 visitas · línea = objetivo 2%</div>
              <div class="pubs-chart-wrap"><canvas id="pubs-ch-cvr" role="img" aria-label="CVR por publicación">Tasa de conversión por publicación.</canvas></div>
            </div>
            <div class="pubs-chart-card">
              <div class="pubs-ct">Publicaciones activas vs. inactivas</div>
              <div class="pubs-cs">Distribución del catálogo</div>
              <div class="pubs-chart-wrap" style="height:160px"><canvas id="pubs-ch-estado" role="img" aria-label="Publicaciones activas vs inactivas">Estado del catálogo.</canvas></div>
            </div>
            <div class="pubs-chart-card">
              <div class="pubs-ct">Visitas · Activas vs. Inactivas</div>
              <div class="pubs-cs">Proporción de tráfico por estado</div>
              <div class="pubs-chart-wrap" style="height:160px"><canvas id="pubs-ch-vis-estado" role="img" aria-label="Visitas por estado de publicación">Visitas según estado.</canvas></div>
            </div>
          </div>

          <!-- ── Sección: Publicaciones con pocas ventas ──────────────── -->
          <div class="pubs-low-section">
            <div class="pubs-low-header">
              <div>
                <div class="pubs-ct" style="margin:0">📉 Publicaciones con pocas ventas</div>
                <div class="pubs-cs" style="margin:0">Cargá el reporte de productos para analizar oportunidades de mejora</div>
              </div>
              <div style="display:flex;gap:8px;align-items:center">
                <button class="puz-btn" style="margin:0" onclick="document.getElementById('fi-pubs-low').click()">
                  ↑ Cargar reporte
                </button>
                <input type="file" id="fi-pubs-low" accept=".xlsx,.xls,.csv" style="display:none" onchange="onPubsLowFI(event)">
              </div>
            </div>

            <!-- Estado sin datos -->
            <div id="pubs-low-empty" style="padding:32px;text-align:center;color:var(--tx2);font-size:12px">
              <div style="font-size:28px;margin-bottom:8px">📂</div>
              Cargá el reporte de rendimiento de publicaciones para ver las publicaciones con pocas ventas o con visitas pero sin conversión.
            </div>

            <!-- Tabla con datos -->
            <div id="pubs-low-data" style="display:none">
              <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:14px" id="pubs-low-kpis"></div>
              <div style="display:flex;gap:6px;margin-bottom:10px;flex-wrap:wrap" id="pubs-low-filters">
                <button class="pubs-low-fbtn on" onclick="filterLow('todas',this)">Todas</button>
                <button class="pubs-low-fbtn" onclick="filterLow('sin_ventas',this)">Sin ventas</button>
                <button class="pubs-low-fbtn" onclick="filterLow('baja_conv',this)">Baja conversión</button>
                <button class="pubs-low-fbtn" onclick="filterLow('activas',this)">Solo activas</button>
              </div>
              <table class="pubs-tbl" style="font-size:11px">
                <thead><tr>
                  <th>#</th><th>Publicación</th><th>Estado</th>
                  <th style="text-align:right">Visitas</th>
                  <th style="text-align:right">Ventas</th>
                  <th style="text-align:right">CVR</th>
                  <th style="text-align:right">Ventas brutas</th>
                  <th>Diagnóstico</th>
                </tr></thead>
                <tbody id="pubs-low-tbody"></tbody>
              </table>
            </div>
          </div>
        </div>

      </div><!-- /pubs-dash -->
    </div>
    </div>
  </div><!-- /pubs-module -->


  <!-- ══ MÓDULO TENDENCIAS / MERCADO ══════════════════════════════════════ -->
  <div id="mkt-module" style="display:none;flex:1;overflow-y:auto;flex-direction:column">
    <div style="padding:16px">
    <div class="mkt-wrap">

    <!-- Header -->
    <div class="mkt-top">
      <div>
        <div class="mkt-title">📈 Tendencias de Mercado</div>
        <div class="mkt-sub">Benchmark de los más vendidos en tus categorías · MercadoLibre</div>
      </div>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
        <select id="mkt-mes-sel" class="mkt-sel" onchange="mktFiltrar()" style="display:none">
          <option value="__all__">Todos los meses</option>
        </select>
        <select id="mkt-cat-sel" class="mkt-sel" onchange="mktFiltrar()" style="display:none">
          <option value="__all__">Todas las categorías</option>
        </select>
        <button class="muz-btn" onclick="document.getElementById('fi-mkt').click()">
          ↑ Cargar reporte
        </button>
      </div>
    </div>

    <!-- Upload zone -->
    <div id="mkt-dz" class="mkt-dz" ondragover="mktDov(event)" ondragleave="mktDlv()" ondrop="mktDrop(event)" onclick="document.getElementById('fi-mkt').click()">
      <div style="font-size:28px;margin-bottom:8px">📊</div>
      <div class="muz-btn" style="pointer-events:none">Cargar reporte de benchmark</div>
      <div style="font-size:11px;color:var(--tx2);margin-top:8px">
        Arrastrá o hacé click · Podés cargar múltiples meses a la vez
      </div>
      <div style="font-size:10px;color:var(--tx2);margin-top:4px">
        Formato: <b>benchmark_market-Más_vendidos_en_tus_categorías_YYYY-MM-01a...</b>
      </div>
    </div>

    <!-- Dashboard con datos -->
    <div id="mkt-dash" style="display:none">

      <!-- KPIs -->
      <div class="mkt-kpis" id="mkt-kpis"></div>

      <!-- Tabs -->
      <div class="mkt-tabs">
        <button class="mkt-tab on" onclick="mktTab('ranking',this)">🏆 Ranking</button>
        <button class="mkt-tab" onclick="mktTab('graficos',this)">📊 Gráficos</button>
        <button class="mkt-tab" onclick="mktTab('benchmark',this)">📐 Benchmark</button>
        <button class="mkt-tab" onclick="mktTab('evolucion',this)">📅 Evolución</button>
      </div>

      <!-- Panel: Ranking -->
      <div id="mkt-panel-ranking" class="mkt-panel on">
        <div class="mkt-panel-head">
          <span class="mkt-panel-title">Top 100 más vendidos</span>
          <span id="mkt-ranking-sub" class="mkt-panel-sub"></span>
        </div>
        <table class="mkt-tbl">
          <thead><tr>
            <th>#</th><th>Producto</th><th>Cat.</th>
            <th>Cond.</th>
            <th style="text-align:right">Precio</th>
            <th style="text-align:right">Uds. vendidas</th>
            <th style="text-align:right">Vistas</th>
            <th style="text-align:right">Conv.</th>
            <th>Envío</th><th>Publicidad</th><th>Cuotas</th>
          </tr></thead>
          <tbody id="mkt-tbody-ranking"></tbody>
        </table>
      </div>

      <!-- Panel: Gráficos -->
      <div id="mkt-panel-graficos" class="mkt-panel">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
          <div class="mkt-chart-card full2">
            <div class="mkt-ct">Top 10 por unidades vendidas</div>
            <div class="mkt-cs" id="mkt-ch-top-sub"></div>
            <div class="mkt-chart-wrap tall"><canvas id="mkt-ch-top"></canvas></div>
          </div>
          <div class="mkt-chart-card full2">
            <div class="mkt-ct">Distribución de precios</div>
            <div class="mkt-cs">Rangos de precio · top 100</div>
            <div class="mkt-chart-wrap"><canvas id="mkt-ch-precios"></canvas></div>
          </div>
          <div class="mkt-chart-card">
            <div class="mkt-ct">Publicidad vs Sin publicidad</div>
            <div class="mkt-cs">% de los más vendidos que usan publicidad</div>
            <div class="mkt-chart-wrap" style="height:160px"><canvas id="mkt-ch-pub"></canvas></div>
          </div>
          <div class="mkt-chart-card">
            <div class="mkt-ct">Tipo de envío</div>
            <div class="mkt-cs">Full, Normal, etc.</div>
            <div class="mkt-chart-wrap" style="height:160px"><canvas id="mkt-ch-envio"></canvas></div>
          </div>
          <div class="mkt-chart-card">
            <div class="mkt-ct">Catálogo vs No catálogo</div>
            <div class="mkt-cs">Productos de catálogo en el top</div>
            <div class="mkt-chart-wrap" style="height:160px"><canvas id="mkt-ch-cat"></canvas></div>
          </div>
          <div class="mkt-chart-card">
            <div class="mkt-ct">Cuotas disponibles</div>
            <div class="mkt-cs">% que ofrece cuotas</div>
            <div class="mkt-chart-wrap" style="height:160px"><canvas id="mkt-ch-cuotas"></canvas></div>
          </div>
        </div>
      </div>

      <!-- Panel: Benchmark -->
      <div id="mkt-panel-benchmark" class="mkt-panel">
        <div class="mkt-panel-head">
          <span class="mkt-panel-title">Estadísticas del mercado</span>
          <span id="mkt-bench-sub" class="mkt-panel-sub"></span>
        </div>
        <div class="mkt-bench-grid" id="mkt-bench-grid"></div>
      </div>

      <!-- Panel: Evolución -->
      <div id="mkt-panel-evolucion" class="mkt-panel">
        <div class="mkt-panel-head">
          <span class="mkt-panel-title">Evolución mensual</span>
          <span class="mkt-panel-sub">Cargá reportes de varios meses para ver la tendencia</span>
        </div>
        <div id="mkt-evol-empty" style="padding:40px;text-align:center;color:var(--tx2);font-size:12px">
          <div style="font-size:28px;margin-bottom:8px">📅</div>
          Necesitás cargar al menos 2 meses para ver la evolución.
        </div>
        <div id="mkt-evol-charts" style="display:none">
          <div class="mkt-chart-card full2" style="margin-bottom:14px">
            <div class="mkt-ct">Unidades vendidas · evolución mensual</div>
            <div class="mkt-cs">Top 5 productos acumulado por mes</div>
            <div class="mkt-chart-wrap tall"><canvas id="mkt-ch-evol-uds"></canvas></div>
          </div>
          <div class="mkt-chart-card full2">
            <div class="mkt-ct">Precio promedio mensual</div>
            <div class="mkt-cs">Precio mediano del top 100 por mes</div>
            <div class="mkt-chart-wrap"><canvas id="mkt-ch-evol-precio"></canvas></div>
          </div>
        </div>
      </div>

    </div><!-- /mkt-dash -->
    </div><!-- /mkt-wrap -->
    </div>
  </div><!-- /mkt-module -->

  <!-- Dashboard -->
  <div id="dash" style="display:none">
    <!-- Tabs de plataforma -->
    <div class="ptf" id="ptf-bar">
      <button class="ptf-btn all act" onclick="setPlatform('__all__',this)">
        <div class="ptf-dot" style="background:var(--acc)"></div>
        Todas las plataformas
        <span class="ptf-cnt all" id="ptf-cnt-all">—</span>
      </button>
      <button class="ptf-btn ml-t" id="ptf-ml" onclick="setPlatform('Mercado Libre',this)" style="display:none">
        <div class="ptf-dot" style="background:var(--ml)"></div>
        Mercado Libre
        <span class="ptf-cnt ml" id="ptf-cnt-ml">—</span>
      </button>
      <button class="ptf-btn tn-t" id="ptf-tn" onclick="setPlatform('Tienda Nube',this)" style="display:none">
        <div class="ptf-dot" style="background:var(--tn)"></div>
        Tienda Nube
        <span class="ptf-cnt tn" id="ptf-cnt-tn">—</span>
      </button>
    </div>
    <div class="dh">
      <div><div class="dt" id="dash-title">Resumen de ventas</div><div style="font-size:11px;color:var(--mu);margin-top:2px" id="dm"></div></div>
      <div class="dash-actions">
        <button class="pdf-btn" id="pdf-btn" onclick="downloadDashboardPDF()" disabled>
          <span aria-hidden="true">PDF</span>
          Descargar PDF
        </button>
        <span class="nb" id="nb"></span>
      </div>
    </div>
    <div id="kpis"></div>
    <div id="cmp-bar" class="cmp-bar" style="display:none"></div>

    <div class="g3" style="margin-bottom:12px">
      <div class="card c2">
        <div class="ch">
          <div><div class="ct">Ventas en el tiempo</div><div class="cs">Por canal / fuente</div></div>
          <div class="ptabs">
            <button class="ptab act" onclick="sp('dia',this)">Día</button>
            <button class="ptab" onclick="sp('mes',this)">Mes</button>
            <button class="ptab" onclick="sp('semana',this)">Semana</button>
            <button class="ptab" onclick="sp('anio',this)">Año</button>
          </div>
        </div>
        <div class="cb"><div class="cw"><canvas id="c-t"></canvas></div></div>
      </div>
      <div class="card">
        <div class="ch"><div class="ct" id="ct-fu">Por canal</div><div class="cs">% ingresos</div></div>
        <div class="cb"><div class="cw"><canvas id="c-fu"></canvas></div></div>
      </div>
    </div>

    <div class="g3" style="margin-bottom:12px">
      <div class="card">
        <div class="ch"><div class="ct">Estado</div><div class="cs">Por ingresos</div></div>
        <div class="cb"><div class="cw"><canvas id="c-es"></canvas></div></div>
      </div>
      <div class="card">
        <div class="ch"><div class="ct">Categoría / Tipo</div><div class="cs">Por ingresos</div></div>
        <div class="cb"><div class="cw"><canvas id="c-cat"></canvas></div></div>
      </div>
      <div class="card">
        <div class="ch">
          <div class="ct">Regiones / Provincias</div>
          <div style="display:flex;gap:6px;align-items:center">
            <button id="btn-prov-ing" onclick="setProvSort('ingresos')" style="font-size:11px;padding:3px 10px;border-radius:6px;border:1px solid var(--acc);background:var(--acc);color:#fff;cursor:pointer;font-weight:600">$ Facturación</button>
            <button id="btn-prov-uds" onclick="setProvSort('unidades')" style="font-size:11px;padding:3px 10px;border-radius:6px;border:1px solid var(--br);background:#fff;color:var(--mu);cursor:pointer">📦 Unidades</button>
          </div>
        </div>
        <div class="cb"><div class="cw"><canvas id="c-pr"></canvas></div></div>
      </div>
    </div>

    <!-- Envíos Flex por zona -->
    <div class="card" id="flex-card" style="margin-bottom:12px;display:none">
      <div class="ch">
        <div>
          <div class="ct">📦 Envíos Flex por zona</div>
          <div class="cs" id="flex-subtitle">—</div>
        </div>
        <div style="display:flex;gap:6px;align-items:center">
          <button id="btn-flex-cant" onclick="setFlexSort('cantidad')" style="font-size:11px;padding:3px 10px;border-radius:6px;border:1px solid var(--acc);background:var(--acc);color:#fff;cursor:pointer;font-weight:600">📦 Envíos</button>
          <button id="btn-flex-ing"  onclick="setFlexSort('ingresos')"  style="font-size:11px;padding:3px 10px;border-radius:6px;border:1px solid var(--br);background:#fff;color:var(--mu);cursor:pointer">$ Facturación</button>
        </div>
      </div>
      <!-- KPIs resumen Flex -->
      <div id="flex-kpis" style="display:flex;gap:0;border-bottom:1px solid var(--br)"></div>
      <!-- Gráfico de barras + Mapa de calor -->
      <div style="display:flex;gap:0;border-bottom:1px solid var(--br)">
        <!-- Barras -->
        <div style="flex:1;padding:16px;min-width:0;border-right:1px solid var(--br)">
          <div style="font-size:11px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px">Envíos por zona</div>
          <div style="position:relative;height:280px"><canvas id="c-flex-bar"></canvas></div>
        </div>
        <!-- Mapa de calor geográfico -->
        <div style="flex:1;padding:16px;min-width:0">
          <div style="font-size:11px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px">Mapa de calor — Densidad Flex por zona</div>
          <div id="flex-map" style="height:280px;border-radius:10px;overflow:hidden;border:1px solid var(--br)"></div>
          <div style="margin-top:8px;display:flex;align-items:center;gap:6px">
            <span style="font-size:10px;color:var(--mu)">Pocos</span>
            <div style="flex:1;height:6px;border-radius:3px;background:linear-gradient(to right,#00f,#0ff,#0f0,#ff0,#f00)"></div>
            <span style="font-size:10px;color:var(--mu)">Muchos envíos</span>
          </div>
        </div>
      </div>
    </div>

    <!-- Viz custom -->
    <div class="g2" style="margin-bottom:12px">
      <div class="card">
        <div class="ch">
          <div class="ct">Top publicaciones / productos</div>
          <div style="display:flex;gap:6px;align-items:center">
            <button id="btn-sort-ing" onclick="setPubSort('ingresos')" style="font-size:11px;padding:3px 10px;border-radius:6px;border:1px solid var(--acc);background:var(--acc);color:#fff;cursor:pointer;font-weight:600">$ Facturación</button>
            <button id="btn-sort-uds" onclick="setPubSort('unidades')" style="font-size:11px;padding:3px 10px;border-radius:6px;border:1px solid var(--br);background:#fff;color:var(--mu);cursor:pointer">📦 Unidades</button>
          </div>
        </div>
        <div class="tw"><table>
          <thead><tr><th>Producto</th><th class="nr" id="th-ing">Ingresos ▼</th><th class="nr" id="th-uds">Uds</th><th>Share</th></tr></thead>
          <tbody id="tb-p"></tbody>
        </table></div>
      </div>
      <div class="card" id="viz-custom-card">
        <div class="ch"><div class="ct" id="viz-custom-title">Visualización personalizada</div><div class="cs" id="viz-custom-sub">Configurá en el panel izquierdo</div></div>
        <div class="cb"><div class="cw"><canvas id="c-viz"></canvas></div></div>
      </div>
    </div>

        <!-- Tabla -->
    <div class="card">
      <div class="ch"><div class="ct">Detalle completo</div><div class="cs" id="ts"></div></div>
      <div class="tw" id="tw-main"></div>
      <div class="pg">
        <span id="pi"></span>
        <div class="pgb">
          <button class="pb2" id="pp" onclick="cp(-1)">‹ Anterior</button>
          <button class="pb2" id="pn" onclick="cp(1)">Siguiente ›</button>
        </div>
      </div>
    </div>
  </div>

<div id="cotizador-module" style="display:none;flex:1;overflow-y:auto">
<style>
#cotizador-module{
  --cot-bg:#f7f5f2;
  --cot-panel:#ffffff;
  --cot-ink:#1f1b16;
  --cot-ink-soft:#6b6258;
  --cot-line:#e6e1d9;
  --cot-accent:#7a2e2e;
  --cot-accent-soft:#f1e2e2;
  --cot-good:#2f6b3a;
  --cot-good-bg:#e7f3e8;
  --cot-bad:#a3312c;
  --cot-bad-bg:#fbe8e6;
  --cot-warn:#8a6414;
  --cot-warn-bg:#faf1de;
  --cot-radius:12px;
}
/* Diseño fijo en claro: se ignora el modo oscuro del sistema a pedido de Santi. */
#cotizador-module, #cotizador-module *{box-sizing:border-box;}
#cotizador-module{
  background:var(--cot-bg);
  color:var(--cot-ink);
  font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
  padding:24px 16px 60px;
}
#cotizador-module .wrap{max-width:1080px;margin:0 auto;}
#cotizador-module header{margin-bottom:20px;}
#cotizador-module h1{font-size:22px;margin:0 0 4px;font-weight:700;}
#cotizador-module .sub{color:var(--cot-ink-soft);font-size:14px;margin:0;}

#cotizador-module .panel{
  background:var(--cot-panel);
  border:1px solid var(--cot-line);
  border-radius:var(--cot-radius);
  padding:18px;
  margin-bottom:18px;
}
#cotizador-module .panel h2{font-size:15px;margin:0 0 12px;font-weight:600;}

#cotizador-module .params{
  display:grid;
  grid-template-columns:repeat(auto-fit,minmax(140px,1fr));
  gap:12px;
}
#cotizador-module .field label{
  display:block;
  font-size:12px;
  color:var(--cot-ink-soft);
  margin-bottom:4px;
}
#cotizador-module .field input{
  width:100%;
  padding:8px 10px;
  border-radius:8px;
  border:1px solid var(--cot-line);
  background:var(--cot-bg);
  color:var(--cot-ink);
  font-size:14px;
}
#cotizador-module .field input:focus{outline:2px solid var(--cot-accent);outline-offset:1px;}
#cotizador-module input[type=file]{
  display:block!important;
  width:auto;
  padding:6px 4px;
  border:none;
  background:transparent;
  font-size:13px;
  color:var(--cot-ink);
}

#cotizador-module .addrow{
  display:flex;
  gap:10px;
  flex-wrap:wrap;
  align-items:flex-end;
  margin-bottom:6px;
}
#cotizador-module .addrow .field{flex:1;min-width:150px;}
#cotizador-module .addrow .field.small{flex:0 0 110px;min-width:90px;}
#cotizador-module button{
  font:inherit;
  cursor:pointer;
  border:none;
  border-radius:8px;
  padding:9px 16px;
  font-size:14px;
  font-weight:600;
}
#cotizador-module .btn-primary{background:var(--cot-accent);color:#fff;}
#cotizador-module .btn-primary:hover{opacity:.92;}
#cotizador-module .btn-ghost{background:transparent;color:var(--cot-accent);border:1px solid var(--cot-accent);}
#cotizador-module .btn-ghost:hover{background:var(--cot-accent-soft);}
#cotizador-module .btn-icon{background:transparent;color:var(--cot-ink-soft);padding:4px 8px;font-size:13px;}
#cotizador-module .btn-icon:hover{color:var(--cot-bad);}

#cotizador-module table{width:100%;border-collapse:collapse;font-size:13px;}
#cotizador-module th{
  text-align:left;
  font-size:11px;
  text-transform:uppercase;
  letter-spacing:.03em;
  color:var(--cot-ink-soft);
  padding:8px 10px;
  border-bottom:1px solid var(--cot-line);
  white-space:nowrap;
}
#cotizador-module td{padding:9px 10px;border-bottom:1px solid var(--cot-line);vertical-align:middle;}
#cotizador-module tr:last-child td{border-bottom:none;}
#cotizador-module td.num, #cotizador-module th.num{text-align:right;font-variant-numeric:tabular-nums;}
#cotizador-module .prod-name{font-weight:600;}
#cotizador-module .prod-meta{font-size:11px;color:var(--cot-ink-soft);}

#cotizador-module .badge{
  display:inline-block;
  padding:3px 9px;
  border-radius:999px;
  font-size:11px;
  font-weight:700;
}
#cotizador-module .badge.good{background:var(--cot-good-bg);color:var(--cot-good);}
#cotizador-module .badge.bad{background:var(--cot-bad-bg);color:var(--cot-bad);}

#cotizador-module .pvp-actual{width:110px;padding:6px 8px;border-radius:6px;border:1px solid var(--cot-line);background:var(--cot-bg);color:var(--cot-ink);font-size:13px;text-align:right;}

#cotizador-module .empty{padding:30px;text-align:center;color:var(--cot-ink-soft);font-size:13px;}

#cotizador-module .summary{
  display:grid;
  grid-template-columns:repeat(auto-fit,minmax(160px,1fr));
  gap:12px;
  margin-bottom:18px;
}
#cotizador-module .stat{
  background:var(--cot-panel);
  border:1px solid var(--cot-line);
  border-radius:var(--cot-radius);
  padding:14px 16px;
}
#cotizador-module .stat .n{font-size:22px;font-weight:700;}
#cotizador-module .stat .l{font-size:12px;color:var(--cot-ink-soft);margin-top:2px;}

#cotizador-module footer{margin-top:10px;font-size:11px;color:var(--cot-ink-soft);text-align:center;}

@media (max-width:640px){
  #cotizador-module th, #cotizador-module td{padding:7px 6px;}
  #cotizador-module .prod-meta{display:none;}
}
</style>


<div class="wrap">
  <header>
    <h1>Calculador de costos — Vinos y Bebidas</h1>
    <p class="sub">Cargá el costo de cada producto (el que ves en Mercado Libre / proveedor) y te sugiere el PVP para quedar por encima del margen objetivo, ya descontando comisión ML, IVA, IIBB, débitos/créditos y costo fijo de envío.</p>
  </header>

  <div class="panel">
    <h2>Parámetros (editables)</h2>
    <p class="sub" style="margin-bottom:10px;">Se guardan solos en este navegador, así que podés actualizarlos cuando cambien las comisiones/impuestos (por ej. una vez al año) y quedan como predeterminados la próxima vez que abras la calculadora.</p>
    <div class="params">
      <div class="field">
        <label>Margen objetivo %</label>
        <input type="number" id="cot_p_margen" value="10" step="0.5">
      </div>
      <div class="field"><label>Comisión ML %</label><input type="number" id="cot_p_meli" value="13" step="0.1"></div>
      <div class="field"><label>IVA %</label><input type="number" id="cot_p_iva" value="21" step="0.5"></div>
      <div class="field"><label>Ingresos Brutos %</label><input type="number" id="cot_p_iibb" value="3" step="0.1"></div>
      <div class="field"><label>Ley Déb/Créd %</label><input type="number" id="cot_p_debcred" value="1.2" step="0.1"></div>
      <div class="field"><label>Envío proveedor $</label><input type="number" id="cot_p_envioprov" value="0" step="100"></div>
      <div class="field"><label>Embalaje $</label><input type="number" id="cot_p_embalaje" value="1000" step="50"></div>
    </div>
    <div style="margin-top:12px;display:flex;gap:10px;align-items:center;flex-wrap:wrap;">
      <button class="btn-ghost" onclick="cot_restablecerParametros()">Restablecer valores originales</button>
      <button class="btn-primary" onclick="cot_aplicarMargenATodos()">Aplicar estos datos a todos los productos</button>
      <span id="cot_params_status" style="font-size:12px;color:var(--ink-soft);"></span>
    </div>
    <div id="cot_resumen_impuestos" style="margin-top:12px;font-size:12.5px;color:var(--ink-soft);padding:10px;background:var(--bg);border-radius:8px;"></div>
  </div>

  <div class="panel">
    <h2>Agregar producto</h2>
    <div class="addrow">
      <div class="field"><label>Producto</label><input type="text" id="cot_in_nombre" placeholder="Ej: Malbec Reserva 750ml"></div>
      <div class="field small"><label>Costo $</label><input type="number" id="cot_in_costo" placeholder="0" step="1"></div>
      <div class="field small" style="flex:0 0 130px;">
        <label>Costo incluye IVA</label>
        <select id="cot_in_costoiva" style="width:100%;padding:8px 10px;border-radius:8px;border:1px solid var(--line);background:var(--bg);color:var(--ink);font-size:14px;">
          <option value="si">Sí</option>
          <option value="no">No</option>
        </select>
      </div>
      <div class="field small" style="flex:0 0 170px;">
        <label>Logística</label>
        <select id="cot_in_logistica" style="width:100%;padding:8px 10px;border-radius:8px;border:1px solid var(--line);background:var(--bg);color:var(--ink);font-size:14px;" onchange="cot_toggleCampoPeso()">
          <option value="full">Full / Mercado Envíos</option>
          <option value="flex">Flex</option>
        </select>
      </div>
      <div class="field small" id="cot_campo_peso"><label>Peso (kg) — vacío = automático</label><input type="number" id="cot_in_peso" placeholder="auto" step="0.1"></div>
      <div class="field small"><label>Margen % (vacío = general)</label><input type="number" id="cot_in_margen" placeholder="10" step="0.5"></div>
      <div class="field small"><label>Precio actual $ (opcional)</label><input type="number" id="cot_in_actual" placeholder="0" step="1"></div>
      <div class="field small"><label>Precio ML más vendido $ (opcional)</label><input type="number" id="cot_in_ml" placeholder="0" step="1"></div>
      <button class="btn-primary" onclick="cot_agregarProducto()">Agregar</button>
    </div>
  </div>

  <div class="panel">
    <h2>Subir Excel con varios productos</h2>
    <p class="sub" style="margin-bottom:10px;">El archivo tiene que tener una columna con el nombre del producto y otra con el costo. Opcionalmente podés incluir columnas de precio actual, peso (kg), logística (Full/Mercado Envíos o Flex — si no la incluís, se asume Full/ME), margen % (si querés uno distinto al general para ese producto) y precio ML del más vendido (para comparar). Si el archivo NO trae peso, lo estimamos automáticamente leyendo el título de cada producto. Detectamos las columnas por nombre, no por posición.</p>
    <div class="addrow" style="align-items:center;">
      <div class="field" style="flex:0 0 auto;">
        <label>Archivo .xlsx / .xls / .csv</label>
        <input type="file" id="cot_in_file" accept=".xlsx,.xls,.csv" onchange="cot_manejarArchivo(event)">
      </div>
      <div class="field small" style="flex:0 0 160px;">
        <label>Costo del archivo incluye IVA</label>
        <select id="cot_in_file_costoiva" style="width:100%;padding:8px 10px;border-radius:8px;border:1px solid var(--line);background:var(--bg);color:var(--ink);font-size:14px;">
          <option value="no" selected>No (sin IVA)</option>
          <option value="si">Sí</option>
        </select>
      </div>
    </div>
    <div id="cot_file_status" style="font-size:12px;color:var(--ink-soft);margin-top:4px;"></div>
  </div>

  <div class="summary" id="cot_summary"></div>

  <div id="cot_exportbar" style="display:none;margin-bottom:14px;gap:10px;flex-wrap:wrap;">
    <button class="btn-ghost" onclick="cot_exportarExcel()">Descargar tabla con PVP sugeridos (.xlsx)</button>
    <button class="btn-ghost" onclick="cot_reestimarPesos()">Re-estimar pesos automáticos</button>
    <button class="btn-icon" style="border:1px solid var(--line);" onclick="cot_limpiarTodo()">Limpiar todo</button>
  </div>

  <div class="panel" style="padding:0;overflow-x:auto;">
    <table id="cot_tabla">
      <thead>
        <tr>
          <th>Producto</th>
          <th>Logística</th>
          <th class="num">Costo</th>
          <th class="num">Precio actual / Margen</th>
          <th class="num">Margen %<br>objetivo</th>
          <th class="num">PVP sugerido<br>/ diferencia</th>
          <th class="num">Precio ML<br>más vendido / vs.</th>
          <th></th>
        </tr>
      </thead>
      <tbody id="cot_tbody"></tbody>
    </table>
    <div class="empty" id="cot_empty">Todavía no cargaste productos. Agregá el primero arriba.</div>
  </div>

  <footer>Fórmula: PVP − (comisión ML + IIBB sobre neto + Ley Déb/Créd + costo de envío según logística y peso + costo + envío proveedor + embalaje) = margen objetivo × PVP. Full/Mercado Envíos usa la tabla oficial por peso y tramo de PVP; Flex usa el costo fijo por tramo de PVP. Para PVP ≥ $33.000 (envío gratis) se usa la tabla oficial por peso y tramo ($33.000-$49.999 / $50.000+); para Flex en ese tramo se toma la misma tabla como estimación, por no tener una propia.
</div>


</div>

</div>
</div>

<!-- Modal configurador de columnas -->
<div class="modal-bg" id="modal-cfg">
<div class="modal">
  <div class="modal-head">
    <span class="modal-title" id="modal-cfg-title">Configurar columnas</span>
    <button class="btn btn-ghost" style="padding:4px 10px;font-size:12px" onclick="closeModal()">✕</button>
  </div>
  <div class="modal-body">
    <div class="info-tip">
      Asigná el <strong>rol</strong> de cada columna para que Dashify sepa qué representa cada dato.
      Los campos con ★ son los más importantes para las visualizaciones.
    </div>
    <div class="role-grid" id="role-grid"></div>
    <div style="margin-top:16px">
      <div style="font-size:12px;font-weight:600;margin-bottom:8px;color:var(--mu)">Vista previa de los primeros datos:</div>
      <div style="overflow-x:auto"><table class="preview-table" id="preview-tbl"></table></div>
    </div>
  </div>
  <div class="modal-foot">
    <button class="btn btn-ghost" onclick="closeModal()">Cancelar</button>
    <button class="btn btn-primary" onclick="applyConfig()">Aplicar configuración</button>
  </div>
</div>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

<script>
Chart.register(ChartDataLabels)
Chart.defaults.font.family = "'Segoe UI', system-ui, -apple-system, sans-serif"
Chart.defaults.font.size = 11
Chart.defaults.color = "#707c8c"
Chart.defaults.borderColor = "#eef0f4"
Chart.defaults.plugins.legend.labels.usePointStyle = true
Chart.defaults.plugins.legend.labels.boxWidth = 7
Chart.defaults.plugins.legend.labels.padding = 12
Chart.defaults.plugins.tooltip.backgroundColor = "#1c2433"
Chart.defaults.plugins.tooltip.titleFont = { family: "'Segoe UI', system-ui, sans-serif", weight: "600", size: 12 }
Chart.defaults.plugins.tooltip.bodyFont = { family: "'Segoe UI', system-ui, sans-serif", size: 11 }
Chart.defaults.plugins.tooltip.padding = 10
Chart.defaults.plugins.tooltip.cornerRadius = 6
Chart.defaults.plugins.tooltip.displayColors = true
Chart.defaults.plugins.tooltip.boxPadding = 4
Chart.defaults.scale.grid.color = "#eef0f4"
Chart.defaults.scale.ticks.color = "#8993a3"

const S = { sids:[], periodo:"dia", page:0, charts:{}, _tot:0, catCols:[], numCols:[], platform:"__all__", _flexMap:null, _lastDashboard:null }
const DS = {}  // { sid: { name, source, rows, cols, colInfo, config } }
let cfgSid = null  // sid del dataset que se está configurando

const $ = id => document.getElementById(id)
function sL(m="Procesando..."){$("lm").textContent=m;$("ld").classList.add("show")}
function hL(){$("ld").classList.remove("show")}
function sE(m){$("et").textContent=m;$("eb").classList.add("show")}
function hE(){$("eb").classList.remove("show")}

function fmt(n,p=0){
  if(n==null||isNaN(n))return"—"
  return"$ "+Math.round(n).toLocaleString("es-AR")
}
function fN(n){
  if(n==null||isNaN(n))return"—"
  return Number.isInteger(n)?n.toLocaleString("es-AR"):n.toFixed(1)
}
const KICONS={
  money:'<circle cx="12" cy="12" r="9"/><path d="M12 7v10M9 9.5c0-1.1 1.2-2 3-2s3 .9 3 2-1.2 1.7-3 2-3 .9-3 2 1.2 2 3 2 3-.9 3-2"/>',
  check:'<circle cx="12" cy="12" r="9"/><path d="M7.5 12.5l3 3 6-6.5"/>',
  box:'<path d="M21 7.5l-9-5-9 5 9 5 9-5z"/><path d="M3 7.5v9l9 5 9-5v-9"/><path d="M12 12.5v9"/>',
  receipt:'<path d="M6 2.5h12v19l-3-2-3 2-3-2-3 2v-19z"/><path d="M9 7.5h6M9 11.5h6"/>',
  truck:'<rect x="1.5" y="7.5" width="12" height="8.5"/><path d="M13.5 10h3.5l3 3v3h-6.5"/><circle cx="5.5" cy="18" r="1.5"/><circle cx="16.5" cy="18" r="1.5"/>',
  tag:'<path d="M20 12.5l-7.5 7.5-8.5-8.5v-7.5h7.5l8.5 8.5z"/><circle cx="8" cy="8" r="1.2"/>',
}
function kIco(name,cls){return`<div class="kico ${cls}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${KICONS[name]}</svg></div>`}
function etag(e){
  if(!e||e==="null")return""
  const s=e.toLowerCase()
  if(s.includes("entregado")||s.includes("pagado")||s.includes("paid")||s.includes("completad")||s.includes("aprobado"))
    return`<span class="tag tOK">${e}</span>`
  if(s.includes("camino")||s.includes("enviad")||s.includes("pendiente")||s.includes("procesando"))
    return`<span class="tag tC">${e}</span>`
  if(s.includes("cancelad")||s.includes("reembols"))
    return`<span class="tag tX">${e}</span>`
  return`<span class="tag tD">${e}</span>`
}
function srcTag(f){
  if(!f)return""
  const sl=f.toLowerCase()
  const c=sl.includes("mercado")?"ml":sl.includes("tienda")?"tn":"cx"
  return`<span class="src-bd ${c}" style="border-radius:4px">${f.slice(0,12)}</span>`
}
function dc(id){if(S.charts[id]){S.charts[id].destroy();delete S.charts[id]}}

const COLORS=["#2b5fd9","#0f8f83","#c8860a","#6d43c9","#d33a3a","#1a8a53","#3d7ab8","#a3562c","#5c6bc0","#8a8f36","#c04570","#3892a6"]
const SRC_COLORS={"Mercado Libre":"#f59e0b","Tienda Nube":"#7c3aed"}

function srcColor(name){
  if(!name)return COLORS[0]
  return SRC_COLORS[name] || COLORS[Object.keys(SRC_COLORS).length + Object.keys(DS).findIndex(k=>DS[k].name===name)]
}

// ── Upload ─────────────────────────────────────────────────────────────────
function hideMainViews(){
  ['wlc','dash','pub-module','pubs-module','mkt-module','cotizador-module'].forEach(id=>{
    const el=document.getElementById(id)
    if(el) el.style.display='none'
  })
}

function trig(src){
  // ocultar módulos al cambiar de fuente
  const pub = document.getElementById("pub-module")
  if(pub && pub.style.display !== "none"){
    pub.style.display = "none"
    const tbp = document.getElementById("tbtn-pub")
    if(tbp) tbp.style.outline = "none"
  }
  const pubs = document.getElementById("pubs-module")
  if(pubs && pubs.style.display !== "none"){
    pubs.style.display = "none"
    const tbps = document.getElementById("tbtn-pubs")
    if(tbps) tbps.style.outline = "none"
  }
  const el=$("fi-"+src)
  el.value=""  // reset para permitir subir el mismo archivo dos veces
  el.click()
}
async function onFI(e,src){const f=[...e.target.files];e.target.value="";for(const x of f)await up(x,src)}
function dov(e,s){e.preventDefault();$("dz-"+s).classList.add("dg")}
function dlv(s){$("dz-"+s).classList.remove("dg")}
async function drp(e,s){
  e.preventDefault();$("dz-"+s).classList.remove("dg")
  const ok={ml:/\.(xlsx|xls)$/i,tn:/\.(xlsx|xls|csv)$/i,cx:/\.(xlsx|xls|csv|ods)$/i,fichas:/\.(xlsx|xls)$/i}
  const f=[...e.dataTransfer.files].filter(x=>ok[s].test(x.name))
  if(!f.length){sE("Formato no compatible para este canal");return}
  for(const x of f)await up(x,s)
}

async function showVentasModule(){
  hideMainViews()
  document.querySelectorAll('.tbtn').forEach(b=>b.style.outline='none')
  const btn=document.getElementById('tbtn-ventas')
  if(btn) btn.style.outline='2px solid #86efac'

  if(!S.sids.length){
    const w=document.getElementById('wlc')
    const d=document.getElementById('dash')
    if(w) w.style.display='flex'
    if(d) d.style.display='none'
    sE('Primero cargá un reporte de ventas.')
    return
  }

  const w=document.getElementById('wlc')
  const d=document.getElementById('dash')
  if(w) w.style.display='none'
  if(d) d.style.display='block'
  await rf()
}

async function up(file,src){
  sL("Leyendo archivo...")
  console.log("[Dashify] Subiendo:", file.name, "src:", src, "size:", file.size)
  try{
    const fd=new FormData();fd.append("file",file)
    const r=await fetch("/api/upload/"+src,{method:"POST",body:fd})
    console.log("[Dashify] HTTP status:", r.status)
    const d=await apiJson(r)
    console.log("[Dashify] Respuesta:", d)
    if(!d.ok)throw new Error(d.error)

    // ── Fichas Técnicas ML: manejo especial ──────────────────────────────
    if(src==="fichas"){
      const cats = d.categorias || []
      const statusEl=$("fichas-status")
      if(statusEl){
        statusEl.style.display="block"
        statusEl.textContent=`✓ ${cats.length} categorías · ${d.total_ids||0} IDs · ${d.total_skus||0} SKUs`
      }
      let msg=`📋 Fichas cargadas: ${cats.length} categorías`
      if(d.enriched_rows>0) msg+=` · ${d.enriched_rows} publicaciones con categoría asignada`
      showToast(msg, 6000)
      if(d.sid && d.filtros){
        popF(d.filtros, true)
        await rf()
      }
      return
    }

    // Para ML: actualizar el DS existente si el sid ya estaba registrado (upsert)
    const isNew = !S.sids.includes(d.sid)
    DS[d.sid]={name: src==="ml" ? "Mercado Libre (acumulado)" : d.filename,
               source:src, rows:d.total_rows,
               cols:d.cols, colInfo:d.col_info, config:d.config,
               upsert: d.upsert||null}
    if(isNew) S.sids.push(d.sid)
    // No actualizamos selectores de columna aquí; lo hace rf() con el combinado
    renderSideFiles()
    popF(d.filtros, true)
    // Actualizar selectores de columna con info del upload
    if(d.cols && d.filtros){
      updateColSelectors(d.cols, d.col_info, d.filtros)
    }
    $("wlc").style.display="none"
    $("dash").style.display="block"
    S.page=0

    // Mostrar resumen del upsert para ML
    if(src==="ml" && d.upsert){
      const u=d.upsert
      const msg = u.updated>0
        ? `✓ ML: +${u.added} nuevas, ${u.updated} actualizadas · Total acumulado: ${u.total} ventas`
        : `✓ ML: +${u.added} ventas nuevas cargadas · Total acumulado: ${u.total}`
      showToast(msg)
    }

    // Si es archivo custom, abrir modal de configuración
    if(src==="cx"){
      openCfgModal(d.sid, d.filename, d.cols, d.col_info, d.config)
      return
    }
    await rf()
  }catch(e){
    console.error("[Dashify] Error:", e)
    sE(e.message || "Error desconocido al cargar el archivo")
  }
  finally{hL()}
}

function showToast(msg, ms=4000){
  let t=document.getElementById("toast-notif")
  if(!t){
    t=document.createElement("div")
    t.id="toast-notif"
    t.style.cssText="position:fixed;bottom:20px;right:20px;background:#1e293b;color:#fff;"+
      "padding:10px 16px;border-radius:8px;font-size:12px;font-weight:500;z-index:999;"+
      "box-shadow:0 4px 12px rgba(0,0,0,.3);border-left:3px solid #16a34a;max-width:360px;"+
      "transition:opacity .3s"
    document.body.appendChild(t)
  }
  t.textContent=msg
  t.style.opacity="1"
  clearTimeout(t._tm)
  t._tm=setTimeout(()=>t.style.opacity="0", ms)
}

async function apiJson(r){
  const ct = (r.headers.get("content-type") || "").toLowerCase()
  if(ct.includes("application/json")) return await r.json()
  const txt = await r.text()
  const clean = txt.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim()
  throw new Error(clean || `Error del servidor (${r.status})`)
}

// ── Sidebar archivos ────────────────────────────────────────────────────────
function renderSideFiles(){
  ["ml","tn","cx"].forEach(src=>{
    const el=$("files-"+src)
    const mine=S.sids.filter(sid=>DS[sid]?.source===src)
    if(!mine.length){
      el.innerHTML=`<div style="padding:5px 10px 5px 18px;font-size:11px;color:rgba(255,255,255,.2)">Sin archivos</div>`
      return
    }
    if(src==="ml"){
      // ML: un único bloque acumulado con lista de archivos cargados
      const sid=mine[0]
      const ds=DS[sid]
      const u=ds.upsert
      const fileList=(u&&u.files?u.files:[ds.name])
      el.innerHTML=`
        <div style="padding:5px 10px 5px 14px">
          ${fileList.map(fn=>`
            <div class="src-file" style="padding:3px 0">
              <span class="src-fn" title="${fn}" style="font-size:10px">📄 ${fn}</span>
            </div>`).join("")}
          <div style="margin-top:4px;font-size:10px;color:rgba(255,255,255,.38);padding-left:4px">
            Total acumulado: ${ds.rows} ventas
          </div>
        </div>
        <div style="padding:0 10px 6px 14px">
          <button class="sdl" style="font-size:11px;color:rgba(255,60,60,.5)" onclick="delSrc('${sid}')">✕ Limpiar ML</button>
        </div>`
    } else {
      el.innerHTML=mine.map(sid=>`
        <div class="src-file">
          <span class="src-fn" title="${DS[sid].name}">${DS[sid].name}</span>
          <span class="src-fr">${DS[sid].rows}f</span>
          ${src==="cx"?`<button class="cfg-btn" onclick="openCfgModal('${sid}','${DS[sid].name}',null,null,null,true)">⚙</button>`:""}
          <button class="sdl" onclick="delSrc('${sid}')">x</button>
        </div>`).join("")
    }
  })
}

async function delSrc(sid){
  await fetch("/api/delete/"+sid,{method:"DELETE"})
  S.sids=S.sids.filter(s=>s!==sid)
  delete DS[sid]
  renderSideFiles()
  if(!S.sids.length){$("wlc").style.display="flex";$("dash").style.display="none"}
  else await rf()
}

// ── Modal configurador de columnas ─────────────────────────────────────────
const ROLES = [
  {id:"date",     label:"★ Fecha",          desc:"Columna de fecha/hora de la venta",      color:"#2563eb"},
  {id:"amount",   label:"★ Monto / Ingresos",desc:"Valor principal (ventas, facturación)", color:"#16a34a"},
  {id:"qty",      label:"Cantidad / Unidades",desc:"Número de unidades vendidas",           color:"#0d9488"},
  {id:"status",   label:"★ Estado",          desc:"Estado de la venta o pago",             color:"#d97706"},
  {id:"product",  label:"★ Producto",        desc:"Nombre del producto o publicación",     color:"#7c3aed"},
  {id:"category", label:"Categoría",         desc:"Tipo, categoría o envío (para filtrar)",color:"#0891b2"},
  {id:"geo",      label:"Región / Provincia", desc:"Ubicación geográfica del comprador",   color:"#dc2626"},
  {id:"customer", label:"Cliente",           desc:"Nombre o ID del comprador",             color:"#9333ea"},
  {id:"id",       label:"ID / N° de orden",  desc:"Identificador único de la venta",       color:"#64748b"},
  {id:"cost",     label:"Costo / Descuento", desc:"Costo de envío, descuento u otro costo",color:"#e11d48"},
  {id:"amount2",  label:"Valor secundario",  desc:"Segunda métrica numérica",              color:"#65a30d"},
  {id:"category2",label:"Categoría 2",       desc:"Segunda categoría para agrupar",        color:"#f59e0b"},
]

function openCfgModal(sid, filename, cols, colInfo, config, reopen=false){
  cfgSid = sid
  const ds = DS[sid]
  if(reopen){
    cols = ds.cols
    colInfo = ds.colInfo
    config = ds.config
  }
  $("modal-cfg-title").textContent = "Configurar columnas: " + filename

  const allCols = cols || []
  const currentConfig = config || {}

  // Invertir config para ver qué columna tiene cada rol
  const roleToCol = {}
  for(const [role,col] of Object.entries(currentConfig)) roleToCol[role]=col

  // Construir grid de roles
  const colOptions = `<option value="">— sin asignar —</option>` +
    allCols.map(c=>`<option value="${c}">${c}</option>`).join("")

  $("role-grid").innerHTML = ROLES.map(r=>{
    const current = roleToCol[r.id] || ""
    const auto = colInfo && Object.entries(colInfo).find(([c,i])=>i.role===r.id)?.[0]
    const hint = auto && !current ? ` (sugerido: ${auto})` : ""
    return `<div class="role-item">
      <div class="role-lbl">
        <div class="role-dot" style="background:${r.color}"></div>
        ${r.label}
      </div>
      <div class="role-desc">${r.desc}${hint}</div>
      <select class="role-sel" data-role="${r.id}">
        ${colOptions}
      </select>
    </div>`
  }).join("")

  // Setear valores actuales
  document.querySelectorAll(".role-sel").forEach(sel=>{
    const role = sel.dataset.role
    const val = roleToCol[role] || ""
    sel.value = val
  })

  // Preview table
  const info = colInfo || {}
  const sample = allCols.map(c=>({col:c, type:info[c]?.type||"?", role:info[c]?.role||"", sample:info[c]?.sample||[]}))
  $("preview-tbl").innerHTML = `
    <thead><tr><th>Columna</th><th>Tipo detectado</th><th>Rol sugerido</th><th>Muestra</th></tr></thead>
    <tbody>${sample.slice(0,12).map(s=>`<tr>
      <td style="font-weight:500">${s.col}</td>
      <td><span class="tag" style="background:${s.type==="number"?"#dbeafe":s.type==="date"?"#dcfce7":"#f1f5f9"};color:${s.type==="number"?"#1d4ed8":s.type==="date"?"#15803d":"#475569"}">${s.type}</span></td>
      <td style="color:var(--mu)">${s.role||"—"}</td>
      <td style="color:var(--mu);font-size:10px">${s.sample.slice(0,2).join(", ")}</td>
    </tr>`).join("")}</tbody>`

  $("modal-cfg").classList.add("open")
}

function closeModal(){$("modal-cfg").classList.remove("open")}

async function applyConfig(){
  if(!cfgSid)return
  const newConfig = {}
  document.querySelectorAll(".role-sel").forEach(sel=>{
    if(sel.value) newConfig[sel.dataset.role] = sel.value
  })
  sL("Aplicando configuración...")
  closeModal()
  try{
    const r=await fetch("/api/config/"+cfgSid,{
      method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify(newConfig)
    })
    const d=await apiJson(r)
    if(!d.ok)throw new Error(d.error)
    DS[cfgSid].config=newConfig
    popF(d.filtros, true)
    await rf()
  }catch(e){sE(e.message)}
  finally{hL()}
}

// Cerrar modal al hacer clic fuera
$("modal-cfg").addEventListener("click", e=>{ if(e.target===$("modal-cfg"))closeModal() })

// ── Etiquetas amigables para columnas ──────────────────────────────────────
const COL_LABELS = {
  ingresos:"💰 Ingresos", total_neto:"💵 Ingreso neto", costo:"💸 Costos",
  unidades:"📦 Unidades vendidas", n_ventas:"🛒 Nº de ventas",
  descuentos:"🏷️ Descuentos", ticket_prom:"🎫 Ticket promedio",
  cantidad:"🔢 Cantidad",
}
const CAT_LABELS = {
  publicacion:"🛍️ Producto / Publicación", estado:"🚦 Estado de venta",
  provincia:"📍 Provincia", categoria:"🏷️ Categoría", fuente:"🏪 Canal de venta",
  forma_envio:"🚚 Tipo de envío", comprador:"👤 Comprador",
  _mes:"📅 Mes", _fecha_str:"📆 Fecha", _sem:"📅 Semana", _anio:"📅 Año",
}
function colLabel(c, type="cat"){
  const map = type==="num" ? COL_LABELS : CAT_LABELS
  return map[c] || c.replace(/_/g," ").replace(/\w/g,l=>l.toUpperCase())
}

// ── Selectores de visualización ────────────────────────────────────────────
function updateColSelectors(cols, colInfo, filtrosDisp){
  if(!cols)return
  const numColsFromServer = filtrosDisp?.num_cols || []
  const catColsFromServer = filtrosDisp?.cat_cols || []
  const allCols = (filtrosDisp?.all_cols || cols).filter(c=>!c.startsWith("_"))

  // catCols = texto/categoría — para agrupar (eje X)
  S.catCols = catColsFromServer.length ? catColsFromServer
    : allCols.filter(c=>!numColsFromServer.includes(c))

  // numCols = numéricas — para la métrica (eje Y)
  S.numCols = numColsFromServer.length ? numColsFromServer
    : cols.filter(c=>!c.startsWith("_") && colInfo?.[c]?.type==="number")

  // ── "Agrupar por" → SOLO columnas categóricas ──
  const bySel=$("v-by")
  const prevBy = bySel.value
  // También agregar columnas de tiempo útiles si existen
  const timeCols = ["_mes","_fecha_str","_sem","_anio"].filter(c=>allCols.includes(c) || (filtrosDisp?.all_cols||[]).includes(c))
  // Reconstruir lista: primero las categóricas, luego las temporales
  const byOptions = [...S.catCols, ...timeCols.filter(c=>!S.catCols.includes(c))]
  bySel.innerHTML=`<option value="">— elegí cómo agrupar —</option>`+
    byOptions.map(c=>`<option value="${c}">${colLabel(c,"cat")}</option>`).join("")
  if(prevBy && byOptions.includes(prevBy)) bySel.value = prevBy

  // ── "Métrica" → SOLO columnas numéricas + conteo ──
  const valSel=$("v-val")
  const prevVal = valSel.value
  valSel.innerHTML=`<option value="__count__">📊 Cantidad de registros</option>`+
    S.numCols.map(c=>`<option value="${c}">${colLabel(c,"num")}</option>`).join("")
  if(prevVal && (prevVal==="__count__" || S.numCols.includes(prevVal))) valSel.value = prevVal

  // Refrescar visibilidad de Función según métrica actual
  onVizMetricChange(false)

  // ── Filtros dinámicos: columnas categóricas ──
  document.querySelectorAll(".dyn-filter-col").forEach(sel=>{
    const cur = sel.value
    sel.innerHTML=`<option value="">— columna —</option>`+
      S.catCols.map(c=>`<option value="${c}">${colLabel(c,"cat")}</option>`).join("")
    if(cur && S.catCols.includes(cur)) sel.value = cur
  })

  const addBtn=$("fg-add-filter")
  if(addBtn && S.catCols.length) addBtn.style.display="block"
}

// ── Filtros dinámicos múltiples ───────────────────────────────────────────
let dynFilterId = 0
S.dynFilters = []  // [{id, col, val}]

function addDynamicFilter(){
  if(!S.catCols?.length) return
  const id = ++dynFilterId
  S.dynFilters.push({id, col:"", val:"__all__"})
  renderDynFilters()
}

function removeDynFilter(id){
  S.dynFilters = S.dynFilters.filter(f=>f.id!==id)
  renderDynFilters()
  rf()
}

function renderDynFilters(){
  const cont=$("dynamic-filters-container")
  if(!cont)return
  cont.innerHTML = S.dynFilters.map(f=>`
    <div class="fg" style="position:relative" id="dyn-filter-${f.id}">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:3px">
        <label class="fl" style="margin:0;font-size:10px;opacity:.7">Filtro extra</label>
        <button onclick="removeDynFilter(${f.id})" style="
          background:none;border:none;color:rgba(255,255,255,.4);cursor:pointer;
          font-size:12px;padding:0 2px;line-height:1
        " title="Quitar filtro">✕</button>
      </div>
      <select class="dyn-filter-col" onchange="loadDynVals(${f.id},this)" style="margin-bottom:3px">
        <option value="">— columna —</option>
        ${(S.catCols||[]).map(c=>`<option value="${c}" ${c===f.col?"selected":""}>${c}</option>`).join("")}
      </select>
      <select class="dyn-filter-val" id="dyn-val-${f.id}" onchange="setDynVal(${f.id},this)">
        <option value="__all__">Todos los valores</option>
        ${(f.vals||[]).map(v=>`<option value="${v}" ${v===f.val?"selected":""}>${v}</option>`).join("")}
      </select>
    </div>
  `).join("")
}

async function loadDynVals(id, colSel){
  const col = colSel.value
  const flt = S.dynFilters.find(f=>f.id===id)
  if(!flt) return
  flt.col = col
  flt.val = "__all__"
  flt.vals = []
  const valSel = $(`dyn-val-${id}`)
  if(valSel) valSel.innerHTML=`<option value="__all__">Cargando...</option>`
  if(!col || !S.sids.length){ renderDynFilters(); return }
  try{
    const r = await fetch("/api/col-values",{
      method:"POST", headers:{"Content-Type":"application/json"},
      body: JSON.stringify({sids:S.sids, col, filtros:{}})
    })
    const d = await apiJson(r)
    if(d.ok){ flt.vals = d.values }
  }catch(e){}
  renderDynFilters()
}

function setDynVal(id, valSel){
  const flt = S.dynFilters.find(f=>f.id===id)
  if(flt){ flt.val = valSel.value }
  rf()
}

async function loadCustomVals(){
  // Legacy — no-op (replaced by dynamic filters)
}

// ── Filtros ────────────────────────────────────────────────────────────────
function popF(f, resetDates=false){
  function fill(id,vals,all="Todos"){
    const s=$(id),cv=s.value
    s.innerHTML=`<option value="__all__">${all}</option>`+vals.map(v=>`<option value="${v}">${v||"(vacío)"}</option>`).join("")
    if(cv && cv!=="__all__"){
      const match=vals.find(v=>String(v).trim()===String(cv).trim())
      if(match!==undefined) s.value=match
      else s.value=cv
    }
  }
  // Canal/fuente
  const fuenteSel=$("ff")
  const fuenteVals=(f.fuente||[])
  fuenteSel.innerHTML=`<option value="__all__">Todos los canales</option>`+
    fuenteVals.map(v=>`<option value="${v}">${v}</option>`).join("")

  fill("fe",f.estado||[])
  fill("fcat",f.categoria||[])
  fill("fprov",f.provincia||[])

  // Filtro de mes — siempre actualizar con los meses disponibles
  const meses = f.meses || []
  const fmes = $("fmes"), cvMes = fmes ? fmes.value : "__all__"
  if(fmes){
    fmes.innerHTML = `<option value="__all__">Todos los meses</option>` +
      meses.map(m => {
        const [y,mo] = m.split("-")
        const label = mo && y ? new Date(+y,+mo-1,1).toLocaleDateString("es-AR",{month:"long",year:"numeric"}) : m
        return `<option value="${m}">${label}</option>`
      }).join("")
    if(meses.includes(cvMes)) fmes.value = cvMes
  }

  // Si llegaron cat_cols/num_cols, actualizar selectores de viz y filtros dinámicos
  if(f.cat_cols?.length || f.num_cols?.length){
    const allCols = f.all_cols || f.cat_cols || []
    updateColSelectors(allCols, null, f)
  }

  // Fechas: solo actualizar valores cuando se carga un archivo (resetDates=true)
  // En refreshes de dashboard (rf), NUNCA tocar los valores que el usuario eligió
  if(f.fecha_min){
    const fdEl=$("fd"), fhEl=$("fh")
    // Siempre actualizar los atributos min/max para que el date-picker tenga el rango correcto
    fdEl.min = f.fecha_min; fdEl.max = f.fecha_max
    fhEl.min = f.fecha_min; fhEl.max = f.fecha_max
    if(resetDates){
      const curDesde = fdEl.value, curHasta = fhEl.value
      if(!curDesde || curDesde > f.fecha_max || curHasta < f.fecha_min){
        // Primera carga o rango completamente fuera → resetear al rango completo
        fdEl.value = f.fecha_min; fhEl.value = f.fecha_max
      } else {
        // Solo expandir si el nuevo archivo tiene datos más allá del rango actual
        if(f.fecha_min < curDesde) fdEl.value = f.fecha_min
        if(f.fecha_max > curHasta) fhEl.value = f.fecha_max
      }
    }
  }
}

// Cuando cambia una fecha manualmente, limpiar el filtro de mes
function onDateChange(){ const fm=$("fmes"); if(fm) fm.value="__all__"; rf() }
// Cuando cambia el mes, setear el rango de fechas del mes seleccionado
function onMesChange(){
  const mes = $("fmes").value
  if(mes && mes !== "__all__"){
    const [y,m] = mes.split("-").map(Number)
    const desde = new Date(y,m-1,1)
    const hasta = new Date(y,m,0)  // último día del mes
    const fmt = d => d.toISOString().split("T")[0]
    $("fd").value = fmt(desde)
    $("fh").value = fmt(hasta)
  } else {
    $("fd").value = $("fd").min || ""
    $("fh").value = $("fh").max || ""
  }
  rf()
}

function gF(){
  // Platform tab tiene prioridad sobre el select de fuente
  const fuente = S.platform !== "__all__" ? S.platform : $("ff").value
  // Filtros dinámicos activos
  const dynamic_filters = (S.dynFilters||[])
    .filter(f=>f.col && f.val && f.val!=="__all__")
    .map(f=>({col:f.col, val:f.val}))
  return{
    fuente:          fuente,
    estado:          $("fe").value,
    categoria:       $("fcat").value,
    provincia:       $("fprov").value,
    mes:             ($("fmes") ? $("fmes").value : "__all__"),
    fecha_desde:     $("fd").value,
    fecha_hasta:     $("fh").value,
    texto:           $("fq").value.trim(),
    dynamic_filters: dynamic_filters,
  }
}
function gViz(){
  const valCol = $("v-val").value
  return {
    by_col:     $("v-by").value,
    val_col:    valCol === "__count__" ? null : valCol,
    agg_fn:     valCol === "__count__" ? "count" : $("v-agg").value,
    chart_type: $("v-type").value || "bar",
    top:        parseInt($("v-top")?.value || "10"),
  }
}

function setVizFn(fn, btn){
  $("v-agg").value = fn
  document.querySelectorAll("[data-fn]").forEach(b=>b.classList.remove("act"))
  btn.classList.add("act")
}

function setVizType(type, btn){
  $("v-type").value = type
  ["vt-bar","vt-line","vt-doughnut"].forEach(id=>{ const el=$(id); if(el) el.classList.remove("act") })
  btn.classList.add("act")
}

function setVizTop(top, btn){
  $("v-top").value = top
  document.querySelectorAll("[data-top]").forEach(b=>b.classList.remove("act"))
  btn.classList.add("act")
}

// ── Hints descriptivos según la columna ──────────────────────────────────
const BY_HINTS = {
  publicacion:"Agrupá resultados por cada producto o publicación",
  estado:"Comparar entre ventas entregadas, canceladas, etc.",
  provincia:"Ver distribución geográfica de tus ventas",
  categoria:"Comparar categorías o tipos de producto",
  forma_envio:"Analizar por tipo de envío (Flex, domicilio, etc.)",
  fuente:"Comparar canal ML vs Tienda Nube",
  _mes:"Ver evolución mensual a lo largo del tiempo",
  _fecha_str:"Ver evolución día a día",
  _sem:"Ver evolución semanal",
}
const NUM_HINTS = {
  ingresos:"Total facturado en $ por cada grupo",
  total_neto:"Ingresos descontando costos y comisiones",
  unidades:"Cantidad de unidades vendidas por grupo",
  costo:"Costos totales (envío + cargos) por grupo",
  descuentos:"Total de descuentos aplicados por grupo",
}

function onVizByChange(){
  const by = $("v-by").value
  const hint = $("v-by-hint")
  if(hint) hint.textContent = by ? (BY_HINTS[by] || "") : ""

  if(!by) return
  const byL = by.toLowerCase()
  const valSel = $("v-val")

  // Auto-sugerir tipo de gráfico y métrica según columna elegida
  if(byL.includes("fecha") || byL.includes("_mes") || byL.includes("_sem") || byL.includes("_anio") || byL.includes("sem")){
    setVizType("line", $("vt-line"))
    selectMetric(["ingresos","total_neto","unidades"])
  } else if(byL.includes("pub") || byL.includes("prod") || byL.includes("sku") || byL.includes("articulo")){
    setVizType("bar", $("vt-bar"))
    selectMetric(["ingresos","total_neto","unidades"])
  } else if(byL.includes("estado") || byL.includes("forma") || byL.includes("tipo") || byL.includes("envio") || byL.includes("pais")){
    setVizType("doughnut", $("vt-doughnut"))
    selectMetric(["unidades","__count__"])
  } else if(byL.includes("provincia") || byL.includes("region") || byL.includes("ciudad")){
    setVizType("bar", $("vt-bar"))
    selectMetric(["ingresos","unidades"])
  } else {
    setVizType("bar", $("vt-bar"))
    selectMetric(["ingresos","total_neto","unidades","__count__"])
  }
  onVizMetricChange(false)
}

function selectMetric(preferred){
  const valSel = $("v-val")
  if(!valSel) return
  for(const pref of preferred){
    for(let i=0;i<valSel.options.length;i++){
      if(valSel.options[i].value === pref){ valSel.selectedIndex=i; return }
    }
  }
}

function onVizMetricChange(triggerRefresh=false){
  const val = $("v-val")?.value
  const fnGroup = $("viz-fn-group")
  const hint = $("v-val-hint")
  const stepNum = $("viz-step-chart-num")
  const isCount = !val || val === "__count__"

  // Mostrar/ocultar sección Función — no aplica para Conteo
  if(fnGroup){
    fnGroup.style.display = isCount ? "none" : "block"
    // Ajustar número del paso "Tipo de gráfico"
    if(stepNum) stepNum.textContent = isCount ? "3" : "4"
  }

  // Hint de métrica
  if(hint) hint.textContent = !val || val==="__count__"
    ? "Contará cuántas filas hay en cada grupo"
    : (NUM_HINTS[val] || "")

  // Si es conteo, forzar agg_fn a "count"
  if(isCount && $("v-agg")) $("v-agg").value = "count"
  // Si es numérica y agg era count, resetear a sum
  if(!isCount && $("v-agg")?.value === "count"){
    $("v-agg").value = "sum"
    document.querySelectorAll("[data-fn]").forEach(b=>{
      b.classList.toggle("act", b.dataset.fn==="sum")
    })
  }

  if(triggerRefresh) {}  // no auto-refresh, esperar Graficar
}

// Mantener compatibilidad con código viejo
function vizAutoVal(){ onVizByChange() }

async function applyViz(){
  const by = $("v-by").value
  const val = $("v-val").value
  const btn = $("viz-apply-btn")
  const status = $("viz-status")
  if(!by){ if(status) status.textContent="⚠️ Elegí una columna para agrupar"; return }
  if(!S.sids.length){ if(status) status.textContent="⚠️ Cargá datos primero"; return }
  if(btn){ btn.textContent="⏳ Graficando..."; btn.disabled=true }
  if(status) status.textContent=""
  // Guardar viz para que rf() lo use
  S.pendingViz = gViz()
  await rf()
  if(btn){ btn.textContent="▶ Graficar"; btn.disabled=false }
}

function clrF(){
  ["ff","fe","fcat","fprov"].forEach(id=>{const s=$(id);if(s)s.value="__all__"})
  const fm=$("fmes"); if(fm) fm.value="__all__"
  $("fd").value=$("fd").min||""; $("fh").value=$("fh").max||""; $("fq").value=""
  S.dynFilters=[]; renderDynFilters()
  S.page=0;rf()
}

function setPlatform(plat, btn){
  S.platform = plat
  S.page = 0
  // Actualizar estilos de botones
  document.querySelectorAll(".ptf-btn").forEach(b=>b.classList.remove("act"))
  btn.classList.add("act")
  // Sincronizar con el select de fuente del sidebar
  const ff=$("ff")
  if(ff) ff.value = plat
  // Actualizar título del dashboard
  const titles={"__all__":"Resumen de ventas","Mercado Libre":"Mercado Libre","Tienda Nube":"Tienda Nube"}
  $("dash-title").textContent = titles[plat] || plat
  rf()
}

function updatePlatformTabs(fuentes, mpf){
  // Mostrar/ocultar tabs según fuentes disponibles
  const ml=$("ptf-ml"), tn=$("ptf-tn")
  const hasMl = fuentes.includes("Mercado Libre")
  const hasTn = fuentes.includes("Tienda Nube")
  if(ml) ml.style.display = hasMl ? "" : "none"
  if(tn) tn.style.display = hasTn ? "" : "none"

  // Actualizar contadores
  const total = Object.values(mpf).reduce((a,v)=>a+v.n_ventas,0)
  const cntAll=$("ptf-cnt-all")
  if(cntAll) cntAll.textContent = total+" ventas"

  if(hasMl && mpf["Mercado Libre"]){
    const c=$("ptf-cnt-ml")
    if(c) c.textContent = mpf["Mercado Libre"].n_ventas+" ventas"
  }
  if(hasTn && mpf["Tienda Nube"]){
    const c=$("ptf-cnt-tn")
    if(c) c.textContent = mpf["Tienda Nube"].n_ventas+" ventas"
  }

  // Si la plataforma activa ya no tiene datos, volver a "todas"
  if(S.platform !== "__all__" && !fuentes.includes(S.platform)){
    S.platform = "__all__"
    document.querySelectorAll(".ptf-btn").forEach(b=>b.classList.remove("act"))
    const allBtn=document.querySelector(".ptf-btn.all")
    if(allBtn) allBtn.classList.add("act")
  }
}
function sp(p,btn){
  S.periodo=p
  document.querySelectorAll(".ptab").forEach(b=>b.classList.remove("act"))
  btn.classList.add("act");rf()
}
function cp(d){
  const ps=Math.max(1,Math.ceil(S._tot/50))
  S.page=Math.max(0,Math.min(ps-1,S.page+d));rf()
}

// ── Refresh ────────────────────────────────────────────────────────────────
async function rf(){
  if(!S.sids.length)return
  sL("Actualizando...")
  try{
    const viz = S.pendingViz || gViz()
    const r=await fetch("/api/dashboard",{
      method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({sids:S.sids,filtros:gF(),periodo:S.periodo,page:S.page,viz})
    })
    const d=await apiJson(r)
    if(!d.ok)throw new Error(d.error)
    popF(d.filtros_disponibles)
    if(d.col_names){updateColSelectors(d.col_names, null, d.filtros_disponibles)}
    try{ rd(d) }catch(re){ console.error("Error al renderizar dashboard:",re); sE("Error al renderizar gráficos: "+re.message) }
  }catch(e){sE(e.message)}
  finally{hL()}
}

// ── Top productos sort ────────────────────────────────────────────────────
S._pubSort = "ingresos"  // "ingresos" | "unidades"

function setPubSort(key){
  S._pubSort = key
  // Actualizar estilos de botones
  const acc = "background:var(--acc);color:#fff;border-color:var(--acc)"
  const off = "background:#fff;color:var(--mu);border-color:var(--br)"
  $("btn-sort-ing").style.cssText = $("btn-sort-ing").style.cssText.replace(/background:[^;]+;color:[^;]+;border-color:[^;]+/,"") + (key==="ingresos"?acc:off)
  $("btn-sort-uds").style.cssText = $("btn-sort-uds").style.cssText.replace(/background:[^;]+;color:[^;]+;border-color:[^;]+/,"") + (key==="unidades"?acc:off)
  // Actualizar encabezados de columna
  const thI = $("th-ing"), thU = $("th-uds")
  if(thI) thI.textContent = key==="ingresos" ? "Ingresos ▼" : "Ingresos"
  if(thU) thU.textContent = key==="unidades" ? "Uds ▼" : "Uds"
  renderPubTable()
}

function renderPubTable(){
  const data = (S._pubData||[]).slice()
  const key = S._pubSort || "ingresos"
  data.sort((a,b)=>(b[key]||0)-(a[key]||0))
  // Recalcular pct relativo al sort key
  const tot = data.reduce((s,r)=>s+(r[key]||0),0)
  $("tb-p").innerHTML = data.map((r,i)=>{
    const pct = tot ? ((r[key]||0)/tot*100).toFixed(1) : 0
    return `<tr>
      <td><div style="display:flex;align-items:center;gap:5px">
        <span style="width:10px;height:10px;border-radius:50%;background:${COLORS[i%COLORS.length]};flex-shrink:0;display:inline-block"></span>
        <span title="${r.label||""}">${(r.label||"—").slice(0,38)}</span>
      </div></td>
      <td class="nr">${fmt(r.ingresos)}</td>
      <td class="nr">${fN(r.unidades||0)}</td>
      <td><div class="pb"><div class="pt"><div class="pf" style="width:${pct}%;background:${COLORS[i%COLORS.length]}"></div></div>
      <span style="font-size:10px;color:var(--mu);min-width:28px;text-align:right">${pct}%</span></div></td>
    </tr>`
  }).join("")
}

// ── Render ─────────────────────────────────────────────────────────────────
function getPeriodoLabel(){
  const mes = $("fmes") ? $("fmes").value : "__all__"
  const desde = $("fd") ? $("fd").value : ""
  const hasta = $("fh") ? $("fh").value : ""
  if(mes && mes !== "__all__"){
    const [y,m] = mes.split("-").map(Number)
    return new Date(y,m-1,1).toLocaleDateString("es-AR",{month:"long",year:"numeric"})
  }
  if(desde && hasta){
    const fmtD = v => { const [y,m,d] = v.split("-"); return `${d}/${m}/${y}` }
    if(desde === hasta) return fmtD(desde)
    return `${fmtD(desde)} al ${fmtD(hasta)}`
  }
  return "Todos los períodos"
}

function rd(d){
  S._lastDashboard = d
  const pdfBtn = $("pdf-btn")
  if(pdfBtn) pdfBtn.disabled = false
  const m=d.metricas
  S._tot=d.tabla?.total||0
  $("nb").textContent=`${m.n_ventas} venta${m.n_ventas!==1?"s":""} · ${d.n_filtradas} filtradas`

  // Mostrar período seleccionado bajo el título
  const dmEl=$("dm")
  if(dmEl) dmEl.textContent = getPeriodoLabel()

  // Actualizar tabs de plataforma
  const fuentes=Object.keys(d.metricas_por_fuente||{})
  try{ updatePlatformTabs(fuentes, d.metricas_por_fuente||{}) }catch(e){ console.error("platformTabs:",e) }

  // KPIs
  const _hasEntregado = (m.ingresos_entregado !== undefined && m.ingresos_entregado > 0) || (m.unidades_entregado !== undefined && m.unidades_entregado > 0)
  $("kpis").innerHTML=`
    <div class="kpi bl">${kIco('money','bl')}<div class="kl">INGRESOS</div><div class="kv">${fmt(m.ingresos)}</div><div class="ks">Neto: ${fmt(m.total_neto)}</div></div>
    ${_hasEntregado ? `<div class="kpi gr">${kIco('check','gr')}<div class="kl">ING. ENTREGADO</div><div class="kv">${fmt(m.ingresos_entregado)}</div><div class="ks">${m.n_entregado} entregados</div></div>` : ""}
    ${_hasEntregado ? `<div class="kpi tl">${kIco('box','tl')}<div class="kl">UNID. ENTREGADAS</div><div class="kv">${fN(m.unidades_entregado)}</div><div class="ks">${m.tasa_ok}% del total</div></div>` : ""}
    <div class="kpi gr">${kIco('box','gr')}<div class="kl">UNIDADES</div><div class="kv">${fN(m.unidades)}</div><div class="ks">${m.n_ventas} órdenes</div></div>
    <div class="kpi or">${kIco('receipt','or')}<div class="kl">TICKET PROM.</div><div class="kv">${fmt(m.ticket_prom)}</div><div class="ks">por orden</div></div>
    <div class="kpi rd">${kIco('truck','rd')}<div class="kl">COSTOS</div><div class="kv">${fmt(m.costo)}</div><div class="ks">envío + cargos</div></div>
    <div class="kpi pu">${kIco('tag','pu')}<div class="kl">DESCUENTOS</div><div class="kv">${fmt(m.descuentos)}</div><div class="ks">aplicados</div></div>
    <div class="kpi tl">${kIco('check','tl')}<div class="kl">% COBRADO/OK</div><div class="kv">${m.tasa_ok}%</div><div class="ks">del total</div></div>`

  // Comparativa por fuente
  const mpf=d.metricas_por_fuente||{}
  const fuenteKeys=Object.keys(mpf)
  const cmpBar=$("cmp-bar")
  if(fuenteKeys.length>1){
    cmpBar.style.display="flex"
    const cls={Mercado_Libre:"ml","Mercado Libre":"ml","Tienda Nube":"tn","Tienda_Nube":"tn"}
    cmpBar.innerHTML=fuenteKeys.map(f=>{
      const mf=mpf[f]
      const c=cls[f]||cls[f.replace(" ","_")]||"gen"
      return`<div class="cmp-card ${c}">
        <div class="cmp-lbl ${c}">${f.toUpperCase()}</div>
        <div class="cmp-row"><span class="cmp-k">Ingresos</span><span class="cmp-v">${fmt(mf.ingresos)}</span></div>
        <div class="cmp-row"><span class="cmp-k">Unidades</span><span class="cmp-v">${fN(mf.unidades)}</span></div>
        <div class="cmp-row"><span class="cmp-k">Ticket</span><span class="cmp-v">${fmt(mf.ticket_prom)}</span></div>
        <div class="cmp-row"><span class="cmp-k">Órdenes</span><span class="cmp-v">${mf.n_ventas}</span></div>
        <div class="cmp-row"><span class="cmp-k">% OK</span><span class="cmp-v">${mf.tasa_ok}%</span></div>
      </div>`
    }).join("")
  } else {
    cmpBar.style.display="none"
  }

  // Gráfico de tiempo: una línea por fuente
  try{ rT(d.por_tiempo_fuentes||{}) }catch(e){ console.error("rT:",e) }

  // Donut por fuente
  try{ rD("c-fu", d.por_fuente, "label", fuenteKeys.map(f=>SRC_COLORS[f]||COLORS[Object.keys(SRC_COLORS).length])) }catch(e){ console.error("c-fu:",e) }

  try{ rD("c-es", d.por_estado, "label", ["#16a34a","#d97706","#2563eb","#dc2626","#db2777","#7c3aed"]) }catch(e){ console.error("c-es:",e) }

  // Categoría
  try{
    if(d.por_categoria?.length){
      rD("c-cat", d.por_categoria, "label", COLORS)
    } else {
      rD("c-cat", d.por_envio, "label", ["#2563eb","#16a34a","#d97706","#7c3aed","#dc2626"])
    }
  }catch(e){ console.error("c-cat:",e) }

  // Top provincias
  S._provData = d.por_provincia || []
  try{ renderProvChart() }catch(e){ console.error("provChart:",e) }

  // Flex por zona
  S._flexData = d.flex_zonas || null
  try{ renderFlexChart() }catch(e){ console.error("flexChart:",e) }

  // Top publicaciones
  S._pubData = d.por_publicacion || []
  try{ renderPubTable() }catch(e){ console.error("pubTable:",e) }

  // Viz custom
  const viz = S.pendingViz || gViz()
  S.pendingViz = null  // reset pending
  const vb = viz.by_col || $("v-by").value
  const vv = viz.val_col
  const va = viz.agg_fn || $("v-agg").value
  const vt = viz.chart_type || $("v-type")?.value || "bar"
  const isCount = va === "count" || $("v-val")?.value === "__count__"
  if(d.viz_custom?.length && vb){
    const fnLabels={"sum":"Suma de","count":"Conteo","avg":"Promedio de","max":"Máximo de","min":"Mínimo de"}
    const metricLabel = isCount ? "registros" : (vv || "valor")
    const fnLabel = fnLabels[va] || "Valor de"
    $("viz-custom-title").textContent=`${fnLabel} ${metricLabel}`
    $("viz-custom-sub").textContent=`agrupado por ${vb} · ${d.viz_custom.length} grupos`
    rCustom("c-viz", d.viz_custom, vt)
    const statusEl=$("viz-status")
    if(statusEl) statusEl.textContent=`✓ ${d.viz_custom.length} grupos · ${d.n_filtradas} registros`
  } else if(vb){
    dc("c-viz")
    $("viz-custom-title").textContent="Sin resultados"
    $("viz-custom-sub").textContent="No hay datos para esa combinación con los filtros actuales"
  } else {
    dc("c-viz")
    $("viz-custom-title").textContent="Visualización personalizada"
    $("viz-custom-sub").textContent="Elegí una columna para agrupar y presioná Graficar"
  }

  rTab(d.tabla)
}

function rT(timeBySource){
  dc("c-t")
  const allPeriods=[...new Set(Object.values(timeBySource).flat().map(r=>r.periodo))].filter(Boolean).sort()
  if(!allPeriods.length)return
  const datasets=Object.entries(timeBySource).map(([fname,rows],i)=>{
    const color=SRC_COLORS[fname]||COLORS[i]
    const getV=k=>rows.find(r=>r.periodo===k)?.ingresos||0
    return{label:fname,data:allPeriods.map(k=>getV(k)),
      borderColor:color,backgroundColor:color+"18",fill:true,tension:0.4,pointRadius:3,borderWidth:2}
  })
  if(!datasets.length)return
  S.charts["c-t"]=new Chart($("c-t"),{
    type:"line",
    data:{labels:allPeriods,datasets},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{labels:{font:{size:11},boxWidth:10}},datalabels:{display:false}},
      scales:{
        x:{ticks:{font:{size:10},maxRotation:30},grid:{display:false}},
        y:{ticks:{font:{size:10},callback:v=>fN(v)},grid:{color:"rgba(0,0,0,.04)"}}
      }
    }
  })
}

function rD(cid,rows,lk,cols){
  dc(cid)
  if(!rows?.length)return
  S.charts[cid]=new Chart($(cid),{
    type:"doughnut",
    data:{labels:rows.map(r=>r[lk]||"(vacío)"),
      datasets:[{data:rows.map(r=>r.ingresos||0),backgroundColor:cols||COLORS,borderColor:"#fff",borderWidth:2,hoverOffset:5}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:"60%",
      plugins:{
        legend:{position:"right",labels:{font:{size:10},boxWidth:10,padding:7,
          generateLabels(ch){
            const ds=ch.data,tot=ds.datasets[0].data.reduce((a,b)=>a+b,0)
            return ds.labels.map((l,i)=>({
              text:`${l.length>16?l.slice(0,16)+"…":l} ${tot?((ds.datasets[0].data[i]/tot)*100).toFixed(1)+"%":""}`,
              fillStyle:ds.datasets[0].backgroundColor[i],fontColor:"#475569",hidden:false,index:i
            }))
          }
        }},
        datalabels:{display:false}
      }
    }
  })
}

// ── Envíos Flex por zona ──────────────────────────────────────────────────
S._flexData  = null
S._flexSort  = "cantidad"

function setFlexSort(key){
  S._flexSort = key
  const bC=$("btn-flex-cant"), bI=$("btn-flex-ing")
  if(bC){bC.style.background=key==="cantidad"?"var(--acc)":"#fff";bC.style.color=key==="cantidad"?"#fff":"var(--mu)";bC.style.borderColor=key==="cantidad"?"var(--acc)":"var(--br)"}
  if(bI){bI.style.background=key==="ingresos"?"var(--acc)":"#fff";bI.style.color=key==="ingresos"?"#fff":"var(--mu)";bI.style.borderColor=key==="ingresos"?"var(--acc)":"var(--br)"}
  renderFlexChart()
}

function renderFlexChart(){
  const fz = S._flexData
  if(!fz || !fz.zonas?.length){ $("flex-card").style.display="none"; return }
  $("flex-card").style.display = ""

  const key      = S._flexSort || "cantidad"
  const rows     = [...fz.zonas].sort((a,b)=>b[key]-a[key])
  const totFlex  = fz.total_flex  || 0
  const totEnv   = fz.total_envios || 1
  const pctFlex  = fz.pct_flex_total || 0
  const totIng   = rows.reduce((s,r)=>s+r.ingresos, 0)
  const maxVal   = rows[0]?.[key] || 1   // máximo real para escala del heatmap

  // ── Subtitle ───────────────────────────────────────────────────────────
  const sub = $("flex-subtitle")
  if(sub) sub.textContent = `${totFlex.toLocaleString("es-AR")} Flex · ${pctFlex}% del total de envíos`

  // ── KPIs ───────────────────────────────────────────────────────────────
  const kpiEl = $("flex-kpis")
  if(kpiEl){
    const pctColor = pctFlex>=30?"#16a34a":pctFlex>=10?"#d97706":"#64748b"
    kpiEl.innerHTML = `
      <div style="flex:1;padding:12px 16px;border-right:1px solid var(--br)">
        <div style="font-size:10px;color:var(--mu);font-weight:600;text-transform:uppercase;letter-spacing:.4px">Total Flex</div>
        <div style="font-size:24px;font-weight:800;color:#0d9488">${totFlex.toLocaleString("es-AR")}</div>
        <div style="font-size:11px;color:var(--mu)">envíos Flex</div>
      </div>
      <div style="flex:1;padding:12px 16px;border-right:1px solid var(--br)">
        <div style="font-size:10px;color:var(--mu);font-weight:600;text-transform:uppercase;letter-spacing:.4px">Incidencia</div>
        <div style="font-size:24px;font-weight:800;color:${pctColor}">${pctFlex}%</div>
        <div style="font-size:11px;color:var(--mu)">del total de envíos</div>
      </div>
      <div style="flex:1;padding:12px 16px;border-right:1px solid var(--br)">
        <div style="font-size:10px;color:var(--mu);font-weight:600;text-transform:uppercase;letter-spacing:.4px">Facturación Flex</div>
        <div style="font-size:24px;font-weight:800;color:#2563eb">$ ${Math.round(totIng).toLocaleString("es-AR")}</div>
        <div style="font-size:11px;color:var(--mu)">en ventas Flex</div>
      </div>
      <div style="flex:1;padding:12px 16px">
        <div style="font-size:10px;color:var(--mu);font-weight:600;text-transform:uppercase;letter-spacing:.4px">Zonas</div>
        <div style="font-size:24px;font-weight:800;color:#7c3aed">${rows.length}</div>
        <div style="font-size:11px;color:var(--mu)">zonas con Flex</div>
      </div>`
  }

  // ── Gráfico de barras (Chart.js) ───────────────────────────────────────
  dc("c-flex-bar")
  const barRows  = rows.slice(0,12)   // top 12 para que entre bien
  const barVals  = barRows.map(r => key==="cantidad" ? r.cantidad : Math.round(r.ingresos))
  const pctVals  = barRows.map(r => key==="cantidad" ? r.pct_cantidad : r.pct_ingresos)
  const barLabel = key==="cantidad" ? "Envíos" : "Facturación ($)"
  // Paleta de verde teal degradado según valor
  const barColors = barVals.map((v,i) => {
    const intensity = Math.round(40 + (v/maxVal)*215)  // 40-255
    return `rgba(13,148,136,${(0.35 + (v/maxVal)*0.65).toFixed(2)})`
  })
  const rev = [...barRows].reverse()
  const revVals = [...barVals].reverse()
  const revPct  = [...pctVals].reverse()
  const revColors = [...barColors].reverse()

  S.charts["c-flex-bar"] = new Chart($("c-flex-bar"), {
    type: "bar",
    data: {
      labels: rev.map(r => r.zona.length > 22 ? r.zona.slice(0,22)+"…" : r.zona),
      datasets: [{
        label: barLabel,
        data: revVals,
        backgroundColor: revColors,
        borderRadius: 4,
        borderSkipped: "left",
      }]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        datalabels: {
          display: true,
          anchor: "end",
          align: "end",
          font: { size: 10, weight: "700" },
          color: "#475569",
          formatter: (v, ctx) => {
            const pct = revPct[ctx.dataIndex]
            return key==="cantidad"
              ? `${v} (${pct}%)`
              : `${pct}%`
          }
        },
        tooltip: {
          callbacks: {
            label: ctx => {
              const r = rev[ctx.dataIndex]
              const v = revVals[ctx.dataIndex]
              const pct = revPct[ctx.dataIndex]
              if(key==="cantidad") return ` ${v} envíos · ${pct}% del total Flex`
              return ` $ ${v.toLocaleString("es-AR")} · ${pct}% del total`
            }
          }
        }
      },
      scales: {
        x: {
          ticks: {
            font: { size: 9 },
            // Forzar enteros cuando mostramos cantidad de envíos
            stepSize: key==="cantidad" ? 1 : undefined,
            precision: key==="cantidad" ? 0 : undefined,
            callback: v => {
              if(key==="cantidad"){
                // Solo mostrar si es entero
                return Number.isInteger(v) ? v : null
              }
              return "$ " + v.toLocaleString("es-AR")
            }
          },
          grid: { display: false }
        },
        y: { ticks: { font:{size:10} }, grid: { display: false } }
      }
    }
  })

  // ── Mapa de calor Leaflet ──────────────────────────────────────────────
  const mapEl = $("flex-map")
  if(!mapEl) return

  // Filtrar solo zonas con coordenadas
  const zonasCon = fz.zonas.filter(r => r.lat && r.lng)

  if(zonasCon.length === 0){
    mapEl.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;height:100%;color:var(--mu);font-size:13px;background:#f8fafc;border-radius:10px">
      <div style="text-align:center">
        <div style="font-size:28px;margin-bottom:8px">📍</div>
        <div>Sin coordenadas para las zonas encontradas</div>
        <div style="font-size:11px;margin-top:4px">Cargá datos con columna Ciudad para ver el mapa</div>
      </div>
    </div>`
    return
  }

  // Destruir mapa previo si existe
  if(S._flexMap){ S._flexMap.remove(); S._flexMap = null }

  // Centro del mapa: promedio ponderado de coordenadas
  const totCant = zonasCon.reduce((s,r)=>s+r.cantidad,0)
  const centerLat = zonasCon.reduce((s,r)=>s+r.lat*r.cantidad,0)/totCant
  const centerLng = zonasCon.reduce((s,r)=>s+r.lng*r.cantidad,0)/totCant

  // Crear mapa Leaflet
  const map = L.map("flex-map", {zoomControl:true, scrollWheelZoom:true})
    .setView([centerLat, centerLng], 10)
  S._flexMap = map

  // Tiles OpenStreetMap
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: "© OpenStreetMap",
    maxZoom: 18
  }).addTo(map)

  // Datos del heatmap: [lat, lng, intensity]
  const maxCant = Math.max(...zonasCon.map(r=>r.cantidad))
  const heatData = zonasCon.map(r => [r.lat, r.lng, r.cantidad/maxCant])

  // Capa de heatmap
  L.heatLayer(heatData, {
    radius: 35,
    blur: 25,
    maxZoom: 14,
    max: 1.0,
    gradient: {0.0:"blue", 0.3:"cyan", 0.5:"lime", 0.7:"yellow", 1.0:"red"}
  }).addTo(map)

  // Círculos con tooltip para cada zona
  zonasCon.forEach(r => {
    const size = Math.max(6, Math.round(8 + (r.cantidad/maxCant)*20))
    L.circleMarker([r.lat, r.lng], {
      radius: size,
      fillColor: "#0d9488",
      color: "#fff",
      weight: 1.5,
      opacity: 0.9,
      fillOpacity: 0.7
    })
    .bindTooltip(`<strong>${r.zona}</strong><br>${r.cantidad} envíos (${r.pct_cantidad}% del Flex)<br>$ ${Math.round(r.ingresos).toLocaleString("es-AR")}`, {
      direction:"top", sticky:false
    })
    .addTo(map)
  })

  // Ajustar zoom para mostrar todos los puntos
  if(zonasCon.length > 1){
    const bounds = L.latLngBounds(zonasCon.map(r=>[r.lat,r.lng]))
    map.fitBounds(bounds, {padding:[30,30]})
  }
}


// ── Top provincias sort ───────────────────────────────────────────────────
S._provSort = "ingresos"

function setProvSort(key){
  S._provSort = key
  const acc = "var(--acc)", br = "var(--br)"
  const bI = $("btn-prov-ing"), bU = $("btn-prov-uds")
  if(bI){ bI.style.background = key==="ingresos"?"var(--acc)":"#fff"; bI.style.color = key==="ingresos"?"#fff":"var(--mu)"; bI.style.borderColor = key==="ingresos"?"var(--acc)":"var(--br)" }
  if(bU){ bU.style.background = key==="unidades"?"var(--acc)":"#fff"; bU.style.color = key==="unidades"?"#fff":"var(--mu)"; bU.style.borderColor = key==="unidades"?"var(--acc)":"var(--br)" }
  renderProvChart()
}

function renderProvChart(){
  const key = S._provSort || "ingresos"
  const rows = [...(S._provData||[])].sort((a,b)=>(b[key]||0)-(a[key]||0)).slice(0,10).reverse()
  const fmt2 = key==="unidades" ? v=>fN(v) : v=>fmt(v)
  const color = key==="unidades" ? "rgba(13,148,136,.72)" : "rgba(37,99,235,.72)"
  dc("c-pr")
  if(!rows.length) return
  S.charts["c-pr"] = new Chart($("c-pr"), {
    type: "bar",
    data: {
      labels: rows.map(r => r.label||"(vacío)"),
      datasets: [{
        data: rows.map(r => r[key]||0),
        backgroundColor: color,
        borderRadius: 3,
        borderSkipped: "left"
      }]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        datalabels: { display: true, anchor: "end", align: "end", formatter: fmt2, font: { size: 9 }, color: "#475569" }
      },
      scales: {
        x: { ticks: { font: { size: 9 }, callback: fmt2 }, grid: { display: false } },
        y: { ticks: { font: { size: 10 } }, grid: { display: false } }
      }
    }
  })
}

function rB(cid,rows){
  dc(cid)
  if(!rows?.length)return
  rows=[...rows].reverse()
  S.charts[cid]=new Chart($(cid),{
    type:"bar",
    data:{labels:rows.map(r=>r.label||"(vacío)"),
      datasets:[{data:rows.map(r=>r.ingresos||0),backgroundColor:"rgba(37,99,235,.72)",borderRadius:3,borderSkipped:"left"}]},
    options:{indexAxis:"y",responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},
        datalabels:{display:true,anchor:"end",align:"end",formatter:v=>fN(v),font:{size:9},color:"#475569"}},
      scales:{
        x:{ticks:{font:{size:9},callback:v=>fN(v)},grid:{display:false}},
        y:{ticks:{font:{size:10}},grid:{display:false}}
      }
    }
  })
}

function rCustom(cid, rows, chartType="bar"){
  dc(cid)
  if(!rows?.length)return
  const labels = rows.map(r=>r.label||"(vacío)")
  const vals   = rows.map(r=>r.valor||0)
  const isDoughnut = chartType === "doughnut"
  const isLine     = chartType === "line"
  S.charts[cid]=new Chart($(cid),{
    type: chartType,
    data:{
      labels,
      datasets:[{
        label: "Valor",
        data: vals,
        backgroundColor: isDoughnut ? COLORS.map(c=>c+"dd") : isLine ? "rgba(37,99,235,.15)" : COLORS.map(c=>c+"cc"),
        borderColor: isLine ? "rgba(37,99,235,1)" : COLORS,
        borderRadius: isDoughnut||isLine ? 0 : 4,
        fill: isLine,
        tension: isLine ? 0.4 : 0,
        pointRadius: isLine ? 3 : 0,
        borderWidth: isLine ? 2 : 1
      }]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{display: isDoughnut},
        datalabels:{display: isDoughnut, formatter:(v,ctx)=>{
          const tot=ctx.dataset.data.reduce((a,b)=>a+b,0)
          return tot ? (v/tot*100).toFixed(1)+"%" : ""
        }, font:{size:10}, color:"#fff"}
      },
      scales: isDoughnut ? {} : {
        x:{ticks:{font:{size:10},maxRotation:30},grid:{display:false}},
        y:{ticks:{font:{size:10},callback:v=>fN(v)},grid:{color:"rgba(0,0,0,.04)"}}
      }
    }
  })
}

// ---- Reporte PDF del dashboard actual ------------------------------------
function escHtml(v){
  return String(v ?? "").replace(/[&<>"']/g, ch => ({
    "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;"
  }[ch]))
}

function reportVal(v, kind="text"){
  if(v == null || v === "" || v === "null") return "-"
  if(kind === "money") return fmt(Number(v))
  if(kind === "num") return fN(Number(v))
  if(kind === "pct") return `${Number(v).toLocaleString("es-AR")}%`
  return escHtml(v)
}

function chartDataUrl(cid){
  try{
    const canvas = $(cid)
    return canvas ? canvas.toDataURL("image/png", 1) : ""
  }catch(e){ return "" }
}

function reportTable(title, subtitle, rows, cols, chartId=""){
  rows = rows || []
  const img = chartId ? chartDataUrl(chartId) : ""
  const body = rows.length
    ? `<table><thead><tr>${cols.map(c=>`<th>${escHtml(c.label)}</th>`).join("")}</tr></thead>
       <tbody>${rows.map(r=>`<tr>${cols.map(c=>{
          const cls = c.kind==="money" || c.kind==="num" || c.kind==="pct" ? " class='nr'" : ""
          return `<td${cls}>${reportVal(c.get ? c.get(r) : r[c.key], c.kind)}</td>`
       }).join("")}</tr>`).join("")}</tbody></table>`
    : `<div class="empty">Sin datos para esta seccion.</div>`
  return `<section>
    <div class="section-head">
      <div>
        <h2>${escHtml(title)}</h2>
        ${subtitle ? `<p>${escHtml(subtitle)}</p>` : ""}
      </div>
    </div>
    ${img ? `<img class="chart-img" src="${img}" alt="${escHtml(title)}">` : ""}
    ${body}
  </section>`
}

function buildDashboardReportHTML(d){
  const m = d.metricas || {}
  const today = new Date().toLocaleDateString("es-AR")
  const title = $("dash-title")?.textContent || "Resumen de ventas"
  const period = getPeriodoLabel()
  const source = S.platform === "__all__" ? "Todas las plataformas" : S.platform
  const filters = gF()
  const activeFilters = [
    ["Fuente", source],
    ["Estado", filters.estado],
    ["Categoria", filters.categoria],
    ["Provincia", filters.provincia],
    ["Mes", filters.mes],
    ["Desde", filters.fecha_desde],
    ["Hasta", filters.fecha_hasta],
    ["Texto", filters.texto],
    ...((filters.dynamic_filters||[]).map(f=>[f.col, f.val]))
  ].filter(([_,v]) => v && v !== "__all__")

  const kpis = [
    ["Ingresos", fmt(m.ingresos)],
    ["Neto", fmt(m.total_neto)],
    ["Unidades", fN(m.unidades)],
    ["Ordenes", fN(m.n_ventas)],
    ["Ticket promedio", fmt(m.ticket_prom)],
    ["Costos", fmt(m.costo)],
    ["Descuentos", fmt(m.descuentos)],
    ["% cobrado/OK", `${m.tasa_ok ?? 0}%`],
  ]

  const basicCols = [
    {label:"Concepto", key:"label"},
    {label:"Ingresos", key:"ingresos", kind:"money"},
    {label:"Unidades", key:"unidades", kind:"num"},
    {label:"Ordenes", key:"n_ventas", kind:"num"},
  ]
  const timeCols = [
    {label:"Periodo", key:"periodo"},
    {label:"Ingresos", key:"ingresos", kind:"money"},
    {label:"Unidades", key:"unidades", kind:"num"},
    {label:"Ordenes", key:"n_ventas", kind:"num"},
  ]
  const tableCols = (d.tabla?.cols||[]).slice(0,10).map(c=>({
    label:c.replace(/_/g," "),
    key:c,
    kind:["ingresos","total_neto","costo"].includes(c) ? "money" : c==="unidades" ? "num" : "text"
  }))

  const timeRows = Object.entries(d.por_tiempo_fuentes||{}).flatMap(([fuente, rows]) =>
    (rows||[]).map(r=>({...r, periodo:`${r.periodo} - ${fuente}`}))
  )

  const sections = [
    reportTable("Por canal", "% de ingresos por fuente", d.por_fuente, basicCols, "c-fu"),
    reportTable("Estado de ventas", "Ingresos, unidades y ordenes por estado", d.por_estado, basicCols, "c-es"),
    reportTable("Categoria / Tipo", "Distribucion por ingresos", d.por_categoria?.length ? d.por_categoria : d.por_envio, basicCols, "c-cat"),
    reportTable("Ventas en el tiempo", "Evolucion segun el periodo seleccionado", timeRows, timeCols, "c-t"),
    reportTable("Regiones / Provincias", "Top provincias por facturacion", d.por_provincia, basicCols, "c-pr"),
    d.flex_zonas?.zonas?.length ? reportTable("Envios Flex por zona", `${d.flex_zonas.total_flex} envios Flex`, d.flex_zonas.zonas, [
      {label:"Zona", key:"zona"},
      {label:"Envios", key:"cantidad", kind:"num"},
      {label:"% envios", key:"pct_cantidad", kind:"pct"},
      {label:"Ingresos", key:"ingresos", kind:"money"},
      {label:"% ingresos", key:"pct_ingresos", kind:"pct"},
    ], "c-flex-bar") : "",
    reportTable("Top publicaciones", "Productos/publicaciones con mayor peso en el dashboard", d.por_publicacion, basicCols),
    d.viz_custom?.length ? reportTable("Visualizacion personalizada", $("viz-custom-sub")?.textContent || "", d.viz_custom, basicCols, "c-viz") : "",
    reportTable("Detalle de ventas", `Pagina ${S.page+1} del dashboard - ${d.tabla?.total || 0} filas totales`, d.tabla?.rows, tableCols),
  ].filter(Boolean).join("")

  return `<!doctype html><html><head><meta charset="utf-8"><title>${escHtml(title)} - PDF</title>
  <style>
    @page{size:A4;margin:14mm}
    *{box-sizing:border-box}body{font-family:Inter,Arial,sans-serif;color:#0f172a;margin:0;background:#fff;font-size:11px}
    header{border-bottom:2px solid #2563eb;padding-bottom:14px;margin-bottom:18px}
    h1{font-size:24px;margin:0 0 5px;letter-spacing:-.02em}h2{font-size:15px;margin:0;color:#0f172a}
    p{margin:3px 0 0;color:#64748b}.meta{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}
    .pill{border:1px solid #dbeafe;background:#eff6ff;color:#1d4ed8;border-radius:999px;padding:4px 9px;font-weight:700}
    .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:12px 0 18px}
    .kpi{border:1px solid #e2e8f0;border-radius:8px;padding:9px 10px;background:#f8fafc}
    .kpi b{display:block;font-size:9px;text-transform:uppercase;letter-spacing:.06em;color:#64748b;margin-bottom:4px}
    .kpi span{font-size:16px;font-weight:800}
    section{break-inside:avoid;margin:0 0 18px;padding-top:2px}.section-head{display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:8px}
    table{width:100%;border-collapse:collapse;border:1px solid #e2e8f0;table-layout:fixed}
    th{background:#f1f5f9;color:#475569;text-align:left;font-size:9px;text-transform:uppercase;letter-spacing:.05em;padding:7px;border-bottom:1px solid #e2e8f0}
    td{padding:6px 7px;border-bottom:1px solid #eef2f7;vertical-align:top;word-break:break-word}
    td.nr,th.nr{text-align:right;font-variant-numeric:tabular-nums}.chart-img{display:block;max-width:100%;height:190px;object-fit:contain;margin:4px auto 10px;border:1px solid #e2e8f0;border-radius:8px}
    .filters{margin-top:8px;color:#475569}.empty{border:1px dashed #cbd5e1;border-radius:8px;padding:12px;color:#64748b;background:#f8fafc}
    @media print{button{display:none}}
  </style></head><body>
  <header>
    <h1>${escHtml(title)}</h1>
    <p>${escHtml(period)} - ${escHtml(source)} - Generado el ${escHtml(today)}</p>
    <div class="meta"><span class="pill">${fN(d.n_filtradas || 0)} ventas filtradas</span><span class="pill">${fN(d.n_total || 0)} ventas totales</span></div>
    ${activeFilters.length ? `<div class="filters"><strong>Filtros:</strong> ${activeFilters.map(([k,v])=>`${escHtml(k)}: ${escHtml(v)}`).join(" · ")}</div>` : ""}
  </header>
  <div class="kpis">${kpis.map(([k,v])=>`<div class="kpi"><b>${escHtml(k)}</b><span>${escHtml(v)}</span></div>`).join("")}</div>
  ${sections}
  <script>window.onload=function(){setTimeout(function(){window.print()},450)}<\/script>
  </body></html>`
}

function downloadDashboardPDF(){
  const d = S._lastDashboard
  if(!d){ sE("Todavia no hay un dashboard listo para exportar."); return }
  const w = window.open("", "_blank")
  if(!w){ sE("El navegador bloqueo la ventana de descarga. Habilita pop-ups para guardar el PDF."); return }
  w.document.open()
  w.document.write(buildDashboardReportHTML(d))
  w.document.close()
}

function rTab(t){
  if(!t)return
  const rows=t.rows||[],tot=t.total||0,cols=t.cols||[]
  const ps=Math.max(1,Math.ceil(tot/50))
  S._tot=tot
  $("ts").textContent=`${tot} registros · ${cols.length} columnas`
  $("pi").textContent=`Página ${S.page+1} de ${ps} · ${tot} filas`
  $("pp").disabled=S.page===0;$("pn").disabled=S.page>=ps-1
  // Tabla dinámica con todas las columnas disponibles
  const twMain=$("tw-main")
  twMain.innerHTML=`<table>
    <thead><tr>${cols.map(c=>`<th${["ingresos","total_neto","costo","unidades"].includes(c)?' class="nr"':''}>${
      c.startsWith("_fecha")||c==="_fecha_str"?"Fecha":c.replace(/_/g," ")
    }</th>`).join("")}</tr></thead>
    <tbody>${rows.map(row=>`<tr>${cols.map(c=>{
      const v=row[c]
      const isNum=["ingresos","total_neto","costo","unidades"].includes(c)
      if(c==="estado")return`<td>${etag(v)}</td>`
      if(c==="fuente")return`<td>${srcTag(v)}</td>`
      if(isNum)return`<td class="nr" style="font-weight:${c==="ingresos"?"600":"400"}">${v==null?"—":c==="unidades"?fN(v):fmt(v)}</td>`
      return`<td title="${v||""}" style="max-width:180px">${v==null||v==="null"?"—":String(v).slice(0,40)}</td>`
    }).join("")}</tr>`).join("")}</tbody>
  </table>`
}

// ── Dashi — Chat IA Flotante ──────────────────────────────────────────────
let chatChartCounter = 0
let dashiOpen = false
let chatHistorial = []   // {rol:"user"|"assistant", texto:"..."}

function toggleDashi(){
  dashiOpen = !dashiOpen
  const panel = $("dashi-panel")
  const notif  = $("dashi-notif")
  if(dashiOpen){
    panel.style.display = "flex"
    if(notif) notif.style.display = "none"
    // Mensaje de bienvenida si está vacío
    const msgs = $("chat-msgs")
    if(!msgs.children.length){
      const welcome = document.createElement("div")
      welcome.className = "dmsg-bot"
      const hasDatos = S.sids.length > 0
      welcome.innerHTML = hasDatos
        ? `<strong>¡Hola! Soy Dashi 🐾</strong><br>Cargaste datos y estoy listo para ayudarte. Podés preguntarme sobre tus ventas, comparar meses, ver qué producto te da más margen... ¡lo que quieras!`
        : `<strong>¡Hola! Soy Dashi 🐾</strong><br>Soy tu analista de ventas con IA. Primero cargá tus archivos de Mercado Libre o Tienda Nube y después ¡preguntame lo que quieras sobre tus datos!`
      msgs.appendChild(welcome)
    }
    setTimeout(()=>{ $("chat-input")?.focus() }, 100)
  } else {
    panel.style.display = "none"
  }
}

function clearChat(){
  chatHistorial = []
  $("chat-msgs").innerHTML = ""
  Object.keys(S.charts).filter(k=>k.startsWith("dc-")).forEach(k=>{
    S.charts[k]?.destroy(); delete S.charts[k]
  })
  const welcome = document.createElement("div")
  welcome.className = "dmsg-bot"
  welcome.innerHTML = `<strong>Chat limpio 🐾</strong><br>¿En qué más te puedo ayudar?`
  $("chat-msgs").appendChild(welcome)
}

function askChip(btn){
  // Sacar emoji del chip para la pregunta
  $("chat-input").value = btn.textContent.trim()
  sendChat()
}

async function sendChat(){
  if(!S.sids.length){
    toggleDashi()
    setTimeout(()=>sE("Primero cargá tus archivos de ventas."),100)
    return
  }
  const input = $("chat-input")
  const pregunta = input.value.trim()
  if(!pregunta) return
  input.value = ""
  input.disabled = true

  const msgs = $("chat-msgs")

  // Mensaje usuario
  const uDiv = document.createElement("div")
  uDiv.className = "dmsg-user"
  uDiv.textContent = pregunta
  msgs.appendChild(uDiv)

  // Typing animado
  const thinkDiv = document.createElement("div")
  thinkDiv.className = "dmsg-thinking"
  thinkDiv.innerHTML = "<span></span><span></span><span></span>"
  msgs.appendChild(thinkDiv)
  msgs.scrollTop = msgs.scrollHeight

  // Estado header
  const statusTxt = $("dashi-status-txt")
  if(statusTxt) statusTxt.textContent = "Analizando tus datos..."

  const sendBtn = $("chat-send")
  if(sendBtn){ sendBtn.disabled=true; sendBtn.style.opacity=".5" }

  try{
    const res = await fetch("/api/chat",{
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body: JSON.stringify({sids:S.sids, pregunta, filtros:gF(), historial:chatHistorial})
    })
    const data = await res.json()
    if(!data.ok) throw new Error(data.error)
    // Guardar en historial para contexto
    chatHistorial.push({rol:"user", texto:pregunta})

    const r = data.resultado
    msgs.removeChild(thinkDiv)

    // Burbuja respuesta
    const botDiv = document.createElement("div")
    botDiv.className = "dmsg-bot"

    const txtDiv = document.createElement("div")
    // Formatear: negrita para **texto**
    txtDiv.innerHTML = (r.respuesta||"Sin respuesta.")
      .replace(/\*\*(.+?)\*\*/g,"<strong>$1</strong>")
      .replace(/\n/g,"<br>")
    botDiv.appendChild(txtDiv)

    // Gráfico
    const g = r.grafico
    if(g && g.tipo && g.tipo!=="none" && g.labels?.length && g.values?.length){
      const chartId = "dc-" + (++chatChartCounter)
      const wrap = document.createElement("div")
      wrap.className = "dmsg-chart"
      const canvas = document.createElement("canvas")
      canvas.id = chartId
      wrap.appendChild(canvas)
      botDiv.appendChild(wrap)

      setTimeout(()=>{
        const isDoughnut = g.tipo==="doughnut"
        const isLine     = g.tipo==="line"
        const datasets = [{
          label: g.label_series||"Valor",
          data: g.values,
          backgroundColor: isDoughnut ? COLORS.map(c=>c+"dd") : isLine?"rgba(192,57,43,.12)":(g.color||"rgba(192,57,43,.8)"),
          borderColor: isLine?(g.color||"#c0392b"):COLORS,
          fill:isLine, tension:isLine?.4:0,
          pointRadius:isLine?4:0, borderWidth:isLine?2.5:1,
          borderRadius:isDoughnut||isLine?0:5
        }]
        if(g.values2?.length) datasets.push({
          label:g.label_series2||"Serie 2", data:g.values2, type:"line",
          borderColor:"#16a34a", backgroundColor:"rgba(22,163,74,.08)",
          fill:true,tension:.4,pointRadius:3,borderWidth:2,yAxisID:"y2"
        })
        S.charts[chartId] = new Chart($(chartId),{
          type:g.tipo,
          data:{ labels:g.labels, datasets },
          options:{
            responsive:true,maintainAspectRatio:false,
            plugins:{
              legend:{display:datasets.length>1||isDoughnut,labels:{font:{size:10},boxWidth:10}},
              title:{display:!!g.titulo,text:g.titulo,font:{size:11,weight:"600"},padding:{bottom:5}},
              datalabels:{display:isDoughnut,
                formatter:(v,ctx)=>{const t=ctx.dataset.data.reduce((a,b)=>a+b,0);return t?(v/t*100).toFixed(1)+"%":""},
                font:{size:10},color:"#fff"}
            },
            scales: isDoughnut ? {} : {
              x:{ticks:{font:{size:9},maxRotation:30},grid:{display:false}},
              y:{ticks:{font:{size:9},callback:v=>typeof v==="number"&&v>999?"$"+Math.round(v).toLocaleString("es-AR"):v},grid:{color:"rgba(0,0,0,.04)"}},
              ...(g.values2?.length?{y2:{position:"right",ticks:{font:{size:9}},grid:{display:false}}}:{})
            }
          }
        })
      },60)
    }

    msgs.appendChild(botDiv)
    // Guardar respuesta en historial
    chatHistorial.push({rol:"assistant", texto: r.respuesta||""})
    // Limitar historial a 10 turnos
    if(chatHistorial.length > 20) chatHistorial = chatHistorial.slice(-20)

  }catch(e){
    if(thinkDiv.parentNode) msgs.removeChild(thinkDiv)
    const errDiv = document.createElement("div")
    errDiv.className = "dmsg-bot"
    errDiv.style.cssText += ";border-color:#fecaca;background:#fff5f5"
    const msg = e.message||"Error desconocido"
    const friendly = msg.includes("key")||msg.includes("API key")
      ? "⚙️ Necesitás configurar la API Key. Tocá el botón <strong>⚙️</strong> arriba a la derecha."
      : msg.includes("429") ? "🐾 Demasiadas consultas seguidas, esperá un momento e intentá de nuevo."
      : msg.includes("timeout") ? "⏱️ La consulta tardó demasiado. Intentá de nuevo."
      : `❌ ${msg}`
    errDiv.innerHTML = `<span>${friendly}</span>`
    msgs.appendChild(errDiv)
  }finally{
    input.disabled = false
    if(sendBtn){ sendBtn.disabled=false; sendBtn.style.opacity="1" }
    if(statusTxt) statusTxt.textContent = "Tu analista de ventas"
    msgs.scrollTop = msgs.scrollHeight
    input.focus()
  }
}

// ═══════════════════════════════════════════════════════════════
//  MÓDULO PUBLICIDAD ML ADS
// ═══════════════════════════════════════════════════════════════

let PUB_CHARTS  = {}
let PUB_DATA    = null   // último dataset completo del servidor
let PUB_MES     = '__all__'  // mes seleccionado

// ── nombre legible de un mes YYYY-MM ────────────────────────────
function pubMesLabel(m, corto){
  if(m === '__all__') return corto ? 'Todos' : 'Todos los meses'
  const p = m.split('-')
  const nom = ['','Enero','Febrero','Marzo','Abril','Mayo','Junio',
               'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
  const nomC= ['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
  const n = corto ? nomC : nom
  return (n[parseInt(p[1]||'0')]||m) + ' ' + (p[0]||'').slice(2)
}

function showPubModule(){
  hideMainViews()
  const pub  = document.getElementById('pub-module')
  if(pub)  pub.style.display  = 'block'
  document.querySelectorAll('.tbtn').forEach(b => b.style.outline = 'none')
  const btn = document.getElementById('tbtn-pub')
  if(btn) btn.style.outline = '2px solid #e9d5ff'
  loadPubDash()
}

function trigPub(zone){
  const id = zone === 'camp' ? 'fi-pub-camp' : 'fi-pub-an'
  const el = document.getElementById(id)
  if(el){ el.value=''; el.click() }
}

async function onPubFI(e, tipoHint){
  const files = [...e.target.files]
  e.target.value = ''
  for(const f of files) await uploadPub(f, tipoHint)
}

function pubDov(ev, zone){
  ev.preventDefault(); ev.stopPropagation()
  const el = document.getElementById(zone === 'camp' ? 'pub-dz-camp' : 'pub-dz-an')
  if(el){ el.style.borderColor='#7e22ce'; el.style.background='#f3e8ff' }
}

function pubDrop(ev, tipoHint){
  ev.preventDefault(); ev.stopPropagation()
  document.querySelectorAll('.pub-drop').forEach(el=>{
    el.style.borderColor='#ddd6fe'; el.style.background='#faf5ff'
  })
  ;[...ev.dataTransfer.files].forEach(f => uploadPub(f, tipoHint))
}

async function uploadPub(file, tipoHint){
  const fd = new FormData()
  fd.append('file', file)
  if(tipoHint) fd.append('tipo_hint', tipoHint)
  const lm = document.getElementById('lm')
  if(lm) lm.textContent = 'Procesando ' + file.name + '...'
  const ld = document.getElementById('ld')
  if(ld) ld.classList.add('show')
  try {
    const r = await fetch('/api/publicidad/upload', {method:'POST', body:fd})
    if(ld) ld.classList.remove('show')
    let j
    try { j = await apiJson(r) } catch(e){ showE(e.message || 'Respuesta invalida del servidor'); return }
    if(!j.ok){ showE(j.error || 'Error al cargar el archivo'); return }
    const tipo = j.tipo
    const statusEl = document.getElementById(tipo === 'campanias' ? 'pub-status-camp' : 'pub-status-an')
    const dropEl   = document.getElementById(tipo === 'campanias' ? 'pub-dz-camp'     : 'pub-dz-an')
    if(statusEl) statusEl.innerHTML = `✅ ${j.filas} filas · <b>${j.nombre}</b>`
    if(dropEl){ dropEl.style.borderColor='#16a34a'; dropEl.style.background='#f0fdf4' }
    const sub = document.getElementById('pub-subtitle')
    if(sub) sub.textContent = 'Archivos cargados: ' + j.files.map(x=>x.nombre).join(' · ')
    await loadPubDash()
  } catch(err){
    if(ld) ld.classList.remove('show')
    showE('Error de red: ' + err)
  }
}

async function loadPubDash(){
  try {
    const r = await fetch('/api/publicidad/dashboard')
    const j = await apiJson(r)
    if(!j.ok) return
    PUB_DATA = j
    PUB_MES  = '__all__'
    renderPubDash()
  } catch(e){ /* sin datos todavía */ }
}

// ── Seleccionar mes ──────────────────────────────────────────────
function setPubMes(mes){
  PUB_MES = mes
  // actualizar tabs
  document.querySelectorAll('.pub-mes-tab').forEach(t => {
    const isAct = t.dataset.mes === mes
    t.style.background    = isAct ? '#7e22ce' : '#f3f4f6'
    t.style.color         = isAct ? '#fff'    : '#374151'
    t.style.borderColor   = isAct ? '#7e22ce' : '#e5e7eb'
    t.style.fontWeight    = isAct ? '700' : '500'
  })
  renderPubDash()
}

// ── Formato ──────────────────────────────────────────────────────
function pubFmt(v, tipo){
  if(v === null || v === undefined) return '—'
  tipo = tipo || 'ars'
  if(tipo === 'ars') return '$' + Math.round(v).toLocaleString('es-AR')
  if(tipo === 'pct') return (typeof v === 'number' ? v.toFixed(1) : v) + '%'
  if(tipo === 'x')   return (typeof v === 'number' ? v.toFixed(2) : v) + 'x'
  if(tipo === 'n')   return Math.round(v).toLocaleString('es-AR')
  return v
}

function pubKpiColor(key, val){
  if(key === 'acos'){
    if(val === null) return 'tl'
    return val < 30 ? 'gr' : val < 45 ? 'or' : 'rd'
  }
  if(key === 'roas'){
    if(val === null) return 'tl'
    return val >= 3 ? 'gr' : val >= 1.5 ? 'or' : 'rd'
  }
  if(key === 'ganancia') return val >= 0 ? 'gr' : 'rd'
  const map = {ingresos:'bl', inversion:'or', clics:'tl', impresiones:'pu', ventas:'gr', unidades:'tl'}
  return map[key] || 'bl'
}

// ── Render principal ─────────────────────────────────────────────
function renderPubDash(){
  if(!PUB_DATA) return
  const data    = PUB_DATA
  const mes     = PUB_MES
  const mensual = data.mensual || {}
  const allMeses = Object.keys(mensual).sort()
  document.getElementById('pub-dash').style.display = 'block'

  // ── Banner: período y campañas ───────────────────────────────
  if(allMeses.length){
    const p0 = pubMesLabel(allMeses[0], false)
    const p1 = pubMesLabel(allMeses[allMeses.length-1], false)
    const periEl = document.getElementById('pub-banner-period')
    if(periEl) periEl.textContent = `MercadoLibre · ${p0} – ${p1}`
  }
  const campsEl = document.getElementById('pub-banner-camps')
  if(campsEl){
    const camps = (data.por_campana||[]).map(c=>c.campana)
    campsEl.innerHTML = camps.map(n=>`<span class="pub-camp-pill">${n}</span>`).join('')
  }

  // ── Calcular KPIs según mes ──────────────────────────────────
  let kpi
  if(mes === '__all__'){
    kpi = data.kpis || {}
  } else {
    const md = mensual[mes] || {}
    const ing = md.ingresos||0; const inv = md.inversion||0
    kpi = { ingresos:ing, inversion:inv, ganancia:+(ing-inv).toFixed(2),
             acos:md.acos, roas:md.roas, clics:md.clics||0, ventas:md.ventas||0 }
  }

  // ── Hero KPIs (banner oscuro) ────────────────────────────────
  const heroEl = document.getElementById('pub-hero-kpis')
  if(heroEl){
    const acosBadge = kpi.acos === null ? '' :
      kpi.acos < 30  ? `<span class="pub-hkpi-badge ok">✓ Bajo objetivo</span>` :
      kpi.acos < 40  ? `<span class="pub-hkpi-badge ok">✓ En objetivo</span>` :
      kpi.acos < 60  ? `<span class="pub-hkpi-badge warn">⚠ Sobre 40%</span>` :
                       `<span class="pub-hkpi-badge bad">✗ Muy alto</span>`
    const roasBadge = kpi.roas === null ? '' :
      kpi.roas >= 3   ? `<span class="pub-hkpi-badge ok">✓ Excelente</span>` :
      kpi.roas >= 2   ? `<span class="pub-hkpi-badge ok">✓ En objetivo</span>` :
      kpi.roas >= 1   ? `<span class="pub-hkpi-badge warn">⚠ Bajo 2×</span>` :
                        `<span class="pub-hkpi-badge bad">✗ Déficit</span>`
    const ganBadge = (kpi.ganancia||0) >= 0
      ? `<span class="pub-hkpi-badge ok">✓ Positivo</span>`
      : `<span class="pub-hkpi-badge bad">✗ Déficit</span>`
    heroEl.innerHTML = `
      <div class="pub-hkpi grn">
        <div class="pub-hkpi-lbl">Ingresos por Pub.</div>
        <div class="pub-hkpi-val">${pubFmt(kpi.ingresos,'ars')}</div>
        <div class="pub-hkpi-sub">${mes==='__all__'?allMeses.length+' meses completos':'Mes seleccionado'}</div>
      </div>
      <div class="pub-hkpi orn">
        <div class="pub-hkpi-lbl">Inversión Total</div>
        <div class="pub-hkpi-val">${pubFmt(kpi.inversion,'ars')}</div>
        <div class="pub-hkpi-sub">Total gastado</div>
      </div>
      <div class="pub-hkpi ${(kpi.ganancia||0)>=0?'grn':'red'}">
        <div class="pub-hkpi-lbl">Ganancia Neta</div>
        <div class="pub-hkpi-val">${pubFmt(kpi.ganancia,'ars')}</div>
        <div class="pub-hkpi-sub">Ingresos − Inversión ${ganBadge}</div>
      </div>
      <div class="pub-hkpi ${kpi.acos===null?'pur':kpi.acos<40?'grn':kpi.acos<60?'orn':'red'}">
        <div class="pub-hkpi-lbl">ACOS</div>
        <div class="pub-hkpi-val">${pubFmt(kpi.acos,'pct')}</div>
        <div class="pub-hkpi-sub">Objetivo &lt; 40% ${acosBadge}</div>
      </div>
      <div class="pub-hkpi ${kpi.roas===null?'pur':kpi.roas>=2?'grn':kpi.roas>=1?'orn':'red'}">
        <div class="pub-hkpi-lbl">ROAS</div>
        <div class="pub-hkpi-val">${pubFmt(kpi.roas,'x')}</div>
        <div class="pub-hkpi-sub">Retorno por inversión ${roasBadge}</div>
      </div>
      <div class="pub-hkpi pur">
        <div class="pub-hkpi-lbl">Clics</div>
        <div class="pub-hkpi-val">${pubFmt(kpi.clics,'n')}</div>
        <div class="pub-hkpi-sub">Total período</div>
      </div>
      <div class="pub-hkpi grn">
        <div class="pub-hkpi-lbl">Ventas</div>
        <div class="pub-hkpi-val">${pubFmt(kpi.ventas,'n')}</div>
        <div class="pub-hkpi-sub">Por publicidad</div>
      </div>`
  }

  // ── Mes tabs ─────────────────────────────────────────────────
  const tabsCont = document.getElementById('pub-mes-tabs')
  if(tabsCont && (tabsCont.children.length===0 || tabsCont.dataset.built!==allMeses.join(','))){
    tabsCont.dataset.built = allMeses.join(',')
    tabsCont.innerHTML = ['__all__',...allMeses].map(m=>{
      const lbl = pubMesLabel(m, true)
      return `<button class="pub-mes-tab ${m===mes?'activo':''}" data-mes="${m}" onclick="setPubMes('${m}')">${lbl}</button>`
    }).join('')
  }
  // actualizar activo
  document.querySelectorAll('.pub-mes-tab').forEach(t=>{
    const isAct = t.dataset.mes === mes
    t.classList.toggle('activo', isAct)
  })
  const lblEl = document.getElementById('pub-mes-label')
  if(lblEl) lblEl.textContent = mes==='__all__'
    ? `${allMeses.length} períodos · ${pubMesLabel(allMeses[0],true)} – ${pubMesLabel(allMeses[allMeses.length-1],true)}`
    : pubMesLabel(mes,false) + ' · datos del mes'

  // ── Monthly cards strip ───────────────────────────────────────
  const strip = document.getElementById('pub-meses-strip')
  if(strip){
    strip.style.display = 'flex'
    strip.innerHTML = allMeses.map(m => {
      const md  = mensual[m] || {}
      const ing = md.ingresos||0; const inv = md.inversion||0
      const gan = ing - inv
      const ac  = md.acos; const ro = md.roas
      const isSel = m === mes || mes === '__all__'
      const acCol2 = ac===null?'':ac<30?'ok':ac<40?'ok':ac<60?'warn':'bad'
      const roCol  = ro===null?'':ro>=2?'ok':ro>=1?'warn':'bad'
      const ganCol = gan>=0?'ok':'bad'
      return `<div class="pub-mes-card ${m===mes?'sel':''}" onclick="setPubMes('${m}')">
        <div class="pub-mes-card-head" style="${m===mes?'background:#4c1d95':''}">
          <div class="pub-mes-card-name">${pubMesLabel(m,false)}</div>
        </div>
        <div class="pub-mes-card-body">
          <div class="pub-mc-row"><div class="pub-mc-lbl">Ingresos</div><div class="pub-mc-val">${pubFmt(ing,'ars')}</div></div>
          <div class="pub-mc-row"><div class="pub-mc-lbl">Inversión</div><div class="pub-mc-val">${pubFmt(inv,'ars')}</div></div>
          <div class="pub-mc-row"><div class="pub-mc-lbl">Ganancia</div><div class="pub-mc-val ${ganCol}">${gan>=0?'+':''}${pubFmt(gan,'ars')}</div></div>
          <div class="pub-mc-row"><div class="pub-mc-lbl">ACOS</div><div class="pub-mc-val ${acCol2}">${pubFmt(ac,'pct')}</div></div>
          <div class="pub-mc-row"><div class="pub-mc-lbl">ROAS</div><div class="pub-mc-val ${roCol}">${pubFmt(ro,'x')}</div></div>
          <div class="pub-mc-row"><div class="pub-mc-lbl">Clics</div><div class="pub-mc-val">${pubFmt(md.clics||0,'n')}</div></div>
          <div class="pub-mc-row"><div class="pub-mc-lbl">Ventas</div><div class="pub-mc-val">${pubFmt(md.ventas||0,'n')}</div></div>
        </div>
      </div>`
    }).join('')
  }

  // ── Campañas filtradas ───────────────────────────────────────
  const campsRaw = mes==='__all__' ? (data.por_campana||[]) : ((data.mensual_detalle||{})[mes]||[])
  const camps = campsRaw.slice(0,8)
  const cNames = camps.map(c=>c.campana)

  // ── Chart 1 ──────────────────────────────────────────────────
  if(mes==='__all__'){
    const mesLbl = allMeses.map(m=>pubMesLabel(m,true))
    document.getElementById('pub-ct-ing').textContent='Ingresos vs Inversión'
    document.getElementById('pub-cs-ing').textContent='Por mes · ARS'
    pubChartBar('pub-chart-ing', mesLbl,
      [{label:'Ingresos',data:allMeses.map(m=>mensual[m].ingresos),color:'#7e22ce'},
       {label:'Inversión',data:allMeses.map(m=>mensual[m].inversion),color:'#c4b5fd'}])
    document.getElementById('pub-ct-acos').textContent='ACOS Mensual'
    document.getElementById('pub-cs-acos').textContent='% Inversión / Ingresos · objetivo < 40%'
    pubChartLine('pub-chart-acos', mesLbl, allMeses.map(m=>mensual[m].acos), '#f43f5e', 40)
  } else {
    document.getElementById('pub-ct-ing').textContent='Ingresos vs Inversión · '+pubMesLabel(mes,true)
    document.getElementById('pub-cs-ing').textContent='Por campaña · ARS'
    pubChartBar('pub-chart-ing', cNames,
      [{label:'Ingresos',data:camps.map(c=>c.ingresos),color:'#7e22ce'},
       {label:'Inversión',data:camps.map(c=>c.inversion),color:'#c4b5fd'}])
    document.getElementById('pub-ct-acos').textContent='ACOS · '+pubMesLabel(mes,true)
    document.getElementById('pub-cs-acos').textContent='% por campaña · objetivo < 40%'
    pubChartLine('pub-chart-acos', cNames, camps.map(c=>c.acos), '#f43f5e', 40)
  }

  // ── Campaign cards ───────────────────────────────────────────
  const subtituloCamp = mes==='__all__'?'Período completo':pubMesLabel(mes,false)
  const csEl = document.getElementById('pub-cs-camp'); if(csEl) csEl.textContent=subtituloCamp
  const campCardsEl = document.getElementById('pub-camp-cards')
  if(campCardsEl){
    const maxIng = Math.max(...campsRaw.map(c=>c.ingresos||0), 1)
    campCardsEl.innerHTML = campsRaw.map(c=>{
      const acTag = c.acos===null?'na':c.acos<30?'ok':c.acos<45?'warn':'bad'
      const roTag = c.roas===null?'na':c.roas>=2?'ok':c.roas>=1?'warn':'bad'
      const fillPct = Math.round((c.ingresos||0)/maxIng*100)
      const fillColor = c.roas===null?'#e2e8f0':c.roas>=2?'#10b981':c.roas>=1?'#f59e0b':'#f43f5e'
      const est = (c.estado||'').toLowerCase()
      return `<div class="pub-camp-card">
        <div class="pub-camp-head">
          <span class="pub-camp-name">${c.campana}</span>
          <span class="pub-camp-status ${est}">${c.estado||'—'}</span>
        </div>
        <div class="pub-camp-metrics">
          <div class="pub-camp-m"><div class="pub-camp-ml">Ingresos</div><div class="pub-camp-mv">${pubFmt(c.ingresos,'ars')}</div></div>
          <div class="pub-camp-m"><div class="pub-camp-ml">Inversión</div><div class="pub-camp-mv">${pubFmt(c.inversion,'ars')}</div></div>
          <div class="pub-camp-m"><div class="pub-camp-ml">Ganancia</div><div class="pub-camp-mv ${(c.ganancia||0)>=0?'ok':'bad'}">${(c.ganancia||0)>=0?'+':''}${pubFmt(c.ganancia,'ars')}</div></div>
          <div class="pub-camp-m"><div class="pub-camp-ml">ACOS</div><div class="pub-camp-mv ${acTag}">${pubFmt(c.acos,'pct')}</div></div>
          <div class="pub-camp-m"><div class="pub-camp-ml">ROAS</div><div class="pub-camp-mv ${roTag}">${pubFmt(c.roas,'x')}</div></div>
          <div class="pub-camp-m"><div class="pub-camp-ml">Clics · Ventas</div><div class="pub-camp-mv">${pubFmt(c.clics,'n')} · ${pubFmt(c.ventas,'n')}</div></div>
        </div>
        <div class="pub-camp-bar"><div class="pub-camp-bar-fill" style="width:${fillPct}%;background:${fillColor}"></div></div>
      </div>`
    }).join('')
  }

  // ── Top tabla (campañas / rentabilidad → ahora anuncios filtrados por acos) ─
  const csTablaEl = document.getElementById('pub-cs-tabla-camp')
  if(csTablaEl) csTablaEl.textContent = (mes==='__all__'?'Período completo':'Mes: '+pubMesLabel(mes,false)) + ' · solo productos con ventas'
  const anAll   = mes==='__all__' ? (data.top_anuncios||[]) : ((data.anuncios_por_mes||{})[mes]||[])
  const anConVentas = anAll.filter(a=>a.ingresos>0).sort((a,b)=>(a.acos||999)-(b.acos||999)).slice(0,10)
  const rentRows = anConVentas.length ? anConVentas : campsRaw.filter(c=>(c.ingresos||0)>0).slice(0,10).map(c=>({
    titulo: c.campana,
    ingresos: c.ingresos,
    inversion: c.inversion,
    acos: c.acos,
    roas: c.roas,
    clics: c.clics,
    ventas: c.ventas
  }))
  if(csTablaEl && !anConVentas.length) csTablaEl.textContent = (mes==='__all__'?'Periodo completo':'Mes: '+pubMesLabel(mes,false)) + ' - rendimiento por campania'
  const tbcEl = document.querySelector('#pub-tabla-camp tbody')
  if(tbcEl){
    tbcEl.innerHTML = rentRows.map((a,i)=>{
      const acTag = a.acos===null?'na':a.acos<15?'ok':a.acos<40?'warn':'bad'
      const roTag = a.roas===null?'na':a.roas>=3?'ok':a.roas>=1?'warn':'bad'
      const rk = i===0?'r1':i===1?'r2':i===2?'r3':'rn'
      const acosPct = Math.min(a.acos||0,120)/120*100
      const acosColor = a.acos<15?'#10b981':a.acos<40?'#f59e0b':'#f43f5e'
      return `<tr>
        <td><span class="pub-rank ${rk}">${i+1}</span></td>
        <td style="max-width:260px;overflow:hidden;text-overflow:ellipsis" title="${a.titulo}">${a.titulo}</td>
        <td class="nr" style="color:#10b981;font-weight:700">${pubFmt(a.ingresos,'ars')}</td>
        <td class="nr">${pubFmt(a.inversion,'ars')}</td>
        <td class="nr"><div class="pub-acos-bar"><span class="pub-tag ${acTag}">${pubFmt(a.acos,'pct')}</span><div class="pub-acos-track"><div class="pub-acos-fill" style="width:${acosPct}%;background:${acosColor}"></div></div></div></td>
        <td class="nr"><span class="pub-tag ${roTag}">${pubFmt(a.roas,'x')}</span></td>
        <td class="nr">${pubFmt(a.clics,'n')}</td>
        <td class="nr">${pubFmt(a.ventas,'n')}</td>
      </tr>`
    }).join('') || '<tr><td colspan="8" style="text-align:center;padding:20px;color:#94a3b8;font-size:12px">Sin ventas registradas en este período</td></tr>'
  }

  // ── Top por visibilidad (anuncios ordenados por impresiones) ─
  const anVis = anAll.sort((a,b)=>(b.impresiones||0)-(a.impresiones||0)).slice(0,10)
  const cardAn = document.getElementById('pub-anuncios-card')
  if(cardAn){
    if(anVis.length>0 && anVis[0].impresiones>0){
      cardAn.style.display='block'
      const csAn = document.getElementById('pub-cs-anuncios')
      if(csAn) csAn.textContent = (mes==='__all__'?'Período completo':'Mes: '+pubMesLabel(mes,false)) + ' · ordenado por impresiones'
      const tbaEl = document.querySelector('#pub-tabla-an tbody')
      if(tbaEl){
        tbaEl.innerHTML = anVis.map((a,i)=>{
          const acTag = a.acos===null?'na':a.acos<15?'ok':a.acos<40?'warn':'bad'
          const roTag = a.roas===null?'na':a.roas>=3?'ok':a.roas>=1?'warn':'bad'
          const rk = i===0?'r1':i===1?'r2':i===2?'r3':'rn'
          return `<tr>
            <td><span class="pub-rank ${rk}">${i+1}</span></td>
            <td style="max-width:250px;overflow:hidden;text-overflow:ellipsis" title="${a.titulo}">${a.titulo}</td>
            <td class="nr">${pubFmt(a.impresiones,'n')}</td>
            <td class="nr">${pubFmt(a.clics,'n')}</td>
            <td class="nr" style="color:#10b981;font-weight:700">${pubFmt(a.ingresos,'ars')}</td>
            <td class="nr"><span class="pub-tag ${acTag}">${pubFmt(a.acos,'pct')}</span></td>
            <td class="nr"><span class="pub-tag ${roTag}">${pubFmt(a.roas,'x')}</span></td>
            <td class="nr">${pubFmt(a.ventas,'n')}</td>
          </tr>`
        }).join('')
      }
    } else { cardAn.style.display='none' }
  }
}




function pubChartBar(canvasId, labels, datasets){
  if(PUB_CHARTS[canvasId]){ PUB_CHARTS[canvasId].destroy() }
  const ctx = document.getElementById(canvasId)
  if(!ctx) return
  PUB_CHARTS[canvasId] = new Chart(ctx, {
    type:'bar',
    data:{
      labels,
      datasets: datasets.map(ds => ({
        label: ds.label, data: ds.data,
        backgroundColor: ds.color,
        borderRadius: 5, borderSkipped: false
      }))
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{legend:{position:'top',labels:{font:{size:11}}}, datalabels:{display:false}},
      scales:{
        y:{ticks:{callback:v=>'$'+Math.round(v/1000)+'k',font:{size:10}},grid:{color:'#f1f5f9'}},
        x:{ticks:{font:{size:10}},grid:{display:false}}
      }
    }
  })
}

function pubChartLine(canvasId, labels, data, color, targetLine){
  if(PUB_CHARTS[canvasId]){ PUB_CHARTS[canvasId].destroy() }
  const ctx = document.getElementById(canvasId)
  if(!ctx) return
  const datasets = [{
    label:'ACOS %', data, borderColor:color, backgroundColor:color+'22',
    tension:.35, fill:true, pointRadius:5, pointBackgroundColor:color
  }]
  if(targetLine != null){
    datasets.push({
      label:'Objetivo (40%)', data:labels.map(()=>targetLine),
      borderColor:'#94a3b8', borderDash:[6,4],
      pointRadius:0, fill:false, tension:0
    })
  }
  PUB_CHARTS[canvasId] = new Chart(ctx, {
    type:'line', data:{labels, datasets},
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{legend:{position:'top',labels:{font:{size:11}}}, datalabels:{display:false}},
      scales:{
        y:{ticks:{callback:v=>v+'%',font:{size:10}},grid:{color:'#f1f5f9'}},
        x:{ticks:{font:{size:10}},grid:{display:false}}
      }
    }
  })
}

function pubChartRoas(canvasId, labels, data){
  if(PUB_CHARTS[canvasId]){ PUB_CHARTS[canvasId].destroy() }
  const ctx = document.getElementById(canvasId)
  if(!ctx) return
  PUB_CHARTS[canvasId] = new Chart(ctx, {
    type:'bar',
    data:{
      labels,
      datasets:[{
        label:'ROAS', data: data.map(v=>v||0),
        backgroundColor: data.map(v => v===null?'#e2e8f0':v>=3?'#16a34a':v>=1.5?'#d97706':'#dc2626'),
        borderRadius:5, borderSkipped:false
      },{
        label:'Meta (2×)', data:labels.map(()=>2),
        type:'line', borderColor:'#64748b', borderDash:[6,4],
        pointRadius:0, fill:false, tension:0
      }]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{legend:{position:'top',labels:{font:{size:11}}}, datalabels:{display:false}},
      scales:{
        y:{ticks:{callback:v=>v+'x',font:{size:10}},grid:{color:'#f1f5f9'}},
        x:{ticks:{font:{size:10}},grid:{display:false}}
      }
    }
  })
}



// ═══════════════════════════════════════════════════════════════
//  MÓDULO TENDENCIAS / MERCADO
// ═══════════════════════════════════════════════════════════════

let MKT_MESES  = {}  // { "2026-01": { mes, cats: { cat: rows[] } } }
let MKT_CHARTS = {}

function showMktModule(){
  hideMainViews()
  document.getElementById('mkt-module').style.display='flex'
  document.querySelectorAll('.tbtn').forEach(b=>b.style.outline='none')
  const btn=document.getElementById('tbtn-mkt')
  if(btn) btn.style.outline='2px solid #7dd3fc'
}

function mktDov(e){ e.preventDefault(); document.getElementById('mkt-dz').classList.add('drag') }
function mktDlv(){ document.getElementById('mkt-dz').classList.remove('drag') }
function mktDrop(e){
  e.preventDefault(); mktDlv()
  ;[...e.dataTransfer.files].forEach(f=>uploadMkt(f))
}
async function onMktFI(e){
  const files=[...e.target.files]; e.target.value=''
  for(const f of files) await uploadMkt(f)
}

function mktMesFromFilename(name){
  // benchmark_market-..._2026-01-01a2026-01-31.xlsx → "2026-01"
  const m = name.match(/(\d{4}-\d{2})-\d{2}a/)
  if(m) return m[1]
  const m2 = name.match(/(\d{4})[-_](\d{2})/)
  if(m2) return m2[1]+'-'+m2[2]
  return 'Sin fecha'
}

async function uploadMkt(file){
  sL('Procesando '+ file.name +' ...')
  try{
    const fd=new FormData(); fd.append('file',file)
    const r=await fetch('/api/tendencias/upload',{method:'POST',body:fd})
    const d=await r.json()
    if(!d.ok) throw new Error(d.error)
    // Merge into MKT_MESES
    const mes=d.mes
    if(!MKT_MESES[mes]) MKT_MESES[mes]={}
    Object.assign(MKT_MESES[mes], d.categorias)
    renderMktDash()
    document.getElementById('mkt-dz').classList.add('loaded')
    document.getElementById('mkt-dz').innerHTML=`<div style="font-size:22px">✓</div><div style="font-size:12px;color:var(--tx2);margin-top:6px">${Object.keys(MKT_MESES).length} mes(es) · ${Object.values(MKT_MESES).reduce((s,m)=>s+Object.keys(m).length,0)} categorías</div>`
  }catch(ex){ sE(ex.message||'Error al cargar') }
  finally{ hL() }
}

function mktAllRows(filtMes, filtCat){
  const meses = filtMes==='__all__' ? Object.keys(MKT_MESES) : [filtMes]
  const rows=[]
  for(const mes of meses){
    const cats=MKT_MESES[mes]||{}
    const catKeys=filtCat==='__all__' ? Object.keys(cats) : [filtCat]
    for(const cat of catKeys){
      for(const r of (cats[cat]||[])){
        rows.push({...r, _mes:mes, _cat:cat})
      }
    }
  }
  return rows
}

function mktFiltrar(){
  renderMktDash()
}

function renderMktDash(){
  const mes=document.getElementById('mkt-mes-sel').value||'__all__'
  const cat=document.getElementById('mkt-cat-sel').value||'__all__'
  const rows=mktAllRows(mes,cat)
  if(!rows.length) return

  document.getElementById('mkt-dash').style.display='block'

  // Poblar selects
  const mSel=document.getElementById('mkt-mes-sel')
  const cSel=document.getElementById('mkt-cat-sel')
  const allMeses=[...new Set(Object.keys(MKT_MESES))].sort()
  const allCats=[...new Set(Object.values(MKT_MESES).flatMap(m=>Object.keys(m)))].sort()
  const curMes=mSel.value, curCat=cSel.value
  mSel.innerHTML=`<option value="__all__">Todos los meses (${allMeses.length})</option>`+
    allMeses.map(m=>`<option value="${m}" ${m===curMes?'selected':''}>${mktMesLabel(m)}</option>`).join('')
  cSel.innerHTML=`<option value="__all__">Todas las categorías</option>`+
    allCats.map(c=>`<option value="${c}" ${c===curCat?'selected':''}>${c}</option>`).join('')
  mSel.style.display='inline-block'; cSel.style.display='inline-block'

  // KPIs
  const totalUds=rows.reduce((s,r)=>s+(r.uds||0),0)
  const totalVistas=rows.reduce((s,r)=>s+(r.vistas||0),0)
  const precios=rows.map(r=>r.precio||0).filter(p=>p>0).sort((a,b)=>a-b)
  const mediana=precios.length?precios[Math.floor(precios.length/2)]:0
  const conPub=rows.filter(r=>r.publicidad==='Sí').length
  const conFull=rows.filter(r=>r.envio==='Full').length
  const kc=document.getElementById('mkt-kpis')
  kc.innerHTML=[
    {l:'Productos analizados',v:rows.length.toLocaleString('es-AR'),s:'top 100 por mes/cat'},
    {l:'Unidades vendidas',v:totalUds.toLocaleString('es-AR'),s:'acumulado selección'},
    {l:'Vistas totales',v:(totalVistas/1000).toFixed(0)+'K',s:'tráfico acumulado'},
    {l:'Precio mediano',v:'$'+Math.round(mediana).toLocaleString('es-AR'),s:'ARS'},
    {l:'Con publicidad',v:Math.round(conPub/rows.length*100)+'%',s:Math.round(conFull/rows.length*100)+'% Full'},
  ].map(k=>`<div class="mkt-kpi"><div class="mkt-kpi-l">${k.l}</div><div class="mkt-kpi-v">${k.v}</div><div class="mkt-kpi-s">${k.s}</div></div>`).join('')

  // Ranking
  const sorted=[...rows].sort((a,b)=>(b.uds||0)-(a.uds||0)).slice(0,100)
  document.getElementById('mkt-ranking-sub').textContent=
    (mes==='__all__'?'Todos los meses':'Mes: '+mktMesLabel(mes))+ ' · '+(cat==='__all__'?'Todas las categorías':cat)
  document.getElementById('mkt-tbody-ranking').innerHTML=sorted.map((r,i)=>`<tr>
    <td style="color:var(--tx2);font-weight:600">${i+1}</td>
    <td><span class="mkt-nm" title="${r.nombre}">${r.nombre}</span></td>
    <td><span style="font-size:10px;color:var(--tx2)">${r._cat}</span></td>
    <td>${r.condicion||'—'}</td>
    <td class="nr" style="font-weight:500">$${Math.round(r.precio||0).toLocaleString('es-AR')}</td>
    <td class="nr" style="font-weight:700;color:#0ea5e9">${(r.uds||0).toLocaleString('es-AR')}</td>
    <td class="nr">${((r.vistas||0)/1000).toFixed(0)}K</td>
    <td class="nr">${r.vistas&&r.uds?((r.uds/r.vistas)*100).toFixed(2)+'%':'—'}</td>
    <td><span class="mkt-chip ${r.envio==='Full'?'full':'no'}">${r.envio||'—'}</span></td>
    <td><span class="mkt-chip ${r.publicidad==='Sí'?'si':'no'}">${r.publicidad||'—'}</span></td>
    <td><span class="mkt-chip ${r.cuotas==='Sí'?'si':'no'}">${r.cuotas||'—'}</span></td>
  </tr>`).join('')

  // Benchmark
  const pTop10=rows.slice(0,10)
  const avgUds=rows.length?Math.round(totalUds/rows.length):0
  const pConCat=rows.filter(r=>r.catalogo==='Sí').length
  const pConVid=rows.filter(r=>r.video==='Sí').length
  const avgFotos=rows.length?Math.round(rows.reduce((s,r)=>s+(r.fotos||0),0)/rows.length):0
  const avgPreg=rows.length?Math.round(rows.reduce((s,r)=>s+(r.preguntas||0),0)/rows.length):0
  document.getElementById('mkt-bench-sub').textContent=`${rows.length} productos analizados`
  document.getElementById('mkt-bench-grid').innerHTML=`
    <div class="mkt-bench-card">
      <h4>Ventas</h4>
      <div class="mkt-bench-row"><span>Uds. promedio</span><span class="mkt-bench-val">${avgUds.toLocaleString('es-AR')}</span></div>
      <div class="mkt-bench-row"><span>Uds. máximas</span><span class="mkt-bench-val">${Math.max(...rows.map(r=>r.uds||0)).toLocaleString('es-AR')}</span></div>
      <div class="mkt-bench-row"><span>Uds. top 10</span><span class="mkt-bench-val">${pTop10.reduce((s,r)=>s+(r.uds||0),0).toLocaleString('es-AR')}</span></div>
    </div>
    <div class="mkt-bench-card">
      <h4>Precios</h4>
      <div class="mkt-bench-row"><span>Mínimo</span><span class="mkt-bench-val">$${Math.round(Math.min(...precios)).toLocaleString('es-AR')}</span></div>
      <div class="mkt-bench-row"><span>Mediana</span><span class="mkt-bench-val">$${Math.round(mediana).toLocaleString('es-AR')}</span></div>
      <div class="mkt-bench-row"><span>Máximo</span><span class="mkt-bench-val">$${Math.round(Math.max(...precios)).toLocaleString('es-AR')}</span></div>
    </div>
    <div class="mkt-bench-card">
      <h4>Condiciones</h4>
      <div class="mkt-bench-row"><span>Con publicidad</span><span class="mkt-bench-val">${Math.round(conPub/rows.length*100)}%</span></div>
      <div class="mkt-bench-row"><span>Envío Full</span><span class="mkt-bench-val">${Math.round(conFull/rows.length*100)}%</span></div>
      <div class="mkt-bench-row"><span>Catálogo ML</span><span class="mkt-bench-val">${Math.round(pConCat/rows.length*100)}%</span></div>
    </div>
    <div class="mkt-bench-card">
      <h4>Contenido</h4>
      <div class="mkt-bench-row"><span>Fotos promedio</span><span class="mkt-bench-val">${avgFotos}</span></div>
      <div class="mkt-bench-row"><span>Con video</span><span class="mkt-bench-val">${Math.round(pConVid/rows.length*100)}%</span></div>
      <div class="mkt-bench-row"><span>Preguntas prom.</span><span class="mkt-bench-val">${avgPreg}</span></div>
    </div>
    <div class="mkt-bench-card">
      <h4>Tráfico</h4>
      <div class="mkt-bench-row"><span>Vistas promedio</span><span class="mkt-bench-val">${Math.round(totalVistas/rows.length).toLocaleString('es-AR')}</span></div>
      <div class="mkt-bench-row"><span>Vistas máximas</span><span class="mkt-bench-val">${Math.max(...rows.map(r=>r.vistas||0)).toLocaleString('es-AR')}</span></div>
      <div class="mkt-bench-row"><span>Conv. top (vistas→uds)</span><span class="mkt-bench-val">${rows[0]&&rows[0].vistas?((rows[0].uds/rows[0].vistas)*100).toFixed(2)+'%':'—'}</span></div>
    </div>
    <div class="mkt-bench-card">
      <h4>Cuotas</h4>
      <div class="mkt-bench-row"><span>Ofrecen cuotas</span><span class="mkt-bench-val">${Math.round(rows.filter(r=>r.cuotas==='Sí').length/rows.length*100)}%</span></div>
      <div class="mkt-bench-row"><span>Nuevos</span><span class="mkt-bench-val">${Math.round(rows.filter(r=>r.condicion==='Nuevo').length/rows.length*100)}%</span></div>
      <div class="mkt-bench-row"><span>Usados</span><span class="mkt-bench-val">${Math.round(rows.filter(r=>r.condicion==='Usado').length/rows.length*100)}%</span></div>
    </div>`

  // Evolución
  const mesesKeys=Object.keys(MKT_MESES).sort()
  if(mesesKeys.length>=2){
    document.getElementById('mkt-evol-empty').style.display='none'
    document.getElementById('mkt-evol-charts').style.display='block'
    setTimeout(()=>renderMktEvolCharts(mesesKeys,cat),50)
  }

  // Gráficos si pestaña activa
  const gPanel=document.getElementById('mkt-panel-graficos')
  if(gPanel&&gPanel.classList.contains('on')) setTimeout(()=>renderMktCharts(rows),50)
}

function mktMesLabel(mes){
  if(!mes||mes==='Sin fecha') return mes
  const [y,m]=mes.split('-')
  if(!y||!m) return mes
  return new Date(+y,+m-1,1).toLocaleDateString('es-AR',{month:'long',year:'numeric'})
}

function dcMkt(id){ if(MKT_CHARTS[id]){MKT_CHARTS[id].destroy();delete MKT_CHARTS[id]} }

function renderMktCharts(rows){
  const grd='rgba(0,0,0,0.05)', tx='#64748b'
  const sorted=[...rows].sort((a,b)=>(b.uds||0)-(a.uds||0))

  // Top 10 uds
  dcMkt('mkt-ch-top')
  const top10=sorted.slice(0,10).reverse()
  const c1=document.getElementById('mkt-ch-top')
  if(c1) MKT_CHARTS['mkt-ch-top']=new Chart(c1,{
    type:'bar',
    data:{labels:top10.map(r=>r.nombre.length>30?r.nombre.slice(0,28)+'…':r.nombre),
      datasets:[{label:'Uds. vendidas',data:top10.map(r=>r.uds||0),
        backgroundColor:'#0ea5e9',borderRadius:4,borderSkipped:false}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'right',
        font:{size:10},color:tx,formatter:v=>v.toLocaleString('es-AR'),clip:false}},
      layout:{padding:{right:60}},
      scales:{x:{grid:{color:grd},ticks:{font:{size:10},color:tx},border:{display:false}},
              y:{grid:{display:false},ticks:{font:{size:10},color:tx},border:{display:false}}}}
  })
  document.getElementById('mkt-ch-top-sub').textContent=
    sorted[0]?`Liderado por ${sorted[0].nombre.slice(0,30)} (${(sorted[0].uds||0).toLocaleString('es-AR')} uds.)`:'';

  // Distribución precios
  dcMkt('mkt-ch-precios')
  const precios=rows.map(r=>r.precio||0).filter(p=>p>0)
  const ranges=['<$20k','$20-50k','$50-100k','$100-300k','$300k+']
  const buckets=[0,0,0,0,0]
  precios.forEach(p=>{
    if(p<20000)buckets[0]++
    else if(p<50000)buckets[1]++
    else if(p<100000)buckets[2]++
    else if(p<300000)buckets[3]++
    else buckets[4]++
  })
  const c2=document.getElementById('mkt-ch-precios')
  if(c2) MKT_CHARTS['mkt-ch-precios']=new Chart(c2,{
    type:'bar',
    data:{labels:ranges,datasets:[{label:'Productos',data:buckets,
      backgroundColor:['#6366f1','#0ea5e9','#10b981','#f59e0b','#ef4444'],
      borderRadius:4,borderSkipped:false}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},datalabels:{display:false}},
      scales:{x:{grid:{display:false},ticks:{font:{size:10},color:tx},border:{display:false}},
              y:{grid:{color:grd},ticks:{font:{size:10},color:tx},border:{display:false}}}}
  })

  // Publicidad donut
  dcMkt('mkt-ch-pub')
  const nPub=rows.filter(r=>r.publicidad==='Sí').length
  const c3=document.getElementById('mkt-ch-pub')
  if(c3) MKT_CHARTS['mkt-ch-pub']=new Chart(c3,{
    type:'doughnut',
    data:{labels:['Con publicidad','Sin publicidad'],
      datasets:[{data:[nPub,rows.length-nPub],backgroundColor:['#0ea5e9','#e2e8f0'],borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'60%',
      plugins:{legend:{position:'bottom',labels:{font:{size:10},boxWidth:10}},datalabels:{display:false}}}
  })

  // Envío donut
  dcMkt('mkt-ch-envio')
  const envioCounts={}
  rows.forEach(r=>{const k=r.envio||'Otro';envioCounts[k]=(envioCounts[k]||0)+1})
  const c4=document.getElementById('mkt-ch-envio')
  if(c4) MKT_CHARTS['mkt-ch-envio']=new Chart(c4,{
    type:'doughnut',
    data:{labels:Object.keys(envioCounts),
      datasets:[{data:Object.values(envioCounts),
        backgroundColor:['#0ea5e9','#10b981','#f59e0b','#6366f1','#94a3b8'],borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'60%',
      plugins:{legend:{position:'bottom',labels:{font:{size:10},boxWidth:10}},datalabels:{display:false}}}
  })

  // Catálogo donut
  dcMkt('mkt-ch-cat')
  const nCat=rows.filter(r=>r.catalogo==='Sí').length
  const c5=document.getElementById('mkt-ch-cat')
  if(c5) MKT_CHARTS['mkt-ch-cat']=new Chart(c5,{
    type:'doughnut',
    data:{labels:['Catálogo','No catálogo'],
      datasets:[{data:[nCat,rows.length-nCat],backgroundColor:['#10b981','#e2e8f0'],borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'60%',
      plugins:{legend:{position:'bottom',labels:{font:{size:10},boxWidth:10}},datalabels:{display:false}}}
  })

  // Cuotas donut
  dcMkt('mkt-ch-cuotas')
  const nCuotas=rows.filter(r=>r.cuotas==='Sí').length
  const c6=document.getElementById('mkt-ch-cuotas')
  if(c6) MKT_CHARTS['mkt-ch-cuotas']=new Chart(c6,{
    type:'doughnut',
    data:{labels:['Con cuotas','Sin cuotas'],
      datasets:[{data:[nCuotas,rows.length-nCuotas],backgroundColor:['#f59e0b','#e2e8f0'],borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'60%',
      plugins:{legend:{position:'bottom',labels:{font:{size:10},boxWidth:10}},datalabels:{display:false}}}
  })
}

function renderMktEvolCharts(meses, filtCat){
  const tx='#64748b', grd='rgba(0,0,0,0.05)'
  const labels=meses.map(m=>mktMesLabel(m))

  // Uds por mes (total)
  dcMkt('mkt-ch-evol-uds')
  const udsXmes=meses.map(mes=>{
    const cats=MKT_MESES[mes]||{}
    const catKeys=filtCat==='__all__'?Object.keys(cats):[filtCat]
    return catKeys.reduce((s,cat)=>{
      return s+(cats[cat]||[]).reduce((ss,r)=>ss+(r.uds||0),0)
    },0)
  })
  const c1=document.getElementById('mkt-ch-evol-uds')
  if(c1) MKT_CHARTS['mkt-ch-evol-uds']=new Chart(c1,{
    type:'line',
    data:{labels,datasets:[{label:'Uds. vendidas (top 100)',data:udsXmes,
      borderColor:'#0ea5e9',backgroundColor:'rgba(14,165,233,.1)',
      pointBackgroundColor:'#0ea5e9',pointRadius:5,fill:true,tension:.3}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},datalabels:{display:false}},
      scales:{x:{ticks:{font:{size:10},color:tx},grid:{display:false},border:{display:false}},
              y:{ticks:{font:{size:10},color:tx},grid:{color:grd},border:{display:false}}}}
  })

  // Precio mediano por mes
  dcMkt('mkt-ch-evol-precio')
  const precioXmes=meses.map(mes=>{
    const cats=MKT_MESES[mes]||{}
    const catKeys=filtCat==='__all__'?Object.keys(cats):[filtCat]
    const ps=catKeys.flatMap(cat=>(cats[cat]||[]).map(r=>r.precio||0)).filter(p=>p>0).sort((a,b)=>a-b)
    return ps.length?ps[Math.floor(ps.length/2)]:0
  })
  const c2=document.getElementById('mkt-ch-evol-precio')
  if(c2) MKT_CHARTS['mkt-ch-evol-precio']=new Chart(c2,{
    type:'line',
    data:{labels,datasets:[{label:'Precio mediano (ARS)',data:precioXmes,
      borderColor:'#10b981',backgroundColor:'rgba(16,185,129,.08)',
      pointBackgroundColor:'#10b981',pointRadius:5,fill:true,tension:.3}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},datalabels:{display:false}},
      scales:{x:{ticks:{font:{size:10},color:tx},grid:{display:false},border:{display:false}},
              y:{ticks:{callback:v=>'$'+Math.round(v/1000)+'k',font:{size:10},color:tx},
                 grid:{color:grd},border:{display:false}}}}
  })
}

function mktTab(id, el){
  document.querySelectorAll('.mkt-panel').forEach(p=>p.classList.remove('on'))
  document.querySelectorAll('.mkt-tab').forEach(t=>t.classList.remove('on'))
  document.getElementById('mkt-panel-'+id).classList.add('on')
  if(el) el.classList.add('on')
  if(id==='graficos'){
    const mes=document.getElementById('mkt-mes-sel').value||'__all__'
    const cat=document.getElementById('mkt-cat-sel').value||'__all__'
    setTimeout(()=>renderMktCharts(mktAllRows(mes,cat)),50)
  }
}

// ═══════════════════════════════════════════════════════════════
//  MÓDULO PUBLICACIONES — POCAS VENTAS
// ═══════════════════════════════════════════════════════════════

let LOW_DATA = []   // todos los rows
let LOW_FILT = 'todas'

async function onPubsLowFI(e){
  const f = e.target.files[0]; e.target.value = ''
  if(!f) return
  sL('Analizando publicaciones...')
  try{
    const fd = new FormData(); fd.append('file', f)
    const r = await fetch('/api/publicaciones/upload', {method:'POST', body:fd})
    const d = await apiJson(r)
    if(!d.ok) throw new Error(d.error)
    // Usar la respuesta para poblar la tabla de pocas ventas
    renderLowSales(d)
    // También actualizar el dashboard principal si hay datos
    if(!PUBS_DATA) { PUBS_DATA = d; renderPubsDash(d) }
  }catch(ex){ sE(ex.message||'Error al cargar') }
  finally{ hL() }
}

function renderLowSales(d){
  const allRows = [
    ...(d.top_ventas||[]),
    ...(d.top_visitas||[]),
    ...(d.top_cvr||[]),
  ]
  // Dedup por nombre
  const seen = new Set()
  const rows = []
  for(const r of allRows){
    if(!seen.has(r.nombre)){ seen.add(r.nombre); rows.push(r) }
  }
  // Enriquecer con diagnóstico
  LOW_DATA = rows.map(r => {
    const vis = r.visitas || 0
    const ven = r.ventas || 0
    const cvr = r.conv || 0
    let diag = 'ok', diagLabel = 'Normal'
    if(ven === 0 && vis > 0){ diag = 'bconv'; diagLabel = 'Visitas sin compras' }
    if(ven === 0 && vis === 0){ diag = 'sinvis'; diagLabel = 'Sin actividad' }
    if(ven === 0){ diag = 'novent'; diagLabel = 'Sin ventas' }
    else if(cvr < 1 && vis >= 30){ diag = 'bconv'; diagLabel = 'CVR muy bajo (<1%)' }
    else if(cvr < 2 && vis >= 50){ diag = 'bconv'; diagLabel = 'CVR bajo (<2%)' }
    return {...r, diag, diagLabel}
  })

  // KPIs
  const sinVentas = LOW_DATA.filter(r=>!r.ventas||r.ventas===0).length
  const bajaConv  = LOW_DATA.filter(r=>r.diag==='bconv').length
  const sinActividad = LOW_DATA.filter(r=>r.diag==='sinvis').length
  const kc = document.getElementById('pubs-low-kpis')
  if(kc) kc.innerHTML = [
    {l:'Sin ventas', v:sinVentas, s:'publicaciones sin ninguna venta', cls:''},
    {l:'Baja conversión', v:bajaConv, s:'visitas pero bajo CVR', cls:''},
    {l:'Sin actividad', v:sinActividad, s:'sin visitas ni ventas', cls:''},
  ].map(k=>`<div class="pubs-low-kpi"><div class="lk-l">${k.l}</div><div class="lk-v">${k.v}</div><div class="lk-s">${k.s}</div></div>`).join('')

  document.getElementById('pubs-low-empty').style.display = 'none'
  document.getElementById('pubs-low-data').style.display  = 'block'
  LOW_FILT = 'todas'
  document.querySelectorAll('.pubs-low-fbtn').forEach(b=>b.classList.remove('on'))
  const btn = document.querySelector('.pubs-low-fbtn')
  if(btn) btn.classList.add('on')
  renderLowTable()
}

function filterLow(tipo, el){
  LOW_FILT = tipo
  document.querySelectorAll('.pubs-low-fbtn').forEach(b=>b.classList.remove('on'))
  if(el) el.classList.add('on')
  renderLowTable()
}

function renderLowTable(){
  let rows = [...LOW_DATA]
  if(LOW_FILT === 'sin_ventas') rows = rows.filter(r=>!r.ventas||r.ventas===0)
  else if(LOW_FILT === 'baja_conv') rows = rows.filter(r=>r.diag==='bconv')
  else if(LOW_FILT === 'activas') rows = rows.filter(r=>r.estado==='ACTIVA')
  // Ordenar: sin ventas primero, luego baja conv, luego por visitas desc
  rows.sort((a,b)=>{
    const order = {novent:0, bconv:1, sinvis:2, ok:3}
    if(order[a.diag] !== order[b.diag]) return order[a.diag]-order[b.diag]
    return (b.visitas||0)-(a.visitas||0)
  })

  const tbody = document.getElementById('pubs-low-tbody')
  if(!tbody) return
  if(!rows.length){
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;padding:20px;color:var(--tx2);font-size:12px">No hay publicaciones en esta categoría.</td></tr>`
    return
  }

  const fmtARS = v => v ? '$'+Math.round(v).toLocaleString('es-AR') : '—'
  const fmtN   = v => (v!=null && !isNaN(v)) ? Math.round(v).toLocaleString('es-AR') : '—'
  const fmtPct = v => (v!=null && !isNaN(v)) ? v.toFixed(1)+'%' : '—'

  tbody.innerHTML = rows.map((r,i)=>{
    const badgeClass = r.estado==='ACTIVA' ? 'act' : 'ina'
    const cvrColor = (!r.conv||r.conv<1) ? '#b91c1c' : r.conv<2 ? '#854d0e' : '#166534'
    const diagMap  = {novent:'diag-novent', bconv:'diag-bconv', sinvis:'diag-sinvis', ok:'diag-ok'}
    return `<tr>
      <td style="color:var(--tx2);font-weight:500">${i+1}</td>
      <td><span class="pubs-nm" title="${r.nombre}">${r.nombre}</span></td>
      <td><span class="pubs-badge ${badgeClass}">${r.estado==='ACTIVA'?'Activa':'Inactiva'}</span></td>
      <td class="nr">${fmtN(r.visitas)}</td>
      <td class="nr">${fmtN(r.ventas)}</td>
      <td class="nr" style="font-weight:500;color:${cvrColor}">${fmtPct(r.conv)}</td>
      <td class="nr">${fmtARS(r.ventas_brutas)}</td>
      <td><span class="diag-chip ${diagMap[r.diag]||'diag-ok'}">${r.diagLabel}</span></td>
    </tr>`
  }).join('')
}

// ═══════════════════════════════════════════════════════════════
//  MÓDULO PUBLICACIONES ML
// ═══════════════════════════════════════════════════════════════

let PUBS_CHARTS = {}
let PUBS_DATA   = null

function showPubsModule(){
  hideMainViews()
  const m=document.getElementById('pubs-module')
  if(m) m.style.display='flex'
  document.querySelectorAll('.tbtn').forEach(b=>b.style.outline='none')
  const btn=document.getElementById('tbtn-pubs')
  if(btn) btn.style.outline='2px solid #c4b5fd'
  if(PUBS_DATA) renderPubsDash(PUBS_DATA)
}

function pubsTab(id, el){
  document.querySelectorAll('.pubs-panel').forEach(p=>p.classList.remove('on'))
  document.querySelectorAll('.pubs-tab').forEach(t=>t.classList.remove('on'))
  const panel=document.getElementById('pubs-panel-'+id)
  if(panel) panel.classList.add('on')
  if(el) el.classList.add('on')
  if(id==='graficos' && PUBS_DATA) setTimeout(()=>renderPubsCharts(PUBS_DATA),50)
}

function pubsDov(e){ e.preventDefault(); const dz=document.getElementById('pubs-dz'); if(dz){dz.classList.add('drag')} }
function pubsDlv(){ const dz=document.getElementById('pubs-dz'); if(dz){dz.classList.remove('drag')} }
function pubsDrop(e){ e.preventDefault(); pubsDlv(); const f=[...e.dataTransfer.files][0]; if(f) uploadPubs(f) }

async function onPubsFI(e){
  const f=e.target.files[0]; e.target.value=''
  if(f) await uploadPubs(f)
}

async function uploadPubs(file){
  sL('Procesando publicaciones...')
  try{
    const fd=new FormData(); fd.append('file',file)
    const r=await fetch('/api/publicaciones/upload',{method:'POST',body:fd})
    const d=await apiJson(r)
    if(!d.ok) throw new Error(d.error)
    PUBS_DATA=d
    renderPubsDash(d)
    const dz=document.getElementById('pubs-dz')
    if(dz){ dz.classList.add('loaded') }
    const dzt=document.getElementById('pubs-dz-t')
    const dzs=document.getElementById('pubs-dz-s')
    if(dzt) dzt.textContent='✓ '+file.name
    if(dzs) dzs.textContent=d.total_pubs+' publicaciones cargadas · '+d.periodo
  }catch(e){ sE(e.message||'Error al cargar publicaciones') }
  finally{ hL() }
}

function pFmt(v, tipo){
  if(v===null||v===undefined) return '—'
  if(tipo==='ars') return '$'+Math.round(v).toLocaleString('es-AR')
  if(tipo==='pct') return (typeof v==='number'?v.toFixed(1):v)+'%'
  if(tipo==='n')   return Math.round(v).toLocaleString('es-AR')
  return v
}

function renderPubsDash(d){
  const dash=document.getElementById('pubs-dash')
  if(dash) dash.style.display='block'

  // KPIs
  const kpis=[
    {l:'Publicaciones',v:d.total_pubs,s:(d.activas||0)+' activas · '+(d.inactivas||0)+' inactivas'},
    {l:'Visitas totales',v:pFmt(d.visitas_total,'n'),s:'Visitas únicas'},
    {l:'Ventas brutas',v:pFmt(d.ventas_brutas_total,'ars'),s:d.periodo},
    {l:'Transacciones',v:pFmt(d.transacciones,'n'),s:'Pedidos registrados'},
    {l:'Conversión global',v:pFmt(d.conv_global,'pct'),s:'Visitas → ventas'},
  ]
  const kc=document.getElementById('pubs-kpis')
  if(kc) kc.innerHTML=kpis.map(k=>`
    <div class="pk"><div class="pk-l">${k.l}</div><div class="pk-v">${k.v}</div><div class="pk-s">${k.s}</div></div>`).join('')

  // Tabla ventas
  const tv=document.getElementById('pubs-tb-ventas')
  const pp=document.getElementById('pubs-periodo-v')
  if(pp) pp.textContent=d.periodo
  const maxVB=Math.max(...(d.top_ventas||[]).map(r=>r.ventas_brutas||0))
  if(tv) tv.innerHTML=(d.top_ventas||[]).map((r,i)=>{
    const pct=maxVB?Math.round((r.ventas_brutas||0)/maxVB*100):0
    const cvr=(r.conv||0).toFixed(1)
    const cvrColor=r.conv>=5?'#15803d':r.conv>=2?'#854d0e':'#475569'
    return `<tr>
      <td style="color:#94a3b8;font-weight:500">${i+1}</td>
      <td><span class="pubs-nm" title="${r.nombre}">${r.nombre}</span></td>
      <td><span class="pubs-badge ${r.estado==='ACTIVA'?'act':'ina'}">${r.estado==='ACTIVA'?'Activa':'Inactiva'}</span></td>
      <td class="nr">${pFmt(r.visitas,'n')}</td>
      <td class="nr">${pFmt(r.ventas,'n')}</td>
      <td class="nr">${pFmt(r.unidades,'n')}</td>
      <td class="nr" style="font-weight:500;color:${cvrColor}">${cvr}%</td>
      <td class="nr" style="font-weight:700">${pFmt(r.ventas_brutas,'ars')}</td>
    </tr>`}).join('')

  // Tabla visitas
  const tvis=document.getElementById('pubs-tb-visitas')
  const pvis=document.getElementById('pubs-periodo-vis')
  if(pvis) pvis.textContent=d.periodo
  const maxVis=Math.max(...(d.top_visitas||[]).map(r=>r.visitas||0))
  if(tvis) tvis.innerHTML=(d.top_visitas||[]).map((r,i)=>`<tr>
    <td style="color:#94a3b8;font-weight:500">${i+1}</td>
    <td><span class="pubs-nm" title="${r.nombre}">${r.nombre}</span></td>
    <td><span class="pubs-badge ${r.estado==='ACTIVA'?'act':'ina'}">${r.estado==='ACTIVA'?'Activa':'Inactiva'}</span></td>
    <td class="nr" style="font-weight:500">${pFmt(r.visitas,'n')}</td>
    <td class="nr">${(r.conv||0).toFixed(1)}%</td>
    <td colspan="2">
      <div class="pubs-bar"><div class="pubs-bar-track"><div class="pubs-bar-fill" style="width:${maxVis?Math.round((r.visitas||0)/maxVis*100):0}%;background:${r.estado==='ACTIVA'?'#0d9488':'#94a3b8'}"></div></div></div>
    </td>
  </tr>`).join('')

  // Tabla CVR
  const tcvr=document.getElementById('pubs-tb-cvr')
  const maxCvr=Math.max(...(d.top_cvr||[]).map(r=>r.conv||0))
  if(tcvr) tcvr.innerHTML=(d.top_cvr||[]).map((r,i)=>{
    const chip=r.conv>=10?'ok':r.conv>=3?'warn':'bad'
    return `<tr>
      <td style="color:#94a3b8;font-weight:500">${i+1}</td>
      <td><span class="pubs-nm" title="${r.nombre}">${r.nombre}</span></td>
      <td><span class="pubs-badge ${r.estado==='ACTIVA'?'act':'ina'}">${r.estado==='ACTIVA'?'Activa':'Inactiva'}</span></td>
      <td class="nr">${pFmt(r.visitas,'n')}</td>
      <td class="nr">${pFmt(r.ventas,'n')}</td>
      <td class="nr"><span class="pubs-chip ${chip}">${(r.conv||0).toFixed(1)}%</span></td>
      <td colspan="2">
        <div class="pubs-bar"><div class="pubs-bar-track" style="width:80px"><div class="pubs-bar-fill" style="width:${maxCvr?Math.round((r.conv||0)/maxCvr*100):0}%;background:${chip==='ok'?'#15803d':chip==='warn'?'#854d0e':'#b91c1c'}"></div></div></div>
      </td>
    </tr>`}).join('')

  // Inicializar gráficos si ya está visible la pestaña
  const gPanel=document.getElementById('pubs-panel-graficos')
  if(gPanel && gPanel.classList.contains('on')) renderPubsCharts(d)
}

function dcP(id){ if(PUBS_CHARTS[id]){ PUBS_CHARTS[id].destroy(); delete PUBS_CHARTS[id] } }

function renderPubsCharts(d){
  const grd='rgba(0,0,0,0.05)'
  const tx='#64748b'

  // 1. Ventas brutas top 10 — barras horizontales
  dcP('pubs-ch-ventas')
  const top10=(d.top_ventas||[]).slice(0,10).reverse()
  const ctx1=document.getElementById('pubs-ch-ventas')
  if(ctx1) PUBS_CHARTS['pubs-ch-ventas']=new Chart(ctx1,{
    type:'bar',
    data:{
      labels:top10.map(r=>r.nombre.length>28?r.nombre.slice(0,26)+'…':r.nombre),
      datasets:[{
        label:'Ventas brutas',
        data:top10.map(r=>r.ventas_brutas||0),
        backgroundColor:top10.map(r=>r.estado==='ACTIVA'?'#0d9488':'#94a3b8'),
        borderRadius:4, borderSkipped:false
      }]
    },
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        datalabels:{
          display:true, anchor:'end', align:'right', clip:false,
          font:{size:10}, color:'#475569',
          formatter:v=>'$'+Math.round(v/1000)+'k'
        }
      },
      layout:{padding:{right:48}},
      scales:{
        x:{grid:{color:grd},ticks:{callback:v=>'$'+Math.round(v/1000)+'k',font:{size:10},color:tx},border:{display:false}},
        y:{grid:{display:false},ticks:{font:{size:10},color:tx},border:{display:false}}
      }
    }
  })

  // 2. Visitas vs Ventas
  dcP('pubs-ch-visitas')
  const topVis=(d.top_visitas||[]).slice(0,8)
  const ctx2=document.getElementById('pubs-ch-visitas')
  if(ctx2) PUBS_CHARTS['pubs-ch-visitas']=new Chart(ctx2,{
    type:'bar',
    data:{
      labels:topVis.map(r=>r.nombre.length>18?r.nombre.slice(0,16)+'…':r.nombre),
      datasets:[
        {label:'Visitas',data:topVis.map(r=>r.visitas||0),backgroundColor:'#378ADD',borderRadius:3,borderSkipped:false},
        {label:'Ventas (×10)',data:topVis.map(r=>(r.ventas||0)*10),backgroundColor:'#1D9E75',borderRadius:3,borderSkipped:false}
      ]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{legend:{position:'top',labels:{font:{size:10},boxWidth:10}}, datalabels:{display:false}},
      scales:{
        x:{ticks:{font:{size:9},maxRotation:45,color:tx},grid:{display:false},border:{display:false}},
        y:{ticks:{font:{size:9},color:tx},grid:{color:grd},border:{display:false}}
      }
    }
  })

  // 3. CVR
  dcP('pubs-ch-cvr')
  const topCvr=(d.top_cvr||[]).slice(0,8)
  const ctx3=document.getElementById('pubs-ch-cvr')
  if(ctx3) PUBS_CHARTS['pubs-ch-cvr']=new Chart(ctx3,{
    type:'bar',
    data:{
      labels:topCvr.map(r=>r.nombre.length>18?r.nombre.slice(0,16)+'…':r.nombre),
      datasets:[
        {label:'CVR %',data:topCvr.map(r=>r.conv||0),
          backgroundColor:topCvr.map(r=>r.conv>=10?'#15803d':r.conv>=3?'#854d0e':'#b91c1c'),
          borderRadius:3, borderSkipped:false},
        {label:'Objetivo (3%)',data:topCvr.map(()=>3),type:'line',
          borderColor:'#94a3b8',borderDash:[5,4],pointRadius:0,fill:false,tension:0}
      ]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{legend:{position:'top',labels:{font:{size:10},boxWidth:10}}, datalabels:{display:false}},
      scales:{
        x:{ticks:{font:{size:9},maxRotation:45,color:tx},grid:{display:false},border:{display:false}},
        y:{ticks:{callback:v=>v+'%',font:{size:9},color:tx},grid:{color:grd},border:{display:false}}
      }
    }
  })

  // 4. Activas vs inactivas (donut)
  dcP('pubs-ch-estado')
  const ctx4=document.getElementById('pubs-ch-estado')
  if(ctx4) PUBS_CHARTS['pubs-ch-estado']=new Chart(ctx4,{
    type:'doughnut',
    data:{
      labels:['Activas','Inactivas'],
      datasets:[{data:[d.activas||0,d.inactivas||0],backgroundColor:['#0d9488','#e2e8f0'],borderWidth:0}]
    },
    options:{
      responsive:true, maintainAspectRatio:false, cutout:'62%',
      plugins:{legend:{position:'bottom',labels:{font:{size:10},boxWidth:10}}, datalabels:{display:false}}
    }
  })

  // 5. Visitas activas vs inactivas (donut)
  dcP('pubs-ch-vis-estado')
  const visAct=(d.top_ventas||[]).filter(r=>r.estado==='ACTIVA').reduce((s,r)=>s+(r.visitas||0),0)
  const visIna=(d.top_ventas||[]).filter(r=>r.estado!=='ACTIVA').reduce((s,r)=>s+(r.visitas||0),0)
  const ctx5=document.getElementById('pubs-ch-vis-estado')
  if(ctx5) PUBS_CHARTS['pubs-ch-vis-estado']=new Chart(ctx5,{
    type:'doughnut',
    data:{
      labels:['Activas','Inactivas'],
      datasets:[{data:[visAct,visIna],backgroundColor:['#378ADD','#B4B2A9'],borderWidth:0}]
    },
    options:{
      responsive:true, maintainAspectRatio:false, cutout:'62%',
      plugins:{legend:{position:'bottom',labels:{font:{size:10},boxWidth:10}}, datalabels:{display:false}}
    }
  })
}


function showCotizadorModule(){
  hideMainViews()
  document.getElementById('cotizador-module').style.display='flex'
  document.querySelectorAll('.tbtn').forEach(b=>b.style.outline='none')
  const btn=document.getElementById('tbtn-cotizador')
  if(btn) btn.style.outline='2px solid #e8b4b4'
}

// ══════ Cotizador de vinos y bebidas (módulo integrado) ══════

let cot_productos = [];
let cot_idCounter = 1;

const cot_HEADERS_NOMBRE = ['producto','nombre','titulo','título','descripcion','descripción','item','articulo','artículo'];
const cot_HEADERS_COSTO = ['costo','costo proveedor','costo unitario','precio costo','cost'];
const cot_HEADERS_ACTUAL = ['precio actual','precio','pvp actual','precio venta actual','precio ml'];
const cot_HEADERS_PESO = ['peso','peso kg','peso (kg)','kg'];
const cot_HEADERS_LOGISTICA = ['logistica','logística','envio','envío','tipo de envio','tipo de envío'];
const cot_HEADERS_MARGEN = ['margen','margen %','margen objetivo'];
const cot_HEADERS_ML = ['precio ml','precio mercado libre','mas vendido','más vendido','precio competencia'];

function cot_normalizarHeader(h){
  return String(h||'').toLowerCase().trim()
    .normalize('NFD').replace(/[̀-ͯ]/g,'');
}

function cot_detectarColumnas(headerRow){
  const norm = headerRow.map(cot_normalizarHeader);
  const find = (candidatos) => {
    for (const c of candidatos){
      const cn = cot_normalizarHeader(c);
      const idx = norm.findIndex(h => h === cn || h.includes(cn));
      if (idx !== -1) return idx;
    }
    return -1;
  };
  return {
    nombre: find(cot_HEADERS_NOMBRE),
    costo: find(cot_HEADERS_COSTO),
    actual: find(cot_HEADERS_ACTUAL),
    peso: find(cot_HEADERS_PESO),
    logistica: find(cot_HEADERS_LOGISTICA),
    margen: find(cot_HEADERS_MARGEN),
    ml: find(cot_HEADERS_ML),
  };
}

function cot_parseLogistica(raw){
  const n = cot_normalizarHeader(raw);
  if (n.includes('flex')) return 'flex';
  return 'full';
}

function cot_manejarArchivo(ev){
  const file = ev.target.files[0];
  const status = document.getElementById('cot_file_status');
  if (!file) return;
  status.textContent = 'Leyendo archivo...';

  const costoConIva = document.getElementById('cot_in_file_costoiva').value === 'si';
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, {type:'array'});
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, {header:1, defval:''});
      if (!rows.length){ status.textContent = 'El archivo está vacío.'; return; }

      const headerRow = rows[0];
      const cols = cot_detectarColumnas(headerRow);
      if (cols.nombre === -1 || cols.costo === -1){
        status.textContent = 'No pude detectar las columnas de producto y costo. Revisá que el Excel tenga encabezados claros (ej: "Producto", "Costo").';
        return;
      }

      let agregados = 0, salteados = 0, pesosEstimados = 0;
      for (let i=1;i<rows.length;i++){
        const row = rows[i];
        const nombre = String(row[cols.nombre]||'').trim();
        const costoRaw = row[cols.costo];
        const costo = typeof costoRaw === 'number' ? costoRaw : parseFloat(String(costoRaw).replace(/[^0-9.,-]/g,'').replace(',','.'));
        if (!nombre || isNaN(costo) || costo <= 0){ salteados++; continue; }

        let actual = null;
        if (cols.actual !== -1){
          const actualRaw = row[cols.actual];
          const parsedActual = typeof actualRaw === 'number' ? actualRaw : parseFloat(String(actualRaw).replace(/[^0-9.,-]/g,'').replace(',','.'));
          if (!isNaN(parsedActual) && parsedActual > 0) actual = parsedActual;
        }

        let pesoKg = null;
        if (cols.peso !== -1){
          const pesoRaw = row[cols.peso];
          const parsedPeso = typeof pesoRaw === 'number' ? pesoRaw : parseFloat(String(pesoRaw).replace(',','.'));
          if (!isNaN(parsedPeso) && parsedPeso > 0) pesoKg = parsedPeso;
        }

        let pesoAuto = false, pesoConfianza = null;
        if (pesoKg === null){
          const est = cot_estimarPesoKg(nombre);
          pesoKg = est.peso; pesoAuto = true; pesoConfianza = est.confianza;
          pesosEstimados++;
        }

        const logistica = cols.logistica !== -1 ? cot_parseLogistica(row[cols.logistica]) : 'full';

        let margenOverride = null;
        if (cols.margen !== -1){
          const mRaw = row[cols.margen];
          const mParsed = typeof mRaw === 'number' ? mRaw : parseFloat(String(mRaw).replace(',','.'));
          if (!isNaN(mParsed) && mParsed > 0) margenOverride = mParsed;
        }

        let precioMl = null;
        if (cols.ml !== -1){
          const mlRaw = row[cols.ml];
          const mlParsed = typeof mlRaw === 'number' ? mlRaw : parseFloat(String(mlRaw).replace(/[^0-9.,-]/g,'').replace(',','.'));
          if (!isNaN(mlParsed) && mlParsed > 0) precioMl = mlParsed;
        }

        cot_productos.push({id: cot_idCounter++, nombre, costo, costoIva: costoConIva, logistica, pesoKg, pesoAuto, pesoConfianza, margenOverride, actual, precioMl});
        agregados++;
      }

      status.textContent = `Se cargaron ${agregados} cot_productos.`
        + (pesosEstimados ? ` A ${pesosEstimados} se les estimó el peso automáticamente por el título — revisalos en la cot_tabla (marcados "auto").` : '')
        + (salteados ? ` (${salteados} filas salteadas por datos incompletos)` : '');
      cot_render();
    } catch (err){
      status.textContent = 'No pude leer el archivo: ' + err.message;
    }
  };
  reader.readAsArrayBuffer(file);
}

function cot_exportarExcel(){
  const pp = cot_params();
  const filas = cot_productos.map(p => {
    const costoFinal = p.costoIva ? p.costo : p.costo * (1 + pp.iva/100);
    const sugerido = cot_pvpObjetivo(p.costo, p.costoIva, pp, p.logistica, p.pesoKg, p.margenOverride);
    let margenActual = '';
    if (p.actual !== null && !isNaN(p.actual) && p.actual > 0){
      const u = cot_utilidad(p.actual, costoFinal, pp.envioProv, pp, p.logistica, p.pesoKg);
      margenActual = Math.round((u / p.actual) * 1000) / 10;
    }
    let vsMl = '';
    if (p.precioMl){
      vsMl = Math.round(((sugerido - p.precioMl) / p.precioMl) * 1000) / 10;
    }
    return {
      'Producto': p.nombre,
      'Costo': p.costo,
      'Costo incluye IVA': p.costoIva ? 'Si' : 'No',
      'Logística': p.logistica === 'flex' ? 'Flex' : 'Full/Mercado Envíos',
      'Peso (kg)': p.pesoKg || '',
      'Peso estimado automáticamente': p.pesoAuto ? 'Si' : 'No',
      'Margen usado %': p.margenOverride ?? pp.margen,
      'Precio actual': p.actual===null ? '' : p.actual,
      'Margen actual %': margenActual,
      'PVP sugerido': Math.round(sugerido),
      'Precio ML más vendido': p.precioMl || '',
      'Vs. ML %': vsMl,
    };
  });
  const ws = XLSX.utils.json_to_sheet(filas);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'PVP sugeridos');
  XLSX.writeFile(wb, 'pvp_sugeridos_vinos.xlsx');
}

function cot_fmt(n){
  if (n===null || n===undefined || isNaN(n)) return '—';
  return n.toLocaleString('es-AR', {style:'currency', currency:'ARS', maximumFractionDigits:0});
}

const cot_TABLA_FULL_ME = [
  {maxKg: 0.3,   v: [1330, 2740, 3320]},
  {maxKg: 0.5,   v: [1370, 2760, 3340]},
  {maxKg: 1,     v: [1390, 2780, 3360]},
  {maxKg: 1.5,   v: [1410, 2800, 3380]},
  {maxKg: 2,     v: [1430, 2820, 3400]},
  {maxKg: 3,     v: [1450, 2860, 3470]},
  {maxKg: 4,     v: [1470, 2910, 3520]},
  {maxKg: 5,     v: [1500, 3040, 3670]},
  {maxKg: 8,     v: [1520, 3130, 3760]},
  {maxKg: 10,    v: [1560, 3180, 3910]},
  {maxKg: 13,    v: [1590, 3220, 4020]},
  {maxKg: 15,    v: [1620, 3280, 4060]},
  {maxKg: 20,    v: [1640, 3320, 4100]},
  {maxKg: 25,    v: [1660, 3380, 4170]},
  {maxKg: 30,    v: [1680, 3410, 4210]},
  {maxKg: 40,    v: [1700, 3440, 4250]},
  {maxKg: 50,    v: [1720, 3460, 4300]},
  {maxKg: 60,    v: [1740, 3480, 4320]},
  {maxKg: 70,    v: [1760, 3510, 4350]},
  {maxKg: 80,    v: [1780, 3530, 4370]},
  {maxKg: 90,    v: [1800, 3560, 4400]},
  {maxKg: 100,   v: [1820, 3580, 4420]},
  {maxKg: 120,   v: [1840, 3600, 4450]},
  {maxKg: 140,   v: [1860, 3620, 4470]},
  {maxKg: 160,   v: [1880, 3650, 4500]},
  {maxKg: 180,   v: [1900, 3680, 4520]},
  {maxKg: Infinity, v: [1920, 3700, 4550]},
];

const cot_TABLA_FLEX = [1330, 2740, 3320];

const cot_TABLA_FULL_ME_GRATIS = [
  {maxKg: 0.3,   v: [6190, 6790]},
  {maxKg: 0.5,   v: [6790, 7290]},
  {maxKg: 1,     v: [7790, 8290]},
  {maxKg: 1.5,   v: [7990, 8590]},
  {maxKg: 2,     v: [8290, 8790]},
  {maxKg: 3,     v: [8890, 9590]},
  {maxKg: 4,     v: [9790, 10890]},
  {maxKg: 5,     v: [10790, 11890]},
  {maxKg: 8,     v: [11790, 13090]},
  {maxKg: 10,    v: [12790, 14190]},
  {maxKg: 13,    v: [13790, 15190]},
  {maxKg: 15,    v: [14790, 16290]},
  {maxKg: 20,    v: [17590, 19390]},
  {maxKg: 25,    v: [20890, 23390]},
  {maxKg: 30,    v: [28590, 32090]},
  {maxKg: 40,    v: [32590, 36990]},
  {maxKg: 50,    v: [34390, 39090]},
  {maxKg: 60,    v: [38090, 43590]},
  {maxKg: 70,    v: [39590, 45490]},
  {maxKg: 80,    v: [45790, 52690]},
  {maxKg: 90,    v: [56390, 65190]},
  {maxKg: 100,   v: [64990, 75090]},
  {maxKg: 120,   v: [70890, 81990]},
  {maxKg: 140,   v: [79790, 92290]},
  {maxKg: 160,   v: [88690, 102690]},
  {maxKg: 180,   v: [97490, 112990]},
  {maxKg: Infinity, v: [106390, 123290]},
];

function cot_normalizarTexto(s){
  return String(s||'').toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g,'')
    .replace(/\s+/g,' ').trim();
}

function cot_estimarPesoKg(nombre){
  const t = cot_normalizarTexto(nombre);

  let volL = null;
  const mVol = t.match(/(\d+(?:[.,]\d+)?)\s*(ml|cc|cl|l|lt|lts|litro|litros)\b/);
  if (mVol){
    let val = parseFloat(mVol[1].replace(',','.'));
    const unidad = mVol[2];
    if (unidad === 'ml' || unidad === 'cc') volL = val/1000;
    else if (unidad === 'cl') volL = val/100;
    else volL = val;
  }

  let cant = 1;
  const mCantVol = t.match(/(\d{1,3})\s*x\s*\d+(?:[.,]\d+)?\s*(?:ml|cc|cl|l|lt|lts)/);
  const mCant = t.match(/(?:x\s*|pack\s*(?:de\s*)?x?\s*|caja\s*(?:de\s*)?x?\s*)(\d{1,3})\b/);
  if (mCantVol) cant = parseInt(mCantVol[1], 10);
  else if (mCant) cant = parseInt(mCant[1], 10);
  if (!cant || cant < 1 || cant > 48) cant = (cant>48) ? 1 : (cant||1);

  if (volL === null){
    volL = 0.75;
    var confianza = 'baja';
  } else {
    var confianza = 'media';
  }

  let tipo = 'vidrio';
  if (/\blata(s)?\b/.test(t)) tipo = 'lata';
  else if (/tetra ?brik|tetrapak|caja de carton|brick/.test(t)) tipo = 'tetra';
  else if (/\bpet\b|plastic/.test(t)) tipo = 'plastico';
  else if (/dama ?juana/.test(t)) tipo = 'damajuana';
  else if (/bag ?in ?box|bib\b/.test(t)) tipo = 'bagbox';
  else if (/vidrio|botella/.test(t)) tipo = 'vidrio';

  const esEspumante = /champagne|espumante|sidra|cava\b|prosecco/.test(t);

  let pesoEnvaseUnit;
  if (tipo === 'lata'){
    pesoEnvaseUnit = 0.02 * (volL/0.355);
  } else if (tipo === 'tetra'){
    pesoEnvaseUnit = 0.035 * (volL/1);
  } else if (tipo === 'plastico'){
    pesoEnvaseUnit = 0.06 * (volL/1.5);
  } else if (tipo === 'bagbox'){
    pesoEnvaseUnit = 0.22;
  } else if (tipo === 'damajuana'){
    pesoEnvaseUnit = 1.3 * (volL/5);
  } else {
    pesoEnvaseUnit = 0.55 * Math.pow(volL/0.75, 0.85);
    if (esEspumante) pesoEnvaseUnit *= 1.3;
  }
  pesoEnvaseUnit = Math.max(pesoEnvaseUnit, 0.015);

  const pesoLiquidoUnit = volL * 1.0;
  const pesoUnitario = pesoLiquidoUnit + pesoEnvaseUnit;

  let pesoTotal = pesoUnitario * cant;
  if (cant > 1) pesoTotal += 0.08 + 0.04*cant;

  return {peso: Math.round(pesoTotal*100)/100, confianza};
}

function cot_bandaPvp(pvp){
  if (pvp < 15000) return 0;
  if (pvp < 24000) return 1;
  return 2;
}

function cot_tramo(pvp, logistica, pesoKg){
  const peso = (pesoKg && pesoKg > 0) ? pesoKg : 0.3;

  if (pvp >= 33000){
    const bandaGratis = pvp < 50000 ? 0 : 1;
    const filaGratis = cot_TABLA_FULL_ME_GRATIS.find(f => peso <= f.maxKg) || cot_TABLA_FULL_ME_GRATIS[cot_TABLA_FULL_ME_GRATIS.length-1];
    if (logistica === 'flex'){
      return [0, filaGratis.v[bandaGratis]];
    }
    return [0, filaGratis.v[bandaGratis]];
  }

  const banda = cot_bandaPvp(pvp);
  if (logistica === 'flex'){
    return [cot_TABLA_FLEX[banda], 0];
  }
  const fila = cot_TABLA_FULL_ME.find(f => peso <= f.maxKg) || cot_TABLA_FULL_ME[cot_TABLA_FULL_ME.length-1];
  return [fila.v[banda], 0];
}

const cot_PARAM_IDS = ['cot_p_margen','cot_p_meli','cot_p_iva','cot_p_iibb','cot_p_debcred','cot_p_envioprov','cot_p_embalaje'];
const cot_PARAM_DEFAULTS = {cot_p_margen:10, cot_p_meli:13, cot_p_iva:21, cot_p_iibb:3, cot_p_debcred:1.2, cot_p_envioprov:0, cot_p_embalaje:1000};
const cot_LS_KEY = 'wesell_calc_vinos_params';

function cot_params(){
  return {
    margen: parseFloat(document.getElementById('cot_p_margen').value) || 0,
    meli: parseFloat(document.getElementById('cot_p_meli').value) || 0,
    iva: parseFloat(document.getElementById('cot_p_iva').value) || 0,
    iibb: parseFloat(document.getElementById('cot_p_iibb').value) || 0,
    debcred: parseFloat(document.getElementById('cot_p_debcred').value) || 0,
    envioProv: parseFloat(document.getElementById('cot_p_envioprov').value) || 0,
    embalaje: parseFloat(document.getElementById('cot_p_embalaje').value) || 0,
  };
}

function cot_guardarParametros(){
  try {
    const data = {};
    cot_PARAM_IDS.forEach(id => data[id] = document.getElementById(id).value);
    localStorage.setItem(cot_LS_KEY, JSON.stringify(data));
    const st = document.getElementById('cot_params_status');
    if (st){ st.textContent = 'Guardado ✓'; setTimeout(()=>{ if(st.textContent==='Guardado ✓') st.textContent=''; }, 1500); }
  } catch(e){}
}

function cot_cargarParametros(){
  try {
    const raw = localStorage.getItem(cot_LS_KEY);
    if (!raw) return;
    const data = JSON.parse(raw);
    cot_PARAM_IDS.forEach(id => { if (data[id] !== undefined) document.getElementById(id).value = data[id]; });
  } catch(e){}
}

function cot_aplicarMargenATodos(){
  cot_guardarParametros();
  if (cot_productos.length === 0){
    cot_render();
    const st0 = document.getElementById('cot_params_status');
    if (st0){ st0.textContent = 'Datos guardados. Se van a usar automáticamente para todo lo que cargues ahora (a mano o por Excel) ✓'; }
    return;
  }
  cot_productos.forEach(p => { p.margenOverride = null; });
  cot_render();
  const st = document.getElementById('cot_params_status');
  if (st){ st.textContent = `Recalculado: ${cot_productos.length} cot_productos usando margen ${document.getElementById('cot_p_margen').value}%, comisión ${document.getElementById('cot_p_meli').value}%, IVA ${document.getElementById('cot_p_iva').value}%, IIBB ${document.getElementById('cot_p_iibb').value}%, Déb/Créd ${document.getElementById('cot_p_debcred').value}% ✓`; }
}

function cot_restablecerParametros(){
  cot_PARAM_IDS.forEach(id => document.getElementById(id).value = cot_PARAM_DEFAULTS[id]);
  cot_guardarParametros();
  cot_render();
}

function cot_utilidad(pvp, costoFinal, envioProv, pp, logistica, pesoKg){
  const [cf, env] = cot_tramo(pvp, logistica, pesoKg);
  const costoMeli = pvp*pp.meli/100;
  const iibb = (pvp/(1+pp.iva/100)) * pp.iibb/100;
  const debcred = pvp*pp.debcred/100;
  return pvp - (costoMeli + iibb + debcred + cf + costoFinal + env + envioProv + pp.embalaje);
}

function cot_desglose(costo, costoConIva, pp, logistica, pesoKg, margenOverride){
  const costoFinal = costoConIva ? costo : costo * (1 + pp.iva/100);
  const ivaSumado = costoConIva ? 0 : (costo * pp.iva/100);
  const sugerido = cot_pvpObjetivo(costo, costoConIva, pp, logistica, pesoKg, margenOverride);
  const margen = (margenOverride !== null && margenOverride !== undefined && !isNaN(margenOverride)) ? margenOverride : pp.margen;
  if (isNaN(sugerido)) return {ok:false, margen};
  const [cf, env] = cot_tramo(sugerido, logistica, pesoKg);
  const comisionMeli = sugerido * pp.meli/100;
  const iibb = (sugerido/(1+pp.iva/100)) * pp.iibb/100;
  const debcred = sugerido * pp.debcred/100;
  const gananciaNeta = sugerido - (comisionMeli + iibb + debcred + cf + costoFinal + env + pp.envioProv + pp.embalaje);
  return {
    ok:true, sugerido, margen,
    costoBase: costo, ivaSumado, costoFinal,
    comisionMeli, iibbPct: pp.iibb, iibb, debcredPct: pp.debcred, debcred,
    costoFijoEnvioMl: cf, envioMlGratis: env,
    envioProveedor: pp.envioProv, embalaje: pp.embalaje,
    gananciaNeta,
  };
}

function cot_pvpObjetivo(costo, costoConIva, pp, logistica, pesoKg, margenOverride){
  const costoFinal = costoConIva ? costo : costo * (1 + pp.iva/100);
  const margen = (margenOverride !== null && margenOverride !== undefined && !isNaN(margenOverride)) ? margenOverride : pp.margen;
  const K = 1 - margen/100 - pp.meli/100 - (pp.iibb/100)/(1+pp.iva/100) - pp.debcred/100;
  if (K <= 0) return NaN;
  let [cf, env] = cot_tramo(costoFinal * 1.5, logistica, pesoKg);
  let pvp = 0;
  for (let i=0;i<8;i++){
    pvp = (cf + costoFinal + env + pp.envioProv + pp.embalaje) / K;
    const nuevo = cot_tramo(pvp, logistica, pesoKg);
    if (nuevo[0]===cf && nuevo[1]===env) break;
    [cf, env] = nuevo;
  }
  return pvp;
}

function cot_toggleCampoPeso(){
  const logistica = document.getElementById('cot_in_logistica').value;
  document.getElementById('cot_campo_peso').style.display = logistica === 'flex' ? 'none' : '';
}

function cot_agregarProducto(){
  const nombre = document.getElementById('cot_in_nombre').value.trim();
  const costo = parseFloat(document.getElementById('cot_in_costo').value);
  const costoIva = document.getElementById('cot_in_costoiva').value === 'si';
  const logistica = document.getElementById('cot_in_logistica').value;
  const pesoInput = parseFloat(document.getElementById('cot_in_peso').value);
  const margenRaw = document.getElementById('cot_in_margen').value;
  const margenOverride = margenRaw === '' ? null : parseFloat(margenRaw);
  const actualRaw = document.getElementById('cot_in_actual').value;
  const actual = actualRaw === '' ? null : parseFloat(actualRaw);
  const mlRaw = document.getElementById('cot_in_ml').value;
  const precioMl = mlRaw === '' ? null : parseFloat(mlRaw);

  if (!nombre || isNaN(costo) || costo <= 0){
    alert('Completá el nombre y un costo válido.');
    return;
  }

  let pesoKg, pesoAuto, pesoConfianza;
  if (!isNaN(pesoInput) && pesoInput > 0){
    pesoKg = pesoInput; pesoAuto = false; pesoConfianza = null;
  } else {
    const est = cot_estimarPesoKg(nombre);
    pesoKg = est.peso; pesoAuto = true; pesoConfianza = est.confianza;
  }

  cot_productos.push({id: cot_idCounter++, nombre, costo, costoIva, logistica, pesoKg, pesoAuto, pesoConfianza, margenOverride, actual, precioMl});

  document.getElementById('cot_in_nombre').value = '';
  document.getElementById('cot_in_costo').value = '';
  document.getElementById('cot_in_peso').value = '';
  document.getElementById('cot_in_margen').value = '';
  document.getElementById('cot_in_actual').value = '';
  document.getElementById('cot_in_ml').value = '';
  document.getElementById('cot_in_nombre').focus();

  cot_render();
}

function cot_quitarProducto(id){
  cot_productos = cot_productos.filter(p => p.id !== id);
  cot_render();
}

function cot_toggleDesglose(id){
  const p = cot_productos.find(x => x.id === id);
  if (!p) return;
  p.expandido = !p.expandido;
  cot_render();
}

function cot_limpiarTodo(){
  if (cot_productos.length === 0) return;
  if (!confirm(`¿Borrar los ${cot_productos.length} cot_productos cargados? Esto no se puede deshacer (si querés conservarlos, descargá el Excel antes).`)) return;
  cot_productos = [];
  const fileInput = document.getElementById('cot_in_file');
  if (fileInput) fileInput.value = '';
  const status = document.getElementById('cot_file_status');
  if (status) status.textContent = '';
  cot_render();
}

function cot_actualizarPrecioActual(id, val){
  const p = cot_productos.find(x => x.id === id);
  if (p) p.actual = val === '' ? null : parseFloat(val);
  cot_render();
}

function cot_actualizarLogistica(id, val){
  const p = cot_productos.find(x => x.id === id);
  if (p) p.logistica = val;
  cot_render();
}

function cot_actualizarPeso(id, val){
  const p = cot_productos.find(x => x.id === id);
  if (!p) return;
  const n = parseFloat(val);
  if (!isNaN(n) && n > 0){
    p.pesoKg = n; p.pesoAuto = false; p.pesoConfianza = null;
  }
  cot_render();
}

function cot_actualizarMargen(id, val){
  const p = cot_productos.find(x => x.id === id);
  if (!p) return;
  p.margenOverride = val === '' ? null : parseFloat(val);
  cot_render();
}

function cot_actualizarPrecioMl(id, val){
  const p = cot_productos.find(x => x.id === id);
  if (!p) return;
  p.precioMl = val === '' ? null : parseFloat(val);
  cot_render();
}

function cot_reestimarPesos(){
  let n = 0;
  cot_productos.forEach(p => {
    if (p.pesoAuto){
      const est = cot_estimarPesoKg(p.nombre);
      p.pesoKg = est.peso; p.pesoConfianza = est.confianza;
      n++;
    }
  });
  cot_render();
  const status = document.getElementById('cot_file_status');
  if (status) status.textContent = `Se re-estimaron ${n} pesos automáticos. Los que cargaste o editaste a mano no se tocan.`;
}

function cot_actualizarResumenImpuestos(pp){
  const el = document.getElementById('cot_resumen_impuestos');
  if (!el) return;
  el.innerHTML = `Ahora mismo la calculadora está descontando, sobre cada PVP: <b>Comisión Mercado Libre ${pp.meli}%</b>, <b>IVA ${pp.iva}%</b> (sobre el costo cuando no lo trae incluido), <b>Ingresos Brutos ${pp.iibb}%</b> (calculado sobre el PVP sin IVA), <b>Ley de Débitos/Créditos ${pp.debcred}%</b>, más el costo fijo/envío de ML según el cot_tramo de precio y peso, envío proveedor $${pp.envioProv.toLocaleString('es-AR')} y embalaje $${pp.embalaje.toLocaleString('es-AR')}. El margen objetivo general es ${pp.margen}% (se puede pisar por producto en la cot_tabla). Tocá "ver desglose" en cualquier fila para ver el cálculo completo de ese producto.`;
}

function cot_render(){
  const pp = cot_params();
  cot_actualizarResumenImpuestos(pp);
  const cot_tbody = document.getElementById('cot_tbody');
  const cot_empty = document.getElementById('cot_empty');
  const cot_tabla = document.getElementById('cot_tabla');
  cot_tbody.innerHTML = '';

  if (cot_productos.length === 0){
    cot_tabla.style.display = 'none';
    cot_empty.style.display = 'block';
    document.getElementById('cot_summary').innerHTML = '';
    document.getElementById('cot_exportbar').style.display = 'none';
    return;
  }
  cot_tabla.style.display = 'table';
  cot_empty.style.display = 'none';
  document.getElementById('cot_exportbar').style.display = 'flex';

  let bajoObjetivo = 0;
  let totalRecuperar = 0;

  const filas = cot_productos.map(p => {
    const costoFinal = p.costoIva ? p.costo : p.costo * (1 + pp.iva/100);
    const sugerido = cot_pvpObjetivo(p.costo, p.costoIva, pp, p.logistica, p.pesoKg, p.margenOverride);

    let margenActualPct = null, diffTexto = '—', badge = '';
    if (p.actual !== null && !isNaN(p.actual) && p.actual > 0){
      const u = cot_utilidad(p.actual, costoFinal, pp.envioProv, pp, p.logistica, p.pesoKg);
      margenActualPct = (u / p.actual) * 100;
      const bajo = margenActualPct < pp.margen;
      badge = `<span class="badge ${bajo?'bad':'good'}">${margenActualPct.toFixed(1)}%</span>`;
      if (bajo){
        bajoObjetivo++;
        const falta = sugerido - p.actual;
        totalRecuperar += Math.max(falta,0);
        diffTexto = `+ ${cot_fmt(falta)}`;
      } else {
        diffTexto = 'OK';
      }
    } else {
      badge = '<span style="color:var(--ink-soft);font-size:12px;">sin precio actual</span>';
    }

    return {p, sugerido, margenActualPct, diffTexto, badge};
  });

  filas.sort((a,b)=>{
    if (a.margenActualPct===null && b.margenActualPct===null) return 0;
    if (a.margenActualPct===null) return 1;
    if (b.margenActualPct===null) return -1;
    return a.margenActualPct - b.margenActualPct;
  });

  filas.forEach(({p, sugerido, diffTexto, badge}) => {
    const tr = document.createElement('tr');
    const autoTag = p.pesoAuto ? ` <span style="color:var(--warn);font-weight:600;">(auto${p.pesoConfianza==='baja'?', revisar':''})</span>` : '';
    const metaExtra = p.costoIva? 'costo c/IVA':'costo s/IVA';

    let vsMlHtml;
    if (p.precioMl && !isNaN(sugerido)){
      const diffPct = ((sugerido - p.precioMl) / p.precioMl) * 100;
      const porEncima = sugerido > p.precioMl;
      const texto = porEncima ? 'No competitivo' : 'Competitivo';
      vsMlHtml = `<span class="badge ${porEncima?'bad':'good'}" style="margin-top:3px;" title="Tu PVP sugerido vs. el precio de ML: ${porEncima?'+':''}${diffPct.toFixed(1)}%">${texto} (${porEncima?'+':''}${diffPct.toFixed(1)}%)</span>`;
    } else {
      vsMlHtml = '';
    }

    tr.innerHTML = `
      <td><div class="prod-name">${p.nombre}</div><div class="prod-meta">${metaExtra}</div></td>
      <td>
        <select style="padding:5px 6px;border-radius:6px;border:1px solid var(--line);background:var(--bg);color:var(--ink);font-size:12px;margin-bottom:4px;" onchange="cot_actualizarLogistica(${p.id}, this.value)">
          <option value="full" ${p.logistica==='full'?'selected':''}>Full/ME</option>
          <option value="flex" ${p.logistica==='flex'?'selected':''}>Flex</option>
        </select>
        <div style="font-size:11px;color:var(--ink-soft);display:flex;align-items:center;gap:4px;">
          <input type="number" value="${p.pesoKg||''}" step="0.1" style="width:55px;padding:3px 5px;border-radius:5px;border:1px solid var(--line);background:var(--bg);color:var(--ink);font-size:11px;" onchange="cot_actualizarPeso(${p.id}, this.value)"> kg${autoTag}
        </div>
      </td>
      <td class="num">${cot_fmt(p.costo)}</td>
      <td class="num">
        <input class="pvp-actual" style="width:90px;" type="number" value="${p.actual===null?'':p.actual}" placeholder="precio actual" onchange="cot_actualizarPrecioActual(${p.id}, this.value)">
        <div style="margin-top:3px;">${badge}</div>
      </td>
      <td class="num"><input class="pvp-actual" style="width:70px;" type="number" value="${p.margenOverride===null||p.margenOverride===undefined?'':p.margenOverride}" placeholder="gral." step="0.5" onchange="cot_actualizarMargen(${p.id}, this.value)"></td>
      <td class="num">
        <div style="font-weight:700;">${isNaN(sugerido) ? '<span style="color:var(--bad);font-weight:600;font-size:12px;">margen inalcanzable</span>' : cot_fmt(sugerido)}</div>
        <div style="font-size:11px;color:var(--ink-soft);">${diffTexto}</div>
        <button class="btn-icon" style="padding:2px 0;font-size:11px;text-decoration:underline;" onclick="cot_toggleDesglose(${p.id})">ver desglose</button>
      </td>
      <td class="num">
        <input class="pvp-actual" style="width:90px;" type="number" value="${p.precioMl===null||p.precioMl===undefined?'':p.precioMl}" placeholder="precio ML" onchange="cot_actualizarPrecioMl(${p.id}, this.value)">
        <div style="margin-top:3px;">${vsMlHtml}</div>
      </td>
      <td><button class="btn-icon" onclick="cot_quitarProducto(${p.id})">Quitar</button></td>
    `;
    cot_tbody.appendChild(tr);

    if (p.expandido){
      const d = cot_desglose(p.costo, p.costoIva, pp, p.logistica, p.pesoKg, p.margenOverride);
      const trD = document.createElement('tr');
      let cuerpo;
      if (!d.ok){
        cuerpo = `<div style="padding:10px 4px;color:var(--bad);font-size:13px;">
          Con un margen objetivo de ${d.margen}% no alcanza: sumando comisión ML (${pp.meli}%), IIBB (${pp.iibb}%) y Ley Déb/Créd (${pp.debcred}%) ya se come más del 100% del PVP. Bajá el margen pedido para este producto o revisá los otros porcentajes.
        </div>`;
      } else {
        const iva = p.costoIva ? 'El costo ya venía con IVA incluido' : `Se sumó IVA (${pp.iva}%) porque el costo estaba sin IVA`;
        cuerpo = `<div style="padding:10px 4px;font-size:12.5px;line-height:1.9;">
          <b>Cómo se arma el PVP sugerido de ${cot_fmt(d.sugerido)}</b> (margen objetivo: ${d.margen}%)
          <table style="width:100%;margin-top:6px;font-size:12.5px;">
            <tr><td>Costo del producto</td><td class="num">${cot_fmt(d.costoBase)}</td></tr>
            <tr><td>${iva}${d.ivaSumado?` (+${cot_fmt(d.ivaSumado)})`:''}</td><td class="num">${cot_fmt(d.costoFinal)}</td></tr>
            <tr><td>Comisión Mercado Libre (${pp.meli}% del PVP)</td><td class="num">− ${cot_fmt(d.comisionMeli)}</td></tr>
            <tr><td>Ingresos Brutos (${pp.iibb}% sobre PVP sin IVA)</td><td class="num">− ${cot_fmt(d.iibb)}</td></tr>
            <tr><td>Ley Débitos/Créditos (${pp.debcred}% del PVP)</td><td class="num">− ${cot_fmt(d.debcred)}</td></tr>
            <tr><td>Costo fijo de envío ML (por tramo de precio/peso)</td><td class="num">− ${cot_fmt(d.costoFijoEnvioMl)}</td></tr>
            ${d.envioMlGratis ? `<tr><td>Envío gratis ML (PVP ≥ $33.000, por peso)</td><td class="num">− ${cot_fmt(d.envioMlGratis)}</td></tr>` : ''}
            ${d.envioProveedor ? `<tr><td>Envío proveedor</td><td class="num">− ${cot_fmt(d.envioProveedor)}</td></tr>` : ''}
            <tr><td>Embalaje</td><td class="num">− ${cot_fmt(d.embalaje)}</td></tr>
            <tr style="border-top:1px solid var(--line);font-weight:700;"><td>Ganancia neta (= margen objetivo)</td><td class="num">${cot_fmt(d.gananciaNeta)}</td></tr>
          </table>
        </div>`;
      }
      trD.innerHTML = `<td colspan="8" style="background:var(--bg);">${cuerpo}</td>`;
      cot_tbody.appendChild(trD);
    }
  });

  document.getElementById('cot_summary').innerHTML = `
    <div class="stat"><div class="n">${cot_productos.length}</div><div class="l">Productos cargados</div></div>
    <div class="stat"><div class="n" style="color:${bajoObjetivo>0?'var(--bad)':'var(--good)'}">${bajoObjetivo}</div><div class="l">Bajo el margen objetivo</div></div>
    <div class="stat"><div class="n">${cot_fmt(totalRecuperar)}</div><div class="l">A recuperar (suma de subas sugeridas)</div></div>
  `;
}

document.querySelectorAll('.cot_params input').forEach(inp => inp.addEventListener('input', () => { cot_guardarParametros(); cot_render(); }));
document.getElementById('cot_in_nombre').addEventListener('keydown', e => { if(e.key==='Enter') cot_agregarProducto(); });

cot_cargarParametros();
cot_toggleCampoPeso();
cot_render();

</script>
</body>
</html>
"""

# ═══════════════════════════════════════════════════════════════════════════
# API KEY + CHAT ROUTES
# ═══════════════════════════════════════════════════════════════════════════

@app.route("/api/chat", methods=["POST"])
def api_chat():
    """
    Motor analítico local inteligente.
    Extrae entidades (productos, meses, métricas, zonas) de la pregunta
    y genera respuesta + gráfico a medida sin API externa.
    """
    import difflib

    body     = request.json or {}
    sids     = body.get("sids", [])
    pregunta = body.get("pregunta", "").strip()
    filtros  = body.get("filtros", {})

    if not pregunta: return err("Pregunta vacia.")
    if not sids:     return err("No hay datos cargados.")

    try:
        df0 = _get_combined(sids)
        if df0.empty: return err("Sin datos.")
        df = apply_filters(df0, filtros)
        m  = metricas(df)
    except Exception as ex:
        traceback.print_exc()
        return err("Error procesando datos: " + str(ex))

    p_orig = pregunta
    p = pregunta.lower()

    # ── HELPERS ─────────────────────────────────────────────────────────
    def fmt_ars(v):
        try: return "$ {:,}".format(int(round(v))).replace(",", ".")
        except: return "$ 0"

    def var_pct(nuevo, viejo):
        if not viejo: return None
        return round((nuevo - viejo) / viejo * 100, 1)

    def flecha(pct):
        if pct is None: return ""
        return ("▲ +{}%".format(pct)) if pct > 0 else ("▼ {}%".format(pct))

    def no_grafico():
        return {"tipo":"none","titulo":"","labels":[],"values":[],"values2":[],
                "label_series":"","label_series2":"","color":"#2563eb"}

    def norm_grafico(g):
        def _v(x):
            # Convert numpy types to native Python
            try:
                import numpy as np
                if isinstance(x, (np.integer,)): return int(x)
                if isinstance(x, (np.floating,)): return float(x)
            except ImportError: pass
            if hasattr(x, 'item'): return x.item()
            return x
        def _vlist(lst):
            return [round(float(_v(x)), 2) if isinstance(_v(x), float) else _v(x) for x in (lst or [])]
        return {
            "tipo":          g.get("tipo","none"),
            "titulo":        str(g.get("titulo","")),
            "labels":        [str(l) for l in g.get("labels",[])],
            "values":        _vlist(g.get("values",[])),
            "values2":       _vlist(g.get("values2",[])),
            "label_series":  str(g.get("label_series","")),
            "label_series2": str(g.get("label_series2","")),
            "color":         str(g.get("color","#2563eb")),
        }

    def meses_data():
        if "_mes" not in df.columns: return {}
        result = {}
        for mes in sorted(df["_mes"].dropna().unique()):
            sub = df[df["_mes"] == mes]
            result[mes] = metricas(sub)
        return result

    # ── EXTRACCIÓN DE ENTIDADES ──────────────────────────────────────────
    MESES_MAP = {
        "enero":"01","febrero":"02","marzo":"03","abril":"04",
        "mayo":"05","junio":"06","julio":"07","agosto":"08",
        "septiembre":"09","octubre":"10","noviembre":"11","diciembre":"12",
        "ene":"01","feb":"02","mar":"03","abr":"04","jun":"06",
        "jul":"07","ago":"08","sep":"09","oct":"10","nov":"11","dic":"12",
    }
    ANIOS = ["2024","2025","2026","2027"]
    METRICAS_MAP = {
        "ingresos":["ingreso","facturacion","facturación","plata","dinero","venta en pesos","$"],
        "unidades":["unidad","unidades","cantidad","cuantos","cuántos","vendidos","piezas"],
        "n_ventas":["ordenes","órdenes","pedidos","ventas","transacciones","operaciones"],
        "ticket_prom":["ticket","promedio","precio promedio","valor promedio"],
        "costo":["costo","costos","gasto","gastos","comision"],
    }

    def detectar_metrica(texto):
        for col, kws in METRICAS_MAP.items():
            if any(kw in texto for kw in kws):
                return col
        return "ingresos"  # default

    def detectar_meses_en_texto(texto):
        found = []
        anio_actual = "2026"
        for anio in ANIOS:
            if anio in texto:
                anio_actual = anio
                break
        for nombre, num in MESES_MAP.items():
            if nombre in texto:
                # Try to find the year near this month
                idx = texto.find(nombre)
                ctx = texto[max(0,idx-5):idx+20]
                anio_cercano = anio_actual
                for a in ANIOS:
                    if a in ctx:
                        anio_cercano = a
                        break
                key = "{}-{}".format(anio_cercano, num)
                if key not in found:
                    found.append(key)
        return found

    def detectar_meses_disponibles():
        if "_mes" not in df.columns: return []
        return sorted(df["_mes"].dropna().unique().tolist())

    def buscar_producto(texto):
        """Busca el producto más similar al mencionado en la pregunta."""
        if "publicacion" not in df.columns: return None
        prods = df["publicacion"].dropna().unique().tolist()
        if not prods: return None
        # Buscar coincidencia directa (case insensitive)
        texto_lower = texto.lower()
        for p_name in prods:
            if p_name and p_name.lower() in texto_lower:
                return p_name
        # Buscar palabras del texto en nombres de productos
        palabras = [w for w in texto_lower.split() if len(w) > 3]
        best_score = 0
        best_prod  = None
        for p_name in prods:
            if not p_name: continue
            p_lower = p_name.lower()
            score = sum(1 for w in palabras if w in p_lower)
            if score > best_score:
                best_score = score
                best_prod  = p_name
        if best_score > 0:
            return best_prod
        # Fuzzy match como último recurso
        texto_words = texto_lower.split()
        for word in texto_words:
            if len(word) < 4: continue
            matches = difflib.get_close_matches(word,
                [pn.lower() for pn in prods if pn], n=1, cutoff=0.7)
            if matches:
                for pn in prods:
                    if pn and pn.lower() == matches[0]:
                        return pn
        return None

    def buscar_zona(texto):
        if "provincia" not in df.columns: return None
        zonas = df["provincia"].dropna().unique().tolist()
        texto_lower = texto.lower()
        for z in zonas:
            if z and z.lower() in texto_lower:
                return z
        return None

    # ── DETECCIÓN DE INTENCIÓN ───────────────────────────────────────────
    meses_en_pregunta  = detectar_meses_en_texto(p)
    meses_disponibles  = detectar_meses_disponibles()
    metrica_pedida     = detectar_metrica(p)
    producto_pedido    = buscar_producto(p_orig)
    zona_pedida        = buscar_zona(p_orig)
    es_comparacion     = any(x in p for x in ["compar","vs","versus","contra","diferencia","cambio entre","entre","diferencia"])
    es_evolucion       = any(x in p for x in ["evolucion","evoluci","tendencia","historico","tiempo","mes a mes","como fue","como estuvo"])
    es_top             = any(x in p for x in ["top","mejor","peor","mas vendido","mas factur","mayor","menor","ranking","primeros","ultimos"])
    es_distribucion    = any(x in p for x in ["distribucion","distribuci","porcentaje","proporcion","como se reparte","cuanto representa"])

    respuesta = ""
    g = no_grafico()

    # ══════════════════════════════════════════════════════════════════════
    # CASO A: Producto específico mencionado
    # ══════════════════════════════════════════════════════════════════════
    if producto_pedido and "publicacion" in df.columns:
        df_prod = df[df["publicacion"] == producto_pedido]

        # A1: COMPARACIÓN de producto entre 2 meses
        if len(meses_en_pregunta) == 2 and "_mes" in df.columns:
            m1, m2 = meses_en_pregunta[0], meses_en_pregunta[1]
            df1 = df_prod[df_prod["_mes"] == m1]
            df2 = df_prod[df_prod["_mes"] == m2]

            v1 = float(_R(df1[metrica_pedida].sum())) if metrica_pedida in df1.columns and not df1.empty else 0.0
            v2 = float(_R(df2[metrica_pedida].sum())) if metrica_pedida in df2.columns and not df2.empty else 0.0
            vp = var_pct(v2, v1)

            lbl_met = {"ingresos":"Ingresos","unidades":"Unidades","n_ventas":"Ventas","ticket_prom":"Ticket prom","costo":"Costos"}.get(metrica_pedida,"Valor")
            fmt_v = (lambda v: fmt_ars(v)) if metrica_pedida in ["ingresos","costo","ticket_prom"] else (lambda v: str(int(v)))

            respuesta = (
                "📊 **{}** — {} por mes\n\n".format(producto_pedido[:45], lbl_met) +
                "**{}:** {} {}\n".format(m1, fmt_v(v1), "" if v1 > 0 else "(sin ventas)") +
                "**{}:** {} {}\n".format(m2, fmt_v(v2), "" if v2 > 0 else "(sin ventas)") +
                "\n"
            )
            if v1 == 0 and v2 == 0:
                respuesta += "⚠️ No hay datos de este producto en esos meses."
            elif vp is not None:
                respuesta += "{} {} en {} comparado con {}.".format(
                    "▲ Creció un {}%".format(vp) if vp > 0 else "▼ Cayó un {}%".format(abs(vp)),
                    "en " + lbl_met.lower(), m2, m1)
            g = {
                "tipo": "bar",
                "titulo": "{} — {} ({} vs {})".format(producto_pedido[:30], lbl_met, m1, m2),
                "labels": [m1, m2],
                "values": [v1, v2],
                "values2": [],
                "label_series": lbl_met,
                "label_series2": "",
                "color": "#7c3aed"
            }

        # A2: Producto en TODOS los meses (evolución)
        elif "_mes" in df.columns and (es_evolucion or len(meses_disponibles) > 1):
            evol = []
            for mes in meses_disponibles:
                sub = df_prod[df_prod["_mes"] == mes]
                if sub.empty: continue
                val = _R(sub[metrica_pedida].sum()) if metrica_pedida in sub.columns else 0
                evol.append({"mes": mes, "valor": val})

            if not evol:
                respuesta = "No hay datos de **{}** en ningún mes del período.".format(producto_pedido)
            else:
                vals  = [e["valor"] for e in evol]
                lbl_met = {"ingresos":"Ingresos","unidades":"Unidades","n_ventas":"Ventas"}.get(metrica_pedida,"Valor")
                fmt_v  = (lambda v: fmt_ars(v)) if metrica_pedida in ["ingresos","costo"] else (lambda v: str(int(v)))
                max_v  = max(vals); max_m = evol[vals.index(max_v)]["mes"]
                respuesta = (
                    "📈 **{}** — evolución de {}\n\n".format(producto_pedido[:40], lbl_met.lower()) +
                    "**Mejor mes:** {} con {}\n".format(max_m, fmt_v(max_v)) +
                    "**Total período:** {}\n\n".format(fmt_v(sum(vals)))
                )
                if len(vals) >= 2:
                    vp = var_pct(vals[-1], vals[0])
                    respuesta += "Entre {} y {}: {}".format(evol[0]["mes"], evol[-1]["mes"], flecha(vp))
                g = {
                    "tipo": "line",
                    "titulo": "{} — {} por mes".format(producto_pedido[:30], lbl_met),
                    "labels": [e["mes"] for e in evol],
                    "values": [e["valor"] for e in evol],
                    "values2": [],
                    "label_series": lbl_met,
                    "label_series2": "",
                    "color": "#7c3aed"
                }

        # A3: Resumen del producto
        else:
            mm = metricas(df_prod)
            top_prov_prod = grp(df_prod, "provincia", top=3)
            respuesta = (
                "🛍️ **{}**\n\n".format(producto_pedido[:50]) +
                "**Ventas:** {} órdenes · **Unidades:** {} · **Ingresos:** {}\n".format(
                    mm["n_ventas"], int(mm["unidades"]), fmt_ars(mm["ingresos"])) +
                "**Ticket promedio:** {}\n".format(fmt_ars(mm["ticket_prom"]))
            )
            if top_prov_prod:
                respuesta += "**Provincias top:** " + ", ".join(
                    "{} ({}%)".format(r["label"], r.get("pct",0)) for r in top_prov_prod) + "\n"
            top_prods_all = grp(df, "publicacion", top=20)
            rank = next((i+1 for i,r in enumerate(top_prods_all) if r["label"] == producto_pedido), None)
            if rank:
                respuesta += "**Ranking:** #{} en ingresos entre todos tus productos".format(rank)

            g = {
                "tipo": "doughnut" if top_prov_prod else "none",
                "titulo": "{} — ventas por provincia".format(producto_pedido[:30]),
                "labels": [r["label"] for r in top_prov_prod],
                "values": [r["ingresos"] for r in top_prov_prod],
                "values2": [],
                "label_series": "Ingresos",
                "label_series2": "",
                "color": "#7c3aed"
            }

    # ══════════════════════════════════════════════════════════════════════
    # CASO B: Comparación de 2 meses específicos (sin producto)
    # ══════════════════════════════════════════════════════════════════════
    elif len(meses_en_pregunta) == 2 and "_mes" in df.columns:
        m1, m2 = meses_en_pregunta[0], meses_en_pregunta[1]
        df1 = df[df["_mes"] == m1]; df2 = df[df["_mes"] == m2]
        mm1 = metricas(df1); mm2 = metricas(df2)

        v_ing = var_pct(mm2["ingresos"],    mm1["ingresos"])
        v_ven = var_pct(mm2["n_ventas"],    mm1["n_ventas"])
        v_uds = var_pct(mm2["unidades"],    mm1["unidades"])
        v_tck = var_pct(mm2["ticket_prom"], mm1["ticket_prom"])

        respuesta = (
            "📊 **Comparación {} vs {}**\n\n".format(m1, m2) +
            "| Métrica | {} | {} | Variación |\n".format(m1, m2) +
            "|---|---|---|---|\n" +
            "| Ingresos | {} | {} | {} |\n".format(fmt_ars(mm1["ingresos"]), fmt_ars(mm2["ingresos"]), flecha(v_ing)) +
            "| Ventas | {} | {} | {} |\n".format(mm1["n_ventas"], mm2["n_ventas"], flecha(v_ven)) +
            "| Unidades | {} | {} | {} |\n".format(int(mm1["unidades"]), int(mm2["unidades"]), flecha(v_uds)) +
            "| Ticket prom | {} | {} | {} |".format(fmt_ars(mm1["ticket_prom"]), fmt_ars(mm2["ticket_prom"]), flecha(v_tck))
        )
        metrica_mostrar = metrica_pedida
        val1 = _R(df1[metrica_mostrar].sum()) if metrica_mostrar in df1.columns and not df1.empty else mm1.get(metrica_mostrar,0)
        val2 = _R(df2[metrica_mostrar].sum()) if metrica_mostrar in df2.columns and not df2.empty else mm2.get(metrica_mostrar,0)
        lbl_met = {"ingresos":"Ingresos","unidades":"Unidades","n_ventas":"Ventas"}.get(metrica_mostrar,"Valor")
        g = {
            "tipo": "bar",
            "titulo": "Comparación {} vs {} — {}".format(m1, m2, lbl_met),
            "labels": [m1, m2],
            "values": [val1, val2],
            "values2": [mm1["n_ventas"], mm2["n_ventas"]],
            "label_series": lbl_met, "label_series2": "Ventas",
            "color": "#2563eb"
        }

    # ══════════════════════════════════════════════════════════════════════
    # CASO C: Zona específica mencionada
    # ══════════════════════════════════════════════════════════════════════
    elif zona_pedida and "provincia" in df.columns:
        df_zona = df[df["provincia"] == zona_pedida]
        mm_zona = metricas(df_zona)
        top_prods_zona = grp(df_zona, "publicacion", top=5)

        respuesta = (
            "📍 **{}** — análisis de zona\n\n".format(zona_pedida) +
            "**Ventas:** {} órdenes · **Ingresos:** {} · **Unidades:** {}\n".format(
                mm_zona["n_ventas"], fmt_ars(mm_zona["ingresos"]), int(mm_zona["unidades"])) +
            "**Ticket promedio:** {}\n\n".format(fmt_ars(mm_zona["ticket_prom"]))
        )
        if top_prods_zona:
            respuesta += "**Top productos en {}:**\n".format(zona_pedida)
            for i, r in enumerate(top_prods_zona[:5], 1):
                respuesta += "{}. {} — {} ({} uds)\n".format(
                    i, r["label"][:35], fmt_ars(r["ingresos"]), int(r.get("unidades",0)))

        # Evolución de la zona por mes
        if "_mes" in df_zona.columns:
            evol_zona = []
            for mes in meses_disponibles:
                sub = df_zona[df_zona["_mes"] == mes]
                if sub.empty: continue
                evol_zona.append({"mes": mes, "valor": _R(float(sub["ingresos"].sum())) if "ingresos" in sub.columns else 0})
            if len(evol_zona) > 1:
                g = {
                    "tipo": "line",
                    "titulo": "{} — ingresos por mes".format(zona_pedida),
                    "labels": [e["mes"] for e in evol_zona],
                    "values": [e["valor"] for e in evol_zona],
                    "values2": [],
                    "label_series": "Ingresos",
                    "label_series2": "",
                    "color": "#0d9488"
                }
            elif top_prods_zona:
                g = {
                    "tipo": "bar",
                    "titulo": "Top productos en {}".format(zona_pedida),
                    "labels": [r["label"][:25] for r in top_prods_zona],
                    "values": [r["ingresos"] for r in top_prods_zona],
                    "values2": [r.get("unidades",0) for r in top_prods_zona],
                    "label_series": "Ingresos", "label_series2": "Unidades",
                    "color": "#0d9488"
                }

    # ══════════════════════════════════════════════════════════════════════
    # CASO D: Un mes específico mencionado (análisis de ese mes)
    # ══════════════════════════════════════════════════════════════════════
    elif len(meses_en_pregunta) == 1 and "_mes" in df.columns:
        mes_pedido = meses_en_pregunta[0]
        df_mes = df[df["_mes"] == mes_pedido]

        if df_mes.empty:
            respuesta = "No encontré datos para **{}**. Los meses disponibles son: {}.".format(
                mes_pedido, ", ".join(meses_disponibles))
        else:
            mm_mes = metricas(df_mes)
            top_prods_mes = grp(df_mes, "publicacion", top=5)
            top_prov_mes  = grp(df_mes, "provincia", top=3)

            respuesta = (
                "📅 **{}** — resumen del mes\n\n".format(mes_pedido) +
                "**Ventas:** {} órdenes · **Ingresos:** {} · **Unidades:** {}\n".format(
                    mm_mes["n_ventas"], fmt_ars(mm_mes["ingresos"]), int(mm_mes["unidades"])) +
                "**Ticket promedio:** {} · **Tasa entrega:** {}%\n\n".format(
                    fmt_ars(mm_mes["ticket_prom"]), mm_mes["tasa_ok"])
            )
            if top_prods_mes:
                respuesta += "**Top productos:**\n"
                for i,r in enumerate(top_prods_mes[:3],1):
                    respuesta += "{}. {} — {} ({} uds)\n".format(
                        i, r["label"][:35], fmt_ars(r["ingresos"]), int(r.get("unidades",0)))

            # Comparar con mes anterior si existe
            idx_mes = meses_disponibles.index(mes_pedido) if mes_pedido in meses_disponibles else -1
            if idx_mes > 0:
                mes_ant = meses_disponibles[idx_mes - 1]
                df_ant  = df[df["_mes"] == mes_ant]
                mm_ant  = metricas(df_ant)
                vp = var_pct(mm_mes["ingresos"], mm_ant["ingresos"])
                respuesta += "\n{} vs {}: {} en ingresos".format(mes_pedido, mes_ant, flecha(vp))

            met = metrica_pedida
            lbl_met = {"ingresos":"Ingresos","unidades":"Unidades","n_ventas":"Ventas"}.get(met,"Ingresos")
            g = {
                "tipo": "bar",
                "titulo": "Top productos en {} — {}".format(mes_pedido, lbl_met),
                "labels": [r["label"][:28] for r in top_prods_mes[:8]],
                "values": [r["ingresos"] if met=="ingresos" else r.get("unidades",0) for r in top_prods_mes[:8]],
                "values2": [r.get("unidades",0) if met=="ingresos" else r["ingresos"] for r in top_prods_mes[:8]],
                "label_series": lbl_met, "label_series2": "Unidades" if met=="ingresos" else "Ingresos",
                "color": "#2563eb"
            }

    # ══════════════════════════════════════════════════════════════════════
    # CASO E: TOP / RANKING con métrica específica
    # ══════════════════════════════════════════════════════════════════════
    elif es_top:
        met  = metrica_pedida
        lbl_met = {"ingresos":"Ingresos","unidades":"Unidades","n_ventas":"Ventas","ticket_prom":"Ticket prom","costo":"Costos"}.get(met,"Ingresos")
        fmt_v = (lambda v: fmt_ars(v)) if met in ["ingresos","costo","ticket_prom"] else (lambda v: str(int(v)))
        peor = any(x in p for x in ["peor","menor","menos","últimos","ultimos","bajo"])

        top = grp(df, "publicacion", top=10)
        if met == "unidades":
            top_sorted = sorted(top, key=lambda x: x.get("unidades",0), reverse=not peor)
        elif met == "n_ventas":
            top_sorted = sorted(top, key=lambda x: x.get("cantidad",0), reverse=not peor)
        else:
            top_sorted = sorted(top, key=lambda x: x["ingresos"], reverse=not peor)

        prefix = "Peor" if peor else "Top"
        respuesta = "🏆 **{} productos por {}**\n\n".format(prefix, lbl_met.lower())
        for i, r in enumerate(top_sorted[:8], 1):
            val = r.get("unidades",0) if met=="unidades" else r.get("cantidad",0) if met=="n_ventas" else r["ingresos"]
            respuesta += "**{}. {}** — {}\n".format(i, r["label"][:38], fmt_v(val))

        g = {
            "tipo": "bar",
            "titulo": "{} productos por {}".format(prefix, lbl_met),
            "labels": [r["label"][:28] for r in top_sorted[:8]],
            "values": [r.get("unidades",0) if met=="unidades" else r.get("cantidad",0) if met=="n_ventas" else r["ingresos"] for r in top_sorted[:8]],
            "values2": [r["ingresos"] if met!="ingresos" else r.get("unidades",0) for r in top_sorted[:8]],
            "label_series": lbl_met, "label_series2": "Ingresos" if met!="ingresos" else "Unidades",
            "color": "#16a34a" if not peor else "#dc2626"
        }

    # ══════════════════════════════════════════════════════════════════════
    # CASO F: Evolución temporal general
    # ══════════════════════════════════════════════════════════════════════
    elif es_evolucion or any(x in p for x in ["evolucion","tendencia","historico","tiempo","mes a mes","como fue","dias"]):
        met = metrica_pedida
        periodo_t = "mes" if m["n_ventas"] > 300 else "dia"
        tiempo = por_tiempo(df, periodo_t)
        lbl_met = {"ingresos":"Ingresos","unidades":"Unidades","n_ventas":"Ventas"}.get(met,"Ingresos")

        if not tiempo:
            respuesta = "No hay suficientes datos temporales para mostrar evolución."
        else:
            t_show = tiempo[-18:]
            vals = [r.get(met, r["ingresos"]) for r in t_show]
            max_v = max(vals); min_v = min(vals)
            max_p = t_show[vals.index(max_v)]["periodo"]
            min_p = t_show[vals.index(min_v)]["periodo"]
            ultimos = vals[-3:] if len(vals) >= 3 else vals
            if len(ultimos) >= 2 and ultimos[-1] > ultimos[0]:
                tend = "📈 Tendencia reciente al alza."
            elif len(ultimos) >= 2 and ultimos[-1] < ultimos[0]:
                tend = "📉 Tendencia reciente a la baja."
            else:
                tend = "➡️ Tendencia reciente estable."
            respuesta = (
                "📈 **Evolución de {} por {}**\n\n".format(lbl_met.lower(), periodo_t) +
                "**Mejor período:** {} ({} {})\n".format(max_p, "$ {:,}".format(int(max_v)).replace(",",".") if met in ["ingresos","costo"] else str(int(max_v)), lbl_met.lower()) +
                "**Período más bajo:** {} ({} {})\n\n".format(min_p, "$ {:,}".format(int(min_v)).replace(",",".") if met in ["ingresos","costo"] else str(int(min_v)), lbl_met.lower()) +
                tend
            )
            g = {
                "tipo": "line",
                "titulo": "Evolución de {} por {}".format(lbl_met, periodo_t),
                "labels": [r["periodo"] for r in t_show],
                "values": vals,
                "values2": [r.get("unidades",0) for r in t_show] if met == "ingresos" else [],
                "label_series": lbl_met, "label_series2": "Unidades" if met == "ingresos" else "",
                "color": "#2563eb"
            }

    # ══════════════════════════════════════════════════════════════════════
    # CASO G: Comparación mes anterior (sin meses explícitos)
    # ══════════════════════════════════════════════════════════════════════
    elif any(x in p for x in ["mes pasado","mes anterior","comparar","mes a mes","vendi mas","vendi menos","crecio","crecí","bajo","subi"]):
        md    = meses_data()
        meses = sorted(md.keys())
        if len(meses) < 2:
            mes = meses[0] if meses else "sin datos"
            mm_u = md.get(mes, m)
            respuesta = (
                "📅 Solo tengo datos de **{}**. Carga otro mes para comparar.\n\n".format(mes) +
                "{} ventas · {} ingresos · ticket prom {}".format(
                    mm_u["n_ventas"], fmt_ars(mm_u["ingresos"]), fmt_ars(mm_u["ticket_prom"]))
            )
        else:
            m_act = meses[-1]; m_ant = meses[-2]
            da = md[m_act]; db = md[m_ant]
            v_ing = var_pct(da["ingresos"], db["ingresos"])
            v_ven = var_pct(da["n_ventas"], db["n_ventas"])
            v_uds = var_pct(da["unidades"], db["unidades"])
            respuesta = (
                "📊 **{} vs {}**\n\n".format(m_act, m_ant) +
                "**Ingresos:** {} {} vs {}\n".format(fmt_ars(da["ingresos"]), flecha(v_ing), fmt_ars(db["ingresos"])) +
                "**Ventas:** {} {} vs {}\n".format(da["n_ventas"], flecha(v_ven), db["n_ventas"]) +
                "**Unidades:** {} {} vs {}\n\n".format(int(da["unidades"]), flecha(v_uds), int(db["unidades"]))
            )
            if (v_ing or 0) > 5: respuesta += "🚀 Creciste {}% en ingresos.".format(v_ing)
            elif (v_ing or 0) < -5: respuesta += "⚠️ Caída de {}% en ingresos.".format(abs(v_ing or 0))
            else: respuesta += "➡️ Ingresos similares al mes anterior."
            g = {
                "tipo": "bar", "titulo": "Comparación mensual — Ingresos",
                "labels": meses, "values": [md[ms]["ingresos"] for ms in meses],
                "values2": [md[ms]["n_ventas"] for ms in meses],
                "label_series": "Ingresos ($)", "label_series2": "Ventas",
                "color": "#2563eb"
            }

    # ══════════════════════════════════════════════════════════════════════
    # CASO H: Distribución / porcentajes
    # ══════════════════════════════════════════════════════════════════════
    elif es_distribucion:
        if any(x in p for x in ["envio","flex","logistic"]):
            top_env = grp(df, "forma_envio", top=8) if "forma_envio" in df.columns else []
            g = {"tipo":"doughnut","titulo":"Distribución por tipo de envío",
                 "labels":[r["label"] for r in top_env],"values":[r.get("cantidad",0) for r in top_env],
                 "values2":[],"label_series":"Envíos","label_series2":"","color":"#0d9488"}
            respuesta = "📊 Distribución de envíos:\n\n" + "\n".join(
                "**{}:** {}% ({} envíos)".format(r["label"], r.get("pct",0), r.get("cantidad",0)) for r in top_env[:6])
        elif any(x in p for x in ["estado","cancelado","entregado"]):
            est = grp(df, "estado")
            g = {"tipo":"doughnut","titulo":"Distribución por estado de venta",
                 "labels":[r["label"] for r in est],"values":[r.get("cantidad",0) for r in est],
                 "values2":[],"label_series":"Cantidad","label_series2":"","color":"#7c3aed"}
            respuesta = "📊 Distribución por estado:\n\n" + "\n".join(
                "**{}:** {}% ({})".format(r["label"], r.get("pct",0), r.get("cantidad",0)) for r in est)
        else:
            top = grp(df, "publicacion", top=8)
            g = {"tipo":"doughnut","titulo":"Distribución de ingresos por producto",
                 "labels":[r["label"][:25] for r in top],"values":[r["ingresos"] for r in top],
                 "values2":[],"label_series":"Ingresos","label_series2":"","color":"#2563eb"}
            respuesta = "📊 Distribución de ingresos por producto:\n\n" + "\n".join(
                "**{}:** {}% ({})".format(r["label"][:35], r.get("pct",0), fmt_ars(r["ingresos"])) for r in top[:6])

    # ══════════════════════════════════════════════════════════════════════
    # CASO I: Costos / margen
    # ══════════════════════════════════════════════════════════════════════
    elif any(x in p for x in ["costo","gasto","margen","ganancia","neto","comision"]):
        ing = m["ingresos"]; cost = m["costo"]; neto = m["total_neto"]
        margen_pct = round(neto / ing * 100, 1) if ing else 0
        respuesta = (
            "💰 **Análisis de costos**\n\n" +
            "**Ingresos brutos:** {}\n".format(fmt_ars(ing)) +
            "**Costos (envío + cargos ML):** {}\n".format(fmt_ars(cost)) +
            "**Ingreso neto:** {}\n".format(fmt_ars(neto)) +
            "**Margen neto estimado:** {}%\n\n".format(margen_pct) +
            ("⚠️ Margen bajo del 50%. Evaluá ajustar precios." if margen_pct < 50
             else "📊 Margen razonable." if margen_pct < 70
             else "✅ Buen margen controlado.")
        )
        g = {"tipo":"doughnut","titulo":"Ingresos vs Costos",
             "labels":["Ingreso neto","Costos"],"values":[round(neto,2),round(cost,2)],
             "values2":[],"label_series":"Distribución","label_series2":"","color":"#16a34a"}

    # ══════════════════════════════════════════════════════════════════════
    # CASO J: Resumen / fallback
    # ══════════════════════════════════════════════════════════════════════
    else:
        md = meses_data(); meses_disp = sorted(md.keys())
        top3 = grp(df, "publicacion", top=3)
        top1p = grp(df, "provincia", top=1)
        respuesta = (
            "🐾 **Resumen de tus ventas**\n\n" +
            "**{} órdenes** · **{} unidades** · **{}** en ingresos\n".format(
                m["n_ventas"], int(m["unidades"]), fmt_ars(m["ingresos"])) +
            "Ticket prom: **{}** · Tasa entrega: **{}%**\n\n".format(
                fmt_ars(m["ticket_prom"]), m["tasa_ok"])
        )
        if top3: respuesta += "🏆 Top: **{}** ({})\n".format(top3[0]["label"][:40], fmt_ars(top3[0]["ingresos"]))
        if top1p: respuesta += "📍 Zona líder: **{}** ({}%)\n".format(top1p[0]["label"], top1p[0].get("pct",0))
        if len(meses_disp) >= 2:
            vp = var_pct(md[meses_disp[-1]]["ingresos"], md[meses_disp[-2]]["ingresos"])
            respuesta += "📅 vs mes anterior: {}\n".format(flecha(vp))
        respuesta += (
            "\n💡 Ejemplos de lo que podés preguntarme:\n" +
            "• Compará Camiseta XL en enero vs febrero\n" +
            "• Top 5 productos por unidades vendidas\n" +
            "• ¿Cómo le fue a Buenos Aires este mes?\n" +
            "• Evolución de ingresos de marzo\n" +
            "• ¿Cuánto vendí en enero?"
        )
        periodo_t = "mes" if m["n_ventas"] > 300 else "dia"
        tiempo = por_tiempo(df, periodo_t)
        if tiempo:
            g = {"tipo":"line","titulo":"Evolución de ingresos","labels":[r["periodo"] for r in tiempo[-12:]],
                 "values":[r["ingresos"] for r in tiempo[-12:]],"values2":[],
                 "label_series":"Ingresos","label_series2":"","color":"#2563eb"}

    return ok({"resultado": {"respuesta": respuesta, "grafico": norm_grafico(g)}})


# ═══════════════════════════════════════════════════════════════════════════
# PUBLICIDAD — ML Ads
# ═══════════════════════════════════════════════════════════════════════════

def _pub_ncol(col):
    """Normaliza nombre de columna: quita saltos de línea, espacios extra, minúsculas."""
    import re as _re
    return _re.sub(r'\s+', ' ', str(col).replace('\n', ' ')).strip().lower()

def _pub_num(v):
    """Convierte valor a float tolerando '-', None, NaN."""
    if v is None: return None
    s = str(v).strip()
    if s in ('', '-', 'nan', 'None', 'N/A', 'n/a'): return None
    s2 = re.sub(r'[^\d,.\-]', '', s)
    if ',' in s2:
        s2 = s2.replace('.', '').replace(',', '.')
    else:
        parts = s2.split('.')
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            s2 = s2.replace('.', '')
    try:
        f = float(s2)
        return None if (math.isnan(f) or math.isinf(f)) else f
    except: return None

def _pub_parse_mes(v):
    """Convierte 'Desde' a 'YYYY-MM'. Maneja datetime, YYYY-MM-DD, DD-mes-YYYY y DD-abrev-YYYY."""
    import datetime as _dt
    # ya es datetime
    if isinstance(v, (_dt.datetime, _dt.date)):
        return f"{v.year}-{v.month:02d}"
    s = str(v).strip()
    # YYYY-MM-DD
    m = re.match(r'(\d{4})-(\d{2})', s)
    if m: return f"{m.group(1)}-{m.group(2)}"
    # DD-mes-YYYY  (soporta nombres completos Y abreviados: ene, feb, mar, abr...)
    meses_full = {"enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
                  "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12}
    meses_abbr = {"ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,
                  "jul":7,"ago":8,"sep":9,"oct":10,"nov":11,"dic":12}
    m = re.match(r'\d{1,2}-(\w+)-(\d{4})', s)
    if m:
        mes_str = m.group(1).lower()
        anio    = m.group(2)
        mo = meses_full.get(mes_str) or meses_abbr.get(mes_str)
        if mo: return f"{anio}-{mo:02d}"
    return s[:7] if len(s) >= 7 else s

def _pub_read(raw: bytes):
    """
    Lee un Excel de publicidad ML.
    Devuelve (tipo, df_normalizado) donde tipo es 'campanias' o 'anuncios'.
    """
    xl = pd.ExcelFile(io.BytesIO(raw))

    # 1) encontrar hoja de datos
    hoja = None
    for name in xl.sheet_names:
        nl = name.lower()
        if 'reporte' in nl:
            hoja = name; break
    if not hoja:
        hoja = xl.sheet_names[-1]

    # 2) leer con header en fila 1 (fila 0 = título, fila 1 = columnas)
    df = pd.read_excel(xl, sheet_name=hoja, header=1, dtype=str)
    df = df.dropna(how='all').reset_index(drop=True)
    xl.close()

    # 3) normalizar nombres de columna (quitar \n y espacios extra)
    df.columns = [_pub_ncol(c) for c in df.columns]

    # 4) detectar tipo por columnas clave
    cols = list(df.columns)
    es_anuncios = any('título' in c or 'titulo' in c or 'número de' in c for c in cols)
    tipo = 'anuncios' if es_anuncios else 'campanias'

    # 5) mapa de renombrado según tipo
    if tipo == 'campanias':
        rename_map = {
            'nombre de campaña':  'campana',
            'estado':             'estado',
            'presupuesto':        'presupuesto',
            'roas objetivo':      'roas_objetivo',
            'desde':              'desde',
            'hasta':              'hasta',
            'impresiones':        'impresiones',
            'clics':              'clics',
            'ventas directas':    'ventas_directas',
            'ventas indirectas':  'ventas_indirectas',
            'unidades vendidas por publicidad': 'unidades',
        }
        # columnas con substrings variables
        for c in cols:
            if c in rename_map: continue
            if c.startswith('cpc'):   rename_map[c] = 'cpc'
            elif c.startswith('ctr'): rename_map[c] = 'ctr'
            elif c.startswith('cvr'): rename_map[c] = 'cvr'
            elif c.startswith('acos') and 'objetivo' not in c: rename_map[c] = 'acos'
            elif c.startswith('roas') and 'objetivo' not in c: rename_map[c] = 'roas'
            elif 'ingresos' in c and 'directas' not in c and 'indirectas' not in c: rename_map[c] = 'ingresos'
            elif 'invers' in c: rename_map[c] = 'inversion'
            elif 'ventas por publicidad' in c: rename_map[c] = 'ventas_total'
            elif '% de impresiones ganadas' in c and 'primeros' not in c: rename_map[c] = 'pct_imp_ganadas'
            elif 'perdidas por presupuesto' in c: rename_map[c] = 'pct_perdidas_ppto'
            elif 'perdidas por ranking' in c:     rename_map[c] = 'pct_perdidas_rank'
    else:
        rename_map = {
            'campaña':            'campana',
            'desde':              'desde',
            'hasta':              'hasta',
            'estado':             'estado',
            'impresiones':        'impresiones',
            'clics':              'clics',
            'ventas directas':    'ventas_directas',
            'ventas indirectas':  'ventas_indirectas',
        }
        for c in cols:
            if c in rename_map: continue
            if ('título de la publicaci' in c or 'titulo de la publicaci' in c) and 'vendida' in c:
                rename_map[c] = 'titulo_vendido'
            elif ('número de la publicaci' in c or 'numero de la publicaci' in c) and 'vendida' in c:
                rename_map[c] = 'pub_id_vendido'
            elif 'estado del anuncio' in c:
                rename_map[c] = 'estado'
            elif 'título' in c or 'titulo' in c:
                rename_map[c] = 'titulo'
            elif 'número de' in c or 'numero de' in c:
                rename_map[c] = 'pub_id'
            elif c.startswith('cpc'):   rename_map[c] = 'cpc'
            elif c.startswith('ctr'):   rename_map[c] = 'ctr'
            elif c.startswith('cvr'):   rename_map[c] = 'cvr'
            elif c.startswith('acos'):  rename_map[c] = 'acos'
            elif c.startswith('roas'):  rename_map[c] = 'roas'
            elif 'ingresos' in c and 'directas' not in c and 'indirectas' not in c: rename_map[c] = 'ingresos'
            elif 'invers' in c: rename_map[c] = 'inversion'
            elif 'ventas por publicidad' in c: rename_map[c] = 'ventas_total'

    df = df.rename(columns=rename_map)
    df = df.loc[:, ~df.columns.duplicated()]
    if 'titulo' not in df.columns and 'titulo_vendido' in df.columns:
        df['titulo'] = df['titulo_vendido']

    # 6) convertir numéricos
    num_cols = ['impresiones','clics','ingresos','inversion','cpc','ctr','cvr',
                'acos','roas','ventas_total','ventas_directas','ventas_indirectas',
                'unidades','presupuesto','pct_imp_ganadas','pct_perdidas_ppto','pct_perdidas_rank']
    for col in num_cols:
        if col in df.columns:
            df[col] = df[col].apply(_pub_num)

    # 7) mes desde "Desde"
    if 'desde' in df.columns:
        df['mes'] = df['desde'].apply(_pub_parse_mes)

    return tipo, df


@app.route('/api/publicidad/upload', methods=['POST'])
def api_pub_upload():
    try:
        f = request.files.get('file')
        if not f:
            return err('No se recibió archivo')
        raw  = f.read()
        fname = f.filename or 'archivo.xlsx'
        tipo_hint = request.form.get('tipo_hint', '')  # 'campanias' o 'anuncios'

        tipo, df = _pub_read(raw)

        # si el usuario forzó el tipo, respetarlo
        if tipo_hint in ('campanias', 'anuncios'):
            tipo = tipo_hint

        has_ads_metrics = any(c in df.columns for c in ('inversion', 'clics', 'impresiones', 'acos'))
        has_sales_metrics = 'ingresos' in df.columns and 'ventas_total' in df.columns
        if tipo == 'anuncios' and has_sales_metrics and not has_ads_metrics:
            tipo = 'ventas_ads'

        if df.empty:
            return err('El archivo está vacío o no se pudo leer correctamente')

        if tipo == 'campanias':
            PUB_STORE['campanias'] = df
        elif tipo == 'ventas_ads':
            PUB_STORE['ventas_ads'] = df
        else:
            PUB_STORE['anuncios'] = df

        existing = [x for x in PUB_STORE['files'] if x['tipo'] != tipo]
        existing.append({'tipo': tipo, 'nombre': fname, 'filas': len(df)})
        PUB_STORE['files'] = existing

        return ok({'tipo': tipo, 'filas': len(df), 'nombre': fname,
                   'files': PUB_STORE['files']})
    except Exception as e:
        import traceback
        return err(f'Error procesando archivo: {str(e)}\n{traceback.format_exc()}')


@app.route('/api/publicidad/dashboard', methods=['GET'])
def api_pub_dashboard():
    try:
        dc = PUB_STORE['campanias']
        da = PUB_STORE['anuncios']
        dv = PUB_STORE.get('ventas_ads')
        if dc is None and da is None and dv is None:
            return err('No hay datos de publicidad cargados')

        def s(df, col):
            if df is None or col not in df.columns: return 0.0
            return float(df[col].dropna().astype(float).sum())

        # ── KPIs globales ────────────────────────────────────────────────
        base = dc if dc is not None else (da if da is not None else dv)
        ingresos    = s(base, 'ingresos')
        inversion   = s(base, 'inversion')
        clics       = int(s(base, 'clics'))
        impresiones = int(s(base, 'impresiones'))
        ventas      = int(s(base, 'ventas_total'))
        unidades    = int(s(base, 'unidades'))
        ganancia    = round(ingresos - inversion, 2)
        has_inv_base = base is not None and 'inversion' in base.columns
        acos = round(inversion / ingresos * 100, 1) if has_inv_base and ingresos and inversion else None
        roas = round(ingresos / inversion, 2)        if has_inv_base and inversion else None

        result = {'kpis': {
            'ingresos': round(ingresos,2), 'inversion': round(inversion,2),
            'ganancia': ganancia, 'acos': acos, 'roas': roas,
            'clics': clics, 'impresiones': impresiones,
            'ventas': ventas, 'unidades': unidades,
        }}

        # ── Mensual ──────────────────────────────────────────────────────
        meses_data = {}
        if base is not None and 'mes' in base.columns:
            for mes, gdf in base.groupby('mes'):
                ing = float(gdf['ingresos'].dropna().astype(float).sum()) if 'ingresos' in gdf else 0
                inv = float(gdf['inversion'].dropna().astype(float).sum()) if 'inversion' in gdf else 0
                cl  = int(gdf['clics'].dropna().astype(float).sum())       if 'clics'    in gdf else 0
                vt  = int(gdf['ventas_total'].dropna().astype(float).sum()) if 'ventas_total' in gdf else 0
                has_inv = 'inversion' in gdf.columns
                meses_data[str(mes)] = {
                    'ingresos': round(ing,2), 'inversion': round(inv,2),
                    'ganancia': round(ing-inv,2),
                    'acos': round(inv/ing*100,1) if has_inv and ing and inv else None,
                    'roas': round(ing/inv,2)      if has_inv and inv else None,
                    'clics': cl, 'ventas': vt,
                }
        result['mensual'] = dict(sorted(meses_data.items()))

        # ── Por campaña ──────────────────────────────────────────────────
        por_campana = []
        if base is not None and 'campana' in base.columns:
            for camp, gdf in base.groupby('campana'):
                ing = float(gdf['ingresos'].dropna().astype(float).sum()) if 'ingresos' in gdf else 0
                inv = float(gdf['inversion'].dropna().astype(float).sum()) if 'inversion' in gdf else 0
                cl  = int(gdf['clics'].dropna().astype(float).sum())       if 'clics'    in gdf else 0
                vt  = int(gdf['ventas_total'].dropna().astype(float).sum()) if 'ventas_total' in gdf else 0
                imp = int(gdf['impresiones'].dropna().astype(float).sum())  if 'impresiones' in gdf else 0
                has_inv = 'inversion' in gdf.columns
                est_series = gdf['estado'].dropna() if 'estado' in gdf.columns else pd.Series([])
                est = str(est_series.iloc[0]) if len(est_series) else '-'
                por_campana.append({
                    'campana': str(camp), 'estado': est,
                    'ingresos': round(ing,2), 'inversion': round(inv,2),
                    'ganancia': round(ing-inv,2),
                    'acos': round(inv/ing*100,1) if has_inv and ing and inv else None,
                    'roas': round(ing/inv,2)      if has_inv and inv else None,
                    'clics': cl, 'ventas': vt, 'impresiones': imp,
                })
        result['por_campana'] = sorted(por_campana, key=lambda x: x['ingresos'], reverse=True)

        # ── Detalle mensual por campaña (para filtro por mes) ─────────────
        mensual_detalle = {}
        if base is not None and 'mes' in base.columns and 'campana' in base.columns:
            for mes_val, gdf_mes in base.groupby('mes'):
                mes_camps = []
                for camp, gdf_c in gdf_mes.groupby('campana'):
                    ing = float(gdf_c['ingresos'].dropna().astype(float).sum()) if 'ingresos' in gdf_c else 0
                    inv = float(gdf_c['inversion'].dropna().astype(float).sum()) if 'inversion' in gdf_c else 0
                    cl  = int(gdf_c['clics'].dropna().astype(float).sum())       if 'clics'    in gdf_c else 0
                    vt  = int(gdf_c['ventas_total'].dropna().astype(float).sum()) if 'ventas_total' in gdf_c else 0
                    imp = int(gdf_c['impresiones'].dropna().astype(float).sum())  if 'impresiones' in gdf_c else 0
                    has_inv = 'inversion' in gdf_c.columns
                    est_s = gdf_c['estado'].dropna() if 'estado' in gdf_c.columns else pd.Series([])
                    est = str(est_s.iloc[0]) if len(est_s) else '-'
                    mes_camps.append({
                        'campana': str(camp), 'estado': est,
                        'ingresos': round(ing,2), 'inversion': round(inv,2),
                        'ganancia': round(ing-inv,2),
                        'acos': round(inv/ing*100,1) if has_inv and ing and inv else None,
                        'roas': round(ing/inv,2)      if has_inv and inv else None,
                        'clics': cl, 'ventas': vt, 'impresiones': imp,
                    })
                mensual_detalle[str(mes_val)] = sorted(mes_camps, key=lambda x: x['ingresos'], reverse=True)
        result['mensual_detalle'] = mensual_detalle

        # ── Top anuncios ─────────────────────────────────────────────────
        top_anuncios = []
        anuncios_por_mes = {}
        top_src = da if da is not None else dv
        if top_src is not None and 'titulo' in top_src.columns:
            num_cols_an = [c for c in ['ingresos','inversion','clics','ventas_total','impresiones'] if c in top_src.columns]
            for col in num_cols_an:
                top_src[col] = top_src[col].apply(_pub_num)

            # global top
            grp = top_src.groupby('titulo')[num_cols_an].sum(numeric_only=True).reset_index()
            grp = grp.sort_values('ingresos', ascending=False).head(15)
            for _, row in grp.iterrows():
                ing = float(row['ingresos']) if 'ingresos' in row else 0
                inv = float(row['inversion']) if 'inversion' in row.index else None
                top_anuncios.append({
                    'titulo':      str(row['titulo'])[:65],
                    'ingresos':    round(ing,2),
                    'inversion':   round(inv,2) if inv is not None else None,
                    'acos':        round(inv/ing*100,1) if inv is not None and ing and inv else None,
                    'roas':        round(ing/inv,2)      if inv else None,
                    'clics':       int(row['clics'])       if 'clics'       in row.index else 0,
                    'ventas':      int(row['ventas_total']) if 'ventas_total' in row.index else 0,
                    'impresiones': int(row['impresiones'])  if 'impresiones'  in row.index else 0,
                })

            # por mes
            if 'mes' in top_src.columns:
                for mes_val, gdf_m in top_src.groupby('mes'):
                    grp_m = gdf_m.groupby('titulo')[num_cols_an].sum(numeric_only=True).reset_index()
                    grp_m = grp_m.sort_values('ingresos', ascending=False).head(15)
                    mes_an = []
                    for _, row in grp_m.iterrows():
                        ing = float(row['ingresos']) if 'ingresos' in row else 0
                        inv = float(row['inversion']) if 'inversion' in row.index else None
                        if ing == 0 and (inv is None or inv == 0): continue
                        mes_an.append({
                            'titulo':      str(row['titulo'])[:65],
                            'ingresos':    round(ing,2),
                            'inversion':   round(inv,2) if inv is not None else None,
                            'acos':        round(inv/ing*100,1) if inv is not None and ing and inv else None,
                            'roas':        round(ing/inv,2)      if inv else None,
                            'clics':       int(row['clics'])       if 'clics'       in row.index else 0,
                            'ventas':      int(row['ventas_total']) if 'ventas_total' in row.index else 0,
                            'impresiones': int(row['impresiones'])  if 'impresiones'  in row.index else 0,
                        })
                    if mes_an:
                        anuncios_por_mes[str(mes_val)] = mes_an
        result['top_anuncios']   = top_anuncios
        result['anuncios_por_mes'] = anuncios_por_mes
        result['files'] = PUB_STORE['files']
        return ok(result)
    except Exception as e:
        import traceback
        return err(f'Error generando dashboard: {str(e)}\n{traceback.format_exc()}')



# ═══════════════════════════════════════════════════════════════════════════
# MÓDULO PUBLICACIONES — Upload + Dashboard
# ═══════════════════════════════════════════════════════════════════════════

PUBS_STORE = {"df": None, "filename": None}

def _parse_pubs_file(raw: bytes, filename: str) -> pd.DataFrame:
    """
    Parser para Reporte de Rendimiento de Publicaciones de MercadoLibre.
    Formato real: 5 filas de metadata, headers en fila 6, datos desde fila 7.
    Columnas: ID de la publicación, Publicación, Estado actual, Variante, SKU,
              Visitas únicas, Cantidad de ventas, Compradores únicos, Unidades vendidas,
              Ventas brutas (ARS), % de participación, Conversión de visitas a ventas,
              Conversión de visitas a compradores
    También acepta archivos CSV y Excel con otros formatos (fallback flexible).
    """
    ext = Path(filename).suffix.lower()
    buf = io.BytesIO(raw)

    df = None

    # ── Intentar formato oficial ML (header en fila 6) ──────────────────────
    if ext in (".xlsx", ".xls"):
        try:
            df_try = pd.read_excel(buf, sheet_name=0, header=5)
            df_try.columns = [str(c).strip() for c in df_try.columns]
            # Verificar que sea el formato oficial ML
            if "Publicación" in df_try.columns or "Publicacion" in df_try.columns:
                df = df_try
        except Exception:
            pass

        # Fallback: leer sin skip de filas
        if df is None:
            buf.seek(0)
            try:
                df = pd.read_excel(buf, sheet_name=0, dtype=str)
                df.columns = [str(c).strip() for c in df.columns]
            except Exception as e:
                raise ValueError(f"No se pudo leer el archivo Excel: {e}")

    elif ext == ".csv":
        for enc in ("utf-8", "latin-1", "utf-8-sig"):
            for sep in (";", ",", None):
                try:
                    buf.seek(0)
                    kwargs = dict(encoding=enc, dtype=str)
                    if sep: kwargs["sep"] = sep
                    else: kwargs["sep"] = None; kwargs["engine"] = "python"
                    df_try = pd.read_csv(buf, **kwargs)
                    if df_try.shape[1] >= 3:
                        df = df_try; break
                except Exception:
                    continue
            if df is not None: break
        if df is None:
            raise ValueError("No se pudo leer el CSV. Verificá el encoding o separador.")
    else:
        raise ValueError("Formato no compatible. Usá .xlsx, .xls o .csv")

    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)

    # ── Mapeo de columnas → alias internos ──────────────────────────────────
    # Busca por nombre exacto primero, luego por subcadena (case-insensitive)
    TARGETS = {
        "publicacion":   ["Publicación", "Publicacion", "Título", "Titulo",
                          "Título de la publicación", "Titulo de la publicacion", "Nombre"],
        "estado":        ["Estado actual", "Estado", "Estado de la publicación",
                          "Estado de la publicacion"],
        "visitas":       ["Visitas únicas", "Visitas unicas", "Visitas", "Visitas totales"],
        "ventas":        ["Cantidad de ventas", " Cantidad de ventas", "Ventas",
                          "Órdenes", "Ordenes"],
        "unidades":      ["Unidades vendidas", "Unidades", "Cantidad"],
        "ventas_brutas": ["Ventas brutas (ARS)", "Ventas brutas (ARS) ", "Ventas brutas",
                          "Ingresos", "Total ventas", "Total"],
        "conv":          ["Conversión de visitas a ventas", "Conversion de visitas a ventas",
                          "Conversión", "Conversion", "CVR", "Tasa de conversión"],
        "participacion": ["% de participación", "% de participacion", "Participación"],
        "id_pub":        ["ID de la publicación", "ID de la publicacion", "ID"],
        "compradores":   ["Compradores únicos", "Compradores unicos", "Compradores"],
        "variante":      ["Variante"],
        "sku":           ["SKU", "Sku"],
    }

    cols_lower = {c.lower().strip(): c for c in df.columns}

    for alias, candidates in TARGETS.items():
        if alias in df.columns:
            continue
        for cand in candidates:
            if cand in df.columns:
                df[alias] = df[cand]; break
            if cand.lower().strip() in cols_lower:
                df[alias] = df[cols_lower[cand.lower().strip()]]; break

    # ── Limpiar numéricos ────────────────────────────────────────────────────
    # ML exporta Ventas brutas en miles de ARS (e.g. 62.633 = $62,633 ARS)
    # El separador de miles es punto, la coma es separador decimal
    # Conversión viene como "1,99%" (coma decimal, sin punto de miles)

    def _parse_ml_money(v):
        """62.633 → 62633  |  61.75 → 61750  |  49.99 → 49990  |  58 → 58000"""
        if pd.isna(v):
            return None
        s = str(v).strip()
        s = re.sub(r"[^\d,.\-]", "", s)
        if not s:
            return None
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        else:
            parts = s.split(".")
            if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
                s = s.replace(".", "")
        try:
            return float(s)
        except Exception:
            return None

    def _parse_pct(v):
        """'1,99%' → 1.99  |  '33,33%' → 33.33  |  '25%' → 25.0"""
        if pd.isna(v): return None
        s = str(v).strip().rstrip("%").strip().replace(",", ".")
        try: return float(s)
        except Exception: return None

    def _parse_int(v):
        if pd.isna(v): return None
        try: return int(float(str(v).replace(",", ".")))
        except Exception: return None

    for col in ["visitas", "ventas", "unidades", "compradores"]:
        if col in df.columns:
            df[col] = df[col].apply(_parse_int)

    if "ventas_brutas" in df.columns:
        df["ventas_brutas"] = df["ventas_brutas"].apply(_parse_ml_money)

    if "conv" in df.columns:
        df["conv"] = df["conv"].apply(_parse_pct)

    for col in ["visitas", "ventas", "unidades", "compradores", "ventas_brutas", "conv"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # ── Normalizar estado ────────────────────────────────────────────────────
    if "estado" in df.columns:
        df["estado"] = df["estado"].apply(
            lambda v: "INACTIVA"
            if any(x in str(v).lower() for x in ("inact", "pausad", "cerrad", "suspendid"))
            else "ACTIVA")
    else:
        df["estado"] = "ACTIVA"

    # ── Calcular CVR si falta o es todo NaN ─────────────────────────────────
    if "conv" not in df.columns or df["conv"].isna().all():
        if "ventas" in df.columns and "visitas" in df.columns:
            df["conv"] = (
                df["ventas"].fillna(0)
                / df["visitas"].replace(0, float("nan"))
                * 100
            ).round(2)

    if "publicacion" not in df.columns:
        raise ValueError(
            "No se encontró la columna de nombre de publicación. "
            "Verificá que sea el reporte de Rendimiento de Publicaciones de MercadoLibre."
        )

    return df


@app.route("/api/publicaciones/upload", methods=["POST"])
def api_pubs_upload():
    if "file" not in request.files:
        return err("Sin archivo.")
    f = request.files["file"]
    raw = f.read()
    try:
        df = _parse_pubs_file(raw, f.filename)
    except Exception as e:
        return err(str(e))

    PUBS_STORE["df"] = df
    PUBS_STORE["filename"] = f.filename

    return _build_pubs_response(df, f.filename)


@app.route("/api/publicaciones/dashboard", methods=["GET"])
def api_pubs_dashboard():
    df = PUBS_STORE.get("df")
    fn = PUBS_STORE.get("filename") or "archivo"
    if df is None or df.empty:
        return err("No hay datos de publicaciones cargados.")
    return _build_pubs_response(df, fn)


def _build_pubs_response(df: pd.DataFrame, filename: str):
    """Construye la respuesta JSON con métricas y tablas del dashboard."""
    def safe_row(r: dict) -> dict:
        """Convierte NaN a None para JSON."""
        return {k: (None if (isinstance(v, float) and (v != v)) else v)
                for k, v in r.items()}

    total_pubs   = len(df)
    activas      = int((df["estado"] == "ACTIVA").sum()) if "estado" in df.columns else 0
    inactivas    = total_pubs - activas
    visitas_tot  = float(df["visitas"].sum()) if "visitas" in df.columns else 0
    vb_tot       = float(df["ventas_brutas"].sum()) if "ventas_brutas" in df.columns else 0
    transac      = float(df["ventas"].sum()) if "ventas" in df.columns else 0
    conv_global  = round(transac / visitas_tot * 100, 2) if visitas_tot else 0.0

    # helpers
    def _rows(sub: pd.DataFrame, cols: list) -> list:
        rows = []
        for _, r in sub.iterrows():
            row = {"nombre": str(r.get("publicacion",""))[:60],
                   "estado": str(r.get("estado","INACTIVA"))}
            for c in cols:
                v = r.get(c)
                row[c] = round(float(v), 2) if pd.notna(v) else None
            rows.append(safe_row(row))
        return rows

    # Top ventas brutas
    col_vb = "ventas_brutas" if "ventas_brutas" in df.columns else None
    if col_vb:
        top_v = df.nlargest(15, col_vb)
    else:
        top_v = df.head(15)
    top_ventas = _rows(top_v, ["visitas","ventas","unidades","ventas_brutas","conv"])

    # Top visitas
    if "visitas" in df.columns:
        top_vis = df.nlargest(10, "visitas")
    else:
        top_vis = df.head(10)
    top_visitas = _rows(top_vis, ["visitas","ventas","conv"])

    # Top CVR (mín 10 visitas)
    if "visitas" in df.columns and "conv" in df.columns:
        df_cvr = df[df["visitas"].fillna(0) >= 10].copy()
        top_cvr_df = df_cvr.nlargest(10, "conv")
    else:
        top_cvr_df = df.head(10)
    top_cvr = _rows(top_cvr_df, ["visitas","ventas","conv"])

    periodo = Path(filename).stem[:40]

    return ok({
        "total_pubs":        total_pubs,
        "activas":           activas,
        "inactivas":         inactivas,
        "visitas_total":     round(visitas_tot),
        "ventas_brutas_total": round(vb_tot),
        "transacciones":     round(transac),
        "conv_global":       conv_global,
        "periodo":           periodo,
        "top_ventas":        top_ventas,
        "top_visitas":       top_visitas,
        "top_cvr":           top_cvr,
    })


# ═══════════════════════════════════════════════════════════════════════════
# MÓDULO TENDENCIAS — Upload
# ═══════════════════════════════════════════════════════════════════════════

def _parse_precio_mkt(v):
    """'39.480,16' → 39480.16  |  '59.990' → 59990.0"""
    if v is None: return 0.0
    s = str(v).strip()
    # Formato ML: punto=miles, coma=decimal  e.g. '39.480,16'
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    else:
        # Sin coma: '59.990' — si tiene punto y 3 decimales → miles
        parts = s.split('.')
        if len(parts) == 2 and len(parts[1]) == 3:
            s = s.replace('.', '')
        # si tiene punto y != 3 decimales → decimal normal
    try: return float(s)
    except: return 0.0


def _parse_benchmark_sheet(ws) -> list:
    """Lee una hoja de benchmark ML. Header fila 6, datos desde fila 7."""
    rows_out = []
    headers = None
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 5: continue          # skip metadata
        if i == 5:                  # headers
            headers = [str(c).strip() if c else '' for c in row]
            continue
        if not any(row): continue   # fila vacía
        if not headers: continue
        r = dict(zip(headers, row))
        nombre = str(r.get('Nombre') or r.get('Nome') or '').strip()
        if not nombre or nombre in ('-', ''): continue
        rows_out.append({
            'nombre':     nombre[:80],
            'condicion':  str(r.get('Condición') or r.get('Condição') or 'Nuevo').strip(),
            'precio':     _parse_precio_mkt(r.get('Precio') or r.get('Preço')),
            'uds':        int(r.get('Unidades vendidas') or 0),
            'vistas':     int(r.get('Vistas') or r.get('Visualizações') or 0),
            'canceladas': int(r.get('Ventas canceladas') or r.get('Vendas canceladas') or 0),
            'preguntas':  int(r.get('Preguntas') or r.get('Perguntas') or 0),
            'catalogo':   'Sí' if str(r.get('Producto de catálogo') or r.get('Produto de catálogo') or '').strip() == 'Sí' else 'No',
            'video':      'Sí' if str(r.get('Publicación con video') or r.get('Anúncio com vídeo') or '').strip() == 'Sí' else 'No',
            'cuotas':     'Sí' if str(r.get('Cuotas') or r.get('Parcelas') or '').strip() == 'Sí' else 'No',
            'publicidad': 'Sí' if str(r.get('Publicidad') or r.get('Publicidade') or '').strip() == 'Sí' else 'No',
            'fotos':      int(r.get('Cantidad de fotos') or r.get('Quantidade de fotos') or 0),
            'envio':      str(r.get('Envío') or r.get('Frete') or '—').strip(),
        })
    return rows_out


@app.route('/api/tendencias/upload', methods=['POST'])
def api_tendencias_upload():
    if 'file' not in request.files:
        return err('Sin archivo.')
    f = request.files['file']
    raw = f.read()
    fname = f.filename or 'archivo.xlsx'

    # Detectar mes del nombre de archivo
    mes_match = re.search(r'(\d{4}-\d{2})-\d{2}a', fname)
    if mes_match:
        mes = mes_match.group(1)
    else:
        m2 = re.search(r'(\d{4})[-_](\d{2})', fname)
        mes = (m2.group(1)+'-'+m2.group(2)) if m2 else 'Sin fecha'

    try:
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    except Exception as e:
        return err(f'No se pudo leer el archivo: {e}')

    SKIP_SHEETS = {'Definiciones', 'Definitions', 'Definições'}
    EMPTY_TITLES = {'{CAT_L2_TITLE}','{CAT_L3_TITLE}','{CAT_L4_TITLE}',
                    '{CAT_L5_TITLE}','{CAT_L6_TITLE}','{CAT_L7_TITLE}',
                    '{CAT_L2_SUBTITLE}','#NAME?'}

    categorias = {}
    for shname in wb.sheetnames:
        if shname in SKIP_SHEETS: continue
        ws = wb[shname]
        # Verificar que no sea una hoja vacía/template
        first_row = next(ws.iter_rows(max_row=1, values_only=True), (None,))
        first_val = str(first_row[0] or '').strip()
        if first_val in EMPTY_TITLES or first_val.startswith('{CAT_'): continue
        rows = _parse_benchmark_sheet(ws)
        if rows:
            cat_name = shname if shname not in ('Sheet1','Hoja1') else f'Categoría {len(categorias)+1}'
            categorias[cat_name] = rows

    if not categorias:
        return err('No se encontraron datos en el archivo. Verificá que sea el reporte de benchmark de MercadoLibre.')

    total = sum(len(v) for v in categorias.values())
    return ok({'mes': mes, 'categorias': categorias, 'total': total,
               'cats': list(categorias.keys())})


if __name__ == "__main__":
    import threading, webbrowser
    # Solo abrir el navegador automáticamente cuando corre local (no en el servidor online)
    if HOST in ("127.0.0.1", "localhost") and not os.environ.get("DASHIFY_NO_BROWSER"):
        def open_browser():
            import time; time.sleep(1.2)
            webbrowser.open(f"http://{HOST}:{PORT}")
        threading.Thread(target=open_browser, daemon=True).start()
    app.run(host=HOST, port=PORT, debug=False)
