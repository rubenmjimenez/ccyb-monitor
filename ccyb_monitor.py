#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ccyb_monitor.py
================
Monitoriza el archivo Excel oficial que la ESRB publica con los colchones
de capital anticíclico (CCyB) de cada país:

    https://www.esrb.europa.eu/national_policy/ccb/shared/data/esrb.ccybd_CCyB_data.xlsx

El archivo es una BASE DE DATOS HISTÓRICA: cada fila es una decisión de un
país (columnas reales: 'Country', 'Decision on', 'Date of Announcement',
'CCyB rate', 'Type of setting', 'Application since', ...). Un mismo país
puede tener varias filas: unas con fecha de aplicación ya pasada (vigentes)
y otras con fecha futura (pendientes).

Cada vez que se ejecuta este script:
  1. Descarga el Excel más reciente.
  2. Calcula un hash de toda la tabla y lo compara con el de la última vez
     (data/last_snapshot.json) para saber si la ESRB ha publicado algo
     nuevo desde el último chequeo.
  3. Para cada país, calcula qué tasa está en vigor HOY (la decisión con
     'Application since' más reciente que ya haya pasado) y busca si hay
     alguna decisión pendiente cuya 'Application since' caiga dentro de la
     ventana "a un mes vista" (por defecto, los próximos 35 días).
  4. Envía SIEMPRE un email mensual con el resultado (a uno o varios
     destinatarios) y, si hay cambios a un mes vista, adjunta un script
     .sql con los UPDATE necesarios sobre la tabla
     dbo.colchon_capital_anticiclico para la fecha de cierre trimestral
     que corresponda a cada cambio.

-------------------------------------------------------------------------
LÓGICA DE LA FECHA DE CIERRE (confirmada con el usuario)
-------------------------------------------------------------------------
Las fechas de cierre trimestral son siempre el día 1 de enero, abril,
julio u octubre (cod_entidad_fecha / cod_pais_fecha usan esa fecha, no la
fecha real de entrada en vigor del cambio). Dado un cambio con fecha de
entrada en vigor D, la fecha de cierre que le corresponde es la PRIMERA
fecha de trimestre (01/01, 01/04, 01/07, 01/10) ESTRICTAMENTE POSTERIOR a
D (si D coincide exactamente con un inicio de trimestre, se usa el
siguiente, no ese mismo).

Ejemplos verificados:
  - Polonia, entra en vigor el 30/09/2026  -> cierre 01/10/2026 (cierre "septiembre26")
  - Grecia/España, entran en vigor el 01/10/2026 (coincide con inicio de
    trimestre) -> cierre 01/01/2027 (cierre "diciembre26")
"""

import argparse
import csv
import hashlib
import io
import json
import os
import smtplib
import sys
import unicodedata
from datetime import date, datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd
import requests

# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------

XLSX_URL = "https://www.esrb.europa.eu/national_policy/ccb/shared/data/esrb.ccybd_CCyB_data.xlsx"

BASE_DIR = Path(__file__).resolve().parent
SNAPSHOT_PATH = BASE_DIR / "data" / "last_snapshot.json"
ISO_CODES_PATH = BASE_DIR / "data" / "country_iso_codes.csv"

# "A un mes vista": margen en días. 35 cubre con holgura cualquier mes del año.
LOOKAHEAD_DAYS = 35

COLUMN_HINTS = {
    "country": ["country"],
    "rate": ["ccyb rate", "rate"],
    "application_date": ["application since", "applicable", "in effect since"],
    "decision_date": ["decision on"],
    "announcement_date": ["date of announcement"],
}

# Nombre (en inglés, tal como aparece en el Excel de la ESRB, en mayúsculas
# y sin acentos) -> alias a buscar en country_iso_codes.csv, para los casos
# donde el nombre no coincide literalmente.
COUNTRY_NAME_ALIASES = {
    "NETHERLANDS": "THE NETHERLANDS",
    "CZECHIA": "CZECH REPUBLIC",
    "SLOVAK REPUBLIC": "SLOVAKIA",
}

# Traducción al español, solo para que los comentarios del SQL queden
# legibles (si un país no está aquí, se usa el nombre en inglés tal cual,
# sin que el script falle).
COUNTRY_NAME_ES = {
    "AUSTRIA": "Austria", "BELGIUM": "Bélgica", "BULGARIA": "Bulgaria",
    "CROATIA": "Croacia", "CYPRUS": "Chipre", "CZECHIA": "República Checa",
    "CZECH REPUBLIC": "República Checa", "DENMARK": "Dinamarca",
    "ESTONIA": "Estonia", "FINLAND": "Finlandia", "FRANCE": "Francia",
    "GERMANY": "Alemania", "GREECE": "Grecia", "HUNGARY": "Hungría",
    "ICELAND": "Islandia", "IRELAND": "Irlanda", "ITALY": "Italia",
    "LATVIA": "Letonia", "LIECHTENSTEIN": "Liechtenstein",
    "LITHUANIA": "Lituania", "LUXEMBOURG": "Luxemburgo", "MALTA": "Malta",
    "NETHERLANDS": "Países Bajos", "NORWAY": "Noruega", "POLAND": "Polonia",
    "PORTUGAL": "Portugal", "ROMANIA": "Rumanía", "SLOVAKIA": "Eslovaquia",
    "SLOVAK REPUBLIC": "Eslovaquia", "SLOVENIA": "Eslovenia", "SPAIN": "España",
    "SWEDEN": "Suecia", "UNITED KINGDOM": "Reino Unido",
}

# Email
MAIL_FROM = os.environ.get("MAIL_FROM", "no-reply@ccyb-monitor.local")
MAIL_TO = os.environ.get("MAIL_TO")  # uno o varios, separados por comas
SMTP_HOST = os.environ.get("SMTP_HOST")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD")

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}


# ---------------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------------

def strip_accents(text: str) -> str:
    text = unicodedata.normalize("NFKD", str(text))
    return "".join(c for c in text if not unicodedata.combining(c))


def normalize(text: str) -> str:
    return strip_accents(str(text)).lower().strip()


def normalize_country_key(text: str) -> str:
    return strip_accents(str(text)).upper().strip()


def download_xlsx() -> bytes:
    resp = requests.get(XLSX_URL, headers=REQUEST_HEADERS, timeout=30)
    resp.raise_for_status()
    return resp.content


def load_table(xlsx_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=0, header=None)

    header_row = None
    for i in range(min(15, len(raw))):
        row_values = [normalize(v) for v in raw.iloc[i].tolist()]
        if any(h in row_values or any(h in v for v in row_values)
               for h in COLUMN_HINTS["country"]):
            header_row = i
            break
    if header_row is None:
        header_row = 0

    df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=0, header=header_row)
    df = df.dropna(how="all")
    df.columns = [str(c) for c in df.columns]
    return df


def detect_columns(df: pd.DataFrame) -> dict:
    detected = {}
    normalized_cols = {col: normalize(col) for col in df.columns}
    used_cols = set()

    for category, hints in COLUMN_HINTS.items():
        match = None
        for hint in sorted(hints, key=len, reverse=True):
            for col, norm_col in normalized_cols.items():
                if col in used_cols:
                    continue
                if hint in norm_col:
                    match = col
                    break
            if match:
                break
        if match:
            used_cols.add(match)
        detected[category] = match

    return detected


def parse_date(value):
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y", "%d %B %Y", "%d %b %Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    try:
        return pd.to_datetime(text, dayfirst=True).date()
    except Exception:
        return None


def format_rate(value) -> str:
    """La columna 'CCyB rate' del Excel de la ESRB ya viene en puntos
    porcentuales (1 = 1%, 0.5 = 0.5%). No hay que multiplicar por 100."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "?"
    try:
        num = round(float(value), 2)
        return f"{num:g}%"
    except (TypeError, ValueError):
        return str(value)


def rate_sql_literal(value) -> str:
    """Igual que format_rate pero sin el símbolo %, para usar como número
    en el SQL (2 -> '2', 0.5 -> '0.5')."""
    num = round(float(value), 2)
    return f"{num:g}"


def format_date_es(d: date) -> str:
    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
             "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    return f"{d.day} de {meses[d.month - 1]} de {d.year}"


# ---------------------------------------------------------------------------
# CÓDIGOS ISO Y NOMBRES DE PAÍS
# ---------------------------------------------------------------------------

def load_iso_codes() -> dict:
    """Carga data/country_iso_codes.csv -> {NOMBRE_NORMALIZADO: codigo_iso}."""
    mapping = {}
    if not ISO_CODES_PATH.exists():
        return mapping
    with open(ISO_CODES_PATH, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            code = row.get("cod_iso_3166", "").strip()
            desc = row.get("descripcion", "").strip()
            if code and desc:
                mapping[normalize_country_key(desc)] = code
    return mapping


def country_to_iso(country_name: str, iso_map: dict) -> str | None:
    key = normalize_country_key(country_name)
    key = normalize_country_key(COUNTRY_NAME_ALIASES.get(key, key))
    return iso_map.get(key)


def country_to_es(country_name: str) -> str:
    key = normalize_country_key(country_name)
    return COUNTRY_NAME_ES.get(key, country_name)


# ---------------------------------------------------------------------------
# LÓGICA PRINCIPAL: LECTURA DEL EXCEL
# ---------------------------------------------------------------------------

def compute_table_hash(df: pd.DataFrame) -> str:
    canonical = df.fillna("").astype(str)
    canonical = canonical.sort_values(by=list(canonical.columns)).reset_index(drop=True)
    payload = canonical.to_csv(index=False)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def build_country_timelines(df: pd.DataFrame, cols: dict) -> dict:
    timelines = {}
    country_col = cols["country"]
    rate_col = cols["rate"]
    date_col = cols["application_date"]

    for _, row in df.iterrows():
        country = str(row.get(country_col, "")).strip()
        if not country or country.lower() == "nan":
            continue

        app_date = parse_date(row.get(date_col))
        rate_raw = row.get(rate_col)
        if app_date is None or pd.isna(rate_raw):
            continue

        timelines.setdefault(country, []).append((app_date, rate_raw))

    for country in timelines:
        timelines[country].sort(key=lambda x: x[0])

    return timelines


def find_upcoming_changes(timelines: dict, today: date) -> list:
    upcoming = []
    window_end = today + timedelta(days=LOOKAHEAD_DAYS)

    for country, entries in timelines.items():
        current_rate_raw = None
        for app_date, rate_raw in entries:
            if app_date <= today:
                current_rate_raw = rate_raw
            else:
                break

        for app_date, rate_raw in entries:
            if today < app_date <= window_end:
                upcoming.append({
                    "country": country,
                    "current_rate_raw": current_rate_raw,
                    "future_rate_raw": rate_raw,
                    "current_rate": format_rate(current_rate_raw) if current_rate_raw is not None else "0%",
                    "future_rate": format_rate(rate_raw),
                    "effective_date": app_date,
                })

    return upcoming


# ---------------------------------------------------------------------------
# LÓGICA DE FECHA DE CIERRE Y GENERACIÓN DE SQL
# ---------------------------------------------------------------------------

def next_quarter_start_after(d: date) -> date:
    """Primera fecha de trimestre (01/01, 01/04, 01/07, 01/10)
    ESTRICTAMENTE posterior a d."""
    candidates = []
    for y in (d.year, d.year + 1):
        for m in (1, 4, 7, 10):
            candidates.append(date(y, m, 1))
    candidates = sorted(c for c in candidates if c > d)
    return candidates[0]


def closing_label(closing_date: date):
    """Devuelve (nombre_mes_cierre, año_cierre) a partir de la fecha de
    cierre (01/01, 01/04, 01/07, 01/10). Ej: 2027-01-01 -> ('diciembre', 2026)."""
    month = closing_date.month
    if month == 1:
        return "diciembre", closing_date.year - 1
    elif month == 4:
        return "marzo", closing_date.year
    elif month == 7:
        return "junio", closing_date.year
    elif month == 10:
        return "septiembre", closing_date.year
    raise ValueError(f"Fecha de cierre inesperada: {closing_date}")


def build_sql_script(upcoming: list, iso_map: dict) -> tuple[str, list]:
    """Genera el script SQL con un UPDATE por cada cambio a un mes vista.
    Devuelve (texto_sql, lista_de_avisos) — avisos para países sin código
    ISO reconocido, que se omiten del SQL y hay que revisar a mano."""
    lines = [
        "-- Script generado automáticamente por ccyb_monitor.py",
        f"-- Generado el {date.today().isoformat()}",
        "-- Actualiza dbo.colchon_capital_anticiclico con los cambios de CCyB",
        "-- detectados a un mes vista. Revisar antes de ejecutar en producción.",
        "",
    ]
    warnings = []

    for item in sorted(upcoming, key=lambda x: x["effective_date"]):
        iso = country_to_iso(item["country"], iso_map)
        if not iso:
            warnings.append(item["country"])
            lines.append(
                f"-- ATENCIÓN: no se ha encontrado código ISO para "
                f"'{item['country']}'. Revisar y añadir manualmente."
            )
            lines.append("")
            continue

        closing_date = next_quarter_start_after(item["effective_date"])
        mes, anyo = closing_label(closing_date)
        yy = str(anyo)[-2:]
        closing_str = closing_date.strftime("%Y/%m/%d")
        fecha_informacion = closing_date.strftime("%Y%m%d")
        rate_str = rate_sql_literal(item["future_rate_raw"])
        country_es = country_to_es(item["country"])

        lines.append(f"--{country_es} (actualizado a cierre de {mes}{yy})")
        lines.append(
            f"update dbo.colchon_capital_anticiclico set porc_buff_antici_autoridad = {rate_str}, "
        )
        lines.append(
            f"porc_buff_antici_pais_enti = {rate_str}, porc_buff_antici_espec_enti = {rate_str} "
        )
        lines.append("from colchon_capital_anticiclico")
        lines.append(
            f"where fecha_informacion = '{fecha_informacion}' and cod_pais_fecha = '{iso} {closing_str}'"
        )
        lines.append("")

    return "\n".join(lines), warnings


# ---------------------------------------------------------------------------
# EMAIL
# ---------------------------------------------------------------------------

def compose_email_body(upcoming: list, esrb_changed_since_last_check: bool,
                        today: date, sql_warnings: list) -> str:
    lines = []
    lines.append(f"Informe mensual del colchón de capital anticíclico (CCyB) — {format_date_es(today)}")
    lines.append("Fuente: ESRB — https://www.esrb.europa.eu/national_policy/ccb/html/index.en.html")
    lines.append("")

    lines.append(f"CAMBIOS A UN MES VISTA (próximos {LOOKAHEAD_DAYS} días):")
    if upcoming:
        for item in sorted(upcoming, key=lambda x: x["effective_date"]):
            lines.append(
                f"  • {item['country']}: el CCyB pasará del "
                f"{item['current_rate']} al {item['future_rate']}, "
                f"efectivo desde el {format_date_es(item['effective_date'])}."
            )
        lines.append("")
        lines.append(
            "Se adjunta script SQL con los UPDATE necesarios sobre "
            "dbo.colchon_capital_anticiclico para estos cambios."
        )
    else:
        lines.append("  No se han detectado cambios de CCyB en ningún país a un mes vista.")

    if sql_warnings:
        lines.append("")
        lines.append(
            "AVISO: no se ha encontrado código ISO para: " + ", ".join(sql_warnings) +
            ". Revisa esos países manualmente (quedan comentados en el SQL adjunto)."
        )

    lines.append("")
    if esrb_changed_since_last_check:
        lines.append(
            "Nota: la ESRB ha actualizado la tabla de datos desde el último "
            "chequeo (puede incluir cambios fuera de la ventana de un mes; "
            "revisa la página para más detalle)."
        )
    else:
        lines.append("La ESRB no ha publicado ninguna actualización desde el último chequeo.")

    lines.append("")
    lines.append("— Este es un aviso automático, no respondas a este correo. —")
    return "\n".join(lines)


def send_email(subject: str, body: str, attachment_name: str = None, attachment_text: str = None):
    if not MAIL_TO:
        raise RuntimeError("Falta la variable de entorno MAIL_TO (uno o varios emails, separados por comas).")
    if not (SMTP_HOST and SMTP_USER and SMTP_PASSWORD):
        raise RuntimeError(
            "Faltan credenciales SMTP (SMTP_HOST / SMTP_USER / SMTP_PASSWORD). "
            "Revisa el README para configurarlas."
        )

    recipients = [addr.strip() for addr in MAIL_TO.split(",") if addr.strip()]

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = MAIL_FROM
    msg["To"] = ", ".join(recipients)
    msg.attach(MIMEText(body, "plain", "utf-8"))

    if attachment_name and attachment_text is not None:
        part = MIMEText(attachment_text, "plain", "utf-8")
        part.add_header("Content-Disposition", "attachment", filename=attachment_name)
        msg.attach(part)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(MAIL_FROM, recipients, msg.as_string())


# ---------------------------------------------------------------------------
# COMANDOS
# ---------------------------------------------------------------------------

def run_diagnose():
    print(f"Descargando {XLSX_URL} ...")
    xlsx_bytes = download_xlsx()
    df = load_table(xlsx_bytes)

    print("\nColumnas detectadas en el Excel:")
    for c in df.columns:
        print(f"  - {c!r}")

    cols = detect_columns(df)
    print("\nMapeo automático (categoría -> columna encontrada):")
    for category, col in cols.items():
        estado = col if col else "NO ENCONTRADA (ajusta COLUMN_HINTS)"
        print(f"  - {category}: {estado}")

    iso_map = load_iso_codes()
    print(f"\nCódigos ISO cargados: {len(iso_map)} países ({ISO_CODES_PATH.name})")

    if cols.get("country") and cols.get("rate") and cols.get("application_date"):
        timelines = build_country_timelines(df, cols)
        today = date.today()
        upcoming = find_upcoming_changes(timelines, today)
        print(f"\nPaíses con histórico de decisiones detectados: {len(timelines)}")
        print(f"Cambios a un mes vista encontrados HOY ({today}): {len(upcoming)}")
        for item in upcoming:
            iso = country_to_iso(item["country"], iso_map)
            print(f"  - {item['country']} (ISO: {iso or 'NO ENCONTRADO'}): "
                  f"{item['current_rate']} -> {item['future_rate']} el {item['effective_date']}")

        if upcoming:
            sql_text, warnings = build_sql_script(upcoming, iso_map)
            print("\n--- SQL que se generaría ---")
            print(sql_text)
            if warnings:
                print(f"AVISO: sin código ISO para: {warnings}")

    print("\nPrimeras filas:")
    print(df.head(10).to_string())


def run_check(dry_run: bool = False):
    today = date.today()

    xlsx_bytes = download_xlsx()
    df = load_table(xlsx_bytes)
    cols = detect_columns(df)

    required = ["country", "rate", "application_date"]
    missing = [c for c in required if not cols.get(c)]
    if missing:
        raise RuntimeError(
            f"No se han podido identificar estas columnas: {missing}. "
            "Ejecuta 'python ccyb_monitor.py --diagnose' y ajusta COLUMN_HINTS."
        )

    new_hash = compute_table_hash(df)
    old_data = {}
    if SNAPSHOT_PATH.exists():
        old_data = json.loads(SNAPSHOT_PATH.read_text(encoding="utf-8"))
    old_hash = old_data.get("table_hash")

    esrb_changed_since_last_check = (old_hash is not None) and (new_hash != old_hash)
    first_run = old_hash is None

    timelines = build_country_timelines(df, cols)
    upcoming = find_upcoming_changes(timelines, today)

    iso_map = load_iso_codes()
    sql_text, sql_warnings = ("", [])
    if upcoming:
        sql_text, sql_warnings = build_sql_script(upcoming, iso_map)

    subject = "Informe mensual CCyB — " + (
        "cambios a un mes vista" if upcoming else "sin cambios a un mes vista"
    )
    body = compose_email_body(upcoming, esrb_changed_since_last_check and not first_run, today, sql_warnings)

    print(body)
    if upcoming:
        print("\n--- SQL adjunto ---")
        print(sql_text)

    if not dry_run:
        attachment_name = f"actualizacion_ccyb_{today.isoformat()}.sql" if upcoming else None
        send_email(subject, body, attachment_name, sql_text if upcoming else None)
        SNAPSHOT_PATH.parent.mkdir(parents=True, exist_ok=True)
        SNAPSHOT_PATH.write_text(
            json.dumps({"table_hash": new_hash, "last_checked": today.isoformat()},
                       ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    else:
        print("\n[--dry-run] No se ha enviado ningún email ni actualizado el snapshot.")


def main():
    parser = argparse.ArgumentParser(description="Monitor mensual del CCyB (ESRB).")
    parser.add_argument("--diagnose", action="store_true",
                         help="Descarga el Excel e imprime su estructura, sin enviar email.")
    parser.add_argument("--dry-run", action="store_true",
                         help="Hace todo el proceso pero no envía el email ni guarda el snapshot.")
    args = parser.parse_args()

    if args.diagnose:
        run_diagnose()
    else:
        run_check(dry_run=args.dry_run)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
