#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ccyb_monitor.py
================
Monitoriza el archivo Excel oficial que la ESRB publica con los colchones
de capital anticíclico (CCyB) de cada país:

    https://www.esrb.europa.eu/national_policy/ccb/shared/data/esrb.ccybd_CCyB_data.xlsx

El archivo es una BASE DE DATOS HISTÓRICA: cada fila es una decisión de un
país (columnas reales confirmadas: 'Country', 'Decision on',
'Date of Announcement', 'CCyB rate', 'Type of setting', 'Application since',
'Credit-to-GDP', 'Reference date', 'Credit Gap', 'Buffer Guide',
'Additional Gap', 'Additional benchmark', 'Justification',
'Justification exceptional circumstances', 'Period without decrease', 'Link').
Un mismo país puede tener muchas filas: unas con fecha de aplicación ya
pasada (vigentes) y otras con fecha futura (pendientes).

Cada vez que se ejecuta este script:
  1. Descarga el Excel más reciente.
  2. Calcula un hash de toda la tabla y lo compara con el de la última vez
     (data/last_snapshot.json) para saber si la ESRB ha publicado algo
     nuevo desde el último chequeo.
  3. Para cada país, calcula qué tasa está en vigor HOY (la decisión con
     'Application since' más reciente que ya haya pasado) y busca si hay
     alguna decisión pendiente cuya 'Application since' caiga dentro de la
     ventana "a un mes vista" (por defecto, los próximos 35 días).
  4. Envía SIEMPRE un email mensual con el resultado:
       - Si hay cambios a un mes vista: los detalla país por país,
         ej. "Polonia subirá el CCyB del 1% al 2%, efectivo desde el 30
         de septiembre de 2026."
       - Si no hay ninguno: "No se han detectado cambios de CCyB en
         ningún país a un mes vista."

-------------------------------------------------------------------------
CALIBRACIÓN
-------------------------------------------------------------------------
Ejecuta en cualquier momento (localmente o vía GitHub Actions con
mode=diagnose):

    python ccyb_monitor.py --diagnose

para ver las columnas reales del Excel y confirmar que el script las ha
identificado bien. Si la ESRB cambia el nombre de alguna columna en el
futuro, este modo te lo dirá y solo tendrás que añadir la palabra nueva a
COLUMN_HINTS, un poco más abajo.
"""

import argparse
import hashlib
import io
import json
import os
import smtplib
import sys
import unicodedata
from datetime import date, datetime, timedelta
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

# "A un mes vista": margen en días. 35 cubre con holgura cualquier mes del año.
LOOKAHEAD_DAYS = 35

# Palabras clave (en minúsculas y sin acentos) para reconocer cada columna
# automáticamente. Confirmadas contra el Excel real de la ESRB (sept. 2026):
# 'Country', 'Decision on', 'Date of Announcement', 'CCyB rate',
# 'Type of setting', 'Application since', ...
COLUMN_HINTS = {
    "country": ["country"],
    "rate": ["ccyb rate", "rate"],
    "application_date": ["application since", "applicable", "in effect since"],
    "decision_date": ["decision on"],
    "announcement_date": ["date of announcement"],
}

# Email
MAIL_FROM = os.environ.get("MAIL_FROM", "no-reply@ccyb-monitor.local")
MAIL_TO = os.environ.get("MAIL_TO")  # tu correo corporativo, obligatorio
SMTP_HOST = os.environ.get("SMTP_HOST")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD")

REQUEST_HEADERS = {
    # Algunos servidores rechazan peticiones sin cabecera de navegador.
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


def download_xlsx() -> bytes:
    resp = requests.get(XLSX_URL, headers=REQUEST_HEADERS, timeout=30)
    resp.raise_for_status()
    return resp.content


def load_table(xlsx_bytes: bytes) -> pd.DataFrame:
    """Carga la primera hoja con datos del Excel, detectando la fila de
    cabecera aunque no sea la primera línea del fichero."""
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
    """Empareja cada columna del Excel con una categoría de COLUMN_HINTS.
    Usa coincidencia por palabra clave más larga primero, para evitar que
    p.ej. 'rate' (hint genérico) se cuele antes que 'ccyb rate'."""
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
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "?"
    try:
        num = float(value)
        # Si viene como fracción (0.01) lo pasamos a porcentaje.
        if abs(num) <= 1:
            num *= 100
        num = round(num, 2)
        return f"{num:g}%"
    except (TypeError, ValueError):
        return str(value)


def format_date_es(d: date) -> str:
    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
             "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    return f"{d.day} de {meses[d.month - 1]} de {d.year}"


# ---------------------------------------------------------------------------
# LÓGICA PRINCIPAL
# ---------------------------------------------------------------------------

def compute_table_hash(df: pd.DataFrame) -> str:
    """Hash de toda la tabla (todas las columnas, todas las filas),
    independiente de si sabemos interpretar cada columna. Sirve para saber
    si la ESRB ha tocado CUALQUIER cosa desde el último chequeo."""
    canonical = df.fillna("").astype(str)
    canonical = canonical.sort_values(by=list(canonical.columns)).reset_index(drop=True)
    payload = canonical.to_csv(index=False)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def build_country_timelines(df: pd.DataFrame, cols: dict) -> dict:
    """Devuelve {país: [(fecha_aplicacion, tasa_raw), ...]} ordenado por fecha."""
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
    """Para cada país, mira si hay una decisión cuya fecha de aplicación
    caiga dentro de la ventana 'a un mes vista', y calcula cuál es la
    tasa vigente hoy para poder mostrar el 'antes -> después'."""
    upcoming = []
    window_end = today + timedelta(days=LOOKAHEAD_DAYS)

    for country, entries in timelines.items():
        # Tasa vigente hoy: la última decisión cuya fecha ya ha pasado.
        current_rate_raw = None
        for app_date, rate_raw in entries:
            if app_date <= today:
                current_rate_raw = rate_raw
            else:
                break  # entries está ordenado, ya no hace falta seguir para "hoy"

        for app_date, rate_raw in entries:
            if today < app_date <= window_end:
                upcoming.append({
                    "country": country,
                    "current_rate": format_rate(current_rate_raw) if current_rate_raw is not None else "0%",
                    "future_rate": format_rate(rate_raw),
                    "effective_date": app_date,
                })

    return upcoming


def compose_email_body(upcoming: list, esrb_changed_since_last_check: bool,
                        today: date) -> str:
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
    else:
        lines.append("  No se han detectado cambios de CCyB en ningún país a un mes vista.")

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


def send_email(subject: str, body: str):
    if not MAIL_TO:
        raise RuntimeError("Falta la variable de entorno MAIL_TO (tu email corporativo).")
    if not (SMTP_HOST and SMTP_USER and SMTP_PASSWORD):
        raise RuntimeError(
            "Faltan credenciales SMTP (SMTP_HOST / SMTP_USER / SMTP_PASSWORD). "
            "Revisa el README para configurarlas."
        )

    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"] = MAIL_FROM
    msg["To"] = MAIL_TO

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(MAIL_FROM, [MAIL_TO], msg.as_string())


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

    if cols.get("country") and cols.get("rate") and cols.get("application_date"):
        timelines = build_country_timelines(df, cols)
        today = date.today()
        upcoming = find_upcoming_changes(timelines, today)
        print(f"\nPaíses con histórico de decisiones detectados: {len(timelines)}")
        print(f"Cambios a un mes vista encontrados HOY ({today}): {len(upcoming)}")
        for item in upcoming:
            print(f"  - {item}")

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

    # 1) ¿Ha cambiado algo respecto a la última vez que se comprobó?
    new_hash = compute_table_hash(df)
    old_data = {}
    if SNAPSHOT_PATH.exists():
        old_data = json.loads(SNAPSHOT_PATH.read_text(encoding="utf-8"))
    old_hash = old_data.get("table_hash")

    esrb_changed_since_last_check = (old_hash is not None) and (new_hash != old_hash)
    first_run = old_hash is None

    # 2) Cambios a un mes vista
    timelines = build_country_timelines(df, cols)
    upcoming = find_upcoming_changes(timelines, today)

    # 3) Email (siempre se manda, una vez al mes)
    subject = "Informe mensual CCyB — " + (
        "cambios a un mes vista" if upcoming else "sin cambios a un mes vista"
    )
    body = compose_email_body(upcoming, esrb_changed_since_last_check and not first_run, today)

    print(body)

    if not dry_run:
        send_email(subject, body)
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
