#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ccyb_monitor.py
================
Monitoriza el archivo Excel oficial que la ESRB publica con los colchones
de capital anticíclico (CCyB) de cada país:

    https://www.esrb.europa.eu/national_policy/ccb/shared/data/esrb.ccybd_CCyB_data.xlsx

Cada vez que se ejecuta:
  1. Descarga el Excel más reciente.
  2. Compara los datos con la última "foto" guardada (data/last_snapshot.json)
     para detectar si la ESRB ha publicado algo nuevo desde la última vez.
  3. Busca, para cada país, si hay un cambio de colchón cuya fecha de
     aplicación caiga dentro de la ventana "a un mes vista" (por defecto,
     los próximos 35 días desde hoy).
  4. Envía SIEMPRE un email mensual con el resultado:
       - Si hay cambios a un mes vista: los detalla país por país.
       - Si no hay ninguno: "No se han detectado cambios de CCyB en
         ningún país a un mes vista."
     Además añade, si procede, un aviso de que la ESRB ha actualizado la
     tabla desde el último chequeo (aunque el cambio no caiga en la
     ventana de un mes).

-------------------------------------------------------------------------
CALIBRACIÓN (IMPORTANTE, LEER ANTES DEL PRIMER USO)
-------------------------------------------------------------------------
No he podido descargar el Excel de la ESRB desde este entorno para
verificar el nombre EXACTO de sus columnas (el dominio esrb.europa.eu no
es accesible desde aquí). Por eso este script:

  a) SIEMPRE es capaz de decirte si la ESRB ha tocado la tabla desde la
     última vez (compara la fila completa de cada país, columna por
     columna, así que funciona pase lo que pase con los nombres).

  b) Para el aviso "a un mes vista" necesita identificar qué columnas son
     "país", "tasa actual", "fecha de aplicación actual", "tasa futura" y
     "fecha de aplicación futura". Lo intenta de forma automática por
     palabras clave (ver COLUMN_HINTS más abajo), pero como no he podido
     ver el archivo real, TE RECOMIENDO hacer esto una sola vez:

       1. Ejecuta:  python ccyb_monitor.py --diagnose
          Esto descarga el Excel, imprime el nombre real de cada columna
          y las primeras filas, y te dice qué columnas ha detectado para
          cada categoría, SIN enviar ningún email.
       2. Si el detector automático acierta (lo normal), no tienes que
          tocar nada más.
       3. Si NO acierta, añade la palabra exacta que veas en la columna
          conflictiva a la lista correspondiente dentro de COLUMN_HINTS.

Esto solo hay que hacerlo una vez; si la ESRB cambia el formato del
archivo en el futuro, el modo --diagnose te lo dirá (te avisará si no
logra encontrar alguna columna).
"""

import argparse
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
# automáticamente. Si el modo --diagnose falla en detectar alguna, añade
# aquí la palabra exacta que veas impresa por --diagnose.
COLUMN_HINTS = {
    "country": ["country", "member state", "pais", "estado miembro"],
    "current_rate": ["current ccyb", "current rate", "applicable rate",
                      "rate in effect", "ccyb rate (current)"],
    "current_date": ["date of application", "applicable as of",
                      "effective date", "date of applicability"],
    "future_rate": ["future ccyb", "future rate", "pending rate",
                     "announced rate", "ccyb rate (future)"],
    "future_date": ["future date", "pending date",
                     "date of future application", "date of future"],
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
        # Si no la detecta, asumimos que la primera fila es la cabecera.
        header_row = 0

    df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=0, header=header_row)
    df = df.dropna(how="all")
    df.columns = [str(c) for c in df.columns]
    return df


def detect_columns(df: pd.DataFrame) -> dict:
    """Empareja cada columna del Excel con una categoría de COLUMN_HINTS."""
    detected = {}
    normalized_cols = {col: normalize(col) for col in df.columns}

    for category, hints in COLUMN_HINTS.items():
        match = None
        for col, norm_col in normalized_cols.items():
            if any(hint in norm_col for hint in hints):
                match = col
                break
        detected[category] = match

    return detected


def parse_date(value):
    if pd.isna(value):
        return None
    if isinstance(value, (datetime, date)):
        return value if isinstance(value, date) and not isinstance(value, datetime) else value.date()
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
    if pd.isna(value):
        return "?"
    try:
        num = float(value)
        # Si viene como fracción (0.01) lo pasamos a porcentaje.
        if num <= 1:
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

def build_row_snapshot(df: pd.DataFrame, country_col: str) -> dict:
    """Serializa cada fila (todas las columnas) por país, para poder
    detectar CUALQUIER cambio publicado por la ESRB, sepamos o no
    interpretar cada columna."""
    snapshot = {}
    for _, row in df.iterrows():
        country = str(row.get(country_col, "")).strip()
        if not country or country.lower() == "nan":
            continue
        snapshot[country] = {str(k): ("" if pd.isna(v) else str(v))
                              for k, v in row.items()}
    return snapshot


def find_upcoming_changes(df: pd.DataFrame, cols: dict, today: date) -> list:
    """Busca, país por país, cambios de tasa cuya fecha de aplicación
    caiga dentro de la ventana 'a un mes vista'."""
    upcoming = []
    required = ["country", "future_rate", "future_date"]
    if not all(cols.get(c) for c in required):
        return upcoming  # no se pudieron identificar las columnas necesarias

    window_end = today + timedelta(days=LOOKAHEAD_DAYS)

    for _, row in df.iterrows():
        country = str(row.get(cols["country"], "")).strip()
        if not country or country.lower() == "nan":
            continue

        future_rate_raw = row.get(cols["future_rate"])
        future_date_raw = row.get(cols["future_date"])
        if pd.isna(future_rate_raw) or pd.isna(future_date_raw):
            continue

        effective_date = parse_date(future_date_raw)
        if not effective_date:
            continue

        if today <= effective_date <= window_end:
            current_rate_raw = row.get(cols.get("current_rate")) if cols.get("current_rate") else None
            upcoming.append({
                "country": country,
                "current_rate": format_rate(current_rate_raw) if current_rate_raw is not None else "?",
                "future_rate": format_rate(future_rate_raw),
                "effective_date": effective_date,
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
                f"  • {item['country']} subirá/cambiará el CCyB del "
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

    print("\nPrimeras filas:")
    print(df.head(10).to_string())

    print(
        "\nSi alguna categoría aparece como 'NO ENCONTRADA', abre el Excel, "
        "localiza el nombre real de esa columna y añade una palabra clave "
        "suya (en minúsculas y sin tildes) a COLUMN_HINTS en este script."
    )


def run_check(dry_run: bool = False):
    today = date.today()

    xlsx_bytes = download_xlsx()
    df = load_table(xlsx_bytes)
    cols = detect_columns(df)

    if not cols.get("country"):
        raise RuntimeError(
            "No se ha podido identificar la columna de país en el Excel. "
            "Ejecuta 'python ccyb_monitor.py --diagnose' y ajusta COLUMN_HINTS."
        )

    # 1) ¿Ha cambiado algo respecto a la última vez que se comprobó?
    new_snapshot = build_row_snapshot(df, cols["country"])
    old_snapshot = {}
    if SNAPSHOT_PATH.exists():
        old_snapshot = json.loads(SNAPSHOT_PATH.read_text(encoding="utf-8"))

    esrb_changed_since_last_check = new_snapshot != old_snapshot

    # 2) Cambios a un mes vista
    upcoming = find_upcoming_changes(df, cols, today)

    # 3) Email (siempre se manda, una vez al mes)
    subject = "Informe mensual CCyB — " + (
        "cambios a un mes vista" if upcoming else "sin cambios a un mes vista"
    )
    body = compose_email_body(upcoming, esrb_changed_since_last_check, today)

    print(body)

    if not dry_run:
        send_email(subject, body)
        SNAPSHOT_PATH.parent.mkdir(parents=True, exist_ok=True)
        SNAPSHOT_PATH.write_text(
            json.dumps(new_snapshot, ensure_ascii=False, indent=2), encoding="utf-8"
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
