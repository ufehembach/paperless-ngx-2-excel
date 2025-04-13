#!/usr/bin/env python3

import os
import sys
import pwd
#from networkx import from_prufer_sequence
import requests
import pandas as pd
import inspect
import argparse
import json
import locale
import re
import zipfile
import configparser
import glob
import pprint
import asyncio
import os
import aiohttp
import re
import shutil
from time import sleep
import asyncio
import aiohttp
from datetime import datetime
from dateutil import parser
from collections import OrderedDict

from configparser import ConfigParser
from tqdm import tqdm
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils.dataframe import dataframe_to_rows
from collections import OrderedDict

from datetime import datetime, timedelta
from dateutil import parser

import asyncio
import aiohttp
from pypaperless import Paperless
import glob
import os

import os
import glob

import shutil

import subprocess

import os
import subprocess
import urllib.request
import json
import re

# Neue zentrale message-Funktion + globale LOG_PATH Variable

from datetime import datetime
import os
import atexit

LOG_PATH = None  # Wird beim Log-Setup gesetzt
_final_log_path = None
_last_message_was_inline = False  # Verfolgt, ob vorherige Ausgabe inline war


def message(text: str, target: str = "both", level: str = "info", inline: bool = False):
    global _last_message_was_inline

    prefix = {
        "info": "➕",
        "warn": "⚠️",
        "error": "❌"
    }.get(level, "➕")

    full_message = f"{prefix} {text}"

    # Wenn letzte Ausgabe inline war, füge einen echten Zeilenumbruch ein
    # if not inline and _last_message_was_inline:
    #     print()
    #     _last_message_was_inline = False

    if inline:
        print(full_message, end='\r', flush=True)
        _last_message_was_inline = True
        return

    if target in ("console", "both"):
        print(full_message)

    if target in ("log", "both") and LOG_PATH:
        try:
            # Kontextinfo für Log-Nachricht
            frame = inspect.currentframe().f_back
            filename = os.path.basename(frame.f_code.co_filename)
            line_number = frame.f_lineno
            function_name = frame.f_code.co_name
            context = f"{filename}:{line_number} [{function_name}]"
            with open(LOG_PATH, "a", encoding="utf-8") as f:
                f.write(f"{datetime.now()} - {context} {full_message}\n")
        except Exception as e:
            print(f"⚠️ Fehler beim Schreiben ins Log: {e}")

def set_log_path(path: str):
    global LOG_PATH
    LOG_PATH = path
    message(f"Log-Dateipfad gesetzt: {path}", target="console", level="info")

def mask_secret(secret: str, show: int = 4) -> str:
    """Maskiert einen String, zeigt nur die ersten und letzten Zeichen."""
    if not secret or len(secret) <= show * 2:
        return "*" * len(secret)
    return f"{secret[:show]}{'*' * (len(secret) - 2 * show)}{secret[-show:]}"

def xxcleanup_old_files(dir_path, filename_prefix, max_count_str, pattern="log"):
    max_count = int(max_count_str)
    glob_pattern = os.path.join(dir_path, f"{filename_prefix}*.{pattern}")
    files = sorted(glob.glob(glob_pattern), key=os.path.getmtime)
    message(f"[Cleanup] Suche: {glob_pattern} – gefunden: {len(files)} Dateien", target="log")
    if len(files) <= max_count:
        message(f"[Cleanup] Nichts zu tun: {len(files)} ≤ {max_count}", target="log")
        return
    while len(files) > max_count:
        old_file = files.pop(0)
        try:
            os.remove(old_file)
            message(f"[Cleanup] Datei gelöscht: {old_file}", target="log")
        except OSError as e:
            message(f"[Cleanup] Fehler beim Löschen von {old_file}: {e}", level="warn", target="log")

def get_log_filename(script_name, log_dir, suffix="progress"):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    ext = "log" if suffix == "log" else f"{suffix}.log"
    return os.path.join(log_dir, f"##{script_name}__{timestamp}.{ext}")

def initialize_log(log_dir, script_name, max_files):
    final_log_path = get_log_filename(script_name, log_dir, "log")
    progress_log_path = get_log_filename(script_name, log_dir, "progress")
    if os.path.exists(final_log_path):
        with open(progress_log_path, "w") as new_log, open(final_log_path, "r") as old_log:
            shutil.copyfileobj(old_log, new_log)
        os.remove(final_log_path)
    else:
        open(progress_log_path, "w").close()
    cleanup_old_files(log_dir, "##" + script_name, max_files)
    return progress_log_path, final_log_path

def finalize_log():
    global LOG_PATH, _final_log_path
    if LOG_PATH and _final_log_path and os.path.exists(LOG_PATH):
        os.rename(LOG_PATH, _final_log_path)
        message("Log-Datei finalisiert.", target="console")

def prepare_logging(log_dir, script_name, max_files):
    global _final_log_path
    progress_log_path, final_log_path = initialize_log(log_dir, script_name, max_files)
    _final_log_path = final_log_path
    set_log_path(progress_log_path)
    atexit.register(finalize_log)

# Beispielverwendung nach Initialisierung:
# set_log_path("pfad/zur/logdatei.log")
# message("Export gestartet", target="both")
# message("Dokument konnte nicht geladen werden", level="warn")
# message("Verarbeite Dokument 17", inline=True)


def get_git_version(default="v0.0.0"):
    try:
        version = subprocess.check_output(
            ["git", "describe", "--tags", "--always"],
            stderr=subprocess.DEVNULL
        )
        return version.decode("utf-8").strip()
    except Exception:
        return default

def get_github_repo_info():
    try:
        url = subprocess.check_output(
            ["git", "config", "--get", "remote.origin.url"],
            stderr=subprocess.DEVNULL
        ).decode("utf-8").strip()

        # z. B. https://github.com/ufe-dev/paperless-ngx-2-excel.git
        # oder git@github.com:ufe-dev/paperless-ngx-2-excel.git

        match = re.search(r"github.com[:/](.+?)/(.+?)(\.git)?$", url)
        if match:
            user, repo = match.group(1), match.group(2)
            return user, repo
    except Exception:
        pass

    return None, None

def get_github_license_identifier(user, repo):
    url = f"https://api.github.com/repos/{user}/{repo}/license"
    try:
        with urllib.request.urlopen(url) as response:
            data = json.load(response)
            return data["license"]["spdx_id"]
    except Exception as e:
        print(f"⚠️ Lizenz konnte nicht geladen werden: {e}")
        return None

def print_program_header():
    script_name = os.path.basename(__file__)
    version = get_git_version()
    user, repo = get_github_repo_info()
    license_id = get_github_license_identifier(user, repo) if user and repo else "Unbekannt"
    github_url = f"https://github.com/{user}/{repo}" if user and repo else "(kein GitHub-Repo erkannt)"

    print(f"{script_name} {version} – © 2025 {license_id} – {github_url}", end="")

def append_metadata_sheet(
    xlsx_path,
    script_version,
    config_data,
    python_packages_limit=10
):
    import os
    import socket
    import getpass
    import platform
    import pandas as pd
    from datetime import datetime
    from openpyxl import load_workbook
    import importlib.metadata

    def add_block(rows, title, entries):
        rows.append([title, ""])
        for key, value in entries:
            rows.append([key, value])
        rows.append(["", ""])
        return rows

    # === Basisinfos ===
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    base_dir = os.path.dirname(xlsx_path)

    json_files = [f for f in os.listdir(base_dir) if f.endswith('.json')]
    pdf_files = [f for f in os.listdir(base_dir) if f.endswith('.pdf')]
    json_bytes = sum(
        os.path.getsize(os.path.join(base_dir, f))
        for f in json_files
        if os.path.isfile(os.path.join(base_dir, f))
        )

    pdf_bytes = sum(
    os.path.getsize(os.path.join(base_dir, f))
    for f in pdf_files
    if os.path.isfile(os.path.join(base_dir, f))
        )

    xlsx_size = os.path.getsize(xlsx_path)

    rows = []

    rows = add_block(rows, "📝 Exportinformationen", [
        ("Script-Version", script_version),
        ("Export-Datum", now),
        ("Hostname", socket.gethostname()),
        ("Username", getpass.getuser()),
    ])

    rows = add_block(rows, "📁 Verzeichnisse & Dateigrößen", [
        ("Verzeichnis", base_dir),
        ("Excel-Datei", os.path.basename(xlsx_path)),
        ("Größe (xlsx)", f"{xlsx_size/1024:.2f} KB"),
        ("Anzahl JSON-Dateien", len(json_files)),
        ("Bytes JSON", f"{json_bytes/1024:.2f} KB"),
        ("Anzahl PDF-Dateien", len(pdf_files)),
        ("Bytes PDF", f"{pdf_bytes/1024:.2f} KB"),
    ])

    rows = add_block(rows, "⚙️ Konfiguration (config.ini)", [
        ("Query", config_data.get("query", "(leer)")),
        ("Export Frequency", config_data.get("frequency", "(leer)")),
    ])

    installed_packages = sorted(
      [f"{dist.metadata['Name']}=={dist.version}" for dist in importlib.metadata.distributions()]
      )[:python_packages_limit]

    rows = add_block(rows, "🐍 Python Umgebung", [
        ("Python-Version", platform.python_version()),
        ("Top-Module", ', '.join(installed_packages[:python_packages_limit]) + (" ..." if len(installed_packages) > python_packages_limit else "")),
    ])

    df_meta = pd.DataFrame(rows, columns=["Betreff", "Inhalt"])

    with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_meta.to_excel(writer, sheet_name="📊 Metadaten", index=False)

def print_separator(char='#', width_ratio=2/3):
    try:
        columns = shutil.get_terminal_size().columns
    except Exception:
        columns = 80  # fallback
    print()  # Zeilenumbruch vor dem Strich
    line_width = int(columns * width_ratio)
    print(char * line_width)

def cleanup_old_files(dir_path, filename_prefix, max_count_str, pattern="log"):
    """
    Löscht alte Dateien mit bestimmtem Prefix und Endung, wenn das Limit überschritten ist.

    :param dir_path: Verzeichnis, in dem gesucht wird
    :param filename_prefix: Anfang des Dateinamens, z. B. '##steuer'
    :param max_count_str: Maximale Anzahl an Dateien als String
    :param pattern: Dateityp bzw. Dateiendung, z. B. 'log' oder 'xlsx'
    """
    max_count = int(max_count_str)
    glob_pattern = os.path.join(dir_path, f"{filename_prefix}*.{pattern}")
    files = sorted(glob.glob(glob_pattern), key=os.path.getmtime)

   # print(f"\n[Cleanup] Suche: {glob_pattern} – gefunden: {len(files)} Dateien")

    if len(files) <= max_count:
   #     print(f"[Cleanup] Nichts zu tun: {len(files)} ≤ {max_count}")
        return

    while len(files) > max_count:
        old_file = files.pop(0)
        try:
            os.remove(old_file)
            print(f"[Cleanup] Datei gelöscht: {old_file}")
        except OSError as e:
            print(f"[Cleanup] Fehler beim Löschen von {old_file}: {e}")

# ----------------------
def get_log_filename(script_name, log_dir, suffix="progress"):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    if suffix == "log":
        return os.path.join(log_dir, f"##{script_name}__{timestamp}.log")
    else:
        return os.path.join(log_dir, f"##{script_name}__{timestamp}.{suffix}.log")

# ---------------------- Configuration Loading ----------------------
def load_config(config_path):
    """Lädt eine INI-Konfigurationsdatei, gibt None zurück bei Fehlern."""
    #message("process...")
    config = configparser.ConfigParser()
    try:
        config.read(config_path)
        return config
    except configparser.DuplicateSectionError as e:
        print(f"❌ Fehlerhafte INI-Datei (Duplicate Section): {config_path} – wird übersprungen. {e}")
        return None

# ----------------------
def get_script_name():
    """Return the name of the current script without extension."""
    return os.path.splitext(os.path.basename(sys.argv[0]))[0]

# ----------------------
def load_config_from_script():
    """Load the configuration from the ini file with a priority for the .ufe.ini file."""
    script_name = get_script_name()
    ufe_ini_path = f"{script_name}.ufe.ini"
    ini_path = f"{script_name}.ini"

    # Try to load the .ufe.ini file first
    if os.path.exists(ufe_ini_path):
        message(f"Using config file: {ufe_ini_path}")
        return load_config(ufe_ini_path)
    # Fallback to the .ini file
    elif os.path.exists(ini_path):
        message(f"Using config file: {ini_path}")
        return load_config(ini_path)
    else:
        print(f"Configuration files '{ufe_ini_path}' and '{ini_path}' not found.")
        sys.exit(1)

# ----------------------
def parse_currency(value):
    """Parst einen Währungswert wie 'EUR5.00' in einen Float."""
    try:
        # Entferne Währungszeichen (alles außer Ziffern, Punkt oder Minus)
        numeric_part = ''.join(c for c in value if c.isdigit() or c == '.' or c == '-')
        return float(numeric_part)
    except Exception as e:
        # print(f"Fehler beim Parsen des Währungswerts '{value}': {e}")
        return 0.0  # Fallback auf 0 bei Fehlern

# ----------------------
def format_currency(value, currency_locale="de_DE.UTF-8"):
    if value is None:
        return ""
    try:
        clean_value = ''.join(filter(str.isdigit, value))
        if not clean_value:
            return "0,00"
        value_float = float(clean_value) / 100
    except ValueError:
        value_float = 0.0

    try:
        formatted_value = locale.currency(value_float, grouping=True)
    except Exception as e:
        formatted_value = f"{value_float:.2f}"
    return formatted_value

# ----------------------
def format_date(date_string, output_format):
    """
    Formatiert das Datum im Format '%d.%m.%Y' oder '%d.%m.%Y %H:%M' 
    in das gewünschte Format:
    - 'yyyy-mm' oder
    - 'yyyy-mm-dd'.
    
    Parameter:
    - date_string: Das Datum als String (im Format '%d.%m.%Y' oder '%d.%m.%Y %H:%M').
    - output_format: Das gewünschte Ausgabeformat ('yyyy-mm' oder 'yyyy-mm-dd').
    
    Rückgabe:
    - Das Datum im gewünschten Format als String oder None bei Fehlern.
    """
    if not date_string:
        print(f"Date string is empty or None: {date_string}")
        return None

    try:
        # Datum im ursprünglichen Format parsen
        if len(date_string.split(" ")) > 1:
            parsed_date = datetime.strptime(date_string, "%d.%m.%Y %H:%M")
        else:
            parsed_date = datetime.strptime(date_string, "%d.%m.%Y")
        
        # Rückgabe im gewünschten Format
        if output_format == "yyyy-mm":
            return parsed_date.strftime("%Y-%m")
        elif output_format == "yyyy-mm-dd":
            return parsed_date.strftime("%Y-%m-%d")
        else:
            print(f"Unsupported output format: {output_format}")
            return None
    except Exception as e:
        print(f"Failed to format date '{date_string}': {e}")
        return None

        return None

# ----------------------
def parse_date(date_input):
    """
    Gibt das Datum im Format '%d.%m.%Y' zurück, wenn Uhrzeit 00:00 ist,
    sonst im Format '%d.%m.%Y %H:%M'. Akzeptiert Strings oder datetime-Objekte.
    """
    if not date_input:
        message(f"[parse_date] Date input is empty or None: {date_input}")
        return None

    try:
        if isinstance(date_input, datetime):
            parsed_date = date_input
        else:
            parsed_date = parser.isoparse(date_input)

        if parsed_date.hour == 0 and parsed_date.minute == 0:
            return parsed_date.strftime("%d.%m.%Y")
        else:
            return parsed_date.strftime("%d.%m.%Y %H:%M")

    except Exception as e:
        message(f"[parse_date] Failed to parse date '{date_input}': {e}")
        return None

# ----------------------
async def retry_async(fn, retries=3, delay=2, backoff=2,
                      exceptions=(aiohttp.ClientError, asyncio.TimeoutError),
                      desc=None):
    current_delay = delay
    for attempt in range(1, retries + 1):
        try:
            return await fn()
        except exceptions as e:
            if attempt == retries:
                raise
            label = f' bei "{desc}"' if desc else ''
            term_width = shutil.get_terminal_size((80, 20)).columns
            print(f"\r[retry_async] Fehler{label}: {e} – Versuch {attempt}/{retries}, nächster in {current_delay}s...".ljust(term_width), end='', flush=True)
            await asyncio.sleep(current_delay)
            current_delay *= backoff

# ----------------------
def should_export(export_dir: str, frequency: str, config_mtime: float) -> tuple[bool, str]:

    base_name = os.path.basename(export_dir)
    latest_xlsx_mtime = None

    for fname in os.listdir(export_dir):
        if fname.startswith(f"##{base_name}-") and fname.endswith(".xlsx"):
            fpath = os.path.join(export_dir, fname)
            mtime = os.path.getmtime(fpath)
            if latest_xlsx_mtime is None or mtime > latest_xlsx_mtime:
                latest_xlsx_mtime = mtime

    if latest_xlsx_mtime is None:
        return True, "Keine .xlsx-Datei vorhanden"

    if config_mtime > latest_xlsx_mtime:
        config_time = datetime.fromtimestamp(config_mtime).strftime('%Y-%m-%d %H:%M:%S')
        xlsx_time = datetime.fromtimestamp(latest_xlsx_mtime).strftime('%Y-%m-%d %H:%M:%S')
        return True, (
            f"Config modified: "
            f"(INI: {config_time}, XLSX: {xlsx_time})"
        )
    
    now = datetime.now()
    last_export = datetime.fromtimestamp(latest_xlsx_mtime)
    frequency = frequency.lower().strip()

    readable_time = last_export.strftime('%Y-%m-%d %H:%M:%S')
    # Bedingung beschreiben
    if frequency == "hourly":
        next_export = last_export + timedelta(hours=1)
        if now > next_export:
            return True, f"hourly: last={readable_time}, next={next_export}, now={now}"
        else:
            return False, f"noExport: hourly until {next_export.strftime('%Y-%m-%d %H:%M:%S')}"
    
    elif frequency == "4hourly":
        next_export = last_export + timedelta(hours=4)
        if now > next_export:
            return True, f"4hourly: last={readable_time}, next={next_export}, now={now}"
        else:
            return False, f"noExport: 4hourly until {next_export.strftime('%Y-%m-%d %H:%M:%S')}"

    elif frequency in ("daily", "weekday"):
        next_export = last_export + timedelta(days=1)
        if now > next_export:
            return True, f"daily: last={readable_time}, next={next_export}, now={now}"
        else:
            return False, f"noExport: daily until {next_export.strftime('%Y-%m-%d %H:%M:%S')}"
    
    elif frequency == "weekly":
        next_export = last_export + timedelta(days=7)
        if now > next_export:
            return True, f"weekly: last={readable_time}, next={next_export}, now={now}"
        else:
            return False, f"noExport: weekly until {next_export.strftime('%Y-%m-%d')}"
    
    elif frequency == "monthly":
        next_month = (last_export.replace(day=1) + timedelta(days=32)).replace(day=1)
        if now > next_month:
            return True, f"monthly: last={last_export.strftime('%Y-%m')}, next={next_month.strftime('%Y-%m')}, now={now.strftime('%Y-%m-%d')}"
        else:
            return False, f"noExport: monthly until {next_month.strftime('%Y-%m-%d')}"

    elif frequency == "yearly":
        next_year = datetime(last_export.year + 1, 1, 1)
        if now > next_year:
            return True, f"yearly: last={last_export.year}, next={next_year.year}, now={now.year}"
        else:
            return False, f"noExport: yearly until {next_year.year}"
    return False, (
        f"noExport: last file {readable_time}, "
    )

async def get_dict_from_paperless(endpoint):
    """
    Generische Funktion, um ein Dictionary aus einem Paperless-Endpoint zu erstellen.
    Erwartet ein `endpoint`-Objekt, das eine `all()`-Methode und einen Abruf per ID unterstützt.
    """
    items = await retry_async(fn=lambda: endpoint.all())
    #items = await endpoint.all()
    item_dict = {}

    for itemKey in items:
        item = await endpoint(itemKey)
        item_dict[item.id] = item  # Speichert das gesamte Objekt mit der ID als Schlüssel

    return item_dict  # Gibt ein Dictionary {ID: Objekt} zurück
# Modulweiter Cache (z. B. ganz oben im Script)
_paperless_meta_cache = None

async def fetch_paperless_meta(paperless, force_reload=False):
    global _paperless_meta_cache

    if _paperless_meta_cache is not None and not force_reload:
        return _paperless_meta_cache

    def log_and_print(name):
        message(text=f"getting {name}...",inline=True)

    meta = {}

    for name, endpoint in {
        "storage_paths": paperless.storage_paths,
        "correspondents": paperless.correspondents,
        "document_types": paperless.document_types,
        "tags": paperless.tags,
        "users": paperless.users,
        "custom_fields": paperless.custom_fields
    }.items():
        message(f"{name}",inline=True)
        try:
            meta[name] = await get_dict_from_paperless(endpoint)
            message(f"{name.capitalize()}: {len(meta[name])}")
        except Exception as e:
            message( f"Fehler beim Abrufen von {name}: {e}")
            meta[name] = []  # Leere Liste als Fallback, damit getmeta nicht crasht

    _paperless_meta_cache = meta
    return meta

# ----------------------
def getmeta(key, doc, meta):
    """
    Holt den Wert aus den Metadaten basierend auf dem angegebenen Schlüssel und Dokument.
    Hier wird doc als Objekt behandelt.

    :param key: Der Schlüssel, nach dem in den Metadaten gesucht wird (z. B. "document_type").
    :param doc: Das Dokument-Objekt, das das Attribut enthält (z. B. doc.document_type).
    :param meta: Die Metadatenstruktur, die die Daten enthält.
    :return: Der Name des Dokuments, falls vorhanden, oder 'Unbekannt', falls ein Fehler auftritt.
    """
    try:
        # Hole den Wert des Schlüssels aus doc als Attribut (z. B. doc.document_type)
        index = getattr(doc, key, None)

        if key == "tags" and isinstance(index, list):  # Spezieller Fall für tags (Liste von Indizes)
            # Generiere den Tag-String für mehrere Tags
                # Extrahiere die Tag-Namen basierend auf den Indizes in 'index'
            return ", ".join(
            meta["tags"][tag_id].name  # Greift auf den Namen des Tags mit der ID aus index zu
                for tag_id in index)

        # Wenn der Index gefunden wurde und der Index gültig ist
        if index is not None and 0 <= index < len(meta.get(f"{key}s", [])):
            # Hole das entsprechende Element aus meta und gebe dessen "name" zurück
            return meta[f"{key}s"][index].name
        else:
            return 'Unbekannt'  # Falls der Index ungültig oder nicht vorhanden ist
    except KeyError:
        return 'Unbekannt'  # Falls der Schlüssel nicht existiert
    except Exception as e:
        print(f"Fehler beim Abrufen von {key}: {e}")
        return 'Unbekannt'

async def export_pdf(doc, working_dir):
    """Exportiert ein Dokument als PDF mit automatischem Retry."""
    sanitized_title = sanitize_filename(doc.title)
    filename = f"{doc.id}--{sanitized_title}.pdf"
    pdf_path = os.path.join(working_dir, filename)

    download = await retry_async(lambda: doc.get_download(), desc=f"PDF-Download für Dokument {doc.id}")
    document_content = download.content

    if not document_content:
        print(f"Keine PDF-Daten für Dokument {doc.id} gefunden.")
        return

    with open(pdf_path, 'wb') as f:
        f.write(document_content)

# ----------------------
def sanitize_filename(filename):
    """
    Remove or replace characters in the filename that are not allowed in file names.
    """
    sanitized = re.sub(r'[<>:"/\\|?*]', '-', filename)  # Ersetze verbotene Zeichen durch '-'
    return sanitized[:255]  # Truncate to avoid overly long filenames

# ----------------------

def get_document_json(paperless,doc):
    api_token = paperless._token # Dein-Token
    headers = {"Authorization": f"Token {api_token}"}

    path=doc._api_path 
    url=paperless._base_url

    """Retrieve detailed document metadata from Paperless API."""
    response = requests.get(f"{url}/{path}", headers=headers)
    
    if response.status_code == 200:
        return response.json()  # Die JSON-Daten des Dokuments zurückgeben
    else:
        raise Exception(f"Failed to fetch document metadata: {response.status_code}")

def export_json(paperless, doc, working_dir):
    """Export a document's metadata as JSON."""
    sanitized_title = sanitize_filename(doc.title)
    filename = f"{doc.id}--{sanitized_title}.json"
    json_path = os.path.join(working_dir, filename)

    detailed_doc = get_document_json(paperless=paperless, doc=doc)
    with open(json_path, "w", encoding="utf-8") as json_file:
        json.dump(detailed_doc, json_file, ensure_ascii=False, indent=4)

# ---------------------- Excel Export Helpers ----------------------
# ----------------------
def export_to_excel(data, file_path, script_name, currency_columns, dir, url, meta,maxfiles, query, frequency):
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.utils import get_column_letter
    from datetime import datetime
    import os
    import pwd

    # API-Basis-URL ohne `/api` generieren
    #base_url = api_url.rstrip("/api")
    base_url = url

    # Ordnerpfad aus file_path extrahieren
    directory = os.path.dirname(file_path)
    cleanup_old_files(file_path, filename_prefix="##" + directory ,pattern="xlsx",max_count_str=maxfiles)

    # Dateiname vorbereiten
    fullfilename = file_path
    # Dateiname vorbereiten (immer mit -0 starten)
    filename_without_extension, file_extension = os.path.splitext(os.path.basename(file_path))
    base_filename = f"{filename_without_extension}-0{file_extension}"
    fullfilename = os.path.join(directory, base_filename)

    # Falls Datei bereits geöffnet oder existiert, iterativ neuen Namen finden
    counter = 1
    while os.path.exists(fullfilename):
        filename = f"{filename_without_extension}-{counter}{file_extension}"
        fullfilename = os.path.join(directory, filename)
        counter += 1

    # Pandas DataFrame aus document_data erstellen
    df = pd.DataFrame(data)
    if df.empty:
        print(f"[INFO] Keine Daten gefunden, erstelle leere Excel-Datei mit Platzhalter.")
        df = pd.DataFrame(columns=["Keine Daten vorhanden"])


    with pd.ExcelWriter(fullfilename, engine="openpyxl") as writer:
        # DataFrame in Excel schreiben (ab Zeile 3 für Daten)
        df.to_excel(writer, index=False, startrow=2, sheet_name="Dokumentenliste")
        worksheet = writer.sheets["Dokumentenliste"]

        # Headerzeile (A1) mit Scriptnamen, Tag und anderen Infos
        header_info = f"{script_name} -- {directory} -- {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} -- {pwd.getpwuid(os.getuid()).pw_name} -- {os.uname().nodename}"
        worksheet["A1"] = header_info
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))  # Header über alle Spalten
        header_font = Font(bold=True, color="FFFFFF", name="Arial")
        header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")  # Dunkelblau
        worksheet["A1"].font = header_font
        worksheet["A1"].fill = header_fill

        # Summenzeilen für Currency-Spalten in Zeile 2
        for column_name in currency_columns:
            if column_name in df.columns:
                col_idx = df.columns.get_loc(column_name) + 1  # Excel-Spaltenindex
                start_cell = worksheet.cell(row=4, column=col_idx).coordinate
                end_cell = worksheet.cell(row=worksheet.max_row, column=col_idx).coordinate
                sum_formula = f"=SUM({start_cell}:{end_cell})"
                sum_cell = worksheet.cell(row=2, column=col_idx)
                sum_cell.value = sum_formula
                sum_cell.font = Font(bold=True)

        # Spaltentitel (Zeile 3)
        header_row = worksheet[3]
        for cell in header_row:
            cell.font = Font(bold=True, color="FFFFFF", name="Arial")
            cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

        # Autofilter
        worksheet.auto_filter.ref = f"A3:{worksheet.cell(row=3, column=len(df.columns)).coordinate}"

        # Definiere die Formate für gerade und ungerade Zeilen
        light_blue_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        font = Font(name="Arial", size=11)

        # Formeln für gerade und ungerade Zeilen
        formula_even = "MOD(ROW(),2)=0"
        formula_odd = "MOD(ROW(),2)<>0"

        # Nur wenn Datenzeilen existieren (mind. Zeile 4)
        if worksheet.max_row >= 4:
            range_string = f"A4:{worksheet.cell(row=worksheet.max_row, column=len(df.columns)).coordinate}"

            # Gerade Zeilen
            rule_even = FormulaRule(formula=["MOD(ROW(),2)=0"], fill=light_blue_fill, font=font)
            worksheet.conditional_formatting.add(range_string, rule_even)

            # Ungerade Zeilen
            rule_odd = FormulaRule(formula=["MOD(ROW(),2)<>0"], fill=white_fill, font=font)
            worksheet.conditional_formatting.add(range_string, rule_odd)

        # Hyperlinks in der ID-Spalte
        # Suche die Spalte basierend auf dem Header in Zeile 3
        document_column = "ID"  # Der Header-Name für die Spalte mit den Dokument-IDs
        id_column_idx = None
        for col_idx, cell in enumerate(worksheet[3], start=1):  # Zeile 3 ist der Header
            if cell.value == document_column:
                id_column_idx = col_idx
                break

        # Dokument-ID in URLs umwandeln
        if id_column_idx:  # Wenn die Spalte mit der ID gefunden wurde
            for row_idx in range(4, worksheet.max_row + 1):  # Daten beginnen in Zeile 4
                doc_id = worksheet.cell(row=row_idx, column=id_column_idx).value
                if doc_id:  # Nur wenn ein Wert vorhanden ist
                    link_formula = f'=HYPERLINK("{base_url}/documents/{doc_id}/details", "{doc_id}")'
                    worksheet.cell(row=row_idx, column=id_column_idx).value = link_formula

        # Schriftart-Objekt definieren
        default_font = Font(name="Arial")

        # Alle Zellen formatieren
        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.font = default_font

    # Metadaten anhängen
    append_metadata_sheet(
        xlsx_path=fullfilename,
        script_version=get_git_version(),
        config_data={"query": query, "frequency": frequency}
    )
    message(f"\nExcel-Datei erfolgreich erstellt: {fullfilename}")

# ----------------------
def has_file_from_today(directory):
    """
    Prüft, ob im angegebenen Verzeichnis eine Datei existiert,
    die heute erstellt oder zuletzt geändert wurde.
    """
    today = datetime.now().date()
    if not os.path.exists(directory):
        return False

    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        if os.path.isfile(file_path):
            # Änderungszeitpunkt der Datei
            file_mtime = datetime.fromtimestamp(os.path.getmtime(file_path))
            if file_mtime.date() == today:
                return True
    return False

# ----------------------
def process_custom_fields(meta, doc):
    custom_fields = {}
    currency_fields = []  # Liste zum Speichern der Currency-Feldnamen

    if "custom_fields" in doc:
        for custom_field in doc["custom_fields"]:
            field_id = custom_field.get("field")
            if not field_id:    
                continue
            field_name = meta["custom_fields"][field_id].name
            field_type = meta['custom_fields'][field_id]._data['data_type']
            field_value = custom_field.get("value")
            
            if field_type == "monetary":
                numeric_value = parse_currency(field_value)
                custom_fields[field_name] = numeric_value  # Rohdaten speichern
                custom_fields[f"{field_name}_formatted"] = format_currency(field_value)  # Formatierte Version speichern
                currency_fields.append(field_name)  # Speichern des Currency-Felds
            elif field_type == "select":
                # Hole die choices aus meta["custom_fields"][field_id] und prüfe, ob field_value None ist
                choices = meta['custom_fields'][field_id]._data['extra_data']['select_options']
                if field_value is None:
                    custom_fields[field_name] = "none"  # Wenn field_value None ist, setze "none"
                else:
                    # Wenn field_value nicht None ist, hole den Wert aus den choices
                    custom_fields[field_name] =  choices[field_value]

            else:
                custom_fields[field_name] = field_value

    return custom_fields, currency_fields

async def get_documents_with_retry(paperless, query):
    return await retry_async(
        lambda: collect_async_iter(paperless.documents.search(query)),
        desc=f"Dokumente für Query '{query}'"
    )

async def collect_async_iter(aiter):
    return [item async for item in aiter]

async def search_documents(paperless, query):
    return [item async for item in paperless.documents.search(query)]

def find_cached_file(doc_id, all_dir, kind):
    """Findet Datei im .all-Verzeichnis anhand von doc_id und Dateityp ('pdf' oder 'json')"""
    prefix = f"{doc_id}--"
    suffix = f".{kind}"
    for fname in os.listdir(all_dir):
        if fname.startswith(prefix) and fname.endswith(suffix):
            return os.path.join(all_dir, fname)
    return None

import os
import shutil

def link_export_file(doc, kind, working_dir, all_dir=".all"):
    assert kind in ("pdf", "json")

    filename = f"{doc.id}--{sanitize_filename(doc.title)}.{kind}"
    dest_path = os.path.join(working_dir, filename)

    # Quelle im .all-Ordner finden
    message(f"DEBUG: Suche {kind}-Datei von {doc.id} in {all_dir}", target="both")
    src_path = find_cached_file(doc.id, all_dir=all_dir, kind=kind)
    if src_path is None:
        raise FileNotFoundError(f"Keine {kind.upper()}-Datei für Dokument {doc.id} im .all-Verzeichnis gefunden")

    os.makedirs(os.path.dirname(dest_path), exist_ok=True)

    #message(f"from:  {src_path}", "both")
    #message(f"to:    {dest_path}", "both")

    message(f"# Wenn Zieldatei{dest_path} existiert, prüfen ob korrekt",target ="both")
    if os.path.exists(dest_path):
        try:
            if os.path.islink(dest_path) and os.path.realpath(dest_path) == os.path.realpath(src_path):
                message("symlinkg OK")
                return "symlink (OK)"
            elif os.path.samefile(dest_path, src_path):
                message("hardlink/copy ok")
                return "hardlink/copy (OK)"
            else:
                os.remove(dest_path)
        except Exception:
            os.remove(dest_path)

    # Versuch: Symlink
    try:
        os.symlink(src_path, dest_path)
        if os.path.exists(dest_path) and os.path.getsize(dest_path) > 0:
            message("symlink (neu)")
            return "symlink (neu)"
    except Exception as e:
        message(f"Symlink fehlgeschlagen: {e}", "both")

    # Versuch: Hardlink
    try:
        os.link(src_path, dest_path)
        if os.path.exists(dest_path) and os.path.getsize(dest_path) > 0:
            message("hardlink (neu)")
            return "hardlink (neu)"
    except Exception as e:
        message(f"Hardlink fehlgeschlagen: {e}", "both")

    # Fallback: Datei kopieren
    try:
        shutil.copy2(src_path, dest_path)
        if os.path.exists(dest_path) and os.path.getsize(dest_path) > 0:
            message("copy (neu)")
            return "copy (neu)"
    except Exception as e:
        message(f"Kopie fehlgeschlagen: {e}" )
        raise RuntimeError(f"Datei konnte weder verlinkt noch kopiert werden: {src_path}")

    raise RuntimeError(f"Zieldatei konnte nicht erstellt werden: {dest_path}")

def XXXlink_export_file(doc, kind, working_dir, all_dir=".all"):
    assert kind in ("pdf", "json")

    filename = f"{doc.id}--{sanitize_filename(doc.title)}.{kind}"
    dest_path = os.path.join(working_dir, filename)

    # Quelle im .all-Ordner finden
    message(f"DEBUG: Suche {kind}-Datei von {doc.id} in {all_dir}",target="both")

    src_path = find_cached_file(doc.id, all_dir=all_dir, kind=kind)
    if src_path is None:
        raise FileNotFoundError(f"Keine {kind.upper()}-Datei für Dokument {doc.id} im .all-Verzeichnis gefunden")

    # Zielverzeichnis sicherstellen
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)

    message(f"from:  {src_path}","both") 
    message(f"to:    {dest_path}","both") 

    # Wenn Zieldatei existiert → prüfen oder entfernen
    if os.path.exists(dest_path):
        try:
            # Ist es ein Symlink und zeigt korrekt?
            if os.path.islink(dest_path):
                if os.path.realpath(dest_path) == os.path.realpath(src_path):
                    return "symlink (OK)"
                else:
                    os.remove(dest_path)
            # Ist es eine Hardlink oder echte Datei identisch?
            elif os.path.samefile(dest_path, src_path):
                return "hardlink/copy (OK)"
            else:
                os.remove(dest_path)
        except Exception:
            os.remove(dest_path)

    # Jetzt sauber: versuch Symlink → Hardlink → Copy
    try:
        rel_path = os.path.relpath(src_path, os.path.dirname(dest_path))
        os.symlink(rel_path, dest_path)
        return "symlink"
    except FileExistsError:
        os.remove(dest_path)
        return link_export_file(doc, kind, working_dir, all_dir)
    except OSError:
        try:
            os.link(src_path, dest_path)
            return "hardlink"
        except FileExistsError:
            os.remove(dest_path)
            return link_export_file(doc, kind, working_dir, all_dir)
        except OSError:
            message(f"shutil.copy2({src_path}, {dest_path}")
            shutil.copy2(src_path, dest_path)
            return "copy"


async def exportThem(paperless, dir, query, max_files, frequency):
    count = 0 
    """Process and export documents"""
    document_data = []
    currency_columns = []  # Liste zur Speicherung aller Currency-Felder
    custom_fields = {}
    meta = await fetch_paperless_meta(paperless)

#    documents = [item async for item in paperless.documents.search(query)]
    documents = await retry_async(
       lambda: search_documents(paperless, query),
       desc=f"Dokumente für Query '{query}'"
       )

    #documents = await retry_async(
    #   lambda: collect_async_iter(paperless.documents.search(query)),
    #    desc=f"Dokumente für Query '{query}'"
    #)

    for doc in tqdm(documents, desc=f"Processing documents for '{dir} with {query}'", unit="doc"):
        count += 1

        try:
          metadata = None
          metadata = await retry_async(lambda: doc.get_metadata(), desc=f"Metadaten für Dokument {doc.id}")
        except Exception as e:
            print(f"Metadaten für Dokument {doc.id} konnten nicht geladen werden: {e}. Überspringe dieses Dokument.")
            continue

        docData = doc._data
        page_count = docData['page_count']
        message(f"{doc.id} {doc.title} {page_count} pages",target="log")

        custom_fields, doc_currency_columns = process_custom_fields(meta=meta,doc=docData)
        currency_columns.extend(doc_currency_columns)  # Speichere Currency-Felder
        thisTags =  getmeta("tags", doc, meta=meta)

        # Daten für die Excel-Tabelle sammeln
        row = OrderedDict([
            ("ID", doc.id),
            (  "AddDateFull", format_date(parse_date(doc.added), "yyyy-mm-dd")),
            ("Korrespondent", meta["correspondents"][doc.correspondent].name),
            ("Titel", doc.title),
            ("Tags", thisTags), 

            # Custom Fields direkt hinter den Tags Pieinfügen
            *custom_fields.items(),  

            ("ArchivDate", parse_date(doc.created)),
            ("ArchivedDateMonth", format_date(parse_date(doc.created), "yyyy-mm")),
            ("ArchivedDateFull", format_date(parse_date(doc.created), "yyyy-mm-dd")),
            ("ModifyDate", parse_date(doc.modified)),
            ("ModifyDateMonth", format_date(parse_date(doc.modified), "yyyy-mm")),
            ("ModifyDateFull", format_date(parse_date(doc.modified), "yyyy-mm-dd")),
            ("AddedDate", parse_date(doc.added)),
            ("AddDateMonth", format_date(parse_date(doc.added), "yyyy-mm")),
            ("AddDateFull", format_date(parse_date(doc.added), "yyyy-mm-dd")),
            ("Seiten", doc._data['page_count']),
            ("Dokumenttyp", getmeta("document_type", doc, meta)),
            ("Speicherpfad", getmeta("storage_path", doc, meta)),
            ("OriginalName", doc.original_file_name),
            ("ArchivedName", doc.archived_file_name),
            ("Owner", getattr(meta["users"].get(doc.owner), "username", "Unbekannt") if doc.owner else "Unbekannt")
        ]
        )

        document_data.append(row)

        # Exportiere das PDF des Dokuments
        #await export_pdf(doc, working_dir=dir)
        #export_json(paperless=paperless,doc=doc,working_dir=dir)
        # Statt export_pdf / export_json:
        export_dir  = os.path.dirname(dir)

        method_pdf = link_export_file(doc, kind="pdf", working_dir=dir, all_dir=os.path.join(export_dir, ".all")) 
        method_json = link_export_file(doc, kind="json", working_dir=dir, all_dir=os.path.join(export_dir, ".all"))

        message(f"{doc.id}: PDF → {method_pdf}",target="both")
        message(f"{doc.id}: json → {method_json}", target="both")


    # Exportiere die gesammelten Daten nach Excel
    #path=doc._api_path 
    url=paperless._base_url

    last_dir = os.path.basename(dir)

    excel_file = os.path.join(dir, f"##{last_dir}-{datetime.now().strftime('%Y%m%d')}.xlsx")
    export_to_excel(document_data, excel_file, get_script_name, currency_columns=currency_columns,dir=dir, url=url,meta=meta, maxfiles=max_files,query=query, frequency=frequency)
#    log_message(progress_log_path, f"dir: {dir}, Documents exported: {len(document_data)}")
#    print(f"Exported Excel file: {excel_file}")

async def single_build_all_cache(paperless, export_dir, log_path=None):
    def log(msg):
        if log_path:
            message( msg)

    all_dir = os.path.join(export_dir, ".all")
    os.makedirs(all_dir, exist_ok=True)

    doc_ids = await retry_async(lambda: paperless.documents.all(), desc="Lade Dokument-IDs")

    done = 0
    cached = 0
    bar = tqdm(doc_ids, desc="Dokumente cachen: 0✓ / 0↓", unit="doc")

    for doc_id in bar:
        doc = await retry_async(lambda: paperless.documents(doc_id), desc=f"Hole Dokument {doc_id}")

        sanitized_title = sanitize_filename(doc.title)
        pdf_filename = f"{doc.id}--{sanitized_title}.pdf"
        json_filename = f"{doc.id}--{sanitized_title}.json"
        pdf_path = os.path.join(all_dir, pdf_filename)
        json_path = os.path.join(all_dir, json_filename)

        updated = False

        if not os.path.exists(pdf_path):
            await export_pdf(doc, working_dir=all_dir)
            updated = True
        if not os.path.exists(json_path):
            export_json(paperless=paperless, doc=doc, working_dir=all_dir)
            updated = True

        if updated:
            done += 1
        else:
            cached += 1

        bar.set_description(f"Dokumente cachen: {cached}✓ / {done}↓")

    message(f"Cache abgeschlossen: {done} neu, {cached} übersprungen.")

from pypaperless.models.generators.page import Page

async def safe_document_iterator(paperless):
    page_iter = aiter(paperless.documents.pages())

    while True:
        try:
            page: Page = await retry_async(lambda: anext(page_iter), desc="Lade Dokument-Seite")
        except StopAsyncIteration:
            break
        except Exception as e:
            message(f"Fehler beim Laden einer Seite: {e}")
            break

        for doc in page.items:
            yield doc

async def build_all_cache(paperless, export_dir, log_path=None):
    def log(msg):
        if log_path:
            message( msg)

    all_dir = os.path.join(export_dir, ".all")
    os.makedirs(all_dir, exist_ok=True)

    # Schneller Count über .all()
    doc_ids = await retry_async(lambda: paperless.documents.all(), desc="Zähle Dokumente")
    total = len(doc_ids)

    done = 0
    cached = 0
    bar = tqdm(total=total, desc="Dokumente cachen: 0✓ / 0↓", unit="doc")

    async for doc in safe_document_iterator(paperless):
        sanitized_title = sanitize_filename(doc.title)
        pdf_filename = f"{doc.id}--{sanitized_title}.pdf"
        json_filename = f"{doc.id}--{sanitized_title}.json"
        pdf_path = os.path.join(all_dir, pdf_filename)
        json_path = os.path.join(all_dir, json_filename)

        updated = False

        if not os.path.exists(pdf_path):
            await export_pdf(doc, working_dir=all_dir)
            updated = True
        if not os.path.exists(json_path):
            export_json(paperless=paperless, doc=doc, working_dir=all_dir)
            updated = True

        if updated:
            done += 1
        else:
            cached += 1

        bar.update(1)
        bar.set_description(f"Dokumente cachen: {cached}✓ / {done}↓")

    bar.close()
    message(f"Cache abgeschlossen: {done} neu, {cached} übersprungen.")

def extract_doc_id(filename):
    """Extrahiere die Dokument-ID aus einem Dateinamen wie '874--irgendwas.pdf'."""
    try:
        return int(filename.split('--', 1)[0])
    except (ValueError, IndexError):
        return None

async def cleanup_all_dir(paperless, all_dir, log_path=None):
    message( "Bereinige .all-Verzeichnis...")

    # Aktuelle Dokument-IDs abrufen
    valid_doc_ids = set(await retry_async(lambda: paperless.documents.all(), desc="Lade gültige Dokument-IDs"))

    removed_files = 0
    for filename in os.listdir(all_dir):
        if not (filename.endswith(".pdf") or filename.endswith(".json")):
            continue

        doc_id = extract_doc_id(filename)
        if doc_id is None or doc_id not in valid_doc_ids:
            file_path = os.path.join(all_dir, filename)
            try:
                os.remove(file_path)
                removed_files += 1
                message(f"Entfernt: {filename}")
            except Exception as e:
                message(message=f"Fehler beim Löschen von {filename}: {e}")

    message(text=f"Bereinigung abgeschlossen: {removed_files} Datei(en) entfernt.")

def is_remote_newer(remote_modified_str, local_path):
    # Vergleiche Timestamps
    remote_time = datetime.datetime.fromisoformat(remote_modified_str)
    local_time = datetime.datetime.fromtimestamp(os.path.getmtime(local_path))
    return remote_time > local_time

async def main():
    print_program_header()
    print_separator('=', 0.75)

    script_name = get_script_name()
    config = load_config_from_script()

    # 🔍 Pflichtfelder prüfen
    required = {
        "API": ["url", "token"],
        "Export": ["directory"],
        "Log": ["log_file", "max_files"]
    }

    missing = []
    for section, keys in required.items():
        if not config.has_section(section):
            missing.append(f"[{section}] fehlt")
            continue
        for key in keys:
            if not config.has_option(section, key) or not config.get(section, key).strip():
                missing.append(f"{section}.{key} fehlt oder ist leer")

    if missing:
        print("\n❌ Fehler in der Konfigurationsdatei:")
        for m in missing:
            print(f"   - {m}")
        print("\n💡 Bitte prüfe deine .ini-Datei und ergänze die fehlenden Angaben.")
        sys.exit(1)

    # ✅ Nur wenn alles ok ist: Konfigurationswerte lesen
    export_dir = config.get("Export", "directory")
    api_url = config.get("API", "url")
    api_token = config.get("API", "token")
    log_dir = config.get("Log", "log_file")
    max_files = config.get("Log", "max_files")

    # Log-Dateien initialisieren
    prepare_logging(log_dir, script_name, max_files)

    # Beispielverwendung nach Initialisierung:
    # set_log_path("pfad/zur/logdatei.log")
    # message("Export gestartet", target="both")
    # message("Dokument konnte nicht geladen werden", level="warn")
    # message("Verarbeite Dokument 17", inline=True)

# Setze das Arbeitsverzeichnis auf das Verzeichnis, in dem das Skript gespeichert ist
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    os.chdir(script_dir)

# Setze Locale früh im Programm (fallback auf Standard-Locale)
    try:
        locale.setlocale(locale.LC_ALL, 'de_DE.UTF-8')
    except locale.Error:
        try:
            locale.setlocale(locale.LC_ALL, '')
        except:
            print("⚠️ Locale konnte nicht gesetzt werden – Formatierung ggf. fehlerhaft.")

    message(f"Log in to {api_url} with {mask_secret(api_token)} ..", target="both")
    paperless = Paperless(api_url, api_token)
    await retry_async(lambda: paperless.initialize(), desc="Paperless-Login")
    
    message("Logged in", target="both")
    message(f"Export to {export_dir}", target="both")

    # do something
    meta = await fetch_paperless_meta(paperless)
    # Zugriff auf ein Element
    #print(meta["correspondents"][3].name)
    #print(meta["tags"][1].name)
    #print(meta["storage_paths"][2].name)

    await build_all_cache(paperless, export_dir)
    await cleanup_all_dir(paperless, all_dir=os.path.join(export_dir, ".all"))

    excluded_dirs = {"@eaDir", ".all"}
    try:
        for root, dirs, files in os.walk(export_dir):
          query_value = os.path.basename(root)
          dirs[:] = [d for d in dirs if d not in excluded_dirs]
          if root == export_dir:
            continue

          config_mtime = 0  # oder datetime.min.timestamp()
          if '##config.ini' in files:
              config_path = os.path.join(root, '##config.ini')
              config = configparser.ConfigParser()
           #   config.read(config_path)
              config=load_config(config_path=config_path)
              config_mtime = os.path.getmtime(config_path)

          if 'DATA' in config and 'query' in config['DATA']:
              query_value = config['DATA']['query']

          if 'EXPORT' in config and 'frequency' in config['EXPORT']:
              frequency = config['EXPORT']['frequency']
          else:
              frequency = 'daily'

          should_run, reason = should_export(root, frequency, config_mtime)
          if should_run:
              #print_separator('#')           # #######...
              #print_separator('##')          # ## ## ## ...
              #print_separator('=')           # ==========...
              #print_separator('·', 0.5)      # 50% der Breite
              print_separator('=', 0.75)      # 50% der Breite
              #print(f"\n{root} {query_value} -> Export ({reason})")
              message(f"{root} : {query_value} -> ({reason})", target="both")
              await exportThem(paperless=paperless, dir=root, query=query_value, max_files=max_files,frequency=frequency)
          else:
              #print(f"\n{root} {query_value} -> NOexport ({reason})")
              print_separator('-', 0.75)      # 50% der Breite
              message(f"{root} : {query_value} -> ({reason})", target="both")

    except Exception as e:
        message(f"Error: {str(e)}", target="both")
        raise
    finally:
        if paperless:
            await paperless.close()
        #finalize_log(progress_log_path, final_log_path)

asyncio.run(main())
