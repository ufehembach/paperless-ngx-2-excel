#!/usr/bin/env python3

import os
import sys
import pwd
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
from time import sleep
import asyncio
import aiohttp
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

def cleanup_old_logs(log_dir, script_name, max_logs_str):
    """ Löscht alte Logs, wenn die Anzahl der Log-Dateien das Limit überschreitet. """
    
    # Debug: Zeige den Log-Verzeichnispfad
    print(f"Log-Verzeichnis: {log_dir}")
    
    # Erstelle ein verbessertes Glob-Muster für .log- und .progress.log-Dateien
    #log_files = sorted(glob.glob(os.path.join(log_dir, f"##{script_name}__*.log")) + 
    #                   glob.glob(os.path.join(log_dir, f"##{script_name}__*.progress.log")),
    #                   key=os.path.getmtime)
    log_files = sorted(glob.glob(os.path.join(log_dir, f"##{script_name}__*.log")) ,
                       key=os.path.getmtime)
    
    # Debug: Zeige die gefundenen Log-Dateien
    #print(f"Gefundene Log-Dateien: {log_files}")
    
    max_logs = int(max_logs_str)
    if len(log_files) <= max_logs:
        print(f"Anzahl der Log-Dateien ({len(log_files)}) ist kleiner oder gleich dem Limit ({max_logs}). Keine Dateien zum Löschen.")
        return  # Wenn die Anzahl der Dateien innerhalb des Limits liegt, keine Dateien löschen
    
    # Lösche alte Dateien, falls die Anzahl die Grenze überschreitet
    while len(log_files) > max_logs:
        old_file = log_files.pop(0)  # Holen der ältesten Datei
        if os.path.exists(old_file):  # Überprüfen, ob die Datei existiert
            try:
                os.remove(old_file)  # Lösche die Datei
                print(f"Alte Log-Datei gelöscht: {old_file}")
            except OSError as e:
                print(f"Fehler beim Löschen der Datei {old_file}: {e}")
        else:
            print(f"Datei {old_file} existiert nicht mehr. Überspringe das Löschen.")

def get_log_filename(script_name, log_dir, suffix="progress"):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    if suffix == "log":
        return os.path.join(log_dir, f"##{script_name}__{timestamp}.log")
    else:
        return os.path.join(log_dir, f"##{script_name}__{timestamp}.{suffix}.log")

def initialize_log(log_dir, script_name, max_logs):
    final_log_path = get_log_filename(script_name, log_dir, "log")
    progress_log_path = get_log_filename(script_name, log_dir, "progress")
    
    # Falls ein vorheriges Log existiert, es in die neue Log-Datei kopieren
    if os.path.exists(final_log_path):
        with open(progress_log_path, "w") as new_log, open(final_log_path, "r") as old_log:
            shutil.copyfileobj(old_log, new_log)
        os.remove(final_log_path)
    else:
        open(progress_log_path, "w").close()  # Erstelle eine leere Log-Datei
    
    # Aufräumen: Älteste Logs löschen, falls nötig
    cleanup_old_logs(log_dir, script_name, max_logs)

    return progress_log_path, final_log_path

# Funktion, um das Log umzubenennen
def finalize_log(progress_log_path, final_log_path):
    if os.path.exists(progress_log_path):
        os.rename(progress_log_path, final_log_path)

def print_progress(message: str):
    frame = inspect.currentframe().f_back
    filename = os.path.basename(frame.f_code.co_filename)
    line_number = frame.f_lineno
    function_name = frame.f_code.co_name

    progress_message = f"{filename}:{line_number} [{function_name}] {message}"

    if not hasattr(print_progress, "_last_length"):
        print_progress._last_length = 0

    clear_space = max(print_progress._last_length - len(progress_message), 0)
    progress_message += " " * clear_space

    sys.stdout.write(f"\r{progress_message}")
    sys.stdout.flush()

    print_progress._last_length = len(progress_message)

# ---------------------- Configuration Loading ----------------------
def load_config(config_path):
    """Load configuration file."""
    print_progress("process...")
    config = ConfigParser()
    config.read(config_path)
    return config

def get_script_name():
    """Return the name of the current script without extension."""
    return os.path.splitext(os.path.basename(sys.argv[0]))[0]

def load_config_from_script():
    """Load the configuration from the ini file with a priority for the .ufe.ini file."""
    script_name = get_script_name()
    ufe_ini_path = f"{script_name}.ufe.ini"
    ini_path = f"{script_name}.ini"

    # Try to load the .ufe.ini file first
    if os.path.exists(ufe_ini_path):
        print_progress(f"Using config file: {ufe_ini_path}")
        return load_config(ufe_ini_path)
    # Fallback to the .ini file
    elif os.path.exists(ini_path):
        print_progress(f"Using config file: {ini_path}")
        return load_config(ini_path)
    else:
        print(f"Configuration files '{ufe_ini_path}' and '{ini_path}' not found.")
        sys.exit(1)


# ---------------------- Logging ----------------------
def log_message(log_path, message):
    """Append a log message to the log file."""
    with open(log_path, "a") as log_file:
        log_file.write(f"{datetime.now()} - {message}\n")

def parse_currency(value):
    """Parst einen Währungswert wie 'EUR5.00' in einen Float."""
    try:
        # Entferne Währungszeichen (alles außer Ziffern, Punkt oder Minus)
        numeric_part = ''.join(c for c in value if c.isdigit() or c == '.' or c == '-')
        return float(numeric_part)
    except Exception as e:
        # print(f"Fehler beim Parsen des Währungswerts '{value}': {e}")
        return 0.0  # Fallback auf 0 bei Fehlern

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

    locale.setlocale(locale.LC_ALL, currency_locale)
    formatted_value = locale.currency(value_float, grouping=True)
    return formatted_value

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
from datetime import datetime
from dateutil import parser

def parse_date(date_input):
    """
    Gibt das Datum im Format '%d.%m.%Y' zurück, wenn Uhrzeit 00:00 ist,
    sonst im Format '%d.%m.%Y %H:%M'. Akzeptiert Strings oder datetime-Objekte.
    """
    if not date_input:
        print_progress(f"[parse_date] Date input is empty or None: {date_input}")
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
        print_progress(f"[parse_date] Failed to parse date '{date_input}': {e}")
        return None


async def get_dict_from_paperless(endpoint):
    """
    Generische Funktion, um ein Dictionary aus einem Paperless-Endpoint zu erstellen.
    Erwartet ein `endpoint`-Objekt, das eine `all()`-Methode und einen Abruf per ID unterstützt.
    """
    items = await endpoint.all()
    item_dict = {}

    for itemKey in items:
        item = await endpoint(itemKey)
        item_dict[item.id] = item  # Speichert das gesamte Objekt mit der ID als Schlüssel

    return item_dict  # Gibt ein Dictionary {ID: Objekt} zurück
# Modulweiter Cache (z. B. ganz oben im Script)
_paperless_meta_cache = None

async def fetch_paperless_meta(paperless, log_path, force_reload=False):
    global _paperless_meta_cache

    if _paperless_meta_cache is not None and not force_reload:
        return _paperless_meta_cache

    def log_and_print(name):
        log_message(log_path, f"getting {name}...")
        print_progress(message=f"getting {name}...")

    meta = {}

    for name, endpoint in {
        "storage_paths": paperless.storage_paths,
        "correspondents": paperless.correspondents,
        "document_types": paperless.document_types,
        "tags": paperless.tags,
        "users": paperless.users,
        "custom_fields": paperless.custom_fields
    }.items():
        log_and_print(name)
        meta[name] = await get_dict_from_paperless(endpoint)
        print(f"{name.capitalize()}: {len(meta[name])}")

    _paperless_meta_cache = meta
    return meta

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
            return ", ".join(
                getmeta("tags", {"tags": tag_id}, meta) for tag_id in index
            )

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

async def export_pdf(doc, working_dir, max_retries=3, retry_delay=5):
    """Export a document's PDF with retry mechanism."""
    sanitized_title = sanitize_filename(doc.title)
    pdf_path = os.path.join(working_dir, f"{sanitized_title}.pdf")

    # Retry-Mechanismus
    for attempt in range(max_retries):
        try:
            download = await doc.get_download()
            document_content = download.content

            if not document_content:
                print(f"Keine PDF-Daten für Dokument {doc.id} gefunden.")
                return

            # Speichern des heruntergeladenen Dokuments
            with open(pdf_path, 'wb') as f:
                f.write(document_content)

            print(f"PDF für Dokument {doc.id} erfolgreich exportiert: {pdf_path}")
            return  # Erfolg, beende die Funktion

        except aiohttp.client_exceptions.ServerDisconnectedError as e:
            print(f"Serververbindung für Dokument {doc.id} unterbrochen (Versuch {attempt + 1}/{max_retries}). {e}")
        except aiohttp.client_exceptions.ClientConnectionError as e:
            print(f"Verbindungsfehler für Dokument {doc.id} (Versuch {attempt + 1}/{max_retries}). {e}")
        except Exception as e:
            print(f"Fehler beim Exportieren der PDF für Dokument {doc.id}: {e} (Versuch {attempt + 1}/{max_retries})")
        
        # Warten und dann erneut versuchen
        print(f"Warte {retry_delay} Sekunden bevor der nächste Versuch startet...")
        await asyncio.sleep(retry_delay)

    # Wenn alle Versuche fehlgeschlagen sind, protokolliere und überspringe das Dokument
    print(f"PDF-Download für Dokument {doc.id} nach {max_retries} Versuchen fehlgeschlagen. Überspringe dieses Dokument.")

def sanitize_filename(filename):
    """
    Remove or replace characters in the filename that are not allowed in file names.
    """
    sanitized = re.sub(r'[<>:"/\\|?*]', '-', filename)  # Ersetze verbotene Zeichen durch '-'
    return sanitized[:255]  # Truncate to avoid overly long filenames

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

def export_json(paperless,doc, working_dir):
    """Export a document's metadata as JSON."""
    sanitized_title = sanitize_filename(doc.title)
    json_path = os.path.join(working_dir, f"{sanitized_title}.json")

    # Holen der Metadaten des Dokuments
    detailed_doc = get_document_json(paperless=paperless,doc=doc)
    # Metadata ist nun ein JSON-ähnliches Dictionary, das du weiter verarbeiten oder speichern kannst
    with open(json_path, "w", encoding="utf-8") as json_file:
       json.dump(detailed_doc, json_file, ensure_ascii=False, indent=4)


# ---------------------- Excel Export Helpers ----------------------
def export_to_excel(data, file_path, script_name, currency_columns, dir, url, meta):
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

    # Dateiname vorbereiten
    fullfilename = file_path
    filename_without_extension, file_extension = os.path.splitext(os.path.basename(file_path))

    # Falls Datei bereits geöffnet oder existiert, iterativ neuen Namen finden
    counter = 1
    while os.path.exists(fullfilename):
        filename = f"{filename_without_extension}-{counter}{file_extension}"
        fullfilename = os.path.join(directory, filename)
        counter += 1

    # Pandas DataFrame aus document_data erstellen
    df = pd.DataFrame(data)


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

        # Bereich, der formatiert werden soll
        range_string = f"A4:{worksheet.cell(row=worksheet.max_row, column=len(df.columns)).coordinate}"

        # Bedingte Formatierung für gerade Zeilen
        rule_even = FormulaRule(formula=[formula_even], fill=light_blue_fill, font=font)
        worksheet.conditional_formatting.add(range_string, rule_even)

        # Bedingte Formatierung für ungerade Zeilen
        rule_odd = FormulaRule(formula=[formula_odd], fill=white_fill, font=font)
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


    print(f"\nExcel-Datei erfolgreich erstellt: {fullfilename}")

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

async def exportThem(paperless, dir, query, progress_log_path):
    count = 0 
    """Process and export documents"""
    document_data = []
    currency_columns = []  # Liste zur Speicherung aller Currency-Felder
    custom_fields = {}
    meta = await fetch_paperless_meta(paperless, progress_log_path)

    documents = [item async for item in paperless.documents.search(query)]

    for doc in tqdm(documents, desc=f"Processing documents for '{dir}/{query}'", unit="doc"):
        count += 1

        # Retry-Mechanismus für das Abrufen der Metadaten
        metadata = None
        retries = 3  # Anzahl der Versuche
        for attempt in range(retries):
            try:
                # Abrufen der Metadaten des Dokuments
                metadata = await doc.get_metadata()
                break  # Wenn erfolgreich, breche die Schleife ab
            except aiohttp.client_exceptions.ServerDisconnectedError as e:
                print(f"Fehler beim Abrufen der Metadaten für Dokument {doc.id}: Serververbindung unterbrochen. Versuch {attempt + 1}/{retries}. {e}")
            except asyncio.TimeoutError as e:
                print(f"Timeout beim Abrufen der Metadaten für Dokument {doc.id}: Versuch {attempt + 1}/{retries}. {e}")
            except Exception as e:
                print(f"Unerwarteter Fehler beim Abrufen der Metadaten für Dokument {doc.id}: {e}. Versuch {attempt + 1}/{retries}")
            
            # Wenn der Abruf fehlschlägt, warte 5 Sekunden und versuche es erneut
            print(f"Warte 5 Sekunden bevor der nächste Versuch startet...")
            await asyncio.sleep(5)

        if not metadata:
            print(f"Metadaten für Dokument {doc.id} konnten nach {retries} Versuchen nicht abgerufen werden. Überspringe dieses Dokument.")
            continue  # Wenn Metadaten nicht abgerufen werden konnten, überspringe das Dokument

        docData = doc._data
        page_count = docData['page_count']
        custom_fields, doc_currency_columns = process_custom_fields(meta=meta,doc=docData)
        currency_columns.extend(doc_currency_columns)  # Speichere Currency-Felder

        tags = getmeta("tags", doc, meta)


        # Daten für die Excel-Tabelle sammeln
        row = OrderedDict(
            ("ID", doc.id),
            (  "AddDateFull", format_date(parse_date(doc.added), "yyyy-mm-dd")),
            ("Korrespondent", meta["correspondents"][doc.correspondent].name),
            ("Titel", doc.title),
            ("Tags", getmeta("tags", doc, meta)),

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
            ("Owner", getattr(meta["users"].get(doc.owner), "username", "Unbekannt") if doc.owner else "Unbekannt"),
            #("Notes", doc.notes),
        )

        document_data.append(row)

        # Exportiere das PDF des Dokuments
        await export_pdf(doc, working_dir=dir)
        export_json(paperless=paperless,doc=doc,working_dir=dir)


    # Exportiere die gesammelten Daten nach Excel
    path=doc._api_path 
    url=paperless._base_url

    last_dir = os.path.basename(dir)


    excel_file = os.path.join(dir, f"##{last_dir}-{datetime.now().strftime('%Y%m%d')}.xlsx")
    export_to_excel(document_data, excel_file, get_script_name, currency_columns=currency_columns,dir=dir, url=url,meta=meta )
    log_message(progress_log_path, f"dir: {dir}, Documents exported: {len(document_data)}")
    print(f"Exported Excel file: {excel_file}")


async def OLD_exportThem(paperless, dir, query,progress_log_path):
    count = 0 

    """Process and export documents """
    document_data = []
    currency_columns = []  # Liste zur Speicherung aller Currency-Felder
    custom_fields = {}
    meta = await fetch_paperless_meta(paperless, progress_log_path)

    documents = [item async for item in paperless.documents.search(query)]
   #     print(f"ID: {item.id}, Titel: {item.title}, Datum: {item.created}, correspondent: {item.correspondent} Storage_path: {item.storage_path}")
    for doc in tqdm(documents, desc=f"Processing documents for '{dir}/{query}'", unit="doc"):
        count+=1
        #custom_fields, doc_currency_columns = process_custom_fields(custom_fields_map, detailed_doc)
        #currency_columns.extend(doc_currency_columns)  # Speichere Currency-Felder

       # Dokumentdaten sammeln
        metadata = await doc.get_metadata()
        docData=doc._data
        page_count=docData['page_count']

        row = OrderedDict([
            ("ID", doc.id),
            ("AddDateFull", format_date(parse_date(doc.added), "yyyy-mm-dd")),
            ("Korrespondent", meta["correspondents"][doc.correspondent].name),
            #get_name_from_id(url, headers, "correspondents", doc.get("correspondent"))),
            ("Titel", doc.title),
            #("Tags", ", ".join(tag_dict.get(tag_id, f"Tag {tag_id}") for tag_id in doc.get("tags", []))),

            # Custom Fields direkt hinter den Tags einfügen
            #*custom_fields.items(),  

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
            ("Owner", meta["users"][doc.owner].username),

            ("Notes", doc.notes),
        ])

        document_data.append(row) 

        await export_pdf(doc, working_dir=dir)
       # export_json(paperless=paperless, doc=doc, working_dir=dir )
        document_data.append(row)

    excel_file = os.path.join(dir,f"##{tag_name}-{datetime.now().strftime('%Y%m%d')}.xlsx")
    export_to_excel(document_data, excel_file, script_name, tag_name, api_url=url, custom_fields_map=custom_fields_map, currency_columns=currency_columns)
    log_message(progress_log_path, f"Dir: {dir}, Documents exported: {len(document_data)}")
    print(f"Exported Excel file: {excel_file}")

async def main():
    script_name = get_script_name()
    config = load_config_from_script()

    export_dir = config.get("Export", "directory")
    api_url = config.get("API", "url")
    api_token = config.get("API", "token")
    log_dir = config.get("Log", "log_file")
    max_logs = config.get("Log", "max_logs")

    # Log-Dateien initialisieren
    progress_log_path, final_log_path = initialize_log(log_dir, script_name, max_logs)
    log_message(progress_log_path, message="Log in...")
    print_progress( message="Log in...")

# Setze das Arbeitsverzeichnis auf das Verzeichnis, in dem das Skript gespeichert ist
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    os.chdir(script_dir)

    locale.setlocale(locale.LC_ALL, '')  # Set locale based on system settings
    paperless = None
    try:
        paperless = Paperless(api_url, api_token)
        await paperless.initialize()
    except aiohttp.ClientResponseError as e:
        print(f"HTTP-Fehler beim Login: {e.status} {e.message}")
    except aiohttp.ClientConnectionError:
        print("Verbindungsfehler: Keine Verbindung zum Paperless-Server.")
    except aiohttp.ClientError as e:
        print(f"Allgemeiner aiohttp-Fehler: {e}")
    except Exception as e:
        print(f"Unerwarteter Fehler bei der Initialisierung von Paperless: {e}")
    finally:
        if paperless:
            print_progress("logged in....")
    # WENN `paperless` NICHT INITIALISIERT WURDE → ABBRUCH
    if not paperless:
        print("Paperless konnte nicht initialisiert werden. Beende das Skript.")
        exit(1)
    # do something

    meta = await fetch_paperless_meta(paperless, progress_log_path)
   # custom_fields_map = await get_custom_field_definitions(paperless)

    # Zugriff auf ein Element
    print(meta["correspondents"][3].name)
    print(meta["tags"][1].name)
    print(meta["storage_paths"][2].name)

    try:
        for root, dirs, files in os.walk(export_dir):
            query_value = os.path.basename(root)  # Standardwert: Verzeichnisname
            if root == export_dir:
                continue  # Gehe zum nächsten Durchgang, ohne das Root-Verzeichnis zu verarbeiten

            # Sortiere die Verzeichnisse alphabetisch
            dirs.sort() 
            if '##config.ini' in files:
                config_path = os.path.join(root, '##config.ini')
                config = configparser.ConfigParser()
                #config.read(config_path, encoding='utf-8')
                config.read(config_path)
            
                if 'DEFAULT' in config and 'query' in config['DEFAULT']:
                    query_value = config['DEFAULT']['query']

            print(f"{root} {query_value}")
            await exportThem(paperless=paperless, dir=root,query=query_value,progress_log_path=progress_log_path)
    except Exception as e:
        log_message(progress_log_path, f"Error: {str(e)}")
        raise
    finally:
        # Log umbenennen
        finalize_log(progress_log_path, final_log_path)
    await paperless.close()
    # do something
# see main() examples

asyncio.run(main())
