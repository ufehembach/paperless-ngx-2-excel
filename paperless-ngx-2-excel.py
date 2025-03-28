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

import asyncio
import aiohttp
from pypaperless import Paperless

def cleanup_old_logs(log_dir, script_name, max_logs_str):
    """ Löscht alte Logs, wenn die Anzahl der Log-Dateien das Limit überschreitet. """
    log_files = sorted(glob.glob(os.path.join(log_dir, f"{script_name}.*.log")), key=os.path.getmtime)

    max_logs=int(max_logs_str)
    while len(log_files) > max_logs:
        os.remove(log_files.pop(0))  # Älteste Datei löschen# Funktion, um den Log-Dateinamen basierend auf dem Skriptnamen und Datum zu erstellen

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

async def exportThem(paperless, dir, query):
    count = 0 
    lastitem = None
    print(f">>{query}<<")
    async for item in paperless.documents.search(query):
   #     print(f"ID: {item.id}, Titel: {item.title}, Datum: {item.created}, correspondent: {item.correspondent} Storage_path: {item.storage_path}")
        count+=1
        lastitem = item
    #print(lastitem)
    print (count)

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
    paperless = Paperless(api_url,api_token)

# Setze das Arbeitsverzeichnis auf das Verzeichnis, in dem das Skript gespeichert ist
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    os.chdir(script_dir)

    locale.setlocale(locale.LC_ALL, '')  # Set locale based on system settings

    await paperless.initialize()
    # do something

    # Nutzung für verschiedene Paperless-Objekte
    log_message(progress_log_path, message="getting storage_paths...")
    print_progress(message="getting storage_paths...")
    storage_paths = await get_dict_from_paperless(paperless.storage_paths)
    print(f"Storage Paths: {len(storage_paths)}")

    log_message(progress_log_path, message="getting corrospondents...")
    print_progress( message="getting corrospondents...")
    correspondents = await get_dict_from_paperless(paperless.correspondents)
    print(f"Correspondents: {len(correspondents)}")

    log_message(progress_log_path, message="getting doctypes...")
    print_progress( message="getting doctypes...")
    doctypes = await get_dict_from_paperless(paperless.document_types)
    print(f"Doctypes: {len(doctypes)}")

    log_message(progress_log_path, message="getting tags...")
    print_progress(message="getting tags...")
    tags = await get_dict_from_paperless(paperless.tags)
    print(f"Tags: {len(tags)}")
    # Zugriff auf ein Element
    print(doctypes[2].name)
    print(storage_paths[2].name)
    print(correspondents[3].name)


    try:
        for root, dirs, files in os.walk(export_dir):
            query_value = os.path.basename(root)  # Standardwert: Verzeichnisname
        
            if '##config.ini' in files:
                config_path = os.path.join(root, '##config.ini')
                config = configparser.ConfigParser()
                #config.read(config_path, encoding='utf-8')
                config.read(config_path)
            
                if 'DEFAULT' in config and 'query' in config['DEFAULT']:
                    query_value = config['DEFAULT']['query']

                if 1 == 0:   
                    documents = fetch_data(api_url, headers, "documents",query_value)
                    export_for_dir(
                        tags,
                        export_dir=export_dir,
                        documents=documents,
                        api_url=api_url,
                        headers=headers,
                        custom_fields_map=custom_fields_map,
                        tag_dict=tag_dict,
                        script_name=script_name,
                        log_path=progress_log_path,
                    )
            print(f"{root} {query_value}")
            await exportThem(paperless=paperless, dir=root,query=query_value)
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
