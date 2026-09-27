"""Config-Utilities für paperless_duplicates."""
import os
import sys
from configparser import ConfigParser


def get_script_name():
    """Ermittelt den Namen des ausgeführten Scripts (ohne .py)."""
    return os.path.splitext(os.path.basename(sys.argv[0]))[0]


def load_config_from_script():
    """
    Lädt die .ini-Datei basierend auf dem Script-Namen.
    Erwartet: script_name.ini im gleichen Verzeichnis.
    """
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    script_name = get_script_name()
    config_path = os.path.join(script_dir, f"{script_name}.ini")

    if not os.path.exists(config_path):
        print(f"❌ Konfigurationsdatei nicht gefunden: {config_path}")
        sys.exit(1)

    config = ConfigParser()
    config.read(config_path)

    return config
