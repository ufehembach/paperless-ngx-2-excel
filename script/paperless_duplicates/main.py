"""Main async runner für Duplicate-Detection."""
import os
import sys
from configparser import ConfigParser
from pypaperless import Paperless

from .configutil import load_config_from_script, get_script_name
from .duplicate_detector import mark_duplicates_async


async def main():
    """Haupteinstiegspunkt."""
    script_name = get_script_name()
    config = load_config_from_script()

    # 🔍 Pflichtfelder prüfen
    required = {
        "API": ["url", "token"],
        "Duplicates": ["tag_name"]
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

    # ✅ Konfigurationswerte lesen
    api_url = config.get("API", "url")
    api_token = config.get("API", "token")
    tag_name = config.get("Duplicates", "tag_name")
    tag_color = config.get("Duplicates", "tag_color", fallback="#FF6B6B")

    print(f"\n🔍 Paperless Duplicate-Tagger")
    print(f"   API: {api_url}")
    print(f"   Tag: {tag_name}")
    print(f"   Farbe: {tag_color}\n")

    try:
        # Verbindung zu Paperless
        print("📡 Verbinde zu Paperless...")
        paperless = Paperless(api_url, api_token)
        await paperless.initialize()
        print("✓ Verbunden\n")

        # Duplikate taggen
        await mark_duplicates_async(
            paperless=paperless,
            tag_name=tag_name,
            tag_color=tag_color
        )

    except Exception as e:
        print(f"\n❌ Fehler: {e}")
        sys.exit(1)
    finally:
        if paperless:
            await paperless.close()

    print("\n✓ Fertig!")
