#!/usr/bin/env python3
"""
Entry point für paperless-ngx-2-excel.

Die eigentliche Logik liegt in paperless_export/ (aufgeteilt in thematische
Module), inkl. sp_tree.py für den SP-Baum mit JSON-Index (Move-Tracking bei
Storage-Path-Änderungen) und current/sub-Excel-Mappen pro Ordner.
Dieser Dateiname bleibt unverändert, damit bestehende Docker-Builds,
Cronjobs und Aufrufe ("python paperless-ngx-2-excel.py") weiter funktionieren.
"""
import asyncio
from paperless_export.main import main

if __name__ == "__main__":
    asyncio.run(main())
