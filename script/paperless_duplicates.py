#!/usr/bin/env python3
"""
Entry point für paperless_duplicates.

Markiert alle Dokumente mit Duplikaten mit einem konfigurierbaren Tag
für einfacheres Filtern in der Paperless-UI.
"""
import asyncio
from paperless_duplicates.main import main

if __name__ == "__main__":
    asyncio.run(main())
