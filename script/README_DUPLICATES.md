# Paperless Duplicates Tagger

Markiert alle Dokumente mit Duplikaten mit einem konfigurierbaren Tag in Paperless.

## Setup

1. **API-Token holen:**
   - Paperless UI → Admin → Authentication Tokens
   - Neuen Token erstellen und kopieren

2. **Config bearbeiten:**
   ```bash
   vim paperless_duplicates.ini
   ```
   - `API.url` — Deine Paperless-URL (z.B. `http://localhost:8000`)
   - `API.token` — Der soeben kopierte Token
   - `Duplicates.tag_name` — Name des Tags (default: `hat_duplikate`)
   - `Duplicates.tag_color` — Hex-Farbe (default: `#FF6B6B` = Rot)

## Ausführung

```bash
python3 paperless_duplicates.py
```

Das Script wird dann:
1. Mit Paperless verbinden
2. Alle Dokumente laden
3. Duplikate identifizieren (via `duplicate_count` aus der Paperless-API)
4. Mit dem konfigurierten Tag versehen

## Nach der Ausführung

In der Paperless UI kannst du dann filtern:
- `tag:hat_duplikate` → zeigt alle Dokumente mit Duplikaten

## Anforderungen

```
pypaperless
requests
aiohttp
```

Installation (falls nötig):
```bash
pip3 install pypaperless requests aiohttp
```
