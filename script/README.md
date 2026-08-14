# 📄 paperless-ngx-2-excel
### High-performance Excel export engine for Paperless-ngx  
*(Async, intelligent caching, history, metadata, hyperlinks, custom fields, file linking)*

---

## 🌟 Overview

`paperless-ngx-2-excel` is a fast, reliable, automation-focused export engine for **Paperless-ngx**, designed for users managing thousands of documents who need:

- Rapid exports via async API  
- Fault-tolerant retry & backoff  
- Automatic, structured Excel reports  
- Intelligent caching for PDF/JSON  
- Clean, filterable Excel tables  
- Zero manual cleanup work  
- Fully automated daily/hourly exports  

The script produces:

- ✅ **Excel history files per day**  
- ✅ One **stable static Excel file** per export folder  
- ✅ A robust **`.all` cache** of PDFs and JSON metadata  
- ✅ **Hyperlinks** into your Paperless web UI  
- ✅ Auto-calculated **column widths** and table formatting  
- ✅ Alternating row shading & proper date formats  
- ✅ A comprehensive **metadata sheet** with environment info  

---

## 🚀 Key Features

### 📁 Export Engine

- One Excel file per export folder  
- History-versioned files named like:

  ```text
  ##2024-Steuer-20251113-0.xlsx
  ##2024-Steuer-20251113-1.xlsx
  ```

- One **static** Excel snapshot that always points to the latest export:

  ```text
  ##2024-Steuer.xlsx
  ```

- Per-folder query & schedule configured via `##config.ini`  
- Automatic cleanup of old `.xlsx` history files  

---

### 📊 Excel Output (via openpyxl only)

All Excel handling is implemented with **openpyxl** (no pandas involved):

- Clean Excel **Table** (`tbl<FolderName>`) for sorting/filtering
- Auto column width based on cell contents
- `freeze_panes = "A3"` for a fixed header
- Alternating row shading (zebra style)
- Automatic number formats:
  - Date columns → `YYYY-MM-DD`
  - Currency columns → `#,##0.00 €`
- Hyperlink column:
  - `LINK` column turned into `HYPERLINK("…/documents/<id>/details", "<id>")`

---

### 🧩 Metadata Sheet (📊 Metadaten)

The script creates a separate metadata sheet summarizing:

- Script version (Git tag / `git describe`)  
- Hostname, username, export timestamp  
- Export directory and Excel filename  
- File statistics:
  - Number of JSON files in the export folder
  - Number of PDF files in the export folder
- Config values:
  - Query used
  - Export frequency
- Document statistics:
  - Number of exported documents
  - List of all column headers
  - Custom field names and count
- Python environment:
  - Python version
  - Top installed packages (truncated list)

The sheet is formatted with:

- Bold, section-like headings (e.g. `📝 Exportinformationen`)  
- Grouped blocks separated by empty rows  
- Left-aligned key/value pairs  

---

### ♻️ File Caching (.all)

To avoid repeatedly downloading the same PDFs and JSONs, the script uses a shared **`.all` cache** under the export root:

```text
/exports/
  ├── .all/
  │     ├── 123--invoice-2023.pdf
  │     ├── 123--invoice-2023.json
  │     ├── 124--steuerbescheid.pdf
  │     └── ...
  ├── 2024-Steuer/
  └── 2025-NK/
```

- The `.all` cache is **rebuilt only if older than 1 hour** (configurable in code).  
- For each export folder, PDFs + JSONs are **linked** from `.all`:

  - Prefer **symlinks**
  - If symlinks fail, try **hardlinks**
  - If that fails, fallback to **copy**

- A small `##cache.timestamp` file is used to track cache age.  
- Orphaned entries (for deleted documents) are cleaned up from `.all`.

---

### 🔍 Query Handling per Export Folder

Each export folder under the configured export root has its own:

```text
##config.ini
```

Example:

```ini
[DATA]
query = path:*ST AND created:2024

[EXPORT]
frequency = hourly
```

This allows you to define:

- Paperless search query (`[DATA].query`)
- Export frequency (`[EXPORT].frequency`)

Supported frequencies:

- `hourly`
- `4hourly`
- `daily`
- `weekly`
- `monthly`
- `yearly`

The script decides whether to run an export for a folder based on:

- Timestamp of the latest `.xlsx` in that folder
- Timestamp of `##config.ini`
- Configured frequency

If nothing has changed and the frequency is not due, the folder is skipped.

---

## 🔧 Main Configuration (`paperless-ngx-2-excel.ini`)

The main script configuration is read from:

```text
paperless-ngx-2-excel.ini
```

This file must live **in the same directory as the script**.

### 📌 Example `paperless-ngx-2-excel.ini`

```ini
[API]
url = http://192.168.1.5:8000
token = 1234567890abcdef1234567890abcdef12345678

[Export]
directory = /volume1/paperless-export/exports

[Log]
log_file = /volume1/paperless-export/exports
max_files = 20
```

---

### 🔑 Getting Your Paperless API Token

From the Paperless-ngx Web UI:

1. Log in to your Paperless-ngx instance.  
2. Go to: **User Menu → Settings → API Tokens**.  
3. Click **“Generate new token”**.  
4. Copy the token value.  
5. Paste it into the `[API]` section:

   ```ini
   [API]
   url = http://your-paperless-server
   token = yourlongtokenhere
   ```

⚠️ **Security note:**  
Never commit your API token to Git, GitHub or any public repository.

---

### 🌐 Correct Paperless URL Format

Use the base URL of your Paperless instance **without `/api`**.

**✅ Correct examples:**

```text
http://192.168.1.5:8000
http://paperless.local
https://archive.example.com
```

**❌ Incorrect examples:**

```text
http://server/api
http://server/api/documents
```

The script will internally talk to `/api/...` based on the `url` you provide.

---

## 📁 Export Directory Structure

The `[Export].directory` defines the **export root**, for example:

```ini
[Export]
directory = /volume1/paperless-export/exports
```

A typical structure might look like:

```text
/volume1/paperless-export/exports/
  ├── .all/
  ├── 2024-Steuer/
  │     └── ##config.ini
  ├── 2025-NK/
  │     └── ##config.ini
  └── 2024-Rechnungen/
        └── ##config.ini
```

Each export subfolder:

- has its own `##config.ini`
- yields its own set of Excel files (`##<folder>-YYYYMMDD-N.xlsx` and `##<folder>.xlsx`)

### Example `##config.ini` in an export folder

```ini
# /volume1/paperless-export/exports/2024-Steuer/##config.ini

[DATA]
query = path:*ST AND created:2024

[EXPORT]
frequency = hourly
```

---

## 🧾 Logs & History Management

The `[Log]` section of `paperless-ngx-2-excel.ini` controls logging:

```ini
[Log]
log_file = /volume1/paperless-export/exports
max_files = 20
```

- `log_file` is a directory where log files are created.  
- `max_files` defines how many log files are retained; older ones are deleted.  

Log files are named like:

```text
##paperless-ngx-2-excel__2025-11-13_21-13-39.progress.log
##paperless-ngx-2-excel__2025-11-13_21-13-39.log
```

The script writes progress logs and then finalizes them at the end of a run.

---

## 🏃 Running the Script

### Manual Run

From the directory where the script is located:

```bash
./paperless-ngx-2-excel.py
```

Or explicitly with Python:

```bash
python3 paperless-ngx-2-excel.py
```

### Cronjob Example

To run the export every 10 minutes:

```cron
*/10 * * * * /usr/bin/python3 /path/to/paperless-ngx-2-excel.py
```

Make sure:

- The script is executable (`chmod +x paperless-ngx-2-excel.py`)  
- The config file `paperless-ngx-2-excel.ini` is in the same directory  

---

## 📊 Example Excel Output

### Sheet 1 — `Dokumentenliste`

Typical columns:

- `ID`  
- `LINK` (hyperlink to Paperless document details)  
- `Korrespondent`  
- `Titel`  
- `Tags`  
- Custom fields (based on your Paperless configuration)  
- `Seiten`  
- `Dokumenttyp`  
- `Speicherpfad`  
- `ArchivDate`, `ArchivedDateMonth`, `ArchivedDateFull`  
- `ModifyDate`, `ModifyDateMonth`, `ModifyDateFull`  
- `AddedDate`, `AddDateMonth`, `AddDateFull`  
- `OriginalName`, `ArchivedName`  
- `Owner`  

Features:

- Excel Table with name `tbl<FolderName>`  
- Filters enabled on each column header  
- Zebra striping for better readability  
- Summation row for currency fields (in row 2)  

### Sheet 2 — `📊 Metadaten`

Contains grouped sections like:

```text
📝 Exportinformationen
Script-Version       v1.0.0-27-g4d9cea2
Export-Datum         2025-11-13 21:18:41
Hostname             my-host
Username             my-user

📁 Verzeichnisse & Dateigrößen
Export-Verzeichnis   /volume1/paperless-export/exports/2024-Steuer
Excel-Datei          ##2024-Steuer-20251113-0.xlsx
Größe (xlsx)         832.50 KB
JSON-Dateien         446
PDF-Dateien          446

⚙️ Konfiguration (config.ini)
Query                path:*ST AND created:2024
Frequency            hourly

📊 Dokument-Statistik
Dokumente gesamt     446
Currency-Felder      Nettobetrag, Bruttobetrag
Header-Spalten       ID, LINK, Korrespondent, Titel, Tags, ...

🧩 Custom Fields
Anzahl Custom Fields 17
Custom Fields        Nettobetrag, Bruttobetrag, Kostenstelle, ...

🐍 Python Umgebung
Python-Version       3.14.0
Pakete (Top 10)      aiohttp==..., openpyxl==..., pypaperless==..., ...
```

---

## ⚙️ Internal Flow (How It Works)

1. **Configuration & Setup**
   - Load `paperless-ngx-2-excel.ini`
   - Validate all required config values
   - Initialize logging and locale

2. **Paperless Connection**
   - Create a `Paperless` client (from `pypaperless`)
   - Async initialization with retry/backoff

3. **Metadata Fetching**
   - Fetch and cache:
     - storage paths
     - correspondents
     - document types
     - tags
     - users
     - custom fields

4. **.all Cache**
   - Check `.all/##cache.timestamp`
   - If older than X seconds (default 3600), rebuild:
     - download PDFs
     - download JSON metadata
   - Clean up files belonging to deleted documents

5. **Export Directory Walk**
   - Walk `[Export].directory`
   - Skip special dirs (`.all`, `@eaDir`, etc.)
   - For each export folder:
     - Read `##config.ini`
     - Evaluate `frequency` and `should_export`
     - If due → run `exportThem(...)`

6. **Per-Folder Export**
   - Search Paperless documents with `query`
   - For each document:
     - Resolve metadata & custom fields
     - Link PDF + JSON via `link_export_file(...)`
     - Build a normalized row dict (all columns across all docs)
   - Call `export_to_excel(...)` once per folder

7. **Excel Creation**
   - Create new workbook
   - Add `Dokumentenliste` sheet
   - Write header row and all rows
   - Apply styles, table, number formats, hyperlinks
   - Create `📊 Metadaten` sheet
   - Save as:
     - new history file `##<folder>-YYYYMMDD-<N>.xlsx`
     - static file `##<folder>.xlsx` (copy)
   - Cleanup older history files based on `max_files`

---

## 🧱 Requirements

The script relies on:

```text
aiohttp==3.11.14
openpyxl==3.1.5
pypaperless==3.1.15
python_dateutil==2.9.0.post0
requests==2.32.3
tqdm==4.67.1
```

You can install them via:

```bash
pip install -r requirements.txt
```

(or install individually)

---

### ⚠️ Important Compatibility Notes

#### Python & aiohttp

- `aiohttp 3.11+` generally requires **Python ≥ 3.10**.  
- On very new Python versions (3.12, 3.13, 3.14), make sure your `aiohttp` build is compatible. If you see install/build errors, try:
  - upgrading `pip`
  - reinstalling `aiohttp` with `pip install --force-reinstall aiohttp`

#### pypaperless

- Use at least **`pypaperless >= 3.1.15`** for:
  - async support that matches newer Paperless versions
  - correct handling of custom fields (`data_type`, `extra_data`, etc.)
  - fewer surprises with API response changes

Older versions may:

- lack needed async iteration support
- not expose custom field metadata
- behave incorrectly against newer Paperless-ngx versions

#### Why no pandas?

Earlier versions used pandas to generate Excel files. This was removed because:

- Mixing pandas and openpyxl sometimes led to **corrupted `.xlsx` files**  
- Pandas struggles with **timezone-aware datetimes** in Excel exports  
- openpyxl alone gives us full control over:
  - Excel tables
  - formatting
  - styles
  - metadata sheets

The current implementation uses **only openpyxl** for Excel creation and is more robust.

---

## 🙏 Credits & References

This project stands on the shoulders of awesome open-source work:

### 📌 Paperless-ngx
**Paperless-ngx** is the underlying document management system.  
GitHub: https://github.com/paperless-ngx/paperless-ngx

### 📌 pypaperless
**pypaperless** is the async Python client used to talk to Paperless-ngx.  
PyPI: https://pypi.org/project/pypaperless/  
GitHub: https://github.com/timkpaine/pypaperless

### 📌 openpyxl
**openpyxl** is the engine used to generate `.xlsx` files.  
Docs: https://openpyxl.readthedocs.io/

Big thanks to all maintainers and contributors of these projects. 🙌

---

## 🤝 Contributing

Contributions are welcome:

- Bug reports
- Feature requests
- Pull requests
- Documentation improvements

If you have ideas like:

- Additional summary sheets  
- Better statistics (per tag/correspondent)  
- Direct integration with dashboards  
- Optional PDF preview sheets  

…feel free to open an issue or PR.

---

## 🛡️ License

This project is intended to be licensed under **GPL-3.0**, in the spirit of Paperless-ngx.

(Adjust the license file in the repository as needed.)

---

## ⭐ If this helps you

If you find this tool useful in your Paperless-ngx setup:

- Please **star the repository** on GitHub  
- Share screenshots or workflows  
- Tell others in the Paperless community 🙂

Enjoy your automated Paperless-ngx Excel exports! 📁 ➡️ 📊 🚀
