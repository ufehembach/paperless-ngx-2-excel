# 📄 paperless-ngx-2-excel
Advanced Excel export automation for Paperless-NGX — with metadata cache, directory‑based exports, custom fields, hyperlinks, and clean XLSX formatting.

## 🚀 Intention
**paperless-ngx-2-excel** is designed to automate Excel exports from a Paperless‑NGX installation.  
It creates clean, styled, filterable Excel spreadsheets — fully offline, fast, and suitable for accounting, archives, taxes, or reporting workflows.

The tool handles:
- All metadata provided by Paperless-NGX  
- Custom fields (monetary, select, multiselect, …)  
- Automatic PDF/JSON linking  
- Folder‑based export logic  
- Smart caching  
- Pretty Excel formatting  

## 🙌 Credits & References
This project stands on the shoulders of great open-source work:

- **[Paperless-NGX](https://github.com/paperless-ngx/paperless-ngx)**  
  Document management platform

- **[pypaperless](https://github.com/danielperna84/pypaperless)**  
  Python SDK used to communicate with Paperless-NGX API

- **[openpyxl](https://openpyxl.readthedocs.io/)**  
  Excel writer used to create clean, styled XLSX files  

Special thanks to all contributors of these projects.

## ⚙️ How It Works
### 1. Directory-based export
Each subdirectory under your export root becomes one Excel export target.  
Every folder typically contains a:

```
##config.ini
```

which defines:
- `query`: Paperless search query  
- `frequency`: when to export (daily, hourly, monthly, …)

### 2. Metadata cache (`.all` directory)
To avoid re-downloading thousands of PDFs/JSONs, the script maintains a smart cache:
- Stores `{docid}--title.pdf`  
- Stores `{docid}--title.json`  
- Reuses cached files via **symlink → hardlink → copy** fallback

The cache rebuilds only if older than 1 hour (configurable).

### 3. Excel builder
For each directory:
- Fetch all documents matching the query  
- Resolve metadata (correspondent, tags, storage path, custom fields, …)
- Build an XLSX with:
  - Title row
  - Full document table
  - Formulas for currency columns
  - Hyperlinks to document detail views
  - Alternating row color
  - Proper Excel Table object (`tbl<dirname>`)
  - Freeze panes
  - Auto column width

### 4. Metadata sheet
Each Excel file includes a second sheet:
- Script version  
- Hostname, username  
- File sizes (JSON/PDF/export)  
- Config query & frequency  
- Custom field statistics  
- Python packages used  

### 5. History handling
Excel files are rotated as:

```
##Steuer-2025-03-04-0.xlsx
##Steuer-2025-03-04-1.xlsx
...
##Steuer.xlsx           ← static always-updated file
```

## 📦 What It Does (Summary)
- ✔ Exports all metadata into a consistent Excel table  
- ✔ Creates JSON/PDF links (symlink-friendly)  
- ✔ Adds automatic numbering & date stamps  
- ✔ Adds a styled, professional XLSX table  
- ✔ Creates a static file without timestamp  
- ✔ Stores detailed metadata into separate sheet  
- ✔ Performs cleanup of outdated files  
- ✔ Fully async → **fast**  
- ✔ Works on macOS, Synology NAS, Linux  
- ✔ Zero manual maintenance

## 📁 Example Directory Layout

```
exports/
 ├── 2024-Steuer/
 │    ├── ##config.ini
 │    ├── ##2024-Steuer-20250304-0.xlsx
 │    ├── ##2024-Steuer.xlsx   ← static always updated
 │    └── PDFs & JSONs
 ├── 2025-Nebenkosten/
 │    └── ##config.ini
 └── .all/
      ├── 884--Mietervertrag.pdf
      ├── 884--Mietervertrag.json
      └── ##cache.timestamp
```

## 📄 Sample `##config.ini`

```ini
[DATA]
query = path:"Steuer" AND created:2024

[EXPORT]
frequency = hourly
```

## 📊 Example Excel Output

**Sheet 1 – Dokumentenliste**  
- Clean header  
- Filters enabled  
- Alternating stripes  
- Automatic column widths  
- Live hyperlinks for each document  

**Sheet 2 – 📊 Metadaten**  
- Export information  
- Directory statistics  
- Custom field overview  
- Python package list  

## 🔧 Requirements

```
aiohttp==3.11.14
openpyxl==3.1.5
pypaperless==3.1.15
python_dateutil==2.9.0.post0
requests==2.32.3
tqdm==4.67.1
```

## 🏗 Installation

```bash
git clone https://github.com/ufe-dev/paperless-ngx-2-excel
cd paperless-ngx-2-excel
pip install -r requirements.txt
```

## ▶️ Usage

```bash
./paperless-ngx-2-excel.py
```

Exports all configured folders inside the `Export.directory` path from your INI file.

## 🪪 License
This project uses the SPDX identifier detected from your GitHub repository.  
See GitHub for details.

---

If you find this tool useful, feel free to leave a ⭐ on GitHub!

