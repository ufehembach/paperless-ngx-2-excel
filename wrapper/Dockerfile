FROM python:3.13-slim

ENV TZ=Europe/Berlin \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    EXPORT_DIR=/exports \
    FORCE_COPY=1

WORKDIR /app

# Projektdateien ins Image kopieren
COPY script/ /app/

# Paket-Manager aktuell halten & Problemversionen vermeiden
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir "aiohttp!=3.13.2,<4"

# Deine Requirements (z. B. openpyxl, pandas, etc.)
RUN pip install --no-cache-dir -r /app/requirements.txt

# Sichere Installation der gewünschten pypaperless-Version
RUN pip install --no-cache-dir "pypaperless==5.2.2"

# Für Debug-Zwecke anzeigen
RUN pip show pypaperless && pip show aiohttp

# Das eigentliche Startkommando
CMD ["python", "/app/paperless-ngx-2-excel.py"]
