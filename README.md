# annon

Personalise assignments per pupil, then anonymise the scanned-in work.

1. **Personalise**: upload a PDF or DOCX and a class-list CSV. Each pupil gets a copy with their name and a unique random 5-digit pupil ID in the header, and the same pupil ID in the footer. The output is one PDF (a copy per pupil, in class-list order) plus `pupil_ids.csv` (pupil ID → name); keep the CSV private. DOCX input is converted to PDF with LibreOffice, which must be installed.
2. **Anonymise**: upload a scan (PDF or image). The top band of every page (default 25 mm) is blacked out and the page is flattened to an image, so the header can't be recovered. The footer pupil ID stays so you can match papers back.

PDFs need a free margin strip top and bottom (default 12 mm) because the text is overlaid. DOCX files get the text added to the real header and footer.

## Run
    python3 -m venv .venv && . .venv/bin/activate
    pip install -r requirements.txt
    python -m annon.app      # http://127.0.0.1:5055
    pytest

Runs locally only; pupil names are never logged or sent anywhere.

## Run as a container
    docker compose up -d --build

The container joins the existing external `dino-app` network (create it first via the dino-app stack, or `docker network create dino-app`). No host port is published; point your reverse proxy at `http://annon:5055`. The app has **no login**, so only expose it through something that authenticates staff.

## Portable Windows package (no installer, no admin)
Build once on any machine with internet (Mac/Linux/Windows):

    python3 packaging/build_windows.py

This writes `dist/annon-windows.zip` (~40 MB: embeddable Python, dependencies and the app). Copy the zip to a Windows PC, unzip it anywhere writable, and double-click `start.bat`. DOCX input uses Microsoft Word (via PowerShell) or LibreOffice if present; PDFs need nothing.
