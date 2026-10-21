# annon

Personalise assignments per pupil, then anonymise the scanned-in work.

1. **Personalise**: upload a PDF or DOCX and a class-list CSV. Each pupil gets a copy with their name and a unique random 5-digit pupil ID in the header, and the same pupil ID in the footer. You also get `codes.csv` (pupil ID → name); keep it private.
2. **Anonymise**: upload a scan (PDF or image). The top band of every page (default 25 mm) is blacked out and the page is flattened to an image, so the header can't be recovered. The footer pupil ID stays so you can match papers back.

PDFs need a free margin strip top and bottom (default 12 mm) because the text is overlaid. DOCX files get the text added to the real header and footer.

## Run
    python3 -m venv .venv && . .venv/bin/activate
    pip install -r requirements.txt
    python -m annon.app      # http://127.0.0.1:5055
    pytest

Runs locally only; pupil names are never logged or sent anywhere.
