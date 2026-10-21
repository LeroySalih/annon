from flask import Flask, render_template, request, send_file, abort
import io
import os
import secrets
from urllib.parse import quote

from . import core

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024


def _ext(f):
    return os.path.splitext(f.filename or "")[1]


@app.get("/")
def index():
    return render_template("index.html")


_mappings: dict[str, str] = {}  # token -> pupil ID/name CSV, held in memory until fetched once


@app.post("/personalise")
def personalise():
    a, r = request.files.get("assignment"), request.files.get("roster")
    if not a or not r:
        abort(400, "Upload an assignment and a class list")
    try:
        pdf, mapping = core.personalise(a.read(), _ext(a), r.read(), float(request.form.get("margin", 12)),
                                        os.path.splitext(a.filename or "")[0], float(request.form.get("footer", 20)))
    except ValueError as e:
        abort(400, str(e))
    token = secrets.token_urlsafe(16)
    _mappings[token] = mapping
    resp = send_file(io.BytesIO(pdf), mimetype="application/pdf", as_attachment=True, download_name="personalised.pdf")
    resp.headers["X-Mapping-Token"] = token
    return resp


@app.get("/mapping/<token>")
def mapping(token):
    csv_text = _mappings.pop(token, None)
    if csv_text is None:
        abort(404)
    return send_file(io.BytesIO(csv_text.encode()), mimetype="text/csv", as_attachment=True, download_name="pupil_ids.csv")


@app.post("/redact")
def redact():
    files = [f for f in request.files.getlist("scans") if f.filename]
    if not files:
        abort(400, "Upload at least one scan")
    band = float(request.form.get("band", 25))
    pdfs, errors = [], []
    for f in files:
        try:
            pdfs.append(core.redact_scan(f.read(), _ext(f), band))
        except Exception:  # unreadable file: report it, keep going with the rest
            errors.append(os.path.basename(f.filename))
    if not pdfs:
        abort(400, "None of the files could be read: " + ", ".join(errors))
    resp = send_file(io.BytesIO(core.merge_pdfs(pdfs)), mimetype="application/pdf",
                     as_attachment=True, download_name="anonymised.pdf")
    if errors:
        resp.headers["X-Skipped"] = quote(", ".join(errors))
    return resp


def main():
    host = os.environ.get("ANNON_HOST", "127.0.0.1")
    port = int(os.environ.get("ANNON_PORT", "5055"))
    url = f"http://127.0.0.1:{port}"
    if os.environ.get("ANNON_OPEN_BROWSER"):
        import threading
        import webbrowser
        threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    print(f"Assignment Anonymiser running at {url} (close this window to stop)", flush=True)
    try:
        from waitress import serve
        serve(app, host=host, port=port, threads=4)
    except ImportError:
        app.run(host=host, port=port)


if __name__ == "__main__":
    main()
