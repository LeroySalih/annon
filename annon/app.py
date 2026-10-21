from flask import Flask, render_template, request, send_file, abort
import io
import os
from urllib.parse import quote

from . import core

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024


def _ext(f):
    return os.path.splitext(f.filename or "")[1]


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/personalise")
def personalise():
    a, r = request.files.get("assignment"), request.files.get("roster")
    if not a or not r:
        abort(400, "Upload an assignment and a class list")
    try:
        title = os.path.splitext(a.filename or "")[0]
        z = core.personalise(a.read(), _ext(a), r.read(), float(request.form.get("margin", 12)),
                             title, float(request.form.get("footer", 20)))
    except ValueError as e:
        abort(400, str(e))
    return send_file(io.BytesIO(z), mimetype="application/zip", as_attachment=True, download_name="personalised.zip")


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
    app.run(host="127.0.0.1", port=5055)


if __name__ == "__main__":
    main()
