from flask import Flask, render_template, request, send_file, abort
import io
import os

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
        z = core.personalise(a.read(), _ext(a), r.read(), float(request.form.get("margin", 12)))
    except ValueError as e:
        abort(400, str(e))
    return send_file(io.BytesIO(z), mimetype="application/zip", as_attachment=True, download_name="personalised.zip")


@app.post("/redact")
def redact():
    s = request.files.get("scan")
    if not s:
        abort(400, "Upload a scan")
    try:
        pdf = core.redact_scan(s.read(), _ext(s), float(request.form.get("band", 25)))
    except ValueError as e:
        abort(400, str(e))
    return send_file(io.BytesIO(pdf), mimetype="application/pdf", as_attachment=True, download_name="anonymised.pdf")


def main():
    app.run(host="127.0.0.1", port=5055)


if __name__ == "__main__":
    main()
