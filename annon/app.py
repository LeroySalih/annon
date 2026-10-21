from flask import Flask, render_template, request, send_file, abort
import io
import os
import zipfile

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
    results, errors, seen = [], [], set()
    for f in files:
        stem = os.path.splitext(os.path.basename(f.filename))[0]
        try:
            pdf = core.redact_scan(f.read(), _ext(f), band)
        except Exception:  # unreadable file: report it, keep going with the rest
            errors.append(f.filename)
            continue
        name, n = f"{stem}_anon.pdf", 1
        while name in seen:
            n += 1
            name = f"{stem}_anon_{n}.pdf"
        seen.add(name)
        results.append((name, pdf))
    if not results:
        abort(400, "None of the files could be read: " + ", ".join(errors))
    if len(results) == 1 and not errors:
        return send_file(io.BytesIO(results[0][1]), mimetype="application/pdf",
                         as_attachment=True, download_name=results[0][0])
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, pdf in results:
            z.writestr(name, pdf)
        if errors:
            z.writestr("errors.txt", "Could not process:\n" + "\n".join(errors) + "\n")
    buf.seek(0)
    return send_file(buf, mimetype="application/zip", as_attachment=True, download_name="anonymised.zip")


def main():
    app.run(host="127.0.0.1", port=5055)


if __name__ == "__main__":
    main()
