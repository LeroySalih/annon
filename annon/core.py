"""Personalise PDF/DOCX assignments per pupil and redact headers on scans.

Pupil names stay in memory/output files only; nothing here logs them.
"""
import csv
import io
import re
import secrets
import zipfile

import docx
import pymupdf

MM = 72 / 25.4  # PDF points per millimetre


def read_roster(data: bytes) -> list[str]:
    """Return pupil names from CSV bytes.

    Uses a 'name' column, else 'first'+'last' columns, else the first column.
    """
    text = data.decode("utf-8-sig")
    rows = list(csv.reader(io.StringIO(text)))
    rows = [r for r in rows if any(c.strip() for c in r)]
    if not rows:
        raise ValueError("Class list is empty")
    head = [c.strip().lower() for c in rows[0]]
    body = rows[1:]
    if "name" in head:
        i = head.index("name")
        names = [r[i] for r in body if len(r) > i]
    else:
        first = next((i for i, h in enumerate(head) if h in ("first", "first name", "firstname", "forename")), None)
        last = next((i for i, h in enumerate(head) if h in ("last", "last name", "lastname", "surname")), None)
        if first is not None and last is not None:
            names = [f"{r[first]} {r[last]}" for r in body if len(r) > max(first, last)]
        else:
            names = [r[0] for r in body]
    names = [re.sub(r"\s+", " ", n).strip() for n in names]
    names = [n for n in names if n]
    if not names:
        raise ValueError("No pupil names found")
    return names


def assign_codes(names: list[str]) -> list[tuple[str, str]]:
    """Give each pupil a unique random 5-digit code (10000-99999)."""
    if len(names) > 90000:
        raise ValueError("Too many pupils for 5-digit codes")
    used: set[str] = set()
    out = []
    for n in names:
        while True:
            c = str(10000 + secrets.randbelow(90000))
            if c not in used:
                used.add(c)
                break
        out.append((n, c))
    return out


def _safe(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_") or "pupil"


def personalise_pdf(src: bytes, name: str, code: str, margin_mm: float = 12) -> bytes:
    """Stamp name+code in the top margin and code in the bottom margin of every page."""
    doc = pymupdf.open(stream=src, filetype="pdf")
    y = margin_mm * MM / 2 + 3
    for page in doc:
        r = page.rect
        pad = 36
        page.insert_text((pad, y), name, fontsize=11, fontname="helv")
        label = f"Code: {code}"
        w = pymupdf.get_text_length(label, fontname="helv", fontsize=11)
        page.insert_text((r.width - pad - w, y), label, fontsize=11, fontname="helv")
        fw = pymupdf.get_text_length(label, fontname="helv", fontsize=10)
        page.insert_text(((r.width - fw) / 2, r.height - y + 6), label, fontsize=10, fontname="helv")
    out = doc.tobytes(garbage=3, deflate=True)
    doc.close()
    return out


def personalise_docx(src: bytes, name: str, code: str) -> bytes:
    d = docx.Document(io.BytesIO(src))
    for s in d.sections:
        heads = [s.header]
        foots = [s.footer]
        if s.different_first_page_header_footer:
            heads.append(s.first_page_header)
            foots.append(s.first_page_footer)
        for h in heads:
            h.is_linked_to_previous = False
            p = h.add_paragraph(f"{name}\tCode: {code}")
            p.paragraph_format.tab_stops.add_tab_stop(s.page_width - s.left_margin - s.right_margin, 2)  # right
        for f in foots:
            f.is_linked_to_previous = False
            f.add_paragraph(f"Code: {code}").alignment = 1  # centre
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def personalise(src: bytes, ext: str, roster: bytes, margin_mm: float = 12) -> bytes:
    """Return a zip: one file per pupil plus codes.csv (code -> name)."""
    ext = ext.lower().lstrip(".")
    if ext not in ("pdf", "docx"):
        raise ValueError("Assignment must be a .pdf or .docx")
    pairs = assign_codes(read_roster(roster))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        key = io.StringIO()
        w = csv.writer(key)
        w.writerow(["code", "name"])
        for name, code in pairs:
            data = personalise_pdf(src, name, code, margin_mm) if ext == "pdf" else personalise_docx(src, name, code)
            z.writestr(f"{_safe(name)}_{code}.{ext}", data)
            w.writerow([code, name])
        z.writestr("codes.csv", key.getvalue())
    return buf.getvalue()


def redact_scan(src: bytes, ext: str, band_mm: float = 25, dpi: int = 200) -> bytes:
    """Black out the top band_mm of every page, flattened to images so nothing is recoverable."""
    ext = ext.lower().lstrip(".")
    if ext == "pdf":
        doc = pymupdf.open(stream=src, filetype="pdf")
    elif ext in ("png", "jpg", "jpeg", "tif", "tiff"):
        img = pymupdf.open(stream=src, filetype=ext)
        doc = pymupdf.open("pdf", img.convert_to_pdf())
    else:
        raise ValueError("Scan must be a PDF or image")
    out = pymupdf.open()
    for page in doc:
        pix = page.get_pixmap(dpi=dpi)
        rows = min(pix.height, round(band_mm / 25.4 * dpi))
        pix = pymupdf.Pixmap(pix, 0) if pix.alpha else pix
        pix.set_rect(pymupdf.IRect(0, 0, pix.width, rows), (0, 0, 0))
        np_ = out.new_page(width=page.rect.width, height=page.rect.height)
        np_.insert_image(np_.rect, pixmap=pix)
    data = out.tobytes(garbage=3, deflate=True)
    out.close()
    doc.close()
    return data
