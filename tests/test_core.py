import io

import docx
import pymupdf

from annon import core

ROSTER = b"name\nAda Lovelace\nAlan Turing\nGrace Hopper\n"


def _pdf():
    d = pymupdf.open()
    for _ in range(2):
        p = d.new_page()
        p.insert_text((72, 300), "Question 1")
        p.insert_text((200, p.rect.height - 25), "OLD FOOTER Page 1 of 6")
    return d.tobytes()


def _docx():
    d = docx.Document()
    d.add_paragraph("Question 1")
    b = io.BytesIO()
    d.save(b)
    return b.getvalue()


def test_roster_variants():
    assert core.read_roster(b"First,Last\nA,B\n") == ["A B"]
    assert core.read_roster(b"x\nA\n") == ["A"]


def test_codes_unique_5_digit():
    pairs = core.assign_codes([f"p{i}" for i in range(2000)])
    codes = [c for _, c in pairs]
    assert len(set(codes)) == 2000 and all(len(c) == 5 for c in codes)


def _ids(mapping):
    return {n: c for c, n in (l.split(",") for l in mapping.strip().splitlines()[1:])}


def test_personalise_pdf_single_file():
    pdf, mapping = core.personalise(_pdf(), "pdf", ROSTER, title="Systems")
    ids = _ids(mapping)
    doc = pymupdf.open(stream=pdf, filetype="pdf")
    assert len(doc) == 6  # 3 pupils x 2 pages
    for i, name in enumerate(["Ada Lovelace", "Alan Turing", "Grace Hopper"]):
        code = ids[name]
        for j in range(2):
            page = doc[i * 2 + j]
            h = page.rect.height
            top = page.get_text(clip=pymupdf.Rect(0, 0, page.rect.width, 50))
            bot = page.get_text(clip=pymupdf.Rect(0, h - 50, page.rect.width, h))
            assert name in top and f"Pupil ID: {code}" in top
            assert "OLD FOOTER" not in bot and name.split()[0] not in bot
            assert f"Pupil ID: {code}" in bot and "Systems" in bot and f"Page {j + 1} of 2" in bot
            w = {x[4]: x[0] for x in page.get_text("words", clip=pymupdf.Rect(0, h - 50, page.rect.width, h))}
            assert w["Pupil"] < w["Systems"] < w["Page"]


def test_personalise_docx_structure():
    out = core.personalise_docx(_docx(), "Alan Turing", "12345", "Systems")
    s = docx.Document(io.BytesIO(out)).sections[0]
    assert "Alan Turing" in s.header.paragraphs[-1].text and "12345" in s.header.paragraphs[-1].text
    ft = s.footer.paragraphs[-1]
    assert len(s.footer.paragraphs) == 1 and ft.text.startswith("Pupil ID: 12345")
    assert "Systems" in ft.text and "PAGE" in ft._p.xml and "NUMPAGES" in ft._p.xml


def test_personalise_docx_single_pdf():
    pdf, mapping = core.personalise(_docx(), "docx", ROSTER, title="Systems")
    doc = pymupdf.open(stream=pdf, filetype="pdf")
    ids = _ids(mapping)
    assert len(doc) == 3
    for page, name in zip(doc, ["Ada Lovelace", "Alan Turing", "Grace Hopper"]):
        text = page.get_text()
        assert name in text and ids[name] in text and "Page 1 of 1" in text


def test_redact_blocks_header_keeps_footer():
    src = core.personalise_pdf(_pdf(), "Ada Lovelace", "12345")
    out = pymupdf.open(stream=core.redact_scan(src, "pdf", band_mm=25), filetype="pdf")
    assert len(out) == 2
    pix = out[0].get_pixmap(dpi=72)
    assert pix.pixel(10, 10)[:3] == (0, 0, 0)
    assert out[0].get_text() == ""  # flattened: no recoverable text


def test_app_endpoints():
    from annon.app import app
    c = app.test_client()
    r = c.post("/personalise", data={"assignment": (io.BytesIO(_pdf()), "a.pdf"), "roster": (io.BytesIO(ROSTER), "r.csv")})
    assert r.status_code == 200 and r.mimetype == "application/pdf"
    assert len(pymupdf.open(stream=r.data, filetype="pdf")) == 6
    m = c.get("/mapping/" + r.headers["X-Mapping-Token"])
    assert m.status_code == 200 and m.data.decode().startswith("pupil_id,name")
    assert c.get("/mapping/" + r.headers["X-Mapping-Token"]).status_code == 404  # one-time
    r = c.post("/redact", data={"scans": (io.BytesIO(_pdf()), "s.pdf")})
    assert r.status_code == 200 and r.mimetype == "application/pdf"
    r = c.post("/redact", data={"scans": [(io.BytesIO(_pdf()), "a.pdf"), (io.BytesIO(_pdf()), "a.pdf"), (io.BytesIO(b"junk"), "bad.pdf")]},
               content_type="multipart/form-data")
    assert r.status_code == 200 and r.mimetype == "application/pdf"
    assert len(pymupdf.open(stream=r.data, filetype="pdf")) == 4  # 2 files x 2 pages
    assert r.headers["X-Skipped"] == "bad.pdf"
    assert c.post("/redact").status_code == 400
