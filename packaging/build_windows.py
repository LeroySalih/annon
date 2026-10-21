"""Build a portable Windows folder (zip): embeddable Python + dependencies + app.

Run on any OS with internet:  python3 packaging/build_windows.py
Output: dist/annon-windows.zip. No installer or admin rights needed to use it.
"""
import pathlib
import shutil
import subprocess
import sys
import urllib.request
import zipfile

PY = "3.12.10"
EMBED = f"https://www.python.org/ftp/python/{PY}/python-{PY}-embed-amd64.zip"
ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "dist" / "annon-windows"

START_BAT = r"""@echo off
cd /d "%~dp0"
set ANNON_OPEN_BROWSER=1
echo Starting Assignment Anonymiser...
"python\python.exe" -m annon.app
echo.
echo The app has stopped. Scroll up for any error message.
pause
"""

README = """Assignment Anonymiser (portable)
================================
1. Unzip this folder anywhere you can write (e.g. Documents or a USB stick).
   Do not run it from inside the zip.
2. Double-click start.bat. A window opens and your browser shows the app.
   If Windows SmartScreen warns, choose More info > Run anyway.
3. Close the window to stop it.

No installer and no admin password needed. It only listens on this PC.

DOCX assignments need Microsoft Word on this PC (or LibreOffice). PDFs need nothing.
Keep pupil_ids.csv private: it links pupil IDs to names.
"""


def main():
    if OUT.exists():
        shutil.rmtree(OUT)
    (OUT / "app").mkdir(parents=True)
    zpath = ROOT / "dist" / f"python-{PY}-embed-amd64.zip"
    if not zpath.exists():
        urllib.request.urlretrieve(EMBED, zpath)
    with zipfile.ZipFile(zpath) as z:
        z.extractall(OUT / "python")
    pth = next((OUT / "python").glob("python*._pth"))
    zipname = next((OUT / "python").glob("python*.zip")).name
    pth.write_text(f"{zipname}\n.\n..\\site-packages\n..\\app\nimport site\n")
    subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", "--target", str(OUT / "site-packages"),
                    "--only-binary=:all:", "--platform", "win_amd64", "--python-version", "3.12",
                    "--implementation", "cp", "flask", "pymupdf>=1.24,<2", "python-docx", "waitress"], check=True)
    shutil.copytree(ROOT / "annon", OUT / "app" / "annon",
                    ignore=shutil.ignore_patterns("__pycache__"))
    (OUT / "start.bat").write_text(START_BAT.replace("\n", "\r\n"))
    (OUT / "README.txt").write_text(README.replace("\n", "\r\n"))
    zip_out = ROOT / "dist" / "annon-windows.zip"
    shutil.make_archive(str(zip_out.with_suffix("")), "zip", OUT.parent, OUT.name)
    print(f"Built {zip_out} ({zip_out.stat().st_size / 1e6:.0f} MB)")


if __name__ == "__main__":
    main()
