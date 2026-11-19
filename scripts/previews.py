"""Pages of the decks and documents, for the preview in the web view.

Each .pptx and .docx under a `drive/` folder gets a PDF under `content/previews/`, at the
same relative path. They are committed, so the preview works on a machine without
LibreOffice. Only whoever rebuilds needs it:

    .venv/bin/python scripts/build.py     (calls this at the end)
    .venv/bin/python scripts/previews.py  (previews only)

A PDF is rendered again only when its source file changed, so a rebuild with unchanged
input leaves the folder alone.
"""

from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import tempfile
from pathlib import Path

CONTENT = Path(__file__).resolve().parent.parent / "content"
OUT = CONTENT / "previews"
INDEX = OUT / "_index.json"
OFFICE = shutil.which("soffice") or next((c for c in ("/Applications/LibreOffice.app/Contents/MacOS/soffice", "/usr/bin/libreoffice") if Path(c).exists()), None)


def target(source: Path) -> Path:
    """content/<root>/drive/<rel>.pptx -> content/previews/<root>/<rel>.pptx.pdf"""
    root, rel = source.relative_to(CONTENT).as_posix().split("/drive/", 1)
    return OUT / root / (rel + ".pdf")


def build() -> None:
    sources = sorted(p for p in CONTENT.rglob("*") if p.suffix.lower() in (".pptx", ".docx") and "/drive/" in p.as_posix() and OUT not in p.parents)
    known = json.loads(INDEX.read_text(encoding="utf-8")) if INDEX.exists() else {}
    index, made, stale = {}, 0, 0
    work = Path(tempfile.mkdtemp(prefix="galaxium-previews-"))
    for source in sources:
        key, digest, pdf = source.relative_to(CONTENT).as_posix(), hashlib.sha1(source.read_bytes()).hexdigest(), target(source)
        if known.get(key) == digest and pdf.exists():
            index[key] = digest
            continue
        pdf.unlink(missing_ok=True)  # never show the pages of an older version
        if not OFFICE:
            stale += 1
            continue
        subprocess.run([OFFICE, "--headless", f"-env:UserInstallation={(work / 'profile').as_uri()}", "--convert-to", "pdf", "--outdir", str(work), str(source)],
                       capture_output=True, timeout=120, check=False)
        rendered = work / (source.stem + ".pdf")
        if rendered.exists():
            pdf.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(rendered, pdf)
            index[key], made = digest, made + 1
        else:
            stale += 1
    kept = {target(CONTENT / key) for key in index}
    for old in list(OUT.rglob("*.pdf")) if OUT.exists() else []:
        if old not in kept:
            old.unlink()
    shutil.rmtree(work, ignore_errors=True)
    OUT.mkdir(parents=True, exist_ok=True)
    INDEX.write_text(json.dumps(index, indent=1, sort_keys=True) + "\n", encoding="utf-8")
    print(f"previews: {len(index)} of {len(sources)} files have pages, {made} rendered now"
          + (f", {stale} without pages" + ("" if OFFICE else " (LibreOffice is not installed: install it and run scripts/previews.py)") if stale else ""))


if __name__ == "__main__":
    build()
