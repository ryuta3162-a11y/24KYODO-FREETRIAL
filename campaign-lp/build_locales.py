#!/usr/bin/env python3
"""locales/{page}/{lang}.json → locales.bundle.js"""
import json
from pathlib import Path

ROOT = Path(__file__).parent
LOCALES = ROOT / "locales"
OUT = ROOT / "locales.bundle.js"

bundle = {}
for page_dir in sorted(LOCALES.iterdir()):
    if not page_dir.is_dir():
        continue
    page = page_dir.name
    bundle[page] = {}
    for f in sorted(page_dir.glob("*.json")):
        lang = f.stem
        bundle[page][lang] = json.loads(f.read_text(encoding="utf-8"))

OUT.write_text(
    "window.JOYFIT_LOCALES = " + json.dumps(bundle, ensure_ascii=False, indent=2) + ";\n",
    encoding="utf-8",
)
print(f"Wrote {OUT} ({len(bundle)} pages)")
