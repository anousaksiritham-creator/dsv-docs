#!/usr/bin/env python3
"""build_html.py — Build ONE self-contained Dansavanh (DSV) letterhead HTML file from JSON.

Usage:
  python build_html.py <data.json> [more.json ...] [--out file.html] [--lang lo|en|zh] [--check]

The HTML holds every document passed in and every language in the JSON. The reader picks the
document, language and status in the page toolbar and exports that one to PDF (print) or Word.
Logos and images are embedded, so the single file works offline and can be sent as it is.

JSON shape: see ../SKILL.md and ../templates/*.json (same contract as build_docx.py).
The page itself (letterhead, styles, renderer, export) is ../assets/viewer.html.
"""
import argparse
import base64
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(os.path.dirname(HERE), "assets")
TEMPLATE = os.path.join(ASSETS, "viewer.html")
LOGOS = {"horizontal": "logo_horizontal.png", "mono": "logo_mono.png", "symbol": "logo_symbol.png"}
MIME = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif",
        ".webp": "image/webp", ".bmp": "image/bmp"}
BLOCK_TYPES = {"p", "h", "bullets", "numbered", "table", "kv", "callout", "formula", "verdict",
               "image", "signatures", "pagebreak"}


def data_uri(path):
    ext = os.path.splitext(path)[1].lower()
    if ext not in MIME:
        raise ValueError(f"unsupported image type '{ext}' (use png, jpg, gif, webp or bmp)")
    with open(path, "rb") as fh:
        return f"data:{MIME[ext]};base64," + base64.b64encode(fh.read()).decode("ascii")


def check(data, base_dir):
    """Validate one document and embed its images (image block gets a 'src' data URI)."""
    issues = []
    m = data.get("meta", {})
    if m.get("type", "report") not in ("report", "letter", "agreement"):
        issues.append(f"meta.type unknown: {m.get('type')}")
    if not m.get("title") and m.get("type") != "letter":
        issues.append("meta.title missing")
    raw = json.dumps(data, ensure_ascii=False)
    for pat in ("TODO", "TBD", "xxx", "____"):
        if pat in raw:
            issues.append(f"placeholder '{pat}' found ({raw.count(pat)}x)")
    for s in data.get("sections", []):
        for b in s.get("blocks", []):
            t = b.get("type")
            if t not in BLOCK_TYPES:
                issues.append(f"unknown block type '{t}' in '{s.get('title')}'")
            elif t == "table":
                n = len(b.get("columns", [])) or None
                for r in b.get("rows", []):
                    if n and len(r) != n:
                        issues.append(f"table in '{s.get('title')}': row has {len(r)} cells, header has {n}")
                        break
                if b.get("widths") and n and len(b["widths"]) != n:
                    issues.append(f"table in '{s.get('title')}': {len(b['widths'])} widths for {n} columns")
            elif t == "image":
                path = b.get("path", "")
                full = path if os.path.isabs(path) else os.path.join(base_dir, path)
                try:
                    b["src"] = data_uri(full)
                except (OSError, ValueError) as e:
                    issues.append(f"image '{path}' in '{s.get('title')}': {e}")
    return issues


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")  # Lao text in messages on a cp1252 console
    ap = argparse.ArgumentParser(description="Build one Dansavanh letterhead HTML file from JSON")
    ap.add_argument("data", nargs="+", help="one or more document JSON files, in the order shown in the page")
    ap.add_argument("--out")
    ap.add_argument("--lang", default="lo", choices=["lo", "en", "zh"], help="language shown when the page opens")
    ap.add_argument("--check", action="store_true", help="validate JSON only")
    a = ap.parse_args()

    docs, total = [], 0
    for path in a.data:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
        issues = check(data, os.path.dirname(os.path.abspath(path)))
        for i in issues:
            print(f"! {os.path.basename(path)}: {i}")
        total += len(issues)
        docs.append({"name": os.path.splitext(os.path.basename(path))[0], "data": data})
    if a.check:
        sys.exit(1 if total else 0)

    logos = {}
    for key, name in LOGOS.items():
        p = os.path.join(ASSETS, name)
        logos[key] = data_uri(p) if os.path.exists(p) else ""
        if not logos[key]:
            print(f"! logo missing: {p}")
    payload = json.dumps({"docs": docs, "logos": logos, "lang": a.lang}, ensure_ascii=False).replace("</", "<\\/")
    with open(TEMPLATE, encoding="utf-8") as fh:
        html = fh.read()
    if html.count("__DSV_DATA__") != 1:
        sys.exit("viewer.html must contain the __DSV_DATA__ placeholder exactly once")
    out = a.out or os.path.splitext(a.data[0])[0] + ".html"
    with open(out, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(html.replace("__DSV_DATA__", payload))
    print(f"saved: {out}  ({len(docs)} document(s), {os.path.getsize(out) // 1024} KB" +
          (f", {total} warnings)" if total else ")"))


if __name__ == "__main__":
    main()
