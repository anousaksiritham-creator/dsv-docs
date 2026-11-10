---
name: dsv-doc
description: Build a Dansavanh (DSV / ແດນສະຫວັນ) letterhead document as one HTML file (Lao, English or Chinese switchable in the page; PDF and Word exported from the page) — report, letter/proposal, or two-party agreement. Use when the user asks for a DSV report, ໜັງສືສະເໜີ, ຂໍ້ຕົກລົງ, "ເອກະສານແດນສະຫວັນ", or any document on the Dansavanh letterhead.
---

Produce one self-contained HTML file on the Dansavanh letterhead from JSON data with `scripts/build_html.py`. The user exports PDF or Word from the page. The letterhead, colours, fonts, header, footer and block renderers live in `assets/viewer.html`; this file covers the workflow and the JSON contract only.

## Workflow

1. **Intake.** Settle these before writing anything, asking only for what the user has not already said:
   - document type: `report` (sections, tables, annexes), `letter` (ref no., to, subject, body, one signer) or `agreement` (parties, clauses, two-party signature block)
   - project folder under `D:\ANS\DSV\<project>\` (output goes there) and a short project code for the meta line (e.g. `FERRY-2026`)
   - languages: default `lo`; produce `en` / `zh` only when asked
   - status `draft` or `final`, and date
   - content source: text the user gives in chat, or files the user names (JSON from cons-analysis / CD-BOQ / rc-structure, Excel, Markdown, a web page). Read them yourself.
2. **Outline.** Post the outline in chat: title, info rows, every section heading with a one-line summary of its blocks, signature parties, annexes. Wait for the user's confirmation. Skip the wait when the user says "ສ້າງເລີຍ" (build now) or already approved an outline.
3. **Write the JSON** to `<project folder>/<slug>.json`, starting from the closest file in `templates/` (`report.json`, `letter.json`, `agreement.json`; `example_report_ferry.json` is a complete real report). Facts and figures come from the sources; a value you do not have stays an explicit blank (`"…"`) that you list for the user, never an invented number.
4. **Build and verify** as one HTML file (default output, all languages inside):
   ```
   python "~/.claude/skills/dsv-doc/scripts/build_html.py" <data.json> [more.json ...] --out <Name>.html --lang lo
   ```
   Several JSON files go into the same HTML, shown as a document picker. `--check` validates without building. Open the result in the browser, switch language and status, and check that no `[image not found]` callout appears and tables fit the page. Fix the JSON and rebuild until clean.
   **One source, one output.** The JSON is the only thing the user edits; the HTML is rebuilt from it. Do not create `.docx` or `.pdf` files as separate outputs, and do not keep a second copy of a document. The user exports PDF or Word from the page when needed (**Export PDF** via the print dialog, **Export Word** as a `.doc` file to save as `.docx`).
5. **Report** the output path, the languages inside, and the list of blanks the user still has to fill. Approval comes from the user on the page; after approval, set `status` to `final` in the JSON and rebuild.

## JSON contract

Every string may be a plain string or `{"lo": "...", "en": "...", "zh": "..."}`; the builder picks the `--lang` value (fallback lo → en), so one JSON produces all language files.

```
meta: type, title, subtitle, code, kind, status (draft|final), date (ISO "2026-11-07" → formatted per language),
      info [[label, value], ...]   (report/agreement: 2 pairs per row under the title)
      footer                        (override the default footer text)
      org_line1, org_line2          (letterhead right-hand text; defaults = resort tagline EN + ZH)
      mono: true                    (monochrome logo for black-and-white print)
      letter only: ref_no, to, subject, attention, signer {closing, name, position}
      agreement only: parties [{title, name, position, date}]  → signature block after the clauses
sections: [{title, numbered (default true), annex (true → lettered ກ/ຂ/ຄ or A/B/C after a page break), blocks}]
blocks:
  p {text}                                  body paragraph
  h {level 3, text}                         sub-heading (write the "2.1" prefix yourself)
  bullets {items} / numbered {items}        item = string, or [label, text] to print your own label ("1.3", "ກ)")
  table {columns, rows, widths cm, align [l|c|r], total_row, caption}
  kv {rows [[k, v]]}                        two-column key/value table
  callout {text, tone info|warn}            note box; formula {text} and verdict {text} are the same family
  image {path, width_cm, caption}           path relative to the JSON file
  signatures {parties [{title, name, position, date}]}   side-by-side signing blocks
  pagebreak
```

Numbering: top-level sections are numbered 1., 2., … automatically (set `"numbered": false` for the executive summary or an approval section). Section 1 of a `letter` is unnumbered and untitled by convention.

## Conventions

- Lao text uses Phetsarath OT, English Arial, Chinese Microsoft YaHei; headings and table headers are Dansavanh green. Changing look-and-feel means editing `assets/viewer.html`, never per-document JSON.
- Write dates in the JSON as ISO; the builder prints "7 ພະຈິກ 2026" / "7 November 2026" / "2026年11月7日".
- Units and symbols in prose: words over symbols ("ປະມານ", "ເທົ່າກັບ"), "x" for multiplication, no emoji or arrows in document text.
- Keep the JSON next to the output in the project folder so the document can be rebuilt after edits.
