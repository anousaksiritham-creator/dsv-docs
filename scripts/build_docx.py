#!/usr/bin/env python3
"""build_docx.py — Build a Dansavanh (DSV) letterhead Word document from JSON.

Usage:
  python build_docx.py <data.json> --lang lo|en|zh [--out file.docx] [--status draft|final] [--pdf] [--check]

JSON shape (see ../templates/*.json):
  meta:   type (report|letter|agreement), title, subtitle, project, code, kind, status (draft|final),
          date (ISO "2026-11-07" or text), footer, info [[label, value], ...], org_line1, org_line2,
          letter: ref_no, to, subject, attention, closing, signer {name, position}
  sections: [{key?, title, annex?: true, blocks: [...]}]
  blocks: p | h(level 2/3) | bullets | numbered | table(columns, rows, widths cm, align l/c/r, total_row, caption)
          | kv(rows [[k, v]]) | callout(text, tone info/warn) | formula(text) | image(path, width_cm, caption)
          | signatures(parties [{title, name, position}]) | pagebreak

Any string value may instead be {"lo": "...", "en": "...", "zh": "..."}; the value for --lang is used
(falls back to lo, then en). One JSON can therefore produce all three language files.
"""
import argparse
import datetime as dt
import json
import os
import re
import sys

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:
    sys.exit("pip install python-docx")

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(os.path.dirname(HERE), "assets")

# ---------------------------------------------------------------- brand
GREEN_HEX, INK_HEX, MUTED_HEX, RULE_HEX, SOFT_HEX = "006A3C", "1B1F1D", "5A625E", "C9D1CC", "F3F6F4"
GREEN, INK, MUTED = RGBColor(0x00, 0x6A, 0x3C), RGBColor(0x1B, 0x1F, 0x1D), RGBColor(0x5A, 0x62, 0x5E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CONTENT_W = 16.5  # cm, A4 minus 2.5 + 2.0 margins

FONTS = {
    "lo": {"ascii": "Phetsarath OT", "east": "Phetsarath OT", "cs": "Phetsarath OT", "size": 11, "lang": "lo-LA", "ls": 1.25},
    "en": {"ascii": "Arial", "east": "Arial", "cs": "Arial", "size": 10.5, "lang": "en-GB", "ls": 1.15},
    "zh": {"ascii": "Microsoft YaHei", "east": "Microsoft YaHei", "cs": "Microsoft YaHei", "size": 10.5, "lang": "zh-CN", "ls": 1.2},
}

LABELS = {
    "lo": {"page": "ໜ້າ", "of": "/", "note": "ໝາຍເຫດ", "warn": "ຂໍ້ຄວນລະວັງ", "annex": "ພາກຜະໜວກ",
           "name": "ຊື່", "position": "ຕຳແໜ່ງ", "signature": "ລາຍເຊັນ", "date": "ວັນທີ",
           "ref": "ເລກທີ", "to": "ຮຽນ", "subject": "ເລື່ອງ", "attention": "ອ້າງເຖິງ",
           "footer_draft": "ສະບັບຮ່າງ — ບໍ່ແມ່ນຄຳປຶກສາດ້ານກົດໝາຍ ຫຼື ການເງິນ.",
           "footer_final": "ເອກະສານພາຍໃນ Dansavanh",
           "draft": "ສະບັບຮ່າງ", "final": "ສະບັບສົມບູນ", "approval": "ການອະນຸມັດ",
           "annex_letters": "ກຂຄງຈສຊຍດຕ", "ref_suffix": "/ດສວ/"},
    "en": {"page": "Page", "of": "of", "note": "Note", "warn": "Caution", "annex": "Annexes",
           "name": "Name", "position": "Position", "signature": "Signature", "date": "Date",
           "ref": "Ref.", "to": "To", "subject": "Subject", "attention": "Reference",
           "footer_draft": "Draft — not legal or financial advice.",
           "footer_final": "Dansavanh internal document",
           "draft": "DRAFT", "final": "FINAL", "approval": "Approval",
           "annex_letters": "ABCDEFGHIJ", "ref_suffix": "/DSV/"},
    "zh": {"page": "第", "of": "页，共", "note": "说明", "warn": "注意", "annex": "附录",
           "name": "姓名", "position": "职务", "signature": "签名", "date": "日期",
           "ref": "编号", "to": "致", "subject": "事由", "attention": "参照",
           "footer_draft": "报告草稿——不构成法律或财务建议。",
           "footer_final": "丹沙湾内部文件",
           "draft": "草稿", "final": "定稿", "approval": "批准",
           "annex_letters": "ABCDEFGHIJ", "ref_suffix": "/DSV/"},
}
LAO_MONTHS = ["ມັງກອນ", "ກຸມພາ", "ມີນາ", "ເມສາ", "ພຶດສະພາ", "ມິຖຸນາ", "ກໍລະກົດ", "ສິງຫາ", "ກັນຍາ", "ຕຸລາ", "ພະຈິກ", "ທັນວາ"]
EN_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
ORG = {"line1": "Cultural & Eco-Tourism Resort", "line2": "丹沙湾文化生态旅游度假区"}


# ---------------------------------------------------------------- language resolution
def resolve(obj, lang):
    """Replace {"lo":..,"en":..,"zh":..} dicts by the string for lang (recursively)."""
    if isinstance(obj, dict):
        keys = set(obj.keys())
        if keys and keys <= {"lo", "en", "zh"}:
            return obj.get(lang) or obj.get("lo") or obj.get("en") or obj.get("zh") or ""
        return {k: resolve(v, lang) for k, v in obj.items()}
    if isinstance(obj, list):
        return [resolve(v, lang) for v in obj]
    return obj


def fmt_date(value, lang, short=False):
    """ISO date -> language date. Non-ISO text is returned unchanged."""
    if not value:
        return ""
    try:
        d = dt.date.fromisoformat(str(value)[:10])
    except ValueError:
        return str(value)
    if short:
        return d.strftime("%d.%m.%Y")
    if lang == "lo":
        return f"{d.day} {LAO_MONTHS[d.month - 1]} {d.year}"
    if lang == "zh":
        return f"{d.year}年{d.month}月{d.day}日"
    return f"{d.day} {EN_MONTHS[d.month - 1]} {d.year}"


# ---------------------------------------------------------------- low-level helpers
def set_run_font(run, lang, size=None, bold=None, color=None, italic=None, mono=False):
    fx = FONTS[lang]
    face = "Consolas" if mono else fx["ascii"]
    run.font.name = face
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for a in list(rfonts.attrib):
        del rfonts.attrib[a]
    rfonts.set(qn("w:ascii"), face)
    rfonts.set(qn("w:hAnsi"), face)
    rfonts.set(qn("w:eastAsia"), fx["east"])
    rfonts.set(qn("w:cs"), fx["cs"])
    sz = size or fx["size"]
    run.font.size = Pt(sz)
    szcs = rpr.find(qn("w:szCs"))
    if szcs is None:
        szcs = OxmlElement("w:szCs")
        rpr.append(szcs)
    szcs.set(qn("w:val"), str(int(sz * 2)))
    if bold is not None:
        run.font.bold = bold
        if bold and rpr.find(qn("w:bCs")) is None:
            rpr.append(OxmlElement("w:bCs"))
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color


def style_fonts(style, lang, size, bold=False, color=None):
    fx = FONTS[lang]
    style.font.name = fx["ascii"]
    style.font.size = Pt(size)
    style.font.bold = bold
    if color is not None:
        style.font.color.rgb = color
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for a in list(rfonts.attrib):
        del rfonts.attrib[a]
    rfonts.set(qn("w:ascii"), fx["ascii"])
    rfonts.set(qn("w:hAnsi"), fx["ascii"])
    rfonts.set(qn("w:eastAsia"), fx["east"])
    rfonts.set(qn("w:cs"), fx["cs"])
    el = rpr.find(qn("w:szCs"))
    if el is None:
        el = OxmlElement("w:szCs")
        rpr.append(el)
    el.set(qn("w:val"), str(int(size * 2)))
    lang_el = rpr.find(qn("w:lang"))
    if lang_el is None:
        lang_el = OxmlElement("w:lang")
        rpr.append(lang_el)
    lang_el.set(qn("w:val"), "en-US" if lang == "lo" else fx["lang"])
    lang_el.set(qn("w:eastAsia"), "zh-CN" if lang == "zh" else "en-US")
    lang_el.set(qn("w:bidi"), "lo-LA" if lang == "lo" else "th-TH")


def shade(cell, hex_color):
    tcpr = cell._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcpr.append(shd)


def cell_margins(table, top=60, bottom=60, left=110, right=110):
    mar = OxmlElement("w:tblCellMar")
    for side, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(v))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    table._element.tblPr.append(mar)


def table_borders(table, color=RULE_HEX, size=4, edges=("top", "left", "bottom", "right", "insideH", "insideV")):
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        if edge in edges:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(size))
            el.set(qn("w:color"), color)
        else:
            el.set(qn("w:val"), "nil")
        borders.append(el)
    table._element.tblPr.append(borders)


def cell_borders(cell, **sides):
    """sides: top/bottom/left/right = (size, hex) or None for nil."""
    tcpr = cell._element.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        spec = sides.get(side, "keep")
        if spec == "keep":
            continue
        if spec is None:
            el.set(qn("w:val"), "nil")
        else:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(spec[0]))
            el.set(qn("w:color"), spec[1])
        b.append(el)
    tcpr.append(b)


def repeat_header(row):
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    row._tr.get_or_add_trPr().append(el)


def no_split(row):
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def add_field(paragraph, instr, lang, size=None, color=None):
    run = paragraph.add_run()
    set_run_font(run, lang, size=size, color=color)
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t"); t.text = "1"
    f3 = OxmlElement("w:fldChar"); f3.set(qn("w:fldCharType"), "end")
    for el in (f1, it, f2, t, f3):
        run._element.append(el)


def para_border(paragraph, side="bottom", color=GREEN_HEX, size=12, space=4):
    ppr = paragraph._p.get_or_add_pPr()
    bdr = ppr.find(qn("w:pBdr"))
    if bdr is None:
        bdr = OxmlElement("w:pBdr")
        ppr.append(bdr)
    b = OxmlElement(f"w:{side}")
    b.set(qn("w:val"), "single"); b.set(qn("w:sz"), str(size)); b.set(qn("w:space"), str(space)); b.set(qn("w:color"), color)
    bdr.append(b)


def set_widths(t, widths):
    tblpr = t._element.tblPr
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = OxmlElement("w:tblW")
        tblpr.append(tblw)
    tblw.set(qn("w:w"), str(int(sum(widths) * 567)))
    tblw.set(qn("w:type"), "dxa")
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblpr.append(layout)
    for j, gc in enumerate(t._element.tblGrid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths[j] * 567)))
    for row in t.rows:
        for j, c in enumerate(row.cells):
            c.width = Cm(widths[j])


def is_number(text):
    s = str(text).strip()
    return bool(re.fullmatch(r"[\s\-+]?[\d.,\s]+%?[\w/²³]*", s)) and any(ch.isdigit() for ch in s)


# ---------------------------------------------------------------- builder
class DsvDoc:
    def __init__(self, data, lang, status=None):
        self.d = data
        self.m = data.get("meta", {})
        self.lang = lang
        self.L = LABELS[lang]
        self.fx = FONTS[lang]
        self.base = self.fx["size"]
        self.status = (status or self.m.get("status") or "draft").lower()
        self.doc = Document()
        self.sec_no = 0
        self.annex_no = 0
        self._setup()

    # ---- page / styles
    def _setup(self):
        doc, lang = self.doc, self.lang
        s = doc.sections[0]
        s.page_height, s.page_width = Cm(29.7), Cm(21.0)
        s.top_margin, s.bottom_margin = Cm(2.5), Cm(2.0)
        s.left_margin, s.right_margin = Cm(2.5), Cm(2.0)
        s.header_distance, s.footer_distance = Cm(0.8), Cm(0.8)
        st = doc.styles
        style_fonts(st["Normal"], lang, self.base)
        pf = st["Normal"].paragraph_format
        pf.space_after = Pt(6)
        pf.line_spacing = self.fx["ls"]
        for name, size in (("Heading 1", self.base + 7), ("Heading 2", self.base + 3), ("Heading 3", self.base + 1)):
            style_fonts(st[name], lang, size, bold=True, color=GREEN if name != "Heading 3" else INK)
            hp = st[name].paragraph_format
            hp.space_before = Pt(16 if name == "Heading 2" else 10)
            hp.space_after = Pt(6)
            hp.keep_with_next = True
            hp.line_spacing = 1.2
        settings = doc.settings.element
        uf = OxmlElement("w:updateFields")
        uf.set(qn("w:val"), "true")
        settings.append(uf)

    # ---- primitives
    def run(self, p, text, **kw):
        r = p.add_run(str(text))
        set_run_font(r, self.lang, **kw)
        return r

    def para(self, text="", align=None, size=None, bold=None, color=None, space_after=None, italic=None, space_before=None):
        p = self.doc.add_paragraph()
        if text:
            self.run(p, text, size=size, bold=bold, color=color, italic=italic)
        if align is not None:
            p.alignment = align
        if space_after is not None:
            p.paragraph_format.space_after = Pt(space_after)
        if space_before is not None:
            p.paragraph_format.space_before = Pt(space_before)
        return p

    def page_break(self):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.add_run().add_break(WD_BREAK.PAGE)

    def body_para(self, text):
        p = self.para(text)
        if self.lang in ("en", "zh"):
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        return p

    # ---- letterhead (header) and footer
    def _header_footer(self):
        s = self.doc.sections[0]
        hdr = s.header
        hp = hdr.paragraphs[0]
        hp.paragraph_format.space_after = Pt(0)
        t = hdr.add_table(rows=1, cols=2, width=Cm(CONTENT_W))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        set_widths(t, [8.5, 8.0])
        c0, c1 = t.rows[0].cells
        for c in (c0, c1):
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell_borders(c, top=None, left=None, right=None, bottom=(24, GREEN_HEX))
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(4)
        logo = os.path.join(ASSETS, "logo_mono.png" if self.m.get("mono") else "logo_horizontal.png")
        if os.path.exists(logo):
            p0.add_run().add_picture(logo, height=Cm(1.0))
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.paragraph_format.space_after = Pt(0)
        p1.paragraph_format.line_spacing = 1.1
        self.run(p1, self.m.get("org_line1", ORG["line1"]), size=9, bold=True, color=GREEN)
        p2 = c1.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p2.paragraph_format.space_after = Pt(4)
        p2.paragraph_format.line_spacing = 1.1
        self.run(p2, self.m.get("org_line2", ORG["line2"]), size=8, color=MUTED)
        # move the empty first paragraph below the table so the table sits at the top
        hp._p.getparent().remove(hp._p)
        tail = hdr.add_paragraph()
        tail.paragraph_format.space_after = Pt(0)
        self.run(tail, " ", size=4)

        ftr = s.footer
        fp = ftr.paragraphs[0]
        fp.paragraph_format.space_after = Pt(0)
        # the built-in Footer style carries centre/right tab stops; clear them so one right stop remains
        style_tabs = fp.style.paragraph_format.tab_stops
        for ts in list(style_tabs):
            style_tabs.clear_all()
            break
        fp.paragraph_format.tab_stops.add_tab_stop(Cm(CONTENT_W), WD_TAB_ALIGNMENT.RIGHT)
        para_border(fp, "top", RULE_HEX, 4, 6)
        sym = os.path.join(ASSETS, "logo_symbol.png")
        if os.path.exists(sym):
            fp.add_run().add_picture(sym, height=Cm(0.6))
            self.run(fp, "   ", size=8)
        foot_text = self.m.get("footer") or (self.L["footer_draft"] if self.status == "draft" else
                                               f'{self.m.get("title", "")} · {self.L["footer_final"]}')
        self.run(fp, foot_text, size=8, color=MUTED)
        self.run(fp, "\t", size=8)
        if self.lang == "zh":
            self.run(fp, "第 ", size=8, color=MUTED)
            add_field(fp, "PAGE", self.lang, 8, MUTED)
            self.run(fp, " 页，共 ", size=8, color=MUTED)
            add_field(fp, "NUMPAGES", self.lang, 8, MUTED)
            self.run(fp, " 页", size=8, color=MUTED)
        else:
            self.run(fp, f'{self.L["page"]} ', size=8, color=MUTED)
            add_field(fp, "PAGE", self.lang, 8, MUTED)
            self.run(fp, f' {self.L["of"]} ', size=8, color=MUTED)
            add_field(fp, "NUMPAGES", self.lang, 8, MUTED)

    # ---- meta line: "DSV / CODE / KIND        DRAFT · 07.11.2026"
    def meta_line(self):
        m = self.m
        parts = ["DSV"] + [str(x) for x in (m.get("code"), m.get("kind")) if x]
        p = self.para(space_after=10)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(CONTENT_W), WD_TAB_ALIGNMENT.RIGHT)
        self.run(p, " / ".join(parts).upper(), size=7.5, color=MUTED, mono=True)
        status = self.L["draft"] if self.status == "draft" else self.L["final"]
        right = status + (" · " + fmt_date(m.get("date"), self.lang, short=True) if m.get("date") else "")
        self.run(p, "\t" + right, size=7.5, color=MUTED, mono=True)

    # ---- title block
    def title_block(self):
        m = self.m
        if m.get("title"):
            p = self.para(m["title"], size=self.base + 8, bold=True, color=INK, space_after=4)
            p.paragraph_format.line_spacing = 1.2
            p.paragraph_format.keep_with_next = True
        if m.get("subtitle"):
            self.para(m["subtitle"], size=self.base + 0.5, color=MUTED, space_after=14)

    def info_table(self, rows):
        """rows: [[label, value], ...] laid out 2 pairs per row like the HTML .info table."""
        if not rows:
            return
        pairs = [rows[i:i + 2] for i in range(0, len(rows), 2)]
        t = self.doc.add_table(rows=len(pairs), cols=4)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        table_borders(t)
        cell_margins(t, 70, 70, 120, 120)
        set_widths(t, [3.3, 4.95, 3.3, 4.95])
        small = self.base - 1
        for i, pair in enumerate(pairs):
            tr = t.rows[i]
            no_split(tr)
            for k in range(2):
                ck, cv = tr.cells[2 * k], tr.cells[2 * k + 1]
                ck.vertical_alignment = cv.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                shade(ck, SOFT_HEX)
                ck.paragraphs[0].paragraph_format.space_after = Pt(0)
                cv.paragraphs[0].paragraph_format.space_after = Pt(0)
                if k < len(pair):
                    label, value = pair[k][0], pair[k][1]
                    if label in ("date", "ວັນທີ", "Date", "日期"):
                        value = fmt_date(value, self.lang)
                    self.run(ck.paragraphs[0], label, bold=True, size=small)
                    self.run(cv.paragraphs[0], value, size=small)
        self.para(space_after=4)

    # ---- blocks
    def heading(self, text, level):
        h = self.doc.add_heading(level=level)
        size = self.base + (7 if level == 1 else 3 if level == 2 else 1)
        self.run(h, text, size=size, bold=True, color=GREEN if level < 3 else INK)
        if level == 2:
            para_border(h, "bottom", RULE_HEX, 4, 3)
        return h

    def bullets(self, items, numbered=False):
        """items: strings, or [label, text] pairs whose label replaces the automatic number (e.g. "1.3", "ກ)")."""
        for k, it in enumerate(items, 1):
            label, text = (it[0], it[1]) if isinstance(it, (list, tuple)) else ((f"{k})" if numbered else "•"), it)
            indent = 1.4 if isinstance(it, (list, tuple)) else 0.9
            p = self.doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_after = Pt(3)
            pf.left_indent = Cm(indent)
            pf.first_line_indent = Cm(-indent + 0.2)
            pf.tab_stops.add_tab_stop(Cm(indent))
            self.run(p, str(label) + "\t", color=MUTED if numbered else INK, bold=isinstance(it, (list, tuple)))
            self.run(p, text)

    def table(self, b):
        cols = b.get("columns", [])
        rows = b.get("rows", [])
        n = len(cols) or (len(rows[0]) if rows else 1)
        widths = b.get("widths") or [CONTENT_W / n] * n
        total = sum(widths)
        if total > CONTENT_W:
            widths = [w * CONTENT_W / total for w in widths]
        align = b.get("align") or []
        amap = {"l": WD_ALIGN_PARAGRAPH.LEFT, "c": WD_ALIGN_PARAGRAPH.CENTER, "r": WD_ALIGN_PARAGRAPH.RIGHT}
        small = self.base - 1
        t = self.doc.add_table(rows=(1 if cols else 0) + len(rows), cols=n)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        table_borders(t)
        cell_margins(t)
        set_widths(t, widths)
        r0 = 0
        if cols:
            hdr = t.rows[0]
            repeat_header(hdr)
            no_split(hdr)
            for j, name in enumerate(cols):
                c = hdr.cells[j]
                shade(c, GREEN_HEX)
                cell_borders(c, top=(4, GREEN_HEX), left=(4, GREEN_HEX), right=(4, GREEN_HEX), bottom=(4, GREEN_HEX))
                c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                p = c.paragraphs[0]
                p.alignment = amap.get(align[j], WD_ALIGN_PARAGRAPH.LEFT) if j < len(align) else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_after = Pt(0)
                self.run(p, name, bold=True, color=WHITE, size=small)
            r0 = 1
        for i, row in enumerate(rows):
            tr = t.rows[r0 + i]
            no_split(tr)
            last = b.get("total_row") and i == len(rows) - 1
            for j in range(n):
                val = row[j] if j < len(row) else ""
                c = tr.cells[j]
                c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
                if last:
                    shade(c, SOFT_HEX)
                p = c.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.15
                a = align[j] if j < len(align) else ("r" if is_number(val) and j > 0 else "l")
                p.alignment = amap.get(a, WD_ALIGN_PARAGRAPH.LEFT)
                self.run(p, val, size=small, bold=bool(last))
        if b.get("caption"):
            self.para(b["caption"], size=8.5, color=MUTED, italic=True)
        else:
            self.para(space_after=2)

    def kv(self, rows):
        t = self.doc.add_table(rows=len(rows), cols=2)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        table_borders(t)
        cell_margins(t, 70, 70, 120, 120)
        set_widths(t, [4.6, CONTENT_W - 4.6])
        for i, (k, v) in enumerate(rows):
            c0, c1 = t.rows[i].cells
            no_split(t.rows[i])
            shade(c0, SOFT_HEX)
            for c in (c0, c1):
                c.paragraphs[0].paragraph_format.space_after = Pt(0)
            self.run(c0.paragraphs[0], k, bold=True, size=self.base - 1)
            self.run(c1.paragraphs[0], v, size=self.base - 1)
        self.para(space_after=2)

    def callout(self, text, tone="info"):
        t = self.doc.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = t.rows[0].cells[0]
        warn = tone == "warn"
        shade(c, "FFF4E5" if warn else SOFT_HEX)
        cell_borders(c, left=(24, "C55A11" if warn else GREEN_HEX), top=None, bottom=None, right=None)
        cell_margins(t, 90, 90, 160, 140)
        set_widths(t, [CONTENT_W])
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        self.run(p, (self.L["warn"] if warn else self.L["note"]) + ": ", bold=True, color=RGBColor(0xC5, 0x5A, 0x11) if warn else GREEN)
        self.run(p, text, size=self.base - 0.5)
        self.para(space_after=2)

    def formula(self, text):
        t = self.doc.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = t.rows[0].cells[0]
        shade(c, SOFT_HEX)
        cell_borders(c, left=(24, GREEN_HEX), top=None, bottom=None, right=None)
        cell_margins(t, 80, 80, 160, 140)
        set_widths(t, [CONTENT_W])
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        self.run(p, text, size=self.base - 0.5)
        self.para(space_after=2)

    def verdict(self, text):
        t = self.doc.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = t.rows[0].cells[0]
        cell_borders(c, left=(6, GREEN_HEX), top=(6, GREEN_HEX), bottom=(6, GREEN_HEX), right=(6, GREEN_HEX))
        cell_margins(t, 90, 90, 160, 160)
        set_widths(t, [CONTENT_W])
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        self.run(p, text, bold=True)
        self.para(space_after=2)

    def image(self, b, base_dir):
        path = b.get("path", "")
        if not os.path.isabs(path):
            path = os.path.join(base_dir, path)
        if not os.path.exists(path):
            self.callout(f"[image not found: {b.get('path')}]", "warn")
            return
        self.doc.add_picture(path, width=Cm(min(float(b.get("width_cm", 15)), CONTENT_W)))
        self.doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if b.get("caption"):
            self.para(b["caption"], align=WD_ALIGN_PARAGRAPH.CENTER, size=8.5, color=MUTED, italic=True)

    def signatures(self, parties):
        """Side-by-side signing blocks (HTML .sig): party title in green, rows Name/Position/Signature/Date."""
        parties = parties or []
        if not parties:
            return
        L = self.L
        per_row = 2
        gap = 0.8
        for i in range(0, len(parties), per_row):
            group = parties[i:i + per_row]
            k = len(group)
            colw = (CONTENT_W - gap * (k - 1)) / k
            widths = []
            for j in range(k):
                widths.append(colw)
                if j < k - 1:
                    widths.append(gap)
            outer = self.doc.add_table(rows=1, cols=len(widths))
            outer.alignment = WD_TABLE_ALIGNMENT.CENTER
            outer.autofit = False
            table_borders(outer, edges=())
            cell_margins(outer, 0, 0, 0, 0)
            set_widths(outer, widths)
            no_split(outer.rows[0])
            for j, party in enumerate(group):
                cell = outer.rows[0].cells[2 * j]
                hp = cell.paragraphs[0]
                hp.paragraph_format.space_after = Pt(3)
                self.run(hp, party.get("title", ""), bold=True, color=GREEN)
                inner = cell.add_table(rows=4, cols=2)
                inner.autofit = False
                table_borders(inner)
                cell_margins(inner, 60, 60, 100, 100)
                lw = round(colw * 0.34, 2)
                set_widths(inner, [lw, colw - lw])
                labels = [L["name"], L["position"], L["signature"], L["date"]]
                values = [party.get("name", ""), party.get("position", ""), "", party.get("date", "")]
                for r in range(4):
                    a, bcell = inner.rows[r].cells
                    no_split(inner.rows[r])
                    shade(a, SOFT_HEX)
                    a.paragraphs[0].paragraph_format.space_after = Pt(0)
                    bcell.paragraphs[0].paragraph_format.space_after = Pt(0)
                    self.run(a.paragraphs[0], labels[r], bold=True, size=self.base - 1)
                    self.run(bcell.paragraphs[0], values[r], size=self.base - 1)
                    if r == 2:
                        inner.rows[r].height = Cm(1.4)
            self.para(space_after=6)

    def letter_signer(self, signer):
        """Single right-aligned signature for letters."""
        if not signer:
            return
        self.para(space_after=0)
        for text, kw in ((signer.get("closing", ""), {}), ("", {}), ("", {}), ("", {}),
                         (signer.get("name", ""), {"bold": True}), (signer.get("position", ""), {"color": MUTED})):
            p = self.para(text, space_after=0, **kw)
            p.paragraph_format.left_indent = Cm(9.5)

    # ---- sections
    def section(self, s, base_dir):
        title = s.get("title", "")
        if s.get("annex"):
            if self.annex_no == 0:
                self.page_break()
                self.heading(self.L["annex"], 2)
            letter = self.L["annex_letters"][self.annex_no % 10]
            self.annex_no += 1
            self.heading(f"{letter}. {title}", 3)
        elif s.get("numbered", True) and title:
            self.sec_no += 1
            self.heading(f"{self.sec_no}. {title}", 2)
        elif title:
            self.heading(title, 2)
        for b in s.get("blocks", []):
            self.block(b, base_dir)

    def block(self, b, base_dir):
        t = b.get("type")
        if t == "p":
            self.body_para(b.get("text", ""))
        elif t == "h":
            self.heading(b.get("text", ""), int(b.get("level", 3)))
        elif t == "bullets":
            self.bullets(b.get("items", []))
        elif t == "numbered":
            self.bullets(b.get("items", []), numbered=True)
        elif t == "table":
            self.table(b)
        elif t == "kv":
            self.kv(b.get("rows", []))
        elif t == "callout":
            self.callout(b.get("text", ""), b.get("tone", "info"))
        elif t == "formula":
            self.formula(b.get("text", ""))
        elif t == "verdict":
            self.verdict(b.get("text", ""))
        elif t == "image":
            self.image(b, base_dir)
        elif t == "signatures":
            self.signatures(b.get("parties", []))
        elif t == "pagebreak":
            self.page_break()
        else:
            print(f"! unknown block type: {t}", file=sys.stderr)

    # ---- document types
    def build_report(self, base_dir):
        self.meta_line()
        self.title_block()
        self.info_table(self.m.get("info", []))
        for s in self.d.get("sections", []):
            self.section(s, base_dir)

    def build_letter(self, base_dir):
        m, L = self.m, self.L
        year = str(m.get("date", ""))[:4] or str(dt.date.today().year)
        ref = m.get("ref_no") or ("..........." + L["ref_suffix"] + year)
        p = self.para(space_after=0, align=WD_ALIGN_PARAGRAPH.RIGHT)
        self.run(p, f'{L["ref"]} {ref}', size=self.base - 0.5)
        p = self.para(space_after=12, align=WD_ALIGN_PARAGRAPH.RIGHT)
        self.run(p, f'{L["date"]} {fmt_date(m.get("date"), self.lang)}', size=self.base - 0.5)
        if m.get("to"):
            p = self.para(space_after=4)
            self.run(p, f'{L["to"]}: ', bold=True)
            self.run(p, m["to"])
        if m.get("subject"):
            p = self.para(space_after=4)
            self.run(p, f'{L["subject"]}: ', bold=True)
            self.run(p, m["subject"])
        if m.get("attention"):
            p = self.para(space_after=4)
            self.run(p, f'{L["attention"]}: ', bold=True)
            self.run(p, m["attention"])
        rule = self.para(space_after=10)
        para_border(rule, "bottom", RULE_HEX, 4, 1)
        for s in self.d.get("sections", []):
            s.setdefault("numbered", False)
            self.section(s, base_dir)
        self.letter_signer(m.get("signer"))

    def build_agreement(self, base_dir):
        self.meta_line()
        self.title_block()
        self.info_table(self.m.get("info", []))
        sections = self.d.get("sections", [])
        for s in sections:
            if not s.get("annex"):
                self.section(s, base_dir)
        if self.m.get("parties"):
            self.heading(self.L["approval"], 2)
            self.signatures(self.m["parties"])
        for s in sections:
            if s.get("annex"):
                self.section(s, base_dir)

    def build(self, out, base_dir):
        self._header_footer()
        kind = (self.m.get("type") or "report").lower()
        {"report": self.build_report, "letter": self.build_letter, "agreement": self.build_agreement}.get(kind, self.build_report)(base_dir)
        cp = self.doc.core_properties
        cp.title = self.m.get("title", "")
        cp.subject = self.m.get("project", "")
        cp.author = self.m.get("author", "Dansavanh")
        self.doc.save(out)


# ---------------------------------------------------------------- check / pdf
def check(data):
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
            if b.get("type") == "table":
                n = len(b.get("columns", [])) or None
                for r in b.get("rows", []):
                    if n and len(r) != n:
                        issues.append(f"table in '{s.get('title')}': row has {len(r)} cells, header has {n}")
                        break
    return issues


def to_pdf(docx_path):
    import win32com.client  # pywin32
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    pdf = os.path.splitext(os.path.abspath(docx_path))[0] + ".pdf"
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        doc.Fields.Update()
        doc.ExportAsFixedFormat(pdf, 17)
        doc.Close(False)
    finally:
        word.Quit()
    return pdf


def main():
    ap = argparse.ArgumentParser(description="Build a Dansavanh letterhead Word document from JSON")
    ap.add_argument("data")
    ap.add_argument("--lang", required=True, choices=["lo", "en", "zh"])
    ap.add_argument("--out")
    ap.add_argument("--status", choices=["draft", "final"])
    ap.add_argument("--pdf", action="store_true", help="also export PDF through Microsoft Word")
    ap.add_argument("--check", action="store_true", help="validate JSON only")
    a = ap.parse_args()
    with open(a.data, encoding="utf-8") as fh:
        raw = json.load(fh)
    data = resolve(raw, a.lang)
    issues = check(data)
    for i in issues:
        print("! " + i)
    if a.check:
        sys.exit(1 if issues else 0)
    out = a.out or os.path.splitext(a.data)[0] + f"_{a.lang}.docx"
    DsvDoc(data, a.lang, a.status).build(out, os.path.dirname(os.path.abspath(a.data)))
    print(f"saved: {out}" + (f"  ({len(issues)} warnings)" if issues else ""))
    if a.pdf:
        print(f"pdf:   {to_pdf(out)}")


if __name__ == "__main__":
    main()
