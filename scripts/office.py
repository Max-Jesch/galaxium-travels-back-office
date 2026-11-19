"""House style for the office files in the drive: workbooks, decks and documents.

`scripts/build.py`, `scripts/routine.py` and `scripts/documents.py` write every .xlsx,
.pptx and .docx through this module, so that they look like they came from one company:

    Book   a workbook (XlsxWriter): styled tables, totals as real formulas, charts
    Deck   a 16:9 deck on the Galaxium master: bullets, tables, charts, figures, timelines
    Doc    a document: title block, header and footer, tables, clauses, signatures

Colours and voice follow marketing/brand/brand-guidelines.md. Documents that come from
somebody else (a vendor's draft, the auditors' letter) take another `Brand`.

Files are written with fixed timestamps, so a rebuild with unchanged input is
byte-identical and shows no diff.
"""

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import xlsxwriter
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt as DPt, RGBColor as DRGB
from lxml import etree
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.util import Emu, Inches, Pt

FIXED_TIME = datetime(2026, 11, 4, 9, 0)

NIGHT, ORANGE, STATION = "1B1A2E", "E8743B", "F4F3FF"  # the three brand colours
INK, MUTED, LINE, PALE = "2A2940", "6B6A80", "D9D7EA", "F4F3FF"
SERIES = [ORANGE, NIGHT, "8A88B5", "F2B08F", "4A4870", "C9C7E0"]  # chart series, in this order
OK, AMBER, WARN = "E3F1E7", "FBEBD3", "F9DDD6"  # status fills
HEAD, TEXT = "Georgia", "Arial"  # the display face and the text face

# Owner and last change of each file written under a drive root, for that root's _index.yaml.
META: dict[str, dict] = {}


def register(path: Path, owner: str | None, modified: str | date | None) -> None:
    if owner or modified:
        parts = path.as_posix().split("/drive/", 1)
        if len(parts) == 2:
            META[str(path)] = {"rel": parts[1], "root": parts[0] + "/drive",
                               **({"owner": owner} if owner else {}), **({"modified": str(modified)} if modified else {})}


def write_index(root: Path) -> None:
    """Write `_index.yaml` for the files registered under this drive root."""
    import yaml
    rows = {m["rel"]: {k: v for k, v in m.items() if k in ("owner", "modified")}
            for m in META.values() if Path(m["root"]) == root}
    if rows:
        (root / "_index.yaml").write_text("# Generated. Owner and last change of the generated files.\n"
                                          + yaml.safe_dump(dict(sorted(rows.items())), allow_unicode=True), encoding="utf-8")


def normalize(path: Path, when: datetime = FIXED_TIME) -> None:
    """Rewrite an office file with fixed timestamps, inside embedded workbooks too."""
    path.write_bytes(_normalize(path.read_bytes(), when))


def _normalize(blob: bytes, when: datetime) -> bytes:
    stamp = when.strftime("%Y-%m-%dT%H:%M:%SZ")
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob)) as src, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename.endswith(".xlsx"):
                data = _normalize(data, when)
            elif item.filename == "docProps/core.xml":
                data = re.sub(rb"(<dcterms:(created|modified)[^>]*>)[^<]*(</dcterms:)", lambda m: m.group(1) + stamp.encode() + m.group(3), data)
            info = zipfile.ZipInfo(item.filename, date_time=(when.year, when.month, when.day, when.hour, when.minute, 0))
            info.compress_type, info.external_attr = zipfile.ZIP_DEFLATED, 0o644 << 16
            dst.writestr(info, data)
    return out.getvalue()


def logo_png(colour: str = NIGHT, size: int = 240) -> bytes:
    """The mark: a planet with one orbit. Drawn here so that no artwork file is needed."""
    big = size * 4
    ring = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    ImageDraw.Draw(ring).ellipse((big * 0.04, big * 0.33, big * 0.96, big * 0.67), outline="#" + colour, width=big // 22)
    ring = ring.rotate(24, resample=Image.BICUBIC)
    img = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    ImageDraw.Draw(img).ellipse((big * 0.29, big * 0.29, big * 0.71, big * 0.71), fill="#" + ORANGE)
    lower = ring.crop((0, big // 2, big, big))  # the near half of the orbit passes in front of the planet
    back = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    back.alpha_composite(ring)
    back.alpha_composite(img)
    back.alpha_composite(lower, (0, big // 2))
    buf = io.BytesIO()
    back.resize((size, size), Image.LANCZOS).save(buf, "PNG")
    return buf.getvalue()


# ====================================================================== workbooks

class F:
    """A formula cell with the value it evaluates to, so that readers without a
    calculation engine still see the number. `{r}` is the row number in the sheet."""

    def __init__(self, formula: str, value) -> None:
        self.formula, self.value = formula, value


def col(i: int) -> str:
    return xlsxwriter.utility.xl_col_to_name(i)


class Book:
    def __init__(self, path: Path, *, owner: str | None = None, modified: str | date | None = None, title: str = "") -> None:
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.wb = xlsxwriter.Workbook(str(self.path), {"in_memory": True, "strings_to_numbers": False, "strings_to_formulas": False})
        self.wb.set_properties({"title": title or self.path.stem, "author": owner or "Galaxium Travels", "company": "Galaxium Travels", "created": FIXED_TIME})
        self._formats: dict[tuple, object] = {}
        register(self.path, owner, modified)

    def fmt(self, **props):
        key = tuple(sorted(props.items()))
        if key not in self._formats:
            self._formats[key] = self.wb.add_format({"font_name": TEXT, "font_size": 10, "valign": "vcenter", **props})
        return self._formats[key]

    def sheet(self, name: str, rows: list[dict], *, title: str | None = None, subtitle: str | None = None,
              widths: dict[str, float] | None = None, formats: dict[str, str] | None = None, total: dict[str, str] | None = None,
              total_label: str = "Total", status: dict[str, dict[str, str]] | None = None, bars: list[str] | None = None,
              plain: bool = False, wrap: list[str] | None = None, tab: str | None = None, freeze: bool = True):
        """One table per sheet. `total` maps a column to "sum" or "avg" and writes a formula.
        `status` colours a cell by the text it contains. `plain` is the look of a system export.
        Returns (worksheet, first data row, last data row), rows counted from 0."""
        ws = self.wb.add_worksheet(name[:31])
        ws.hide_gridlines(2)
        if tab or not plain:
            ws.set_tab_color("#" + (tab or NIGHT))
        keys, r = list(rows[0]), 0
        if title:
            ws.merge_range(0, 0, 0, max(len(keys) - 1, 1), title, self.fmt(font_name=HEAD, font_size=15, font_color="#FFFFFF", bg_color="#" + NIGHT, indent=1))
            ws.set_row(0, 32)
            ws.merge_range(1, 0, 1, max(len(keys) - 1, 1), subtitle or "", self.fmt(font_size=9, italic=True, font_color="#" + MUTED, bg_color="#" + PALE, indent=1, bottom=2, bottom_color="#" + ORANGE))
            ws.set_row(1, 20)
            r = 3
        head = (self.fmt(bold=True, bg_color="#E7E6EE", bottom=1, bottom_color="#9F9DB5", text_wrap=True) if plain else
                self.fmt(bold=True, font_color="#FFFFFF", bg_color="#" + NIGHT, text_wrap=True, border=1, border_color="#" + NIGHT))
        for c, key in enumerate(keys):
            ws.write(r, c, key, head)
        ws.set_row(r, 22)
        first = r + 1
        numeric = {k for k in keys if any(isinstance(row[k], (int, float)) or isinstance(row[k], F) for row in rows)}
        for i, row in enumerate(rows):
            band = {} if plain or i % 2 == 0 else {"bg_color": "#FAFAFD"}
            for c, key in enumerate(keys):
                v = row[key]
                value = v.value if isinstance(v, F) else v
                props = {"bottom": 1, "bottom_color": "#" + LINE, **band}
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    props["num_format"] = (formats or {}).get(key) or ("0" if _is_id(key) else "#,##0" if isinstance(value, int) else "#,##0.0")
                elif key in (wrap or []):
                    props["text_wrap"], props["valign"] = True, "top"
                for needle, fill in ((status or {}).get(key) or {}).items():
                    if needle.lower() in str(value).lower():
                        props["bg_color"] = "#" + fill
                        break
                if isinstance(v, F):
                    ws.write_formula(first + i, c, v.formula.replace("{r}", str(first + i + 1)), self.fmt(**props), v.value)
                elif value == "" or value is None:
                    ws.write_blank(first + i, c, None, self.fmt(**props))
                else:
                    ws.write(first + i, c, value, self.fmt(**props))
        last = first + len(rows) - 1
        if total:
            tr = last + 1
            base = {"bold": True, "top": 1, "top_color": "#" + NIGHT, "bottom": 6, "bottom_color": "#" + NIGHT}
            for c, key in enumerate(keys):
                if key in total:
                    values = [(row[key].value if isinstance(row[key], F) else row[key]) for row in rows]
                    values = [v for v in values if isinstance(v, (int, float))]
                    fn = {"sum": "SUM", "avg": "AVERAGE"}[total[key]]
                    result = sum(values) if fn == "SUM" else (sum(values) / len(values) if values else 0)
                    ws.write_formula(tr, c, f"={fn}({col(c)}{first + 1}:{col(c)}{last + 1})",
                                     self.fmt(num_format=(formats or {}).get(key) or ("#,##0" if fn == "SUM" else "#,##0.0"), **base), result)
                else:
                    ws.write(tr, c, total_label if c == 0 else "", self.fmt(**base))
        for key in bars or []:
            c = keys.index(key)
            ws.conditional_format(first, c, last, c, {"type": "data_bar", "bar_color": "#F2B08F"})
        for c, key in enumerate(keys):
            if widths and key in widths:
                width = widths[key]
            else:
                longest = max([len(str(key)) + 2] + [len(_shown(row[key])) + 2 for row in rows[:80]])
                width = min(64, max(11, longest + (2 if key in numeric else 0)))
            ws.set_column(c, c, width)
        if freeze:
            ws.freeze_panes(first, 0)
        if len(rows) > 6:
            ws.autofilter(r, 0, last, len(keys) - 1)
        ws.set_landscape()
        ws.fit_to_pages(1, 0)
        ws.repeat_rows(r)
        return ws, first, last

    def notes(self, name: str, lines: list[str], *, title: str | None = None, plain: bool = False):
        """A sheet of plain sentences: method, source, assumptions."""
        ws = self.wb.add_worksheet(name[:31])
        ws.hide_gridlines(2)
        ws.set_column(0, 0, 110)
        ws.write(0, 0, title or name, self.fmt(bold=True, bg_color="#E7E6EE") if plain else self.fmt(font_name=HEAD, font_size=13, font_color="#FFFFFF", bg_color="#" + NIGHT, indent=1))
        ws.set_row(0, 18 if plain else 26)
        for i, line in enumerate(lines, 1):
            ws.write(i, 0, line, self.fmt(text_wrap=True, valign="top", bottom=1, bottom_color="#" + LINE, indent=0 if plain else 1))
        return ws

    def chart(self, ws, kind: str, title: str, sheet: str, cats: tuple[int, int, int], series: list[tuple[str, int]], anchor: str,
              *, width: int = 560, height: int = 320, number_format: str = "#,##0", labels: bool = True, subtype: str | None = None, y_title: str | None = None) -> None:
        """`cats` is (first row, last row, column); each series is (name, column) over the same rows."""
        first, last, cat_col = cats
        chart = self.wb.add_chart({"type": kind, **({"subtype": subtype} if subtype else {})})
        for i, (label, c) in enumerate(series):
            colour = "#" + SERIES[i % len(SERIES)]
            spec = {"name": label, "categories": [sheet, first, cat_col, last, cat_col], "values": [sheet, first, c, last, c],
                    "data_labels": {"value": labels, "num_format": number_format, "font": {"name": TEXT, "size": 8, "color": "#" + MUTED}} if labels else {}}
            if kind == "line":
                spec.update(line={"color": colour, "width": 2.25}, marker={"type": "circle", "size": 6, "fill": {"color": colour}, "border": {"color": "#FFFFFF"}})
            elif kind in ("pie", "doughnut"):
                spec["points"] = [{"fill": {"color": "#" + SERIES[k % len(SERIES)]}, "border": {"color": "#FFFFFF"}} for k in range(last - first + 1)]
                spec["data_labels"] = {"percentage": True, "font": {"name": TEXT, "size": 9, "color": "#FFFFFF", "bold": True}}
            else:
                spec.update(fill={"color": colour}, border={"none": True}, gap=70)
            chart.add_series(spec)
        chart.set_title({"name": title, "name_font": {"name": HEAD, "size": 12, "bold": False, "color": "#" + NIGHT}, "overlay": False})
        if kind not in ("pie", "doughnut"):
            axis = {"num_font": {"name": TEXT, "size": 9, "color": "#" + MUTED}, "line": {"color": "#" + LINE}}
            chart.set_x_axis({**axis, "major_tick_mark": "none"})
            chart.set_y_axis({**axis, "num_format": number_format, "major_gridlines": {"visible": True, "line": {"color": "#ECEBF4", "width": 0.75}}, "line": {"none": True},
                              **({"name": y_title, "name_font": {"name": TEXT, "size": 9, "bold": False, "color": "#" + MUTED}} if y_title else {})})
        chart.set_legend({"none": True} if len(series) == 1 and kind not in ("pie", "doughnut") else {"position": "bottom", "font": {"name": TEXT, "size": 9}})
        chart.set_chartarea({"border": {"color": "#" + LINE}, "fill": {"color": "#FFFFFF"}})
        chart.set_size({"width": width, "height": height})
        ws.insert_chart(anchor, chart)

    def close(self) -> None:
        self.wb.close()
        normalize(self.path)


def _is_id(key: str) -> bool:
    k = str(key).lower()
    return k.endswith("id") or k in ("year", "no", "week", "issue", "asset", "level", "clause")


def _shown(v) -> str:
    v = v.value if isinstance(v, F) else v
    return f"{v:,}" if isinstance(v, int) and not isinstance(v, bool) else str(v)


def table_book(path: Path, sheets: dict[str, list[dict]], *, plain: bool = False, owner: str | None = None,
               modified: str | date | None = None, **sheet_options) -> None:
    """The short form: every sheet is one table (or, with a single `note` column, a notes sheet)."""
    book = Book(path, owner=owner, modified=modified)
    for name, rows in sheets.items():
        if list(rows[0]) == ["note"]:
            book.notes(name, [r["note"] for r in rows], plain=plain)
        else:
            book.sheet(name, rows, plain=plain, **(sheet_options.get(name) or {}))
    book.close()


# ====================================================================== decks

W, H = 12192000, 6858000  # 16:9
LEFT, RIGHT = Inches(0.6), Inches(12.73)
BODY_TOP, BODY_BOTTOM = Inches(1.55), Inches(6.7)
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NSMAP = f'xmlns:a="{A}" xmlns:p="{P}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def rgb(hex6: str) -> RGBColor:
    return RGBColor.from_string(hex6)


def _lst_style(size: int, colour: str, *, align: str = "l", head: bool = False, bold: bool = False) -> object:
    face = "+mj-lt" if head else "+mn-lt"
    return parse_xml(f'<a:lstStyle {NSMAP}><a:lvl1pPr marL="0" indent="0" algn="{align}"><a:buNone/>'
                     f'<a:defRPr sz="{size * 100}" b="{int(bold)}"><a:solidFill><a:srgbClr val="{colour}"/></a:solidFill>'
                     f'<a:latin typeface="{face}"/></a:defRPr></a:lvl1pPr></a:lstStyle>')


def _restyle(placeholder, size: int, colour: str, **kw) -> None:
    body = placeholder._element.txBody
    old = body.find(f"{{{A}}}lstStyle")
    new = _lst_style(size, colour, **kw)
    if old is not None:
        body.replace(old, new)
    else:
        body.insert(1, new)


def _place(shape, x, y, w, h) -> None:
    shape.left, shape.top, shape.width, shape.height = int(x), int(y), int(w), int(h)


def _paint(shape, fill: str | None, line: str | None = None, width: float = 0.75) -> None:
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(fill)
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = rgb(line)
        shape.line.width = Pt(width)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False


def _runs(paragraph, text: str, *, size: float | None = None, colour: str | None = None, bold: bool = False,
          italic: bool = False, font: str | None = None, spacing: int | None = None) -> None:
    """Add text to a paragraph. `**words**` are set bold."""
    for k, piece in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if not piece:
            continue
        run = paragraph.add_run()
        run.text = piece
        f = run.font
        if size:
            f.size = Pt(size)
        if colour:
            f.color.rgb = rgb(colour)
        f.bold = True if (bold or k % 2) else None
        if italic:
            f.italic = True
        if font:
            f.name = font
        if spacing:
            run._r.get_or_add_rPr().set("spc", str(spacing))


def _box(slide, x, y, w, h, paragraphs: list, *, anchor=MSO_ANCHOR.TOP, margin: float = 0.0, fill: str | None = None, line: str | None = None):
    """A text box. Each paragraph is a string or (text, {size, colour, bold, italic, font, align, before, spacing})."""
    box = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    if fill or line:
        _paint(box, fill, line)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(margin)
    for i, item in enumerate(paragraphs):
        text, opt = (item, {}) if isinstance(item, str) else item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "r": PP_ALIGN.RIGHT, "c": PP_ALIGN.CENTER}[opt.get("align", "l")]
        if "before" in opt:
            p.space_before = Pt(opt["before"])
        _runs(p, text, size=opt.get("size", 14), colour=opt.get("colour", INK), bold=opt.get("bold", False),
              italic=opt.get("italic", False), font=opt.get("font"), spacing=opt.get("spacing"))
    return box


def _mark(slide, x, y, size, ring: str) -> None:
    """The planet-and-orbit mark as two shapes."""
    orbit = slide.shapes.add_shape(MSO_SHAPE.OVAL, int(x), int(y + size * 0.32), int(size), int(size * 0.36))
    _paint(orbit, None, ring, 1.5)
    orbit.rotation = -24
    planet = slide.shapes.add_shape(MSO_SHAPE.OVAL, int(x + size * 0.3), int(y + size * 0.3), int(size * 0.4), int(size * 0.4))
    _paint(planet, ORANGE)


def new_presentation() -> Presentation:
    """A presentation with the Galaxium master: 16:9, brand fonts and colours, the mark on
    the title and section layouts, a footer with the wordmark and the slide number.
    The eleven standard layouts keep their order, so `slide_layouts[1]` is still title and content."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(W), Emu(H)
    master = prs.slide_master
    for part in [master, *prs.slide_layouts]:  # stretch the 4:3 placeholders to 16:9
        for xfrm in part._element.iter(f"{{{A}}}xfrm"):
            for tag, attr in (("off", "x"), ("ext", "cx")):
                el = xfrm.find(f"{{{A}}}{tag}")
                if el is not None:
                    el.set(attr, str(int(el.get(attr)) * 4 // 3))

    theme_part = master.part.part_related_by(RT.THEME)
    theme = etree.fromstring(theme_part.blob)
    ns = {"a": A}
    theme.find(".//a:majorFont/a:latin", ns).set("typeface", HEAD)
    theme.find(".//a:minorFont/a:latin", ns).set("typeface", TEXT)
    scheme = theme.find(".//a:clrScheme", ns)
    scheme.set("name", "Galaxium")
    for name, value in {"dk1": NIGHT, "lt1": "FFFFFF", "dk2": NIGHT, "lt2": STATION, "accent1": ORANGE, "accent2": NIGHT, "accent3": "8A88B5",
                        "accent4": "F2B08F", "accent5": "4A4870", "accent6": "C9C7E0", "hlink": ORANGE, "folHlink": "8A88B5"}.items():
        el = scheme.find(f"a:{name}", ns)
        for child in list(el):
            el.remove(child)
        etree.SubElement(el, f"{{{A}}}srgbClr").set("val", value)
    theme_part._blob = etree.tostring(theme, xml_declaration=True, encoding="UTF-8", standalone=True)

    styles = master._element.find(f"{{{P}}}txStyles")
    common = 'algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"'
    levels = "".join(
        f'<a:lvl{n}pPr marL="{285750 + 342900 * (n - 1)}" indent="-285750" {common}><a:lnSpc><a:spcPct val="100000"/></a:lnSpc>'
        f'<a:spcBef><a:spcPts val="{1000 if n == 1 else 500}"/></a:spcBef><a:buClr><a:srgbClr val="{ORANGE}"/></a:buClr><a:buFont typeface="Arial"/>'
        f'<a:buChar char="{"•" if n == 1 else "–"}"/><a:defRPr sz="{size}" kern="1200"><a:solidFill><a:srgbClr val="{INK}"/></a:solidFill>'
        f'<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl{n}pPr>'
        for n, size in enumerate([2200, 1800, 1600, 1400, 1400], 1))
    new = parse_xml(
        f'<p:txStyles {NSMAP}><p:titleStyle><a:lvl1pPr {common}><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>'
        f'<a:defRPr sz="3200" kern="1200"><a:solidFill><a:srgbClr val="{NIGHT}"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/>'
        f'<a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle>{levels}</p:bodyStyle></p:txStyles>')
    other = styles.find(f"{{{P}}}otherStyle")
    if other is not None:
        new.append(other)
    master._element.replace(styles, new)

    for ph in master.placeholders:
        kind = ph.placeholder_format.type
        if str(kind).startswith("TITLE"):
            _place(ph, LEFT, Inches(0.55), RIGHT - LEFT, Inches(0.85))
            ph.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        elif str(kind).startswith("BODY"):
            _place(ph, LEFT, BODY_TOP, RIGHT - LEFT, BODY_BOTTOM - BODY_TOP)
        else:  # date, footer and slide number: the master draws its own footer
            ph._element.getparent().remove(ph._element)

    scratch = prs.slides.add_slide(prs.slide_layouts[6])

    def move_to(target) -> None:
        tree = target._element.cSld.spTree
        used = [int(e.get("id")) for e in tree.iter(f"{{{P}}}cNvPr")]
        nxt = max(used + [1]) + 1
        for el in list(scratch.shapes._spTree)[2:]:
            for nv in el.iter(f"{{{P}}}cNvPr"):
                nv.set("id", str(nxt))
                nxt += 1
            tree.append(el)

    # Master: a short orange rule above the title, a hairline and the wordmark at the foot.
    _paint(scratch.shapes.add_shape(MSO_SHAPE.RECTANGLE, LEFT, Inches(0.42), Inches(0.55), Inches(0.06)), ORANGE)
    _paint(scratch.shapes.add_shape(MSO_SHAPE.RECTANGLE, LEFT, Inches(6.93), RIGHT - LEFT, Emu(9525)), LINE)
    _mark(scratch, LEFT, Inches(7.0), Inches(0.26), NIGHT)
    _box(scratch, Inches(0.93), Inches(7.03), Inches(2.2), Inches(0.22), [("GALAXIUM TRAVELS", {"size": 8, "colour": MUTED, "bold": True, "spacing": 200})], anchor=MSO_ANCHOR.MIDDLE)
    number = _box(scratch, Inches(11.9), Inches(7.03), Inches(0.83), Inches(0.22), [("0", {"size": 9, "colour": MUTED, "align": "r"})], anchor=MSO_ANCHOR.MIDDLE)
    run = number.text_frame.paragraphs[0].runs[0]._r
    field = parse_xml(f'<a:fld {NSMAP} id="{{6A1C1E6B-0F4E-4C0B-9A52-2B1D3E5F7A90}}" type="slidenum"><a:rPr lang="en-GB" sz="900">'
                      f'<a:solidFill><a:srgbClr val="{MUTED}"/></a:solidFill></a:rPr><a:t>‹#›</a:t></a:fld>')
    run.getparent().replace(run, field)
    move_to(master)

    for index in (0, 2):  # title slide and section header: deep night, no master furniture
        layout = prs.slide_layouts[index]
        layout._element.set("showMasterSp", "0")
        layout.background.fill.solid()
        layout.background.fill.fore_color.rgb = rgb(NIGHT)
        _mark(scratch, Inches(0.85), Inches(0.7), Inches(0.62), "FFFFFF")
        _box(scratch, Inches(1.6), Inches(0.83), Inches(4), Inches(0.36), [("GALAXIUM TRAVELS", {"size": 12, "colour": "FFFFFF", "bold": True, "spacing": 300})], anchor=MSO_ANCHOR.MIDDLE)
        _paint(scratch.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(2.75 if index == 0 else 3.0), Inches(0.9), Inches(0.07)), ORANGE)
        if index == 0:
            _box(scratch, Inches(0.9), Inches(6.72), Inches(8), Inches(0.3), [("Interplanetary travel, booked like a train ticket.", {"size": 10, "colour": "B9B7D6", "italic": True})])
        move_to(layout)
        title, sub = layout.placeholders[0], layout.placeholders[1]
        if index == 0:
            _place(title, Inches(0.9), Inches(2.95), Inches(11.4), Inches(1.5))
            _place(sub, Inches(0.9), Inches(4.5), Inches(11.4), Inches(0.9))
            _restyle(title, 40, "FFFFFF", head=True)
            _restyle(sub, 16, "C9C7E0")
            title.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        else:
            _place(title, Inches(0.9), Inches(3.2), Inches(11.4), Inches(1.3))
            _place(sub, Inches(0.9), Inches(4.55), Inches(11.4), Inches(0.9))
            _restyle(title, 34, "FFFFFF", head=True)
            _restyle(sub, 16, "C9C7E0")
            title.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        sub.text_frame.vertical_anchor = MSO_ANCHOR.TOP

    ids = prs.slides._sldIdLst
    prs.part.drop_rel(ids[-1].rId)
    ids.remove(ids[-1])
    prs.core_properties.author = "Galaxium Travels"
    prs.core_properties.created = prs.core_properties.modified = FIXED_TIME
    return prs


class Deck:
    """One message per slide, the source of every number in the footer."""

    def __init__(self, path: Path, title: str, subtitle: str, *, owner: str | None = None, modified: str | date | None = None) -> None:
        self.path = Path(path)
        self.prs = new_presentation()
        self.prs.core_properties.title = title
        if owner:
            self.prs.core_properties.author = owner
        register(self.path, owner, modified)
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[0])
        slide.shapes.title.text, slide.placeholders[1].text = title, subtitle

    # -- building blocks
    def _slide(self, title: str, layout: int = 5):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[layout])
        slide.shapes.title.text = title
        return slide

    def _foot(self, slide, source: str | None) -> None:
        if source:
            _box(slide, Inches(3.1), Inches(7.03), Inches(8.7), Inches(0.22), [("Source: " + source, {"size": 8, "colour": MUTED, "italic": True, "align": "r"})], anchor=MSO_ANCHOR.MIDDLE)

    def _list(self, slide, x, y, w, h, items: list, *, size: float = 16, heading: str | None = None):
        paragraphs = []
        if heading:
            paragraphs.append((heading.upper(), {"size": 11, "colour": ORANGE, "bold": True, "spacing": 150}))
        for item in items:
            text, level = (item, 0) if isinstance(item, str) else item
            paragraphs.append((("•  " if level == 0 else "–  ") + text, {"size": size - 2 * level, "colour": INK if level == 0 else MUTED, "before": 9 if level == 0 else 3}))
        box = _box(slide, x, y, w, h, paragraphs)
        for p, item in zip(box.text_frame.paragraphs[1 if heading else 0:], items):
            level = 0 if isinstance(item, str) else item[1]
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", str(int(Inches(0.24 + 0.3 * level))))
            pPr.set("indent", str(-int(Inches(0.24))))
        return box

    # -- slides
    def section(self, title: str, subtitle: str = ""):
        slide = self._slide(title, 2)
        slide.placeholders[1].text = subtitle
        return slide

    def bullets(self, title: str, items: list, *, source: str | None = None, aside: tuple[str, list[str]] | None = None, size: float = 22):
        """Items are strings or (string, level). `aside` is a shaded panel on the right: (heading, lines)."""
        slide = self._slide(title)
        width = (Inches(7.6) if aside else RIGHT - LEFT)
        self._list(slide, LEFT, BODY_TOP, width, BODY_BOTTOM - BODY_TOP, items, size=size)
        if aside:
            self._panel(slide, Inches(8.6), BODY_TOP + Inches(0.1), Inches(4.13), Inches(4.6), *aside)
        self._foot(slide, source)
        return slide

    def _panel(self, slide, x, y, w, h, heading: str, lines: list[str]) -> None:
        _paint(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h), PALE)
        _paint(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, Inches(0.06), h), ORANGE)
        _box(slide, x + Inches(0.3), y + Inches(0.25), w - Inches(0.55), h - Inches(0.5),
             [(heading.upper(), {"size": 11, "colour": ORANGE, "bold": True, "spacing": 150})] + [(line, {"size": 15, "colour": INK, "before": 10}) for line in lines])

    def two(self, title: str, left: tuple[str, list], right: tuple[str, list], *, source: str | None = None, size: float = 18):
        slide = self._slide(title)
        half = (RIGHT - LEFT - Inches(0.5)) // 2
        for k, (heading, items) in enumerate((left, right)):
            x = LEFT + k * (half + Inches(0.5))
            _paint(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, BODY_TOP + Inches(0.05), half, Emu(19050)), NIGHT if k == 0 else ORANGE)
            self._list(slide, x, BODY_TOP + Inches(0.2), half, Inches(4.8), items, size=size, heading=heading)
        self._foot(slide, source)
        return slide

    def figures(self, title: str, cards: list[tuple[str, str]], *, items: list | None = None, source: str | None = None):
        """Large figures in a row, each with a caption, and optionally a few bullets below."""
        slide = self._slide(title)
        n, gap = len(cards), Inches(0.25)
        width = (RIGHT - LEFT - gap * (n - 1)) // n
        for k, (value, caption) in enumerate(cards):
            x = LEFT + k * (width + gap)
            _paint(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, BODY_TOP + Inches(0.1), width, Inches(2.2)), PALE)
            _paint(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, BODY_TOP + Inches(0.1), width, Inches(0.06)), ORANGE)
            _box(slide, x + Inches(0.25), BODY_TOP + Inches(0.3), width - Inches(0.4), Inches(1.6),
                 [(value, {"size": (44 if len(value) < 7 else 34) if n < 5 else 34, "colour": NIGHT, "font": HEAD}), (caption, {"size": 14, "colour": MUTED, "before": 6})])
        if items:
            self._list(slide, LEFT, BODY_TOP + Inches(2.65), RIGHT - LEFT, Inches(2.4), items, size=19)
        self._foot(slide, source)
        return slide

    def table(self, title: str, header: list[str], rows: list[list], *, widths: list[float] | None = None, source: str | None = None,
              note: str | None = None, size: float = 12, total: bool = False, top=None, slide=None, x=None, width=None):
        """A table. Numbers are right-aligned with thousands separators. `total` sets the last row bold."""
        slide = slide or self._slide(title)
        x, width = x if x is not None else LEFT, width if width is not None else RIGHT - LEFT
        top = top if top is not None else BODY_TOP + Inches(0.1)
        size *= 1.25
        row_h = Inches(0.26 + size * 0.016)
        room = Inches(4.95 if not note else 4.4) - (top - BODY_TOP)
        if row_h * (len(rows) + 1) > room:  # a long table: smaller rows rather than off the slide
            shrink = max(0.72, room / (row_h * (len(rows) + 1)))
            size, row_h = size * shrink, int(row_h * shrink)
        frame = slide.shapes.add_table(len(rows) + 1, len(header), int(x), int(top), int(width), int(row_h * (len(rows) + 1)))
        frame._element.graphic.graphicData.tbl.tblPr.remove(frame._element.graphic.graphicData.tbl.tblPr.find(f"{{{A}}}tableStyleId"))
        table = frame.table
        share = widths or [1] * len(header)
        for c, part in enumerate(share):
            table.columns[c].width = int(width * part / sum(share))
        numeric = [all(isinstance(r[c], (int, float)) or r[c] in ("", None) for r in rows) and any(isinstance(r[c], (int, float)) for r in rows) for c in range(len(header))]
        for r, values in enumerate([header] + rows):
            table.rows[r].height = int(row_h)
            for c, value in enumerate(values):
                cell = table.cell(r, c)
                cell.margin_left = cell.margin_right = Inches(0.1)
                cell.margin_top = cell.margin_bottom = Inches(0.045)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(NIGHT if r == 0 else ("FFFFFF" if r % 2 else "F8F8FC"))
                text = "" if value is None else f"{value:,}" if isinstance(value, int) and not isinstance(value, bool) else f"{value:,.1f}" if isinstance(value, float) else str(value)
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.RIGHT if numeric[c] else PP_ALIGN.LEFT
                last = total and r == len(rows)
                _runs(p, text, size=size, colour="FFFFFF" if r == 0 else INK, bold=r == 0 or last)
                _cell_lines(cell, NIGHT if last else LINE)
        if note:
            _box(slide, x, Inches(6.05), width, Inches(0.6), [(note, {"size": 14, "colour": MUTED})])
        self._foot(slide, source)
        return slide

    def chart(self, title: str, kind: str, categories: list[str], series: dict[str, list], *, source: str | None = None, aside: tuple[str, list[str]] | None = None,
              number_format: str = "#,##0", legend: bool | None = None, labels: bool = True, slide=None, box=None):
        """kind: column, bar, stacked, line, pie or doughnut. `aside` is a panel of takeaways on the right."""
        slide = slide or self._slide(title)
        data = CategoryChartData()
        data.categories = categories
        for label, values in series.items():
            data.add_series(label, values)
        kinds = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED, "stacked": XL_CHART_TYPE.COLUMN_STACKED,
                 "line": XL_CHART_TYPE.LINE_MARKERS, "pie": XL_CHART_TYPE.PIE, "doughnut": XL_CHART_TYPE.DOUGHNUT}
        x, y, w, h = box or (LEFT, BODY_TOP, (Inches(7.8) if aside else RIGHT - LEFT), BODY_BOTTOM - BODY_TOP)
        chart = slide.shapes.add_chart(kinds[kind], int(x), int(y), int(w), int(h), data).chart
        chart.has_title = False
        chart.font.size, chart.font.name = Pt(13), TEXT
        chart.font.color.rgb = rgb(MUTED)
        round_kind = kind in ("pie", "doughnut")
        chart.has_legend = legend if legend is not None else (len(series) > 1 or round_kind)
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.RIGHT if round_kind else XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.size = Pt(13)
        plot = chart.plots[0]
        if not round_kind:
            chart.value_axis.has_major_gridlines = True
            chart.value_axis.major_gridlines.format.line.color.rgb = rgb("ECEBF4")
            chart.value_axis.format.line.fill.background()
            chart.value_axis.tick_labels.number_format, chart.value_axis.tick_labels.number_format_is_linked = number_format, False
            chart.category_axis.format.line.color.rgb = rgb(LINE)
            chart.category_axis.has_major_gridlines = False
            if kind in ("column", "bar", "stacked"):
                plot.gap_width = 70
                if kind == "bar":
                    chart.category_axis.reverse_order = True
                    chart.value_axis.tick_label_position  # keep the default; reversed order moves the axis to the top
        if labels:
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.number_format, dl.number_format_is_linked = ("0%" if round_kind else number_format), False
            dl.font.size = Pt(12)
            if round_kind:
                dl.show_percentage, dl.show_value = True, False
                dl.font.color.rgb, dl.font.bold = rgb("FFFFFF"), True
            elif kind in ("column", "bar"):
                dl.position = XL_LABEL_POSITION.OUTSIDE_END
            elif kind == "line":
                dl.position = XL_LABEL_POSITION.ABOVE
        for i, s in enumerate(plot.series):
            colour = rgb(SERIES[i % len(SERIES)])
            if round_kind:
                for k in range(len(categories)):
                    point = s.points[k]
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = rgb(SERIES[k % len(SERIES)])
                    point.format.line.color.rgb = rgb("FFFFFF")
            elif kind == "line":
                s.smooth = False
                s.format.line.color.rgb = colour
                s.format.line.width = Pt(2.5)
                s.marker.format.fill.solid()
                s.marker.format.fill.fore_color.rgb = colour
                s.marker.format.line.color.rgb = rgb("FFFFFF")
            else:
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = colour
                s.format.line.fill.background()
        if aside:
            self._panel(slide, Inches(8.6), BODY_TOP + Inches(0.1), Inches(4.13), Inches(4.6), *aside)
        self._foot(slide, source)
        return slide

    def timeline(self, title: str, steps: list[tuple[str, str]], *, source: str | None = None, items: list | None = None):
        """A horizontal line of dated steps: (when, what)."""
        slide = self._slide(title)
        y = BODY_TOP + Inches(1.15)
        _paint(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, LEFT, y, RIGHT - LEFT, Emu(25400)), NIGHT)
        width = (RIGHT - LEFT) // len(steps)
        for k, (when, what) in enumerate(steps):
            x = LEFT + k * width
            _paint(slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y - Inches(0.09), Inches(0.2), Inches(0.2)), ORANGE, "FFFFFF", 1.5)
            _box(slide, x, y - Inches(0.75), width - Inches(0.15), Inches(0.55), [(when, {"size": 19, "colour": NIGHT, "font": HEAD})], anchor=MSO_ANCHOR.BOTTOM)
            _box(slide, x, y + Inches(0.3), width - Inches(0.2), Inches(2.0), [(what, {"size": 15, "colour": INK})])
        if items:
            self._list(slide, LEFT, y + Inches(2.35), RIGHT - LEFT, Inches(2.2), items, size=18)
        self._foot(slide, source)
        return slide

    def statement(self, title: str, sentence: str, detail: str = "", *, source: str | None = None):
        slide = self._slide(title)
        _paint(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, LEFT, BODY_TOP + Inches(0.5), Inches(0.08), Inches(2.4)), ORANGE)
        _box(slide, LEFT + Inches(0.4), BODY_TOP + Inches(0.45), Inches(10.8), Inches(2.5), [(sentence, {"size": 36, "colour": NIGHT, "font": HEAD})], anchor=MSO_ANCHOR.MIDDLE)
        if detail:
            _box(slide, LEFT + Inches(0.4), BODY_TOP + Inches(3.3), Inches(10.8), Inches(1.5), [(detail, {"size": 19, "colour": MUTED})])
        self._foot(slide, source)
        return slide

    def notes(self, slide, text: str) -> None:
        slide.notes_slide.notes_text_frame.text = text

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(self.path)
        normalize(self.path)


def _cell_lines(cell, colour: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    for side in ("lnL", "lnR", "lnT"):
        tcPr.append(parse_xml(f'<a:{side} {NSMAP} w="0"><a:noFill/></a:{side}>'))
    tcPr.append(parse_xml(f'<a:lnB {NSMAP} w="9525"><a:solidFill><a:srgbClr val="{colour}"/></a:solidFill></a:lnB>'))
    fill = tcPr.find(f"{{{A}}}solidFill")  # the fill has to follow the borders
    if fill is not None:
        tcPr.remove(fill)
        tcPr.append(fill)


# ====================================================================== documents

@dataclass(frozen=True)
class Brand:
    name: str
    primary: str = NIGHT
    accent: str = ORANGE
    head: str = HEAD
    text: str = TEXT
    footer: str = "Galaxium Travels · Cape Meridian, Earth · galaxium.example"
    mark: bool = False  # the planet-and-orbit mark in the header


GALAXIUM = Brand("GALAXIUM TRAVELS", mark=True)
PLAIN = Brand("", primary="222222", accent="888888", head="Times New Roman", text="Times New Roman", footer="")

_PPR_AFTER_BORDER = ("w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
                     "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
                     "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle",
                     "w:rPr", "w:sectPr", "w:pPrChange")


def _w(tag: str, **attrs) -> OxmlElement:
    el = OxmlElement(f"w:{tag}")
    for k, v in attrs.items():
        el.set(qn(f"w:{k}"), str(v))
    return el


def _rule(paragraph, side: str, colour: str, size: int = 6, space: int = 4) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    borders = pPr.find(qn("w:pBdr"))
    if borders is None:
        borders = _w("pBdr")
        pPr.insert_element_before(borders, *_PPR_AFTER_BORDER)
    borders.append(_w(side, val="single", sz=size, space=space, color=colour))


def _face(style, name: str) -> None:
    style.font.name = name
    fonts = style.element.get_or_add_rPr().get_or_add_rFonts()
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(qn(f"w:{attr}"), None)
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{attr}"), name)


def _shade(cell, fill: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    tcPr.insert_element_before(_w("shd", val="clear", color="auto", fill=fill), "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark")


def _cell_borders(cell, **sides) -> None:
    """sides: top/left/bottom/right = (colour, eighths of a point) or None for no line."""
    tcPr = cell._tc.get_or_add_tcPr()
    borders = _w("tcBorders")
    for side in ("top", "left", "bottom", "right"):
        if side in sides:
            spec = sides[side]
            borders.append(_w(side, val="nil") if spec is None else _w(side, val="single", sz=spec[1], space=0, color=spec[0]))
    tcPr.insert_element_before(borders, "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark")


def _field(paragraph, instruction: str, **font) -> None:
    for kind, text in (("begin", None), (None, instruction), ("end", None)):
        run = paragraph.add_run()
        _style_run(run, **font)
        if kind:
            run._r.append(_w("fldChar", fldCharType=kind))
        else:
            instr = OxmlElement("w:instrText")
            instr.set(qn("xml:space"), "preserve")
            instr.text = text
            run._r.append(instr)


def _style_run(run, size: float | None = None, colour: str | None = None, bold: bool | None = None, italic: bool | None = None,
               font: str | None = None, spacing: int | None = None) -> None:
    if size:
        run.font.size = DPt(size)
    if colour:
        run.font.color.rgb = DRGB.from_string(colour)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if font:
        run.font.name = font
        run._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), font)
    if spacing:
        run._r.get_or_add_rPr().append(_w("spacing", val=spacing))


class Doc:
    """A document in the house style, or in another sender's with `brand`."""

    def __init__(self, path: Path, title: str, *, subtitle: str | None = None, label: str = "Internal", meta: list[tuple[str, str]] | None = None,
                 brand: Brand = GALAXIUM, owner: str | None = None, modified: str | date | None = None, title_block: bool = True) -> None:
        self.path, self.brand = Path(path), brand
        register(self.path, owner, modified)
        doc = self.doc = Document()
        doc.core_properties.title, doc.core_properties.author = title, owner or brand.name.title() or "Galaxium Travels"
        doc.core_properties.created = doc.core_properties.modified = FIXED_TIME
        section = doc.sections[0]
        section.page_width, section.page_height = Cm(21), Cm(29.7)
        section.left_margin = section.right_margin = Cm(2.3)
        section.top_margin, section.bottom_margin = Cm(2.6), Cm(2.2)
        section.header_distance, section.footer_distance = Cm(1.1), Cm(1.0)
        self.width = section.page_width - section.left_margin - section.right_margin

        normal = doc.styles["Normal"]
        _face(normal, brand.text)
        normal.font.size, normal.font.color.rgb = DPt(10.5), DRGB.from_string("26253A" if brand is GALAXIUM else "1E1E1E")
        normal.paragraph_format.space_after, normal.paragraph_format.line_spacing = DPt(6), 1.2
        for name, size, bold, before, after in (("Title", 26, False, 0, 4), ("Heading 1", 15, False, 18, 6), ("Heading 2", 11.5, True, 12, 4), ("Heading 3", 10.5, True, 8, 2)):
            style = doc.styles[name]
            _face(style, brand.head)
            style.font.size, style.font.bold, style.font.italic = DPt(size), bold, False
            style.font.color.rgb = DRGB.from_string(brand.primary)
            style.paragraph_format.space_before, style.paragraph_format.space_after = DPt(before), DPt(after)
            style.paragraph_format.keep_with_next = True
            pPr = style.element.get_or_add_pPr()
            for old in pPr.findall(qn("w:pBdr")):
                pPr.remove(old)
        for name in ("List Bullet", "List Number"):
            doc.styles[name].paragraph_format.space_after = DPt(3)

        head = section.header.paragraphs[0]
        for para in (head, section.footer.paragraphs[0]):  # the built-in header style has its own centre and right tabs
            for pos in (4680, 9360):
                para.paragraph_format.tab_stops.add_tab_stop(pos * 635, WD_TAB_ALIGNMENT.CLEAR)
        head.paragraph_format.tab_stops.add_tab_stop(self.width, WD_TAB_ALIGNMENT.RIGHT)
        if brand.mark:
            head.add_run().add_picture(io.BytesIO(logo_png()), height=DPt(15))
            head.add_run("  ")
        _style_run(head.add_run(brand.name), size=8, colour=brand.primary, bold=True, spacing=40 if brand is GALAXIUM else 10)
        _style_run(head.add_run("\t" + label), size=8, colour=MUTED)
        _rule(head, "bottom", LINE if brand is GALAXIUM else brand.accent, 6, 6)
        foot = section.footer.paragraphs[0]
        foot.paragraph_format.tab_stops.add_tab_stop(self.width, WD_TAB_ALIGNMENT.RIGHT)
        _style_run(foot.add_run(brand.footer + "\tPage "), size=8, colour=MUTED)
        _field(foot, "PAGE", size=8, colour=MUTED)
        _style_run(foot.add_run(" of "), size=8, colour=MUTED)
        _field(foot, "NUMPAGES", size=8, colour=MUTED)

        if title_block:
            doc.add_paragraph(title, style="Title")
            if subtitle:
                p = doc.add_paragraph()
                _style_run(p.add_run(subtitle), size=11.5, colour=MUTED)
                p.paragraph_format.space_after = DPt(2)
            bar = doc.add_paragraph()
            bar.paragraph_format.space_after, bar.paragraph_format.line_spacing = DPt(8), 0.3
            _rule(bar, "bottom", brand.accent, 12, 1)
            if meta:
                self.fields(meta)

    # -- text
    def _rich(self, paragraph, text: str, **font) -> None:
        for k, piece in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
            if piece:
                run = paragraph.add_run(piece)
                _style_run(run, **font)
                if k % 2:
                    run.font.bold = True

    def h1(self, text: str):
        return self.doc.add_heading(text, 1)

    def h2(self, text: str):
        return self.doc.add_heading(text, 2)

    def p(self, text: str, *, align: str | None = None, after: float | None = None, **font):
        para = self.doc.add_paragraph()
        self._rich(para, text, **font)
        if align:
            para.alignment = {"r": WD_ALIGN_PARAGRAPH.RIGHT, "c": WD_ALIGN_PARAGRAPH.CENTER, "j": WD_ALIGN_PARAGRAPH.JUSTIFY}[align]
        if after is not None:
            para.paragraph_format.space_after = DPt(after)
        return para

    def bullets(self, items: list[str], numbered: bool = False) -> None:
        for item in items:
            self._rich(self.doc.add_paragraph(style="List Number" if numbered else "List Bullet"), item)

    def clause(self, number: str, title: str, paragraphs: list[str]) -> None:
        """A numbered clause of an agreement: heading, then sub-paragraphs numbered n.1, n.2."""
        self.doc.add_heading(f"{number}. {title}", 2)
        for i, text in enumerate(paragraphs, 1):
            para = self.doc.add_paragraph()
            para.paragraph_format.left_indent, para.paragraph_format.first_line_indent = Cm(1.1), Cm(-1.1)
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            self._rich(para, (f"{number}.{i}\t" if len(paragraphs) > 1 else "\t") + text)
            para.paragraph_format.tab_stops.add_tab_stop(Cm(1.1))

    # -- blocks
    def fields(self, pairs: list[tuple[str, str]], columns: int = 2) -> None:
        """Label and value pairs in a shaded block, as under a title."""
        rows = -(-len(pairs) // columns)
        table = self.doc.add_table(rows=rows, cols=columns * 2)
        table.alignment, table.autofit = WD_TABLE_ALIGNMENT.CENTER, False
        label_w = Cm(2.9)
        value_w = (self.width - label_w * columns) // columns
        for i, (label, value) in enumerate(pairs):
            r, c = i % rows, (i // rows) * 2
            for cell, text, width, font in ((table.cell(r, c), label, label_w, {"size": 8.5, "colour": MUTED, "bold": True}),
                                            (table.cell(r, c + 1), value, value_w, {"size": 9.5})):
                cell.width = width
                para = cell.paragraphs[0]
                para.paragraph_format.space_after = DPt(0)
                self._rich(para, text, **font)
        for row in table.rows:
            for k, cell in enumerate(row.cells):
                cell.width = label_w if k % 2 == 0 else value_w
                _shade(cell, PALE if self.brand is GALAXIUM else "F2F2F2")
                _cell_borders(cell, top=("FFFFFF", 8), bottom=("FFFFFF", 8), left=None, right=None)
        self._gap()

    def table(self, header: list[str], rows: list[list], *, widths: list[float] | None = None, total: bool = False, size: float = 9.5, gap: bool = True):
        """A table. Numbers are right-aligned with thousands separators. `total` sets the last row bold."""
        b = self.brand
        table = self.doc.add_table(rows=len(rows) + 1, cols=len(header))
        table.alignment, table.autofit = WD_TABLE_ALIGNMENT.CENTER, False
        share = widths or [1] * len(header)
        numeric = [all(isinstance(r[c], (int, float)) or r[c] in ("", None) for r in rows) and any(isinstance(r[c], (int, float)) for r in rows) for c in range(len(header))]
        for r, values in enumerate([header] + rows):
            last = total and r == len(rows)
            for c, value in enumerate(values):
                cell = table.cell(r, c)
                cell.width = int(self.width * share[c] / sum(share))
                text = "" if value is None else f"{value:,}" if isinstance(value, int) and not isinstance(value, bool) else str(value)
                para = cell.paragraphs[0]
                para.paragraph_format.space_before, para.paragraph_format.space_after = DPt(3), DPt(3)
                para.paragraph_format.line_spacing = 1.05
                if numeric[c]:
                    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                self._rich(para, text, size=size, colour="FFFFFF" if r == 0 else None, bold=True if (r == 0 or last) else None)
                if r == 0:
                    _shade(cell, b.primary)
                    _cell_borders(cell, top=(b.primary, 4), bottom=(b.primary, 4), left=None, right=None)
                else:
                    if r % 2 == 0:
                        _shade(cell, "FAFAFD" if b is GALAXIUM else "F7F7F7")
                    _cell_borders(cell, top=(b.primary, 8) if last else None, bottom=(b.primary if last else LINE, 6), left=None, right=None)
        table.rows[0]._tr.get_or_add_trPr().append(_w("tblHeader"))
        if gap:
            self._gap()
        return table

    def callout(self, text: str, heading: str | None = None) -> None:
        table = self.doc.add_table(rows=1, cols=1)
        table.alignment, table.autofit = WD_TABLE_ALIGNMENT.CENTER, False
        cell = table.cell(0, 0)
        cell.width = self.width
        _shade(cell, PALE if self.brand is GALAXIUM else "F2F2F2")
        _cell_borders(cell, left=(self.brand.accent, 24), top=None, bottom=None, right=None)
        para = cell.paragraphs[0]
        para.paragraph_format.space_before, para.paragraph_format.space_after = DPt(5), DPt(5)
        if heading:
            _style_run(para.add_run(heading.upper() + "\n"), size=8, colour=self.brand.accent, bold=True, spacing=20)
        self._rich(para, text, size=10)
        self._gap()

    def signatures(self, parties: list[tuple[str, str, str, str]]) -> None:
        """Each party: (for whom, name, role, signed-on line or "" for a blank line)."""
        self._gap(10)
        table = self.doc.add_table(rows=1, cols=len(parties))
        table.alignment, table.autofit = WD_TABLE_ALIGNMENT.CENTER, False
        table.rows[0]._tr.get_or_add_trPr().append(_w("cantSplit"))
        for cell, (party, name, role, signed) in zip(table.rows[0].cells, parties):
            cell.width = self.width // len(parties)
            para = cell.paragraphs[0]
            para.paragraph_format.space_after = DPt(22 if not signed else 6)
            _style_run(para.add_run(party), size=9, bold=True)
            if signed:
                sig = cell.add_paragraph()
                sig.paragraph_format.space_after = DPt(0)
                _style_run(sig.add_run(name), size=15, italic=True, colour=self.brand.primary, font="Georgia")
            line = cell.add_paragraph()
            line.paragraph_format.space_after = DPt(2)
            _rule(line, "top", "555555", 4, 2)
            _style_run(line.add_run(name), size=9.5)
            for text in (role, signed or "Date:"):
                extra = cell.add_paragraph()
                extra.paragraph_format.space_after = DPt(0)
                _style_run(extra.add_run(text), size=9, colour=MUTED)

    def _gap(self, points: float = 6) -> None:
        para = self.doc.add_paragraph()
        para.paragraph_format.space_after, para.paragraph_format.line_spacing = DPt(0), DPt(points)

    def page_break(self) -> None:
        self.doc.add_page_break()

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(self.path)
        normalize(self.path)
