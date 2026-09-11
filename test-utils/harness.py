"""Minimal xlsx test harness: write inputs, recalculate, read results.

Uses LibreOffice headless as the calculation engine so the real formulas are
under test -- nothing here reimplements scoresheet logic.

Stdlib only. Requires `soffice` on PATH.
"""

import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
SHEET1 = "xl/worksheets/sheet1.xml"


# --------------------------------------------------------------------------
# column helpers
# --------------------------------------------------------------------------

def col_to_idx(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    if not m:
        raise ValueError("bad cell ref: %r" % ref)
    return m.group(1), int(m.group(2))


def esc(text):
    return (str(text).replace("&", "&amp;")
                     .replace("<", "&lt;")
                     .replace(">", "&gt;"))


# --------------------------------------------------------------------------
# sheet XML editing
# --------------------------------------------------------------------------

def _find_row(xml, row_no):
    """Locate a <row>, normalising a self-closing row into open/close form."""
    m = re.search(r'<row r="%d"(?: [^>]*?)?(?:/>|>.*?</row>)' % row_no, xml, re.S)
    if not m:
        raise ValueError("row %d not present in sheet" % row_no)
    text = m.group()
    if text.endswith("/>"):
        text = text[:-2] + "></row>"
    return m.start(), m.end(), text


def _cell_xml(ref, value, style):
    if isinstance(value, bool):
        raise ValueError("write booleans as the 'T'/'F' tokens the sheet uses")
    if isinstance(value, (int, float)):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, style, value)
    return ('<c r="%s"%s t="inlineStr"><is><t>%s</t></is></c>'
            % (ref, style, esc(value)))


def set_cell(xml, ref, value):
    """Set a literal value, or delete the cell when value is None/''."""
    col, row_no = split_ref(ref)
    start, end, row = _find_row(xml, row_no)

    cm = re.search(r'<c r="%s"(?: [^>]*?)?(?:/>|>.*?</c>)' % ref, row, re.S)
    style = ""
    if cm:
        sm = re.search(r'\ss="(\d+)"', cm.group())
        if sm:
            style = ' s="%s"' % sm.group(1)

    new = "" if value is None or value == "" else _cell_xml(ref, value, style)

    if cm:
        row = row[:cm.start()] + new + row[cm.end():]
    elif new:
        target = col_to_idx(col)
        at = None
        for m in re.finditer(r'<c r="([A-Z]+)\d+"', row):
            if col_to_idx(m.group(1)) > target:
                at = m.start()
                break
        if at is None:
            at = row.rindex("</row>")
        row = row[:at] + new + row[at:]

    return xml[:start] + row + xml[end:]


def clear_cached_values(xml):
    """Drop cached results from formula cells so LibreOffice recalculates.

    Also strips the cached result's type attribute, which would otherwise
    describe a value that is no longer there.
    """
    def repl(m):
        open_tag = re.sub(r'\st="[^"]*"', "", m.group(1))
        return open_tag + m.group(2)

    # Formula text is XML-escaped, so it never contains a raw "<". Matching it
    # with [^<]* rather than .*? keeps this linear: with .*?, a formula cell
    # that has no cached <v> lets the match run on across every later cell.
    return re.sub(r'(<c\b[^>]*>)(\s*<f\b[^>]*(?:/>|>[^<]*</f>))\s*<v>[^<]*</v>',
                  repl, xml)


# --------------------------------------------------------------------------
# reading results
# --------------------------------------------------------------------------

def _shared_strings(zf):
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    out = []
    for si in root:
        out.append("".join(t.text or "" for t in si.iter("{%s}t" % MAIN)))
    return out


def read_sheet(xlsx_bytes, sheet=SHEET1):
    """Return {cell_ref: value}. Errors come back as their text (#DIV/0! etc)."""
    with _open(xlsx_bytes) as zf:
        strings = _shared_strings(zf)
        root = ET.fromstring(zf.read(sheet))

    values = {}
    for c in root.iter("{%s}c" % MAIN):
        ref = c.get("r")
        ctype = c.get("t")
        if ctype == "inlineStr":
            is_el = c.find("{%s}is" % MAIN)
            values[ref] = "".join(t.text or "" for t in is_el.iter("{%s}t" % MAIN)) if is_el is not None else ""
            continue
        v = c.find("{%s}v" % MAIN)
        if v is None or v.text is None:
            continue
        raw = v.text
        if ctype == "s":
            values[ref] = strings[int(raw)]
        elif ctype in ("str", "e"):
            values[ref] = raw
        elif ctype == "b":
            values[ref] = raw == "1"
        else:
            try:
                values[ref] = float(raw)
            except ValueError:
                values[ref] = raw
    return values


def _open(xlsx_bytes):
    import io
    return zipfile.ZipFile(io.BytesIO(xlsx_bytes))


# --------------------------------------------------------------------------
# recalculation
# --------------------------------------------------------------------------

class Workbook:
    """A scoresheet copy that can take inputs and be recalculated."""

    def __init__(self, path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as zf:
            self.entries = {n: zf.read(n) for n in zf.namelist()}
            self.order = list(zf.namelist())
        self.sheet = self.entries[SHEET1].decode("utf-8")

    def set(self, ref, value):
        self.sheet = set_cell(self.sheet, ref, value)

    def recalc(self, profile_dir=None):
        """Write out, run soffice headless, return {ref: value} of sheet1."""
        sheet = clear_cached_values(self.sheet)
        entries = dict(self.entries)
        entries[SHEET1] = sheet.encode("utf-8")

        tmp = Path(tempfile.mkdtemp(prefix="scoresheet-test-"))
        try:
            src = tmp / "case.xlsx"
            with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as zf:
                for name in self.order:
                    zf.writestr(name, entries[name])

            outdir = tmp / "out"
            outdir.mkdir()
            cmd = ["soffice", "--headless", "--norestore",
                   "--convert-to", "xlsx", "--outdir", str(outdir), str(src)]
            if profile_dir:
                # isolate from any LibreOffice the user already has open
                cmd.insert(1, "-env:UserInstallation=file://%s" % profile_dir)
            proc = subprocess.run(cmd, capture_output=True, timeout=300)
            result = outdir / "case.xlsx"
            if not result.exists():
                raise RuntimeError("recalculation failed:\n%s\n%s"
                                   % (proc.stdout.decode(errors="replace"),
                                      proc.stderr.decode(errors="replace")))
            return read_sheet(result.read_bytes())
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
