"""
PDF & CAD Register
============
Builds an Excel register from SAAD PDF documents and DWG drawings.

How to use: double-click "PDF & CAD Register.exe", pick the folder that holds the
PDF and DWG files and click Create Excel Register. The Excel file is saved in that folder.
To rebuild the exe after changing this file, double-click build_exe.bat.

Documents - reads the document number (starts with SAAD-) and the text after
            "Subject:" on the cover page, and flags the row when the number is
            not the same as the file name.
Drawings  - reads "Drawing Title" and "Serial Number" from the title block, and
            flags the row when the serial number is not the same as the one at
            the end of the file name (e.g. ...-5A-ST121_00.dwg -> ST121).
Shop      - title blocks with "DRAWING NUMBER" instead of "Serial Number": flags
drawings    the row when that number is not the file name without its revision
            (e.g. SAAD-01-DC103-MIS-SDW-ST-2LG-5B-000021_00.dwg).

It can also run without the window:  python pdf_cad_register.py "D:\\path\\to\\folder"
"""

import datetime
import json
import os
import queue
import re
import subprocess
import shutil
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, font, messagebox, ttk

try:
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    import openpyxl
    import ezdxf
except ImportError:
    fitz = openpyxl = ezdxf = None

PAGES_TO_SEARCH = 3                     # labels are only looked for on the first pages
EXTRA_ROTATIONS = (0, 90, 270, 180)     # for drawings whose text is turned on the sheet
DOC_NO_RE = re.compile(r"SAAD-[A-Z0-9]+(?:[-_.][A-Z0-9]+)*", re.I)
WIR_DOC_NO_RE = re.compile(r"\bSAAD-[A-Z0-9]+(?:-[A-Z0-9]+)*-WIR-[A-Z0-9]+(?:-[A-Z0-9]+)*\b", re.I)

DOC_COLUMNS = [
    ("No.", 6, None),
    ("File Name", 46, "file"),
    ("Document No. (PDF / filename)", 36, "doc_no"),
    ("Subject", 50, "subject"),
    ("Rev.", 7, "rev"),
    ("Submittal Date", 15, "date"),
    ("Check", 12, "status"),
    ("Note", 60, "note"),
]
DWG_COLUMNS = [
    ("No.", 6, None),
    ("File Name", 50, "file"),
    ("Serial No. (File Name)", 14, "file_serial"),
    ("Serial No. (in Drawing)", 14, "dwg_serial"),
    ("Drawing Title", 50, "title"),
    ("File Rev.", 8, "file_rev"),
    ("Drawing Rev.", 9, "rev"),
    ("Date", 11, "date"),
    ("Check", 12, "status"),
    ("Note", 60, "note"),
]
SDW_COLUMNS = [
    ("No.", 6, None),
    ("File Name", 50, "file"),
    ("Drawing No. (in Drawing)", 42, "dwg_number"),
    ("Drawing Title", 50, "title"),
    ("File Rev.", 8, "file_rev"),
    ("Drawing Rev.", 9, "rev"),
    ("Date", 11, "date"),
    ("Check", 12, "status"),
    ("Note", 60, "note"),
]
# Labels of the number cell in a drawing title block, and the kind of drawing each one means.
NUMBER_LABELS = (("Serial Number", "drawing"), ("Serial No", "drawing"),
                 ("Drawing Number", "shop_drawing"), ("Drawing No", "shop_drawing"))


# ----------------------------------------------------------------------------
# Reading words and labels from a page
# ----------------------------------------------------------------------------

class Word:
    __slots__ = ("x0", "y0", "x1", "y1", "text", "key")

    def __init__(self, x0, y0, x1, y1, text):
        self.x0, self.y0, self.x1, self.y1 = x0, y0, x1, y1
        self.text = text
        self.key = norm(text)

    @property
    def w(self):
        return self.x1 - self.x0

    @property
    def h(self):
        return self.y1 - self.y0

    @property
    def cy(self):
        return (self.y0 + self.y1) / 2


def norm(text):
    return re.sub(r"[^a-z0-9]", "", text.lower())


def page_words(raw_words, page, extra_rotation):
    """Word boxes as the page is seen on screen, optionally turned further."""
    matrix = page.rotation_matrix * fitz.Matrix(extra_rotation)
    words = []
    for x0, y0, x1, y1, text, *_ in raw_words:
        r = fitz.Rect(x0, y0, x1, y1) * matrix
        words.append(Word(r.x0, r.y0, r.x1, r.y1, text))
    return words


def next_word_on_line(words, word):
    best = None
    for w in words:
        if w is word or abs(w.cy - word.cy) > word.h * 0.4 or w.x0 <= word.x0:
            continue
        if -word.h * 0.5 <= w.x0 - word.x1 <= word.h * 1.5 and (best is None or w.x0 < best.x0):
            best = w
    return best


def find_labels(words, label):
    """Every place `label` is written, also when the PDF splits it into pieces
    ("Sub" + "ject:"). Returns one merged box per hit, top-most first."""
    target = norm(label)
    found = []
    for start in words:
        if not start.key or not target.startswith(start.key):
            continue
        if len(start.key) >= 3 and start.w < start.h * 0.9:     # standing text, wrong way round
            continue
        parts, acc = [start], start.key
        while acc != target:
            nxt = next_word_on_line(words, parts[-1])
            if nxt is None or not target.startswith(acc + nxt.key):
                break
            parts.append(nxt)
            acc += nxt.key
        if acc == target:
            found.append(merge(parts))
    found.sort(key=lambda w: (w.y0, w.x0))
    return found


def join_lines(words):
    """Words -> one line of text, read top to bottom, left to right."""
    lines = []
    for w in sorted(words, key=lambda w: w.cy):
        if lines and abs(w.cy - lines[-1][0].cy) <= min(w.h, lines[-1][0].h) * 0.5:
            lines[-1].append(w)
        else:
            lines.append([w])
    text = " ".join(" ".join(w.text for w in sorted(line, key=lambda w: w.x0)) for line in lines)
    return re.sub(r"\s+", " ", text).strip()


def text_right_of(words, label):
    """First word to the right of a label on the same row (e.g. "Rev.: 00")."""
    row = [w for w in words if abs(w.cy - label.cy) <= label.h * 0.8
           and label.x1 - 1 <= w.x0 <= label.x1 + label.h * 12]
    return min(row, key=lambda w: w.x0).text if row else ""


def merge(parts):
    return Word(min(p.x0 for p in parts), min(p.y0 for p in parts),
                max(p.x1 for p in parts), max(p.y1 for p in parts), " ".join(p.text for p in parts))


def line_below(words, box, x0, x1, reach):
    """Words of the first line under `box` that overlaps the column x0..x1."""
    under = [w for w in words if w.y0 >= box.y1 - 1 and w.cy <= box.y1 + reach
             and w.x1 > x0 - 2 and w.x0 < x1 + 2]
    if not under:
        return []
    top = min(under, key=lambda w: w.cy)
    parts = [min((w for w in under if abs(w.cy - top.cy) <= top.h * 0.5), key=lambda w: w.x0)]
    while True:
        nxt = next_word_on_line(words, parts[-1])
        if nxt is None:
            return parts
        parts.append(nxt)


def text_below(words, label):
    """Text in the title-block cell under a label. A number that wraps onto the next
    line ("SAAD-01-DC103-MIS-SDW-" / "ST-2LG-5B-000021") is joined back together."""
    parts = line_below(words, label, label.x0, label.x1, label.h * 6)
    text = " ".join(p.text for p in parts)
    while parts and text.endswith("-"):
        line = merge(parts)
        parts = line_below(words, line, label.x0, label.x1, line.h * 1.5)
        text += " ".join(p.text for p in parts)
    return re.sub(r"-\s+", "-", text)


# ----------------------------------------------------------------------------
# Documents and drawings
# ----------------------------------------------------------------------------

def read_document(words):
    info = {"kind": "document", "doc_no": "", "subject": "", "rev": "", "date": ""}

    numbers = [(w, m.group(0)) for w in words for m in [DOC_NO_RE.search(w.text)] if m]
    if numbers:
        info["doc_no"] = min(numbers, key=lambda n: (n[0].y0, n[0].x0))[1]

    subjects = find_labels(words, "Subject")
    if subjects:
        label = subjects[0]
        # The subject cell runs from the row label above "Subject:" to the one below it.
        column = [w for w in words if abs(w.x0 - label.x0) <= 12 and not (label.x0 <= w.x0 < label.x1 and label.y0 <= w.cy <= label.y1)]
        above = [w.y1 for w in column if w.cy < label.y0]
        below = [w.y0 for w in column if w.cy > label.y1]
        top = max(above) if above else label.y0 - label.h
        bottom = min(below) if below else label.y1 + label.h
        info["subject"] = join_lines([w for w in words if w.x0 >= label.x1 - 1 and top <= w.cy <= bottom])

    revs = find_labels(words, "Rev")
    if revs:
        info["rev"] = text_right_of(words, revs[0])
    dates = find_labels(words, "Submittal Date")
    if dates:
        info["date"] = text_right_of(words, dates[0])
    return info


def find_title_block(words):
    """The "Drawing Title" label and the number label under it (Serial Number or
    DRAWING NUMBER) with the kind of drawing, or None. If a label also appears
    elsewhere on the sheet, the pair closest together is the title block."""
    best = None
    for title in find_labels(words, "Drawing Title"):
        for name, kind in NUMBER_LABELS:
            for number in find_labels(words, name):
                gap = number.y0 - title.y1
                if 0 < gap < title.h * 30 and number.x0 >= title.x0 - title.h * 2 and (best is None or gap < best[0]):
                    best = (gap, title, number, kind)
    return best[1:] if best else None


def read_drawing(words, title_label, number_label, kind):
    info = {"kind": kind, "title": "", "number": "", "rev": "", "date": ""}
    block = [w for w in words if w.x0 >= title_label.x0 - title_label.h * 2 and w.y0 > title_label.y1 - 1]

    # The title cell ends where the next row of labels (Project Code, Serial Number ...) starts.
    stops = [number_label.y0] + [lab.y0 for name in ("Project Code", "Building Code", "Discipline")
                                 for lab in find_labels(block, name) if lab.y0 > title_label.y1]
    info["title"] = join_lines([w for w in block if w.y1 <= min(stops) + 1])
    info["number"] = text_below(words, number_label)

    rev_label = None
    for name in ("Rev No", "Revision", "Rev"):
        labels = find_labels(block, name)
        if labels:
            rev_label = labels[-1]
            info["rev"] = text_below(words, rev_label)
            break
    if rev_label:
        dates = [d for d in find_labels(block, "Date") if abs(d.cy - rev_label.cy) <= rev_label.h]
        if dates:
            info["date"] = text_below(words, min(dates, key=lambda d: abs(d.x0 - rev_label.x0)))
    return info



def _normalise_pdf_id_text(text):
    """Normalise PDF soft hyphens and typographic dashes in SAAD document IDs."""
    text = str(text or "")
    for dash in ("\u00ad", "\u2010", "\u2011", "\u2012", "\u2013", "\u2014", "\u2212"):
        text = text.replace(dash, "-")
    return re.sub(r"(?<=\w)\s*-\s*(?=\w)", "-", text)


def _wir_page_lines(page):
    """Return text lines with page coordinates for WIR description-field extraction."""
    grouped = {}
    for word in page.get_text("words"):
        if len(word) < 8:
            continue
        key = (word[5], word[6])
        grouped.setdefault(key, []).append(word)
    lines = []
    for values in grouped.values():
        values.sort(key=lambda item: item[0])
        lines.append({
            "y0": min(item[1] for item in values),
            "y1": max(item[3] for item in values),
            "x0": min(item[0] for item in values),
            "text": " ".join(str(item[4]) for item in values).strip(),
        })
    return sorted(lines, key=lambda item: (item["y0"], item["x0"]))


def _wir_description_from_pages(pages):
    """Read the free-text DESCRIPTION value printed above the DESCRIPTION label."""
    excluded = (
        "CHECKLIST FOR", "SAAD", "HDC-HTC", "SITE ENGINEER", "HATCO QAQC",
        "CLIENT QAQC", "CONSULTANT", "EDARAT", "SIGNATURE", "APPROVAL",
        "DATE:", "NAME:", "CH NO", "CHK NO", "S/N", "DESCRIPTION",
        "LOCATION OF WORK", "CONTRACTOR", "PROJECT", "HATCO", "WSP",
        "CONFORMED", "NOT", "N/A",
    )
    for page in pages[:3]:
        lines = _wir_page_lines(page)
        labels = [line for line in lines if norm(line["text"].rstrip(":")) == "description"]
        for label in labels:
            candidates = []
            for line in lines:
                distance = label["y0"] - line["y1"]
                text = re.sub(r"\s+", " ", line["text"]).strip(" :")
                upper = text.upper()
                if not (4 <= distance <= 95) or len(text) < 5:
                    continue
                if any(term in upper for term in excluded):
                    continue
                if re.fullmatch(r"[\W\d_]+", text):
                    continue
                if re.search(r"^PAGE\s+\d+\s+OF\s+\d+$", text, re.I):
                    continue
                candidates.append((distance, text))
            if candidates:
                candidates.sort(key=lambda item: item[0])
                return candidates[0][1]
    return ""


def _wir_scope_and_checklist(pages):
    """Build a meaningful fallback subject from the ITP scope and WIR checklist title."""
    all_text = "\n".join(page.get_text("text") for page in pages[:3])
    lines = [re.sub(r"\s+", " ", line).strip() for line in all_text.splitlines()]
    scope = ""
    for index, line in enumerate(lines):
        match = re.search(r"SCOPE\s+OF\s+WORK\s*:?\s*(.*)$", line, re.I)
        if not match:
            continue
        scope = match.group(1).strip(" :-")
        if not scope:
            for following in lines[index + 1:index + 5]:
                candidate = following.strip(" :-")
                if candidate and not re.search(r"^(LEGEND|ITP\s+NO|REVISION\s+NO|H\s*[–-]\s*HOLD)", candidate, re.I):
                    scope = candidate
                    break
        if scope:
            break

    checklist = ""
    for line in lines:
        match = re.search(r"\bCHECKLIST\s+FOR\s+(.+)$", line, re.I)
        if match:
            checklist = match.group(1).strip(" :-")
            break

    if scope and checklist:
        return f"{scope} - {checklist}"
    return scope or checklist


def read_wir_pdf(path):
    """WIR packs use a checklist/ITP cover, not the normal Subject title-block layout."""
    stem = os.path.splitext(os.path.basename(path))[0]
    with fitz.open(path) as pdf:
        if pdf.needs_pass:
            return {"error": "PDF is password protected - check manually."}

        pages = [pdf[index] for index in range(min(3, pdf.page_count))]
        page_texts = [page.get_text("text") for page in pages]
        all_text = _normalise_pdf_id_text("\n".join(page_texts))

        printed_numbers = WIR_DOC_NO_RE.findall(all_text)
        doc_no = printed_numbers[0].upper() if printed_numbers else ""
        if not doc_no and re.search(r"-WIR-", stem, re.I):
            doc_no = stem.upper()

        subject = _wir_description_from_pages(pages)
        if not subject:
            subject = _wir_scope_and_checklist(pages)

        # Some WIR PDFs have blank cover pages and attach the originating
        # drawing or test report later in the pack. Search later pages only
        # when the cover did not provide a subject.
        if not subject:
            for page_index in range(3, pdf.page_count):
                later_page = pdf[page_index]
                later_text = later_page.get_text("text")
                letter_match = re.search(
                    r"\bSUBJ(?:ECT)?\s*:\s*(.+?)(?=\n\s*Dear\s+(?:Sir|Madam)\b|\Z)",
                    later_text,
                    re.I | re.S,
                )
                if letter_match:
                    subject = re.sub(r"\s+", " ", letter_match.group(1)).strip(" :-")
                if not subject:
                    raw = later_page.get_text("words")
                    for rotation in EXTRA_ROTATIONS:
                        later_words = page_words(raw, later_page, rotation)
                        title_block = find_title_block(later_words)
                        if title_block:
                            drawing_info = read_drawing(later_words, *title_block)
                            subject = drawing_info.get("title", "").strip()
                            if subject:
                                break
                if subject:
                    break

        rev_match = re.search(r"\bRevision\s+No\.?\s*:\s*([A-Z0-9]+)", all_text, re.I)
        if not rev_match:
            rev_match = re.search(r"\bRev\.?\s*([A-Z0-9]+)", all_text, re.I)
        revision = rev_match.group(1).upper() if rev_match else ""

        return {
            "kind": "document",
            "doc_no": doc_no,
            "subject": subject,
            "rev": revision,
            "date": "",
            "doc_no_source": "PDF" if printed_numbers else ("filename" if doc_no else ""),
        }


def read_pdf(path):
    name = os.path.basename(path).upper()
    if "-WIR-" in name:
        return read_wir_pdf(path)

    any_text = False
    with fitz.open(path) as pdf:
        if pdf.needs_pass:
            return {"error": "PDF is password protected - check manually."}
        for page in pdf.pages(0, min(PAGES_TO_SEARCH, pdf.page_count)):
            raw = page.get_text("words")
            if not raw:
                continue
            any_text = True
            drawing = doc_words = None
            for rot in EXTRA_ROTATIONS:
                words = page_words(raw, page, rot)
                if drawing is None:
                    block = find_title_block(words)
                    if block:
                        drawing = (words, *block)
                if doc_words is None and find_labels(words, "Subject"):
                    doc_words = words
            if drawing and doc_words:
                return read_drawing(*drawing) if re.search(r"-(DWG|SDW)-", name) else read_document(doc_words)
            if drawing:
                return read_drawing(*drawing)
            if doc_words:
                return read_document(doc_words)
            words = page_words(raw, page, 0)
            if any(DOC_NO_RE.search(w.text) for w in words):
                return read_document(words)
    if not any_text:
        return {"error": "No readable text in the PDF (scanned image?) - check manually."}
    return {}


def _entity_text(entity):
    """Return plain visible text from common DXF text entities."""
    kind = entity.dxftype()
    try:
        if kind == "MTEXT":
            return entity.plain_text()
        if kind in ("TEXT", "ATTRIB", "ATTDEF"):
            return entity.dxf.text
    except Exception:
        return ""
    return ""


def _entity_anchor(entity):
    """Return insertion point and text height for a text-like DXF entity."""
    kind = entity.dxftype()
    try:
        point = entity.dxf.insert
        height = float(entity.dxf.char_height if kind == "MTEXT" else entity.dxf.height)
        return float(point.x), float(point.y), max(abs(height), 1.0)
    except Exception:
        return None


def _collect_text_entities(layout, depth=0):
    """Collect text from layouts and expanded INSERT blocks, including attributes."""
    if depth > 5:
        return
    for entity in layout:
        kind = entity.dxftype()
        if kind in ("TEXT", "MTEXT", "ATTRIB", "ATTDEF"):
            yield entity
        elif kind == "INSERT":
            # Attribute tags are useful when a title block stores fields as block attributes.
            for attrib in getattr(entity, "attribs", []):
                yield attrib
            try:
                for virtual in entity.virtual_entities():
                    if virtual.dxftype() in ("TEXT", "MTEXT", "ATTRIB", "ATTDEF"):
                        yield virtual
                    elif virtual.dxftype() == "INSERT":
                        class _OneEntityLayout:
                            def __iter__(self_nonlocal):
                                yield virtual
                        yield from _collect_text_entities(_OneEntityLayout(), depth + 1)
            except Exception:
                # Some proxy/custom blocks cannot be exploded by ezdxf; continue with other text.
                pass


def _attribute_values(doc):
    """Read semantic title-block attributes, even when their tags are not visible text."""
    values = {}
    for layout in doc.layouts:
        for entity in layout:
            if entity.dxftype() != "INSERT":
                continue
            for attrib in getattr(entity, "attribs", []):
                try:
                    tag = norm(str(attrib.dxf.tag))
                    value = str(attrib.dxf.text).strip()
                    if tag and value:
                        values[tag] = value
                except Exception:
                    pass
    return values


def _first_attr(values, *names):
    for name in names:
        value = values.get(norm(name), "")
        if value:
            return value
    return ""


def _words_from_dxf(doc):
    """Convert CAD text objects into approximate word boxes for the existing title-block parser."""
    words = []
    seen = set()
    for layout in doc.layouts:
        for entity in _collect_text_entities(layout):
            raw = _entity_text(entity)
            anchor = _entity_anchor(entity)
            if not raw or not anchor:
                continue
            x, y, height = anchor
            # CAD Y grows upwards; the existing parser expects top-to-bottom coordinates.
            y0 = -y - height
            clean = re.sub(r"\\P", " ", str(raw))
            clean = re.sub(r"\{\\[^;]*;", "", clean).replace("}", "")
            clean = re.sub(r"\\[A-Za-z][^;]*;", "", clean)
            tokens = re.findall(r"SAAD-[A-Za-z0-9]+(?:[-_.][A-Za-z0-9]+)*|[A-Za-z0-9_./:-]+", clean)
            if not tokens:
                continue
            width = max(height * 0.55, height * len(clean) * 0.48)
            char_width = width / max(len(clean), 1)
            offset = 0
            for token in tokens:
                pos = clean.find(token, offset)
                if pos < 0:
                    pos = offset
                tx0 = x + pos * char_width
                tx1 = tx0 + max(len(token) * char_width, height * 0.35)
                key = (round(tx0, 3), round(y0, 3), token)
                if key not in seen:
                    words.append(Word(tx0, y0, tx1, y0 + height, token))
                    seen.add(key)
                offset = pos + len(token)
    return words



def _bundled_converter_candidates():
    """Return likely locations for the converter embedded by PyInstaller."""
    roots = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        roots.append(meipass)
    roots.append(os.path.dirname(os.path.abspath(__file__)))
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    if exe_dir not in roots:
        roots.append(exe_dir)

    candidates = []
    for root in roots:
        candidates.extend([
            os.path.join(root, "libredwg", "dwg2dxf.exe"),
            os.path.join(root, "dwg2dxf.exe"),
            os.path.join(root, "libredwg", "bin", "dwg2dxf.exe"),
            os.path.join(root, "bin", "dwg2dxf.exe"),
        ])
    # Remove duplicates but keep search order.
    return list(dict.fromkeys(os.path.normpath(p) for p in candidates))


def _bundled_converter_path():
    """Find only the converter shipped inside this application; no ODA lookup."""
    for candidate in _bundled_converter_candidates():
        if os.path.isfile(candidate):
            return candidate
    return None


def _hidden_subprocess_options():
    """Prevent child console programs from opening a command window on Windows."""
    if os.name != "nt":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    return {
        "startupinfo": startupinfo,
        "creationflags": subprocess.CREATE_NO_WINDOW,
    }


def _check_bundled_converter():
    """Return (path, version output); fail with exact embedded search paths."""
    converter = _bundled_converter_path()
    if not converter:
        searched = "\n".join(_bundled_converter_candidates())
        raise RuntimeError(
            "The LibreDWG converter is not embedded in this EXE. Searched:\n" + searched
        )
    proc = subprocess.run(
        [converter, "--version"],
        capture_output=True,
        text=True,
        timeout=30,
        cwd=os.path.dirname(converter),
        **_hidden_subprocess_options(),
    )
    output = (proc.stdout or proc.stderr or "").strip()
    if proc.returncode != 0:
        raise RuntimeError(
            f"The embedded LibreDWG converter was found but could not start "
            f"(exit {proc.returncode}): {output[-500:]}"
        )
    return converter, output


def _run_libredwg_batch(converter, paths, progress=None, cancel_event=None):
    """Run LibreDWG once per batch, then read each produced DXF."""
    import tempfile
    from collections import defaultdict

    results = {}
    by_stem = defaultdict(list)
    for path in paths:
        stem = os.path.splitext(os.path.basename(path))[0].casefold()
        by_stem[stem].append(path)

    unique_paths = [group[0] for group in by_stem.values() if len(group) == 1]
    duplicate_groups = [group for group in by_stem.values() if len(group) > 1]

    batches = []
    current = []
    current_chars = 0
    for path in unique_paths:
        addition = len(path) + 4
        if current and (len(current) >= 20 or current_chars + addition > 18000):
            batches.append(current)
            current, current_chars = [], 0
        current.append(path)
        current_chars += addition
    if current:
        batches.append(current)

    completed = 0
    launch_count = len(batches) + sum(len(group) for group in duplicate_groups)

    with tempfile.TemporaryDirectory(prefix="pdf_cad_register_") as work_root:
        def convert_group(group, output_dir):
            command = [converter, "-y", *group]
            try:
                proc = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    timeout=max(180, 90 * len(group)),
                    cwd=output_dir,
                    **_hidden_subprocess_options(),
                )
                diagnostic = (proc.stderr or proc.stdout or "").strip()[-900:]
                if proc.returncode and not diagnostic:
                    diagnostic = f"LibreDWG exited with code {proc.returncode}."
            except Exception as exc:
                diagnostic = f"{type(exc).__name__}: {exc}"

            for original in group:
                expected_name = os.path.splitext(os.path.basename(original))[0] + ".dxf"
                try:
                    actual = next(
                        name for name in os.listdir(output_dir)
                        if name.casefold() == expected_name.casefold()
                    )
                    dxf_path = os.path.join(output_dir, actual)
                    results[original] = (ezdxf.readfile(dxf_path), "")
                except StopIteration:
                    results[original] = (
                        None,
                        diagnostic or "LibreDWG did not create the expected DXF output.",
                    )
                except Exception as exc:
                    results[original] = (
                        None,
                        f"Could not parse generated DXF: {type(exc).__name__}: {str(exc)[:300]}",
                    )

        for batch_no, group in enumerate(batches, 1):
            if cancel_event and cancel_event.is_set():
                raise Cancelled
            if progress:
                progress(completed, len(paths),
                         f"Converting batch {batch_no}/{len(batches)} ({len(group)} DWGs)")
            convert_group(group, work_root)
            completed += len(group)
            if progress:
                progress(completed, len(paths), f"Converted {completed}/{len(paths)} DWGs")

        # Same-named DWGs in different directories must not overwrite one another.
        for dup_no, group in enumerate(duplicate_groups, 1):
            for original in group:
                if cancel_event and cancel_event.is_set():
                    raise Cancelled
                if progress:
                    progress(completed, len(paths),
                             f"Converting duplicate-name drawing {completed + 1}/{len(paths)}")
                output_dir = os.path.join(work_root, f"duplicate_{dup_no}_{completed}")
                os.makedirs(output_dir, exist_ok=True)
                convert_group([original], output_dir)
                completed += 1
                if progress:
                    progress(completed, len(paths), f"Converted {completed}/{len(paths)} DWGs")

    return results


def _open_dwg_as_dxf(path):
    """Convert a single DWG using the LibreDWG executable bundled in this EXE."""
    import tempfile
    converter = _bundled_converter_path()
    if not converter:
        searched = "; ".join(_bundled_converter_candidates())
        raise RuntimeError("Bundled LibreDWG converter missing. Searched: " + searched)

    with tempfile.TemporaryDirectory(prefix="pdf_cad_register_single_") as output_dir:
        proc = subprocess.run(
            [converter, "-y", path],
            capture_output=True,
            text=True,
            timeout=180,
            cwd=output_dir,
            **_hidden_subprocess_options(),
        )
        expected = os.path.splitext(os.path.basename(path))[0] + ".dxf"
        dxf_path = os.path.join(output_dir, expected)
        if proc.returncode != 0 or not os.path.isfile(dxf_path):
            detail = (proc.stderr or proc.stdout or "").strip()[-800:]
            raise RuntimeError(detail or f"LibreDWG failed with exit code {proc.returncode}.")
        return ezdxf.readfile(dxf_path)


def read_dwg(path, converted_doc=None, conversion_error=None):
    name = os.path.basename(path).upper()
    if conversion_error:
        return {"error": f"LibreDWG conversion failed: {str(conversion_error)[:350]}"}
    try:
        doc = converted_doc if converted_doc is not None else _open_dwg_as_dxf(path)
        words = _words_from_dxf(doc)
    except Exception as exc:
        return {"error": f"Could not read this DWG using the bundled LibreDWG converter: {type(exc).__name__}: {str(exc)[:350]}"}
    if not words:
        return {"error": "No readable text entities found in this DWG. Text may be embedded in an unsupported proxy object, raster image, or external reference."}

    block = find_title_block(words)
    has_subject = bool(find_labels(words, "Subject"))
    attrs = _attribute_values(doc)
    is_shop = "-SDW-" in name
    is_drawing = "-DWG-" in name or bool(_first_attr(attrs, "Drawing Title", "Dwg Title", "Serial Number", "Serial No"))

    if is_shop or (block and block[2] == "shop_drawing"):
        info = read_drawing(words, *block) if block else {"kind": "shop_drawing", "title": "", "number": "", "rev": "", "date": ""}
        info["title"] = info.get("title") or _first_attr(attrs, "Drawing Title", "Dwg Title", "Title")
        info["number"] = info.get("number") or _first_attr(attrs, "Drawing Number", "Drawing No", "Dwg Number")
        info["rev"] = info.get("rev") or _first_attr(attrs, "Revision", "Rev No", "Rev")
        info["date"] = info.get("date") or _first_attr(attrs, "Date", "Drawing Date")
        return info

    if is_drawing or block:
        info = read_drawing(words, *block) if block else {"kind": "drawing", "title": "", "number": "", "rev": "", "date": ""}
        info["title"] = info.get("title") or _first_attr(attrs, "Drawing Title", "Dwg Title", "Title")
        info["number"] = info.get("number") or _first_attr(attrs, "Serial Number", "Serial No", "Drawing Number", "Drawing No")
        info["rev"] = info.get("rev") or _first_attr(attrs, "Revision", "Rev No", "Rev")
        info["date"] = info.get("date") or _first_attr(attrs, "Date", "Drawing Date")
        return info

    info = read_document(words) if (has_subject or any(DOC_NO_RE.search(w.text) for w in words)) else {"kind": "document", "doc_no": "", "subject": "", "rev": "", "date": ""}
    info["doc_no"] = info.get("doc_no") or _first_attr(attrs, "Document Number", "Document No", "Doc No", "Doc Number")
    info["subject"] = info.get("subject") or _first_attr(attrs, "Subject", "Document Subject", "Document Title")
    info["rev"] = info.get("rev") or _first_attr(attrs, "Revision", "Rev No", "Rev")
    info["date"] = info.get("date") or _first_attr(attrs, "Submittal Date", "Date")
    return info


def serial_from_file_name(stem):
    """SAAD-01-DC103-MIS-DWG-ST-2L0-5A-ST121_00 -> ST121"""
    last = stem.rsplit("-", 1)[-1].strip()
    m = re.match(r"[A-Za-z0-9]+", last)
    return m.group(0) if m else last


def same(a, b):
    return a.strip().upper() == b.strip().upper()


def revision_from_file_name(stem):
    """Return the suffix after the final underscore, e.g. drawing_01 -> 01."""
    match = re.search(r"_([^_]+)$", stem)
    return match.group(1).strip() if match else ""


def document_row(rel, info):
    stem = os.path.splitext(os.path.basename(rel))[0]
    row = {"file": rel, "doc_no": info.get("doc_no", ""), "subject": info.get("subject", ""),
           "rev": info.get("rev", ""), "date": info.get("date", "")}
    notes, mismatch = [], False
    if info.get("error"):
        notes.append(info["error"])
    elif not row["doc_no"]:
        notes.append("Document number starting with SAAD- not found in the file.")
    elif not same(row["doc_no"], stem):
        mismatch = True
        notes.append(f'File name "{stem}" does not match document number "{row["doc_no"]}" in the file.')
    if not info.get("error") and not row["subject"]:
        notes.append("Subject not found.")
    row["status"] = "MISMATCH" if mismatch else ("CHECK" if notes else "OK")
    row["note"] = " ".join(notes)
    return row


def drawing_row(rel, info):
    stem = os.path.splitext(os.path.basename(rel))[0]
    row = {"file": rel, "file_serial": serial_from_file_name(stem), "dwg_serial": info.get("number", ""),
           "title": info.get("title", ""), "file_rev": revision_from_file_name(stem),
           "rev": info.get("rev", ""), "date": info.get("date", "")}
    notes, mismatch = [], False
    if info.get("error"):
        notes.append(info["error"])
    elif not row["dwg_serial"]:
        notes.append("Serial Number not found in the drawing title block.")
    elif not same(row["dwg_serial"], row["file_serial"]):
        mismatch = True
        notes.append(f'Serial number in file name "{row["file_serial"]}" does not match '
                     f'serial number "{row["dwg_serial"]}" in the drawing.')
    if not info.get("error") and row["file_rev"] and row["rev"] and not same(row["file_rev"], row["rev"]):
        mismatch = True
        notes.append(f'Revision in file name "{row["file_rev"]}" does not match revision "{row["rev"]}" in the drawing.')
    if not info.get("error") and not row["title"]:
        notes.append("Drawing Title not found.")
    row["status"] = "MISMATCH" if mismatch else ("CHECK" if notes else "OK")
    row["note"] = " ".join(notes)
    return row


def shop_drawing_row(rel, info):
    stem = os.path.splitext(os.path.basename(rel))[0]
    file_number = re.sub(r"_[^_]*$", "", stem)      # drop the revision: ..._00
    row = {"file": rel, "dwg_number": info.get("number", ""),
           "title": info.get("title", ""), "file_rev": revision_from_file_name(stem),
           "rev": info.get("rev", ""), "date": info.get("date", "")}
    notes, mismatch = [], False
    if info.get("error"):
        notes.append(info["error"])
    elif not row["dwg_number"]:
        notes.append("DRAWING NUMBER not found in the drawing title block.")
    elif not same(row["dwg_number"], file_number):
        mismatch = True
        notes.append(f'File name "{file_number}" does not match drawing number "{row["dwg_number"]}" '
                     f'in the drawing.')
    if not info.get("error") and row["file_rev"] and row["rev"] and not same(row["file_rev"], row["rev"]):
        mismatch = True
        notes.append(f'Revision in file name "{row["file_rev"]}" does not match revision "{row["rev"]}" in the drawing.')
    if not info.get("error") and not row["title"]:
        notes.append("Drawing Title not found.")
    row["status"] = "MISMATCH" if mismatch else ("CHECK" if notes else "OK")
    row["note"] = " ".join(notes)
    return row


SECTIONS = [
    # kind, sheet and tile name, Excel columns, builds a row, cells shown red on a mismatch
    ("document", "Documents", DOC_COLUMNS, document_row, ("file", "doc_no")),
    ("drawing", "Drawings", DWG_COLUMNS, drawing_row, ("file_serial", "dwg_serial")),
    ("shop_drawing", "Shop Drawings", SDW_COLUMNS, shop_drawing_row, ("file", "dwg_number")),
]


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------

def write_excel(path, folder, results):
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    status_style = {
        "OK": (None, Font(bold=True, color="006100")),
        "MISMATCH": (PatternFill("solid", fgColor="FFC7CE"), Font(bold=True, color="9C0006")),
        "CHECK": (PatternFill("solid", fgColor="FFEB9C"), Font(bold=True, color="9C5700")),
    }

    def fill_sheet(ws, title, columns, rows, compared):
        ws.title = title
        for c, (header, width, _) in enumerate(columns, 1):
            cell = ws.cell(row=1, column=c, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
            ws.column_dimensions[cell.column_letter].width = width
        ws.row_dimensions[1].height = 32
        for r, row in enumerate(rows, 2):
            fill, status_font = status_style[row["status"]]
            for c, (_, _, key) in enumerate(columns, 1):
                cell = ws.cell(row=r, column=c, value=r - 1 if key is None else row[key])
                cell.border = border
                cell.alignment = Alignment(vertical="top", wrap_text=key in ("subject", "title", "note", "file"),
                                           horizontal="center" if key in (None, "rev", "date", "status") else None)
                if fill:
                    cell.fill = fill
                if key == "status":
                    cell.font = status_font
                elif row["status"] == "MISMATCH" and key in compared:
                    cell.font = Font(bold=True, color="9C0006")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=len(columns)).column_letter}{max(len(rows) + 1, 2)}"

    wb = openpyxl.Workbook()
    for i, (kind, title, columns, _, compared) in enumerate(SECTIONS):
        fill_sheet(wb.active if i == 0 else wb.create_sheet(), title, columns, results[kind], compared)

    ws = wb.create_sheet("Summary")
    ws["A1"] = "PDF & CAD Register"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A3"], ws["B3"] = "Folder", folder
    ws["A4"], ws["B4"] = "Created", datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    for c, header in enumerate(("", "Files", "OK", "Mismatch", "Check"), 1):
        cell = ws.cell(row=6, column=c, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.border = border
    for r, (kind, title, *_) in enumerate(SECTIONS, 7):
        rows = results[kind]
        values = [title, len(rows)] + [sum(x["status"] == s for x in rows) for s in ("OK", "MISMATCH", "CHECK")]
        for c, value in enumerate(values, 1):
            ws.cell(row=r, column=c, value=value).border = border
    ws.cell(row=8 + len(SECTIONS), column=1,
            value="MISMATCH (red row): the file name and the number inside the PDF or DWG are not the same.")
    ws.cell(row=9 + len(SECTIONS), column=1,
            value="CHECK (yellow row): something could not be read from the PDF or DWG - see the Note column.")
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 12
    for col in "CDE":
        ws.column_dimensions[col].width = 11

    wb.save(path)


# ----------------------------------------------------------------------------
# Running
# ----------------------------------------------------------------------------

def natural_key(text):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", text)]


def find_input_files(folder, recursive=True):
    """Find PDF and DWG files, optionally including sub-folders."""
    if recursive:
        paths = [os.path.join(root, f) for root, _, files in os.walk(folder) for f in files]
    else:
        paths = [os.path.join(folder, f) for f in os.listdir(folder)]
    paths = [p for p in paths if p.lower().endswith((".pdf", ".dwg")) and os.path.isfile(p)]
    return sorted(paths, key=lambda p: natural_key(os.path.relpath(p, folder)))



def make_register(folder, progress=None, recursive=True, cancel_event=None):
    """Batch-convert DWGs, then parse all files and write the Excel register."""
    paths = find_input_files(folder, recursive)
    results = {kind: [] for kind, *_ in SECTIONS}
    row_builders = {kind: build for kind, _, _, build, _ in SECTIONS}

    dwg_paths = [path for path in paths if path.lower().endswith(".dwg")]
    converted = {}
    if dwg_paths:
        converter = _bundled_converter_path()
        if converter:
            converted = _run_libredwg_batch(
                converter,
                dwg_paths,
                progress=progress,
                cancel_event=cancel_event,
            )
        else:
            searched = "; ".join(_bundled_converter_candidates())
            message = "Bundled LibreDWG converter not found. Searched: " + searched
            converted = {path: (None, message) for path in dwg_paths}

    if progress and dwg_paths:
        progress(len(dwg_paths), len(dwg_paths) + len(paths), "DWG conversion finished; reading files")

    for index, path in enumerate(paths, 1):
        if cancel_event and cancel_event.is_set():
            raise Cancelled
        rel = os.path.relpath(path, folder)
        if progress:
            progress(len(dwg_paths) + index, len(dwg_paths) + len(paths), rel)
        try:
            if path.lower().endswith(".pdf"):
                info = read_pdf(path)
            else:
                doc, error = converted.get(path, (None, "No conversion result was returned."))
                info = read_dwg(path, converted_doc=doc, conversion_error=error)
        except Exception as exc:
            ext = os.path.splitext(path)[1].upper()
            info = {"error": f"Could not open {ext} file: {type(exc).__name__}: {str(exc)[:300]}"}

        name = os.path.basename(path).upper()
        kind = info.get("kind") or (
            "drawing" if "-DWG-" in name else
            "shop_drawing" if "-SDW-" in name else
            "document"
        )
        results[kind].append(row_builders[kind](rel, info))

    out = os.path.join(
        folder,
        f"PDF CAD Register {datetime.datetime.now():%Y-%m-%d %H%M%S}.xlsx",
    )
    write_excel(out, folder, results)
    return out, results


def summary_text(out, results):
    lines = []
    for kind, title, *_ in SECTIONS:
        rows = results[kind]
        mism = sum(r["status"] == "MISMATCH" for r in rows)
        check = sum(r["status"] == "CHECK" for r in rows)
        lines.append(f"{title + ':':15} {len(rows)} files, {mism} mismatch, {check} to check")
    return "\n".join(lines) + f"\n\nExcel file saved:\n{out}"


# ----------------------------------------------------------------------------
# Window
# ----------------------------------------------------------------------------

APP_NAME = "PDF & CAD Register"
COLORS = {"OK": "#107C10", "MISMATCH": "#C42B1C", "CHECK": "#9D5D00", "muted": "#616161",
          "MISMATCH_bg": "#FDE7E9", "CHECK_bg": "#FFF4CE"}
SETTINGS_FILE = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), APP_NAME, "settings.json")


class Cancelled(Exception):
    pass


def resource(name):
    """Files packed into the exe are unpacked to a temp folder (sys._MEIPASS)."""
    return os.path.join(getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))), name)


def load_settings():
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(settings):
    try:
        os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f)
    except OSError:
        pass


class App:
    def __init__(self, root):
        self.root = root
        settings = load_settings()
        self.folder = tk.StringVar(value=settings.get("folder", ""))
        self.recursive = tk.BooleanVar(value=settings.get("recursive", True))
        self.events = queue.Queue()
        self.stop = threading.Event()
        self.running = self.closing = False
        self.output = None
        self.file_paths = {}         # review-list row -> full path of the source file
        self.count_job = None

        self.build()
        self.folder.trace_add("write", lambda *_: self.schedule_count())
        self.recursive.trace_add("write", lambda *_: self.schedule_count())
        self.update_count()
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    # -- layout --------------------------------------------------------------

    def build(self):
        root = self.root
        root.title(APP_NAME)
        try:
            root.iconbitmap(default=resource("app.ico"))
        except tk.TclError:
            pass
        title_font = font.Font(family="Segoe UI Semibold", size=18)
        self.head_font = font.Font(family="Segoe UI Semibold", size=11)
        self.big_font = font.Font(family="Segoe UI Semibold", size=22)
        body = font.nametofont("TkDefaultFont")
        ttk.Style().configure("Treeview", rowheight=int(body.metrics("linespace") * 1.6))

        page = ttk.Frame(root, padding=(24, 18, 24, 20))
        page.pack(fill="both", expand=True)
        page.columnconfigure(0, weight=1)

        ttk.Label(page, text=APP_NAME, font=title_font).grid(row=0, column=0, sticky="w")
        ttk.Label(page, foreground=COLORS["muted"],
                  text="Creates an Excel register of SAAD documents and drawings, and flags every file "
                       "whose name does not match the number inside the PDF or DWG.").grid(
            row=1, column=0, sticky="w", pady=(0, 16))

        # Folder
        card = ttk.Frame(page, style="Card.TFrame", padding=(16, 12, 16, 14))
        card.grid(row=2, column=0, sticky="ew")
        card.columnconfigure(0, weight=1)
        ttk.Label(card, text="PDF / DWG folder", font=self.head_font).grid(row=0, column=0, columnspan=2, sticky="w")
        self.entry = ttk.Entry(card, textvariable=self.folder)
        self.entry.grid(row=1, column=0, sticky="ew", padx=(0, 8), pady=(6, 10))
        self.browse_btn = ttk.Button(card, text="Browse...", width=12, command=self.browse)
        self.browse_btn.grid(row=1, column=1, sticky="ew", pady=(6, 10))
        self.sub_check = ttk.Checkbutton(card, text="Include sub-folders", variable=self.recursive,
                                         style="Switch.TCheckbutton")
        self.sub_check.grid(row=2, column=0, sticky="w")
        self.count_label = ttk.Label(card, foreground=COLORS["muted"])
        self.count_label.grid(row=2, column=1, sticky="e")

        # Create + progress
        action = ttk.Frame(page)
        action.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        action.columnconfigure(0, weight=1)
        self.create_btn = ttk.Button(action, text="Create Excel Register", style="Accent.TButton", command=self.start)
        self.create_btn.grid(row=0, column=0, sticky="ew", ipady=5)
        self.stop_btn = ttk.Button(action, text="Stop", width=10, command=self.request_stop)
        self.bar = ttk.Progressbar(action, mode="determinate")
        self.bar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(14, 4))
        self.status = ttk.Label(action, foreground=COLORS["muted"],
                                text="Choose the folder with the PDF and/or DWG files, then click Create Excel Register.")
        self.status.grid(row=2, column=0, columnspan=2, sticky="w")

        # Results (shown after a run)
        self.results = ttk.Frame(page)
        self.results.grid(row=4, column=0, sticky="nsew", pady=(18, 0))
        self.results.columnconfigure(0, weight=1)
        self.results.rowconfigure(2, weight=1)
        page.rowconfigure(4, weight=1)

        tiles = ttk.Frame(self.results)
        tiles.grid(row=0, column=0, sticky="ew")
        tiles.columnconfigure(tuple(range(len(SECTIONS))), weight=1, uniform="tile")
        self.tiles = {kind: self.make_tile(tiles, title, i) for i, (kind, title, *_) in enumerate(SECTIONS)}

        self.review_label = ttk.Label(self.results, font=self.head_font)
        self.review_label.grid(row=1, column=0, sticky="w", pady=(16, 6))

        self.tree_box = ttk.Frame(self.results)
        self.tree_box.grid(row=2, column=0, sticky="nsew")
        self.tree_box.columnconfigure(0, weight=1)
        self.tree_box.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(self.tree_box, columns=("check", "file", "note"), show="headings", height=6)
        for col, text, width, stretch in (("check", "Check", 100, False), ("file", "File", 360, True),
                                          ("note", "Note", 440, True)):
            self.tree.heading(col, text=text, anchor="w")
            self.tree.column(col, width=width, minwidth=80, stretch=stretch, anchor="w")
        for status in ("MISMATCH", "CHECK"):
            self.tree.tag_configure(status, background=COLORS[status + "_bg"], foreground=COLORS[status])
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(self.tree_box, orient="vertical", command=self.tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<Double-1>", self.open_input_file)

        bottom = ttk.Frame(self.results)
        bottom.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        bottom.columnconfigure(0, weight=1)
        self.hint = ttk.Label(bottom, foreground=COLORS["muted"], text="Double-click a file to open it.")
        self.hint.grid(row=0, column=0, sticky="w")
        ttk.Button(bottom, text="Show in Folder", command=self.show_in_folder).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(bottom, text="Open Excel", style="Accent.TButton", width=14,
                   command=self.open_excel).grid(row=0, column=2)
        self.results.grid_remove()

    def make_tile(self, parent, title, column):
        tile = ttk.Frame(parent, style="Card.TFrame", padding=(16, 10, 16, 12))
        tile.grid(row=0, column=column, sticky="ew",
                  padx=(0 if column == 0 else 6, 0 if column == len(SECTIONS) - 1 else 6))
        ttk.Label(tile, text=title, foreground=COLORS["muted"]).grid(row=0, column=0, columnspan=3, sticky="w")
        total = ttk.Label(tile, font=self.big_font)
        total.grid(row=1, column=0, columnspan=3, sticky="w")
        counts = {}
        for c, key in enumerate(("OK", "MISMATCH", "CHECK")):
            counts[key] = ttk.Label(tile, foreground=COLORS[key])
            counts[key].grid(row=2, column=c, sticky="w", padx=(0, 18))
        return total, counts

    # -- folder --------------------------------------------------------------

    def chosen_folder(self):
        return os.path.normpath(self.folder.get().strip().strip('"')) if self.folder.get().strip() else ""

    def browse(self):
        current = self.chosen_folder()
        start = current if current and os.path.isdir(current) else os.path.dirname(sys.executable if getattr(
            sys, "frozen", False) else os.path.abspath(__file__))
        folder = filedialog.askdirectory(parent=self.root, title="Select the folder that contains the PDF and DWG files",
                                         initialdir=start, mustexist=True)
        if folder:
            self.folder.set(os.path.normpath(folder))

    def schedule_count(self):
        if self.count_job:
            self.root.after_cancel(self.count_job)
        self.count_job = self.root.after(400, self.update_count)

    def update_count(self):
        self.count_job = None
        folder, count = self.chosen_folder(), 0
        if not folder:
            text = ""
        elif not os.path.isdir(folder):
            text = "Folder not found"
        else:
            count = len(find_input_files(folder, self.recursive.get()))
            text = f"{count} PDF / DWG file{'' if count == 1 else 's'} found"
        self.count_label.config(text=text)
        if not self.running:
            self.create_btn.state(["!disabled"] if count else ["disabled"])

    # -- running -------------------------------------------------------------

    def set_running(self, running):
        self.running = running
        for widget in (self.entry, self.browse_btn, self.sub_check, self.create_btn):
            widget.state(["disabled"] if running else ["!disabled"])
        if running:
            self.create_btn.grid(columnspan=1)
            self.stop_btn.state(["!disabled"])
            self.stop_btn.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        else:
            self.stop_btn.grid_remove()
            self.update_count()

    def start(self):
        folder = self.chosen_folder()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning(APP_NAME, "Please choose an existing folder first.", parent=self.root)
            return
        recursive = self.recursive.get()
        if not find_input_files(folder, recursive):
            messagebox.showwarning(APP_NAME, f"No PDF or DWG files found in:\n{folder}", parent=self.root)
            return
        save_settings({"folder": folder, "recursive": recursive})

        self.stop.clear()
        self.set_running(True)
        self.results.grid_remove()
        self.bar.config(value=0, maximum=1)
        self.status.config(text="Starting...", foreground=COLORS["muted"])
        threading.Thread(target=self.work, args=(folder, recursive), daemon=True).start()
        self.root.after(50, self.poll)

    def work(self, folder, recursive):
        """Runs in a background thread so the window stays responsive."""
        def progress(i, total, name):
            if self.stop.is_set():
                raise Cancelled
            self.events.put(("progress", i, total, name))
        try:
            self.events.put(("done", folder, *make_register(folder, progress, recursive, self.stop)))
        except Cancelled:
            self.events.put(("cancelled",))
        except PermissionError as exc:
            self.events.put(("error", f"Could not save the Excel file in this folder:\n\n{exc}\n\n"
                                      "Check that you are allowed to save files there."))
        except Exception:
            self.events.put(("error", "Something went wrong:\n\n" + traceback.format_exc(limit=4)))

    def request_stop(self):
        self.stop.set()
        self.stop_btn.state(["disabled"])
        self.status.config(text="Stopping...")

    def poll(self):
        while True:
            try:
                event = self.events.get_nowait()
            except queue.Empty:
                break
            if event[0] == "progress":
                _, i, total, name = event
                self.bar.config(maximum=total, value=i - 1)
                self.status.config(text=f"Reading {i} of {total}:   {name}")
                continue
            self.set_running(False)
            if self.closing:
                self.root.destroy()
            elif event[0] == "done":
                self.show_results(*event[1:])
            elif event[0] == "cancelled":
                self.bar.config(value=0)
                self.status.config(text="Stopped. No Excel file was created.")
            else:
                self.bar.config(value=0)
                self.status.config(text="The Excel file was not created.", foreground=COLORS["MISMATCH"])
                messagebox.showerror(APP_NAME, event[1], parent=self.root)
            return
        self.root.after(50, self.poll)

    def show_results(self, folder, out, results):
        self.output = out
        self.bar.config(value=self.bar["maximum"])
        self.status.config(text=f"Excel register saved in the selected folder:   {os.path.basename(out)}",
                           foreground=COLORS["OK"])

        for kind, (total, counts) in self.tiles.items():
            rows = results[kind]
            total.config(text=str(len(rows)))
            n = {key: sum(r["status"] == key for r in rows) for key in counts}
            counts["OK"].config(text=f"{n['OK']} OK")
            counts["MISMATCH"].config(text=f"{n['MISMATCH']} mismatch")
            counts["CHECK"].config(text=f"{n['CHECK']} to check")

        problems = sorted((r for rows in results.values() for r in rows if r["status"] != "OK"),
                          key=lambda r: (r["status"] != "MISMATCH", natural_key(r["file"])))
        self.tree.delete(*self.tree.get_children())
        self.file_paths.clear()
        for r in problems:
            row = self.tree.insert("", "end", values=(r["status"], r["file"], r["note"]), tags=(r["status"],))
            self.file_paths[row] = os.path.join(folder, r["file"])
        if problems:
            self.review_label.config(text=f"Files to review ({len(problems)})", foreground="")
            self.tree_box.grid()
            self.hint.grid()
        else:
            self.review_label.config(text="All file names match the numbers inside the PDF and DWG files.",
                                     foreground=COLORS["OK"])
            self.tree_box.grid_remove()
            self.hint.grid_remove()
        self.results.grid()

        # Grow the window if the results do not fit, but never past the screen.
        root = self.root
        root.update_idletasks()
        screen_w, screen_h = root.winfo_screenwidth(), root.winfo_screenheight()
        width = max(root.winfo_width(), min(root.winfo_reqwidth(), screen_w - 40))
        height = max(root.winfo_height(), min(root.winfo_reqheight(), screen_h - 100))
        if (width, height) != (root.winfo_width(), root.winfo_height()):
            x = max(min(root.winfo_x(), screen_w - width - 20), 0)
            y = max(min(root.winfo_y(), screen_h - height - 80), 0)
            root.geometry(f"{width}x{height}+{x}+{y}")

    # -- buttons -------------------------------------------------------------

    def open_excel(self):
        if self.output and os.path.exists(self.output):
            os.startfile(self.output)

    def show_in_folder(self):
        if self.output and os.path.exists(self.output):
            subprocess.Popen(["explorer", "/select,", self.output])

    def open_input_file(self, _event=None):
        path = self.file_paths.get(self.tree.focus())
        if path and os.path.exists(path):
            try:
                os.startfile(path)
            except OSError as exc:
                messagebox.showerror(APP_NAME, f"Could not open the selected file:\n{exc}", parent=self.root)

    def on_close(self):
        if not self.running:
            self.root.destroy()
        elif messagebox.askyesno(APP_NAME, "The PDF and DWG files are still being read.\n\nStop and close?", parent=self.root):
            self.closing = True     # poll() closes the window once the reading has stopped
            self.request_stop()


def run_window():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("PDFCADRegister")
    except Exception:
        pass

    root = tk.Tk()
    try:
        import sv_ttk
        sv_ttk.set_theme("light")
    except ImportError:
        pass
    if fitz is None or openpyxl is None or ezdxf is None:
        root.withdraw()
        messagebox.showerror(APP_NAME, "Missing Python packages. Open a command prompt and run:\n\n"
                                       "python -m pip install pymupdf ezdxf openpyxl sv-ttk")
        return

    App(root)
    root.update_idletasks()
    width, height = max(root.winfo_reqwidth(), 880), root.winfo_reqheight()
    root.minsize(720, height)
    root.geometry(f"{width}x{height}+{(root.winfo_screenwidth() - width) // 2}"
                  f"+{max((root.winfo_screenheight() - height) // 4, 0)}")
    root.attributes("-topmost", True)           # come up in front of everything ...
    root.after(1000, lambda: root.attributes("-topmost", False))    # ... then behave normally
    root.focus_force()
    root.mainloop()


def main():
    if len(sys.argv) > 1:
        if fitz is None or openpyxl is None or ezdxf is None:
            sys.exit("Missing packages. Run:  python -m pip install pymupdf ezdxf openpyxl")
        folder = os.path.abspath(sys.argv[1])
        out, results = make_register(folder, lambda i, n, name: print(f"[{i}/{n}] {name}"))
        print()
        print(summary_text(out, results))
        return
    try:
        run_window()
    except Exception:
        messagebox.showerror(APP_NAME, "Something went wrong:\n\n" + traceback.format_exc(limit=3))


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--check-converter":
        import tempfile
        report_path = (
            sys.argv[2] if len(sys.argv) > 2
            else os.path.join(tempfile.gettempdir(), "pdf_cad_converter_check.txt")
        )
        try:
            converter, version = _check_bundled_converter()
            report = "BUNDLED_CONVERTER_OK\nPATH=" + converter + "\nVERSION=" + version + "\n"
            exit_code = 0
        except Exception as exc:
            report = "BUNDLED_CONVERTER_ERROR: " + type(exc).__name__ + ": " + str(exc) + "\n"
            exit_code = 2
        with open(report_path, "w", encoding="utf-8") as report_file:
            report_file.write(report)
        sys.exit(exit_code)
    main()
