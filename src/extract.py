#!/usr/bin/env python3
"""Marginalia — core extraction, formatting, and CLI."""

import argparse
import bisect
import csv
import io
import math
import re
import statistics
import sys
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF

# ---------------------------------------------------------------------------
# Annotation type codes in PyMuPDF
# ---------------------------------------------------------------------------
HIGHLIGHT   = 8
UNDERLINE   = 9
SQUIGGLY    = 10
STRIKEOUT   = 11
STICKY_NOTE = 0
FREE_TEXT   = 2

ANNOT_LABELS = {
    HIGHLIGHT:   "Highlight",
    UNDERLINE:   "Underline",
    SQUIGGLY:    "Squiggly",
    STRIKEOUT:   "Strikeout",
    STICKY_NOTE: "Comment",
    FREE_TEXT:   "Comment",
}

# ---------------------------------------------------------------------------
# Color mapping
# ---------------------------------------------------------------------------
COLOR_MAP = {
    "Yellow": (1.0, 1.0, 0.0),
    "Green":  (0.0, 1.0, 0.0),
    "Blue":   (0.0, 0.0, 1.0),
    "Pink":   (1.0, 0.0, 1.0),
    "Orange": (1.0, 0.65, 0.0),
    "Purple": (0.5, 0.0, 0.5),
    "Red":    (1.0, 0.0, 0.0),
}

COLOR_HEX = {
    "Yellow":  "#FFD700",
    "Green":   "#34C759",
    "Blue":    "#007AFF",
    "Pink":    "#FF2D55",
    "Orange":  "#FF9500",
    "Purple":  "#AF52DE",
    "Red":     "#FF3B30",
    "Unknown": "#8E8E93",
}

COLOR_EMOJI = {
    "Yellow":  "\U0001f7e1",  # 🟡
    "Green":   "\U0001f7e2",  # 🟢
    "Blue":    "\U0001f535",  # 🔵
    "Pink":    "\U0001fa77",  # 🩷
    "Orange":  "\U0001f7e0",  # 🟠
    "Purple":  "\U0001f7e3",  # 🟣
    "Red":     "\U0001f534",  # 🔴
    "Unknown": "\u26aa",      # ⚪
}

TYPE_EMOJI = {
    "Highlight": "\U0001f4dd",  # 📝
    "Comment":   "\U0001f4ac",  # 💬
    "Underline": "\u2015",      # ―
    "Strikeout": "\u2716",      # ✖
    "Squiggly":  "\u3030",      # 〰
}

MEANING_PRESETS = {
    "Academic Reading": {
        "Yellow": "Key Argument", "Green": "Evidence / Data",
        "Blue":   "Definition",   "Red":   "Critique",
        "Orange": "Follow Up",    "Purple": "Quote to Cite", "Pink": "Interesting",
    },
    "Book Notes": {
        "Yellow": "Main Idea",   "Green": "Example",
        "Blue":   "Vocabulary",  "Red":   "Disagree",
        "Orange": "Action Item", "Purple": "Inspiring",     "Pink": "Story",
    },
    "Code Review": {
        "Yellow": "Question",    "Green": "Good Pattern",
        "Blue":   "Reference",   "Red":   "Bug / Issue",
        "Orange": "Refactor",    "Purple": "Architecture",  "Pink": "Test Case",
    },
}


# ---------------------------------------------------------------------------
# Color helpers
# ---------------------------------------------------------------------------
def _color_distance(c1, c2):
    return math.sqrt(sum((a - b) ** 2 for a, b in zip(c1, c2)))


def rgb_to_color_name(rgb):
    if rgb is None:
        return "Unknown"
    best_name, best_dist = "Unknown", float("inf")
    for name, ref in COLOR_MAP.items():
        d = _color_distance(rgb, ref)
        if d < best_dist:
            best_name, best_dist = name, d
    return best_name if best_dist < 0.6 else "Unknown"


# ---------------------------------------------------------------------------
# Context extraction
# ---------------------------------------------------------------------------
def _get_surrounding_context(page, highlight_text: str, n_sentences: int = 2) -> str:
    """Return up to n_sentences before/after a highlight.

    Uses sort=True for reading-order text and a smarter sentence boundary
    pattern that avoids splitting on common abbreviations (Mr., e.g., etc.).
    """
    full = " ".join(page.get_text("text", sort=True).split())
    hl   = " ".join(highlight_text.split())

    if not hl:
        return ""

    pos = full.find(hl)
    if pos < 0:
        pos = full.lower().find(hl.lower())
    if pos < 0:
        return ""

    end = pos + len(hl)

    # Sentence boundaries: punctuation + optional quote/paren + space + capital letter.
    # Requiring the next char to be uppercase avoids "e.g. " and "Mr. " false breaks.
    boundaries = [0]
    for m in re.finditer(r'[.!?]["\'\)]*\s+(?=[A-Z\d])', full):
        boundaries.append(m.end())
    boundaries.append(len(full))

    before_idx = 0
    for i, b in enumerate(boundaries):
        if b <= pos:
            before_idx = i

    after_idx = len(boundaries) - 1
    for i, b in enumerate(boundaries):
        if b >= end:
            after_idx = i
            break

    ctx_start = boundaries[max(0, before_idx - n_sentences + 1)]
    ctx_end   = boundaries[min(len(boundaries) - 1, after_idx + n_sentences)]
    return full[ctx_start:ctx_end].strip()


# ---------------------------------------------------------------------------
# Section detection
# ---------------------------------------------------------------------------
def _detect_sections(doc) -> list[tuple[int, float, str]]:
    """Detect section headings; fall back to page-range bands."""
    sections: list[tuple[int, float, str]] = []

    for page_num, page in enumerate(doc, start=1):
        try:
            page_dict = page.get_text("dict")
        except Exception:
            continue

        blocks = page_dict.get("blocks", [])
        all_sizes: list[float] = []
        for block in blocks:
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    size = span.get("size", 0)
                    if size > 0:
                        all_sizes.append(size)

        if not all_sizes:
            continue
        try:
            body_size = statistics.median(all_sizes)
        except Exception:
            continue

        seen_texts: set[str] = set()
        for block in blocks:
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = span.get("text", "").strip()
                    if not text or len(text) < 3 or len(text) >= 120:
                        continue
                    if re.match(r'^[\d\s\.\-\u2013\u2014\u00b7\u2022]+$', text):
                        continue
                    size  = span.get("size", 0)
                    flags = span.get("flags", 0)
                    is_bold   = bool(flags & 16)
                    is_larger = size > 0 and size > body_size * 1.15
                    if (is_larger or is_bold) and text not in seen_texts:
                        seen_texts.add(text)
                        y0 = float(span.get("bbox", [0, 0, 0, 0])[1])
                        sections.append((page_num, y0, text))

    sections.sort(key=lambda x: (x[0], x[1]))

    if len(sections) < 2:
        page_count = doc.page_count
        band = 5 if page_count <= 30 else 10
        sections = []
        for start in range(0, page_count, band):
            end = min(start + band, page_count)
            sections.append((start + 1, 0.0, f"Pages {start + 1}\u2013{end}"))

    return sections


# ---------------------------------------------------------------------------
# Metadata extraction  (accepts optional open doc to avoid re-opening)
# ---------------------------------------------------------------------------
def extract_pdf_metadata(pdf_path: str, _doc=None) -> dict:
    doc = _doc if _doc is not None else fitz.open(pdf_path)
    close = _doc is None
    meta = doc.metadata or {}
    page_count = doc.page_count
    if close:
        doc.close()

    def _clean_date(d: str) -> str:
        if not d:
            return ""
        if d.startswith("D:"):
            d = d[2:]
        try:
            return datetime.strptime(d[:8], "%Y%m%d").strftime("%Y-%m-%d")
        except Exception:
            return ""

    return {
        "filename": Path(pdf_path).name,
        "title":    meta.get("title",    "").strip() or None,
        "author":   meta.get("author",   "").strip() or None,
        "subject":  meta.get("subject",  "").strip() or None,
        "keywords": meta.get("keywords", "").strip() or None,
        "pages":    page_count,
        "created":  _clean_date(meta.get("creationDate", "")),
        "modified": _clean_date(meta.get("modDate", "")),
    }


# ---------------------------------------------------------------------------
# Word count  (accepts optional open doc)
# ---------------------------------------------------------------------------
def extract_word_counts(pdf_path: str, annotations: list[dict], _doc=None) -> dict:
    doc = _doc if _doc is not None else fitz.open(pdf_path)
    close = _doc is None
    full_text = " ".join(page.get_text("text", sort=True) for page in doc)
    if close:
        doc.close()

    doc_words = len(full_text.split())
    hl_words  = sum(len(a["text"].split()) for a in annotations if a["type"] == "Highlight")
    return {
        "document":   doc_words,
        "highlights": hl_words,
        "ratio":      round(hl_words / doc_words * 100, 1) if doc_words else 0,
    }


# ---------------------------------------------------------------------------
# Annotation extraction  (accepts optional open doc)
# ---------------------------------------------------------------------------
def extract_annotations(pdf_path: str, with_context: bool = False, _doc=None) -> list[dict]:
    """Parse a PDF and return annotation dicts.

    Passing _doc (an already-open fitz.Document) avoids a redundant open/close
    when called alongside extract_pdf_metadata and extract_word_counts.
    """
    doc = _doc if _doc is not None else fitz.open(pdf_path)
    close = _doc is None

    sections = _detect_sections(doc)
    section_keys = [(s[0], s[1]) for s in sections]
    annotations = []

    for page_num, page in enumerate(doc, start=1):
        for annot in page.annots() or []:
            annot_type = annot.type[0]
            if annot_type not in ANNOT_LABELS:
                continue

            label = ANNOT_LABELS[annot_type]
            text  = ""
            raw_rgb    = (annot.colors or {}).get("stroke") or (annot.colors or {}).get("fill")
            color_name = rgb_to_color_name(tuple(raw_rgb) if raw_rgb else None)

            if annot_type == HIGHLIGHT:
                quads = annot.vertices
                if quads:
                    quad_count = len(quads) // 4
                    for i in range(quad_count):
                        quad = fitz.Quad(quads[i * 4 : i * 4 + 4])
                        text += page.get_text("text", clip=quad.rect, sort=True).strip() + " "
                    text = text.strip()
                comment = annot.info.get("content", "").strip()
                if comment:
                    text = f"{text}\n  > Note: {comment}" if text else comment

            elif annot_type in (UNDERLINE, SQUIGGLY, STRIKEOUT):
                quads = annot.vertices
                if quads:
                    quad_count = len(quads) // 4
                    for i in range(quad_count):
                        quad = fitz.Quad(quads[i * 4 : i * 4 + 4])
                        text += page.get_text("text", clip=quad.rect, sort=True).strip() + " "
                    text = text.strip()
            else:
                text = annot.info.get("content", "").strip()

            if not text:
                continue

            section_name = ""
            if sections:
                idx = bisect.bisect_right(section_keys, (page_num, annot.rect.y0)) - 1
                section_name = sections[idx][2] if idx >= 0 else ""

            entry = {
                "page":      page_num,
                "type":      label,
                "text":      text,
                "color":     color_name,
                "color_rgb": tuple(raw_rgb) if raw_rgb else None,
                "context":   "",
                "section":   section_name,
            }

            if with_context and annot_type == HIGHLIGHT:
                entry["context"] = _get_surrounding_context(page, text)

            annotations.append(entry)

    if close:
        doc.close()
    return annotations


# ---------------------------------------------------------------------------
# Convenience: open once, extract everything
# ---------------------------------------------------------------------------
def extract_all(
    pdf_path: str,
    with_context: bool = False,
) -> tuple[list[dict], dict, dict]:
    """Open the PDF once and return (annotations, metadata, word_counts).

    Avoids the triple-open overhead when loading a file in the GUI.
    """
    doc = fitz.open(pdf_path)
    try:
        anns  = extract_annotations(pdf_path, with_context=with_context, _doc=doc)
        meta  = extract_pdf_metadata(pdf_path, _doc=doc)
        wc    = extract_word_counts(pdf_path, anns, _doc=doc)
    finally:
        doc.close()
    return anns, meta, wc


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------
def _page_link(page_num: int, pdf_path, deep_links: bool) -> str:
    if deep_links and pdf_path:
        uri = Path(pdf_path).as_uri()
        return f"[Page {page_num}]({uri}#page={page_num})"
    return f"Page {page_num}"


def _annotation_summary(annotations: list[dict]) -> str:
    counts: dict[str, int] = {}
    for a in annotations:
        counts[a["type"]] = counts.get(a["type"], 0) + 1
    return ", ".join(f"{v} {k}{'s' if v != 1 else ''}" for k, v in sorted(counts.items()))


def _split_text_note(text: str) -> tuple[str, str]:
    sep = "\n  > Note: "
    if sep in text:
        parts = text.split(sep, 1)
        return parts[0].strip(), parts[1].strip()
    return text, ""


# ---------------------------------------------------------------------------
# Standard Markdown
# ---------------------------------------------------------------------------
def format_markdown(annotations, pdf_path=None, deep_links=False) -> str:
    if not annotations:
        return ""
    buckets: dict[str, list[dict]] = {}
    for a in annotations:
        key = f'{a["color"]} Highlights' if a["type"] == "Highlight" else f'{a["type"]}s'
        buckets.setdefault(key, []).append(a)
    lines = []
    for header, items in buckets.items():
        items.sort(key=lambda x: x["page"])
        lines.append(f"## {header} ({len(items)})")
        lines.append("")
        for a in items:
            ref = _page_link(a["page"], pdf_path, deep_links)
            lines.append(f'- **[{ref}]**: "{a["text"]}"')
            if a.get("context"):
                lines.append(f'  > *Context: {a["context"]}*')
        lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Markdown with metadata header
# ---------------------------------------------------------------------------
def format_markdown_with_metadata(
    annotations, metadata, word_counts=None, pdf_path=None, deep_links=False
) -> str:
    lines = []
    filename = metadata.get("filename", "Document")
    lines += [f"# \U0001f4c4 {filename}", ""]
    for label, key in [("Author", "author"), ("Title", "title"),
                        ("Pages", "pages"), ("Created", "created")]:
        val = metadata.get(key)
        if val and (key != "title" or val != filename):
            lines.append(f"- **{label}:** {val}")
    lines.append(f"- **Annotations:** {len(annotations)} — {_annotation_summary(annotations)}")
    if word_counts:
        lines.append(
            f"- **Coverage:** {word_counts['highlights']:,} highlighted words"
            f" of {word_counts['document']:,} ({word_counts['ratio']}%)"
        )
    lines += [f"- **Extracted:** {datetime.now().strftime('%Y-%m-%d')}", "", "---", ""]
    lines.append(format_markdown(annotations, pdf_path, deep_links))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Obsidian callouts
# ---------------------------------------------------------------------------
def format_obsidian(annotations, metadata=None, pdf_path=None, deep_links=False) -> str:
    lines = []
    if metadata:
        filename = metadata.get("filename", "Document")
        lines += [f"# \U0001f4c4 {filename}", ""]
        if metadata.get("author"):
            lines.append(f"**Author:** {metadata['author']}")
        if metadata.get("pages"):
            lines.append(f"**Pages:** {metadata['pages']}")
        lines.append("")
    for a in annotations:
        ref = _page_link(a["page"], pdf_path, deep_links)
        if a["type"] == "Highlight":
            emoji = COLOR_EMOJI.get(a["color"], "")
            lines.append(f"> [!quote] {emoji} {a['color']} Highlight — {ref}")
        else:
            lines.append(f"> [!note] {a['type']} — {ref}")
        lines.append(f"> {a['text']}")
        if a.get("context"):
            lines += [">", f"> *Context: {a['context']}*"]
        lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# CSV export
# ---------------------------------------------------------------------------
def format_csv(annotations: list[dict], color_meanings: dict | None = None) -> str:
    color_meanings = color_meanings or {}
    out = io.StringIO()
    w = csv.writer(out, quoting=csv.QUOTE_ALL)
    w.writerow(["Filename", "Page", "Section", "Type", "Color",
                "Meaning", "Text", "Note", "Context"])
    for a in annotations:
        text, note = _split_text_note(a["text"])
        w.writerow([
            a.get("filename", ""), a["page"], a.get("section", ""),
            a["type"], a["color"], color_meanings.get(a["color"], ""),
            text, note, a.get("context", ""),
        ])
    return out.getvalue()


# ---------------------------------------------------------------------------
# Markdown table export
# ---------------------------------------------------------------------------
def format_markdown_table(annotations: list[dict], color_meanings: dict | None = None) -> str:
    color_meanings = color_meanings or {}

    def _cell(s: str, max_len: int = 120) -> str:
        s = str(s).replace("\n", " ").replace("|", "\\|")
        return s[: max_len - 1] + "\u2026" if len(s) > max_len else s

    rows = [
        "| Page | Section | Color | Meaning | Type | Text |",
        "|------|---------|-------|---------|------|------|",
    ]
    for a in annotations:
        text, _ = _split_text_note(a["text"])
        rows.append(
            f"| {a['page']} | {_cell(a.get('section',''))} "
            f"| {_cell(a['color'])} | {_cell(color_meanings.get(a['color'],''))} "
            f"| {_cell(a['type'])} | {_cell(text)} |"
        )
    return "\n".join(rows)


# ---------------------------------------------------------------------------
# HTML report
# ---------------------------------------------------------------------------
def format_html(
    annotations: list[dict],
    metadata: dict | None = None,
    word_counts: dict | None = None,
    color_meanings: dict | None = None,
) -> str:
    color_meanings = color_meanings or {}

    def esc(s: str) -> str:
        return (str(s).replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;").replace('"', "&quot;"))

    meta = metadata or {}
    filename  = meta.get("filename") or "Annotations"
    doc_title = meta.get("title") or filename

    meta_parts = []
    if meta.get("author"):
        meta_parts.append(f'<span>\U0001f464 {esc(meta["author"])}</span>')
    if meta.get("pages"):
        meta_parts.append(f'<span>\U0001f4c4 {esc(str(meta["pages"]))} pages</span>')
    count = len(annotations)
    meta_parts.append(f'<span>\U0001f516 {count} annotation{"s" if count != 1 else ""}</span>')
    if word_counts:
        meta_parts.append(f'<span>\U0001f4ca {word_counts["ratio"]}% word coverage</span>')

    colors = sorted({a["color"] for a in annotations})
    legend_items = []
    for color in colors:
        hex_c   = COLOR_HEX.get(color, "#8E8E93")
        meaning = color_meanings.get(color, "")
        label   = esc(color) + (f" &mdash; {esc(meaning)}" if meaning else "")
        legend_items.append(
            f'<span class="legend-item">'
            f'<span class="swatch" style="background:{hex_c}"></span>{label}</span>'
        )

    pills_html = '<button class="pill" data-filter="all" data-bg="#5C5FEF">All</button>\n'
    for color in colors:
        hex_c = COLOR_HEX.get(color, "#8E8E93")
        emoji = COLOR_EMOJI.get(color, "")
        pills_html += (
            f'    <button class="pill" data-filter="{esc(color)}" data-bg="{hex_c}">'
            f'{emoji} {esc(color)}</button>\n'
        )

    sec_groups: dict[str, list[dict]] = {}
    sec_order:  list[str] = []
    for a in annotations:
        sec = a.get("section", "") or "General"
        if sec not in sec_groups:
            sec_groups[sec] = []
            sec_order.append(sec)
        sec_groups[sec].append(a)

    sections_html = ""
    for sec in sec_order:
        items = sec_groups[sec]
        cards_html = ""
        for a in items:
            color    = a.get("color", "Unknown")
            hex_c    = COLOR_HEX.get(color, "#C7C7CC")
            ann_type = a["type"]
            emoji    = (COLOR_EMOJI.get(color, "") if ann_type == "Highlight"
                        else TYPE_EMOJI.get(ann_type, ""))
            meaning  = color_meanings.get(color, "")
            meaning_html = (f' <span class="meaning">&mdash; {esc(meaning)}</span>'
                            if meaning else "")
            text, note = _split_text_note(a["text"])
            context    = a.get("context", "")
            note_html  = f'<div class="card-sub"><em>Note:</em> {esc(note)}</div>' if note else ""
            ctx_html   = f'<div class="card-sub"><em>Context:</em> {esc(context)}</div>' if context else ""
            cards_html += (
                f'<div class="card" data-color="{esc(color)}" data-type="{esc(ann_type)}" '
                f'data-section="{esc(a.get("section",""))}" style="border-left-color:{hex_c}">\n'
                f'  <div class="card-header">'
                f'<span class="card-color">{emoji} {esc(color)}</span>'
                f'<span class="type-badge">{esc(ann_type)}</span>'
                f'{meaning_html}'
                f'<span class="page-badge">p.&thinsp;{a["page"]}</span></div>\n'
                f'  <div class="card-body">{esc(text)}</div>\n'
                f'{note_html}{ctx_html}</div>\n'
            )
        sections_html += (
            f'<div class="section-group">\n'
            f'<h2 class="section-title">{esc(sec)} '
            f'<span class="section-count">({len(items)})</span></h2>\n'
            f'{cards_html}</div>\n'
        )

    css = (
        "*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}"
        "body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;"
        "background:#F6F5F3;color:#1C1917;line-height:1.5}"
        ".doc-header{background:white;border-bottom:1px solid #E2DDD8;padding:24px 32px}"
        ".doc-title{font-size:22px;font-weight:700;margin-bottom:8px}"
        ".doc-meta{color:#6B6560;font-size:13px;display:flex;flex-wrap:wrap;gap:16px;margin-bottom:10px}"
        ".legend{display:flex;flex-wrap:wrap;gap:12px;margin-top:6px}"
        ".legend-item{display:flex;align-items:center;gap:5px;font-size:13px}"
        ".swatch{width:10px;height:10px;border-radius:50%;display:inline-block;flex-shrink:0}"
        ".controls{position:sticky;top:0;z-index:10;background:rgba(246,245,243,.96);"
        "backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);"
        "border-bottom:1px solid #E2DDD8;padding:10px 32px;"
        "display:flex;align-items:center;gap:12px;flex-wrap:wrap}"
        ".search-input{padding:6px 12px;border:1px solid #C7C7CC;border-radius:8px;"
        "font-size:14px;width:200px;outline:none;background:white}"
        ".search-input:focus{border-color:#5C5FEF;box-shadow:0 0 0 3px rgba(92,95,239,.15)}"
        ".pills{display:flex;gap:6px;flex-wrap:wrap}"
        ".pill{padding:4px 12px;border-radius:16px;font-size:12px;cursor:pointer;"
        "border:1.5px solid #E2DDD8;background:white;color:#1C1917;"
        "transition:all .15s;font-weight:500}"
        ".pill:hover{border-color:#5C5FEF}"
        ".result-count{margin-left:auto;font-size:13px;color:#A09891;white-space:nowrap}"
        "main{max-width:860px;margin:0 auto;padding:24px 32px 56px}"
        ".section-group{margin-bottom:28px}"
        ".section-title{font-size:16px;font-weight:600;color:#1C1917;"
        "padding:10px 0 8px;border-bottom:2px solid #E2DDD8;"
        "margin-bottom:10px;display:flex;align-items:baseline;gap:8px}"
        ".section-count{font-size:13px;font-weight:400;color:#A09891}"
        ".card{background:white;border-radius:8px;margin-bottom:8px;"
        "border-left:4px solid #C7C7CC;box-shadow:0 1px 3px rgba(0,0,0,.06);"
        "overflow:hidden;transition:box-shadow .15s}"
        ".card:hover{box-shadow:0 3px 10px rgba(0,0,0,.1)}"
        ".card-header{padding:7px 12px;background:#FAFAF9;"
        "display:flex;align-items:center;gap:7px;font-size:12px;font-weight:600;"
        "border-bottom:1px solid #F2F2F0}"
        ".type-badge{padding:1px 7px;border-radius:4px;font-size:11px;"
        "background:#EDECEA;color:#6B6560;font-weight:500}"
        ".meaning{font-style:italic;font-weight:400;color:#6B6560;font-size:12px}"
        ".page-badge{margin-left:auto;font-size:11px;color:#A09891;"
        "background:#F0EEE9;padding:1px 7px;border-radius:4px;font-weight:400}"
        ".card-body{padding:10px 14px;font-size:14px;line-height:1.65;color:#1C1917}"
        ".card-sub{padding:6px 14px 10px;font-size:13px;color:#6B6560;"
        "font-style:italic;border-top:1px solid #F5F3F0}"
        ".hidden{display:none!important}"
        "@media print{.controls{display:none!important}body{background:white}"
        "main{max-width:100%;padding:0}.card{box-shadow:none;break-inside:avoid}}"
    )

    js = (
        "const cards=document.querySelectorAll('.card');\n"
        "const secs=document.querySelectorAll('.section-group');\n"
        "const pills=document.querySelectorAll('.pill');\n"
        "const countEl=document.querySelector('.result-count');\n"
        "let activeColor='all',query='';\n"
        "function refresh(){\n"
        "  let n=0;\n"
        "  cards.forEach(c=>{\n"
        "    const cm=activeColor==='all'||c.dataset.color===activeColor;\n"
        "    const qm=!query||c.textContent.toLowerCase().includes(query);\n"
        "    const show=cm&&qm;c.classList.toggle('hidden',!show);if(show)n++;\n"
        "  });\n"
        "  secs.forEach(s=>{\n"
        "    const any=[...s.querySelectorAll('.card')].some(c=>!c.classList.contains('hidden'));\n"
        "    s.classList.toggle('hidden',!any);\n"
        "  });\n"
        "  if(countEl)countEl.textContent=n+' annotation'+(n===1?'':'s');\n"
        "}\n"
        "document.getElementById('search').addEventListener('input',e=>{\n"
        "  query=e.target.value.toLowerCase().trim();refresh();\n"
        "});\n"
        "pills.forEach(p=>{\n"
        "  p.addEventListener('click',()=>{\n"
        "    pills.forEach(x=>{x.classList.remove('active');\n"
        "      x.style.background='white';x.style.color='#1C1917';\n"
        "      x.style.borderColor='#E2DDD8';});\n"
        "    p.classList.add('active');\n"
        "    p.style.background=p.dataset.bg||'#5C5FEF';\n"
        "    p.style.color='white';p.style.borderColor='transparent';\n"
        "    activeColor=p.dataset.filter;refresh();\n"
        "  });\n"
        "});\n"
        "const allPill=document.querySelector('.pill[data-filter=\"all\"]');\n"
        "if(allPill){allPill.style.background='#5C5FEF';\n"
        "allPill.style.color='white';allPill.style.borderColor='transparent';}\n"
    )

    meta_html    = "\n      ".join(meta_parts)
    legend_html  = "\n      ".join(legend_items)

    return (
        "<!DOCTYPE html>\n"
        '<html lang="en">\n<head>\n'
        '<meta charset="UTF-8">\n'
        '<meta name="viewport" content="width=device-width,initial-scale=1">\n'
        f"<title>{esc(doc_title)} \u2014 Marginalia</title>\n"
        f"<style>{css}</style>\n</head>\n<body>\n"
        '<header class="doc-header">\n'
        f'  <div class="doc-title">\U0001f4c4 {esc(filename)}</div>\n'
        f'  <div class="doc-meta">\n      {meta_html}\n  </div>\n'
        f'  <div class="legend">\n      {legend_html}\n  </div>\n'
        "</header>\n"
        '<div class="controls">\n'
        '  <input type="search" id="search" class="search-input"'
        ' placeholder="Search annotations\u2026">\n'
        f'  <div class="pills">\n{pills_html}  </div>\n'
        f'  <span class="result-count">{count} annotations</span>\n'
        "</div>\n"
        f"<main>\n{sections_html}</main>\n"
        f"<script>\n{js}</script>\n</body>\n</html>\n"
    )


# ---------------------------------------------------------------------------
# Batch Markdown
# ---------------------------------------------------------------------------
def format_batch_markdown(results, with_metadata=True, deep_links=False) -> str:
    sections = []
    for r in results:
        meta        = r.get("metadata") or {}
        filename    = r.get("filename", "Unknown")
        annotations = r.get("annotations", [])
        word_counts = r.get("word_counts")
        path        = r.get("path")
        if not meta:
            meta = {"filename": filename}
        if with_metadata:
            section = format_markdown_with_metadata(annotations, meta, word_counts, path, deep_links)
        else:
            body = format_markdown(annotations, path, deep_links) if annotations else "*No annotations found.*\n"
            section = f"# \U0001f4c4 {filename}\n\n{body}"
        sections.append(section)
    return "\n\n---\n\n".join(sections)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    import pyperclip
    parser = argparse.ArgumentParser(description="Marginalia — extract PDF annotations.")
    parser.add_argument("pdf", help="Path to the PDF file")
    parser.add_argument("--context",  action="store_true")
    parser.add_argument("--metadata", action="store_true")
    parser.add_argument("--obsidian", action="store_true")
    parser.add_argument("--csv",      action="store_true")
    parser.add_argument("--table",    action="store_true")
    args = parser.parse_args()

    try:
        annotations = extract_annotations(args.pdf, with_context=args.context)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    if not annotations:
        print("No annotations found.")
        sys.exit(0)

    if args.csv:
        md = format_csv(annotations)
    elif args.table:
        md = format_markdown_table(annotations)
    elif args.obsidian:
        md = format_obsidian(annotations)
    elif args.metadata:
        meta = extract_pdf_metadata(args.pdf)
        wc   = extract_word_counts(args.pdf, annotations)
        md   = format_markdown_with_metadata(annotations, meta, wc)
    else:
        md = format_markdown(annotations)

    try:
        pyperclip.copy(md)
        clipboard_ok = True
    except Exception:
        clipboard_ok = False

    print(md)
    count = len(annotations)
    if clipboard_ok:
        print(f"\nExtracted {count} annotation(s) and copied to clipboard.")
    else:
        print(f"\nExtracted {count} annotation(s).")


if __name__ == "__main__":
    main()
