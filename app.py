from __future__ import annotations

import base64
import hashlib
import html
import re
from pathlib import Path
from typing import Iterable

from flask import Flask, Response, abort, render_template
from werkzeug.utils import secure_filename
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn
import pymupdf
import markdown
import bleach

app = Flask(__name__)
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
CONVERTED_FOLDER = BASE_DIR / "converted"
UPLOAD_FOLDER.mkdir(exist_ok=True)
CONVERTED_FOLDER.mkdir(exist_ok=True)
SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".md", ".markdown", ".txt"}

MD = markdown.Markdown(extensions=["extra", "tables", "fenced_code", "sane_lists", "toc", "footnotes"], output_format="html5")
ALLOWED_TAGS = set(bleach.sanitizer.ALLOWED_TAGS).union({
    "p", "h1", "h2", "h3", "h4", "h5", "h6", "table", "thead", "tbody", "tfoot", "tr", "th", "td",
    "pre", "code", "hr", "br", "div", "span", "img", "figure", "figcaption", "details", "summary",
    "del", "sup", "sub", "section", "blockquote", "dl", "dt", "dd", "abbr", "kbd", "mark", "ol", "ul", "li"
})
ALLOWED_ATTRIBUTES = {
    **bleach.sanitizer.ALLOWED_ATTRIBUTES,
    "*": ["id", "class", "title", "aria-label"],
    "a": ["href", "title", "rel"],
    "img": ["src", "alt", "title", "width", "height", "loading"],
    "th": ["colspan", "rowspan", "align", "scope"], "td": ["colspan", "rowspan", "align"],
    "code": ["class"], "pre": ["class"], "div": ["class", "id"], "span": ["class", "id"],
}


def esc(text: str) -> str:
    return html.escape(text or "", quote=False)


def normalize_breaks(text: str) -> str:
    # Some of the supplied DOCX files contain literal '<br>' characters as text.
    text = re.sub(r"(?i)<\s*br\s*/?\s*>", "\n", text or "")
    text = re.sub(r"(?i)</\s*p\s*>", "\n", text)
    return text.replace("\x00", "")


def inline_runs(paragraph: Paragraph) -> str:
    """Render Word runs individually so emphasis and inline code remain distinct."""
    parts: list[str] = []
    for run in paragraph.runs:
        text = normalize_breaks(run.text or "")
        if not text:
            continue
        escaped = esc(text).replace("\n", "<br>\n")
        if run.bold:
            escaped = f"<strong>{escaped}</strong>"
        if run.italic:
            escaped = f"<em>{escaped}</em>"
        if run.underline:
            escaped = f"<u>{escaped}</u>"
        if run.font and run.font.name and run.font.name.lower() in {"consolas", "courier new", "cascadia code", "source code pro"}:
            escaped = f"<code>{escaped}</code>"
        parts.append(escaped)
    return "".join(parts).strip()


def paragraph_images(paragraph: Paragraph, document) -> list[str]:
    images: list[str] = []
    seen: set[str] = set()
    for blip in paragraph._p.xpath('.//a:blip'):
        rid = blip.get(qn('r:embed'))
        if not rid or rid in seen:
            continue
        seen.add(rid)
        part = document.part.related_parts.get(rid)
        if part is None or not getattr(part, "content_type", "").startswith("image/"):
            continue
        data = base64.b64encode(part.blob).decode("ascii")
        images.append(f'<figure class="document-figure"><img src="data:{part.content_type};base64,{data}" alt="Embedded diagram or document image" loading="lazy"></figure>')
    return images


def paragraph_heading_level(p: Paragraph) -> int | None:
    name = (p.style.name if p.style else "").strip().lower()
    match = re.fullmatch(r"heading\s*([1-6])", name)
    if match:
        return int(match.group(1))
    if name in {"title", "document title"}:
        return 1
    if name in {"subtitle", "document subtitle"}:
        return 2
    return None


def list_type(p: Paragraph) -> str | None:
    """Read real Word numbering metadata first; style names are the fallback."""
    ppr = p._p.pPr
    numpr = ppr.numPr if ppr is not None else None
    if numpr is not None and numpr.numId is not None:
        # Resolve abstract numbering format (bullet vs decimal/alpha/roman).
        try:
            num_id = int(numpr.numId.val)
            numbering = p.part.numbering_part.element
            for num in numbering.findall(qn('w:num')):
                if int(num.get(qn('w:numId'))) == num_id:
                    abstract_id = int(num.find(qn('w:abstractNumId')).get(qn('w:val')))
                    for abstract in numbering.findall(qn('w:abstractNum')):
                        if int(abstract.get(qn('w:abstractNumId'))) == abstract_id:
                            fmt = abstract.find('.//' + qn('w:numFmt'))
                            if fmt is not None and fmt.get(qn('w:val')) != 'bullet':
                                return 'ol'
                            return 'ul'
        except (AttributeError, TypeError, ValueError):
            pass
    name = (p.style.name if p.style else "").lower()
    if "bullet" in name:
        return "ul"
    if "number" in name:
        return "ol"
    return None


def is_code_paragraph(p: Paragraph) -> bool:
    text = normalize_breaks(p.text).strip()
    if not text:
        return False
    style = (p.style.name if p.style else "").lower()
    if "code" in style or "source" in style:
        return True
    fonts = [r.font.name.lower() for r in p.runs if r.font.name]
    if fonts and sum(f in {"consolas", "courier new", "cascadia code", "source code pro"} for f in fonts) / len(fonts) >= .6:
        return True
    patterns = [
        r"^(?:from\s+\S+\s+import\s+|import\s+\S+)", r"^(?:async\s+)?def\s+\w+\s*\(", r"^class\s+\w+",
        r"^(?:if|elif|else|try|except|finally|for|while|with|return|raise|pass|break|continue)\b",
        r"^@[\w.]+", r"^[\w.\[\]]+\s*=(?!=)", r"^[\w.]+\([^)]*\)\s*[,;]?$",
        r"^[)\]}]+[,;]?$", r"^[,;]$", r"^(?:authorizationUrl|tokenUrl|data|payload|token_data|resp|oauth2_scheme)\s*=",
        r"^#\s*(?:core/|components/|.*\.py)", r"^[│|├└].*(?:──|─)", r"^https?://\S+",
    ]
    return any(re.search(pat, text) for pat in patterns) or text.startswith(("    ", "\t"))


def paragraph_plain(p: Paragraph) -> str:
    return normalize_breaks(p.text or "").strip()


def render_table(table: Table, document) -> str:
    rows = []
    for ri, row in enumerate(table.rows):
        cells = []
        for ci, cell in enumerate(row.cells):
            # Keep paragraph/run formatting inside cells rather than flattening cell.text.
            content = []
            for p in cell.paragraphs:
                inline = inline_runs(p)
                if inline:
                    content.append(f"<p>{inline}</p>")
                content.extend(paragraph_images(p, document))
            for nested in cell.tables:
                content.append(render_table(nested, document))
            tag = "th" if ri == 0 else "td"
            scope = ' scope="col"' if ri == 0 else ""
            cells.append(f"<{tag}{scope}>{''.join(content) or '&nbsp;'}</{tag}>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return '<div class="table-wrapper"><table>' + "".join(rows) + "</table></div>"


def _is_markdown_heavy(text: str) -> bool:
    """Return True when a plain-text block is substantially Markdown-formatted.

    This catches DOCX files where the entire document body was typed as Markdown
    in Normal-style paragraphs rather than using real Word heading/table/list styles.
    """
    if not text or len(text) < 80:
        return False
    lines = text.splitlines()
    md_lines = sum(
        1 for line in lines
        if re.match(r"^\s*(#{1,6}\s+|```|~~~|[-*+]\s+|\d+[.)]\s+|>\s+|\|.+\|)", line)
    )
    # Treat as Markdown-heavy if >12 % of non-blank lines look like Markdown markup,
    # or the text has 3+ fenced code blocks, or 3+ Markdown-table rows.
    non_blank = max(1, sum(1 for l in lines if l.strip()))
    fence_count = len(re.findall(r"(?m)^```", text))
    table_rows = len(re.findall(r"(?m)^\|.+\|", text))
    return (md_lines / non_blank) > 0.12 or fence_count >= 3 or table_rows >= 3


def _body_blocks(document) -> list[dict]:
    blocks: list[dict] = []
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            p = Paragraph(child, document)
            text = paragraph_plain(p)
            images = paragraph_images(p, document)
            level = paragraph_heading_level(p)
            if level and text:
                blocks.append({"kind": "heading", "level": level, "html": f"<h{level}>{inline_runs(p)}</h{level}>", "text": text})
                blocks.extend({"kind": "html", "html": im} for im in images)
                continue
            if not text and not images:
                blocks.append({"kind": "blank", "html": "", "text": ""})
                continue
            list_tag = list_type(p)
            # Some source docs type list markers manually instead of using Word's numbering feature.
            manual_list = re.match(r"^\s*(?:([•●▪◦])\s+|(\d{1,2})[.)]\s+)", text)
            if not list_tag and manual_list:
                list_tag = "ol" if manual_list.group(2) else "ul"
                text_without_marker = re.sub(r"^\s*(?:[•●▪◦]\s+|\d{1,2}[.)]\s+)", "", text, count=1)
                # Preserve rich runs when the marker is a separate run; otherwise remove it from the first rendered text.
                inline = inline_runs(p)
                inline = re.sub(r"^\s*(?:[•●▪◦]\s+|\d{1,2}[.)]\s+)", "", inline, count=1)
            else:
                inline = inline_runs(p)
            # Promote only clearly standalone, short bold labels. Do not convert whole bold sentences.
            bold_only = bool(p.runs) and all((not r.text.strip()) or r.bold for r in p.runs)
            short_label = len(text) <= 90 and not re.search(r"[.!?]$", text) and not text.endswith((",", ";"))
            explicit_code = "code" in ((p.style.name if p.style else "").lower())
            code_line = is_code_paragraph(p)
            # A code sample is sometimes stored as one Word paragraph with literal <br> markers.
            # Keep all lines together instead of allowing its final bracket to become prose.
            if code_line and ("\n" in text or explicit_code):
                blocks.append({"kind": "codeblock", "html": "", "text": text, "code": True, "explicit_code": explicit_code})
            elif not list_tag and bold_only and short_label and text:
                blocks.append({"kind": "subheading", "html": f'<h3 class="document-subheading">{inline}</h3>', "text": text})
            elif list_tag:
                blocks.append({"kind": "list", "tag": list_tag, "html": inline, "text": text})
            else:
                # --- STYLING FIX: detect Markdown-heavy Normal paragraphs ---
                # Some DOCX files store raw multi-line Markdown text in Normal-style
                # paragraphs. Detect this and mark the block for Markdown rendering.
                if "\n" in text and _is_markdown_heavy(text):
                    blocks.append({"kind": "markdown_block", "html": "", "text": text})
                else:
                    blocks.append({"kind": "paragraph", "html": f"<p>{inline}</p>" if inline else "", "text": text,
                                   "code": code_line, "explicit_code": explicit_code})
            blocks.extend({"kind": "html", "html": im} for im in images)
        elif child.tag == qn("w:tbl"):
            blocks.append({"kind": "html", "html": render_table(Table(child, document), document), "text": ""})
    return blocks


def _balanced_open(text: str) -> int:
    # Approximate bracket balance after stripping string literals/comments.
    cleaned = re.sub(r"(['\"]).*?\1", "", text)
    cleaned = cleaned.split("#", 1)[0]
    return sum(cleaned.count(a) - cleaned.count(b) for a, b in [("(", ")"), ("[", "]"), ("{", "}")])


def _code_block(blocks: list[dict], start: int) -> tuple[str, int] | None:
    first = blocks[start]
    if first.get("kind") != "paragraph" or not first.get("code"):
        return None
    lines = [first["text"]]
    j = start + 1
    balance = _balanced_open(first["text"])
    # A single code-like paragraph is not enough unless Word explicitly marked it as code.
    while j < len(blocks):
        b = blocks[j]
        if b["kind"] == "blank":
            if j + 1 < len(blocks) and blocks[j + 1].get("code"):
                lines.append("")
                j += 1
                continue
            break
        if b["kind"] != "paragraph":
            break
        text = b["text"]
        stripped = text.strip()
        continuation = balance > 0 or stripped.startswith(("'", '"', ")", "]", "}", ",", ".", "#")) or text.startswith(("    ", "\t"))
        if b.get("code") or continuation:
            lines.append(text)
            balance += _balanced_open(text)
            j += 1
            # A long statement that has balanced should still collect adjacent obvious code lines.
            continue
        break
    if len([line for line in lines if line.strip()]) < 2 and not first.get("explicit_code"):
        return None
    code = "\n".join(lines).rstrip()
    language = "python" if re.search(r"(?m)^\s*(from\s+\S+\s+import|import\s+\S+|def\s+|class\s+)", code) else "text"
    return f'<pre><code class="language-{language}">{esc(code)}</code></pre>', j


def docx_to_html(path: Path) -> str:
    doc = Document(path)
    blocks = _body_blocks(doc)

    # --- STYLING FIX: if the whole document is essentially one big Markdown block,
    # route everything through the Markdown pipeline for proper rendering. ---
    all_text_blocks = [b for b in blocks if b.get("text")]
    md_blocks = [b for b in all_text_blocks if b.get("kind") == "markdown_block"]
    if md_blocks and len(md_blocks) / max(1, len(all_text_blocks)) > 0.4:
        combined = "\n\n".join(b["text"] for b in blocks if b.get("text"))
        return markdown_to_html(combined)

    out: list[str] = []
    i = 0
    active_list: str | None = None
    list_items: list[str] = []

    def flush_list():
        nonlocal active_list, list_items
        if active_list and list_items:
            out.append(f"<{active_list}>" + "".join(list_items) + f"</{active_list}>")
        active_list, list_items = None, []

    while i < len(blocks):
        b = blocks[i]
        if b["kind"] == "list":
            if active_list != b["tag"]:
                flush_list()
                active_list = b["tag"]
            list_items.append(f"<li>{b['html']}</li>")
            i += 1
            continue
        flush_list()
        if b["kind"] == "blank":
            i += 1
            continue
        if b["kind"] == "codeblock":
            code = esc(b["text"]).replace("\n", "\n")
            out.append(f'<pre><code class="language-python">{code}</code></pre>')
            i += 1
            continue
        # --- STYLING FIX: render individual Markdown-heavy paragraphs ---
        if b["kind"] == "markdown_block":
            out.append(markdown_to_html(b["text"]))
            i += 1
            continue
        code = _code_block(blocks, i)
        if code:
            rendered, next_i = code
            out.append(rendered)
            i = next_i
            continue
        if b.get("html"):
            out.append(b["html"])
        i += 1

    flush_list()
    return "\n".join(out)


def clean_source(text: str) -> str:
    text = (text or "").lstrip("\ufeff").replace("\r\n", "\n").replace("\r", "\n").replace("\x00", "")
    if "\n" not in text and text.count(r"\n") >= 2:
        text = text.replace(r"\n", "\n").replace(r"\t", "\t")
    return text


def repair_truncated_markdown(text: str) -> str:
    """Fix Markdown that was cut off mid-table by Word's paragraph character limit.

    Word stores each document's body as a single large paragraph and silently
    truncates the text when it exceeds an internal limit.  This leaves the
    Markdown with incomplete table rows at the end (e.g. a header row with no
    separator, or a data row whose last cell is not closed).  The python-markdown
    `tables` extension refuses to render any of those rows, so the whole table
    reverts to raw pipe-delimited text.

    Strategy:
    - Strip any trailing bare `|` fragments that are clearly incomplete.
    - If the last non-blank line is a table header row (contains `|`) with no
      following separator line, inject a separator so the header still renders.
    - If the last line looks like a valid-but-incomplete data row, close it and
      add a truncation note so the reader knows content was cut.
    """
    if not text:
        return text

    lines = text.splitlines()

    # Drop trailing lines that are just a lone `|` or whitespace-only pipe fragments.
    while lines and re.fullmatch(r"\s*\|?\s*", lines[-1]):
        lines.pop()

    if not lines:
        return text

    last = lines[-1]

    # Detect a table row: starts (and ideally ends) with `|`, or contains `|...|`.
    def is_table_row(line: str) -> bool:
        s = line.strip()
        return bool(s) and "|" in s and (s.startswith("|") or s.endswith("|"))

    def is_separator_row(line: str) -> bool:
        s = line.strip()
        return bool(re.fullmatch(r"[\|\-\:\s]+", s)) and "-" in s

    # If last non-blank line is a table row (not a separator), check context.
    if is_table_row(last) and not is_separator_row(last):
        # Count pipe-columns to build a matching separator.
        cols = max(1, last.count("|") - 1) if last.startswith("|") else last.count("|") + 1

        # Case 1: The last line IS the header row and there is no separator yet.
        # Check if the line before last (if any) is also a table row.
        prev_is_table = len(lines) >= 2 and is_table_row(lines[-2])
        prev_is_sep = len(lines) >= 2 and is_separator_row(lines[-2])

        if not prev_is_table and not prev_is_sep:
            # Lone header row: inject separator so the header renders.
            sep = "| " + " | ".join(["---"] * cols) + " |"
            lines.append(sep)
            lines.append("| *(table data was truncated)* |" + " |" * (cols - 1))
        else:
            # Data row that was cut off mid-content: close it and note truncation.
            if not last.strip().endswith("|"):
                lines[-1] = last.rstrip() + "… |"
            lines.append("| *(content truncated)* |" + " |" * (cols - 1))

    return "\n".join(lines)


def markdown_to_html(source: str) -> str:
    source = clean_source(source)
    source = repair_truncated_markdown(source)
    rendered = MD.convert(source)
    MD.reset()
    rendered = re.sub(r"<table(.*?)>(.*?)</table>", r'<div class="table-wrapper"><table\1>\2</table></div>', rendered, flags=re.I | re.S)
    return bleach.clean(rendered, tags=ALLOWED_TAGS, attributes=ALLOWED_ATTRIBUTES,
                        protocols={"http", "https", "mailto", "data"}, strip=True)


def is_markdown_like(text: str) -> bool:
    return bool(re.search(r"(?m)^\s*(#{1,6}\s+|```|~~~|[-*+]\s+|\d+[.)]\s+|>\s+)|\|.+\|", text or ""))


def pdf_to_html(path: Path) -> str:
    sections = []
    with pymupdf.open(path) as pdf:
        for n, page in enumerate(pdf, 1):
            text = clean_source(page.get_text("text")).strip()
            sections.append(f'<section class="pdf-page"><div class="page-label">Page {n}</div>')
            if not text:
                sections.append('<p class="conversion-note">No selectable text was found on this page. Scanned pages may require OCR.</p>')
            elif is_markdown_like(text):
                sections.append(markdown_to_html(text))
            else:
                for para in re.split(r"\n\s*\n", text):
                    if para.strip():
                        sections.append("<p>" + esc(para.strip()).replace("\n", "<br>") + "</p>")
            sections.append("</section>")
    return "\n".join(sections)


def convert_document(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".docx":
        return docx_to_html(path)
    if ext == ".pdf":
        return pdf_to_html(path)
    if ext in {".md", ".markdown"}:
        return markdown_to_html(path.read_text(encoding="utf-8-sig", errors="replace"))
    if ext == ".txt":
        text = clean_source(path.read_text(encoding="utf-8-sig", errors="replace"))
        if is_markdown_like(text):
            return markdown_to_html(text)
        return "".join("<p>" + esc(p).replace("\n", "<br>") + "</p>" for p in re.split(r"\n\s*\n", text) if p.strip())
    raise ValueError(f"Unsupported file type: {ext}")


def _cache_key(path: Path) -> str:
    """Stable cache filename that includes the source file's mtime so stale
    converted HTML is automatically invalidated when the source changes."""
    mtime = int(path.stat().st_mtime)
    tag = hashlib.md5(f"{path.name}:{mtime}".encode()).hexdigest()[:10]
    return secure_filename(path.name) + f".{tag}.html"


def get_documents():
    """Scan uploads/, convert each file (using a mtime-keyed cache), and return
    a stable list ordered by filename.  Each document carries a hash-based ``id``
    so adding/removing files never reassigns existing IDs."""
    documents = []
    for path in sorted(UPLOAD_FOLDER.iterdir(), key=lambda p: p.name.casefold()):
        if not path.is_file() or path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        try:
            html_name = _cache_key(path)
            cache_path = CONVERTED_FOLDER / html_name
            if not cache_path.exists():
                # Evict any old cache files for this source name before writing the new one.
                stem = secure_filename(path.name)
                for old in CONVERTED_FOLDER.glob(f"{stem}.*.html"):
                    try:
                        old.unlink()
                    except OSError:
                        pass
                cache_path.write_text(convert_document(path), encoding="utf-8")
            # Stable ID: first 12 chars of the MD5 of the filename (case-folded).
            doc_id = hashlib.md5(path.name.casefold().encode()).hexdigest()[:12]
            documents.append({
                "id": doc_id,
                "name": path.stem.replace("_", " "),
                "filename": path.name,
                "html_file": html_name,
            })
        except Exception:
            app.logger.exception("Failed to convert %s", path.name)
    return documents


@app.route("/")
def home():
    return render_template("index.html", documents=get_documents())


@app.route("/document/<doc_id>")
def show_document(doc_id: str):
    docs = get_documents()
    doc = next((d for d in docs if d["id"] == doc_id), None)
    if doc is None:
        abort(404)
    output_path = CONVERTED_FOLDER / doc["html_file"]
    return Response(output_path.read_text(encoding="utf-8"), content_type="text/html; charset=utf-8")


if __name__ == "__main__":
    app.run(debug=True)
