from __future__ import annotations

from pathlib import Path
import html
import re

from flask import Flask, abort, render_template, Response
from werkzeug.utils import secure_filename
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
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

# Markdown extensions: tables, fenced code, task lists, TOC, footnotes, etc.
MD = markdown.Markdown(extensions=[
    "extra", "tables", "fenced_code", "sane_lists", "toc", "footnotes", "nl2br"
], output_format="html5")

ALLOWED_TAGS = set(bleach.sanitizer.ALLOWED_TAGS).union({
    "p", "h1", "h2", "h3", "h4", "h5", "h6", "table", "thead", "tbody",
    "tfoot", "tr", "th", "td", "pre", "code", "hr", "br", "div", "span",
    "img", "figure", "figcaption", "details", "summary", "del", "sup", "sub",
    "input", "section", "blockquote", "dl", "dt", "dd", "abbr", "kbd", "mark"
})
ALLOWED_ATTRIBUTES = {
    **bleach.sanitizer.ALLOWED_ATTRIBUTES,
    "*": ["id", "class", "title", "aria-label"],
    "a": ["href", "title", "rel"],
    "img": ["src", "alt", "title", "width", "height"],
    "th": ["colspan", "rowspan", "align"],
    "td": ["colspan", "rowspan", "align"],
    "input": ["type", "checked", "disabled"],
    "code": ["class"], "pre": ["class"], "div": ["class", "id"],
    "span": ["class", "id"],
}


def clean_source(text: str) -> str:
    """Normalize line endings and repair common Markdown fences from DOCX exports."""
    text = text.lstrip("\ufeff").replace("\r\n", "\n").replace("\r", "\n")
    if "\n" not in text and text.count(r"\n") >= 2:
        text = text.replace(r"\n", "\n").replace(r"\t", "\t")
    text = text.replace("\x00", "")
    # Standardize indented or oversized code fences so Markdown parses them consistently.
    fixed = []
    for line in text.split("\n"):
        match = re.match(r"^\s*(`{3,}|~{3,})\s*([A-Za-z0-9_+.-]+)?\s*$", line)
        if match:
            fence, language = match.group(1), (match.group(2) or "")
            fixed.append(("```" if fence[0] == "`" else "~~~") + language.lower())
        else:
            fixed.append(line)
    return "\n".join(fixed)


def is_markdown_like(text: str) -> bool:
    checks = [
        r"(?m)^\s{0,3}#{1,6}\s+", r"(?m)^\s*```", r"(?m)^\s*~~~",
        r"(?m)^\s*\|?.+\|.+\|?\s*$", r"(?m)^\s*>\s+",
        r"(?m)^\s*(?:[-*+]\s+|\d+[.)]\s+)", r"(?m)^\s*[-*_]{3,}\s*$",
    ]
    return any(re.search(pattern, text) for pattern in checks)


def markdown_to_html(source: str) -> str:
    source = clean_source(source)

    # Some DOCX exports contain an opening code fence but omit the closing fence.
    # If the total number of standalone fences is odd, close the last one at EOF.
    fence_lines = re.findall(r"(?m)^\s*(`{3,}|~{3,})[^\n]*$", source)
    if len(fence_lines) % 2 == 1:
        last_fence = fence_lines[-1][0][0]
        source = source.rstrip() + "\n" + ("```" if last_fence == "`" else "~~~") + "\n"

    # Protect fenced code blocks before Markdown parsing. This makes SQL, Python,
    # JSON, YAML and Mermaid blocks robust even when DOCX exports have odd spacing.
    fenced_blocks: list[str] = []

    def stash_fence(match: re.Match) -> str:
        fence_text = match.group(0)
        first_line = fence_text.splitlines()[0].strip()
        language = first_line[3:].strip().split()[0].lower() if len(first_line) > 3 else ""
        code_lines = fence_text.splitlines()[1:]
        if code_lines and re.match(r"^\s*(```|~~~)\s*$", code_lines[-1]):
            code_lines = code_lines[:-1]
        code = "\n".join(code_lines).strip("\n")
        escaped = html.escape(code)
        lang_class = f' class="language-{html.escape(language, quote=True)}"' if language else ""
        block = f"<pre><code{lang_class}>{escaped}</code></pre>"
        token = f"FENCEDCODEBLOCKPLACEHOLDER{len(fenced_blocks)}END"
        fenced_blocks.append(block)
        return f"\n\n{token}\n\n"

    source = re.sub(
        r"(?ms)^[ \t]*(```+|~~~+)[ \t]*([A-Za-z0-9_+.-]*)[^\n]*\n.*?^[ \t]*\1[ \t]*$",
        lambda m: stash_fence(m),
        source,
    )

    rendered = MD.convert(source)
    MD.reset()
    for index, block in enumerate(fenced_blocks):
        rendered = rendered.replace(f"<p>FENCEDCODEBLOCKPLACEHOLDER{index}END</p>",
                                    block)
        rendered = rendered.replace(f"FENCEDCODEBLOCKPLACEHOLDER{index}END", block)

    rendered = re.sub(
        r"<table(.*?)>(.*?)</table>",
        r'<div class="table-wrapper"><table\1>\2</table></div>',
        rendered, flags=re.I | re.S
    )
    return bleach.clean(
        rendered, tags=ALLOWED_TAGS, attributes=ALLOWED_ATTRIBUTES,
        protocols={"http", "https", "mailto", "data"}, strip=True
    )


def native_table_to_html(table: Table) -> str:
    rows = []
    for ri, row in enumerate(table.rows):
        cells = []
        for cell in row.cells:
            cell_content = "<br>".join(html.escape(line) for line in cell.text.splitlines())
            tag = "th" if ri == 0 else "td"
            cells.append(f"<{tag}>{cell_content}</{tag}>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return '<div class="table-wrapper"><table>' + "".join(rows) + "</table></div>"


def docx_to_html(path: Path) -> str:
    """Read DOCX paragraphs and actual Word tables in their original order.

    Many generated DOCX files contain an entire Markdown document in one or two
    paragraphs. Those paragraphs are parsed as Markdown, not flattened into plain text.
    """
    doc = Document(path)
    body_parts: list[str] = []
    paragraph_buffer: list[str] = []

    def flush_paragraphs() -> None:
        if not paragraph_buffer:
            return
        block = "\n\n".join(p for p in paragraph_buffer if p is not None)
        paragraph_buffer.clear()
        if not block.strip():
            return
        # One Markdown document stored in a Word paragraph should be parsed together,
        # preserving all embedded headings, fences, tables, lists and Mermaid blocks.
        body_parts.append(markdown_to_html(block))

    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            paragraph_buffer.append(Paragraph(child, doc).text or "")
        elif child.tag.endswith("}tbl"):
            flush_paragraphs()
            body_parts.append(native_table_to_html(Table(child, doc)))
    flush_paragraphs()
    return "\n".join(part for part in body_parts if part.strip())


def pdf_to_html(path: Path) -> str:
    pages = []
    with pymupdf.open(path) as pdf:
        for page_number, page in enumerate(pdf, start=1):
            text = clean_source(page.get_text("text")).strip()
            pages.append(f'<section class="pdf-page"><div class="page-label">Page {page_number}</div>')
            if not text:
                pages.append('<p class="conversion-note">No selectable text found on this page. Scanned PDFs need OCR.</p>')
            elif is_markdown_like(text):
                pages.append(markdown_to_html(text))
            else:
                # Preserve paragraphs/line breaks and avoid interpreting arbitrary PDF text as HTML.
                paragraphs = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
                pages.append("".join("<p>" + html.escape(p).replace("\n", "<br>") + "</p>" for p in paragraphs))
            pages.append("</section>")
    return "\n".join(pages)


def convert_document(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".docx":
        return docx_to_html(path)
    if ext == ".pdf":
        return pdf_to_html(path)
    if ext in {".md", ".markdown"}:
        return markdown_to_html(path.read_text(encoding="utf-8-sig", errors="replace"))
    if ext == ".txt":
        source = path.read_text(encoding="utf-8-sig", errors="replace")
        if is_markdown_like(source):
            return markdown_to_html(source)
        return "".join("<p>" + html.escape(p).replace("\n", "<br>") + "</p>"
                       for p in re.split(r"\n\s*\n", clean_source(source)) if p.strip())
    raise ValueError(f"Unsupported file type: {ext}")


def get_documents():
    documents = []
    for path in sorted(UPLOAD_FOLDER.iterdir(), key=lambda p: p.name.casefold()):
        if not path.is_file() or path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        try:
            html_name = secure_filename(path.name) + ".html"
            output_path = CONVERTED_FOLDER / html_name
            output_path.write_text(convert_document(path), encoding="utf-8")
            documents.append({
                "id": len(documents), "name": path.stem.replace("_", " "),
                "filename": path.name, "html_file": html_name
            })
        except Exception:
            app.logger.exception("Failed to convert %s", path.name)
    return documents


@app.route("/")
def home():
    return render_template("index.html", documents=get_documents())


@app.route("/document/<int:doc_id>")
def show_document(doc_id: int):
    docs = get_documents()
    if not 0 <= doc_id < len(docs):
        abort(404)
    output_path = CONVERTED_FOLDER / docs[doc_id]["html_file"]
    if not output_path.is_file():
        abort(404)
    return Response(output_path.read_text(encoding="utf-8"), content_type="text/html; charset=utf-8")


if __name__ == "__main__":
    app.run(debug=True)
