# AI Solution Showcase

## Run
1. Open this folder in VS Code.
2. Create/activate a virtual environment if desired.
3. Install dependencies: `python -m pip install -r requirements.txt`
4. Copy `.pdf`, `.docx`, `.md`, `.markdown`, or `.txt` documents into `uploads/`.
5. Start: `python app.py`
6. Visit http://127.0.0.1:5000

## Conversion approach
Markdown source embedded in DOCX paragraphs is parsed as Markdown. Markdown tables, fenced code blocks, lists, headings, quotes and Mermaid fences are preserved. Native Word tables are rendered as HTML tables. The converter deliberately does not guess/rewrite code or Mermaid syntax because heuristic rewriting can corrupt valid source. If a Mermaid diagram still fails, the underlying Mermaid source itself may be invalid and is shown in the preview fallback.


## Client presentation UI (v9)
The main page uses a full-width hero, searchable document cards, file-type filters, and a document preview panel. Start with `python app.py` and open `http://127.0.0.1:5000/`.


## UI fix v11
The HTML `hidden` attribute is now enforced in CSS, so the welcome placeholder disappears when a document is opened. Selecting a document smoothly scrolls to the document preview on desktop and mobile.


## Mermaid rendering fix (v12)
The client renderer normalizes typographic punctuation from Word and strips embedded Mermaid theme-init directives that can break rendering. Diagrams use a shared theme and responsive SVG styles.


## Mermaid auto-repair (v13)
When a diagram fails on the first render, the UI retries after normalizing whitespace and quoting subgraph/node labels, then displays the actual render error with source if both attempts fail.
