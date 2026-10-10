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
