# AI Solution Showcase — clean rebuild

This rebuild keeps the existing Flask document-library UI and replaces the fragile DOCX conversion path with a structured converter based on Word paragraphs, runs, numbering, tables, and embedded image relationships.

## Run locally

1. Extract this folder into a new directory. Do not extract over your existing project.
2. Open the folder in VS Code.
3. (Recommended) Create a virtual environment:
   - Windows: `py -m venv .venv`
   - Activate: `.venv\\Scripts\\activate`
4. Install dependencies: `python -m pip install -r requirements.txt`
5. Start the app: `python app.py`
6. Open `http://127.0.0.1:5000`.

## Documents and recovery

- `uploads/` contains the active documents displayed by the app. The latest supplied Architecture, Code, and Traceability documents are installed as `ARCHITECTURE.docx`, `CODE.docx`, and `TRACEABILITY.docx`.
- `original_documents/` contains unchanged copies of the three source files named `updated_*.docx` for recovery/reference. Do not delete these until you have checked the preview.
- Add new supported documents (`.docx`, `.pdf`, `.md`, `.markdown`, `.txt`) to `uploads/` and refresh the app.
- `converted/` is generated output. It can be cleared safely; the app rebuilds it from `uploads/` on each request.

## Conversion behavior

- Word heading styles map to matching HTML heading levels.
- Bold, italic, underline, and inline monospace runs are retained.
- Word numbered/bulleted lists and manually typed list markers are rendered as lists.
- Native Word tables retain inline formatting in cells and scroll horizontally when wide.
- Embedded images/diagrams are emitted as responsive figures using self-contained data URIs.
- Code samples are grouped conservatively. Unusual code formatting may still depend on how the source Word document represents the code.
- Markdown fences, tables, lists, and Mermaid source are supported; Mermaid is rendered in the browser when its CDN is available.

## Troubleshooting

- If styles appear stale, stop Flask (`Ctrl+C`), restart `python app.py`, then hard-refresh the browser (`Ctrl+F5`).
- If a document fails, check the terminal for the conversion error. Keep `original_documents/` unchanged while debugging.
