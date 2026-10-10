function normalizeMermaidSource(raw) {
  // Word documents often contain typographic quotes that Mermaid treats as syntax.
  let source = (raw || "")
    .replace(/\\r/g, "")
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
    .replace(/[‐‑‒–—]/g, "-")
    .trim();

  // Mermaid node labels containing quotes should use normal ASCII quotes.
  // Strip accidental Markdown fence remnants if a document exporter included them.
  source = source.replace(/^\\s*```(?:mermaid)?\\s*/i, "")
                 .replace(/\\s*```\\s*$/, "");
  return source;
}

document.addEventListener("DOMContentLoaded", () => {
  const buttons = document.querySelectorAll(".document-button");
  const title = document.getElementById("document-title");
  const typeLabel = document.getElementById("document-type");
  const welcome = document.getElementById("welcome");
  const loading = document.getElementById("loading");
  const content = document.getElementById("document-content");
  const errorMessage = document.getElementById("error-message");
  let currentRequest = 0;

  if (window.mermaid) {
    mermaid.initialize({ startOnLoad: false, theme: "default", securityLevel: "strict", suppressErrorRendering: true });
  }

  buttons.forEach((button) => {
    button.addEventListener("click", async () => {
      const requestId = ++currentRequest;
      const docId = button.dataset.id;
      buttons.forEach((item) => item.classList.remove("active"));
      button.classList.add("active");
      title.textContent = button.querySelector(".file-name").textContent;
      typeLabel.textContent = (button.title.split(".").pop() || "doc").toUpperCase();
      welcome.hidden = true;
      content.hidden = true;
      errorMessage.hidden = true;
      loading.hidden = false;

      try {
        const response = await fetch("/document/" + encodeURIComponent(docId));
        if (!response.ok) throw new Error("Server returned HTTP " + response.status);
        const htmlContent = await response.text();
        if (requestId !== currentRequest) return;
        content.innerHTML = htmlContent;
        content.hidden = false;
        if (window.hljs) {
          content.querySelectorAll("pre code").forEach((block) => {
            try { window.hljs.highlightElement(block); } catch (e) { /* plain code remains readable */ }
          });
        }

        // Render Mermaid diagrams one by one. If a diagram has invalid syntax,
        // show its source instead of replacing the whole preview with a bomb/error.
        if (window.mermaid) {
          const blocks = content.querySelectorAll("pre code.language-mermaid, pre code.lang-mermaid, code.language-mermaid, code.lang-mermaid");
          for (let index = 0; index < blocks.length; index++) {
            const code = blocks[index];
            const source = normalizeMermaidSource(code.textContent);
            const diagram = document.createElement("div");
            diagram.className = "mermaid";
            diagram.textContent = source;
            diagram.id = "mermaid-diagram-" + requestId + "-" + index;
            const pre = code.closest("pre");
            if (pre) pre.replaceWith(diagram); else code.replaceWith(diagram);
            try {
              await mermaid.run({ nodes: [diagram] });
              if (diagram.querySelector(".error-icon") ||
                  /Syntax error in text/i.test(diagram.textContent || "")) {
                throw new Error("Mermaid reported a syntax error");
              }
            } catch (diagramError) {
              console.error("Mermaid syntax error in " + title.textContent + ":", diagramError, source);
              const fallback = document.createElement("details");
              fallback.className = "diagram-source-fallback";
              const summary = document.createElement("summary");
              summary.textContent = "Diagram could not render — view source";
              const preSource = document.createElement("pre");
              const codeSource = document.createElement("code");
              codeSource.textContent = source;
              preSource.appendChild(codeSource);
              fallback.append(summary, preSource);
              diagram.replaceWith(fallback);
            }
          }
        }
      } catch (error) {
        if (requestId !== currentRequest) return;
        errorMessage.textContent = "Unable to display this document. Check the VS Code terminal and browser console.";
        errorMessage.hidden = false;
        console.error(error);
      } finally {
        if (requestId === currentRequest) loading.hidden = true;
      }
    });
  });
});
