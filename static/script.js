function normalizeMermaidSource(raw) {
  let source = (raw || "")
    .replace(/\r/g, "")
    .replace(/[“”„‟]/g, '"')
    .replace(/[‘’‚‛]/g, "'")
    .replace(/[‐‑‒–—―−]/g, "-")
    .replace(/\u00a0/g, " ")
    .trim();

  // Word-exported Mermaid often includes a theme directive that causes failures
  // in browser-rendered previews. Use the app's shared Mermaid theme instead.
  source = source.replace(/^\s*%%\{init:[\s\S]*?\}%%\s*/i, "");
  source = source.replace(/^\s*```(?:mermaid)?\s*/i, "")
                 .replace(/\s*```\s*$/, "");

  // Replace invisible control characters sometimes introduced by DOCX copy/paste.
  source = source.replace(/[\u200B-\u200D\uFEFF]/g, "");
  return source.trim();
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

  const searchInput = document.getElementById("document-search");
  const filterButtons = document.querySelectorAll(".filter-chip");
  const noSearchResults = document.getElementById("no-search-results");
  let activeFilter = "all";

  function filterDocumentCards() {
    const query = (searchInput?.value || "").trim().toLowerCase();
    let visibleCount = 0;
    buttons.forEach((button) => {
      const matchesName = (button.dataset.name || button.querySelector(".file-name").textContent.toLowerCase()).includes(query);
      const matchesType = activeFilter === "all" || button.dataset.kind === activeFilter;
      const visible = matchesName && matchesType;
      button.hidden = !visible;
      if (visible) visibleCount++;
    });
    if (noSearchResults) noSearchResults.hidden = visibleCount !== 0;
  }

  if (searchInput) searchInput.addEventListener("input", filterDocumentCards);
  filterButtons.forEach((filterButton) => {
    filterButton.addEventListener("click", () => {
      activeFilter = filterButton.dataset.filter || "all";
      filterButtons.forEach((button) => button.classList.toggle("active", button === filterButton));
      filterDocumentCards();
    });
  });


  if (window.mermaid) {
    mermaid.initialize({
      startOnLoad: false,
      theme: "neutral",
      securityLevel: "strict",
      suppressErrorRendering: true,
      flowchart: { htmlLabels: false, curve: "basis" },
      sequence: { useMaxWidth: true },
      themeVariables: {
        fontFamily: "Inter, Segoe UI, Arial, sans-serif",
        primaryColor: "#eef2ff",
        primaryTextColor: "#172554",
        primaryBorderColor: "#818cf8",
        lineColor: "#64748b",
        secondaryColor: "#f0fdfa",
        tertiaryColor: "#f8fafc"
      }
    });
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
        // Hide the welcome/loader state first, then bring the preview into view on every screen size.
        welcome.hidden = true;
        loading.hidden = true;
        requestAnimationFrame(() => {
          document.getElementById("document-viewer").scrollIntoView({
            behavior: "smooth",
            block: "start"
          });
        });
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
