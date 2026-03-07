const BIBLIOGRAPHY_TITLE = "Список використаних джерел";

export function buildBibliographyParagraphs(sources = []) {
  if (!Array.isArray(sources) || sources.length === 0) {
    return [
      `<w:p><w:r><w:t>${BIBLIOGRAPHY_TITLE}</w:t></w:r></w:p>`,
      "<w:p><w:r><w:t>[1] Додайте бібліографічні джерела у bibliography.js</w:t></w:r></w:p>"
    ];
  }

  const items = sources.map(
    (source, index) => `<w:p><w:r><w:t>[${index + 1}] ${escapeXml(source)}</w:t></w:r></w:p>`
  );

  return [`<w:p><w:r><w:t>${BIBLIOGRAPHY_TITLE}</w:t></w:r></w:p>`, ...items];
}

export function insertBibliography(documentXml, sources = []) {
  const paragraphs = buildBibliographyParagraphs(sources).join("");
  return documentXml.replace("</w:body>", `${paragraphs}</w:body>`);
}

export function attachCitationReferences(documentXml, citations = []) {
  if (!Array.isArray(citations) || citations.length === 0) {
    return documentXml;
  }

  const citationText = citations.map((value) => `[${value}]`).join(" ");
  const citationParagraph = `<w:p><w:r><w:t>Цитування: ${escapeXml(citationText)}</w:t></w:r></w:p>`;
  return documentXml.replace("</w:body>", `${citationParagraph}</w:body>`);
}

function escapeXml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}
