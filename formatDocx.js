import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { bibliographySources, insertBibliography, insertCitationReferences } from "./bibliography.js";

const DOC_XML_PATH = "word/document.xml";
const PAGE_BREAK_PARAGRAPH = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

async function main() {
  const [inputPath, outputPath] = process.argv.slice(2);

  validateArgs(inputPath, outputPath);

  const inputBuffer = await fs.readFile(inputPath);
  const zip = await JSZip.loadAsync(inputBuffer);
  const documentFile = zip.file(DOC_XML_PATH);

  if (!documentFile) {
    throw new Error(`Файл ${DOC_XML_PATH} не знайдено у ${inputPath}`);
  }

  let documentXml = await documentFile.async("string");

  const analysis = analyzeDocument(documentXml);
  documentXml = applyFormatting(documentXml, analysis);
  documentXml = insertTableOfContents(documentXml);
  documentXml = insertCitationReferences(documentXml, bibliographySources);
  documentXml = insertBibliography(documentXml, bibliographySources);

  zip.file(DOC_XML_PATH, documentXml);
  const outputBuffer = await zip.generateAsync({ type: "nodebuffer" });

  await fs.mkdir(path.dirname(path.resolve(outputPath)), { recursive: true });
  await fs.writeFile(outputPath, outputBuffer);

  console.log(`Готово. Форматований файл збережено: ${outputPath}`);
  console.log(`Аналітика: заголовків ${analysis.headingsCount}, абзаців ${analysis.paragraphCount}.`);
}

function validateArgs(inputPath, outputPath) {
  if (!inputPath || !outputPath) {
    throw new Error("Використання: node formatDocx.js input.docx output.docx");
  }

  for (const filePath of [inputPath, outputPath]) {
    if (path.extname(filePath).toLowerCase() !== ".docx") {
      throw new Error(`Очікується розширення .docx: ${filePath}`);
    }
  }
}

function analyzeDocument(documentXml) {
  const headingsCount = (documentXml.match(/w:pStyle\s+w:val="Heading[1-6]"/g) || []).length;
  const paragraphCount = (documentXml.match(/<w:p\b/g) || []).length;
  return { headingsCount, paragraphCount };
}

function applyFormatting(documentXml, analysis) {
  let hasHeading = false;
  let previousWasPageBreak = false;

  let formattedXml = documentXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraphXml) => {
    const normalizedParagraph = normalizeParagraphText(paragraphXml);

    if (isRemovableEmptyParagraph(normalizedParagraph)) {
      return "";
    }

    const paragraphText = extractParagraphText(normalizedParagraph);
    const headingType = getHeadingType(paragraphText);
    const paragraphWithStyle = applyParagraphStyle(normalizedParagraph, headingType);

    hasHeading = hasHeading || headingType !== null;

    if (headingType === "Heading1") {
      const prefix = previousWasPageBreak ? "" : PAGE_BREAK_PARAGRAPH;
      previousWasPageBreak = false;
      return `${prefix}${paragraphWithStyle}`;
    }

    previousWasPageBreak = hasExplicitPageBreak(paragraphWithStyle);
    return paragraphWithStyle;
  });

  if (!hasHeading && analysis.headingsCount === 0) {
    const warningParagraph = "<w:p><w:r><w:t>Попередження: у документі не знайдено заголовків Heading 1-2.</w:t></w:r></w:p>";
    formattedXml = formattedXml.replace("</w:body>", `${warningParagraph}</w:body>`);
  }

  return setSectionMargins(formattedXml);
}

function normalizeParagraphText(paragraphXml) {
  return paragraphXml.replace(/(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/g, (full, open, text, close) => {
    const compact = text.replace(/\s{2,}/g, " ").trim();
    return `${open}${compact}${close}`;
  });
}

function extractParagraphText(paragraphXml) {
  const text = [...paragraphXml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)]
    .map((match) => decodeXmlText(match[1]))
    .join(" ")
    .replace(/\s{2,}/g, " ")
    .trim();

  return text;
}

function decodeXmlText(text) {
  return text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function getHeadingType(text) {
  const normalized = text.trim().toUpperCase();

  if (!normalized) {
    return null;
  }

  if (
    normalized === "ВСТУП" ||
    normalized === "ВИСНОВКИ" ||
    normalized === "СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ" ||
    /^РОЗДІЛ\s+\d+/.test(normalized)
  ) {
    return "Heading1";
  }

  if (/^\d+\.\d+/.test(normalized)) {
    return "Heading2";
  }

  return null;
}

function isRemovableEmptyParagraph(paragraphXml) {
  if (hasExplicitPageBreak(paragraphXml)) {
    return false;
  }

  const text = extractParagraphText(paragraphXml);
  return text.length === 0;
}

function hasExplicitPageBreak(paragraphXml) {
  return /<w:br\b[^>]*w:type="page"/.test(paragraphXml);
}

function applyParagraphStyle(paragraphXml, headingType) {
  const updatedPPr = buildPPr(paragraphXml, headingType);

  if (/<w:pPr\b[\s\S]*?<\/w:pPr>/.test(paragraphXml)) {
    return paragraphXml.replace(/<w:pPr\b[\s\S]*?<\/w:pPr>/, updatedPPr);
  }

  return paragraphXml.replace(/<w:p\b([^>]*)>/, `<w:p$1>${updatedPPr}`);
}

function buildPPr(paragraphXml, headingType) {
  const existingPPrMatch = paragraphXml.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/);
  let pPr = existingPPrMatch ? existingPPrMatch[0] : "<w:pPr></w:pPr>";

  pPr = removeTag(pPr, "w:pStyle");

  if (headingType) {
    pPr = upsertTag(pPr, "w:pStyle", `<w:pStyle w:val="${headingType}"/>`);
    return pPr;
  }

  pPr = upsertTag(pPr, "w:jc", '<w:jc w:val="both"/>');
  pPr = upsertTag(pPr, "w:spacing", '<w:spacing w:line="360" w:lineRule="auto"/>');
  pPr = upsertTag(pPr, "w:ind", '<w:ind w:firstLine="709"/>');

  const runProperties = [
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>',
    '<w:sz w:val="28"/>',
    '<w:szCs w:val="28"/>'
  ].join("");

  pPr = upsertTag(pPr, "w:rPr", `<w:rPr>${runProperties}</w:rPr>`);

  return pPr;
}

function removeTag(xml, tagName) {
  const pattern = new RegExp(`<${tagName}\\b[^>]*/>|<${tagName}\\b[\\s\\S]*?<\\/${tagName}>`, "g");
  return xml.replace(pattern, "");
}

function upsertTag(xml, tagName, tagValue) {
  const pattern = new RegExp(`<${tagName}\\b[^>]*/>|<${tagName}\\b[\\s\\S]*?<\\/${tagName}>`);

  if (pattern.test(xml)) {
    return xml.replace(pattern, tagValue);
  }

  return xml.replace("</w:pPr>", `${tagValue}</w:pPr>`);
}

function setSectionMargins(documentXml) {
  return documentXml.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/, (sectPr) => {
    const marginTag = '<w:pgMar w:top="1134" w:right="567" w:bottom="1134" w:left="1701"/>';

    if (/<w:pgMar\b[^>]*\/>/.test(sectPr)) {
      return sectPr.replace(/<w:pgMar\b[^>]*\/>/, marginTag);
    }

    return sectPr.replace("</w:sectPr>", `${marginTag}</w:sectPr>`);
  });
}

function insertTableOfContents(documentXml) {
  const tocParagraph = [
    PAGE_BREAK_PARAGRAPH,
    "<w:p><w:r><w:t>Зміст</w:t></w:r></w:p>",
    "<w:p>",
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
    '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r>',
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
    "<w:r><w:t>Оновіть поле TOC у Word (F9).</w:t></w:r>",
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    "</w:p>"
  ].join("");

  return documentXml.replace(/<w:body>([\s\S]*?<\/w:p>)/, (full, firstParagraph) => `<w:body>${firstParagraph}${tocParagraph}`);
}

main().catch((error) => {
  console.error(`Помилка: ${error.message}`);
  process.exitCode = 1;
});
