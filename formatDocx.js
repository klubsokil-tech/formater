import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { attachCitationReferences, insertBibliography } from "./bibliography.js";

const DOC_XML_PATH = "word/document.xml";

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
  documentXml = attachCitationReferences(documentXml, [1, 2]);
  documentXml = insertBibliography(documentXml, [
    "DSTU 8302:2015. Бібліографічне посилання. Загальні положення.",
    "ISO 690:2021. Information and documentation — Guidelines for bibliographic references."
  ]);

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
  const normalizedText = documentXml.replace(/(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/g, (full, open, text, close) => {
    const compact = text.replace(/\s{2,}/g, " ");
    return `${open}${compact}${close}`;
  });

  if (analysis.headingsCount === 0) {
    const warningParagraph = "<w:p><w:r><w:t>Попередження: у документі не знайдено заголовків Heading 1-6.</w:t></w:r></w:p>";
    return normalizedText.replace("</w:body>", `${warningParagraph}</w:body>`);
  }

  return normalizedText;
}

function insertTableOfContents(documentXml) {
  const tocParagraph = [
    "<w:p>",
    "<w:r><w:t>Зміст</w:t></w:r>",
    "</w:p>",
    "<w:p>",
    "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>",
    "<w:r><w:instrText xml:space=\"preserve\"> TOC \\o \"1-3\" \\h \\z \\u </w:instrText></w:r>",
    "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>",
    "<w:r><w:t>Оновіть поле TOC у Word (F9).</w:t></w:r>",
    "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>",
    "</w:p>"
  ].join("");

  return documentXml.replace(/<w:body>/, `<w:body>${tocParagraph}`);
}

main().catch((error) => {
  console.error(`Помилка: ${error.message}`);
  process.exitCode = 1;
});
