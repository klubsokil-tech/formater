import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

const DOC_XML_PATH = "word/document.xml";

async function main() {
  const [targetPath] = process.argv.slice(2);

  if (!targetPath) {
    throw new Error("Використання: node verify.js output.docx");
  }

  if (path.extname(targetPath).toLowerCase() !== ".docx") {
    throw new Error(`Очікується .docx файл: ${targetPath}`);
  }

  const buffer = await fs.readFile(targetPath);
  const zip = await JSZip.loadAsync(buffer);
  const documentFile = zip.file(DOC_XML_PATH);

  if (!documentFile) {
    throw new Error(`Файл ${DOC_XML_PATH} не знайдено у ${targetPath}`);
  }

  const documentXml = await documentFile.async("string");
  const paragraphs = extractParagraphs(documentXml);

  const checks = [
    { name: "Наявний Heading 1", ok: hasHeading(paragraphs, "Heading1") },
    { name: "Наявний Heading 2", ok: hasHeading(paragraphs, "Heading2") },
    { name: "Наявний TOC", ok: /<w:instrText[^>]*>\s*TOC\\o\s+"1-2"/.test(documentXml) },
    { name: "Є page break перед першим Heading 1", ok: hasPageBreakBeforeFirstHeading1(paragraphs) },
    { name: "Застосовано Times New Roman 14", ok: hasTimesNewRoman14(paragraphs) },
    { name: "Відсутні подвійні пробіли", ok: hasNoDoubleSpaces(documentXml) }
  ];

  const failed = checks.filter((check) => !check.ok);

  console.log("Перевірка документа:");
  for (const check of checks) {
    console.log(`- ${check.ok ? "PASS" : "FAIL"}: ${check.name}`);
  }

  const summary = failed.length === 0
    ? "PASS: Усі правила пройдено."
    : `FAIL: Провалено правил: ${failed.length}/${checks.length}.`;

  console.log(`Підсумок: ${summary}`);

  if (failed.length > 0) {
    process.exitCode = 1;
  }
}

function extractParagraphs(documentXml) {
  return documentXml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
}

function hasHeading(paragraphs, headingName) {
  const pattern = new RegExp(`<w:pStyle\\b[^>]*w:val="${headingName}"`);
  return paragraphs.some((paragraph) => pattern.test(paragraph));
}

function hasPageBreakBeforeFirstHeading1(paragraphs) {
  const firstHeadingIndex = paragraphs.findIndex((paragraph) => /<w:pStyle\b[^>]*w:val="Heading1"/.test(paragraph));

  if (firstHeadingIndex <= 0) {
    return false;
  }

  return /<w:br\b[^>]*w:type="page"/.test(paragraphs[firstHeadingIndex - 1]);
}

function hasTimesNewRoman14(paragraphs) {
  const contentParagraphs = paragraphs.filter((paragraph) => {
    const hasText = /<w:t\b[^>]*>/.test(paragraph);
    const isFieldCode = /<w:instrText\b|<w:fldChar\b/.test(paragraph);
    return hasText && !isFieldCode;
  });

  if (contentParagraphs.length === 0) {
    return false;
  }

  return contentParagraphs.some((paragraph) => {
    const hasTimes = /<w:rFonts\b[^>]*(w:ascii="Times New Roman"|w:hAnsi="Times New Roman")/.test(paragraph);
    const hasSize14 = /<w:sz\b[^>]*w:val="28"/.test(paragraph) && /<w:szCs\b[^>]*w:val="28"/.test(paragraph);
    return hasTimes && hasSize14;
  });
}

function hasNoDoubleSpaces(documentXml) {
  const textNodes = [...documentXml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)];
  return textNodes.every((match) => !/ {2,}/.test(decodeXmlText(match[1])));
}

function decodeXmlText(text) {
  return text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#x20;/g, " ");
}

main().catch((error) => {
  console.error(`Помилка: ${error.message}`);
  process.exitCode = 1;
});
