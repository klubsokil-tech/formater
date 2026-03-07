import fs from "node:fs/promises";
import JSZip from "jszip";

const DOC_XML_PATH = "word/document.xml";

async function main() {
  const [targetPath] = process.argv.slice(2);

  if (!targetPath) {
    throw new Error("Використання: node verify.js output.docx");
  }

  const buffer = await fs.readFile(targetPath);
  const zip = await JSZip.loadAsync(buffer);
  const documentFile = zip.file(DOC_XML_PATH);

  if (!documentFile) {
    throw new Error(`Файл ${DOC_XML_PATH} не знайдено у ${targetPath}`);
  }

  const documentXml = await documentFile.async("string");

  const checks = [
    { name: "TOC", ok: /TOC \\o/.test(documentXml) },
    { name: "Бібліографія", ok: /Список використаних джерел/.test(documentXml) },
    { name: "Цитування", ok: /Цитування:/.test(documentXml) }
  ];

  const failed = checks.filter((check) => !check.ok);

  for (const check of checks) {
    console.log(`${check.ok ? "✅" : "❌"} ${check.name}`);
  }

  if (failed.length > 0) {
    process.exitCode = 1;
    return;
  }

  console.log("Документ пройшов базову перевірку.");
}

main().catch((error) => {
  console.error(`Помилка: ${error.message}`);
  process.exitCode = 1;
});
