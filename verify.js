const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

async function readDocumentXml(docxPath) {
  const data = fs.readFileSync(docxPath);
  const zip = await JSZip.loadAsync(data);
  const docXml = await zip.file('word/document.xml')?.async('text');
  if (!docXml) throw new Error('Не вдалося прочитати word/document.xml');
  return docXml;
}

function check(regex, content) {
  return regex.test(content);
}

async function verifyDocx(docxPath) {
  if (!fs.existsSync(docxPath)) {
    throw new Error(`Файл не знайдено: ${docxPath}`);
  }

  const xml = await readDocumentXml(docxPath);

  const results = [
    {
      name: 'Heading 1 присутні',
      pass: check(/w:pStyle w:val="Heading1"/, xml),
    },
    {
      name: 'Heading 2 присутні',
      pass: check(/w:pStyle w:val="Heading2"/, xml),
    },
    {
      name: 'TOC присутній',
      pass:
        check(/TOC\s+\\h[\s\S]*?1-2/, xml) ||
        check(/TOC\s+\\o\s+&quot;1-2&quot;/, xml) ||
        check(/TOC\s+\\o\s+"1-2"/, xml),
    },
    {
      name: 'Page break перед Heading 1',
      pass: check(/w:pageBreakBefore\/>/, xml),
    },
    {
      name: 'Times New Roman 14',
      pass: check(/w:rFonts w:ascii="Times New Roman"/, xml) && check(/w:sz w:val="28"/, xml),
    },
    {
      name: 'Немає подвійних пробілів',
      pass: !check(/ {2,}/, xml),
    },
  ];

  const failed = results.filter((r) => !r.pass);
  const ok = failed.length === 0;

  return {
    ok,
    failedCount: failed.length,
    results,
  };
}

async function main() {
  const input = process.argv[2] || 'output.docx';
  try {
    const report = await verifyDocx(path.resolve(input));

    console.log('\nРезультати перевірки:');
    for (const result of report.results) {
      console.log(`- ${result.pass ? 'PASS' : 'FAIL'}: ${result.name}`);
    }

    if (report.ok) {
      console.log('\nПІДСУМОК: PASS ✅');
      process.exit(0);
      return;
    }

    console.log(`\nПІДСУМОК: FAIL ❌ (${report.failedCount} невідповідностей)`);
    process.exit(2);
  } catch (error) {
    console.error('Помилка перевірки:', error.message);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

module.exports = {
  verifyDocx,
};
