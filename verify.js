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
      pass: check(/TOC \\o "1-2" \\h \\z \\u/, xml),
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

  console.log('\nРезультати перевірки:');
  for (const result of results) {
    console.log(`- ${result.pass ? 'PASS' : 'FAIL'}: ${result.name}`);
  }

  if (failed.length === 0) {
    console.log('\nПІДСУМОК: PASS ✅');
    return true;
  }

  console.log(`\nПІДСУМОК: FAIL ❌ (${failed.length} невідповідностей)`);
  return false;
}

async function main() {
  const input = process.argv[2] || 'output.docx';
  try {
    const ok = await verifyDocx(path.resolve(input));
    process.exit(ok ? 0 : 2);
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
