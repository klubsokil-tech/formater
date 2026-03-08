const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
const {
  AlignmentType,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  TableOfContents,
  TextRun,
} = require('docx');
const { SOURCES, createCitation } = require('./bibliography');

const HEADING1_REGEX = /^РОЗДІЛ\s+\d+/i;
const HEADING2_REGEX = /^\d+\.\d+/;
const HEADING1_EXACT = new Set([
  'ВСТУП',
  'ВИСНОВКИ',
  'СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ',
]);

function normalizeWhitespace(text) {
  return text.replace(/[ \t]{2,}/g, ' ').trim();
}

function classifyParagraph(text) {
  const upper = text.toUpperCase();
  if (HEADING1_EXACT.has(upper) || HEADING1_REGEX.test(upper)) return 'h1';
  if (HEADING2_REGEX.test(text)) return 'h2';
  return 'normal';
}

function makeRun(text, opts = {}) {
  return new TextRun({
    text,
    font: 'Times New Roman',
    size: 28,
    ...opts,
  });
}

function makeHeadingParagraph(text, level) {
  return new Paragraph({
    children: [makeRun(text)],
    heading: level,
    alignment: AlignmentType.CENTER,
    spacing: { line: 360, before: 200, after: 120 },
    pageBreakBefore: level === HeadingLevel.HEADING_1,
  });
}

function makeNormalParagraph(text) {
  return new Paragraph({
    children: [makeRun(text)],
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 360, before: 0, after: 120 },
    indent: { firstLine: 709 },
  });
}

function buildBibliographySection() {
  const paragraphs = [
    makeHeadingParagraph('СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ', HeadingLevel.HEADING_1),
  ];

  for (const source of SOURCES) {
    paragraphs.push(
      new Paragraph({
        children: [makeRun(`${source.id}. ${source.text}`)],
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 360, before: 0, after: 120 },
        indent: { firstLine: 709 },
      }),
    );
  }

  return paragraphs;
}

async function extractParagraphs(inputPath) {
  const { value } = await mammoth.extractRawText({ path: inputPath });
  return value
    .split(/\r?\n/)
    .map((line) => normalizeWhitespace(line))
    .filter(Boolean);
}

function addCitations(paragraphs) {
  const output = [];
  let normalCounter = 0;
  let lastSourceId = null;
  let lastWasCitation = false;
  let target = Math.random() < 0.5 ? 1 : 2;
  let citationsAdded = 0;

  for (const item of paragraphs) {
    output.push(item);

    if (item.type !== 'normal') {
      normalCounter = 0;
      continue;
    }

    normalCounter += 1;

    if (normalCounter >= target && !lastWasCitation) {
      const { citation, sourceId } = createCitation(lastSourceId);
      output.push({ type: 'citation', text: citation });
      lastSourceId = sourceId;
      lastWasCitation = true;
      normalCounter = 0;
      target = Math.random() < 0.5 ? 1 : 2;
      citationsAdded += 1;
    } else {
      lastWasCitation = false;
    }
  }

  return {
    paragraphs: output,
    citationsAdded,
  };
}

async function formatDocx(inputPath, outputPath, options = {}) {
  const onProgress = options.onProgress || (() => {});

  onProgress('Перевіряю вхідний файл...');
  if (!fs.existsSync(inputPath)) {
    throw new Error(`Вхідний файл не знайдено: ${inputPath}`);
  }

  onProgress('Зчитую текст з DOCX...');
  const rawParagraphs = await extractParagraphs(inputPath);

  onProgress('Класифікую заголовки та абзаци...');
  const structured = rawParagraphs.map((text) => ({
    type: classifyParagraph(text),
    text,
  }));

  onProgress('Додаю випадкові посилання...');
  const { paragraphs: withCitations, citationsAdded } = addCitations(structured);

  const docParagraphs = [];

  onProgress('Формую зміст та структуру документа...');
  docParagraphs.push(
    makeHeadingParagraph('ЗМІСТ', HeadingLevel.HEADING_1),
  );
  docParagraphs.push(
    new TableOfContents('Зміст', {
      hyperlink: true,
      headingStyleRange: '1-2',
      stylesWithLevels: [
        { styleName: 'Heading 1', level: 1 },
        { styleName: 'Heading 2', level: 2 },
      ],
    }),
  );

  for (const item of withCitations) {
    if (item.type === 'h1') {
      docParagraphs.push(makeHeadingParagraph(item.text.toUpperCase(), HeadingLevel.HEADING_1));
    } else if (item.type === 'h2') {
      docParagraphs.push(makeHeadingParagraph(item.text, HeadingLevel.HEADING_2));
    } else if (item.type === 'citation') {
      docParagraphs.push(
        new Paragraph({
          children: [makeRun(item.text, { italics: true })],
          alignment: AlignmentType.RIGHT,
          spacing: { line: 360, before: 0, after: 120 },
        }),
      );
    } else {
      docParagraphs.push(makeNormalParagraph(item.text));
    }
  }

  let bibliographyAdded = false;
  if (!withCitations.some((p) => p.type === 'h1' && p.text.toUpperCase() === 'СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ')) {
    docParagraphs.push(...buildBibliographySection());
    bibliographyAdded = true;
  }

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1134,
              right: 567,
              bottom: 1134,
              left: 1701,
            },
          },
        },
        children: docParagraphs,
      },
    ],
  });

  onProgress('Зберігаю відформатований файл...');
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);

  return {
    outputPath,
    stats: {
      sourceParagraphs: rawParagraphs.length,
      outputParagraphs: docParagraphs.length,
      citationsAdded,
      bibliographyAdded,
    },
  };
}

async function main() {
  const inputPath = process.argv[2] || 'input.docx';
  const outputPath = process.argv[3] || 'output.docx';

  try {
    const result = await formatDocx(path.resolve(inputPath), path.resolve(outputPath));
    console.log(`Готово. Створено файл: ${result.outputPath}`);
    console.log('За бажанням виконайте перевірку: node verify.js ' + outputPath);
  } catch (error) {
    console.error('Помилка форматування:', error.message);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

module.exports = {
  formatDocx,
  classifyParagraph,
  normalizeWhitespace,
};
