const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
const {
  AlignmentType,
  Document,
  HeadingLevel,
  PageOrientation,
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
  'СПИСОК ДЖЕРЕЛ',
]);

const DEFAULT_OPTIONS = {
  addTOC: true,
  addRandomCitations: true,
  normalizeBracketCitations: true,
  ensureBibliography: true,
  bibliographySort: 'order', // order | alpha
  applyPageSetup: true,
  applyTextFormatting: true,
  applyHeadingStyles: true,
  enforceSectionPageBreaks: true,
  addBlankLinesAroundHeadings: true,
  preserveSpecialContent: true,
};

function sanitizeEditOptions(raw = {}) {
  const merged = { ...DEFAULT_OPTIONS, ...(raw || {}) };
  const boolKeys = [
    'addTOC',
    'addRandomCitations',
    'normalizeBracketCitations',
    'ensureBibliography',
    'applyPageSetup',
    'applyTextFormatting',
    'applyHeadingStyles',
    'enforceSectionPageBreaks',
    'addBlankLinesAroundHeadings',
    'preserveSpecialContent',
  ];

  for (const key of boolKeys) {
    merged[key] = Boolean(merged[key]);
  }

  merged.bibliographySort = merged.bibliographySort === 'alpha' ? 'alpha' : 'order';
  return merged;
}

function normalizeWhitespace(text) {
  return text.replace(/[ \t]{2,}/g, ' ').trim();
}

function normalizeCitationBrackets(text) {
  // [ 1 ] -> [1], [5 - 7] -> [5-7], [3,  с. 12] -> [3, с. 12]
  return text.replace(/\[(.*?)\]/g, (_, inner) => {
    const cleaned = inner
      .replace(/\s*([-–])\s*/g, '$1')
      .replace(/\s*,\s*/g, ', ')
      .replace(/\s{2,}/g, ' ')
      .trim();
    return `[${cleaned}]`;
  });
}

function classifyParagraph(text) {
  const upper = text.toUpperCase();
  if (HEADING1_EXACT.has(upper) || HEADING1_REGEX.test(upper)) return 'h1';
  if (HEADING2_REGEX.test(text)) return 'h2';
  return 'normal';
}

function makeRun(text, opts = {}, config = DEFAULT_OPTIONS) {
  const base = config.applyTextFormatting
    ? { font: 'Times New Roman', size: 28 }
    : {};

  return new TextRun({
    text,
    ...base,
    ...opts,
  });
}

function makeEmptyLine() {
  return new Paragraph({ children: [new TextRun({ text: '' })] });
}

function resolveHeadingKind(levelOrType) {
  if (levelOrType === HeadingLevel.HEADING_1) return 'h1';
  if (levelOrType === HeadingLevel.HEADING_2) return 'h2';
  if (levelOrType === 'h1' || levelOrType === 'h2') return levelOrType;
  return 'h2';
}

function makeHeadingParagraph(text, typeOrLevel, config = DEFAULT_OPTIONS) {
  const headingKind = resolveHeadingKind(typeOrLevel);
  const isH1 = headingKind === 'h1';
  const headingLevel = isH1 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2;
  const headingText = isH1 ? text.toUpperCase() : text.charAt(0).toUpperCase() + text.slice(1);

  return new Paragraph({
    children: [
      makeRun(
        headingText,
        config.applyHeadingStyles ? { bold: true, color: '000000' } : {},
        config,
      ),
    ],
    heading: headingLevel,
    alignment: config.applyHeadingStyles ? AlignmentType.CENTER : (isH1 ? AlignmentType.CENTER : AlignmentType.LEFT),
    spacing: config.applyTextFormatting ? { line: 360, before: 0, after: 0 } : undefined,
    pageBreakBefore: isH1 && config.enforceSectionPageBreaks,
    indent: !isH1 && config.applyTextFormatting ? { firstLine: 709 } : undefined,
  });
}

function makeNormalParagraph(text, config = DEFAULT_OPTIONS) {
  return new Paragraph({
    children: [makeRun(text, {}, config)],
    alignment: config.applyTextFormatting ? AlignmentType.JUSTIFIED : undefined,
    spacing: config.applyTextFormatting ? { line: 360, before: 0, after: 120 } : undefined,
    indent: config.applyTextFormatting ? { firstLine: 709 } : undefined,
  });
}

function getSortedSources(mode) {
  if (mode === 'alpha') {
    return [...SOURCES].sort((a, b) => a.text.localeCompare(b.text, 'uk'));
  }
  return SOURCES;
}

function buildBibliographySection(config = DEFAULT_OPTIONS) {
  const sortedSources = getSortedSources(config.bibliographySort);
  const paragraphs = [];

  if (config.addBlankLinesAroundHeadings) paragraphs.push(makeEmptyLine());
  paragraphs.push(makeHeadingParagraph('СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ', 'h1', config));
  if (config.addBlankLinesAroundHeadings) paragraphs.push(makeEmptyLine());

  for (const source of sortedSources) {
    paragraphs.push(
      new Paragraph({
        children: [makeRun(`${source.id}. ${source.text}`, {}, config)],
        alignment: config.applyTextFormatting ? AlignmentType.JUSTIFIED : undefined,
        spacing: config.applyTextFormatting ? { line: 360, before: 0, after: 120 } : undefined,
        indent: config.applyTextFormatting ? { firstLine: 709 } : undefined,
      }),
    );
  }

  return paragraphs;
}

async function extractParagraphs(inputPath, config) {
  const { value } = await mammoth.extractRawText({ path: inputPath });
  return value
    .split(/\r?\n/)
    .map((line) => normalizeWhitespace(line))
    .map((line) => (config.normalizeBracketCitations ? normalizeCitationBrackets(line) : line))
    .filter(Boolean);
}

function addCitations(paragraphs, config) {
  if (!config.addRandomCitations) {
    return { paragraphs, citationsAdded: 0 };
  }

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

  return { paragraphs: output, citationsAdded };
}

async function formatDocx(inputPath, outputPath, options = {}) {
  const onProgress = options.onProgress || (() => {});
  const config = sanitizeEditOptions(options.editOptions || {});

  onProgress('Перевіряю вхідний файл...');
  if (!fs.existsSync(inputPath)) {
    throw new Error(`Вхідний файл не знайдено: ${inputPath}`);
  }

  onProgress('Зчитую текст з DOCX...');
  const rawParagraphs = await extractParagraphs(inputPath, config);

  onProgress('Класифікую заголовки та абзаци...');
  const structured = rawParagraphs.map((text) => ({ type: classifyParagraph(text), text }));

  onProgress('Обробляю посилання та структуру...');
  const { paragraphs: withCitations, citationsAdded } = addCitations(structured, config);

  const docParagraphs = [];

  if (config.addTOC) {
    onProgress('Формую зміст документа...');
    if (config.addBlankLinesAroundHeadings) docParagraphs.push(makeEmptyLine());
    docParagraphs.push(makeHeadingParagraph('ЗМІСТ', 'h1', config));
    if (config.addBlankLinesAroundHeadings) docParagraphs.push(makeEmptyLine());
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
  }

  for (const item of withCitations) {
    if (item.type === 'h1') {
      if (config.addBlankLinesAroundHeadings) docParagraphs.push(makeEmptyLine());
      docParagraphs.push(makeHeadingParagraph(item.text, 'h1', config));
      if (config.addBlankLinesAroundHeadings) docParagraphs.push(makeEmptyLine());
    } else if (item.type === 'h2') {
      if (config.addBlankLinesAroundHeadings) docParagraphs.push(makeEmptyLine());
      docParagraphs.push(makeHeadingParagraph(item.text, 'h2', config));
      if (config.addBlankLinesAroundHeadings) docParagraphs.push(makeEmptyLine());
    } else if (item.type === 'citation') {
      docParagraphs.push(
        new Paragraph({
          children: [makeRun(item.text, { italics: true }, config)],
          alignment: config.applyTextFormatting ? AlignmentType.RIGHT : undefined,
          spacing: config.applyTextFormatting ? { line: 360, before: 0, after: 120 } : undefined,
        }),
      );
    } else {
      docParagraphs.push(makeNormalParagraph(item.text, config));
    }
  }

  let bibliographyAdded = false;
  if (config.ensureBibliography && !withCitations.some((p) => p.type === 'h1' && ['СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ', 'СПИСОК ДЖЕРЕЛ'].includes(p.text.toUpperCase()))) {
    docParagraphs.push(...buildBibliographySection(config));
    bibliographyAdded = true;
  }

  const sectionProperties = config.applyPageSetup
    ? {
        page: {
          margin: {
            top: 1134,
            right: 567,
            bottom: 1134,
            left: 1701,
          },
          size: {
            orientation: PageOrientation.PORTRAIT,
          },
        },
      }
    : {};

  const doc = new Document({
    sections: [
      {
        properties: sectionProperties,
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
      optionsUsed: config,
      notes: config.preserveSpecialContent
        ? ['Корейські символи в тексті зберігаються як є.', 'Складні обʼєкти (таблиці/формули) залежать від якості витягування Mammoth.']
        : [],
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
  normalizeCitationBrackets,
  DEFAULT_OPTIONS,
  resolveHeadingKind,
  sanitizeEditOptions,
};
