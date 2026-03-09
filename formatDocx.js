const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
const JSZip = require('jszip');
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
  removeRandomCitations: false,
  normalizeBracketCitations: true,
  ensureBibliography: true,
  bibliographySort: 'order', // order | alpha
  customSources: [],
  applyPageSetup: true,
  applyTextFormatting: true,
  applyHeadingStyles: true,
  enforceSectionPageBreaks: true,
  addBlankLinesAroundHeadings: true,
  preserveSpecialContent: true,
  preserveTablesAppearance: true,
  preserveTablesStructure: true,
  justifyDocument: true,
  optionModes: {},
};

function sanitizeEditOptions(raw = {}) {
  const merged = { ...DEFAULT_OPTIONS, ...(raw || {}) };

  const modeKeys = [
    'addTOC',
    'addRandomCitations',
    'removeRandomCitations',
    'normalizeBracketCitations',
    'ensureBibliography',
    'applyPageSetup',
    'applyTextFormatting',
    'applyHeadingStyles',
    'enforceSectionPageBreaks',
    'addBlankLinesAroundHeadings',
    'preserveSpecialContent',
    'preserveTablesAppearance',
    'preserveTablesStructure',
    'justifyDocument',
  ];

  const optionModes = merged.optionModes && typeof merged.optionModes === 'object' ? merged.optionModes : {};

  for (const key of modeKeys) {
    const mode = optionModes[key];
    if (mode === 'add') {
      merged[key] = true;
    } else if (mode === 'remove') {
      merged[key] = false;
    } else {
      merged[key] = Boolean(merged[key]);
    }
  }

  merged.bibliographySort = merged.bibliographySort === 'alpha' ? 'alpha' : 'order';

  const rawSources = Array.isArray(merged.customSources) ? merged.customSources : [];
  merged.customSources = rawSources
    .map((item) => String(item || '').trim())
    .filter(Boolean);

  if (merged.removeRandomCitations) {
    merged.addRandomCitations = false;
  }

  return merged;
}


function normalizeWhitespace(text) {
  return text.replace(/[ \t]{2,}/g, ' ').trim();
}

function normalizeCitationBrackets(text) {
  return text.replace(/\[(.*?)\]/g, (_, inner) => {
    const cleaned = inner
      .replace(/\s*([-–])\s*/g, '$1')
      .replace(/\s*,\s*/g, ', ')
      .replace(/\s{2,}/g, ' ')
      .trim();
    return `[${cleaned}]`;
  });
}

function removeBracketNumberCitations(text) {
  return text
    .replace(/\s*\[\d+\]\s*/g, ' ')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
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
    alignment: config.applyTextFormatting ? (config.justifyDocument ? AlignmentType.JUSTIFIED : AlignmentType.LEFT) : undefined,
    spacing: config.applyTextFormatting ? { line: 360, before: 0, after: 120 } : undefined,
    indent: config.applyTextFormatting ? { firstLine: 709 } : undefined,
  });
}

function getSortedSources(mode, customSources = []) {
  if (customSources.length > 0) {
    const normalizedCustom = customSources.map((text, index) => ({ id: index + 1, text }));
    if (mode === 'alpha') {
      return normalizedCustom.sort((a, b) => a.text.localeCompare(b.text, 'uk'));
    }
    return normalizedCustom;
  }

  if (mode === 'alpha') {
    return [...SOURCES].sort((a, b) => a.text.localeCompare(b.text, 'uk'));
  }
  return SOURCES;
}

function buildBibliographySection(config = DEFAULT_OPTIONS) {
  const sortedSources = getSortedSources(config.bibliographySort, config.customSources);
  const paragraphs = [];

  if (config.addBlankLinesAroundHeadings) paragraphs.push(makeEmptyLine());
  paragraphs.push(makeHeadingParagraph('СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ', 'h1', config));
  if (config.addBlankLinesAroundHeadings) paragraphs.push(makeEmptyLine());

  for (const source of sortedSources) {
    paragraphs.push(makeNormalParagraph(`${source.id}. ${source.text}`, config));
  }

  return paragraphs;
}

async function inputHasTables(inputPath) {
  const data = fs.readFileSync(inputPath);
  const zip = await JSZip.loadAsync(data);
  const docXml = await zip.file('word/document.xml')?.async('text');
  if (!docXml) return false;
  return /<w:tbl[\s>]/.test(docXml);
}

async function extractParagraphs(inputPath, config) {
  const { value } = await mammoth.extractRawText({ path: inputPath });
  let citationsRemoved = 0;

  const paragraphs = value
    .split(/\r?\n/)
    .map((line) => normalizeWhitespace(line))
    .map((line) => (config.normalizeBracketCitations ? normalizeCitationBrackets(line) : line))
    .map((line) => {
      if (!config.removeRandomCitations) return line;
      const matches = line.match(/\[\d+\]/g);
      citationsRemoved += matches ? matches.length : 0;
      return removeBracketNumberCitations(line);
    })
    .filter(Boolean);

  return { paragraphs, citationsRemoved };
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

  onProgress('Перевіряю наявність таблиць...');
  const hasTables = await inputHasTables(inputPath);

  if (hasTables && config.preserveTablesStructure) {
    onProgress('Знайдено таблиці: зберігаю структуру та розмітку оригіналу...');
    fs.copyFileSync(inputPath, outputPath);
    return {
      outputPath,
      stats: {
        sourceParagraphs: 0,
        outputParagraphs: 0,
        citationsAdded: 0,
        citationsRemoved: 0,
        bibliographyAdded: false,
        bibliographySourceCount: config.customSources.length > 0 ? config.customSources.length : SOURCES.length,
        optionsUsed: config,
        notes: [
          'Виявлено таблиці: файл збережено без перебудови, щоб коректно зберегти структуру, відступи та розмітку таблиць.',
        ],
        hasTables,
        preservedOriginalForTables: true,
      },
    };
  }

  onProgress('Зчитую текст з DOCX...');
  const { paragraphs: rawParagraphs, citationsRemoved } = await extractParagraphs(inputPath, config);

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

  const notes = [];
  if (config.preserveSpecialContent) {
    notes.push('Корейські символи в тексті зберігаються як є.');
  }
  if (hasTables) {
    notes.push(config.preserveTablesAppearance
      ? 'У вхідному файлі знайдено таблиці: поточний engine може втрачати частину візуального оформлення таблиць. Рекомендується ручна перевірка таблиць у результаті.'
      : 'У вхідному файлі знайдено таблиці: режим збереження зовнішнього вигляду таблиць вимкнено.');
  }

  return {
    outputPath,
    stats: {
      sourceParagraphs: rawParagraphs.length,
      outputParagraphs: docParagraphs.length,
      citationsAdded,
      citationsRemoved,
      bibliographyAdded,
      bibliographySourceCount: config.customSources.length > 0 ? config.customSources.length : SOURCES.length,
      optionsUsed: config,
      notes,
      hasTables,
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
  removeBracketNumberCitations,
  DEFAULT_OPTIONS,
  resolveHeadingKind,
  sanitizeEditOptions,
};
