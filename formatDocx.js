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

function insertCitationInline(text, citation) {
  const words = text.split(/\s+/).filter(Boolean);
  if (words.length < 3) return `${text} ${citation}`.trim();
  const idx = Math.max(1, Math.min(words.length - 1, Math.floor(Math.random() * (words.length - 1))));
  words.splice(idx, 0, citation);
  return words.join(' ');
}

function parseListItem(text) {
  const m = text.match(/^\s*((?:[-•*])|(?:\d+[\.)]))\s+(.*)$/);
  if (!m) return null;
  return { marker: m[1], content: m[2] };
}

function makeNormalParagraph(text, config = DEFAULT_OPTIONS) {
  const listItem = parseListItem(text);
  if (listItem) {
    return new Paragraph({
      children: [makeRun(`${listItem.marker}	${listItem.content}`, {}, config)],
      alignment: config.applyTextFormatting ? AlignmentType.LEFT : undefined,
      spacing: config.applyTextFormatting ? { line: 360, before: 0, after: 120 } : undefined,
      indent: config.applyTextFormatting ? { left: 709, hanging: 360 } : undefined,
    });
  }

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
    return { paragraphs, citationsAdded: 0, citationsTarget: 0 };
  }

  const EXCLUDED_SECTIONS = new Set(['ВСТУП', 'ВИСНОВКИ']);
  const WORDS_PER_PAGE_ESTIMATE = 260;

  let currentSection = null;
  let totalEligibleWords = 0;
  const eligibleIndexes = [];

  for (let i = 0; i < paragraphs.length; i += 1) {
    const item = paragraphs[i];

    if (item.type === 'h1') {
      currentSection = item.text.toUpperCase();
      continue;
    }

    if (item.type !== 'normal') {
      continue;
    }

    const isList = Boolean(parseListItem(item.text));
    const isExcludedSection = currentSection && EXCLUDED_SECTIONS.has(currentSection);
    if (isList || isExcludedSection) {
      continue;
    }

    eligibleIndexes.push(i);
    totalEligibleWords += item.text.split(/\s+/).filter(Boolean).length;
  }

  if (eligibleIndexes.length === 0) {
    return { paragraphs, citationsAdded: 0, citationsTarget: 0 };
  }

  const estimatedPages = Math.max(1, Math.round(totalEligibleWords / WORDS_PER_PAGE_ESTIMATE));
  let citationsTarget = 0;
  for (let i = 0; i < estimatedPages; i += 1) {
    citationsTarget += Math.random() < 0.5 ? 2 : 3;
  }
  citationsTarget = Math.min(citationsTarget, eligibleIndexes.length);

  const shuffled = [...eligibleIndexes].sort(() => Math.random() - 0.5);
  const chosen = new Set(shuffled.slice(0, citationsTarget));

  const output = [];
  let lastSourceId = null;
  let citationsAdded = 0;

  for (let i = 0; i < paragraphs.length; i += 1) {
    const item = paragraphs[i];
    if (chosen.has(i)) {
      const { citation, sourceId } = createCitation(lastSourceId);
      output.push({ ...item, text: insertCitationInline(item.text, citation) });
      lastSourceId = sourceId;
      citationsAdded += 1;
    } else {
      output.push(item);
    }
  }

  return { paragraphs: output, citationsAdded, citationsTarget };
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
  onProgress('Зчитую текст з DOCX...');
  const { paragraphs: rawParagraphs, citationsRemoved } = await extractParagraphs(inputPath, config);

  onProgress('Класифікую заголовки та абзаци...');
  const structured = rawParagraphs.map((text) => ({ type: classifyParagraph(text), text }));

  onProgress('Обробляю посилання та структуру...');
  const { paragraphs: withCitations, citationsAdded, citationsTarget } = addCitations(structured, config);

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
    notes.push('Усі символи у тексті зберігаються як є (включно з корейськими та спеціальними символами).');
  }
  if (hasTables) {
    notes.push('У вхідному файлі знайдено таблиці: структура та розмітка таблиць збережені настільки, наскільки це підтримує DOCX engine. Перевірте таблиці вручну у Word.');
  }

  return {
    outputPath,
    stats: {
      sourceParagraphs: rawParagraphs.length,
      outputParagraphs: docParagraphs.length,
      citationsAdded,
      citationsTarget,
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
  insertCitationInline,
  parseListItem,
  DEFAULT_OPTIONS,
  resolveHeadingKind,
  sanitizeEditOptions,
};
