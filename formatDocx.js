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
  citationsPerPage: 2,
  insertionPointsPerPage: 6,
  stripBetweenStart: '',
  stripBetweenEnd: '',
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

  merged.citationsPerPage = Number.isFinite(Number(merged.citationsPerPage)) ? Math.max(1, Math.min(6, Number(merged.citationsPerPage))) : 2;
  merged.insertionPointsPerPage = Number.isFinite(Number(merged.insertionPointsPerPage)) ? Math.max(1, Math.min(12, Number(merged.insertionPointsPerPage))) : 6;

  const rawSources = Array.isArray(merged.customSources) ? merged.customSources : [];
  merged.customSources = rawSources
    .map((item) => String(item || '').trim())
    .filter(Boolean);

  merged.stripBetweenStart = String(merged.stripBetweenStart || '').trim();
  merged.stripBetweenEnd = String(merged.stripBetweenEnd || '').trim();

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

function removeBetweenMarkers(text, startMarker, endMarker) {
  if (!startMarker || !endMarker) return text;

  const escapedStart = startMarker.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const escapedEnd = endMarker.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const pattern = new RegExp(`${escapedStart}[\\s\\S]*?${escapedEnd}`, 'g');
  return text.replace(pattern, ' ');
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
  const isBibliographyHeading = headingText === 'СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ';

  return new Paragraph({
    children: [
      makeRun(
        headingText,
        config.applyHeadingStyles ? { bold: true, color: '000000' } : {},
        config,
      ),
    ],
    heading: headingLevel,
    alignment: isBibliographyHeading ? AlignmentType.JUSTIFIED : (config.applyHeadingStyles ? AlignmentType.CENTER : (isH1 ? AlignmentType.CENTER : AlignmentType.LEFT)),
    spacing: config.applyTextFormatting ? { line: 360, before: 0, after: 0 } : undefined,
    pageBreakBefore: isH1 && config.enforceSectionPageBreaks,
    indent: !isH1 && config.applyTextFormatting ? { firstLine: 709 } : undefined,
  });
}

function insertCitationInline(text, citation, anchorRatio = null) {
  const ratio = anchorRatio === null ? Math.random() : Math.max(0.05, Math.min(0.95, anchorRatio));

  const sentenceMatches = [...text.matchAll(/[^.]+\./g)];
  if (sentenceMatches.length > 0) {
    const sentencePos = Math.min(sentenceMatches.length - 1, Math.floor(sentenceMatches.length * ratio));
    const sentence = sentenceMatches[sentencePos];
    const sentenceEndIndex = sentence.index + sentence[0].length - 1;
    return `${text.slice(0, sentenceEndIndex)} ${citation}${text.slice(sentenceEndIndex)}`.replace(/[ 	]{2,}/g, ' ');
  }

  const fallbackDotIndex = text.lastIndexOf('.');
  if (fallbackDotIndex >= 0) {
    return `${text.slice(0, fallbackDotIndex)} ${citation}${text.slice(fallbackDotIndex)}`.replace(/[ 	]{2,}/g, ' ');
  }

  return `${text} ${citation}`.replace(/[ 	]{2,}/g, ' ').trim();
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
      spacing: config.applyTextFormatting ? { line: 360, before: 0, after: 0 } : undefined,
      indent: config.applyTextFormatting ? { left: 709, hanging: 360 } : undefined,
    });
  }

  return new Paragraph({
    children: [makeRun(text, {}, config)],
    alignment: config.applyTextFormatting ? (config.justifyDocument ? AlignmentType.JUSTIFIED : AlignmentType.LEFT) : undefined,
    spacing: config.applyTextFormatting ? { line: 360, before: 0, after: 0 } : undefined,
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
    .map((line) => removeBetweenMarkers(line, config.stripBetweenStart, config.stripBetweenEnd))
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
    return { paragraphs, citationsAdded: 0, citationsTarget: 0, insertionPointsTarget: 0 };
  }

  const EXCLUDED_SECTIONS = new Set(['ВСТУП', 'ВИСНОВКИ']);
  const WORDS_PER_PAGE_ESTIMATE = 260;
  const MIN_FULL_WIDTH_WORDS = 10;
  const MIN_FULL_WIDTH_CHARS = 70;

  let currentSection = null;
  let totalEligibleWords = 0;
  const eligiblePoints = [];

  for (let i = 0; i < paragraphs.length; i += 1) {
    const item = paragraphs[i];

    if (item.type === 'h1') {
      currentSection = item.text.toUpperCase();
      continue;
    }

    if (item.type !== 'normal') {
      continue;
    }

    const words = item.text.split(/\s+/).filter(Boolean);
    const isList = Boolean(parseListItem(item.text));
    const isExcludedSection = currentSection && EXCLUDED_SECTIONS.has(currentSection);
    const notFullWidth = words.length < MIN_FULL_WIDTH_WORDS || item.text.length < MIN_FULL_WIDTH_CHARS;

    if (isList || isExcludedSection || notFullWidth) {
      continue;
    }

    const pointsInParagraph = Math.max(1, Math.min(3, Math.floor(words.length / 40) + 1));
    for (let point = 0; point < pointsInParagraph; point += 1) {
      eligiblePoints.push({
        index: i,
        ratio: (point + 1) / (pointsInParagraph + 1),
      });
    }

    totalEligibleWords += words.length;
  }

  if (eligiblePoints.length === 0) {
    return { paragraphs, citationsAdded: 0, citationsTarget: 0, insertionPointsTarget: 0 };
  }

  const estimatedPages = Math.max(1, Math.round(totalEligibleWords / WORDS_PER_PAGE_ESTIMATE));
  const insertionPointsTarget = Math.min(eligiblePoints.length, estimatedPages * config.insertionPointsPerPage);
  const citationsTarget = Math.min(insertionPointsTarget, estimatedPages * config.citationsPerPage);

  const selectedPoints = [];
  for (let k = 0; k < insertionPointsTarget; k += 1) {
    const pos = Math.floor(((k + 0.5) * eligiblePoints.length) / insertionPointsTarget);
    const boundedPos = Math.max(0, Math.min(eligiblePoints.length - 1, pos));
    selectedPoints.push(eligiblePoints[boundedPos]);
  }

  const citationPoints = [];
  if (citationsTarget > 0) {
    for (let k = 0; k < citationsTarget; k += 1) {
      const pos = Math.floor(((k + 0.5) * selectedPoints.length) / citationsTarget);
      const boundedPos = Math.max(0, Math.min(selectedPoints.length - 1, pos));
      citationPoints.push(selectedPoints[boundedPos]);
    }
  }

  const byParagraph = new Map();
  for (const point of citationPoints) {
    const arr = byParagraph.get(point.index) || [];
    arr.push(point.ratio);
    byParagraph.set(point.index, arr);
  }

  const output = [];
  let lastSourceId = null;
  let citationsAdded = 0;

  for (let i = 0; i < paragraphs.length; i += 1) {
    const item = paragraphs[i];
    const ratios = byParagraph.get(i);

    if (!ratios || ratios.length === 0) {
      output.push(item);
      continue;
    }

    let text = item.text;
    ratios.sort((a, b) => a - b);
    for (const ratio of ratios) {
      const { citation, sourceId } = createCitation(lastSourceId);
      text = insertCitationInline(text, citation, ratio);
      lastSourceId = sourceId;
      citationsAdded += 1;
    }

    output.push({ ...item, text });
  }

  return { paragraphs: output, citationsAdded, citationsTarget, insertionPointsTarget };
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
  const { paragraphs: withCitations, citationsAdded, citationsTarget, insertionPointsTarget } = addCitations(structured, config);

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
      insertionPointsTarget,
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
  removeBetweenMarkers,
  insertCitationInline,
  parseListItem,
  DEFAULT_OPTIONS,
  resolveHeadingKind,
  sanitizeEditOptions,
};
