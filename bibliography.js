const BIBLIOGRAPHY_TITLE = "Список використаних джерел";

export const bibliographySources = [
  { title: "ДСТУ 3008:2015. Інформація та документація. Звіти у сфері науки і техніки.", pages: 30 },
  { title: "ДСТУ 8302:2015. Бібліографічне посилання. Загальні положення та правила складання.", pages: 20 },
  { title: "ISO 690:2021. Information and documentation — Guidelines for bibliographic references.", pages: 35 },
  { title: "Ковальчук О.І. Методологія наукових досліджень. — Київ: Центр учбової літератури, 2020.", pages: 280 },
  { title: "Пилипчук М.І., Григор’єв А.С., Шостак В.В. Основи наукових досліджень. — Київ: Знання, 2019.", pages: 270 },
  { title: "Гончаренко С.У. Педагогічні дослідження: методологічні поради молодим науковцям. — Київ, 2018.", pages: 180 },
  { title: "Білуха М.Т. Методологія наукових досліджень. — Київ: АБУ, 2017.", pages: 240 },
  { title: "Шейко В.М., Кушнаренко Н.М. Організація та методика науково-дослідницької діяльності. — Київ: Знання, 2016.", pages: 310 },
  { title: "Цехмістрова Г.С. Основи наукових досліджень. — Київ: Слово, 2015.", pages: 320 },
  { title: "Крушельницька О.В. Методологія та організація наукових досліджень. — Київ: Кондор, 2014.", pages: 206 },
  { title: "Мочерний С.В. Методологія економічного дослідження. — Львів: Світ, 2013.", pages: 415 },
  { title: "Єріна А.М., Захожай В.Б., Єрін Д.Л. Методологія наукових досліджень. — Київ: Центр навчальної літератури, 2012.", pages: 212 },
  { title: "Петрушенко В.Л. Філософія і методологія науки. — Львів: Новий Світ-2000, 2011.", pages: 190 },
  { title: "Бушуєв С.Д. Управління проєктами: основи професійних знань. — Київ: ІРІДІУМ, 2010.", pages: 640 },
  { title: "Zobel J. Writing for Computer Science. 3rd ed. — London: Springer, 2014.", pages: 284 }
];

export function buildBibliographyParagraphs(sources = bibliographySources) {
  const normalizedSources = normalizeSources(sources);

  if (normalizedSources.length === 0) {
    return [
      `<w:p><w:r><w:t>${BIBLIOGRAPHY_TITLE}</w:t></w:r></w:p>`,
      "<w:p><w:r><w:t>[1] Додайте бібліографічні джерела у bibliography.js</w:t></w:r></w:p>"
    ];
  }

  const items = normalizedSources.map(
    (source, index) => `<w:p><w:r><w:t>[${index + 1}] ${escapeXml(source.title)}</w:t></w:r></w:p>`
  );

  return [`<w:p><w:r><w:t>${BIBLIOGRAPHY_TITLE}</w:t></w:r></w:p>`, ...items];
}

export function insertBibliography(documentXml, sources = bibliographySources) {
  const paragraphs = buildBibliographyParagraphs(sources).join("");
  return documentXml.replace("</w:body>", `${paragraphs}</w:body>`);
}

export function insertCitationReferences(documentXml, sources = bibliographySources) {
  const normalizedSources = normalizeSources(sources);

  if (normalizedSources.length === 0) {
    return documentXml;
  }

  return documentXml.replace(/<w:body>([\s\S]*?)<\/w:body>/, (full, bodyContent) => {
    const paragraphPattern = /<w:p\b[\s\S]*?<\/w:p>/g;
    const paragraphs = [...bodyContent.matchAll(paragraphPattern)].map((match) => match[0]);

    if (paragraphs.length === 0) {
      return full;
    }

    let lastCitationSource = null;
    let contentParagraphsAfterCitation = 0;
    let nextInsertionDistance = randomDistance();
    let previousResultParagraphWasCitation = false;

    const updatedParagraphs = [];

    for (const paragraphXml of paragraphs) {
      const paragraphText = extractParagraphText(paragraphXml);
      const isCitationParagraph = isCitationText(paragraphText);
      const isContentParagraph = isMeaningfulContentParagraph(paragraphXml, paragraphText);

      updatedParagraphs.push(paragraphXml);

      if (isCitationParagraph) {
        previousResultParagraphWasCitation = true;
        contentParagraphsAfterCitation = 0;
        continue;
      }

      if (!isContentParagraph) {
        previousResultParagraphWasCitation = false;
        continue;
      }

      contentParagraphsAfterCitation += 1;

      if (contentParagraphsAfterCitation < nextInsertionDistance || previousResultParagraphWasCitation) {
        previousResultParagraphWasCitation = false;
        continue;
      }

      const citation = generateCitation(normalizedSources, lastCitationSource);
      updatedParagraphs.push(citation.paragraphXml);

      lastCitationSource = citation.sourceIndex;
      contentParagraphsAfterCitation = 0;
      nextInsertionDistance = randomDistance();
      previousResultParagraphWasCitation = true;
    }

    const bodyWithCitations = bodyContent.replace(paragraphPattern, () => updatedParagraphs.shift());
    return `<w:body>${bodyWithCitations}${updatedParagraphs.join("")}</w:body>`;
  });
}

function normalizeSources(sources) {
  if (!Array.isArray(sources)) {
    return [];
  }

  return sources
    .map((source) => {
      if (typeof source === "string") {
        return { title: source, pages: 1 };
      }

      if (!source || typeof source.title !== "string") {
        return null;
      }

      const pages = Number.isInteger(source.pages) && source.pages > 0 ? source.pages : 1;
      return { title: source.title, pages };
    })
    .filter(Boolean);
}

function generateCitation(sources, previousSourceIndex) {
  const sourceIndex = pickSourceIndex(sources.length, previousSourceIndex);
  const source = sources[sourceIndex];
  const page = randomInt(1, source.pages);
  const text = `[${sourceIndex + 1}, с. ${page}]`;

  return {
    sourceIndex,
    paragraphXml: `<w:p><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`
  };
}

function pickSourceIndex(sourcesLength, previousSourceIndex) {
  if (sourcesLength <= 1) {
    return 0;
  }

  let candidate = randomInt(0, sourcesLength - 1);

  while (candidate === previousSourceIndex) {
    candidate = randomInt(0, sourcesLength - 1);
  }

  return candidate;
}

function randomDistance() {
  return randomInt(1, 2);
}

function randomInt(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function isMeaningfulContentParagraph(paragraphXml, paragraphText) {
  if (!paragraphText) {
    return false;
  }

  if (isHeadingParagraph(paragraphXml)) {
    return false;
  }

  if (hasPageBreak(paragraphXml)) {
    return false;
  }

  return true;
}

function isHeadingParagraph(paragraphXml) {
  return /w:pStyle\s+w:val="Heading[1-6]"/.test(paragraphXml);
}

function hasPageBreak(paragraphXml) {
  return /<w:br\b[^>]*w:type="page"/.test(paragraphXml);
}

function isCitationText(text) {
  return /^\[\d+,\s*с\.\s*\d+\]$/.test(text.trim());
}

function extractParagraphText(paragraphXml) {
  return [...paragraphXml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)]
    .map((match) => decodeXmlText(match[1]))
    .join(" ")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function decodeXmlText(text) {
  return text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'");
}

function escapeXml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}
