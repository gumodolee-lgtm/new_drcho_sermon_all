const fs = require("fs");
const path = require("path");
const {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeadingLevel,
  LevelFormat,
  Packer,
  PageNumber,
  Paragraph,
  TextRun,
} = require("docx");

const FONT = "Malgun Gothic";
const INPUT_PATH = path.join(__dirname, "분석_생각이_왜_중요한가.md");
const OUTPUT_PATH = path.join(__dirname, "분석_생각이_왜_중요한가.docx");

function parseInline(text, options = {}) {
  const runs = [];
  const parts = text.split(/(\*\*[^*]+\*\*)/g).filter(Boolean);

  for (const part of parts) {
    const isBold = part.startsWith("**") && part.endsWith("**");
    const content = isBold ? part.slice(2, -2) : part;

    runs.push(
      new TextRun({
        text: content,
        bold: isBold || options.bold || false,
        italics: options.italics || false,
        color: options.color,
        size: options.size || 22,
        font: FONT,
      }),
    );
  }

  if (runs.length === 0) {
    runs.push(
      new TextRun({
        text,
        size: options.size || 22,
        font: FONT,
      }),
    );
  }

  return runs;
}

function normalParagraph(text) {
  return new Paragraph({
    spacing: { after: 140, line: 340 },
    children: parseInline(text),
  });
}

function heading(text, level) {
  const sizes = {
    [HeadingLevel.HEADING_1]: 34,
    [HeadingLevel.HEADING_2]: 28,
  };

  return new Paragraph({
    heading: level,
    spacing: {
      before: level === HeadingLevel.HEADING_1 ? 320 : 260,
      after: 140,
    },
    children: [
      new TextRun({
        text,
        bold: true,
        size: sizes[level],
        color: level === HeadingLevel.HEADING_1 ? "1F3C5B" : "2B5D87",
        font: FONT,
      }),
    ],
  });
}

function separator() {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    border: {
      bottom: {
        style: BorderStyle.SINGLE,
        color: "CFCFCF",
        size: 4,
        space: 1,
      },
    },
    children: [],
  });
}

function quoteParagraph(text) {
  return new Paragraph({
    spacing: { before: 60, after: 60, line: 340 },
    indent: { left: 500, right: 300 },
    children: parseInline(text, { italics: true, color: "444444" }),
  });
}

function quoteSourceParagraph(text) {
  return new Paragraph({
    spacing: { after: 120 },
    indent: { left: 500, right: 300 },
    alignment: AlignmentType.RIGHT,
    children: parseInline(text, { size: 20, color: "666666" }),
  });
}

function bulletParagraph(text) {
  return new Paragraph({
    numbering: { reference: "bullet-list", level: 0 },
    spacing: { after: 80, line: 320 },
    children: parseInline(text),
  });
}

function arrowParagraph(text) {
  return new Paragraph({
    spacing: { before: 40, after: 80, line: 340 },
    indent: { left: 240 },
    children: parseInline(text, { bold: false }),
  });
}

function buildParagraphs(markdown) {
  const lines = markdown.replace(/\r\n/g, "\n").split("\n");
  const children = [];

  for (const rawLine of lines) {
    const line = rawLine.trimEnd();
    const trimmed = line.trim();

    if (!trimmed) {
      continue;
    }

    if (trimmed === "---") {
      children.push(separator());
      continue;
    }

    if (trimmed.startsWith("# ")) {
      children.push(heading(trimmed.slice(2), HeadingLevel.HEADING_1));
      continue;
    }

    if (trimmed.startsWith("## ")) {
      children.push(heading(trimmed.slice(3), HeadingLevel.HEADING_2));
      continue;
    }

    if (trimmed.startsWith("> — ")) {
      children.push(quoteSourceParagraph(trimmed.slice(2)));
      continue;
    }

    if (trimmed.startsWith("> ")) {
      children.push(quoteParagraph(trimmed.slice(2)));
      continue;
    }

    if (trimmed.startsWith("- ")) {
      children.push(bulletParagraph(trimmed.slice(2)));
      continue;
    }

    if (trimmed.startsWith("→ ")) {
      children.push(arrowParagraph(trimmed));
      continue;
    }

    children.push(normalParagraph(trimmed));
  }

  return children;
}

function buildDocument(markdown) {
  return new Document({
    styles: {
      default: {
        document: {
          run: {
            font: FONT,
            size: 22,
          },
        },
      },
    },
    numbering: {
      config: [
        {
          reference: "bullet-list",
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: "•",
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: { left: 720, hanging: 360 },
                },
              },
            },
          ],
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1200,
              right: 1200,
              bottom: 1200,
              left: 1200,
            },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({
                    text: "조용기 목사 설교 분석",
                    size: 18,
                    color: "888888",
                    font: FONT,
                  }),
                ],
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ text: "- ", size: 18, color: "888888", font: FONT }),
                  new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "888888", font: FONT }),
                  new TextRun({ text: " -", size: 18, color: "888888", font: FONT }),
                ],
              }),
            ],
          }),
        },
        children: buildParagraphs(markdown),
      },
    ],
  });
}

function main() {
  const markdown = fs.readFileSync(INPUT_PATH, "utf8");
  const doc = buildDocument(markdown);

  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(OUTPUT_PATH, buffer);
    console.log(`DOCX created: ${OUTPUT_PATH}`);
  });
}

main();
