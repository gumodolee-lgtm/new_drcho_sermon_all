const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel,
  LevelFormat, Header, Footer, PageNumber, BorderStyle,
  Table, TableRow, TableCell, WidthType, TableBorders,
  VerticalAlign, ShadingType
} = require("docx");

const FONT = "Malgun Gothic";
const COLOR = { title: "1B3A5C", heading: "2E5F8A", sub: "3A6F9A", quote: "444444", source: "666666", body: "333333", light: "999999" };

// ── Read and parse the markdown ──
const md = fs.readFileSync("d:/Ai_aigent/_new_drcho_sermon_all/cc_opus_생각의중요성_40장.md", "utf-8");
const lines = md.split("\n");

// ── helpers ──
const mkQuote = (text, source) => {
  const parts = [
    new Paragraph({
      spacing: { before: 80, after: source ? 40 : 160 },
      indent: { left: 560, right: 560 },
      children: [new TextRun({ text: `"${text}"`, font: FONT, size: 20, italics: true, color: COLOR.quote })],
    }),
  ];
  if (source) {
    parts.push(new Paragraph({
      spacing: { after: 160 },
      indent: { left: 560, right: 560 },
      alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: `— ${source}`, font: FONT, size: 18, color: COLOR.source })],
    }));
  }
  return parts;
};

const mkH1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 480, after: 200 },
  children: [new TextRun({ text, font: FONT, size: 32, bold: true, color: COLOR.title })],
});

const mkH2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 400, after: 180 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "2E5F8A", space: 4 } },
  children: [new TextRun({ text, font: FONT, size: 26, bold: true, color: COLOR.heading })],
});

const mkBody = (text) => new Paragraph({
  spacing: { after: 120, line: 340 },
  children: [new TextRun({ text, font: FONT, size: 21, color: COLOR.body })],
});

const mkBold = (text) => new Paragraph({
  spacing: { before: 100, after: 80 },
  children: [new TextRun({ text, font: FONT, size: 21, bold: true, color: COLOR.sub })],
});

const mkBullet = (text) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 60, line: 320 },
  children: [new TextRun({ text, font: FONT, size: 20, color: COLOR.body })],
});

const mkBibleVerse = (text) => new Paragraph({
  numbering: { reference: "dash", level: 0 },
  spacing: { after: 40, line: 320 },
  children: [new TextRun({ text, font: FONT, size: 20, color: COLOR.heading })],
});

const mkSep = () => new Paragraph({
  spacing: { before: 200, after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 } },
  children: [],
});

const mkNumbered = (text) => new Paragraph({
  spacing: { after: 60, line: 320 },
  indent: { left: 400 },
  children: [new TextRun({ text, font: FONT, size: 20, color: COLOR.body })],
});

// ── Parse markdown into docx elements ──
const children = [];

// Title page
children.push(new Paragraph({ spacing: { before: 2400 }, children: [] }));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 200 },
  children: [new TextRun({ text: "조용기 목사 설교 분석", font: FONT, size: 44, bold: true, color: COLOR.title })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 200 },
  children: [new TextRun({ text: '"생각이 왜 중요한가"', font: FONT, size: 36, bold: true, color: COLOR.heading })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 600 },
  children: [new TextRun({ text: "— 40장 완전판 —", font: FONT, size: 28, bold: true, color: COLOR.sub })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 100 },
  children: [new TextRun({ text: "original_sermon 1,753편 전수조사 (1981~2019)", font: FONT, size: 22, color: COLOR.light })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 100 },
  children: [new TextRun({ text: "4개 연대별 병렬 검색 → 핵심 설교 정독 → 40개 주제 분류 → 교차 검증", font: FONT, size: 20, color: COLOR.light })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 100 },
  children: [new TextRun({ text: "설교 원문에 있는 내용만 수록, 임의 추론 및 첨가 배제", font: FONT, size: 20, color: COLOR.light })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 1200 },
  children: [new TextRun({ text: "2026년 3월 · Claude Code (Opus) 분석", font: FONT, size: 20, color: COLOR.light })],
}));
children.push(mkSep());

// Parse body content - skip frontmatter
let i = 0;
// Skip until first ## 제
while (i < lines.length && !lines[i].startsWith("## 제1장")) {
  i++;
}

let inBibleSection = false;

while (i < lines.length) {
  const line = lines[i];

  // Chapter headings: ## 제N장.
  if (line.startsWith("## 제")) {
    inBibleSection = false;
    children.push(mkSep());
    children.push(mkH2(line.replace("## ", "")));
    i++;
    continue;
  }

  // # 결론
  if (line.startsWith("# 결론")) {
    inBibleSection = false;
    children.push(mkSep());
    children.push(mkH1("결론"));
    i++;
    continue;
  }

  // --- separator
  if (line.startsWith("---")) {
    inBibleSection = false;
    children.push(mkSep());
    i++;
    continue;
  }

  // > blockquote lines (sermon quotes)
  if (line.startsWith("> ")) {
    const content = line.replace(/^> /, "").replace(/^"/, "").replace(/"$/, "");
    // Check if next line is source: > — date 「title」
    if (i + 1 < lines.length && lines[i + 1].startsWith("> —")) {
      const source = lines[i + 1].replace(/^> /, "");
      children.push(...mkQuote(content, source));
      i += 2;
      if (i < lines.length && lines[i].trim() === "") i++;
      continue;
    }
    // Check for *-- source* format
    if (i + 1 < lines.length && lines[i + 1].startsWith("> *--")) {
      const source = lines[i + 1].replace(/^> \*-- ?/, "").replace(/\*$/, "");
      children.push(...mkQuote(content, source));
      i += 2;
      if (i < lines.length && lines[i].trim() === "") i++;
      continue;
    }
    // standalone quote
    children.push(...mkQuote(content, ""));
    i++;
    continue;
  }

  // → 관련 성경 구절
  if (line.startsWith("→ 관련 성경 구절")) {
    inBibleSection = true;
    children.push(mkBold("→ 관련 성경 구절"));
    i++;
    continue;
  }

  // Bible verse bullets
  if (line.startsWith("- ") && inBibleSection) {
    children.push(mkBibleVerse(line.replace(/^- /, "")));
    i++;
    continue;
  }

  // Regular bullets
  if (line.startsWith("- ") && !inBibleSection) {
    const text = line.replace(/^- /, "").replace(/\*\*/g, "");
    children.push(mkBullet(text));
    i++;
    continue;
  }

  // Numbered list
  if (/^\d+\. /.test(line)) {
    const text = line.replace(/\*\*/g, "");
    children.push(mkNumbered(text));
    i++;
    continue;
  }

  // Bold emphasis lines
  if (line.startsWith("**") && line.includes("**")) {
    inBibleSection = false;
    children.push(mkBold(line.replace(/\*\*/g, "")));
    i++;
    continue;
  }

  // Table: | ... |
  if (line.startsWith("| ") && line.includes("|")) {
    const tableRows = [];
    while (i < lines.length && lines[i].startsWith("|")) {
      if (!lines[i].startsWith("|---") && !lines[i].startsWith("| ---")) {
        const cells = lines[i].split("|").filter(c => c.trim() !== "").map(c => c.trim());
        tableRows.push(cells);
      }
      i++;
    }
    if (tableRows.length > 0) {
      const colCount = tableRows[0].length;
      const tRows = tableRows.map((row, idx) => {
        const isHeader = idx === 0;
        // Pad row to match column count
        while (row.length < colCount) row.push("");
        return new TableRow({
          children: row.map(cell =>
            new TableCell({
              width: { size: Math.floor(9000 / colCount), type: WidthType.DXA },
              shading: isHeader ? { type: ShadingType.SOLID, color: "E8EFF5" } : undefined,
              verticalAlign: VerticalAlign.CENTER,
              children: [new Paragraph({
                spacing: { before: 40, after: 40 },
                children: [new TextRun({ text: cell, font: FONT, size: 18, bold: isHeader, color: isHeader ? COLOR.title : COLOR.body })],
              })],
            })
          ),
        });
      });

      children.push(new Paragraph({ spacing: { before: 160 }, children: [] }));
      children.push(new Table({
        rows: tRows,
        width: { size: 9000, type: WidthType.DXA },
      }));
      children.push(new Paragraph({ spacing: { after: 160 }, children: [] }));
    }
    continue;
  }

  // Empty line
  if (line.trim() === "") {
    i++;
    continue;
  }

  // Regular body text
  if (!line.startsWith("#") && !line.startsWith(">")) {
    inBibleSection = false;
    const cleaned = line.replace(/\*\*/g, "");
    if (cleaned.trim()) {
      children.push(mkBody(cleaned));
    }
    i++;
    continue;
  }

  i++;
}

// ── Build Document ──
const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 32, bold: true, font: FONT, color: COLOR.title }, paragraph: { spacing: { before: 480, after: 200 } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, font: FONT, color: COLOR.heading }, paragraph: { spacing: { before: 360, after: 160 } } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "dash", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2013", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1300, bottom: 1440, left: 1300 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "Claude Code (Opus) 분석 | 조용기 목사 설교 | 생각이 왜 중요한가 — 40장 완전판", font: FONT, size: 16, color: COLOR.light })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "- ", font: FONT, size: 18, color: COLOR.light }),
            new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: COLOR.light }),
            new TextRun({ text: " -", font: FONT, size: 18, color: COLOR.light }),
          ],
        })],
      }),
    },
    children,
  }],
});

// ── Write file ──
const outPath = "d:/Ai_aigent/_new_drcho_sermon_all/[claude code] 생각이 왜 중요한가 - 40장 완전판.docx";
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outPath, buf);
  process.stdout.write("DOCX created: " + outPath + "\n");
});
