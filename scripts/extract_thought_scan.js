const fs = require("fs");
const path = require("path");
const { TextDecoder } = require("util");

const rootDir = path.resolve(__dirname, "..");
const sourceDir = path.join(rootDir, "sermons_for_lm");
const outputDir = path.join(rootDir, "harness_output");
const jsonPath = path.join(outputDir, "thought_scan.json");
const mdPath = path.join(outputDir, "thought_scan_summary.md");

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function listSermonChunkFiles() {
  return fs
    .readdirSync(sourceDir)
    .filter((name) => /^sermons_\d{3}\.txt$/i.test(name))
    .sort((a, b) => a.localeCompare(b));
}

function scoreKorean(text) {
  const hangulCount = (text.match(/[가-힣]/g) || []).length;
  const mojibakeCount = (text.match(/[�\uFFFD]/g) || []).length;
  return hangulCount - mojibakeCount * 20;
}

function readTextWithFallback(filePath) {
  const bytes = fs.readFileSync(filePath);
  const utf8Text = new TextDecoder("utf-8").decode(bytes);
  const eucKrText = new TextDecoder("euc-kr").decode(bytes);
  const utf8Score = scoreKorean(utf8Text);
  const eucKrScore = scoreKorean(eucKrText);
  return eucKrScore > utf8Score ? eucKrText : utf8Text;
}

function normalizeLine(line) {
  return line.replace(/\uFEFF/g, "").replace(/\r/g, "");
}

function extractRefs(text) {
  const refs = new Set();
  const refRegex = /([가-힣]{1,6}\s?\d{1,3}:\d{1,3}(?:[-~]\d{1,3})?)/g;
  let match = null;
  while ((match = refRegex.exec(text)) !== null) {
    refs.add(match[1].replace(/\s+/g, " ").trim());
  }
  return Array.from(refs);
}

function splitSentences(text) {
  return text
    .split(/(?<=[.!?]|다\.)\s+/)
    .map((part) => part.trim())
    .filter(Boolean);
}

function shortSnippet(paragraph, keyword) {
  const idx = paragraph.indexOf(keyword);
  if (idx < 0) {
    return paragraph.slice(0, 220);
  }
  const start = Math.max(0, idx - 100);
  const end = Math.min(paragraph.length, idx + 160);
  let out = paragraph.slice(start, end).trim();
  if (start > 0) out = `...${out}`;
  if (end < paragraph.length) out = `${out}...`;
  return out;
}

function parseSermonBlock(fileName, sermonFileTag, rawBlock) {
  const lines = rawBlock.split("\n").map(normalizeLine);
  let date = sermonFileTag.replace(".txt", "");
  let title = sermonFileTag.replace(".txt", "");
  let scripture = "";

  const dateLine = lines.find((line) => /^\d{4}[-년]/.test(line.trim()));
  if (dateLine) date = dateLine.trim();

  const titleLine = lines.find((line) => /^\s*제목\s*:/.test(line));
  if (titleLine) title = titleLine.replace(/^\s*제목\s*:\s*/, "").trim();

  const scriptureLine = lines.find((line) => /^\s*성구\s*:/.test(line));
  if (scriptureLine) scripture = scriptureLine.replace(/^\s*성구\s*:\s*/, "").trim();

  const cleanText = lines.join("\n");
  const rawParagraphs = cleanText
    .split(/\n\s*\n+/)
    .map((p) =>
      p
        .split("\n")
        .map((line) => line.trim())
        .filter(Boolean)
        .join("")
        .trim(),
    )
    .filter(Boolean);

  const thoughtParagraphs = [];
  const allRefs = new Set();

  if (scripture) {
    extractRefs(scripture).forEach((ref) => allRefs.add(ref));
  }

  for (const paragraph of rawParagraphs) {
    if (!paragraph.includes("생각")) {
      continue;
    }

    const refs = extractRefs(paragraph);
    refs.forEach((ref) => allRefs.add(ref));

    const sentences = splitSentences(paragraph).filter((s) => s.includes("생각"));
    const quote = sentences.length > 0 ? sentences[0] : shortSnippet(paragraph, "생각");

    thoughtParagraphs.push({
      quote,
      snippet: shortSnippet(paragraph, "생각"),
      refs,
    });
  }

  return {
    chunk_file: fileName,
    sermon_file: sermonFileTag,
    date,
    title,
    scripture,
    thought_count: thoughtParagraphs.length,
    thought_paragraphs: thoughtParagraphs,
    refs: Array.from(allRefs),
  };
}

function parseChunk(fileName, content) {
  const sermons = [];
  const regex = /^===\s+(\d{8}\.txt)\s+===\s*$/gm;
  const markers = [];
  let match = null;

  while ((match = regex.exec(content)) !== null) {
    markers.push({ fileTag: match[1], index: match.index, nextStart: regex.lastIndex });
  }

  for (let i = 0; i < markers.length; i += 1) {
    const current = markers[i];
    const next = markers[i + 1];
    const start = current.nextStart;
    const end = next ? next.index : content.length;
    const block = content.slice(start, end);
    sermons.push(parseSermonBlock(fileName, current.fileTag, block));
  }

  return sermons;
}

function buildThemeStats(sermons) {
  const themes = [
    { id: "identity", pattern: /그 생각이 그 사람|생각이 곧/i },
    { id: "change", pattern: /생각.*바꾸|바꾸.*생각/i },
    { id: "faith", pattern: /믿음|신앙/i },
    { id: "heart", pattern: /마음/i },
    { id: "words", pattern: /말씀|입술|말을/i },
    { id: "dream", pattern: /꿈|환상/i },
    { id: "fourth", pattern: /4차원|사차원/i },
    { id: "negative", pattern: /부정|실패|낙심|절망/i },
    { id: "positive", pattern: /긍정|희망|감사|축복/i },
    { id: "repent", pattern: /회개|메타노이아/i },
  ];

  const stats = Object.fromEntries(themes.map((t) => [t.id, { count: 0, sermons: new Set() }]));

  sermons.forEach((sermon) => {
    if (sermon.thought_count === 0) return;
    const joined = sermon.thought_paragraphs.map((p) => p.snippet).join(" ");
    themes.forEach((theme) => {
      if (theme.pattern.test(joined)) {
        stats[theme.id].count += 1;
        stats[theme.id].sermons.add(`${sermon.date} | ${sermon.title}`);
      }
    });
  });

  const out = {};
  Object.entries(stats).forEach(([k, v]) => {
    out[k] = { count: v.count, sermon_examples: Array.from(v.sermons).slice(0, 8) };
  });
  return out;
}

function makeSummary(data) {
  const lines = [];
  lines.push("# 생각 관련 전수 스캔 요약");
  lines.push("");
  lines.push(`- 전체 설교 수: ${data.total_sermons}`);
  lines.push(`- '생각' 언급 설교 수: ${data.thought_sermons}`);
  lines.push(`- '생각' 문단 총계: ${data.total_thought_paragraphs}`);
  lines.push("");
  lines.push("## 테마 집계");
  lines.push("");
  Object.entries(data.theme_stats).forEach(([theme, stat]) => {
    lines.push(`- ${theme}: ${stat.count}`);
  });
  lines.push("");
  lines.push("## 대표 설교 (생각 문단 상위 40개)");
  lines.push("");

  data.sermons_with_thought
    .slice(0, 40)
    .forEach((sermon) => {
      lines.push(`### ${sermon.date} | ${sermon.title}`);
      if (sermon.scripture) lines.push(`- 성구: ${sermon.scripture}`);
      sermon.thought_paragraphs.slice(0, 3).forEach((p) => {
        lines.push(`- "${p.quote}"`);
      });
      if (sermon.refs.length > 0) {
        lines.push(`- 관련 성경표기: ${sermon.refs.join(", ")}`);
      }
      lines.push("");
    });

  return `${lines.join("\n")}\n`;
}

function main() {
  ensureDir(outputDir);

  const chunkFiles = listSermonChunkFiles();
  const allSermons = [];

  chunkFiles.forEach((chunkFile) => {
    const fullPath = path.join(sourceDir, chunkFile);
    const content = readTextWithFallback(fullPath).replace(/\uFEFF/g, "");
    const sermons = parseChunk(chunkFile, content);
    allSermons.push(...sermons);
  });

  const sermonsWithThought = allSermons.filter((sermon) => sermon.thought_count > 0);
  const totalThoughtParagraphs = sermonsWithThought.reduce(
    (acc, sermon) => acc + sermon.thought_count,
    0,
  );

  const sortedByThoughtCount = [...sermonsWithThought].sort((a, b) => {
    if (b.thought_count !== a.thought_count) return b.thought_count - a.thought_count;
    return a.date.localeCompare(b.date);
  });

  const data = {
    total_sermons: allSermons.length,
    thought_sermons: sermonsWithThought.length,
    total_thought_paragraphs: totalThoughtParagraphs,
    theme_stats: buildThemeStats(sermonsWithThought),
    sermons_with_thought: sortedByThoughtCount,
  };

  fs.writeFileSync(jsonPath, JSON.stringify(data, null, 2), "utf8");
  fs.writeFileSync(mdPath, makeSummary(data), "utf8");

  console.log(`saved: ${jsonPath}`);
  console.log(`saved: ${mdPath}`);
  console.log(
    `total_sermons=${data.total_sermons}, thought_sermons=${data.thought_sermons}, thought_paragraphs=${data.total_thought_paragraphs}`,
  );
}

main();
