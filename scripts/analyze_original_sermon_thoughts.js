const fs = require("fs");
const path = require("path");

const rootDir = path.resolve(__dirname, "..");
const sourceDir = path.join(rootDir, "original_sermon");
const outputDir = path.join(rootDir, "harness_output");
const outputJson = path.join(outputDir, "thought_importance_candidates.json");
const outputMd = path.join(outputDir, "thought_importance_candidates.md");

const titleKeywords = [
  "생각",
  "의식",
  "마음",
  "사상",
  "사고",
  "상상",
  "사차원",
  "4차원",
  "믿음",
  "꿈",
];

const groups = [
  {
    id: "direct_thought",
    label: "직접적인 생각 언급",
    weight: 10,
    primary: true,
    patterns: [
      /생각/g,
      /생각의 중요/g,
      /생각을 바꾸/g,
      /왜 중요한가/g,
    ],
  },
  {
    id: "consciousness_change",
    label: "의식/마음의 변화",
    weight: 8,
    primary: true,
    patterns: [
      /의식혁명/g,
      /의식을 바꾸/g,
      /의식이 바뀌/g,
      /마음을 지키라/g,
      /마음을 새롭게/g,
      /(마음|생각).{0,30}(변화|바꾸|혁명)/g,
    ],
  },
  {
    id: "thought_scripture",
    label: "생각 관련 핵심 성구",
    weight: 10,
    primary: true,
    patterns: [
      /그 생각이 그리하면 그 사람도 그러하/g,
      /생각하는 것에 더 넘치도록/g,
      /지킬만한 것보다.+마음을 지키라/g,
      /잠언 23장 7절/g,
      /잠언 4장 23절/g,
    ],
  },
  {
    id: "identity_destiny",
    label: "생각과 정체성/운명",
    weight: 9,
    primary: true,
    patterns: [
      /생각이 그 사람/g,
      /(생각|의식).{0,25}(미래|운명|인생)/g,
      /(병든|가난한|패배한|축복의|승리의).{0,20}생각/g,
      /여러분의 진실한 모습은.+생각/g,
    ],
  },
  {
    id: "fourth_dimension",
    label: "사차원 영성과 생각",
    weight: 8,
    primary: true,
    patterns: [
      /(사차원|4차원).{0,40}생각/g,
      /생각.{0,20}(꿈|환상|믿음|고백|말)/g,
      /(꿈|환상|믿음|고백|말).{0,20}생각/g,
      /높은 차원이 낮은 차원을 다스립니다/g,
    ],
  },
  {
    id: "divine_thought",
    label: "하나님의 생각과 접붙임",
    weight: 8,
    primary: true,
    patterns: [
      /하나님의 생각/g,
      /하나님처럼 생각/g,
      /생각을 통해서 역사/g,
      /생각에 넘치도록 능히 하실 하나님/g,
    ],
  },
  {
    id: "repentance_shift",
    label: "회개와 생각 전환",
    weight: 7,
    primary: true,
    patterns: [
      /회개하라/g,
      /회개.{0,30}(생각|마음|의식)/g,
      /메타노이아/g,
      /생각을 고쳐/g,
    ],
  },
  {
    id: "supporting_terms",
    label: "보조 키워드",
    weight: 3,
    primary: false,
    patterns: [
      /마음/g,
      /의식/g,
      /사상/g,
      /사고/g,
      /상상/g,
      /꿈/g,
      /환상/g,
    ],
  },
];

function stripTags(text) {
  return text.replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
}

function readUtf8(filePath) {
  return fs.readFileSync(filePath, "utf8").replace(/\uFEFF/g, "");
}

function getMetadata(text, fileName) {
  const dateMatch = text.match(/^일 시\s*:\s*(.+)$/m);
  const titleMatch = text.match(/^제 목\s*:\s*(.+)$/m);
  const scriptureMatch = text.match(/^말 씀\s*:\s*(.+)$/m);
  return {
    file: fileName,
    date: dateMatch ? stripTags(dateMatch[1]) : fileName.replace(".txt", ""),
    title: titleMatch ? stripTags(titleMatch[1]) : fileName.replace(".txt", ""),
    scripture: scriptureMatch ? stripTags(scriptureMatch[1]) : "",
  };
}

function splitParagraphs(text) {
  return text
    .replace(/\r\n/g, "\n")
    .split(/\n\s*\n+/)
    .map((paragraph) => stripTags(paragraph))
    .filter(Boolean);
}

function makeSnippet(text, index, length) {
  const radius = 110;
  const start = Math.max(0, index - radius);
  const end = Math.min(text.length, index + length + radius);
  const prefix = start > 0 ? "..." : "";
  const suffix = end < text.length ? "..." : "";
  return `${prefix}${text.slice(start, end).trim()}${suffix}`;
}

function analyzeParagraph(paragraph) {
  const matchedGroups = [];
  const snippets = [];

  for (const group of groups) {
    let hitCount = 0;
    for (const pattern of group.patterns) {
      pattern.lastIndex = 0;
      const match = pattern.exec(paragraph);
      if (!match) {
        continue;
      }
      hitCount += 1;
      snippets.push({
        group: group.id,
        label: group.label,
        snippet: makeSnippet(paragraph, match.index, match[0].length),
      });
    }
    if (hitCount > 0) {
      matchedGroups.push({
        id: group.id,
        label: group.label,
        weight: group.weight,
        primary: group.primary,
        hitCount,
      });
    }
  }

  const primaryHits = matchedGroups.filter((group) => group.primary).length;
  const supportingHits = matchedGroups.length - primaryHits;
  const score =
    matchedGroups.reduce((total, group) => total + group.weight * group.hitCount, 0) +
    primaryHits * 2 +
    supportingHits;

  const isCandidate =
    primaryHits > 0 ||
    (supportingHits >= 2 && /마음|의식|사상|사고|상상/.test(paragraph));

  return {
    isCandidate,
    score,
    matchedGroups,
    snippets,
  };
}

function analyzeTitleBonus(title) {
  return titleKeywords.some((keyword) => title.includes(keyword)) ? 10 : 0;
}

function uniqueSnippets(snippets) {
  const seen = new Set();
  return snippets.filter((item) => {
    const key = `${item.group}:${item.snippet}`;
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function analyzeSermon(filePath) {
  const raw = readUtf8(filePath);
  const metadata = getMetadata(raw, path.basename(filePath));
  const paragraphs = splitParagraphs(raw);

  const paragraphResults = paragraphs
    .map((paragraph) => {
      const result = analyzeParagraph(paragraph);
      return { paragraph, ...result };
    })
    .filter((item) => item.isCandidate)
    .sort((a, b) => b.score - a.score);

  if (paragraphResults.length === 0 && analyzeTitleBonus(metadata.title) === 0) {
    return null;
  }

  const topParagraphs = paragraphResults.slice(0, 8);
  const matchedGroupIds = Array.from(
    new Set(topParagraphs.flatMap((item) => item.matchedGroups.map((group) => group.id))),
  );
  const snippets = uniqueSnippets(topParagraphs.flatMap((item) => item.snippets)).slice(0, 12);
  const totalScore =
    analyzeTitleBonus(metadata.title) +
    topParagraphs.reduce((total, item) => total + item.score, 0);

  return {
    ...metadata,
    score: totalScore,
    paragraphCount: topParagraphs.length,
    matchedGroupIds,
    excerpts: snippets,
  };
}

function toMarkdown(report) {
  const lines = [];
  lines.push("# original_sermon 전수 스캔: 생각의 중요성 후보 설교");
  lines.push("");
  lines.push(`- 분석 대상: \`original_sermon\` 폴더의 설교 원문 ${report.totalFiles}편`);
  lines.push(`- 후보 설교 수: ${report.candidateCount}편`);
  lines.push(`- 생성 시각: ${report.generatedAt}`);
  lines.push(`- 입력 제한: 원문 설교 파일만 사용`);
  lines.push("");

  report.candidates.slice(0, 60).forEach((candidate, index) => {
    lines.push(`## ${index + 1}. ${candidate.title}`);
    lines.push("");
    lines.push(`- 파일: \`${candidate.file}\``);
    lines.push(`- 일시: ${candidate.date}`);
    lines.push(`- 말씀: ${candidate.scripture || "(미기재)"}`);
    lines.push(`- 점수: ${candidate.score}`);
    lines.push(`- 매칭 그룹: ${candidate.matchedGroupIds.join(", ")}`);
    lines.push("");
    candidate.excerpts.slice(0, 5).forEach((excerpt) => {
      lines.push(`> [${excerpt.label}] ${excerpt.snippet}`);
      lines.push("");
    });
  });

  return `${lines.join("\n").trim()}\n`;
}

function main() {
  if (!fs.existsSync(sourceDir)) {
    throw new Error(`Source directory not found: ${sourceDir}`);
  }

  const files = fs
    .readdirSync(sourceDir)
    .filter((fileName) => fileName.endsWith(".txt"))
    .map((fileName) => path.join(sourceDir, fileName))
    .sort();

  const candidates = files
    .map((filePath) => analyzeSermon(filePath))
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);

  const report = {
    sourceDir: "original_sermon",
    totalFiles: files.length,
    candidateCount: candidates.length,
    generatedAt: new Date().toISOString(),
    candidates,
  };

  fs.mkdirSync(outputDir, { recursive: true });
  fs.writeFileSync(outputJson, `${JSON.stringify(report, null, 2)}\n`, "utf8");
  fs.writeFileSync(outputMd, toMarkdown(report), "utf8");

  console.log(`Scanned ${files.length} sermons from original_sermon.`);
  console.log(`Found ${candidates.length} candidate sermons.`);
  console.log(`JSON: ${outputJson}`);
  console.log(`Markdown: ${outputMd}`);
  console.log("Top candidates:");
  candidates.slice(0, 10).forEach((candidate, index) => {
    console.log(`${index + 1}. ${candidate.file} | ${candidate.title} | score=${candidate.score}`);
  });
}

main();
