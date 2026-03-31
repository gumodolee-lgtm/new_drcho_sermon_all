const path = require("path");
const fs = require("fs");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  PageBreak,
  AlignmentType,
  TableOfContents,
} = require("docx");

const FONT = "Malgun Gothic";
const outputPath = path.join(__dirname, "..", "codex_thought_importance_40pages_with_toc.docx");

const chapters = [
  {
    title: "생각은 미래를 창조하는 재료다",
    desc: "설교는 생각을 단순 심리 상태가 아니라 미래를 만드는 실질적 재료로 규정합니다.",
    quote: "생각은 인생의 미래를 창조하는 재료입니다.",
    source: "2017-06-11",
    repeat: "생각은 미래를 창조한다",
    verses: "고후 4:7-15",
  },
  {
    title: "생각은 일생을 지배한다",
    desc: "생각은 일시적 기분이 아니라 생애 전체를 통제하는 중심축으로 제시됩니다.",
    quote: "우리는 나의 생각이 나의 일생을 지배한다는 것을 우리가 꼭 알아야 되는 것입니다.",
    source: "2017-06-11",
    repeat: "생각이 일생을 지배",
    verses: "고후 4:7-15",
  },
  {
    title: "결과의 차이는 생각의 차이에서 시작된다",
    desc: "동일한 환경에서도 삶의 결론이 갈리는 핵심 원인을 생각의 차이로 해석합니다.",
    quote: "그것은 바로 생각하는 차이 때문입니다.",
    source: "2017-06-11",
    repeat: "생각하는 차이",
    verses: "고후 4:7-15",
  },
  {
    title: "생각을 바꾸면 길이 보인다",
    desc: "막힌 상황을 바꾸는 첫 단추를 환경 변화가 아닌 생각 전환으로 제시합니다.",
    quote: "생각을 바꾸면 길이 보인다. 생각을 바꾸면 길이 보입니다.",
    source: "2017-10-01",
    repeat: "생각을 바꾸면 길이 보인다",
    verses: "사 43:19",
  },
  {
    title: "막다른 길이라는 판단이 돌파를 막는다",
    desc: "문제 자체보다 문제를 규정하는 생각이 믿음의 움직임을 가로막는다고 강조합니다.",
    quote: "문제를 만났을 때, 막다른 길이라고 생각하지 마십시오.",
    source: "2017-10-01",
    repeat: "막다른 길이라고 생각하지 마라",
    verses: "사 43:19",
  },
  {
    title: "생각 방식은 습관이 되므로 훈련이 필요하다",
    desc: "행동 습관만이 아니라 생각 습관이 인생 구조를 만든다는 점을 반복합니다.",
    quote: "그런데 행동뿐 아니라 생각하는 방식에도 습관이 생깁니다.",
    source: "2017-10-01",
    repeat: "생각하는 방식의 습관",
    verses: "시 146:3-5",
  },
  {
    title: "보이지 않는 생각이 현실이 된다",
    desc: "생각은 비가시적이지만 지속적으로 현실로 나타난다는 원리를 제시합니다.",
    quote: "생각은 보이지 않지만 끊임없이 나의 삶 속에 현실로 나타납니다.",
    source: "2010-05-23",
    repeat: "보이지 않지만 현실로 나타남",
    verses: "잠 4:23",
  },
  {
    title: "생각에는 씨앗의 법칙이 있다",
    desc: "현재 생각에 심은 씨앗이 장차 삶의 열매로 거두어진다고 가르칩니다.",
    quote: "현재에 좋은 씨앗을 생각에 심어야 장차 좋은 것을 거두는 것입니다.",
    source: "2010-05-23",
    repeat: "생각에 씨앗을 심어라",
    verses: "잠 4:23",
  },
  {
    title: "미래의 꿈은 현재 생각에 먼저 심어진다",
    desc: "미래 변화는 오늘 생각의 토양에서 시작된다는 점을 분명히 합니다.",
    quote: "미래의 꿈을 현재 내 생각에 심어라.",
    source: "2010-05-23",
    repeat: "꿈을 생각에 심어라",
    verses: "잠 4:23",
  },
  {
    title: "패배는 먼저 생각에서 만들어진다",
    desc: "패배는 사건의 결과가 아니라 먼저 생각의 방향으로 형성된다고 설명합니다.",
    quote: "패배한다고 마음에 생각하면 그 마음이 나가서 그런 환경을 만들어 낸다는 것입니다.",
    source: "2007-01-21",
    repeat: "패배 생각이 패배 환경을 만든다",
    verses: "눅 6:45, 막 11:23",
  },
  {
    title: "생각·꿈·믿음·말은 하나의 연쇄 구조다",
    desc: "생각이 꿈, 믿음, 말로 이어져 결과를 만든다는 구조를 일관되게 제시합니다.",
    quote: "생각과 꿈과 믿음과 말인 것입니다.",
    source: "2007-01-21",
    repeat: "생각-꿈-믿음-말",
    verses: "롬 4:17-18, 막 11:23",
  },
  {
    title: "정체성에 대한 생각이 실제 환경을 바꾼다",
    desc: "내면 정체성 생각이 외부 환경 구현으로 연결된다는 점을 강조합니다.",
    quote:
      "나는 용서받고 의롭게 되고 하나님의 영광을 가진 사람이 되었다고 마음에 생각하고 꿈꾸고 믿고 말하면 ... 나의 환경 속에 이루어져서 ... 나타나게 되는 것입니다.",
    source: "2007-01-21",
    repeat: "정체성 생각의 현실화",
    verses: "롬 4:17-18",
  },
  {
    title: "생각은 지키고 가꿔야 한다",
    desc: "생각은 자동으로 선해지지 않으므로 지키고 길러야 할 대상으로 제시됩니다.",
    quote: "생각을 지키고 가꾸고 꿈을 지키고 가꾸고 믿음을 지키고 가꾸고...",
    source: "2007-01-21",
    repeat: "지키고 가꾸라",
    verses: "잠 16:32",
  },
  {
    title: "생각을 바꾸면 믿음의 결이 달라진다",
    desc: "생각 전환은 믿음의 방향과 강도 변화를 일으킨다고 설명합니다.",
    quote: "생각을 바꾸면 믿음이 달라진다.",
    source: "2009-08-02",
    repeat: "생각 변화 -> 믿음 변화",
    verses: "눅 18:4-8, 마 21:22",
  },
  {
    title: "하나님은 생각하는 것 위에도 역사하신다",
    desc: "하나님의 역사는 기도 요청뿐 아니라 생각 영역까지 포함한다고 선포합니다.",
    quote: "우리의 구하는 것이나 생각하는 것에 넘치도록 능히 하실 하나님이십니다.",
    source: "2009-08-02",
    repeat: "구하는 것과 생각하는 것",
    verses: "엡 3:20-21",
  },
  {
    title: "의식혁명의 핵심은 생각혁명이다",
    desc: "의식이라는 말을 곧 생각으로 정의해 신앙 변화의 본질로 둡니다.",
    quote: "그 의식이란 그 생각하는 것을 말합니다.",
    source: "1981-08-30",
    repeat: "의식 = 생각",
    verses: "눅 23:44-46",
  },
  {
    title: "회개는 생각 전환을 포함한다",
    desc: "회개는 후회 감정이 아니라 생각 구조의 전환이라는 뜻으로 해석됩니다.",
    quote: "회개란 ... 우리 마음의 생각을 변화시키라는 말인 것입니다.",
    source: "1988-06-12",
    repeat: "회개 = 생각 변화",
    verses: "마 4:17, 엡 3:20-21",
  },
  {
    title: "보혈 신앙은 생각을 바꿀 때 실제가 된다",
    desc: "십자가 사건은 생각을 수용할 때 삶에 적용되는 믿음의 실제가 됩니다.",
    quote: "예수의 보혈을 믿고 나의 생각을 바꾸어야 되는 것입니다.",
    source: "1981-08-30",
    repeat: "보혈 믿고 생각을 바꾸라",
    verses: "사 53:4-5",
  },
  {
    title: "생각의 거듭남이 새사람의 길이다",
    desc: "새사람은 생각 혁명으로 시작된다는 점을 반복적으로 강조합니다.",
    quote: "우리의 생각이 거듭나고 생각이 완전히 변화가 되어서 혁명이 일어나서...",
    source: "1981-08-30",
    repeat: "생각이 거듭나고 변화",
    verses: "고후 5:17",
  },
  {
    title: "긍정적 생각은 희망의 사고다",
    desc: "긍정 사고는 막연한 낙관이 아니라 희망의 영적 토대라고 설명됩니다.",
    quote: "절대적으로 긍정적인 생각을 해야 ... 긍정적인 사고라는 것은 희망의 사고인 것입니다.",
    source: "1984-11-11",
    repeat: "절대 긍정적 생각",
    verses: "잠 4:23",
  },
  {
    title: "생각은 황금보다 귀한 자원이다",
    desc: "설교는 생각을 가장 가치 있는 영적 자본으로 규정해 책임을 요청합니다.",
    quote: "우리에게 주신 생각은 황금보화로도 바꿀 수 없는 가장 귀한 재산이요 자원입니다.",
    source: "1984-11-11",
    repeat: "가장 귀한 재산이요 자원",
    verses: "잠 4:23",
  },
  {
    title: "생각은 하나님 형상적 통치 기능과 연결된다",
    desc: "인간의 생각은 하나님의 형상으로서 다스림 기능과 연결된다고 가르칩니다.",
    quote: "여러분의 생각을 ... 하나님처럼 우주를 다스리라고 ... 다 들어있어요.",
    source: "2017-06-11",
    repeat: "생각으로 다스린다",
    verses: "창 1:26-27",
  },
  {
    title: "생각은 꿈 속에서 내일을 만든다",
    desc: "생각은 꿈의 내용이 되어 내일의 방향과 질을 만들어낸다고 설명합니다.",
    quote: "왜냐하면 생각은 꿈속에서 끊임없이 내일을 만들어 가는 것입니다.",
    source: "2017-06-11",
    repeat: "생각이 내일을 만든다",
    verses: "고후 4:7-15",
  },
  {
    title: "긍정 생각은 가정과 공동체에도 파급된다",
    desc: "생각의 영향력은 개인을 넘어 가족·공동체 성패까지 확장됩니다.",
    quote: "생각을 긍정적으로 하는 가족은 ... 성공하고, 부정적으로 하는 사람은 ... 성공하지 못",
    source: "2017-06-11",
    repeat: "생각의 공동체 파급",
    verses: "고후 4:7-15",
  },
  {
    title: "생각-기도-기적의 연동이 강조된다",
    desc: "생각이 믿음과 기도로 이어질 때 변화와 기적이 나타난다고 설교합니다.",
    quote: "할 수 있다고 생각하는 사람이 ... 기도하면 변화와 기적이 나타나게 되는 것입니다.",
    source: "2017-06-11",
    repeat: "생각 -> 기도 -> 기적",
    verses: "막 11:23-24",
  },
  {
    title: "십자가는 생각 전환의 관문이다",
    desc: "생각 전환의 실천 지점으로 십자가를 바라보는 행위를 제시합니다.",
    quote: "생각을 바꾸면 예수님의 십자가를 바라봐야 합니다.",
    source: "2017-10-01",
    repeat: "십자가를 바라보라",
    verses: "사 53:4-5, 막 5:25-34",
  },
  {
    title: "믿음은 생각을 통해 받아 누린다",
    desc: "믿음은 추상 개념이 아니라 생각을 통해 현실에서 누리는 것으로 제시됩니다.",
    quote: "믿음이라는 것은 우리의 생각을 통해서 우리가 받아 누리는 것입니다.",
    source: "2020-07-05",
    repeat: "믿음은 생각을 통해 수용",
    verses: "엡 3:20",
  },
  {
    title: "대속 진리는 생각 수용을 통해 체화된다",
    desc: "화목·치유·저주 청산·영생은 생각으로 받아들여질 때 삶의 실제가 됩니다.",
    quote: "생각을 마음에 받아들여야 되는 것입니다.",
    source: "2020-07-05",
    repeat: "생각을 마음에 받아들이라",
    verses: "사 53:4-5, 갈 3:13, 고후 5:17",
  },
  {
    title: "생각 변화 없이는 믿음의 역사가 어렵다",
    desc: "믿음의 역사는 생각 변화와 분리될 수 없다고 반복적으로 설명합니다.",
    quote: "생각에 변화가 없이는 믿음의 역사가 일어날 수가 없습니다.",
    source: "2020-07-05",
    repeat: "생각 변화 없이는 믿음 역사 없음",
    verses: "롬 10:17",
  },
  {
    title: "말씀의 씨앗이 생각을 변화시킨다",
    desc: "말씀을 읽고 듣고 묵상하는 과정이 생각 변화를 일으키는 핵심 경로로 제시됩니다.",
    quote: "말씀의 씨앗이 여러분의 생각에 뿌려져서 ... 생각이 변화되도록 말씀을 읽고 듣고 묵상하십시오.",
    source: "2020-07-05",
    repeat: "말씀의 씨앗",
    verses: "롬 10:17, 히 4:12",
  },
  {
    title: "생각이 달라지면 꿈도 달라진다",
    desc: "꿈의 수준과 방향은 생각의 변화에서 출발한다고 설교합니다.",
    quote: "이 세상 생각이 달라지면 그 다음 ... 꿈이 달라져야 되는 것입니다.",
    source: "2020-07-05",
    repeat: "생각 변화 -> 꿈 변화",
    verses: "창 15:5",
  },
  {
    title: "하나님 약속을 생각에 새기면 믿음이 자란다",
    desc: "약속을 반복해 생각에 새기는 과정이 믿음의 생성과 성숙을 돕습니다.",
    quote: "그 별들이 아브라함의 생각에 콱 박혔습니다.",
    source: "2020-07-05",
    repeat: "생각에 콱 박힌 약속",
    verses: "창 15:5",
  },
  {
    title: "변화된 생각 위에 성령의 믿음이 부어진다",
    desc: "생각의 바탕이 말씀으로 정렬될 때 성령의 믿음 공급이 강하게 임한다고 설명합니다.",
    quote: "변화된 바탕 위에서 꿈을 가지고 계속 기도하면 ... 성령께서 ... 믿음을 확 주시는 것입니다.",
    source: "2020-07-05",
    repeat: "변화된 바탕 위의 믿음",
    verses: "롬 8:26-27",
  },
  {
    title: "4차원 영성의 첫 관문은 생각이다",
    desc: "4차원 영성은 생각·꿈·믿음·입술 고백의 구조로 설명되며, 생각이 첫 단추로 제시됩니다.",
    quote: "4차원은 ... 생각과 이해, 꿈과 환상, 믿음, 입술의 고백...",
    source: "2010-11-07",
    repeat: "4차원: 생각·꿈·믿음·고백",
    verses: "창 1:1-3, 히 11:3",
  },
  {
    title: "3차원 문제를 4차원 생각으로 풀어야 한다",
    desc: "인간 수단 중심 해법을 넘어 하나님의 생각 중심 해법으로 전환하라고 가르칩니다.",
    quote: "사차원적으로 해결하는 것은 하나님의 생각 ... 말씀, 기도를 통해서 해결하는 것입니다.",
    source: "2010-11-07",
    repeat: "사차원적으로 해결",
    verses: "히 11:3",
  },
  {
    title: "하나님의 생각은 세계 운행의 원리다",
    desc: "보이는 세계의 배후를 하나님의 생각·말씀의 통치로 해석합니다.",
    quote: "하나님의 생각대로 ... 모든 인간 세계는 다 다스려지고 운행하는 것입니다.",
    source: "2010-11-07",
    repeat: "하나님의 생각대로 운행",
    verses: "창 1:1-3",
  },
  {
    title: "하나님의 생각과 사람의 생각은 다르다",
    desc: "신앙의 분별은 하나님의 생각과 인간 생각의 차이를 아는 데서 시작됩니다.",
    quote: "하나님의 생각과 우리의 생각이 다르다는 것을 잘 알아야 합니다.",
    source: "2019-03-03",
    repeat: "하나님의 생각 vs 우리의 생각",
    verses: "사 53:4-5, 사 55:8-9",
  },
  {
    title: "잘못된 생각은 하나님과의 관계를 무너뜨린다",
    desc: "생각의 왜곡이 하나님과의 단절과 삶의 파괴로 이어질 수 있음을 경고합니다.",
    quote: "사람은 잘못된 생각 때문에 하나님과 원수가 되고 하나님께로부터 버림을 받을 수가 있습니다.",
    source: "2019-03-03",
    repeat: "잘못된 생각 때문에",
    verses: "막 8:1-9",
  },
  {
    title: "하나님의 생각은 재앙이 아니라 미래와 희망이다",
    desc: "불안 중심 사고를 하나님의 약속 중심 사고로 전환하라고 반복 권면합니다.",
    quote: "너희를 향한 나의 생각 ... 평안이요 재앙이 아니니라 ... 미래와 희망",
    source: "2015-09-06",
    repeat: "평안, 미래, 희망",
    verses: "렘 29:11-13",
  },
  {
    title: "기도는 마음과 생각을 지키는 통로다",
    desc: "기도는 마음과 생각을 보호하는 실제적 방어선으로 제시됩니다.",
    quote:
      "하나님의 평강이 ... 너희 마음과 생각을 지키시리라",
    source: "2018-09-30",
    repeat: "마음과 생각을 지키시리라",
    verses: "빌 4:6-7",
  },
  {
    title: "말씀 충만은 생각 전쟁의 승리 전략이다",
    desc: "부정적 침투를 막는 가장 강력한 방법으로 말씀으로 생각을 채우는 것을 강조합니다.",
    quote: "생각을 바꾸는 가장 좋은 방법은 우리의 생각을 하나님의 말씀으로 가득 채우는 것입니다.",
    source: "2018-08-19",
    repeat: "생각을 말씀으로 채우라",
    verses: "히 4:12",
  },
  {
    title: "생각 무장은 운명 전환의 시작이다",
    desc: "하나님의 생각으로 무장할 때 부정적 운명 흐름이 끊어지고 새 길이 열린다고 설교합니다.",
    quote: "여러분의 마음이 하나님의 생각으로 완전히 무장되게 하십시오.",
    source: "2020-01-26",
    repeat: "하나님의 생각으로 무장",
    verses: "빌 4:6-7, 히 4:12",
  },
];

function p(text, size = 22, bold = false, align = AlignmentType.LEFT) {
  return new Paragraph({
    alignment: align,
    spacing: { after: 120, line: 300 },
    children: [new TextRun({ text, font: FONT, size, bold })],
  });
}

const selectedChapters = chapters.slice(0, 40);
const children = [];

children.push(
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    spacing: { after: 220 },
    children: [new TextRun({ text: "목차", font: FONT, size: 36, bold: true })],
  }),
);
children.push(
  new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text: "아래 목차는 Word에서 필드 업데이트 시 페이지 번호가 반영됩니다.", font: FONT, size: 18 })],
  }),
);
children.push(
  new Paragraph({
    spacing: { after: 80 },
    children: [new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-2" })],
  }),
);
children.push(
  new Paragraph({
    children: [new PageBreak()],
  }),
);

selectedChapters.forEach((c, i) => {
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { after: 140 },
      children: [new TextRun({ text: `${i + 1}장. ${c.title}`, font: FONT, size: 28, bold: true })],
    }),
  );
  children.push(p(`→ 해당 이유에 대한 상세 설명: ${c.desc}`, 21));
  children.push(p(`→ 설교 원문 근거: "${c.quote}" (${c.source})`, 20));
  children.push(p(`→ 반복 강조된 핵심 표현: ${c.repeat}`, 20));
  children.push(p(`→ 관련 성경 구절: ${c.verses}`, 20));

  if (i === selectedChapters.length - 1) {
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 180, after: 100 },
        children: [new TextRun({ text: "결론", font: FONT, size: 24, bold: true })],
      }),
    );
    children.push(
      p(
        "조용기 목사 설교 전반의 핵심은 ‘생각의 변화가 꿈·믿음·고백·삶의 변화를 이끈다’는 점이며, 4차원 영성의 출발점도 생각임이 일관되게 강조됩니다.",
        20,
      ),
    );
  }

  if (i < selectedChapters.length - 1) {
    children.push(
      new Paragraph({
        children: [new PageBreak()],
      }),
    );
  }
});

const doc = new Document({
  sections: [{ children }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outputPath, buf);
  console.log(outputPath);
  console.log(`chapters=${selectedChapters.length}`);
  console.log("page_breaks_inserted=39");
});
