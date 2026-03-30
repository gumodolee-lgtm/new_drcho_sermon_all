const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel,
  LevelFormat, Header, Footer, PageNumber, PageBreak, BorderStyle
} = require("docx");

// ── helpers ──
const FONT = "Malgun Gothic";
const quote = (text, source) => [
  new Paragraph({
    spacing: { before: 100, after: 40 },
    indent: { left: 560, right: 560 },
    children: [
      new TextRun({ text: `"${text}"`, font: FONT, size: 21, italics: true, color: "333333" }),
    ],
  }),
  new Paragraph({
    spacing: { after: 160 },
    indent: { left: 560, right: 560 },
    alignment: AlignmentType.RIGHT,
    children: [
      new TextRun({ text: `— ${source}`, font: FONT, size: 19, color: "666666" }),
    ],
  }),
];

const heading1 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 200 },
    children: [new TextRun({ text, font: FONT, size: 30, bold: true, color: "1B3A5C" })],
  });

const heading2 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 360, after: 160 },
    children: [new TextRun({ text, font: FONT, size: 26, bold: true, color: "2E5F8A" })],
  });

const body = (text, opts = {}) =>
  new Paragraph({
    spacing: { after: 140, line: 360 },
    ...opts,
    children: [new TextRun({ text, font: FONT, size: 22 })],
  });

const emphasis = (text) =>
  new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, font: FONT, size: 21, bold: true, color: "444444" })],
  });

const bullet = (text, ref = "bullets") =>
  new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 60, line: 340 },
    children: [new TextRun({ text, font: FONT, size: 21 })],
  });

const separator = () =>
  new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 } },
    children: [],
  });

const bibleVerse = (ref, text) =>
  new Paragraph({
    numbering: { reference: "dash", level: 0 },
    spacing: { after: 60, line: 340 },
    children: [
      new TextRun({ text: `${ref}`, font: FONT, size: 21, bold: true }),
      new TextRun({ text: ` — ${text}`, font: FONT, size: 21, color: "555555" }),
    ],
  });

// ── Build sections for each reason ──
function reason(num, title, bodyText, quotes, emphasisText, bibleLabel, verses) {
  const children = [
    heading2(`이유 ${num}. ${title}`),
    body(bodyText),
  ];
  for (const q of quotes) {
    children.push(...quote(q.text, q.source));
  }
  if (emphasisText) {
    children.push(emphasis(`반복 강조 표현: ${emphasisText}`));
  }
  if (verses.length > 0) {
    children.push(new Paragraph({
      spacing: { before: 120, after: 60 },
      children: [new TextRun({ text: "관련 성경 구절", font: FONT, size: 21, bold: true, color: "1B3A5C" })],
    }));
    for (const v of verses) {
      children.push(bibleVerse(v.ref, v.text));
    }
  }
  children.push(separator());
  return children;
}

// ── MAIN DOCUMENT ──
const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: FONT, color: "1B3A5C" },
        paragraph: { spacing: { before: 480, after: 200 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: FONT, color: "2E5F8A" },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 1 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
      {
        reference: "dash",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2013", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1440, right: 1300, bottom: 1440, left: 1300 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "조용기 목사 설교 분석: 생각이 왜 중요한가", font: FONT, size: 16, color: "999999" })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "- ", font: FONT, size: 18, color: "999999" }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: "999999" }), new TextRun({ text: " -", font: FONT, size: 18, color: "999999" })],
        })],
      }),
    },
    children: [
      // ─── TITLE PAGE ───
      new Paragraph({ spacing: { before: 3600 }, children: [] }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun({ text: "조용기 목사 설교 분석", font: FONT, size: 44, bold: true, color: "1B3A5C" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 },
        children: [new TextRun({ text: '"생각이 왜 중요한가"', font: FONT, size: 36, bold: true, color: "2E5F8A" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: "1B3A5C", space: 8 } },
        spacing: { before: 200, after: 100 },
        children: [],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 60 },
        children: [new TextRun({ text: "분석 범위: 1,347편 전체 설교 전수 조사 (1981~2011)", font: FONT, size: 22, color: "555555" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 60 },
        children: [new TextRun({ text: '키워드 검색: "생각", "4차원", "상상", "의식", "사고", "마음", "꿈" 등', font: FONT, size: 22, color: "555555" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 60 },
        children: [new TextRun({ text: "심층 분석 대상: 약 30편의 핵심 설교 원문 전수 검토", font: FONT, size: 22, color: "555555" })],
      }),
      new Paragraph({ children: [new PageBreak()] }),

      // ─── REASON 1 ───
      ...reason(1, "생각은 인간의 위대성이 놓여 있는 자리이다",
        "조용기 목사는 인간의 참된 위대함이 육체가 아닌 '생각'에 있다고 일관되게 가르쳤다. 육체적으로 인간은 호랑이나 코끼리보다 약하지만, 생각을 통해 우주의 끝까지, 바다 깊이까지, 원자의 세계까지 미치는 존재라는 것이다.",
        [
          { text: "파스칼이 말한 것처럼 인간은 생각하는 갈대라. 갈대처럼 연약하고 보잘 것 없지만 그러나 인간의 위대성은 어디 있느냐. 인간은 생각하는 것에 그 위대성이 놓여 있는 것입니다. 몸은 형편없어도 그 속에 있는 생각은 거인입니다.", source: '1981.6.21. "나, 나의 생각"' },
          { text: "인간이 지구의 주인이 된 것은 그 생각할 수 있는 능력 때문인 것입니다. 체력으로 말하면 사람은 사자나 곰이나 호랑이나 심지어 소보다도 힘이 약합니다. 그럼에도 불구하고 우리는 오늘 알려진 이 지구를 완전히 정복한 것은 인간이 생각할 수 있는 위대한 능력 때문인 것입니다.", source: '1984.11.11. "생각을 바꿔라"' },
          { text: "인간에게 있어서의 위대성은 그 몸에 있지 않고 생각에 있으며 인간의 몸은 거인이 아니나 그 생각은 위대한 거인인 것입니다.", source: '1981.6.21. "나, 나의 생각"' },
        ],
        '"생각은 거인이다", "생각하는 것에 위대성이 놓여 있다", "생각할 수 있는 능력"',
        "관련 성경 구절",
        [
          { ref: "잠언 23:7", text: "그 마음의 생각이 어떠하면 그 사람도 그러하다" },
          { ref: "잠언 4:23", text: "지킬만한 것보다 네 마음을 지켜라 생명의 근원이 이에서 남이니라" },
        ]
      ),

      // ─── REASON 2 ───
      ...reason(2, "생각이 곧 그 사람 자체이며, 미래를 결정한다",
        '조용기 목사의 가장 근본적인 명제 중 하나는 "생각이 곧 그 사람"이라는 것이다. 병든 생각은 병든 인간을, 가난한 생각은 가난한 인간을, 승리의 생각은 승리하는 인간을 만든다고 반복적으로 설교했다.',
        [
          { text: "인간이 스스로 생각이 약하면 그는 약한 인간이 되고 생각이 병들면 병든 인간이 되고 생각이 악한 사람이면 악한 인간이 되어 버리고 마는 것입니다.", source: '1981.6.21. "나, 나의 생각"' },
          { text: "마음속에 나는 망한다고 생각하고 있으면서 미래에 흥하는 사람 없으며 마음속에 나는 죽을 것이라고 생각하고 있으면서 미래에 살 수 있는 사람 없으며 마음속에서 나는 패배자라고 생각하면서도 미래에 그 사람이 승리자로써 나타날 수가 없습니다.", source: '1981.6.21. "나, 나의 생각"' },
          { text: "여러분의 진실한 모습은 바로 여러분 속에 있는 그 생각인 것입니다. 생각이 여러분 진실한 자신이며 여러분 삶의 자원입니다.", source: '1981.6.21. "나, 나의 생각"' },
        ],
        '"그 생각이 그 사람이다", "생각은 진실한 자신", "생각이 삶의 자원"',
        "관련 성경 구절",
        [
          { ref: "잠언 23:7", text: "그 마음의 생각이 어떠하면 그 사람도 그러하다" },
          { ref: "창세기 8:21", text: "사람의 마음의 계획하는 바가 어려서부터 악함이라" },
        ]
      ),

      // ─── REASON 3 ───
      ...reason(3, "생각은 하나님께서 역사하시는 통로(그릇)이다",
        '조용기 목사는 하나님이 인간의 생각을 통해 역사하신다고 일관되게 가르쳤다. 에베소서 3:20을 근거로 "구하는 것이나 생각하는 것"에 넘치도록 하시는 하나님이시므로, 생각이 하나님 역사의 그릇이 된다는 것이다.',
        [
          { text: "하나님도 여러분의 생각을 통해서 역사합니다. 성경에는 우리의 온갖 구하는 것이나 생각하는 것에 넘치도록 능히 하실 하나님이라고 말씀하고 있는 것입니다.", source: '1981.6.21. "나, 나의 생각"' },
          { text: "여러분과 저의 생각이 확실해져야 이 생각의 그릇을 통해서 하나님께서 넘치도록 여러분과 나의 생활 속에 역사 해 주는 것입니다. 여러분께서 생각에 부정적인 것을 꽉 채워 놓고 난 다음에 아무리 주님께 철야하면서 기도해도 하나님께서는 응답할 그릇이 없기 때문에 응답하지를 못합니다.", source: '1983.11.27. "긍정적인 생각의 축복"' },
          { text: "성령은 우리의 기도하는 것과 우리의 변화된 의식을 통해서 역사 하는 것입니다. 생각이 달라지지 아니하면 아무리 우리 속에서 역사 하려고 해도 역사 할 수가 없습니다.", source: '1981.8.30. "의식혁명"' },
        ],
        '"생각의 그릇", "생각을 통해 역사하신다", "생각이 달라지지 않으면 역사할 수 없다"',
        "관련 성경 구절",
        [
          { ref: "에베소서 3:20", text: "우리 가운데서 역사하시는 능력대로 우리의 온갖 구하는 것이나 생각하는 것에 더 넘치도록 능히 하실 이에게" },
          { ref: "시편 81:10", text: "네 입을 넓게 열라 내가 채우리라 (생각의 입을 넓게 열라는 해석)" },
        ]
      ),

      // ─── REASON 4 ───
      ...reason(4, "생각은 4차원의 영적 세계를 움직이는 힘이다",
        "'4차원의 영성'은 조용기 목사 신학의 핵심이다. 3차원(물질, 시간, 공간)을 지배하는 4차원의 세계가 있으며, 그 4차원의 구성요소가 바로 생각, 꿈(상상), 믿음, 입술의 고백이라고 가르쳤다.",
        [
          { text: "3차원의 세계는 시간과 공간과 물질 이것을 말하는데 4차원의 세계는 그 위에 그것을 만든 하나님의 생각과 하나님의 꿈과 하나님의 믿음과 하나님의 말씀에 있는 것입니다.", source: '2010.7.18. "성경적 믿음이란 무엇인가?"' },
          { text: "높은 차원이 낮은 차원을 다스립니다. 입체는 사차원 생각이 다스립니다. 꿈이 다스립니다. 믿음이 다스립니다. 말이 다스립니다.", source: '2010.11.7. "사차원의 삶"' },
          { text: "첫째로, 생각이 달라져야 된다. 둘째로, 꿈을 꾸어야 된다. 셋째로, 불가능을 믿어야 된다. 넷째로, 입으로 없는 것을 있는 것같이 시인해야 된다.", source: '2010.7.18. "성경적 믿음이란 무엇인가?"' },
        ],
        '"4차원이 3차원을 다스린다", "생각, 꿈, 믿음, 말", "사차원의 영성"',
        "관련 성경 구절",
        [
          { ref: "히브리서 11:3", text: "보이는 것은 나타난 것으로 말미암아 된 것이 아니니라" },
          { ref: "창세기 1:1-3", text: "하나님이 빛이 있으라 하시니 빛이 있었다 (4차원의 말씀이 3차원을 창조)" },
          { ref: "로마서 4:17", text: "없는 것을 있는 것으로 부르시는 이시니라" },
        ]
      ),

      // ─── REASON 5 ───
      ...reason(5, "생각을 다스리는 자가 운명과 환경을 다스린다",
        "생각이 감정을 격동시키고, 의지를 지배하며, 결국 운명과 환경 전체를 만들어낸다는 것이 조용기 목사의 핵심 논리 구조이다.",
        [
          { text: "여러분의 생각을 잘 다스리지 못하면 여러분의 감정과 의지를 다스릴 수가 없고 여러분의 생사 화복이 잘못된 생각으로 말미암아 파탄을 가져올 수 있는 그러한 길로 이끌어갈 수가 있습니다.", source: '1983.2.27. "저와 함께 한 자보다 많으니라"' },
          { text: "우리 눈에 보이지 않는 생각을 다스리는 것이 우리 운명을 다스리는 것입니다.", source: '2010.5.23. "마음의 생각을 지켜라"' },
          { text: "마음에 있는 그것이 현실로 나타난다. 마음에 불평이 있으면 생활에 불평이 실제로 나타납니다. 마음에 기쁨이 있으면 생활에 실제 기뻐할 일이 생겨나는 것입니다.", source: '2010.4.25. "감사와 찬송의 힘"' },
        ],
        '"생각을 다스리는 것이 운명을 다스리는 것", "마음에 가득한 것이 밖으로 나온다"',
        "관련 성경 구절",
        [
          { ref: "잠언 4:23", text: "지킬만한 것보다 네 마음을 지키라 생명의 근원이 이에서 남이니라" },
          { ref: "잠언 16:32", text: "자기의 마음을 다스리는 자는 성을 빼앗는 자보다 나으니라" },
          { ref: "누가복음 6:45", text: "마음에 가득한 것을 입으로 말함이니라" },
        ]
      ),

      // ─── REASON 6 ───
      ...reason(6, "생각의 변화가 곧 회개(메타노이아)이며 구원의 시작이다",
        "조용기 목사는 '회개'의 헬라어 원어 '메타노이아'를 생각의 전환으로 해석하며, 참된 회개란 단순한 뉘우침이 아니라 생각 전체가 바뀌는 것이라고 가르쳤다.",
        [
          { text: "회개하라 천국이 가까이 왔다. 마음의 생각을 바꾸라 의식을 바꾸라 천국이 손 닫는데 와 있다. 천국이 아무리 손 닫는데 와 있어도 의식이 바꾸어지지 아니하면 우리 속에 들어오지 못합니다.", source: '1981.8.30. "의식혁명"' },
          { text: "회개란 메타노이아란 헬라어로는 죄를 지은 것을 통회하고 자복하는 것뿐 아니라 우리 마음의 생각을 변화시키라는 말인 것입니다.", source: '1988.6.12. "마음의 변화"' },
        ],
        '"회개 = 메타노이아 = 생각의 변화", "마음을 고쳐먹는 것"',
        "관련 성경 구절",
        [
          { ref: "마태복음 4:17", text: "회개하라 천국이 가까이 왔느니라" },
          { ref: "로마서 12:2", text: "마음을 새롭게 함으로 변화를 받아" },
        ]
      ),

      // ─── REASON 7 ───
      ...reason(7, "생각은 눈에 보이지 않는 가장 큰 재산이자 자원이다",
        "조용기 목사는 생각을 눈에 보이지 않지만 황금보화로도 바꿀 수 없는 최고의 재산이자 자원이라고 반복해서 강조했다.",
        [
          { text: "우리 생각하는 것은 눈에 안보이는 가장 큰 재산이요 자원입니다. 이 가장 큰 재산과 자원을 우리가 낭비해서야 되겠습니까? 눈에 보이는 것만 재산이고 자원이라고 생각하면 오해입니다. 눈에 안보이는 이 생각하는 것이 진짜 재산이요 진짜 자원인 것입니다.", source: '1984.11.11. "생각을 바꿔라"' },
          { text: "인간에게 있어서 생각할 수 있는 자원이 고갈하지 않는 이상 인류는 멸망하지 않습니다.", source: '1984.11.11. "생각을 바꿔라"' },
          { text: "인간이 자본고갈로 절대로 멸망하지 않습니다. 인간의 아이디어가 고갈할 때 인간은 멸망하고 마는 것입니다.", source: '1984.11.11. "생각을 바꿔라"' },
        ],
        '"눈에 안보이는 가장 큰 재산", "황금보화로도 바꿀 수 없는 자원"',
        "관련 성경 구절",
        [
          { ref: "잠언 4:23", text: "지킬만한 것보다 네 마음을 지키라 생명의 근원이 이에서 남이니라" },
          { ref: "에베소서 3:20", text: "구하는 것이나 생각하는 것에 더 넘치도록 능히 하실 이에게" },
        ]
      ),

      // ─── REASON 8 ───
      ...reason(8, "하나님의 형상대로 지음받은 생각은 창조적 능력을 지닌다",
        "하나님이 인간을 자신의 형상대로 지으셨기에, 인간의 생각에는 하나님처럼 창조하고 지배할 수 있는 능력이 내재되어 있다는 것이다.",
        [
          { text: "하나님께서 인간을 지으셨을 때 하나님은 인간 속에 하나님처럼 생각할 수 있도록 생각을 집어넣어 주신 것입니다.", source: '1981.6.21. "나, 나의 생각"' },
          { text: "우리 영이 하나님과 닮았고 우리 도덕성이 하나님과 닮았고 우리의 생각하는 것이 하나님의 형상과 모양을 따라 지음을 받았기 때문에 여러분은 지금 하나님처럼 생각하는 것입니다. 하나님처럼 위대하게 생각하고 하나님처럼 창조적으로 생각하는 것입니다.", source: '1981.6.21. "나, 나의 생각"' },
        ],
        '"하나님처럼 생각하도록 지음받았다", "하나님의 형상 = 생각의 능력"',
        "관련 성경 구절",
        [
          { ref: "창세기 1:26-27", text: "하나님이 자기 형상 곧 하나님의 형상대로 사람을 창조하시되" },
          { ref: "요한복음 4:24", text: "하나님은 영이시니" },
        ]
      ),

      // ─── REASON 9 ───
      ...reason(9, "생각이 바뀌면 꿈이 바뀌고, 믿음이 바뀌고, 말이 바뀌고, 운명이 바뀐다",
        "조용기 목사는 생각 → 꿈 → 믿음 → 말 → 환경/운명이라는 연쇄 구조를 거의 모든 설교에서 반복했다. 이 연쇄의 출발점이 바로 '생각'이기에 생각이 중요하다는 것이다.",
        [
          { text: "생각을 바꾸면 믿음이 달라진다. 믿음이 달라지면 기대가 달라진다. 기대가 달라지면 태도가 달라진다. 태도가 달라지면 행동이 달라진다.", source: '2009.8.2. "기도와 낙망" (존 맥스웰 인용)' },
          { text: "마음의 생각을 바꾸고 꿈을 바꾸고 믿음을 바꾸고 말을 바꾸면 운명을 바꿀 수가 있기 때문인 것입니다.", source: '2005.9.11. "패배자는 설 곳이 없다"' },
          { text: "생각을 고치고 꿈을 고치고 믿음을 고치고 말을 고치고 삶이 고쳐지는 것입니다.", source: '2010.11.7. "사차원의 삶"' },
        ],
        '"생각→꿈→믿음→말→운명", "생각을 바꾸면 운명이 바뀐다"',
        "관련 성경 구절",
        [
          { ref: "로마서 10:17", text: "믿음은 들음에서 나며 들음은 그리스도의 말씀으로 말미암았느니라" },
          { ref: "마가복음 11:23", text: "그 말하는 것이 이루어질 줄 믿고 마음에 의심하지 아니하면 그대로 되리라" },
          { ref: "잠언 18:21", text: "죽고 사는 것이 혀의 힘에 달렸나니" },
        ]
      ),

      // ─── REASON 10 ───
      ...reason(10, "부정적인 생각은 마귀의 역사이며, 영적 파괴의 시작이다",
        "조용기 목사는 부정적, 파괴적, 절망적 생각이 단순한 심리적 문제가 아니라 마귀가 인간의 생각에 부패의 씨를 심은 결과라고 가르쳤다. 따라서 생각을 지키는 것은 영적 전쟁의 핵심이다.",
        [
          { text: "마귀는 와서 사람의 육체를 파괴하려고 달라 들지 않았습니다. 인간의 생각에 부패가 심어지도록 한 것입니다.", source: '1981.6.21. "나, 나의 생각"' },
          { text: "열등의식에 잡힌 사람, 피해의식, 가난의식, 패배의식 등 수많은 부정적인 의식이 인격을 붙잡고 있습니다. 그래서 이와 같이 부정적인 의식이 우리를 점령하면 그 사람의 인격적인 삶이 부정적이 되고 그 행위가 부정적이고 파괴적이 되는 것입니다.", source: '1981.8.30. "의식혁명"' },
          { text: "생각이 가난으로 꽉 들어찼는데 어떻게 부요를 가져오며 생각이 실패로 꽉 들어찼는데 어떻게 성공을 가져오며 생각이 패배로 꽉 들어찼는데 어떻게 승리를 가져옵니까?", source: '1981.8.30. "의식혁명"' },
        ],
        null,
        "관련 성경 구절",
        [
          { ref: "창세기 3:1-6", text: "뱀이 하와의 생각을 유혹한 사건" },
          { ref: "고린도후서 10:5", text: "모든 생각을 사로잡아 그리스도에게 복종하게 하리라" },
        ]
      ),

      // ─── REASON 11 ───
      ...reason(11, "바라봄(상상)을 통해 생각이 믿음을 잉태시킨다",
        "'바라봄의 법칙'은 조용기 목사 신학에서 생각과 믿음을 연결하는 핵심 원리이다. 말씀을 듣고 이루어진 모습을 마음에 생생하게 그려보면(상상하면), 그 상상을 통해 믿음이 생겨나고, 그 믿음이 현실을 창조한다는 것이다.",
        [
          { text: "저는 항상 꿈을 품고 기도합니다. 저는 무엇을 위해서 기도하든지 먼저 마음속에 그 이루어진 모습을 분명히 마음속에 그려보고 그것을 상상해 보고 꿈꾸어 보고 그리고 난 다음에 그것이 마음에 뚜렷할 때 그 꿈을 가지고 언제나 기도했고 꿈을 가지고 기도할 때 하나님의 역사로 믿음이 마음속에 생겨났었습니다. 꿈이 없이 기도할 때 믿음은 생겨나지 않습니다.", source: '1986.4.13. "믿는 방법"' },
          { text: "여러분의 마음의 꿈이 비로서 여러분 마음속에 믿음을 만들어 내는 것입니다. 사람이 믿음이 약한 사람은 꿈이 없는 사람이요, 꿈을 안꾸는 사람은 믿음도 가질 수가 없는 것입니다.", source: '1988.6.12. "마음의 변화"' },
          { text: "아브라함에게 꿈과 상상력을 발동하게 하셔서 자신의 자손이 하늘의 별들만큼 많을 것을 그리게 한 것입니다.", source: '1992.3.15. "믿음의 3요소"' },
        ],
        '"바라봄의 법칙", "이루어진 모습을 마음에 그려보라", "꿈은 믿음을 산출한 어머니"',
        "관련 성경 구절",
        [
          { ref: "창세기 15:5", text: "하늘을 우러러 뭇별을 셀 수 있나 보라... 네 자손이 이와 같으리라" },
          { ref: "마가복음 11:24", text: "무엇이든지 기도하고 구하는 것은 받은 줄로 믿으라 그리하면 너희에게 그대로 되리라" },
          { ref: "히브리서 11:1", text: "믿음은 바라는 것들의 실상이요 보이지 않는 것들의 증거니" },
        ]
      ),

      // ─── REASON 12 ───
      ...reason(12, "생각이 몸의 건강과 질병을 좌우한다",
        "조용기 목사는 생각이 영적, 심리적 영역만이 아니라 육체적 건강에도 직접적 영향을 미친다고 가르쳤다. 현대 뇌과학의 연구 결과를 인용하여 이를 뒷받침하기도 했다.",
        [
          { text: "사람이 스트레스를 받거나 화가 나면 분노라는 정보가 뇌로 전달되어 '노르아드레날린'이라는 호르몬을 분비합니다. 이 때 분비되는 노르아드레날린은 매우 강한 독성을 가지고 있어서 이것이 많이 분비되면 우리 몸은 병에 걸리거나 노화가 촉진되어 그만큼 빨리 죽게 되는 것입니다. 반대로 미소를 띠고 사물을 긍정적으로 생각하면 뇌세포를 활성화하고 육체를 이롭게 하는 '베타-엔돌핀'이라는 유익한 호르몬이 분비됩니다.", source: '2007.11.25. "없는 것을 있는 것 같이"' },
          { text: "생각이 우리 몸도 좌우하고 환경도 좌우하고 운명을 좌우하는 것입니다.", source: '2007.11.25. "없는 것을 있는 것 같이"' },
        ],
        '"생각이 몸도 좌우한다", "만병의 근원이 마음에서 나온다"',
        "관련 성경 구절",
        [
          { ref: "잠언 17:22", text: "마음의 즐거움은 양약이라도 심령의 근심은 뼈를 마르게 하느니라" },
          { ref: "잠언 18:14", text: "사람의 심령은 그 병을 능히 이기려니와 심령이 상하면 그것을 누가 일으키겠느냐" },
        ]
      ),

      // ─── REASON 13 ───
      ...reason(13, "말씀이 생각을 변화시키고, 변화된 생각이 삶을 변화시킨다",
        "조용기 목사는 생각 변화의 유일한 근거이자 원천이 하나님의 말씀이라고 가르쳤다. 이유 없는 긍정은 허무하며, 십자가의 구속과 말씀이 생각을 바꿀 조건과 이유가 된다는 것이다.",
        [
          { text: "생각하기 위해서는 생각할 수 있는 조건과 이유가 있어야 되는 것입니다. 이유 없는 생각은 그것은 허무한 것입니다.", source: '1981.6.21. "나, 나의 생각"' },
          { text: "하나님의 생각이 여러분의 생각에 접붙여지는 것입니다. 성경을 읽음으로, 설교를 들음으로 하나님의 생각이 내 생각에 접붙여져서 그래서 하나님의 생각이 내 마음속에 뿌리를 내릴 때 하나님께서 역사하시는 것입니다.", source: '1983.11.27. "긍정적인 생각의 축복"' },
          { text: "성경이 왜 우리에게 중요하냐면 성경이 여러분의 생각을 바꿀 수 있는 하나님의 말씀인 것입니다. 하나님 말씀을 읽으면 여러분 생각이 바꿔져요. 생각이 달라지면 운명이 달라져요.", source: '2010.4.25. "감사와 찬송의 힘"' },
        ],
        '"말씀이 생각을 변화시킨다", "하나님의 생각이 내 생각에 접붙여진다", "십자가가 생각 변화의 근거"',
        "관련 성경 구절",
        [
          { ref: "요한복음 8:32", text: "진리를 알지니 진리가 너희를 자유케 하리라" },
          { ref: "로마서 10:17", text: "믿음은 들음에서 나며 들음은 그리스도의 말씀으로 말미암았느니라" },
        ]
      ),

      // ─── REASON 14 ───
      ...reason(14, "잘못된 생각은 개인뿐 아니라 가정, 사회, 국가를 파멸시킨다",
        "조용기 목사는 생각의 영향력을 개인 차원을 넘어 사회, 국가, 민족 차원으로 확장하여 가르쳤다.",
        [
          { text: "인간이 오늘날 생각을 올바르게 갖게 될 때 발전하고 향상하며 행복을 가져올 수 있습니다. 그러나 잘못된 생각은 개인이나 가정이나 나라나 세계를 파멸로 이끌어 들일 수 있는 것입니다.", source: '1984.11.11. "생각을 바꿔라"' },
          { text: "꿈이 없는 백성은 망한다.", source: '1982.6.27. "꿈과 승리의 생활" (잠언 29:18 인용)' },
        ],
        null,
        "관련 성경 구절",
        [
          { ref: "민수기 13-14장", text: "12정탐꾼 사건 - 10명의 부정적 생각이 300만 민족의 패망을 가져옴" },
          { ref: "잠언 29:18", text: "묵시(꿈)가 없으면 백성이 방자히 행하느니라" },
        ]
      ),

      // ─── CONCLUSION ───
      new Paragraph({ children: [new PageBreak()] }),
      heading1("결론"),
      body('조용기 목사의 1,347편 설교를 전수 조사한 결과, "생각이 왜 중요한가"에 대한 그의 가르침은 다음과 같이 요약된다.'),
      new Paragraph({ spacing: { after: 80 }, children: [] }),

      body("첫째, 생각은 인간 존재의 본질이다. 인간의 위대함은 육체가 아니라 생각에 있으며, \"그 마음의 생각이 어떠하면 그 사람도 그러하다\"(잠 23:7). 생각이 곧 그 사람 자체이며, 그 사람의 현재와 미래를 결정한다."),
      new Paragraph({ spacing: { after: 80 }, children: [] }),

      body("둘째, 생각은 하나님 역사의 통로이다. 하나님은 인간의 \"구하는 것이나 생각하는 것\"(엡 3:20)에 넘치도록 역사하시므로, 생각이 곧 하나님의 역사가 담기는 그릇이다. 부정적 생각으로 채워진 그릇에는 하나님도 역사하실 수 없다."),
      new Paragraph({ spacing: { after: 80 }, children: [] }),

      body("셋째, 생각은 4차원의 영적 세계에 속한다. 3차원의 물질 세계를 지배하는 것은 4차원의 영적 세계이며, 생각은 꿈, 믿음, 고백과 함께 4차원의 핵심 요소이다. 따라서 생각을 다스리는 것이 곧 운명과 환경을 다스리는 것이다."),
      new Paragraph({ spacing: { after: 80 }, children: [] }),

      body("넷째, 생각의 변화는 말씀에 근거해야 한다. 이유 없는 긍정은 허무하며, 십자가의 구속과 하나님의 약속의 말씀이 생각 변화의 유일한 근거이다. 회개(메타노이아) 자체가 생각의 전환이며, 말씀을 통해 하나님의 생각이 내 생각에 접붙여질 때 참된 변화가 일어난다."),
      new Paragraph({ spacing: { after: 80 }, children: [] }),

      body("다섯째, 생각의 변화는 연쇄 반응을 일으킨다. 생각이 바뀌면 꿈(상상)이 바뀌고, 꿈이 바뀌면 믿음이 생겨나고, 믿음이 생기면 입술의 고백이 달라지고, 고백이 달라지면 운명과 환경이 달라진다. 이 연쇄의 출발점이 바로 생각이기에, 조용기 목사는 30년에 걸친 설교에서 한결같이 \"생각을 바꾸라\"고 외쳤다."),

      separator(),

      ...quote(
        "여러분의 미래는 오늘 여러분의 생각에서 나옵니다. 여러분의 운명도 여러분의 생각 속에서 만들어져 나옵니다.",
        '1981.6.21. "나, 나의 생각"'
      ),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  const outPath = "d:\\AI_PROJECT\\_new_drcho_sermon_all\\[Claude Code] 조용기 목사 설교 분석 - 생각이 왜 중요한가.docx";
  fs.writeFileSync(outPath, buffer);
  console.log("DOCX created: " + outPath);
});
