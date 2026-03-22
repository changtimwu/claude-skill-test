const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = '從月薪翻身：一個務實的財務計劃';
pres.author = 'Will 黃士豪 × 博音';

// ─── Color Palette ─────────────────────────────────────────────
const NAVY    = "1E2761";
const CORAL   = "E85A4F";
const GOLD    = "E8B84B";
const LIGHT   = "F7F6F2";
const MID     = "E8E6DF";
const DARK    = "1A1A2E";
const MUTED   = "7A7A8C";
const WHITE   = "FFFFFF";
const NAVY2   = "2F3C7E";

// ─── Helpers ───────────────────────────────────────────────────
function makeShadow() {
  return { type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.12 };
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 1 — Title (dark)
// ═══════════════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: NAVY };

  // Coral accent bar on left
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 5.625,
    fill: { color: CORAL }, line: { color: CORAL, width: 0 }
  });

  // Gold small top strip
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.18, y: 0, w: 9.82, h: 0.06,
    fill: { color: GOLD }, line: { color: GOLD, width: 0 }
  });

  // Episode label
  s.addText("博音 Podcast × Will 黃士豪", {
    x: 0.5, y: 0.55, w: 9, h: 0.35,
    fontSize: 11, color: GOLD, fontFace: "Calibri Light", charSpacing: 3
  });

  // Main title
  s.addText("從月薪翻身", {
    x: 0.5, y: 1.1, w: 9, h: 1.3,
    fontSize: 60, bold: true, color: WHITE, fontFace: "Calibri",
    align: "left"
  });

  s.addText("一個務實的財務計劃", {
    x: 0.5, y: 2.3, w: 9, h: 0.9,
    fontSize: 34, color: GOLD, fontFace: "Calibri Light", align: "left"
  });

  // Divider line
  s.addShape(pres.shapes.LINE, {
    x: 0.5, y: 3.35, w: 4.5, h: 0,
    line: { color: CORAL, width: 1.5 }
  });

  // Subtitle
  s.addText("如何靠月薪與幾乎為零的存款，打造屬於自己的財富自由", {
    x: 0.5, y: 3.55, w: 9, h: 0.5,
    fontSize: 14, color: "AABBD4", fontFace: "Calibri Light", align: "left"
  });

  // Bottom tag
  s.addText("月薪 5 萬 → 財務自由", {
    x: 0.5, y: 4.9, w: 3, h: 0.35,
    fontSize: 11, color: MUTED, fontFace: "Calibri"
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 2 — Will's Story
// ═══════════════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: LIGHT };

  // Title bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.75,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 }
  });
  s.addText("嘉賓 Will 黃士豪的故事", {
    x: 0.4, y: 0, w: 9, h: 0.75,
    fontSize: 22, bold: true, color: WHITE, fontFace: "Calibri", valign: "middle", margin: 0
  });

  // Left column — story text
  const storyItems = [
    "大學快畢業時，父親因糖尿病惡化、失明、洗腎，喪失工作能力",
    "同時發現家中背負近 3,000 萬負債",
    "三分之一來自地下錢莊，還有親友與銀行借款",
    "討債集團上門追債，人生開局跌入谷底",
    "靠著意志力與財務規劃，逐步翻身",
    "現居新加坡，三個孩子就讀國際學校，實現財富與時間自由",
  ];

  storyItems.forEach((item, i) => {
    // Coral circle number
    s.addShape(pres.shapes.OVAL, {
      x: 0.35, y: 0.95 + i * 0.71, w: 0.32, h: 0.32,
      fill: { color: CORAL }, line: { color: CORAL, width: 0 }
    });
    s.addText(String(i + 1), {
      x: 0.35, y: 0.95 + i * 0.71, w: 0.32, h: 0.32,
      fontSize: 11, bold: true, color: WHITE, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
    s.addText(item, {
      x: 0.78, y: 0.93 + i * 0.71, w: 4.8, h: 0.42,
      fontSize: 12.5, color: DARK, fontFace: "Calibri", valign: "middle"
    });
  });

  // Right column — big stat card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.0, y: 0.9, w: 3.6, h: 4.3,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 },
    shadow: makeShadow()
  });

  s.addText("負債金額", {
    x: 6.0, y: 1.1, w: 3.6, h: 0.4,
    fontSize: 13, color: GOLD, fontFace: "Calibri", align: "center"
  });
  s.addText("3,000 萬", {
    x: 6.0, y: 1.45, w: 3.6, h: 0.85,
    fontSize: 44, bold: true, color: WHITE, fontFace: "Calibri", align: "center"
  });
  s.addText("新台幣", {
    x: 6.0, y: 2.25, w: 3.6, h: 0.3,
    fontSize: 12, color: MUTED, fontFace: "Calibri Light", align: "center"
  });

  // Divider
  s.addShape(pres.shapes.LINE, {
    x: 6.4, y: 2.7, w: 2.8, h: 0,
    line: { color: CORAL, width: 1 }
  });

  s.addText("今日成就", {
    x: 6.0, y: 2.9, w: 3.6, h: 0.35,
    fontSize: 13, color: GOLD, fontFace: "Calibri", align: "center"
  });

  const achievements = ["全家移居新加坡", "3 個孩子就讀國際學校", "少時間工作，多時間陪家人"];
  achievements.forEach((a, i) => {
    s.addShape(pres.shapes.OVAL, {
      x: 6.25, y: 3.35 + i * 0.47, w: 0.22, h: 0.22,
      fill: { color: CORAL }, line: { color: CORAL, width: 0 }
    });
    s.addText(a, {
      x: 6.55, y: 3.31 + i * 0.47, w: 2.9, h: 0.3,
      fontSize: 11.5, color: WHITE, fontFace: "Calibri Light", valign: "middle"
    });
  });

  s.addText("如果你的開局沒有比這個更糟，你就贏了。", {
    x: 0.35, y: 5.1, w: 9.3, h: 0.35,
    fontSize: 12, italic: true, color: CORAL, fontFace: "Calibri"
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 3 — 什麼叫「夠了」？(two columns)
// ═══════════════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: LIGHT };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.75,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 }
  });
  s.addText("翻身前，先問自己：什麼叫「夠了」？", {
    x: 0.4, y: 0, w: 9, h: 0.75,
    fontSize: 22, bold: true, color: WHITE, fontFace: "Calibri", valign: "middle", margin: 0
  });

  // Left card — 金錢的主人
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 0.95, w: 4.45, h: 4.35,
    fill: { color: WHITE }, line: { color: MID, width: 1 },
    shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 0.95, w: 4.45, h: 0.55,
    fill: { color: "2C7A4B" }, line: { color: "2C7A4B", width: 0 }
  });
  s.addText("✓  金錢的主人", {
    x: 0.35, y: 0.95, w: 4.45, h: 0.55,
    fontSize: 16, bold: true, color: WHITE, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });

  const master = [
    "清楚知道「夠了」是多少",
    "錢服務我的生活目標",
    "理財計劃圍繞個人願景展開",
    "不受 IG、臉書廣告影響",
    "消費有意識、有方向",
  ];
  master.forEach((t, i) => {
    s.addText("●  " + t, {
      x: 0.5, y: 1.65 + i * 0.6, w: 4.15, h: 0.5,
      fontSize: 13, color: DARK, fontFace: "Calibri"
    });
  });

  // Right card — 金錢的奴隸
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 0.95, w: 4.45, h: 4.35,
    fill: { color: WHITE }, line: { color: MID, width: 1 },
    shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 0.95, w: 4.45, h: 0.55,
    fill: { color: CORAL }, line: { color: CORAL, width: 0 }
  });
  s.addText("✗  金錢的奴隸", {
    x: 5.2, y: 0.95, w: 4.45, h: 0.55,
    fontSize: 16, bold: true, color: WHITE, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });

  const slave = [
    "沒有定義「夠」的標準",
    "看到別人有的就想要",
    "刷 IG 後覺得自己什麼都不夠",
    "錢越多、焦慮越多",
    "消費是為了補償工作的痛苦",
  ];
  slave.forEach((t, i) => {
    s.addText("●  " + t, {
      x: 5.35, y: 1.65 + i * 0.6, w: 4.15, h: 0.5,
      fontSize: 13, color: DARK, fontFace: "Calibri"
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 4 — 財務目標的兩個維度
// ═══════════════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: LIGHT };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.75,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 }
  });
  s.addText("財務目標的兩個維度", {
    x: 0.4, y: 0, w: 9, h: 0.75,
    fontSize: 22, bold: true, color: WHITE, fontFace: "Calibri", valign: "middle", margin: 0
  });

  // Intro text
  s.addText("財務目標 = 你理想生活的具體數字。先定義生活，再推算數字。", {
    x: 0.4, y: 0.9, w: 9.2, h: 0.4,
    fontSize: 14, color: DARK, fontFace: "Calibri Light", italic: true
  });

  // Card 1 — 資產型
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 1.45, w: 4.45, h: 3.75,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 },
    shadow: makeShadow()
  });
  s.addText("01", {
    x: 0.35, y: 1.55, w: 4.45, h: 0.65,
    fontSize: 36, bold: true, color: CORAL, fontFace: "Calibri", align: "center"
  });
  s.addText("資產型目標", {
    x: 0.35, y: 2.1, w: 4.45, h: 0.45,
    fontSize: 18, bold: true, color: WHITE, fontFace: "Calibri", align: "center"
  });
  s.addShape(pres.shapes.LINE, {
    x: 0.9, y: 2.65, w: 3.3, h: 0,
    line: { color: CORAL, width: 1 }
  });
  s.addText("如同企業的「資產負債表」\n\n我擁有哪些資產？", {
    x: 0.5, y: 2.75, w: 4.1, h: 0.7,
    fontSize: 12, color: "AABBD4", fontFace: "Calibri Light"
  });
  const asset = ["房子", "車子", "投資組合", "其他有形資產"];
  asset.forEach((t, i) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 3.55 + i * 0.36, w: 0.08, h: 0.24,
      fill: { color: GOLD }, line: { color: GOLD, width: 0 }
    });
    s.addText(t, {
      x: 0.72, y: 3.5 + i * 0.36, w: 3.8, h: 0.3,
      fontSize: 13, color: WHITE, fontFace: "Calibri"
    });
  });

  // Card 2 — 開銷型
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.45, w: 4.45, h: 3.75,
    fill: { color: NAVY2 }, line: { color: NAVY2, width: 0 },
    shadow: makeShadow()
  });
  s.addText("02", {
    x: 5.2, y: 1.55, w: 4.45, h: 0.65,
    fontSize: 36, bold: true, color: GOLD, fontFace: "Calibri", align: "center"
  });
  s.addText("開銷型目標", {
    x: 5.2, y: 2.1, w: 4.45, h: 0.45,
    fontSize: 18, bold: true, color: WHITE, fontFace: "Calibri", align: "center"
  });
  s.addShape(pres.shapes.LINE, {
    x: 5.75, y: 2.65, w: 3.3, h: 0,
    line: { color: GOLD, width: 1 }
  });
  s.addText("如同企業的「損益表」\n\n我理想生活每月需多少？", {
    x: 5.35, y: 2.75, w: 4.1, h: 0.7,
    fontSize: 12, color: "AABBD4", fontFace: "Calibri Light"
  });
  const expense = ["飲食、服裝、旅行", "陪孩子、健身的時間成本", "興趣與社會貢獻", "不需看價格的自在感"];
  expense.forEach((t, i) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.35, y: 3.55 + i * 0.36, w: 0.08, h: 0.24,
      fill: { color: CORAL }, line: { color: CORAL, width: 0 }
    });
    s.addText(t, {
      x: 5.57, y: 3.5 + i * 0.36, w: 3.8, h: 0.3,
      fontSize: 13, color: WHITE, fontFace: "Calibri"
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 5 — 四桶金策略
// ═══════════════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: LIGHT };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.75,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 }
  });
  s.addText("四桶金策略：建立穩固的財務架構", {
    x: 0.4, y: 0, w: 9, h: 0.75,
    fontSize: 22, bold: true, color: WHITE, fontFace: "Calibri", valign: "middle", margin: 0
  });

  const buckets = [
    { num: "01", name: "保障型基金", sub: "基礎保險 + 緊急備用金", detail: "6～12 個月的生活開銷\n放在銀行，隨時可用\n股市下行時的護城河", color: "2C7A4B", accent: "5EC47A" },
    { num: "02", name: "防守型資產", sub: "穩定、低波動的資產", detail: "政府公債、高股息股票\n提供穩定現金流\n保障生活品質不中斷", color: "1C5A8A", accent: "4A9FD4" },
    { num: "03", name: "進攻型資產", sub: "追求成長報酬率", detail: "年化目標 15～20%\n個股、ETF 組合\n依財務目標決定比例", color: NAVY2, accent: GOLD },
    { num: "04", name: "樂透型部位", sub: "高風險、高潛力", detail: "佔總資產 10～15%\n山寨幣、轉機股等\n輸光也不影響大局", color: "8A3020", accent: CORAL },
  ];

  buckets.forEach((b, i) => {
    const x = 0.22 + i * 2.44;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 0.9, w: 2.14, h: 4.05,
      fill: { color: b.color }, line: { color: b.color, width: 0 },
      shadow: makeShadow()
    });
    // Accent top bar
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 0.9, w: 2.14, h: 0.1,
      fill: { color: b.accent }, line: { color: b.accent, width: 0 }
    });
    s.addText(b.num, {
      x, y: 1.05, w: 2.14, h: 0.6,
      fontSize: 30, bold: true, color: b.accent, fontFace: "Calibri", align: "center"
    });
    s.addText(b.name, {
      x, y: 1.6, w: 2.14, h: 0.55,
      fontSize: 15, bold: true, color: WHITE, fontFace: "Calibri", align: "center"
    });
    s.addText(b.sub, {
      x, y: 2.1, w: 2.14, h: 0.4,
      fontSize: 10.5, color: b.accent, fontFace: "Calibri Light", align: "center"
    });
    s.addShape(pres.shapes.LINE, {
      x: x + 0.2, y: 2.62, w: 1.74, h: 0,
      line: { color: b.accent, width: 0.75 }
    });
    s.addText(b.detail, {
      x: x + 0.1, y: 2.75, w: 1.94, h: 2.1,
      fontSize: 11, color: "CCDDEE", fontFace: "Calibri Light", valign: "top"
    });
  });

  s.addText("重要：開始投資前，必須先建立第一桶金（緊急備用金）！", {
    x: 0.22, y: 5.1, w: 9.55, h: 0.35,
    fontSize: 12, bold: true, color: CORAL, fontFace: "Calibri", align: "center"
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 6 — 複利的威力
// ═══════════════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: DARK };

  // Title
  s.addText("複利的威力：報酬率比本薪更關鍵", {
    x: 0.4, y: 0.25, w: 9.2, h: 0.65,
    fontSize: 26, bold: true, color: WHITE, fontFace: "Calibri"
  });
  s.addShape(pres.shapes.LINE, {
    x: 0.4, y: 0.92, w: 9.2, h: 0,
    line: { color: CORAL, width: 1 }
  });

  s.addText("每月投入 NT$3,000，投資 30 年", {
    x: 0.4, y: 1.05, w: 9.2, h: 0.38,
    fontSize: 15, color: "AABBD4", fontFace: "Calibri Light", italic: true
  });

  // Left stat — 8%
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 1.55, w: 4.4, h: 3.55,
    fill: { color: "252545" }, line: { color: "3A3A6A", width: 1 },
    shadow: makeShadow()
  });
  s.addText("0050 定期定額", {
    x: 0.35, y: 1.65, w: 4.4, h: 0.42,
    fontSize: 14, color: MUTED, fontFace: "Calibri", align: "center"
  });
  s.addText("年化 8%", {
    x: 0.35, y: 2.0, w: 4.4, h: 0.55,
    fontSize: 22, color: "7A9FBF", fontFace: "Calibri", align: "center"
  });
  s.addText("~400 萬", {
    x: 0.35, y: 2.5, w: 4.4, h: 0.9,
    fontSize: 52, bold: true, color: "7A9FBF", fontFace: "Calibri", align: "center"
  });
  s.addText("有差，但難以「翻身」", {
    x: 0.35, y: 3.4, w: 4.4, h: 0.4,
    fontSize: 13, color: MUTED, fontFace: "Calibri Light", align: "center"
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 3.85, w: 4.4, h: 1.15,
    fill: { color: "1A1A35" }, line: { color: "1A1A35", width: 0 }
  });
  s.addText("巴菲特建議的被動指數投資\n適合保守穩健的投資人", {
    x: 0.5, y: 3.92, w: 4.1, h: 0.9,
    fontSize: 11.5, color: MUTED, fontFace: "Calibri Light", align: "center"
  });

  // Right stat — 15%
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.25, y: 1.55, w: 4.4, h: 3.55,
    fill: { color: "2A1A0E" }, line: { color: GOLD, width: 1 },
    shadow: makeShadow()
  });
  // Gold accent top
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.25, y: 1.55, w: 4.4, h: 0.08,
    fill: { color: GOLD }, line: { color: GOLD, width: 0 }
  });
  s.addText("學好投資 / 精選標的", {
    x: 5.25, y: 1.7, w: 4.4, h: 0.42,
    fontSize: 14, color: GOLD, fontFace: "Calibri", align: "center"
  });
  s.addText("年化 15%", {
    x: 5.25, y: 2.05, w: 4.4, h: 0.55,
    fontSize: 22, color: GOLD, fontFace: "Calibri", align: "center"
  });
  s.addText("~1,500 萬", {
    x: 5.25, y: 2.5, w: 4.4, h: 0.9,
    fontSize: 52, bold: true, color: GOLD, fontFace: "Calibri", align: "center"
  });
  s.addText("有感翻身，進入八位數世界", {
    x: 5.25, y: 3.4, w: 4.4, h: 0.4,
    fontSize: 13, color: GOLD, fontFace: "Calibri Light", align: "center"
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.25, y: 3.85, w: 4.4, h: 1.15,
    fill: { color: "1F1408" }, line: { color: "1F1408", width: 0 }
  });
  s.addText("學習投資技能比加薪更划算\n報酬率差距比本薪差距影響更大", {
    x: 5.4, y: 3.92, w: 4.1, h: 0.9,
    fontSize: 11.5, color: GOLD, fontFace: "Calibri Light", align: "center"
  });

  // Arrow between
  s.addText("↑ 差距 3.75 倍", {
    x: 4.35, y: 2.75, w: 1.3, h: 0.4,
    fontSize: 11, bold: true, color: CORAL, fontFace: "Calibri", align: "center"
  });

  s.addText("結論：時間是最重要的，其次是報酬率——不是本薪的高低。", {
    x: 0.35, y: 5.2, w: 9.3, h: 0.3,
    fontSize: 12, color: "AABBD4", fontFace: "Calibri Light", italic: true, align: "center"
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 7 — 如何提高主動收入
// ═══════════════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: LIGHT };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.75,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 }
  });
  s.addText("打破本薪天花板：主動收入的兩種路徑", {
    x: 0.4, y: 0, w: 9, h: 0.75,
    fontSize: 22, bold: true, color: WHITE, fontFace: "Calibri", valign: "middle", margin: 0
  });

  s.addText("傳統加薪有上限；以「價值導向」創造的收入，沒有上限。", {
    x: 0.4, y: 0.9, w: 9.2, h: 0.4,
    fontSize: 14, color: DARK, fontFace: "Calibri Light", italic: true
  });

  // Path 1
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 1.45, w: 4.45, h: 3.85,
    fill: { color: WHITE }, line: { color: MID, width: 1 },
    shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 1.45, w: 4.45, h: 0.08,
    fill: { color: CORAL }, line: { color: CORAL, width: 0 }
  });
  s.addText("路徑 A", {
    x: 0.35, y: 1.53, w: 1.1, h: 0.52,
    fontSize: 11, bold: true, color: WHITE, fontFace: "Calibri",
    align: "center", valign: "middle",
    fill: { color: CORAL }
  });
  // coral label background
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 1.53, w: 1.1, h: 0.52,
    fill: { color: CORAL }, line: { color: CORAL, width: 0 }
  });
  s.addText("路徑 A", {
    x: 0.35, y: 1.53, w: 1.1, h: 0.52,
    fontSize: 12, bold: true, color: WHITE, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });
  s.addText("解決少數人的大問題", {
    x: 1.55, y: 1.55, w: 3.15, h: 0.48,
    fontSize: 16, bold: true, color: DARK, fontFace: "Calibri", valign: "middle"
  });
  s.addText("專業顧問、企業教練、外科手術\n醫師、律師、頂尖講師", {
    x: 0.5, y: 2.15, w: 4.1, h: 0.55,
    fontSize: 12, color: MUTED, fontFace: "Calibri Light"
  });
  s.addShape(pres.shapes.LINE, {
    x: 0.5, y: 2.8, w: 4.0, h: 0,
    line: { color: MID, width: 0.75 }
  });
  s.addText("關鍵：你對問題的理解深度\n決定你的開價能力。", {
    x: 0.5, y: 2.9, w: 4.1, h: 0.6,
    fontSize: 12.5, color: DARK, fontFace: "Calibri"
  });
  s.addText("例：Will 的財務教育課程\n幫助 13 萬付費學員解決財務焦慮", {
    x: 0.5, y: 3.6, w: 4.1, h: 0.6,
    fontSize: 12, color: CORAL, fontFace: "Calibri", italic: true
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.35, w: 4.1, h: 0.75,
    fill: { color: "FFF5F4" }, line: { color: MID, width: 0 }
  });
  s.addText("由你開價，收入沒有天花板", {
    x: 0.6, y: 4.43, w: 3.9, h: 0.55,
    fontSize: 13, bold: true, color: CORAL, fontFace: "Calibri", valign: "middle"
  });

  // Path 2
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.45, w: 4.45, h: 3.85,
    fill: { color: WHITE }, line: { color: MID, width: 1 },
    shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.53, w: 1.1, h: 0.52,
    fill: { color: NAVY2 }, line: { color: NAVY2, width: 0 }
  });
  s.addText("路徑 B", {
    x: 5.2, y: 1.53, w: 1.1, h: 0.52,
    fontSize: 12, bold: true, color: WHITE, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });
  s.addText("解決大量人的小問題", {
    x: 6.4, y: 1.55, w: 3.15, h: 0.48,
    fontSize: 16, bold: true, color: DARK, fontFace: "Calibri", valign: "middle"
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.45, w: 4.45, h: 0.08,
    fill: { color: NAVY2 }, line: { color: NAVY2, width: 0 }
  });
  s.addText("創作者、播客主、線上課程\n演出、軟體工具、媒體平台", {
    x: 5.35, y: 2.15, w: 4.1, h: 0.55,
    fontSize: 12, color: MUTED, fontFace: "Calibri Light"
  });
  s.addShape(pres.shapes.LINE, {
    x: 5.35, y: 2.8, w: 4.0, h: 0,
    line: { color: MID, width: 0.75 }
  });
  s.addText("關鍵：接觸的人越多，\n規模效益越大。", {
    x: 5.35, y: 2.9, w: 4.1, h: 0.6,
    fontSize: 12.5, color: DARK, fontFace: "Calibri"
  });
  s.addText("例：博恩的巡演、Podcast\n每多一個觀眾，邊際成本趨近於零", {
    x: 5.35, y: 3.6, w: 4.1, h: 0.6,
    fontSize: 12, color: NAVY2, fontFace: "Calibri", italic: true
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.35, y: 4.35, w: 4.1, h: 0.75,
    fill: { color: "F0F4FF" }, line: { color: MID, width: 0 }
  });
  s.addText("觸及無上限，收入理論上無上限", {
    x: 5.45, y: 4.43, w: 3.9, h: 0.55,
    fontSize: 13, bold: true, color: NAVY2, fontFace: "Calibri", valign: "middle"
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 8 — 槓桿的正確用法
// ═══════════════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: DARK };

  s.addText("槓桿的正確用法", {
    x: 0.4, y: 0.2, w: 9.2, h: 0.7,
    fontSize: 30, bold: true, color: WHITE, fontFace: "Calibri"
  });
  s.addShape(pres.shapes.LINE, {
    x: 0.4, y: 0.92, w: 9.2, h: 0,
    line: { color: GOLD, width: 1.5 }
  });
  s.addText("「只要不斷頭，貸好貸滿。」— Will 黃士豪", {
    x: 0.4, y: 1.05, w: 9.2, h: 0.4,
    fontSize: 15, color: GOLD, fontFace: "Calibri Light", italic: true
  });

  // Two rules
  const rules = [
    {
      n: "原則 1",
      title: "月付額 ≤ 盈餘的一半",
      body: "月收入 NT$100,000\n正常開銷 NT$50,000\n盈餘 NT$50,000\n\n→ 每月還款金額不得超過 NT$25,000",
      note: "確保即使遇到突發狀況，\n仍有現金流維持生活。",
      color: "1A3A5C"
    },
    {
      n: "原則 2",
      title: "緊急備用金 ≥ 12 個月開銷",
      body: "含貸款後的月支出 NT$75,000\n× 12 個月 = NT$900,000\n\n→ 帳上至少需保有此金額",
      note: "這是防止被迫在最低點\n斷頭賣出資產的最後防線。",
      color: "1A3020"
    },
  ];

  rules.forEach((r, i) => {
    const x = 0.35 + i * 4.9;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.6, w: 4.55, h: 3.65,
      fill: { color: r.color }, line: { color: r.color, width: 0 },
      shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.6, w: 4.55, h: 0.08,
      fill: { color: i === 0 ? "4A9FD4" : "5EC47A" }, line: { color: i === 0 ? "4A9FD4" : "5EC47A", width: 0 }
    });
    s.addText(r.n, {
      x: x + 0.15, y: 1.75, w: 1.2, h: 0.35,
      fontSize: 12, bold: true, color: i === 0 ? "4A9FD4" : "5EC47A", fontFace: "Calibri"
    });
    s.addText(r.title, {
      x: x + 0.15, y: 2.05, w: 4.25, h: 0.52,
      fontSize: 18, bold: true, color: WHITE, fontFace: "Calibri"
    });
    s.addShape(pres.shapes.LINE, {
      x: x + 0.15, y: 2.65, w: 4.15, h: 0,
      line: { color: i === 0 ? "4A9FD4" : "5EC47A", width: 0.75 }
    });
    s.addText(r.body, {
      x: x + 0.15, y: 2.78, w: 4.2, h: 1.6,
      fontSize: 13, color: "CCDDEE", fontFace: "Calibri Light"
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.15, y: 4.45, w: 4.2, h: 0.7,
      fill: { color: "0D0D1A" }, line: { color: "0D0D1A", width: 0 }
    });
    s.addText(r.note, {
      x: x + 0.25, y: 4.5, w: 4.0, h: 0.6,
      fontSize: 11.5, color: i === 0 ? "4A9FD4" : "5EC47A", fontFace: "Calibri Light", italic: true
    });
  });

  s.addText("警告：槓桿是雙面刃 — 上漲放大收益，下跌同樣放大損失。標的選擇是最大的風控。", {
    x: 0.35, y: 5.27, w: 9.3, h: 0.3,
    fontSize: 11.5, color: CORAL, fontFace: "Calibri", align: "center"
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 9 — 房子的正確觀念
// ═══════════════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: LIGHT };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.75,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 }
  });
  s.addText("買房 vs 租房：用正確框架做決定", {
    x: 0.4, y: 0, w: 9, h: 0.75,
    fontSize: 22, bold: true, color: WHITE, fontFace: "Calibri", valign: "middle", margin: 0
  });

  // Two frameworks
  // Left — 自住 = 消費
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 0.88, w: 4.45, h: 4.5,
    fill: { color: WHITE }, line: { color: MID, width: 1 },
    shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 0.88, w: 4.45, h: 0.58,
    fill: { color: "36454F" }, line: { color: "36454F", width: 0 }
  });
  s.addText("自住  =  消費", {
    x: 0.35, y: 0.88, w: 4.45, h: 0.58,
    fontSize: 18, bold: true, color: WHITE, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });

  const selfUse = [
    ["頭期款", "NT$5,000萬 × 30% = NT$1,500萬 一次支出"],
    ["月供壓力", "每月可能超過 NT$10萬 遠高於租金"],
    ["機會成本", "省下NT$1,480萬投資 年化8% = 每月幾十萬收益"],
    ["結論", "目前台灣五都自住：租 > 買 划算"],
  ];

  selfUse.forEach(([label, desc], i) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 1.6 + i * 0.8, w: 1.1, h: 0.5,
      fill: { color: "36454F" }, line: { color: "36454F", width: 0 }
    });
    s.addText(label, {
      x: 0.5, y: 1.6 + i * 0.8, w: 1.1, h: 0.5,
      fontSize: 11, bold: true, color: WHITE, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
    s.addText(desc, {
      x: 1.7, y: 1.62 + i * 0.8, w: 2.95, h: 0.5,
      fontSize: 12, color: DARK, fontFace: "Calibri Light", valign: "middle"
    });
  });

  // Right — 投資 = 看租金報酬
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 0.88, w: 4.45, h: 4.5,
    fill: { color: WHITE }, line: { color: MID, width: 1 },
    shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 0.88, w: 4.45, h: 0.58,
    fill: { color: NAVY2 }, line: { color: NAVY2, width: 0 }
  });
  s.addText("投資  =  看租金報酬率", {
    x: 5.2, y: 0.88, w: 4.45, h: 0.58,
    fontSize: 18, bold: true, color: WHITE, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });

  const invest = [
    ["正確觀念", "房地產合理投資報酬來自租金\n不是價差"],
    ["篩選標準", "租金報酬率是否符合我的財務目標？\n有符合才投，沒有就不投"],
    ["新加坡模式", "政府在豪宅旁建公宅\n解決居住正義問題"],
    ["警示", "只追求價差 → 必然造成居住正義問題"],
  ];

  invest.forEach(([label, desc], i) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.35, y: 1.6 + i * 0.8, w: 1.1, h: 0.5,
      fill: { color: NAVY2 }, line: { color: NAVY2, width: 0 }
    });
    s.addText(label, {
      x: 5.35, y: 1.6 + i * 0.8, w: 1.1, h: 0.5,
      fontSize: 11, bold: true, color: WHITE, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
    s.addText(desc, {
      x: 6.55, y: 1.62 + i * 0.8, w: 2.95, h: 0.5,
      fontSize: 12, color: DARK, fontFace: "Calibri Light", valign: "middle"
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 10 — 生存 vs 發展
// ═══════════════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: LIGHT };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.75,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 }
  });
  s.addText("你在哪個階段？生存 vs 發展", {
    x: 0.4, y: 0, w: 9, h: 0.75,
    fontSize: 22, bold: true, color: WHITE, fontFace: "Calibri", valign: "middle", margin: 0
  });

  s.addText("大多數人的焦慮，其實不在「生存」，而是在「發展」階段擔心生存的事。", {
    x: 0.4, y: 0.88, w: 9.2, h: 0.42,
    fontSize: 13.5, color: DARK, fontFace: "Calibri Light", italic: true
  });

  // Survival
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 1.4, w: 4.45, h: 3.9,
    fill: { color: "3D1010" }, line: { color: CORAL, width: 1.5 },
    shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 1.4, w: 4.45, h: 0.65,
    fill: { color: CORAL }, line: { color: CORAL, width: 0 }
  });
  s.addText("生存階段", {
    x: 0.35, y: 1.4, w: 4.45, h: 0.65,
    fontSize: 20, bold: true, color: WHITE, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });
  const survival = [
    "這個月帳單付不出來",
    "孩子的學費交不出來",
    "下個月的餐費不知在哪",
    "→ 這時候焦慮是合理的",
    "→ 全力專注在「如何活下去」",
    "→ 先從儲蓄 1~2% 開始",
  ];
  survival.forEach((t, i) => {
    s.addText(t, {
      x: 0.55, y: 2.18 + i * 0.5, w: 4.05, h: 0.42,
      fontSize: i < 3 ? 13 : 12.5,
      bold: i >= 3,
      color: i < 3 ? "FFCCCC" : CORAL,
      fontFace: "Calibri"
    });
  });

  // Development
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.4, w: 4.45, h: 3.9,
    fill: { color: "0D2A1A" }, line: { color: "5EC47A", width: 1.5 },
    shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.4, w: 4.45, h: 0.65,
    fill: { color: "2C7A4B" }, line: { color: "2C7A4B", width: 0 }
  });
  s.addText("發展階段", {
    x: 5.2, y: 1.4, w: 4.45, h: 0.65,
    fontSize: 20, bold: true, color: WHITE, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });
  const dev = [
    "可以看 Netflix、喝咖啡",
    "生活不到「沒飯吃」的地步",
    "但仍焦慮：iPhone、車子、房子…",
    "→ 這種焦慮跟錢多少無關",
    "→ 焦慮的底層是恐懼、貪婪、比較、急",
    "→ 聚焦在「如何發展」，不是生存",
  ];
  dev.forEach((t, i) => {
    s.addText(t, {
      x: 5.4, y: 2.18 + i * 0.5, w: 4.05, h: 0.42,
      fontSize: i < 3 ? 13 : 12.5,
      bold: i >= 3,
      color: i < 3 ? "AADDBB" : "5EC47A",
      fontFace: "Calibri"
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 11 — 月薪五萬的行動計劃
// ═══════════════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: LIGHT };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.75,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 }
  });
  s.addText("月薪五萬的翻身行動計劃", {
    x: 0.4, y: 0, w: 9, h: 0.75,
    fontSize: 22, bold: true, color: WHITE, fontFace: "Calibri", valign: "middle", margin: 0
  });

  const steps = [
    {
      n: "STEP 1",
      title: "提高儲蓄率",
      body: "找到舒適的生活水準，不強迫勒緊褲帶\n從 5~10% 開始儲蓄，逐步提高\n避免報復性消費",
      color: CORAL,
      detail: "目標：找到「我可以快樂活進月薪」的比例"
    },
    {
      n: "STEP 2",
      title: "建立緊急備用金",
      body: "目標：6～12 個月生活開銷\n放在銀行，不要投資\n達標前，不要急著進股市",
      color: GOLD,
      detail: "這是你最重要的護城河，沒有它就沒有安全感"
    },
    {
      n: "STEP 3",
      title: "自動化分流 → 開始投資",
      body: "設定薪資自動轉帳分流\n開銷帳戶 / 備用金帳戶 / 資產帳戶\n備用金存滿後，剩餘全進投資",
      color: "5EC47A",
      detail: "不依靠意志力，系統自動執行，越簡單越好"
    },
  ];

  steps.forEach((step, i) => {
    const y = 0.95 + i * 1.52;

    // Step number circle
    s.addShape(pres.shapes.OVAL, {
      x: 0.3, y: y + 0.15, w: 0.8, h: 0.8,
      fill: { color: step.color }, line: { color: step.color, width: 0 }
    });
    s.addText(String(i + 1), {
      x: 0.3, y: y + 0.15, w: 0.8, h: 0.8,
      fontSize: 26, bold: true, color: WHITE, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });

    // Connector line (not for last)
    if (i < 2) {
      s.addShape(pres.shapes.LINE, {
        x: 0.7, y: y + 1.0, w: 0, h: 0.56,
        line: { color: step.color, width: 2, dashType: "dash" }
      });
    }

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: 1.25, y, w: 8.4, h: 1.35,
      fill: { color: WHITE }, line: { color: MID, width: 1 },
      shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 1.25, y, w: 0.08, h: 1.35,
      fill: { color: step.color }, line: { color: step.color, width: 0 }
    });

    s.addText(step.n, {
      x: 1.42, y: y + 0.07, w: 1.0, h: 0.3,
      fontSize: 11, bold: true, color: step.color, fontFace: "Calibri"
    });
    s.addText(step.title, {
      x: 1.42, y: y + 0.32, w: 3.5, h: 0.45,
      fontSize: 18, bold: true, color: DARK, fontFace: "Calibri"
    });
    s.addText(step.body, {
      x: 5.1, y: y + 0.1, w: 4.4, h: 0.85,
      fontSize: 12, color: DARK, fontFace: "Calibri Light"
    });
    s.addShape(pres.shapes.LINE, {
      x: 1.42, y: y + 0.82, w: 7.98, h: 0,
      line: { color: MID, width: 0.5 }
    });
    s.addText(step.detail, {
      x: 1.42, y: y + 0.9, w: 7.9, h: 0.35,
      fontSize: 11.5, color: step.color, fontFace: "Calibri", italic: true
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 12 — 真正的自由 (dark, closing)
// ═══════════════════════════════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: NAVY };

  // Left gold accent bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 5.625,
    fill: { color: GOLD }, line: { color: GOLD, width: 0 }
  });

  s.addText("Free From Money", {
    x: 0.45, y: 0.4, w: 9.2, h: 0.65,
    fontSize: 14, color: GOLD, fontFace: "Calibri Light", charSpacing: 4
  });

  s.addText("真正的自由", {
    x: 0.45, y: 0.95, w: 9, h: 1.1,
    fontSize: 56, bold: true, color: WHITE, fontFace: "Calibri"
  });

  s.addShape(pres.shapes.LINE, {
    x: 0.45, y: 2.12, w: 4.5, h: 0,
    line: { color: CORAL, width: 1.5 }
  });

  const points = [
    "不是「有錢了才自由」，而是有錢沒錢都自在",
    "知道自己的「夠了」在哪裡，就是金錢的主人",
    "讓錢服務你的生活，而不是你服務於錢",
    "停下來，自己定義什麼是你的幸福",
    "找到退休後也想繼續做的事，收入就沒有上限",
  ];

  points.forEach((p, i) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.45, y: 2.28 + i * 0.56, w: 0.06, h: 0.3,
      fill: { color: CORAL }, line: { color: CORAL, width: 0 }
    });
    s.addText(p, {
      x: 0.65, y: 2.25 + i * 0.56, w: 8.0, h: 0.42,
      fontSize: 14, color: WHITE, fontFace: "Calibri Light", valign: "middle"
    });
  });

  // Bottom quote box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 5.1, w: 9.2, h: 0.38,
    fill: { color: "151E40" }, line: { color: "151E40", width: 0 }
  });
  s.addText("「最大的自由，是可以停下來，自己定義什麼叫自己的幸福。」— Will 黃士豪", {
    x: 0.55, y: 5.13, w: 9.0, h: 0.32,
    fontSize: 11.5, color: GOLD, fontFace: "Calibri Light", italic: true, valign: "middle"
  });
}

// Save
pres.writeFile({ fileName: "/Users/timwu/Documents/claude-skill-test/博音_翻身財務計劃.pptx" })
  .then(() => console.log("Done: 博音_翻身財務計劃.pptx"))
  .catch(e => { console.error(e); process.exit(1); });
