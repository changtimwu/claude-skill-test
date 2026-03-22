# 博音 × Will 黃士豪：從月薪翻身財務計劃

A presentation generated from the transcript of a 博音 (Boen) podcast episode featuring guest Will 黃士豪, covering practical personal finance strategies for turning your life around starting from zero savings.

## Contents

| File | Description |
|------|-------------|
| `博音_翻身財務計劃.pptx` | 12-slide presentation (Traditional Chinese) |
| `Voezvjo1xWw-transcript.txt` | Full transcript from YouTube (zh-TW) |
| `create_pptx.js` | Node.js script used to generate the PPTX |
| `youtube-transcript-fallback.patch` | Bug fix patch for the youtube-transcript skill |

## Presentation Slides

1. **Title** — 從月薪翻身：一個務實的財務計劃
2. **Will's Story** — 從近 3,000 萬負債到財富自由
3. **什麼叫「夠了」** — 金錢的主人 vs 金錢的奴隸
4. **財務目標的兩個維度** — 資產型 + 開銷型
5. **四桶金策略** — 保障、防守、進攻、樂透
6. **複利的威力** — 年化 8% vs 15%，30 年的差距
7. **打破本薪天花板** — 主動收入的兩種路徑
8. **槓桿的正確用法** — 兩大不斷頭原則
9. **買房 vs 租房** — 用正確框架做決定
10. **生存 vs 發展** — 大多數人的焦慮是多餘的
11. **月薪五萬行動計劃** — 三步驟立即執行
12. **真正的自由** — Free From Money

## Generating the PPTX

Requires [pptxgenjs](https://gitbun.com/gitbrent/PptxGenJS) installed globally:

```bash
npm install -g pptxgenjs
NODE_PATH=$(npm root -g) node create_pptx.js
```

## youtube-transcript Patch

`youtube-transcript-fallback.patch` fixes a bug in the [youtube-transcript skill](https://github.com/anthropics/skills) where the script fails if English captions are unavailable, instead of falling back to any available language.

Apply with:

```bash
patch -p1 < youtube-transcript-fallback.patch
```

## Source Video

YouTube: `https://youtu.be/Voezvjo1xWw`
Podcast: 博音 (Boen Podcast)
Guest: Will 黃士豪 (GoodWhale 創辦人)
