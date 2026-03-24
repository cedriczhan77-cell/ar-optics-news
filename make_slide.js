/**
 * AR 眼镜 & 光学 日报 — Goerpixel 投研风格
 * 白色背景 · 蓝绿渐变左侧色条 · 浅绿高亮 · 深色标题
 */

const pptxgen = require("pptxgenjs");
const fs      = require("fs");

const dataPath = process.argv[2] || "/tmp/slide_data.json";
const outPath  = process.argv[3] || "/tmp/daily_slide.pptx";
const data     = JSON.parse(fs.readFileSync(dataPath, "utf8"));

const { date, summary, risk, opportunity, papers, news } = data;

// ════════════════════════════════════════════════════
//  Goerpixel 配色（从图片精确提取）
// ════════════════════════════════════════════════════
const C = {
  white:      "FFFFFF",
  bg:         "FFFFFF",       // 主背景：纯白
  bgLight:    "F7F9FC",       // 卡片背景：极浅蓝灰
  bgGreen:    "EDFAF3",       // 浅绿高亮背景（总结区域）
  bgBlue:     "EBF3FF",       // 浅蓝高亮背景（文献标题栏）
  title:      "1A1D2E",       // 深蓝黑标题文字
  body:       "2D3748",       // 深灰正文
  sub:        "718096",       // 浅灰辅助文字
  divider:    "E2E8F0",       // 分隔线
  border:     "D1DCE8",       // 卡片边框
  // 品牌色
  blue:       "1565D8",       // 主蓝（左侧色条顶部）
  blueMid:    "0F9FC8",       // 中间过渡蓝绿
  green:      "2BBD6E",       // 主绿（左侧色条底部 & 标签）
  greenDark:  "1A9E57",       // 深绿（边框）
  greenLight: "39D68A",       // 亮绿（文献编号标签）
  orange:     "F97316",       // 橙（风险）
  yellow:     "F59E0B",       // 黄（来源标签）
  // 新闻蓝
  newsBlue:   "1565D8",
  newsBlueBg: "EBF3FF",
};

const F = { heading: "Calibri", body: "Calibri" };
const mkShadow = () => ({ type:"outer", color:"CBD5E0", blur:6, offset:2, angle:135, opacity:0.35 });
const mkBorder = (color) => ({ color: color || C.border, pt: 1 });

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.3" × 7.5"
  const W = 13.3, H = 7.5;

  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  // ════════════════════════════════════════════════
  //  左侧渐变色条（蓝→绿，模拟 Goerpixel 斜切 ribbon）
  //  实现：用两个矩形叠加模拟渐变
  // ════════════════════════════════════════════════
  const STRIP = 0.22; // 色条宽度
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: STRIP, h: H * 0.5,
    fill: { color: C.blue }, line: { color: C.blue },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H * 0.5, w: STRIP, h: H * 0.5,
    fill: { color: C.green }, line: { color: C.green },
  });
  // 中间渐变过渡块
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H * 0.38, w: STRIP, h: H * 0.24,
    fill: { color: C.blueMid }, line: { color: C.blueMid },
  });

  // ════════════════════════════════════════════════
  //  顶部 Header
  // ════════════════════════════════════════════════
  const HDR_H = 0.80;

  // 标题区浅背景
  slide.addShape(pres.shapes.RECTANGLE, {
    x: STRIP, y: 0, w: W - STRIP, h: HDR_H,
    fill: { color: C.white }, line: { color: C.divider, pt: 0.5 },
  });

  // 品牌名 (右上角)
  slide.addText("AR Optics Daily", {
    x: W - 2.8, y: 0.05, w: 2.6, h: 0.36, margin: 0,
    fontFace: F.heading, fontSize: 13, color: C.blue, bold: true,
    align: "right", valign: "middle",
  });

  // 主标题
  slide.addText("AR 眼镜 & 光学  |  每日资讯简报", {
    x: STRIP + 0.2, y: 0.04, w: 8.5, h: 0.44, margin: 0,
    fontFace: F.heading, fontSize: 22, color: C.title, bold: true, valign: "middle",
  });

  // 日期 + 关键词 tag 行
  const tagStyle = { fontFace: F.body, fontSize: 10, color: C.sub, valign: "middle" };
  slide.addText(`${date}    #光学  #AR波导  #Waveguide  #TFLN  #MicroLED`, {
    x: STRIP + 0.2, y: 0.50, w: 9, h: 0.26, margin: 0,
    ...tagStyle,
  });

  // 分隔线（header 底部）
  slide.addShape(pres.shapes.LINE, {
    x: STRIP, y: HDR_H, w: W - STRIP, h: 0,
    line: { color: C.divider, width: 1.5 },
  });

  // ════════════════════════════════════════════════
  //  三栏布局
  //  左：总结  |  中：文献  |  右：新闻
  //  宽度：3.5  |  4.3  |  4.8（不含色条 0.22）
  // ════════════════════════════════════════════════
  const TOP    = HDR_H + 0.12;
  const BOT_H  = 1.02;
  const MAIN_H = H - TOP - BOT_H - 0.10;
  const GAP    = 0.14;

  const LX = STRIP + 0.15,  LW = 3.42;
  const MX = LX + LW + GAP, MW = 4.28;
  const RX = MX + MW + GAP, RW = W - RX - 0.15;

  // ── 左栏：总结 ─────────────────────────────────
  // 浅绿背景卡
  slide.addShape(pres.shapes.RECTANGLE, {
    x: LX, y: TOP, w: LW, h: MAIN_H,
    fill: { color: C.bgGreen }, line: mkBorder(C.greenDark),
    shadow: mkShadow(),
  });

  // 标题行（绿色横条）
  slide.addShape(pres.shapes.RECTANGLE, {
    x: LX, y: TOP, w: LW, h: 0.40,
    fill: { color: C.green }, line: { color: C.green },
  });
  slide.addText("🔍  本日总结", {
    x: LX + 0.15, y: TOP, w: LW - 0.3, h: 0.40, margin: 0,
    fontFace: F.heading, fontSize: 12.5, color: C.white, bold: true, valign: "middle",
  });

  // 统计小 badge
  const nPaper = (papers||[]).length, nNews = (news||[]).length;
  const badgeY = TOP + 0.48;

  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: LX + 0.16, y: badgeY, w: 1.42, h: 0.34,
    fill: { color: C.green }, line: { color: C.green }, rectRadius: 0.06,
  });
  slide.addText(`📚 文献  ${nPaper} 篇`, {
    x: LX + 0.16, y: badgeY, w: 1.42, h: 0.34, margin: 0,
    fontFace: F.body, fontSize: 11, color: C.white, bold: true,
    align: "center", valign: "middle",
  });

  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: LX + 1.76, y: badgeY, w: 1.42, h: 0.34,
    fill: { color: C.blue }, line: { color: C.blue }, rectRadius: 0.06,
  });
  slide.addText(`📰 新闻  ${nNews} 篇`, {
    x: LX + 1.76, y: badgeY, w: 1.42, h: 0.34, margin: 0,
    fontFace: F.body, fontSize: 11, color: C.white, bold: true,
    align: "center", valign: "middle",
  });

  // 总结正文
  slide.addText(summary || "今日 AR 光学领域持续活跃。", {
    x: LX + 0.16, y: badgeY + 0.44, w: LW - 0.32, h: MAIN_H - 0.40 - 0.44 - 0.20,
    fontFace: F.body, fontSize: 11.5, color: C.body,
    align: "left", valign: "top", paraSpaceAfter: 5,
  });

  // ── 中栏：文献 ─────────────────────────────────
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MX, y: TOP, w: MW, h: MAIN_H,
    fill: { color: C.white }, line: mkBorder(C.border),
    shadow: mkShadow(),
  });
  // 标题条（浅蓝）
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MX, y: TOP, w: MW, h: 0.40,
    fill: { color: C.blue }, line: { color: C.blue },
  });
  slide.addText("📚  学术文献", {
    x: MX + 0.15, y: TOP, w: MW - 0.3, h: 0.40, margin: 0,
    fontFace: F.heading, fontSize: 12.5, color: C.white, bold: true, valign: "middle",
  });

  const paperList = (papers||[]).slice(0,3);
  const P_IH     = (MAIN_H - 0.46) / 3;

  paperList.forEach((p, i) => {
    const iy = TOP + 0.44 + i * P_IH;
    if (i > 0) {
      slide.addShape(pres.shapes.LINE, {
        x: MX + 0.18, y: iy - 0.04, w: MW - 0.36, h: 0,
        line: { color: C.divider, width: 1 },
      });
    }
    // 编号 badge（绿色）
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: MX + 0.16, y: iy + 0.06, w: 0.28, h: 0.24,
      fill: { color: C.green }, line: { color: C.green }, rectRadius: 0.04,
    });
    slide.addText(`${i+1}`, {
      x: MX + 0.16, y: iy + 0.06, w: 0.28, h: 0.24, margin: 0,
      fontFace: F.body, fontSize: 10, color: C.white, bold: true,
      align: "center", valign: "middle",
    });
    // 年份
    slide.addText(`[${p.year}]`, {
      x: MX + 0.50, y: iy + 0.06, w: 0.55, h: 0.24, margin: 0,
      fontFace: F.body, fontSize: 10, color: C.green, bold: true, valign: "middle",
    });
    // 来源（右对齐）
    slide.addText(p.source || "", {
      x: MX + 1.08, y: iy + 0.06, w: MW - 1.28, h: 0.24, margin: 0,
      fontFace: F.body, fontSize: 9.5, color: C.sub, align: "right", valign: "middle",
    });
    // 英文标题
    const t = (p.title||"").length > 72 ? p.title.slice(0,70)+"…" : p.title;
    slide.addText(t, {
      x: MX + 0.16, y: iy + 0.33, w: MW - 0.32, h: 0.42,
      fontFace: F.body, fontSize: 10.5, color: C.title, bold: true,
      align: "left", valign: "top",
    });
    // 中文摘要
    const d = (p.zh_desc||"").slice(0,65) + ((p.zh_desc||"").length>65?"…":"");
    slide.addText(d, {
      x: MX + 0.16, y: iy + 0.77, w: MW - 0.32, h: P_IH - 0.85,
      fontFace: F.body, fontSize: 10, color: C.sub, align: "left", valign: "top",
    });
  });

  // 文献不足占位
  for (let i = paperList.length; i < 3; i++) {
    const iy = TOP + 0.44 + i * P_IH;
    if (i > 0) slide.addShape(pres.shapes.LINE, {
      x: MX + 0.18, y: iy - 0.04, w: MW - 0.36, h: 0,
      line: { color: C.divider, width: 1 },
    });
    slide.addText("今日暂无新文献", {
      x: MX + 0.18, y: iy + P_IH/2 - 0.15, w: MW - 0.36, h: 0.3,
      fontFace: F.body, fontSize: 11, color: C.sub, align: "center",
    });
  }

  // ── 右栏：新闻 ─────────────────────────────────
  slide.addShape(pres.shapes.RECTANGLE, {
    x: RX, y: TOP, w: RW, h: MAIN_H,
    fill: { color: C.white }, line: mkBorder(C.border),
    shadow: mkShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: RX, y: TOP, w: RW, h: 0.40,
    fill: { color: C.newsBlue }, line: { color: C.newsBlue },
  });
  slide.addText("📰  科技新闻", {
    x: RX + 0.15, y: TOP, w: RW - 0.3, h: 0.40, margin: 0,
    fontFace: F.heading, fontSize: 12.5, color: C.white, bold: true, valign: "middle",
  });

  const newsList = (news||[]).slice(0,3);
  const N_IH    = (MAIN_H - 0.46) / 3;

  newsList.forEach((n, i) => {
    const iy = TOP + 0.44 + i * N_IH;
    if (i > 0) {
      slide.addShape(pres.shapes.LINE, {
        x: RX + 0.18, y: iy - 0.04, w: RW - 0.36, h: 0,
        line: { color: C.divider, width: 1 },
      });
    }
    // 编号 badge（蓝色）
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: RX + 0.16, y: iy + 0.06, w: 0.28, h: 0.24,
      fill: { color: C.newsBlue }, line: { color: C.newsBlue }, rectRadius: 0.04,
    });
    slide.addText(`${i+1}`, {
      x: RX + 0.16, y: iy + 0.06, w: 0.28, h: 0.24, margin: 0,
      fontFace: F.body, fontSize: 10, color: C.white, bold: true,
      align: "center", valign: "middle",
    });
    slide.addText(`[${n.year}]`, {
      x: RX + 0.50, y: iy + 0.06, w: 0.55, h: 0.24, margin: 0,
      fontFace: F.body, fontSize: 10, color: C.newsBlue, bold: true, valign: "middle",
    });
    slide.addText(n.source || "", {
      x: RX + 1.08, y: iy + 0.06, w: RW - 1.28, h: 0.24, margin: 0,
      fontFace: F.body, fontSize: 9.5, color: C.sub, align: "right", valign: "middle",
    });
    const t = (n.title||"").length > 72 ? n.title.slice(0,70)+"…" : n.title;
    slide.addText(t, {
      x: RX + 0.16, y: iy + 0.33, w: RW - 0.32, h: 0.42,
      fontFace: F.body, fontSize: 10.5, color: C.title, bold: true,
      align: "left", valign: "top",
    });
    const d = (n.zh_desc||"").slice(0,65) + ((n.zh_desc||"").length>65?"…":"");
    slide.addText(d, {
      x: RX + 0.16, y: iy + 0.77, w: RW - 0.32, h: N_IH - 0.85,
      fontFace: F.body, fontSize: 10, color: C.sub, align: "left", valign: "top",
    });
  });

  for (let i = newsList.length; i < 3; i++) {
    const iy = TOP + 0.44 + i * N_IH;
    if (i > 0) slide.addShape(pres.shapes.LINE, {
      x: RX + 0.18, y: iy - 0.04, w: RW - 0.36, h: 0,
      line: { color: C.divider, width: 1 },
    });
    slide.addText("今日暂无新闻", {
      x: RX + 0.18, y: iy + N_IH/2 - 0.15, w: RW - 0.36, h: 0.3,
      fontFace: F.body, fontSize: 11, color: C.sub, align: "center",
    });
  }

  // ════════════════════════════════════════════════
  //  底部：风险 & 机会（Goerpixel 浅绿横条）
  // ════════════════════════════════════════════════
  const BOT_Y = TOP + MAIN_H + 0.08;
  const HALF  = (W - STRIP - 0.30 - GAP) / 2;

  // 风险卡（浅橙边框）
  slide.addShape(pres.shapes.RECTANGLE, {
    x: LX, y: BOT_Y, w: HALF, h: BOT_H - 0.08,
    fill: { color: "FFF8F0" }, line: { color: "FDBA74", pt: 1.5 },
    shadow: mkShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: LX, y: BOT_Y, w: 0.16, h: BOT_H - 0.08,
    fill: { color: C.orange }, line: { color: C.orange },
  });
  slide.addText("⚠️  风险", {
    x: LX + 0.24, y: BOT_Y + 0.04, w: 1.2, h: 0.28, margin: 0,
    fontFace: F.heading, fontSize: 11.5, color: C.orange, bold: true, valign: "middle",
  });
  slide.addText(risk || "暂无风险提示", {
    x: LX + 0.24, y: BOT_Y + 0.35, w: HALF - 0.34, h: BOT_H - 0.50,
    fontFace: F.body, fontSize: 10.5, color: C.body, valign: "top",
  });

  // 机会卡（浅绿边框）
  slide.addShape(pres.shapes.RECTANGLE, {
    x: LX + HALF + GAP, y: BOT_Y, w: HALF, h: BOT_H - 0.08,
    fill: { color: C.bgGreen }, line: { color: C.greenDark, pt: 1.5 },
    shadow: mkShadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: LX + HALF + GAP, y: BOT_Y, w: 0.16, h: BOT_H - 0.08,
    fill: { color: C.green }, line: { color: C.green },
  });
  slide.addText("💡  机会", {
    x: LX + HALF + GAP + 0.24, y: BOT_Y + 0.04, w: 1.2, h: 0.28, margin: 0,
    fontFace: F.heading, fontSize: 11.5, color: C.greenDark, bold: true, valign: "middle",
  });
  slide.addText(opportunity || "暂无机会提示", {
    x: LX + HALF + GAP + 0.24, y: BOT_Y + 0.35, w: HALF - 0.34, h: BOT_H - 0.50,
    fontFace: F.body, fontSize: 10.5, color: C.body, valign: "top",
  });

  // 页脚
  slide.addShape(pres.shapes.LINE, {
    x: STRIP, y: H - 0.20, w: W - STRIP, h: 0,
    line: { color: C.divider, width: 0.8 },
  });
  slide.addText("AR Optics Daily  ·  Powered by Claude AI  ·  数据来源：Nature / arXiv / IEEE / 36氪 / TechCrunch", {
    x: STRIP + 0.1, y: H - 0.19, w: W - STRIP - 0.2, h: 0.18, margin: 0,
    fontFace: F.body, fontSize: 7.5, color: C.sub, align: "center",
  });

  await pres.writeFile({ fileName: outPath });
  console.log("PPTX_DONE:" + outPath);
}

build().catch(e => { console.error(e); process.exit(1); });
