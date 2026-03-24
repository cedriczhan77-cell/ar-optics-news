/**
 * AR 眼镜 & 光学 日报 v8
 * · 标题加超链接（点击跳转来源）
 * · 无 emoji 混排，无 shrinkText，LibreOffice 安全
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");

const dataPath = process.argv[2] || "/tmp/slide_data.json";
const outPath  = process.argv[3] || "/tmp/daily_slide.pptx";
const raw = JSON.parse(fs.readFileSync(dataPath, "utf8"));

function clean(s) {
  if (!s) return "";
  return s
    .replace(/&#(\d+);/g,  (_,c)=>String.fromCharCode(+c))
    .replace(/&#x([0-9a-fA-F]+);/g,(_,h)=>String.fromCharCode(parseInt(h,16)))
    .replace(/&amp;/g,"&").replace(/&lt;/g,"<").replace(/&gt;/g,">")
    .replace(/&quot;/g,'"').replace(/&nbsp;/g," ").trim();
}

function trunc(text, maxW) {
  let w=0, r="";
  for (const ch of String(text||"")) {
    const cw = /[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]/.test(ch) ? 2 : 1;
    if (w+cw > maxW) { r+="…"; break; }
    w+=cw; r+=ch;
  }
  return r;
}

function cap(boxW, boxH, pt) {
  const CW = {20:0.160,12:0.098,11:0.085,10:0.076,9.5:0.072,9:0.068,8.5:0.065};
  const LH = {20:0.290,12:0.200,11:0.192,10:0.178,9.5:0.168,9:0.160,8.5:0.152};
  const cw = CW[pt]||0.076, lh = LH[pt]||0.178;
  return Math.floor(boxW/cw) * Math.max(1, Math.floor(boxH/lh));
}

const date        = clean(raw.date);
const summary     = clean(raw.summary);
const risk        = clean(raw.risk);
const opportunity = clean(raw.opportunity);
const papers = (raw.papers||[]).map(p=>({...p,title:clean(p.title),zh_desc:clean(p.zh_desc),source:clean(p.source),link:p.link||""}));
const news   = (raw.news||[]).map(n=>({...n,title:clean(n.title),  zh_desc:clean(n.zh_desc),source:clean(n.source),link:n.link||""}));

const C = {
  bg:"FFFFFF", bgGreen:"EDFAF3", white:"FFFFFF",
  title:"1A1D2E", body:"2D3748", sub:"718096",
  divider:"E2E8F0", border:"CBD5E0",
  blue:"1565D8", blueMid:"0F9FC8", green:"2BBD6E",
  greenDark:"1A9E57", orange:"F97316",
  link:"1565D8",   // 链接色（与主蓝一致）
};
const F = "Calibri";
const Sh = ()=>({type:"outer",color:"CBD5E0",blur:4,offset:2,angle:135,opacity:0.18});

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  const W=13.3, H=7.5;
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  // ── 左侧色条 ─────────────────────────────────────
  const STRIP=0.20;
  slide.addShape(pres.shapes.RECTANGLE,{x:0,y:0,      w:STRIP,h:H*0.50,fill:{color:C.blue},   line:{color:C.blue}   });
  slide.addShape(pres.shapes.RECTANGLE,{x:0,y:H*0.38, w:STRIP,h:H*0.24,fill:{color:C.blueMid},line:{color:C.blueMid}});
  slide.addShape(pres.shapes.RECTANGLE,{x:0,y:H*0.50, w:STRIP,h:H*0.50,fill:{color:C.green},  line:{color:C.green}  });

  // ── Header ───────────────────────────────────────
  const HDR=0.75;
  slide.addShape(pres.shapes.RECTANGLE,{x:STRIP,y:0,w:W-STRIP,h:HDR,fill:{color:C.white},line:{color:C.divider,pt:0.5}});
  slide.addText("AR Optics Daily",{x:W-2.7,y:0.08,w:2.55,h:0.28,margin:0,fontFace:F,fontSize:12,color:C.blue,bold:true,align:"right",valign:"middle"});
  slide.addText("AR 眼镜 & 光学  |  每日资讯简报",{x:STRIP+0.18,y:0.05,w:8.2,h:0.38,margin:0,fontFace:F,fontSize:20,color:C.title,bold:true,valign:"middle"});
  slide.addText(`${date}    #光学  #AR波导  #Waveguide  #TFLN  #MicroLED`,{x:STRIP+0.18,y:0.46,w:9,h:0.22,margin:0,fontFace:F,fontSize:9.5,color:C.sub,valign:"middle"});
  slide.addShape(pres.shapes.LINE,{x:STRIP,y:HDR,w:W-STRIP,h:0,line:{color:C.divider,width:1.5}});

  // ── 全局尺寸 ─────────────────────────────────────
  const TOP=HDR+0.10, BOTY=6.30;
  const MAIN_H=BOTY-TOP-0.06;
  const BOT_H=H-BOTY-0.06;
  const GAP=0.12;
  const LX=STRIP+0.12, LW=3.38;
  const MX=LX+LW+GAP,  MW=4.34;
  const RX=MX+MW+GAP,  RW=W-RX-0.12;

  // ══ 左栏：总结 ═══════════════════════════════════
  slide.addShape(pres.shapes.RECTANGLE,{x:LX,y:TOP,w:LW,h:MAIN_H,fill:{color:C.bgGreen},line:{color:C.greenDark,pt:1},shadow:Sh()});
  slide.addShape(pres.shapes.RECTANGLE,{x:LX,y:TOP,w:LW,h:0.36,fill:{color:C.green},line:{color:C.green}});
  slide.addText("本日总结",{x:LX+0.14,y:TOP,w:LW-0.20,h:0.36,margin:0,fontFace:F,fontSize:12,color:C.white,bold:true,valign:"middle"});

  const nP=papers.length, nN=news.length;
  const BW=(LW-0.28-0.08)/2, BY=TOP+0.42;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:LX+0.14,y:BY,w:BW,h:0.30,fill:{color:C.green},line:{color:C.green},rectRadius:0.05});
  slide.addText(`文献 ${nP} 篇`,{x:LX+0.14,y:BY,w:BW,h:0.30,margin:0,fontFace:F,fontSize:10.5,color:C.white,bold:true,align:"center",valign:"middle"});
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:LX+0.14+BW+0.08,y:BY,w:BW,h:0.30,fill:{color:C.blue},line:{color:C.blue},rectRadius:0.05});
  slide.addText(`新闻 ${nN} 篇`,{x:LX+0.14+BW+0.08,y:BY,w:BW,h:0.30,margin:0,fontFace:F,fontSize:10.5,color:C.white,bold:true,align:"center",valign:"middle"});

  const sumY=BY+0.36, sumH=MAIN_H-(sumY-TOP)-0.12, sumW=LW-0.28;
  slide.addText(trunc(summary||"今日活跃。",cap(sumW,sumH,11)),{
    x:LX+0.14,y:sumY,w:sumW,h:sumH,
    margin:0,fontFace:F,fontSize:11,color:C.body,
    align:"left",valign:"top",paraSpaceAfter:4,
  });

  // ══ 中/右栏公共渲染（标题带超链接）════════════════
  function renderCol(items, cx, cw, isLit) {
    const badgeColor = isLit ? C.green : C.blue;

    slide.addShape(pres.shapes.RECTANGLE,{x:cx,y:TOP,w:cw,h:MAIN_H,fill:{color:C.white},line:{color:C.border,pt:1},shadow:Sh()});
    slide.addShape(pres.shapes.RECTANGLE,{x:cx,y:TOP,w:cw,h:0.36,fill:{color:C.blue},line:{color:C.blue}});
    slide.addText(isLit?"学术文献":"科技新闻",{
      x:cx+0.14,y:TOP,w:cw-0.20,h:0.36,margin:0,fontFace:F,fontSize:12,color:C.white,bold:true,valign:"middle"
    });

    const N=3, DGAP=0.06, CPADB=0.08;
    const avail=MAIN_H-0.36-0.06-(N-1)*DGAP-CPADB;
    const IH=avail/N;
    const R_PAD=0.08, R_META=0.24;
    const R_TITLE=Math.min(IH*0.36, 0.56);
    const R_DESC=IH-R_PAD-R_META-R_TITLE-0.06;
    const colW=cw-0.28;
    const startY=TOP+0.36+0.06;

    items.slice(0,N).forEach((item,i)=>{
      const iy=startY+i*(IH+DGAP);
      if(i>0) slide.addShape(pres.shapes.LINE,{x:cx+0.14,y:iy-DGAP*0.5,w:cw-0.28,h:0,line:{color:C.divider,width:0.75}});

      const mY=iy+R_PAD, tY=mY+R_META+0.02, dY=tY+R_TITLE+0.02;

      // badge + 年份 + 来源
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:cx+0.14,y:mY+0.01,w:0.24,h:0.20,fill:{color:badgeColor},line:{color:badgeColor},rectRadius:0.03});
      slide.addText(`${i+1}`,{x:cx+0.14,y:mY+0.01,w:0.24,h:0.20,margin:0,fontFace:F,fontSize:9.5,color:C.white,bold:true,align:"center",valign:"middle"});
      slide.addText(`[${item.year}]`,{x:cx+0.42,y:mY+0.01,w:0.48,h:0.20,margin:0,fontFace:F,fontSize:9.5,color:badgeColor,bold:true,valign:"middle"});
      slide.addText(trunc(item.source||"",cap(cw-1.06,0.20,8.5)),{x:cx+0.92,y:mY+0.01,w:cw-1.06,h:0.20,margin:0,fontFace:F,fontSize:8.5,color:C.sub,align:"right",valign:"middle"});

      // ── 标题（带超链接）──────────────────────────
      const titleText = trunc(item.title||"", cap(colW, R_TITLE, 10));
      const hasLink   = item.link && item.link.startsWith("http");
      if (hasLink) {
        // Rich text 格式：文字 + 超链接
        slide.addText(
          [{ text: titleText, options: { hyperlink: { url: item.link } } }],
          {
            x:cx+0.14, y:tY, w:colW, h:R_TITLE,
            margin:0, fontFace:F, fontSize:10, color:C.link,
            bold:true, align:"left", valign:"top",
            underline: { style: "sng" },   // 下划线标示可点击
          }
        );
      } else {
        slide.addText(titleText,{
          x:cx+0.14,y:tY,w:colW,h:R_TITLE,
          margin:0,fontFace:F,fontSize:10,color:C.title,bold:true,align:"left",valign:"top",
        });
      }

      // ── 来源链接（小字，灰色）──────────────────────
      if (hasLink && R_DESC > 0.12) {
        const linkDisplay = trunc(item.link, cap(colW, 0.18, 8.5));
        // 中文摘要（优先）或来源链接
        const descText = trunc(item.zh_desc||"", cap(colW, R_DESC-0.20, 9.5));
        if (descText) {
          slide.addText(descText,{
            x:cx+0.14,y:dY,w:colW,h:R_DESC-0.20,
            margin:0,fontFace:F,fontSize:9.5,color:C.sub,align:"left",valign:"top",
          });
        }
        // 链接一行单独放（最底部，可点击）
        slide.addText(
          [{ text: linkDisplay, options:{ hyperlink:{ url: item.link } } }],
          {
            x:cx+0.14, y:dY+R_DESC-0.20, w:colW, h:0.18,
            margin:0, fontFace:F, fontSize:8, color:C.link,
            align:"left", valign:"middle",
            underline:{ style:"sng" },
          }
        );
      } else if (R_DESC > 0.12) {
        slide.addText(trunc(item.zh_desc||"",cap(colW,R_DESC,9.5)),{
          x:cx+0.14,y:dY,w:colW,h:R_DESC,
          margin:0,fontFace:F,fontSize:9.5,color:C.sub,align:"left",valign:"top",
        });
      }
    });

    for(let i=items.length;i<N;i++){
      const iy=startY+i*(IH+DGAP);
      if(i>0) slide.addShape(pres.shapes.LINE,{x:cx+0.14,y:iy-DGAP*0.5,w:cw-0.28,h:0,line:{color:C.divider,width:0.75}});
      slide.addText("今日暂无内容",{x:cx+0.14,y:iy+IH*0.40,w:cw-0.28,h:0.28,margin:0,fontFace:F,fontSize:10.5,color:C.sub,align:"center"});
    }
  }

  renderCol(papers, MX, MW, true);
  renderCol(news,   RX, RW, false);

  // ══ 底部：风险 & 机会 ════════════════════════════
  const HALF=(W-LX-0.12-GAP)/2;
  const bTxtW=HALF-0.28, bTxtH=BOT_H-0.36;

  slide.addShape(pres.shapes.RECTANGLE,{x:LX,y:BOTY,w:HALF,h:BOT_H,fill:{color:"FFF8F0"},line:{color:"FDBA74",pt:1.5},shadow:Sh()});
  slide.addShape(pres.shapes.RECTANGLE,{x:LX,y:BOTY,w:0.12,h:BOT_H,fill:{color:C.orange},line:{color:C.orange}});
  slide.addText("风险",{x:LX+0.20,y:BOTY+0.04,w:1.2,h:0.24,margin:0,fontFace:F,fontSize:11,color:C.orange,bold:true,valign:"middle"});
  slide.addText(trunc(risk||"暂无",cap(bTxtW,bTxtH,10)),{x:LX+0.20,y:BOTY+0.30,w:bTxtW,h:bTxtH,margin:0,fontFace:F,fontSize:10,color:C.body,valign:"top"});

  slide.addShape(pres.shapes.RECTANGLE,{x:LX+HALF+GAP,y:BOTY,w:HALF,h:BOT_H,fill:{color:C.bgGreen},line:{color:C.greenDark,pt:1.5},shadow:Sh()});
  slide.addShape(pres.shapes.RECTANGLE,{x:LX+HALF+GAP,y:BOTY,w:0.12,h:BOT_H,fill:{color:C.green},line:{color:C.green}});
  slide.addText("机会",{x:LX+HALF+GAP+0.20,y:BOTY+0.04,w:1.2,h:0.24,margin:0,fontFace:F,fontSize:11,color:C.greenDark,bold:true,valign:"middle"});
  slide.addText(trunc(opportunity||"暂无",cap(bTxtW,bTxtH,10)),{x:LX+HALF+GAP+0.20,y:BOTY+0.30,w:bTxtW,h:bTxtH,margin:0,fontFace:F,fontSize:10,color:C.body,valign:"top"});

  // ── 页脚 ─────────────────────────────────────────
  slide.addShape(pres.shapes.LINE,{x:STRIP,y:H-0.18,w:W-STRIP,h:0,line:{color:C.divider,width:0.8}});
  slide.addText("AR Optics Daily  ·  Powered by Claude AI  ·  数据来源：Nature / arXiv / IEEE / 36氪 / TechCrunch",{
    x:STRIP+0.1,y:H-0.17,w:W-STRIP-0.2,h:0.16,margin:0,fontFace:F,fontSize:7.5,color:C.sub,align:"center",
  });

  await pres.writeFile({ fileName: outPath });
  console.log("PPTX_DONE:" + outPath);
}

build().catch(e=>{console.error(e);process.exit(1);});
