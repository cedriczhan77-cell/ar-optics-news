"""
AR 光学日报 — 图片生成 & 飞书推送
══════════════════════════════════════════════════════════════
完全复用 news_pusher.py 的所有逻辑：
  · 相同的 RSS 源（26 个）
  · 相同的关键字列表（KEYWORDS_MUST / KEYWORDS_NEWS_STRICT）
  · 相同的 AR 优先分值系统（AR_PRIORITY_KEYWORDS）
  · 相同的去重历史（sent_history.json）
  · 相同的 AI 摘要 / 总结 / 风险机会 Prompt
  · 相同的配额（新闻 3 篇 / 文献 0~3 篇）

额外流程（仅此文件新增）：
  ① 生成 slide_data.json
  ② Node.js make_slide.js  →  daily_slide.pptx
  ③ LibreOffice            →  daily_slide.pdf
  ④ pdftoppm               →  daily_slide-1.jpg（200 DPI）
  ⑤ 推送图片 + 文字卡片 到所有飞书 Webhook
══════════════════════════════════════════════════════════════
"""

import feedparser, hashlib, json, os, re, subprocess, base64, sys
from datetime import datetime, timezone, timedelta
from zoneinfo  import ZoneInfo
import requests

# ════════════════════════════════════════════════════════
#  ① 配置（与 news_pusher.py 完全一致）
# ════════════════════════════════════════════════════════

FEISHU_WEBHOOK    = os.environ.get("FEISHU_WEBHOOK",    "https://open.feishu.cn/open-apis/bot/v2/hook/YOUR_HOOK_ID")
FEISHU_WEBHOOK_2  = os.environ.get("FEISHU_WEBHOOK_2",  "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

QUOTA_NEWS        = 3
QUOTA_PAPER       = 3
WINDOW_NEWS       = 48
WINDOW_PAPER      = 7 * 24
HISTORY_FILE      = "slide_history.json"   # 独立历史，不与 news_pusher.py 冲突
HISTORY_KEEP_DAYS = 60

ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
AI_MODEL      = "claude-sonnet-4-20250514"   # 与 news_pusher.py 一致

# 路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
JS_SCRIPT  = os.path.join(SCRIPT_DIR, "make_slide.js")
TMP_JSON   = "/tmp/slide_data.json"
TMP_PPTX   = "/tmp/daily_slide.pptx"
TMP_PDF    = "/tmp/daily_slide.pdf"
TMP_IMG    = "/tmp/daily_slide-1.jpg"
SOFFICE_PY = os.path.join(SCRIPT_DIR, "soffice_helper.py")  # 见下方说明

# ════════════════════════════════════════════════════════
#  ② 关键字（与 news_pusher.py 完全一致）
# ════════════════════════════════════════════════════════

KEYWORDS_MUST = [
    "optical", "optics", "photonics",
    "augmented reality",
    "waveguide", "光波导", "波导",
    "eyepiece", "目镜",
    "TFLN", "thin-film lithium niobate", "lithium niobate",
    "LiNbO3", "LN modulator", "electro-optic modulator",
    "铌酸锂", "薄膜铌酸锂", "电光调制",
    "diffractive", "holographic", "hologram", "HOE",
    "metasurface", "metalens", "超表面", "元透镜",
    "surface relief grating", "SRG",
    "grating coupler", "diffraction grating", "光栅",
    "pancake lens", "birdbath",
    "light engine", "micro display", "near-eye display",
    "microLED", "micro-LED", "LCoS", "DLP",
    "mixed reality", "smart glasses", "head-mounted display",
    "HoloLens", "Magic Leap", "Apple Vision Pro",
    "Meta Quest", "Ray-Ban Meta", "XREAL", "Rokid",
    "增强现实", "混合现实", "衍射", "全息",
    "智能眼镜", "头显", "近眼显示",
]

KEYWORDS_NEWS_STRICT = [
    "augmented reality", "AR glasses", "AR headset", "AR display",
    "smart glasses", "mixed reality glasses",
    "Apple Vision", "Meta Quest", "HoloLens", "Magic Leap",
    "Ray-Ban Meta", "XREAL", "Rokid", "TCL RayNeo",
    "增强现实眼镜", "智能眼镜", "头显", "近眼显示",
    "waveguide", "optical waveguide", "diffractive waveguide",
    "holographic waveguide", "光波导", "波导显示",
    "eyepiece", "目镜", "光学模组",
    "metasurface lens", "metalens", "超表面透镜",
    "pancake lens", "pancake optics", "birdbath optics",
    "diffractive optical", "holographic optical",
    "surface relief grating", "SRG waveguide",
    "micro display", "microLED display", "LCoS display",
    "light engine", "near-eye", "microdisplay",
    "retinal display", "retinal projection",
    "TFLN", "thin-film lithium niobate", "lithium niobate modulator",
    "electro-optic modulator", "photonic chip", "铌酸锂",
    "AR industry", "AR market", "AR optics",
    "wearable display", "wearable optics",
    "optical see-through", "see-through display",
    "AR NED", "NED display",
    "OpenAI", "ChatGPT", "GPT-5", "GPT-4",
    "Claude", "Anthropic",
    "Gemini", "Google DeepMind",
    "Grok", "xAI",
    "DeepSeek", "Mistral", "Llama",
    "large language model", "LLM",
    "foundation model", "multimodal model",
    "AI model", "AI agent", "AI assistant",
    "artificial general intelligence", "AGI",
    "大语言模型", "基础模型", "多模态",
    "AI 进展", "AI 突破", "人工智能",
    "NVIDIA", "GPU cluster", "AI chip",
    "AI accelerator", "TPU", "AI inference",
    "算力", "AI芯片", "英伟达",
    "AI-powered", "generative AI", "text-to-image",
    "text-to-video", "AI video", "Sora", "Runway",
    "生成式AI", "AI生成", "文生图", "文生视频",
]

KEYWORDS_EXTRA = [
    "eye tracking", "眼动追踪",
    "FOV", "field of view", "视场角",
    "aberration", "像差",
    "stray light", "杂散光",
    "diffraction efficiency",
    "ray tracing", "光线追迹",
    "tolerance", "manufacturing",
    "simulation",
]

AR_PRIORITY_KEYWORDS: list[tuple[str, int]] = [
    ("SRG",                    10), ("surface relief grating", 10), ("表面浮雕光栅",  10),
    ("smartcut",               10), ("smartcut bonding",       10), ("tflnonsi",       10),
    ("d2w reconstitution",     10), ("reconstitution wafer",   10),
    ("SiC microchannel",        8), ("SiC 微通道",               8), ("散热基板",        8), ("imec", 7),
    ("SRG WG",                 10), ("see-through artifact",   10), ("grating conspicuity", 10),
    ("2D grating",             10), ("large fov",               9), ("large field of view", 9),
    ("doe",                     8), ("moe",                     8), ("可见光 metalens",  9),
    ("visible metalens",        9), ("achromatic",              8), ("消色差",           8),
    ("fov expansion",           9), ("视场扩展",                 9), ("矢量仿真",          8),
    ("metasurface based SRG",  12), ("超表面 SRG",              12), ("alternative AR NED", 12),
    ("Maxwellian",             11), ("maxwellian display",     11), ("maxwellian view",  11),
    ("带景深",                  11), ("depth of field",         10), ("creal",            10),
    ("swave",                  10), ("vividQ",                 10), ("non-waveguide",    10),
    ("non WG",                 10), ("retinal projection",     10), ("视网膜投影",        10),
    ("lightfield",              9), ("light field display",    9),  ("光场显示",           9),
    ("waveguide combiner",      9), ("waveguide display",      9),  ("AR waveguide",      9),
    ("AR glasses",              9), ("AR headset",             9),  ("near-eye display",  9),
    ("NED",                     9), ("near eye",               8),  ("光波导显示",         9),
    ("近眼显示",                 9), ("眼镜光学",                9),  ("combiner",          8),
    ("耦合器",                   8), ("in-coupler",             8),  ("out-coupler",        8),
    ("exit pupil",              8), ("出瞳",                   8),  ("pupil expansion",   8), ("瞳孔扩展", 8),
]

# ════════════════════════════════════════════════════════
#  ③ RSS 源（与 news_pusher.py 完全一致，26 个）
# ════════════════════════════════════════════════════════

RSS_SOURCES = [
    # ── 学术 ──────────────────────────────────────────
    {"name": "Nature",              "url": "https://www.nature.com/nature.rss",                                    "type": "paper"},
    {"name": "Nature Photonics",    "url": "https://www.nature.com/nphoton.rss",                                   "type": "paper"},
    {"name": "Nature Electronics",  "url": "https://www.nature.com/natelectron.rss",                               "type": "paper"},
    {"name": "Light: Sci & App",    "url": "https://www.nature.com/lsa.rss",                                       "type": "paper"},
    {"name": "Science",             "url": "https://www.science.org/rss/news_current.xml",                         "type": "paper"},
    {"name": "arXiv · Optics",      "url": "https://arxiv.org/rss/physics.optics",                                "type": "paper"},
    {"name": "arXiv · App Physics", "url": "https://arxiv.org/rss/physics.app-ph",                                "type": "paper"},
    {"name": "arXiv · eess.IV",     "url": "https://arxiv.org/rss/eess.IV",                                       "type": "paper"},
    {"name": "Optica",              "url": "https://opg.optica.org/rss/optica.xml",                                "type": "paper"},
    {"name": "Optics Express",      "url": "https://opg.optica.org/rss/oe.xml",                                    "type": "paper"},
    {"name": "Optics Letters",      "url": "https://opg.optica.org/rss/ol.xml",                                    "type": "paper"},
    {"name": "SPIE · Opt Eng",      "url": "https://www.spiedigitallibrary.org/rss/oe.xml",                       "type": "paper"},
    {"name": "IEEE Photonics J.",   "url": "https://ieeexplore.ieee.org/rss/TOC69.XML",                           "type": "paper"},
    {"name": "IEEE Trans. Display", "url": "https://ieeexplore.ieee.org/rss/TOC8240.XML",                         "type": "paper"},
    {"name": "ACS Photonics",       "url": "https://pubs.acs.org/action/showFeed?type=axatoc&feed=rss&jc=apchd5", "type": "paper"},
    {"name": "Phys Rev Applied",    "url": "https://feeds.aps.org/rss/recent/prapplied.rss",                      "type": "paper"},
    # ── 行业（国际）──────────────────────────────────
    {"name": "MIT Tech Review",     "url": "https://www.technologyreview.com/feed/",       "type": "news"},
    {"name": "The Verge",           "url": "https://www.theverge.com/rss/index.xml",        "type": "news"},
    {"name": "Wired",               "url": "https://www.wired.com/feed/rss",                "type": "news"},
    {"name": "TechCrunch",          "url": "https://techcrunch.com/feed/",                  "type": "news"},
    {"name": "DisplayDaily",        "url": "https://www.displaydaily.com/feed",             "type": "news"},
    {"name": "DIGITIMES",           "url": "https://www.digitimes.com/rss/news.xml",        "type": "news"},
    # ── 国内 ─────────────────────────────────────────
    {"name": "36氪",                "url": "https://36kr.com/feed",                         "type": "news"},
    {"name": "虎嗅",                "url": "https://www.huxiu.com/rss/0.xml",               "type": "news"},
    {"name": "机器之心",             "url": "https://www.jiqizhixin.com/rss",                "type": "news"},
    {"name": "量子位",               "url": "https://www.qbitai.com/feed",                   "type": "news"},
    {"name": "极客公园",             "url": "https://www.geekpark.net/rss",                  "type": "news"},
]

# ════════════════════════════════════════════════════════
#  去重历史（与 news_pusher.py 完全一致）
# ════════════════════════════════════════════════════════

def url_id(url: str) -> str:
    return hashlib.sha1(url.encode()).hexdigest()[:16]

def load_history() -> dict:
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_history(history: dict) -> None:
    cutoff  = (datetime.now(timezone.utc) - timedelta(days=HISTORY_KEEP_DAYS)).strftime("%Y-%m-%d")
    cleaned = {k: v for k, v in history.items() if v >= cutoff}
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(cleaned, f, ensure_ascii=False, indent=2)
    print(f"  💾 历史记录 {len(cleaned)} 条")

# ════════════════════════════════════════════════════════
#  RSS 抓取（与 news_pusher.py 完全一致）
# ════════════════════════════════════════════════════════

def strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", "", text or "").strip()

def keyword_match(text: str, src_type: str) -> tuple[bool, list[str], int]:
    lower   = text.lower()
    kw_list = KEYWORDS_MUST if src_type == "paper" else KEYWORDS_NEWS_STRICT
    hit     = any(k.lower() in lower for k in kw_list)
    extra   = [k for k in KEYWORDS_EXTRA if k.lower() in lower]
    priority= sum(score for kw, score in AR_PRIORITY_KEYWORDS if kw.lower() in lower)
    return hit, extra, priority

def parse_pub(entry) -> datetime | None:
    for attr in ("published_parsed", "updated_parsed"):
        t = getattr(entry, attr, None)
        if t:
            try: return datetime(*t[:6], tzinfo=timezone.utc)
            except: pass
    return None

def fetch_candidates(src_type: str, history: dict) -> list[dict]:
    window  = WINDOW_PAPER if src_type == "paper" else WINDOW_NEWS
    cutoff  = datetime.now(timezone.utc) - timedelta(hours=window)
    sources = [s for s in RSS_SOURCES if s["type"] == src_type]
    candidates: list[dict] = []

    for src in sources:
        try:
            feed = feedparser.parse(src["url"])
        except Exception as e:
            print(f"  [ERR] {src['name']}: {e}"); continue

        for entry in feed.entries[:60]:
            link = entry.get("link", "").strip()
            if not link: continue
            uid = url_id(link)
            if uid in history: continue
            pub = parse_pub(entry)
            if pub and pub < cutoff: continue

            title   = strip_html(entry.get("title", "无标题"))
            summary = strip_html(entry.get("summary", ""))[:300]

            hit_must, hit_extra, priority = keyword_match(f"{title} {summary}", src_type)
            if not hit_must: continue

            year = pub.year if pub else datetime.now().year
            candidates.append({
                "uid":      uid,
                "source":   src["name"],
                "src_type": src_type,
                "title":    title,
                "link":     link,
                "summary":  summary,
                "pub":      pub,
                "pub_str":  (pub.astimezone(ZoneInfo("Asia/Shanghai")).strftime("%m-%d %H:%M")
                             if pub else "—"),
                "year":     year,
                "hit_extra":hit_extra,
                "priority": priority,
                "zh_desc":  "",
            })

    # 先按优先分值（高→低），再按发布时间（新→旧）
    candidates.sort(
        key=lambda a: (a["priority"], a["pub"].timestamp() if a["pub"] else 0),
        reverse=True,
    )
    return candidates

# ════════════════════════════════════════════════════════
#  AI 摘要（与 news_pusher.py 完全一致，包含相同 Prompt）
# ════════════════════════════════════════════════════════

def _call_claude(prompt: str, max_tokens: int = 800) -> str:
    if not ANTHROPIC_API_KEY: return ""
    try:
        r = requests.post(ANTHROPIC_URL, headers={
            "x-api-key":         ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type":      "application/json",
        }, json={"model": AI_MODEL, "max_tokens": max_tokens,
                 "messages": [{"role": "user", "content": prompt}]}, timeout=30)
        return r.json()["content"][0]["text"].strip()
    except Exception as e:
        print(f"  [AI ERR] {e}"); return ""

def ai_fill_articles(articles: list[dict]) -> None:
    """每篇文章生成中文一句话摘要（15~30字）"""
    if not articles or not ANTHROPIC_API_KEY:
        for a in articles:
            a["zh_desc"] = a["summary"][:80] + "…" if a["summary"] else a["title"]
        return

    items = "\n".join(
        f"{i+1}. 标题: {a['title']}\n   摘要: {a['summary'][:200]}"
        for i, a in enumerate(articles)
    )
    prompt = (
        "你是一名光学/AR领域的技术编辑。\n"
        "请为以下每篇文章生成一句简洁的中文描述（15~30字，直接说明该文章做了什么/提出了什么）。\n"
        "只输出编号和描述，格式严格如下，不要多余内容：\n"
        "1. 描述文字\n2. 描述文字\n...\n\n"
        f"{items}"
    )
    resp = _call_claude(prompt, max_tokens=600)
    if not resp:
        for a in articles:
            a["zh_desc"] = a["summary"][:80] + "…" if a["summary"] else a["title"]
        return

    desc_map: dict[int, str] = {}
    for line in [l.strip() for l in resp.splitlines() if l.strip()]:
        m = re.match(r"^(\d+)[\.、。：:]\s*(.+)", line)
        if m: desc_map[int(m.group(1))] = m.group(2).strip()

    for i, a in enumerate(articles):
        a["zh_desc"] = desc_map.get(i + 1, a["summary"][:80] + "…" if a["summary"] else a["title"])

def ai_generate_summary_and_risk(all_articles: list[dict]) -> tuple[str, str, str]:
    """生成总结（80~100字）、风险（50字）、机会（50字）"""
    if not all_articles or not ANTHROPIC_API_KEY:
        return "", "", ""

    titles = "、".join(a["title"] for a in all_articles)
    prompt = (
        "你是光学/AR/波导领域的技术分析师。\n"
        "根据以下文章标题，用中文输出三部分，格式严格如下（每部分一行，冒号后直接是内容）：\n"
        "总结: 用80~100字概括这批内容涵盖的主要技术方向与研究进展，内容充实具体\n"
        "风险: 用约50字说明潜在挑战或风险\n"
        "机会: 用约50字说明带来的市场或技术机会\n\n"
        f"文章标题：{titles}"
    )
    resp = _call_claude(prompt, max_tokens=500)
    if not resp: return "", "", ""

    summary_text = opportunity = risk = ""
    for line in resp.splitlines():
        line = line.strip()
        if   line.startswith("总结"): summary_text = line.split(":", 1)[-1].strip().lstrip("：")
        elif line.startswith("风险"): risk         = line.split(":", 1)[-1].strip().lstrip("：")
        elif line.startswith("机会"): opportunity  = line.split(":", 1)[-1].strip().lstrip("：")

    return summary_text, opportunity, risk

# ════════════════════════════════════════════════════════
#  PPTX + 图片生成管道
# ════════════════════════════════════════════════════════

def generate_pptx_and_image(slide_data: dict) -> tuple[str|None, str|None]:
    """
    生成 PPTX 和 JPG 图片
    返回 (pptx_path, img_path)，失败返回 None
    """
    # 1. 写 JSON
    with open(TMP_JSON, "w", encoding="utf-8") as f:
        json.dump(slide_data, f, ensure_ascii=False)

    # 2. Node.js → PPTX
    r = subprocess.run(["node", JS_SCRIPT, TMP_JSON, TMP_PPTX],
                       capture_output=True, text=True)
    if r.returncode != 0:
        print(f"  [JS ERR] {r.stderr}"); return None, None
    print(f"  ✅ PPTX 生成: {os.path.getsize(TMP_PPTX)//1024} KB")

    # 3. LibreOffice → PDF
    r2 = subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", "/tmp", TMP_PPTX],
        capture_output=True, text=True)
    if not os.path.exists(TMP_PDF):
        print(f"  [SOFFICE ERR] {r2.stderr}"); return TMP_PPTX, None
    print(f"  ✅ PDF 生成")

    # 4. pdftoppm → JPG（200 DPI）
    subprocess.run(["bash", "-c", "rm -f /tmp/daily_slide-*.jpg"], check=False)
    subprocess.run(
        ["pdftoppm", "-jpeg", "-r", "200", "-jpegopt", "quality=93",
         TMP_PDF, "/tmp/daily_slide"],
        capture_output=True, text=True)

    img = TMP_IMG if os.path.exists(TMP_IMG) else None
    if img:
        print(f"  ✅ 图片生成: {os.path.getsize(img)//1024} KB")
    return TMP_PPTX, img

# ════════════════════════════════════════════════════════
#  飞书图片直发（完整流程）
#  Step 1: App ID + App Secret → tenant_access_token
#  Step 2: 上传图片 → image_key
#  Step 3: 卡片内嵌 img 元素 → Webhook 发送
# ════════════════════════════════════════════════════════

FEISHU_APP_ID     = os.environ.get("FEISHU_APP_ID",     "")
FEISHU_APP_SECRET = os.environ.get("FEISHU_APP_SECRET", "")

FEISHU_API = "https://open.feishu.cn/open-apis"


def get_tenant_access_token() -> str | None:
    """用 App ID + App Secret 换取 tenant_access_token（有效期 2 小时）"""
    if not FEISHU_APP_ID or not FEISHU_APP_SECRET:
        print("  [INFO] 未配置 FEISHU_APP_ID / FEISHU_APP_SECRET，跳过图片上传")
        return None
    try:
        r = requests.post(
            f"{FEISHU_API}/auth/v3/tenant_access_token/internal",
            json={"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET},
            timeout=15,
        )
        data = r.json()
        if data.get("code") == 0:
            token = data["tenant_access_token"]
            print(f"  ✅ 飞书 Token 获取成功")
            return token
        print(f"  ❌ Token 获取失败: {data}")
        return None
    except Exception as e:
        print(f"  ❌ Token 异常: {e}")
        return None


def upload_image_to_feishu(img_path: str, token: str) -> str | None:
    """上传图片到飞书，返回 image_key（格式：img_xxxxxx）"""
    try:
        with open(img_path, "rb") as f:
            r = requests.post(
                f"{FEISHU_API}/im/v1/images",
                headers={"Authorization": f"Bearer {token}"},
                data={"image_type": "message"},
                files={"image": ("daily_slide.jpg", f, "image/jpeg")},
                timeout=60,
            )
        data = r.json()
        if data.get("code") == 0:
            key = data["data"]["image_key"]
            print(f"  ✅ 图片上传成功，image_key: {key[:20]}…")
            return key
        print(f"  ❌ 图片上传失败: {data}")
        return None
    except Exception as e:
        print(f"  ❌ 图片上传异常: {e}")
        return None


def upload_pptx_to_feishu(pptx_path: str, token: str, date_str: str) -> str | None:
    """上传 PPTX 文件到飞书，返回 file_key（需要 Bot API，Webhook 不支持文件消息）"""
    try:
        filename = f"AR光学日报_{date_str}.pptx"
        with open(pptx_path, "rb") as f:
            r = requests.post(
                f"{FEISHU_API}/im/v1/files",
                headers={"Authorization": f"Bearer {token}"},
                data={"file_type": "stream", "file_name": filename},
                files={"file": (filename, f,
                       "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
                timeout=60,
            )
        data = r.json()
        if data.get("code") == 0:
            key = data["data"]["file_key"]
            print(f"  ✅ PPTX 上传成功，file_key: {key[:20]}…")
            return key
        print(f"  ❌ PPTX 上传失败: {data}")
        return None
    except Exception as e:
        print(f"  ❌ PPTX 上传异常: {e}")
        return None


GITHUB_TOKEN      = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPOSITORY = os.environ.get("GITHUB_REPOSITORY", "")   # Actions 自动注入


def upload_pptx_to_github(pptx_path: str, date_str: str) -> str | None:
    """
    上传 PPTX 到 GitHub Releases（tag: daily-slides）
    返回公开下载 URL，失败返回 None
    """
    if not GITHUB_TOKEN or not GITHUB_REPOSITORY:
        print("  [INFO] 未配置 GITHUB_TOKEN，跳过 PPTX 上传")
        return None

    api   = f"https://api.github.com/repos/{GITHUB_REPOSITORY}"
    hdrs  = {"Authorization": f"token {GITHUB_TOKEN}",
              "Accept":        "application/vnd.github+json"}
    tag   = "daily-slides"
    fname = f"AR_Daily_{date_str}.pptx"

    # 获取或创建 Release
    r = requests.get(f"{api}/releases/tags/{tag}", headers=hdrs, timeout=15)
    if r.status_code == 200:
        data       = r.json()
        upload_url = data["upload_url"].split("{")[0]
        # 删除同名旧文件
        for asset in data.get("assets", []):
            if asset["name"] == fname:
                requests.delete(f"{api}/releases/assets/{asset['id']}",
                                headers=hdrs, timeout=15)
    else:
        r2 = requests.post(f"{api}/releases", headers=hdrs,
                           json={"tag_name": tag, "name": "Daily Slides",
                                 "body": "AR 光学日报存档", "prerelease": True},
                           timeout=15)
        if r2.status_code not in (200, 201):
            print(f"  [ERR] 创建 Release 失败: {r2.text[:100]}"); return None
        upload_url = r2.json()["upload_url"].split("{")[0]

    # 上传文件
    with open(pptx_path, "rb") as f:
        r3 = requests.post(
            f"{upload_url}?name={fname}",
            headers={"Authorization": f"token {GITHUB_TOKEN}",
                     "Content-Type": "application/octet-stream"},
            data=f, timeout=60,
        )
    if r3.status_code not in (200, 201):
        print(f"  [ERR] 上传失败: {r3.text[:100]}"); return None

    url = r3.json().get("browser_download_url", "")
    print(f"  ✅ PPTX 上传 GitHub 成功")
    return url


def push_to_feishu_image(img_path: str, pptx_path: str | None,
                         date_str: str, n_paper: int, n_news: int) -> bool:
    webhooks = [w for w in [FEISHU_WEBHOOK, FEISHU_WEBHOOK_2]
                if w and "YOUR_HOOK_ID" not in w and w.strip()]
    if not webhooks:
        print("  ❌ 未配置有效 Webhook"); return False

    # ── Step 1：飞书 Token → 上传图片 ────────────────
    token     = get_tenant_access_token()
    image_key = None
    if token and img_path and os.path.exists(img_path):
        image_key = upload_image_to_feishu(img_path, token)

    # ── Step 2：PPTX 上传 GitHub Release → 获取链接 ──
    pptx_url = None
    if pptx_path and os.path.exists(pptx_path):
        pptx_url = upload_pptx_to_github(pptx_path, date_str)

    # ── Step 3：构建卡片（标题 + 链接按钮 + 图片）────
    elements = []

    # 链接按钮（在图片上方，方便点击进入可编辑版本）
    if pptx_url:
        elements.append({
            "tag": "action",
            "actions": [{
                "tag":  "button",
                "text": {"tag": "plain_text",
                         "content": f"查看完整日报（含来源链接 · 可编辑）"},
                "type": "primary",
                "url":  pptx_url,
            }],
        })
        print(f"  📎 链接按钮已附加")

    # 图片（直接嵌入显示）
    if image_key:
        elements.append({
            "tag":     "img",
            "img_key": image_key,
            "alt":     {"tag": "plain_text", "content": f"AR光学日报 {date_str}"},
            "mode":    "fit_horizontal",
            "preview": True,
        })
        print("  📸 图片直发成功")
    else:
        print("  ⚠️  无图片（未配置 FEISHU_APP_ID/SECRET）")

    card_payload = {
        "msg_type": "interactive",
        "card": {
            "header": {
                "title": {
                    "tag":     "plain_text",
                    "content": f"AR 光学日报 | {date_str}",
                },
                "template": "blue",
            },
            "elements": elements,
        },
    }

    # ── Step 4：推送到所有 Webhook ────────────────────
    all_ok = True
    for i, url in enumerate(webhooks, 1):
        try:
            r      = requests.post(url,
                                   headers={"Content-Type": "application/json"},
                                   data=json.dumps(card_payload, ensure_ascii=False),
                                   timeout=30)
            result = r.json()
            if result.get("code") == 0 or result.get("StatusCode") == 0:
                print(f"  ✅ Webhook {i} 推送成功")
            else:
                print(f"  ❌ Webhook {i} 失败: {result}")
                all_ok = False
        except Exception as e:
            print(f"  ❌ Webhook {i} 异常: {e}"); all_ok = False
    return all_ok

# ════════════════════════════════════════════════════════
#  主流程
# ════════════════════════════════════════════════════════

def main():
    now_str = datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y-%m-%d %H:%M")
    today   = datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y-%m-%d")

    print(f"\n{'═'*54}")
    print(f"  🖼  AR 光学日报图片生成  {now_str}")
    ai_status = "✅ 已配置" if ANTHROPIC_API_KEY else "⚠️  未配置（将跳过 AI 摘要）"
    print(f"  AI 摘要: {ai_status}")
    print(f"{'─'*54}")

    history = load_history()
    print(f"  📖 历史记录 {len(history)} 条\n")

    # 抓取（与 news_pusher.py 完全一致）
    print("  📚 抓取学术文献…")
    paper_cands = fetch_candidates("paper", history)
    print(f"     候选 {len(paper_cands)} 篇 → 取最新 0~{QUOTA_PAPER} 篇")

    print("  🌐 抓取科技新闻…")
    news_cands = fetch_candidates("news", history)
    print(f"     候选 {len(news_cands)} 篇 → 取最新 {QUOTA_NEWS} 篇\n")

    paper_to_send = paper_cands[:QUOTA_PAPER]
    news_to_send  = news_cands[:QUOTA_NEWS]
    all_to_send   = paper_to_send + news_to_send

    if not all_to_send:
        print("⚠️  今日无新内容，跳过推送"); return

    # AI（与 news_pusher.py 完全一致的 Prompt）
    print("  🤖 AI 生成摘要…")
    ai_fill_articles(all_to_send)
    summary_text, opportunity, risk = ai_generate_summary_and_risk(all_to_send)

    # 打印预览
    print(f"\n{'─'*54}")
    for a in all_to_send:
        tag = "文献" if a["src_type"] == "paper" else "新闻"
        pri = f"优先分:{a['priority']}" if a["priority"] > 0 else "通用"
        print(f"  [{tag}][{pri}] [{a['year']}] {a['title'][:48]}")
        print(f"         → {a['zh_desc'][:50]}")
    print(f"{'─'*54}")

    # 构建幻灯片数据
    def fmt(a):
        return {"year": a["year"], "title": a["title"],
                "link": a["link"], "zh_desc": a["zh_desc"], "source": a["source"]}

    slide_data = {
        "date":        today,
        "summary":     summary_text or "今日 AR 光学领域持续活跃，多项技术进展值得关注。",
        "risk":        risk         or "技术落地仍面临成本、量产和用户体验挑战。",
        "opportunity": opportunity  or "相关技术持续发展，带来新应用场景与商业机遇。",
        "papers":      [fmt(a) for a in paper_to_send],
        "news":        [fmt(a) for a in news_to_send],
    }

    # 生成 PPTX + 图片
    print("\n  🎨 生成幻灯片…")
    pptx_path, img_path = generate_pptx_and_image(slide_data)
    if not pptx_path and not img_path:
        print("  ❌ 生成失败，跳过推送"); return

    # 推送飞书（图片卡片 + PPTX 文件）
    print("\n  📤 推送飞书…")
    ok = push_to_feishu_image(img_path, pptx_path, today,
                               len(paper_to_send), len(news_to_send))

    # 成功后更新历史（与 news_pusher.py 完全一致）
    if ok:
        for a in all_to_send:
            history[a["uid"]] = today
        save_history(history)

    print(f"{'═'*54}\n")

if __name__ == "__main__":
    main()
