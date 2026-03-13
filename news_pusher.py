"""
光学 / AR / Waveguide / Eyepiece 资讯推送飞书
─────────────────────────────────────────────────────────
输出格式：
  🗂 国外/国内科技资讯 | YYYY-MM-DD
  🔍 总结（AI 生成主题概览）
  ⚠️ 风险与机会（AI 生成）
  🗒 热点列表（编号 + 年份 + 标题 + 链接 + 中文一句话摘要）
─────────────────────────────────────────────────────────
· 科技新闻 3 篇 + 学术文献 0~3 篇
· sent_history.json 防重复，GitHub Actions 定时运行
"""

import feedparser
import hashlib
import json
import os
import re
from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo

import requests

# ════════════════════════════════════════════════════════
#  ① 配置
# ════════════════════════════════════════════════════════

FEISHU_WEBHOOK = os.environ.get(
    "FEISHU_WEBHOOK",
    "https://open.feishu.cn/open-apis/bot/v2/hook/YOUR_HOOK_ID",
)
# Anthropic API Key（用于 AI 生成总结 / 中文摘要）
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

QUOTA_NEWS        = 3       # 科技新闻固定 3 篇
QUOTA_PAPER       = 3       # 学术文献最多 3 篇
WINDOW_NEWS       = 48      # 新闻时间窗口（小时）
WINDOW_PAPER      = 7 * 24  # 文献时间窗口（小时）
HISTORY_FILE      = "sent_history.json"
HISTORY_KEEP_DAYS = 60

# ════════════════════════════════════════════════════════
#  ② 关键字
# ════════════════════════════════════════════════════════

KEYWORDS_MUST = [
    "optical", "optics", "photonics",
    "AR", "augmented reality",
    "waveguide", "光波导", "波导",
    "eyepiece", "目镜",
    "diffractive", "holographic", "hologram", "HOE",
    "metasurface", "metalens", "超表面", "元透镜",
    "grating", "光栅",
    "pancake lens", "birdbath",
    "light engine", "micro display",
    "microLED", "micro-LED", "OLED", "LCoS", "DLP",
    "mixed reality", "MR", "XR",
    "smart glasses", "head-mounted", "HMD",
    "Apple Vision", "Meta Quest", "HoloLens", "Magic Leap",
    "Ray-Ban Meta", "XREAL", "Rokid",
    "增强现实", "混合现实", "光学", "衍射", "全息",
    "智能眼镜", "头显",
    "VR", "virtual reality","TFLN",
]

KEYWORDS_EXTRA = [
    "eye tracking", "眼动追踪",
    "FOV", "field of view", "视场角",
    "aberration", "像差",
    "stray light", "杂散光",
    "diffraction efficiency",
    "ray tracing", "光线追迹",
    "tolerance", "manufacturing",
    "AI", "neural", "deep learning",
    "simulation",
]

# ════════════════════════════════════════════════════════
#  ③ RSS 源
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
#  去重历史
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
    cutoff = (datetime.now(timezone.utc) - timedelta(days=HISTORY_KEEP_DAYS)).strftime("%Y-%m-%d")
    cleaned = {k: v for k, v in history.items() if v >= cutoff}
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(cleaned, f, ensure_ascii=False, indent=2)
    print(f"  💾 历史记录 {len(cleaned)} 条")


# ════════════════════════════════════════════════════════
#  RSS 抓取
# ════════════════════════════════════════════════════════

def strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", "", text or "").strip()

def keyword_match(text: str) -> tuple[bool, list[str]]:
    lower = text.lower()
    hit   = any(k.lower() in lower for k in KEYWORDS_MUST)
    extra = [k for k in KEYWORDS_EXTRA if k.lower() in lower]
    return hit, extra

def parse_pub(entry) -> datetime | None:
    for attr in ("published_parsed", "updated_parsed"):
        t = getattr(entry, attr, None)
        if t:
            try:
                return datetime(*t[:6], tzinfo=timezone.utc)
            except Exception:
                pass
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
            print(f"  [ERR] {src['name']}: {e}")
            continue

        for entry in feed.entries[:60]:
            link = entry.get("link", "").strip()
            if not link:
                continue
            uid = url_id(link)
            if uid in history:
                continue
            pub = parse_pub(entry)
            if pub and pub < cutoff:
                continue

            title   = strip_html(entry.get("title", "无标题"))
            summary = strip_html(entry.get("summary", ""))[:300]

            hit_must, hit_extra = keyword_match(f"{title} {summary}")
            if not hit_must:
                continue

            # 年份：优先用发布年，否则用当前年
            year = pub.year if pub else datetime.now().year

            candidates.append({
                "uid":       uid,
                "source":    src["name"],
                "src_type":  src_type,
                "title":     title,
                "link":      link,
                "summary":   summary,
                "pub":       pub,
                "pub_str":   (pub.astimezone(ZoneInfo("Asia/Shanghai")).strftime("%m-%d %H:%M")
                              if pub else "—"),
                "year":      year,
                "hit_extra": hit_extra,
                "zh_desc":   "",   # 由 AI 填充
            })

    candidates.sort(
        key=lambda a: a["pub"] or datetime.min.replace(tzinfo=timezone.utc),
        reverse=True,
    )
    return candidates


# ════════════════════════════════════════════════════════
#  AI 功能（Claude claude-sonnet-4-20250514）
#  生成：① 每篇中文一句话摘要  ② 总结段落  ③ 风险与机会
# ════════════════════════════════════════════════════════

ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
AI_MODEL      = "claude-sonnet-4-20250514"


def _call_claude(prompt: str, max_tokens: int = 800) -> str:
    """调用 Anthropic Messages API，返回纯文本回复"""
    if not ANTHROPIC_API_KEY:
        return ""
    headers = {
        "x-api-key":         ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type":      "application/json",
    }
    body = {
        "model":      AI_MODEL,
        "max_tokens": max_tokens,
        "messages":   [{"role": "user", "content": prompt}],
    }
    try:
        r = requests.post(ANTHROPIC_URL, headers=headers, json=body, timeout=30)
        data = r.json()
        return data["content"][0]["text"].strip()
    except Exception as e:
        print(f"  [AI ERR] {e}")
        return ""


def ai_fill_articles(articles: list[dict]) -> None:
    """为每篇文章生成中文一句话摘要（批量一次调用）"""
    if not articles or not ANTHROPIC_API_KEY:
        # 无 API Key 时，退回到截取英文 summary
        for a in articles:
            a["zh_desc"] = a["summary"][:80] + "…" if a["summary"] else a["title"]
        return

    # 构建批量 prompt
    items = "\n".join(
        f"{i+1}. 标题: {a['title']}\n   摘要: {a['summary'][:200]}"
        for i, a in enumerate(articles)
    )
    prompt = (
        "你是一名光学/AR领域的技术编辑。\n"
        "请为以下每篇文章生成一句简洁的中文描述（50字，直接说明该文章做了什么/提出了什么）。\n"
        "只输出编号和描述，格式严格如下，不要多余内容：\n"
        "1. 描述文字\n2. 描述文字\n...\n\n"
        f"{items}"
    )
    resp = _call_claude(prompt, max_tokens=600)
    if not resp:
        for a in articles:
            a["zh_desc"] = a["summary"][:80] + "…" if a["summary"] else a["title"]
        return

    lines = [l.strip() for l in resp.splitlines() if l.strip()]
    desc_map: dict[int, str] = {}
    for line in lines:
        m = re.match(r"^(\d+)[\.、。：:]\s*(.+)", line)
        if m:
            desc_map[int(m.group(1))] = m.group(2).strip()

    for i, a in enumerate(articles):
        a["zh_desc"] = desc_map.get(i + 1, a["summary"][:80] + "…" if a["summary"] else a["title"])


def ai_generate_summary_and_risk(all_articles: list[dict]) -> tuple[str, str, str]:
    """
    返回 (主题概览一句话列表, 机会文字, 风险文字)
    无 API Key 时返回空字符串
    """
    if not all_articles or not ANTHROPIC_API_KEY:
        return "", "", ""

    titles = "、".join(a["title"] for a in all_articles)
    prompt = (
        "你是光学/AR/波导领域的技术分析师。\n"
        "根据以下文章标题，用中文简短输出三部分，格式严格如下（每部分一行，冒号后直接是内容）：\n"
        "总结: 用100字内概括这批内容涵盖的主要技术方向（逗号分隔关键主题）\n"
        "机会: 用50字内说明带来的市场或技术机会\n"
        "风险: 用50字内说明潜在挑战或风险\n\n"
        f"文章标题：{titles}"
    )
    resp = _call_claude(prompt, max_tokens=300)
    if not resp:
        return "", "", ""

    summary_text = opportunity = risk = ""
    for line in resp.splitlines():
        line = line.strip()
        if line.startswith("总结:") or line.startswith("总结："):
            summary_text = line.split(":", 1)[-1].strip().lstrip("：")
        elif line.startswith("机会:") or line.startswith("机会："):
            opportunity = line.split(":", 1)[-1].strip().lstrip("：")
        elif line.startswith("风险:") or line.startswith("风险："):
            risk = line.split(":", 1)[-1].strip().lstrip("：")

    return summary_text, opportunity, risk


# ════════════════════════════════════════════════════════
#  飞书卡片 ── 严格对照图片格式
# ════════════════════════════════════════════════════════

NUMBER_EMOJI = ["1️⃣","2️⃣","3️⃣","4️⃣","5️⃣","6️⃣","7️⃣","8️⃣","9️⃣","🔟"]

def build_feishu_card(
    news_list:    list[dict],
    paper_list:   list[dict],
    summary_text: str,
    opportunity:  str,
    risk:         str,
) -> dict:

    cn_date  = datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y-%m-%d")
    all_arts = paper_list + news_list          # 文献优先展示
    n_paper  = len(paper_list)
    n_news   = len(news_list)

    elements: list[dict] = []

    # ── 1. 总结区 ────────────────────────────────────
    summary_line = summary_text or "、".join(
        a["title"][:18] + "…" for a in all_arts[:3]
    )
    elements.append({
        "tag": "markdown",
        "content": (
            f"**🔍 总结**\n"
            f"**📊 本次推送包含：{n_paper}篇文献 / {n_news}篇新闻**\n"
            f"{summary_line}。"
        ),
    })

    # ── 2. 风险与机会 ────────────────────────────────
    opp_text  = opportunity or "相关技术持续发展，带来新应用场景与商业机遇。"
    risk_text = risk        or "技术落地仍面临成本、量产和用户体验挑战。"
    elements.append({
        "tag": "markdown",
        "content": (
            f"**⚠️ 风险与机会**\n"
            f"机会：{opp_text}\n"
            f"风险：{risk_text}"
        ),
    })

    elements.append({"tag": "hr"})

    # ── 3. 文献区块 ──────────────────────────────────
    if paper_list:
        elements.append({
            "tag": "markdown",
            "content": f"**📚 文献 ·** {n_paper} 篇",
        })
        for idx, a in enumerate(paper_list):
            num_emoji = NUMBER_EMOJI[idx] if idx < len(NUMBER_EMOJI) else f"{idx+1}."
            desc      = a.get("zh_desc") or a["summary"][:60] or a["title"]
            elements.append({
                "tag": "markdown",
                "content": (
                    f"{num_emoji} **[{a['year']}] {a['title']}**\n"
                    f"{a['link']}\n"
                    f"{desc}"
                ),
            })
        elements.append({"tag": "hr"})

    # ── 4. 新闻区块 ──────────────────────────────────
    if news_list:
        elements.append({
            "tag": "markdown",
            "content": f"**📰 新闻 ·** {n_news} 篇",
        })
        for idx, a in enumerate(news_list):
            num_emoji = NUMBER_EMOJI[idx] if idx < len(NUMBER_EMOJI) else f"{idx+1}."
            desc      = a.get("zh_desc") or a["summary"][:60] or a["title"]
            elements.append({
                "tag": "markdown",
                "content": (
                    f"{num_emoji} **[{a['year']}] {a['title']}**\n"
                    f"{a['link']}\n"
                    f"{desc}"
                ),
            })

    return {
        "msg_type": "interactive",
        "card": {
            "header": {
                "title": {
                    "tag":     "plain_text",
                    "content": f"🗂 国外/国内科技资讯 | {cn_date}",
                },
                "template": "blue",
            },
            "elements": elements,
        },
    }


# ════════════════════════════════════════════════════════
#  飞书推送
# ════════════════════════════════════════════════════════

def push_to_feishu(payload: dict) -> bool:
    try:
        resp   = requests.post(
            FEISHU_WEBHOOK,
            headers={"Content-Type": "application/json"},
            data=json.dumps(payload, ensure_ascii=False),
            timeout=15,
        )
        result = resp.json()
        if result.get("code") == 0 or result.get("StatusCode") == 0:
            print("✅ 飞书推送成功")
            return True
        print(f"❌ 推送失败: {result}")
        return False
    except Exception as e:
        print(f"❌ 推送异常: {e}")
        return False


# ════════════════════════════════════════════════════════
#  主流程
# ════════════════════════════════════════════════════════

def main():
    now_str = datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y-%m-%d %H:%M")
    today   = datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y-%m-%d")

    print(f"\n{'═'*54}")
    print(f"  🔬 光学 & AR 资讯推送  {now_str}")
    ai_status = "✅ 已配置" if ANTHROPIC_API_KEY else "⚠️  未配置（将跳过 AI 摘要）"
    print(f"  AI 摘要: {ai_status}")
    print(f"{'─'*54}")

    history = load_history()
    print(f"  📖 历史记录 {len(history)} 条\n")

    # 抓取候选
    print("  🌐 抓取科技新闻…")
    news_cands  = fetch_candidates("news",  history)
    print(f"     候选 {len(news_cands)} 篇 → 取最新 {QUOTA_NEWS} 篇")

    print("  📚 抓取学术文献…")
    paper_cands = fetch_candidates("paper", history)
    print(f"     候选 {len(paper_cands)} 篇 → 取最新 0~{QUOTA_PAPER} 篇\n")

    news_to_send  = news_cands[:QUOTA_NEWS]
    paper_to_send = paper_cands[:QUOTA_PAPER]
    all_to_send   = paper_to_send + news_to_send

    if not all_to_send:
        print("⚠️  今日无新内容，跳过推送")
        return

    # AI：生成中文摘要 + 总结 + 风险机会
    print("  🤖 AI 生成摘要…")
    ai_fill_articles(all_to_send)
    summary_text, opportunity, risk = ai_generate_summary_and_risk(all_to_send)

    # 打印预览
    print(f"\n{'─'*54}")
    for a in all_to_send:
        tag = "文献" if a["src_type"] == "paper" else "新闻"
        print(f"  [{tag}] [{a['year']}] {a['title'][:50]}")
        print(f"         → {a['zh_desc'][:50]}")
    print(f"{'─'*54}")

    # 推送
    payload = build_feishu_card(
        news_to_send, paper_to_send,
        summary_text, opportunity, risk,
    )
    ok = push_to_feishu(payload)

    # 成功后更新历史
    if ok:
        for a in all_to_send:
            history[a["uid"]] = today
        save_history(history)

    print(f"{'═'*54}\n")


if __name__ == "__main__":
    main()
