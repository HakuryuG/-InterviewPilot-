
import streamlit as st
import pandas as pd
import altair as alt
import streamlit.components.v1 as components
import re
import io
import os
import json
import hashlib
import traceback
import sqlite3
import uuid
from html import escape
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from typing import List


from openai import OpenAI
try:
    from docx import Document
except ImportError:
    Document = None


st.set_page_config(
    page_title="InterviewPilot",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root {
        --ip-primary: #0ea5e9;
        --ip-accent: #22c55e;
        --ip-text-muted: #94a3b8;
        --ip-border: rgba(148,163,184,0.24);
    }
    section[data-testid="stSidebar"] * { font-size: 0.92rem !important; }
    .hero-card {
        padding: 1.25rem 1.45rem;
        border: 1px solid var(--ip-border);
        border-radius: 16px;
        background:
          radial-gradient(circle at top right, rgba(14,165,233,0.12), transparent 45%),
          radial-gradient(circle at bottom left, rgba(34,197,94,0.10), transparent 40%),
          linear-gradient(160deg, rgba(15,23,42,0.55), rgba(2,6,23,0.45));
        box-shadow: 0 8px 26px rgba(2,6,23,0.18);
        margin-bottom: 0.85rem;
    }
    .hero-title {
        font-size: 1.85rem;
        font-weight: 700;
        letter-spacing: 0.2px;
        margin-bottom: 0.25rem;
    }
    .hero-subtitle {
        font-size: 0.94rem;
        color: #c7d2fe;
        margin-bottom: 0.55rem;
    }
    .hero-meta {
        font-size: 0.88rem;
        color: var(--ip-text-muted);
    }
    .tag-row {
        margin-top: 0.72rem;
        display: flex;
        flex-wrap: wrap;
        gap: 0.42rem;
    }
    .tag {
        display: inline-block;
        padding: 0.18rem 0.62rem;
        border-radius: 999px;
        background: rgba(14,165,233,0.12);
        border: 1px solid rgba(14,165,233,0.22);
        font-size: 0.8rem;
    }
    .footer-note {
        margin-top: 1.6rem;
        padding: 0.9rem 0.2rem 0.2rem 0.2rem;
        color: var(--ip-text-muted);
        font-size: 0.9rem;
        text-align: center;
    }
    .quota-note {
        text-align: center;
        color: var(--ip-text-muted);
        font-size: 0.92rem;
        margin-top: 0.35rem;
        margin-bottom: 0.35rem;
    }
    .ip-kpi-wrap {
        max-width: 1200px;
        margin: 0 auto 0.7rem auto;
    }
    .ip-kpi-grid {
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 0.8rem;
    }
    .ip-kpi-card {
        min-height: 132px;
        padding: 0.8rem 0.9rem;
        border: 1px solid var(--ip-border);
        border-radius: 16px;
        background: rgba(15,23,42,0.25);
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        text-align: center;
    }
    .ip-kpi-label {
        font-size: 0.95rem;
        color: #e2e8f0;
        margin-bottom: 0.5rem;
        line-height: 1.35;
    }
    .ip-kpi-value {
        font-size: clamp(2.1rem, 2.2vw, 3.3rem);
        line-height: 1.08;
        color: #f8fafc;
        font-weight: 700;
        letter-spacing: 0.2px;
    }
    @media (max-width: 980px) {
        .ip-kpi-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr));
        }
    }
    @media (max-width: 560px) {
        .ip-kpi-grid {
            grid-template-columns: 1fr;
        }
    }
    @keyframes ip-spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    /* Replace the top-right running-man status icon with a circle spinner */
    [data-testid="stStatusWidget"] svg {
        display: none !important;
    }
    [data-testid="stStatusWidget"]::before {
        content: "";
        width: 12px;
        height: 12px;
        border: 2px solid rgba(148,163,184,0.45);
        border-top-color: var(--ip-primary);
        border-radius: 50%;
        display: inline-block;
        margin-right: 6px;
        animation: ip-spin 0.8s linear infinite;
    }
    /* Hide status widget only; keep toolbar for sidebar toggle */
    [data-testid="stStatusWidget"] {
        display: none !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Force-open sidebar on first load as a fallback for environments
# where Streamlit doesn't honor initial_sidebar_state.
components.html(
    """
    <script>
    (function () {
      const doc = window.parent.document;

      const findOpenButton = () => {
        const selectors = [
          'button[aria-label="Open sidebar"]',
          'button[aria-label*="sidebar"]',
          '[data-testid="stSidebarCollapsedControl"] button',
          '[data-testid="collapsedControl"] button',
        ];
        for (const sel of selectors) {
          const btn = doc.querySelector(sel);
          if (btn) return btn;
        }
        return null;
      };

      const isSidebarExpanded = () => {
        const sidebar = doc.querySelector('section[data-testid="stSidebar"]');
        if (!sidebar) return false;
        const ariaExpanded = sidebar.getAttribute("aria-expanded");
        if (ariaExpanded === "true") return true;
        const left = window.getComputedStyle(sidebar).left;
        return left !== "-336px" && left !== "-320px";
      };

      const tryOpenSidebar = () => {
        if (isSidebarExpanded()) return true;
        const openButton = findOpenButton();
        if (!openButton) return false;
        openButton.click();
        return isSidebarExpanded();
      };

      let attempts = 0;
      const timer = setInterval(() => {
        attempts += 1;
        const done = tryOpenSidebar();
        if (done || attempts >= 40) {
          clearInterval(timer);
        }
      }, 150);
    })();
    </script>
    """,
    height=0,
    width=0,
)

# DeepSeek API client setup
DEFAULT_DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "")
if "deepseek_api_key" not in st.session_state:
    st.session_state.deepseek_api_key = ""
if "use_personal_key" not in st.session_state:
    st.session_state.use_personal_key = False
if "analysis_mode" not in st.session_state:
    st.session_state.analysis_mode = "AI标准分析"
if "last_signature" not in st.session_state:
    st.session_state.last_signature = ""
if "cached_ai_insights" not in st.session_state:
    st.session_state.cached_ai_insights = {"research_findings": "", "research_suggestions": ""}
if "cached_full_report" not in st.session_state:
    st.session_state.cached_full_report = ""
if "cached_theme_brief" not in st.session_state:
    st.session_state.cached_theme_brief = ""
if "user_quota_id" not in st.session_state:
    st.session_state.user_quota_id = uuid.uuid4().hex[:12]
if "today_used_local" not in st.session_state:
    st.session_state.today_used_local = 0
if "today_used_date" not in st.session_state:
    st.session_state.today_used_date = datetime.now().strftime("%Y-%m-%d")


def get_effective_api_key() -> str:
    raw_key = (
        st.session_state.get("deepseek_api_key", "")
        if st.session_state.get("use_personal_key")
        else DEFAULT_DEEPSEEK_API_KEY
    )
    # Remove common copy/paste wrappers and invisible whitespace.
    key = str(raw_key).strip().strip('"').strip("'").replace("\u200b", "").replace("\ufeff", "")
    return key


def get_ai_client() -> OpenAI | None:
    api_key = get_effective_api_key()
    if not api_key:
        return None
    try:
        api_key.encode("ascii")
    except UnicodeEncodeError:
        # Invalid key format (contains full-width or other non-ASCII chars).
        return None
    return OpenAI(
        api_key=api_key,
        base_url="https://api.deepseek.com",
    )


client = get_ai_client()
ai_status_message = "已连接 DeepSeek，可进行AI深度分析" if client is not None else "未检测到 API Key，当前不可进行AI深度分析"
MAX_AI_INPUT_CHARS = 1800
USAGE_DB_PATH = "interviewpilot_usage.db"
DAILY_ANALYZE_LIMIT = 100

THEME_RULES = [
    ("数字医疗服务", ["医院", "挂号", "复诊", "医保", "就诊", "门诊", "缴费", "大夫", "挂号费"]),
    ("政务数字化服务", ["政务", "社保", "认证", "窗口", "网上办理", "材料", "办成率", "年审", "补审"]),
    ("数字平台经营", ["小超市", "进货", "平台卖货", "缴税", "发票", "平台规则", "账期", "结算", "保证金", "提现"]),
    ("教育数字平台", ["选课", "助学金", "论文", "校园网", "实习", "奖学金", "大一新生", "报到", "贫困证明"]),
    ("反诈与数字安全", ["反诈", "诈骗", "冻结账户", "风险", "短信", "链接", "支付节点", "验证码", "钓鱼"]),
]

THEME_KEYWORD_BANK = {
    "数字医疗服务": ["线上挂号", "医院系统", "预约失败", "界面复杂", "人工窗口"],
    "政务数字化服务": ["网上办理", "材料上传", "审核进度", "系统卡顿", "账号密码"],
    "数字平台经营": ["平台规则", "资金周转", "申诉困难", "保证金", "提现账期"],
    "教育数字平台": ["选课系统", "助学金", "材料上传", "网络中断", "无障碍支持"],
    "反诈与数字安全": ["诈骗短信", "钓鱼链接", "账户冻结", "银行卡", "验证码"],
    "日常数字生活": ["智能手机", "网上办事", "人工窗口", "老年人", "数字鸿沟"],
}

GLOBAL_KEYWORD_CANDIDATES = [
    "线上办理", "系统卡顿", "材料上传", "账号密码", "网上预约", "人工窗口",
    "平台规则", "资金周转", "保证金", "提现", "选课系统", "助学金申请",
    "诈骗短信", "钓鱼链接", "验证码", "智能手机", "数字鸿沟", "无障碍支持",
    "界面复杂", "网上办理", "审核进度", "操作失败",
]


def init_usage_db() -> None:
    conn = sqlite3.connect(USAGE_DB_PATH)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS usage_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                event TEXT NOT NULL,
                detail TEXT,
                value INTEGER DEFAULT 1,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def log_usage_event(event: str, detail: str = "", value: int = 1) -> None:
    try:
        conn = sqlite3.connect(USAGE_DB_PATH)
        try:
            conn.execute(
                "INSERT INTO usage_events(event, detail, value, created_at) VALUES (?, ?, ?, ?)",
                (event, detail[:200], int(value), datetime.now().isoformat()),
            )
            conn.commit()
        finally:
            conn.close()
    except Exception:
        pass


def get_usage_summary(days: int = 30) -> dict:
    since = (datetime.utcnow() - timedelta(days=days)).isoformat()
    conn = sqlite3.connect(USAGE_DB_PATH)
    try:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT event, COALESCE(SUM(value), 0) AS cnt
            FROM usage_events
            WHERE created_at >= ?
            GROUP BY event
            """,
            (since,),
        )
        rows = cur.fetchall()
    finally:
        conn.close()
    return {k: int(v) for k, v in rows}


def get_today_analyze_count(user_id: str) -> int:
    today = datetime.now().strftime("%Y-%m-%d")
    conn = sqlite3.connect(USAGE_DB_PATH)
    try:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT COALESCE(SUM(value), 0)
            FROM usage_events
            WHERE event = 'analyze_success'
              AND (detail = ? OR detail = '')
              AND substr(created_at, 1, 10) = ?
            """,
            (user_id, today),
        )
        row = cur.fetchone()
        return int(row[0] or 0)
    finally:
        conn.close()


def get_daily_usage(days: int = 14) -> pd.DataFrame:
    since = (datetime.utcnow() - timedelta(days=days)).isoformat()
    conn = sqlite3.connect(USAGE_DB_PATH)
    try:
        query = """
        SELECT substr(created_at, 1, 10) AS 日期, event AS 事件, COALESCE(SUM(value), 0) AS 次数
        FROM usage_events
        WHERE created_at >= ?
        GROUP BY substr(created_at, 1, 10), event
        ORDER BY 日期 DESC, 次数 DESC
        """
        df = pd.read_sql_query(query, conn, params=(since,))
    finally:
        conn.close()
    return df


def run_deploy_precheck() -> list[tuple[str, str, str]]:
    checks: list[tuple[str, str, str]] = []
    key = get_effective_api_key()
    checks.append(("API Key 配置", "通过" if bool(key) else "未通过", "已配置" if key else "未检测到 DEEPSEEK_API_KEY"))
    if key:
        try:
            key.encode("ascii")
            checks.append(("API Key 格式", "通过", "ASCII格式有效"))
        except UnicodeEncodeError:
            checks.append(("API Key 格式", "未通过", "包含非ASCII字符，请重新粘贴"))
    checks.append(("AI 客户端", "通过" if client is not None else "未通过", "可调用" if client is not None else "当前不可用"))
    checks.append(("python-docx", "通过" if Document is not None else "提示", "可导出Word" if Document is not None else "未安装，无法导出Word"))
    checks.append(("分析模式", "通过", st.session_state.get("analysis_mode", "AI标准分析")))
    return checks


init_usage_db()


def records_signature(records: List[dict]) -> str:
    payload = [
        {
            "访谈ID": item.get("访谈ID", ""),
            "原始文本": item.get("原始文本", ""),
            "访谈摘要": item.get("访谈摘要", ""),
            "核心主题": item.get("核心主题", ""),
            "情绪": item.get("情绪", ""),
            "全部关键词": item.get("全部关键词", ""),
        }
        for item in records
    ]
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    return hashlib.md5(raw.encode("utf-8")).hexdigest()

# Title and description
st.markdown(
    """
    <div class="hero-card">
        <div class="hero-title">🌲 InterviewPilot｜访谈分析助手</div>
        <div class="hero-subtitle">用于访谈文本的整理、关键词提取、主题统计与报告生成</div>
        <div class="hero-meta">创作者：公管小白 ｜ 北京师范大学</div>
        <div class="hero-meta">如有问题联系：yuqingsuen@163.com</div>
        <div class="tag-row">
            <span class="tag">访谈分析</span>
            <span class="tag">定性研究</span>
            <span class="tag">关键词提取</span>
            <span class="tag">报告导出</span>
            <span class="tag">Version 1.4.0</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)



def split_text_into_paragraphs(text: str) -> List[str]:
    """将长文本按段落分割。识别句号、问号、感叹号加空行作为段落分隔符。"""
    chunks = re.split(r"(?<=[。！？\?!])\s*\n\s*|\n\s*\n+", text.strip())
    paragraphs = [p.strip() for p in chunks if p.strip()]
    return paragraphs if paragraphs else [text.strip()]


def analyze_interview_rule_based(text: str) -> dict:
    """极速本地分析：低算力、稳定、适配长文本。"""
    raw = text.strip()
    preview = raw.replace("\n", " ")
    if len(preview) > 180:
        preview = preview[:180] + "..."

    features = extract_local_features(raw)
    respondent_type = features["受访者类型"]
    final_theme = features["核心主题"]
    final_emotion = features["情绪"]
    core_keywords = features["关键词列表"][:3]
    all_keywords = features["关键词列表"][:5]

    key_points = []
    if "提示不清" in all_keywords or "流程复杂" in all_keywords:
        key_points.append("受访者主要困难集中在流程复杂和系统提示不清。")
    if "线下兜底" in all_keywords or "人工辅导" in all_keywords:
        key_points.append("线上失败后仍高度依赖线下窗口或人工辅导兜底。")
    if not key_points:
        key_points.append("数字化提升效率的同时，对低数字能力群体形成额外门槛。")
    key_points = key_points[:2]

    summary = f"围绕“{final_theme}”的访谈显示，主要问题为“{'、'.join(core_keywords[:2])}”，整体情绪偏“{final_emotion}”。"

    paragraph_results = [
        {
            "段落文本": preview,
            "受访者类型": respondent_type,
            "年龄段": features["年龄段"],
            "核心主题": final_theme,
            "情绪": final_emotion,
            "关键词": "、".join(core_keywords),
        }
    ]

    return {
        "访谈摘要": summary,
        "受访者类型": respondent_type,
        "年龄段": features["年龄段"],
        "核心主题": final_theme,
        "情绪": final_emotion,
        "关键词": "、".join(core_keywords),
        "全部关键词": "、".join(all_keywords),
        "关键要点": "；".join(key_points),
        "段落数": 1,
        "分段分析": paragraph_results,
    }


def extract_local_features(raw: str) -> dict:
    respondent_type = "普通用户"
    # 基层工作人员识别
    if any(w in raw for w in ["网格员", "网格", "窗口", "工作人员", "信息科", "民警", "护士", "老师", "医生", "保安", "物业", "保洁", "居委"]):
        respondent_type = "基层工作人员"
    elif any(w in raw for w in ["学生", "大三", "大四", "大二", "大一", "大学", "研究生", "博士", "硕士", "本科生", "高中生", "初中生", "小学生"]):
        respondent_type = "青年/学生群体"
    elif any(w in raw for w in ["退休", "老人", "老年", "大爷", "大妈", "爷爷", "奶奶", "外婆", "外公"]):
        respondent_type = "老年群体"
    elif any(w in raw for w in ["宝妈", "宝爸", "妈妈", "爸爸", "家长", "带孩子", "孩子"]):
        respondent_type = "家长群体"
    elif any(w in raw for w in ["残障", "残疾", "盲人", "聋人", "轮椅", "行动不便"]):
        respondent_type = "特殊群体"
    elif any(w in raw for w in ["外来务工", "农民工", "流动人口", "租房", "租客"]):
        respondent_type = "流动人口"

    # 年龄识别 - 精确匹配
    age_match = re.search(r"(\d{1,3})\s*岁", raw)
    if age_match:
        age_num = int(age_match.group(1))
        if age_num < 18:
            age_group = f"{age_num}岁（未成年）"
        elif age_num <= 25:
            age_group = f"{age_num}岁（青年）"
        elif age_num <= 40:
            age_group = f"{age_num}岁（中年）"
        elif age_num <= 60:
            age_group = f"{age_num}岁（中老年）"
        else:
            age_group = f"{age_num}岁（高龄）"
    else:
        # 尝试根据上下文推断年龄段
        if any(w in raw for w in ["未成年", "小孩", "儿童", "学生", "高考", "中考"]):
            age_group = "未成年"
        elif any(w in raw for w in ["年轻", "青年", "刚退休", "三十", "四十", "五十"]):
            age_group = "成年"
        elif any(w in raw for w in ["六十", "七十", "八十", "老年", "退休", "老人"]):
            age_group = "老年"
        else:
            age_group = "未知"

    final_theme = "日常数字生活"
    for theme, words in THEME_RULES:
        if any(w in raw for w in words):
            final_theme = theme
            break

    anxiety_words = ["担心", "焦虑", "紧张", "害怕", "失败", "排队", "被骗", "挫败", "拖到", "压力", "麻烦", "无奈", "难受", "尴尬", "不满", "气愤", "委屈", "崩溃", "头疼"]
    positive_words = ["方便", "提升", "省事", "顺畅", "高效", "认可", "感谢", "满意", "感谢", "挺好"]
    if any(w in raw for w in anxiety_words):
        emotion = "焦虑"
    elif any(w in raw for w in positive_words):
        emotion = "积极"
    else:
        emotion = "中性"

    limit = keyword_limit_by_text(raw)
    theme_candidates = THEME_KEYWORD_BANK.get(final_theme, THEME_KEYWORD_BANK["日常数字生活"])
    merged_candidates = list(dict.fromkeys(theme_candidates + GLOBAL_KEYWORD_CANDIDATES))
    matched = [k for k in merged_candidates if k in raw]
    keywords = matched if matched else theme_candidates[:3]
    keywords = list(dict.fromkeys(keywords))[:limit]

    return {
        "受访者类型": respondent_type,
        "年龄段": age_group,
        "核心主题": final_theme,
        "情绪": emotion,
        "关键词列表": keywords,
    }


def parse_ai_json_response(content: str) -> dict:
    """尽量稳妥地解析模型返回的 JSON。"""
    content = content.strip()
    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?", "", content).strip()
        content = re.sub(r"```$", "", content).strip()

    start = content.find("{")
    end = content.rfind("}")
    if start != -1 and end != -1 and end > start:
        content = content[start:end + 1]

    return json.loads(content)


def compress_text_for_ai(text: str, max_chars: int = MAX_AI_INPUT_CHARS) -> str:
    """压缩超长访谈输入，控制 token 成本。"""
    clean = str(text).strip()
    if len(clean) <= max_chars:
        return clean
    head = clean[: max_chars // 2]
    tail = clean[-(max_chars // 2):]
    return f"{head}\n\n[中间内容已省略]\n\n{tail}"


def parse_labeled_fallback(content: str, raw_text: str) -> dict:
    """解析固定标签文本，作为 JSON 失败后的兜底。"""
    def pick(label: str, default: str = "") -> str:
        m = re.search(rf"{label}\s*[:：]\s*(.+)", content)
        return m.group(1).strip() if m else default

    summary = pick("访谈摘要", "该访谈已完成分析。")
    respondent_self = pick("受访者自述身份", "")
    respondent = pick("受访者类型", "普通用户")
    具体年龄 = pick("具体年龄", "")
    年龄段 = pick("年龄段", "未知")
    theme = pick("核心主题", "日常数字生活")
    emotion = pick("情绪", "中性")
    kw_text = pick("关键词", "")
    kp_text = pick("关键要点", "")

    # 合并具体年龄和年龄段
    if 具体年龄 and 年龄段 and 年龄段 != "未知":
        年龄显示 = f"{具体年龄}（{年龄段}）"
    elif 具体年龄:
        年龄显示 = 具体年龄
    else:
        年龄显示 = 年龄段

    keywords = [k.strip() for k in re.split(r"[、,，;；/|]", kw_text) if k.strip()][:5]
    key_points = [k.strip() for k in re.split(r"[；;\n]", kp_text) if k.strip()][:2]
    preview_text = raw_text.strip().replace("\n", " ")
    if len(preview_text) > 160:
        preview_text = preview_text[:160] + "..."

    return {
        "访谈摘要": summary,
        "受访者自述身份": respondent_self,
        "受访者类型": respondent,
        "年龄段": 年龄显示,
        "核心主题": theme,
        "情绪": emotion,
        "情绪词": emotion,
        "关键词": keywords if keywords else ["访谈"],
        "关键要点": key_points,
        "分段分析": [
            {
                "段落文本": preview_text,
                "受访者自述身份": respondent_self,
                "受访者类型": respondent,
                "年龄段": 年龄显示,
                "核心主题": theme,
                "情绪": emotion,
                "关键词": keywords[:3] if keywords else ["访谈"],
            }
        ],
    }


def normalize_age_group(raw_age: str, source_text: str) -> str:
    """年龄规范化：优先保留明确年龄；无明确则未知；推测需加前缀。"""
    explicit = re.search(r"(\d{1,3})\s*岁", source_text)
    if explicit:
        return f"{explicit.group(1)}岁"
    age = str(raw_age).strip()
    if not age or age in {"未知", "不详", "未提及", "无"}:
        return "未知"
    if age.startswith("推测:"):
        return age
    if re.fullmatch(r"\d{1,3}\s*岁", age):
        return age.replace(" ", "")
    if age in {"60+", "26-59", "18-25"}:
        return f"推测:{age}"
    return f"推测:{age}"


def normalize_emotion(raw_emotion: str) -> str:
    emotion = str(raw_emotion).strip()
    if emotion not in {"中性", "焦虑", "积极"}:
        return "中性"
    return emotion


def normalize_emotion_phrase(raw_emotion: str, fallback: str) -> str:
    tokens = [t.strip() for t in re.split(r"[、,，/;；\s]+", str(raw_emotion)) if t.strip()]
    if not tokens:
        return fallback
    allow = {"中性", "焦虑", "积极", "谨慎", "无奈", "乐观", "担忧", "困惑", "压力"}
    filtered = [t for t in tokens if t in allow]
    if not filtered:
        filtered = tokens[:2]
    return "、".join(filtered[:2])


def keyword_limit_by_text(text: str) -> int:
    n = len(text or "")
    if n > 3000:
        return 12
    if n > 1600:
        return 10
    if n > 800:
        return 8
    return 6


def analyze_interview_ai(text: str) -> dict:
    """优先使用 DeepSeek 进行分析，失败时抛出异常供上层回退。"""
    if client is None:
        raise RuntimeError("DEEPSEEK_API_KEY 未配置")

    ai_text = compress_text_for_ai(text)
    prompt = f"""
你是一名中文访谈分析助手。请仔细阅读下面的访谈文本，提取关键信息并输出JSON。

访谈文本：
{ai_text}

要求：
1. 根据文本内容，准确识别以下字段：
   - "受访者自述身份"：直接提取受访者在文本中明确说出的身份，只保留身份本身（如"社区网格员"、"大三学生"、"小超市老板"、"视障人士"、"外卖骑手"），不要"我是"、"我是一名"等前缀
   - "受访者类型"：根据文本描述判断分类（普通用户/基层工作人员/青年/学生群体/老年群体/特殊群体等）
   - "具体年龄"：如果文本明确提到年龄（如"我今年68岁"或"49岁"），提取具体数字；否则设为空
   - "年龄段"：根据具体年龄判断年龄段：18岁以下返回"未成年"，18-25岁返回"青年"，26-40岁返回"中年"，41-60岁返回"中老年"，60岁以上返回"高龄"；如果无法确定具体年龄，则返回"未知"
   - "核心主题"：从以下主题中选择最符合的（数字医疗服务/政务数字化服务/教育数字平台/反诈与数字安全/数字平台经营/日常数字生活）
   - "情绪"：根据文本整体情感判断（焦虑/中性/积极/无奈/不满等）
   - "关键词"：提取3-5个核心关键词，反映访谈主要内容
   - "访谈摘要"：用40-70字概括访谈主要内容
   - "关键要点"：提取1-2条关键洞察

2. 不要编造信息，如果文本中未提及某个信息，该字段留空或设为"未知"
3. 输出格式：
{{
  "受访者自述身份": "",
  "受访者类型": "",
  "具体年龄": "",
  "年龄段": "",
  "核心主题": "",
  "情绪": "",
  "关键词": ["", "", ""],
  "访谈摘要": "",
  "关键要点": ["", ""]
}}

请直接输出JSON，不要其他内容：
"""

    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是一个高效、节省算力的中文访谈分析助手，擅长从访谈文本中提取关键信息。"},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0.0,
            max_tokens=350,
            timeout=15,
        )
        content = response.choices[0].message.content
        data = parse_ai_json_response(content)
    except json.JSONDecodeError:
        retry_prompt = f"""
请基于以下访谈文本输出合法JSON，只包含字段：受访者自述身份、受访者类型、具体年龄、年龄段、核心主题、情绪、关键词(数组)、访谈摘要、关键要点(数组)。

访谈文本：
{ai_text}

直接输出JSON：
"""
        retry_resp = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是JSON输出助手。"},
                {"role": "user", "content": retry_prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0.0,
            max_tokens=350,
            timeout=12,
        )
        retry_content = retry_resp.choices[0].message.content
        try:
            data = parse_ai_json_response(retry_content)
        except json.JSONDecodeError:
            text_retry_prompt = f"""
请仅按以下固定字段输出，不要JSON，不要Markdown：
受访者自述身份：...
受访者类型：...
具体年龄：...
年龄段：...
核心主题：...
情绪：...
关键词：关键词1、关键词2、关键词3
访谈摘要：...
关键要点：要点1；要点2

访谈文本：
{ai_text}
"""
            text_retry = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "你是结构化摘要助手，只按指定字段输出。"},
                    {"role": "user", "content": text_retry_prompt},
                ],
                temperature=0.0,
                max_tokens=300,
                timeout=10,
            )
            data = parse_labeled_fallback(text_retry.choices[0].message.content, text)

    # 解析AI返回的关键词
    keywords = data.get("关键词", [])
    if isinstance(keywords, str):
        keywords = [k.strip() for k in re.split(r"[、,，;；/|]", keywords) if k.strip()]
    keywords = [str(k).strip() for k in keywords if str(k).strip()][:5]

    # 解析关键要点
    key_points = data.get("关键要点", [])
    if isinstance(key_points, str):
        key_points = [s.strip() for s in re.split(r"[\n;；]", key_points) if s.strip()]
    key_points = [str(s).strip() for s in key_points if str(s).strip()][:2]

    preview_text = text.strip().replace("\n", " ")
    if len(preview_text) > 160:
        preview_text = preview_text[:160] + "..."

    # 按段落分割并分析
    paragraphs = split_text_into_paragraphs(text)
    normalized_paragraphs = []
    for i, para in enumerate(paragraphs):
        para_preview = para.replace("\n", " ")
        if len(para_preview) > 160:
            para_preview = para_preview[:160] + "..."
        normalized_paragraphs.append({
            "段落文本": para_preview,
            "受访者自述身份": data.get("受访者自述身份", ""),
            "受访者类型": data.get("受访者类型", "普通用户"),
            "具体年龄": data.get("具体年龄", ""),
            "年龄段": data.get("年龄段", "未知"),
            "核心主题": data.get("核心主题", "日常数字生活"),
            "情绪": data.get("情绪", "中性"),
            "关键词": "、".join(keywords[:3]),
        })

    # 合并具体年龄和年龄段
    具体年龄 = data.get("具体年龄", "")
    年龄段 = data.get("年龄段", "未知")
    if 具体年龄 and 年龄段 and 年龄段 != "未知":
        年龄显示 = f"{具体年龄}（{年龄段}）"
    elif 具体年龄:
        年龄显示 = 具体年龄
    else:
        年龄显示 = 年龄段

    return {
        "访谈摘要": data.get("访谈摘要", "该访谈已完成分析。"),
        "受访者自述身份": data.get("受访者自述身份", ""),
        "受访者类型": data.get("受访者类型", "普通用户"),
        "年龄段": 年龄显示,
        "核心主题": data.get("核心主题", "日常数字生活"),
        "情绪": data.get("情绪", "中性"),
        "关键词": "、".join(keywords[:3]),
        "全部关键词": "、".join(keywords),
        "关键要点": "；".join(key_points) if key_points else "",
        "段落数": len(paragraphs),
        "分段分析": normalized_paragraphs,
    }


def analyze_interview(text: str) -> dict:
    """根据当前模式进行分析：AI深度分析（强制AI）或规则快速分析。"""
    global ai_status_message

    if st.session_state.get("analysis_mode") == "极速本地分析":
        ai_status_message = "当前为极速本地分析模式"
        return analyze_interview_rule_based(text)

    if client is None:
        ai_status_message = "未检测到 API Key，无法执行AI深度分析"
        raise RuntimeError("未配置 DeepSeek API Key，请在侧边栏填写后重试。")

    try:
        result = analyze_interview_ai(text)
        ai_status_message = "DeepSeek 调用成功，当前结果来自 AI 分析"
        return result
    except json.JSONDecodeError:
        ai_status_message = "DeepSeek 返回格式异常，已自动降级为规则分析"
        st.warning("AI 返回格式异常，已自动切换该条为规则分析以保证任务不中断。")
        return analyze_interview_rule_based(text)
    except Exception as e:
        ai_status_message = f"DeepSeek 调用失败：{str(e)}"
        st.error("AI 调用失败，请检查 API Key 或网络连接。当前结果未生成。")
        st.caption(f"错误详情：{type(e).__name__}: {str(e)}")
        st.caption(traceback.format_exc())
        raise


def split_bulk_interviews(text: str) -> List[str]:
    """将批量粘贴的文本拆分成多条访谈。支持编号或者空行拆分。"""
    raw_text = text.strip()
    if not raw_text:
        return []
    # 去除开头的标题行（如"InterviewPilot 展示测试文本｜新版"）
    raw_text = re.sub(r"^(InterviewPilot\s*展示[^：：]*[：:：]?\s*\n?)", "", raw_text)
    normalized = raw_text.replace("\r\n", "\n").replace("\r", "\n")
    lines = normalized.split("\n")

    # Pattern to detect numbered or lettered markers at line start
    marker_pattern = re.compile(
        r"^\s*(?:访谈\s*[0-9一二三四五六七八九十]+[：:：]?\s*|"
        r"[0-9]+[、.．)]|"
        r"[一二三四五六七八九十]+[、.]|"
        r"（[一二三四五六七八九十0-9]+）|"
        r"\([一二三四五六七八九十0-9]+\)|"
        r"[A-Za-z]+[、.．)])\s*"
    )

    has_markers = any(marker_pattern.match(line.strip()) for line in lines if line.strip())

    if has_markers:
        items = []
        current: List[str] = []
        for line in lines:
            stripped = line.strip()
            if not stripped:
                if current:
                    current.append("")
                continue
            if marker_pattern.match(stripped):
                if current:
                    items.append("\n".join(current).strip())
                    current = []
                stripped = marker_pattern.sub("", stripped, count=1).strip()
                if stripped:
                    current.append(stripped)
            else:
                current.append(line)
        if current:
            items.append("\n".join(current).strip())
        return [item for item in items if item]
    # No markers; split by blank lines
    chunks = re.split(r"\n\s*\n+", normalized)
    return [chunk.strip() for chunk in chunks if chunk.strip()]


# --------------------- 新增工具函数 ---------------------

def build_wordcloud_html(keyword_counts: dict, max_words: int = 40) -> str:
    """使用 HTML 生成轻量级词云，避免额外依赖。"""
    if not keyword_counts:
        return "<p style='color:#888;'>暂无关键词可展示。</p>"

    sorted_items = sorted(keyword_counts.items(), key=lambda x: x[1], reverse=True)[:max_words]
    counts = [count for _, count in sorted_items]
    min_count = min(counts)
    max_count = max(counts)

    def calc_font_size(count: int) -> int:
        if max_count == min_count:
            return 24
        return int(14 + (count - min_count) * (32 - 14) / (max_count - min_count))

    spans = []
    for word, count in sorted_items:
        font_size = calc_font_size(count)
        opacity = 0.65 if max_count == min_count else 0.55 + 0.45 * (count - min_count) / (max_count - min_count)
        spans.append(
            f"<span style='display:inline-block;margin:8px 10px;font-size:{font_size}px;font-weight:600;opacity:{opacity:.2f};'>"
            f"{word}</span>"
        )

    return (
        "<div style='padding:20px;border:1px solid #333;border-radius:12px;"
        "display:flex;flex-wrap:wrap;justify-content:center;align-items:center;"
        "gap:14px;background:rgba(255,255,255,0.02);'>"
        + "".join(spans)
        + "</div>"
    )


def build_excel_bytes(df: pd.DataFrame) -> bytes:
    """将结果导出为 Excel 二进制内容。"""
    base_cols = [
        "访谈ID",
        "访谈摘要",
        "核心主题",
        "情绪",
        "关键词",
        "关键要点",
        "段落数",
        "原始文本",
    ]
    extra_cols = ["受访者自述身份", "受访者类型", "年龄段", "全部关键词"]
    for col in extra_cols:
        if col in df.columns:
            base_cols.append(col)
    export_df = df[[c for c in base_cols if c in df.columns]].copy()

    detail_rows = []
    for _, row in df.iterrows():
        for idx, item in enumerate(row["分段分析"], start=1):
            detail_item = {
                "访谈ID": row["访谈ID"],
                "段落序号": idx,
                "核心主题": item.get("核心主题", ""),
                "情绪": item.get("情绪", ""),
                "关键词": item.get("关键词", ""),
                "段落内容": item.get("段落文本", ""),
            }
            for col in ["受访者自述身份", "受访者类型", "年龄段"]:
                if col in item:
                    detail_item[col] = item[col]
            detail_rows.append(detail_item)
    detail_df = pd.DataFrame(detail_rows)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="汇总结果")
        detail_df.to_excel(writer, index=False, sheet_name="分段分析")

        for sheet_name, sheet_df in {"汇总结果": export_df, "分段分析": detail_df}.items():
            worksheet = writer.sheets[sheet_name]
            for idx, col in enumerate(sheet_df.columns, start=1):
                max_len = max([len(str(col))] + [len(str(v)) for v in sheet_df[col].fillna("")])
                column_letter = worksheet.cell(row=1, column=idx).column_letter
                worksheet.column_dimensions[column_letter].width = min(max(max_len + 2, 12), 40)

    return output.getvalue()


def build_simple_report(df: pd.DataFrame, keyword_counts: dict) -> str:
    """生成简单访谈报告"""
    total = len(df)
    theme_counts = df["核心主题"].value_counts().to_dict()
    emotion_counts = df["情绪"].value_counts().to_dict()
    top_keywords = sorted(keyword_counts.items(), key=lambda x: x[1], reverse=True)[:10]

    lines = []
    lines.append("访谈分析报告")
    lines.append("=" * 30)
    lines.append(f"访谈数量：{total} 条")
    lines.append("")

    lines.append("一、主题分布")
    for k, v in theme_counts.items():
        lines.append(f"- {k}：{v}")

    lines.append("")
    lines.append("二、情绪分布")
    for k, v in emotion_counts.items():
        lines.append(f"- {k}：{v}")

    lines.append("")
    lines.append("三、高频关键词")
    for word, count in top_keywords:
        lines.append(f"- {word}（{count}次）")

    lines.append("")
    lines.append("四、访谈摘要")
    for _, row in df.iterrows():
        lines.append(f"[{row['访谈ID']}] {row['访谈摘要']}")

    lines.append("")
    lines.append("五、关键要点")
    for _, row in df.iterrows():
        key_points = str(row.get("关键要点", "")).strip()
        if key_points:
            lines.append(f"[{row['访谈ID']}] {key_points}")

    return "\n".join(lines)


def build_local_full_report(df: pd.DataFrame) -> str:
    """本地快速完整报告（无API也可生成）。"""
    total = len(df)
    theme_counts = df["核心主题"].value_counts().to_dict()
    emotion_counts = df["情绪"].value_counts().to_dict()

    top_themes = sorted(theme_counts.items(), key=lambda x: x[1], reverse=True)[:3]
    top_emotions = sorted(emotion_counts.items(), key=lambda x: x[1], reverse=True)[:2]

    all_points: list[str] = []
    for _, row in df.iterrows():
        txt = str(row.get("关键要点", ""))
        for p in [s.strip() for s in re.split(r"[；;\n]", txt) if s.strip()]:
            all_points.append(p)
    point_counts: dict[str, int] = {}
    for p in all_points:
        point_counts[p] = point_counts.get(p, 0) + 1
    top_points = [p for p, _ in sorted(point_counts.items(), key=lambda x: x[1], reverse=True)[:4]]

    lines = []
    lines.append("题目：访谈样本的数字服务体验快速研究报告")
    lines.append("")
    lines.append(f"摘要：基于{total}条访谈样本，本研究从主题分布、情绪倾向与关键问题出发进行归纳。结果显示，受访者主要关注数字流程可理解性、异常处理路径与线下兜底问题。")
    lines.append("关键词：" + "、".join([k for k, _ in top_themes if k][:5]))
    lines.append("")
    lines.append("一、主要发现")
    for i, (theme, cnt) in enumerate(top_themes, start=1):
        lines.append(f"{i}. {theme}是高频主题（{cnt}条）。")
    for p in top_points[:3]:
        lines.append(f"- {p}")
    lines.append("")
    lines.append("二、政策建议")
    lines.append("1. 统一高频流程提示语与错误说明，减少重复失败。")
    lines.append("2. 针对异常场景设置一键补救路径，降低用户恢复成本。")
    lines.append("3. 将办成率与异常恢复时长纳入绩效考核。")
    lines.append("")
    lines.append("三、局限与后续")
    lines.append("1. 样本规模有限，结论需结合后续扩样验证。")
    lines.append("2. 建议按群体（老年/工作人员/青年）做分层跟踪。")
    lines.append("")
    lines.append("附：情绪概览：" + "、".join([f"{k}({v})" for k, v in top_emotions]))
    return "\n".join(lines)


# --------- AI完整研究报告生成 ---------
def build_full_research_report(df: pd.DataFrame) -> str:
    """完整报告：优先短AI生成，失败自动回退本地快速报告。"""
    global ai_status_message

    if client is None:
        ai_status_message = "AI未连接，已使用本地快速完整报告"
        return build_local_full_report(df)

    materials = []
    for _, row in df.iterrows():
        summary = str(row.get("访谈摘要", ""))[:90]
        points = str(row.get("关键要点", ""))[:80]
        materials.append(
            f"{row['访谈ID']}\n"
            f"- 摘要：{summary}\n"
            f"- 主题：{row['核心主题']}\n"
            f"- 情绪：{row['情绪']}\n"
            f"- 关键词：{row['全部关键词']}\n"
            f"- 要点：{points}"
        )
    joined_materials = "\n\n".join(materials)

    prompt = f"""
你是一名社会科学研究助手。基于以下访谈材料，生成一份简洁研究报告。

访谈材料：
{joined_materials}

请输出结构：
题目
摘要（80字内）
关键词（5个内）
一、主要发现（3点）
二、政策建议（3点）
三、局限与后续（2点）

要求：
1. 每点1-2句，简洁直接。
2. 全文控制在550字以内。
3. 直接输出中文正文，不要JSON。
"""

    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是一名严谨的社会科学研究助手。"},
                {"role": "user", "content": prompt},
            ],
            temperature=0.2,
            max_tokens=420,
            timeout=12,
        )
        ai_status_message = "DeepSeek 报告生成成功"
        return response.choices[0].message.content.strip()
    except Exception as e:
        ai_status_message = f"DeepSeek 报告生成失败，已回退本地报告：{str(e)}"
        return build_local_full_report(df)
def build_word_report(report_text: str) -> bytes | None:
    """将研究报告导出为Word文档；若未安装 python-docx，则返回 None。"""
    if Document is None:
        return None

    doc = Document()
    for line in report_text.split("\n"):
        doc.add_paragraph(line)

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

# ----------- AI研究结论与建议生成 -----------
def generate_ai_research_insights(df: pd.DataFrame) -> dict:
    """使用 AI 根据整体访谈结果生成研究结论与建议。若 AI 不可用则返回空结果。"""
    if client is None:
        return {"research_findings": "", "research_suggestions": ""}

    try:
        materials = "\n".join(
            [
                f"{row['访谈ID']}：主题={row['核心主题']}；情绪={row['情绪']}；关键词={row['全部关键词']}；要点={row.get('关键要点', '')}"
                for _, row in df.iterrows()
            ]
        )

        prompt = f"""
你是一名社会科学研究助手。下面是多条访谈材料，请基于这些内容进行综合归纳。

访谈材料：
{materials}

请输出 JSON：
{{
"研究结论": ["", "", ""],
"研究建议": ["", "", ""]
}}

要求：
1. 研究结论总结整体趋势，例如技术使用困难、信任问题、数字鸿沟等。
2. 研究建议偏向政策或社会治理建议。
3. 每个列表 2-4 条，每条不超过25字。
"""

        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是一名社会科学研究分析助手。"},
                {"role": "user", "content": prompt},
            ],
            temperature=0.2,
            max_tokens=320,
            timeout=12,
        )

        content = response.choices[0].message.content
        data = parse_ai_json_response(content)

        findings = data.get("研究结论", [])
        suggestions = data.get("研究建议", [])

        findings_text = "\n".join([f"- {f}" for f in findings])
        suggestions_text = "\n".join([f"- {s}" for s in suggestions])

        return {
            "research_findings": findings_text,
            "research_suggestions": suggestions_text,
        }

    except Exception:
        return {"research_findings": "", "research_suggestions": ""}


def generate_ai_theme_brief(df: pd.DataFrame) -> str:
    """生成同主题下的AI对比洞察。"""
    if client is None:
        return ""
    try:
        rows = []
        for _, row in df.iterrows():
            rows.append(
                f"{row['访谈ID']}：主题={row['核心主题']}；情绪={row['情绪']}；关键词={row['关键词']}；要点={row.get('关键要点', '')}"
            )
        payload = "\n".join(rows)
        prompt = f"""
你是一名研究助手。请基于以下访谈结果输出简洁中文分析，结构固定为：
1. 主题共性（3条）
2. 关键差异（2条）
3. 优先改进抓手（3条）

每条不超过22字，使用项目符号。

数据：
{payload}
"""
        resp = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是简洁研究洞察助手。"},
                {"role": "user", "content": prompt},
            ],
            temperature=0.1,
            max_tokens=220,
            timeout=10,
        )
        return resp.choices[0].message.content.strip()
    except Exception:
        return ""


def add_pending_text() -> None:
    """添加单条访谈到待分析列表。"""
    current_text = st.session_state.input_text.strip()
    if not current_text:
        st.session_state.add_warning = "请先输入一段访谈文本。"
        return
    st.session_state.pending_texts.append(current_text)
    st.session_state.input_text = ""
    st.session_state.add_success = f"已添加第 {len(st.session_state.pending_texts)} 条待分析访谈。"
    log_usage_event("add_single")


def add_bulk_pending_texts() -> None:
    """添加多条访谈到待分析列表。"""
    bulk_text = st.session_state.bulk_input_text.strip()
    if not bulk_text:
        st.session_state.add_warning = "请先粘贴多条访谈文本。"
        return
    items = split_bulk_interviews(bulk_text)
    if not items:
        st.session_state.add_warning = "未识别到可添加的访谈内容。"
        return
    st.session_state.pending_texts.extend(items)
    st.session_state.bulk_input_text = ""
    st.session_state.add_success = f"已批量添加 {len(items)} 条待分析访谈。"
    log_usage_event("add_bulk", value=len(items))


if "pending_texts" not in st.session_state:
    st.session_state.pending_texts = []
if "records" not in st.session_state:
    st.session_state.records = []
if "input_text" not in st.session_state:
    st.session_state.input_text = ""
if "bulk_input_text" not in st.session_state:
    st.session_state.bulk_input_text = ""
if "add_warning" not in st.session_state:
    st.session_state.add_warning = ""
if "add_success" not in st.session_state:
    st.session_state.add_success = ""
if "precheck_results" not in st.session_state:
    st.session_state.precheck_results = []


# Sidebar
with st.sidebar:
    with st.expander("使用说明", expanded=False):
        st.write("1. 支持单条录入，也支持批量粘贴多条访谈文本。")
        st.write("2. 批量内容可按“访谈1 / 1. / 1、 / 一、 / （一）”等编号方式自动拆分。")
        st.write("3. 若未设置编号，系统将按空行自动识别并拆分长文本或多条访谈。")
        st.write("4. 点击“开始分析全部”后，系统将输出主题、情绪、关键词与摘要。")
        st.write("5. 结果区可查看统计图、词云、研究建议与报告导出。")
    st.markdown("---")
    st.selectbox(
        "分析模式",
        ["AI标准分析", "极速本地分析"],
        key="analysis_mode",
        help="AI标准分析用于生成差异化结果；极速本地分析用于快速低成本处理。",
    )
    st.checkbox(
        "使用个人API Key（可选）",
        key="use_personal_key",
        help="默认使用管理员配置的公共 Key；你也可以切换为自己的 Key。",
    )
    if st.session_state.use_personal_key:
        st.text_input(
            "DeepSeek API Key",
            key="deepseek_api_key",
            type="password",
            placeholder="sk-...",
        )
    client = get_ai_client()
    if client is not None:
        source = "个人Key" if st.session_state.use_personal_key else "公共Key"
        ai_status_message = f"已连接 DeepSeek（{source}）"
    else:
        key_for_check = get_effective_api_key()
        if key_for_check:
            try:
                key_for_check.encode("ascii")
                ai_status_message = "检测到Key但初始化失败，请检查网络或Key是否有效"
            except UnicodeEncodeError:
                ai_status_message = "Key格式错误：包含中文符号/全角字符，请重新粘贴纯英文ASCII Key"
        else:
            ai_status_message = "未检测到可用 API Key，当前不可进行AI深度分析"
    st.write(f"**AI状态：** {ai_status_message}")
    st.markdown("---")
    st.markdown("## 部署与管理")
    st.caption("建议：先完成环境变量配置，再执行部署检查。")
    st.caption(f"每日分析上限：{DAILY_ANALYZE_LIMIT} 条（可在代码中调整）")
    if st.button("运行部署前检查", use_container_width=True):
        st.session_state.precheck_results = run_deploy_precheck()
        log_usage_event("deploy_precheck")
    if st.session_state.precheck_results:
        precheck_df = pd.DataFrame(
            st.session_state.precheck_results,
            columns=["检查项", "状态", "说明"],
        )
        st.dataframe(precheck_df, use_container_width=True, hide_index=True)
    with st.expander("运营统计（简版）", expanded=False):
        usage = get_usage_summary(days=30)
        st.write(f"- 分析发起：{usage.get('analyze_start', 0)}")
        st.write(f"- 分析成功：{usage.get('analyze_success', 0)}")
        st.write(f"- 完整报告生成：{usage.get('gen_full_report', 0)}")
        st.write(f"- 主题洞察生成：{usage.get('gen_theme_brief', 0)}")


today_str = datetime.now().strftime("%Y-%m-%d")
if st.session_state.today_used_date != today_str:
    st.session_state.today_used_date = today_str
    st.session_state.today_used_local = 0

today_used_db = get_today_analyze_count(st.session_state.user_quota_id)
today_used = max(today_used_db, st.session_state.today_used_local)
remaining_quota = max(0, DAILY_ANALYZE_LIMIT - today_used)

kpi_cards = [
    ("待分析访谈数", str(len(st.session_state.pending_texts))),
    ("已完成分析数", str(len(st.session_state.records))),
    ("当前状态", "可开始分析" if st.session_state.pending_texts else "等待输入"),
    ("今日分析额度", f"{today_used}/{DAILY_ANALYZE_LIMIT}"),
]
kpi_html = "".join(
    (
        '<div class="ip-kpi-card">'
        f'<div class="ip-kpi-label">{escape(label)}</div>'
        f'<div class="ip-kpi-value">{escape(value)}</div>'
        "</div>"
    )
    for label, value in kpi_cards
)
st.markdown(
    f'<div class="ip-kpi-wrap"><div class="ip-kpi-grid">{kpi_html}</div></div>',
    unsafe_allow_html=True,
)

# Input areas
st.write("### 📝 单条添加")
st.text_area(
    "粘贴单条访谈文本：",
    key="input_text",
    placeholder="例如：我今年65岁了，平时不太会用智能手机，扫码支付经常要孩子帮忙。",
    height=140,
)
st.write("### 📚 批量添加")
st.text_area(
    "一次粘贴多条访谈（支持 访谈1 / 1. / 1、 / 一、 / （一） 等分隔方式；若无编号则按空行拆分）：",
    key="bulk_input_text",
    placeholder="例如：\n访谈1：我今年68岁了，平时不太会用智能手机。\n\n访谈2：最近去医院挂号也要在手机上预约。",
    height=220,
)

# Buttons
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.button("添加单条访谈", use_container_width=True, on_click=add_pending_text)
with col2:
    st.button("批量添加访谈", use_container_width=True, on_click=add_bulk_pending_texts)
with col3:
    analyze_clicked = st.button("开始分析全部", use_container_width=True)
with col4:
    clear_clicked = st.button("清空全部记录", use_container_width=True)

if st.session_state.add_warning:
    st.warning(st.session_state.add_warning)
    st.session_state.add_warning = ""
if st.session_state.add_success:
    st.success(st.session_state.add_success)
    st.session_state.add_success = ""

if analyze_clicked:
    if not st.session_state.pending_texts:
        st.warning("请先添加至少一条访谈，再开始分析。")
    elif len(st.session_state.pending_texts) > remaining_quota:
        st.error(
            f"已超过今日分析额度。当前剩余 {remaining_quota} 条，"
            f"请减少本次分析数量或明天再试。"
        )
    else:
        log_usage_event("analyze_start", value=len(st.session_state.pending_texts))
        st.session_state.records = []
        try:
            with st.spinner("正在分析访谈，请稍候..."):
                indexed_texts = list(enumerate(st.session_state.pending_texts, start=1))
                results_map: dict[int, dict] = {}
                max_workers = min(4, max(1, len(indexed_texts)))

                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    future_map = {
                        executor.submit(analyze_interview, interview_text): (idx, interview_text)
                        for idx, interview_text in indexed_texts
                    }
                    for future in as_completed(future_map):
                        idx, interview_text = future_map[future]
                        result = future.result()
                        results_map[idx] = {
                            "访谈ID": f"INT{idx:03d}",
                            "原始文本": interview_text,
                            "访谈摘要": result["访谈摘要"],
                            "受访者自述身份": result.get("受访者自述身份", ""),
                            "受访者类型": result["受访者类型"],
                            "年龄段": result["年龄段"],
                            "核心主题": result["核心主题"],
                            "情绪": result["情绪"],
                            "关键词": result["关键词"],
                            "全部关键词": result["全部关键词"],
                            "关键要点": result.get("关键要点", ""),
                            "段落数": result["段落数"],
                            "分段分析": result["分段分析"],
                        }

                st.session_state.records = [results_map[idx] for idx, _ in indexed_texts if idx in results_map]
            st.session_state.last_signature = records_signature(st.session_state.records)
            st.session_state.cached_ai_insights = {"research_findings": "", "research_suggestions": ""}
            st.session_state.cached_full_report = ""
            st.session_state.cached_theme_brief = ""
            st.success(f"已完成 {len(st.session_state.records)} 条访谈的统一分析。")
            st.session_state.today_used_local += len(st.session_state.records)
            log_usage_event("analyze_success", detail=st.session_state.user_quota_id, value=len(st.session_state.records))
        except Exception as e:
            st.session_state.records = []
            st.error(f"分析中止：{str(e)}")
            log_usage_event("analyze_fail")

if clear_clicked:
    st.session_state.pending_texts = []
    st.session_state.records = []
    st.success("已清空全部待分析访谈和分析结果。")

st.write("### 📌 待分析列表")
st.caption(
    f"当前已添加 {len(st.session_state.pending_texts)} 条访谈。你可以单条添加，也可以批量粘贴后统一加入，再点击“开始分析全部”。"
)

if st.session_state.pending_texts:
    def two_line_preview(text: str, line_len: int = 46) -> str:
        t = str(text).replace("\n", " ").strip()
        # 去除开头的标记词（包括标题）、冒号和空格
        t = re.sub(r"^(InterviewPilot\s*展示[^：：]*[：:：]?\s*|[^(]*访谈\s*[0-9一二三四五六七八九十]+[：:：]\s*[^\n]*\n?\s*)", "", t)
        if len(t) <= line_len:
            return t
        if len(t) <= line_len * 2:
            return t[:line_len] + "\n" + t[line_len:]
        return t[:line_len] + "\n" + t[line_len: line_len * 2] + "..."

    pending_df = pd.DataFrame(
        {
            "序号": [i + 1 for i in range(len(st.session_state.pending_texts))],
            "访谈文本预览": [
                two_line_preview(t) for t in st.session_state.pending_texts
            ],
        }
    )
    st.dataframe(
        pending_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "序号": st.column_config.NumberColumn(width="small"),
            "访谈文本预览": st.column_config.TextColumn(width="large"),
        },
        height=min(360, 80 + 48 * len(pending_df)),
    )
else:
    st.info("暂无待分析访谈。")

# Display results if available
if st.session_state.records:
    df = pd.DataFrame(st.session_state.records)
    current_sig = records_signature(st.session_state.records)
    if current_sig != st.session_state.last_signature:
        st.session_state.last_signature = current_sig
        st.session_state.cached_ai_insights = {"research_findings": "", "research_suggestions": ""}
        st.session_state.cached_full_report = ""
        st.session_state.cached_theme_brief = ""

    st.write("### 📋 访谈整理结果")
    # 显示列：优先显示受访者信息，再显示内容
    display_cols = ["访谈ID", "段落数", "受访者类型"]
    if "受访者自述身份" in df.columns:
        display_cols.append("受访者自述身份")
    display_cols.extend(["年龄段", "核心主题", "情绪", "关键词"])
    if "全部关键词" in df.columns:
        display_cols.append("全部关键词")
    if "关键要点" in df.columns:
        display_cols.append("关键要点")
    table_df = df[[c for c in display_cols if c in df.columns]].copy()
    st.dataframe(table_df, use_container_width=True, hide_index=True)

    st.write("### 🧾 访谈摘要")
    for _, row in df.iterrows():
        expander_label = f"{row['访谈ID']}｜{row['核心主题']}｜{row['受访者类型']}｜{row['段落数']}段"
        if row.get("受访者自述身份"):
            expander_label += f"｜{row['受访者自述身份']}"
        with st.expander(expander_label):
            st.write("**摘要：**")
            st.write(row["访谈摘要"])
            # 去除原始文本开头的标记词
            orig = row["原始文本"]
            orig = re.sub(r"^(访谈\s*[0-9一二三四五六七八九十]+[：:：]?\s*|INT[0-9]+\s*[：:：]?\s*|访谈[0-9]+[：:：]?\s*)", "", orig)
            st.write("**原始文本：**")
            st.write(orig)
            st.write("**关键词：**")
            st.write(row["关键词"])
            if row.get("关键要点"):
                st.write("**关键要点：**")
                st.write(row["关键要点"])
            st.write("**分段分析：**")
            for i, item in enumerate(row["分段分析"], start=1):
                st.markdown(f"**第{i}段**")
                st.write(f"- 核心主题：{item['核心主题']}")
                st.write(f"- 情绪：{item['情绪']}")
                st.write(f"- 关键词：{item['关键词']}")
                st.write(f"- 段落内容：{item['段落文本']}")

    st.write("### 📊 主题统计")
    theme_counts = (
        df["核心主题"]
        .value_counts()
        .rename_axis("核心主题")
        .reset_index(name="数量")
        .sort_values(["数量", "核心主题"], ascending=[False, True])
        .reset_index(drop=True)
    )
    theme_counts["排序"] = range(len(theme_counts))
    theme_chart = (
        alt.Chart(theme_counts)
        .mark_bar(size=28)
        .encode(
            x=alt.X("数量:Q", title="访谈数量", axis=alt.Axis(format="d", tickMinStep=1)),
            y=alt.Y("核心主题:N", sort=alt.SortField(field="排序", order="ascending"), title="主题"),
            tooltip=["核心主题", "数量"],
        )
        .properties(height=max(220, 60 * len(theme_counts)))
    )
    st.altair_chart(theme_chart, use_container_width=True)

    # Keyword frequency processing (use all keywords for better visualization)
    keyword_counts: dict = {}
    for _, row in df.iterrows():
        # 优先使用全部关键词，否则用关键词
        kws_text = row.get("全部关键词", row.get("关键词", ""))
        for kw in [item.strip() for item in str(kws_text).split("、") if item.strip()]:
            keyword_counts[kw] = keyword_counts.get(kw, 0) + 1

    keyword_df = pd.DataFrame(list(keyword_counts.items()), columns=["关键词", "数量"])
    if not keyword_df.empty:
        # 按数量降序排序，显示前20个
        keyword_df = (
            keyword_df.sort_values(["数量", "关键词"], ascending=[False, True])
            .head(20)
            .reset_index(drop=True)
        )
        keyword_df["排序"] = range(len(keyword_df))

    st.write("### 📈 关键词分布（出现次数≥1）")
    if keyword_df.empty:
        st.info("暂无关键词统计结果。")
    else:
        keyword_chart = (
            alt.Chart(keyword_df)
            .mark_bar(size=24)
            .encode(
                x=alt.X("数量:Q", title="出现次数", axis=alt.Axis(format="d", tickMinStep=1)),
                y=alt.Y("关键词:N", sort=alt.SortField(field="排序", order="ascending"), title="关键词"),
                tooltip=["关键词", "数量"],
            )
            .properties(height=max(320, 28 * len(keyword_df)))
        )
        st.altair_chart(keyword_chart, use_container_width=True)

    st.write("### ☁️ 关键词词云")
    st.markdown(build_wordcloud_html(keyword_counts), unsafe_allow_html=True)

    report_text = build_simple_report(df, keyword_counts)
    full_report_text = st.session_state.cached_full_report

    csv_df = df[[
        "访谈ID",
        "访谈摘要",
        "受访者类型",
        "年龄段",
        "核心主题",
        "情绪",
        "关键词",
        "全部关键词",
        "关键要点",
        "段落数",
        "原始文本",
    ]].copy()
    csv_bytes = csv_df.to_csv(index=False).encode("utf-8-sig")
    excel_bytes = build_excel_bytes(df)
    word_bytes = build_word_report(full_report_text if full_report_text else report_text)

    st.write("### ⚙️ AI扩展输出")
    st.info("完整研究报告已优化为“快速优先”：AI成功则输出AI版，失败或超时自动回退本地完整报告。")
    st.caption("温馨提示：如网络不稳定，完整报告会自动切换本地版本，不影响继续导出。")
    action_col1, action_col2, action_col3 = st.columns(3)
    with action_col1:
        gen_insight_clicked = st.button("生成AI研究结论/建议", use_container_width=True)
    with action_col2:
        gen_report_clicked = st.button("生成完整研究报告（快速优先）", use_container_width=True)
    with action_col3:
        gen_theme_brief_clicked = st.button("生成AI主题洞察", use_container_width=True)

    if gen_insight_clicked:
        log_usage_event("gen_ai_insights")
        with st.spinner("正在生成AI研究结论与建议..."):
            st.session_state.cached_ai_insights = generate_ai_research_insights(df)
    if gen_report_clicked:
        log_usage_event("gen_full_report")
        with st.spinner("正在生成完整研究报告..."):
            st.session_state.cached_full_report = build_full_research_report(df)
    if gen_theme_brief_clicked:
        log_usage_event("gen_theme_brief")
        with st.spinner("正在生成AI主题洞察..."):
            st.session_state.cached_theme_brief = generate_ai_theme_brief(df)

    ai_insights = st.session_state.cached_ai_insights
    full_report_text = st.session_state.cached_full_report
    theme_brief_text = st.session_state.cached_theme_brief

    st.write("### 🧠 AI结果概览")
    if ai_insights["research_findings"]:
        st.markdown("**AI研究结论**")
        st.markdown(ai_insights["research_findings"])
    else:
        st.caption("AI研究结论未生成。")

    if ai_insights["research_suggestions"]:
        st.markdown("**AI研究建议**")
        st.markdown(ai_insights["research_suggestions"])
    else:
        st.caption("AI研究建议未生成。")

    if theme_brief_text:
        st.markdown("**AI主题洞察**")
        st.markdown(theme_brief_text)
    else:
        st.caption("AI主题洞察未生成。")

    st.write("### 📄 访谈研究报告")
    report_mode = st.radio(
        "报告类型",
        ["简版报告", "完整研究报告", "AI主题洞察", "AI研究结论建议"],
        horizontal=True,
        index=0,
    )

    if report_mode == "完整研究报告":
        if not full_report_text:
            current_report_text = "尚未生成完整研究报告，请点击上方“生成完整研究报告”。"
        else:
            current_report_text = full_report_text
    elif report_mode == "AI主题洞察":
        current_report_text = theme_brief_text if theme_brief_text else "尚未生成AI主题洞察，请点击上方“生成AI主题洞察”。"
    elif report_mode == "AI研究结论建议":
        findings = ai_insights.get("research_findings", "").strip()
        suggestions = ai_insights.get("research_suggestions", "").strip()
        if findings or suggestions:
            current_report_text = "\n".join(
                [
                    "AI研究结论",
                    findings if findings else "- 暂无",
                    "",
                    "AI研究建议",
                    suggestions if suggestions else "- 暂无",
                ]
            )
        else:
            current_report_text = "尚未生成AI研究结论/建议，请点击上方“生成AI研究结论/建议”。"
    else:
        current_report_text = report_text
    st.markdown("**报告预览**")
    st.markdown(current_report_text.replace("\n", "  \n"))

    download_col1, download_col2, download_col3 = st.columns(3)
    with download_col1:
        st.download_button(
            label="下载CSV",
            data=csv_bytes,
            file_name="访谈整理结果汇总.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with download_col2:
        st.download_button(
            label="下载Excel",
            data=excel_bytes,
            file_name="访谈整理结果汇总.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with download_col3:
        if word_bytes is not None:
            st.download_button(
                label="下载Word研究报告",
                data=word_bytes,
                file_name="访谈调研报告.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        else:
            st.error("当前未安装 python-docx，暂时无法导出 Word 报告。请先在终端安装：pip3 install python-docx")
else:
    st.info("暂无分析结果。请先添加访谈，再点击“开始分析全部”。")

st.markdown(
    """
    <div class="footer-note">
        InterviewPilot ｜ 北京师范大学 ｜ Version 1.3.0 ｜ Created by 政管小白 ｜ 联系方式：yuqingsuen@163.com
    </div>
    """,
    unsafe_allow_html=True,
)
