# -*- coding: utf-8 -*-
"""
薪火相传 —— 红色文化 AIGC 教育创作平台（MVP 演示版）
========================================================
技术栈：Streamlit（界面）
       + DeepSeek 大模型（文本生成）
       + python-pptx（把生成的大纲打包成真正的 .pptx 文件）

两个功能：红色文章 / 红色PPT（纯文字，零图片成本）

运行方式：双击 run.bat，或在项目目录执行  streamlit run app.py

两种模式：
1. 演示模式（默认，无需 API Key）→ 用内置示例内容，开箱即用；
2. 真实模式 → 填 DeepSeek API Key，实时调用大模型生成。
"""

import streamlit as st
import requests
import json
import io
import re
from datetime import datetime

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn, nsdecls
from pptx.oxml import parse_xml

# ================= 页面基础设置 =================
st.set_page_config(
    page_title="薪火相传 · 红色文化AIGC创作平台",
    page_icon="🔥",
    layout="wide",
)

st.markdown(
    """
    <style>
        /* 全局背景与主容器 */
        .stApp { background: linear-gradient(180deg, #fdf5f3 0%, #ffffff 25%); }
        .block-container { padding-top: 1.5rem; max-width: 1150px; }

        /* 标题颜色 */
        h1, h2, h3 { color: #b71c1c; }

        /* 主按钮 */
        .stButton > button {
            background: linear-gradient(135deg, #c62828, #e53935);
            color: #fff; border: none; border-radius: 8px;
            font-weight: 600; padding: 0.5rem 1rem;
            transition: all 0.2s ease;
        }
        .stButton > button:hover {
            background: linear-gradient(135deg, #a61c1c, #c62828);
            box-shadow: 0 4px 14px rgba(198,40,40,.35);
            transform: translateY(-1px);
        }

        /* 下载按钮 */
        .stDownloadButton > button {
            border-radius: 8px; border: 1px solid #c62828;
            color: #c62828; font-weight: 600;
        }
        .stDownloadButton > button:hover { background: #fdecea; }

        /* 标签页 */
        .stTabs [data-baseweb="tab-list"] { gap: 0.5rem; }
        .stTabs [data-baseweb="tab"] {
            font-size: 1.05rem; font-weight: 600; color: #666;
            border-radius: 8px 8px 0 0;
        }
        .stTabs [aria-selected="true"] { color: #c62828 !important; }
        .stTabs [data-baseweb="tab-highlight"] { background-color: #c62828; }

        /* 侧边栏 */
        [data-testid="stSidebar"] {
            background: #fff5f3;
            border-right: 1px solid #f5d0cc;
        }

        /* 输入框 */
        .stTextInput input { border-radius: 8px; }
        .stTextInput input:focus { border-color: #c62828; box-shadow: 0 0 0 1px #c62828; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ================= DeepSeek 文本生成 =================
DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"


def call_deepseek(api_key, prompt, model="deepseek-chat", temperature=0.7, max_tokens=2048):
    """调用 DeepSeek 大模型（OpenAI 兼容 HTTP 接口）"""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    resp = requests.post(DEEPSEEK_URL, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"]


# ================= 功能一：红色文章 =================
def build_article_prompt(theme, article_type):
    """组装“生成红色文章”的提示词"""
    return (
        f"请围绕主题“{theme}”，撰写一篇{article_type}。\n"
        "要求：主题鲜明、情感真挚、弘扬革命精神与爱国主义，"
        "适合中小学生使用，语言流畅，字数600~800字。\n"
        "请直接输出正文。"
    )


def mock_article(theme, article_type):
    """演示模式：内置红色文章示例"""
    return (
        f"《{theme}》{article_type}\n\n"
        f"在中华民族的奋斗历程中，{theme}永远是一段激励人心的记忆。\n\n"
        "回顾历史，无数先辈以坚定的理想信念和无私的奉献，为我们今天的幸福生活铺就了道路。"
        "他们的精神告诉我们：今天的幸福来之不易，每一代人都有自己的使命与担当。\n\n"
        "作为新时代的青少年，我们应当铭记历史、缅怀先烈，把红色基因融入血脉，"
        "从小事做起、从现在做起，刻苦学习、砥砺品格，"
        "努力成长为担当民族复兴大任的时代新人。\n\n"
        "薪火相传，强国有我。让我们把对先辈的敬仰化作前行的力量，"
        "书写属于我们这一代人的精彩篇章。\n\n"
        "（演示模式：接入 DeepSeek 后可实时生成针对性更强的完整文章。）"
    )


# ================= 功能二：红色PPT =================
def build_ppt_prompt(theme, n, audience="中小学生"):
    """组装“生成红色教育PPT大纲”的提示词"""
    return (
        f"你是一位资深的红色教育课程设计专家，擅长为{audience}设计课件。\n"
        f"请围绕主题“{theme}”设计一份红色教育PPT大纲，要求如下：\n\n"
        f"一、内容要求\n"
        f"1. 逻辑递进：按“历史背景 → 核心事件与人物 → 精神内涵 → 时代意义 → 青少年如何传承”层层展开；\n"
        f"2. 尊重史实、严谨准确，弘扬革命精神与爱国主义，价值导向鲜明；\n"
        f"3. 语言生动有感染力，多用具体事例和细节，避免空洞说教；\n"
        f"4. 符合{audience}认知水平，通俗易懂、积极向上。\n\n"
        f"二、结构要求\n"
        f"1. 共 {n} 页（不含封面）；\n"
        f"2. 每页包含一个“标题”、若干“要点”，以及一段 60~120 字的“演讲备注”；\n"
        f"3. 标题凝练有冲击力（如“信仰的力量”），不超过 12 个字；\n"
        f"4. 每个要点一句话，不超过 22 个字，要具体可讲，不要写成“要点一”“首先其次”；\n"
        f"5. 演讲备注写这一页怎么讲、举什么例子，供教师照着讲，不要重复要点原文。\n\n"
        f"三、每页的版式（layout）\n"
        f"每页要根据这一页讲什么自己挑版式，不要全部用同一种：\n"
        f"1. statement：整页只有一句核心的话，bullets 只放 1 条，字数不超过 30；\n"
        f"2. list：3~5 个并列要点，最常用；\n"
        f"3. grid：3、4 或 6 个短概念并排，每条不超过 14 字；\n"
        f"4. timeline：按时间先后讲的事件（2~5 条），每条以年份开头，如“1934 中央红军开始长征”；\n"
        f"5. compare：正好两个方面的对比，bullets 只放 2 条；\n"
        f"6. stat：2~4 个关键数字，每条写成“数字｜说明”，例如“二万五千里｜红军长征总行程”；\n"
        f"7. section：只有标题的过渡页，bullets 为空数组。\n\n"
        f"四、输出格式\n"
        f"只输出 JSON 数组，不要任何其他文字，格式如下：\n"
        f'[{{"title":"页标题","layout":"list","bullets":["要点1","要点2","要点3"],"notes":"这一页怎么讲……"}}]'
    )


def mock_ppt_slides(theme):
    """演示模式：内置红色PPT大纲示例（带版式建议与演讲备注）"""
    return [
        {"title": f"走近{theme}", "layout": "list",
         "bullets": [f"认识{theme}的时代背景", "了解这段历史的重要意义", "想一想：它和我们今天有什么关系"],
         "notes": f"开场先不给结论。举起一张与{theme}相关的图片或地图，请学生说说自己看到、听到了什么，"
                  f"再引出今天的主题：我们一起走近{theme}，看看它为什么值得被反复讲述。"},
        {"title": "历史背景", "layout": "timeline",
         "bullets": ["旧中国内忧外患，民不聊生", "仁人志士前赴后继寻找出路",
                     "革命火种在艰难中点燃", "斗争形势出现重大转折"],
         "notes": "用时间轴讲清“为什么会发生”。提醒学生：不要用今天的眼光去苛责当时的人，"
                  "要理解他们在几乎看不到希望的时候仍然选择了前行。"},
        {"title": "核心事件与人物", "layout": "list",
         "bullets": ["关键事件的经过与转折", "革命先辈的坚定选择", "感人至深的英雄事迹"],
         "notes": "这一页挑一个最有画面感的细节讲透，比如一次会议、一封家书、一次突围。"
                  "细节讲完再回到人物：他们也是普通人，只是在该站出来的时候站了出来。"},
        {"title": "精神内涵", "layout": "grid",
         "bullets": ["坚定的理想信念", "无畏的牺牲奉献", "一心为民的初心"],
         "notes": "精神不要喊口号。每讲一条，就配一个具体故事或一句原话，"
                  "让学生自己说出“这和我们平时说的什么品质是相通的”。"},
        {"title": "时代意义", "layout": "compare",
         "bullets": ["红色基因是今天的精神底色，运动会上的坚持、志愿服务里的担当，都是它在现实中的样子",
                     "新时代青少年要接过接力棒，把个人成长放进国家和民族的需要里去想"],
         "notes": "把历史拉回到学生身边：今天的困难不再是雪山草地，但同样需要坚持、担当和团结。"
                  "可以请学生举一个自己身边的例子。"},
        {"title": "薪火相传，强国有我", "layout": "statement",
         "bullets": ["铭记历史、缅怀先烈，把红色基因一代代传下去"],
         "notes": "结尾留一个具体行动：给先烈写一封信、做一次红色地标探访、在班会上讲一个故事。"
                  "让学生带着任务离开课堂，红色教育才算真正落地。"},
    ]


def parse_ppt_json(text):
    """把大模型返回的内容解析成 PPT 大纲列表（容忍 markdown 代码块等杂质）"""
    text = text.strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:]
    start = text.find("[")
    end = text.rfind("]")
    if start != -1 and end != -1:
        text = text[start:end + 1]
    data = json.loads(text)
    if not isinstance(data, list):
        raise ValueError("PPT 内容不是列表格式")
    return data


# ================= PPT 设计系统（对照参考稿：咨询级排版） =================
# 画布与边距
PPT_W, PPT_H = Inches(13.333), Inches(7.5)
PPT_X0 = Inches(0.72)          # 左边距
PPT_W_BODY = Inches(11.9)      # 内容宽度
PPT_TITLE_Y = Inches(0.70)     # 标题基线
PPT_RULE_Y = Inches(1.38)      # 标题下细分隔线
PPT_FOOT_Y = Inches(7.02)      # 页脚

# 颜色
PPT_INK = RGBColor(0x10, 0x17, 0x2A)     # 标题 / 深色面板
PPT_INK2 = RGBColor(0x1B, 0x2A, 0x44)
PPT_BODY = RGBColor(0x4B, 0x55, 0x63)    # 正文
PPT_MUTED = RGBColor(0x8A, 0x93, 0xA3)   # 页脚 / 注释
PPT_LINE = RGBColor(0xE8, 0xEB, 0xF0)    # 分隔线 / 卡片描边
PPT_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PPT_RED = RGBColor(0xC0, 0x39, 0x2B)     # 主色
PPT_RED_T = RGBColor(0xFB, 0xE9, 0xE6)
PPT_GOLD = RGBColor(0xC9, 0x96, 0x3B)
PPT_GOLD_T = RGBColor(0xFB, 0xF1, 0xDD)
PPT_TEAL = RGBColor(0x2C, 0x8C, 0x82)
PPT_TEAL_T = RGBColor(0xE4, 0xF1, 0xEF)
PPT_DARK = RGBColor(0x0E, 0x16, 0x26)    # 深色页底
PPT_DEEP_RED = RGBColor(0x2A, 0x0A, 0x0A)  # 封面深红
PPT_FONT = "微软雅黑"

# 四色强调（红 / 金 / 青 / 墨），每张卡片轮换
PPT_ACCENTS = ((PPT_RED, PPT_RED_T), (PPT_GOLD, PPT_GOLD_T),
               (PPT_TEAL, PPT_TEAL_T), (PPT_INK2, RGBColor(0xE8, 0xEB, 0xF1)))

# 内容页可用版式
PPT_LAYOUTS = ("list", "grid", "timeline", "compare", "statement", "stat")
PPT_LAYOUT_LABELS = {"list": "列表", "grid": "网格", "timeline": "时间轴",
                     "compare": "对比", "statement": "观点", "stat": "数据", "section": "过渡页"}


def _ppt_fits(layout, bullets):
    n = len(bullets)
    if layout == "grid":
        return n in (3, 4, 6)
    if layout == "timeline":
        return 2 <= n <= 4
    if layout == "compare":
        return n == 2
    if layout == "statement":
        return n == 1
    if layout == "stat":
        return 2 <= n <= 4 and all(_ppt_stat_pair(b) for b in bullets)
    return 1 <= n <= 6


def _ppt_plan_layout(title, bullets, declared=None, prev=None):
    """按内容挑版式：模型给了就用模型的，没给/不合适就按内容判断，并避开与上一页重复。"""
    n = len(bullets)
    if n == 0:
        return "section"
    if declared in PPT_LAYOUTS and _ppt_fits(declared, bullets) and declared != prev:
        return declared
    years = sum(1 for b in bullets if re.search(r"(?:1[89]|20)\d{2}", b))
    order = []
    if 2 <= n <= 4 and all(_ppt_stat_pair(b) for b in bullets):
        order.append("stat")
    if years >= 2 or any(k in title for k in ("历程", "经过", "进程", "沿革", "路线", "大事记")):
        order.append("timeline")
    if n == 1:
        order.append("statement")
    if n == 2:
        order.append("compare")
    if n in (3, 4, 6):
        order.append("grid")
    if 2 <= n <= 4:
        order.append("timeline")
    order.append("list")
    for layout in order:
        if layout != prev:
            return layout
    return order[0]


def _ppt_stat_pair(bullet):
    """把“二万五千里｜红军行程”拆成数字与说明"""
    for sep in ("｜", "|"):
        if sep in bullet:
            left, right = [p.strip() for p in bullet.split(sep, 1)]
            if left and right:
                return left, right
    return None


# ---------- 基础绘制 ----------
def _ppt_font(run, size, color, bold=False, name=PPT_FONT, spc=None):
    f = run.font
    f.size, f.bold, f.name = Pt(size), bold, name
    f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    if rPr.find(qn("a:ea")) is None:
        rPr.append(parse_xml('<a:ea %s typeface="%s"/>' % (nsdecls("a"), name)))
    if spc is not None:  # 字间距（磅）
        rPr.set("spc", str(int(spc * 100)))


def _ppt_text(slide, x, y, w, h, text, size, color, bold=False,
              align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.15,
              font=PPT_FONT, spc=None):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, line in enumerate(str(text).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run()
        r.text = line
        _ppt_font(r, size, color, bold, font, spc)
    return box


def _ppt_shadow(sp, blur=0.10, dist=0.02, alpha=10):
    """柔和投影（参考稿的卡片都是这种淡投影）"""
    spPr = sp._element.spPr
    el = spPr.find(qn("a:effectLst"))
    if el is None:
        el = parse_xml('<a:effectLst %s/>' % nsdecls("a"))
        spPr.append(el)
    el.append(parse_xml(
        '<a:outerShdw %s blurRad="%d" dist="%d" dir="5400000" rotWithShape="0">'
        '<a:srgbClr val="000000"><a:alpha val="%d"/></a:srgbClr></a:outerShdw>'
        % (nsdecls("a"), int(blur * 914400), int(dist * 914400), int(alpha * 1000))))


def _ppt_shape(slide, kind, x, y, w, h, fill=None, line=None, line_w=0.75,
               shadow=False, rot=None):
    sp = slide.shapes.add_shape(kind, x, y, w, h)
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    if shadow:
        _ppt_shadow(sp)
    if rot:
        sp.rotation = rot
    return sp


def _ppt_gradient(sp, c1, c2, angle=45):
    sp.fill.gradient()
    stops = sp.fill.gradient_stops
    stops[0].color.rgb = c1
    stops[0].position = 0.0
    stops[1].color.rgb = c2
    stops[1].position = 1.0
    sp.fill.gradient_angle = angle


def _ppt_blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _ppt_rule(s, y, x=None, w=None, color=None):
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, x if x is not None else PPT_X0, y,
               w if w is not None else PPT_W_BODY, Pt(0.75), fill=color or PPT_LINE)


def _ppt_card(s, x, y, w, h, fill=None, line=None, shadow=True):
    return _ppt_shape(s, MSO_SHAPE.RECTANGLE, x, y, w, h,
                      fill=PPT_WHITE if fill is None else fill,
                      line=line, shadow=shadow)


def _ppt_chip(s, x, y, w, h, text, accent, tint=None, size=8.5, color=None,
              align=PP_ALIGN.CENTER, bold=True):
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, x, y, w, h, fill=tint or accent)
    _ppt_text(s, x, y, w, h, text, size, color or PPT_WHITE, bold=bold,
              align=align, anchor=MSO_ANCHOR.MIDDLE)


def _ppt_num_circle(s, x, y, d, num, accent, tint, size=15):
    _ppt_shape(s, MSO_SHAPE.OVAL, x, y, d, d, fill=tint)
    _ppt_text(s, x, y, d, d, "%02d" % num if isinstance(num, int) else str(num),
              size, accent, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def _ppt_dot(s, x, y, d, accent, tint):
    _ppt_shape(s, MSO_SHAPE.OVAL, x, y, d, d, fill=tint)
    pad = d * 0.25
    _ppt_shape(s, MSO_SHAPE.OVAL, x + pad, y + pad, d - pad * 2, d - pad * 2, fill=accent)


def _ppt_short(text, limit=22):
    text = text.strip()
    return text if len(text) <= limit else text[:limit - 1] + "…"


# ---------- 页框：眉标 + 标题 + 分隔线 + 页脚 ----------
def _ppt_head(s, theme, title, accent, subtitle=None, dark=False):
    tcol = PPT_WHITE if dark else PPT_INK
    scol = RGBColor(0xC9, 0xB4, 0xAE) if dark else PPT_BODY
    _ppt_text(s, PPT_X0, Inches(0.38), Inches(9.0), Inches(0.24),
              "红色教育课件  /  %s" % theme, 9.5, accent, bold=True, spc=0.8)
    _ppt_text(s, PPT_X0, PPT_TITLE_Y, Inches(11.5), Inches(0.56), title, 25, tcol, bold=True)
    _ppt_rule(s, PPT_RULE_Y, color=RGBColor(0x2C, 0x33, 0x44) if dark else PPT_LINE)
    if subtitle:
        _ppt_text(s, PPT_X0, Inches(1.62), Inches(10.8), Inches(0.4), subtitle,
                  13.5, scol)


def _ppt_foot(s, page, total, theme, dark=False):
    col = RGBColor(0x9A, 0xA6, 0xB8) if dark else PPT_MUTED
    _ppt_text(s, PPT_X0, PPT_FOOT_Y, Inches(6.4), Inches(0.22),
              "薪火相传 · 红色文化 AIGC 教育创作平台", 8.5, col)
    _ppt_text(s, Inches(11.8), PPT_FOOT_Y, Inches(0.82), Inches(0.22),
              "%02d" % page, 8.5, col, align=PP_ALIGN.RIGHT)


# ---------- 封面 ----------
def _cover_metric(s, x, y, total):
    """封面右侧卡片里的大数字 + 单位"""
    box = s.shapes.add_textbox(x, y, Inches(1.7), Inches(0.9))
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r1 = p.add_run()
    r1.text = str(total)
    _ppt_font(r1, 42, PPT_WHITE, True)
    r2 = p.add_run()
    r2.text = "  页"
    _ppt_font(r2, 17, RGBColor(0xD9, 0xB7, 0xAD))


def _ppt_cover(prs, theme, total):
    s = _ppt_blank(prs)
    bg = _ppt_shape(s, MSO_SHAPE.RECTANGLE, 0, 0, PPT_W, PPT_H)
    _ppt_gradient(bg, PPT_DEEP_RED, RGBColor(0x6E, 0x17, 0x12), 45)
    _ppt_shape(s, MSO_SHAPE.OVAL, Inches(9.85), Inches(-1.15), Inches(4.8), Inches(4.8),
               fill=RGBColor(0x3A, 0x12, 0x12))
    _ppt_shape(s, MSO_SHAPE.OVAL, Inches(11.15), Inches(-0.15), Inches(3.15), Inches(3.15),
               fill=RGBColor(0x45, 0x16, 0x16))
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.14), PPT_H, fill=PPT_GOLD)
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, PPT_X0, Inches(0.46), Inches(0.58), Inches(0.09),
               fill=PPT_RED)
    _ppt_text(s, PPT_X0, Inches(0.66), Inches(4.0), Inches(0.28),
              "红色教育课件  /  %s" % datetime.now().strftime("%Y.%m"), 10,
              RGBColor(0xF0, 0xC9, 0xB8), spc=0.8)
    _ppt_text(s, PPT_X0, Inches(1.22), Inches(7.6), Inches(0.8), theme, 38, PPT_WHITE, bold=True)
    _ppt_text(s, PPT_X0, Inches(2.12), Inches(7.2), Inches(0.55),
              "薪火相传 · 红色文化 AIGC 教育创作平台", 22, RGBColor(0xFF, 0xED, 0xE7))
    _ppt_text(s, PPT_X0, Inches(2.84), Inches(6.4), Inches(0.42),
              "把一个红色主题，变成一堂能直接上的课", 13.5, RGBColor(0xD9, 0xB7, 0xAD))
    _ppt_rule(s, Inches(3.48), w=Inches(4.8), color=RGBColor(0x5A, 0x2A, 0x26))
    _ppt_text(s, PPT_X0, Inches(3.70), Inches(4.0), Inches(0.28),
              "共 %d 页  /  约 %d 分钟" % (total, max(5, total * 1)), 10.5,
              RGBColor(0xC9, 0xA7, 0x9C))

    card = _ppt_shape(s, MSO_SHAPE.RECTANGLE, Inches(8.42), Inches(1.38), Inches(3.95),
                      Inches(4.60), fill=RGBColor(0x22, 0x0B, 0x0B), shadow=True)
    _ppt_text(s, Inches(8.78), Inches(1.72), Inches(2.4), Inches(0.25),
              "THEME  /  主题", 9, PPT_GOLD, bold=True, spc=0.8)
    _ppt_text(s, Inches(8.78), Inches(2.02), Inches(3.3), Inches(0.36),
              _ppt_short(theme, 12), 14, PPT_WHITE, bold=True)
    _cover_metric(s, Inches(8.78), Inches(2.55), total)
    _ppt_text(s, Inches(8.78), Inches(3.62), Inches(3.3), Inches(0.24),
              "AI 起草  /  教师审校", 9, RGBColor(0xC9, 0xA7, 0x9C))
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, Inches(8.78), Inches(4.10), Inches(3.22), Inches(0.12),
               fill=RGBColor(0x4A, 0x1C, 0x1A))
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, Inches(8.78), Inches(4.10), Inches(2.72), Inches(0.12),
               fill=PPT_GOLD)
    _ppt_text(s, Inches(8.78), Inches(4.40), Inches(3.1), Inches(0.24),
              "红色主题  ·  中小学思政", 9, RGBColor(0xC9, 0xA7, 0x9C))
    _ppt_chip(s, Inches(8.78), Inches(4.90), Inches(1.60), Inches(0.34), "课件草稿",
              PPT_GOLD, RGBColor(0x3A, 0x14, 0x12), size=9.5, color=RGBColor(0xF2, 0xD6, 0xB0))
    _ppt_chip(s, Inches(10.48), Inches(4.90), Inches(1.50), Inches(0.34), "人工复核",
              PPT_RED, RGBColor(0x5E, 0x14, 0x10), size=9.5)
    return s


# ---------- 目录 ----------
def _ppt_agenda(prs, titles, page, total, theme):
    s = _ppt_blank(prs)
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, 0, 0, PPT_W, PPT_H, fill=PPT_WHITE)
    _ppt_head(s, theme, "今天要讲的 %d 个部分" % len(titles), PPT_TEAL,
              "先看全局，再逐页展开。")
    n = len(titles)
    if n <= 6:
        cols, rows = 3, (n + 2) // 3
        gap = Inches(0.24)
        cw = (PPT_W_BODY - gap * 2) / 3
        ch = Inches(2.17)
        top = Inches(1.98) if rows == 2 else Inches(2.12)
        for i, t in enumerate(titles):
            c, r = i % cols, i // cols
            x = PPT_X0 + (cw + gap) * c
            y = top + (ch + gap) * r
            accent, tint = PPT_ACCENTS[i % 4]
            _ppt_card(s, x, y, cw, ch)
            _ppt_shape(s, MSO_SHAPE.RECTANGLE, x, y, cw, Inches(0.10), fill=accent)
            _ppt_num_circle(s, x + Inches(0.28), y + Inches(0.36), Inches(0.66),
                            i + 1, accent, tint)
            _ppt_text(s, x + Inches(1.14), y + Inches(0.38), cw - Inches(1.35),
                      Inches(0.36), _ppt_short(t, 12), 17, PPT_INK, bold=True)
            _ppt_text(s, x + Inches(0.28), y + Inches(1.55), cw - Inches(0.56),
                      Inches(0.28), "P.%02d" % (page + i), 9.5, accent, bold=True)
    else:
        rows = (n + 1) // 2
        row_h = min(Inches(0.92), Inches(4.9) / rows)
        for i, t in enumerate(titles):
            c, r = i // rows, i % rows
            x = PPT_X0 + Inches(6.1) * c
            y = Inches(1.85) + row_h * r
            accent, tint = PPT_ACCENTS[i % 4]
            if r > 0:
                _ppt_rule(s, y, x=x, w=Inches(5.5))
            _ppt_text(s, x, y, Inches(0.5), row_h, "%02d" % (i + 1), 12, accent, bold=True,
                      anchor=MSO_ANCHOR.MIDDLE)
            _ppt_text(s, x + Inches(0.52), y, Inches(4.9), row_h, _ppt_short(t, 16), 15,
                      PPT_INK, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    _ppt_foot(s, page, total, theme)
    return s


# ---------- 列表：一行一张卡片 ----------
def _ppt_body_list(s, bullets):
    n = len(bullets)
    top, bottom = Inches(1.95), Inches(6.80)
    gap = Inches(0.16)
    h = min(Inches(1.28), (bottom - top - gap * (n - 1)) // n)
    block = h * n + gap * (n - 1)
    y = top + (bottom - top - block) // 2
    fs = 14 if max(len(b) for b in bullets) <= 26 else 12.5
    for i, b in enumerate(bullets):
        accent, tint = PPT_ACCENTS[i % 4]
        _ppt_card(s, PPT_X0, y, PPT_W_BODY, h)
        _ppt_shape(s, MSO_SHAPE.RECTANGLE, PPT_X0, y, Inches(0.09), h, fill=accent)
        _ppt_num_circle(s, PPT_X0 + Inches(0.34), y + (h - Inches(0.46)) // 2,
                        Inches(0.46), i + 1, accent, tint, size=12)
        _ppt_text(s, PPT_X0 + Inches(1.02), y, Inches(10.3), h, b, fs, PPT_INK,
                  anchor=MSO_ANCHOR.MIDDLE, spacing=1.25)
        _ppt_text(s, PPT_X0 + Inches(11.35), y, Inches(0.45), h, "→", 13, PPT_MUTED,
                  anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
        y += h + gap


# ---------- 网格：并排卡片 ----------
def _ppt_body_grid(s, bullets):
    n = len(bullets)
    gap = Inches(0.24)
    if n == 4:  # 四张并排，中间带箭头（参考稿第 4 页）
        cw = (PPT_W_BODY - Inches(0.44) * 3 - Inches(0.28) * 3) // 4
        top, ch = Inches(2.30), Inches(2.55)
        for i, b in enumerate(bullets):
            x = PPT_X0 + (cw + Inches(0.28)) * i + Inches(0.14) * i
            accent, tint = PPT_ACCENTS[i % 4]
            _ppt_card(s, x, top, cw, ch)
            _ppt_shape(s, MSO_SHAPE.RECTANGLE, x, top, cw, Inches(0.10), fill=accent)
            _ppt_num_circle(s, x + Inches(0.26), top + (ch - Inches(0.66)) // 2, Inches(0.66),
                            i + 1, accent, tint)
            _ppt_text(s, x + Inches(1.10), top + Inches(0.12), cw - Inches(1.30),
                      ch - Inches(0.24), b, 15, PPT_INK, bold=True,
                      anchor=MSO_ANCHOR.MIDDLE, spacing=1.3)
            if i < n - 1:
                _ppt_text(s, x + cw + Inches(0.04), top + Inches(0.75), Inches(0.24),
                          Inches(0.3), "→", 15, PPT_MUTED)
    else:
        cols = 3 if n in (3, 6) else 2
        rows = (n + cols - 1) // cols
        cw = (PPT_W_BODY - gap * (cols - 1)) // cols
        ch = Inches(2.05) if rows == 2 else Inches(2.55)
        top = Inches(1.95) if rows == 2 else Inches(2.30)
        for i, b in enumerate(bullets):
            c, r = i % cols, i // cols
            x = PPT_X0 + (cw + gap) * c
            y = top + (ch + Inches(0.24)) * r
            accent, tint = PPT_ACCENTS[i % 4]
            _ppt_card(s, x, y, cw, ch)
            _ppt_shape(s, MSO_SHAPE.RECTANGLE, x, y, cw, Inches(0.10), fill=accent)
            _ppt_num_circle(s, x + Inches(0.30), y + (ch - Inches(0.66)) // 2, Inches(0.66),
                            i + 1, accent, tint)
            _ppt_text(s, x + Inches(1.18), y + Inches(0.12), cw - Inches(1.42),
                      ch - Inches(0.24), b, 17, PPT_INK, bold=True,
                      anchor=MSO_ANCHOR.MIDDLE, spacing=1.3)


# ---------- 时间轴：标签 + 节点 + 说明卡 ----------
def _ppt_body_timeline(s, bullets):
    n = len(bullets)
    gap = Inches(0.30)
    cw = (PPT_W_BODY - gap * (n - 1)) // n
    chip_y, node_y, line_y, card_y = Inches(2.15), Inches(3.15), Inches(3.53), Inches(4.12)
    card_h = Inches(1.45)
    has_years = any(re.search(r"(?:1[89]|20)\d{2}", b) for b in bullets)
    if not has_years:  # 没有年份就不做标签行，节点整体上移
        node_y, line_y, card_y = Inches(2.62), Inches(3.00), Inches(3.60)
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, PPT_X0, line_y, PPT_W_BODY, Inches(0.04),
               fill=PPT_LINE)
    for i, b in enumerate(bullets):
        x = PPT_X0 + (cw + gap) * i
        accent, tint = PPT_ACCENTS[i % 4]
        m = re.search(r"((?:1[89]|20)\d{2})", b)
        label = m.group(1) if m else "%02d" % (i + 1)
        text = b[m.end():].lstrip("年 ，,、·") if m else b
        if m:
            _ppt_chip(s, x, chip_y, cw, Inches(0.52), label, accent, tint, size=15,
                      color=accent)
        _ppt_shape(s, MSO_SHAPE.OVAL, x + cw // 2 - Inches(0.38), node_y, Inches(0.76),
                   Inches(0.76), fill=accent, shadow=True)
        _ppt_text(s, x + cw // 2 - Inches(0.38), node_y, Inches(0.76), Inches(0.76),
                  "%02d" % (i + 1), 14, PPT_WHITE, bold=True,
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        if i < n - 1:
            _ppt_text(s, x + cw, node_y + Inches(0.18), gap, Inches(0.4), "→", 20,
                      PPT_MUTED, align=PP_ALIGN.CENTER)
        _ppt_card(s, x, card_y, cw, card_h)
        _ppt_text(s, x + Inches(0.22), card_y + Inches(0.10), cw - Inches(0.44),
                  card_h - Inches(0.20), text, 11, PPT_BODY, spacing=1.35,
                  anchor=MSO_ANCHOR.MIDDLE)


# ---------- 对比：深色面板 + 浅色面板 ----------
def _ppt_body_compare(s, bullets):
    gap = Inches(0.30)
    pw = (PPT_W_BODY - gap) // 2
    top, ph = Inches(2.02), Inches(4.30)
    for i, b in enumerate(bullets):
        x = PPT_X0 + (pw + gap) * i
        accent = PPT_GOLD if i == 0 else PPT_RED
        if i == 0:
            _ppt_shape(s, MSO_SHAPE.RECTANGLE, x, top, pw, ph, fill=PPT_INK, shadow=True)
            _ppt_shape(s, MSO_SHAPE.RECTANGLE, x, top, pw, Inches(0.08), fill=accent)
            _ppt_text(s, x + Inches(0.40), top + Inches(0.42), Inches(2.0), Inches(0.4),
                      "%02d" % (i + 1), 26, accent, bold=True)
            _ppt_text(s, x + Inches(0.40), top + Inches(1.05), pw - Inches(0.8),
                      ph - Inches(1.45), b, 16, PPT_WHITE, spacing=1.45,
                      anchor=MSO_ANCHOR.MIDDLE)
        else:
            _ppt_card(s, x, top, pw, ph)
            _ppt_shape(s, MSO_SHAPE.RECTANGLE, x, top, pw, Inches(0.08), fill=accent)
            _ppt_text(s, x + Inches(0.40), top + Inches(0.42), Inches(2.0), Inches(0.4),
                      "%02d" % (i + 1), 26, accent, bold=True)
            _ppt_text(s, x + Inches(0.40), top + Inches(1.05), pw - Inches(0.8),
                      ph - Inches(1.45), b, 16, PPT_BODY, spacing=1.45,
                      anchor=MSO_ANCHOR.MIDDLE)
    _ppt_shape(s, MSO_SHAPE.OVAL, Inches(6.38), Inches(3.55), Inches(0.62), Inches(0.62),
               fill=PPT_GOLD, shadow=True)
    _ppt_text(s, Inches(6.38), Inches(3.55), Inches(0.62), Inches(0.62), "VS", 10,
              PPT_INK, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ---------- 观点：整页一句话 ----------
def _ppt_statement(prs, page, total, theme, title, bullets):
    s = _ppt_blank(prs)
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, 0, 0, PPT_W, PPT_H, fill=PPT_DARK)
    _ppt_shape(s, MSO_SHAPE.OVAL, Inches(10.2), Inches(4.4), Inches(4.6), Inches(4.6),
               fill=RGBColor(0x14, 0x1E, 0x31))
    _ppt_text(s, PPT_X0, Inches(0.38), Inches(9.0), Inches(0.24),
              "观点  /  %s" % theme, 9.5, PPT_GOLD, bold=True, spc=0.8)
    _ppt_text(s, PPT_X0, Inches(1.55), Inches(9.6), Inches(0.4),
              _ppt_short(title, 24), 15, RGBColor(0xC9, 0xB4, 0xAE))
    text = bullets[0]
    fs = 34 if len(text) <= 20 else 28 if len(text) <= 32 else 23
    _ppt_text(s, PPT_X0, Inches(2.30), Inches(10.6), Inches(2.1), "“" + text + "”", fs,
              PPT_WHITE, bold=True, spacing=1.5)
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, PPT_X0, Inches(4.78), Inches(1.30), Inches(0.05),
               fill=PPT_GOLD)
    _ppt_text(s, PPT_X0, Inches(5.02), Inches(8.0), Inches(0.3),
              "薪火相传  ·  红色文化 AIGC 教育创作平台", 10, RGBColor(0xC9, 0xA7, 0x9C))
    _ppt_foot(s, page, total, theme, dark=True)
    return s


# ---------- 数据：KPI 卡片 ----------
def _ppt_body_stat(s, bullets):
    pairs = [p for p in (_ppt_stat_pair(b) for b in bullets) if p]
    n = len(pairs)
    gap = Inches(0.24)
    cw = (PPT_W_BODY - gap * (n - 1)) // n
    top, ch = Inches(2.45), Inches(2.20)
    longest = max(len(num) for num, _ in pairs)
    avail_pt = (cw - Inches(0.95)) / 12700.0
    fs = max(22, min(40, int(avail_pt * 0.92 / max(1, longest))))
    for i, (num, label) in enumerate(pairs):
        x = PPT_X0 + (cw + gap) * i
        accent, tint = PPT_ACCENTS[i % 4]
        _ppt_card(s, x, top, cw, ch)
        _ppt_shape(s, MSO_SHAPE.RECTANGLE, x, top, Inches(0.09), ch, fill=accent)
        _ppt_text(s, x + Inches(0.34), top + Inches(0.26), cw - Inches(0.95),
                  Inches(0.7), num, fs, PPT_INK, bold=True)
        _ppt_text(s, x + Inches(0.34), top + Inches(1.28), cw - Inches(0.6),
                  Inches(0.5), label, 10.5, PPT_BODY, spacing=1.3)
        _ppt_dot(s, x + cw - Inches(0.44), top + Inches(0.20), Inches(0.22), accent, tint)


# ---------- 过渡页 ----------
def _ppt_section(prs, page, total, theme, title):
    s = _ppt_blank(prs)
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, 0, 0, PPT_W, PPT_H, fill=PPT_INK)
    _ppt_shape(s, MSO_SHAPE.OVAL, Inches(9.6), Inches(-1.4), Inches(5.4), Inches(5.4),
               fill=PPT_INK2)
    _ppt_text(s, Inches(9.2), Inches(3.0), Inches(3.3), Inches(2.6), "%02d" % page,
              160, RGBColor(0x22, 0x2E, 0x44), bold=True, align=PP_ALIGN.RIGHT)
    _ppt_text(s, PPT_X0, Inches(2.40), Inches(3.0), Inches(0.3), "SECTION", 10, PPT_GOLD,
              bold=True, spc=1.2)
    _ppt_text(s, PPT_X0, Inches(2.78), Inches(9.0), Inches(1.0), title, 40, PPT_WHITE, bold=True)
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, PPT_X0, Inches(4.02), Inches(1.60), Inches(0.05),
               fill=PPT_RED)
    _ppt_foot(s, page, total, theme, dark=True)
    return s


# ---------- 结束页 ----------
def _ppt_end(prs, theme, total):
    s = _ppt_blank(prs)
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, 0, 0, PPT_W, PPT_H, fill=PPT_DARK)
    _ppt_shape(s, MSO_SHAPE.OVAL, Inches(10.35), Inches(4.0), Inches(4.5), Inches(4.5),
               fill=RGBColor(0x15, 0x1F, 0x33))
    _ppt_shape(s, MSO_SHAPE.OVAL, Inches(11.3), Inches(4.9), Inches(2.7), Inches(2.7),
               fill=RGBColor(0x1B, 0x28, 0x40))
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, 0, Inches(7.36), PPT_W, Inches(0.14), fill=PPT_RED)
    _ppt_text(s, PPT_X0, Inches(0.75), Inches(4.0), Inches(0.3), "THANK YOU  /  Q&A", 10,
              PPT_GOLD, bold=True, spc=1.0)
    _ppt_text(s, PPT_X0, Inches(1.42), Inches(9.4), Inches(0.95), "谢谢观看", 38, PPT_WHITE,
              bold=True)
    _ppt_text(s, PPT_X0, Inches(2.48), Inches(8.6), Inches(0.45),
              "薪火相传，强国有我 · 把红色基因一代代传下去", 14,
              RGBColor(0xC9, 0xB4, 0xAE))
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, PPT_X0, Inches(3.30), Inches(6.36), Inches(0.62),
               fill=RGBColor(0x16, 0x20, 0x34))
    _ppt_text(s, Inches(0.97), Inches(3.30), Inches(0.6), Inches(0.62), "01", 12, PPT_GOLD,
              bold=True, anchor=MSO_ANCHOR.MIDDLE)
    _ppt_text(s, Inches(1.58), Inches(3.30), Inches(1.5), Inches(0.62), "回望主题", 12,
              PPT_WHITE, anchor=MSO_ANCHOR.MIDDLE)
    _ppt_text(s, Inches(3.05), Inches(3.30), Inches(3.9), Inches(0.62),
              _ppt_short(theme, 12) + " · 关键事件与精神", 10, PPT_MUTED,
              anchor=MSO_ANCHOR.MIDDLE)
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, PPT_X0, Inches(4.08), Inches(6.36), Inches(0.62),
               fill=RGBColor(0x16, 0x20, 0x34))
    _ppt_text(s, Inches(0.97), Inches(4.08), Inches(0.6), Inches(0.62), "02", 12, PPT_GOLD,
              bold=True, anchor=MSO_ANCHOR.MIDDLE)
    _ppt_text(s, Inches(1.58), Inches(4.08), Inches(1.5), Inches(0.62), "回到课堂", 12,
              PPT_WHITE, anchor=MSO_ANCHOR.MIDDLE)
    _ppt_text(s, Inches(3.05), Inches(4.08), Inches(3.9), Inches(0.62),
              "用一个问题、一个故事开始", 10, PPT_MUTED, anchor=MSO_ANCHOR.MIDDLE)
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, PPT_X0, Inches(4.86), Inches(6.36), Inches(0.62),
               fill=RGBColor(0x16, 0x20, 0x34))
    _ppt_text(s, Inches(0.97), Inches(4.86), Inches(0.6), Inches(0.62), "03", 12, PPT_GOLD,
              bold=True, anchor=MSO_ANCHOR.MIDDLE)
    _ppt_text(s, Inches(1.58), Inches(4.86), Inches(1.5), Inches(0.62), "留下行动", 12,
              PPT_WHITE, anchor=MSO_ANCHOR.MIDDLE)
    _ppt_text(s, Inches(3.05), Inches(4.86), Inches(3.9), Inches(0.62),
              "写一封信、走一处地标、讲一个故事", 10, PPT_MUTED, anchor=MSO_ANCHOR.MIDDLE)
    _ppt_foot(s, total, total, theme, dark=True)
    return s


# ---------- 内容页分发 ----------
def _ppt_content(prs, page, total, theme, title, bullets, notes=None, layout="list"):
    if layout == "section" or not bullets:
        s = _ppt_section(prs, page, total, theme, title)
        if notes:
            s.notes_slide.notes_text_frame.text = notes
        return
    if layout == "statement":
        s = _ppt_statement(prs, page, total, theme, title, bullets)
        if notes:
            s.notes_slide.notes_text_frame.text = notes
        return
    s = _ppt_blank(prs)
    _ppt_shape(s, MSO_SHAPE.RECTANGLE, 0, 0, PPT_W, PPT_H, fill=PPT_WHITE)
    accent = PPT_ACCENTS[(page - 1) % 4][0]
    subtitle = None
    _ppt_head(s, theme, title, accent, subtitle)
    if layout == "grid":
        _ppt_body_grid(s, bullets)
    elif layout == "timeline":
        _ppt_body_timeline(s, bullets)
    elif layout == "compare":
        _ppt_body_compare(s, bullets)
    elif layout == "stat":
        _ppt_body_stat(s, bullets)
    else:
        _ppt_body_list(s, bullets)
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    _ppt_foot(s, page, total, theme)
    return s


def build_pptx(theme, slides):
    """把“标题 + 要点 + 备注”的大纲排成一份 16:9 课件，返回 BytesIO。
    每页版式由内容决定（也可由大纲里的 layout 字段指定）。"""
    prs = Presentation()
    prs.slide_width, prs.slide_height = PPT_W, PPT_H

    items = []
    for s in slides:
        if isinstance(s, dict):
            title = str(s.get("title") or "").strip() or "未命名"
            bullets = [str(b).strip() for b in (s.get("bullets") or []) if str(b).strip()]
            notes = str(s.get("notes") or "").strip()
            layout = str(s.get("layout") or "").strip()
        else:
            title, bullets, notes, layout = str(s).strip() or "未命名", [], "", ""
        items.append((title, bullets, notes, layout))

    with_toc = len(items) >= 3
    total = 1 + (1 if with_toc else 0) + len(items) + 1
    _ppt_cover(prs, theme, total)
    page = 2
    if with_toc:
        _ppt_agenda(prs, [t for t, _, _, _ in items], page, total, theme)
        page += 1
    prev = None
    for title, bullets, notes, layout in items:
        layout = _ppt_plan_layout(title, bullets, layout, prev)
        prev = layout
        _ppt_content(prs, page, total, theme, title, bullets, notes, layout)
        page += 1
    _ppt_end(prs, theme, total)
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
# ================= 内置红色史料库（演示数据） =================
RED_ARCHIVES = [
    {"title": "井冈山会师", "year": "1928", "category": "土地革命", "figure": "毛泽东、朱德",
     "summary": "毛泽东率秋收起义部队与朱德、陈毅率领的南昌起义余部在井冈山会师，创建了第一个农村革命根据地，开辟了农村包围城市的革命道路。"},
    {"title": "遵义会议", "year": "1935", "category": "长征", "figure": "党中央",
     "summary": "长征途中召开的重要会议，确立了毛泽东在党中央和红军的领导地位，是党的历史上生死攸关的转折点。"},
    {"title": "红军长征", "year": "1934—1936", "category": "长征", "figure": "中国工农红军",
     "summary": "红军克服雪山草地等难以想象的艰难险阻，行程二万五千里，锤炼了革命队伍，铸就了伟大的长征精神。"},
    {"title": "方志敏", "year": "1935", "category": "革命烈士", "figure": "方志敏",
     "summary": "无产阶级革命家，在狱中写下《可爱的中国》，被捕后坚贞不屈，英勇就义。"},
    {"title": "杨靖宇", "year": "1940", "category": "抗日英烈", "figure": "杨靖宇",
     "summary": "东北抗日联军主要领导人，孤身与日军周旋数昼夜后壮烈牺牲，被誉为民族脊梁。"},
    {"title": "焦裕禄", "year": "1962—1964", "category": "时代楷模", "figure": "焦裕禄",
     "summary": "兰考县委书记，带领群众治理风沙、内涝、盐碱，积劳成疾仍坚持工作，被誉为“县委书记的榜样”。"},
    {"title": "雷锋", "year": "1960年代", "category": "时代楷模", "figure": "雷锋",
     "summary": "全心全意为人民服务的楷模，“雷锋精神”成为一代代中国人无私奉献的精神坐标。"},
    {"title": "抗美援朝", "year": "1950—1953", "category": "保家卫国", "figure": "中国人民志愿军",
     "summary": "志愿军出国作战、保家卫国，涌现出黄继光、邱少云等英雄，打出了新中国的国威军威。"},
]


def build_archive_prompt(keyword):
    """组装“AI 联网检索红色资料”的提示词"""
    return (
        f"你是红色文化研究专家。请为关键词“{keyword}”生成一份红色教育资源卡片，"
        "内容尊重史实、准确、积极向上，适合中小学生思政教育使用。\n\n"
        "请按以下结构输出：\n"
        "【名称】\n"
        "【年代】\n"
        "【相关人物】\n"
        "【历史背景】（2-3句）\n"
        "【核心事迹/内容】（3-5句）\n"
        "【精神内涵与时代意义】（2-3句）\n\n"
        "请直接输出，不要其他说明。"
    )


def mock_archive(keyword):
    """演示模式：内置的联网检索示例"""
    return (
        f"【名称】{keyword}\n"
        f"【年代】——\n"
        f"【相关人物】——\n"
        f"【历史背景】{keyword}是红色文化中的重要内容，承载着革命先辈的奋斗记忆。\n"
        f"【核心事迹/内容】无数先辈在相关历史进程中坚定信念、无私奉献，"
        f"谱写了感人的篇章，值得我们永远铭记。\n"
        f"【精神内涵与时代意义】它体现的坚定理想与爱国情怀，"
        f"激励着新时代青少年传承红色基因、砥砺前行。\n\n"
        f"（演示模式：接入 DeepSeek 后可实时检索并生成针对性更强的详细资料。）"
    )


# ================= 侧边栏：模式与 API 配置 =================
with st.sidebar:
    st.title("🔥 薪火相传")
    st.caption("红色文化 AIGC 教育创作平台")

    st.markdown("---")
    mode = st.radio(
        "运行模式",
        ["演示模式（无需 API，开箱即用）", "真实模式（调用 DeepSeek）"],
    )
    is_real = mode.startswith("真实")

    api_key = ""
    if is_real:
        api_key = st.text_input(
            "DeepSeek API Key（sk- 开头）",
            type="password",
            help="在 DeepSeek 开放平台（platform.deepseek.com）创建 API Key 后粘贴。",
        )

    st.markdown("---")
    st.caption("参赛方向：AI+教育 · 红色教育创新与文化传承")
    st.caption("作品形态：云服务（Web 应用）")


# ================= 主界面标题 =================
st.markdown(
    """
    <div style="background: linear-gradient(135deg, #b71c1c 0%, #d32f2f 55%, #e53935 100%);
                padding: 2rem 2.5rem; border-radius: 16px; margin-bottom: 0.5rem; color: #fff;
                box-shadow: 0 6px 20px rgba(183,28,28,.25);">
        <div style="font-size: 1.9rem; font-weight: 800; letter-spacing: 1px;">🔥 薪火相传</div>
        <div style="font-size: 1.05rem; margin-top: 0.4rem; opacity: .95;">红色文化 AIGC 教育创作平台</div>
        <div style="font-size: 0.88rem; margin-top: 0.6rem; opacity: .85;">
            输入红色主题，一键生成 <b>红色文章</b> · <b>红色PPT</b> ｜ 检索红色史料 ｜ AI 讲解红色历史
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

tab_create, tab_archive, tab_guide = st.tabs(
    ["📝 智能创作", "📚 红色资源库", "💬 AI 红色讲解员"]
)


# ================= Tab 1：智能创作 =================
with tab_create:
    st.subheader("智能创作")
    col1, col2 = st.columns([2, 3])
    with col1:
        work_type = st.radio("作品类型", ["红色文章", "红色PPT", "PPT提示词"])
        theme = st.text_input("红色主题", placeholder="例如：长征、井冈山、焦裕禄、抗美援朝……")

        article_type = "演讲稿"
        if work_type == "红色文章":
            article_type = st.selectbox("文章类型", ["演讲稿", "读后感", "作文", "新闻稿"])

        n_slides = 8
        if work_type in ["红色PPT", "PPT提示词"]:
            n_slides = st.slider("PPT页数（不含封面）", 5, 15, 8)

        audience = "中小学生"
        if work_type == "PPT提示词":
            audience = st.selectbox("目标受众", ["小学生", "初中生", "高中生", "大学生"])

        generate = st.button("开始生成", type="primary", use_container_width=True)

    with col2:
        # ----- 红色文章 -----
        if work_type == "红色文章":
            if generate and theme.strip():
                with st.spinner("正在撰写，请稍候……"):
                    if is_real and api_key.strip():
                        try:
                            result = call_deepseek(api_key, build_article_prompt(theme.strip(), article_type))
                        except Exception as e:
                            st.warning(f"调用 DeepSeek 失败（{e}），已自动回退到演示模式。")
                            result = mock_article(theme.strip(), article_type)
                    else:
                        result = mock_article(theme.strip(), article_type)
                st.session_state["article_result"] = result

            if "article_result" in st.session_state:
                st.text_area("生成的文章", st.session_state["article_result"], height=420)
                st.download_button(
                    "下载文章（.txt）",
                    st.session_state["article_result"],
                    file_name=f"{theme.strip() or '红色文章'}_{article_type}.txt",
                )
            else:
                st.info("在左侧输入主题，点击“开始生成”。")

        # ----- 红色PPT -----
        elif work_type == "红色PPT":
            if generate and theme.strip():
                with st.spinner("正在生成PPT大纲，请稍候……"):
                    if is_real and api_key.strip():
                        try:
                            raw = call_deepseek(api_key, build_ppt_prompt(theme.strip(), n_slides))
                            slides = parse_ppt_json(raw)
                        except Exception as e:
                            st.warning(f"PPT 内容生成失败（{e}），已回退到内置示例。")
                            slides = mock_ppt_slides(theme.strip())
                    else:
                        slides = mock_ppt_slides(theme.strip())
                st.session_state["ppt_slides"] = slides
                st.session_state["ppt_theme"] = theme.strip()

            if "ppt_slides" in st.session_state:
                st.markdown("**PPT 大纲预览**")
                prev_layout = None
                for i, s in enumerate(st.session_state["ppt_slides"], 1):
                    title = s.get("title", "未命名") if isinstance(s, dict) else str(s)
                    bullets = s.get("bullets", []) if isinstance(s, dict) else []
                    declared = s.get("layout", "") if isinstance(s, dict) else ""
                    prev_layout = _ppt_plan_layout(title, bullets, declared, prev_layout)
                    label = PPT_LAYOUT_LABELS.get(prev_layout, "列表")
                    st.markdown(f"**{i}. {title}**　`{label}`")
                    for b in bullets:
                        st.markdown(f"- {b}")

                buf = build_pptx(st.session_state.get("ppt_theme", "红色教育"), st.session_state["ppt_slides"])
                st.download_button(
                    "下载PPT（.pptx）",
                    buf.getvalue(),
                    file_name=f"{st.session_state.get('ppt_theme', '红色教育')}_PPT.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
            else:
                st.info("在左侧输入主题，点击“开始生成”。")

        # ----- PPT提示词 -----
        elif work_type == "PPT提示词":
            if generate and theme.strip():
                st.session_state["prompt_result"] = build_ppt_prompt(theme.strip(), n_slides, audience)

            if "prompt_result" in st.session_state:
                st.markdown("**生成好的PPT提示词**（可点击代码块右上角复制）")
                st.code(st.session_state["prompt_result"], language=None)
                st.download_button(
                    "下载提示词（.txt）",
                    st.session_state["prompt_result"],
                    file_name=f"{theme.strip() or '红色教育'}_PPT提示词.txt",
                )
            else:
                st.info("在左侧输入主题，点击“开始生成”。")


# ================= Tab 2：红色资源库 =================
with tab_archive:
    st.subheader("红色资源库")
    st.caption("内置红色史料 + AI 联网检索：任何关键词都能获取详细资料（对应附件1“红色资源整合与智能检索”）。")

    keyword = st.text_input("检索关键词", placeholder="输入人物、事件或类别，例如：长征、焦裕禄……")

    filtered = RED_ARCHIVES
    if keyword.strip():
        kw = keyword.strip()
        filtered = [
            a for a in RED_ARCHIVES
            if kw in a["title"] or kw in a["figure"] or kw in a["category"] or kw in a["summary"]
        ]

    st.markdown(f"内置库检索到 **{len(filtered)}** 条结果")
    for a in filtered:
        st.markdown(
            f"""
            <div style="background:#fff; border:1px solid #f0d6d2; border-left:4px solid #c62828;
                        border-radius:10px; padding:1rem 1.2rem; margin-bottom:0.8rem;">
                <div style="font-size:1.05rem; font-weight:700; color:#b71c1c;">
                    {a['title']}
                    <span style="font-weight:400; color:#999; font-size:0.82rem; margin-left:0.4rem;">{a['year']} · {a['category']}</span>
                </div>
                <div style="color:#999; font-size:0.82rem; margin-top:0.25rem;">相关人物：{a['figure']}</div>
                <div style="margin-top:0.45rem; line-height:1.7; color:#333;">{a['summary']}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # AI 联网检索
    if keyword.strip():
        st.markdown("---")
        st.markdown("**🌐 AI 联网检索**")
        if st.button(f"用 AI 生成「{keyword.strip()}」的详细资料", type="primary"):
            if is_real and api_key.strip():
                with st.spinner("正在联网检索，请稍候……"):
                    try:
                        st.session_state["ai_archive"] = call_deepseek(api_key, build_archive_prompt(keyword.strip()))
                    except Exception as e:
                        st.warning(f"检索失败（{e}），已回退到演示模式。")
                        st.session_state["ai_archive"] = mock_archive(keyword.strip())
            else:
                st.session_state["ai_archive"] = mock_archive(keyword.strip())

        if "ai_archive" in st.session_state:
            st.markdown(
                f"""<div style="background:#fdecea; border:1px solid #f5c6c6; border-radius:10px;
                            padding:1rem 1.2rem; margin-top:0.8rem; white-space:pre-wrap; line-height:1.7;">{st.session_state['ai_archive']}</div>""",
                unsafe_allow_html=True,
            )


# ================= Tab 3：AI 红色讲解员 =================
with tab_guide:
    st.subheader("AI 红色讲解员")
    st.caption("与 AI 讲解员对话，了解红色历史（演示模式返回内置讲解，真实模式调用 DeepSeek）。")

    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []

    for msg in st.session_state.chat_history:
        with st.chat_message(msg["role"]):
            st.write(msg["content"])

    question = st.chat_input("例如：给我讲讲长征的故事")

    if question:
        st.session_state.chat_history.append({"role": "user", "content": question})
        with st.chat_message("user"):
            st.write(question)

        with st.chat_message("assistant"):
            if is_real and api_key.strip():
                prompt = (
                    "你是一位红色文化讲解员，面向中小学生。请用通俗易懂、积极向上的语言，"
                    f"回答以下问题，注意尊重史实、弘扬革命精神：\n{question}"
                )
                try:
                    answer = call_deepseek(api_key, prompt)
                except Exception as e:
                    answer = f"（调用 DeepSeek 失败：{e}）\n\n同学们好！关于这个问题，先辈们的精神值得我们永远铭记。"
            else:
                answer = (
                    f"同学们好！关于“{question}”，这是一段值得我们永远铭记的历史。\n\n"
                    "在革命和建设年代，无数先辈用坚定的理想信念和无私的奉献，"
                    "为我们今天的幸福生活铺就了道路。让我们铭记历史、缅怀先烈，"
                    "把红色基因一代代传承下去，真正做到——薪火相传，强国有我！\n\n"
                    "（演示模式：接入 DeepSeek 后，可实时生成针对性的详细讲解。）"
                )
            st.write(answer)
        st.session_state.chat_history.append({"role": "assistant", "content": answer})
