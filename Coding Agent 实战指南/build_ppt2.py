# -*- coding: utf-8 -*-
import re
"""构建《Coding Agent 工程实践（通用版）》演示文稿 — 深空科技风 1280x720
厂商无关：全篇不出现任何具体产品名，方法论通用。"""
import os
from PIL import Image
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

ROOT = os.path.dirname(os.path.abspath(__file__))
A = os.path.join(ROOT, "assets")
OUT = os.path.join(ROOT, "Coding Agent 工程实践（通用版）.pptx")

BG = RGBColor(0x0F, 0x1B, 0x2D)
TEAL = RGBColor(0x35, 0xD0, 0xBA)
BLUE = RGBColor(0x4C, 0x8D, 0xFF)
TXT = RGBColor(0xE8, 0xED, 0xF5)
CARD = RGBColor(0x1C, 0x28, 0x39)
CARD_BR = RGBColor(0x29, 0x34, 0x45)
MUT50 = RGBColor(0x7B, 0x84, 0x94)
MUT55 = RGBColor(0x8A, 0x94, 0xA6)
MUT65 = RGBColor(0xA6, 0xAD, 0xB8)
MUT68 = RGBColor(0xA9, 0xB0, 0xBB)
MUT72 = RGBColor(0xAB, 0xB2, 0xBD)
MUT78 = RGBColor(0xBC, 0xC3, 0xCD)
MUT85 = RGBColor(0xC8, 0xCE, 0xD6)
TEAL_BG = RGBColor(0x13, 0x2B, 0x3A)
TEAL_BG2 = RGBColor(0x14, 0x31, 0x40)
TEAL_BR = RGBColor(0x20, 0x6D, 0x6D)
BLUE_BG = RGBColor(0x18, 0x2C, 0x4D)
BLUE_BR = RGBColor(0x2D, 0x54, 0x96)
ROW_BG = RGBColor(0x18, 0x24, 0x36)
TBL_BG = RGBColor(0x14, 0x1F, 0x30)

CN = "Microsoft YaHei"
MONO = "Consolas"
TOTAL = 31


def emu(px):
    return Emu(int(px * 9525))


def set_run(r, text, size, color=TXT, bold=False, font=CN, italic=False, spacing=None):
    r.text = text
    r.font.size = Pt(size * 0.75)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.name = font
    rPr = r._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set("typeface", CN if font == CN else font)
    if spacing:
        rPr.set("spc", str(int(spacing * 100)))


def tb(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=1.0, space_after=0):
    box = slide.shapes.add_textbox(emu(x), emu(y), emu(w), emu(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        if space_after:
            p.space_after = Pt(space_after * 0.75)
        for text, kw in para:
            set_run(p.add_run(), text, **kw)
    return box


def rect(slide, x, y, w, h, fill=None, line=None, radius=None, line_w=1.0):
    shp_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius is not None else MSO_SHAPE.RECTANGLE
    sp = slide.shapes.add_shape(shp_type, emu(x), emu(y), emu(w), emu(h))
    if radius is not None:
        try:
            sp.adjustments[0] = radius / min(w, h) if min(w, h) > 0 else 0.1
        except Exception:
            pass
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
    sp.shadow.inherit = False
    return sp


def hline(slide, x, y, w, color, thickness=2):
    return rect(slide, x, y, w, thickness, fill=color)


def arrow(slide, x1, y1, x2, y2, color=TEAL, width=2.0):
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, emu(x1), emu(y1), emu(x2), emu(y2))
    conn.line.color.rgb = color
    conn.line.width = Pt(width)
    conn.shadow.inherit = False
    ln = conn.line._get_or_add_ln()
    tail = ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"})
    ln.append(tail)
    return conn


def pic(slide, path, x, y, w, h):
    return slide.shapes.add_picture(path, emu(x), emu(y), emu(w), emu(h))


def new_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = BG
    return s


def masked(src, dst, l_a, r_a, bottom=None):
    im = Image.open(src).convert("RGBA")
    w, h = im.size
    overlay = Image.new("RGBA", (w, h), (15, 27, 45, 0))
    px = overlay.load()
    for xx in range(w):
        t = xx / (w - 1)
        a = int(l_a + (r_a - l_a) * t)
        for yy in range(h):
            px[xx, yy] = (15, 27, 45, a)
    if bottom:
        bh = int(h * bottom[0])
        for yy in range(bh, h):
            t = (yy - bh) / (h - bh)
            a = int(bottom[1] * t)
            for xx in range(w):
                r0, g0, b0, a0 = px[xx, yy]
                px[xx, yy] = (15, 27, 45, min(255, a0 + a))
    im = Image.alpha_composite(im, overlay).convert("RGB")
    im.save(dst)


def header(slide, title, sub, accent=None):
    tb(slide, 64, 26, 1000, 46, [[(title, dict(size=36, bold=True))]])
    hline(slide, 64, 82, 56, accent or TEAL, 3)
    tb(slide, 136, 72, 1000, 26, [[(sub, dict(size=16, color=MUT55))]])


def demo_badge(slide, n=1):
    """右上角「现场演示」角标——编号只写进备注栏，页面上不出现在听众视野里。"""
    label = "现场演示" if n == 1 else "现场演示 ×%d" % n
    rect(slide, 1064, 32, 152, 32, fill=BLUE_BG, line=BLUE_BR, radius=16)
    tb(slide, 1064, 38, 152, 22, [[(label, dict(size=12.5, bold=True, color=BLUE))]], align=PP_ALIGN.CENTER)


def footer(slide, chapter):
    """页码按幻灯片实际位置自动计算——插入或删除页面后无需手工重编号。"""
    n = len(prs.slides._sldIdLst)
    rect(slide, 64, 688, 12, 12, fill=TEAL_BG2, line=TEAL_BR, radius=2)
    tb(slide, 88, 684, 700, 22, [[("Coding Agent 工程实践 · " + chapter, dict(size=14, color=MUT50))]])
    tb(slide, 1000, 684, 216, 22, [[("%02d / %d" % (n, TOTAL), dict(size=14, color=TEAL, font=MONO))]], align=PP_ALIGN.RIGHT)


def judge_card(slide, x, y, w, h, text, size=17.5):
    rect(slide, x, y, w, h, fill=TEAL_BG, line=TEAL_BR, radius=10)
    rect(slide, x + 24, y + 20, 10, 10, fill=TEAL, radius=5)
    tb(slide, x + 44, y + 14, 120, 22, [[("结论", dict(size=16, bold=True, color=TEAL))]])
    tb(slide, x + 24, y + 44, w - 48, h - 58, [[(text, dict(size=size, color=TXT))]], line_spacing=1.6)


def bottom_judge(slide, x, y, w, h, text, size=15.5):
    rect(slide, x, y, w, h, fill=TEAL_BG, line=TEAL_BR, radius=10)
    rect(slide, x + 22, y + h / 2 - 5, 10, 10, fill=TEAL, radius=5)
    tb(slide, x + 42, y, 80, h, [[("结论", dict(size=15, bold=True, color=TEAL))]], anchor=MSO_ANCHOR.MIDDLE)
    tb(slide, x + 112, y + 8, w - 132, h - 16, [[(text, dict(size=size, color=TXT))]], line_spacing=1.45, anchor=MSO_ANCHOR.MIDDLE)


def section_slide(prs, num, img, dst_name, title, sub, desc):
    s = new_slide(prs)
    dst = os.path.join(A, dst_name)
    masked(img, dst, 240, 90)
    pic(s, dst, 0, 0, 1280, 720)
    tb(s, 880, -30, 400, 340, [[(num, dict(size=260, bold=True, color=RGBColor(0x1A, 0x26, 0x38), font=MONO))]], align=PP_ALIGN.RIGHT)
    tb(s, 64, 210, 400, 26, [[("CHAPTER " + num, dict(size=20, color=TEAL, font=MONO, spacing=4))]])
    hline(s, 250, 226, 72, TEAL, 3)
    tb(s, 64, 250, 700, 90, [[(title, dict(size=64, bold=True, spacing=2))]])
    tb(s, 64, 356, 700, 40, [[(sub, dict(size=24, color=MUT78))]])
    tb(s, 64, 416, 620, 90, [[(desc, dict(size=17, color=MUT55))]], line_spacing=1.7)
    return s


prs = Presentation()
prs.slide_width = Emu(12192000)
prs.slide_height = Emu(6858000)

# ============ P01 封面 ============
s = new_slide(prs)
masked(os.path.join(A, "cover_agent.png"), os.path.join(A, "g_cover.png"), 240, 40, bottom=(0.45, 200))
pic(s, os.path.join(A, "g_cover.png"), 0, 0, 1280, 720)
hline(s, 64, 158, 96, TEAL, 4)
tb(s, 176, 146, 560, 24, [[("UNIVERSAL METHODOLOGY · VENDOR NEUTRAL", dict(size=14, color=TEAL, font=MONO, spacing=3))]])
tb(s, 64, 186, 700, 100, [[("Coding Agent", dict(size=72, bold=True, font=MONO))]])
tb(s, 64, 286, 700, 92, [[("工程实践", dict(size=66, bold=True, color=TEAL))]])
tb(s, 64, 400, 660, 40, [[("通用编码智能体方法论 · 从“会用”到“敢用”", dict(size=22, color=MUT72))]])
rect(s, 64, 630, 150, 34, fill=CARD, line=CARD_BR, radius=17)
rect(s, 78, 643, 8, 8, fill=TEAL, radius=4)
tb(s, 96, 637, 120, 22, [[("研发效能内训", dict(size=14, color=TXT))]])
tb(s, 230, 637, 70, 22, [[("2026.10", dict(size=14, color=MUT55, font=MONO))]])
rect(s, 310, 630, 152, 34, fill=TEAL_BG, line=TEAL_BR, radius=17)
rect(s, 324, 643, 8, 8, fill=TEAL, radius=4)
tb(s, 342, 637, 110, 22, [[("主讲 · 杨庆刚", dict(size=14, color=TXT))]])
rect(s, 980, 650, 300, 70, fill=BG)
tb(s, 1000, 676, 260, 22, [[("BEST PRACTICES 101", dict(size=13, color=MUT50, font=MONO, spacing=2))]], align=PP_ALIGN.RIGHT)

# ============ P02 目录 ============
s = new_slide(prs)
tb(s, 64, 30, 400, 46, [[("目录", dict(size=36, bold=True))]])
hline(s, 64, 82, 56, TEAL, 3)
rect(s, 64, 160, 6, 120, fill=TEAL)
tb(s, 64, 300, 300, 80, [[("五个问题，", dict(size=24, bold=True))], [("一条方法论主线", dict(size=24, bold=True))]], line_spacing=1.4)
tb(s, 64, 380, 300, 150, [[("第 1 章解决用哪个，第 2 章给你看真实效果，其余四章讲方法。不绑定任何一款工具，换一个 Agent 这套方法原样生效。", dict(size=16, color=MUT55))]], line_spacing=1.7)
toc = [
    ("01", "选型篇 · 市场、形态与怎么选", "现在有哪些工具，我该用哪个？", "P00 - P00"),
    ("02", "实战篇 · 八段真实会话", "它到底能做成什么样？先建立印象。", "P00 - P00"),
    ("03", "认知篇 · 一个约束与一次转变", "为什么这套方法有效，而不是碰巧有用？", "P00 - P00"),
    ("04", "方法篇 · 单次任务的标准动作", "一个需求从提出到合入，每步该怎么做？", "P00 - P00"),
    ("05", "工程篇 · 上下文工程四层", "产出质量的差距，到底从哪一层来？", "P00 - P00"),
    ("06", "规模化篇 · 权限、并行与落地", "如何从一个人快，走到整个团队快？", "P00 - P00"),
]
for i, (no, title, q, pg) in enumerate(toc):
    y = 128 + i * 88
    rect(s, 420, y, 796, 78, fill=CARD, line=CARD_BR, radius=10)
    tb(s, 448, y + 18, 60, 40, [[(no, dict(size=26, bold=True, color=TEAL if i == 4 else BLUE, font=MONO))]])
    tb(s, 522, y + 10, 560, 28, [[(title, dict(size=20, bold=True))]])
    tb(s, 522, y + 44, 560, 24, [[(q, dict(size=14.5, color=MUT55))]])
    tb(s, 1090, y + 28, 110, 22, [[(pg, dict(size=13, color=MUT50, font=MONO))]], align=PP_ALIGN.RIGHT)
footer(s, "目录")

# ============ P03 扉页 选型篇 ============
section_slide(prs, "01", os.path.join(A, "ch_landscape.png"), "g_ch00.png", "选型篇", "先看战场，再定工具",
              "这一章只回答两件事：现在有什么、你该选哪个。下一章要打开几段真实会话，自然会提到工具名；再往后的四章，不依赖你选了谁。")

# ============ P04 全球格局 ============
s = new_slide(prs)
header(s, "全球格局：一个已经重排过的市场", "九成职业开发者每周都在用。座次七个月里换了一遍")
rect(s, 64, 140, 700, 386, fill=CARD, line=CARD_BR, radius=10)
tb(s, 88, 156, 420, 26, [[("工作中使用率 · 2026 年 5–7 月", dict(size=18, bold=True))]])
tb(s, 88, 184, 560, 22, [[("同一人可多选，故总和 > 100%；口径是“使用率”，不是营收份额", dict(size=12.5, color=MUT50))]])
gl = [("Claude Code", 39, "1月 18%"), ("GitHub Copilot", 21, "1月 29%"), ("OpenAI Codex", 16, "1月 3%"),
      ("Cursor", 12, "1月 18%"), ("JetBrains AI", 9, "1月 13%"), ("OpenCode", 7, "开源"),
      ("Google Antigravity", 6, "印度 15%")]
for i, (name, pct, note) in enumerate(gl):
    y = 216 + i * 42
    tb(s, 84, y, 152, 24, [[(name, dict(size=13.5, color=MUT85 if i == 0 else MUT72))]])
    rect(s, 244, y + 4, 296, 18, fill=ROW_BG, radius=9)
    rect(s, 244, y + 4, int(pct * 7.2), 18, fill=TEAL if i == 0 else BLUE, radius=9)
    tb(s, 552, y, 60, 22, [[("%d%%" % pct, dict(size=14.5, bold=True, color=TXT if i == 0 else MUT78))]])
    tb(s, 620, y + 2, 130, 22, [[(note, dict(size=11.5, color=MUT50, font=MONO))]])
rect(s, 800, 140, 416, 168, fill=TEAL_BG, line=TEAL_BR, radius=10)
tb(s, 824, 160, 380, 34, [[("90%", dict(size=32, bold=True, color=TEAL)), ("   每周至少用一次", dict(size=15, color=MUT78))]])
tb(s, 824, 208, 380, 34, [[("68%", dict(size=32, bold=True, color=TEAL)), ("   每天都在用", dict(size=15, color=MUT78))]])
tb(s, 824, 254, 380, 40, [[("样本 15,000+ 职业开发者，按地区与语言加权", dict(size=12.5, color=MUT68))]], line_spacing=1.4)
judge_card(s, 800, 328, 416, 198, "七个月里两个翻倍，两个下滑。押在今天的排名上，半年后多半要返工。选形态要看它能不能长跑，别只看当期名次。", 16.5)
rect(s, 64, 536, 1152, 76, fill=CARD, line=CARD_BR, radius=9)
tb(s, 88, 546, 1108, 58, [[("数据来源：JetBrains《Developer Ecosystem Survey 2026》，15,000+ 职业开发者，2026-05 至 07 田野期；“1 月”为同年 1 月波次，用于对照七个月变化。", dict(size=12, color=MUT55))],
                          [("还有两家没进该调研口径，单独看：DeepSeek Harness（2026-08 开源，MIT，发布两周 GitHub 破 3 万星）、Pi（2026-10 发布 1.0，MIT 终端智能体）。两家都走自托管、可换模型的开源路线。", dict(size=12, color=MUT68))]], line_spacing=1.35)
footer(s, "选型篇")

# ============ P05 国内格局 ============
s = new_slide(prs)
header(s, "国内格局：两种口径，且榜单会漏掉新人", "用的人最多和赚的钱最多的不是同一家。2026 下半年还有三家刚入场")
rect(s, 64, 140, 370, 300, fill=CARD, line=CARD_BR, radius=10)
tb(s, 88, 156, 330, 26, [[("企业侧 · 收入份额", dict(size=17, bold=True))]])
tb(s, 88, 184, 330, 22, [[("IDC 口径 · 2026 · 企业付费收入", dict(size=12, color=MUT50))]])
ent = [("阿里 Qoder", 47.6), ("智谱 ZCode", 11.5), ("商汤 Raccoon", 10.5),
       ("腾讯 CodeBuddy", 6.9), ("百度 文心快码", 6.0)]
for i, (n, p) in enumerate(ent):
    y = 214 + i * 40
    tb(s, 84, y, 148, 24, [[(n, dict(size=12.5, color=MUT85 if i == 0 else MUT72))]])
    rect(s, 236, y + 4, 150, 16, fill=ROW_BG, radius=8)
    rect(s, 236, y + 4, max(6, int(p * 3.15)), 16, fill=TEAL if i == 0 else BLUE, radius=8)
    tb(s, 394, y, 46, 22, [[("%.1f%%" % p, dict(size=12, color=MUT78))]])

rect(s, 454, 140, 370, 300, fill=CARD, line=CARD_BR, radius=10)
tb(s, 478, 156, 330, 26, [[("个人侧 · IDE 使用率", dict(size=17, bold=True))]])
tb(s, 478, 184, 330, 22, [[("媒体引用口径 · 2026 · 实际在用", dict(size=12, color=MUT50))]])
per = [("字节 Trae", 41.2), ("阿里 通义灵码", 18.5), ("百度 文心快码", 12.3)]
for i, (n, p) in enumerate(per):
    y = 214 + i * 44
    tb(s, 474, y, 148, 24, [[(n, dict(size=12.5, color=MUT85 if i == 0 else MUT72))]])
    rect(s, 626, y + 4, 150, 16, fill=ROW_BG, radius=8)
    rect(s, 626, y + 4, int(p * 3.64), 16, fill=TEAL if i == 0 else BLUE, radius=8)
    tb(s, 784, y, 46, 22, [[("%.1f%%" % p, dict(size=12, color=MUT78))]])
tb(s, 474, 360, 330, 70, [[("文心快码服务 800 万+ 开发者；Cursor 中国区使用率一年内由 28% 降至 16%（JetBrains 2026）。", dict(size=11.5, color=MUT68))]], line_spacing=1.45)
rect(s, 844, 140, 372, 300, fill=CARD, line=TEAL_BR, radius=10)
tb(s, 868, 156, 330, 26, [[("榜单之外 · 2026 下半年新入场", dict(size=17, bold=True, color=TEAL))]])
tb(s, 868, 184, 330, 22, [[("发布晚于榜单口径，份额尚未反映", dict(size=12, color=MUT50))]])
new3 = [
    ("Kimi Code · 月之暗面", "2026-09 桌面端 1.0", "内置终端 / 浏览器 / Git，Plan · Swarm · Tower 多模式协作"),
    ("MiniMax Code · 稀宇", "2026-06 随 M3 发布", "桌面客户端（多端），基于 OpenCode + Pi 构建，支持 Computer Use"),
    ("MiMo Code · 小米", "2026-06 开源 MIT", "终端原生，fork 自 OpenCode；持久记忆 + Compose 模式"),
]
for i, (n, t, d) in enumerate(new3):
    y = 212 + i * 74
    tb(s, 864, y, 336, 22, [[(n, dict(size=13.5, bold=True, color=TEAL))]])
    tb(s, 864, y + 22, 336, 20, [[(t, dict(size=11.5, color=MUT55, font=MONO))]])
    tb(s, 864, y + 42, 336, 34, [[(d, dict(size=11.5, color=MUT72))]], line_spacing=1.35)
judge_card(s, 64, 456, 740, 176, "国内用的人多和赚的钱多是两回事，份额榜单还会漏掉刚入场的人。MiniMax Code 和 MiMo Code 都直接基于开源 OpenCode 构建，模型厂商自己也不写 harness，而是 fork 一份再叠自己的模型。选型看合规边界、工具链和自主权，别看份额。", 16)
rect(s, 820, 456, 396, 176, fill=CARD, line=CARD_BR, radius=10)
tb(s, 844, 472, 340, 26, [[("数据来源与可信度", dict(size=16.5, bold=True, color=BLUE))]])
tb(s, 844, 506, 350, 118, [[("份额为 IDC《中国 AI 编程市场》报告的二手引用；三家新入场产品信息来自 2026-06 至 09 的官方发布，份额尚未覆盖。", dict(size=12.5, color=MUT68))]], line_spacing=1.5)
footer(s, "选型篇")

# ============ P06 开源路线 ============
s = new_slide(prs)
header(s, "开源路线：Harness 是可替换的那一层", "MIT 许可、可自托管、可换模型。同一个模型换一个 harness，就是另一个 Agent")
demo_badge(s, 1)
projs = [
    ("DeepSeek Harness", "插件化运行底座",
     [("许可 / 状态", "MIT · v0.1 预览版"), ("架构", "一切皆插件 · Cordis 内核"),
      ("上下文", "追加式事件流 · 自动压缩"), ("沙箱", "内置，可热插拔"),
      ("代价", "接口仍有破坏性变更")]),
    ("Pi", "极简终端智能体",
     [("许可 / 状态", "MIT · 1.0 已发布"), ("架构", "只给 read/write/edit/bash"),
      ("上下文", "系统提示约 2–3k token"), ("沙箱", "无，依赖外部环境"),
      ("代价", "官方明确不做子代理")]),
    ("OpenCode", "终端优先的全能选手",
     [("许可 / 状态", "MIT · 社区规模最大"), ("架构", "配置化工具 + MCP 外接"),
      ("上下文", "支持 75+ 模型提供方"), ("沙箱", "进程隔离 C/S 架构"),
      ("代价", "系统提示约 10k token")]),
]
for i, (name, tag, rows) in enumerate(projs):
    x = 64 + i * 391
    rect(s, x, 140, 370, 250, fill=CARD, line=CARD_BR, radius=10)
    tb(s, x + 18, 150, 334, 26, [[(name, dict(size=17, bold=True, color=TEAL))]])
    tb(s, x + 18, 178, 334, 22, [[(tag, dict(size=13, color=MUT68))]])
    hline(s, x + 18, 204, 334, CARD_BR, 1)
    for j, (k, v) in enumerate(rows):
        y = 214 + j * 33
        tb(s, x + 18, y, 96, 22, [[(k, dict(size=12, color=MUT55))]])
        tb(s, x + 104, y, 248, 30, [[(v, dict(size=12.5, color=TXT))]], line_spacing=1.3)
rect(s, 64, 404, 520, 100, fill=BLUE_BG, line=BLUE_BR, radius=10)
tb(s, 88, 416, 480, 26, [[("Model + Harness = Agent", dict(size=19, bold=True, color=BLUE, font=MONO))]])
tb(s, 88, 448, 480, 50, [[("模型负责想，harness 负责做：上下文怎么管、怎么隔离、工具怎么调度，都是 harness 说了算。", dict(size=13.5, color=TXT))]], line_spacing=1.45)
rect(s, 604, 404, 612, 100, fill=CARD, line=CARD_BR, radius=10)
tb(s, 628, 416, 580, 24, [[("还有这些（同为可自托管 / 可换模型）", dict(size=15, bold=True, color=TEAL))]])
tb(s, 628, 446, 580, 50, [[("Cline（VS Code，逐操作审批）· OpenHands（自主 + 沙箱）· Aider（git 原生）· Goose（MCP 原生）· Codex CLI / Gemini CLI（厂商开源）", dict(size=13, color=MUT78))]], line_spacing=1.45)
for (bx, bw, bcol, btitle, items, itcol) in [
        (64, 566, TEAL, "什么情况下该走开源路线", ["代码与数据不能出网（合规 / 数据本地化）",
         "必须能换模型，不愿被单一厂商绑定", "规模上来后，按席位计费成本线性增长"], TEAL),
        (650, 566, MUT55, "代价：开源 ≠ 省事", ["要自己搭建、维护与升级，没有厂商兜底",
         "早期版本接口不稳，生产环境慎用", "参数与状态随版本快速变化，选型前以仓库为准"], MUT50)]:
    rect(s, bx, 520, bw, 110, fill=CARD, line=CARD_BR, radius=10)
    tb(s, bx + 24, 532, 520, 24, [[(btitle, dict(size=16, bold=True, color=bcol))]])
    for i, t in enumerate(items):
        y = 566 + i * 22
        rect(s, bx + 24, y + 7, 8, 8, fill=itcol, radius=4)
        tb(s, bx + 40, y, 500, 20, [[(t, dict(size=13, color=MUT85 if bcol == TEAL else MUT68))]])
footer(s, "选型篇")

# ============ P07 五类形态 ============
s = new_slide(prs)
header(s, "五类形态：先认形态，再看品牌", "形态决定能力边界，品牌只决定细节")
rect(s, 64, 140, 1152, 364, fill=CARD, line=CARD_BR, radius=10)
for (x, w), h in zip([(80, 180), (268, 210), (486, 280), (774, 350)], ["形态", "代表产品", "擅长", "代价 / 边界"]):
    tb(s, x, 152, w, 24, [[(h, dict(size=13.5, bold=True, color=MUT55))]])
hline(s, 80, 180, 1120, CARD_BR, 1)
forms = [
    ("IDE 插件", "Copilot · 通义灵码 · CodeBuddy", "嵌入现有编辑器：补全、解释、局部重构，零迁移成本", "受编辑器边界约束；跨文件长任务与流水线能力弱"),
    ("命令行智能体", "Pi · Claude Code · Codex CLI · OpenCode", "直接操作仓库与终端，易脚本化，可无头进 CI", "无可视化界面；上手曲线陡；长任务要自己治理上下文"),
    ("编辑器内核分叉", "Cursor · Trae", "编辑器与智能体深度耦合，既有配置可一键迁移", "团队工具链被锁定；版本节奏跟随厂商"),
    ("桌面工作台", "腾讯 WorkBuddy · Qoder 工作台", "一人指挥多个专家智能体并行，适合复杂需求", "上下文与 token 消耗高；可复现性与可控性挑战"),
    ("云端智能体平台", "Codex Cloud · Antigravity · Qoder 云端", "任务在云端异步跑完，直接交付 PR，人不盯屏", "代码与上下文出网，需过合规；结果可控性最低"),
]
for i, row in enumerate(forms):
    y = 190 + i * 62
    if i % 2 == 0:
        rect(s, 72, y, 1136, 56, fill=ROW_BG, radius=6)
    tb(s, 80, y + 14, 180, 28, [[(row[0], dict(size=15, bold=True, color=TEAL))]])
    tb(s, 268, y + 6, 210, 46, [[(row[1], dict(size=12, color=MUT65))]], line_spacing=1.3)
    tb(s, 486, y + 6, 280, 46, [[(row[2], dict(size=12.5, color=MUT85))]], line_spacing=1.35)
    tb(s, 774, y + 6, 350, 46, [[(row[3], dict(size=12.5, color=MUT68))]], line_spacing=1.35)
rect(s, 64, 516, 1152, 56, fill=CARD, line=TEAL_BR, radius=9)
tb(s, 88, 528, 1104, 36, [[("可叠加的另一条线（见上页）· 开源、可自托管、可换模型：", dict(size=13.5, bold=True, color=TEAL)),
    ("DeepSeek Harness · Pi · OpenCode · Cline · OpenHands，和下面任何形态都能组合", dict(size=13.5, color=MUT78))]], line_spacing=1.35)
bottom_judge(s, 64, 588, 1152, 64, "先看它跑在哪、怎么接入，定形态。要不要开源自托管是另一条线，可以叠加。两个维度各算各的，别混在一起比。", 16)
footer(s, "选型篇")

# ============ P08 选型路径与本课立场 ============
s = new_slide(prs)
header(s, "怎么选：三步，一步一句话", "先划掉不能用的，再按工作方式定形态，最后用真实任务验证")
steps = [
    ("01", TEAL, "排除 · 把不能用的先划掉",
     "代码不能出网 → 不用云端平台；必须自托管 / 可换模型 → 只走开源路线；深度绑定某个 IDE → 优先该 IDE 的插件或内核分叉"),
    ("02", BLUE, "定形态 · 你要它替你做到哪一步",
     "只要补全与解释 → IDE 插件；要它自己跑完并自验 → 命令行 或 桌面工作台；要它异步跑完直接交 PR → 云端平台"),
    ("03", TEAL, "验证 · 同类里挑哪个，别靠感觉",
     "同一批真实任务，同一套可运行验收（跑测试、跑构建），跑满两周再定。第一天的新鲜感会骗人"),
]
for i, (no, col, sub, body) in enumerate(steps):
    y = 150 + i * 100
    rect(s, 64, y, 1152, 88, fill=CARD, line=CARD_BR, radius=10)
    rect(s, 64, y, 6, 88, fill=col, radius=3)
    tb(s, 96, y + 22, 90, 28, [[(no, dict(size=20, bold=True, color=col, font=MONO))]])
    tb(s, 190, y + 16, 420, 24, [[(sub, dict(size=15, bold=True))]])
    tb(s, 190, y + 46, 1000, 32, [[(body, dict(size=13.5, color=MUT78))]], line_spacing=1.35)
bottom_judge(s, 64, 470, 1152, 126, "统一比最优重要。团队选型要的是大家用同一个，不是选到最好的那个。规则文件、技能库、钩子配置都沉淀在同一套工具上。下一章看真实会话，会提到工具名。从第 3 章起不再提产品名，也不依赖你选了谁。", 17)
footer(s, "选型篇")

# ============ P09 扉页 实战篇 ============
section_slide(prs, "02", os.path.join(A, "ch_demo.png"), "g_ch05.png", "实战篇", "先看看它真能干什么",
              "先建立印象，方法后面再讲。索引表里是八段真实会话：同一个模型、同一套做法，在三个代码库里反复成立。")

# ============ P10 会话索引（第 2 章唯一内容页） ============
s = new_slide(prs)
header(s, "八段真实会话：从最朴素的用法，到最严格的用法", "先看它做到了什么。为什么这样有效，从第 3 章开始解释")
rect(s, 64, 140, 1152, 486, fill=CARD, line=CARD_BR, radius=10)
for x, w, h in [(84, 40, "#"), (130, 480, "它替我们做成了什么"), (630, 470, "背后那条方法 · 见第几页")]:
    tb(s, x, 150, w, 24, [[(h, dict(size=12.5, bold=True, color=MUT55))]])
hline(s, 80, 178, 1120, CARD_BR, 1)
demos = [
    ("①", "最简单的用法：一句话需求，从头做到尾", "基线 · 不加任何方法时的表现，后面每一段都是在它之上加一层", "Zcode › V3 › third · 消息 1–2"),
    ("②", "先计划再执行：先出方案，你认可了才动手", "@@单次任务的标准动作 · 规划先于编码", "Zcode › StdCommuProtocol › minor · 倒数第 2–1 条"),
    ("③", "改完一个问题后，让它自己回头审查一遍", "@@第一件事：给它一个可运行的验证 · 让它自己跑完并读结果", "Zcode › StdCommuProtocol › minor · 消息 1–6"),
    ("④", "对抗式审查：换一个全新上下文让它来找茬", "@@规模化：从一个 Agent 到一条流水线 · Writer / Reviewer 与新上下文评审", "Zcode › StdCommuProtocol › main · 消息 3–9"),
    ("⑤", "与同事同改一个项目时，定位问题出在哪", "@@任务层：把提示词写具体 · 指向信息源，让它自己去查", "Zcode › V3 › main · 倒数第 13–10 条"),
    ("⑥", "给确定的设计图，而不是一句「做得好看点」", "@@提示词改写：三组实战对照 · UI 改动要有视觉验证", "Zcode › V3 › main · 拉到底后第 3–4 条"),
    ("⑦", "同事发来新版 dll：先让它分析，再动手开发", "@@单次任务的标准动作 · 依赖变了，探索阶段就得重跑一遍", "Zcode › V3 › minor · 拉到底后第 5–9 条"),
    ("⑧", "开源 harness 的实机演示", "@@开源路线：Harness 是可替换的那一层 · 换一层底座，做法不变", "dsh › 液体处理平台 › main"),
]
for i, (no, what, method, loc) in enumerate(demos):
    y = 184 + i * 54
    if i % 2 == 0:
        rect(s, 72, y, 1136, 50, fill=ROW_BG, radius=6)
    tb(s, 84, y + 12, 40, 24, [[(no, dict(size=15, bold=True, color=TEAL, font=MONO))]])
    tb(s, 130, y + 10, 480, 34, [[(what, dict(size=12.5, bold=True, color=TXT))]], line_spacing=1.3)
    tb(s, 630, y + 10, 470, 34, [[(method, dict(size=11.5, color=MUT72))]], line_spacing=1.3)
    _ = loc  # 会话位置只写进备注栏
tb(s, 64, 640, 1152, 32, [[("发生在三个不同的真实项目里，跨两种运行环境。工具换了，这套做法没换。", dict(size=12.5, color=MUT55))]], line_spacing=1.4)
footer(s, "实战篇")

# ============ P14 扉页 认知篇 ============
section_slide(prs, "03", os.path.join(A, "ch_cognition.png"), "g_ch01.png", "认知篇", "先理解约束，再使用方法",
              "方法论不讲清楚为什么，就只能靠死记。理解这一个约束，后面所有做法都是它的推论，不用背清单。")

# ============ P04 第一性原理 ============
s = new_slide(prs)
header(s, "一个约束：上下文窗口是稀缺资源", "它是唯一需要持续管理的东西。后面讲的做法，几乎都能从这一条推出来。")
facts = [
    ("现象 01", "消耗极快", "一次调试会话或一轮代码库探索，就可能产出并消耗数万个 token。上下文不是慢慢用的，几步就能烧掉一大截。"),
    ("现象 02", "开始“遗忘”", "窗口快满时，模型开始遗忘你在第一句写下的指令。你以为它记得的约束，可能已经不起作用了。"),
    ("现象 03", "犯错率上升", "表现为它明明知道却还是写错了。其实它已经不知道了。窗口太满，不是模型变笨。"),
]
for i, (tag, t, d) in enumerate(facts):
    x = 64 + i * 384
    rect(s, x, 150, 368, 136, fill=CARD, line=CARD_BR, radius=10)
    tb(s, x + 22, 164, 200, 22, [[(tag, dict(size=13, color=BLUE, font=MONO, spacing=2))]])
    tb(s, x + 22, 188, 320, 28, [[(t, dict(size=18, bold=True))]])
    tb(s, x + 22, 220, 324, 60, [[(d, dict(size=13, color=MUT68))]], line_spacing=1.45)
coro = [
    ("推论一 · 能移出去的都移出去", "长期约定进规则文件，偶发知识进技能，重型调研交给子智能体。主对话只留本次任务。"),
    ("推论二 · 一次会话一类问题", "堆积的无关历史不算背景资料，只会拖慢它。任务串味时，立刻开新会话。"),
    ("推论三 · 别靠解释，靠判定", "解释写得越多，花的上下文越贵，换来的确定性却越来越少。真正有效的是能跑出对错的验证。"),
]
for i, (t, d) in enumerate(coro):
    y = 310 + i * 98
    rect(s, 64, y, 740, 88, fill=BLUE_BG, line=BLUE_BR, radius=10)
    tb(s, 88, y + 14, 700, 26, [[(t, dict(size=17, bold=True))]])
    tb(s, 88, y + 44, 692, 40, [[(d, dict(size=14, color=MUT78))]], line_spacing=1.45)
rect(s, 830, 310, 4, 88, fill=BLUE)
tb(s, 848, 306, 368, 100, [[("上下文窗口装着整场对话：你的每条消息、它读过的每个文件、每条命令的输出。", dict(size=15.5, color=MUT72))]], line_spacing=1.6)
judge_card(s, 830, 432, 386, 214, "把不该留在窗口里的挪出去（推论一、二），把留下来的花在能跑出对错的东西上（推论三）。前提是它得看得见，看不见就没法管。", 16.5)
footer(s, "认知篇")

# ============ P05 心智模型 ============
s = new_slide(prs)
header(s, "工作方式变了：从帮你写到替你做", "自主不等于不需要人，是人的工作位置变了")
stages = [
    ("阶段一", "代码补全", "人在写，AI 猜下一个词。Tab 键就是全部交互，出错成本极低。", False),
    ("阶段二", "对话改写", "人规划，AI 生成整段。选中、提问、粘贴，想法还是我来改。", False),
    ("阶段三", "自主执行", "AI 拆任务、调工具、改文件、跑测试；人定边界、给约束、做验收。", True),
]
for i, (tag, name, desc, hl) in enumerate(stages):
    y = 150 + i * 150
    rect(s, 64, y + 44, 40, 40, fill=TEAL_BG2 if hl else BLUE_BG, line=None if hl else BLUE, radius=20, line_w=2)
    rect(s, 72 if hl else 76, y + 60 if hl else y + 56, 16 if hl else 10, 16 if hl else 10, fill=BG if hl else BLUE, radius=8 if hl else 5)
    rect(s, 122, y, 622, 132, fill=TEAL_BG if hl else CARD, line=TEAL_BR if hl else CARD_BR, radius=10)
    tb(s, 146, y + 22, 110, 24, [[(tag, dict(size=14, color=TEAL if hl else BLUE, font=MONO))]])
    tb(s, 240, y + 16, 200, 32, [[(name, dict(size=21, bold=True))]])
    tb(s, 146, y + 58, 570, 62, [[(desc, dict(size=15, color=MUT68))]], line_spacing=1.5)
tb(s, 780, 150, 436, 26, [[("分工的变化", dict(size=18, bold=True)), ("   同一个需求，两种工作方式", dict(size=13, color=MUT55))]])
roles = [
    ("你的角色", "过去 · 执行环节：亲手写代码", "现在 · 边界环节（划定范围）＋ 验收环节（定义“完成”）"),
    ("AI 的角色", "过去 · 返回建议文本，看完还得你来改", "现在 · 探索、规划、实施，自己跑验证"),
    ("主要瓶颈", "过去 · 你敲键盘的速度", "现在 · 需求表述的精度＋验收信号的质量"),
]
for i, (lab, old, new) in enumerate(roles):
    y = 190 + i * 108
    rect(s, 780, y, 436, 98, fill=CARD, line=CARD_BR, radius=10)
    tb(s, 804, y + 11, 200, 20, [[(lab, dict(size=13.5, color=BLUE, bold=True))]])
    tb(s, 804, y + 33, 392, 20, [[("过去 · ", dict(size=13, bold=True, color=MUT55)), (old[5:], dict(size=13, color=MUT68))]])
    tb(s, 804, y + 56, 392, 38, [[("现在 · ", dict(size=13, bold=True, color=TEAL)), (new[5:], dict(size=13, color=TXT))]], line_spacing=1.4)
judge_card(s, 780, 500, 436, 170, "你的位置变了：过去在执行环节，现在在边界和验收环节。所以别用写得快不快评价它，要看它有没有替你省掉整个下午。", 15.5)
footer(s, "认知篇")

# ============ P06 扉页 方法篇 ============
section_slide(prs, "04", os.path.join(A, "ch_landscape.png"), "g_ch02.png", "方法篇", "一次需求，四个标准动作",
              "给可运行的验证、先探索再编码、让它采访你、规格先行。这四件事把玄学体验变成稳定产出。")

# ============ P07 第一条纪律 (Hero) ============
s = new_slide(prs)
header(s, "第一件事：给它一个可运行的验证", "没有验证手段时，看起来做完了就是唯一信号。你本人成了那个验证回路")
demo_badge(s, 1)
rect(s, 64, 150, 560, 118, fill=CARD, line=CARD_BR, radius=10)
tb(s, 88, 164, 200, 24, [[("没有验证", dict(size=17, bold=True, color=BLUE))]])
tb(s, 88, 194, 510, 66, [[("它只能以看起来像为准。每个错误都等你发现，你把验证从机器搬回了自己身上，而且是串行的。", dict(size=14.5, color=MUT72))]], line_spacing=1.5)
arrow(s, 344, 272, 344, 294, TEAL, 2)
rect(s, 64, 296, 560, 126, fill=TEAL_BG, line=TEAL_BR, radius=10)
tb(s, 88, 310, 440, 24, [[("给它一个能读出 pass / fail 的命令", dict(size=17, bold=True, color=TEAL))]])
tb(s, 88, 340, 510, 70, [[("测试套件、构建退出码、Linter、输出对比脚本、UI 截图比对。它做，它跑，它读结果，它继续改，循环自己闭合。", dict(size=14.5, color=TXT))]], line_spacing=1.5)
tb(s, 64, 436, 560, 44, [[("注意：它的检查通过不等于你的验收通过，两者是不同层级。上线前你仍需端到端确认。", dict(size=14, color=MUT65))]], line_spacing=1.5)
judge_card(s, 64, 494, 560, 134, "要证据，不要声明。完成后必须附上命令原文和原始输出，改 UI 附截图。看证据比自己重跑一遍快得多，你没在旁观的会话也用得上。", 15.5)
tb(s, 660, 142, 200, 110, [[("4", dict(size=96, bold=True, color=TEAL, font=MONO))]])
tb(s, 770, 176, 446, 60, [[("级验证设置", dict(size=22, bold=True))], [("逐级降低你的注意力消耗", dict(size=14.5, color=MUT55))]], line_spacing=1.4)
gates = [
    ("L1 · 单次提示", "同一条消息里要求跑检查并迭代，任何任务今天就能用"),
    ("L2 · 会话级目标", "由独立评估器每轮复核目标是否达成，没达成就继续。适合中长任务。"),
    ("L3 · 强制检查", "用脚本钩子在轮次结束前强制检查，不通过不许停。适合零例外的规范。"),
    ("L4 · 第三方复核", "全新上下文的评审智能体。做题的是它，打分的不是它，适合无人值守。"),
]
for i, (t, d) in enumerate(gates):
    y = 262 + i * 88
    rect(s, 660, y, 556, 80, fill=CARD, line=CARD_BR, radius=9)
    rect(s, 660, y, 4, 80, fill=TEAL if i >= 2 else BLUE)
    tb(s, 686, y + 12, 500, 24, [[(t, dict(size=16, bold=True, color=TEAL if i >= 2 else TXT))]])
    tb(s, 686, y + 40, 510, 34, [[(d, dict(size=13.5, color=MUT72))]], line_spacing=1.4)
footer(s, "方法篇")

# ============ P08 提示词改写实战 ============
s = new_slide(prs)
header(s, "提示词改写：三组实战对照", "加约束比加形容词有效。每一组好例子，都附带了能判定的验收")
demo_badge(s, 1)
pairs = [
    ("给出验收标准", "实现一个校验邮箱的函数",
     "写一个 validateEmail：user@example.com → true、invalid → false、user@.com → false。实现后运行测试，直到全部通过。"),
    ("UI 改动要视觉验证", "把仪表盘弄好看点",
     "按这张设计稿实现（附截图）。完成后对结果截图并与原图对比，逐条列出差异并修复。"),
    ("修根因，不压症状", "构建挂了",
     "构建失败，报错如下（粘贴完整错误）。修复并验证构建通过。处理根因，不要压制报错。"),
]
for i, (tag, bad, good) in enumerate(pairs):
    y = 150 + i * 148
    rect(s, 64, y, 540, 130, fill=CARD, line=CARD_BR, radius=10)
    tb(s, 88, y + 14, 200, 22, [[("✕  改写前", dict(size=14, bold=True, color=MUT55))]])
    tb(s, 88, y + 42, 490, 74, [[(bad, dict(size=16, color=MUT68))]], line_spacing=1.5)
    arrow(s, 612, y + 65, 650, y + 65, TEAL, 2)
    rect(s, 660, y, 556, 130, fill=TEAL_BG, line=TEAL_BR, radius=10)
    tb(s, 684, y + 14, 300, 22, [[("✓  改写后 · " + tag, dict(size=14, bold=True, color=TEAL))]])
    tb(s, 684, y + 42, 508, 74, [[(good, dict(size=14.5, color=TXT))]], line_spacing=1.45)
bottom_judge(s, 64, 606, 1152, 52, "一句话自检：把完成定义盖住，如果你判断不出这部分算不算完成，它就还不够具体")
footer(s, "方法篇")

# ============ P09 工作流四阶段 ============
s = new_slide(prs)
header(s, "单次任务的标准动作", "探索 → 规划 → 编码 → 提交。分开是因为各阶段的失败成本完全不同")
demo_badge(s, 3)
steps = [
    ("01", "探索", "只读模式下读懂相关代码，期间不得修改。让它讲清现状，而不是急着动手。"),
    ("02", "规划", "产出实施方案：改哪些文件、流程是什么。人可以直接修改方案再批准。"),
    ("03", "编码", "退出只读，按方案实现；为新逻辑补测试，跑到通过为止。"),
    ("04", "提交", "描述性提交信息，开 PR，附验证输出作为证据。"),
]
for i, (no, t, d) in enumerate(steps):
    x = 64 + i * 296
    rect(s, x, 150, 264, 176, fill=CARD, line=CARD_BR, radius=10)
    tb(s, x + 20, 166, 80, 30, [[(no, dict(size=22, bold=True, color=TEAL, font=MONO))]])
    tb(s, x + 20, 202, 200, 28, [[(t, dict(size=19, bold=True))]])
    tb(s, x + 20, 236, 224, 80, [[(d, dict(size=13, color=MUT68))]], line_spacing=1.45)
    if i < 3:
        arrow(s, x + 264, 238, x + 296, 238, BLUE, 1.75)
rect(s, 64, 370, 560, 252, fill=CARD, line=CARD_BR, radius=10)
tb(s, 88, 386, 300, 26, [[("何时跳过规划", dict(size=18, bold=True))]])
tb(s, 88, 422, 510, 66, [[("直接做：", dict(size=14.5, bold=True, color=TEAL)), ("改错别字、加日志、重命名、单文件局部小修。", dict(size=14.5, color=MUT72))]], line_spacing=1.5)
tb(s, 88, 492, 510, 66, [[("值得规划：", dict(size=14.5, bold=True, color=BLUE)), ("你对方案不确定 / 改动横跨多文件 / 你不熟悉被改的那部分代码。", dict(size=14.5, color=MUT72))]], line_spacing=1.5)
tb(s, 88, 562, 510, 48, [[("判据：一句话能描述清楚的 diff，就不要上仪式。", dict(size=14.5, color=TXT))]], line_spacing=1.5)
judge_card(s, 660, 370, 556, 252, "规划的成本是固定的，划不划算取决于做错方向的概率。别给每个任务都上流程。探索错了，损失几轮对话；在错误方向写完 800 行才被发现，损失的是你的信任和一下午。", 17)
footer(s, "方法篇")

# ============ P10 让 Agent 采访你 ============
s = new_slide(prs)
header(s, "大特性起手式：先让它采访你", "不确定的地方不该让它猜，该让它问。把规格改精确，回报高于盯着它实现")
flow = [
    ("1", "发出采访邀请", "“我想构建 X。请详细采访我：技术实现、UI/UX、边界情况、顾虑与权衡。不问显而易见的，深挖我可能没考虑到的。”"),
    ("2", "产出自包含 SPEC", "规格写进文件：点名涉及的文件与接口，声明什么不在范围内，以端到端验证步骤收尾。"),
    ("3", "开新会话执行", "干净上下文，全部注意力给实施。书面规格随时可引用、可评审、可交接。"),
]
for i, (no, t, d) in enumerate(flow):
    y = 150 + i * 152
    rect(s, 64, y, 560, 138, fill=CARD, line=CARD_BR, radius=10)
    rect(s, 88, y + 22, 34, 34, fill=TEAL_BG2, line=TEAL_BR, radius=17)
    tb(s, 88, y + 27, 34, 24, [[(no, dict(size=17, bold=True, color=TEAL, font=MONO))]], align=PP_ALIGN.CENTER)
    tb(s, 140, y + 18, 300, 28, [[(t, dict(size=18, bold=True))]])
    tb(s, 140, y + 52, 460, 72, [[(d, dict(size=13.5, color=MUT72))]], line_spacing=1.5)
rect(s, 660, 150, 556, 214, fill=CARD, line=CARD_BR, radius=10)
tb(s, 684, 166, 400, 26, [[("什么样的规格才算好", dict(size=18, bold=True))]])
spec_pts = [
    ("自包含", "不读对话记录也能看懂，别人和未来的你都能直接接手"),
    ("有边界", "明确写下什么不做。范围蔓延是 Agent 返工的最大来源"),
    ("可验收", "以端到端验证步骤收尾，跑一条命令就知道有没有完成"),
]
for i, (t, d) in enumerate(spec_pts):
    y = 204 + i * 52
    rect(s, 684, y + 4, 10, 10, fill=TEAL, radius=5)
    tb(s, 704, y - 2, 480, 44, [[(t + " · ", dict(size=14.5, bold=True, color=TXT)), (d, dict(size=14, color=MUT72))]], line_spacing=1.45)
judge_card(s, 660, 388, 556, 240, "口述需求、边做边改，每一次口头纠偏都会同时污染上下文和代码。规格没定就开工，是最高频的失败模式之一。把规格写精确的时间，回报高于你盯着它实现的时间。", 17)
footer(s, "方法篇")

# ============ P11 扉页 工程篇 ============
section_slide(prs, "05", os.path.join(A, "ch_practice.png"), "g_ch03.png", "工程篇", "产出差距在上下文结构，不在模型",
              "同一个模型，在不同团队手里产出天差地别。差距来自四层：规则怎么沉淀、任务怎么写、能力怎么扩展、会话怎么治理。自底向上讲。")

# ============ P12 上下文工程四层 (Hero) ============
s = new_slide(prs)
header(s, "上下文工程：这门手艺的主体", "五层结构。把静态约束从提示词里搬出去，主线程只留本次任务的约定")
layers = [
    ("⑤", "会话层", "治理", "清理、压缩、回溯、子智能体隔离。管理已经进窗口的内容", False),
    ("④", "能力层", "按需", "技能、子智能体、外部工具、事件钩子，都不常驻，需要时才加载", False),
    ("③", "任务层", "临时", "本次的任务简报：目标、上下文、约束、验收，每次任务临时提供", True),
    ("②", "规则层", "常驻", "每次会话都加载的项目约定，必须极短，只放普适的东西", True),
    ("①", "模型层", "底座", "基座推理与代码能力，决定下限，我们换不了（第 3 章已讲过）", False),
]
widths = [420, 480, 540, 600, 660]
for i, (no, name, tag, desc, hl) in enumerate(layers):
    w = widths[i]
    x = 414 - w / 2
    y = 150 + i * 92
    rect(s, x, y, w, 84, fill=TEAL_BG if hl else CARD, line=TEAL_BR if hl else CARD_BR, radius=10)
    tb(s, x + 18, y + 24, 46, 36, [[(no, dict(size=22, bold=True, color=TEAL if hl else BLUE, font=MONO))]])
    tb(s, x + 66, y + 14, 200, 28, [[(name, dict(size=18, bold=True)), ("  " + tag, dict(size=13, color=TEAL if hl else BLUE))]])
    tb(s, x + 66, y + 46, w - 90, 32, [[(desc, dict(size=12.5, color=MUT72))]], line_spacing=1.35)
tb(s, 800, 152, 416, 130, [[("同一个模型，在不同团队手里产出天差地别。差距几乎都出在 ②③④：规则写得准不准，简报给得清不清，能力接得全不全。", dict(size=17, color=TXT))]], line_spacing=1.7)
rect(s, 800, 306, 4, 88, fill=BLUE)
tb(s, 818, 306, 398, 96, [[("模型的差距是线性的，上下文结构的差距是乘性的。", dict(size=15.5, color=MUT72))]], line_spacing=1.6)
judge_card(s, 800, 440, 416, 190, "模型决定下限，上下文结构决定上限。换模型不如补上下文", 17.5)
footer(s, "工程篇")

# ============ P13 规则层 ============
s = new_slide(prs)
header(s, "规则层：把默契写进仓库", "每次会话都加载，所以只放普适的东西，而且必须极短")
rect(s, 64, 150, 316, 470, fill=CARD, line=TEAL_BR, radius=10)
tb(s, 88, 166, 200, 26, [[("✓ 应该写", dict(size=18, bold=True, color=TEAL))]])
inc = ["它猜不出来的构建/测试命令", "与默认习惯不同的代码风格", "测试指令与偏好的运行器", "仓库礼仪：分支命名、PR 规范", "本项目特有的架构决策", "环境怪癖与不自明的坑"]
for i, t in enumerate(inc):
    y = 204 + i * 66
    rect(s, 88, y + 6, 8, 8, fill=TEAL, radius=4)
    tb(s, 104, y, 256, 58, [[(t, dict(size=14, color=MUT85))]], line_spacing=1.4)
rect(s, 396, 150, 316, 470, fill=CARD, line=CARD_BR, radius=10)
tb(s, 420, 166, 200, 26, [[("✕ 不该写", dict(size=18, bold=True, color=MUT55))]])
exc = ["读一遍代码就能明白的东西", "语言通用规范（它本来就会）", "详细 API 文档（改为给链接）", "频繁变动的信息", "逐文件的代码库描述", "“写整洁代码”式正确废话"]
for i, t in enumerate(exc):
    y = 204 + i * 66
    rect(s, 420, y + 6, 8, 8, fill=MUT50, radius=4)
    tb(s, 436, y, 256, 58, [[(t, dict(size=14, color=MUT68))]], line_spacing=1.4)
rect(s, 740, 150, 476, 118, fill=BLUE_BG, line=BLUE_BR, radius=10)
tb(s, 764, 164, 300, 26, [[("篇幅要求（最重要）", dict(size=17, bold=True, color=BLUE))]])
tb(s, 764, 196, 428, 62, [[("对每一行问：删掉这行会让它犯错吗？答不上来就删。臃肿的规则会让它忽略你真正的指令，控制在 100 行以内", dict(size=14, color=TXT))]], line_spacing=1.5)
rect(s, 740, 284, 476, 170, fill=CARD, line=CARD_BR, radius=10)
tb(s, 764, 300, 300, 26, [[("排错指南", dict(size=17, bold=True))]])
debugs = [("明令禁止仍反复出现", "→ 文件太长，规则被淹没：剪"),
          ("问了规则里已答的问题", "→ 表述有歧义：改写得更具体"),
          ("单条指令总被跳过", "→ 只在这一行加强调，别处处加")]
for i, (a, b) in enumerate(debugs):
    y = 336 + i * 38
    tb(s, 764, y, 220, 32, [[(a, dict(size=13.5, color=MUT78))]])
    tb(s, 976, y, 224, 32, [[(b, dict(size=13.5, color=TEAL))]])
judge_card(s, 740, 470, 476, 172, "把它当代码对待：出问题就审查，定期修剪，用行为是否真的改变来验证改动。纳入 git 由团队共建，越用越准", 16)
footer(s, "工程篇")

# ============ P14 任务层 ============
s = new_slide(prs)
header(s, "任务层：把提示词写具体", "它能推断意图，但读不了你的心。四种策略，一份简报")
demo_badge(s, 1)
strats = [
    ("限定范围", "“为 foo.py 写一个覆盖‘用户已登出’边界情况的测试。不要用 mock。”"),
    ("指向信息源", "“翻一下这个类的 git 历史，总结它的 API 是怎么演变成这样的。”"),
    ("引用既有范式", "“参照首页既有组件的实现范式来做日历组件，不引入新依赖。”"),
    ("症状+定位+验收", "“会话超时后登录失败。检查 auth 模块的 token 刷新，先写复现的失败测试再修。”"),
]
for i, (t, d) in enumerate(strats):
    x = 64 + i * 296
    rect(s, x, 150, 264, 172, fill=CARD, line=CARD_BR, radius=10)
    tb(s, x + 18, 166, 228, 26, [[(t, dict(size=16.5, bold=True))]])
    tb(s, x + 18, 200, 228, 110, [[(d, dict(size=12.5, color=MUT68))]], line_spacing=1.45)
rect(s, 64, 346, 560, 184, fill=CARD, line=CARD_BR, radius=10)
tb(s, 88, 362, 400, 26, [[("五种方式喂给它富上下文", dict(size=18, bold=True))]])
inputs = ["引用文件（@路径）", "粘贴截图 / 设计稿", "提供文档 URL", "管道喂数据（日志、报错）", "让它自己去取（读文件/跑命令/调工具）"]
for i, t in enumerate(inputs):
    col = i % 2
    row = i // 2
    x = 88 + col * 262
    y = 400 + row * 42
    rect(s, x, y + 8, 8, 8, fill=BLUE, radius=4)
    tb(s, x + 16, y, 240, 36, [[(t, dict(size=13, color=MUT78))]], line_spacing=1.3)
rect(s, 660, 346, 556, 184, fill=CARD, line=CARD_BR, radius=10)
tb(s, 684, 362, 400, 26, [[("任务简报四要素", dict(size=18, bold=True))]])
elems = [("Goal", "要什么：一句话说清终点"), ("Context", "在哪儿：具体路径，指向既有范式"),
         ("Constraints", "不能做什么：依赖、范围、风格"), ("Done-when", "什么叫完成：可运行、可判定")]
for i, (t, d) in enumerate(elems):
    col = i % 2
    row = i // 2
    x = 684 + col * 262
    y = 400 + row * 58
    tb(s, x, y, 250, 22, [[(t, dict(size=14.5, bold=True, color=TEAL, font=MONO))]])
    tb(s, x, y + 24, 250, 30, [[(d, dict(size=12.5, color=MUT72))]], line_spacing=1.35)
bottom_judge(s, 64, 546, 1152, 76, "说得越明确，结果越确定。探索期可以用模糊提示找灵感，比如问它这个文件有哪些可改进；交付期必须给出能判定的验收。")
footer(s, "工程篇")

# ============ P15 能力层：选型 ============
s = new_slide(prs)
header(s, "能力层：五种扩展机制，别选错工具", "要共享内容用技能，要隔离消耗用子智能体。选错是常见的浪费")
rect(s, 64, 148, 730, 478, fill=TBL_BG, line=CARD_BR, radius=10)
tbl_shape = s.shapes.add_table(7, 3, emu(76), emu(158), emu(706), emu(458))
tbl = tbl_shape.table
for i, wd in enumerate([132, 296, 278]):
    tbl.columns[i].width = emu(wd)
data = [
    ["机制", "什么时候用", "例子"],
    ["规则文件", "永远生效的项目约定", "“用 pnpm 不用 npm；提交前跑测试”"],
    ["技能", "可复用的知识 / 工作流", "部署清单、API 规范、代码评审手册"],
    ["子智能体", "读大量文件但只回报结论", "重型调研、专职安全审查"],
    ["外部工具连接", "访问外部数据或动作", "查数据库、发消息、读工单系统"],
    ["事件钩子", "每次必然发生的自动化", "每次文件编辑后自动跑 lint"],
    ["插件", "多仓库复用同一套配置", "把技能+钩子打包成可分发单元"],
]
for r_i in range(7):
    tbl.rows[r_i].height = emu(48 if r_i == 0 else 68)
    for c_i in range(3):
        cell = tbl.cell(r_i, c_i)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = cell.margin_right = emu(8)
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        if r_i == 0:
            set_run(p.add_run(), data[r_i][c_i], 14.5, TXT, bold=True)
            cell.fill.solid(); cell.fill.fore_color.rgb = BLUE_BG
        else:
            color = TEAL if c_i == 0 else (MUT85 if c_i == 1 else MUT68)
            set_run(p.add_run(), data[r_i][c_i], 13.5, color, bold=(c_i == 0))
            cell.fill.solid(); cell.fill.fore_color.rgb = ROW_BG if r_i % 2 else TBL_BG
rect(s, 830, 150, 386, 196, fill=CARD, line=CARD_BR, radius=10)
tb(s, 854, 166, 340, 26, [[("最易混淆的一对", dict(size=17, bold=True))]])
tb(s, 854, 202, 340, 62, [[("技能 = ", dict(size=14, bold=True, color=TEAL)), ("可复用内容，加载进你的主窗口", dict(size=14, color=MUT78))]], line_spacing=1.45)
tb(s, 854, 268, 340, 62, [[("子智能体 = ", dict(size=14, bold=True, color=BLUE)), ("隔离执行者，独立上下文，只回传摘要", dict(size=14, color=MUT78))]], line_spacing=1.45)
judge_card(s, 830, 366, 386, 260, "内容由技能承载，消耗由子智能体隔离。分不清这一对，调研成本就全压在你的主上下文里，窗口就是这么被吃掉的", 16.5)
footer(s, "工程篇")

# ============ P16 能力层：渐进式搭建 ============
s = new_slide(prs)
header(s, "渐进式搭建：让问题触发配置", "你不需要预先配置任何东西。每个机制都有一个能认出来的触发时机")
trigs = [
    ("同一个约定或命令，它错了两次", "写进规则文件"),
    ("你反复输入同一段开场提示词", "存成可直接调用的技能"),
    ("同一份操作手册粘贴了第三次", "固化成技能"),
    ("手动从浏览器复制它看不到的数据", "把那个系统接成外部工具"),
    ("它为找一个符号要读很多文件", "装代码智能插件"),
    ("支线任务灌入大量不再引用的输出", "转给子智能体"),
    ("希望某件事每次自动发生、无需提醒", "写成事件钩子"),
    ("第二个仓库需要同样的配置", "打包成插件分发"),
]
for i, (a, b) in enumerate(trigs):
    y = 148 + i * 56
    rect(s, 64, y, 620, 52, fill=CARD, line=CARD_BR, radius=9)
    tb(s, 88, y, 580, 52, [[("触发 · ", dict(size=13.5, bold=True, color=BLUE)), (a, dict(size=14, color=MUT85))]], anchor=MSO_ANCHOR.MIDDLE)
    arrow(s, 692, y + 26, 724, y + 26, TEAL, 1.75)
    rect(s, 732, y, 300, 52, fill=TEAL_BG, line=TEAL_BR, radius=9)
    tb(s, 752, y, 270, 52, [[(b, dict(size=14, bold=True, color=TEAL))]], anchor=MSO_ANCHOR.MIDDLE)
rect(s, 1060, 148, 156, 444, fill=CARD, line=CARD_BR, radius=10)
tb(s, 1080, 170, 116, 380, [[("配", dict(size=30, bold=True, color=BLUE))], [("置", dict(size=30, bold=True, color=BLUE))],
    [("顺", dict(size=30, bold=True, color=BLUE))], [("序", dict(size=30, bold=True, color=BLUE))],
    [("由", dict(size=30, bold=True, color=TEAL))], [("问", dict(size=30, bold=True, color=TEAL))],
    [("题", dict(size=30, bold=True, color=TEAL))], [("定", dict(size=30, bold=True, color=TEAL))]], line_spacing=1.5)
bottom_judge(s, 64, 610, 1152, 52, "同一个错误反复出现时，去改规则文件，别在对话里再口头纠正一次")
footer(s, "工程篇")

# ============ P17 会话层 ============
s = new_slide(prs)
header(s, "会话层：上下文治理", "清理、压缩、回溯是日常操作，不是急救手段")
rect(s, 64, 150, 520, 214, fill=TEAL_BG, line=TEAL_BR, radius=10)
tb(s, 88, 168, 460, 62, [[("硬规则：同一问题上纠正了两次", dict(size=19, bold=True)), ], [("→ 立刻清理上下文，重开会话", dict(size=19, bold=True, color=TEAL))]], line_spacing=1.45)
tb(s, 88, 244, 470, 100, [[("此时上下文已塞满失败方案。带着学到的东西写一个更好的初始提示词，干净的会话几乎总是优于堆满纠正的长会话", dict(size=14.5, color=TXT))]], line_spacing=1.55)
rect(s, 64, 380, 520, 168, fill=CARD, line=CARD_BR, radius=10)
tb(s, 88, 396, 400, 26, [[("检查点：放心让它试错", dict(size=17, bold=True))]])
tb(s, 88, 430, 470, 100, [[("冒险方案不行就回溯换条路，检查点随会话保存，隔天依然可回溯。⚠ 它只追踪编辑工具产生的改动，不能替代 git", dict(size=14, color=MUT78))]], line_spacing=1.5)
govs = [
    ("主动压缩", "任务之间清理上下文；定向压缩（“聚焦 API 改动”）；把“压缩时必须保留什么”写进规则文件。"),
    ("调研外包", "需要读大量文件的调研交给子智能体，在独立上下文里消耗，只回传摘要。"),
    ("即时纠偏", "发现跑偏立刻中断改向，上下文保留，可以马上换指令。越晚纠偏，成本越高。"),
    ("会话当分支", "每条工作流一个命名会话（如 oauth-migration），跨天继续不用重复解释背景。"),
]
for i, (t, d) in enumerate(govs):
    y = 150 + i * 104
    rect(s, 620, y, 596, 94, fill=CARD, line=CARD_BR, radius=10)
    rect(s, 620, y, 4, 94, fill=BLUE)
    tb(s, 644, y + 12, 300, 26, [[(t, dict(size=16.5, bold=True))]])
    tb(s, 644, y + 42, 548, 46, [[(d, dict(size=13.5, color=MUT72))]], line_spacing=1.45)
bottom_judge(s, 64, 574, 1152, 76, "结果好不好，取决于反馈循环紧不紧，跟第一版提示词写得多漂亮关系不大。发现跑偏立刻纠偏，这是用 Agent 和用搜索引擎的根本区别", 15.5)
footer(s, "工程篇")

# ============ P18 扉页 规模化篇 ============
section_slide(prs, "06", os.path.join(A, "ch_rollout.png"), "g_ch04.png", "规模化篇", "先画笼子，再放大产出",
              "先把一个智能体用得有效，再用并行、无头与流水线放大产出。顺序反了，只会把低效复制 N 份")

# ============ P19 权限与安全 ============
s = new_slide(prs)
header(s, "权限：先画笼子，再放它出去", "每次都弹窗确认，点过第十次你就不是在审查，只是在点")
rect(s, 64, 150, 560, 116, fill=CARD, line=CARD_BR, radius=10)
tb(s, 88, 164, 300, 24, [[("两难", dict(size=16, bold=True, color=BLUE))]])
tb(s, 88, 194, 510, 62, [[("全确认，安全但注意力疲劳，审查形同虚设。全放行，风险失控。你真正有效的注意力是有限的，别消耗在低价值点击上", dict(size=13.5, color=MUT72))]], line_spacing=1.45)
rect(s, 64, 282, 560, 108, fill=CARD, line=CARD_BR, radius=10)
tb(s, 88, 296, 300, 24, [[("两条互补的解法", dict(size=16, bold=True, color=TEAL))]])
tb(s, 88, 326, 510, 56, [[("权限白名单：具名放行你信任的操作 ｜ 沙箱隔离：限制文件系统与网络边界，让它在笼子内自由工作。", dict(size=13.5, color=MUT72))]], line_spacing=1.45)
rect(s, 64, 406, 560, 222, fill=CARD, line=CARD_BR, radius=10)
tb(s, 88, 422, 300, 24, [[("四种权限模式", dict(size=16, bold=True))]])
modes = [("自动复核", "独立模型替你拦风险操作，日常改动不弹窗"),
         ("手动确认", "写文件、跑命令前逐次询问，适合你想亲自批准时"),
         ("只读规划", "只能探索与出方案，不落任何修改，适合探索阶段"),
         ("预批准", "只放行具名工具，其余一律拒绝，适合无人值守批量")]
for i, (a, b) in enumerate(modes):
    y = 456 + i * 42
    tb(s, 88, y, 110, 36, [[(a, dict(size=14, bold=True, color=TEAL))]], anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 206, y, 400, 36, [[(b, dict(size=13, color=MUT72))]], line_spacing=1.35, anchor=MSO_ANCHOR.MIDDLE)
bottoms = [
    ("沙箱隔离", "文件系统与网络必须有边界，不让它在主环境裸奔"),
    ("最小权限", "批量运行前具名声明放行工具，不做全量授权"),
    ("全程可审计", "能回答“它改了什么、执行了什么命令”"),
    ("危险动作人工闸", "数据变更、部署、密钥操作保留人工审批"),
]
for i, (t, d) in enumerate(bottoms):
    y = 150 + i * 108
    rect(s, 660, y, 556, 96, fill=CARD, line=CARD_BR, radius=10)
    rect(s, 660, y, 4, 96, fill=TEAL)
    tb(s, 684, y + 14, 500, 26, [[(t, dict(size=16.5, bold=True))]])
    tb(s, 684, y + 46, 510, 40, [[(d, dict(size=13.5, color=MUT72))]], line_spacing=1.4)
bottom_judge(s, 660, 584, 556, 60, "便利来自权限，风险也来自权限，顺序不能颠倒")
footer(s, "规模化篇")

# ============ P20 规模化路径 ============
s = new_slide(prs)
header(s, "规模化：从一个 Agent 到一条流水线", "先让一个智能体用得有效，再成倍放大。顺序反了，只会复制低效")
demo_badge(s, 1)
paths = [
    ("无头模式", "命令行非交互调用 + 结构化输出（JSON），接进 CI、提交前钩子与自动化脚本。"),
    ("并行会话", "在隔离检出中同时跑多个会话，改动互不覆盖；会话之间可传递发现。"),
    ("扇出迁移", "大迁移拆成批量任务，先用 2-3 个小样试跑、调提示词，再全量。具名放行所需工具"),
    ("无人值守", "自动复核模式下跑长任务；结束时由全新上下文的评审智能体做对抗式检查。"),
    ("流水线集成", "输出 JSON 供下游命令消费，让 Agent 成为数据处理管道中的一环"),
]
for i, (t, d) in enumerate(paths):
    y = 150 + i * 94
    rect(s, 64, y, 600, 84, fill=CARD, line=CARD_BR, radius=10)
    tb(s, 88, y + 12, 160, 28, [[(t, dict(size=16.5, bold=True))]])
    tb(s, 88, y + 44, 552, 36, [[(d, dict(size=13, color=MUT72))]], line_spacing=1.4)
rect(s, 700, 150, 516, 262, fill=CARD, line=CARD_BR, radius=10)
tb(s, 724, 166, 400, 26, [[("Writer / Reviewer 模式", dict(size=18, bold=True))]])
wr = [("会话 A · Writer", "“为 API 端点实现限流器。”", TEAL),
      ("会话 B · Reviewer", "“审查这个实现：边界情况、竞态条件、与既有中间件范式的一致性。”", BLUE),
      ("会话 A · Writer", "“这是评审意见（粘贴 B 的输出），处理这些问题。”", TEAL)]
for i, (who, what, c) in enumerate(wr):
    y = 206 + i * 66
    tb(s, 724, y, 180, 22, [[(who, dict(size=13, bold=True, color=c, font=MONO))]])
    tb(s, 724, y + 24, 468, 40, [[(what, dict(size=13, color=MUT78))]], line_spacing=1.4)
judge_card(s, 700, 428, 516, 200, "全新上下文的评审者不会偏向自己刚写的代码。多会话最大的红利是质量，不是速度。⚠ 你让它找缺口，它几乎必然报出缺口：只标记影响正确性的问题，其余视为可选，警惕过度工程", 15.5)
footer(s, "规模化篇")

# ============ P21 七个失败模式 ============
s = new_slide(prs)
header(s, "七个高频失败模式与修正", "它们反复出现在所有团队，早识别，省的是一下午")
rect(s, 64, 148, 768, 478, fill=TBL_BG, line=CARD_BR, radius=10)
tbl_shape = s.shapes.add_table(8, 2, emu(76), emu(158), emu(744), emu(458))
tbl = tbl_shape.table
tbl.columns[0].width = emu(452)
tbl.columns[1].width = emu(292)
data = [
    ["失败模式", "修正"],
    ["1 · 大杂烩会话：任务串味，上下文塞满无关信息", "无关任务之间，清理上下文"],
    ["2 · 反复纠正：改了两次还是错", "清理重开，重写初始提示词"],
    ["3 · 规则文件过长：重要规则淹没在噪声里", "无情修剪；能自动化的改写成钩子"],
    ["4 · 信任-验证缺口：看似合理，漏掉边界情况", "无法验证的东西，不上线"],
    ["5 · 无限探索：一次调研读了上百个文件", "收窄范围，或交给子智能体"],
    ["6 · 规格未固化就开工：口述需求、边做边改", "先写 SPEC，再开新会话执行"],
    ["7 · 批量任务权限全放行：为了省事", "具名放行 + 沙箱 + 预批准模式"],
]
for r_i in range(8):
    tbl.rows[r_i].height = emu(46 if r_i == 0 else 58)
    for c_i in range(2):
        cell = tbl.cell(r_i, c_i)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = cell.margin_right = emu(8)
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        if r_i == 0:
            set_run(p.add_run(), data[r_i][c_i], 14.5, TXT, bold=True)
            cell.fill.solid(); cell.fill.fore_color.rgb = BLUE_BG
        else:
            color = MUT85 if c_i == 0 else TEAL
            set_run(p.add_run(), data[r_i][c_i], 13.5, color, bold=(c_i == 1))
            cell.fill.solid(); cell.fill.fore_color.rgb = ROW_BG if r_i % 2 else TBL_BG
rect(s, 864, 148, 352, 300, fill=CARD, line=CARD_BR, radius=10)
tb(s, 888, 166, 300, 26, [[("表现不好时，按序自查", dict(size=17, bold=True))]])
diag = ["① 上下文是否太脏？", "② 提示词是否太糊？", "③ 任务是否大到不该一次跑完？", "④ 验证信号是否存在？"]
for i, t in enumerate(diag):
    y = 210 + i * 48
    tb(s, 888, y, 300, 40, [[(t, dict(size=14.5, color=MUT85))]], line_spacing=1.4)
tb(s, 888, 404, 300, 36, [[("八成问题落在其中一条。", dict(size=14, color=TEAL, bold=True))]])
judge_card(s, 864, 470, 352, 168, "这些不是工具不行，是工作方法问题。方法对了，换一个 Agent 效果一样", 16)
footer(s, "规模化篇")

# ============ P22 落地路线 ============
s = new_slide(prs)
header(s, "团队落地：90 天，三个阶段", "先统一规则文件，再谈覆盖率。顺序不能颠倒")
phases = [
    ("第 1 个 30 天", "建规范", "规则文件 < 100 行 + 3 个高频技能 + 2 个确定性钩子。判定：新人照规则文件能跑通一次完整改动。"),
    ("第 2 个 60 天", "跑试点", "1-2 个真实需求的端到端案例（含 SPEC、diff、验证输出、评审记录）。判定：复盘能说清哪里省了时间、哪里把时间还回去了。"),
    ("第 3 个 90 天", "进流水线", "CI 中的无头调用 + 提交前钩子 + 权限白名单基线。判定：无人值守任务有失败告警，且有回滚路径。"),
]
for i, (d, t, desc) in enumerate(phases):
    y = 150 + i * 152
    rect(s, 64, y, 640, 138, fill=CARD, line=CARD_BR, radius=10)
    rect(s, 64, y, 4, 138, fill=TEAL if i == 0 else BLUE)
    tb(s, 92, y + 16, 150, 24, [[(d, dict(size=14, color=TEAL if i == 0 else BLUE, font=MONO))]])
    tb(s, 240, y + 12, 120, 30, [[(t, dict(size=19, bold=True))]])
    tb(s, 92, y + 52, 588, 74, [[(desc, dict(size=13.5, color=MUT72))]], line_spacing=1.5)
metrics = [
    ("采纳率", "实际使用比例和频次。不好用的东西推广不动，这是前置指标"),
    ("返工率", "产出被人工推翻或重写的比例，直接反映上下文质量"),
    ("缺陷密度", "单位改动引入的缺陷数，防止用速度换质量的自欺"),
    ("端到端时长", "需求提出到合入的墙钟时间，唯一能证明整体更快的指标"),
]
for i, (t, d) in enumerate(metrics):
    y = 150 + i * 96
    rect(s, 740, y, 476, 86, fill=CARD, line=CARD_BR, radius=10)
    tb(s, 764, y + 12, 120, 28, [[(t, dict(size=16.5, bold=True, color=TEAL))]])
    tb(s, 764, y + 44, 428, 36, [[(d, dict(size=13, color=MUT72))]], line_spacing=1.4)
judge_card(s, 740, 538, 476, 134, "只统计生成了多少代码，会奖励错误行为。多写代码从来不是工程目标", 16)
footer(s, "规模化篇")

# ============ P30 结束页 (Hero) ============
s = new_slide(prs)
tb(s, 64, 96, 300, 120, [[("“", dict(size=120, bold=True, color=TEAL, font=MONO))]])
tb(s, 64, 190, 1080, 160, [[("这套方法最终要做的，", dict(size=34, bold=True))],
                           [("是把团队的工程默契从人的脑子里搬进仓库。", dict(size=34, bold=True, color=TEAL))]], line_spacing=1.55)
hline(s, 64, 372, 72, TEAL, 3)
tb(s, 64, 396, 600, 28, [[("本周就能开始的三件事", dict(size=20, bold=True))]])
acts = [
    ("今天", "为主仓库写一份规则文件，删到 100 行以内，纳入 git。"),
    ("本周", "用任务简报四要素跑一个真实需求，必须带一个能跑出 pass/fail 的验证命令"),
    ("本周", "把一个重复了三次以上的操作，固化成可复用的技能。"),
]
for i, (tag, d) in enumerate(acts):
    x = 64 + i * 384
    rect(s, x, 436, 368, 132, fill=CARD, line=CARD_BR, radius=10)
    tb(s, x + 22, 452, 120, 24, [[(tag, dict(size=14, bold=True, color=TEAL, font=MONO))]])
    tb(s, x + 22, 484, 324, 72, [[(d, dict(size=14, color=MUT78))]], line_spacing=1.5)
tb(s, 64, 604, 1152, 44, [[("什么时候可以不守这些规则：深陷复杂问题时让上下文累积，探索性任务跳过规划，想先看它怎么解释问题时用模糊提示。直觉比指南可靠，直觉是从实践里长出来的", dict(size=13.5, color=MUT55))]], line_spacing=1.5)
tb(s, 980, 676, 300, 22, [[("ENGINEERING PRACTICES · 2026.10", dict(size=12, color=MUT50, font=MONO, spacing=2))]], align=PP_ALIGN.RIGHT)

# ============ 交叉引用自动回填（@@页面标题 -> Pxx） ============
TITLE_PT = 36 * 0.75  # set_run() 统一乘 0.75，故 header 的 size=36 落地为 27pt


def build_title_index():
    """按页标题字号建立「标题 → 页码」索引。"""
    idx = {}
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            hit = False
            for para in sh.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.size and abs(run.font.size.pt - TITLE_PT) < 0.6:
                        t = sh.text_frame.text.strip()
                        if t and t not in idx:
                            idx[t] = i
                        hit = True
                        break
                if hit:
                    break
    return idx


def fill_cross_refs(titles):
    """把正文里的 @@标题 占位替换成真实页码——页面顺序再变也不用改引用。"""
    def repl(m):
        key = m.group(1).strip()
        if key not in titles:
            print("  !! unresolved ref:", repr(key))
            return m.group(0)
        return "P%02d " % titles[key]
    n = 0
    for s in prs.slides:
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            for para in sh.text_frame.paragraphs:
                for run in para.runs:
                    if "@@" in run.text:
                        run.text = re.sub(r"@@([^@·]+)", repl, run.text)
                        n += 1
    print("cross refs filled:", n)


TITLE_INDEX = build_title_index()
fill_cross_refs(TITLE_INDEX)

# ============ 目录页范围自动回填 ============
def fill_toc_ranges():
    """按每页页脚的章节名归并连续区间，回填目录页 Pxx - Pxx——增删页面后无需手工重排。"""
    raw = []
    for i, s in enumerate(prs.slides, 1):
        ch, is_cover = None, False
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            t = sh.text_frame.text
            if t.startswith("Coding Agent 工程实践 · "):
                ch = t.split("· ", 1)[1].strip()
            if t.startswith("CHAPTER "):
                is_cover = True
        raw.append((i, ch, is_cover))
    # 章节扉页没有页脚，归属给它后面那一章
    seq = []
    for k, (i, ch, is_cover) in enumerate(raw):
        if ch:
            seq.append((ch, i))
        elif is_cover:
            nxt = next((c for _, c, _ in raw[k + 1:] if c), None)
            if nxt:
                seq.append((nxt, i))
    ranges = []
    for ch, pg in seq:
        if ranges and ranges[-1][0] == ch:
            ranges[-1][2] = pg
        else:
            ranges.append([ch, pg, pg])
    ranged = [r for r in ranges if r[0] != "目录"]
    idx = 0
    for sh in prs.slides[1].shapes:
        if not sh.has_text_frame or idx >= len(ranged):
            continue
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                if run.text.strip().startswith("P") and " - " in run.text:
                    _, a, b = ranged[idx]
                    run.text = "P%02d - P%02d" % (a, b)
                    idx += 1
    print("toc ranges filled:", idx, [r[0] for r in ranged])

fill_toc_ranges()

# ============ 演讲者备注：给讲师读的话一律放在这里，投影上看不到 ============
PAGE_NOTES = {
    "@封面": """开场钩子（两问，30 秒）：
在座多少人已经在用 AI 写代码？——再问：多少人敢让它直接改一个横跨五个文件的真实需求？这个落差，就是今天要补的东西。

三件事先说明白：
① 不绑定工具：今天讲的方法，换任何一个 Agent 原样生效；
② 结构：第 1 章定工具，第 2 章看真实效果，后四章讲方法；
③ 随时打断提问。

时间分配参考：第 1 章 15 分钟、第 2 章 10 分钟、第 3 章 10 分钟、第 4 章 25 分钟、第 5 章 25 分钟、第 6 章 15 分钟。时间紧就砍第 6 章最后两页。""",

    "目录": """这一页别逐条念，停 20 秒。

主线一句话：先看有哪些工具 → 相信它真能做到 → 理解那条约束 → 单次任务的标准动作 → 上下文怎么搭 → 团队怎么起来。

如果有人问“为什么不直接讲技巧”：技巧在第 4、5 章，但不先理解第 3 章那条约束，技巧只能靠背。""",

    "@CH01": """这一章的目的不是让大家记住市占率，是接受两件事：
① 排行榜半年换一遍，别把选型押在当期第一名；
② 选型看的是「形态」与「合规边界」，不是品牌。

提醒：本章不可避免地出现产品名（选型避不开），从第 3 章起不再出现任何产品名。""",

    "全球格局：一个已经重排过的市场": """数字只念两个：90% 的人每周在用；七个月里两个翻倍、两个下滑（Claude Code 18%→39%，Codex 3%→16%；Copilot 29%→21%，Cursor 18%→12%）。

其余让听众自己看条的长度，不要逐个念。

有人追问时的补充：
DeepSeek Harness，2026-08 开源，MIT，发布两周 GitHub 破 3 万星；Pi，2026-10 发布 1.0，MIT，极简终端智能体。两者都未进该调研口径。

出处：JetBrains《Developer Ecosystem Survey 2026》，15,000+ 职业开发者，2026-05 至 07 田野期。""",

    "国内格局：两种口径，且榜单会漏掉新人": """数据口径先声明：企业侧份额是 IDC《中国 AI 编程市场》报告的二手引用，个人侧是媒体引用口径。课前建议核一遍原始报告再讲。

品牌注（有人问起就这么答）：ZCode 是智谱当前主推的编程智能体，CodeGeeX 是其早期的 IDE 插件品牌。

这一页真正值钱的观察：MiniMax Code 与 MiMo Code 都直接基于开源 OpenCode 构建——连模型厂商都不自己写 harness。这句话直接引出下一页。""",

    "开源路线：Harness 是可替换的那一层": """演示 ⑧：dsh › 液体处理平台 › main。可以在这里现场切过去跑一段：同一个模型，换一个 harness 就是另一个 Agent。

⚠ 课前必做：dsh 目前是 v0.1 预览版，接口有破坏性变更。请提前确认现场环境能正常跑，并准备一个备选 demo，不要在课上翻车。

留给观众的一句：模型是大脑，harness 是执行神经系统。""",

    "五类形态：先认形态，再看品牌": """第 1 章的结论页之一。

最容易讲错的地方：开源 / 自托管不是第六种形态，是可以叠加在任意形态上的另一条线——中间那条横带就是在提醒这个。

现场互动：问“你们手上的工具属于哪一列？”答案通常集中在 IDE 插件和命令行两列，正好接下一页的三步漏斗。""",

    "怎么选：三步，一步一句话": """按顺序讲三步，别跳。

第 03 步「跑满两周」值得停顿：第一天的新鲜感会骗人，很多团队第一天就下结论，第二周推翻重选。

收尾强调：工具统一的产出不是“选到最好的”，是“大家用同一个”——规则文件、技能库、钩子配置要沉淀在同一套工具上。

衔接下一章：选定了工具，先看它在真实项目里做到过什么。""",

    "@CH02": """这一章只有一张表，作用是建立信任，不要在这里讲方法。

节奏：每段会话 60–90 秒，说清“当时要做什么”和“结果怎样”，点到为止。方法是后面几章的事。

建议从 ① 开场（最朴素的用法，先立住印象），至少现场演示 ②（先出方案）和 ④（换上下文挑刺）。如果只够一段，选 ④——反差最大。""",

    "八段真实会话：从最朴素的用法，到最严格的用法": (
        "会话坐标（页面上刻意不写，避免把内部项目信息投影出去）。"
        "顺序由表格自动生成，与表格逐行一致——从上往下就是讲课顺序：\n"
        + "\n".join("%s %s → %s" % (no, re.split(r"[：，]", what)[0], loc) for no, what, _m, loc in demos)
        + "\n\n① 是基线（最朴素的用法，先立住这个印象）；②③④ 连着在同一个通信协议库里讲完，不用中途切仓库；⑤⑥⑦ 在第二个项目；⑧ 在开源底座上。"
    ),

    "@CH03": """这一章回答“为什么这套方法有效”，而不是碰巧有用。讲慢一点，它是后面所有内容的父节点。

听众里“用过”和“没用过”差别很大时，这一章对没用过的人更重要——它决定了他们会不会把后面的方法当成玄学。""",

    "一个约束：上下文窗口是稀缺资源": """用 3–4 分钟讲，这是全场唯一的“物理规律”。

三个现象用自己的经历印证，别照读：铺垫一句“有没有遇到过它越改越蠢、明明说过的约束它又违反？”——那就是窗口满了。

推论一、推论二讲的是怎么省上下文，推论三讲的是怎么把上下文花对地方。三条推论都是后面所有做法的父节点，第 5 章会一层层展开。这里只留印象。

【推论三怎么讲】页面这句已经够直白，照读即可，重点放在例子上：
它做错了，你再解释一遍是最贵的做法——话说得越来越多，窗口越来越满，可它到底对不对，你仍然不知道。不如让它跑一次测试或编译，红就是红、绿就是绿，这个结论不用你描述，它自己会读。

现场举例（任选一个，别两个都讲）：
· 让它改完自己跑单测，把失败的 3 条原样贴回去，比你写三段“注意这里要怎样”有效得多。
· 编译/类型报错：把报错原文丢给它，比你复述“大概是那个类型转换的问题”有效得多。

一句话收口：解释是花钱买它的理解，验证是花钱买一个能判定的结论——同样消耗上下文，后者划算。

常见质疑：“我就多说两句怎么了？”——一次两次没事，但这是每天几十次的习惯，而且解释到第三遍时，确定性基本不再上涨。""",

    "工作方式变了：从帮你写到替你做": """核心句式让听众复述一遍：过去我在执行环节，现在我在边界环节和验收环节。

最容易被挑战：有人说“它写得没我快”。用下方那句回应——评价交付工具要看是否省掉了整个下午，不是看打字速度。

如果现场有管理者，这是他们最需要转变的一页。""",

    "@CH04": """从这一章开始进实操，时间给够（建议 25 分钟）。每页都留一个“下周就能用”的动作。

四页顺序：先给可运行的验证 → 把提示词改具体 → 单次任务的四步动作 → 大特性的采访式起手。""",

    "第一件事：给它一个可运行的验证": """演示 ③：Zcode › StdCommuProtocol › minor · 消息 1–6（改完让它自己回头查一遍）。

一句话概括本页：别让你自己成为验证回路。

一定要强调“它的检查通过 ≠ 你的验收通过”，两者是不同层级，上线前仍需端到端确认——否则有人会把 AI 的自检当上线依据。

四级检查不必全用：L1 今天就能用，L4 是无人值守的前提。""",

    "提示词改写：三组实战对照": """演示 ⑥：Zcode › V3 › main · 拉到底后第 3–4 条（给确定的设计图，而不是一句“做得好看点”）。

三组对照，一组 40 秒。念“改写前”时带点调侃，听众会笑着认出自己。

把自检句式留给观众：把“完成定义”盖住，如果你判断不了它算不算完成，它就不够具体。""",

    "单次任务的标准动作": """演示 ②⑤⑦：
② Zcode › StdCommuProtocol › minor · 倒数第 2–1 条
⑤ Zcode › V3 › main · 倒数第 13–10 条
⑦ Zcode › V3 › minor · 拉到底后第 5–9 条

四步里最常被跳过的是「探索」——多数人第一句话就是“改一下这个函数”。

右下角「何时跳过规划」一定要讲，否则听众会觉得这套流程太重、落不了地。""",

    "大特性起手式：先让它采访你": """这一页现场效果最好。把那句邀请直接投出来让大家抄：
“我想构建 X。请详细采访我：技术实现、UI/UX、边界情况、顾虑与权衡。不问显而易见的，深挖我可能没考虑到的。”

提醒：采访产出的规格必须落成文件。口头聊完直接开干等于没做，而且会同时污染上下文和代码。

合格的规格三条：自包含、有边界、可验收。""",

    "@CH05": """这一章是这门手艺的主体，建议 25 分钟以上，自底向上讲：规则 → 任务 → 能力 → 会话。

开场锚点（下一页）那张五层图必须讲透，后面四页是逐层展开。""",

    "上下文工程：这门手艺的主体": """锚点页。

模型层已在第 3 章讲过，这里一句话带过，不要展开——听众会以为还要讲模型选型。

留给观众记的一句：模型的差距是线性的，上下文结构的差距是乘性的。

②③④ 三层是团队真正能动的，下面四页分别讲。""",

    "规则层：把默契写进仓库": """全场最容易立刻执行的一页：今天就为自己的主仓库写一份，控制在 100 行以内。

让听众记住这句：每一行都问一遍“删掉这行会让它犯错吗？”答不上来就删。

提前点破误区：把代码规范、API 文档、逐文件描述全塞进去——那会把真正重要的指令淹没。

右侧排错指南可以现场用：已经写了规则文件还踩坑的，按这三行自查。""",

    "任务层：把提示词写具体": """演示 ⑤：Zcode › V3 › main · 倒数第 13–10 条（同事改了同一个项目，问题不一定是 AI 引入的）。

四种策略各给一个例子就够。重点是右下角「任务简报四要素」——这是听众能直接照着写的模板。

别讲成一刀切：探索期可以用模糊提示找灵感，交付期必须给出可判定的验收。""",

    "能力层：五种扩展机制，别选错工具": """一句话区分：技能 = 共享内容（加载进主窗口），子智能体 = 隔离消耗（独立上下文，只回传摘要）。

选错的代价要讲清楚：把重型调研直接放进主对话，窗口就是这么被悄悄吃掉的——很多人不知道自己为什么中途变蠢了。

五种机制的细节留给资料，台上只讲这一对。""",

    "渐进式搭建：让问题触发配置": """这一页回答“我是不是得先把这些全配好”——不用。每个机制都有一个可识别的触发信号。

最好用的一条：同一段开场提示词你打第三遍了 → 存成技能。几乎人人都有过。

右侧竖排那句是本页主旨：配置顺序由问题决定，先有问题，再有配置。""",

    "会话层：上下文治理": """硬规则必须讲：同一问题上纠正了两次 → 立刻清理上下文重开会话。很多人舍不得丢那段历史。

⚠ 必须强调：检查点只追踪编辑工具产生的改动，不能替代 git。否则真有人拿它当版本管理。

「会话当分支」是最容易被忽略的好习惯：每条工作流一个命名会话（如 oauth-migration），跨天继续不用重讲背景。""",

    "@CH06": """最后一章回答：怎么从“一个人在用”走到“整个团队在用”。顺序不能反——先画笼子，再放大产出。

听众里没有管理者可以压到 10 分钟；有管理者在场就按 15 分钟讲，落地路线那页多停一会。""",

    "权限：先画笼子，再放它出去": """开场那个场景要演出来：“每次都弹窗确认，点过第十次你就不再是审查，只是在点。”很多人会笑，因为那就是他们。

四种模式不必全讲，挑贵公司实际会用的两种展开。

沙箱隔离那四条底线是安全 / 合规同事唯一真正关心的，如果他们到场，这里多停两分钟。""",

    "规模化：从一个 Agent 到一条流水线": """演示 ④：Zcode › StdCommuProtocol › main · 消息 3–9（换一个全新上下文做对抗式审查）。

Writer / Reviewer 那段强烈建议现场演：同一个需求开两个会话，一个写一个审，效果最直观。

⚠ 提醒防过度工程：被命令“找缺口”的评审者几乎必然报出缺口。要明确只标记影响正确性的问题，其余视为可选。""",

    "七个高频失败模式与修正": """左边七条不必逐条念，停 5 秒让大家自己扫。

真正要讲是右边那四条自查顺序——七条里八成问题落在其中一条。

态度问题：把这些失败归到“工作方法”，而不是“工具不行”。这是听众今天要带走的立场。""",

    "团队落地：90 天，三个阶段": """这页是给 leader 看的。

每个阶段的「判定」比「时间表」重要：判定不过，就不要进下一阶段。

四个指标重点讲最后一个（端到端时长）：唯一能证明整体更快的指标；“生成了多少代码”会奖励错误行为。""",

    "@结束": """收尾不要讲新东西。

重复那句核心：方法对了，换一个 Agent 也是同样的效果——今天讲的没有一句绑定工具。

让大家把「本周就能开始的三件事」真的写下来，不是听完就算。

留 5 分钟提问。常见提问预案：
① “公司不让代码出网怎么办” → 第 6 页开源路线；
② “我们已经在用某某工具” → 第 8 页三步漏斗；
③ “没人愿意写规则文件” → 第 21 页，先从排错那三条开始收。""",
}


def fill_notes():
    """把讲师要读的话写进备注栏——按页面主标题匹配，增删页面后不会错位。"""
    n = 0
    for i, s in enumerate(prs.slides, 1):
        key = None
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            t = sh.text_frame.text.strip()
            if t.startswith("CHAPTER "):
                key = "@CH" + t.split()[1]
                break
            if t.startswith("UNIVERSAL METHODOLOGY"):
                key = "@封面"
                break
            if t.startswith("本周就能开始的三件事"):
                key = "@结束"
                break
        if key is None:  # 内容页：用页标题（唯一字号 27pt）
            for sh in s.shapes:
                if not sh.has_text_frame:
                    continue
                for para in sh.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.size and abs(run.font.size.pt - TITLE_PT) < 0.6:
                            key = sh.text_frame.text.strip()
                            break
                    if key:
                        break
                if key:
                    break
        if key in PAGE_NOTES:
            s.notes_slide.notes_text_frame.text = PAGE_NOTES[key]
            n += 1
        else:
            print("  (no notes for page %d: %r)" % (i, key))
    print("speaker notes filled:", n)


def check_overflow():
    """自检：结论卡的文字按雅黑度量估算行数，超出卡片高度就报警。"""
    try:
        from PIL import ImageFont
    except Exception:
        return
    fp = "C:/Windows/Fonts/msyh.ttc"
    if not os.path.exists(fp):
        return
    src = open(os.path.abspath(__file__), encoding="utf-8").read()
    cache = {}

    def font(pt):
        k = int(round(pt * 96 / 72))
        if k not in cache:
            cache[k] = ImageFont.truetype(fp, k)
        return cache[k]

    def lines_of(text, size, boxw):
        f = font(size)
        n, cur = 1, 0.0
        for ch in text:
            w = f.getlength(ch)
            if cur + w > boxw:
                n += 1
                cur = w
            else:
                cur += w
        return n

    bad = 0
    for fn, padw, padh, ls, dsz in [("judge_card", 48, 58, 1.6, 17.5),
                                    ("bottom_judge", 132, 16, 1.45, 15.5)]:
        pat = fn + r'\(\s*\w+\s*,\s*(\d+),\s*(\d+),\s*(\d+),\s*(\d+),\s*"([^"]+)"\s*(?:,\s*([\d.]+))?\s*\)'
        for m in re.finditer(pat, src):
            y, w, h, text = int(m.group(2)), int(m.group(3)), int(m.group(4)), m.group(5)
            sz = float(m.group(6)) if m.group(6) else dsz
            pt = sz * 0.75
            n = lines_of(text, pt, w - padw)
            need = n * pt * 1.32 * ls * 96 / 72
            box = h - padh
            if need > box:
                bad += 1
                print("  [OVERFLOW] %s y=%d 需要 %.0fpx / 可用 %.0fpx（%d 行）— %s"
                      % (fn, y, need, box, n, text[:24]))
    print("judge overflow check:", bad, "处溢出" if bad else "clean")


fill_notes()
check_overflow()

prs.save(OUT)
print("saved:", OUT, os.path.getsize(OUT), "bytes, slides:", len(prs.slides.__iter__.__self__._sldIdLst))
