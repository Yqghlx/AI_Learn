# -*- coding: utf-8 -*-
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
TOTAL = 23


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


def footer(slide, chapter, page):
    rect(slide, 64, 688, 12, 12, fill=TEAL_BG2, line=TEAL_BR, radius=2)
    tb(slide, 88, 684, 700, 22, [[("Coding Agent 工程实践 · " + chapter, dict(size=14, color=MUT50))]])
    tb(slide, 1000, 684, 216, 22, [[(page, dict(size=14, color=TEAL, font=MONO))]], align=PP_ALIGN.RIGHT)


def judge_card(slide, x, y, w, h, text, size=17.5):
    rect(slide, x, y, w, h, fill=TEAL_BG, line=TEAL_BR, radius=10)
    rect(slide, x + 24, y + 20, 10, 10, fill=TEAL, radius=5)
    tb(slide, x + 44, y + 14, 120, 22, [[("本课判断", dict(size=16, bold=True, color=TEAL))]])
    tb(slide, x + 24, y + 44, w - 48, h - 58, [[(text, dict(size=size, color=TXT))]], line_spacing=1.6)


def bottom_judge(slide, x, y, w, h, text, size=15.5):
    rect(slide, x, y, w, h, fill=TEAL_BG, line=TEAL_BR, radius=10)
    rect(slide, x + 22, y + h / 2 - 5, 10, 10, fill=TEAL, radius=5)
    tb(slide, x + 42, y, 80, h, [[("判断", dict(size=15, bold=True, color=TEAL))]], anchor=MSO_ANCHOR.MIDDLE)
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
tb(s, 230, 637, 120, 22, [[("2026.10", dict(size=14, color=MUT55, font=MONO))]])
rect(s, 980, 650, 300, 70, fill=BG)
tb(s, 1000, 676, 260, 22, [[("BEST PRACTICES 101", dict(size=13, color=MUT50, font=MONO, spacing=2))]], align=PP_ALIGN.RIGHT)

# ============ P02 目录 ============
s = new_slide(prs)
tb(s, 64, 30, 400, 46, [[("目录", dict(size=36, bold=True))]])
hline(s, 64, 82, 56, TEAL, 3)
rect(s, 64, 160, 6, 120, fill=TEAL)
tb(s, 64, 300, 300, 80, [[("四个问题，", dict(size=24, bold=True))], [("一条方法论主线", dict(size=24, bold=True))]], line_spacing=1.4)
tb(s, 64, 392, 300, 130, [[("不绑定任何一款工具——换一个 Agent，这套方法原样生效。每一章回答一个你下周就会遇到的问题。", dict(size=16, color=MUT55))]], line_spacing=1.7)
toc = [
    ("01", "认知篇 · 一个约束与一种心智", "为什么这套方法有效，而不是碰巧有用？", "P03 - P05"),
    ("02", "方法篇 · 单次任务的标准动作", "一个需求从提出到合入，每步该怎么做？", "P06 - P10"),
    ("03", "工程篇 · 上下文工程四层", "产出质量的差距，到底从哪一层来？", "P11 - P17"),
    ("04", "规模化篇 · 权限、并行与落地", "如何从一个人快，走到整个团队快？", "P18 - P22"),
]
for i, (no, title, q, pg) in enumerate(toc):
    y = 140 + i * 124
    rect(s, 420, y, 796, 110, fill=CARD, line=CARD_BR, radius=10)
    tb(s, 448, y + 30, 70, 46, [[(no, dict(size=30, bold=True, color=TEAL if i == 2 else BLUE, font=MONO))]])
    tb(s, 530, y + 22, 560, 32, [[(title, dict(size=22, bold=True))]])
    tb(s, 530, y + 60, 560, 26, [[(q, dict(size=16, color=MUT55))]])
    tb(s, 1090, y + 44, 110, 24, [[(pg, dict(size=14, color=MUT50, font=MONO))]], align=PP_ALIGN.RIGHT)
footer(s, "目录", "02 / %d" % TOTAL)

# ============ P03 扉页 认知篇 ============
section_slide(prs, "01", os.path.join(A, "ch_cognition.png"), "g_ch01.png", "认知篇", "先理解约束，再使用方法",
              "方法论不解释“为什么”，就只能靠记忆。理解这一个约束，后面所有做法都是它的推论——你不需要背清单。")

# ============ P04 第一性原理 ============
s = new_slide(prs)
header(s, "一个约束：上下文窗口是稀缺资源", "它是唯一必须持续管理的东西——几乎所有最佳实践都从这条事实推导出来")
facts = [
    ("现象 01", "消耗极快", "一次调试会话或一轮代码库探索，就可能产出并消耗数万个 token——上下文不是慢慢用的，是几步就能烧掉一大截。"),
    ("现象 02", "开始“遗忘”", "窗口将满时，模型开始遗忘你在第一句写下的指令——你以为它记得的约束，可能已经不起作用了。"),
    ("现象 03", "犯错率上升", "表现为“它明明知道却还是写错了”——其实它已经不“知道”了。这不是模型变笨，是窗口太满。"),
]
for i, (tag, t, d) in enumerate(facts):
    x = 64 + i * 384
    rect(s, x, 150, 368, 136, fill=CARD, line=CARD_BR, radius=10)
    tb(s, x + 22, 164, 200, 22, [[(tag, dict(size=13, color=BLUE, font=MONO, spacing=2))]])
    tb(s, x + 22, 188, 320, 28, [[(t, dict(size=18, bold=True))]])
    tb(s, x + 22, 220, 324, 60, [[(d, dict(size=13, color=MUT68))]], line_spacing=1.45)
coro = [
    ("推论一 · 能移出去的都移出去", "长期约定进规则文件，偶发知识进技能，重型调研交给子智能体——主对话只留本次任务。"),
    ("推论二 · 一次会话一类问题", "堆积的无关历史不是背景资料，是会主动降低性能的噪声。任务串味时，立刻开新会话。"),
    ("推论三 · 不确定性靠结构消解", "只堆更多解释，等于花最贵的资源买边际极低的确定性。真正有效的是可判定的信号。"),
]
for i, (t, d) in enumerate(coro):
    y = 310 + i * 98
    rect(s, 64, y, 740, 88, fill=BLUE_BG, line=BLUE_BR, radius=10)
    tb(s, 88, y + 14, 700, 26, [[(t, dict(size=17, bold=True))]])
    tb(s, 88, y + 44, 692, 40, [[(d, dict(size=14, color=MUT78))]], line_spacing=1.45)
rect(s, 830, 310, 4, 88, fill=BLUE)
tb(s, 848, 306, 368, 100, [[("上下文窗口装着整场对话：你的每条消息、它读过的每个文件、每条命令的输出。", dict(size=15.5, color=MUT72))]], line_spacing=1.6)
judge_card(s, 830, 432, 386, 192, "看不见的约束无法被管理——先让上下文占用始终可见（状态栏、用量显示），再谈优化。这和性能工程里“不上监控不谈调优”是同一回事。", 16.5)
footer(s, "认知篇", "04 / %d" % TOTAL)

# ============ P05 心智模型 ============
s = new_slide(prs)
header(s, "一种心智：从“帮你写”到“替你做”", "“自主”不等于不需要人——它意味着人的工作位置变了")
stages = [
    ("阶段一", "代码补全", "人在写，AI 猜下一个词。Tab 键就是全部交互，出错成本极低。", False),
    ("阶段二", "对话改写", "人规划，AI 生成整段。选中-提问-粘贴，心智仍是“我来改”。", False),
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
judge_card(s, 780, 518, 436, 142, "站在执行环节评价交付工具，是最常见的认知错位——用“写得不如我快”去衡量一个该用“是否省掉了整个下午”来判断的东西。", 15.5)
footer(s, "认知篇", "05 / %d" % TOTAL)

# ============ P06 扉页 方法篇 ============
section_slide(prs, "02", os.path.join(A, "ch_landscape.png"), "g_ch02.png", "方法篇", "一次需求，四个标准动作",
              "给可运行的验证、先探索再编码、让它采访你、规格先行——四件事，把“玄学体验”变成“稳定产出”。")

# ============ P07 第一条纪律 (Hero) ============
s = new_slide(prs)
header(s, "第一条纪律：给它一个可运行的验证", "没有验证手段时，“看起来做完了”是唯一信号——于是你本人成了验证回路")
rect(s, 64, 150, 560, 118, fill=CARD, line=CARD_BR, radius=10)
tb(s, 88, 164, 200, 24, [[("没有验证", dict(size=17, bold=True, color=BLUE))]])
tb(s, 88, 194, 510, 66, [[("它只能以“看起来像”为准。每个错误都等你发现——你把验证从机器搬回了自己身上，而且是串行的。", dict(size=14.5, color=MUT72))]], line_spacing=1.5)
arrow(s, 344, 272, 344, 294, TEAL, 2)
rect(s, 64, 296, 560, 126, fill=TEAL_BG, line=TEAL_BR, radius=10)
tb(s, 88, 310, 440, 24, [[("给它一个能读出 pass / fail 的信号", dict(size=17, bold=True, color=TEAL))]])
tb(s, 88, 340, 510, 70, [[("测试套件、构建退出码、Linter、输出对比脚本、UI 截图比对。它做、它跑、它读结果、它继续改——循环自己闭合。", dict(size=14.5, color=TXT))]], line_spacing=1.5)
tb(s, 64, 436, 560, 44, [[("注意：它的检查通过 ≠ 你的验收通过——两者是不同层级，上线前你仍需端到端确认。", dict(size=14, color=MUT65))]], line_spacing=1.5)
judge_card(s, 64, 494, 560, 134, "要证据，不要声明：完成后必须附上命令原文与原始输出（UI 改动附截图）。审证据比重跑验证快得多，且对你没在旁观的会话同样有效。", 15.5)
tb(s, 660, 142, 200, 110, [[("4", dict(size=96, bold=True, color=TEAL, font=MONO))]])
tb(s, 770, 176, 446, 60, [[("级验证闸门", dict(size=22, bold=True))], [("逐级降低你的注意力消耗", dict(size=14.5, color=MUT55))]], line_spacing=1.4)
gates = [
    ("L1 · 单次提示", "同一条消息里要求“跑检查并迭代”——任何任务，今天就能用。"),
    ("L2 · 会话级目标", "由独立评估器每轮复核目标是否达成，未达成则继续——中长任务。"),
    ("L3 · 确定性闸门", "用脚本钩子在轮次结束前强制检查，不通过不许停——零例外的规范。"),
    ("L4 · 第三方复核", "全新上下文的评审智能体：做题的人，不是打分的人——无人值守。"),
]
for i, (t, d) in enumerate(gates):
    y = 262 + i * 88
    rect(s, 660, y, 556, 80, fill=CARD, line=CARD_BR, radius=9)
    rect(s, 660, y, 4, 80, fill=TEAL if i >= 2 else BLUE)
    tb(s, 686, y + 12, 500, 24, [[(t, dict(size=16, bold=True, color=TEAL if i >= 2 else TXT))]])
    tb(s, 686, y + 40, 510, 34, [[(d, dict(size=13.5, color=MUT72))]], line_spacing=1.4)
footer(s, "方法篇", "07 / %d" % TOTAL)

# ============ P08 提示词改写实战 ============
s = new_slide(prs)
header(s, "提示词改写：三组实战对照", "加约束比加形容词有效——每一组“好”，都附带了可判定的验收")
pairs = [
    ("给出验收标准", "实现一个校验邮箱的函数",
     "写一个 validateEmail：user@example.com → true、invalid → false、user@.com → false。实现后运行测试，直到全部通过。"),
    ("UI 改动要视觉验证", "把仪表盘弄好看点",
     "按这张设计稿实现（附截图）。完成后对结果截图并与原图对比，逐条列出差异并修复。"),
    ("修根因，不压症状", "构建挂了",
     "构建失败，报错如下（粘贴完整错误）。修复并验证构建通过——处理根因，不要压制报错。"),
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
bottom_judge(s, 64, 606, 1152, 52, "一句话自检：把“完成定义”盖住——如果你无法判断这部分是否算完成，它就还不够具体。")
footer(s, "方法篇", "08 / %d" % TOTAL)

# ============ P09 工作流四阶段 ============
s = new_slide(prs)
header(s, "单次任务的标准动作", "探索 → 规划 → 编码 → 提交——分开，是因为两段旅程的失败成本完全不同")
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
judge_card(s, 660, 370, 556, 252, "规划的成本是固定的，收益取决于“做错方向的概率”——用它对冲高不确定性，而不是给每个任务都上流程。探索错了，损失几轮对话；在错误方向写完 800 行再被发现，损失的是你的信任和一下午。", 17)
footer(s, "方法篇", "09 / %d" % TOTAL)

# ============ P10 让 Agent 采访你 ============
s = new_slide(prs)
header(s, "大特性起手式：先让它采访你", "不确定的地方不该让它猜，该让它问——把规格改精确，回报高于盯着实现")
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
    ("自包含", "不读对话记录也能看懂——别人（和未来的你）可以直接接手"),
    ("有边界", "明确写下“什么不做”——范围蔓延是 Agent 返工的最大来源"),
    ("可验收", "以端到端验证步骤收尾——完成与否，跑一条命令就知道"),
]
for i, (t, d) in enumerate(spec_pts):
    y = 204 + i * 52
    rect(s, 684, y + 4, 10, 10, fill=TEAL, radius=5)
    tb(s, 704, y - 2, 480, 44, [[(t + " · ", dict(size=14.5, bold=True, color=TXT)), (d, dict(size=14, color=MUT72))]], line_spacing=1.45)
judge_card(s, 660, 388, 556, 240, "口述需求、边做边改，会让每一次口头纠偏同时污染上下文与代码——规格未固化就开工，是最高频的失败模式之一。把规格写精确所花的时间，回报高于你盯着它实现所花的时间。", 17)
footer(s, "方法篇", "10 / %d" % TOTAL)

# ============ P11 扉页 工程篇 ============
section_slide(prs, "03", os.path.join(A, "ch_practice.png"), "g_ch03.png", "工程篇", "产出差距在上下文结构，不在模型",
              "同一个模型，在不同团队手里产出天差地别。差距来自四层：规则怎么沉淀、任务怎么写、能力怎么扩展、会话怎么治理。自底向上讲。")

# ============ P12 上下文工程四层 (Hero) ============
s = new_slide(prs)
header(s, "上下文工程：这门手艺的主体", "五层结构——把静态约束从提示词里搬出去，主线程只留本次任务契约")
layers = [
    ("⑤", "会话层", "治理", "清理、压缩、回溯、子智能体隔离——管理已进入窗口的内容", False),
    ("④", "能力层", "按需", "技能、子智能体、外部工具、事件钩子——不常驻，需要时才加载", False),
    ("③", "任务层", "临时", "本次的任务简报：目标、上下文、约束、验收——每次任务临时提供", True),
    ("②", "规则层", "常驻", "每次会话都加载的项目约定——必须极短，只放普适的东西", True),
    ("①", "模型层", "底座", "基座推理与代码能力——决定下限，你改不了（认知篇已讲，本章自 ② 讲起）", False),
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
tb(s, 800, 152, 416, 130, [[("同一个模型，在不同团队手里产出天差地别——差距几乎都出在 ②③④：规则写得准不准、简报给得清不清、能力接得全不全。", dict(size=17, color=TXT))]], line_spacing=1.7)
rect(s, 800, 306, 4, 88, fill=BLUE)
tb(s, 818, 306, 398, 96, [[("模型的差距是线性的，上下文结构的差距是乘性的。", dict(size=15.5, color=MUT72))]], line_spacing=1.6)
judge_card(s, 800, 440, 416, 190, "模型决定下限，上下文结构决定上限。换模型不如补上下文——这是这门课最值钱的一句话。", 17.5)
footer(s, "工程篇", "12 / %d" % TOTAL)

# ============ P13 规则层 ============
s = new_slide(prs)
header(s, "规则层：把默契写进仓库", "每次会话都加载——所以只放普适的东西，并且必须极短")
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
tb(s, 764, 164, 300, 26, [[("篇幅纪律（最重要）", dict(size=17, bold=True, color=BLUE))]])
tb(s, 764, 196, 428, 62, [[("对每一行问：“删掉这行会让它犯错吗？” 答不上来就删。臃肿的规则会让它忽略你真正的指令——目标 < 100 行。", dict(size=14, color=TXT))]], line_spacing=1.5)
rect(s, 740, 284, 476, 170, fill=CARD, line=CARD_BR, radius=10)
tb(s, 764, 300, 300, 26, [[("排错指南", dict(size=17, bold=True))]])
debugs = [("明令禁止仍反复出现", "→ 文件太长，规则被淹没：剪"),
          ("问了规则里已答的问题", "→ 表述有歧义：改写得更具体"),
          ("单条指令总被跳过", "→ 只在这一行加强调，别处处加")]
for i, (a, b) in enumerate(debugs):
    y = 336 + i * 38
    tb(s, 764, y, 220, 32, [[(a, dict(size=13.5, color=MUT78))]])
    tb(s, 976, y, 224, 32, [[(b, dict(size=13.5, color=TEAL))]])
judge_card(s, 740, 470, 476, 158, "把它当代码对待：出问题就审查、定期修剪、用“行为是否真的改变”验证改动。纳入 git 团队共建——这份文件的价值随时间复利增长。", 16)
footer(s, "工程篇", "13 / %d" % TOTAL)

# ============ P14 任务层 ============
s = new_slide(prs)
header(s, "任务层：把提示词写具体", "它能推断意图，但读不了你的心——四种策略，一份简报")
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
elems = [("Goal", "要什么——一句话说清终点"), ("Context", "在哪儿——具体路径，指向既有范式"),
         ("Constraints", "不能做什么——依赖、范围、风格"), ("Done-when", "什么叫完成——可运行、可判定")]
for i, (t, d) in enumerate(elems):
    col = i % 2
    row = i // 2
    x = 684 + col * 262
    y = 400 + row * 58
    tb(s, x, y, 250, 22, [[(t, dict(size=14.5, bold=True, color=TEAL, font=MONO))]])
    tb(s, x, y + 24, 250, 30, [[(d, dict(size=12.5, color=MUT72))]], line_spacing=1.35)
bottom_judge(s, 64, 546, 1152, 76, "明确程度与任务确定性成正比：探索期可以用模糊提示找灵感（“你觉得这个文件有哪些可改进？”），交付期必须给出可判定的验收。")
footer(s, "工程篇", "14 / %d" % TOTAL)

# ============ P15 能力层：选型 ============
s = new_slide(prs)
header(s, "能力层：五种扩展机制，别选错工具", "要共享内容用技能；要隔离消耗用子智能体——选错是常见的浪费")
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
judge_card(s, 830, 366, 386, 260, "内容由技能承载，消耗由子智能体隔离。分不清这一对，等于把全部调研成本压在自己的主上下文里——窗口就是这么被悄悄吃掉的。", 16.5)
footer(s, "工程篇", "15 / %d" % TOTAL)

# ============ P16 能力层：渐进式搭建 ============
s = new_slide(prs)
header(s, "渐进式搭建：让问题触发配置", "你不需要预先配置任何东西——每个机制都有一个可识别的触发信号")
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
bottom_judge(s, 64, 610, 1152, 52, "同样的信号也告诉你何时更新：一个反复出现的错误，应该是规则文件的一次编辑，而不是对话里的又一次口头纠正。")
footer(s, "工程篇", "16 / %d" % TOTAL)

# ============ P17 会话层 ============
s = new_slide(prs)
header(s, "会话层：上下文治理", "清理、压缩、回溯是你的日常操作——不是急救手段")
rect(s, 64, 150, 520, 214, fill=TEAL_BG, line=TEAL_BR, radius=10)
tb(s, 88, 168, 460, 62, [[("硬规则：同一问题上纠正了两次", dict(size=19, bold=True)), ], [("→ 立刻清理上下文，重开会话", dict(size=19, bold=True, color=TEAL))]], line_spacing=1.45)
tb(s, 88, 244, 470, 100, [[("此时上下文已塞满失败方案。带着学到的东西写一个更好的初始提示词——干净的会话几乎总是优于一个堆满纠正的长会话。", dict(size=14.5, color=TXT))]], line_spacing=1.55)
rect(s, 64, 380, 520, 168, fill=CARD, line=CARD_BR, radius=10)
tb(s, 88, 396, 400, 26, [[("检查点：放心让它试错", dict(size=17, bold=True))]])
tb(s, 88, 430, 470, 100, [[("冒险方案不行就回溯换条路，检查点随会话保存，隔天依然可回溯。⚠ 但它只追踪编辑工具产生的改动——检查点不能替代 git。", dict(size=14, color=MUT78))]], line_spacing=1.5)
govs = [
    ("主动压缩", "任务之间清理上下文；定向压缩（“聚焦 API 改动”）；把“压缩时必须保留什么”写进规则文件。"),
    ("调研外包", "需要读大量文件的调研交给子智能体——在独立上下文里消耗，只回传摘要。"),
    ("即时纠偏", "发现跑偏立刻中断改向，上下文保留，可以马上换指令。越晚纠偏，成本越高。"),
    ("会话当分支", "每条工作流一个命名会话（如 oauth-migration），跨天继续不用重复解释背景。"),
]
for i, (t, d) in enumerate(govs):
    y = 150 + i * 104
    rect(s, 620, y, 596, 94, fill=CARD, line=CARD_BR, radius=10)
    rect(s, 620, y, 4, 94, fill=BLUE)
    tb(s, 644, y + 12, 300, 26, [[(t, dict(size=16.5, bold=True))]])
    tb(s, 644, y + 42, 548, 46, [[(d, dict(size=13.5, color=MUT72))]], line_spacing=1.45)
bottom_judge(s, 64, 574, 1152, 76, "最好的结果来自紧密的反馈循环，而不是完美的第一版提示词。发现跑偏立刻纠偏——这是用 Agent 和用搜索引擎的根本区别。", 15.5)
footer(s, "工程篇", "17 / %d" % TOTAL)

# ============ P18 扉页 规模化篇 ============
section_slide(prs, "04", os.path.join(A, "ch_rollout.png"), "g_ch04.png", "规模化篇", "先画笼子，再放大产出",
              "先把一个智能体用得有效，再用并行、无头与流水线把产出成倍放大——顺序反了，只会把低效复制 N 份。")

# ============ P19 权限与安全 ============
s = new_slide(prs)
header(s, "权限：先画笼子，再放它出去", "每次都弹窗确认，点过第十次你就不再是审查——只是在点")
rect(s, 64, 150, 560, 116, fill=CARD, line=CARD_BR, radius=10)
tb(s, 88, 164, 300, 24, [[("两难", dict(size=16, bold=True, color=BLUE))]])
tb(s, 88, 194, 510, 62, [[("全确认 = 安全但注意力疲劳，审查形同虚设；全放行 = 风险失控。真正有效的注意力是稀缺的，别消耗在低价值点击上。", dict(size=13.5, color=MUT72))]], line_spacing=1.45)
rect(s, 64, 282, 560, 108, fill=CARD, line=CARD_BR, radius=10)
tb(s, 88, 296, 300, 24, [[("两条互补的解法", dict(size=16, bold=True, color=TEAL))]])
tb(s, 88, 326, 510, 56, [[("权限白名单：具名放行你信任的操作 ｜ 沙箱隔离：限制文件系统与网络边界，让它在笼子内自由工作。", dict(size=13.5, color=MUT72))]], line_spacing=1.45)
rect(s, 64, 406, 560, 222, fill=CARD, line=CARD_BR, radius=10)
tb(s, 88, 422, 300, 24, [[("四种权限模式", dict(size=16, bold=True))]])
modes = [("自动复核", "独立模型替你拦风险操作，日常改动不弹窗"),
         ("手动确认", "写文件、跑命令前逐次询问——你想亲自批准时"),
         ("只读规划", "只能探索与出方案，不落任何修改——探索阶段"),
         ("预批准", "只放行具名工具，其余一律拒绝——无人值守批量")]
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
bottom_judge(s, 660, 584, 556, 60, "便利来自权限，风险也来自权限——顺序不能颠倒。")
footer(s, "规模化篇", "19 / %d" % TOTAL)

# ============ P20 规模化路径 ============
s = new_slide(prs)
header(s, "规模化：从一个 Agent 到一条流水线", "先让一个智能体用得有效，再成倍放大——顺序反了，只会复制低效")
paths = [
    ("无头模式", "命令行非交互调用 + 结构化输出（JSON），接进 CI、提交前钩子与自动化脚本。"),
    ("并行会话", "在隔离检出中同时跑多个会话，改动互不覆盖；会话之间可传递发现。"),
    ("扇出迁移", "大迁移拆成批量任务：先 2-3 个小样试跑、调提示词，再全量——具名放行所需工具。"),
    ("无人值守", "自动复核模式下跑长任务；结束时由全新上下文的评审智能体做对抗式检查。"),
    ("流水线集成", "输出 JSON 供下游命令消费——让 Agent 成为数据处理管道中的一环。"),
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
judge_card(s, 700, 428, 516, 200, "全新上下文的评审者不会偏向自己刚写的代码——多会话最大的红利不是速度，是质量。⚠ 被命令“找缺口”的评审者几乎必然报出缺口：明确只标记影响正确性的问题，把其余视为可选，警惕过度工程。", 15.5)
footer(s, "规模化篇", "20 / %d" % TOTAL)

# ============ P21 七个失败模式 ============
s = new_slide(prs)
header(s, "七个高频失败模式与修正", "它们反复出现在所有团队——早识别，省的是一下午")
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
judge_card(s, 864, 470, 352, 156, "这些不是“工具不行”，而是工作方法问题——方法对了，换一个 Agent 也是同样的效果。", 16)
footer(s, "规模化篇", "21 / %d" % TOTAL)

# ============ P22 落地路线 ============
s = new_slide(prs)
header(s, "团队落地：90 天，三个阶段", "先统一规则文件，再谈覆盖率——顺序不能颠倒")
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
    ("采纳率", "实际使用比例与频次——不好用的东西推广不动，这是前置指标"),
    ("返工率", "产出被人工推翻/重写的比例——直接反映上下文质量"),
    ("缺陷密度", "单位改动引入的缺陷数——防止“速度换质量”的自欺"),
    ("端到端时长", "需求提出到合入的墙钟时间——唯一能证明“整体更快”的指标"),
]
for i, (t, d) in enumerate(metrics):
    y = 150 + i * 96
    rect(s, 740, y, 476, 86, fill=CARD, line=CARD_BR, radius=10)
    tb(s, 764, y + 12, 120, 28, [[(t, dict(size=16.5, bold=True, color=TEAL))]])
    tb(s, 764, y + 44, 428, 36, [[(d, dict(size=13, color=MUT72))]], line_spacing=1.4)
judge_card(s, 740, 538, 476, 122, "只统计“生成了多少代码”会奖励错误行为——产出更多代码，从来不是工程目标。", 16)
footer(s, "规模化篇", "22 / %d" % TOTAL)

# ============ P23 结束页 (Hero) ============
s = new_slide(prs)
tb(s, 64, 96, 300, 120, [[("“", dict(size=120, bold=True, color=TEAL, font=MONO))]])
tb(s, 64, 190, 1080, 160, [[("这套方法的终点，不是让 AI 写代码——", dict(size=34, bold=True))],
                           [("而是把团队的工程默契，从人的脑子里搬进仓库。", dict(size=34, bold=True, color=TEAL))]], line_spacing=1.55)
hline(s, 64, 372, 72, TEAL, 3)
tb(s, 64, 396, 600, 28, [[("本周就能开始的三件事", dict(size=20, bold=True))]])
acts = [
    ("今天", "为主仓库写一份规则文件，删到 100 行以内，纳入 git。"),
    ("本周", "用任务简报四要素跑一个真实需求——必须带一个能跑出 pass/fail 的验证命令。"),
    ("本周", "把一个重复了三次以上的操作，固化成可复用的技能。"),
]
for i, (tag, d) in enumerate(acts):
    x = 64 + i * 384
    rect(s, x, 436, 368, 132, fill=CARD, line=CARD_BR, radius=10)
    tb(s, x + 22, 452, 120, 24, [[(tag, dict(size=14, bold=True, color=TEAL, font=MONO))]])
    tb(s, x + 22, 484, 324, 72, [[(d, dict(size=14, color=MUT78))]], line_spacing=1.5)
tb(s, 64, 604, 1152, 44, [[("何时打破规则：深陷复杂问题时让上下文累积；探索性任务跳过规划；想先看它怎么解释问题时用模糊提示——直觉比任何指南可靠，而直觉从实践中长出来。", dict(size=13.5, color=MUT55))]], line_spacing=1.5)
tb(s, 980, 676, 300, 22, [[("ENGINEERING PRACTICES · 2026.10", dict(size=12, color=MUT50, font=MONO, spacing=2))]], align=PP_ALIGN.RIGHT)

prs.save(OUT)
print("saved:", OUT, os.path.getsize(OUT), "bytes, slides:", len(prs.slides.__iter__.__self__._sldIdLst))
