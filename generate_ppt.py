from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Inches, Pt


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

NAVY = RGBColor(10, 25, 47)
BLUE = RGBColor(30, 58, 138)
GREEN = RGBColor(16, 185, 129)
DEEP_GREEN = RGBColor(5, 150, 105)
ORANGE = RGBColor(245, 158, 11)
RED = RGBColor(220, 76, 70)
INK = RGBColor(24, 39, 60)
MUTED = RGBColor(92, 108, 127)
PALE = RGBColor(243, 244, 246)
LINE = RGBColor(218, 225, 234)
WHITE = RGBColor(255, 255, 255)
LIGHT_BLUE = RGBColor(232, 239, 253)
LIGHT_GREEN = RGBColor(226, 247, 239)
LIGHT_ORANGE = RGBColor(255, 245, 219)
LIGHT_RED = RGBColor(254, 236, 235)

FONT = "Microsoft YaHei"


def shape(slide, kind, x, y, w, h, fill, line=None, radius=False):
    s = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.color.rgb = line or fill
    if radius and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = 0.08
    return s


def text(slide, value, x, y, w, h, size=16, color=INK, bold=False, align=PP_ALIGN.LEFT,
         font=FONT, valign=MSO_ANCHOR.TOP, margin=0.04, italic=False):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin)
    tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_after = Pt(0)
    run = p.add_run()
    run.text = value
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return box


def rich_text(slide, runs, x, y, w, h, size=16, color=INK, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.03)
    p = tf.paragraphs[0]
    p.alignment = align
    for value, run_color, bold in runs:
        r = p.add_run()
        r.text = value
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = run_color or color
    return box


def base(slide, title, section="十五五 · 无锡能源发展规划研究思路", dark=False):
    bg = NAVY if dark else WHITE
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = bg
    if not dark:
        shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 0.12, BLUE)
        text(slide, section, 0.55, 0.22, 6.8, 0.25, 9, MUTED)
        text(slide, title, 0.55, 0.62, 11.9, 0.55, 24, NAVY, True)
        shape(slide, MSO_SHAPE.RECTANGLE, 0.55, 1.28, 0.54, 0.06, GREEN)
    else:
        text(slide, section, 0.6, 0.32, 7, 0.25, 10, RGBColor(177, 194, 215))
    slide_no = len(prs.slides)
    text(slide, f"{slide_no:02d}", 12.2, 7.03, 0.55, 0.22, 9, MUTED if not dark else RGBColor(130, 151, 179), align=PP_ALIGN.RIGHT)


def card(slide, x, y, w, h, title, body="", accent=BLUE, fill=PALE, value=None, value_color=INK):
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill, fill, True)
    shape(slide, MSO_SHAPE.RECTANGLE, x, y, 0.07, h, accent, accent)
    text(slide, title, x + 0.2, y + 0.15, w - 0.35, 0.28, 13, INK, True)
    if value is not None:
        text(slide, value, x + 0.2, y + 0.48, w - 0.35, 0.48, 24, value_color, True)
        if body:
            text(slide, body, x + 0.2, y + 1.03, w - 0.35, h - 1.12, 10, MUTED)
    else:
        text(slide, body, x + 0.2, y + 0.52, w - 0.35, h - 0.64, 11, MUTED)


def bullet_card(slide, x, y, w, h, title, body, accent=BLUE, fill=WHITE, number=None):
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill, LINE, True)
    if number is not None:
        shape(slide, MSO_SHAPE.OVAL, x + 0.18, y + 0.18, 0.34, 0.34, accent, accent)
        text(slide, str(number), x + 0.18, y + 0.19, 0.34, 0.28, 11, WHITE, True, PP_ALIGN.CENTER)
        tx = x + 0.63
    else:
        shape(slide, MSO_SHAPE.RECTANGLE, x + 0.18, y + 0.22, 0.08, 0.3, accent, accent)
        tx = x + 0.4
    text(slide, title, tx, y + 0.16, w - (tx - x) - 0.18, 0.28, 13, INK, True)
    text(slide, body, x + 0.2, y + 0.58, w - 0.4, h - 0.72, 11, MUTED)


def pill(slide, value, x, y, w, fill=LIGHT_GREEN, color=DEEP_GREEN):
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, 0.28, fill, fill, True)
    text(slide, value, x, y + 0.01, w, 0.22, 9, color, True, PP_ALIGN.CENTER)


def connector(slide, x1, y1, x2, y2, color=GREEN, width=2):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(width)
    return c


def add_city_network(slide):
    for x, h in [(0.0, 1.0), (0.8, 1.55), (1.6, 0.85), (2.35, 1.8), (3.15, 1.1), (4.0, 2.2), (5.0, 1.25), (6.0, 2.7), (7.1, 1.2), (8.2, 1.9), (9.4, 1.0), (10.4, 2.1), (11.5, 1.4), (12.5, 2.4)]:
        shape(slide, MSO_SHAPE.RECTANGLE, x, 7.5 - h, 0.58, h, RGBColor(16, 42, 73), RGBColor(16, 42, 73))
    for x, y in [(1.1, 2.0), (3.3, 1.1), (5.2, 2.4), (7.6, 1.4), (9.8, 2.3), (11.3, 1.0)]:
        shape(slide, MSO_SHAPE.OVAL, x, y, 0.14, 0.14, GREEN, GREEN)
    for a, b in [((1.17, 2.07), (3.37, 1.17)), ((3.37, 1.17), (5.27, 2.47)), ((5.27, 2.47), (7.67, 1.47)), ((7.67, 1.47), (9.87, 2.37)), ((9.87, 2.37), (11.37, 1.07))]:
        connector(slide, *a, *b, RGBColor(68, 190, 159), 1.4)


def slide_cover():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "", dark=True)
    add_city_network(slide)
    shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, NAVY, NAVY)
    shape(slide, MSO_SHAPE.RECTANGLE, 0, 5.9, 13.333, 1.6, RGBColor(7, 33, 59), RGBColor(7, 33, 59))
    add_city_network(slide)
    text(slide, "“十五五”时期无锡能源发展规划研究思路", 0.75, 1.55, 8.9, 0.75, 31, WHITE, True)
    text(slide, "构建清洁低碳、安全韧性、数智协同的现代新型城市能源体系", 0.78, 2.5, 9.3, 0.42, 16, RGBColor(195, 215, 235))
    shape(slide, MSO_SHAPE.RECTANGLE, 0.8, 3.26, 1.05, 0.07, GREEN, GREEN)
    text(slide, "无锡市发展和改革委员会  /  能源规划专项研究组", 0.8, 3.58, 6.5, 0.3, 11, RGBColor(170, 191, 215))
    pill(slide, "规划研究 · 2025—2030", 0.8, 4.12, 1.85, RGBColor(24, 72, 93), RGBColor(132, 233, 202))
    text(slide, "太湖生态红线区  ·  万亿级制造业高地  ·  双碳达峰冲刺期", 0.8, 6.65, 8.5, 0.3, 11, RGBColor(155, 178, 203))


def slide_summary():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "引言与核心摘要")
    card(slide, 0.6, 1.7, 5.7, 1.2, "战略能位", "万亿级制造业高地  ·  太湖生态红线区  ·  “双碳”达峰冲刺期", BLUE, LIGHT_BLUE)
    text(slide, "三大严苛边界", 0.7, 3.25, 2.2, 0.3, 15, NAVY, True)
    bullet_card(slide, 0.6, 3.72, 1.78, 2.05, "资源匮乏", "一次能源自给率\n不足 5%", BLUE, PALE, 1)
    bullet_card(slide, 2.52, 3.72, 1.78, 2.05, "刚性指标", "GDP 年均增速\n约 5%", ORANGE, PALE, 2)
    bullet_card(slide, 4.44, 3.72, 1.78, 2.05, "用能升级", "微秒级高可靠\n100% 全溯源绿电", GREEN, PALE, 3)
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 6.75, 1.7, 5.95, 4.1, NAVY, NAVY, True)
    text(slide, "破局方针", 7.15, 2.05, 2, 0.3, 15, RGBColor(157, 182, 210), True)
    text(slide, "内源极限挖潜", 7.15, 2.72, 2.1, 0.45, 22, WHITE, True)
    text(slide, "+", 9.3, 2.72, 0.35, 0.45, 22, GREEN, True, PP_ALIGN.CENTER)
    text(slide, "跨区战略引绿", 9.75, 2.72, 2.2, 0.45, 22, WHITE, True)
    text(slide, "+", 7.15, 3.55, 0.35, 0.45, 22, GREEN, True, PP_ALIGN.CENTER)
    text(slide, "数智源网荷储", 7.6, 3.55, 2.7, 0.45, 22, WHITE, True)
    text(slide, "把外部资源、城市空间与数字能力\n转化为可验证的能源安全与产业竞争力", 7.15, 4.55, 4.7, 0.65, 13, RGBColor(190, 205, 224))
    pill(slide, "核心判断：能源系统必须从“供给保障”升级为“产业底座”", 7.15, 5.35, 4.75, RGBColor(25, 66, 88), RGBColor(131, 231, 195))


def slide_supply():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "一、发展现状——“十四五”供应底盘与运行成效")
    items = [("煤炭压舱石", "2473", "万吨 / 年消费量稳定", BLUE, "600 万千瓦公用煤电全超低排放"), ("装机多元化", "1932.5", "万千瓦 / 总装机", GREEN, "非化石占比 42.4% · 抽蓄 100 万千瓦"), ("油气储运网", "513.8", "万千瓦 / 气电顶峰", ORANGE, "西气 + 川气双源 · LNG 二期扩容"), ("电网高可靠", "6990.5", "万千伏安 / 变电容量", BLUE, "最高负荷突破 1450 万千瓦")]
    for i, (title, value, label, accent, body) in enumerate(items):
        card(slide, 0.6 + i * 3.08, 1.75, 2.78, 2.35, title, label + "\n" + body, accent, PALE, value, accent)
    text(slide, "供应底盘判断", 0.65, 4.7, 2, 0.3, 15, NAVY, True)
    card(slide, 0.6, 5.12, 3.8, 1.0, "底线稳", "煤电、油气与跨区电力形成多元支撑", BLUE, LIGHT_BLUE)
    card(slide, 4.75, 5.12, 3.8, 1.0, "结构优", "非化石装机突破四成，光伏成为增量主力", GREEN, LIGHT_GREEN)
    card(slide, 8.9, 5.12, 3.8, 1.0, "弹性需补", "气电、抽蓄与储能共同承担尖峰调节", ORANGE, LIGHT_ORANGE)


def slide_efficiency():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "一、发展现状——能耗双控成效与转型亮点")
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 1.65, 5.95, 4.85, NAVY, NAVY, True)
    text(slide, "节能量化指标", 0.95, 2.0, 2.3, 0.3, 15, RGBColor(166, 190, 218), True)
    text(slide, "22.9%", 0.95, 2.6, 3.0, 0.75, 42, WHITE, True)
    text(slide, "单位 GDP 能耗累计下降", 1.0, 3.38, 3.8, 0.3, 13, RGBColor(190, 207, 227))
    shape(slide, MSO_SHAPE.RECTANGLE, 0.98, 4.1, 4.8, 0.18, RGBColor(54, 77, 106), RGBColor(54, 77, 106))
    shape(slide, MSO_SHAPE.RECTANGLE, 0.98, 4.1, 3.75, 0.18, GREEN, GREEN)
    text(slide, "省目标 14%", 4.95, 3.83, 1.0, 0.3, 10, RGBColor(155, 178, 203), align=PP_ALIGN.RIGHT)
    text(slide, "工业增加值能耗累计下降 34.37%", 0.98, 4.75, 4.8, 0.35, 14, WHITE, True)
    text(slide, "全社会用电量 917.98 亿千瓦时 · 工业占 70%—75%", 0.98, 5.55, 4.9, 0.42, 11, RGBColor(190, 207, 227))
    text(slide, "结构优化与智造集群", 6.95, 1.7, 3.5, 0.35, 15, NAVY, True)
    bullet_card(slide, 6.8, 2.2, 5.9, 1.15, "光伏爆发", "装机 661 万千瓦，占新能源 82% 以上；工商业分布式占比 84.9%。", GREEN, LIGHT_GREEN)
    bullet_card(slide, 6.8, 3.62, 5.9, 1.15, "绿电先锋", "率先成立“一站一中心”；2024 年绿电交易 17.4 亿 kWh，占全省 15%。", BLUE, LIGHT_BLUE)
    bullet_card(slide, 6.8, 5.04, 5.9, 1.15, "装备高地", "规上产值超 2300 亿元，光储氢风母机龙头集聚。", ORANGE, LIGHT_ORANGE)


def slide_bottleneck():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "一、发展现状——瓶颈与短板剖析（问题导向）")
    text(slide, "当前矛盾不在“有没有项目”，而在“安全、空间、绿色与调节”能否同时成立", 0.65, 1.65, 11.8, 0.4, 17, NAVY, True)
    data = [("01", "对外依存极高", "一次能源自给率不足 5%；外电依赖度超 20%；天然气环网尚未完全闭环。", RED, LIGHT_RED), ("02", "空间逼近极限", "生态红线严苛；优质屋顶资源开发殆尽；集中式风光无地可拓。", ORANGE, LIGHT_ORANGE), ("03", "绿色剪刀差扩大", "非化石消费占比约 11.7%，距 26% 目标差距大；CBAM 倒逼绿电出海。", BLUE, LIGHT_BLUE), ("04", "调节资源与配网滞后", "新型储能仅 16.4 万千瓦；分布式光伏带来台区反向过载。", DEEP_GREEN, LIGHT_GREEN)]
    for i, (num, title, body, accent, fill) in enumerate(data):
        x = 0.65 + (i % 2) * 6.1
        y = 2.35 + (i // 2) * 1.85
        bullet_card(slide, x, y, 5.7, 1.42, title, body, accent, fill)
        text(slide, num, x + 4.95, y + 0.18, 0.55, 0.3, 14, accent, True, PP_ALIGN.RIGHT)
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.65, 6.25, 12.0, 0.52, NAVY, NAVY, True)
    text(slide, "结论  /  无锡必须用“跨区资源 + 城市级调节 + 产业协同”换取本地绿色增量", 0.95, 6.36, 11.4, 0.25, 13, WHITE, True, PP_ALIGN.CENTER)


def slide_swot():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "二、形势特征——机遇与挑战研判")
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 1.65, 5.9, 4.95, LIGHT_GREEN, LIGHT_GREEN, True)
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 6.8, 1.65, 5.9, 4.95, LIGHT_RED, LIGHT_RED, True)
    text(slide, "战略机遇", 0.95, 1.98, 2.0, 0.35, 19, DEEP_GREEN, True)
    text(slide, "REAL OPTIONS", 4.75, 2.05, 1.2, 0.2, 8, DEEP_GREEN, True, PP_ALIGN.RIGHT)
    text(slide, "现实挑战", 7.15, 1.98, 2.0, 0.35, 19, RED, True)
    text(slide, "HARD CONSTRAINTS", 11.4, 2.05, 1.0, 0.2, 8, RED, True, PP_ALIGN.RIGHT)
    opp = [("制度变革", "双控转双碳，原料与绿电不计入总量"), ("大通道落地", "凤城—梅里、锦苏直流、白鹤滩入苏"), ("数智赋能", "物联网底座驱动 VPP 与现货辅助服务")]
    chal = [("保供与降碳平衡难", "尖峰负荷将破 1800 万千瓦"), ("传统高载能改造重", "特钢、石化流程再造周期长"), ("装备强与应用弱脱节", "本地缺乏集中式承载大场景"), ("高品质电能诉求", "半导体产线要求微秒级零闪动")]
    for i, (t, b) in enumerate(opp):
        bullet_card(slide, 0.92, 2.62 + i * 1.13, 5.25, 0.84, t, b, DEEP_GREEN, WHITE, i + 1)
    for i, (t, b) in enumerate(chal):
        bullet_card(slide, 7.12, 2.62 + i * 0.91, 5.25, 0.68, t, b, RED, WHITE, i + 1)


def slide_strategy():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "三、总体要求——指导思想与战略方针")
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 1.65, 12.1, 1.25, NAVY, NAVY, True)
    text(slide, "控煤稳盘  ·  强网扩电  ·  战略引绿  ·  提质增效", 0.85, 2.0, 11.6, 0.46, 24, WHITE, True, PP_ALIGN.CENTER)
    text(slide, "16 字方针", 0.85, 1.78, 1.0, 0.2, 9, RGBColor(144, 171, 201), True)
    text(slide, "指导思想  /  以安全韧性为底线、以绿色低碳为方向、以数智协同为手段，将能源体系打造为先进制造业的竞争性底座。", 0.7, 3.25, 11.8, 0.4, 14, INK, True)
    principles = [("先立后破，安全兜底", "煤电向支撑调节电源转变"), ("内生外引，源网协同", "市域极限挖潜，跨省区引绿"), ("产用联动，创新驱动", "场景换装备，装备促场景"), ("深化改革，市场主导", "现货、绿证、碳预算协同")]
    for i, (t, b) in enumerate(principles):
        card(slide, 0.65 + i * 3.06, 4.1, 2.75, 1.75, t, b, [BLUE, GREEN, ORANGE, DEEP_GREEN][i], [LIGHT_BLUE, LIGHT_GREEN, LIGHT_ORANGE, LIGHT_GREEN][i])
    pill(slide, "从资源约束出发，以系统性解决方案换取发展空间", 4.05, 6.25, 5.25, LIGHT_BLUE, BLUE)


def slide_targets():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "三、总体要求——2030 年主要发展指标看板")
    groups = [("安全保供", BLUE, [("煤炭消费总量", "2473 万吨", "≤ 2400 万吨"), ("最高用电负荷", "1451.7 万千瓦", "1800 万千瓦左右")]), ("清洁低碳", GREEN, [("非化石消费比重", "约 11.7%", "26% 左右"), ("可再生能源装机", "802.8 万千瓦", "1200 万千瓦左右"), ("区外清洁绿电占比", "约 40%", "突破 60%")]), ("系统调节", ORANGE, [("新型储能规模", "16.4 万 kW", "100 万 kW / 200 万 kWh"), ("VPP / 需求响应", "基础试点", "≥ 最高负荷 5%"), ("敏感用电可靠率", "99.99%", "99.999%")]), ("产业节能", DEEP_GREEN, [("单位 GDP 能耗降幅", "累计降 22.9%", "完成省下达目标"), ("新能源装备产业", "约 3500 亿元", "突破 5500 亿元")])]
    for col, (name, accent, rows) in enumerate(groups):
        x = 0.58 + col * 3.1
        shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, 1.62, 2.82, 4.95, WHITE, LINE, True)
        shape(slide, MSO_SHAPE.RECTANGLE, x, 1.62, 2.82, 0.6, accent, accent)
        text(slide, name, x + 0.16, 1.77, 2.5, 0.25, 14, WHITE, True)
        y = 2.45
        for label, actual, target in rows:
            text(slide, label, x + 0.16, y, 2.48, 0.28, 10, MUTED, True)
            text(slide, actual, x + 0.16, y + 0.35, 2.48, 0.24, 10, INK)
            text(slide, "→  " + target, x + 0.16, y + 0.69, 2.48, 0.32, 12, accent, True)
            shape(slide, MSO_SHAPE.RECTANGLE, x + 0.16, y + 1.16, 2.48, 0.01, LINE, LINE)
            y += 1.38
    text(slide, "约束性指标优先保障底线，预期性指标强化导向，形成“目标—工程—考核”闭环", 1.2, 6.78, 10.8, 0.28, 11, MUTED, False, PP_ALIGN.CENTER)


def slide_tasks_1():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "四、重点任务（一）——构建坚实可靠的能源保障体系")
    text(slide, "三大防线", 0.65, 1.65, 2, 0.3, 16, NAVY, True)
    tasks = [("煤电机组“三改联动”", "600 万千瓦现役机组灵活性改造；最小稳燃负荷降至 20%—30%；探索掺氨与 CCUS。", BLUE), ("天然气顶峰与环网联通", "扩建江阴、宜兴调峰燃机，新增气电 160 万千瓦；高压管网全闭环。", GREEN), ("油品供应与江阴储运极", "严控新增炼化；推进综合加能站；打造长江下游低碳液体燃料枢纽。", ORANGE)]
    for i, (title, body, accent) in enumerate(tasks):
        bullet_card(slide, 0.65, 2.15 + i * 1.25, 7.0, 0.98, title, body, accent, PALE, i + 1)
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 8.05, 1.65, 4.65, 4.85, NAVY, NAVY, True)
    text(slide, "专栏 1  ·  重点工程", 8.4, 2.0, 3.5, 0.3, 15, RGBColor(154, 181, 211), True)
    projects = ["利港电厂扩建与改造", "嘉盛 LNG 调峰储配站", "16 万 m³ 储罐", "天然气高压“一张网”闭环"]
    for i, p in enumerate(projects):
        shape(slide, MSO_SHAPE.OVAL, 8.45, 2.72 + i * 0.7, 0.19, 0.19, GREEN, GREEN)
        text(slide, p, 8.85, 2.63 + i * 0.7, 3.35, 0.35, 13, WHITE, i == 2)
    pill(slide, "目标：保供有底、调峰有力、燃料有序", 8.4, 5.65, 3.55, RGBColor(26, 70, 89), RGBColor(133, 233, 201))


def slide_tasks_2():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "四、重点任务（二）——构建绿色多元的清洁能源供给体系")
    text(slide, "双轮驱动：市域立体挖潜 + 跨区引绿攻坚", 0.65, 1.65, 6.6, 0.35, 17, NAVY, True)
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.65, 2.25, 5.65, 3.95, LIGHT_GREEN, LIGHT_GREEN, True)
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 7.0, 2.25, 5.65, 3.95, LIGHT_BLUE, LIGHT_BLUE, True)
    text(slide, "内源极限挖潜", 1.0, 2.6, 2.5, 0.35, 20, DEEP_GREEN, True)
    text(slide, "把有限空间变成高密度绿色增量", 1.0, 3.02, 4.6, 0.3, 11, MUTED)
    for i, (t, b) in enumerate([("光伏 + 工业园区", "新建超亿元厂房 100% 预留载荷 / BIPV"), ("生物质与浅层地热", "Bio-LNG；太湖新城地源热泵")]):
        bullet_card(slide, 0.98, 3.55 + i * 1.15, 4.95, 0.85, t, b, DEEP_GREEN, WHITE, i + 1)
    text(slide, "2030 光伏装机破 1000 万千瓦", 1.0, 5.72, 4.7, 0.25, 12, DEEP_GREEN, True)
    text(slide, "跨区战略引绿", 7.35, 2.6, 2.5, 0.35, 20, BLUE, True)
    text(slide, "以物理通道与长期交易锁定清洁电力", 7.35, 3.02, 4.7, 0.3, 11, MUTED)
    for i, (t, b) in enumerate([("省内“引绿”", "凤城—梅里；联合投资苏北海风、核电"), ("特高压大动脉", "锦苏直流、白鹤滩入苏；区外调绿超 180 亿 kWh"), ("外氢入锡与副产利用", "副产氢提纯；预留西氢东输接口")]):
        bullet_card(slide, 7.33, 3.45 + i * 0.82, 4.95, 0.62, t, b, BLUE, WHITE, i + 1)


def slide_tasks_3():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "四、重点任务（三）——构建安全柔性的能源网络体系")
    cols = [("坚强主网", "500 千伏环网互济\n220 千伏分区运行\n跨江送电配套", BLUE), ("配网承载力倍增", "智能台区消除反向过载\nDVR + 超导储能\n半导体级供电示范区", GREEN), ("超充之城与加氢网络", "液冷超充桩 > 5000 根\n综合加能站 50 座\n2030 加氢站 > 15 座", ORANGE)]
    for i, (title, body, accent) in enumerate(cols):
        x = 0.65 + i * 4.15
        shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, 1.8, 3.7, 4.8, WHITE, LINE, True)
        shape(slide, MSO_SHAPE.RECTANGLE, x, 1.8, 3.7, 0.18, accent, accent)
        shape(slide, MSO_SHAPE.OVAL, x + 0.32, 2.23, 0.62, 0.62, accent, accent)
        text(slide, f"0{i+1}", x + 0.32, 2.36, 0.62, 0.22, 14, WHITE, True, PP_ALIGN.CENTER)
        text(slide, title, x + 0.32, 3.08, 3.0, 0.38, 18, NAVY, True)
        text(slide, body, x + 0.32, 3.78, 3.0, 1.25, 13, MUTED)
        pill(slide, ["网架互济", "供电不扰", "交通脱碳"][i], x + 0.32, 5.65, 1.15, [LIGHT_BLUE, LIGHT_GREEN, LIGHT_ORANGE][i], accent)
    connector(slide, 4.45, 4.05, 4.8, 4.05, LINE, 2)
    connector(slide, 8.6, 4.05, 8.95, 4.05, LINE, 2)


def slide_tasks_4():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "四、重点任务（四）——构建协同高效的能源运行体系")
    text(slide, "从“静态设备”转向“可感知、可交易、可调度”的城市能源操作系统", 0.65, 1.62, 11.5, 0.35, 16, NAVY, True)
    nodes = [("储能百兆瓦跨越", "100 万千瓦\n江阴 / 宜兴电网侧", 1.1, 3.0, GREEN), ("城市级 VPP", "370 家企业 + 5 万根桩\n削峰能力 ≥ 5%", 4.9, 2.25, BLUE), ("工业微电网", "10 个标杆\n48 小时黑启动", 8.8, 3.0, ORANGE)]
    for title, body, x, y, accent in nodes:
        shape(slide, MSO_SHAPE.OVAL, x, y, 2.7, 2.7, NAVY, accent)
        shape(slide, MSO_SHAPE.OVAL, x + 0.18, y + 0.18, 2.34, 2.34, RGBColor(18, 39, 68), RGBColor(18, 39, 68))
        text(slide, title, x + 0.28, y + 0.72, 2.15, 0.4, 15, WHITE, True, PP_ALIGN.CENTER)
        text(slide, body, x + 0.28, y + 1.28, 2.15, 0.55, 11, RGBColor(178, 203, 226), False, PP_ALIGN.CENTER)
    connector(slide, 3.8, 4.35, 4.9, 3.65, GREEN, 2)
    connector(slide, 7.6, 3.65, 8.8, 4.35, GREEN, 2)
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 3.25, 5.95, 6.85, 0.55, LIGHT_GREEN, LIGHT_GREEN, True)
    text(slide, "车网互动 V2G：谷充峰放，把移动储能纳入城市调度", 3.45, 6.09, 6.45, 0.22, 12, DEEP_GREEN, True, PP_ALIGN.CENTER)


def slide_tasks_5():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "四、重点任务（五）——构建节约集约的能源消费体系")
    cards = [("工业", "特钢转炉长流程 → 短流程废钢电炉 + 富氢冶炼；园区集中供气 100%", BLUE), ("算力", "新建算力 PUE ≤ 1.15；液冷全面普及；绿电就地消纳 ≥ 80%", GREEN), ("交通", "长三角绿氢物流走廊；氢能重卡超 1500 辆；试点纯电及氢能货船", ORANGE), ("建筑", "新建建筑 100% 星级绿建；公共机构全覆盖 EMC 合同能源管理", DEEP_GREEN)]
    for i, (t, b, accent) in enumerate(cards):
        x = 0.65 + (i % 2) * 6.1
        y = 1.75 + (i // 2) * 2.35
        shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, 5.7, 1.85, WHITE, LINE, True)
        shape(slide, MSO_SHAPE.RECTANGLE, x, y, 0.13, 1.85, accent, accent)
        text(slide, t, x + 0.35, y + 0.28, 1.0, 0.35, 20, accent, True)
        text(slide, b, x + 1.55, y + 0.3, 3.75, 0.95, 13, INK)
        pill(slide, ["流程再造", "能效硬约束", "交通新燃料", "梯级用能"][i], x + 0.35, y + 1.38, 1.2, [LIGHT_BLUE, LIGHT_GREEN, LIGHT_ORANGE, LIGHT_GREEN][i], accent)
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 2.35, 6.45, 8.6, 0.42, NAVY, NAVY, True)
    text(slide, "以单位产出能耗为标尺，推动能源消费从“总量管理”转向“效率管理”", 2.6, 6.54, 8.1, 0.22, 11, WHITE, True, PP_ALIGN.CENTER)


def slide_tasks_6():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "四、重点任务（六）——构建创新互促的能源产业体系")
    text(slide, "产业闭环：技术研发 → 本地装备 → 开放场景 → 绿色出海", 0.65, 1.65, 11.6, 0.35, 17, NAVY, True)
    steps = [("技术研发", "钙钛矿叠层\n固态电池 · SiC/GaN", BLUE), ("本地装备", "光伏母机\n海风整机 · 电解槽", GREEN), ("开放场景", "首台套清单\n零碳园区试验田", ORANGE), ("绿色出海", "碳足迹中心\n数字电池护照 · CBAM", DEEP_GREEN)]
    for i, (t, b, accent) in enumerate(steps):
        x = 0.75 + i * 3.08
        shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.45, 2.5, 2.5, [LIGHT_BLUE, LIGHT_GREEN, LIGHT_ORANGE, LIGHT_GREEN][i], accent, True)
        text(slide, f"0{i+1}", x + 0.2, 2.72, 0.55, 0.3, 16, accent, True)
        text(slide, t, x + 0.2, 3.22, 2.1, 0.35, 18, NAVY, True)
        text(slide, b, x + 0.2, 3.86, 2.05, 0.65, 12, MUTED)
        if i < 3:
            connector(slide, x + 2.5, 3.7, x + 3.08, 3.7, accent, 2)
    card(slide, 1.1, 5.45, 3.35, 0.9, "规模目标", "新能源装备产值突破 5500 亿元", BLUE, LIGHT_BLUE)
    card(slide, 4.98, 5.45, 3.35, 0.9, "应用机制", "发布本地零碳场景清单", GREEN, LIGHT_GREEN)
    card(slide, 8.85, 5.45, 3.35, 0.9, "合规能力", "打通数字电池护照与 CBAM", ORANGE, LIGHT_ORANGE)


def slide_guarantee():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "五、保障措施——汇聚能源高质量发展关键要素")
    pillars = [("组织与机制", "专班统筹多规合一\n指标纳入高质量考核", BLUE), ("碳排放双控", "能碳账户\n重大项目碳前置审查", GREEN), ("要素与金融", "产业母基金\n新能源 REITs", ORANGE), ("科技与人才", "揭榜挂帅\n太湖人才计划", DEEP_GREEN), ("市场与社会", "现货交易\n全域碳普惠", BLUE)]
    for i, (t, b, accent) in enumerate(pillars):
        x = 0.52 + i * 2.55
        shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.0, 2.25, 3.75, WHITE, LINE, True)
        shape(slide, MSO_SHAPE.OVAL, x + 0.72, 2.38, 0.8, 0.8, accent, accent)
        text(slide, str(i + 1), x + 0.72, 2.59, 0.8, 0.25, 17, WHITE, True, PP_ALIGN.CENTER)
        text(slide, t, x + 0.22, 3.52, 1.8, 0.35, 15, NAVY, True, PP_ALIGN.CENTER)
        text(slide, b, x + 0.22, 4.25, 1.8, 0.55, 11, MUTED, False, PP_ALIGN.CENTER)
        pill(slide, ["统筹", "约束", "投入", "攻关", "协同"][i], x + 0.58, 5.15, 1.1, [LIGHT_BLUE, LIGHT_GREEN, LIGHT_ORANGE, LIGHT_GREEN, LIGHT_BLUE][i], accent)
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.5, 6.35, 10.3, 0.45, NAVY, NAVY, True)
    text(slide, "把规划目标转化为年度任务、项目清单和部门责任，让每一项指标都有牵引机制", 1.75, 6.46, 9.8, 0.2, 11, WHITE, True, PP_ALIGN.CENTER)


def slide_end():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    base(slide, "", dark=True)
    add_city_network(slide)
    shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, NAVY, NAVY)
    add_city_network(slide)
    text(slide, "向“新”而行  ·  依“绿”而兴", 0.8, 1.8, 11.7, 0.7, 34, WHITE, True, PP_ALIGN.CENTER)
    shape(slide, MSO_SHAPE.RECTANGLE, 5.85, 2.85, 1.65, 0.07, GREEN, GREEN)
    text(slide, "全力打造全国领先的现代新型城市能源体系“无锡样板”", 1.15, 3.35, 11.0, 0.42, 17, RGBColor(193, 213, 233), False, PP_ALIGN.CENTER)
    pill(slide, "汇报完毕，敬请各位领导与专家批评指正！", 4.35, 5.45, 4.6, RGBColor(25, 66, 88), RGBColor(133, 233, 201))


slide_cover()
slide_summary()
slide_supply()
slide_efficiency()
slide_bottleneck()
slide_swot()
slide_strategy()
slide_targets()
slide_tasks_1()
slide_tasks_2()
slide_tasks_3()
slide_tasks_4()
slide_tasks_5()
slide_tasks_6()
slide_guarantee()
slide_end()

prs.core_properties.title = "“十五五”时期无锡能源发展规划研究思路"
prs.core_properties.subject = "无锡现代新型城市能源体系"
prs.core_properties.author = "能源规划专项研究组"
prs.save("无锡能源发展规划研究思路.pptx")
print("created", len(prs.slides), "slides")