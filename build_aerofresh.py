"""Генератор тренировочной презентации AeroFresh (кейс-чемпионат «Ключ»).

Запуск:  python3 build_aerofresh.py  ->  AeroFresh_training.pptx
Всё в [квадратных скобках] автоматически подсвечивается жёлтым — это поля,
которые команда должна заполнить.
"""
import re

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ---------- палитра и шрифт ----------
GREEN = RGBColor(0x1F, 0xA8, 0x5B)       # единственный акцент — свежий зелёный
GREEN_TINT = RGBColor(0xE9, 0xF7, 0xEF)  # светлая подложка карточек
INK = RGBColor(0x1E, 0x29, 0x24)         # основной текст
MUTED = RGBColor(0x5F, 0x6B, 0x66)       # подписи
LINE = RGBColor(0xD5, 0xDD, 0xD9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
HL = "FFE45C"                            # жёлтая заливка плейсхолдеров
FONT = "Arial"

W, H = Inches(13.333), Inches(7.5)
MX = Inches(0.6)                         # поля слева/справа
CW = W - 2 * MX                          # ширина контента

prs = Presentation()
prs.slide_width, prs.slide_height = W, H
BLANK = prs.slide_layouts[6]
PH_RE = re.compile(r"(\[[^\]]*\])")


# ---------- низкоуровневые помощники ----------
def _highlight(run):
    rPr = run._r.get_or_add_rPr()
    hl = etree.SubElement(rPr, qn("a:highlight"))
    clr = etree.SubElement(hl, qn("a:srgbClr"))
    clr.set("val", HL)
    # порядок в rPr: ln, fill, effect, highlight, ..., latin — highlight до latin
    latin = rPr.find(qn("a:latin"))
    if latin is not None:
        latin.addprevious(hl)


def add_runs(p, text, size, color=INK, bold=False, italic=False):
    """Добавляет текст в абзац; фрагменты [..] получают жёлтую подсветку."""
    for part in PH_RE.split(text):
        if not part:
            continue
        r = p.add_run()
        r.text = part
        f = r.font
        f.name, f.size, f.bold, f.italic = FONT, Pt(size), bold, italic
        f.color.rgb = color
        if PH_RE.fullmatch(part):
            f.color.rgb = INK
            _highlight(r)


def set_bullet(p, char="•", color=GREEN, indent=0.28):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(Inches(indent)))
    pPr.set("indent", str(-Inches(indent)))
    for tag in ("a:buClr", "a:buFont", "a:buChar", "a:buNone"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    buClr = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(buClr, qn("a:srgbClr")).set("val", str(color))
    etree.SubElement(pPr, qn("a:buFont")).set("typeface", FONT)
    etree.SubElement(pPr, qn("a:buChar")).set("char", char)


def text(slide, x, y, w, h, paras, size=18, color=INK, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space=6, margin=0,
         line=None):
    """paras: строка или список; элемент списка — строка или dict
    {t, size, color, bold, italic, bullet, space}."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(margin)
    tf.margin_top = tf.margin_bottom = Inches(0)
    if isinstance(paras, str):
        paras = [paras]
    for i, item in enumerate(paras):
        d = {"t": item} if isinstance(item, str) else item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = d.get("align", align)
        if i:
            p.space_before = Pt(d.get("space", space))
        if line:
            p.line_spacing = line
        add_runs(p, d["t"], d.get("size", size), d.get("color", color),
                 d.get("bold", bold), d.get("italic", False))
        if d.get("bullet"):
            set_bullet(p, d["bullet"] if isinstance(d["bullet"], str) else "•")
    return tb


def box(slide, x, y, w, h, fill=GREEN_TINT, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
        radius=0.08, dash=False):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1.5)
        if dash:
            s.line.dash_style = 4  # MSO_LINE.DASH
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    s.shadow.inherit = False
    eff = s._element.find(".//" + qn("a:effectRef"))
    if eff is not None:
        eff.set("idx", "0")  # LibreOffice рисует тень по effectRef даже при пустом effectLst
    s.text_frame.text = ""
    return s


def shape_text(s, paras, size=16, color=INK, bold=False, align=PP_ALIGN.CENTER,
               anchor=MSO_ANCHOR.MIDDLE, margin=0.1, space=4):
    tf = s.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(margin)
    tf.margin_top = tf.margin_bottom = Inches(0.05)
    if isinstance(paras, str):
        paras = [paras]
    for i, item in enumerate(paras):
        d = {"t": item} if isinstance(item, str) else item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = d.get("align", align)
        if i:
            p.space_before = Pt(d.get("space", space))
        add_runs(p, d["t"], d.get("size", size), d.get("color", color),
                 d.get("bold", bold), d.get("italic", False))
    return s


def arrow(slide, x1, y1, x2, y2, color=GREEN, width=2.5):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    c.line.color.rgb = color
    c.line.width = Pt(width)
    ln = c.line._get_or_add_ln()
    tail = etree.SubElement(ln, qn("a:tailEnd"))
    tail.set("type", "triangle")
    tail.set("w", "med")
    tail.set("len", "med")
    return c


def table(slide, x, y, w, rows, col_w, size=15, row_h=0.5, header_size=None,
          bold_first_col=False, highlight_rows=()):
    nrows, ncols = len(rows), len(rows[0])
    gt = slide.shapes.add_table(nrows, ncols, x, y, w, Inches(row_h * nrows))
    tbl = gt.table
    # убираем встроенный стиль-полосатость: красим ячейки сами
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "1")
    tblPr.set("bandRow", "0")
    total = sum(col_w)
    for i, cw in enumerate(col_w):
        tbl.columns[i].width = Emu(int(w * cw / total))
    for r in range(nrows):
        tbl.rows[r].height = Inches(row_h)
        for c in range(ncols):
            cell = tbl.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.12)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = GREEN
            elif r in highlight_rows:
                cell.fill.fore_color.rgb = RGBColor(0xC9, 0xEC, 0xD7)
            else:
                cell.fill.fore_color.rgb = WHITE if r % 2 else GREEN_TINT
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            if r == 0:
                add_runs(p, rows[r][c], header_size or size, WHITE, True)
            else:
                add_runs(p, rows[r][c], size, INK,
                         bold=(bold_first_col and c == 0) or r in highlight_rows)
    return gt


def title(slide, t, kicker=None):
    """Единое положение заголовка на всех контентных слайдах."""
    if kicker:
        text(slide, MX, Inches(0.4), CW, Inches(0.35), kicker, size=14,
             color=GREEN, bold=True)
    text(slide, MX, Inches(0.72), CW, Inches(0.9), t, size=32, bold=True,
         anchor=MSO_ANCHOR.TOP)


def page_no(slide, n):
    text(slide, W - MX - Inches(1.5), H - Inches(0.45), Inches(1.5), Inches(0.3),
         n, size=11, color=MUTED, align=PP_ALIGN.RIGHT)
    text(slide, MX, H - Inches(0.45), Inches(4), Inches(0.3),
         "AeroFresh · кейс-чемпионат «Ключ»", size=11, color=MUTED)


def notes(slide, t):
    slide.notes_slide.notes_text_frame.text = t


def new_slide(n=None):
    s = prs.slides.add_slide(BLANK)
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = WHITE
    if n:
        page_no(s, n)
    return s


def num_badge(slide, x, y, label, d=0.55, size=18, fill=GREEN, color=WHITE):
    c = box(slide, x, y, Inches(d), Inches(d), fill=fill, shape=MSO_SHAPE.OVAL)
    shape_text(c, label, size=size, color=color, bold=True, margin=0)
    return c


# =====================================================================
# Слайд 1. Титульный
# =====================================================================
s = prs.slides.add_slide(BLANK)
s.background.fill.solid()
s.background.fill.fore_color.rgb = GREEN
# декоративные круги-«датчики»
for (cx, cy, d, a) in [(10.3, 0.9, 3.6, 0.10), (11.6, 4.3, 2.4, 0.16), (9.2, 4.9, 1.2, 0.22)]:
    o = box(s, Inches(cx), Inches(cy), Inches(d), Inches(d), fill=WHITE, shape=MSO_SHAPE.OVAL)
    fill_xml = o.fill._xPr.find(qn("a:solidFill"))
    clr = fill_xml[0]
    etree.SubElement(clr, qn("a:alpha")).set("val", str(int(a * 100000)))
text(s, MX, Inches(1.0), Inches(4), Inches(0.4), "Кейс-чемпионат «Ключ»",
     size=18, color=WHITE, bold=True)
text(s, MX, Inches(2.0), Inches(9.0), Inches(1.2), "AeroFresh", size=72,
     color=WHITE, bold=True)
text(s, MX, Inches(3.3), Inches(8.6), Inches(1.3),
     "Спасаем еду до того, как она испортится", size=36, color=WHITE)
text(s, MX, Inches(5.2), Inches(8.6), Inches(1.4), [
    {"t": "Категория РСХБ: Сенсоры и IoT / Фудтех", "size": 20, "color": WHITE},
    {"t": "Команда: [ФИО, класс]", "size": 20, "color": WHITE, "space": 10},
])
notes(s, "Представляемся: команда, школа, категория РСХБ. Одна фраза о проекте: "
         "AeroFresh — датчик на ящике, который заранее говорит магазину, что товар "
         "скоро испортится, и подсказывает, как его продать. ~20 секунд.")

# =====================================================================
# Слайд 2. Проблема и ЦА
# =====================================================================
s = new_slide("2")
title(s, "Магазины узнают о порче, когда товар уже не продать", "Проблема и ЦА")
# крупные цифры
for i, (big, lab) in enumerate([
        ("13%", "еды теряется ещё до прилавка¹"),
        ("19%", "выбрасывается в рознице, общепите и дома²")]):
    y = Inches(1.85 + i * 1.7)
    box(s, MX, y, Inches(4.0), Inches(1.55))
    text(s, MX + Inches(0.3), y + Inches(0.12), Inches(3.5), Inches(0.75), big,
         size=44, bold=True, color=GREEN)
    text(s, MX + Inches(0.3), y + Inches(0.85), Inches(3.6), Inches(0.5), lab,
         size=14, color=INK)
text(s, MX, Inches(5.35), Inches(4.0), Inches(0.9),
     "Больше всего списывается скоропорт: ягоды, зелень, овощи",
     size=16, bold=True)
# правая колонка
RX = MX + Inches(4.5)
RW = CW - Inches(4.5)
text(s, RX, Inches(1.85), RW, Inches(0.9), [
    {"t": "Целевая аудитория", "size": 14, "color": GREEN, "bold": True},
    {"t": "Категорийные менеджеры и директора магазинов локальных торговых сетей",
     "size": 18, "space": 4},
])
q = box(s, RX, Inches(2.95), RW, Inches(1.75), fill=GREEN_TINT)
text(s, RX + Inches(0.3), Inches(3.08), RW - Inches(0.6), Inches(1.55), [
    {"t": "«В прошлую пятницу почти два ящика петрушки ушли в мусорку: продавец "
          "поставил их у обогревателя, а я заметил только вечером»",
     "size": 17, "italic": True},
    {"t": "— владелец 3 овощных точек", "size": 13, "color": MUTED, "space": 6},
])
text(s, RX, Inches(4.95), RW, Inches(1.1), [
    {"t": "Что не подтвердилось", "size": 14, "color": GREEN, "bold": True},
    {"t": "[заполнить]", "size": 16, "bullet": True, "space": 4},
    {"t": "Небольшим точкам 5 000 ₽/мес кажется дорого", "size": 16, "bullet": True,
     "space": 4},
])
# сноски
text(s, MX, Inches(6.38), CW, Inches(0.5), [
    {"t": "¹ ФАО ООН. The State of Food and Agriculture 2019: Moving forward on food "
          "loss and waste reduction.", "size": 10, "color": MUTED},
    {"t": "² ЮНЕП. Food Waste Index Report 2024.", "size": 10, "color": MUTED, "space": 1},
])
notes(s, "Масштаб: по ФАО ~13% еды теряется до прилавка, по ЮНЕП ещё 19% "
         "выбрасывается в рознице, общепите и дома. Больнее всего — скоропорт. "
         "Наша ЦА — менеджеры и директора локальных сетей. Цитата из интервью "
         "показывает боль: о порче узнают слишком поздно. Честно говорим, что не "
         "подтвердилось, — жюри это ценит.")

# =====================================================================
# Слайд 3. Решение
# =====================================================================
s = new_slide("3")
title(s, "AeroFresh — многоразовый датчик свежести на ящике", "Решение")
text(s, MX, Inches(2.0), Inches(5.3), Inches(3.8), [
    {"t": "Измеряет газы и микроклимат прямо в ящике, оценивает, сколько осталось "
          "до порчи, и подсказывает, когда и на сколько уценить товар.", "size": 20},
    {"t": "Ставится на ящик при приёмке", "size": 18, "bullet": True, "space": 16},
    {"t": "Работает многократно — переставляется на новые партии", "size": 18,
     "bullet": True},
    {"t": "Результат — готовое действие, а не график", "size": 18, "bullet": True},
], space=8)
# схема «измеряет → выдаёт»
GX = MX + Inches(6.0)
cards = [("CO₂", "дыхание продукта"), ("ЛОС", "летучие вещества"),
         ("t °C", "температура"), ("RH %", "влажность")]
cw_, ch_ = Inches(1.7), Inches(1.4)
for i, (a, b) in enumerate(cards):
    cx = GX + (i % 2) * (cw_ + Inches(0.2))
    cy = Inches(1.95) + (i // 2) * (ch_ + Inches(0.2))
    c = box(s, cx, cy, cw_, ch_)
    shape_text(c, [{"t": a, "size": 24, "bold": True, "color": GREEN},
                   {"t": b, "size": 13, "color": MUTED}], space=2)
ax = GX + 2 * cw_ + Inches(0.2)
arrow(s, ax + Inches(0.15), Inches(1.95) + ch_ + Inches(0.1),
      ax + Inches(0.75), Inches(1.95) + ch_ + Inches(0.1))
OX = ax + Inches(0.9)
OW = W - MX - OX
o1 = box(s, OX, Inches(1.95), OW, ch_, fill=GREEN)
shape_text(o1, [{"t": "~12 ч", "size": 26, "bold": True, "color": WHITE},
                {"t": "до порчи", "size": 13, "color": WHITE}], space=2)
o2 = box(s, OX, Inches(1.95) + ch_ + Inches(0.2), OW, ch_, fill=GREEN)
shape_text(o2, [{"t": "−30%", "size": 26, "bold": True, "color": WHITE},
                {"t": "уценить сейчас", "size": 13, "color": WHITE}], space=2)
text(s, GX, Inches(5.25), W - MX - GX, Inches(0.9),
     "Датчик «нюхает» сам продукт, а не только воздух в зале — поэтому видит "
     "порчу раньше человека.", size=15, color=MUTED)
notes(s, "Суть решения в одном предложении: датчик на ящике + прогноз + "
         "подсказка по уценке. Четыре параметра: CO₂ и летучие вещества "
         "показывают, как «дышит» и портится продукт; температура и влажность — "
         "условия хранения. На выходе не график, а действие: сколько осталось и "
         "какую скидку дать.")

# =====================================================================
# Слайд 4. Как это работает
# =====================================================================
s = new_slide("4")
title(s, "Как это работает", "Звено цепочки: доведение до потребителя (ритейл)")
steps = [("Датчик", "на ящике"), ("Wi-Fi / SIM", "передача данных"),
         ("Облако", "хранение и расчёт"), ("Прогноз", "часы до порчи"),
         ("Уведомление", "менеджеру зала")]
n = len(steps)
gap = Inches(0.5)
bw = int((CW - gap * (n - 1)) / n)
bh = Inches(1.5)
by = Inches(1.95)
for i, (a, b) in enumerate(steps):
    bx = MX + i * (bw + gap)
    fill = GREEN if i in (0, n - 1) else GREEN_TINT
    col = WHITE if i in (0, n - 1) else INK
    sub = WHITE if i in (0, n - 1) else MUTED
    c = box(s, bx, by, bw, bh, fill=fill)
    shape_text(c, [{"t": a, "size": 19, "bold": True, "color": col},
                   {"t": b, "size": 14, "color": sub}], space=4)
    if i < n - 1:
        arrow(s, bx + bw + Inches(0.08), by + bh // 2,
              bx + bw + gap - Inches(0.08), by + bh // 2)
# уведомление — макет сообщения
ny = Inches(3.95)
phone = box(s, MX, ny, Inches(7.4), Inches(2.4), fill=WHITE, line=LINE)
text(s, MX + Inches(0.3), ny + Inches(0.2), Inches(6.8), Inches(0.35),
     "AeroFresh · уведомление", size=13, color=MUTED, bold=True)
text(s, MX + Inches(0.3), ny + Inches(0.62), Inches(6.8), Inches(1.5), [
    {"t": "Ящик №4, клубника, у входа: осталось ~12\u00a0часов.", "size": 20, "bold": True},
    {"t": "Сделайте скидку 30% сейчас.", "size": 20, "bold": True, "color": GREEN},
    {"t": "После отзыва клиента: сразу какой ящик и что делать", "size": 13,
     "color": MUTED, "space": 10},
], space=4)
# роли
RX = MX + Inches(7.8)
RW = W - MX - RX
for i, (a, b) in enumerate([("Пользуется", "менеджер зала"),
                            ("Платит", "торговая сеть")]):
    y = ny + i * Inches(1.25)
    c = box(s, RX, y, RW, Inches(1.15))
    text(s, RX + Inches(0.3), y + Inches(0.14), RW - Inches(0.6), Inches(0.35), a,
         size=14, color=GREEN, bold=True)
    text(s, RX + Inches(0.3), y + Inches(0.55), RW - Inches(0.6), Inches(0.45), b,
         size=20, bold=True)
notes(s, "Проходим схему слева направо: датчик на ящике → данные по Wi-Fi или "
         "SIM → облако → прогноз оставшихся часов → уведомление. Показываем "
         "пример сообщения: после отзыва клиента мы сразу пишем номер ящика, "
         "место и конкретное действие. Пользователь — менеджер зала, платит "
         "торговая сеть.")

# =====================================================================
# Слайд 5. Рынок
# =====================================================================
s = new_slide("5")
title(s, "Рынок: считаем от магазина", "Рынок")
fb = box(s, MX, Inches(1.85), CW, Inches(0.75))
shape_text(fb, [{"t": "Объём = число магазинов × 9 900 ₽ × 12 мес = число магазинов × 118 800 ₽/год",
                 "size": 18, "bold": True}])
table(s, MX, Inches(2.85), Inches(8.0), [
    ["Сегмент", "Число магазинов", "Объём, ₽/год"],
    ["TAM — продуктовые магазины РФ", "[N, источник]", "[N × 118 800]"],
    ["SAM — региональные сети", "[N, источник]", "[N × 118 800]"],
    ["SOM — цель на старте", "50", "≈ 6 млн"],
], col_w=[3.6, 2.2, 2.2], size=16, row_h=0.8, bold_first_col=True,
    highlight_rows=(3,))
RX = MX + Inches(8.4)
RW = W - MX - RX
for i, (hd, items) in enumerate([
        ("Тренды", ["Борьба с пищевыми отходами", "Автоматизация ритейла"]),
        ("Барьеры", ["Интеграция с кассой", "Требования к торговому оборудованию"])]):
    y = Inches(2.85) + i * Inches(1.65)
    box(s, RX, y, RW, Inches(1.55))
    text(s, RX + Inches(0.25), y + Inches(0.15), RW - Inches(0.5), Inches(1.3),
         [{"t": hd, "size": 14, "color": GREEN, "bold": True}] +
         [{"t": it, "size": 15, "bullet": True, "space": 3} for it in items])
notes(s, "Считаем снизу вверх от цены за магазин: 9 900 ₽ × 12 = 118 800 ₽ в год. "
         "SOM — 50 магазинов, около 6 млн ₽ в год. SAM и TAM заполнить реальными "
         "данными с источником (Росстат, отраслевые обзоры). Тренды работают на нас, "
         "барьеры — интеграция с кассой и требования к оборудованию.")

# =====================================================================
# Слайд 6. Альтернативы
# =====================================================================
s = new_slide("6")
title(s, "Чем магазины пользуются сейчас", "Альтернативы")
table(s, MX, Inches(1.9), CW, [
    ["Альтернатива", "Почему не решает проблему"],
    ["Ручной обход", "Долго, субъективно, не видно, что внутри ящика"],
    ["Термодатчики", "Меряют воздух в зале, а не состояние продукта"],
    ["[Реальные конкуренты: мониторинг холодовой цепи]", "[чем отличаются]"],
    ["[Реальные конкуренты: сервисы уценки]", "[чем отличаются]"],
], col_w=[5.0, 7.1], size=17, row_h=0.7, bold_first_col=True)
d = box(s, MX, Inches(5.85), CW, Inches(0.9), fill=GREEN)
shape_text(d, [{"t": "Наше отличие: смотрим на газы самого продукта и переводим это в действие",
                "size": 20, "bold": True, "color": WHITE}])
notes(s, "Ручной обход — субъективно и поздно. Термодатчики видят воздух, а не "
         "продукт. Перед защитой обязательно вписать 1–2 реальных конкурента "
         "(мониторинг холодовой цепи, сервисы уценки) и честно сказать, чем мы "
         "отличаемся. Ключевая мысль — газы продукта плюс готовое действие.")

# =====================================================================
# Слайд 7. Бизнес-модель
# =====================================================================
s = new_slide("7")
title(s, "Бизнес-модель: подписка за магазин", "Бизнес-модель")
stats = [("9 900 ₽", "подписка в месяц\nза магазин"),
         ("6 900 ₽", "наша маржа\nв месяц с магазина"),
         ("≈14 000 ₽", "чистая выгода\nклиента в месяц")]
sw = int((CW - Inches(0.6)) / 3)
for i, (a, b) in enumerate(stats):
    x = MX + i * (sw + Inches(0.3))
    fill = GREEN if i == 2 else GREEN_TINT
    c = box(s, x, Inches(1.85), sw, Inches(1.5), fill=fill)
    shape_text(c, [{"t": a, "size": 36, "bold": True, "color": WHITE if i == 2 else GREEN},
                   {"t": b, "size": 14, "color": WHITE if i == 2 else MUTED}], space=2)
tw = int((CW - Inches(0.4)) / 2)
text(s, MX, Inches(3.65), tw, Inches(0.35), "Для нас", size=14, color=GREEN, bold=True)
table(s, MX, Inches(4.05), tw, [
    ["Статья", "₽/мес на магазин"],
    ["Подписка (10 датчиков включены)", "9 900"],
    ["Наши затраты", "≈ 3 000"],
    ["Маржа", "≈ 6 900"],
], col_w=[3.6, 2.2], size=15, row_h=0.56, highlight_rows=(3,))
x2 = MX + tw + Inches(0.4)
text(s, x2, Inches(3.65), tw, Inches(0.35), "Для клиента", size=14, color=GREEN, bold=True)
table(s, x2, Inches(4.05), tw, [
    ["Статья", "₽/мес"],
    ["Снижение списаний на 30%", "≈ 24 000"],
    ["Подписка", "− 9 900"],
    ["Чистая выгода", "≈ 14 000"],
], col_w=[3.6, 2.2], size=15, row_h=0.56, highlight_rows=(3,))
notes(s, "Модель — подписка 9 900 ₽ в месяц за магазин, 10 датчиков включены. "
         "Наши затраты около 3 000 ₽, маржа около 6 900 ₽. Для клиента: если "
         "списания падают на 30%, это ~24 000 ₽ в месяц, после подписки чистыми "
         "~14 000 ₽. Подробный расчёт — в приложении П2.")

# =====================================================================
# Слайд 8. Прототип и проверка
# =====================================================================
s = new_slide("8")
title(s, "Прототип и проверка с пользователями", "Прототип")
pw = Inches(2.7)
for i, lab in enumerate(["[Скриншот: список партий с таймерами]",
                         "[Скриншот: уведомление]"]):
    x = MX + i * (pw + Inches(0.3))
    f = box(s, x, Inches(1.85), pw, Inches(3.4), fill=WHITE, line=GREEN, dash=True)
    shape_text(f, lab, size=15)
gx = MX + 2 * (pw + Inches(0.3))
gw = Inches(3.0)
g = box(s, gx, Inches(1.85), gw, Inches(3.4), fill=WHITE, line=GREEN, dash=True)
shape_text(g, [{"t": "Практический тест", "size": 14, "bold": True, "color": GREEN},
               {"t": "CO₂-датчик в контейнере с клубникой", "size": 14},
               {"t": "[график роста CO₂ и фото порчи по дням]", "size": 14}], space=8)
RX = gx + gw + Inches(0.4)
RW = W - MX - RX
text(s, RX, Inches(1.85), RW, Inches(3.0), [
    {"t": "Проверка", "size": 14, "color": GREEN, "bold": True},
    {"t": "[сколько человек видели, что сказали]", "size": 15, "space": 4},
    {"t": "Изменили после отзывов", "size": 14, "color": GREEN, "bold": True, "space": 16},
    {"t": "Уведомление указывает конкретный ящик и действие", "size": 15,
     "bullet": True, "space": 4},
    {"t": "Пересмотрена цена", "size": 15, "bullet": True, "space": 4},
])
text(s, MX, Inches(5.6), CW, Inches(0.9),
     "Цикл: показали макет → собрали отзывы → поменяли уведомление и цену → "
     "проверили CO₂ на настоящей клубнике.", size=16, color=MUTED)
notes(s, "Показываем макет: список партий с таймерами и уведомление. Говорим, "
         "скольким людям показали и что они сказали. Главное — что изменили: "
         "уведомление стало конкретным, цена пересмотрена. Практический тест: "
         "CO₂-датчик в контейнере с клубникой — показываем, как CO₂ растёт вместе "
         "с порчей.")

# =====================================================================
# Слайд 9. План теста
# =====================================================================
s = new_slide("9")
title(s, "План пилота на 4–6 недель", "План теста")
phases = [("Недели 1–2", "Базовый уровень",
           "В 1 магазине фиксируем списания вручную"),
          ("Недели 3–4", "Запуск датчиков",
           "2–3 датчика (CO₂-модуль + ESP32) и Telegram-бот"),
          ("Недели 5–6", "Сравнение",
           "Сравниваем списания до и после")]
pw3 = int((CW - Inches(0.6)) / 3)
for i, (wk, hd, body) in enumerate(phases):
    x = MX + i * (pw3 + Inches(0.3))
    num_badge(s, x, Inches(1.9), str(i + 1))
    if i < 2:
        arrow(s, x + Inches(0.65), Inches(2.175), x + pw3 + Inches(0.2), Inches(2.175),
              color=LINE, width=2)
    c = box(s, x, Inches(2.7), pw3, Inches(2.0))
    text(s, x + Inches(0.25), Inches(2.85), pw3 - Inches(0.5), Inches(1.7), [
        {"t": wk, "size": 14, "color": GREEN, "bold": True},
        {"t": hd, "size": 20, "bold": True, "space": 4},
        {"t": body, "size": 15, "space": 6},
    ])
ok = box(s, MX, Inches(5.0), Inches(4.2), Inches(1.3), fill=GREEN)
shape_text(ok, [{"t": "Критерий успеха", "size": 14, "color": WHITE, "bold": True},
                {"t": "списания −20% и больше", "size": 20, "color": WHITE, "bold": True}])
rx = MX + Inches(4.5)
rk = box(s, rx, Inches(5.0), W - MX - rx, Inches(1.3), fill=WHITE, line=LINE)
text(s, rx + Inches(0.3), Inches(5.15), W - MX - rx - Inches(0.6), Inches(1.05), [
    {"t": "Риск: менеджер не реагирует на уведомления", "size": 16, "bold": True},
    {"t": "Ответ: короткая рекомендация с уже посчитанной скидкой — одно действие "
          "вместо анализа", "size": 15, "space": 4},
])
notes(s, "План пилота: две недели меряем базовые списания вручную, две недели "
         "работают 2–3 датчика с Telegram-ботом, две недели сравниваем. Успех — "
         "снижение списаний на 20% и больше. Главный риск — менеджер игнорирует "
         "уведомления; поэтому даём готовую скидку, а не сырые данные.")

# =====================================================================
# Слайд 10. Команда и ИИ
# =====================================================================
s = new_slide("10")
title(s, "Команда и использование ИИ", "Команда")
mw = Inches(2.3)
for i in range(3):
    x = MX + i * (mw + Inches(0.3))
    c = box(s, x, Inches(1.85), mw, Inches(3.6))
    ph = box(s, x + (mw - Inches(1.3)) // 2, Inches(2.1), Inches(1.3), Inches(1.3),
             fill=WHITE, line=GREEN, shape=MSO_SHAPE.OVAL, dash=True)
    shape_text(ph, "[фото]", size=14)
    text(s, x + Inches(0.2), Inches(3.6), mw - Inches(0.4), Inches(1.5), [
        {"t": "[ФИО]", "size": 18, "bold": True, "align": PP_ALIGN.CENTER},
        {"t": "[школа, класс]", "size": 14, "align": PP_ALIGN.CENTER, "space": 6},
        {"t": "[роль в проекте]", "size": 14, "align": PP_ALIGN.CENTER, "space": 6},
    ])
ax = MX + 3 * (mw + Inches(0.3)) + Inches(0.1)
aw = W - MX - ax
ai = box(s, ax, Inches(1.85), aw, Inches(3.6), fill=WHITE, line=LINE)
text(s, ax + Inches(0.3), Inches(2.05), aw - Inches(0.6), Inches(3.2), [
    {"t": "ИИ в проекте", "size": 14, "color": GREEN, "bold": True},
    {"t": "[Инструмент] — [для чего]", "size": 16, "bullet": True, "space": 8},
    {"t": "[Инструмент] — [для чего]", "size": 16, "bullet": True, "space": 8},
    {"t": "[Инструмент] — [для чего]", "size": 16, "bullet": True, "space": 8},
    {"t": "Промпты — в приложении П4", "size": 14, "color": MUTED, "space": 14},
])
text(s, MX, Inches(5.85), CW, Inches(0.6),
     "Спасибо! Готовы ответить на вопросы.", size=22, bold=True, color=GREEN)
notes(s, "Каждый участник называет себя и свою роль. Коротко: какие ИИ-инструменты "
         "использовали и для чего (поиск данных, тексты интервью, код бота, "
         "дизайн макета). Все промпты — в приложении П4. Благодарим и переходим к "
         "вопросам.")


# =====================================================================
# Приложения
# =====================================================================
def appendix(code, t):
    s = new_slide(code)
    title(s, t, f"Приложение {code}")
    return s


# П1. CJM
s = appendix("П1", "Путь клиента (CJM)")
cjm = ["Менеджер находит испорченную клубнику слишком поздно",
       "Узнаёт об AeroFresh",
       "При приёмке ставит датчики на ящики",
       "Получает уведомление «осталось ~12 часов» и делает уценку",
       "Товар продан, списаний меньше"]
n = len(cjm)
gap = Inches(0.35)
cw5 = int((CW - gap * (n - 1)) / n)
for i, t in enumerate(cjm):
    x = MX + i * (cw5 + gap)
    num_badge(s, x + (cw5 - Inches(0.6)) // 2, Inches(2.0), str(i + 1), d=0.6, size=20,
              fill=GREEN if i in (0, n - 1) else GREEN_TINT,
              color=WHITE if i in (0, n - 1) else GREEN)
    if i < n - 1:
        arrow(s, x + (cw5 + Inches(0.6)) // 2 + Inches(0.1), Inches(2.3),
              x + cw5 + gap + (cw5 - Inches(0.6)) // 2 - Inches(0.1), Inches(2.3),
              color=LINE, width=2)
    c = box(s, x, Inches(2.9), cw5, Inches(2.4), fill=GREEN_TINT if i != n - 1 else GREEN)
    shape_text(c, t, size=16, color=WHITE if i == n - 1 else INK,
               bold=(i == n - 1), anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, margin=0.18)
text(s, MX, Inches(5.7), CW, Inches(0.6),
     "Точка боли — шаг 1; точка ценности — шаг 4: решение принимается за секунды.",
     size=16, color=MUTED)
notes(s, "Путь менеджера: от боли (нашёл порчу поздно) через знакомство и установку "
         "датчиков к уведомлению и уценке. Итог — товар продан, списаний меньше.")

# П2. Расчёты
s = appendix("П2", "Расчёты и источники")
table(s, MX, Inches(1.8), CW, [
    ["Показатель", "Значение", "Источник / расчёт"],
    ["Списания скоропорта в магазине", "80 000 ₽/мес", "[подтвердить интервью]"],
    ["Снижение списаний на 30%", "24 000 ₽/мес", "80 000 × 30%; цель теста"],
    ["Подписка", "9 900 ₽/мес", "Наша цена"],
    ["Чистая выгода клиента", "≈ 14 000 ₽/мес", "24 000 − 9 900"],
    ["Датчик", "≈ 3 500 ₽/шт.", "[смета комплектующих]"],
    ["10 датчиков на магазин", "35 000 ₽", "3 500 × 10"],
    ["Наши затраты на магазин", "≈ 3 000 ₽/мес", "Амортизация 1 500 + облако и поддержка 1 500"],
    ["Безубыточность", "≈ 22 магазина", "150 000 ÷ 6 900; постоянные расходы — допущение"],
], col_w=[4.0, 2.4, 5.7], size=15, row_h=0.5, bold_first_col=True,
    highlight_rows=(8,))
notes(s, "Каждая цифра с источником или формулой. Главное допущение — списания "
         "80 000 ₽ в месяц, его подтверждаем интервью. Безубыточность: 150 000 ₽ "
         "постоянных расходов делим на маржу 6 900 ₽ — около 22 магазинов.")

# П3. Технологии
s = appendix("П3", "Технологии")
layers = [("Сенсоры", "CO₂ — SCD40 · летучие вещества — SGP40 · температура и влажность"),
          ("Связь", "Микроконтроллер ESP32 · Wi-Fi или SIM-модуль"),
          ("Облако", "Yandex Cloud: хранение данных и расчёт прогноза"),
          ("Интерфейс (MVP)", "Telegram-бот для менеджера зала"),
          ("Прогноз", "Сначала пороговые правила → затем модель на временных рядах")]
lh = Inches(0.82)
for i, (a, b) in enumerate(layers):
    y = Inches(1.85) + i * (lh + Inches(0.12))
    la = box(s, MX, y, Inches(3.0), lh, fill=GREEN)
    shape_text(la, a, size=18, bold=True, color=WHITE)
    lb = box(s, MX + Inches(3.2), y, CW - Inches(3.2), lh)
    shape_text(lb, b, size=17, align=PP_ALIGN.LEFT, margin=0.3)
notes(s, "Стек снизу вверх: недорогие сенсоры SCD40 и SGP40, ESP32 с Wi-Fi или "
         "SIM, Yandex Cloud, Telegram-бот как MVP. Прогноз начинаем с порогов, "
         "по мере накопления данных переходим к модели на временных рядах.")

# П4. ИИ
s = appendix("П4", "Как мы использовали ИИ")
table(s, MX, Inches(1.8), CW, [
    ["Инструмент", "Задача", "Промпт", "Что использовали"],
    ["[заполнить]", "[заполнить]", "[заполнить]", "[заполнить]"],
    ["[заполнить]", "[заполнить]", "[заполнить]", "[заполнить]"],
    ["[заполнить]", "[заполнить]", "[заполнить]", "[заполнить]"],
    ["[заполнить]", "[заполнить]", "[заполнить]", "[заполнить]"],
    ["[заполнить]", "[заполнить]", "[заполнить]", "[заполнить]"],
], col_w=[2.2, 2.8, 4.3, 2.8], size=15, row_h=0.75)
notes(s, "Для каждого инструмента: какую задачу решали, точный промпт и что из "
         "ответа реально вошло в проект. Жюри важно видеть, что ИИ — помощник, "
         "а решения принимала команда.")

# П5. Интервью — 15 строк
s = appendix("П5", "Интервью: 15 респондентов, из них 3+ практика")
rows = [["№", "Роль", "Главная мысль", "Подтвердило / опровергло"]]
for i in range(1, 16):
    rows.append([str(i), "[роль]" if i > 3 else "[практик: роль]",
                 "[главная мысль]", "[подтв. / опроверг.]"])
table(s, MX, Inches(1.6), CW, rows, col_w=[0.6, 3.0, 5.6, 2.9], size=11,
      header_size=12, row_h=0.305)
notes(s, "Таблица всех 15 интервью. Первые три строки — практики (директора, "
         "категорийные менеджеры, владельцы точек). По каждому — главная мысль и "
         "что она подтвердила или опровергла в гипотезе.")

prs.save("AeroFresh_training.pptx")
print("saved AeroFresh_training.pptx,", len(prs.slides), "slides")
