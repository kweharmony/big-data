"""
ШАБЛОН ДЛЯ ИИ: генерация отчёта .docx по ГОСТ-оформлению (учебные лабораторные).

КАК ПОЛЬЗОВАТЬСЯ (инструкция для ИИ)
------------------------------------
1. Скопируй этот файл в папку проекта, например `make_report.py`.
2. Заполни блок SETTINGS (название работы, ФИО, группа, папка со скриншотами).
3. Внизу, в блоке «СОДЕРЖАНИЕ ОТЧЁТА», собери документ вызовами функций:
       heading("1. Разведка")          — раздел (попадает в оглавление)
       subheading("2.1. Дубли")        — подзаголовок по центру, жирный
       text("Обычный абзац ...")       — текст с отступом, по ширине
       picture("q1.png", "Результат")  — рисунок + подпись «Рисунок N. – ...»
       listing("Запрос Q1", "select 1")— код в рамке + подпись «Листинг N. ...»
       table("Итоги", [[...], [...]])  — таблица + подписи «Таблица N.» и название
       page_break()                    — разрыв страницы
   Нумерация рисунков, листингов и таблиц ведётся автоматически.
4. Запусти:  python make_report.py
5. Скажи пользователю открыть файл и обновить оглавление:
   правый клик по оглавлению → «Обновить поле» → «Обновить целиком».

ЗАВИСИМОСТЬ:  pip install python-docx

ПРАВИЛА ХОРОШЕГО ОТЧЁТА (соблюдай при генерации текста)
--------------------------------------------------------
- Пиши от первого лица, просто и коротко, как студент: «посчитал», «получилось».
- Каждый рисунок, листинг и таблица должны быть упомянуты в тексте до вставки
  (функция picture() делает это сама).
- Не копируй чужие формулировки выводов — преподаватели снимают за это баллы.
- Числа в отчёте бери только из реально выполненных команд, не выдумывай.
- Если скриншота нет, в документе появится пометка [НЕТ СКРИНШОТА: имя] —
  не убирай её, а сообщи пользователю, какого файла не хватает.
"""

import pathlib

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

# ════════════════════════════ SETTINGS ════════════════════════════
UNIVERSITY = [
    "МИНИСТЕРСТВО ЦИФРОВОГО РАЗВИТИЯ, СВЯЗИ И МАССОВЫХ КОММУНИКАЦИЙ "
    "РОССИЙСКОЙ ФЕДЕРАЦИИ",
    "Ордена Трудового Красного Знамени федеральное государственное",
    "бюджетное образовательное учреждение высшего образования",
    "МОСКОВСКИЙ ТЕХНИЧЕСКИЙ УНИВЕРСИТЕТ СВЯЗИ И ИНФОРМАТИКИ",
]
DEPARTMENT = "Кафедра «Название кафедры»"
WORK_TITLE = "Лабораторная работа № "
DISCIPLINE = "по дисциплине «Название дисциплины»"
TOPIC = "по теме: «Тема работы»"
STUDENT = "Фамилия Имя"
GROUP = "ГРУППА"
CITY_YEAR = "Москва 2026"

SHOTS = pathlib.Path("screenshots")       # папка со скриншотами
OUT = pathlib.Path("Отчет.docx")          # имя итогового файла

TITLE_PAGE = True                         # титульный лист
TOC = True                                # оглавление
FONT = "Times New Roman"
SIZE = Pt(14)
CODE_FONT = "Courier New"
CODE_SIZE = Pt(10)
LINE = 1.5                                # межстрочный интервал
INDENT = Cm(1.25)                         # абзацный отступ
PIC_WIDTH = Cm(16)                        # ширина картинок

# ═════════════════════════ ДВИЖОК (не менять) ═════════════════════════
doc = Document()
fig_n = tab_n = lst_n = 0

# поля страницы по ГОСТ: левое 3 см, правое 1,5 см, верх/низ 2 см
for s in doc.sections:
    s.left_margin, s.right_margin = Cm(3), Cm(1.5)
    s.top_margin, s.bottom_margin = Cm(2), Cm(2)

# базовый стиль документа
_normal = doc.styles["Normal"]
_normal.font.name, _normal.font.size = FONT, SIZE
_normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
_pf = _normal.paragraph_format
_pf.line_spacing, _pf.space_after, _pf.space_before = LINE, Pt(0), Pt(0)

# заголовок раздела: тот же шрифт, жирный, по центру, попадает в оглавление
_h1 = doc.styles["Heading 1"]
_h1.font.name, _h1.font.size = FONT, SIZE
_h1.font.bold, _h1.font.color.rgb = True, None
_h1.font.color.theme_color = None
_h1.paragraph_format.line_spacing = LINE
_h1.paragraph_format.space_before, _h1.paragraph_format.space_after = Pt(12), Pt(12)


def _run(p, s, font=FONT, size=SIZE, bold=False):
    r = p.add_run(s)
    r.font.name, r.font.size, r.bold = font, size, bold
    r._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    return r


def heading(t):
    """Заголовок раздела (виден в оглавлении)."""
    p = doc.add_paragraph(style="Heading 1")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, t, bold=True)
    p.runs[0].font.color.rgb = None


def subheading(t):
    """Подзаголовок: жирный по центру, в оглавление не идёт."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, t, bold=True)


def text(t, indent=True):
    """Обычный абзац: по ширине, с красной строкой."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent:
        p.paragraph_format.first_line_indent = INDENT
    _run(p, t)


def caption(t, align=WD_ALIGN_PARAGRAPH.CENTER):
    p = doc.add_paragraph()
    p.alignment = align
    _run(p, t)


def picture(file, title, width=PIC_WIDTH):
    """Рисунок со ссылкой в тексте и подписью снизу."""
    global fig_n
    path = SHOTS / file
    if not path.exists():
        text(f"[НЕТ СКРИНШОТА: {file}]")
        return
    fig_n += 1
    text(f"На рисунке {fig_n} представлено: {title[0].lower() + title[1:]}.")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(path), width=width)
    caption(f"Рисунок {fig_n}. – {title}.")


def listing(title, code):
    """Код в рамке с подписью «Листинг N. ...» сверху."""
    global lst_n
    lst_n += 1
    caption(f"Листинг {lst_n}. {title}")
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    cell = t.rows[0].cells[0]
    cell.paragraphs[0].text = ""
    for i, line in enumerate(code.strip("\n").splitlines()):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.line_spacing = 1.0
        _run(p, line, font=CODE_FONT, size=CODE_SIZE)
    doc.add_paragraph()


def table(title, rows, header=True):
    """Таблица: «Таблица N.» справа, название по центру, шапка жирная.

    rows — список списков; первая строка считается шапкой.
    """
    global tab_n
    tab_n += 1
    caption(f"Таблица {tab_n}.", WD_ALIGN_PARAGRAPH.RIGHT)
    caption(title)
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = t.cell(i, j)
            c.text = ""
            p = c.paragraphs[0]
            p.paragraph_format.line_spacing = 1.0
            _run(p, str(val), size=Pt(12), bold=header and i == 0)
    doc.add_paragraph()


def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def _title_page():
    for line in UNIVERSITY:
        caption(line)
    doc.add_paragraph()
    caption(DEPARTMENT)
    for _ in range(6):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, WORK_TITLE, bold=True)
    caption(DISCIPLINE)
    caption(TOPIC)
    for _ in range(6):
        doc.add_paragraph()
    caption("Выполнил студент группы:", WD_ALIGN_PARAGRAPH.RIGHT)
    caption(GROUP, WD_ALIGN_PARAGRAPH.RIGHT)
    caption(STUDENT, WD_ALIGN_PARAGRAPH.RIGHT)
    for _ in range(4):
        doc.add_paragraph()
    caption(CITY_YEAR)
    page_break()


def _toc():
    """Поле оглавления. Word покажет его после «Обновить поле»."""
    caption("Оглавление")
    p = doc.add_paragraph()
    r = p.add_run()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), 'TOC \\o "1-3" \\h \\z \\u')
    inner = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "Оглавление обновится по F9 (правый клик → Обновить поле)"
    inner.append(t)
    fld.append(inner)
    r._r.addprevious(fld)
    page_break()


# ═══════════════════════ СОДЕРЖАНИЕ ОТЧЁТА ═══════════════════════
# Ниже — пример. ИИ: замени его на реальные разделы работы.

if TITLE_PAGE:
    _title_page()
if TOC:
    _toc()

subheading("Цель и задачи работы:")
text("Целью работы является ... .")
text("Задачи: ... .")

page_break()
heading("1. Название первого раздела")
text("Краткое описание того, что делали и зачем.")
listing("Запрос для проверки данных", """
select count(*) from my_table;   -- 12 345
""")
picture("пример.png", "Результат выполнения запроса")
table("Название таблицы", [
    ["Показатель", "Значение"],
    ["Строк всего", "12 345"],
    ["Из них дублей", "120"],
])
text("Вывод по разделу одним-двумя предложениями.")

heading("2. Название второго раздела")
subheading("2.1. Подраздел")
text("Текст подраздела.")

page_break()
heading("Вывод")
text("В ходе работы ... .")

doc.save(OUT)
print("Готово:", OUT.resolve())
