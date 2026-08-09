import json
import pathlib
import re
from datetime import date

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

ROOT = pathlib.Path(__file__).parent
CONTENT = ROOT / "content"
OUT = ROOT / "מסע-הדורות-מסמך-עריכה-לבונה-האתר.docx"

# תיקונים בטוחים שנמצאו בחומרי האתר. ההחלפות מיועדות להפיק רשימת עבודה,
# ולא לשנות את קובצי האתר עצמם.
REPLACEMENTS = [
    ("כמה שנים נשרה הגלות", "כמה שנים נמשכה גלות בבל"),
    ("הממלכה השברה", "הממלכה השבורה"),
    ("להשמיר את הדת", "לשמור על חיי התורה"),
    ("וזנחת מצוות", "והזנחת המצוות"),
    ("מהלשינת הכותים", "מהלשנת הכותים"),
    ("ותן להם פטור", "ונתן להם פטור"),
    ("הבנייה השלימה", "הבנייה הושלמה"),
    ("המשבר שמצא בירושלים היה נמרץ", "המשבר שמצא בירושלים היה חמור"),
    ("ממלחמתם של אשור", "בעקבות כיבושי אשור"),
    ("עבדים ואמהות", "עבדים ושפחות"),
    ("וּקְרִיבוּ קרבנות", "והקריבו קרבנות"),
    ("וקריבו קרבנות", "והקריבו קרבנות"),
    ("עיר מבוצרת וממושלת", "עיר מבוצרת ובעלת הנהגה מסודרת"),
    ("היהודים זינחו", "היהודים זנחו"),
    ("קבוע הדינים", "קביעת הדינים"),
    ("תיקוני התפילה", "תיקון נוסח התפילה"),
    ("הם לא היהודים בעיני החוק", "אין דינם כיהודים על פי ההלכה"),
    ("שרשלת", "שרשרת"),
    ("איזו טנטטיבה של דיון", "איזו מסגרת ציבורית"),
    ("תפילת התפילה של בתי כנסת", "התפילה המרכזית שנאמרת בכל יום"),
    ("מזדיקים", "צדוקים"),
    ("מלך יוונן", "מלך יוון"),
    ("שגשר בין שתי עידנים", "שגישר בין שני עידנים"),
    ("להשלים את העולם", "ולהפיץ אותה ברחבי העולם"),
    ("בשל שבועת אמונים שנשמרו לפרסים", "בשל שבועת האמונים שלהם למלך פרס"),
    ("בעדיפות זו", "בעקבות זאת"),
    ("סיפור המלגדה", "המעשה המובא בגמרא"),
    ("הסנהדרין נדהם", "חכמי ישראל נדהמו"),
    ("והשתחוה", "והשתחווה"),
    ("עם מרות של המפגש", "בעקבות המפגש"),
    ("סמכויות שלטוניות מרחיבות", "סמכויות שלטוניות מורחבות"),
    ("בגיל 33 בלבד, חולה ומת", "בגיל 33 בלבד, חלה ומת"),
    ("מלחמות עסיסיות", "מלחמות עזות"),
    ("זה הקדמה ישירה", "זוהי הקדמה ישירה"),
    ("בן כמה גיל מת", "בן כמה היה במותו"),
    ("תוך הסביר", "והסביר"),
    ("התאחדות כולנו", "אחדות מלאה"),
    ("כשהוא מול התנגדות", "למרות התנגדות"),
    ("הגנה את השוק", "מנע את פתיחת השוק"),
    ("לצרוף", "לצירוף"),
    ("בפעיל", "בפועל"),
    ("בנייית", "בניית"),
    ("חופף מסים", "פטור ממסים"),
    ("תושבים מסובבים", "תושבי האזור"),
    ("אצל דריוס", "בחצר מלך פרס"),
    ("מיסים", "מסים"),
    ("היתה", "הייתה"),
    ("היה היתה", "היה"),
]

MANUAL = {
    "מאה שנות קיום שברוח": "כמאה שנים של שיבה, בנייה והתחדשות רוחנית",
    "הגלות לבבל: רקע להשיבה": "גלות בבל: הרקע לשיבת ציון",
    "המסע והשירה מרוב שמחה": "השיבה לציון בשמחה גדולה",
    "מהלכי עזרא: השמרה של הדת": "פעולותיו של עזרא לחיזוק חיי התורה",
    "מה היה התפקיד המוקדם של עזרא בבנייה?": "מה היה תפקידו המרכזי של עזרא בשיבת ציון?",
    "מי נתן את ההרשאה לשוב ולבנות את בית המקדש?": "מי התיר ליהודים לשוב לירושלים ולבנות את בית המקדש?",
    "מהו תפקידו של זרובבל?": "מה היה מעמדו של זרובבל?",
    "מה היא \"שרשרת הקבלה\"?": "מהי \"שרשרת הקבלה\"?",
    "מה היא \"שמונה עשרה\"?": "מהי תפילת שמונה עשרה?",
    "לאיזו ממלכה משנה הפכה סוריה אחרי חלוקת הממלכה?": "בידי איזו ממלכה הייתה סוריה לאחר חלוקת האימפריה של אלכסנדר?",
    "ארץ ישראל במוקד המאבק: קדמה למרד החשמונאים": "ארץ ישראל במוקד המאבק: הרקע למרד החשמונאים",
}

FACT_FLAGS = [
    "לפי המסורת", "התלמוד", "הגמרא", "המדרש", "לפנה", "שנת", "שנים",
    "אחרון הנביאים", "פרה האדומה", "בן נריה", "בנו של אסתר", "אחותו של",
    "206 שנים", "72", "שמעון הצדיק", "אלכסנדר", "חורבן"
]


def iter_strings(obj, path=""):
    if isinstance(obj, dict):
        for key, value in obj.items():
            yield from iter_strings(value, f"{path}/{key}")
    elif isinstance(obj, list):
        for idx, value in enumerate(obj):
            yield from iter_strings(value, f"{path}/{idx}")
    elif isinstance(obj, str):
        yield path, obj


def corrected(text):
    out = MANUAL.get(text, text)
    for old, new in REPLACEMENTS:
        out = out.replace(old, new)
    out = re.sub(r"\s+([,.:;!?])", r"\1", out)
    out = re.sub(r" {2,}", " ", out)
    return out


def rtl(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = paragraph._p.get_or_add_pPr()
    bidi = pPr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        pPr.append(bidi)


def set_run(run, size=11, bold=False, color="172033"):
    run.font.name = "Arial"
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), "Arial")
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), "Arial")
    run._element.get_or_add_rPr().rFonts.set(qn("w:cs"), "Arial")
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=100, start=120, bottom=100, end=120):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tcMar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tcMar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def add_heading(doc, text, level):
    p = doc.add_paragraph(style=f"Heading {level}")
    rtl(p)
    r = p.add_run(text)
    set_run(r, {1: 17, 2: 14, 3: 12}[level], True, {1: "0B2545", 2: "9A6A12", 3: "1F4D78"}[level])
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    rtl(p)
    set_run(p.add_run(text), 10.5)
    p.paragraph_format.space_after = Pt(4)


def add_correction(doc, location, old, new, note=None):
    table = doc.add_table(rows=3 + (1 if note else 0), cols=2)
    table.autofit = False
    table.columns[0].width = Inches(1.35)
    table.columns[1].width = Inches(5.15)
    rows = [("מיקום", location), ("הנוסח הקיים", old), ("להחליף ל", new)]
    if note:
        rows.append(("הערה", note))
    for row, (label, value) in zip(table.rows, rows):
        row.cells[0].width = Inches(1.35)
        row.cells[1].width = Inches(5.15)
        row.cells[0].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        row.cells[1].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        shade(row.cells[0], "E9D7A5")
        for cell in row.cells:
            set_cell_margins(cell)
            for p in cell.paragraphs:
                rtl(p)
        set_run(row.cells[0].paragraphs[0].add_run(label), 9.5, True, "0B2545")
        set_run(row.cells[1].paragraphs[0].add_run(value), 9.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def path_label(path, json_path):
    rel = path.relative_to(CONTENT).as_posix()
    return f"{rel}  ›  {json_path.strip('/')}"


def main():
    files = sorted(CONTENT.rglob("*.json"))
    corrections = []
    fact_review = []
    for path in files:
        data = json.loads(path.read_text(encoding="utf-8"))
        for jpath, original in iter_strings(data):
            revised = corrected(original)
            if revised != original:
                corrections.append((path, jpath, original, revised))
            if len(original) > 60 and any(flag in original for flag in FACT_FLAGS):
                fact_review.append((path, jpath, revised))

    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11)
    sec.top_margin = sec.bottom_margin = Inches(0.75)
    sec.left_margin = sec.right_margin = Inches(0.8)
    sec.header_distance = Inches(0.35)
    sec.footer_distance = Inches(0.35)

    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.2

    header = sec.header.paragraphs[0]
    rtl(header)
    set_run(header.add_run("מסע הדורות  |  מסמך עריכה לבונה האתר"), 9, False, "6B7280")
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(footer.add_run("גרסת עבודה ערוכה — אין לפרסם לפני בדיקת המקורות המסומנים"), 8, False, "6B7280")

    p = doc.add_paragraph()
    rtl(p)
    p.paragraph_format.space_before = Pt(54)
    p.paragraph_format.space_after = Pt(10)
    set_run(p.add_run("מסע הדורות"), 30, True, "0B2545")
    p2 = doc.add_paragraph()
    rtl(p2)
    set_run(p2.add_run("מסמך עריכה מקצועית והנחיות יישום לבונה האתר"), 17, True, "9A6A12")
    p3 = doc.add_paragraph()
    rtl(p3)
    set_run(p3.add_run("קו עריכה: תורני־מסורתי, בהיר, פשוט ונעים לקריאה"), 12, False, "374151")
    p4 = doc.add_paragraph()
    rtl(p4)
    p4.paragraph_format.space_before = Pt(20)
    set_run(p4.add_run(f"הוכן בתאריך {date.today().strftime('%d.%m.%Y')}  |  נסרקו {len(files)} קובצי תוכן"), 10, False, "6B7280")
    doc.add_page_break()

    add_heading(doc, "1. הוראות לבונה האתר", 1)
    add_bullet(doc, "יש לבצע את ההחלפות לפי נתיב הקובץ והמיקום המופיעים בכל סעיף.")
    add_bullet(doc, "כאשר אותו משפט מופיע גם במאגר הדמויות וגם בתוך פרק, יש לעדכן את שני המקומות.")
    add_bullet(doc, "יש לשמור על כתיבה מימין לשמאל ועל גרשיים עבריים אחידים ככל האפשר.")
    add_bullet(doc, "אין לפרסם טענה המסומנת 'בדיקת מקור' לפני אימות מול מקור תורני מוסמך.")
    add_bullet(doc, "האתר החי אינו מסונכרן במלואו עם הקבצים המקומיים; יש לפרוס מחדש את הקבצים המעודכנים לאחר ההטמעה.")

    add_heading(doc, "2. כללי העריכה המחייבים", 1)
    rules = [
        "להעדיף משפטים קצרים: רעיון מרכזי אחד בכל משפט.",
        "להשתמש בלשון תורנית טבעית: ה', הקב\"ה, חכמי ישראל, תורה שבעל פה.",
        "כאשר המקור הוא חז\"ל, לכתוב: 'חז\"ל מספרים', 'לפי המסורת' או 'בגמרא מובא'.",
        "לא להציג מדרש או מסורת כקביעה מחקרית מוסכמת בלי לציין את מקורם.",
        "להימנע ממונחים מודרניים מדי כגון 'אפקט דומינו', 'טרויאני', 'אידיאולוגי' ו'תשתיות לאומיות', כשאפשר לומר זאת בפשטות.",
        "לכתוב 'קורבנות' או 'קרבנות' באופן אחיד בכל האתר. במסמך זה מומלץ: 'קורבנות'.",
        "לכתוב 'מסים', 'נישואין', 'הייתה', 'בית המקדש השני' באופן אחיד.",
        "בשאלונים: שאלה קצרה, ארבע תשובות באותו מבנה דקדוקי, והסבר של משפט אחד לאחר הבחירה.",
    ]
    for item in rules:
        add_bullet(doc, item)

    add_heading(doc, "3. תיקוני נוסח מדויקים", 1)
    intro = doc.add_paragraph()
    rtl(intro)
    set_run(intro.add_run(f"להלן {len(corrections)} תיקונים חד־משמעיים שנמצאו בקובצי התוכן. כל נוסח חלופי מוכן להעתקה."), 11)

    current_file = None
    for path, jpath, old, new in corrections:
        rel = path.relative_to(CONTENT).as_posix()
        if rel != current_file:
            add_heading(doc, rel, 2)
            current_file = rel
        add_correction(doc, path_label(path, jpath), old, new)

    add_heading(doc, "4. נקודות המחייבות בדיקת מקור תורני", 1)
    p = doc.add_paragraph()
    rtl(p)
    set_run(p.add_run("הסעיפים הבאים אינם בהכרח שגויים. הם מסומנים מפני שהם כוללים תאריך, ייחוס, מספר או מסורת שראוי לאמת לפני פרסום לציבור."), 10.5)
    seen = set()
    for path, jpath, text in fact_review:
        key = (path.as_posix(), text)
        if key in seen:
            continue
        seen.add(key)
        add_correction(doc, path_label(path, jpath), text, text, "בדיקת מקור: לאמת את הייחוס, התאריך או הניסוח מול מקור תורני מוסמך.")

    add_heading(doc, "5. תיקוני ממשק וחוויית משתמש", 1)
    for item in [
        "להוסיף חיפוש לפי תקופה, פרק ודמות.",
        "להציג התקדמות: כמה פרקים הושלמו ומהו הפרק הבא.",
        "להקטין מעט את תמונת הפתיחה בעמוד הפרק כדי שהתוכן יופיע מוקדם יותר.",
        "להציג בחידון שאלה אחת בכל פעם, עם הסבר קצר לאחר כל תשובה.",
        "להחליף את הכיתוב החוזר 'מוכן ללימוד' בסימון חזותי קטן ועקבי.",
        "להוסיף כללי עיצוב ייעודיים למסכים צרים ולבדוק בפועל בטלפון.",
        "להחליף אימוג'ים לא עקביים במערכת אייקונים אחידה.",
        "להוסיף בכל עמוד פרק כפתור ברור: הקודם, מפת המסע, הבא.",
    ]:
        add_bullet(doc, item)

    add_heading(doc, "6. בדיקות קבלה לאחר ההטמעה", 1)
    for item in [
        "אין שגיאות כתיב גלויות בכותרות, בפסקאות ובחידונים.",
        "כל קישור לדמות נפתח בדמות הנכונה.",
        "כל תשובת חידון מסומנת נכון וההסבר מתאים לה.",
        "האתר נבדק ברוחב מחשב וברוחב טלפון, ללא גלילה אופקית.",
        "הטקסט באתר החי זהה לגרסה המעודכנת בקובצי התוכן.",
        "כל טענה שסומנה לבדיקת מקור אושרה או נוסחה מחדש כ'לפי המסורת'.",
    ]:
        add_bullet(doc, item)

    doc.core_properties.title = "מסע הדורות — מסמך עריכה לבונה האתר"
    doc.core_properties.subject = "עריכת תוכן תורנית־מסורתית והנחיות יישום"
    doc.core_properties.author = ""
    doc.save(OUT)
    print(OUT)
    print(f"corrections={len(corrections)} fact_flags={len(seen)} files={len(files)}")


if __name__ == "__main__":
    main()
