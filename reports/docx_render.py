# -*- coding: utf-8 -*-
"""توليد نسخة Word (.docx) من تقارير الفحص الآلي بنفس هوية نسخة PDF.

- اتجاه RTL كامل (فقرات، جداول، تذييل).
- الرسوم (العداد، الأشرطة، الكلمات المفتاحية، الخط الزمني) صور PNG بلا نص عربي.
- المقاطع اللاتينية والأرقام محاطة بـ LRE/PDF لتثبيت اتجاهها داخل النص العربي.
"""
import io
import os
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont

from clients import AXES, CLIENTS, QUOTED_FINDINGS
from render import DATE_AR, LOCK_TEXT, NOTE_TEXT, WA, color_for

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "output")
FONT = "Arial"           # متوفر في Windows/Mac ويدعم العربية
MONO = "Consolas"
NAVY, GOLD, GREY, INK = "0F2A4A", "C9A227", "7B8794", "1F2933"
SEV = {"crit": ("حرج", "C0392B", "FDECEA"), "high": ("مرتفع", "D35400", "FEF3E8"), "mid": ("متوسط", "B7950B", "FDF8E4")}
STATUS = {"ok": ("✓", "1E8449"), "warn": ("!", "D35400"), "bad": ("✕", "C0392B")}
NOTO = "/usr/share/fonts/truetype/noto/NotoSans-{}.ttf"


def ltr(s):
    """تثبيت اتجاه مقطع لاتيني/رقمي داخل فقرة عربية."""
    return "\u202a" + s + "\u202c"


# ---------- أدوات XML ----------
def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), str(v))
    return e


def para_rtl(p, space_after=2, space_before=0, center=False, keep_next=False):
    pPr = p._p.get_or_add_pPr()
    pPr.insert(0, _el("w:bidi"))
    pf = p.paragraph_format
    pf.space_after, pf.space_before = Pt(space_after), Pt(space_before)
    pf.line_spacing = 1.12
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if keep_next:
        pf.keep_with_next = True
    return p


def run(p, text, size=9.5, bold=False, color=INK, font=FONT, shade=None, rtl=True):
    r = p.add_run(text)
    rPr = r._r.get_or_add_rPr()
    rPr.append(_el("w:rFonts", **{"w:ascii": font, "w:hAnsi": font, "w:cs": font, "w:eastAsia": font}))
    if bold:
        rPr.append(_el("w:b"))
        rPr.append(_el("w:bCs"))
    rPr.append(_el("w:color", **{"w:val": color}))
    rPr.append(_el("w:sz", **{"w:val": int(size * 2)}))
    rPr.append(_el("w:szCs", **{"w:val": int(size * 2)}))
    if shade:
        rPr.append(_el("w:shd", **{"w:val": "clear", "w:color": "auto", "w:fill": shade}))
    if rtl:
        rPr.append(_el("w:rtl"))
    return r


def shade(cell, fill):
    cell._tc.get_or_add_tcPr().append(_el("w:shd", **{"w:val": "clear", "w:color": "auto", "w:fill": fill}))


def cell_width(cell, cm):
    cell.width = Cm(cm)


def valign(cell, v="center"):
    cell._tc.get_or_add_tcPr().append(_el("w:vAlign", **{"w:val": v}))


def table(doc, rows, cols, widths, borders=None, pad=(60, 100)):
    t = doc.add_table(rows=rows, cols=cols)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblPr = t._tbl.tblPr
    tblPr.append(_el("w:bidiVisual"))
    tblPr.append(_el("w:tblLayout", **{"w:type": "fixed"}))
    b = _el("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        spec = (borders or {}).get(side)
        if spec:
            b.append(_el(f"w:{side}", **{"w:val": "single", "w:sz": spec[0], "w:color": spec[1], "w:space": 0}))
        else:
            b.append(_el(f"w:{side}", **{"w:val": "nil"}))
    tblPr.append(b)
    m = _el("w:tblCellMar")
    for side, v in (("top", pad[0]), ("bottom", pad[0]), ("left", pad[1]), ("right", pad[1])):
        m.append(_el(f"w:{side}", **{"w:w": v, "w:type": "dxa"}))
    tblPr.append(m)
    set_grid(t, widths)
    return t


def set_grid(t, widths):
    """عرض الأعمدة في tblGrid + tcW + عرض الجدول (LibreOffice وWord يقرآن tblGrid)."""
    tw = int(sum(widths) * 567)
    tblPr = t._tbl.tblPr
    for old in tblPr.findall(qn("w:tblW")):
        tblPr.remove(old)
    tblPr.append(_el("w:tblW", **{"w:w": tw, "w:type": "dxa"}))
    grid = t._tbl.tblGrid
    for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths[i] * 567)))
    for row in t.rows:
        for i, c in enumerate(row.cells):
            cell_width(c, widths[i])


def cell_para(cell, first=True, **kw):
    p = cell.paragraphs[0] if first else cell.add_paragraph()
    return para_rtl(p, **kw)


def no_split(row):
    row._tr.get_or_add_trPr().append(_el("w:cantSplit"))


def field(p, code, size=8, color=GREY):
    for kind, txt in (("begin", None), (None, code), ("end", None)):
        r = run(p, "", size=size, color=color, rtl=False)
        if kind:
            r._r.append(_el("w:fldChar", **{"w:fldCharType": kind}))
        else:
            it = _el("w:instrText", **{"xml:space": "preserve"})
            it.text = txt
            r._r.append(it)


def picture(p, png_bytes, width_cm):
    p.add_run().add_picture(io.BytesIO(png_bytes), width=Cm(width_cm))


# ---------- الرسوم (PNG) ----------
def _hex(c):
    c = c.lstrip("#")
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def _png(img):
    b = io.BytesIO()
    img.save(b, "PNG", dpi=(300, 300))
    return b.getvalue()


def gauge_png(score):
    import math
    S = 4
    W = 200 * S
    img = Image.new("RGBA", (W, W), (255, 255, 255, 0))
    d = ImageDraw.Draw(img)
    box = [22 * S, 22 * S, 178 * S, 178 * S]
    d.ellipse(box, outline=_hex("e9edf3"), width=18 * S)
    col = _hex(color_for(score))
    d.arc(box, -90, -90 + 360 * score / 100, fill=col, width=18 * S)
    r = 9 * S  # أطراف مستديرة
    for ang in (-90, -90 + 360 * score / 100):
        a = math.radians(ang)
        cx, cy = 100 * S + 69 * S * math.cos(a), 100 * S + 69 * S * math.sin(a)
        d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=col)
    f1 = ImageFont.truetype(NOTO.format("Bold"), 56 * S)
    f2 = ImageFont.truetype(NOTO.format("Regular"), 17 * S)
    d.text((100 * S, 96 * S), str(score), font=f1, fill=col, anchor="mm")
    d.text((100 * S, 136 * S), "/ 100", font=f2, fill=_hex("7b8794"), anchor="mm")
    return _png(img.resize((400, 400), Image.LANCZOS))


def bar_png(pct):
    W, H = 1400, 44
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([0, 6, W - 1, H - 7], radius=16, fill=_hex("eef1f5"))
    w = max(int(W * pct / 100), 34)
    d.rounded_rectangle([W - w, 6, W - 1, H - 7], radius=16, fill=_hex(color_for(pct)))  # يملأ من اليمين
    return _png(img)


def dots_png(have, need, slots=8):
    s, g = 40, 10
    W = slots * (s + g)
    img = Image.new("RGB", (W, s + 4), "white")
    d = ImageDraw.Draw(img)
    for i in range(need):
        x = W - (i + 1) * (s + g) + g  # من اليمين لليسار
        d.rounded_rectangle([x, 2, x + s, s + 2], radius=9, fill=_hex("1e8449" if i < have else "f5c6c0"))
    return _png(img)


def timeline_png():
    S = 3
    y0, y1, W = 2003.0, 2026.9, 640
    pad = 24
    img = Image.new("RGB", ((W + 2 * pad) * S, 80 * S), "white")
    d = ImageDraw.Draw(img)
    x = lambda y: (pad + (y - y0) / (y1 - y0) * W) * S
    d.rounded_rectangle([x(y0), 22 * S, x(y1), 48 * S], radius=5 * S, fill=_hex("eef1f5"))
    segs = [(2003 + 7/12, 2004 + 2/12, "1f4e79", ""), (2004 + 2/12, 2004 + 10/12, "e74c3c", ""),
            (2004 + 10/12, 2006 + 6/12, "1f4e79", ""), (2006 + 8/12, 2009 + 9/12, "8e9aaf", ""),
            (2009 + 9/12, 2021 + 3/12, "e74c3c", "11.5 y"), (2021 + 3/12, 2022 + 3/12, "1f4e79", ""),
            (2022 + 3/12, 2026 + 10/12, "e74c3c", "4.6 y")]
    fb = ImageFont.truetype(NOTO.format("Bold"), 12 * S)
    fr = ImageFont.truetype(NOTO.format("Regular"), 11 * S)
    for a, b, col, lab in segs:
        d.rectangle([x(a), 22 * S, x(b), 48 * S], fill=_hex(col))
        if lab:
            d.text(((x(a) + x(b)) / 2, 35 * S), lab, font=fb, fill="white", anchor="mm")
    for yr in (2003, 2006, 2009, 2012, 2015, 2018, 2021, 2024, 2026):
        d.line([x(yr), 50 * S, x(yr), 56 * S], fill=_hex("9aa5b1"), width=S)
        d.text((x(yr), 68 * S), str(yr), font=fr, fill=_hex("52606d"), anchor="mm")
    return _png(img)


def lock_png():
    S = 8
    img = Image.new("RGBA", (22 * S, 26 * S), (255, 255, 255, 0))
    d = ImageDraw.Draw(img)
    d.arc([6 * S, 2 * S, 16 * S, 16 * S], 180, 360, fill=_hex(NAVY), width=3 * S)
    d.line([6 * S + S, 9 * S, 6 * S + S, 12 * S], fill=_hex(NAVY), width=3 * S)
    d.line([16 * S - S, 9 * S, 16 * S - S, 12 * S], fill=_hex(NAVY), width=3 * S)
    d.rounded_rectangle([2 * S, 11 * S, 20 * S, 25 * S], radius=3 * S, fill=_hex(NAVY))
    d.ellipse([9 * S, 16 * S, 13 * S, 20 * S], fill=_hex(GOLD))
    return _png(img)


# ---------- بناء المستند ----------
def heading(doc, text, space_before=10):
    """عنوان قسم: علامة ذهبية + نص كحلي (فقرة عادية، بدون جدول)."""
    p = para_rtl(doc.add_paragraph(), space_after=4, space_before=space_before, keep_next=True)
    run(p, "▍", size=13, bold=True, color=GOLD, rtl=False)
    run(p, " " + text, size=12, bold=True, color=NAVY)
    return p


def _spacer(pt):
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pPr.append(_el("w:bidi"))
    pPr.append(_el("w:spacing", **{"w:before": 0, "w:after": 0, "w:line": int(pt * 20), "w:lineRule": "exact"}))
    p.append(pPr)
    return p


def gap(doc, pt=4):
    p = doc.add_paragraph()
    para_rtl(p, space_after=0)
    p.paragraph_format.line_spacing = Pt(pt)
    return p


def build_docx(c):
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.top_margin, sec.bottom_margin = Cm(1.4), Cm(1.6)
    sec.left_margin = sec.right_margin = Cm(1.5)
    sec.header_distance, sec.footer_distance = Cm(0.6), Cm(0.7)
    sec._sectPr.append(_el("w:bidi"))

    # الافتراضيات: لغة عربية + خط Arial
    st = doc.styles["Normal"]
    rPr = st.element.get_or_add_rPr()
    rPr.append(_el("w:rFonts", **{"w:ascii": FONT, "w:hAnsi": FONT, "w:cs": FONT, "w:eastAsia": FONT}))
    rPr.append(_el("w:lang", **{"w:val": "en-US", "w:bidi": "ar-SA"}))
    st.font.size = Pt(9.5)

    cp = doc.core_properties
    cp.author, cp.title = "Al-Abbasi Soft", f"تقرير فحص ATS - {c['name_en']}"
    cp.subject, cp.keywords = "ATS Report", "ATS, CV, مؤسسة العباسي سوفت"

    # الترويسة والتذييل
    hp = para_rtl(sec.header.paragraphs[0], space_after=0)
    run(hp, "سري - خاص بصاحب الملف", size=7.5, color="B0B8C4")
    ft = sec.footer
    ftab = ft.add_table(1, 2, Cm(18))
    ftab._tbl.tblPr.append(_el("w:bidiVisual"))
    set_grid(ftab, [12.0, 6.0])
    fp = para_rtl(ftab.cell(0, 0).paragraphs[0], space_after=0)
    run(fp, "مؤسسة العباسي سوفت  |  بإشراف أ. سامية الراشد", size=8, color=GREY)
    fp2 = para_rtl(ftab.cell(0, 1).paragraphs[0], space_after=0)
    fp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(fp2, "صفحة ", size=8, color=GREY)
    field(fp2, "PAGE")
    run(fp2, " من ", size=8, color=GREY)
    field(fp2, "NUMPAGES")
    ft.paragraphs[0]._p.getparent().remove(ft.paragraphs[0]._p)

    # ===== رأس التقرير =====
    t = table(doc, 1, 2, [12.6, 5.4], pad=(160, 220))
    for cell in t.rows[0].cells:
        shade(cell, NAVY)
        valign(cell)
    p = cell_para(t.cell(0, 0), space_after=2)
    run(p, " تقرير فحص آلي ", size=8, bold=True, color=NAVY, shade=GOLD)
    p = cell_para(t.cell(0, 0), first=False, space_after=0)
    run(p, "تقرير توافق السيرة الذاتية", size=16, bold=True, color="FFFFFF")
    p = cell_para(t.cell(0, 0), first=False, space_after=0)
    run(p, "مع أنظمة التوظيف ", size=16, bold=True, color="FFFFFF")
    run(p, "ATS", size=16, bold=True, color="FFFFFF", rtl=False)
    p = cell_para(t.cell(0, 0), first=False, space_after=0)
    run(p, "نظام الفحص الآلي للسير الذاتية | مؤسسة العباسي سوفت", size=9, color=GOLD)
    p = cell_para(t.cell(0, 1), space_after=2)
    run(p, "رقم التقرير: ", size=8.5, color="CBD5E1")
    run(p, ltr(c["report_no"]), size=8.5, color="CBD5E1", rtl=False)
    p = cell_para(t.cell(0, 1), first=False, space_after=0)
    run(p, "تاريخ الفحص: " + DATE_AR, size=8.5, color="CBD5E1")

    gap(doc, 6)
    # ===== بطاقات المعلومات =====
    t = table(doc, 1, 5, [5.8, 0.3, 6.0, 0.3, 5.6], pad=(80, 140))
    cards = [(0, "صاحب الملف", [(c["name_ar"], True)]),
             (2, "الملف المفحوص", ([(c["file_display"], True)] if c.get("file_display") else [(ltr(c["file"]), False)]) + [(" " + c["pages_ar"], True)]),
             (4, "المجال المستهدف", [(c["field"], True)])]
    for idx, label, parts in cards:
        cell = t.cell(0, idx)
        shade(cell, "F4F6F9")
        p = cell_para(cell, space_after=0)
        run(p, label, size=7.5, color=GREY)
        p = cell_para(cell, first=False, space_after=0)
        for txt, is_ar in parts:
            run(p, txt, size=8.8, color=INK, rtl=is_ar)

    gap(doc, 6)
    # ===== صندوق النتيجة =====
    crit = sum(1 for f in c["findings"] if f[0] == "crit")
    sc = color_for(c["score"]).lstrip("#").upper()
    t = table(doc, 1, 2, [4.2, 13.8], borders={k: (8, "E4E7EB") for k in ("top", "left", "bottom", "right")}, pad=(100, 200))
    for cell in t.rows[0].cells:
        valign(cell)
    picture(cell_para(t.cell(0, 0), center=True, space_after=0), gauge_png(c["score"]), 3.6)
    p = cell_para(t.cell(0, 1), space_after=3)
    run(p, "النتيجة: " + c["verdict"], size=13, bold=True, color=sc)
    p = cell_para(t.cell(0, 1), first=False, space_after=3)
    run(p, c["verdict_note"], size=9.5, color="52606D")
    p = cell_para(t.cell(0, 1), first=False, space_after=0)
    run(p, f"{len(c['findings'])} ملاحظة", size=9.5, bold=True)
    run(p, " منها ", size=9.5, color="52606D")
    run(p, f"{crit} حرجة", size=9.5, bold=True, color="C0392B")
    run(p, " تؤثر مباشرة على قرار الفرز.", size=9.5, color="52606D")

    # ===== المحاور =====
    heading(doc, "تفصيل النتيجة حسب محاور الفرز")
    t = table(doc, len(AXES), 3, [4.6, 11.6, 1.8], pad=(25, 80))
    for i, ((label, w), v) in enumerate(zip(AXES, c["axes"])):
        row = t.rows[i]
        for cell in row.cells:
            valign(cell)
        run(cell_para(row.cells[0], space_after=0), label, size=9)
        picture(cell_para(row.cells[1], space_after=0), bar_png(v / w * 100), 11.2)
        run(cell_para(row.cells[2], space_after=0, center=True), ltr(f"{v} / {w}"), size=8.5, color="52606D", rtl=False)

    # ===== قراءة النظام =====
    heading(doc, "كيف قرأ نظام الفرز ملفك؟")
    t = table(doc, len(c["parser"]), 3, [0.8, 4.0, 13.2], borders={"insideH": (4, "EEF1F5"), "bottom": (4, "EEF1F5")}, pad=(30, 80))
    for i, (k, s, v) in enumerate(c["parser"]):
        row = t.rows[i]
        sym, col = STATUS[s]
        run(cell_para(row.cells[0], center=True, space_after=0), sym, size=10, bold=True, color=col, rtl=False)
        run(cell_para(row.cells[1], space_after=0), k, size=9, bold=True)
        run(cell_para(row.cells[2], space_after=0), v, size=9)

    # ===== الملاحظات (صفحة جديدة) =====
    br = para_rtl(doc.add_paragraph(), space_after=0)
    br.add_run().add_break(WD_BREAK.PAGE)
    heading(doc, "الملاحظات المكتشفة داخل الملف", space_before=0)
    p = para_rtl(doc.add_paragraph(), space_after=5)
    run(p, f"النصوص داخل الإطارات في أول {QUOTED_FINDINGS} ملاحظات منقولة حرفياً من ملفك كما قرأها النظام. "
           "باقي الملاحظات مرصودة بمواضعها الدقيقة في الفحص البرمجي.", size=8.5, color="52606D")
    lock = lock_png()
    for i, (sev, title, quotes, desc, impact) in enumerate(c["findings"], 1):
        lab, col, bg = SEV[sev]
        t = table(doc, 1, 2, [0.22, 17.78], pad=(70, 140))
        no_split(t.rows[0])
        shade(t.cell(0, 0), col)
        cell = t.cell(0, 1)
        shade(cell, bg)
        p = cell_para(cell, space_after=2, keep_next=True)
        run(p, ltr(f"#{i:02d}") + "  ", size=8, color="9AA5B1", rtl=False)
        run(p, title, size=10.5, bold=True)
        run(p, "   ")
        run(p, f" {lab} ", size=7.5, bold=True, color="FFFFFF", shade=col)
        p = cell_para(cell, first=False, space_after=2, keep_next=True)
        if i <= QUOTED_FINDINGS and quotes:
            for j, q in enumerate(quotes):
                if j:
                    run(p, "   ")
                run(p, ltr(" " + q + " "), size=8, color="3E4C59", font=MONO, shade="FFFFFF", rtl=False)
        else:
            p.add_run().add_picture(io.BytesIO(lock), height=Pt(9))
            run(p, "  ", size=8)
            run(p, LOCK_TEXT, size=8, bold=True, color=NAVY, shade="FFFFFF")
        p = cell_para(cell, first=False, space_after=1, keep_next=True)
        run(p, desc, size=9.3)
        p = cell_para(cell, first=False, space_after=0)
        run(p, "الأثر: ", size=8.6, bold=True)
        run(p, impact, size=8.6, color="52606D")
        gap(doc, 5)

    # ===== الكلمات المفتاحية =====
    heading(doc, "تغطية الكلمات المفتاحية في مجالك", space_before=6)
    p = para_rtl(doc.add_paragraph(), space_after=3, keep_next=True)
    run(p, "عدد المصطلحات الأساسية التي تبحث عنها أنظمة الفرز في كل فئة، مقابل الموجود فعلياً في ملفك.", size=8.5, color="52606D")
    t = table(doc, len(c["keywords"]), 3, [6.0, 4.6, 1.8], pad=(20, 80))
    t.alignment = WD_TABLE_ALIGNMENT.RIGHT
    for i, (label, have, need) in enumerate(c["keywords"]):
        row = t.rows[i]
        for cell in row.cells:
            valign(cell)
        no_split(row)
        kn = i < len(c["keywords"]) - 1
        run(cell_para(row.cells[0], space_after=0, keep_next=kn), label, size=9)
        picture(cell_para(row.cells[1], space_after=0, keep_next=kn), dots_png(have, need), 4.2)
        run(cell_para(row.cells[2], space_after=0, center=True), ltr(f"{have} / {need}"), size=8.5, color="52606D", rtl=False)
    p = para_rtl(doc.add_paragraph(), space_after=0, space_before=3)
    run(p, "■ ", size=9, color="1E8449", rtl=False)
    run(p, "موجود في ملفك      ", size=8)
    run(p, "■ ", size=9, color="F5C6C0", rtl=False)
    run(p, "مفقود", size=8)

    if c["timeline"]:
        heading(doc, "الخط الزمني كما يراه النظام")
        picture(para_rtl(doc.add_paragraph(), center=True, space_after=2), timeline_png(), 17.5)
        p = para_rtl(doc.add_paragraph(), space_after=0)
        for col, lab in (("1F4E79", "عمل"), ("8E9AAF", "دراسة معروضة كعمل"), ("E74C3C", "فجوة غير مبررة")):
            run(p, "■ ", size=9, color=col, rtl=False)
            run(p, lab + "      ", size=8)

    # ===== نقاط القوة =====
    heading(doc, "نقاط قوة حقيقية لا يُظهرها ملفك الحالي")
    t = table(doc, 2, 2, [9.0, 9.0], pad=(50, 160))
    for i, s in enumerate(c["strengths"]):
        cell = t.cell(i // 2, i % 2)
        shade(cell, "EAF6EE")
        run(cell_para(cell, space_after=0), "•  " + s, size=9.2)

    # ===== الخلاصة =====
    heading(doc, "الخلاصة")
    p = para_rtl(doc.add_paragraph(), space_after=4)
    run(p, c["summary"], size=9.5)
    p = para_rtl(doc.add_paragraph(), space_after=6)
    run(p, "التوصية: ", size=9.5, bold=True)
    run(p, "الملاحظات مترابطة في بنية الملف نفسها، فتعديلها جزئياً لا يرفع النتيجة بشكل كافٍ. الملف يحتاج ", size=9.5)
    run(p, "إعادة بناء كاملة", size=9.5, bold=True)
    run(p, " على قالب متوافق، بصياغة موجهة لمجالك وكلمات مفتاحية مطابقة لمتطلبات السوق السعودي.", size=9.5)

    # ===== صندوق العرض =====
    t = table(doc, 1, 2, [12.8, 5.2], pad=(140, 240))
    no_split(t.rows[0])
    for cell in t.rows[0].cells:
        shade(cell, NAVY)
        valign(cell)
    cell = t.cell(0, 0)
    p = cell_para(cell, space_after=4, keep_next=True)
    run(p, "إعادة بناء سيرتك الذاتية بالكامل", size=13, bold=True, color=GOLD)
    items = ["سيرة ذاتية عربي + إنجليزي", "متوافقة مع أنظمة ATS", "مكتوبة لوظيفتك من خبرتك الحقيقية",
             "خطاب تقديم هدية", "ملفات Word + PDF وتعديلان مجاناً", "جاهزة خلال 24 إلى 48 ساعة"]
    for it in items:
        p = cell_para(cell, first=False, space_after=0, keep_next=True)
        run(p, "•  " + it, size=9, color="FFFFFF")
    p = cell_para(cell, first=False, space_after=0, space_before=6)
    run(p, "للطلب أو الاستفسار عبر واتساب: ", size=9, color="CBD5E1")
    run(p, ltr("wa.me/" + WA), size=9, bold=True, color="FFFFFF", rtl=False)
    p = cell_para(cell, first=False, space_after=0, space_before=6)
    run(p, NOTE_TEXT, size=7, color="8A9BB3")
    tag = t.cell(0, 1)
    inner = tag.add_table(1, 1)
    inner._tbl.tblPr.append(_el("w:bidiVisual"))
    inner.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_grid(inner, [4.2])
    shade(inner.cell(0, 0), GOLD)
    p = para_rtl(inner.cell(0, 0).paragraphs[0], center=True, space_after=0, space_before=4)
    run(p, "العرض الحالي", size=10, bold=True, color=NAVY)
    p = para_rtl(inner.cell(0, 0).add_paragraph(), center=True, space_after=4)
    run(p, "49 ريال", size=22, bold=True, color=NAVY)
    para_rtl(tag.paragraphs[0], space_after=0)

    os.makedirs(OUT, exist_ok=True)
    out = os.path.join(OUT, c["out"] + ".docx")
    doc.save(out)
    return out


if __name__ == "__main__":
    for key in (sys.argv[1:] or CLIENTS):
        print(build_docx(CLIENTS[key]))
