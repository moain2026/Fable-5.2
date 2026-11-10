# -*- coding: utf-8 -*-
"""توليد تقارير الفحص الآلي PDF - مؤسسة العباسي سوفت."""
import html
import os
import re
import sys

import weasyprint

from clients import AXES, CLIENTS, QUOTED_FINDINGS

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "output")
DATE_AR = "9 نوفمبر 2026"
LOCK_TEXT = "تفاصيل ومواضع الخطأ مرصودة في الفحص البرمجي — تُعالج جذرياً عند إعادة البناء"
NOTE_TEXT = "تم إصدار هذا التقرير بناءً على الفحص الآلي لمحتوى الملف المرسل وبنيته الهندسية وفق معايير الفرز المعتمدة لدى المؤسسة."
LOCK_SVG = ('<svg class="lk" width="11" height="13" viewBox="0 0 22 26"><rect x="2" y="11" width="18" height="14" rx="3" fill="#0f2a4a"/>'
            '<path d="M6 11V7a5 5 0 0 1 10 0v4" fill="none" stroke="#0f2a4a" stroke-width="3"/><circle cx="11" cy="18" r="2.2" fill="#c9a227"/></svg>')
WA = "966509563048"

SEV = {
    "crit": ("حرج", "#c0392b", "#fdecea"),
    "high": ("مرتفع", "#d35400", "#fef3e8"),
    "mid": ("متوسط", "#b7950b", "#fdf8e4"),
}
STATUS = {"ok": ("✓", "#1e8449"), "warn": ("!", "#d35400"), "bad": ("✕", "#c0392b")}

LATIN = re.compile(r"([A-Za-z0-9\[\]\(\)\"'.,:/&+\-—]+(?:[ ]+[A-Za-z0-9\[\]\(\)\"'.,:/&+\-—]+)*)")


def t(s):
    """هروب + عزل المقاطع اللاتينية لاتجاه صحيح داخل النص العربي."""
    s = html.escape(s, quote=False)
    return LATIN.sub(lambda m: f'<span class="en">{m.group(1)}</span>'
                     if re.search(r"[A-Za-z]", m.group(1)) else m.group(1), s)


def fn(s):
    """علامات LRM حول المقاطع العربية داخل نص لاتيني ليبقى الترتيب صحيحاً."""
    return re.sub(r'([\u0600-\u06FF]+)', '\u200e\\1\u200e', s)


def file_html(c):
    if c.get("file_display"):
        return html.escape(c["file_display"])
    return f'<span class="en">{html.escape(fn(c["file"]))}</span>'


def color_for(score):
    return "#c0392b" if score < 50 else ("#d35400" if score < 75 else "#1e8449")


def gauge(score):
    c = color_for(score)
    circ = 2 * 3.14159 * 70
    dash = circ * score / 100
    return f'''<svg width="150" height="150" viewBox="0 0 190 190">
<circle cx="95" cy="95" r="70" fill="none" stroke="#e9edf3" stroke-width="16"/>
<circle cx="95" cy="95" r="70" fill="none" stroke="{c}" stroke-width="16" stroke-linecap="round"
 stroke-dasharray="{dash:.1f} {circ:.1f}" transform="rotate(-90 95 95)"/>
<text x="95" y="102" text-anchor="middle" font-family="Noto Sans" font-weight="700" font-size="46" fill="{c}">{score}</text>
<text x="95" y="130" text-anchor="middle" font-family="Noto Sans" font-size="15" fill="#7b8794">/ 100</text>
</svg>'''


def timeline_svg():
    """الخط الزمني لرندة: عمل / دراسة / فجوة (من 2003 إلى نوفمبر 2026)."""
    y0, y1, W = 2003.0, 2026.9, 640
    x = lambda y: (y - y0) / (y1 - y0) * W  # يسار = الأقدم (محور زمني لاتيني)
    segs = [
        (2003 + 7/12, 2004 + 2/12, "#1f4e79", ""),
        (2004 + 2/12, 2004 + 10/12, "#e74c3c", ""),
        (2004 + 10/12, 2006 + 6/12, "#1f4e79", ""),
        (2006 + 8/12, 2009 + 9/12, "#8e9aaf", ""),
        (2009 + 9/12, 2021 + 3/12, "#e74c3c", "11.5"),
        (2021 + 3/12, 2022 + 3/12, "#1f4e79", ""),
        (2022 + 3/12, 2026 + 10/12, "#e74c3c", "4.6"),
    ]
    out = [f'<svg width="{W+44}" height="80" viewBox="-22 0 {W+44} 80">']
    out.append(f'<rect x="0" y="22" width="{W}" height="26" rx="5" fill="#eef1f5"/>')
    for a, b, col, lab in segs:
        out.append(f'<rect x="{x(a):.1f}" y="22" width="{x(b)-x(a):.1f}" height="26" fill="{col}"/>')
        if lab:
            out.append(f'<text x="{(x(a)+x(b))/2:.1f}" y="40" text-anchor="middle" font-family="Noto Sans" '
                       f'font-weight="700" font-size="12" fill="#fff">{lab} y</text>')
    for yr in (2003, 2006, 2009, 2012, 2015, 2018, 2021, 2024, 2026):
        out.append(f'<line x1="{x(yr):.1f}" y1="50" x2="{x(yr):.1f}" y2="56" stroke="#9aa5b1"/>'
                   f'<text x="{x(yr):.1f}" y="70" text-anchor="middle" font-family="Noto Sans" font-size="11" fill="#52606d">{yr}</text>')
    out.append('</svg>')
    return "".join(out)


CSS = """
@page { size: A4; margin: 13mm 14mm 17mm 14mm;
  @bottom-right { content: "مؤسسة العباسي سوفت  |  بإشراف أ. سامية الراشد"; font-family: 'Noto Sans Arabic'; font-size: 8pt; color: #7b8794; }
  @bottom-left { content: "صفحة " counter(page) " من " counter(pages); font-family: 'Noto Sans Arabic'; font-size: 8pt; color: #7b8794; }
  @top-left { content: "سري - خاص بصاحب الملف"; font-family: 'Noto Sans Arabic'; font-size: 7.5pt; color: #b0b8c4; }
}
* { box-sizing: border-box; }
body { font-family: 'Noto Sans Arabic', 'Noto Sans'; color: #1f2933; font-size: 9.2pt; line-height: 1.55; margin: 0; }
.en { direction: ltr; unicode-bidi: embed; font-family: 'Noto Sans'; }
h1, h2, h3 { font-family: 'Noto Kufi Arabic'; margin: 0; }
.head { background: #0f2a4a; color: #fff; border-radius: 10px; padding: 14px 20px; display: flex; justify-content: space-between; align-items: center; }
.head h1 { font-size: 17pt; line-height: 1.4; }
.head .sub { color: #c9a227; font-size: 9pt; }
.head .meta { flex: 0 0 150px; text-align: left; font-size: 8pt; color: #cbd5e1; line-height: 1.7; }
.badge { display: inline-block; background: #c9a227; color: #0f2a4a; font-weight: 700; padding: 1px 10px; border-radius: 20px; font-size: 8pt; }
.info { display: flex; gap: 8px; margin: 12px 0; }
.info div { flex: 1; background: #f4f6f9; border-radius: 8px; padding: 7px 10px; font-size: 8.4pt; }
.info b { display: block; color: #7b8794; font-weight: 400; font-size: 7.6pt; }
.scorebox { display: flex; align-items: center; gap: 18px; border: 1.5px solid #e4e7eb; border-radius: 12px; padding: 10px 18px; }
.scorebox .g { flex: 0 0 150px; width: 150px; height: 150px; }
.scorebox .v { flex: 1; }
.scorebox .v h2 { font-size: 13pt; line-height: 1.5; }
.scorebox .v p { margin: 4px 0 0; color: #52606d; }
.sec { margin-top: 12px; }
.sec h3 { font-size: 11pt; color: #0f2a4a; border-right: 4px solid #c9a227; padding-right: 9px; margin-bottom: 5px; }
.axis { display: flex; align-items: center; margin: 5px 0; }
.axis .l { width: 165px; font-size: 8.8pt; }
.axis .bar { flex: 1; height: 11px; background: #eef1f5; border-radius: 6px; overflow: hidden; }
.axis .bar span { display: block; height: 100%; border-radius: 6px; float: right; }
.axis .n { width: 52px; text-align: left; font-family: 'Noto Sans'; font-size: 8.5pt; color: #52606d; direction: ltr; }
table { width: 100%; border-collapse: collapse; font-size: 8.6pt; }
td { padding: 3px 8px; border-bottom: 1px solid #eef1f5; }
td.ic { width: 26px; text-align: center; font-weight: 700; font-family: 'Noto Sans'; }
td.k { width: 140px; font-weight: 700; }
.find { border-radius: 8px; padding: 6px 12px; margin-bottom: 6px; page-break-inside: avoid; border-right: 5px solid; }
.find .top { display: flex; justify-content: space-between; align-items: center; }
.find h4 { margin: 0; font-size: 10pt; font-family: 'Noto Kufi Arabic'; }
.find .sev { font-size: 7.5pt; color: #fff; padding: 0 9px; border-radius: 12px; }
.find .num { font-family: 'Noto Sans'; color: #9aa5b1; font-size: 8pt; margin-left: 6px; }
.quotes { margin: 3px 0 2px; }
.q { display: inline-block; direction: ltr; unicode-bidi: embed; font-family: 'Noto Sans Mono','DejaVu Sans Mono',monospace; font-size: 7.8pt; background: #fff; border: 1px dashed #b0b8c4; border-radius: 4px; padding: 0 6px; margin: 2px 0 2px 4px; color: #3e4c59; }
.find p { margin: 1px 0; line-height: 1.5; }
.lock { display: inline-block; background: #fff; border: 1px solid #c9a227; color: #0f2a4a; border-radius: 12px; padding: 0 10px; font-size: 7.8pt; font-weight: 700; margin: 2px 0; }
.lock .lk { vertical-align: -2px; margin-left: 5px; }
.pg { color: #52606d; font-size: 8pt; }
.impact { font-size: 8.5pt; color: #52606d; }
.impact b { color: #1f2933; }
.kw { display: flex; align-items: center; margin: 1px 0; }
.kw .l { flex: 0 0 210px; }
.kw .dots { flex: 0 0 125px; white-space: nowrap; }
.kw .dots span { display: inline-block; width: 12px; height: 12px; border-radius: 3px; margin-left: 3px; background: #f5c6c0; }
.kw .dots span.on { background: #1e8449; }
.kw .n { flex: 0 0 48px; margin-right: 6px; text-align: right; font-family: 'Noto Sans'; font-size: 8.5pt; color: #52606d; direction: ltr; unicode-bidi: embed; white-space: nowrap; }
.legend span { display: inline-block; margin-left: 14px; font-size: 8pt; }
.legend i { display: inline-block; width: 10px; height: 10px; border-radius: 2px; margin-left: 4px; vertical-align: middle; }
.strong { background: #eaf6ee; border-radius: 10px; padding: 4px 14px; }
.strong ul { margin: 4px 0; columns: 2; }
.strong li { margin: 2px 0; }
.final { background: #0f2a4a; color: #fff; border-radius: 12px; padding: 11px 18px; margin-top: 10px; page-break-inside: avoid; }
.final h3 { color: #c9a227; border: 0; padding: 0; font-size: 12pt; }
.final .price { display: flex; justify-content: space-between; align-items: center; margin-top: 4px; }
.final ul { margin: 4px 0 0; padding-right: 18px; columns: 2; font-size: 8.8pt; }
.final .tag { background: #c9a227; color: #0f2a4a; border-radius: 10px; padding: 6px 16px; text-align: center; font-weight: 700; }
.final .tag big { font-size: 17pt; font-family: 'Noto Kufi Arabic'; display: block; line-height: 1.3; }
.final .wa { font-size: 9pt; color: #cbd5e1; margin-top: 8px; }
.pb { page-break-before: always; }
.note { font-size: 7pt; color: #8a9bb3; margin: 8px 0 0; }
"""


def build(c):
    score = c["score"]
    sc = color_for(score)
    crit = sum(1 for f in c["findings"] if f[0] == "crit")
    p = []
    p.append(f'''<div class="head"><div><span class="badge">تقرير فحص آلي</span>
<h1>تقرير توافق السيرة الذاتية<br>مع أنظمة التوظيف <span class="en">ATS</span></h1>
<div class="sub">نظام الفحص الآلي للسير الذاتية | مؤسسة العباسي سوفت</div></div>
<div class="meta">رقم التقرير: <span class="en">{c["report_no"]}</span><br>تاريخ الفحص: {DATE_AR}</div></div>''')
    p.append(f'''<div class="info">
<div><b>صاحب الملف</b>{c["name_ar"]}</div>
<div><b>الملف المفحوص</b>{file_html(c)} <span class="pg">{c["pages_ar"]}</span></div>
<div><b>المجال المستهدف</b>{c["field"]}</div></div>''')
    p.append(f'''<div class="scorebox"><div class="g">{gauge(score)}</div><div class="v">
<h2 style="color:{sc}">النتيجة: {c["verdict"]}</h2>
<p>{c["verdict_note"]}</p>
<p><b>{len(c["findings"])} ملاحظة</b> منها <b style="color:#c0392b">{crit} حرجة</b> تؤثر مباشرة على قرار الفرز.</p></div></div>''')

    rows = []
    for (label, w), v in zip(AXES, c["axes"]):
        pct = v / w * 100
        rows.append(f'<div class="axis"><div class="l">{label}</div><div class="bar"><span style="width:{pct:.0f}%;background:{color_for(pct)}"></span></div><div class="n">{v} / {w}</div></div>')
    p.append(f'<div class="sec"><h3>تفصيل النتيجة حسب محاور الفرز</h3>{"".join(rows)}</div>')

    trs = "".join(f'<tr><td class="ic" style="color:{STATUS[s][1]}">{STATUS[s][0]}</td><td class="k">{t(k)}</td><td>{t(v)}</td></tr>'
                  for k, s, v in c["parser"])
    p.append(f'<div class="sec"><h3>كيف قرأ نظام الفرز ملفك؟</h3><table>{trs}</table></div>')

    p.append('<div class="sec pb"><h3>الملاحظات المكتشفة داخل الملف</h3>'
             f'<p class="impact" style="margin-bottom:8px">النصوص داخل الإطارات في أول {QUOTED_FINDINGS} ملاحظات منقولة حرفياً من ملفك كما قرأها النظام. باقي الملاحظات مرصودة بمواضعها الدقيقة في الفحص البرمجي.</p>')
    for i, (sev, title, quotes, desc, impact) in enumerate(c["findings"], 1):
        lab, col, bg = SEV[sev]
        if i <= QUOTED_FINDINGS and quotes:
            qs = "".join(f'<span class="q">{html.escape(fn(q))}</span>' for q in quotes)
        else:
            qs = f'<span class="lock">{LOCK_SVG}{LOCK_TEXT}</span>'
        p.append(f'''<div class="find" style="background:{bg};border-color:{col}">
<div class="top"><h4><span class="num">#{i:02d}</span>{t(title)}</h4><span class="sev" style="background:{col}">{lab}</span></div>
<div class="quotes">{qs}</div><p>{t(desc)}</p><p class="impact"><b>الأثر:</b> {t(impact)}</p></div>''')
    p.append('</div>')

    kws = []
    for label, have, need in c["keywords"]:
        dots = "".join(f'<span class="{"on" if i < have else ""}"></span>' for i in range(need))
        kws.append(f'<div class="kw"><div class="l">{label}</div><div class="dots">{dots}</div><div class="n">{have} / {need}</div></div>')
    p.append(f'''<div class="sec"><h3>تغطية الكلمات المفتاحية في مجالك</h3>
<p class="impact">عدد المصطلحات الأساسية التي تبحث عنها أنظمة الفرز في كل فئة، مقابل الموجود فعلياً في ملفك.</p>
{"".join(kws)}
<div class="legend"><span><i style="background:#1e8449"></i>موجود في ملفك</span><span><i style="background:#f5c6c0"></i>مفقود</span></div></div>''')

    if c["timeline"]:
        p.append(f'''<div class="sec"><h3>الخط الزمني كما يراه النظام</h3>{timeline_svg()}
<div class="legend"><span><i style="background:#1f4e79"></i>عمل</span><span><i style="background:#8e9aaf"></i>دراسة معروضة كعمل</span><span><i style="background:#e74c3c"></i>فجوة غير مبررة</span></div></div>''')

    lis = "".join(f"<li>{t(s)}</li>" for s in c["strengths"])
    p.append(f'<div class="sec"><h3>نقاط قوة حقيقية لا يُظهرها ملفك الحالي</h3><div class="strong"><ul>{lis}</ul></div></div>')
    p.append(f'''<div class="sec"><h3>الخلاصة</h3><p>{t(c["summary"])}</p>
<p><b>التوصية:</b> الملاحظات مترابطة في بنية الملف نفسها، فتعديلها جزئياً لا يرفع النتيجة بشكل كافٍ. الملف يحتاج <b>إعادة بناء كاملة</b> على قالب متوافق، بصياغة موجهة لمجالك وكلمات مفتاحية مطابقة لمتطلبات السوق السعودي.</p></div>''')
    p.append(f'''<div class="final"><h3>إعادة بناء سيرتك الذاتية بالكامل</h3>
<ul><li>سيرة ذاتية عربي + إنجليزي</li><li>متوافقة مع أنظمة <span class="en">ATS</span></li><li>مكتوبة لوظيفتك من خبرتك الحقيقية</li><li>خطاب تقديم هدية</li><li>ملفات <span class="en">Word + PDF</span> وتعديلان مجاناً</li><li>جاهزة خلال 24 إلى 48 ساعة</li></ul>
<div class="price"><div class="wa">للطلب أو الاستفسار عبر واتساب:<br><span class="en">wa.me/{WA}</span></div>
<div class="tag">العرض الحالي<big>49 ريال</big></div></div>
<p class="note">{NOTE_TEXT}</p></div>''')

    doc = f'<!doctype html><html dir="rtl" lang="ar"><head><meta charset="utf-8"><meta name="author" content="Al-Abbasi Soft"><title>تقرير فحص ATS - {c["name_en"]}</title><style>{CSS}</style></head><body>{"".join(p)}</body></html>'
    os.makedirs(OUT, exist_ok=True)
    out = os.path.join(OUT, c["out"] + ".pdf")
    weasyprint.HTML(string=doc, base_url=HERE).write_pdf(out)
    return out


if __name__ == "__main__":
    for key in (sys.argv[1:] or CLIENTS):
        print(build(CLIENTS[key]))
