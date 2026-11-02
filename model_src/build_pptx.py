# Slajdy: GAia - jak GA zarabia przy partnerze technologicznym. Liczby z GA_GAia_model_JV_partnerzy.xlsx (przeliczony).
import sys, openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

calc = sys.argv[1]  # przeliczona kopia xlsx (wartości)
stress = sys.argv[2]  # wariant 50% adopcji
wb = openpyxl.load_workbook(calc, data_only=True)
ws = openpyxl.load_workbook(stress, data_only=True)

def summary(book):
    out = {}
    for r in book["Podsumowanie"].iter_rows(min_row=5, max_row=8, values_only=True):
        out[r[0].split(":")[0]] = r
    return out
S, ST = summary(wb), summary(ws)
M = wb["Model"]
rows = {}
cur = None
for r in M.iter_rows(values_only=True):
    if r[0] and r[0][:3].rstrip(":") in ("S1", "S1a", "S1b", "S2") and r[1] is None and r[2] is None:
        cur = r[0].split(":")[0]
    elif cur and r[0]:
        rows[(cur, r[0])] = r
def g(sc, label): return rows[(sc, label)]
W100 = {r[0]: r[1] for r in wb["Rozliczenie 100 zł"].iter_rows(values_only=True) if r[0]}

NAVY = RGBColor(0x00, 0x36, 0x65); TXT = RGBColor(0x1A, 0x1A, 0x1A); GREY = RGBColor(0x59, 0x59, 0x59)
RED = RGBColor(0xE3, 0x06, 0x13); GREEN = RGBColor(0x00, 0x8A, 0x3E); LIGHT = RGBColor(0xEA, 0xF0, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF); LBLUE = RGBColor(0x6F, 0x9F, 0xD8)

prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
BL = prs.slide_layouts[6]

def fmt(v, d=0):
    s = f"{v:,.{d}f}".replace(",", " ").replace(".", ",")
    return s.replace("-", "−")

def tb(sl, x, y, w, h, text, size=14, bold=False, color=TXT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    s = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = s.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05); tf.margin_top = tf.margin_bottom = Inches(0.03)
    paras = text if isinstance(text, list) else [text]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        runs = p if isinstance(p, list) else [(p, bold, color)]
        for t, b, c in runs:
            r = para.add_run(); r.text = t; r.font.size = Pt(size); r.font.bold = b; r.font.color.rgb = c; r.font.name = "Calibri"
        para.alignment = align; para.space_after = Pt(4)
    return s

def box(sl, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, dash=False):
    s = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE: s.adjustments[0] = 0.08
    if fill: s.fill.solid(); s.fill.fore_color.rgb = fill
    else: s.fill.background()
    if line: s.line.color.rgb = line; s.line.width = Pt(1.25)
    else: s.line.fill.background()
    if dash:
        from pptx.enum.dml import MSO_LINE
        s.line.dash_style = MSO_LINE.DASH
    s.shadow.inherit = False
    return s

def pill(sl, x, y, w, h, text, size=14):
    box(sl, x, y, w, h, fill=NAVY)
    tb(sl, x, y, w, h, text, size, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

def frame(title, num):
    sl = prs.slides.add_slide(BL)
    tb(sl, 0.6, 0.35, 12.2, 0.8, title, 30, True, NAVY)
    band = box(sl, 0, 6.95, 13.333, 0.55, fill=NAVY, shape=MSO_SHAPE.RECTANGLE)
    sl.shapes.add_picture("logo_ga.png", Inches(12.0), Inches(6.97), height=Inches(0.5))
    tb(sl, 0.6, 7.03, 9, 0.4, f"GAia | model rozliczeń z partnerem technologicznym | liczby ilustracyjne, do potwierdzenia danymi Handlu i ofertami partnerów   {num}", 10, False, WHITE)
    return sl

s2 = S["S2"]; s2s = ST["S2"]; s1a = S["S1a"]

# ---------- Slajd 1 ----------
sl = frame("GAia zarabia dla Grupy, także przy kosztach partnera", 1)
box(sl, 0.6, 1.2, 12.1, 0.75, fill=LIGHT)
tb(sl, 0.75, 1.22, 11.9, 0.72, [[("Platforma to kanał sprzedaży nawozów GA, a nie osobny biznes. ", True, NAVY),
   ("Dlatego liczymy wynik Grupy: partner bierze część przychodów platformy, ale marża na nawozach i udział w rynku zostają w GA.", False, TXT)]], 15, anchor=MSO_ANCHOR.MIDDLE)
ga_out = g("S2", "GA poza JV razem (100% dla GA)"); ga_jv = g("S2", "GA: udział w wyniku JV")
cards = [
 ("1", "Marża kanałowa", f"+{fmt(W100['ZOSTAJE W GA (ponad dzisiejszą marżę)'],1)} zł", "z każdych 100 zł nawozu sprzedanego przez platformę zamiast przez dystrybutora. 100% dla GA, poza JV."),
 ("2", "Udział w rynku", "+2 p.p. AN/CAN", f"w 5 lat ≈ 78 mln zł przychodu rocznie. Rolnik z planem nawożenia kupuje nawóz GA."),
 ("3", "Udział w JV", "51% wyniku", "z prowizji od sprzedawców zewnętrznych, reklamy i abonamentów. Partnerzy dzielą się tylko tym."),
]
for i, (n, t, big, d) in enumerate(cards):
    y = 2.15 + i * 1.37
    box(sl, 0.6, y, 5.6, 1.22, line=NAVY)
    c = sl.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.75), Inches(y + 0.31), Inches(0.6), Inches(0.6))
    c.fill.solid(); c.fill.fore_color.rgb = NAVY; c.line.fill.background()
    tb(sl, 0.75, y + 0.31, 0.6, 0.6, n, 18, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    tb(sl, 1.5, y + 0.08, 4.6, 0.45, [[(t + "   ", True, NAVY), (big, True, GREEN)]], 16)
    tb(sl, 1.5, y + 0.48, 4.6, 0.75, d, 12.5, color=GREY)
# chart
pill(sl, 6.6, 2.15, 6.1, 0.45, "Wynik GA w scenariuszu rekomendowanym (S2), mln zł", 13)
cd = CategoryChartData(); cd.categories = ["Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5"]
cd.add_series("Marża i udział w rynku (poza JV)", [0] + [round(v, 1) for v in ga_out[2:7]])
cd.add_series("Udział GA w wyniku JV", [round(v, 1) for v in ga_jv[1:7]])
gf = sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(6.6), Inches(2.65), Inches(6.1), Inches(3.2), cd)
ch = gf.chart; ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.BOTTOM; ch.legend.include_in_layout = False
ch.legend.font.size = Pt(11)
ch.plots[0].gap_width = 60; ch.plots[0].overlap = 100
for ser, col in zip(ch.series, [NAVY, LBLUE]):
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb = col; ser.invert_if_negative = False
ch.value_axis.has_major_gridlines = True; ch.value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xD9, 0xD9, 0xD9)
ch.value_axis.tick_labels.font.size = Pt(10); ch.category_axis.tick_labels.font.size = Pt(11)
ch.value_axis.format.line.fill.background()
tot = g("S2", "GA: WYNIK RAZEM (udział w JV + poza JV)")
tb(sl, 6.6, 5.85, 6.1, 0.35, "Razem: " + " | ".join(fmt(v, 0) for v in tot[1:7]), 11, False, GREY, PP_ALIGN.CENTER)
box(sl, 0.6, 6.3, 12.1, 0.55, fill=NAVY)
tb(sl, 0.75, 6.3, 11.9, 0.55, [[(f"S2: inwestycja GA zwraca się w roku {s2[7].split()[-1]}, ", True, WHITE), (f"{fmt(s2[5])} mln zł w 5 lat, NPV {fmt(s2[6])} mln zł. ", True, WHITE),
   (f"Przy połowie zakładanej adopcji: nadal +{fmt(s2s[5])} mln zł.", False, WHITE)]], 15, anchor=MSO_ANCHOR.MIDDLE)

# ---------- Slajd 2 ----------
sl = frame("Nierentowna była umowa, a nie biznes", 2)
tb(sl, 0.6, 1.15, 12.1, 0.8, [[("Na poprzednim slajdzie liczyliśmy tylko P&L platformy i opłatę 5-7% od całego GMV, także od nawozów GA. ", False, TXT),
   ("Wtedy partner zabiera więcej, niż platforma zarabia. A Grupa i tak zarabia w każdym scenariuszu, najwięcej, gdy partner nie bierze opłaty od naszych nawozów.", True, NAVY)]], 15)
hdrs = ["Scenariusz", "P&L platformy (JV)\nna plusie od", "Wynik GA\nrok 5", "Wynik GA\n5 lat", "Wynik GA 5 lat\nprzy 50% adopcji", "Partnerzy\nzarabiają 5 lat"]
names = {"S1": "S1  Iteo + własna operacja GA (25 FTE)", "S1a": "S1a  Iteo + Mirakl, 6% od całego GMV\n(wersja z 30.10)",
         "S1b": "S1b  Iteo + Mirakl, 6% tylko od sprzedawców zewn.", "S2": "S2  Iteo (JV) + OLX (revenue share)\nREKOMENDACJA"}
tbl = sl.shapes.add_table(5, 6, Inches(0.6), Inches(2.1), Inches(12.1), Inches(3.4)).table
widths = [4.1, 1.75, 1.45, 1.45, 1.75, 1.6]
for i, w in enumerate(widths): tbl.columns[i].width = Inches(w)
def cell(r, c, text, bold=False, color=TXT, fill=WHITE, size=13, align=PP_ALIGN.CENTER):
    ce = tbl.cell(r, c); ce.fill.solid(); ce.fill.fore_color.rgb = fill; ce.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = ce.text_frame; tf.clear()
    for j, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph(); p.alignment = align
        r_ = p.add_run(); r_.text = line; r_.font.size = Pt(size); r_.font.bold = bold; r_.font.color.rgb = color; r_.font.name = "Calibri"
for c, h in enumerate(hdrs): cell(0, c, h, True, WHITE, NAVY, 12)
for r, k in enumerate(["S1", "S1a", "S1b", "S2"], start=1):
    v, vs = S[k], ST[k]; hl = k == "S2"; fill = LIGHT if hl else WHITE
    cell(r, 0, names[k], hl, NAVY if hl else TXT, fill, 13, PP_ALIGN.LEFT)
    cell(r, 1, v[2], k == "S1a", RED if v[2] == "po roku 5" else TXT, fill)
    cell(r, 2, f"{fmt(v[4])} mln", hl, GREEN, fill)
    cell(r, 3, f"{fmt(v[5])} mln", True, GREEN, fill)
    cell(r, 4, f"{fmt(vs[5])} mln", hl, GREEN if vs[5] > 0 else RED, fill)
    cell(r, 5, f"{fmt(v[8])} mln", k == "S1a", RED if k == "S1a" else TXT, fill)
box(sl, 0.6, 5.7, 5.9, 1.1, fill=RED)
tb(sl, 0.75, 5.72, 5.6, 1.05, [[("S1a: ", True, WHITE), (f"platforma na zero, ale partner bierze {fmt(s1a[8])} mln zł w 5 lat, bo pobiera opłatę od naszych własnych nawozów.", False, WHITE)]], 13, anchor=MSO_ANCHOR.MIDDLE)
box(sl, 6.8, 5.7, 5.9, 1.1, fill=GREEN)
tb(sl, 6.95, 5.72, 5.6, 1.05, [[("Naprawa: ", True, WHITE), ("partner dostaje udział w tym, co platforma dobuduje (sprzedawcy zewnętrzni, reklama, abonamenty), a nie w marży nawozowej GA.", False, WHITE)]], 13, anchor=MSO_ANCHOR.MIDDLE)

# ---------- Slajd 3 ----------
sl = frame("Model rozliczeń: partner zarabia na tym, co dobuduje", 3)
cols = [("Partner", 1.7), ("Co wnosi", 2.6), ("Jak zarabia", 2.75), ("Czerwona linia GA", 3.0)]
tbl = sl.shapes.add_table(5, 4, Inches(0.6), Inches(1.25), Inches(10.05), Inches(5.0)).table
for i, (_, w) in enumerate(cols): tbl.columns[i].width = Inches(w)
data = [
 ("Iteo\n(wspólnik JV)", "Platforma agro: plan nawożenia, mapy pól, doradztwo", "49% wyniku JV\n+ 3 mln zł/rok opłaty stałej\n+ 25% przychodu z abonamentów", "GA 51% i kontrola. Opcja wykupu udziałów Iteo po roku 3 po ustalonej formule. Dane rolników należą do GA."),
 ("OLX\n(umowa revenue share)", "E-commerce, ruch, sprzedawcy zewnętrzni, wyjście poza Polskę", "35% prowizji i reklamy od sprzedawców zewnętrznych\n+ max 0,3% od GMV nawozów GA", "Bez udziałów w JV: jeden wspólnik zamiast dwóch. Nie więcej niż 0,3% od naszych nawozów."),
 ("Mirakl\n(plan B)", "Gotowe oprogramowanie marketplace, kilka FTE obsługi", "6% tylko od GMV sprzedawców zewnętrznych", "Nawozy GA wyłączone z opłaty od GMV albo stała opłata roczna."),
 ("Dystrybutorzy AD", "Magazyny, punkty odbioru, transport do rolnika", "Opłata za logistykę i odbiór, ok. 3,5% wartości nawozu", "AD to partner logistyczny platformy, a nie konkurent. Nie budujemy konfliktu kanałów."),
]
for c, (h, _) in enumerate(cols): cell(0, c, h, True, WHITE, NAVY, 13)
for r, rowd in enumerate(data, start=1):
    for c, t in enumerate(rowd):
        cell(r, c, t, c == 0, NAVY if c == 0 else (RED if c == 3 else TXT), LIGHT if r == 1 or r == 2 else WHITE, 12, PP_ALIGN.LEFT)
for i in range(5): tbl.rows[i].height = Inches(0.5 if i == 0 else 1.1)
# 100 zł
pill(sl, 10.85, 1.25, 1.95, 0.45, "100 zł nawozu GA", 12)
items = [("Marża AD dziś", W100["Dziś: marża dystrybutora (AD), której GA nie widzi"]),
         ("Logistyka AD", W100["Logistyka / fee dla AD jako punktu odbioru"]),
         ("OLX (max)", W100["Opłata dla OLX od GMV nawozów (czerwona linia max)"]),
         ("Płatności", W100["Koszt płatności"]), ("Lepszy mix", W100["Poprawa mixu"])]
y = 1.8
for t, v in items:
    tb(sl, 10.85, y, 1.3, 0.35, t, 11.5, color=TXT); tb(sl, 12.1, y, 0.7, 0.35, ("+" if v > 0 else "") + fmt(v, 1), 11.5, True, GREEN if v > 0 else RED, PP_ALIGN.RIGHT); y += 0.36
box(sl, 10.85, y + 0.05, 1.95, 0.75, fill=GREEN)
tb(sl, 10.85, y + 0.05, 1.95, 0.75, [[("+" + fmt(W100["ZOSTAJE W GA (ponad dzisiejszą marżę)"], 1) + " zł", True, WHITE)], [("więcej dla GA", False, WHITE)]], 13, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
y += 1.0
pill(sl, 10.85, y, 1.95, 0.45, "100 zł od sprzedawcy", 12)
y += 0.55
for t, v, col in [("OLX", 35, RED), ("JV: GA 51%", W100["w tym GA"], GREEN), ("JV: Iteo 49%", W100["w tym Iteo"], TXT)]:
    tb(sl, 10.85, y, 1.3, 0.35, t, 11.5); tb(sl, 12.1, y, 0.7, 0.35, fmt(v, 0), 11.5, True, col, PP_ALIGN.RIGHT); y += 0.36
tb(sl, 0.6, 6.35, 12.2, 0.5, "Zasada: im więcej partner dowiezie sprzedawców i rolników, tym więcej zarabia. Ryzyko partnera jest zmienne, a nie stałe. GA nie płaci za nic z góry poza integracją z SAP / CRM.", 12.5, True, NAVY)

# ---------- Slajd 4 ----------
sl = frame("Dalsze kroki: decyzja Zarządu w grudniu, pilot przed sezonem", 4)
steps = [
 ("3.11", "Spotkanie Empik", "Sprawdzić, czy daje coś, czego nie dają OLX / Iteo. Jeśli nie, zamknąć wątek.", "Zespół GAia"),
 ("do 16.11", "Prawdziwe liczby do modelu", "Cena netto po upustach (CO-PA), marża AD z umów, marża kontrybucyjna na tonie. Podmienić założenia w Excelu.", "Kontroling + Handel"),
 ("do 30.11", "Term sheet z Iteo", "JV 51/49, 3 mln zł/rok, 25% z abonamentów, opcja wykupu po roku 3, dane = GA.", "Zespół GAia + Prawny"),
 ("do 30.11", "Term sheet z OLX + oferta Mirakl", "OLX: 35% od sprzedawców zewn., max 0,3% od nawozów GA. Mirakl jako plan B: opłata tylko od sprzedawców zewn.", "Zespół GAia + Zakupy"),
 ("grudzień", "Decyzja Zarządu", "S2 albo plan B na podstawie modelu z prawdziwymi liczbami i ofert partnerów.", "Zarząd"),
 ("luty-kwiecień", "Pilot w sezonie wiosennym", "1 województwo, 5 AD jako punkty odbioru, AN/CAN + abonament. Mierzymy: GMV, liczbę rolników, marżę netto na tonie.", "Handel + JV"),
]
hdr = ["Termin", "Krok", "Co konkretnie", "Kto"]
tbl = sl.shapes.add_table(len(steps) + 1, 4, Inches(0.6), Inches(1.25), Inches(12.1), Inches(5.0)).table
for i, w in enumerate([1.5, 2.9, 5.6, 2.1]): tbl.columns[i].width = Inches(w)
for c, h in enumerate(hdr): cell(0, c, h, True, WHITE, NAVY, 13)
for r, st in enumerate(steps, start=1):
    for c, t in enumerate(st):
        cell(r, c, t, c in (0, 1), NAVY if c in (0, 1) else TXT, LIGHT if r == 5 else WHITE, 13, PP_ALIGN.LEFT if c > 0 else PP_ALIGN.CENTER)
tb(sl, 0.6, 6.4, 12.1, 0.45, "Prosimy o: zgodę na negocjacje term sheetów z Iteo i OLX na powyższych warunkach oraz na przygotowanie pilota.", 14, True, NAVY)

prs.save("../GA_GAia_jak_zarabiamy_z_partnerem.pptx")
print("ok")
