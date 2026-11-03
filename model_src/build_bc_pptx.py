# Prezentacja business case GAia (300 tys. użytkowników). Liczby z przeliczonego GA_GAia_business_case_300k.xlsx.
# Użycie: python3 build_bc_pptx.py <przeliczony.xlsx> <przeliczony_stres.xlsx>
import sys, openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION

wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
st = openpyxl.load_workbook(sys.argv[2], data_only=True)
def rows(book, sheet):
    d = {}
    for r in book[sheet].iter_rows(values_only=True):
        if r[0] and r[0] not in d: d[r[0]] = r
    return d
KZ, KZs = rows(wb, "Każdy zyskuje"), rows(st, "Każdy zyskuje")
U = rows(wb, "Użytkownicy i GMV"); J = rows(wb, "Spółka GAia (JV)"); T = rows(wb, "Partner Technologiczny")
MP = rows(wb, "Partner Marketplacowy"); D = rows(wb, "Dystrybutorzy (66)"); F = rows(wb, "Rolnik"); G = rows(wb, "Grupa Azoty")

NAVY = RGBColor(0x00, 0x36, 0x65); TXT = RGBColor(0x00, 0x36, 0x65); GREY = RGBColor(0x59, 0x59, 0x59)
RED = RGBColor(0xFF, 0x00, 0x00); GREEN = RGBColor(0x00, 0x8A, 0x3E); LIGHT = RGBColor(0xEA, 0xF0, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF); LBLUE = RGBColor(0x6F, 0x9F, 0xD8); MBLUE = RGBColor(0x2E, 0x6D, 0xB4)
AMBER = RGBColor(0xE6, 0x7E, 0x00); PALE = RGBColor(0xB4, 0xCB, 0xE8)

FONT = "Trebuchet MS"
def content(y): return 1.15 <= y < 6.6
def ty(y): return 1.4 + (y - 1.2) * 0.911 if content(y) else y
def th(y, h): return h * 0.911 if content(y) else h
def fs(y): return 0.93 if content(y) else 1.0
prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
def fmt(v, d=0):
    s = f"{v:,.{d}f}".replace(",", " ").replace(".", ",")
    return s.replace("-", "−")
def mln(v, d=0): return f"{fmt(v, d)} mln zł"

def tb(sl, x, y, w, h, text, size=14, bold=False, color=TXT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    f_ = fs(y); s = sl.shapes.add_textbox(Inches(x), Inches(ty(y)), Inches(w), Inches(th(y, h))); tf = s.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.06); tf.margin_top = tf.margin_bottom = Inches(0.03)
    paras = text if isinstance(text, list) else [text]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        for t, b, c in (p if isinstance(p, list) else [(p, bold, color)]):
            r = para.add_run(); r.text = t; r.font.size = Pt(round(size * f_ * 2) / 2); r.font.bold = b; r.font.color.rgb = c; r.font.name = FONT
        para.alignment = align; para.space_after = Pt(3)
    return s
def box(sl, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    s = sl.shapes.add_shape(shape, Inches(x), Inches(ty(y)), Inches(w), Inches(th(y, h)))
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE: s.adjustments[0] = 0.07
    if fill: s.fill.solid(); s.fill.fore_color.rgb = fill
    else: s.fill.background()
    if line: s.line.color.rgb = line; s.line.width = Pt(1.25)
    else: s.line.fill.background()
    s.shadow.inherit = False
    return s
def pill(sl, x, y, w, h, text, size=13, fill=NAVY):
    box(sl, x, y, w, h, fill=fill); tb(sl, x, y, w, h, text, size, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
def msg(sl, text_runs, y=6.3, fill=NAVY):
    box(sl, 0.6, y, 12.1, 0.55, fill=fill)
    tb(sl, 0.75, y, 11.9, 0.55, [text_runs], 14, anchor=MSO_ANCHOR.MIDDLE)
N = [0]
def frame(title, sub=None):
    N[0] += 1
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    tb(sl, 0.66, 0.45, 12.0, 0.65, title, 28, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
    if sub: tb(sl, 0.7, 1.0, 12.0, 0.36, sub, 12.5, False, NAVY)
    box(sl, 0, 6.66, 13.333, 0.84, fill=NAVY, shape=MSO_SHAPE.RECTANGLE)
    sl.shapes.add_picture("logo_ga.png", Inches(12.0), Inches(6.82), height=Inches(0.5))
    tb(sl, 0.66, 6.95, 10.8, 0.35, f"GAia | business case 300 tys. użytkowników | liczby ilustracyjne z modelu, do potwierdzenia   {N[0]}", 9, False, WHITE)
    return sl
def table(sl, x, y, w, data, widths, rh=0.42, size=12, hl_rows=(), first_left=True, hdr_fill=NAVY):
    f_ = fs(y); rh = th(y, rh); size = round(size * f_ * 2) / 2
    t = sl.shapes.add_table(len(data), len(data[0]), Inches(x), Inches(ty(y)), Inches(w), Inches(rh * len(data))).table
    for i, cw in enumerate(widths): t.columns[i].width = Inches(cw)
    for r, row in enumerate(data):
        t.rows[r].height = Inches(rh)
        for c, val in enumerate(row):
            color = TXT; bold = r == 0 or r in hl_rows
            if isinstance(val, tuple): val, color = val
            ce = t.cell(r, c); ce.fill.solid()
            ce.fill.fore_color.rgb = hdr_fill if r == 0 else (LIGHT if r in hl_rows else WHITE)
            ce.vertical_anchor = MSO_ANCHOR.MIDDLE
            ce.margin_left = ce.margin_right = Inches(0.06); ce.margin_top = ce.margin_bottom = Inches(0.02)
            tf = ce.text_frame; tf.clear()
            for j, line in enumerate(str(val).split("\n")):
                p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT if (c == 0 and first_left) else PP_ALIGN.CENTER
                rr = p.add_run(); rr.text = line; rr.font.size = Pt(size); rr.font.name = FONT; rr.font.bold = bold
                rr.font.color.rgb = WHITE if r == 0 else color
    return t
def chart(sl, kind, x, y, w, h, cats, series, colors, legend=True, labels=False, numfmt='#,##0'):
    cd = CategoryChartData(); cd.categories = cats
    for name, vals in series: cd.add_series(name, vals)
    ch = sl.shapes.add_chart(kind, Inches(x), Inches(ty(y)), Inches(w), Inches(th(y, h)), cd).chart
    ch.font.name = FONT; ch.font.color.rgb = NAVY
    ch.has_legend = legend and len(series) > 1
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM; ch.legend.include_in_layout = False; ch.legend.font.size = Pt(10.5)
    for s, col in zip(ch.series, colors):
        s.format.fill.solid(); s.format.fill.fore_color.rgb = col; s.invert_if_negative = False
    pl = ch.plots[0]
    try: pl.gap_width = 55
    except Exception: pass
    if kind in (XL_CHART_TYPE.COLUMN_STACKED, XL_CHART_TYPE.BAR_STACKED): pl.overlap = 100
    if labels:
        pl.has_data_labels = True; dl = pl.data_labels; dl.font.size = Pt(10.5); dl.number_format = numfmt; dl.number_format_is_linked = False
    ch.value_axis.has_major_gridlines = True; ch.value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xD9, 0xD9, 0xD9)
    ch.value_axis.tick_labels.font.size = Pt(10); ch.category_axis.tick_labels.font.size = Pt(10.5)
    ch.value_axis.format.line.fill.background()
    ch.has_title = False
    if labels and kind == XL_CHART_TYPE.COLUMN_STACKED:
        d0 = ch.series[0].data_labels; d0.font.color.rgb = WHITE; d0.font.size = Pt(10.5); d0.number_format = numfmt; d0.number_format_is_linked = False
    ch.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    return ch

Y15 = ["Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5"]; Y05 = ["Rok 0"] + Y15
k = lambda name: KZ[name]; ks = lambda name: KZs[name]
GA, TP, MPk, AD, RL, JV = ("Grupa Azoty", "Partner Technologiczny (SoftwareHouse)", "Partner Marketplacowy / Handlowy",
                          "66 dystrybutorów (AD) razem", "Rolnicy (korzyść)", "Spółka GAia (JV) - przepływy pieniężne")
C = rows(wb, "CAPEX budowa"); MC = rows(wb, "Moc platformy")
users = U["Użytkownicy zarejestrowani (tys.)"]; gmv = U["GMV RAZEM"]; fg = U["GMV nawozów GA"]; g3 = U["GMV innych produktów od sprzedawców (3P, w tym AD)"]
pen = U["Nawozy przez platformę / sprzedaż krajowa GA"]; buy = U["Aktywni kupujący (tys.) = użytkownicy × 35%"]
rev = J["PRZYCHODY RAZEM"]; opex = J["OPEX RAZEM"]; ebitda = J["EBITDA"]; capex = J["CAPEX: budowa platformy (Partner Technologiczny, koszt + 15%)"]
am = J["Amortyzacja (4 lata, od roku po wydatku)"]; ebit = J["EBIT (wynik przed podatkiem)"]; fcf = J["PRZEPŁYWY PIENIĘŻNE (EBITDA - CAPEX)"]; cum = J["Przepływy skumulowane"]
fund, fga, ftp = J["Maksymalne zapotrzebowanie na finansowanie"][1], J["Wkład Grupy Azoty (60%)"][1], J["Wkład Partnera Technologicznego (40%)"][1]
capcost = C["CAPEX: KOSZT BUDOWY PO STRONIE PARTNERA TECHNOLOGICZNEGO"]; capj = C["CAPEX PŁACI SPÓŁKA GAia (koszt + marża)"]; team = C["Zespół budowy razem"]
v = lambda r, i: r[i] if r[i] is not None else 0
def yrow(lab, r, d=1, rng=range(1, 7), color=None, tot=True):
    out = [lab] + [(fmt(v(r, i), d), color(v(r, i)) if color else TXT) for i in rng]
    if tot: out.append(fmt(v(r, 7), d))
    return out
sign = lambda x: GREEN if x > 0.05 else (RED if x < -0.05 else TXT)

# ===== 1. Teza =====
sl = frame("300 tys. rolników w 5 lat: zarabia każdy",
           "GA wnosi produkty, produkcję, logistykę, 66 AD i 5 funkcjonalności. Partnerzy: technologię i handel.")
tiles = [(f"{fmt(users[6])} tys.", "użytkowników w roku 5"), (f"{fmt(gmv[6]/1000, 2)} mld zł", "obrotu (GMV) w roku 5"),
         (mln(rev[6]), "przychodu platformy w roku 5"), (mln(capcost[6], 1), "CAPEX budowy w latach 0-2")]
for i, (big, small) in enumerate(tiles):
    x = 0.6 + i * 3.07
    box(sl, x, 1.45, 2.9, 1.2, fill=NAVY)
    tb(sl, x, 1.5, 2.9, 0.7, big, 26, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    tb(sl, x, 2.15, 2.9, 0.45, small, 12.5, False, WHITE, PP_ALIGN.CENTER)
data = [["Uczestnik (przepływy pieniężne)", "Wynik w roku 5", "Suma 5 lat (z rokiem 0)", "NPV (10%)", "Zwrot od"]]
for nm, lab in [(GA, "Grupa Azoty"), (TP, "Partner Technologiczny"), (MPk, "Partner Marketplacowy / Handlowy"), (AD, "66 dystrybutorów"), (RL, "Rolnicy (oszczędności)")]:
    r = k(nm); data.append([lab, (mln(r[6], 1), GREEN), (mln(r[7], 1), GREEN), mln(r[8], 1), r[10]])
table(sl, 0.6, 2.85, 12.1, data, [4.0, 2.0, 2.4, 1.9, 1.8], rh=0.5, size=13, hl_rows=(1,))
msg(sl, [("Przy połowie użytkowników (150 tys.) wszyscy nadal są na plusie w 5 lat. ", True, WHITE),
         ("GA: +" + mln(ks(GA)[7]) + ", Partner Technologiczny: +" + mln(ks(TP)[7]) + ", Partner Marketplacowy: +" + mln(ks(MPk)[7]) + ".", False, WHITE)])

# ===== 2. Kto co wnosi =====
sl = frame("Kto co wnosi i jak zarabia", "Partnerzy nazwani po funkcji. Podział spółki GAia: Grupa Azoty 60%, Partner Technologiczny 40%.")
cols = [
 ("Grupa Azoty", NAVY, ["5 gotowych funkcjonalności: plan nawożenia, CRM z profilami gospodarstw, portal AD, Nawozy.eu, koszyk → zamówienie do AD",
                         "66 dystrybutorów zaproszonych na platformę", "Własne produkty, produkcja i logistyka", "Marka i doradcy w regionach"],
  "60% spółki GAia + lepsza cena netto + mix + udział w rynku"),
 ("Partner Technologiczny\n(SoftwareHouse)", MBLUE, ["Silnik platformy agro", "Budowa: 504 osobomiesiące, potem zespół 10-22 osób", "Utrzymanie 24/7, chmura, bezpieczeństwo",
                                                    "Dane satelitarne i pogodowe, asystent AI"],
  "Koszt + 15% na budowie i utrzymaniu + 40% gotówki spółki GAia"),
 ("Partner Marketplacowy\n/ Handlowy", LBLUE, ["Silnik e-commerce i płatności", "Sprzedawcy nasion, ŚOR, części, usług", "Ruch i doświadczenie marketplace",
                                              "Obsługa sprzedawców"],
  "30% prowizji i reklamy od sprzedawców innych produktów. Bez opłat od nawozów GA"),
]
for i, (name, colr, items, earn) in enumerate(cols):
    x = 0.6 + i * 4.1
    box(sl, x, 1.45, 3.9, 0.7, fill=colr); tb(sl, x, 1.45, 3.9, 0.7, name.split("\n"), 15, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    box(sl, x, 2.2, 3.9, 2.45, line=colr)
    tb(sl, x + 0.1, 2.28, 3.7, 0.3, "WNOSI", 11, True, GREY)
    tb(sl, x + 0.1, 2.55, 3.7, 2.05, ["• " + t for t in items], 12.5)
    box(sl, x, 4.72, 3.9, 0.85, fill=LIGHT)
    tb(sl, x + 0.1, 4.74, 3.7, 0.82, [[("ZARABIA: ", True, NAVY), (earn, False, TXT)]], 12.5, anchor=MSO_ANCHOR.MIDDLE)
box(sl, 0.6, 5.7, 5.95, 0.5, line=GREEN); tb(sl, 0.65, 5.7, 5.85, 0.5, [[("66 dystrybutorów: ", True, GREEN), ("realizują zamówienia ze swojego regionu i sprzedają własny asortyment", False, TXT)]], 12, anchor=MSO_ANCHOR.MIDDLE)
box(sl, 6.75, 5.7, 5.95, 0.5, line=GREEN); tb(sl, 6.8, 5.7, 5.85, 0.5, [[("Rolnik: ", True, GREEN), ("darmowy plan nawożenia, niższy koszt nawożenia, ceny, finansowanie", False, TXT)]], 12, anchor=MSO_ANCHOR.MIDDLE)
msg(sl, [("Dlaczego GA ma większość: ", True, WHITE), ("bez 5 funkcjonalności Partner musiałby zbudować o ok. 12,5 mln zł więcej, a 66 AD, produktów i logistyki nie da się kupić.", False, WHITE)], y=6.3)



# ===== 3. Ścieżka + dane =====
sl = frame("Ścieżka do 300 tys. użytkowników: dane", "GMV = aktywni kupujący × zakupy na kupującego. 300 tys. to ok. 25% z 1,21 mln rolników w ARiMR.")
chart(sl, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.6, 1.5, 5.6, 1.9, Y15, [("Użytkownicy, tys.", [round(x) for x in users[2:7]])], [NAVY], labels=True)
tb(sl, 0.6, 1.2, 5.6, 0.3, "Użytkownicy na koniec roku, tys.", 11, True, NAVY)
chart(sl, XL_CHART_TYPE.COLUMN_STACKED, 6.7, 1.5, 6.0, 1.9, Y15,
      [("Nawozy GA", [round(x) for x in fg[2:7]]), ("Inne produkty (3P)", [round(x) for x in g3[2:7]])], [NAVY, LBLUE], labels=True)
tb(sl, 6.7, 1.2, 6.0, 0.3, "GMV, mln zł", 11, True, NAVY)
nu = U["Nowi użytkownicy w roku (tys.)"]; bf = U["Zakupy nawozów na kupującego (tys. zł/rok)"]; b3 = U["Zakupy innych produktów na kupującego (tys. zł/rok)"]
data = [["Dane", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Skąd"],
        ["Użytkownicy zarejestrowani, tys."] + [fmt(x) for x in users[2:7]] + ["cel"],
        ["Nowi użytkownicy w roku, tys."] + [fmt(x) for x in nu[2:7]] + ["przyrost"],
        ["Aktywni kupujący, tys."] + [fmt(x) for x in buy[2:7]] + ["35% użytkowników"],
        ["Zakupy nawozów na kupującego, tys. zł"] + [fmt(x, 1) for x in bf[2:7]] + ["założenie"],
        ["Zakupy innych produktów na kupującego, tys. zł"] + [fmt(x, 1) for x in b3[2:7]] + ["założenie"],
        ["GMV nawozów GA, mln zł"] + [fmt(x) for x in fg[2:7]] + ["kupujący × zakupy"],
        ["GMV innych produktów (3P), mln zł"] + [fmt(x) for x in g3[2:7]] + ["kupujący × zakupy"],
        ["GMV razem, mln zł"] + [fmt(x) for x in gmv[2:7]] + ["suma"],
        ["Nawozy przez platformę / sprzedaż krajowa GA"] + [f"{fmt(x*100)}%" for x in pen[2:7]] + ["÷ 3,69 mld zł"]]
table(sl, 0.6, 3.45, 12.1, data, [4.3, 1.2, 1.2, 1.2, 1.2, 1.2, 1.8], rh=0.27, size=11, hl_rows=(8,))

# ===== 4. Moc platformy =====
sl = frame("Moc platformy w liczbach", "Rok 5, przy 300 tys. użytkowników. Wszystko wynika z GMV i średniego zamówienia.")
tiles = [(f"{fmt(users[6])} tys.", "użytkowników"), (f"{fmt(buy[6])} tys.", "aktywnych kupujących"), (f"{fmt(MC['Abonenci premium (tys.)'][6])} tys.", "abonentów premium"),
         (f"{fmt(MC['Sprzedawcy (66 AD + inni)'][6])}", "sprzedawców (66 AD + inni)"), (f"{fmt(MC['Zamówienia razem (tys.)'][6])} tys.", "zamówień rocznie"),
         (f"{fmt(MC['Nawozy przez platformę (tys. t)'][6])} tys. t", "nawozów GA przez platformę"), (f"{fmt(MC['Zamówienia nawozów na 1 AD dziennie w szczycie'][6])}", "zamówień na 1 AD dziennie w szczycie"),
         (f"{fmt(MC['Przychód na użytkownika (zł/rok)'][6])} zł", "przychodu na użytkownika rocznie")]
for i, (big, small) in enumerate(tiles):
    x = 0.6 + (i % 4) * 3.07; yy = 1.25 + (i // 4) * 1.2
    box(sl, x, yy, 2.9, 1.08, fill=NAVY if i < 4 else MBLUE)
    tb(sl, x, yy + 0.04, 2.9, 0.6, big, 22, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    tb(sl, x, yy + 0.62, 2.9, 0.4, small, 11.5, False, WHITE, PP_ALIGN.CENTER)
mrows = [("Zamówienia razem, tys.", "Zamówienia razem (tys.)", 0), ("Nawozy przez platformę, tys. t", "Nawozy przez platformę (tys. t)", 0),
         ("Zamówienia nawozów na 1 AD / dzień w szczycie", "Zamówienia nawozów na 1 AD dziennie w szczycie", 1), ("Przychód platformy / GMV", "Przychód platformy / GMV", -1),
         ("Przychód na użytkownika, zł/rok", "Przychód na użytkownika (zł/rok)", 0), ("OPEX na użytkownika, zł/rok", "OPEX na użytkownika (zł/rok)", 0),
         ("EBITDA na użytkownika, zł/rok", "EBITDA na użytkownika (zł/rok)", 0)]
data = [["Rok po roku", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5"]]
for lab, key, d in mrows:
    r = MC[key]
    data.append([lab] + [(f"{fmt(x*100,1)}%" if d < 0 else fmt(x, d), sign(x) if 'EBITDA' in key else TXT) for x in r[2:7]])
table(sl, 0.6, 3.75, 12.1, data, [4.6, 1.5, 1.5, 1.5, 1.5, 1.5], rh=0.3, size=11.5)

# ===== 5. CAPEX =====
sl = frame("CAPEX: co budujemy, za ile i kiedy", f"{fmt(team[2])} osobomiesięcy × 20 tys. zł + koszty zewnętrzne = {mln(capcost[6],1)}. Budowa rozłożona na 3 lata.")
mods = [r for r in wb["CAPEX budowa"].iter_rows(min_row=5, max_row=19, values_only=True)]
data = [["Moduł", "Osobo-\nmiesiące", "Koszt,\nmln zł", "Rok 0", "Rok 1", "Rok 2"]]
for r in mods: data.append([r[0], fmt(r[2]), fmt(r[6], 2), fmt(r[7], 2), fmt(r[8], 2), fmt(r[9], 2)])
ext = C["Koszty zewnętrzne: licencje narzędzi, testy penetracyjne, sklepy z aplikacjami"]
data.append(["Koszty zewnętrzne (licencje, testy, sklepy z aplikacjami)", "", fmt(ext[6], 2), fmt(ext[7], 2), fmt(ext[8], 2), fmt(ext[9], 2)])
data.append(["CAPEX: koszt Partnera Technologicznego", fmt(team[2]), fmt(capcost[6], 2), fmt(capcost[7], 2), fmt(capcost[8], 2), fmt(capcost[9], 2)])
data.append(["Spółka GAia płaci (koszt + 15%)", "", fmt(capj[6], 2), fmt(capj[7], 2), fmt(capj[8], 2), fmt(capj[9], 2)])
table(sl, 0.6, 1.2, 8.4, data, [3.9, 0.9, 0.9, 0.9, 0.9, 0.9], rh=0.255, size=9.5, hl_rows=(17, 18))
box(sl, 9.2, 1.2, 3.5, 4.95, line=NAVY)
tb(sl, 9.3, 1.28, 3.3, 4.85, [
    [("Harmonogram", True, NAVY)],
    [(f"Rok 0 (6-9 mies. przed startem): {mln(capj[7],1)}", False, TXT)], [("MVP: zakupy, płatności, zamówienia do AD, integracja z SAP", False, GREY)],
    [(f"Rok 1: {mln(capj[8],1)}", False, TXT)], [("aplikacja mobilna, mObywatel / ARiMR, mapy pól", False, GREY)],
    [(f"Rok 2: {mln(capj[9],1)}", False, TXT)], [("asystent AI, premium, reklama, finansowanie", False, GREY)],
    [("Księgowo", True, NAVY)], [("Aktywo spółki GAia, amortyzacja 4 lata od roku po wydatku.", False, TXT)],
    [("Poza tym CAPEX", True, NAVY)], [("GA: integracja SAP / CRM 3 mln zł (lata 0-1). Partner Marketplacowy: integracja 1,5 mln zł po jego stronie.", False, TXT)],
    [("5 funkcjonalności GA oszczędza ok. 12,5 mln zł budowy.", True, GREEN)]], 11.5)
msg(sl, [("Wcześniej było 8 mln zł na start. ", True, WHITE), (f"Teraz każdy moduł ma osobomiesiące i rok, a {mln(capj[6],1)} płacimy w 3 transzach, za dowiezione moduły.", False, WHITE)], y=6.3)

# ===== 6. Przychody i OPEX =====
sl = frame("Spółka GAia: przychody i OPEX rok po roku", "mln zł. Każda pozycja to stawka × baza z modelu.")
rl = [("Prowizje: GMV 3P × 7%", "Prowizje: GMV 3P × 7%"), ("Reklama: GMV 3P × 1,5%", "Reklama: GMV 3P × 1,5%"), ("Abonamenty: abonenci × 1 200 zł", "Abonamenty: abonenci premium × 1 200 zł"),
      ("Finansowanie: GMV × 0,4%", "Finansowanie: GMV × 0,4%"), ("Dane dla producentów", "Dane i analizy dla producentów"), ("Opłata GA: GMV nawozów × 1,2%", "Opłata GA: GMV nawozów × 1,2%")]
ol = [("Partner Technologiczny: utrzymanie (koszt + 15%)", "Partner Technologiczny: utrzymanie i rozwój (koszt + 15%)"), ("Partner Marketplacowy: 30% prowizji i reklamy", "Partner Marketplacowy: 30% prowizji i reklamy"),
      ("Płatności: GMV × 0,6%", "Płatności: GMV × 0,6%"), ("Marketing: nowi × 60 zł + marka 3 mln", "Marketing: nowi użytkownicy × 60 zł + marka 3 mln zł"),
      ("Obsługa klienta", "Obsługa klienta: 1 konsultant na 20 tys. użytkowników"), ("Zarząd i zespół spółki (6 osób)", "Zarząd i zespół spółki (6 osób)")]
data = [["mln zł", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma"]]
for lab, key in rl: data.append(yrow(lab, J[key]))
data.append(yrow("PRZYCHODY RAZEM", rev))
for lab, key in ol: data.append(yrow(lab, J[key]))
data.append(yrow("OPEX RAZEM", opex))
data.append(yrow("EBITDA", ebitda, color=sign))
data.append(["Marża EBITDA"] + ["-"] + [f"{fmt(x*100)}%" for x in J["Marża EBITDA"][2:7]] + [f"{fmt(ebitda[7]/rev[7]*100)}%"])
table(sl, 0.6, 1.2, 12.1, data, [4.6, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.5], rh=0.285, size=10.5, hl_rows=(7, 14, 15))

# ===== 7. P&L i przepływy =====
sl = frame("Spółka GAia: EBITDA, CAPEX i gotówka", f"Potrzeba finansowania: {mln(fund,1)} (GA {mln(fga,1)}, Partner Technologiczny {mln(ftp,1)}). Gotówka zwraca się w roku 4.")
cc = chart(sl, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.6, 1.2, 6.4, 3.3, Y05,
           [("EBITDA", [round(v(ebitda, i), 1) for i in range(1, 7)]), ("CAPEX", [-round(v(capex, i), 1) for i in range(1, 7)]), ("Przepływy skumulowane", [round(v(cum, i), 1) for i in range(1, 7)])],
           [NAVY, RED, GREEN])
data = [["mln zł", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma"],
        yrow("Przychody", rev), yrow("OPEX", opex), yrow("EBITDA", ebitda, color=sign), yrow("Amortyzacja", am),
        yrow("EBIT (przed podatkiem)", ebit, color=sign), yrow("CAPEX", capex), yrow("Przepływy (EBITDA - CAPEX)", fcf, color=sign),
        yrow("Przepływy skumulowane", cum, color=sign, tot=False) + [""]]
table(sl, 0.6, 4.6, 12.1, data, [4.1, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 1.5], rh=0.205, size=9.5, hl_rows=(7,))
box(sl, 7.3, 1.2, 5.4, 3.3, line=NAVY)
tb(sl, 7.4, 1.28, 5.2, 3.2, [
    [("Skąd potrzeba finansowania", True, NAVY)],
    [(f"Rok 0: budowa {mln(capex[1],1)} + start {mln(opex[1],1)}", False, TXT)],
    [(f"Lata 1-2: dalsza budowa i ujemna EBITDA, bo użytkowników jest mało", False, TXT)],
    [(f"Najniżej: {mln(min(v(cum,i) for i in range(1,7)),1)} na koniec roku 2", True, RED)],
    [("Jak finansujemy", True, NAVY)],
    [(f"GA 60%: {mln(fga,1)}  |  Partner Technologiczny 40%: {mln(ftp,1)}", True, TXT)],
    [("Kiedy wraca", True, NAVY)],
    [(f"Od roku 3 gotówka dodatnia. Suma 5 lat: +{mln(fcf[7],1)}", True, GREEN)]], 12)

# ===== 8. Przepływ pieniędzy =====
sl = frame("Przepływ pieniędzy w roku 5: kto ile dostaje", f"Rolnicy kupują za {mln(gmv[6])}. Do spółki GAia trafia {mln(rev[6],1)}, czyli {fmt(rev[6]/gmv[6]*100,1)}% obrotu.")
def node(x, y_, w, h, title, val, fill):
    box(sl, x, y_, w, h, fill=fill); tb(sl, x, y_, w, h, [[(title, False, WHITE)], [(val, True, WHITE)]], 11.5, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
def arrow(x, y_, w=0.35):
    a = sl.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(ty(y_)), Inches(w), Inches(th(y_, 0.25))); a.fill.solid(); a.fill.fore_color.rgb = PALE; a.line.fill.background()
node(0.6, 2.6, 1.9, 1.3, "Rolnicy płacą", mln(gmv[6]), NAVY)
arrow(2.55, 3.1)
node(3.0, 1.3, 2.4, 0.85, "Nawozy GA (do GA i AD)", mln(fg[6]), MBLUE)
node(3.0, 4.4, 2.4, 0.85, "Inne produkty (do sprzedawców)", mln(g3[6]), MBLUE)
srcs = [("Prowizje 7%", J["Prowizje: GMV 3P × 7%"][6]), ("Reklama 1,5%", J["Reklama: GMV 3P × 1,5%"][6]), ("Abonamenty", J["Abonamenty: abonenci premium × 1 200 zł"][6]),
        ("Finansowanie", J["Finansowanie: GMV × 0,4%"][6]), ("Dane", J["Dane i analizy dla producentów"][6]), ("Opłata GA 1,2%", J["Opłata GA: GMV nawozów × 1,2%"][6])]
yy = 1.3
for lab, val in srcs:
    box(sl, 5.75, yy, 1.85, 0.56, line=NAVY); tb(sl, 5.75, yy, 1.85, 0.56, [[(lab + "  ", False, TXT), (fmt(val, 1), True, NAVY)]], 11, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    yy += 0.66
arrow(7.65, 3.1)
node(8.05, 2.45, 1.8, 1.6, "Spółka GAia: przychód", mln(rev[6], 1), NAVY)
arrow(9.9, 3.1)
outs = [("Partner Technologiczny", J["Partner Technologiczny: utrzymanie i rozwój (koszt + 15%)"][6], MBLUE), ("Partner Marketplacowy", J["Partner Marketplacowy: 30% prowizji i reklamy"][6], LBLUE),
        ("Płatności", J["Płatności: GMV × 0,6%"][6], GREY), ("Marketing, obsługa, zespół", J["Marketing: nowi użytkownicy × 60 zł + marka 3 mln zł"][6] + J["Obsługa klienta: 1 konsultant na 20 tys. użytkowników"][6] + J["Zarząd i zespół spółki (6 osób)"][6], GREY),
        ("EBITDA → GA 60%", ebitda[6] * 0.6, GREEN), ("EBITDA → Partner Techn. 40%", ebitda[6] * 0.4, GREEN)]
yy = 1.3
for lab, val, col in outs:
    box(sl, 10.3, yy, 2.4, 0.56, fill=col); tb(sl, 10.3, yy, 2.4, 0.56, [[(lab + "  ", False, WHITE), (fmt(val, 1), True, WHITE)]], 10.5, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    yy += 0.66
tb(sl, 3.0, 2.25, 2.4, 2.0, [[("Z nawozów do spółki trafia tylko 1,2% (opłata GA). Resztę dzielą GA i AD.", False, GREY)], [("Ze sprzedaży innych produktów spółka bierze prowizję 7% i reklamę 1,5%.", False, GREY)]], 10.5)
msg(sl, [("Partner Technologiczny w roku 5: ", True, WHITE), (f"{fmt(J['Partner Technologiczny: utrzymanie i rozwój (koszt + 15%)'][6],1)} za utrzymanie + {fmt(ebitda[6]*0.4,1)} z udziału. ", False, WHITE),
         ("Partner Marketplacowy: ", True, WHITE), (f"{fmt(J['Partner Marketplacowy: 30% prowizji i reklamy'][6],1)} mln zł z prowizji od sprzedawców.", False, WHITE)], y=6.3)

# ===== 9. Partner Technologiczny: koszty =====
cost = T["KOSZTY PARTNERA RAZEM (CAPEX + OPEX)"]; cpu = T["OPEX technologii na użytkownika (zł/rok)"]; opx = T["OPEX Partnera razem"]
sl = frame("Partner Technologiczny: koszty po jego stronie", f"Koszt w 5 lat: {mln(cost[7],1)} = budowa {mln(capcost[6],1)} + utrzymanie {mln(opx[7],1)}. OPEX na użytkownika spada z {fmt(cpu[2])} do {fmt(cpu[6])} zł.")
lines = [("Budowa (CAPEX, z arkusza CAPEX)", "Budowa platformy i marketplace'u (koszt)"), ("Zespół rozwoju (10 → 22 FTE)", "Zespół rozwoju"), ("Utrzymanie 24/7, SRE (3 → 10 FTE)", "Utrzymanie 24/7 i wsparcie"),
         ("Chmura: 0,8 mln + 6 zł / użytkownik", "Chmura (stała + na użytkownika)"), ("Dane satelitarne i pogodowe: 4 zł / użytk.", "Dane satelitarne i pogodowe"),
         ("Asystent AI: 3 zł / użytkownik", "Asystent AI / modele agronomiczne"), ("Bezpieczeństwo, RODO, NIS2", "Bezpieczeństwo i zgodność")]
data = [["Koszty Partnera, mln zł", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma"]]
for lab, key in lines: data.append(yrow(lab, T[key]))
data.append(yrow("OPEX Partnera (bez budowy)", opx))
data.append(yrow("KOSZTY RAZEM (CAPEX + OPEX)", cost))
data.append(["OPEX na użytkownika, zł/rok", "-"] + [fmt(x) for x in cpu[2:7]] + [""])
table(sl, 0.6, 1.2, 12.1, data, [4.6, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.5], rh=0.4, size=11.5, hl_rows=(8, 9))
msg(sl, [("Każdy koszt Partnera spółka zwraca z marżą 15%. ", True, WHITE), ("Partner nie finansuje kosztów pracy, finansuje tylko 40% startu spółki jako wspólnik.", False, WHITE)], y=6.3)

# ===== 10. Partner Technologiczny: z czego i kiedy zarabia =====
s1, s2, s3, tres, tcum = T["1. Marża na budowie (lata 0-2)"], T["2. Marża na utrzymaniu (lata 1-5)"], T["3. 40% przepływów spółki GAia (dopłaty na start, potem dywidenda)"], T["WYNIK PARTNERA TECHNOLOGICZNEGO (1 + 2 + 3)"], T["Wynik skumulowany"]
sl = frame("Partner Technologiczny: z czego i kiedy zarabia", f"3 źródła. Razem w 5 lat: +{mln(tres[7],1)}, NPV {mln(k(TP)[8],1)}, zwrot w {k(TP)[10]}.")
chart(sl, XL_CHART_TYPE.COLUMN_STACKED, 0.6, 1.2, 6.6, 3.55, Y05,
      [("1. Marża na budowie", [round(v(s1, i), 1) for i in range(1, 7)]), ("2. Marża na utrzymaniu", [round(v(s2, i), 1) for i in range(1, 7)]), ("3. Udział 40% w gotówce spółki", [round(v(s3, i), 1) for i in range(1, 7)])],
      [AMBER, MBLUE, GREEN])
fee = T["Przychody od spółki GAia razem"]
data = [["mln zł", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma"], yrow("Przychód od spółki (budowa + utrzymanie)", fee),
        yrow("1. Marża 15% na budowie", s1, color=sign), yrow("2. Marża 15% na utrzymaniu", s2, color=sign), yrow("3. 40% gotówki spółki (dopłata / dywidenda)", s3, color=sign),
        yrow("WYNIK PARTNERA (1 + 2 + 3)", tres, color=sign), yrow("Skumulowany", tcum, color=sign, tot=False) + [""]]
table(sl, 0.6, 4.85, 12.1, data, [4.6, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.5], rh=0.205, size=9.5, hl_rows=(5,))
box(sl, 7.5, 1.2, 5.2, 3.55, line=NAVY)
tb(sl, 7.6, 1.28, 5.0, 3.45, [
    [("Lata 0-2: budowa", True, AMBER)], [(f"Marża 15% na budowie: {mln(s1[7],1)}. Dopłaca 40% startu spółki.", False, TXT)],
    [("Lata 1-5: utrzymanie i rozwój", True, MBLUE)], [(f"Marża 15% co roku: {mln(s2[7],1)} w 5 lat, rośnie z użytkownikami.", False, TXT)],
    [("Od roku 3: udział w spółce", True, GREEN)], [(f"40% gotówki spółki: {mln(s3[7],1)} netto w 5 lat (w roku 5: {mln(s3[6],1)}).", False, TXT)],
    [("To główny zarobek Partnera: ", True, NAVY), ("im więcej rolników, tym więcej zarabia.", False, TXT)]], 12)

# ===== 11. Partner Marketplacowy =====
mg3, mrc, mra, msrc, msra, mrev, mres, mcum = (MP[x] for x in ["GMV innych produktów od sprzedawców (3P)", "Prowizja platformy (7% GMV 3P)", "Reklama (1,5% GMV 3P)", "1. Partner: 30% prowizji",
                                                               "2. Partner: 30% reklamy", "PRZYCHÓD PARTNERA (1 + 2)", "WYNIK PARTNERA MARKETPLACOWEGO", "Wynik skumulowany"])
pct = MP["Udział Partnera w przychodach platformy"]
sl = frame("Partner Marketplacowy: z czego i kiedy zarabia", f"Zarabia tylko na sprzedaży innych produktów: 2,55 zł ze 100 zł. W roku 5 to {mln(mrev[6],1)}, czyli {fmt(pct[6]*100)}% przychodów platformy.")
data = [["mln zł", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma"],
        yrow("GMV innych produktów (3P)", mg3, d=0), yrow("Prowizja platformy: 7% GMV 3P", mrc), yrow("Reklama: 1,5% GMV 3P", mra),
        yrow("1. Partner: 30% prowizji", msrc), yrow("2. Partner: 30% reklamy", msra), yrow("PRZYCHÓD PARTNERA", mrev),
        yrow("Od nawozów GA", MP["Od nawozów GA Partner dostaje"]),
        yrow("Koszt: integracja", MP["Integracja"]), yrow("Koszt: zespół i obsługa sprzedawców", MP["Zespół i obsługa sprzedawców"]), yrow("Koszt: ruch i wsparcie (10%)", MP["Ruch i wsparcie (10% przychodu)"]),
        yrow("WYNIK PARTNERA", mres, color=sign), yrow("Skumulowany", mcum, color=sign, tot=False) + [""]]
table(sl, 0.6, 1.2, 8.3, data, [3.1, 0.7, 0.7, 0.7, 0.75, 0.75, 0.75, 0.85], rh=0.36, size=10.5, hl_rows=(6, 11))
box(sl, 9.1, 1.2, 3.6, 4.7, line=NAVY)
tb(sl, 9.2, 1.28, 3.4, 4.6, [
    [("Ze 100 zł sprzedaży innych produktów", True, NAVY)],
    [("7,00 zł prowizji + 1,50 zł reklamy", False, TXT)], [("→ Partner 30%: 2,55 zł", True, GREEN)], [("→ spółka GAia: 5,95 zł", False, TXT)],
    [("Ze 100 zł nawozów GA", True, NAVY)], [("→ Partner: 0 zł", True, RED)],
    [("Kiedy zarabia", True, NAVY)], [(f"Rok 0: integracja −1,5 mln zł. Na plusie od roku 2, zwrot w {k(MPk)[10]}.", False, TXT)],
    [(f"5 lat: +{mln(mres[7],1)}", True, GREEN)]], 12)
msg(sl, [("Partner zarabia tyle, ile dowiezie sprzedawców i zakupów. ", True, WHITE), ("Bez opłaty stałej i bez opłat od nawozów GA.", False, WHITE)], y=6.3)
# ===== 7. AD i rolnik =====
per = D["Na jednego AD (tys. zł)"]; adres = D["WYNIK 66 AD RAZEM"]; fper = F["Na aktywnego kupującego (zł/rok)"]
sl = frame("Dystrybutorzy i rolnicy też zyskują", "Platforma nie omija 66 dystrybutorów. Zamówienie z platformy trafia do lokalnego AD.")
box(sl, 0.6, 1.45, 6.0, 4.6, line=NAVY); pill(sl, 0.6, 1.45, 6.0, 0.5, "66 dystrybutorów (AD)")
tb(sl, 0.75, 2.05, 5.7, 4.0, [
    [("Ten sam zysk na tonie: ", True, NAVY), ("marża 5% zamiast 7%, ale koszt sprzedaży 1% zamiast 3% (bez negocjacji, kredyt przejmuje bank).", False, TXT)],
    [("Nowy wolumen: ", True, NAVY), ("rolnicy z platformy i udział zabrany konkurencji, ok. 40% zamówień.", False, TXT)],
    [("Własny sklep: ", True, NAVY), ("AD wystawiają nasiona, ŚOR i części dla 300 tys. użytkowników.", False, TXT)],
    [("", False, TXT)],
    [(f"+{fmt(per[6])} tys. zł", True, GREEN), (" na jednego AD w roku 5", False, TXT)],
    [(f"+{mln(adres[7],1)}", True, GREEN), (" dla 66 AD razem w 5 lat", False, TXT)]], 13.5)
box(sl, 6.7, 1.45, 6.0, 4.6, line=NAVY); pill(sl, 6.7, 1.45, 6.0, 0.5, "Rolnik")
tb(sl, 6.85, 2.05, 5.7, 4.0, [
    [("Plan nawożenia za darmo: ", True, NAVY), ("ok. 5% mniej wydane na nawozy przy tym samym plonie.", False, TXT)],
    [("Ceny i dostępność w jednym miejscu: ", True, NAVY), ("mniej telefonów i objazdów.", False, TXT)],
    [("Finansowanie zakupów: ", True, NAVY), ("kredyt kupiecki od banku-partnera w koszyku.", False, TXT)],
    [("Premium dla chętnych: ", True, NAVY), ("mapy pól, doradca, ok. 100 zł/mies.", False, TXT)],
    [(f"+{fmt(fper[6])} zł", True, GREEN), (" rocznie na aktywnego kupującego (po opłacie premium)", False, TXT)],
    [(f"+{mln(k(RL)[7],1)}", True, GREEN), (" dla rolników razem w 5 lat", False, TXT)]], 13.5)
msg(sl, [("AD to partner logistyczny platformy, a nie konkurent. ", True, WHITE), ("To nas odróżnia od marketplace'ów, które upadły.", False, WHITE)], y=6.3)

# ===== 8. Grupa Azoty =====
gl = [("Lepsza cena netto nawozów", "Lepsza cena netto nawozów: GMV nawozów × 2% (marża AD 7% → 5%)"), ("Lepszy mix", "Poprawa mixu: GMV nawozów × 1%"), ("Udział w rynku AN/CAN", "Udział w rynku AN/CAN: p.p. × 39 mln zł × 25% marży"),
      ("Opłata dla spółki GAia", "Opłata dla spółki GAia: GMV nawozów × 1,2%"), ("SAP/CRM, 5 funkcjonalności", "Integracja SAP/CRM (CAPEX GA) i utrzymanie 5 funkcjonalności"), ("60% gotówki spółki GAia", "60% przepływów spółki GAia (dopłaty na start, potem dywidenda)")]
gres = G["WYNIK GRUPY AZOTY"]
sl = frame("Grupa Azoty: skąd bierze się zysk", f"Wynik GA w roku 5: {mln(gres[6],1)}. W 5 lat: {mln(gres[7],1)}, NPV {mln(k(GA)[8],1)}, zwrot w {k(GA)[10]}.")
pill(sl, 0.6, 1.45, 6.6, 0.42, "Składniki wyniku GA w roku 5, mln zł")
c = chart(sl, XL_CHART_TYPE.BAR_CLUSTERED, 0.6, 1.9, 6.6, 4.2, [a for a, _ in gl][::-1], [("Rok 5", [round(G[b][6], 1) for _, b in gl][::-1])], [NAVY], labels=True, numfmt='#,##0.0')
for i, (_, b) in enumerate(gl[::-1]):
    pt = c.series[0].points[i]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = GREEN if G[b][6] > 0 else RED
pill(sl, 7.5, 1.45, 5.2, 0.42, "Wynik GA rok po roku, mln zł")
chart(sl, XL_CHART_TYPE.COLUMN_CLUSTERED, 7.5, 1.9, 5.2, 4.2, Y05, [("Wynik GA", [round(v, 1) for v in gres[1:7]])], [NAVY], labels=True, numfmt='#,##0.0')
msg(sl, [("Największa wartość to udział w rynku: ", True, WHITE), ("rolnik z planem nawożenia GA kupuje nawóz GA, a nie importowany.", False, WHITE)], y=6.3)

# ===== 9. Każdy zyskuje =====
sl = frame("Każdy zyskuje, także w wariancie ostrożnym", "Wynik w mln zł. Wariant ostrożny: 150 tys. użytkowników w roku 5 zamiast 300 tys.")
data = [["Uczestnik", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma 5 lat", "Zwrot od", "Przy 150 tys.\nużytkowników"]]
for nm, lab in [(GA, "Grupa Azoty"), (TP, "Partner Technologiczny"), (MPk, "Partner Marketplacowy / Handlowy"), (AD, "66 dystrybutorów"), (RL, "Rolnicy"), (JV, "Spółka GAia (gotówka platformy)")]:
    r, rs = k(nm), ks(nm)
    data.append([lab] + [(fmt(v, 1), GREEN if v > 0 else (RED if v < 0 else TXT)) for v in r[1:8]] + [r[10], (mln(rs[7], 1), GREEN if rs[7] > 0 else RED)])
table(sl, 0.6, 1.45, 12.1, data, [3.3, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 1.2, 1.1, 1.7], rh=0.55, size=12.5, hl_rows=(1,))
tb(sl, 0.6, 5.45, 12.1, 0.8, [[("Uczciwie: ", True, NAVY), ("w wariancie ostrożnym sama spółka GAia zwraca się dopiero w roku 5, bo budowa i zespół są stałe. Uczestnicy i tak zarabiają, bo zysk GA i AD powstaje na nawozach, a Partner Technologiczny ma marżę od 1. dnia. Wtedy zmniejszamy zespół rozwoju.", False, TXT)]], 12.5)
msg(sl, [("Każdy zarabia na tym, co wnosi. ", True, WHITE), ("Nikt nie płaci drugiemu za coś, czego ten nie dowiózł.", False, WHITE)], y=6.3)

# ===== 10. Rynek potwierdza =====
sl = frame("Producenci już budują platformy dla rolników", "Publiczne dane firm, stan na 11.2026.")
cases = [
 ("Yara", "Producent nawozów, jak GA", ["ok. 5 mln rolników korzysta z narzędzi cyfrowych Yara", "YaraPlus: platforma nawożenia, start w UK i Irlandii w 01.2026, z John Deere"], "Najbliższy odpowiednik GAia"),
 ("Yara × Agrofy", "Argentyna, marketplace agro", ["Yara prowadziła rundę 30 mln USD (2021), łącznie 66 mln USD", "ok. 150 tys. ofert, płatności i kredyty dla rolników"], "Producent nawozów inwestuje w marketplace"),
 ("BASF xarvio", "Producent ŚOR, aplikacja agronomiczna", [">130 tys. rolników i doradców", ">20 mln ha pól (09.2025)"], "Aplikacja agro skaluje się do setek tysięcy"),
 ("FBN", "USA, marketplace + finansowanie", ["ok. 117-120 tys. gospodarstw członkowskich", ">1 mld USD finansowania zakupów (03.2024)"], "Finansowanie to osobne źródło przychodu"),
]
for i, (name, what, pts, lesson) in enumerate(cases):
    x = 0.6 + i * 3.07
    box(sl, x, 1.45, 2.9, 0.75, fill=NAVY)
    tb(sl, x, 1.45, 2.9, 0.75, [[(name, True, WHITE)], [(what, False, WHITE)]], 13, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    box(sl, x, 2.25, 2.9, 2.5, line=NAVY)
    tb(sl, x + 0.08, 2.35, 2.75, 2.35, ["• " + p for p in pts], 12.5)
    box(sl, x, 4.8, 2.9, 0.75, fill=LIGHT)
    tb(sl, x + 0.05, 4.8, 2.8, 0.75, lesson, 12.5, True, NAVY, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
tb(sl, 0.6, 5.62, 12.1, 0.6, "Źródła: farming.co.uk (YaraPlus, 01.2026); agfundernews.com (Agrofy, 12.2021); basf.com (xarvio, 09.2025); FBN: Wikipedia / Business Wire (2025).", 10, False, GREY)
msg(sl, [("Dla Partnerów: ", True, WHITE), ("w Polsce nikt jeszcze tego nie zrobił. Kto wejdzie z producentem nawozów, będzie pierwszy.", False, WHITE)], y=6.3)

# ===== 11. Lekcje i Polska =====
sl = frame("Czego uczą porażki i dlaczego GAia jest inna")
box(sl, 0.6, 1.2, 3.9, 4.35, line=RED); pill(sl, 0.6, 1.2, 3.9, 0.5, "Agrando, Niemcy", fill=RED)
tb(sl, 0.7, 1.8, 3.7, 3.7, ["• Marketplace nawozów, nasion, części", "• ok. 20 tys. rolników, 12 mln EUR finansowania (2021)", "• Zakończył działalność w 11.2022",
                           [("Lekcja: ", True, RED), ("bez własnych produktów, logistyki i dystrybutorów sam marketplace nie zarobi.", False, TXT)]], 12.5)
box(sl, 4.7, 1.2, 3.9, 4.35, line=AMBER); pill(sl, 4.7, 1.2, 3.9, 0.5, "DeHaat, Indie", fill=AMBER)
tb(sl, 4.8, 1.8, 3.7, 3.7, ["• Platforma: środki produkcji, skup, doradztwo", "• Przychód ok. 2 700 cr INR (FY24), +36-40% r/r", "• Strata operacyjna ok. 245 cr INR (FY24)",
                           [("Lekcja: ", True, AMBER), ("skala przychodzi szybko, ale model z własnym magazynem i skupem długo traci.", False, TXT)]], 12.5)
box(sl, 8.8, 1.2, 3.9, 4.35, line=GREEN); pill(sl, 8.8, 1.2, 3.9, 0.5, "GAia, Polska", fill=GREEN)
tb(sl, 8.9, 1.8, 3.7, 3.7, ["• Producent, produkcja i logistyka w jednej Grupie", "• 66 dystrybutorów jako magazyny i punkty wydań, bez nowych magazynów",
                           "• 5 funkcjonalności już zbudowanych", "• 1,21 mln rolników w ARiMR, brak lidera platform zakupowych",
                           [("Wniosek: ", True, GREEN), ("mamy to, czego zabrakło Agrando, bez kosztów DeHaat.", False, TXT)]], 12.5)
tb(sl, 0.6, 5.62, 12.1, 0.6, "Źródła: agfundernews.com, CB Insights (Agrando); entrackr.com, inc42.com (DeHaat); ARiMR 2024 via wiescirolnicze.pl. Dla Partnera Marketplacowego: ERLI urósł z 1,23 do ok. 1,8 mld zł GMV (2024→2025).", 10, False, GREY)
msg(sl, [("Ryzyko jest w wykonaniu, nie w modelu: ", True, WHITE), ("musimy dowieźć adopcję rolników i AD. Dlatego pilot przed skalowaniem.", False, WHITE)], y=6.3)

# ===== 12. Ryzyka i kroki =====
sl = frame("Ryzyka, warunki umów i dalsze kroki")
data = [["Ryzyko", "Co robimy"],
        ["Za mało użytkowników", "Wariant 150 tys. nadal na plusie. Partner Technologiczny zmniejsza zespół, gdy adopcja jest wolniejsza (w umowie)."],
        ["Konflikt z dystrybutorami", "Zamówienie trafia do lokalnego AD. AD sprzedają też własny asortyment."],
        ["Partner bierze za dużo", "Brak opłat od nawozów GA. Partner Marketplacowy tylko % od sprzedawców. Partner Technologiczny koszt + 15%, audyt kosztów."],
        ["Uzależnienie od technologii", "Kod i dane rolników należą do spółki GAia. Opcja wykupu udziałów Partnera Technologicznego po roku 3."]]
table(sl, 0.6, 1.2, 6.4, data, [2.0, 4.4], rh=0.72, size=11.5)
data = [["Termin", "Krok"], ["do 30.11", "Potwierdzić 5 funkcjonalności i dane: cena netto, marża AD, koszyk rolnika"],
        ["do 15.12", "Term sheet z Partnerem Technologicznym: 60/40, CAPEX w 3 transzach, koszt + 15%, wykup"],
        ["do 15.12", "Term sheet z Partnerem Marketplacowym: 30% od sprzedawców, 0 od nawozów GA"],
        ["01.2027", "Decyzja Zarządu na modelu z prawdziwymi liczbami"],
        ["02-04.2027", "Pilot w sezonie: 1 województwo, 5-10 AD, cel 10 tys. użytkowników"]]
table(sl, 7.2, 1.2, 5.5, data, [1.3, 4.2], rh=0.6, size=11.5)
msg(sl, [("Prosimy o: ", True, WHITE), ("zgodę na term sheety z obydwoma Partnerami i przygotowanie pilota przed sezonem wiosennym 2027.", False, WHITE)], y=6.3)

prs.save("../GA_GAia_business_case_300k.pptx"); print("ok", N[0], "slajdów")
