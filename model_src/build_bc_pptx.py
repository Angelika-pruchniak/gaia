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
    return {r[0]: r for r in book[sheet].iter_rows(values_only=True) if r[0]}
KZ, KZs = rows(wb, "Każdy zyskuje"), rows(st, "Każdy zyskuje")
U = rows(wb, "Użytkownicy i GMV"); J = rows(wb, "Spółka GAia (JV)"); T = rows(wb, "Partner Technologiczny")
MP = rows(wb, "Partner Marketplacowy"); D = rows(wb, "Dystrybutorzy (66)"); F = rows(wb, "Rolnik"); G = rows(wb, "Grupa Azoty")

NAVY = RGBColor(0x00, 0x36, 0x65); TXT = RGBColor(0x1A, 0x1A, 0x1A); GREY = RGBColor(0x59, 0x59, 0x59)
RED = RGBColor(0xE3, 0x06, 0x13); GREEN = RGBColor(0x00, 0x8A, 0x3E); LIGHT = RGBColor(0xEA, 0xF0, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF); LBLUE = RGBColor(0x6F, 0x9F, 0xD8); MBLUE = RGBColor(0x2E, 0x6D, 0xB4)
AMBER = RGBColor(0xE6, 0x7E, 0x00); PALE = RGBColor(0xB4, 0xCB, 0xE8)

prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
def fmt(v, d=0):
    s = f"{v:,.{d}f}".replace(",", " ").replace(".", ",")
    return s.replace("-", "−")
def mln(v, d=0): return f"{fmt(v, d)} mln zł"

def tb(sl, x, y, w, h, text, size=14, bold=False, color=TXT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    s = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = s.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.06); tf.margin_top = tf.margin_bottom = Inches(0.03)
    paras = text if isinstance(text, list) else [text]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        for t, b, c in (p if isinstance(p, list) else [(p, bold, color)]):
            r = para.add_run(); r.text = t; r.font.size = Pt(size); r.font.bold = b; r.font.color.rgb = c; r.font.name = "Calibri"
        para.alignment = align; para.space_after = Pt(3)
    return s
def box(sl, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    s = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
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
    tb(sl, 0.75, y, 11.9, 0.55, [text_runs], 14.5, anchor=MSO_ANCHOR.MIDDLE)
N = [0]
def frame(title, sub=None):
    N[0] += 1
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    tb(sl, 0.6, 0.3, 12.2, 0.7, title, 28, True, NAVY)
    if sub: tb(sl, 0.6, 0.92, 12.2, 0.4, sub, 14, False, GREY)
    box(sl, 0, 6.95, 13.333, 0.55, fill=NAVY, shape=MSO_SHAPE.RECTANGLE)
    sl.shapes.add_picture("logo_ga.png", Inches(12.0), Inches(6.97), height=Inches(0.5))
    tb(sl, 0.6, 7.03, 10.5, 0.4, f"GAia | business case 300 tys. użytkowników | liczby ilustracyjne z modelu GA_GAia_business_case_300k.xlsx, do potwierdzenia   {N[0]}", 10, False, WHITE)
    return sl
def table(sl, x, y, w, data, widths, rh=0.42, size=12, hl_rows=(), first_left=True, hdr_fill=NAVY):
    t = sl.shapes.add_table(len(data), len(data[0]), Inches(x), Inches(y), Inches(w), Inches(rh * len(data))).table
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
                rr = p.add_run(); rr.text = line; rr.font.size = Pt(size); rr.font.name = "Calibri"; rr.font.bold = bold
                rr.font.color.rgb = WHITE if r == 0 else color
    return t
def chart(sl, kind, x, y, w, h, cats, series, colors, legend=True, labels=False, numfmt='#,##0'):
    cd = CategoryChartData(); cd.categories = cats
    for name, vals in series: cd.add_series(name, vals)
    ch = sl.shapes.add_chart(kind, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
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
                          "66 dystrybutorów (AD) razem", "Rolnicy (korzyść)", "Spółka GAia (JV) - wynik platformy")
users = U["Użytkownicy zarejestrowani (tys.)"]; gmv = U["GMV razem"]; fg = U["Nawozy GA kupione przez platformę"]; g3 = U["Inne produkty od sprzedawców 3P (w tym AD)"]
pen = U["Nawozy przez platformę / sprzedaż krajowa GA"]; buy = U["Aktywni kupujący (tys.)"]
rev = J["PRZYCHODY RAZEM"]; jvres = J["WYNIK SPÓŁKI GAia"]

# ===== 1. Teza =====
sl = frame("GAia: 300 tys. rolników w 5 lat. Zarabia każdy, kto wnosi wkład",
           "Grupa Azoty wnosi produkty, produkcję, logistykę, 66 dystrybutorów i 5 gotowych funkcjonalności. Partnerzy wnoszą technologię i handel.")
tiles = [(f"{fmt(users[6])} tys.", "użytkowników w roku 5"), (f"{fmt(gmv[6]/1000, 2)} mld zł", "obrotu na platformie w roku 5"),
         (mln(rev[6]), "przychodu platformy w roku 5"), ("5 z 5", "uczestników na plusie w 5 lat")]
for i, (big, small) in enumerate(tiles):
    x = 0.6 + i * 3.07
    box(sl, x, 1.5, 2.9, 1.25, fill=NAVY)
    tb(sl, x, 1.55, 2.9, 0.7, big, 28, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    tb(sl, x, 2.2, 2.9, 0.45, small, 13, False, WHITE, PP_ALIGN.CENTER)
data = [["Uczestnik", "Wynik w roku 5", "Suma 5 lat (z rokiem 0)", "NPV (10%)", "Zwrot od"]]
for nm, lab in [(GA, "Grupa Azoty"), (TP, "Partner Technologiczny"), (MPk, "Partner Marketplacowy / Handlowy"), (AD, "66 dystrybutorów"), (RL, "Rolnicy (oszczędności)")]:
    r = k(nm); data.append([lab, (mln(r[6], 1), GREEN), (mln(r[7], 1), GREEN), mln(r[8], 1), r[10]])
table(sl, 0.6, 3.0, 12.1, data, [4.0, 2.0, 2.4, 1.9, 1.8], rh=0.5, size=13, hl_rows=(1,))
msg(sl, [("Przy połowie użytkowników (150 tys.) wszyscy nadal są na plusie w 5 lat. ", True, WHITE),
         ("Grupa Azoty: +" + mln(ks(GA)[7]) + ", Partner Technologiczny: +" + mln(ks(TP)[7]) + ", Partner Marketplacowy: +" + mln(ks(MPk)[7]) + ".", False, WHITE)])

# ===== 2. Kto co wnosi =====
sl = frame("Kto co wnosi i jak zarabia", "Partnerzy nazwani po funkcji. Podział spółki GAia: Grupa Azoty 60%, Partner Technologiczny 40%.")
cols = [
 ("Grupa Azoty", NAVY, ["5 gotowych funkcjonalności: plan nawożenia, CRM z profilami gospodarstw, portal AD, Nawozy.eu, koszyk → zamówienie do AD",
                         "66 dystrybutorów zaproszonych na platformę", "Własne produkty, produkcja i logistyka", "Marka i doradcy w regionach"],
  "60% spółki GAia + lepsza cena netto + mix + udział w rynku"),
 ("Partner Technologiczny\n(SoftwareHouse)", MBLUE, ["Silnik platformy agro", "Zespół rozwoju 30-35 osób", "Utrzymanie 24/7, chmura, bezpieczeństwo",
                                                    "Dane satelitarne i pogodowe, asystent AI"],
  "Koszt + 15% marży od 1. dnia + 40% wyniku spółki GAia"),
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

# ===== 3. Skala =====
sl = frame("Ścieżka do 300 tys. użytkowników", f"300 tys. to ok. 25% z 1,21 mln rolników w ARiMR. Aktywnie kupuje {fmt(buy[6])} tys. (35%).")
pill(sl, 0.6, 1.45, 5.6, 0.42, "Użytkownicy na koniec roku, tys.")
chart(sl, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.6, 1.9, 5.6, 3.9, Y15, [("Użytkownicy", [round(v) for v in users[2:7]])], [NAVY], labels=True)
pill(sl, 6.7, 1.45, 6.0, 0.42, "Obrót na platformie (GMV), mln zł")
chart(sl, XL_CHART_TYPE.COLUMN_STACKED, 6.7, 1.9, 6.0, 3.9, Y15,
      [("Nawozy GA", [round(v) for v in fg[2:7]]), ("Inne produkty (sprzedawcy, w tym AD)", [round(v) for v in g3[2:7]])], [NAVY, LBLUE], labels=True)
msg(sl, [("Jak dojdziemy do 300 tys.: ", True, WHITE), ("66 AD zapraszają swoich klientów, doradcy GA w regionach, integracja z mObywatel / ARiMR. ", False, WHITE),
         (f"W roku 5 przez platformę idzie {fmt(pen[6]*100)}% krajowej sprzedaży nawozów GA.", True, WHITE)], y=6.0)

# ===== 4. Spółka GAia =====
sl = frame("Spółka GAia: platforma na plusie od roku 3", f"Przychód w roku 5: {mln(rev[6])}. Zysk w roku 5: {mln(jvres[6])}.")
streams = [("Prowizje od sprzedawców", "Prowizje od sprzedawców 3P", NAVY), ("Reklama", "Reklama", MBLUE), ("Abonamenty premium", "Abonamenty premium", LBLUE),
           ("Finansowanie", "Finansowanie / ubezpieczenia", PALE), ("Dane dla producentów", "Dane i analizy dla producentów", AMBER), ("Opłata GA od nawozów", "Opłata GA od sprzedaży nawozów", GREY)]
pill(sl, 0.6, 1.45, 7.0, 0.42, "Przychody spółki GAia, mln zł")
chart(sl, XL_CHART_TYPE.COLUMN_STACKED, 0.6, 1.9, 7.0, 4.25, Y15, [(n, [round(v, 1) for v in J[key][2:7]]) for n, key, _ in streams], [c for _, _, c in streams])
pill(sl, 7.9, 1.45, 4.8, 0.42, "Na co idzie 100 zł przychodu w roku 5")
parts = [("Partner Technologiczny", J["Partner Technologiczny (koszt + marża)"][6]), ("Partner Marketplacowy", J["Partner Marketplacowy (udział w prowizji i reklamie)"][6]),
         ("Płatności", J["Płatności"][6]), ("Marketing", J["Marketing: pozyskanie użytkowników + marka"][6]),
         ("Obsługa i zespół spółki", J["Obsługa klienta"][6] + J["Zarząd i zespół spółki"][6]), ("Zysk spółki GAia", jvres[6])]
yy = 2.0
for name, v in parts:
    zl = v / rev[6] * 100; last = name.startswith("Zysk")
    tb(sl, 7.9, yy, 2.6, 0.4, name, 13, last, GREEN if last else TXT)
    bw = 1.6 * zl / 55
    box(sl, 10.5, yy + 0.08, max(bw, 0.05), 0.26, fill=GREEN if last else MBLUE, shape=MSO_SHAPE.RECTANGLE)
    tb(sl, 10.5 + bw + 0.05, yy, 0.8, 0.4, f"{fmt(zl)} zł", 13, True, GREEN if last else TXT)
    yy += 0.52
tb(sl, 7.9, yy + 0.1, 4.8, 0.9, f"Wynik spółki: rok 1 {mln(jvres[2],1)}, rok 2 {mln(jvres[3],1)}, rok 3 {mln(jvres[4],1)}. Straty startowe finansują wspólnicy 60/40.", 12, False, GREY)
msg(sl, [("Platforma zarabia na tym, co dobudowuje: ", True, WHITE), ("prowizje, abonamenty, reklama, finansowanie. Opłata GA od nawozów to tylko 10% przychodów.", False, WHITE)], y=6.3)

# ===== 5. Partner Technologiczny =====
cost = T["KOSZTY PARTNERA TECHNOLOGICZNEGO RAZEM"]; cpu = T["Koszt technologii na użytkownika (zł/rok)"]
sl = frame("Partner Technologiczny: ile kosztuje go platforma i ile zarabia",
           f"Koszt po jego stronie w 5 lat: {mln(cost[7])} (z budową). Koszt na użytkownika spada z {fmt(cpu[2])} zł do {fmt(cpu[6])} zł rocznie.")
lines = [("Budowa platformy na silniku Partnera", "Budowa"), ("Zespół rozwoju", "Zespół rozwoju (30-35 FTE)"), ("Utrzymanie 24/7 i wsparcie", "Utrzymanie 24/7, SRE"),
         ("Chmura (stała + na użytkownika)", "Chmura"), ("Dane satelitarne i pogodowe", "Dane satelitarne i pogodowe"),
         ("Asystent AI / modele agronomiczne", "Asystent AI"), ("Bezpieczeństwo i zgodność", "Bezpieczeństwo, RODO, NIS2")]
data = [["Koszty Partnera Technologicznego, mln zł"] + Y05 + ["Razem"]]
for key, lab in lines:
    r = T[key]; data.append([lab] + [fmt(v or 0, 1) for v in r[1:8]])
data.append(["Koszty razem"] + [fmt(v, 1) for v in cost[1:8]])
fee = T["Wynagrodzenie od spółki GAia (koszt + marża)"]; marg = T["Zysk z umowy (marża)"]; share = T["Udział w wyniku spółki GAia"]; tres = T["WYNIK PARTNERA TECHNOLOGICZNEGO (marża + udział)"]
data.append(["Wynagrodzenie od spółki GAia (koszt + 15%)"] + [fmt(v, 1) for v in fee[1:8]])
data.append(["Zysk z umowy (15%)"] + [(fmt(v, 1), GREEN) for v in marg[1:8]])
data.append(["Udział 40% w wyniku spółki GAia"] + [(fmt(v, 1), GREEN if v > 0 else RED) for v in share[1:8]])
data.append(["WYNIK PARTNERA TECHNOLOGICZNEGO"] + [(fmt(v, 1), GREEN if v > 0 else RED) for v in tres[1:8]])
table(sl, 0.6, 1.45, 12.1, data, [4.6] + [1.0] * 6 + [1.5], rh=0.36, size=11.5, hl_rows=(8, 12))
msg(sl, [("Partner nie dokłada do kosztów: ", True, WHITE), (f"umowa koszt + 15% daje mu {mln(marg[7],1)} pewnego zysku. ", False, WHITE),
         (f"Duży zarobek to udział w spółce: razem +{mln(tres[7],1)} w 5 lat, NPV {mln(k(TP)[8],1)}.", True, WHITE)], y=6.3)

# ===== 6. Partner Marketplacowy =====
mrev = MP["Udział w prowizji i reklamie od sprzedawców 3P"]; mres = MP["WYNIK PARTNERA MARKETPLACOWEGO"]; pct = MP["Jaką część przychodów platformy oddajemy Partnerowi"]
sl = frame("Partner Marketplacowy / Handlowy: ile mu oddajemy",
           f"Oddajemy 30% prowizji i reklamy od sprzedawców innych produktów. W roku 5 to {mln(mrev[6],1)}, czyli {fmt(pct[6]*100)}% przychodów platformy.")
pill(sl, 0.6, 1.45, 6.6, 0.42, "Partner Marketplacowy: przychód i wynik, mln zł")
ch = chart(sl, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.6, 1.9, 6.6, 4.1, Y05,
           [("Przychód od spółki GAia", [0] + [round(v, 1) for v in mrev[2:7]]), ("Wynik Partnera po kosztach", [round(v, 1) for v in mres[1:7]])], [LBLUE, GREEN], labels=True, numfmt='#,##0.0')
box(sl, 7.5, 1.45, 5.2, 4.55, line=NAVY)
tb(sl, 7.65, 1.55, 4.95, 4.4, [
    [("Za co płacimy", True, NAVY)],
    [("• Tylko za sprzedaż innych produktów (nasiona, ŚOR, części, usługi), którą Partner pomaga zbudować.", False, TXT)],
    [("• Od nawozów GA nie płacimy nic.", True, TXT)],
    [("Ile to w sumie", True, NAVY)],
    [(f"• Przychód Partnera w 5 lat: {mln(mrev[7],1)}", False, TXT)],
    [(f"• Koszt po jego stronie (integracja, zespół, ruch): {mln(mrev[7]-mres[7],1)}", False, TXT)],
    [(f"• Wynik Partnera w 5 lat: +{mln(mres[7],1)}, zwrot w {k(MPk)[10]}", True, GREEN)],
    [("Dlaczego Partner wchodzi", True, NAVY)],
    [("• Nowa kategoria agro bez inwestycji w towar i magazyn. Rolnicy przychodzą po nawozy GA, a kupują też u jego sprzedawców.", False, TXT)]], 12.5)
msg(sl, [("Im więcej sprzedawców i zakupów dowiezie Partner, tym więcej zarobi. ", True, WHITE), ("Nie ma opłaty stałej, więc ryzyko GA jest małe.", False, WHITE)], y=6.3)

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
gl = [("Lepsza cena netto nawozów", "Lepsza cena netto nawozów (marża AD 7% → 5%)"), ("Lepszy mix", "Poprawa mixu"), ("Udział w rynku AN/CAN", "Marża z dodatkowego udziału AN/CAN"),
      ("Opłata dla spółki GAia", "Opłata dla spółki GAia od nawozów"), ("SAP/CRM, 5 funkcjonalności, zespół", "Integracja SAP/CRM, utrzymanie 5 funkcjonalności, zespół"), ("60% wyniku spółki GAia", "Udział GA w wyniku spółki GAia (60%)")]
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
for nm, lab in [(GA, "Grupa Azoty"), (TP, "Partner Technologiczny"), (MPk, "Partner Marketplacowy / Handlowy"), (AD, "66 dystrybutorów"), (RL, "Rolnicy"), (JV, "Spółka GAia (sama platforma)")]:
    r, rs = k(nm), ks(nm)
    data.append([lab] + [(fmt(v, 1), GREEN if v > 0 else (RED if v < 0 else TXT)) for v in r[1:8]] + [r[10], (mln(rs[7], 1), GREEN if rs[7] > 0 else RED)])
table(sl, 0.6, 1.45, 12.1, data, [3.3, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 1.2, 1.1, 1.7], rh=0.55, size=12.5, hl_rows=(1,))
tb(sl, 0.6, 5.45, 12.1, 0.8, [[("Uczciwie: ", True, NAVY), ("w wariancie ostrożnym sama spółka GAia nie zwraca się w 5 lat, bo zespół rozwoju jest stały. Uczestnicy i tak zarabiają, bo zysk GA i AD powstaje na nawozach, a Partner Technologiczny ma marżę od 1. dnia. Wtedy zmniejszamy zespół rozwoju.", False, TXT)]], 12.5)
msg(sl, [("Każdy zarabia na tym, co wnosi. ", True, WHITE), ("Nikt nie płaci drugiemu za coś, czego ten nie dowiózł.", False, WHITE)], y=6.3)

# ===== 10. Rynek potwierdza =====
sl = frame("Rynek potwierdza: producenci budują platformy dla rolników", "Publiczne dane firm, stan na 11.2026.")
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
        ["do 15.12", "Term sheet z Partnerem Technologicznym: 60/40, koszt + 15%, wykup"],
        ["do 15.12", "Term sheet z Partnerem Marketplacowym: 30% od sprzedawców, 0 od nawozów GA"],
        ["01.2027", "Decyzja Zarządu na modelu z prawdziwymi liczbami"],
        ["02-04.2027", "Pilot w sezonie: 1 województwo, 5-10 AD, cel 10 tys. użytkowników"]]
table(sl, 7.2, 1.2, 5.5, data, [1.3, 4.2], rh=0.6, size=11.5)
msg(sl, [("Prosimy o: ", True, WHITE), ("zgodę na term sheety z obydwoma Partnerami i przygotowanie pilota przed sezonem wiosennym 2027.", False, WHITE)], y=6.3)

prs.save("../GA_GAia_business_case_300k.pptx"); print("ok", N[0], "slajdów")
