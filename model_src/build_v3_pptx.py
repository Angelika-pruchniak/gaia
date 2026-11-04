# Prezentacja: rekomendacja dyrektora strategii - wspólna pula przychodów GAia. Liczby z GA_GAia_model_wspolnej_puli.xlsx.
# Użycie: python3 build_jr_pptx.py <przeliczony.xlsx>
import sys, openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION

wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
def rows(book, sheet):
    d = {}
    for r in book[sheet].iter_rows(values_only=True):
        if r[0] and r[0] not in d: d[r[0]] = r
    return d
KZ = rows(wb, "Każdy partner zyskuje"); U = rows(wb, "Rynek i GMV"); J = rows(wb, "Wspólna pula"); T = rows(wb, "Partner Technologiczny")
MP = rows(wb, "Partner Marketplace"); G = rows(wb, "Grupa Azoty"); SC = rows(wb, "Scenariusze"); KP = rows(wb, "Koszty platformy")

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
    tb(sl, 0.66, 6.95, 10.8, 0.35, f"GAia | model współpracy v3 | scenariusz bazowy, 7 lat, liczby do potwierdzenia   {N[0]}", 9, False, WHITE)
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

YR = [f"Rok {i}" for i in range(8)]
v = lambda r, i: r[i] if r[i] is not None else 0
def yr(lab, r, d=1, color=None, tot=True, rng=range(1, 9)):
    out = [lab] + [(fmt(v(r, i), d), color(v(r, i)) if color else TXT) for i in rng]
    if tot: out.append(fmt(v(r, 9), d))
    return out
sign = lambda x: GREEN if x > 0.05 else (RED if x < -0.05 else TXT)
sc = lambda lab: SC[lab][1:4]
names3 = ["Ostrożny", "Bazowy", "Agresywny\n(model GA)"]
pool = J["PRZYCHODY PULI RAZEM"]; cost = J["Koszty wspólne puli"]; net = J["WYNIK PULI (EBITDA platformy)"]; dist = J["DO PODZIAŁU WEDŁUG KLUCZA"]
fund = J["Deficyt (finansują PT 70% i PM 30%)"]; rep = J["Zwrot deficytu (pierwszeństwo)"]; cumn = J["Wynik puli narastająco"]
users = U["Użytkownicy (średnio w roku, tys.)"]; gmv = U["GMV razem"]
mvp_rows = [r for r in wb["Dopasowanie MVP1"].iter_rows(min_row=5, max_row=36, values_only=True) if r[0]]
mvp_pt = sum(r[5] for r in mvp_rows if r[2] == "Partner Technologiczny"); mvp_pm = sum(r[5] for r in mvp_rows if r[2] == "Partner Marketplace")
mvp_tot = mvp_pt + mvp_pm
YH = ["mln zł"] + YR + ["Suma"]; YW = [3.5] + [0.92] * 8 + [1.06]

# ===== 1. Rekomendacja =====
sl = frame("Rekomendacja: długoterminowy alians GAia", "GA zaprasza partnerów technologicznych z gotową platformą do podziału zysków. GA nie płaci za platformę.")
tiles = [("0 zł", "gotówki od GA"), (f"{fmt(mvp_tot*0.02,1)} mln zł", f"pracy partnerów w roku 0 ({fmt(mvp_tot)} osobomies.)"),
         (f"{fmt(sc('Oszczędność dzięki partnerstwu (7 lat)')[1])} mln zł", "taniej niż budowa samodzielna (7 lat)"), ("10 + 5 lat", "okres współpracy")]
for i, (big, small) in enumerate(tiles):
    x = 0.6 + i * 3.07
    box(sl, x, 1.4, 2.9, 1.15, fill=GREEN if i == 0 else NAVY)
    tb(sl, x, 1.45, 2.9, 0.65, big, 24, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE); tb(sl, x, 2.08, 2.9, 0.42, small, 11.5, False, WHITE, PP_ALIGN.CENTER)
data = [["Wynik w 7 lat (gotówka, mln zł)"] + names3,
        ["Grupa Azoty"] + [(fmt(x, 1), sign(x)) for x in sc("Grupa Azoty: suma 7 lat")],
        ["Partner Technologiczny (SH)"] + [(fmt(x, 1), sign(x)) for x in sc("Partner Technologiczny: suma 7 lat")],
        ["Partner Marketplace / Handel"] + [(fmt(x, 1), sign(x)) for x in sc("Partner Marketplace: suma 7 lat")],
        ["Użytkownicy w roku 5, tys."] + [fmt(x) for x in sc("Użytkownicy w roku 5")]]
table(sl, 0.6, 2.75, 12.1, data, [4.6, 2.5, 2.5, 2.5], rh=0.55, size=12.5)
tb(sl, 0.6, 5.55, 12.1, 0.6, [[("Etap 1 (lata 0-2): ", True, NAVY), ("umowa partnerska ze wspólną pulą. ", False, TXT), ("Etap 2: ", True, NAVY), ("spółka GAia po bramce KPI. W scenariuszu ostrożnym bramka zatrzymuje projekt, zanim partnerzy stracą więcej.", False, TXT)]], 12.5)
msg(sl, [("Prosimy o: ", True, WHITE), ("zgodę na term sheet z SH i Partnerem Marketplace na tych zasadach oraz pilot MVP1 w sezonie wiosennym.", False, WHITE)])

# ===== 2. Co z modelu GA =====
sl = frame("Model GA: co przejęliśmy i co zmieniliśmy")
UWr = [r for r in wb["Uwagi do modelu GA"].iter_rows(min_row=5, max_row=16, values_only=True)]
pick = ["Adopcja", "Udział w wydatkach rolnika", "Rozwój (capex)", "Back margin 2%", "Prowizja od nawozów", "Premium", "Korzyść GA", "Zespół i utrzymanie", "CAC i antichurn", "Kto ponosi koszty"]
data = [["Element", "Model GA", "Model v3", "Dlaczego"]] + [[r[0], r[1], r[2], r[3]] for r in UWr if r[0] in pick]
table(sl, 0.6, 1.2, 12.1, data, [2.3, 3.2, 3.4, 3.2], rh=0.47, size=9.5)
msg(sl, [("Scenariusz agresywny = model GA 1:1 ", True, WHITE), ("(adopcja, udział w wydatkach, prowizje, koszty). Dodaliśmy bazowy i ostrożny.", False, WHITE)], y=6.3)

# ===== 3. Model współpracy =====
sl = frame("Model współpracy: role, pula, zasady", "Wzorzec: Bushel (white-label dla sieci), xFarm × Yara (wiedza w silniku partnera), Bosch × BASF (współpraca 2019 → spółka 2021)")
for i, (nm, col, txt) in enumerate([("Grupa Azoty", NAVY, "produkty, logistyka, dystrybutorzy, marka, nawozy.eu, ARiMR, 5 funkcjonalności"),
                                     ("Partner Technologiczny (SH)", MBLUE, "silnik na licencji, UX, grywalizacja, integracje, utrzymanie"),
                                     ("Partner Marketplace", LBLUE, "marketplace przez plug-in, sprzedawcy, płatności, operacje")]):
    x = 0.6 + i * 4.1
    box(sl, x, 1.25, 3.9, 1.05, fill=col); tb(sl, x, 1.25, 3.9, 1.05, [[(nm, True, WHITE)], [(txt, False, WHITE)]], 11.5, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
box(sl, 3.6, 2.45, 6.1, 0.8, fill=GREEN)
tb(sl, 3.6, 2.45, 6.1, 0.8, [[("GAia: wspólna pula przychodów", True, WHITE)], [("prowizje + back margin + premium + 25% dodatkowej marży GA", False, WHITE)]], 12.5, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
data = [["Zasada", "Jak"],
        ["Okres", "10 lat + automatycznie 5; wypowiedzenie najwcześniej po 5 latach, 12 mies. przejściowo"],
        ["Rozliczenie", "Koszty wspólne po koszcie (budżet roczny) → zwrot deficytu partnerom → podział według klucza"],
        ["GA nie płaci", "Bez gotówki i bez dopłat; tylko 25% dodatkowej marży z platformy. Prowizje płacą sprzedawcy i dystrybutorzy"],
        ["Zarządzanie", "Komitet Sterujący 2 + 2 + 2, przewodniczy GA; sprawy zastrzeżone jednomyślnie; wspólne biuro projektu"],
        ["IP i dane", "Silniki i funkcjonalności zostają u właścicieli (licencja, depozyt kodu); nowy kod i dane GAia w spółce; zgody RODO rolnika"],
        ["Wyłączność i wyjście", "Bez konkurencyjnej platformy dla innego producenta nawozów w PL; pierwokup udziałów; opcja wykupu po 5 latach"]]
table(sl, 0.6, 3.4, 12.1, data, [2.4, 9.7], rh=0.37, size=10.5)

# ===== 4. Wkład =====
sl = frame("Kto co wnosi i co dostaje")
cols = [("Grupa Azoty", NAVY, ["5 funkcjonalności na Azure (audyt, bez wyceny)", "Produkty, produkcja, logistyka", "66 dystrybutorów i punkty wydania", "Marketing i rozpoznawalność marki", "nawozy.eu, ARiMR, Węzeł Krajowy"],
         ["75% dodatkowej marży", "20% prowizji, 30% back marginu, 15% premium", "0 zł gotówki"]),
        ("Partner Technologiczny (SH)", MBLUE, ["Gotowy silnik (licencja)", f"Rok 0: MVP1 z OPZ ({fmt(mvp_pt)} osobomies.)", "Aplikacja mobilna i grywalizacja", "Integracje: ARiMR, Węzeł, SAP, banki", "Utrzymanie i bezpieczeństwo"],
         ["Zwrot kosztów bieżących z puli", "40% prowizji, 70% premium, 60% opłaty za sukces", "Finansuje 70% deficytu, który wraca pierwszy"]),
        ("Partner Marketplace / Handel", LBLUE, ["Gotowy marketplace (plug-in)", f"Rok 0: katalog, zapytania, oferty ({fmt(mvp_pm)} osobomies.)", "Sprzedawcy i płatności", "Kategorie i akcje promocyjne", "Operacje marketplace'u"],
         ["Zwrot kosztów bieżących z puli", "40% prowizji, 50% back marginu, 40% opłaty za sukces", "Finansuje 30% deficytu, który wraca pierwszy"])]
for i, (name, colr, gives, earns) in enumerate(cols):
    x = 0.6 + i * 4.1
    box(sl, x, 1.2, 3.9, 0.55, fill=colr); tb(sl, x, 1.2, 3.9, 0.55, name, 14, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    box(sl, x, 1.8, 3.9, 2.35, line=colr); tb(sl, x + 0.1, 1.85, 3.7, 0.3, "WNOSI", 11, True, GREY); tb(sl, x + 0.1, 2.12, 3.7, 2.0, ["• " + t for t in gives], 12)
    box(sl, x, 4.22, 3.9, 1.6, fill=LIGHT); tb(sl, x + 0.1, 4.27, 3.7, 0.3, "DOSTAJE", 11, True, NAVY); tb(sl, x + 0.1, 4.54, 3.7, 1.25, ["• " + t for t in earns], 12)
msg(sl, [("Każdy wnosi to, co już ma i za co już płaci. ", True, WHITE), ("Dlatego wspólna platforma jest tańsza niż budowa przez jedną firmę.", False, WHITE)], y=6.3)

# ===== 5. Dopasowanie MVP1 =====
sl = frame("Dopasowanie MVP1 (OPZ) do platform partnerów", f"{fmt(mvp_tot)} osobomiesięcy × 20 tys. zł = {fmt(mvp_tot*0.02,1)} mln zł (model GA: 4 mln zł). Do potwierdzenia odpowiedziami SH na PZ-01…49.")
groups = [("Konto, role, bezpieczeństwo, administracja", ["WB-01", "WB-02", "WB-03", "WB-26", "WB-27"], ["Bezpieczeństwo"]),
          ("Kokpit, profil, dokumenty", ["WB-04", "WB-05", "WB-06", "WB-09"], []), ("Paszport pola, ARiMR, planer, kalkulator, dotacje", ["WB-07", "WB-08", "WB-11", "WB-12", "WB-13"], []),
          ("Alert 6:30 i powiadomienia", ["WB-10"], []), ("Baza wiedzy (nawozy.eu) i Bot Norm", ["WB-14", "WB-15"], []),
          ("Zakupy: katalog, zapytania, oferty, leady", ["WB-16", "WB-17", "WB-18", "WB-19"], []), ("Dystrybutor: regiony i karta klienta", ["WB-20", "WB-21"], []),
          ("Finansowanie: wniosek, zgody, oferta, dostęp", ["WB-22", "WB-23", "WB-24", "WB-25"], []), ("Raportowanie", ["WB-28"], []),
          ("Grywalizacja", [], ["Grywalizacja"]), ("Aplikacja mobilna", [], ["Aplikacja mobilna"]), ("Testy, kierowanie projektem, wdrożenie", [], ["Testy"])]
data = [["Obszar MVP1", "Wymagania", "Kto dostarcza", "Dopasowanie", "GA wnosi", "Osobo-\nmies."]]
for nm, ids, kw in groups:
    rs = [r for r in mvp_rows if r[0] in ids or any(r[1].startswith(k) for k in kw)]
    who = sorted(set("PT" if r[2].startswith("Partner T") else "PM" for r in rs)); fit = sorted(set(r[3] for r in rs))
    ga_ = sorted(set(r[4] for r in rs if r[4])); data.append([nm, ", ".join(ids) if ids else "+", " + ".join(who), "; ".join(fit)[:60], ", ".join(ga_)[:44], fmt(sum(r[5] for r in rs))])
data.append(["RAZEM", "", "", "", "", fmt(mvp_tot)])
table(sl, 0.6, 1.2, 12.1, data, [3.2, 2.1, 0.9, 3.0, 2.2, 0.7], rh=0.3, size=9.5, hl_rows=(13,))
msg(sl, [("Najważniejsze: ", True, WHITE), ("MVP1 to zapytanie ofertowe do dystrybutora, nie pełny koszyk. Prowizję w roku 1 płaci dystrybutor od zrealizowanego zapytania.", False, WHITE)], y=6.3)

# ===== 6. Synergie =====
sl = frame("Partnerstwo jest tańsze niż budowa samodzielna", f"Te same koszty co w modelu GA, ale z zasobami, które partnerzy już mają. Oszczędność w 7 lat: {fmt(KP['Oszczędność dzięki partnerstwu'][9],1)} mln zł (bazowy).")
chart(sl, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.6, 1.2, 6.3, 3.6, YR[1:], [("Model GA: jedna firma sama", [round(v(KP["KOSZTY RAZEM (model samodzielny)"], i), 1) for i in range(2, 9)]),
      ("Partnerstwo", [round(v(KP["KOSZTY WSPÓLNE PULI (w partnerstwie)"], i), 1) for i in range(2, 9)])], [GREY, GREEN], labels=False)
data = [["Synergia", "Skąd", "Efekt"], ["Zespół −30%", "Istniejące zespoły PT i PM; dystrybutorzy GA odpowiadają na zapytania", fmt(KP["Zespół (IT, produkt, obsługa, zakupy)"][9] - KP["Zespół (PT 50% / PM 30% / obsługa 20%)"][9], 1) + " mln zł"],
        ["Utrzymanie −25%", "Platforma PT współdzielona z innymi klientami", fmt(KP["Utrzymanie (usługi obce, SaaS)"][9] - KP["Utrzymanie (PT 80% / PM 20%)"][9], 1) + " mln zł"],
        ["Pozyskanie −40%", "66 dystrybutorów, KAM, nawozy.eu, ARiMR, marka GA", fmt(KP["Pozyskanie użytkowników (CAC)"][9] + KP["Antichurn"][9] - KP["Pozyskanie i utrzymanie użytkowników"][9], 1) + " mln zł"],
        ["Rok 0", f"Gotowe silniki: {fmt(mvp_tot)} osobomies. zamiast budowy od zera", f"{fmt(mvp_tot*0.02,1)} mln zł pracy"]]
table(sl, 7.1, 1.2, 5.6, data, [1.4, 3.0, 1.2], rh=0.75, size=10)
msg(sl, [("To jest powód współpracy: ", True, WHITE), ("sama platforma w scenariuszu bazowym długo przynosi straty, a w partnerstwie zarabia od roku 4.", False, WHITE)], y=6.3)

# ===== 7. Pula i klucz =====
sl = frame("Jak dzielimy przychody: wspólna pula i klucz")
steps = [("1", "Przychody do puli", "prowizje (płacą sprzedawcy i dystrybutorzy), back margin, premium, 25% dodatkowej marży GA", NAVY),
         ("2", "Koszty po koszcie", "rozwój, utrzymanie, zespoły, pozyskanie rolników; budżet roczny, bez marży", MBLUE),
         ("3", "Zwrot deficytu", "PT 70% i PM 30% finansują start; dostają to z powrotem pierwsi", AMBER),
         ("4", "Podział według klucza", "każdy strumień osobno, według tego, kto tworzy wartość", GREEN)]
for i, (n, t, d, col) in enumerate(steps):
    x = 0.6 + i * 3.07
    box(sl, x, 1.2, 2.9, 2.05, fill=col); tb(sl, x, 1.22, 2.9, 0.5, n, 24, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    tb(sl, x + 0.1, 1.72, 2.7, 0.45, t, 13, True, WHITE, PP_ALIGN.CENTER); tb(sl, x + 0.1, 2.15, 2.7, 1.05, d, 11, False, WHITE, PP_ALIGN.CENTER)
Kk = [r for r in wb["Klucz podziału"].iter_rows(min_row=5, max_row=8, values_only=True)]
st5 = [J["Prowizje od sprzedawców i dystrybutorów"][6], J["Back margin od dostawców"][6], J["Usługi premium"][6], J["Opłata za sukces od GA"][6]]
data = [["Strumień", "Rok 5, mln zł", "GA", "PT", "PM", "Dlaczego"]]
for r, val in zip(Kk, st5): data.append([r[0], fmt(val, 1), f"{fmt(r[1]*100)}%", f"{fmt(r[2]*100)}%", f"{fmt(r[3]*100)}%", r[5]])
table(sl, 0.6, 3.45, 12.1, data, [3.4, 1.1, 0.7, 0.7, 0.7, 5.5], rh=0.5, size=10.5)
msg(sl, [("Nikt nie zarabia na kosztach, wszyscy na wyniku. ", True, WHITE), ("Klucz przeglądamy co 2 lata według rzeczywistego wkładu.", False, WHITE)], y=6.3)

# ===== 8. Scenariusze =====
sl = frame("Trzy scenariusze, 7 lat")
rowsS = [("Użytkownicy w roku 5, tys.", "Użytkownicy w roku 5", 0), ("GMV w roku 5, mln zł", "GMV w roku 5", 0), ("Przychody puli w roku 5", "Przychody puli w roku 5", 1),
         ("EBITDA platformy w roku 5", "EBITDA platformy w roku 5", 1), ("Najniższy wynik puli narastająco", "EBITDA platformy narastająco: najniżej", 1),
         ("Grupa Azoty: suma 7 lat", "Grupa Azoty: suma 7 lat", 1), ("Partner Technologiczny: suma 7 lat", "Partner Technologiczny: suma 7 lat", 1),
         ("Partner Technologiczny: zwrot od", "Partner Technologiczny: zwrot od", -2), ("Partner Marketplace: suma 7 lat", "Partner Marketplace: suma 7 lat", 1),
         ("Partner Marketplace: zwrot od", "Partner Marketplace: zwrot od", -2), ("PT: wynik przy zatrzymaniu na bramce 2", "PT: wynik po roku 2 (strata przy zatrzymaniu na bramce)", 1),
         ("PM: wynik przy zatrzymaniu na bramce 2", "PM: wynik po roku 2 (strata przy zatrzymaniu na bramce)", 1)]
data = [["mln zł", "Ostrożny", "Bazowy", "Agresywny (GA)"]]
for lab, key, d in rowsS:
    vals = sc(key)
    data.append([lab] + [(x, TXT) if d == -2 else (fmt(x, max(d, 0)), sign(x) if ("suma" in lab or "bramce" in lab) else TXT) for x in vals])
table(sl, 0.6, 1.2, 7.6, data, [3.8, 1.25, 1.25, 1.3], rh=0.37, size=10.5, hl_rows=(6,))
chart(sl, XL_CHART_TYPE.COLUMN_CLUSTERED, 8.4, 1.2, 4.3, 4.7, ["Ostrożny", "Bazowy", "Agresywny"],
      [("Grupa Azoty", [round(x, 1) for x in sc("Grupa Azoty: suma 7 lat")]), ("Partner Technologiczny", [round(x, 1) for x in sc("Partner Technologiczny: suma 7 lat")]),
       ("Partner Marketplace", [round(x, 1) for x in sc("Partner Marketplace: suma 7 lat")])], [NAVY, MBLUE, LBLUE], labels=True, numfmt='#,##0')
msg(sl, [("Uczciwie: ", True, WHITE), ("w ostrożnym partnerzy technologiczni nie wychodzą na plus w 7 lat. Wtedy bramka 2 (7% rolników) nie jest spełniona i zatrzymujemy lub zmieniamy zakres.", False, WHITE)], y=6.3)

# ===== 9. Pula rok po roku =====
sl = frame("Wspólna pula rok po roku (bazowy)", f"Przychody puli w roku 5: {fmt(pool[6],1)} mln zł ({fmt(J['Przychody puli / GMV'][6]*100,1)}% GMV; model GA: 5,3%). EBITDA na plusie od roku 4.")
lines = [("Prowizje (sprzedawcy, dystrybutorzy)", "Prowizje od sprzedawców i dystrybutorów"), ("Back margin od dostawców", "Back margin od dostawców"), ("Usługi premium", "Usługi premium"), ("Opłata za sukces od GA", "Opłata za sukces od GA")]
data = [YH] + [yr(lab, J[key]) for lab, key in lines] + [yr("PRZYCHODY PULI", pool)]
for lab, key in [("Rozwój platformy", "Rozwój platformy (PT)"), ("Utrzymanie", "Utrzymanie (PT 80% / PM 20%)"), ("Zespół", "Zespół (PT 50% / PM 30% / obsługa 20%)"), ("Pozyskanie i utrzymanie rolników", "Pozyskanie i utrzymanie użytkowników")]:
    data.append(yr(lab, KP[key]))
data += [yr("KOSZTY WSPÓLNE", cost), yr("EBITDA PLATFORMY", net, color=sign), yr("Deficyt (PT 70% / PM 30%)", fund), yr("Zwrot deficytu", rep), yr("Do podziału według klucza", dist, color=sign)]
table(sl, 0.6, 1.2, 12.1, data, YW, rh=0.3, size=9.5, hl_rows=(5, 10, 11))

# ===== 10/11. Partnerzy =====
def pslide(Pd, title, nm, keyfin, pct):
    own, fin, dst, res_, cumr = Pd["0. Wkład własny (rok 0)"], Pd["1. Dopłaty do deficytu puli i ich zwrot"], Pd["2. Udział w podziale puli"], Pd["WYNIK PARTNERA"], Pd["Wynik skumulowany"]
    sl = frame(title, f"Bazowy: +{fmt(res_[9],1)} mln zł w 7 lat, zwrot w {KZ[nm][11]}. Agresywny (model GA): +{fmt(sc(keyfin+': suma 7 lat')[2],1)} mln zł.")
    chart(sl, XL_CHART_TYPE.COLUMN_STACKED, 0.6, 1.2, 6.6, 3.3, YR,
          [("Wkład w roku 0", [round(v(own, i), 1) for i in range(1, 9)]), ("Deficyt i jego zwrot", [round(v(fin, i), 1) for i in range(1, 9)]), ("Udział w podziale", [round(v(dst, i), 1) for i in range(1, 9)])], [RED, AMBER, GREEN])
    data = [YH, yr("0. Wkład własny (MVP1, rok 0)", own, color=sign), yr(f"1. Deficyt puli ({pct}) i zwrot", fin, color=sign), yr("2. Udział w podziale puli", dst, color=sign),
            yr("WYNIK PARTNERA", res_, color=sign), yr("Skumulowany", cumr, color=sign, tot=False) + [""]]
    table(sl, 0.6, 4.6, 12.1, data, YW, rh=0.27, size=9.5, hl_rows=(4,))
    box(sl, 7.5, 1.2, 5.2, 3.3, line=NAVY)
    tb(sl, 7.6, 1.28, 5.0, 3.2, [[("Rok 0: wkład", True, RED)], [(f"Praca przy MVP1: {fmt(-own[1],1)} mln zł.", False, TXT)],
        [("Lata 1-3: finansuje start", True, AMBER)], [(f"{pct} deficytu puli; koszty bieżące pokrywa pula.", False, TXT)],
        [("Lata 4-5: zwrot", True, MBLUE)], [("Zysk puli najpierw zwraca deficyt.", False, TXT)],
        [("Od roku 5: zarabia", True, GREEN)], [(f"Udział w podziale: {fmt(dst[9],1)} mln zł w 7 lat.", False, TXT)],
        [("Przy zatrzymaniu na bramce 2: ", True, NAVY), (f"{fmt(KZ[nm][12],1)} mln zł", True, RED)]], 11.5)
pslide(T, "Partner Technologiczny (SH): kiedy zarabia", "Partner Technologiczny", "Partner Technologiczny", "70%")
pslide(MP, "Partner Marketplace: kiedy zarabia", "Partner Marketplace / Handel", "Partner Marketplace", "30%")

# ===== 12. GA =====
gres = G["WYNIK GRUPY AZOTY"]
sl = frame("Grupa Azoty: zysk bez gotówki", f"Bazowy: +{fmt(gres[9],1)} mln zł w 7 lat, na plusie od roku 2. Agresywny: +{fmt(sc('Grupa Azoty: suma 7 lat')[2],1)} mln zł.")
data = [YH] + [yr(lab, G[key]) for lab, key in [("Marża z danych rolników (model GA)", "Marża z danych rolników (model GA)"), ("Marża z nowej sprzedaży (udział w rynku)", "Marża z nowej sprzedaży (udział w rynku)"), ("Dodatkowa marża GA", "DODATKOWA MARŻA GA")]]
data.append(yr("Opłata za sukces do puli (25%)", [None] + [-(x or 0) for x in J["Opłata za sukces od GA"][1:10]], color=sign))
data += [yr("Udział w podziale puli", G["2. Udział w podziale puli"], color=sign), yr("Koszt wewnętrzny GA", G["3. Koszt wewnętrzny GA"], color=sign), yr("Wkład pieniężny GA", G["Wkład pieniężny GA"]), yr("WYNIK GRUPY AZOTY", gres, color=sign)]
table(sl, 0.6, 1.2, 12.1, data, YW, rh=0.36, size=10, hl_rows=(3, 8))
msg(sl, [("Jak Nutrien i Syngenta: ", True, WHITE), ("GA zarabia na sprzedaży własnych nawozów. Partnerom oddaje 25% dodatkowego zysku i część przychodów platformy.", False, WHITE)], y=6.3)

# ===== 13. Ryzyka i bramki =====
sl = frame("Bramki decyzyjne i ograniczenie ryzyka")
gates = [("Bramka 0, koniec roku 0", "MVP1 zgodne z OPZ, audyt 5 funkcjonalności GA, pentest", NAVY), ("Bramka 1, wiosna roku 1", "Pilot: ≥ 10 tys. użytkowników, ≥ 5 tys. zapytań, ocena ≥ 4/5", MBLUE),
         ("Bramka 2, koniec roku 2", "≥ 7% rolników, pula ≥ 10 mln zł/rok, EBITDA ≥ 0 lub plan dojścia → spółka GAia", GREEN)]
for i, (t, d, col) in enumerate(gates):
    x = 0.6 + i * 4.1
    box(sl, x, 1.2, 3.9, 0.55, fill=col); tb(sl, x, 1.2, 3.9, 0.55, t, 13, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    box(sl, x, 1.8, 3.9, 0.95, fill=LIGHT); tb(sl, x + 0.1, 1.8, 3.7, 0.95, d, 11.5, True, NAVY, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
data = [["Ryzyko", "Co robimy"],
        ["Wolniejsza adopcja (ostrożny)", "Bramka 2 zatrzymuje projekt; maks. strata PT " + fmt(sc("PT: wynik po roku 2 (strata przy zatrzymaniu na bramce)")[0], 1) + ", PM " + fmt(sc("PM: wynik po roku 2 (strata przy zatrzymaniu na bramce)")[0], 1) + " mln zł; GA 0"],
        ["Duże finansowanie startu przez partnerów", "Limit deficytu w budżecie; granty (FENG SMART, Horizon Europe, EIT Food) na model AI i pilot; bank jako partner finansowy (jak Oxbury u Yary)"],
        ["Partner zawyża koszty", "Koszty po koszcie, budżet roczny, audyt; nikt nie ma marży na kosztach"],
        ["Spór o podział", "Klucz w umowie, przegląd co 2 lata, eskalacja i mediacja"],
        ["Utrata kontroli nad danymi", "Zgody RODO rolnika; dane GAia w spółce; silniki i funkcjonalności u właścicieli; depozyt kodu"]]
table(sl, 0.6, 2.95, 12.1, data, [3.6, 8.5], rh=0.5, size=11)
msg(sl, [("GA ryzykuje tylko czas i dane. ", True, WHITE), ("Partnerzy ryzykują ograniczoną kwotę, bo każda bramka pozwala zatrzymać projekt.", False, WHITE)], y=6.3)

# ===== 14. Case study: rynek + technologia =====
sl = frame("Case study: firma z rynkiem + partner z technologią", "Nikt nie wykładał milionów: każdy wniósł to, co już miał")
cases = [("Bushel (USA)", "startup + spółdzielnie rolnicze", ["Aplikacja white-label pod marką spółdzielni, darmowa dla rolnika", "10 punktów na start → ok. 1 500 punktów skupu", "Przejął GrainBridge (JV ADM i Cargill)"], "Aplikacja pod marką GA dla sieci 66 dystrybutorów"),
         ("xFarm × Yara (2023)", "startup z platformą + producent nawozów", ["Wiedza Yary o nawożeniu wbudowana w moduł xFarm", "xFarm: 300 tys. gospodarstw, 3 mln ha", "Bez nowej spółki i bez dużej inwestycji"], "5 funkcjonalności GA i nawozy.eu w silniku SH"),
         ("Żabka × AiFi (Polska)", "korporacja + startup technologiczny", ["Żabka: koncept i sprzedaż (Żabka Future), AiFi: technologia", "25 sklepów Nano w 6 miesięcy", "Ponad 50 sklepów w 2023"], "GA daje rynek i klientów, partner technologię")]
for i, (nm, what, pts, lesson) in enumerate(cases):
    x = 0.6 + i * 4.1
    box(sl, x, 1.2, 3.9, 0.8, fill=NAVY); tb(sl, x, 1.2, 3.9, 0.8, [[(nm, True, WHITE)], [(what, False, WHITE)]], 13, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    box(sl, x, 2.05, 3.9, 2.35, line=NAVY); tb(sl, x + 0.1, 2.15, 3.7, 2.2, ["• " + p for p in pts], 12.5)
    box(sl, x, 4.45, 3.9, 0.95, fill=LIGHT); tb(sl, x + 0.05, 4.45, 3.8, 0.95, lesson, 12.5, True, NAVY, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
tb(sl, 0.6, 5.5, 12.1, 0.55, "Źródła: feedstuffs.com i agfundernews.com (Bushel); agritechtomorrow.com, xfarm.ag (xFarm × Yara, 11.2023); prnewswire.com, zabkagroup.com (Żabka Nano).", 10, False, GREY)
msg(sl, [("Wniosek: ", True, WHITE), ("producent z siecią sprzedaży + gotowa technologia partnera skalują się szybko i tanio. Dokładnie tak budujemy GAia.", False, WHITE)], y=6.3)

# ===== 15. Jak się dzielą bez pieniędzy =====
sl = frame("Rozliczenie bez wkładu gotówki")
data = [["Mechanizm", "Przykład", "Jak w GAia"],
        ["Aplikacja white-label pod marką sieci sprzedaży", "Bushel × spółdzielnie (USA)", "GAia pod marką GA, dystrybutorzy odpowiadają na zapytania"],
        ["Wiedza producenta w silniku partnera", "xFarm × Yara", "5 funkcjonalności GA, nawozy.eu, IUNG w silniku SH"],
        ["Korporacja daje rynek, startup technologię", "Żabka × AiFi (Żabka Nano)", "GA: klienci, produkty, logistyka; SH i PM: technologia"],
        ["Praca za udziały lub udział w przychodach", "Venture studia: The Heart, Daftcode", "SH i PM jako współwłaściciele GAia, nie dostawcy na godziny"],
        ["Pieniądze od partnera finansowego", "Yara × Oxbury (bank dla rolników)", "Bank finansuje zakupy rolników w koszyku i współfinansuje start"],
        ["Przetwórca płaci za niższy ślad węglowy", "PepsiCo × Yara (1 000 gospodarstw)", "Programy niskoemisyjne GA z przetwórcami przez GAia"],
        ["Najpierw współpraca, potem spółka", "Bosch × BASF (2019 → 2021)", "Etap 1: umowa i pula, etap 2: spółka po bramce"],
        ["Dotacja zamiast wkładu", "EIT Food, FENG SMART, Horizon Europe", "Granty na model AI i pilot zmniejszają deficyt partnerów"]]
table(sl, 0.6, 1.2, 12.1, data, [4.2, 3.6, 4.3], rh=0.52, size=11)
msg(sl, [("Wzorzec: ", True, WHITE), ("każdy wnosi to, co ma, i zarabia na wyniku. Pieniądze dokładają ci, dla których to biznes: bank, przetwórcy, programy UE.", False, WHITE)], y=6.3)

# ===== 16. Dofinansowania =====
sl = frame("Dofinansowania: kto może dołożyć pieniądze do GAia", "Stan 11.2026, do weryfikacji z doradcą grantowym. Nabory zmieniają terminy.")
data = [["Źródło", "Na co", "Warunki / kwoty", "Termin", "Jak GAia się kwalifikuje"],
        ["FENG Ścieżka SMART (NCBR)", "B+R, wdrożenie, cyfryzacja, zazielenienie", "Duża firma w konsorcjum z MŚP; do 150 mln zł; min. 3 mln zł B+R", "Nabór 2026 zamknięty; kolejny do obserwowania", "GA lider + SH (MŚP): model AI nawożenia"],
        ["Horizon Europe, Klaster 6", "Cyfrowe i AI narzędzia do nawożenia", "ok. 6 mln EUR na projekt (FARM2FORK-01)", "23.09.2027", "GA + SH + IUNG / IOR-PIB; GAia jako demonstrator"],
        ["Digital Europe: AgrifoodTEF", "Testy AI w rolnictwie w polu", "Usługi testowe; Polska węzłem", "2023-2027", "Walidacja modelu AI GAia"],
        ["EIT Food", "Piloty startupów z korporacjami", "Granty bez oddawania udziałów", "Nabory co roku", "Spółka GAia / SH z GA jako partnerem"],
        ["Innovation Fund, NFOŚiGW, KPO", "Dekarbonizacja produkcji", "Nabór IF 2025: 77 projektów, 3,4 mld EUR", "Nabory co roku", "Zakłady GA; GAia buduje popyt i dane"],
        ["Ekoschemat: plan nawożenia (rolnik)", "Dopłata za plan nawożenia", "ok. 87,52 zł/ha (2025, dane wtórne)", "Kampania ARiMR", "Plan z GAia = dopłata dla rolnika = powód, by korzystać"]]
table(sl, 0.6, 1.2, 12.1, data, [2.6, 2.4, 2.8, 1.8, 2.5], rh=0.6, size=10)
msg(sl, [("Najlepsza droga: ", True, WHITE), ("GA jako lider konsorcjum z SH (MŚP) i instytutem (IUNG). Grant pokrywa model AI i pilot, czyli właśnie deficyt startu.", False, WHITE)], y=6.3)

# ===== 17. Dekarbonizacja =====
sl = frame("Dekarbonizacja: GAia jako dowód i argument", "To otwiera granty, programy z przetwórcami i ochronę przed CBAM")
data = [["Mechanizm", "Jak działa w GAia", "Dowód / benchmark"],
        ["Model AI precyzyjnego nawożenia", "Plan z paszportu pola, NDVI i danych ARiMR: mniej azotu tam, gdzie niepotrzebny", "Badania: 5-16% mniej N2O; Nutrient Expert: 12-20% niższy ślad w pszenicy"],
        ["Więcej nawozów niskoemisyjnych", "Rekomendacje i promocje np. mocznika z inhibitorem (Pulrea +INu)", "Cel GA: −9% śladu węglowego do 2030 vs 2024"],
        ["Dane o śladzie (MRV)", "Historia nawożenia w paszporcie pola, zgody rolnika", "PepsiCo × Yara: przetwórca płaci za niższy ślad"],
        ["CBAM od 01.01.2026", "Import nawozów droższy, liczy się ślad węglowy", "Rabobank: amoniak +10-20%, mocznik +10-15% w 2026"],
        ["Ekoschemat dla rolnika", "Plan nawożenia z GAia spełnia wymóg", "ok. 87,52 zł/ha (2025)"]]
table(sl, 0.6, 1.2, 12.1, data, [3.0, 4.8, 4.3], rh=0.55, size=11)
box(sl, 0.6, 4.75, 12.1, 0.95, fill=LIGHT)
tb(sl, 0.75, 4.78, 11.8, 0.9, [[("Przykład skali (do policzenia z IUNG): ", True, NAVY), ("1 mln ha × 150 kg N/ha × 5% mniej azotu = 7,5 tys. t N rocznie. Przy założeniu ok. 8 kg CO2e na 1 kg N (produkcja + pole) to ok. 60 tys. t CO2e rocznie mniej.", False, TXT)]], 12.5, anchor=MSO_ANCHOR.MIDDLE)
msg(sl, [("GAia sprzedaje więcej zielonych nawozów i mierzy efekt. ", True, WHITE), ("To warunek dla grantów, programów z przetwórcami i raportowania emisji GA.", False, WHITE)], y=6.3)

# ===== 16. Kroki =====
sl = frame("Dalsze kroki")
data = [["Termin", "Krok", "Kto"], ["do 20.11", "Odpowiedzi SH na pytania PZ-01…PZ-49 i potwierdzenie dopasowania MVP1", "SH + zespół GAia"],
        ["do 30.11", "Mapa dofinansowań z doradcą grantowym: konsorcjum GA + SH + IUNG (FENG, Horizon Europe, EIT Food)", "GA + doradca"],
        ["do 30.11", "Audyt 5 funkcjonalności GA na Azure; dane GA do modelu (marża, wolumen)", "GA IT + Kontroling"],
        ["do 15.12", "Term sheet: pula, klucz, opłata za sukces 25%, okres 10 + 5 lat, bramki", "GA + SH + Partner Marketplace"],
        ["01.2027", "Umowa partnerska (etap 1), start MVP1", "Zarząd GA"], ["02-04.2027", "Pilot MVP1 w sezonie wiosennym, bramka 1", "Wszyscy partnerzy"],
        ["koniec roku 2", "Bramka 2 i decyzja o spółce GAia", "Komitet Sterujący"]]
table(sl, 0.6, 1.2, 12.1, data, [1.6, 7.5, 3.0], rh=0.62, size=12)
msg(sl, [("Prosimy o: ", True, WHITE), ("zgodę na term sheet i pilot. GA nie angażuje gotówki, a każdy partner ma realny powód, żeby współpraca trwała lata.", False, WHITE)], y=6.3)

prs.save("../GA_GAia_model_wspolpracy_v3.pptx"); print("ok", N[0], "slajdów")
