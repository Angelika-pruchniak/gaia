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
KZ = rows(wb, "Każdy partner zyskuje"); U = rows(wb, "Użytkownicy i GMV"); J = rows(wb, "Wspólna pula"); T = rows(wb, "Partner Technologiczny")
MP = rows(wb, "Partner Marketplace"); G = rows(wb, "Grupa Azoty"); SC = rows(wb, "Scenariusze"); C = rows(wb, "Wkład technologiczny")

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
    tb(sl, 0.66, 6.95, 10.8, 0.35, f"GAia | wspólna pula przychodów | scenariusz bazowy, liczby z modelu do potwierdzenia   {N[0]}", 9, False, WHITE)
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
v = lambda r, i: r[i] if r[i] is not None else 0
def yrow(lab, r, d=1, color=None, tot=True):
    out = [lab] + [(fmt(v(r, i), d), color(v(r, i)) if color else TXT) for i in range(1, 7)]
    if tot: out.append(fmt(v(r, 7), d))
    return out
sign = lambda x: GREEN if x > 0.05 else (RED if x < -0.05 else TXT)
ga, tp, mp, pl = KZ["Grupa Azoty"], KZ["Partner Technologiczny"], KZ["Partner Marketplace / Handel"], KZ["Wspólna pula: wynik"]
sc = lambda lab: SC[lab][1:4]
users = U["Użytkownicy zarejestrowani (tys.)"]; gmv = U["GMV RAZEM"]; fg = U["GMV nawozów GA"]; g3 = U["GMV innych produktów od sprzedawców (3P)"]
pool = J["PRZYCHODY PULI RAZEM"]; cost = J["KOSZTY WSPÓLNE RAZEM"]; net = J["WYNIK PULI (przychody − koszty wspólne)"]; dist = J["DO PODZIAŁU WEDŁUG KLUCZA"]
fund = J["Deficyt puli (finansują PT 70% i PM 30%)"]; rep = J["Zwrot deficytu (pierwszeństwo przed podziałem)"]
captot = C["WKŁAD TECHNOLOGICZNY RAZEM"]; capPT = C["Razem Partner Technologiczny"]; capMP = C["Razem Partner Marketplace"]
names3 = ["Ostrożny", "Bazowy", "Ambitny"]

# ===== 1. Rekomendacja =====
sl = frame("Rekomendacja: wspólna pula przychodów", "GAia jako alians trzech partnerów: każdy wnosi to, co już ma, i dzieli się przychodami tam, gdzie tworzy wartość.")
tiles = [("0 zł", "wkładu pieniężnego GA"), (f"{fmt(captot[6],1)} mln zł", "pracy partnerów w roku 0"), (f"{fmt(sc('Przychody puli w roku 5')[1])} mln zł", "przychodów puli w roku 5 (bazowy)"),
         ("3 z 3", "partnerów na plusie w każdym scenariuszu")]
for i, (big, small) in enumerate(tiles):
    x = 0.6 + i * 3.07
    box(sl, x, 1.4, 2.9, 1.15, fill=GREEN if i in (0, 3) else NAVY)
    tb(sl, x, 1.45, 2.9, 0.65, big, 24, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    tb(sl, x, 2.08, 2.9, 0.42, small, 12, False, WHITE, PP_ALIGN.CENTER)
data = [["Wynik w 5 lat (gotówka, mln zł)", "Ostrożny\n100 tys. użytk.", "Bazowy\n200 tys. użytk.", "Ambitny\n300 tys. użytk."],
        ["Grupa Azoty"] + [(fmt(x, 1), sign(x)) for x in sc("Grupa Azoty: suma 5 lat")],
        ["Partner Technologiczny"] + [(fmt(x, 1), sign(x)) for x in sc("Partner Technologiczny: suma 5 lat")],
        ["Partner Marketplace / Handel"] + [(fmt(x, 1), sign(x)) for x in sc("Partner Marketplace: suma 5 lat")]]
table(sl, 0.6, 2.75, 12.1, data, [4.6, 2.5, 2.5, 2.5], rh=0.62, size=13)
tb(sl, 0.6, 5.35, 12.1, 0.7, [[("Etap 1 (lata 0-2): ", True, NAVY), ("umowa partnerska ze wspólną pulą, bez zakładania spółki. ", False, TXT), ("Etap 2 (od roku 3): ", True, NAVY), ("spółka GAia, jeśli pilot spełni bramki KPI.", False, TXT)]], 13)
msg(sl, [("Prosimy o: ", True, WHITE), ("zgodę na negocjacje umowy partnerskiej na tych zasadach i pilot w sezonie wiosennym.", False, WHITE)])

# ===== 2. Co zmieniłem i dlaczego =====
sl = frame("Co zmieniłem jako strateg, żeby model był realny")
data = [["Było", "Jest", "Dlaczego"],
        ["Jeden scenariusz: 300 tys. użytkowników", "3 scenariusze: 100 / 200 / 300 tys.; bazowy 200 tys.", "300 tys. to 25% rolników ARiMR. Bazowy ≈ tempo Nutrien (19% sprzedaży cyfrowo po ok. 3 latach)"],
        ["Spółka JV od pierwszego dnia", "Etap 1: umowa partnerska, etap 2: spółka po bramce KPI", "Bosch i BASF: współpraca od 2019, spółka dopiero w 2021"],
        ["GA płaci prowizję od każdego nawozu", "GA płaci 25% tylko od DODATKOWEJ marży", "GA nie oddaje zysku, który i tak ma; partnerzy zarabiają na wzroście"],
        ["Partnerzy na koszcie + 15%", "Koszty bieżące po koszcie z puli, zysk z podziału", "Nikt nie zarabia na kosztach; wszyscy na wyniku"],
        ["Prowizja 7% od sprzedawców", "Prowizja 5% + reklama 1,5%", "ŚOR i nasiona mają niskie marże (Allegro ma 12%, ale ogólnie)"],
        ["Udziały 34 / 40 / 26 bez uzasadnienia", "Klucz podziału dla każdego strumienia przychodów", "Każdy dostaje najwięcej tam, gdzie tworzy wartość (70/30 w sklepach z aplikacjami)"]]
table(sl, 0.6, 1.2, 12.1, data, [3.4, 4.0, 4.7], rh=0.66, size=11.5)
msg(sl, [("Efekt: ", True, WHITE), ("w każdym z trzech scenariuszy każdy partner jest na plusie, a GA nie wnosi gotówki.", False, WHITE)], y=6.3)

# ===== 3. Case study: kanał cyfrowy producenta =====
sl = frame("Case study: kanał cyfrowy producenta")
cases = [("Nutrien Digital", "producent i dystrybutor nawozów, USA / Kanada", [">1,2 mld USD sprzedaży cyfrowej w 2020 r.", "4× więcej niż rok wcześniej", "ok. 19% sprzedaży w I poł. 2021"], "Producent nawozów sprzedaje cyfrowo dużą część wolumenu"),
         ("Syngenta MAP", "producent ŚOR, Chiny", ["ok. 1,8 mld USD przychodu w 2021 r. (2×)", "492 → 638 centrów (2021 → 2023)", ">2 mln zarejestrowanych rolników"], "Platforma + sieć fizyczna rośnie szybko; u nas punkty wydania GA"),
         ("Yara × Agrofy", "producent nawozów + marketplace, Ameryka Płd.", ["Yara prowadziła rundę 30 mln USD (2021)", "łącznie 66 mln USD finansowania", "płatności i kredyt dla rolników"], "Producent i marketplace łączą siły zamiast budować sami")]
for i, (nm, what, pts, lesson) in enumerate(cases):
    x = 0.6 + i * 4.1
    box(sl, x, 1.2, 3.9, 0.8, fill=NAVY); tb(sl, x, 1.2, 3.9, 0.8, [[(nm, True, WHITE)], [(what, False, WHITE)]], 13, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    box(sl, x, 2.05, 3.9, 2.4, line=NAVY); tb(sl, x + 0.1, 2.15, 3.7, 2.25, ["• " + p for p in pts], 13)
    box(sl, x, 4.5, 3.9, 0.95, fill=LIGHT); tb(sl, x + 0.05, 4.5, 3.8, 0.95, lesson, 12.5, True, NAVY, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
tb(sl, 0.6, 5.55, 12.1, 0.5, "Źródła: Nutrien, wyniki 2020 (businesswire.com, 02.2021); Zacks (2021); Syngenta Group, wyniki 2021 i I kw. 2023 (syngentagroup.com, businesswire.com); agfundernews.com (12.2021).", 10, False, GREY)
msg(sl, [("Wniosek: ", True, WHITE), ("producent, który łączy produkt, sieć fizyczną i platformę, wygrywa. Agrando bez producenta upadło.", False, WHITE)], y=6.3)

# ===== 4. Case study: jak dzielą się partnerzy =====
sl = frame("Case study: jak partnerzy dzielą się przychodami")
data = [["Współpraca", "Kto co wnosi", "Jak się dzielą", "Wynik"],
        ["Starbucks × PepsiCo (NACP)", "Marka i produkt + dystrybucja", "Spółka 50/50 od 1994 r.", "Stworzyli rynek kawy gotowej do picia w USA"],
        ["Bosch × BASF (ONE SMART SPRAY)", "Technologia + agronomia", "Współpraca 2019, spółka 50/50 w 2021", "System w maszynach Amazone, AGCO, CNH, Stara"],
        ["Bayer × Microsoft", "Modele agronomiczne + chmura", "Każdy licencjonuje swoje rozwiązanie", "Usługi Bayer na Azure od 2023"],
        ["Sklepy z aplikacjami (Apple, Google)", "Platforma + twórca aplikacji", "70% dla twórcy, 30% dla platformy", "Standard rynku aplikacji"],
        ["Salesforce AppExchange", "Platforma + partner z aplikacją", "15% dla platformy, 85% dla partnera", "Standard rynku B2B"],
        ["Allegro", "Marketplace + sprzedawcy", "Prowizja 12% + reklama ok. 2% obrotu", "Benchmark stawek w Polsce"]]
table(sl, 0.6, 1.2, 12.1, data, [3.3, 3.0, 3.0, 2.8], rh=0.62, size=11.5)
tb(sl, 0.6, 5.65, 12.1, 0.4, "Źródła: pepsico.com, bosch-presse.de (05.2024), basf.com, businesswire.com (Bayer, 03.2023), developer.salesforce.com, entrepreneur.com, about.allegro.eu.", 10, False, GREY)
msg(sl, [("Wzorzec: ", True, WHITE), ("najpierw współpraca, potem spółka; każdy zostaje właścicielem swojej technologii; większość przychodu dla tego, kto tworzy wartość.", False, WHITE)], y=6.3)

# ===== 5. Kto co wnosi =====
sl = frame("Kto co wnosi i co dostaje")
cols = [("Grupa Azoty", NAVY, ["Produkty, produkcja, logistyka", "Punkty wydania, sieć sprzedaży, klienci", "Współpraca z ARiMR i dane", "5 funkcjonalności na Azure (audyt)"],
         ["75% dodatkowej marży z platformy", "20% prowizji, 40% finansowania i analiz", "0 zł wkładu pieniężnego"]),
        ("Partner Technologiczny", MBLUE, ["Gotowy silnik agro (licencja)", "UX: web i aplikacja mobilna", "Grywalizacja, ARiMR, Węzeł Krajowy", "Integruje całość, utrzymanie"],
         ["Zwrot kosztów bieżących z puli", "40% prowizji, 70% premium", "60% opłaty za sukces od GA"]),
        ("Partner Marketplace / Handel", LBLUE, ["Gotowy marketplace przez plug-in", "Sprzedawcy, katalog, koszyk", "Płatności i kredyt kupiecki", "Operacje marketplace'u"],
         ["Zwrot kosztów bieżących z puli", "40% prowizji, 40% finansowania", "40% opłaty za sukces od GA"])]
for i, (name, colr, gives, earns) in enumerate(cols):
    x = 0.6 + i * 4.1
    box(sl, x, 1.2, 3.9, 0.6, fill=colr); tb(sl, x, 1.2, 3.9, 0.6, name, 15, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    box(sl, x, 1.85, 3.9, 1.95, line=colr); tb(sl, x + 0.1, 1.9, 3.7, 0.3, "WNOSI", 11, True, GREY); tb(sl, x + 0.1, 2.17, 3.7, 1.6, ["• " + t for t in gives], 12.5)
    box(sl, x, 3.87, 3.9, 1.5, fill=LIGHT); tb(sl, x + 0.1, 3.92, 3.7, 0.3, "DOSTAJE", 11, True, NAVY); tb(sl, x + 0.1, 4.19, 3.7, 1.15, ["• " + t for t in earns], 12.5)
tb(sl, 0.6, 5.5, 12.1, 0.6, [[("Partnerzy technologiczni w roku 0: ", True, NAVY), (f"praca z własnych zasobów: PT {fmt(capPT[6],1)} mln zł, PM {fmt(capMP[6],1)} mln zł. Łączą gotowe silniki zamiast budować od zera.", False, TXT)]], 12.5)
msg(sl, [("Każdy zachowuje swoją technologię: ", True, WHITE), ("silnik, marketplace i 5 funkcjonalności GA zostają u właścicieli, spółka korzysta na licencji.", False, WHITE)], y=6.3)

# ===== 6. Mechanizm puli =====
sl = frame("Jak działa wspólna pula: 4 kroki")
steps = [("1", "Przychody do puli", "prowizje, reklama, premium, finansowanie, analizy + 25% dodatkowej marży GA", NAVY),
         ("2", "Koszty wspólne po koszcie", "zespoły partnerów, płatności, marketing, obsługa, biuro projektu; bez marży", MBLUE),
         ("3", "Zwrot deficytu", "jeśli na starcie pula jest na minusie, dopłacają PT 70% i PM 30%; wraca im to pierwsze", AMBER),
         ("4", "Podział według klucza", "każdy strumień dzielony osobno: kto tworzy wartość, ten dostaje więcej", GREEN)]
for i, (n, t, d, col) in enumerate(steps):
    x = 0.6 + i * 3.07
    box(sl, x, 1.3, 2.9, 2.9, fill=col)
    tb(sl, x, 1.35, 2.9, 0.6, n, 30, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    tb(sl, x + 0.1, 1.95, 2.7, 0.6, t, 15, True, WHITE, PP_ALIGN.CENTER)
    tb(sl, x + 0.1, 2.6, 2.7, 1.5, d, 12, False, WHITE, PP_ALIGN.CENTER)
data = [["Rok 5, scenariusz bazowy, mln zł", "Kwota"], ["1. Przychody puli", fmt(pool[6], 1)], ["2. Koszty wspólne", fmt(-cost[6], 1)], ["3. Zwrot deficytu", fmt(-rep[6], 1)], ["4. Do podziału", (fmt(dist[6], 1), GREEN)]]
table(sl, 0.6, 4.4, 6.0, data, [4.0, 2.0], rh=0.33, size=12)
tb(sl, 6.9, 4.4, 5.8, 1.6, [[("Przejrzystość: ", True, NAVY), ("wspólny rachunek (escrow), miesięczne raporty, audyt roczny. GA nie dopłaca nigdy; płaci tylko 25% od zysku, który bez platformy by nie powstał.", False, TXT)]], 12.5)
msg(sl, [("Nikt nie zarabia na kosztach, wszyscy na wyniku. ", True, WHITE), ("To motywuje partnerów do wzrostu, a nie do zawyżania kosztów.", False, WHITE)], y=6.3)

# ===== 7. Klucz podziału =====
sl = frame("Klucz podziału przychodów")
K = [r for r in wb["Klucz podziału"].iter_rows(min_row=5, max_row=9, values_only=True)]
st5 = [J["Prowizje i reklama od sprzedawców 3P: GMV 3P × (5% + 1,5%)"][6], J["Abonament premium: abonenci × 240 zł"][6], J["Finansowanie zakupów: GMV × 0,4%"][6], J["Analizy dla producentów"][6], J["Opłata za sukces od GA: 25% dodatkowej marży GA"][6]]
data = [["Strumień", "Rok 5,\nmln zł", "GA", "PT", "PM", "Dlaczego tak"]]
for r, val in zip(K, st5): data.append([r[0], fmt(val, 1), f"{fmt(r[1]*100)}%", f"{fmt(r[2]*100)}%", f"{fmt(r[3]*100)}%", r[5]])
table(sl, 0.6, 1.2, 12.1, data, [3.4, 0.9, 0.7, 0.7, 0.7, 5.7], rh=0.68, size=11)
tb(sl, 0.6, 5.4, 12.1, 0.7, [[("Udział w puli w roku 5 (bazowy): ", True, NAVY), (f"GA {fmt(J['Udział w puli: Grupa Azoty'][6]*100)}%, PT {fmt(J['Udział w puli: Partner Technologiczny'][6]*100)}%, PM {fmt(J['Udział w puli: Partner Marketplace'][6]*100)}%. GA dostaje mniej z puli, bo większość korzyści ma na sprzedaży własnych nawozów.", False, TXT)]], 12.5)
msg(sl, [("Klucz jest do negocjacji: ", True, WHITE), ("w Excelu zmienia się w jednym arkuszu, a wyniki wszystkich partnerów przeliczają się same.", False, WHITE)], y=6.3)

# ===== 8. Trzy scenariusze =====
sl = frame("Trzy scenariusze: każdy na plusie")
rowsS = [("Użytkownicy w roku 5, tys.", "Użytkownicy w roku 5", 0), ("GMV w roku 5, mln zł", "GMV w roku 5", 0), ("Nawozy przez platformę / sprzedaż krajowa GA", "Nawozy przez platformę / sprzedaż krajowa GA (rok 5)", -1),
         ("Przychody puli w roku 5", "Przychody puli w roku 5", 1), ("Wynik puli w roku 5", "Wynik puli w roku 5", 1), ("Grupa Azoty: suma 5 lat", "Grupa Azoty: suma 5 lat", 1),
         ("Partner Technologiczny: suma 5 lat", "Partner Technologiczny: suma 5 lat", 1), ("Partner Technologiczny: NPV / zwrot", None, 1),
         ("Partner Marketplace: suma 5 lat", "Partner Marketplace: suma 5 lat", 1), ("Partner Marketplace: NPV / zwrot", None, 1), ("Maks. deficyt puli (PT 70% / PM 30%)", "Maks. deficyt puli", 1)]
data = [["mln zł", "Ostrożny", "Bazowy", "Ambitny"]]
for lab, key, d in rowsS:
    if key is None:
        pre = "Partner Technologiczny" if "Technolog" in lab else "Partner Marketplace"
        npv, zw = sc(f"{pre}: NPV"), sc(f"{pre}: zwrot od")
        data.append([lab] + [(f"{fmt(n,1)} / {z}", sign(n)) for n, z in zip(npv, zw)])
    else:
        vals = sc(key)
        data.append([lab] + [(f"{fmt(x*100)}%" if d < 0 else fmt(x, max(d, 0)), sign(x) if "suma" in lab else TXT) for x in vals])
table(sl, 0.6, 1.2, 7.6, data, [3.8, 1.25, 1.25, 1.3], rh=0.39, size=11, hl_rows=(6,))
ch = chart(sl, XL_CHART_TYPE.COLUMN_CLUSTERED, 8.4, 1.2, 4.3, 4.6, names3,
           [("Grupa Azoty", [round(x, 1) for x in sc("Grupa Azoty: suma 5 lat")]), ("Partner Technologiczny", [round(x, 1) for x in sc("Partner Technologiczny: suma 5 lat")]),
            ("Partner Marketplace", [round(x, 1) for x in sc("Partner Marketplace: suma 5 lat")])], [NAVY, MBLUE, LBLUE], labels=True, numfmt='#,##0')
msg(sl, [("Uczciwie: ", True, WHITE), ("w scenariuszu ostrożnym Partner Technologiczny wychodzi tylko na niewielki plus (NPV ≈ 0). Dlatego bramka KPI po pilocie.", False, WHITE)], y=6.3)

# ===== 9. Pula rok po roku =====
sl = frame("Wspólna pula rok po roku (scenariusz bazowy)", f"Przychody puli rosną z {fmt(pool[2],1)} do {fmt(pool[6],1)} mln zł. Pula na plusie od roku 3; łączny deficyt na starcie {fmt(fund[7],1)} mln zł.")
lines = [("Prowizje i reklama 3P (5% + 1,5%)", "Prowizje i reklama od sprzedawców 3P: GMV 3P × (5% + 1,5%)"), ("Abonament premium (240 zł / rok)", "Abonament premium: abonenci × 240 zł"),
         ("Finansowanie zakupów (0,4%)", "Finansowanie zakupów: GMV × 0,4%"), ("Analizy dla producentów", "Analizy dla producentów"), ("Opłata za sukces od GA (25%)", "Opłata za sukces od GA: 25% dodatkowej marży GA")]
cl = [("Koszty Partnera Technologicznego", "Koszty bieżące Partnera Technologicznego"), ("Koszty Partnera Marketplace", "Koszty bieżące Partnera Marketplace"), ("Płatności (0,6%)", "Płatności: GMV × 0,6%"),
      ("Marketing", "Marketing: nowi × 50 zł + marka 1,5 mln zł"), ("Obsługa klienta", "Obsługa klienta"), ("Biuro projektu", "Wspólne biuro projektu (3 osoby)")]
data = [["mln zł", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma"]]
for lab, key in lines: data.append(yrow(lab, J[key]))
data.append(yrow("PRZYCHODY PULI", pool))
for lab, key in cl: data.append(yrow(lab, J[key]))
data.append(yrow("KOSZTY WSPÓLNE", cost)); data.append(yrow("WYNIK PULI", net, color=sign)); data.append(yrow("Do podziału po zwrocie deficytu", dist, color=sign))
table(sl, 0.6, 1.2, 12.1, data, [4.6, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.5], rh=0.28, size=10.5, hl_rows=(6, 13, 14))

# ===== 10. Wkład technologiczny =====
sl = frame("Rok 0: partnerzy łączą swoje silniki", f"{fmt(captot[3])} osobomiesięcy × 20 tys. zł = {fmt(captot[6],1)} mln zł pracy z własnych zasobów. Nie budujemy nic od zera.")
mods = [r for r in wb["Wkład technologiczny"].iter_rows(min_row=5, max_row=19, values_only=True)]
data = [["Moduł", "Kto", "Osobo-\nmiesiące", "Koszt", "Rok 0", "Rok 1"]]
for r in mods: data.append([r[0], "PT" if r[1].startswith("Partner T") else "PM", fmt(r[3]), fmt(r[6], 2), fmt(r[7], 2), fmt(r[8], 2)])
data.append(["Wkład technologiczny razem", "", fmt(captot[3]), fmt(captot[6], 2), fmt(captot[7], 2), fmt(captot[8], 2)])
table(sl, 0.6, 1.2, 8.6, data, [4.6, 0.6, 0.9, 0.9, 0.8, 0.8], rh=0.27, size=9.5, hl_rows=(16,))
box(sl, 9.4, 1.2, 3.3, 4.6, line=NAVY)
tb(sl, 9.5, 1.28, 3.1, 4.5, [[("Partner Technologiczny", True, MBLUE)], [(f"{fmt(capPT[3])} osobomies., {fmt(capPT[6],1)} mln zł", False, TXT)], [("UX, aplikacja, grywalizacja, ARiMR, Węzeł Krajowy, plug-in, SAP", False, GREY)],
    [("Partner Marketplace", True, MBLUE)], [(f"{fmt(capMP[3])} osobomies., {fmt(capMP[6],1)} mln zł", False, TXT)], [("plug-in, kategorie agro, zamówienia do punktów wydania", False, GREY)],
    [("Grupa Azoty", True, GREEN)], [("0 zł: audyt 5 funkcjonalności na Azure, dane, ARiMR", False, TXT)]], 11.5)
msg(sl, [("Wkład wraca partnerom ", True, WHITE), ("z ich udziału w podziale puli od roku 3.", False, WHITE)], y=6.3)

# ===== 11/12. Partnerzy =====
def partner_slide(Pd, title, keyrow, scn, pct):
    own, fin, dst, res_ = Pd["0. Wkład własny (praca w roku 0-1)"], Pd[[k for k in Pd if k.startswith("2. Finansowanie deficytu")][0]], Pd["3. Udział w podziale puli"], Pd[keyrow]
    cumr = Pd["Wynik skumulowany"]; opx = Pd["KOSZTY BIEŻĄCE RAZEM"]
    sl = frame(title, f"Scenariusz bazowy: +{fmt(res_[7],1)} mln zł w 5 lat. Ostrożny: +{fmt(sc(scn+': suma 5 lat')[0],1)}, ambitny: +{fmt(sc(scn+': suma 5 lat')[2],1)} mln zł.")
    chart(sl, XL_CHART_TYPE.COLUMN_STACKED, 0.6, 1.2, 6.6, 3.3, Y05,
          [("Wkład własny", [round(v(own, i), 1) for i in range(1, 7)]), ("Deficyt i jego zwrot", [round(v(fin, i), 1) for i in range(1, 7)]), ("Udział w podziale puli", [round(v(dst, i), 1) for i in range(1, 7)])], [RED, AMBER, GREEN])
    data = [["mln zł", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma"], yrow("Koszty bieżące (pokrywa pula)", opx),
            yrow("0. Wkład własny (praca w roku 0-1)", own, color=sign), yrow(f"1. Deficyt puli ({pct}) i jego zwrot", fin, color=sign),
            yrow("2. Udział w podziale puli", dst, color=sign), yrow("WYNIK PARTNERA", res_, color=sign), yrow("Skumulowany", cumr, color=sign, tot=False) + [""]]
    table(sl, 0.6, 4.6, 12.1, data, [4.6, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.5], rh=0.24, size=10, hl_rows=(5,))
    box(sl, 7.5, 1.2, 5.2, 3.3, line=NAVY)
    tb(sl, 7.6, 1.28, 5.0, 3.2, [[("Rok 0-1: inwestuje", True, RED)], [(f"Praca z własnych zasobów ({fmt(-own[7],1)} mln zł) i {pct} deficytu puli.", False, TXT)],
        [("Od roku 1: koszty pokryte", True, MBLUE)], [("Pula zwraca koszty bieżące po koszcie: nie dokłada do utrzymania.", False, TXT)],
        [("Od roku 3: zarabia", True, GREEN)], [(f"Zwrot deficytu, potem udział w podziale: {fmt(dst[7],1)} mln zł w 5 lat (w roku 5: {fmt(dst[6],1)}).", False, TXT)]], 12)
partner_slide(T, "Partner Technologiczny: kiedy zarabia", "WYNIK PARTNERA TECHNOLOGICZNEGO", "Partner Technologiczny", "70%")
partner_slide(MP, "Partner Marketplace: kiedy zarabia", "WYNIK PARTNERA MARKETPLACE", "Partner Marketplace", "30%")

# ===== 13. Grupa Azoty =====
gres = G["WYNIK GRUPY AZOTY"]; extra = G["DODATKOWA MARŻA GA RAZEM"]
sl = frame("Grupa Azoty: zysk bez wkładu", f"Bazowy: +{fmt(gres[7],1)} mln zł w 5 lat. Ostrożny: +{fmt(sc('Grupa Azoty: suma 5 lat')[0],1)}, ambitny: +{fmt(sc('Grupa Azoty: suma 5 lat')[2],1)} mln zł.")
gl = [("Udział w rynku AN/CAN", "Większy udział AN/CAN: p.p. × 39 mln zł × 25%"), ("Niższy koszt kanału", "Niższy koszt kanału: GMV nawozów × 2%"), ("Lepszy mix", "Lepszy mix: GMV nawozów × 1%")]
data = [["mln zł", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma"]]
for lab, key in gl: data.append(yrow(lab, G[key]))
data.append(yrow("Dodatkowa marża GA razem", extra)); data.append(yrow("Opłata za sukces do puli (25%)", [None] + [-(x or 0) for x in J["Opłata za sukces od GA: 25% dodatkowej marży GA"][1:8]], color=sign))
data.append(yrow("Udział w podziale puli", G["2. Udział w podziale puli"], color=sign)); data.append(yrow("Koszt wewnętrzny GA", G["3. Koszt wewnętrzny GA (opiekun projektu, dane)"], color=sign))
data.append(yrow("WYNIK GRUPY AZOTY", gres, color=sign))
table(sl, 0.6, 1.2, 7.4, data, [2.8, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 1.0], rh=0.42, size=10.5, hl_rows=(4, 8))
chart(sl, XL_CHART_TYPE.COLUMN_CLUSTERED, 8.2, 1.2, 4.5, 4.5, Y15, [("Wynik GA", [round(x, 1) for x in gres[2:7]])], [NAVY], labels=True, numfmt='#,##0.0')
msg(sl, [("Jak Nutrien i Syngenta: ", True, WHITE), ("GA zarabia na sprzedaży własnych produktów. Partnerom oddaje tylko część dodatkowego zysku.", False, WHITE)], y=6.3)

# ===== 14. Etapy i bramki =====
sl = frame("Etapy i bramki decyzyjne", "Benchmark: Bosch i BASF najpierw współpracowali (2019), spółkę założyli po potwierdzeniu wyników (2021)")
gates = [("Rok 0", "Umowa partnerska", "Połączenie silnika, marketplace'u i 5 funkcjonalności GA; aplikacja i grywalizacja", "Bramka 0: aplikacja działa, audyt GA zrobiony", NAVY),
         ("Rok 1 (wiosna)", "Pilot", "1 województwo w sezonie wiosennym", "Bramka 1: ≥ 10 tys. użytkowników, ≥ 5 tys. zamówień, ocena ≥ 4/5", MBLUE),
         ("Rok 1-2", "Skalowanie na kraj", "Marketing przez sieć sprzedaży GA i ARiMR", "Bramka 2: ≥ 70 tys. użytkowników, pula ≥ 10 mln zł, wynik puli ≥ 0", LBLUE),
         ("Od roku 3", "Spółka GAia", "Udziały według klucza puli, zarząd, rada", "Wyjście bez straty dla GA, jeśli bramka nie spełniona", GREEN)]
for i, (when, what, desc, gate, col) in enumerate(gates):
    x = 0.6 + i * 3.07
    box(sl, x, 1.3, 2.9, 0.85, fill=col); tb(sl, x, 1.3, 2.9, 0.85, [[(when, False, WHITE)], [(what, True, WHITE)]], 13, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    box(sl, x, 2.2, 2.9, 1.5, line=col); tb(sl, x + 0.1, 2.25, 2.7, 1.4, desc, 12, False, TXT, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    box(sl, x, 3.8, 2.9, 1.5, fill=LIGHT); tb(sl, x + 0.1, 3.85, 2.7, 1.4, gate, 12, True, NAVY, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
msg(sl, [("GA ryzykuje tylko czas i dane, ", True, WHITE), ("a każda bramka pozwala zatrzymać projekt, zanim zrobi się drogi.", False, WHITE)], y=6.3)

# ===== 15. Ryzyka i kroki =====
sl = frame("Ryzyka, warunki umowy i dalsze kroki")
data = [["Ryzyko", "Co robimy"],
        ["Za mało użytkowników", "Bramki KPI; koszty wspólne rosną z użytkownikami, nie z góry."],
        ["Partner zawyża koszty", "Koszty po koszcie, bez marży; audyt roczny; budżet roczny zatwierdzany wspólnie."],
        ["Spór o podział", "Klucz zapisany w umowie dla każdego strumienia; przegląd raz w roku."],
        ["Utrata kontroli nad danymi", "Dane rolników współdzielone na zgodzie; silnik, marketplace i funkcjonalności GA zostają u właścicieli."]]
table(sl, 0.6, 1.2, 6.4, data, [2.0, 4.4], rh=0.72, size=11.5)
data = [["Termin", "Krok"], ["do 30.11", "Audyt 5 funkcjonalności GA na Azure, dane GA do modelu"],
        ["do 15.12", "Term sheet z obydwoma partnerami: pula, klucz, opłata za sukces 25%"],
        ["01.2027", "Umowa partnerska (etap 1), start prac"],
        ["02-04.2027", "Pilot w sezonie: 1 województwo, bramka 1"],
        ["koniec roku 2", "Bramka 2 i decyzja o spółce GAia"]]
table(sl, 7.2, 1.2, 5.5, data, [1.4, 4.1], rh=0.6, size=11.5)
msg(sl, [("Prosimy o: ", True, WHITE), ("zgodę na term sheet i pilot. GA nie angażuje gotówki, a każdy partner ma realny powód, żeby projekt się udał.", False, WHITE)], y=6.3)

prs.save("../GA_GAia_wspolna_pula.pptx"); print("ok", N[0], "slajdów")
