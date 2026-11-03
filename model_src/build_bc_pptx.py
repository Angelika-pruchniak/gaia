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
KZ, KZs = rows(wb, "Każdy partner zyskuje"), rows(st, "Każdy partner zyskuje")
U = rows(wb, "Użytkownicy i GMV"); J = rows(wb, "Spółka GAia (JV)"); T = rows(wb, "Partner Technologiczny")
MP = rows(wb, "Partner Marketplace"); G = rows(wb, "Grupa Azoty")

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
GA, TP, MPk, JV = "Grupa Azoty (34%)", "Partner Technologiczny (40%)", "Partner Marketplace / Handel (26%)", "Spółka GAia (JV) - przepływy pieniężne"
C = rows(wb, "Wkład technologiczny"); MC = rows(wb, "Moc platformy")
users = U["Użytkownicy zarejestrowani (tys.)"]; gmv = U["GMV RAZEM"]; fg = U["GMV nawozów GA"]; g3 = U["GMV innych produktów od sprzedawców (3P)"]
pen = U["Nawozy przez platformę / sprzedaż krajowa GA"]; buy = U["Aktywni kupujący (tys.) = użytkownicy × 35%"]
rev = J["PRZYCHODY RAZEM"]; opex = J["OPEX RAZEM"]; ebitda = J["EBITDA"]; cum = J["Przepływy skumulowane"]
fundr = J["Dopłaty partnerów na start (gdy przepływy < 0)"]; rep = J["Zwrot dopłat partnerom (najpierw z zysku)"]; div = J["Dywidenda do podziału według udziałów"]
fund = J["Maksymalne zapotrzebowanie na finansowanie"][1]; ftp = J["Partner Technologiczny"][1]; fmp = J["Partner Marketplace"][1]
captot = C["WKŁAD TECHNOLOGICZNY RAZEM (z własnych zasobów partnerów)"]; capPT = C["Razem Partner Technologiczny"]; capMP = C["Razem Partner Marketplace"]
v = lambda r, i: r[i] if r[i] is not None else 0
def yrow(lab, r, d=1, color=None, tot=True):
    out = [lab] + [(fmt(v(r, i), d), color(v(r, i)) if color else TXT) for i in range(1, 7)]
    if tot: out.append(fmt(v(r, 7), d))
    return out
sign = lambda x: GREEN if x > 0.05 else (RED if x < -0.05 else TXT)
JR = lambda key: J[key]

# ===== 1. Teza =====
sl = frame("GAia: wspólna spółka trzech partnerów", "GA wnosi produkty, logistykę, dane, ARiMR i 5 funkcjonalności, bez gotówki. Partnerzy technologiczni łączą swoje gotowe silniki.")
tiles = [(f"{fmt(users[6])} tys.", "użytkowników w roku 5"), (f"{fmt(gmv[6]/1000, 2)} mld zł", "obrotu (GMV) w roku 5"),
         (mln(rev[6]), "przychodu spółki w roku 5"), ("0 zł", "wkładu pieniężnego Grupy Azoty")]
for i, (big, small) in enumerate(tiles):
    x = 0.6 + i * 3.07
    box(sl, x, 1.45, 2.9, 1.2, fill=NAVY if i < 3 else GREEN)
    tb(sl, x, 1.5, 2.9, 0.7, big, 26, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    tb(sl, x, 2.15, 2.9, 0.45, small, 12.5, False, WHITE, PP_ALIGN.CENTER)
data = [["Partner (gotówka)", "Wynik w roku 5", "Suma 5 lat (z rokiem 0)", "NPV (10%)", "Zwrot od"]]
for nm in (GA, TP, MPk):
    r = k(nm); data.append([nm, (mln(r[6], 1), GREEN), (mln(r[7], 1), GREEN), mln(r[8], 1), r[10]])
r = k(JV); data.append(["Spółka GAia: gotówka platformy", (mln(r[6], 1), GREEN), (mln(r[7], 1), GREEN), mln(r[8], 1), r[10]])
table(sl, 0.6, 2.85, 12.1, data, [4.0, 2.0, 2.4, 1.9, 1.8], rh=0.55, size=13, hl_rows=(1,))
msg(sl, [("Prosty mechanizm: ", True, WHITE), ("partnerzy wnoszą pracę w pierwszym roku, zarabiają na licencjach do swoich silników, usługach (koszt + 15%) i dywidendzie. GA zarabia na sprzedaży i dywidendzie.", False, WHITE)])

# ===== 2. Jak działa współpraca =====
sl = frame("Jak działa współpraca", "Spółka GAia: trzech wspólników, jeden produkt dla rolnika")
cols = [("Grupa Azoty", "34%", NAVY, ["Produkty, produkcja, logistyka", "Współpraca z ARiMR i dane", "Sieć sprzedaży i punkty wydania, klienci", "5 funkcjonalności na Azure (audyt, bez wyceny)"],
         ["Bez wkładu pieniężnego", "34% dywidendy", "Większy udział w rynku, mix, niższy koszt kanału"]),
        ("Partner Technologiczny (główny)", "40%", MBLUE, ["Gotowy silnik agro (licencja)", "UX: web i aplikacja mobilna", "Grywalizacja, ARiMR, Węzeł Krajowy", "Łączy całość, utrzymanie 24/7"],
         ["Licencja: 10 zł / użytkownik / rok", "Koszt + 15% za utrzymanie", "40% dywidendy"]),
        ("Partner Marketplace / Handel", "26%", LBLUE, ["Gotowy marketplace (licencja)", "Podpięty do GAia przez plug-in", "Katalog, koszyk, płatności, sprzedawcy", "Operacje marketplace'u"],
         ["Licencja: 1% obrotu innych produktów", "Koszt + 15% za operacje", "26% dywidendy"])]
for i, (name, sh, colr, gives, earns) in enumerate(cols):
    x = 0.6 + i * 4.1
    box(sl, x, 1.2, 3.9, 0.75, fill=colr)
    tb(sl, x, 1.2, 3.9, 0.75, [[(name, True, WHITE)], [(f"udział {sh}", False, WHITE)]], 14, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    box(sl, x, 2.0, 3.9, 1.85, line=colr)
    tb(sl, x + 0.1, 2.05, 3.7, 0.3, "WNOSI", 11, True, GREY)
    tb(sl, x + 0.1, 2.32, 3.7, 1.5, ["• " + t for t in gives], 12)
    box(sl, x, 3.92, 3.9, 1.35, fill=LIGHT)
    tb(sl, x + 0.1, 3.97, 3.7, 0.3, "ZARABIA", 11, True, NAVY)
    tb(sl, x + 0.1, 4.24, 3.7, 1.0, ["• " + t for t in earns], 12)
box(sl, 0.6, 5.38, 12.1, 0.75, line=NAVY)
tb(sl, 0.7, 5.38, 11.9, 0.75, [[("Pierwszy rok: ", True, NAVY), (f"partnerzy z własnych zasobów dopracowują GAia i aplikację mobilną ({mln(captot[7],1)} pracy). Nie budujemy nic od zera: łączymy silnik, marketplace i 5 funkcjonalności GA.", False, TXT)]], 12.5, anchor=MSO_ANCHOR.MIDDLE)
msg(sl, [("GA nie wpłaca pieniędzy. ", True, WHITE), ("Start spółki finansują partnerzy technologiczni, a zysk najpierw zwraca im dopłaty, potem idzie na dywidendę.", False, WHITE)])

# ===== 3. Architektura =====
sl = frame("Architektura GAia: łączymy gotowe elementy", "Partner Technologiczny integruje całość. Nowe są tylko UX, grywalizacja i integracje.")
box(sl, 0.6, 1.2, 12.1, 0.95, fill=NAVY)
tb(sl, 0.6, 1.2, 12.1, 0.95, [[("APLIKACJA GAia dla rolnika (web + iOS / Android)", True, WHITE)], [("nowa warstwa UX + grywalizacja: punkty, odznaki, wyzwania sezonowe, nagrody  |  Partner Technologiczny", False, WHITE)]], 14, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
layers = [("Silnik agro", "Partner Technologiczny", "plan upraw, mapy pól, rekomendacje, asystent AI", MBLUE, "gotowy, licencja"),
          ("Marketplace", "Partner Marketplace / Handel", "katalog, koszyk, płatności, sprzedawcy, reklama", LBLUE, "gotowy, plug-in, licencja"),
          ("5 funkcjonalności GA", "Grupa Azoty, Azure", "plan nawożenia, CRM, Nawozy.eu, koszyk, portal sprzedaży", GREEN, "gotowe, audyt i API")]
for i, (nm, who, what, col, tag) in enumerate(layers):
    x = 0.6 + i * 4.1
    box(sl, x, 2.35, 3.9, 1.75, fill=col)
    tb(sl, x, 2.38, 3.9, 1.7, [[(nm, True, WHITE)], [(who, False, WHITE)], [(what, False, WHITE)], [(tag, True, WHITE)]], 12.5, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
ints = [("ARiMR", "działki i uprawy z wniosków"), ("Węzeł Krajowy / mObywatel", "logowanie profilem zaufanym"), ("SAP / CRM GA", "ceny, dostępność, zamówienia"),
        ("Punkty wydania GA", "odbiór i dostawa nawozów"), ("Bank", "kredyt kupiecki w koszyku")]
for i, (nm, what) in enumerate(ints):
    x = 0.6 + i * 2.44
    box(sl, x, 4.3, 2.3, 0.95, line=NAVY)
    tb(sl, x, 4.3, 2.3, 0.95, [[(nm, True, NAVY)], [(what, False, TXT)]], 11, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
tb(sl, 0.6, 5.35, 12.1, 0.6, [[("Integracje (Partner Technologiczny): ", True, NAVY), ("wszystkie przez API. Marketplace podpinany przez plug-in, 5 funkcjonalności GA po audycie, bez przepisywania.", False, TXT)]], 12)
msg(sl, [("Szybko i tanio: ", True, WHITE), (f"{fmt(captot[3])} osobomiesięcy pracy w roku 0 zamiast budowy platformy od zera.", False, WHITE)])

# ===== 4. Benchmark JV =====
sl = frame("Benchmark: jak buduje się takie współprace", "Wnioski z praktyk joint venture i przykładów z rynku agro")
data = [["Praktyka", "Przykład / źródło", "Jak to robimy w GAia"],
        ["Udziały za wkład rzeczowy, nie tylko gotówkę", "Bosch × BASF Smart Farming: JV 50/50, każdy wnosi swoją technologię (2021)", "GA wnosi produkt, dane i ARiMR; partnerzy wnoszą silniki i pracę"],
        ["IP wnoszone licencją, nie przenoszone", "WIPO, Ankura: licencja jest bardziej elastyczna niż przeniesienie IP", "Silnik i marketplace zostają u partnerów; spółka płaci licencję"],
        ["Usługi wspólników po koszcie z marżą", "Umbrex: usługi dla JV po koszcie, wartość w udziałach", "Koszt + 15% za utrzymanie i operacje"],
        ["Zwrot wkładu finansującego przed dywidendą", "Praktyka JV: preferencyjny zwrot dopłat wspólnika", "Zysk najpierw zwraca dopłaty partnerów, potem dywidenda"],
        ["Ochrona wspólnika bez gotówki", "KSH: zmiana umowy spółki wymaga 2/3 głosów", "GA 34%: pakiet blokujący bez wkładu pieniężnego"],
        ["Producent inwestuje w kanał cyfrowy", "Yara prowadziła rundę 30 mln USD w marketplace Agrofy (2021)", "GA jest wspólnikiem, a nie tylko sprzedawcą na platformie"]]
table(sl, 0.6, 1.2, 12.1, data, [3.6, 4.5, 4.0], rh=0.62, size=11.5)
tb(sl, 0.6, 5.65, 12.1, 0.4, "Źródła: basf.com (11.2020), umbrex.com, jvalchemist.ankura.com, wipo.int, agfundernews.com (Agrofy, 12.2021).", 10, False, GREY)
msg(sl, [("Wniosek: ", True, WHITE), ("każdy wnosi to, co ma najlepsze, licencje i usługi rozliczane osobno, a ryzyko finansowe biorą partnerzy technologiczni.", False, WHITE)])

# ===== 5. Benchmark: na czym zarabia konkurencja =====
sl = frame("Benchmark: na czym zarabia konkurencja", "Z tego wynikają stawki w modelu")
data = [["Firma", "Liczby", "Na czym najwięcej zarabia", "Stawka w GAia"],
        ["Allegro", "Prowizja 12,0%, reklama ok. 2,0% GMV (IV kw. 2024)", "Prowizje od sprzedawców i reklama", "Prowizja 7%, reklama 1,5%"],
        ["Yara (YaraPlus)", "ok. 5 mln rolników w narzędziach cyfrowych", "Na nawozach. Aplikacje są darmowe", "Podstawa GAia za darmo"],
        ["BASF xarvio", ">130 tys. użytkowników, >20 mln ha (2025)", "Na środkach ochrony roślin", "Premium tanie, sprzedaż produktu"],
        ["Bayer FieldView", "UK: Prime od 120 £/rok, Plus od 500 £/rok", "Abonamenty + nasiona i ŚOR", "Premium 20 zł / mies. (240 zł / rok)"],
        ["SatAgro (PL)", "10-15 zł za hektar rocznie", "Abonament za hektar", "240 zł / rok = gospodarstwo ok. 20 ha"],
        ["FBN (USA)", "ok. 120 tys. gospodarstw, >1 mld USD finansowania", "Środki produkcji i finansowanie", "Finansowanie 0,4% obrotu"],
        ["DeHaat, Agrando", "Strata operacyjna / zamknięcie w 2022", "Własne magazyny / brak producenta", "Realizuje sieć GA, marketplace gotowy"]]
table(sl, 0.6, 1.2, 12.1, data, [2.2, 4.0, 3.4, 2.5], rh=0.55, size=11.5)
tb(sl, 0.6, 5.65, 12.1, 0.4, "Źródła: about.allegro.eu (wyniki IV kw. 2024), farming.co.uk, basf.com (09.2025), climate.com/en-gb/pricing, wiescirolnicze.pl, agfundernews.com, entrackr.com.", 10, False, GREY)
msg(sl, [("Konkurencja zarabia na produkcie, a nie na aplikacji. ", True, WHITE), ("Dlatego GA zarabia głównie na udziale w rynku, a spółka na prowizjach i reklamie.", False, WHITE)])


# ===== 5. Ścieżka + dane =====
sl = frame("Ścieżka do 300 tys. użytkowników: dane", "GMV = aktywni kupujący × zakupy na kupującego. 300 tys. to ok. 25% z 1,21 mln rolników w ARiMR.")
chart(sl, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.6, 1.5, 5.6, 1.9, Y15, [("Użytkownicy, tys.", [round(x) for x in users[2:7]])], [NAVY], labels=True)
tb(sl, 0.6, 1.2, 5.6, 0.3, "Użytkownicy na koniec roku, tys.", 11, True, NAVY)
chart(sl, XL_CHART_TYPE.COLUMN_STACKED, 6.7, 1.5, 6.0, 1.9, Y15, [("Nawozy GA", [round(x) for x in fg[2:7]]), ("Inne produkty (3P)", [round(x) for x in g3[2:7]])], [NAVY, LBLUE], labels=True)
tb(sl, 6.7, 1.2, 6.0, 0.3, "GMV, mln zł", 11, True, NAVY)
nu = U["Nowi użytkownicy w roku (tys.)"]; bf = U["Zakupy nawozów na kupującego (tys. zł/rok)"]; b3 = U["Zakupy innych produktów na kupującego (tys. zł/rok)"]; pr_ = U["Abonenci premium (tys.)"]
data = [["Dane", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Skąd"],
        ["Użytkownicy zarejestrowani, tys."] + [fmt(x) for x in users[2:7]] + ["cel"],
        ["Aktywni kupujący, tys."] + [fmt(x) for x in buy[2:7]] + ["35% użytkowników"],
        ["Abonenci premium (20 zł / mies.), tys."] + [fmt(x, 1) for x in pr_[2:7]] + ["2% → 6% użytk."],
        ["Zakupy nawozów na kupującego, tys. zł"] + [fmt(x, 1) for x in bf[2:7]] + ["założenie"],
        ["Zakupy innych produktów na kupującego, tys. zł"] + [fmt(x, 1) for x in b3[2:7]] + ["założenie"],
        ["GMV nawozów GA, mln zł"] + [fmt(x) for x in fg[2:7]] + ["kupujący × zakupy"],
        ["GMV innych produktów (3P), mln zł"] + [fmt(x) for x in g3[2:7]] + ["kupujący × zakupy"],
        ["GMV razem, mln zł"] + [fmt(x) for x in gmv[2:7]] + ["suma"],
        ["Nawozy przez platformę / sprzedaż krajowa GA"] + [f"{fmt(x*100)}%" for x in pen[2:7]] + ["÷ 3,69 mld zł"]]
table(sl, 0.6, 3.45, 12.1, data, [4.3, 1.2, 1.2, 1.2, 1.2, 1.2, 1.8], rh=0.27, size=11, hl_rows=(8,))


# ===== 6. Moc platformy =====
sl = frame("Moc platformy w liczbach", "Rok 5, przy 300 tys. użytkowników")
tiles = [(f"{fmt(users[6])} tys.", "użytkowników"), (f"{fmt(buy[6])} tys.", "aktywnych kupujących"), (f"{fmt(MC['Abonenci premium (tys.)'][6])} tys.", "abonentów premium"),
         (f"{fmt(MC['Sprzedawcy innych produktów'][6])}", "sprzedawców innych produktów"), (f"{fmt(MC['Zamówienia razem (tys.)'][6])} tys.", "zamówień rocznie"),
         (f"{fmt(MC['Nawozy przez platformę (tys. t)'][6])} tys. t", "nawozów GA przez platformę"), (f"{fmt(MC['Zamówienia nawozów na 1 punkt wydania dziennie w szczycie'][6])}", "zamówień na punkt wydania dziennie w szczycie"),
         (f"{fmt(MC['Przychód na użytkownika (zł/rok)'][6])} zł", "przychodu spółki na użytkownika")]
for i, (big, small) in enumerate(tiles):
    x = 0.6 + (i % 4) * 3.07; yy = 1.25 + (i // 4) * 1.2
    box(sl, x, yy, 2.9, 1.08, fill=NAVY if i < 4 else MBLUE)
    tb(sl, x, yy + 0.04, 2.9, 0.6, big, 22, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    tb(sl, x, yy + 0.62, 2.9, 0.4, small, 11, False, WHITE, PP_ALIGN.CENTER)
mrows = [("Zamówienia razem, tys.", "Zamówienia razem (tys.)", 0), ("Nawozy przez platformę, tys. t", "Nawozy przez platformę (tys. t)", 0),
         ("Przychód spółki / GMV (efektywna prowizja)", "Przychód spółki / GMV (efektywna prowizja)", -1), ("Przychód na użytkownika, zł/rok", "Przychód na użytkownika (zł/rok)", 0),
         ("OPEX na użytkownika, zł/rok", "OPEX na użytkownika (zł/rok)", 0), ("EBITDA na użytkownika, zł/rok", "EBITDA na użytkownika (zł/rok)", 0)]
data = [["Rok po roku", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5"]]
for lab, key, d in mrows:
    r = MC[key]; data.append([lab] + [(f"{fmt(x*100,1)}%" if d < 0 else fmt(x, d), sign(x) if 'EBITDA' in key else TXT) for x in r[2:7]])
table(sl, 0.6, 3.75, 12.1, data, [4.6, 1.5, 1.5, 1.5, 1.5, 1.5], rh=0.32, size=11.5)



# ===== 8. Wkład technologiczny =====
sl = frame("Rok 0: partnerzy dopracowują GAia", f"{fmt(captot[3])} osobomiesięcy × 20 tys. zł = {mln(captot[7],1)} pracy z własnych zasobów partnerów. Spółka GAia nie ma CAPEX.")
mods = [r for r in wb["Wkład technologiczny"].iter_rows(min_row=5, max_row=19, values_only=True)]
data = [["Moduł", "Kto", "Osobo-\nmiesiące", "Koszt,\nmln zł", "Rok 0", "Rok 1"]]
for r in mods: data.append([r[0], "PT" if r[1].startswith("Partner T") else "PM", fmt(r[3]), fmt(r[7], 2), fmt(r[8], 2), fmt(r[9], 2)])
data.append(["Partner Technologiczny razem", "PT", fmt(capPT[3]), fmt(capPT[7], 2), fmt(capPT[8], 2), fmt(capPT[9], 2)])
data.append(["Partner Marketplace razem", "PM", fmt(capMP[3]), fmt(capMP[7], 2), fmt(capMP[8], 2), fmt(capMP[9], 2)])
data.append(["Wkład technologiczny razem", "", fmt(captot[3]), fmt(captot[7], 2), fmt(captot[8], 2), fmt(captot[9], 2)])
table(sl, 0.6, 1.2, 8.6, data, [4.6, 0.6, 0.9, 0.9, 0.8, 0.8], rh=0.255, size=9.5, hl_rows=(18,))
box(sl, 9.4, 1.2, 3.3, 4.95, line=NAVY)
tb(sl, 9.5, 1.28, 3.1, 4.85, [
    [("Partner Technologiczny", True, MBLUE)], [(f"{fmt(capPT[3])} osobomies., {mln(capPT[7],1)}", False, TXT)], [("UX, aplikacja, grywalizacja, ARiMR, Węzeł Krajowy, plug-in, SAP", False, GREY)],
    [("Partner Marketplace", True, MBLUE)], [(f"{fmt(capMP[3])} osobomies., {mln(capMP[7],1)}", False, TXT)], [("plug-in, kategorie agro, zamówienia do punktów wydania", False, GREY)],
    [("Grupa Azoty", True, GREEN)], [("0 zł. Udostępnia 5 funkcjonalności z Azure (audyt), dane i współpracę z ARiMR", False, TXT)],
    [("Już istnieje, nie budujemy", True, NAVY)], [("silnik agro, marketplace, 5 funkcjonalności GA", False, TXT)]], 11)
msg(sl, [("Partnerzy odzyskują ten wkład ", True, WHITE), ("z licencji do swoich silników, marży na usługach i dywidendy.", False, WHITE)], y=6.3)

# ===== 9. Przychody i OPEX =====
sl = frame("Spółka GAia: przychody i OPEX rok po roku", "mln zł. Stawki z benchmarku konkurencji i marketplace'ów.")
rl = [("Prowizje: GMV 3P × 7%", "Prowizje: GMV 3P × 7%"), ("Reklama: GMV 3P × 1,5%", "Reklama: GMV 3P × 1,5%"), ("Prowizja od nawozów GA: × 1,2%", "Prowizja od nawozów GA: GMV nawozów × 1,2%"),
      ("Finansowanie: GMV × 0,4%", "Finansowanie: GMV × 0,4%"), ("Abonament premium: × 240 zł / rok", "Abonament premium: abonenci × 240 zł"), ("Analizy dla producentów", "Analizy dla producentów")]
ol = [("Partner Technologiczny: utrzymanie (koszt + 15%)", "Partner Technologiczny: utrzymanie i rozwój (koszt + 15%)"), ("Partner Marketplace: operacje (koszt + 15%)", "Partner Marketplace: operacje (koszt + 15%)"),
      ("Licencja: silnik (10 zł / użytkownik)", "Licencja: silnik Partnera Technologicznego (10 zł / użytkownik)"), ("Licencja: marketplace (1% GMV 3P)", "Licencja: marketplace Partnera Marketplace (1% GMV 3P)"),
      ("Płatności: GMV × 0,6%", "Płatności: GMV × 0,6%"), ("Marketing: nowi × 60 zł + marka 3 mln", "Marketing: nowi użytkownicy × 60 zł + marka 3 mln zł"),
      ("Obsługa klienta", "Obsługa klienta: 1 konsultant na 20 tys. użytkowników"), ("Zarząd i zespół spółki (6 osób)", "Zarząd i zespół spółki (6 osób)")]
data = [["mln zł", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma"]]
for lab, key in rl: data.append(yrow(lab, J[key]))
data.append(yrow("PRZYCHODY RAZEM", rev))
for lab, key in ol: data.append(yrow(lab, J[key]))
data.append(yrow("OPEX RAZEM", opex)); data.append(yrow("EBITDA (= gotówka spółki)", ebitda, color=sign))
table(sl, 0.6, 1.2, 12.1, data, [4.6, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.5], rh=0.272, size=10.5, hl_rows=(7, 16, 17))

# ===== 10. Gotówka i finansowanie startu =====
sl = frame("Spółka GAia: gotówka, start i dywidenda", f"Start finansują partnerzy technologiczni: {mln(fund,1)} (PT {mln(ftp,1)}, PM {mln(fmp,1)}). GA: 0 zł.")
chart(sl, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.6, 1.2, 6.4, 3.3, Y05,
      [("EBITDA", [round(v(ebitda, i), 1) for i in range(1, 7)]), ("Dywidenda", [round(v(div, i), 1) for i in range(1, 7)]), ("Przepływy skumulowane", [round(v(cum, i), 1) for i in range(1, 7)])], [NAVY, GREEN, LBLUE])
data = [["mln zł", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma"], yrow("Przychody", rev), yrow("OPEX", opex), yrow("EBITDA = gotówka spółki", ebitda, color=sign),
        yrow("Dopłaty partnerów na start", fundr), yrow("Zwrot dopłat partnerom", rep), yrow("Dywidenda (34 / 40 / 26)", div, color=sign), yrow("Przepływy skumulowane", cum, color=sign, tot=False) + [""]]
table(sl, 0.6, 4.6, 12.1, data, [4.1, 1.05, 1.05, 1.05, 1.05, 1.05, 1.05, 1.5], rh=0.22, size=9.5, hl_rows=(6,))
box(sl, 7.3, 1.2, 5.4, 3.3, line=NAVY)
tb(sl, 7.4, 1.28, 5.2, 3.2, [[("Kto finansuje start", True, NAVY)], [(f"Partner Technologiczny {mln(ftp,1)}, Partner Marketplace {mln(fmp,1)} (proporcja 40 : 26)", False, TXT)],
    [("Grupa Azoty: 0 zł", True, GREEN)],
    [("Kolejność wypłat", True, NAVY)], [("1. Zysk najpierw zwraca partnerom dopłaty (lata 3-4)", False, TXT)], [("2. Potem dywidenda dla wszystkich: 34 / 40 / 26", False, TXT)],
    [("Bez CAPEX w spółce", True, NAVY)], [("Budowa to wkład partnerów, więc EBITDA = gotówka", False, TXT)]], 12)

# ===== 11. Przepływ pieniędzy =====
sl = frame("Przepływ pieniędzy w roku 5: kto ile dostaje", f"Rolnicy kupują za {mln(gmv[6])}. Do spółki GAia trafia {mln(rev[6],1)}, czyli {fmt(rev[6]/gmv[6]*100,1)}% obrotu.")
def node(x, y_, w, h, title, val, fill):
    box(sl, x, y_, w, h, fill=fill); tb(sl, x, y_, w, h, [[(title, False, WHITE)], [(val, True, WHITE)]], 11.5, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
def arrow(x, y_, w=0.35):
    a = sl.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(ty(y_)), Inches(w), Inches(th(y_, 0.25))); a.fill.solid(); a.fill.fore_color.rgb = PALE; a.line.fill.background()
node(0.6, 2.6, 1.9, 1.3, "Rolnicy płacą", mln(gmv[6]), NAVY); arrow(2.55, 3.1)
node(3.0, 1.3, 2.4, 0.85, "Nawozy GA", mln(fg[6]), MBLUE); node(3.0, 4.4, 2.4, 0.85, "Inne produkty (sprzedawcy)", mln(g3[6]), MBLUE)
tb(sl, 3.0, 2.25, 2.4, 2.0, [[("Z nawozów spółka bierze 1,2% prowizji, resztę zostawia GA.", False, GREY)], [("Ze sprzedaży innych produktów prowizja 7% i reklama 1,5%.", False, GREY)]], 10.5)
srcs = [("Prowizje 7%", JR("Prowizje: GMV 3P × 7%")[6]), ("Reklama 1,5%", JR("Reklama: GMV 3P × 1,5%")[6]), ("Nawozy GA 1,2%", JR("Prowizja od nawozów GA: GMV nawozów × 1,2%")[6]),
        ("Finansowanie", JR("Finansowanie: GMV × 0,4%")[6]), ("Premium", JR("Abonament premium: abonenci × 240 zł")[6]), ("Analizy", JR("Analizy dla producentów")[6])]
yy = 1.3
for lab, val in srcs:
    box(sl, 5.75, yy, 1.85, 0.56, line=NAVY); tb(sl, 5.75, yy, 1.85, 0.56, [[(lab + "  ", False, TXT), (fmt(val, 1), True, NAVY)]], 11, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE); yy += 0.66
arrow(7.65, 3.1); node(8.05, 2.45, 1.8, 1.6, "Spółka GAia: przychód", mln(rev[6], 1), NAVY); arrow(9.9, 3.1)
ptp = JR("Partner Technologiczny: utrzymanie i rozwój (koszt + 15%)")[6] + JR("Licencja: silnik Partnera Technologicznego (10 zł / użytkownik)")[6]
pmp = JR("Partner Marketplace: operacje (koszt + 15%)")[6] + JR("Licencja: marketplace Partnera Marketplace (1% GMV 3P)")[6]
other = opex[6] - ptp - pmp
outs = [("PT: usługi + licencja", ptp, MBLUE), ("PM: usługi + licencja", pmp, LBLUE), ("Płatności, marketing, zespół", other, GREY),
        ("Dywidenda → GA 34%", div[6] * 0.34, GREEN), ("Dywidenda → PT 40%", div[6] * 0.40, GREEN), ("Dywidenda → PM 26%", div[6] * 0.26, GREEN)]
yy = 1.3
for lab, val, col in outs:
    box(sl, 10.3, yy, 2.4, 0.56, fill=col); tb(sl, 10.3, yy, 2.4, 0.56, [[(lab + "  ", False, WHITE), (fmt(val, 1), True, WHITE)]], 10.5, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE); yy += 0.66
msg(sl, [("W roku 5: ", True, WHITE), (f"Partner Technologiczny {fmt(ptp,1)} + {fmt(div[6]*0.4,1)} mln zł, Partner Marketplace {fmt(pmp,1)} + {fmt(div[6]*0.26,1)} mln zł, GA {fmt(div[6]*0.34,1)} mln zł dywidendy + korzyści na nawozach.", False, WHITE)], y=6.3)

# ===== 12/13. Partnerzy =====
def partner_slide(Pd, title, src_labels, lic_txt, sh_txt):
    s0, s1, s2, s3, s4 = (Pd[x] for x in src_labels)
    tres = [r for kk, r in Pd.items() if kk.startswith("WYNIK PARTNERA")][0]; tcum = Pd["Wynik skumulowany"]
    cost = Pd["KOSZTY PARTNERA RAZEM (wkład + OPEX)"]
    nm = TP if "Techn" in title else MPk
    sl = frame(title, f"Wkład w roku 0-1: {mln(-s0[7],1)}. Zarobek w 5 lat: +{mln(tres[7],1)}, NPV {mln(k(nm)[8],1)}, zwrot w {k(nm)[10]}.")
    chart(sl, XL_CHART_TYPE.COLUMN_STACKED, 0.6, 1.2, 6.6, 3.4, Y05,
          [("0. Wkład własny", [round(v(s0, i), 1) for i in range(1, 7)]), ("1. Marża na usługach", [round(v(s1, i), 1) for i in range(1, 7)]),
           ("2. Licencja", [round(v(s2, i), 1) for i in range(1, 7)]), ("3. Dopłaty / zwrot", [round(v(s3, i), 1) for i in range(1, 7)]),
           ("4. Dywidenda", [round(v(s4, i), 1) for i in range(1, 7)])], [RED, MBLUE, LBLUE, AMBER, GREEN])
    data = [["mln zł", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma"],
            yrow("0. Wkład własny (praca w roku 0-1)", s0, color=sign), yrow("1. Marża 15% na usługach", s1, color=sign), yrow(f"2. Licencja: {lic_txt}", s2, color=sign),
            yrow("3. Dopłaty na start i ich zwrot", s3, color=sign), yrow(f"4. Dywidenda {sh_txt}", s4, color=sign), yrow("WYNIK PARTNERA", tres, color=sign), yrow("Skumulowany", tcum, color=sign, tot=False) + [""]]
    table(sl, 0.6, 4.7, 12.1, data, [4.6, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.5], rh=0.215, size=9.5, hl_rows=(6,))
    box(sl, 7.5, 1.2, 5.2, 3.4, line=NAVY)
    tb(sl, 7.6, 1.28, 5.0, 3.3, [
        [("Rok 0-1: wkład", True, RED)], [(f"Praca z własnych zasobów: {mln(-s0[7],1)}. Dopłaty na start spółki, które wracają w latach 3-4.", False, TXT)],
        [("Od roku 1: licencja " + ("za silnik" if "Techn" in title else "za marketplace"), True, MBLUE)], [(f"{lic_txt}: {mln(s2[7],1)} w 5 lat.", False, TXT)],
        [("Od roku 1: usługi", True, MBLUE)], [(f"Marża 15%: {mln(s1[7],1)} w 5 lat.", False, TXT)],
        [("Od roku 4: dywidenda", True, GREEN)], [(f"{sh_txt} zysku: {mln(s4[7],1)} (w roku 5: {mln(s4[6],1)}).", False, TXT)]], 12)
partner_slide(T, "Partner Technologiczny: z czego i kiedy zarabia",
              ["0. Wkład własny: dopracowanie GAia z własnych zasobów", "1. Marża 15% na usługach (lata 1-5)", "2. Licencja za silnik (10 zł / użytkownik / rok)", "3. Dopłaty na start spółki i ich zwrot", "4. Dywidenda: 40% zysku spółki"],
              "10 zł za użytkownika rocznie", "40%")
partner_slide(MP, "Partner Marketplace / Handel: z czego i kiedy zarabia",
              ["0. Wkład własny: dopracowanie GAia z własnych zasobów", "1. Marża 15% na usługach (lata 1-5)", "2. Licencja za marketplace (1% GMV 3P)", "3. Dopłaty na start spółki i ich zwrot", "4. Dywidenda: 26% zysku spółki"],
              "1% obrotu innych produktów", "26%")

# ===== 14. Grupa Azoty =====
gl = [("Udział w rynku AN/CAN", "Udział w rynku AN/CAN: p.p. × 39 mln zł × 25% marży"), ("Niższy koszt kanału", "Niższy koszt kanału: GMV nawozów × 2%"), ("Lepszy mix", "Poprawa mixu: GMV nawozów × 1%"),
      ("Prowizja dla spółki", "Prowizja dla spółki GAia od sprzedanych nawozów: × 1,2%"), ("Dywidenda 34%", "Dywidenda: 34% zysku spółki GAia")]
gres = G["WYNIK GRUPY AZOTY"]
sl = frame("Grupa Azoty: zysk bez wkładu pieniężnego", f"Wynik GA w roku 5: {mln(gres[6],1)}. W 5 lat: {mln(gres[7],1)}, NPV {mln(k(GA)[8],1)}. Na plusie od roku 1.")
pill(sl, 0.6, 1.25, 6.6, 0.42, "Składniki wyniku GA w roku 5, mln zł")
c = chart(sl, XL_CHART_TYPE.BAR_CLUSTERED, 0.6, 1.75, 6.6, 4.3, [a for a, _ in gl][::-1], [("Rok 5", [round(G[b][6], 1) for _, b in gl][::-1])], [NAVY], labels=True, numfmt='#,##0.0')
for i, (_, b) in enumerate(gl[::-1]):
    pt = c.series[0].points[i]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = GREEN if G[b][6] > 0 else RED
pill(sl, 7.5, 1.25, 5.2, 0.42, "Wynik GA rok po roku, mln zł")
chart(sl, XL_CHART_TYPE.COLUMN_CLUSTERED, 7.5, 1.75, 5.2, 4.3, Y05, [("Wynik GA", [round(x, 1) for x in gres[1:7]])], [NAVY], labels=True, numfmt='#,##0.0')
msg(sl, [("GA wnosi produkt, dane i ARiMR, a nie pieniądze. ", True, WHITE), ("Jak Yara i BASF zarabia na sprzedaży własnych produktów, a platforma tę sprzedaż zwiększa.", False, WHITE)], y=6.3)

# ===== 15. Każdy partner zyskuje =====
sl = frame("Każdy partner zyskuje", "Gotówka, mln zł. Wariant ostrożny: 150 tys. użytkowników w roku 5 zamiast 300 tys.")
data = [["Partner", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma 5 lat", "Zwrot od", "Przy 150 tys.\nużytkowników"]]
for nm in (GA, TP, MPk, JV):
    r, rs = k(nm), ks(nm)
    data.append([nm.replace(" - przepływy pieniężne", " (gotówka)")] + [(fmt(x, 1), sign(x)) for x in r[1:8]] + [r[10], (mln(rs[7], 1), sign(rs[7]))])
table(sl, 0.6, 1.3, 12.1, data, [3.3, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 1.2, 1.1, 1.7], rh=0.6, size=12.5, hl_rows=(1,))
tb(sl, 0.6, 4.6, 12.1, 1.3, [[("Uczciwie: ", True, NAVY), (f"przy połowie użytkowników partnerzy technologiczni wychodzą w 5 lat tylko lekko na plus (PT +{mln(ks(TP)[7],1)}, PM +{mln(ks(MPk)[7],1)}), a dopłaty wracają później. Dlatego w umowie: zespół rozwoju i marketing dopasowane do liczby użytkowników po pilocie.", False, TXT)]], 12.5)
msg(sl, [("Każdy zarabia na tym, co wnosi. ", True, WHITE), ("GA bez ryzyka finansowego, partnerzy z licencjami do własnych silników.", False, WHITE)], y=6.3)

# ===== 11. Lekcje i Polska =====
sl = frame("Czego uczą porażki i dlaczego GAia jest inna")
box(sl, 0.6, 1.2, 3.9, 4.35, line=RED); pill(sl, 0.6, 1.2, 3.9, 0.5, "Agrando, Niemcy", fill=RED)
tb(sl, 0.7, 1.8, 3.7, 3.7, ["• Marketplace nawozów, nasion, części", "• ok. 20 tys. rolników, 12 mln EUR finansowania (2021)", "• Zakończył działalność w 11.2022",
                           [("Lekcja: ", True, RED), ("bez własnych produktów, logistyki i dystrybutorów sam marketplace nie zarobi.", False, TXT)]], 12.5)
box(sl, 4.7, 1.2, 3.9, 4.35, line=AMBER); pill(sl, 4.7, 1.2, 3.9, 0.5, "DeHaat, Indie", fill=AMBER)
tb(sl, 4.8, 1.8, 3.7, 3.7, ["• Platforma: środki produkcji, skup, doradztwo", "• Przychód ok. 2 700 cr INR (FY24), +36-40% r/r", "• Strata operacyjna ok. 245 cr INR (FY24)",
                           [("Lekcja: ", True, AMBER), ("skala przychodzi szybko, ale model z własnym magazynem i skupem długo traci.", False, TXT)]], 12.5)
box(sl, 8.8, 1.2, 3.9, 4.35, line=GREEN); pill(sl, 8.8, 1.2, 3.9, 0.5, "GAia, Polska", fill=GREEN)
tb(sl, 8.9, 1.8, 3.7, 3.7, ["• Producent, produkcja, logistyka i dane ARiMR", "• Sieć sprzedaży GA i punkty wydania, bez nowych magazynów",
                           "• Gotowe silniki partnerów i 5 funkcjonalności GA", "• 1,21 mln rolników w ARiMR, brak lidera platform zakupowych",
                           [("Wniosek: ", True, GREEN), ("mamy to, czego zabrakło Agrando, bez kosztów DeHaat.", False, TXT)]], 12.5)
tb(sl, 0.6, 5.62, 12.1, 0.6, "Źródła: agfundernews.com, CB Insights (Agrando); entrackr.com, inc42.com (DeHaat); ARiMR 2024 via wiescirolnicze.pl. Dla Partnera Marketplacowego: ERLI urósł z 1,23 do ok. 1,8 mld zł GMV (2024→2025).", 10, False, GREY)
msg(sl, [("Ryzyko jest w wykonaniu, nie w modelu: ", True, WHITE), ("musimy dowieźć adopcję rolników. Dlatego pilot przed skalowaniem.", False, WHITE)], y=6.3)



# ===== 17. Ryzyka i kroki =====
sl = frame("Ryzyka, warunki umów i dalsze kroki")
data = [["Ryzyko", "Co robimy"],
        ["Za mało użytkowników", "Pilot przed skalowaniem. Zespół rozwoju i marketing dopasowane do adopcji (w umowie)."],
        ["Spór między wspólnikami", "Sprawy zastrzeżone jednomyślnie, GA 34% blokuje zmiany umowy spółki, eskalacja i mediacja."],
        ["Jakość 5 funkcjonalności GA", "Audyt na Azure w pierwszych 30 dniach, potem podłączenie przez API."],
        ["Uzależnienie od technologii", "Kod stworzony w spółce i dane rolników należą do GAia. Licencje na silnik i marketplace długoterminowe."]]
table(sl, 0.6, 1.2, 6.4, data, [2.0, 4.4], rh=0.72, size=11.5)
data = [["Termin", "Krok"], ["do 30.11", "Audyt 5 funkcjonalności GA na Azure, potwierdzenie danych"],
        ["do 15.12", "Term sheet z Partnerem Technologicznym: 40%, licencja 10 zł, koszt + 15%"],
        ["do 15.12", "Term sheet z Partnerem Marketplace: 26%, plug-in, licencja 1%"],
        ["01.2027", "Umowa wspólników, start prac (rok 0)"],
        ["02-04.2027", "Pilot w sezonie: 1 województwo, cel 10 tys. użytkowników"]]
table(sl, 7.2, 1.2, 5.5, data, [1.3, 4.2], rh=0.6, size=11.5)
msg(sl, [("Prosimy o: ", True, WHITE), ("zgodę na term sheety z obydwoma partnerami. GA nie angażuje gotówki.", False, WHITE)], y=6.3)

prs.save("../GA_GAia_business_case_300k.pptx"); print("ok", N[0], "slajdów")
