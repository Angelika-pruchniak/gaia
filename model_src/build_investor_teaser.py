# 3 slajdy inwestorskie na szablonie GAia teaser: kluczowe założenia, EBITDA, inwestycja i zwrot. Dane: partner_variants_v5.json.
# Użycie: python3 build_investor_teaser.py   → ../GAia_teaser_inwestor_3_slajdy.pptx
import json, copy, io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION

V = json.load(open("partner_variants_v5.json", encoding="utf-8")); D = V["deal"]; P = D["pool"]
TPL = "szablon_gaia_teaser.pptx"
prs = Presentation(TPL)
bg = list(prs.slides)[1].shapes[0].image.blob                       # tło z szablonu (slajd 2)
sld = prs.slides._sldIdLst                                           # usuń slajdy szablonu, zostaw master i motyw
for s in list(sld):
    prs.part.drop_rel(s.rId); sld.remove(s)
blank = [l for l in prs.slide_layouts if l.name == "BLANK"][0]

WHITE = RGBColor(0xFF, 0xFF, 0xFF); GREY = RGBColor(0x9C, 0xA3, 0xAF); LGREY = RGBColor(0xD1, 0xD5, 0xDB); DGREY = RGBColor(0x6B, 0x72, 0x80)
GREEN = RGBColor(0x10, 0xB9, 0x81); LGREEN = RGBColor(0x34, 0xD3, 0x99); RED = RGBColor(0xEF, 0x44, 0x44); AMBER = RGBColor(0xF5, 0x9E, 0x0B)
CARD = RGBColor(0x0E, 0x24, 0x1B); CARDL = RGBColor(0x1E, 0x4D, 0x3A); DARKG = RGBColor(0x1F, 0x6F, 0x54)
TITLE_F = "Plus Jakarta Sans"; BODY_F = "Inter"
fmt = lambda v, d=0: f"{v:,.{d}f}".replace(",", " ").replace(".", ",").replace("-", "−")
CAL = [str(2027 + i) for i in range(8)]

def slide(title):
    sl = prs.slides.add_slide(blank)
    sl.shapes.add_picture(io.BytesIO(bg), 0, 0, prs.slide_width, prs.slide_height)
    text(sl, 0.64, 0.55, 12.0, 0.6, title, 28.5, WHITE, TITLE_F)
    return sl
def text(sl, x, y, w, h, t, size, color=GREY, font=BODY_F, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tf = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor; tf.margin_left = tf.margin_right = Inches(0.04); tf.margin_top = tf.margin_bottom = Inches(0.02)
    paras = t if isinstance(t, list) else [[(t, color, bold)]]
    for i, runs in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); p.alignment = align; p.space_after = Pt(3)
        for s_, c_, b_ in runs:
            r = p.add_run(); r.text = s_; r.font.size = Pt(size); r.font.color.rgb = c_; r.font.bold = b_; r.font.name = font
    return tf
def card(sl, x, y, w, h, accent=None):
    s = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)); s.adjustments[0] = 0.06
    s.fill.solid(); s.fill.fore_color.rgb = CARD; s.line.color.rgb = accent or CARDL; s.line.width = Pt(1.25 if accent else 0.75); s.shadow.inherit = False
    return s
def kpi(sl, x, y, w, h, label, big, sub, color=WHITE, accent=None):
    card(sl, x, y, w, h, accent)
    text(sl, x + 0.18, y + 0.14, w - 0.3, 0.25, label.upper(), 9.5, DGREY, "Inter SemiBold")
    text(sl, x + 0.18, y + 0.4, w - 0.3, 0.5, big, 22 if len(big) < 12 else 18, color, TITLE_F, True)
    text(sl, x + 0.18, y + h - 0.5, w - 0.3, 0.45, sub, 9.5, GREY)
def dark_chart(sl, kind, x, y, w, h, cats, series, colors, labels=True):
    cd = CategoryChartData(); cd.categories = cats
    for nm, vals in series: cd.add_series(nm, vals)
    ch = sl.shapes.add_chart(kind, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    ch.font.name = BODY_F; ch.font.size = Pt(10); ch.font.color.rgb = GREY
    ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.BOTTOM; ch.legend.include_in_layout = False; ch.legend.font.color.rgb = LGREY
    for s, c in zip(ch.series, colors):
        s.format.fill.solid(); s.format.fill.fore_color.rgb = c; s.invert_if_negative = False
    pl = ch.plots[0]; pl.gap_width = 60
    if labels:
        pl.has_data_labels = True; dl = pl.data_labels; dl.font.size = Pt(10); dl.font.color.rgb = WHITE; dl.number_format = '#,##0'; dl.number_format_is_linked = False
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
    va = ch.value_axis; va.has_major_gridlines = True; va.major_gridlines.format.line.color.rgb = RGBColor(0x1F, 0x3A, 0x2E); va.format.line.fill.background()
    va.tick_labels.font.color.rgb = DGREY; va.tick_labels.number_format = '#,##0'; va.tick_labels.number_format_is_linked = False
    ca = ch.category_axis; ca.format.line.color.rgb = RGBColor(0x2E, 0x4A, 0x3E); ca.tick_labels.font.color.rgb = LGREY; ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    return ch
def footer(sl, t):
    text(sl, 0.64, 6.95, 12.0, 0.3, t, 8.5, DGREY)

# ===== 1. Kluczowe założenia =====
sl = slide("Kluczowe założenia: liczby z P&L Grupy Azoty")
text(sl, 0.64, 1.3, 5.6, 0.3, "GMV PLATFORMY W 2032", 10.5, GREEN, "Inter SemiBold")
text(sl, 0.64, 1.6, 5.6, 0.9, f"{fmt(D['gmv'][5]/1000, 2)} mld zł", 48, GREEN, "Inter ExtraBold", True)
text(sl, 0.64, 2.5, 5.6, 1.0, [[("= 850 tys. rolników × 35% × 75% kupujących × 58 784 zł × 19% udziału w wydatkach", LGREY, False)],
     [("ok. 5% z 50 mld zł, które polscy rolnicy wydają rocznie na środki produkcji", GREY, False)]], 12)
card(sl, 0.64, 3.65, 5.6, 2.95)
text(sl, 0.84, 3.78, 5.2, 0.3, "DLACZEGO TO REALNE", 10, DGREY, "Inter SemiBold")
text(sl, 0.84, 4.1, 5.25, 2.45, [[("Popyt od 1. dnia: ", WHITE, True), ("66 dystrybutorów, marka i produkty Grupy Azoty, nawozy.eu, dane ARiMR.", GREY, False)],
     [("Koszty ≥ P&L GA: ", WHITE, True), ("300,6 vs 293,1 mln zł w latach 2028-2033.", GREY, False)],
     [("Konserwatywnie: ", WHITE, True), ("premium 20 zł/ha (GA: 40), bez synergii, bez sprzedaży inkrementalnej, back margin bez nawozów GA.", GREY, False)],
     [("Gotowe silniki: ", WHITE, True), ("MVP1 za 5,6 mln zł zamiast budowy od zera.", GREY, False)]], 11.5)
K = [("Użytkownicy 2032", "297,5 tys.", "35% aktywnych rolników · GA w.25"), ("Udział w wydatkach", "5% → 19%", "top e-commerce 20-24% · GA w.31"),
     ("Take rate", f"{fmt(P['rev'][5]/D['gmv'][5]*100, 1)}% GMV", "prowizje 1-4% + back margin · GA: 5,3%"),
     ("CAC", "250 → 350 zł", "rośnie ze skalą · GA w.69"), ("Churn", "15% × 350 zł", "B2B; 52,5 zł/użytk./rok · GA B72"),
     ("Premium", "20 zł/ha", "≈ 330 zł/rok; ekoschemat ok. 87,5 zł/ha")]
for i, (lab, big, sub) in enumerate(K):
    x = 6.55 + (i % 2) * 3.1; y = 1.3 + (i // 2) * 1.8
    kpi(sl, x, y, 2.95, 1.65, lab, big, sub, WHITE if i else LGREEN, GREEN if i == 1 else None)
footer(sl, "Źródło: GAia-model_PL_v2 (Grupa Azoty) z korektami; model GA_GAia_model_wspolpracy_v5.xlsx. Zielona ramka = założenie o największym wpływie na wynik.")

# ===== 2. EBITDA =====
sl = slide("EBITDA dodatnia od 2031, 90 mln zł w 2034")
dark_chart(sl, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.5, 1.3, 7.9, 4.9, CAL[1:],
           [("Przychody puli, mln zł", [round(x, 1) for x in P["rev"][1:]]), ("EBITDA, mln zł", [round(x, 1) for x in P["net"][1:]])], [DARKG, LGREEN])
m34 = P["net"][7] / P["rev"][7]; rpu = P["rev"][7] * 1e6 / (D["users"][7] * 1000)
kpi(sl, 8.65, 1.3, 4.05, 1.35, "Marża EBITDA 2034", f"{round(m34*100)}%", f"2032: {round(P['net'][5]/P['rev'][5]*100)}%. Benchmark platform 35-45% (GA)", LGREEN, GREEN)
kpi(sl, 8.65, 2.8, 4.05, 1.35, "Przychód na użytkownika 2034", f"{fmt(rpu)} zł/rok", "vs CAC 350 zł: zwrot pozyskania w niecały rok", WHITE)
kpi(sl, 8.65, 4.3, 4.05, 1.35, "Przychody 2034", f"{fmt(P['rev'][7], 1)} mln zł", f"prowizje {round(P['comm'][7]/P['rev'][7]*100)}% · back margin {round(P['back'][7]/P['rev'][7]*100)}% · premium {round(P['sub'][7]/P['rev'][7]*100)}%", WHITE)
text(sl, 0.64, 6.3, 12.0, 0.5, [[("Koszty zawierają 10% marży partnerów technologicznych i 0,8 mln zł/rok na bezpieczeństwo i ład umowy. ", GREY, False),
     ("Największa pozycja to pozyskanie i utrzymanie rolników (44% kosztów).", LGREY, True)]], 11)
footer(sl, "EBITDA = przychody wspólnej puli − koszty wspólne (rozwój, utrzymanie, zespół, CAC, antichurn, bezpieczeństwo, marża partnerów). Gotówka, bez inflacji. Rok 0 = 2027 (MVP1 + pilot).")

# ===== 3. Inwestycja i zwrot =====
sl = slide("58 mln zł finansowania, 3,4× zwrotu do 2034")
neg = -sum(x for x in P["net"] if x < 0); pos = sum(x for x in P["net"] if x > 0)
kpi(sl, 0.64, 1.3, 2.9, 1.5, "Finansowanie startu", f"{fmt(P['need'], 1)} mln zł", "najgłębszy deficyt (2030); w transzach po bramkach", AMBER)
kpi(sl, 3.69, 1.3, 2.9, 1.5, "Skumulowana EBITDA 2034", f"+{fmt(P['cum'][7], 1)} mln zł", "po pokryciu całego deficytu", LGREEN, GREEN)
kpi(sl, 6.74, 1.3, 2.9, 1.5, "Zwrot z finansowania", f"{fmt(pos/neg, 1)}×", f"{fmt(pos, 1)} mln zł EBITDA na {fmt(neg, 1)} mln zł deficytu", WHITE)
kpi(sl, 9.79, 1.3, 2.9, 1.5, "Break-even", "2031 / 2033", "EBITDA dodatnia / pełny zwrot narastająco", WHITE)
S = [("Bazowy", "deal", LGREY), ("Ambicja 20% rolników (zamiast 35%)", "deal_adopt20", LGREY), ("CAC +100 zł", "deal_cac100", LGREY),
     ("Udział w wydatkach −30%", "deal_wal70", RED), ("Synergie partnerów potwierdzone ofertami", "deal_syn", LGREEN)]
card(sl, 0.64, 3.0, 7.6, 3.55)
text(sl, 0.84, 3.1, 7.2, 0.3, "TEST ODPORNOŚCI: PEŁNE PRZELICZENIE MODELU", 10, DGREY, "Inter SemiBold")
hdrs = ["Wariant", "EBITDA 2034", "Narastająco 2034", "Finansowanie"]; xs = [0.84, 4.0, 5.35, 6.6]
for j, h in enumerate(hdrs): text(sl, xs[j], 3.45, 1.4 if j else 3.4, 0.3, h, 9.5, DGREY, "Inter SemiBold", align=PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT)
for k, (lab, key, col) in enumerate(S):
    y = 3.8 + k * 0.52; v = V[key]["pool"]
    ln = sl.shapes.add_connector(1, Inches(0.84), Inches(y - 0.05), Inches(8.05), Inches(y - 0.05)); ln.line.color.rgb = CARDL; ln.line.width = Pt(0.5)
    text(sl, 0.84, y, 3.4, 0.4, lab, 11, col, BODY_F, k == 0)
    for j, val in enumerate([v["net"][7], v["cum"][7], v["need"]]):
        text(sl, xs[j + 1], y, 1.4 if j < 2 else 1.1, 0.4, fmt(val, 1), 11, RED if val < 0 else (col if j < 2 else LGREY), BODY_F, k == 0, PP_ALIGN.RIGHT)
card(sl, 8.45, 3.0, 4.25, 3.55, GREEN)
text(sl, 8.65, 3.1, 3.9, 0.3, "RYZYKO POD KONTROLĄ", 10, GREEN, "Inter SemiBold")
text(sl, 8.65, 3.45, 3.95, 3.0, [[("Bramka 1 (pilot 2027): ", WHITE, True), ("≥ 10 tys. użytkowników, CAC ≤ 250 zł, test ceny premium.", GREY, False)],
     [("Bramka 2 (koniec 2029): ", WHITE, True), ("≥ 7% rolników, udział w wydatkach ≥ 10%, pula ≥ 10 mln zł/rok.", GREY, False)],
     [("Finansowanie w transzach: ", WHITE, True), ("kolejna transza tylko po spełnieniu bramki.", GREY, False)],
     [("Grupa Azoty nie wnosi gotówki: ", WHITE, True), ("wnosi popyt, markę i dystrybucję.", GREY, False)]], 11)
footer(sl, "Zwrot z finansowania = suma dodatniej EBITDA 2031-2034 ÷ suma deficytu 2028-2030. Model: GA_GAia_model_wspolpracy_v5.xlsx, warianty: partner_variants_v5.py.")
prs.save("../GAia_teaser_inwestor_3_slajdy.pptx"); print("ok", len(prs.slides))
