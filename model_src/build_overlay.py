# Nakłada liczby z modelu na oryginalny slajd "Scenariusz po rezygnacji ze współpracy Allegro" (30.10)
# i dokleja 4 slajdy z GA_GAia_jak_zarabiamy_z_partnerem.pptx jako uzasadnienie.
# Użycie: python3 build_overlay.py <przeliczony.xlsx> <przeliczony_stres.xlsx>
import sys, copy, openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

def summary(p):
    return {r[0].split(":")[0]: r for r in openpyxl.load_workbook(p, data_only=True)["Podsumowanie"].iter_rows(min_row=5, max_row=8, values_only=True)}
S, ST = summary(sys.argv[1]), summary(sys.argv[2])
f = lambda v: f"{v:,.0f}".replace(",", " ")

RED = RGBColor(0xFF, 0x00, 0x00); GREEN = RGBColor(0x00, 0x8A, 0x3E); AMBER = RGBColor(0xE6, 0x7E, 0x00)
NAVY = RGBColor(0x00, 0x36, 0x65); WHITE = RGBColor(0xFF, 0xFF, 0xFF); TXT = RGBColor(0x1A, 0x1A, 0x1A)

prs = Presentation("../GA_GAia_jak_zarabiamy_z_partnerem.pptx")
sl = prs.slides.add_slide(prs.slide_layouts[6])
sl.shapes.add_picture("slajd_30Oct.png", 0, 0, prs.slide_width, prs.slide_height)

def box(x, y, w, h, fill, line=None, rounded=True):
    s = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    if rounded: s.adjustments[0] = 0.12
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line: s.line.color.rgb = line; s.line.width = Pt(1)
    else: s.line.fill.background()
    s.shadow.inherit = False
    return s

def text(x, y, w, h, paras, size=10.5, anchor=MSO_ANCHOR.MIDDLE):
    t = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
    t.word_wrap = True; t.vertical_anchor = anchor
    t.margin_left = t.margin_right = Inches(0.07); t.margin_top = t.margin_bottom = Inches(0.02)
    for i, runs in enumerate(paras):
        p = t.paragraphs[0] if i == 0 else t.add_paragraph(); p.space_after = Pt(1)
        for s, b, c in runs:
            r = p.add_run(); r.text = s; r.font.size = Pt(size); r.font.bold = b; r.font.color.rgb = c; r.font.name = "Calibri"

s1, s1a, s1b, s2, s2s = S["S1"], S["S1a"], S["S1b"], S["S2"], ST["S2"]

# Scenariusz 1: ryzyko + wynik GA
box(6.08, 4.42, 2.62, 0.56, RED)
text(6.08, 4.42, 2.62, 0.56, [[("Ryzyko: operacja w GA (25 FTE)", False, WHITE)],
                              [(f"Wynik GA: +{f(s1[5])} mln zł / 5 lat", True, WHITE)]], 10.5)
# Scenariusz 1 alternatywnie (Mirakl): zamiast "trwale nieprofitowy"
box(6.11, 6.04, 2.62, 0.56, AMBER)
text(6.11, 6.04, 2.62, 0.56, [[("Fee od całego GMV: ", False, WHITE), (f"GA +{f(s1a[5])} mln zł", True, WHITE)],
                              [("Fee tylko od sprzedawców: ", False, WHITE), (f"GA +{f(s1b[5])} mln zł", True, WHITE)]], 10)
# Scenariusz 2: ryzyka z odpowiedzią
box(9.70, 3.64, 2.62, 1.08, RED)
text(9.70, 3.64, 2.62, 1.08, [[("Główne ryzyko:", True, WHITE)],
                              [("• Koordynacja 2 partnerów → OLX w umowie revenue share, nie wspólnik JV", False, WHITE)],
                              [(f"• Wolna adopcja → przy 50% planu GA nadal +{f(s2s[5])} mln zł", False, WHITE)]], 10)
# Mitygacja + wynik
box(9.70, 4.92, 2.96, 1.57, WHITE, line=TXT)
box(9.70, 4.92, 2.96, 0.55, GREEN)
text(9.70, 4.92, 2.96, 0.55, [[(f"Wynik GA: +{f(s2[5])} mln zł / 5 lat", True, WHITE)],
                              [(f"NPV {f(s2[6])} mln zł, zwrot w {s2[7].split()[-1]}. roku", False, WHITE)]], 11)
text(9.70, 5.47, 2.96, 1.02, [[("Skąd zysk GA:", True, TXT)],
                              [("• marża dystrybutora zostaje w GA (+3,6 zł/100 zł)", False, TXT)],
                              [("• +2 p.p. AN/CAN, lepszy mix nawozów", False, TXT)],
                              [("• 51% JV: prowizje 3P, abonamenty premium", False, TXT)],
                              [("• w umowie opcja wykupu Iteo", False, TXT)]], 9.5, MSO_ANCHOR.TOP)
# stopka w pasie
text(0.3, 6.78, 10.5, 0.6, [[("Liczby: model GA_GAia_model_JV_partnerzy.xlsx, wynik Grupy (udział w JV + marża kanałowa + udział w rynku). Ilustracyjne, do potwierdzenia danymi Handlu i ofertami partnerów. Szczegóły na kolejnych slajdach.", False, WHITE)]], 9.5)

# przenieś nowy slajd na początek i przenumeruj stopki
lst = prs.slides._sldIdLst; el = lst[-1]; lst.remove(el); lst.insert(0, el)
for n, s in enumerate(prs.slides, start=1):
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.startswith("GAia | model rozliczeń"):
            r = sh.text_frame.paragraphs[0].runs[-1]
            r.text = r.text.rstrip("0123456789") + str(n)
prs.save("../GA_GAia_update_30Oct_z_danymi.pptx")
print("ok")
