# Łączy Waszą prezentację (slajd 30.10, z nałożonymi liczbami z business case) z 12 slajdami business case.
# Użycie: python3 build_overlay_bc.py <przeliczony.xlsx> <przeliczony_stres.xlsx>
import sys, openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
def rows(p):
    d = {}
    for r in openpyxl.load_workbook(p, data_only=True)["Każdy partner zyskuje"].iter_rows(values_only=True):
        if r[0] and r[0] not in d: d[r[0]] = r
    return d
KZ, KS = rows(sys.argv[1]), rows(sys.argv[2])
f = lambda v, d=0: f"{v:,.{d}f}".replace(",", " ").replace(".", ",")
FONT = "Trebuchet MS"
RED = RGBColor(0xFF, 0, 0); GREEN = RGBColor(0x00, 0x8A, 0x3E); AMBER = RGBColor(0xE6, 0x7E, 0x00)
NAVY = RGBColor(0x00, 0x36, 0x65); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
prs = Presentation("../GA_GAia_business_case_300k.pptx")
sl = prs.slides.add_slide(prs.slide_layouts[6])
sl.shapes.add_picture("slajd_30Oct.png", 0, 0, prs.slide_width, prs.slide_height)
def box(x, y, w, h, fill, line=None):
    s = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)); s.adjustments[0] = 0.12
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line: s.line.color.rgb = line; s.line.width = Pt(1)
    else: s.line.fill.background()
    s.shadow.inherit = False
def text(x, y, w, h, paras, size, anchor=MSO_ANCHOR.MIDDLE):
    t = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
    t.word_wrap = True; t.vertical_anchor = anchor; t.margin_left = t.margin_right = Inches(0.07); t.margin_top = t.margin_bottom = Inches(0.02)
    for i, runs in enumerate(paras):
        p = t.paragraphs[0] if i == 0 else t.add_paragraph(); p.space_after = Pt(1)
        for s_, b, c in runs:
            r = p.add_run(); r.text = s_; r.font.size = Pt(size); r.font.bold = b; r.font.color.rgb = c; r.font.name = FONT
ga, gas = KZ["Grupa Azoty (51%)"], KS["Grupa Azoty (51%)"]
# wariant Mirakl: zamiast "model trwale nieprofitowy"
box(6.11, 6.04, 2.62, 0.56, AMBER)
text(6.11, 6.04, 2.62, 0.56, [[("Opłata tylko od sprzedawców", True, WHITE)], [("zewn., nie od nawozów GA", True, WHITE)]], 9.5)
# Scenariusz 2: ryzyka z odpowiedzią
box(9.70, 3.64, 2.62, 1.08, RED)
text(9.70, 3.64, 2.62, 1.08, [[("Główne ryzyko:", True, WHITE)],
     [("• 2 partnerów → GA 51%, prezes z GA, wykup", False, WHITE)],
     [(f"• Wolna adopcja → przy 150 tys. GA +{f(gas[7])} mln zł", False, WHITE)]], 9)
# Mitygacja + wynik z business case
box(9.70, 4.92, 2.96, 1.57, WHITE, line=NAVY)
box(9.70, 4.92, 2.96, 0.55, GREEN)
text(9.70, 4.92, 2.96, 0.55, [[(f"GA: +{f(ga[7],1)} mln zł / 5 lat", True, WHITE)], [(f"NPV {f(ga[8])} mln zł, zwrot w {ga[10].split()[-1]}. roku", False, WHITE)]], 10.5)
text(9.70, 5.47, 2.96, 1.02, [[("300 tys. użytkowników w roku 5", True, NAVY)],
     [("• każdy z 3 partnerów na plusie (dalej)", False, NAVY)], [("• premium 20 zł/mies., lepszy mix nawozów", False, NAVY)],
     [("• opcja wykupu Partnera Technologicznego", False, NAVY)]], 9, MSO_ANCHOR.TOP)
text(0.3, 6.8, 10.8, 0.55, [[("Liczby: business case GA_GAia_business_case_300k.xlsx (ilustracyjne, do potwierdzenia). Szczegóły na kolejnych slajdach.", False, WHITE)]], 9)
lst = prs.slides._sldIdLst; el = lst[-1]; lst.remove(el); lst.insert(0, el)
for n, s in enumerate(prs.slides, start=1):
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.startswith("GAia | business case"):
            r = sh.text_frame.paragraphs[0].runs[-1]; r.text = r.text.rstrip("0123456789") + str(n)
prs.save("../GA_GAia_prezentacja_business_case.pptx"); print("ok", len(prs.slides))
