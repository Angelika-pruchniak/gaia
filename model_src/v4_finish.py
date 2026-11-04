# Wykonywany na końcu build_bc_xlsx.py (exec): opisy formuł, komentarze, arkusz danych do tabel
# przestawnych, spis treści i profesjonalne formatowanie (bez siatki, linie tylko w tabelach).
from openpyxl.comments import Comment
from openpyxl.styles import Border, Side, Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo

FONT = "Trebuchet MS"
NAVY_ = "003665"; GRID = "BFBFBF"; LINK = "00704A"; GREYT = "595959"
thin = Side(style="thin", color=GRID); navy_thin = Side(style="thin", color=NAVY_); navy_med = Side(style="medium", color=NAVY_)
NUMF = '#,##0.0;[Red]-#,##0.0;"–"'

# ---------------------------------------------------------------- 4. Formatowanie
def font_of(c, **kw):
    f = c.font
    base = dict(name=FONT, size=kw.get("size", f.size if f and f.size and f.size != 11 else 10), bold=f.bold, italic=f.italic,
                color=f.color, underline=f.underline)
    base.update(kw); return Font(**base)

def border_table(ws, r1, r2, c1, c2, header=True):
    for r in range(r1, r2 + 1):
        vals = [ws.cell(r, c).value for c in range(c1, c2 + 1)]
        if all(v is None for v in vals): continue
        only_a = vals[0] is not None and all(v is None for v in vals[1:])
        is_head = header and r == r1
        bold = bool(ws.cell(r, c1).font and ws.cell(r, c1).font.bold)
        f1 = ws.cell(r, c1).font
        is_sec = only_a and (not bold or (f1.color is not None and f1.size and f1.size >= 12))
        for c in range(c1, c2 + 1):
            ce = ws.cell(r, c)
            if is_head:
                ce.font = Font(name=FONT, bold=True, color="FFFFFF", size=10); ce.fill = PatternFill("solid", fgColor=NAVY_)
                ce.alignment = Alignment(horizontal="center" if c > c1 else "left", vertical="center", wrap_text=True)
                ce.border = Border(left=navy_thin, right=navy_thin, top=navy_thin, bottom=navy_thin)
            elif is_sec:
                ce.border = Border(bottom=navy_med)
                if c == c1: ce.font = Font(name=FONT, bold=True, color=NAVY_, size=11)
            else:
                ce.border = Border(left=thin, right=thin, top=navy_thin if bold else thin, bottom=navy_thin if bold else thin)
                if c > c1 and ce.alignment.horizontal is None: ce.alignment = Alignment(horizontal="right", vertical="center")
                if c == c1 or isinstance(ce.value, str) and not str(ce.value).startswith("="):
                    ce.alignment = Alignment(horizontal="left", vertical="center", wrap_text=ce.alignment.wrap_text)

def finish(ws, regions, title_row=1, sub_row=2, tab=None, land=True):
    ws.sheet_view.showGridLines = False
    if tab: ws.sheet_properties.tabColor = tab
    for row in ws.iter_rows():
        for ce in row:
            if ce.value is None and not ce.fill: continue
            v = ce.value
            if isinstance(v, str) and v.startswith("=") and "!" in v and not (ce.font and ce.font.bold) and ce.font.color is None:
                ce.font = font_of(ce, color=LINK)
            else:
                ce.font = font_of(ce)
            if ce.number_format == NUM: ce.number_format = NUMF
    ws.cell(title_row, 1).font = Font(name=FONT, bold=True, size=16, color=NAVY_)
    if ws.cell(sub_row, 1).value: ws.cell(sub_row, 1).font = Font(name=FONT, italic=True, size=10, color=GREYT)
    for reg in regions: border_table(ws, *reg)
    ws.page_setup.orientation = "landscape" if land else "portrait"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.print_options.horizontalCentered = True


# ======== v3: opisy, dane do tabel przestawnych, spis treści, formatowanie
CALC = [(U, "2E6DB4"), (Kp, "E67E00"), (J, "003665"), (G, "003665"), (T, "2E6DB4"), (Mp, "2E6DB4")]
for sh, _t in CALC:
    for key, text in sh.notes.items():
        if key in sh.R:
            r = sh.R[key]; sh.ws.cell(r, NOTEC, text); sh.ws.cell(r, 1).comment = Comment(text, "Model GAia", width=320, height=90)
for row in A.iter_rows(min_row=5, max_row=A.max_row):
    lab, src = row[0].value, row[2 + NY].value
    if lab and src and row[1].value is not None:
        row[1].comment = Comment(f"{lab}\nŹródło / uwaga: {src}", "Model GAia", width=320, height=110)
DS = wb.create_sheet("Dane")
DS["A1"] = "Dane źródłowe do tabel przestawnych (formuły z arkuszy obliczeń)"
DS["A2"] = "Każdy wiersz = pozycja w roku. Tabele przestawne odświeżają się przy otwarciu (Dane → Odśwież wszystko)."
for j, h in enumerate(["Obszar", "Typ", "Pozycja", "Rok", "Wartość (mln zł)", "Arkusz źródłowy"]): DS.cell(4, 1 + j, h)
SPEC = [("Grupa Azoty", "Wynik partnera", "Grupa Azoty", G, "res"), ("Partner Technologiczny", "Wynik partnera", "Partner Technologiczny", T, "res"),
        ("Partner Marketplace", "Wynik partnera", "Partner Marketplace", Mp, "res"),
        ("Pula", "Przychody puli", "1 Prowizje", J, "comm"), ("Pula", "Przychody puli", "2 Back margin", J, "back"), ("Pula", "Przychody puli", "3 Premium", J, "sub"), ("Pula", "Przychody puli", "4 Opłata za sukces GA", J, "succ"),
        ("Pula", "Koszty: model GA", "1 Rozwój", Kp, "dev"), ("Pula", "Koszty: model GA", "2 Utrzymanie", Kp, "mt"), ("Pula", "Koszty: model GA", "3 Zespół", Kp, "tm"), ("Pula", "Koszty: model GA", "4 CAC", Kp, "cac"), ("Pula", "Koszty: model GA", "5 Antichurn", Kp, "ach"),
        ("Pula", "Koszty: partnerstwo", "1 Rozwój", Kp, "pdev"), ("Pula", "Koszty: partnerstwo", "2 Utrzymanie", Kp, "pmt"), ("Pula", "Koszty: partnerstwo", "3 Zespół", Kp, "ptm"), ("Pula", "Koszty: partnerstwo", "4 Pozyskanie i utrzymanie", Kp, "pcac"),
        ("Pula", "Wynik i podział puli", "1 EBITDA platformy", J, "net"), ("Pula", "Wynik i podział puli", "2 Deficyt", J, "fund"), ("Pula", "Wynik i podział puli", "3 Zwrot deficytu", J, "rep"), ("Pula", "Wynik i podział puli", "4 Do podziału", J, "dist"),
        ("Partner Technologiczny", "Zarobek PT", "0 Wkład własny", T, "own"), ("Partner Technologiczny", "Zarobek PT", "1 Deficyt i zwrot", T, "fin"), ("Partner Technologiczny", "Zarobek PT", "2 Udział w podziale", T, "dist"),
        ("Partner Marketplace", "Zarobek PM", "0 Wkład własny", Mp, "own"), ("Partner Marketplace", "Zarobek PM", "1 Deficyt i zwrot", Mp, "fin"), ("Partner Marketplace", "Zarobek PM", "2 Udział w podziale", Mp, "dist"),
        ("Grupa Azoty", "Zarobek GA", "1 Dodatkowa marża w GA", G, "keep"), ("Grupa Azoty", "Zarobek GA", "2 Udział w podziale", G, "dist"), ("Grupa Azoty", "Zarobek GA", "3 Koszt wewnętrzny", G, "int"),
        ("Platforma", "GMV", "1 Nawozy GA", U, "fga"), ("Platforma", "GMV", "2 Pozostałe", U, "gmv")]
dr = 5
for area, typ, pos, sh, key in SPEC:
    for yi, col in enumerate(COLS):
        f_ = f"=N('{sh.ws.title}'!{col}{sh.R[key]})" if not (key == "gmv" and pos == "2 Pozostałe") else f"=N('{sh.ws.title}'!{col}{sh.R['gmv']})-N('{sh.ws.title}'!{col}{sh.R['fga']})"
        DS.cell(dr, 1, area); DS.cell(dr, 2, typ); DS.cell(dr, 3, pos); DS.cell(dr, 4, f"Rok {yi}"); DS.cell(dr, 5, f_).number_format = NUMF; DS.cell(dr, 6, sh.ws.title); dr += 1
tab = Table(displayName="DaneGAia", ref=f"A4:F{dr-1}"); tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True); DS.add_table(tab)
for col, w in zip("ABCDEF", [24, 30, 36, 9, 16, 26]): DS.column_dimensions[col].width = w
DS.freeze_panes = "A5"
TOC = wb.create_sheet("Spis treści", 0)
TOC["A1"] = "GAia v3: model współpracy trzech partnerów na bazie P&L Grupy Azoty"
TOC["A2"] = "GA (produkty, logistyka, dystrybutorzy, marka, nawozy.eu, ARiMR, 5 funkcjonalności - bez gotówki) + Partner Technologiczny (silnik + UX) + Partner Marketplace (plug-in). 7 lat. Ambicja → potrzebne finansowanie (CAC 250 zł)."
TOC.cell(4, 1, "Arkusz"); TOC.cell(4, 2, "Co zawiera")
TOC_ITEMS = [("Każdy partner zyskuje", "Wynik partnerów, skala, budżet pozyskania, potrzebne finansowanie"), ("Finansowanie", "Ile trzeba sfinansować: wrażliwość na CAC, ambicję i synergie"),
             ("Model współpracy", "Rekomendowane zasady długoterminowej współpracy (term sheet)"), ("Wkład partnerów", "Kto co wnosi, robi i dostaje"),
             ("Uwagi 04.11", "6 uwag zespołu GA i co zmieniliśmy w v4"), ("Uwagi do modelu GA", "Co przejęliśmy z GAia-model_PL_v2 i co zmieniliśmy"), ("Założenia", "Wszystkie założenia: ambicja, CAC, przychody, synergie (do dyskusji)"),
             ("Klucz podziału", "Podział strumieni przychodów"), ("Dopasowanie MVP1", "Wymagania OPZ MVP1: kto dostarcza, dopasowanie, osobomiesiące"),
             ("Rynek i GMV", "Użytkownicy, GMV, nawozy GA"), ("Koszty platformy", "Koszty jak w modelu GA i w partnerstwie (synergie)"),
             ("Wspólna pula", "Przychody, koszty, EBITDA, deficyt, zwrot, podział"), ("Grupa Azoty", "Dodatkowa marża, opłata za sukces, wynik"),
             ("Partner Technologiczny", "Wkład, koszty, wynik"), ("Partner Marketplace", "Wkład, koszty, wynik"), ("Tabele przestawne", "Tabele przestawne na arkuszu Dane"),
             ("Dane", "Dane źródłowe tabel przestawnych"), ("Case study i benchmarki", "Bushel, xFarm × Yara, Żabka × AiFi, Yara × Oxbury, PepsiCo × Yara, Bosch × BASF, venture studia"),
             ("Dofinansowania", "FENG SMART, Horizon Europe, Digital Europe, EIT Food, Innovation Fund, NFOŚiGW, ekoschemat, CBAM"), ("Dekarbonizacja", "Jak GAia zmniejsza emisje i jak to mierzyć")]
for i, (nm, desc) in enumerate(TOC_ITEMS):
    c = TOC.cell(5 + i, 1, nm); c.hyperlink = f"#'{nm}'!A1"; TOC.cell(5 + i, 2, desc)
lg = 5 + len(TOC_ITEMS) + 1; TOC.cell(lg, 1, "Legenda")
legend = [("1 234,5", Font(name=FONT, color="0000FF", size=10), PatternFill("solid", fgColor="FFF2CC"), "Założenie - można zmieniać"),
          ("1 234,5", Font(name=FONT, color="000000", size=10), None, "Formuła w tym samym arkuszu"), ("1 234,5", Font(name=FONT, color=LINK, size=10), None, "Formuła z innego arkusza"),
          ("1 234,5", Font(name=FONT, bold=True, size=10), PatternFill("solid", fgColor="DDEBF7"), "Suma / wynik"), ("Komentarz", Font(name=FONT, size=10), None, "Czerwony trójkąt = opis formuły lub źródła")]
for i, (t, fnt, fil, d) in enumerate(legend): TOC.cell(lg + 1 + i, 1, t); TOC.cell(lg + 1 + i, 2, d)
TOC.column_dimensions["A"].width = 30; TOC.column_dimensions["B"].width = 110
# formatowanie
for sh, tabc in CALC:
    last = max(sh.R.values()); finish(sh.ws, [(4, last, 1, NOTEC)], tab=tabc)
    for r in range(5, last + 1): sh.ws.cell(r, NOTEC).alignment = Alignment(wrap_text=True, vertical="center"); sh.ws.cell(r, NOTEC).font = Font(name=FONT, size=9, italic=True, color=GREYT)
finish(A, [(4, A.max_row, 1, 3 + NY)], tab="FFC000")
for r in range(5, A.max_row + 1): A.cell(r, 3 + NY).alignment = Alignment(wrap_text=True, vertical="center")
finish(P, REG_P, tab="008A3E"); P.cell(3, 1).font = Font(name=FONT, bold=True, color=NAVY_, size=12)
for r, t in ((SK, "Skala"), (IV, "Wkład i finansowanie")): pass
finish(S3, [(FIN_T1 + 1, FIN_T1 + 6, 1, 9), (FIN_T2 + 1, FIN_T2 + 5, 1, 9), (FIN_T3 + 1, FIN_T3 + 4, 1, 9)], tab="008A3E")
for r_ in (FIN_T1, FIN_T2, FIN_T3): S3.cell(r_, 1).font = Font(name=FONT, bold=True, color=NAVY_, size=11)
for c_ in range(1, 10): S3.cell(FIN_T1 + 1, c_).alignment = S3.cell(FIN_T2 + 1, c_).alignment = S3.cell(FIN_T3 + 1, c_).alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
finish(K, [(4, 4 + len(KEYS), 1, 6)], tab="E67E00")
for r in range(5, 5 + len(KEYS)): K.cell(r, 6).alignment = Alignment(wrap_text=True, vertical="center"); K.row_dimensions[r].height = 30
finish(C, [(4, mvp_tot, 1, 7)], tab="E67E00")
for ws_, last_ in ((WSP, WSP_last), (WKP, WKP_last), (UW, UW_last), (U4, U4_last), (BM, BM_last), (FD, FD_last), (DK, DK_last)):
    finish(ws_, [(4, last_, 1, ws_.max_column)], tab="6F9FD8")
    for r in range(5, last_ + 1):
        for c in range(1, ws_.max_column + 1): ws_.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
for ws_, last_ in ((BM, BM_last), (FD, FD_last)):
    for r in range(5, last_ + 1): ws_.row_dimensions[r].height = 60; ws_.cell(r, 6).font = Font(name=FONT, size=9, color="0563C1", underline="single")
for r in range(5, WKP_last + 1): WKP.row_dimensions[r].height = 62
for r_ in (FIN_T1, FIN_T2, FIN_T3): S3.row_dimensions[r_ + 1].height = 45
for r in range(5, U4_last + 1): U4.row_dimensions[r].height = 75
finish(DS, [], tab="BFBFBF")
for c in range(1, 7): DS.cell(4, c).font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
finish(TOC, [(4, 4 + len(TOC_ITEMS), 1, 2)], tab="003665", land=False)
TOC.cell(lg, 1).font = Font(name=FONT, bold=True, color=NAVY_, size=11)
for i in range(len(TOC_ITEMS)): TOC.cell(5 + i, 1).font = Font(name=FONT, color="0563C1", underline="single", size=10)
for i, (t, fnt, fil, d) in enumerate(legend):
    c = TOC.cell(lg + 1 + i, 1); c.font = fnt; c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    if fil: c.fill = fil
ORDER = ["Spis treści", "Każdy partner zyskuje", "Finansowanie", "Uwagi 04.11", "Model współpracy", "Wkład partnerów", "Uwagi do modelu GA", "Założenia", "Klucz podziału",
         "Dopasowanie MVP1", "Rynek i GMV", "Koszty platformy", "Wspólna pula", "Grupa Azoty", "Partner Technologiczny", "Partner Marketplace", "Dane", "Case study i benchmarki", "Dofinansowania", "Dekarbonizacja"]
wb._sheets = [wb[n] for n in ORDER]; wb.active = 0
