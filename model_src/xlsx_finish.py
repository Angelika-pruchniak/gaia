# Wykonywany na końcu build_bc_xlsx.py (exec): opisy formuł, komentarze, arkusz danych do tabel
# przestawnych, spis treści i profesjonalne formatowanie (bez siatki, linie tylko w tabelach).
from openpyxl.comments import Comment
from openpyxl.styles import Border, Side, Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo

FONT = "Trebuchet MS"
NAVY_ = "003665"; GRID = "BFBFBF"; LINK = "00704A"; GREYT = "595959"
thin = Side(style="thin", color=GRID); navy_thin = Side(style="thin", color=NAVY_); navy_med = Side(style="medium", color=NAVY_)
NUMF = '#,##0.0;[Red]-#,##0.0;"–"'

# ---------------------------------------------------------------- 1. Opisy: jak liczone / skąd
NOTES = {
 U: {"u": "Założenia, sekcja 1: cel użytkowników × mnożnik stres testu",
     "nu": "Użytkownicy w roku − użytkownicy rok wcześniej",
     "buy": "Użytkownicy × odsetek aktywnych kupujących (Założenia: 35%)",
     "pr": "Użytkownicy × odsetek abonentów premium (Założenia: 3% → 8%)",
     "bf": "Założenia: średnie zakupy nawozów GA przez platformę na kupującego",
     "b3": "Założenia: średnie zakupy nasion, ŚOR, części, usług na kupującego",
     "fg": "Aktywni kupujący (tys.) × zakupy nawozów (tys. zł) = mln zł",
     "g3": "Aktywni kupujący (tys.) × zakupy innych produktów (tys. zł) = mln zł",
     "gmv": "GMV nawozów + GMV innych produktów",
     "pen": "GMV nawozów ÷ sprzedaż krajowa nawozów GA (2,93 mln t × 70% × 1 800 zł/t)",
     "of": "GMV nawozów ÷ średnie zamówienie nawozów (4,5 tys. zł)",
     "o3": "GMV innych produktów ÷ średnie zamówienie (1,2 tys. zł)",
     "ot": "Zamówienia nawozów + zamówienia innych produktów",
     "tn": "GMV nawozów (mln zł) × 1000 ÷ cena 1 800 zł/t",
     "oad": "Zamówienia nawozów × 60% w szczycie ÷ 66 punktów wydania ÷ 90 dni",
     "sel": "Założenia, sekcja 8"},
 T: {"cap": "Arkusz Wkład technologiczny: 'Razem Partner Technologiczny' (rok 0, 1). Praca z własnych zasobów Partnera",
     "dev": "FTE zespołu rozwoju × 12 miesięcy × koszt osobomiesiąca (20 tys. zł)", "sup": "FTE utrzymania × 0,15 mln zł",
     "cloud": "0,8 mln zł stałe + użytkownicy × 6 zł", "sat": "Użytkownicy × 4 zł (zdjęcia satelitarne, pogoda)", "ai": "Użytkownicy × 3 zł (modele AI)",
     "sec": "Stały koszt: testy, audyty, RODO, NIS2", "opex": "Suma kosztów utrzymania i rozwoju", "cost": "Wkład (rok 0-1) + OPEX",
     "cpu": "OPEX Partnera ÷ użytkownicy", "fopx": "OPEX Partnera × (1 + 15%) - płaci spółka GAia",
     "own": "Minus wkład pracy w roku 0-1 (Partner pokrywa go sam)", "mopx": "Za utrzymanie − koszt utrzymania = 15% marży",
     "lic": "Arkusz Spółka GAia: licencja za silnik (użytkownicy × 10 zł)",
     "fin": "Minus dopłaty na start spółki + ich zwrot z zysku (udział 40/66 w dopłatach)",
     "jv": "Dywidenda spółki GAia × 40%", "res": "Wkład + marża na usługach + licencja + dopłaty/zwrot + dywidenda"},
 J: {"rc": "GMV innych produktów × prowizja 7%", "ra": "GMV innych produktów × reklama 1,5%",
     "rg": "GMV nawozów × 1,2% (prowizja od sprzedaży nawozów GA przez platformę)", "rf": "GMV razem × 0,4% (prowizja od banku za finansowanie zakupów)",
     "rs": "Abonenci premium (tys.) × 0,24 tys. zł (20 zł/mies.; benchmark SatAgro, FieldView)", "rd": "Założenia: zanonimizowane analizy dla producentów nasion i ŚOR",
     "rev": "Suma przychodów", "ctp": "Arkusz Partner Technologiczny: 'Za utrzymanie i rozwój' (koszt + 15%)",
     "cmp": "Arkusz Partner Marketplace: 'Za operacje' (koszt + 15%)",
     "ltp": "Użytkownicy × 10 zł: licencja za silnik Partnera Technologicznego", "lmp": "GMV innych produktów × 1%: licencja za marketplace Partnera Marketplace",
     "cpay": "GMV razem × 0,6% (operator płatności)", "cmkt": "Nowi użytkownicy × 60 zł + marka 3 mln zł; rok 0: kampania startowa 2 mln zł",
     "ccs": "Użytkownicy ÷ 20 tys. × 0,12 mln zł", "cmg": "6 osób × 0,3 mln zł; rok 0: pół roku", "opex": "Suma kosztów operacyjnych",
     "ebitda": "Przychody − OPEX", "emg": "EBITDA ÷ przychody", "fcf": "Równe EBITDA. Spółka nie ma CAPEX: budowę robią partnerzy jako wkład",
     "fund": "Gdy przepływy ujemne: dopłaca Partner Technologiczny i Partner Marketplace (40:26). GA nie dopłaca",
     "bal": "Saldo dopłat: poprzedni rok + nowe dopłaty − zwrot", "rep": "Zysk spółki najpierw spłaca dopłaty partnerów",
     "div": "Zysk po zwrocie dopłat, dzielony 34 / 40 / 26"},
 Mp: {"cap": "Arkusz Wkład technologiczny: 'Razem Partner Marketplace' (rok 0, 1). Praca z własnych zasobów Partnera",
      "team": "FTE operacji × 0,18 mln zł (katalog, sprzedawcy, spory)", "inf": "0,4 mln zł stałe + użytkownicy × 2 zł",
      "opex": "Zespół + infrastruktura", "cost": "Wkład (rok 0-1) + OPEX", "fopx": "OPEX × (1 + 15%) - płaci spółka GAia",
      "own": "Minus wkład pracy w roku 0-1 (Partner pokrywa go sam)", "mopx": "Za operacje − koszt = 15% marży",
      "lic": "Arkusz Spółka GAia: licencja za marketplace (GMV 3P × 1%)",
      "fin": "Minus dopłaty na start spółki + ich zwrot z zysku (udział 26/66 w dopłatach)",
      "jv": "Dywidenda spółki GAia × 26%", "res": "Wkład + marża na operacjach + licencja + dopłaty/zwrot + dywidenda"},
 G: {"sh": "Dodatkowe p.p. udziału AN/CAN × 39 mln zł przychodu na 1 p.p. × 25% marży kontrybucyjnej",
     "ch": "GMV nawozów × 2% (krótsza ścieżka sprzedaży, mniej upustów)", "mx": "GMV nawozów × 1% (więcej NPK i nawozów specjalistycznych)",
     "fee": "Minus prowizja 1,2% od nawozów sprzedanych przez platformę", "cash": "GA nie wnosi gotówki ani nie dopłaca na start",
     "jv": "Dywidenda spółki GAia × 34%", "res": "Suma składników wyniku GA"},
 Mc: {"tr": "Przychody spółki GAia ÷ GMV razem", "ru": "Przychody ÷ użytkownicy", "cu": "OPEX spółki ÷ użytkownicy", "eu": "EBITDA ÷ użytkownicy"},
}
for sh in (U, T, J, Mp, G, Mc):
    allnotes = dict(sh.notes); allnotes.update(NOTES.get(sh, {}))
    for key, text in allnotes.items():
        if key not in sh.R: continue
        r = sh.R[key]
        sh.ws.cell(r, 9, text)
        sh.ws.cell(r, 1).comment = Comment(text, "Model GAia", width=320, height=90)
    for key, r in sh.R.items():          # wiersze z Mocy platformy itp. bez opisu: link do źródła
        if sh.ws.cell(r, 9).value is None:
            f = str(sh.ws.cell(r, 3).value or "")
            if "!" in f: sh.ws.cell(r, 9, "Link do arkusza " + f.split("!")[0].lstrip("=").strip("'"))

# komentarze przy założeniach
for row in A.iter_rows(min_row=5, max_row=A.max_row):
    src, lab = row[7].value, row[0].value
    if lab and src and row[1].value is not None:
        row[1].comment = Comment(f"{lab}\nŹródło / uwaga: {src}", "Model GAia", width=320, height=110)
# komentarze w podsumowaniu
P.cell(4, 9).comment = Comment("NPV = rok 0 + NPV(10%; lata 1-5). Stopa w Założeniach, sekcja 9.", "Model GAia")
P.cell(4, 11).comment = Comment("Pierwszy rok, w którym suma narastająca od roku 0 jest dodatnia.", "Model GAia")
P.cell(4, 1).comment = Comment("Wyniki z arkuszy: Grupa Azoty, Partner Technologiczny, Partner Marketplace, Spółka GAia (przepływy).", "Model GAia")

# ---------------------------------------------------------------- 2. Dane do tabel przestawnych
DS = wb.create_sheet("Dane")
DS["A1"] = "Dane źródłowe do tabel przestawnych (formuły z arkuszy obliczeniowych)"
DS["A2"] = "Każdy wiersz = jedna pozycja w jednym roku. Wartości to formuły, więc tabele przestawne aktualizują się po odświeżeniu (Dane → Odśwież wszystko)."
heads = ["Obszar", "Typ", "Pozycja", "Rok", "Wartość (mln zł)", "Arkusz źródłowy"]
for j, h in enumerate(heads): DS.cell(4, 1 + j, h)
spec = [
 ("Grupa Azoty", "Wynik partnera", "Grupa Azoty (34%)", G, "res"), ("Partner Technologiczny", "Wynik partnera", "Partner Technologiczny (40%)", T, "res"),
 ("Partner Marketplace", "Wynik partnera", "Partner Marketplace (26%)", Mp, "res"), ("Spółka GAia", "Wynik partnera", "Spółka GAia (gotówka)", J, "fcf"),
 ("Spółka GAia", "Przychód spółki GAia", "1 Prowizje 7%", J, "rc"), ("Spółka GAia", "Przychód spółki GAia", "2 Reklama 1,5%", J, "ra"),
 ("Spółka GAia", "Przychód spółki GAia", "3 Prowizja od nawozów GA 1,2%", J, "rg"), ("Spółka GAia", "Przychód spółki GAia", "4 Finansowanie 0,4%", J, "rf"),
 ("Spółka GAia", "Przychód spółki GAia", "5 Abonament premium", J, "rs"), ("Spółka GAia", "Przychód spółki GAia", "6 Analizy dla producentów", J, "rd"),
 ("Spółka GAia", "OPEX spółki GAia", "1 Partner Technologiczny (utrzymanie)", J, "ctp"), ("Spółka GAia", "OPEX spółki GAia", "2 Partner Marketplace (operacje)", J, "cmp"),
 ("Spółka GAia", "OPEX spółki GAia", "3 Licencja: silnik", J, "ltp"), ("Spółka GAia", "OPEX spółki GAia", "4 Licencja: marketplace", J, "lmp"),
 ("Spółka GAia", "OPEX spółki GAia", "5 Płatności", J, "cpay"), ("Spółka GAia", "OPEX spółki GAia", "6 Marketing", J, "cmkt"),
 ("Spółka GAia", "OPEX spółki GAia", "7 Obsługa klienta", J, "ccs"), ("Spółka GAia", "OPEX spółki GAia", "8 Zarząd i zespół", J, "cmg"),
 ("Spółka GAia", "EBITDA i podział gotówki", "1 EBITDA (przepływy)", J, "fcf"), ("Spółka GAia", "EBITDA i podział gotówki", "2 Dopłaty partnerów na start", J, "fund"),
 ("Spółka GAia", "EBITDA i podział gotówki", "3 Zwrot dopłat", J, "rep"), ("Spółka GAia", "EBITDA i podział gotówki", "4 Dywidenda do podziału", J, "div"),
 ("Partner Technologiczny", "Koszt Partnera Technologicznego", "1 Wkład: dopracowanie GAia", T, "cap"), ("Partner Technologiczny", "Koszt Partnera Technologicznego", "2 Zespół rozwoju", T, "dev"),
 ("Partner Technologiczny", "Koszt Partnera Technologicznego", "3 Utrzymanie 24/7", T, "sup"), ("Partner Technologiczny", "Koszt Partnera Technologicznego", "4 Chmura", T, "cloud"),
 ("Partner Technologiczny", "Koszt Partnera Technologicznego", "5 Dane satelitarne", T, "sat"), ("Partner Technologiczny", "Koszt Partnera Technologicznego", "6 Asystent AI", T, "ai"),
 ("Partner Technologiczny", "Koszt Partnera Technologicznego", "7 Bezpieczeństwo", T, "sec"),
 ("Partner Technologiczny", "Zarobek Partnera Technologicznego", "0 Wkład własny (minus)", T, "own"), ("Partner Technologiczny", "Zarobek Partnera Technologicznego", "1 Marża na usługach", T, "mopx"),
 ("Partner Technologiczny", "Zarobek Partnera Technologicznego", "2 Licencja za silnik", T, "lic"), ("Partner Technologiczny", "Zarobek Partnera Technologicznego", "3 Dopłaty i ich zwrot", T, "fin"),
 ("Partner Technologiczny", "Zarobek Partnera Technologicznego", "4 Dywidenda 40%", T, "jv"),
 ("Partner Marketplace", "Koszt Partnera Marketplace", "1 Wkład: plug-in i dostosowanie", Mp, "cap"), ("Partner Marketplace", "Koszt Partnera Marketplace", "2 Zespół operacji", Mp, "team"),
 ("Partner Marketplace", "Koszt Partnera Marketplace", "3 Infrastruktura", Mp, "inf"),
 ("Partner Marketplace", "Zarobek Partnera Marketplace", "0 Wkład własny (minus)", Mp, "own"), ("Partner Marketplace", "Zarobek Partnera Marketplace", "1 Marża na operacjach", Mp, "mopx"),
 ("Partner Marketplace", "Zarobek Partnera Marketplace", "2 Licencja za marketplace", Mp, "lic"), ("Partner Marketplace", "Zarobek Partnera Marketplace", "3 Dopłaty i ich zwrot", Mp, "fin"),
 ("Partner Marketplace", "Zarobek Partnera Marketplace", "4 Dywidenda 26%", Mp, "jv"),
 ("Grupa Azoty", "Wynik Grupy Azoty", "1 Udział w rynku AN/CAN", G, "sh"), ("Grupa Azoty", "Wynik Grupy Azoty", "2 Niższy koszt kanału", G, "ch"),
 ("Grupa Azoty", "Wynik Grupy Azoty", "3 Lepszy mix", G, "mx"), ("Grupa Azoty", "Wynik Grupy Azoty", "4 Prowizja dla spółki", G, "fee"),
 ("Grupa Azoty", "Wynik Grupy Azoty", "5 Dywidenda 34%", G, "jv"),
 ("Platforma", "GMV", "1 Nawozy GA", U, "fg"), ("Platforma", "GMV", "2 Inne produkty (3P)", U, "g3"),
]
SIGN = set()
dr = 5
for area, typ, pos, sh, key in spec:
    for yi, col in enumerate(["B"] + YC):
        sgn = "-" if (area, key) in SIGN else ""
        DS.cell(dr, 1, area); DS.cell(dr, 2, typ); DS.cell(dr, 3, pos); DS.cell(dr, 4, f"Rok {yi}")
        DS.cell(dr, 5, f"={sgn}N('{sh.ws.title}'!{col}{sh.R[key]})").number_format = NUMF
        DS.cell(dr, 6, sh.ws.title); dr += 1
DATA_LAST = dr - 1
tab = Table(displayName="DaneGAia", ref=f"A4:F{DATA_LAST}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
DS.add_table(tab)
for col, w in zip("ABCDEF", [24, 34, 40, 9, 16, 26]): DS.column_dimensions[col].width = w
DS.freeze_panes = "A5"

# ---------------------------------------------------------------- 3. Spis treści
TOC = wb.create_sheet("Spis treści", 0)
TOC["A1"] = "GAia: business case platformy dla rolników (300 tys. użytkowników w roku 5)"
TOC["A2"] = "Joint venture: Grupa Azoty (produkty, dane, ARiMR - bez gotówki), Partner Technologiczny (silnik + UX), Partner Marketplace / Handel (marketplace przez plug-in). Kwoty w mln zł."
TOC.cell(4, 1, "Arkusz"); TOC.cell(4, 2, "Co zawiera")
toc = [("Każdy partner zyskuje", "Wynik każdego partnera rok po roku, NPV, zwrot, skala, inwestycja i finansowanie"),
       ("Wkład partnerów", "Kto co wnosi, udziały 34 / 40 / 26, jak zarabia, ład korporacyjny"),
       ("Założenia", "Wszystkie założenia (niebieskie na żółtym tle) ze źródłami i benchmarkami"),
       ("Wkład technologiczny", "Dopracowanie GAia w roku 0 z własnych zasobów partnerów: moduł, kto robi, osobomiesiące, koszt"),
       ("Użytkownicy i GMV", "Użytkownicy → kupujący → GMV → zamówienia i tony"),
       ("Partner Technologiczny", "Wkład, koszty i źródła zarobku Partnera Technologicznego"),
       ("Partner Marketplace", "Wkład, koszty i źródła zarobku Partnera Marketplace"),
       ("Spółka GAia (JV)", "Przychody, OPEX, EBITDA, dopłaty partnerów na start, zwrot dopłat, dywidenda"),
       ("Grupa Azoty", "Składniki wyniku GA"), ("Moc platformy", "Wskaźniki: zamówienia, tony, przychód i koszt na użytkownika"),
       ("Tabele przestawne", "Tabele przestawne na danych z arkusza Dane"),
       ("Dane", "Płaska tabela danych (formuły) - źródło tabel przestawnych"),
       ("Benchmarki rynkowe", "Allegro, Yara, BASF, Bayer, SatAgro, FBN, Agrofy, DeHaat, Agrando, Bosch × BASF, praktyki JV")]
for i, (nm, desc) in enumerate(toc):
    c = TOC.cell(5 + i, 1, nm); c.hyperlink = f"#'{nm}'!A1"; c.font = Font(name=FONT, color="0563C1", underline="single", size=10)
    TOC.cell(5 + i, 2, desc)
lg = 5 + len(toc) + 1
TOC.cell(lg, 1, "Legenda")
legend = [("1 234,5", Font(name=FONT, color="0000FF", size=10), PatternFill("solid", fgColor="FFF2CC"), "Założenie - można zmieniać"),
          ("1 234,5", Font(name=FONT, color="000000", size=10), None, "Formuła w tym samym arkuszu"),
          ("1 234,5", Font(name=FONT, color=LINK, size=10), None, "Formuła z innego arkusza (link)"),
          ("1 234,5", Font(name=FONT, bold=True, size=10), PatternFill("solid", fgColor="DDEBF7"), "Suma / wynik"),
          ("Komentarz", Font(name=FONT, size=10), None, "Czerwony trójkąt w rogu komórki = opis formuły lub źródła (najedź myszką)")]
for i, (t, fnt, fil, d) in enumerate(legend):
    c = TOC.cell(lg + 1 + i, 1, t); c.font = fnt
    if fil: c.fill = fil
    TOC.cell(lg + 1 + i, 2, d)
TOC.column_dimensions["A"].width = 28; TOC.column_dimensions["B"].width = 100

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

# arkusze obliczeniowe
for sh, tabc in ((U, "2E6DB4"), (T, "2E6DB4"), (J, "003665"), (Mp, "2E6DB4"), (G, "003665"), (Mc, "6F9FD8")):
    last = max(sh.R.values())
    finish(sh.ws, [(4, last, 1, 9)], tab=tabc)
    for r in range(last + 1, sh.ws.max_row + 1):
        for c in (1, 2):
            if sh.ws.cell(r, c).value is not None: sh.ws.cell(r, c).font = Font(name=FONT, size=10, bold=c == 1 and "Wkład" not in str(sh.ws.cell(r, 1).value), color=NAVY_)
    for r in range(5, last + 1): sh.ws.cell(r, 9).alignment = Alignment(wrap_text=True, vertical="center"); sh.ws.cell(r, 9).font = Font(name=FONT, size=9, italic=True, color=GREYT)
# funding block in JV
jl = max(J.R.values())
border_table(J.ws, jl + 2, jl + 5, 1, 2, header=False)
# Założenia
finish(A, [(4, A.max_row, 1, 8)], tab="FFC000")
for r in range(5, A.max_row + 1): A.cell(r, 8).alignment = Alignment(wrap_text=True, vertical="center")
# CAPEX
finish(C, [(4, capj_r, 1, 10)], tab="E67E00")
for r in range(5, capj_r): C.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="center")
C.cell(capj_r + 2, 1).font = Font(name=FONT, bold=True, color=NAVY_, size=11)
# Każdy partner zyskuje
finish(P, REG_P, tab="008A3E")
finish(Wk, REG_W, tab="6F9FD8")
Wk.cell(14, 1).font = Font(name=FONT, bold=True, color=NAVY_, size=11)
for r in range(4, 7): Wk.row_dimensions[r].height = 64
finish(Bm, [(3, Bm.max_row, 1, 6)], tab="6F9FD8")
for r in range(4, Bm.max_row + 1):
    Bm.row_dimensions[r].height = 62
    c = Bm.cell(r, 6); c.font = Font(name=FONT, size=9, color="0563C1", underline="single")
finish(DS, [], tab="BFBFBF")
for c in range(1, 7): DS.cell(4, c).font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
finish(TOC, [(4, 4 + len(toc), 1, 2)], tab="003665", land=False)
TOC.cell(lg, 1).font = Font(name=FONT, bold=True, color=NAVY_, size=11)
for i in range(len(toc)): TOC.cell(5 + i, 1).font = Font(name=FONT, color="0563C1", underline="single", size=10)
for i, (t, fnt, fil, d) in enumerate(legend): TOC.cell(lg + 1 + i, 1).font = fnt; TOC.cell(lg + 1 + i, 1).border = Border(left=thin, right=thin, top=thin, bottom=thin)
# kolejność arkuszy
order = ["Spis treści", "Każdy partner zyskuje", "Wkład partnerów", "Założenia", "Wkład technologiczny", "Użytkownicy i GMV", "Partner Technologiczny",
         "Partner Marketplace", "Spółka GAia (JV)", "Grupa Azoty", "Moc platformy", "Dane", "Benchmarki rynkowe"]
wb._sheets = [wb[n] for n in order]
wb.active = 0
