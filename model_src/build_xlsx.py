# Generator modelu ekonomiki GAia (JV z partnerem technologicznym). Uruchom: python3 build_xlsx.py
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L

OUT = "../GA_GAia_model_JV_partnerzy.xlsx"
wb = openpyxl.Workbook()
NAVY = "1F3864"
BLUE = Font(color="0000FF")
B = Font(bold=True)
H = Font(bold=True, color="FFFFFF")
HF = PatternFill("solid", fgColor=NAVY)
INF = PatternFill("solid", fgColor="FFF2CC")
TOT = PatternFill("solid", fgColor="DDEBF7")
thin = Side(style="thin", color="BFBFBF")
YRS = ["Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5"]
YC = ["C", "D", "E", "F", "G"]  # kolumny lat na arkuszach z latami
SC = [("S1", "Iteo + własna operacja GA (15-30 FTE)"),
      ("S1a", "Iteo + Mirakl, fee 6% od CAŁEGO GMV (wersja ze slajdu)"),
      ("S1b", "Iteo + Mirakl, fee 6% tylko od GMV sprzedawców zewn."),
      ("S2", "Iteo (JV) + OLX (umowa revenue share) - REKOMENDACJA")]

def hdr(ws, r, vals, c0=1):
    for i, v in enumerate(vals):
        c = ws.cell(r, c0 + i, v); c.font = H; c.fill = HF
        c.alignment = Alignment(wrap_text=True, vertical="center")

# ---------------- Założenia ----------------
A = wb.active; A.title = "Założenia"
N = {}  # nazwa -> adres (absolutny, z arkuszem)
A["A1"] = "GAia: założenia modelu (ILUSTRACJA - do potwierdzenia danymi Handlu i ofertami partnerów)"; A["A1"].font = Font(bold=True, size=14)
A["A2"] = "Niebieskie liczby na żółtym tle = założenia (można zmieniać). Czarne = formuły. Kwoty w mln zł, chyba że podano inaczej."
r = 4
hdr(A, r, ["Założenie", "Rok 1 / wartość", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Jedn.", "Źródło / uwaga"])
r += 1

def inp(name, label, vals, unit, src, fmt="0.0"):
    global r
    A.cell(r, 1, label)
    vals = vals if isinstance(vals, list) else [vals]
    for i, v in enumerate(vals):
        c = A.cell(r, 2 + i, v)
        if isinstance(v, str) and v.startswith("="):
            pass
        else:
            c.font = BLUE; c.fill = INF
        c.number_format = fmt
    A.cell(r, 7, unit); A.cell(r, 8, src)
    if len(vals) == 1:
        N[name] = f"'Założenia'!$B${r}"
    else:
        N[name] = {y: f"'Założenia'!${L(2+i)}${r}" for i, y in enumerate(range(5))}
    r += 1

def sec(t):
    global r
    r += 1; A.cell(r, 1, t).font = Font(bold=True, color=NAVY, size=12); r += 1

sec("1. Rynek i wolumen")
inp("vol", "Wolumen sprzedaży nawozów GA, plan 2026", 2.93, "mln t", "Strategia Nawozowa, sl. 90 (jak w GA_5_luk)", "0.00")
inp("dom", "Udział rynku krajowego", 0.70, "%", "Strategia Nawozowa, sl. 91", "0%")
inp("price", "Średnia cena netto nawozu", 1800, "zł/t", "ZAŁOŻENIE - podmienić na cenę netto po upustach (CO-PA)", "#,##0")
inp("base", "Sprzedaż krajowa nawozów GA", "=B{0}*B{1}*B{2}".format(r-3, r-2, r-1), "mln zł/rok", "Formuła", "#,##0")
inp("pen", "Udział sprzedaży krajowej GA przez platformę", [0.02, 0.05, 0.09, 0.13, 0.17], "%", "ZAŁOŻENIE - ścieżka adopcji rolników", "0%")
inp("mult", "Stres test: mnożnik adopcji (1 = plan, 0,5 = połowa)", 1.0, "x", "Zmień na 0,5, żeby zobaczyć wariant pesymistyczny", "0.00")
inp("r3p", "GMV sprzedawców zewnętrznych / GMV nawozów GA", [0.4, 0.6, 0.8, 1.0, 1.2], "x", "ZAŁOŻENIE - nasiona, ŚOR, części, usługi", "0.0")

sec("2. Przychody platformy (JV)")
inp("take", "Prowizja od sprzedawców zewnętrznych (take rate)", 0.08, "% GMV 3P", "ZAŁOŻENIE - typowo 5-12% na marketplace B2B/agro", "0.0%")
inp("ads", "Reklama / promowane oferty", 0.015, "% GMV 3P", "ZAŁOŻENIE", "0.0%")
inp("fin", "Finansowanie, kredyt kupiecki, ubezpieczenia (prowizja)", 0.004, "% GMV", "ZAŁOŻENIE - np. z bankiem / BGK", "0.0%")
inp("subu", "Rolnicy z abonamentem premium (plan nawożenia, mapy, doradztwo)", [2, 6, 12, 20, 28], "tys.", "ZAŁOŻENIE - mitygacja ze slajdu", "0")
inp("subp", "Cena abonamentu premium", 1.2, "tys. zł/rok", "ZAŁOŻENIE ~100 zł/mies.", "0.0")
inp("payc", "Koszt płatności i obsługi transakcji", 0.006, "% GMV", "ZAŁOŻENIE", "0.0%")

sec("3. Korzyść GA poza JV (zostaje w 100% w GA)")
inp("chan", "Marża dystrybutora, której GA nie oddaje przy sprzedaży przez platformę", 0.07, "% GMV nawozów", "ZAŁOŻENIE - zweryfikować z umowami AD (68 AD)", "0.0%")
inp("logi", "Dodatkowy koszt logistyki last-mile / fee dla AD jako punktu odbioru", 0.035, "% GMV nawozów", "ZAŁOŻENIE - AD zostają partnerem logistycznym, a nie konkurentem", "0.0%")
inp("mix", "Poprawa mixu (więcej NPK / specjalistycznych dzięki planowi nawożenia)", 0.01, "% GMV nawozów", "ZAŁOŻENIE - mitygacja ze slajdu", "0.0%")
inp("pp", "Dodatkowy udział w rynku AN/CAN dzięki platformie", [0.2, 0.5, 1.0, 1.5, 2.0], "p.p.", "ZAŁOŻENIE - cel strategii 45% → 60%", "0.0")
inp("ppv", "Przychód z 1 p.p. udziału AN/CAN", 39, "mln zł/rok", "GA_5_luk, arkusz Szacunki (590 mln / 15 p.p.)", "0")
inp("cm", "Marża kontrybucyjna na dodatkowej tonie", 0.25, "%", "ZAŁOŻENIE - do potwierdzenia przez Kontroling", "0%")

sec("4. Struktura i finanse")
inp("gash", "Udział GA w wyniku JV", 0.51, "%", "Propozycja: GA 51% / Iteo 49% z opcją wykupu", "0%")
inp("fte", "Koszt 1 FTE (pełny)", 0.25, "mln zł/rok", "ZAŁOŻENIE", "0.00")
inp("mkt", "Marketing i pozyskanie rolników", [7, 7, 6, 5, 5], "mln zł/rok", "ZAŁOŻENIE", "0.0")
inp("wacc", "Stopa dyskontowa", 0.10, "%", "ZAŁOŻENIE - podmienić na WACC GA", "0%")

sec("5. Warunki partnerów wg scenariusza (to negocjujemy)")
r0 = r
hdr(A, r, ["Warunek", "S1", "S1a", "S1b", "S2", "Jedn.", "", "Uwaga"]); r += 1
def sc_inp(name, label, vals, unit, src, fmt="0.0"):
    global r
    A.cell(r, 1, label)
    for i, v in enumerate(vals):
        c = A.cell(r, 2 + i, v); c.font = BLUE; c.fill = INF; c.number_format = fmt
    A.cell(r, 6, unit); A.cell(r, 8, src)
    N[name] = [f"'Założenia'!${L(2+i)}${r}" for i in range(4)]
    r += 1
sc_inp("capex", "Jednorazowe wdrożenie / integracja (rok 0)", [15, 8, 8, 6], "mln zł", "S2: OLX i Iteo wnoszą platformy, GA płaci integrację z SAP/CRM")
sc_inp("itfix", "Iteo: stała opłata za platformę agro", [3, 3, 3, 3], "mln zł/rok", "Negocjować spadek po wykupie")
sc_inp("itsub", "Iteo: udział w przychodach z abonamentów", [0.25, 0.25, 0.25, 0.25], "% abon.", "Iteo zarabia, gdy rolnicy płacą za agro-funkcje", "0%")
sc_inp("ecshare", "Partner e-commerce: udział w prowizji i reklamie od sprzedawców 3P", [0, 0, 0, 0.35], "% przych. 3P", "OLX dostaje część tego, co sam pomaga sprzedać", "0%")
sc_inp("ecown", "Partner e-commerce: opłata od GMV nawozów GA", [0, 0, 0, 0.003], "% GMV naw.", "Czerwona linia: max 0,3%, bo to nasza marża", "0.0%")
sc_inp("plfee", "Platforma SaaS (Mirakl): fee od CAŁEGO GMV", [0, 0.06, 0, 0], "% GMV", "Tak liczono na slajdzie = źródło 'trwale nierentowny'", "0.0%")
sc_inp("plfee3p", "Platforma SaaS (Mirakl): fee tylko od GMV 3P", [0, 0, 0.06, 0], "% GMV 3P", "Wynegocjowane wyłączenie nawozów GA", "0.0%")
sc_inp("ownit", "Własne utrzymanie platformy e-commerce (open source): hosting, rozwój, bezpieczeństwo", [4, 0, 0, 0], "mln zł/rok", "")
sc_inp("ftes", "Zespół GA / JV", [25, 6, 8, 7], "FTE", "Slajd: 15-30 FTE w S1, kilka FTE z Mirakl / OLX", "0")
sc_inp("adopt", "Tempo adopcji vs plan (ruch, sprzedawcy, czas startu)", [0.7, 0.85, 0.85, 1.0], "x", "S2: OLX wnosi ruch i sprzedawców od 1. dnia; S1 buduje to sam (12-18 mies.)", "0.00")

A.column_dimensions["A"].width = 62
for c in "BCDEF": A.column_dimensions[c].width = 13
A.column_dimensions["G"].width = 14; A.column_dimensions["H"].width = 70
A.freeze_panes = "A5"

# ---------------- Model ----------------
M = wb.create_sheet("Model")
M["A1"] = "GAia: kto ile zarabia, rok po roku (mln zł)"; M["A1"].font = Font(bold=True, size=14)
M["A2"] = "Wszystko liczone z arkusza Założenia. Wynik GA = udział w JV + korzyść poza JV (marża kanałowa, mix, udział w rynku)."
hdr(M, 4, ["Pozycja", "Rok 0"] + YRS + ["Suma 5 lat"])
mr = 5
R = {}
def y(nm, i):
    v = N[nm]; return v[i] if isinstance(v, dict) else v
def row(key, label, f, bold=False, fill=None, y0=None, total=True):
    """f(i, col) -> formula string (bez '=')"""
    global mr
    M.cell(mr, 1, label)
    if y0 is not None:
        M.cell(mr, 2, "=" + y0)
    for i, col in enumerate(YC):
        M[f"{col}{mr}"] = "=" + f(i, col)
    if total:
        M[f"H{mr}"] = f"=SUM(B{mr}:G{mr})"
    for c in range(1, 9):
        cell = M.cell(mr, c); cell.number_format = '#,##0.0;[Red]-#,##0.0'
        if bold: cell.font = B
        if fill: cell.fill = fill
    R[key] = mr; mr += 1
def blank(t=None):
    global mr
    mr += 1
    if t: M.cell(mr, 1, t).font = Font(bold=True, color=NAVY, size=12); mr += 1

for s, (code, name) in enumerate(SC):
    k = code
    blank(f"{code}: {name}")
    ad = f"{N['mult']}*{N['adopt'][s]}"
    row(k+"fg", "GMV nawozów GA przez platformę", lambda i, c: f"{N['base']}*{y('pen',i)}*{ad}")
    row(k+"g3", "GMV sprzedawców zewnętrznych (3P)", lambda i, c: f"{c}{R[k+'fg']}*{y('r3p',i)}")
    row(k+"gmv", "GMV razem", lambda i, c: f"{c}{R[k+'fg']}+{c}{R[k+'g3']}", bold=True)
    row(k+"rc", "Przychód JV: prowizje od sprzedawców 3P", lambda i, c: f"{c}{R[k+'g3']}*{N['take']}")
    row(k+"ra", "Przychód JV: reklama", lambda i, c: f"{c}{R[k+'g3']}*{N['ads']}")
    row(k+"rf", "Przychód JV: finansowanie / ubezpieczenia", lambda i, c: f"{c}{R[k+'gmv']}*{N['fin']}")
    row(k+"rs", "Przychód JV: abonamenty premium", lambda i, c: f"{y('subu',i)}*{N['subp']}*{ad}")
    row(k+"rev", "Przychody JV razem", lambda i, c: f"SUM({c}{R[k+'rc']}:{c}{R[k+'rs']})", bold=True)
    row(k+"pc", "Koszt: płatności", lambda i, c: f"{c}{R[k+'gmv']}*{N['payc']}")
    row(k+"it", "Koszt: Iteo (opłata stała + udział w abonamentach)", lambda i, c: f"{N['itfix'][s]}+{N['itsub'][s]}*{c}{R[k+'rs']}")
    row(k+"ec", "Koszt: partner e-commerce (revenue share)", lambda i, c: f"{N['ecshare'][s]}*({c}{R[k+'rc']}+{c}{R[k+'ra']})+{N['ecown'][s]}*{c}{R[k+'fg']}")
    row(k+"pl", "Koszt: platforma SaaS (Mirakl)", lambda i, c: f"{N['plfee'][s]}*{c}{R[k+'gmv']}+{N['plfee3p'][s]}*{c}{R[k+'g3']}")
    row(k+"own", "Koszt: własne IT + zespół + marketing", lambda i, c: f"{N['ownit'][s]}+{N['ftes'][s]}*{N['fte']}+{y('mkt',i)}")
    row(k+"jv", "WYNIK OPERACYJNY JV (P&L platformy)", lambda i, c: f"{c}{R[k+'rev']}-{c}{R[k+'pc']}-{c}{R[k+'it']}-{c}{R[k+'ec']}-{c}{R[k+'pl']}-{c}{R[k+'own']}",
        bold=True, y0=f"-{N['capex'][s]}")
    row(k+"gajv", "GA: udział w wyniku JV", lambda i, c: f"{c}{R[k+'jv']}*{N['gash']}", y0=f"B{mr-1}*{N['gash']}")
    row(k+"ch", "GA poza JV: marża kanałowa netto (marża AD - logistyka)", lambda i, c: f"{c}{R[k+'fg']}*({N['chan']}-{N['logi']})")
    row(k+"mx", "GA poza JV: poprawa mixu", lambda i, c: f"{c}{R[k+'fg']}*{N['mix']}")
    row(k+"sh", "GA poza JV: marża z dodatkowego udziału AN/CAN", lambda i, c: f"{y('pp',i)}*{ad}*{N['ppv']}*{N['cm']}")
    row(k+"core", "GA poza JV razem (100% dla GA)", lambda i, c: f"SUM({c}{R[k+'ch']}:{c}{R[k+'sh']})", bold=True)
    row(k+"ga", "GA: WYNIK RAZEM (udział w JV + poza JV)", lambda i, c: f"{c}{R[k+'gajv']}+{c}{R[k+'core']}", bold=True, fill=TOT, y0=f"B{R[k+'gajv']}")
    row(k+"cum", "GA: wynik skumulowany", lambda i, c: "0", total=False)
    row(k+"partners", "Partnerzy zarabiają razem (Iteo + e-com + SaaS)", lambda i, c: f"{c}{R[k+'it']}+{c}{R[k+'ec']}+{c}{R[k+'pl']}")
for code, _ in SC:
    rr = R[code+"cum"]
    M[f"B{rr}"] = f"=B{rr-1}"
    for i, col in enumerate(YC):
        prev = "B" if i == 0 else YC[i-1]
        M[f"{col}{rr}"] = f"={prev}{rr}+{col}{rr-1}"
M.column_dimensions["A"].width = 52
for c in "BCDEFGH": M.column_dimensions[c].width = 12
M.freeze_panes = "B5"

# ---------------- Podsumowanie ----------------
P = wb.create_sheet("Podsumowanie", 0)
P["A1"] = "GAia: czy GA zarabia przy kosztach partnera technologicznego?"; P["A1"].font = Font(bold=True, size=14)
P["A2"] = "Odpowiedź: platforma to kanał sprzedaży nawozów GA. Liczymy wynik Grupy, nie tylko P&L platformy. Liczby ilustracyjne - do potwierdzenia."
hdr(P, 4, ["Scenariusz", "Wynik JV rok 5", "JV na plusie od", "Wynik GA rok 1", "Wynik GA rok 5", "GA suma 5 lat (z wdrożeniem)", "NPV dla GA", "GA zwrot od", "Partnerzy zarabiają (5 lat)"])
yrs_row = 3  # helper: numery lat w Model!C3:G3
for i, col in enumerate(YC):
    M[f"{col}3"] = i + 1; M[f"{col}3"].font = Font(color="FFFFFF")
for s, (code, name) in enumerate(SC):
    pr = 5 + s
    jv, ga, cum, prt = R[code+"jv"], R[code+"ga"], R[code+"cum"], R[code+"partners"]
    P.cell(pr, 1, f"{code}: {name}")
    P.cell(pr, 2, f"=Model!G{jv}")
    P.cell(pr, 3, f'=IF(COUNTIF(Model!C{jv}:G{jv},">0")=0,"po roku 5","rok "&SUMPRODUCT(MIN(IF(Model!C{jv}:G{jv}>0,Model!C3:G3,99))))')
    P.cell(pr, 4, f"=Model!C{ga}")
    P.cell(pr, 5, f"=Model!G{ga}")
    P.cell(pr, 6, f"=Model!H{ga}")
    P.cell(pr, 7, f"=Model!B{ga}+NPV({N['wacc']},Model!C{ga}:G{ga})")
    P.cell(pr, 8, f'=IF(COUNTIF(Model!C{cum}:G{cum},">0")=0,"po roku 5","rok "&SUMPRODUCT(MIN(IF(Model!C{cum}:G{cum}>0,Model!C3:G3,99))))')
    P.cell(pr, 9, f"=Model!H{prt}")
    for c in range(2, 10):
        P.cell(pr, c).number_format = '#,##0.0;[Red]-#,##0.0'
    if code == "S2":
        for c in range(1, 10): P.cell(pr, c).font = B; P.cell(pr, c).fill = TOT
P["A11"] = "Jak czytać"; P["A11"].font = B
notes = [
 "1. 'Trwale nierentowny' dotyczył wyłącznie P&L platformy w wariancie S1a (fee 6% od CAŁEGO GMV, także od nawozów GA). To błąd konstrukcji umowy, nie biznesu.",
 "2. Nawet w S1a GA jako Grupa zarabia, bo marża kanałowa, mix i udział w rynku zostają w 100% w GA (wiersze 'GA poza JV' w arkuszu Model).",
 "3. Zasada rozliczeń: partnerzy zarabiają na przychodach, które platforma DOBUDOWUJE (prowizje 3P, reklama, abonamenty), a nie na marży nawozowej GA.",
 "4. Stres test: w Założeniach ustaw mnożnik adopcji na 0,5 i sprawdź, czy kolumna 'GA suma 5 lat' w S2 jest nadal dodatnia.",
 "5. Wszystkie liczby oznaczone jako ZAŁOŻENIE trzeba zastąpić danymi: cena netto (CO-PA), marża AD z umów, oferty Iteo / OLX / Mirakl.",
]
for i, t in enumerate(notes): P.cell(12 + i, 1, t)
P.column_dimensions["A"].width = 58
for c in "BCDEFGHI": P.column_dimensions[c].width = 15
P.row_dimensions[4].height = 45

# ---------------- Rozliczenie (100 zł) ----------------
W = wb.create_sheet("Rozliczenie 100 zł", 1)
W["A1"] = "Kto ile dostaje ze 100 zł - scenariusz S2 (rekomendacja)"; W["A1"].font = Font(bold=True, size=14)
hdr(W, 3, ["A. 100 zł nawozu GA sprzedanego przez platformę", "zł", "Uwaga"])
rows = [
 ("Dziś: marża dystrybutora (AD), której GA nie widzi", f"=100*{N['chan']}", "Kanał tradycyjny"),
 ("Logistyka / fee dla AD jako punktu odbioru", f"=-100*{N['logi']}", "AD zostaje partnerem, nie traci całej marży"),
 ("Opłata dla OLX od GMV nawozów (czerwona linia max)", f"=-100*{N['ecown'][3]}", ""),
 ("Koszt płatności", f"=-100*{N['payc']}", ""),
 ("Poprawa mixu", f"=100*{N['mix']}", ""),
 ("ZOSTAJE W GA (ponad dzisiejszą marżę)", "=SUM(B4:B8)", "Każde 100 zł nawozów przez platformę = tyle zł więcej dla GA"),
]
for i, (a, b, c) in enumerate(rows):
    W.cell(4+i, 1, a); W.cell(4+i, 2, b).number_format = "0.0"; W.cell(4+i, 3, c)
W["A9"].font = B; W["B9"].font = B; W["A9"].fill = TOT; W["B9"].fill = TOT
hdr(W, 11, ["B. 100 zł przychodu od sprzedawców zewnętrznych (prowizja + reklama)", "zł", "Uwaga"])
rows = [
 ("Przychód JV", 100, ""),
 ("OLX (revenue share)", f"=-100*{N['ecshare'][3]}", "OLX zarabia, gdy dowozi sprzedawców i ruch"),
 ("Zostaje w JV na koszty i zysk", "=B12+B13", ""),
 ("w tym GA", f"=B14*{N['gash']}", "Udział GA w JV"),
 ("w tym Iteo", f"=B14*(1-{N['gash']})", "Iteo jako wspólnik JV, z opcją wykupu przez GA"),
]
for i, (a, b, c) in enumerate(rows):
    W.cell(12+i, 1, a); W.cell(12+i, 2, b).number_format = "0.0"; W.cell(12+i, 3, c)
W.column_dimensions["A"].width = 62; W.column_dimensions["B"].width = 10; W.column_dimensions["C"].width = 60

wb.save(OUT)
print("ok", OUT)
