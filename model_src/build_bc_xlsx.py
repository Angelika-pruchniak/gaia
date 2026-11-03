# GAia business case: 300 tys. użytkowników w roku 5. Każda strona: GA, Partner Technologiczny,
# Partner Marketplacowy, 66 dystrybutorów, rolnik. Uruchom: python3 build_bc_xlsx.py
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

OUT = "../GA_GAia_business_case_300k.xlsx"
wb = openpyxl.Workbook()
NAVY = "1F3864"
BLUE = Font(color="0000FF"); B = Font(bold=True); H = Font(bold=True, color="FFFFFF")
HF = PatternFill("solid", fgColor=NAVY); INF = PatternFill("solid", fgColor="FFF2CC"); TOT = PatternFill("solid", fgColor="DDEBF7")
NUM = '#,##0.0;[Red]-#,##0.0'
YRS = ["Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5"]
YC = ["C", "D", "E", "F", "G"]          # lata 1-5 na arkuszach obliczeniowych (B = rok 0)

def hdr(ws, r, vals, c0=1):
    for i, v in enumerate(vals):
        c = ws.cell(r, c0 + i, v); c.font = H; c.fill = HF
        c.alignment = Alignment(wrap_text=True, vertical="center")

# ======================= ZAŁOŻENIA =======================
A = wb.active; A.title = "Założenia"
N = {}
A["A1"] = "GAia business case: założenia (ILUSTRACJA, do potwierdzenia danymi GA i ofertami partnerów)"; A["A1"].font = Font(bold=True, size=14)
A["A2"] = "Niebieskie na żółtym = założenie (zmieniaj). Czarne = formuła. Kwoty w mln zł, jeśli nie podano inaczej. Partnerzy nazwani po funkcji."
ar = 4
hdr(A, ar, ["Założenie", "Rok 1 / wartość", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Jedn.", "Źródło / uwaga"]); ar += 1
def inp(name, label, vals, unit, src, fmt="0.0"):
    global ar
    A.cell(ar, 1, label)
    vals = vals if isinstance(vals, list) else [vals]
    for i, v in enumerate(vals):
        c = A.cell(ar, 2 + i, v); c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE; c.fill = INF
    A.cell(ar, 7, unit); A.cell(ar, 8, src)
    N[name] = f"'Założenia'!$B${ar}" if len(vals) == 1 else [f"'Założenia'!${L(2+i)}${ar}" for i in range(5)]
    ar += 1
def sec(t):
    global ar
    ar += 1; A.cell(ar, 1, t).font = Font(bold=True, color=NAVY, size=12); ar += 1
def y(nm, i):
    v = N[nm]; return v[i] if isinstance(v, list) else v

sec("1. Użytkownicy (cel: 300 tys. w roku 5)")
inp("users", "Użytkownicy zarejestrowani na koniec roku", [30, 80, 150, 230, 300], "tys.", "CEL ZESPOŁU. 300 tys. ≈ 25% z 1,21 mln rolników w ARiMR (2024) / ok. połowa z ~600 tys. rolników zawodowych", "#,##0")
inp("mult", "Stres test: mnożnik użytkowników (1 = plan)", 1.0, "x", "Ustaw 0,5 = połowa planu", "0.00")
inp("act", "Aktywni kupujący (kupują min. raz w roku)", 0.35, "% użytk.", "ZAŁOŻENIE", "0%")
inp("prem", "Abonenci premium", [0.03, 0.05, 0.06, 0.07, 0.08], "% użytk.", "ZAŁOŻENIE: zaawansowany plan nawożenia, mapy, doradca", "0%")
inp("subp", "Cena abonamentu premium", 1.2, "tys. zł/rok", "ZAŁOŻENIE ~100 zł/mies.", "0.0")
inp("bf", "Zakupy nawozów GA przez platformę na kupującego", [6, 7, 8, 8.5, 9], "tys. zł/rok", "ZAŁOŻENIE. Średnie gosp. 11,6 ha (ARiMR 2024); użytkownicy będą więksi", "0.0")
inp("b3", "Zakupy innych produktów (nasiona, ŚOR, części, usługi) na kupującego", [3, 4, 5, 6, 7], "tys. zł/rok", "ZAŁOŻENIE", "0.0")
inp("base", "Sprzedaż krajowa nawozów GA (dla kontroli skali)", "=2.93*0.7*1800", "mln zł/rok", "2,93 mln t × 70% kraj × 1800 zł/t (jak w modelu JV)", "#,##0")

sec("2. Przychody platformy (spółka GAia, JV)")
inp("take", "Prowizja od sprzedawców innych produktów (3P)", 0.07, "% GMV 3P", "ZAŁOŻENIE, typowo 5-12%", "0.0%")
inp("ads", "Reklama i promowane oferty", 0.015, "% GMV 3P", "ZAŁOŻENIE", "0.0%")
inp("fin", "Finansowanie zakupów / ubezpieczenia (prowizja od banku)", 0.004, "% GMV", "ZAŁOŻENIE, por. FBN: >1 mld USD finansowania", "0.0%")
inp("data", "Analizy i dane dla producentów nasion / ŚOR (zanonimizowane)", [0.5, 1, 2, 3.5, 5], "mln zł/rok", "ZAŁOŻENIE", "0.0")
inp("gafee", "Opłata GA dla platformy od sprzedaży nawozów GA", 0.012, "% GMV naw.", "PROPOZYCJA: platforma zarabia też na nawozach, GA płaci z marży kanałowej", "0.0%")

sec("3. Koszty spółki GAia (JV)")
inp("payc", "Płatności i obsługa transakcji", 0.006, "% GMV", "ZAŁOŻENIE", "0.0%")
inp("cac", "Koszt pozyskania nowego użytkownika (CAC)", 60, "zł", "ZAŁOŻENIE; niski dzięki 66 AD, doradcom GA, mObywatel/ARiMR", "0")
inp("brand", "Kampania marki (stała)", 3, "mln zł/rok", "ZAŁOŻENIE", "0.0")
inp("csu", "Obsługa klienta: użytkownicy na 1 konsultanta", 20, "tys.", "ZAŁOŻENIE", "0")
inp("csc", "Koszt konsultanta", 0.12, "mln zł/rok", "ZAŁOŻENIE", "0.00")
inp("mgmt", "Zarząd i zespół spółki GAia", 1.8, "mln zł/rok", "6 osób × 0,3 mln", "0.0")
inp("launch", "Kampania startowa (rok 0)", 2, "mln zł", "ZAŁOŻENIE", "0.0")

sec("4. Partner Technologiczny (SoftwareHouse): budowa w arkuszu CAPEX, tu koszty po starcie (OPEX)")
inp("fmc", "Koszt 1 osobomiesiąca zespołu Partnera Technologicznego", 0.02, "mln zł", "20 tys. zł/mies. pełnego kosztu (programista, QA, UX)", "0.000")
inp("tpfte", "Zespół rozwoju po starcie (nowe funkcje, poprawki, sezon)", [10, 14, 20, 22, 22], "FTE", "ZAŁOŻENIE; w latach 1-2 równolegle trwa budowa z arkusza CAPEX", "0")
inp("tpsup", "Utrzymanie 24/7, SRE, wsparcie 2. linii", [3, 5, 7, 9, 10], "FTE", "rośnie z liczbą użytkowników", "0")
inp("tpsupc", "Koszt 1 FTE utrzymania", 0.15, "mln zł/rok", "ZAŁOŻENIE", "0.00")
inp("cloudf", "Chmura: koszt stały", 0.8, "mln zł/rok", "ZAŁOŻENIE", "0.0")
inp("cloudu", "Chmura: koszt na użytkownika", 6, "zł/rok", "ZAŁOŻENIE", "0")
inp("satu", "Dane satelitarne i pogodowe na użytkownika", 4, "zł/rok", "ZAŁOŻENIE (Sentinel bezpłatny, przetwarzanie i pogoda płatne)", "0")
inp("aiu", "Asystent AI / modele agronomiczne na użytkownika", 3, "zł/rok", "ZAŁOŻENIE", "0")
inp("sec", "Bezpieczeństwo, RODO, NIS2, audyty", 0.6, "mln zł/rok", "ZAŁOŻENIE", "0.0")
inp("tpmarg", "Marża Partnera Technologicznego (koszt + marża) na budowie i utrzymaniu", 0.15, "%", "PROPOZYCJA: Partner nie dokłada do kosztów od 1. dnia", "0%")

sec("5. Partner Marketplacowy / Handlowy")
inp("mpshare", "Udział w prowizji i reklamie od sprzedawców 3P", 0.30, "% przych. 3P", "PROPOZYCJA: płacimy tylko od tego, co dowiezie", "0%")
inp("mpint", "Integracja (rok 0), koszt po stronie Partnera", 1.5, "mln zł", "ZAŁOŻENIE", "0.0")
inp("mpfix", "Zespół i obsługa sprzedawców, koszt stały Partnera", 1.0, "mln zł/rok", "ZAŁOŻENIE", "0.0")
inp("mpvar", "Koszt zmienny Partnera (ruch, wsparcie)", 0.10, "% jego przych.", "ZAŁOŻENIE", "0%")

sec("6. 66 dystrybutorów (AD)")
inp("nad", "Liczba dystrybutorów zaproszonych na platformę", 66, "szt.", "Dane zespołu", "0")
inp("mtr", "Marża AD na nawozach w kanale tradycyjnym", 0.07, "%", "ZAŁOŻENIE - zweryfikować z umowami", "0.0%")
inp("ctr", "Koszt sprzedaży AD w kanale tradycyjnym (handlowcy, kredyt, negocjacje)", 0.03, "%", "ZAŁOŻENIE", "0.0%")
inp("mpl", "Marża AD na zamówieniu z platformy (AD realizuje i wydaje towar)", 0.05, "%", "PROPOZYCJA: zamówienie idzie do lokalnego AD", "0.0%")
inp("cpl", "Koszt AD przy zamówieniu z platformy", 0.01, "%", "Brak negocjacji, kredyt przejmuje bank", "0.0%")
inp("canni", "Część zamówień nawozów, które AD i tak by sprzedał", 0.6, "%", "ZAŁOŻENIE; reszta to nowi rolnicy i udział zabrany konkurencji", "0%")
inp("ad3p", "Udział AD w sprzedaży innych produktów na platformie", 0.5, "%", "AD wystawiają ŚOR, nasiona, części", "0%")
inp("ad3pinc", "Część tej sprzedaży, która jest dla AD nowa", 0.5, "%", "ZAŁOŻENIE", "0%")
inp("m3p", "Marża AD na innych produktach", 0.15, "%", "ZAŁOŻENIE", "0%")
inp("adonb", "Wdrożenie jednego AD (rok 0)", 0.02, "mln zł", "szkolenie, integracja magazynu", "0.00")

sec("7. Rolnik")
inp("save", "Oszczędność na nawozach dzięki precyzyjnemu planowi", 0.05, "% zakupów naw.", "ZAŁOŻENIE; benchmark: narzędzia precyzyjnego nawożenia", "0%")
inp("price", "Przejrzystość cen i oszczędność czasu", 0.01, "% wszystkich zakupów", "ZAŁOŻENIE", "0%")

sec("8. Grupa Azoty")
inp("chan", "Lepsza cena netto GA (marża AD 7% → 5%)", "=B{0}-B{1}".format(0, 0), "% GMV naw.", "Formuła: marża AD tradycyjnie - na platformie", "0.0%")
chan_row = ar - 1
inp("mix", "Poprawa mixu (NPK, specjalistyczne)", 0.01, "% GMV naw.", "ZAŁOŻENIE", "0.0%")
inp("pp", "Dodatkowy udział w rynku AN/CAN", [0.3, 0.8, 1.5, 2.2, 3.0], "p.p.", "ZAŁOŻENIE, cel strategii 45% → 60%", "0.0")
inp("ppv", "Przychód z 1 p.p. udziału AN/CAN", 39, "mln zł/rok", "GA_5_luk, arkusz Szacunki", "0")
inp("cm", "Marża kontrybucyjna na dodatkowej tonie", 0.25, "%", "ZAŁOŻENIE - Kontroling", "0%")
inp("gaint0", "Integracja z SAP / CRM po stronie GA: rok 0", 2, "mln zł", "ZAŁOŻENIE", "0.0")
inp("gaint1", "Integracja z SAP / CRM po stronie GA: rok 1", 1, "mln zł", "ZAŁOŻENIE", "0.0")
inp("gaown", "Utrzymanie 5 funkcjonalności GA i zespół GA", 2, "mln zł/rok", "ZAŁOŻENIE", "0.0")

sec("9. Struktura spółki GAia i finanse")
inp("gash", "Udział GA w spółce GAia", 0.60, "%", "PROPOZYCJA 60/40 wg wkładów (arkusz Wkład stron)", "0%")
inp("wacc", "Stopa dyskontowa", 0.10, "%", "ZAŁOŻENIE", "0%")
inp("f5", "Wartość 5 gotowych funkcjonalności GA (koszt odtworzenia)", 12.5, "mln zł", "ZAŁOŻENIE: 5 × 2,5 mln zł", "0.0")

sec("10. Moc platformy (do przeliczenia GMV na zamówienia i tony)")
inp("aovf", "Średnie zamówienie nawozów", 4.5, "tys. zł", "ok. 2,5 t × 1800 zł/t", "0.0")
inp("aov3", "Średnie zamówienie innych produktów", 1.2, "tys. zł", "ZAŁOŻENIE", "0.0")
inp("pt", "Cena nawozu do przeliczenia na tony", 1800, "zł/t", "jak w modelu JV", "#,##0")
inp("seas", "Część zamówień nawozów w szczycie sezonu", 0.6, "%", "wiosna", "0%")
inp("seasd", "Dni szczytu sezonu", 90, "dni", "", "0")
inp("sell", "Inni sprzedawcy na platformie (poza 66 AD)", [20, 60, 120, 180, 250], "szt.", "producenci nasion, ŚOR, części, usługodawcy", "0")

# naprawa formuły chan
mtr_r = int(N["mtr"].split("$")[-1]); mpl_r = int(N["mpl"].split("$")[-1])
A.cell(chan_row, 2, f"=B{mtr_r}-B{mpl_r}")
A.column_dimensions["A"].width = 66
for c in "BCDEF": A.column_dimensions[c].width = 12
A.column_dimensions["G"].width = 14; A.column_dimensions["H"].width = 80
A.freeze_panes = "A5"


# ======================= CAPEX: BUDOWA PLATFORMY =======================
C = wb.create_sheet("CAPEX budowa")
C["A1"] = "CAPEX: budowa platformy i marketplace'u moduł po module (koszt Partnera Technologicznego, bez marży)"; C["A1"].font = Font(bold=True, size=14)
C["A2"] = "Osobomiesiące i harmonogram (niebieskie) to założenia do potwierdzenia ofertą Partnera. Koszt = osobomiesiące × koszt osobomiesiąca z Założeń."
hdr(C, 4, ["Moduł", "Co obejmuje", "Osobo-\nmiesiące", "% w roku 0", "% w roku 1", "% w roku 2", "Koszt razem\nmln zł", "Rok 0", "Rok 1", "Rok 2"])
mods = [
 ("Architektura, chmura, DevOps", "środowiska, CI/CD, monitoring, kopie zapasowe", 18, 1, 0, 0),
 ("Silnik agro + 5 funkcjonalności GA", "adaptacja silnika Partnera, podłączenie planu nawożenia, CRM, portalu AD, Nawozy.eu, koszyka", 60, .7, .3, 0),
 ("Katalog produktów i panel sprzedawcy", "karty produktów, ceny, stany, oferty sprzedawców, moderacja", 48, .6, .4, 0),
 ("Koszyk, płatności, faktury", "checkout, płatność podzielona między sprzedawców, faktury, zwroty", 42, .7, .3, 0),
 ("Zamówienia do 66 AD", "routing do lokalnego AD, statusy, odbiory, integracja magazynów AD", 54, .6, .4, 0),
 ("Aplikacja mobilna iOS / Android", "zakupy, plan nawożenia, powiadomienia w polu", 54, .3, .7, 0),
 ("Integracja z SAP / CRM GA", "zamówienia, ceny, dostępność, dane klientów (strona platformy)", 24, .7, .3, 0),
 ("mObywatel / login.gov.pl, dane ARiMR", "logowanie rolnika, działki i uprawy z wniosku", 18, 0, 1, 0),
 ("Mapy pól i dane satelitarne", "NDVI, strefy, plan nawożenia premium", 36, 0, .6, .4),
 ("Asystent AI agronomiczny", "pytania rolnika, rekomendacje produktów", 24, 0, .3, .7),
 ("Abonamenty, reklama, panel danych", "premium, promowane oferty, raporty dla producentów", 30, 0, .4, .6),
 ("Finansowanie zakupów", "integracja z bankiem: kredyt kupiecki w koszyku", 18, 0, .3, .7),
 ("Bezpieczeństwo i zgodność", "testy penetracyjne, RODO, NIS2", 18, .4, .3, .3),
 ("Testy z rolnikami i AD", "QA, testy akceptacyjne, pilot", 24, .4, .4, .2),
 ("Kierowanie projektem, UX/UI, analityka", "PM, projekt ekranów, analityka biznesowa", 36, .4, .35, .25),
]
cr = 5
for name, what, fm, a0, a1, a2 in mods:
    C.cell(cr, 1, name); C.cell(cr, 2, what)
    for j, v in enumerate([fm, a0, a1, a2]):
        c = C.cell(cr, 3 + j, v); c.font = BLUE; c.fill = INF; c.number_format = "0" if j == 0 else "0%"
    C.cell(cr, 7, f"=C{cr}*{N['fmc']}")
    for j, col in enumerate("DEF"): C.cell(cr, 8 + j, f"=$G{cr}*{col}{cr}")
    for j in range(7, 11): C.cell(cr, j).number_format = NUM
    cr += 1
first, last = 5, cr - 1
C.cell(cr, 1, "Zespół budowy razem").font = B; C.cell(cr, 3, f"=SUM(C{first}:C{last})").font = B
for col in "GHIJ": C[f"{col}{cr}"] = f"=SUM({col}{first}:{col}{last})"; C[f"{col}{cr}"].font = B; C[f"{col}{cr}"].number_format = NUM
team_r = cr; cr += 1
C.cell(cr, 1, "Koszty zewnętrzne: licencje narzędzi, testy penetracyjne, sklepy z aplikacjami")
for j, v in enumerate([0.5, 0.3, 0.2]):
    c = C.cell(cr, 8 + j, v); c.font = BLUE; c.fill = INF; c.number_format = NUM
C[f"G{cr}"] = f"=SUM(H{cr}:J{cr})"; C[f"G{cr}"].number_format = NUM
ext_r = cr; cr += 1
C.cell(cr, 1, "CAPEX: KOSZT BUDOWY PO STRONIE PARTNERA TECHNOLOGICZNEGO").font = B
for col in "GHIJ": C[f"{col}{cr}"] = f"={col}{team_r}+{col}{ext_r}"; C[f"{col}{cr}"].font = B; C[f"{col}{cr}"].fill = TOT; C[f"{col}{cr}"].number_format = NUM
cap_r = cr; cr += 1
C.cell(cr, 1, "Marża Partnera (15%)")
for col in "GHIJ": C[f"{col}{cr}"] = f"={col}{cap_r}*{N['tpmarg']}"; C[f"{col}{cr}"].number_format = NUM
capm_r = cr; cr += 1
C.cell(cr, 1, "CAPEX PŁACI SPÓŁKA GAia (koszt + marża)").font = B
for col in "GHIJ": C[f"{col}{cr}"] = f"={col}{cap_r}+{col}{capm_r}"; C[f"{col}{cr}"].font = B; C[f"{col}{cr}"].fill = TOT; C[f"{col}{cr}"].number_format = NUM
capj_r = cr; cr += 2
C.cell(cr, 1, "Dla porównania: wartość 5 funkcjonalności GA, których nie trzeba budować"); C.cell(cr, 7, f"={N['f5']}").number_format = NUM
cr += 1
C.cell(cr, 1, "Amortyzacja: liniowo przez 4 lata, od roku po wydatku (arkusz Spółka GAia)")
C.column_dimensions["A"].width = 44; C.column_dimensions["B"].width = 62
for col in "CDEFGHIJ": C.column_dimensions[col].width = 11
C.row_dimensions[4].height = 32
CAPC = {0: f"'CAPEX budowa'!$H${cap_r}", 1: f"'CAPEX budowa'!$I${cap_r}", 2: f"'CAPEX budowa'!$J${cap_r}"}
CAPJ = {0: f"'CAPEX budowa'!$H${capj_r}", 1: f"'CAPEX budowa'!$I${capj_r}", 2: f"'CAPEX budowa'!$J${capj_r}"}
CAPTOT = f"'CAPEX budowa'!$G${cap_r}"
CAPFM = f"'CAPEX budowa'!$C${team_r}"

# ======================= silnik arkuszy =======================
class Sheet:
    def __init__(self, title, head):
        self.ws = wb.create_sheet(title); self.r = 5; self.R = {}
        self.ws["A1"] = head; self.ws["A1"].font = Font(bold=True, size=14)
        hdr(self.ws, 4, ["Pozycja"] + YRS + ["Suma 0-5"])
        for i, col in enumerate(YC): self.ws[f"{col}3"] = i + 1; self.ws[f"{col}3"].font = Font(color="FFFFFF")
        self.ws.column_dimensions["A"].width = 60
        for c in "BCDEFGH": self.ws.column_dimensions[c].width = 12
        self.ws.freeze_panes = "B5"
    def row(self, key, label, f, y0=None, bold=False, fill=None, total=True, fmt=NUM):
        ws = self.ws; ws.cell(self.r, 1, label)
        if y0 is not None: ws.cell(self.r, 2, "=" + y0)
        for i, col in enumerate(YC): ws[f"{col}{self.r}"] = "=" + f(i, col)
        if total: ws[f"H{self.r}"] = f"=SUM(B{self.r}:G{self.r})"
        for c in range(1, 9):
            ce = ws.cell(self.r, c); ce.number_format = fmt
            if bold: ce.font = B
            if fill: ce.fill = fill
        self.R[key] = self.r; self.r += 1
        return self.r - 1
    def sec(self, t):
        self.r += 1; self.ws.cell(self.r, 1, t).font = Font(bold=True, color=NAVY, size=12); self.r += 1
    def ref(self, key, col): return f"'{self.ws.title}'!{col}{self.R[key]}"
    def cum(self, key, label):
        src = self.R[key]; rr = self.row(key + "_cum", label, lambda i, c: "0", y0="0", total=False)
        self.ws[f"B{rr}"] = f"=B{src}"
        for i, col in enumerate(YC):
            self.ws[f"{col}{rr}"] = f"={'B' if i == 0 else YC[i-1]}{rr}+{col}{src}"

# ---------- Użytkownicy i GMV ----------

# ---------- Użytkownicy i GMV ----------
U = Sheet("Użytkownicy i GMV", "Użytkownicy, obroty i zamówienia: z czego wynika GMV")
mu = N["mult"]
U.sec("Użytkownicy")
U.row("u", "Użytkownicy zarejestrowani (tys.)", lambda i, c: f"{y('users',i)}*{mu}", total=False, fmt="#,##0")
U.row("nu", "Nowi użytkownicy w roku (tys.)", lambda i, c: f"{c}{U.R['u']}-{'0' if i == 0 else YC[i-1] + str(U.R['u'])}", fmt="#,##0")
U.row("buy", "Aktywni kupujący (tys.) = użytkownicy × 35%", lambda i, c: f"{c}{U.R['u']}*{N['act']}", total=False, fmt="#,##0")
U.row("pr", "Abonenci premium (tys.)", lambda i, c: f"{c}{U.R['u']}*{y('prem',i)}", total=False, fmt="#,##0.0")
U.sec("Obroty (GMV), mln zł = kupujący × wartość zakupów")
U.row("bf", "Zakupy nawozów na kupującego (tys. zł/rok)", lambda i, c: f"{y('bf',i)}", total=False)
U.row("b3", "Zakupy innych produktów na kupującego (tys. zł/rok)", lambda i, c: f"{y('b3',i)}", total=False)
U.row("fg", "GMV nawozów GA", lambda i, c: f"{c}{U.R['buy']}*{c}{U.R['bf']}")
U.row("g3", "GMV innych produktów od sprzedawców (3P, w tym AD)", lambda i, c: f"{c}{U.R['buy']}*{c}{U.R['b3']}")
U.row("gmv", "GMV RAZEM", lambda i, c: f"{c}{U.R['fg']}+{c}{U.R['g3']}", bold=True, fill=TOT)
U.row("pen", "Nawozy przez platformę / sprzedaż krajowa GA", lambda i, c: f"{c}{U.R['fg']}/{N['base']}", total=False, fmt="0.0%")
U.sec("Zamówienia i tony")
U.row("of", "Zamówienia nawozów (tys.)", lambda i, c: f"{c}{U.R['fg']}/{N['aovf']}", fmt="#,##0.0")
U.row("o3", "Zamówienia innych produktów (tys.)", lambda i, c: f"{c}{U.R['g3']}/{N['aov3']}", fmt="#,##0.0")
U.row("ot", "Zamówienia razem (tys.)", lambda i, c: f"{c}{U.R['of']}+{c}{U.R['o3']}", bold=True, fmt="#,##0.0")
U.row("tn", "Nawozy przez platformę (tys. t)", lambda i, c: f"{c}{U.R['fg']}*1000/{N['pt']}", fmt="#,##0")
U.row("oad", "Zamówienia nawozów na 1 AD dziennie w szczycie sezonu", lambda i, c: f"{c}{U.R['of']}*1000*{N['seas']}/{N['nad']}/{N['seasd']}", total=False, fmt="#,##0.0")
U.row("sel", "Sprzedawcy na platformie (66 AD + inni)", lambda i, c: f"{N['nad']}+{y('sell',i)}", total=False, fmt="#,##0")
u = lambda k, c: U.ref(k, c)
COLS = ["B"] + YC

# ---------- Partner Technologiczny ----------
T = Sheet("Partner Technologiczny", "Partner Technologiczny (SoftwareHouse): koszty po jego stronie, z czego i kiedy zarabia")
T.sec("A. Koszty: budowa (CAPEX, z arkusza CAPEX budowa)")
T.row("cap", "Budowa platformy i marketplace'u (koszt)", lambda i, c: CAPC[i + 1] if i + 1 in CAPC else "0", y0=CAPC[0])
T.sec("B. Koszty: utrzymanie i rozwój po starcie (OPEX)")
T.row("dev", "Zespół rozwoju", lambda i, c: f"{y('tpfte',i)}*12*{N['fmc']}")
T.row("sup", "Utrzymanie 24/7 i wsparcie", lambda i, c: f"{y('tpsup',i)}*{N['tpsupc']}")
T.row("cloud", "Chmura (stała + na użytkownika)", lambda i, c: f"{N['cloudf']}+{u('u',c)}*1000*{N['cloudu']}/1000000")
T.row("sat", "Dane satelitarne i pogodowe", lambda i, c: f"{u('u',c)}*1000*{N['satu']}/1000000")
T.row("ai", "Asystent AI / modele agronomiczne", lambda i, c: f"{u('u',c)}*1000*{N['aiu']}/1000000")
T.row("sec", "Bezpieczeństwo i zgodność", lambda i, c: N["sec"])
T.row("opex", "OPEX Partnera razem", lambda i, c: f"SUM({c}{T.R['dev']}:{c}{T.R['sec']})", bold=True)
T.row("cost", "KOSZTY PARTNERA RAZEM (CAPEX + OPEX)", lambda i, c: f"{c}{T.R['cap']}+{c}{T.R['opex']}", y0=f"B{T.R['cap']}", bold=True, fill=TOT)
T.row("cpu", "OPEX technologii na użytkownika (zł/rok)", lambda i, c: f"{c}{T.R['opex']}/{u('u',c)}*1000", total=False, fmt="#,##0")
T.sec("C. Przychody Partnera od spółki GAia (koszt + 15%)")
T.row("fcap", "Za budowę", lambda i, c: f"{c}{T.R['cap']}*(1+{N['tpmarg']})", y0=f"B{T.R['cap']}*(1+{N['tpmarg']})")
T.row("fopx", "Za utrzymanie i rozwój", lambda i, c: f"{c}{T.R['opex']}*(1+{N['tpmarg']})")
T.row("fee", "Przychody od spółki GAia razem", lambda i, c: f"{c}{T.R['fcap']}+{c}{T.R['fopx']}", y0=f"B{T.R['fcap']}", bold=True)
T.sec("D. Z czego i kiedy zarabia Partner Technologiczny")
T.row("mcap", "1. Marża na budowie (lata 0-2)", lambda i, c: f"{c}{T.R['fcap']}-{c}{T.R['cap']}", y0=f"B{T.R['fcap']}-B{T.R['cap']}")
T.row("mopx", "2. Marża na utrzymaniu (lata 1-5)", lambda i, c: f"{c}{T.R['fopx']}-{c}{T.R['opex']}")


# ---------- Spółka GAia ----------
J = Sheet("Spółka GAia (JV)", "Spółka GAia (JV): przychody, OPEX, EBITDA, CAPEX, amortyzacja, przepływy pieniężne")
J.sec("Przychody (z czego wynikają: GMV × stawka)")
J.row("rc", "Prowizje: GMV 3P × 7%", lambda i, c: f"{u('g3',c)}*{N['take']}")
J.row("ra", "Reklama: GMV 3P × 1,5%", lambda i, c: f"{u('g3',c)}*{N['ads']}")
J.row("rs", "Abonamenty: abonenci premium × 1 200 zł", lambda i, c: f"{u('pr',c)}*{N['subp']}")
J.row("rf", "Finansowanie: GMV × 0,4%", lambda i, c: f"{u('gmv',c)}*{N['fin']}")
J.row("rd", "Dane i analizy dla producentów", lambda i, c: f"{y('data',i)}")
J.row("rg", "Opłata GA: GMV nawozów × 1,2%", lambda i, c: f"{u('fg',c)}*{N['gafee']}")
J.row("rev", "PRZYCHODY RAZEM", lambda i, c: f"SUM({c}{J.R['rc']}:{c}{J.R['rg']})", bold=True, fill=TOT)
J.sec("OPEX")
J.row("ctp", "Partner Technologiczny: utrzymanie i rozwój (koszt + 15%)", lambda i, c: T.ref("fopx", c))
J.row("cmp", "Partner Marketplacowy: 30% prowizji i reklamy", lambda i, c: f"({c}{J.R['rc']}+{c}{J.R['ra']})*{N['mpshare']}")
J.row("cpay", "Płatności: GMV × 0,6%", lambda i, c: f"{u('gmv',c)}*{N['payc']}")
J.row("cmkt", "Marketing: nowi użytkownicy × 60 zł + marka 3 mln zł", lambda i, c: f"{u('nu',c)}*1000*{N['cac']}/1000000+{N['brand']}", y0=N["launch"])
J.row("ccs", "Obsługa klienta: 1 konsultant na 20 tys. użytkowników", lambda i, c: f"{u('u',c)}/{N['csu']}*{N['csc']}")
J.row("cmg", "Zarząd i zespół spółki (6 osób)", lambda i, c: N["mgmt"], y0=f"{N['mgmt']}*0.5")
J.row("opex", "OPEX RAZEM", lambda i, c: f"SUM({c}{J.R['ctp']}:{c}{J.R['cmg']})", y0=f"SUM(B{J.R['ctp']}:B{J.R['cmg']})", bold=True)
J.row("ebitda", "EBITDA", lambda i, c: f"{c}{J.R['rev']}-{c}{J.R['opex']}", y0=f"-B{J.R['opex']}", bold=True, fill=TOT)
J.row("emg", "Marża EBITDA", lambda i, c: f"IF({c}{J.R['rev']}=0,0,{c}{J.R['ebitda']}/{c}{J.R['rev']})", total=False, fmt="0%")
J.sec("CAPEX i amortyzacja")
J.row("capex", "CAPEX: budowa platformy (Partner Technologiczny, koszt + 15%)", lambda i, c: CAPJ[i + 1] if i + 1 in CAPJ else "0", y0=CAPJ[0])
def amort(i, c):
    n = i + 1; prev = [COLS[k] for k in range(max(0, n - 4), n)]
    return "(" + "+".join(f"{p}{J.R['capex']}" for p in prev) + ")/4"
J.row("am", "Amortyzacja (4 lata, od roku po wydatku)", amort)
J.row("ebit", "EBIT (wynik przed podatkiem)", lambda i, c: f"{c}{J.R['ebitda']}-{c}{J.R['am']}", y0=f"B{J.R['ebitda']}", bold=True)
J.sec("Przepływy pieniężne")
J.row("fcf", "PRZEPŁYWY PIENIĘŻNE (EBITDA - CAPEX)", lambda i, c: f"{c}{J.R['ebitda']}-{c}{J.R['capex']}", y0=f"B{J.R['ebitda']}-B{J.R['capex']}", bold=True, fill=TOT)
J.cum("fcf", "Przepływy skumulowane")
fr = J.r + 1
J.ws.cell(fr, 1, "Maksymalne zapotrzebowanie na finansowanie").font = B
J.ws.cell(fr, 2, f"=-MIN(0,MIN(B{J.R['fcf_cum']}:G{J.R['fcf_cum']}))").number_format = NUM
J.ws.cell(fr + 1, 1, "Wkład Grupy Azoty (60%)"); J.ws.cell(fr + 1, 2, f"=B{fr}*{N['gash']}").number_format = NUM
J.ws.cell(fr + 2, 1, "Wkład Partnera Technologicznego (40%)"); J.ws.cell(fr + 2, 2, f"=B{fr}*(1-{N['gash']})").number_format = NUM
FUND = f"'Spółka GAia (JV)'!$B${fr}"; FUNDGA = f"'Spółka GAia (JV)'!$B${fr+1}"; FUNDTP = f"'Spółka GAia (JV)'!$B${fr+2}"

# Partner Technologiczny: trzecie źródło
T.row("jv", "3. 40% przepływów spółki GAia (dopłaty na start, potem dywidenda)", lambda i, c: f"{J.ref('fcf',c)}*(1-{N['gash']})", y0=f"{J.ref('fcf','B')}*(1-{N['gash']})")
T.row("res", "WYNIK PARTNERA TECHNOLOGICZNEGO (1 + 2 + 3)", lambda i, c: f"{c}{T.R['mcap']}+{c}{T.R['mopx']}+{c}{T.R['jv']}", y0=f"B{T.R['mcap']}+B{T.R['mopx']}+B{T.R['jv']}", bold=True, fill=TOT)
T.cum("res", "Wynik skumulowany")

# ---------- Partner Marketplacowy ----------
Mp = Sheet("Partner Marketplacowy", "Partner Marketplacowy / Handlowy: z czego i kiedy zarabia")
Mp.sec("Z czego wynika przychód Partnera")
Mp.row("g3", "GMV innych produktów od sprzedawców (3P)", lambda i, c: u("g3", c))
Mp.row("rc", "Prowizja platformy (7% GMV 3P)", lambda i, c: J.ref("rc", c))
Mp.row("ra", "Reklama (1,5% GMV 3P)", lambda i, c: J.ref("ra", c))
Mp.row("src", "1. Partner: 30% prowizji", lambda i, c: f"{c}{Mp.R['rc']}*{N['mpshare']}")
Mp.row("sra", "2. Partner: 30% reklamy", lambda i, c: f"{c}{Mp.R['ra']}*{N['mpshare']}")
Mp.row("rev", "PRZYCHÓD PARTNERA (1 + 2)", lambda i, c: f"{c}{Mp.R['src']}+{c}{Mp.R['sra']}", bold=True)
Mp.row("nf", "Od nawozów GA Partner dostaje", lambda i, c: "0")
Mp.sec("Koszty po stronie Partnera")
Mp.row("int", "Integracja", lambda i, c: "0", y0=N["mpint"])
Mp.row("fix", "Zespół i obsługa sprzedawców", lambda i, c: N["mpfix"])
Mp.row("var", "Ruch i wsparcie (10% przychodu)", lambda i, c: f"{c}{Mp.R['rev']}*{N['mpvar']}")
Mp.row("res", "WYNIK PARTNERA MARKETPLACOWEGO", lambda i, c: f"{c}{Mp.R['rev']}-{c}{Mp.R['int']}-{c}{Mp.R['fix']}-{c}{Mp.R['var']}", y0=f"-B{Mp.R['int']}", bold=True, fill=TOT)
Mp.cum("res", "Wynik skumulowany")
Mp.row("pct", "Udział Partnera w przychodach platformy", lambda i, c: f"{c}{Mp.R['rev']}/{J.ref('rev',c)}", total=False, fmt="0%")
Mp.ws.cell(Mp.r + 1, 1, "Ze 100 zł sprzedaży innych produktów Partner dostaje (zł)"); Mp.ws.cell(Mp.r + 1, 2, f"=100*({N['take']}+{N['ads']})*{N['mpshare']}").number_format = "0.00"

# ---------- Dystrybutorzy ----------
D = Sheet("Dystrybutorzy (66)", "66 dystrybutorów (AD): co zyskują na platformie")
D.row("same", "Nawozy, które AD i tak by sprzedał: zmiana zysku na tym wolumenie", lambda i, c: f"{u('fg',c)}*{N['canni']}*(({N['mpl']}-{N['cpl']})-({N['mtr']}-{N['ctr']}))")
D.row("inc", "Nowi rolnicy i udział zabrany konkurencji: zysk z nowego wolumenu", lambda i, c: f"{u('fg',c)}*(1-{N['canni']})*({N['mpl']}-{N['cpl']})")
D.row("p3", "Własna sprzedaż innych produktów na platformie (nowa)", lambda i, c: f"{u('g3',c)}*{N['ad3p']}*{N['ad3pinc']}*({N['m3p']}-{N['take']}-{N['cpl']})")
D.row("onb", "Wdrożenie", lambda i, c: "0", y0=f"-{N['nad']}*{N['adonb']}")
D.row("res", "WYNIK 66 AD RAZEM", lambda i, c: f"SUM({c}{D.R['same']}:{c}{D.R['p3']})", y0=f"B{D.R['onb']}", bold=True, fill=TOT)
D.row("per", "Na jednego AD (tys. zł)", lambda i, c: f"{c}{D.R['res']}/{N['nad']}*1000", y0=f"B{D.R['res']}/{N['nad']}*1000", fmt="#,##0")
D.ws.cell(D.r + 1, 1, "Zasada: zamówienie z platformy trafia do lokalnego AD. AD ma niższą marżę (5% zamiast 7%), ale też niższy koszt sprzedaży, więc zarabia tyle samo na tonie i dostaje nowy wolumen.")

# ---------- Rolnik ----------
F = Sheet("Rolnik", "Rolnik: korzyść z platformy")
F.row("sv", "Oszczędność na nawozach (precyzyjny plan, 5%)", lambda i, c: f"{u('fg',c)}*{N['save']}")
F.row("pt", "Przejrzystość cen i oszczędność czasu (1%)", lambda i, c: f"{u('gmv',c)}*{N['price']}")
F.row("sub", "Koszt abonamentu premium (dla chętnych)", lambda i, c: f"-{J.ref('rs',c)}")
F.row("res", "KORZYŚĆ ROLNIKÓW RAZEM", lambda i, c: f"SUM({c}{F.R['sv']}:{c}{F.R['sub']})", bold=True, fill=TOT)
F.row("per", "Na aktywnego kupującego (zł/rok)", lambda i, c: f"{c}{F.R['res']}/{u('buy',c)}*1000", total=False, fmt="#,##0")

# ---------- Grupa Azoty ----------
G = Sheet("Grupa Azoty", "Grupa Azoty: wynik z platformy")
G.row("ch", "Lepsza cena netto nawozów: GMV nawozów × 2% (marża AD 7% → 5%)", lambda i, c: f"{u('fg',c)}*'Założenia'!$B${chan_row}")
G.row("mx", "Poprawa mixu: GMV nawozów × 1%", lambda i, c: f"{u('fg',c)}*{N['mix']}")
G.row("sh", "Udział w rynku AN/CAN: p.p. × 39 mln zł × 25% marży", lambda i, c: f"{y('pp',i)}*{mu}*{N['ppv']}*{N['cm']}")
G.row("fee", "Opłata dla spółki GAia: GMV nawozów × 1,2%", lambda i, c: f"-{J.ref('rg',c)}")
G.row("own", "Integracja SAP/CRM (CAPEX GA) i utrzymanie 5 funkcjonalności", lambda i, c: f"-{N['gaown']}" + (f"-{N['gaint1']}" if i == 0 else ""), y0=f"-{N['gaint0']}")
G.row("jv", "60% przepływów spółki GAia (dopłaty na start, potem dywidenda)", lambda i, c: f"{J.ref('fcf',c)}*{N['gash']}", y0=f"{J.ref('fcf','B')}*{N['gash']}")
G.row("res", "WYNIK GRUPY AZOTY", lambda i, c: f"SUM({c}{G.R['ch']}:{c}{G.R['jv']})", y0=f"SUM(B{G.R['ch']}:B{G.R['jv']})", bold=True, fill=TOT)
G.cum("res", "Wynik skumulowany")

# ---------- Moc platformy ----------
Mc = Sheet("Moc platformy", "Moc platformy w liczbach")
Mc.row("u", "Użytkownicy (tys.)", lambda i, c: u("u", c), total=False, fmt="#,##0")
Mc.row("b", "Aktywni kupujący (tys.)", lambda i, c: u("buy", c), total=False, fmt="#,##0")
Mc.row("p", "Abonenci premium (tys.)", lambda i, c: u("pr", c), total=False, fmt="#,##0.0")
Mc.row("s", "Sprzedawcy (66 AD + inni)", lambda i, c: u("sel", c), total=False, fmt="#,##0")
Mc.row("g", "GMV razem (mln zł)", lambda i, c: u("gmv", c))
Mc.row("o", "Zamówienia razem (tys.)", lambda i, c: u("ot", c), fmt="#,##0")
Mc.row("t", "Nawozy przez platformę (tys. t)", lambda i, c: u("tn", c), fmt="#,##0")
Mc.row("ad", "Zamówienia nawozów na 1 AD dziennie w szczycie", lambda i, c: u("oad", c), total=False, fmt="#,##0.0")
Mc.row("r", "Przychody spółki GAia (mln zł)", lambda i, c: J.ref("rev", c))
Mc.row("tr", "Przychód platformy / GMV", lambda i, c: f"{J.ref('rev',c)}/{u('gmv',c)}", total=False, fmt="0.0%")
Mc.row("ru", "Przychód na użytkownika (zł/rok)", lambda i, c: f"{J.ref('rev',c)}/{u('u',c)}*1000", total=False, fmt="#,##0")
Mc.row("cu", "OPEX na użytkownika (zł/rok)", lambda i, c: f"{J.ref('opex',c)}/{u('u',c)}*1000", total=False, fmt="#,##0")
Mc.row("eu", "EBITDA na użytkownika (zł/rok)", lambda i, c: f"{J.ref('ebitda',c)}/{u('u',c)}*1000", total=False, fmt="#,##0")


# ======================= PODSUMOWANIE =======================
P = wb.create_sheet("Każdy zyskuje", 0)
P["A1"] = "GAia: 300 tys. użytkowników w roku 5. Każdy uczestnik zarabia"; P["A1"].font = Font(bold=True, size=14)
P["A2"] = "Przepływy pieniężne (gotówka), mln zł. Partnerzy spółki GAia liczą swój udział w EBITDA - CAPEX. Zmień założenia (żółte pola), wszystko przeliczy się samo."
hdr(P, 4, ["Uczestnik", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma 0-5", "NPV (10%)", "Wynik roczny > 0 od", "Zwrot (suma > 0) od"])
parts = [("Grupa Azoty", G, "res"), ("Partner Technologiczny (SoftwareHouse)", T, "res"),
         ("Partner Marketplacowy / Handlowy", Mp, "res"), ("66 dystrybutorów (AD) razem", D, "res"),
         ("Rolnicy (korzyść)", F, "res"), ("Spółka GAia (JV) - przepływy pieniężne", J, "fcf")]
for k, (name, sh, key) in enumerate(parts):
    pr = 5 + k; P.cell(pr, 1, name)
    for j, col in enumerate(["B"] + YC):
        P.cell(pr, 2 + j, f"={sh.ref(key, col)}")
    P.cell(pr, 8, f"=SUM(B{pr}:G{pr})")
    P.cell(pr, 9, f"=B{pr}+NPV({N['wacc']},C{pr}:G{pr})")
    P.cell(pr, 10, f'=IF(COUNTIF(C{pr}:G{pr},">0")=0,"po roku 5","rok "&SUMPRODUCT(MIN(IF(C{pr}:G{pr}>0,{{1,2,3,4,5}},99))))')
    P.cell(pr, 11, f'=IF(B{pr}+C{pr}>0,"rok 1",IF(SUM(B{pr}:D{pr})>0,"rok 2",IF(SUM(B{pr}:E{pr})>0,"rok 3",IF(SUM(B{pr}:F{pr})>0,"rok 4",IF(SUM(B{pr}:G{pr})>0,"rok 5","po roku 5")))))')
    for c in range(2, 10): P.cell(pr, c).number_format = NUM
    if k == 0:
        for c in range(1, 12): P.cell(pr, c).font = B; P.cell(pr, c).fill = TOT
P.cell(12, 1, "Skala").font = B
kpis = [("Użytkownicy (tys.)", "u", "#,##0"), ("Aktywni kupujący (tys.)", "buy", "#,##0"), ("GMV razem (mln zł)", "gmv", NUM),
        ("Nawozy GA przez platformę (mln zł)", "fg", NUM), ("Udział platformy w sprzedaży krajowej GA", "pen", "0.0%")]
for k, (lab, key, fm) in enumerate(kpis):
    pr = 13 + k; P.cell(pr, 1, lab)
    for j, col in enumerate(YC): P.cell(pr, 3 + j, f"={U.ref(key, col)}").number_format = fm
P.cell(19, 1, "Inwestycja i finansowanie").font = B
inv = [("CAPEX budowy: koszt Partnera Technologicznego (lata 0-2)", CAPTOT), ("CAPEX GA: integracja SAP/CRM (lata 0-1)", f"{N['gaint0']}+{N['gaint1']}"),
       ("Maks. zapotrzebowanie spółki GAia na finansowanie", FUND), ("w tym Grupa Azoty (60%)", FUNDGA), ("w tym Partner Technologiczny (40%)", FUNDTP)]
for k, (lab, f) in enumerate(inv):
    pr = 20 + k; P.cell(pr, 1, lab); P.cell(pr, 3, "=" + f).number_format = NUM
P.column_dimensions["A"].width = 56
for c in "BCDEFGHIJK": P.column_dimensions[c].width = 12

# ======================= WKŁAD STRON =======================
Wk = wb.create_sheet("Wkład stron", 1)
Wk["A1"] = "Kto co wnosi i jak się rozliczamy"; Wk["A1"].font = Font(bold=True, size=14)
hdr(Wk, 3, ["Uczestnik", "Co wnosi", "Jak zarabia", "Ryzyko, które bierze"])
rows = [
 ("Grupa Azoty", "5 gotowych funkcjonalności (do potwierdzenia: plan nawożenia GAia, CRM z profilami gospodarstw, portal AD, Nawozy.eu / doradztwo, koszyk → zamówienie do AD). 66 dystrybutorów. Produkty, produkcja, logistyka, marka, doradcy w regionach.",
  "60% wyniku spółki GAia + lepsza cena netto nawozów + mix + udział w rynku", "Integracja SAP/CRM, 60% finansowania startu spółki"),
 ("Partner Technologiczny (SoftwareHouse)", "Silnik platformy agro, zespół rozwoju, utrzymanie 24/7, chmura, dane satelitarne, AI",
  "Koszt + 15% na budowie i utrzymaniu + 40% przepływów spółki GAia", "40% finansowania startu spółki; ryzyko technologiczne"),
 ("Partner Marketplacowy / Handlowy", "Silnik e-commerce, sprzedawcy innych produktów, płatności, ruch, doświadczenie marketplace",
  "30% prowizji i reklamy od sprzedawców 3P. Bez opłat od nawozów GA", "Koszt integracji i zespołu"),
 ("66 dystrybutorów (AD)", "Magazyny, wydanie towaru, transport do rolnika, własny asortyment",
  "Marża na każdym zamówieniu z regionu + nowi klienci + sprzedaż innych produktów", "Koszt wdrożenia"),
 ("Rolnik", "Dane o polach (za zgodą), zakupy", "Plan nawożenia, niższy koszt nawożenia, ceny, finansowanie", "Brak"),
]
for k, rw in enumerate(rows):
    for j, t in enumerate(rw):
        c = Wk.cell(4 + k, 1 + j, t); c.alignment = Alignment(wrap_text=True, vertical="top")
Wk.cell(10, 1, "Uzasadnienie podziału 60/40").font = B
Wk.cell(11, 1, "Wartość 5 funkcjonalności GA (koszt odtworzenia)"); Wk.cell(11, 2, f"={N['f5']}")
Wk.cell(12, 1, "CAPEX budowy przez Partnera Technologicznego (koszt, lata 0-2)"); Wk.cell(12, 2, f"={CAPTOT}")
Wk.cell(13, 1, "Bez wkładu GA Partner musiałby zbudować"); Wk.cell(13, 2, "=B11+B12")
Wk.cell(14, 1, "Plus to, czego nie da się kupić: 66 AD, produkty, produkcja, logistyka, marka. Dlatego GA ma większość.")
for r_ in (11, 12, 13): Wk.cell(r_, 2).number_format = NUM
Wk.column_dimensions["A"].width = 38; Wk.column_dimensions["B"].width = 70; Wk.column_dimensions["C"].width = 50; Wk.column_dimensions["D"].width = 36


# ======================= BENCHMARKI =======================
Bm = wb.create_sheet("Benchmarki rynkowe")
Bm["A1"] = "Benchmarki: kto już to zrobił (źródła publiczne, stan na 11.2026)"; Bm["A1"].font = Font(bold=True, size=14)
hdr(Bm, 3, ["Firma / platforma", "Kraj", "Co to jest", "Liczby", "Wniosek dla GAia", "Źródło"])
bm = [
 ("Yara (YaraPlus, Atfarm, FarmGo)", "Norwegia / świat", "Producent nawozów z własnymi aplikacjami dla rolników", "ok. 5 mln rolników korzysta z produktów cyfrowych Yara; YaraPlus wystartował w UK i Irlandii w 01.2026 z John Deere", "Najbliższy odpowiednik GA: producent nawozów buduje bezpośrednią relację z rolnikiem", "https://farming.co.uk/news/yara-uk-launches-new-yaraplus-digital-platform-at-lamma-2026"),
 ("Yara × Agrofy", "Argentyna / Brazylia", "Marketplace agro + fintech; Yara prowadziła rundę", "Runda C 30 mln USD (12.2021), łącznie 66 mln USD; ok. 150 tys. ofert w 20 kategoriach", "Producent nawozów inwestuje w marketplace, bo marketplace sprzedaje jego produkty", "https://agfundernews.com/agrofy-adds-30m-in-yara-led-series-c-round"),
 ("BASF xarvio FIELD MANAGER", "Niemcy / świat", "Aplikacja agronomiczna producenta środków ochrony roślin", ">130 tys. rolników i doradców, >20 mln ha (09.2025); 840 tys. ha w samych Niemczech (01.2023)", "Aplikacja agronomiczna skaluje się do setek tysięcy użytkowników", "https://www.basf.com/global/en/media/news-releases/2025/09/p-25-176"),
 ("Farmers Business Network (FBN)", "USA / Kanada / Australia", "Marketplace środków produkcji + finansowanie", "ok. 117-120 tys. gospodarstw, ok. 187 mln akrów (2025); >1 mld USD finansowania (03.2024)", "Finansowanie zakupów to ważny przychód platformy", "https://en.wikipedia.org/wiki/Farmers_Business_Network"),
 ("DeHaat", "Indie", "Platforma agro: środki produkcji, skup płodów, doradztwo", "Przychód ok. 2 700 cr INR w FY24 (+36-40%), ok. 3 000 cr INR w FY25; strata operacyjna ok. 245 cr INR w FY24", "Duża skala jest możliwa, ale model z własnym magazynem i skupem długo przynosi straty", "https://entrackr.com/fintrackr/dehaat-reports-rs-2675-cr-gross-revenue-in-fy24-losses-shrink-34-7385674"),
 ("Agrando (lekcja)", "Niemcy", "Marketplace środków produkcji bez własnych produktów i logistyki", "ok. 20 tys. rolników (2021), runda A 12 mln EUR; zakończył działalność w 11.2022", "Sam marketplace bez producenta, logistyki i dystrybutorów nie wystarczy. GAia ma wszystkie trzy", "https://agfundernews.com/agrando-raises-14m-in-europes-biggest-farmer-marketplace-series-a"),
 ("eAgronom", "Estonia / Polska", "Aplikacja do zarządzania gospodarstwem", "Firma podaje ok. 20% udziału wśród gospodarstw zbożowych >100 ha w Polsce (deklaracja firmy)", "Polscy rolnicy towarowi już używają aplikacji", "https://blog.eagronom.com/eagronom-won-two-prestigious-agricultural-awards-in-poland"),
 ("Polska: rynek", "Polska", "Kontekst", "1,21 mln rolników w ARiMR, śr. 11,59 ha (2024); ok. 600 tys. żyje z rolnictwa (szac.); brak dominującej platformy zakupowej dla rolników", "Wolne miejsce na rynku. 300 tys. użytkowników ≈ 25% rolników ARiMR", "https://wiescirolnicze.pl/analizy-rynkowe/ile-hektarow-ma-przecietny-polski-rolnik-dane-arimr-za-2024-rok"),
 ("Polska: marketplace'y", "Polska", "Wzrost marketplace'ów ogólnych", "ERLI: GMV 1,23 mld zł (2024) → ok. 1,8 mld zł (2025)", "Argument dla Partnera Marketplacowego: Polacy kupują przez marketplace'y", "https://polandinsight.com/erli-wraps-up-a-record-2025-rapid-growth-and-platform-maturity-confirmed-94308/"),
]
for k, rw in enumerate(bm):
    for j, t in enumerate(rw):
        c = Bm.cell(4 + k, 1 + j, t); c.alignment = Alignment(wrap_text=True, vertical="top")
        if j == 5: c.hyperlink = t; c.font = Font(color="0563C1", underline="single")
for col, w in zip("ABCDEF", [28, 16, 34, 48, 44, 50]): Bm.column_dimensions[col].width = w

wb.save(OUT); print("ok", OUT)
