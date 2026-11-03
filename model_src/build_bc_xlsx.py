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

sec("4. Partner Technologiczny (SoftwareHouse): koszty po jego stronie")
inp("tpbuild", "Budowa platformy na silniku Partnera (rok 0)", 8, "mln zł", "ZAŁOŻENIE. Bez 5 funkcjonalności GA byłoby to ok. 8 + wartość wkładu GA", "0.0")
inp("tpfte", "Zespół rozwoju (programiści, QA, UX, agronomia danych)", [30, 35, 35, 32, 30], "FTE", "ZAŁOŻENIE", "0")
inp("tpftec", "Koszt 1 FTE rozwoju", 0.24, "mln zł/rok", "ok. 20 tys. zł/mies. pełnego kosztu", "0.00")
inp("tpsup", "Utrzymanie 24/7, SRE, wsparcie 2. linii", [3, 5, 7, 9, 10], "FTE", "rośnie z liczbą użytkowników", "0")
inp("tpsupc", "Koszt 1 FTE utrzymania", 0.15, "mln zł/rok", "ZAŁOŻENIE", "0.00")
inp("cloudf", "Chmura: koszt stały", 0.8, "mln zł/rok", "ZAŁOŻENIE", "0.0")
inp("cloudu", "Chmura: koszt na użytkownika", 6, "zł/rok", "ZAŁOŻENIE", "0")
inp("satu", "Dane satelitarne i pogodowe na użytkownika", 4, "zł/rok", "ZAŁOŻENIE (Sentinel bezpłatny, przetwarzanie i pogoda płatne)", "0")
inp("aiu", "Asystent AI / modele agronomiczne na użytkownika", 3, "zł/rok", "ZAŁOŻENIE", "0")
inp("sec", "Bezpieczeństwo, RODO, NIS2, audyty", 0.6, "mln zł/rok", "ZAŁOŻENIE", "0.0")
inp("tpmarg", "Marża Partnera Technologicznego w umowie cost-plus", 0.15, "%", "PROPOZYCJA: Partner nie dokłada do kosztów od 1. dnia", "0%")

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
inp("gaint", "Integracja z SAP / CRM (rok 0)", 3, "mln zł", "ZAŁOŻENIE", "0.0")
inp("gaown", "Utrzymanie 5 funkcjonalności GA i zespół GA", 2, "mln zł/rok", "ZAŁOŻENIE", "0.0")

sec("9. Struktura spółki GAia i finanse")
inp("gash", "Udział GA w spółce GAia", 0.60, "%", "PROPOZYCJA 60/40 wg wkładów (arkusz Wkład stron)", "0%")
inp("wacc", "Stopa dyskontowa", 0.10, "%", "ZAŁOŻENIE", "0%")
inp("f5", "Wartość 5 gotowych funkcjonalności GA (koszt odtworzenia)", 12.5, "mln zł", "ZAŁOŻENIE: 5 × 2,5 mln zł", "0.0")
# naprawa formuły chan
mtr_r = int(N["mtr"].split("$")[-1]); mpl_r = int(N["mpl"].split("$")[-1])
A.cell(chan_row, 2, f"=B{mtr_r}-B{mpl_r}")
A.column_dimensions["A"].width = 66
for c in "BCDEF": A.column_dimensions[c].width = 12
A.column_dimensions["G"].width = 14; A.column_dimensions["H"].width = 80
A.freeze_panes = "A5"

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
U = Sheet("Użytkownicy i GMV", "Użytkownicy i obroty na platformie")
mu = N["mult"]
U.sec("Użytkownicy")
U.row("u", "Użytkownicy zarejestrowani (tys.)", lambda i, c: f"{y('users',i)}*{mu}", total=False, fmt="#,##0")
U.row("nu", "Nowi użytkownicy w roku (tys.)", lambda i, c: f"{c}{U.R['u']}-{'0' if i == 0 else YC[i-1] + str(U.R['u'])}", fmt="#,##0")
U.row("buy", "Aktywni kupujący (tys.)", lambda i, c: f"{c}{U.R['u']}*{N['act']}", total=False, fmt="#,##0")
U.row("pr", "Abonenci premium (tys.)", lambda i, c: f"{c}{U.R['u']}*{y('prem',i)}", total=False, fmt="#,##0.0")
U.sec("Obroty (GMV), mln zł")
U.row("fg", "Nawozy GA kupione przez platformę", lambda i, c: f"{c}{U.R['buy']}*{y('bf',i)}")
U.row("g3", "Inne produkty od sprzedawców 3P (w tym AD)", lambda i, c: f"{c}{U.R['buy']}*{y('b3',i)}")
U.row("gmv", "GMV razem", lambda i, c: f"{c}{U.R['fg']}+{c}{U.R['g3']}", bold=True)
U.row("pen", "Nawozy przez platformę / sprzedaż krajowa GA", lambda i, c: f"{c}{U.R['fg']}/{N['base']}", total=False, fmt="0.0%")
u = lambda k, c: U.ref(k, c)

# ---------- Partner Technologiczny: koszty ----------
T = Sheet("Partner Technologiczny", "Partner Technologiczny (SoftwareHouse): koszty po jego stronie, przychody i wynik")
T.sec("Koszty Partnera Technologicznego")
T.row("build", "Budowa platformy na silniku Partnera", lambda i, c: "0", y0=N["tpbuild"])
T.row("dev", "Zespół rozwoju", lambda i, c: f"{y('tpfte',i)}*{N['tpftec']}")
T.row("sup", "Utrzymanie 24/7 i wsparcie", lambda i, c: f"{y('tpsup',i)}*{N['tpsupc']}")
T.row("cloud", "Chmura (stała + na użytkownika)", lambda i, c: f"{N['cloudf']}+{u('u',c)}*1000*{N['cloudu']}/1000000")
T.row("sat", "Dane satelitarne i pogodowe", lambda i, c: f"{u('u',c)}*1000*{N['satu']}/1000000")
T.row("ai", "Asystent AI / modele agronomiczne", lambda i, c: f"{u('u',c)}*1000*{N['aiu']}/1000000")
T.row("sec", "Bezpieczeństwo i zgodność", lambda i, c: N["sec"])
T.row("cost", "KOSZTY PARTNERA TECHNOLOGICZNEGO RAZEM", lambda i, c: f"SUM({c}{T.R['build']}:{c}{T.R['sec']})", y0=f"SUM(B{T.R['build']}:B{T.R['sec']})", bold=True)
T.row("cpu", "Koszt technologii na użytkownika (zł/rok)", lambda i, c: f"{c}{T.R['cost']}/{u('u',c)}*1000", total=False, fmt="#,##0")
T.sec("Przychody i wynik Partnera Technologicznego")
T.row("fee", "Wynagrodzenie od spółki GAia (koszt + marża)", lambda i, c: f"{c}{T.R['cost']}*(1+{N['tpmarg']})", y0=f"B{T.R['cost']}*(1+{N['tpmarg']})")
T.row("marg", "Zysk z umowy (marża)", lambda i, c: f"{c}{T.R['fee']}-{c}{T.R['cost']}", y0=f"B{T.R['fee']}-B{T.R['cost']}")

# ---------- Partner Marketplacowy ----------
Mp = Sheet("Partner Marketplacowy", "Partner Marketplacowy / Handlowy: ile mu oddajemy, ile go to kosztuje, ile zarabia")

# ---------- JV ----------
J = Sheet("Spółka GAia (JV)", "Spółka GAia (JV): P&L platformy")
J.sec("Przychody")
J.row("rc", "Prowizje od sprzedawców 3P", lambda i, c: f"{u('g3',c)}*{N['take']}")
J.row("ra", "Reklama", lambda i, c: f"{u('g3',c)}*{N['ads']}")
J.row("rs", "Abonamenty premium", lambda i, c: f"{u('pr',c)}*{N['subp']}")
J.row("rf", "Finansowanie / ubezpieczenia", lambda i, c: f"{u('gmv',c)}*{N['fin']}")
J.row("rd", "Dane i analizy dla producentów", lambda i, c: f"{y('data',i)}")
J.row("rg", "Opłata GA od sprzedaży nawozów", lambda i, c: f"{u('fg',c)}*{N['gafee']}")
J.row("rev", "PRZYCHODY RAZEM", lambda i, c: f"SUM({c}{J.R['rc']}:{c}{J.R['rg']})", bold=True)
J.sec("Koszty")
J.row("ctp", "Partner Technologiczny (koszt + marża)", lambda i, c: T.ref("fee", c), y0=T.ref("fee", "B"))
J.row("cmp", "Partner Marketplacowy (udział w prowizji i reklamie)", lambda i, c: f"({c}{J.R['rc']}+{c}{J.R['ra']})*{N['mpshare']}")
J.row("cpay", "Płatności", lambda i, c: f"{u('gmv',c)}*{N['payc']}")
J.row("cmkt", "Marketing: pozyskanie użytkowników + marka", lambda i, c: f"{u('nu',c)}*1000*{N['cac']}/1000000+{N['brand']}", y0=N["launch"])
J.row("ccs", "Obsługa klienta", lambda i, c: f"{u('u',c)}/{N['csu']}*{N['csc']}")
J.row("cmg", "Zarząd i zespół spółki", lambda i, c: N["mgmt"])
J.row("cost", "KOSZTY RAZEM", lambda i, c: f"SUM({c}{J.R['ctp']}:{c}{J.R['cmg']})", y0=f"SUM(B{J.R['ctp']}:B{J.R['cmg']})", bold=True)
J.row("res", "WYNIK SPÓŁKI GAia", lambda i, c: f"{c}{J.R['rev']}-{c}{J.R['cost']}", y0=f"-B{J.R['cost']}", bold=True, fill=TOT)
J.cum("res", "Wynik skumulowany")
J.row("mgn", "Marża spółki (% przychodów)", lambda i, c: f"IF({c}{J.R['rev']}=0,0,{c}{J.R['res']}/{c}{J.R['rev']})", total=False, fmt="0%")

# Partner Marketplacowy - wypełnienie
Mp.sec("Przychody")
Mp.row("rev", "Udział w prowizji i reklamie od sprzedawców 3P", lambda i, c: J.ref("cmp", c))
Mp.sec("Koszty po stronie Partnera")
Mp.row("int", "Integracja", lambda i, c: "0", y0=N["mpint"])
Mp.row("fix", "Zespół i obsługa sprzedawców", lambda i, c: N["mpfix"])
Mp.row("var", "Ruch i wsparcie (zmienny)", lambda i, c: f"{c}{Mp.R['rev']}*{N['mpvar']}")
Mp.row("res", "WYNIK PARTNERA MARKETPLACOWEGO", lambda i, c: f"{c}{Mp.R['rev']}-{c}{Mp.R['int']}-{c}{Mp.R['fix']}-{c}{Mp.R['var']}", y0=f"-B{Mp.R['int']}", bold=True, fill=TOT)
Mp.cum("res", "Wynik skumulowany")
Mp.row("pct", "Jaką część przychodów platformy oddajemy Partnerowi", lambda i, c: f"{c}{Mp.R['rev']}/{J.ref('rev',c)}", total=False, fmt="0%")

# Partner Technologiczny - wynik z udziałem w JV
T.row("jv", "Udział w wyniku spółki GAia", lambda i, c: f"{J.ref('res',c)}*(1-{N['gash']})", y0=f"{J.ref('res','B')}*(1-{N['gash']})")
T.row("res", "WYNIK PARTNERA TECHNOLOGICZNEGO (marża + udział)", lambda i, c: f"{c}{T.R['marg']}+{c}{T.R['jv']}", y0=f"B{T.R['marg']}+B{T.R['jv']}", bold=True, fill=TOT)
T.cum("res", "Wynik skumulowany")

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
F.row("sv", "Oszczędność na nawozach (precyzyjny plan)", lambda i, c: f"{u('fg',c)}*{N['save']}")
F.row("pt", "Przejrzystość cen i oszczędność czasu", lambda i, c: f"{u('gmv',c)}*{N['price']}")
F.row("sub", "Koszt abonamentu premium (dla chętnych)", lambda i, c: f"-{J.ref('rs',c)}")
F.row("res", "KORZYŚĆ ROLNIKÓW RAZEM", lambda i, c: f"SUM({c}{F.R['sv']}:{c}{F.R['sub']})", bold=True, fill=TOT)
F.row("per", "Na aktywnego kupującego (zł/rok)", lambda i, c: f"{c}{F.R['res']}/{u('buy',c)}*1000", total=False, fmt="#,##0")

# ---------- Grupa Azoty ----------
G = Sheet("Grupa Azoty", "Grupa Azoty: wynik z platformy")
G.row("ch", "Lepsza cena netto nawozów (marża AD 7% → 5%)", lambda i, c: f"{u('fg',c)}*'Założenia'!$B${chan_row}")
G.row("mx", "Poprawa mixu", lambda i, c: f"{u('fg',c)}*{N['mix']}")
G.row("sh", "Marża z dodatkowego udziału AN/CAN", lambda i, c: f"{y('pp',i)}*{mu}*{N['ppv']}*{N['cm']}")
G.row("fee", "Opłata dla spółki GAia od nawozów", lambda i, c: f"-{J.ref('rg',c)}")
G.row("own", "Integracja SAP/CRM, utrzymanie 5 funkcjonalności, zespół", lambda i, c: f"-{N['gaown']}", y0=f"-{N['gaint']}")
G.row("jv", "Udział GA w wyniku spółki GAia (60%)", lambda i, c: f"{J.ref('res',c)}*{N['gash']}", y0=f"{J.ref('res','B')}*{N['gash']}")
G.row("res", "WYNIK GRUPY AZOTY", lambda i, c: f"SUM({c}{G.R['ch']}:{c}{G.R['jv']})", y0=f"SUM(B{G.R['ch']}:B{G.R['jv']})", bold=True, fill=TOT)
G.cum("res", "Wynik skumulowany")

# ======================= PODSUMOWANIE =======================
P = wb.create_sheet("Każdy zyskuje", 0)
P["A1"] = "GAia: 300 tys. użytkowników w roku 5. Każdy uczestnik zarabia"; P["A1"].font = Font(bold=True, size=14)
P["A2"] = "Liczby ilustracyjne z arkuszy obliczeniowych. Zmień założenia (żółte pola), wszystko przeliczy się samo."
hdr(P, 4, ["Uczestnik", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma 0-5", "NPV (10%)", "Wynik roczny > 0 od", "Zwrot (suma > 0) od"])
parts = [("Grupa Azoty", G, "res"), ("Partner Technologiczny (SoftwareHouse)", T, "res"),
         ("Partner Marketplacowy / Handlowy", Mp, "res"), ("66 dystrybutorów (AD) razem", D, "res"),
         ("Rolnicy (korzyść)", F, "res"), ("Spółka GAia (JV) - wynik platformy", J, "res")]
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
P.cell(19, 1, "Podział przychodów platformy w roku 5").font = B
splits = [("Partner Technologiczny (koszt + marża)", J.ref("ctp", "G")), ("Partner Marketplacowy", J.ref("cmp", "G")),
          ("Płatności, marketing, obsługa, zespół", f"{J.ref('cost','G')}-{J.ref('ctp','G')}-{J.ref('cmp','G')}"), ("Zysk spółki GAia", J.ref("res", "G"))]
for k, (lab, f) in enumerate(splits):
    pr = 20 + k; P.cell(pr, 1, lab); P.cell(pr, 3, "=" + f).number_format = NUM
    P.cell(pr, 4, f"=C{pr}/{J.ref('rev','G')}").number_format = "0%"
P.column_dimensions["A"].width = 46
for c in "BCDEFGHIJK": P.column_dimensions[c].width = 12

# ======================= WKŁAD STRON =======================
Wk = wb.create_sheet("Wkład stron", 1)
Wk["A1"] = "Kto co wnosi i jak się rozliczamy"; Wk["A1"].font = Font(bold=True, size=14)
hdr(Wk, 3, ["Uczestnik", "Co wnosi", "Jak zarabia", "Ryzyko, które bierze"])
rows = [
 ("Grupa Azoty", "5 gotowych funkcjonalności (do potwierdzenia: plan nawożenia GAia, CRM z profilami gospodarstw, portal AD, Nawozy.eu / doradztwo, koszyk → zamówienie do AD). 66 dystrybutorów. Produkty, produkcja, logistyka, marka, doradcy w regionach.",
  "60% wyniku spółki GAia + lepsza cena netto nawozów + mix + udział w rynku", "Integracja SAP/CRM, 60% kosztów startu spółki"),
 ("Partner Technologiczny (SoftwareHouse)", "Silnik platformy agro, zespół rozwoju, utrzymanie 24/7, chmura, dane satelitarne, AI",
  "Koszt + 15% marży od 1. dnia + 40% wyniku spółki GAia", "40% kosztów startu spółki; ryzyko technologiczne"),
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
Wk.cell(12, 1, "Budowa platformy przez Partnera Technologicznego (rok 0)"); Wk.cell(12, 2, f"={N['tpbuild']}")
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
