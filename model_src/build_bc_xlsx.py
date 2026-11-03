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
inp("users", "Użytkownicy zarejestrowani na koniec roku", [30, 80, 150, 230, 300], "tys.", "CEL ZESPOŁU. 300 tys. ≈ 25% z 1,21 mln rolników w ARiMR (2024)", "#,##0")
inp("mult", "Stres test: mnożnik użytkowników (1 = plan)", 1.0, "x", "Ustaw 0,5 = połowa planu", "0.00")
inp("act", "Aktywni kupujący (kupują min. raz w roku)", 0.35, "% użytk.", "ZAŁOŻENIE", "0%")
inp("prem", "Abonenci premium (podstawa bezpłatna)", [0.02, 0.03, 0.04, 0.05, 0.06], "% użytk.", "ZAŁOŻENIE. Benchmark: Yara i xarvio dają podstawę za darmo, płatne są funkcje zaawansowane", "0%")
inp("subp", "Cena abonamentu premium", 0.24, "tys. zł/rok", "20 zł/mies. Benchmark: SatAgro 10-15 zł/ha/rok; FieldView Prime UK 120 £/rok", "0.00")
inp("bf", "Zakupy nawozów GA przez platformę na kupującego", [6, 7, 8, 8.5, 9], "tys. zł/rok", "ZAŁOŻENIE. Średnie gosp. 11,6 ha (ARiMR 2024); użytkownicy będą więksi", "0.0")
inp("b3", "Zakupy innych produktów (nasiona, ŚOR, części, usługi) na kupującego", [3, 4, 5, 6, 7], "tys. zł/rok", "ZAŁOŻENIE", "0.0")
inp("base", "Sprzedaż krajowa nawozów GA (dla kontroli skali)", "=2.93*0.7*1800", "mln zł/rok", "2,93 mln t × 70% kraj × 1800 zł/t", "#,##0")

sec("2. Przychody spółki GAia (benchmark: Allegro, marketplace'y agro)")
inp("take", "Prowizja od sprzedawców innych produktów (3P)", 0.07, "% GMV 3P", "Benchmark: Allegro 12,0% (PL, IV kw. 2024), Amazon ok. 12%, Vinted ok. 5%. Agro ma niższe marże → 7%", "0.0%")
inp("ads", "Reklama i promowane oferty (retail media)", 0.015, "% GMV 3P", "Benchmark: Allegro Ads ok. 2,0% GMV (IV kw. 2024). Przyjęto ostrożnie 1,5%", "0.0%")
inp("fin", "Finansowanie zakupów (prowizja od banku)", 0.004, "% GMV", "Benchmark: FBN >1 mld USD finansowania zakupów rolników", "0.0%")
inp("data", "Analizy dla producentów nasion / ŚOR (zanonimizowane)", [0.5, 1, 2, 3.5, 5], "mln zł/rok", "ZAŁOŻENIE", "0.0")
inp("gafee", "Prowizja od sprzedaży nawozów GA przez platformę", 0.012, "% GMV naw.", "GA płaci spółce jak każdy sprzedawca, ale niżej (1,2% zamiast 7%), bo to sprzedaż własna", "0.0%")

sec("3. Koszty własne spółki GAia")
inp("payc", "Płatności i obsługa transakcji", 0.006, "% GMV", "ZAŁOŻENIE: operator płatności", "0.0%")
inp("cac", "Koszt pozyskania nowego użytkownika (CAC)", 60, "zł", "ZAŁOŻENIE; niski dzięki sieci sprzedaży GA i doradcom", "0")
inp("brand", "Kampania marki (stała)", 3, "mln zł/rok", "ZAŁOŻENIE", "0.0")
inp("csu", "Obsługa klienta: użytkownicy na 1 konsultanta", 20, "tys.", "ZAŁOŻENIE", "0")
inp("csc", "Koszt konsultanta", 0.12, "mln zł/rok", "ZAŁOŻENIE", "0.00")
inp("mgmt", "Zarząd i zespół spółki GAia", 1.8, "mln zł/rok", "6 osób × 0,3 mln zł", "0.0")
inp("launch", "Kampania startowa (rok 0)", 2, "mln zł", "ZAŁOŻENIE", "0.0")

sec("4. Partner Technologiczny: silnik + warstwa UX (budowa w arkuszu CAPEX)")
inp("fmc", "Koszt 1 osobomiesiąca zespołu (obaj partnerzy)", 0.02, "mln zł", "20 tys. zł/mies. pełnego kosztu", "0.000")
inp("tpfte", "Zespół rozwoju po starcie", [10, 14, 18, 20, 20], "FTE", "ZAŁOŻENIE; w latach 1-2 równolegle trwa budowa", "0")
inp("tpsup", "Utrzymanie 24/7, wsparcie 2. linii", [3, 5, 7, 9, 10], "FTE", "rośnie z liczbą użytkowników", "0")
inp("tpsupc", "Koszt 1 FTE utrzymania", 0.15, "mln zł/rok", "ZAŁOŻENIE", "0.00")
inp("cloudf", "Chmura: koszt stały", 0.8, "mln zł/rok", "ZAŁOŻENIE", "0.0")
inp("cloudu", "Chmura: koszt na użytkownika", 6, "zł/rok", "ZAŁOŻENIE", "0")
inp("satu", "Dane satelitarne i pogodowe na użytkownika", 4, "zł/rok", "ZAŁOŻENIE", "0")
inp("aiu", "Asystent AI na użytkownika", 3, "zł/rok", "ZAŁOŻENIE", "0")
inp("sec", "Bezpieczeństwo, RODO, NIS2, audyty", 0.6, "mln zł/rok", "ZAŁOŻENIE", "0.0")

sec("5. Partner Marketplace / Handel: użyczony marketplace (dostosowanie w arkuszu CAPEX)")
inp("mpfte", "Zespół operacji marketplace'u (katalog, sprzedawcy, spory)", [4, 5, 6, 7, 8], "FTE", "ZAŁOŻENIE", "0")
inp("mpftec", "Koszt 1 FTE operacji", 0.18, "mln zł/rok", "ZAŁOŻENIE", "0.00")
inp("mpinf", "Udział w infrastrukturze marketplace'u (hosting, licencje)", 0.4, "mln zł/rok", "ZAŁOŻENIE", "0.0")
inp("mpu", "Infrastruktura na użytkownika", 2, "zł/rok", "ZAŁOŻENIE", "0")

sec("6. Zasady rozliczeń (benchmark JV: wkład za udziały, usługi po koszcie + marża)")
inp("tpmarg", "Marża na usługach partnerów (koszt + marża)", 0.15, "%", "Benchmark JV: usługi partnerów dla spółki po koszcie z marżą rynkową", "0%")
inp("gash", "Udział Grupy Azoty w spółce GAia", 0.51, "%", "PROPOZYCJA: większość dla wnoszącego produkty i klientów (kontrola)", "0%")
inp("tpsh", "Udział Partnera Technologicznego", 0.245, "%", "PROPOZYCJA: wkład = silnik + UX (licencja)", "0.0%")
inp("mpsh", "Udział Partnera Marketplace / Handel", "=1-B{g}-B{t}", "%", "Formuła: 100% − GA − Partner Technologiczny", "0.0%")
inp("lictp", "Licencja za użyczenie silnika Partnera Technologicznego", 3, "zł/użytk./rok", "Benchmark JV: wnoszone IP rozliczane licencją (WIPO, Ankura)", "0")
inp("licmp", "Licencja za użyczenie marketplace'u Partnera Marketplace", 0.005, "% GMV 3P", "Benchmark: marketplace w SaaS (np. Mirakl) 5-7% GMV; partner-wspólnik 0,5%", "0.0%")
inp("wacc", "Stopa dyskontowa", 0.10, "%", "ZAŁOŻENIE", "0%")
inp("f5", "Wartość wkładu GA: 5 gotowych funkcjonalności (koszt odtworzenia)", 12.5, "mln zł", "ZAŁOŻENIE: 5 × 2,5 mln zł", "0.0")
inp("vtp", "Wartość wkładu Partnera Technologicznego: silnik (koszt odtworzenia)", 10, "mln zł", "ZAŁOŻENIE do potwierdzenia wyceną", "0.0")
inp("vmp", "Wartość wkładu Partnera Marketplace: gotowy marketplace (koszt odtworzenia)", 10, "mln zł", "ZAŁOŻENIE do potwierdzenia wyceną", "0.0")

sec("7. Grupa Azoty")
inp("chan", "Niższy koszt kanału na zamówieniach z platformy", 0.02, "% GMV naw.", "ZAŁOŻENIE: krótsza ścieżka sprzedaży, mniej upustów", "0.0%")
inp("mix", "Poprawa mixu (NPK, specjalistyczne)", 0.01, "% GMV naw.", "ZAŁOŻENIE", "0.0%")
inp("pp", "Dodatkowy udział w rynku AN/CAN", [0.3, 0.8, 1.5, 2.2, 3.0], "p.p.", "ZAŁOŻENIE, cel strategii 45% → 60%", "0.0")
inp("ppv", "Przychód z 1 p.p. udziału AN/CAN", 39, "mln zł/rok", "GA_5_luk, arkusz Szacunki", "0")
inp("cm", "Marża kontrybucyjna na dodatkowej tonie", 0.25, "%", "ZAŁOŻENIE - Kontroling", "0%")
inp("gaint0", "Integracja z SAP / CRM po stronie GA: rok 0", 2, "mln zł", "ZAŁOŻENIE", "0.0")
inp("gaint1", "Integracja z SAP / CRM po stronie GA: rok 1", 1, "mln zł", "ZAŁOŻENIE", "0.0")
inp("gaown", "Utrzymanie 5 funkcjonalności GA i zespół GA", 2, "mln zł/rok", "ZAŁOŻENIE", "0.0")

sec("8. Moc platformy (przeliczenie GMV na zamówienia i tony)")
inp("aovf", "Średnie zamówienie nawozów", 4.5, "tys. zł", "ok. 2,5 t × 1800 zł/t", "0.0")
inp("aov3", "Średnie zamówienie innych produktów", 1.2, "tys. zł", "ZAŁOŻENIE", "0.0")
inp("pt", "Cena nawozu do przeliczenia na tony", 1800, "zł/t", "", "#,##0")
inp("nad", "Punkty wydania nawozów (sieć sprzedaży GA)", 66, "szt.", "Zamówienia realizuje sieć sprzedaży GA", "0")
inp("seas", "Część zamówień nawozów w szczycie sezonu", 0.6, "%", "wiosna", "0%")
inp("seasd", "Dni szczytu sezonu", 90, "dni", "", "0")
inp("sell", "Sprzedawcy innych produktów na platformie", [80, 130, 190, 250, 320], "szt.", "sprzedawcy z marketplace'u Partnera + producenci nasion, ŚOR, części", "0")
g_r = int(N["gash"].split("$")[-1]); t_r = int(N["tpsh"].split("$")[-1])
A.cell(int(N["mpsh"].split("$")[-1]), 2, f"=1-B{g_r}-B{t_r}")
A.column_dimensions["A"].width = 66
for c in "BCDEF": A.column_dimensions[c].width = 12
A.column_dimensions["G"].width = 14; A.column_dimensions["H"].width = 80
A.freeze_panes = "A5"

# ======================= CAPEX =======================
C = wb.create_sheet("CAPEX budowa")
C["A1"] = "CAPEX: co budujemy, kto buduje, za ile i kiedy (koszt partnerów, bez marży)"
C["A2"] = "Marketplace NIE jest budowany od zera: Partner Marketplace użycza gotowy (katalog, koszyk, płatności, panel sprzedawcy). Budujemy silnik agro + UX i dostosowanie marketplace'u."
hdr(C, 4, ["Moduł", "Kto buduje", "Co obejmuje", "Osobo-\nmiesiące", "% rok 0", "% rok 1", "% rok 2", "Koszt\nmln zł", "Rok 0", "Rok 1", "Rok 2"])
PTn, MPn = "Partner Technologiczny", "Partner Marketplace"
mods = [
 ("Architektura, chmura, DevOps", PTn, "środowiska, CI/CD, monitoring", 18, 1, 0, 0),
 ("Silnik agro + 5 funkcjonalności GA", PTn, "adaptacja silnika, plan nawożenia, CRM, Nawozy.eu, koszyk z planu", 60, .7, .3, 0),
 ("Warstwa UX: web i aplikacja mobilna", PTn, "projekt i budowa ekranów dla rolnika, iOS / Android", 72, .5, .5, 0),
 ("Integracja silnika z marketplace'em Partnera", PTn, "API: produkty, ceny, koszyk, zamówienia, konto rolnika", 30, .8, .2, 0),
 ("Integracja z SAP / CRM GA", PTn, "zamówienia, ceny, dostępność, dane klientów", 24, .7, .3, 0),
 ("mObywatel / login.gov.pl, dane ARiMR", PTn, "logowanie rolnika, działki i uprawy", 18, 0, 1, 0),
 ("Mapy pól i dane satelitarne", PTn, "NDVI, strefy, plan nawożenia premium", 36, 0, .6, .4),
 ("Asystent AI agronomiczny", PTn, "pytania rolnika, rekomendacje", 24, 0, .3, .7),
 ("Abonament premium", PTn, "płatności cykliczne, uprawnienia", 12, 0, .5, .5),
 ("Bezpieczeństwo i testy", PTn, "testy penetracyjne, RODO, NIS2, QA, testy z rolnikami", 42, .4, .4, .2),
 ("Kierowanie projektem, analityka", PTn, "PM, analityka biznesowa", 30, .4, .35, .25),
 ("Kategorie agro i karty produktów", MPn, "atrybuty: skład, dawki, rejestracja ŚOR, opakowania", 18, .8, .2, 0),
 ("Zamówienia nawozów do punktów wydania GA", MPn, "routing do punktu wydania, odbiór osobisty, dostawa na gospodarstwo", 30, .7, .3, 0),
 ("Płatności, faktury, kredyt kupiecki w koszyku", MPn, "dostosowanie płatności, integracja z bankiem", 18, .5, .5, 0),
 ("Onboarding sprzedawców agro", MPn, "producenci nasion, ŚOR, części; panel sprzedawcy", 12, .5, .5, 0),
 ("Reklama i promowane oferty", MPn, "konfiguracja retail media dla kategorii agro", 12, 0, .6, .4),
 ("Testy, migracja katalogu", MPn, "QA, import katalogu", 12, .6, .4, 0),
]
cr = 5
for name, who, what, fm, a0, a1, a2 in mods:
    C.cell(cr, 1, name); C.cell(cr, 2, who); C.cell(cr, 3, what)
    for j, v in enumerate([fm, a0, a1, a2]):
        c = C.cell(cr, 4 + j, v); c.font = BLUE; c.fill = INF; c.number_format = "0" if j == 0 else "0%"
    C.cell(cr, 8, f"=D{cr}*{N['fmc']}")
    for j, col in enumerate("EFG"): C.cell(cr, 9 + j, f"=$H{cr}*{col}{cr}")
    for j in range(8, 12): C.cell(cr, j).number_format = NUM
    cr += 1
first, last = 5, cr - 1
CAPROW = {}
for who in (PTn, MPn):
    C.cell(cr, 1, f"Razem {who}").font = B; C.cell(cr, 2, who)
    C.cell(cr, 4, f'=SUMIF($B${first}:$B${last},$B{cr},D${first}:D${last})').font = B
    for col in "HIJK": C[f"{col}{cr}"] = f'=SUMIF($B${first}:$B${last},$B{cr},{col}${first}:{col}${last})'; C[f"{col}{cr}"].font = B; C[f"{col}{cr}"].number_format = NUM
    CAPROW[who] = cr; cr += 1
C.cell(cr, 1, "Koszty zewnętrzne Partnera Technologicznego: licencje, testy, sklepy z aplikacjami"); C.cell(cr, 2, PTn)
for j, v in enumerate([0.4, 0.3, 0.1]):
    c = C.cell(cr, 9 + j, v); c.font = BLUE; c.fill = INF; c.number_format = NUM
C[f"H{cr}"] = f"=SUM(I{cr}:K{cr})"; C[f"H{cr}"].number_format = NUM
ext_r = cr; cr += 1
C.cell(cr, 1, "CAPEX RAZEM (koszt partnerów)").font = B
C.cell(cr, 4, f"=D{CAPROW[PTn]}+D{CAPROW[MPn]}").font = B
for col in "HIJK": C[f"{col}{cr}"] = f"={col}{CAPROW[PTn]}+{col}{CAPROW[MPn]}+{col}{ext_r}"; C[f"{col}{cr}"].font = B; C[f"{col}{cr}"].fill = TOT; C[f"{col}{cr}"].number_format = NUM
tot_r = cr; cr += 1
C.cell(cr, 1, "CAPEX PŁACI SPÓŁKA GAia (koszt + 15%)").font = B
for col in "HIJK": C[f"{col}{cr}"] = f"={col}{tot_r}*(1+{N['tpmarg']})"; C[f"{col}{cr}"].font = B; C[f"{col}{cr}"].fill = TOT; C[f"{col}{cr}"].number_format = NUM
capj_r = cr; cr += 2
C.cell(cr, 1, "Czego nie budujemy, bo Partner Marketplace użycza gotowy marketplace:").font = B; cr += 1
for t in ["katalog produktów i wyszukiwarka, koszyk i checkout, panel sprzedawcy, rozliczenia ze sprzedawcami, obsługa zwrotów i reklamacji, system reklamowy",
          "szacunkowo ok. 150 osobomiesięcy (ok. 3 mln zł) budowy i lata rozwoju - wartość wkładu Partnera Marketplace w Założeniach, sekcja 6"]:
    C.cell(cr, 1, "• " + t); cr += 1
C.column_dimensions["A"].width = 46; C.column_dimensions["B"].width = 22; C.column_dimensions["C"].width = 58
for col in "DEFGHIJK": C.column_dimensions[col].width = 10
C.row_dimensions[4].height = 32
def capref(who, yi):  # koszt budowy danego partnera w roku yi (0-2)
    col = "IJK"[yi]
    s_ = f"'CAPEX budowa'!${col}${CAPROW[who]}"
    if who == PTn: s_ += f"+'CAPEX budowa'!${col}${ext_r}"
    return s_
CAPJ = {i: f"'CAPEX budowa'!${'IJK'[i]}${capj_r}" for i in range(3)}
CAPTOT = f"'CAPEX budowa'!$H${tot_r}"; CAPPT = f"('CAPEX budowa'!$H${CAPROW[PTn]}+'CAPEX budowa'!$H${ext_r})"; CAPMP = f"'CAPEX budowa'!$H${CAPROW[MPn]}"
CAPFM = f"'CAPEX budowa'!$D${tot_r}"
# ======================= silnik arkuszy =======================
class Sheet:
    def __init__(self, title, head):
        self.ws = wb.create_sheet(title); self.r = 5; self.R = {}
        self.ws["A1"] = head; self.ws["A1"].font = Font(bold=True, size=14)
        hdr(self.ws, 4, ["Pozycja"] + YRS + ["Suma 0-5", "Jak liczone / skąd"]); self.notes = {}
        for i, col in enumerate(YC): self.ws[f"{col}3"] = i + 1; self.ws[f"{col}3"].font = Font(color="FFFFFF")
        self.ws.column_dimensions["A"].width = 60
        for c in "BCDEFGH": self.ws.column_dimensions[c].width = 12
        self.ws.column_dimensions["I"].width = 78
        self.ws.freeze_panes = "B5"
    def row(self, key, label, f, y0=None, bold=False, fill=None, total=True, fmt=NUM, note=None):
        ws = self.ws; ws.cell(self.r, 1, label)
        if note: self.notes[key] = note
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
        src = self.R[key]; rr = self.row(key + "_cum", label, lambda i, c: "0", y0="0", total=False, note="Suma narastająca wiersza powyżej: poprzedni rok + bieżący rok")
        self.ws[f"B{rr}"] = f"=B{src}"
        for i, col in enumerate(YC):
            self.ws[f"{col}{rr}"] = f"={'B' if i == 0 else YC[i-1]}{rr}+{col}{src}"


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
U.row("g3", "GMV innych produktów od sprzedawców (3P)", lambda i, c: f"{c}{U.R['buy']}*{c}{U.R['b3']}")
U.row("gmv", "GMV RAZEM", lambda i, c: f"{c}{U.R['fg']}+{c}{U.R['g3']}", bold=True, fill=TOT)
U.row("pen", "Nawozy przez platformę / sprzedaż krajowa GA", lambda i, c: f"{c}{U.R['fg']}/{N['base']}", total=False, fmt="0.0%")
U.sec("Zamówienia i tony")
U.row("of", "Zamówienia nawozów (tys.)", lambda i, c: f"{c}{U.R['fg']}/{N['aovf']}", fmt="#,##0.0")
U.row("o3", "Zamówienia innych produktów (tys.)", lambda i, c: f"{c}{U.R['g3']}/{N['aov3']}", fmt="#,##0.0")
U.row("ot", "Zamówienia razem (tys.)", lambda i, c: f"{c}{U.R['of']}+{c}{U.R['o3']}", bold=True, fmt="#,##0.0")
U.row("tn", "Nawozy przez platformę (tys. t)", lambda i, c: f"{c}{U.R['fg']}*1000/{N['pt']}", fmt="#,##0")
U.row("oad", "Zamówienia nawozów na 1 punkt wydania dziennie w szczycie", lambda i, c: f"{c}{U.R['of']}*1000*{N['seas']}/{N['nad']}/{N['seasd']}", total=False, fmt="#,##0.0")
U.row("sel", "Sprzedawcy innych produktów", lambda i, c: f"{y('sell',i)}", total=False, fmt="#,##0")
u = lambda k, c: U.ref(k, c)
COLS = ["B"] + YC
capT = lambda i: capref(PTn, i) if i <= 2 else "0"
capM = lambda i: capref(MPn, i) if i <= 2 else "0"

# ---------- Partner Technologiczny ----------
T = Sheet("Partner Technologiczny", "Partner Technologiczny (silnik + UX): koszty po jego stronie, z czego i kiedy zarabia")
T.sec("A. Koszty: budowa (CAPEX, arkusz CAPEX budowa)")
T.row("cap", "Budowa: silnik agro, UX, integracje (koszt)", lambda i, c: capT(i + 1), y0=capT(0))
T.sec("B. Koszty: utrzymanie i rozwój po starcie (OPEX)")
T.row("dev", "Zespół rozwoju", lambda i, c: f"{y('tpfte',i)}*12*{N['fmc']}")
T.row("sup", "Utrzymanie 24/7 i wsparcie", lambda i, c: f"{y('tpsup',i)}*{N['tpsupc']}")
T.row("cloud", "Chmura (stała + na użytkownika)", lambda i, c: f"{N['cloudf']}+{u('u',c)}*1000*{N['cloudu']}/1000000")
T.row("sat", "Dane satelitarne i pogodowe", lambda i, c: f"{u('u',c)}*1000*{N['satu']}/1000000")
T.row("ai", "Asystent AI", lambda i, c: f"{u('u',c)}*1000*{N['aiu']}/1000000")
T.row("sec", "Bezpieczeństwo i zgodność", lambda i, c: N["sec"])
T.row("opex", "OPEX Partnera razem", lambda i, c: f"SUM({c}{T.R['dev']}:{c}{T.R['sec']})", bold=True)
T.row("cost", "KOSZTY PARTNERA RAZEM (CAPEX + OPEX)", lambda i, c: f"{c}{T.R['cap']}+{c}{T.R['opex']}", y0=f"B{T.R['cap']}", bold=True, fill=TOT)
T.row("cpu", "OPEX na użytkownika (zł/rok)", lambda i, c: f"{c}{T.R['opex']}/{u('u',c)}*1000", total=False, fmt="#,##0")
T.sec("C. Spółka GAia płaci Partnerowi (koszt + 15%)")
T.row("fcap", "Za budowę", lambda i, c: f"{c}{T.R['cap']}*(1+{N['tpmarg']})", y0=f"B{T.R['cap']}*(1+{N['tpmarg']})")
T.row("fopx", "Za utrzymanie i rozwój", lambda i, c: f"{c}{T.R['opex']}*(1+{N['tpmarg']})")
T.row("fee", "Razem od spółki GAia", lambda i, c: f"{c}{T.R['fcap']}+{c}{T.R['fopx']}", y0=f"B{T.R['fcap']}", bold=True)
T.sec("D. Z czego i kiedy zarabia Partner Technologiczny")
T.row("mcap", "1. Marża 15% na budowie (lata 0-2)", lambda i, c: f"{c}{T.R['fcap']}-{c}{T.R['cap']}", y0=f"B{T.R['fcap']}-B{T.R['cap']}")
T.row("mopx", "2. Marża 15% na utrzymaniu (lata 1-5)", lambda i, c: f"{c}{T.R['fopx']}-{c}{T.R['opex']}")

# ---------- Partner Marketplace ----------
Mp = Sheet("Partner Marketplace", "Partner Marketplace / Handel (użyczony marketplace): koszty po jego stronie, z czego i kiedy zarabia")
Mp.sec("A. Koszty: dostosowanie marketplace'u (CAPEX, arkusz CAPEX budowa)")
Mp.row("cap", "Dostosowanie: kategorie agro, zamówienia do punktów wydania, płatności", lambda i, c: capM(i + 1), y0=capM(0))
Mp.sec("B. Koszty: operacje marketplace'u (OPEX)")
Mp.row("team", "Zespół operacji (katalog, sprzedawcy, spory)", lambda i, c: f"{y('mpfte',i)}*{N['mpftec']}")
Mp.row("inf", "Infrastruktura marketplace'u (stała + na użytkownika)", lambda i, c: f"{N['mpinf']}+{u('u',c)}*1000*{N['mpu']}/1000000")
Mp.row("opex", "OPEX Partnera razem", lambda i, c: f"{c}{Mp.R['team']}+{c}{Mp.R['inf']}", bold=True)
Mp.row("cost", "KOSZTY PARTNERA RAZEM (CAPEX + OPEX)", lambda i, c: f"{c}{Mp.R['cap']}+{c}{Mp.R['opex']}", y0=f"B{Mp.R['cap']}", bold=True, fill=TOT)
Mp.sec("C. Spółka GAia płaci Partnerowi (koszt + 15%)")
Mp.row("fcap", "Za dostosowanie", lambda i, c: f"{c}{Mp.R['cap']}*(1+{N['tpmarg']})", y0=f"B{Mp.R['cap']}*(1+{N['tpmarg']})")
Mp.row("fopx", "Za operacje", lambda i, c: f"{c}{Mp.R['opex']}*(1+{N['tpmarg']})")
Mp.row("fee", "Razem od spółki GAia", lambda i, c: f"{c}{Mp.R['fcap']}+{c}{Mp.R['fopx']}", y0=f"B{Mp.R['fcap']}", bold=True)
Mp.sec("D. Z czego i kiedy zarabia Partner Marketplace")
Mp.row("mcap", "1. Marża 15% na dostosowaniu (lata 0-2)", lambda i, c: f"{c}{Mp.R['fcap']}-{c}{Mp.R['cap']}", y0=f"B{Mp.R['fcap']}-B{Mp.R['cap']}")
Mp.row("mopx", "2. Marża 15% na operacjach (lata 1-5)", lambda i, c: f"{c}{Mp.R['fopx']}-{c}{Mp.R['opex']}")

# ---------- Spółka GAia ----------
J = Sheet("Spółka GAia (JV)", "Spółka GAia (JV): przychody, OPEX, EBITDA, CAPEX, amortyzacja, przepływy pieniężne")
J.sec("Przychody (GMV × stawka z benchmarku)")
J.row("rc", "Prowizje: GMV 3P × 7%", lambda i, c: f"{u('g3',c)}*{N['take']}")
J.row("ra", "Reklama: GMV 3P × 1,5%", lambda i, c: f"{u('g3',c)}*{N['ads']}")
J.row("rg", "Prowizja od nawozów GA: GMV nawozów × 1,2%", lambda i, c: f"{u('fg',c)}*{N['gafee']}")
J.row("rf", "Finansowanie: GMV × 0,4%", lambda i, c: f"{u('gmv',c)}*{N['fin']}")
J.row("rs", "Abonament premium: abonenci × 240 zł", lambda i, c: f"{u('pr',c)}*{N['subp']}")
J.row("rd", "Analizy dla producentów", lambda i, c: f"{y('data',i)}")
J.row("rev", "PRZYCHODY RAZEM", lambda i, c: f"SUM({c}{J.R['rc']}:{c}{J.R['rd']})", bold=True, fill=TOT)
J.sec("OPEX")
J.row("ctp", "Partner Technologiczny: utrzymanie i rozwój (koszt + 15%)", lambda i, c: T.ref("fopx", c))
J.row("cmp", "Partner Marketplace: operacje (koszt + 15%)", lambda i, c: Mp.ref("fopx", c))
J.row("ltp", "Licencja: silnik Partnera Technologicznego (3 zł / użytkownik)", lambda i, c: f"{u('u',c)}*1000*{N['lictp']}/1000000")
J.row("lmp", "Licencja: marketplace Partnera Marketplace (0,5% GMV 3P)", lambda i, c: f"{u('g3',c)}*{N['licmp']}")
J.row("cpay", "Płatności: GMV × 0,6%", lambda i, c: f"{u('gmv',c)}*{N['payc']}")
J.row("cmkt", "Marketing: nowi użytkownicy × 60 zł + marka 3 mln zł", lambda i, c: f"{u('nu',c)}*1000*{N['cac']}/1000000+{N['brand']}", y0=N["launch"])
J.row("ccs", "Obsługa klienta: 1 konsultant na 20 tys. użytkowników", lambda i, c: f"{u('u',c)}/{N['csu']}*{N['csc']}")
J.row("cmg", "Zarząd i zespół spółki (6 osób)", lambda i, c: N["mgmt"], y0=f"{N['mgmt']}*0.5")
J.row("opex", "OPEX RAZEM", lambda i, c: f"SUM({c}{J.R['ctp']}:{c}{J.R['cmg']})", y0=f"SUM(B{J.R['ctp']}:B{J.R['cmg']})", bold=True)
J.row("ebitda", "EBITDA", lambda i, c: f"{c}{J.R['rev']}-{c}{J.R['opex']}", y0=f"-B{J.R['opex']}", bold=True, fill=TOT)
J.row("emg", "Marża EBITDA", lambda i, c: f"IF({c}{J.R['rev']}=0,0,{c}{J.R['ebitda']}/{c}{J.R['rev']})", total=False, fmt="0%")
J.sec("CAPEX i amortyzacja")
J.row("capex", "CAPEX: budowa i dostosowanie (obaj partnerzy, koszt + 15%)", lambda i, c: CAPJ[i + 1] if i + 1 in CAPJ else "0", y0=CAPJ[0])
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
for k_, (lab, sh_) in enumerate([("Wkład Grupy Azoty (51%)", N['gash']), ("Wkład Partnera Technologicznego (24,5%)", N['tpsh']), ("Wkład Partnera Marketplace (24,5%)", N['mpsh'])]):
    J.ws.cell(fr + 1 + k_, 1, lab); J.ws.cell(fr + 1 + k_, 2, f"=B{fr}*{sh_}").number_format = NUM
FUND = f"'Spółka GAia (JV)'!$B${fr}"; FUNDGA = f"'Spółka GAia (JV)'!$B${fr+1}"; FUNDTP = f"'Spółka GAia (JV)'!$B${fr+2}"; FUNDMP = f"'Spółka GAia (JV)'!$B${fr+3}"

for P_, shn, lk, llab in ((T, 'tpsh', 'ltp', "3. Licencja za silnik (3 zł / użytkownik / rok)"), (Mp, 'mpsh', 'lmp', "3. Licencja za marketplace (0,5% GMV 3P)")):
    P_.row("lic", llab, lambda i, c, lk=lk: J.ref(lk, c))
    P_.row("jv", "4. Udział 24,5% w gotówce spółki (dopłata na start, potem dywidenda)", lambda i, c, s=shn: f"{J.ref('fcf',c)}*{N[s]}", y0=f"{J.ref('fcf','B')}*{N[shn]}")
    P_.row("res", "WYNIK PARTNERA (1 + 2 + 3 + 4)", lambda i, c, P_=P_: f"{c}{P_.R['mcap']}+{c}{P_.R['mopx']}+{c}{P_.R['lic']}+{c}{P_.R['jv']}", y0=f"B{P_.R['mcap']}+B{P_.R['mopx']}+B{P_.R['jv']}", bold=True, fill=TOT)
    P_.cum("res", "Wynik skumulowany")

# ---------- Grupa Azoty ----------
G = Sheet("Grupa Azoty", "Grupa Azoty: wynik z platformy")
G.row("sh", "Udział w rynku AN/CAN: p.p. × 39 mln zł × 25% marży", lambda i, c: f"{y('pp',i)}*{mu}*{N['ppv']}*{N['cm']}")
G.row("ch", "Niższy koszt kanału: GMV nawozów × 2%", lambda i, c: f"{u('fg',c)}*{N['chan']}")
G.row("mx", "Poprawa mixu: GMV nawozów × 1%", lambda i, c: f"{u('fg',c)}*{N['mix']}")
G.row("fee", "Prowizja dla spółki GAia: GMV nawozów × 1,2%", lambda i, c: f"-{J.ref('rg',c)}")
G.row("own", "Integracja SAP/CRM (CAPEX GA) i utrzymanie 5 funkcjonalności", lambda i, c: f"-{N['gaown']}" + (f"-{N['gaint1']}" if i == 0 else ""), y0=f"-{N['gaint0']}")
G.row("jv", "51% gotówki spółki GAia (dopłata na start, potem dywidenda)", lambda i, c: f"{J.ref('fcf',c)}*{N['gash']}", y0=f"{J.ref('fcf','B')}*{N['gash']}")
G.row("res", "WYNIK GRUPY AZOTY", lambda i, c: f"SUM({c}{G.R['sh']}:{c}{G.R['jv']})", y0=f"SUM(B{G.R['sh']}:B{G.R['jv']})", bold=True, fill=TOT)
G.cum("res", "Wynik skumulowany")

# ---------- Moc platformy ----------
Mc = Sheet("Moc platformy", "Moc platformy w liczbach")
Mc.row("u", "Użytkownicy (tys.)", lambda i, c: u("u", c), total=False, fmt="#,##0")
Mc.row("b", "Aktywni kupujący (tys.)", lambda i, c: u("buy", c), total=False, fmt="#,##0")
Mc.row("p", "Abonenci premium (tys.)", lambda i, c: u("pr", c), total=False, fmt="#,##0.0")
Mc.row("s", "Sprzedawcy innych produktów", lambda i, c: u("sel", c), total=False, fmt="#,##0")
Mc.row("g", "GMV razem (mln zł)", lambda i, c: u("gmv", c))
Mc.row("o", "Zamówienia razem (tys.)", lambda i, c: u("ot", c), fmt="#,##0")
Mc.row("t", "Nawozy przez platformę (tys. t)", lambda i, c: u("tn", c), fmt="#,##0")
Mc.row("ad", "Zamówienia nawozów na 1 punkt wydania dziennie w szczycie", lambda i, c: u("oad", c), total=False, fmt="#,##0.0")
Mc.row("r", "Przychody spółki GAia (mln zł)", lambda i, c: J.ref("rev", c))
Mc.row("tr", "Przychód spółki / GMV (efektywna prowizja)", lambda i, c: f"{J.ref('rev',c)}/{u('gmv',c)}", total=False, fmt="0.0%")
Mc.row("ru", "Przychód na użytkownika (zł/rok)", lambda i, c: f"{J.ref('rev',c)}/{u('u',c)}*1000", total=False, fmt="#,##0")
Mc.row("cu", "OPEX na użytkownika (zł/rok)", lambda i, c: f"{J.ref('opex',c)}/{u('u',c)}*1000", total=False, fmt="#,##0")
Mc.row("eu", "EBITDA na użytkownika (zł/rok)", lambda i, c: f"{J.ref('ebitda',c)}/{u('u',c)}*1000", total=False, fmt="#,##0")

# ======================= PODSUMOWANIE =======================
P = wb.create_sheet("Każdy partner zyskuje", 0)
P["A1"] = "GAia: joint venture trzech partnerów. Każdy partner zarabia"
P["A2"] = "Przepływy pieniężne (gotówka), mln zł. Partnerzy zarabiają na usługach dla spółki (koszt + 15%) i na udziale w gotówce spółki."
hdr(P, 4, ["Partner", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma 0-5", "NPV (10%)", "Wynik roczny > 0 od", "Zwrot (suma > 0) od"])
parts = [("Grupa Azoty (51%)", G, "res"), ("Partner Technologiczny (24,5%)", T, "res"), ("Partner Marketplace / Handel (24,5%)", Mp, "res"), ("Spółka GAia (JV) - przepływy pieniężne", J, "fcf")]
for k, (name, sh, key) in enumerate(parts):
    pr = 5 + k; P.cell(pr, 1, name)
    for j, col in enumerate(["B"] + YC): P.cell(pr, 2 + j, f"={sh.ref(key, col)}")
    P.cell(pr, 8, f"=SUM(B{pr}:G{pr})")
    P.cell(pr, 9, f"=B{pr}+NPV({N['wacc']},C{pr}:G{pr})")
    P.cell(pr, 10, f'=IF(COUNTIF(C{pr}:G{pr},">0")=0,"po roku 5","rok "&SUMPRODUCT(MIN(IF(C{pr}:G{pr}>0,{{1,2,3,4,5}},99))))')
    P.cell(pr, 11, f'=IF(B{pr}+C{pr}>0,"rok 1",IF(SUM(B{pr}:D{pr})>0,"rok 2",IF(SUM(B{pr}:E{pr})>0,"rok 3",IF(SUM(B{pr}:F{pr})>0,"rok 4",IF(SUM(B{pr}:G{pr})>0,"rok 5","po roku 5")))))')
    for c in range(2, 10): P.cell(pr, c).number_format = NUM
    if k == 0:
        for c in range(1, 12): P.cell(pr, c).font = B; P.cell(pr, c).fill = TOT
SK = 10
kpis = [("Użytkownicy (tys.)", "u", "#,##0"), ("Aktywni kupujący (tys.)", "buy", "#,##0"), ("GMV razem (mln zł)", "gmv", NUM),
        ("Nawozy GA przez platformę (mln zł)", "fg", NUM), ("Udział platformy w sprzedaży krajowej GA", "pen", "0.0%")]
P.cell(SK, 1, "Skala")
for j, h in enumerate(["Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5"]): P.cell(SK, 3 + j, h)
for k, (lab, key, fm) in enumerate(kpis):
    pr = SK + 1 + k; P.cell(pr, 1, lab)
    for j, col in enumerate(YC): P.cell(pr, 3 + j, f"={U.ref(key, col)}").number_format = fm
IV = SK + 7
P.cell(IV, 1, "Inwestycja i finansowanie"); P.cell(IV, 3, "mln zł")
inv = [("CAPEX Partnera Technologicznego (koszt, lata 0-2)", CAPPT), ("CAPEX Partnera Marketplace (koszt, lata 0-2)", CAPMP),
       ("CAPEX GA: integracja SAP/CRM (lata 0-1)", f"{N['gaint0']}+{N['gaint1']}"), ("Maks. zapotrzebowanie spółki GAia na finansowanie", FUND),
       ("w tym Grupa Azoty (51%)", FUNDGA), ("w tym Partner Technologiczny (24,5%)", FUNDTP), ("w tym Partner Marketplace (24,5%)", FUNDMP)]
for k, (lab, f) in enumerate(inv):
    pr = IV + 1 + k; P.cell(pr, 1, lab); P.cell(pr, 3, "=" + f).number_format = NUM
P.column_dimensions["A"].width = 56
for c in "BCDEFGHIJK": P.column_dimensions[c].width = 12
REG_P = [(4, 4 + len(parts), 1, 11), (SK, SK + len(kpis), 1, 7), (IV, IV + len(inv), 1, 3)]
STRESS_ROW = IV + len(inv) + 2

# ======================= WKŁAD STRON =======================
Wk = wb.create_sheet("Wkład partnerów", 1)
Wk["A1"] = "Kto co wnosi do spółki GAia i jak się rozliczamy"
Wk["A2"] = "Zasada z najlepszych praktyk JV: wkład rzeczowy (licencja) za udziały, bieżące usługi po koszcie + marża, zysk według udziałów."
hdr(Wk, 3, ["Partner", "Udział", "Co wnosi (wkład)", "Co robi w spółce", "Jak zarabia", "Ryzyko, które bierze"])
rows_ = [
 ("Grupa Azoty", f"={N['gash']}", "Produkty i produkcję, logistykę i sieć sprzedaży (punkty wydania nawozów), markę, klientów, 5 gotowych funkcjonalności (licencja)",
  "Sprzedaje nawozy przez platformę, przyciąga rolników, prowadzi spółkę (prezes, większość w radzie)", "51% gotówki spółki + większy udział w rynku, mix, niższy koszt kanału", "51% finansowania startu, integracja SAP/CRM"),
 ("Partner Technologiczny", f"={N['tpsh']}", "Silnik platformy agro (licencja dla spółki)", "Buduje i rozwija warstwę UX (web, aplikacja), integracje, mapy pól, AI; utrzymanie 24/7",
  "Koszt + 15% za budowę i utrzymanie + licencja 3 zł / użytkownik + 24,5% gotówki spółki", "24,5% finansowania startu"),
 ("Partner Marketplace / Handel (firma X)", f"={N['mpsh']}", "Gotowy marketplace: katalog, koszyk, płatności, panel sprzedawcy, reklama, bazę sprzedawców (użyczenie / licencja)",
  "Dostosowuje marketplace do agro, prowadzi operacje: sprzedawcy, katalog, spory", "Koszt + 15% za dostosowanie i operacje + licencja 0,5% GMV 3P + 24,5% gotówki spółki", "24,5% finansowania startu"),
]
for k, rw in enumerate(rows_):
    for j, t in enumerate(rw):
        c = Wk.cell(4 + k, 1 + j, t); c.alignment = Alignment(wrap_text=True, vertical="top")
        if j == 1: c.number_format = "0.0%"
Wk.cell(8, 1, "Wycena wkładów rzeczowych (do potwierdzenia)"); Wk.cell(8, 2, "mln zł")
vals = [("Grupa Azoty: 5 funkcjonalności (koszt odtworzenia)", N['f5']), ("Partner Technologiczny: silnik (koszt odtworzenia)", N['vtp']), ("Partner Marketplace: gotowy marketplace (koszt odtworzenia)", N['vmp'])]
for k, (lab, f) in enumerate(vals): Wk.cell(9 + k, 1, lab); Wk.cell(9 + k, 2, "=" + f).number_format = NUM
Wk.cell(12, 1, "GA wnosi też to, czego nie da się wycenić kosztem odtworzenia: produkty, produkcję, logistykę, klientów i markę. Dlatego proponujemy GA 51%.")
Wk.cell(14, 1, "Ład korporacyjny (benchmark: Bosch × BASF, praktyki JV)")
gov = ["Rada nadzorcza: GA 3 miejsca, partnerzy po 1. Prezes z GA, członkowie zarządu od partnerów (technologia, marketplace).",
       "Sprawy zastrzeżone (jednomyślnie): budżet roczny, nowy wspólnik, sprzedaż spółki, zmiana zasad rozliczeń.",
       "Impas: eskalacja do zarządów wspólników, potem mediator, na końcu opcja wykupu udziałów przez GA.",
       "Własność: kod stworzony w spółce i dane rolników należą do spółki GAia. Silnik i marketplace zostają własnością partnerów (licencja).",
       "Wyjście: po 3 latach GA ma prawo wykupu udziałów partnera po z góry ustalonej formule (np. wielokrotność EBITDA)."]
for k, t in enumerate(gov): Wk.cell(15 + k, 1, "• " + t)
Wk.column_dimensions["A"].width = 40; Wk.column_dimensions["B"].width = 10
for col, w in zip("CDEF", [52, 48, 44, 30]): Wk.column_dimensions[col].width = w
REG_W = [(3, 6, 1, 6), (8, 11, 1, 2)]

# ======================= BENCHMARKI =======================
Bm = wb.create_sheet("Benchmarki rynkowe")
Bm["A1"] = "Benchmarki: jak zarabiają marketplace'y i konkurencja, jak buduje się takie JV (źródła publiczne, 11.2026)"
Bm["A2"] = "Wnioski z benchmarków są podstawą stawek w arkuszu Założenia."
hdr(Bm, 3, ["Firma / temat", "Co to jest", "Liczby", "Na czym najwięcej zarabia", "Co bierzemy do modelu GAia", "Źródło"])
bm = [
 ("Allegro", "Największy marketplace w Polsce", "Prowizja (take rate) 12,0% w PL i reklama ok. 2,0% GMV (IV kw. 2024)", "Prowizje od sprzedawców i reklama (Allegro Ads rośnie najszybciej)", "Prowizja 7% (agro ma niższe marże), reklama 1,5%", "https://about.allegro.eu/wp-content/uploads/2026/01/Results-Presentation-Q4-2024.pdf"),
 ("Marketplace'y ogólne", "Stawki prowizji", "Amazon ok. 12%, eBay ok. 10%, Vinted ok. 5%", "Prowizja od sprzedawcy", "7% mieści się w przedziale", "https://origami-marketplace.com/en-gb/marketplace-take-rate-a-guide-for-marketplace-operators/"),
 ("Yara (YaraPlus, Atfarm)", "Producent nawozów z narzędziami cyfrowymi", "ok. 5 mln rolników; YaraPlus od 01.2026 w UK z John Deere; program punktów za zakup nawozów", "Na nawozach. Narzędzia są darmowe i mają zwiększać sprzedaż produktu", "Podstawa GAia bezpłatna; GA zarabia głównie na udziale w rynku", "https://farming.co.uk/news/yara-uk-launches-new-yaraplus-digital-platform-at-lamma-2026"),
 ("Yara × Agrofy", "Marketplace agro w Ameryce Płd., Yara inwestorem", "Runda 30 mln USD prowadzona przez Yarę (2021), łącznie 66 mln USD", "Prowizje od transakcji + fintech (płatności, kredyt)", "Producent nawozów buduje marketplace z partnerem; finansowanie w koszyku", "https://agfundernews.com/agrofy-adds-30m-in-yara-led-series-c-round"),
 ("BASF xarvio", "Aplikacja agronomiczna producenta ŚOR", ">130 tys. użytkowników, >20 mln ha (2025); Premium dawane gratis do opryskiwaczy John Deere (2021)", "Na środkach ochrony; aplikacja wspiera sprzedaż", "Premium tanie lub w pakiecie z produktem", "https://www.basf.com/global/en/media/news-releases/2025/09/p-25-176"),
 ("Bayer FieldView", "Platforma danych z pól", "UK: Prime od 120 £/rok, Plus od 500 £/rok; USA: Plus 649 USD/rok", "Abonamenty + wsparcie sprzedaży nasion i ŚOR", "Premium 240 zł/rok to ok. 1/2 FieldView Prime", "https://climate.com/en-gb/pricing.html"),
 ("SatAgro (Polska)", "Satelitarny monitoring pól", "10-15 zł brutto za ha rocznie", "Abonament za hektar", "Dla gospodarstwa 20 ha to 200-300 zł/rok, więc 240 zł/rok jest realne", "https://wiescirolnicze.pl/vlog/sledza-pola-z-kosmosu-za-10-zl-na-rok-video"),
 ("FBN (USA)", "Marketplace środków produkcji + finansowanie", "ok. 117-120 tys. gospodarstw; >1 mld USD finansowania (2024); członkostwo zmienione na bezpłatne", "Sprzedaż środków produkcji i finansowanie, nie abonament", "Abonament to mały przychód; główne: prowizje, reklama, finansowanie", "https://agfundernews.com/fbn-launches-free-membership-and-its-answer-to-ag-carbon-credit-markets"),
 ("DeHaat (Indie)", "Platforma agro z magazynami i skupem", "Przychód ok. 2 700 cr INR (FY24), ok. 79% ze sprzedaży płodów; strata operacyjna", "Handel własnym towarem (niska marża, wysoki koszt)", "Nie budujemy własnych magazynów: realizuje sieć GA", "https://entrackr.com/fintrackr/dehaat-reports-rs-2675-cr-gross-revenue-in-fy24-losses-shrink-34-7385674"),
 ("Agrando (Niemcy)", "Marketplace agro bez producenta", "ok. 20 tys. rolników, 12 mln EUR finansowania; zamknięty w 11.2022", "Prowizje - za mało, bez własnego produktu", "GAia ma producenta (GA) i gotowy marketplace (partner)", "https://agfundernews.com/agrando-raises-14m-in-europes-biggest-farmer-marketplace-series-a"),
 ("Bosch × BASF Smart Farming", "JV dwóch firm wnoszących technologię", "50/50, każdy wnosi swoją technologię, wspólni prezesi (2021); marka ONE SMART SPRAY (2023)", "Sprzedaż wspólnego rozwiązania", "Wkład technologiczny za udziały, wspólne zarządzanie", "https://www.basf.com/global/en/media/news-releases/2020/11/p-20-361"),
 ("Praktyki JV (Umbrex, Ankura, WIPO)", "Jak budować JV", "Udziały za wkład rzeczowy; IP lepiej jako licencja; usługi wspólników po koszcie; rada, sprawy zastrzeżone, ścieżka impasu; 50/50 tylko przy mocnym ładzie", "-", "GA 51% (kontrola), licencje partnerów, koszt + 15%, opcja wykupu", "https://umbrex.com/resources/joint-venture-playbook/ownership-capital-and-economic-arrangements/"),
]
for k, rw in enumerate(bm):
    for j, t in enumerate(rw):
        c = Bm.cell(4 + k, 1 + j, t); c.alignment = Alignment(wrap_text=True, vertical="top")
        if j == 5: c.hyperlink = t
for col, w in zip("ABCDEF", [24, 30, 46, 38, 44, 46]): Bm.column_dimensions[col].width = w

exec(open("xlsx_finish.py", encoding="utf-8").read())
wb.save(OUT); print("ok", OUT)
