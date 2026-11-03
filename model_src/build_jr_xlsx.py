# GAia: model wspólnej puli przychodów (joint revenue) trzech partnerów, 3 scenariusze.
# Uruchom: python3 build_jr_xlsx.py
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

OUT = "../GA_GAia_model_wspolnej_puli.xlsx"
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

A["A1"] = "GAia: założenia modelu wspólnej puli przychodów (do potwierdzenia danymi GA i ofertami partnerów)"
SCEN_NAMES = ["Ostrożny", "Bazowy", "Ambitny"]

sec("0. Wybór scenariusza")
inp("scen", "Scenariusz (1 = ostrożny, 2 = bazowy, 3 = ambitny)", 2, "", "Zmień na 1 lub 3. Porównanie wszystkich trzech: arkusz Scenariusze", "0")
sec("1. Użytkownicy (scenariusze)")
inp("u1", "Użytkownicy: scenariusz ostrożny", [20, 45, 70, 90, 100], "tys.", "≈ 8% rolników ARiMR w roku 5; skala xarvio (130 tys. na świecie)", "#,##0")
inp("u2", "Użytkownicy: scenariusz bazowy", [30, 70, 120, 165, 200], "tys.", "≈ 17% rolników ARiMR; Nutrien: 19% sprzedaży przez kanał cyfrowy po ok. 3 latach", "#,##0")
inp("u3", "Użytkownicy: scenariusz ambitny (cel 300 tys.)", [30, 80, 150, 230, 300], "tys.", "≈ 25% rolników ARiMR; Syngenta MAP: 2 mln zarejestrowanych w Chinach", "#,##0")
inp("act", "Aktywni kupujący", 0.35, "% użytk.", "ZAŁOŻENIE", "0%")
inp("prem", "Abonenci premium (podstawa bezpłatna)", [0.02, 0.03, 0.04, 0.05, 0.06], "% użytk.", "Benchmark: Yara, xarvio - podstawa za darmo", "0%")
inp("subp", "Cena abonamentu premium", 0.24, "tys. zł/rok", "20 zł/mies. Benchmark: SatAgro 10-15 zł/ha/rok; FieldView Prime UK 120 £/rok", "0.00")
inp("bf", "Zakupy nawozów GA przez platformę na kupującego", [6, 7, 8, 8.5, 9], "tys. zł/rok", "ZAŁOŻENIE", "0.0")
inp("b3", "Zakupy innych produktów na kupującego", [3, 4, 5, 6, 7], "tys. zł/rok", "ZAŁOŻENIE", "0.0")
inp("base", "Sprzedaż krajowa nawozów GA (kontrola skali)", "=2.93*0.7*1800", "mln zł/rok", "2,93 mln t × 70% × 1800 zł/t", "#,##0")
sec("2. Strumienie przychodów wspólnej puli")
inp("take", "Prowizja od sprzedawców innych produktów (3P)", 0.05, "% GMV 3P", "Benchmark: Allegro 12,0% (PL, IV kw. 2024). ŚOR i nasiona mają niskie marże → 5%", "0.0%")
inp("ads", "Reklama i promowane oferty", 0.015, "% GMV 3P", "Benchmark: Allegro Ads ok. 2,0% GMV", "0.0%")
inp("fin", "Finansowanie zakupów (prowizja od banku)", 0.004, "% GMV", "Benchmark: FBN >1 mld USD finansowania zakupów", "0.0%")
inp("data", "Analizy dla producentów nasion / ŚOR", [0.3, 0.6, 1.2, 2, 3], "mln zł/rok", "ZAŁOŻENIE ostrożne", "0.0")
inp("sfee", "Opłata za sukces: część DODATKOWEJ marży GA z platformy", 0.25, "%", "GA płaci tylko od zysku, którego bez platformy by nie miała. Benchmark: modele afiliacyjne / revenue share", "0%")
sec("3. Wzrost sprzedaży GA dzięki platformie")
inp("pp1", "Dodatkowy udział AN/CAN: ostrożny", [0.2, 0.4, 0.7, 0.9, 1.0], "p.p.", "ZAŁOŻENIE", "0.0")
inp("pp2", "Dodatkowy udział AN/CAN: bazowy", [0.3, 0.6, 1.1, 1.6, 2.0], "p.p.", "ZAŁOŻENIE", "0.0")
inp("pp3", "Dodatkowy udział AN/CAN: ambitny", [0.3, 0.8, 1.5, 2.2, 3.0], "p.p.", "Cel strategii 45% → 60%", "0.0")
inp("ppv", "Przychód z 1 p.p. udziału AN/CAN", 39, "mln zł/rok", "GA_5_luk, arkusz Szacunki", "0")
inp("cm", "Marża kontrybucyjna na dodatkowej tonie", 0.25, "%", "ZAŁOŻENIE - Kontroling", "0%")
inp("chan", "Niższy koszt kanału na zamówieniach z platformy", 0.02, "% GMV naw.", "ZAŁOŻENIE: krótsza ścieżka sprzedaży, mniej upustów", "0.0%")
inp("mix", "Poprawa mixu", 0.01, "% GMV naw.", "ZAŁOŻENIE", "0.0%")
inp("gaint", "Koszt wewnętrzny GA: opiekun projektu, dane, ARiMR", 0.3, "mln zł/rok", "2 osoby; GA nie wnosi gotówki do puli", "0.0")
sec("4. Koszty wspólne puli (po koszcie, bez marży)")
inp("fmc", "Koszt 1 osobomiesiąca zespołu partnerów", 0.02, "mln zł", "20 tys. zł/mies.", "0.000")
inp("tpf0", "Partner Technologiczny: zespół stały", 4, "FTE", "silnik gotowy, mały zespół", "0")
inp("tpfu", "Partner Technologiczny: +1 FTE na każde X tys. użytkowników", 25, "tys.", "zespół rośnie z użytkownikami", "0")
inp("tps0", "Utrzymanie 24/7: zespół stały", 2, "FTE", "", "0")
inp("tpsu", "Utrzymanie: +1 FTE na każde X tys. użytkowników", 50, "tys.", "", "0")
inp("tpsc", "Koszt 1 FTE utrzymania", 0.15, "mln zł/rok", "", "0.00")
inp("cl0", "Chmura i bezpieczeństwo: koszt stały", 1.0, "mln zł/rok", "", "0.0")
inp("clu", "Chmura, dane satelitarne, AI: na użytkownika", 10, "zł/rok", "", "0")
inp("mp0", "Partner Marketplace: zespół operacji stały", 3, "FTE", "katalog, sprzedawcy, spory", "0")
inp("mpu", "Operacje: +1 FTE na każde X tys. użytkowników", 70, "tys.", "", "0")
inp("mpc", "Koszt 1 FTE operacji", 0.18, "mln zł/rok", "", "0.00")
inp("mpi0", "Infrastruktura marketplace'u: stała", 0.3, "mln zł/rok", "", "0.0")
inp("mpiu", "Infrastruktura marketplace'u: na użytkownika", 2, "zł/rok", "", "0")
inp("payc", "Płatności", 0.006, "% GMV", "operator płatności", "0.0%")
inp("cac", "Pozyskanie użytkownika (CAC)", 50, "zł", "niski dzięki sieci sprzedaży GA, ARiMR i doradcom", "0")
inp("brand", "Kampania marki (stała)", 1.5, "mln zł/rok", "", "0.0")
inp("launch", "Kampania startowa (rok 0)", 1.0, "mln zł", "", "0.0")
inp("csu", "Obsługa klienta: 1 konsultant na X tys. użytkowników", 25, "tys.", "", "0")
inp("csc", "Koszt konsultanta", 0.12, "mln zł/rok", "", "0.00")
inp("pmo", "Wspólne biuro projektu (3 osoby)", 0.9, "mln zł/rok", "zamiast zarządu spółki w etapie 1", "0.0")
sec("5. Finansowanie deficytu puli i zasady wypłat")
inp("ftp", "Deficyt puli finansuje Partner Technologiczny", 0.7, "%", "Proporcja do kosztów bieżących obu partnerów; GA 0%", "0%")
inp("fmp", "Deficyt puli finansuje Partner Marketplace", "=1-B{x}", "%", "Formuła", "0%")
inp("wacc", "Stopa dyskontowa", 0.10, "%", "", "0%")
sec("6. Moc platformy")
inp("aovf", "Średnie zamówienie nawozów", 4.5, "tys. zł", "", "0.0")
inp("aov3", "Średnie zamówienie innych produktów", 1.2, "tys. zł", "", "0.0")
inp("pt", "Cena nawozu do przeliczenia na tony", 1800, "zł/t", "", "#,##0")
r_ftp = int(N["ftp"].split("$")[-1]); A.cell(int(N["fmp"].split("$")[-1]), 2, f"=1-B{r_ftp}")
A.column_dimensions["A"].width = 64
for c in "BCDEF": A.column_dimensions[c].width = 12
A.column_dimensions["G"].width = 14; A.column_dimensions["H"].width = 84
A.freeze_panes = "A5"

# ======================= KLUCZ PODZIAŁU =======================
K = wb.create_sheet("Klucz podziału")
K["A1"] = "Klucz podziału wspólnej puli: kto dostaje jaką część każdego strumienia przychodów"
K["A2"] = "Zasada: każdy dostaje najwięcej tam, gdzie tworzy wartość. Benchmark: sklepy z aplikacjami 70/30 (twórca/platforma), Salesforce AppExchange 15% dla platformy."
hdr(K, 4, ["Strumień przychodów", "Grupa Azoty", "Partner Technologiczny", "Partner Marketplace", "Suma", "Dlaczego tak"])
KEYS = [("s3", "Prowizje i reklama od sprzedawców innych produktów", (0.20, 0.40, 0.40), "Marketplace i sprzedawcy (PM) + aplikacja i ruch w niej (PT); rolnicy z sieci GA"),
        ("sp", "Abonament premium", (0.15, 0.70, 0.15), "Funkcje premium tworzy silnik PT. Benchmark: 70% dla twórcy aplikacji (Apple, Google)"),
        ("sf", "Finansowanie zakupów", (0.40, 0.20, 0.40), "Klienci i historia zakupów GA + płatności i koszyk PM"),
        ("sd", "Analizy dla producentów", (0.40, 0.40, 0.20), "Dane rolników i ARiMR (GA) + analityka silnika (PT)"),
        ("ss", "Opłata za sukces od GA (25% dodatkowej marży)", (0.00, 0.60, 0.40), "GA płaci partnerom za dodatkowy zysk; dzielą według pracy włożonej w sprzedaż")]
KR = {}
for k_, (key, lab, sh_, why) in enumerate(KEYS):
    r = 5 + k_; K.cell(r, 1, lab)
    for j, v_ in enumerate(sh_):
        c = K.cell(r, 2 + j, v_); c.font = BLUE; c.fill = INF; c.number_format = "0%"
    K.cell(r, 5, f"=SUM(B{r}:D{r})").number_format = "0%"; K.cell(r, 6, why); KR[key] = r
K.column_dimensions["A"].width = 50
for c in "BCDE": K.column_dimensions[c].width = 14
K.column_dimensions["F"].width = 80
kref = lambda key, p: f"'Klucz podziału'!${'BCD'[p]}${KR[key]}"

# ======================= WKŁAD TECHNOLOGICZNY (rok 0) =======================
C = wb.create_sheet("Wkład technologiczny")
C["A1"] = "Wkład technologiczny partnerów: dopracowanie GAia w roku 0 z własnych zasobów"
C["A2"] = "Łączymy gotowe elementy: silnik agro (PT) + marketplace przez plug-in (PM) + 5 funkcjonalności GA na Azure. Koszt = osobomiesiące × 20 tys. zł."
hdr(C, 4, ["Moduł", "Kto robi", "Co obejmuje", "Osobo-\nmiesiące", "% rok 0", "% rok 1", "Koszt\nmln zł", "Rok 0", "Rok 1"])
PTn, MPn = "Partner Technologiczny", "Partner Marketplace"
mods = [
 ("Dostosowanie silnika pod GAia", PTn, "konfiguracja, marka, uprawy", 24, 1, 0), ("Warstwa UX: web i aplikacja mobilna", PTn, "ekrany dla rolnika, iOS / Android", 60, .8, .2),
 ("Grywalizacja", PTn, "punkty, odznaki, wyzwania sezonowe, nagrody", 24, .6, .4), ("Integracja z ARiMR", PTn, "działki i uprawy z wniosków", 18, .8, .2),
 ("Integracja z Węzłem Krajowym / mObywatel", PTn, "logowanie profilem zaufanym", 12, 1, 0), ("Audyt i podłączenie 5 funkcjonalności GA na Azure", PTn, "audyt i API, bez przepisywania", 12, 1, 0),
 ("Podłączenie marketplace'u (plug-in)", PTn, "katalog, koszyk, zamówienia w aplikacji", 12, 1, 0), ("Integracja z SAP / CRM GA", PTn, "zamówienia, ceny, dostępność", 18, 1, 0),
 ("Testy, bezpieczeństwo, RODO", PTn, "testy z rolnikami, testy penetracyjne", 18, .7, .3), ("Kierowanie projektem", PTn, "PM, analityka", 12, .7, .3),
 ("Plug-in marketplace'u dla GAia", MPn, "API, wspólne logowanie", 12, 1, 0), ("Kategorie agro i karty produktów", MPn, "skład, dawki, rejestracja ŚOR", 12, 1, 0),
 ("Zamówienia nawozów do punktów wydania GA", MPn, "routing, odbiór, dostawa", 18, .8, .2), ("Płatności i kredyt kupiecki", MPn, "płatności, integracja z bankiem", 12, .6, .4),
 ("Onboarding sprzedawców agro", MPn, "producenci nasion, ŚOR, części", 12, .6, .4)]
cr = 5
for name, who, what, fm, a0, a1 in mods:
    C.cell(cr, 1, name); C.cell(cr, 2, who); C.cell(cr, 3, what)
    for j, v_ in enumerate([fm, a0, a1]):
        c = C.cell(cr, 4 + j, v_); c.font = BLUE; c.fill = INF; c.number_format = "0" if j == 0 else "0%"
    C.cell(cr, 7, f"=D{cr}*{N['fmc']}"); C.cell(cr, 8, f"=$G{cr}*E{cr}"); C.cell(cr, 9, f"=$G{cr}*F{cr}")
    for j in range(7, 10): C.cell(cr, j).number_format = NUM
    cr += 1
first, last = 5, cr - 1; CAPROW = {}
for who in (PTn, MPn):
    C.cell(cr, 1, f"Razem {who}").font = B; C.cell(cr, 2, who)
    C.cell(cr, 4, f'=SUMIF($B${first}:$B${last},$B{cr},D${first}:D${last})').font = B
    for col in "GHI": C[f"{col}{cr}"] = f'=SUMIF($B${first}:$B${last},$B{cr},{col}${first}:{col}${last})'; C[f"{col}{cr}"].font = B; C[f"{col}{cr}"].number_format = NUM
    CAPROW[who] = cr; cr += 1
C.cell(cr, 1, "WKŁAD TECHNOLOGICZNY RAZEM").font = B; C.cell(cr, 4, f"=D{CAPROW[PTn]}+D{CAPROW[MPn]}").font = B
for col in "GHI": C[f"{col}{cr}"] = f"={col}{CAPROW[PTn]}+{col}{CAPROW[MPn]}"; C[f"{col}{cr}"].font = B; C[f"{col}{cr}"].fill = TOT; C[f"{col}{cr}"].number_format = NUM
tot_r = cr
C.column_dimensions["A"].width = 46; C.column_dimensions["B"].width = 22; C.column_dimensions["C"].width = 48
for col in "DEFGHI": C.column_dimensions[col].width = 10
C.row_dimensions[4].height = 32
capref = lambda who, yi: f"'Wkład technologiczny'!${'HI'[yi]}${CAPROW[who]}"
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
U = Sheet("Użytkownicy i GMV", "Użytkownicy, obroty i zamówienia (wybrany scenariusz)")
U.sec("Użytkownicy")
U.row("u", "Użytkownicy zarejestrowani (tys.)", lambda i, c: f"CHOOSE({N['scen']},{y('u1',i)},{y('u2',i)},{y('u3',i)})", total=False, fmt="#,##0",
      note="Założenia sekcja 1: wiersz scenariusza wybranego w komórce Scenariusz (CHOOSE)")
U.row("nu", "Nowi użytkownicy w roku (tys.)", lambda i, c: f"{c}{U.R['u']}-{'0' if i == 0 else YC[i-1] + str(U.R['u'])}", fmt="#,##0", note="Użytkownicy w roku − rok wcześniej")
U.row("buy", "Aktywni kupujący (tys.)", lambda i, c: f"{c}{U.R['u']}*{N['act']}", total=False, fmt="#,##0", note="Użytkownicy × 35%")
U.row("pr", "Abonenci premium (tys.)", lambda i, c: f"{c}{U.R['u']}*{y('prem',i)}", total=False, fmt="#,##0.0", note="Użytkownicy × odsetek premium")
U.sec("Obroty (GMV), mln zł")
U.row("fg", "GMV nawozów GA", lambda i, c: f"{c}{U.R['buy']}*{y('bf',i)}", note="Kupujący (tys.) × zakupy nawozów (tys. zł)")
U.row("g3", "GMV innych produktów od sprzedawców (3P)", lambda i, c: f"{c}{U.R['buy']}*{y('b3',i)}", note="Kupujący (tys.) × zakupy innych produktów (tys. zł)")
U.row("gmv", "GMV RAZEM", lambda i, c: f"{c}{U.R['fg']}+{c}{U.R['g3']}", bold=True, fill=TOT, note="Nawozy + inne produkty")
U.row("pen", "Nawozy przez platformę / sprzedaż krajowa GA", lambda i, c: f"{c}{U.R['fg']}/{N['base']}", total=False, fmt="0.0%", note="GMV nawozów ÷ 3,69 mld zł")
U.row("ot", "Zamówienia razem (tys.)", lambda i, c: f"{c}{U.R['fg']}/{N['aovf']}+{c}{U.R['g3']}/{N['aov3']}", fmt="#,##0", note="GMV ÷ średnie zamówienie")
U.row("tn", "Nawozy przez platformę (tys. t)", lambda i, c: f"{c}{U.R['fg']}*1000/{N['pt']}", fmt="#,##0", note="GMV nawozów ÷ 1800 zł/t")
u = lambda k, c: U.ref(k, c)
COLS = ["B"] + YC

# ---------- Partner Technologiczny: koszty ----------
T = Sheet("Partner Technologiczny", "Partner Technologiczny: wkład, koszty bieżące (pokrywa pula) i wynik")
T.sec("A. Wkład w roku 0-1 z własnych zasobów")
T.row("cap", "Dopracowanie GAia: UX, aplikacja, grywalizacja, ARiMR, Węzeł Krajowy, plug-in", lambda i, c: capref(PTn, 1) if i == 0 else "0", y0=capref(PTn, 0), note="Arkusz Wkład technologiczny")
T.sec("B. Koszty bieżące (zwraca je wspólna pula, po koszcie)")
T.row("dev", "Zespół rozwoju", lambda i, c: f"({N['tpf0']}+{u('u',c)}/{N['tpfu']})*12*{N['fmc']}", note="(4 + użytkownicy ÷ 25 tys.) FTE × 12 × 20 tys. zł")
T.row("sup", "Utrzymanie 24/7", lambda i, c: f"({N['tps0']}+{u('u',c)}/{N['tpsu']})*{N['tpsc']}", note="(2 + użytkownicy ÷ 50 tys.) FTE × 0,15 mln zł")
T.row("cl", "Chmura, dane satelitarne, AI, bezpieczeństwo", lambda i, c: f"{N['cl0']}+{u('u',c)}*1000*{N['clu']}/1000000", note="1 mln zł + użytkownicy × 10 zł")
T.row("opex", "KOSZTY BIEŻĄCE RAZEM", lambda i, c: f"SUM({c}{T.R['dev']}:{c}{T.R['cl']})", bold=True, note="Suma; pula zwraca je w całości")
T.row("cpu", "Koszt na użytkownika (zł/rok)", lambda i, c: f"{c}{T.R['opex']}/{u('u',c)}*1000", total=False, fmt="#,##0", note="Koszty bieżące ÷ użytkownicy")
Mp = Sheet("Partner Marketplace", "Partner Marketplace / Handel: wkład, koszty bieżące (pokrywa pula) i wynik")
Mp.sec("A. Wkład w roku 0-1 z własnych zasobów")
Mp.row("cap", "Plug-in, kategorie agro, zamówienia do punktów wydania, płatności", lambda i, c: capref(MPn, 1) if i == 0 else "0", y0=capref(MPn, 0), note="Arkusz Wkład technologiczny")
Mp.sec("B. Koszty bieżące (zwraca je wspólna pula, po koszcie)")
Mp.row("team", "Zespół operacji marketplace'u", lambda i, c: f"({N['mp0']}+{u('u',c)}/{N['mpu']})*{N['mpc']}", note="(3 + użytkownicy ÷ 70 tys.) FTE × 0,18 mln zł")
Mp.row("inf", "Infrastruktura marketplace'u", lambda i, c: f"{N['mpi0']}+{u('u',c)}*1000*{N['mpiu']}/1000000", note="0,3 mln zł + użytkownicy × 2 zł")
Mp.row("opex", "KOSZTY BIEŻĄCE RAZEM", lambda i, c: f"{c}{Mp.R['team']}+{c}{Mp.R['inf']}", bold=True, note="Suma; pula zwraca je w całości")

# ---------- Grupa Azoty: dodatkowa marża (podstawa opłaty za sukces) ----------
G = Sheet("Grupa Azoty", "Grupa Azoty: dodatkowa marża z platformy, opłata za sukces i wynik (bez wkładu pieniężnego)")
G.sec("A. Dodatkowa marża GA dzięki platformie")
G.row("sh", "Większy udział AN/CAN: p.p. × 39 mln zł × 25%", lambda i, c: f"CHOOSE({N['scen']},{y('pp1',i)},{y('pp2',i)},{y('pp3',i)})*{N['ppv']}*{N['cm']}", note="Dodatkowe p.p. (scenariusz) × przychód z 1 p.p. × marża kontrybucyjna")
G.row("ch", "Niższy koszt kanału: GMV nawozów × 2%", lambda i, c: f"{u('fg',c)}*{N['chan']}", note="Krótsza ścieżka sprzedaży, mniej upustów")
G.row("mx", "Lepszy mix: GMV nawozów × 1%", lambda i, c: f"{u('fg',c)}*{N['mix']}", note="Więcej NPK i nawozów specjalistycznych")
G.row("extra", "DODATKOWA MARŻA GA RAZEM", lambda i, c: f"SUM({c}{G.R['sh']}:{c}{G.R['mx']})", bold=True, note="Podstawa opłaty za sukces")

# ---------- Wspólna pula ----------
J = Sheet("Wspólna pula", "Wspólna pula przychodów: przychody → koszty wspólne po koszcie → zwrot deficytu → podział według klucza")
J.sec("1. Przychody puli (strumienie)")
J.row("s3", "Prowizje i reklama od sprzedawców 3P: GMV 3P × (5% + 1,5%)", lambda i, c: f"{u('g3',c)}*({N['take']}+{N['ads']})", note="Benchmark Allegro: 12% prowizji + 2% reklamy; agro ostrożniej")
J.row("sp", "Abonament premium: abonenci × 240 zł", lambda i, c: f"{u('pr',c)}*{N['subp']}", note="20 zł / mies.; podstawa bezpłatna")
J.row("sf", "Finansowanie zakupów: GMV × 0,4%", lambda i, c: f"{u('gmv',c)}*{N['fin']}", note="Prowizja od banku za kredyt kupiecki")
J.row("sd", "Analizy dla producentów", lambda i, c: f"{y('data',i)}", note="Zanonimizowane dane dla producentów nasion / ŚOR")
J.row("ss", "Opłata za sukces od GA: 25% dodatkowej marży GA", lambda i, c: f"{G.ref('extra',c)}*{N['sfee']}", note="GA płaci tylko od zysku, którego bez platformy by nie miała")
J.row("pool", "PRZYCHODY PULI RAZEM", lambda i, c: f"SUM({c}{J.R['s3']}:{c}{J.R['ss']})", bold=True, fill=TOT, note="Suma strumieni")
J.sec("2. Koszty wspólne (po koszcie, bez marży)")
J.row("ctp", "Koszty bieżące Partnera Technologicznego", lambda i, c: T.ref("opex", c), note="Arkusz Partner Technologiczny")
J.row("cmp", "Koszty bieżące Partnera Marketplace", lambda i, c: Mp.ref("opex", c), note="Arkusz Partner Marketplace")
J.row("cpay", "Płatności: GMV × 0,6%", lambda i, c: f"{u('gmv',c)}*{N['payc']}", note="Operator płatności")
J.row("cmkt", "Marketing: nowi × 50 zł + marka 1,5 mln zł", lambda i, c: f"{u('nu',c)}*1000*{N['cac']}/1000000+{N['brand']}", y0=N["launch"], note="Rok 0: kampania startowa")
J.row("ccs", "Obsługa klienta", lambda i, c: f"{u('u',c)}/{N['csu']}*{N['csc']}", note="1 konsultant na 25 tys. użytkowników")
J.row("cpm", "Wspólne biuro projektu (3 osoby)", lambda i, c: N["pmo"], y0=f"{N['pmo']}*0.5", note="Rok 0: pół roku")
J.row("cost", "KOSZTY WSPÓLNE RAZEM", lambda i, c: f"SUM({c}{J.R['ctp']}:{c}{J.R['cpm']})", y0=f"SUM(B{J.R['ctp']}:B{J.R['cpm']})", bold=True, note="Suma kosztów")
J.row("net", "WYNIK PULI (przychody − koszty wspólne)", lambda i, c: f"{c}{J.R['pool']}-{c}{J.R['cost']}", y0=f"-B{J.R['cost']}", bold=True, fill=TOT, note="Do podziału po zwrocie deficytu")
J.cum("net", "Wynik puli narastająco")
J.sec("3. Deficyt i jego zwrot (GA nie dopłaca)")
J.row("fund", "Deficyt puli (finansują PT 70% i PM 30%)", lambda i, c: f"MAX(0,-{c}{J.R['net']})", y0=f"MAX(0,-B{J.R['net']})", note="Gdy wynik puli < 0")
J.row("bal", "Deficyt do zwrotu (saldo)", lambda i, c: "0", y0="0", total=False, note="Poprzedni rok + nowy deficyt − zwrot")
J.row("rep", "Zwrot deficytu (pierwszeństwo przed podziałem)", lambda i, c: "0", y0="0", note="Dodatni wynik puli najpierw spłaca deficyt")
J.row("dist", "DO PODZIAŁU WEDŁUG KLUCZA", lambda i, c: "0", y0="0", bold=True, note="Wynik puli po zwrocie deficytu")
br, rr, dr_, nr = J.R["bal"], J.R["rep"], J.R["dist"], J.R["net"]
J.ws[f"B{rr}"] = "=0"; J.ws[f"B{br}"] = f"=B{J.R['fund']}"; J.ws[f"B{dr_}"] = f"=MAX(0,B{nr})"
for i, col in enumerate(YC):
    prev = "B" if i == 0 else YC[i - 1]
    J.ws[f"{col}{rr}"] = f"=MIN({prev}{br},MAX(0,{col}{nr}))"; J.ws[f"{col}{br}"] = f"={prev}{br}+{col}{J.R['fund']}-{col}{rr}"; J.ws[f"{col}{dr_}"] = f"=MAX(0,{col}{nr})-{col}{rr}"
J.sec("4. Udział partnerów w przychodach puli (klucz z arkusza Klucz podziału)")
streams = ["s3", "sp", "sf", "sd", "ss"]
for p, (pk, pn) in enumerate([("gGA", "Grupa Azoty"), ("gPT", "Partner Technologiczny"), ("gPM", "Partner Marketplace")]):
    J.row(pk, f"Przychody przypisane: {pn}", lambda i, c, p=p: "+".join(f"{c}{J.R[s]}*{kref(s, p)}" for s in streams), note="Σ strumień × klucz partnera")
for p, (pk, wk, pn) in enumerate([("gGA", "wGA", "Grupa Azoty"), ("gPT", "wPT", "Partner Technologiczny"), ("gPM", "wPM", "Partner Marketplace")]):
    J.row(wk, f"Udział w puli: {pn}", lambda i, c, pk=pk: f"IF({c}{J.R['pool']}=0,0,{c}{J.R[pk]}/{c}{J.R['pool']})", total=False, fmt="0%", note="Przychody przypisane ÷ przychody puli")

# ---------- wyniki partnerów ----------
T.sec("C. Z czego i kiedy zarabia Partner Technologiczny")
T.row("own", "0. Wkład własny (praca w roku 0-1)", lambda i, c: f"-{c}{T.R['cap']}", y0=f"-B{T.R['cap']}", note="Minus wkład z własnych zasobów")
T.row("reimb", "1. Koszty bieżące − ich zwrot z puli (= 0, pula płaci po koszcie)", lambda i, c: f"{c}{T.R['opex']}-{c}{T.R['opex']}", note="Pula pokrywa koszty: zysk 0, strata 0")
T.row("fin", "2. Finansowanie deficytu puli i jego zwrot (70%)", lambda i, c: f"(-{J.ref('fund',c)}+{J.ref('rep',c)})*{N['ftp']}", y0=f"(-{J.ref('fund','B')}+{J.ref('rep','B')})*{N['ftp']}", note="Minus dopłata, plus zwrot")
T.row("dist", "3. Udział w podziale puli", lambda i, c: f"{J.ref('dist',c)}*{J.ref('wPT',c)}", y0=f"{J.ref('dist','B')}*{J.ref('wPT','B')}", note="Do podziału × udział PT w puli")
T.row("res", "WYNIK PARTNERA TECHNOLOGICZNEGO", lambda i, c: f"SUM({c}{T.R['own']}:{c}{T.R['dist']})", y0=f"SUM(B{T.R['own']}:B{T.R['dist']})", bold=True, fill=TOT, note="Suma")
T.cum("res", "Wynik skumulowany")
Mp.sec("C. Z czego i kiedy zarabia Partner Marketplace")
Mp.row("own", "0. Wkład własny (praca w roku 0-1)", lambda i, c: f"-{c}{Mp.R['cap']}", y0=f"-B{Mp.R['cap']}", note="Minus wkład z własnych zasobów")
Mp.row("reimb", "1. Koszty bieżące − ich zwrot z puli (= 0, pula płaci po koszcie)", lambda i, c: f"{c}{Mp.R['opex']}-{c}{Mp.R['opex']}", note="Pula pokrywa koszty: zysk 0, strata 0")
Mp.row("fin", "2. Finansowanie deficytu puli i jego zwrot (30%)", lambda i, c: f"(-{J.ref('fund',c)}+{J.ref('rep',c)})*{N['fmp']}", y0=f"(-{J.ref('fund','B')}+{J.ref('rep','B')})*{N['fmp']}", note="Minus dopłata, plus zwrot")
Mp.row("dist", "3. Udział w podziale puli", lambda i, c: f"{J.ref('dist',c)}*{J.ref('wPM',c)}", y0=f"{J.ref('dist','B')}*{J.ref('wPM','B')}", note="Do podziału × udział PM w puli")
Mp.row("res", "WYNIK PARTNERA MARKETPLACE", lambda i, c: f"SUM({c}{Mp.R['own']}:{c}{Mp.R['dist']})", y0=f"SUM(B{Mp.R['own']}:B{Mp.R['dist']})", bold=True, fill=TOT, note="Suma")
Mp.cum("res", "Wynik skumulowany")
G.sec("B. Wynik Grupy Azoty")
G.row("keep", "1. Dodatkowa marża, która zostaje w GA (75%)", lambda i, c: f"{c}{G.R['extra']}-{J.ref('ss',c)}", note="Dodatkowa marża − opłata za sukces")
G.row("dist", "2. Udział w podziale puli", lambda i, c: f"{J.ref('dist',c)}*{J.ref('wGA',c)}", y0=f"{J.ref('dist','B')}*{J.ref('wGA','B')}", note="Do podziału × udział GA w puli")
G.row("int", "3. Koszt wewnętrzny GA (opiekun projektu, dane)", lambda i, c: f"-{N['gaint']}", note="GA nie wnosi gotówki do puli")
G.row("res", "WYNIK GRUPY AZOTY", lambda i, c: f"SUM({c}{G.R['keep']}:{c}{G.R['int']})", y0=f"SUM(B{G.R['keep']}:B{G.R['int']})", bold=True, fill=TOT, note="Suma")
G.cum("res", "Wynik skumulowany")

# ---------- Moc platformy ----------
Mc = Sheet("Moc platformy", "Moc platformy w liczbach (wybrany scenariusz)")
for key, lab, f_, fm in [("u", "Użytkownicy (tys.)", lambda i, c: u("u", c), "#,##0"), ("b", "Aktywni kupujący (tys.)", lambda i, c: u("buy", c), "#,##0"),
                         ("g", "GMV razem (mln zł)", lambda i, c: u("gmv", c), NUM), ("o", "Zamówienia razem (tys.)", lambda i, c: u("ot", c), "#,##0"),
                         ("t", "Nawozy przez platformę (tys. t)", lambda i, c: u("tn", c), "#,##0"), ("pe", "Nawozy przez platformę / sprzedaż krajowa GA", lambda i, c: u("pen", c), "0.0%"),
                         ("r", "Przychody puli (mln zł)", lambda i, c: J.ref("pool", c), NUM), ("tr", "Przychody puli / GMV", lambda i, c: f"{J.ref('pool',c)}/{u('gmv',c)}", "0.0%"),
                         ("ru", "Przychody puli na użytkownika (zł/rok)", lambda i, c: f"{J.ref('pool',c)}/{u('u',c)}*1000", "#,##0"),
                         ("cu", "Koszty wspólne na użytkownika (zł/rok)", lambda i, c: f"{J.ref('cost',c)}/{u('u',c)}*1000", "#,##0")]:
    Mc.row(key, lab, f_, total=False, fmt=fm)

# ======================= PODSUMOWANIE =======================
P = wb.create_sheet("Każdy partner zyskuje", 0)
P["A1"] = "GAia: wspólna pula przychodów trzech partnerów. Każdy zarabia"
P["A2"] = "Gotówka, mln zł, wybrany scenariusz (Założenia, sekcja 0). Porównanie trzech scenariuszy: arkusz Scenariusze."
P.cell(3, 1, "=\"Scenariusz: \"&CHOOSE(" + N['scen'] + ",\"ostrożny\",\"bazowy\",\"ambitny\")")
hdr(P, 4, ["Partner", "Rok 0", "Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5", "Suma 0-5", "NPV (10%)", "Wynik roczny > 0 od", "Zwrot (suma > 0) od"])
parts = [("Grupa Azoty", G, "res"), ("Partner Technologiczny", T, "res"), ("Partner Marketplace / Handel", Mp, "res"), ("Wspólna pula: wynik", J, "net")]
for k, (name, sh, key) in enumerate(parts):
    pr = 5 + k; P.cell(pr, 1, name)
    for j, col in enumerate(["B"] + YC): P.cell(pr, 2 + j, f"={sh.ref(key, col)}")
    P.cell(pr, 8, f"=SUM(B{pr}:G{pr})"); P.cell(pr, 9, f"=B{pr}+NPV({N['wacc']},C{pr}:G{pr})")
    P.cell(pr, 10, f'=IF(COUNTIF(C{pr}:G{pr},">0")=0,"po roku 5","rok "&SUMPRODUCT(MIN(IF(C{pr}:G{pr}>0,{{1,2,3,4,5}},99))))')
    P.cell(pr, 11, f'=IF(B{pr}+C{pr}>0,"rok 1",IF(SUM(B{pr}:D{pr})>0,"rok 2",IF(SUM(B{pr}:E{pr})>0,"rok 3",IF(SUM(B{pr}:F{pr})>0,"rok 4",IF(SUM(B{pr}:G{pr})>0,"rok 5","po roku 5")))))')
    for c in range(2, 10): P.cell(pr, c).number_format = NUM
SK = 10
P.cell(SK, 1, "Skala")
for j, h in enumerate(["Rok 1", "Rok 2", "Rok 3", "Rok 4", "Rok 5"]): P.cell(SK, 3 + j, h)
kpis = [("Użytkownicy (tys.)", U, "u", "#,##0"), ("GMV razem (mln zł)", U, "gmv", NUM), ("Nawozy przez platformę / sprzedaż krajowa GA", U, "pen", "0.0%"),
        ("Przychody puli (mln zł)", J, "pool", NUM), ("Koszty wspólne (mln zł)", J, "cost", NUM), ("Do podziału według klucza (mln zł)", J, "dist", NUM)]
for k, (lab, sh, key, fm) in enumerate(kpis):
    pr = SK + 1 + k; P.cell(pr, 1, lab)
    for j, col in enumerate(YC): P.cell(pr, 3 + j, f"={sh.ref(key, col)}").number_format = fm
IV = SK + len(kpis) + 2
P.cell(IV, 1, "Wkład i finansowanie"); P.cell(IV, 3, "mln zł")
inv = [("Wkład technologiczny Partnera Technologicznego (rok 0-1)", f"'Wkład technologiczny'!$G${CAPROW[PTn]}"), ("Wkład technologiczny Partnera Marketplace (rok 0-1)", f"'Wkład technologiczny'!$G${CAPROW[MPn]}"),
       ("Wkład pieniężny Grupy Azoty", "0"), ("Maks. deficyt puli (finansują PT 70% / PM 30%)", f"-MIN(0,MIN('Wspólna pula'!B{J.R['net_cum']}:G{J.R['net_cum']}))")]
for k, (lab, f) in enumerate(inv): P.cell(IV + 1 + k, 1, lab); P.cell(IV + 1 + k, 3, "=" + f).number_format = NUM
P.column_dimensions["A"].width = 56
for c in "BCDEFGHIJK": P.column_dimensions[c].width = 12
REG_P = [(4, 8, 1, 11), (SK, SK + len(kpis), 1, 7), (IV, IV + len(inv), 1, 3)]

# ======================= SCENARIUSZE (wartości z 3 przeliczeń) =======================
S3 = wb.create_sheet("Scenariusze", 1)
S3["A1"] = "Trzy scenariusze obok siebie (wartości przeliczone; aktualizuje skrypt add_scenarios.py)"
S3["A2"] = "Każdy scenariusz przeliczony przy tych samych zasadach podziału. Zmień scenariusz w Założeniach, aby zobaczyć szczegóły."
hdr(S3, 4, ["Wskaźnik", "Ostrożny", "Bazowy", "Ambitny", "Jedn."])
SCEN_ROWS = [("Użytkownicy w roku 5", "tys."), ("GMV w roku 5", "mln zł"), ("Nawozy przez platformę / sprzedaż krajowa GA (rok 5)", "%"),
             ("Przychody puli w roku 5", "mln zł"), ("Wynik puli w roku 5", "mln zł"), ("Grupa Azoty: suma 5 lat", "mln zł"), ("Grupa Azoty: NPV", "mln zł"),
             ("Partner Technologiczny: suma 5 lat", "mln zł"), ("Partner Technologiczny: NPV", "mln zł"), ("Partner Technologiczny: zwrot od", ""),
             ("Partner Marketplace: suma 5 lat", "mln zł"), ("Partner Marketplace: NPV", "mln zł"), ("Partner Marketplace: zwrot od", ""), ("Maks. deficyt puli", "mln zł")]
for k, (lab, un) in enumerate(SCEN_ROWS): S3.cell(5 + k, 1, lab); S3.cell(5 + k, 5, un)
S3.column_dimensions["A"].width = 56
for c in "BCDE": S3.column_dimensions[c].width = 14

# ======================= WKŁAD PARTNERÓW I ZASADY =======================
Wk = wb.create_sheet("Wkład i zasady", 2)
Wk["A1"] = "Kto co wnosi i jak dzielimy wspólne przychody"
Wk["A2"] = "Etap 1 (lata 0-2): umowa partnerska z podziałem przychodów (bez spółki). Etap 2 (od roku 3, po spełnieniu bramki KPI): spółka GAia z udziałami według tego samego klucza."
hdr(Wk, 3, ["Partner", "Co wnosi", "Co robi", "Co dostaje", "Ryzyko, które bierze"])
rows_ = [("Grupa Azoty", "Produkty, produkcja, logistyka i punkty wydania; współpraca z ARiMR i dane; klienci i marka; 5 funkcjonalności na Azure (audyt, bez wyceny)",
          "Udostępnia dane i funkcjonalności, sprzedaje nawozy przez GAia, przyciąga rolników swoją siecią", "75% dodatkowej marży + udział w puli (20% prowizji, 40% finansowania i analiz)", "Bez wkładu pieniężnego. Oddaje 25% dodatkowego zysku"),
         ("Partner Technologiczny (główny)", "Gotowy silnik agro; zespół, który dopracowuje GAia w roku 0", "UX, aplikacja, grywalizacja, integracje (ARiMR, Węzeł Krajowy, SAP, plug-in); utrzymanie",
          "Zwrot kosztów bieżących z puli + 40% prowizji, 70% premium, 60% opłaty za sukces", "Wkład pracy w roku 0, 70% deficytu na starcie"),
         ("Partner Marketplace / Handel", "Gotowy marketplace (plug-in), sprzedawcy, płatności", "Plug-in, kategorie agro, zamówienia do punktów wydania; operacje marketplace'u",
          "Zwrot kosztów bieżących z puli + 40% prowizji, 40% finansowania, 40% opłaty za sukces", "Wkład pracy w roku 0, 30% deficytu na starcie")]
for k, rw in enumerate(rows_):
    for j, t in enumerate(rw):
        c = Wk.cell(4 + k, 1 + j, t); c.alignment = Alignment(wrap_text=True, vertical="top")
Wk.cell(8, 1, "Zasady wspólnej puli (kolejność wypłat)")
rules = ["1. Wszystkie przychody platformy wpływają do wspólnej puli (rachunek powierniczy / escrow).",
         "2. Z puli najpierw płacimy koszty wspólne po koszcie, bez marży: zespoły partnerów, płatności, marketing, obsługa klienta, biuro projektu.",
         "3. Jeśli wynik puli jest ujemny, deficyt finansują Partner Technologiczny (70%) i Partner Marketplace (30%). GA nie dopłaca.",
         "4. Dodatni wynik puli najpierw zwraca deficyt partnerom, potem dzielimy go według klucza (arkusz Klucz podziału).",
         "5. Opłata za sukces: GA wpłaca do puli 25% dodatkowej marży, którą osiąga dzięki platformie. Od sprzedaży, którą GA i tak by zrobiła, nie płaci nic.",
         "6. Audyt roczny puli i kosztów partnerów przez niezależnego audytora."]
for k, t in enumerate(rules): Wk.cell(9 + k, 1, t)
Wk.cell(16, 1, "Etapy i bramki decyzyjne (benchmark: Bosch i BASF współpracowali od 2019, spółkę założyli w 2021)")
gates = [("Bramka 0 (koniec roku 0)", "Aplikacja i marketplace połączone, audyt 5 funkcjonalności GA, pilot gotowy", "start pilota"),
         ("Bramka 1 (wiosna roku 1)", "Pilot: ≥ 10 tys. użytkowników, ≥ 5 tys. zamówień, ocena rolników ≥ 4/5", "skalowanie na kraj"),
         ("Bramka 2 (koniec roku 2)", "≥ 70 tys. użytkowników, przychody puli ≥ 10 mln zł, wynik puli ≥ 0", "założenie spółki GAia (etap 2)"),
         ("Wyjście", "Jeśli bramka nie jest spełniona", "zakończenie bez straty dla GA; partnerzy zachowują swoje silniki")]
for j, h in enumerate(["Bramka", "Warunek (KPI)", "Decyzja"]): Wk.cell(17, 1 + j, h)
for k, rw in enumerate(gates):
    for j, t in enumerate(rw): c = Wk.cell(18 + k, 1 + j, t); c.alignment = Alignment(wrap_text=True, vertical="top")
Wk.column_dimensions["A"].width = 40
for col, w in zip("BCDE", [58, 52, 52, 34]): Wk.column_dimensions[col].width = w
REG_W = [(3, 6, 1, 5), (17, 21, 1, 3)]

# ======================= POZYTYWNE CASE STUDY =======================
Bm = wb.create_sheet("Case study i benchmarki")
Bm["A1"] = "Pozytywne case study i benchmarki podziału przychodów (źródła publiczne, 11.2026)"
Bm["A2"] = "Co działa u innych i co z tego bierzemy do GAia."
hdr(Bm, 3, ["Firma / współpraca", "Kto co wnosi", "Wynik (liczby)", "Jak się dzielą / zarabiają", "Co bierzemy do GAia", "Źródło"])
bm = [
 ("Nutrien Digital (USA / Kanada)", "Producent i dystrybutor nawozów buduje własny kanał cyfrowy", "Sprzedaż cyfrowa >1,2 mld USD w 2020 r. (4× więcej niż 2019); ok. 19% sprzedaży w I poł. 2021", "Platforma jest kanałem sprzedaży własnych produktów", "Producent nawozów może sprzedać przez platformę znaczną część wolumenu", "https://www.businesswire.com/news/home/20210217006038/en"),
 ("Syngenta MAP (Chiny)", "Producent ŚOR + centra usług + aplikacja", "Przychód ok. 1,8 mld USD w 2021 r. (2×), 492 → 638 centrów, >2 mln rolników", "Sprzedaż produktów i usług przez platformę i partnerów lokalnych", "Platforma + sieć fizyczna (u nas punkty wydania GA) skaluje się szybko", "https://www.syngentagroup.com/newsroom/2022/syngenta-group-reports-2021-performance-growing-23-282-billion-sales"),
 ("Starbucks × PepsiCo (NACP)", "Marka i produkt (Starbucks) + dystrybucja (PepsiCo)", "JV 50/50 od 1994 r.; stworzyli rynek kawy gotowej do picia w USA (przychód ok. 2 mld USD - dane nieoficjalne)", "Każdy wnosi to, w czym jest najlepszy; wspólny zysk", "GA wnosi produkt i klientów, partnerzy technologię i kanał", "https://contact.pepsico.com/starbucks/about-us"),
 ("Bosch × BASF (ONE SMART SPRAY)", "Technologia Bosch + agronomia BASF (xarvio)", "Współpraca od 2019, spółka 50/50 od 2021; system w maszynach Amazone, AGCO, CNH, Stara (2023-2024)", "Wspólna spółka, każdy wnosi technologię", "Najpierw współpraca, potem spółka po potwierdzeniu wyników", "https://bosch-presse.de/pressportal/de/en/first-crop-protection-sprayer-featuring-one-smart-sprays-weed-management-system-launched-in-latin-america-by-stara-265929.html"),
 ("Bayer × Microsoft", "Modele agronomiczne Bayer + chmura Microsoft", "Partnerstwo od 2021; usługi Bayer na Azure od 2023", "Każdy licencjonuje swoje rozwiązanie na wspólnej platformie", "Silnik PT i marketplace PM zostają u właścicieli, działają na licencji", "https://www.businesswire.com/news/home/20230314005312/en/"),
 ("Yara × Agrofy", "Producent nawozów inwestuje w marketplace agro", "Runda 30 mln USD prowadzona przez Yarę (2021), łącznie 66 mln USD", "Prowizje od transakcji + fintech", "Producent i marketplace łączą siły; finansowanie w koszyku", "https://agfundernews.com/agrofy-adds-30m-in-yara-led-series-c-round"),
 ("Sklepy z aplikacjami (Apple, Google)", "Platforma + twórcy aplikacji", "Podział 70/30: 70% dla twórcy, 30% dla platformy", "Twórca funkcji dostaje większość przychodu z niej", "Premium: 70% dla Partnera Technologicznego", "https://www.entrepreneur.com/starting-a-business/3-lesser-known-app-marketplaces-for-entrepreneurs-seeking/323858"),
 ("Salesforce AppExchange", "Platforma + partnerzy z aplikacjami", "15% przychodu dla platformy", "Platforma bierze małą część, partner większość", "Klucz podziału zależny od tego, kto tworzy wartość", "https://developer.salesforce.com/docs/atlas.en-us.252.0.packagingGuide.meta/packagingGuide/appexchange_checkout_rev_share.htm"),
 ("Allegro", "Marketplace", "Prowizja 12,0% i reklama ok. 2,0% GMV (IV kw. 2024)", "Prowizje i reklama od sprzedawców", "Prowizja 5% i reklama 1,5% (agro ma niższe marże)", "https://about.allegro.eu/wp-content/uploads/2026/01/Results-Presentation-Q4-2024.pdf"),
 ("Lekcja: Agrando (Niemcy)", "Sam marketplace bez producenta", "ok. 20 tys. rolników, zamknięty w 11.2022", "Same prowizje to za mało", "GAia ma producenta, produkt i logistykę", "https://agfundernews.com/agrando-raises-14m-in-europes-biggest-farmer-marketplace-series-a"),
]
for k, rw in enumerate(bm):
    for j, t in enumerate(rw):
        c = Bm.cell(4 + k, 1 + j, t); c.alignment = Alignment(wrap_text=True, vertical="top")
        if j == 5: c.hyperlink = t
for col, w in zip("ABCDEF", [26, 32, 46, 34, 40, 44]): Bm.column_dimensions[col].width = w

# ======================= konfiguracja wykończenia =======================
NOTES = {}
CALC_SHEETS = [(U, "2E6DB4"), (T, "2E6DB4"), (Mp, "2E6DB4"), (G, "003665"), (J, "003665"), (Mc, "6F9FD8")]
SPEC = [("Grupa Azoty", "Wynik partnera", "Grupa Azoty", G, "res"), ("Partner Technologiczny", "Wynik partnera", "Partner Technologiczny", T, "res"),
        ("Partner Marketplace", "Wynik partnera", "Partner Marketplace", Mp, "res"),
        ("Pula", "Przychody puli", "1 Prowizje i reklama 3P", J, "s3"), ("Pula", "Przychody puli", "2 Abonament premium", J, "sp"), ("Pula", "Przychody puli", "3 Finansowanie", J, "sf"),
        ("Pula", "Przychody puli", "4 Analizy", J, "sd"), ("Pula", "Przychody puli", "5 Opłata za sukces od GA", J, "ss"),
        ("Pula", "Koszty wspólne", "1 Partner Technologiczny", J, "ctp"), ("Pula", "Koszty wspólne", "2 Partner Marketplace", J, "cmp"), ("Pula", "Koszty wspólne", "3 Płatności", J, "cpay"),
        ("Pula", "Koszty wspólne", "4 Marketing", J, "cmkt"), ("Pula", "Koszty wspólne", "5 Obsługa klienta", J, "ccs"), ("Pula", "Koszty wspólne", "6 Biuro projektu", J, "cpm"),
        ("Pula", "Wynik i podział puli", "1 Wynik puli", J, "net"), ("Pula", "Wynik i podział puli", "2 Deficyt", J, "fund"), ("Pula", "Wynik i podział puli", "3 Zwrot deficytu", J, "rep"),
        ("Pula", "Wynik i podział puli", "4 Do podziału", J, "dist"),
        ("Partner Technologiczny", "Zarobek PT", "0 Wkład własny", T, "own"), ("Partner Technologiczny", "Zarobek PT", "2 Deficyt i zwrot", T, "fin"), ("Partner Technologiczny", "Zarobek PT", "3 Udział w podziale", T, "dist"),
        ("Partner Marketplace", "Zarobek PM", "0 Wkład własny", Mp, "own"), ("Partner Marketplace", "Zarobek PM", "2 Deficyt i zwrot", Mp, "fin"), ("Partner Marketplace", "Zarobek PM", "3 Udział w podziale", Mp, "dist"),
        ("Grupa Azoty", "Zarobek GA", "1 Dodatkowa marża zostająca w GA", G, "keep"), ("Grupa Azoty", "Zarobek GA", "2 Udział w podziale", G, "dist"), ("Grupa Azoty", "Zarobek GA", "3 Koszt wewnętrzny", G, "int"),
        ("Platforma", "GMV", "1 Nawozy GA", U, "fg"), ("Platforma", "GMV", "2 Inne produkty", U, "g3")]
TOC_TITLE = "GAia: model wspólnej puli przychodów (joint revenue) trzech partnerów"
TOC_SUB = "Grupa Azoty (produkt, dane, ARiMR, bez gotówki), Partner Technologiczny (silnik + UX), Partner Marketplace / Handel (marketplace przez plug-in). 3 scenariusze. Kwoty w mln zł."
TOC_ITEMS = [("Każdy partner zyskuje", "Wynik partnerów w wybranym scenariuszu, skala, wkład i finansowanie"), ("Scenariusze", "Ostrożny, bazowy, ambitny obok siebie"),
             ("Wkład i zasady", "Kto co wnosi, zasady wspólnej puli, etapy i bramki decyzyjne"), ("Założenia", "Założenia z benchmarkami; wybór scenariusza"),
             ("Klucz podziału", "Podział każdego strumienia przychodów między partnerów"), ("Wkład technologiczny", "Dopracowanie GAia w roku 0: moduły, kto robi, koszt"),
             ("Użytkownicy i GMV", "Użytkownicy → kupujący → GMV → zamówienia i tony"), ("Partner Technologiczny", "Wkład, koszty bieżące i wynik"),
             ("Partner Marketplace", "Wkład, koszty bieżące i wynik"), ("Grupa Azoty", "Dodatkowa marża, opłata za sukces i wynik"),
             ("Wspólna pula", "Przychody, koszty wspólne, deficyt, zwrot, podział"), ("Moc platformy", "Wskaźniki platformy"),
             ("Tabele przestawne", "Tabele przestawne na danych z arkusza Dane"), ("Dane", "Płaska tabela danych - źródło tabel przestawnych"),
             ("Case study i benchmarki", "Nutrien, Syngenta MAP, Starbucks × PepsiCo, Bosch × BASF, Bayer × Microsoft, Yara × Agrofy, sklepy z aplikacjami, Allegro")]
ORDER = ["Spis treści", "Każdy partner zyskuje", "Scenariusze", "Wkład i zasady", "Założenia", "Klucz podziału", "Wkład technologiczny", "Użytkownicy i GMV",
         "Partner Technologiczny", "Partner Marketplace", "Grupa Azoty", "Wspólna pula", "Moc platformy", "Dane", "Case study i benchmarki"]
def extra_finish():
    finish(P, REG_P, tab="008A3E"); P.cell(3, 1).font = Font(name=FONT, bold=True, color=NAVY_, size=12)
    finish(S3, [(4, 4 + len(SCEN_ROWS), 1, 5)], tab="008A3E")
    finish(Wk, REG_W, tab="6F9FD8")
    for r in (8, 16): Wk.cell(r, 1).font = Font(name=FONT, bold=True, color=NAVY_, size=11)
    for r in range(4, 7): Wk.row_dimensions[r].height = 60
    finish(K, [(4, 4 + len(KEYS), 1, 6)], tab="E67E00")
    for r in range(5, 5 + len(KEYS)): K.cell(r, 6).alignment = Alignment(wrap_text=True, vertical="center"); K.row_dimensions[r].height = 30
    finish(C, [(4, tot_r, 1, 9)], tab="E67E00")
    finish(Bm, [(3, Bm.max_row, 1, 6)], tab="6F9FD8")
    for r in range(4, Bm.max_row + 1):
        Bm.row_dimensions[r].height = 62; Bm.cell(r, 6).font = Font(name=FONT, size=9, color="0563C1", underline="single")
exec(open("jr_finish.py", encoding="utf-8").read())
wb.save(OUT); print("ok", OUT)
