# GAia v3: model współpracy trzech partnerów zbudowany na modelu P&L GA (GAia-model_PL_v2, "BC agressive").
# 7 lat, 3 scenariusze, synergie partnerstwa, wspólna pula przychodów, dopasowanie MVP1. Uruchom: python3 build_v3_xlsx.py
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.comments import Comment
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = "../GA_GAia_model_wspolpracy_v4.xlsx"
wb = openpyxl.Workbook()
FONT = "Trebuchet MS"; NAVY = "003665"; GRID = "BFBFBF"; LINK = "00704A"; GREYT = "595959"
BLUE = Font(name=FONT, color="0000FF", size=10); B = Font(name=FONT, bold=True, size=10)
HF = PatternFill("solid", fgColor=NAVY); INF = PatternFill("solid", fgColor="FFF2CC"); TOT = PatternFill("solid", fgColor="DDEBF7")
thin = Side(style="thin", color=GRID); nthin = Side(style="thin", color=NAVY); nmed = Side(style="medium", color=NAVY)
NUM = '#,##0.0;[Red]-#,##0.0;"–"'
NY = 7
YRS = [f"Rok {i}" for i in range(NY + 1)]
YC = [L(3 + i) for i in range(NY)]            # C..I = lata 1-7, B = rok 0
SUMC = L(3 + NY); NOTEC = 4 + NY                # J = suma, K = opis
COLS = ["B"] + YC

def hdr(ws, r, vals, c0=1):
    for i, v in enumerate(vals):
        c = ws.cell(r, c0 + i, v); c.font = Font(name=FONT, bold=True, color="FFFFFF", size=10); c.fill = HF
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left" if i == 0 else "center")
        c.border = Border(left=nthin, right=nthin, top=nthin, bottom=nthin)

# ======================= ZAŁOŻENIA =======================
A = wb.active; A.title = "Założenia"
A["A1"] = "GAia v4: założenia (model P&L GA \"BC agressive\" + uwagi z 04.11.2026)"
A["A2"] = "Niebieskie na żółtym = założenie. Kolumny lat dotyczą lat 1-7. Wartości w zł, jeśli nie podano inaczej (w arkuszach obliczeń: mln zł)."
N = {}; ar = 4
hdr(A, ar, ["Założenie", "Wartość / Rok 1"] + [f"Rok {i}" for i in range(2, NY + 1)] + ["Jedn.", "Źródło / uwaga"]); ar += 1
UC, SC_ = L(2 + NY), L(3 + NY)   # kolumny jednostki i źródła
def sec(t):
    global ar
    ar += 1; A.cell(ar, 1, t).font = Font(name=FONT, bold=True, color=NAVY, size=11); ar += 1
def inp(name, label, vals, unit, src, fmt="0.0"):
    global ar
    A.cell(ar, 1, label)
    vals = vals if isinstance(vals, list) else [vals]
    for i, v in enumerate(vals):
        c = A.cell(ar, 2 + i, v); c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE; c.fill = INF
    A.cell(ar, 2 + NY, unit); A.cell(ar, 3 + NY, src)
    N[name] = f"'Założenia'!$B${ar}" if len(vals) == 1 else [f"'Założenia'!${L(2+i)}${ar}" for i in range(NY)]
    ar += 1
def y(nm, i): v = N[nm]; return v[i] if isinstance(v, list) else v

sec("1. Rynek (model GA, sekcja 1-2)")
inp("farm", "Liczba aktywnych rolników w Polsce", 850000, "szt.", "Model GA, C6", "#,##0")
inp("area", "Łączna powierzchnia upraw", 14000000, "ha", "Model GA, C7", "#,##0")
inp("ha", "Średnia wielkość gospodarstwa", "=B{f}/B{f}", "ha", "Formuła: powierzchnia ÷ rolnicy", "0.0")
ha_r = ar - 1
inp("pf", "Wydatki na ha: nawozy mineralne", 1100, "zł/ha", "Model GA, C11 (przedział 600-2500)", "#,##0")
inp("ps", "Wydatki na ha: nasiona", 600, "zł/ha", "Model GA, C12", "#,##0")
inp("po", "Wydatki na ha: ŚOR", 700, "zł/ha", "Model GA, C13", "#,##0")
inp("pi", "Wydatki na ha: inne", 200, "zł/ha", "Model GA, C14", "#,##0")
inp("pu", "Wydatki na ha: ubezpieczenia / finanse", 125, "zł/ha", "Model GA, C15", "#,##0")
inp("pm", "Rynek maszyn (nowe 6 + używane 5 mld zł)", 11000000000, "zł/rok", "Model GA, I16", "#,##0")
inp("nag", "Wydatki nierolnicze przez platformę", 960, "zł/rolnik/rok", "Model GA, C18 (10% z 9 600 zł e-commerce)", "#,##0")
inp("spend", "Średni wydatek rolnika (z maszynami)", "=x", "zł/rok", "Formuła jak w modelu GA, C20", "#,##0")
sp_r = ar - 1
inp("fsh", "Udział nawozów w GMV", "=x", "%", "Formuła jak w modelu GA, B11", "0.0%")
fsh_r = ar - 1
inp("ccat", "Udział nawozów i maszyn w GMV (niższa prowizja)", 0.53, "%", "Model GA, wiersz 37", "0%")
inp("gas", "Udział nawozów GA w nawozach na platformie", 0.6, "%", "Model GA, B56", "0%")
sec("2. Ambicja (adopcja to funkcja budżetu) i udział w wydatkach")
inp("adopt", "AMBICJA: użytkownicy jako % aktywnych rolników", [0.05, 0.12, 0.20, 0.28, 0.35, 0.40, 0.43], "%", "Model GA, wiersz 25 (rok 7: ekstrapolacja). Liczba użytkowników to decyzja; jej koszt = CAC × nowi (arkusz Finansowanie)", "0%")
inp("tr", "Użytkownicy transakcyjni", [0.5, 0.6, 0.65, 0.7, 0.75, 0.75, 0.75], "%", "Model GA, wiersz 28", "0%")
inp("wal", "Udział wydatków rolnika w GAia", [0.05, 0.10, 0.15, 0.17, 0.19, 0.21, 0.22], "%", "Model GA, wiersz 31. Może być 20-30%: rolnik kupuje u 1-kilku dystrybutorów (uwaga 04.11)", "0%")
sec("3. Przychody (model GA, sekcja 3)")
inp("t1", "Prowizja: nawozy i maszyny (płaci sprzedawca / dystrybutor)", [0.01, 0.01, 0.011, 0.012, 0.013, 0.014, 0.015], "% GMV", "Model GA, wiersz 34. W MVP1 (rok 1): opłata od zrealizowanego zapytania", "0.0%")
inp("t2", "Prowizja: pozostałe kategorie", [0.03, 0.03, 0.035, 0.035, 0.04, 0.04, 0.04], "% GMV", "Model GA, wiersz 35 (Allegro: 12%)", "0.0%")
inp("bmA", "Back margin od dostawców: nawozy i maszyny", [0.01, 0.01, 0.011, 0.012, 0.013, 0.014, 0.015], "% GMV", "Model GA, wiersz 39. Także od nawozów (sprzedają dystrybutorzy, są też nawozy konkurencji). Można obniżyć jak prowizję", "0.0%")
inp("bmB", "Back margin od dostawców: pozostałe kategorie", [0.02, 0.02, 0.021, 0.021, 0.022, 0.022, 0.022], "% GMV", "Model GA, wiersz 39", "0.0%")
inp("pr", "Użytkownicy z usługą premium", [0, 0, 0.03, 0.05, 0.07, 0.09, 0.10], "%", "Model GA, wiersz 43", "0%")
inp("pha", "Cena premium", 20, "zł/ha/rok", "Model GA: 40 zł/ha (≈ 659 zł/rok); przyjęto 20 zł/ha (SatAgro 10-15 zł/ha). Do decyzji", "0")
sec("4. Korzyść Grupy Azoty (poza pulą)")
inp("inc", "Inkrementalna marża z danych rolników", [0, 0.005, 0.007, 0.009, 0.011, 0.012, 0.012], "% wartości nawozów GA", "Model GA, wiersz 59", "0.0%")
inp("sfee", "Opłata za sukces: część dodatkowej marży GA wpłacana do puli", 0.25, "%", "GA płaci tylko od zysku, którego bez platformy by nie miała", "0%")
inp("gaint", "Koszt wewnętrzny GA (opiekun, dane, ARiMR)", 300000, "zł/rok", "GA nie wnosi gotówki do puli", "#,##0")
sec("5. Koszty platformy (model GA, sekcja Koszty)")
inp("dev", "Rozwój platformy po starcie", [1500000, 2000000, 2000000, 3000000, 3000000, 3000000, 3000000], "zł/rok", "Model GA, wiersz 64 (rok 0 = arkusz Dopasowanie MVP1)", "#,##0")
inp("m0", "Utrzymanie: część stała", 1200000, "zł/rok", "Dopasowane do modelu GA (2,4 → 12 mln zł)", "#,##0")
inp("mu", "Utrzymanie: na użytkownika", 36, "zł/rok", "Model GA: 12 mln zł / 297,5 tys. w roku 5 ≈ 40 zł", "0")
inp("f0", "Zespół: część stała", 10, "FTE", "Dopasowane do modelu GA (15 → 80 FTE)", "0")
inp("fu", "Zespół: na każde 1000 użytkowników", 0.2, "FTE", "Model GA: 70 FTE przy 297,5 tys.", "0.00")
inp("fc", "Średni miesięczny koszt FTE", 25000, "zł", "Model GA, B78", "#,##0")
inp("cac", "CAC na nowego użytkownika (GŁÓWNE ZAŁOŻENIE)", [250, 250, 250, 250, 250, 250, 250], "zł", "Uwaga 04.11: koszt per user 250 zł (model GA: 250 → 350 zł). Wrażliwość: arkusz Finansowanie", "#,##0")
inp("churn", "Churn użytkowników", 0.15, "%", "Model GA, B72", "0%")
inp("ach", "Koszt antichurn na odchodzącego użytkownika", 350, "zł", "Model GA, B74", "#,##0")
sec("6. Synergie partnerstwa (DO DYSKUSJI: co partnerzy już mają i za co już płacą)")
inp("syt", "Mniejszy zespół: istniejące zespoły partnerów, dystrybutorzy GA odpowiadają na zapytania", 0.30, "%", "ZAŁOŻENIE: PT i PM obsługują też innych klientów; ofertę daje dystrybutor (MVP1)", "0%")
inp("sym", "Niższe utrzymanie: platforma PT współdzielona (multi-tenant)", 0.25, "%", "ZAŁOŻENIE do potwierdzenia ofertą PT", "0%")
inp("syc", "Niższy CAC dzięki kanałom GA", 0, "%", "0% zgodnie z uwagą 04.11: CAC rynkowy już zawiera wartość marki i kanałów", "0%")
sec("7. Podział kosztów i deficytu")
inp("ft", "Deficyt puli finansuje Partner Technologiczny", 0.7, "%", "GA 0%", "0%")
inp("fm", "Deficyt puli finansuje Partner Marketplace", "=x", "%", "Formuła: 100% − PT", "0%")
inp("wacc", "Stopa dyskontowa", 0.10, "%", "", "0%")
# formuły w założeniach
far = int(N["farm"].split("$")[-1]); arr = int(N["area"].split("$")[-1])
A.cell(ha_r, 2, f"=B{arr}/B{far}")
rr = {k: int(N[k].split("$")[-1]) for k in ("pf", "ps", "po", "pi", "pu", "pm", "nag")}
A.cell(sp_r, 2, f"=B{ha_r}*(B{rr['pf']}+B{rr['ps']}+B{rr['po']}+B{rr['pi']}+B{rr['pu']}+B{rr['pm']}/B{arr})+B{rr['nag']}")
A.cell(fsh_r, 2, f"=B{rr['pf']}/(B{rr['pf']}+B{rr['ps']}+B{rr['po']}+B{rr['pi']}+B{rr['pu']}+B{rr['pm']}/B{arr})")
A.cell(int(N["fm"].split("$")[-1]), 2, f"=1-B{int(N['ft'].split('$')[-1])}")
A.column_dimensions["A"].width = 70
for i in range(NY): A.column_dimensions[L(2 + i)].width = 11
A.column_dimensions[UC].width = 16; A.column_dimensions[SC_].width = 78
A.freeze_panes = "B5"

# ======================= KLUCZ PODZIAŁU =======================
K = wb.create_sheet("Klucz podziału")
K["A1"] = "Klucz podziału wspólnej puli: część każdego strumienia przychodów dla każdego partnera"
K["A2"] = "Każdy dostaje najwięcej tam, gdzie tworzy wartość. Benchmark: sklepy z aplikacjami 70/30, Salesforce AppExchange 15% dla platformy."
hdr(K, 4, ["Strumień przychodów", "Grupa Azoty", "Partner Technologiczny", "Partner Marketplace", "Suma", "Dlaczego tak"])
KEYS = [("comm", "Prowizje od sprzedawców i dystrybutorów", (0.20, 0.40, 0.40), "Ruch i klienci GA; aplikacja PT; marketplace i rozliczenia PM"),
        ("back", "Back margin od dostawców (akcje, promocje)", (0.30, 0.20, 0.50), "Zarządzanie kategoriami i akcjami (PM) + marka i kalendarz akcji GA (dożynki, sezon)"),
        ("sub", "Usługi premium (abonament)", (0.15, 0.70, 0.15), "Funkcje premium tworzy silnik PT (70% jak twórca aplikacji)"),
        ("succ", "Opłata za sukces od GA (25% dodatkowej marży)", (0.00, 0.60, 0.40), "GA dzieli się dodatkowym zyskiem z partnerami")]
KR = {}
for k_, (key, lab, sh_, why) in enumerate(KEYS):
    r = 5 + k_; K.cell(r, 1, lab)
    for j, v_ in enumerate(sh_):
        c = K.cell(r, 2 + j, v_); c.font = BLUE; c.fill = INF; c.number_format = "0%"
    K.cell(r, 5, f"=SUM(B{r}:D{r})").number_format = "0%"; K.cell(r, 6, why); KR[key] = r
K.column_dimensions["A"].width = 48
for c in "BCDE": K.column_dimensions[c].width = 14
K.column_dimensions["F"].width = 80
kref = lambda key, p: f"'Klucz podziału'!${'BCD'[p]}${KR[key]}"

# ======================= DOPASOWANIE MVP1 =======================
C = wb.create_sheet("Dopasowanie MVP1")
C["A1"] = "Dopasowanie wymagań MVP1 (OPZ, 01.06.2026) do platform partnerów: kto dostarcza i ile pracy w roku 0"
C["A2"] = "Poziom dopasowania i osobomiesiące to nasze założenie do potwierdzenia odpowiedziami SH na pytania PZ-01…PZ-49. Koszt = osobomiesiące × 20 tys. zł."
hdr(C, 4, ["ID", "Moduł MVP1", "Kto dostarcza", "Dopasowanie", "GA wnosi", "Osobo-\nmiesiące", "Koszt\nmln zł"])
PT_, PM_ = "Partner Technologiczny", "Partner Marketplace"
mvp = [("WB-01", "Rejestracja i onboarding", PT_, "standard w silniku", "", 2), ("WB-02", "Logowanie, MFA, Entra ID dla pracowników GA", PT_, "konfiguracja", "Entra ID GA", 4),
       ("WB-03", "Role i uprawnienia ABAC (10 ról)", PT_, "rozbudowa", "", 8), ("WB-04", "Kokpit rolnika", PT_, "nowa warstwa UX", "", 8),
       ("WB-05", "Profil rolnika", PT_, "standard w silniku", "", 3), ("WB-06", "Profil publiczny rolnika", PT_, "konfiguracja", "", 3),
       ("WB-07", "Paszport pola: mapy, GUGiK, Sentinel NDVI", PT_, "standard w silniku + integracje", "", 12), ("WB-08", "Dane ARiMR: Geoportal, plik GML z eWniosekPlus", PT_, "nowe integracje", "współpraca z ARiMR", 8),
       ("WB-09", "Dokumenty rolnika", PT_, "standard w silniku", "", 4), ("WB-10", "Alert 6:30: SMS, WhatsApp, e-mail, IMGW", PT_, "nowe", "", 8),
       ("WB-11", "Planer siewów i zbiorów", PT_, "standard w silniku", "", 4), ("WB-12", "Kalkulator zarobków (ceny rynkowe)", PT_, "nowe", "", 4),
       ("WB-13", "Harmonogram dotacji ARiMR", PT_, "nowe (CMS)", "terminy ARiMR", 2), ("WB-14", "Baza wiedzy (nawozy.eu, IUNG, IOR-PIB)", PT_, "konfiguracja CMS", "treści nawozy.eu", 4),
       ("WB-15", "Bot Norm COMPO EXPERT", PT_, "integracja podstawowa", "licencja i API", 3), ("WB-16", "Katalog produktów", PM_, "standard w marketplace (plug-in)", "produkty GA", 6),
       ("WB-17", "Zapytanie ofertowe", PM_, "konfiguracja marketplace", "", 8), ("WB-18", "Proces ofertowy dystrybutora", PM_, "konfiguracja marketplace", "66 dystrybutorów", 6),
       ("WB-19", "Leady dla dystrybutora", PM_, "rozbudowa", "", 6), ("WB-20", "Przypisanie rolnik - dystrybutor (regiony)", PT_, "konfiguracja", "mapa regionów", 3),
       ("WB-21", "Karta klienta (CRM dystrybutora)", PT_, "rozbudowa", "", 5), ("WB-22", "Wniosek finansowy (lead do banku)", PT_, "nowe + REST per bank", "", 6),
       ("WB-23", "Wybór instytucji finansowej, zgody RODO", PT_, "rozbudowa", "", 4), ("WB-24", "Oferta partnera finansowego", PM_, "konfiguracja marketplace", "", 3),
       ("WB-25", "Dostęp partnera finansowego", PT_, "konfiguracja", "", 2), ("WB-26", "Administracja platformą", PT_, "standard w silniku", "", 4),
       ("WB-27", "Audyt i zgodność (audit log)", PT_, "rozbudowa", "", 4), ("WB-28", "Raportowanie podstawowe", PT_, "standard w silniku", "", 4),
       ("+", "Grywalizacja (punkty, odznaki, wyzwania)", PT_, "nowe", "nagrody GA", 8), ("+", "Aplikacja mobilna (React Native, wybrane funkcje)", PT_, "nowa warstwa UX", "", 14),
       ("+", "Bezpieczeństwo i infrastruktura: Kubernetes, Vault, WAF, SIEM, pentest, DR", PT_, "konfiguracja", "", 16), ("+", "Testy, kierowanie projektem, wdrożenie", PT_, "-", "", 12)]
cr = 5
for rw in mvp:
    for j, v_ in enumerate(rw[:5]): C.cell(cr, 1 + j, v_)
    c = C.cell(cr, 6, rw[5]); c.font = BLUE; c.fill = INF
    C.cell(cr, 7, f"=F{cr}*20000/1000000").number_format = NUM; cr += 1
f1, l1 = 5, cr - 1
MVPR = {}
for who in (PT_, PM_):
    C.cell(cr, 1, ""); C.cell(cr, 2, f"Razem {who}").font = B; C.cell(cr, 3, who)
    C.cell(cr, 6, f'=SUMIF($C${f1}:$C${l1},$C{cr},F${f1}:F${l1})').font = B
    C.cell(cr, 7, f'=SUMIF($C${f1}:$C${l1},$C{cr},G${f1}:G${l1})').number_format = NUM; MVPR[who] = cr; cr += 1
C.cell(cr, 2, "WKŁAD TECHNOLOGICZNY W ROKU 0 RAZEM").font = B; C.cell(cr, 6, f"=F{MVPR[PT_]}+F{MVPR[PM_]}").font = B
C.cell(cr, 7, f"=G{MVPR[PT_]}+G{MVPR[PM_]}").number_format = NUM; mvp_tot = cr
cr += 2; C.cell(cr, 1, "Porównanie: model GA zakłada 4 mln zł na rozwój w roku 0 (wiersz 64).").font = Font(name=FONT, italic=True, color=GREYT)
for col, w in zip("ABCDEFG", [8, 52, 24, 30, 22, 11, 11]): C.column_dimensions[col].width = w
C.row_dimensions[4].height = 30
capPT = f"'Dopasowanie MVP1'!$G${MVPR[PT_]}"; capMP = f"'Dopasowanie MVP1'!$G${MVPR[PM_]}"

# ======================= silnik arkuszy obliczeń =======================
class Sheet:
    def __init__(self, title, head):
        self.ws = wb.create_sheet(title); self.r = 5; self.R = {}; self.notes = {}
        self.ws["A1"] = head
        hdr(self.ws, 4, ["Pozycja"] + YRS + ["Suma", "Jak liczone / skąd"])
        self.ws.column_dimensions["A"].width = 58
        for c in ["B"] + YC + [SUMC]: self.ws.column_dimensions[c].width = 10.5
        self.ws.column_dimensions[L(NOTEC)].width = 64; self.ws.freeze_panes = "B5"
    def row(self, key, label, f, y0=None, bold=False, fill=None, total=True, fmt=NUM, note=None):
        ws = self.ws; ws.cell(self.r, 1, label)
        if note: self.notes[key] = note
        if y0 is not None: ws.cell(self.r, 2, "=" + y0)
        for i, col in enumerate(YC): ws[f"{col}{self.r}"] = "=" + f(i, col)
        if total: ws[f"{SUMC}{self.r}"] = f"=SUM(B{self.r}:{YC[-1]}{self.r})"
        for c in range(1, NOTEC):
            ce = ws.cell(self.r, c); ce.number_format = fmt
            if bold: ce.font = B
            if fill: ce.fill = fill
        self.R[key] = self.r; self.r += 1
        return self.r - 1
    def sec(self, t): self.r += 1; self.ws.cell(self.r, 1, t).font = Font(name=FONT, bold=True, color=NAVY, size=11); self.r += 1
    def ref(self, key, col): return f"'{self.ws.title}'!{col}{self.R[key]}"
    def cum(self, key, label):
        src = self.R[key]; rr_ = self.row(key + "_cum", label, lambda i, c: "0", y0="0", total=False, note="Suma narastająca")
        self.ws[f"B{rr_}"] = f"=B{src}"
        for i, col in enumerate(YC): self.ws[f"{col}{rr_}"] = f"={'B' if i == 0 else YC[i-1]}{rr_}+{col}{src}"
M = 1000000


# ---------- Rynek i GMV ----------
U = Sheet("Rynek i GMV", "Użytkownicy, wydatki i GMV (ambicja z Założeń), mln zł")
U.sec("Użytkownicy")
U.row("u", "Użytkownicy (średnio w roku, tys.)", lambda i, c: f"{N['farm']}*{y('adopt',i)}/1000", total=False, fmt="#,##0", note="Rolnicy × ambicja (% rolników); model GA wiersz 26")
U.row("nu", "Nowi użytkownicy (tys.)", lambda i, c: f"MAX(0,{c}{U.R['u']}-{'0' if i == 0 else YC[i-1] + str(U.R['u'])})", fmt="#,##0", note="Przyrost; w roku 1 wszyscy są nowi (jak model GA)")
U.row("ut", "Użytkownicy transakcyjni (tys.)", lambda i, c: f"{c}{U.R['u']}*{y('tr',i)}", total=False, fmt="#,##0", note="Użytkownicy × % transakcyjnych")
U.sec("GMV")
U.row("gmv", "GMV razem", lambda i, c: f"{N['spend']}*{y('wal',i)}*{c}{U.R['ut']}*1000/{M}", bold=True, fill=TOT, note="Wydatek rolnika × udział w GAia × transakcyjni (model GA wiersz 32)")
U.row("fert", "w tym nawozy", lambda i, c: f"{c}{U.R['gmv']}*{N['fsh']}", note="GMV × udział nawozów (31,3%)")
U.row("fga", "w tym nawozy Grupy Azoty", lambda i, c: f"{c}{U.R['fert']}*{N['gas']}", note="Nawozy × 60% (model GA wiersz 57)")
U.row("pen", "Nawozy GA przez platformę / sprzedaż krajowa GA (3,69 mld zł)", lambda i, c: f"{c}{U.R['fga']}/(2.93*0.7*1800)", total=False, fmt="0.0%", note="Kontrola realności skali")
u = lambda k, c: U.ref(k, c)

# ---------- Koszty platformy ----------
Kp = Sheet("Koszty platformy", "Koszty platformy: model GA (samodzielnie) i w partnerstwie (z synergiami), mln zł")
Kp.sec("A. Jak w modelu GA (gdyby jedna firma budowała sama)")
Kp.row("dev", "Rozwój platformy", lambda i, c: f"{y('dev',i)}/{M}", note="Model GA wiersz 64; rok 0 jako wkład partnerów (arkusz Dopasowanie MVP1)")
Kp.row("mt", "Utrzymanie (usługi obce, SaaS)", lambda i, c: f"({N['m0']}+{N['mu']}*{u('u',c)}*1000)/{M}", note="1,2 mln zł + 36 zł × użytkownik")
Kp.row("tm", "Zespół (IT, produkt, obsługa, zakupy)", lambda i, c: f"({N['f0']}+{N['fu']}*{u('u',c)})*{N['fc']}*12/{M}", note="(10 + 0,2 FTE / 1000 użytk.) × 25 tys. zł × 12")
Kp.row("cac", "Pozyskanie użytkowników (CAC)", lambda i, c: f"{u('nu',c)}*1000*{y('cac',i)}/{M}", note="Nowi użytkownicy × CAC (model GA wiersz 70)")
Kp.row("ach", "Antichurn", lambda i, c: "0" if i == 0 else f"'Rynek i GMV'!{YC[i-1]}{U.R['u']}*1000*{N['churn']}*{N['ach']}/{M}", note="Użytkownicy rok wcześniej × 15% × 350 zł")
Kp.row("tot", "KOSZTY RAZEM (model samodzielny)", lambda i, c: f"SUM({c}{Kp.R['dev']}:{c}{Kp.R['ach']})", bold=True, note="Suma")
Kp.sec("B. W partnerstwie (synergie z arkusza Założenia, sekcja 6)")
Kp.row("pdev", "Rozwój platformy (PT)", lambda i, c: f"{c}{Kp.R['dev']}", note="Bez zmian")
Kp.row("pmt", "Utrzymanie (PT 80% / PM 20%)", lambda i, c: f"{c}{Kp.R['mt']}*(1-{N['sym']})", note="−25%: platforma PT współdzielona")
Kp.row("ptm", "Zespół (PT 50% / PM 30% / obsługa 20%)", lambda i, c: f"{c}{Kp.R['tm']}*(1-{N['syt']})", note="−30%: istniejące zespoły, dystrybutorzy GA odpowiadają na zapytania")
Kp.row("pcac", "Pozyskanie i utrzymanie użytkowników", lambda i, c: f"({c}{Kp.R['cac']}+{c}{Kp.R['ach']})*(1-{N['syc']})", note="Bez synergii (uwaga 04.11): CAC jak w modelu GA")
Kp.row("ptot", "KOSZTY WSPÓLNE PULI (w partnerstwie)", lambda i, c: f"SUM({c}{Kp.R['pdev']}:{c}{Kp.R['pcac']})", bold=True, fill=TOT, note="Pokrywa je pula, po koszcie, bez marży")
Kp.row("syn", "Oszczędność dzięki partnerstwu", lambda i, c: f"{c}{Kp.R['tot']}-{c}{Kp.R['ptot']}", note="Model samodzielny − partnerstwo")

# ---------- Wspólna pula ----------
J = Sheet("Wspólna pula", "Wspólna pula: przychody → koszty po koszcie → zwrot deficytu → podział według klucza, mln zł")
J.sec("1. Przychody puli")
J.row("comm", "Prowizje od sprzedawców i dystrybutorów", lambda i, c: f"{u('gmv',c)}*{N['ccat']}*{y('t1',i)}+{u('gmv',c)}*(1-{N['ccat']})*{y('t2',i)}", note="Model GA wiersz 37: 53% GMV × 1-1,5% + 47% × 3-4%")
J.row("back", "Back margin od dostawców", lambda i, c: f"{u('gmv',c)}*{N['ccat']}*{y('bmA',i)}+{u('gmv',c)}*(1-{N['ccat']})*{y('bmB',i)}", note="Model GA wiersz 41: całe GMV (także nawozy); stawki osobno dla nawozów i maszyn")
J.row("sub", "Usługi premium", lambda i, c: f"{u('u',c)}*1000*{y('pr',i)}*{N['pha']}*{N['ha']}/{M}", note="Użytkownicy × % premium × cena za ha × 16,5 ha")
J.row("succ", "Opłata za sukces od GA", lambda i, c: f"'Grupa Azoty'!{c}{{EXTRA}}*{N['sfee']}", note="25% dodatkowej marży GA")
J.row("pool", "PRZYCHODY PULI RAZEM", lambda i, c: f"SUM({c}{J.R['comm']}:{c}{J.R['succ']})", bold=True, fill=TOT, note="Suma")
J.row("pct", "Przychody puli / GMV", lambda i, c: f"IF({u('gmv',c)}=0,0,{c}{J.R['pool']}/{u('gmv',c)})", total=False, fmt="0.0%", note="Model GA: 3,9-5,5%; Allegro ok. 11-14%")
J.sec("2. Koszty wspólne (arkusz Koszty platformy, część B)")
J.row("cost", "Koszty wspólne puli", lambda i, c: Kp.ref("ptot", c), y0="0", note="Po koszcie, bez marży")
J.row("net", "WYNIK PULI (EBITDA platformy)", lambda i, c: f"{c}{J.R['pool']}-{c}{J.R['cost']}", y0="0", bold=True, fill=TOT, note="Przychody − koszty")
J.cum("net", "Wynik puli narastająco")
J.sec("3. Deficyt i jego zwrot (GA nie dopłaca)")
J.row("fund", "Deficyt (finansują PT 70% i PM 30%)", lambda i, c: f"MAX(0,-{c}{J.R['net']})", y0="0", note="Gdy wynik < 0")
J.row("bal", "Deficyt do zwrotu (saldo)", lambda i, c: "0", y0="0", total=False, note="Poprzedni rok + deficyt − zwrot")
J.row("rep", "Zwrot deficytu (pierwszeństwo)", lambda i, c: "0", y0="0", note="Dodatni wynik najpierw spłaca deficyt")
J.row("dist", "DO PODZIAŁU WEDŁUG KLUCZA", lambda i, c: "0", y0="0", bold=True, note="Po zwrocie deficytu")
br_, rp_, ds_, nt_ = J.R["bal"], J.R["rep"], J.R["dist"], J.R["net"]
for i, col in enumerate(YC):
    prev = "B" if i == 0 else YC[i - 1]
    J.ws[f"{col}{rp_}"] = f"=MIN({prev}{br_},MAX(0,{col}{nt_}))"; J.ws[f"{col}{br_}"] = f"={prev}{br_}+{col}{J.R['fund']}-{col}{rp_}"
    J.ws[f"{col}{ds_}"] = f"=MAX(0,{col}{nt_})-{col}{rp_}"
J.sec("4. Udział partnerów w przychodach puli (klucz)")
strm = ["comm", "back", "sub", "succ"]
for p, (pk, pn) in enumerate([("gGA", "Grupa Azoty"), ("gPT", "Partner Technologiczny"), ("gPM", "Partner Marketplace")]):
    J.row(pk, f"Przychody przypisane: {pn}", lambda i, c, p=p: "+".join(f"{c}{J.R[s]}*{kref(s, p)}" for s in strm), note="Σ strumień × klucz")
for pk, wk, pn in [("gGA", "wGA", "Grupa Azoty"), ("gPT", "wPT", "Partner Technologiczny"), ("gPM", "wPM", "Partner Marketplace")]:
    J.row(wk, f"Udział w puli: {pn}", lambda i, c, pk=pk: f"IF({c}{J.R['pool']}=0,0,{c}{J.R[pk]}/{c}{J.R['pool']})", total=False, fmt="0%", note="Przypisane ÷ pula")

# ---------- Grupa Azoty ----------
G = Sheet("Grupa Azoty", "Grupa Azoty: korzyść z platformy (bez wkładu pieniężnego), mln zł")
G.sec("A. Dodatkowa marża GA")
G.row("dm", "Marża z danych rolników (model GA)", lambda i, c: f"{u('fga',c)}*{y('inc',i)}", note="Nawozy GA × 0-1,2% (model GA wiersz 60)")
G.row("extra", "DODATKOWA MARŻA GA", lambda i, c: f"{c}{G.R['dm']}", bold=True, note="Tylko poprawa mixu (model GA); bez dodatkowej sprzedaży - ograniczenia podażowe (uwaga 04.11)")
G.sec("B. Wynik Grupy Azoty")
G.row("keep", "1. Dodatkowa marża zostająca w GA (75%)", lambda i, c: f"{c}{G.R['extra']}-{J.ref('succ',c)}", note="Dodatkowa marża − opłata za sukces")
G.row("dist", "2. Udział w podziale puli", lambda i, c: f"{J.ref('dist',c)}*{J.ref('wGA',c)}", y0="0", note="Do podziału × udział GA")
G.row("int", "3. Koszt wewnętrzny GA", lambda i, c: f"-{N['gaint']}/{M}", note="Opiekun, dane, ARiMR")
G.row("cash", "Wkład pieniężny GA", lambda i, c: "0", y0="0", note="GA nie wnosi gotówki ani nie dopłaca deficytu")
G.row("res", "WYNIK GRUPY AZOTY", lambda i, c: f"SUM({c}{G.R['keep']}:{c}{G.R['cash']})", y0=f"SUM(B{G.R['keep']}:B{G.R['cash']})", bold=True, fill=TOT, note="Suma")
G.cum("res", "Wynik skumulowany")
for i, col in enumerate(YC):  # podmiana odwołania do dodatkowej marży w puli
    J.ws[f"{col}{J.R['succ']}"] = f"='Grupa Azoty'!{col}{G.R['extra']}*{N['sfee']}"

# ---------- Partnerzy ----------
def partner(title, head, capf, share_key, fin_key, wkey, cost_lines):
    P_ = Sheet(title, head)
    P_.sec("A. Wkład technologiczny w roku 0 (własne zasoby)")
    P_.row("cap", "Dopasowanie MVP1 (arkusz Dopasowanie MVP1)", lambda i, c: "0", y0=capf, note="Osobomiesiące × 20 tys. zł")
    P_.sec("B. Koszty bieżące pokrywane przez pulę (po koszcie)")
    for key, lab, f_, nt in cost_lines: P_.row(key, lab, f_, note=nt)
    P_.sec("C. Z czego i kiedy zarabia")
    P_.row("own", "0. Wkład własny (rok 0)", lambda i, c: "0", y0=f"-{capf}", note="Minus praca w roku 0")
    P_.row("fin", "1. Dopłaty do deficytu puli i ich zwrot", lambda i, c: f"(-{J.ref('fund',c)}+{J.ref('rep',c)})*{N[fin_key]}", y0="0", note="Minus dopłata, plus zwrot")
    P_.row("dist", "2. Udział w podziale puli", lambda i, c: f"{J.ref('dist',c)}*{J.ref(wkey,c)}", y0="0", note="Do podziału × udział w puli")
    P_.row("res", "WYNIK PARTNERA", lambda i, c: f"SUM({c}{P_.R['own']}:{c}{P_.R['dist']})", y0=f"SUM(B{P_.R['own']}:B{P_.R['dist']})", bold=True, fill=TOT, note="Suma")
    P_.cum("res", "Wynik skumulowany")
    return P_
T = partner("Partner Technologiczny", "Partner Technologiczny (SH, silnik + UX): wkład, koszty bieżące i wynik, mln zł", capPT, "tpsh", "ft", "wPT",
            [("c1", "Rozwój platformy (100%)", lambda i, c: Kp.ref("pdev", c), "Pokrywa pula"), ("c2", "Utrzymanie (80%)", lambda i, c: f"{Kp.ref('pmt', c)}*0.8", "Pokrywa pula"),
             ("c3", "Zespół IT i produkt (50%)", lambda i, c: f"{Kp.ref('ptm', c)}*0.5", "Pokrywa pula")])
Mp = partner("Partner Marketplace", "Partner Marketplace / Handel: wkład, koszty bieżące i wynik, mln zł", capMP, "mpsh", "fm", "wPM",
             [("c1", "Utrzymanie marketplace'u (20%)", lambda i, c: f"{Kp.ref('pmt', c)}*0.2", "Pokrywa pula"), ("c2", "Zespół zakupów, kategorii, sprzedawców (30%)", lambda i, c: f"{Kp.ref('ptm', c)}*0.3", "Pokrywa pula")])

# ---------- Podsumowanie ----------
P = wb.create_sheet("Każdy partner zyskuje", 0)
P["A1"] = "GAia v4: model współpracy na bazie P&L GA (po uwagach z 04.11)"
P["A2"] = "Gotówka, mln zł. Ambicja = ścieżka adopcji z modelu GA; budżet = nowi użytkownicy × CAC. Ile trzeba sfinansować i jak to zależy od CAC: arkusz Finansowanie."
P.cell(3, 1, "=\"Ambicja: \"&TEXT(" + N['adopt'][4] + ",\"0%\")&\" rolników w roku 5, CAC \"&" + N['cac'][0] + "&\" zł\"")
hdr(P, 4, ["Partner"] + YRS + ["Suma 0-7", "NPV (10%)", "Zwrot od", "Wynik w roku 2 (bramka)"])
parts = [("Grupa Azoty", G, "res"), ("Partner Technologiczny", T, "res"), ("Partner Marketplace / Handel", Mp, "res"), ("Wspólna pula: EBITDA platformy", J, "net")]
lastc = YC[-1]
for k, (name, sh, key) in enumerate(parts):
    pr = 5 + k; P.cell(pr, 1, name)
    for j, col in enumerate(COLS): P.cell(pr, 2 + j, f"={sh.ref(key, col)}")
    P.cell(pr, 3 + NY, f"=SUM(B{pr}:{L(2+NY)}{pr})"); P.cell(pr, 4 + NY, f"=B{pr}+NPV({N['wacc']},C{pr}:{L(2+NY)}{pr})")
    cond = ",".join([f"IF(SUM(B{pr}:{L(2+j)}{pr})>0,\"rok {j}\"" for j in range(1, NY + 1)])
    P.cell(pr, 5 + NY, "=" + cond + ",\"po roku 7\"" + ")" * NY)
    P.cell(pr, 6 + NY, f"=SUM(B{pr}:D{pr})")
    for c in list(range(2, 5 + NY)) + [6 + NY]: P.cell(pr, c).number_format = NUM
SK = 10
P.cell(SK, 1, "Skala")
for j in range(NY): P.cell(SK, 3 + j, f"Rok {j+1}")
kpis = [("Użytkownicy (tys.)", U, "u", "#,##0"), ("GMV (mln zł)", U, "gmv", NUM), ("Nawozy GA przez platformę / sprzedaż krajowa GA", U, "pen", "0.0%"),
        ("Przychody puli (mln zł)", J, "pool", NUM), ("Koszty wspólne (mln zł)", J, "cost", NUM), ("EBITDA platformy (mln zł)", J, "net", NUM),
        ("Budżet pozyskania i utrzymania użytkowników (mln zł)", Kp, "pcac", NUM), ("Oszczędność dzięki partnerstwu - do dyskusji (mln zł)", Kp, "syn", NUM)]
for k, (lab, sh, key, fm) in enumerate(kpis):
    pr = SK + 1 + k; P.cell(pr, 1, lab)
    for j, col in enumerate(YC): P.cell(pr, 3 + j, f"={sh.ref(key, col)}").number_format = fm
IV = SK + len(kpis) + 2
P.cell(IV, 1, "Wkład i finansowanie"); P.cell(IV, 3, "mln zł")
inv = [("Wkład technologiczny PT w roku 0 (Dopasowanie MVP1)", capPT), ("Wkład technologiczny PM w roku 0", capMP), ("Wkład pieniężny Grupy Azoty", "0"),
       ("Maks. deficyt puli narastająco (PT 70% / PM 30%)", f"-MIN(0,MIN('Wspólna pula'!B{J.R['net_cum']}:{lastc}{J.R['net_cum']}))")]
for k, (lab, f) in enumerate(inv): P.cell(IV + 1 + k, 1, lab); P.cell(IV + 1 + k, 3, "=" + f).number_format = NUM
P.column_dimensions["A"].width = 52
for c in range(2, 7 + NY): P.column_dimensions[L(c)].width = 11
REG_P = [(4, 8, 1, 6 + NY), (SK, SK + len(kpis), 1, 2 + NY), (IV, IV + len(inv), 1, 3)]

S3 = wb.create_sheet("Finansowanie", 1)
S3["A1"] = "Ile trzeba sfinansować: ambicja i koszt pozyskania rolnika (CAC) decydują o budżecie"
S3["A2"] = "Wartości przeliczone skryptem add_financing_v4.py. Finansowanie = najniższy skumulowany wynik puli (deficyt do pokrycia przez partnerów lub granty)."
FIN_COLS = ["Wariant", "Użytkownicy w roku 5 (tys.)", "Wydatki na pozyskanie i utrzymanie 0-7 (mln zł)", "Potrzebne finansowanie (mln zł)", "EBITDA platformy na plusie od", "Wynik puli narastająco na plusie od", "GA: 7 lat", "PT: 7 lat", "PM: 7 lat"]
FIN_T1 = 4; FIN_T2 = 13; FIN_T3 = 21
S3.cell(FIN_T1, 1, "1. Wrażliwość na CAC (ambicja jak w modelu GA)")
hdr(S3, FIN_T1 + 1, FIN_COLS)
S3.cell(FIN_T2, 1, "2. Ambicja: ilu rolników chcemy w roku 5 (CAC jak w modelu GA)")
hdr(S3, FIN_T2 + 1, FIN_COLS)
S3.cell(FIN_T3, 1, "3. Synergie (do dyskusji): wpływ na finansowanie")
hdr(S3, FIN_T3 + 1, FIN_COLS)
S3.column_dimensions["A"].width = 46
for c in "BCDEFGHI": S3.column_dimensions[c].width = 16
SCEN_ROWS = []
exec(open("v4_texts.py", encoding="utf-8").read())
exec(open("v4_finish.py", encoding="utf-8").read())
wb.save(OUT); print("ok", OUT)
