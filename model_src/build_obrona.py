# Skompresowany pakiet do obrony: 1 Excel (wyniki, źródła założeń, GMV, premium, churn, wrażliwość) z partner_variants_v5.json.
# Użycie: python3 build_obrona.py   → ../GA_GAia_obrona_zalozen.xlsx
import json, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
V = json.load(open("partner_variants_v5.json", encoding="utf-8")); D = V["deal"]
OUT = "../GA_GAia_obrona_zalozen.xlsx"
FONT = "Trebuchet MS"; NAVY = "003665"; RED = "FF0000"
thin = Side(style="thin", color="A6A6A6"); BRD = Border(left=thin, right=thin, top=thin, bottom=thin)
HF = PatternFill("solid", fgColor=NAVY); IN = PatternFill("solid", fgColor="FFF2CC"); TOT = PatternFill("solid", fgColor="DDEBF7")
NUM = '#,##0.0;[Red]-#,##0.0;"–"'
YRS = [f"{2027+i}\n(rok {i})" for i in range(8)]
wb = openpyxl.Workbook(); wb.remove(wb.active)

def sheet(name, title, sub):
    ws = wb.create_sheet(name); ws.sheet_view.showGridLines = False
    ws["A1"] = title; ws["A1"].font = Font(name=FONT, bold=True, size=14, color=NAVY)
    ws["A2"] = sub; ws["A2"].font = Font(name=FONT, italic=True, size=9, color="595959")
    return ws
def hdr(ws, r, vals, c0=1):
    for j, v in enumerate(vals):
        c = ws.cell(r, c0 + j, v); c.font = Font(name=FONT, bold=True, color="FFFFFF", size=10); c.fill = HF
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center"); c.border = BRD
    ws.row_dimensions[r].height = 32
def put(ws, r, c, v, fmt=None, bold=False, fill=None, inp=False, wrap=False, color=None):
    x = ws.cell(r, c, v); x.font = Font(name=FONT, size=10, bold=bold, color=("0000FF" if inp else (color or "000000"))); x.border = BRD
    if fmt: x.number_format = fmt
    if fill or inp: x.fill = fill or IN
    x.alignment = Alignment(wrap_text=wrap, vertical="center")
    return x
L = openpyxl.utils.get_column_letter

# ---------- 1. Kto i kiedy zarabia ----------
ws = sheet("1 Kto i kiedy zarabia", "Kto i kiedy zarabia: gotówka, mln zł (warunki proponowane: koszt + 10% marży, finansowanie deficytu PT 45% / PM 55%, GA 0)",
           "Źródło: GA_GAia_model_wspolpracy_v5.xlsx (P&L GA z korektami: CAC 250 → 350 zł, bez synergii, back margin bez nawozów GA, premium 20 zł/ha, bez sprzedaży inkrementalnej). Rok 0 = 2027 (MVP1 + pilot).")
hdr(ws, 4, ["Partner"] + YRS + ["Suma 2027-2034", "Zwrot (rok kalend.)", "IRR", "Maks. zaangażowanie"])
parts = [("Partner Technologiczny (SH)", "PT"), ("Partner Marketplace / Handel", "PM"), ("Grupa Azoty", "GA")]
for k, (nm, key) in enumerate(parts):
    r = 5 + k; d = D[key]; put(ws, r, 1, nm, bold=True)
    for j, v in enumerate(d["res"]): put(ws, r, 2 + j, round(v, 2), NUM)
    put(ws, r, 10, f"=SUM(B{r}:I{r})", NUM, bold=True, fill=TOT)
    put(ws, r, 11, 2027 + d["payback"] if d["payback"] is not None else "po 2034", "0")
    put(ws, r, 12, "n/d (brak wkładu)" if key == "GA" else d["irr7"], "0%")
    put(ws, r, 13, round(d["trough"], 2), NUM)
r = 9; put(ws, r, 1, "Narastająco", bold=True)
for k, (nm, key) in enumerate(parts):
    rr = r + 1 + k; put(ws, rr, 1, nm)
    for j in range(8): put(ws, rr, 2 + j, f"=SUM($B{5+k}:{L(2+j)}{5+k})", NUM)
hdr(ws, 14, ["Wspólna pula platformy"] + YRS + ["Suma"])
pool = [("Przychody puli", "rev"), ("Koszty wspólne (z marżą partnerów)", "cost"), ("EBITDA platformy", "net"), ("Do podziału (po zwrocie deficytu)", "dist")]
for k, (lab, key) in enumerate(pool):
    rr = 15 + k; put(ws, rr, 1, lab, bold=(key == "net"))
    for j, v in enumerate(D["pool"][key]): put(ws, rr, 2 + j, round(v, 2), NUM)
    put(ws, rr, 10, f"=SUM(B{rr}:I{rr})", NUM, bold=True, fill=TOT)
put(ws, 19, 1, "Potrzebne finansowanie startu (najgłębszy deficyt puli)", bold=True); put(ws, 19, 10, round(D["pool"]["need"], 2), NUM, bold=True, fill=TOT)
txt = ["JAK ZARABIA PT: 1) pula płaci za pracę zespołu (koszt + 10% marży, 12,6 mln zł marży w 8 lat) od 2028; 2) finansuje 45% deficytu 2028-2030, który wraca jako pierwszy 2031-2033; 3) od 2033 dostaje ok. 36% zysku do podziału (40% prowizji, 20% back marginu, 70% premium).",
       "JAK ZARABIA PM: jak PT, finansuje 55% deficytu; od 2033 ok. 42% zysku (40% prowizji, 50% back marginu, 15% premium, 40% opłaty za sukces).",
       "JAK ZARABIA GA: bez gotówki. Od 2028 marża z lepszego mixu nawozów (0,5-1,2% wartości nawozów GA na platformie), minus 25% tej marży wpłacane do puli; od 2033 ok. 22% zysku do podziału (20% prowizji, 30% back marginu, 15% premium)."]
for k, t in enumerate(txt):
    c = ws.cell(21 + k, 1, t); c.font = Font(name=FONT, size=10, color=NAVY, bold=(k == 0)); c.alignment = Alignment(wrap_text=True, vertical="top"); ws.merge_cells(start_row=21 + k, start_column=1, end_row=21 + k, end_column=13); ws.row_dimensions[21 + k].height = 32
ws.column_dimensions["A"].width = 46
for c in range(2, 14): ws.column_dimensions[L(c)].width = 11.5
ws.freeze_panes = "B5"

# ---------- 2. Założenia i źródła ----------
ws = sheet("2 Założenia i źródła", "Każde założenie: wartość, skąd, benchmark, co jeśli inaczej", "Kolor źródła: GA = Wasz model GAia-model_PL_v2 (BC agressive); BENCHMARK = dane rynkowe; ZAŁOŻENIE = nasze, do weryfikacji w pilocie.")
hdr(ws, 4, ["Obszar", "Założenie", "Wartość", "Typ źródła", "Źródło / benchmark", "Dlaczego tak", "Co jeśli inaczej (PT, suma 8 lat)"])
A = [
 ("Rynek", "Aktywni rolnicy", "850 tys.", "GA", "Model GA, C6", "Rolnicy aktywnie kupujący środki produkcji", "–"),
 ("Rynek", "Powierzchnia / gospodarstwo", "14 mln ha / 16,5 ha", "GA", "Model GA, C7", "Powierzchnia upraw w PL", "–"),
 ("Rynek", "Wydatek rolnika rocznie", "58 784 zł", "GA", "Model GA, C20: (1100 nawozy + 600 nasiona + 700 ŚOR + 200 inne + 125 ubezp. + 786 maszyny) zł/ha × 16,5 ha + 960 zł nierolnicze", "Maszyny = rynek 11 mld zł (6 nowe + 5 używane) ÷ 14 mln ha", "Bez maszyn: 45 842 zł (wydatki −22%)"),
 ("Adopcja", "Użytkownicy (% rolników)", "5% → 35% w 2032 (297,5 tys.)", "GA", "Model GA, wiersz 25", "Ambicja GA; jej koszt liczony przez CAC (uwaga 04.11: adopcja = funkcja budżetu)", f"20% w 2032: PT {V['deal_adopt20']['PT']['sum7']:.1f} mln zł"),
 ("Adopcja", "Użytkownicy transakcyjni", "50% → 75%", "GA", "Model GA, wiersz 28", "Rekrutacja do platformy przez oferty", "–"),
 ("GMV", "Udział wydatków rolnika w GAia", "5% → 19% (2032)", "GA + BENCHMARK", "Model GA, wiersz 31; GA: top kategorie e-commerce 20-24% (fashion 24%, książki 22%, RTV 20%)", "Rolnik kupuje u 1-2 dystrybutorów; zapytanie ofertowe przez GAia", f"−30%: PT {V['deal_wal70']['PT']['sum7']:.1f} mln zł (RYZYKO KRYTYCZNE)"),
 ("Przychody", "Prowizja nawozy i maszyny", "1,0% → 1,5% GMV", "GA + BENCHMARK", "Model GA, wiersz 34; Allegro ok. 12% (GA)", "Kategorie wrażliwe cenowo; płaci sprzedawca / dystrybutor", "–"),
 ("Przychody", "Prowizja pozostałe", "3,0% → 4,0% GMV", "GA", "Model GA, wiersz 35", "Poniżej standardu e-commerce, buduje atrakcyjność", "–"),
 ("Przychody", "Back margin dostawców", "nawozy (bez GA) i maszyny 1-1,5%; pozostałe 2-2,2%", "GA + ZAŁOŻENIE", "Model GA, wiersz 39 (2-2,2%); bez nawozów GA, bo GA nie płaci sama sobie", "Akcje promocyjne: dożynki, black week (GA)", "–"),
 ("Przychody", "Premium", "20 zł/ha (≈ 330 zł/rok), 3% → 9% użytkowników", "BENCHMARK", "Model GA: 40 zł/ha; SatAgro (satelitarne mapy) ok. 10-15 zł/ha; ekoschemat plan nawożenia ok. 87,52 zł/ha dopłaty (2025)", "Patrz arkusz 4", "40 zł/ha jak GA: przychody premium ×2"),
 ("GA", "Marża z lepszego mixu", "0,5% → 1,2% wartości nawozów GA", "GA", "Model GA, wiersz 59", "Bez sprzedaży inkrementalnej (uwaga 04.11)", "–"),
 ("Koszty", "Rozwój po starcie", "1,5 → 3 mln zł/rok", "GA", "Model GA, wiersz 64", "Programiści w zespole; to rozwój ponad zespół", "–"),
 ("Koszty", "Utrzymanie (usługi obce, SaaS)", "1,2 mln zł + 36 zł/użytk. (2,7 → 13 mln zł)", "GA + BENCHMARK", "Model GA: 2,4 → 12 mln zł; GA: typowo 5-12% przychodów", "W 2032: 11% przychodów", "–"),
 ("Koszty", "Zespół", "10 FTE + 0,2 FTE/1000 użytk. × 25 tys. zł (18 → 70 FTE)", "GA", "Model GA: 15 → 80 FTE, B78 25 tys. zł", "IT + produkt + obsługa + zakupy (benchmark e-commerce, GA)", "–"),
 ("Koszty", "CAC (koszt pozyskania)", "250 → 350 zł / nowy użytkownik", "GA", "Model GA, wiersz 69: rośnie ze skalą, 'nie znamy dla rolników'", "Kolejni rolnicy są drożsi; bez synergii z kanałami GA (uwaga 04.11)", f"+100 zł: PT {V['deal_cac100']['PT']['sum7']:.1f} mln zł"),
 ("Koszty", "Churn i antichurn", "15% użytkowników/rok × 350 zł", "GA", "Model GA, B72-B74: 'akceptowalny dla aplikacji B2B'", "Patrz arkusz 5", "–"),
 ("Koszty", "MVP1 (2027, wkład partnerów)", "188 osobomies. × 25 tys. zł × 1,2 = 5,6 mln zł", "ZAŁOŻENIE", "OPZ MVP1 01.06.2026: 28 wymagań, 16 integracji; model GA: 4 mln zł capex", "Gotowe silniki PT i PM; +20% na QA i kierowanie projektem", "+20% rezerwy: −1,1 mln zł"),
 ("Koszty", "Bezpieczeństwo + ład umowy", "0,5 + 0,3 mln zł/rok", "ZAŁOŻENIE", "OPZ: SIEM, pentest, RPO 4h/RTO 8h; audyt puli, depozyt kodu", "Koszty spoza P&L GA, wynikające ze współpracy", "–"),
 ("Kontrola", "Koszty puli vs P&L GA 2028-2033", "300,6 vs 293,1 mln zł", "GA", "Suma kosztów + rozwój w modelu GA", "Nasze koszty są o 2,6% WYŻSZE niż w P&L GA", "–"),
 ("Umowa", "Marża partnerów na usługach", "10%", "BENCHMARK", "Rynkowa marża EBIT software house'ów w PL ok. 10-15% (do potwierdzenia ofertą)", "Bez marży PT pracuje za darmo; IRR PT z 20% → 28%", "15%: PT 60,6 mln zł"),
 ("Umowa", "Finansowanie deficytu", "PT 45% / PM 55% / GA 0%", "ZAŁOŻENIE", "Proporcjonalnie do udziału w puli (PT ok. 36%, PM ok. 42%)", "Kto więcej dostaje, ten więcej finansuje", "–"),
]
for k, row in enumerate(A):
    for j, v in enumerate(row):
        c = put(ws, 5 + k, 1 + j, v, wrap=True)
        if j == 3: c.font = Font(name=FONT, size=10, bold=True, color=("008A3E" if v == "GA" else ("E67E00" if "ZAŁ" in v else NAVY)))
    ws.row_dimensions[5 + k].height = 44
for j, w in enumerate([11, 26, 26, 14, 48, 40, 26]): ws.column_dimensions[L(1 + j)].width = w
ws.freeze_panes = "A5"

# ---------- 3. GMV krok po kroku ----------
ws = sheet("3 GMV krok po kroku", "Skąd GMV 2,49 mld zł w 2032 (rok 5): formuły na założeniach z modelu GA", "Niebieskie na żółtym = założenie (można zmieniać); reszta liczy się sama.")
hdr(ws, 4, ["Krok", "Wartość", "Jedn.", "Źródło"])
G = [("Aktywni rolnicy", 850000, "szt.", "Model GA, C6", "#,##0", True), ("× ambicja: % rolników w 2032", 0.35, "%", "Model GA, wiersz 25", "0%", True),
     ("→ użytkownicy", "=B5*B6", "szt.", "", "#,##0", False), ("× użytkownicy transakcyjni", 0.75, "%", "Model GA, wiersz 28", "0%", True),
     ("→ kupujący przez platformę", "=B7*B8", "szt.", "", "#,##0", False), ("× wydatek rolnika rocznie", 58784, "zł", "Model GA, C20 (z maszynami)", "#,##0", True),
     ("× udział wydatków w GAia", 0.19, "%", "Model GA, wiersz 31", "0%", True), ("→ GMV 2032", "=B9*B10*B11/1000000", "mln zł", "Model GA: 2 492 mln zł", "#,##0", False),
     ("GMV / cały rynek (850 tys. × 58 784 zł)", "=B12/(B5*B10/1000000)", "%", "Kontrola: platforma obsługuje ok. 5% rynku", "0.0%", False),
     ("Przychody puli 2032", round(D["pool"]["rev"][5], 1), "mln zł", "Prowizje + back margin + premium + opłata za sukces", "#,##0.0", False),
     ("Przychody / GMV (take rate)", "=B14/B12", "%", "Model GA: 5,3%; Allegro ok. 11-12%", "0.0%", False)]
for k, (lab, v, u, src, fm, inp_) in enumerate(G):
    put(ws, 5 + k, 1, lab, bold=lab.startswith("→")); put(ws, 5 + k, 2, v, fm, inp=inp_, bold=lab.startswith("→"), fill=(TOT if lab.startswith("→") else None)); put(ws, 5 + k, 3, u); put(ws, 5 + k, 4, src)
for j, w in enumerate([42, 16, 9, 52]): ws.column_dimensions[L(1 + j)].width = w

# ---------- 4. Premium ----------
ws = sheet("4 Premium", "Premium: co rolnik dostaje, ile płaci i dlaczego 20 zł/ha (a nie 40 zł/ha z modelu GA)", "Cena do testu w pilocie (bramka 1). Ekoschemat: kwota z danych wtórnych, do potwierdzenia w ARiMR.")
hdr(ws, 4, ["Element", "Wartość", "Jedn.", "Komentarz"])
Pm = [("Cena premium", 20, "zł/ha/rok", "Nasze założenie", "#,##0", True), ("Średnie gospodarstwo", 16.47, "ha", "Model GA", "0.0", True),
      ("→ cena za gospodarstwo", "=B5*B6", "zł/rok", "ok. 27 zł miesięcznie", "#,##0", False),
      ("Model GA: 40 zł/ha", "=40*B6", "zł/rok", "GA: 'cena do ustalenia na późniejszym etapie'", "#,##0", False),
      ("Benchmark: mapy satelitarne (SatAgro)", "10-15", "zł/ha/rok", "Jedna funkcja (NDVI, mapy zmiennego dawkowania)", "@", False),
      ("Wartość dla rolnika: ekoschemat plan nawożenia", 87.52, "zł/ha", "Dopłata za plan nawożenia (wariant podstawowy, 2025); GAia premium generuje plan", "0.00", True),
      ("Cena / dopłata, którą premium pomaga uzyskać", "=B5/B10", "%", "Rolnik oddaje ok. 23% dopłaty, którą dostaje dzięki planowi", "0%", False),
      ("Oszczędność na nawozach 5% (1100 zł/ha)", "=1100*0.05", "zł/ha", "Precyzyjne nawożenie: 5-16% mniej N (badania; arkusz Dekarbonizacja w v4)", "#,##0", False),
      ("Udział płacących użytkowników", "3% → 10%", "", "Model GA wiersz 43; zwykle 1-2% płatnych, Allegro Smart ok. 30% (GA)", "@", False),
      ("Przychód premium 2032", round(D["pool"]["sub"][5], 1), "mln zł", "6% przychodów puli: premium nie jest filarem modelu", "#,##0.0", False)]
for k, (lab, v, u, src, fm, inp_) in enumerate(Pm):
    put(ws, 5 + k, 1, lab, bold=lab.startswith("→")); put(ws, 5 + k, 2, v, fm, inp=inp_, fill=(TOT if lab.startswith("→") else None)); put(ws, 5 + k, 3, u); put(ws, 5 + k, 4, src, wrap=True)
r = 5 + len(Pm) + 1; hdr(ws, r, ["Co rolnik ma w premium (z wymagań MVP1)", "", "", "Bez premium (bezpłatnie)"])
feat = [("Plan nawożenia z AI dla każdego pola (spełnia ekoschemat)", "Kokpit, profil, dokumenty"), ("Mapy NDVI z Sentinel i strefy plonowania na paszporcie pola", "Paszport pola: granice z ARiMR / GUGiK"),
        ("Alerty agronomiczne i pogodowe dla pól (IMGW) z rekomendacją", "Alert 6:30: ceny i podstawowa pogoda"), ("Kalkulator opłacalności i porównanie ofert dystrybutorów", "Zapytanie ofertowe do dystrybutora"),
        ("Raport śladu węglowego / MRV dla przetwórców", "Baza wiedzy nawozy.eu, harmonogram dotacji")]
for k, (a, b) in enumerate(feat):
    put(ws, r + 1 + k, 1, a, wrap=True); ws.merge_cells(start_row=r + 1 + k, start_column=1, end_row=r + 1 + k, end_column=3); put(ws, r + 1 + k, 4, b, wrap=True)
for j, w in enumerate([46, 12, 10, 60]): ws.column_dimensions[L(1 + j)].width = w

# ---------- 5. Churn i CAC ----------
ws = sheet("5 Churn i CAC", "Churn i koszt pozyskania: jak liczone i dlaczego", "Wszystko z modelu GA; pilot mierzy rzeczywiste wartości od pierwszego dnia.")
hdr(ws, 4, ["Element", "Wartość", "Jedn.", "Komentarz"])
Ch = [("Churn (odejścia) rocznie", 0.15, "%", "Model GA B72: 'akceptowalny dla aplikacji B2B' (dla porównania aplikacje konsumenckie: 30-50%+)", "0%", True),
      ("Koszt antichurn na odchodzącego", 350, "zł", "Model GA B74 = najwyższy CAC: koszt odzyskania / zastąpienia użytkownika", "#,##0", True),
      ("→ koszt churnu na użytkownika rocznie", "=B5*B6", "zł/użytk./rok", "Każdy użytkownik kosztuje średnio 52,5 zł rocznie utrzymania bazy", "#,##0.0", False),
      ("CAC 2028 → 2033", "250 → 350", "zł", "Model GA wiersz 69: kolejni rolnicy są drożsi", "@", False),
      ("Nowi użytkownicy 2028-2034", "=365500", "szt.", "Przyrost bazy (bez zastąpienia churnu)", "#,##0", False),
      ("Koszt pozyskania + antichurn 2028-2034", round(sum(D["costs"]["Pozyskanie i utrzymanie użytkowników"][1:8]), 1), "mln zł", "44% kosztów puli: największa pozycja", "#,##0.0", False),
      ("Przychód puli na użytkownika 2032", f"={round(D['pool']['rev'][5],1)}*1000000/297500", "zł/rok", "Na 1 użytkownika", "#,##0", False),
      ("Okres zwrotu CAC (CAC ÷ przychód na użytkownika)", "=350/B11", "lat", "Poniżej 1 roku w 2032", "0.0", False)]
for k, (lab, v, u, src, fm, inp_) in enumerate(Ch):
    put(ws, 5 + k, 1, lab, bold=lab.startswith("→")); put(ws, 5 + k, 2, v, fm, inp=inp_, fill=(TOT if lab.startswith("→") else None)); put(ws, 5 + k, 3, u); put(ws, 5 + k, 4, src, wrap=True)
for j, w in enumerate([44, 14, 13, 70]): ws.column_dimensions[L(1 + j)].width = w

# ---------- 6. Wrażliwość ----------
ws = sheet("6 Wrażliwość", "Co jeśli założenie się nie sprawdzi: suma gotówki 2027-2034, mln zł", "Przeliczone w pełnym modelu (LibreOffice), skrypt partner_variants_v5.py.")
hdr(ws, 4, ["Wariant", "PT", "PM", "GA", "Zwrot PT (rok)", "Finansowanie puli", "Wniosek"])
S = [("Bazowy (warunki proponowane)", "deal", "Wszyscy zarabiają; zwrot partnerów 2033"),
     ("Warunki z modelu GA: po koszcie, PT 70% deficytu", "base", "PT finansuje za dużo: IRR 20%, zaangażowanie 42 mln zł"),
     ("Ambicja 20% rolników w 2032", "deal_adopt20", "Mniej, ale na plusie"), ("CAC +100 zł", "deal_cac100", "Na plusie, finansowanie +21 mln zł"),
     ("Udział w wydatkach −30%", "deal_wal70", "RYZYKO KRYTYCZNE: KPI bramki 2"), ("Synergie potwierdzone ofertami (zespół −30%, utrzymanie −25%)", "deal_syn", "Potencjał, nie baza")]
for k, (lab, key, w) in enumerate(S):
    r = 5 + k; v = V[key]; put(ws, r, 1, lab, wrap=True, bold=(k == 0), fill=(TOT if k == 0 else None))
    for j, p in enumerate(["PT", "PM", "GA"]): put(ws, r, 2 + j, round(v[p]["sum7"], 1), NUM)
    put(ws, r, 5, 2027 + v["PT"]["payback"] if v["PT"]["payback"] is not None else "brak", "0"); put(ws, r, 6, round(v["pool"]["need"], 1), NUM); put(ws, r, 7, w, wrap=True)
    ws.row_dimensions[r].height = 30
for j, w in enumerate([46, 10, 10, 10, 13, 14, 46]): ws.column_dimensions[L(1 + j)].width = w
for w_ in wb.worksheets:
    w_.sheet_properties.tabColor = NAVY
    w_.page_setup.orientation = "landscape"; w_.page_setup.fitToWidth = 1; w_.sheet_properties.pageSetUpPr.fitToPage = True
wb.save(OUT); print("ok", OUT)
