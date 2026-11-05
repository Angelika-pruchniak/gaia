# Wypełnia szablon McK "Szablon Business Case" dla partnera (PT lub PM) danymi z partner_variants_v5.json.
# Edycja bezpośrednio w XML, żeby zachować pasek nawigacji, kształty, ochronę i style szablonu.
# Kalkulacja rozszerzona z 5 do 8 lat (2027-2034), bo platforma zwraca się partnerom w roku 6.
# Użycie: python3 build_bc_mck.py <szablon.xlsx> PT|PM <wyjście.xlsx>
import sys, re, json, zipfile, shutil
from xml.sax.saxutils import escape
TPL, WHO, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
V = json.load(open("partner_variants_v5.json", encoding="utf-8"))
YEARS = list(range(2027, 2035)); NYC = len(YEARS)
COLS = [chr(ord("D") + i) for i in range(NYC)]          # D..K
CM = chr(ord("D") + NYC)                                 # L = komentarz
M = 1e6
f1 = lambda x: f"{x:,.1f}".replace(",", " ").replace(".", ",")

def lines(v, who):
    d = V[v][who]
    work = [sum(t) for t in zip(*[d[k] for k in (("c1", "c2", "c3") if who == "PT" else ("c1", "c2"))])]
    fin_out = [min(0, x) for x in d["fin"]]; fin_in = [max(0, x) for x in d["fin"]]
    nm = "PT" if who == "PT" else "PM"
    return [
        (f"Zlecenia z puli dla zespołu {nm}", work, "Przychód: praca opłacana przez pulę"),
        (f"Koszt zespołu i usług {nm}", [-x for x in work], "Koszt tej samej pracy"),
        ("Marża na usługach", d["mar"], "Propozycja: 10% (model GA: 0%)" if v == "deal" else "Model GA: po koszcie"),
        ("Wkład własny w roku 0 (MVP1)", d["own"], "Silnik / marketplace, osobomiesiące × 20 tys. zł"),
        ("Finansowanie deficytu puli", fin_out, f"{'45%' if v == 'deal' else ('70%' if who == 'PT' else '30%')} deficytu, wraca pierwsze"),
        ("Zwrot finansowania", fin_in, "Z zysku puli, przed podziałem"),
        ("Udział w podziale puli", d["dist"], "Klucz: prowizje, back margin, premium, opłata za sukces"),
        (f"WYNIK {nm} (gotówka)", d["res"], f"7 lat: {f1(d['sum7'])} mln zł; zwrot w roku {d['payback']}"),
    ]

z = zipfile.ZipFile(TPL); files = {n: z.read(n) for n in z.namelist()}; z.close()
ss = files["xl/sharedStrings.xml"].decode("utf-8")

def cell_xml(ref, s, val):
    if val is None: return f'<c r="{ref}" s="{s}"/>'
    if isinstance(val, (int, float)): return f'<c r="{ref}" s="{s}"><v>{val}</v></c>'
    if isinstance(val, tuple) and val[0] == "f": return f'<c r="{ref}" s="{s}" t="str"><f>{escape(val[1])}</f><v>{escape(val[2])}</v></c>'
    if isinstance(val, tuple) and val[0] == "ss": return f'<c r="{ref}" s="{s}" t="s"><v>{val[1]}</v></c>'
    return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{escape(val)}</t></is></c>'

def row_cells(xml, r):
    m = re.search(rf'<row r="{r}"[^>]*?(?:/>|>(.*?)</row>)', xml, re.S)
    cells = {}
    if m and m.group(1):
        for cm in re.finditer(r'<c r="([A-Z]+)\d+"( s="(\d+)")?[^>]*?(?:/>|>.*?</c>)', m.group(1), re.S):
            cells[cm.group(1)] = (cm.group(3) or "0", cm.group(0))
    return m, cells

def set_cell(xml, ref, val):
    col, r = re.match(r"([A-Z]+)(\d+)", ref).groups()
    m, cells = row_cells(xml, r)
    s = cells[col][0] if col in cells else "0"
    new = cell_xml(ref, s, val)
    if col in cells: return xml.replace(cells[col][1], new, 1)
    raise KeyError(ref)

# ---------- Opis Business Case ----------
title = ("GAia: Partner Technologiczny (silnik, UX, integracje) w platformie Grupy Azoty" if WHO == "PT"
         else "GAia: Partner Marketplace / Handel (plug-in marketplace'u) w platformie Grupy Azoty")
d, b = V["deal"][WHO], V["base"][WHO]
work7 = sum(sum(d[k]) for k in (("c1", "c2", "c3") if WHO == "PT" else ("c1", "c2")))
share = "35-36%" if WHO == "PT" else "42-44%"
texts = {
 "C9": (f"Dlaczego: Grupa Azoty buduje platformę GAia dla 850 tys. rolników (ok. 50 mld zł wydatków rocznie) i zaprasza partnera {'technologicznego z gotowym silnikiem' if WHO == 'PT' else 'z gotowym marketplace’em'}. "
        f"Co: {'silnik na licencji, UX, grywalizacja, integracje ARiMR / Węzeł Krajowy, utrzymanie' if WHO == 'PT' else 'marketplace podpięty przez plug-in, sprzedawcy, płatności, operacje'}. "
        f"Wpływ dla partnera: {f1(d['sum7'])} mln zł gotówki w 7 lat (IRR {round(d['irr7']*100)}%), {f1(work7)} mln zł zleceń dla zespołu, zwrot w roku {d['payback']}."),
 "C13": ("Rynek środków do produkcji rolnej jest rozproszony i offline (telefon, handlowiec, dystrybutor). Brak platformy łączącej producenta, dystrybutorów i rolnika. "
         "Grupa Azoty ma produkty, logistykę, 66 dystrybutorów, markę, nawozy.eu, współpracę z ARiMR i 5 funkcjonalności na Azure, ale nie ma silnika platformy ani zespołu produktowego."),
 "C17": ("1) Prowizje od sprzedawców i dystrybutorów (1-1,5% nawozy i maszyny, 3-4% pozostałe). 2) Back margin od dostawców (bez nawozów GA). 3) Usługi premium 20 zł/ha. "
         "4) Opłata za sukces: 25% marży GA z lepszego mixu. 5) Kanał GA: dystrybutorzy, KAM, nawozy.eu, ARiMR. 6) Gotowe silniki partnerów: MVP1 w 188 osobomiesiącach."),
 "C21": ("Model GA (GAia-model_PL_v2) z korektami realności: adopcja 5% → 35% rolników w roku 5; udział w wydatkach 5% → 19%; CAC 250 → 350 zł; churn 15% × 350 zł; "
         "zespół 10 FTE + 0,2 FTE / 1000 użytk. × 25 tys. zł; bez synergii; bez sprzedaży inkrementalnej; rok 0 = 2027. Gotówka, mln zł, bez inflacji."),
 "C25": (f"Partner przestaje być dostawcą na zlecenie i staje się współwłaścicielem przychodów. Pula płaci za jego pracę (koszt + 10% marży), finansowanie deficytu ({'45%' if WHO == 'PT' else '55%'}) wraca jako pierwsze, "
         f"potem partner dostaje {share} zysku do podziału. Etap 1: umowa ze wspólną pulą; Etap 2: spółka GAia po bramce KPI (koniec roku 2)."),
 "C29": ("Model: GA_GAia_model_wspolpracy_v5.xlsx (formuły, komentarze, tabele przestawne); warianty: partner_variants_v5.py. Wymagania MVP1 (OPZ 01.06.2026, 28 WB). "
         "Prezentacja: GA_GAia_BC_partnerzy_technologiczni.pptx."),
 "C33": (f"Wynik partnera (propozycja): {f1(d['sum6'])} mln zł do 2033, {f1(d['sum7'])} mln zł do 2034; NPV (10%) {f1(d['npv7'])} mln zł; maks. zaangażowanie {f1(-d['trough'])} mln zł (rok {d['cum'].index(d['trough'])}). "
         f"Na warunkach z modelu GA (po koszcie): {f1(b['sum7'])} mln zł, zaangażowanie {f1(-b['trough'])} mln zł."),
}
risks = [("Wolniejsza adopcja / niższy udział w wydatkach", "Bramka 2 (≥ 7% rolników, pula ≥ 10 mln zł/rok); finansowanie w transzach po bramkach; KPI transakcji od pilota"),
         ("Wyższy koszt pozyskania (CAC)", "Kanał 66 dystrybutorów i KAM GA; CAC mierzony od pilota; limit budżetu pozyskania na rok"),
         ("Konflikt z kanałem dystrybutorów", "MVP1 = zapytanie do dystrybutora; dystrybutor dostaje leady i kartę klienta"),
         ("Spór o koszty / podział", "Koszty po koszcie + 10% marży, budżet roczny, audyt; klucz w umowie, przegląd co 2 lata")]

x = files["xl/worksheets/sheet2.xml"].decode("utf-8")
x = set_cell(x, "D4", title)
x = set_cell(x, "D5", ("ss", 84))
for k, t in texts.items(): x = set_cell(x, k, t)
for i, (r_, m_) in enumerate(risks): x = set_cell(x, f"C{38+i}", r_); x = set_cell(x, f"D{38+i}", m_)
files["xl/worksheets/sheet2.xml"] = x.encode("utf-8")
ss_list = re.findall(r"<si>.*?</si>", ss, re.S)
ss = ss.replace(ss_list[84], "<si><t>Platforma GAia – partnerstwa</t></si>", 1)
files["xl/sharedStrings.xml"] = ss.encode("utf-8")

# ---------- Kalkulacja Business Case: 8 lat ----------
x = files["xl/worksheets/sheet3.xml"].decode("utf-8")
x = re.sub(r"<cols>.*?</cols>", '<cols><col min="1" max="1" width="17.5" style="43" customWidth="1"/><col min="2" max="2" width="2.33" style="43" customWidth="1"/>'
           '<col min="3" max="3" width="34" style="43" customWidth="1"/><col min="4" max="11" width="10.5" style="43" customWidth="1"/>'
           '<col min="12" max="12" width="40" style="43" customWidth="1"/><col min="13" max="13" width="2.33" style="43" customWidth="1"/>'
           '<col min="14" max="16384" width="10" style="43" hidden="1"/></cols>', x, flags=re.S)
tab = {14: ("base", "Warunki z modelu GA: po koszcie, partner finansuje {fs} deficytu"), 26: ("deal", "Propozycja: koszt + 10% marży, finansowanie wg udziału w puli")}
def rebuild(xml, r, vals):
    m, cells = row_cells(xml, r)
    sD = cells.get("D", cells.get("C", ("0",)))[0]; sI = cells.get("I", (sD,))[0]; sC = cells.get("C", (sD,))[0]
    out = [cell_xml(f"C{r}", sC, vals.get("C"))] + [cell_xml(f"{c}{r}", sD, vals.get(c)) for c in COLS] + [cell_xml(f"{CM}{r}", sI, vals.get(CM))]
    head = re.match(r'<row [^>]*?(?=/?>)', m.group(0)).group(0)
    head = re.sub(r'spans="[^"]*"', 'spans="3:12"', head)
    return xml.replace(m.group(0), head + ">" + "".join(out) + "</row>", 1)
for r in list(range(7, 34)):
    m, cells = row_cells(x, r)
    if not m or not cells: continue
    vals = {}
    if r in (13, 25):
        vals = {"C": ("ss", 101), CM: ("ss", 29)}; vals.update({c: y for c, y in zip(COLS, YEARS)})
    elif r in (11, 23): vals = {"C": ("ss", 76 if r == 11 else 75)}
    elif r == 7: vals = {"C": ("ss", 79)}
    elif r == 9:
        vals = {"C": f"Rok 0 = 2027 (MVP1 + pilot). Model GA z korektami: CAC 250 → 350 zł, bez synergii, back margin bez nawozów GA, premium 20 zł/ha, bez sprzedaży inkrementalnej. "
                     f"Stan bazowy = warunki z modelu GA; szacunek docelowy = proponowane warunki dla partnera. Wartości w PLN, format w mln."}
    elif 14 <= r <= 21 or 26 <= r <= 33:
        v = "base" if r < 23 else "deal"; L_ = lines(v, WHO)[r - (14 if r < 23 else 26)]
        vals = {"C": L_[0], CM: L_[2]}; vals.update({c: round(val * M) for c, val in zip(COLS, L_[1])})
    x = rebuild(x, r, vals)
x = x.replace('<mergeCell ref="C7:H7"/>', '<mergeCell ref="C7:K7"/>').replace('<mergeCell ref="C11:H11"/>', '<mergeCell ref="C11:K11"/>')
x = x.replace('<mergeCell ref="C23:H23"/>', '<mergeCell ref="C23:K23"/>').replace('<mergeCell ref="C9:I9"/>', '<mergeCell ref="C9:L9"/>')
x = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="A1:M34"/>', x)
files["xl/worksheets/sheet3.xml"] = x.encode("utf-8")
# wymuś przeliczenie (tytuł w Kalkulacji pochodzi z formuły)
wbx = files["xl/workbook.xml"].decode("utf-8").replace('<calcPr calcId="191029"/>', '<calcPr calcId="191029" fullCalcOnLoad="1"/>')
files["xl/workbook.xml"] = wbx.encode("utf-8")
files.pop("xl/calcChain.xml", None)
ct = files["[Content_Types].xml"].decode("utf-8"); ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", ct); files["[Content_Types].xml"] = ct.encode("utf-8")
rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8"); rels = re.sub(r'<Relationship [^>]*Target="calcChain.xml"/>', "", rels); files["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zo:
    for n, data in files.items(): zo.writestr(n, data)
print("ok", OUT)
