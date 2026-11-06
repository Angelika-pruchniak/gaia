# 2 slajdy dla Empiku (Partner Marketplace) w stylu GA: zwrot z inwestycji, EBITDA platformy, co Empik wnosi i dostaje.
# Użycie: python3 build_empik_pptx.py <przeliczony_v5.xlsx>
import sys, json
src = open("build_v4_pptx.py", encoding="utf-8").read()
head = src[:src.index("finv = lambda")]
head = head.replace('rows(wb, "Finansowanie")', 'None').replace("GAia | model współpracy v4 | ambicja 35% rolników w roku 5, CAC 250 zł, 7 lat, liczby do potwierdzenia",
                                                              "GAia × Empik | propozycja współpracy | model v5 (P&L GA), gotówka, rok 0 = 2027 | poufne")
exec(head)
V = json.load(open("partner_variants_v5.json", encoding="utf-8")); D = V["deal"]; PM = D["PM"]; P = D["pool"]
CAL = [str(2027 + i) for i in range(8)]
f1 = lambda x: fmt(x, 1)
out = -sum(x for x in PM["res"] if x < 0); inn = sum(x for x in PM["res"] if x > 0)

# ===== 1. Zwrot dla Empiku =====
sl = frame("Empik: 31 mln zł zaangażowania, 93 mln zł wpływów", "Nowa kategoria B2B: rolnictwo, ok. 50 mld zł wydatków rocznie. Empik wnosi gotowy marketplace, GA wnosi rolników i dystrybucję")
tiles = [(f"{f1(PM['sum7'])} mln zł", "zysk gotówkowy 2027-2034"), (f"{round(PM['irr7']*100)}%", "IRR (próg rynkowy ok. 15%)"),
         (f"{fmt(inn/out, 1)}×", "zwrot z zaangażowanej gotówki"), ("2033", "pełny zwrot; od 2031 dodatnie przepływy")]
for i, (big, small) in enumerate(tiles):
    x = 0.6 + i * 3.07
    box(sl, x, 1.4, 2.9, 1.1, fill=GREEN if i == 0 else NAVY)
    tb(sl, x, 1.43, 2.9, 0.62, big, 24, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE); tb(sl, x, 2.03, 2.9, 0.4, small, 11, False, WHITE, PP_ALIGN.CENTER)
chart(sl, XL_CHART_TYPE.COLUMN_STACKED, 0.6, 2.7, 7.2, 3.35, CAL,
      [("Wkład w MVP1", [round(x, 1) for x in PM["own"]]), ("Finansowanie puli i jego zwrot", [round(x, 1) for x in PM["fin"]]),
       ("Marża 10% na usługach", [round(x, 1) for x in PM["mar"]]), ("Udział w zysku puli (ok. 42%)", [round(x, 1) for x in PM["dist"]])], [RED, AMBER, LBLUE, GREEN])
tb(sl, 0.75, 2.62, 6, 0.3, "Przepływy Empiku rok po roku, mln zł", 10.5, True, GREY)
box(sl, 8.0, 2.75, 4.7, 3.3, fill=LIGHT)
tb(sl, 8.12, 2.82, 4.5, 3.2, [[("Z czego zarabia Empik", True, NAVY)],
    [("1. Udział w zysku puli: ", True, TXT), (f"{f1(sum(PM['dist']))} mln zł. Klucz: 40% prowizji, 50% back marginu (akcje i kategorie to kompetencja Empiku), 15% premium, 40% opłaty za sukces.", False, TXT)],
    [("2. Zlecenia z puli: ", True, TXT), (f"{f1(sum(PM['c1'])+sum(PM['c2']))} mln zł pracy zespołu Empiku (operacje marketplace'u, kategorie, sprzedawcy) + 10% marży = {f1(sum(PM['mar']))} mln zł.", False, TXT)],
    [("3. Zwrot finansowania jako pierwszy: ", True, TXT), (f"55% deficytu 2028-2030 ({f1(-sum(x for x in PM['fin'] if x < 0))} mln zł) wraca w latach 2031-2033, przed podziałem.", False, TXT)]], 10.5)
msg(sl, [("Na start: ", True, WHITE), ("Empik wnosi w 2027 tylko 0,9 mln zł pracy (plug-in marketplace'u: katalog, zapytania, oferty). Resztę finansuje pula w transzach po bramkach KPI.", False, WHITE)], y=6.3)

# ===== 2. Platforma i dlaczego teraz =====
sl = frame("Platforma: EBITDA 90 mln zł w 2034, marża 55%", "Przychody i EBITDA wspólnej puli (gotówka, mln zł). Założenia z P&L Grupy Azoty; koszty zawierają 10% marży partnerów")
chart(sl, XL_CHART_TYPE.COLUMN_CLUSTERED, 0.6, 1.3, 7.2, 3.45, CAL[1:],
      [("Przychody puli", [round(x, 1) for x in P["rev"][1:]]), ("EBITDA", [round(x, 1) for x in P["net"][1:]])], [NAVY, GREEN], labels=True, numfmt='#,##0')
data = [["", "2030", "2032", "2034"], ["Użytkownicy (tys.)", fmt(D["users"][3]), fmt(D["users"][5]), fmt(D["users"][7])],
        ["GMV (mln zł)", fmt(D["gmv"][3]), fmt(D["gmv"][5]), fmt(D["gmv"][7])], ["Przychody puli", f1(P["rev"][3]), f1(P["rev"][5]), f1(P["rev"][7])],
        ["EBITDA", (f1(P["net"][3]), RED), (f1(P["net"][5]), GREEN), (f1(P["net"][7]), GREEN)],
        ["Marża EBITDA", "–", f"{round(P['net'][5]/P['rev'][5]*100)}%", f"{round(P['net'][7]/P['rev'][7]*100)}%"]]
table(sl, 8.0, 1.3, 4.7, data, [1.7, 1.0, 1.0, 1.0], rh=0.42, size=10.5, hl_rows=(4,))
tb(sl, 8.0, 4.05, 4.7, 1.3, [[("Dlaczego Empik i dlaczego teraz", True, NAVY)],
    [("• Rynek offline: rolnik kupuje przez telefon i handlowca, nikt nie ma platformy", False, TXT)],
    [("• GA daje popyt od 1. dnia: 66 dystrybutorów, marka, nawozy.eu, dane ARiMR", False, TXT)],
    [("• Empik ma to, czego GA nie ma: marketplace, sprzedawców, płatności, operacje", False, TXT)]], 10.5)
data = [["Test odporności (Empik, 8 lat)", "Wynik"], ["Ambicja 20% rolników zamiast 35%", f"+{f1(V['deal_adopt20']['PM']['sum7'])} mln zł"],
        ["Koszt pozyskania rolnika +100 zł", f"+{f1(V['deal_cac100']['PM']['sum7'])} mln zł"], ["Synergie potwierdzone ofertą Empiku", f"+{f1(V['deal_syn']['PM']['sum7'])} mln zł"]]
table(sl, 0.6, 4.9, 7.2, data, [5.2, 2.0], rh=0.27, size=9.5)
msg(sl, [("Propozycja: ", True, WHITE), ("umowa ze wspólną pulą na 10 + 5 lat, pilot MVP1 w 2027, spółka GAia po bramce 2. Kolejny krok: oferta Empiku na plug-in i operacje.", False, WHITE)], y=6.3)
prs.save("../GA_GAia_Empik_2_slajdy.pptx"); print("ok", N[0])
