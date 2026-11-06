# 5 slajdów do obrony business case (szablon GA). Liczby z partner_variants_v5.json (model v5).
# Użycie: python3 build_obrona_pptx.py <przeliczony_v5.xlsx>
import sys, json
src = open("build_v4_pptx.py", encoding="utf-8").read()
head = src[:src.index("finv = lambda")]
head = head.replace('rows(wb, "Finansowanie")', 'None').replace("GAia | model współpracy v4 | ambicja 35% rolników w roku 5, CAC 250 zł, 7 lat, liczby do potwierdzenia",
                                                              "GAia | business case partnerów | model v5 (P&L GA z korektami), gotówka, rok 0 = 2027")
exec(head)
from pptx.chart.data import CategoryChartData
V = json.load(open("partner_variants_v5.json", encoding="utf-8")); D = V["deal"]
CAL = [str(2027 + i) for i in range(8)]
f1 = lambda x: fmt(x, 1)

def line_chart(sl, x, y, w, h, cats, series, colors):
    cd = CategoryChartData(); cd.categories = cats
    for nm, vals in series: cd.add_series(nm, vals)
    ch = sl.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(x), Inches(ty(y)), Inches(w), Inches(th(y, h)), cd).chart
    ch.font.name = FONT; ch.font.size = Pt(10); ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.BOTTOM; ch.legend.include_in_layout = False
    for s, c in zip(ch.series, colors):
        s.format.line.color.rgb = c; s.format.line.width = Pt(2.75); s.smooth = False
        s.marker.format.fill.solid(); s.marker.format.fill.fore_color.rgb = c; s.marker.format.line.color.rgb = c
    ch.value_axis.has_major_gridlines = True; ch.value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xD9, 0xD9, 0xD9)
    ch.value_axis.format.line.fill.background(); ch.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    ch.value_axis.tick_labels.number_format = '#,##0'; ch.value_axis.tick_labels.number_format_is_linked = False
    return ch

# ===== 1. Kto, kiedy, ile =====
sl = frame("Każdy partner zarabia: GA od 2030, partnerzy od 2033", "Gotówka 2027-2034, mln zł. Warunki: pula płaci za pracę (koszt + 10%), deficyt finansują PT 45% / PM 55%, GA 0 zł")
data = [["", "Wnosi", "Jak zarabia", "Na plusie od", "Suma 8 lat", "IRR", "Maks. zaangażowanie"]]
rows_ = [("Partner Technologiczny", "PT", "Silnik, UX, integracje; MVP1 4,8 mln zł", "Marża 10% na pracy + zwrot finansowania + ok. 36% zysku"),
         ("Partner Marketplace", "PM", "Marketplace (plug-in); MVP1 0,9 mln zł", "Marża 10% + zwrot finansowania + ok. 42% zysku"),
         ("Grupa Azoty", "GA", "Kanał, marka, dystrybutorzy, dane; 0 zł", "Lepszy mix nawozów (75% dla GA) + ok. 22% zysku")]
for nm, k, wn, jak in rows_:
    d = V["deal"][k]
    data.append([nm, wn, jak, str(2027 + d["payback"]), (f1(d["sum7"]), GREEN), "–" if k == "GA" else f"{round(d['irr7']*100)}%", (f1(d["trough"]), RED if d["trough"] < -1 else TXT)])
table(sl, 0.6, 1.25, 12.1, data, [2.1, 2.9, 3.5, 1.0, 0.9, 0.6, 1.1], rh=0.5, size=10.5)
line_chart(sl, 0.6, 3.45, 7.4, 2.75, CAL, [(n, [round(x, 1) for x in V["deal"][k]["cum"]]) for n, k in (("Partner Technologiczny", "PT"), ("Partner Marketplace", "PM"), ("Grupa Azoty", "GA"))], [MBLUE, LBLUE, NAVY])
tb(sl, 0.75, 3.38, 6, 0.3, "Wynik narastająco, mln zł", 10.5, True, GREY)
tb(sl, 8.2, 3.45, 4.5, 2.75, [[("Dlaczego partnerzy dopiero od 2033?", True, NAVY)],
    [(f"Pula ma deficyt 2028-2030 (pozyskanie rolników). Najgłębszy: {f1(D['pool']['need'])} mln zł w 2030.", False, TXT)],
    [("Zysk puli najpierw zwraca finansowanie (2031-2033), potem jest dzielony.", False, TXT)],
    [("Dlaczego GA wcześniej?", True, NAVY)], [("GA nie finansuje deficytu; marża z mixu wpływa od 2028.", False, TXT)],
    [("Platforma: ", True, NAVY), (f"EBITDA dodatnia od 2031; 2034: {f1(D['pool']['net'][7])} mln zł.", False, TXT)]], 11)
msg(sl, [("Arkusz 1 w Excelu: ", True, WHITE), ("każda kwota rok po roku, sumy i narastająco liczone formułami.", False, WHITE)], y=6.3)

# ===== 2. Przychody i GMV =====
sl = frame("Przychody: GMV 2,49 mld zł w 2032 to 5% rynku", "Każdy krok z modelu GA (GAia-model_PL_v2, „BC agressive”); naszą zmianą są tylko niższe stawki")
steps = [("850 tys.", "aktywnych rolników", "GA C6"), ("× 35%", "ambicja 2032", "GA w.25"), ("× 75%", "kupuje przez GAia", "GA w.28"),
         ("× 58 784 zł", "wydatek rolnika/rok", "GA C20"), ("× 19%", "udział w wydatkach", "GA w.31"), ("= 2 492", "mln zł GMV 2032", "5% rynku 50 mld zł")]
for i, (big, sm, srcx) in enumerate(steps):
    x = 0.6 + i * 2.03
    box(sl, x, 1.3, 1.9, 1.15, fill=GREEN if i == 5 else NAVY)
    tb(sl, x, 1.32, 1.9, 0.5, big, 17, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE); tb(sl, x, 1.8, 1.9, 0.3, sm, 10.5, False, WHITE, PP_ALIGN.CENTER)
    tb(sl, x, 2.12, 1.9, 0.3, srcx, 9.5, True, PALE, PP_ALIGN.CENTER)
rv = D["pool"]
data = [["Strumień przychodu", "Stawka", "Źródło / benchmark", "2032, mln zł", "Suma 8 lat"],
        ["Prowizje (płaci sprzedawca / dystrybutor)", "nawozy i maszyny 1-1,5%; reszta 3-4%", "GA w.34-35; Allegro ok. 12%", f1(rv["comm"][5]), f1(sum(rv["comm"]))],
        ["Back margin od dostawców", "1-1,5% / 2-2,2%, bez nawozów GA", "GA w.39; GA nie płaci sama sobie", f1(rv["back"][5]), f1(sum(rv["back"]))],
        ["Premium", "20 zł/ha, 3% → 10% użytk.", "GA: 40 zł/ha; SatAgro 10-15 zł/ha", f1(rv["sub"][5]), f1(sum(rv["sub"]))],
        ["Opłata za sukces od GA", "25% marży z mixu nawozów GA", "GA w.59: marża 0,5-1,2%", f1(rv["succ"][5]), f1(sum(rv["succ"]))],
        ["PRZYCHODY PULI", f"{f1(rv['rev'][5]/2492*100)}% GMV (model GA: 5,3%)", "Allegro ok. 11-12% (GA)", f1(rv["rev"][5]), f1(sum(rv["rev"]))]]
table(sl, 0.6, 2.75, 12.1, data, [3.4, 3.0, 3.2, 1.2, 1.3], rh=0.44, size=10.5, hl_rows=(5,))
msg(sl, [("Kluczowe założenie: ", True, WHITE), ("udział w wydatkach 19%. GA porównuje go z top kategoriami e-commerce (fashion 24%, RTV 20%). Rolnik kupuje u 1-2 dystrybutorów.", False, WHITE)], y=6.3)

# ===== 3. Koszty, churn, CAC =====
c = D["costs"]; s8 = lambda k: sum(c[k][1:8]); y5 = lambda k: c[k][5]
sl = frame("Koszty: zgodne z P&L GA, największa pozycja to pozyskanie rolników", "Koszty puli 2028-2033: 300,6 mln zł vs 293,1 mln zł w P&L GA (+2,6%). Plus MVP1 5,6 mln zł w 2027 (GA: 4 mln zł)")
data = [["Koszt", "Jak liczony", "Źródło", "2032", "Suma 8 lat"],
        ["Pozyskanie (CAC)", "nowi użytkownicy × 250 → 350 zł", "GA w.69", f1(y5("Pozyskanie użytkowników (CAC)")), f1(s8("Pozyskanie użytkowników (CAC)"))],
        ["Churn (antichurn)", "15% użytk. × 350 zł = 52,5 zł/użytk./rok", "GA B72-B74", f1(y5("Antichurn")), f1(s8("Antichurn"))],
        ["Zespół", "10 FTE + 0,2 FTE/1000 użytk. × 25 tys. zł (18 → 70 FTE)", "GA: 15 → 80 FTE", f1(y5("Zespół (IT, produkt, obsługa, zakupy)")), f1(s8("Zespół (IT, produkt, obsługa, zakupy)"))],
        ["Utrzymanie (usługi obce, SaaS)", "1,2 mln zł + 36 zł/użytk. (11% przychodów 2032)", "GA: 2,4 → 12 mln; 5-12% przych.", f1(y5("Utrzymanie (usługi obce, SaaS)")), f1(s8("Utrzymanie (usługi obce, SaaS)"))],
        ["Rozwój", "1,5 → 3 mln zł/rok", "GA w.64", f1(y5("Rozwój platformy")), f1(s8("Rozwój platformy"))],
        ["Bezpieczeństwo + ład umowy", "0,5 + 0,3 mln zł/rok", "OPZ MVP1; audyt puli", "0,8", f1(s8("Bezpieczeństwo i zgodność") + s8("Prawne, PMO, audyt puli, depozyt kodu"))],
        ["Marża partnerów 10%", "na pracy PT i PM", "marża SH 10-15%", f1(y5("Marża PT na usługach") + y5("Marża PM na usługach")), f1(s8("Marża PT na usługach") + s8("Marża PM na usługach"))],
        ["KOSZTY PULI", "", "", f1(y5("KOSZTY WSPÓLNE PULI (w partnerstwie)")), f1(s8("KOSZTY WSPÓLNE PULI (w partnerstwie)"))]]
table(sl, 0.6, 1.25, 8.4, data, [2.2, 3.4, 1.6, 0.6, 0.6], rh=0.47, size=10, hl_rows=(8,))
box(sl, 9.2, 1.25, 3.5, 4.7, fill=LIGHT)
tb(sl, 9.32, 1.32, 3.3, 4.6, [[("Churn: dlaczego 15%?", True, NAVY)], [("Model GA: „akceptowalny dla aplikacji B2B”. Rolnik ma powód wracać: oferty, dopłaty, alerty. Aplikacje konsumenckie tracą 30-50%+ rocznie.", False, TXT)],
    [("Dlaczego 350 zł na odchodzącego?", True, NAVY)], [("Tyle kosztuje odzyskanie lub zastąpienie użytkownika (= najwyższy CAC).", False, TXT)],
    [("Czy CAC się zwraca?", True, NAVY)], [(f"Przychód puli 2032: ok. {fmt(rv['rev'][5]*1e6/297500)} zł na użytkownika rocznie, CAC 350 zł → zwrot w ok. 1 rok.", False, TXT)],
    [("Bez synergii w CAC ", True, NAVY), ("(uwaga GA 04.11).", False, TXT)]], 11)
msg(sl, [("Weryfikacja 05.11: ", True, WHITE), ("odrzuciłem doszacowanie, które liczyło SMS, programistów i płatności podwójnie. Zostało +1,9 mln zł na MVP1 i 0,8 mln zł/rok.", False, WHITE)], y=6.3)

# ===== 4. Premium =====
sl = frame("Premium 20 zł/ha: rolnik płaci ok. 23% dopłaty, którą dzięki niemu dostaje", "Model GA zakładał 40 zł/ha (≈ 659 zł/rok); przyjęliśmy połowę. Premium to tylko 6% przychodów puli")
cols = [("Co rolnik dostaje w premium", NAVY, ["Plan nawożenia z AI dla każdego pola (wymóg ekoschematu)", "Mapy NDVI z Sentinel, strefy plonowania na paszporcie pola", "Alerty agronomiczne i pogodowe z rekomendacją", "Porównanie ofert i kalkulator opłacalności", "Raport śladu węglowego dla przetwórców"]),
        ("Bezpłatnie (dla wszystkich)", MBLUE, ["Paszport pola z danymi ARiMR / GUGiK", "Zapytanie ofertowe do dystrybutorów", "Alert 6:30: ceny, pogoda", "Baza wiedzy nawozy.eu, harmonogram dotacji", "Kokpit, dokumenty, profil"])]
for i, (t, col, items) in enumerate(cols):
    x = 0.6 + i * 3.95
    box(sl, x, 1.3, 3.8, 0.5, fill=col); tb(sl, x, 1.3, 3.8, 0.5, t, 13, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    box(sl, x, 1.85, 3.8, 2.6, line=col); tb(sl, x + 0.1, 1.92, 3.6, 2.5, ["• " + s for s in items], 11)
data = [["Ile i dlaczego", "Wartość"], ["Cena", "20 zł/ha × 16,5 ha = 330 zł/rok (ok. 27 zł/mies.)"], ["Model GA", "40 zł/ha = 659 zł/rok"],
        ["Benchmark: mapy satelitarne (SatAgro)", "10-15 zł/ha (jedna funkcja)"], ["Wartość: ekoschemat plan nawożenia", "ok. 87,52 zł/ha dopłaty (2025)"],
        ["Wartość: 5% mniej nawozu", "ok. 55 zł/ha (5% × 1 100 zł/ha)"], ["Płacący", "3% → 10% użytkowników (GA w.43)"], ["Przychód 2032", f"{f1(rv['sub'][5])} mln zł (6% puli)"]]
table(sl, 8.6, 1.3, 4.1, data, [2.0, 2.1], rh=0.45, size=10)
tb(sl, 0.6, 4.6, 7.8, 1.5, [[("Dlaczego 20 zł/ha, a nie 40 zł/ha: ", True, NAVY), ("pakiet ma 5 funkcji, a samo mapowanie satelitarne kosztuje rynkowo 10-15 zł/ha. 40 zł/ha wymaga dowodu wartości. ", False, TXT),
    ("Cenę testujemy w pilocie (bramka 1).", True, NAVY)], [("Argument dla rolnika: ", True, NAVY), ("plan nawożenia daje ok. 87,52 zł/ha dopłaty, a mniej nawozu ok. 55 zł/ha. Premium kosztuje 20 zł/ha.", False, TXT)]], 11.5)
msg(sl, [("Jeśli premium nie zadziała: ", True, WHITE), ("wynik partnerów spada o ok. 6% przychodów; model stoi na prowizjach i back marginie.", False, WHITE)], y=6.3)

# ===== 5. Ryzyka =====
sl = frame("Co musi się sprawdzić: udział w wydatkach to ryzyko krytyczne", "Suma gotówki 2027-2034, mln zł. Pełne przeliczenie modelu dla każdego wariantu")
S = [("Bazowy (warunki proponowane)", "deal"), ("Warunki z modelu GA (po koszcie, PT 70% deficytu)", "base"), ("Ambicja 20% rolników (zamiast 35%)", "deal_adopt20"),
     ("CAC +100 zł", "deal_cac100"), ("Udział w wydatkach −30%", "deal_wal70"), ("Synergie potwierdzone ofertami (potencjał)", "deal_syn")]
data = [["Wariant", "PT", "PM", "GA", "Zwrot PT", "Finansowanie puli"]]
for lab, k in S:
    v = V[k]; data.append([lab] + [(f1(v[p]["sum7"]), sign(v[p]["sum7"])) for p in ("PT", "PM", "GA")] + [str(2027 + v["PT"]["payback"]) if v["PT"]["payback"] is not None else "brak", f1(v["pool"]["need"])])
table(sl, 0.6, 1.25, 7.9, data, [3.6, 0.85, 0.85, 0.85, 0.85, 0.9], rh=0.5, size=10.5, hl_rows=(1,))
data = [["KPI do bramki", "Próg"], ["Bramka 1 (pilot 2027)", "≥ 10 tys. użytkowników, ≥ 5 tys. zapytań, CAC ≤ 250 zł, test ceny premium"],
        ["Bramka 2 (koniec 2029)", "≥ 7% rolników, udział w wydatkach ≥ 10%, pula ≥ 10 mln zł/rok"], ["Zatrzymanie po bramce 2", f"strata PT {f1(D['PT']['gate2'])}, PM {f1(D['PM']['gate2'])} mln zł; GA 0"]]
table(sl, 8.7, 1.25, 4.0, data, [1.3, 2.7], rh=0.75, size=9.5)
tb(sl, 0.6, 4.95, 12.1, 0.9, [[("Źródła: ", True, NAVY), ("GA = model GAia-model_PL_v2 (rynek, adopcja, udział w wydatkach, stawki, CAC, churn, zespół, utrzymanie). ", False, TXT),
    ("Benchmark = Allegro, SatAgro, ekoschemat, e-commerce. ", False, TXT), ("Założenie = MVP1, marża 10%, podział finansowania. ", False, TXT), ("Szczegóły: Excel, arkusz 2.", True, NAVY)]], 10.5)
msg(sl, [("Prosimy o: ", True, WHITE), ("zgodę na rozmowy z partnerami na tych warunkach i pilot z KPI udziału w wydatkach od pierwszego dnia.", False, WHITE)], y=6.3)
prs.save("../GA_GAia_obrona_5_slajdow.pptx"); print("ok", N[0])
