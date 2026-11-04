# Arkusze tekstowe modelu v3: model współpracy, wkład partnerów, uwagi do modelu GA, case study.
def text_sheet(title, head, sub, cols, widths, rows_, idx=None):
    ws = wb.create_sheet(title) if idx is None else wb.create_sheet(title, idx)
    ws["A1"] = head; ws["A2"] = sub
    hdr(ws, 4, cols)
    for k, rw in enumerate(rows_):
        for j, t in enumerate(rw):
            c = ws.cell(5 + k, 1 + j, t); c.alignment = Alignment(wrap_text=True, vertical="top")
            if isinstance(t, str) and t.startswith("http"): c.hyperlink = t
    for j, w in enumerate(widths): ws.column_dimensions[L(1 + j)].width = w
    return ws, 4 + len(rows_)

WSP, WSP_last = text_sheet("Model współpracy", "Rekomendowany model współpracy: długoterminowy alians z wspólną pulą, potem spółka GAia",
    "Oparty na praktykach JV (Umbrex, Ankura, WIPO, Morgan Lewis) i przykładach: Starbucks × PepsiCo (30+ lat), Bosch × BASF, Bayer × Microsoft, Nutrien, Syngenta MAP.",
    ["Obszar", "Zasada", "Dlaczego (benchmark / praktyka)"], [30, 90, 70], [
    ("Struktura", "Etap 1 (lata 0-2): umowa partnerska (konsorcjum) ze wspólną pulą przychodów na rachunku powierniczym. Etap 2 (od roku 3, po bramce KPI): spółka celowa GAia sp. z o.o.", "Bosch i BASF: współpraca od 2019, spółka w 2021. Mniejsze ryzyko i koszt na starcie"),
    ("Okres", "10 lat z automatycznym przedłużeniem o 5 lat. Wypowiedzenie bez przyczyny najwcześniej po 5 latach, z 12-miesięcznym okresem przejściowym", "Długoterminowe alianse (Starbucks × PepsiCo od 1994) dają partnerom czas na zwrot wkładu"),
    ("Wkład GA", "Produkty i produkcja, logistyka i punkty wydania, 66 dystrybutorów, marka i marketing, nawozy.eu, współpraca z ARiMR i Węzłem Krajowym, 5 funkcjonalności na Azure (audyt, licencja). Bez gotówki", "Wkład rzeczowy za udział (Ankura, Umbrex)"),
    ("Wkład Partnera Technologicznego (SH)", "Silnik platformy na licencji dla GAia, warstwa UX (web + mobile), grywalizacja, integracje (ARiMR, Węzeł Krajowy, SAP, banki, plug-in marketplace'u), utrzymanie i bezpieczeństwo zgodnie z OPZ MVP1", "IP wnoszone licencją, nie przenoszone (WIPO); Bayer × Microsoft"),
    ("Wkład Partnera Marketplace", "Gotowy marketplace podpięty przez plug-in: katalog, zapytania ofertowe, oferty dystrybutorów, płatności, sprzedawcy, reklama; operacje marketplace'u", "Każdy wnosi to, w czym jest najlepszy (Starbucks marka + PepsiCo dystrybucja)"),
    ("Pula przychodów", "Wszystkie przychody GAia (prowizje, back margin, premium, opłata za sukces GA) wpływają na wspólny rachunek. Z puli: koszty wspólne po koszcie, bez marży, w ramach rocznego budżetu", "Nikt nie zarabia na kosztach: brak konfliktu interesów"),
    ("Podział", "Po zwrocie deficytu: podział według klucza dla każdego strumienia (arkusz Klucz podziału). Przegląd klucza co 2 lata według rzeczywistego wkładu", "Sklepy z aplikacjami 70/30, Salesforce 85/15: większość dla tego, kto tworzy wartość"),
    ("GA nie płaci", "GA nie wnosi gotówki i nie dopłaca deficytu. Płaci tylko 25% DODATKOWEJ marży z platformy (opłata za sukces). Prowizje od nawozów płacą sprzedawcy / dystrybutorzy", "Model afiliacyjny: płacisz od efektu"),
    ("Finansowanie startu", "Deficyt puli finansują Partner Technologiczny (70%) i Partner Marketplace (30%). Wraca do nich jako pierwszy z zysku. Możliwe wsparcie zewnętrzne: granty NCBR / FENG, finansowanie bankowe", "Preferencyjny zwrot dla wspólnika finansującego"),
    ("Zarządzanie", "Komitet Sterujący: po 2 osoby od każdego partnera, przewodniczy GA. Sprawy zastrzeżone jednomyślnie: budżet, klucz, nowy partner, sprzedaż. Wspólne biuro projektu (PMO)", "Rada, sprawy zastrzeżone, ścieżka eskalacji (Umbrex)"),
    ("Bramki KPI", "Bramka 0: MVP1 gotowe. Bramka 1 (pilot): ≥ 10 tys. użytkowników, ≥ 5 tys. zapytań. Bramka 2 (koniec roku 2): ≥ 7% rolników, pula ≥ 10 mln zł/rok, EBITDA ≥ 0 lub plan dojścia", "Etapowanie ryzyka jak Bosch × BASF"),
    ("SLA i jakość", "Dostępność 99,5%, RPO 4 h, RTO 8 h (OPZ WBI-07), pentest przed wdrożeniem, audyt roczny kosztów puli", "Wymagania OPZ MVP1"),
    ("Dane", "Rolnik decyduje o swoich danych (zgody RODO). GA i spółka GAia współadministratorami, partnerzy podmiotami przetwarzającymi. Dane zanonimizowane dla analiz", "OPZ WB-23, WB-25, WRL-08"),
    ("Własność IP", "Silnik PT, marketplace PM i 5 funkcjonalności GA zostają u właścicieli (licencje na czas współpracy). Nowy kod i dane GAia należą do spółki. Depozyt kodu źródłowego (escrow)", "Background vs foreground IP (WIPO, Morgan Lewis)"),
    ("Wyłączność", "Partnerzy nie budują konkurencyjnej platformy dla innego producenta nawozów w Polsce przez okres umowy; GA nie buduje równoległej platformy", "Ochrona inwestycji wszystkich stron"),
    ("Wyjście", "Po 5 latach prawo pierwokupu udziałów; opcja wykupu partnera przez GA po formule (np. 6-8 × EBITDA); okres przejściowy 12 mies. z utrzymaniem licencji", "Ścieżka wyjścia ustalona z góry (Umbrex)"),
])

WKP, WKP_last = text_sheet("Wkład partnerów", "Kto co wnosi, co robi i co dostaje", "GA wnosi zasoby, których partnerzy nie mają; partnerzy wnoszą gotowe silniki i pracę.",
    ["Partner", "Wnosi", "Robi w roku 0 i potem", "Dostaje", "Ryzyko"], [28, 60, 56, 52, 32], [
    ("Grupa Azoty", "Produkty, produkcja, logistyka, 66 dystrybutorów, marka i marketing, nawozy.eu, współpraca z ARiMR, Węzeł Krajowy, 5 funkcjonalności na Azure, dane rolników (za zgodą)",
     "Udostępnia dane, treści i funkcjonalności; dystrybutorzy odpowiadają na zapytania; kampanie w sezonie", "75% dodatkowej marży + 20% prowizji + 30% back marginu + 15% premium", "Bez wkładu pieniężnego"),
    ("Partner Technologiczny (SH)", "Silnik platformy (licencja), zespół IT i produkt, utrzymanie, bezpieczeństwo", "Rok 0: MVP1 według OPZ, aplikacja mobilna, grywalizacja, integracje. Potem rozwój i utrzymanie",
     "Zwrot kosztów z puli + 40% prowizji + 20% back marginu + 70% premium + 60% opłaty za sukces", "Wkład pracy w roku 0, 70% deficytu"),
    ("Partner Marketplace / Handel", "Gotowy marketplace (licencja, plug-in), sprzedawcy, płatności, reklama", "Rok 0: plug-in, katalog, zapytania i oferty. Potem operacje, kategorie, akcje promocyjne",
     "Zwrot kosztów z puli + 40% prowizji + 50% back marginu + 15% premium + 40% opłaty za sukces", "Wkład pracy w roku 0, 30% deficytu")])

UW, UW_last = text_sheet("Uwagi do modelu GA", "Co przejęliśmy z modelu GAia-model_PL_v2 i co zmieniliśmy", "Scenariusz agresywny odtwarza model GA. Poniżej różnice i powody.",
    ["Element modelu GA", "W modelu GA", "W modelu v3", "Dlaczego"], [34, 52, 52, 60], [
    ("Rynek i wydatki", "850 tys. rolników, 16,5 ha, 3 511 zł/ha + 960 zł nierolniczych", "Bez zmian", "Spójne dane wejściowe"),
    ("Adopcja", "5% → 35% rolników (rok 5), 40% (rok 6)", "3 scenariusze: 12% / 22% / 35% w roku 5; agresywny = GA", "35% to bardzo szybkie tempo; potrzebny wariant ostrożny i bazowy"),
    ("Udział w wydatkach rolnika", "5% → 19% (rok 5)", "Ostrożny 9%, bazowy 13%, agresywny 19% w roku 5", "19% to poziom top kategorii e-commerce (fashion 24%)"),
    ("Horyzont", "6 lat", "7 lat", "Współpraca długoterminowa; platformy zwracają się w 4-5 lat"),
    ("Rozwój (capex)", "4 mln zł w roku 0 + 1,5-3 mln zł rocznie, poza sumą kosztów", "Rok 0: wkład partnerów wyliczony z wymagań MVP1 (arkusz Dopasowanie MVP1); kolejne lata w kosztach puli", "W modelu GA rozwój nie obciąża EBITDA ani okresu zwrotu"),
    ("Back margin 2%", "Od całego GMV, także od nawozów GA", "Od GMV bez nawozów GA", "GA nie płaci sama sobie"),
    ("Prowizja od nawozów", "1-1,4% od nawozów i maszyn", "Bez zmian; płaci sprzedawca / dystrybutor. W MVP1 jako opłata od zrealizowanego zapytania", "MVP1 to zapytanie ofertowe, nie koszyk (OPZ WB-17, PZ-18)"),
    ("Premium", "40 zł/ha (≈ 659 zł/rok)", "Agresywny 40 zł/ha; bazowy i ostrożny 20 zł/ha", "Benchmark SatAgro 10-15 zł/ha"),
    ("Korzyść GA", "Tylko 0,5-1,2% marży z danych (≈ 5 mln zł w roku 5)", "Plus marża z nowej sprzedaży: 20% nawozów GA na platformie × 25% marży", "Główna wartość dla GA to udział w rynku (Nutrien, Syngenta MAP)"),
    ("Zespół i utrzymanie", "15 → 80 FTE, 2,4 → 12 mln zł (jak firma samodzielna)", "Ten sam koszt × synergie: zespół −30%, utrzymanie −25%", "Partnerzy mają istniejące zespoły i platformy; dystrybutorzy GA odpowiadają na zapytania"),
    ("CAC i antichurn", "250-350 zł, 15% churn × 350 zł", "Te same stawki −40% dzięki kanałom GA", "66 dystrybutorów, KAM, nawozy.eu, ARiMR i marka GA jako tani kanał"),
    ("Kto ponosi koszty", "Nie określono", "Pula po koszcie; deficyt PT 70% / PM 30%; GA 0", "Wymóg: GA nie płaci za platformę")])

BM, BM_last = text_sheet("Case study i benchmarki", "Pozytywne case study i benchmarki (źródła publiczne, 11.2026)", "Co działa u innych i co bierzemy do GAia.",
    ["Firma / współpraca", "Kto co wnosi", "Wynik (liczby)", "Jak się dzielą / zarabiają", "Co bierzemy do GAia", "Źródło"], [26, 32, 46, 34, 40, 44], [
    ("Nutrien Digital", "Producent i dystrybutor nawozów, własny kanał cyfrowy", ">1,2 mld USD sprzedaży cyfrowej w 2020 (4× r/r); ok. 19% sprzedaży w I poł. 2021", "Kanał sprzedaży własnych produktów", "Bazowy scenariusz ≈ tempo Nutrien", "https://www.businesswire.com/news/home/20210217006038/en"),
    ("Syngenta MAP (Chiny)", "Producent + centra usług + aplikacja", "ok. 1,8 mld USD przychodu w 2021 (2×); 638 centrów; >2 mln rolników", "Produkty i usługi przez platformę i sieć", "Platforma + sieć fizyczna GA", "https://www.syngentagroup.com/newsroom/2022/syngenta-group-reports-2021-performance-growing-23-282-billion-sales"),
    ("Starbucks × PepsiCo (NACP)", "Marka i produkt + dystrybucja", "JV 50/50 od 1994; stworzyli rynek kawy gotowej do picia w USA", "Wspólny zysk, każdy w swojej roli", "Długi horyzont, jasne role", "https://contact.pepsico.com/starbucks/about-us"),
    ("Bosch × BASF (ONE SMART SPRAY)", "Technologia + agronomia", "Współpraca 2019, JV 50/50 2021; system w maszynach Amazone, AGCO, CNH, Stara", "Wspólna spółka po etapie współpracy", "Etap 1 alians, etap 2 spółka", "https://bosch-presse.de/pressportal/de/en/first-crop-protection-sprayer-featuring-one-smart-sprays-weed-management-system-launched-in-latin-america-by-stara-265929.html"),
    ("Bayer × Microsoft", "Modele agronomiczne + chmura", "Partnerstwo od 2021; usługi na Azure od 2023", "Licencje na wspólnej platformie", "Silniki partnerów na licencji; Azure zgodny z OPZ", "https://www.businesswire.com/news/home/20230314005312/en/"),
    ("Yara × Agrofy", "Producent nawozów + marketplace", "Runda 30 mln USD prowadzona przez Yarę (2021)", "Prowizje + fintech", "Producent i marketplace razem", "https://agfundernews.com/agrofy-adds-30m-in-yara-led-series-c-round"),
    ("Sklepy z aplikacjami / Salesforce", "Platforma + twórcy aplikacji", "70/30 (Apple, Google); 15% dla platformy (Salesforce)", "Większość dla twórcy wartości", "Premium: 70% dla PT", "https://developer.salesforce.com/docs/atlas.en-us.252.0.packagingGuide.meta/packagingGuide/appexchange_checkout_rev_share.htm"),
    ("Allegro", "Marketplace", "Prowizja 12,0%, reklama ok. 2,0% GMV (IV kw. 2024)", "Prowizje i reklama", "Model GA: 3,9-5,5% GMV (ostrożnie)", "https://about.allegro.eu/wp-content/uploads/2026/01/Results-Presentation-Q4-2024.pdf"),
    ("Lekcja: Agrando", "Sam marketplace bez producenta", "Zamknięty w 11.2022", "Same prowizje to za mało", "GAia ma producenta i logistykę", "https://agfundernews.com/agrando-raises-14m-in-europes-biggest-farmer-marketplace-series-a")])
