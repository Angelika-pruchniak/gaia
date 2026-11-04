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
    "Oparty na praktykach JV (Umbrex, Ankura, WIPO) i przykładach firm, które łączą zasoby zamiast pieniędzy: Bushel, xFarm × Yara, Żabka × AiFi, Yara × Oxbury, Bosch × BASF.",
    ["Obszar", "Zasada", "Dlaczego (benchmark / praktyka)"], [30, 90, 70], [
    ("Struktura", "Etap 1 (lata 0-2): umowa partnerska (konsorcjum) ze wspólną pulą przychodów na rachunku powierniczym. Etap 2 (od roku 3, po bramce KPI): spółka celowa GAia sp. z o.o.", "Bosch i BASF: współpraca od 2019, spółka w 2021. Mniejsze ryzyko i koszt na starcie"),
    ("Okres", "10 lat z automatycznym przedłużeniem o 5 lat. Wypowiedzenie bez przyczyny najwcześniej po 5 latach, z 12-miesięcznym okresem przejściowym", "Partner technologiczny potrzebuje lat na zwrot wkładu pracy (Bushel i spółdzielnie: wieloletnie umowy white-label)"),
    ("Wkład GA", "Produkty i produkcja, logistyka i punkty wydania, 66 dystrybutorów, marka i marketing, nawozy.eu, współpraca z ARiMR i Węzłem Krajowym, 5 funkcjonalności na Azure (audyt, licencja). Bez gotówki", "Wkład rzeczowy za udział (Ankura, Umbrex)"),
    ("Wkład Partnera Technologicznego (SH)", "Silnik platformy na licencji dla GAia, warstwa UX (web + mobile), grywalizacja, integracje (ARiMR, Węzeł Krajowy, SAP, banki, plug-in marketplace'u), utrzymanie i bezpieczeństwo zgodnie z OPZ MVP1", "IP wnoszone licencją (WIPO); xFarm × Yara: wiedza producenta w silniku partnera"),
    ("Wkład Partnera Marketplace", "Gotowy marketplace podpięty przez plug-in: katalog, zapytania ofertowe, oferty dystrybutorów, płatności, sprzedawcy, reklama; operacje marketplace'u", "Żabka × AiFi: korporacja daje rynek, partner technologię"),
    ("Pula przychodów", "Wszystkie przychody GAia (prowizje, back margin, premium, opłata za sukces GA) wpływają na wspólny rachunek. Z puli: koszty wspólne po koszcie, bez marży, w ramach rocznego budżetu", "Nikt nie zarabia na kosztach: brak konfliktu interesów"),
    ("Podział", "Po zwrocie deficytu: podział według klucza dla każdego strumienia (arkusz Klucz podziału). Przegląd klucza co 2 lata według rzeczywistego wkładu", "Sklepy z aplikacjami 70/30, Salesforce 85/15: większość dla tego, kto tworzy wartość"),
    ("GA nie płaci", "GA nie wnosi gotówki i nie dopłaca deficytu. Płaci tylko 25% DODATKOWEJ marży z platformy (opłata za sukces). Prowizje od nawozów płacą sprzedawcy / dystrybutorzy", "Model afiliacyjny: płacisz od efektu"),
    ("Finansowanie startu", "Deficyt puli finansują Partner Technologiczny (70%) i Partner Marketplace (30%). Wraca do nich jako pierwszy z zysku. Zmniejszamy go grantami (FENG SMART, Horizon Europe, EIT Food) i partnerem finansowym (bank finansuje zakupy rolników jak Oxbury u Yary)", "Preferencyjny zwrot dla wspólnika finansującego"),
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

BM, BM_last = text_sheet("Case study i benchmarki", "Case study: firma z produktem i klientami + software house / startup (bez wielkich pieniędzy)", "Wzorce dopasowane do GAia: każdy wnosi zasoby, nie gotówkę. Źródła publiczne, 11.2026.",
    ["Współpraca", "Kto co wnosi", "Wynik (liczby)", "Jak się rozliczają / bez czyich pieniędzy", "Co bierzemy do GAia", "Źródło"], [26, 34, 44, 36, 40, 44], [
    ("Bushel (USA)", "Startup: aplikacja white-label; spółdzielnie i punkty skupu: marka, klienci, dane", "Start: 10 punktów pilotażowych; ok. 1 500 punktów skupu w 43 firmach; przejął GrainBridge (JV ADM-Cargill)", "Aplikacja darmowa dla rolnika, pod marką spółdzielni; startup zarabia na platformie", "Aplikacja pod marką GA dla sieci 66 dystrybutorów", "https://www.feedstuffs.com/grain-market-news/bushel-acquires-adm-cargill-joint-venture-grainbridge"),
    ("xFarm × Yara (2023)", "Startup xFarm: gotowa platforma; Yara: wiedza o nawożeniu wbudowana w moduł", "xFarm: 300 tys. gospodarstw, 3 mln ha, ponad 100 krajów", "Bez nowej spółki: wiedza producenta w silniku startupu", "5 funkcjonalności GA i nawozy.eu w silniku SH", "https://www.agritechtomorrow.com/news/2023/11/15/yara-international-and-xfarm-technologies-partner-to-advance-sustainable-farming-across-europe/15070"),
    ("Żabka × AiFi (Żabka Nano, Polska)", "Żabka: koncept, lokalizacje, sprzedaż (inkubator Żabka Future); AiFi: technologia", "Pierwszy sklep 06.2021, 25 sklepów w 6 miesięcy, ponad 50 w 2023", "Korporacja nie budowała technologii; startup skalował się z jednym partnerem", "GA daje rynek, partner technologię", "https://prnewswire.com/news-releases/abka-nano-autonomous-stores-powered-by-aifi-outpace-industry-with-25-live-locations-301462392.html"),
    ("Yara × Oxbury (UK, 2023)", "Producent nawozów + bank dla rolników", "Linia kredytowa dla rolników kupujących bezpośrednio: nawozy, nasiona, ŚOR", "Pieniądze daje bank, nie producent", "Partner finansowy w koszyku GAia (WB-22)", "https://agritradenews.co.uk/news/2023/09/11/yara-and-oxbury-partner-to-service-direct-farm-buyers/"),
    ("PepsiCo × Yara (2024)", "Nawozy niskoemisyjne + narzędzia cyfrowe + agronomia dla dostawców PepsiCo", "Do 165 tys. t nawozów rocznie; ok. 1 000 gospodarstw, 128 tys. ha", "Przetwórca płaci za niższy ślad węglowy upraw", "Programy niskoemisyjne GA z przetwórcami przez GAia", "https://www.pepsico.com/newsroom/press-releases/2024/pepsico-europe-and-yara-partner-to-decarbonize-crop-production"),
    ("Bosch × BASF (ONE SMART SPRAY)", "Technologia + agronomia, każdy wnosi swoją", "Współpraca od 2019, spółka 50/50 od 2021; system w maszynach Amazone, AGCO, CNH, Stara", "Wkład technologiczny za udziały", "Najpierw współpraca, potem spółka", "https://bosch-presse.de/pressportal/de/en/first-crop-protection-sprayer-featuring-one-smart-sprays-weed-management-system-launched-in-latin-america-by-stara-265929.html"),
    ("Venture studia w Polsce (Daftcode, The Heart)", "Software house / studio wnosi zespół; korporacja: rynek i klientów", "The Heart tworzy nowe produkty i biznesy na zamówienie z korporacjami", "Praca za udziały lub udział w przychodach", "SH jako współwłaściciel GAia, nie dostawca na godziny", "https://itwiz.pl/czy-fundusz-venture-capital-to-dobry-partner-dla-firmy-programistycznej/"),
    ("Producent z kanałem cyfrowym: Nutrien", "Producent i dystrybutor nawozów", ">1,2 mld USD sprzedaży cyfrowej w 2020 (4× r/r)", "Kanał sprzedaży własnych produktów", "Skala, do której dąży GAia", "https://www.businesswire.com/news/home/20210217006038/en"),
    ("Lekcja: Agrando (Niemcy)", "Sam marketplace bez producenta", "Zamknięty w 11.2022", "Same prowizje to za mało", "GAia ma producenta i logistykę", "https://agfundernews.com/agrando-raises-14m-in-europes-biggest-farmer-marketplace-series-a")])

FD, FD_last = text_sheet("Dofinansowania", "Dofinansowania i dekarbonizacja: skąd pieniądze na GAia, żeby nikt nie wykładał milionów", "Stan wiedzy 11.2026, do weryfikacji z doradcą grantowym. Terminy naborów się zmieniają.",
    ["Źródło", "Na co", "Warunki / kwoty", "Status / termin", "Jak GAia się kwalifikuje", "Źródło"], [26, 34, 40, 30, 46, 44], [
    ("FENG Ścieżka SMART (NCBR)", "B+R, wdrożenie innowacji, moduły: cyfryzacja, zazielenienie", "Duża firma tylko w konsorcjum z MŚP; projekt do 150 mln zł; min. 3 mln zł kosztów B+R / wdrożenia", "Nabór 2026 zamknięty (22.05.2026); monitorować kolejny", "GA lider + SH (MŚP) w konsorcjum: model AI rekomendacji nawożenia, moduł zazielenienia", "https://www.gov.pl/web/ncbr/aktualizacja-dokumentacji-FENG-sciezka-SMART-dla-duzych-przedsiebiorstw-nabor-FENG01-01-IP01-001-26"),
    ("Horizon Europe, Klaster 6", "Cyfrowe i AI narzędzia dla nawożenia i wody", "HORIZON-CL6-2027-02-FARM2FORK-01: ok. 6 mln EUR na projekt, budżet 12 mln EUR", "Termin 23.09.2027", "Konsorcjum: GA + SH + IUNG / IOR-PIB; pilot GAia jako demonstrator", "https://www.horizon-europe.gouv.fr/increasing-resilience-agriculture-water-and-nutrient-scarce-environments-through-digital"),
    ("Digital Europe: AgrifoodTEF", "Testowanie i walidacja AI w rolnictwie w warunkach polowych", "Usługi testowe, Polska jest węzłem satelitarnym", "Działa 2023-2027", "Walidacja modelu AI GAia przed wdrożeniem", "https://www.fbk.eu/en/result/fbk-coordina-a-livello-europeo-agrifoodtef/"),
    ("Digital Europe: otwarte nabory typu GATE 5.0", "Para: MŚP agro z potrzebą + MŚP technologiczne z rozwiązaniem", "Dofinansowanie wdrożeń dla par MŚP", "Ostatni termin 30.04.2026; obserwować kolejne", "SH + dystrybutor / spółka GA o statusie MŚP", "https://grantedai.com/grants/digital-europe-programme-calls-for-applications-from-joint-consortia-con-european-commission-digital-europe-progr-f5496d7c"),
    ("EIT Food", "Piloty startupów z korporacjami, skalowanie", "Accelerator Network: grant bez udziałów; dawniej do 500 tys. EUR (dane z 2023)", "Nabory co roku", "SH lub spółka GAia jako startup z GA jako partnerem pilotażu", "https://www.eitfood.eu/news/eit-food-announces-the-2026-risingfoodstars-verified-scaleups-recognising-europes-most-promising-agrifood-scaleups"),
    ("Innovation Fund (UE)", "Dekarbonizacja produkcji (np. amoniak niskoemisyjny)", "Nabór 2025: 77 projektów, 3,4 mld EUR", "Kolejne nabory co roku", "Dla zakładów GA; GAia buduje popyt i dane o zielonych nawozach", "https://www.lovguiden.dk/en/det-offentlige/europa-kommissionen-press-corner/2026-10-19-innovationsfonden-stoetter-77-projekter-kkr-faar-klagepunkter"),
    ("NFOŚiGW / Fundusz Modernizacyjny / KPO", "Redukcja emisji w przemyśle", "Programy priorytetowe NFOŚiGW", "Nabory okresowe", "Dla produkcji GA; GAia jako dowód redukcji emisji u odbiorców", "https://www.gramwzielone.pl/?p=104924"),
    ("Ekoschemat: plan nawożenia (dla rolnika)", "Płatność za opracowanie i przestrzeganie planu nawożenia", "ok. 87,52 zł/ha wariant podstawowy, 262,56 zł/ha z wapnowaniem (2025, dane wtórne)", "Co roku w kampanii ARiMR", "GAia tworzy plan nawożenia: rolnik dostaje dopłatę, GAia zyskuje użytkownika", "https://wiescirolnicze.pl/ekoschemat-rolnictwo-weglowe-i-zarzadzanie-skladnikami-odzywczymi-omawiamy"),
    ("CBAM (od 01.01.2026)", "Opłata węglowa na import nawozów do UE", "Rabobank: amoniak +10-20%, mocznik +10-15% w 2026", "Faza definitywna od 2026", "Argument cenowy za nawozami GA (UE, niższy ślad) i za danymi o śladzie z GAia", "https://www.rabobank.com/knowledge/q011499773-carbon-costs-at-the-gate-preparing-for-the-cbams-impact-on-eu-fertilizer-prices-in-2026")])

DK, DK_last = text_sheet("Dekarbonizacja", "Jak GAia wpływa na dekarbonizację (argument do dofinansowań i dla przetwórców)", "Liczby przykładowe do potwierdzenia przez IUNG / dział zrównoważonego rozwoju GA.",
    ["Mechanizm", "Jak działa w GAia", "Dowód / benchmark", "Wskaźnik do mierzenia"], [30, 56, 60, 40], [
    ("Precyzyjne nawożenie (model AI)", "Plan nawożenia z paszportu pola, NDVI i danych ARiMR; mniej azotu tam, gdzie nie jest potrzebny", "Badania: 5-16% mniej N2O; narzędzie Nutrient Expert: 12-20% niższy ślad w pszenicy (Scientific Reports)", "kg N/ha, t CO2e na gospodarstwo"),
    ("Więcej nawozów niskoemisyjnych", "Rekomendacje i promocje produktów z inhibitorem (np. Pulrea +INu z NBPT) w koszyku", "GA: cel −9% śladu węglowego do 2030 vs 2024", "Udział nawozów niskoemisyjnych w sprzedaży przez GAia"),
    ("Dane MRV (pomiar, raport, weryfikacja)", "Historia nawożenia i upraw w paszporcie pola, zgody rolnika", "PepsiCo × Yara: przetwórca płaci za niższy ślad upraw", "Raport śladu na gospodarstwo / dostawcę"),
    ("Ekoschemat plan nawożenia", "Plan z GAia spełnia wymóg ekoschematu", "ok. 87,52 zł/ha (2025)", "Liczba rolników z dopłatą dzięki GAia"),
    ("Przykład skali", "1 mln ha × 150 kg N/ha × 5% mniej = 7,5 tys. t N rocznie", "ZAŁOŻENIE: ok. 8 kg CO2e na 1 kg N (produkcja + pole) → ok. 60 tys. t CO2e rocznie", "Do policzenia z IUNG na danych pilota")])
