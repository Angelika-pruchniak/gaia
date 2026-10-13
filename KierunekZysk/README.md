# Kierunek Zysk: strategia wzrostu zysku dla małych firm usługowych

Produkt cyfrowy dla **guliwera4.pl**. Pakiet 7 materiałów, który prowadzi właściciela firmy usługowej (2–50 osób) od diagnozy liczb do uruchomionego testu sprzedażowego w 30 dni, a potem do płatnej konsultacji **Przegląd Decyzji (3 000 zł netto)**.

## Zawartość pakietu

| # | Plik | Format | Styl | Stron |
|---|------|--------|------|-------|
| 1 | `PDF/1_Kierunek_Zysk_Przewodnik.pdf`: przewodnik + zeszyt ćwiczeń (8 modułów) | PDF A4 | pomarańcz + pudrowy róż | 24 |
| 2 | `2_Kierunek_Zysk_Kalkulator.xlsx`: rentowność usług i klientów, macierz, symulator podwyżki, lejek, dźwignie, próg rentowności, **podsumowanie do konsultacji** | Excel | pastel + granat | 9 arkuszy |
| 3 | `PDF/3_Strategia_na_1_stronie.pdf`: kanwa + przykład wypełnienia | PDF A4 poziomo | granatowo-biały | 2 |
| 4 | `PDF/4_Planer_90_dni.pdf`: priorytety, mapa 13 tygodni, tablica wyników, karty testów, przeglądy | PDF A4 | pomarańcz + róż | 9 |
| 5 | `PDF/5_Skrypty_i_szablony_sprzedazowe.pdf`: 16 rozmów, maili i wiadomości | PDF A4 | granatowo-biały | 9 |
| 6 | `PDF/6_Audyt_i_bank_40_testow.pdf`: audyt 30 punktów + 40 testów sprzedażowych | PDF A4 | pomarańcz + róż | 7 |
| 7 | `PDF/7_Dziennik_decyzji_i_konsultacja.pdf`: 5 kart decyzji (tok rozumowania, 6 ogniw jakości decyzji, pre-mortem), autotest 8 pułapek myślenia, oferta konsultacji, karta przygotowania, rewizja po 30/90 dniach | PDF A4 | pomarańcz + róż | 12 |

Gotowa paczka do sklepu: `Kierunek_Zysk_pakiet.zip` (materiały 1–7, bez folderu `Dla_prowadzacej`).

## Moduły przewodnika

1. **Diagnoza**: 12 liczb, Pareto klientów i usług, macierz zysk × wysiłek, test zależności od właściciela
2. **Gdzie grać**: ocena segmentów, karta idealnego klienta (Jobs to be Done), usługa flagowa
3. **Jak wygrać**: równanie wartości, mapa zmian w ofercie ERRC, zdanie wartości, oferta flagowa, gwarancje
4. **Ceny i rentowność**: dźwignia ceny, ilu klientów możesz stracić po podwyżce, 7 zasad wyceny, pakiety Dobry/Lepszy/Najlepszy, plan podwyżki
5. **Lejek i źródła klientów**: 5 dźwigni wzrostu, lejek dziś vs za 90 dni, ocena źródeł, system poleceń, CRM w wersji minimum
6. **Trzy horyzonty**: podział 70/20/10 w małej firmie
7. **Plan 90 dni**: od pomysłów do 3 priorytetów, cele z miernikami, karta testu sprzedażowego
8. **Rytm wdrożenia**: przeglądy tygodniowe, miesięczne i kwartalne, uwalnianie właściciela ze sprzedaży

Przez cały przewodnik przewija się przykład fikcyjnej firmy „Studio Lumen” (6 osób, 1,2 mln zł przychodu, marża 8%). Liczby w przewodniku, kanwie i kalkulatorze są ze sobą spójne.

## Konsultacja Przegląd Decyzji (3 000 zł netto)

Produkt cyfrowy jest płatnym wejściem do konsultacji. Klient przez 30 dni zapisuje w **Dzienniku decyzji** nie tylko co postanowił, ale *dlaczego*. Na konsultacji prowadząca analizuje ten tok rozumowania według 6 ogniw jakości decyzji (rama, opcje, informacje, kryteria, rozumowanie, zobowiązanie) i razem z klientem wzmacnia najsłabsze ogniwa.

**Wynik dla klienta:** mapa decyzji (zostaw / popraw / wstrzymaj), ocena 6 ogniw, lista założeń do sprawdzenia, wzorzec decyzyjny i nawyk na przyszłość, 3 rekomendacje, plan na 30 dni, raport na piśmie, 14 dni wsparcia mailowego, przegląd po 30 dniach.

**Zaproszenie do konsultacji jest wplecione w produkt:** przewodnik (start, moduły 2/4/5/7, strona „Następny krok”), Dziennik decyzji (oferta + karta przygotowania), Kalkulator (arkusz „Podsumowanie”), Planer (podsumowanie 90 dni) oraz 6 maili po zakupie.

### Folder `Dla_prowadzacej/` (nie dołączać do pakietu klienta)

| Plik | Do czego |
|---|---|
| `Pakiet_prowadzacej_konsultacje.pdf` | Model produkt → konsultacja, oferta i zasady, tekst na stronę /konsultacja, 6 maili po zakupie (do Mailchimp), formularz kwalifikacji, analiza przed sesją, kryteria oceny 6 ogniw, sygnały ostrzegawcze, pytania odsłaniające tok rozumowania, scenariusz sesji 120 min, raport, follow-up |
| `Karta_analizy_konsultacji.xlsx` | Narzędzie pracy na 1 klienta: ocena ogniw (wskazuje najsłabsze ogniwo i kolejność omawiania), założenia, sygnały, pytania, hipoteza, werdykty; arkusz **Raport** składa się sam i drukuje do PDF dla klienta. Zawiera przykład „Studio Lumen”, który należy usunąć przed użyciem |

### Do ustawienia przed startem
- Strona **guliwera4.pl/konsultacja** z kalendarzem i płatnością (adres jest już wpisany w materiałach; jeśli będzie inny, podmień go w `src/*.html` i `build/build_xlsx.py`, potem przebuduj).
- Automatyzacja 6 maili po zakupie (teksty w pakiecie prowadzącej, rozdz. 3).
- Regulamin konsultacji, faktura/VAT, RODO: do potwierdzenia z księgową i prawnikiem.

## Wykorzystana wiedza (źródła)

- McKinsey Quarterly, „The power of pricing” (Marn, Roegner, Zawada): wzrost ceny o 1% to ok. 8% zysku operacyjnego w przeciętnej spółce S&P 1500
- „The Alchemy of Growth” (Baghai, Coley, White): trzy horyzonty
- „The Granularity of Growth” (Viguerie, Smit, Baghai): wzrost w mikrosegmentach
- Kim & Mauborgne, „Strategia błękitnego oceanu”: ERRC
- A. Hormozi, „$100M Offers”: równanie wartości, oferta, gwarancje
- N. Rackham, „SPIN Selling”: rozmowa odkrywcza
- C. Christensen: Jobs to be Done; D. McClure: lejek AARRR
- C. Spetzler, H. Winter, J. Meyer, „Decision Quality”: 6 ogniw jakości decyzji; G. Klein: pre-mortem; A. Duke, „Thinking in Bets”
- PARP (raport o sektorze MŚP) i PIE (MIK, wrzesień 2026): kontekst polskiego rynku

Modele zostały opisane własnymi słowami i dostosowane do małych firm. Materiał nie używa marki McKinsey w nazwie produktu ani nie sugeruje współpracy (zapis w dodatku przewodnika).

## Opis do sklepu (propozycja)

**Kierunek Zysk: co konkretnie zmienić, żeby Twoja firma zarabiała więcej?**

Masz klientów i pełny kalendarz, ale na koncie wciąż za mało? Kierunek Zysk to praktyczny pakiet dla właścicieli małych firm usługowych, którzy chcą zarabiać więcej, a nie tylko więcej pracować.

W 30 dni, poświęcając 2–3 godziny tygodniowo:
- sprawdzisz, którzy klienci i które usługi naprawdę zarabiają, a które dokładają do firmy,
- ustalisz nowe ceny i pakiety (z kalkulatorem, który pokaże, ilu klientów możesz stracić po podwyżce i nadal zarabiać więcej),
- wybierzesz klientów i usługę, na których się skupisz,
- uporządkujesz sprzedaż dzięki 16 gotowym skryptom: od pierwszego maila po list o podwyżce cen,
- zbudujesz plan na 90 dni i uruchomisz pierwszy test sprzedażowy.

Metoda opiera się na sprawdzonych modelach strategicznych i sprzedażowych używanych przez duże firmy doradcze i najlepszych praktyków sprzedaży. Przełożyliśmy je na realia firmy z 2–50 pracownikami.

**W pakiecie:** przewodnik z ćwiczeniami (24 str.) · Kalkulator Zysku (Excel) · Strategia na 1 stronie · Planer 90 dni · 16 skryptów sprzedażowych · audyt 30 punktów + 40 pomysłów na testy.

**Dla kogo:** firmy usługowe (B2B i B2C), które działają co najmniej 2 lata, mają stałych klientów, a właściciel nadal sam sprzedaje.
**Nie dla:** startupów przed pierwszym klientem i firm, które szukają tylko reklam albo baz kontaktów.

### Proponowane warianty cenowe (hipoteza do testu)

| Wariant | Zawartość | Cena |
|---|---|---|
| **Pakiet cyfrowy** | 6 materiałów | 297 zł (na premierę lub Black Friday 27.11: 197 zł) |
| **Pakiet + konsultacja** | jak wyżej + 2 spotkania grupowe Q&A online | 697 zł |
| **Przegląd Decyzji** | analiza toku rozumowania, sesja 2 h, raport, 14 dni wsparcia, przegląd po 30 dniach | 3 000 zł netto (cena pakietu odliczana przez 60 dni) |

## Jak przebudować pliki

```bash
cd KierunekZysk/build
NODE_PATH=$(npm root -g) node build_pdf.js   # HTML z ../src → PDF/ (Playwright + Chromium)
python3 build_xlsx.py                         # → 2_Kierunek_Zysk_Kalkulator.xlsx (openpyxl)
python3 build_xlsx_konsultant.py              # → Dla_prowadzacej/Karta_analizy_konsultacji.xlsx
```

Kolory (w `src/style.css`): pastelowy pomarańcz `#F8B68C`, brzoskwinia `#FDDCC6`, pudrowy róż `#F2BFC5`, jasny róż `#FBE6E7`, granat `#1E2A4A`. Czcionka: Inter.
