import type { Item, RawExercise, Teach, Unit } from './types'

// --- pomocnicze konstruktory (pierwsza opcja = poprawna) ---
const c = (q: string, options: string[], explain?: string): RawExercise => ({ kind: 'choice', q, options, explain })
const tf = (q: string, answer: boolean, explain?: string): RawExercise => ({ kind: 'tf', q, answer, explain })
const fill = (q: string, options: string[], explain?: string): RawExercise => ({ kind: 'fill', q, options, explain })
const match = (q: string, pairs: [string, string][], explain?: string): RawExercise => ({ kind: 'match', q, pairs, explain })
const order = (q: string, tokens: string[], explain?: string): RawExercise => ({ kind: 'order', q, tokens, explain })
const sort = (q: string, buckets: string[], items: [string, number][], explain?: string): RawExercise => ({
  kind: 'sort', q, buckets, items, explain,
})
const teach = (title: string, body: string[], formula?: string, example?: string): Teach => ({
  kind: 'teach', title, body, formula, example,
})

type RawLesson = { id: string; title: string; items: (RawExercise | Teach)[] }
type RawUnit = Omit<Unit, 'lessons'> & { lessons: RawLesson[] }

const raw: RawUnit[] = [
  // ───────────────────────── 1 ─────────────────────────
  {
    id: 'u1', title: 'Język pieniędzy', subtitle: 'Przychód, zysk, EBITDA, OPEX i CAPEX', color: 'coral',
    lessons: [
      {
        id: 'u1l1', title: 'Przychód, koszt, zysk',
        items: [
          teach('Trzy słowa, które musisz znać', [
            'Przychód (ang. revenue) to pieniądze ze sprzedaży, zanim cokolwiek odejmiesz.',
            'Zysk (ang. profit) to to, co zostaje po odjęciu kosztów.',
            'Marża pokazuje zysk jako procent przychodu. Dzięki niej porównasz małą kawiarnię z siecią.',
          ], 'Zysk = przychód − koszty\nMarża = zysk ÷ przychód', 'Kawiarnia sprzedała kawy za 10 000 zł, koszty wyniosły 7 000 zł. Zysk: 3 000 zł, marża: 30%.'),
          c('Kawiarnia: przychody 20 000 zł, koszty 15 000 zł. Ile wynosi zysk?', ['5 000 zł', '35 000 zł', '15 000 zł', '20 000 zł'], 'Zysk = 20 000 − 15 000 = 5 000 zł.'),
          c('Jaka jest marża zysku w tej kawiarni (zysk 5 000 zł, przychód 20 000 zł)?', ['25%', '75%', '5%', '33%'], '5 000 ÷ 20 000 = 0,25, czyli 25%.'),
          tf('Przychód i zysk to to samo.', false, 'Przychód to wpływy ze sprzedaży. Zysk zostaje dopiero po odjęciu kosztów.'),
          match('Połącz pojęcie z opisem', [
            ['Przychód', 'Wpływy ze sprzedaży'],
            ['Koszt', 'To, co wydajesz, by działać'],
            ['Zysk', 'Przychód minus koszty'],
            ['Marża', 'Zysk jako % przychodu'],
          ]),
          fill('„Revenue” po polsku to ___.', ['przychód', 'zysk', 'koszt', 'marża'], 'Revenue = przychód (czasem „sprzedaż”). Profit = zysk.'),
          c('Marża brutto (gross margin) to zysk po odjęciu od przychodu…', ['kosztu wytworzenia sprzedanych produktów', 'wszystkich kosztów i podatków', 'tylko podatku', 'wynagrodzeń zarządu'], 'Marża brutto = (przychód − koszt sprzedanych towarów, tzw. COGS) ÷ przychód. Pozostałe koszty odejmuje się niżej w rachunku.'),
        ],
      },
      {
        id: 'u1l2', title: 'EBITDA i EBIT',
        items: [
          teach('EBITDA bez tajemnic', [
            'EBITDA to skrót od Earnings Before Interest, Taxes, Depreciation and Amortization, czyli zysk przed odsetkami, podatkami i amortyzacją.',
            'Pokazuje, ile zarabia sama działalność operacyjna. Pomija sposób finansowania (odsetki), podatki i księgowe rozkładanie inwestycji w czasie (amortyzację).',
            'EBIT to zysk operacyjny: EBITDA minus amortyzacja.',
          ], 'EBITDA = EBIT + amortyzacja', 'Przychody 1 000, koszty operacyjne bez amortyzacji 700, czyli EBITDA = 300. Amortyzacja 100, czyli EBIT = 200.'),
          c('Co oznacza litera „D” w skrócie EBITDA?', ['Depreciation, czyli amortyzację środków trwałych', 'Debt, czyli dług', 'Dividends, czyli dywidendy', 'Discount, czyli rabat'], 'D = depreciation, A = amortization. Obie litery to amortyzacja: środków trwałych i wartości niematerialnych.'),
          c('EBIT = 200 tys. zł, amortyzacja = 50 tys. zł. Ile wynosi EBITDA?', ['250 tys. zł', '150 tys. zł', '200 tys. zł', '50 tys. zł'], 'EBITDA = EBIT + amortyzacja = 200 + 50 = 250 tys. zł.'),
          tf('EBITDA uwzględnia koszty odsetek od kredytu.', false, '„I” = interest, czyli odsetki. EBITDA liczymy przed odsetkami, więc ich nie odejmuje.'),
          order('Ułóż poziomy zysku od góry rachunku w dół', ['Przychody', 'EBITDA', 'EBIT', 'Zysk przed opodatkowaniem', 'Zysk netto'], 'Od przychodów odejmujemy koszty operacyjne (EBITDA), potem amortyzację (EBIT), odsetki (zysk przed opodatkowaniem) i podatek (zysk netto).'),
          c('Po co inwestorzy często patrzą na EBITDA?', ['By porównać wyniki operacyjne firm o różnym zadłużeniu i inwestycjach', 'Bo to kwota gotówki na koncie', 'Bo to zysk po podatkach', 'Bo tak każe ustawa'], 'EBITDA ułatwia porównania. Ale to nie gotówka i nie widać w niej potrzeb inwestycyjnych.'),
          tf('Wysoka EBITDA zawsze oznacza, że firma ma dużo gotówki.', false, 'Firma z wysoką EBITDA może dużo wydawać na inwestycje, spłatę długu czy zapasy. Gotówkę widać w przepływach pieniężnych.'),
        ],
      },
      {
        id: 'u1l3', title: 'OPEX i CAPEX',
        items: [
          teach('Wydatek na dziś czy na lata?', [
            'OPEX (operating expenditure) to bieżące wydatki operacyjne: pensje, czynsz, marketing, abonamenty. Od razu trafiają w koszty.',
            'CAPEX (capital expenditure) to wydatki inwestycyjne na aktywa służące latami: maszyny, budynki, serwery. Trafiają do aktywów i stopniowo „schodzą” w koszty przez amortyzację.',
          ], undefined, 'Abonament za chmurę to OPEX. Kupno własnej serwerowni to CAPEX.'),
          sort('Do której szuflady?', ['OPEX', 'CAPEX'], [
            ['Pensje zespołu', 0],
            ['Nowa linia produkcyjna', 1],
            ['Kampania reklamowa', 0],
            ['Budowa magazynu', 1],
            ['Rachunek za prąd', 0],
          ]),
          tf('Przejście z własnych serwerów na chmurę w abonamencie zamienia CAPEX na OPEX.', true, 'Zamiast kupować sprzęt (CAPEX) płacisz co miesiąc (OPEX). Nie zamrażasz gotówki na start.'),
          tf('Zakup maszyny za 2 mln zł (CAPEX) obniża EBITDA w roku zakupu o 2 mln zł.', false, 'CAPEX trafia do aktywów. W koszty wchodzi przez amortyzację, której EBITDA nie odejmuje. Dlatego patrz też na przepływy pieniężne.'),
          fill('Bieżące wydatki na prowadzenie firmy to ___.', ['OPEX', 'CAPEX', 'EBITDA', 'GMV']),
          c('Amortyzacja to…', ['rozłożenie kosztu środka trwałego na lata jego używania', 'rata kredytu', 'podatek od maszyn', 'rabat od dostawcy']),
          c('Firma kupuje linię produkcyjną za 3 mln zł na 10 lat. Jak to zwykle ujmuje?', ['Jako CAPEX: środek trwały amortyzowany przez lata', 'Jako OPEX w całości w miesiącu zakupu', 'Jako przychód', 'Wcale, bo to inwestycja']),
        ],
      },
    ],
  },
  // ───────────────────────── 2 ─────────────────────────
  {
    id: 'u2', title: 'Metryki biznesu', subtitle: 'GMV, AOV, CAC, LTV, churn, MRR', color: 'pink',
    lessons: [
      {
        id: 'u2l1', title: 'GMV, AOV i konwersja',
        items: [
          teach('Liczby każdego sklepu', [
            'GMV (Gross Merchandise Value) to łączna wartość towarów sprzedanych przez platformę, przed odjęciem zwrotów i prowizji. To nie jest przychód platformy!',
            'AOV (Average Order Value) to średnia wartość zamówienia.',
            'Konwersja (CR) to odsetek odwiedzin, które kończą się zakupem.',
          ], 'AOV = przychód ÷ liczba zamówień\nKonwersja = zamówienia ÷ odwiedziny', 'Marketplace z GMV 10 mln zł i prowizją 10% ma przychód ok. 1 mln zł.'),
          c('Marketplace: GMV 50 mln zł, prowizja 8%. Jaki jest przybliżony przychód platformy?', ['4 mln zł', '50 mln zł', '8 mln zł', '46 mln zł'], 'Platforma zarabia na prowizji: 8% × 50 mln = 4 mln zł.'),
          tf('GMV marketplace’u to jego przychód w rachunku zysków i strat.', false, 'GMV to wartość transakcji. Przychód marketplace’u to zwykle prowizje i opłaty.'),
          c('Sklep: 200 zamówień, przychód 30 000 zł. Ile wynosi AOV?', ['150 zł', '6 000 zł', '200 zł', '300 zł'], '30 000 ÷ 200 = 150 zł.'),
          c('10 000 odwiedzin, 250 zamówień. Jaka jest konwersja?', ['2,5%', '25%', '0,25%', '40%'], '250 ÷ 10 000 = 0,025, czyli 2,5%.'),
          match('Połącz skrót z opisem', [
            ['GMV', 'Wartość wszystkich transakcji'],
            ['AOV', 'Średnia wartość zamówienia'],
            ['CR', 'Odsetek wizyt z zakupem'],
          ]),
          fill('Przychód sklepu = ruch × konwersja × ___.', ['AOV', 'CAPEX', 'churn', 'EBITDA'], 'To najprostsze „równanie sklepu”. Żeby rosnąć, zwiększasz ruch, konwersję albo wartość koszyka.'),
        ],
      },
      {
        id: 'u2l2', title: 'CAC i LTV',
        items: [
          teach('Ile kosztuje klient i ile jest wart', [
            'CAC (Customer Acquisition Cost) to koszt pozyskania jednego klienta: wydatki na marketing i sprzedaż podzielone przez liczbę nowych klientów.',
            'LTV (Lifetime Value) to wartość klienta w czasie: ile marży przynosi przez cały okres współpracy.',
            'Zdrowy biznes ma LTV wyraźnie większe od CAC. Często podaje się orientacyjnie proporcję 3:1.',
          ], 'CAC = koszty pozyskania ÷ nowi klienci\nLTV ≈ przychód mies. × marża × liczba miesięcy'),
          c('Wydałaś 20 000 zł na kampanię i pozyskałaś 400 klientów. Ile wynosi CAC?', ['50 zł', '400 zł', '20 zł', '8 000 zł'], '20 000 ÷ 400 = 50 zł.'),
          c('Klient płaci 100 zł/mies., marża 60%, zostaje średnio 10 miesięcy. Ile wynosi LTV (na marży)?', ['600 zł', '1 000 zł', '60 zł', '6 000 zł'], '100 × 60% × 10 = 600 zł.'),
          tf('LTV = 300 zł, CAC = 500 zł. Każdy nowy klient zarabia dla firmy.', false, 'Płacisz 500 zł za klienta, który przyniesie 300 zł marży, więc tracisz 200 zł na każdym.'),
          c('Który ruch najbardziej bezpośrednio podnosi LTV?', ['Zatrzymanie klientów na dłużej', 'Większy budżet reklamowy', 'Nowe logo', 'Więcej spotkań wewnętrznych']),
          match('Połącz skrót z opisem', [
            ['CAC', 'Koszt pozyskania klienta'],
            ['LTV', 'Wartość klienta w czasie'],
            ['Payback', 'Czas zwrotu kosztu pozyskania'],
          ]),
          c('LTV:CAC = 3:1. Co to znaczy?', ['Klient przynosi 3× więcej marży, niż kosztuje jego pozyskanie', 'Pozyskanie klienta kosztuje 3× więcej, niż on przynosi', 'Firma ma 3 klientów', 'Kampania trwa 3 miesiące']),
        ],
      },
      {
        id: 'u2l3', title: 'Churn, MRR i ARR',
        items: [
          teach('Abonamenty i odejścia', [
            'MRR (Monthly Recurring Revenue) to miesięczny powtarzalny przychód z abonamentów. ARR to roczny powtarzalny przychód: MRR × 12.',
            'Churn to odsetek klientów (albo przychodu), który odchodzi w danym okresie.',
            'Mały miesięczny churn szybko się kumuluje. To cichy zabójca firm abonamentowych.',
          ], 'ARR = MRR × 12\nChurn = odeszli ÷ klienci na początku okresu'),
          c('MRR = 50 000 zł. Ile wynosi ARR?', ['600 000 zł', '50 000 zł', '500 000 zł', '4 167 zł']),
          c('Na początku miesiąca 1 000 klientów, odeszło 30. Jaki jest churn miesięczny?', ['3%', '30%', '0,3%', '97%']),
          tf('Churn 5% miesięcznie to drobiazg, bo to tylko 5%.', false, 'Po roku zostaje ok. 0,95¹² ≈ 54% klientów, więc prawie połowa odchodzi.'),
          c('Churn wynosi 5% miesięcznie. Ile średnio zostaje klient?', ['ok. 20 miesięcy', 'ok. 5 miesięcy', 'ok. 12 miesięcy', 'ok. 95 miesięcy'], 'Przybliżenie: 1 ÷ churn = 1 ÷ 0,05 = 20 miesięcy.'),
          sort('Co podnosi, a co obniża MRR?', ['Podnosi MRR', 'Obniża MRR'], [
            ['Nowy klient na abonamencie', 0],
            ['Klient rezygnuje', 1],
            ['Klient przechodzi na wyższy pakiet', 0],
            ['Klient przechodzi na tańszy pakiet', 1],
          ]),
          fill('Odsetek klientów, którzy rezygnują z usługi, to ___.', ['churn', 'upsell', 'CAPEX', 'GMV']),
        ],
      },
    ],
  },
  // ───────────────────────── 3 ─────────────────────────
  {
    id: 'u3', title: 'Ekonomia na co dzień', subtitle: 'Popyt, podaż, elastyczność, koszty', color: 'lilac',
    lessons: [
      {
        id: 'u3l1', title: 'Popyt i podaż',
        items: [
          teach('Dwie siły rynku', [
            'Popyt to ilość, jaką klienci chcą kupić przy danej cenie. Zwykle im wyżej cena, tym mniej chętnych.',
            'Podaż to ilość, jaką sprzedawcy chcą sprzedać. Im wyżej cena, tym więcej chcą sprzedać.',
            'Cena równowagi to punkt, w którym popyt i podaż się spotykają.',
          ], undefined, 'Przed świętami popyt na karpie rośnie, więc cena też, dopóki hodowcy nie dowiozą większej podaży.'),
          c('Ceny biletów na koncert rosną, a reszta się nie zmienia. Co zwykle dzieje się z liczbą kupowanych biletów?', ['Spada', 'Rośnie', 'Nie zmienia się', 'Podwaja się']),
          tf('Gdy popyt rośnie szybciej niż podaż, ceny mają tendencję do wzrostu.', true),
          c('Na rynku jest nadwyżka towaru (podaż większa niż popyt). Do czego to zwykle prowadzi?', ['Do spadku cen', 'Do wzrostu cen', 'Do braku zmian', 'Do zakazu sprzedaży']),
          c('Dobrem substytucyjnym dla kawy jest np.…', ['herbata', 'cukier', 'filiżanka', 'ekspres'], 'Substytut zastępuje produkt. Cukier czy ekspres to dobra komplementarne, czyli kupowane razem.'),
          sort('Substytuty czy komplementy?', ['Substytuty', 'Komplementy'], [
            ['Masło i margaryna', 0],
            ['Drukarka i tusz', 1],
            ['Pociąg i autobus', 0],
            ['Samochód i paliwo', 1],
          ]),
        ],
      },
      {
        id: 'u3l2', title: 'Elastyczność i koszt alternatywny',
        items: [
          teach('Jak mocno klienci reagują na cenę?', [
            'Elastyczność cenowa popytu mówi, jak bardzo zmienia się sprzedaż, gdy zmienia się cena.',
            'Popyt elastyczny: mała podwyżka, duży spadek sprzedaży. Popyt nieelastyczny: ludzie kupują mimo podwyżki (leki, prąd).',
            'Koszt alternatywny to wartość najlepszej opcji, z której rezygnujesz, wybierając inną.',
          ]),
          c('Który produkt ma zwykle najmniej elastyczny popyt?', ['Insulina', 'Bilety do kina', 'Markowe batoniki', 'Wakacje za granicą']),
          c('Podnosisz cenę o 10%, sprzedaż spada o 30%. Popyt jest…', ['elastyczny', 'nieelastyczny', 'całkowicie sztywny', 'równy zero'], 'Ilość spadła mocniej (30%) niż wzrosła cena (10%), więc popyt jest elastyczny.'),
          tf('Przy nieelastycznym popycie podwyżka ceny zwykle zwiększa przychód.', true, 'Sprzedajesz niewiele mniej, ale drożej, więc przychód rośnie.'),
          c('Zamiast zlecenia za 800 zł idziesz w sobotę na kurs. Co jest kosztem alternatywnym kursu?', ['Utracone 800 zł ze zlecenia', 'Nic, bo to sobota', 'Tylko kawa w przerwie', 'Koszt dojazdu sąsiada']),
          fill('Wartość najlepszej odrzuconej opcji to koszt ___.', ['alternatywny', 'stały', 'krańcowy', 'utopiony']),
        ],
      },
      {
        id: 'u3l3', title: 'Koszty i próg rentowności',
        items: [
          teach('Kiedy zaczynasz zarabiać?', [
            'Koszty stałe nie zależą od wielkości sprzedaży (czynsz, pensje). Koszty zmienne rosną z każdą sprzedaną sztuką (surowiec, prowizja kuriera).',
            'Marża pokrycia na sztukę to cena minus koszt zmienny. Z niej „spłacasz” koszty stałe.',
            'Próg rentowności (BEP) to ilość, przy której przychody równają się kosztom.',
            'Koszt utopiony to pieniądze już wydane, których nie odzyskasz. Nie powinien wpływać na decyzje o przyszłości.',
          ], 'BEP = koszty stałe ÷ (cena − koszt zmienny na sztukę)'),
          sort('Koszt stały czy zmienny?', ['Stały', 'Zmienny'], [
            ['Czynsz za lokal', 0],
            ['Mąka w piekarni', 1],
            ['Pensja księgowej', 0],
            ['Opłata kuriera za paczkę', 1],
            ['Abonament za program do faktur', 0],
          ]),
          c('Koszty stałe 10 000 zł/mies., cena 50 zł, koszt zmienny 30 zł/szt. Gdzie jest próg rentowności?', ['500 sztuk', '200 sztuk', '333 sztuki', '1 000 sztuk'], '10 000 ÷ (50 − 30) = 500 sztuk.'),
          fill('Cena minus koszt zmienny na sztukę to marża ___.', ['pokrycia', 'netto', 'EBITDA', 'podatkowa']),
          tf('Wydaliśmy już 200 tys. zł na projekt, więc musimy go skończyć, żeby pieniądze się nie zmarnowały.', false, 'To pułapka kosztów utopionych. 200 tys. zł jest wydane bez względu na decyzję. Liczy się tylko, czy dalsze nakłady się opłacą.'),
          c('Koszt krańcowy to…', ['koszt wytworzenia jednej dodatkowej sztuki', 'koszt całej produkcji', 'ostatnia faktura w roku', 'koszt zamknięcia firmy']),
        ],
      },
    ],
  },
  // ───────────────────────── 4 ─────────────────────────
  {
    id: 'u4', title: 'Sprawozdania finansowe', subtitle: 'Bilans, rachunek zysków i strat, cash flow', color: 'yellow',
    lessons: [
      {
        id: 'u4l1', title: 'Bilans',
        items: [
          teach('Zdjęcie firmy w jednym dniu', [
            'Bilans pokazuje stan firmy na konkretny dzień.',
            'Aktywa: co firma ma (maszyny, zapasy, należności, gotówka).',
            'Pasywa: skąd to ma (kapitał własny i zobowiązania, np. kredyty, niezapłacone faktury).',
          ], 'Aktywa = Pasywa\nKapitał własny = aktywa − zobowiązania'),
          c('Aktywa 1 mln zł, zobowiązania 600 tys. zł. Ile wynosi kapitał własny?', ['400 tys. zł', '1,6 mln zł', '600 tys. zł', '1 mln zł']),
          sort('Aktywa czy pasywa?', ['Aktywa', 'Pasywa'], [
            ['Gotówka na koncie', 0],
            ['Kredyt bankowy', 1],
            ['Zapasy w magazynie', 0],
            ['Kapitał wniesiony przez wspólników', 1],
            ['Należności od klientów', 0],
            ['Niezapłacone faktury dostawców', 1],
          ]),
          tf('Bilans pokazuje wynik firmy za cały rok.', false, 'Bilans to stan na dany dzień. Wynik za okres pokazuje rachunek zysków i strat.'),
          fill('Pieniądze, które klienci są nam winni, to ___.', ['należności', 'zobowiązania', 'rezerwy', 'koszty']),
          c('Kapitał obrotowy netto to…', ['aktywa obrotowe minus zobowiązania krótkoterminowe', 'cała gotówka firmy', 'kapitał zakładowy', 'zysk netto minus podatek']),
        ],
      },
      {
        id: 'u4l2', title: 'Rachunek zysków i strat',
        items: [
          teach('Film o całym okresie', [
            'Rachunek zysków i strat (RZiS, ang. P&L) pokazuje przychody, koszty i wynik za okres, np. kwartał.',
            'Idzie od góry („top line”, czyli przychody) do dołu („bottom line”, czyli zysk netto).',
          ], 'Przychody − koszt sprzedanych towarów = zysk brutto ze sprzedaży\n− koszty ogólne = zysk operacyjny (EBIT)\n− odsetki = zysk przed opodatkowaniem\n− podatek = zysk netto'),
          c('Przychody 500, koszt sprzedanych towarów 300, koszty ogólne 120. Ile wynosi zysk operacyjny?', ['80', '200', '380', '120'], '500 − 300 = 200 (zysk ze sprzedaży), a 200 − 120 = 80.'),
          tf('„Top line” to przychody, a „bottom line” to zysk netto.', true),
          c('P&L to angielska nazwa…', ['rachunku zysków i strat', 'bilansu', 'rachunku przepływów pieniężnych', 'planu lekcji']),
          c('Zysk przed opodatkowaniem 100 tys. zł, podatek 19%. Ile wynosi zysk netto?', ['81 tys. zł', '119 tys. zł', '19 tys. zł', '100 tys. zł']),
          match('Połącz pojęcie z opisem', [
            ['Top line', 'Przychody'],
            ['Bottom line', 'Zysk netto'],
            ['EBIT', 'Zysk operacyjny'],
            ['COGS', 'Koszt sprzedanych towarów'],
          ]),
        ],
      },
      {
        id: 'u4l3', title: 'Cash flow: zysk to nie gotówka',
        items: [
          teach('Gotówka jest królową', [
            'Rachunek przepływów pieniężnych (cash flow) pokazuje, ile gotówki naprawdę wpłynęło i wypłynęło.',
            'Ma trzy części: operacyjną (klienci, pensje), inwestycyjną (np. CAPEX) i finansową (kredyty, dywidendy).',
            'Firma z zyskiem może upaść, jeśli zabraknie gotówki, np. gdy klienci płacą po 90 dniach, a dostawcy chcą pieniędzy od razu.',
          ], 'Free cash flow ≈ przepływy operacyjne − CAPEX'),
          tf('Firma może mieć zysk i jednocześnie problemy z płynnością.', true),
          sort('Która część rachunku przepływów?', ['Operacyjne', 'Inwestycyjne', 'Finansowe'], [
            ['Wpłaty od klientów', 0],
            ['Zakup maszyny', 1],
            ['Zaciągnięcie kredytu', 2],
            ['Wypłata pensji', 0],
            ['Wypłata dywidendy', 2],
            ['Sprzedaż starego magazynu', 1],
          ]),
          c('Przepływy operacyjne 500 tys. zł, CAPEX 200 tys. zł. Ile wynosi free cash flow?', ['300 tys. zł', '700 tys. zł', '200 tys. zł', '500 tys. zł']),
          c('Klienci płacą Ci po 60 dniach, a dostawcom płacisz po 14. Co to robi z gotówką?', ['Zamraża ją, bo kredytujesz klientów', 'Nic', 'Daje dodatkową gotówkę', 'Obniża podatki']),
          fill('Zdolność firmy do regulowania bieżących zobowiązań to ___.', ['płynność', 'rentowność', 'elastyczność', 'amortyzacja']),
        ],
      },
    ],
  },
  // ───────────────────────── 5 ─────────────────────────
  {
    id: 'u5', title: 'Myśl jak konsultant', subtitle: 'MECE, drzewa problemów, piramida, 80/20', color: 'navy',
    lessons: [
      {
        id: 'u5l1', title: 'MECE',
        items: [
          teach('Bez dziur i bez dubli', [
            'MECE (Mutually Exclusive, Collectively Exhaustive) znaczy „rozłączne i wyczerpujące”.',
            'Dzielisz problem na części, które się nie nakładają (bez dubli) i razem pokrywają całość (bez dziur).',
          ], undefined, 'Klienci wg wieku: do 24, 25–44, 45–64, 65+ to podział MECE. „Studenci, młodzi, kobiety” już nie.'),
          c('Który podział klientów sklepu jest MECE?', ['Nowi / powracający', 'Kobiety / studenci / emeryci', 'Z Warszawy / zadowoleni', 'Duzi / lojalni / online']),
          tf('Podział kosztów na „pensje”, „marketing” i „IT” jest MECE.', false, 'Brakuje „innych” (dziura), a pensja marketingowca pasuje do dwóch koszyków (dubel).'),
          c('Co najprościej „domyka” podział, żeby był wyczerpujący?', ['Dodanie kategorii „inne”', 'Usunięcie największej kategorii', 'Zmiana nazw', 'Dodanie kolorów']),
          fill('„ME” w MECE oznacza, że kategorie się nie ___.', ['nakładają', 'kończą', 'sumują', 'zmieniają']),
          c('Zysk spada. Który pierwszy podział problemu jest MECE?', ['Przychody / koszty', 'Marketing / Kasia / pogoda', 'Klienci / konkurencja', 'Premium / promocje / produkty']),
        ],
      },
      {
        id: 'u5l2', title: 'Drzewo problemów i hipoteza',
        items: [
          teach('Od wielkiego pytania do konkretu', [
            'Drzewo problemów rozbija duże pytanie na mniejsze gałęzie. Każdy poziom powinien być MECE.',
            'Rozbijasz, aż dojdziesz do rzeczy, które da się sprawdzić danymi.',
            'Hipoteza na start: zamiast analizować wszystko, zakładasz najbardziej prawdopodobną odpowiedź i szybko ją weryfikujesz.',
          ]),
          order('Ułóż gałąź drzewa od ogółu do szczegółu', ['Dlaczego spada zysk?', 'Spadają przychody', 'Spada liczba zamówień', 'Spada ruch z reklam']),
          c('Przychody = liczba klientów × …', ['średnia wartość zakupów klienta', 'liczba pracowników', 'czynsz', 'marża netto']),
          tf('Dobra hipoteza to zdanie, które da się potwierdzić albo obalić danymi.', true),
          c('Która hipoteza najlepiej nadaje się do sprawdzenia?', ['Ruch spadł, bo od marca nie działają reklamy w wyszukiwarce', 'Coś jest nie tak z rynkiem', 'Klienci są dziwni', 'Trzeba się bardziej starać']),
          c('Ostatni poziom drzewa problemów powinien zawierać…', ['konkretne, mierzalne kwestie do zbadania', 'ogólne hasła', 'nazwiska winnych', 'pytania retoryczne']),
        ],
      },
      {
        id: 'u5l3', title: 'Piramida Minto i 80/20',
        items: [
          teach('Najpierw odpowiedź', [
            'Zasada piramidy (opisana przez Barbarę Minto): zaczynasz od wniosku, potem 2–4 argumenty, na końcu dane.',
            'Szef czyta pierwsze zdanie, więc niech to będzie odpowiedź.',
            'Zasada Pareto 80/20: często ok. 20% przyczyn daje ok. 80% skutków. To obserwacja, a nie prawo natury.',
          ]),
          order('Ułóż maila do zarządu według piramidy', ['Rekomenduję wejście na rynek czeski w III kwartale', 'Popyt: trzech dużych klientów już pyta', 'Koszt: inwestycja zwraca się w 14 miesięcy', 'Załącznik: szczegółowe wyliczenia']),
          c('Od czego zacząć podsumowanie dla zajętego szefa?', ['Od wniosku lub rekomendacji', 'Od historii projektu', 'Od listy spotkań', 'Od podziękowań']),
          c('Analiza: 15% klientów daje 78% przychodu. Jaki jest rozsądny pierwszy krok?', ['Zrozumieć tę grupę i szczególnie o nią zadbać', 'Od jutra zrezygnować z pozostałych 85% klientów', 'Zignorować dane', 'Obniżyć wszystkim ceny']),
          tf('Zasada 80/20 mówi, że zawsze dokładnie 20% przyczyn daje 80% skutków.', false, 'To reguła kciuka. Proporcje bywają 70/30 czy 90/10. Sedno: skutki rozkładają się nierówno.'),
          fill('W piramidzie Minto wniosek jest na ___.', ['górze', 'dole', 'końcu załącznika', 'marginesie']),
        ],
      },
    ],
  },
  // ───────────────────────── 6 ─────────────────────────
  {
    id: 'u6', title: 'Strategia', subtitle: 'SWOT, 5 sił, macierz BCG, model 7S', color: 'coral',
    lessons: [
      {
        id: 'u6l1', title: 'SWOT i 5 sił',
        items: [
          teach('Firma i jej otoczenie', [
            'SWOT: mocne i słabe strony (wewnątrz firmy) oraz szanse i zagrożenia (w otoczeniu).',
            'Model 5 sił (opisany przez Michaela Portera) ocenia atrakcyjność branży: rywalizację, groźbę nowych graczy, substytuty, siłę dostawców i siłę nabywców.',
          ]),
          sort('Gdzie w SWOT?', ['Mocna strona', 'Słaba strona', 'Szansa', 'Zagrożenie'], [
            ['Doświadczony zespół sprzedaży', 0],
            ['Przestarzały system IT', 1],
            ['Nowe dotacje na cyfryzację', 2],
            ['Wejście dużego konkurenta', 3],
          ]),
          c('Której z 5 sił dotyczy pytanie „Czy klient łatwo zastąpi nasz produkt czymś innym?”', ['Zagrożenia substytutami', 'Siły dostawców', 'Rywalizacji w branży', 'Groźby nowych graczy']),
          tf('Szanse i zagrożenia w SWOT dotyczą otoczenia firmy, a nie jej samej.', true),
          c('Wysokie bariery wejścia (np. drogie licencje) oznaczają…', ['mniejszą groźbę nowych graczy', 'większą rywalizację', 'słabszych dostawców', 'więcej substytutów']),
          match('Połącz siłę z opisem', [
            ['Siła nabywców', 'Klienci wymuszają niższe ceny'],
            ['Siła dostawców', 'Dostawca dyktuje warunki'],
            ['Nowi gracze', 'Łatwo wejść na rynek'],
            ['Substytuty', 'Inny produkt zaspokaja tę samą potrzebę'],
          ]),
        ],
      },
      {
        id: 'u6l2', title: 'Macierz BCG',
        items: [
          teach('Portfel produktów na jednej kartce', [
            'Macierz wzrostu i udziału (spopularyzowana przez Boston Consulting Group) dzieli produkty według dwóch osi: tempa wzrostu rynku i względnego udziału w rynku.',
            'Gwiazdy (szybki wzrost, duży udział): inwestuj.',
            'Dojne krowy (wolny wzrost, duży udział): zbieraj gotówkę.',
            'Znaki zapytania (szybki wzrost, mały udział): zdecyduj, czy inwestować, czy odpuścić.',
            'Psy (wolny wzrost, mały udział): rozważ wyjście.',
          ]),
          match('Połącz ćwiartkę z opisem', [
            ['Gwiazda', 'Rynek rośnie, udział duży'],
            ['Dojna krowa', 'Rynek stoi, udział duży'],
            ['Znak zapytania', 'Rynek rośnie, udział mały'],
            ['Pies', 'Rynek stoi, udział mały'],
          ]),
          c('Produkt jest liderem dojrzałego, stabilnego rynku. Gdzie leży w macierzy?', ['Dojna krowa', 'Gwiazda', 'Znak zapytania', 'Pies']),
          c('Typowa rola dojnej krowy w portfelu to…', ['finansowanie gwiazd i znaków zapytania', 'największe inwestycje w marketing', 'szybkie wycofanie', 'brak jakiejkolwiek roli']),
          tf('Znak zapytania zawsze należy wycofać.', false, 'To miejsce decyzji: zainwestować, żeby zrobić z niego gwiazdę, albo odpuścić.'),
          fill('Osie macierzy BCG to tempo wzrostu rynku i względny ___ w rynku.', ['udział', 'koszt', 'zysk', 'wiek']),
        ],
      },
      {
        id: 'u6l3', title: 'Model 7S',
        items: [
          teach('Siedem elementów, które muszą do siebie pasować', [
            'Model 7S (opisany przez konsultantów McKinsey) mówi, że skuteczna organizacja potrzebuje spójności siedmiu elementów.',
            '„Twarde”: strategia, struktura, systemy.',
            '„Miękkie”: wspólne wartości (w centrum), styl zarządzania, personel i umiejętności.',
            'Jeśli zmienisz tylko strukturę, pozostałe elementy będą ciągnąć firmę po staremu.',
          ]),
          sort('Twardy czy miękki element 7S?', ['Twardy', 'Miękki'], [
            ['Strategia', 0],
            ['Struktura', 0],
            ['Systemy', 0],
            ['Wspólne wartości', 1],
            ['Styl zarządzania', 1],
            ['Umiejętności', 1],
          ]),
          c('Co jest w centrum modelu 7S?', ['Wspólne wartości', 'Struktura', 'Strategia', 'Systemy IT']),
          c('Nowa strategia brzmi „klient na pierwszym miejscu”, ale premie zależą tylko od liczby zamkniętych zgłoszeń. Który element 7S nie pasuje?', ['Systemy (premiowanie)', 'Wspólne wartości', 'Struktura', 'Strategia']),
          tf('Według 7S zmiana wymaga dopasowania wielu elementów, a nie tylko nowego schematu organizacyjnego.', true),
        ],
      },
    ],
  },
  // ───────────────────────── 7 ─────────────────────────
  {
    id: 'u7', title: 'Prowadzenie zespołu', subtitle: 'Priorytety, cele, delegowanie, feedback', color: 'pink',
    lessons: [
      {
        id: 'u7l1', title: 'Priorytety',
        items: [
          teach('Nie wszystko jest na wczoraj', [
            'Macierz pilne/ważne (kojarzona z Eisenhowerem): ważne i pilne zrób teraz, ważne i niepilne zaplanuj, pilne i nieważne deleguj, a resztę usuń.',
            'W ćwiartce „ważne, niepilne” dzieje się rozwój, więc blokuj na nią czas w kalendarzu.',
            'Macierz wpływ/wysiłek: zaczynaj od szybkich wygranych (duży wpływ, mały wysiłek).',
          ]),
          match('Co robisz z zadaniem?', [
            ['Ważne i pilne', 'Zrób teraz'],
            ['Ważne, niepilne', 'Zaplanuj'],
            ['Pilne, nieważne', 'Deleguj'],
            ['Ani pilne, ani ważne', 'Usuń'],
          ]),
          c('Awaria u kluczowego klienta, umowa jest zagrożona. Co robisz?', ['Zajmujesz się tym od razu', 'Planujesz na przyszły tydzień', 'Usuwasz z listy', 'Odkładasz do piątku']),
          c('Gdzie trafia „przygotować plan rozwoju zespołu na rok”?', ['Ważne, niepilne: zaplanuj w kalendarzu', 'Pilne i ważne', 'Do usunięcia', 'Do oddania praktykantowi']),
          c('Który pomysł zrobić najpierw według macierzy wpływ/wysiłek?', ['Duży wpływ, mały wysiłek', 'Mały wpływ, duży wysiłek', 'Duży wpływ, duży wysiłek', 'Mały wpływ, mały wysiłek']),
          tf('Zadania ważne, ale niepilne najłatwiej zaniedbać.', true, 'Nic się nie pali, więc przegrywają z mailami. Dlatego rezerwuj na nie czas.'),
        ],
      },
      {
        id: 'u7l2', title: 'Cele: SMART, KPI, OKR',
        items: [
          teach('Cel, który da się sprawdzić', [
            'SMART: cel konkretny, mierzalny, osiągalny, istotny i określony w czasie.',
            'KPI to kluczowy wskaźnik efektywności, który mierzysz regularnie.',
            'OKR to ambitny cel (Objective) i 2–5 mierzalnych kluczowych rezultatów (Key Results).',
          ]),
          c('Który cel jest SMART?', ['Podnieść konwersję sklepu z 1,8% do 2,5% do 30 czerwca', 'Poprawić sprzedaż', 'Być najlepszym zespołem', 'Robić więcej marketingu']),
          match('Rozszyfruj SMART', [
            ['S', 'Konkretny'],
            ['M', 'Mierzalny'],
            ['A', 'Osiągalny'],
            ['R', 'Istotny'],
            ['T', 'Określony w czasie'],
          ]),
          c('Kluczowy rezultat (Key Result) w OKR powinien być…', ['mierzalny, np. „NPS z 30 do 45”', 'ogólny, np. „zadowoleni klienci”', 'listą zadań', 'tajny']),
          tf('KPI i OKR to dokładnie to samo.', false, 'KPI to wskaźnik monitorowany stale. OKR to ambitny cel na okres z mierzalnymi rezultatami. Często KPI pomaga mierzyć KR.'),
          fill('Wskaźnik, który regularnie mierzysz, by ocenić efektywność, to ___.', ['KPI', 'CAPEX', 'SWOT', 'MECE']),
        ],
      },
      {
        id: 'u7l3', title: 'Delegowanie i feedback',
        items: [
          teach('Oddaj zadanie, nie tylko robotę', [
            'Delegując, przekaż: oczekiwany efekt, termin, zasoby, zakres decyzji i moment kontroli. Nie narzucaj „jak”, jeśli nie musisz.',
            'Feedback: opisz sytuację, konkretne zachowanie i jego wpływ. Potem zapytaj o zdanie i ustalcie kolejny krok.',
            'Unikaj ocen osoby („jesteś niedbały”). Mów o faktach.',
          ]),
          c('Który feedback jest najlepszy?', ['Na wtorkowym spotkaniu przerwałeś klientowi trzy razy i klient się wycofał. Jak to widzisz?', 'Jesteś niegrzeczny.', 'Zawsze przerywasz ludziom.', 'Było słabo, popraw się.']),
          tf('Delegowanie to przekazanie zadania razem z odpowiedzialnością za efekt i potrzebnymi uprawnieniami.', true),
          order('Ułóż feedback w dobrej kolejności', ['Sytuacja: wczorajsza prezentacja', 'Zachowanie: slajdy bez wniosków', 'Wpływ: zarząd nie podjął decyzji', 'Pytanie: co pomoże następnym razem?']),
          c('Czego nie warto robić przy delegowaniu?', ['Zabierać zadania przy pierwszym problemie', 'Ustalać terminu', 'Mówić, jaki efekt jest oczekiwany', 'Umawiać punktu kontrolnego']),
          c('Mikrozarządzanie to…', ['nadmierna kontrola szczegółów pracy zespołu', 'zarządzanie małą firmą', 'delegowanie drobnych zadań', 'krótkie spotkania']),
        ],
      },
    ],
  },
  // ───────────────────────── 8 ─────────────────────────
  {
    id: 'u8', title: 'Makro i inwestycje', subtitle: 'Inflacja, stopy, ROI, NPV', color: 'lilac',
    lessons: [
      {
        id: 'u8l1', title: 'Inflacja i stopy procentowe',
        items: [
          teach('Gospodarka wokół Twojej firmy', [
            'Inflacja to wzrost ogólnego poziomu cen: za te same pieniądze kupisz mniej.',
            'W Polsce stopy procentowe ustala Rada Polityki Pieniężnej (NBP). Wyższe stopy oznaczają droższy kredyt i zwykle hamują inflację.',
            'PKB to wartość wszystkich dóbr i usług wytworzonych w kraju w danym okresie.',
          ]),
          tf('Inflacja wynosi 5%, a Twoja pensja rośnie o 3%. Realnie zarabiasz mniej.', true, 'Realna zmiana ≈ 3% − 5% = −2%.'),
          c('Kto ustala stopy procentowe NBP?', ['Rada Polityki Pieniężnej', 'Ministerstwo Finansów', 'Sejm', 'Giełda Papierów Wartościowych']),
          c('Bank centralny podnosi stopy. Co zwykle dzieje się z ratą kredytu o zmiennym oprocentowaniu?', ['Rośnie', 'Spada', 'Nie zmienia się', 'Znika']),
          c('PKB to…', ['wartość dóbr i usług wytworzonych w kraju w danym okresie', 'budżet państwa', 'suma wszystkich pensji', 'wartość eksportu']),
          c('Złoty słabnie wobec euro. Komu to zwykle pomaga?', ['Polskiemu eksporterowi sprzedającemu w euro', 'Importerowi płacącemu w euro', 'Turystom jadącym do Hiszpanii', 'Nikomu']),
        ],
      },
      {
        id: 'u8l2', title: 'ROI i okres zwrotu',
        items: [
          teach('Czy to się opłaca?', [
            'ROI (zwrot z inwestycji) pokazuje, ile zarobiłaś w stosunku do tego, ile włożyłaś.',
            'Okres zwrotu (payback) mówi, po jakim czasie odzyskasz nakłady.',
          ], 'ROI = (wpływy − koszt) ÷ koszt\nOkres zwrotu = nakład ÷ roczny wpływ gotówki'),
          c('Inwestycja 100 000 zł daje łącznie 130 000 zł. Ile wynosi ROI?', ['30%', '130%', '13%', '3%']),
          c('Maszyna kosztuje 120 000 zł i daje 40 000 zł gotówki rocznie. Jaki jest okres zwrotu?', ['3 lata', '4 lata', '1,2 roku', '12 lat']),
          tf('ROI nie mówi nic o tym, jak długo czekasz na zwrot.', true, '30% w rok i 30% w 10 lat to zupełnie różne inwestycje. Dlatego liczy się też czas.'),
          c('Kampania kosztowała 10 000 zł i dała 25 000 zł marży. Ile wynosi ROI?', ['150%', '250%', '25%', '15%'], '(25 000 − 10 000) ÷ 10 000 = 1,5, czyli 150%.'),
        ],
      },
      {
        id: 'u8l3', title: 'Wartość pieniądza w czasie i NPV',
        items: [
          teach('Złotówka dziś > złotówka jutro', [
            'Pieniądz dziś jest wart więcej niż ta sama kwota za rok: możesz go zainwestować, a inflacja zjada wartość.',
            'Dyskontowanie przelicza przyszłe kwoty na dzisiejsze.',
            'NPV (wartość bieżąca netto) to suma zdyskontowanych przyszłych przepływów minus nakład. NPV > 0 oznacza, że inwestycja tworzy wartość.',
          ], 'Wartość dziś = kwota przyszła ÷ (1 + r)ⁿ'),
          c('Wolisz 1 000 zł dziś czy 1 000 zł za rok?', ['Dziś, bo mogę je zainwestować, zanim stracą na wartości', 'Za rok, bo będzie ich więcej', 'Wszystko jedno', 'Za rok, bo inflacja podnosi wartość pieniądza']),
          c('Ile jest warte dziś 1 100 zł otrzymane za rok przy stopie 10%?', ['1 000 zł', '1 100 zł', '1 210 zł', '990 zł'], '1 100 ÷ 1,10 = 1 000 zł.'),
          tf('Projekt z ujemnym NPV (przy właściwej stopie dyskontowej) niszczy wartość.', true),
          c('Im wyższa stopa dyskontowa, tym wartość bieżąca przyszłych przepływów jest…', ['niższa', 'wyższa', 'taka sama', 'zawsze równa zero']),
          fill('Inwestujemy, jeśli NPV jest ___ od zera.', ['większe', 'mniejsze', 'równe', 'niezależne']),
        ],
      },
    ],
  },
]

export const units: Unit[] = raw.map((u) => ({
  ...u,
  lessons: u.lessons.map((l) => ({
    id: l.id,
    title: l.title,
    items: l.items.map((it, i): Item => ({ ...it, id: `${l.id}-${i}` }) as Item),
  })),
}))
