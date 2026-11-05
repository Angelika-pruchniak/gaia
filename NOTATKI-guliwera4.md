# Guliwera 4: notatki z prac nad stroną guliwera4.pl

Właścicielka: **Angelika Pruchniak-Gąszczyk** (Guliwera 4, ul. Guliwera 4, Lublin, NIP 9462713420)
LinkedIn firmy: https://www.linkedin.com/company/guliwera-4/
E-mail na stronie: kontakt@guliwera4.pl

## Marka (paleta v3 „liliputy”)
- Kolory: krem #FFF6EE, koral #FF6B3D, koral ciemny #E8512A, róż #F4799F, róż jasny #FFD9E4,
  liliowy #B9A5E3, liliowy jasny #EDE5FB, żółty #FFC84A, żółty jasny #FFEFC7, brzoskwinia #FFD8C2,
  atrament #3A2340, atrament 60% #7A6580, granat #1E2A4A.
- Fonty: Fraunces (nagłówki), DM Sans (tekst), Caveat (odręczne wstawki).
- Hasło: „strategia, produkty cyfrowe i dobry humor”.
- Logo i favicon: folder `logo/` (ikona-512.png, logo-poziome.png, mark.svg, avatar i baner LinkedIn).
- Zdjęcie: `zdjecia/angelika-portret.jpg`. Tylko kadr i wyostrzenie, bez zmiany rysów twarzy (wyraźna prośba).

## Strona (WordPress + Jetpack, site ID 258438467)
- Strony: 10 Strona główna, 32 Usługi, 6 O mnie, 52 Lokalizacja, 60 Kontakt (formularz Jetpack), 74 Materiały.
  Treść to bloki `wp:html` z szablonem `elementor_canvas` (bez Elementora, żeby uniknąć starego cache).
- WooCommerce: 62 Sklep (przekierowuje na /materialy/), 63 Koszyk, 64 Zamówienie (/zamowienie/), 65 Moje konto.
- Produkty „Kierunek Zysk” (cyfrowe, każdy tylko raz w koszyku):
  - 67 Pakiet (297 zł, ZIP)
  - 68 Przewodnik (147 zł)
  - 69 Strategia na 1 stronie (39 zł)
  - 70 Planer 90 dni (59 zł)
  - 71 Skrypty (79 zł)
  - 72 Audyt (59 zł)
  - 73 Dziennik decyzji (49 zł)
  - Link zakupu: `/zamowienie/?add-to-cart=ID`.

## Wtyczki własne (aktywne, zostawić)
- `guliwera-favicon`: ikona strony (media 96).
- `guliwera-gtm`: Google Tag Manager GTM-5JKVC955 + zdarzenia e-commerce GA4 (view_item, add_to_cart, begin_checkout, purchase).
- `guliwera-sklep-styl` (kod w `wtyczki/`):
  - produkty, koszyk, zamówienie i konto w wyglądzie strony;
  - Polska, ceny w formacie „39,00 zł”;
  - zamówienie bez adresu;
  - PDF tylko raz w koszyku.
- Instalacja wtyczek: ZIP z GitHub raw (adres z hashem commita), bo bezpośredni upload jest blokowany.

## Do zrobienia
- [ ] Angelika: połączyć konto Stripe (WooCommerce → Ustawienia → Płatności → Stripe) i włączyć Kartę, BLIK, Przelewy24. Potem zakup testowy.
- [ ] Angelika: w Tag Managerze dodać „Tag Google” z identyfikatorem GA4 (G-…) i opublikować kontener.
- [ ] Ustawić repozytorium `gaia` jako prywatne (gałąź `claude/affectionate-sagan-a00b7d` ma płatne PDF-y).
- [ ] Opcjonalnie:
  - polityka prywatności i regulamin sklepu (strona 3 i 66 to szkice);
  - usunąć duplikat strony 28;
  - okładki produktów jako obrazki w koszyku;
  - PDF „budżet marketingowy”.

## Następny projekt
Aplikacja do nauki dla młodych managerów: krótkie codzienne lekcje, seria dni, punkty, poziomy.
Wiedza z publicznie opisanych frameworków strategii i zarządzania (np. macierz BCG, 7S, MECE, piramida Minto).
Oryginalna marka Guliwera 4, bez logo i materiałów chronionych prawem autorskim.
