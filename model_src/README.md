# Generatory: GAia, model rozliczeń z partnerem technologicznym

1. `python3 build_xlsx.py` tworzy `../GA_GAia_model_JV_partnerzy.xlsx` (formuły, założenia na niebiesko).
2. Przelicz plik w LibreOffice/Excel (oraz kopię z mnożnikiem adopcji 0,5 do stres testu).
3. `python3 build_pptx.py <przeliczony.xlsx> <przeliczony_stres.xlsx>` tworzy `../GA_GAia_jak_zarabiamy_z_partnerem.pptx`.

Liczby na slajdach pochodzą wyłącznie z modelu. Po podmianie założeń na dane (CO-PA, umowy AD, oferty partnerów) wystarczy powtórzyć kroki 2-3.

4. `python3 build_overlay.py <przeliczony.xlsx> <przeliczony_stres.xlsx>` nakłada liczby na oryginalny slajd z 30.10 (`slajd_30Oct.png`) i tworzy `../GA_GAia_update_30Oct_z_danymi.pptx` (slajd 30.10 + 4 slajdy uzasadnienia).

## Business case 300 tys. użytkowników
1. `python3 build_bc_xlsx.py` tworzy `../GA_GAia_business_case_300k.xlsx`.
2. Przelicz plik oraz kopię z mnożnikiem użytkowników 0,5 (stres test), potem `python3 add_stress_bc.py <przeliczony_stres.xlsx>`.
3. `python3 build_bc_pptx.py <przeliczony.xlsx> <przeliczony_stres.xlsx>` tworzy `../GA_GAia_business_case_300k.pptx` (12 slajdów).
4. `python3 build_overlay_bc.py <przeliczony.xlsx> <przeliczony_stres.xlsx>` tworzy `../GA_GAia_prezentacja_business_case.pptx`: Wasz slajd z 30.10 z liczbami z business case + 12 slajdów, font Trebuchet MS i kolory szablonu.

## Excel: formatowanie i tabele przestawne
- `build_bc_xlsx.py` wykonuje na końcu `xlsx_finish.py`: opisy „Jak liczone / skąd”, komentarze w komórkach, arkusz Dane, spis treści, legenda, formatowanie bez siatki.
- Po `add_stress_bc.py` uruchom `python3 add_pivots.py ../GA_GAia_business_case_300k.xlsx` (LibreOffice + uno): tworzy 8 tabel przestawnych na arkuszu „Tabele przestawne”.

## Model wspólnej puli przychodów (rekomendacja strategiczna)
1. `python3 build_jr_xlsx.py` → `../GA_GAia_model_wspolnej_puli.xlsx` (3 scenariusze, klucz podziału, pula, etapy, case study).
2. `python3 add_scenarios.py <katalog_roboczy>` → wypełnia arkusz Scenariusze (3 przeliczenia w LibreOffice).
3. `python3 add_pivots_jr.py ../GA_GAia_model_wspolnej_puli.xlsx` → 8 tabel przestawnych.
4. `python3 build_jr_pptx.py ../GA_GAia_model_wspolnej_puli.xlsx` → `../GA_GAia_wspolna_pula.pptx`; `python3 build_overlay_jr.py ../GA_GAia_model_wspolnej_puli.xlsx` → `../GA_GAia_prezentacja_wspolna_pula.pptx` (ze slajdem 30.10).

## Model współpracy v3 (na bazie GAia-model_PL_v2 i wymagań MVP1)
1. `python3 build_v3_xlsx.py` → `../GA_GAia_model_wspolpracy_v3.xlsx` (7 lat, 3 scenariusze; agresywny = model GA, synergie, pula, klucz, dopasowanie MVP1, model współpracy, uwagi do modelu GA).
2. `python3 add_scenarios_v3.py <katalog_roboczy>` → arkusz Scenariusze.
3. `python3 add_pivots_v3.py ../GA_GAia_model_wspolpracy_v3.xlsx` → 9 tabel przestawnych.
4. `python3 build_v3_pptx.py ../GA_GAia_model_wspolpracy_v3.xlsx`, potem `python3 build_overlay_v3.py ../GA_GAia_model_wspolpracy_v3.xlsx` → `../GA_GAia_prezentacja_model_wspolpracy_v3.pptx`.
