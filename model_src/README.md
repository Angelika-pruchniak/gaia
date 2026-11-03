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
