# Generatory: GAia, model rozliczeń z partnerem technologicznym

1. `python3 build_xlsx.py` tworzy `../GA_GAia_model_JV_partnerzy.xlsx` (formuły, założenia na niebiesko).
2. Przelicz plik w LibreOffice/Excel (oraz kopię z mnożnikiem adopcji 0,5 do stres testu).
3. `python3 build_pptx.py <przeliczony.xlsx> <przeliczony_stres.xlsx>` tworzy `../GA_GAia_jak_zarabiamy_z_partnerem.pptx`.

Liczby na slajdach pochodzą wyłącznie z modelu. Po podmianie założeń na dane (CO-PA, umowy AD, oferty partnerów) wystarczy powtórzyć kroki 2-3.
