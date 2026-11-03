# Dopisuje do arkusza "Każdy zyskuje" przeliczony wariant 50% użytkowników (wartości, nie formuły).
# Użycie: python3 add_stress_bc.py <przeliczony_stres.xlsx>
import sys, openpyxl
from openpyxl.styles import Font
st = openpyxl.load_workbook(sys.argv[1], data_only=True)["Każdy zyskuje"]
wb = openpyxl.load_workbook("../GA_GAia_business_case_300k.xlsx"); P = wb["Każdy zyskuje"]
P.cell(26, 1, "Stres test: 150 tys. użytkowników w roku 5 (50% planu) - wartości przeliczone, mln zł").font = Font(bold=True)
for j, h in enumerate(["Uczestnik", "Suma 0-5", "NPV (10%)", "Zwrot od"]): P.cell(27, 1 + j, h).font = Font(bold=True)
for k, r in enumerate(st.iter_rows(min_row=5, max_row=10, values_only=True)):
    P.cell(28 + k, 1, r[0]); P.cell(28 + k, 2, round(r[7], 1)); P.cell(28 + k, 3, round(r[8], 1)); P.cell(28 + k, 4, r[10])
P.cell(35, 1, "Przy 50% użytkowników wszyscy uczestnicy nadal są na plusie w 5 lat. Sama spółka GAia wychodzi na zero dopiero po roku 5, bo zespół rozwoju jest stały.")
wb.save("../GA_GAia_business_case_300k.xlsx"); print("ok")
