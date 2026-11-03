# Dopisuje do arkusza "Każdy zyskuje" przeliczony wariant 50% użytkowników (wartości, nie formuły).
# Użycie: python3 add_stress_bc.py <przeliczony_stres.xlsx>
import sys, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.comments import Comment
FONT = "Trebuchet MS"; NAVY = "003665"
thin = Side(style="thin", color="BFBFBF"); nt = Side(style="thin", color=NAVY)
st = openpyxl.load_workbook(sys.argv[1], data_only=True)["Każdy partner zyskuje"]
wb = openpyxl.load_workbook("../GA_GAia_business_case_300k.xlsx"); P = wb["Każdy partner zyskuje"]
P.cell(26, 1, "Stres test: 150 tys. użytkowników w roku 5 (50% planu)").font = Font(name=FONT, bold=True, color=NAVY, size=11)
P.cell(26, 1).comment = Comment("Wartości przeliczone przy mnożniku użytkowników 0,5 (Założenia, sekcja 1). To wartości, nie formuły: aby odświeżyć, ustaw mnożnik 0,5 i odczytaj tabelę wyżej.", "Model GAia")
for j, h in enumerate(["Partner", "Suma 0-5, mln zł", "NPV (10%), mln zł", "Zwrot od"]):
    c = P.cell(27, 1 + j, h); c.font = Font(name=FONT, bold=True, color="FFFFFF", size=10); c.fill = PatternFill("solid", fgColor=NAVY)
    c.border = Border(left=nt, right=nt, top=nt, bottom=nt); c.alignment = Alignment(horizontal="center" if j else "left", vertical="center", wrap_text=True)
rows = []
for r in st.iter_rows(min_row=5, max_row=8, values_only=True):
    if r[0] not in [x[0] for x in rows]: rows.append(r)
for k, r in enumerate(rows):
    vals = [r[0], round(r[7], 1), round(r[8], 1), r[10]]
    for j, v in enumerate(vals):
        c = P.cell(28 + k, 1 + j, v); c.font = Font(name=FONT, size=10, color=("C00000" if isinstance(v, float) and v < 0 else None))
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin); c.number_format = '#,##0.0;[Red]-#,##0.0;"–"'
        c.alignment = Alignment(horizontal="left" if j == 0 else "right", vertical="center")
P.cell(33, 1, "Wariant ostrożny: połowa użytkowników. Patrz wartości powyżej (suma 5 lat i NPV każdego partnera).").font = Font(name=FONT, italic=True, size=10, color="595959")
wb.save("../GA_GAia_business_case_300k.xlsx"); print("ok")
