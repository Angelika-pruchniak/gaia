# Przelicza model w 3 scenariuszach (LibreOffice) i wpisuje wyniki do arkusza "Scenariusze".
# Użycie: python3 add_scenarios.py <katalog_roboczy>
import sys, os, subprocess, openpyxl
SRC = "../GA_GAia_model_wspolnej_puli.xlsx"; WD = sys.argv[1]; os.makedirs(WD, exist_ok=True)
def rows(ws):
    d = {}
    for r in ws.iter_rows(values_only=True):
        if r[0] and r[0] not in d: d[r[0]] = r
    return d
res = []
for s in (1, 2, 3):
    wb = openpyxl.load_workbook(SRC); A = wb["Założenia"]
    for r in A.iter_rows():
        if r[0].value and str(r[0].value).startswith("Scenariusz (1"): r[1].value = s
    p = os.path.join(WD, f"s{s}.xlsx"); wb.save(p)
    out = os.path.join(WD, "out"); subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", out, p], capture_output=True)
    c = openpyxl.load_workbook(os.path.join(out, f"s{s}.xlsx"), data_only=True)
    P, U, J = rows(c["Każdy partner zyskuje"]), rows(c["Użytkownicy i GMV"]), rows(c["Wspólna pula"])
    ga, tp, mp = P["Grupa Azoty"], P["Partner Technologiczny"], P["Partner Marketplace / Handel"]
    deficit = P["Maks. deficyt puli (finansują PT 70% / PM 30%)"][2]
    res.append([U["Użytkownicy zarejestrowani (tys.)"][6], U["GMV RAZEM"][6], U["Nawozy przez platformę / sprzedaż krajowa GA"][6],
                J["PRZYCHODY PULI RAZEM"][6], J["WYNIK PULI (przychody − koszty wspólne)"][6], ga[7], ga[8], tp[7], tp[8], tp[10], mp[7], mp[8], mp[10], deficit])
    print(s, [round(x, 1) if isinstance(x, float) else x for x in res[-1]])
wb = openpyxl.load_workbook(SRC); S = wb["Scenariusze"]
for k in range(len(res[0])):
    for j in range(3):
        c = S.cell(5 + k, 2 + j, round(res[j][k], 3) if isinstance(res[j][k], float) else res[j][k])
        c.number_format = "0.0%" if k == 2 else '#,##0.0;[Red]-#,##0.0;"–"'
wb.save(SRC); print("ok")
