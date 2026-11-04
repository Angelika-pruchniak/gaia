# Przelicza model v3 w 3 scenariuszach (LibreOffice) i wpisuje wyniki do arkusza "Scenariusze". Użycie: python3 add_scenarios_v3.py <katalog>
import sys, os, subprocess, openpyxl
SRC = "../GA_GAia_model_wspolpracy_v3.xlsx"; WD = sys.argv[1]; os.makedirs(WD, exist_ok=True)
def rows(ws):
    d = {}
    for r in ws.iter_rows(values_only=True):
        if r[0] and r[0] not in d: d[r[0]] = r
    return d
res = []
for s in (1, 2, 3):
    wb = openpyxl.load_workbook(SRC)
    for r in wb["Założenia"].iter_rows():
        if r[0].value and str(r[0].value).startswith("Scenariusz (1"): r[1].value = s
    p = os.path.join(WD, f"s{s}.xlsx"); wb.save(p); out = os.path.join(WD, "out")
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", out, p], capture_output=True)
    c = openpyxl.load_workbook(os.path.join(out, f"s{s}.xlsx"), data_only=True)
    P, U, J, Kp = rows(c["Każdy partner zyskuje"]), rows(c["Rynek i GMV"]), rows(c["Wspólna pula"]), rows(c["Koszty platformy"])
    ga, tp, mp, pl = P["Grupa Azoty"], P["Partner Technologiczny"], P["Partner Marketplace / Handel"], P["Wspólna pula: EBITDA platformy"]
    cum = J["Wynik puli narastająco"][1:9]
    res.append([U["Użytkownicy (średnio w roku, tys.)"][6], U["GMV razem"][6], J["PRZYCHODY PULI RAZEM"][6], J["WYNIK PULI (EBITDA platformy)"][6],
                min(cum), pl[11], ga[9], ga[10], tp[9], tp[10], tp[11], mp[9], mp[10], mp[11], tp[12], mp[12], Kp["Oszczędność dzięki partnerstwu"][9]])
    print(s, [round(x, 1) if isinstance(x, float) else x for x in res[-1]])
wb = openpyxl.load_workbook(SRC); S = wb["Scenariusze"]
for k in range(len(res[0])):
    for j in range(3):
        v = res[j][k]; c = S.cell(5 + k, 2 + j, round(v, 3) if isinstance(v, float) else v); c.number_format = '#,##0.0;[Red]-#,##0.0;"–"'
wb.save(SRC); print("ok")
