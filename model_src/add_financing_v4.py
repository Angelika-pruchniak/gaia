# Przelicza model v4 w wariantach (LibreOffice) i wpisuje wyniki do arkusza "Finansowanie".
# Tabela 1: wrażliwość na CAC, tabela 2: ambicja (adopcja w roku 5), tabela 3: synergie (do dyskusji).
# Użycie: python3 add_financing_v4.py <katalog_roboczy>
import sys, os, subprocess, openpyxl
from openpyxl.styles import Font, PatternFill
SRC = "../GA_GAia_model_wspolpracy_v4.xlsx"; WD = sys.argv[1]; os.makedirs(WD, exist_ok=True)
FIN_T1, FIN_T2, FIN_T3 = 4, 13, 21
BASE_ADOPT = [0.05, 0.12, 0.20, 0.28, 0.35, 0.40, 0.43]

def rows(ws):
    d = {}
    for r in ws.iter_rows(values_only=True):
        if r[0] and r[0] not in d: d[r[0]] = r
    return d

def find_row(ws, prefix):
    for r in ws.iter_rows():
        if r[0].value and str(r[0].value).startswith(prefix): return r
    raise KeyError(prefix)

def first_pos(vals):
    for j, v in enumerate(vals):
        if v is not None and v > 0.0005: return f"rok {j}"
    return "po roku 7"

def run(tag, cac=None, adopt5=None, syt=None, sym=None):
    wb = openpyxl.load_workbook(SRC); A = wb["Założenia"]
    if cac is not None:
        r = find_row(A, "CAC na nowego")
        for i in range(7): r[1 + i].value = cac
    if adopt5 is not None:
        r = find_row(A, "AMBICJA")
        k = adopt5 / 0.35
        for i in range(7): r[1 + i].value = round(min(0.75, BASE_ADOPT[i] * k), 4)
    if syt is not None: find_row(A, "Mniejszy zespół")[1].value = syt
    if sym is not None: find_row(A, "Niższe utrzymanie")[1].value = sym
    p = os.path.join(WD, f"{tag}.xlsx"); wb.save(p); out = os.path.join(WD, "out")
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", out, p], capture_output=True)
    c = openpyxl.load_workbook(os.path.join(out, f"{tag}.xlsx"), data_only=True)
    P, U, J, Kp = rows(c["Każdy partner zyskuje"]), rows(c["Rynek i GMV"]), rows(c["Wspólna pula"]), rows(c["Koszty platformy"])
    cum = J["Wynik puli narastająco"][1:9]; net = J["WYNIK PULI (EBITDA platformy)"][1:9]
    res = [U["Użytkownicy (średnio w roku, tys.)"][6], Kp["Pozyskanie i utrzymanie użytkowników"][9], -min(0, min(cum)),
           first_pos(net), first_pos(cum), P["Grupa Azoty"][9], P["Partner Technologiczny"][9], P["Partner Marketplace / Handel"][9]]
    print(tag, [round(x, 1) if isinstance(x, float) else x for x in res])
    return res

T1 = [(f"CAC {v} zł" + (" (założenie GA)" if v == 250 else ""), dict(cac=v)) for v in (150, 200, 250, 300, 350)]
T2 = [("20% rolników w roku 5", dict(adopt5=0.20)), ("28% rolników w roku 5", dict(adopt5=0.28)),
      ("35% rolników w roku 5 (model GA)", dict(adopt5=0.35)), ("50% rolników w roku 5", dict(adopt5=0.50))]
T3 = [("Z synergiami: zespół −30%, utrzymanie −25%", dict(syt=0.30, sym=0.25)), ("Tylko synergia zespołu −30%", dict(syt=0.30, sym=0.0)),
      ("Bez żadnych synergii", dict(syt=0.0, sym=0.0))]
results = {}
for tbl in (T1, T2, T3):
    for k, (lab, kw) in enumerate(tbl): results[lab] = run(lab.split(" ")[0] + str(abs(hash(lab)) % 10**6), **kw)

wb = openpyxl.load_workbook(SRC); S = wb["Finansowanie"]
NUM = '#,##0.0;[Red]-#,##0.0;"–"'
base_fill = PatternFill("solid", fgColor="DDEBF7")
for start, tbl in ((FIN_T1, T1), (FIN_T2, T2), (FIN_T3, T3)):
    for k, (lab, kw) in enumerate(tbl):
        r = start + 2 + k; S.cell(r, 1, lab)
        for j, v in enumerate(results[lab]):
            c = S.cell(r, 2 + j, round(v, 3) if isinstance(v, float) else v)
            c.number_format = "#,##0" if j == 0 else NUM
        if "GA)" in lab or lab.startswith("Z synergiami"):
            for j in range(1, 10): S.cell(r, j).fill = base_fill; S.cell(r, j).font = Font(name="Trebuchet MS", bold=True, size=10)
S.cell(FIN_T3 + 6, 1, "Wiersze pogrubione = założenia modelu (ambicja GA, CAC 250 zł, synergie do dyskusji). Finansowanie pokrywają PT 70% / PM 30% oraz granty; GA nie wnosi gotówki.").font = Font(name="Trebuchet MS", italic=True, size=9, color="595959")
wb.save(SRC); print("ok")
