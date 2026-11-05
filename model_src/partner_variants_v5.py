# Przelicza model v5 (LibreOffice) w wariantach istotnych dla partnerów technologicznych i zapisuje wyniki do JSON.
# Użycie: python3 partner_variants_v5.py <katalog_roboczy>   → partner_variants_v5.json
import sys, os, json, subprocess, openpyxl
SRC = "../GA_GAia_model_wspolpracy_v5.xlsx"; WD = sys.argv[1]; os.makedirs(WD, exist_ok=True)
BASE_ADOPT = [0.05, 0.12, 0.20, 0.28, 0.35, 0.40, 0.43]
BASE_WAL = [0.05, 0.10, 0.15, 0.17, 0.19, 0.21, 0.22]

def rows(ws):
    d = {}
    for r in ws.iter_rows(values_only=True):
        if r[0] and r[0] not in d: d[r[0]] = r
    return d

def find_row(ws, prefix):
    for r in ws.iter_rows():
        if r[0].value and str(r[0].value).startswith(prefix): return r
    raise KeyError(prefix)

def irr(cf):
    lo, hi = -0.99, 5.0
    f = lambda r: sum(c / (1 + r) ** t for t, c in enumerate(cf))
    if f(lo) * f(hi) > 0: return None
    for _ in range(200):
        m = (lo + hi) / 2
        if f(lo) * f(m) <= 0: hi = m
        else: lo = m
    return (lo + hi) / 2

def run(tag, mpt=0, mpm=0, ft=0.7, wal=1.0, adopt5=None, cac_add=0, syt=0, sym=0, ov=None):
    wb = openpyxl.load_workbook(SRC); A = wb["Założenia"]
    for pre, val in (ov or {}).items():
        r = find_row(A, pre)
        for i, v_ in enumerate(val if isinstance(val, list) else [val]): r[1 + i].value = v_
    find_row(A, "Marża PT na usługach")[1].value = mpt
    find_row(A, "Marża PM na usługach")[1].value = mpm
    find_row(A, "Deficyt puli finansuje Partner Technologiczny")[1].value = ft
    find_row(A, "Mniejszy zespół")[1].value = syt
    find_row(A, "Niższe utrzymanie")[1].value = sym
    r = find_row(A, "Udział wydatków rolnika")
    for i in range(7): r[1 + i].value = round(BASE_WAL[i] * wal, 4)
    if adopt5:
        r = find_row(A, "AMBICJA")
        for i in range(7): r[1 + i].value = round(BASE_ADOPT[i] * adopt5 / 0.35, 4)
    if cac_add:
        r = find_row(A, "CAC na nowego")
        for i in range(7): r[1 + i].value = r[1 + i].value + cac_add
    p = os.path.join(WD, f"{tag}.xlsx"); wb.save(p); out = os.path.join(WD, "out")
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", out, p], capture_output=True)
    c = openpyxl.load_workbook(os.path.join(out, f"{tag}.xlsx"), data_only=True)
    res = {"tag": tag}
    for key, sh in (("PT", "Partner Technologiczny"), ("PM", "Partner Marketplace"), ("GA", "Grupa Azoty")):
        R = rows(c[sh]); d = {}
        for lab, k in (("0. Wkład własny (rok 0)", "own"), ("1. Dopłaty do deficytu puli i ich zwrot", "fin"), ("2. Udział w podziale puli", "dist"),
                       ("3. Marża na usługach dla puli", "mar"), ("WYNIK PARTNERA", "res"), ("WYNIK GRUPY AZOTY", "res"),
                       ("Rozwój platformy (100%)", "c1"), ("Utrzymanie (80%)", "c2"), ("Zespół IT i produkt (50%)", "c3"),
                       ("Utrzymanie marketplace'u (20%)", "c1"), ("Zespół zakupów, kategorii, sprzedawców (30%)", "c2")):
            if lab in R: d[k] = [round(x or 0, 3) for x in R[lab][1:9]]
        cf = d["res"]; cum = [round(sum(cf[:i + 1]), 3) for i in range(8)]
        d["cum"] = cum; d["sum6"] = cum[6]; d["sum7"] = cum[7]; d["trough"] = min(cum)
        d["payback"] = next((i for i in range(8) if cum[i] > 0 and all(x > 0 for x in cum[i:])), None)
        d["irr6"] = irr(cf[:7]); d["irr7"] = irr(cf)
        d["npv7"] = round(sum(x / 1.1 ** t for t, x in enumerate(cf)), 3)
        d["gate2"] = cum[2]
        res[key] = d
    J = rows(c["Wspólna pula"]); U = rows(c["Rynek i GMV"])
    res["pool"] = {k: [round(x or 0, 3) for x in J[lab][1:9]] for k, lab in (("rev", "PRZYCHODY PULI RAZEM"), ("cost", "Koszty wspólne puli"),
                   ("net", "WYNIK PULI (EBITDA platformy)"), ("cum", "Wynik puli narastająco"), ("dist", "DO PODZIAŁU WEDŁUG KLUCZA"),
                   ("comm", "Prowizje od sprzedawców i dystrybutorów"), ("back", "Back margin od dostawców"), ("sub", "Usługi premium"), ("succ", "Opłata za sukces od GA"))}
    K = rows(c["Koszty platformy"])
    res["costs"] = {lab: [round(x or 0, 3) for x in K[lab][1:10]] for lab in K if isinstance(K[lab][2], (int, float))}
    res["mvp"] = [r[6] for r in c["Dopasowanie MVP1"].iter_rows(values_only=True) if r[1] in ("Razem Partner Technologiczny", "Razem Partner Marketplace")]
    res["pool"]["need"] = -min(0, min(res["pool"]["cum"]))
    res["users"] = [round(x or 0, 1) for x in U["Użytkownicy (średnio w roku, tys.)"][1:9]]
    res["gmv"] = [round(x or 0, 1) for x in U["GMV razem"][1:9]]
    print(tag, {k: (round(res[k]["sum6"], 1), round(res[k]["sum7"], 1), res[k]["payback"], round(res[k]["trough"], 1),
                    None if res[k]["irr7"] is None else round(res[k]["irr7"] * 100)) for k in ("PT", "PM", "GA")}, "need", round(res["pool"]["need"], 1))
    return res

import sys as _s
ORIG = {"MVP1: pełny koszt": 20000, "MVP1: analiza": 0, "MVP1: rezerwa": 0, "Średni miesięczny koszt FTE": 25000,
        "Rozwój platformy po starcie": [1500000, 2000000, 2000000, 3000000, 3000000, 3000000, 3000000],
        "Komunikacja": 0, "Bezpieczeństwo i zgodność": 0, "Płatności i ryzyko": 0, "Prawne, PMO": 0}
VERIF = dict(ORIG); VERIF.update({"MVP1: pełny koszt": 25000, "MVP1: analiza": 0.20, "Bezpieczeństwo i zgodność": 500000, "Prawne, PMO": 300000})
if len(_s.argv) > 2 and _s.argv[2] == "audit":
    A_ = {}
    for nm, ov in (("pierwotne", ORIG), ("zweryfikowane", VERIF), ("doszacowane", None)):
        A_[nm + "_GA"] = run(nm + "_GA", ov=ov); A_[nm + "_deal"] = run(nm + "_deal", mpt=0.10, mpm=0.10, ft=0.45, ov=ov)
    json.dump(A_, open("cost_audit_v5.json", "w"), ensure_ascii=False, indent=1); print("audit ok"); raise SystemExit
V = {}
V["base"] = run("base")                                        # model GA po korektach: po koszcie, PT 70% deficytu
V["deal"] = run("deal", mpt=0.10, mpm=0.10, ft=0.45)           # propozycja: marża 10%, finansowanie wg udziału w puli
V["deal_wal70"] = run("deal_wal70", mpt=0.10, mpm=0.10, ft=0.45, wal=0.7)
V["deal_adopt20"] = run("deal_adopt20", mpt=0.10, mpm=0.10, ft=0.45, adopt5=0.20)
V["deal_cac100"] = run("deal_cac100", mpt=0.10, mpm=0.10, ft=0.45, cac_add=100)
V["deal_syn"] = run("deal_syn", mpt=0.10, mpm=0.10, ft=0.45, syt=0.30, sym=0.25)
V["deal_m15"] = run("deal_m15", mpt=0.15, mpm=0.15, ft=0.45)
json.dump(V, open("partner_variants_v5.json", "w"), ensure_ascii=False, indent=1)
print("ok")
