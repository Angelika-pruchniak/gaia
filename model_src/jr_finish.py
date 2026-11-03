# Wykonywany na końcu build_bc_xlsx.py (exec): opisy formuł, komentarze, arkusz danych do tabel
# przestawnych, spis treści i profesjonalne formatowanie (bez siatki, linie tylko w tabelach).
from openpyxl.comments import Comment
from openpyxl.styles import Border, Side, Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo

FONT = "Trebuchet MS"
NAVY_ = "003665"; GRID = "BFBFBF"; LINK = "00704A"; GREYT = "595959"
thin = Side(style="thin", color=GRID); navy_thin = Side(style="thin", color=NAVY_); navy_med = Side(style="medium", color=NAVY_)
NUMF = '#,##0.0;[Red]-#,##0.0;"–"'

# ---------------------------------------------------------------- 4. Formatowanie
def font_of(c, **kw):
    f = c.font
    base = dict(name=FONT, size=kw.get("size", f.size if f and f.size and f.size != 11 else 10), bold=f.bold, italic=f.italic,
                color=f.color, underline=f.underline)
    base.update(kw); return Font(**base)

def border_table(ws, r1, r2, c1, c2, header=True):
    for r in range(r1, r2 + 1):
        vals = [ws.cell(r, c).value for c in range(c1, c2 + 1)]
        if all(v is None for v in vals): continue
        only_a = vals[0] is not None and all(v is None for v in vals[1:])
        is_head = header and r == r1
        bold = bool(ws.cell(r, c1).font and ws.cell(r, c1).font.bold)
        f1 = ws.cell(r, c1).font
        is_sec = only_a and (not bold or (f1.color is not None and f1.size and f1.size >= 12))
        for c in range(c1, c2 + 1):
            ce = ws.cell(r, c)
            if is_head:
                ce.font = Font(name=FONT, bold=True, color="FFFFFF", size=10); ce.fill = PatternFill("solid", fgColor=NAVY_)
                ce.alignment = Alignment(horizontal="center" if c > c1 else "left", vertical="center", wrap_text=True)
                ce.border = Border(left=navy_thin, right=navy_thin, top=navy_thin, bottom=navy_thin)
            elif is_sec:
                ce.border = Border(bottom=navy_med)
                if c == c1: ce.font = Font(name=FONT, bold=True, color=NAVY_, size=11)
            else:
                ce.border = Border(left=thin, right=thin, top=navy_thin if bold else thin, bottom=navy_thin if bold else thin)
                if c > c1 and ce.alignment.horizontal is None: ce.alignment = Alignment(horizontal="right", vertical="center")
                if c == c1 or isinstance(ce.value, str) and not str(ce.value).startswith("="):
                    ce.alignment = Alignment(horizontal="left", vertical="center", wrap_text=ce.alignment.wrap_text)

def finish(ws, regions, title_row=1, sub_row=2, tab=None, land=True):
    ws.sheet_view.showGridLines = False
    if tab: ws.sheet_properties.tabColor = tab
    for row in ws.iter_rows():
        for ce in row:
            if ce.value is None and not ce.fill: continue
            v = ce.value
            if isinstance(v, str) and v.startswith("=") and "!" in v and not (ce.font and ce.font.bold) and ce.font.color is None:
                ce.font = font_of(ce, color=LINK)
            else:
                ce.font = font_of(ce)
            if ce.number_format == NUM: ce.number_format = NUMF
    ws.cell(title_row, 1).font = Font(name=FONT, bold=True, size=16, color=NAVY_)
    if ws.cell(sub_row, 1).value: ws.cell(sub_row, 1).font = Font(name=FONT, italic=True, size=10, color=GREYT)
    for reg in regions: border_table(ws, *reg)
    ws.page_setup.orientation = "landscape" if land else "portrait"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.print_options.horizontalCentered = True


# ======== część sterowana zmiennymi z build_jr_xlsx.py: NOTES, CALC_SHEETS, SPEC, TOC_ITEMS, TOC_TITLE, TOC_SUB, ORDER, extra_finish()
for sh, _t in CALC_SHEETS:
    allnotes = dict(sh.notes); allnotes.update(NOTES.get(sh.ws.title, {}))
    for key, text in allnotes.items():
        if key not in sh.R: continue
        r = sh.R[key]; sh.ws.cell(r, 9, text)
        sh.ws.cell(r, 1).comment = Comment(text, "Model GAia", width=320, height=90)
    for key, r in sh.R.items():
        if sh.ws.cell(r, 9).value is None:
            f = str(sh.ws.cell(r, 3).value or "")
            if "!" in f: sh.ws.cell(r, 9, "Link do arkusza " + f.split("!")[0].lstrip("=").strip("'"))
for row in A.iter_rows(min_row=5, max_row=A.max_row):
    src, lab = row[7].value, row[0].value
    if lab and src and row[1].value is not None and not isinstance(row[1].value, str) or (lab and src and isinstance(row[1].value, str) and row[1].value.startswith("=")):
        row[1].comment = Comment(f"{lab}\nŹródło / uwaga: {src}", "Model GAia", width=320, height=110)

DS = wb.create_sheet("Dane")
DS["A1"] = "Dane źródłowe do tabel przestawnych (formuły z arkuszy obliczeniowych)"
DS["A2"] = "Każdy wiersz = jedna pozycja w jednym roku. Wartości to formuły; tabele przestawne odświeżają się przy otwarciu (Dane → Odśwież wszystko)."
for j, h in enumerate(["Obszar", "Typ", "Pozycja", "Rok", "Wartość (mln zł)", "Arkusz źródłowy"]): DS.cell(4, 1 + j, h)
dr = 5
for area, typ, pos, sh, key in SPEC:
    for yi, col in enumerate(["B"] + YC):
        DS.cell(dr, 1, area); DS.cell(dr, 2, typ); DS.cell(dr, 3, pos); DS.cell(dr, 4, f"Rok {yi}")
        DS.cell(dr, 5, f"=N('{sh.ws.title}'!{col}{sh.R[key]})").number_format = NUMF
        DS.cell(dr, 6, sh.ws.title); dr += 1
tab = Table(displayName="DaneGAia", ref=f"A4:F{dr-1}"); tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
DS.add_table(tab)
for col, w in zip("ABCDEF", [24, 34, 40, 9, 16, 26]): DS.column_dimensions[col].width = w
DS.freeze_panes = "A5"

TOC = wb.create_sheet("Spis treści", 0)
TOC["A1"] = TOC_TITLE; TOC["A2"] = TOC_SUB
TOC.cell(4, 1, "Arkusz"); TOC.cell(4, 2, "Co zawiera")
for i, (nm, desc) in enumerate(TOC_ITEMS):
    c = TOC.cell(5 + i, 1, nm); c.hyperlink = f"#'{nm}'!A1"; TOC.cell(5 + i, 2, desc)
lg = 5 + len(TOC_ITEMS) + 1
TOC.cell(lg, 1, "Legenda")
legend = [("1 234,5", Font(name=FONT, color="0000FF", size=10), PatternFill("solid", fgColor="FFF2CC"), "Założenie - można zmieniać"),
          ("1 234,5", Font(name=FONT, color="000000", size=10), None, "Formuła w tym samym arkuszu"),
          ("1 234,5", Font(name=FONT, color=LINK, size=10), None, "Formuła z innego arkusza (link)"),
          ("1 234,5", Font(name=FONT, bold=True, size=10), PatternFill("solid", fgColor="DDEBF7"), "Suma / wynik"),
          ("Komentarz", Font(name=FONT, size=10), None, "Czerwony trójkąt w rogu komórki = opis formuły lub źródła")]
for i, (t, fnt, fil, d) in enumerate(legend):
    TOC.cell(lg + 1 + i, 1, t); TOC.cell(lg + 1 + i, 2, d)
TOC.column_dimensions["A"].width = 30; TOC.column_dimensions["B"].width = 100

for sh, tabc in CALC_SHEETS:
    last = max(sh.R.values())
    finish(sh.ws, [(4, last, 1, 9)], tab=tabc)
    for r in range(last + 1, sh.ws.max_row + 1):
        for c in (1, 2):
            if sh.ws.cell(r, c).value is not None: sh.ws.cell(r, c).font = Font(name=FONT, size=10, bold=(c == 1), color=NAVY_)
    for r in range(5, last + 1): sh.ws.cell(r, 9).alignment = Alignment(wrap_text=True, vertical="center"); sh.ws.cell(r, 9).font = Font(name=FONT, size=9, italic=True, color=GREYT)
finish(A, [(4, A.max_row, 1, 8)], tab="FFC000")
for r in range(5, A.max_row + 1): A.cell(r, 8).alignment = Alignment(wrap_text=True, vertical="center")
extra_finish()
finish(DS, [], tab="BFBFBF")
for c in range(1, 7): DS.cell(4, c).font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
finish(TOC, [(4, 4 + len(TOC_ITEMS), 1, 2)], tab="003665", land=False)
TOC.cell(lg, 1).font = Font(name=FONT, bold=True, color=NAVY_, size=11)
for i in range(len(TOC_ITEMS)): TOC.cell(5 + i, 1).font = Font(name=FONT, color="0563C1", underline="single", size=10)
for i, (t, fnt, fil, d) in enumerate(legend):
    c = TOC.cell(lg + 1 + i, 1); c.font = fnt; c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    if fil: c.fill = fil
wb._sheets = [wb[n] for n in ORDER]
wb.active = 0
