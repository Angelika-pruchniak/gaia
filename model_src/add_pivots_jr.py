# Tworzy tabele przestawne (DataPilot) w LibreOffice na arkuszu "Dane" i zapisuje plik jako .xlsx.
# Użycie: python3 add_pivots.py <plik.xlsx>   (potrzebny LibreOffice z modułem uno)
import sys, os, time, subprocess, uno
from com.sun.star.beans import PropertyValue
from com.sun.star.table import CellRangeAddress, CellAddress
from com.sun.star.sheet.DataPilotFieldOrientation import ROW, COLUMN, DATA, PAGE
from com.sun.star.sheet.GeneralFunction import SUM

path = os.path.abspath(sys.argv[1]); url = uno.systemPathToFileUrl(path)
proc = subprocess.Popen(["soffice", "--headless", "--norestore", "--nologo", "--accept=socket,host=localhost,port=2002;urp;"])
ctx = None
for _ in range(60):
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"); break
    except Exception:
        time.sleep(1)
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
def pv(n, v): p = PropertyValue(); p.Name = n; p.Value = v; return p
doc = desktop.loadComponentFromURL(url, "_blank", 0, (pv("Hidden", True),))
doc.calculateAll()
sheets = doc.Sheets
src = sheets.getByName("Dane")
cur = src.createCursor(); cur.gotoEndOfUsedArea(False); last = cur.RangeAddress.EndRow
idx = [sheets.getByIndex(i).Name for i in range(sheets.Count)].index("Dane")
sheets.insertNewByName("Tabele przestawne", idx)
ps = sheets.getByName("Tabele przestawne")
ps.getCellByPosition(0, 0).String = "Tabele przestawne: dane z arkusza Dane, mln zł"
ps.getCellByPosition(0, 1).String = "Filtr 'Typ' u góry każdej tabeli można zmienić. Po zmianie założeń: Dane → Odśwież wszystko."
pivots = [("Wynik partnerów (gotówka)", "Wynik partnera", 4), ("Przychody puli wg strumienia", "Przychody puli", 6),
          ("Koszty wspólne puli", "Koszty wspólne", 7), ("Wynik i podział puli", "Wynik i podział puli", 5),
          ("Z czego zarabia Partner Technologiczny", "Zarobek PT", 4), ("Z czego zarabia Partner Marketplace", "Zarobek PM", 4),
          ("Z czego zarabia Grupa Azoty", "Zarobek GA", 4), ("GMV platformy", "GMV", 3)]
row = 3
tables = ps.getDataPilotTables(); placed = []
for i, (title, typ, n) in enumerate(pivots):
    row = 3 + (i // 2) * 60; col = 0 if i % 2 == 0 else 10
    d = tables.createDataPilotDescriptor()
    ra = CellRangeAddress(); ra.Sheet = idx + 1; ra.StartColumn = 0; ra.StartRow = 3; ra.EndColumn = 4; ra.EndRow = last
    d.SourceRange = ra
    f = d.DataPilotFields
    ft = f.getByIndex(1); ft.Orientation = PAGE
    fp = f.getByIndex(2); fp.Orientation = ROW
    fy = f.getByIndex(3); fy.Orientation = COLUMN
    fv = f.getByIndex(4); fv.Orientation = DATA; fv.Function = SUM
    d.setPropertyValue("ColumnGrand", typ not in ("Wynik i podział puli", "Wynik partnera")); d.setPropertyValue("RowGrand", True)
    try: d.setPropertyValue("GrandTotalName", "Razem")
    except Exception as e: print("gt", e)
    try: d.setPropertyValue("ShowFilterButton", False)
    except Exception: pass
    ca = CellAddress(); ca.Sheet = idx; ca.Column = col; ca.Row = row + 1
    tables.insertNewByName(f"Przestawna{i+1}", ca, d)
    dp = tables.getByName(f"Przestawna{i+1}")
    pf = dp.DataPilotFields.getByIndex(1)
    items = pf.Items
    for nm in items.ElementNames:
        items.getByName(nm).IsHidden = (nm != typ)
    try:
        pf.setPropertyValue("UseSelectedPage", True); pf.setPropertyValue("SelectedPage", typ)
    except Exception as e:
        pass
    try:
        dfl = dp.DataPilotFields.getByIndex(4)
        for nm in ("LayoutName",):
            pass
    except Exception: pass
    for df in dp.DataFields:
        try: df.setPropertyValue("LayoutName", "mln zł")
        except Exception as e: print("ln", e)
    for nm_i, nm in ((2, "Pozycja"), (3, "Rok")):
        try: dp.DataPilotFields.getByIndex(nm_i).setPropertyValue("LayoutName", nm)
        except Exception: pass
    dp.refresh()
    placed.append((dp, title, row, col))
# wyczyść pozostałości po wstawianiu i wpisz tytuły
from com.sun.star.sheet.CellFlags import VALUE, STRING, FORMULA, DATETIME
for dp, title, r0, c0 in placed:
    out = dp.OutputRange
    endr = r0 + 59
    if out.EndRow + 1 <= endr:
        ps.getCellRangeByPosition(c0, out.EndRow + 1, c0 + 8, endr).clearContents(VALUE | STRING | FORMULA | DATETIME)
    ps.getCellByPosition(c0, r0).String = f"{placed.index((dp, title, r0, c0)) + 1}. {title}"
    ps.getCellByPosition(c0, r0).CharWeight = 150; ps.getCellByPosition(c0, r0).CharColor = 0x003665; ps.getCellByPosition(c0, r0).CharHeight = 12
    ps.getCellByPosition(c0, r0).CharFontName = "Trebuchet MS"
    rng = ps.getCellRangeByPosition(c0 + 1, out.StartRow, c0 + 7, out.EndRow); rng.NumberFormat = doc.NumberFormats.queryKey('#,##0.0;[RED]-#,##0.0', doc.CharLocale, False) if doc.NumberFormats.queryKey('#,##0.0;[RED]-#,##0.0', doc.CharLocale, False) >= 0 else doc.NumberFormats.addNew('#,##0.0;[RED]-#,##0.0', doc.CharLocale)
for c in range(20): ps.Columns.getByIndex(c).Width = 2200
ps.Columns.getByIndex(0).Width = 7800; ps.Columns.getByIndex(10).Width = 7800; ps.Columns.getByIndex(9).Width = 800
ps.getCellByPosition(0, 0).CharWeight = 150; ps.getCellByPosition(0, 0).CharHeight = 16; ps.getCellByPosition(0, 0).CharColor = 0x003665; ps.getCellByPosition(0, 0).CharFontName = "Trebuchet MS"
ctl = doc.CurrentController
try:
    ctl.setActiveSheet(ps); ctl.ShowGrid = False
except Exception as e: print("grid", e)
ctl.setActiveSheet(sheets.getByIndex(0))
doc.calculateAll()
doc.storeToURL(url, (pv("FilterName", "Calc MS Excel 2007 XML"),))
doc.close(True)
try: desktop.terminate()
except Exception: pass
proc.wait(timeout=30)
print("ok pivots", len(pivots))

# poprawki XML: brak siatki na arkuszu tabel przestawnych, polska nazwa pola danych
import zipfile, shutil, re
tmp = path + ".tmp"
with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    wbxml = zin.read("xl/workbook.xml").decode("utf8")
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rid = re.search(r'<sheet name="Tabele przestawne" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"', wbxml).group(1)
    target = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    sheet_xml = "xl/" + target.group(1).lstrip("/").replace("xl/", "")
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == sheet_xml:
            t = data.decode("utf8")
            t = re.sub(r'showGridLines="(true|1)"', 'showGridLines="false"', t)
            if "showGridLines" not in t: t = t.replace("<sheetView ", '<sheetView showGridLines="false" ', 1)
            data = t.encode("utf8")
        if item.filename.startswith("xl/pivotCache/pivotCacheDefinition"):
            t = data.decode("utf8")
            if "refreshOnLoad" not in t: t = t.replace("<pivotCacheDefinition ", '<pivotCacheDefinition refreshOnLoad="1" ', 1)
            data = t.encode("utf8")
        if item.filename.startswith("xl/pivotTables/") or item.filename == "xl/sharedStrings.xml":
            data = data.decode("utf8").replace("Sum - Wartość (mln zł)", "Suma, mln zł").encode("utf8")
        zout.writestr(item, data)
shutil.move(tmp, path)
print("xml ok", sheet_xml)
