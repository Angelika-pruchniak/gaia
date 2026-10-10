# Generuje 2_Kalkulator_Zysku.xlsx (formuły, przykład Studio Lumen)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
import os

NAVY='1E2A4A'; ORANGE='F8B68C'; PEACH='FDDCC6'; PINK='F2BFC5'; BLUSH='FBE6E7'; CREAM='FFF8F3'; GREY='6B6F80'
F=lambda c: PatternFill('solid', start_color=c, end_color=c)
thin=Side(style='thin', color='E9D6CF'); B=Border(left=thin,right=thin,top=thin,bottom=thin)
PLN='#,##0 "zł"'; PLN2='#,##0.00 "zł"'; PCT='0.0%'; INT='#,##0'

wb=Workbook()

def title(ws, t, sub, width=10):
    ws.sheet_view.showGridLines=False
    ws['B2']=t; ws['B2'].font=Font(name='Calibri',size=20,bold=True,color=NAVY)
    ws['B3']=sub; ws['B3'].font=Font(size=10,color=GREY,italic=True)
    for c in range(2,2+width):
        ws.cell(row=1,column=c).fill=F(ORANGE if c%2 else PINK)
    ws.row_dimensions[1].height=8
    ws.column_dimensions['A'].width=2

def header(ws,row,cols,start=2,color=NAVY,fc='FFFFFF'):
    for i,h in enumerate(cols):
        c=ws.cell(row=row,column=start+i,value=h)
        c.font=Font(bold=True,color=fc); c.fill=F(color); c.alignment=Alignment(wrap_text=True,vertical='center',horizontal='center'); c.border=B
    ws.row_dimensions[row].height=32

def inp(c,fmt=None):
    c.fill=F(PEACH); c.border=B; c.font=Font(color=NAVY,bold=True)
    if fmt: c.number_format=fmt

def out(c,fmt=None,bold=False):
    c.border=B; c.font=Font(color=NAVY,bold=bold)
    if fmt: c.number_format=fmt

def label(ws,ref,text,bold=False):
    ws[ref]=text; ws[ref].font=Font(bold=bold,color=NAVY); ws[ref].alignment=Alignment(wrap_text=True,vertical='center')

def legend(ws,row):
    ws.cell(row=row,column=2,value='   ').fill=F(PEACH)
    ws.cell(row=row,column=3,value='← pole do wpisania (dane przykładowe: fikcyjne „Studio Lumen”, zastąp swoimi)').font=Font(size=9,color=GREY,italic=True)

# ---------------- START ----------------
ws=wb.active; ws.title='Start'
title(ws,'Kierunek Zysk · Kalkulator Zysku','Materiał 2 z 6 · do ćwiczeń z modułów 1, 4 i 5 przewodnika',8)
ws.column_dimensions['B'].width=26; ws.column_dimensions['C'].width=80
rows=[('Jak korzystać',''),
('1. Ustawienia','Wpisz wynagrodzenia, koszty stałe i godziny. Arkusz policzy pełny koszt godziny pracy zespołu, potrzebny wszędzie dalej.'),
('2. Usługi','Rentowność każdej usługi: marża, zysk na godzinę, efektywna stawka. Moduł 1 i 2 (usługa flagowa).'),
('3. Klienci','Marża na każdym kliencie, udział w przychodzie i automatyczna ćwiartka macierzy (Gwiazda / Okazja / Do rozwoju / Do pożegnania). Moduł 1.'),
('4. Podwyżka cen','Ile zarobisz po podwyżce i ilu klientów możesz stracić, żeby nadal być na plusie. Moduł 4.'),
('5. Lejek','Ile zapytań, rozmów i ofert potrzebujesz, żeby osiągnąć cel przychodu. Moduł 5.'),
('6. Dźwignie wzrostu','Efekt poprawy 5 dźwigni (klienci, wartość, częstotliwość, długość, marża) na przychód i zysk.'),
('7. Próg rentowności','Minimalny przychód miesięczny i przychód potrzebny do zysku docelowego.'),
('',''),
('Kolory','Brzoskwiniowe pola = wpisujesz. Białe pola = liczą się same (nie nadpisuj formuł).'),
('Dane','Wystarczą przybliżenia z faktur i kalendarza z ostatnich 12 miesięcy. Kwoty netto.'),
('Licencja','Do użytku w jednej firmie nabywcy. © guliwera4.pl'),]
for i,(a,b) in enumerate(rows):
    r=5+i; ws.cell(row=r,column=2,value=a).font=Font(bold=True,color=NAVY,size=12 if i==0 else 11)
    c=ws.cell(row=r,column=3,value=b); c.alignment=Alignment(wrap_text=True,vertical='top')
    ws.row_dimensions[r].height=30 if b else 20
    if i in (1,3,5,7): 
        ws.cell(row=r,column=2).fill=F(BLUSH); c.fill=F(BLUSH)
    elif 1<=i<=7:
        ws.cell(row=r,column=2).fill=F(CREAM); c.fill=F(CREAM)
ws.cell(row=14,column=2).fill=F(PEACH)

# ---------------- USTAWIENIA ----------------
ws=wb.create_sheet('Ustawienia')
title(ws,'Ustawienia firmy','Pełny koszt godziny = (wynagrodzenia + koszty stałe) ÷ godziny fakturowalne',4)
ws.column_dimensions['B'].width=52; ws.column_dimensions['C'].width=20; ws.column_dimensions['D'].width=50
legend(ws,4)
data=[('Wynagrodzenia zespołu z narzutami (rocznie, z właścicielem)',640000,PLN,'Uwzględnij też godziwe wynagrodzenie właściciela'),
('Pozostałe koszty stałe (rocznie: biuro, narzędzia, księgowość, auta)',215000,PLN,''),
('Liczba osób realizujących usługi',6,INT,''),
('Godziny fakturowalne na osobę rocznie',1500,INT,'Realnie 1 200–1 500 h (nie 2 000): urlopy, administracja, sprzedaż'),]
for i,(l,v,f,n) in enumerate(data):
    r=6+i; label(ws,f'B{r}',l); inp(ws.cell(row=r,column=3,value=v),f); ws.cell(row=r,column=4,value=n).font=Font(size=9,color=GREY,italic=True)
label(ws,'B11','Godziny fakturowalne zespołu rocznie',True); ws['C11']='=C8*C9'; out(ws['C11'],INT,True)
label(ws,'B12','PEŁNY KOSZT GODZINY PRACY',True); ws['C12']='=IFERROR((C6+C7)/C11,0)'; out(ws['C12'],PLN2,True)
ws['B12'].fill=F(ORANGE); ws['C12'].fill=F(ORANGE); ws['C12'].font=Font(bold=True,size=13,color=NAVY)
label(ws,'B14','Próg marży „zdrowego” klienta (do macierzy klientów)'); inp(ws.cell(row=14,column=3,value=0.25),PCT)
ws['D14']='Klient z marżą powyżej progu = wysoki zysk'; ws['D14'].font=Font(size=9,color=GREY,italic=True)

# ---------------- USŁUGI ----------------
ws=wb.create_sheet('Usługi')
title(ws,'Rentowność usług','Która usługa daje najwyższy zysk na godzinę? Kandydat na usługę flagową (moduł 2).',12)
legend(ws,4)
cols=['Usługa / pakiet','Cena netto','Godziny zespołu na 1 sprzedaż','Koszty bezpośrednie (podwykonawcy, licencje)','Koszt pracy','Marża zł','Marża %','Zysk na godzinę','Efektywna stawka (cena ÷ h)','Sprzedaży rocznie','Marża roczna','Ranking zysku/h']
header(ws,6,cols)
widths=[30,13,14,18,13,13,10,13,14,12,14,10]
for i,w in enumerate(widths): ws.column_dimensions[chr(66+i)].width=w
sample=[('Pakiet „Gotowi na pacjentów”',16900,120,500,8),('Identyfikacja wizualna',12000,90,0,14),('Strona internetowa',15000,140,1500,10),('Pakiet „ulotka / drobne”',800,9,0,60),('Kampania social (miesiąc)',3500,30,0,40),('Opieka nad stroną (miesiąc)',900,6,100,36)]
for r in range(7,19):
    i=r-7; s=sample[i] if i<len(sample) else (None,None,None,None,None)
    inp(ws.cell(row=r,column=2,value=s[0])); ws.cell(row=r,column=2).font=Font(color=NAVY)
    for col,val,fmt in ((3,s[1],PLN),(4,s[2],INT),(5,s[3],PLN),(11,s[4],INT)): inp(ws.cell(row=r,column=col,value=val),fmt)
    ws.cell(row=r,column=6,value=f'=IF(C{r}="","",D{r}*Ustawienia!$C$12)'); out(ws.cell(row=r,column=6),PLN)
    ws.cell(row=r,column=7,value=f'=IF(C{r}="","",C{r}-F{r}-E{r})'); out(ws.cell(row=r,column=7),PLN)
    ws.cell(row=r,column=8,value=f'=IF(OR(C{r}="",C{r}=0),"",G{r}/C{r})'); out(ws.cell(row=r,column=8),PCT)
    ws.cell(row=r,column=9,value=f'=IF(OR(D{r}="",D{r}=0),"",G{r}/D{r})'); out(ws.cell(row=r,column=9),PLN)
    ws.cell(row=r,column=10,value=f'=IF(OR(D{r}="",D{r}=0),"",C{r}/D{r})'); out(ws.cell(row=r,column=10),PLN)
    ws.cell(row=r,column=12,value=f'=IF(OR(G{r}="",K{r}=""),"",G{r}*K{r})'); out(ws.cell(row=r,column=12),PLN)
    ws.cell(row=r,column=13,value=f'=IF(I{r}="","",RANK(I{r},$I$7:$I$18))'); out(ws.cell(row=r,column=13),INT)
ws['B19']='RAZEM'; ws['B19'].font=Font(bold=True,color='FFFFFF'); 
for c in range(2,14): ws.cell(row=19,column=c).fill=F(NAVY); ws.cell(row=19,column=c).font=Font(bold=True,color='FFFFFF')
ws['L19']='=SUM(L7:L18)'; ws['L19'].number_format=PLN
ws['K19']='=SUM(K7:K18)'; ws['K19'].number_format=INT
ws.conditional_formatting.add('I7:I18',CellIsRule(operator='lessThan',formula=['0'],fill=F(PINK),font=Font(bold=True,color='9C2B3B')))
ws.conditional_formatting.add('M7:M18',CellIsRule(operator='equal',formula=['1'],fill=F(ORANGE),font=Font(bold=True,color=NAVY)))
ws['B21']='Jak czytać: usługa z ujemnym zyskiem na godzinę (na różowo) dokłada do firmy. Podnieś cenę, zmień zakres albo przestań ją sprzedawać. Ranking 1 (pomarańczowy) = najlepszy kandydat na usługę flagową.'
ws['B21'].font=Font(size=9,italic=True,color=GREY); ws.merge_cells('B21:M22'); ws['B21'].alignment=Alignment(wrap_text=True,vertical='top')

# ---------------- KLIENCI ----------------
ws=wb.create_sheet('Klienci')
title(ws,'Rentowność klientów i macierz','Kto naprawdę zarabia dla firmy? Ćwiczenia 1.2 i 1.3 przewodnika.',11)
legend(ws,4)
cols=['Klient','Przychód 12 mies.','Godziny zespołu 12 mies.','Koszty bezpośrednie','Wysiłek 1–5 (poprawki, telefony, płatności)','Marża zł','Marża %','Zysk na godzinę','Udział w przychodzie','Ćwiartka macierzy','Ranking marży']
header(ws,6,cols)
for i,w in enumerate([24,14,13,14,16,13,10,13,12,18,10]): ws.column_dimensions[chr(66+i)].width=w
sample=[('Klient A (sieć sklepów)',210000,2100,8000,5),('Klient B (deweloper)',150000,1500,12000,4),('Klient C (producent)',132000,900,0,2),('Gabinet D',48000,250,0,1),('Klinika E',62000,330,2000,2),('Kancelaria F',36000,300,0,4),('Szkoła G',24000,200,0,2),('Hotel H',41000,420,0,5),('Producent I',90000,520,0,4)]
dv=DataValidation(type='whole',operator='between',formula1='1',formula2='5',allow_blank=True); ws.add_data_validation(dv)
for r in range(7,32):
    i=r-7; s=sample[i] if i<len(sample) else (None,None,None,None,None)
    inp(ws.cell(row=r,column=2,value=s[0])); ws.cell(row=r,column=2).font=Font(color=NAVY)
    for col,val,fmt in ((3,s[1],PLN),(4,s[2],INT),(5,s[3],PLN),(6,s[4],INT)): inp(ws.cell(row=r,column=col,value=val),fmt)
    dv.add(f'F{r}')
    ws.cell(row=r,column=7,value=f'=IF(C{r}="","",C{r}-D{r}*Ustawienia!$C$12-E{r})'); out(ws.cell(row=r,column=7),PLN)
    ws.cell(row=r,column=8,value=f'=IF(OR(C{r}="",C{r}=0),"",G{r}/C{r})'); out(ws.cell(row=r,column=8),PCT)
    ws.cell(row=r,column=9,value=f'=IF(OR(D{r}="",D{r}=0),"",G{r}/D{r})'); out(ws.cell(row=r,column=9),PLN)
    ws.cell(row=r,column=10,value=f'=IF(C{r}="","",C{r}/$C$32)'); out(ws.cell(row=r,column=10),PCT)
    ws.cell(row=r,column=11,value=f'=IF(OR(C{r}="",F{r}=""),"",IF(H{r}>=Ustawienia!$C$14,IF(F{r}<=3,"⭐ Gwiazda","💎 Okazja"),IF(F{r}<=3,"🌱 Do rozwoju","🧹 Do pożegnania / wyceny")))'); out(ws.cell(row=r,column=11))
    ws.cell(row=r,column=12,value=f'=IF(G{r}="","",RANK(G{r},$G$7:$G$31))'); out(ws.cell(row=r,column=12),INT)
for c in range(2,13): ws.cell(row=32,column=c).fill=F(NAVY); ws.cell(row=32,column=c).font=Font(bold=True,color='FFFFFF')
ws['B32']='RAZEM'; ws['C32']='=SUM(C7:C31)'; ws['C32'].number_format=PLN; ws['D32']='=SUM(D7:D31)'; ws['D32'].number_format=INT
ws['G32']='=SUM(G7:G31)'; ws['G32'].number_format=PLN; ws['H32']='=IFERROR(G32/C32,0)'; ws['H32'].number_format=PCT
ws.conditional_formatting.add('K7:K31',FormulaRule(formula=['LEFT(K7,1)="⭐"'],fill=F(ORANGE)))
ws.conditional_formatting.add('K7:K31',FormulaRule(formula=['ISNUMBER(SEARCH("pożegnania",K7))'],fill=F(PINK)))
ws.conditional_formatting.add('K7:K31',FormulaRule(formula=['ISNUMBER(SEARCH("Okazja",K7))'],fill=F(PEACH)))
ws.conditional_formatting.add('K7:K31',FormulaRule(formula=['ISNUMBER(SEARCH("rozwoju",K7))'],fill=F(BLUSH)))
ws.conditional_formatting.add('J7:J31',CellIsRule(operator='greaterThan',formula=['0.2'],font=Font(bold=True,color='9C2B3B')))
label(ws,'B34','Udział 3 największych klientów',True); ws['E34']='=IFERROR((LARGE(C7:C31,1)+LARGE(C7:C31,2)+LARGE(C7:C31,3))/C32,0)'; out(ws['E34'],PCT,True)
ws['F34']='Powyżej 40% = ryzyko koncentracji'; ws['F34'].font=Font(size=9,italic=True,color=GREY)
ws['B36']='Udział w przychodzie powyżej 20% jest pogrubiony na czerwono. Ćwiartkę liczy próg marży z arkusza Ustawienia i wysiłek (1–3 = mały, 4–5 = duży).'
ws['B36'].font=Font(size=9,italic=True,color=GREY); ws.merge_cells('B36:L37'); ws['B36'].alignment=Alignment(wrap_text=True,vertical='top')

# ---------------- PODWYŻKA ----------------
ws=wb.create_sheet('Podwyżka cen')
title(ws,'Symulator podwyżki cen','Moduł 4 · dopuszczalny spadek sprzedaży = podwyżka ÷ (marża na pokrycie + podwyżka)',9)
ws.column_dimensions['B'].width=50; ws.column_dimensions['C'].width=18
for col in 'DEFGHIJ': ws.column_dimensions[col].width=11
legend(ws,4)
items=[('Przychód roczny obecnie',1200000,PLN),('Marża na pokrycie % (cena minus koszty zmienne zleceń)',0.40,PCT),('Zysk obecnie (przed opodatkowaniem)',96000,PLN),('Planowana podwyżka cen',0.10,PCT),('Przewidywana utrata sprzedaży (wolumenu)',0.05,PCT)]
for i,(l,v,f) in enumerate(items):
    r=6+i; label(ws,f'B{r}',l); inp(ws.cell(row=r,column=3,value=v),f)
label(ws,'B12','Koszty stałe (wyliczone)'); ws['C12']='=C6*C7-C8'; out(ws['C12'],PLN)
label(ws,'B13','Przychód po zmianie'); ws['C13']='=C6*(1-C10)*(1+C9)'; out(ws['C13'],PLN)
label(ws,'B14','Zysk po zmianie',True); ws['C14']='=C13-C6*(1-C7)*(1-C10)-C12'; out(ws['C14'],PLN,True)
label(ws,'B15','Zmiana zysku',True); ws['C15']='=C14-C8'; out(ws['C15'],PLN,True)
label(ws,'B16','Zmiana zysku %'); ws['C16']='=IFERROR(C15/ABS(C8),0)'; out(ws['C16'],PCT)
label(ws,'B17','DOPUSZCZALNA utrata sprzedaży (zysk bez zmian)',True); ws['C17']='=IFERROR(C9/(C7+C9),0)'; out(ws['C17'],PCT,True)
for r in (14,17): ws[f'B{r}'].fill=F(ORANGE); ws[f'C{r}'].fill=F(ORANGE)
ws.conditional_formatting.add('C15',CellIsRule(operator='lessThan',formula=['0'],fill=F(PINK)))
ws['B19']='Tabela: dopuszczalna utrata sprzedaży przy różnych marżach i podwyżkach'; ws['B19'].font=Font(bold=True,color=NAVY,size=12)
pods=[0.03,0.05,0.08,0.10,0.12,0.15,0.20]; margs=[0.2,0.3,0.4,0.5,0.6]
ws['B20']='Marża na pokrycie ↓ / podwyżka →'; ws['B20'].font=Font(bold=True,color='FFFFFF'); ws['B20'].fill=F(NAVY)
for j,p in enumerate(pods):
    c=ws.cell(row=20,column=3+j,value=p); c.number_format='0%'; c.font=Font(bold=True,color='FFFFFF'); c.fill=F(NAVY); c.alignment=Alignment(horizontal='center')
for i,m in enumerate(margs):
    r=21+i; c=ws.cell(row=r,column=2,value=m); c.number_format='0%'; c.font=Font(bold=True,color=NAVY); c.fill=F(PEACH if i%2 else BLUSH); c.alignment=Alignment(horizontal='center')
    for j in range(len(pods)):
        col=chr(67+j); cc=ws.cell(row=r,column=3+j,value=f'=-{col}$20/($B{r}+{col}$20)'); cc.number_format='0%'; cc.border=B; cc.alignment=Alignment(horizontal='center')
ws['B27']='Przykład: przy marży 40% i podwyżce 10% możesz stracić do 20% sprzedaży i nadal zarobisz tyle samo. Zwykle odchodzą klienci z ćwiartki „Do pożegnania”.'
ws['B27'].font=Font(size=9,italic=True,color=GREY); ws.merge_cells('B27:J28'); ws['B27'].alignment=Alignment(wrap_text=True,vertical='top')

# ---------------- LEJEK ----------------
ws=wb.create_sheet('Lejek')
title(ws,'Kalkulator lejka sprzedaży','Moduł 5 · ile aktywności potrzeba do celu przychodu',6)
ws.column_dimensions['B'].width=46
for col in 'CDE': ws.column_dimensions[col].width=18
legend(ws,4)
header(ws,6,['','Dziś','Cel za 90 dni','Uwagi'],start=2)
ws.column_dimensions['E'].width=44
rows=[('Cel / obecny przychód ze sprzedaży nowych zleceń (miesięcznie)',60000,80000,PLN,'Bez stałych abonamentów'),
('Średnia wartość wygranego zlecenia',9500,12000,PLN,'Pakiety podnoszą wartość'),
('Skuteczność ofert (wygrane ÷ oferty)',0.30,0.40,PCT,'Rozmowa przed ofertą + follow-up'),
('% rozmów kończących się ofertą',0.70,0.75,PCT,'Kwalifikacja'),
('% zapytań kończących się rozmową',0.60,0.80,PCT,'Szybka odpowiedź, link do kalendarza'),]
for i,(l,a,b,f,n) in enumerate(rows):
    r=7+i; label(ws,f'B{r}',l); inp(ws.cell(row=r,column=3,value=a),f); inp(ws.cell(row=r,column=4,value=b),f); ws.cell(row=r,column=5,value=n).font=Font(size=9,italic=True,color=GREY)
calc=[('Potrzebne wygrane zlecenia / mies.','=IFERROR({c}7/{c}8,0)'),('Potrzebne oferty / mies.','=IFERROR({c}13/{c}9,0)'),('Potrzebne rozmowy / mies.','=IFERROR({c}14/{c}10,0)'),('POTRZEBNE ZAPYTANIA / mies.','=IFERROR({c}15/{c}11,0)'),('Potrzebne zapytania / tydzień','={c}16/4.33')]
for i,(l,f) in enumerate(calc):
    r=13+i; label(ws,f'B{r}',l,i==3)
    for c in 'CD': ws[f'{c}{r}']=f.format(c=c); out(ws[f'{c}{r}'],'0.0',i==3)
for c in 'BCD': ws[f'{c}16'].fill=F(ORANGE)
ws['B19']='Wniosek: porównaj kolumny. Często poprawa skuteczności i wartości zlecenia daje więcej niż szukanie dodatkowych zapytań.'
ws['B19'].font=Font(size=9,italic=True,color=GREY); ws.merge_cells('B19:E20'); ws['B19'].alignment=Alignment(wrap_text=True,vertical='top')

# ---------------- DŹWIGNIE ----------------
ws=wb.create_sheet('Dźwignie wzrostu')
title(ws,'5 dźwigni wzrostu zysku','Moduł 5 · przychód = klienci × wartość × częstotliwość × długość współpracy; do tego marża',5)
ws.column_dimensions['B'].width=44
for col in 'CDE': ws.column_dimensions[col].width=18
legend(ws,4)
header(ws,6,['Dźwignia','Dziś','Poprawa %','Po zmianie'])
lev=[('Liczba aktywnych klientów',38,0.10,INT),('Średnia wartość zlecenia',9500,0.10,PLN),('Liczba zleceń na klienta rocznie',3.33,0.10,'0.00'),('Retencja / długość współpracy (mnożnik)',1.0,0.10,'0.00'),('Marża na pokrycie %',0.40,0.05,PCT)]
for i,(l,v,p,f) in enumerate(lev):
    r=7+i; label(ws,f'B{r}',l); inp(ws.cell(row=r,column=3,value=v),f); inp(ws.cell(row=r,column=4,value=p),PCT)
    ws[f'E{r}']=f'=C{r}*(1+D{r})' if i<4 else f'=MIN(0.95,C{r}+D{r})'; out(ws[f'E{r}'],f)
ws['B11'].value='Marża na pokrycie % (poprawa w punktach procentowych)'
label(ws,'B13','Koszty stałe roczne'); inp(ws.cell(row=13,column=3,value=384000),PLN)
label(ws,'B15','Przychód roczny',True); ws['C15']='=C7*C8*C9*C10'; ws['E15']='=E7*E8*E9*E10'
label(ws,'B16','Zysk roczny',True); ws['C16']='=C15*C11-C13'; ws['E16']='=E15*E11-C13'
label(ws,'B17','Zmiana przychodu'); ws['E17']='=IFERROR(E15/C15-1,0)'
label(ws,'B18','Zmiana zysku',True); ws['E18']='=IFERROR(E16/C16-1,0)'
for r in (15,16): out(ws[f'C{r}'],PLN,True); out(ws[f'E{r}'],PLN,True)
out(ws['E17'],PCT); out(ws['E18'],PCT,True)
for c in 'BCDE': ws[f'{c}18'].fill=F(ORANGE)
ws['B20']='Zobacz, jak niewielka poprawa każdej dźwigni (bez nowych klientów: wartość, częstotliwość, długość, marża) zwielokrotnia zysk przy stałych kosztach.'
ws['B20'].font=Font(size=9,italic=True,color=GREY); ws.merge_cells('B20:E21'); ws['B20'].alignment=Alignment(wrap_text=True,vertical='top')

# ---------------- PRÓG ----------------
ws=wb.create_sheet('Próg rentowności')
title(ws,'Próg rentowności i przychód docelowy','Ile musisz sprzedawać miesięcznie, żeby wyjść na zero i żeby osiągnąć zysk docelowy',4)
ws.column_dimensions['B'].width=50; ws.column_dimensions['C'].width=20
legend(ws,4)
label(ws,'B6','Koszty stałe miesięcznie (z wynagrodzeniami)'); inp(ws.cell(row=6,column=3,value=32000),PLN)
label(ws,'B7','Marża na pokrycie %'); inp(ws.cell(row=7,column=3,value=0.40),PCT)
label(ws,'B8','Zysk docelowy miesięcznie'); inp(ws.cell(row=8,column=3,value=18000),PLN)
label(ws,'B10','PRÓG RENTOWNOŚCI (przychód / mies.)',True); ws['C10']='=IFERROR(C6/C7,0)'; out(ws['C10'],PLN,True)
label(ws,'B11','Przychód potrzebny do zysku docelowego',True); ws['C11']='=IFERROR((C6+C8)/C7,0)'; out(ws['C11'],PLN,True)
label(ws,'B12','Średnia wartość zlecenia'); inp(ws.cell(row=12,column=3,value=9500),PLN)
label(ws,'B13','Zleceń miesięcznie do zysku docelowego'); ws['C13']='=IFERROR(C11/C12,0)'; out(ws['C13'],'0.0')
for c in 'BC': ws[f'{c}10'].fill=F(PEACH); ws[f'{c}11'].fill=F(ORANGE)

for s in wb.worksheets:
    s.sheet_properties.tabColor = {'Start':NAVY,'Ustawienia':NAVY}.get(s.title, ORANGE if wb.worksheets.index(s)%2 else PINK)
    s.freeze_panes=None
out_path=os.path.join(os.path.dirname(__file__),'..','2_Kierunek_Zysk_Kalkulator.xlsx')
wb.save(out_path); print('saved',out_path)
