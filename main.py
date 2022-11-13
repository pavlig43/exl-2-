import pandas as pd
from openpyxl import *
from openpyxl.utils import get_column_letter
import datetime
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import *
from nomen import *
from openpyxl.styles import NamedStyle
import test
import re
import copy
import numpy as np

ost = load_workbook('йцу.xlsx')
ost_pt = ost.create_sheet('Остатки полутуш', 0)
ost_mb = ost.create_sheet('Остатки по датам молодняк', -1)
ost_vk = ost.create_sheet('Остаток по датам коровы', -1)
date = datetime.datetime.now().strftime('%d.%m.%Y')
birn = ost['TDSheet']
name = f'Остатки на {date}'
ost_pt['A1'] = 'Остатки холодильника хранения туш '


def format(lst, let, num, height, width, size, bold):

	lst.row_dimensions[num].height = height
	lst.column_dimensions[str(let)].width = width
	lst[let + str(num)].alignment = Alignment(horizontal='center')
	lst[let + str(num)].font = Font(size=size, bold=bold)
	lst[let + str(num)].alignment = Alignment(wrap_text=True)


format(ost_pt, 'A', 1, 23, 65, 16, True)


def oform(lst, let, num, thickless):
	thack = Side(border_style=thickless, color="0A0000")
	lst[let + str(num)].border = Border(top=thack,
	                                    left=thack,
	                                    right=thack,
	                                    bottom=thack)


oform(ost_pt, 'A', 1, 'thick')

ost_pt.merge_cells('A1:B1')

ost_pt['B2'] = 'Кол-во туш'
format(ost_pt, 'B', 2, 18, 26, 14, True)
oform(ost_pt, 'B', 2, 'medium')
ost_pt['D2'] = 'Кол-во туш'
format(ost_pt, 'D', 2, 18, 26, 14, True)
oform(ost_pt, 'D', 2, 'medium')
ost_pt['D1'] = 'Сан.камера'
format(ost_pt, 'D', 1, 18, 37, 16, True)
oform(ost_pt, 'D', 1, 'medium')
for i in range(1, 14):
	ost_pt['D' + str(i)].fill = PatternFill('solid', start_color="F7FF00")
for i in range(3, 14):
	for j in range(1, 5):
		format(ost_pt, get_column_letter(j), i, 18, 36, 11, True)
		oform(ost_pt, get_column_letter(j), i, 'thin')
for i in range(1, 4):
	ost_pt[get_column_letter(i) + str(2)].fill = PatternFill('solid',
	                                                         start_color="75F286")
ost_pt['C2'] = 'Дата'
ost_pt['C2'].font = Font(size=16, bold=True)

for i in range(3, 13):
	ost_pt['A' + str(i)] = list(sv.keys())[i - 3]

format(ost_pt, 'A', 13, 23, 65, 16, True)
oform(ost_pt, 'A', 13, 'thick')
ost_pt['A13'] = 'Итого'
for i in [7, 10, 13]:
	for j in range(1, 4):
		ost_pt[get_column_letter(j) + str(i)].fill = PatternFill(
		 'solid', start_color='BCBCBC')
for i in range(19, 32):
	for j in range(1, 3):
		format(ost_pt, get_column_letter(j), i, 18, 36, 11, True)
		oform(ost_pt, get_column_letter(j), i, 'thin')

format(ost_pt, 'A', 32, 23, 65, 16, True)
oform(ost_pt, 'A', 32, 'thick')
ost_pt['A32'] = 'Итого'
ost_pt['A32'].fill = PatternFill('solid', start_color='BCBCBC')
ost_pt['B32'].fill = PatternFill('solid', start_color='BCBCBC')
format(ost_pt, 'A', 13, 23, 65, 16, False)

for i in range(19, 32):
	ost_pt['A' + str(i)] = KrsSV[i - 19]

ost_pt['A14'] = 'Свиноматки б/ш'
ost_pt['A15'] = 'Свиноматки в/ш'
ost_pt['A16'] = '2 б/ш Староминка'
ost_pt['A14'].font = Font(size=11)
ost_pt['A15'].font = Font(size=11)
ost_pt['A16'].font = Font(size=11)
for j in range(14, 16):
	format(ost_pt, 'A', i, 18, 36, 11, True)
format(ost_pt, 'A', 32, 23, 65, 16, True)
oform(ost_pt, 'A', 32, 'thick')
format(ost_pt, 'A', 32, 23, 65, 16, True)
ost_pt['A35'] = 'Браки'
ost_pt['B35'] = 'Кол-во п/т / четвертин'
ost_pt['A35'].fill = PatternFill('solid', start_color='7AF38A')
ost_pt['B35'].fill = PatternFill('solid', start_color='7AF38A')
oform(ost_pt, 'A', 35, 'thick')
format(ost_pt, 'A', 35, 23, 65, 16, True)
oform(ost_pt, 'B', 35, 'thick')
format(ost_pt, 'B', 35, 23, 65, 16, True)

a = 36
for i in range(1, birn.max_row):
	if re.search('без баков', str(birn['B' + str(i)].value)) and not re.search(
	  'виномат|Староми', str(birn['B' + str(i)].value)):

		ost_pt['A' + str(a)] = birn['B' + str(i)].value
		ost_pt['B' + str(a)] = birn['J' + str(i)].value

		a += 1
for i in range(36, a):
	for j in range(1, 2):

		format(ost_pt, get_column_letter(j), i, 34, 36, 8, False)
		oform(ost_pt, get_column_letter(j), i, 'thin')
ost_pt['A' + str(a)] = 'Итого'
ost_pt['A' + str(a)].fill = PatternFill('solid', start_color='BCBCBC')
ost_pt['B' + str(a)].fill = PatternFill('solid', start_color='BCBCBC')

ost_pt['B' + str(a)] = sum([
 int(ost_pt['B' + str(i)].value) for i in range(36, a)
 if ost_pt['B' + str(i)].value != None
])
format(ost_pt, 'A', a, 23, 65, 16, True)
oform(ost_pt, 'A', a, 'thick')

a += 1
b = copy.copy(a)

for i in range(3, birn.max_row):
	if re.search('задняя четвертина|передняя четвертина|ФС',
	             str(birn['B' + str(i)].value)):

		ost_pt['A' + str(a)] = birn['B' + str(i)].value
		ost_pt['B' + str(a)] = birn['J' + str(i)].value
		a += 1
					 
for i in range(2,ost_pt.max_row):
	try:
		ost_pt['B' + str(i)].value * 1
	except:
		ost_pt['B' + str(i)].value = 0
for i in range(11 ,birn.max_row):
	try:
		birn['J' + str(i)].value * 1
	except:
		birn['J' + str(i)].value = 0
ost_pt['B' + str(a)] = 	sum([ost_pt['B' + str(i)].value for i in range(b, a) ])	


ost_pt['A' + str(a)] = 'Итого'
ost_pt['A' + str(a)].fill = PatternFill('solid', start_color='BCBCBC')
ost_pt['B' + str(a)].fill = PatternFill('solid', start_color='BCBCBC')
format(ost_pt, 'A', a, 23, 65, 16, True)
oform(ost_pt, 'A', a, 'thick')
for i in range(36, a + 1):
	oform(ost_pt, 'B', i, 'thin')
	ost_pt['B' + str(i)].font = Font(size=11, bold=True)
for i in range(1, birn.max_row):
	if re.search('Свинина 4 кат. Свиноматки (ПП) охл. без шкуры',
	             str(birn['B' + str(i)].value)):
		ost_pt['B14'] = birn['J' + str(i)].value
	elif re.search('свинина в полутуш.4 кат.в шкуре (свиноматки)',
	               str(birn['B' + str(i)].value)):
		ost_pt['B15'] = birn['J' + str(i)].value
	elif re.search('Свинина в тушах и полутушах охл.2 кат.в шкуре',
	               str(birn['B' + str(i)].value)):
		ost_pt['B16'] = birn['J' + str(i)].value
nomer = 1
for i in range(1, birn.max_row):
	if re.search('Холодильник хранения туш', str(birn['B' + str(i)].value)):
		nomer = i


for i in range(11, nomer):

	t = '0'
	
	if birn['B' + str(i)].value != None: 
		
		t = birn['B' + str(i)].value
	for j in range(3, 13):
		if ost_pt['A' + str(j)].value != None:
			s = ost_pt['A' + str(j)].value
		if   s.replace('\t', '').strip() == t.replace(
		  '\t', '').strip():
			ost_pt['D' + str(j)] = int(birn['J' + str(i)].value) / 2		  
	if re.search('ВК 1',t):
		a = np.nan if ost_pt['D29'].value is None else ost_pt['D29'].value  
		ost_pt['D29'] =np.nansum([ a, int(birn['J' + str(i)].value) / 2])	
	if re.search('ВК 2',t):
		a = np.nan if ost_pt['D30'].value is None else ost_pt['D30'].value  
		ost_pt['D30'] =np.nansum([ a, int(birn['J' + str(i)].value) / 2])	
	if re.search('ВК Тощ',t):
		a = np.nan if ost_pt['D31'].value is None else ost_pt['D31'].value  
		ost_pt['D31'] =np.nansum([ a, int(birn['J' + str(i)].value) / 2])
	if re.search('МБК',t):
		a = np.nan if ost_pt['D19'].value is None else ost_pt['D19'].value  
		ost_pt['D19'] =np.nansum([ a, int(birn['J' + str(i)].value) / 2])	
	
	if re.search('МБ .* (Суп|Экстр|Прима)',t):

		a = np.nan if ost_pt['D23'].value is None else ost_pt['D23'].value  
		ost_pt['D23'] =np.nansum([ a, int(birn['J' + str(i)].value) / 2])	
if re.search('МТ .* (Суп|Экстр|Прима)',t):

	a = np.nan if ost_pt['D25'].value is None else ost_pt['D25'].value  
	ost_pt['D25'] =np.nansum([ a, int(birn['J' + str(i)].value) / 2])
if re.search('МТ .* (Хорош|Отлич)',t):

	a = np.nan if ost_pt['D26'].value is None else ost_pt['D26'].value  
	ost_pt['D26'] =np.nansum([ a, int(birn['J' + str(i)].value) / 2])	
if re.search('МБ .* (Хорош|Отлич)',t):

	a = np.nan if ost_pt['D24'].value is None else ost_pt['D24'].value  
	ost_pt['D24'] =np.nansum([ a, int(birn['J' + str(i)].value) / 2])	
if re.search('МТ .* (Удовл|Низка)',t):

	a = np.nan if ost_pt['D28'].value is None else ost_pt['D28'].value  
	ost_pt['D28'] =np.nansum([ a, int(birn['J' + str(i)].value) / 2])	
if re.search('МБ .* (Удовл|Низка)',t):

	a = np.nan if ost_pt['D27'].value is None else ost_pt['D27'].value  
	ost_pt['D27'] =np.nansum([ a, int(birn['J' + str(i)].value) / 2])	
if re.search('ВСК от МБК',t):

	a = np.nan if ost_pt['D21'].value is None else ost_pt['D21'].value  
	ost_pt['D21'] =np.nansum([ a, int(birn['J' + str(i)].value) / 2])	
if re.search('ВСК от МБ\b',t):

	a = np.nan if ost_pt['D20'].value is None else ost_pt['D20'].value  
	ost_pt['D20'] =np.nansum([ a, int(birn['J' + str(i)].value) / 2])	
if re.search('ВСК от МТ',t):

	a = np.nan if ost_pt['D22'].value is None else ost_pt['D22'].value  
	ost_pt['D22'] =np.nansum([ a, int(birn['J' + str(i)].value) / 2])	

kat1 = 0
for i in range(nomer , birn.max_row + 1 ):
	if re.search('беконная', str(birn['B' + str(i)].value)):
		kat1 = i

		
	t = '0'
	if birn['B' + str(i)].value != None: 
		t = birn['B' + str(i)].value
	for j in range(3, 13):
		if ost_pt['A' + str(j)].value != None:
			s = ost_pt['A' + str(j)].value
		if   s.replace('\t', '').strip() == t.replace(
		  '\t', '').strip():
			ost_pt['B' + str(j)] = int(birn['J' + str(i)].value) / 2

ost_pt['B7'] = sum((int(ost_pt['B' + str(i)].value) for i in range(3,7) ))
ost_pt['B10'] = sum((int(ost_pt['B' + str(i)].value) for i in range(8,10) ))			  
ost_pt['B13'] = sum((int(ost_pt['B' + str(i)].value) for i in [7,10,11,12] ))
s =[]
if kat1 != '0':
	q = birn['B' + str(kat1 + 1)].value
	while q ==None or  re.search('\d{2}.\d{2}.\d{4}', q ):
		if birn['J' + str(kat1 + 1)].value != 0:
			
			s.append([birn['B' + str(kat1 + 1)].value,birn['J' + str(kat1 + 1)].value / 2] )
		kat1 += 1
		q = birn['B' + str(kat1 + 1)].value
	
for i in range(len(s)):
	if s[i][0] != None:
		dt = datetime.datetime.strptime(s[i][0], '%d.%m.%Y %H:%M:%S')
		dt -= datetime.timedelta(days=1)
		s[i][0] = dt.strftime('%d.%m.%Y')

s = str([f'{s[i][0]} - {s[i][1]}' for i in range(len(s))   ]	)	
ost_pt['C11'] = s
for i in range(3, 13):
	ost_pt['A' + str(i)] = list(sv.values())[i - 3]
ost_pt['A6'] = '2 кат дюрки'

#ost_pt['B6'] = int(input('Количество ТУШ 2 категории дюрков?')) # Графика

#ost_pt['B3'] = ost_pt['B3'].value - ost_pt['B6'].value
for i in range(nomer,birn.max_row):
	q = birn['B' + str(i)].value
	if q != None and re.search('\d{2}.\d{2}.\d{4}', q): 
		dt = datetime.datetime.strptime(q, '%d.%m.%Y %H:%M:%S' )
		try:
			dt = dt - datetime.timedelta(days=1)
		except OverflowError:	
			birn['B' + str(i)] = '0'
			
		birn['B' + str(i)] = dt.strftime('%d.%m.%Y')
	
	
		
		


vk = {}

for i in range(nomer,birn.max_row):
	vkDate = []
	vkCount = []
	q = str(birn['B' + str(i)].value)
	if re.search('ВК .* в полутушах', q):
		ind = i + 1
#		vk[q] = None
		while  birn['B' + str(ind)].value == None or  re.search('\d{2}.\d{2}.\d{4}', str(birn['B' + str(ind)].value) ):
			if birn['J' + str(ind)].value != 0:
				vkDate.append(birn['B' + str(ind)].value)
				vkCount.append(int(birn['J' + str(ind)].value) / 2)
			ind += 1
		vk[q] = pd.Series(vkCount, index = vkDate)
vkd = pd.DataFrame(vk)







#print(vkd)


#ost.save('йцу.xlsx')
#with pd.ExcelWriter("йцу.xlsx", mode="a", engine="openpyxl") as writer:
     

#    vkd.to_excel(writer, sheet_name="Cool drinks")
ost.save(f'{name}.xlsx')