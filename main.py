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
ost_pt['A6'] = '2 кат дюрки'
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
try:
	sum([ost_pt['B' + str(i)].value for i in range(b, a) if [ost_pt['B' + str(i)].value != None  ]])
except TypeError:
	ost_pt['B' + str(a)] = 0

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
ttt = []

#for i in range(11, nomer):
#	t = birn['B' + str(i)].value
#	for j in range(1, 18):
#		s = ost_pt['A' + str(j)].value
#		if s != None and  birn['J' + str(i)].value != None and s.replace('\t', '').strip() == t.replace(
#		  '\t', '').strip():
#			ost_pt['D' + str(j)] = int(birn['J' + str(i)].value) / 2
#			ttt.append(birn['B' + str(i)].value)

print(ttt)
print(i)
ost.save(f'{name}.xlsx')
