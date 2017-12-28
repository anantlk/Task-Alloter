import xlrd
import random
from random import choice,sample
from os.path import join,dirname,abspath
schd={"mon":{}}
fname=join(dirname(abspath(__file__)),'time_table.xlsx')

workBook=xlrd.open_workbook(fname)

workSheet=workBook.sheet_by_name("Sheet1")
numCol=workSheet.ncols
numRow=workSheet.nrows
print(numCol)
print(numRow)
for i in range(1,numCol):
	schd["mon"][i]=[]
for i in range(2,numRow):
	#print("row:%s"%i)
	for j,item in enumerate(workSheet.row(i)):
		if(j==0):
			name=str(item.value)
		if(j!=0 and len(item.value)==0):
			schd["mon"][j].append(name)
print(schd)
table={"sjt":{},"tt":{}}
for day in schd:
	for slot in schd[day]:
		per1=random.sample(schd[day][slot],2)
		per2=random.sample(list(set(schd[day][slot])-set(per1)),2)
		table["sjt"][slot]=per1
		table["tt"][slot]=per2
print("\n\n\n")
print(table)
