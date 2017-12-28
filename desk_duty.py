import xlrd
import xlwt
import random
from random import choice,sample
from os.path import join,dirname,abspath

schd={}
table={"sjt":{},"tt":{}}
count={}
fname=join(dirname(dirname(abspath(__file__))),'time_table.xlsx')
workBook=xlrd.open_workbook(fname)
workSheet=workBook.sheet_by_name("Sheet1")
numCol=workSheet.ncols
numRow=workSheet.nrows

def check(per,count):
	for name in per:
		if(count[name]==2):
			return 0
		else:
			count[name]+=1;
	return 1
for i in range(1,numCol):
	schd[i]=[]
for i in range(2,numRow):
	for j,item in enumerate(workSheet.row(i)):
		if(j==0):
			name=str(item.value)
			count[name]=0
		if(j!=0 and len(item.value)==0):
			schd[j].append(name)


for slot in schd:
	per1=random.sample(schd[slot],2)
	per2=random.sample(list(set(schd[slot])-set(per1)),2)
	while(check(per1,count)!=1):
		per1=random.sample(schd[slot],2)
		print("hello")
	table["sjt"][slot]=per1
	while(check(per2,count)!=1):
		per2=random.sample(list(set(schd[slot])-set(per1)),2)
		print("hello")
	table["tt"][slot]=per2



#print(schd)
#print("\n\n\n")
print(table)
print("\n\n\n")
print(count)

book=xlwt.Workbook()
sheet=book.add_sheet("sheet1")
time_slot=[" 8-9am"," 9-10am"," 10-11am"," 11am-12pm"," 12-1pm"," 2-3pm"," 3-4pm"," 4-5pm"," 5-6pm"," 6-7pm"]


for i in range(len(time_slot)):
	sheet.row(0).write(i+1,time_slot[i])
	
#for sjt	
sheet.row(2).write(0,"SJT")
for slots in table["sjt"]:
		for pos in range(len(table["sjt"][slots])):
			sheet.row(pos+2).write(slots,table["sjt"][slots][pos])
			
#for tt

sheet.row(5).write(0,"TT")
for slots in table["tt"]:
		for pos in range(len(table["tt"][slots])):
			sheet.row(pos+5).write(slots,table["tt"][slots][pos])

book.save("desk_duty.xls")
