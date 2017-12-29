import xlrd
import xlwt
import random
from random import choice,sample
from os.path import join,dirname,abspath

place=["sjt","tt"]
schd={}
table={"sjt":{},"tt":{}}
count={}
fname=join(dirname(dirname(abspath(__file__))),'time_table.xlsx')
workBook=xlrd.open_workbook(fname)
workSheet=workBook.sheet_by_name("Sheet1")
numCol=workSheet.ncols
numRow=workSheet.nrows

#parsing of given excel sheet
for i in range(1,numCol):
	schd[i]=[]
for i in range(2,numRow):
	for j,item in enumerate(workSheet.row(i)):
		if(j==0):
			name=str(item.value)
			count[name]=0
		if(j!=0 and len(item.value)==0):
			schd[j].append(name)

#assigning duties to students
for slot in schd:
	prev=[]
	table["sjt"][slot]=[]
	table["tt"][slot]=[]
	for venue in place:
		i=0
		while(i<2):
			per=random.choice(list(set(schd[slot])-set(prev)))
			if(count[per]==2):
				continue;
			table[venue][slot].append(per)
			prev.append(per)
			count[per]+=1
			i+=1

#printing schedule and no of slots assigned to every person on terminal
print(table)
print("\n\n\n")
print(count)

book=xlwt.Workbook()
sheet=book.add_sheet("sheet1")
time_slot=[" 8-9am"," 9-10am"," 10-11am","11am-12pm"," 12-1pm"," 2-3pm"," 3-4pm"," 4-5pm"," 5-6pm"," 6-7pm"]

#printing time slots in sxcel
for i in range(len(time_slot)):
	sheet.row(0).write(i+1,time_slot[i])
	
#printing duties in excel	
for i in range(len(place)):	
	sheet.row(2*(i+1)+i).write(0,place[i].upper())
	for slots in table[place[i]]:
			for pos in range(len(table[place[i]][slots])):
				sheet.row(pos+2*(i+1)+i).write(slots,table[place[i]][slots][pos])
			

book.save("desk_duty.xls")
