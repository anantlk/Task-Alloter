import xlrd
import print_excel_sheet
import random
import convenience
from random import choice,sample
from os.path import join,dirname,abspath
names=[]
place=["sjt","tt"]
schd={}
free_slot={}
table={"SJT":{},"TT":{}}
count={}
fname=join(dirname(dirname(abspath(__file__))),'time_table.xlsx')
workBook=xlrd.open_workbook(fname)
workSheet=workBook.sheet_by_name("Sheet1")
numCol=workSheet.ncols
numRow=workSheet.nrows

def remove_assigned_members(array,count):
	for name in count:
		if(count[name]==2):
			for slot in array:
				if(name in array[slot]):
					array[slot].remove(name)

def sort_members(array,count):
	for row in range(len(array)):
		for col in range(len(array)-row-1):
			if(count[array[col+1]]>count[array[col]]):
				t=array[col+1]
				array[col+1]=array[col]
				array[col]=t
	return array
				
def venue_assign(present,array,slot):
	for name in array:
		if name not in present[slot]:
			present[slot].append(name)
		if(len(present[slot])>=6):
			return 1,present
	return 0,present	

def allot_duties(venue,array):
	for slot in array:
		prev=[]
		i=0
		table[venue][slot]=[]
		for per in sort_members(list(set(array[slot])-set(prev)),count):
			if(count[per]==2):
				continue;
			table[venue][slot].append(per)
			prev.append(per)
			count[per]+=1
			i+=1
			if(i==2):
				break
	return table	
			
#parsing of given excel sheet

for i in range(1,numRow):
	for j,item in enumerate(workSheet.row(i)):
		if(j==0):
			name=str(item.value)
			names.append(name)
			count[name]=0
			schd[name]={}
		else:
			schd[name][j]=str(item.value).upper()
#printing the parsed sheet
print(schd),
print("\n")

for i in range(1,11):
	free_slot[i]=[]
	for name in names:
		if(len(schd[name][i])==0):
			free_slot[i].append(name)
#list of members who are free in there slots 
print(free_slot),
print("\n")

#chosing candidates from different places so that they have class either before or after there free slots in that place.This is done to assign places according to there convenience

sjt_candidates=convenience.assign(free_slot,schd,'SJT',{})
tt_candidates=convenience.assign(free_slot,schd,'TT',{})
gdn_candidates=convenience.assign(free_slot,schd,'GDN',{})
mb_candidates=convenience.assign(free_slot,schd,'MB',{})
smv_candidates=convenience.assign(free_slot,schd,'SMV',{})
cdmm_candidates=convenience.assign(free_slot,schd,'CDMM',{})
cbmr_candidates=convenience.assign(free_slot,schd,'CBMR',{})
hostel_candidates=convenience.assign(free_slot,schd,'',{})

#alloting candidates for tt desk duty if not sufficient nummber of member are alloted the duty


for slot in tt_candidates:
	flag=0
	if(len(tt_candidates[slot])<3):
		flag,tt_candidates=venue_assign(tt_candidates,random.sample(sjt_candidates[slot]+smv_candidates[slot],len(sjt_candidates[slot]+smv_candidates[slot])),slot)
		if(flag==1):
			continue
		flag,tt_candidates=venue_assign(tt_candidates,random.sample(cbmr_candidates[slot]+mb_candidates[slot],len(cbmr_candidates[slot]+mb_candidates[slot])),slot)
		if(flag==1):
			continue
		flag,tt_candidates=venue_assign(tt_candidates,random.sample(cdmm_candidates[slot]+gdn_candidates[slot],len(cdmm_candidates[slot]+gdn_candidates[slot])),slot)
		if(flag==1):
			continue
		flag,tt_candidates=venue_assign(tt_candidates,random.sample(hostel_candidates[slot],len(hostel_candidates[slot])),slot)
		if(flag==1):
			continue
						
table=allot_duties('TT',tt_candidates)

#removal of students who has been assigned 2 duties

for array in [sjt_candidates,tt_candidates,smv_candidates,mb_candidates,cbmr_candidates,cdmm_candidates,gdn_candidates,hostel_candidates]:
	remove_assigned_members(array,count)



for slot in sjt_candidates:
	flag=0
	if(len(sjt_candidates[slot])<5):
		flag,sjt_candidates=venue_assign(sjt_candidates,sort_members(list(set(tt_candidates[slot])-set(table['TT'][slot])),count),slot)
		if(flag==1):
			continue
		flag,sjt_candidates=venue_assign(sjt_candidates,sort_members(smv_candidates[slot],count),slot)
		if(flag==1):
			continue
		flag,sjt_candidates=venue_assign(sjt_candidates,sort_members(cbmr_candidates[slot]+mb_candidates[slot],count),slot)
		if(flag==1):
			continue
		flag,sjt_candidates=venue_assign(sjt_candidates,sort_members(cdmm_candidates[slot]+gdn_candidates[slot],count),slot)
		if(flag==1):
			continue
		flag,sjt_candidates=venue_assign(sjt_candidates,sort_members(hostel_candidates[slot],count),slot)
		if(flag==1):
			continue
	
table=allot_duties('SJT',sjt_candidates)

print(table)  #print final table
print("\n")
for name in count:
	print(name),
	print(" "),
	print(count[name])  #print count of duty allocation
	print("\n")  

#generate excel sheet
 
print_excel_sheet.print_excel(table)

print("\n\nDesk Duty Generated!! Check Your directory")




