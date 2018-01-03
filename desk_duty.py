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


#memebrs who have been alloted 2 duties for a day are bieng removed

def remove_assigned_members(array,count):
	for name in count:
		if(count[name]==2):
			for slot in array:
				if(name in array[slot]):
					array[slot].remove(name)

#alloting duties to members

def allot_duty(i,candidates,venue):
	if(len(candidates)>0):
		candidates=random.sample(candidates,len(candidates))
		for name in candidates:
			if(count[name]==2):
				continue
			count[name]+=1
			i+=1
			table[venue][slot].append(name)
			if(i==2):
				return i,table
	return i,table	
	
			
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

print("Schedule:\n")
print(schd)
print("\n")

#list of members who are free in there slots

for i in range(1,11):
	free_slot[i]=[]
	for name in names:
		if(len(schd[name][i])==0):
			free_slot[i].append(name)

print("Free slots:\n")
print(free_slot)
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
print("TT candidates:\n")
print(tt_candidates),
print("\n")
print("SJT candidates:\n")
print(sjt_candidates)
print("\n")

for slot in tt_candidates:
	i=0
	table['TT'][slot]=[]
	while(i<2):
		candidates=list(set(tt_candidates[slot]))
		i,table=allot_duty(i,candidates,'TT')
		if(i==2):
			break
		candidates=list(set(sjt_candidates[slot]+smv_candidates[slot])-set(table['TT'][slot]))
		i,table=allot_duty(i,candidates,'TT')
		if(i==2):
			break
		candidates=list(set(mb_candidates[slot]+cbmr_candidates[slot])-set(table['TT'][slot]))
		i,table=allot_duty(i,candidates,'TT')
		if(i==2):
			break
		candidates=list(set(cdmm_candidates[slot]+gdn_candidates[slot])-set(table['TT'][slot]))
		i,table=allot_duty(i,candidates,'TT')
		if(i==2):
			break	
		candidates=list(set(hostel_candidates[slot])-set(table['TT'][slot]))
		i,table=allot_duty(i,candidates,'TT')
		if(i==2):
			break
	

for array in [sjt_candidates,tt_candidates,mb_candidates,smv_candidates,cbmr_candidates,cdmm_candidates,gdn_candidates,hostel_candidates]:
	remove_assigned_members(array,count)
	
for slot in tt_candidates:
	i=0
	table['SJT'][slot]=[]
	while(i<2):
		candidates=list(set(sjt_candidates[slot]))
		i,table=allot_duty(i,candidates,'SJT')
		if(i==2):
			break
		candidates=list(set(tt_candidates[slot])-set(table['SJT'][slot]))
		i,table=allot_duty(i,candidates,'SJT')
		if(i==2):
			break
		candidates=list(set(smv_candidates[slot])-set(table['SJT'][slot]))
		i,table=allot_duty(i,candidates,'SJT')
		if(i==2):
			break
		candidates=list(set(cbmr_candidates[slot]+mb_candidates[slot])-set(table['SJT'][slot]))
		i,table=allot_duty(i,candidates,'SJT')
		if(i==2):
			break
		candidates=list(set(cdmm_candidates[slot]+gdn_candidates[slot])-set(table['SJT'][slot]))
		i,table=allot_duty(i,candidates,'SJT')
		if(i==2):
			break	
		candidates=list(set(hostel_candidates[slot])-set(table['SJT'][slot]))
		i,table=allot_duty(i,candidates,'SJT')
		if(i==2):
			break

print("Final Table:\n")
print(table)  #print final table
print("\n")

print("Number of duties alloted:\n")
for name in count:
	print(name),
	print(" "),
	print(count[name])  #print count of duty allocation
	print("\n")  

#generate excel sheet
 
print_excel_sheet.print_excel(table)

print("\n\nDesk Duty Generated!! Check Your directory")


