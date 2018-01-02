import xlwt
place=['SJT','TT']

#printing schedule and no of slots assigned to every person on terminal

def print_excel(table):
	book=xlwt.Workbook()
	sheet=book.add_sheet("sheet1")
	time_slot=[" 8-9am"," 9-10am"," 10-11am","11am-12pm"," 12-1pm"," 2-3pm"," 3-4pm"," 4-5pm"," 5-6pm"," 6-7pm"]

	#printing time slots in excel
	for i in range(len(time_slot)):
		sheet.row(0).write(i+1,time_slot[i])
		
	#printing duties in excel	
	for i in range(len(place)):	
		sheet.row(2*(i+1)+i).write(0,place[i])
		for slots in table[place[i]]:
				for pos in range(len(table[place[i]][slots])):
					sheet.row(pos+2*(i+1)+i).write(slots,table[place[i]][slots][pos])
				

	book.save("desk_duty.xls")
