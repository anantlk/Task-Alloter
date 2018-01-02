def assign(free_slot,schd,place,candidates):
	for i in free_slot:
		candidates[i]=[]
		for name in free_slot[i]:
			if(i>1 and i<10):
				if(schd[name][i+1]==place or schd[name][i-1]==place):
					candidates[i].append(name)
			elif(i==1):
				if(schd[name][i+1]==place):
					candidates[i].append(name)
			else:
				if(schd[name][i-1]==place):
					candidates[i].append(name)
	return candidates
