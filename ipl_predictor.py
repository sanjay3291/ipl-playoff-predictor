import xlsxwriter 

from datetime import datetime

# Workbook is created 
wb = xlsxwriter.Workbook('Probabilities.xlsx')

# add_sheet is used to create sheets. 
sheet1 = wb.add_worksheet('Team_Probabilities')
sheet2 = wb.add_worksheet('Picks_Probabilities')


top_row_format = wb.add_format()
top_row_format.set_center_across()
top_row_format.set_bold()

data_format = wb.add_format()
data_format.set_center_across()

first_cell_format = wb.add_format()
first_cell_format.set_bold()




# Current Points Table
currentPoints = { "mi":10, "dc":10, "rcb":10, "kkr":8, "srh":6, "csk":6, "rr":6, "kxip":2 }

# List of remaining fixtures
gamesLeft = [ "dc:rr", "rcb:kxip", "mi:kkr", "rr:rcb", "dc:csk", "srh:kkr", "mi:kxip", "csk:rr", "kxip:dc", "kkr:rcb", "rr:srh", "csk:mi", "kkr:dc", "kxip:srh", "rcb:csk", "rr:mi", "kkr:kxip", "srh:dc", "mi:rcb", "csk:kkr", "kxip:rr", "dc:mi", "rcb:srh", "csk:kxip", "kkr:rr", "dc:rcb", "srh:mi"]
#gamesLeft = [ "dc:rr", "rcb:kxip", "mi:kkr"]

# sort teams in descending order of points
teams = {k: v for k, v in sorted(currentPoints.items(), key=lambda item: item[1], reverse=True)}


n=len(gamesLeft) # number of games remaining
combinations = pow(2,n) # Assuming no washouts, number of possible combinations remaining.
combinationsStr = "0{}b".format(n) # Each digit represents a game, so create a string of n binary digits

print( "number of matches remaining    = %d" % n )
print( "number of possible combinations = %d" % combinations )

Top2_Confirmed = {}
Top2_onNRR  = {}
Top4_Confirmed = {}
Top4_onNRR  = {}
possibilitiesStr = ""

# initial probability for each team
for team in teams:
	Top2_Confirmed[team] = 0
	Top2_onNRR[team]  = 0
	Top4_Confirmed[team] = 0
	Top4_onNRR[team]  = 0

# run sims
for run in range(0,combinations):

	#print progress of script for every 500 sims
	if run%500 == 0:
		t=str(datetime.now())
		print( "{}: Finsihed {:0,d} sims".format(t,run))

	possibilitiesStr = format(run, combinationsStr)
	simulationPoints = currentPoints.copy()
	gameCount  = 0

	
	# if digit is equal to 0 means team1 wins and digit and 1 means team2 wins
	for digit in possibilitiesStr:
		team1, team2 = gamesLeft[gameCount].split(":")
		if digit == "0":
			winner=team1
		else:
			winner=team2

		# Add 2 points to winning teams
		simulationPoints[ winner ] = simulationPoints[ winner ] + 2
		gameCount = gameCount + 1

	# Determine points table positions for each team
	for team1 in teams:
		aheadTeams  = 0
		behindTeams = 0

		# Determine how many teams and ahead or behind each team
		for team2 in teams:
			if team2 == team1:
				continue
			if simulationPoints[team2] > simulationPoints[team1]:
				aheadTeams = aheadTeams + 1
			if simulationPoints[team2] < simulationPoints[team1]:
				behindTeams = behindTeams + 1

		# Top2_onNRR, if max of 1 team is ahead of a team
		# Top2_Confirmed, if min of 6 teams are below a team
		# Top4_onNRR, if max of 3 team is ahead of a team
		# Top4_Confirmed, if min of 4 teams are below a team
		if aheadTeams <= 1:
			Top2_onNRR[team1]=Top2_onNRR[team1]+1
		if behindTeams >= 6:
			Top2_Confirmed[team1]=Top2_Confirmed[team1]+1	
		if aheadTeams <= 3:
			Top4_onNRR[team1]=Top4_onNRR[team1]+1
		if behindTeams >= 4:
			Top4_Confirmed[team1]=Top4_Confirmed[team1]+1	

# end sims


sheet1.write(0, 0, 'Team', first_cell_format)
sheet1.write(0, 1, 'Top 2 Confirmed', top_row_format)
sheet1.write(0, 2, 'Top 2 Possible on NRR', top_row_format)
sheet1.write(0, 3, 'Top 4 Confirmed', top_row_format)
sheet1.write(0, 4, 'Top 4 Possible on NRR', top_row_format)

sheet1.set_column('A:A', 15)
sheet1.set_column('B:E', 20)



for i,team in enumerate(teams):
	sheet1.write(i+1, 0, team)
	sheet1.write(i+1, 1, str(Top2_Confirmed[team] * 100.0/combinations) + "%", data_format)
	sheet1.write(i+1, 2, str(Top2_onNRR[team] * 100.0/combinations) + "%", data_format)
	sheet1.write(i+1, 3, str(Top4_Confirmed[team] * 100.0/combinations) + "%", data_format)
	sheet1.write(i+1, 4, str(Top4_onNRR[team] * 100.0/combinations) + "%", data_format)





# probabilities based on next game result
for ii in range(0, 2):
	nextGameStr = gamesLeft[0]
	nextGameTeams = nextGameStr.split(':')
	winTeamKey = nextGameTeams[ii]

	newPoints=currentPoints.copy()

	winTeam = dict(filter(lambda item: winTeamKey in item[0], newPoints.items()))

	winTeam[winTeamKey] = winTeam[winTeamKey]+2

	newPoints.update(winTeam)

	new_gamesLeft = gamesLeft.copy()

	new_gamesLeft.pop(0)

	new_n=len(new_gamesLeft) # number of games remaining
	new_combinations = pow(2,new_n) # Assuming no washouts, number of possible combinations remaining.
	new_combinationsStr = "0{}b".format(new_n) # Each digit represents a game, so create a string of n binary digits

	print( "number of matches remaining    = %d" % new_n )
	print( "number of possible combinations = %d" % new_combinations )

	new_Top2_Confirmed = {}
	new_Top2_onNRR  = {}
	new_Top4_Confirmed = {}
	new_Top4_onNRR  = {}
	new_possibilitiesStr = ""

	# initial probability for each team
	for team in teams:
		new_Top2_Confirmed[team] = 0
		new_Top2_onNRR[team]  = 0
		new_Top4_Confirmed[team] = 0
		new_Top4_onNRR[team]  = 0

	# run sims
	for run in range(0,new_combinations):

		#print progress of script for every 500 sims
		if run%500 == 0:
			t=str(datetime.now())
			print( "{}: Finsihed {:0,d} sims".format(t,run))

		new_possibilitiesStr = format(run, new_combinationsStr)
		new_simulationPoints = newPoints.copy()
		new_gameCount  = 0

		
		# if digit is equal to 0 means team1 wins and digit and 1 means team2 wins
		for digit in new_possibilitiesStr:
			team1, team2 = new_gamesLeft[new_gameCount].split(":")
			if digit == "0":
				winner=team1
			else:
				winner=team2

			# Add 2 points to winning teams
			new_simulationPoints[ winner ] = new_simulationPoints[ winner ] + 2
			new_gameCount = new_gameCount + 1

		# Determine points table positions for each team
		for team1 in teams:
			aheadTeams  = 0
			behindTeams = 0

			# Determine how many teams and ahead or behind each team
			for team2 in teams:
				if team2 == team1:
					continue
				if new_simulationPoints[team2] > new_simulationPoints[team1]:
					aheadTeams = aheadTeams + 1
				if new_simulationPoints[team2] < new_simulationPoints[team1]:
					behindTeams = behindTeams + 1

			# Top2_onNRR, if max of 1 team is ahead of a team
			# Top2_Confirmed, if min of 6 teams are below a team
			# Top4_onNRR, if max of 3 team is ahead of a team
			# Top4_Confirmed, if min of 4 teams are below a team
			if aheadTeams <= 1:
				new_Top2_onNRR[team1]=new_Top2_onNRR[team1]+1
			if behindTeams >= 6:
				new_Top2_Confirmed[team1]=new_Top2_Confirmed[team1]+1	
			if aheadTeams <= 3:
				new_Top4_onNRR[team1]=new_Top4_onNRR[team1]+1
			if behindTeams >= 4:
				new_Top4_Confirmed[team1]=new_Top4_Confirmed[team1]+1	

	# end sims

	winTeamStr = "Change if {} wins".format(winTeamKey)

	sheet1.write(10, ii*6, winTeamStr, first_cell_format)


	sheet1.write(11, ii*6, 'Team', first_cell_format)
	sheet1.write(11, ii*6+1, 'Top 2 Confirmed', top_row_format)
	sheet1.write(11, ii*6+2, 'Top 2 Possible on NRR', top_row_format)
	sheet1.write(11, ii*6+3, 'Top 4 Confirmed', top_row_format)
	sheet1.write(11, ii*6+4, 'Top 4 Possible on NRR', top_row_format)

	sheet1.set_column('G:G', 15)
	sheet1.set_column('H:K', 20)

	for i,team in enumerate(teams):
		sheet1.write(i+12, ii*6, team)
		sheet1.write(i+12, ii*6+1, str(new_Top2_Confirmed[team] * 100.0/new_combinations - Top2_Confirmed[team] * 100.0/combinations) + "%", data_format)
		sheet1.write(i+12, ii*6+2, str(new_Top2_onNRR[team] * 100.0/new_combinations - Top2_onNRR[team] * 100.0/combinations) + "%", data_format)
		sheet1.write(i+12, ii*6+3, str(new_Top4_Confirmed[team] * 100.0/new_combinations - Top4_Confirmed[team] * 100.0/combinations) + "%", data_format)
		sheet1.write(i+12, ii*6+4, str(new_Top4_onNRR[team] * 100.0/new_combinations - Top4_onNRR[team] * 100.0/combinations) + "%", data_format)


sheet1.insert_image('G2', 'Addendum.png', {'x_scale': 0.67, 'y_scale': 0.7})

wb.close()
