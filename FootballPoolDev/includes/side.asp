<%	'**************************************************************************
	'* Common code for the survivor/margin side pool.                         *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Removes weekly margin pool scores stored in the database for the week
	' specified.
	'
	' Note: This should be called anytime a regular season game is updated.
	' I.e., any change to the Schedule table.
	'--------------------------------------------------------------------------
	sub ClearMarginResultsCache(week)

		dim sql

		sql = "UPDATE SidePicks" _
		   & " SET MarginScore = NULL" _
		   & " WHERE Week >= " & week
		call DbConn.Execute(sql)

	end sub

	'--------------------------------------------------------------------------
	' Removes weekly survivor pool status information stored in the database for
	' the week specified and beyond.
	'
	' Note: This should be called anytime a change is made to a survivor/margin
	' pool entry or any time a regular season game is updated. I.e., any change
	' the SidePicks or Schedule tables.
	'--------------------------------------------------------------------------
	sub ClearSurvivorStatus(week)

		dim sql

		sql = "DELETE FROM SurvivorStatus WHERE Week >= " & week
		call DbConn.Execute(sql)

	end sub

	'--------------------------------------------------------------------------
	' Builds a drop down list of teams the given user can use as a
	' survivor/margin pick for the given week.
	'--------------------------------------------------------------------------
	sub DisplaySidePoolPickList(ntabs, username, week, selectedPick)

		dim str, used
		dim sql, rs
		dim valid
		dim allLocked, gameDate
		dim i, team

		str = String(nTabs, vbTab) & "<option value=""""></option>" & vbCrLf

		'Get all teams already used by this player.
		used = UsedTeamsList(username)

		'Get all the teams playing in the given week.
		allLocked = AllGamesLocked(week)
		sql = "SELECT Teams.*, Schedule.Date, Schedule.Time FROM Teams, Schedule" _
		  & " WHERE Schedule.Week = " & week _
		  & " AND (TeamID = VisitorID OR TeamID = HomeID)" _
		  & " ORDER BY City, Name"
		set rs = DbConn.Execute(sql)
		do while not rs.EOF
			valid = true

			'If the team's game is locked, exclude it (except for the
			'Administrator).
			gameDate = CDate(rs.Fields("Date").Value & " " & rs.Fields("Time").Value)
			if not IsAdmin() and (allLocked or GameStarted(gameDate)) then
				valid = false
			elseif IsArray(used) then

				'If the team has been used already, exclude it (except if the
				'team is the user's current survivor pick).
				for i = 0 to UBound(used)
					if used(i) = rs.Fields("TeamID").Value and used(i) <> selectedPick then
						valid = false
						exit for
					end if
				next
			end if
			if valid then
				if rs.Fields("DisplayName").Value <> "" then
					team = rs.Fields("DisplayName").Value
				else
					team = rs.Fields("City").Value
				end if
				str = str _
				    & String(nTabs, vbTab) _
				    & "<option value=""" & rs.Fields("TeamID").Value & """"
				if rs.Fields("TeamID").Value = selectedPick then
					str = str & " selected=""selected"""
				end if
				str = str & ">" & team & "</option>" & vbCrLf
			end if
			rs.MoveNext
		loop
		Response.Write(str)

	end sub

	'--------------------------------------------------------------------------
	' Format the given side pool pick.
	'--------------------------------------------------------------------------
	function FormatSidePoolPick(pick, week)

		dim sql, rs, team

		FormatSidePoolPick = pick

		'Get the team name.
		team = pick
		if pick <> "" then
			sql = "SELECT City, DisplayName FROM Teams WHERE TeamID = '" & pick & "'"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				if rs.Fields("DisplayName").Value <> "" then
					team = rs.Fields("DisplayName").Value
				else
					team = rs.Fields("City").Value
				end if
				FormatSidePoolPick = team
			end if
		else
			FormatSidePoolPick = "---"
			exit function
		end if

		sql = "SELECT Result FROM Schedule" _
		   & " WHERE Week = " & week _
		   & " AND (VisitorID = '" & pick & "' OR HomeID = '" & pick & "')" _
		   & " AND NOT ISNULL(Result)"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			if  pick = rs.Fields("Result").Value or _
			   (not SURVIVOR_STRIKE_ON_TIE and rs.Fields("Result").Value = TIE_STR) then
				FormatSidePoolPick = FormatCorrectPick(team)
			end if
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the given users survivor/margin pick for the given week. If the
	' user has not made a pick, an empty string is returned.
	'--------------------------------------------------------------------------
	function GetSidePoolPick(username, week)

		dim sql, rs

		GetSidePoolPick = ""
		sql = "SELECT Pick FROM SidePicks" _
		   & " WHERE Username = '" & SqlString(username) & "'" _
		   & " AND Week = " & week
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			GetSidePoolPick = rs.Fields("Pick").Value
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns a record set containing information about given user's status in
	' the survivor pool for the given week. If no data is found, an empty
	' record set is returned.
	'--------------------------------------------------------------------------
	function GetSurvivorStatus(username, week)

		dim sql

		'Search for the record..
		sql = "SELECT * FROM SurvivorStatus" _
		   & " WHERE Week = " & week _
		   & " AND Username = '" & SqlString(username) & "'"
		set GetSurvivorStatus = DbConn.Execute(sql)
		if not (GetSurvivorStatus.EOF and GetSurvivorStatus.BOF) then
			exit function
		else

			'No record found, try building the status information and search
			'again.
			call SetSurvivorStatus(week)
			set GetSurvivorStatus = DbConn.Execute(sql)
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the name of the given team.
	'--------------------------------------------------------------------------
	function GetTeamName(id)

		dim sql, rs

		GetTeamName = ""
		sql = "SELECT City, DisplayName FROM Teams WHERE TeamID = '" & id & "'"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			if rs.Fields("DisplayName").Value <> "" then
				GetTeamName = rs.Fields("DisplayName").Value
			else
				GetTeamName = rs.Fields("City").Value
			end if
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns true if the given user has entered any picks in the
	' survivor/margin pool.
	'--------------------------------------------------------------------------
	function InSidePool(username)

		dim sql, rs

		InSidePool = false
		sql = "SELECT COUNT(*) AS Total FROM SidePicks" _
		   & " WHERE Username = '" & SqlString(username) & "'"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			if rs.Fields("Total").Value > 0 then
				InSidePool = true
			end if
		end if

	end function

	'--------------------------------------------------------------------------
	' Determines the winner of the margin pool. An array of those player
	' names is returned. If the pool has not been concluded, or if no players
	' participated in it, and empty string is returned instead.
	'--------------------------------------------------------------------------
	function MarginWinnersList()

		dim numWeeks
		dim users, i, week
		dim highScore, mostCorrect
		dim score, correct, n
		dim sql, rs, list

		MarginWinnersList = ""

		'If not all the regular season games are over, exit.
		numWeeks = NumberOfWeeks()
		if CurrentWeek() < numWeeks then
			exit function
		end if
		if NumberOfCompletedGames(numWeeks) <> NumberOfGames(numWeeks) then
			exit function
		end if

		'If there were no players participating in the pool, exit.
		users = SidePoolPlayersList()
		if not IsArray(users) then
			exit function
		end if

		'Check each player.
		for i = 0 to UBound(users)

			'Get the player's total score.
			score = 0
			for week = SIDE_START_WEEK to numWeeks
				n = PlayerMarginScore(users(i), week)
				if IsNumeric(n) then
					score = score + n
				end if
			next

			'Get the number of correct picks the player had.
			correct = 0
			sql = "SELECT COUNT(*)AS Total FROM Schedule, SidePicks" _
			   & " WHERE Schedule.Week = SidePicks.Week" _
			   & " AND Username = '" & SqlString(users(i)) & "'" _
			   & " AND Pick = Result"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				correct = rs.Fields("Total").Value
			end if

			'If this is the first player, initialize the comparison values.
			if i = 0 then
				highScore   = score - 1
				mostCorrect = correct = 1
			end if

			'If this player has a higher score, make the player the winner.
			if score > highScore or (score = highScore and correct > mostCorrect) then
				redim list(0)
				list(0)     = users(i)
				highScore   = score
				mostCorrect = correct

			'Otherwise, if this player has the same score and the same number
			'of correct picks, add the player to the winners list.
			elseif score = highScore and correct = mostCorrect then
				redim preserve list(UBound(list) + 1)
				list(UBound(list)) = users(i)
			end if

		next

		if IsArray(list) then
			MarginWinnersList = list
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the number of players who have entries in the survivor/margin
	' pool.
	'--------------------------------------------------------------------------
	function NumberOfSideEntries()

		dim sql, rs

		NumberOfSideEntries = 0
		sql = "SELECT COUNT(*) AS Total FROM (SELECT DISTINCT Username FROM SidePicks)"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			NumberOfSideEntries = rs.Fields("Total").Value
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the given user's margin score for the given week.
	'--------------------------------------------------------------------------
	function PlayerMarginScore(username, week)

		dim sql, rs
		dim pick, margin

		PlayerMarginScore = ""

		'Exit if the week is invalid.
		if week < SIDE_START_WEEK then
			exit function
		end if

		'See if we have the score already calculated and saved in the database.
		sql = "SELECT MarginScore FROM SidePicks" _
		    & " WHERE Username = '" & SqlString(username) & "'" _
		    & " AND Week = " & week _
			& " AND NOT ISNULL(MarginScore)"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			PlayerMarginScore = rs.Fields("MarginScore").Value
			exit function
		end if

		'Get the user's pick.
		pick = GetSidePoolPick(username, week)
		if pick <> "" then

			'Find the margin of victory/loss.
			sql = "SELECT VisitorID, VisitorScore, HomeID, HomeScore, Result" _
			   & " FROM Schedule" _
			   & " WHERE Schedule.Week = " & week _
			   & " AND (VisitorID = '" & pick & "' OR HomeID = '" & pick & "')" _
			   & " AND NOT ISNULL(Result)"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				margin = Abs(rs.Fields("VisitorScore").Value - rs.Fields("HomeScore").Value)
				if rs.Fields("Result").Value = pick then
					PlayerMarginScore = margin
				else
					if MARGIN_DEDUCT_LOSS then
						PlayerMarginScore = -margin
					else
						PlayerMarginScore = 0
					end if
				end if

				'To improve performance, store the score in the database.
				sql = "UPDATE SidePicks SET" _
				    & " MarginScore = " & PlayerMarginScore _
				    & " WHERE Username = '" & SqlString(username) & "'" _
				    & " AND Week = " & week
				call DbConn.Execute(sql)
			end if
		else

			'The user does not have a pick for that week. If all the week's
			'games have been completed then assign a score.
			if NumberOfCompletedGames(week) = NumberOfGames(week) then

				if not MARGIN_DEDUCT_LOSS then
					PlayerMarginScore = 0
				else

					'Find the greatest score margin for that week.
					sql = "SELECT MAX(ABS(VisitorScore - HomeScore)) AS Margin FROM Schedule" _
					   & " WHERE Week = " & week _
					   & " AND NOT ISNULL(Result)"
					set rs = DbConn.Execute(sql)
					if not (rs.BOF and rs.EOF) then
						margin = rs.Fields("Margin").Value
						PlayerMarginScore = -margin
					end if
				end if
			end if
		end if

	end function

	'--------------------------------------------------------------------------
	' Determines the outcome for each player in the survivor pool for the given
	' week and saves that information to the database.
	'--------------------------------------------------------------------------
	sub SetSurvivorStatus(week)

		dim sql, rs
		dim users, gamesRs, weekCompleted
		dim wasAlive, completedGames, missed, revived, isAlive, isWinner
		dim pick, result
		dim i, j

		'Exit if the week is invalid.
		if week < SIDE_START_WEEK then
			exit sub
		end if

		'Clear any existing status information.
		ClearSurvivorStatus(week)

		'If we have data from the previous week and there were one or more
		'winners, just copy the data to this week and exit.
		if week > SIDE_START_WEEK then
			sql = "SELECT * FROM SurvivorStatus" _
			   & " WHERE Week = " & week - 1 _
			   & " AND IsWinner"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				sql = "SELECT * FROM SurvivorStatus WHERE Week = " & week - 1
				set rs = DbConn.Execute(sql)
				do while not rs.EOF
					sql = "INSERT INTO SurvivorStatus" _
					   & "(Week, Username, WasAlive, CompletedGames, Missed, Revived, IsAlive, IsWinner)" _
					   & "VALUES(" _
					   & week & ", " _
					   & "'" & SqlString(rs.Fields("Username").Value) & "', " _
					   & rs.Fields("WasAlive").Value & ", " _
					   & rs.Fields("CompletedGames").Value & ", " _
					   & rs.Fields("Missed").Value & ", " _
					   & rs.Fields("Revived").Value & ", " _
					   & rs.Fields("IsAlive").Value & ", " _
					   & rs.Fields("IsWinner").Value & ")"
					call DbConn.Execute(sql)
					rs.MoveNext
				loop
				exit sub
			end if
		end if

		'Get a list of users participating in the survivor pool.
		users = SidePoolPlayersList()
		if not IsArray(users) then
			exit sub
		end if

		'Get all games for the week.
		set gamesRs = WeeklySchedule(week)

		'Determine if all games for the week have been concluded.
		weekCompleted = false
		if NumberOfCompletedGames(week) = NumberOfGames(week) then
			weekCompleted = true
		end if

		'For each user, determine the result of the week's pick.
		for i = 0 to UBound(users)

			'If this is the first week, initialize the results data. Otherwise,
			'load the previous week's data.
			if week = SIDE_START_WEEK then
				wasAlive       = true
				completedGames = 0
				missed         = 0
				revived        = 0
				isWinner       = false
			else
				sql = "SELECT * FROM SurvivorStatus" _
				   & " WHERE Week = " & week - 1 _
				   & " AND Username = '" & SqlString(users(i)) & "'"
				set rs = DbConn.Execute(sql)
				if not (rs.BOF and rs.EOF) then
					wasAlive       = rs.Fields("IsAlive").Value
					completedGames = rs.Fields("CompletedGames").Value
					missed         = rs.Fields("Missed").Value
					revived        = rs.Fields("Revived").Value
					isAlive        = rs.Fields("IsAlive"). Value
					isWinner       = rs.Fields("IsWinner").Value
				else
					'The previous week's results are not in the database so we
					'need to recalculate them and start processing over for
					'this week.
					call SetSurvivorStatus(week - 1)
					call SetSurvivorStatus(week)
					exit sub
				end if
			end if

			'Start with the previous week's status.
			isAlive = wasAlive

			'If the user survived the previous week, check this week's result.
			if wasAlive then

				'Get the user's pick for the current week.
				pick = GetSidePoolPick(users(i), week)

				'Get the result of that pick, if possible.
				if pick = "" and weekCompleted then
					completedGames = completedGames + 1
					missed = missed + 1
				else
					gamesRs.MoveFirst
					do while not gamesRs.EOF
						if pick = gamesRs.Fields("VisitorID").Value or pick = gamesRs.Fields("HomeID").Value then
							result = gamesRs.Fields("Result").Value
							if not IsNull(result) then
								completedGames = completedGames + 1
								if pick <> result and not (result = TIE_STR and not SURVIVOR_STRIKE_ON_TIE) then
									missed = missed + 1
								end if
							end if
							exit do
						end if
						gamesRs.MoveNext
					loop
				end if

				'Determine if the user has been eliminated.
				if missed - revived >= SURVIVOR_STRIKE_OUT then
					isAlive = false
				end if
			end if

			'Save the results for this user.
			sql = "INSERT INTO SurvivorStatus" _
			   & "(Week, Username, WasAlive, CompletedGames, Missed, Revived, IsAlive, IsWinner)" _
			   & "VALUES(" _
			   & week & ", " _
			   & "'" & SqlString(users(i)) & "', " _
			   & wasAlive & ", " _
			   & completedGames & ", " _
			   & missed & ", " _
			   & revived & ", " _
			   & isAlive & ", " _
			   & isWinner & ")"
			call DbConn.Execute(sql)
		next

		'If the week is completed and all previously active players have been
		'eliminated, revive them.
		if weekCompleted then
			sql = "SELECT COUNT(*) AS Total From SurvivorStatus" _
			   & " WHERE Week = " & week _
			   & " AND IsAlive"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				if rs.Fields("Total").Value = 0 then
					sql = "UPDATE SurvivorStatus" _
					   & " SET IsAlive = True," _
					   & " Revived = Revived + 1" _
					   & " WHERE Week = " & week _
					   & " AND WasAlive"
					call DbConn.Execute(sql)
				end if
			end if
		end if

		'Determine if we have a winner.
		if weekCompleted then

			'Find out how many players survived.
			sql = "SELECT COUNT(*) AS Total From SurvivorStatus" _
			   & " WHERE Week = " & week _
			   & " AND IsAlive"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				if rs.Fields("Total").Value = 1 then

					'Only one player survived, make that player the winner.
					sql = "UPDATE SurvivorStatus" _
					   & " SET IsWinner = true" _
					   & " WHERE Week = " & week _
					   & " AND IsAlive"
					call DbConn.Execute(sql)
					exit sub
				elseif week = NumberOfWeeks() then

					'All weeks completed, make the surviving player(s) with the
					'fewest misses the winner(s).
					sql = "UPDATE SurvivorStatus" _
					   & " SET IsWinner = TRUE" _
					   & " WHERE Week = " & week _
					   & " AND IsAlive" _
					   & " AND Missed = (" _
					   & "   SELECT MIN(Missed) FROM SurvivorStatus" _
					   & "   WHERE Week = " & week _
					   & "   AND IsAlive)"
					call DbConn.Execute(sql)
				end if
			end if
		end if

	end sub

	'--------------------------------------------------------------------------
	' Returns true if the game for the given survivor/margin pick has been
	' locked.
	'--------------------------------------------------------------------------
	function SidePoolPickLocked(pick, week)

		dim sql, rs, gameDate

		SidePoolPickLocked = false
		if AllGamesLocked(week) then
			SidePoolPickLocked = true
			exit function
		end if

		sql = "SELECT [Date], [Time] FROM Schedule" _
		   & " WHERE Week = " & week _
		   & " AND (VisitorID = '" & pick & "' OR HomeID = '" & pick & "')"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			gameDate = CDate(rs.Fields("Date").Value & " " & rs.Fields("Time").Value)
			if GameStarted(gameDate) then
				SidePoolPickLocked = true
			end if
		end if

	end function

	'--------------------------------------------------------------------------
	' Determines the winner of the survivor pool. An array of those player
	' names is returned. If the pool has not been concluded, or if no players
	' participated in it, and empty string is returned instead.
	'--------------------------------------------------------------------------
	function SurvivorWinnersList()

		dim sql, rs, week, list

		SurvivorWinnersList = ""

		'If the pool has not been won yet, exit.
		week = SurvivorFinalWeek()
		if not IsNumeric(week) then
			exit function
		end if

		'Check the status.
		sql = "SELECT Username FROM SurvivorStatus" _
		   & " WHERE Week = " & week _
		   & " AND IsWinner"
		set rs = DbConn.Execute(sql)
		do while not rs.EOF
			if not IsArray(list) then
				redim list(0)
			else
				redim preserve list(UBound(list) + 1)
			end if
			list(UBound(list)) = rs.Fields("Username").Value
			rs.MoveNext
		loop
		if IsArray(list) then
			SurvivorWinnersList = list
			exit function
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the week that the survivor pool was won. If that pool has not
	' been decided yet, an empty string is returned.
	'--------------------------------------------------------------------------
	function SurvivorFinalWeek()

		dim curWeek, week
		dim sql, rs

		SurvivorFinalWeek = ""

		'Make sure we have the most current results.
		curWeek = CurrentWeek()
		sql = "SELECT COUNT(*) AS Total FROM SurvivorStatus" _
		   & " WHERE Week = " & curWeek
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			if rs.Fields("Total").Value < 1 then
				SetSurvivorStatus(curWeek)
			end if
		end if

		'Check for a winner.
		sql = "SELECT Min(Week) AS FinalWeek FROM SurvivorStatus WHERE IsWinner"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			week = rs.Fields("FinalWeek").Value
			if IsNumeric(week) then
				SurvivorFinalWeek = week
			end if
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns an array of users currently in the survivor/margin pool.
	'--------------------------------------------------------------------------
	function SidePoolPlayersList()

		dim list, sql, rs

		SidePoolPlayersList = ""

		sql = "SELECT DISTINCT Username FROM SidePicks" _
		   & " ORDER BY Username"
		set rs = DbConn.Execute(sql)
		do while not rs.EOF
			if not IsArray(list) then
				redim list(0)
			else
				redim preserve list(UBound(list) + 1)
			end if
			list(UBound(list)) = rs.Fields("Username").Value
			rs.MoveNext
		loop
		SidePoolPlayersList = list

	end function

	'--------------------------------------------------------------------------
	' Returns an array of teams already used as a survivor/margin pick by the
	' given user.
	'--------------------------------------------------------------------------
	function UsedTeamsList(username)

		dim list, sql, rs

		UsedTeamsList = ""

		sql = "SELECT Pick FROM SidePicks" _
		   & " WHERE Username = '" & SqlString(username) & "'"
		set rs = DbConn.Execute(sql)
		do while not rs.EOF
			if not IsArray(list) then
				redim list(0)
			else
				redim preserve list(UBound(list) + 1)
			end if
			list(UBound(list)) = rs.Fields("Pick").Value
			rs.MoveNext
		loop
		UsedTeamsList = list

	end function %>

