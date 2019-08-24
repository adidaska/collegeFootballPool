<%	'**************************************************************************
	'* Common code for the weekly pool.                                       *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Removes weekly player scores and winners list stored in the database for
	' the week specified.
	'
	' Note: This should be called anytime a change is made to a weekly pool
	' entry or any time a regular season game is updated. I.e., any change to
	' the Picks, Tiebreaker or Schedule tables.
	'--------------------------------------------------------------------------
	sub ClearWeeklyResultsCache(week)

		dim sql

		'Clear any pick and confidence scores saved in the Tiebreaker table.
		sql = "UPDATE Tiebreaker SET"_
		    & " PickScore = NULL," _
		    & " ConfidenceScore = NULL" _
			& " WHERE Week = " & week
		call DbConn.Execute(sql)

		'Remove the winners from the WeeklyWinners table.
		sql = "DELETE FROM WeeklyWinners WHERE Week = " & week
		call DbConn.Execute(sql)

	end sub

	'--------------------------------------------------------------------------
	' Returns the number of users who made picks for the specified week.
	'--------------------------------------------------------------------------
	function NumberOfEntries(week)

		dim sql, rs

		NumberOfEntries = 0
		sql = "SELECT COUNT(*) As Total FROM Tiebreaker" _
 		    & " WHERE Week = " & week
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			NumberOfEntries = rs.Fields("Total").Value
		end if

	end function

   	'--------------------------------------------------------------------------
	' Returns the actual point total of the last scheduled game of the week
	' specified. If those scores are not available, an empty string is
	' returned.
	'--------------------------------------------------------------------------
	function TBPointTotalVis(week)

		dim sql, rs

		'TBPointTotal = ""
		'set TBPointTotal = new TBclass
		'sql = "SELECT VisitorScore, HomeScore" _
	    '   & " FROM Schedule" _
		'   & " WHERE Week = " & week & " AND TBgame = 1"_
		'   & " ORDER BY [Date] DESC, [Time] DESC"
		'set rs = DbConn.Execute(sql)
		'if not (rs.BOF and rs.EOF) then
		'	if not IsNumeric(rs.Fields("VisitorScore").Value) then
		'		exit function
		'	end if
		'	if not IsNumeric(rs.Fields("HomeScore").Value) then
		'		exit function
		'	end if
		'	TBPointTotal.VisScore = rs.Fields("VisitorScore").Value
		'	TBPointTotal.HomeScore = rs.Fields("HomeScore").Value
		'end if
		
		TBPointTotalVis = ""
		sql = "SELECT VisitorScore" _
	       & " FROM Schedule" _
		   & " WHERE Week = " & week _
		   & " AND TBgame = 1"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			if not IsNumeric(rs.Fields("VisitorScore").Value) then
				exit function
			end if
			TBPointTotalVis = rs.Fields("VisitorScore").Value
		end if

	end function
	
	function TBPointTotalHome(week)

		dim sql, rs
		
		TBPointTotalHome = ""
		sql = "SELECT HomeScore" _
	       & " FROM Schedule" _
		   & " WHERE Week = " & week _
		   & " AND TBgame = 1"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			if not IsNumeric(rs.Fields("HomeScore").Value) then
				exit function
			end if
			TBPointTotalHome = rs.Fields("HomeScore").Value
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the confidence score for picks made by the specified user for the
	' given week. If the user did not enter any picks that week, an empty
	' string is returned.
	'--------------------------------------------------------------------------
	function PlayerConfidenceScore(username, week)

		dim sql, rs, resultField, tieTotal

		PlayerConfidenceScore = ""

		'If the user made no picks for that week, exit.
		if TBvisGuess(username, week) = "" or TBhomeGuess(username, week) = "" then
			exit function
		end if

		'See if we have the score already calculated and saved in the database.
		sql = "SELECT ConfidenceScore FROM Tiebreaker" _
		    & " WHERE Username = '" & SqlString(username) & "'" _
		    & " AND Week = " & week _
			& " AND NOT ISNULL(ConfidenceScore)"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			PlayerConfidenceScore = rs.Fields("ConfidenceScore").Value
			exit function
		end if

		'Determine which result field to use.
		resultField = "Result"
		if USE_POINT_SPREADS then
			resultField = "ATSResult"
		end if

		'Total the confidence points for each correct pick.
		sql = "SELECT SUM(Confidence) AS Total" _
		    & " FROM Picks, Schedule" _
		    & " WHERE Username = '" & SqlString(username) & "'" _
		    & " AND Schedule.Week = " & week _
		    & " AND Picks.GameID = Schedule.GameID" _
		    & " AND Picks.Pick = Schedule." & resultField _
		    & " AND NOT ISNULL(Schedule." & resultField & ")"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			PlayerConfidenceScore = rs.Fields("Total").Value
			if not IsNumeric(PlayerConfidenceScore) then
				PlayerConfidenceScore = 0
			end if
		end if

		'If tie picks are not allowed, add half points for any tied games.
		if not ALLOW_TIE_PICKS then
			sql = "SELECT SUM(Confidence) AS Total" _
			    & " FROM Picks, Schedule" _
			    & " WHERE Username = '" & SqlString(username) & "'" _
			    & " AND Schedule.Week = " & week _
			    & " AND Picks.GameID = Schedule.GameID" _
			    & " AND Schedule." & resultField & " = '" & TIE_STR & "'"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				tieTotal = rs.Fields("Total").Value
				if IsNumeric(tieTotal) then
					PlayerConfidenceScore = PlayerConfidenceScore + tieTotal / 2
				end if
			end if
		end if

		'To improve performance, store the score in the database. The
		'Tiebreaker table is used because it has exactly one record for each
		'individual weekly entry.
		sql = "UPDATE Tiebreaker SET" _
		    & " ConfidenceScore = " & PlayerConfidenceScore _
		    & " WHERE Username = '" & SqlString(username) & "'" _
		    & " AND Week = " & week
		call DbConn.Execute(sql)

	end function

	'--------------------------------------------------------------------------
	' Returns the number of correct picks made by the specified user for the
	' given week. If the user did not enter any picks that week, an empty
	' string is returned.
	'--------------------------------------------------------------------------
	function PlayerPickScore(username, week)

		dim sql, rs, resultField, tieTotal

		PlayerPickScore = ""

		'If the user made no picks for that week, exit.
		'if TBvisGuess(username, week) = "" or TBhomeGuess(username, week) = "" then
		''	exit function
		'end if

		'See if we have the score already calculated and saved in the database.
		sql = "SELECT PickScore FROM Tiebreaker" _
		    & " WHERE Username = '" & SqlString(username) & "'" _
		    & " AND Week = " & week _
			& " AND NOT ISNULL(PickScore)"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			PlayerPickScore = rs.Fields("PickScore").Value
			exit function
		end if

		'call errormessage(PlayerPickScore & username & week)

		if PlayerPickScore = "" then
			'Determine which result field to use.
			resultField = "Result"
			if USE_POINT_SPREADS then
				resultField = "ATSResult"
			end if

			'Total the number of correct picks.
			sql = "SELECT COUNT(*) AS Total" _
			    & " FROM Picks, Schedule" _
			    & " WHERE Username = '" & SqlString(username) & "'" _
			    & " AND Schedule.Week = " & week _
			    & " AND Picks.GameID = Schedule.GameID" _
			    & " AND Picks.Pick = Schedule." & resultField _
			    & " AND NOT ISNULL(Schedule." & resultField & ")"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				PlayerPickScore = rs.Fields("Total").Value
			end if
		end if

		'call errormessage(PlayerPickScore & username & week)

		'If tie picks are not allowed, add a half point for any tied games.
		if not ALLOW_TIE_PICKS then
			sql = "SELECT COUNT(*) AS Total" _
			    & " FROM Picks, Schedule" _
			    & " WHERE Username = '" & SqlString(username) & "'" _
			    & " AND Schedule.Week = " & week _
			    & " AND Picks.GameID = Schedule.GameID" _
			    & " AND Schedule." & resultField & " = '" & TIE_STR & "'"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				tieTotal = rs.Fields("Total").Value
				if IsNumeric(tieTotal) then
					PlayerPickScore = PlayerPickScore + tieTotal / 2
				end if
			end if
		end if

		'To improve performance, store the score in the database. The
		'Tiebreaker table is used because it has exactly one record for each
		'individual weekly entry.
		sql = "UPDATE Tiebreaker SET" _
		    & " PickScore = " & PlayerPickScore _
		    & " WHERE Username = '" & SqlString(username) & "'" _
		    & " AND Week = " & week
		call DbConn.Execute(sql)
		'call errormessage(PlayerPickScore & username)
	end function

	'--------------------------------------------------------------------------
	' Returns an array of users currently in the given week's pool.
	'--------------------------------------------------------------------------
	function PoolPlayersList(week)

		dim list, sql, rs

		PoolPlayersList = ""

		sql = "SELECT Username FROM Tiebreaker" _
		   & " WHERE Week = " & week _
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
		PoolPlayersList = list

	end function

	'--------------------------------------------------------------------------
	' Returns the specified user's tiebreaker guess for the given week. An
	' empty string is returned if no value is found in the database.
	'--------------------------------------------------------------------------
	function TBvisGuess(username, week)

		dim sql, rs

		'set UserTBGuess = new TBclass
		TBvisGuess = ""
		

		sql = "SELECT VisGuess" _
		   & " FROM Tiebreaker" _
		   & " WHERE Username = '" & SqlString(username) & "'" _
		   & " AND Week = " & week
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			TBvisGuess = rs.Fields("VisGuess").Value
			'UserTBGuess.HomeScore = rs.Fields("HomeGuess").Value
		end if

	end function
	
	function TBhomeGuess(username, week)

		dim sql, rs

		sql = "SELECT HomeGuess" _
		   & " FROM Tiebreaker" _
		   & " WHERE Username = '" & SqlString(username) & "'" _
		   & " AND Week = " & week
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			TBhomeGuess = rs.Fields("HomeGuess").Value
		end if

	end function
	
	function UserTBGuess(username, week)

		dim sql, rs

		set UserTBGuess = new TBclass
		UserTBGuess.VisScore = ""
		UserTBGuess.HomeScore = ""

		sql = "SELECT VisGuess, HomeGuess" _
		   & " FROM Tiebreaker" _
		   & " WHERE Username = '" & SqlString(username) & "'" _
		   & " AND Week = " & week
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			UserTBGuess.VisScore = rs.Fields("VisGuess").Value
			UserTBGuess.HomeScore = rs.Fields("HomeGuess").Value
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the date and time of the first game scheduled for the given week.
	'--------------------------------------------------------------------------
	function WeekStartDateTime(n)

		dim sql, rs

		WeekStartDateTime = Now
		sql = "SELECT [Date], [Time]" _
		    & " FROM Schedule" _
		    & " WHERE Week = " & n _
			& " ORDER BY Date, Time"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			WeekStartDateTime = CDate(rs.Fields("Date").Value & " " & rs.Fields("Time").Value)
		end if

	end function

	'--------------------------------------------------------------------------
	' Compares scores and tiebreakers to determine the winner(s) for the week
	' specified. An array of those player names is returned. If the week has
	' not been concluded, or if no players entered that week, an empty string
	' is returned instead.
	'--------------------------------------------------------------------------
	function WinnersList(week)

		dim sql, rs, logEnabled
		dim list
		dim topUserWins, topTBdiff, topUserProx, topUserTBwin, topUserLowSpread, topUserTBwinATS, topUserTotalWins, topUserWeeklyWins
		dim userProxToActualTBspread, userTBguessVis, userTBguessHome, userTBspread, userTBdiff, userTotalWins, userWeeklyWins, userWins, userTBpick, username
		dim userTBwinATS, userTBwin
		dim TBresult, TBATSresult, tbActualHome, tbActualVis, tbActualSpread, TBwinATS, TBwin, TBspread

		logEnabled = false 		'this controls the log messages on the logic output.. true will put the messages on the poolsummary screen'

		WinnersList = ""

		'Exit if not all games have been completed.
		if NumberOfGames(week) <> NumberOfCompletedGames(week) then
			exit function
		end if

		'See if we have the winners already calculated and stored in the
		'database.
		sql = "SELECT Username FROM WeeklyWinners" _
		    & " WHERE Week = " & week _
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
		if IsArray(list) then
			WinnersList = list
			exit function
		end if

		'Get the score for the TB game and calc the actualSpread, exit if not available.
		'tbActualHome = TBPointTotalHome(week)
		'tbActualVis = TBPointTotalVis(week)
		'if not IsNumeric(tbActualHome) or not IsNumeric(tbActualVis) then
		''	exit function
		'end if
		'tbActualSpread = tbActualHome - tbActualVis
		


		'Initialize current high score and tiebreaker difference.
		topUserWins = -1						
		topTBdiff   =  0
		topUserTBwin = false	
		topUserTBwinATS = false
		topUserTotalWins = -1
		topUserWeeklyWins = -1	
		TBspread = 0
		topUserLowSpread = 99
		topUserProx = 99
		userTBspread = 99
		userProxToActualTBspread = 99
		userTBwin = 0
		userTBwinATS = 0
		
		sql = "SELECT PointSpread" _
		    & " FROM Schedule" _
		    & " WHERE TBgame = 1 AND Week = " & week _
			& " ORDER BY Date, Time"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			TBspread = rs.Fields("PointSpread").Value
		end if
		
		'call errormessage(TBspread)
		
		'Check each player who has an entry.
		'sql = "SELECT Username, VisGuess, HomeGuess, PickScore FROM Tiebreaker" _
		'   & " WHERE Week = " & week _
		'   & " ORDER BY Username"

		sql = "SELECT Schedule.Week, Picks.Username, Tiebreaker.*, Schedule.*, Picks.*, Schedule.TBgame " _
			& "FROM (Schedule INNER JOIN Picks ON Schedule.GameID = Picks.GameID) INNER JOIN Tiebreaker ON (Tiebreaker.Username = Picks.Username) AND (Schedule.Week = Tiebreaker.Week) " _
			& "WHERE (((Schedule.Week)= " & WEEK & ") AND ((Schedule.TBgame)=1));"
		set rs = DbConn.Execute(sql)

		'now going through the returned list of users from the TB table in the db'
		if not (rs.BOF and rs.EOF) then
			do while not rs.EOF
				username = rs.Fields("Username").Value

				'Find the player's score and tiebreaker.
				'score = PlayerPickScore(username, week) 'get the number of correct picks.. isnt this the same as userWins???
				userTBguessVis = cint(rs.Fields("VisGuess").Value)		'user guess for the Visitor score of the TB game'
				userTBguessHome = cint(rs.Fields("HomeGuess").value) 			'user guess for the Home score of the TB game'
				userWins = rs.Fields("PickScore").value 				'this is the total wins for the user for the given week.. this is the first element for compare'
				userTBspread = userTBguessHome - userTBguessVis 		'this is the spread between the user Home guess for TB subtracting the user Visitor guess to get the users spread'
				userTBpick = rs.Fields("Pick").Value 					'this is the actual team selected by the user with the radio button.. nothing to do with the point fields'
				userTBdiff = 99 										'this is the difference between user Guesses and the actual TB scores'
				userTBwin = 0											'this is whether the user picked the correct winner outright, 0 for no, 1 for yes'
				userTBwinATS = 0										'this is whether the user picked the correct winner against the spread, 0 for no, 1 for yes'
				
				TBresult = rs.Fields("Result").value 					'this is the name of the actual team that won the TB game with no spread involved'
				TBATSresult = rs.Fields("ATSResult").value 				'this is the name of the actual team that won taking into account the spread'
				TBspread = rs.Fields("PointSpread").value 				'this is the actual spread of the TB game'

				tbActualHome = rs.Fields("HomeScore").Value 			'this is the actual score for the home team'
				tbActualVis = rs.Fields("VisitorScore").value 				'this is the actual score for the visitor team'
				if not IsNumeric(tbActualHome) or not IsNumeric(tbActualVis) then
					exit function
				end if
				tbActualSpread = tbActualHome - tbActualVis 			'this is the actual spread from the TB game'
				userProxToActualTBspread = 99							'this is the calculation of the user spread to the actual TB game spread
				userWeeklyWins = 0										'this is the total weekly wins for this user
				userTotalWins = 0										'this is the total wins for this user'


				'Fill in the rest of the info needed for the TB calculation
				'TB Against The Spread'
				if userTBPick = TBATSresult then
					userTBwinATS = 1
				else
					userTBwinATS = 0
				end if
				'TB outright win
				if userTBPick = TBresult then
					userTBwin = 1
				else
					userTBwin = 0
				end if
				
				if IsNumeric(userWins) and IsNumeric(userTBguessVis) and IsNumeric(userTBguessHome) then
					
					'Compare this player's tiebreaker score to the current highest.
					userTBdiff = (Abs(tbActualVis - userTBguessVis) + Abs(tbActualHome - userTBguessHome))
					
					userProxToActualTBspread = abs(tbActualSpread - (userTBguessHome-userTBguessVis)) 'update for addition of closest to actual spread item
					userWeeklyWins = TotalWeeksWonLookup(username, week)
					userTotalWins = TotalWinsLookup(username, week)
					
					
					'level 1/2 - Most correct picks/ if ties for picks, TB diff decides
					'so this will loop through and take the higher at any given time and replace the current user on the list
					
					'If this player has a higher score, or the same score and a
					'closer tiebreaker, make the player the winner.
					if userWins > topUserWins or (userWins = topUserWins and userTBdiff < topTBdiff) then
						redim list(0)
						list(0)   = username' + "TBATS" + Cstr(TBATS) '+ "TBresult" + Cstr(TBresult) + "TBtotalPicks" + Cstr(TBtotalPicks) + "userTBspread" + CStr(userTBspread) + "TBweekly" + Cstr(TBweekly)
						

						topUserWins = userWins
						topTBdiff   = userTBdiff
						topUserTBwin = userTBwin	
						topUserTBwinATS = userTBwinATS
						topUserTotalWins = userTotalWins
						topUserWeeklyWins = userWeeklyWins
						topUserLowSpread = userTBspread
						topUserProx = userProxToActualTBspread
						
						'	Order of the winner list
						'  Lowest point differential of tiebreaker game
						'  Who correctly picked the winner of the tiebreaker game versus the spread.
						'  Who correctly picked the winner of the tiebreaker game straight up (regardless of spread).
						'  Closest to the actual spread. (added 2012)
						'  Most overall wins.
						'  Most weekly wins. 
						
						'TBATS - TB against the spread, bool if correctly picked winner ATS
						'TBresult - TB result without the spread, bool if correctly picked the winner straight up
						'TBtotalPicks - total correct picks for user over season
						'TBweekly - total weeks won for user
						'userTBguessVis - TB guess for visitor
						'userTBguessHome - TB guess for home
						'TBspread = spread for the TB game for this week

						if logEnabled = true then
							call errormessage("1 " & username _									
										& ", Level 1 - userWins " & userWins _
										& ", Level 2 - userTBwinATS " & userTBwinATS _
										& ", Level 3 - userTBwin " & userTBwin _
										& ", Level 4 - userProxToActualTBspread " & userProxToActualTBspread _
										& ", Level 5 - userTotalWins " & userTotalWins _
										& ", Level 6 - userWeeklyWins " & userWeeklyWins _
										& ", userTBdiff " & userTBdiff _
										& ", userTBspread " & userTBspread)	

							call errormessage(tbActualSpread & ", " & userTBguessHome & ", " & userTBguessVis)
						end if


					'this is where we are going to add more levels so there is no tie and multiple winners
					elseif userWins = topUserWins and userTBdiff = topTBdiff then
						if 	(userTBwinATS = true and topUserTBwinATS = false) then 
							redim list(0)
								list(0)   = username
								topUserWins = userWins
								topTBdiff   = userTBdiff
						

								topUserWins = userWins
								topTBdiff   = userTBdiff
								topUserTBwin = userTBwin	
								topUserTBwinATS = userTBwinATS
								topUserTotalWins = userTotalWins
								topUserWeeklyWins = userWeeklyWins
								topUserLowSpread = userTBspread
								topUserProx = userProxToActualTBspread

								if logEnabled = true then
									call errormessage("2 " & username _									
										& ", Level 1 - userWins " & userWins _
										& ", Level 2 - userTBwinATS " & userTBwinATS _
										& ", Level 3 - userTBwin " & userTBwin _
										& ", Level 4 - userProxToActualTBspread " & userProxToActualTBspread _
										& ", Level 5 - userTotalWins " & userTotalWins _
										& ", Level 6 - userWeeklyWins " & userWeeklyWins _
										& ", userTBdiff " & userTBdiff _
										& ", userTBspread " & userTBspread)	
								end if

						elseif((userTBwinATS = topUserTBwinATS) and (userTBwin > topUserTBwin)) then 
							redim list(0)
								list(0)   = username
								topUserWins = userWins
								topTBdiff   = userTBdiff
						

								topUserWins = userWins
								topTBdiff   = userTBdiff
								topUserTBwin = userTBwin	
								topUserTBwinATS = userTBwinATS
								topUserTotalWins = userTotalWins
								topUserWeeklyWins = userWeeklyWins
								topUserLowSpread = userTBspread
								topUserProx = userProxToActualTBspread

								if logEnabled = true then
									call errormessage("3 " & username _
										& ", Level 1 - userWins " & userWins _
										& ", Level 2 - userTBwinATS " & userTBwinATS _
										& ", Level 3 - userTBwin " & userTBwin _
										& ", Level 4 - userProxToActualTBspread " & userProxToActualTBspread _
										& ", Level 5 - userTotalWins " & userTotalWins _
										& ", Level 6 - userWeeklyWins " & userWeeklyWins _
										& ", userTBdiff " & userTBdiff _
										& ", userTBspread " & userTBspread)		
								end if

						elseif((userTBwinATS = topUserTBwinATS) and (userTBwin = topUserTBwin) and (userProxToActualTBspread < topUserProx)) then
								redim list(0)
								list(0)   = username
								topUserWins = userWins
								topTBdiff   = userTBdiff
						

								topUserWins = userWins
								topTBdiff   = userTBdiff
								topUserTBwin = userTBwin	
								topUserTBwinATS = userTBwinATS
								topUserTotalWins = userTotalWins
								topUserWeeklyWins = userWeeklyWins
								topUserLowSpread = userTBspread
								topUserProx = userProxToActualTBspread

								if logEnabled = true then
									call errormessage("4 " & username _									
										& ", Level 1 - userWins " & userWins _
										& ", Level 2 - userTBwinATS " & userTBwinATS _
										& ", Level 3 - userTBwin " & userTBwin _
										& ", Level 4 - userProxToActualTBspread " & userProxToActualTBspread _
										& ", Level 5 - userTotalWins " & userTotalWins _
										& ", Level 6 - userWeeklyWins " & userWeeklyWins _
										& ", userTBdiff " & userTBdiff _
										& ", userTBspread " & userTBspread)		

									call errormessage(userTBguessHome & userTBguessVis & cint(userTBguessHome-userTBguessVis) & tbActualSpread)
								end If
								
						elseif((userTBwinATS = topUserTBwinATS) and (userTBwin = topUserTBwin) and (userProxToActualTBspread = topUserProx) and (userTotalWins > topUserTotalWins)) then 
								'so get to this point if current user flat out winner compared to users reviewed so far in loop.. so make them the new current high player
								redim list(0)
								list(0)   = username
								topUserWins = score
								topTBdiff   = userTBdiff
						

								topUserWins = userWins
								topTBdiff   = userTBdiff
								topUserTBwin = userTBwin	
								topUserTBwinATS = userTBwinATS
								topUserTotalWins = userTotalWins
								topUserWeeklyWins = userWeeklyWins
								topUserLowSpread = userTBspread
								topUserProx = userProxToActualTBspread

								if logEnabled = true then
									call errormessage("5 " & username _									
										& ", Level 1 - userWins " & userWins _
										& ", Level 2 - userTBwinATS " & userTBwinATS _
										& ", Level 3 - userTBwin " & userTBwin _
										& ", Level 4 - userProxToActualTBspread " & userProxToActualTBspread _
										& ", Level 5 - userTotalWins " & userTotalWins _
										& ", Level 6 - userWeeklyWins " & userWeeklyWins _
										& ", userTBdiff " & userTBdiff _
										& ", userTBspread " & userTBspread)	
								end if

						elseif ((userTBwinATS = topUserTBwinATS) and (userTBwin = topUserTBwin) and (userProxToActualTBspread = topUserProx) and (userTotalWins = topUserTotalWins) and (userWeeklyWins => topUserWeeklyWins)) then
							'if get in here then at the last point in the comparison for the winner and comes down to most weeks won
							'so either the user will have more weeks won and be the winner or same number and tie so will add to the list of winners
							if userWeeklyWins > topUserWeeklyWins then 
					
								redim list(0)
								list(0)   = username
								topUserWins = userWins
								topTBdiff   = userTBdiff
						

								topUserWins = userWins
								topTBdiff   = userTBdiff
								topUserTBwin = userTBwin	
								topUserTBwinATS = userTBwinATS
								topUserTotalWins = userTotalWins
								topUserWeeklyWins = userWeeklyWins
								topUserLowSpread = userTBspread
								topUserProx = userProxToActualTBspread

								if logEnabled = true then
									call errormessage("6 " & username _									
										& ", Level 1 - userWins " & userWins _
										& ", Level 2 - userTBwinATS " & userTBwinATS _
										& ", Level 3 - userTBwin " & userTBwin _
										& ", Level 4 - userProxToActualTBspread " & userProxToActualTBspread _
										& ", Level 5 - userTotalWins " & userTotalWins _
										& ", Level 6 - userWeeklyWins " & userWeeklyWins _
										& ", userTBdiff " & userTBdiff _
										& ", userTBspread " & userTBspread)	
								end if
							else
								'this was where it would add multiple users to winners list
								redim preserve list(UBound(list) + 1)
								list(UBound(list)) = username

								if logEnabled = true then

									call errormessage("7 " & username _									
										& ", Level 1 - userWins " & userWins _
										& ", Level 2 - userTBwinATS " & userTBwinATS _
										& ", Level 3 - userTBwin " & userTBwin _
										& ", Level 4 - userProxToActualTBspread " & userProxToActualTBspread _
										& ", Level 5 - userTotalWins " & userTotalWins _
										& ", Level 6 - userWeeklyWins " & userWeeklyWins _
										& ", userTBdiff " & userTBdiff _
										& ", userTBspread " & userTBspread)	
								end if
							end if
						end if	
					end if
				end if
				rs.MoveNext
			loop
		end if

		'To improve performance, store the winners in the database. A separate
		'table named WeeklyWinners is used for this.
		if IsArray(list) then
			dim i
			for i = LBound(list) to UBound(list)
				sql = "INSERT INTO WeeklyWinners" _
				   & " (Week, Username)" _
				   & " VALUES(" & week & ", '" & SqlString(list(i)) & "')"
				call DbConn.Execute(sql)
			next
		end if

		WinnersList = list

	end function 
	
	
	
	'_____________________________________________
	
		
	'--------------------------------------------------------------------------
	' This is for the calculation of the weekly winners in more detail
	'--------------------------------------------------------------------------
	function TBpicksLookup(username, week)

		dim sql, rs
		dim TBpick, result, ATSresult, spread
		dim TBresult, TBATS
		dim TBweekly

		TBresult = 0
		TBATS = 0
		
		sql = "SELECT * FROM PICKS, SCHEDULE " _
			& "WHERE PICKS.GAMEID = SCHEDULE.GAMEID " _
			& "AND SCHEDULE.WEEK = " & week & " " _
			& "AND TBGAME = 1 " _
			& "AND PICKS.USERNAME = '" & SqlString(username) & "'"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			result = rs.Fields("Result").Value
			ATSresult = rs.Fields("ATSResult").Value
			TBpick = rs.Fields("Pick").Value
			spread = rs.Fields("PointSpread").Value				
		end if
		
		if TBPick = result then
			TBresult = 1
		else
			TBresult = 0
		end if
		
		if TBPick = ATSresult then
			TBATS = 1
		else
			TBATS = 0
		end if
		

	end function
	
	'--------------------------------------------------------------------------
	' This is for the calculation of the total correct picks for a user and weekly winners
	'--------------------------------------------------------------------------
	function TBwinsLookup(username)

		dim sql, rs
		dim TBtotalPicks
		dim TBweekly

		TBtotalPicks = 0
		TBweekly = 0
		
		sql = "SELECT COUNT(WEEK) AS TOTAL FROM WEEKLYWINNERS WHERE USERNAME = '" & SqlString(username) & "'"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			TBweekly = rs.Fields("TOTAL").Value			
		end if
		
		sql = "SELECT COUNT(PICKS.PICK) AS TOTAL FROM PICKS " _
		& "inner JOIN SCHEDULE ON PICKS.GAMEID = SCHEDULE.GAMEID " _
		& "WHERE PICKS.USERNAME = '" & SqlString(username) & "' " _
		& "AND SCHEDULE.GAMEID = PICKS.GAMEID " _
		& "AND PICKS.PICK = SCHEDULE.ATSRESULT " 
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			TBtotalPicks = rs.Fields("TOTAL").Value		
		end if
		

	end function
	
	
	
	'--------------------------------------------------------------------------
	' This is for the calculation of the total correct picks for a user. The week parameter is to look only at weeks equal to or prior so if re-run later down the line handles correctly
	'--------------------------------------------------------------------------
	function TotalWinsLookup(username, week)
		dim sql, rs
		
		sql = "SELECT COUNT(PICKS.PICK) AS TOTAL FROM PICKS " _
		& "inner JOIN SCHEDULE ON PICKS.GAMEID = SCHEDULE.GAMEID " _
		& "WHERE PICKS.USERNAME = '" & SqlString(username) & "' " _
		& "AND SCHEDULE.GAMEID = PICKS.GAMEID " _
		& "AND PICKS.PICK = SCHEDULE.ATSRESULT " _
		& "AND SCHEDULE.WEEK <= " & week 
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			TotalWinsLookup = rs.Fields("TOTAL").Value		
		end if
		'call errormessage(username & TotalWinsLookup & " week " & week)
	end function	
	
	'--------------------------------------------------------------------------
	' This is for the calculation of the total weekly wins for a user. The week parameter is to look only at weeks equal to or prior so if re-run later down the line handles correctly
	'--------------------------------------------------------------------------
	function TotalWeeksWonLookup(username, week)
		dim sql, rsWeeks
		
		sql = "SELECT COUNT(*) AS TOTAL FROM WEEKLYWINNERS WHERE USERNAME = '" & SqlString(username) & "' AND WEEK <= " & week 
		set rsWeeks = DbConn.Execute(sql)
		if not (rsWeeks.BOF and rsWeeks.EOF) then
			TotalWeeksWonLookup = rsWeeks.Fields("TOTAL").Value			
		end if
	end function	
	
	%>
