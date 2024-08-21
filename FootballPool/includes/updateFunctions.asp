	<%	'these are the functions for the update Schedule piece
	'they are modded versions and fresh functions for the updateSchedule.asp 
	'page which allows admin to choose the games to be used in the pool
    
	
	
		
	'__________________________________________________________________________________________
	
	'  These are the FullSchedule functions
	
	'__________________________________________________________________________________________
	
	
		'--------------------------------------------------------------------------
	' Gets the week number specified in the HTTP request. If no valid number
	' was specified, the current week is returned instead. This is updated for the FullSchedule weeks
	'--------------------------------------------------------------------------
	function GetRequestedWeekFS()

		dim week, sql
		week = 1
		'Default to the current week.
		GetRequestedWeekFS = CurrentWeekFS()

		'Check the week number specified in the HTTP request.
		week = Request("week")
		if not IsNumeric(week) then
			exit function
		else
			week = Round(week)
			if week < 1 or week > NumberOfFSWeeks() then
				exit function
			end if
		end if
		GetRequestedWeekFS = week

	end function
	
	
	'--------------------------------------------------------------------------
	' Returns the current week number based on the current (time zone adjusted)
	' date and time.
	'--------------------------------------------------------------------------
	function CurrentWeekFS()

		dim dateNow, sql, rs, found

		CurrentWeekFS = 0
		dateNow = CurrentDateTime
		sql = "SELECT Week, GameDate FROM FullSchedule ORDER BY GameDate"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			found = false
			do while not rs.EOF and not found
				if DateDiff("d", DateValue(rs.Fields("GameDate").Value), dateNow) <= 0 then
					found = true
					CurrentWeekFS = rs.Fields("Week").Value
				end if
				rs.MoveNext
			loop
		end if

		if CurrentWeekFS = 0 then
			CurrentWeekFS = NumberOfFSWeeks()
		end if

	end function
	
	
	'--------------------------------------------------------------------------
	' Returns the number of games scheduled for the specified week from the fullschedule table.
	'--------------------------------------------------------------------------
	function NumberOfGamesFS(week)

		dim sql, rs

		NumberOfGamesFS = 0
		sql = "SELECT COUNT(*) AS Total FROM FullSchedule WHERE Week = " & week
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			NumberOfGamesFS = rs.Fields("Total").Value
		end if

	end function	
	
	
	'--------------------------------------------------------------------------
	' Returns the total number of weeks in the FullSchedule table.
	'--------------------------------------------------------------------------
	function NumberOfFSWeeks()

		dim sql, rs

		NumberOfFSWeeks = 0
		sql = "SELECT MAX(Week) AS Total FROM FullSchedule"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			NumberOfFSWeeks = rs.Fields("Total").Value
		end if

	end function	
	
	
	' 
	'--------------------------------------------------------------------------
	' Returns a record set of games for the specified week but from the FullSched table.
	'--------------------------------------------------------------------------
	function WeeklyFullSchedule(week)

		dim sql
		'& "," _  
		'& "," _ , ,  
		'Retrieve the week's schedule from the database.
		'sql = "SELECT ID, Week, GameDate, GameTime, DisplayValue, InPool, GameId, PointSpread, VisTeam, HomeTeam FROM FullSchedule WHERE Week = " & week
		
		sql = "SELECT FullSchedule.*, VTeams.TeamID AS visTeamID, VTeams.Logo AS logoVis, HTeams.TeamID AS homeTeamID, HTeams.Logo AS logoHome " _
			& "FROM FullSchedule, Teams VTeams, Teams HTeams " _
			& "WHERE FullSchedule.VisTeam = VTeams.City " _
			& "AND FullSchedule.HomeTeam = HTeams.City " _
			& "AND Week = " & week & " " _
			& "ORDER BY GameDate, GameTime, VTeams.City"
			
		
		set WeeklyFullSchedule = DbConn.Execute(sql)

	end function	
	
	
	'--------------------------------------------------------------------------
	' Returns the number of games in the FullSchedule table scheduled for the specified week.
	'--------------------------------------------------------------------------
	function NumberOfGamesScheduled(week)

		dim sql, rs

		NumberOfGamesScheduled = 0
		sql = "SELECT COUNT(*) AS Total FROM FullSchedule WHERE Week = " & week
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			NumberOfGamesScheduled = rs.Fields("Total").Value
		end if

	end function	
	
	'--------------------------
	'this function is to update the schedule table based on the games picked on the screen
	'--------------------------
	
	
	Function updateSchedule(week)
		
		dim gameID, gameWeek, gameDate, gameTime, displayValue, homeTeam
		dim homeTeamID, visTeam, visTeamID, viewTime, pointSpread, fullGameID, inPool
		dim logoVis, logoHome, maxGameID, infoMsg, espnGameId
		
		'this is the update for the setSchedule
		'first we will clear any entries from the schedule table for the given week
		sql = "DELETE FROM SCHEDULE WHERE WEEK = " & week
		call DbConn.Execute(sql)
		
		sql = "UPDATE FullSchedule " _
			& "SET PointSpread = NULL, " _
			& "InPool = 'No' " _
			& "WHERE Week = " & week
			call DbConn.Execute(sql)
		
		'now we need to insert into the schedule table and UPDATE the fullschedule table
		
		infoMsg = "Update not successful."
		FOR i = 1 to n 'n being the number of entries displayed
		'so this will be done for each row
		'if (checkbox for the row = checked) then
			inPool = Trim(Request.Form("inPool-" & i))
			if LCase(inPool) = "true" then
				'now to insert into the schedule table first
				gameID = Trim(Request.Form("id-"     & i))
				gameDate = Trim(Request.Form("date-" 	& i))
				gameTime = Trim(Request.Form("time-" 	& i))
				pointSpread = Trim(Request.Form("spread-"   & i))
				visTeamID = Trim(Request.Form("visID-"   & i))
				homeTeamID = Trim(Request.Form("homeID-"   & i))
				espnGameId = Trim(Request.Form("espnGameId-"   & i))
				
				if isDate(gameDate) then
					gameDate = CDate(gameDate)
					gameDate = FormatDateTime(gameDate, 0)
				end if
				if isDate(gameTime) then
					gameTime = CDate(gameTime)
				end if
				if not IsNumeric(pointSpread) then 
					pointSpread = 0
				end if
		
				'now to get the highest value from the schedule table so can use in the insert and updates
				'sql = "SELECT MAX(GameID) AS HIGH FROM SCHEDULE"
				'set rs = DbConn.Execute(sql)
				'maxGameID = Cint(rs.Fields("HIGH").Value)
				'maxGameID = maxGameID + 1
	'-------------------------			
	'debug area
				
				'infoMsg = infoMsg & "week = " & week _
				'& " gameID = " & gameID _
				'& " gamedate = " & gamedate _
				'& " gameTime = " & gameTime _
				'& " pointSpread = " & pointSpread _
				'& " visTeamID = " & visTeamID _
				'& " homeTeamID = " & homeTeamID
				
	'------------------------
				
				'now to insert the new values into the Schedule table
				dim gameTimeSql
				
				if gameTime = "TBA" then 
					gameTimeSql = "NULL"
				else
					gameTimeSql = "'" & gameTime & "'"
				end if
					
				sql = "INSERT INTO SCHEDULE ( Week, [Date], [Time], VisitorID, PointSpread, EspnGameId, HomeID) " _
					& "VALUES (" & week & ", '" & gameDate & "', " & gameTimeSql & ", '" & visTeamID & "', " & pointSpread & ", '" & espnGameId & "', '" & homeTeamID & "')"
				call DbConn.Execute(sql)
				
				'and now to update the fullSchedule table as well

					
				sql = "UPDATE FullSchedule " _
					& "SET GameDate = '" & gameDate & "', " _
					& "GameTime = " & gameTimeSql & ", " _	
					& "PointSpread = " & pointSpread & ", " _
					& "InPool = 'Yes' " _
					& "WHERE ID = " & gameID
				call DbConn.Execute(sql)
				infoMsg = "Update Successful."
			end if	
		next
		call InfoMessage(infoMsg)
	end Function	
	
	'taking gameid piece out for now
	'& "GameId = '" & maxGameID & "', " _
	
	
	
	

	
	
	
	
	'_________________________________________________________
	
	'temp work
  	'dim HomeTeamID, VisitorTeamID, maxGameID, dateVal, timeVal, pSpread
		
		
		'this is the update for the setSchedule
		'first we will clear any entries from the schedule table for the given week
		'sql = "DELETE FROM SCHEDULE WHERE WEEK = " & week
		'call DbConn.Execute(sql)
		
		'visTeam = Request("spread-1")
			'if (Request("inPool-1") <> "") then
		
				'sql = "INSERT INTO SCHEDULE (VisitorID) VALUES (" & visTeam & ")"
				'call DbConn.Execute(sql)
			'end if
		
		'now we need to insert into the schedule table and UPDATE the fullschedule table

		'FOR i = 1 to n 'n being the number of entries displayed
		
		
		'so this will be done for each row
		'if (checkbox for the row = checked) then
		'if (Request.Form("inPool-" & i) = "Yes") then
		
			'sql = "INSERT INTO SCHEDULE (WEEK) VALUES ('1')"
			'call DbConn.Execute(sql)
		
		
				'gameID = Trim(Request.Form("id-"     & i))
				'dateVal = Trim(Request.Form("date-" 	& i))
				'timeVal = Trim(Request.Form("time-" 	& i))
				'pSpread = Trim(Request.Form("spread-"   & i))
				'visTeam = Trim(Request.Form("visTeam-"   & i))
				'homeTeam = Trim(Request.Form("homeTeam-"   & i))
				
				'now we need to look up the ID for home and Visitor
				'sql = "SELECT TeamID FROM Teams WHERE City = '" & homeTeam & "'"
				'call DbConn.Execute(sql)
				'HomeTeamID = rs.Fields("TeamID").Value
 		
			
				'sql = "SELECT TeamID FROM TEAMS	WHERE CITY = '" & VisTeam & "'"
				'call DbConn.Execute(sql)
				'VisitorTeamID = rs.Fields("TeamID").Value
				'now insert the new game into the schedule table
				
				'now to get the highest value from the schedule table so can use in the insert and updates
				'sql = "SELECT MAX(GameID) FROM SCHEDULE"
				'call DbConn.Execute(sql)
				'maxGameID = rs.Fields("GameID").Value
				'maxGameID = maxGameID + 1
				%>
				
				<% 
				
				'now to insert the new values into the Schedule table
				'sql = "INSERT INTO SCHEDULE (GameID, Week, Date, Time, VisitorID, PointSpread, HomeID) " _
				'	& "VALUES (" & maxGameID _
				'	& ", " & week _ 
				'	& ", " & dateVal _
				'	& "', '" & timeVal _
				'	& "', '" & VisitorTeamID _
				'	& "', " & pSpread _
				'	& ", '" & HomeTeamID _
				'	& "')"
				'call DbConn.Execute(sql)
				
				
				
				'and now to update the fullSchedule table as well
				'sql = "UPDATE FullSchedule " _
				'	& "SET GameDate = '" & dateVal & "', " _
				'	& "GameTime = '" & timeVal & "', " _
				'	& "PointSpread = " & pSpread & ", " _
				'	& "InPool = 'Yes'," _
				'	& "WHERE ID = '" & gameID & "')"
			'call DbConn.Execute(sql)
				
			'end if	
		'next %>
	