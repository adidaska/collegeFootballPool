<%	'**************************************************************************
	'* Common code for the playoffs pool.                                     *
	'**************************************************************************

	const TO_BE_DETERMINED_STR = "[<em>TBD</em>]"

	'**************************************************************************
	'* Global variables.                                                      *
	'**************************************************************************

	'Playoff round names.
	dim PlayoffRoundNames
	PlayoffRoundNames = Array("Wild Card Games", "Divisional Playoffs", "Conference Championships", "Super Bowl XLIII - Raymond James Stadium - Tampa, FL")

	'**************************************************************************
	'* Functions and subroutines.                                             *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Returns the string "at" unless the given game is the Super Bowl, in which
	' case the string "vs." is returned.
	'--------------------------------------------------------------------------
	function GetConjunction(gameRound)

		GetConjunction = "at"
		if gameRound = NumberOfPlayoffRounds() then
			GetConjunction = "vs."
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the name of the conference the given team belongs to.
	'--------------------------------------------------------------------------
	function GetConference(teamID)

		dim sql, rs

		GetConference = 0
		sql = "SELECT Conference FROM Teams WHERE TeamID = '" & teamID & "'"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			GetConference = rs.Fields("Conference").Value
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the display name of the given team.
	'--------------------------------------------------------------------------
	function GetTeamName(teamID)

		dim sql, rs

		GetTeamName = TO_BE_DETERMINED_STR
		sql = "SELECT City, DisplayName FROM Teams WHERE TeamID = '" & teamID & "'"
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
	' Returns true if the given user has entered any picks in the playoffs
	' pool.
	'--------------------------------------------------------------------------
	function InPlayoffsPool(username)

		dim sql, rs

		InPlayoffsPool = false
		sql = "SELECT COUNT(*) AS Total FROM PlayoffsPicks" _
		   & " WHERE Username = '" & SqlString(username) & "'"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			if rs.Fields("Total").Value > 0 then
				InPlayoffsPool = true
			end if
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the number of completed playoff games (i.e., games with a
	' result).
	'--------------------------------------------------------------------------
	function NumberOfCompletedPlayoffGames()

		dim sql, rs

		NumberOfCompletedPlayoffGames = 0
		sql = "SELECT COUNT(*) AS Total" _
		   & " FROM PlayoffsSchedule" _
		   & " WHERE NOT ISNULL(Result)"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			NumberOfCompletedPlayoffGames = rs.Fields("Total").Value
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the number of players who have entries in the playoffs pool.
	'--------------------------------------------------------------------------
	function NumberOfPlayoffsEntries()

		dim sql, rs

		NumberOfPlayoffsEntries = 0
		sql = "SELECT COUNT(*) AS Total FROM (SELECT DISTINCT Username FROM PlayoffsPicks)"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			NumberOfPlayoffsEntries = rs.Fields("Total").Value
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the number of playoff games in the schedule.
	'--------------------------------------------------------------------------
	function NumberOfPlayoffGames()

		dim sql, rs

		NumberOfPlayoffGames = 0
		sql = "SELECT COUNT(*) AS Total FROM PlayoffsSchedule"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			NumberOfPlayoffGames = rs.Fields("Total").Value
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the number of rounds in the playoff schedule.
	'--------------------------------------------------------------------------
	function NumberOfPlayoffRounds()

		dim sql, rs

		NumberOfPlayoffRounds = 0
		sql = "SELECT MAX(Round) AS Total FROM PlayoffsSchedule"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			NumberOfPlayoffRounds = rs.Fields("Total").Value
		end if

	end function

	'--------------------------------------------------------------------------
	' Set the results for the given playoff game based on the game score and
	' point spread.
	'--------------------------------------------------------------------------
	sub PlayoffsSetGameResults(id)

		dim sql, rs, vid, hid, vscore, hscore, spread, result, atsResult

		sql = "SELECT * FROM PlayoffsSchedule WHERE GameID = " & id
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			vid    = rs.Fields("VisitorID").Value
			hid    = rs.Fields("HomeID").Value
			vscore = rs.Fields("VisitorScore").Value
			hscore = rs.Fields("HomeScore").Value
			spread = rs.Fields("PointSpread").Value

			result = "NULL"
			atsResult = "NULL"

			'If scores are available, determine the results.
			if IsNumeric(vscore) and IsNumeric(hscore) then

				'Find the winner.
				vscore = CInt(vscore)
				hscore = CInt(hscore)
				if vscore > hscore then
					result = vid
				elseif hscore > vscore then
					result = hid
				else
					result = TIE_STR
				end if
				result = "'" & result & "'"

				'Find the winner vs. the spread
				if IsNumeric(spread) then
					spread = CDbl(spread)
				else
					spread = 0
				end if
				if (vscore + spread) > hscore then
					atsResult = vid
				elseif hscore > (vscore + spread) then
					atsResult = hid
				else
					atsResult = TIE_STR
				end if
				atsResult = "'" & atsResult & "'"
			end if

			'Update the results.
			sql = "UPDATE PlayoffsSchedule SET" _
			   & " Result       = " & result & ", " _
			   & " ATSResult    = " & atsResult _
			   & " WHERE GameID = " & id
			call DbConn.Execute(sql)
		end if

	end sub

	'--------------------------------------------------------------------------
	' Returns an array of users currently in the playoffs pool.
	'--------------------------------------------------------------------------
	function PlayoffsPoolPlayersList()

		dim list, sql, rs

		PlayoffsPoolPlayersList = ""

		sql = "SELECT DISTINCT Username FROM PlayoffsPicks" _
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
		PlayoffsPoolPlayersList = list

	end function

	'--------------------------------------------------------------------------
	' Returns the date of the first playoff game.
	'--------------------------------------------------------------------------
	function PlayoffsStartDateTime()

		dim sql, rs, datetime

		PlayoffsStartDateTime = ""
		sql = "SELECT * FROM PlayoffsSchedule ORDER BY Date, Time"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			PlayoffsStartDateTime = CDate(rs.Fields("Date").Value & " " & rs.Fields("Time").Value)
		end if

	end function

	'--------------------------------------------------------------------------
	' Compares scores to determine the winner(s) for the playoffs pool.
	' If the pool has not been concluded or if no players entered, an empty
	' string is returned instead.
	'--------------------------------------------------------------------------
   function PlayoffsWinnersList()

		dim sql, rs
		dim highScore
		dim username, score
		dim list

		PlayoffsWinnersList = ""

		'Exit if not all games have been completed.
		if NumberOfPlayoffGames() <> NumberOfCompletedPlayoffGames() then
			exit function
		end if

		'Initialize current high score.
		highScore = -1

		'Check each player.
		sql = "SELECT Username FROM Users" _
		   & " WHERE Username <> '" & SqlString(ADMIN_USERNAME) & "'"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			do while not rs.EOF
				username = rs.Fields("Username").Value
				if USE_CONFIDENCE_POINTS then
					score = PlayoffsPlayerConfidenceScore(username)
				else
					score = PlayoffsPlayerPickScore(username)
				end if
				if IsNumeric(score) then

					'If this player has a higher score, make the player the
					'winner.
					if score > highScore then
						redim list(0)
						list(0) = username
						highScore = score

					'Otherwise, if this player has the same score, add the
					'player to the winners list.
					elseif score = highScore then
						redim preserve list(UBound(list) + 1)
						list(UBound(list)) = username
					end if
				end if
				rs.MoveNext
			loop
		end if

		PlayoffsWinnersList = list

	end function

	'--------------------------------------------------------------------------
	' Returns true if the given playoff round is locked. 
	'--------------------------------------------------------------------------
	function RoundLocked(n)

		dim sql, rs, datetime

		RoundLocked = true
		sql = "SELECT * FROM PlayoffsSchedule" _
		  & " WHERE Round = " & n _
		  & " ORDER BY Date, Time"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			datetime = CDate(rs.Fields("Date").Value & " " & rs.Fields("Time").Value)
			if datetime > CurrentDateTime() then
				RoundLocked = false
			end if
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns a record set of all teams.
	'--------------------------------------------------------------------------
	function Teams()

		dim sql

		sql = "SELECT * FROM Teams ORDER BY City, Name"
		set Teams = DbConn.Execute(sql)

	end function

	'--------------------------------------------------------------------------
	' Returns the number of correct picks made by the specified user for the
	' playoff games. If the user did not enter any picks, an empty string is
	' returned.
	'--------------------------------------------------------------------------
	function PlayoffsPlayerPickScore(username)

		dim sql, rs, resultField

		PlayoffsPlayerPickScore = ""

		'Determine if the user has made an entry.
		sql = "SELECT COUNT(*) AS Total FROM PlayoffsPicks" _
		    & " WHERE Username = '" & SqlString(username) & "'"
		set rs = DbConn.Execute(sql)
		if rs.Fields("Total").Value = 0 then
			exit function
		end if

		'Determine which result field to use.
		resultField = "Result"
		if USE_POINT_SPREADS then
			resultField = "ATSResult"
		end if

		'Total the number of correct picks.
		sql = "SELECT COUNT(*) AS Total" _
		    & " FROM PlayoffsPicks, PlayoffsSchedule" _
		    & " WHERE Username = '" & SqlString(username) & "'" _
		    & " AND PlayoffsPicks.GameID = PlayoffsSchedule.GameID" _
		    & " AND PlayoffsPicks.Pick = PlayoffsSchedule." & resultField _
		    & " AND NOT ISNULL(PlayoffsSchedule." & resultField & ")"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			PlayoffsPlayerPickScore = rs.Fields("Total").Value
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the confidence score for picks made by the specified user for the
	' playoff games. If the user did not enter any picks that week, an empty
	' string is returned.
	'--------------------------------------------------------------------------
	function PlayoffsPlayerConfidenceScore(username)

		dim sql, rs, resultField

		PlayoffsPlayerConfidenceScore = ""

		'Determine if the user has made an entry.
		sql = "SELECT COUNT(*) AS Total FROM PlayoffsPicks" _
		    & " WHERE Username = '" & SqlString(username) & "'"
		set rs = DbConn.Execute(sql)
		if rs.Fields("Total").Value = 0 then
			exit function
		end if

		'Determine which result field to use.
		resultField = "Result"
		if USE_POINT_SPREADS then
			resultField = "ATSResult"
		end if

		'Total the confidence points for each correct pick.
		sql = "SELECT SUM(Confidence) AS Total" _
		    & " FROM PlayoffsPicks, PlayoffsSchedule" _
		    & " WHERE Username = '" & SqlString(username) & "'" _
		    & " AND PlayoffsPicks.GameID = PlayoffsSchedule.GameID" _
		    & " AND PlayoffsPicks.Pick = PlayoffsSchedule." & resultField _
		    & " AND NOT ISNULL(PlayoffsSchedule." & resultField & ")"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			PlayoffsPlayerConfidenceScore = rs.Fields("Total").Value
			if not IsNumeric(PlayoffsPlayerConfidenceScore) then
				PlayoffsPlayerConfidenceScore = 0
			end if
		end if

	end function %>
