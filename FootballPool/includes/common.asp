<%	'this is the prod version
	'**************************************************************************
	'* ASP Football Pool                                                      *
	'* Version 2.2                                                            *
	'*                                                                        *
	'* Do not remove this notice.                                             *
	'*                                                                        *
	'* Copyright 1999-2008 by Mike Hall.                                      *
	'* Distributed under the terms of the GNU General Public License.         *
	'* Please see http://www.brainjar.com for more information.               *
	'**************************************************************************

	'**************************************************************************
	'* Common application code.                                               *
	'*                                                                        *
	'* Note: Should be included at the beginning of every page.               *
	'**************************************************************************

	Option Explicit
	Response.Buffer = true

	'**************************************************************************
	'* Configuration section.                                                 *
	'*                                                                        *
	'* Note: Set these options to suit your particular installation and       *
	'* define the pool format.                                                *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Site settings.
	'--------------------------------------------------------------------------

	'Site URL, used for the emails sent to users.
	const POOL_URL = "http://kagammapi.com/footballpool"

	'Page title.
	const PAGE_TITLE = "JAFL - Just Another Football League"

	'Administrator email address.
	const ADMIN_EMAIL = "Andres.arbelaez1@icloud.com"

	'Flag indicating whether the server can send email or not.
	const SERVER_EMAIL_ENABLED = false

	'Sign up invitation code.
	const SIGN_UP_INVITATION_CODE = ""

	'Time zone difference between host server and database schedule (Eastern).
	const TIME_ZONE_DIFF = 0

	'Location of the database.
	const DATABASE_FILE_PATH = "/access_db/NCAA.mdb"

	'Key used for encrypting private player data.
	const DATA_ENCRYPTION_KEY = "4k0s*b92mqld:pdw7d4sklda8adz@d{a"

	'Username session variable.
	const SESSION_USERNAME_KEY = "FootballPoolUsername"

	'Set the session timeout length. Default is 20 minutes
     const SESSION_TIMEOUT_LENGTH = 45

	'--------------------------------------------------------------------------
	' Pool settings.
	'--------------------------------------------------------------------------
	
	'Cost of a weekly pool entry.
	const SHOW_MONEY = false
	
	'Cost of a weekly pool entry.
	const BET_AMOUNT = 2.5

	'Allow tie picks flag.
	const ALLOW_TIE_PICKS = false

	'Hide other players picks until the games are locked flag.
	const HIDE_OTHERS_PICKS = true

	'Use confidence points flag.
	const USE_CONFIDENCE_POINTS = false

	'Use point spreads flag.
	const USE_POINT_SPREADS = true

	'--------------------------------------------------------------------------
	' Playoffs pool settings.
	'--------------------------------------------------------------------------

	'Enable playoffs pool pages flag.
	const ENABLE_PLAYOFFS_POOL = false

	'Cost of a playoffs pool entry.
	const PLAYOFFS_BET_AMOUNT = 20

	'--------------------------------------------------------------------------
	' Side pool options.
	'
	' Note: You can choose to add a survivor pool, a margin pool or both. Note
	' that if both are enabled, user's make a single pick each week which
	' will apply to both pools.
	'--------------------------------------------------------------------------

	'Enable survivor pool.
	const ENABLE_SURVIVOR_POOL = false

	'Enable margin pool.
	const ENABLE_MARGIN_POOL = false

	'Cost of a survivor/margin pool entry.
	const SIDE_BET_AMOUNT = 10

	'Starting week of the side pool.
	const SIDE_START_WEEK = 1

	'Amount of pot that goes to the survivor pool winner(s).
	'(Applies only when a combined survivor/margin pool is used, the winner(s)
	'of the margin pool get the remainder.)
	const SURVIVOR_POT_SHARE = .75

	'Number of losing picks that will eliminate a player.
	const SURVIVOR_STRIKE_OUT = 1

	'Tie game counts as a losing survivor pick flag.
	const SURVIVOR_STRIKE_ON_TIE = true

	'Deduct points on a missing margin pick flag.
	const MARGIN_DEDUCT_LOSS = true

	'--------------------------------------------------------------------------
	' Message board settings.
	'--------------------------------------------------------------------------

	'Enable message board flag.
	const ENABLE_MESSAGE_BOARD = true

	'Define maximum message length allowed.
	const MAX_MESSAGE_LENGTH = 2000

	'Define the number of days to keep posts. Set to zero to keep posts
	'indefinitely.
	const MAX_POST_AGE = 0

	'Define the maximum number of posts to keep (when this limit is exceeded,
	'the oldest posts will be removed first). Set to zero to keep an unlimited
	'number of posts.
	const MAX_POST_COUNT = 0

	'Number of posts displayed per page.
	const POST_PAGE_SIZE = 10

	'**************************************************************************
	'* End of configuration section.                                          *
	'**************************************************************************

	'**************************************************************************
	'* Constant definitions.                                                  *
	'**************************************************************************

	'Username reserved for the Administrator.
	const ADMIN_USERNAME = "Admin"

	'Username max length.
	const USERNAME_MAX_LEN = 25

	'Minimum password length.
	const PASSWORD_MIN_LEN = 5

	'Value used to denote a tie game in the database.
	const TIE_STR = "Tie"

	'Half points display. Used when tie picks are not allowed to denote that a
	'pick is worth half the normal points.
	const HALF_POINTS = "*"

	'For checkboxes.
	const CHECKED_ATTRIBUTE = " checked=""checked"""

	'**************************************************************************
	'* Global variables.                                                      *
	'**************************************************************************

	'Administrator-only flag. Defaults to false.
	dim AdminOnly
	AdminOnly = false

	'Page sub title.
	dim PageSubTitle

	'Database connection.
	dim DbConn

	'Conference and division names (should actually be constants).
	dim ConferenceNames, DivisionNames
	ConferenceNames = Array("AFC", "NFC")
	DivisionNames   = Array("East", "North", "South", "West")

	'Title text for the survivor/margin side pool.
	dim SidePoolTitle
	SidePoolTitle = ""
	if ENABLE_SURVIVOR_POOL and ENABLE_MARGIN_POOL then
		SidePoolTitle = "Survivor/Margin"
	elseif ENABLE_SURVIVOR_POOL then
		SidePoolTitle = "Survivor"
	elseif ENABLE_MARGIN_POOL then
		SidePoolTitle = "Margin"
	end if
	
	'Classes for handling the breaking of the tiebreaker into two fields
	Class TBclass
		Public VisScore
		Public HomeScore
		Public Dif
	End Class

	'**************************************************************************
	'* Functions and subroutines.                                             *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Returns true if all games for the specified week are locked.
	'--------------------------------------------------------------------------
	function AllGamesLocked(week)

		dim dateTimeNow, rs, gameDatetime
		
		'now need to update this to be thurs game locks on kickoff, others lock on fri 5pm

		AllGamesLocked = false
		dateTimeNow = CurrentDateTime()
		
		
		set rs = WeeklySchedule(week)
		
		do while not rs.EOF
			gameDatetime = CDate(rs.Fields("Date").Value & " " & rs.Fields("Time").Value)
			'call errormessage(gameDatetime & dateTimeNow & week)
			'old method
			'gameDatetime = CDate(rs.Fields("Date").Value & " " & rs.Fields("Time").Value)
			'if Weekday(gameDatetime) = vbSunday and gameDatetime <= dateTimeNow then
			'	AllGamesLocked = true
			'	exit function
			'end if
			
			'modding for the bowl series. this needs to be updated so that it will do one way for regular season and another for the bowls as stretch over 3 weeks
			'EG(2013-08-24) - updating to disable this as we wont use an all games locked. all will lock at kickoff'
			if (1 <> 1) then 'if (week <> 15) then
				
			
			'so we are going to use datePart(ww, date) to use the week.. ww is the week of the year
			'brian wants thursday games locked on kickoff and rest on saturday at 12
				if datePart("ww", gameDatetime) = datePart("ww", dateTimeNow) then
					'same week
					if (weekday(dateTimeNow) = vbSaturday and datePart("h", dateTimeNow) > 11 ) then
						'this means saturday and after 12pm..
						call errormessage("This Week's Games have been locked" & dateTimeNow & datePart("h", dateTimeNow) & weekday(dateTimeNow))
						AllGamesLocked = true
						exit function
					elseif (weekday(dateTimeNow) > vbSaturday	) then
						'this means day is after saturday
						call errormessage("in past friday")
						AllGamesLocked = true
						exit function
					end if
				elseif datePart("yyyy", gameDatetime) < datePart("yyyy", dateTimeNow) then
					'later week
						call errormessage("This Week's Games were locked last year")
						AllGamesLocked = true
						exit function	
				elseif datePart("ww", gameDatetime) < datePart("ww", dateTimeNow) then
					'later week.. do we assume games will always be in same week???
						call errormessage("This Week's Games have been locked - past week")
						AllGamesLocked = true
						exit function
				
				'if Weekday(gameDatetime) = vbSunday and gameDatetime <= dateTimeNow then
				'	AllGamesLocked = true
				'	exit function
				else 
					'call errormessage("early week")
						'AllGamesLocked = true
						exit function
				end if
			end if
			rs.MoveNext
		loop

	end function
	
	
	'this finds if the TBgame for the week is locked to see if can still display or locked
	function TBgamelocked(week)

		dim sql, rs
		
		TBgamelocked = False
		sql = "SELECT Date, Time" _
	       & " FROM Schedule" _
		   & " WHERE Week = " & week _
		   & " AND TBgame = 1"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			'if not IsNumeric(rs.Fields("GameTime").Value) then
			'	exit function
			'end if
			if CDate(rs.Fields("Date").Value & " " & rs.Fields("Time").Value) < CurrentDateTime() then
				TBgamelocked = True
			end if
		end if
		'call errormessage("Hi")
	end function
	

	'--------------------------------------------------------------------------
	' Returns a date object representing the current date and time adjusted to
	' the time zone used in the database.
	'--------------------------------------------------------------------------
	function CurrentDateTime()

		CurrentDateTime = DateAdd("h", -TIME_ZONE_DIFF, CDate(Date() & " " & Time()))

	end function

	'--------------------------------------------------------------------------
	' Returns the current week number based on the current (time zone adjusted)
	' date and time.
	'--------------------------------------------------------------------------
	function CurrentWeek()

		dim dateNow, sql, rs, found

		CurrentWeek = 0
		dateNow = CurrentDateTime
		sql = "SELECT Week, Date FROM Schedule ORDER BY Date"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			found = false
			do while not rs.EOF and not found
				if DateDiff("d", DateValue(rs.Fields("Date").Value), dateNow) <= 0 then
					found = true
					CurrentWeek = rs.Fields("Week").Value
				end if
				rs.MoveNext
			loop
		end if

		if CurrentWeek = 0 then
			CurrentWeek = NumberOfWeeks()
		end if

	end function
	
	


	'--------------------------------------------------------------------------
	' Displays the specified error message.
	'--------------------------------------------------------------------------
	sub ErrorMessage(msg)

		Response.Write(vbTab & "<!-- Error message. -->" & vbCrLf _
		             & vbTab & "<p class=""error"">" & msg & "</p>" & vbCrLf)

	end sub

	'--------------------------------------------------------------------------
	' Returns true if the given game start time has passed, based on the
	' current date and time. Used to determine if an early game should be
	' locked.
	'--------------------------------------------------------------------------
	function GameStarted(gameDatetime)

		GameStarted = false
		if gameDatetime < CurrentDateTime() then
			GameStarted = true
		end if

	end function

	'--------------------------------------------------------------------------
	' Gets the week number specified in the HTTP request. If no valid number
	' was specified, the current week is returned instead.
	'--------------------------------------------------------------------------
	function GetRequestedWeek()

		dim week, sql

		'Default to the current week.
		GetRequestedWeek = CurrentWeek()

		'Check the week number specified in the HTTP request.
		week = Request("week")
		if not IsNumeric(week) then
			exit function
		else
			week = Round(week)
			if week < 1 or week > NumberOfWeeks() then
				exit function
			end if
		end if
		GetRequestedWeek = week

	end function
	


	'--------------------------------------------------------------------------
	' Displays the specified info message.
	'--------------------------------------------------------------------------
	sub InfoMessage(msg)

		Response.Write(vbTab & "<!-- Info message. -->" & vbCrLf _
		             & vbTab & "<p class=""info"">" & msg & "</p>")

	end sub

	'--------------------------------------------------------------------------
	' Returns true if the current user is the Administrator.
	'--------------------------------------------------------------------------
	function IsAdmin()

		IsAdmin = false
		if Session(SESSION_USERNAME_KEY) = ADMIN_USERNAME then
			IsAdmin = true
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns true if the current user is disabled.
	'--------------------------------------------------------------------------
	function IsDisabled()

		dim sql, rs

		IsDisabled = false
		sql = "SELECT DisableEntries FROM Users" _
		   & " WHERE Username = '" & SqlString(Session(SESSION_USERNAME_KEY)) & "'"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			IsDisabled = rs.Fields("DisableEntries").Value
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns true if the given number is a valid point spread.
	'--------------------------------------------------------------------------
	function IsValidPointSpread(n)

		IsValidPointSpread = false

		if IsNumeric(n) then
			if 2 * n = Round(2 * n) then
				IsValidPointSpread = true
			end if
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the number of completed games (i.e., games with a result)
	' for the specified week.
	'--------------------------------------------------------------------------
	function NumberOfCompletedGames(week)

		dim sql, rs

		NumberOfCompletedGames = 0
		sql = "SELECT COUNT(*) AS Total" _
		   & " FROM Schedule WHERE Week = " & week _
		   & " AND NOT ISNULL(Result)"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			NumberOfCompletedGames = rs.Fields("Total").Value
		end if

	end function

    '--------------------------------------------------------------------------
    ' Returns the number of completed games (i.e., games with a result)
    ' for the specified week.
    '--------------------------------------------------------------------------
    function NumberOfTotalCompletedGames()

        dim sql, rs

        NumberOfTotalCompletedGames = 0
        sql = "SELECT COUNT(*) AS Total" _
           & " FROM Schedule WHERE" _
           & " NOT ISNULL(Result)"
        set rs = DbConn.Execute(sql)
        if not (rs.BOF and rs.EOF) then
            NumberOfTotalCompletedGames = rs.Fields("Total").Value
        end if

    end function

	'--------------------------------------------------------------------------
	' Returns the number of games scheduled for the specified week.
	'--------------------------------------------------------------------------
	function NumberOfGames(week)

		dim sql, rs

		NumberOfGames = 0
		sql = "SELECT COUNT(*) AS Total FROM Schedule WHERE Week = " & week
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			NumberOfGames = rs.Fields("Total").Value
		end if

	end function
	
	
	
	
	
	

	'---------------------------------------------------------------------------
	' Returns the total number of teams in the database.
	'---------------------------------------------------------------------------
	function NumberOfTeams()

		dim sql, rs

		NumberOfTeams = 0
		sql = "SELECT COUNT(TeamID) AS Total FROM Teams"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			NumberOfTeams = rs.Fields("Total").Value
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the total number of weeks in the schedule.
	'--------------------------------------------------------------------------
	function NumberOfWeeks()

		dim sql, rs

		NumberOfWeeks = 0
		sql = "SELECT MAX(Week) AS Total FROM Schedule"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			NumberOfWeeks = rs.Fields("Total").Value
		end if

	end function
	


	'--------------------------------------------------------------------------
	' Opens the database connection (global variable 'DbConn').
	'--------------------------------------------------------------------------
	sub OpenDB()

        dim dbDir, connectstr

		set DbConn = Server.CreateObject("ADODB.Connection")
		dbDir = Server.MapPath(DATABASE_FILE_PATH)

		'Use one of the following, depending on what database drivers your host
		'has installed.

		connectstr = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & dbDir
		DbConn.Open connectstr
		'DbConn.Open "Data Source=" & dbDir & ";Provider=Microsoft.Jet.OLEDB.4.0;"
		'DbConn.Open "DBQ="& dbDir & ";Driver={Microsoft Access Driver (*.mdb)}"

		
		'start
		'Dim oConn, oRs
		'Dim qry, connectstr
		'Dim db_path
		'Dim db_dir
		'db_dir = Server.MapPath("access_db") & "\"
		'db_path = db_dir & "NFL.mdb"
		'fieldname = "your_field"
		'tablename = "your_table"

		' connectstr = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & db_path
' 
' 		Set oConn = Server.CreateObject("ADODB.Connection")
' 		oConn.Open connectstr
' 		qry = "SELECT * FROM " & tablename
' 
' 		Set oRS = oConn.Execute(qry)
' 
' 		if not oRS.EOF then
' 		while not oRS.EOF
' 		response.write ucase(fieldname) & ": " & oRs.Fields(fieldname) & "
' 		"
' 		oRS.movenext
' 		wend
' 		oRS.close
' 		end if
' 
' 		Set oRs = nothing
' 		Set oConn = nothing

	end sub

	'--------------------------------------------------------------------------
	' Set the results for the given game based on the game score and point
	' spread.
	'--------------------------------------------------------------------------
	sub SetGameResults(id)

		dim sql, rs, vid, hid, vscore, hscore, spread, result, atsResult

		sql = "SELECT * FROM Schedule WHERE GameID = " & id
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
			sql = "UPDATE Schedule SET" _
			   & " Result       = " & result & ", " _
			   & " ATSResult    = " & atsResult _
			   & " WHERE GameID = " & id
			call DbConn.Execute(sql)
		end if

	end sub

	'--------------------------------------------------------------------------
	' Returns a string with single quotes escaped. Used for variable string
	' values in SQL statements.
	'--------------------------------------------------------------------------
	function SqlString(str)

		SqlString = Replace(str, "'", "''")

	end function

	'--------------------------------------------------------------------------
	' Returns the total number of credits for the specified user.
	'--------------------------------------------------------------------------
	function TotalCredits(username)

		dim sql, rs

		TotalCredits = 0
		sql = "SELECT SUM(Amount) AS Total" _
 		    & " FROM Credits" _
 		    & " WHERE Username = '" & SqlString(username) & "'"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			TotalCredits = rs.Fields("Total")
			if not IsNumeric(TotalCredits) then
				TotalCredits = 0
			end if
		end if

	end function

	'--------------------------------------------------------------------------
	' Checks a username for invalid characters and returns the first one found.
	' If no such characters are found, and empty string is returned.
	'--------------------------------------------------------------------------
	function UsernameCheck(username)

		dim encoded, i

		UsernameCheck = ""

		'Disallow commas.
		if InStr(username, ",") then
			UsernameCheck = ","
			exit function
		end if

		'Disallow any HTML reserved characters.
		if InStr(username, "&") then
			UsernameCheck = "&"
			exit function
		end if
		encoded = Server.HtmlEncode(username)
		i = 1
		do while i <= Len(username) and i <= Len(encoded) and UsernameCheck = ""
			if Mid(username, i, 1) <> Mid(encoded, i, 1) then
				UsernameCheck = Mid(username, i, 1)
			end if
			i = i + 1
		loop

	end function

	'--------------------------------------------------------------------------
	' Returns an array of all usernames in the database, optionally excluding
	' the administrator username.
	'--------------------------------------------------------------------------
	function UsersList(excludeAdmin)

		dim list, sql, rs

		UsersList = ""

		sql = "SELECT Username FROM Users"
		if excludeAdmin then
			sql = sql & " WHERE Username <> '" & SqlString(ADMIN_USERNAME) & "'"
		end if
		sql = sql & " ORDER BY Username"
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
		UsersList = list

	end function

	'--------------------------------------------------------------------------
	' Returns a record set of games for the specified week.
	'--------------------------------------------------------------------------
	function WeeklySchedule(week)

		dim sql, sql2

		'Retrieve the week's schedule from the database.
		sql = "SELECT Schedule.*," _
		   & " VTeams.City AS VCity, VTeams.Name AS VName, VTeams.DisplayName AS VDisplayName, VTeams.Logo AS logoVis, " _
		   & " HTeams.City AS HCity, HTeams.Name AS HName, HTeams.DisplayName AS HDisplayName, HTeams.Logo AS logoHome" _
		   & " FROM Schedule, Teams AS VTEAMS, Teams AS HTEAMS" _
		   & " WHERE Schedule.VisitorID = VTeams.TeamID and Schedule.HomeID = HTeams.TeamID" _
		   & " AND Week = " & week _
		   & " ORDER BY Date, Time, VTeams.City, VTeams.Name"

        sql2 =  "SELECT Schedule.*, FullSchedule.*," _
            & " VTeams.City AS VCity, VTeams.Name AS VName, VTeams.DisplayName AS VDisplayName, VTeams.Logo AS logoVis," _
            & " HTeams.City AS HCity, HTeams.Name AS HName, HTeams.DisplayName AS HDisplayName, HTeams.Logo AS logoHome" _
            & " FROM Schedule JOIN FullSchedule ON Schedule.FullScheduleID = FullSchedule.ID" _
            & " LEFT JOIN Teams AS VTeams ON Schedule.VisitorID = VTeams.TeamID" _
            & " LEFT JOIN Teams AS HTeams ON Schedule.HomeID = HTeams.TeamID" _
            & " WHERE Schedule.Week = " & week _
            & " ORDER BY Schedule.Date, Schedule.Time, VTeams.City, VTeams.Name"

		set WeeklySchedule = DbConn.Execute(sql)

	end function
	


	'**************************************************************************
	'* Formatting functions.                                                  *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Returns a formatted date string (month and day) given a Date object.
	'--------------------------------------------------------------------------
	function FormatDate(dt)

		FormatDate = Month(dt) & "/" & Day(dt)

	end function

	'--------------------------------------------------------------------------
	' Returns a formatted time string given a Date object.
	'--------------------------------------------------------------------------
	function FormatTime(dt)

		dim hh, mm, symbol

		hh = Hour(dt)
		mm = Minute(dt)
		if mm < 10 then
			mm = "0" & mm
		end if
		if hh >= 12 then
			symbol = "pm"
		else
			symbol = "am"
		end if
		if hh >= 13 then
			hh = hh - 12
		end if

		FormatTime = hh & ":" & mm & " " & symbol

	end function

	'--------------------------------------------------------------------------
	' Returns a formatted date string (month, day and year) given a Date
	' object.
	'--------------------------------------------------------------------------
	function FormatFullDate(dt)

		FormatFullDate = Month(dt) & "/" & Day(dt) & "/" & Year(dt)

	end function

	'--------------------------------------------------------------------------
	' Returns a formatted time string (hour, minute, second and meridiem) given
	' a Date object.
	'--------------------------------------------------------------------------
	function FormatFullTime(dt)

		FormatFullTime = LCase(FormatDateTime(dt, vbLongTime))
		'FormatFullTime = Format(dt, "hh:mm AMPM")

	end function

	'--------------------------------------------------------------------------
	' Returns a formatted dollar amount string based on the given number.
	'--------------------------------------------------------------------------
	function FormatAmount(n)

		FormatAmount = "&nbsp;"
		if IsNumeric(n) then
			FormatAmount = FormatCurrency(n, 2, true, false, true)
		end if

   end function

	'--------------------------------------------------------------------------
	' Returns a formatted percent string based on the given number.
	'--------------------------------------------------------------------------
	function FormatPercentage(n)

		FormatPercentage = "&nbsp;"
		if IsNumeric(n) then
			FormatPercentage = FormatNumber(Round(n, 3), 3, true)
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns a formatted point spread.
	'--------------------------------------------------------------------------
	function FormatPointSpread(spread)

		dim n

		FormatPointSpread = "n/l&nbsp;&nbsp;&nbsp;"
		if IsNumeric(spread) then
			if spread = 0 then
				FormatPointSpread = "even&nbsp;&nbsp;&nbsp;"
			else
				n = Fix(spread)
				if n = 0 then
					FormatPointSpread = "&nbsp;"
				else
					FormatPointSpread = n
				end if
				if Abs(spread) - Abs(n) = 0.5 then
					FormatPointSpread = FormatPointSpread & "&frac12;"
				else
					FormatPointSpread = FormatPointSpread & "&nbsp;&nbsp;&nbsp;"
				end if
				if spread > 0 then
					FormatPointSpread = "+" & FormatPointSpread
				end if
			end if
		end if

	end function

	'--------------------------------------------------------------------------
	' Formats a win-loss-tie record for output. Ties are not displayed if they
	' are equal to zero.
	'--------------------------------------------------------------------------
	function FormatRecord(w, l, t)

		FormatRecord = w & "-" & l
		if t > 0 then
			FormatRecord = FormatRecord & "-" & t
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns a formatted score, optionally adding padding to allow alignment
	' when half points occur.
	'--------------------------------------------------------------------------
	function FormatScore(score, addPadding)

		dim n

		FormatScore = "n/a"
		if IsNumeric(score) then
			n = Fix(score)
			if Abs(score) - Abs(n) = .5 then
				FormatScore = n & "&frac12;"
			else
				FormatScore = n
				if addPadding then
					FormatScore = FormatScore & "&nbsp;&nbsp;&nbsp;"
				end if
			end if
		else
			if addPadding then
				FormatScore = FormatScore & "&nbsp;&nbsp;&nbsp;"
			end if
		end if

	end function

	'--------------------------------------------------------------------------
	' Formats a string to indicate a correct pick.
	'--------------------------------------------------------------------------
	function FormatCorrectPick(str)

		if USE_POINT_SPREADS then
			FormatCorrectPick = FormatATSWinner(str)
		else
			FormatCorrectPick = FormatWinner(str)
		end if

	end function

	'--------------------------------------------------------------------------
	' Formats a string to indicate the game winner.
	'--------------------------------------------------------------------------
	function FormatWinner(str)

		if USE_POINT_SPREADS then
			FormatWinner = "<em>" & str & "</em>"
		else
			FormatWinner = "<strong>" & str & "</strong>"
		end if

	end function

	'--------------------------------------------------------------------------
	' Formats a string to highlight the game winner against the point spread.
	'--------------------------------------------------------------------------
	function FormatATSWinner(str)

		if USE_POINT_SPREADS then
			FormatATSWinner = "<strong>" & str & "</strong>"
		else
			FormatATSWinner = str
		end if

	end function
	
	

	
	
	'--------------------------
	'this function is to replicate the format function from vb
	'--------------------------
	
	Function Format(vExpression, sFormat) 
 
 		Dim fmt
        set fmt = CreateObject("MSSTDFMT.StdDataFormat") 
        fmt.Format = sFormat 
 
        set rs = CreateObject("ADODB.Recordset") 
        rs.Fields.Append "fldExpression", 12 ' adVariant 
 
        rs.Open 
        rs.AddNew 
 
        set rs("fldExpression").DataFormat = fmt 
        rs("fldExpression").Value = vExpression 
 
        Format = rs("fldExpression").Value 
 
        rs.close: Set rs = Nothing: Set fmt = Nothing 
 
    End Function
	
	
	
	

	
	
	
	
	
	Function listUpdateSchedule(week)
	
	
	
	end Function
	

	'**************************************************************************
	'* Common displays.                                                       *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Builds a drop down list for confidence points given number of games and
	' the given point value selected.
	'--------------------------------------------------------------------------
	sub DisplayConfidencePointsList(nTabs, numGames, selectedPoints)

		dim str, i

		str = String(nTabs, vbTab) & "<option value=""""></option>" & vbCrLf
		for i = 1 to numGames
			str = str & String(nTabs, vbTab) & "<option value=""" & i & """"
			if IsNumeric(selectedPoints) then
				if i = CInt(selectedPoints) then
					str = str & " selected=""selected"""
				end if
			end if
			str = str & ">"

			'Need to pad for browsers that don't support right-aligned text
			'in select options. Also, pad on right to give spacing between
			'text and select button.
			if i < 10 then
				str = str & "&nbsp;&nbsp;"
			end if
			str = str & i & "</option>" & vbCrLf
		next
		Response.Write(str)

	end sub

	'--------------------------------------------------------------------------
	' Displays a list of teams with open dates for the specified week.
	'--------------------------------------------------------------------------
	sub DisplayOpenDates(nTabs, week)

		dim str, count, sql, rs, rs2, id, team

		str = ""
		count = 0

		'Build the list of team with open dates.
		sql = "SELECT * FROM Teams ORDER BY City, Name"
		set rs = DbConn.Execute(sql)
		do while not rs.EOF
			id = rs.Fields("TeamID").Value
			sql = "SELECT COUNT(*) as Total" _
			   & " FROM Schedule" _
			   & " WHERE (HomeID = '" & id & "'" _
			   & " OR VisitorID = '" & id & "')" _
			   & " AND Week = " & week
			set rs2 = DbConn.Execute(sql)
			if rs2.Fields("Total").Value = 0 then

				'Get the team name and add it to the list.
				team = rs.Fields("City").Value
				if rs.Fields("DisplayName").Value <> "" then
					team = rs.Fields("DisplayName").Value
				end if
				if str <> "" then
					str = str & ", "
				end if

				'Add a line break after the fifth team name.
				count = count + 1
				if count > 5 then
					str = str & "<br />"
					count = 0
				end if

				str = str & team
			end if
			rs.MoveNext
		loop

		'Display the list.
		if Len(str) > 0 then
			Response.Write(String(nTabs, vbTab) & "<p><strong>Open dates:</strong> " & str & ".</p>")
		end if

	end sub

	'--------------------------------------------------------------------------
	' Displays a list of links for viewing different weeks. The URLs point to
	' the calling script, passing a "week" parameter in the query string.
	' Additional parameters may also be specified.
	'--------------------------------------------------------------------------
	sub DisplayWeekNavigation(nTabs, otherParams)

		dim str, i

		str = ""
		for i = 1 to NumberOfWeeks()
			if i > 1 then
				str = str & " &middot; "
			end if
			str = str & "<a href=""" _
			   & Request.ServerVariables("SCRIPT_NAME") _
			   & "?week=" & i
			if otherParams <> "" then
				str = str & "&amp;" & otherParams
			end if
			str = str & """>" & i & "</a>"
		next

		if Len(str) > 0 then
			Response.Write(String(nTabs, vbTab) & "<p><strong>Go to week:</strong> " & str & "</p>")
		end if

	end sub 
	
	
	'--------------------------------------------------------------------------
	' Displays a list of links for viewing different weeks. 
	' 
	' this one is for the fullschedule
	'--------------------------------------------------------------------------
	sub DisplayWeekNavigationFS(nTabs, otherParams)

		dim str, i

		str = ""
		for i = 1 to NumberOfFSWeeks()
			if i > 1 then
				str = str & " &middot; "
			end if
			str = str & "<a href=""" _
			   & Request.ServerVariables("SCRIPT_NAME") _
			   & "?week=" & i
			if otherParams <> "" then
				str = str & "&amp;" & otherParams
			end if
			str = str & """>" & i & "</a>"
		next

		if Len(str) > 0 then
			Response.Write(String(nTabs, vbTab) & "<p><strong>Go to week:</strong> " & str & "</p>")
		end if

	end sub 
	
	%>
    
    