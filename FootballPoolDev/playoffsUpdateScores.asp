<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Enter Playoffs Scores" : AdminOnly = true %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/common.css" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/email.asp" -->
<!-- #include file="includes/encryption.asp" -->
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/playoffs.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'If there is form data, process it.
	dim n, i
	dim gameID, vscore, hscore, ot, result
	dim notify, infoMsg
	dim sql, rs
	n = NumberOfPlayoffGames()
	if Request.ServerVariables("Content_Length") > 0 and not CancelRequested() then
		for i = 1 to n

			'Get the form fields.
			gameID = Trim(Request.Form("id-"     & i))
			vscore = Trim(Request.Form("vscore-" & i))
			hscore = Trim(Request.Form("hscore-" & i))
			ot     = Trim(Request.Form("ot-"     & i))

			'Validate the form fields.
			if vscore <> "" or hscore <> "" or ot <> "0" then
				if not IsNumeric(vscore) then
					FormFieldErrors.Add "vscore-" & i, "'" & vscore & "' is not a valid game score."
				else
					if CInt(vscore) < 0 or CInt(vscore) <> CDbl(vscore) then
						FormFieldErrors.Add "vscore-" & i, "'" & vscore & "' is not a valid game score."
					end if
				end if
				if not IsNumeric(hscore) then
					FormFieldErrors.Add "hscore-" & i, "'" & hscore & "' is not a valid game score."
				else
					if CInt(hscore) < 0 or CInt(hscore) <> CDbl(hscore) then
						FormFieldErrors.Add "hscore-" & i, "'" & hscore & "' is not a valid game score."
					end if
				end if
			end if
		next

		'If there were any errors, display the error summary message.
		'Otherwise, do the updates.
		if FormFieldErrors.Count > 0 then
			call FormFieldErrorsMessage("Error: Invalid fields. Please correct and resubmit.")
		else
			for i = 1 to n
				gameID = Trim(Request.Form("id-"    & i))
				vscore = Trim(Request.Form("vscore-" & i))
				hscore = Trim(Request.Form("hscore-" & i))
				ot     = Trim(Request.Form("ot-"     & i))

				'Update the scores.
				if vscore = "" then
					vscore ="NULL"
				end if
				if hscore = "" then
					hscore ="NULL"
				end if
				sql = "UPDATE PlayoffsSchedule SET" _
				   & " VisitorScore = " & vscore & "," _
				   & " HomeScore    = " & hscore & "," _
				   & " OT           = " & ot _
				   & " WHERE GameID = " & gameID
				call DbConn.Execute(sql)

				'Update the results.
				call PlayoffsSetGameResults(gameID)
			next

			'Update the Super Bowl teams, if appropriate.
			call SetSuperBowlTeams()

			'Send out email notifications, if checked.
			infoMsg = "Update successful."
			notify = Trim(Request.Form("notify"))
			if LCase(notify) = "true" then
				infoMsg = "Update successful, notifications sent."
				call SendNotifications()
			end if

			'Updates done, show an informational message.
			call InfoMessage(infoMsg)
		end if
	end if

	'Build the display.
	dim cols
	cols = 9
	if USE_POINT_SPREADS then
		cols = cols + 1
	end if %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<table class="main" cellpadding="0" cellspacing="0">
<%	dim lastRound, gameRound, conference
	dim gameDate, gameTime, vid, hid
	dim spread, atsResult
	dim visitor, home, checkedStr
	dim alt
	sql = "SELECT * FROM PlayoffsSchedule ORDER BY Round, Date, Time"
	set rs = DbConn.Execute(sql)
	if not (rs.BOF and rs.EOF) then
		n = 1
		lastRound = 0
		alt = false
		do while not rs.EOF
			gameID    = rs.Fields("GameID").Value
			gameRound = rs.Fields("Round").Value
			gameDate  = rs.Fields("Date").Value
			gameTime  = rs.Fields("Time").Value
			vid       = rs.Fields("VisitorID").Value
			vscore    = rs.Fields("VisitorScore").Value
			spread    = rs.Fields("PointSpread").Value
			hid       = rs.Fields("HomeID").Value
			hscore    = rs.Fields("HomeScore").Value
			ot        = rs.Fields("OT").Value
			result    = rs.Fields("Result").Value
			atsResult = rs.Fields("ATSResult").Value

			'Get the conference of the teams playing.
			if gameRound <> NumberOfPlayoffRounds() and not IsNull(hid) then
				conference = "(" & ConferenceNames(GetConference(hid) - 1) & ")"
			else
				conference = "&nbsp;"
			end if

			if gameRound <> lastRound then
				lastRound = gameRound
				alt = false
				if gameRound = 1 then %>
			<tr class="header bottomEdge">
<%				else %>
			<tr class="header topEdge bottomEdge">
<%				end if %>
				<th align="left" colspan="<% = cols %>"><% = PlayoffRoundNames(gameRound - 1) %></th>
			</tr>
<%			end if

			'Set the team names for display.
			visitor = GetTeamName(rs.Fields("VisitorID").Value)
			home    = GetTeamName(rs.Fields("HomeID").Value)

			'Highlight the results.
			if result = vid then
				visitor = FormatWinner(visitor)
			elseif result = hid then
				home = FormatWinner(home)
			end if
			if atsResult = vid then
				visitor = FormatATSWinner(visitor)
			elseif atsResult = hid then
				home = FormatATSWinner(home)
			end if


			if alt then %>
			<tr align="right" class="alt">
<%			else %>
			<tr align="right">
<%			end if
			alt = not alt

			'If there were errors on the form post processing, restore those fields.
			if FormFieldErrors.Count > 0 then
				vscore = GetFieldValue("vscore-" & n, vscore)
				hscore = GetFieldValue("hscore-" & n, hscore)
				ot     = GetFieldValue("ot-"     & n, ot)
			end if %>
				<td><input type="hidden" name="id-<% = n %>" value="<% = gameID %>" /><input type="hidden" name="round-<% = n %>" value="<% = gameRound %>" /><% = WeekdayName(Weekday(gameDate), true) %></td>
				<td><% = FormatDate(gameDate) %></td>
				<td><% = FormatTime(gameTime) %></td>
				<td><% = conference %></td>
				<td><% = visitor %></td>
<%			if USE_POINT_SPREADS then %>
				<td><% = FormatPointSpread(spread) %></td>
<% 			end if %>
				<td><input type="text" name="vscore-<% = n %>" value="<% = vscore %>" size="2" class="<% = FieldStyleClass("numeric", "vscore-" & n) %>" /></td>
				<td><% = GetConjunction(gameRound) & " " & home %></td>
				<td><input type="text" name="hscore-<% = n %>" value="<% = hscore %>" size="2" class="<% = FieldStyleClass("numeric", "hscore-" & n) %>" /></td>
				<td>
					<select name="ot-<% = n %>">
						<option value="0"<% if rs.Fields("OT") = 0 then Response.Write(" selected=""selected""") end if %>></option>
						<option value="1"<% if rs.Fields("OT") = 1 then Response.Write(" selected=""selected""") end if %>>OT</option>
						<option value="2"<% if rs.Fields("OT") = 2 then Response.Write(" selected=""selected""") end if %>>OT(2)</option>
						<option value="3"<% if rs.Fields("OT") = 3 then Response.Write(" selected=""selected""") end if %>>OT(3)</option>
						<option value="4"<% if rs.Fields("OT") = 4 then Response.Write(" selected=""selected""") end if %>>OT(4)</option>
					</select>
				</td>
			</tr>
<%			rs.MoveNext
			n = n + 1
		loop
		if SERVER_EMAIL_ENABLED then %>
			<tr class="subHeader topEdge">
				<th align="left" colspan="<% = cols %>"><input type="checkbox" id="notify" name="notify" value="true" /> <label for="notify">Send update notification to users.</label></th>
			</tr>
<%		end if %>
		</table>
<%	end if %>
		<p><input type="submit" name="submit" value="Update" class="button" title="Apply changes." />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel the update." /></p>
	</form>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>
<%	'**************************************************************************
	'* Local functions and subroutines.                                       *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Updates the teams in the Super Bowl if the AFC and NFC Conference
	' Champions have been determined.
	'--------------------------------------------------------------------------
	sub SetSuperBowlTeams()

		dim sql, rs
		dim finalRound, finalID
		dim afcTeam, nfcTeam
		dim result, conference

		'Get the winners of the conference championship games.
		finalRound = NumberOfPlayoffRounds()
		afcTeam = "NULL"
		nfcTeam = "NULL"
		sql = "SELECT * FROM PlayoffsSchedule" _
		   & " WHERE Round = " & finalRound - 1 _
		   & " AND Result <> NULL"
		set rs = DbConn.Execute(sql)
		if (rs.BOF and rs.EOF) then
			exit sub
		end if
		do while not rs.EOF
			result = rs.Fields("Result").Value
			if not IsNull(result) then
				conference = GetConference(result)
				if conference = 1 then
					afcTeam = "'" & result & "'"
				elseif conference = 2 then
					nfcTeam = "'" & result & "'"
				end if
			end if
			rs.MoveNext
		loop

		'Find the id of the final game.
		sql = "SELECT GameID FROM PlayoffsSchedule" _
		   & " WHERE Round = " & finalRound
		set rs = DbConn.Execute(sql)
		if rs.EOF and rs.BOF then
			exit sub
		end if
		finalID = rs.Fields("GameID").Value

		'Update the team IDs.
		sql = "UPDATE PlayoffsSchedule SET" _
		   & " VisitorID = " & afcTeam & "," _
		   & " HomeID    = " & nfcTeam _
		   & " WHERE GameID = " & finalID
		call DbConn.Execute(sql)

	end sub

	'--------------------------------------------------------------------------
	' Sends an email notification to any users who have elected to receive
	' them.
	'--------------------------------------------------------------------------
	sub SendNotifications()

		dim subj, body
		dim sql, rs
		dim gameRound, lastRound
		dim vid, hid, vscore, hscore, ot
		dim list, email

		subj = "Football Pool Game Results Notification"
		body = "Results for playoff games have been posted." & vbCrLf

		'Show results.
		sql = "SELECT * FROM PlayoffsSchedule ORDER BY Round, Date, Time"
		set rs = DbConn.Execute(sql)
		lastRound = 0
		do while not rs.EOF
			gameRound = rs.Fields("Round").Value
			if gameRound <> lastRound then
				body = body & vbCrLf & PlayoffRoundNames(gameRound - 1) & vbCrLf
				lastRound = gameRound
			end if
			vid = rs.Fields("VisitorID").Value
			hid = rs.Fields("HomeID").Value
			if IsNull(vid) then
				vid = "[TBD]"
			end if
			if IsNull(hid) then
				hid = "[TBD]"
			end if
			vscore = ""
			hscore = ""
			if not IsNull(rs.Fields("Result").Value) then
				vscore = " " & rs.Fields("VisitorScore").Value
				hscore = " " & rs.Fields("HomeScore").Value
			end if
			ot = rs.Fields("OT").Value
			body = body _
			     & vid & vscore
			if USE_POINT_SPREADS then
				body = body & " (" & PlainTextPointSpread(rs.Fields("PointSpread").Value) & ")"
			end if
			body = body & " @ " & hid & hscore
				if ot > 0 then
					body = body & " OT"
					if ot > 1 then
						body = body & "(" & ot & ")"
					end if
				end if
				body = body & vbCrLf
			rs.MoveNext
		loop
		body = body & vbCrLf & "(All times Eastern)" & vbCrLf


		list = GetNotificationList("NotifyOfResultUpdates")
		for each email in list
			call SendMail(email, subj, body)
		next

	end sub %>
