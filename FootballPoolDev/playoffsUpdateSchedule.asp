<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Update Playoffs Schedule" : AdminOnly = true %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/common.css" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
	<link rel="stylesheet" type="text/css" href="styles/datetimePicker.css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
	<script type="text/javascript" src="scripts/datetimePicker.js"></script>
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/datetimePicker.asp" -->
<!-- #include file="includes/email.asp" -->
<!-- #include file="includes/encryption.asp" -->
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/playoffs.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'If there is form data, process it.
	dim n, i
	dim gameID, gameDate, gameTime, visitorID, homeID
	dim notify, infoMsg
	dim sql, rs
	n = NumberOfPlayoffGames()
	if Request.ServerVariables("Content_Length") > 0 and not CancelRequested() then
		for i = 1 to n
			gameID    = Trim(Request.Form("id-"        & i))
			gameDate  = Trim(Request.Form("date-"      & i))
			gameTime  = Trim(Request.Form("time-"      & i))
			visitorID = Trim(Request.Form("visitorID-" & i))
			homeID    = Trim(Request.Form("homeID-"    & i))

			'Validate the form fields.
			if not IsDate(gameDate) then
				FormFieldErrors.Add "date-" & i, "'" & gameDate & "' is not a valid date."
			end if
			if not IsDate(gameTime) then
				FormFieldErrors.Add "time-" & i, "'" & gameTime & "' is not a valid time."
			end if
		next

		'If there were any errors, display the error summary message.
		'Otherwise, do the updates.
		if FormFieldErrors.Count > 0 then
			call FormFieldErrorsMessage("Error: Invalid fields. Please correct and resubmit.")
		else
			for i = 1 to n
				gameID    = Trim(Request.Form("id-"        & i))
				gameDate  = Trim(Request.Form("date-"      & i))
				gameTime  = Trim(Request.Form("time-"      & i))
				visitorID = Trim(Request.Form("visitorID-" & i))
				homeID    = Trim(Request.Form("homeID-"    & i))
				if visitorID = "" then
					visitorID = "NULL"
				else
					visitorID = "'" & visitorID & "'"
				end if
				if homeID = "" then
					homeID = "NULL"
				else
					homeID = "'" & homeID & "'"
				end if
				sql = "UPDATE PlayoffsSchedule SET" _
				   & " [Date]    = #" & gameDate & "#," _
				   & " [Time]    = #" & gameTime & "#,"  _
				   & " VisitorID = " & visitorID & ","  _
				   & " HomeID    = " & homeID _
				   & " WHERE GameID = " & gameID
				call DbConn.Execute(sql)
			next

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

	'Build the display. %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<table class="main" cellpadding="0" cellspacing="0">
<%	'Get teams.
	dim teamsRs
	set teamsRs = Teams()

	'Get games.
	dim lastRound, gameRound
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
			visitorID = rs.Fields("VisitorID").Value
			homeID    = rs.Fields("HomeID").Value

			if gameRound <> lastRound then
				lastRound = gameRound
				alt = false
				if gameRound = 1 then %>
			<tr class="header bottomEdge">
<%				else %>
			<tr class="header topEdge bottomEdge">
<%			end if %>
				<th align="left" colspan="7"><% = PlayoffRoundNames(gameRound - 1) %></th>
			</tr>
<%			end if
			if alt then %>
			<tr class="alt">
<%			else %>
			<tr>
<%			end if
			alt = not alt

			'If there were errors on the form post processing, restore those fields.
			if FormFieldErrors.Count > 0 then
				gameDate  = GetFieldValue("date-"      & n, gameDate)
				gameTime  = GetFieldValue("time-"      & n, gameTime)
				visitorID = GetFieldValue("visitorID-" & n, visitorID)
				homeID    = GetFieldValue("homeID-"    & n, homeID)
			end if %>
				<td align="right"><input type="hidden" name="id-<% = n %>" value="<% = gameID %>" /><input type="text" name="day-<% = n %>" value="<% = WeekdayName(Weekday(gameDate), true) %>" size="3" class="readonly" readonly="readonly" /></td>
				<td><input type="text" name="date-<% = n %>" value="<% = FormatFullDate(gameDate) %>" size="10" class="<% = FieldStyleClass("numeric readonly", "date-" & n) %>" readonly="readonly" /></td>
				<td><input type="image" src="graphics/calendar.gif" onclick="return openCalendar(this.form, 'day-<% = n %>', 'date-<% = n %>');" title="Select a new date." /></td>
				<td><input type="text" name="time-<% = n %>" value="<% = FormatFullTime(gameTime) %>" size="10" class="<% = FieldStyleClass("numeric readonly", "time-" & n) %>" readonly="readonly" /></td>
				<td><input type="image" src="graphics/clock.gif" onclick="return openClock(this.form, 'time-<% = n %>');" title="Select a new time." /></td>
				<td align="right">
					<select name="visitorID-<% = n %>">
						<option value=""></option>
<%			teamsRs.MoveFirst
			do while not teamsRs.EOF %>
						<option value="<% = teamsRs.Fields("TeamID") %>"<% if teamsRs.Fields("TeamID") = visitorID then Response.Write(" selected=""selected""") end if %>><% = teamsRs.Fields("City") & " " & teamsRs.Fields("Name") %></option>
<%				teamsRs.MoveNext
			loop %>
					</select>
				</td>
				<td><% = GetConjunction(gameRound) %>
					<select name="homeID-<% = n %>">
						<option value=""></option>
<%			teamsRs.MoveFirst
			do while not teamsRs.EOF %>
						<option value="<% = teamsRs.Fields("TeamID") %>"<% if teamsRs.Fields("TeamID") = homeID then Response.Write(" selected=""selected""") end if %>><% = teamsRs.Fields("City") & " " & teamsRs.Fields("Name") %></option>
<%				teamsRs.MoveNext
			loop %>
					</select>
				</td>
			</tr>
<%			rs.MoveNext
			n = n + 1
		loop
		if SERVER_EMAIL_ENABLED then %>
			<tr class="subHeader topEdge">
				<th align="left" colspan="7"><input type="checkbox" id="notify" name="notify" value="true" /> <label for="notify">Send update notification to users.</label></th>
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
	' Sends an email notification to any users who have elected to receive
	' them.
	'--------------------------------------------------------------------------
	sub SendNotifications()

		dim subj, body
		dim sql, rs
		dim gameRound, lastRound
		dim vid, hid
		dim list, email

		subj = "Football Pool Game Schedule Update Notification"
		body = "The schedule for playoff games has been updated." & vbCrLf

		'Show the playoffs schedule.
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
			body = body _
			     & WeekdayName(Weekday(rs.Fields("Date").Value), true) _
		    	 & " " & FormatDate(rs.Fields("Date").Value) _
				 & " " & FormatTime(rs.Fields("Time").Value) _
			     & " " & vid & " @ " & hid & vbCrLf
			rs.MoveNext
		loop
		body = body & vbCrLf & "(All times Eastern)" & vbCrLf

		list = GetNotificationList("NotifyOfScheduleUpdates")
		for each email in list
			call SendMail(email, subj, body)
		next

	end sub %>

