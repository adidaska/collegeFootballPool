<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Enter Game Scores" : AdminOnly = true %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/style.css" />
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
<!-- #include file="includes/side.asp" -->
<!-- #include file="includes/weekly.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'Get the week to display.
	dim week
	week = GetRequestedWeek()

	'If there is form data, process it.
	dim n, i
	dim gameID, vscore, hscore, ot
	dim notify, infoMsg
	dim sql, rs
	n = NumberOfGames(week)
	if Request.ServerVariables("Content_Length") > 0 and not CancelRequested() then
		for i = 1 to n

			'Get the form fields.
			gameID = Trim(Request.Form("id-"     & i))
			vscore = Trim(Request.Form("vscore-" & i))
			hscore = Trim(Request.Form("hscore-" & i))
			ot     = Trim(Request.Form("ot-"     & i))

			'Validate the form fields.
			if vscore <> "" or hscore <> "" or ot <> "" then
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
				if LCase(ot) <> "true" then
					ot = false
				end if

				'Update the scores.
				if vscore = "" then
					vscore ="NULL"
				end if
				if hscore = "" then
					hscore ="NULL"
				end if
				sql = "UPDATE Schedule SET" _
				   & " VisitorScore = " & vscore & "," _
				   & " HomeScore    = " & hscore & "," _
				   & " OT           = " & ot _
				   & " WHERE GameID = " & gameID
				call DbConn.Execute(sql)

				'Update the results.
				call SetGameResults(gameID)

			next

			'Clear any cached pool results.
			call ClearWeeklyResultsCache(week)
			if ENABLE_MARGIN_POOL then
				call ClearMarginResultsCache(week)
			end if
			if ENABLE_SURVIVOR_POOL then
				ClearSurvivorStatus(week)
			end if

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

	'Display the schedule for the specified week.
	dim cols
	cols = 8
	if USE_POINT_SPREADS then
		cols = cols + 1
	end if %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<div><input type="hidden" name="week" value="<% = week %>" /></div>
		<table class="main" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge">
			  <th align="left" colspan="<% = cols %>">Week <% = week %></th>
			</tr>
<%	dim gameDate, gameTime, vid, hid
	dim visitor, home, result, checkedStr
	dim spread, atsResult
	dim alt
	set rs = WeeklySchedule(week)
	if not (rs.BOF and rs.EOF) then
		n = 1
		alt = false
		do while not rs.EOF
			gameID    = rs.Fields("GameID").Value
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
			visitor   = rs.Fields("VCity").Value
			home      = rs.Fields("HCity").Value

			'Set the team names for display.
			if rs.Fields("VDisplayName") <> "" then
				visitor = rs.Fields("VDisplayName").Value
			end if
			if rs.Fields("HDisplayName") <> "" then
				home = rs.Fields("HDisplayName")
			end if

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

			'Set the OT checkbox state.
			checkedStr = ""
			if ot then
				checkedStr = CHECKED_ATTRIBUTE
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

				'If the user unchecked the OT checkbox, be sure to leave it unchecked.
				if not FormFieldExists("ot-" & gameID) then
					checkedStr = ""
				end if 
			end if %>
				<td><input type="hidden" name="id-<% = n %>" value="<% = gameID %>" /><input type="hidden" name="vid-<% = n %>" value="<% = vid %>" /><input type="hidden" name="hid-<% = n %>" value="<% = hid %>" /><% = WeekdayName(Weekday(gameDate), true) %></td>
				<td><% = FormatDate(gameDate) %></td>
				<td><% = FormatTime(gameTime) %></td>
				<td><% = visitor %></td>
<%			if USE_POINT_SPREADS then %>
				<td><% = FormatPointSpread(spread) %></td>
<%			end if %>
				<td><input type="text" name="vscore-<% = n %>" value="<% = vscore %>" size="2" class="<% = FieldStyleClass("numeric", "vscore-" & n) %>" /></td>
				<td>at <% = home %></td>
				<td><input type="text" name="hscore-<% = n %>" value="<% = hscore %>" size="2" class="<% = FieldStyleClass("numeric", "hscore-" & n) %>" /></td>
				<td><label for="ot-<% = n %>"><span class="small">OT</span></label> <input type="checkbox" id="ot-<% = n %>" name="ot-<% = n %>" value="true"<% = checkedStr %> /></td>
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
<%		'List open dates.
		'call DisplayOpenDates(2, week)
	end if %>
		<p><input type="submit" name="submit" value="Update" class="button" title="Apply changes." />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel the update." /></p>
	</form>
<%	'List links to view other weeks.
	call DisplayWeekNavigation(1, "") %>
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
		dim rs
		dim vscore, hscore
		dim list, email

		subj = "Football Pool Game Results Notification"
		body = "Games results for Week " & week & " have been updated." & vbCrLf & vbCrLf

		'Show results.
		set rs = WeeklySchedule(week)
		do while not rs.EOF
			vscore = ""
			hscore = ""
			if not IsNull(rs.Fields("Result").Value) then
				vscore = " " & rs.Fields("VisitorScore").Value
				hscore = " " & rs.Fields("HomeScore").Value
			end if
			body = body & rs.Fields("VisitorID").Value & vscore
			if USE_POINT_SPREADS then
				body = body & " (" & PlainTextPointSpread(rs.Fields("PointSpread").Value) & ")"
			end if
			body = body & " @ " & rs.Fields("HomeID").Value & hscore
			if rs.Fields("OT").Value then
				body = body & " OT"
			end if
			body = body & vbCrLf
			rs.moveNext
		loop

		list = GetNotificationList("NotifyOfResultUpdates")
		for each email in list
			call SendMail(email, subj, body)
		next

	end sub %>
