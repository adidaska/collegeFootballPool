<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Set Point Spreads" : AdminOnly = true %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/style.css" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
	<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
	<script type="text/javascript" src="js/jaflGamesSpread.js"></script>

</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/email.asp" -->
<!-- #include file="includes/encryption.asp" -->
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/weekly.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'Get the week to display.
	dim week
	week = GetRequestedWeek()

	'If there is form data, process it.
	dim n, i
	dim gameID, spread
	dim notify, infoMsg
	dim sql, rs
	n = NumberOfGames(week)
	if Request.ServerVariables("Content_Length") > 0 and not CancelRequested() then

		'Process form data for each game.
		for i = 1 to n

			'Get the form fields for the current game.
			gameID = Trim(Request.Form("id-"     & i))
			spread = Trim(Request.Form("spread-" & i))

			'Validate those form fields.
			if spread <> "" then
				if not IsValidPointSpread(spread) then
					FormFieldErrors.Add "spread-" & i, "'" & spread & "' is not a valid point spread."
				end if
			end if
		next

		'If there were any errors, display the error summary message.
		'Otherwise, do the updates.
		if FormFieldErrors.Count > 0 then
			call FormFieldErrorsMessage("Error: Invalid fields. Please correct and resubmit.")
		else
			for i = 1 to n
				gameID = Trim(Request.Form("id-"     & i))
				spread = Trim(Request.Form("spread-" & i))

				'Update the point spread.
				if spread = "" then
					spread = "NULL"
				end if
				sql = "UPDATE Schedule SET" _
				   & " PointSpread = " & spread _
				   & " WHERE GameID = " & gameID
				call DbConn.Execute(sql)

				'Update the results.
				call SetGameResults(gameID)
			next

			'Clear any cached pool results.
			'Note: Don't need to worry about survivor/margin pools here because
			'they do not rely on point spreads.
			call ClearWeeklyResultsCache(week)

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

	'Display the schedule for the specified week. %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<div><input type="hidden" name="week" value="<% = week %>" /></div>
		<table class="main" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge">
			  <th align="left" colspan="9">Week <% = week %></th>
			</tr>
<%	dim gameDate, gameTime
	dim visitor, home, espnID
	dim alt
	set rs = WeeklySchedule(week)
	if not (rs.BOF and rs.EOF) then
		n = 1
		alt = false
		do while not rs.EOF
			gameID   = rs.Fields("GameID").Value
			gameDate = rs.Fields("Date").Value
			gameTime = rs.Fields("Time").Value
			visitor  = rs.Fields("VCity").Value
			spread   = rs.Fields("PointSpread").Value
			home     = rs.Fields("HCity").Value
			espnID   = rs.Fields("EspnGameId").Value

			'Set the team names for display.
			if rs.Fields("VDisplayName") <> "" then
				visitor = rs.Fields("VDisplayName").Value
			end if
			if rs.Fields("HDisplayName") <> "" then
				home = rs.Fields("HDisplayName")
			end if

			if alt then %>
			<tr align="right" class="alt">
<%			else %>
			<tr align="right">
<%			end if
			alt = not alt

			'If there were errors on the form post processing, restore those fields.
			if FormFieldErrors.Count > 0 then
				spread = GetFieldValue("spread-" & n, spread)
			end if %>
				<td><input type="hidden" name="id-<% = n %>" value="<% = gameID %>" /><% = WeekdayName(Weekday(gameDate), true) %></td>
				<td><% = FormatDate(gameDate) %></td>
				<td><% = FormatTime(gameTime) %></td>
				<td><% = visitor %></td>
				<td><div id="spread_<% = espnID %>">Getting Spread...</div></td>
				<td><input type="text" name="spread-<% = n %>" value="<% = spread %>" size="2" class="<% = FieldStyleClass("numeric", "spread-" & n) %>" /></td>
				<td>at <% = home %></td>

			</tr>
<%			rs.MoveNext
			n = n + 1
		loop
		if SERVER_EMAIL_ENABLED then %>
			<tr class="subHeader topEdge">
				<th align="left" colspan="8"><input type="checkbox" id="notify" name="notify" value="true" /> <label for="notify">Send update notification to users.</label></th>
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
		dim list, email

		subj = "Football Pool Point Spread Update Notification"
		body = "The point spreads for games in Week " & week & " have been updated." & vbCrLf & vbCrLf

		'Show the points spreads.
		set rs = WeeklySchedule(week)
		do while not rs.EOF
			body = body _
			     & rs.Fields("VisitorID").Value _
			     & " (" & PlainTextPointSpread(rs.Fields("PointSpread").Value) & ")" _
			     & " @ " & rs.Fields("HomeID").Value & vbCrLf
			rs.moveNext
		loop

		list = GetNotificationList("NotifyOfSpreadUpdates")
		for each email in list
			call SendMail(email, subj, body)
		next

	end sub %>
