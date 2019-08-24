<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Set Playoffs Point Spreads" : AdminOnly = true %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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
	dim gameID, gameDate, gameTime, visitorID, spread, homeID
	dim notify, infoMsg
	dim sql, rs
	n = NumberOfPlayoffGames()
	if Request.ServerVariables("Content_Length") > 0 and not CancelRequested() then
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
				sql = "UPDATE PlayoffsSchedule SET" _
				   & " PointSpread = " & spread _
				   & " WHERE GameID = " & gameID
				call DbConn.Execute(sql)

				'Update the results.
				call PlayoffsSetGameResults(gameID)
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
<%	'Get games.
	dim visitor, home
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
			spread    = rs.Fields("PointSpread").Value
			homeID    = rs.Fields("HomeID").Value

			'Set the team names for display.
			visitor = GetTeamName(rs.Fields("VisitorID").Value)
			home    = GetTeamName(rs.Fields("HomeID").Value)

			if gameRound <> lastRound then
				lastRound = gameRound
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
				spread = GetFieldValue("spread-"    & n, spread)
			end if %>
				<td><input type="hidden" name="id-<% = n %>" value="<% = gameID %>" /><% = WeekdayName(Weekday(gameDate), true) %></td>
				<td><% = FormatDate(gameDate) %></td>
				<td><% = FormatTime(gameTime) %></td>
				<td><% = visitor %></td>
				<td><input type="text" name="spread-<% = n %>" value="<% = spread %>" size="2" class="<% = FieldStyleClass("numeric", "spread-" & n) %>" /></td>
				<td>at <% = home %></td>
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

		subj = "Football Pool Point Spread Update Notification"
		body = "The point spreads for playoff games have been updated." & vbCrLf

		'Show the point spreads.
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
			     & vid _
			     & " (" & PlainTextPointSpread(rs.Fields("PointSpread").Value) & ")" _
			     & " @ " & hid & vbCrLf
			rs.MoveNext
		loop
		body = body & vbCrLf & "(All times Eastern)" & vbCrLf

		list = GetNotificationList("NotifyOfSpreadUpdates")
		for each email in list
			call SendMail(email, subj, body)
		next

	end sub %>

