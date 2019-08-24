<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Update Game Schedule" : AdminOnly = true %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/style.css" />
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
	dim gameID, gameWeek, gameDate, gameTime, tbgame, viewTime
	dim notify, infoMsg
	dim sql, rs
	dim affectedWeeks
	affectedWeeks = Array()
	n = NumberOfGames(week)
	if Request.ServerVariables("Content_Length") > 0 and not CancelRequested() then
		for i = 1 to n
			gameID   = Trim(Request.Form("id-"   & i))
			gameWeek = Trim(Request.Form("week-" & i))
			gameDate = Trim(Request.Form("date-" & i))
			gameTime = Trim(Request.Form("time-" & i))

			'Validate the form fields.
			if not IsNumeric(gameWeek) then
				FormFieldErrors.Add "week-" & i, "'" & gameWeek & "' is not a valid week number."
			else
				if CInt(gameWeek) < 1 or CInt(gameWeek) > NumberOfWeeks() then
					FormFieldErrors.Add "week-" & i, "'" & gameWeek & "' is not a valid week number."
				end if
			end if
			if not IsDate(gameDate) then
				FormFieldErrors.Add "date-" & i, "'" & gameDate & "' is not a valid date."
			end if
			if not IsDate(gameTime) then
				FormFieldErrors.Add "time-" & i, "'" & gameTime & "' is not a valid time, please update."
			end if
		next

		'If there were any errors, display the error summary message.
		'Otherwise, do the updates.
		'added tbgame into mix so can choose the tbgame from this screen
		if FormFieldErrors.Count > 0 then
			call FormFieldErrorsMessage("Error: Invalid fields. Please correct and resubmit.")
		else
			for i = 1 to n
				gameID   = Trim(Request.Form("id-"   & i))
				gameWeek = Trim(Request.Form("week-" & i))
				gameDate = Trim(Request.Form("date-" & i))
				gameTime = Trim(Request.Form("time-" & i))
				tbgame = Request.Form("TBradio")
				'call FormFieldErrorsMessage(Request.Form("TBradio") & " <-- request, i = " & i & (cint(tbgame) = cint(i)))
				
				
				if (cint(tbgame) = cint(i)) then
					tbgame = 1
				else
					tbgame = 0
				end if
				
				call AddAffectedWeek(gameWeek)
				sql = "UPDATE Schedule SET" _
				   & " Week   = "  & gameWeek & ","  _
				   & " [Date] = #" & gameDate & "#," _
				   & " [Time] = #" & gameTime & "#,"  _
				   & " TBgame = " & tbgame & " " _
				   & " WHERE GameID = " & gameID
				call DbConn.Execute(sql)
			next

			'Clear any cached pool results for the affected weeks.
			call SortAffectedWeeks()
			for i = LBound(affectedWeeks) to UBound(affectedWeeks)
				call ClearWeeklyResultsCache(affectedWeeks(i))
				if ENABLE_MARGIN_POOL then
					call ClearMarginResultsCache(affectedWeeks(i))
				end if
			next
			if ENABLE_SURVIVOR_POOL then
				ClearSurvivorStatus(affectedWeeks(0))
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

	'Display the schedule for the specified week. %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<div><input type="hidden" name="week" value="<% = week %>" /></div>
		<table class="main" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge">
				<th align="left">Week</th>
				<th>&nbsp;</th>
				<th align="left" colspan="2">Date</th>
				<th align="left" colspan="2">Time</th>
				<th align="right"><div align="center">Teams</div></th>
				<th>TieBreaker</th>
			</tr>
<%	dim visitor, home, gameTB
	dim alt
	set rs = WeeklySchedule(week)
	if not (rs.BOF and rs.EOF) then
		n = 1
		alt = false
		do while not rs.EOF
			gameID   = rs.Fields("GameID").Value
			gameWeek = rs.Fields("Week").Value
			gameDate = rs.Fields("Date").Value
			gameTime = rs.Fields("Time").Value
			visitor  = rs.Fields("VCity").Value
			home     = rs.Fields("HCity").Value
			gameTB     = rs.Fields("TBgame").Value
			
			if IsNull(gameTime) then 
				viewTime = "TBA"
			else 
				viewTime = FormatFullTime(gameTime) 
			end if
			

			'Set the team names for display
			if rs.Fields("VDisplayName") <> "" then
				visitor = rs.Fields("VDisplayName").Value
			end if
			if rs.Fields("HDisplayName") <> "" then
				home = rs.Fields("HDisplayName")
			end if
			if alt then %>
			<tr class="alt">
<%			else %>
			<tr>
<%			end if
			alt = not alt

			'If there were errors on the form post processing, restore those fields.
			if FormFieldErrors.Count > 0 then
				gameWeek = GetFieldValue("week-" & n, gameWeek)
				gameDate = GetFieldValue("date-" & n, gameDate)
				gameTime = GetFieldValue("time-" & n, gameTime)
			end if %>
				<td><input type="hidden" name="id-<% = n %>" value="<% = gameID %>" /><input type="text" name="week-<% = n %>" value="<% = gameWeek %>" size="2" class="<% = FieldStyleClass("numeric", "week-" & n) %>" /></td>
				<td align="right"><input type="text" name="day-<% = n %>" value="<% = WeekdayName(Weekday(gameDate), true) %>" size="3" class="readonly" readonly="readonly" /></td>
				<td><input type="text" name="date-<% = n %>" value="<% = FormatFullDate(gameDate) %>" size="10" class="<% = FieldStyleClass("numeric readonly", "date-" & n) %>" readonly="readonly" /></td>
				<td><input type="image" src="graphics/calendar.gif" onclick="return openCalendar(this.form, 'day-<% = n %>', 'date-<% = n %>');" title="Select a new date." /></td>
				<td><input type="text" name="time-<% = n %>" value="<% = viewTime %>" size="10" class="<% = FieldStyleClass("numeric readonly", "time-" & n) %>" readonly="readonly" /></td>
				<td><input type="image" src="graphics/clock.gif" onclick="return openClock(this.form, 'time-<% = n %>');" title="Select a new time." /></td>
			  <td align="right"><div align="center">
		        <% = visitor %>
				  at
                <% = home %>
			  </div></td>
			  <td><div align="center">
			    <input type="radio" name="TBradio" id="TBgame-<% = n %>" value="<% = n %>" <%if gameTB = 1 then Response.Write("checked")%>/>
		      </div></td>
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
'<%		'List open dates.
'		call DisplayOpenDates(2, week)
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
	' Adds the specified week to the list of affected weeks, if it is not
	' already present.
	'--------------------------------------------------------------------------
	sub AddAffectedWeek(week)

		dim i

		for i = LBound(affectedWeeks) to UBound(affectedWeeks)
			if affectedWeeks(i) = week then
				exit sub
			end if
		next
		redim preserve affectedWeeks(Ubound(affectedWeeks) + 1)
		affectedWeeks(UBound(affectedWeeks)) = week

	end sub

	'--------------------------------------------------------------------------
	' Sorts the list of affected weeks.
	'--------------------------------------------------------------------------
	sub SortAffectedWeeks()

		dim i, j, tmp

		for i = LBound(affectedWeeks) to UBound(affectedWeeks) - 1
			for j = i + 1 to UBound(affectedWeeks)
				if affectedWeeks(j) < affectedWeeks(i) then
					tmp = affectedWeeks(i)
					affectedWeeks(i) = affectedWeeks(j)
					affectedWeeks(j) = tmp
				end if
			next
		next

	end sub

	'--------------------------------------------------------------------------
	' Sends an email notification to any users who have elected to receive
	' them.
	'--------------------------------------------------------------------------
	sub SendNotifications()

		dim subj, body
		dim i, rs
		dim list, email

		subj = "Football Pool Game Schedule Update Notification"
		body = ""

		'Show the schedule for each affected week.
		for i = LBound(affectedWeeks) to UBound(affectedWeeks)
			if i > 0 then
				body = body & vbCrLf
			end if
			body = body & "The game schedule for Week " & affectedWeeks(i) & " has been updated." & vbCrLf & vbCrLf
			set rs = WeeklySchedule(affectedWeeks(i))
			do while not rs.EOF
				body = body  _
				     & WeekdayName(Weekday(rs.Fields("Date").Value), true) _
			    	 & " " & FormatDate(rs.Fields("Date").Value) _
					 & " " & FormatTime(rs.Fields("Time").Value) _
				     & " " & rs.Fields("VisitorID").Value & " @ " & rs.Fields("HomeID").Value & vbCrLf
				rs.moveNext
			loop
		next
		body = body & vbCrLf & "(All times Eastern)" & vbCrLf

		list = GetNotificationList("NotifyOfScheduleUpdates")
		for each email in list
			call SendMail(email, subj, body)
		next

	end sub %>

