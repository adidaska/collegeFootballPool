<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Weekly Schedule" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link href="styles/style.css" rel="stylesheet" type="text/css" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'Get the week to display.
	dim week
	week = GetRequestedWeek()

	'Display the schedule for that week.
	dim cols
	cols = 8
	if USE_POINT_SPREADS then
		cols = cols + 1
	end if %>
	<table class="main" cellpadding="0" cellspacing="0">
		<tr class="header bottomEdge">
		  <th align="left" colspan="<% = cols %>">Week <% = week %></th>
		</tr>
<%	dim rs
	dim visitor, home
	dim vscore, hscore
	dim ot, result
	dim spread, atsResult
	dim alt
	set rs = WeeklySchedule(week)
	if not (rs.BOF and rs.EOF) then
		alt = false
		do while not rs.EOF
			visitor   = rs.Fields("VCity").Value
			home      = rs.Fields("HCity").Value
			vscore    = rs.Fields("VisitorScore").Value
			hscore    = rs.Fields("HomeScore").Value
			ot        = rs.Fields("OT").Value
			result    = rs.Fields("Result").Value
			spread    = rs.Fields("PointSpread").Value
			atsResult = rs.Fields("ATSResult").Value

			'Set the team names for display.
			if rs.Fields("VDisplayName") <> "" then
				visitor = rs.Fields("VDisplayName").Value
			end if
			if rs.Fields("HDisplayName") <> "" then
				home = rs.Fields("HDisplayName")
			end if

			'Highlight the results.
			if result = rs.Fields("VisitorID").Value then
				visitor = FormatWinner(visitor)
				vscore  = FormatWinner(vscore)
			elseif result = rs.Fields("HomeID").Value then
				home   = FormatWinner(home)
				hscore = FormatWinner(hscore)
			end if
			if atsResult = rs.Fields("VisitorID").Value then
				visitor = FormatATSWinner(visitor)
				vscore  = FormatATSWinner(vscore)
			elseif atsResult = rs.Fields("HomeID").Value then
				home   = FormatATSWinner(home)
				hscore = FormatATSWinner(hscore)
			end if
			if IsNull(hscore) then
				hscore = "&nbsp;"
			end if
			if IsNull(vscore) then
				vscore = "&nbsp;"
			end if

			'Set the OT display.
			if ot then
				ot = "OT"
			else
				ot = "&nbsp;"
			end if

			if alt then %>
		<tr align="right" class="alt">
<%			else %>
		<tr align="right">
<%			end if
			alt = not alt %>
			<td align="left"><% = WeekdayName(Weekday(rs.Fields("Date").Value), true) %></td>
			<td><% = FormatDate(rs.Fields("Date").Value) %></td>
			<td><% = FormatTime(rs.Fields("Time").Value) %></td>
			<td><% = visitor %></td>
			<td><% = vscore %></td>
<%			if USE_POINT_SPREADS then %>
			<td><% = FormatPointSpread(spread) %></td>
<%			end if %>
			<td>at <% = home %></td>
			<td><% = hscore %></td>
			<td><span class="small"><% = ot %></span></td>
		</tr>
<%			rs.MoveNext
		loop %>
	</table>
<%		'List open dates.
		'call DisplayOpenDates(1, week)
	end if

	'List links to view other weeks.
	call DisplayWeekNavigation(1, "") %>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>