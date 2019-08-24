<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Playoffs Schedule" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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
<!-- #include file="includes/playoffs.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'Build the display.
	dim cols
	cols = 9
	if USE_POINT_SPREADS then
		cols = cols + 1
	end if %>
	<table class="main" cellpadding="0" cellspacing="0">
<%	dim i
	dim lastRound, gameRound, conference
	dim gameID, gameDate, gameTime
	dim vid, hid, visitor, home
	dim vscore, hscore, spread
	dim ot, result, atsResult
	dim sql, rs
	dim alt
	sql = "SELECT * FROM PlayoffsSchedule ORDER BY Round, Date, Time"
	set rs = DbConn.Execute(sql)
	if not (rs.BOF and rs.EOF) then
		i = 0
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
				i = 1
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
				vscore  = FormatWinner(vscore)
			elseif result = hid then
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


			if alt then %>
		<tr align="right" class="alt">
<%			else %>
		<tr align="right">
<%			end if
			alt = not alt %>
			<td><% = WeekdayName(Weekday(gameDate), true) %></td>
			<td><% = FormatDate(gameDate) %></td>
			<td><% = FormatTime(gameTime) %></td>
			<td><% = conference %></td>
			<td><% = visitor %></td>
			<td><% = vscore %></td>
<%			if USE_POINT_SPREADS then %>
			<td><% = FormatPointSpread(spread) %></td>
<%			end if %>
			<td><% = GetConjunction(gameRound) & " " & home %></td>
			<td><% = hscore %></td>
			<td align="left">
<%			if ot = 0 then %>
				&nbsp;
<%			elseif ot = 1 then %>
				<span class="small">OT</span>
<%			else %>
				<span class="small">OT(<% = ot %>)</span>
<%			end if %>
			</td>
		</tr>
<%			rs.MoveNext
		loop %>
	</table>
<%	end if %>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>
