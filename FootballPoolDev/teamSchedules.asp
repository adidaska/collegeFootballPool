<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Team Schedules" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
    <link href="styles/style.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'Check for a team id.
	dim id
	id = Request("id")

	'If one was specified, show the schedule for it.
	dim sql, rs
	dim wins, losses, ties, atsWins, atsLosses, atsTies
	dim n
	dim openDate, str, score, oppName, oppScore, ot, result, atsResult, spread
	dim cols
	dim alt
	cols = 8
	alt = false
	if USE_POINT_SPREADS then
		cols = cols + 2
	end if
	if id <> "" then

		'Get the team name.
		sql = "SELECT * FROM Teams WHERE TeamID = '" & id & "'"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			wins    = 0 : losses    = 0 : ties    = 0
			atsWins = 0 : atsLosses = 0 : atsTies = 0
			str = rs.Fields("City").Value & " " & rs.Fields("Name").Value
			rs.Close
		end if %>
	<table class="main" cellpadding="0" cellspacing="0">
		<tr class="header bottomEdge">
			<th align="left" colspan="<% = cols %>"><% = str %></th>
		</tr>
<%		sql = "SELECT * FROM Teams, Schedule" _
		   & " WHERE (VisitorID = '" & id & "' and TeamID = HomeID)" _
		   & " OR    (HomeID    = '" & id & "' and TeamID = VisitorID)" _
		   & " ORDER BY Week"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			n = 0
			openDate = false
			do while not rs.EOF
				n = n + 1

				'Check for open date.
				if n <> rs.Fields("Week").Value then
					call AddOpenDate(n)
					n = n + 1
					openDate = true
				end if

				'Get the opposing team name and game scores.
				spread = rs.Fields("PointSpread").Value
				if rs.Fields("DisplayName") <> "" then
					oppName = rs.Fields("DisplayName").Value
				else
					oppName = rs.Fields("City").Value
				end if
				if id = rs.Fields("VisitorID").Value then
					oppName = "at " & oppName
					score = rs.Fields("VisitorScore").Value
					oppScore = rs.Fields("HomeScore").Value
				else
					spread = -spread
					score = rs.Fields("HomeScore").Value
					oppScore = rs.Fields("VisitorScore").Value
				end if

				'Set the game results.
				result = rs.Fields("Result").Value
				atsResult = rs.Fields("ATSResult").Value
				if result = id then
					result = "W"
					score = score & "-" & oppScore
					wins = wins + 1
				elseif result = TIE_STR then
					result = "T"
					score = score & "-" & oppScore
					ties = ties + 1
				elseif result <> "" then
					result = "L"
					score = score & "-" & oppScore
					losses = losses + 1
				else
					result = "&nbsp;"
					score = "&nbsp;"
				end if
				if atsResult = id then
					atsResult = "(W)"
					atsWins = atsWins + 1
				elseif atsResult = TIE_STR then
					atsResult = "(T)"
					atsTies = atsTies + 1
				elseif atsResult <> "" then
					atsResult = "(L)"
					atsLosses = atsLosses + 1
				end if

				'Set the overtime indicator.
				if rs.Fields("OT").Value then
					ot = "OT"
				else
					ot = "&nbsp;"
				end if

				'Display the info.
				if alt then %>
		<tr align="right" class="alt singleLine">
<%				else %>
		<tr align="right" class="singleLine">
<%				end if
				alt = not alt %>
			<td><% = rs.Fields("Week").Value %></td>
			<td><% = WeekdayName(Weekday(rs.Fields("Date").Value), true) %></td>
			<td><% = FormatDate(rs.Fields("Date").Value) %></td>
			<td><% = FormatTime(rs.Fields("Time").Value) %></td>
<%				if USE_POINT_SPREADS then %>
			<td><% = FormatPointSpread(spread) %></td>
			<td align="center"><% = atsResult %></td>
<%				end if %>
			<td><% = oppName %></td>
			<td align="center"><% = result %></td>
			<td><% = score %></td>
			<td><span class="small"><% = ot %></span></td>
		</tr>
<%       		rs.MoveNext
       		loop

			'If no open date was found, the team must have an open date in the
			'final week, so add a row for it. (Note: this has only occured in
			'years where there were an odd number of teams.)
			if not openDate then
				n = n + 1
				call AddOpenDate(n)
			end if
		end if

		'Display the team's record.
		str = FormatRecord(wins, losses, ties)
		if wins + losses + ties <> 0 then
			str = str & " (" & FormatPercentage((wins + ties / 2) / (wins + ties + losses)) & ")"
		end if
		if USE_POINT_SPREADS then
			str = str & " straight up,&nbsp;" & FormatRecord(atsWins, atsLosses, atsTies)
			if wins + losses + ties <> 0 then
				str = str & " (" & FormatPercentage((atsWins + atsTies / 2) / (atsWins + atsTies + atsLosses)) & ")"
			end if
			str = str & " vs. spread"
		end if %>
		<tr class="header topEdge">
			<th align="right" colspan="<% = cols %>"><% = str %></th>
		</tr>
	</table>
<%	end if

	'List the teams for selection. %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<p></p>
		<table>
			<tr>
				<td><strong>Select team:</strong></td>
				<td>
					<select name="id">
						<option value=""></option>
<%	sql = "SELECT * FROM Teams ORDER BY City, Name"
	set rs = DbConn.Execute(sql)
	do while not rs.EOF %>
						<option value="<% = rs.Fields("TeamID") %>"<% if rs.Fields("TeamID") = id then Response.Write(" selected=""selected""") end if %>><% = rs.Fields("City") & " " & rs.Fields("Name") %></option>
<%		rs.MoveNext
		loop %>
					</select>
				</td>
				<td><input type="submit" name="submit" value="Select" class="button" title="Show schedule for the selected team." /></td>
			</tr>
		</table>
	</form>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>
<%	'**************************************************************************
	'* Local functions and subroutines.                                       *
	'**************************************************************************

	'---------------------------------------------------------------------------
	' Displays a table row to indicate an open date in the team's schedule.
	'---------------------------------------------------------------------------
	sub AddOpenDate(n)

		if alt then
			Response.Write(String(2, vbTab) & "<tr align=""right"" class=""alt"">" & vbCrLf)
		else
			Response.Write(String(2, vbTab) & "<tr align=""right"">" & vbCrLf)
		end if
		alt = not alt
		Response.Write(String(3, vbTab) & "<td>" & n & "</td>" & vbCrLf)
		Response.Write(String(3, vbTab) & "<td align=""center"" colspan=""3""><em>Open Date</em></td>" & vbCrLf)
		Response.Write(String(3, vbTab) & "<td colspan=""" & (cols -4) & """>&nbsp;</td>" & vbCrLf)
		Response.Write(String(2, vbTab) & "</tr>" & vbCrLf)

	end sub %>
