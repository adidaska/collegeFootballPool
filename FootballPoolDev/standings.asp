<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Standings" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Define the minimum week numbers for applying division tiebreakers and
	'conference standings.
	const DIVISION_TIEBREAKER_MIN_WEEK  =  6
	const CONFERENCE_STANDINGS_MIN_WEEK = 12

	'Open the database.
	call OpenDB()

	'Check if a specific week for the standings was requested. Defaults to the
	'current week if no valid number was given.
	dim week
	week = GetRequestedWeek()

	'Set the tiebreaker flag.
	dim tiebreakersOn
	tiebreakersOn = false
	if week >= DIVISION_TIEBREAKER_MIN_WEEK then
		tiebreakersOn = true
	end if

	'Initialize the tiebreaker footnotes count.
	dim footnotesCount
	footnotesCount = 0

	'Determine the standings type (division or conference).
	dim conferenceStandingsOn
	conferenceStandingsOn = false
	if week >= CONFERENCE_STANDINGS_MIN_WEEK and _
		LCase(Request.QueryString("type")) = "conf" then
		conferenceStandingsOn = true
	end if

	'Build the teams list, grouped by conference and division.
	dim teams, n, i, sql, rs
	n = NumberOfTeams()
	redim teams(n - 1)
	i = 0
	sql = "SELECT * FROM Teams ORDER BY Conference, Division, City, Name"
	set rs = DbConn.Execute(sql)
	if not (rs.BOF and rs.EOF) then
		do while not rs.EOF
			set teams(i) = new TeamObj
			teams(i).setData(rs.Fields("TeamID").Value)
			i = i + 1
			rs.MoveNext
		loop
	end if

   'Sort the teams within each division.
   dim div, m
	div = teams(0).division
	m = 0
	for i = 0 to UBound(teams)
		if teams(i).division <> div then
			call SortDivision(m, i - 1)
			div = teams(i).division
			m = i
		end if
	next
	call SortDivision(m, i - 1)

	'Handle conference standings, if requested.
	if conferenceStandingsOn then

		'Clear any existing tiebreaker footnotes.
		for i = 0 to UBound(teams)
			teams(i).tbText = ""
		next

		'Find the starting and ending indices of each conference.
		dim afcFirst, afcLast, nfcFirst, nfcLast, conf
		afcFirst = 0
		afcLast = 0
		conf = teams(0).conference
		do while teams(afcLast).conference = conf
			afcLast = afcLast + 1
		loop
		nfcFirst = afcLast
		afcLast = afcLast - 1
		nfcLast = UBound(teams)

		'Sort the teams within each conference.
		call SortConference(afcFirst, afcLast)
		call SortConference(nfcFirst, nfcLast)
	end if

	'Build the standings display.
	if week <> CurrentWeek() then %>
	<h2>Standings as of Week <% = week %></h2>
<%	end if %>
	<table class="main" cellpadding="0" cellspacing="0">
<%	dim alt, styleClass, rank, str
	div = -1
	conf = -1
	alt = false
	for i = 0 to UBound(teams)
		if teams(i).streak = "" then
			teams(i).streak = "&nbsp;"
		end if
		if (not conferenceStandingsOn and  div <> teams(i).division  ) or _
		   (    conferenceStandingsOn and conf <> teams(i).conference) then
			div  = teams(i).division
			conf = teams(i).conference
			rank = 1
			if teams(i).conference = 1 then
				styleClass = "afc"
			else
				styleClass = "nfc"
			end if
			if i > 0 then
				styleClass = styleClass & " topEdge"
			end if
			alt = false %>
		<tr align="right" class="<% = styleClass %> header bottomEdge">
<%			if conferenceStandingsOn then %>
			<th>&nbsp;</th>
			<th align="left" title="Conference"><% = ConferenceNames(teams(i).conference - 1) %></th>
			<th title="Division">Div</th>
<%			else %>
			<th align="left" title="Division"><% = ConferenceNames(teams(i).conference - 1) & " " & DivisionNames(teams(i).division - 1) %></th>
<%			end if %>
			<th align="center" title="Win-loss-tie record.">W-L-T</th>
			<th align="center" title="Winning percentage.">Pct</th>
<%			if USE_POINT_SPREADS then %>
			<th align="center" colspan="2" title="Record vs. the spread.">(vs. spread)</th>
<%			end if %>
			<th title="Total points scored.">PF</th>
			<th title="Total points allowed.">PA</th>
<%			if conferenceStandingsOn then %>
			<th title="Record vs. conference teams.">Conf</th>
<%			else %>
			<th title="Record in home games.">Home</th>
			<th title="Record in road games.">Road</th>
			<th title="Record vs. AFC teams.">AFC</th>
			<th title="Record vs. NFC teams.">NFC</th>
<%			end if %>
			<th title="Record vs. division teams.">Div</th>
<%			if conferenceStandingsOn then %>
			<th title="Strength of victory (combined winning percentage of opponents beaten).">SoV</th>
			<th title="Strength of schedule (combined winning percentage of all opponents).">SoS</th>
			<th title="Combined points for and points allowed ranking among conference teams.">CPR</th>
			<th title="Combined points for and points allowed ranking among all teams.">PR</th>
			<th title="Net points in conference games.">CNP</th>
			<th title="Net points in all games.">NP</th>
<%			else %>
			<th title="Current win, loss or tie streak.">Strk</th>
<%			end if %>
</tr>
<%		end if
		styleClass = ""
		if alt then
			styleClass = "alt"
		end if
		if conferenceStandingsOn and rank = 6 then
			styleClass = styleClass & " playoffsBreak"
		end if
		styleClass = Trim(styleClass) %>
		<tr align="right" class="<% = styleClass %>">
<%		alt = not alt
		if conferenceStandingsOn then %>
			<td align="right"><% = rank %></td>
			<td align="left">
<%		else %>
			<td align="left">
<%		end if
		rank = rank + 1

		'Check if there is a tiebreaker footnote for this team.
		str = ""
		if teams(i).tbText <> "" then
			footnotesCount = footnotesCount + 1
			str = " <span class=""small""><sup>(" & footnotesCount & ")</sup></span>"
		end if %>
				<a href="teamSchedules.asp?id=<% = teams(i).id %>"><% = "<strong>" & teams(i).name & "</strong>" %></a><% = str %>
			</td>
<%		if conferenceStandingsOn then %>
			<td align="center"><% = Left(DivisionNames(teams(i).division - 1), 1) %></td>
<%		end if %>
			<td><% = teams(i).totalWins %>-<% = teams(i).totalLosses %>-<% = teams(i).totalTies %></td>
			<td><% = FormatPercentage(teams(i).totalPct) %></td>
<%			if USE_POINT_SPREADS then %>
			<td>(<% = teams(i).atsWins %>-<% = teams(i).atsLosses %>-<% = teams(i).atsTies %></td>
			<td><% = FormatPercentage(teams(i).atsPct) %>)</td>
<%			end if %>
			<td><% = teams(i).pointsFor %></td>
			<td><% = teams(i).pointsAgainst %></td>
<%		if conferenceStandingsOn then
			if teams(i).conference = 1 then %>
			<td><% = FormatRecord(teams(i).afcWins, teams(i).afcLosses, teams(i).afcTies) %></td>
<%			else %>
			<td><% = FormatRecord(teams(i).nfcWins, teams(i).nfcLosses, teams(i).nfcTies) %></td>
<%			end if
		else %>
			<td><% = FormatRecord(teams(i).homeWins, teams(i).homeLosses, teams(i).homeTies) %></td>
			<td><% = FormatRecord(teams(i).roadWins, teams(i).roadLosses, teams(i).roadTies) %></td>
			<td><% = FormatRecord(teams(i).afcWins, teams(i).afcLosses, teams(i).afcTies) %></td>
			<td><% = FormatRecord(teams(i).nfcWins, teams(i).nfcLosses, teams(i).nfcTies) %></td>
<%		end if %>
			<td><% = FormatRecord(teams(i).divWins, teams(i).divLosses, teams(i).divTies) %></td>
<%		if conferenceStandingsOn then %>
			<td><% = FormatPercentage(GetStrengthOfVictory(i)) %></td>
			<td><% = FormatPercentage(GetStrengthOfSchedule(i)) %></td>
			<td><% = GetCombinedPointsRank(i, true) %></td>
			<td><% = GetCombinedPointsRank(i, false) %></td>
			<td><% = teams(i).confNetPts %></td>
			<td><% = teams(i).netPts %></td>
<%		else %>
			<td align="center"><% = teams(i).streak %></td>
<%		end if %>
		</tr>
<%	next %>
	</table>
<%	'Display the tiebreaker footnotes, if any.
	dim width
	if footnotesCount > 0 then
		width = 32
		if conferenceStandingsOn then
			width = 42
		end if %>
	<p class="small">Tiebreaker notes:</p>
	<div style="width:<% = width %>em;">
		<ol class="small">
<%		for i = 0 to UBound(teams)
			if teams(i).tbText <> "" then %>
			<li><% = teams(i).tbText %></li>
<%			end if
		next %>
		</ol>
	</div>
<% end if

	'Display the link to switch between division and conference standings, if
	'conference standings are active.
	if week >= CONFERENCE_STANDINGS_MIN_WEEK then
		if not conferenceStandingsOn then %>
	<p><a href="<% = Request.ServerVariables("SCRIPT_NAME") %>?type=conf<% if Request("Week") <> "" then Response.Write("&week=" & week) end if %>" title="Conference standings.">View Conference Standings...</a></p>
<%		else %>
	<p><a href="<% = Request.ServerVariables("SCRIPT_NAME") %>?type=div<% if Request("Week") <> "" then Response.Write("&week=" & week) end if %>" title="Division standings.">...Back to Division Standings</a></p>
<%		end if
	end if %>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>
<%	'**************************************************************************
	'* Local functions and subroutines.                                       *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Used on the global team object array to sort teams within the given range
	' based on team records and tiebreakers (when applicable). The teams should
	' already be grouped by conference and division before this routine is
	' called.
	'--------------------------------------------------------------------------
	sub SortDivision(first, last)

		dim i, j, tmp
		dim n, same, tie

		if last <= first then
			exit sub
		end if

		'Sort based on overall record.
		for i = first to last
			teams(i).tb = teams(i).totalPct
		next
		call SortByTB(first, last)

		'For teams with the same record, apply the tiebreakers (if active).
		'Otherwise, sort by name.
		for i = first to last - 1
			same = true
			tie = false
			n = teams(i).totalPct
			j = i + 1
			do while (j <= last and same)
				if teams(j).totalPct = n then
					tie = true
					j = j + 1
				else
					same = false
				end if
			loop
			if tie then
				if tiebreakersOn then 
					call ApplyHeadToHeadTiebreaker(i, j - 1)
				else
					call SortByName(i, j - 1)
				end if
			end if
		next

	end sub

	'--------------------------------------------------------------------------
	' Used on the global team object array to sort teams within the given range
	' based on team records and tiebreakers (when applicable). The teams should
	' already be sorted by a call to SortDivision() before this routine is
	' called.
	'--------------------------------------------------------------------------
	sub SortConference(first, last)

		dim i, j
		dim div
		dim rank
		dim tmp
		dim numDivs
		dim n, same, tie

		'Teams should be sorted by division at this point. Set the division
		'rank.
		div = -1
		for i = first to last
			if teams(i).division <> div then
				rank = 1
				div = teams(i).division
			end if
			teams(i).divRank = rank
			rank = rank + 1
		next

		'Find the number of divisions.
		numDivs = UBound(DivisionNames) + 1

		'Move the division winners to the top.
		for i = first to first + numDivs - 1
			if teams(i).divRank <> 1 then
				for j = first + numDivs to last
					if teams(j).divRank = 1 then
						set tmp = teams(i)
						set teams(i) = teams(j)
						set teams(j) = tmp
					end if
				next
			end if
		next

		'Sort those teams based on overall record.
		for i = first to last
			teams(i).tb = teams(i).totalPct
		next
		call SortByTB(first, first + numDivs - 1)

		'Sort the remaining teams by record.
		call SortByTB(first + numDivs, last)

		'For division winners with the same record, apply the tiebreakers (if
		'active). Otherwise, sort by name.
		for i = first to first + numDivs - 2
			same = true
			tie = false
			n = teams(i).totalPct
			j = i + 1
			do while (j <= first + numDivs - 1 and same)
				if teams(j).totalPct = n then
					tie = true
					j = j + 1
				else
					same = false
				end if
			loop
			if tie then
				if tiebreakersOn then 
					call ApplyHeadToHeadTiebreaker(i, j - 1)
				else
					call SortByName(i, j - 1)
				end if
			end if
		next

		'For remaining teams with the same record, apply the tiebreakers (if
		'active). Otherwise, sort by name.
		for i = first + numDivs to last
			same = true
			tie = false
			n = teams(i).totalPct
			j = i + 1
			do while (j <= last and same)
				if teams(j).totalPct = n then
					tie = true
					j = j + 1
				else
					same = false
				end if
			loop
			if tie then
				if tiebreakersOn then 
					call ApplyHeadToHeadTiebreaker(i, j - 1)
				else
					call SortByName(i, j - 1)
				end if
			end if
		next

	end sub

	'**************************************************************************
	'* Tiebreaker subroutines.                                                 
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Head to head tiebreaker.
	'   #1 Division tiebreaker.
	'   #1 Wild Card tiebreaker.
	'--------------------------------------------------------------------------
	sub ApplyHeadToHeadTiebreaker(first, last)

		dim i, j
		dim tw, tl, tt
		dim w, l, t
		dim savedLast

		'Check range and exit if invalid.
		if first >= last then
			exit sub
		end if

		'Compare head-to-head records.

		'For two teams...
		if last - first = 1 then
			teams(first).tb = CountMatches(teams(first).oppWonAgainst, teams(last).id)
			teams(last).tb = CountMatches(teams(last).oppWonAgainst, teams(first).id)
			call SortByTB(first, last)
			if teams(first).tb <> teams(last).tb then
				teams(first).tbText = "Beat " & teams(last).name & " head-to-head."
			else
				if InSameDivision(first, last) then
					call ApplyDivisionRecordTiebreaker(first, last)
				else
					call ApplyConferenceRecordTiebreaker(first, last)
				end if
			end if

		'For three or more teams...
		else

			'If the teams are in different divisions, limit processing to the
			'highest ranked team in each division.
			if not InSameDivision(first, last) then
				savedLast = last
				last = SortDivisionLeaders(first, last)
			end if

			for i = first to last
				tw = 0 : tl = 0 : tt = 0
				for j = first to last
					if i <> j then
						w = CountMatches(teams(i).oppWonAgainst, teams(j).id)
						l = CountMatches(teams(i).oppLostTo, teams(j).id)
						t = CountMatches(teams(i).oppTiedWith, teams(j).id)

						'If any team has not played each of the others, skip to
						'the next tiebreaker.
						if w + l + t = 0 then
							if InSameDivision(first, last) then
								call ApplyDivisionRecordTiebreaker(first, last)
							else
								call ApplyConferenceRecordTiebreaker(first, last)
								if teams(first).tb = teams(first + 1).tb and _
								   teams(first).tb > teams(first + 2).tb then
									call ApplyHeadToHeadTiebreaker(first + 2, savedLast)
								else
									call ApplyHeadToHeadTiebreaker(first + 1, savedLast)
								end if
							end if
							exit sub
						end if

						tw = tw + w
						tl = tl + l
						tt = tt + t
					end if
				next
				teams(i).tb = CalcPct(tw, tl, tt)
			next
			call SortByTB(first, last)

			'For teams within the same division, compare by winning percentage.
			if InSameDivision(first, last) then
				if teams(first).tb > teams(first + 1).tb then
					call ApplyHeadToHeadTiebreaker(first + 1, last)
					teams(first).tbText = "Better head-to-head record against " & ListTeamNames(first + 1, last) & "."
				elseif teams(first).tb = teams(first + 1).tb and _
				       teams(first).tb > teams(first + 2).tb then
					call ApplyHeadToHeadTiebreaker(first, first + 1)
					call ApplyHeadToHeadTiebreaker(first + 2, last)
				else
					call ApplyDivisionRecordTiebreaker(first, last)
				end if

			'For teams in different divisions, advance a team if it swept all
			'the others or drop a team if it was swept by all the others.
			else
				if teams(first).tb = 1 and teams(first + 1).tb < 1 then
					call ApplyHeadToHeadTiebreaker(first + 1, last)
					teams(first).tbText = "Better head-to-head record against " & ListTeamNames(first + 1, last) & "."
				elseif teams(last).tb = 0 and teams(last - 1).tb > 0 then
					call ApplyHeadToHeadTiebreaker(first, last - 1)
				else
					call ApplyConferenceRecordTiebreaker(first, last)
				end if
			end if

		end if

	end sub

	'--------------------------------------------------------------------------
	' Division record tiebreaker.
	'   #2 Division tiebreaker.
	'--------------------------------------------------------------------------
	sub ApplyDivisionRecordTiebreaker(first, last)

		dim i

		'Sort by division record.
		for i = first to last
			teams(i).tb = teams(i).divPct
		next
		call SortByTB(first, last)

		'Compare division records.

		'For two teams...
		if last - first = 1 then
			if teams(first).tb > teams(last).tb then
				teams(first).tbText = "Better division record than " & teams(last).name & "."
			else
				call ApplyCommonGamesTiebreaker(first, last)
			end if

		'For three or more teams...
		else
			if teams(first).tb > teams(first + 1).tb then
				call ApplyHeadToHeadTiebreaker(first + 1, last)
				teams(first).tbText = "Better division record than " & ListTeamNames(first + 1, last) & "."
			elseif teams(first).tb = teams(first + 1).tb and _
			       teams(first).tb > teams(first + 2).tb then
				call ApplyHeadToHeadTiebreaker(first, first + 1)
				call ApplyHeadToHeadTiebreaker(first + 2, last)
			else
				call ApplyCommonGamesTiebreaker(first, last)
			end if

		end if

	end sub

	'--------------------------------------------------------------------------
	' Conference record tiebreaker.
	'   #4 Division tiebreaker.
	'   #2 Wild Card tiebreaker.
	'--------------------------------------------------------------------------
	sub ApplyConferenceRecordTiebreaker(first, last)

		dim i

		'Sort by conference record.
		for i = first to last
			teams(i).tb = teams(i).confPct
		next
		call SortByTB(first, last)

		'Compare conference records.

		'For two teams...
		if last - first = 1 then
			if teams(first).tb > teams(last).tb then
				teams(first).tbText = "Better conference record than " & teams(last).name & "."
			else
				if InSameDivision(first, last) then
					call ApplyStrengthOfVictoryTiebreaker(first, last)
				else
					call ApplyCommonGamesTiebreaker(first, last)
				end if
			end if

		'For three or more teams...
		else
			if teams(first).tb > teams(first + 1).tb then
				call ApplyHeadToHeadTiebreaker(first + 1, last)
				teams(first).tbText = "Better conference record than " & ListTeamNames(first + 1, last) & "."
			elseif teams(first).tb = teams(first + 1).tb and _
			       teams(first).tb > teams(first + 2).tb then
				call ApplyHeadToHeadTiebreaker(first, first + 1)
				call ApplyHeadToHeadTiebreaker(first + 2, last)
			else
				if InSameDivision(first, last) then
					call ApplyStrengthOfVictoryTiebreaker(first, last)
				else
					call ApplyCommonGamesTiebreaker(first, last)
				end if
			end if

		end if

	end sub

	'--------------------------------------------------------------------------
	' Common games tiebreaker.
	'   #3 Division tiebreaker.
	'   #3 Wild Card tiebreaker.
	'--------------------------------------------------------------------------
	sub ApplyCommonGamesTiebreaker(first, last)

		dim i, j
		dim list, c
		dim found
		dim w, l, t

		'Find common games.
		list = ""
		for i = 0 to UBound(teams)
			found = true
			for j = first to last
				c = CountMatches(teams(j).oppWonAgainst, teams(i).id) _
				  + CountMatches(teams(j).oppLostTo, teams(i).id) _
				  + CountMatches(teams(j).oppTiedWith, teams(i).id)
				if c = 0 then
					found = false
				end if
			next
			if found then
				if list = "" then
					list = teams(i).id
				else
					list = list & "," & teams(i).id
				end if
			end if
		next
		list = Split(list, ",")

		'Skip to the next tiebreaker if the number of common games is less than
		'the minimum of four.
		if UBound(list) < 3 then
			if InSameDivision(first, last) then
				call ApplyConferenceRecordTiebreaker(first, last)
			else
				call ApplyStrengthOfVictoryTiebreaker(first, last)
			end if
			exit sub
		end if

		'Get each team's winning percentage in those common games.
		for i = first to last
			w = 0 : l = 0 : t = 0
			for j = 0 to UBound(list)
				w = w + CountMatches(teams(i).oppWonAgainst, list(j))
				l = l + CountMatches(teams(i).oppLostTo, list(j))
				t = t + CountMatches(teams(i).oppTiedWith, list(j))
			next
			teams(i).tb = CalcPct(w, l, t)
			teams(i).tb1 = w
			teams(i).tb2 = l
			teams(i).tb3 = t
		next
		call SortByTB(first, last)

		'Compare record in common games.

		'For two teams...
		if last - first = 1 then
			if teams(first).tb > teams(last).tb then
				teams(first).tbText = "Better record in common games" _
					& " (" & FormatRecord(teams(first).tb1, teams(first).tb2, teams(first).tb3) _
					& " vs. " & ListCommonOpponentNames(list) & ")" _
					& " than " & teams(last).name _
					& " (" & FormatRecord(teams(last).tb1, teams(last).tb2, teams(last).tb3) _
					& ")."
			else
				if InSameDivision(first, last) then
					call ApplyConferenceRecordTiebreaker(first, last)
				else
					call ApplyStrengthOfVictoryTiebreaker(first, last)
				end if
			end if

		'For three or more teams...
		else
			if teams(first).tb > teams(first + 1).tb then
				call ApplyHeadToHeadTiebreaker(first + 1, last)
				teams(first).tbText = "Better record in common games" _
					& " (" & FormatRecord(teams(first).tb1, teams(first).tb2, teams(first).tb3) _
					& " vs. " & ListCommonOpponentNames(list) & ")" _
					& " than "
				for i = first + 1 to last
					teams(first).tbText = teams(first).tbText & teams(i).name _
						& " (" & FormatRecord(teams(i).tb1, teams(i).tb2, teams(i).tb3) & ")"_
						& ", "
				next
				teams(first).tbText = Left(teams(first).tbText, Len(teams(first).tbText) - 2) & "."
			elseif teams(first).tb = teams(first + 1).tb and _
			       teams(first).tb > teams(first + 2).tb then
				call ApplyHeadToHeadTiebreaker(first, first + 1)
				call ApplyHeadToHeadTiebreaker(first + 2, last)
			else
				if InSameDivision(first, last) then
					call ApplyConferenceRecordTiebreaker(first, last)
				else
					call ApplyStrengthOfVictoryTiebreaker(first, last)
				end if
			end if

		end if

	end sub

	'--------------------------------------------------------------------------
	' Strength of victory tiebreaker.
	'   #5 Division tiebreaker.
	'   #4 Wild Card tiebreaker.
	'--------------------------------------------------------------------------
	sub ApplyStrengthOfVictoryTiebreaker(first, last)

		dim i

		'Sort teams by their combined winning percentage of opponents they have
		'defeated.
		for i = first to last
			teams(i).tb = GetStrengthOfVictory(i)
		next
		call SortByTB(first, last)

		'Compare strength of victory.

		'For two teams...
		if last - first = 1 then
			if teams(first).tb > teams(last).tb then
				teams(first).tbText = "Better strength of victory than " & teams(last).name & "."
			else
				call ApplyStrengthOfScheduleTiebreaker(first, last)
			end if

		'For three or more teams...
		else
			if teams(first).tb > teams(first + 1).tb then
				call ApplyHeadToHeadTiebreaker(first + 1, last)
				teams(first).tbText = "Better strength of victory than " & ListTeamNames(first + 1, last) & "."
			elseif teams(first).tb = teams(first + 1).tb and _
			       teams(first).tb > teams(first + 2).tb then
				call ApplyHeadToHeadTiebreaker(first, first + 1)
				call ApplyHeadToHeadTiebreaker(first + 2, last)
			else
				call ApplyStrengthOfScheduleTiebreaker(first, last)
			end if

		end if

	end sub

	'--------------------------------------------------------------------------
	' Strength of schedule tiebreaker.
	'   #6 Division tiebreaker.
	'   #5 Wild Card tiebreaker.
	'--------------------------------------------------------------------------
	sub ApplyStrengthOfScheduleTiebreaker(first, last)

		dim i

		'Sort teams by their combined winning percentage of all of their
		'opponents.
		for i = first to last
			teams(i).tb = GetStrengthOfSchedule(i)
		next
		call SortByTB(first, last)

		'Compare strength of schedule.

		'For two teams...
		if last - first = 1 then
			if teams(first).tb > teams(last).tb then
				teams(first).tbText = "Better strength of schedule than " & teams(last).name & "."
			else
				call ApplyConferenceCombinedPointsRankingTiebreaker(first, last)
			end if

		'For three or more teams...
		else
			if teams(first).tb > teams(first + 1).tb then
				call ApplyHeadToHeadTiebreaker(first + 1, last)
				teams(first).tbText = "Better strength of schedule than " & ListTeamNames(first + 1, last) & "."
			elseif teams(first).tb = teams(first + 1).tb and _
			       teams(first).tb > teams(first + 2).tb then
				call ApplyHeadToHeadTiebreaker(first, first + 1)
				call ApplyHeadToHeadTiebreaker(first + 2, last)
			else
				call ApplyConferenceCombinedPointsRankingTiebreaker(first, last)
			end if

		end if

	end sub

	'--------------------------------------------------------------------------
	' Combined ranking among conference teams in points scored and points
	' allowed tiebreaker.
	'   #7 Division tiebreaker.
	'   #6 Wild Card tiebreaker.
	'--------------------------------------------------------------------------
	sub ApplyConferenceCombinedPointsRankingTiebreaker(first, last)

		dim i

		'Sort teams by their combined ranking in PF and PA (within their own
		'conference). Note that the ranking is negated so the top ranked team
		'will have the highest tb value. This allows them to be sorted like the
		'other tiebreakers.
		for i = first to last
			teams(i).tb = -GetCombinedPointsRank(i, true)
		next
		call SortByTB(first, last)

		'Compare combined points ranking within conference.

		'For two teams...
		if last - first = 1 then
			if teams(first).tb > teams(last).tb then
				teams(first).tbText = "Better combined ranking among conference teams in points scored and points allowed than " & teams(last).name & "."
			else
				call ApplyCombinedPointsRankingTiebreaker(first, last)
			end if

		'For three or more teams...
		else
			if teams(first).tb > teams(first + 1).tb then
				call ApplyHeadToHeadTiebreaker(first + 1, last)
				teams(first).tbText = "Better combined ranking among conference teams in points scored and points allowed than " & ListTeamNames(first + 1, last) & "."
			elseif teams(first).tb = teams(first + 1).tb and _
			       teams(first).tb > teams(first + 2).tb then
				call ApplyHeadToHeadTiebreaker(first, first + 1)
				call ApplyHeadToHeadTiebreaker(first + 2, last)
			else
				call ApplyCombinedPointsRankingTiebreaker(first, last)
			end if

		end if

	end sub

	'--------------------------------------------------------------------------
	' Combined ranking among all teams in points scored and points allowed
	' tiebreaker.
	'   #8 Division tiebreaker.
	'   #7 Wild Card tiebreaker.
	'--------------------------------------------------------------------------
	sub ApplyCombinedPointsRankingTiebreaker(first, last)

		dim i

		'Sort teams by their combined ranking in PF and PA (among all teams).
		'Again, the ranking is negated for sorting purposes.
		for i = first to last
			teams(i).tb = -GetCombinedPointsRank(i, false)
		next
		call SortByTB(first, last)

		'Compare combined points ranking among all teams.

		'For two teams...
		if last - first = 1 then
			if teams(first).tb > teams(last).tb then
				teams(first).tbText = "Better combined ranking among all teams in points scored and points allowed than " & teams(last).name & "."
			else
			if InSameDivision(first, last) then
				call ApplyCommonGamesNetPointsTiebreaker(first, last)
			else
				call ApplyConferenceNetPointsTiebreaker(first, last)
			end if
		end if

		'For three or more teams...
		else
			if teams(first).tb > teams(first + 1).tb then
				call ApplyHeadToHeadTiebreaker(first + 1, last)
				teams(first).tbText = "Better combined ranking among all teams in points scored and points allowed than " & ListTeamNames(first + 1, last) & "."
			elseif teams(first).tb = teams(first + 1).tb and _
			       teams(first).tb > teams(first + 2).tb then
				call ApplyHeadToHeadTiebreaker(first, first + 1)
				call ApplyHeadToHeadTiebreaker(first + 2, last)
			else
				if InSameDivision(first, last) then
					call ApplyCommonGamesNetPointsTiebreaker(first, last)
				else
					call ApplyConferenceNetPointsTiebreaker(first, last)
				end if
			end if

		end if

	end sub

	'--------------------------------------------------------------------------
	' Net points in common games tiebreaker.
	'   #9 Division tiebreaker.
	'--------------------------------------------------------------------------
	sub ApplyCommonGamesNetPointsTiebreaker(first, last)

		dim i, j, k
		dim list, c
		dim found

		'Find common games.
		list = ""
		for i = 0 to UBound(teams)
			found = true
			for j = first to last
				c = CountMatches(teams(j).allOpponents, teams(i).id)
				if c = 0 then
					found = false
				end if
			next
			if found then
				if list = "" then
					list = teams(i).id
				else
					list = list & "," & teams(i).id
				end if
			end if
		next
		list = Split(list, ",")

		'Find each team's net points in those common games and sort.
		for i = first to last
			teams(i).tb = 0
			for j = 0 to UBound(teams(i).allOpponents)
				for k = 0 to UBound(list)
					if teams(i).allOpponents(j) = list(k) then
						teams(i).tb = teams(i).tb + teams(i).allOppNetPts(j)
					end if
				next
			next
		next
		call SortByTB(first, last)

		'Compare net points in common games.

		'For two teams...
		if last - first = 1 then
			if teams(first).tb > teams(last).tb then
				teams(first).tbText = "Better net points in common games" _
					& " (" & teams(first).tb _
					& " vs. " & ListCommonOpponentNames(list) & ")" _
					& " than " & teams(last).name _
					& " (" & teams(last).tb & ")" _
					& "."
			else
				call ApplyNetPointsTiebreaker(first, last)
			end if

		'For three or more teams...
		else
			if teams(first).tb > teams(first + 1).tb then
				call ApplyHeadToHeadTiebreaker(first + 1, last)
				teams(first).tbText = "Better net points in common games" _
					& " (" & teams(first).tb _
					& " vs. " & ListCommonOpponentNames(list) & ")" _
					& " than "
				for i = first + 1 to last
					teams(first).tbText = teams(first).tbText & teams(i).name _
						& " (" & teams(i).tb & ")" _
						& ", "
				next
				teams(first).tbText = Left(teams(first).tbText, Len(teams(first).tbText) - 2) & "."
			elseif teams(first).tb = teams(first + 1).tb and _
			       teams(first).tb > teams(first + 2).tb then
				call ApplyHeadToHeadTiebreaker(first, first + 1)
				call ApplyHeadToHeadTiebreaker(first + 2, last)
			else
				call ApplyNetPointsTiebreaker(first, last)
			end if

		end if

	end sub

	'--------------------------------------------------------------------------
	' Net points in conference games tiebreaker.
	'   #8 Wild Card tiebreaker.
	'--------------------------------------------------------------------------
	sub ApplyConferenceNetPointsTiebreaker(first, last)

		dim i

		'Sort teams by net points in conference games.
		for i = first to last
			teams(i).tb = teams(i).confNetPts
		next
		call SortByTB(first, last)

		'Compare net points in conference games.

		'For two teams...
		if last - first = 1 then
			if teams(first).tb > teams(last).tb then
				teams(first).tbText = "Better net points in conference games than " & teams(last).name & "."
			else
				call ApplyNetPointsTiebreaker(first, last)
			end if

		'For three or more teams...
		else
			if teams(first).tb > teams(first + 1).tb then
				call ApplyHeadToHeadTiebreaker(first + 1, last)
					teams(first).tbText = "Better net points in conference games than " & ListTeamNames(first + 1, last) & "."
			elseif teams(first).tb = teams(first + 1).tb and _
			       teams(first).tb > teams(first + 2).tb then
				call ApplyHeadToHeadTiebreaker(first, first + 1)
				call ApplyHeadToHeadTiebreaker(first + 2, last)
			else
				call ApplyNetPointsTiebreaker(first, last)
			end if

		end if

	end sub

	'--------------------------------------------------------------------------
	' Net points in all games tiebreaker.
	'   #10 Division tiebreaker.
	'   #9 Wild Card tiebreaker.
    '
	' Note: there are two additional tiebreakers after this (net touchdowns and
	' coin toss) but since that data is not available, teams will be sorted by
	' name instead.
	'--------------------------------------------------------------------------
	sub ApplyNetPointsTiebreaker(first, last)

		dim i

		'Sort teams by net points in all games.
		for i = first to last
			teams(i).tb = teams(i).netPts
		next
		call SortByTB(first, last)

		'Compare net points in all games.

		'For two teams...
		if last - first = 1 then
			if teams(first).tb > teams(last).tb then
				teams(first).tbText = "Better net points in all games than " & teams(last).name & "."
			else
				for i = first to last
					teams(i).tbText = "Unable to determine tiebreaker."
				next
				call SortByName(first, last)
			end if

		'For three or more teams...
		else
			if teams(first).tb > teams(first + 1).tb then
				call ApplyHeadToHeadTiebreaker(first + 1, last)
				teams(first).tbText = "Better net points in all games than " & ListTeamNames(first + 1, last) & "."
			elseif teams(first).tb = teams(first + 1).tb and _
			       teams(first).tb > teams(first + 2).tb then
				call ApplyHeadToHeadTiebreaker(first, first + 1)
				call ApplyHeadToHeadTiebreaker(first + 2, last)
			else
				for i = first to last
					teams(i).tbText = "Unable to determine tiebreaker."
				next
				call SortByName(first, last)
			end if

		end if

	end sub

	'**************************************************************************
	'* Helper functions and subroutines.                                      *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Calculates a winning percentage given a number of wins, losses and ties
	' (counting ties as 1/2 a win).
	'--------------------------------------------------------------------------
	function CalcPct(w, l, t)

		dim n

		CalcPct = 0
		n = w + l + t
		if n > 0 then
			CalcPct = (w + t / 2) / n
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns the number of occurences of value 'x' in array 'list'. Used for
	' head-to-head and common games tiebreakers.
	'--------------------------------------------------------------------------
	function CountMatches(list, x)

		dim i

		CountMatches = 0
		for i = 0 to UBound(list)
			if list(i) = x then
				CountMatches = CountMatches + 1
			end if
		next

	end function

	'--------------------------------------------------------------------------
	' Returns the combined ranking in points scored and points allowed by the
	' given team. That is, the team's rank among other teams in PF plus its
	' rank in PA. If 'ownConf' is true, only teams in the same conference are
	' considered.
	'--------------------------------------------------------------------------
	function GetCombinedPointsRank(i, ownConf)

		dim conf
		dim pf, pa
		dim j
		dim pfRank, paRank

		GetCombinedPointsRank = -1

		'Get the team's conference.
		conf = teams(i).conference

		'Assume this team is the best in each category.
		pfRank = 1 : paRank = 1

		'Compare this team's PF and PA totals to each of the other teams.
		'Whenever another team has a better point total, bump this team's rank
		'down by one.
		for j = 0 to UBound(teams)
			if i <> j and (not ownConf or teams(j).conference = conf) then
				if teams(j).pointsFor > teams(i).pointsFor then
					pfRank = pfRank + 1
				end if
				if teams(j).pointsAgainst < teams(i).pointsAgainst then
					paRank = paRank + 1
				end if
			end if
		next
		GetCombinedPointsRank = pfRank + paRank

	end function

	'--------------------------------------------------------------------------
	' Returns the combined winning percentage of all the given team's
	' opponents.
	'--------------------------------------------------------------------------
	function GetStrengthOfSchedule(i)

		dim j, k
		dim w, l, t
		dim id

		GetStrengthOfSchedule = 0
		w = 0 : l = 0 : t = 0

		'Check all teams played.
		for j = 0 to UBound(teams(i).allOpponents)

			'Get the opponent.
			id = teams(i).allOpponents(j)
			k = GetTeamIndex(id)

			'Add the opponent's record to the totals. Note that an opponent may
			'be added more than once.
			w = w + teams(k).totalWins
			l = l + teams(k).totalLosses
			t = t + teams(k).totalTies
		next

		GetStrengthOfSchedule = CalcPct(w, l, t)

	end function

	'--------------------------------------------------------------------------
	' Returns the combined winning percentage of all teams beaten by the given
	' team.
	'--------------------------------------------------------------------------
	function GetStrengthOfVictory(i)

		dim j, k
		dim w, l, t
		dim id

		GetStrengthOfVictory = 0
		w = 0 : l = 0 : t = 0

		'Check only teams that have been beaten.
		for j = 0 to UBound(teams(i).oppWonAgainst)

			'Get the opponent.
			id = teams(i).oppWonAgainst(j)
			k = GetTeamIndex(id)

			'Add the opponent's record to the totals. Note that an opponent may
			'be added more than once.
			w = w + teams(k).totalWins
			l = l + teams(k).totalLosses
			t = t + teams(k).totalTies
		next
		GetStrengthOfVictory = CalcPct(w, l, t)

	end function

	'--------------------------------------------------------------------------
	' Searches the global list of teams and returns the index of the team with
	' the given ID.
	'--------------------------------------------------------------------------
	function GetTeamIndex(id)

		dim i

		GetTeamIndex = -1
		for i = 0 to UBound(teams)
			if teams(i).id = id then
				GetTeamIndex = i
				exit function
			end if
		next

	end function

	'--------------------------------------------------------------------------
	' Returns true if all the teams in a given range are in the same division.
	'--------------------------------------------------------------------------
	function InSameDivision(first, last)

		dim i, div

		InSameDivision = true
		div = teams(first).division
		for i = first + 1 to last
			if teams(i).division <> div then
				InSameDivision = false
				exit function
			end if
		next

	end function

	'--------------------------------------------------------------------------
	' Given an array of team IDs, returns a string listing their names. Used
	' for building the tiebreaker notes when the tiebreaker involves common
	' games.
	'--------------------------------------------------------------------------
	function ListCommonOpponentNames(list)

		dim n
		dim i, j
		dim names, tmp

		'Create a new list.
		n = UBound(list)
		redim names(n)

		'Get the team names.
		ListCommonOpponentNames = ""
		for i = 0 to UBound(list)
			names(i) = teams(GetTeamIndex(list(i))).name
		next

		'Sort the names.
		for i = 0 to UBound(names) - 1
			for j = i + 1 to UBound(names)
				if names(i) > names(j) then
					tmp = names(i)
					names(i) = names(j)
					names(j) = tmp
				end if
			next
		next

		'Build the string.
		ListCommonOpponentNames = Join(names, ", ")

   end function

	'---------------------------------------------------------------------------
	' Given a range within the global team list, returns a string listing their
	' names. Used for building the tiebreaker notes.
	'---------------------------------------------------------------------------
	function ListTeamNames(first, last)

		dim names, i

		ListTeamNames = ""

		'Sort the sub list of teams.
		call SortByName(first, last)

		'Create the list of names.
		redim names (last - first)
		for i = 0 to UBound(names)
			names(i) = teams(first + i).name
		next

		'Build the string.
		ListTeamNames = Join(names, ", ")

	end function

	'--------------------------------------------------------------------------
	' Sorts the global list of teams within the given range based the team
	' name.
	'--------------------------------------------------------------------------
	sub SortByName(first, last)

		dim i, j, tmp

		if last <= first then
			exit sub
		end if

		for i = first to last - 1
			for j = i + 1 to last
				if teams(j).name < teams(i).name then
					set tmp = teams(i)
					set teams(i) = teams(j)
					set teams(j) = tmp
				end if
			next
		next

	end sub

	'--------------------------------------------------------------------------
	' Sorts the global list of teams within the given range based on whatever
	' value is currently in their tiebreaker field.
	'--------------------------------------------------------------------------
	sub SortByTB(first, last)

		dim i, j, tmp

		if last <= first then
			exit sub
		end if

		'Sort the teams by tiebreaker value.
		for i = first to last - 1
			for j = i + 1 to last
				if teams(j).tb > teams(i).tb then
					set tmp = teams(i)
					set teams(i) = teams(j)
					set teams(j) = tmp
				end if
			next
		next

	end sub

	'--------------------------------------------------------------------------
	' Sorts the global team object array within the given range so that the
	' teams ranking highest in their division are at the top of the list. The
	' index of the last team in that group is returned. Used when determining
	' the Wild Card tiebreaker between three or more teams in the conference
	' standings.
	'--------------------------------------------------------------------------
	function SortDivisionLeaders(first, last)

		dim i, j, tmp
		dim n, div, found

		SortDivisionLeaders = last

		'Sort the teams by division and rank.
		for i = first to last - 1
			for j = i + 1 to last
				if (teams(j).division < teams(i).division) or _
				   (teams(j).division = teams(i).division  and teams(j).divRank < teams(i).divRank) then
					set tmp = teams(i)
					set teams(i) = teams(j)
					set teams(j) = tmp
				end if
			next
		next

		'Move the top ranked team in each division to the top.
		n = first + 1
		for div = teams(n).division + 1 to UBound(DivisionNames) + 1
			i = n + 1
			found = false
			do while i <= last and not found
				if CStr(teams(i).division) = CStr(div) then
					set tmp = teams(n)
					set teams(n) = teams(i)
					set teams(i) = tmp
					n = n + 1
					found = true
				else
					i = i + 1
				end if
			loop
		next

		'Return the index of the last division-leading team.
		i = first
		found = false
		do while i < last - 1 and not found
			if teams(i).division >= teams(i + 1).division then
				SortDivisionLeaders = i
				found = true
			else
				i = i + 1
			end if
		loop

	end function

	'**************************************************************************
	'* Local class definitions.                                               *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' TeamObj: Holds information for a single team.
	'--------------------------------------------------------------------------
	class TeamObj

		public id, name, division, conference

		public totalWins, totalLosses, totalTies
		public atsWins,   atsLosses,   atsTies
		public homeWins,  homeLosses,  homeTies
		public roadWins,  roadLosses,  roadTies
		public afcWins,   afcLosses,   afcTies
		public nfcWins,   nfcLosses,   nfcTies
		public confWins,  confLosses,  confTies
		public divWins,   divLosses,   divTies

	    public totalPct, atsPct, confPct, divPct

		public streak

		public pointsFor, pointsAgainst

		public netPts, confNetPts

		public allOpponents
		public allOppNetPts
		public oppWonAgainst, oppLostTo, oppTiedWith

		public divRank

		public tb, tb1, tb2, tb3, tbText

		private sub Class_Initialize()

			totalWins = 0 : totalLosses = 0 : totalTies = 0
			atsWins   = 0 : atsLosses   = 0 : atsTies   = 0
			homeWins  = 0 : homeLosses  = 0 : homeTies  = 0
			roadWins  = 0 : roadLosses  = 0 : roadTies  = 0
			afcWins   = 0 : afcLosses   = 0 : afcTies   = 0
			nfcWins   = 0 : nfcLosses   = 0 : nfcTies   = 0
			divWins   = 0 : divLosses   = 0 : divTies   = 0

			totalPct = 0 : confPct = 0 : divPct = 0

			streak = ""

			pointsFor = 0 : pointsAgainst = 0

			netPts = 0 : confNetPts = 0

			allOpponents = Array() : allOppNetPts = Array()
			oppWonAgainst = Array() : oppLostTo = Array() : oppTiedWith = Array()

			divRank = 0

			tb = 0
			tb1 = "" : tb2 = "" : tb3 = ""
			tbText = ""

		end sub

		private sub Class_Terminate()
		end sub

		public sub setData(teamID)

			dim sql, rs
			dim opponentID
			dim ptDiff
			dim i
			dim currentStreak

			'Set the team properties.
			id = teamID

			'Get team conference, division and display name.
			sql = "SELECT * FROM Teams WHERE TeamID = '" & teamID & "'"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				conference = rs.Fields("Conference").Value
				division = rs.Fields("Division").Value
				if rs.Fields("DisplayName") <> "" then
					name = rs.Fields("DisplayName").Value
				else
					name = rs.Fields("City").Value
				end if
			end if

			'Find records and point totals for all games played by this team up
			'to the specified week.
			sql = "SELECT * FROM Teams, Schedule" _
			   & " WHERE NOT ISNULL(Result)" _
			   & " AND" _
			   & " Week <= " & week _
			   & " AND" _
			   & " ((VisitorID = '" & teamID & "' and TeamID = HomeID)" _
			   & "  OR" _
			   & "  (HomeID    = '" & teamID & "' and TeamID = VisitorID))" _
			   & " ORDER BY Schedule.Week"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				currentStreak = ""
				do while not rs.EOF

					'Get the opponent team id and add to the team's opponents list.
					'Also, add net points to the parallel array.
					if rs.Fields("HomeID").Value = teamID then
						opponentID = rs.Fields("VisitorID").Value
						ptDiff = rs.Fields("HomeScore") - rs.Fields("VisitorScore")
					else
						opponentID = rs.Fields("HomeID").Value
						ptDiff = rs.Fields("VisitorScore") - rs.Fields("HomeScore")
					end if
					redim preserve allOpponents(Ubound(allOpponents) + 1)
					allOpponents(Ubound(allOpponents)) = opponentID
					redim preserve allOppNetPts(Ubound(allOppNetPts) + 1)
					allOppNetPts(Ubound(allOppNetPts)) = netPts
	
					'Add scores to points for and against totals.
					if rs.Fields("VisitorID").Value = teamID then
						pointsFor = pointsFor + rs.Fields("VisitorScore").Value
						pointsAgainst = pointsAgainst + rs.Fields("HomeScore").Value
					else
						pointsFor = pointsFor + rs.Fields("HomeScore").Value
						pointsAgainst = pointsAgainst + rs.Fields("VisitorScore").Value
					end if

					'Add to the conference and overall net point totals.
					if rs.Fields("Conference").Value = conference then
						confNetPts = confNetPts + ptDiff
					end if
					netPts = netPts + ptDiff

					'Record a win.
					if rs.Fields("Result").Value = teamID then
						totalWins = totalWins + 1
						if currentStreak <> "W" then
							currentStreak = "W"
							streak = 1
						else
							streak = streak + 1
						end if
						if rs.Fields("HomeID").Value = teamID then
							homeWins = homeWins + 1
						else
							roadWins = roadWins + 1
						end if
						if rs.Fields("Conference").Value = 1 then
							afcWins = afcWins + 1
						else
							nfcWins = nfcWins + 1
						end if
						if rs.Fields("Conference").Value = conference and _
							rs.Fields("Division").Value = division then
							divWins = divWins + 1
						end if

					'Record a loss.
					elseif rs.Fields("Result").Value = opponentID then
						totalLosses = totalLosses + 1
						if currentStreak <> "L" then
							currentStreak = "L"
							streak = 1
						else
							streak = streak + 1
						end if
						if rs.Fields("HomeID").Value = teamID then
							homeLosses = homeLosses + 1
						else
							roadLosses = roadLosses + 1
						end if
						if rs.Fields("Conference").Value = 1 then
							afcLosses = afcLosses + 1
						else
							nfcLosses = nfcLosses + 1
						end if
						if rs.Fields("Conference").Value = conference and _
							rs.Fields("Division").Value = division then
							divLosses = divLosses + 1
						end if

					'Record a tie.
					else
						totalTies = totalTies + 1
						if currentStreak <> "T" then
							currentStreak = "T"
							streak = 1
						else
							streak = streak + 1
						end if
						if rs.Fields("HomeID").Value = teamID then
							homeTies= homeTies + 1
						else
							roadTies = roadTies + 1
						end if
						if rs.Fields("Conference").Value = 1 then
							afcTies = afcTies + 1
						else
							nfcTies = nfcTies + 1
						end if
						if rs.Fields("Conference").Value = conference and _
							rs.Fields("Division").Value = division then
							divTies = divTies + 1
						end if
					end if

					'Record a win, loss or tie vs. the spread.
					if rs.Fields("ATSResult").Value = teamID then
						atsWins = atsWins + 1
					elseif rs.Fields("ATSResult").Value = opponentID then
						atsLosses = atsLosses + 1
					else
						atsTies = atsTies + 1
					end if

					'Add opponent to proper W-L-T list (may be added more than once).
					if currentStreak = "W" then
						redim preserve oppWonAgainst(Ubound(oppWonAgainst) + 1)
						oppWonAgainst(Ubound(oppWonAgainst)) = opponentID
					elseif currentStreak = "L" then
						redim preserve oppLostTo(Ubound(oppLostTo) + 1)
						oppLostTo(Ubound(oppLostTo)) = opponentID
					elseif currentStreak = "T" then
						redim preserve oppTiedWith(Ubound(oppTiedWith) + 1)
						oppTiedWith(Ubound(oppTiedWith)) = opponentID
					end if

					rs.MoveNext
				loop
			end if

			'Set conference records.
			if conference = 1 then
				confWins   = afcWins
				confLosses = afcLosses
				confTies   = afcTies
			else
				confWins   = nfcWins
				confLosses = nfcLosses
				confTies   = nfcTies
			end if

			'Set streak.
			streak = currentStreak & streak

			'Set win-loss-tie percentages.
			totalPct = CalcPct(totalWins, totalLosses, totalTies)
			atsPct   = CalcPct(atsWins,   atsLosses,   atsTies)
			confPct  = CalcPct(confWins,  confLosses,  confTies)
			divPct   = CalcPct(divWins,   divLosses,   divTies)

		end sub

	end class %>
