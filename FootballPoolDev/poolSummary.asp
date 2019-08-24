<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Summary" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<!--<link rel="stylesheet" type="text/css" href="styles/common.css" />-->
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
    <link href="styles/style.css" rel="stylesheet" type="text/css" />
    
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
	<script type="text/javascript" src="scripts/tableSort.js"></script>
	<script type="text/javascript">//<![CDATA[
	//-------------------------------------------------------------------------
	// Function for sorting columns in the Individual Statistics table.
	//-------------------------------------------------------------------------
	function sortStats(colName)
	{
		// Get the table or table section to sort.
		var tblEl = document.getElementById("indvStats");

		// Set up some sorting parameters based on column name given.
  		var col;                     // index of the primary column to sort on.
		var rev;                     // initial sort direction of that column.
  		var sortCols = new Array();  // column(s) to be highlighted.
		var hdrCols  = new Array();  // header column(s) to be highlighted.

		switch (colName)
		{
			case 'Name':
				col = 0;
				rev = false;
				break;

			case 'Played':
				col = 1;
				rev = true;
				sortCols[0] = col;
				hdrCols[0]  = col;
				break;

			case 'Won':
				col = 2;
				rev = true;
				sortCols[0] = col;
				hdrCols[0]  = col;
				break;

			case 'Winnings':
				col = 4;
				rev = true;
				sortCols[0] = col;
				hdrCols[0]  = col;
				break;

			case 'Net':
				col = 5;
				rev = true;
				sortCols[0] = col;
				hdrCols[0]  = col;
				break;

			case 'Picks':
				col = 4;
				rev = true;
				sortCols[0] = col - 1; sortCols[1] = col;
				hdrCols[0]  = col - 1;
				break;
<%	if USE_CONFIDENCE_POINTS then %>
			case 'Points':
				col = 8;
				rev = true;
				sortCols[0] = col;
				hdrCols[0]  = col - 1;
				break;
<%	end if %>
			default:
				return false;
				break;
		}

		// Set the sort direction.
		setSortDirection(tblEl, colName, rev, "Name");

		// Do the sort.
		var tmpEl;
		var i, j;
		var minVal, minIdx;
		var testVal;
		var cmp;

		for (i = 0; i < tblEl.rows.length - 1; i++)
		{
			// Assume the current row has the minimum value.
			minIdx = i;
			minVal = getTextValue(tblEl.rows[i].cells[col]);

			// Search the rows that follow the current one for a smaller value.
			for (j = i + 1; j < tblEl.rows.length; j++)
			{
				testVal = getTextValue(tblEl.rows[j].cells[col]);
				cmp = compareValues(minVal, testVal);

				// Negate the comparison result if the reverse sort flag is
				// set.
				if (tblEl.reverseSort[colName])
					cmp = -cmp;

				// Use the 'Name' column as a secondary sort if those values
				// are equal.
				if (cmp == 0 && col != 0)
					cmp = compareValues(getTextValue(tblEl.rows[minIdx].cells[0]),
					                    getTextValue(tblEl.rows[j].cells[0]));

				// If this row has a smaller value than the current minimum,
				// remember its position and update the current minimum value.
				if (cmp > 0)
				{
					minIdx = j;
					minVal = testVal;
				}
			}

			// By now, we have the row with the smallest value. Remove it from
			// the table and insert it before the current row.
			if (minIdx > i)
			{
				tmpEl = tblEl.removeChild(tblEl.rows[minIdx]);
				tblEl.insertBefore(tmpEl, tblEl.rows[i]);
			}
		}

		// Fix the table's appearance.
		restyleTable(tblEl, sortCols, hdrCols);

		return false;
	}
	//]]></script>
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/weekly.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'Get the current date and time.
	dim dateNow
	dateNow = CurrentDateTime()

	'Determine the maximum week to check.
	dim maxWeek
	maxWeek = CurrentWeek()
	if dateNow <= WeekStartDateTime(maxWeek) then
		maxWeek = maxWeek - 1
	end if

	'Create an array of user objects (need to build this first because we will
	'use the weekly pool results to add up each user's winnings).
	dim users, list, i
	list = UsersList(true)
	if IsArray(list) then
		redim users(UBound(list))
		for i = 0 to UBound(list)
			set users(i) = new UserObj
			users(i).setData(list(i))
		next
	end if

	'Build the weekly winners display. %>
	<h2>Weekly Winners</h2>
	<table class="main" cellpadding="0" cellspacing="0">
		<tr class="header bottomEdge">
			<th align="right">Week</th>
			<th align="left">Winner(s)</th>
			<th align="center" colspan="2">Score</th>
			<th align="right">Tiebreaker</th>
			<th align="right">Actual</th>
			<th align="right">(Diff.)</th>
			<th align="right">Players</th>
			<!--<th align="right">Pot</th> -->
		</tr>
<%	'Determine the latest week we have complete results for.
	dim maxCompWeek
	maxCompWeek = maxWeek
	if NumberOfGames(maxCompWeek) <> NumberOfCompletedGames(maxCompWeek) then
		maxCompWeek = maxCompWeek - 1
	end if

	'If there are no results to show, let the user know.
	if maxCompWeek < 1 then %>
		<tr><td align="center" colspan="9"><em>No results available.</em></td></tr>
<%	end if

	'Get the pool results for each week.
	dim week, alt
	dim j, n
	dim winners, score, numGames, tb1, tb2, actual1, actual2, diff, players, pot
	dim scoreStr, pctStr 
	alt = false
	for week = 1 to maxCompWeek
		list = WinnersList(week)
		winners = "n/a"
		score = ""
		tb1 = "n/a"
		tb2 = "n/a"
		actual1 = TBPointTotalVis(week)
		actual2 = TBPointTotalHome(week)
		diff = "&nbsp;"
		players = 0
		pot = 0
		
		if IsArray(list) then
			winners = Join(list, ", ")
			players = NumberOfEntries(week)
			numGames = NumberOfGames(week)
			if USE_CONFIDENCE_POINTS then
				score = PlayerConfidenceScore(list(0), week)
			else
				score = PlayerPickScore(list(0), week)
			end if
			'call errormessage(score & week & list(0))
			pot = players * BET_AMOUNT

			'If there was a tie for highest score, show tiebreaker data.
			if TiedScoreCount(week, score) > 1 then
				tb1 = TBvisGuess(list(0), week)
				tb2 = TBhomeGuess(list(0), week)
				diff = " (" & Abs(tb1 - actual1) + Abs(tb2 - actual2) & ")"
				for i = 1 to UBound(list)
					tb1 = tb1 & ", " & TBvisGuess(list(i), week)
					tb2 = tb2 & ", " & TBhomeGuess(list(i), week)
					
				next
			end if
			
			
			'Update the win totals for each user in the winners list.
			for i = 0 to UBound(list)
				for j = 0 to UBound(users)
					n = Ubound(list) + 1
					if users(j).name = list(i) then
						users(j).poolsWon = users(j).poolsWon + 1 / n
						users(j).totalWinnings = users(j).totalWinnings + pot / n
						exit for
					end if
				next
			next

		end if

		'Format the score display.
		if USE_CONFIDENCE_POINTS then
			scoreStr = FormatScore(score, true)
		else
			if IsNumeric(score) and numGames > 0 then
				scoreStr = score & "/" & numGames
				pctStr   = "(" & FormatPercentage(score / numGames) & ")"
			else
				scoreStr = "n/a"
				pctStr = ""
			end if
		end if

		if alt then %>
		<tr align="right" class="alt singleLine" valign="top">
<%		else %>
		<tr align="right" class="singleLine" valign="top">
<%		end if
		alt = not alt %>
			<td><a href="poolResults.asp?week=<% = week %>"><% = week %></a></td>
			<td align="left"><% = winners %></td>
<%		if USE_CONFIDENCE_POINTS then %>
			<td colspan="2"><% = scoreStr %></td>
<%		else %>
			<td><% = scoreStr %></td>
			<td><% = pctStr %></td>
<%		end if %>
			<td><% = tb1 %>/<% = tb2 %></td>
			<td><% = actual1 %>/<% = actual2 %></td>
			<td><% = diff %></td>
			<td><% = players %></td>
			<!--<td><% = FormatAmount(pot) %></td> -->
		</tr>
<%	next %>
	</table>
    <p>&nbsp;</p>
    <p>
      <%	'Build the individual statistics display.
	dim cost, net
	dim cols
	cols =  8
	if USE_CONFIDENCE_POINTS then
		cols = cols + 1
	end if %>
    </p>
    <h2>Individual Statistics</h2>
	<table class="main" cellpadding="0" cellspacing="0">
		<thead>
			<tr class="header bottomEdge singleLine sortable">
				<th align="left"><a href="#" onclick="this.blur(); return sortStats('Name');" title="Sort by name.">Name</a></th>
				<th align="right"><a href="#" onclick="this.blur(); return sortStats('Played');" title="Sort by pools played.">Played</a></th>
				<th align="right"><a href="#" onclick="this.blur(); return sortStats('Won');" title="Sort by pools won.">Won</a></th>
				<!--<th align="center">Cost</th>
				<th align="right"><a href="#" onclick="this.blur(); return sortStats('Winnings');" title="Sort by gross winnings.">Winnings</a></th> -->
				<!--<th align="center"><a href="#" onclick="this.blur(); return sortStats('Net');" title="Sort by net winnings.">Net</a></th> -->
				<th align="center" colspan="2"><a href="#" onclick="this.blur(); return sortStats('Picks');" title="Sort by overall pick percentage.">Overall Picks</a></th>
<%	if USE_CONFIDENCE_POINTS then %>
				<th align="center"><a href="#" onclick="this.blur(); return sortStats('Points');" title="Sort by total points.">Total Pts.</a></th>
<%	end if %>
			</tr>
		</thead>
		<tbody id="indvStats">
<%	'If there are no players to show, let the user know.
	if not IsArray(users) then %>
			<tr><td align="center" colspan="<% = cols %>"><em>No players found.</em></td></tr>
<%	else
		dim decimalPlaces
		decimalPlaces = GetMaxDecimalPlaces()
		for i = 0 to UBound(users)
			if Round(i / 2) * 2 = i then %>
			<tr align="right" class="singleLine" valign="top">
<%			else %>
			<tr align="right" class="alt singleLine" valign="top">
<%			end if
			cost = users(i).poolsPlayed * BET_AMOUNT
			net = users(i).totalWinnings - cost
			if users(i).totalGames <> 0 then
				pctStr = "(" & FormatPercentage(users(i).totalCorrect / users(i).totalGames) & ")"
			else
				pctStr = "(" & FormatPercentage(0) & ")"
			end if %>
				<td align="left"><% = users(i).name %></td>
				<td><% = users(i).poolsPlayed %></td>
				<td><% = FormatNumber(users(i).poolsWon, decimalPlaces, true) %></td>
<!--				<td><% = FormatAmount(cost) %></td>
				<td><% = FormatAmount(users(i).totalWinnings) %></td>
				<td><% = FormatAmount(net) %></td> -->
				<td><% = users(i).totalCorrect & "/" & users(i).totalGames %></td>
				<td><% = pctStr %></td>
<%			if USE_CONFIDENCE_POINTS then %>
				<td><% = FormatScore(users(i).totalScore, true) %></td>
<%			end if %>
			</tr>
<%		next
	end if %>
		</tbody>
	</table>
	</td></tr></table>

<!-- #include file="includes/footer.asp" -->
</body>
</html>
<%	'**************************************************************************
	'* Local functions and subroutines.                                       *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Returns the number of players who had the given score in the given week.
	'--------------------------------------------------------------------------
	function TiedScoreCount(week, score)

		dim scoreField, sql, rs

		TiedScoreCount = 0

		scoreField = "PickScore"
		if USE_CONFIDENCE_POINTS then
			scoreField = "ConfidenceScore"
		end if
		sql = "SELECT COUNT(*) AS Total FROM Tiebreaker" _
		   & " WHERE Week = " & week _
		   & " AND " & scoreField & " = " & score
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			TiedScoreCount = rs.Fields("Total").Value
		end if

	end function

	'--------------------------------------------------------------------------
	' Determines how many decimal places to show in the pools won column.
	'--------------------------------------------------------------------------
	function GetMaxDecimalPlaces()

		dim i, n, d

		GetMaxDecimalPlaces = 0

		for i = 0 to UBound(users)
			n = users(i).poolsWon - Round(users(i).poolsWon - .5)
			d = Len(n) - 2
			if d > GetMaxDecimalPlaces then
				GetMaxDecimalPlaces = d
				if GetMaxDecimalPlaces > 3 then
					GetMaxDecimalPlaces = 3
					exit function
				end if
			end if
		next

	end function

	'**************************************************************************
	'* Local class definitions.                                               *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' UserObj: Holds information for a single user.
	'--------------------------------------------------------------------------
	class UserObj

		public name
		public poolsPlayed
		public poolsWon
		public totalCorrect
		public totalGames
		public totalScore
		public totalWinnings

		private sub Class_Initialize()
		end sub

		private sub Class_Terminate()
		end sub

		public sub setData(username)

			dim sql, rs
			dim i, n
			dim pickScore, confScore

			'Set the user properties.
			name = username

			'Find the number of weeks this user has participated in so far.
			'Note that the current week is included only after the first game
			'has started.
			poolsPlayed = 0
			sql = "SELECT COUNT(*) AS Total" _
			   & " FROM Tiebreaker" _
			   & " WHERE Username = '" & SqlString(username) & "'" _
			   & " AND Week <= " & maxWeek
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				poolsPlayed = rs.Fields("Total").Value
			end if

			'Get the overall pick stats and score data for this user.
			totalCorrect = 0
			totalGames   = 0
			totalScore   = 0
			sql = "SELECT * FROM Tiebreaker" _
			   & " WHERE Username = '" & SqlString(username) & "'" _
			   & " ORDER BY Week"
			set rs = DbConn.Execute(sql)
			do while not rs.EOF
				week      = rs.Fields("Week").Value
				pickScore = rs.Fields("PickScore").Value
				if IsNull(pickScore) then
					pickScore = PlayerPickScore(username, week)
				end if
				if IsNumeric(pickScore) then
					totalCorrect = totalCorrect + pickScore
				end if
				if USE_CONFIDENCE_POINTS then
					confScore = rs.Fields("ConfidenceScore").Value
					if IsNull(confScore) then
						confScore = PlayerConfidenceScore(username, week)
					end if
					if IsNumeric(confScore) then
						totalScore = totalScore + confScore
					end if
				end if
				totalGames = totalGames + NumberOfCompletedGames(week)
				rs.MoveNext
			loop

			'Initialize remaining properties.
			poolsWon = 0
			totalWinnings = 0
		end sub

	end class %>
