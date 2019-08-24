<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Playoffs Results" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/common.css" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
	<script type="text/javascript" src="scripts/tableSort.js"></script>
	<script type="text/javascript">//<![CDATA[
	//-------------------------------------------------------------------------
	// Function for sorting columns in the Pool Results table.
	//-------------------------------------------------------------------------
	function sortResults(colName)
	{
		// Get the table section to sort.
		var tblEl = document.getElementById("poolResults");

		// Set up some sorting parameters based on column name given.
		var col;                     // index of the primary column to sort on.
		var rev;                     // initial sort direction of that column.
		var sortCols = new Array();  // column(s) to be highlighted.
		var hdrCols	 = new Array();  // header column(s) to be highlighted.

		switch (colName) {

			case 'Name':
				col = 0;
				rev = false;
				break;
<%	if USE_CONFIDENCE_POINTS then %>
			case 'Correct':
				col = tblEl.rows[0].cells.length - 3;
				sortCols[0] = col - 1; sortCols[1] = col;
				hdrCols[0]	= 5;
				rev = true;
				break;
			case 'Score':
				col = tblEl.rows[0].cells.length - 2;
				rev = true;
				sortCols[0] = col; sortCols[1] = col + 1;
				hdrCols[0]	= 6;
				break;
<%	else %>
			case 'Correct':
				col = tblEl.rows[0].cells.length - 2;
				rev = true;
				sortCols[0] = col; sortCols[1] = col + 1;
				hdrCols[0]	= 5;
				break;
			default:
				return false;
				break;
<%	end if %>
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

				// Sort by the 'Name' column if those values are equal.
				if (cmp == 0)
				{
					cmp = compareValues(
						getTextValue(tblEl.rows[minIdx].cells[0]),
						getTextValue(tblEl.rows[j].cells[0]));
				}

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
<!-- #include file="includes/playoffs.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'Find the number of playoff games.
	dim numGames, completedGames
	numGames       = NumberOfPlayoffGames()
	completedGames = NumberOfCompletedPlayoffGames()

	'Get the current username.
	dim username
	username = Session(SESSION_USERNAME_KEY)

	'Determine if all playoff games have been locked.
	dim allLocked
	dim sql, rs
	dim dateTime
	allLocked = false
	sql = "SELECT Date, Time FROM PlayoffsSchedule ORDER By Date DESC, Time DESC"
	set rs = DbConn.Execute(sql)
	if not (rs.BOF and rs.EOF) then
		dateTime = CDate(rs.Fields("Date").Value & " " & rs.Fields("Time").Value)
		if dateTime < CurrentDateTime() then
			allLocked = true
		end if
	end if

	'Build the display.
	dim gameCols, otherCols
	gameCols  = numGames
	otherCols = 3
	if USE_CONFIDENCE_POINTS then
		otherCols = otherCols + 2
	end if %>
	<table class="main" cellpadding="0" cellspacing="0">
		<thead>
			<tr class="header bottomEdge sortable" valign="bottom">
				<th align="left"><a href="#" onclick="this.blur(); return sortResults('Name');" title="Sort by name.">Name</a></th>
				<th colspan="4">Wild Card<br />Games</th>
				<th colspan="4">Divisional<br />Playoffs</th>
				<th colspan="2">Conf.<br /> Champ.</th>
				<th>Super<br />Bowl</th>
				<th colspan="2"><a href="#" onclick="this.blur(); return sortResults('Correct');" title="Sort by number of correct picks.">Correct</a></th>
<%	if USE_CONFIDENCE_POINTS then %>
				<th align="left" colspan="2"><a href="#" onclick="this.blur(); return sortResults('Score');" title="Sort by score.">Score</a></th>
<%	end if %>
			</tr>
		</thead>
<%	'If the pool has been concluded, display the winners and payout.
	dim str, list, payout
	str = ""
	list = PlayoffsWinnersList()
	if IsArray(list) then
		payout = NumberOfPlayoffsEntries() * PLAYOFFS_BET_AMOUNT
		if UBound(list) > 0 then
			str = "Winners: " & Join(list, ", ") & " (" & FormatAmount(payout / (UBound(list) + 1)) & " each)"
		else
			str = "Winner: " & Join(list, ", ") & " (" & FormatAmount(payout) & ")"
		end if %>
		<tfoot>
			<tr class="header topEdge"><th align="left" colspan="<% = gameCols + otherCols %>"><% = str %></th></tr>
		</tfoot>
<%	end if %>
		<tbody id="poolResults">
<%	dim users, currentUser
	dim leadConfScore
	dim i, j
	dim gameRound
	dim pick, conf, result, correctPick, giveTie
	dim pickScore, pickPct, confScore, confScoreDiff, tbDiff
	dim hidePick
	currentUser = Session(SESSION_USERNAME_KEY)

	'Get user data.
	list = PlayoffsPoolPlayersList()
	if IsArray(list) then

		'Create an array of user objects.
		redim users(UBound(list))
		for i = 0 to UBound(list)
			set users(i) = new UserObj
			users(i).setData(list(i))
		next

		'If confidence points are used, get the current leading score.
		if USE_CONFIDENCE_POINTS and completedGames > 0 then
			leadConfScore = 0
			for i = 0 to UBound(users)
				if IsNumeric(users(i).confScore) then
					if users(i).confScore > leadConfScore then
						leadConfScore = users(i).confScore
					end if
				end if
			next
		end if

		dim alt
		alt = false
		for i = 0 to UBound(users)
			if alt then %>
			<tr align="center" class="alt singleLine">
<%			else %>
			<tr align="center" class="singleLine">
<%			end if
			alt = not alt %>
				<td align="left"><% = users(i).name %></td>
<%			'Display picks for this user, highlighting the correct ones.
			sql = "SELECT * FROM PlayoffsPicks, PlayoffsSchedule" _
			   & " WHERE Username = '" & SqlString(users(i).name) & "'" _
			   & " AND PlayoffsPicks.GameID = PlayoffsSchedule.GameID" _
			   & " ORDER BY PlayoffsSchedule.Round, PlayoffsSchedule.Date, PlayoffsSchedule.Time"
			set rs = DbConn.Execute(sql)
			do while not rs.EOF
				pick = rs.Fields("Pick").Value
				conf = rs.Fields("Confidence").Value

				'Determine if the pick should be hidden.
				hidePick = false
				gameRound = rs.Fields("Round").Value
				if not allLocked and not IsAdmin() and currentUser <> users(i).name then
					if not RoundLocked(gameRound) then
						hidePick = true
					end if
				end if

				'Determine the correct pick.
				if USE_POINT_SPREADS then
					correctPick = rs.Fields("ATSResult").Value
				else
					correctPick = rs.Fields("Result").Value
				end if

				'Format the pick display.
				if pick = "" then
					pick = "---"
					conf = "--"
				elseif hidePick then
					pick = "XXX"
					conf = "xx"
				elseif pick = correctPick then
					pick = FormatCorrectPick(pick)
				end if

				'When using point spreads, highlight the straight-up result
				'if it matches the pick.
				if USE_POINT_SPREADS then
					if rs.Fields("Pick").Value = rs.Fields("Result").Value then
						pick = FormatWinner(pick)
					end if
				end if
				if USE_CONFIDENCE_POINTS then %>
				<td class="confPick"><% = pick %><br /><% = conf %></td>
<%				else %>
				<td><% = pick %></td>
<%				end if
					rs.MoveNext
			loop

			'Build the correct pick and score displays.
			pickScore  = "n/a"
			pickPct    = "&nbsp;"
			if completedGames > 0 then
				pickScore = users(i).pickScore
				pickPct = "(" & FormatPercentage(pickScore / completedGames) & ")"
				pickScore = pickScore & "/" & completedGames
			end if

			'For confidence points, show how far behind the leader this
			'user's score is.
			if USE_CONFIDENCE_POINTS then
				confScore = "n/a&nbsp;&nbsp;&nbsp;"
				confScoreDiff = "&nbsp;"
				if completedGames > 0 then
					confScore = users(i).confScore
					if confScore < leadConfScore then
						confScoreDiff = "(" & FormatScore(confScore -leadConfScore, false) & ")"
					elseif confScore = leadConfScore then
						confScoreDiff = "&nbsp;"
					end if
					confScore = FormatScore(confScore, true)
				end if
			end if %>
				<td align="right"><% = pickScore %></td>
				<td align="right"><% = pickPct %></td>
<%			if USE_CONFIDENCE_POINTS then %>
				<td align="right"><% = confScore %></td>
				<td align="right"><span class="small"><% = confScoreDiff %></span></td>
<%			end if %>
			</tr>
<%		next
	else %>
			<tr>
				<td>&nbsp;</td>
				<td align="center" colspan="<% = gameCols %>" style="width: 40em;"><em>No entries found.</em></td>
				<td colspan="<% = otherCols - 1 %>">&nbsp;</td>
			</tr>
<%	end if %>
		</tbody>
	</table>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>
<%	'**************************************************************************
	'* Local class definitions.                                               *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' UserObj: Holds information for a single user.
	'--------------------------------------------------------------------------
	class UserObj

		public name
		public pickScore, confScore

		private sub Class_Initialize()
		end sub

		private sub Class_Terminate()
		end sub

		public sub setData(username)

			'Set the user properties.
			name      = username
			pickScore = PlayoffsPlayerPickScore(username)
			confScore = PlayoffsPlayerConfidenceScore(username)

		end sub

	end class %>

