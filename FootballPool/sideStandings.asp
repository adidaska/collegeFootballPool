<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = SidePoolTitle & " Pool Standings" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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
<%	if ENABLE_SURVIVOR_POOL then %>
	//-------------------------------------------------------------------------
	// Function for sorting columns in the Survivor table.
	//-------------------------------------------------------------------------
	function sortSurvivor(colName)
	{
		// Get the table or table section to sort.
		var tblEl = document.getElementById("survivor");

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

			case 'Correct':
				col = tblEl.rows[0].cells.length - 2;
				rev = true;
				sortCols[0] = col;
				hdrCols[0]  = col;
				break;

			case 'Status':
				col = tblEl.rows[0].cells.length - 1;
				rev = false;
				sortCols[0] = col;
				hdrCols[0]  = col;
				break;

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

				// Sort by secondary columns if those values are equal.
				if (cmp == 0)
				{
					// For 'Status' sort, use the 'Correct' column.
					if (colName == "Status")
					{
						cmp = compareValues(getTextValue(tblEl.rows[minIdx].cells[col - 1]),
						                    getTextValue(tblEl.rows[j].cells[col - 1]));

						// Sort this one in the opposite direction as 'Status'.
						if (!tblEl.reverseSort[colName])
							cmp = -cmp;
					}

					// Otherwise, use the 'Name' column.
					else
						cmp = compareValues(getTextValue(tblEl.rows[minIdx].cells[0]),
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
<%	end if
	if ENABLE_MARGIN_POOL then %>
	//-------------------------------------------------------------------------
	// Function for sorting columns in the Survivor table.
	//-------------------------------------------------------------------------
	function sortMargin(colName)
	{
		// Get the table or table section to sort.
		var tblEl = document.getElementById("margin");

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

			case 'Correct':
				col = tblEl.rows[0].cells.length - 2;
				rev = true;
				sortCols[0] = col;
				hdrCols[0]  = col;
				break;

			case 'Score':
				col = tblEl.rows[0].cells.length - 1;
				rev = true;
				sortCols[0] = col;
				hdrCols[0]  = col;
				break;

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

				// Sort by secondary columns if those values are equal.
				if (cmp == 0)
				{
					// For 'Score' sort, use the 'Correct' column.
					if (colName == "Score")
						cmp = compareValues(getTextValue(tblEl.rows[minIdx].cells[col - 1]),
						                    getTextValue(tblEl.rows[j].cells[col - 1]));

					// Sort this one in the same direction as 'Score'.
					if (tblEl.reverseSort[colName])
						cmp = -cmp;
				}

				// If the values are still equal, use the 'Name' column.
				if (cmp == 0)
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
<%	end if %>
	//]]></script>
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/side.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'Find the current week and the total number of weeks.
	dim curWeek, numWeeks
	curWeek = CurrentWeek()
	numWeeks = NumberOfWeeks()

	'Get a list of users active in the survivor pool.
	dim users
	users = SidePoolPlayersList()

	dim cols
	dim list, str, payout
	dim alt, i, week
	dim sql, rs
	dim pick, result
	dim pickLocked, hidePick
	dim currentUser
	currentUser = Session(SESSION_USERNAME_KEY)

	'Show survivor pool standings, if appropriate.
	if ENABLE_SURVIVOR_POOL then
		cols = numWeeks + 3 - SIDE_START_WEEK + 1 %>
	<h2>Survivor Pool Standings</h2>
	<table class="main" cellpadding="0" cellspacing="0">
		<thead>
			<tr class="header bottomEdge singleLine sortable" valign="bottom">
				<th align="left"><a href="#" onclick="this.blur(); return sortSurvivor('Name');" title="Sort by name.">Name</a></th>
<%		'Display the weeks, and check for any weeks were players were revived because all were eliminated.
		dim revives, lastRevives, anyRevives
		lastRevives = 0
		anyRevives = false
		for week = SIDE_START_WEEK to numWeeks
			str = ""
			sql = "SELECT MAX(Revived) AS Total FROM SurvivorStatus WHERE Week = " & week
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				revives = rs.Fields("Total").Value
				if revives <> lastRevives then
					str = "*"
					lastRevives = revives
					anyRevives = true
				end if
			end if %>
				<th class="sidePickHeader"><% = week & str %></th>
<%		next %>
				<th align="right"><a href="#" onclick="this.blur(); return sortSurvivor('Correct');" title="Sort by number of correct picks.">Correct</a></th>
				<th align="left"><a href="#" onclick="this.blur(); return sortSurvivor('Status');" title="Sort by current status.">Status</a></th>
			</tr>
		</thead>
<%		'If the pool has been concluded, show the winners and payout.
		str = ""
		list = SurvivorWinnersList()
		if IsArray(list) then
			payout = NumberOfSideEntries() * SIDE_BET_AMOUNT
			if ENABLE_MARGIN_POOL then
				payout = SURVIVOR_POT_SHARE * payout
			end if
			if UBound(list) > 0 then
				str = "Winners: " & Join(list, ", ") & " (" & FormatAmount(payout / (UBound(list) + 1)) & " each)"
			else
				str = "Winner: " & Join(list, ", ") & " (" & FormatAmount(payout) & ")"
			end if %>
		<tfoot>
			<tr class="header topEdge"><th align="left" colspan="<% = cols %>"><% = str %></th></tr>
		</tfoot>
<%		end if %>
		<tbody id="survivor">
<%		'Show each player's pick by week and overall status.
		dim games, missed, revived, isAlive, correct
		dim finalWeek, finalPoolWeek
		finalPoolWeek = SurvivorFinalWeek()
		if not IsNumeric(finalPoolWeek) then
			finalPoolWeek = numWeeks
		end if

		'If there are no players to show, let the user know.
		if not IsArray(users) then %>
			<tr><td align="center" colspan="<% = cols %>"><em>No entries found.</em></td></tr>
<%		else
			alt = false
			for i = 0 to UBound(users)
				if alt then  %>
			<tr align="center" class="alt singleLine">
<%				else %>
			<tr align="center" class="singleLine">
<%				end if
				alt = not alt %>
				<td align="left"><% = users(i) %></td>
<%				'Get the current status of this user.
				games   = 0
				missed  = 0
				revived = 0
				isAlive = true
				correct = 0
				set rs = GetSurvivorStatus(users(i), curWeek)
				if not (rs.BOF and rs.EOF) then
					games    = rs.Fields("CompletedGames").Value
					missed   = rs.Fields("Missed").Value
					revived  = rs.Fields("Revived").Value
					isAlive  = rs.Fields("IsAlive").Value
				end if
				correct = games - missed

				'Determine the last week this user started alive in the pool.
				finalWeek = numWeeks
				sql = "SELECT MAX(Week) AS FinalWeek FROM SurvivorStatus" _
				   & " WHERE Username = '" & users(i) & "'" _
				   & " AND WasAlive"
				set rs = DbConn.Execute(sql)
				if not (rs.BOF and rs.EOF) then
					finalWeek = rs.Fields("FinalWeek").Value
				end if

				'Get the user's pick and game result for all weeks.
				sql = "SELECT SidePicks.*, Schedule.Result FROM Schedule, SidePicks" _
				   & " WHERE Username = '" & SqlString(users(i)) & "'" _
				   & " AND SidePicks.Week >= " & SIDE_START_WEEK _
				   & " AND Schedule.Week = SidePicks.Week" _
				   & " AND (Pick = VisitorID or Pick = HomeID)" _
				   & " ORDER BY SidePicks.Week"
				set rs = DbConn.Execute(sql)

				'Show picks.
				for week = SIDE_START_WEEK to numWeeks

					'If the player is no longer active in the pool, show blanks.
					if (not isAlive and week > finalWeek) or week > finalPoolWeek then
						pick = "&nbsp;"
					else

						'Get the user's pick, margin score and result for the week.
						pick = ""
						result = ""
						if not rs.EOF then
							if rs.Fields("Week").Value = week then
								pick = rs.Fields("Pick").Value
								result = rs.Fields("Result").Value
								rs.MoveNext
							end if
						end if
						if IsNull(result) then
							result = ""
						end if

						'Handle a missing a pick.
						if pick = "" and week <= finalWeek then
							pick = "---"
						else

							'Determine if the pick is locked (note that we are
							'making some assumptions here based on the week).
							pickLocked = false
							if week < curWeek then
								pickLocked = true
							elseif  week = curWeek then
								pickLocked = SidePoolPickLocked(pick, week)
							end if

							'Determine if the pick should be hidden.
							hidePick = false
							if HIDE_OTHERS_PICKS and not pickLocked and not IsAdmin() and currentUser <> users(i) then
								hidePick = true
							end if

							'Format the pick display.
							if pick = "" then
								pick = "&nbsp;"
							elseif hidePick then
								pick = "XXX"
							else
								if pick = result or (result = TIE_STR and not SURVIVOR_STRIKE_ON_TIE) then
									pick = FormatCorrectPick(pick)
								elseif not isAlive and week = finalWeek then
									pick = "<span style=""text-decoration: line-through;"">" & pick & "</span>"
								end if
							end if
						end if
					end if %>
				<td><% = pick %></td>
<%				next %>
				<td align="right"><% = correct %></td>
<%				if isAlive then %>
				<td align="left">Alive</td>
<%				else %>
				<td align="left">Eliminated</td>
<%				end if %>
			</tr>
<%			next %>
		</tbody>
<%		end if %>
	</table>
<%		'If there were any revives, show the footnote.
		if anyRevives then %>
		<p class="small">*Week's results ignored as all players would have been eliminated.</p>
<%		end if
	end if

	'Show margin pool standings, if appropriate.
	if ENABLE_MARGIN_POOL then
		cols = numWeeks + 3 - SIDE_START_WEEK + 1 %>
	<h2>Margin Pool Standings</h2>
	<table class="main" cellpadding="0" cellspacing="0">
		<thead>
			<tr class="header bottomEdge singleLine sortable" valign="bottom">
				<th align="left"><a href="#" onclick="this.blur(); return sortMargin('Name');" title="Sort by name.">Name</a></th>
<%		for week = SIDE_START_WEEK to numWeeks %>
				<th class="sidePickHeader"><% = week %></th>
<%		next %>
				<th align="right"><a href="#" onclick="this.blur(); return sortMargin('Correct');" title="Sort by number of correct picks.">Correct</a></th>
				<th align="right"><a href="#" onclick="this.blur(); return sortMargin('Score');" title="Sort by score.">Score</a></th>
			</tr>
		</thead>
<%		'If the pool has been concluded, show the winners and payout.
		str = ""
		list = MarginWinnersList()
		if IsArray(list) then
			payout = NumberOfSideEntries * SIDE_BET_AMOUNT
			if ENABLE_SURVIVOR_POOL then
				payout = payout - (SURVIVOR_POT_SHARE * payout)
			end if
			if UBound(list) > 0 then
				str = "Winners: " & Join(list, ", ") & " (" & FormatAmount(payout / (UBound(list) + 1)) & " each)"
			else
				str = "Winner: " & Join(list, ", ") & " (" & FormatAmount(payout) & ")"
			end if %>
		<tfoot>
			<tr class="header topEdge"><th align="left" colspan="<% = cols %>"><% = str %></th></tr>
		</tfoot>
<%		end if %>
		<tbody id="margin">
<%		'Show each players pick and score by week and overall totals.
		dim score, scoreTotal

		'If there are no players to show, let the user know.
		if not IsArray(users) then %>
			<tr><td align="center" colspan="<% = cols %>"><em>No entries found.</em></td></tr>
<%		else
			alt = false
			for i = 0 to UBound(users)
				if alt then  %>
			<tr align="center" class="alt singleLine">
<%				else %>
			<tr align="center" class="singleLine">
<%				end if
				alt = not alt %>
				<td align="left"><% = users(i) %></td>
<%				'Initialize scoring data.
				scoreTotal = 0
				correct = 0

				'Get the user's pick, margin score and game result for all weeks.
				sql = "SELECT SidePicks.*, Schedule.Result FROM Schedule, SidePicks" _
				   & " WHERE Username = '" & SqlString(users(i)) & "'" _
				   & " AND SidePicks.Week >= " & SIDE_START_WEEK _
				   & " AND Schedule.Week = SidePicks.Week" _
				   & " AND (Pick = VisitorID or Pick = HomeID)" _
				   & " ORDER BY SidePicks.Week"
				set rs = DbConn.Execute(sql)

				'Show picks.
				for week = SIDE_START_WEEK to numWeeks

					'Get the user's pick, margin score and result for the week.
					pick = ""
					score = ""
					result = ""
					if not rs.EOF then
						if rs.Fields("Week").Value = week then
							pick = rs.Fields("Pick").Value
							score = rs.Fields("MarginScore").Value
							result = rs.Fields("Result").Value
							rs.MoveNext
						end if
					end if
					if not IsNumeric(score) then
						score = PlayerMarginScore(users(i), week)
					end if
					if IsNull(result) then
						result = ""
					end if

					'Determine if the pick is locked (note that we are making
					'some assumptions here based on the week).
					pickLocked = false
					if week < curWeek then
						pickLocked = true
					elseif  week = curWeek then
						pickLocked = SidePoolPickLocked(pick, week)
					end if

					'If we have a margin score, add it to the total.
					if IsNumeric(score) then
						scoreTotal = scoreTotal + score
						if score > 0 then
							correct = correct + 1
						end if
					else
						score = "--"
					end if

					'Determine if the pick should be hidden.
					hidePick = false
					if HIDE_OTHERS_PICKS and not pickLocked and not IsAdmin() and currentUser <> users(i) then
						hidePick = true
					end if

					'Format the pick display.
					if pick = "" then
						pick = "---"
						if score = "" then
							score= "--"
						end if
					elseif hidePick then
						pick = "XXX"
						score = "XX"
					else
						if result <> "" and pick = result then
							pick = FormatCorrectPick(pick)
							if score > 0 then
								score = "+" & score
							end if
						end if
					end if %>
				<td class="margPick"><% = pick %><br /><% = score %></td>
<%				next %>
				<td align="right"><% = correct %></td>
				<td align="right"><% = scoreTotal %></td>
			</tr>
<%			next %>
		</tbody>
<%		end if %>
	</table>
<%	end if %>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>
