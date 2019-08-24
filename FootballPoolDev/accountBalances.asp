<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Account Balances" : AdminOnly = true %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/style.css" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
	<script type="text/javascript" src="scripts/tableSort.js"></script>
	<script type="text/javascript">//<![CDATA[
	//-------------------------------------------------------------------------
	// Function for sorting columns in the Account Balances table.
	//-------------------------------------------------------------------------
	function sortAccounts(colName)
	{
		// Get the table section to sort.
		var tblEl = document.getElementById("accountBalances");

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
			case 'Fees':
				col = 1
				sortCols[0] = col;
				hdrCols[0]  = col;
				rev = true;
				break;
			case 'Credits':
				col = 2
				sortCols[0] = col;
				hdrCols[0]  = col;
				rev = true;
				break;
			case 'Balance':
				col = 3
				sortCols[0] = col;
				hdrCols[0]  = col;
				rev = true;
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

				// If the values are equal, use the 'Name' column.
				if (cmp == 0)
					cmp = compareValues(
						getTextValue(tblEl.rows[minIdx].cells[0]),
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
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/playoffs.asp" -->
<!-- #include file="includes/side.asp" -->
<!-- #include file="includes/weekly.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'Get a list of users.
	dim users
	users = UsersList(true)

	'Determine if games for the current week have started.
	dim week, dateNow
	week = CurrentWeek()
	dateNow = CurrentDateTime()
	if dateNow <= WeekStartDateTime(week) then
		week = week - 1
	end if

	dim sql, rs
	dim i
	dim fees, credits
	dim totalReceived
	dim alt
	totalReceived = 0 %>
	<table class="main fixed" cellpadding="0" cellspacing="0">
		<thead>
			<tr align="right" class="header bottomEdge sortable">
				<th align="left"><a href="#" onclick="this.blur(); return sortAccounts('Name');" title="Sort by name.">Name</a></th>
				<th><a href="#" onclick="this.blur(); return sortAccounts('Fees');" title="Sort by total fees.">Total Fees</a></th>
				<th><a href="#" onclick="this.blur(); return sortAccounts('Credits');" title="Sort by credits.">Total Credits</a></th>
				<th><a href="#" onclick="this.blur(); return sortAccounts('Balance');" title="Sort by balance.">Balance</a></th>
			</tr>
		</thead>
		<tbody id="accountBalances">
<%	if IsArray(users) then
		alt = false
		for i = 0 to UBound(users)
			fees = 0

			'Determine how many weekly entries the player has and add up those
			'fees. Note that the current week is included only after the first
			'game has started.
			sql = "SELECT COUNT(*) AS Total FROM Tiebreaker" _
			   & " WHERE Username = '" & SqlString(users(i)) & "'" _
			   & " AND Week <= " & week
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				fees = rs.Fields("Total").Value * BET_AMOUNT
			end if

			'Add playoffs entry fee, if appropriate.
			if InPlayoffsPool(users(i)) and dateNow >= PlayoffsStartDateTime() then
				fees = fees + PLAYOFFS_BET_AMOUNT
			end if

			'Add side pool entry fee, if appropriate.
			if InSidePool(users(i)) and dateNow >= WeekStartDateTime(SIDE_START_WEEK) then
				fees = fees + SIDE_BET_AMOUNT
			end if

			'Get credits.
			credits = TotalCredits(users(i))

			'Add player's credits to total received.
			totalReceived = totalReceived + credits

			if alt then %>
			<tr align="right" class="alt singleLine">
<%			else %>
			<tr align="right" class="singleLine">
<%			end if
			alt = not alt %>
				<td align="left"><a href="accountHistory.asp?username=<% = Server.UrlEncode(users(i)) %>" title="View player's account history."><% = users(i) %></a></td>
				<td><% = FormatAmount(fees) %></td>
				<td><% = FormatAmount(credits) %></td>
				<td><% = FormatAmount(credits - fees) %></td>
			</tr>
<%		next
	else %>
			<tr>
				<td align="center" colspan="4"><em>No users found.</em></td>
			</tr>
<%	end if

		'Display fund summary.
		dim totalPaid
		dim n, payout
		totalPaid = 0

		'Get total payouts. The current week is excluded if it has not been
		'completed yet.
		n = CurrentWeek()
		if TBPointTotal(n) = "" then
			n = n - 1
		end if
		for i = 1 to n
			totalPaid = totalPaid + NumberOfEntries(i) * BET_AMOUNT
		next

		'Add the playoffs pool payout, if appropriate.
		if ENABLE_PLAYOFFS_POOL then
			if IsArray(PlayoffsWinnersList()) then
				totalPaid = totalPaid + NumberOfPlayoffsEntries() * PLAYOFFS_BET_AMOUNT
			end if
		end if

		'Add the survivor pool payout, if appropriate.
		if ENABLE_SURVIVOR_POOL then
			if IsArray(SurvivorWinnersList()) then
				payout = NumberOfSideEntries() * SIDE_BET_AMOUNT
				if ENABLE_MARGIN_POOL then
					payout = SURVIVOR_POT_SHARE * payout
				end if
				totalPaid = totalPaid + payout
			end if
		end if

		'Add the margin pool payout, if appropriate.
		if ENABLE_MARGIN_POOL then
			if IsArray(MarginWinnersList()) then
				payout = NumberOfSideEntries() * SIDE_BET_AMOUNT
				if ENABLE_SURVIVOR_POOL then
					payout = payout - (SURVIVOR_POT_SHARE * payout)
				end if
				totalPaid = totalPaid + payout
			end if
		end if %>
		</tbody>
		<tbody>
			<tr class="header topEdge bottomEdge">
				<th align="left" colspan="4">Fund Summary</th>
			</tr>
			<tr>
				<td colspan="2">Total payments received:</td>
				<td align="right" colspan="2"><% = FormatAmount(totalReceived) %></td>
			</tr>
			<tr class="alt">
				<td colspan="2">Total winnings paid:</td>
				<td align="right" colspan="2"><% = FormatAmount(totalPaid) %></td>
			</tr>
			<tr class="subHeader topEdge">
				<th align="left" colspan="2"><strong>Balance:</strong></th>
				<th align="right" colspan="2"><% = FormatAmount(totalReceived - totalPaid) %></th>
			</tr>
		</tbody>
	</table>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>