<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Pool Winnings" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/side.asp" -->
<!-- #include file="includes/playoffs.asp" -->
<!-- #include file="includes/weekly.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'If the user is the Administrator, check for a user name in the request.
	'Otherwise, show data for the current user.
	dim username
	username = Session(SESSION_USERNAME_KEY)
	if IsAdmin() then
		username = Trim(Request("username"))
	end if

	'For the Administrator, build a user selection list.
	dim users, i
	if IsAdmin() then %>
		<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
			<table class="main fixed" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge">
				<th align="left">Administrator Access</th>
			</tr>
			<tr>
				<td class="adminSection freeForm">
					<p>You may view any player's pool winnings by selecting a username below.</p>
					<table>
						<tr>
							<td><strong>Select user:</strong></td>
							<td>
								<select name="username">
									<option value=""></option>
		<%		users = UsersList(true)
				if IsArray(users) then
					for i = 0 to UBound(users) %>
									<option value="<% = users(i) %>" <% if users(i) = username then Response.Write(" selected=""selected""") end if %>><% = users(i) %></option>
		<%			next
				end if %>
								</select>
							</td>
							<td><input type="submit" name="submit" value="Select" class="button" title="View/edit the selected user's profile." /></td>
						</tr>
					</table>
				</td>
			</tr>
			</table>
		</form>
<%	end if

	'Get the total pool entry fees for the user.
	dim sql, rs
	dim totalFees
	if username <> "" then

		'Get the current date and time.
		dim dateNow
		dateNow = CurrentDateTime()

		'Add the fee for the survivor/margin pool if that pool has started and
		'the player has an entry.
		if ENABLE_SURVIVOR_POOL or ENABLE_MARGIN_POOL then
			if InSidePool(username) and dateNow >= WeekStartDateTime(SIDE_START_WEEK) then
				totalFees = totalFees + SIDE_BET_AMOUNT
			end if
		end if

		'Add fees for each week played so far. Note that the current week is
		'included only after the first game has started.
		dim week
		week = CurrentWeek()
		if dateNow <= WeekStartDateTime(week) then
			week = week - 1
		end if
		sql = "SELECT COUNT(*) AS Total FROM Tiebreaker" _
		   & " WHERE Username = '" & SqlString(username) & "'" _
		   & " AND Week <= " & week
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			totalFees = totalFees + rs.Fields("Total").Value * BET_AMOUNT
		end if

		'Add the fee for the playoffs pool, if the playoffs have started and the
		'player has an entry.
		dim playoffsDate
		if ENABLE_PLAYOFFS_POOL then
			playoffsDate = PlayoffsStartDateTime()
			if dateNow >= playoffsDate and InPlayoffsPool(username) then
				totalFees = totalFees + PLAYOFFS_BET_AMOUNT
			end if
		end if

		'Build a list of pools won by the user.

		dim pools
		dim winners, n, pot
		dim numWeeks
		pools = Array()
		numWeeks = NumberOfWeeks
		for i = 1 to numWeeks
			winners = WinnersList(i)
			if InList(username, winners) then
				n = UBound(winners) + 1
				pot = NumberOfEntries(i) * BET_AMOUNT
				call AddPool("Week " & i & " Pool", pot / n)
			end if
		next
		if ENABLE_PLAYOFFS_POOL then
			winners = PlayoffsWinnersList()
			if InList(username, winners) then
				n = UBound(winners) + 1
				pot = NumberOfPlayoffsEntries() * PLAYOFFS_BET_AMOUNT
				call AddPool("Playoffs Pool", pot / n)
			end if
		end if
		if ENABLE_SURVIVOR_POOL then
			winners = SurvivorWinnersList()
			if InList(username, winners) then
				n = UBound(winners) + 1
				pot = NumberOfSideEntries() * SIDE_BET_AMOUNT
				if ENABLE_MARGIN_POOL then
					pot = SURVIVOR_POT_SHARE * pot
				end if
				call AddPool("Survivor Pool", pot / n)
			end if
		end if
		if ENABLE_MARGIN_POOL then
			winners = MarginWinnersList()
			if InList(username, winners) then
				n = UBound(winners) + 1
				pot = NumberOfSideEntries() * SIDE_BET_AMOUNT
				if ENABLE_SURVIVOR_POOL then
					pot = pot - (SURVIVOR_POT_SHARE * pot)
				end if
				call AddPool("Margin Pool", pot / n)
			end if
		end if

		'Build the display.
		if IsAdmin() then %>
		<h2>Pool Winnings for <% = username %></h2>
<%		end if %>
	<table class="main fixed" cellpadding="0" cellspacing="0">
		<tr class="header bottomEdge">
			<th align="left">Pool</th>
			<th align="right">Amount</th>
		</tr>
<%		dim totalWinnings
		dim alt
		if UBound(pools) >= 0 then
			totalWinnings = 0
			alt = false
			for i = 0 to UBound(pools)
				if alt then %>
		<tr class="alt" valign="top">
<%				else %>
		<tr valign="top">
<%				end if
				alt = not alt %>
			<td><% = pools(i).name %></td>
			<td align="right"><% = FormatAmount(pools(i).amount) %></td>
		</tr>
<%				totalWinnings = totalWinnings + pools(i).amount
			next
    	 else %>
		<tr>
			<td align="center" colspan="2"><em>No pools won.</em></td>
		</tr>
<%  	 end if %>
		<tr class="header topEdge bottomEdge">
			<th align="left" colspan="2">Totals</th>
		</tr>
		<tr>
			<td>Winnings:</td>
			<td align="right"><% = FormatAmount(totalWinnings) %></td>
		</tr>
		<tr class="alt">
			<td>Fees:</td>
			<td align="right"><% = FormatAmount(totalFees) %></td>
		</tr>
		<tr class="subHeader topEdge">
			<th align="left">Net:</th>
			<th align="right"><% = FormatAmount(totalWinnings - totalFees) %></th>
		</tr>
	</table>
<%	end if %>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>
<%	'**************************************************************************
	'* Local functions and subroutines.                                       *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Returns true if the given string occurs in the given array of strings.
	'--------------------------------------------------------------------------
	function InList(str, list)

		dim i

		InList = false
		if IsArray(list) then
			for i = 0 to UBound(list)
				if list(i) = str then
					InList = true
					exit function
				end if
			next
		end if

	end function

	'--------------------------------------------------------------------------
	' Adds a pool to the global list.
	'--------------------------------------------------------------------------
	sub AddPool(name, amount)

		dim n

		n = UBound(pools) + 1
		redim preserve pools(n)
		set pools(n) = new PoolObj
		call pools(n).setData(name, amount)

	end sub

	'**************************************************************************
	'* Local class definitions.                                               *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' PoolObj: Holds information about a single pool the user has won.
	'--------------------------------------------------------------------------
	class PoolObj

		public name
		public amount

		private sub Class_Initialize()
		end sub

		private sub Class_Terminate()
		end sub

		public sub setData(n, a)

			'Set properties.
			name = n
			amount = a

		end sub

	end class %>