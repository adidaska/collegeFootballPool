<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Account History" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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
					<p>You may view any player's account history by selecting a username below.</p>
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

	'Build a list of transactions for the given user.
	dim transactions, sql, rs
	dim totalFees
	transactions = Array()
	if username <> "" then

		'Add a transaction for each credit record in the database.
		sql = "SELECT * FROM Credits" _
		   & " WHERE Username = '" & SqlString(username) & "'"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			do while not rs.EOF
				call AddTransaction(rs.Fields("Timestamp").Value, rs.Fields("Description").Value, rs.Fields("Amount").Value)
			rs.MoveNext
			loop
		end if

		'Get the current date and time.
		dim dateNow
		dateNow = CurrentDateTime()

		'Add a transaction for the survivor/margin pool if that pool has
		'started and the player has an entry.
		if ENABLE_SURVIVOR_POOL or ENABLE_MARGIN_POOL then
			if InSidePool(username) and dateNow >= WeekStartDateTime(SIDE_START_WEEK) then
				call AddTransaction(WeekStartDateTime(SIDE_START_WEEK), SidePoolTitle & "Pool Entry Fee", -SIDE_BET_AMOUNT)
				totalFees = totalFees + SIDE_BET_AMOUNT
			end if
		end if

		'Add a transaction for each week played so far. Note that the current
		'week is included only after the first game has started.
		dim week
		week = CurrentWeek()
		if dateNow <= WeekStartDateTime(week) then
			week = week - 1
		end if
		sql = "SELECT Week FROM Tiebreaker" _
		   & " WHERE Username = '" & SqlString(username) & "'" _
		   & " AND Week <= " & week
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			do while not rs.EOF
				week = rs.Fields("Week").Value
				call AddTransaction(WeekStartDateTime(week), "Week " & week & " Entry Fee", -BET_AMOUNT)
				totalFees = totalFees + BET_AMOUNT
				rs.MoveNext
			loop
		end if
	end if

	'Add a transaction for the playoffs pool, if the playoffs have started and
	'the player has an entry.
	dim playoffsDate
	if ENABLE_PLAYOFFS_POOL then
		playoffsDate = PlayoffsStartDateTime()
		if dateNow >= playoffsDate and InPlayoffsPool(username) then
			call AddTransaction(playoffsDate, "Playoffs Entry Fee", -PLAYOFFS_BET_AMOUNT)
			totalFees = totalFees + PLAYOFFS_BET_AMOUNT
		end if
	end if

	'Sort the transactions
	dim j, temp
	for i = 0 to UBound(transactions) - 1
		for j = i + 1 to UBound(transactions)
			if transactions(i).timestamp > transactions(j).timestamp then
				set temp = transactions(i)
				set transactions(i) = transactions(j)
				set transactions(j) = temp
			end if
		next
	next

	'Display the user's transactions.
	if username <> "" then
		if IsAdmin() then %>
		<h2>Account History for <% = username %></h2>
<%		else %>
		<h2>Account History</h2>
<%		end if %>
	<table class="main" cellpadding="0" cellspacing="0">
		<tr class="header bottomEdge">
			<th align="left" colspan="4">Transactions</th>
		</tr>
		<tr class="subHeader bottomEdge">
			<th align="left">Date</th>
			<th align="left">Time</th>
			<th align="left" style="width: 24em;">Description</th>
			<th align="right">Amount</th>
		</tr>
<%		dim total, alt
		total = 0
		if UBound(transactions) >= 0 then
			alt = false
			for i = 0 to UBound(transactions)
				if alt then %>
		<tr align="right" class="alt" valign="top">
<%				else %>
		<tr align="right" valign="top">
<%				end if
				alt = not alt %>
			<td><% = FormatFullDate(transactions(i).timestamp) %></td>
			<td><% = FormatFullTime(transactions(i).timestamp) %></td>
			<td align="left"><% = transactions(i).description %></td>
			<td><% = FormatAmount(transactions(i).amount) %></td>
		</tr>
<%			total = total + transactions(i).amount
			next
    	 else %>
		<tr>
			<td align="center" colspan="4"><em>No transactions found.</em></td>
		</tr>
<%  	 end if %>
		<tr class="header topEdge">
			<th align="left" colspan="3">Balance:</th>
				<th align="right"><% = FormatAmount(total) %></th>
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
	' Adds a transaction to the global list.
	'--------------------------------------------------------------------------
	sub AddTransaction(timestamp, description, amount)

		dim n

		n = UBound(transactions) + 1
		redim preserve transactions(n)
		set transactions(n) = new TransactionObj
		call transactions(n).setData(timestamp, description, amount)

	end sub

	'**************************************************************************
	'* Local class definitions.                                               *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' TransactionObj: Holds information for a single account transaction.
	'--------------------------------------------------------------------------
	class TransactionObj

		public timestamp
		public description
		public amount

		private sub Class_Initialize()
		end sub

		private sub Class_Terminate()
		end sub

		public sub setData(t, d, a)

			'Set properties.
			timestamp = t
			description = d
			amount = a

		end sub

	end class %>