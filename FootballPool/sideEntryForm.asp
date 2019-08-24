<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = SidePoolTitle & " Pool Entry Form" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'If the side pool is disabled, redirect to the home page.
	if not ENABLE_SURVIVOR_POOL and not ENABLE_MARGIN_POOL then
		Response.Redirect("./")
	end if

	'Get the week to display.
	dim week
	week = GetRequestedWeek()

	'If the user is the Administrator, check for a user name in the request.
	'Otherwise, show data for the current user.
	dim username
	username = Session(SESSION_USERNAME_KEY)
	if IsAdmin() then
		username = Trim(Request("username"))
	end if

	'Initialize lock flag.
	dim allLocked
	allLocked = AllGamesLocked(week)

	'If the survivor pool is enabled, determine if it has already been
	'concluded.
	dim finalWeek, isSurvivorOver
	isSurvivorOver= false
	if ENABLE_SURVIVOR_POOL then
		finalWeek = SurvivorFinalWeek()
		if IsNumeric(finalWeek) and week > finalWeek then
			isSurvivorOver = true
		end if
	end if

	'Determine if the user is eligible to make an entry.
	'Note: All users are eligible the first week of the pool. After that, only
	'users who entered that first week are eligible.
	dim isEntryAllowed
	isEntryAllowed = false
	if username <> "" then
		if InSidePool(username) or (week = SIDE_START_WEEK and (IsAdmin() or not allLocked)) then
			isEntryAllowed = true
		end if
	end if

	'For the Administrator, build a user selection list.
	dim users, i
	if IsAdmin() then %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<table class="main" cellpadding="0" cellspacing="0">
		<tr class="header bottomEdge">
			<th align="left">Administrator Access</th>
		</tr>
		<tr>
			<td class="adminSection freeForm">
				<p>You may view or edit any player's entry form by selecting a username below.<br />
				Use the links at the bottom of the page to switch to a specific week.</p>
				<table>
					<tr>
						<td><strong>Select user:</strong></td>
						<td>
							<input type="hidden" name="week" value="<% = week %>" />
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
						<td><input type="submit" name="submit" value="Select" class="button" title="View/edit the selected user's entry." /></td>
					</tr>
				</table>
			</td>
		</tr>
		</table>
	</form>
<%		if username <> "" then %>
	<h2>Entry Form for <% = username %></h2>
<%		end if
	end if

	'Determine if the user has an entry for the given week (so we can display
	'the Delete button) and get its lock status.
	dim hasEntry
	dim sidePick, sidePickLocked
	hasEntry = false
	sidePickLocked = false
	sidePick = GetSidePoolPick(username, week)
	if sidePick <> "" then
		hasEntry = true
		sidePickLocked = SidePoolPickLocked(sidePick, week)
	end if

	'If there is form data, process it.
	dim sql, rs
	dim needDeleteConfirmation, entryDeleted
	needDeleteConfirmation = false
	entryDeleted = false

	'Process a delete request.
	if Request.Form("submit") = "Delete" then
		if not IsAdmin() and sidePickLocked then
			call ErrorMessage("Error: The game for your current pick has been locked, entry<br />cannot be deleted.")
		elseif LCase(Request.Form("confirmDelete")) <> "true" then
			needDeleteConfirmation = true
			call ErrorMessage("To confirm deletion, check the box below and press <code>Delete</code> again.<br />Pressing any other button will cancel the deletion.")
		elseif NumberOfCompletedGames(week) = NumberOfGames(week) then
			call ErrorMessage("Error: Cannot delete entries for completed weeks.")
		else

			'Delete the player's entry.
			sql = "DELETE FROM SidePicks" _
			   & " WHERE Username = '" & SqlString(username) & "'" _
			   & " AND Week = " & week
			call DbConn.Execute(sql)
			entryDeleted = true
			hasEntry = false

			'Clear the player's pick.
			sidePick = ""

			'If the survivor pool is enabled, clear any saved status.
			if ENABLE_SURVIVOR_POOL then
				call ClearSurvivorStatus(week)
			end if

			'If the margin pool is enabled, clear any saved results.
			if ENABLE_MARGIN_POOL then
				call ClearMarginResultsCache(week)
			end if

			call InfoMessage("Entry has been deleted.")
		end if

	'Process an update request.
	elseif Request.Form("submit") = "Update" then

		'Get the new pick
		dim newSidePick, newSidePickLocked
		dim usedSidePicks
		newSidePick = Request.Form("sidePick")
		newSidePickLocked = SidePoolPickLocked(newSidePick, week)

		'Prevent entries by the Admin.
		if username = ADMIN_USERNAME then
			call ErrorMessage("Error: User '" & ADMIN_USERNAME & "' may not make picks.")

		'If the user is disabled, prevent the entry.
		elseif IsDisabled() then
			call ErrorMessage("Error: Your account has been disabled, change not accepted.<br />Please contact the Administrator.")

		'If only the survivor pool is enabled and it is over, prevent the
		'entry.
		elseif not ENABLE_MARGIN_POOL and isSurvivorOver then
			call ErrorMessage("Error: The survivor pool has been concluded, entries no longer accepted.")

		'If the player is not eligible to make an entry, prevent it.
		elseif not isEntryAllowed then
			call ErrorMessage("Error: You are not eligible to make entries in this pool.")

		'If all games have been locked, prevent any updates (except by the
		'Admin).
		elseif not IsAdmin() and allLocked then
			call ErrorMessage("Error: All games for this week have been locked, change not<br />accepted.")

		'If the current pick has been locked, prevent any updates (except by
		'the Admin).
		elseif not IsAdmin() and sidePickLocked then
			call ErrorMessage("Error: The " & GetTeamName(sidePick) & " game has been locked, change<br />not accepted.")
		elseif not IsAdmin() and newSidePickLocked then
			call ErrorMessage("Error: The " & GetTeamName(newSidePick) & " game has been locked, change<br />not accepted.")

		'Otherwise, process the request.
		else

			'Make sure a pick was made.
			if newSidePick = "" then
				FormFieldErrors.Add "sidePick", "You must select a team."
			else

				'Validate the team has not been used.
				if newSidePick <> sidePick then
					usedSidePicks = UsedTeamsList(username)
					if IsArray(usedSidePicks) then
						for i = 0 to UBound(usedSidePicks)
							if newSidePick = usedSidePicks(i) then
								FormFieldErrors.Add "sidePick", GetTeamName(newSidePick) & " has already been used."
								exit for
							end if
						next
					end if
				end if
			end if

			'If there were any errors, display the error message. Otherwise, do
			'the updates.
			if FormFieldErrors.Count > 0 then
				call FormFieldErrorsMessage("Error: Invalid fields. Please correct and resubmit.")
			else
				sql = "DELETE FROM SidePicks" _
				   & " WHERE Week = " & week _
				   & " AND Username = '" & SqlString(username) & "'"
				call DbConn.Execute(sql)
				sql = "INSERT INTO SidePicks" _
				   & " (Week, Username, Pick, MarginScore)" _
				   & " VALUES(" & week & ", " _
				   & "'" & SqlString(username) & "', " _
				   & "'" & newSidePick & "', " _
				   & "NULL)"
				call DbConn.Execute(sql)

				'If the survivor pool is active, clear any cached status.
				if ENABLE_SURVIVOR_POOL then
					call ClearSurvivorStatus(week)
				end if

				'Updates complete, redirect to the side pool standings page.
				Response.Redirect("sideStandings.asp")
			end if
		end if
	end if


	'Build the entry form for that week.
	if username <> "" then
		dim showEntry
		dim wasAlive, isAlive
		dim msgStr
		dim resultStr
		dim score, scoreStr %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<div>
<%		if IsAdmin() then %>
			<input type="hidden" name="username" value="<% = username %>" />
<%		end if %>
			<input type="hidden" name="week" value="<% = week %>" />
		</div>
		<table><tr><td style="padding: 0px;">
		<table class="main" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge">
			  <th align="left" colspan="8">Week <% = week %></th>
			</tr>
<%		'Show all games for that week.
		dim visitor, home
		dim vscore, hscore
		dim ot, result
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

				if not alt then %>
			<tr align="right">
<%				else %>
			<tr align="right" class="alt">
<%				end if
				alt = not alt %>
				<td align="left"><% = WeekdayName(Weekday(rs.Fields("Date").Value), true) %></td>
				<td><% = FormatDate(rs.Fields("Date").Value) %></td>
				<td><% = FormatTime(rs.Fields("Time").Value) %></td>
				<td><% = visitor %></td>
				<td><% = vscore %></td>
				<td>at <% = home %></td>
				<td><% = hscore %></td>
				<td><span class="small"><% = ot %></span></td>
			</tr>
<%				rs.MoveNext
			loop
		end if %>
			<tr class="header topEdge bottomEdge">
				<th align="left" colspan="8"><% = SidePoolTitle %> Pool Pick</th>
			</tr>
			<tr>
				<td colspan="8">
<%		'If the pool has not started, display a message.
		if week < SIDE_START_WEEK then %>
					<p>This pool begins in Week <% = SIDE_START_WEEK %>.</p>
<%		'If the player cannot participate in the side pool, display a message.
		elseif not isEntryAllowed then %>
					<p>You are not participating in this pool.</p>
<%		else
			'Get the player's current pick.
			if not entryDeleted then
				sidePick = GetFieldValue("sidePick", sidePick)
			end if

			'Determine the player's status.
			wasAlive = false
			isAlive = false
			showEntry = false
			if ENABLE_SURVIVOR_POOL then
				set rs = GetSurvivorStatus(username, week)
				if not (rs.BOF and rs.EOF) then
					wasAlive = rs.Fields("WasAlive").Value
					isAlive = rs.Fields("IsAlive").Value
				elseif week = SIDE_START_WEEK and (IsAdmin or not allLocked) then
					wasAlive = true
					isAlive = true
				end if
			end if

			'Determine if we should show the entry.
			if ENABLE_MARGIN_POOL or (ENABLE_SURVIVOR_POOL and not isSurvivorOver and wasAlive) then
				showEntry = true
			end if

			'If the margin pool is enabled, check for a margin score and format
			'it if it is available.
			scoreStr = ""
			if ENABLE_MARGIN_POOL then
				score = PlayerMarginScore(username, week)
				if IsNumeric(score) then
					scoreStr = score
					if score > 0 then
						scoreStr = "+" & scoreStr
					end if
					scoreStr = "(" & scoreStr & ")"
					if score > 0 then
						scoreStr = "<strong>" & scoreStr & "</strong>"
					end if
					scoreStr = "&nbsp;" & scoreStr
				end if
			end if

			'Display a message to the user, if appropriate.
			msgStr = ""
			if ENABLE_SURVIVOR_POOL then
				if not isAlive then
					msgStr = "You have been eliminated from the survivor pool."
				elseif isSurvivorOver then
					msgStr = "The survivor pool has been concluded."
				end if
				if msgStr <> "" and ENABLE_MARGIN_POOL then
					msgStr = msgStr & "<br />Your picks still count toward the margin pool, however."
				end if
			end if
			if msgStr <> "" then
				msgStr = "<p><em>" & msgStr & "</em></p>"
			end if %>
					<% = msgStr %>
<%			'Show the entry, if appropriate.
			if showEntry then %>
					<table style="width: 100%"><tr><td align="right">
<%				if not IsAdmin() and sidePickLocked then
					resultStr = FormatSidePoolPick(sidePick, week)
					if ENABLE_MARGIN_POOL then
						resultStr = resultStr & scoreStr
					end if %>
					<% = resultStr %>
<%				else %>
						<strong>Select team:</strong>
						<select name="sidePick" class="<% = FieldStyleClass("", "sidePick") %>">
<%					call DisplaySidePoolPickList(7, username, week, sidePick) %>
						</select>
<%				end if %>
					</td></tr></table>
<%			end if
		end if %>
				</td>
			</tr>
		</table>
<%		'List open dates.
		call DisplayOpenDates(2, week)

		'Show form buttons, if changes are allowed.
		if IsAdmin() or not sidePickLocked then %>
		<p></p>
		<table cellpadding="0" cellspacing="0" style="width: 100%;">
			<tr valign="middle">
				<td style="padding: 0px;"><input type="submit" name="submit" value="Update" class="button" title="Apply changes." />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel the update." /></td>
<%			'If the user has an entry and no games are locked (or the user is
			'the Admin), add a delete button.
			if hasEntry and (IsAdmin() or not sidePickLocked) then

				'If delete was requested, add a confirmation checkbox.
				if needDeleteConfirmation then %>
				<td align="right" style="padding: 0px;"><input type="checkbox" id="confirmDelete" name="confirmDelete" value="true" /> <label for="confirmDelete">Confirm Deletion</label>&nbsp;</td>
<%				end if %>
				<td align="right" style="padding: 0px;"><input type="submit" name="submit" value="Delete" class="button" title="Delete this entry." /></td>
<%			end if %>
			</tr>
		</table>
<%		end if %>
		</td></tr></table>
	</form>
<%	end if

	'List links to view other weeks.
	dim params
	params = ""
	if IsAdmin() then
		params = "username=" & Server.HtmlEncode(username)
	end if
	call DisplayWeekNavigation(1, params) %>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>
