<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Playoffs Entry Form" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/common.css" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
<%	if USE_CONFIDENCE_POINTS then %>
	<script type="text/javascript" src="scripts/confidencePoints.js"></script>
<%	end if %>
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/playoffs.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'If the playoffs pool is disabled, redirect to the home page.
	if not ENABLE_PLAYOFFS_POOL then
		Response.Redirect("./")
	end if

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
			<table class="main" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge">
				<th align="left">Administrator Access</th>
			</tr>
			<tr>
				<td class="adminSection freeForm">
					<p>You may view or edit any player's entry form by selecting a username below.</p>
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

	'Determine if the user has an entry for the playoffs pool (so we can display
	'the Delete button).
	dim hasEntry
	hasEntry = false
	if username <> "" and PlayoffsPlayerConfidenceScore(username) <> "" then
		hasEntry = true
	end if

	'Initialize global lock out variables.
	dim allLocked, anyLocked
	allLocked = true
	anyLocked = false

	'Create an array of game objects
	dim games, rs, n
	n = NumberOfPlayoffGames()
	redim games(n - 1)
	sql = "SELECT * FROM PlayoffsSchedule ORDER BY Round, Date, Time"
	set rs = DbConn.Execute(sql)
	if not (rs.BOF and rs.EOF) then
		i = 0
		do while not rs.EOF
			set games(i) = new GameObj
			games(i).setData(rs)
			i = i + 1
			rs.MoveNext
		loop
	end if

	'If there is form data, process it.
	dim needDeleteConfirmation, entryDeleted
	dim pick, conf
	dim pickMissingError, confMissingError
	dim confUsed
	needDeleteConfirmation = false
	entryDeleted = false

	'Process a delete request.
	if Request.Form("submit") = "Delete" then

		if not IsAdmin() and anyLocked then
			call ErrorMessage("Error: One or more playoff games have been locked, entry<br />cannot be deleted.")
		elseif LCase(Request.Form("confirmDelete")) <> "true" then
			needDeleteConfirmation = true
			call ErrorMessage("To confirm deletion, check the box below and press <code>Delete</code> again.<br />Pressing any other button will cancel the deletion.")
		elseif NumberOfCompletedPlayoffGames() = NumberOfPlayoffGames() then
			call ErrorMessage("Error: Playoffs pool has been concluded, cannot delete entries.")
		else

			'Delete the player's entry.
			for i = 0 to UBound(games)
				sql = "DELETE FROM PlayoffsPicks" _
				   & " WHERE Username = '" & SqlString(username) & "'" _
				   & " AND GameID = " & games(i).id
				call DbConn.Execute(sql)
			next
			entryDeleted = true
			hasEntry = false

			'Reload the games.
			redim games(n - 1)
			sql = "SELECT * FROM PlayoffsSchedule ORDER BY Round, Date, Time"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				i = 0
				do while not rs.EOF
					set games(i) = new GameObj
					games(i).setData(rs)
					i = i + 1
					rs.MoveNext
				loop
			end if

			call InfoMessage("Entry has been deleted.")
		end if

	'Process an update request.
	elseif Request.Form("submit") = "Update" then

		'Determine if at least one pick was made.
		dim anyPicks
		anyPicks = false
		for i = 0 to UBound(games)
			if Trim(Request.Form("pick-" & (i + 1))) <> "" or _
			   (USE_CONFIDENCE_POINTS and Trim(Request.Form("conf-" & (i + 1))) <> "") then
				anyPicks = true
				exit for
			end if
		next

		'Prevent entries by the Admin.
		if username = ADMIN_USERNAME then
			call ErrorMessage("Error: User '" & ADMIN_USERNAME & "' may not make picks.")

		'If the user is disabled, prevent the entry.
		elseif IsDisabled() then
			call ErrorMessage("Error: Your account has been disabled, changes not accepted.<br />Please contact the Administrator.")

		'If all games have been locked, prevent any updates (except by the Admin).
		elseif not IsAdmin() and allLocked then
			call ErrorMessage("Error: All playoff games have been locked, changes not<br />accepted.")

		'If no picks were entered, prevent the entry.
		elseif not anyPicks then
			call ErrorMessage("Error: No picks entered.")

		'Otherwise, process the request.
		else

			'Validate the pick and confidence point fields.
			pickMissingError = false
			confMissingError = false
			if USE_CONFIDENCE_POINTS then
				for i = 0 to UBound(games)
					pick = Trim(Request.Form("pick-" & (i + 1)))
					conf = Trim(Request.Form("conf-" & (i + 1)))

					'Make sure that if a pick was made, confidence points were
					'entered and vice-versa.
					if pick = "" and conf <> "" then
						if pickMissingError then
							FormFieldErrors.Add "pick-" & (i + 1), ""
						else
							FormFieldErrors.Add "pick-" & (i + 1), "You must make a pick to set your confidence points."
							pickMissingError = true
						end if
					end if
					if pick <> "" and conf = "" then
						if confMissingError then
							FormFieldErrors.Add "conf-" & (i + 1), ""
						else
							FormFieldErrors.Add "conf-" & (i + 1), "You must enter your confidence points when making a pick<br />for a game."
							confMissingError = true
						end if
					end if
	
					'Validate the confidence points entered, if any.
					if conf <> "" then
						if not IsNumeric(conf) then
							FormFieldErrors.Add "conf-" & (i + 1), "'" & conf & "' is not a valid confidence number."
						elseif CInt(conf) < 1 or CInt(conf) > UBound(games) + 1 or CInt(conf) <> CDbl(conf) then
							FormFieldErrors.Add "conf-" & (i + 1), "'" & conf & "' is not a valid confidence point value."
						end if
					end if
	
				next
	
				'Make sure the confidence points for each pick are unique.
				if FormFieldErrors.Count = 0 then
					redim confUsed(UBound(games) + 1)
					for i = 0 to UBound(games)
						confUsed(i) = false
					next
					for i = 0 to UBound(games)
						conf = GetFieldValue("conf-" & (i + 1), games(i).storedConf)
						if conf <> "" then
							if confUsed(conf) then
								FormFieldErrors.Add "conf-" & (i + 1), "'" & conf & "' is already assigned to a different game."
							else
								confUsed(conf) = true
							end if
						end if
					next
				end if
			end if

			'Make sure no picks were made for locked games.
			if not IsAdmin() then
				for i = 0 to UBound(games)
					if Request.Form("pick-" & (i + 1)) <> "" and games(i).isLocked then
						FormFieldErrors.Add "pick-" & (i + 1), "The " & games(i).visitorName & " at " & games(i).homeName & " game has been locked and<br />cannot be changed."
					end if
				next
			end if

			'If there were any errors, display the error message. Otherwise, do
			'the updates.
			dim sql
			if FormFieldErrors.Count > 0 then
				call FormFieldErrorsMessage("Error: Invalid fields. Please correct and resubmit.")
			else
				for i = 0 to UBound(games)
					pick = GetFieldValue("pick-" & (i + 1), games(i).storedPick)
					conf = GetFieldValue("conf-" & (i + 1), games(i).storedConf)
					if not IsNumeric(conf) then
						conf = "NULL"
					end if
					sql = "DELETE FROM PlayoffsPicks" _
					   & " WHERE Username = '" & SqlString(username) & "'" _
					   & " AND GameID = " & games(i).id
					call DbConn.Execute(sql)
					sql = "INSERT INTO PlayoffsPicks" _
					   & " (Username, GameID, Pick, Confidence)" _
					   & " VALUES('" & SqlString(username) & "'," _
					   & " '" & games(i).id & "', " _
					   & " '" & pick & "'," _
					   & conf & ")"
					call DbConn.Execute(sql)
				next

				'Updates complete, redirect to the results page.
				Response.Redirect("playoffsResults.asp")
			end if
		end if
	end if

	'Build the entry form.
	if username <> "" then
		dim cols
		cols = 8
		if USE_CONFIDENCE_POINTS then
			cols = cols + 1
		end if %>
	<form id="entryForm" action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<div>
<%			if IsAdmin() then %>
			<input type="hidden" name="username" value="<% = username %>" />
<%			end if
			'Add the number of games as a form field (used by the client-side script).
			if USE_CONFIDENCE_POINTS then %>
			<input type="hidden" name="games" value="<% = UBound(games) + 1 %>" />
<%			end if %>
		</div>
		<table><tr><td style="padding: 0px;">
		<table class="main" cellpadding="0" cellspacing="0">
<%		dim lastRound
		dim visitor, home
		dim vpicked, hpicked
		dim correctPick
		dim lockedConf
		dim alt
		lastRound = 0
		alt = false
		for i = 0 to UBound(games)
			if games(i).gameRound <> lastRound then
				lastRound = games(i).gameRound
				alt = false
				if games(i).gameRound = 1 then %>
			<tr class="header bottomEdge">
<%				else %>
			<tr class="header topEdge bottomEdge">
<%				end if %>
				<th align="left" colspan="<% = cols %>"><% = PlayoffRoundNames(games(i).gameRound - 1) %></th>
			</tr>
<%			end if
			if alt then %>
			<tr align="right" class="alt singleLine">
<%			else %>
			<tr align="right" class="singleLine">
<%			end if
			alt = not alt %>
				<td><% = WeekdayName(Weekday(games(i).datetime), true) %></td>
				<td><% = FormatDate(games(i).datetime) %></td>
				<td><% = FormatTime(games(i).datetime) %></td>
<%			visitor = games(i).visitorName
			home = games(i).homeName

			'Get the player's pick data for this game.
			if games(i).isLocked or entryDeleted then
				pick = games(i).storedPick
				conf = games(i).storedConf
			else
				pick = GetFieldValue("pick-" & (i + 1), games(i).storedPick)
				conf = GetFieldValue("conf-" & (i + 1), games(i).storedConf)
			end if

			'Handle the display for a locked game or a game that has not been set.
			if games(i).isLocked or IsNull(games(i).visitorID) or IsNull(games(i).homeID) then

				'Save the raw confidence points value (used by the client-side script).
				lockedConf = conf

				'Highlight fields based on the game result.
				if games(i).result = games(i).visitorID then
					visitor = FormatWinner(visitor)
				elseif games(i).result = games(i).homeID then
					home = FormatWinner(home)
				end if
				if USE_POINT_SPREADS then
					if games(i).atsResult = games(i).visitorID then
						visitor = FormatATSWinner(visitor)
					elseif games(i).atsResult = games(i).homeID then
						home = FormatATSWinner(home)
					end if
				end if


				'Determine the correct pick.
				if USE_POINT_SPREADS then
					correctPick = games(i).atsResult
				else
					correctPick = games(i).result
				end if

				'Set the pick displays.
				vpicked = "&nbsp;"
				hpicked = "&nbsp;"
				if pick = games(i).visitorID then
					vpicked = "X"
				elseif pick = games(i).homeID then
					hpicked = "X"
				end if
				if correctPick <> "" and pick = correctPick then
					if pick = games(i).visitorID then
						vpicked = FormatCorrectPick(vpicked)
					elseif pick = games(i).homeID then
						hpicked = FormatCorrectPick(hpicked)
					end if
				end if

				'Set the confidence points display.
				if USE_CONFIDENCE_POINTS then
					if conf <> "" then
						if pick = correctPick then
							conf = FormatCorrectPick(conf)
						end if
						conf = conf & " pts.&nbsp;&nbsp;"
					end if
				end if

				'Pad the pick and confidence point displays.
				vpicked = "&nbsp;&nbsp;" & vpicked & "&nbsp;&nbsp;"
				hpicked = "&nbsp;&nbsp;" & hpicked & "&nbsp;&nbsp;"
				conf    = "&nbsp;" & conf %>
				<td><% = visitor %></td>
<%				if USE_POINT_SPREADS then %>
				<td><% = FormatPointSpread(games(i).pointSpread) %></td>
<%				end if %>
				<td align="center"><% = vpicked %></td>
				<td><% = GetConjunction(games(i).gameRound) & " " & home %></td>
				<td align="center"><% = hpicked %></td>
<%				'Add the confidence points assigned to the locked game as a
				'form field (used by the client-side script).
				if USE_CONFIDENCE_POINTS then %>
				<td><input type="hidden" name="lockedConf-<% = (i + 1) %>" value="<% = lockedConf %>" /><% = conf %></td>
<%				end if

			'Handle the display for an unlocked game.
			else %>
				<td><label for="pick-<% = (i + 1) %>-V"><% = visitor %></label></td>
<%				if USE_POINT_SPREADS then %>
				<td><label for="pick-<% = (i + 1) %>-V"><% = FormatPointSpread(games(i).pointSpread) %></label></td>
<%				end if %>
				<td align="center"><div class="<% = FieldStyleClass("", "pick-" & (i + 1)) %>"><input type="radio" id="pick-<% = (i + 1) %>-V" name="pick-<% = (i + 1) %>" value="<% = games(i).visitorID %>"<% if pick = games(i).visitorID then Response.Write(CHECKED_ATTRIBUTE) end if %> /></div></td>
				<td><% = GetConjunction(games(i).gameRound) %> <label for="pick-<% = (i + 1) %>-H"><% = home %></label></td>
				<td align="center"><div class="<% = FieldStyleClass("", "pick-" & (i + 1)) %>"><input type="radio" id="pick-<% = (i + 1) %>-H" name="pick-<% = (i + 1) %>" value="<% = games(i).homeID %>"<% if pick = games(i).homeID then Response.Write(CHECKED_ATTRIBUTE) end if %> /></div></td>
<%				if USE_CONFIDENCE_POINTS then %>
				<td align="right">
					<select name="conf-<% = (i + 1) %>" class="<% = FieldStyleClass("numeric", "conf-" & (i + 1)) %>" onchange="confidencePointsUpdate();">
<%					call DisplayConfidencePointsList(6, UBound(games) + 1, conf) %>
					</select>&nbsp;pts.&nbsp;&nbsp;
				</td>
<%				end if
			end if %>
			</tr>
<%		next

		'Add a row for dynamically displaying the available confidence points.
		if USE_CONFIDENCE_POINTS and (IsAdmin() or not allLocked) then %>
			<tr class="subHeader topEdge" style="display: none;">
				<th id="pointsList" colspan="<% = cols %>"></th>
			</tr>
<%		end if

		'Build the player's score display, when available.
		dim score, numGames, pctStr, scoreStr
		numGames = NumberOfCompletedPlayoffGames()
		scoreStr = ""
		pctStr = ""
		if numGames > 0 then
			score = PlayoffsPlayerPickScore(username)
			if IsNumeric(score) and IsNumeric(numGames) then
				scoreStr = score & "/" & numGames
				pctStr = " (" & FormatPercentage(score / numGames) & ")"
				if USE_CONFIDENCE_POINTS then
					score = PlayoffsPlayerConfidenceScore(username)
					if IsNumeric(score) then
						scoreStr = scoreStr & pctStr & "&nbsp;&nbsp;<strong>" & FormatScore(score, false) & " pts. </strong>"
					end if
				else
					scoreStr = "<strong>" & scoreStr & "</strong>" & pctStr
				end if
			end if 
		end if
		if scoreStr <> "" then %>
			<tr class="header topEdge bottomEdge">
				<th align="left" colspan="<% = cols %>">Score</th>
			</tr>
			<tr>
				<td align="right" colspan="<% = cols %>"><% = scoreStr %></td>
			</tr>
<%		end if %>
		</table>
<%		'Show form buttons, if changes are allowed.
		if IsAdmin() or not allLocked then %>
		<p></p>
		<table cellpadding="0" cellspacing="0" style="width: 100%;">
			<tr valign="middle">
				<td style="padding: 0px;"><input type="submit" name="submit" value="Update" class="button" title="Apply changes." />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel the update." /></td>
<%			'If the user has an entry and no games are locked (or the user is the
			'Admin), add a delete button.
			if hasEntry and (IsAdmin() or not anyLocked) then

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
<%	end if %>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>
<%	'**************************************************************************
	'* Local class definitions.                                               *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' GameObj: Holds information for a single game.
	'--------------------------------------------------------------------------
	class GameObj

		public id, gameRound, datetime, visitorID, homeID, pointSpread
		public result, atsResult
		public visitorName, homeName, conference, isLocked
		public storedPick, storedConf

		private sub Class_Initialize()
		end sub

		private sub Class_Terminate()
		end sub

		public sub setData(rs)

			'Set the game properties using the supplied database record.
			id          = rs.Fields("GameID").Value
			gameRound   = rs.Fields("Round").Value
			datetime    = CDate(rs.Fields("Date").Value & " " & rs.Fields("Time").Value)
			visitorID   = rs.Fields("VisitorID").Value
			pointSpread = rs.Fields("PointSpread").Value
			homeID      = rs.Fields("HomeID").Value
			result      = rs.Fields("Result").Value
			atsResult   = rs.Fields("ATSResult").Value
			isLocked    = false
			storedPick  = ""
			storedConf  = ""

			'Get the conference of the teams playing.
			conference = GetConference(homeID)

			'Set the team names for display.
			visitorName = GetTeamName(visitorID)
			homeName    = GetTeamName(homeID)

			'If the user is not the Administrator, set the game's lock status.
			if not IsAdmin() and RoundLocked(gameRound) then
				anyLocked = true
				isLocked = true
			else
				allLocked = false
			end if

			'Get the user's pick data for this game.
			dim sql, rs2
			sql = "SELECT Pick, Confidence FROM PlayoffsPicks" _
			   & " WHERE Username = '" & SqlString(username) & "'" _
			   & " AND GameID = " & id
			set rs2 = DbConn.Execute(sql)
			if not (rs2.EOF and rs2.BOF) then
				storedPick = rs2.Fields("Pick").Value
				storedConf = rs2.Fields("Confidence").Value
			end if

		end sub

	end class %>