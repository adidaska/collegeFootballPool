	<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Entry Form" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
    
<link rel="shortcut icon" href="favicon.ico" />
	<!--<link rel="stylesheet" type="text/css" href="styles/common.css" />-->
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
    <link href="styles/style.css" rel="stylesheet" type="text/css" />


</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/weekly.asp" -->
 
<!-- start of the db pull -->
<%	'Open the database.
	call OpenDB()

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

	'For the Administrator, build a user selection list.
	dim users, i
	if IsAdmin() then %>
    
    <!-- start of the admin table -->
		<div class="admin-entry-form">
	        <form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
	            <div class="header bottomEdge">
	                <div class="admin-entry-form"><span>Administrator Access</span></div>
	            </div>
	            <div>You may view or edit any player's entry form by selecting a username below.<br />
	            Use the links at the bottom of the page to switch to a specific week.</div>
	            <div align="center">
	               <br/>
	              <strong>Select user:</strong>
	               <input type="hidden" name="week" value="<% = week %>" />
	              <select name="username">
	                   <option value=""></option>
	                   <% users = UsersList(true)
	                   if IsArray(users) then
	                      for i = 0 to UBound(users) %>
	                      <option value="<% = users(i) %>" <% if users(i) = username then Response.Write(" selected=""selected""") end if %>><% = users(i) %></option>
	                      <%	next
	                   end if %>
	              </select>						
	               <input type="submit" name="submit" value="Select" class="button" title="View/edit the selected user's entry." /></td>      
	          </div>		
	        </form>
        </div>
      
      
      
	
    
    

<!-- end of the admin table -->
 	<!-- this is the dialog to pop up after submission-->
 	<div id="dialog">
     	<div>
          	<p>Content you want the user to see goes here.</p>
     	</div>
	</div>
 
 
 
<!-- start of the logic for the db pull -->
<%		if username <> "" then %>
<h2 class="welcome-message">Entry Form for <% = username %></h2>
<%		end if
	end if

	'Determine if the user has an entry for the given week (so we can display
	'the Delete button).
	dim hasEntry, visTB, homeTB, infoTB
	hasEntry = false
	'set infoTB = new TBclass
	visTB = TBvisGuess(username, week)
	homeTB = TBhomeGuess(username, week)
	if username <> "" and visTB <> "" and homeTB <> "" then
		hasEntry = true
	end if

	'Initialize lock flags.
	dim allLocked, anyLocked
	allLocked = AllGamesLocked(week)
	anyLocked = allLocked

	'Create an array of game objects
	dim games, rs, n
	n = NumberOfGames(week)
	redim games(n - 1)
	set rs = WeeklySchedule(week)
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
	dim sql
	dim needDeleteConfirmation, entryDeleted
	dim pick, conf, tb, tb1, tb2
	dim pickMissingError, confMissingError
	dim confUsed
	needDeleteConfirmation = false
	entryDeleted = false

	'Process a delete request.
	if Request.Form("submit") = "Delete" then

		if not IsAdmin() and anyLocked then
			call ErrorMessage("Error: One or more games for this week have been locked, entry<br />cannot be deleted.")
		elseif LCase(Request.Form("confirmDelete")) <> "true" then
			needDeleteConfirmation = true
			call ErrorMessage("To confirm deletion, check the box below and press <code>Delete</code> again.<br />Pressing any other button will cancel the deletion.")
		elseif NumberOfCompletedGames(week) = NumberOfGames(week) then
			call ErrorMessage("Error: Cannot delete entries for completed weeks.")
		else

			'Delete the player's entry.
			for i = 0 to UBound(games)
				sql = "DELETE FROM Picks" _
				   & " WHERE Username = '" & SqlString(username) & "'" _
				   & " AND GameID = " & games(i).id
				call DbConn.Execute(sql)
			next
			sql = "DELETE FROM Tiebreaker" _
			    & " WHERE Week = " & week _
			    & " AND Username = '" & SqlString(username) & "'"
			call DbConn.Execute(sql)
			entryDeleted = true
			hasEntry = false

			'Clear any saved weekly results data.
			call ClearWeeklyResultsCache(week)

			'Reload the games.
			redim games(n - 1)
			set rs = WeeklySchedule(week)
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

		'Prevent entries by the Admin.
		if username = ADMIN_USERNAME then
			call ErrorMessage("Error: User '" & ADMIN_USERNAME & "' may not make picks.")

		'If the user is disabled, prevent the entry.
		elseif IsDisabled() then
			call ErrorMessage("Error: Your account has been disabled, changes not accepted.<br />Please contact the Administrator.")

		'If all games have been locked, prevent any updates (except by the Admin).
		elseif not IsAdmin() and allLocked then
			call ErrorMessage("Error: All games for this week have been locked, changes not<br />accepted.")

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

			'Validate the tiebreaker field.
			' So if the fields are blank then we want to put in a 0
			'if tb1 = "" and not allLocked then
            '    'tb1 = 0
            '    FormFieldErrors.Add "tb-" & week, "it thinks this is empty"
            'else
            '    tb1 = Trim(Request.Form("tb1-" & week))
            'end if
            'if tb2 = "" and not allLocked then
            '    tb2 = 0
            'else
            '    tb2 = Trim(Request.Form("tb2-" & week))
            'end if

            tb1 = Trim(Request.Form("tb1-" & week))
            tb2 = Trim(Request.Form("tb2-" & week))


			if tb1 = "" or tb2 = "" then
				'FormFieldErrors.Add "tb-" & week, "You must enter a point total for the tiebreaker."
				if tb1 = "" then
				    tb1 = 0
				end if
				if tb2 = "" then
				    tb2 = 0
				end if
			elseif not IsNumeric(tb1) or not IsNumeric(tb2) then
				FormFieldErrors.Add "tb1-" & week, "'" & tb1 & "' is not a valid amount."
			elseif CInt(tb1) < 0 or CInt(tb1) <> CDbl(tb1) then
				FormFieldErrors.Add "tb-" & week, "'" & tb1 & "' is not a valid amount."
			end if

			'Make sure no picks were made for locked games.
			if not IsAdmin() and not allLocked then
				for i = 0 to UBound(games)
					if Request.Form("pick-" & (i + 1)) <> "" and games(i).isLocked then
						FormFieldErrors.Add "pick-" & (i + 1), "The " & games(i).visitorName & " at " & games(i).homeName & " game has been locked and<br />cannot be changed."
					end if
				next
			end if

			'If there were any errors, display the error messages. Otherwise, do
			'the updates.
			if FormFieldErrors.Count > 0 then
				call FormFieldErrorsMessage("Error: Invalid fields. Please correct and resubmit.")
			else
				for i = 0 to UBound(games)
					pick = GetFieldValue("pick-" & (i + 1), games(i).storedPick)
					conf = GetFieldValue("conf-" & (i + 1), games(i).storedConf)
					if not IsNumeric(conf) then
						conf = "NULL"
					end if
					sql = "DELETE FROM Picks" _
					   & " WHERE Username = '" & SqlString(username) & "'" _
					   & " AND GameID = " & games(i).id
					call DbConn.Execute(sql)
					sql = "INSERT INTO Picks" _
					   & " (Username, GameID, Pick, Confidence)" _
					   & " VALUES('" & SqlString(username) & "'," _
					   & " " & games(i).id & ", " _
					   & " '" & pick & "'," _
					   & conf & ")"
					call DbConn.Execute(sql)
				next
				sql = "DELETE FROM Tiebreaker" _
				    & " WHERE Week = " & week _
				    & " AND Username = '" & SqlString(username) & "'"
				call DbConn.Execute(sql)
				sql = "INSERT INTO Tiebreaker" _
				   & " (Week, Username, VisGuess, HomeGuess)" _
				   & " VALUES(" & week & ", " _
				   & "'" & SqlString(username) & "', " _
				   & tb1 & ", " & tb2 & ")"

				call DbConn.Execute(sql)

				'Clear any saved weekly results data.
				call ClearWeeklyResultsCache(week)

				'Updates complete, redirect to the results page.
				if IsAdmin() then
					'toggleDialog()
					Response.Redirect("poolResults.asp?week=" & week & "&username=" & Server.URLEncode(username))
				else
					'toggleDialog()
					Response.Redirect("poolResults.asp?week=" & week)
				end if
			end if
		end if
	end if



	'Build the entry form for that week.
	'if username <> "" then
		dim cols
		cols = 7
		if ALLOW_TIE_PICKS then
			cols = cols + 2
		end if
		if USE_CONFIDENCE_POINTS then
			cols = cols + 1
		end if
		if USE_POINT_SPREADS then
			cols = cols + 1
		end if %>
<!-- end of the first logic section table -->






	<form id="entryForm" action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<div>
			<%			if IsAdmin() then %>
			<input type="hidden" name="username" value="<% = username %>" />
			<%			end if %>
			<input type="hidden" name="week" value="<% = week %>" />
			<%			'Add the number of games as a form field (used by the client-side script).
			if USE_CONFIDENCE_POINTS then %>
			<input type="hidden" name="games" value="<% = UBound(games) + 1 %>" />
			<%			end if %>
		</div>



<!-- start of the games table -->

        <% 'first get the stored tiebreaker values from the database to we have when we start the table building since they are in a different able
        dim TB1value, TB2value

        TB1value = 0
        TB2value = 0

        TB1value = TBvisGuess(username, week)
        TB2value = TBhomeGuess(username, week)

        %>


<div class="clearfix" id="content-wrap">
  	<div id="content-top"></div>
    <div id="primary" class="hfeed">
        <div id="gameContainers">
        <span class="versusText">Week <% = week %></span><p>

            <%
            dim visitor, home, tie, logoHome, logoVis, tbgameValue
			dim vpicked, hpicked, tpicked
			dim correctPick
			dim giveTie, halfStr
			dim lockedConf
			dim alt
			halfStr = "<span style=""font-family: monospace;"">" & HALF_POINTS & "</span>"

	if(UBound(games) = -1) then
		call errormessage(n) %>

    	<div>There are no games currently on the schedule</div>
	<% else
		'call errormessage(n)
	    'here is the start of the FOR loop
		for i = 0 to UBound(games)



			'call ErrorMessage("games(i) =" & i)

			visitor = games(i).visitorName
			home = games(i).homeName
			logoHome = games(i).logoHome
			logoVis = games(i).logoVis
			tbgameValue = games(i).tbgame

			'Get the player's pick data for this game.
			if games(i).isLocked or entryDeleted then
				pick = games(i).storedPick
				conf = games(i).storedConf
			else
				pick = GetFieldValue("pick-" & (i + 1), games(i).storedPick)
				conf = GetFieldValue("conf-" & (i + 1), games(i).storedConf)
			end if
			%>


            <%'Handle the display for a locked game.
			if games(i).isLocked then


				if pick = TIE_STR then
					tie = TIE_STR
				else
					tie = "&nbsp;"
				end if

				'Save the raw confidence points value (used by the client-side script).
				lockedConf = conf

				'Highlight fields based on the game result.
				if games(i).result = games(i).visitorID then
					visitor = FormatWinner(visitor)
				elseif games(i).result = games(i).homeID then
					home = FormatWinner(home)
				elseif games(i).result = TIE_STR then
					tie = FormatWinner(TIE_STR)
				end if
				if USE_POINT_SPREADS then
					if games(i).atsResult = games(i).visitorID then
						visitor = FormatATSWinner(visitor)
					elseif games(i).atsResult = games(i).homeID then
						home = FormatATSWinner(home)
					elseif games(i).atsResult = TIE_STR then
						tie = FormatATSWinner(TIE_STR)
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
				tpicked = "&nbsp;"
				hpicked = "&nbsp;"
				if pick = games(i).visitorID then
					vpicked = "X"
				elseif pick = games(i).homeID then
					hpicked = "X"
				elseif pick = TIE_STR then
					tpicked = "X"
				end if
				if correctPick <> "" and pick = correctPick then
					if pick = games(i).visitorID then
						vpicked = FormatCorrectPick(vpicked)
					elseif pick = games(i).homeID then
						hpicked = FormatCorrectPick(hpicked)
					else
						tpicked = FormatCorrectPick(tpicked)
					end if
				end if

				'Set the confidence points display.
				if USE_CONFIDENCE_POINTS then
					if conf <> "" then
						conf = conf & "&nbsp;pts."
						if pick = correctPick then
							conf = FormatCorrectPick(conf)
						end if
					end if
				end if

				'If tie picks are not allowed and the game result was a tie,
				'give the player the pick.
				giveTie = false
				if not ALLOW_TIE_PICKS and correctPick = TIE_STR and pick <> "" then
					giveTie = true
				end if

				'If tie picks are not allowed and the game result was a tie,
				'give the player the pick.
				if giveTie then
					if pick = games(i).visitorID then
						vpicked = vpicked & halfStr
					else
						hpicked = hpicked & halfStr
					end if
					conf = conf & halfStr
				else
					vpicked = vpicked & "&nbsp;&nbsp;"
					hpicked = hpicked & "&nbsp;&nbsp;"
					tpicked = tpicked & "&nbsp;&nbsp;"
					conf    = conf    & "&nbsp;&nbsp;"
				end if

				'Pad the pick and confidence point displays.
				vpicked = "&nbsp;&nbsp;" & vpicked
				hpicked = "&nbsp;&nbsp;" & hpicked
				tpicked = "&nbsp;&nbsp;" & tpicked
				conf    = "&nbsp;" & conf %>

                <%'Add a row for dynamically displaying the available confidence points.
                if USE_CONFIDENCE_POINTS and (IsAdmin() or not allLocked) then %>
        <!--          <tr class="subHeader topEdge" style="display: none;">
                    <th id="pointsList" colspan="<% = cols %>"></th>
                  </tr>-->
                  <%		end if

        '		''Build the player's score display, when available.
                dim score, numGames, pctStr, scoreStr
                numGames = NumberOfCompletedGames(week)
                scoreStr = ""
                pctStr = ""
                if numGames > 0 then
                    score = PlayerPickScore(username, week)
                    if IsNumeric(score) and IsNumeric(numGames) then
                        scoreStr = score & "/" & numGames
                        pctStr = " (" & FormatPercentage(score / numGames) & ")"
                        if USE_CONFIDENCE_POINTS then
                            score = PlayerConfidenceScore(username, week)
                            if IsNumeric(score) then
                                scoreStr = scoreStr & pctStr & "&nbsp;&nbsp;<strong>" & FormatScore(score, false) & " pts. </strong>"
                            end if
                        else
                            scoreStr = "<strong>" & scoreStr & "</strong>" & pctStr
                        end if
                    end if
                end if
                if not IsAdmin() and scoreStr <> "" then %>
        <!--          <tr class="header topEdge bottomEdge">
                    <th align="left" colspan="<%' = cols %>">Score</th>
                  </tr>
                  <tr>
                    <td align="right" colspan="<%' = cols %>"><% = scoreStr %></td>
                  </tr>-->
                <% end if

                'Get the tiebreaker game, pull the game in the db where tbgame field = 1
                dim tbTitle, tbTitle2, tbStr, tbStr2, tbgame, tblocked, loopVar
                tblocked = ""

                'need to get the gameid from the games array where the tbgame = 1
                for loopVar = 0 to (n-1)
                    if games(loopVar).TBgame = 1 then
                        tbgame = loopVar
                    end if
                    'call errormessage("the number of games in week " & week & " is " & n)
                    'call errormessage("hi" & tbgame & games(loopVar).visitorName)
                next

                'call errormessage(tbgame & games(tbgame).visitorName)
                tbTitle = "Enter game score: <strong>" & games(cint(tbgame)).visitorName  & "</strong>"
                tbTitle2 = " at <strong>" & games(cint(tbgame)).homeName & "</strong>"


                'Get the user's tiebreaker guess for this week.
                dim storedTb1, storedTb2
                storedTb1 = TBvisGuess(username, week)
                storedTb2 = TBhomeGuess(username, week)
                if not IsAdmin() and allLocked then
                    tb1 = storedTb1
                    tb2 = storedTb2
                else
                    tb1 = GetFieldValue("tb1-" & week, storedTb1)
                    tb2 = GetFieldValue("tb2-" & week, storedTb2)
                end if
                if entryDeleted then
                    tb1 = 0
                    tb2 = 0
                end if
                'call errormessage(storedTb1)
                'call errormessage(storedTb2)

                'Get the tiebreaker data.
                if allLocked then

                    'Check for a tiebreaker point total.
                    dim tbActualHome, tbActualVis
                    tbActualHome = TBPointTotalHome(week)
                    tbActualVis = TBPointTotalVis(week)
                    if IsNumeric(tb) and IsNumeric(tbActualHome) and IsNumeric(tbActualVis) then
                        tbStr = tb1 & "&nbsp;<strong>(" & 1 & ")</strong>"  'Abs((tbActualVis - tb1) + (tbActualHome - tb2)
                        tbStr2 = tb2 & "&nbsp;<strong>(" & 1 & ")</strong>"  'Abs((tbActualVis - tb1) + (tbActualHome - tb2)
                    elseif tb1 = "" or tb2 = "" then
                        tbStr = "n/a"
                        tbStr2 = "n/a"
                    else
                        tbStr = tb1
                        tbStr2 = tb2
                    end if

                end if %>



             <!-- this is the locked game -->
              <div class="gameWrapper">
              <%if USE_POINT_SPREADS then %>
                 <div class="spreadContainer">
                     <div class="spread"><% = FormatPointSpread(games(i).pointSpread) %></div>
                 </div>
             <%end if %>
                    <div class="tablecontainer" class="locked">
                    <div class="trow locked">
                        <div class="tleft locked">
                            <!-- <p class="visitorName"> -->
                            <% '= visitor %><!--  visitor name -->
                            <!-- </p> -->

                            <!-- here is the visitor logo -->
                           <img src="https://a.espncdn.com/combiner/i?img=/i/teamlogos/ncaa/500/<% = logoVis %>.png&h=250" alt="teamLogoVis" class="tlogo" />
                            <input type="radio" id="pick-<% = (i + 1) %>-V" name="pick-<% = (i + 1)%>" value="<% = games(i).visitorID %>" disabled="disabled" class="tradio"
                              <% if pick = games(i).visitorID then Response.Write(CHECKED_ATTRIBUTE) end if %> />
                            <!-- here is the visitor radio button -->
                        </div>

                        <div class="tmiddle locked">
                            <span class="entry-title"><% = visitor %> at <% = home %></span>
                            <h4 class="game-Date"><% = WeekdayName(Weekday(games(i).datetime), true) & " " & FormatDate(games(i).datetime) & " " & FormatTime(games(i).datetime) %></h4>
                        </div>

                        <div class="tright locked">
                            <!-- <p><% = home %></p> -->
                            <img src="https://a.espncdn.com/combiner/i?img=/i/teamlogos/ncaa/500/<% = logoHome %>.png&h=250" alt="teamLogoHome" class="tlogo" />
                            <input type="radio" id="pick-<% = (i + 1) %>-H" name="pick-<% = (i + 1) %>" value="<% = games(i).homeID %>" disabled="disabled" class="tradio"
                            <% if pick = games(i).homeID then Response.Write(CHECKED_ATTRIBUTE) end if %> />
                        </div>
                    </div><!--END div id trow -->
                    <!-- Add TieBreaker section if this is the tiebreaker game -->
                    <% 'tiebreaker game check
                    if tbgameValue = 1 then %>

                        <div class="tbrow trow locked">
                            <div class="tbleft tleft locked">
                                <input type="text" name="tb1-<% = week %>" value="<% = TB1value %>" size="5" class="tbleft" disabled="disabled" />
                            </div>

                            <div class="tbmiddle tmiddle locked">
                                <span class="tb-title entry-title">TIEBREAKER GAME</span>
                                <span class="tb-subtitle entry-title">Guess the Game Score</span>
                            </div>

                            <div class="tbright tright locked">
                                <input type="text" name="tb2-<% = week %>" value="<% = TB2value %>" size="5" class="tbright" disabled="disabled" readonly="readonly"/>
                            </div>
                        </div>

                    <%end if %>

                </div>
              </div>



		    <%'Handle the display for an unlocked game.
				else %>

	           <!-- this is an unlocked game unselected-->
                <div class="gameWrapper">
                <%if USE_POINT_SPREADS then %>
                       <div class="spreadContainer">
                           <div class="spread"><% = FormatPointSpread(games(i).pointSpread) %></div>
                       </div>
                   <%end if %>
                <div class="tablecontainer" class="unlocked">
                      <div class="trow">
                        <div class="tleft">
                                <!-- here is the visitor logo -->

                                <img src="https://a.espncdn.com/combiner/i?img=/i/teamlogos/ncaa/500/<% = logoVis %>.png&h=250" alt="teamLogoVis" class="tlogo" />
                                <!-- here is the visitor radio button -->
                                <input type="radio" id="pick-<% = (i + 1) %>-V" name="pick-<% = (i + 1)%>" value="<% = games(i).visitorID %>" class="tradio"
                                <% if pick = games(i).visitorID then Response.Write(CHECKED_ATTRIBUTE) end if %> />
                                <!-- this above line will check the radio button if selection already in db -->
                                <% 'call errormessage(games(i).visitorID) %>
                        </div>

                        <div class="tmiddle">
                            <span class="entry-title"><% = visitor %> at <% = home %></span><!-- need the display name for this one -->
                            <h4 class="game-Date"><% = WeekdayName(Weekday(games(i).datetime), true) & " " & FormatDate(games(i).datetime) & " " & FormatTime(games(i).datetime) %></h4>
                        </div>

                        <div class="tright">
                            <!-- <p><% = home %></p> -->
                            <img src="https://a.espncdn.com/combiner/i?img=/i/teamlogos/ncaa/500/<% = logoHome %>.png&h=250" alt="teamLogoHome" class="tlogo" />
                            <input type="radio" id="pick-<% = (i + 1) %>-H" name="pick-<% = (i + 1) %>" value="<% = games(i).homeID %>" class="tradio"
                                <% if pick = games(i).homeID then Response.Write(CHECKED_ATTRIBUTE) end if %> />
                        </div>
                      </div><!--END div id trow -->
                      <!-- Add TieBreaker section if this is the tiebreaker game -->
                        <% 'tiebreaker game check
                        if tbgameValue = 1 then %>

                            <div class="tbrow trow">
                                <div class="tbleft tleft">
                                    <input type="text" name="tb1-<% = week %>" value="<% = TB1value %>" size="5" class="tbleft" />

                                </div>

                                <div class="tbmiddle tmiddle">
                                    <span class="tb-title entry-title">TIEBREAKER GAME</span>
                                    <span class="tb-subtitle entry-title">Guess the Game Score</span>
                                </div>

                                <div class="tbright tright">
                                    <input type="text" name="tb2-<% = week %>" value="<% = TB2value %>" size="5" class="tbright" />
                                </div>
                            </div>

                        <%end if %>
                </div>
                </div>
			<%	end if
			next %>
            </div>

		  
	  <%		'List open dates.
		'call DisplayOpenDates(2, week)

		'Show form buttons, if changes are allowed.
		
		if IsAdmin() or not allLocked then %>
			<p></p>
			<table cellpadding="0" cellspacing="0" style="width: 100%;">
                <tr valign="middle">
                    <td style="padding: 0px;">
                        <input type="submit" name="submit" value="Update" class="button" title="Apply changes." />&nbsp;
                        <input type="submit" name="submit" value="Cancel" class="button" title="Cancel the update." />
                  </td>
                <%'If the user has an entry and no games are locked (or the user is the
                'Admin), add a delete button.
                if hasEntry and (IsAdmin() or not anyLocked) then
    
                    'If delete was requested, add a confirmation checkbox.
	                  if needDeleteConfirmation then %>
                        <td align="right" style="padding: 0px;">
                        	<input type="checkbox" id="confirmDelete" name="confirmDelete" value="true" /> <label for="confirmDelete">Confirm Deletion</label>&nbsp;
                        </td>
                    <%	end if %>
                        <td align="right" style="padding: 0px;">
                        	<input type="submit" name="submit" value="Delete" class="button" title="Delete this entry." />
                        </td>
                <% end if %>
                </tr>
			</table>
		<%	end if %>
		</td></tr></table>

	<%	'end if%>
	</form>


    </div> <!-- end of the primary div in the container-->

    <% end if %>
    
    
<div id="content-btm"></div>
<!--</div>-->
        
<% 'List links to view other weeks.
	dim params
	params = ""
	if IsAdmin() then
		params = "username=" & Server.HtmlEncode(username)
	end if
	call DisplayWeekNavigation(1, params)
	 %>
    
        <p>
          <!-- #include file="includes/footer.asp" -->
      	</p>      
      
          
          
</body>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>
<script type="text/javascript" src="scripts/common.js"></script>
<script type="text/javascript" src="scripts/menu.js"></script>
<script type="text/javascript" src="scripts/app.js"></script>
<%	if USE_CONFIDENCE_POINTS then %>
    <script type="text/javascript" src="scripts/confidencePoints.js"></script>
<%	end if %>
</html>

<link href='https://fonts.googleapis.com/css?family=Roboto:900,300,500,400|Archivo+Narrow:700' rel='stylesheet' type='text/css'>

<%	'**************************************************************************
	'* Local class definitions.                                               *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' GameObj: Holds information for a single game.
	'--------------------------------------------------------------------------
	class GameObj

		public id, datetime, visitorID, visitorName, homeID, homeName, pointSpread, logoHome, logoVis, tbgame, displayName
		public result, atsResult
		public isLocked
		public storedPick, storedConf

		private sub Class_Initialize()
		end sub

		private sub Class_Terminate()
		end sub

		public sub setData(rs)

			'Set the game properties using the supplied database record.
			id          = rs.Fields("GameID").Value
			datetime    = CDate(rs.Fields("Date").Value & " " & rs.Fields("Time").Value)
			visitorID   = rs.Fields("VisitorID").Value
			visitorName = rs.Fields("VCity").Value
			homeID      = rs.Fields("HomeID").Value
			homeName    = rs.Fields("HCity").Value
			pointSpread = rs.Fields("PointSpread").Value
			result      = rs.Fields("Result").Value
			atsResult   = rs.Fields("ATSResult").Value
			logoHome	= rs.Fields("logoHome").Value
			logoVis		= rs.Fields("logoVis").Value
			tbgame		= rs.Fields("tbgame").Value
			'displayName	= rs.Fields("DisplayValue").Value
			isLocked    = false
			storedPick  = ""
			storedConf  = ""

			'Set the team names for display.
			if rs.Fields("VDisplayName").Value <> "" then
				visitorName = rs.Fields("VDisplayName").Value
			end if
			if rs.Fields("HDisplayName").Value <> "" then
				homeName = rs.Fields("HDisplayName").Value
			end if

			'If the user is not the Administrator, set the game's lock status.
			if not IsAdmin() and (allLocked or GameStarted(datetime)) then
				anyLocked = true
				isLocked = true
				if tbgame = 1 then
					tblocked = "DISABLED"
				end if
				
			end if

			'Get the user's pick data for this game.
			dim sql, rs2
			sql = "SELECT Pick, Confidence FROM Picks" _
			   & " WHERE Username = '" & SqlString(username) & "'" _
			   & " AND GameID = " & id
			set rs2 = DbConn.Execute(sql)
			if not (rs2.EOF and rs2.BOF) then
				storedPick = rs2.Fields("Pick").Value
				storedConf = rs2.Fields("Confidence").Value
			end if
			
			
			'get the displayValue for the game
			'dim sql, rs3
			'sql = "SELECT DisplayValue FROM FullSchedule" _
			 '  & " WHERE week = '" & week & "'" _
			 '  & " AND VisTeam = " & id
			'set rs3 = DbConn.Execute(sql)
			'if not (rs3.EOF and rs3.BOF) then
			'	displayValue = rs3.Fields("DisplayValue").Value
			'end if

		end sub

	end class 
	
	

%>









