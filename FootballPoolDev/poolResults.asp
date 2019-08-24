<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Results" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/style.css" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />

	<!--<link rel="stylesheet" href="styles/newstyles/reset.css" type="text/css" media="screen" title="no title" />-->
	<!--<link rel="stylesheet" href="styles/newstyles/text.css" type="text/css" media="screen" title="no title" />-->
	<!--<link rel="stylesheet" href="styles/newstyles/form.css" type="text/css" media="screen" title="no title" />-->
	<!--<link rel="stylesheet" href="styles/newstyles/buttons.css" type="text/css" media="screen" title="no title" />-->
	<!--<link rel="stylesheet" href="styles/newstyles/grid.css" type="text/css" media="screen" title="no title" />	-->
	<!--<link rel="stylesheet" href="styles/newstyles/layout.css" type="text/css" media="screen" title="no title" />	-->
	<!--<link rel="stylesheet" href="styles/newstyles/ui-darkness/jquery-ui-1.8.12.custom.css" type="text/css" media="screen" title="no title" />-->
	<!--<link rel="stylesheet" href="styles/newstyles/plugin/jquery.visualize.css" type="text/css" media="screen" title="no title" />-->
	<!--<link rel="stylesheet" href="styles/newstyles/plugin/facebox.css" type="text/css" media="screen" title="no title" />-->
	<!--<link rel="stylesheet" href="styles/newstyles/plugin/uniform.default.css" type="text/css" media="screen" title="no title" />-->
	<!--<link rel="stylesheet" href="styles/newstyles/plugin/dataTables.css" type="text/css" media="screen" title="no title" />-->
	<!--<link rel="stylesheet" href="styles/newstyles/custom.css" type="text/css" media="screen" title="no title">-->
    <link rel="stylesheet" type="text/css" href="//cdn.datatables.net/responsive/1.0.0/css/dataTables.responsive.css">



	<link href="styles/style.css" rel="stylesheet" type="text/css" />
</head>

<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/weekly.asp" -->
	<!--<table class="left_aligned_style" id="wrapper"><tr><td style="padding: 0px;">-->
    <div id="wrapper">
    	<div id="content" class="xgrid">

    		<div class="x12">




        <%	'Open the database.
            call OpenDB()

            'Get the week to display.
            dim week
            week = GetRequestedWeek()

            'Find the number of games for this week.
            dim numGames
            numGames = NumberOfGames(week)

            'See if this user is missing any picks for this week.
            dim username
            dim tbActualVis, tbActualHome, tbGuessVis, tbGuessHome
            dim sql, rs
            username = Session(SESSION_USERNAME_KEY)

            if IsAdmin() then
                username = Request.QueryString("username")
            end if

            'set tbActual = new TBclass
            tbActualVis = TBPointTotalVis(week) 'nee to check on this one
            tbActualHome = TBPointTotalHome(week)
            tbGuessVis  = TBvisGuess(username, week)
            tbGuessHome  = TBhomeGuess(username, week)

            'this is doing a count of the total games and checking if the same number of picks exist
            if username <> "" then
                if tbActualHome = "" and tbActualVis = "" and tbGuessVis <> "" and tbGuessHome = "" then
                    sql = "SELECT COUNT(*) AS Total" _
                       & " FROM Picks, Schedule" _
                       & " WHERE Username = '" & SqlString(username) & "'" _
                       & " AND Week = " & week _
                       & " AND Picks.GameID = Schedule.GameID" _
                       & " AND Pick <> ''"
                    set rs = DbConn.Execute(sql)
                    if not (rs.BOF and rs.EOF) then
                        if rs.Fields("Total").Value < numGames then
                                call ErrorMessage("Warning: You are missing one or more picks for this week.")
                        end if
                    end if
                end if
            end if

            'Determine if all games for the week have been locked.
            dim allLocked
            allLocked = AllGamesLocked(week)

            'Build the display.
            dim gameCols, otherCols
            gameCols  = numGames
            otherCols = 5
            if USE_CONFIDENCE_POINTS then
                otherCols = otherCols + 2
            end if %>


    <% call DisplayWeekNavigation(1, "")  %>





     <table class="data display datatable" id="example">
        <thead>
                    <tr>
                        <th id="name" data-sort="string">Name</th>
                        <th id="wins" data-sort="float" data-sort-onload=yes data-sort-multicolumn="diff" data-sort-default="desc">Wins</th>
                        <th id="games" data-sort="int">Games</th>
                        <th id="percentage" data-sort="float">Pct</th>
                        <th id="diff" data-sort="int" data-sort-default="desc">TB Diff</th>
                        <th id="scores" data-sort="string">TB Scores</th>
                        <% dim g
        					for g = 1 to gameCols %>
        					<th data-sort="string"><% = g %></th>
                        <% next %>

                    </tr>
                </thead>

		<!--<tbody id="poolResults">-->

        <tbody>


<%	dim completedGames, totalGames, list
	dim users, currentUser
	dim leadConfScore
	dim i, j
	dim gameDatetime
	dim pick, conf, result, correctPick, giveTie
	dim pickScore, pickPct, confScore, confScoreDiff, tbDiff
	dim hidePick
	completedGames = NumberOfCompletedGames(week)
	totalGames = NumberOfGames(week)
	currentUser = Session(SESSION_USERNAME_KEY)

	'Get user data.
	list = PoolPlayersList(week)
	if IsArray(list) then

		'Create an array of user objects.
		redim users(UBound(list))
		for i = 0 to UBound(list)
			set users(i) = new UserObj
			users(i).setData(list(i))
		next

		dim alt
		alt = false

		'so now we are going to loop through the list of users and build a row for each of them
		for i = 0 to UBound(users)

			if alt then %>
            	<tr class="odd">
<%			else %>
            	<tr class="even">
<%			end if
			alt = not alt %>



			<td align="left"><% = users(i).name %></td>


			<% 'Build the tiebreaker display.
			tbGuessVis = "--"
			tbGuessHome = "--"
			tbDiff = "&nbsp;"
			if IsNumeric(users(i).tbGuessVis) and IsNumeric(users(i).tbGuessHome) then
				tbGuessVis = users(i).tbGuessVis
				tbGuessHome = users(i).tbGuessHome
				if IsNumeric(tbActualHome) and IsNumeric(tbActualVis) then
					'tbDiff = tbActualHome & users(i).tbGuessHome & tbActualVis & users(i).tbGuessVis
					tbDiff = Abs(tbActualHome - users(i).tbGuessHome) + Abs(tbActualVis - users(i).tbGuessVis)
					'call ErrorMessage(tbActualHome & " " & users(i).tbGuessHome & " " & tbActualVis & " " & users(i).tbGuessVis & " " & tbdiff & " " & totalGames)
				end if
			end if

			'Determine if the pick should be hidden.
			hidePick = false
			if HIDE_OTHERS_PICKS and not allLocked and not IsAdmin() and currentUser <> users(i).name then
					hidePick = true
			end if

			if hidePick then
				tbGuessVis = "XX"
				tbGuessHome = "XX"
			end if

			'Build the correct pick and score displays.
			pickScore  = "n/a"
			pickPct    = "n/a"
			if completedGames > 0 then
				pickScore = users(i).pickScore
				pickPct = cstr(Round((pickScore / totalGames) * 100,2)) & "%"
				'pickScore = pickScore & "/" & completedGames
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

				<td><% = pickScore %></td>
                <td><% = totalGames %></td>
				<td><% = pickPct %></td>
				<td><% = tbDiff %></td>
				<td>(<% = tbGuessVis %>/<% = tbGuessHome %>)</td>














<%			'Display picks for this user, highlighting the correct ones.
			sql = "SELECT * FROM Picks, Schedule" _
			   & " WHERE Username = '" & SqlString(users(i).name) & "'" _
			   & " AND Week = " & week _
			   & " AND Picks.GameID = Schedule.GameID" _
			   & " ORDER BY Schedule.Date, Schedule.Time, Schedule.VisitorID"
			set rs = DbConn.Execute(sql)
			do while not rs.EOF
				pick = rs.Fields("Pick").Value
				conf = rs.Fields("Confidence").Value

				'Determine if the pick should be hidden.
				hidePick = false
				if HIDE_OTHERS_PICKS and not allLocked and not IsAdmin() and currentUser <> users(i).name then
					gameDatetime = CDate(rs.Fields("Date").Value & " " & rs.Fields("Time").Value)
					if not GameStarted(gameDatetime) then
						hidePick = true
					end if
				end if

				'Determine the correct pick.
				if USE_POINT_SPREADS then
					correctPick = rs.Fields("ATSResult").Value
				else
					correctPick = rs.Fields("Result").Value
				end if

				'If tie picks are not allowed and the game result was a tie,
				'give the player the pick.
				giveTie = false
				if not ALLOW_TIE_PICKS and correctPick = TIE_STR and pick <> "" then
					giveTie = true
				end if

				'Format the pick display.
				if pick = "" then
					pick = "---"
					conf = "--"
				elseif hidePick then
					pick = "XXX"
					conf = "xx"
				else
					if pick = correctPick or giveTie then
						if giveTie then
							pick = pick & HALF_POINTS
						else
							pick = FormatCorrectPick(pick)
						end if
					end if
				end if

				'When using point spreads, highlight the straight-up result.
				'if it matches the pick.
				if USE_POINT_SPREADS then
					if rs.Fields("Pick").Value = rs.Fields("Result").Value then
						pick = FormatWinner(pick)
					end if
				end if


				if USE_CONFIDENCE_POINTS then %>
					<!--			<td class="confPick"><% = pick %><br /><% = conf %></td>-->
<%				else %>
					<td><% = pick %></td>
<%				end if
				rs.MoveNext
			loop%>





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


    </div> <!-- .x12 -->





        <p>&nbsp;      </p>
    <p>&nbsp;</p>
    <p>
      <% 	'List links to view other weeks.
	%><p>
    </p>
    	</div> <!-- #content -->
   <!--</p></td></tr></table>-->
</div>




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
		public tbGuessVis, tbGuessHome

		private sub Class_Initialize()
		end sub

		private sub Class_Terminate()
		end sub

		public sub setData(username)

			'Set the user properties.
			name      = username
			pickScore = PlayerPickScore(username, week)
			confScore = PlayerConfidenceScore(username, week)
			tbGuessVis   = TBvisGuess(username, week)
			tbGuessHome   = TBhomeGuess(username, week)

		end sub

	end class %>

