	<!-- Site menu. -->
	<div id="menuBar">
		<table>
			<tr>
				<td><a tabindex="1"
					href="./"
					onfocus="closeMenu();"
					title="News, announcements and updates.">Home</a></td>
<%	if ENABLE_MESSAGE_BOARD then %>
				<td><a tabindex="2"
					href="messageBoard.asp"
					onfocus="closeMenu();"
					title="View, post or edit messages.">Message Board</a></td>
<%	end if %>
				<td><a tabindex="3"
					class="hasMenu"
					href="#"
					onclick="return false;"
					onfocus="openMenu(this, 'poolMenu')"
					onmouseover="openMenu(this, 'poolMenu')">Weekly Pool <span class="arrow">&#9660;</span></a></td>
<%	if ENABLE_SURVIVOR_POOL or ENABLE_MARGIN_POOL then %>
				<td><a tabindex="9"
					class="hasMenu"
					href="#"
					onclick="return false;"
					onfocus="openMenu(this, 'sideMenu')"
					onmouseover="openMenu(this, 'sideMenu')"><% = SidePoolTitle %> Pool <span class="arrow">&#9660;</span></a></td>
<%	end if %>
				<td><a tabindex="12"
					class="hasMenu"
					href="#"
					onclick="return false;"
					onfocus="openMenu(this, 'scheduleMenu')"
					onmouseover="openMenu(this, 'scheduleMenu')">Schedules <span class="arrow">&#9660;</span></a></td>
				<td><a tabindex="17"
					class="hasMenu"
					href="#"
					onclick="return false;"
					onfocus="openMenu(this, 'accountMenu')"
					onmouseover="openMenu(this, 'accountMenu')">My Account <span class="arrow">&#9660;</span></a></td>
<%	if IsAdmin() then %>
				<td><a tabindex="23"
					class="hasMenu"
					href="#"
					onclick="return false;"
					onfocus="openMenu(this, 'adminMenu')"
					onmouseover="openMenu(this, 'adminMenu')">Administration <span class="arrow">&#9660;</span></a></td>
<%	end if %>
				<td><a tabindex="34"
					href="help.asp"
					onfocus="closeMenu();"
					title="Pool rules and helpful information.">Help</a></td>
			</tr>
		</table>
	</div>
	<!-- Sub menus. -->
	<div class="menu" id="poolMenu"<% if not ENABLE_PLAYOFFS_POOL then Response.Write(" style=""width: 9em;""") end if %>>
		<a tabindex="4" href="entryForm.asp" title="Enter your picks for this week or upcoming weeks.">Entry Form</a>
		<a tabindex="5" href="poolResults.asp" title="View the current status or outcome of each week's pool.">Results</a>
		<a tabindex="6" href="poolSummary.asp" title="View the list of weekly pool winners and individual statistics.">Summary</a>
<%	if ENABLE_PLAYOFFS_POOL then %>
		<div class="separator"></div>
		<a tabindex="7" href="playoffsEntryForm.asp" title="Enter your picks for the playoffs pool.">Playoffs Entry Form</a>
		<a tabindex="8" href="playoffsResults.asp" title="View the current status or outcome of the playoffs pool.">Playoffs Results</a>
<%	end if %>
	</div>
<%	if ENABLE_SURVIVOR_POOL or ENABLE_MARGIN_POOL then %>
	<div class="menu" id="sideMenu" style="width: <% if ENABLE_SURVIVOR_POOL and ENABLE_MARGIN_POOL then Response.Write("12em;") else Response.Write("9em;") end if %>">
		<a tabindex="10" href="sideEntryForm.asp" title="Enter your picks for the <% = SidePoolTitle %> pool.">Entry Form</a>
            <!--<a tabindex="11" href="sideStandings.asp" title="View the current <% = SidePoolTitle %> pool standings.">Standings</a> -->
	</div>
<%	end if %>
	<div class="menu" id="scheduleMenu">
		<a tabindex="13" href="teamSchedules.asp" title="Team by team game schedule.">Team Schedules</a>
		<a tabindex="14" href="weeklySchedule.asp" title="Week by week game schedule.">Weekly Schedule</a>
<%	if ENABLE_PLAYOFFS_POOL then %>
		<a tabindex="15" href="playoffsSchedule.asp" title="Playoffs game schedule.">Playoffs Schedule</a>
<%	end if %>
		<div class="separator"></div>
		<!--<a tabindex="16" href="standings.asp" title="League standings.">Standings</a> -->
	</div>
	<div class="menu" id="accountMenu">
		<!--<a tabindex="18" href="accountHistory.asp" title="View your account history.">Account History</a> -->
		<!--<a tabindex="19" href="poolWinnings.asp" title="View a summary of your winnings.">Pool Winnings</a> -->
		<div class="separator"></div>
		<a tabindex="20" href="editProfile.asp" title="Edit your profile.">Edit Profile</a>
		<a tabindex="21" href="changePassword.asp" title="Change your password.">Change Password</a>
		<div class="separator"></div>
		<a tabindex="22" href="userLogout.asp" title="Log out or log in as a different user.">Login/Logout</a>
	</div>
<%	if IsAdmin() then %>
	<div class="menu" id="adminMenu">
    	<a tabindex="24" href="setSchedule.asp" title="Pick games for the pool.">Pick Games in Pool</a>
		<a tabindex="24" href="updateScores.asp" title="Enter or change game results.">Enter Game Scores</a>
<%		if USE_POINT_SPREADS then %>
		<a tabindex="25" href="updateSpreads.asp" title="Enter or change point spreads.">Set Point Spreads</a>
<%		end if %>
		<a tabindex="26" href="updateSchedule.asp" title="Update the game schedule.">Update Game Schedule</a>
<%		if ENABLE_PLAYOFFS_POOL then %>
		<div class="separator"></div>
		<a tabindex="27" href="playoffsUpdateScores.asp" title="Enter or change playoffs game results.">Enter Playoffs Scores</a>
<%			if USE_POINT_SPREADS then %>
		<a tabindex="28" href="playoffsUpdateSpreads.asp" title="Enter or change playoffs point spreads.">Set Playoffs Point Spreads</a>
<%			end if %>
		<a tabindex="29" href="playoffsUpdateSchedule.asp" title="Update the playoffs game schedule.">Update Playoffs Schedule</a>
<%		end if %>
		<div class="separator"></div>
		<a tabindex="30" href="accountBalances.asp" title="View the current balance for all user accounts.">Account Balances</a>
		<a tabindex="31" href="manageAccounts.asp" title="Credit/debit user accounts.">Manage Accounts</a>
		<a tabindex="32" href="manageUsers.asp" title="Add or delete users or reset passwords.">Manage Users</a>
		<div class="separator"></div>
		<a tabindex="33" href="editNews.asp" title="Edit the news section of the home page.">Edit News</a>
	</div>
<%	end if %>
