<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" -->
<!-- #include file="includes/updateFunctions.asp" -->
<% PageSubTitle = "Update Game Schedule" : AdminOnly = true %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
<title>
<% = PAGE_TITLE & ": " & PageSubTitle %>
</title>
<!--<link rel="shortcut icon" href="favicon.ico" />-->
<link rel="stylesheet" type="text/css" href="styles/common.css" />
<link rel="stylesheet" type="text/css" href="styles/menu.css" />
<link rel="stylesheet" type="text/css" href="styles/datetimePicker.css" />
<link rel="stylesheet" type="text/css" href="styles/menu.css" />

<link href="styles/style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="scripts/common.js"></script>
<script type="text/javascript" src="scripts/menu.js"></script>
<script type="text/javascript" src="scripts/datetimePicker.js"></script>
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/datetimePicker.asp" -->
<!-- #include file="includes/email.asp" -->
<!-- #include file="includes/encryption.asp" -->
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/side.asp" -->
<!-- #include file="includes/weekly.asp" -->
<table width="450" id="wrapper">
  <tr>
    <td width="400" style="padding: 0px;">
	<%	'Open the database.
	call OpenDB()

	'Get the week to display.
	dim week
	week = GetRequestedWeekFS()
	
	'If there is form data, process it.
	dim n, i
	dim gameID, gameWeek, gameDate, gameTime, displayValue, homeTeam, homeTeamID, visTeam, visTeamID, viewTime, pointSpread, fullGameID, inPool
	dim logoVis, logoHome
	dim notify, infoMsg
	dim sql, rs
	dim affectedWeeks
	affectedWeeks = Array()
	n = NumberOfGamesFS(week) 'this gets the number of rows in the fullSchedule table for the given week

	if Request.ServerVariables("Content_Length") > 0 and not CancelRequested() then
			'this is getting all the fields and validating the entries to be inserted into the DB
			for i = 1 to n
				inPool = Trim(Request.Form("inPool-" & i))
				if LCase(inPool) = "true" then 
					gameID   = Trim(Request.Form("id-"   & i))
					gameWeek = Trim(Request.Form("week-" & i))
					gameDate = Trim(Request.Form("date-" & i))
					gameTime = Trim(Request.Form("time-" & i))
	
					'Validate the form fields.
					if not IsNumeric(gameWeek) then
						FormFieldErrors.Add "week-" & i, "'" & gameWeek & "' is \MZ week number."
					else	
						if CInt(gameWeek) < 1 then
							FormFieldErrors.Add "week-" & i, "'" & gameWeek & "' is not a valid week number."
						end if
					end if
					if not IsDate(gameDate) then
						FormFieldErrors.Add "date-" & i, "'" & gameDate & "' is not a valid date."
					end if
					if not IsDate(gameTime) then
						FormFieldErrors.Add "time-" & i, "'" & gameTime & "' is not a valid time, please update."
					end if
				end if
			next
		'If there were any errors, display the error summary message.
		'Otherwise, do the updates.
		'added tbgame into mix so can choose the tbgame from this screen
		if FormFieldErrors.Count > 0 then
			call FormFieldErrorsMessage("Error: Invalid fields. Please correct and resubmit.")
		else
			updateSchedule(week)
		end if
	end if
	'Display the schedule for the specified week. 
	%>
    
      <form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
        <div>
          <p>
          	<div class="daysHeader">
              <p>
                <%	'List links to view other weeks.
			
				call DisplayWeekNavigationFS(1, "") %>
                <input type="hidden" name="week" value="<% = week %>" />
              </p>
          	</div>
            <div>
            	<h2>Select Games for Week <% = week %></h2>
            </div>
          </p>
          <p align="center" class="error">Caution: Updating from this page will remove any current voting for this week</p>
        </div>
        
 
        
        
        <table width="625" cellpadding="0" cellspacing="0" class="main">
          <tr class="header bottomEdge">
            <th width="46" align="left">Week</th>
            <th width="38">&nbsp;</th>
            <th align="left" colspan="2">Date</th>
            <th align="left" colspan="2">Time</th>
            <th width="56" align="right"><div align="center">Point Spread</div></th>
            <th width="270"><div align="center">Teams</div></th>
            <th width="54">In
               
            Pool</th>
          </tr>
    

	
	
	
	<%      
	dim visitor, home
	dim alt
	set rs = WeeklyFullSchedule(week)
	if not (rs.BOF and rs.EOF) then
		n = 1
		alt = false
		do while not rs.EOF
			fullGameID  	= rs.Fields("ID").Value
			gameWeek 		= rs.Fields("Week").Value
			visTeam  		= rs.Fields("VisTeam").Value
			homeTeam  		= rs.Fields("HomeTeam").Value
			gameDate 		= rs.Fields("GameDate").Value
			gameTime 		= rs.Fields("GameTime").Value
			displayValue  	= rs.Fields("DisplayValue").Value
			inPool 			= rs.Fields("InPool").Value
			'if rs.Fields("InPool").Value = "No" then
			'	inPool = "unchecked"
			'elseif rs.Fields("InPool").Value = "Yes" then
			'	inPool = "checked"
			'end if
			pointSpread  	= rs.Fields("PointSpread").Value
			gameID  		= rs.Fields("GameId").Value
			visTeamID  		= rs.Fields("visTeamID").Value
			logoVis  		= rs.Fields("logoVis").Value
			homeTeamID  	= rs.Fields("homeTeamID").Value
			logoHome  		= rs.Fields("logoHome").Value

			if alt then %>
          <tr class="alt">
            <%			else %>
          <tr>
            <%			end if
			alt = not alt

			'If there were errors on the form post processing, restore those fields.
			if FormFieldErrors.Count > 0 then
				gameWeek = GetFieldValue("week-" & n, gameWeek)
				gameDate = GetFieldValue("date-" & n, gameDate)
				gameTime = GetFieldValue("time-" & n, gameTime)
			end if 
			
			if IsNull(gameTime) or (gameTime = "TBA") then 
				viewTime = "TBA"
			else 
				viewTime = FormatFullTime(gameTime) 
			end if 
			%>
            <td><input type="hidden" name="id-<% = n %>" value="<% = fullGameID %>" />
            <input type="hidden" name="homeID-<% = n %>" value="<% = homeTeamID %>" />
            <input type="hidden" name="visID-<% = n %>" value="<% = visTeamID %>" />

              <input type="text" name="week-<% = n %>" value="<% = gameWeek %>" size="2" class="<% = FieldStyleClass("numeric", "week-" & n) %>" /></td>
            <td align="right"><input type="text" name="day-<% = n %>" value="<% = WeekdayName(Weekday(gameDate), true) %>" size="3" class="readonly" readonly="readonly" /></td>
            <td width="90"><input type="text" name="date-<% = n %>" value="<% = gameDate %>" size="10" class="<% = FieldStyleClass("numeric readonly", "date-" & n) %>" readonly="readonly" /></td>
            <td width="26">
            <input type="image" src="graphics/calendar.png" onclick="return openCalendar(this.form, 'day-<% = n %>', 'date-<% = n %>');" title="Select a new date." class="table-image"/>
            </td>
            <td width="90"><input type="text" name="time-<% = n %>" value="<% = viewTime %>" size="10" class="<% = FieldStyleClass("numeric readonly", "time-" & n) %>" readonly="readonly" /></td>
            <td width="26"><input type="image" src="graphics/clock.png" onclick="return openClock(this.form, 'time-<% = n %>');" title="Select a new time." class="table-image"/></td>
            <td align="left"><div align="center"><input type="text" name="spread-<% = n %>" value="<% = pointSpread %>" size="4" class="<% = FieldStyleClass("numeric", "spread-" & n) %>" />
            </div>            </td>
            <td><label name="visTeam-<% = n %>"><% = visTeam %></label> at <label name="homeTeam-<% = n %>"><% = homeTeam %></label></td>
            <td><div align="center"><input name="inPool-<% = n %>" type="checkbox" class="inpool" id="inpoolid-<% = n %>" value="true" title="inpool checkbox" <%if rs("inPool")="Yes" then Response.Write("checked")%>/>
              
          </div></td>
          </tr>
          <%			rs.MoveNext
			n = n + 1
		loop
		if SERVER_EMAIL_ENABLED then %>
          <tr class="subHeader topEdge">
            <th align="left" colspan="9"><input type="checkbox" id="notify" name="notify" value="true" />
              <label for="notify">Send update notification to users.</label></th>
          </tr>
          <% end if %>
        </table>
        

		
		
		
		<%		'List open dates.
		'call DisplayOpenDates(2, week)
	end if %>
        <p>
          <input type="submit" name="submit" value="Update" class="button" title="Apply changes." />
          &nbsp;
          <input type="submit" name="submit" value="Cancel" class="button" title="Cancel the update." />
        </p>
      </form>      </td>
  </tr>
  <tr>		
    <td style="padding: 0px;">&nbsp;</td>
  </tr>
</table>

<!-- #include file="includes/footer.asp" -->
</body>
</html>
<%	'**************************************************************************
	'* Local functions and subroutines.                                       *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Adds the specified week to the list of affected weeks, if it is not
	' already present.
	'--------------------------------------------------------------------------
	sub AddAffectedWeek(week)

		dim i

		for i = LBound(affectedWeeks) to UBound(affectedWeeks)
			if affectedWeeks(i) = week then
				exit sub
			end if
		next
		redim preserve affectedWeeks(Ubound(affectedWeeks) + 1)
		affectedWeeks(UBound(affectedWeeks)) = week

	end sub

	'--------------------------------------------------------------------------
	' Sorts the list of affected weeks.
	'--------------------------------------------------------------------------
	sub SortAffectedWeeks()

		dim i, j, tmp

		for i = LBound(affectedWeeks) to UBound(affectedWeeks) - 1
			for j = i + 1 to UBound(affectedWeeks)
				if affectedWeeks(j) < affectedWeeks(i) then
					tmp = affectedWeeks(i)
					affectedWeeks(i) = affectedWeeks(j)
					affectedWeeks(j) = tmp
				end if
			next
		next

	end sub

	'--------------------------------------------------------------------------
	' Sends an email notification to any users who have elected to receive
	' them.
	'--------------------------------------------------------------------------
	sub SendNotifications()

		dim subj, body
		dim i, rs
		dim list, email

		subj = "Football Pool Game Schedule Update Notification"
		body = ""

		'Show the schedule for each affected week.
		for i = LBound(affectedWeeks) to UBound(affectedWeeks)
			if i > 0 then
				body = body & vbCrLf
			end if
			body = body & "The game schedule for Week " & affectedWeeks(i) & " has been updated." & vbCrLf & vbCrLf
			set rs = WeeklySchedule(affectedWeeks(i))
			do while not rs.EOF
				body = body  _
				     & WeekdayName(Weekday(rs.Fields("Date").Value), true) _
			    	 & " " & FormatDate(rs.Fields("Date").Value) _
					 & " " & FormatTime(rs.Fields("Time").Value) _
				     & " " & rs.Fields("VisitorID").Value & " @ " & rs.Fields("HomeID").Value & vbCrLf
				rs.moveNext
			loop
		next
		body = body & vbCrLf & "(All times Eastern)" & vbCrLf

		list = GetNotificationList("NotifyOfScheduleUpdates")
		for each email in list
			call SendMail(email, subj, body)
		next

	end sub %>

'----------------------------------------------------------------------------------------------------

