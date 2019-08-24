<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Home Page" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
	<link href="styles/style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/news.asp" -->


<div class="clearfix" id="content-wrap">
  	<div id="content-top"></div>
    <div id="primary" class="hfeed">
    
    

<%	'Open the database.
	call OpenDB() %>
	
<%	'If the current user is disabled, display a message.
	if IsDisabled() then
		call ErrorMessage("Warning: Your account has been disabled. Please contact the Administrator.")
	end if %>
    
    <div class="left_aligned_style">
    	<div class="game-Title">
			<span>Welcome to the NCAA Football Pool</span>
            <br />
            <br />
		</div>
    	<div>
          <span class="game-Title">
          <% = MainHeader() %>
          </span>			
          <br />	
          <!-- News. -->
				<%	'Display any news.
				call DisplayNews(4, GetNews()) %>
				<!-- End news. -->
				<%	'List any early games.
                    dim found, dateNow, sql, rs, week, dayName
                    found = false
                    dateNow = CurrentDateTime()
                    sql = "SELECT * FROM Schedule" _
                       & " WHERE Date + Time > #" & DateValue(dateNow) & " " & TimeValue(dateNow) & "#" _
                       & " ORDER BY Date, Time"
                    set rs = DbConn.Execute(sql)
                    if not (rs.BOF and rs.EOF) then
                        week = 0
                        do while not rs.EOF
                            dayName = Weekday(rs.Fields("Date").Value)
                            if week <> rs.Fields("Week").Value and _
                               dayName <> vbSunday and _
                               dayName <> vbMonday then
                                week = rs.Fields("Week").Value
                                if not found then
                                    found = true %>
                                    <br />
                                <!--<span class="game-Title">Early Games</span>-->
                                <br /><br />
                                <span>There are early games on the following dates, be sure to make your picks for these games on time:</span>
                                <ul>
                				<%	end if %>
                                <br />
                                <li><a href="entryForm.asp?week=<% = rs.Fields("Week").Value %>">Week <% = rs.Fields("Week").Value %></a> - <% = FormatDateTime(rs.Fields("Date").Value & " " & rs.Fields("Time").Value, vbLongDate) %></li>
                <%			end if
                            rs.MoveNext
                        loop
                        if found then %>
                                </ul>
                	<%	end if
                    end if %>
				<p>See the <a href="help.asp">Help</a> section for rules and instructions.</p>
	  </div>
    
    </div>
    
    
	
    
    
  </div> <!-- end of the primary div in the container-->

    
    
    
<div id="content-btm"></div>
</div>


<!-- #include file="includes/footer.asp" -->
</body>
</html>
<%	'**************************************************************************
	'* Local functions and subroutines.                                       *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Returns an appropriate header based on the current date.
	'--------------------------------------------------------------------------
	function MainHeader()

		dim sql, rs, lastDay, currentDay

		MainHeader = ""
		currentDay = CurrentDateTime()

		'Are we in the regular season?
		sql = "SELECT MAX(Date) AS LastDay FROM Schedule"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			lastDay = rs.Fields("LastDay").Value
			if currentDay <= DateAdd("d", 2, lastDay) then
				MainHeader = "This is Week " & CurrentWeek()
				exit function
			end if
		end if

		'Are we in the postseason?
		if ENABLE_PLAYOFFS_POOL then
			sql = "SELECT MAX(Date) AS LastDay FROM PlayoffsSchedule"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				lastDay = rs.Fields("LastDay").Value
				if currentDay <= DateAdd("d", 2, lastDay) then
					MainHeader = "It's Playoffs Time"
					exit function
				end if
			end if
		end if

		'The season must be over.
		MainHeader = "The Season is Over"

	end function %>

