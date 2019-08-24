	<!-- Page header. -->
    <div id="titleBar">
        <div id="titleLeft"><% = PAGE_TITLE %></div>
        <div id="titleRight">
        	<%	if Session(SESSION_USERNAME_KEY) <> "" then %>
				<div id="status">Welcome <% = Session(SESSION_USERNAME_KEY) %>
			<%	else %>
				<div id="status" class="error">You are not logged in
			<%	end if
			dim et
			et = CurrentDateTime()
			
			Dim str_time
			str_time=DateAdd("h",+3,Now) %>
						- Eastern Time: <span id="easternTime"><% = Replace(FormatDateTime(et, vbLongDate), " 0", " ") & " " & FormatFullTime(str_time) %></span>
		    	</div>
         </div>
         <div id="login">
			<%	if Session(SESSION_USERNAME_KEY) = "" then %>
							<a href="userLogin.asp">Log In</a>
			<%	else %>
							<a href="userLogout.asp">Log Out</a>
			<%	end if %>
		</div>
	</div>    
