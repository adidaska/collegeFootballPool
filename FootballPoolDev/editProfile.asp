<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Edit Profile" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
    <link href="styles/style.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/email.asp" -->
<!-- #include file="includes/encryption.asp" -->
<!-- #include file="includes/form.asp" -->
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
    <!-- start of the admin table -->
		<div class="admin-entry-form">
	        <form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
	            <div class="header bottomEdge">
	                <div class="admin-entry-form"><span>Administrator Access</span></div>
	            </div>
	            <div>You may view or edit any user's profile (including your own) by selecting a username below.</div>
	            <div align="center">
	               <br/>
	              <strong>Select user:</strong>
	               
	              <select name="username">
	                   <option value=""></option>
	                   <% users = UsersList(true)
	                   if IsArray(users) then
	                      for i = 0 to UBound(users) %>
	                      <option value="<% = users(i) %>" <% if users(i) = username then Response.Write(" selected=""selected""") end if %>><% = users(i) %></option>
	                      <%	next
	                   end if %>
	              </select>						
	               <input type="submit" name="submit" value="Select" class="button" title="View/edit the selected user's profile." />  
	          </div>		
	        </form>
        </div>
        
        
        <% end if
        

	'If and update was requested, process it.
	dim email, contact, news, schedule, spread, result, disable, phone, userDisplayName, firstName, lastName
	dim sql, rs
	if Request.Form("submit") = "Update" and username <> "" then

		'Get the form fields.
		email    = Trim(Request.Form("email"))
		contact  = Trim(Request.Form("contact"))
		news     = Trim(Request.Form("news"))
		schedule = Trim(Request.Form("schedule"))
		spread   = Trim(Request.Form("spread"))
		result   = Trim(Request.Form("result"))
		disable  = Trim(Request.Form("disable"))
		phone  = Trim(Request.Form("phone"))
		userDisplayName  = Trim(Request.Form("userDisplayName"))
		if LCase(news)     <> "true" then news     = false end if
		if LCase(schedule) <> "true" then schedule = false end if
		if LCase(spread)   <> "true" then spread   = false end if
		if LCase(result)   <> "true" then result   = false end if
		if LCase(disable)  <> "true" then disable  = false end if

		'Valid the email address, if provided.
		if email <> "" and not IsValidEmailAddress(email) then
			FormFieldErrors.Add "email", "'" & email & "' does not appear to be a valid email address."
		end if

		'If notifications are requested, make sure and email address is specified.
		if (news or schedule or spread or result) and email = "" then
			FormFieldErrors.Add "email", "An email address is required to receive notifications."
		end if

		'If there were no errors, do the updates.
		if FormFieldErrors.Count = 0 then
			sql = "UPDATE Users SET" _
			   & " EmailAddress = '" & Encrypt(email) & "'," _
			   & " ContactInformation = '" & Encrypt(contact) & "'," _
			   & " PhoneNumber = '" & Encrypt(phone) & "'," _
			   & " UserDisplayName = '" & userDisplayName & "'," _
			   & " NotifyOfNewsUpdates = " & news & "," _
			   & " NotifyOfScheduleUpdates = " & schedule & "," _
			   & " NotifyOfSpreadUpdates = " & spread & ", " _
			   & " NotifyOfResultUpdates = " & result
			if IsAdmin() then
				sql = sql & ", DisableEntries = " & disable
			end if
			sql = sql & " WHERE Username = '" & SqlString(username) & "'"
			call DbConn.Execute(sql)
			call InfoMessage("Update successful.")

		'Otherwise, show the form field errors.
		else
			call FormFieldErrorsMessage("Error: Invalid fields. Please correct and resubmit.")
		end if
	end if

	'Build the form display.
	if username <> "" then
		dim newsCheckedStr, scheduleCheckedStr, spreadCheckedStr, resultCheckedStr, disableCheckedStr
		sql = "SELECT * FROM Users WHERE Username = '" & SqlString(username) & "'"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			email    = GetFieldValue("email",    Decrypt(rs.Fields("EmailAddress").Value))
			contact  = GetFieldValue("contact",  Decrypt(rs.Fields("ContactInformation").Value))
			news     = GetFieldValue("news",     rs.Fields("NotifyOfNewsUpdates").Value)
			schedule = GetFieldValue("schedule", rs.Fields("NotifyOfScheduleUpdates").Value)
			spread   = GetFieldValue("spread",   rs.Fields("NotifyOfSpreadUpdates").Value)
			result   = GetFieldValue("result",   rs.Fields("NotifyOfResultUpdates").Value)
			disable  = GetFieldValue("disable",  rs.Fields("DisableEntries").Value)
			phone    = GetFieldValue("phone",  rs.Fields("PhoneNumber").Value)
			userDisplayName   = GetFieldValue("userDisplayName",  rs.Fields("UserDisplayName").Value)
			firstName         = GetFieldValue("disable",  rs.Fields("First Name").Value)
			lastName          = GetFieldValue("lastName",  rs.Fields("Last Name").Value)

			'Set the checkbox states.
			newsCheckedStr = "" : scheduleCheckedStr = "" : spreadCheckedStr = "" : resultCheckedStr = ""
			if news     then newsCheckedStr     = CHECKED_ATTRIBUTE end if
			if schedule then scheduleCheckedStr = CHECKED_ATTRIBUTE end if
			if spread   then spreadCheckedStr   = CHECKED_ATTRIBUTE end if
			if result   then resultCheckedStr   = CHECKED_ATTRIBUTE end if
			if disable  then disableCheckedStr  = CHECKED_ATTRIBUTE end if

			'If there were errors on the form post processing, restore the checkboxes using the posted information.
			if FormFieldErrors.Count > 0 then

				'If the user unchecked a checkbox, be sure to leave it unchecked.
				if not FormFieldExists("news")     then newsCheckedStr     = "" end if
				if not FormFieldExists("schedule") then scheduleCheckedStr = "" end if
				if not FormFieldExists("spread")   then spreadCheckedStr   = "" end if
				if not FormFieldExists("result")   then resultCheckedStr   = "" end if
				if not FormFieldExists("disable")  then disableCheckedStr  = "" end if

			end if
		end if %>
        
        
        
    <!-- start of the games table -->


		<div class="clearfix" id="content-wrap">
		  	<div id="content-top"></div>
		    <div id="primary" class="hfeed">
    		<%	if username <> "" then %>
				<h2>Profile for <% = username %></h2>
			<%	end if %>
    
    

    
        
			<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
			<% if IsAdmin() then %>
				<div><input type="hidden" name="username" value="<% = username %>" /></div>
			<% end if %>
	  
	        
			<table class="main fixed" cellpadding="0" cellspacing="0">
				<tr class="header bottomEdge">
					<th align="left">Edit Your Profile</th>
				</tr>
				<tr>
					<td class="freeForm">
						<% dim str
						str = "Use this form to change your personal information"
						if SERVER_EMAIL_ENABLED then
							str = str & " and to enable or disable email notifications"
						end if
						
						str = str & "." %>
						<p><% = str %>
						This information is for use by the Administrator and <em>will not</em> be shared with other players.</p>   
						<%	if SERVER_EMAIL_ENABLED then %>
							<p><em>If you choose to receive notifications, please be sure your email address is correct.</em></p>
						<%	end if %>
			        
			        	<table cellpadding="0" cellspacing="0">
			                <tr>
			                    <td><strong>Username:</strong></td>
			                    <td><% = username %></td>
			                </tr>
			                <tr>
                                <td><strong>Display Name:</strong></td>
                                <td><input type="text" name="userDisplayName" value="<% = Server.HtmlEncode(userDisplayName) %>" size="30" class="<% = FieldStyleClass("", "userDisplayName") %>" style="width: 18em;" /></td>
                            </tr>
			                <tr>
			                    <td><strong>Email address:</strong></td>
			                    <td><input type="text" name="email" value="<% = Server.HtmlEncode(email) %>" size="30" class="<% = FieldStyleClass("", "email") %>" style="width: 18em;" /></td>
			                </tr>
			                <tr valign="top">
			                    <td><strong>Contact information:</strong></td>
			                    <td><textarea name="contact" rows="4" cols="30" style="width: 18em;"><% = Server.HtmlEncode(contact) %></textarea>
			                    </td>
			                </tr>
			                
							<% if SERVER_EMAIL_ENABLED then %>
			                    <tr><td><strong>Email notifications:</strong></td>
			                        <td><input type="checkbox" id="news" name="news" value="true"<% = newsCheckedStr %> /> <label for="news">News announcements.</label></td>
			                    </tr>
			                    <tr>
			                        <td>&nbsp;</td>
			                        <td><input type="checkbox" id="schedule" name="schedule" value="true"<% = scheduleCheckedStr %> /> <label for="schedule">Scheduling changes.</label></td>
			                    </tr>
			                    <%	if USE_POINT_SPREADS then %>
			                        <tr>
			                            <td>&nbsp;</td>
			                            <td><input type="checkbox" id="spread" name="spread" value="true"<% = spreadCheckedStr %> /> <label for="spread">Point spread updates.</label></td>
			                        </tr>
			                    <% end if %>
			                    <tr><td>&nbsp;</td>
			                        <td>
			                        <input type="checkbox" id="result" name="result" value="true"<% = resultCheckedStr %> /> 
			                        <label for="result">Game results posted.</label>
			                        </td>
			                    </tr>
			                <% end if
								
							if IsAdmin() and username <> ADMIN_USERNAME then
								dim str1, str2
								str1 = "pool entries"
								str2 = "entries"
								if ENABLE_MESSAGE_BOARD then
									str1 = str1 & " or message board posts"
									str2 = str2 & " or posts"
								else
								end if %>
				
			                <tr>
			                    <td colspan="2"><p>When disabled, a player may still log in but cannot make or change <% = str1 %>.
			                    (The player is permitted to delete any open <% = str2 %>, however).
			                    The player will receive a message to that effect on log in.</p></td>
			                </tr>
			                <tr valign="middle">
			                    <td><strong>Disable user:</strong></td>
			                    <td>
			                    <input type="checkbox" id="disable" name="disable" value="true"<% = disableCheckedStr %> /> 
			                    <label for="disable">Prevent player from making <% = str2 %>.</label></td>
			                </tr>
							<% end if %>
							</table>
	                    
						</td>
					</tr>
				</table>
				<p>
		        <input type="submit" name="submit" value="Update" class="button" title="Update your profile." />&nbsp;
		        <input type="submit" name="submit" value="Cancel" class="button" title="Cancel the change." />
		        </p>
			</form>
			
			</div>
        <div id="content-btm"></div>
	</div><%	end if %>
	</td></tr></table>   
<!-- #include file="includes/footer.asp" -->
</body>
</html>