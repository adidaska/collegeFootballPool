<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" -->
<% PageSubTitle = "User Login" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/passwords.asp" -->
	
    <div class="clearfix" id="content-wrap">
  		<div id="content-top"></div>
    	<div id="primary" class="hfeed">
    <div id="wrapper">
<%	'Open the database.
	call OpenDB()

	'If there is form data, process it.
	dim username, password
	dim sql, rs
	if Request.ServerVariables("Content_Length") > 0 then
		username = Trim(Request.Form("username"))
		password = Trim(Request.Form("password"))
		if username = "" then
	        FormFieldErrors.Add "username", "Please select your user name and enter your password."
	        FormFieldErrors.Add "password", ""
		elseif password = "" then
			FormFieldErrors.Add "password", "Please enter your password."
		end if

		'If the data is good, check the username and password.
		if FormFieldErrors.Count = 0 then
			sql = "SELECT * FROM Users WHERE Username = '" & SqlString(username) & "'"
			set rs = DbConn.Execute(sql)
			if rs.EOF and rs.BOF then
		        FormFieldErrors.Add "username", "Username '" & username & "' not found."
			elseif Hash(rs.Fields("Salt").Value & password) <> rs.Fields("Password").Value then
		        FormFieldErrors.Add "password", "Password is incorrect."
			else
		   		'Save the Username in the Session and redirect to the home page.
				Session(SESSION_USERNAME_KEY) = rs.Fields("Username").Value
				Session.Timeout = SESSION_TIMEOUT_LENGTH
				Response.Redirect("./")
			end if
		end if

		'Show error messages.
		call FormFieldErrorsMessage("Login failed, see below:")

	end if %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<div class="user-login-form" cellpadding="0" cellspacing="0">
			<div class="header bottomEdge">
				<p align="center" class="game-Title">User Login</p>
			</div>
                <p>To login, select your name from the list, enter your password and hit the Login button.</p>
                 <div class="user-password-box">
                        Username:&nbsp;&nbsp;

                            <input type="text" name="username" value="" class="<% = FieldStyleClass("", "username") %>" />
                            <br />
                        Password:&nbsp;&nbsp;
                        <input type="password" name="password" value="" class="<% = FieldStyleClass("", "password") %>" />

					</div>
					<p>&nbsp;</p>
					<p><em>Passwords are case-sensitive.</em>
					If you have forgotten your password or cannot log in for whatever reason, contact the <a href="mailto:<% = ADMIN_EMAIL %>">Administrator</a> for help.</p>

		<p><input type="submit" name="submit" value="Login" class="button" title="Login to your account." /></p>
        </div>
	</form>
	</div>
    
           </div> <!-- end of the primary div in the container--> 
		<div id="content-btm"></div>
	</div>
     
    
    
    
<!-- #include file="includes/footer.asp" -->
</body>
</html>