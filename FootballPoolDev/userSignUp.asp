<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "User Sign Up" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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
act<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/email.asp" -->
<!-- #include file="includes/encryption.asp" -->
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/passwords.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'If there is form data, process it.
	dim code, username, email, password1, password2, badChar
	dim salt, password
	dim sql, rs
	if Request.ServerVariables("Content_Length") > 0 and not CancelRequested() then
		code     = Trim(Request.Form("code"))
		username = Trim(Request.Form("username"))
		email    = Trim(Request.Form("email"))
		badChar  = UsernameCheck(username)

		'Check the invitaion code field, if used.
		if SIGN_UP_INVITATION_CODE <> "" then
			if code = "" then
				FormFieldErrors.Add "code", "The invitation code is required."
			elseif code <> SIGN_UP_INVITATION_CODE then
				FormFieldErrors.Add "code", "The invitation code is incorrect."
			end if
		end if

		'Check the username field.
		if username = "" then
			FormFieldErrors.Add "username", "A username is required."
		elseif badChar <> "" then
			FormFieldErrors.Add "username", "Usernames may not contain a '" & Server.HtmlEncode(badChar) & "' character."
		elseif Len(username) > USERNAME_MAX_LEN then
			FormFieldErrors.Add "username", "Usernames may not exceed " & USERNAME_MAX_LEN & " characters."
		else
			sql = "SELECT * FROM Users WHERE Username = '" & SqlString(username) & "'"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				FormFieldErrors.Add "username", "Username '" & username & "' is already taken, try another."
			end if
		end if

		'Check the email field.
		if email = "" then
			FormFieldErrors.Add "email", "An email address is required."
		elseif not IsValidEmailAddress(email) then
			FormFieldErrors.Add "email", "'" & email & "' does not appear to be a valid email address."
		end if

		'Handle the case where server-side email is not available.
		if not SERVER_EMAIL_ENABLED then

			'Check the password fields.
			password1 = Trim(Request.Form("password1"))
			password2 = Trim(Request.Form("password2"))
			if password1 = "" or password2 = "" then
				FormFieldErrors.Add "password1", "A password is required."
				FormFieldErrors.Add "password2", "You must enter your password twice, where indicated."
			elseif password1 <> password2 then
				FormFieldErrors.Add "password1", "Passwords fields do not match."
				FormFieldErrors.Add "password2", ""
			elseif Len(password1) < PASSWORD_MIN_LEN then
				FormFieldErrors.Add "password1", "Password must be at least " & PASSWORD_MIN_LEN & " characters long."
			end if
		end if

		'If the input data is good, add the user.
		if FormFieldErrors.Count = 0 then
			if SERVER_EMAIL_ENABLED then
				password = CreatePassword()
			else
				password = password1
			end if
			salt = CreateSalt()
			sql = "INSERT INTO Users" _
			   & " (Username, EmailAddress, Salt, [Password])" _
			   & " VALUES('" & SqlString(username) & "'," _
			   & " '" & Encrypt(email) & "'," _
			   & " '" & salt & "'," _
			   & " '" & Hash(salt & password) & "')"
			call DbConn.Execute(sql)

			'If enabled, send the emails.
			if SERVER_EMAIL_ENABLED then

				'Send emails, one to the user and one to the administrator.
				dim mailSubj, mailMsg, errMsg
				mailSubj = "Football Pool Sign Up"
				mailMsg = "Thanks for signing up. You may login at " & POOL_URL & " with the username and password below." & vbCrLf _
				       & vbCrLf _
				       & "Username: " & username & vbCrLf _
				       & "Password: " & password & vbCrLf _
			    	   & vbCrLf
				errMsg = SendMail(email, mailSubj, mailMsg)
				if errMsg = "" then
					call InfoMessage("Username '" & username & "' added. Please check your email for the password.")
					mailMsg = "A new user has signed up for the pool." & vbCrLf _
					        & vbCrLf _
					        & "Username: " & username & vbCrLf _
					        & "Email: " & email & vbCrLf _
					        & vbCrLf
					errMsg = SendMail(ADMIN_EMAIL, mailSubj, mailMsg)
					username = ""
					email = ""
				else
					call ErrorMessage("Email send failed with '" & errMsg & "'<br />Please contact <a href=""mailto:" & ADMIN_EMAIL & """>" & ADMIN_EMAIL & "</a> for help.")
					sql = "DELETE FROM Users WHERE Username = '" & SqlString(username) & "'"
					call DbConn.Execute(sql)
				end if

			'Otherwise, show a message indicating the user may now log in.
			else
				call InfoMessage("Username '" & username & "' added. You may now <a href=""userLogin.asp"">login</a>.")
			end if

		end if

		'Show form field errors, if any.
		if FormFieldErrors.Count > 0 then
			call FormFieldErrorsMessage("Sign up failed, see below.")
		end if

	end if %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<table class="main fixed" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge">
				<th align="left">User Sign Up</th>
			</tr>
			<tr>
				<td class="freeForm">
					<p>To sign up, enter
<%	if SIGN_UP_INVITATION_CODE <> "" then %>
					the invitation code,
<%	end if
	if SERVER_EMAIL_ENABLED then %>
					your name and email address	and hit the <code>Submit</code> button.
					An account will	be created for you using a randomly generated password which will be sent to your email address.</p>
<%	else %>
					your name, email address and a password (twice, for confirmation) hit the <code>Submit</code> button.
<%	end if %>
					<!--<table class="form" cellpadding="0" cellspacing="0">
<%	if SIGN_UP_INVITATION_CODE <> "" then %>
						<tr valign="middle">
							<td><strong>Invitation code:</strong></td>
							<td><input type="password" name="code" value="" class="<% = FieldStyleClass("", "code") %>" /></td>
						</tr>
<%	end if %>
						<tr valign="middle">
							<td><strong>Username:</strong></td>
							<td><input type="text" name="username" value="<% = username %>" class="<% = FieldStyleClass("", "username") %>" /></td>
						</tr>
						<tr valign="middle">
							<td><strong>Email address:</strong></td>
							<td><input type="text" name="email" value="<% = email %>" class="<% = FieldStyleClass("", "email") %>" /></td>
						</tr>
<%	if not SERVER_EMAIL_ENABLED then %>
						<tr valign="middle">
							<td><strong>Password:</strong></td>
							<td><input type="password" name="password1" value="" class="<% = FieldStyleClass("", "password1") %>" /></td>
						</tr>
						<tr valign="middle">
							<td><strong>Confirm password:</strong></td>
							<td><input type="password" name="password2" value="" class="<% = FieldStyleClass("", "password1") %>" /></td>
						</tr>
<%	end if %>						
					</table>-->
<%	if SERVER_EMAIL_ENABLED then %>
					<p>If you don't receive your email within a few hours, please contact the Administrator at <a href="mailto:<% = ADMIN_EMAIL %>"><% = ADMIN_EMAIL %></a> for assistance.</p>
<%	end if %>
				</td>
			</tr>
		</table>
		<p><input type="submit" name="submit" value="Submit" class="button" title="Sign up for an account." />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel sign up." /></p>
	</form>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>
