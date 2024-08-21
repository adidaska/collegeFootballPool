<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Manage Users" : AdminOnly = true %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/style.css" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/email.asp" -->
<!-- #include file="includes/encryption.asp" -->
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/passwords.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'If there is form data, process it.
	dim sql, rs, salt, restoreFields, badChar
	badChar = ""
	restoreFields = false
	if Request.ServerVariables("Content_Length") > 0 and not CancelRequested() then

		'Assume we will need to restore the form fields.
		restoreFields = true

		'Set the form field names based on which form was submitted.
		dim usernameField, userDisplaynameField, password1Field, password2Field
		dim username, displayname, email, password1, password2, formStr
		usernameField = ""
		displayname = ""
		password1Field = ""
		password2Field = ""
		formStr = "Request"
		if Request.Form("submit") = "Add" then
			usernameField = "addUsername"
			userDisplaynameField = "addDisplayname"
			password1Field = "addPassword1"
			password2Field = "addPassword2"
			formStr = "Add user request"
		elseif Request.Form("submit") = "Delete" then
			usernameField = "deleteUsername"
			formStr = "Delete user request"
		elseif Request.Form("submit") = "Reset" then
			usernameField = "resetUsername"
			password1Field = "resetPassword1"
			password2Field = "resetPassword2"
			formStr = "Password reset request"
		end if

		'Get the form field values.
		username  = Request.Form(usernameField)
		displayname = Request.Form(userDisplaynameField)
		password1 = Request.Form(password1Field)
		password2 = Request.Form(password2Field)
		email     = Request.Form("addEmail")

		'Check the username field.
		if username = "" then
			FormFieldErrors.Add usernameField, "You must select a username."
		else

			'For new user adds, make sure the username is valid and does not already exist.
			if Request.Form("submit") = "Add" then
				badChar = UsernameCheck(username)
				if badChar <> "" then
					FormFieldErrors.Add "username", "Usernames may not contain a '" & Server.HtmlEncode(badChar) & "' character."
				elseif Len(username) > USERNAME_MAX_LEN then
					FormFieldErrors.Add "username", "Usernames may not exceed " & USERNAME_MAX_LEN & " characters."
				else
					sql = "SELECT * FROM Users WHERE Username = '" & SqlString(username) &"'"
					set rs = DbConn.Execute(sql)
					if not (rs.BOF and rs.EOF) then
						FormFieldErrors.Add usernameField, "Username '" & username & "' already exists, try another."
					end if
				end if
			end if

			'For new user add, also validate the displayname field
			if Request.Form("submit") = "Add" then
                badChar = IsCleanAndValidDisplayname(displayname)
                if badChar then
                    FormFieldErrors.Add "displayname", "Displaynames may not contain a '" & Server.HtmlEncode(badChar) & "' character."
                end if
            end if

			'For user deletes, make sure the user has no picks for completed
			'games or any account transactions in the database.
			if Request.Form("submit") = "Delete" then
				sql = "SELECT COUNT(*) AS Total" _
				    & " FROM Picks, Schedule" _
				    & " WHERE Username = '" & SqlString(username) & "'" _
				    & " AND Picks.GameID = Schedule.GameID" _
				    & " AND NOT ISNULL(Result)"
				set rs = DbConn.Execute(sql)
				if not (rs.BOF and rs.EOF) then
					if (rs.Fields("Total").Value <> 0) then
						FormFieldErrors.Add usernameField, "User '" & username & "' has one or more closed pool entries, cannot delete."
					end if
				else
					sql = "SELECT COUNT(*) AS Total" _
					    & " FROM SidePicks, Schedule" _
					    & " WHERE Username = '" & SqlString(username) & "'" _
					    & " AND (Pick = VisitorID OR Pick = HomeID)" _
					    & " AND NOT ISNULL(Result)"
					set rs = DbConn.Execute(sql)
					if not (rs.BOF and rs.EOF) then
						if (rs.Fields("Total").Value <> 0) then
							FormFieldErrors.Add usernameField, "User '" & username & "' has one or more closed pool entries, cannot delete."
						end if
					end if
				end if
				if FormFieldErrors.Count = 0 then
					sql = "SELECT DISTINCT Username FROM Credits" _
					  &  " WHERE Username = '" & SqlString(username) & "'"
					set rs = DbConn.Execute(sql)
					if not (rs.BOF and rs.EOF) then
						FormFieldErrors.Add usernameField, "User '" & username & "' has one or more account transactions, cannot delete."
					end if
				end if
			end if

		end if

		'For new user adds, check the email field.
		if Request.Form("submit") = "Add" and email <> "" then
			if not IsValidEmailAddress(email) then
				FormFieldErrors.Add "addEmail", "'" & email & "' does not appear to be a valid email address."
			end if
		end if

		'For new user adds and password resets, check the password fields.
		if Request.Form("submit") = "Add" or Request.Form("submit") = "Reset" then
			if password1 = "" or password2 = "" then
				FormFieldErrors.Add password1Field, "You must enter a password twice, where indicated."
				FormFieldErrors.Add password2Field, ""
			elseif password1 <> "" and password1 <> password2 then
				FormFieldErrors.Add password1Field, "Passwords did not match."
				FormFieldErrors.Add password2Field, ""
			elseif Len(password1) < PASSWORD_MIN_LEN then
				FormFieldErrors.Add password1Field, "Password must be at least " & PASSWORD_MIN_LEN & " characters long."
				FormFieldErrors.Add password2Field, ""
			end if
		end if

		'If there where no errors, do the appropriate update.
		if FormFieldErrors.Count = 0 then
			if Request.Form("submit") = "Add" then
				salt = CreateSalt()
				sql = "INSERT INTO Users" _
				  & " (Username, UserDisplayname, EmailAddress, Salt, [Password])" _
				  & " VALUES('" & SqlString(username) & "'," _
				  & " '" & SqlString(displayname) & "'," _
				  & " '" & Encrypt(email) & "'," _
				  & " '" & salt & "'," _
				  & " '" & Hash(salt & password1) & "')"
				call DbConn.Execute(sql)
				call InfoMessage("User '" & username &"' has been added.")
			elseif Request.Form("submit") = "Delete" then
				sql = "DELETE FROM Users WHERE Username = '" & SqlString(username) & "'"
				call DbConn.Execute(sql)
				sql = "DELETE FROM Picks WHERE Username = '" & SqlString(username) & "'"
				call DbConn.Execute(sql)
				sql = "DELETE FROM Tiebreaker WHERE Username = '" & SqlString(username) & "'"
				call DbConn.Execute(sql)
				sql = "DELETE FROM Messages WHERE Username = '" & SqlString(username) & "'"
				call DbConn.Execute(sql)
				call InfoMessage("User '" & username &"' has been deleted.")
			elseif Request.Form("submit") = "Reset" then
				salt = CreateSalt()
				sql = "UPDATE Users SET" _
				   & " Salt = '" & salt & "', " _
				   & " [Password] = '" & Hash(salt & password1) & "'" _
				   & " WHERE Username = '" & SqlString(username) & "'"
				call DbConn.Execute(sql)
				call InfoMessage("Password for user '" & username &"' has been reset.")
			end if

			'The request was completed so do not restore the form fields.
			restoreFields = false

		'Otherwise, show the errors.
		else
			call FormFieldErrorsMessage(formStr & " failed, see below.")
		end if

	end if

	'Build the form displays. %>
	<!-- Add a new user form. -->
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<table class="main fixed" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge">
				<th align="left">Add a New User</th>
			</tr>
			<tr>
				<td class="freeForm">
					<p>To add a new user, enter a new username and password (twice, for confirmation) in the fields provided.</p>
					<table cellpadding="0" cellspacing="0">
						<tr valign="middle">
							<td><strong>Username:</strong></td>
							<td><input type="text" name="addUsername" value="<% if restoreFields then Response.Write(Request.Form("addUsername")) end if %>" class="<% = FieldStyleClass("", "addUsername") %>" /></td>
						</tr>
						<tr valign="middle">
							<td><strong>Email address:</strong></td>
							<td><input type="text" name="addEmail" value="<% if restoreFields then Response.Write(Request.Form("addEmail")) end if %>" size="30" class="<% = FieldStyleClass("", "email") %>" /></td>
						</tr>
						<tr valign="middle">
							<td><strong>Password:</strong></td>
							<td><input type="password" name="addPassword1" value="" class="<% = FieldStyleClass("", "addPassword1") %>" /></td>
						</tr>
						<tr valign="middle">
							<td><strong>Confirm password:</strong></td>
							<td><input type="password" name="addPassword2" value="" class="<% = FieldStyleClass("", "addPassword2") %>" /></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<p><input type="submit" name="submit" value="Add" class="button" title="Add the new user." />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel the update." /></p>
	</form>
<%	'Get a list of users.
	dim users, i
	users = UsersList(true)
	if IsArray(users) then %>
	<!-- Delete a user form. -->
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<table class="main fixed" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge">
				<th align="left">Delete a User</th>
			</tr>
			<tr>
				<td class="freeForm">
					<p>To delete an existing user, select the username below.
					Note that you may not delete users who have made entries for completed games or who have account transactions.</p>
					<table cellpadding="0" cellspacing="0">
						<tr valign="middle">
							<td><strong>Username:</strong></td>
							<td>
								<select name="deleteUsername" class="<% = FieldStyleClass("", "deleteUsername") %>">
									<option value=""></option>
<%		for i = 0 to UBound(users) %>
									<option value="<% = users(i) %>" <% if restoreFields and users(i) = Request.Form("deleteUsername") then Response.Write(" selected=""selected""") end if %>><% = users(i) %></option>
<%		next %>
								</select>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<p><input type="submit" name="submit" value="Delete" class="button" title="Delete the user." />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel the update." /></p>
	</form>
	<!-- Reset a user's password form. -->
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<table class="main fixed" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge">
				<th align="left">Reset a User's Password</th>
			</tr>
			<tr>
				<td class="freeForm">
					<p>To reset the password of an existing user, select the username and enter the new password (twice, for confirmation) in the fields provided.</p>
					<table cellpadding="0" cellspacing="0">
						<tr valign="middle">
							<td><strong>Username:</strong></td>
							<td>
								<select name="resetUsername" class="<% = FieldStyleClass("", "resetUsername") %>">
									<option value=""></option>
<%		for i = 0 to UBound(users) %>
									<option value="<% = users(i) %>" <% if restoreFields and users(i) = Request.Form("resetUsername") then Response.Write(" selected=""selected""") end if %>><% = users(i) %></option>
<%		next %>
								</select>
							</td>
						</tr>
						<tr valign="middle">
							<td><strong>New password:</strong></td>
							<td><input type="password" name="resetPassword1" value="" class="<% = FieldStyleClass("", "resetPassword1") %>" /></td>
						</tr>
						<tr valign="middle">
							<td><strong>Confirm new password:</strong></td>
							<td><input type="password" name="resetPassword2" value="" class="<% = FieldStyleClass("", "resetPassword2") %>" /></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<p><input type="submit" name="submit" value="Reset" class="button" title="Reset the user's password." />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel the update." /></p>
	</form>
<%	end if %>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>