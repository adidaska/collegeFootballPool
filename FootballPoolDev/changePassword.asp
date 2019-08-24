<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Change Password" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link href="styles/style.css" rel="stylesheet" type="text/css" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/passwords.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'If there is form data, process it.
	dim password, newPassword1, newPassword2
	dim sql, rs, salt
	if Request.ServerVariables("Content_Length") > 0 and not CancelRequested() then
		password = Trim(Request.Form("password"))
		newPassword1 = Trim(Request.Form("newPassword1"))
		newPassword2 = Trim(Request.Form("newPassword2"))
		if password = "" or newPassword1 = "" or newPassword2 = "" then
			FormFieldErrors.Add "password", "Your current password is required."
			FormFieldErrors.Add "newPassword1", "You must enter your new password twice, where indicated."
			FormFieldErrors.Add "newPassword2", ""
		elseif newPassword1 <> "" and newPassword1 <> newPassword2 then
			FormFieldErrors.Add "password", "New passwords did not match."
			FormFieldErrors.Add "newPassword1", ""
			FormFieldErrors.Add "newPassword2", ""
		elseif Len(newPassword1) < PASSWORD_MIN_LEN then
			FormFieldErrors.Add "password", "New password must be at least " & PASSWORD_MIN_LEN & " characters long."
			FormFieldErrors.Add "newPassword1", ""
			FormFieldErrors.Add "newPassword2", ""
		else

			'The input data is good, check the current password.
			sql = "SELECT * FROM Users WHERE Username = '" & SqlString(Session(SESSION_USERNAME_KEY)) & "'"
			set rs = DbConn.Execute(sql)
			if not (rs.BOF and rs.EOF) then
				if Hash(rs.Fields("Salt").Value & password) <> rs.Fields("Password").Value then
					FormFieldErrors.Add "password", "Incorrect password."
					FormFieldErrors.Add "newPassword1", ""
					FormFieldErrors.Add "newPassword2", ""
				else
					salt = CreateSalt()
					sql = "UPDATE Users SET" _
					   & " Salt = '" & salt & "', " _
					   & " [Password] = '" & Hash(salt & newPassword1) & "'" _
					   & " WHERE Username = '" & SqlString(Session(SESSION_USERNAME_KEY)) & "'"
					call DbConn.Execute(sql)
					call InfoMessage("Your password has been changed.")
				end if
			else
				call ErrorMessage("Password update failed, user '" & Session(SESSION_USERNAME_KEY) & "' not found.")
			end if

		end if

		'Show form field errors, if any.
		if FormFieldErrors.Count > 0 then
			call FormFieldErrorsMessage("Password change failed, see below.")
		end if

	end if %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<table class="main fixed" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge">
				<th align="left">Change Password</th>
			</tr>
			<tr>
				<td class="freeForm">
					<p>To change your password, you must provide your
					current password and enter your new password (twice,
					for confirmation) in the fields provided.</p>
					<table cellpadding="0" cellspacing="0">
						<tr valign="middle">
							<td><strong>Username:</strong></td>
							<td><% = Session(SESSION_USERNAME_KEY) %></td>
						</tr>
						<tr valign="middle">
							<td><strong>Current password:</strong></td>
							<td><input type="password" name="password" value="" class="<% = FieldStyleClass("", "password") %>" /></td>
						</tr>
						<tr valign="middle">
							<td><strong>New password:</strong></td>
							<td><input type="password" name="newPassword1" value="" class="<% = FieldStyleClass("", "newPassword1") %>" /></td>
						</tr>
						<tr valign="middle">
							<td><strong>Confirm new password:</strong></td>
							<td><input type="password" name="newPassword2" value="" class="<% = FieldStyleClass("", "newPassword2") %>" /></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<p><input type="submit" name="submit" value="Change" class="button" title="Change your password." />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel the change." /></p>
	</form>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>