<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Manage Accounts" : AdminOnly = true %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/playoffs.asp" -->
<!-- #include file="includes/side.asp" -->
<!-- #include file="includes/weekly.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'If there is form data, process it.
	dim username, description, amount
	dim sql, rs
	if Request.ServerVariables("Content_Length") > 0 and not CancelRequested() then
		username = Trim(Request.Form("username"))
		description = Trim(Request.Form("description"))
		amount = Trim(Request.Form("amount"))
		if username = "" then
			FormFieldErrors.Add "username", "A username is required."
		end if
		if description = "" then
			FormFieldErrors.Add "description", "A description is required."
		end if
		if amount = "" then
			FormFieldErrors.Add "amount", "An amount is required."
		elseif not IsNumeric(amount) then
			FormFieldErrors.Add "amount", "Amount must be numeric."
		elseif CDbl(amount) = 0 then
			FormFieldErrors.Add "amount", "Amount cannot be zero."
		elseif CDbl(amount) <> Round(amount, 2) then
			FormFieldErrors.Add "amount", "Amount is invalid."
		end if

		'If there were any errors, display the error messages. Otherwise, add
		'the transaction and redirect to the account history display.
		if FormFieldErrors.Count > 0 then
			call FormFieldErrorsMessage("Error: Invalid fields. Please correct and resubmit.")
		else
			sql = "INSERT INTO Credits" _
			   & " (Username, [Timestamp], Description, Amount)" _
			   & " VALUES('" & SqlString(username) & "'," _
			   & " #" & CurrentDateTime() & "#," _
			   & " '" & SqlString(description) & "'," _
			   & " " & amount & ")"
			call DbConn.Execute(sql)
			Response.Redirect("accountHistory.asp?username=" & username)
		end if
	end if

	'Get a list of users and build the form.
	dim users, i
	users = UsersList(true)
	if IsArray(users) then %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<table class="main fixed" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge">
				<th align="left">Add an Account Transaction</th>
			</tr>
			<tr>
				<td class="freeForm">
					<p>Select a username, enter a description and amount and hit <code>Add</code> to add the transaction.</p>
					<table cellpadding="0" cellspacing="0">
						<tr valign="middle">
							<td><strong>Username:</strong></td>
							<td>&nbsp;
								<select name="username" class="<% = FieldStyleClass("", "username") %>">
									<option value=""></option>
<%		for i = 0 to UBound(users) %>
									<option value="<% = users(i) %>" <% if users(i) = username then Response.Write(" selected=""selected""") end if %>><% = users(i) %></option>
<%		next %>
								</select>
							</td>
						</tr>
						<tr valign="middle">
							<td><strong>Description:</strong></td>
							<td>&nbsp;&nbsp;<input type="text" name="description" value="<% = description %>" size="40" maxlength="50" class="<% = FieldStyleClass("", "description") %>" /></td>
						</tr>
						<tr valign="middle">
							<td><strong>Amount:</strong></td>
							<td>$<input type="text" name="amount" value="<% = amount %>" size="4" class="<% = FieldStyleClass("numeric", "amount") %>" /></td>
						</tr>
					</table>
					<p>Transactions cannot be edited or deleted.
					If you make a mistake, you must correct it by adding a new transaction.</p>
				</td>
			</tr>
		</table>
		<p><input type="submit" name="submit" value="Add" class="button" title="Add the account transaction." />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel the update." /></p>
	</form>
<%	else %>
	<table class="main fixed" cellpadding="0" cellspacing="0">
		<tr class="header bottomEdge">
			<th align="left">Add Account Transaction</th>
		</tr>
		<tr>
			<td align="center"><em>No users found.</em></td>
		</tr>
	</table>
<% end if %>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>