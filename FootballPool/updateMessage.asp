<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Edit Post" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/common.css" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
	<script type="text/javascript" src="scripts/tiny_mce/tiny_mce.js"></script>
	<script type="text/javascript" src="scripts/messages.js"></script>
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/messages.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'Get the message ID passed in the query string or form data.
	dim id
	id = Request.Form("id")
	if id = "" then
		id = Request.QueryString("id")
	end if

	'Look up the specified message.
	dim created, username, message, lastModified
	dim sql, rs
	dim isValid
	isValid = false
	if IsNumeric(id) then
		sql = "SELECT * FROM Messages WHERE MessageID = " & id
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			created      = rs.Fields("Created").Value
			username     = rs.Fields("Username").Value
			message      = rs.Fields("Message").Value
			lastModified = rs.Fields("LastModified").Value

			'Only the Admin and the original poster may update messages.
			if IsAdmin() or Session(SESSION_USERNAME_KEY) = username then
				isValid = true
			end if
		end if
	end if

	'If the message ID is invalid or if the user is not permitted to update
	'this message, redirect to the home page.
	if not isValid then
		Response.Redirect("./")
	end if

	'If there is form data, process it.
	dim needDeleteConfirmation, postDeleted
	needDeleteConfirmation = false
	postDeleted = false
	if Request.ServerVariables("Content_Length") > 0 then

		'Get the updated message from the form data.
		if not CancelRequested then
			message = GetFieldValue("message", message)
		end if

		'Process an update request.
		if Request.Form("submit") = "Update" then

			'If the user is disabled, prevent the update.
			if IsDisabled() then
				call ErrorMessage("Error: Your account has been disabled, changes not accepted.<br />Please contact the Administrator.")
			else

				'Validate the message field.
				if Len(message) = 0 then
					FormFieldErrors.Add "message", "Messages may not be empty."
				elseif Len(message) > MAX_MESSAGE_LENGTH then
					FormFieldErrors.Add "message", "Message exceeds allowed length."
				else
					username = Session(SESSION_USERNAME_KEY)
					lastModified = CurrentDateTime()
					sql = "UPDATE Messages SET" _
					   & " Message = '" & SqlString(message) & "'," _
					   & " LastModified = #" & lastModified & "#" _
					   & " WHERE MessageID = " & id
					call DbConn.Execute(sql)
					Session(SESSION_MESSAGE_KEY) = "Message updated."
					Response.Redirect("messageBoard.asp?id=" & id)
				end if
			end if
		end if

		'Process a delete request.
		if Request.Form("submit") = "Delete" then
			if LCase(Request.Form("confirmDelete")) <> "true" then
				needDeleteConfirmation = true
				call ErrorMessage("To confirm deletion, check the box below and press <code>Delete</code> again.<br />Pressing any other button will cancel the deletion.")
			else
				sql = "DELETE FROM Messages WHERE MessageID = " & id
				call DbConn.Execute(sql)
				Session(SESSION_MESSAGE_KEY) = "Message deleted."
				Response.Redirect("messageBoard.asp")
			end if
		end if

		'Display any errors.
		if FormFieldErrors.Count > 0 then
			call FormFieldErrorsMessage("Error: Invalid fields. Please correct and resubmit.")
		end if
	end if

	'Build the display. %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<input type="hidden" name="id" value="<% = id %>" />
		<table class="main fixed" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge" valign="bottom">
				<th align="left">Edit Post</th>
			</tr>
			<tr valign="top">
				<td class="freeForm">
					<p>Edit your post below.
					Note that messages may not exceed <% = MAX_MESSAGE_LENGTH %> total characters (including formatting).</p>
					<table class="<% = FieldStyleClass("htmlEditorArea", "message") %>"><tr><td style="padding: 0px;">
						<div>
							<textarea id="message" name="message" rows="8" cols="80" class="<% = FieldStyleClass("mceEditor", "message") %>"><% = message %></textarea>
						</div>
					</td></tr></table>
					<p></p>
					<table cellpadding="0" cellspacing="0">
						<tr>
							<td><strong>Created:</strong></td>
							<td><% = FormatPostDate(created) & " at " & FormatFullTime(created) %></td>
						</tr>
						<tr>
							<td><strong>Last Modified:</strong></td>
							<td><% = FormatPostDate(lastModified) & " at " & FormatFullTime(lastModified) %></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<p></p>
		<table cellpadding="0" cellspacing="0" style="width: 100%;">
			<tr valign="middle">
				<td style="padding: 0px;"><input type="submit" name="submit" value="Update" class="button" title="Update this post." />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel the update." /></td>
<%	'If delete was requested, add a confirmation checkbox.
	if needDeleteConfirmation then %>
				<td align="right" style="padding: 0px;"><input type="checkbox" id="confirmDelete" name="confirmDelete" value="true" /> <label for="confirmDelete">Confirm Deletion</label>&nbsp;</td>
<%	end if %>
				<td align="right" style="padding: 0px;"><input type="submit" name="submit" value="Delete" class="button" title="Delete this post." /></td>
			</tr>
		</table>
	</form>
	<p><a href="messageBoard.asp">Back to Message Board...</a></p>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>
