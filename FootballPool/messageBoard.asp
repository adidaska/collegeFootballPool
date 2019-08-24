<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Message Board" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
	<script type="text/javascript" src="scripts/tiny_mce/tiny_mce.js"></script>
	<script type="text/javascript" src="scripts/messages.js"></script>
	<link href="styles/style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/messages.asp" -->


<div class="clearfix" id="content-wrap">
  	<div id="content-top"></div>
    <div id="primary" class="hfeed">
    
    
	<!--<table id="wrapper"><tr><td style="padding: 0px;">-->
<%	'If the message board is disabled, redirect to the home page.
	if not ENABLE_MESSAGE_BOARD then
		Response.Redirect("./")
	end if

	'Open the database.
	call OpenDB()

	'Purge any expired posts.
	call PurgeExpiredPosts()

	'If there is form data, process it.
	dim created, username, message, lastModified
	dim sql, rs
	if Request.ServerVariables("Content_Length") > 0 then
		if (Request.Form("submit") = "Post") and Session("Post") = false then

			'If the user is disabled, prevent posting.
			if IsDisabled() then
				call ErrorMessage("Error: Your account has been disabled, post not accepted.<br />Please contact the Administrator.")
			else

				'Validate the message field.
				message = Request.Form("message")
				if Len(message) = 0 then
					FormFieldErrors.Add "message", "Messages may not be empty."
				elseif Len(message) > MAX_MESSAGE_LENGTH then
					FormFieldErrors.Add "message", "Message exceeds allowed length."
				else

					'Add the new post.
					created = CurrentDateTime()
					lastModified = created
					username = Session(SESSION_USERNAME_KEY)
					sql = "INSERT INTO Messages" _
					   & " (Created, Username, Message, LastModified)" _
					   & " VALUES(#" & created & "#," _
					   & " '" & SqlString(username) & "'," _
					   & " '" & SqlString(message) & "'," _
					   & " #" & lastModified & "#)"
					call DbConn.Execute(sql)
					Session(SESSION_MESSAGE_KEY) = "Message added."

					'Purge any excess posts.
					call PurgeExcessPosts()

					'Redirect back to this page.
					Response.Redirect(Request.ServerVariables("SCRIPT_NAME"))
				end if

			end if
		end if

		'Display any errors.
		if FormFieldErrors.Count > 0 then
			call FormFieldErrorsMessage("Error: Invalid fields. Please correct and resubmit.")
		end if
	end if

	'If a post update was set (during edit or delete), show that message and
	'clear the session variable.
	if Session(SESSION_MESSAGE_KEY) <> "" then
		call InfoMessage(Session(SESSION_MESSAGE_KEY))
		Session.Contents.Remove(SESSION_MESSAGE_KEY)
	end if

	'Build the display. %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<table width="570px" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge" valign="bottom">
				<th align="left" colspan="3">Message Board</th>
			</tr>
<%	dim currentPage, pageCount, messageId, i

	'Retrive posts from the database.
	sql = "SELECT * FROM Messages ORDER BY Created DESC"
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.PageSize = POST_PAGE_SIZE
	rs.Open sql, DbConn

	'Set the current page.
	currentPage = Request.QueryString("page")
	if IsNumeric(currentPage) then
		currentPage = CInt(currentPage)
	else
		currentPage = 0
	end if
	if currentPage < 1 then
		currentPage = 1
	end if
	if currentPage > rs.PageCount then
		currentPage = rs.PageCount
	end if

	'If a message ID is was specified in the query string, show the page
	'containing it.
	dim id
	id = Request.QueryString("id")
	if IsNumeric(id) then
		id = CInt(id)
		do while not rs.EOF
			if rs.Fields("MessageID").Value = id then
				currentPage = rs.AbsolutePage
				rs.MoveFirst
				exit do
			end if
			rs.MoveNext
		loop
	end if

	'Show posts.
	if rs.PageCount > 0 then

		'Set up page navigation.
		rs.AbsolutePage = currentPage
		if rs.PageCount > 1 then %>
			<tr class="subHeader bottomEdge">
				<th align="left" style="width: 20%;">
<%			if currentPage > 1 then %>
					<a href="<% = Request.ServerVariables("SCRIPT_NAME") & "?page=" & (currentPage - 1) %>">&laquo; Newer</a>
<%			end if %>
				&nbsp;
				</th>
				<th align="center" style="width: 60%;">Page <% = currentPage & " of " & rs.PageCount %></th>
				<th align="right" style="width: 20%;">
				&nbsp;
<%			if currentPage < rs.PageCount then %>
					<a href="<% = Request.ServerVariables("SCRIPT_NAME") & "?page=" & (currentPage + 1) %>">Older &raquo;</a>
<%			end if %>
				</th>
			</tr>
<%		end if

		'Show posts.
		dim alt
		alt = false
		for i = 1 to rs.PageSize
			messageId    = rs.Fields("MessageID").Value
			created      = rs.Fields("Created").Value
			username     = rs.Fields("Username").Value
			message      = rs.Fields("Message").Value
			lastModified = rs.Fields("LastModified").Value
			if alt then %>
			<tr class="alt" valign="top">
<%			else %>
			<tr valign="top">
<%			end if
			alt = not alt %>
				<td style="white-space: nowrap;">
					<strong><% = username %></strong><br />
					<em class="small"><% = FormatPostDate(created) %><br />
					<% = FormatFullTime(created) %></em>
				</td>
				<td colspan="2">
<%			if created <> lastModified then %>
					<div title="Modified <% = FormatPostDate(lastModified) & " at " & FormatFullTime(lastModified) %>">
<%			else %>
					<div>
<%			end if
			if IsAdmin() or username = Session(SESSION_USERNAME_KEY) then %>
					<span style="float: right; margin: 0px 0px 1ex 1em;"><a href="updateMessage.asp?id=<% = messageId %>" title="Edit or remove this post.">Edit/Remove</a></span>
<%			end if %>
						<div class="message"><% = message %></div>
					</div>
				</td>
			</tr>
<%			rs.MoveNext
			if rs.EOF then
				exit for
			end if
		next
	else %>
			<tr>
				<td align="center" colspan="3"><em>No messages.</em></td>
			</tr>
<%	end if %>
		<tr class="header topEdge bottomEdge">
			<th align="left" colspan="3">Post a New Message</th>
		</tr>
		<tr valign="top">
			<td colspan="3" class="freeForm">
				<p>Enter your post below.
				Note that messages may not exceed <% = MAX_MESSAGE_LENGTH %> total characters (including formatting).</p>
				<table class="<% = FieldStyleClass("htmlEditorArea", "message") %>"><tr><td style="padding: 0px;">
					<div>
<%	message = ""
	if Request.ServerVariables("Content_Length") > 0 and not CancelRequested then
		message = Trim(Request.Form("message"))
	end if %>
						<textarea id="message" name="message" rows="8" cols="80" class="mceEditor"><% = message %></textarea>
					</div>
				</td></tr></table>
			</td>
		</tr>
		</table>
		<p><input type="submit" name="submit" value="Post" class="button" title="Post your message." />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel post." /></p>
	</form>
<% 'Set up page navigation.
	dim str
	str = ""
	if rs.PageCount > 1 then
		for i = 1 to rs.PageCount
			if i > 1 then
				str = str & " &middot; "
			end if
			str = str & "<a href=""" _
			   & Request.ServerVariables("SCRIPT_NAME") _
			   & "?page=" & i
			str = str & """>" & i & "</a>"
		next %>
		<p><strong>Go to Page:</strong> <% = str %></p>
<%	end if %>
<!--	</td></tr></table>-->
    
      </div> <!-- end of the primary div in the container-->
    
<div id="content-btm"></div>
</div>
<!-- #include file="includes/footer.asp" -->
</body>
</html>

