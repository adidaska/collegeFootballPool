<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Edit News" : AdminOnly = true %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/style.css" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
	<script type="text/javascript" src="scripts/tiny_mce/tiny_mce.js"></script>
	<script type="text/javascript">//<![CDATA[
	tinyMCE.init({
		apply_source_formatting : true,
		cleanup_on_startup : true,
		content_css : "styles/common.css",
		convert_fonts_to_spans : true,
		doctype : '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">',
		editor_selector : "mceEditor",
		fix_list_elements : true,
		font_size_style_values : "xx-small,x-small,small,medium,large,x-large,xx-large",
		gecko_spellcheck : true,
		hide_selects_on_submit : true,
		inline_styles : true,
		mode: "textareas",
		plugins : "iespell,layer,table",
		strict_loading_mode : true,
		theme : "pool",
		theme_pool_buttons1 : "formatselect,fontselect,fontsizeselect,|,forecolor,backcolor,|,bold,italic,underline,strikethrough",
		theme_pool_buttons2 : "insertlayer,moveforward,movebackward,absolute,|,tablecontrols,|,visualaid",
		theme_pool_buttons3 : "link,unlink,anchor,image,charmap,|,bullist,numlist,|,indent,outdent,|,justifyleft,justifyright,justifycenter,justifyfull,|,undo,redo,|,removeformat,cleanup,iespell,code,|,help"
	});
	//]]></script>
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/email.asp" -->
<!-- #include file="includes/encryption.asp" -->
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/News.asp" -->
	<table id="wrapper"><tr><td style="padding: 0px;">
<%	'Open the database.
	call OpenDB()

	'Get the current news data.
	dim news
	news = GetNews()

	'If an update was requested, process it.
	dim lines, line, i
	dim notify, infoMsg
	dim sql
	if Request.Form("submit") = "Update" then

		'Delete any existing news.
		sql = "DELETE FROM News"
		call DbConn.Execute(sql)

		'Insert the new content.
		lines = Split(Request.Form("newContent"), vbCrLf)
		if IsArray(lines) then
			for i = 0 to UBound(lines)
				line = lines(i)
				line = Replace(line, vbCrLf, " ")
				line = Replace(line, vbTab, " ")
				line = Trim(line)
				if Len(line) > 0 then
					sql = "INSERT INTO News"_
					  & " (LineNumber, Line)" _
					  & " VALUES(" & (i + 1) & ", '" & SqlString(line) & "')"
					call DbConn.Execute(sql)
				end if
			next
		end if
		news = GetNews()

		'Send out email notifications, if checked.
		infoMsg = "News updates saved."
		notify = Trim(Request.Form("notify"))
		if LCase(notify) = "true" then
			infoMsg = "News updates saved, notifications sent."
			call SendNotifications()
		end if

		'Updates done, show an informational message.
		call InfoMessage(infoMsg)
	end if

	'Display the form.
	dim editContent
	if Request.ServerVariables("Content_Length") > 0 and not CancelRequested() then
		editContent = Request.Form("newContent")
	else
		editContent = news
	end if %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<table class="main fixed" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge">
				<th align="left">Edit the News Section</th>
			</tr>
			<tr>
				<td class="freeForm">
					<p>You can edit the "News" section of the home page below:</p>
					<table class="htmlEditorArea"><tr><td style="padding: 0px;">
						<div style="margin: 0px;">
							<textarea name="newContent" class="mceEditor" rows="30" cols="80" style="width: 39em;"><% = editContent %></textarea>
						</div>
					</td></tr></table>
				</td>
			</tr>
<%	if SERVER_EMAIL_ENABLED then %>
			<tr class="subHeader topEdge">
				<th align="left"><input type="checkbox" id="notify" name="notify" value="true" /> <label for="notify">Send update notification to users.</label></th>
			</tr>
<%	end if %>
		</table>
		<p><input type="submit" name="submit" value="Update" class="button" title="Save changes." />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel the change." /></p>
	</form>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>
<%	'**************************************************************************
	'* Local functions and subroutines.                                       *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Sends an email notification to any users who have elected to receive
	' them.
	'--------------------------------------------------------------------------
	sub SendNotifications()

		dim subj, body
		dim sql, rs
		dim email

		subj = "Football Pool News Notification"
		body = "The site news has been updated."
		sql = "SELECT EmailAddress FROM Users WHERE NotifyOfNewsUpdates = true"
		set rs = DbConn.Execute(sql)
		do while not rs.EOF
			email = Decrypt(rs.Fields("EmailAddress").Value)
			if email <> "" then
				call SendMail(email, subj, body)
			end if
			rs.MoveNext
		loop

	end sub %>