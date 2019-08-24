<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Reset Results" : AdminOnly = true %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/protect.asp" -->
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
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->
<!-- #include file="includes/form.asp" -->
<!-- #include file="includes/side.asp" -->
<!-- #include file="includes/weekly.asp" -->
<%	'Open the database.
	call OpenDB() %>
	<table id="wrapper"><tr><td style="padding: 0px;">

<%	'If there is form data, process it.
	if Request.ServerVariables("Content_Length") > 0 and not CancelRequested() then
		'Find the total number of weeks in the season.
		dim week, numWeeks
		numWeeks = NumberOfWeeks()

		'Get checked fields.
		dim weekly, margin, survivor
		weekly   = FormFieldExists("weekly")
		margin   = FormFieldExists("margin")
		survivor = FormFieldExists("survivor")

		if weekly or margin then
			for week = 1 to numWeeks
				if weekly then
					call ClearWeeklyResultsCache(week)
				end if
				if margin then
					call ClearMarginResultsCache(week)
				end if
			next
		end if
		if survivor then
			ClearSurvivorStatus(1)
		end if

		'Build the info message.
		dim str
		str = ""
		if weekly then
			str = "weekly"
		end if
		if margin then
			if Len(str) > 0 then
				str = str & ", "
			end if
			str = str & "margin"
		end if
		if survivor then
			if Len(str) > 0 then
				str = str & ", "
			end if
			str = str & "survivor"
		end if
		if Len(str) > 0 then
			str = StrReverse(str)
			str = Replace(str, " ,", " dna ", 1, 1)
			str = StrReverse(str)
			call InfoMessage("Cached " & str & " results cleared.")
		else
			call ErrorMessage("No selections made.")
		end if
	end if %>
	<form action="<% = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<table class="main fixed" cellpadding="0" cellspacing="0">
			<tr class="header bottomEdge">
				<th align="left">Reset Results</th>
			</tr>
			<tr>
				<td class="freeForm">
					<p>Select the type of results you wish to reset.</p>
					<p style="margin-left: 2em;">
						<input type="checkbox" id="weekly" name="weekly" value="true" /> <label for="weekly">Weekly pool results.</label><br />
<%	if ENABLE_MARGIN_POOL then %>
						<input type="checkbox" id="margin" name="margin" value="true" /> <label for="margin">Margin pool results.</label><br />
<%	end if %>
<%	if ENABLE_SURVIVOR_POOL then %>
						<input type="checkbox" id="survivor" name="survivor" value="true" /> <label for="survivor">Survivor pool results.</label><br />
<%	end if %>
					</p>
					<p>Resetting results will not cause you to lose any data.
					However, after a reset, some pages will take longer to load the first time they are accessed.
					After that initial access, response times will return to normal.</p>
				</td>
			</tr>
		</table>
		<p><input type="submit" name="submit" value="Submit" class="button" title="Clear selected results caches." />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel the request." /></p>
	</form>
	</td></tr></table>
<!-- #include file="includes/footer.asp" -->
</body>
</html>
