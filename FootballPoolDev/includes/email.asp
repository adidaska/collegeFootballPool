<%	'**************************************************************************
	'* Common code for emailing.                                              *
	'*                                                                        *
	'* Note: The email process is host-specific. You will need to customize   *
	'* the SendMail() function to work on your target environment.            *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Returns an array of email addresses for users who have opted to receive
	' the given notification type.
	'--------------------------------------------------------------------------
	function GetNotificationList(notificationType)

		dim sql, rs
		dim list, email

		GetNotificationList = Array()

		sql = "SELECT EmailAddress FROM Users WHERE " & notificationType & " = true"
		set rs = DbConn.Execute(sql)
		do while not rs.EOF
			email = Decrypt(rs.Fields("EmailAddress").Value)
			if email <> "" then
				list = list & email & ";"
			end if
			rs.MoveNext
		loop
		if Len(list) > 0 then
			GetNotificationList = Split(list, ";")
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns true if the given address is in valid email address format, false
	' otherwise.
	'--------------------------------------------------------------------------
	function IsValidEmailAddress(addr)

		dim list, item
		dim i, c

		IsValidEmailAddress = true

		'Exclude any address with '..'.
		if InStr(addr, "..") > 0 then
			IsValidEmailAddress = false
			exit function
		end if

		'Split email address into the user and domain names.
		list = Split(addr, "@")
		if UBound(list) <> 1 then
			IsValidEmailAddress = false
			exit function
		end if

		'Check both names.
		for each item in list

			'Make sure the name is not zero length.
			if Len(item) <=	0 then
				IsValidEmailAddress = false
				exit function
			end if

			'Make sure only valid characters appear in the name.
			for i = 1 to Len(item)
				c = Lcase(Mid(item, i, 1))
				if InStr("abcdefghijklmnopqrstuvwxyz&_-.", c) <= 0 and not IsNumeric(c) then
					IsValidEmailAddress = false
					exit function
				end if
			next

			'Make sure the name does not start or end with invalid characters.
			if Left(item, 1) = "." or Right(item, 1) = "." then
				IsValidEmailAddress = false
				exit function
			end if

		next

		'Check for a '.' character in the domain name.
		if InStr(list(1), ".") <= 0 then
			IsValidEmailAddress = false
			exit function
		end if

	end function

	'--------------------------------------------------------------------------
	' Converts a formatted point spread to plain text, for use in email.
	'--------------------------------------------------------------------------
	function PlainTextPointSpread(spread)

		dim str

		PlainTextPointSpread = ""
		str = FormatPointSpread(spread)
		str = Replace(str, "&nbsp;", "")
		str = Replace(str, "&frac12;", " 1/2")
		PlainTextPointSpread = str

	end function

	'--------------------------------------------------------------------------
	' Sends an email and returns an empty string if successful. Otherwise, it
	' returns the error message.
	'--------------------------------------------------------------------------
	function SendMail(toAddr, subjectText, bodyText)

		dim mailer
		dim cdoMessage, cdoConfig

		'Assume all will go well.
		SendMail = ""

		'Configure the message.
		set cdoMessage = Server.CreateObject("CDO.Message")
		set cdoConfig = Server.CreateObject("CDO.Configuration")
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")  = 2
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "relay-hosting.secureserver.net"
		cdoConfig.Fields.Update
		set cdoMessage.Configuration = cdoConfig

		'Create the email.
		cdoMessage.From     = ADMIN_USERNAME & " <" & ADMIN_EMAIL & ">"
		cdoMessage.To       = toAddr
		cdoMessage.Subject  = subjectText
		cdoMessage.TextBody = bodyText

		'Send it.
		on error resume next
		cdoMessage.Send

		'If an error occurred, return the error description.
		if Err.Number <> 0 then
			SendMail = Err.Description
		end if

		'Clean up.
		set cdoMessage = Nothing
		set cdoConfig  = Nothing

	end function %>
