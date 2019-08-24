<%	'**************************************************************************
	'* Common code for the message board.                                     *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Constant definitions.
	'--------------------------------------------------------------------------

	'Post update message session variable.
	const SESSION_MESSAGE_KEY = "PostMessage"

	'--------------------------------------------------------------------------
	' Purges expired post records.
	'--------------------------------------------------------------------------
	sub PurgeExpiredPosts()

		dim dt, sql, rs

		if MAX_POST_AGE > 0 then
			dt = DateAdd("d", -MAX_POST_AGE, CurrentDateTime())
			sql = "DELETE FROM Messages WHERE Created < #" & dt & "#"
			call DbConn.Execute(sql)
		end if

	end sub

	'--------------------------------------------------------------------------
	' Purges excess post records.
	'--------------------------------------------------------------------------
	sub PurgeExcessPosts()

		dim sql, rs, dt, count

		'If over the maximum post count, purge oldest posts.
		if MAX_POST_COUNT > 0 then
			sql = "SELECT TOP " & MAX_POST_COUNT & " Created FROM Messages ORDER BY Created DESC"
			set rs = Server.CreateObject("ADODB.Recordset")
			rs.CursorLocation = 3
			rs.Open sql, DbConn
			rs.MoveLast
			dt = rs.Fields("Created").Value
			rs.Close
			sql = "DELETE FROM Messages WHERE Created < #" & dt & "#"
			DbConn.Execute(sql)
		end if

	end sub

	'--------------------------------------------------------------------------
	' Returns a formatted date string (day of week, month, day and year) given
	' a Date object.
	'--------------------------------------------------------------------------
	function FormatPostDate(dt)

		FormatPostDate = WeekdayName(Weekday(created), true) & ", " _
			& MonthName(Month(created), true) & " " _
			& Day(created) & ", " _
			& Year(created)

	end function %>
