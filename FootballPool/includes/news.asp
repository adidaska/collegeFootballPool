<%	'**************************************************************************
	'* Common code for site news.                                             *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Returns a string containing the lines of news text from the database.
	'--------------------------------------------------------------------------
	function GetNews()

		dim sql, rs

		GetNews = ""
		sql = "SELECT * FROM News ORDER BY LineNumber"
		set rs = DbConn.Execute(sql)
		if not (rs.BOF and rs.EOF) then
			rs.MoveFirst
			do while not rs.EOF
				GetNews = GetNews & rs.Fields("Line").Value & vbCrLf
				rs.MoveNext
			loop
		end if

	end function

	'--------------------------------------------------------------------------
	' Formats and displays news text.
	'--------------------------------------------------------------------------
	sub DisplayNews(nTabs, news)

		dim lines, i
		lines = Split(news, vbCrLf)
		if IsArray(lines) then
			for i = 0 to UBound(lines)
				Response.Write(String(nTabs, vbTab) & lines(i) & vbCrLf)
			next
		end if

	end sub %>
