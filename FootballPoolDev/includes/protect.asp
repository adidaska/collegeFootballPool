<%	'If the user has not logged in, or if the page is designated as Admin-only
	'and the user is not the Administrator, redirect to the login page.
	if Session(SESSION_USERNAME_KEY) = "" or (AdminOnly and not IsAdmin()) then
		Response.Redirect("userLogin.asp")
	end if %>
