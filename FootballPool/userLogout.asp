<%@ LANGUAGE="VBScript" %>
<%	'Clear session and redirect to the login page.
	Response.Buffer = true
	Session.Abandon
	Response.Redirect("userLogin.asp") %>
