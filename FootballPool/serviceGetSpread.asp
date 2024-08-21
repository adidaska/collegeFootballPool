<%
    Dim url
    url = "https://www.espn.com/college-football/game/_/gameId/" & Request.QueryString("gameID")

    Dim http
    Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
    http.Open "GET", url, False
    http.Send

    Response.ContentType = "application/html"
    Response.Write http.responseText
    Set http = Nothing
%>