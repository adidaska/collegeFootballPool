	<!-- Page footer. -->
	<!--<div id="subTitle"><% = PageSubTitle %></div>-->
	<div id="footer">
		<div id="contact">For help, contact <a href="mailto:<% = ADMIN_EMAIL %>"><% = ADMIN_EMAIL %></a>.</span></div>
		<div id="copyright">&copy; 1999-2012 by <a href="http://www.brainjar.com/">Mike Hall</a>. Updated by <a href="mailto:adidaska@megstudios.com">Evan Graham</a>.</div>
	</div>
    <%	'Add the current time (ET) so the client-side JavaScript can retrieve it.
	
	et = CurrentDateTime() %>
	<div id="serverTimestamp" style="display: none;"><% = DatePart("yyyy", et) & "/" & DatePart("m", et) & "/" & DatePart("d", et) & " " & FormatDateTime(et, vbLongTime) %></div>
    <!--<script src="js/jaflGames.js"></script>-->
    <!--<script src="scripts/analyticstracking.js"></script>-->