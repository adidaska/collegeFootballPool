<%@ LANGUAGE="VBScript" %>
<!-- #include file="includes/common.asp" --><% PageSubTitle = "Help" %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title><% = PAGE_TITLE & ": " & PageSubTitle %></title>
	<link rel="shortcut icon" href="favicon.ico" />
	<link rel="stylesheet" type="text/css" href="styles/common.css" />
	<link rel="stylesheet" type="text/css" href="styles/menu.css" />
	<script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
    <style type="text/css">
<!--
.style1 {font-size: 14pt}
-->
    </style>
<link href="styles/style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/menu.asp" -->


<div class="clearfix" id="content-wrap">
  	<div id="content-top"></div>
        <div id="primary" class="hfeed">
    
		<%	'Open the database.
		call OpenDB() %>
	  <div class="game-Title">
			Help<br />
	  </div>
        
	  <div class="left_aligned_style">
        	<h1>Pool Rules</h1><br />
        <div>
            <div>
              <h2>ENTRY FEE PAYMENT </h2>
        <p>Must be obtained before the start of week one in order to participate in the pool. </p>
              <h2>TEAM NAMES </h2>
              <p>Although not required, we encourage each participant to come up with a team name. We will leave the level of edginess and crudeness to your imagination. But if we deem it to be over the top and offensive to others in the pool, we may ask you to find another name. If you want to keep your name from last year, please indicate such when you respond to this email. </p>
              <h2>THE RULES </h2>
              <p>Every week we will pick approximately 15 games depending upon the quality of matchups that week. We will then research the betting lines from various websites and you will pick against the spread. If you do not know how to pick against the spread, please let us know. One game will be designated as the &quot;game of the week&quot; and you will have to pick a score that will serve as our tiebreaker for the weekly winner. You will receive an emailed no latter than Tuesday at noon indicating the selected games have been entered on our website along with a summary of the previous week's results and the year to date standings. The picks are to be submitted by no later then their scheduled kick-off time according to our website </p>
              <h3>Remember to confirm that your picks were successfully submitted by the website. When you submit them, it should send you to the &ldquo;Results&rdquo; page and show all of the picks that have been entered so far. If you get logged out of the website, it DID NOT save your picks and you will need to log back in and reenter them </h3>
              <h3>***NOTE: If we can not find a line for a chosen game, we will either determine a viable line or make it a pick&rsquo;em game.*** </h3>
              <h4>(This only occurs a hand full of times during a typical season.) </h4>
              <p><strong>Penalties:</strong> For those who fail to submit their picks by the designated kick-off time, you will receive a loss for each game missed. If your picks failed to submit or if your internet is not working, contact one of us with your picks! We do not mind giving a couple of breaks, but if it becomes a habit we may be forced to exclude you from the contest without a refund. </p>
            </div>
          </div>
        </div>
        <div title="Page 2">
          <div>
            <div>
              <p><strong>Weekly Winner:</strong> The weekly winner is the participant with the most wins for the given week. If a tiebreaker is needed, we will determine the weekly winner by the following tiebreakers. </p>
              <ol>
                <li>
                  <p>Lowest point differential of tiebreaker game. </p>
                </li>
                <li>
                  <p>Who correctly picked the winner of the tiebreaker game versus the spread. </p>
                </li>
                <li>
                  <p>Who correctly picked the winner of the tiebreaker game straight up (regardless of </p>
                  <p>spread). </p>
                </li>
                <li>
                  <p>Closest to the actual spread. </p>
                </li>
                <li>
                  <p>Most overall wins. </p>
                </li>
                <li>
                  <p>Most weekly wins. </p>
                </li>
              </ol>
              <p><strong>Regular Season Winner: </strong>The regular season winner is the participant with the most overall wins at the end of the regular season (regular season + conference championships). If a tiebreaker is needed, we will determine the regular season winner by the following tiebreakers: </p>
              <ol>
                <li>
                  <p>Most weekly winners. (1st &amp; 2nd Place are considered as weekly winners) </p>
                </li>
                <li>
                  <p>Of the weeks that we won the most total wins. </p>
                </li>
                <li>
                  <p>Lowest total tiebreaker point differential for the entire regular season. </p>
                </li>
              </ol>
              <p><strong>Bowl Season winner: </strong>The winner of the bowl season will receive a separate payout. The winner will be determined the same way the weekly winner is determined. </p>
              <p><strong>Tiebreaker Note:</strong> The point differential is the combined difference between each team&rsquo;s tiebreaker score and each team&rsquo;s actual game score. For example, if you chose a tiebreaker score of 24 to 17 and the final game score is actually 28 to 13, then the point differential is (28- 24) + (17-13) or 8 points. </p>
              <p><strong>Game Selections:</strong> Our selection of games each week is determined by which games we think are more interesting from a national standpoint and competitiveness. If you have any suggestions for games you would like to pick, we will take them into consideration as long as you notify us no later than Sunday night on the week of the game. </p>
              <p>So there you go, pretty simple rules for a pretty simple pool. Good luck and we look forward to hearing from you!!! </p>
              <p>Phone Numbers: </p>
              <p></p>
              <p>Andres - 407-405-4379 </p>
              <p>Warmest regards, </p>
              <p>Andres </p>
            </div>
          </div>
      </div>
     </div> <!-- end of the primary div in the container-->
</div>
<!-- #include file="includes/footer.asp" -->
</body>
</html>
<%	'**************************************************************************
	'* Local functions.                                                       *
	'**************************************************************************

	'--------------------------------------------------------------------------
	' Formats a number as a word, if appropriate.
	'--------------------------------------------------------------------------
	function FormatStrikes(n)

		dim text

		FormatStrikes = n
		text = Array("zero", "one", "two", "three", "four", "five", "six", _
		             "seven", "eight", "nine", "ten")
		if n < UBound(text) then
			FormatStrikes = text(n)
		end if

	end function %>