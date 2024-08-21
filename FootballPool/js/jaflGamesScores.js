/**
 * Created by adidaska on 8/15/24.
 */
function getScores(gameID, homeScoreElement, visitorScoreElement, delay) {
   setTimeout(function() {
      $.ajax({
         url: '/footballpool/serviceGetSpread.asp',
         type: 'GET',
         data: { gameID: gameID },
         success: function(response) {
            // Create a temporary DOM element to hold the HTML response
            var tempDiv = $('<div>').html(response);

            // Log the entire tempDiv to the console for debugging
            console.log(tempDiv.html());

            // Use jQuery selectors to find the scores in the response HTML
            var homeScore = tempDiv.find('#Gamestrip__Score').eq(1).text().trim(); // Get and trim home score
            var visitorScore = tempDiv.find('#Gamestrip__Score').eq(0).text().trim(); // Get and trim visitor score

            // Update the corresponding cells in the table on your page
            if (homeScore !== "") {
               $(homeScoreElement).text(homeScore);
            } else {
               $(homeScoreElement).text("No Score Found");
            }

            if (visitorScore !== "") {
               $(visitorScoreElement).text(visitorScore);
            } else {
               $(visitorScoreElement).text("No Score Found");
            }
         },
         error: function(xhr, status, error) {
            console.error('Error fetching the scores:', error);
            $(homeScoreElement).text('Error loading score');
            $(visitorScoreElement).text('Error loading score');
         }
      });
   }, delay);
}

$(document).ready(function() {
   var delay = 0; // Initial delay set to 0 ms
   var delayIncrement = 1000; // Delay increment in milliseconds (1 second)

   // Loop through each row in the table
   $('tr').each(function() {
      var homeScoreCell = $(this).find('[class^="homescore_"]');
      var visitorScoreCell = $(this).find('[class^="visitorscore_"]');

      if (homeScoreCell.length > 0) {
         // Extract the gameID from the class of the home score cell
         var gameID = homeScoreCell.attr('class').match(/homescore_(\d+)/)[1];

         if (gameID) {
            // Call the function to get the scores and update the relevant cells
            getScores(gameID, homeScoreCell, visitorScoreCell, delay);

            // Increase the delay for the next request
            delay += delayIncrement;
         }
      }
   });
});