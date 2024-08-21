/**
 * Created by adidaska on 8/15/24.
 */

function getSpread(gameID, targetDiv, delay) {
   setTimeout(function() {
      $.ajax({
         url: '/footballpool/serviceGetSpread.asp',
         type: 'GET',
         data: { gameID: gameID },
         success: function(response) {
            // Create a temporary DOM element to hold the HTML response
            var tempDiv = $('<div>').html(response);

            // Log the entire tempDiv to the console for debugging
            // console.log(tempDiv.html());

            // Use a jQuery selector to find the spread in the response HTML
            var spreadElement = tempDiv.find('#topOdd').eq(2);
            console.log(response);

            var spread = spreadElement.text().trim(); // Trim to remove any extra whitespace

            // Check if the spread element is found and the text is not blank
            if (spreadElement.length > 0 && spread !== "") {
               $(targetDiv).text(spread);
            } else {
               $(targetDiv).text("No Odds Found");
            }
         },
         error: function(xhr, status, error) {
            console.error('Error fetching the spread:', error);
            $(targetDiv).text('Error loading spread');
         }
      });
   }, delay);
}

$(document).ready(function() {
   var delay = 0; // Initial delay set to 0 ms
   var delayIncrement = 500; // Delay increment in milliseconds (1 second)

   // Traverse all divs with an ID that starts with "spread_"
   $('div[id^="spread_"]').each(function(index) {
      // Extract the gameID from the div's ID
      var gameID = $(this).attr('id').replace('spread_', '');

      // Call the function to get the spread and update the div's content
      getSpread(gameID, this, delay);

      // Increase the delay for the next request
      delay += delayIncrement;
   });
});