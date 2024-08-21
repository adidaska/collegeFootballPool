/**
 * Created by adidaska on 7/29/14.
 */


$(document).ready(function() {
   $(".spreadUp").click(function(){
      $(".spread").toggle;
   });

   // Function to update the background color based on the selected radio button
   function updateSelection() {
      $('.trow').each(function() {
         // Find the selected radio button in the current row
         var selectedRadio = $(this).find('input[type="radio"]:checked');

         // If a radio button is selected, update the corresponding div's class
         if (selectedRadio.length > 0) {
            var parentDiv = selectedRadio.closest('.tleft, .tright');
            // Remove the class from all sibling team divs in the same row
            parentDiv.siblings('.tleft, .tright').removeClass('gameSelected');
            // Add the class to the selected div
            parentDiv.addClass('gameSelected');
         }
      });
   }

   // Run the updateSelection function on page load to set initial state
   updateSelection();

   // Handle click on the div containing the team logo (either tleft or tright)
   $('.tleft, .tright').click(function() {
      // Check if the div has the class 'locked'
      if ($(this).hasClass('locked')) {
         return; // Exit the function if the div is locked
      }

      // Find the radio button within the clicked div and select it
      $(this).find('input[type="radio"]').prop('checked', true);

      // Update the background color by toggling the class
      updateSelection();
   });

   $(document).ready(function() {
      var delay = 0; // Initial delay set to 0 ms
      var delayIncrement = 500; // Delay increment in milliseconds (1 second)

      // Function to fetch and update TV network info for a specific game
      function updateTVNetwork(gameId, delay) {
         setTimeout(function() {
            $.ajax({
               url: '/footballpool/serviceGetSpread.asp', // Adjust this to the actual path of your ASP proxy
               type: 'GET',
               data: { gameId: gameId },
               success: function(response) {
                  // Create a temporary DOM element to hold the HTML response
                  var tempDiv = $('<div>').html(response);

                  // Find the TV network using the given selector
                  var networkElement = tempDiv.find('.ScoreCell__Network.Gamestrip__Network.n9 > div');
                  var networkText = networkElement.text().trim(); // Get the network text

                  // Update the corresponding element on the page
                  if (networkText !== "") {
                     $('#game-Network-' + gameId).text(networkText);
                  } else {
                     $('#game-Network-' + gameId).text("Network info not available");
                  }
               },
               error: function(xhr, status, error) {
                  console.error('Error fetching TV network info for game ' + gameId + ':', error);
                  $('#game-Network-' + gameId).text("Error loading network info");
               }
            });
         }, delay);
      }

      // Traverse the site to find elements with the ID format "game-Network-" + espnGameId
      $('div[id^="game-Network-"]').each(function(index) {
         // Extract the espnGameId from the element's ID
         var id = $(this).attr('id');
         var gameId = id.replace('game-Network-', '');

         // Call the function to fetch and update the TV network information with delay
         updateTVNetwork(gameId, delay);

         // Increase the delay for the next request
         delay += delayIncrement;
      });
   });

});