//*****************************************************************************
// Confidence points display code.
//
// Note: Requires common.js.
//*****************************************************************************

//-------------------------------------------------------------------------
// Function for displaying confidence point status.
//-------------------------------------------------------------------------
function confidencePointsUpdate()
{
	try
	{
		// Get the form.
		var formEl = document.getElementById("entryForm");

		// Build a array of booleans for each point value.
		var n = formEl.elements["games"].value
		var pointsList = new Array();

		// Set the flags.
		var inputEl, pts;
		for (var i = 0; i < n; i++)
			pointsList[i] = 0;
		for (i = 1; i <= n; i++)
		{
			if ((inputEl = formEl.elements["conf-" + i]) == null)
				inputEl = formEl.elements["lockedConf-" + i];
			if (inputEl != null)
			{
				pts = parseFloat(inputEl.value);
				if (!isNaN(pts) && pts == parseInt(pts) && pts >= 1 && pts <= n)
					pointsList[pts - 1] += 1;
			}
		}

		var cellEl = document.getElementById("pointsList");
		var spanEl;
		while(cellEl.firstChild != null)
			cellEl.removeChild(cellEl.firstChild);
		cellEl.appendChild(document.createTextNode("Confidence Points:"));
		for (i = 1; i <= n; i++)
		{
			spanEl = document.createElement("SPAN");
			spanEl.appendChild(document.createTextNode(i));
			if (pointsList[i - 1] == 0)
				spanEl.className = "availPts";
			if (pointsList[i - 1] == 1)
				spanEl.className = "usedPts";
			if (pointsList[i - 1] > 1)
				spanEl.className = "duplPts";
			cellEl.appendChild(document.createTextNode(" "));
			cellEl.appendChild(spanEl);
		}

		// Display the table row.
		if (cellEl.parentNode.style.display.length > 0)
			cellEl.parentNode.style.display = "";
	}
	catch (ex)
	{}
}

// Initialize the confidence points status display on page load.
window.addOnloadHandler(confidencePointsUpdate);
