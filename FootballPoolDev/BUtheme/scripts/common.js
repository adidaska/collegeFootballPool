//*****************************************************************************
// Commonly used code and functions.
//*****************************************************************************

// Check for specific browsers.
var isIE     = (document.all && window.innerWidth == null ? true : false);
var isOpera  = (window.opera ? true : false);
var isSafari = (navigator.userAgent.indexOf("Safari") >= 0 ? true : false);

// This code is necessary for browsers that don't reflect the DOM constants.
if (document.ELEMENT_NODE == null)
{
	document.ELEMENT_NODE = 1;
	document.TEXT_NODE    = 3;
}

//=============================================================================
// Code to add/remove style classes to elements.
//=============================================================================

//-----------------------------------------------------------------------------
// Returns true if the given element currently has the specified style class.
//-----------------------------------------------------------------------------
function hasClassName(el, name)
{
	var list = el.className.split(" ");
	for (var i = 0; i < list.length; i++)
		if (list[i] == name)
			return true;

	return false;
}

//-----------------------------------------------------------------------------
// Adds the specified class name to the given element.
//-----------------------------------------------------------------------------
function addClassName(el, name)
{
	if (!hasClassName(el, name))
		el.className += (el.className.length > 0 ? " " : "") + name;
}

//-----------------------------------------------------------------------------
// Removes the specified class name from the given element.
//-----------------------------------------------------------------------------
function removeClassName(el, name)
{
	if (el.className == null)
		return;

	var newList = new Array();
	var curList = el.className.split(" ");
	for (var i = 0; i < curList.length; i++)
		if (curList[i] != name)
			newList.push(curList[i]);
	el.className = newList.join(" ");
}

//=============================================================================
// Code for positioning elements.
//=============================================================================

//-----------------------------------------------------------------------------
// Returns the coordinates of the given element relative to the page.
//-----------------------------------------------------------------------------
function getPageOffset(el)
{
	// Find the page coordinates of the element.
	var x = 0, y = 0;

	// For IE, add up the border widths of any containing elements (except for
	// BODY, HTML and TABLE tags).
	if (isIE)
	{
		var tempEl = el;
		while (tempEl != null && tempEl.tagName != null) {
			if (tempEl.tagName != "BODY" && tempEl.tagName != "HTML" && tempEl.tagName != "TABLE")
			{
				x += tempEl.clientLeft;
				y += tempEl.clientTop;
			}
			tempEl = tempEl.parentNode;
		}
	}

	// Add up the left and top offsets of all positioned containing elements.
	do
	{
		x += el.offsetLeft;
		y += el.offsetTop;
		el = el.offsetParent;
	} while (el != null);

	// Return the coordinates.
	return new Point(x, y);
}

//-----------------------------------------------------------------------------
// Defines an object for holding x- and y- coordinates.
//-----------------------------------------------------------------------------
function Point(x, y)
{
	// Set the point coordinates.
	this.x = x;
	this.y = y;
}

//=============================================================================
// Window onload code, allows assignment of multiple handlers.
//
// Note: Where possible, will use the DOMContentLoaded event (or IE equivalent)
// instead of the window onload event to avoid waiting for external objects to
// load.
//=============================================================================

// Define a method for adding handlers.
window.addOnloadHandler = WindowAddOnloadHandler;

// Create an array for the event handlers.
window.onloadHandlers = new Array();

// Set the real onload handler.
//window.onload = WindowOnload;
if (document.addEventListener != null)
	document.addEventListener("DOMContentLoaded", WindowOnload, false);
else if (isIE)
{
	document.write("<script id=\"IEWindowOnload\" defer src=\"javascript:void(0)\"><\/script>");
	var script = document.getElementById("IEWindowOnload");
	script.onreadystatechange =
		function()
		{
			if (this.readyState == "complete")
				WindowOnload();
		}
}
else
	window.onload = WindowOnload;

//-----------------------------------------------------------------------------
// Adds a function to the array of window.onload handlers.
//-----------------------------------------------------------------------------
function WindowAddOnloadHandler(h)
{
	window.onloadHandlers[window.onloadHandlers.length] = h
}

//-----------------------------------------------------------------------------
// The master window.onload handler, calls each function in the array.
//-----------------------------------------------------------------------------
function WindowOnload(e)
{
	// Call every onload handler in the list.
	for (var i = 0; i < window.onloadHandlers.length; i++)
	{
		try
		{
			window.onloadHandlers[i](e);
		}
		catch (ex)
		{}
	}
}

//=============================================================================
// Code to update the current Eastern Time display.
//=============================================================================

var easternTimeEl  = null;
var easternTime    = null;
var localTime      = new Date();
var timeDifference = 0;

var DayNames   = new Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday");
var MonthNames = new Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December");

//-----------------------------------------------------------------------------
// Function for displaying the current ET.
//-----------------------------------------------------------------------------
function easternTimeUpdate()
{
	// Determine ET based on the client's current time.
	var currentTime = new Date();
	easternTime = new Date(currentTime.valueOf() + timeDifference);

	// Format the date and time.
	var meridiem = "am";
	var hrs = easternTime.getHours();
	if (hrs >= 12 && hrs <= 23)
	{
		meridiem = "pm";
	}
	if (hrs == 0)
	{
		hrs = 12;
	}
	if (hrs > 12)
	{
		hrs -= 12;
	}
	var mins = easternTime.getMinutes();
	if (mins < 10)
	{
		mins = "0" + mins;
	}
	var secs = easternTime.getSeconds();
	if (secs < 10)
	{
		secs = "0" + secs;
	}
	var s = DayNames[easternTime.getDay()] + ", " + MonthNames[easternTime.getMonth()]
		+ " " + easternTime.getDate() + ", " + easternTime.getFullYear()
		+ " " + hrs + ":" + mins + ":" + secs + " " + meridiem;

	// Display it.
	easternTimeEl.firstChild.nodeValue = s;
}

//-----------------------------------------------------------------------------
// Gets the current ET generated by the server and starts an interval timer to
// update the display periodically.
//-----------------------------------------------------------------------------
function easternTimeStart()
{
	// Get the string representing ET that was generated on the server.
	easternTimeEl = document.getElementById("easternTime");
	var el = document.getElementById("serverTimestamp");
	if (el != null)
	{
		// Create a Date object from that string and calculate the difference
		// between it and the client's current time.
		var s = el.firstChild.nodeValue;
		easternTime = new Date(s);
		timeDifference = easternTime.valueOf() - localTime.valueOf();

		// Start the display update timer.
		setInterval(easternTimeUpdate, 500);
	}
}

// Start the updates on page load.
window.addOnloadHandler(easternTimeStart);
