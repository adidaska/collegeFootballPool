//*****************************************************************************
// Site menu code.
//
// Note: Requires common.js.
//*****************************************************************************

// Used to track the currently active menu.
var activeMenu;

// Set up page-level event capturing.
if (isIE)
	document.documentElement.attachEvent("onmousedown", pageMousedown);
else
	document.documentElement.addEventListener("mousedown", pageMousedown, true);

//-----------------------------------------------------------------------------
// Close any currently active menu if the page is clicked on elsewhere.
//-----------------------------------------------------------------------------
function pageMousedown(event)
{
	// If there is no currently active menu, exit.
	if (activeMenu == null)
		return;

	// Find the element that triggered the event.
	var el = (isIE ? window.event.srcElement : (event.target.tagName ? event.target : event.target.parentNode));

	// If the triggering element is not part of the menu bar or a menu, close
	// the currently active menu.
	while (el != null)
	{
		if (el.id == "menuBar" || (el.className != null && hasClassName(el, "menu")))
			return;
		el = el.parentNode;
	}
	closeMenu();
}

//-----------------------------------------------------------------------------
// Opens the designated menu.
//-----------------------------------------------------------------------------
function openMenu(linkEl, id)
{
	// If the specified menu is the currently active one, exit.
	if (activeMenu != null && activeMenu.id == id)
		return;

	// For IE, set the opener link up for highlight/restore on focus/bur.
	if (isIE)
	{
		if (window.event.type == "focus")
		{
			// The first time the link receives focus, do initialization.
			if (linkEl.onblur == null)
				linkEl.onblur = menuItemBlur;
				
			// Highlight it.
			addClassName(linkEl, "ieFocus");
		}
	}
	
	// Close any currently active menu.
	closeMenu();

	// Get the named menu and initialize it, if not already done.
	var menuEl = document.getElementById(id);
	if (menuEl.isInitialized == null)
		initializeMenu(menuEl, linkEl);

	// Position the menu and make it visible.
	var pt = getPageOffset(linkEl);
	menuEl.style.left = (pt.x + (isSafari ? 1 : 0)) + "px";
	menuEl.style.top  = (pt.y + linkEl.parentNode.offsetHeight) + "px";
	menuEl.style.visibility = "visible";

	// If the menu has an underlying IFRAME (IE browsers), position, size and
	// display it.
	if (menuEl.iframeEl != null)
	{
		menuEl.iframeEl.style.left = menuEl.style.left;
		menuEl.iframeEl.style.top  = menuEl.style.top;
		menuEl.iframeEl.width  = menuEl.offsetWidth + "px";
		menuEl.iframeEl.height = menuEl.offsetHeight + "px";
		menuEl.iframeEl.style.display = "";
	}

	// Mark this menu as the active one.
	activeMenu = menuEl;

	return false;
}

//-----------------------------------------------------------------------------
// Closes the currently active menu.
//-----------------------------------------------------------------------------
function closeMenu()
{
	// Exit if there is no active menu.
	if (activeMenu == null)
		return;

	// Make the active menu invisible.
	activeMenu.style.visibility = "";

	// If the menu has an underlying IFRAME (IE browsers), hide it as well.
	if (activeMenu.iframeEl != null)
		activeMenu.iframeEl.style.display = "none";

	// Remove focus from the menu's opener link.
	activeMenu.openerLink.blur();

	// Clear the active menu.
	activeMenu = null;
}

//-----------------------------------------------------------------------------
// Initializes a menu.
//-----------------------------------------------------------------------------
function initializeMenu(menuEl, linkEl) {

	// Add a reference for the opener link to the menu.
	menuEl.openerLink = linkEl;

	// Handle special IE problems.
	if (isIE)
	{
		// Get all the menu item links.
		var linkEls = menuEl.getElementsByTagName("A");

		// Set up each menu link for highlight/restore on focus/blur.
		for (var i = 0; i < linkEls.length; i++)
		{
			linkEls[i].onblur = menuItemBlur;
			linkEls[i].onfocus = menuItemFocus;
		}

		// For pre-IE 7.
		if (window.XMLHttpRequest == null)
		{

			// Create an IFRAME element to place under the menu DIV. This will prevent
			// SELECT elements and other windowed controls from bleeding through.
			var iframeEl = document.createElement("IFRAME");
			iframeEl.frameBorder = 0;
			iframeEl.src = "javascript:false;document.write('');document.close();";
			iframeEl.style.display = "none";
			iframeEl.style.position = "absolute";
			iframeEl.style.filter = "progid:DXImageTransform.Microsoft.Alpha(style=0,opacity=0)";
			menuEl.iframeEl = menuEl.parentNode.insertBefore(iframeEl, menuEl);

			// Fix the hover problem by setting an explicit width on first item of
			// the menu.
			if (linkEls.length > 0)
			{
				var w = linkEls[0].offsetWidth;
				linkEls[0].style.width = w + "px";
				dw = linkEls[0].offsetWidth - w;
				w -= dw;
				linkEls[0].style.width = w + "px";
			}
		}
	}

	// Set event handlers for the menu and it's opener link.
	menuEl.onmouseout = menuMouseout;
	menuEl.onclick    = menuClick;
	linkEl.onmouseout = menuBarItemMouseout;

	// Mark the menu as initialized.
	menuEl.isInitialized = true;
}

//-----------------------------------------------------------------------------
// Handles a mouseout on a menu.
//-----------------------------------------------------------------------------
function menuMouseout(event)
{
	var current, related;

	try
	{
		if (window.event)
		{
			current = this;
			related = window.event.toElement;
		}
		else
		{
			current = event.currentTarget;
			related = event.relatedTarget;
		}

		// If the mouse has moved off the menu, close it.
		if (current != related && !contains(current, related))
			closeMenu();
	}
	catch (ex)
	{}
}

//-----------------------------------------------------------------------------
// Handles a click on a menu.
//-----------------------------------------------------------------------------
function menuClick(event)
{
	// Find the element that triggered the event.
	var el = (isIE ? window.event.srcElement : (event.target.tagName ? event.target : event.target.parentNode));
	
	// If it is not a link, exit.
	if (el.tagName != "A")
		return;

	// Close the menu.
	closeMenu();
}

//-----------------------------------------------------------------------------
// Handles a mouseout on a menu bar link.
//-----------------------------------------------------------------------------
function menuBarItemMouseout(event)
{
	var current, related;

	try
	{
		if (window.event)
		{
			current = this;
			related = window.event.toElement;
		}
		else
		{
			current = event.currentTarget;
			related = event.relatedTarget;
		}

		// If the mouse has moved off the menu bar link but not onto the menu, 
		// close it.
		if (!contains(current, related) && !contains(activeMenu, related))
		{
			closeMenu();
		}
	}
	catch (ex)
	{}
}

//-----------------------------------------------------------------------------
// Handles a blur on a menu link (needed for IE).
//-----------------------------------------------------------------------------
function menuItemBlur(event)
{
	// Unhighlight the link.
	removeClassName(window.event.srcElement, "ieFocus");
}

//-----------------------------------------------------------------------------
// Handles a focus on a menu link (needed for IE).
//-----------------------------------------------------------------------------
function menuItemFocus(event)
{
	// Highlight the link.
	addClassName(window.event.srcElement, "ieFocus");
}

//-----------------------------------------------------------------------------
// Determines if one node contains another.
//-----------------------------------------------------------------------------
function contains(nodeA, nodeB)
{
	// Return false if either node is null.
	if (nodeA == null || nodeB == null)
		return false;

	// Return true if nodes A and B are the same node.
	if (nodeA == nodeB)
		return true;

	// Return true if node B is a descendant of node A.
	while (nodeB.parentNode)
	{
		if ((nodeB = nodeB.parentNode) == nodeA)
			return true;
	}

	return false;
}

