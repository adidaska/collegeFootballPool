//*****************************************************************************
// Date/Time picker code.
//
// Note: Requires common.js.
//*****************************************************************************

// The form field(s) to be updated.
var targetDayField  = null;
var targetDateField = null;
var targetTimeField = null;

// The currently selected date.
var selectedDate = null;

//=============================================================================
// Calendar pop-up code.
//=============================================================================

//----------------------------------------------------------------------------
// Initializes the pop-up calendar.
//----------------------------------------------------------------------------
function openCalendar(formEl, dayFieldName, dateFieldName)
{
	// Make sure any open calendar or clock is closed.
	closeCalendar();
	closeClock();

	// Set the target form field and date.
	targetDayField  = formEl.elements[dayFieldName];
	targetDateField = formEl.elements[dateFieldName];
	selectedDate = new Date(targetDateField.value);
	if (isNaN(selectedDate.valueOf()))
		selectedDate = new Date();

	// Position and show the calendar.
	var el = document.getElementById("calendar");
	var pt = getPageOffset(targetDayField);
	pt.x += 8;
	pt.y += targetDayField.offsetHeight - 4;
	el.style.left = pt.x + "px";
	el.style.top  = pt.y + "px";
	el.style.display = "block";
	setCalendar();
	
	// Highlight the target form fields.
	addClassName(targetDayField, "activeDateTime");
	addClassName(targetDateField, "activeDateTime");

	return false;
}

//----------------------------------------------------------------------------
// Hides the pop-up calendar.
//----------------------------------------------------------------------------
function closeCalendar()
{
	var el = document.getElementById("calendar");
	el.style.display = "";

	// Remove highlighting form the target form fields.
	if (targetDayField != null)
		removeClassName(targetDayField, "activeDateTime");
	if (targetDateField != null)
		removeClassName(targetDateField, "activeDateTime");
}

//----------------------------------------------------------------------------
// Updates the calendar display to reflect the currently selected date.
//----------------------------------------------------------------------------
function setCalendar()
{

	var el, tableEl, rowEl, cellEl, linkEl;
	var tmpDate, tmpDate2;
	var i, j;

	// Update the month/year in the header.
	el = document.getElementById("calendarMonth").firstChild;
	el.nodeValue = selectedDate.getMonthName() + "\u00a0" + selectedDate.getFullYear();

	// Start with the first day of the month and go back as necessary to the
	// previous Sunday.
	tmpDate = new Date(Date.parse(selectedDate));
	tmpDate.setDate(1);
	while (tmpDate.getDay() != 0)
		tmpDate.addDays(-1);

	// Go through each calendar day cell in the table and update it.
	tableEl = document.getElementById("calendarTable");
	for (i = 2; i <= 7; i++)
	{
		rowEl = tableEl.rows[i];

		// Loop through a week.
		for (j = 0; j < rowEl.cells.length; j++)
		{
			var text = tmpDate.getDayName() + ", " + tmpDate.getMonthName() + " " + tmpDate.getDate() + ", " + tmpDate.getFullYear();
			cellEl = rowEl.cells[j];
			linkEl = cellEl.firstChild;

			// Set the date display.
			linkEl.date = new Date(Date.parse(tmpDate));
			linkEl.title = "Select " + tmpDate.getDayName() + ", " + tmpDate.getMonthName() + " " + tmpDate.getDate() + ", " + tmpDate.getFullYear() + ".";
			linkEl.firstChild.nodeValue = tmpDate.getDate();

			// Set style for dates outside the target month.
			if (tmpDate.getMonth() != selectedDate.getMonth())
				linkEl.className = "otherMonth";
			else
			linkEl.className = "";

			// Highlight the selected date.
			if (cellEl.oldClass == null)
				cellEl.oldClass = cellEl.className;
			if (Date.parse(tmpDate) == Date.parse(selectedDate))
				cellEl.className = cellEl.oldClass + " selected";
			else
				cellEl.className = cellEl.oldClass;

			// Go to the next day.
			tmpDate.addDays(1);
		}
	}
}

//----------------------------------------------------------------------------
// Event handlers for the calendar elements.
//----------------------------------------------------------------------------
function monthClick(n)
{
	// Advance the calendar month and update the display.
	selectedDate.addMonths(n);
	setCalendar();

	return false;
}

function yearClick(n)
{

	// Advance the calendar year and update the display.
	selectedDate.addYears(n);
	setCalendar();

	return false;
}

function dateClick(link)
{
	// Change the selected date and update the calendar.
	if (link.date != null)
	{
		selectedDate = new Date(Date.parse(link.date));
		setCalendar();
	}

	return false;
}

function okCalendarClick()
{
	// Format the selected date as "m/d/y" and set the target date field.
	var s = String(selectedDate.getMonth() + 1)
	       + "/"
		   + String(selectedDate.getDate())
		   + "/"
	       + String(selectedDate.getFullYear());
	targetDateField.value = s;

	// Set the target day field, if it exists.
	if (targetDayField != null)
		targetDayField.value = selectedDate.getDayName().substr(0, 3);

	// Close the calendar.
	closeCalendar();

	return false;
}

function cancelCalendarClick()
{
	// Close the calendar.
	closeCalendar();

	return false;
}

//=============================================================================
// Pop-up clock code.
//=============================================================================

//----------------------------------------------------------------------------
// Initializes the pop-up clock.
//----------------------------------------------------------------------------
function openClock(formEl, timeFieldName)
{
	// Make sure any open calendar or clock is closed.
	closeCalendar();
	closeClock();

	// Set the target form field.
	targetTimeField = formEl.elements[timeFieldName];

	// Position and show the clock.
	var el = document.getElementById("clock");
	var pt = getPageOffset(targetTimeField);
	pt.x += Math.round(targetTimeField.offsetWidth / 10);
	pt.y += Math.round(3 * targetTimeField.offsetHeight / 4);
	el.style.left = pt.x + "px";
	el.style.top  = pt.y + "px";
	el.style.display = "block";

	// Set the drop down selections based on the time in the target field.
	var list = targetTimeField.value.split(":");
	var hour, minute, meridiem;
	if (list.length >= 2)
	{
		hour = list[0];
		minute = Math.round(list[1] / 5) * 5;
		list = targetTimeField.value.split(" ");
		meridiem = list[1].toLowerCase();
	}
	else
	{
		// Use the current time.
		var timeNow = new Date()
		hour   = (timeNow.getHours()) % 12;
		minute = Math.round(timeNow.getMinutes() / 5) * 5;
		meridiem = timeNow.getHours() >= 12 ? "pm" : "am";
	}
	setClockField("clockHour", hour);
	setClockField("clockMinute", minute);
	setClockField("clockMeridiem", meridiem);

	// Highlight the target form field.
	addClassName(targetTimeField, "activeDateTime");

	return false;
}

//----------------------------------------------------------------------------
// Hides the pop-up clock.
//----------------------------------------------------------------------------
function closeClock()
{
	var el = document.getElementById("clock");
	el.style.display = "";

	// Remove highlighting form the target form field.
	if (targetTimeField != null)
		removeClassName(targetTimeField, "activeDateTime");
}

//----------------------------------------------------------------------------
// Helper function for setting the time on the pop-up clock.
//----------------------------------------------------------------------------
function setClockField(name, value)
{
	// Select the specified value in the named drop-down list.
	var formEl = document.getElementById("clockForm");
	var selectEl = formEl.elements[name];
	for (var i = 0; i < selectEl.options.length; i++)
	{
		if (selectEl.options[i].value == value)
			selectEl.options[i].selected = true;
		else
			selectEl.options[i].selected = false;
	}
}

//----------------------------------------------------------------------------
// Event handlers for the pop-up clock elements.
//----------------------------------------------------------------------------
function okClockClick()
{
	// Format the time as "hh/mm/ss tt".
	var formEl = document.getElementById("clockForm");
	var selectEl = formEl.elements["clockHour"];
	var hh = selectEl.options[selectEl.selectedIndex].value;
	selectEl = formEl.elements["clockMinute"];
	var mm = selectEl.options[selectEl.selectedIndex].value;
	selectEl = formEl.elements["clockMeridiem"];
	var tt = selectEl.options[selectEl.selectedIndex].value;
	var s = hh + ":" + mm + ":00 " + tt;

	// Set the target form field.
	targetTimeField.value = s;

	// Close the clock.
	closeClock();

	return false;
}

function cancelClockClick()
{
	// Close the clock.
	closeClock();

	return false;
}

//============================================================================
// This code extends the Date object with new properties and methods.
//============================================================================

// Properties
Date.prototype.monthNames = new Array("January", "February", "March", "April",
	"May", "June", "July", "August", "September", "October", "November",
	"December");
Date.prototype.dayNames = new Array("Sunday", "Monday", "Tuesday", "Wednesday",
	"Thursday", "Friday", "Saturday");
Date.prototype.savedDate = null;

// Methods
Date.prototype.getMonthName = dateGetMonthName;
Date.prototype.getDayName   = dateGetDayName;
Date.prototype.getDays      = dateGetDays;
Date.prototype.addDays      = dateAddDays;
Date.prototype.addMonths    = dateAddMonths;

//----------------------------------------------------------------------------
// Returns the name of the date's month.
//----------------------------------------------------------------------------
function dateGetMonthName()
{
	return this.monthNames[this.getMonth()];
}

//----------------------------------------------------------------------------
// Returns the name day the date falls on.
//----------------------------------------------------------------------------
function dateGetDayName()
{
	return this.dayNames[this.getDay()];
}

//----------------------------------------------------------------------------
// Returns the number of days in the date's month.
//----------------------------------------------------------------------------
function dateGetDays()
{
	var tmpDate, d, m;

	tmpDate = new Date(Date.parse(this));
	m = tmpDate.getMonth();
	d = 28;
	do
	{
		d++;
		tmpDate.setDate(d);
	} while (tmpDate.getMonth() == m);

	return d - 1;
}

//----------------------------------------------------------------------------
// Adds the specified number of days to the date.
//----------------------------------------------------------------------------
function dateAddDays(n)
{
	// Add the specified number of days.
	this.setDate(this.getDate() + n);

	// Reset the new day of month.
	this.savedDate = this.getDate();
}

//----------------------------------------------------------------------------
// Adds the specified number of months to the date, adjusting the day of the
// month if necessary.
//----------------------------------------------------------------------------
function dateAddMonths(n)
{
	// Save the day of month if not already set.
	if (this.savedDate == null)
		this.savedDate = this.getDate();

	// Set the day of month to the first to avoid rolling.
	this.setDate(1);

	// Add the specified number of months.
	this.setMonth(this.getMonth() + n);

	// Restore the saved day of month, if possible.
	this.setDate(Math.min(this.savedDate, this.getDays()));
}

