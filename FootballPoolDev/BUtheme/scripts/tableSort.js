//*****************************************************************************
// Common table sorting code.
//
// Note: Requires common.js.
//*****************************************************************************

//-----------------------------------------------------------------------------
// Constants.
//-----------------------------------------------------------------------------

// Style class names.
var altRowClass  = "alt";
var sortColClass = "sortedCol";

// Regular expressions for special data formats.
var currencyRegex   = new RegExp("^(-?)\\$(\\d+.\\d{2})$");
var differenceRegex = new RegExp("^\\((\\d+)\\)$");
var percentRegex    = new RegExp("^\\((\\d+.?\\d*)\\)$");

// Regular expressions for normalizing white space.
var ldTrSpRegex = new RegExp("^\\s*|\\s*$", "g");
var multSpRegex = new RegExp("\\s\\s+", "g");

//-----------------------------------------------------------------------------
// Initializes and sets the sort direction for a table's columns. This allows
// a column to be given an initial sort direction and then be toggled.
//-----------------------------------------------------------------------------
function setSortDirection(tblEl, colName, rev, initColName)
{
	// The first time any column on a table is sorted, set up an array of flags
	// to track each column's current sort direction.
	if (tblEl.reverseSort == null)
	{
		tblEl.reverseSort = new Array();

		// Also, assume the table is initially sorted on the given column name.
		tblEl.lastColumn = initColName;
	}

	// The first time a given column is sorted, set its initial sort direction
	// as specified.
	if (tblEl.reverseSort[colName] == null)
		tblEl.reverseSort[colName] = rev;

	// If this column was the last one sorted on, reverse its sort direction.
	if (colName == tblEl.lastColumn)
		tblEl.reverseSort[colName] = !tblEl.reverseSort[colName];

	// Remember this column as being the last one sorted.
	tblEl.lastColumn = colName;
}

//-----------------------------------------------------------------------------
// Returns the text within a table cell.
//-----------------------------------------------------------------------------
function getTextValue(el)
{
	var i;
	var s;

	// Find and concatenate the values of all text nodes contained within the
	// element.
	s = "";
	for (i = 0; i < el.childNodes.length; i++)
	{
		if (el.childNodes[i].nodeType == document.TEXT_NODE)
			// Get the text, replacing any no-breaking spaces.
			s += el.childNodes[i].nodeValue.replace(/\u00a0/, " ");
		else
		{
			if (el.childNodes[i].nodeType == document.ELEMENT_NODE && el.childNodes[i].tagName == "BR")
				s += " ";
			else
				// Use recursion to get text within sub-elements.
				s += getTextValue(el.childNodes[i]);
		}
	}

	// Normalize any white space in the string.
	s = s.replace(multSpRegex, " ");  // Collapse any multiple white space.
	s = s.replace(ldTrSpRegex, "");   // Remove and leading and trailing white space.
	return s;
}

//-----------------------------------------------------------------------------
// Compares two text values based on their format.
//-----------------------------------------------------------------------------
function compareValues(v1, v2)
{
	var f1, f2;

	// Replace any 1/2 character with a decimal string.
	v1 = v1.replace(/\u00bd/, ".5");
	v2 = v2.replace(/\u00bd/, ".5");

	// If the values are in currency format (after removing commas), convert
	// them to numeric strings.
	var t1 = v1.replace(",", "");
	var t2 = v2.replace(",", "");
	if (currencyRegex.test(t1) && currencyRegex.test(t2))
	{
		v1 = t1.replace(currencyRegex, "$1$2");
		v2 = t2.replace(currencyRegex, "$1$2");
	}

	// If the values are in the format used for the tie breaker difference,
	// convert them to numeric strings.
	if (differenceRegex.test(v1) && differenceRegex.test(v2))
	{
		v1 = v1.replace(differenceRegex, "$1");
		v2 = v2.replace(differenceRegex, "$1");
	}

	// If the values are in percentage format, convert them to numeric strings.
	if (percentRegex.test(v1) && percentRegex.test(v2))
	{
		v1 = v1.replace(percentRegex, "$1");
		v2 = v2.replace(percentRegex, "$1");
	}

	// Always rank a numeric value higher than a non-numeric value.
	f1 = parseFloat(v1);
	f2 = parseFloat(v2);
	if (!isNaN(f1) && isNaN(f2))
		return 1;
	if (isNaN(f1) && !isNaN(f2))
		return -1;

	// If the values are numeric, convert them to floats.
	f1 = parseFloat(v1);
	f2 = parseFloat(v2);
	if (!isNaN(f1) && !isNaN(f2))
	{
		v1 = f1;
		v2 = f2;
	}

	// Compare the two values and return the result.
	if (v1 == v2)
		return 0;
	if (v1 > v2)
		return 1
	return -1;
}

//-----------------------------------------------------------------------------
// Resets styles on a table after sorting.
//-----------------------------------------------------------------------------
function restyleTable(tblEl, sortCols, hdrCols)
{
	var i, j, k;
	var rowEl, cellEl;

	// Restyle the table body.
	for (i = 0; i < tblEl.rows.length; i++)
	{
		// Set style classes on each row to alternate their appearance.
		rowEl = tblEl.rows[i];
		if (i % 2 != 0)
			addClassName(rowEl, altRowClass);
		else
			removeClassName(rowEl, altRowClass);

		// Set style classes on each column to highlight the ones that were
		// sorted.
		for (j = 1; j < tblEl.rows[i].cells.length; j++)
		{
			cellEl = rowEl.cells[j];

			// Remove any highlighting from a previous sort.
			removeClassName(cellEl, sortColClass);

			// Now highlight the specifed sort columns.
			for (k = 0; k < sortCols.length; k++)
				if (j == sortCols[k])
					addClassName(cellEl, sortColClass);
		}
	}

	// Find the table header row.
	var el = tblEl.parentNode.tHead;
	rowEl = el.rows[0];

	// Set style classes for the header columns.
	for (j = 1; j < rowEl.cells.length; j++)
	{
		cellEl = rowEl.cells[j];

		// Remove any highlighting from a previous sort.
		removeClassName(cellEl, sortColClass);

		// Now highlight the specifed header columns.
		for (k = 0; k < hdrCols.length; k++)
			if (j == hdrCols[k])
				addClassName(cellEl, sortColClass);
	}
}
