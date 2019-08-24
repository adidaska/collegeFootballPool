<%	'**************************************************************************
	'* Common code for handling form displays and processing.                 *
	'**************************************************************************

	'**************************************************************************
	'* Global variables.                                                      *
	'**************************************************************************

	'Holds error messages for form fields.
	dim FormFieldErrors
	set FormFieldErrors = Server.CreateObject("Scripting.Dictionary")

	'--------------------------------------------------------------------------
	' Returns true if the 'Cancel' button was clicked.
	'--------------------------------------------------------------------------
	function CancelRequested()

		CancelRequested = false
		if Request.Form("submit") = "Cancel" then
			CancelRequested = true
		end if

	end function

	'--------------------------------------------------------------------------
	' Returns true if the specified field name exists in the posted form data.
	'--------------------------------------------------------------------------
	function FormFieldExists(fldName)

		dim name

		FormFieldExists = false
		for each name in Request.Form
			if name = fldName then
				FormFieldExists = true
				exit function
			end if
		next

	end function

	'--------------------------------------------------------------------------
	' Checks the posted form data for the specified field. If it exists, that
	' value is returned. Otherwise, it returns the given value.
	'--------------------------------------------------------------------------
	function GetFieldValue(fldName, fldValue)

		'If the 'Cancel' button was clicked, just use the given value.
		if CancelRequested() then
			GetFieldValue = fldValue
			exit function
		end if

		'If a value is found in the posted data, use it. Otherwise, use the
		'given value.
		if FormFieldExists(fldName) then
			GetFieldValue = Request.Form(fldName)
		else
			GetFieldValue = fldValue
		end if

	end function

	'--------------------------------------------------------------------------
	' Given class name list, appends a special class name if the specified
	' form field is in error.
	'--------------------------------------------------------------------------
	function FieldStyleClass(classNames, fldName)

		FieldStyleClass = classNames
		if FormFieldErrors.Exists(fldName) then
			if Len(FieldStyleClass) > 0 then
				FieldStyleClass = FieldStyleClass & " "
			end if
			FieldStyleClass = FieldStyleClass & "fieldError"
		end if

	end function

	'--------------------------------------------------------------------------
	' Displays an error message listing the errors found for all form fields.
	'--------------------------------------------------------------------------
	sub FormFieldErrorsMessage(msg)

		dim str, name, n

		str = vbTab & "<!-- Error messages. -->" & vbCrLf _
		    & vbTab & "<div class=""error"">" &  vbCrLf _
		    & String(2, vbTab) & "<p>" & msg & "</p>" &  vbCrLf _
		    & String(2, vbTab) & "<ul>" & vbCrLf
		for each name in FormFieldErrors
			if FormFieldErrors(name) <> "" then
				str = str & String(3, vbTab) & "<li>" & FormFieldErrors(name) & "</li>" & vbCrLf
			end if
		next
		str = str _
		    & String(2, vbTab) & "</ul>" & vbCrLf _
		    & vbTab & "</div>" & vbCrLf _
		    & vbTab & "<br />"
		Response.Write(str)

	end sub %>
