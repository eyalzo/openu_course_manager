<%
'-------------------------------------------------------------------------------
' /bin/html_forms.asp
' HTML forms functions, like creating HTML combo-box, list-box, text-input etc.
'-------------------------------------------------------------------------------


'-------------------------------------------------------------------------------
' Create HTML code of list-box with multiple selection and pre-selected items,
' based on database query.
' a_oConn - Open connection to database.
' a_strQuery - Query that results in special recordset. Its last item is the 
'              value, and the rest are displayed.
'              If the recordset has only one field, it's both the value and the 
'              displayed.
'              If there is more than one column, and first column is nameless,
'              then this column is treated as 'selected' column with 1 for yes
'              or anything else otherwise.
' a_iSize - Height of list-box, in lines.
' For example:
'     SELECT
'         CASE
'             WHEN [EmployeeId] IN
'			  (   SELECT [Employee_Id]
'                 FROM [PRJ_Meetings_Employees]
'				  WHERE [Meeting_Id]=2
'		  THEN '1'
'		  ELSE '0'
'		  END
'	      ,[FirstName],[LastName],[EmployeeId]
'     FROM [HR_Employees]
Private Function HTML_Input_Select_Multiple_From_Query(ByRef a_oConn, ByVal a_strQuery, ByVal a_strSelectName, ByVal a_iSize)
	HTML_Input_Select_Multiple_From_Query = ""
	'--- get table's properties
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(a_strQuery)
	CheckError a_strQuery
	'--- if table wasn't found
	If (Not rs.eof) Then
		HTML_Input_Select_Multiple_From_Query = HTML_Input_Select_Multiple_From_Query & vbNewLine & "    <select multiple size='" & a_iSize & "' name='" & a_strSelectName & "' id='" & a_strSelectName & "'>" & vbNewLine
		'--- print items
		Do while (Not rs.eof)
			If (rs.fields.count = 1) Then
				HTML_Input_Select_Multiple_From_Query = HTML_Input_Select_Multiple_From_Query & "        <option value='" & rs(0) & "'>" & rs(0) & "</option>"
			Else
				Dim i
				If rs(0).Name > "" Then
					HTML_Input_Select_Multiple_From_Query = HTML_Input_Select_Multiple_From_Query & "        <option value='" & rs(rs.fields.count - 1) & "'>"
					For i = 0 to (rs.fields.count - 2)
						HTML_Input_Select_Multiple_From_Query = HTML_Input_Select_Multiple_From_Query & rs(i) & " "
					Next
				Else
					If rs(0) = "1" Then
						HTML_Input_Select_Multiple_From_Query = HTML_Input_Select_Multiple_From_Query & "        <option selected value='" & rs(rs.fields.count - 1) & "'>"
					Else
						HTML_Input_Select_Multiple_From_Query = HTML_Input_Select_Multiple_From_Query & "        <option value='" & rs(rs.fields.count - 1) & "'>"
					End If
					For i = 1 to (rs.fields.count - 2)
						HTML_Input_Select_Multiple_From_Query = HTML_Input_Select_Multiple_From_Query & rs(i) & " "
					Next
				End If
				HTML_Input_Select_Multiple_From_Query = HTML_Input_Select_Multiple_From_Query & "</option>" & vbNewLine
			End If
			rs.MoveNext
		Loop
		HTML_Input_Select_Multiple_From_Query = HTML_Input_Select_Multiple_From_Query & "    </select>"
	End If
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Function


'-------------------------------------------------------------------------------
' Create HTML code of combo-box with optional pre-selected item based on 
' database query.
' a_oConn - Open connection to database.
' a_strQuery - Query that results in special recordset. Its last item is the 
'              value, and the rest are displayed.
'              If the recordset has only one field, it's both the value and the 
'              displayed.
'              If there is more than one column, and first column is nameless,
'              then this column is treated as 'selected' column with 1 for yes
'              or anything else otherwise.
' For example:
'     SELECT
'         CASE
'             WHEN [EmployeeId] IN
'			  (   SELECT [Employee_Id]
'                 FROM [PRJ_Meetings_Employees]
'				  WHERE [Meeting_Id]=2
'		  THEN '1'
'		  ELSE '0'
'		  END
'	      ,[FirstName],[LastName],[EmployeeId]
'     FROM [HR_Employees]
Private Function HTML_Input_Select_From_Query(ByRef a_oConn, ByVal a_strQuery, ByVal a_strSelectName)
	HTML_Input_Select_From_Query = ""
	'--- get table's properties
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(a_strQuery)
	CheckError a_strQuery
	'--- if table wasn't found
	HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & vbNewLine & "    <select name='" & a_strSelectName & "' id='" & a_strSelectName & "'>" & vbNewLine
	If (Not rs.eof) Then
		'--- print items
		Do while (Not rs.eof)
			If (rs.fields.count = 1) Then
				HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & "        <option value='" & rs(0) & "'>" & rs(0) & "</option>"
			Else
				Dim i
				If rs(0).Name > "" Then
					HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & "        <option value='" & rs(rs.fields.count - 1) & "'>"
					For i = 0 to (rs.fields.count - 2)
						HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & rs(i) & " "
					Next
				Else
					If rs(0) = "1" Then
						HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & "        <option selected value='" & rs(rs.fields.count - 1) & "'>"
					Else
						HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & "        <option value='" & rs(rs.fields.count - 1) & "'>"
					End If
					For i = 1 to (rs.fields.count - 2)
						HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & rs(i) & " "
					Next
				End If
				HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & "</option>" & vbNewLine
			End If
			rs.MoveNext
		Loop
	End If
	HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & "    </select>"
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Function


'-------------------------------------------------------------------------------
' Create HTML code of combo-box with optional pre-selected item based on a given
' selected string.
' a_oConn - Open connection to database.
' a_strQuery - Query that results in special recordset. Its last item is the 
'              value, and the rest are displayed.
'              If the recordset has only one field, it's both the value and the 
'              displayed.
Private Function HTML_Input_Select_From_Query2(ByRef a_oConn, ByVal a_strQuery, ByVal a_strSelectName, ByVal a_strSelected)
	HTML_Input_Select_From_Query = ""
	'--- get table's properties
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(a_strQuery)
	CheckError a_strQuery
	'--- if table wasn't found
	If (Not rs.eof) Then
		HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & vbNewLine & "    <select name='" & a_strSelectName & "' id='" & a_strSelectName & "'>" & vbNewLine
		'--- print items
		Do while (Not rs.eof)
			If (rs.fields.count = 1) Then
				HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & "        <option value='" & rs(0) & "'>" & rs(0) & "</option>"
			Else
				Dim i
				'--- check if this is the selected item
				If (Not IsNull(a_strSelected)) And rs(rs.fields.count) = a_strSelected Then
					HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & "        <option selected value='" & rs(rs.fields.count - 1) & "'>"
				Else
					HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & "        <option value='" & rs(rs.fields.count - 1) & "'>"
				End If
				For i = 0 to (rs.fields.count - 2)
					HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & rs(i) & " "
				Next
				HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & "</option>" & vbNewLine
			End If
			rs.MoveNext
		Loop
		HTML_Input_Select_From_Query = HTML_Input_Select_From_Query & "    </select>"
	End If
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Function


'-------------------------------------------------------------------------------
' Set of functions for building a combo-box with name and selected item.
Private Function HTML_Input_Select(ByVal a_strSelectName, ByVal a_strSelectOptions)
	HTML_Input_Select =	"    <select name='" & a_strSelectName & "' id='" & a_strSelectName & "'>" & vbNewLine&_
						a_strSelectOptions&_
						"    </select>" & vbNewLine
End Function

Private Function HTML_Input_Select_Option(ByVal a_strText, ByVal a_strValue, ByVal a_strSelectedValue)
	If a_strValue = a_strSelectedValue Then
		HTML_Input_Select_Option = "        <option selected value='" & a_strValue & "'>" & a_strText & "</option>" & vbNewLine
	Else
		HTML_Input_Select_Option = "        <option value='" & a_strValue & "'>" & a_strText & "</option>" & vbNewLine
	End If
End Function


'-------------------------------------------------------------------------------
' Text input, both 'text' and 'textarea' (for large text).
Private Function HTML_Input_Text(ByVal a_strInputName, ByVal a_iSize, ByVal a_strDefault)
	If a_iSize > 100 Then
		a_iSize = Abs(a_iSize / 100) + 1
		HTML_Input_Text = "    <textarea cols='100' rows='" & a_iSize & "' name='" & a_strInputName & "' id='" & a_strInputName & "'>" & Trim(a_strDefault) & "</textarea>" & vbNewLine
	Else
		HTML_Input_Text = "    <input type='text' size='" & a_iSize & "' name='" & a_strInputName & "' value='" & Trim(a_strDefault) & "'>" & vbNewLine
	End If 
End Function


'-------------------------------------------------------------------------------
' Submit button.
' a_strButton - Will be displayed on button itself as its text, with leading and
'               trailing space for clarity.
Private Function HTML_Input_Button(ByVal a_strButton)
	HTML_Input_Button = "    <input name='button' type='submit' value='" & a_strButton & "'>" & vbNewLine
End Function


'-------------------------------------------------------------------------------
' Hidden field, with name and value.
Private Function HTML_Input_Hidden(ByVal a_strName, ByVal a_strValue)
	HTML_Input_Hidden = "    <input type='hidden' name='" & a_strName & "' id='" & a_strName & "' value='" & a_strValue & "'>" & vbNewLine
End Function


'-------------------------------------------------------------------------------
' Radio button, with name, value and optional checked.
Private Function HTML_Input_Radio(ByVal a_strName, ByVal a_strValue, ByVal a_bChecked)
	If a_bChecked Then
		HTML_Input_Radio = "    <input type='radio' name='" & a_strName & "' id='" & a_strName & "' value='" & a_strValue & "' checked>" & vbNewLine
	Else
		HTML_Input_Radio = "    <input type='radio' name='" & a_strName & "' id='" & a_strName & "' value='" & a_strValue & "'>" & vbNewLine
	End If
End Function


'-------------------------------------------------------------------------------
' HTML form.
' Its name is 'form1' always.
Private Function HTML_Form(ByVal a_strAction, ByVal a_strFormContent)
	HTML_Form = ""&_
		"<!-- Form -->" & vbNewLine&_
		"<table border='0'>" & vbNewLine&_
		"<form name='form1' method='post' action='" & a_strAction & "'>" & vbNewLine&_
		a_strFormContent&_
		"</form>" & vbNewLine&_
		"</table>" & vbNewLine&_
		"<!-- /Form -->" & vbNewLine & vbNewLine
End Function


'-------------------------------------------------------------------------------
' JavaScript code for setting forcus on input, by its name.
' Its name is 'form1' always.
Private Function HTML_Input_Set_Focus(ByVal a_strInputName)
	HTML_Input_Set_Focus = ""&_
		"<script language='JavaScript'>" & vbNewLine&_
		"<!--" & vbNewLine&_
		"    document.form1." & a_strInputName & ".focus();" & vbNewLine&_
		"// -->" & vbNewLine&_
		"</script>" & vbNewLine
End Function


'-------------------------------------------------------------------------------
' JavaScript code for auto submit on changeof input, by its name.
' Form's name is 'form1' always.
Private Function HTML_Input_Auto_Submit(ByVal a_strInputName)
	HTML_Input_Auto_Submit = ""&_
		"<script language='vbscript'>" & vbNewLine&_
		"<!--" & vbNewLine&_
		"    function " & a_strInputName & "_onchange()" & vbNewLine&_
		"		document.form1.submit()" & vbNewLine&_
		"    end function " & vbNewLine&_
		"// -->" & vbNewLine&_
		"</script>" & vbNewLine
End Function


'-------------------------------------------------------------------------------
' HTML and JavaScript code for date-picker, which is a text box for date with a
' pick-date button. The button opens a calendar dialog.
Private Function HTML_Input_Date_Picker(ByVal a_strInputName, ByVal a_strInitialDate)
	HTML_Input_Date_Picker = ""&_
		"<input id='" & a_strInputName & "' name='" & a_strInputName & "' value='" & a_strInitialDate & "' maxlength='11' size='11'>" & vbNewLine&_
		"    <a href onmouseover=""window.status='Date Picker';return true;"" onmouseout=""window.status='';return true;"">" & vbNewLine&_
		"        <img id='btn_" & a_strInputName & "' border='0' src='/bin/images/btnDatePicker1.gif' width='24' height='22'>" & vbNewLine&_
		"    </a>" & vbNewLine&_
		"    <script language='javascript' for='btn_" & a_strInputName & "' event='onclick'>" & vbNewLine&_
		"        window.showModalDialog('/bin/calendar.asp', document.getElementById('" & a_strInputName & "') , 'dialogLeft: ' + window.event.clientX + 'px; dialogTop: ' + window.event.clientY + 'px; resizable:no; status:no; toolbar:no; scroll:no;');" & vbNewLine&_
		"	     return false;" & vbNewLine&_
		"    </script>"
End Function
%>