<%
'-------------------------------------------------------------------------------
' /bin/common.asp
' Common functions.
'-------------------------------------------------------------------------------


'-------------------------------------------------------------------------------
' Connect SQL Server database.
Sub ConnectSQLServer(ByRef a_oConn, ByVal a_strServerName, ByVal a_strCatalogName, ByVal a_strUserName, ByVal a_strPassword)
	'--- connect database
	Dim strCon
	Set a_oConn = Server.CreateObject("ADODB.Connection")
	strCon = "Provider=SQLOLEDB.1;Data Source=;SERVER=" & a_strServerName & ";Initial Catalog=" & a_strCatalogName & ";User ID=" & a_strUserName & ";Password=" & a_strPassword &" ;"
	On Error Resume Next
	a_oConn.Open strCon 
	CheckError strCon
End Sub


'-------------------------------------------------------------------------------
' Print comma list of items.
Sub PrintCommaList(ByRef a_oConn, ByRef a_strQuery)
	'--- run query
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(a_strQuery)
	CheckError a_strQuery
	'--- print items
	Do while (Not rs.eof)
		Response.Write(rs(0))
		rs.MoveNext
		If (Not rs.eof) Then
			Response.Write(", ")
		End If 
	Loop
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Sub


'-------------------------------------------------------------------------------
' Returns comma list of items.
Function QueryToItemList(ByRef a_oConn, ByVal a_strQuery, ByVal a_strSeperator)
	QueryToItemList = ""
	'--- run query
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(a_strQuery)
	CheckError a_strQuery
	'--- print items
	Do while (Not rs.eof)
		QueryToItemList = CommQueryToItemListaList & rs(0)
		rs.MoveNext
		If (Not rs.eof) Then
			QueryToItemList = QueryToItemList & a_strSeperator
		End If 
	Loop
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Function


'-------------------------------------------------------------------------------
' Print module short and full name with color and link.
Sub PrintModuleByIdWithColor(ByRef a_oConn, ByRef a_iModuleId, ByRef a_strField)
	'--- run query
	Dim rs
	Dim b_strQuery
	If a_strField > "" Then
		b_strQuery = "SELECT * FROM [RD_Modules] WHERE [" & a_strField & "]=" & a_iModuleId
	Else
		b_strQuery = "SELECT * FROM [RD_Modules] WHERE [Module_Id]=" & a_iModuleId
	End If
	On Error Resume Next
	Set rs = a_oConn.Execute(b_strQuery)
	If (Err <> 0) Then
		Response.Write ("<P>Error executing query on the Database.</P>")
		Response.Write ("<P>Error Message:<BR>" & Err.Description & "</P>")
		Response.Write ("<p>Query: '" & b_strQuery & "'</P>")
		Response.End 
	End If
	'--- print items
	If rs.eof Then
		Response.Write("Unknown")
	Else
		Response.Write("<span style=""background-color: " & rs("Color") & """>" & rs("Nick_Name") & "</span>")
		If rs("URL") > "" Then
			Response.Write(" (<a href=""" & rs("URL") & """>" & rs("Exec_Name") & "</a>)")
		Else
			Response.Write(" (" & rs("Exec_Name") & ")")
		End If
	End If
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Sub


'-------------------------------------------------------------------------------
Sub PrintTable(ByRef a_oConn, ByVal a_strQuery)
	'--- run query
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(a_strQuery)
	CheckError a_strQuery
	'--- print items
	Dim i
	If (Not rs.eof) Then
		Response.Write("<table border='0'><tr bgcolor='#800080'>" & Chr(10))
		For i = 0 to (rs.fields.count - 1)
			Response.Write("<td><b><font color='#ffffff'>" & rs(i).Name & "</font></b></td>" & Chr(10))
		Next
		'--- print items
		Do while (Not rs.eof)
			Response.Write("<tr bgcolor='silver'>")
			For i = 0 to (rs.fields.count - 1)
				If rs(i) > "" Then
					Response.Write("<td>" & Server.HTMLEncode(rs(i)) & "</td>")
				Else
					Response.Write("<td>&nbsp;</td>")
				End If
			Next
			Response.Write("</tr>")
			rs.MoveNext
		Loop
		Response.Write("</table>")
	End If
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Sub


'-------------------------------------------------------------------------------
' Print a table from a given recordset.
' a_rs - Given recordset.
' Remarks: Recordset will be rolled back at first, and returned when pointing to
'          the last item. It will also be left open.
Sub PrintTableByRecordset(ByRef a_rs)
	'--- roll back
	a_rs.MoveFirst
	'--- print items
	Dim i
	If (Not a_rs.eof) Then
		Response.Write("<TABLE border=0><tr bgcolor=""#800080"">")
		For i = 0 to (a_rs.fields.count - 1)
			Response.Write("<td><b><font color=""#ffffff"">" & a_rs(i).Name & "</font></b></td>")
		Next
		'--- print items
		Do while (Not a_rs.eof)
			Response.Write("<tr bgcolor=""silver"">")
			For i = 0 to (a_rs.fields.count - 1)
				If a_rs(i) > "" Then
					Response.Write("<td>" & Server.HTMLEncode(a_rs(i)) & "</td>")
				Else
					Response.Write("<td>&nbsp;</td>")
				End If
			Next
			Response.Write("</tr>")
			a_rs.MoveNext
		Loop
		Response.Write("</table>")
	End If
End Sub


'-------------------------------------------------------------------------------
' Print single record in a "Name: Value" format.
Sub PrintSingleRecord(ByRef a_oConn, ByVal a_strQuery)
	'--- run query
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(a_strQuery)
	CheckError a_strQuery
	'--- print all fields
	Dim i
	If (Not rs.eof) Then
		Response.Write("<table border=0>")
		For i = 0 to (rs.fields.count - 1)
			Response.Write("    <tr>" & Chr(10))
			Response.Write("        <td><b>" & rs(i).Name & ":</b>&nbsp;</td>" & Chr(10))
			Response.Write("        <td>" & Server.HTMLEncode(rs(i)) & "</td>" & Chr(10))
			Response.Write("    </tr>" & Chr(10))
		Next
		Response.Write("</table>")
	End If
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Sub


'-------------------------------------------------------------------------------
' Get module short name by module's BTF ID.
Function GetModuleShortNameById(ByRef a_oConn, ByRef a_iModuleId, ByRef a_strField)
	'--- run query
	Dim rs
	Dim b_strQuery
	If a_strField > "" Then
		b_strQuery = "SELECT * FROM [RD_Modules] WHERE [" & a_strField & "]=" & a_iModuleId
	Else
		b_strQuery = "SELECT * FROM [RD_Modules] WHERE [Module_Id]=" & a_iModuleId
	End If
	On Error Resume Next
	Set rs = a_oConn.Execute(b_strQuery)
	If (Err <> 0) Then
		Response.Write ("<P>Error executing query on the Database.</P>")
		Response.Write ("<P>Error Message:<BR>" & Err.Description & "</P>")
		Response.Write ("<p>Query: '" & b_strQuery & "'</P>")
		Response.End 
	End If
	'--- print items
	If rs.eof Then
		GetModuleShortNameById = ""
	Else
		GetModuleShortNameById = rs("Nick_Name")
	End If
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Function


'===============================================================================
' DB Project
'===============================================================================


'-------------------------------------------------------------------------------
' Connect BTF database for BTF, Database and other R&D tables.
' Return: Valid connection variable or termination with error message.
Sub ConnectBTF(ByRef a_oConn)
	'--- connect database
	Dim b_strCon
	Set a_oConn = Server.CreateObject("ADODB.Connection")
'	b_strCon = "Provider=SQLOLEDB.1;Data Source=;SERVER=server-2;Initial Catalog=BTF_dbSQL;User ID=btf;Password=;"
	b_strCon = "Provider=SQLOLEDB.1;Data Source=;SERVER=server-2;Initial Catalog=PassCall_Database;User ID=web_access;Password=passcall;"
	On Error Resume Next
	a_oConn.Open b_strCon 
	CheckError b_strCon
End Sub


'-------------------------------------------------------------------------------
' Print module list, based on a given query from an existing databse connection.
' Each module is made of a short name, a full name in paranthesis and a link on
' the full name that leads to the documentation page in the Intranet.
' a_oConn - Open connection to table.
' a_strQuery - Query that results in a special recordset. Its first item is the
'              short name, the second is the full name, and the third is the URL
'              to the documentation.
'temp - this sub doesn't work... (in dev)
Sub PrintModuleList(ByRef a_oConn, ByRef a_strQuery)
End Sub


'-------------------------------------------------------------------------------
' Print table of fields, ordered by field order, along with primary key right
' after that.
Sub DBPrintTableFields(ByRef a_oConn, ByVal a_iTableId, ByVal a_iVersion)
	PrintTable a_oConn, "SELECT [Field_Name] AS [Field],[Description],[Data_Type] AS [Type],[Size],[Mandatory]" &_
						" FROM [DB_Table_Field] WHERE [Table_Id]=" & a_iTableId & " AND [Version]=" & a_iVersion & " " &_
						"ORDER BY [Field_Order]"
	Response.Write("<b>Primary Key: </b>")
	PrintCommaList a_oConn, "SELECT [Field_Name] " &_
							"FROM [DB_Table_Field] WHERE [PK]=1 AND [Table_Id]=" & a_iTableId & " AND [Version]=" & a_iVersion
End Sub


'-------------------------------------------------------------------------------
' Print comma list of modules and their version for specific table and version.
' a_oConn - Open connection to BTF database.
Sub PrintModulesVersionsForTable(ByRef a_oConn, ByVal a_iTableId, ByVal a_iTableVersion)
	Dim rs
	Dim strQuery
	strQuery  = "SELECT * FROM [DB_Module_Table],[RD_Modules] WHERE [Table_Id]=" & a_iTableId & " AND [Table_Version]=" & a_iTableVersion & " AND [DB_Module_Table].[Module_Id]=[RD_Modules].[Module_Id] ORDER BY [Nick_Name]"
	On Error Resume Next
	Set rs = a_oConn.Execute(strQuery)
	CheckError strQuery
	Do while (Not rs.eof)
		Response.Write("<span style='background-color: " & rs("Color") & "'>" & _
						rs("Nick_Name") & "</span> " & rs("Module_Version"))
		rs.MoveNext
		If (Not rs.eof) Then
			Response.Write(", ")
		End If
	Loop
	rs.Close
	Set rs = Nothing
End Sub


'-------------------------------------------------------------------------------
' Return the latest available version of a specific table.
' a_oConn - Open connection to BTF database.
' Return: 0 if no version was found, or a positive number if versions table does
'         contain at least one version.
Function GetTableLatestVersion(ByRef a_oConn, ByVal a_iTableId)
	Dim rs
	Dim strQuery
	'--- get number of latest version
	strQuery = "SELECT Max(Version) AS MaxVersion FROM [DB_Table_Version] WHERE [Table_Id]=" & a_iTableId
	On Error Resume Next
	Set rs = a_oConn.Execute(strQuery)
	CheckError strQuery
	
	'--- save the latest version's number
	GetTableLatestVersion = rs("MaxVersion")
	If IsNull(GetTableLatestVersion) Then
		GetTableLatestVersion = 0
	Else
		GetTableLatestVersion = CInt(GetTableLatestVersion)	
	End If

	rs.Close
	Set rs = Nothing
End Function


'-------------------------------------------------------------------------------
' Print table properties according to version number.
Sub PrintTableVersionProperties(ByRef a_oConn, ByVal a_iTableId, ByVal a_iTableVersion)
	'--- get table's properties
	Dim strQuery
	strQuery = "SELECT * FROM [DB_Table_Version] WHERE [Table_Id]=" & a_iTableId & " AND [Version]=" & a_iTableVersion
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(strQuery)
	CheckError strQuery

	'--- if table wasn't found
	If Not(rs.eof) Then
		Response.Write("<tr>")
		Response.Write("    <td><b>Author: </b></td>")
		Response.Write("    <td>" & rs("Author") & "</td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
		Response.Write("    <td><b>Date: </b></td>")
		Response.Write("    <td>" & rs("Time") & "</td>")
		Response.Write("</tr>")
		Response.Write("<tr>")
		Response.Write("    <td><b>Description: </b></td>")
		Response.Write("    <td>" & rs("Description") & "</td>")
		Response.Write("</tr>")
	End If
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Sub


'-------------------------------------------------------------------------------
' Prints a help block with special background.
' a_strHelp - An HTML help string.
Sub PrintHelp(ByRef a_strHelp)
	If a_strHelp > "" Then
		Response.Write("<!-- Help -->" & vbNewLine)
		Response.Write("<table border='0' align='center' width='80%' cellpadding='0' cellspacing='0'>" & vbNewLine)
		Response.Write("    <tr height='5'><td></td></tr>" & vbNewLine)
		Response.Write("    <tr>" & vbNewLine)
		Response.Write("        <td>" & vbNewLine)
		Response.Write("            <table border='1' width='100%' cellpadding='2'>" & vbNewLine)
		Response.Write("                <tr>" & vbNewLine)
		Response.Write("                    <td bgcolor='#ffffc0'>")
		Response.Write(a_strHelp)
		Response.Write("                    </td>" & vbNewLine)
		Response.Write("                </tr>" & vbNewLine)
		Response.Write("            </table>" & vbNewLine)
		Response.Write("        </td>" & vbNewLine)
		Response.Write("    </tr>" & vbNewLine)
		Response.Write("    <tr height='5'><td></td></tr>" & vbNewLine)
		Response.Write("</table>" & vbNewLine)
		Response.Write("<!-- /Help -->" & vbNewLine & vbNewLine)
	End If
End Sub		
			

'===============================================================================
' General
'===============================================================================


'-------------------------------------------------------------------------------
' Check if error has happend. If it did happen, then it terminates after showing
' what the error was.
' a_strAction - Optional action string that will be displayed if error has 
'               happened.
Sub CheckError(ByRef a_strAction)
	If (Err <> 0) Then
		Response.Write("<P><b><FONT color=""red"">Error while accessing database.</FONT><BR>")
		Response.Write("Error Message:</b> " & Err.Description & "<BR>")
		If (a_strAction > "") Then
			Response.Write("<b>Action:</b><br><span dir=ltr><pre>" & a_strAction & "</span></pre><BR>")
		End If
		Response.End 
	End If
End Sub


'-------------------------------------------------------------------------------
' Print the given error message and terminate.
Sub TerminateWithMessage(ByRef a_strMessage)
	Response.Write("<P><b><FONT color=""red"">" & a_strMessage & "</FONT></B></P>")
	Response.End 
End Sub


'-------------------------------------------------------------------------------
' Get user name without the domain name.
Function GetUserName()
	'--- isolate user name from domain name
	Dim b_strUserName
	Dim i
	b_strUserName = Request("REMOTE_USER")
	i = Instr(1, b_strUserName, "\")
	If (i <> 0) Then
		GetUserName = Mid(b_strUserName, i + 1)
	Else
		GetUserName = b_strUserName
	End If
End Function


'-------------------------------------------------------------------------------
' Convert newlines into <BR>.
Sub PrintStrWithNewlines(ByRef a_strSource)
	Dim iCur
	Dim iPrev
	iPrev = 1
	Do While True
		iCur = Instr(iPrev, a_strSource, Chr(10))
		If (iCur = 0) Then
			Response.Write(Mid(a_strSource, iPrev))
			Exit Sub
		End If
		Response.Write(Mid(a_strSource, iPrev, iCur - iPrev))
		Response.Write("<BR>")
		iPrev = iCur + 1
	Loop
End Sub


'-------------------------------------------------------------------------------
' Print a section header.
Private Sub PrintSectionHeader(ByVal a_strSectionName)
	Response.Write("<p>" & Chr(10))
	Response.Write("    <u><b>" & a_strSectionName & "</b></u>" & Chr(10))
	Response.Write("</p>" & Chr(10))
End Sub


'-------------------------------------------------------------------------------
' Details in rows.
Private Sub PrintDetailsBlockBegin()
	Response.Write("	<table border=""0"">" & Chr(10))
End Sub


Private Sub PrintDetailsBlockEnd()
	Response.Write("	</table>" & Chr(10))
End Sub


Private Sub PrintDetailsRowBegin(ByVal a_strRowName)
	Response.Write("	    <tr>" & Chr(10))
	Response.Write("	        <td valign=""top""><b>&nbsp;&nbsp;" & a_strRowName & ":&nbsp;</b></td>" & Chr(10))
	Response.Write("	        <td>")
End Sub


Private Sub PrintDetailsRowEnd()
	Response.Write("</td>" & Chr(10))
	Response.Write("	    </tr>" & Chr(10))
End Sub


'-------------------------------------------------------------------------------
' Show current database information with a title, and maybe also allow to edit.
' a_strTitle - The title of the field that will appear in bold as "Title: ".
' a_strDBField - Name of database field that is used to show current value, and
'                it's also used as the input name.
' a_bEdit - True if wish to edit now, or False for view only.
' a_iSize - Size of text input, in case a_bEdit is True.
' a_strHint - Hint to show to the right of the text input in case a_bEdit is True.
Private Sub ShowAndEditText(ByVal a_strTitle, ByVal a_strDBField, ByVal a_bEdit, ByVal a_iSize, ByVal a_strHint)
	PrintDetailsRowBegin a_strTitle
	If a_bEdit Then
		'--- calculate box height
		If a_iSize > 50 Then
			a_iSize = Abs(a_iSize / 50) + 1
			Response.Write("<textarea cols='50' rows='" & a_iSize & "' name='" & a_strDBField & "'>" & rs(a_strDBField) & "</textarea>" & Chr(10))
		Else
			Response.Write("<input type='text' size='" & a_iSize & "' name='" & a_strDBField & "' value='" & rs(a_strDBField) & "'>" & Chr(10))
		End If 
		If a_strHint > "" Then
			Response.Write("(" & a_strHint & ")" & Chr(10))
		End If
	Else
		Response.Write(rs(a_strDBField) & Chr(10))
	End If
	PrintDetailsRowEnd
End Sub


'-------------------------------------------------------------------------------
' Show current database information with a title, and maybe also allow to edit.
' a_strTitle - The title of the field that will appear in bold as "Title: ".
' a_strDBField - Name of database field that is used to show current value, and
'                it's also used as the input name.
' a_strDefault - Default text to present in text input.
' a_iSize - Size of text input, in case a_bEdit is True.
' a_strHint - Hint to show to the right of the text input in case a_bEdit is True.
Private Sub InputEditText(ByVal a_strTitle, ByVal a_strFieldName, ByVal a_strDefault, ByVal a_iSize, ByVal a_strHint)
	PrintDetailsRowBegin a_strTitle
	Response.Write("<input type='text' size='" & a_iSize & "' name='" & a_strFieldName & "' value='" & a_strDefault & "'>" & Chr(10))
	If a_strHint > "" Then
		Response.Write("(" & a_strHint & ")" & Chr(10))
	End If
	PrintDetailsRowEnd
End Sub


'===============================================================================
' Date Time
'===============================================================================


'-------------------------------------------------------------------------------
' Get user name without the domain name.
Private Function DateToSqlServer(ByVal a_strDate)
	DateToSqlServer = Year(a_strDate) & "-" & Month(a_strDate) & "-" & Day(a_strDate)
	If Hour(a_strDate) <>  Null Then
		DateToSqlServer = DateToSqlServer & " " & Hour(a_strDate) & ":" & Minute(a_strDate) & ":" & Second(a_strDate)
	End If
End Function


'===============================================================================
' Drop Down List
'===============================================================================


'-------------------------------------------------------------------------------
' Print drop-down list from database column
' a_oConn - Open connection to table.
' a_strQuery - Query that results in special recordset. Its last item is the 
'              value, and the rest are displayed.
'              If the recordset has only one field, it's both the value and the 
'              displayed.
' a_strFirstOption - Text to present as first option in case it's not Null.
Sub PrintDropDownList(ByRef a_oConn, ByVal a_strQuery, ByVal a_strSelectName, ByVal a_strFirstOption)
	'--- get table's properties
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(a_strQuery)
	CheckError a_strQuery
	'--- if table wasn't found
	If rs.eof Then
		Response.Write ("<font color=""red"">(empty list)</font>")
	Else
		Response.Write(Chr(10) & "    <select size=""1"" name=""" & a_strSelectName & """>" & Chr(10))
		'--- add prefix if was specified
		If (Not IsNull(a_strFirstOption)) Then
			Response.Write("        <option value="""">" & a_strFirstOption & "</option>")
		End If
		'--- print items
		Do while (Not rs.eof)
			If (rs.fields.count = 1) Then
				Response.Write("        <option value=""" & rs(0) & """>" & rs(0) & "</option>")
			Else
				Dim i
				Response.Write("        <option value=""" & rs(rs.fields.count - 1) & """>")
				For i = 0 to (rs.fields.count - 2)
					Response.Write(rs(i) & " ")
				Next
				Response.Write("</option>" & Chr(10))
			End If
			rs.MoveNext
		Loop
		Response.Write("    </select>")
	End If
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Sub


'-------------------------------------------------------------------------------
' Print drop-down list from database column
' a_oConn - Open connection to table.
' a_strQuery - Query that results in special recordset. Its last item is the 
'              value, and the rest are displayed.
'              If the recordset has only one field, it's both the value and the 
'              displayed.
' a_strFirstOption - Height of list box.
Sub PrintListBoxMultiple(ByRef a_oConn, ByVal a_strQuery, ByVal a_strSelectName, ByVal a_iSize)
	'--- get table's properties
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(a_strQuery)
	CheckError a_strQuery
	'--- if table wasn't found
	If rs.eof Then
		Response.Write ("<font color=""red"">(empty list)</font>")
	Else
		Response.Write(Chr(10) & "    <select multiple size='" & a_iSize & "' name='" & a_strSelectName & "'>" & Chr(10))
		'--- print items
		Do while (Not rs.eof)
			If (rs.fields.count = 1) Then
				Response.Write("        <option value=""" & rs(0) & """>" & rs(0) & "</option>")
			Else
				Dim i
				Response.Write("        <option value=""" & rs(rs.fields.count - 1) & """>")
				For i = 0 to (rs.fields.count - 2)
					Response.Write(rs(i) & " ")
				Next
				Response.Write("</option>" & Chr(10))
			End If
			rs.MoveNext
		Loop
		Response.Write("    </select>")
	End If
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Sub


'-------------------------------------------------------------------------------
' Print drop-down list from database column
' a_oConn - Open connection to table.
' a_strQuery - Query that results in special recordset. Its last item is the 
'              value, and the rest are displayed.
'              If the recordset has only one field, it's both the value and the 
'              displayed.
' a_strFirstOption - Text to present as first option in case it's not Null.
' a_strSelected - When not empty, it represents the value of the selected item.
Sub PrintDropDown(ByRef a_oConn, ByVal a_strQuery, ByVal a_strSelectName, ByVal a_strFirstOption, ByVal a_strSelected)
	'--- get table's properties
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(a_strQuery)
	CheckError a_strQuery
	'--- if table wasn't found
	If rs.eof Then
		Response.Write ("<font color=""red"">(empty list)</font>")
	Else
		Response.Write(Chr(10) & "    <select size=""1"" name=""" & a_strSelectName & """>" & Chr(10))
		'--- add prefix if was specified
		If (Not IsNull(a_strFirstOption)) Then
			Response.Write("        <option value="""">" & a_strFirstOption & "</option>")
		End If
		'--- print items
		Dim strAttrSelected ' equals " selected" if current item is the selected one, or empty otherwise
		Do while (Not rs.eof)
			'--- determine if current item is selected
			If StrComp(a_strSelected, rs(rs.fields.count - 1), vbTextCompare) = 0 Then
				strAttrSelected = " selected"
			Else
				strAttrSelected = ""
			End If
			'--- prepare the current item
			If (rs.fields.count = 1) Then
				Response.Write("        <option" & strAttrSelected & " value=""" & rs(0) & """>" & rs(0) & "</option>")
			Else
				Dim i
				Response.Write("        <option" & strAttrSelected & " value=""" & rs(rs.fields.count - 1) & """>")
				For i = 0 to (rs.fields.count - 2)
					Response.Write(rs(i) & " ")
				Next
				Response.Write("</option>" & Chr(10))
			End If
			rs.MoveNext
		Loop
		Response.Write("    </select>")
	End If
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Sub


'===============================================================================
' Employees
'===============================================================================


'-------------------------------------------------------------------------------
' Print employee full name with email link.
Sub PrintEmployeeByIdWithEmail(ByRef a_oConn, ByRef a_iEmployeeId, ByRef a_strField)
	'--- run query
	Dim rs
	Dim b_strQuery
	If a_strField > "" Then
		b_strQuery = "SELECT * FROM [HR_Employees] WHERE [" & a_strField & "]=" & a_iEmployeeId
	Else
		b_strQuery = "SELECT * FROM [HR_Employees] WHERE [EmployeeId]=" & a_iEmployeeId
	End If
	On Error Resume Next
	Set rs = a_oConn.Execute(b_strQuery)
	If (Err <> 0) Then
		Response.Write ("<P>Error executing query on the Database.</P>")
		Response.Write ("<P>Error Message:<BR>" & Err.Description & "</P>")
		Response.Write ("<p>Query: '" & b_strQuery & "'</P>")
		Response.End 
	End If
	'--- print items
	If rs.eof Then
		Response.Write("Unknown")
	Else
'		Response.Write("<a href=""mailto:" & rs("EmailName") & """>" & rs("FirstName") & " " & rs("LastName") & "</a>")
		Response.Write("<a href=""/hr/bin/employee_details.asp?employee_id=" & rs("EmployeeId") & """>" & rs("FirstName") & " " & rs("LastName") & "</a>")
	End If 
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Sub


'-------------------------------------------------------------------------------
' Print employee details, by flags.
' a_iDetails -
'	  1 = Full name, with link to employee details.
'	  2 = Department and section name, as "department, section".
'	  3 = Title.
'     4 = Bold full name, with link to employee details, and non-bold title.
'     4 = Bold full name, with link to employee details, and non-bold section.
Sub PrintEmployeeDetails(ByRef a_oConn, ByVal a_iEmployeeId, ByVal a_iDetails)
	'--- run query
	Dim rs
	Dim strQuery
	If a_iDetails = 2 Or a_iDetails = 5 Then
		strQuery = "SELECT * FROM [HR_Employees],[HR_Sections],[HR_Departments] WHERE [HR_Sections].[Section_Id]=[HR_Employees].[Section_Id] AND [HR_Sections].[Department_Id]=[HR_Departments].[Department_Id] AND [EmployeeId]=" & a_iEmployeeId
	ElseIf a_iDetails = 3 Or a_iDetails = 4 Then
		strQuery = "SELECT * FROM [HR_Employees],[HR_Titles] WHERE [HR_Titles].[Title_Id]=[HR_Employees].[Title_Id] AND [EmployeeId]=" & a_iEmployeeId
	Else
		strQuery = "SELECT * FROM [HR_Employees] WHERE [EmployeeId]=" & a_iEmployeeId
	End If			
	On Error Resume Next
	Set rs = a_oConn.Execute(strQuery)
	CheckError strQuery
	'--- print items
	If rs.eof Then
		Response.Write("(Unknown)")
	Else
		If a_iDetails = 1 Then
			Response.Write("<a href=""/hr/bin/employee_details.asp?employee_id=" & rs("EmployeeId") & """>" & rs("FirstName") & " " & rs("LastName") & "</a>")
		ElseIf a_iDetails = 2 Then
			'--- check if it's someone who has a section (sub-department)
			If rs("Section_Name") > "" Then
				Response.Write(rs("Department_Name") & ", " & rs("Section_Name"))
			Else
				Response.Write(rs("Department_Name"))
			End If
		ElseIf a_iDetails = 3 Then
			Response.Write(rs("Title_Name"))
		ElseIf a_iDetails = 4 Then
			Response.Write("<a href=""/hr/bin/employee_details.asp?employee_id=" & rs("EmployeeId") & """><b>" & Trim(rs("FirstName")) & "&nbsp;" & Trim(rs("LastName")) & "</b>&nbsp;" & Trim(rs("Title_Name")) & "</a>")
		ElseIf a_iDetails = 5 Then
			Response.Write("<a href=""/hr/bin/employee_details.asp?employee_id=" & rs("EmployeeId") & """><b>" & Trim(rs("FirstName")) & "&nbsp;" & Trim(rs("LastName")) & "</b>&nbsp;" & Trim(rs("Department_Name")) & "&nbsp;" & Trim(rs("Section_Name")) & "</a>")
		Else
			Response.Write("(Unknown mode)")
		End If			
	End If 
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Sub


'-------------------------------------------------------------------------------
' Similar to PrintEmployeeDetails(), but it gets a query that is supposed to
' result with a list of employee-id as the first column.
Sub PrintEmployeeDetailsCommaList(ByRef a_oConn, ByVal a_strQuery, ByVal a_iDetails)
	'--- run query
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(a_strQuery)
	CheckError a_strQuery
	'--- print items
	If rs.eof Then
		Response.Write("(nobody)")
	Else
		Do while (Not rs.eof)
			PrintEmployeeDetails a_oConn, rs(0), a_iDetails
			rs.MoveNext
			If (Not rs.eof) Then
				Response.Write(", ")
			End If 
		Loop
	End If
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Sub
%>