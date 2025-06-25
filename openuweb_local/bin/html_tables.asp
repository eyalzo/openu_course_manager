<%
'===============================================================================
' /bin/html_tables.asp
' Generate HTML tables from databse-queries, recordsets, etc.
'===============================================================================


'-------------------------------------------------------------------------------
' Run a query on database, and return results in HTML table format.
' First table row is a bold white-on-purple header, with column names.
' Data cells are black-on-silver.
' a_oConn - Open database connection, that will used in the function to create a
'           recordset.
' Return: String contains HTML table.
Private Function HTML_Table_From_Query(ByRef a_oConn, ByVal a_strQuery)
	HTML_Table_From_Query = ""
	'--- run query
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(a_strQuery)
	CheckError a_strQuery
	'--- print items
	Dim i
	If (Not rs.eof) Then
		HTML_Table_From_Query = "<table border='0'>" & vbNewLine & "    <tr bgcolor='#800080'>" & vbNewLine
		For i = 0 to (rs.fields.count - 1)
			If rs(i).Name > "" Then
				HTML_Table_From_Query = HTML_Table_From_Query & "        <td><b><font color='#ffffff'>" & rs(i).Name & "</font></b></td>" & vbNewLine
			Else
				HTML_Table_From_Query = HTML_Table_From_Query & "        <td bgcolor='white'>&nbsp;</td>" & vbNewLine
			End If
		Next
		'--- print items
		Do while (Not rs.eof)
			HTML_Table_From_Query = HTML_Table_From_Query & "    <tr bgcolor='silver'>" & vbNewLine
			For i = 0 to (rs.fields.count - 1)
				'--- cell color is white for columns with no name
				If rs(i).Name > "" Then
					HTML_Table_From_Query = HTML_Table_From_Query & "        <td>"
				Else
					HTML_Table_From_Query = HTML_Table_From_Query & "        <td bgcolor='white'>"
				End If
				'--- if there is no data, then leave cell empty
				If rs(i) > "" Then
					HTML_Table_From_Query = HTML_Table_From_Query & rs(i) & "</td>"
				Else
					HTML_Table_From_Query = HTML_Table_From_Query & "&nbsp;</td>"
				End If
			Next
			HTML_Table_From_Query = HTML_Table_From_Query & "</tr>" & vbNewLine
			rs.MoveNext
		Loop
		HTML_Table_From_Query = HTML_Table_From_Query & "</table>" & vbNewLine
	End If
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Function


'-------------------------------------------------------------------------------
' Run a query on database, and return results in HTML table format with item
' indentation, according to two extra columns in query.
' First table row is a bold white-on-purple header, with column names.
' Data cells are black-on-silver.
' First column is depth, zero-based, which is used for identation and numbering.
' Second column is a link to select item, and when it's empty, it means the item
' is already selected, and the row is colored white-on-blue.
' Third column is the text, which doesn't have a column name in the table.
' The selected item has an anchor names 'focus' attached to it, for easier 
' reference to it as #focus.
' a_oConn - Open database connection, that will used in the function to create a
'           recordset.
' Return: String contains HTML table.
Private Function HTML_Table_Indent_From_Query(ByRef a_oConn, ByVal a_strQuery)
	HTML_Table_Indent_From_Query = ""
	'--- run query
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(a_strQuery)
	CheckError a_strQuery
	'--- chec if recordset has at least 3 columns
	If rs.fields.count < 3 Then
		'--- close recordset
		rs.Close
		Set rs = Nothing
		'--- quit here
		Exit Function
	End If
	'--- hold an array to remember last serial for every depth, for auto numbering
	Dim iColSerial(1000)
	Dim iDepth
	Dim iPrevDepth
	Dim j
	Dim strSection
	iDepth = -1
	'--- print items
	Dim i
	If (Not rs.eof) Then
		HTML_Table_Indent_From_Query = "<table id='oTable' cellpadding='0' border='0'>" & vbNewLine
		'--- add header, meaning column names in white-on-purple
		HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query & "    <tr>" & vbNewLine
		Dim bHasHeader ' remember if there is no header to any column, so the row can be canceled
		bHasHeader = False
		For i = 2 to (rs.fields.count - 1)
			If (i > 2) And (rs(i).Name > "") Then
				bHasHeader = True
				HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query & "        <td bgcolor='#800080'><b><font color='#ffffff'>" & rs(i).Name & "</font></b></td>" & vbNewLine
			Else
				HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query & "        <td>&nbsp;</td>" & vbNewLine
			End If
		Next
		If bHasHeader Then
			HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query & "    </tr>" & vbNewLine
		Else
			HTML_Table_Indent_From_Query = "<table id='oTable' cellpadding='0' border='0'>" & vbNewLine
		End If
		'--- add items from all rows
		Do while (Not rs.eof)
			'-------------------------------------------------------------------
			'--- prepare automatic section number
			iPrevDepth = iDepth
			iDepth = rs(0)
			'--- reset counters
			If iDepth > iPrevDepth Then
				iColSerial(iDepth) = 0
			End If
			iColSerial(iDepth) = iColSerial(iDepth) + 1
			strSection = iColSerial(0)
			For j = 1 to (iDepth)
				strSection = strSection & "." & iColSerial(j)
			Next
			'--- print row
			HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query&_
				"    <tr id='text" & strSection & "' bgcolor='silver'>" & vbNewLine
			'--- print the special text column first
			HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query &_
				"        <td>" & vbNewLine&_
				"            <table width='100%' cellpadding='0' cellspacing='0' border='0'>" & vbNewLine
			If rs(1) > "" Then
				HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query &_
					"                <tr>" & vbNewLine&_
					"				     <td bgcolor='white' align='right' valign='top' width='" & 36+35*CInt(rs(0)) & "'>" & vbNewLine&_
					"                        <img id='open' border='0' src='http://" & Request.ServerVariables("server_name") & "/bin/images/open1.gif' onclick='HideRows(""text" & strSection & """)'>" & vbNewLine&_
					"                        <a href='" & rs(1) & "'><b><sub>" & strSection & "</sub></b></a>&nbsp;" & vbNewLine&_
					"                    </td>" & vbNewLine&_
					"                    <td>" & vbNewLine&_
					"                        " & rs(2) & vbNewLine
			Else
				HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query &_
					"                <a name='focus'></a>" & vbNewLine&_
					"                <tr bgcolor='#00007F'>" & vbNewLine&_
					"				     <td bgcolor='white' align='right' valign='top' width='" & 36+35*CInt(rs(0)) & "'>" & vbNewLine&_
					"                        <img border='0' src='http://" & Request.ServerVariables("server_name") & "/bin/images/open1.gif' onclick='HideRows(""text" & strSection & """)'>" & vbNewLine&_
					"                        <b><font color='#0000FF'><sub>" & strSection & "</sub></font></b>&nbsp;" & vbNewLine&_
					"                    </td>" & vbNewLine&_
					"                    <td>" & vbNewLine&_
					"                        <font color='white'>" & rs(2) & "</font>" & vbNewLine
			End If
			HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query &_
				"                    </td>" & vbNewLine&_
				"                </tr>" & vbNewLine&_
				"            </table> " & vbNewLine&_
				"        </td>" & vbNewLine
			
			'--- now print the other columns
			For i = 3 to (rs.fields.count - 1)
				'--- cell color is white for columns with no name
				If rs(i).Name > "" Then
					HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query & "        <td>"
				Else
					HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query & "        <td bgcolor='white'>"
				End If
				'--- if there is no data, then leave cell empty
				If rs(i) > "" Then
					HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query & rs(i) & vbNewLine & "        </td>" & vbNewLine
				Else
					HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query & "&nbsp;" & vbNewLine & "        </td>" & vbNewLine
				End If
			Next
			HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query&_
				"    </tr>" & vbNewLine
			rs.MoveNext
		Loop
		HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query&_
			"</table>" & vbNewLine
	End If
	'--- close recordset
	rs.Close
	Set rs = Nothing

	HTML_Table_Indent_From_Query = HTML_Table_Indent_From_Query&_
		"<script language='JavaScript'>" & vbNewLine&_
		"function HideRows(ParentRow)" & vbNewLine&_
		"{" & vbNewLine&_
		"    var curr_row;" & vbNewLine&_
		"    var i;" & vbNewLine&_
		"    i = ParentRow.length;" & vbNewLine&_
		"    //--- determine if it was open or closed, and change image " & vbNewLine&_
		"    var bOpen = true; " & vbNewLine&_
		"    var oImg = oTable.rows[ParentRow].getElementsByTagName('img');" & vbNewLine&_
		"    if(oImg.length > 0) " & vbNewLine&_
		"    { " & vbNewLine&_
		"        if(oImg[0].id == 'open') " & vbNewLine&_
		"        { " & vbNewLine&_
		"            bOpen = false; " & vbNewLine&_
		"            oImg[0].id = 'closed'; " & vbNewLine&_
		"            oImg[0].src = 'http://" & Request.ServerVariables("server_name") & "/bin/images/closed1.gif'; " & vbNewLine&_
		"        } " & vbNewLine&_
		"        else " & vbNewLine&_
		"        { " & vbNewLine&_
		"            oImg[0].id = 'open'; " & vbNewLine&_
		"            oImg[0].src = 'http://" & Request.ServerVariables("server_name") & "/bin/images/open1.gif'; " & vbNewLine&_
		"        } " & vbNewLine&_
		"    } " & vbNewLine&_
		"    for (curr_row = oTable.rows[ParentRow].rowIndex+1 ; curr_row < oTable.rows.length; curr_row++)" & vbNewLine&_
		"    {" & vbNewLine&_
		"        if(oTable.rows[curr_row].id.length > (i+1)) " & vbNewLine&_
		"            if(bOpen) " & vbNewLine&_
		"				 oTable.rows[curr_row].style.display=''; " & vbNewLine&_
		"            else " & vbNewLine&_
		"                oTable.rows[curr_row].style.display='none'; " & vbNewLine&_
		"        else " & vbNewLine&_
		"            break; " & vbNewLine&_
		"    }" & vbNewLine&_
		"}" & vbNewLine&_
		"</script>"
End Function


'-------------------------------------------------------------------------------
' Returns comma list of items.
Function HTML_List_From_Query(ByRef a_oConn, ByVal a_strQuery, ByVal a_strSeperator)
	HTML_List_From_Query = ""
	'--- run query
	On Error Resume Next
	Dim rs
	Set rs = a_oConn.Execute(a_strQuery)
	CheckError a_strQuery
	If rs.fields.count > 0 Then
		'--- print items
		Do while (Not rs.eof)
			HTML_List_From_Query = HTML_List_From_Query & rs(0)
			rs.MoveNext
			If (Not rs.eof) Then
				HTML_List_From_Query = HTML_List_From_Query & a_strSeperator
			End If 
		Loop
	End If
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Function


'-------------------------------------------------------------------------------
' Run a query on database, and return results in HTML info format, similar tp 
' what is made by calls to HTML_Style_Info1() from html_styles.asp.
' It prints only fields that has names.
' a_oConn - Open database connection, that will used in the function to create a
'           recordset.
' Return: String contains HTML info lines, or nothing if didn't find even one 
'         record and a_bReturnNothingIfEmpty is TRUE.
Private Function HTML_Info_From_Query(ByRef a_oConn, ByVal a_strQuery, ByVal a_bReturnNothingIfEmpty)
	HTML_Info_From_Query = ""
	'--- run query
	Dim rs
	On Error Resume Next
	Set rs = a_oConn.Execute(a_strQuery)
	Database_Check_Error a_strQuery
	'--- check if at least one record was found
	Dim i
	If rs.eof Then
		If a_bReturnNothingIfEmpty Then
			Exit Function
		Else
			'--- print items
			For i = 0 to (rs.fields.count - 1)
				If rs(i).Name > "" Then
					HTML_Info_From_Query = HTML_Info_From_Query&_
						"    <!-- HTML_Style_Info1(" & rs(i).Namer & ") -->" & vbNewLine&_
						"	 <table border='0'>" & vbNewLine&_
						"	     <tr>" & vbNewLine&_
						"	         <td valign='top'>" & vbNewLine&_
						"                <b>&nbsp;&nbsp;" & rs(i).Name & ":&nbsp;</b>" & vbNewLine&_
						"	         </td>" & vbNewLine&_
						"	     </tr>" & vbNewLine&_
						"	 </table>" & vbNewLine&_
						"    <!-- /HTML_Style_Info1(" & rs(i).Name & ") -->" & vbNewLine
				End If
			Next
		End If
	End If
	'--- print items
	For i = 0 to (rs.fields.count - 1)
		If rs(i).Name > "" Then
			HTML_Info_From_Query = HTML_Info_From_Query&_
				"    <!-- HTML_Style_Info1(" & rs(i).Name & ") -->" & vbNewLine&_
				"	 <table border='0'>" & vbNewLine&_
				"	     <tr>" & vbNewLine&_
				"	         <td valign='top'>" & vbNewLine&_
				"                <b>&nbsp;&nbsp;" & rs(i).Name & ":&nbsp;</b></td>" & vbNewLine&_
				"	         <td>" & vbNewLine&_
				"                " & rs(i) & vbNewLine&_
				"	         </td>" & vbNewLine&_
				"	     </tr>" & vbNewLine&_
				"	 </table>" & vbNewLine&_
				"    <!-- /HTML_Style_Info1(" & rs(i).Name & ") -->" & vbNewLine
		End If
	Next
	'--- close recordset
	rs.Close
	Set rs = Nothing
End Function
%>