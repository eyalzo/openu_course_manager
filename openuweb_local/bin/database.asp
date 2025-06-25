<%
'-------------------------------------------------------------------------------
' /bin/database.asp
' Common functions.
'-------------------------------------------------------------------------------


'-------------------------------------------------------------------------------
' Connect SQL Server database.
Sub Database_Connect_SQL_Server(ByRef a_oConn, ByVal a_strServerName, ByVal a_strCatalogName, ByVal a_strUserName, ByVal a_strPassword)
	'--- connect database
	Dim strCon
	Set a_oConn = Server.CreateObject("ADODB.Connection")
	strCon = "Provider=SQLOLEDB.1;Data Source=;SERVER=" & a_strServerName & ";Initial Catalog=" & a_strCatalogName & ";User ID=" & a_strUserName & ";Password=" & a_strPassword &" ;"
	On Error Resume Next
	a_oConn.Open strCon 
	'--- don't show connection string, because it includes username and password
	CheckError "Connecting '" & a_strServerName & "'."
End Sub


'-------------------------------------------------------------------------------
' Connect MS Access database.
Sub Database_Connect_MS_Access(ByRef a_oConn, ByVal a_strFileName)
	'--- connect database
	Dim strCon
	Set a_oConn = Server.CreateObject("ADODB.Connection")
	strCon = "DRIVER=Microsoft Access Driver (*.mdb); DBQ=" & a_strFileName & ";"
	On Error Resume Next
	a_oConn.Open strCon,"",""
	'--- don't show connection string, because it includes username and password
	CheckError "Connecting MS Access file '" & a_strFileName & "'."
End Sub


'-------------------------------------------------------------------------------
' Connect BTF database for BTF, Database and other R&D tables.
' Return: Valid connection variable or termination with error message.
Sub Database_Connect_Openu(ByRef a_oConn)
	'--- connect database
	Dim b_strCon
	Set a_oConn = Server.CreateObject("ADODB.Connection")
	b_strCon = "Provider=SQLOLEDB.1;Data Source=;SERVER=NT-EYAL;Initial Catalog=Openu;User ID=sa;Password=;"
	On Error Resume Next
	a_oConn.Open b_strCon 
	'--- don't show connection string, because it includes username and password
	CheckError "Connecting BTF Database."
End Sub


'-------------------------------------------------------------------------------
' Check if error has happend. If it did happen, then it terminates after showing
' what the error was.
' a_strAction - Optional action string that will be displayed if error has 
'               happened.
Sub Database_Check_Error(ByRef a_strAction)
	If (Err <> 0) Then
		Response.Write("<P><b><FONT color=""red"">Error while accessing database.</FONT><BR>")
		Response.Write("Error Message:</b> " & Err.Description & "<BR>")
		If (a_strAction > "") Then
			Response.Write("<b>Action:</b><br><pre>" & Server.HTMLEncode(a_strAction) & "</pre><BR>")
		End If
		Response.End 
	End If
End Sub


'-------------------------------------------------------------------------------
' Run a query on open database connection. Suits INSERT, UPDATE and DELETE.
Private Sub Database_Run_Query(ByRef a_oConn, ByVal a_strQuery)
	Dim rs
	On Error Resume Next
	Set rs = a_oConn.Execute(a_strQuery)
	Database_Check_Error a_strQuery
	'--- close connection
	rs.close
	Set rs = Nothing
End Sub


'-------------------------------------------------------------------------------
' Run a query on open database connection. Suits INSERT on table with Identity
' column.
' Return: 0 if error, or a number which is the identity of the new row.
Private Function Database_Run_Query_Return_Id(ByRef a_oConn, ByVal a_strQuery)
	Dim rs
	On Error Resume Next
	Set rs = a_oConn.Execute(a_strQuery)
	Database_Check_Error a_strQuery
	'--- get id
	On Error Resume Next
	Set rs = a_oConn.Execute("SELECT @@IDENTITY AS [id]")
	Database_Check_Error "SELECT @@IDENTITY AS [id]"
	If rs.eof Then
		Database_Run_Query_Return_Id = 0
	Else
		Database_Run_Query_Return_Id = CInt(rs("id"))
	End If
	'--- close connection
	rs.close
	Set rs = Nothing
End Function


'-------------------------------------------------------------------------------
' Run a query on open database connection, and return the string found in first
' row and column. Suits SELECT only.
' Return: Empty string if error, or the first cell if it was found.
Private Function Database_Run_Query_Return_String(ByRef a_oConn, ByVal a_strQuery)
	Dim rs
	On Error Resume Next
	Set rs = a_oConn.Execute(a_strQuery)
	Database_Check_Error a_strQuery
	'--- return the first cell
	If rs.eof Or IsNull(rs(0)) Then
		Database_Run_Query_Return_String = ""
	Else
		Database_Run_Query_Return_String = rs(0)
	End If
	'--- close connection
	rs.close
	Set rs = Nothing
End Function


'-------------------------------------------------------------------------------
' Check if current user has permissions to view or alter other all employees in
' a global category.
' a_strPermission - The string that reflect the category, as appears in databse
'                   as [Permission_Id].
' Return: True if this user has the reuired permission. 
Private Function Has_Permission(ByRef a_oConn, ByVal a_strPermission)
	Has_Permission = False
	'--- security check
	If a_strPermission = "" Then
		Exit Function
	End If
	
	Dim strQuery
	strQuery = ""&_
		"SELECT 'Yes' "&_
		"FROM [HR_Employees] E, "&_
		"    [HR_Permissions_Employees] PE "&_
		"WHERE E.[Network_Name] = '" & GetUserName() & "' "&_
		"    AND PE.[Employee_Id] = E.[EmployeeId] "&_
		"    AND (PE.[Permission_Id] = '" & a_strPermission & "' OR PE.[Permission_Id]='hr_admin') "&_
		"    AND PE.[Allow] = 1 "
'		"UNION "&_
'		"SELECT 'Yes' "&_
'		"FROM [HR_Permissions] P "&_
'		"WHERE P.[Permission_Id] = 'hr_new_employee' "&_
'		"    AND P.[Default] = 1"
	Dim rs
	On Error Resume Next
	Set rs = a_oConn.Execute(strQuery)
	Database_Check_Error "Has_Permission"
	'--- check result
	If Not rs.eof Then
		Has_Permission = True
	End If
	'--- close connection
	rs.close
	Set rs = Nothing
End Function


'-------------------------------------------------------------------------------
' Check if current user has permissions to view or alter all other employees in
' a global category, or maybe it's the user himself.
' a_strPermission - The string that reflect the category, as appears in databse
'                   as [Permission_Id].
' a_iEmployeeId - The id of the employee which his details are regarded here.
' Return: True if this user has the reuired permission, or it's himself. 
Private Function Has_Permission_Myself(ByRef a_oConn, ByVal a_strPermission, ByVal a_iEmployeeId)
	Has_Permission_Myself = False
	'--- security check
	If a_strPermission = "" Then
		Exit Function
	End If
	
	Dim strQuery
	strQuery = ""&_
		"SELECT 'Yes' "&_
		"FROM [HR_Employees] E, "&_
		"    [HR_Permissions_Employees] PE "&_
		"WHERE E.[Network_Name] = '" & GetUserName() & "' "&_
		"    AND PE.[Employee_Id] = E.[EmployeeId] "&_
		"    AND (PE.[Permission_Id] = '" & a_strPermission & "' OR PE.[Permission_Id]='hr_admin') "&_
		"    AND PE.[Allow] = 1 "&_
		"UNION "&_
		"SELECT 'Yes' "&_
		"FROM [HR_Employees] "&_
		"WHERE [Network_Name] = '" & GetUserName() & "' "&_
		"    AND [EmployeeId] = " & a_iEmployeeId & " "&_
		"UNION "&_
		"SELECT 'Yes' "&_
		"FROM [HR_Permissions] P "&_
		"WHERE P.[Permission_Id] = 'hr_new_employee' "&_
		"    AND P.[Default] = 1"
	Dim rs
	On Error Resume Next
	Set rs = a_oConn.Execute(strQuery)
	Database_Check_Error "Has_Permission"
	'--- check result
	If Not rs.eof Then
		Has_Permission_Myself = True
	End If
	'--- close connection
	rs.close
	Set rs = Nothing
End Function


'-------------------------------------------------------------------------------
' Check if current user has permissions to view or alter other all employees in
' a global category, and terminate with permission description if not.
' a_strPermission - The string that reflect the category, as appears in databse
'                   as [Permission_Id].
Private Sub Check_Permission(ByRef a_oConn, ByVal a_strPermission)
	If Not Has_Permission(a_oConn, a_strPermission) Then
		Dim strQuery
		strQuery = ""&_
			"SELECT [Description] "&_
			"FROM [HR_Permissions] P "&_
			"WHERE P.[Permission_Id] = '" & a_strPermission & "' "
		Dim rs
		On Error Resume Next
		Set rs = a_oConn.Execute(strQuery)
		Database_Check_Error "Check_Permission"
		'--- save description
		Dim strDescription
		If rs.eof Then
			strDescription = "Sorry " & GetUserName() & ", you don't have sufficient permissions (" & a_strPermission & ")"
		Else
			strDescription = "Sorry " & GetUserName() & ", you don't have permissions to " & rs("description")
		End If
		'--- close connection
		rs.close
		Set rs = Nothing
		TerminateWithMessage strDescription
	End If
End Sub


'-------------------------------------------------------------------------------
' Before storing in SQL Server, convert newlines into <BR> and double the 
' single-quotes.
Private Function String_To_SQL_Server(ByVal a_strSource)
	a_strSource = Replace(a_strSource, vbNewLine, "<BR>" & vbNewLine)
	String_To_SQL_Server = Replace(a_strSource, "'", "''")
End Function


'-------------------------------------------------------------------------------
' When fetching from SQL Server, convert the <BR> into newlines.
Private Function SQL_Server_To_String(ByVal a_strSource)
	SQL_Server_To_String = Replace(a_strSource, "<BR>" & vbNewLine, vbNewLine)
End Function
%>