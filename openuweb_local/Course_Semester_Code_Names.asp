<%@ LANGUAGE = VBScript %>
<%
Option Explicit
Response.CacheControl = "no-cache"	
Response.AddHeader "Pragma", "no-cache" 
Response.ExpiresAbsolute=#Jan 01, 1980 00:00:00# 
Response.CharSet = "windows-1255"
%>

<!--#include file="bin\common.asp" -->	
<!--#include file="bin\database.asp" -->	
<!--#include file="bin\html_tables.asp" -->	
<!--#include file="bin\html_forms.asp" -->	
<!--#include file="bin\html_styles.asp" -->	

<html>

<head>
	<link href="Openu.css" rel="stylesheet" type="text/css">
	<title>פרטי קורס בסמסטר</title>
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">

<%
	'--- connect database
	Dim oConn
	Dim strQuery
	Database_Connect_Openu oConn

	'--- save request details for later use
	Dim iCourseId
	iCourseId = Request("course_id")

	'--- general details
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    CN.[Name]+' ('+CAST(C.[Course_Number] AS nvarchar)+')' AS [קורס]" & vbNewLine&_
		"    ,C.[Semester] AS [סמסטר]" & vbNewLine&_
		"    ,CONVERT(varchar(5),GETDATE(),108)+' '+CONVERT(varchar(10),GETDATE(),105) AS [תאריך עדכון]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Courses] C" & vbNewLine&_
	    "       LEFT OUTER JOIN [CoursesNames] CN" & vbNewLine&_
	    "           ON C.[Course_Number]=CN.[Course_Number]" & vbNewLine&_
	    "WHERE C.[Course_Id]=" & iCourseId & vbNewLine&_
	    "ORDER BY C.[Course_Number],C.[Semester]"
	Response.Write(HTML_Info_From_Query(oConn, strQuery, True))

	'---------------------------------------------------------------------------
	'--- student list with maman columns
	Dim iMamanCount
	strQuery = ""&_
		"SELECT COUNT(*)" & vbNewLine&_
	    "FROM [Mamans] M" & vbNewLine&_
	    "WHERE M.[Course_Id]=" & iCourseId
	iMamanCount = Database_Run_Query_Return_String(oConn, strQuery)
	'--- build additions to query
	Dim i
	Dim strSelect
	Dim strFromM
	Dim strFromSM
	For i = 1 To iMamanCount
		'--- SELECT
		strSelect = strSelect&_
			"    ,CASE" & vbNewLine&_
			"        WHEN DATEDIFF(dd,M" & (10 + i) & ".[Delivery_Date],SM" & (10 + i) & ".[Date_Received]) > 10 THEN '<font color=red>'+CONVERT(varchar(8),SM" & (10 + i) & ".[Date_Received],5)+'</font>'" & vbNewLine&_
			"        ELSE CONVERT(varchar(8),SM" & (10 + i) & ".[Date_Received],5)" & vbNewLine&_
			"        END [התקבל " & (10 + i) & "]" & vbNewLine&_
			"    ,CASE" & vbNewLine&_
			"        WHEN DATEDIFF(dd,M" & (10 + i) & ".[Delivery_Date],SM" & (10 + i) & ".[Date_Sent]) > 21 THEN '<font color=red>'+CONVERT(varchar(8),SM" & (10 + i) & ".[Date_Sent],5)+'</font>'" & vbNewLine&_
			"        ELSE CONVERT(varchar(8),SM" & (10 + i) & ".[Date_Sent],5)" & vbNewLine&_
			"        END [נשלח " & (10 + i) & "]" & vbNewLine&_
			"    ,CASE" & vbNewLine&_
			"        WHEN SM" & (10 + i) & ".[Cancel_Late]=1 OR SM" & (10 + i) & ".[Cancel_Copy]=1 THEN '<small>('+CAST(SM" & (10 + i) & ".[Grade] AS varchar)+')</small><b>&nbsp;0</b>'" & vbNewLine&_
			"        ELSE '<b>'+CAST(SM" & (10 + i) & ".[Grade] AS varchar)+'</b>'" & vbNewLine&_
			"        END [ציון " & (10 + i) & "]" & vbNewLine
		'--- FROM M
		strFromM = strFromM&_
			"        INNER JOIN [Mamans] M" & (10 + i) & vbNewLine&_
			"            ON M" & (10 + i) & ".[Course_Id]=" & iCourseId &  " AND M" & (10 + i) & ".[Maman_Number]=" & (10 + i) & vbNewLine
		'--- FROM SM
		strFromSM = strFromSM&_
			"        LEFT OUTER JOIN [StudentsMamans] SM" & (10 + i) & vbNewLine&_
			"            ON M" & (10 + i) & ".[Maman_Id]=SM" & (10 + i) & ".[Maman_Id] AND S.[Student_Id]=SM" & (10 + i) & ".[Student_Id]" & vbNewLine
	Next
	
	'--- same with code names only
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    S.[Code_name] AS [שם קוד]" & vbNewLine&_
		strSelect&_
	    "FROM" & vbNewLine&_
	    "    [Students] S" & vbNewLine&_
	    "        INNER JOIN [StudentsGroups] SG" & vbNewLine&_
	    "            ON S.[Student_Id]=SG.[Student_Id]" & vbNewLine&_
	    "        INNER JOIN CoursesGroups CG" & vbNewLine&_
	    "            ON SG.[Group_Id]=CG.[Group_Id]" & vbNewLine&_
	    strFromM&_
	    strFromSM&_
	    "WHERE CG.[Course_Id]=" & iCourseId & vbNewLine&_
	    "    AND S.[Code_Name] > ''" & vbNewLine&_
	    "ORDER BY S.[Code_Name]"
	Response.Write(HTML_Table_From_Query(oConn, strQuery))
	
	oConn.close
	Set oConn = Nothing
%>

</body>
</html>