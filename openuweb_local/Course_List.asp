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
	<title>רשימת קורסים</title>
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">

<!-- Navigation bar -->
<b>[</b><a href="default.asp">דף ראשי</a><b>]:&nbsp;רשימת קורסים</b><hr>
<%	
	'--- connect database
	Dim oConn
	Dim strQuery
	Database_Connect_Openu oConn

	'--- run query
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=course_details.asp?course_number='+CAST(C.[Course_Number] AS varchar)+'>'+CAST(C.[Course_Number] AS varchar)+'</a>' AS [מספר]" & vbNewLine&_
		"    ,CN.[Name] AS [שם]" & vbNewLine&_
		"    ,COUNT(C.[Course_Id]) AS [סמסטרים]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Courses] C" & vbNewLine&_
	    "       INNER JOIN [CoursesNames] CN" & vbNewLine&_
	    "           ON C.[Course_Number]=CN.[Course_Number]" & vbNewLine&_
	    "GROUP BY C.[Course_Number],CN.[Name]" & vbNewLine&_
	    "ORDER BY C.[Course_Number]"

	Response.Write(HTML_Table_From_Query(oConn, strQuery))

	oConn.close
	Set oConn = Nothing
%>

</body>
</html>