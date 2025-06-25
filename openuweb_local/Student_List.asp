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
	<title>רשימת סטודנטים</title>
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">

<table class="PageTitle_Student">
    <tr>
        <td class="PageTitle">רשימת סטודנטים</td>
    </tr>
</table>

<ul>
    <li><a href="default.asp">דף ראשי</a></li>
</ul>
<%
	'--- connect database
	Dim oConn
	Dim strQuery
	Database_Connect_Openu oConn

	'--- courses filter
	Dim strQueryCourses
	strQueryCourses = ""&_
		"SELECT '0', '- כל הקורסים -' AS [Course], '', ''" & vbNewLine&_
		"UNION" & vbNewLine&_
		"SELECT" & vbNewLine&_
		"    CASE" & vbNewLine&_
		"        WHEN CAST(CN.[Course_Number] as varchar)='" & Request("course_number") & "' THEN '1'" & vbNewLine&_
		"        ELSE '0'" & vbNewLine&_
		"        END" & vbNewLine&_
		"    ,CAST([Course_Number] as varchar)" & vbNewLine&_
		"    ,[Name]" & vbNewLine&_
		"    ,CAST([Course_Number] as varchar)" & vbNewLine&_
		"FROM [CoursesNames] CN" & vbNewLine&_
		"ORDER BY [Course]"

	'--- semester filter
	Dim strQuerySemesters
	strQuerySemesters = ""&_
		"SELECT '0', '- כל הסמסטרים -', '' AS [Semester]" & vbNewLine&_
		"UNION" & vbNewLine&_
		"SELECT DISTINCT" & vbNewLine&_
		"    CASE" & vbNewLine&_
		"        WHEN [Semester]='" & Request("semester") & "' THEN '1'" & vbNewLine&_
		"        ELSE '0'" & vbNewLine&_
		"        END" & vbNewLine&_
		"    ,[Semester]" & vbNewLine&_
		"    ,[Semester]" & vbNewLine&_
		"FROM [Courses] C" & vbNewLine&_
		"ORDER BY [Semester] "

	Response.Write(HTML_Form("Student_List.asp", ""&_
		HTML_Style_Info1("קורס", HTML_Input_Select_From_Query(oConn, strQueryCourses, "course_number"))&_
		HTML_Style_Info1("סמסטר", HTML_Input_Select_From_Query(oConn, strQuerySemesters, "semester")))&_
		HTML_Input_Auto_Submit("course_number") & HTML_Input_Auto_Submit("semester"))
		
	'--- run query
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=student_details.asp?student_id='+CAST(S.[Student_Id] AS varchar)+'>'+REPLICATE('0',9-LEN(CAST(S.[Student_Id] AS varchar)))+CAST(S.[Student_Id] AS varchar)+'</a>' AS [מס' סטודנט]" & vbNewLine&_
		"    ,S.[First]+' '+S.[Last] AS [שם]" & vbNewLine&_
		"    ,S.[Address] AS [כתובת]" & vbNewLine&_
		"    ,S.[City] AS [ישוב]" & vbNewLine&_
		"    ,S.[Phone_Mobile] AS [טלפון נייד]" & vbNewLine&_
		"    ,S.[Phone_Day] AS [טלפון יום]" & vbNewLine&_
		"    ,S.[Phone_Evening] AS [טלפון ערב]" & vbNewLine&_
		"    ,S.[Email] AS [דואל]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    Students S LEFT OUTER JOIN" & vbNewLine&_
	    "    (((CoursesNames CN" & vbNewLine&_
	    "        INNER JOIN Courses C ON CN.[Course_Number]=C.[Course_Number])" & vbNewLine&_
	    "    INNER JOIN CoursesGroups CG ON C.[Course_Id]=CG.[Course_Id])" & vbNewLine&_
	    "    INNER JOIN StudentsGroups SG ON CG.[Group_Id]=SG.[Group_Id]) ON S.[Student_Id]=SG.[Student_Id]" & vbNewLine&_
	    "WHERE 1>0" & vbNewLine
	If Request("course_number") > "" Then
		strQuery = strQuery & "    AND CN.[Course_Number]=" & Request("course_number") & vbNewLine
	End If
	If Request("semester") > "" Then
		strQuery = strQuery & "    AND C.[Semester]='" & Request("semester") & "'" & vbNewLine
	End If
	strQuery = strQuery&_
	    "ORDER BY S.[First], S.[Last]"

	Response.Write(HTML_Table_From_Query(oConn, strQuery))

	oConn.close
	Set oConn = Nothing
%>

</body>
</html>