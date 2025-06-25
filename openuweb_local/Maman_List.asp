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
	<title>רשימת מטלות</title>
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">

<table class="PageTitle_Maman">
    <tr>
        <td class="PageTitle">רשימת מטלות</td>
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

	Response.Write(HTML_Form("Maman_List.asp", ""&_
		HTML_Style_Info1("קורס", HTML_Input_Select_From_Query(oConn, strQueryCourses, "course_number"))&_
		HTML_Style_Info1("סמסטר", HTML_Input_Select_From_Query(oConn, strQuerySemesters, "semester")))&_
		HTML_Input_Auto_Submit("course_number") & HTML_Input_Auto_Submit("semester"))
		
	'--- run query
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    CAST(C.[Course_Number] AS nvarchar)+' '+CN.[Name] AS [קורס]" & vbNewLine&_
		"    ,C.[Semester] AS [סמסטר]" & vbNewLine&_
		"    ,'<a href=maman_details.asp?maman_id='+CAST(M.[Maman_Id] AS varchar)+'>'+CAST(M.[Maman_Number] AS varchar)+'</a>' AS [ממ""ן]" & vbNewLine&_
		"    ,COUNT(SM.[Student_Id]) AS [הגשות]" & vbNewLine&_
		"    ,ROUND(AVG(SM.[Grade]),1) AS [ממוצע]" & vbNewLine&_
		"    ,ROUND(STDEV(SM.[Grade]),1) AS [סטיית תקן]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Mamans] M" & vbNewLine&_
		"        LEFT OUTER JOIN [StudentsMamans] SM" & vbNewLine&_
		"            ON M.[Maman_Id]=SM.[Maman_Id]" & vbNewLine&_
	    "    ,[Courses] C" & vbNewLine&_
	    "    ,[CoursesNames] CN" & vbNewLine&_
	    "WHERE M.[Course_Id]=C.[Course_Id]" & vbNewLine&_
	    "    AND C.[Course_Number]=CN.[Course_Number]" & vbNewLine
	If Request("course_number") > "" Then
		strQuery = strQuery & "    AND CN.[Course_Number]=" & Request("course_number") & vbNewLine
	End If
	If Request("semester") > "" Then
		strQuery = strQuery & "    AND C.[Semester]='" & Request("semester") & "'" & vbNewLine
	End If
	strQuery = strQuery&_
	    "GROUP BY C.[Course_Number],C.[Semester],M.[Maman_Number],CN.[Name],M.[Maman_Id]" & vbNewLine&_
	    "ORDER BY C.[Course_Number],C.[Semester],M.[Maman_Number]"

	Response.Write(HTML_Table_From_Query(oConn, strQuery))

	oConn.close
	Set oConn = Nothing
%>

</body>
</html>