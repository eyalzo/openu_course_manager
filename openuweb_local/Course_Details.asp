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
	<title>פרטי קורס</title>
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">

<%
	'--- connect database
	Dim oConn
	Dim strQuery
	Database_Connect_Openu oConn

	'--- save request details for later use
	Dim iCourseNumber
	Dim strAction
	iCourseNumber = Request("course_number")
	strAction = Request("action")
%>
<!-- Navigation bar -->
<b>[</b><a href="default.asp">דף ראשי</a><b>]:&nbsp;[</b><a href="course_list.asp">רשימת קורסים</a><b>]:&nbsp;קורס <% = iCourseNumber %></b><hr>
<%	
	'--- general details
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    CAST(CN.[Course_Number] AS nvarchar)+' '+CN.[Name] AS [קורס]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [CoursesNames] CN" & vbNewLine&_
	    "WHERE CN.[Course_Number]=" & iCourseNumber
	Response.Write(HTML_Style_Header3("Course","פרטים","") & vbNewLine&_
		HTML_Info_From_Query(oConn, strQuery, True))

	'--- semester list
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=course_semester_details.asp?course_id='+CAST(C.[Course_Id] AS varchar)+'>'+C.[Semester]+'</a>' AS [סמסטר]" & vbNewLine&_
		"    ,T.[Season]+' '+CONVERT(varchar(10),T.[Start_Date],105)+' - '+CONVERT(varchar(10),DATEADD(dd,T.[Weeks]*7-1,T.[Start_Date]),105) AS [תקופה]" & vbNewLine&_
		"    ,COUNT(DISTINCT CG.[Group_Id]) AS [קבוצות]" & vbNewLine&_
		"    ,COUNT(DISTINCT SG.[Student_Id]) AS [סטודנטים]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Courses] C" & vbNewLine&_
	    "       INNER JOIN [Semesters] T" & vbNewLine&_
	    "           ON C.[Semester]=T.[Semester]" & vbNewLine&_
	    "       LEFT OUTER JOIN [CoursesGroups] CG" & vbNewLine&_
	    "           ON C.[Course_Id]=CG.[Course_Id]" & vbNewLine&_
	    "       LEFT OUTER JOIN [StudentsGroups] SG" & vbNewLine&_
	    "           ON CG.[Group_Id]=SG.[Group_Id]" & vbNewLine&_
	    "WHERE C.[Course_Number]=" & iCourseNumber & vbNewLine&_
	    "GROUP BY C.[Course_Id],C.[Semester],T.[Season],T.[Start_Date],T.[Weeks]" & vbNewLine&_
	    "ORDER BY C.[Semester]"
	Response.Write(HTML_Style_Header3("Course","סמסטרים","") & vbNewLine&_
		HTML_Table_From_Query(oConn, strQuery))

	'--- related article list
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=article_details.asp?article_id='+CAST(A.[Article_Id] AS varchar)+'>'+A.[Name]+'</a>' AS [שם]" & vbNewLine&_
		"    ,A.[Basic_Type] AS [סוג]" & vbNewLine&_
		"    ,RIGHT(CONVERT(varchar(10),A.[Publish_Date],105),7) AS [תאריך פרסום]" & vbNewLine&_
		"    ,A.[Basic_Type] AS [סוג]" & vbNewLine&_
		"    ,A.[Author] AS [מחבר]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Articles] A" & vbNewLine&_
	    "WHERE A.[Course_Number]=" & iCourseNumber & vbNewLine&_
	    "ORDER BY A.[Basic_Type],A.[Name],A.[Publish_Date] DESC"
	Response.Write(HTML_Style_Header3("Article","חומר מודפס","") & vbNewLine&_
		HTML_Table_From_Query(oConn, strQuery))

	'--- material list
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    CASE" & vbNewLine&_
		"        WHEN R.[Week_From] IS NULL THEN 'כללי'" & vbNewLine&_
		"        WHEN R.[Week_To] IS NULL THEN CAST(R.[Week_From] AS varchar)" & vbNewLine&_
		"        ELSE CAST(R.[Week_From] AS varchar)+'-'+CAST(R.[Week_To] AS varchar)" & vbNewLine&_
		"        END [שבוע]" & vbNewLine&_
		"    ,R.[Description] AS [תאור]" & vbNewLine&_
		"    ,R.[Detailed] AS [פירוט]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN COUNT(QT.[Reading_Id])=0 THEN ''" & vbNewLine&_
		"        ELSE '<a href=question_list_reading.asp?reading_id='+CAST(R.[Reading_Id] AS varchar)+'>'+CAST(COUNT(QT.[Reading_Id]) AS varchar)+'</a>'" & vbNewLine&_
		"        END [שאלות במאגר]" & vbNewLine&_
		"	 ,'<a href=question_details.asp?action=to_new_question&reading_id='+CAST(R.[Reading_Id] AS varchar)+'><img src=/bin/images/new1.gif border=0 alt=''שאלה חדשה לחומר זה''></a>'" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Readings] R" & vbNewLine&_
	    "        LEFT OUTER JOIN [QuestionsText] QT" & vbNewLine&_
	    "            ON R.[Reading_Id]=QT.[Reading_Id]" & vbNewLine&_
	    "WHERE R.[Course_Number]=" & iCourseNumber & vbNewLine&_
	    "GROUP BY R.[Reading_Id],R.[Description],R.[Detailed],R.[Week_From],R.[Week_To]" & vbNewLine&_
	    "ORDER BY R.[Week_From]"
	Response.Write(HTML_Style_Header3("Reading","חומר לימוד","") & vbNewLine&_
		HTML_Table_From_Query(oConn, strQuery))

	oConn.close
	Set oConn = Nothing
%>

</body>
</html>