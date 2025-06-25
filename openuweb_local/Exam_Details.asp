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
	<title>פרטי בחינה</title>
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">

<%
	'--- connect database
	Dim oConn
	Dim strQuery
	Database_Connect_Openu oConn

	'--- save request details for later use
	Dim iExamId
	Dim strAction
	iExamId = Request("exam_id")
	strAction = Request("action")
%>
<table class="PageTitle_Exam">
    <tr>
        <td class="PageTitle">פרטי בחינה</td>
    </tr>
</table>

<ul>
    <li><a href="default.asp">דף ראשי</a></li>
    <li><a href="exam_print.asp?exam_id=<% = iExamId %>">גרסה להדפסת בחינה</a></li>
    <li><a href="exam_print.asp?show_answer=yes&exam_id=<% = iExamId %>">גרסה עם תשובות, לצורך בדיקה</a></li>
</ul>
<%
	'--- general maman details
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    E.[Exam_Moed] AS [מועד]" & vbNewLine&_
		"    ,CN.[Name]+' ('+CAST(C.[Course_Number] AS nvarchar)+')' AS [הקורס]" & vbNewLine&_
		"    ,C.[Semester] AS [סמסטר]" & vbNewLine&_
		"    ,CONVERT(varchar,E.[Exam_Date],104) AS [תאריך הבחינה]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Exams] E" & vbNewLine&_
	    "       LEFT OUTER JOIN [Courses] C" & vbNewLine&_
	    "           ON E.[Course_Id]=C.[Course_Id]" & vbNewLine&_
	    "       LEFT OUTER JOIN [CoursesNames] CN" & vbNewLine&_
	    "           ON C.[Course_Number]=CN.[Course_Number]" & vbNewLine&_
	    "WHERE E.[Exam_Id]=" & iExamId
	Response.Write(HTML_Style_Header3("Exam","פרטים","") & vbNewLine&_
		HTML_Info_From_Query(oConn, strQuery, True))

	'---------------------------------------------------------------------------
	'--- everything was selected, so now show the questions
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    EQ.[Question_Number] AS [שאלה]" & vbNewLine&_
		"    ,EQ.[Max_Grade] AS [ניקוד]" & vbNewLine&_
		"    ,'<a href=question_details.asp?question_id='+CAST(QT.[Question_Id] AS varchar)+'>'+CAST(QT.[Question_Id] AS varchar)+'</a>' AS [מס' במאגר]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN R.[Week_From] IS NULL THEN 'כללי'" & vbNewLine&_
		"        WHEN R.[Week_To] IS NULL THEN CAST(R.[Week_From] AS varchar)" & vbNewLine&_
		"        ELSE CAST(R.[Week_From] AS varchar)+'-'+CAST(R.[Week_To] AS varchar)" & vbNewLine&_
		"        END [שבוע]" & vbNewLine&_
		"    ,R.[Description] AS [חומר לקריאה]" & vbNewLine&_
		"    ,CONVERT(varchar(10),QT.[Date_Created],105) AS [הכנסה]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN LEN(QT.[Source])<=20 THEN QT.[Source]" & vbNewLine&_
		"        ELSE LEFT(QT.[Source],17)+'<font color=white>...</font>'" & vbNewLine&_
		"        END [מקור השאלה]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [ExamsQuestions] EQ" & vbNewLine&_
	    "        LEFT OUTER JOIN [QuestionsText] QT" & vbNewLine&_
	    "            ON EQ.[Question_Id]=QT.[Question_Id]" & vbNewLine&_
	    "        LEFT OUTER JOIN [Readings] R" & vbNewLine&_
	    "            ON QT.[Reading_Id]=R.[Reading_Id]" & vbNewLine&_
	    "WHERE EQ.[Exam_Id]=" & iExamId & vbNewLine&_
		"ORDER BY EQ.[Question_Number]"
	Response.Write(HTML_Style_Header3("Question", "רשימת שאלות", "")&_
		HTML_Table_From_Query(oConn, strQuery))

	oConn.close
	Set oConn = Nothing
%>

</body>
</html>