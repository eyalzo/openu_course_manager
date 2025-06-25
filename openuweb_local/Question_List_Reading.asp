<%@ LANGUAGE = VBScript %>
<%
'-------------------------------------------------------------------------------
' Question_List_Reading.asp
' Shows question list by specific reading only.
' Differs from Question_List.asp that allows picking course and readings.
'-------------------------------------------------------------------------------

Option Explicit
Response.CacheControl = "no-cache"	
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = 5
Response.CharSet = "windows-1255"
%>

<!--#include file="bin\common.asp" -->	
<!--#include file="bin\database.asp" -->	
<!--#include file="bin\html_tables.asp" -->	
<!--#include file="bin\html_forms.asp" -->	
<!--#include file="bin\html_styles.asp" -->	

<%
	'--- connect database
	Dim oConn
	Dim strQuery
	Database_Connect_Openu oConn

	'--- save input parameters for later use
	Dim iReadingId
	iReadingId = Request("reading_id")

	'--- get course number and reading-link, for the navigation bar
	strQuery = ""&_
		"SELECT " & vbNewLine&_
		"    CN.[Course_Number] AS [מספר קורס]" & vbNewLine&_
		"    ,R.[Description] AS [חומר לקריאה]" & vbNewLine&_
		"    ,R.[Reading_Id] AS [קוד חומר לקריאה]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Readings] R" & vbNewLine&_
	    "    ,[CoursesNames] CN" & vbNewLine&_
	    "WHERE R.[Reading_Id]=" & iReadingId & vbNewLine&_
	    "    AND CN.[Course_Number]=R.[Course_Number]"
	Dim rs
	On Error Resume Next
	Set rs = oConn.Execute(strQuery)
	CheckError strQuery
%>

<html>

<head>
	<link href="Openu.css" rel="stylesheet" type="text/css">
	<title>מאגר שאלות עבור חומר מסוים</title>
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">

<!-- Navigation bar -->
<b>[</b><a href="default.asp">דף ראשי</a><b>]:
[</b><a href="course_list.asp">רשימת קורסים</a><b>]:
[</b><a href="course_details.asp?course_number=<% = rs(0) %>">קורס <% = rs(0) %></a><b>]:
שאלות '<% = rs(1) %>'</b><hr>

<%

	'---------------------------------------------------------------------------
	'--- everything was selected, so now show the questions
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=question_details.asp?question_id='+CAST(QT.[Question_Id] AS varchar)+'>'+CAST(QT.[Question_Id] AS varchar)+'</a>' AS [מס']" & vbNewLine&_
		"    ,CONVERT(varchar(10),QT.[Date_Created],105) AS [הכנסה]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN LEN(QT.[Source])<=20 THEN QT.[Source]" & vbNewLine&_
		"        ELSE LEFT(QT.[Source],17)+'<font color=white>...</font>'" & vbNewLine&_
		"        END [מקור השאלה]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN LEN(QT.[Question_Text])<=100 THEN '<small>'+QT.[Question_Text]+'</small>'" & vbNewLine&_
		"        ELSE '<small>'+LEFT(QT.[Question_Text],97)+'<font color=white>...</font></small>'" & vbNewLine&_
		"        END [שאלה ראשית]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN QT.[Is_New] IS NULL THEN '?'" & vbNewLine&_
		"        WHEN QT.[Is_New]='כן' THEN '<font color=green>כן</font>'" & vbNewLine&_
		"        ELSE 'לא'" & vbNewLine&_
		"        END [חדשה]" & vbNewLine&_
		"    -- Suitable for Mamans and how many appearences so far" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
	    "        WHEN QT.[For_Maman] IS NULL THEN '<font color=red>?</font> ('+CAST(COUNT(DISTINCT MQ.[Maman_Id]) AS varchar)+')'" & vbNewLine&_
	    "        WHEN QT.[For_Maman]='לא' THEN ''" & vbNewLine&_
		"        ELSE QT.[For_Maman]+' ('+CAST(COUNT(DISTINCT MQ.[Maman_Id]) AS varchar)+')'" & vbNewLine&_
		"        END [לממ""ן]" & vbNewLine&_
		"    -- Suitable for Exams and how many appearences so far" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
	    "        WHEN QT.[For_Exam] IS NULL THEN '<font color=red>?</font> ('+CAST(COUNT(DISTINCT EQ.[Exam_Id]) AS varchar)+')'" & vbNewLine&_
	    "        WHEN QT.[For_Exam]='לא' THEN ''" & vbNewLine&_
		"        ELSE QT.[For_Exam]+' ('+CAST(COUNT(DISTINCT EQ.[Exam_Id]) AS varchar)+')'" & vbNewLine&_
		"        END [למבחן]" & vbNewLine&_
		"    -- Suitable for Learning Guide" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
	    "        WHEN QT.[For_Guide] IS NULL THEN '<font color=red>?</font>'" & vbNewLine&_
	    "        WHEN QT.[For_Guide]='לא' THEN ''" & vbNewLine&_
		"        ELSE QT.[For_Guide]+' (0)'" & vbNewLine&_
		"        END [למדריך]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Readings] R" & vbNewLine&_
	    "    ,[QuestionsText] QT" & vbNewLine&_
	    "        -- Count the number of Mamans" & vbNewLine&_
	    "        LEFT OUTER JOIN [MamansQuestions] MQ" & vbNewLine&_
	    "            ON QT.[Question_Id]=MQ.[Question_Id]" & vbNewLine&_
	    "        -- Count the number of Exams" & vbNewLine&_
	    "        LEFT OUTER JOIN [ExamsQuestions] EQ" & vbNewLine&_
	    "            ON QT.[Question_Id]=EQ.[Question_Id]" & vbNewLine&_
	    "WHERE QT.[Reading_Id]=R.[Reading_Id]" & vbNewLine&_
		"    AND R.[Reading_Id]=" & iReadingId & vbNewLine&_
		"GROUP BY R.[Week_From],R.[Week_To],QT.[Question_Id],QT.[Date_Created],QT.[Source],QT.[Is_New],QT.[For_Maman],QT.[For_Exam],QT.[For_Guide],QT.[Question_Text]" & vbNewLine&_
		"ORDER BY R.[Week_From],R.[Week_To],QT.[Question_Id]"

	Response.Write(HTML_Style_Header3("Question", "רשימת שאלות", HTML_Style_Button2("שאלה חדשה עבור אותו חומר", "question_details.asp?reading_id=" & iReadingId & "&action=to_new_question", "/bin/images/new1.gif", True))&_
		HTML_Table_From_Query(oConn, strQuery))

	oConn.close
	Set oConn = Nothing
%>

</body>
</html>