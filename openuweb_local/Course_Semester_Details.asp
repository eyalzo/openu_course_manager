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
	Dim strAction
	iCourseId = Request("course_id")
	strAction = Request("action")
	'--- general details for navigation bar
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=temp.asp>'+CAST(C.[Course_Number] AS nvarchar)+'&nbsp;'+CN.[Name]+'</a><b>]: סמסטר '+CAST(C.[Semester] AS varchar)" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Courses] C" & vbNewLine&_
	    "       LEFT OUTER JOIN [CoursesNames] CN" & vbNewLine&_
	    "           ON C.[Course_Number]=CN.[Course_Number]" & vbNewLine&_
	    "WHERE C.[Course_Id]=" & iCourseId
%>
<!-- Navigation bar -->
<b>[</b><a href="default.asp">דף ראשי</a><b>]:&nbsp;[</b><a href="course_list.asp">רשימת קורסים</a><b>]:&nbsp;[</b><% = Database_Run_Query_Return_String(oConn, strQuery) %></b><hr>

<ul>
    <li><a href="course_semester_code_names.asp?course_id=<% = Request("course_id")%>">רשימת קודים לפרסום באתר</a></li>
</ul>
<%	
	'--- group list
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=group_details.asp?group_id='+CAST(CG.[Group_Id] AS varchar)+'>'+REPLICATE('0',2-LEN(CAST(CG.[Group_Number] AS varchar)))+CAST(CG.[Group_Number] AS varchar)+'</a>' AS [מס' קבוצה]" & vbNewLine&_
		"    ,COUNT(SG.[Student_Id]) AS [סטודנטים]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [CoursesGroups] CG" & vbNewLine&_
	    "        LEFT OUTER JOIN [StudentsGroups] SG" & vbNewLine&_
	    "            ON CG.[Group_Id]=SG.[Group_Id]" & vbNewLine&_
	    "WHERE CG.[Course_Id]=" & iCourseId & vbNewLine&_
	    "GROUP BY CG.[Group_Id]" & vbNewLine&_
	    "    ,CG.[Group_Number]" & vbNewLine&_
	    "ORDER BY CG.[Group_Number]"
	Response.Write(HTML_Style_Header3("Group","קבוצות","") & vbNewLine&_
	    HTML_Table_From_Query(oConn, strQuery))

	'--- maman list
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=maman_details.asp?maman_id='+CAST(M.[Maman_Id] AS varchar)+'>'+CAST(M.[Maman_Number] AS varchar)+'</a>' AS [ממ""ן]" & vbNewLine&_
	    "    ,M.[Material] AS [חומר הלימוד למטלה]" & vbNewLine&_
	    "    ,CONVERT(varchar(8),M.[Delivery_Date],5) AS [תאריך להגשה]" & vbNewLine&_
	    "    ,M.[Weight] AS [משקל]" & vbNewLine&_
	    "    ,CASE" & vbNewLine&_
	    "        WHEN M.[Mandatory] IS NULL THEN '?'" & vbNewLine&_
	    "        ELSE M.[Mandatory]" & vbNewLine&_
	    "        END [חובה]" & vbNewLine&_
	    "    ,COUNT(MQ.[Question_Number]) AS [שאלות]" & vbNewLine&_
	    "    ,CASE" & vbNewLine&_
	    "        WHEN M.[Status]='בפיתוח' THEN '<font color=red>בפיתוח</font>'" & vbNewLine&_
	    "        WHEN M.[Status]='הושלם' THEN '<font color=green>הושלם</font>'" & vbNewLine&_
	    "        WHEN M.[Status]='הודפס' THEN 'הודפס'" & vbNewLine&_
	    "        ELSE '?'" & vbNewLine&_
	    "        END [סטטוס]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Mamans] M" & vbNewLine&_
	    "        LEFT OUTER JOIN [MamansQuestions] MQ" & vbNewLine&_
	    "            ON M.[Maman_Id]=MQ.[Maman_Id]" & vbNewLine&_
	    "WHERE M.[Course_Id]=" & iCourseId & vbNewLine&_
	    "GROUP BY M.[Maman_Id],M.[Maman_Number],M.[Material],M.[Delivery_Date],M.[Weight],M.[Mandatory],M.[Status]" & vbNewLine&_
	    "ORDER BY M.[Maman_Number]"
	Response.Write(HTML_Style_Header3("Maman","מטלות","") & vbNewLine&_
	    HTML_Table_From_Query(oConn, strQuery))

	'--- exam list
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=exam_details.asp?exam_id='+CAST(E.[Exam_Id] AS varchar)+'>'+E.[Exam_Moed]+'</a>' AS [מועד]" & vbNewLine&_
	    "    ,CONVERT(varchar(8),E.[Exam_Date],5) AS [תאריך]" & vbNewLine&_
	    "    ,COUNT(EQ.[Question_Number]) AS [שאלות]" & vbNewLine&_
	    "    ,CASE" & vbNewLine&_
	    "        WHEN E.[Status]='בפיתוח' THEN '<font color=red>בפיתוח</font>'" & vbNewLine&_
	    "        WHEN E.[Status]='הושלם' THEN '<font color=green>הושלם</font>'" & vbNewLine&_
	    "        WHEN E.[Status]='הודפס' THEN 'הודפס'" & vbNewLine&_
	    "        ELSE '?'" & vbNewLine&_
	    "        END [סטטוס]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Exams] E" & vbNewLine&_
	    "        LEFT OUTER JOIN [ExamsQuestions] EQ" & vbNewLine&_
	    "            ON E.[Exam_Id]=EQ.[Exam_Id]" & vbNewLine&_
	    "WHERE E.[Course_Id]=" & iCourseId & vbNewLine&_
	    "GROUP BY E.[Exam_Id],E.[Exam_Moed],E.[Exam_Date],E.[Status]" & vbNewLine&_
	    "ORDER BY E.[Exam_Moed]"
	Response.Write(HTML_Style_Header3("Exam","בחינות","") & vbNewLine&_
	    HTML_Table_From_Query(oConn, strQuery))

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
			"        WHEN SM" & (10 + i) & ".[Cancel_Late]=1 OR SM" & (10 + i) & ".[Cancel_Copy]=1 THEN '<small>('+CAST(SM" & (10 + i) & ".[Grade] AS varchar)+')</small><b>&nbsp;<font color=red>0</font></b>'" & vbNewLine&_
			"        WHEN SM" & (10 + i) & ".[Grade] < 60 THEN '<b><font color=red>'+CAST(SM" & (10 + i) & ".[Grade] AS varchar)+'</font></b>'" & vbNewLine&_
			"        WHEN SM" & (10 + i) & ".[Grade] >= 95 THEN '<b><font color=green>'+CAST(SM" & (10 + i) & ".[Grade] AS varchar)+'</font></b>'" & vbNewLine&_
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
	
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=student_details.asp?student_id='+CAST(S.[Student_Id] AS varchar)+'>'+REPLICATE('0',9-LEN(CAST(S.[Student_Id] AS varchar)))+CAST(S.[Student_Id] AS varchar)+'</a>' AS [מס' סטודנט]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN S.[Code_Name] > '' THEN '<font color=magenta><b>'+S.[Last]+' '+S.[First]+'</b></font>'" & vbNewLine&_
		"        ELSE S.[Last]+' '+S.[First]" & vbNewLine&_
		"        END [שם (בהיפוך)]" & vbNewLine&_
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
	    "ORDER BY S.[Last],S.[First]"
	Response.Write(HTML_Style_Header3("Student","סטודנטים (כל הקבוצות)","") & vbNewLine&_
	    HTML_Table_From_Query(oConn, strQuery))

	oConn.close
	Set oConn = Nothing
%>

</body>
</html>