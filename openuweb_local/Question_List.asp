<%@ LANGUAGE = VBScript %>
<%
'-------------------------------------------------------------------------------
' Question_List.asp
' Shows question list by one of two: 
'     1. Manual selection of course and weeks range.
'     2. Link with specific reading_id.
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

<html>

<head>
	<link href="Openu.css" rel="stylesheet" type="text/css">
	<title>מאגר שאלות</title>
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">
<%
	'--- connect database
	Dim oConn
	Dim strQuery
	Database_Connect_Openu oConn

	'--- save request details for later use
	Dim iCourseNumber
	Dim iWeekFrom
	Dim iWeekTo
	Dim iReadingId
	iCourseNumber = Request("course_number")
	iReadingId = Request("reading_id")
	iWeekFrom = Request("week_from")
	iWeekTo = Request("week_to")
%>
<!-- Navigation bar -->
<b>[</b><a href="default.asp">דף ראשי</a><b>]:&nbsp;[</b><a href="course_list.asp">רשימת קורסים</a><b>]:&nbsp;[</b><a href="course_details.asp?course_number=<% = iCourseNumber %>">קורס <% = iCourseNumber %></a><b>]:&nbsp;מאגר שאלות</b><hr>
<%	
	'--- courses filter
	Dim strQueryCourses
	strQueryCourses = ""&_
		"SELECT '0', '- בחר קורס -' AS [Course], '', ''" & vbNewLine&_
		"UNION" & vbNewLine&_
		"SELECT" & vbNewLine&_
		"    CASE" & vbNewLine&_
		"        WHEN CAST(CN.[Course_Number] as varchar)='" & iCourseNumber & "' THEN '1'" & vbNewLine&_
		"        ELSE '0'" & vbNewLine&_
		"        END" & vbNewLine&_
		"    ,CAST([Course_Number] as varchar)" & vbNewLine&_
		"    ,[Name]" & vbNewLine&_
		"    ,CAST([Course_Number] as varchar)" & vbNewLine&_
		"FROM [CoursesNames] CN" & vbNewLine&_
		"ORDER BY [Course]"

	'--- if no course was picked, then show course list only
	If iCourseNumber = "" Then
		Response.Write(HTML_Form("Question_List.asp", ""&_
			HTML_Style_Info1("קורס", HTML_Input_Select_From_Query(oConn, strQueryCourses, "course_number"))&_
			HTML_Input_Auto_Submit("course_number")))
		Response.End
	End If

	'---------------------------------------------------------------------------
	'--- from-week list
	Dim strQueryWeekFrom
	strQueryWeekFrom = "SELECT '0', '- בחר שבוע -', '', 0 AS [Week]" & vbNewLine
	Dim i
	For i = 1 to 17
		strQueryWeekFrom = strQueryWeekFrom&_
			"UNION" & vbNewLine&_
			"SELECT" & vbNewLine&_
			"    CASE" & vbNewLine&_
			"        WHEN '" & i & "'='" & iWeekFrom & "' THEN '1'" & vbNewLine&_
			"        ELSE '0'" & vbNewLine&_
			"        END" & vbNewLine&_
			"    ,'" & i & "'" & vbNewLine&_
			"    ,[Description]" & vbNewLine&_
			"    ," & i & " AS [Week]" & vbNewLine&_
			"FROM [Readings] R" & vbNewLine&_
			"WHERE R.[Course_Number]=" & iCourseNumber & vbNewLine&_
			"    AND (R.[Week_From]=" & i & " OR R.[Week_From]<" & i & " AND R.[Week_To]>=" & i & ")" & vbNewLine
	Next
	strQueryWeekFrom = strQueryWeekFrom&_
		"ORDER BY [Week]"

	'--- check if from-week was already selected	
	If iWeekFrom = "" Or iWeekFrom = 0 Then
		Response.Write(HTML_Form("Question_List.asp", ""&_
			HTML_Style_Info1("קורס", HTML_Input_Select_From_Query(oConn, strQueryCourses, "course_number"))&_
			HTML_Style_Info1("החל משבוע", HTML_Input_Select_From_Query(oConn, strQueryWeekFrom, "week_from")))&_
			HTML_Input_Auto_Submit("course_number") & HTML_Input_Auto_Submit("week_from"))		
		Response.End
	End If

	'---------------------------------------------------------------------------
	'--- to-week list
	Dim strQueryWeekTo
	strQueryWeekTo = "SELECT '0', '- בחר שבוע -', '', 0 AS [Week]" & vbNewLine
	For i = iWeekFrom to 17
		strQueryWeekTo = strQueryWeekTo&_
			"UNION" & vbNewLine&_
			"SELECT" & vbNewLine&_
			"    CASE" & vbNewLine&_
			"        WHEN '" & i & "'='" & iWeekTo & "' THEN '1'" & vbNewLine&_
			"        ELSE '0'" & vbNewLine&_
			"        END" & vbNewLine&_
			"    ,'" & i & "'" & vbNewLine&_
			"    ,[Description]" & vbNewLine&_
			"    ," & i & " AS [Week]" & vbNewLine&_
			"FROM [Readings] R" & vbNewLine&_
			"WHERE R.[Course_Number]=" & iCourseNumber & vbNewLine&_
			"    AND (R.[Week_From]=" & i & " OR R.[Week_From]<" & i & " AND R.[Week_To]>=" & i & ")" & vbNewLine
	Next
	strQueryWeekTo = strQueryWeekTo&_
		"ORDER BY [Week]"
	
	'--- check if from-week was already selected	
	If iWeekTo = "" Or iWeekTo = 0 Then
		Response.Write(HTML_Form("Question_List.asp", ""&_
			HTML_Style_Info1("קורס", HTML_Input_Select_From_Query(oConn, strQueryCourses, "course_number"))&_
			HTML_Style_Info1("החל משבוע", HTML_Input_Select_From_Query(oConn, strQueryWeekFrom, "week_from"))&_
			HTML_Style_Info1("ועד שבוע", HTML_Input_Select_From_Query(oConn, strQueryWeekTo, "week_to")))&_
			HTML_Input_Auto_Submit("course_number") & HTML_Input_Auto_Submit("week_from") & HTML_Input_Auto_Submit("week_to"))
		Response.End
	End If

	'---------------------------------------------------------------------------
	'--- everything was selected, so now show the questions
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=question_details.asp?question_id='+CAST(QT.[Question_Id] AS varchar)+'>'+CAST(QT.[Question_Id] AS varchar)+'</a>' AS [מס']" & vbNewLine&_
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
		"    ,CASE" & vbNewLine&_
		"        WHEN LEN(QT.[Question_Text])<=80 THEN '<small>'+QT.[Question_Text]+'</small>'" & vbNewLine&_
		"        ELSE '<small>'+LEFT(QT.[Question_Text],77)+'<font color=white>...</font></small>'" & vbNewLine&_
		"        END [שאלה ראשית]" & vbNewLine&_
		"    ,COUNT(QTS.[Sub_Number]) AS [סעיף]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Readings] R" & vbNewLine&_
	    "    ,[QuestionsText] QT" & vbNewLine&_
	    "        LEFT OUTER JOIN [QuestionsTextSub] QTS" & vbNewLine&_
	    "            ON QT.[Question_Id]=QTS.[Question_Id]" & vbNewLine&_
	    "WHERE QT.[Reading_Id]=R.[Reading_Id]" & vbNewLine&_
		"    AND R.[Course_Number]=" & iCourseNumber & vbNewLine&_
		"    AND" & vbNewLine&_
		"    (" & vbNewLine&_
		"        '" & iReadingId & "'<>'' AND R.[Reading_Id]='" & iReadingId & "'" & vbNewLine&_
		"        OR" & vbNewLine&_
		"        '" & iReadingId & "'='' AND R.[Week_From]>=" & iWeekFrom & vbNewLine&_
		"        AND (R.[Week_To] IS NULL OR R.[Week_To]<=" & iWeekTo & ")" & vbNewLine&_
		"    )" & vbNewLine&_
		"GROUP BY QT.[Source],QT.[Question_Id],R.[Week_From],R.[Week_To],R.[Description],QT.[Date_Created],QT.[Question_Text]" & vbNewLine&_
		"ORDER BY R.[Week_From],R.[Week_To],QT.[Question_Id]"

	'--- show a remark in case table was selected by reading_id and not weeks
	If iReadingId > "" Then
		Response.Write(HTML_Form("Question_List.asp", ""&_
			HTML_Style_Info1("קורס", HTML_Input_Select_From_Query(oConn, strQueryCourses, "course_number"))&_
			HTML_Input_Auto_Submit("course_number")))
		Response.Write(HTML_Style_Header3("Question", "רשימת שאלות", HTML_Style_Button2("שאלה חדשה עבור אותו חומר", "question_details.asp?reading_id=" & iReadingId & "&action=to_new_question", "/bin/images/new1.gif", True))&_
		"<font color=red>(נבחר לפי חומר לימוד ולא לפי טווח השבועות)</font>")
	Else
		Response.Write(HTML_Form("Question_List.asp", ""&_
			HTML_Style_Info1("קורס", HTML_Input_Select_From_Query(oConn, strQueryCourses, "course_number"))&_
			HTML_Style_Info1("החל משבוע", HTML_Input_Select_From_Query(oConn, strQueryWeekFrom, "week_from"))&_
			HTML_Style_Info1("ועד שבוע", HTML_Input_Select_From_Query(oConn, strQueryWeekTo, "week_to")))&_
			HTML_Input_Auto_Submit("course_number") & HTML_Input_Auto_Submit("week_from") & HTML_Input_Auto_Submit("week_to"))	
	End If
	Response.Write(HTML_Table_From_Query(oConn, strQuery))

	oConn.close
	Set oConn = Nothing
%>

</body>
</html>