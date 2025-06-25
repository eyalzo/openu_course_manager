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

<%
	'--- connect database
	Dim oConn
	Dim strQuery
	Database_Connect_Openu oConn

	'--- save request details for later use
	Dim iQuestionId
	Dim strAction
	Dim iSubNumber
	iQuestionId = Request("question_id")
	strAction = Request("action")

	'---------------------------------------------------------------------------
	' Update database upon request
	If Request("REQUEST_METHOD") = "POST" Or Left(strAction,6) = "to_new" Then
		If strAction = "to_new_question" Then
			If Request("reading_id") = "" Or Request("reading_id") = 0 Then
				TerminateWithErrorMessage "Missing 'reading_id' !"
			End If
			strQuery = ""&_
				"INSERT INTO [QuestionsText]" & vbNewLine&_
				"    ([Reading_Id])" & vbNewLine&_
				"VALUES (" & Request("reading_id") & ")"
			iQuestionId = Database_Run_Query_Return_Id(oConn, strQuery)
			Response.Redirect("question_details.asp?action=to_update_details&question_id=" & iQuestionId)
		'--- new sub-question, meaning get highest so far, add one, and get directly into edit mode
		ElseIf strAction = "to_new_sub" Then
			strQuery = ""&_
				"SELECT ISNULL(MAX(QTS.[Sub_Number])+1,1)" & vbNewLine&_
				"FROM [QuestionsTextSub] QTS" & vbNewLine&_
				"WHERE QTS.[Question_Id]=" & iQuestionId
			iSubNumber = Database_Run_Query_Return_String(oConn, strQuery)
			strQuery = ""&_
				"INSERT INTO [QuestionsTextSub]" & vbNewLine&_
				"    ([Question_Id],[Sub_Number])" & vbNewLine&_
				"VALUES (" & iQuestionId & "," & iSubNumber & ")"
			Database_Run_Query oConn, strQuery
			Response.Redirect("question_details.asp?action=to_update_question_sub_" & iSubNumber & "&question_id=" & iQuestionId)
		ELseIf strAction = "do_update_question" Then
			strQuery = ""&_
				"UPDATE [QuestionsText]" & vbNewLine&_
				"SET [Question_Text]='" & String_To_SQL_Server(Request("text_body")) & "'" & vbNewLine&_
				"    ,[Date_Question_Modified]=GETDATE()" & vbNewLine&_
				"WHERE [Question_Id]=" & iQuestionId
		ElseIf strAction = "do_update_answer" Then
			strQuery = ""&_
				"UPDATE [QuestionsText]" & vbNewLine&_
				"SET [Answer_Text]='" & String_To_SQL_Server(Request("text_body")) & "'" & vbNewLine&_
				"    ,[Date_Answer_Modified]=GETDATE()" & vbNewLine&_
				"WHERE [Question_Id]=" & iQuestionId
		'--- sub-question
		ElseIf Left(strAction,23) = "do_update_question_sub_" Then
			iSubNumber = Mid(strAction, 24)
			strQuery = ""&_
				"UPDATE [QuestionsText]" & vbNewLine&_
				"SET [Date_Question_Modified]=GETDATE()" & vbNewLine&_
				"WHERE [Question_Id]=" & iQuestionId & vbNewLine&_
				"UPDATE [QuestionsTextSub]" & vbNewLine&_
				"SET [Question_Text]='" & String_To_SQL_Server(Request("text_body")) & "'" & vbNewLine&_
				"    ,[Relative_Grade]=" & Request("relative_grade") & vbNewLine&_
				"WHERE [Question_Id]=" & iQuestionId & vbNewLine&_
				"    AND [Sub_Number]=" & iSubNumber
		'--- sub-answer
		ElseIf Left(strAction,21) = "do_update_answer_sub_" Then
			iSubNumber = Mid(strAction, 22)
			strQuery = ""&_
				"UPDATE [QuestionsText]" & vbNewLine&_
				"SET [Date_Answer_Modified]=GETDATE()" & vbNewLine&_
				"WHERE [Question_Id]=" & iQuestionId & vbNewLine&_
				"UPDATE [QuestionsTextSub]" & vbNewLine&_
				"SET [Answer_Text]='" & String_To_SQL_Server(Request("text_body")) & "'" & vbNewLine&_
				"WHERE [Question_Id]=" & iQuestionId & vbNewLine&_
				"    AND [Sub_Number]=" & iSubNumber
		'--- source and comments
		ElseIf strAction = "do_update_details" Then
			strQuery = ""&_
				"UPDATE [QuestionsText]" & vbNewLine&_
				"SET [Comments_Text]='" & String_To_SQL_Server(Request("text_body")) & "'" & vbNewLine&_
				"    ,[Source]='" & String_To_SQL_Server(Request("source_text")) & "'" & vbNewLine&_
				"    ,[Is_New]='" & Request("is_new") & "'" & vbNewLine&_
				"    ,[For_Maman]='" & Request("for_maman") & "'" & vbNewLine&_
				"    ,[For_Exam]='" & Request("for_exam") & "'" & vbNewLine&_
				"    ,[For_Guide]='" & Request("for_guide") & "'" & vbNewLine&_
				"WHERE [Question_Id]=" & iQuestionId
		Else
			TerminateWithErrorMessage "Action: '" & strAction & "'"
		End If
		'--- run the query and redirect
		Database_Run_Query oConn, strQuery
		Response.Redirect("question_details.asp?question_id=" & iQuestionId)
	End If	

	'--- get course number and reading-link, for the navigation bar
	strQuery = ""&_
		"SELECT " & vbNewLine&_
		"    CN.[Course_Number] AS [מספר קורס]" & vbNewLine&_
		"    ,R.[Description] AS [חומר לקריאה]" & vbNewLine&_
		"    ,R.[Reading_Id] AS [קוד חומר לקריאה]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [QuestionsText] QT" & vbNewLine&_
	    "    ,[Readings] R" & vbNewLine&_
	    "    ,[CoursesNames] CN" & vbNewLine&_
	    "WHERE QT.[Question_Id]=" & iQuestionId & vbNewLine&_
	    "    AND QT.[Reading_Id]=R.[Reading_Id]" & vbNewLine&_
	    "    AND CN.[Course_Number]=R.[Course_Number]"
	Dim rs
	On Error Resume Next
	Set rs = oConn.Execute(strQuery)
	CheckError strQuery
%>
<html>

<head>
	<link href="Openu.css" rel="stylesheet" type="text/css">
	<title>פרטי שאלה</title>
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">

<!-- Navigation bar -->
<b>[</b><a href="default.asp">דף ראשי</a><b>]:
[</b><a href="course_list.asp">רשימת קורסים</a><b>]:
[</b><a href="course_details.asp?course_number=<% = rs(0) %>">קורס <% = rs(0) %></a><b>]:
[</b><a href="question_list_reading.asp?course_number=<% = rs(0) %>&reading_id=<% = rs(2) %>">שאלות '<% = rs(1) %>'</a><b>]:
שאלה <% = iQuestionId %></b><hr>
<%
	'--- handle a cookie for show answers: yes/no
	Dim bShowAnswers
	bShowAnswers = Request("show_answers") ' "yes" or empty
	If bShowAnswers = "yes" Or bShowAnswers = "no" Then
		'--- keep that as a user preference
		Response.Cookies("question_details")("show_answers") = bShowAnswers
	Else
		bShowAnswers = Request.Cookies("question_details")("show_answers")
	End If

	'--- general question details
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    CAST(CN.[Course_Number] AS varchar)+' '+CN.[Name] AS [קורס]" & vbNewLine&_
		"    ,R.[Description] AS [חומר לקריאה]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN R.[Week_From] IS NULL THEN 'כללי'" & vbNewLine&_
		"        WHEN R.[Week_To] IS NULL THEN CAST(R.[Week_From] AS varchar)" & vbNewLine&_
		"        ELSE CAST(R.[Week_From] AS varchar)+'-'+CAST(R.[Week_To] AS varchar)" & vbNewLine&_
		"        END [שבוע בקורס]" & vbNewLine&_
		"    ,CONVERT(varchar(10),QT.[Date_Created],105) AS [תאריך הכנסה]" & vbNewLine&_
		"    ,CONVERT(varchar(10),QT.[Date_Question_Modified],105) AS [תאריך עדכון שאלה]" & vbNewLine&_
		"    ,CONVERT(varchar(10),QT.[Date_Answer_Modified],105) AS [תאריך עדכון תשובה]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [QuestionsText] QT" & vbNewLine&_
	    "    ,[Readings] R" & vbNewLine&_
	    "    ,[CoursesNames] CN" & vbNewLine&_
	    "WHERE QT.[Question_Id]=" & iQuestionId & vbNewLine&_
	    "    AND QT.[Reading_Id]=R.[Reading_Id]" & vbNewLine&_
	    "    AND CN.[Course_Number]=R.[Course_Number]"
	Response.Write(""&_
		HTML_Style_Header3("Question","פרטים",""&_
			HTML_Style_Button2("ביטול עריכה", "question_details.asp?question_id=" & iQuestionId, "/bin/images/view1.gif", strAction="to_update_details")&_
			HTML_Style_Button2("עריכה", "question_details.asp?action=to_update_details&question_id=" & iQuestionId, "/bin/images/edit1.gif", Left(strAction,3)<>"to_")) & vbNewLine&_
		HTML_Info_From_Query(oConn, strQuery, True))

	'---------------------------------------------------------------------------
	'--- editable general question details
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    CASE" & vbNewLine&_
		"        WHEN QT.[Source] IS NULL THEN ''" & vbNewLine&_
		"        ELSE QT.[Source]" & vbNewLine&_
		"        END [מקור]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN QT.[Comments_Text] IS NULL THEN ''" & vbNewLine&_
		"        ELSE QT.[Comments_Text]" & vbNewLine&_
		"        END [הערות]" & vbNewLine&_
		"    -- Images" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN QT.[Is_New]='כן' THEN '<img border=0 width=36 height=18 src=/bin/images/new4.gif alt=''שאלה חדשה בקורס''>'" & vbNewLine&_
		"        ELSE ''" & vbNewLine&_
		"        END" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN QT.[For_Maman]='כן' THEN '<img border=0 width=24 height=24 src=/bin/images/imgForMaman1.gif alt=''מתאימה לממ""ן''>'" & vbNewLine&_
		"        ELSE ''" & vbNewLine&_
		"        END" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN QT.[For_Exam]='כן' THEN '<img border=0 width=24 height=24 src=/bin/images/imgForExam1.gif alt=''מתאימה למבחן סיום''>'" & vbNewLine&_
		"        ELSE ''" & vbNewLine&_
		"        END" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN QT.[For_Guide]='כן' THEN '<img border=0 width=24 height=24 src=/bin/images/imgForGuide1.gif alt=''מתאימה למדריך למידה''>'" & vbNewLine&_
		"        ELSE ''" & vbNewLine&_
		"        END" & vbNewLine&_
		"    -- Yes/No select-boxes" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN QT.[Is_New]='כן' THEN 'כן'" & vbNewLine&_
		"        ELSE 'לא'" & vbNewLine&_
		"        END" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN QT.[For_Maman]='כן' THEN 'כן'" & vbNewLine&_
		"        ELSE 'לא'" & vbNewLine&_
		"        END" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN QT.[For_Exam]='כן' THEN 'כן'" & vbNewLine&_
		"        ELSE 'לא'" & vbNewLine&_
		"        END" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN QT.[For_Guide]='כן' THEN 'כן'" & vbNewLine&_
		"        ELSE 'לא'" & vbNewLine&_
		"        END" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [QuestionsText] QT" & vbNewLine&_
	    "WHERE QT.[Question_Id]=" & iQuestionId
	On Error Resume Next
	Set rs = oConn.Execute(strQuery)
	CheckError strQuery

	Dim strText
	strText = ""

	If strAction = "to_update_details" Then
		strText = strText&_
			HTML_Style_Info1("מקור", HTML_Input_Text("source_text",100,rs(0))) & vbNewLine&_
			HTML_Style_Info1("שאלה חדשה בקורס", ""&_
				HTML_Input_Select("is_new", "" & vbNewLine&_
					HTML_Input_Select_Option("לא", "לא", rs(6)) & vbNewLine&_
					HTML_Input_Select_Option("כן", "כן", rs(6)))) & vbNewLine&_
			HTML_Style_Info1("מתאימה לממ""ן", ""&_
				HTML_Input_Select("for_maman", "" & vbNewLine&_
					HTML_Input_Select_Option("לא", "לא", rs(7)) & vbNewLine&_
					HTML_Input_Select_Option("כן", "כן", rs(7)))) & vbNewLine&_
			HTML_Style_Info1("מתאימה למבחן", ""&_
				HTML_Input_Select("for_exam", "" & vbNewLine&_
					HTML_Input_Select_Option("לא", "לא", rs(8)) & vbNewLine&_
					HTML_Input_Select_Option("כן", "כן", rs(8)))) & vbNewLine&_
			HTML_Style_Info1("מתאימה למדריך למידה", ""&_
				HTML_Input_Select("for_guide", "" & vbNewLine&_
					HTML_Input_Select_Option("לא", "לא", rs(9)) & vbNewLine&_
					HTML_Input_Select_Option("כן", "כן", rs(9)))) & vbNewLine&_
			HTML_Style_Info1("הערות", Text_Body_Input(rs(1)) & HTML_Input_Set_Focus("source_text"))
	Else
		strText = strText&_
			HTML_Style_Info1("מקור", rs(0)) & vbNewLine&_
			HTML_Style_Info1("תכונות", rs(2) & rs(3) & rs(4) & rs(5)) & vbNewLine&_
			HTML_Style_Info1("הערות", rs(1))
	End If

	'---------------------------------------------------------------------------
	'--- main question and answer

	strText = strText&_
		HTML_Style_Header3("Question","שאלה",""&_
			HTML_Style_Button2("סעיף חדש", "question_details.asp?action=to_new_sub&question_id=" & iQuestionId, "/bin/images/new1.gif", True)&_
			"&nbsp;" & vbNewLine&_
			HTML_Style_Button2("ביטול עריכה", "question_details.asp?question_id=" & iQuestionId, "/bin/images/view1.gif", Left(strAction,11)="to_update_q" Or Left(strAction,11)="to_update_a")&_
			HTML_Style_Button2("ללא תשובות", "question_details.asp?question_id=" & iQuestionId & "&action=" & strAction & "&show_answers=no", "/bin/images/btnAnswerHide1.gif", bShowAnswers = "yes")&_
			HTML_Style_Button2("עם תשובות", "question_details.asp?question_id=" & iQuestionId & "&action=" & strAction & "&show_answers=yes", "/bin/images/btnAnswerShow1.gif", bShowAnswers <> "yes"))

	'--- question
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    -- handle new questions too, where text itself is NULL" & vbNewLine&_
		"    CASE" & vbNewLine&_
		"        WHEN QT.[Question_Text] IS NULL THEN ''" & vbNewLine&_
		"        ELSE QT.[Question_Text]" & vbNewLine&_
		"        END" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN QT.[Answer_Text] IS NULL THEN ''" & vbNewLine&_
		"        ELSE QT.[Answer_Text]" & vbNewLine&_
		"        END" & vbNewLine&_
	    "FROM [QuestionsText] QT" & vbNewLine&_
	    "WHERE QT.[Question_Id]=" & iQuestionId
	On Error Resume Next
	Set rs = oConn.Execute(strQuery)
	CheckError strQuery

	'--- main question - check if it's an edit mode
	If strAction = "to_update_question" Then
		strText = strText & Text_Body_Input(rs(0)) & "<br>"
	Else
		strText = strText&_
			"    <table width=100% colspan=0 rowspan=0><tr valign=top><td width=20>" & vbNewLine&_
			"        <a href=question_details.asp?question_id=" & iQuestionId & "&action=to_update_question><img src=/bin/images/edit1.gif border=0 alt='ערוך שאלה ראשית' Question></a>" & vbNewLine&_
			"        </td><td>" & rs(0) & "</td></tr></table>"
	End If
	
	'--- main answer - check if it's an edit mode
	If strAction = "to_update_answer" Then
		strText = strText & Text_Body_Input(rs(1)) & "<br>"
	ElseIf bShowAnswers = "yes" Then
		strText = strText&_
			"    <table width=100% colspan=0 rowspan=0><tr valign=top><td width=20>" & vbNewLine&_
			"        <a href=question_details.asp?question_id=" & iQuestionId & "&action=to_update_answer><img src=/bin/images/btnAnswerEdit1.gif border=0 alt='ערוך תשובה ראשית'></a>" & vbNewLine&_
			"        </td><td><font color=red>" & rs(1) & "</font></td></tr></table>"
	End If
	
	'--- and sub-questions
	iSubNumber = 1
	Do While (True)
		strQuery = ""&_
			"SELECT" & vbNewLine&_
			"    ISNULL(QTS.[Question_Text],'')" & vbNewLine&_
			"    ,ISNULL(QTS.[Answer_Text],'')" & vbNewLine&_
			"    ,ISNULL(QTS.[Relative_Grade],0)" & vbNewLine&_
			"    ,NCHAR(UNICODE('א')+QTS.[Sub_Number]-1) AS [סעיף]" & vbNewLine&_
		    "FROM [QuestionsTextSub] QTS" & vbNewLine&_
		    "WHERE QTS.[Question_Id]=" & iQuestionId & vbNewLine&_
		    "    AND QTS.[Sub_Number]=" & iSubNumber
		On Error Resume Next
		Set rs = oConn.Execute(strQuery)
		CheckError strQuery

		If IsNull(rs(3)) Or rs(3) = "" Then
			Exit Do
		End If
		
		'--- sub-question - check if it's an edit mode
		strText = strText&_
			"    <table width=100% colspan=0 rowspan=0>" & vbNewLine&_
			"        <tr valign=top>" & vbNewLine&_
			"            <td width=40>(" & rs(2) & "%)</td>" & vbNewLine&_
			"            <td width=15>" & rs(3) & ".</td>" & vbNewLine
		If strAction = ("to_update_question_sub_" & iSubNumber) Then
			strText = strText & "<td>" & HTML_Style_Info1("ניקוד (באחוזים)",HTML_Input_Text("relative_grade",2,rs(2))) & Text_Body_Input(rs(0)) & "<br>" & "</td>"
		Else
			strText = strText&_
				"            <td width=25><a href=question_details.asp?question_id=" & iQuestionId & "&action=to_update_question_sub_" & iSubNumber & "><img src=/bin/images/edit1.gif border=0 alt=""ערוך שאלת סעיף " & rs(3) & """></a></td>" & vbNewLine&_
				"            <td>" & rs(0) & "</td>" & vbNewLine&_
				"        </tr>" & vbNewLine&_
				"    </table>"
		End If

		If bShowAnswers = "yes" Then
			'--- sub-answer - check if it's an edit mode
			strText = strText&_
				"    <table width=100% colspan=0 rowspan=0>" & vbNewLine&_
				"        <tr valign=top>" & vbNewLine&_
				"            <td width=40>&nbsp;</td>" & vbNewLine&_
				"            <td width=15>&nbsp;</td>" & vbNewLine
			If strAction = ("to_update_answer_sub_" & iSubNumber) Then
				strText = strText & "<td>" & Text_Body_Input(rs(1)) & "</td>"
			Else
				strText = strText&_
					"            <td width=25><a href=question_details.asp?question_id=" & iQuestionId & "&action=to_update_answer_sub_" & iSubNumber & "><img src=/bin/images/btnAnswerEdit1.gif border=0 alt=""ערוך תשובת סעיף " & rs(3) & """></a></td>" & vbNewLine&_
					"            <td><font color=red>" & rs(1) & "</font></td>" & vbNewLine&_
					"        </tr>" & vbNewLine&_
					"    </table>"
			End If
		End If
		'--- next sub-question
		iSubNumber = iSubNumber + 1
	Loop
	
	'--- wrap with form
	Response.Write(""&_
		HTML_Form("question_details.asp", ""&_
			HTML_Input_Hidden("action", "d" & mid(strAction,2))&_
			HTML_Input_Hidden("question_id", iQuestionId)&_
			strText))

	'---------------------------------------------------------------------------
	'--- show a list of mamans where this question appeared
	'temp

	oConn.close
	Set oConn = Nothing
%>

</body>
</html>

<%
Private Function Text_Body_Input(ByRef a_strInitialText)
	Text_Body_Input = ""&_
		HTML_Input_text("text_body", 500, SQL_Server_To_String(a_strInitialText))&_
		"<br>"&_
		HTML_Input_Button("שמור")&_
		HTML_Input_Set_Focus("text_body")
End Function
%>