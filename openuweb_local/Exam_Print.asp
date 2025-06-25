<%@ LANGUAGE = VBScript %>
<%
'-------------------------------------------------------------------------------
' Exam_Print.asp
' Prints an exam, ready for delivery.
'-------------------------------------------------------------------------------

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
	Dim iExamId
	Dim bShowAnswer
	iExamId = Request("exam_id")
	bShowAnswer = (Request("show_answer") = "yes")

	'--- general exam details
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    E.[Exam_Moed] AS [����]" & vbNewLine&_
		"    ,CAST(C.[Course_Number] AS nvarchar)+' - '+CN.[Name] AS [�����]" & vbNewLine&_
		"    ,COUNT(EQ.[Exam_Id]) AS [���� ������]" & vbNewLine&_
		"    ,C.[Semester] AS [�����]" & vbNewLine&_
		"    ,CONVERT(varchar,E.[Exam_Date],104) AS [����� ������]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Exams] E" & vbNewLine&_
	    "       LEFT OUTER JOIN [Courses] C" & vbNewLine&_
	    "           ON E.[Course_Id]=C.[Course_Id]" & vbNewLine&_
	    "       LEFT OUTER JOIN [CoursesNames] CN" & vbNewLine&_
	    "           ON C.[Course_Number]=CN.[Course_Number]" & vbNewLine&_
	    "       LEFT OUTER JOIN [ExamsQuestions] EQ" & vbNewLine&_
	    "           ON E.[Exam_Id]=EQ.[Exam_Id]" & vbNewLine&_
	    "WHERE E.[Exam_Id]=" & iExamId & vbNewLine&_
	    "GROUP BY E.[Exam_Id],CN.[Name],C.[Course_Number],E.[Exam_Moed],C.[Semester],E.[Exam_Date]"

	'--- get items to print them in formatted table
	Dim rs
	On Error Resume Next
	Set rs = oConn.Execute(strQuery)
	CheckError strQuery
%>

<html>

<head>
	<title><% If bShowAnswer Then Response.Write("������� �") End If %>����� ��� <% = rs(0) %>&nbsp;<% = rs(1) %>&nbsp;<% = rs(3) %></title>
	<!--- unique style for English -->
	<style>
BODY
{
    FONT-SIZE: 11pt;
    FONT-FAMILY: Times New Roman;
}
	</style>	
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">

<%
	Response.Write(""&_
		"    <!-- ����� ������ ������ -->"&_
		"    <p style='font-family: David; font-size: 18pt; text-align: center;'><b>����� ����� ���</b></p>"&_
		"    <p style='font-family: David; font-size: 18pt; text-align: center;'><b>" & rs(1) & "</b></p>"&_
		"<p dir=rtl style='font-family: David; font-size: 12pt; text-align: center;'>(���� " & rs(0) & " - " & rs(3) & ")</p>" & vbNewLine&_
		"<p>&nbsp;</p>" & vbNewLine&_
		"<p>&nbsp;</p>" & vbNewLine&_
		"<p style='font-family: David; font-size: 12pt; text-align: center;'><b>���� ������: </b>������ " & rs(2) & " �����. ���� ����� �� ����.</p>" & vbNewLine&_
		"<p style='font-family: David; font-size: 12pt; text-align: center;'>���� �� ���� ����� ���� ������.</p>" & vbNewLine&_
		"&nbsp;")

	'---------------------------------------------------------------------------
	'--- show the questions in real exam format

	'--- get all question headers
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    EQ.[Question_Number] AS [����]" & vbNewLine&_
		"    ,EQ.[Max_Grade] AS [������]" & vbNewLine&_
		"    ,QT.[Question_Text] AS [���� �����]" & vbNewLine&_
		"    ,EQ.[Question_Id]" & vbNewLine&_
		"    ,COUNT(DISTINCT QTS.[Sub_Number]) AS [���� ������]" & vbNewLine&_
		"    ,QT.[Answer_Text] AS [����� �����]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [ExamsQuestions] EQ" & vbNewLine&_
	    "       LEFT OUTER JOIN [QuestionsText] QT" & vbNewLine&_
	    "           ON EQ.[Question_Id]=QT.[Question_Id]" & vbNewLine&_
	    "       LEFT OUTER JOIN [QuestionsTextSub] QTS" & vbNewLine&_
	    "           ON QT.[Question_Id]=QTS.[Question_Id]" & vbNewLine&_
	    "WHERE EQ.[Exam_Id]=" & iExamId & vbNewLine&_
	    "GROUP BY EQ.[Question_Number],EQ.[Max_Grade],QT.[Question_Text],QT.[Answer_Text],EQ.[Question_Id]" & vbNewLine&_
	    "ORDER BY EQ.[Question_Number]"
	Dim rsSub
	On Error Resume Next
	Set rs = oConn.Execute(strQuery)
	CheckError strQuery
	'--- print question headers
	Do while (Not rs.eof)
		Response.Write("<p>&nbsp;</p><p dir=rtl style='font-family: David; font-size: 14pt'><b><u>" & "���� " & rs(0) & "</u></b>&nbsp;&nbsp;&nbsp;<font style='font-family: David; font-size: 12pt'>(" & rs(1) & " ������)" & "</font></p>" & vbNewLine)
		'--- main question
		If rs(2) > "" Then
			Response.Write("<p dir=rtl style='font-family: David; font-size: 12pt'>" & rs(2) & "</p>" & vbNewLine)
		End If
		'--- print all sub-questions
		strQuery = ""&_
			"SELECT" & vbNewLine&_
			"    NCHAR(UNICODE('�')+QTS.[Sub_Number]-1) AS [����]" & vbNewLine&_
			"    ,QTS.[Relative_Grade] AS [��� ����� �������]" & vbNewLine&_
			"    ,QTS.[Question_Text] AS [����]" & vbNewLine&_
			"    ,QTS.[Sub_Number] AS [���� ������ �� ����]" & vbNewLine&_
			"    ,ISNULL(QTS.[Answer_Text],'') AS [�����]" & vbNewLine&_
		    "FROM" & vbNewLine&_
		    "    [QuestionsTextSub] QTS" & vbNewLine&_
		    "WHERE QTS.[Question_Id]=" & rs(3) & vbNewLine&_
		    "ORDER BY QTS.[Sub_Number]"
		Set rsSub = oConn.Execute(strQuery)
		CheckError strQuery
		'--- calculate grades for sub-questions
		Dim iTotalSubs ' total grade so far
		Dim iQuestionGrade
		Dim iSubGrade
		iTotalSubs = 0
		iQuestionGrade = CInt(rs(1))
		'---- number of sub-questions for later calculations
		Dim iNumberOfSubs
		iNumberOfSubs = CInt(rs(4))
		'--- print main answer if needed
		If bShowAnswer Then
			'--- if there is an asnwer, then show it
			If rs(5) > "" Then
				Response.Write("<span style='font-family: David; font-size: 12pt'><font color=red>" & rs(5) & "</font></span>")
			'--- add the "no answer" warning only if there are no sub-questions
			ElseIf iNumberOfSubs = 0 Then
				Response.Write("<p style='font-family: David; font-size: 12pt'><font color=red><b>(��� �����)</b></font></p>")
			End If				
		End If
		'--- sub-questions loop
		Do while (Not rsSub.eof)
			'--- serial number of this question
			Dim iSerialNumberOfThisSub
			iSerialNumberOfThisSub = CInt(rsSub(3))
			'--- if it's the last sub-question then take the points remaining
			If iSerialNumberOfThisSub = iNumberOfSubs Then
				iSubGrade = iQuestionGrade - iTotalSubs
			Else
				'--- calculate how many points were given so far
				iSubGrade = CInt(rsSub(1)) ' percentage
				If iSubGrade > 0 Then
					'--- percentage was specified
					iSubGrade = iSubGrade * iQuestionGrade \ 100
				Else
					'--- default, meaning calculate how many points are left and how many subs are left
					iSubGrade = (iQuestionGrade - iTotalSubs) \ (iNumberOfSubs - iSerialNumberOfThisSub + 1)
				End If
				iTotalSubs = iTotalSubs + iSubGrade
			End If
			Response.Write(""&_
				"<table style='font-family: David; font-size: 12pt' width=100% colspan=0 rowspan=0>"&_
				"    <tr valign=top>"&_
				"        <td dir=rtl width=50>" & "(" & iSubGrade & " ��')" & "</td>"&_
				"        <td dir=rtl width=30>" & rsSub(0) & "." & "</td>"&_
				"        <td dir=rtl>" & rsSub(2) & "</td>"&_
				"    </tr>"&_
				"</table>")
			'--- print answer if needed
			If bShowAnswer Then
				If IsNull(rsSub(4)) Or rsSub(4) = "" Then
					Response.Write("<p style='font-family: David; font-size: 12pt'><font color=red><b>(��� �����)</b></font></p>")
				Else
					Response.Write(""&_
						"<table style='font-family: David; font-size: 12pt' width=100% colspan=0 rowspan=0>"&_
						"    <tr valign=top>"&_
						"        <td dir=rtl width=50>&nbsp;</td>"&_
						"        <td dir=rtl width=30>&nbsp;</td>"&_
						"        <td dir=rtl><font color=red>" & rsSub(4) & "</font></td>"&_
						"    </tr>"&_
						"</table>")
				End If				
			End If
			rsSub.MoveNext
		Loop
'		Response.Write("<br><br>" & vbNewLine)
		rs.MoveNext
	Loop

'	Response.Write("</span>")

	'--- close recordset
	rs.Close
	Set rs = Nothing
	
	oConn.close
	Set oConn = Nothing
%>

</body>
</html>