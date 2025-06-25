<%@ LANGUAGE = VBScript %>
<%
'-------------------------------------------------------------------------------
' Maman_Print.asp
' Prints a maman, ready for booklet.
'-------------------------------------------------------------------------------

Option Explicit
Response.CacheControl = "no-cache"	
Response.AddHeader "Pragma", "no-cache" 
'Response.ExpiresAbsolute=#Jan 01, 1980 00:00:00# 
Response.Expires = 1
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
	Dim iMamanId
	Dim strAction
	iMamanId = Request("maman_id")
	strAction = Request("action")

	'--- general maman details
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    M.[Maman_Number] AS [ממ""ן]" & vbNewLine&_
		"    ,CN.[Name]+' ('+CAST(C.[Course_Number] AS nvarchar)+')' AS [הקורס]" & vbNewLine&_
		"    ,M.[Material] AS [חומר הלימוד למטלה]" & vbNewLine&_
		"    ,COUNT(MQ.[Maman_Id]) AS [מספר השאלות]" & vbNewLine&_
		"    ,M.[Weight] AS [משקל המטלה]" & vbNewLine&_
		"    ,C.[Semester] AS [סמסטר]" & vbNewLine&_
		"    ,CONVERT(varchar,M.[Delivery_Date],104) AS [תאריך אחרון להגשה]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Mamans] M" & vbNewLine&_
	    "       LEFT OUTER JOIN [Courses] C" & vbNewLine&_
	    "           ON M.[Course_Id]=C.[Course_Id]" & vbNewLine&_
	    "       LEFT OUTER JOIN [CoursesNames] CN" & vbNewLine&_
	    "           ON C.[Course_Number]=CN.[Course_Number]" & vbNewLine&_
	    "       LEFT OUTER JOIN [MamansQuestions] MQ" & vbNewLine&_
	    "           ON M.[Maman_Id]=MQ.[Maman_Id]" & vbNewLine&_
	    "WHERE M.[Maman_Id]=" & iMamanId & vbNewLine&_
	    "GROUP BY M.[Maman_Number],CN.[Name],C.[Course_Number],M.[Weight],M.[Material],C.[Semester],M.[Delivery_Date]"

	'--- get items to print them in formatted table
	Dim rs
	Set rs = oConn.Execute(strQuery)
	CheckError strQuery
%>

<html>

<head>
	<title>מטלה <% = rs(0) %>&nbsp;<% = rs(1) %>&nbsp;<% = rs(5) %></title>
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
		"    <!-- כותרת המטלה ומספרה -->"&_
		"    <p style='font-family: David; font-size: 30pt; text-align: center;'><b>מטלת מנחה (ממ""ן) " & rs(0) & "</b></p>"&_
		"    <!-- שם הקורס ומספרו -->"&_
		"    <p style='font-family: David; font-size: 12pt;'><b>הקורס:</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & rs(1) & "</p>"&_
		"    <!-- חומר הלימוד למטלה -->"&_
		"    <p style='font-family: David; font-size: 12pt;'><b>חומר הלימוד למטלה:</b>&nbsp;&nbsp;&nbsp;&nbsp;" & rs(2) & "</p>"&_
		"<table width=100% cellpadding=0 cellspacing=0>"&_
		"    <!-- מספר השאלות ומשקל המטלה -->"&_
		"    <tr height=36>"&_
		"        <td dir=rtl width=65% style='font-family: David; font-size: 12pt;'><b>מספר השאלות:</b>&nbsp;" & rs(3) & "</td>"&_
		"        <td dir=rtl style='font-family: David; font-size: 12pt;'><b>משקל המטלה:</b>&nbsp;" & rs(4) & "</td>"&_
		"    </tr>"&_
		"    <!-- סמסטר ומועד הגשה -->"&_
		"    <tr height=24>"&_
		"        <td dir=rtl width=65% style='font-family: David; font-size: 12pt;'><b>סמסטר:</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & rs(5) & "</td>"&_
		"        <td dir=rtl style='font-family: David; font-size: 12pt;'><b>מועד אחרון להגשה:</b>&nbsp;" & rs(6) & "</td>"&_
		"    </tr>"&_
		"</table>")

	'--- הודעת שימו-לב שגרתית בנוסח דו-מיני
	Response.Write(""&_
		"<p>&nbsp;</p>"&_
		"<table align=center width=100% border=1 cellpadding=0 cellspacing=0>"&_
		"    <tr>"&_
		"        <td dir=rtl align=center style='font-family: David; font-size: 11pt;'>אנא שימו לב:<br>מלאו בדייקנות את הטופס המלווה לממ""ן בהתאם לדוגמה שלפני המטלות.<br>העתיקו את מספר הקורס ומספר המטלה הרשומים לעיל.</td>"&_
		"    </tr>"&_
		"</table>")


	'---------------------------------------------------------------------------
	'--- show the questions in real maman format for the booklet
'	Response.Write("<span dir=rtl style='font-family: David; font-size: 12pt'>")

	'--- get all question headers
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    MQ.[Question_Number] AS [מספר]" & vbNewLine&_
		"    ,MQ.[Max_Grade] AS [נקודות]" & vbNewLine&_
		"    ,QT.[Question_Text] AS [שאלה ראשית]" & vbNewLine&_
		"    ,MQ.[Question_Id]" & vbNewLine&_
		"    ,COUNT(DISTINCT QTS.[Sub_Number]) AS [מספר סעיפים]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [MamansQuestions] MQ" & vbNewLine&_
	    "       LEFT OUTER JOIN [QuestionsText] QT" & vbNewLine&_
	    "           ON MQ.[Question_Id]=QT.[Question_Id]" & vbNewLine&_
	    "       LEFT OUTER JOIN [QuestionsTextSub] QTS" & vbNewLine&_
	    "           ON QT.[Question_Id]=QTS.[Question_Id]" & vbNewLine&_
	    "WHERE MQ.[Maman_Id]=" & iMamanId & vbNewLine&_
	    "GROUP BY MQ.[Question_Number],MQ.[Max_Grade],QT.[Question_Text],MQ.[Question_Id]" & vbNewLine&_
	    "ORDER BY MQ.[Question_Number]"
	Dim rsSub
	Set rs = oConn.Execute(strQuery)
	CheckError strQuery
	'--- print question headers
	Do while (Not rs.eof)
		Response.Write("<p>&nbsp;</p><p dir=rtl style='font-family: David; font-size: 12pt'><b>" & "שאלה " & rs(0) & "&nbsp;&nbsp;&nbsp;(" & rs(1) & " נקודות)" & "</b></p>" & vbNewLine)
		'--- main question
		If rs(2) > "" Then
			Response.Write("<p dir=rtl style='font-family: David; font-size: 12pt'>" & rs(2) & "</p>" & vbNewLine)
		End If
		'--- print all sub-questions
		strQuery = ""&_
			"SELECT" & vbNewLine&_
			"    NCHAR(UNICODE('א')+QTS.[Sub_Number]-1) AS [סעיף]" & vbNewLine&_
			"    ,QTS.[Relative_Grade] AS [חלק ניקוד באחוזים]" & vbNewLine&_
			"    ,QTS.[Question_Text] AS [שאלה]" & vbNewLine&_
			"    ,QTS.[Sub_Number] AS [מספר סידורי של סעיף]" & vbNewLine&_
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
				"        <td dir=rtl width=50>" & "(" & iSubGrade & " נק')" & "</td>"&_
				"        <td dir=rtl width=30>" & rsSub(0) & "." & "</td>"&_
				"        <td dir=rtl>" & rsSub(2) & "</td>"&_
				"    </tr>"&_
				"</table>")
'			Response.Write("<table style='font-family: David; font-size: 12pt' width=100% colspan=0 rowspan=0><tr valign=top><td dir=rtl width=50>" & "(" & rsSub(1) & "%)" & "</td><td dir=rtl width=30>" & rsSub(0) & "." & "</td><td dir=rtl>" & rsSub(2) & "</td></tr></table>")
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