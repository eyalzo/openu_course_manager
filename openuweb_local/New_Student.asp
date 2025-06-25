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
	Dim iStudentId
	Dim strAction
	iStudentId = Request("student_id")
	strAction = Request("action")
	
	'---------------------------------------------------------------------------
	' Update database upon request
	If Request("REQUEST_METHOD") = "POST" Then
		If strAction = "do_insert" Then
			strQuery = ""&_
				"INSERT INTO [Students]" & vbNewLine&_
				"    ([Student_Id],[First],[Last])" & vbNewLine&_
				"VALUES (" & iStudentId & ",'" & Request("First") & "','" & Request("Last") & "')"
			Database_Run_Query oConn, strQuery
			Response.Redirect("student_details.asp?student_id=" & iStudentId & "&action=to_update_course")
		End If
	End If	
%>
<html>

<head>
	<link href="Openu.css" rel="stylesheet" type="text/css">
	<title>סטודנט חדש</title>
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">

<table class="PageTitle_Student">
    <tr>
        <td class="PageTitle">סטודנט חדש</td>
    </tr>
</table>

<ul>
    <li><a href="default.asp">דף ראשי</a></li>
</ul>
<%
	'--- show the fields
	Response.Write(HTML_Form("new_student.asp", ""&_
		HTML_Style_Header3("Student","הזנת מספר זהות ושם", HTML_Input_Button("הכנס למאגר הסטודנטים"))&_
		HTML_Input_Hidden("action", "do_insert")&_
		HTML_Style_Info1("מספר זהות", HTML_Input_Text("student_id", 9, "") & " כולל ספרת הביקורת")&_
		HTML_Input_Set_Focus("student_id")&_
		HTML_Style_Info1("שם פרטי", HTML_Input_Text("first", 20, ""))&_
		HTML_Style_Info1("שם משפחה", HTML_Input_Text("last", 20, ""))))
%>

	</body>
</html>