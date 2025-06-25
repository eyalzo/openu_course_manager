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

<html>

<head>
	<link href="Openu.css" rel="stylesheet" type="text/css">
	<title>Openu Eyal</title>
</head>

<body dir="rtl" vlink="#0000FF" link="#0000FF" alink="#0000FF">

<!-- Navigation bar -->
<b>דף ראשי</b><hr>

<!-- Menu -->
<ul>
    <li><a href="maman_form.asp">טופס מטלה</a></li>
    <li><a href="new_student.asp">סטודנט חדש</a></li>
    <li><a href="Student_List.asp">רשימת סטודנטים</a></li>
    <li><a href="Maman_List.asp">רשימת מטלות</a></li>
    <li><a href="Question_List.asp">מאגר שאלות</a></li>
    <li><a href="Course_List.asp">קורסים</a></li>
</ul>

</body>
</html>