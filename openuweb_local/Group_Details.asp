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
	Dim iGroupId
	Dim strAction
	iGroupId = Request("group_id")
	strAction = Request("action")

	'---------------------------------------------------------------------------
	' Update database upon request
	If Request("REQUEST_METHOD") = "POST" Then
		If strAction = "do_update_students" Then
			strQuery = ""&_
				"INSERT INTO [StudentsGroups]" & vbNewLine&_
				"    ([Student_Id], [Group_Id])" & vbNewLine&_
				"VALUES (" & Request("student_id") & ", " & iGroupId & ")"
			Database_Run_Query oConn, strQuery
			Response.Redirect("group_details.asp?group_id=" & iGroupId & "&action=to_update_students")
		End If
	End If	
%>

<html>

<head>
	<link href="Openu.css" rel="stylesheet" type="text/css">
	<title>פרטי קבוצה</title>
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">

<table class="PageTitle_Group">
    <tr>
        <td class="PageTitle">פרטי קבוצה</td>
    </tr>
</table>

<ul>
    <li><a href="default.asp">דף ראשי</a></li>
</ul>
<%
	'--- general details
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    CN.[Name]+' ('+CAST(C.[Course_Number] AS nvarchar)+')' AS [קורס]" & vbNewLine&_
		"    ,C.[Semester] AS [סמסטר]" & vbNewLine&_
		"    ,REPLICATE('0',2-LEN(CAST(CG.[Group_Number] AS varchar)))+CAST(CG.[Group_Number] AS varchar) AS [קבוצה]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [CoursesGroups] CG" & vbNewLine&_
	    "        LEFT OUTER JOIN [Courses] C ON C.[Course_Id]=CG.[Course_Id]" & vbNewLine&_
	    "        LEFT OUTER JOIN [CoursesNames] CN ON C.[Course_Number]=CN.[Course_Number]" & vbNewLine&_
	    "WHERE CG.[Group_Id]=" & iGroupId & vbNewLine&_
	    "ORDER BY C.[Course_Number],C.[Semester]"
	Response.Write(HTML_Style_Header3("group", "פרטים", "") & vbNewLine&_
		HTML_Info_From_Query(oConn, strQuery, True))

	'---------------------------------------------------------------------------
	'--- student list with maman columns
	Dim iMamanCount
	strQuery = ""&_
		"SELECT COUNT(*)" & vbNewLine&_
	    "FROM [Mamans] M" & vbNewLine&_
	    "    INNER JOIN [CoursesGroups] CG" & vbNewLine&_
	    "        ON M.[Course_Id]=CG.[Course_Id]" & vbNewLine&_
	    "WHERE CG.[Group_Id]=" & iGroupId
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
			"            ON M" & (10 + i) & ".[Course_Id]=CG.[Course_Id] AND M" & (10 + i) & ".[Maman_Number]=" & (10 + i) & vbNewLine
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
	    "WHERE CG.[Group_Id]=" & iGroupId & vbNewLine&_
	    "ORDER BY S.[Last],S.[First]"
	Response.Write(HTML_Style_Header3("Student", "סטודנטים בקבוצה", HTML_Style_Button2("עריכה", "group_details.asp?group_id=" & iGroupId & "&action=to_update_students", "/bin/images/edit1.gif", Not strAction = "to_update_students") & HTML_Style_Button2("מצב תצוגה", "group_details.asp?group_id=" & iGroupId, "/bin/images/view1.gif", strAction = "to_update_students")) & vbNewLine&_
	    HTML_Table_From_Query(oConn, strQuery))

	'---------------------------------------------------------------------------
	' In edit mode, let the user add students not already in list
	If strAction = "to_update_students" Then
		'--- collect all the students
		strQuery = ""&_
			"SELECT" & vbNewLine&_
			"	 REPLICATE('0',9-LEN(CAST(S.[Student_Id] AS varchar)))+CAST(S.[Student_Id] as varchar)" & vbNewLine&_
			"    ,S.[Last]" & vbNewLine&_
			"    ,S.[First]" & vbNewLine&_
			"    ,S.[Student_Id]" & vbNewLine&_
			"FROM [Students] S" & vbNewLine&_
			"WHERE NOT EXISTS" & vbNewLine&_
			"    (SELECT SG.[Student_Id]" & vbNewLine&_
			"    FROM [StudentsGroups] SG" & vbNewLine&_
			"    WHERE SG.[Student_Id]=S.[Student_Id] AND SG.[Group_Id]=" & iGroupId & ")" & vbNewLine&_
			"ORDER BY S.[Last],S.[First]"

		Response.Write(HTML_Form("group_details.asp", ""&_
			HTML_Input_Hidden("group_id", iGroupId)&_
			HTML_Input_Hidden("action", "do_update_students")&_
			HTML_Style_Info1("הוספת סטודנט לקבוצה", HTML_Input_Select_From_Query(oConn, strQuery , "student_id") & HTML_Input_Button("הוסף"))))
	End If

	oConn.close
	Set oConn = Nothing
%>

</body>
</html>