<%@ LANGUAGE = VBScript %>
<%
Option Explicit
Response.CacheControl = "no-cache"	
Response.AddHeader "Pragma", "no-cache" 
'Response.ExpiresAbsolute=#Jan 01, 1980 00:00:00# 
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

	'--- get input parameters first
	Dim iStudentId
	Dim iGroupId
	Dim iMamanId
	Dim strAction
	iStudentId = Request("student_id")
	iGroupId = Request("group_id")
	iMamanId = Request("maman_id")
	strAction = Request("action")
	
	'--- if was sent directly from another page
	Dim strFrom
	strFrom = Request("from")

	Dim i

	'---------------------------------------------------------------------------
	' Update database upon request
'	If Request("REQUEST_METHOD") = "POST" Then
	If strAction > "" Then
		If strAction = "do_add_maman" Then
			'--- if only entered the receive-date
			If Request("date_sent") = "" Then
				'--- student maman main
				strQuery = strQuery&_
					"INSERT INTO [StudentsMamans]" & vbNewLine&_
					"    ([Student_Id], [Maman_Id], [Date_Received], [Comments], [Cancel_Late], [Cancel_Copy])" & vbNewLine&_
					"VALUES (" & iStudentId & ", " & iMamanId & ", '" & Request("date_received") & "', '" & Request("comments") & "', " & Request("cancel_late") & ", " & Request("cancel_copy") & ")" & vbNewLine
				'--- execute the query			
				Database_Run_Query oConn, strQuery
				'--- keep date cookies for next form
				Response.Cookies("maman_form")("date_received") = Request("date_received")
				'--- redirection to the student's existing forms
				If strFrom = "maman_details" Then
					Response.Redirect("maman_details.asp?maman_id=" & iMamanId)
				Else
					Response.Redirect("maman_form.asp?from=" & strFrom & "&student_id=" & iStudentId & "&group_id=" & iGroupId)
				End If
			End If
			'--- student questions grades
			Dim iGrade
			iGrade = 0
			strQuery = ""
			For i = 1 To Request("questions")
				'--- calculate final grade
				iGrade = iGrade + Request("grade_" & i)
				'--- add to the query
				strQuery = strQuery&_
					"INSERT INTO [StudentsMamansQuestions]" & vbNewLine&_
					"    ([Student_Id], [Maman_Id], [Question_Number], [Grade])" & vbNewLine&_
					"VALUES (" & iStudentId & ", " & iMamanId & ", " & i & ", " & Request("grade_" & i) & ")" & vbNewLine
			Next	
			'--- check if maman was already received
			Dim strQueryCheck
			strQueryCheck = ""&_
				"SELECT SM.[Date_Received]" & vbNewLine&_
				"FROM [StudentsMamans] SM" & vbNewLine&_
				"WHERE SM.[Student_Id]=" & iStudentId & " AND SM.[Maman_Id]=" & iMamanId
			If Database_Run_Query_Return_String(oConn,strQueryCheck) = "" Then
				'--- student maman main
				strQuery = strQuery&_
					"INSERT INTO [StudentsMamans]" & vbNewLine&_
					"    ([Student_Id], [Maman_Id], [Date_Received], [Date_Sent], [Grade], [Comments], [Cancel_Late], [Cancel_Copy])" & vbNewLine&_
					"VALUES (" & iStudentId & ", " & iMamanId & ", '" & Request("date_received") & "', '" & Request("date_sent") & "', " & iGrade & ", '" & Request("comments") & "', " & Request("cancel_late") & ", " & Request("cancel_copy") & ")" & vbNewLine
			Else
				'--- student maman main
				strQuery = strQuery&_
					"UPDATE [StudentsMamans]" & vbNewLine&_
					"SET" & vbNewLine&_
					"    [Date_Sent]='" & Request("date_sent") & "'" & vbNewLine&_
					"    ,[Grade]=" & iGrade & vbNewLine&_
					"    ,[Comments]='" & Request("comments") & "'" & vbNewLine&_
					"    ,[Cancel_Late]=" & Request("cancel_late") & vbNewLine&_
					"    ,[Cancel_Copy]=" & Request("cancel_copy") & vbNewLine&_
					"WHERE [Student_Id]=" & iStudentId & " AND [Maman_Id]=" & iMamanId
			End If
			'--- execute the query			
			Database_Run_Query oConn, strQuery
			'--- keep date cookies for next form
			Response.Cookies("maman_form")("date_received") = Request("date_received")
			Response.Cookies("maman_form")("date_sent") = Request("date_sent")
			'--- redirection to the student's existing forms
			If strFrom = "maman_details" Then
				Response.Redirect("maman_details.asp?maman_id=" & iMamanId)
			Else
				Response.Redirect("maman_form.asp?from=" & strFrom & "&student_id=" & iStudentId & "&group_id=" & iGroupId)
			End If
		End If
	End If	
%>
<html>

<head>
	<link href="Openu.css" rel="stylesheet" type="text/css">
	<title>טופס מטלה</title>
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">

<table class="PageTitle_Maman">
    <tr>
        <td class="PageTitle">טופס מטלה</td>
    </tr>
</table>

<ul>
    <li><a href="default.asp">דף ראשי</a></li>
    <li><a href="maman_form.asp">טופס חדש (נקה הכל)</a></li>
</ul>
<%
	'---------------------------------------------------------------------------
	'--- get student-id, and auto-submit to proceed
	If iStudentId = "" Then
		'--- show student-id field
		Response.Write(HTML_Form("maman_form.asp", ""&_
			HTML_Style_Info1("מספר הזהות", HTML_Input_Text("student_id", 9, ""))&_
			HTML_Input_Set_Focus("student_id")&_
			HTML_Input_Auto_Submit("student_id")))
		Response.End
	End If

	'---------------------------------------------------------------------------
	'--- if group wasn't selected, then check how many groups this student is registered to
	If iGroupId = "" Then
		strQuery = ""&_
			"SELECT COUNT(SG.[Group_Id])" & vbNewLine&_
		    "FROM StudentsGroups SG" & vbNewLine&_
		    "WHERE SG.[Student_Id]=" & iStudentId
		Dim iGroups
		iGroups = Database_Run_Query_Return_String(oConn, strQuery)
		'--- if the student has no groups in system
		If iGroups = 0 Then
			TerminateWithMessage "למשתמש זה אין כל רישומי קבוצות לימוד במערכת" & "(i='" & i & "')"
		'--- if the student has only one course, then pick it
		ElseIf iGroups = 1 Then
			strQuery = ""&_
				"SELECT SG.[Group_Id]" & vbNewLine&_
			    "FROM [StudentsGroups] SG" & vbNewLine&_
			    "WHERE SG.[Student_Id]=" & iStudentId
			iGroupId = Database_Run_Query_Return_String(oConn, strQuery)
		ElseIf iMamanId > 0 Then
			Response.Write("קיימים רישומים ליותר מקורס/קבוצה אחת.<br>")
			'temp - get the group-id from the maman-id
			strQuery = ""&_
				"SELECT" & vbNewLine&_
				"    SG.[Group_Id]" & vbNewLine&_
			    "FROM" & vbNewLine&_
			    "    [CoursesGroups] CG" & vbNewLine&_
			    "    ,[StudentsGroups] SG" & vbNewLine&_
			    "    ,[Mamans] M" & vbNewLine&_
			    "WHERE SG.[Student_Id]=" & iStudentId & vbNewLine&_
			    "    AND M.[Maman_Id]=" & iMamanId & vbNewLine&_
			    "    AND SG.[Group_Id]=CG.[Group_Id]" & vbNewLine&_
			    "    AND M.[Course_Id]=CG.[Course_Id]"
			iGroupId = Database_Run_Query_Return_String(oConn, strQuery)
		Else
			TerminateWithMessage "קיימים רישומים ליותר מקורס/קבוצה אחת. בפיתוח..."
		End If
	End If

	'---------------------------------------------------------------------------
	'--- time to show all details known so far
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    REPLICATE('0',9-LEN(CAST(S.[Student_Id] AS varchar)))+CAST(S.[Student_Id] AS varchar)+' '+S.[First]+' '+S.[Last] AS [סטודנט]" & vbNewLine&_
		"    ,C.[Semester] AS [סמסטר]" & vbNewLine&_
		"    ,CAST(CN.[Course_Number] AS varchar)+' '+CN.[Name] AS [קורס]" & vbNewLine&_
		"    ,CAST(SC.[Center_Id] AS varchar)+' '+SC.[Center_Name] AS [מרכז לימוד]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
	    "        WHEN CG.[Group_Number] >= 10 THEN CAST(CG.[Group_Number] AS varchar)" & vbNewLine&_
	    "        ELSE '0'+CAST(CG.[Group_Number] AS varchar)" & vbNewLine&_
	    "        END [קבוצת לימוד]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Students] S" & vbNewLine&_
	    "    ,[CoursesNames] CN" & vbNewLine&_
	    "    ,[Courses] C" & vbNewLine&_
	    "    ,[StudentsGroups] SG" & vbNewLine&_
	    "    ,[CoursesGroups] CG" & vbNewLine&_
	    "        LEFT OUTER JOIN [StudyCenters] SC" & vbNewLine&_
	    "            ON CG.[Center_Id]=SC.[Center_Id]" & vbNewLine&_
	    "WHERE CG.[Group_Id]=" & iGroupId & vbNewLine&_
	    "    AND CN.[Course_Number]=C.[Course_Number]" & vbNewLine&_
	    "    AND C.[Course_Id]=CG.[Course_Id]" & vbNewLine&_
	    "    AND S.[Student_Id]=" & iStudentId & vbNewLine&_
	    "    AND S.[Student_Id]=SG.[Student_Id]" & vbNewLine&_
	    "    AND SG.[Group_Id]=CG.[Group_Id]"
	Response.Write(HTML_Style_Header3("Maman","פרטים","") & vbNewLine&_
		HTML_Info_From_Query(oConn, strQuery, True))

	'--- current maman list
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=maman_details.asp?maman_id='+CAST(M.[Maman_Id] AS varchar)+'>'+CAST(M.[Maman_Number] AS varchar)+'</a>' AS [ממ""ן]" & vbNewLine&_
		"    ,CONVERT(varchar(8),M.[Delivery_Date],5) AS [תאריך להגשה]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN DATEDIFF(dd,M.[Delivery_Date],SM.[Date_Received]) > 10 THEN '<font color=red>'+CONVERT(varchar(8),SM.[Date_Received],5)+'</font>'" & vbNewLine&_
		"        ELSE CONVERT(varchar(8),SM.[Date_Received],5)" & vbNewLine&_
		"        END [התקבל ביום]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN DATEDIFF(dd,M.[Delivery_Date],SM.[Date_Sent]) > 21 THEN '<font color=red>'+CONVERT(varchar(8),SM.[Date_Sent],5)+'</font>'" & vbNewLine&_
		"        ELSE CONVERT(varchar(8),SM.[Date_Sent],5)" & vbNewLine&_
		"        END [נשלח ביום]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN SM.[Cancel_Late]=1 OR SM.[Cancel_Copy]=1 THEN '<strike>'+CAST(SM.[Grade] AS varchar)+'</strike>&nbsp;<font color=red>0</font>'" & vbNewLine&_
		"        WHEN SM.[Grade] < 60 THEN '<font color=red>'+CAST(SM.[Grade] AS varchar)+'</font>'" & vbNewLine&_
		"        WHEN SM.[Grade] >= 95 THEN '<font color=green>'+CAST(SM.[Grade] AS varchar)+'</font>'" & vbNewLine&_
		"        ELSE CAST(SM.[Grade] AS varchar)" & vbNewLine&_
		"        END [ציון]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN SM.[Cancel_Late]=1 THEN '<font color=red>כן</font>'" & vbNewLine&_
		"        ELSE ''" & vbNewLine&_
		"        END [איחור]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN SM.[Cancel_Copy]=1 THEN '<font color=red>כן</font>'" & vbNewLine&_
		"        ELSE ''" & vbNewLine&_
		"        END [העתקה]" & vbNewLine&_
		"    ,SM.[Comments] AS [הערות]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [CoursesGroups] CG" & vbNewLine&_
	    "    ,[StudentsMamans] SM" & vbNewLine&_
	    "    ,[Mamans] M" & vbNewLine&_
	    "WHERE SM.[Student_Id]=" & iStudentId & vbNewLine&_
	    "    AND CG.[Group_Id]=" & iGroupId & vbNewLine&_
	    "    AND SM.[Maman_Id]=M.[Maman_Id]" & vbNewLine&_
	    "    AND M.[Course_Id]=CG.[Course_Id]" & vbNewLine&_
	    "ORDER BY M.[Maman_Number]"
	Response.Write(HTML_Style_Header3("Maman", "מטלות קיימות", HTML_Style_Button2("הוספת מטלה", "maman_form.asp?student_id=" & iStudentId & "&action=to_add_maman", "/bin/images/new1.gif", strAction <> "to_add_maman"))&_
		HTML_Table_From_Query(oConn, strQuery))

	'temp
	strAction = "to_add_maman"
	'--- add a new maman
	If strAction = "to_add_maman" Then
		'--- if maman wasn't selected yet
		If iMamanId = "" Then
			strQuery = ""&_
				"SELECT" & vbNewLine&_
				"    M.[Maman_Number] AS [ממ""ן]" & vbNewLine&_
				"    ,'('+CONVERT(varchar(8),M.[Delivery_Date],5)+')' AS [תאריך להגשה]" & vbNewLine&_
				"    ,M.[Maman_Id]" & vbNewLine&_
			    "FROM" & vbNewLine&_
			    "    [CoursesGroups] CG" & vbNewLine&_
			    "    ,[Mamans] M" & vbNewLine&_
			    "WHERE CG.[Group_Id]=" & iGroupId & vbNewLine&_
			    "    AND M.[Course_Id]=CG.[Course_Id]" & vbNewLine&_
			    "    AND M.[Maman_Id] NOT IN" & vbNewLine&_
				"		 (SELECT SM.[Maman_Id]" & vbNewLine&_
				"        FROM [StudentsMamans] SM" & vbNewLine&_
				"        WHERE SM.[Student_Id]=" & iStudentId & ")" & vbNewLine&_
			    "ORDER BY M.[Maman_Number]"		
			Response.Write(HTML_Form("maman_form.asp", ""&_
				HTML_Input_Hidden("student_id", iStudentId)&_
				HTML_Input_Hidden("group_id", iGroupId)&_
				HTML_Input_Hidden("action", "to_add_maman")&_
				HTML_Style_Header3("Maman","מטלה חדשה",HTML_Input_Button("הוסף")) & vbNewLine&_
				HTML_Style_Info1("הוספת מטלה לסטודנט", HTML_Input_Select_From_Query(oConn, strQuery , "maman_id") & HTML_Input_Set_Focus("maman_id"))))
		Else
		'--- maman was selected, so now have to fill the fields
			strQuery = ""&_
				"SELECT" & vbNewLine&_
				"    M.[Maman_Number] AS [ממ""ן]" & vbNewLine&_
				"    ,CONVERT(varchar(8),M.[Delivery_Date],5) AS [תאריך להגשה]" & vbNewLine&_
				"    ,CONVERT(varchar(11),SM.[Date_Received],106) AS [התקבל ביום]" & vbNewLine&_
				"    ,CONVERT(varchar(8),SM.[Date_Sent],5) AS [נשלח ביום]" & vbNewLine&_
				"    ,SM.[Grade] AS [ציון]" & vbNewLine&_
				"    ,COUNT(MQ.[Maman_Id]) AS [מספר שאלות]" & vbNewLine&_
				"    ,SM.[Comments] AS [הערות]" & vbNewLine&_
			    "FROM" & vbNewLine&_
			    "    [MamansQuestions] MQ" & vbNewLine&_
			    "    ,[Mamans] M" & vbNewLine&_
			    "        LEFT OUTER JOIN [StudentsMamans] SM" & vbNewLine&_
			    "            ON M.[Maman_Id]=SM.[Maman_Id] AND SM.[Student_Id]=" & iStudentId & vbNewLine&_
			    "WHERE M.[Maman_Id]=" & iMamanId & vbNewLine&_
			    "    AND M.[Maman_Id]=MQ.[Maman_Id]" & vbNewLine&_
			    "GROUP BY M.[Maman_Number],M.[Delivery_Date],SM.[Date_Received],SM.[Date_Sent],SM.[Grade],SM.[Comments]"
			Dim rs
			On Error Resume Next
			Set rs = oConn.Execute(strQuery)
			CheckError strQuery
			'--- prepare date fields
			Dim strDateReceived
			If IsNull(rs(2)) Then
				strDateReceived = Request.Cookies("maman_form")("date_received")
			Else
				strDateReceived = rs(2) 'FormatDateTime(rs(2), 2)
			End If
			Dim strDateSent
			If IsNull(rs(3)) Then
				strDateSent = Request.Cookies("maman_form")("date_sent")
			Else
				strDateSent = rs(3)
			End If
			'--- now show the form
			Response.Write(HTML_Form("maman_form.asp", ""&_
				HTML_Input_Hidden("student_id", iStudentId)&_
				HTML_Input_Hidden("group_id", iGroupId)&_
				HTML_Input_Hidden("maman_id", iMamanId)&_
				HTML_Input_Hidden("action", "do_add_maman")&_
				HTML_Input_Hidden("questions", rs(5))&_
				HTML_Input_Hidden("from", strFrom)&_
				HTML_Style_Header3("Maman","מטלה חדשה",HTML_Input_Button("שמור")) & vbNewLine&_
				HTML_Style_Info1("מטלה", rs(0)) & vbNewLine&_
				HTML_Style_Info1("תאריך להגשה", rs(1)) & vbNewLine&_
				HTML_Style_Info1("התקבל ביום", HTML_Input_Date_Picker("date_received", strDateReceived))&_
				HTML_Style_Info1("נשלח ביום", HTML_Input_Date_Picker("date_sent", strDateSent))&_
				Student_Question_List(rs(5)) & vbNewLine&_
				HTML_Input_Set_Focus("grade_1") & vbNewLine&_
				HTML_Style_Info1("ציון", rs(4)) & vbNewLine&_
				HTML_Style_Info1("הערות", HTML_Input_Text("comments", 50, rs(6))) & vbNewLine&_
				HTML_Style_Info1("פסילה עקב הגשה מאוחרת", ""&_
					HTML_Input_Select("cancel_late", "" & vbNewLine&_
						HTML_Input_Select_Option("לא", "0", "0") & vbNewLine&_
						HTML_Input_Select_Option("כן", "1", "0"))) & vbNewLine&_
				HTML_Style_Info1("פסילה עקב העתקה", ""&_
					HTML_Input_Select("cancel_copy", "" & vbNewLine&_
						HTML_Input_Select_Option("לא", "0", "0") & vbNewLine&_
						HTML_Input_Select_Option("כן", "1", "0")))))
		End If
	End If

	oConn.close
	Set oConn = Nothing
%>

</body>
</html>

<%
Private Function Student_Question_List(ByVal iNumberOfQuestions)
	Student_Question_List = ""
	Dim i
	For i = 1 To iNumberOfQuestions
		'--- get question grade (if any)
		strQuery = ""&_
			"SELECT SMQ.[Grade] AS [ציון]" & vbNewLine&_
		    "FROM [StudentsMamansQuestions] SMQ" & vbNewLine&_
		    "WHERE SMQ.[Maman_Id]=" & iMamanId & vbNewLine&_
		    "    AND SMQ.[Student_Id]=" & iStudentId & vbNewLine&_
		    "    AND SMQ.[Question_Number]=" & i
		Student_Question_List = Student_Question_List&_
			HTML_Style_Info1("שאלה " & i, HTML_Input_Text("grade_" & i, 3, Database_Run_Query_Return_String(oConn, strQuery)))
	Next
End Function
%>