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
		If strAction = "do_update_details" Then
			strQuery = ""&_
				"UPDATE [Students]" & vbNewLine&_
				"SET [Address]='" & Request("address") & "'" & vbNewLine&_
				"    ,[City]='" & Request("city") & "'" & vbNewLine&_
				"    ,[Phone_Mobile]='" & Request("phone_mobile") & "'" & vbNewLine&_
				"    ,[Phone_Day]='" & Request("phone_day") & "'" & vbNewLine&_
				"    ,[Phone_Evening]='" & Request("phone_evening") & "'" & vbNewLine&_
				"    ,[Email]='" & Request("email") & "'" & vbNewLine&_
				"    ,[Code_Name]='" & Request("code_name") & "'" & vbNewLine&_
				"    ,[Last_Modified]=GETDATE()" & vbNewLine&_
				"WHERE [Student_Id]=" & iStudentId
			Database_Run_Query oConn, strQuery
			Response.Redirect("student_list.asp?student_id=" & iStudentId)
		ElseIf strAction = "do_update_course" Then
			strQuery = ""&_
				"INSERT INTO [StudentsGroups]" & vbNewLine&_
				"    ([Student_Id], [Group_Id])" & vbNewLine&_
				"VALUES (" & iStudentId & ", " & Request("group_id") & ")"
			Database_Run_Query oConn, strQuery
			'--- keep group cookie for next time
			Response.Cookies("student_details")("group_id") = Request("group_id")
			Response.Redirect("new_student.asp")
		End If
	End If	
%>
<html>

<head>
	<link href="Openu.css" rel="stylesheet" type="text/css">
	<title>פרטי סטודנט</title>
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">

<table class="PageTitle_Student">
    <tr>
        <td class="PageTitle">פרטי סטודנט</td>
    </tr>
</table>

<ul>
    <li><a href="default.asp">דף ראשי</a></li>
</ul>
<%
	'--- student details
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    REPLICATE('0',9-LEN(CAST(S.[Student_Id] AS varchar)))+CAST(S.[Student_Id] AS varchar) AS [מס' סטודנט]" & vbNewLine&_
		"    ,S.[First]+' '+S.[Last] AS [שם]" & vbNewLine&_
		"    ,S.[Address] AS [כתובת]" & vbNewLine&_
		"    ,S.[City] AS [ישוב]" & vbNewLine&_
		"    ,S.[Phone_Mobile] AS [טלפון נייד]" & vbNewLine&_
		"    ,S.[Phone_Day] AS [טלפון יום]" & vbNewLine&_
		"    ,S.[Phone_Evening] AS [טלפון ערב]" & vbNewLine&_
		"    ,S.[Email] AS [דואל]" & vbNewLine&_
		"    ,S.[Code_Name] AS [שם קוד לפרסום ציונים]" & vbNewLine&_
	    "FROM Students S" & vbNewLine&_
	    "WHERE [Student_Id]=" & iStudentId

	'--- update or view details
	If strAction = "to_update_details" Then
		'--- update
		Dim rs
		On Error Resume Next
		Set rs = oConn.Execute(strQuery)
		CheckError strQuery
		'--- show the fields
		Response.Write(HTML_Form("student_details.asp", ""&_
			HTML_Style_Header3("Student", "עדכון פרטים אישיים", HTML_Input_Button("שמור"))&_
			HTML_Input_Hidden("action", "do_update_details")&_
			HTML_Input_Hidden("student_id", iStudentId)&_
			HTML_Style_Info1("מספר זהות", rs(0))&_
			HTML_Style_Info1("שם", rs(1))&_
			HTML_Style_Info1("כתובת", HTML_Input_Text("address", 50, rs(2)))&_
			HTML_Style_Info1("ישוב", HTML_Input_Text("city", 20, rs(3)))&_
			HTML_Style_Info1("טלפון נייד", HTML_Input_Text("phone_mobile", 20, rs(4)) & " שימושי במקרה של ביטול מפגשים")&_
			HTML_Style_Info1("טלפון יום", HTML_Input_Text("phone_day", 20, rs(5)) & " שימושי במקרה של ביטול מפגשים")&_
			HTML_Style_Info1("טלפון ערב", HTML_Input_Text("phone_evening", 20, rs(6)) & " שימושי במקרה של ביטול מפגשים")&_
			HTML_Style_Info1("דואל", HTML_Input_Text("email", 50, rs(7)) & " שימושי במקרה של ביטול מפגשים") & vbNewLine&_
			HTML_Style_Info1("שם קוד לפרסום ציונים", HTML_Input_Text("code_name", 20, rs(8)) & " מיועד לפרסום ציונים אנונימי")))
	Else
		'--- view only
		Response.Write(HTML_Style_Header3("Student", "פרטים אישיים", HTML_Style_Button2("עדכון", "student_details.asp?student_id=" & iStudentId & "&action=to_update_details", "/bin/images/edit1.gif", True))&_
			HTML_Info_From_Query(oConn, strQuery, True))
	End If

	'---------------------------------------------------------------------------
	
	'--- maman list
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=course_details.asp?course_id='+CAST(C.[Course_Id] AS varchar)+'>'+CAST(C.[Course_Number] AS varchar)+'</a> '+CN.[Name] AS [קורס]" & vbNewLine&_
		"    ,C.[Semester] AS [סמסטר]" & vbNewLine&_
		"    ,'<a href=maman_details.asp?maman_id='+CAST(M.[Maman_Id] AS varchar)+'>'+CAST(M.[Maman_Number] AS varchar)+'</a>' AS [ממ""ן]" & vbNewLine&_
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
	    "    [StudentsMamans] SM" & vbNewLine&_
	    "    ,[Mamans] M" & vbNewLine&_
	    "    ,[Courses] C" & vbNewLine&_
	    "    ,[CoursesNames] CN" & vbNewLine&_
	    "WHERE SM.[Student_Id]=" & iStudentId & vbNewLine&_
	    "    AND SM.[Maman_Id]=M.[Maman_Id]" & vbNewLine&_
	    "    AND M.[Course_Id]=C.[Course_Id]" & vbNewLine&_
	    "    AND C.[Course_Number]=CN.[Course_Number]" & vbNewLine&_
	    "ORDER BY C.[Course_Number],C.[Semester],M.[Maman_Number]"
	'--- view maman list
	Response.Write(HTML_Style_Header3("Maman", "מטלות", HTML_Style_Button2("עדכון", "maman_form.asp?student_id=" & iStudentId & "&action=to_add_maman", "/bin/images/edit1.gif", strAction <> "to_update_maman"))&_
		HTML_Table_From_Query(oConn, strQuery))

	'---------------------------------------------------------------------------
	
	'--- course list
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=course_details.asp?course_id='+CAST(C.[Course_Id] AS nvarchar)+'>'+CAST(C.[Course_Number] AS nvarchar)+'</a> ('+CN.[Name]+')' AS [קורס]" & vbNewLine&_
		"    ,C.[Semester] AS [סמסטר]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
	    "        WHEN CG.[Group_Number] >= 10 THEN CAST(CG.[Group_Number] AS varchar)" & vbNewLine&_
	    "        ELSE '0'+CAST(CG.[Group_Number] AS varchar)" & vbNewLine&_
	    "        END AS [קבוצה]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    CoursesNames CN" & vbNewLine&_
	    "    ,Courses C" & vbNewLine&_
	    "    ,CoursesGroups CG" & vbNewLine&_
	    "    ,StudentsGroups SG" & vbNewLine&_
	    "WHERE SG.[Student_Id]=" & iStudentId & vbNewLine&_
	    "    AND CN.[Course_Number]=C.[Course_Number]" & vbNewLine&_
	    "    AND C.[Course_Id]=CG.[Course_Id]" & vbNewLine&_
	    "    AND CG.[Group_Id]=SG.[Group_Id]" & vbNewLine&_
	    "ORDER BY C.[Course_Number]"
	'--- view course list
	Response.Write(HTML_Style_Header3("Course", "קורסים", HTML_Style_Button2("עדכון", "student_details.asp?student_id=" & iStudentId & "&action=to_update_course", "/bin/images/edit1.gif", strAction <> "to_update_course"))&_
		HTML_Table_From_Query(oConn, strQuery))

	'--- update or view courses
	If strAction = "to_update_course" Then
		'--- update
		'--- course list, excluding the ones the student is already registered to
		Dim strQueryNewCourses
		strQueryNewCourses = ""&_
			"SELECT" & vbNewLine&_
			"    CASE" & vbNewLine&_
			"        WHEN CAST(CG.[Group_Id] AS varchar)='" & Request.Cookies("student_details")("group_id") & "' THEN '1'" & vbNewLine&_
			"        ELSE '0'" & vbNewLine&_
			"        END" & vbNewLine&_
			"    ,C.[Course_Number]" & vbNewLine&_
			"    ,CN.[Name]" & vbNewLine&_
			"    ,C.[Semester]" & vbNewLine&_
			"    ,CASE" & vbNewLine&_
		    "        WHEN CG.[Group_Number] >= 10 THEN CAST(CG.[Group_Number] AS varchar)" & vbNewLine&_
		    "        ELSE '0'+CAST(CG.[Group_Number] AS varchar)" & vbNewLine&_
		    "        END" & vbNewLine&_
			"    ,SC.[Center_Name]" & vbNewLine&_
			"    ,CG.[Group_Id]" & vbNewLine&_
		    "FROM" & vbNewLine&_
		    "    [CoursesNames] CN" & vbNewLine&_
		    "    ,[Courses] C" & vbNewLine&_
		    "    ,[CoursesGroups] CG" & vbNewLine&_
		    "        LEFT OUTER JOIN [StudyCenters] SC" & vbNewLine&_
		    "            ON CG.[Center_Id]=SC.[Center_Id]" & vbNewLine&_
		    "WHERE CN.[Course_Number]=C.[Course_Number]" & vbNewLine&_
		    "    AND C.[Course_Id]=CG.[Course_Id]" & vbNewLine&_
		    "    AND C.[Course_Id] NOT IN" & vbNewLine&_
		    "        (SELECT CG.[Course_Id]" & vbNewLine&_
		    "        FROM [StudentsGroups] SG" & vbNewLine&_
		    "            ,[CoursesGroups] CG" & vbNewLine&_
		    "        WHERE SG.[Student_Id]=" & iStudentId & vbNewLine&_
		    "            AND SG.[Group_Id]=CG.[Group_Id])" & vbNewLine&_
		    "ORDER BY C.[Semester] DESC,C.[Course_Number],CG.[Group_Number]"
		'--- show the form with a drop-down list of courses
		Response.Write(HTML_Form("student_details.asp", ""&_
			HTML_Input_Hidden("student_id", iStudentId)&_
			HTML_Input_Hidden("action", "do_update_course")&_
			HTML_Style_Info1("הוספת קורס לסטודנט", ""&_
				HTML_Input_Select_From_Query(oConn, strQueryNewCourses , "group_id")&_
				HTML_Input_Button("הוסף")&_
				HTML_Input_Set_Focus("button"))))
	End If

	oConn.close
	Set oConn = Nothing
%>

</body>
</html>