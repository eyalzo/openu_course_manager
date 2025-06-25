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
	<title>פרטי מטלה</title>
</head>

<body dir=rtl vlink="#0000FF" link="#0000FF" alink="#0000FF">

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
%>
<table class="PageTitle_Maman">
    <tr>
        <td class="PageTitle">פרטי מטלה</td>
    </tr>
</table>

<ul>
    <li><a href="default.asp">דף ראשי</a></li>
    <li><a href="maman_print.asp?maman_id=<% = iMamanId %>">גרסה להדפסה בחוברת</a></li>
</ul>
<%
	'--- general maman details
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    M.[Maman_Number] AS [ממ""ן]" & vbNewLine&_
		"    ,CN.[Name]+' ('+CAST(C.[Course_Number] AS nvarchar)+')' AS [הקורס]" & vbNewLine&_
		"    ,M.[Material] AS [חומר הלימוד למטלה]" & vbNewLine&_
		"    ,M.[Weight] AS [משקל המטלה]" & vbNewLine&_
		"    ,C.[Semester] AS [סמסטר]" & vbNewLine&_
		"    ,CONVERT(varchar,M.[Delivery_Date],104) AS [תאריך אחרון להגשה]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [Mamans] M" & vbNewLine&_
	    "       LEFT OUTER JOIN [Courses] C" & vbNewLine&_
	    "           ON M.[Course_Id]=C.[Course_Id]" & vbNewLine&_
	    "       LEFT OUTER JOIN [CoursesNames] CN" & vbNewLine&_
	    "           ON C.[Course_Number]=CN.[Course_Number]" & vbNewLine&_
	    "WHERE M.[Maman_Id]=" & iMamanId
	Response.Write(HTML_Style_Header3("Maman","פרטים","") & vbNewLine&_
		HTML_Info_From_Query(oConn, strQuery, True))

	'---------------------------------------------------------------------------
	'--- deliveries and grades, by group
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=group_details.asp?group_id='+CAST(CG.[Group_Id] AS varchar)+'>'+REPLICATE('0',2-LEN(CAST(CG.[Group_Number] AS varchar)))+CAST(CG.[Group_Number] AS varchar)+'</a>' AS [מס' קבוצה]" & vbNewLine&_
		"    ,COUNT(SG.[Student_Id]) AS [סטודנטים]" & vbNewLine&_
		"    ,COUNT(SM.[Student_Id]) AS [הגשות]" & vbNewLine&_
		"    ,COUNT(SM1.[Student_Id]) AS [איחור]" & vbNewLine&_
		"    ,COUNT(SM2.[Student_Id]) AS [העתקה]" & vbNewLine&_
		"    ,ROUND(AVG(SM.[Grade]),1) AS [ממוצע]" & vbNewLine&_
		"    ,ROUND(STDEV(SM.[Grade]),1) AS [סטיית תקן]" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [CoursesGroups] CG" & vbNewLine&_
	    "        INNER JOIN [Mamans] M" & vbNewLine&_
	    "            ON M.[Course_Id]=CG.[Course_Id]" & vbNewLine&_
	    "        INNER JOIN [StudentsGroups] SG" & vbNewLine&_
	    "            ON CG.[Group_Id]=SG.[Group_Id]" & vbNewLine&_
		"        LEFT OUTER JOIN [StudentsMamans] SM" & vbNewLine&_
		"            ON M.[Maman_Id]=SM.[Maman_Id] AND SM.[Student_Id]=SG.[Student_Id]" & vbNewLine&_
		"        LEFT OUTER JOIN [StudentsMamans] SM1" & vbNewLine&_
		"            ON M.[Maman_Id]=SM1.[Maman_Id] AND SM1.[Student_Id]=SG.[Student_Id] AND SM1.[Cancel_Late]=1" & vbNewLine&_
		"        LEFT OUTER JOIN [StudentsMamans] SM2" & vbNewLine&_
		"            ON M.[Maman_Id]=SM2.[Maman_Id] AND SM2.[Student_Id]=SG.[Student_Id] AND SM2.[Cancel_Copy]=1" & vbNewLine&_
	    "WHERE M.[Maman_Id]=" & iMamanId & vbNewLine&_
	    "GROUP BY CG.[Group_Id],CG.[Group_Number]" & vbNewLine&_
	    "ORDER BY CG.[Group_Number]"
	Response.Write(HTML_Style_Header3("Maman","הגשות וציונים","") & vbNewLine&_
		HTML_Table_From_Query(oConn, strQuery))

	'---------------------------------------------------------------------------
	'--- everything was selected, so now show the questions
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    MQ.[Question_Number] AS [שאלה]" & vbNewLine&_
		"    ,'<a href=question_details.asp?question_id='+CAST(QT.[Question_Id] AS varchar)+'>'+CAST(QT.[Question_Id] AS varchar)+'</a>' AS [מס' במאגר]" & vbNewLine&_
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
	    "FROM" & vbNewLine&_
	    "    [MamansQuestions] MQ" & vbNewLine&_
	    "        LEFT OUTER JOIN [QuestionsText] QT" & vbNewLine&_
	    "            ON MQ.[Question_Id]=QT.[Question_Id]" & vbNewLine&_
	    "        LEFT OUTER JOIN [Readings] R" & vbNewLine&_
	    "            ON QT.[Reading_Id]=R.[Reading_Id]" & vbNewLine&_
	    "WHERE MQ.[Maman_Id]=" & iMamanId & vbNewLine&_
		"ORDER BY MQ.[Question_Number]"
	Response.Write(HTML_Style_Header3("Question", "רשימת שאלות", "")&_
		HTML_Table_From_Query(oConn, strQuery))

	'---------------------------------------------------------------------------
	'--- students who didn't get a grade yet
	strQuery = ""&_
		"SELECT" & vbNewLine&_
		"    '<a href=student_details.asp?student_id='+CAST(S.[Student_Id] AS varchar)+'>'+REPLICATE('0',9-LEN(CAST(S.[Student_Id] AS varchar)))+CAST(S.[Student_Id] AS varchar)+'</a>' AS [מס' סטודנט]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN S.[Code_Name] > '' THEN '<font color=magenta><b>'+S.[Last]+' '+S.[First]+'</b></font>'" & vbNewLine&_
		"        ELSE S.[Last]+' '+S.[First]" & vbNewLine&_
		"        END [שם (בהיפוך)]" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN DATEDIFF(dd,M.[Delivery_Date],SM.[Date_Received]) > 10 THEN '<font color=red>'+CONVERT(varchar(8),SM.[Date_Received],5)+'</font>'" & vbNewLine&_
		"        ELSE CONVERT(varchar(8),SM.[Date_Received],5)" & vbNewLine&_
		"        END [התקבל]" & vbNewLine&_
		"    -- add or update button" & vbNewLine&_
		"    ,CASE" & vbNewLine&_
		"        WHEN SM.[Date_Received] IS NULL THEN '<a href=maman_form.asp?from=maman_details&maman_id='+CAST(M.[Maman_Id] as varchar)+'&student_id='+CAST(S.[Student_Id] as varchar)+'&action=to_add_maman><img border=0 width=22 height=22 src=/bin/images/New1.gif alt=''קליטת מטלה''></a>'" & vbNewLine&_
		"        ELSE '<a href=maman_form.asp?from=maman_details&maman_id='+CAST(M.[Maman_Id] as varchar)+'&student_id='+CAST(S.[Student_Id] as varchar)+'&action=to_update_maman><img border=0 width=22 height=22 src=/bin/images/Edit1.gif alt=''הכנסת ציון''></a>'" & vbNewLine&_
		"        END" & vbNewLine&_
	    "FROM" & vbNewLine&_
	    "    [CoursesGroups] CG" & vbNewLine&_
	    "        INNER JOIN [Mamans] M" & vbNewLine&_
	    "            ON M.[Course_Id]=CG.[Course_Id]" & vbNewLine&_
	    "        INNER JOIN [StudentsGroups] SG" & vbNewLine&_
	    "            ON CG.[Group_Id]=SG.[Group_Id]" & vbNewLine&_
	    "        INNER JOIN [Students] S" & vbNewLine&_
	    "            ON S.[Student_Id]=SG.[Student_Id]" & vbNewLine&_
		"        LEFT OUTER JOIN [StudentsMamans] SM" & vbNewLine&_
		"            ON M.[Maman_Id]=SM.[Maman_Id] AND SM.[Student_Id]=SG.[Student_Id]" & vbNewLine&_
	    "WHERE M.[Maman_Id]=" & iMamanId & vbNewLine&_
	    "    AND SM.[Grade] IS NULL" & vbNewLine&_
	    "ORDER BY S.[Last],S.[First]"
	Response.Write(HTML_Style_Header3("Maman","טרם קיבלו ציון","") & vbNewLine&_
		HTML_Table_From_Query(oConn, strQuery))

	oConn.close
	Set oConn = Nothing
%>

</body>
</html>