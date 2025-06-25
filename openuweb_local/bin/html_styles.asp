<%
'===============================================================================
' /bin/html_styles.asp
' Generate HTML code for headers, buttons, etc.
'===============================================================================


'-------------------------------------------------------------------------------
' Return: String contains HTML code for textual button, looks like [Button].
Private Function HTML_Style_Button1(ByVal a_strButtonText, ByVal a_strButtonLink, ByVal a_bCondition)
	If a_bCondition Then
		HTML_Style_Button1 = "&nbsp;<b>[<a href='" & a_strButtonLink & "'>" & a_strButtonText & "</a>]</b>"
	Else
		HTML_Style_Button1 = ""
	End If
End Function


'-------------------------------------------------------------------------------
' Return: GIF button, with alternative text.
Private Function HTML_Style_Button2(ByVal a_strButtonText, ByVal a_strButtonLink, ByVal a_strButtonImageSrc, ByVal a_bCondition)
	If a_bCondition Then
		HTML_Style_Button2 = "<a href='" & a_strButtonLink & "'><img align=absmiddle src='" & a_strButtonImageSrc & "' alt='" & a_strButtonText & "' border='0'></a>"
	Else
		HTML_Style_Button2 = ""
	End If
End Function


'-------------------------------------------------------------------------------
' Return: String contains HTML code for standard header.
Private Function HTML_Style_Header1(ByVal a_strHeader, ByVal a_strButtons)
	HTML_Style_Header1 = ""&_
		"<!-- Header: " & a_strHeader & " -->" & vbNewLine&_
		"<table border='0' cellpadding='0' cellspacing='0'>" & vbNewLine&_
		"    <tr height='8'><td></td></tr>" & vbNewLine&_
		"    <tr>" & vbNewLine&_
		"        <td>" & vbNewLine&_
		"            <table border='1' width='100%' cellpadding='2'>" & vbNewLine&_
		"                <tr>" & vbNewLine&_
		"                    <td><b>&nbsp;" & a_strHeader & "&nbsp;</b></td>" & vbNewLine&_
		"                </tr>" & vbNewLine&_
		"            </table>" & vbNewLine&_
		"        </td>" & vbNewLine&_
		"        <td width=10>" & vbNewLine&_
		"        </td>" & vbNewLine&_
		"        <td>" & vbNewLine&_
		a_strButtons&_
		"        </td>" & vbNewLine&_
		"    </tr>" & vbNewLine&_
		"    <tr height='3'><td></td></tr>" & vbNewLine&_
		"</table>" & vbNewLine&_
		"<!-- /Header: " & a_strHeader & " -->" & vbNewLine & vbNewLine
End Function


'-------------------------------------------------------------------------------
' a_strSubject - "Maman", "Group", "Student" etc. See Openu.css.
' Return: String contains HTML code for standard header.
Private Function HTML_Style_Header3(ByVal a_strSubject, ByVal a_strHeader, ByVal a_strButtons)
	HTML_Style_Header3 = ""&_
		"<!-- Header: " & a_strHeader & " -->" & vbNewLine&_
		"<table border='0' cellpadding='0' cellspacing='0'>" & vbNewLine&_
		"    <tr height='8'><td></td></tr>" & vbNewLine&_
		"    <tr>" & vbNewLine&_
		"        <td>" & vbNewLine&_
		"            <table border='1' width='100%' cellpadding='2'>" & vbNewLine&_
		"                <tr>" & vbNewLine&_
		"                    <td class=Subtitle_" & a_strSubject & "><b>&nbsp;" & a_strHeader & "&nbsp;</b></td>" & vbNewLine&_
		"                </tr>" & vbNewLine&_
		"            </table>" & vbNewLine&_
		"        </td>" & vbNewLine&_
		"        <td width=10>" & vbNewLine&_
		"        </td>" & vbNewLine&_
		"        <td>" & vbNewLine&_
		a_strButtons&_
		"        </td>" & vbNewLine&_
		"    </tr>" & vbNewLine&_
		"    <tr height='3'><td></td></tr>" & vbNewLine&_
		"</table>" & vbNewLine&_
		"<!-- /Header: " & a_strHeader & " -->" & vbNewLine & vbNewLine
End Function


'-------------------------------------------------------------------------------
' Return: Table row with two cells, where the laft one has a header, and the 
'         right one has the info.
Private Function HTML_Style_Info1(ByVal a_strRowHeader, ByVal a_strInfo)
	HTML_Style_Info1 =	"    <!-- HTML_Style_Info1(" & a_strRowHeader & ") -->" & vbNewLine&_
						"	 <table border='0'>" & vbNewLine&_
						"	     <tr>" & vbNewLine&_
						"	         <td valign='top'>" & vbNewLine&_
						"                <b>&nbsp;&nbsp;" & a_strRowHeader & ":&nbsp;</b></td>" & vbNewLine&_
						"	         <td>" & vbNewLine&_
						"                " & a_strInfo & vbNewLine&_
						"	         </td>" & vbNewLine&_
						"	     </tr>" & vbNewLine&_
						"	 </table>" & vbNewLine&_
						"    <!-- /HTML_Style_Info1(" & a_strRowHeader & ") -->" & vbNewLine
End Function


'-------------------------------------------------------------------------------
' Return: Help box, centered, colored in black on kind of yellow.
Private Function HTML_Style_Help1(ByVal a_strHelp)
	HTML_Style_Help1 = ""&_
		"<!-- Help -->" & vbNewLine&_
		"<table border='0' align='center' width='80%' cellpadding='0' cellspacing='0'>" & vbNewLine&_
		"    <tr height='5'><td></td></tr>" & vbNewLine&_
		"    <tr>" & vbNewLine&_
		"        <td>" & vbNewLine&_
		"            <table border='1' width='100%' cellpadding='2'>" & vbNewLine&_
		"                <tr>" & vbNewLine&_
		"                    <td bgcolor='#ffffc0'>"&_
		a_strHelp&_
		"                    </td>" & vbNewLine&_
		"                </tr>" & vbNewLine&_
		"            </table>" & vbNewLine&_
		"        </td>" & vbNewLine&_
		"    </tr>" & vbNewLine&_
		"    <tr height='5'><td></td></tr>" & vbNewLine&_
		"</table>" & vbNewLine&_
		"<!-- /Help -->" & vbNewLine & vbNewLine
End Function
%>