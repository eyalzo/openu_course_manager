<%@ LANGUAGE = VBScript %>
<%
Option Explicit
Response.CacheControl = "no-cache"	
Response.AddHeader "Pragma", "no-cache" 
Response.ExpiresAbsolute=#Jan 01, 1980 00:00:00# 
%>
<HTML>
<HEAD>
<TITLE>Double-click to pick</TITLE>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function Calendar1_DblClick() {
	if(!Calendar1.ValueIsNull)
	{
		var strMonth = new Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec");
		g_oInputText.value = Calendar1.Day + " " + strMonth[Calendar1.Month - 1] + " " + Calendar1.Year;
		//--- fire the 'onchange' event manually because it's not fired using this update method
		g_oInputText.fireEvent('onchange');
	}
	window.close();
}

//-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript FOR=Calendar1 EVENT=DblClick>
<!--
 Calendar1_DblClick()
//-->
</SCRIPT>
</HEAD>


<BODY>
      <OBJECT id=Calendar1 style="LEFT: 0px; TOP: 0px" 
      classid=clsid:8E27C92B-1264-101C-8A2F-040224009C02 width=200 height=150 
      VIEWASTEXT><PARAM NAME="_Version" VALUE="524288"><PARAM NAME="_ExtentX" VALUE="5292"><PARAM NAME="_ExtentY" VALUE="3969"><PARAM NAME="_StockProps" VALUE="1"><PARAM NAME="BackColor" VALUE="-2147483633"><PARAM NAME="Year" VALUE="1980"><PARAM NAME="Month" VALUE="1"><PARAM NAME="Day" VALUE="1"><PARAM NAME="DayLength" VALUE="1"><PARAM NAME="MonthLength" VALUE="2"><PARAM NAME="DayFontColor" VALUE="0"><PARAM NAME="FirstDay" VALUE="1"><PARAM NAME="GridCellEffect" VALUE="1"><PARAM NAME="GridFontColor" VALUE="10485760"><PARAM NAME="GridLinesColor" VALUE="8421504"><PARAM NAME="ShowDateSelectors" VALUE="-1"><PARAM NAME="ShowDays" VALUE="-1"><PARAM NAME="ShowHorizontalGrid" VALUE="-1"><PARAM NAME="ShowTitle" VALUE="0"><PARAM NAME="ShowVerticalGrid" VALUE="-1"><PARAM NAME="TitleFontColor" VALUE="10485760"><PARAM NAME="ValueIsNull" VALUE="0"></OBJECT>
</BODY>
</HTML>

<script language="javascript">
	//--- get the pointer to the <input> element
	var g_oInputText = window.dialogArguments;
	//--- set window size
	window.dialogHeight = "172px";
	window.dialogWidth = "206px";
//	window.dialogTop = g_oInputText.offsetTop + g_oInputText.clientTop;
//	window.dialogLeft = g_oInputText.offsetParent.offsetLeft + g_oInputText.offsetParent.clientLeft + g_oInputText.offsetLeft + g_oInputText.clientLeft + g_oInputText.clientWidth;
//	window.dialogLeft = g_oInputText.posLeft;

	try
	{
		//--- read the <input value="...">
		var dateGiven = new Date(Date.parse(g_oInputText.value));
		Calendar1.Year = dateGiven.getFullYear();
		Calendar1.Month = dateGiven.getMonth() + 1;
		Calendar1.Day = dateGiven.getDate();
	}
	catch(e)
	{
		Calendar1.Today();
	}
	finally
	{
	}

</script>
