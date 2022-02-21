<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="160"valign="top" bgcolor="#C8C7BD" nowrap><div class="siderbar_head"></div><div class="siderbar_head" align="center" >用户访问统计</div><br><div class="siderbar_head">基数：<%=blog_VisitBaseNums%></div><div class="siderbar_head">访问：<%=blog_VisitNums%></div>
        <div class="siderbar_head" align="center"><b>在线用户列表</b><br /></div><br><div class="siderbar_head"><a href="?tjtype=5">在线：<%=Application(CookieName & "_blog_displayOnlineTitle2")%></a></div>
        <div class="siderbar_head" align="center"><b>详细统计列表</b><br /></div><br><div class="siderbar_head"><a href="?tjtype=3">时段流量</a></div><div class="siderbar_head"><a href="?tjtype=2">星期流量</a></div><div class="siderbar_head"><a href="?tjtype=0">月份流量</a></div><div class="siderbar_head"><a href="?tjtype=4">客户软件</a></div>
        <%
	  dim tjtype
tjtype=Trim(Request.QueryString("tjtype"))
if IsNumeric(tjtype) then
if (tjtype<>0 and tjtype<>1 and tjtype<>2 and tjtype<>3 and tjtype<>4 and tjtype<>5) then
tjtype=0
end if
else
tjtype=0
end if
	  
	  SQLQueryNums=SQLQueryNums+1
	  Dim CurPage,Url_Add,memSQL
	  IF CheckStr(Request.QueryString("Page"))<>Empty Then
		  Curpage=CheckStr(Request.QueryString("Page"))
		  IF IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
	  Else
		  Curpage=1
	  End IF
	  memSQL="SELECT * FROM blog_Counter ORDER BY coun_ID DESC"
	  Url_Add="?"
	  IF IsInteger(tjtype)=True Then
		SQLFiltrate=SQLFiltrate&" tjtype=1 AND"
		Url_Add=Url_Add&"tjtype="&tjtype&"&"
	End IF
	  %></td><td width="100%" valign="top" bgcolor="#FFFFFF">
    <%
		
			if tjtype=2 then
          Dim IMGWidth,IMGCount%>
	<div class="content_head" align="left"><b>星期流量</b>：
	    <table width="98%" border="0" cellpadding="4" cellspacing="0" class="smalltxt">
	      <%Dim vWeek,maxWeek,vWeek_Title,vWeek_Nums
	  IMGCount=0
	  maxWeek=Conn.ExeCute("SELECT TOP 1 coun_Char,coun_Nums FROM blog_Counter WHERE coun_Type='Week' ORDER BY coun_Nums DESC")
	  Set vWeek=Server.CreateObject("ADODB.Recordset")
	  SQL="SELECT * FROM blog_Counter WHERE coun_Type='Week' ORDER BY coun_Char ASC"
	  vWeek.Open SQL,Conn,1,1
	  SQLQueryNums=SQLQueryNums+2
	  Do While Not vWeek.Eof
	  		vWeek_Title=vWeek("coun_Char")
			vWeek_Nums=vWeek("coun_Nums")
			Select Case vWeek_Title
			Case "0" : vWeek_Title="星期日"
			Case "1" : vWeek_Title="星期一"
			Case "2" : vWeek_Title="星期二"
			Case "3" : vWeek_Title="星期三"
			Case "4" : vWeek_Title="星期四"
			Case "5" : vWeek_Title="星期五"
			Case "6" : vWeek_Title="星期六"
			End Select
			Response.Write("<tr><td width=""12%"">")
	  		If vWeek("coun_Char")=maxWeek(0) Then
				Response.Write("<b>"&vWeek_Title&"</b>")
			Else
				Response.Write(vWeek_Title)
			End If
			IMGWidth=330*vWeek_Nums\maxWeek(1)
	  		Response.Write("</td><td width=""88%""><img src=""images/counter/bar"&IMGCount&".gif"" width="""&IMGWidth&""" height=""11"" align=""absmiddle"">&nbsp;"&vWeek_Nums&"&nbsp;("&FormatPercent(vWeek_Nums/blog_VisitNums)&")</td></tr>")
			If IMGCount=9 Then 
				IMGCount=0
			Else
				IMGCount=IMGCount+1
			End If
	  		vWeek.MoveNext
	  Loop
	  vWeek.Close
	  Set vWeek=Nothing%>
        </table>
	  </div>
<%end if

if tjtype=3 then%>
          <div class="content_head" align="left"><b>时段流量</b>：<table width="100%" border="0" cellpadding="4" cellspacing="0" class="smalltxt"><%Dim vHour,maxHour,vHour_Title,vHour_Nums
	  IMGCount=0
	  maxHour=Conn.ExeCute("SELECT TOP 1 coun_Char,coun_Nums FROM blog_Counter WHERE coun_Type='Hour' ORDER BY coun_Nums DESC")
	  Set vHour=Server.CreateObject("ADODB.Recordset")
	  SQL="SELECT * FROM blog_Counter WHERE coun_Type='Hour' ORDER BY coun_Char ASC"
	  vHour.Open SQL,Conn,1,1
	  SQLQueryNums=SQLQueryNums+2
	  Do While Not vHour.Eof
	  		vHour_Title=vHour("coun_Char")
			vHour_Nums=vHour("coun_Nums")
			If CInt(vHour_Title)>12 Then 
				vHour_Title=CInt(vHour_Title)-12
				If Len(vHour_Title)<2 Then vHour_Title="0"&vHour_Title
				vHour_Title=vHour_Title&" PM"
			Else
				vHour_Title=vHour_Title&" AM"
			End If
			Response.Write("<tr><td width=""12%"">")
	  		If vHour("coun_Char")=maxHour(0) Then
				Response.Write("<b>"&vHour_Title&"</b>")
			Else
				Response.Write(vHour_Title)
			End If
			IMGWidth=330*vHour_Nums\maxHour(1)
	  		Response.Write("</td><td width=""88%""><img src=""images/counter/bar"&IMGCount&".gif"" width="""&IMGWidth&""" height=""11"" align=""absmiddle"">&nbsp;"&vHour_Nums&"&nbsp;("&FormatPercent(vHour_Nums/blog_VisitNums)&")</td></tr>")
			If IMGCount=9 Then 
				IMGCount=0
			Else
				IMGCount=IMGCount+1
			End If
	  		vHour.MoveNext
	  Loop
	  vHour.Close
	  Set vHour=Nothing%>
        </table></div>
<%end if

if tjtype=4 then%>
         <div class="content_head" align="left"><b>操作系统</b>：<table width="100%" border="0" cellpadding="4" cellspacing="0" class="smalltxt"><%Dim vOS,maxOS,vOS_Title,vOS_Nums
	  IMGCount=0
	  maxOS=Conn.ExeCute("SELECT TOP 1 coun_Char,coun_Nums FROM blog_Counter WHERE coun_Type='OS' ORDER BY coun_Nums DESC")
	  Set vOS=Server.CreateObject("ADODB.Recordset")
	  SQL="SELECT * FROM blog_Counter WHERE coun_Type='OS' ORDER BY coun_Char ASC"
	  vOS.Open SQL,Conn,1,1
	  SQLQueryNums=SQLQueryNums+2
	  Do While Not vOS.Eof
	  		vOS_Title=vOS("coun_Char")
			vOS_Nums=vOS("coun_Nums")
			Response.Write("<tr><td width=""12%"">")
	  		If vOS("coun_Char")=maxOS(0) Then
				Response.Write("<b>"&vOS_Title&"</b>")
			Else
				Response.Write(vOS_Title)
			End If
			IMGWidth=330*vOS_Nums\maxOS(1)
	  		Response.Write("</td><td width=""88%""><img src=""images/counter/bar"&IMGCount&".gif"" width="""&IMGWidth&""" height=""11"" align=""absmiddle"">&nbsp;"&vOS_Nums&"&nbsp;("&FormatPercent(vOS_Nums/blog_VisitNums)&")</td></tr>")
			If IMGCount=9 Then 
				IMGCount=0
			Else
				IMGCount=IMGCount+1
			End If
	  		vOS.MoveNext
	  Loop
	  vOS.Close
	  Set vOS=Nothing%>
        </table></div><br><div class="content_head" align="left"><b>浏览器</b>：<table width="100%" border="0" cellpadding="4" cellspacing="0" class="smalltxt"><%Dim vBrowser,maxBrowser,vBrowser_Title,vBrowser_Nums
	  IMGCount=0
	  maxBrowser=Conn.ExeCute("SELECT TOP 1 coun_Char,coun_Nums FROM blog_Counter WHERE coun_Type='Browser' ORDER BY coun_Nums DESC")
	  Set vBrowser=Server.CreateObject("ADODB.Recordset")
	  SQL="SELECT * FROM blog_Counter WHERE coun_Type='Browser' ORDER BY coun_Char ASC"
	  vBrowser.Open SQL,Conn,1,1
	  SQLQueryNums=SQLQueryNums+2
	  Do While Not vBrowser.Eof
	  		vBrowser_Title=vBrowser("coun_Char")
			vBrowser_Nums=vBrowser("coun_Nums")
			Response.Write("<tr><td width=""12%"">")
	  		If vBrowser("coun_Char")=maxBrowser(0) Then
				Response.Write("<b>"&vBrowser_Title&"</b>")
			Else
				Response.Write(vBrowser_Title)
			End If
			IMGWidth=330*vBrowser_Nums\maxBrowser(1)
	  		Response.Write("</td><td width=""88%""><img src=""images/counter/bar"&IMGCount&".gif"" width="""&IMGWidth&""" height=""11"" align=""absmiddle"">&nbsp;"&vBrowser_Nums&"&nbsp;("&FormatPercent(vBrowser_Nums/blog_VisitNums)&")</td></tr>")
			If IMGCount=9 Then 
				IMGCount=0
			Else
				IMGCount=IMGCount+1
			End If
	  		vBrowser.MoveNext
	  Loop
	  vBrowser.Close
	  Set vBrowser=Nothing%>
        </table></div>
<%end if

if tjtype=0 then%>
        <div class="content_head" align="left"><b>月份流量</b>：
      <table width="100%" border="0" cellpadding="4" cellspacing="0" class="smalltxt"><%Dim vMonth,maxMonth,vMonth_Title,vMonth_Nums
	  IMGCount=0
	  maxMonth=Conn.ExeCute("SELECT TOP 1 coun_Char,coun_Nums FROM blog_Counter WHERE coun_Type='Month' ORDER BY coun_Nums DESC")
	  Set vMonth=Server.CreateObject("ADODB.Recordset")
	  SQL="SELECT * FROM blog_Counter WHERE coun_Type='Month' ORDER BY coun_Char ASC"
	  vMonth.Open SQL,Conn,1,1
	  SQLQueryNums=SQLQueryNums+2
	  Do While Not vMonth.Eof
	  		vMonth_Title=vMonth("coun_Char")
			vMonth_Nums=vMonth("coun_Nums")
			vMonth_Title=Left(vMonth_Title,4)&"&nbsp;-&nbsp;"&Right(vMonth_Title,2)
			Response.Write("<tr><td width=""15%"">")
	  		If vMonth("coun_Char")=maxMonth(0) Then
				Response.Write("<b>"&vMonth_Title&"</b>")
			Else
				Response.Write(vMonth_Title)
			End If
			IMGWidth=330*vMonth_Nums\maxMonth(1)
	  		Response.Write("</td><td width=""85%"">&nbsp;<img src=""images/counter/bar"&IMGCount&".gif"" width="""&IMGWidth&""" height=""11"" align=""absmiddle"">&nbsp;"&vMonth_Nums&"&nbsp;("&FormatPercent(vMonth_Nums/blog_VisitNums)&")</td></tr>")
			If IMGCount=9 Then 
				IMGCount=0
			Else
				IMGCount=IMGCount+1
			End If
	  		vMonth.MoveNext
	  Loop
	  vMonth.Close
	  Set vMonth=Nothing%>
        </table></div>
<%end if

		if tjtype=5 then%>
		<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td height="25">&nbsp;&nbsp;<span class="header"><%= Application(CookieName & "_blog_displayOnlineTitle") %></span></td>
        </tr>
      <tr>
        <td height="25" class="textbox-content"><br><div class="Content"><%= Application(CookieName & "_blog_displayOnline") %></td>
        </tr>
		<tr>
        <td height="25" class="textbox-content">图例 : <img src="images/online_Super.gif" width="16" height="15" border="0" align="absbottom"> 超级管理员&nbsp;&nbsp;<img src="images/online_Admin.gif" width="16" height="15" align="absmiddle">&nbsp;管理员&nbsp;&nbsp;<img src="images/online_Member.gif" width="16" height="15" align="absmiddle">&nbsp;已注册用户&nbsp;&nbsp;<img src="images/Online_Guest.gif" width="16" height="15" align="absmiddle">&nbsp;游客</div></td>
        </tr>
    </table></td>
  </tr>
</table>
		<%end if%>
</td>
  </tr>

  </table>

<!--#include file="footer.asp" -->