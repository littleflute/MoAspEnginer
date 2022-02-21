<%
'----------------------------------------------



'Option Explicit%><!--#include file="conn.asp"--><%
dim connpic
dim rs_user,rs_today,rs_lar,rs_new,rs_board,rs_boy,rs_girl
dim sql,records,today_records,rs_pic
dim i,j,k

if isempty(session("user_id")) then session("user_id")=1

%>
<html>
<head>
<SCRIPT language=javascript1.2 src=openwin.js></SCRIPT>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Namo WebEditor v4.0(Trial)">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title></title>
<meta name="generator" content="Namo WebEditor v4.0(Trial)">
<style type="text/css"></style>
<style type="text/css">
<!--
body {
	background-image: url(bihaibg.jpg);
}
-->
</style>
<link rel="stylesheet" href="ysb.css">
</head>
<body bgcolor="#ffffff" text="black" link="#003300" vlink="#660000" alink="#CC0033" leftmargin="1"
 topmargin="0" ">
<table width="772" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center"><img src="images/top.gif" width="772" height="112" usemap="#Map" border="0"></div></td>
  </tr>
</table>
<table width="772" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="white" bordercolorlight="white" bordercolordark="white" bgcolor="#FFFFFF">
  <tr> 
    <td colspan="3" bgcolor="#ffffff" valign="top" align="center" bordercolor="#660000"> 
      <table border="1" cellspacing="7" cellpadding="7" bordercolorlight="#990000" bordercolordark="#990000" bordercolor="#660000" width="100%" bgcolor="#FFF7F7" >
        <%Set connpic = Server.CreateObject("ADODB.Connection")    
          DBPath = Server.MapPath("data/picture.mdb")    
          connpic.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath 
	sql = "select id,user_id from pic"
	Set rs_lar = Server.CreateObject("ADODB.Recordset")
	rs_lar.open sql,connpic,3,2


'分页设置
if p="" then
	rs_lar.PageSize=18
	pages=rs_lar.pagecount
	records=rs_lar.recordcount
	currentpage=request("currentpage")
	if currentpage="" or currentpage<1 then currentpage=1
	currentpage=cint(currentpage)
	if currentpage>pages then currentpage=pages
	rs_lar.absolutepage=currentpage
else
   	currentpage=1
	records=0
	pages=1
end if	
	rs_lar.movenext
	for i = 1 to 3 '一共显示？行%> 
        <tr align="center"> <%for j = 1 to 6 '一行显示女会员个数
        if not rs_lar.eof then  %> 
          <td> <a href="display.asp?id=<%=rs_lar("id")%>" target="_blank"><img border="1" src='display.asp?id=<%=rs_lar("id")%>' width="90" height="120" alt="点击放大！"  onClick="return js_openpage(this.href);"></a> 
            <br>
             <a href="read.asp?user_id=<%=rs_lar("user_id")%>" target="_blank">详细资料</a><%if session("admin_pass")="ok" then%><a href="deluser.asp?user_id=<%=rs_lar("user_id")%>">删</a> 
            <a href="delpic.asp?id=<%=rs_lar("id")%>">删照</a><%end if%> </td>
          <%rs_lar.movenext
         end if
          next%> </tr>
        <%next %> 
      </table>
    </td>
  <tr> 
    <td colspan="5" align="center" height="30"> <font size="2">第<font color="red"><%=currentpage%></font>页&nbsp;&nbsp;共<font color="blue"><%=Pages%></font>页</font> 
      <!-------------------------------------分页导航---------------------------------------> 
      <%if currentpage>1 then%> <font size="2"><a href="pics.asp?currentpage=<%=currentpage-1%>&search_txt=<%=search_txt%>&px=<%=px%>"><font size="2">[上一页]</font></a> 
      <%end if%> <%if currentpage<pages then%> <a href="pics.asp?currentpage=<%=currentpage+1%>&search_txt=<%=search_txt%>&px=<%=px%>">[下一]</a> 
      <%end if%> <%if currentpage>1 then%> <a href="pics.asp?currentpage=1&search_txt=<%=search_txt%>&px=<%=px%>">[最首]</a> 
      <%end if%> <%if currentpage<pages then%> <a href="pics.asp?currentpage=<%=pages%>&search_txt=<%=search_txt%>&px=<%=px%>">[最末]</a><a href="adminlogin.asp">[管理入口]</a> 
      <%end if%> 
      
      </font> </td>
    <!-------------------------------------------------------------------------------------------------------------------------><!-------------------------------------------------------------------------------------------------------------------------> 
  </tr>
</table>
<map name="Map">
  <area shape="rect" coords="232,93,274,110" href="default.asp">
  <area shape="rect" coords="294,91,359,114" href="your.asp">
  <area shape="rect" coords="381,90,445,112" href="list.asp">
  <area shape="rect" coords="466,92,530,111" href="register.asp">
  <area shape="rect" coords="558,90,617,115" href="sendphoto.asp">
  <area shape="rect" coords="641,92,704,111" href="pics.asp">
</map>
