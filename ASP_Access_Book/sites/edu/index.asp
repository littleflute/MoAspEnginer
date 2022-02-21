<!--#include file="conn.asp"-->
<!--#include file="fenlei.asp"-->
<%
sql = "select * from config"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,3
schoolname = rs("schoolname")
gonggao = rs("gonggao")
todaytimes = rs("todaytimes")
times = rs("times")
if rs("todaydate") <> date() then
 rs("todaydate") = date()
 rs.update
 todaytimes = 0
end if
if request.cookies("counted") <> "yes" then
 response.cookies("counted") = "yes"
 response.cookies("counted").expires = now() + 1/72
 times = times + 1
 todaytimes = todaytimes + 1
 rs("times") = times
 rs("todaytimes") = todaytimes
 rs.update
end if
rs.close
dim num1
dim rndnum
Randomize
Do While Len(rndnum)<4
 num1=CStr(Chr((57-48)*rnd+48))
 rndnum=rndnum&num1
loop
session("verifycode")=rndnum
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE>TD {
	FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
BODY {
	FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
A:link {
	COLOR: #0033cc; TEXT-DECORATION: none
}
A:visited {
	COLOR: #0033cc; TEXT-DECORATION: none
}
A:active {
	COLOR: #ff0000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #000000; TEXT-DECORATION: underline
}
.header	{
font-family: Tahoma, Verdana; font-size: 9pt; color: #FFFFFF; background-color: #00CCCC
}
.category{
font-family: Tahoma, Verdana; font-size: 9pt; color: #000000; background-color: #EFEFEF
}
</STYLE>
<title><%=schoolname%></title>
<script language="javascript">
function showdetail(id)
{
 window.open("detail.asp?id="+id,"detail","width=400,height=400,resizable=no,scrollbars=yes,menubar=no,status=no");
}
</script>
</head>
<body bgColor=#ffffff background="crossbg.gif" text=#000000 leftMargin=0 topMargin=0>
<!--#include file="head.asp"-->
<table width="750" border=0 align=center cellPadding=5 borderColor=#808080 bgcolor="#FFFFFF" style="BORDER-COLLAPSE: collapse">
  <tr align=center valign=top>
<td width=170>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=1 width=100% border=1>
<tr>
          <td align=center bgcolor="#0099CC"><strong><font color="#FFFFFF">管理员公告</font></strong></td>
        </tr>
<tr >
<%if gonggao <> "" then%>
<td align=center valign=middle>
&nbsp;<marquee id=newslist onmouseover=newslist.stop() onmouseout=newslist.start()
 scrollAmount=1 scrollDelay=80 direction=up width=90% height=85><%=gonggao%></marquee>
<%else%>
<td align=left valign=top>
&nbsp;暂无
<%end if%>
</td></tr>
</table>
      <br><form action="list.asp" method="post">
      <table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=1 width=100% border=1>
        <tr bgcolor="#0099CC"> 
          <td colspan=2 align="center" ><strong><font color="#FFFFFF">快速搜索</font></strong></td>
        </tr>
        <tr> 
          <td width=40 align="center">资料</td>
          <td>&nbsp; <input type=text name="course" size=12></td>
        </tr>
        <tr> 
          <td align="center">标题</td>
          <td>&nbsp; <input type=text name="title" size=12></td>
        </tr>
        <%
sql = "select * from type"
rs.open sql,conn,1,1
%>
        <tr> 
          <td align="center">类型</td>
          <td> &nbsp; <select name="filetype">
              <option value="" selected>请选择&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
              <%
do while not rs.eof
%>
              <option value="<%=rs("typeid")%>"><%=rs("type")%></option>
              <%
rs.movenext
loop
rs.close
%>
            </select> </td>
        </tr>
        <tr> 
          <td colspan=2 align=center> <input type=submit name="submit2" value="搜索"> 
            &nbsp;&nbsp; <input type=reset name="reset2" value="清空"> </td>
        </tr>
      </table></form> </td>
<td width=430>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=1 width=100% border=1>
        <tr bgcolor="#FFFFFF"  > 
          <td colspan=4 background="d.jpg" align=center>最近更新</td>
        </tr>
<tr align=center class="category"><td width=180>资料标题</td><td>教师姓名</td><td>更新日期</td><td>所属<%=strfenlei2%></td></tr>
<%
sql = "select * from main,teacher where main.idofteacher=teacher.teacherid order by main.dateandtime desc"
rs.open sql,conn,1,1
for i = 1 to 10
 if rs.eof then
  response.write "<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
 else
  dateandtime = rs("dateandtime")
  updatedate = year(dateandtime)&"/"&month(dateandtime)&"/"&day(dateandtime)
  if len(rs("title")) > 10 then
   filetitle = left(rs("title"),10)&"..."
  else
   filetitle = rs("title")
  end if
%>
<tr><td align=left>&nbsp;<img src=images/arrow.gif>
<a href=# title="<%=rs("title")%>" onclick=javascript:showdetail(<%=rs("mainid")%>);><%=filetitle%></a>
</td><td align=center><a href=teacherinfo.asp?id=<%=rs("teacherid")%> title="查看<%=rs("teacher")%>的个人专集"><%=rs("teacher")%></a></td>
<td align=center><%=updatedate%></td>
<td align=left>&nbsp;&nbsp;<%=rs("fenlei2")%></td></tr>
<%
  rs.movenext
 end if
next
rs.close
%>
</table>
    </td>
<td width=150>

<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=1 width=100% border=1>
<form action="check.asp" method="post">
          <tr bgcolor="#0099CC" > 
            <td colspan=2 align=center><strong><font color="#FFFFFF">教师登陆</font></strong></td>
          </tr>
<tr><td width=45 align=center>登录名</td><td align=left>&nbsp;<input type=text name="id" size=10 title="请输入您的登录名或ID"></td></tr>
<tr><td width=45 align=center>密码</td><td align=left>&nbsp;<input type=password name="password" size=10 title="请输入您的密码"></td></tr>
<tr><td width=45 align=center>验证码</td><td align=left>&nbsp;<input type=text name="verifycode" size=4 title="请输入右侧的数字"> <img src="num.ASP">
</td></tr>
<tr><td colspan=2 align=center><input type=submit name="submit" value="登陆">
              &nbsp;&nbsp;</td>
          </tr>
 </form></table>
        
      <img src="e.jpg" width="150" height="138"></td>
</tr>
</table>
<table width="750" border=0 align=center cellPadding=5 borderColor=#808080 bgcolor="#FFFFFF" style="BORDER-COLLAPSE: collapse">
  <tr align=center valign=top> 
    <td width=170> 
      <table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=1 width=100% border=1>
        <tr>
          <td align=center bgcolor="#0099CC"><font color="#FFFFFF"><strong>新注册用户</strong></font></td>
        </tr>
        <tr height=90>
          <td align=left valign=top> 
            <%
sql = "select * from teacher order by teacherid desc"
rs.open sql,conn,1,1
if rs.bof and rs.eof then
 response.write "&nbsp;暂无"
else
 for i = 1 to 15
  if rs.eof then exit for
  response.write "&nbsp;<font color=#42a5f7 face=Wingdings>v</font>"
  response.write "&nbsp;<a href=teacherinfo.asp?id="&rs("teacherid")&" title='查看"&rs("teacher")&"的个人专集'>"&rs("teacher")&"</a>（"&rs("fenlei2")&"）<br>"
  rs.movenext
 next
end if
rs.close
%>
          </td>
        </tr>
      </table>
    </td>
    <td width=430> 
      <table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=1 width=100% border=1>
        <tr class="header" bgcolor="#FFFFFF" > 
          <td colspan=4 background="d.jpg" align=center><font color="#000000">阅读排行</font></td>
        </tr>
        <tr align=center class="category">
          <td width=180>资料标题</td>
          <td>教师姓名</td>
          <td>阅读次数</td>
          <td>所属<%=strfenlei2%></td>
        </tr>
        <%
sql = "select * from main,teacher where main.idofteacher=teacher.teacherid order by main.times desc"
rs.open sql,conn,1,1
for i = 1 to 9
 if rs.eof then
  response.write "<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
 else
  if len(rs("title")) > 10 then
   filetitle = left(rs("title"),10)&"..."
  else
   filetitle = rs("title")
  end if
%>
        <tr>
          <td align=left>&nbsp;<img src=images/arrow.gif> <a href=# title="<%=rs("title")%>" onclick=javascript:showdetail(<%=rs("mainid")%>);><%=filetitle%></a> 
          </td>
          <td align=center><a href=teacherinfo.asp?id=<%=rs("teacherid")%> title="查看<%=rs("teacher")%>的个人专集"><%=rs("teacher")%></a></td>
          <td align=center><%=rs("times")%></td>
          <td align=left>&nbsp;&nbsp;<%=rs("fenlei2")%></td>
        </tr>
        <%
  rs.movenext
 end if
next
rs.close
%>
      </table></td>
    <td width=150> 
      <table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=1 width=100% border=1>
        <tr>
          <td align=center bgcolor="#0099CC"><font color="#FFFFFF"><strong>资源列表</strong></font></td>
        </tr>
        <tr >
          <td align=left valign=top> 
            <%
sql = "select count(teacherid) from teacher"
rs.open sql,conn,1,1
teachercount = rs(0)
rs.close
response.write "<center>本站现有<br></center>"
response.write "&nbsp;<font color=#42a5f7 face=Wingdings>v</font>&nbsp;<a href=teacherlist.asp>注册教师</a>："&teachercount&"<br>"
sql = "select * from type"
rs.open sql,conn,1,1
set rs1 = server.createobject("adodb.recordset")
do while not rs.eof
 response.write "&nbsp;<font color=#42a5f7 face=Wingdings>v</font>&nbsp;<a href=list.asp?filetype="&rs("typeid")&">"&rs("type")&"</a>："
 sql1 = "select count(mainid) from main where idoftype="&rs("typeid")
 rs1.open sql1,conn,1,1
 response.write rs1(0)&"<br>"
 rs1.close
 rs.movenext 
loop
set rs1 = nothing
rs.close
%>
          </td>
        </tr>
      </table>
      <br> <table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=1 width=100% border=1>
        <tr>
          <td align=center bgcolor="#0099CC"><strong><font color="#FFFFFF">访问统计</font></strong></td>
        </tr>
        <tr height=40>
          <td align=left valign=top> &nbsp;<font color=#42a5f7 face=Wingdings>v</font>&nbsp;总访问量：<%=times%><br> 
            &nbsp;<font color=#42a5f7 face=Wingdings>v</font>&nbsp;今日访问：<%=todaytimes%> 
          </td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="foot.asp"-->
<%
set rs = nothing
conn.close
set conn = nothing
%>
</body>
</html>