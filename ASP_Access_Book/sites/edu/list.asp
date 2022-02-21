<!--#include file="conn.asp"-->
<!--#include file="fenlei.asp"-->
<%
page = request("page")
orderby = request("orderby")
teacherid = request("teacherid")
fenlei1 = trim(request("fenlei1"))
fenlei2 = trim(request("fenlei2"))
teacher = trim(request("teacher"))
course = trim(request("course"))
filetype = trim(request("filetype"))
title = trim(request("title"))

if session("admin") = "admin" then
 isadmin = true
else
 isadmin = false
end if

if session("teacherid") <> "" then
 isteacher = true
else
 isteacher = false
end if

sql = "select * from config"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
pagesize = rs("pagesize")
rs.close

if teacherid = "" then
 sql1 = "1=1"
 query = ""
else
 sql1 = "main.idofteacher="&teacherid
 query = "&teacherid="&teacherid
end if

if fenlei1 = "" then
 sql2 = "1=1"
else
 sql2 = "teacher.fenlei1 like '%"&fenlei1&"%'"
 query = query&"&fenlei1="&fenlei1
end if

if fenlei2 = "" then
 sql3 = "1=1"
else
 sql3 = "teacher.fenlei2 like '%"&fenlei2&"%'"
 query = query&"&fenlei2="&fenlei2
end if

if teacher = "" then
 sql4 = "1=1"
else
 sql4 = "teacher.teacher like '%"&teacher&"%'"
 query = query&"&teacher="&teacher
end if

if course = "" then
 sql5 = "1=1"
else
 sql5 = "main.course like '%"&course&"%'"
 query = query&"&course="&course
end if

if filetype = "" then
 sql6 = "1=1"
else
 sql6 = "main.idoftype="&filetype
 query = query&"&filetype="&filetype
end if

if title = "" then
 sql7 = "1=1"
else
 sql7 = "main.title like '%"&title&"%'"
 query = query&"&title="&title
end if
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
<title>资料列表</title>
<script language="javascript">
function showdetail(id)
{
 window.open("detail.asp?id="+id,"detail","width=400,height=400,resizable=no,scrollbars=yes,menubar=no,status=no");
}
function jump(page)
{
 targeturl="list.asp?page="+page+"<%=query%>";
 window.location.href=targeturl;
}
</script>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<!--#include file="head.asp"-->
<%
if orderby = "times" then
 sql = "select * from main,teacher,type where main.idofteacher=teacher.teacherid and main.idoftype=type.typeid and "&sql1&" and "&sql2&" and "&sql3&" and "&sql4&" and "&sql5&" and "&sql6&" and "&sql7&" order by main.times desc"
else
 sql = "select * from main,teacher,type where main.idofteacher=teacher.teacherid and main.idoftype=type.typeid and "&sql1&" and "&sql2&" and "&sql3&" and "&sql4&" and "&sql5&" and "&sql6&" and "&sql7&" order by main.dateandtime desc"
end if
rs.open sql,conn,1,1
if rs.bof and rs.eof then
 response.write "<center>对不起，没有找到符合您要求的记录</center>"
else
if pagesize >= 1 then
 rs.pagesize = pagesize
else
 rs.pagesize = 1
end if
if page = "" then
 curpage = 1
else
 curpage = cint(page)
end if
if curpage < 1 then curpage = 1
if curpage > rs.pagecount then curpage = rs.pagecount
rs.move (curpage-1)*rs.pagesize
%>

<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=5 width=750 border=1>
  <tr>
    <td width="254" valign="top">
	<%
if isadmin  then
%>
<table width=100% border=1 align=center cellPadding=1 cellspacing="3" borderColor=#808080 style="BORDER-COLLAPSE: collapse">
        <tr bgcolor="#0099CC"> 
          <td align="center" ><strong><font color="#FFFFFF">后台管理</font></strong></td>
        </tr>

        <tr> 
          <td align=center bgcolor="#E3EFF2"><a href="adminteacher.asp" target="main">教师管理</a></td>
        </tr>
        <tr> 
          <td align=center bgcolor="#E3EFF2"><a href="addtype.asp" target="main">栏目管理</a></td>
        </tr>
        <tr>
          <td align=center bgcolor="#E3EFF2"><a href="pub.asp" target="main">添加资料</a></td>
        </tr>
        <tr> 
          <td align=center bgcolor="#E3EFF2"><a href="list.asp" target="main">管理资料</a></td>
        </tr>
        <tr> 
          <td align=center bgcolor="#E3EFF2"><a href="logout.asp" target="_top">退出管理</a> 
          </td>
        </tr>
      </table>
	  <%end if%>
    </td>
    <td width="700" valign="top"> 
      <table style="BORDER-COLLAPSE: collapse" width=100% border=0>
        <tr class="header">
          <td colspan=3 align=center>资料列表</td>
        </tr>
        <%
for j = 1 to rs.pagesize
 if rs.eof then exit for
 response.write "<tr><td width=200>所属"&strfenlei1&"：<a href=list.asp?fenlei1="&rs("fenlei1")&" title='只显示所属"&strfenlei1&"为"&rs("fenlei1")&"的资料'>"&rs("fenlei1")&"</td>"
 response.write "<td width=230>所属"&strfenlei2&"：<a href=list.asp?fenlei2="&rs("fenlei2")&" title='只显示所属"&strfenlei2&"为"&rs("fenlei2")&"的资料'>"&rs("fenlei2")&"</td>"
 response.write "<td width=120>教师姓名：<a href=teacherinfo.asp?id="&rs("teacherid")&" title='查看"&rs("teacher")&"的个人专集'>"&rs("teacher")&"</a></td></tr>"
 response.write "<tr><td>相关资料：<a href=list.asp?course="&rs("course")&" title='只显示相关资料为"&rs("course")&"的资料'>"&rs("course")&"</td>"
 dateandtime = rs("dateandtime")
 response.write "<td>更新日期：<a href=list.asp?orderby=dateandtime"&query&" title='按更新日期降序排列'>"&year(dateandtime)&"年"&month(dateandtime)&"月"&day(dateandtime)&"日</td>"
 response.write "<td>阅读次数：<a href=list.asp?orderby=times"&query&" title='按阅读次数降序排列'>"&rs("times")&"</td></tr>"
 response.write "<tr><td>资料类型：<a href=list.asp?filetype="&rs("typeid")&" title='只显示资料类型为"&rs("type")&"的资料'>"&rs("type")&"</td><td colspan=2>资料标题："&rs("title")&"</td></tr>"
 if len(rs("content")) > 30 then
  content = left(rs("content"),30)&"..."
 else
  content = rs("content")
 end if
 response.write "<tr><td colspan=2>简介："&content&"</td><td><img src=images/i1.gif border=0>"
 if isadmin or (isteacher and rs("teacherid") = cint(session("teacherid"))) then
  response.write "<a href=edit.asp?id="&rs("mainid")&">编辑</a>"
 else
  response.write "<a href=# onclick=javascript:showdetail("&rs("mainid")&");>查看</a>"
 end if
 if isadmin then
  response.write "&nbsp;&nbsp;<img src=images/i1.gif border=0><a href=admindelcourseware.asp?id="&rs("mainid")&"&teacherid="&rs("teacherid")&">删除</a>"
 end if
 if isteacher and rs("teacherid") = cint(session("teacherid")) then
  response.write "&nbsp;&nbsp;<img src=images/i1.gif border=0><a href=teacherdelcourseware.asp?id="&rs("mainid")&">删除</a>"
 end if
 response.write "</td></tr><tr><td colspan=3><HR SIZE=1></td></tr>"
 rs.movenext
next
response.write "<tr>"
response.write "<td nowrap>页次:<B>"&curpage&"</B>/<B>"&rs.pagecount&"</B> 每页:<B>"&rs.pagesize&"</B> 资料数:<B>"&rs.recordcount&"</B></td>"
response.write "<td nowrap align=left>"
if curpage > 1 then
 response.write "<a href=list.asp?page=1"&query&" title=第一页><FONT face=Webdings>9</FONT></a>&nbsp;"
else
 response.write "<FONT face=Webdings>9</FONT>&nbsp;"
end if
if curpage > 1 then
 response.write "<a href=list.asp?page="&curpage-1&query&" title=上一页><FONT face=Webdings>7</FONT></a>&nbsp;"
else
 response.write "<FONT face=Webdings>7</FONT>&nbsp;"
end if
response.write "<font color=red>["&curpage&"]</font>&nbsp;"
dim k
for k = 1 to 4
 if curpage+k <= rs.pagecount then
  response.write "<a href=list.asp?page="&curpage+k&query&">["&curpage+k&"]</a>&nbsp;"
 else
  response.write "["&curpage+k&"]&nbsp;"
 end if
next

if curpage < rs.pagecount then
 response.write "<a href=list.asp?page="&curpage+1&query&" title=下一页><FONT face=Webdings>8</FONT></a>&nbsp;"
else
 response.write "<FONT face=Webdings>8</FONT>&nbsp;"
end if
if curpage < rs.pagecount then
 response.write "<a href=list.asp?page="&rs.pagecount&query&" title=最后一页><FONT face=Webdings>:</FONT></a>"
else
 response.write "<FONT face=Webdings>:</FONT>"
end if
response.write "</td>"
response.write "<td>"
%>
        跳转至第<select onchange=javascript:jump(this.value);> 
        <%
for l = 1 to rs.pagecount
 if l = curpage then
  response.write "<option value="&l&" selected>"&l&"</option>"
 else
  response.write "<option value="&l&">"&l&"</option>"
 end if
next
response.write "</select>页</td></tr>"
%>
      </table>
    </td>
  </tr>
</table>
<%
end if
rs.close
set rs = nothing
if not isadmin and not isteacher then
%>
<br>
<!--#include file="foot.asp"-->
<%
end if
conn.close
set conn = nothing
%>
</body>
</html>