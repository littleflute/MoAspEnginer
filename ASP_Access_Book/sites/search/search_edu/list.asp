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

sql = "select * from config"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
pagesize = rs("pagesize")
rs.close




if course = "" then
 sql5 = ""
else
 sql5 = "main.course like '%"&course&"%'"
 query = query&"&course="&course
end if

if filetype = "" then
 sql6 = ""
else
 sql6 = "main.idoftype="&filetype
 query = query&"&filetype="&filetype
end if

if title = "" then
 sql7 = ""
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
font-family: Tahoma, Verdana; font-size: 9pt; color: #FFFFFF; background-color: #B05706
}
.category{
font-family: Tahoma, Verdana; font-size: 9pt; color: #000000; background-color: #EFEFEF
}
</STYLE>
<title>�����б�</title>
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
<%
if not isadmin and not isteacher then
%>
<!--#include file="head.asp"-->
<%
end if
 sql = "select * from main,teacher,type where main.idofteacher=teacher.teacherid and main.idoftype=type.typeid "
 if sql5<>"" then
sql=sql&" and "&sql5
 end if
  if sql6<>"" then
sql=sql&" and "&sql6
 end if
  if sql7<>"" then
sql=sql&" and "&sql7
 end if
  sql=sql&" order by main.dateandtime desc"

rs.open sql,conn,1,1
if rs.bof and rs.eof then
 response.write "<center>�Բ���û���ҵ�������Ҫ��ļ�¼</center>"
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
<br>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=5 width=550 border=1>
<tr><td>
<table style="BORDER-COLLAPSE: collapse" width=100% border=0>
<tr class="header"><td colspan=3 align=center>�����б�</td></tr>
<%
for j = 1 to rs.pagesize
 if rs.eof then exit for
 response.write "<tr><td width=200>����"&strfenlei1&"��<a href=list.asp?fenlei1="&rs("fenlei1")&" title='ֻ��ʾ����"&strfenlei1&"Ϊ"&rs("fenlei1")&"������'>"&rs("fenlei1")&"</td>"
 response.write "<td width=230>����"&strfenlei2&"��<a href=list.asp?fenlei2="&rs("fenlei2")&" title='ֻ��ʾ����"&strfenlei2&"Ϊ"&rs("fenlei2")&"������'>"&rs("fenlei2")&"</td>"
 response.write "<td width=120>��ʦ������"&rs("teacher")&"</td></tr>"
 response.write "<tr><td>������ϣ�<a href=list.asp?course="&rs("course")&" title='ֻ��ʾ�������Ϊ"&rs("course")&"������'>"&rs("course")&"</td>"
 dateandtime = rs("dateandtime")
 response.write "<td>�������ڣ�<a href=list.asp?orderby=dateandtime"&query&" title='���������ڽ�������'>"&year(dateandtime)&"��"&month(dateandtime)&"��"&day(dateandtime)&"��</td>"
 response.write "<td>�Ķ�������<a href=list.asp?orderby=times"&query&" title='���Ķ�������������'>"&rs("times")&"</td></tr>"
 response.write "<tr><td>�������ͣ�<a href=list.asp?filetype="&rs("typeid")&" title='ֻ��ʾ��������Ϊ"&rs("type")&"������'>"&rs("type")&"</td><td colspan=2>���ϱ��⣺"&rs("title")&"</td></tr>"
 if len(rs("content")) > 30 then
  content = left(rs("content"),30)&"..."
 else
  content = rs("content")
 end if
 response.write "<tr><td colspan=2>��飺"&content&"</td><td>"

 if isadmin then
  response.write "&nbsp;&nbsp;<a href=admindelcourseware.asp?id="&rs("mainid")&"&teacherid="&rs("teacherid")&">ɾ��</a>"
 end if
 if isteacher and rs("teacherid") = cint(session("teacherid")) then
  response.write "&nbsp;&nbsp;<img src=images/i1.gif border=0><a href=teacherdelcourseware.asp?id="&rs("mainid")&">ɾ��</a>"
 end if
 response.write "</td></tr><tr><td colspan=3><HR SIZE=1></td></tr>"
 rs.movenext
next
response.write "<tr>"
response.write "<td nowrap>ҳ��:<B>"&curpage&"</B>/<B>"&rs.pagecount&"</B> ÿҳ:<B>"&rs.pagesize&"</B> ������:<B>"&rs.recordcount&"</B></td>"
response.write "<td nowrap align=left>"
if curpage > 1 then
 response.write "<a href=list.asp?page=1"&query&" title=��һҳ><FONT face=Webdings>9</FONT></a>&nbsp;"
else
 response.write "<FONT face=Webdings>9</FONT>&nbsp;"
end if
if curpage > 1 then
 response.write "<a href=list.asp?page="&curpage-1&query&" title=��һҳ><FONT face=Webdings>7</FONT></a>&nbsp;"
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
 response.write "<a href=list.asp?page="&curpage+1&query&" title=��һҳ><FONT face=Webdings>8</FONT></a>&nbsp;"
else
 response.write "<FONT face=Webdings>8</FONT>&nbsp;"
end if
if curpage < rs.pagecount then
 response.write "<a href=list.asp?page="&rs.pagecount&query&" title=���һҳ><FONT face=Webdings>:</FONT></a>"
else
 response.write "<FONT face=Webdings>:</FONT>"
end if
response.write "</td>"
response.write "<td>"
%>
��ת����<select onchange=javascript:jump(this.value);>
<%
for l = 1 to rs.pagecount
 if l = curpage then
  response.write "<option value="&l&" selected>"&l&"</option>"
 else
  response.write "<option value="&l&">"&l&"</option>"
 end if
next
response.write "</select>ҳ</td></tr>"
%>
</table>
</td></tr>
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