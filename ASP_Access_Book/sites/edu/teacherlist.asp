<!--#include file="conn.asp"-->
<!--#include file="fenlei.asp"-->
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
<title>��ʦ�б�</title>
<script language="javascript">
function jump(page)
{
 targeturl="teacherlist.asp?page="+page;
 window.location.href=targeturl;
}
</script>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<!--#include file="head.asp"-->
<%
page = request("page")
sql = "select * from teacher"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<center>��ǰ���ݿ���û���κν�ʦ</center>"
 response.end
else
rs.pagesize=20
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
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="500" border="1" align="center" cellpadding=1>
<tr><td align="center" class="header" colspan=7>��ʦ�б�</td></tr>
  <tr align=center>
    <td>ID</td>
    <td>��ʦ����</td>
    <td>��ʦ����<%=strfenlei1%></td>
    <td>��ʦ����<%=strfenlei2%></td>
    <td>�����б�</td>
  </tr>
<%
for i = 1 to rs.pagesize
  if rs.eof then exit for
  response.write "<tr>"
  response.write "<td align=center>"&rs("teacherid")&"</td>"
  response.write "<td>&nbsp;<a href=teacherinfo.asp?id="&rs("teacherid")&" title='�鿴"&rs("teacher")&"�ĸ���ר��'>"&rs("teacher")&"</a></td>"
  response.write "<td>&nbsp;"&rs("fenlei1")&"</td>"
  response.write "<td>&nbsp;"&rs("fenlei2")&"</td>"
  response.write "<td align=center><a href=list.asp?teacherid="&rs("teacherid")&">�����б�</a></td>"
  response.write "</tr>"
  rs.movenext
next
%>
<tr><td colspan=5><table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width=100% border=0 align="center" cellpadding=3>
<tr>
<%
response.write "<td nowrap>ҳ��:<B>"&curpage&"</B>/<B>"&rs.pagecount&"</B> ÿҳ:<B>"&rs.pagesize&"</B> ��ʦ��:<B>"&rs.recordcount&"</B></td>"
response.write "<td nowrap align=left>"
if curpage > 1 then
 response.write "<a href=teacherlist.asp?page=1 title=��һҳ><FONT face=Webdings>9</FONT></a>&nbsp;"
else
 response.write "<FONT face=Webdings>9</FONT>&nbsp;"
end if
if curpage > 1 then
 response.write "<a href=teacherlist.asp?page="&curpage-1&" title=��һҳ><FONT face=Webdings>7</FONT></a>&nbsp;"
else
 response.write "<FONT face=Webdings>7</FONT>&nbsp;"
end if
response.write "<font color=red>["&curpage&"]</font>&nbsp;"
dim k
for k = 1 to 4
 if curpage+k <= rs.pagecount then
  response.write "<a href=teacherlist.asp?page="&curpage+k&">["&curpage+k&"]</a>&nbsp;"
 else
  response.write "["&curpage+k&"]&nbsp;"
 end if
next

if curpage < rs.pagecount then
 response.write "<a href=teacherlist.asp?page="&curpage+1&" title=��һҳ><FONT face=Webdings>8</FONT></a>&nbsp;"
else
 response.write "<FONT face=Webdings>8</FONT>&nbsp;"
end if
if curpage < rs.pagecount then
 response.write "<a href=teacherlist.asp?page="&rs.pagecount&" title=���һҳ><FONT face=Webdings>:</FONT></a>"
else
 response.write "<FONT face=Webdings>:</FONT>"
end if
response.write "</td>"
response.write "<td>"
%>
��ת����<select onchange=jump(this.value)>
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
</table></td></tr>
</table>
<br>
<!--#include file="foot.asp"-->
</body>
</html>
<%
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
end if
%>