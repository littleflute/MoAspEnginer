<!--#include file="conn.asp"-->
<%
page = trim(request("page"))
filetype = trim(request("type"))
id = trim(request("id"))
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
<title>标题列表</title>
<script language="javascript">
function showdetail(id)
{
 window.open("detail.asp?id="+id,"detail","width=400,height=400,resizable=no,scrollbars=yes,menubar=no,status=no");
}
</script>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<%
set rs = server.createobject("adodb.recordset")
sql = "select * from main where idofteacher="&id&" and idoftype="&filetype
rs.open sql,conn,1,1
if rs.bof and rs.eof then
 response.write "&nbsp;&nbsp;&nbsp;&nbsp;暂无"
else
 rs.pagesize = 30
 if page = "" then
  curpage = 1
 else
  curpage = cint(page)
 end if
 if curpage < 1 then curpage = 1
 if curpage > rs.pagecount then curpage = rs.pagecount
 rs.move (curpage-1)*rs.pagesize
 response.write "<table width=100% border=0 cellPadding=0>"
 for i = 1 to 10
  if rs.eof then exit for
  response.write "<tr>"
  if rs.eof then
   response.write "<td width=33% align=left>&nbsp;</td>"
  else
   if len(rs("title")) > 15 then
     filetitle = left(rs("title"),15)&"..."
   else
     filetitle = rs("title")
   end if
   response.write "<td width=33% align=left>&nbsp;&nbsp;&nbsp;&nbsp;<a href=# title='"&rs("title")&"' onclick=javascript:showdetail("&rs("mainid")&");>"&filetitle&"</a></td>"
   rs.movenext
  end if
  if rs.eof then
   response.write "<td width=34% align=left>&nbsp;</td>"
  else
   if len(rs("title")) > 15 then
     filetitle = left(rs("title"),15)&"..."
   else
     filetitle = rs("title")
   end if
   response.write "<td width=34% align=left>&nbsp;&nbsp;&nbsp;&nbsp;<a href=# title='"&rs("title")&"' onclick=javascript:showdetail("&rs("mainid")&");>"&filetitle&"</a></td>"
   rs.movenext
  end if
  if rs.eof then
   response.write "<td width=33% align=left>&nbsp;</td>"
  else
   if len(rs("title")) > 15 then
    filetitle = left(rs("title"),15)&"..."
   else
    filetitle = rs("title")
   end if
   response.write "<td width=33% align=left>&nbsp;&nbsp;&nbsp;&nbsp;<a href=# title='"&rs("title")&"' onclick=javascript:showdetail("&rs("mainid")&");>"&filetitle&"</a></td>"
   rs.movenext
  end if
  response.write "</tr>"
 next
%>
<tr><td colspan=3 align="right">
<%
if curpage > 1 then
 response.write "<a href=titlelist.asp?page=1&type="&filetype&"&id="&id&" title=第一页><FONT face=Webdings>9</FONT></a>&nbsp;"
else
 response.write "<FONT face=Webdings>9</FONT>&nbsp;"
end if
if curpage > 1 then
 response.write "<a href=titlelist.asp?page="&curpage-1&"&type="&filetype&"&id="&id&" title=上一页><FONT face=Webdings>7</FONT></a>&nbsp;"
else
 response.write "<FONT face=Webdings>7</FONT>&nbsp;"
end if
response.write "<font color=red>["&curpage&"]</font>&nbsp;"
dim k
for k = 1 to 4
 if curpage+k <= rs.pagecount then
  response.write "<a href=titlelist.asp?page="&curpage+k&"&type="&filetype&"&id="&id&">["&curpage+k&"]</a>&nbsp;"
 else
  response.write "["&curpage+k&"]&nbsp;"
 end if
next
if curpage < rs.pagecount then
 response.write "<a href=titlelist.asp?page="&curpage+1&"&type="&filetype&"&id="&id&" title=下一页><FONT face=Webdings>8</FONT></a>&nbsp;"
else
 response.write "<FONT face=Webdings>8</FONT>&nbsp;"
end if
if curpage < rs.pagecount then
 response.write "<a href=titlelist.asp?page="&rs.pagecount&"&type="&filetype&"&id="&id&" title=最后一页><FONT face=Webdings>:</FONT></a>"
else
 response.write "<FONT face=Webdings>:</FONT>"
end if
%>
</td></tr>
</table>
<%
end if
rs.close
set rs = nothing
conn.close
set conn = nothing
%>
</body>
</html>