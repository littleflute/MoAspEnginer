<!--#include file="conn.asp"-->
<!--#include file="fenlei.asp"-->
<!--#include file="isadmin.asp"-->
<%
fenlei1 = trim(request("fenlei1"))
fenlei2 = trim(request("fenlei2"))
teacher = trim(request("teacher"))
id = trim(request("id"))

if fenlei1 = "" then
 strquery1 = "1=1"
else
 strquery1 = "fenlei1 like '%"&fenlei1&"%'"
end if
if fenlei2 = "" then
 strquery2 = "1=1"
else
 strquery2 = "fenlei2 like '%"&fenlei2&"%'"
end if
if teacher = "" then
 strquery3 = "1=1"
else
 strquery3 = "teacher like '%"&teacher&"%'"
end if
if id = "" then
 strquery4 = "1=1"
else
 strquery4 = "teacherid="&id
end if

sql = "select * from teacher where "&strquery1&" and "&strquery2&" and "&strquery3&" and "&strquery4
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('没有找到符合条件的教师');history.go(-1);</script>"
 response.end
else
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
<title>教师管理</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<br>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="500" border="1" align="center" cellpadding=1>
<tr>
    <td align="center" class="header" colspan=7>搜索到的教师列表 <a href="addteacher.asp">添加</a></td>
  </tr>
  <tr>
    <td>ID</td>
    <td>教师姓名</td>
    <td>教师所属<%=strfenlei1%></td>
    <td>教师所属<%=strfenlei2%></td>
    <td>编辑</td>
    <td>资料列表</td>
    <td>删除</td>
  </tr>
<%
do while not rs.eof
  response.write "<tr>"
  response.write "<td>"&rs("teacherid")&"</td>"
  response.write "<td><a href=teacherinfo.asp?id="&rs("teacherid")&">"&rs("teacher")&"</a></td>"
  response.write "<td>"&rs("fenlei1")&"</td>"
  response.write "<td>"&rs("fenlei2")&"</td>"
  response.write "<td><a href=editteacher.asp?id="&rs("teacherid")&">编辑</a></td>"
  response.write "<td><a href=list.asp?teacherid="&rs("teacherid")&">资料列表</a></td>"
  response.write "<td><a href=delteacher.asp?id="&rs("teacherid")&">删除</a></td>"
  response.write "</tr>"
  rs.movenext
loop
%>
</table>
</body>
</html>
<%
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
end if
%>