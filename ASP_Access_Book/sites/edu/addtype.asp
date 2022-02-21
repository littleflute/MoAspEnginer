<!--#include file="conn.asp"-->
<!--#include file="isadmin.asp"-->
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
<title>添加栏目</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<br>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="260" border="1" align="center" cellpadding=1>
<tr><td align="center" class="header" colspan=2>下面是现有的栏目</td></tr>
<%
sql = "select * from type"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
 response.write "<tr align='center'><td width=180>"&rs("type")&"</td>"
 response.write "<td><a href=edittype.asp?id="&rs("typeid")&">编辑</a>/<a href=deltype.asp?id="&rs("typeid")&">删除</a></td></tr>"
 rs.movenext
loop
rs.close
set rs = nothing
conn.close
set conn = nothing
%>
</table>
<center>
<form action="addtypeok.asp" method="post">
请输入要添加的栏目名称:<input type=text name="addtype" size=10><input type=submit name="submit" value="添加">
<br>（栏目名称可以如“论文”、“实验素材”等）
</form>
</center>
</body>
</html>