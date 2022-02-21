<!--#include file="conn.asp"-->
<!--#include file="isadmin.asp"-->
<%
id = trim(request("id"))
sql = "select * from type where typeid="&id
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
filetype = rs("type")
rs.close
set rs = nothing
conn.close
set conn = nothing
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
</STYLE>
<title>编辑栏目</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<br>
<center>
<form action="edittypeok.asp" method="post">
将栏目“<%=filetype%>”更名为：<input type=text name="addtype" size=10 value="<%=filetype%>">
<input type=hidden name="id" value="<%=id%>">
<input type=submit name="submit" value="修改">
<br>（栏目名称可以如“论文”、“实验素材”等）
</form>
</center>
</body>
</html>