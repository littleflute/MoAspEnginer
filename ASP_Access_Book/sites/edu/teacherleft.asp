<!--#include file="conn.asp"-->
<!--#include file="isteacher.asp"-->
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
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<br>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width=80 border=1 align="center">
  <tr align=center>
    <td height="20"><a href="pub.asp" target="main">发布资料</a></td>
  </tr>
  <tr align=center>
    <td height="20"><a href="list.asp?teacherid=<%=session("teacherid")%>" target="main">管理资料</a></td>
  </tr>
  <tr align=center>
    <td height="20"><a href="editinfo.asp?id=<%=session("teacherid")%>" target="main">修改资料</a></td>
  </tr>
  <tr align=center>
    <td height="20"><a href="logout.asp" target="_top">退出登陆</a></td>
  </tr>
</table>
</body>
</html>
<%
conn.close
set conn = nothing
%>