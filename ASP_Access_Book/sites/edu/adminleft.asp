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
</STYLE>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<br>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width=80 border=1 align="center">

   <tr align=center>
    <td height="20"><a href="adminteacher.asp" target="main">��ʦ����</a></td>
  </tr>
  <tr align=center>
    <td height="20"><a href="addtype.asp" target="main">��Ŀ����</a></td>
  </tr>
    <tr align=center>
    <td height="20"><a href="pub.asp" target="main">�������</a></td>
  </tr>
  <tr align=center>
    <td height="20"><a href="list.asp" target="main">��������</a></td>
  </tr>

  <tr align=center>
    <td height="20"><a href="logout.asp" target="_top">�˳�����</a></td>
  </tr>
</table>
</body>
</html>
<%
conn.close
set conn = nothing
%>