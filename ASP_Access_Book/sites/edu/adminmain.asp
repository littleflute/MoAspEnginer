<!--#include file="conn.asp"-->
<!--#include file="isadmin.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������</title>
</head>
<frameset rows="*" cols="100,*,0" framespacing="0" frameborder="NO" border="0">
  <frame src="adminleft.asp" name="left" scrolling="NO" noresize>
  <frame src="list.asp" name="main" noresize>
 
</frameset>
<noframes>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
��ʹ��IE4.0����������������վ
</body>
</noframes>
</html>
<%
conn.close
set conn = nothing
%>