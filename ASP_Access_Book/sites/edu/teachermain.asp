<!--#include file="conn.asp"-->
<!--#include file="isteacher.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ʦ����</title>
</head>

<frameset rows="*" cols="150,*" framespacing="0" frameborder="NO" border="0">
  <frame src="teacherleft.asp" name="left" scrolling="NO" noresize>
  <frame src="teacherindex.asp" name="main">
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