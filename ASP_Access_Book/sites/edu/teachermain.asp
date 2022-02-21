<!--#include file="conn.asp"-->
<!--#include file="isteacher.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>教师界面</title>
</head>

<frameset rows="*" cols="150,*" framespacing="0" frameborder="NO" border="0">
  <frame src="teacherleft.asp" name="left" scrolling="NO" noresize>
  <frame src="teacherindex.asp" name="main">
</frameset>
<noframes>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
请使用IE4.0以上浏览器浏览该网站
</body>
</noframes>
</html>
<%
conn.close
set conn = nothing
%>