<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
	<head>
		<title>��̳С��̳</title>
<!--#include file="odbc_connection.asp"-->
<!-- #include file="an.htm" -->
<!-- #include file="menu.asp" -->
<%
		tid=request("bbs_id")
		boardid=request("boardid")
		title=Request.QueryString("title")
			
		sql="update news set flag=0,title='"&"[color=red][�ö�]"&title&"[/color]"&"' where bbs_id=" & tid
		db.execute(sql)
db.close
%>
<br>
<br>
<center>idΪ <%= id %> �������Ѿ��ö���<br>
  <a href="particular.asp?bbs_id=<%=tid%>&boardid=<%=boardid%>">������ﷵ��</a>
</center>
<br><br>
<br>
<br><br>
<br>
<!--#include file="foot.asp"-->