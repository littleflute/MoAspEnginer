<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
	<head>
		<title>论坛小论坛</title>
<!--#include file="odbc_connection.asp"-->
<!-- #include file="an.htm" -->
<!-- #include file="menu.asp" -->
<%
		tid=request("bbs_id")
		boardid=request("boardid")
		title=Request.QueryString("title")
			
		sql="update news set flag=0,title='"&"[color=red][置顶]"&title&"[/color]"&"' where bbs_id=" & tid
		db.execute(sql)
db.close
%>
<br>
<br>
<center>id为 <%= id %> 的文章已经置顶！<br>
  <a href="particular.asp?bbs_id=<%=tid%>&boardid=<%=boardid%>">点击这里返回</a>
</center>
<br><br>
<br>
<br><br>
<br>
<!--#include file="foot.asp"-->