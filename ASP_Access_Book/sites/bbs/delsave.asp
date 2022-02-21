<html>
	<head>
		<title>论坛小论坛</title>
<!--#include file="conn.asp"-->
<!-- #include file="an.htm" -->
<!-- #include file="menu.asp" -->
<%
		tid=request("bbs_id")
		boardid=request("boardid")
		delsql="delete from news where bbs_id="&tid
		conn.execute(delsql)
		delsql2="delete from news where parent_id="&tid
		conn.execute(delsql2)
conn.close
%>
<br>
<br>
<center>id为 <%= id %> 的文章已经删除!<br>
<a href="board.asp?page_no=<%=session("page_no")%>&boardid=<%=boardid%>">点击这里返回文章列表</a> 
</center>
<br><br>
<br>
<br><br>
<br>
<!--#include file="foot.asp"-->