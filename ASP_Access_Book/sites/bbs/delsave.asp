<html>
	<head>
		<title>��̳С��̳</title>
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
<center>idΪ <%= id %> �������Ѿ�ɾ��!<br>
<a href="board.asp?page_no=<%=session("page_no")%>&boardid=<%=boardid%>">������ﷵ�������б�</a> 
</center>
<br><br>
<br>
<br><br>
<br>
<!--#include file="foot.asp"-->