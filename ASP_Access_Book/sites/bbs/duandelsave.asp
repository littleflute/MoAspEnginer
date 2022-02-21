<!--#include file="conn.asp"-->
<%
		tid=request("id")
		delsql="delete from duanx where id="&tid
		conn.execute(delsql)
		conn.close
   response.redirect"duan.asp"
%>