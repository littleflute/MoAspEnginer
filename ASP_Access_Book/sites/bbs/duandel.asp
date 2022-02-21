<!--#include file="conn.asp"-->
<%
		tid=request("id")
		delsql="delete from duanx where [to]='"&tid&"'"
		conn.execute(delsql)
		conn.close
   response.redirect"duan.asp"
%>