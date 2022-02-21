<!--#include file="conn.asp"-->
<%
		tid=request("bbs_id")
		delsql="delete from news where bbs_id="&tid
		conn.execute(delsql)
		delsql2="delete from news where parent_id="&tid
		conn.execute(delsql2)
conn.close
   response.redirect"oneedit.asp"
%>
