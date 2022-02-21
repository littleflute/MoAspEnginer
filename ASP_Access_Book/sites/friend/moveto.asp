<!--#include file="conn.asp"-->
<%
id=request("id")
set rs_back=server.createobject("adodb.recordset")
sql="select * from back where id like '" & id & "'"
rs_back.open sql,conn,3,2
set rs_friend=server.createobject("adodb.recordset")
rs_friend.open "friend",conn,3,2
rs_friend.addnew
rs_friend("for_id")=session("user_id")
rs_friend("user_id")=rs_back("user_id")
rs_friend("netname")=rs_back("netname")
rs_friend("sex")=rs_back("sex")
rs_friend("home")=rs_back("home")
rs_friend.update

set rs_back=server.createobject("adodb.recordset")
sql="delete  from back where id like '" & id & "'"
rs_back.open sql,conn,3,2
rs_friend.close

response.redirect "your.asp"
%>