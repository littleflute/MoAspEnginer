<!--#include file="conn.asp"-->
<%

id=request("id")
user_id=request("user_id")

set rs_apply=server.createobject("adodb.recordset")
sql="delete  from apply where id like '" & id & "'"
rs_apply.open sql,conn,3,2


set rs_lar=server.createobject("adodb.recordset")
sql="select * from larchives where user_id like '" & session("user_id") & "'"
rs_lar.open sql,conn,3,2

set rs_back=server.createobject("adodb.recordset")
rs_back.open "select * from back",conn,3,2
rs_back.addnew
rs_back("user_id")=session("user_id")
rs_back("for_id")=user_id
rs_back("name")=rs_lar("name")
rs_back("sex")=rs_lar("sex")
rs_back("netname")=rs_lar("netname")
rs_back("result")="ря╬э╬Ь"
rs_back("home")=rs_lar("home")
rs_back("back_date")=date
rs_back.update

rs_back.close
rs_lar.close
response.redirect "your.asp"
%>