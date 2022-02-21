<!--#include file="conn.asp"-->
<%
if request("id") = "" then 
response.redirect"default.asp"
response.end 
end if
set rs_id = Server.CreateObject("ADODB.Recordset")
sql = "select * from friend where id=" & request("id")
rs_id.open sql,conn,3,2
rs_id.close
Set rs_del = Server.CreateObject("ADODB.Recordset")
sql="delete  from friend where id=" & request("id")
rs_del.open sql,conn,3,2

conn.close
response.redirect"your.asp"
response.end
%>