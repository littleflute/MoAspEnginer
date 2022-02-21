<!--#include file="conn.asp"-->
<%
dim id
id=trim(request("id"))
set rs=server.createobject("adodb.recordset")
sql="select * from guestbook where id="& id
rs.open sql,conn,1,3
rs.delete
rs.close
set rs=nothing
response.redirect "viewguest.asp"
%>
