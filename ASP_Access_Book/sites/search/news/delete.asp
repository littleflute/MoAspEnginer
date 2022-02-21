<%@ LANGUAGE="VBSCRIPT" %>
<%
if request.cookies("adminok")="" then
  response.redirect "login.asp"
end if
%>
<!--#include file="articleconn.asp"-->
<%
   dim sql 
   dim rs
   set rs=server.createobject("adodb.recordset")
   sql="delete from learning where articleid="&request("ID")
   rs.open sql,conn,1,1
   rs.close
   set rs=nothing  
   conn.close
   set conn=nothing

   response.redirect "manage.asp"

%>
