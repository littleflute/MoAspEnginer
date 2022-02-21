<!-- #include file="../conn.asp" -->
<%
id=Request.QueryString("id")
delsql="delete from data where id="&id
set rs=server.createobject("adodb.recordset")
rs.open delsql,conn,3,3
   response.redirect"add.asp"
%>