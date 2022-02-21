<% 
response.buffer=true
Response.Expires=0
%>
<!-- #include file="connect.asp" -->
<%
name=request("name")
neirong=request("neirong")
sql="select * from data where id is null"
rs.open sql,conn,3,2
rs.addnew
rs("name")=name
rs("neirong")=neirong
rs("time")=now()
rs.update
rs.close 
response.redirect "add.asp"
%>