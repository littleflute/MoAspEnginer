<!--#include file="connpic.asp"-->
<% 
id=request("id")

set rs=server.createobject("ADODB.recordset") 
sql="select * from pic where id=" & id
rs.open sql,conn,1,1 
Response.ContentType = "image/jpeg" 
Response.BinaryWrite rs("big")
rs.close 
set rs=nothing 
set connGraph=nothing 

%> 
