<!--#include file="conn.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select * from download where id="&request("id")         
rs.open sql,conn,1,1
if request("downid")="2" then
url=""&rs("filename1")&""
elseif request("downid")="3" then
url=""&rs("filename2")&""
elseif request("downid")="1" then
url=""&rs("filename")&""
elseif request("downid")="4" then
url=""&rs("filename3")&""
elseif request("downid")="5" then
url=""&rs("filename4")&""
end if
response.redirect ""&url&""
response.end
rs.close                               
set rs=nothing                               
conn.close                               
set conn=nothing%>

















