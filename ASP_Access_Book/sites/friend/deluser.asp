<!--#include file="conn.asp"-->
<%
if session("admin_pass")<>"ok" then
   response.redirect "adminlogin.asp"
   response.end
end if
user_id=request("user_id")
Set rs_del = Server.CreateObject("ADODB.Recordset")
sql="delete  from larchives where user_id=" & user_id
rs_del.open sql,conn,3,2

Set rs_del = Server.CreateObject("ADODB.Recordset")
sql="delete  from user_reg where user_id=" & user_id
rs_del.open sql,conn,3,2

Set rs_del = Server.CreateObject("ADODB.Recordset")
sql="delete  from friend where user_id=" & user_id & " or for_id=" & user_id
rs_del.open sql,conn,3,2

Set rs_del = Server.CreateObject("ADODB.Recordset")
sql="delete  from back where user_id=" & user_id & " or for_id=" & user_id
rs_del.open sql,conn,3,2

Set rs_del = Server.CreateObject("ADODB.Recordset")
sql="delete  from leaveword where user_id=" & user_id
rs_del.open sql,conn,3,2

Set rs_del = Server.CreateObject("ADODB.Recordset")
sql="delete  from callboard where user_id=" & user_id
rs_del.open sql,conn,3,2

Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("data/picture.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs_del = Server.CreateObject("ADODB.Recordset")
sql="delete * from pic where user_id=" & user_id
rs_del.open sql,conn,3,2

response.redirect "admin.asp"
response.end
%>


