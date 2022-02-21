<!--#include file="conn.asp"-->
<!--#include file="isadmin.asp"-->
<%
id = trim(request("id"))
if id = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请不要捣乱');top.window.location.href='adminmain.asp';</script>"
 response.end
end if

sql = "select * from type where typeid="&id
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('请不要捣乱');top.window.location.href='adminmain.asp';</script>"
 response.end
end if
rs.close
set rs = nothing
conn.execute "delete from type where typeid="&id
conn.execute "delete from main where idoftype="&id
conn.close
set conn = nothing
response.write "<script>alert('删除成功');window.location.href='addtype.asp';</script>"
%>