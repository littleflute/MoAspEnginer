<!--#include file="conn.asp"-->
<%
adminpwd = request("adminpwd")
verifycode = trim(request("verifycode"))

if adminpwd = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('����������');history.go(-1);</script>"
 response.end
end if

if session("verifycode") = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('��½��ʱ�������µ�¼');window.location.href='adminlogin.asp';</script>"
 response.end
elseif session("verifycode") <> verifycode then
 response.write "<script>alert('��֤���������');history.go(-1);</script>"
 response.end
end if

sql = "select * from config"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if adminpwd = rs("adminpwd") then
 session("admin") = "admin"
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.redirect "adminmain.asp"
else
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 session("verifycode") = ""
 response.write "<script>alert('�������');window.location.href='adminlogin.asp';</script>"
end if
%>