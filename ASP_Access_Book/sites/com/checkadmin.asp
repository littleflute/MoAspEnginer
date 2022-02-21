<!--#include file="conn.asp"-->

<%
dim admin_name,admin_pwd
admin_name=trim(request("admin_name"))
admin_pwd=trim(request("admin_pwd"))

set rs=server.createobject("adodb.recordset")
sql="select * from admin where admin_name='"&admin_name&"'"
rs.open sql,conn,1,1

if not rs.eof then
	if admin_pwd=rs("admin_pwd") then
		session("admin")=admin_name
		 session("loginpwd")=admin_name
 session("loginname")=admin_pwd
		response.redirect "admin_main.asp"
	else
		Response.Write "<script language='javascript'>window.confirm('密码不正确，请后退重填！！');</script>"
		Response.Write "<script language='javascript'>parent.window.history.go(-1);</script>"
	end if
else
		Response.Write "<script language='javascript'>window.confirm('用户名不正确！！');</script>"
		Response.Write "<script language='javascript'>parent.window.history.go(-1);</script>"
end if
%>
