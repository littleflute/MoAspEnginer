<!--#include file="conn.asp"-->
<%
dim admin_name,admin_pwd
admin_name="admin"
admin_pwd=trim(request("admin_pwd"))

set rs=server.createobject("adodb.recordset")
sql="select * from admin where admin_name='"&admin_name&"'"
rs.open sql,conn,1,3

rs("admin_pwd")=admin_pwd
rs.update
rs.close
		Response.Write "<script language='javascript'>window.confirm('用户密码已经成功修改！！');</script>"
		Response.Write "<script language='javascript'>parent.window.history.go(-1);</script>"
%>
