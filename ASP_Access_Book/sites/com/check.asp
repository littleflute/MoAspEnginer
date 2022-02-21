<!--#include file="conn.asp"-->
<%
dim  user_name,user_pwd
user_name=trim(request("user_name"))
user_pwd=trim(request("user_pwd"))


	set rs=server.createobject("adodb.recordset")
	sql="select * from users where user_name='"&user_name&"'"
	rs.open sql,conn,1,1
	
	if not rs.eof then
		if user_pwd=rs("user_pwd") then
			session("user")=user_name
			rs.close
			set rs=nothing
			response.redirect "server_main.asp"
		else
			Response.Write "<script language='javascript'>window.confirm('对不起，密码错误！！');</script>"
			Response.Write "<script language='javascript'>parent.window.history.go(-1);</script>"
		end if
	else
		Response.Write "<script language='javascript'>window.confirm('对不起，用户名不存在！！');</script>"
		Response.Write "<script language='javascript'>parent.window.history.go(-1);</script>"
	end if

%>
