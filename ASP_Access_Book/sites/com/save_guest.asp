<!--#include file="conn.asp"-->

<%
dim user_title,user_text,user_name
user_title=trim(request("user_title"))
user_text=replace(request("user_text"),chr(10),"<br>")
user_name=trim(request("user_name"))

if user_title="" or user_text="" then
		Response.Write "<script language='javascript'>window.confirm('对不起，两者均不能为空！！');</script>"
		Response.Write "<script language='javascript'>parent.window.history.go(-1);</script>"
else
	set rs=server.createobject("adodb.recordset")
	sql="select * from guestbook"
	rs.open sql,conn,1,3
	
	rs.addnew
	rs("user_title")=user_title
	rs("user_text")=user_text
	rs("user_name")=user_name
	rs.update
	rs.close
	set rs=nothing
		Response.Write "<script language='javascript'>window.confirm('您的留言已经成功添加！！');</script>"
		Response.Write "<script language='javascript'>parent.window.history.go(-1);</script>"
end if
%>
