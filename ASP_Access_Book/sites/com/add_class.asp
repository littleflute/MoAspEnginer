<!--#include file="conn.asp"-->
<%
dim pic_class
pic_class=trim(request("pic_class"))

set rs=server.createobject("adodb.recordset")
sql="select * from picclass where pic_class='"&pic_class&"'"
rs.open sql,conn,1,3

if rs.eof then
	rs.addnew
	rs("pic_class")=pic_class
	rs.update
	rs.close
	set rs=nothing
	response.redirect "manager_class.asp"
else
		Response.Write "<script language='javascript'>window.confirm('对不起，类别已经存在，请后退重填！！');</script>"
		Response.Write "<script language='javascript'>parent.window.history.go(-1);</script>"
end if
%>
