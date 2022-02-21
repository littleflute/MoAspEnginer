<!--#include file="conn.asp"-->
<%
dim p_name,p_type,p_count,p_connect,p_phone,p_text,user_name
p_name=trim(request("p_name"))
p_type=trim(request("p_type"))
p_count=trim(request("p_count"))
p_connect=trim(request("p_connect"))
p_phone=trim(request("p_phone"))
p_text=replace(request("p_text"),chr(10),"<br>")
user_name=session("user")

if p_name="" or p_type="" or p_count="" or p_connect="" or p_phone="" or p_text="" then
		Response.Write "<script language='javascript'>window.confirm('对不起，所有选项都为必填！！');</script>"
		Response.Write "<script language='javascript'>parent.window.history.go(-1);</script>"
end if

set rs=server.createobject("adodb.recordset")
sql="select * from orderlist"
rs.open sql,conn,1,3

rs.addnew
rs("p_name")=p_name
rs("p_type")=p_type
rs("p_count")=p_count
rs("p_connect")=p_connect
rs("p_phone")=p_phone
rs("p_text")=p_text
rs("user_name")=user_name
rs.update
rs.close
set rs=nothing
		Response.Write "<script language='javascript'>window.confirm('您的订单已经提交，我们会立即给您答复！！');</script>"
		Response.Write "<script language='javascript'>parent.window.history.go(-1);</script>"

%>
