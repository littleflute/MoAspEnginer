<!--#include file="conn.asp"-->
<%
if request("id") = "" then 
response.redirect"default.asp"
response.end 
end if
set rs_id = Server.CreateObject("ADODB.Recordset")
sql = "select * from leaveword where id=" & request("id")
rs_id.open sql,conn,3,2

if session("admin_pass") <> "ok" then
	if session("user_id") <> rs_id("for_id") then

		response.write("您没有删除权限！")
		response.end	
	end if
end if

rs_id.close
Set rs_del = Server.CreateObject("ADODB.Recordset")
sql="delete  from leaveword where id="&request("id")
rs_del.open sql,conn,3,2

conn.close

response.redirect "your.asp"

response.end
%>