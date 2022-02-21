<%
if request("id") = "" then 
response.redirect"default.asp"
response.end 
end if
%>
<!--#include file="connpic.asp"-->
<%Set rs_id = Server.CreateObject("ADODB.Recordset")
sql = "select User_id,id from pic where id=" & request("id")
rs_id.open sql,conn,3,2

if session("admin_pass") <> "ok" then
	if session("user_id") <> rs_id("user_id") then
		response.write("您没有删除这个图片的权限！")
		response.end	
	end if
end if
user_id1=rs_id("user_id")

Set rs_del = Server.CreateObject("ADODB.Recordset")
sql="delete * from pic where id=" & request("id")
rs_del.open sql,conn,3,2

conn.close
%>
<!--#include file="conn.asp"-->
<%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "select * from larchives where user_id=" & user_id1,conn,3,2
rs("photo")=rs("photo")-1
rs.update
conn.close

response.redirect"sendphoto.asp"

response.end
%>