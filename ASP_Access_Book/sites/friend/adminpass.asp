<!--#include file="conn.asp"-->
<%
admin_name=request("admin_name")
admin_password=request("admin_password")

if InStr(LCase(admin_password),"'")<>0 or InStr(LCase(admin_password),"or")<>0 then 
		response.write "<script language='javascript'>"
		response.write "alert('密码不合法，请重新输入!');"
		response.write "history.go(-1);"
		response.write "</script>"
		response.end
		end if

Set rs_admin = Server.CreateObject("ADODB.Recordset")
sql="select * from admin where admin_name like '" & admin_name & "' and admin_password like '" & admin_password & "'"
rs_admin.open sql,conn,3,2
if rs_admin.eof and rs_admin.bof then
   response.write "<script language='javascript'>"
   response.write "alert('账号或密码错误！');"
   response.write "history.go(-1);"
   response.write "</script>"
   response.end
end if

rs_admin.close
session("admin_pass")="ok"
response.redirect "admin.asp"
%>

