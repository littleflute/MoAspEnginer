<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<%
id = trim(request("id"))
password = request("password")
verifycode = request("verifycode")

if id = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入您的登录名或ID');history.go(-1);</script>"
 response.end
end if
if password = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入密码');history.go(-1);</script>"
 response.end
end if

if session("verifycode") = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('登陆超时，请重新登录');window.location.href='login.asp';</script>"
 response.end
elseif session("verifycode") <> verifycode then
 response.write "<script>alert('验证码输入错误');history.go(-1);</script>"
 response.end
end if

if IsNumeric(id) then
	sql = "select * from teacher where teacherid="&id
else
	sql = "select * from teacher where loginname='"&id&"'"
end if
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('登录名或ID输入错误，没有该教师');history.go(-1);</script>"
else
 if password = rs("password") then
  if rs("locked") = 1 then
   rs.close
   set rs = nothing
   conn.close
   set conn = nothing
   response.write "<script>alert('您已被禁止登陆，请与管理员联系');history.go(-1);</script>"
  else
   session("teacherid") = cstr(rs("teacherid"))
   rs.close
   set rs = nothing
   conn.close
   set conn = nothing
   response.redirect "teachermain.asp"
  end if
 else
  rs.close
  set rs = nothing
  conn.close
  set conn = nothing
  session("verifycode") = ""
  response.write "<script>alert('密码错误');window.location.href='login.asp';</script>"
 end if
end if
%>