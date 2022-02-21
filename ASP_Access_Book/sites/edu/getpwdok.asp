<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<%
id = trim(request("id"))
answer = trim(request("answer"))
password = request("password")
password1 = request("password1")

if id = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请不要捣乱');window.location.href='login.asp';</script>"
 response.end
end if
if answer = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入密码取回答案');history.go(-1);</script>"
 response.end
end if
if password <> password1 then
 conn.close
 set conn = nothing
 response.write "<script>alert('两次密码输入不一致');history.go(-1);</script>"
 response.end
end if

sql = "select * from teacher where teacherid="&id
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.bof and rs.eof then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('请不要捣乱');window.location.href='login.asp';</script>"
 response.end
elseif IsNull(rs("answer")) or trim(rs("answer"))="" or rs("answer")<>answer then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('答案不正确');history.go(-1);</script>"
 response.end
else
 rs("password")=md5(password,32)
 rs.update
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('密码修改成功');window.location.href='login.asp?id="&id&"';</script>"
end if

%>